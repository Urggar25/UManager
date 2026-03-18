const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');
const { initDb, all, get, run } = require('./db');

const app = express();
const PORT = process.env.PORT || 3000;

const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const upload = multer({ dest: uploadDir });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

function normalizeForId(value) {
  return (value || '')
    .toString()
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '_')
    .replace(/[^a-zA-Z0-9_]/g, '')
    .toUpperCase();
}

function formatDatePart(dateInput) {
  const date = new Date(dateInput);
  if (Number.isNaN(date.getTime())) {
    throw new Error(`Date invalide: ${dateInput}`);
  }
  const day = String(date.getUTCDate()).padStart(2, '0');
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const year = date.getUTCFullYear();
  return {
    iso: `${year}-${month}-${day}`,
    part: `${day}_${month}_${year}`
  };
}

function buildUniqueId(lastNameBirth, firstName, birthDate) {
  const last = normalizeForId(lastNameBirth);
  const first = normalizeForId(firstName);
  const datePart = formatDatePart(birthDate).part;
  return `${last}_${first}_${datePart}`;
}

async function listPeople(search = '') {
  const like = `%${search.trim()}%`;
  const people = await all(
    `
      SELECT *
      FROM people
      WHERE ? = ''
         OR unique_id LIKE ?
         OR last_name_birth LIKE ?
         OR first_name LIKE ?
         OR keywords_json LIKE ?
      ORDER BY last_name_birth COLLATE NOCASE, first_name COLLATE NOCASE
    `,
    [search.trim(), like, like, like, like]
  );

  const categoryValues = await all(`
    SELECT pcv.person_id, c.id AS category_id, c.label AS category_label, pcv.value
    FROM person_category_values pcv
    JOIN categories c ON c.id = pcv.category_id
  `);

  const groupedValues = categoryValues.reduce((acc, row) => {
    if (!acc[row.person_id]) {
      acc[row.person_id] = [];
    }
    acc[row.person_id].push({
      categoryId: row.category_id,
      categoryLabel: row.category_label,
      value: row.value
    });
    return acc;
  }, {});

  return people.map((person) => ({
    id: person.id,
    uniqueId: person.unique_id,
    lastNameBirth: person.last_name_birth,
    firstName: person.first_name,
    birthDate: person.birth_date,
    keywords: JSON.parse(person.keywords_json || '[]'),
    customValues: groupedValues[person.id] || []
  }));
}

app.get('/api/categories', async (_req, res) => {
  try {
    const categories = await all(`SELECT * FROM categories ORDER BY required DESC, label COLLATE NOCASE`);
    res.json(categories);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/categories', async (req, res) => {
  const { label } = req.body;
  if (!label || !label.trim()) {
    return res.status(400).json({ error: 'Le nom de catégorie est obligatoire.' });
  }
  try {
    const result = await run(
      `INSERT INTO categories (code, label, required) VALUES (?, ?, 0)`,
      [null, label.trim()]
    );
    const category = await get(`SELECT * FROM categories WHERE id = ?`, [result.lastID]);
    return res.status(201).json(category);
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

app.put('/api/categories/:id', async (req, res) => {
  const { label } = req.body;
  if (!label || !label.trim()) {
    return res.status(400).json({ error: 'Le nom de catégorie est obligatoire.' });
  }
  try {
    await run(`UPDATE categories SET label = ? WHERE id = ?`, [label.trim(), req.params.id]);
    const category = await get(`SELECT * FROM categories WHERE id = ?`, [req.params.id]);
    return res.json(category);
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

app.post('/api/people', async (req, res) => {
  const {
    lastNameBirth,
    firstName,
    birthDate,
    keywords = [],
    customValues = {}
  } = req.body;

  if (!lastNameBirth || !firstName || !birthDate) {
    return res.status(400).json({
      error: 'Nom de naissance, prénom et date de naissance sont obligatoires.'
    });
  }

  try {
    const { iso } = formatDatePart(birthDate);
    const uniqueId = buildUniqueId(lastNameBirth, firstName, iso);
    const insertPerson = await run(
      `
      INSERT INTO people (unique_id, last_name_birth, first_name, birth_date, keywords_json)
      VALUES (?, ?, ?, ?, ?)
    `,
      [
        uniqueId,
        lastNameBirth.trim(),
        firstName.trim(),
        iso,
        JSON.stringify(
          Array.isArray(keywords)
            ? keywords.map((k) => k.toString().trim()).filter(Boolean)
            : []
        )
      ]
    );

    const personId = insertPerson.lastID;
    const categories = await all(`SELECT id FROM categories WHERE required = 0`);
    for (const category of categories) {
      const value = customValues[category.id] ?? customValues[String(category.id)] ?? '';
      if (value !== undefined && value !== null && String(value).trim() !== '') {
        await run(
          `INSERT INTO person_category_values (person_id, category_id, value) VALUES (?, ?, ?)`,
          [personId, category.id, String(value).trim()]
        );
      }
    }

    const person = await get(`SELECT * FROM people WHERE id = ?`, [personId]);
    return res.status(201).json(person);
  } catch (error) {
    if (error.message.includes('UNIQUE constraint failed: people.unique_id')) {
      return res.status(409).json({
        error: "Une personne avec cet identifiant unique existe déjà."
      });
    }
    return res.status(500).json({ error: error.message });
  }
});

app.get('/api/people', async (req, res) => {
  try {
    const people = await listPeople(req.query.search || '');
    return res.json(people);
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

app.post('/api/import/preview', upload.single('excelFile'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Fichier Excel manquant.' });
  }

  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const headers = rows.length > 0 ? Object.keys(rows[0]) : [];

    return res.json({
      headers,
      sampleRows: rows.slice(0, 5)
    });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  } finally {
    fs.unlink(req.file.path, () => {});
  }
});

app.post('/api/import/commit', upload.single('excelFile'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Fichier Excel manquant.' });
  }

  const mapping = req.body.mapping ? JSON.parse(req.body.mapping) : null;
  if (!mapping) {
    return res.status(400).json({ error: 'Le mapping des colonnes est obligatoire.' });
  }

  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    const categories = await all(`SELECT id, required, code FROM categories`);
    const customCategories = categories.filter((c) => c.required === 0);

    await run('BEGIN TRANSACTION');
    let imported = 0;
    const skipped = [];

    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i];
      const lastNameBirth = row[mapping.last_name_birth] ?? '';
      const firstName = row[mapping.first_name] ?? '';
      const birthDateRaw = row[mapping.birth_date] ?? '';
      const keywordsRaw = mapping.keywords ? row[mapping.keywords] : '';

      if (!lastNameBirth || !firstName || !birthDateRaw) {
        skipped.push({ row: i + 2, reason: 'Champs obligatoires manquants' });
        continue;
      }

      let isoDate;
      try {
        isoDate = formatDatePart(birthDateRaw).iso;
      } catch {
        skipped.push({ row: i + 2, reason: 'Date invalide' });
        continue;
      }

      const uniqueId = buildUniqueId(lastNameBirth, firstName, isoDate);
      try {
        const insertPerson = await run(
          `
            INSERT INTO people (unique_id, last_name_birth, first_name, birth_date, keywords_json)
            VALUES (?, ?, ?, ?, ?)
          `,
          [
            uniqueId,
            String(lastNameBirth).trim(),
            String(firstName).trim(),
            isoDate,
            JSON.stringify(
              String(keywordsRaw)
                .split(',')
                .map((k) => k.trim())
                .filter(Boolean)
            )
          ]
        );

        for (const category of customCategories) {
          const mapKey = `cat_${category.id}`;
          const columnName = mapping[mapKey];
          if (!columnName) continue;
          const value = row[columnName];
          if (value === undefined || value === null || String(value).trim() === '') continue;

          await run(
            `INSERT INTO person_category_values (person_id, category_id, value) VALUES (?, ?, ?)`,
            [insertPerson.lastID, category.id, String(value).trim()]
          );
        }

        imported += 1;
      } catch {
        skipped.push({ row: i + 2, reason: `Doublon identifiant (${uniqueId})` });
      }
    }

    await run('COMMIT');

    return res.json({ imported, skipped });
  } catch (error) {
    await run('ROLLBACK');
    return res.status(500).json({ error: error.message });
  } finally {
    fs.unlink(req.file.path, () => {});
  }
});

initDb()
  .then(() => {
    app.listen(PORT, () => {
      // eslint-disable-next-line no-console
      console.log(`UManager lancé sur http://localhost:${PORT}`);
    });
  })
  .catch((error) => {
    // eslint-disable-next-line no-console
    console.error('Erreur init DB:', error);
    process.exit(1);
  });
