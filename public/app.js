let categories = [];
let importHeaders = [];

const categoriesList = document.getElementById('categories-list');
const customFieldsContainer = document.getElementById('custom-fields');
const personCategoriesOverview = document.getElementById('person-categories-overview');
const peopleBody = document.getElementById('people-body');
const personCreateResult = document.getElementById('person-create-result');
const categoriesCount = document.getElementById('categories-count');

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    headers: { 'Content-Type': 'application/json' },
    ...options
  });

  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || 'Erreur API');
  }

  return response.json();
}

function initTabs() {
  const tabs = document.querySelectorAll('.tab');
  const panels = document.querySelectorAll('.tab-panel');

  tabs.forEach((tab) => {
    tab.addEventListener('click', () => {
      const targetId = tab.dataset.tab;

      tabs.forEach((item) => item.classList.remove('is-active'));
      panels.forEach((panel) => panel.classList.remove('is-active'));

      tab.classList.add('is-active');
      document.getElementById(targetId)?.classList.add('is-active');
    });
  });
}

async function loadCategories() {
  categories = await fetchJson('/api/categories');
  renderCategories();
  if (categoriesCount) {
    categoriesCount.textContent = String(categories.length);
  }
  renderPersonCategoriesOverview();
  renderCustomFields();
  renderImportMapping();
}

function renderCategories() {
  categoriesList.innerHTML = '';

  categories.forEach((category) => {
    const wrap = document.createElement('div');
    wrap.className = 'category-item';

    const left = document.createElement('div');
    left.innerHTML = `<strong>${category.label}</strong>${
      category.required ? '<span class="required">(obligatoire)</span>' : ''
    }`;

    const right = document.createElement('div');
    const input = document.createElement('input');
    input.value = category.label;
    const save = document.createElement('button');
    save.className = 'btn';
    save.textContent = 'Éditer';
    save.onclick = async () => {
      try {
        await fetchJson(`/api/categories/${category.id}`, {
          method: 'PUT',
          body: JSON.stringify({ label: input.value })
        });
        await loadCategories();
      } catch (error) {
        alert(error.message);
      }
    };

    right.appendChild(input);
    right.appendChild(save);

    wrap.appendChild(left);
    wrap.appendChild(right);
    categoriesList.appendChild(wrap);
  });
}

function renderPersonCategoriesOverview() {
  personCategoriesOverview.innerHTML = '';

  categories.forEach((category) => {
    const item = document.createElement('span');
    item.className = `tag ${category.required ? 'tag-default' : 'tag-user'}`;
    item.textContent = `${category.label}${category.required ? ' • défaut' : ' • ajout utilisateur'}`;
    personCategoriesOverview.appendChild(item);
  });
}

function renderCustomFields() {
  customFieldsContainer.innerHTML = '';
  const customCategories = categories.filter((c) => !c.required);

  customCategories.forEach((category) => {
    const label = document.createElement('label');
    label.textContent = category.label;

    const input = document.createElement('input');
    input.id = `cat_${category.id}`;

    label.appendChild(input);
    customFieldsContainer.appendChild(label);
  });
}

async function loadPeople(search = '') {
  const url = new URL('/api/people', window.location.origin);
  if (search.trim()) {
    url.searchParams.set('search', search.trim());
  }

  const people = await fetchJson(url.toString());
  peopleBody.innerHTML = '';

  people.forEach((person) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${person.uniqueId}</td>
      <td>${person.lastNameBirth}</td>
      <td>${person.firstName}</td>
      <td>${person.birthDate}</td>
      <td>${person.keywords.join(', ')}</td>
      <td>${person.customValues
        .map((value) => `<div><strong>${value.categoryLabel}:</strong> ${value.value}</div>`)
        .join('')}</td>
    `;

    peopleBody.appendChild(tr);
  });
}

document.getElementById('add-category-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  const labelInput = document.getElementById('new-category-label');

  try {
    await fetchJson('/api/categories', {
      method: 'POST',
      body: JSON.stringify({ label: labelInput.value })
    });
    labelInput.value = '';
    await loadCategories();
  } catch (error) {
    alert(error.message);
  }
});

document.getElementById('person-form').addEventListener('submit', async (event) => {
  event.preventDefault();

  const keywords = document
    .getElementById('keywords')
    .value.split(',')
    .map((k) => k.trim())
    .filter(Boolean);

  const customValues = {};
  categories
    .filter((c) => !c.required)
    .forEach((category) => {
      customValues[category.id] = document.getElementById(`cat_${category.id}`)?.value || '';
    });

  try {
    const created = await fetchJson('/api/people', {
      method: 'POST',
      body: JSON.stringify({
        lastNameBirth: document.getElementById('last_name_birth').value,
        firstName: document.getElementById('first_name').value,
        birthDate: document.getElementById('birth_date').value,
        keywords,
        customValues
      })
    });

    personCreateResult.textContent = `Personne créée : ${created.unique_id}`;
    document.getElementById('person-form').reset();
    await loadPeople(document.getElementById('search-input').value);
  } catch (error) {
    personCreateResult.textContent = error.message;
  }
});

let searchTimer = null;
document.getElementById('search-input').addEventListener('input', (event) => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    loadPeople(event.target.value).catch((error) => alert(error.message));
  }, 200);
});

async function previewImport() {
  const fileInput = document.getElementById('excel-file');
  const file = fileInput.files[0];
  if (!file) {
    alert('Sélectionnez un fichier Excel.');
    return;
  }

  const formData = new FormData();
  formData.append('excelFile', file);

  const response = await fetch('/api/import/preview', {
    method: 'POST',
    body: formData
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || 'Erreur preview import');
  }

  importHeaders = data.headers;
  renderImportMapping(data.sampleRows);
}

function mappingSelect(id, label, required = false) {
  const requiredMark = required ? '*' : '';
  const options = ['<option value="">-- choisir --</option>']
    .concat(importHeaders.map((h) => `<option value="${h}">${h}</option>`))
    .join('');

  return `<label>${label}${requiredMark}<select id="${id}">${options}</select></label>`;
}

function renderImportMapping(sampleRows = null) {
  const wrap = document.getElementById('import-mapping');

  if (importHeaders.length === 0) {
    wrap.innerHTML = '<p class="info">Analysez un fichier pour configurer le mapping.</p>';
    return;
  }

  const customSelects = categories
    .filter((c) => !c.required)
    .map((c) => mappingSelect(`map_cat_${c.id}`, c.label))
    .join('');

  wrap.innerHTML = `
    <h3>Mapping des colonnes</h3>
    <div class="grid">
      ${mappingSelect('map_last_name_birth', 'Nom de naissance', true)}
      ${mappingSelect('map_first_name', 'Prénom', true)}
      ${mappingSelect('map_birth_date', 'Date de naissance', true)}
      ${mappingSelect('map_keywords', 'Mots-clés (optionnel)')}
      ${customSelects}
    </div>
    ${
      sampleRows
        ? `<details><summary>Aperçu des lignes</summary><pre>${JSON.stringify(sampleRows, null, 2)}</pre></details>`
        : ''
    }
  `;

  document.getElementById('commit-import').hidden = false;
}

async function commitImport() {
  const file = document.getElementById('excel-file').files[0];
  if (!file) {
    alert('Sélectionnez un fichier Excel.');
    return;
  }

  const mapping = {
    last_name_birth: document.getElementById('map_last_name_birth').value,
    first_name: document.getElementById('map_first_name').value,
    birth_date: document.getElementById('map_birth_date').value,
    keywords: document.getElementById('map_keywords').value
  };

  if (!mapping.last_name_birth || !mapping.first_name || !mapping.birth_date) {
    alert('Les mappings obligatoires doivent être renseignés.');
    return;
  }

  categories
    .filter((c) => !c.required)
    .forEach((category) => {
      mapping[`cat_${category.id}`] = document.getElementById(`map_cat_${category.id}`)?.value || '';
    });

  const formData = new FormData();
  formData.append('excelFile', file);
  formData.append('mapping', JSON.stringify(mapping));

  const response = await fetch('/api/import/commit', {
    method: 'POST',
    body: formData
  });
  const data = await response.json();

  if (!response.ok) {
    throw new Error(data.error || 'Erreur import');
  }

  document.getElementById('import-result').textContent = JSON.stringify(data, null, 2);
  await loadPeople(document.getElementById('search-input').value);
}

document.getElementById('preview-import').addEventListener('click', () => {
  previewImport().catch((error) => alert(error.message));
});

document.getElementById('commit-import').addEventListener('click', () => {
  commitImport().catch((error) => alert(error.message));
});

(async function init() {
  initTabs();
  await loadCategories();
  await loadPeople();
})();
