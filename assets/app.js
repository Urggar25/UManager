(function () {
  const STORAGE_KEY = 'umanager-data-store';
  const defaultData = {
    metrics: {
      peopleCount: 0,
      phoneCount: 0,
      emailCount: 0,
    },
    keywords: [],
    lastUpdated: null,
  };

  const pages = document.querySelectorAll('.page');
  const navButtons = document.querySelectorAll('.nav-button');
  const metricValues = document.querySelectorAll('[data-metric]');
  const metricForms = document.querySelectorAll('.metric-form');
  const totalDatasetsEl = document.getElementById('total-datasets');
  const keywordCountEl = document.getElementById('keywords-count');
  const keywordsActiveCountEl = document.getElementById('keywords-active-count');
  const lastUpdatedEl = document.getElementById('last-updated');
  const keywordForm = document.getElementById('keyword-form');
  const keywordList = document.getElementById('keyword-list');
  const keywordEmptyState = document.getElementById('keyword-empty-state');
  const keywordTemplate = document.getElementById('keyword-item-template');

  let data = loadData();
  renderNavigation();
  renderMetrics();
  renderKeywords();

  function loadData() {
    try {
      const stored = window.localStorage.getItem(STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        return {
          metrics: { ...defaultData.metrics, ...(parsed.metrics || {}) },
          keywords: Array.isArray(parsed.keywords) ? parsed.keywords : [],
          lastUpdated: parsed.lastUpdated || null,
        };
      }
    } catch (error) {
      console.warn('Impossible de charger les données locales :', error);
    }
    return JSON.parse(JSON.stringify(defaultData));
  }

  function saveData() {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    } catch (error) {
      console.warn('Impossible de sauvegarder les données locales :', error);
    }
  }

  function renderNavigation() {
    navButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const target = button.dataset.target;
        if (!target) return;

        navButtons.forEach((btn) => btn.classList.toggle('active', btn === button));
        pages.forEach((page) => page.classList.toggle('active', page.id === target));
      });
    });
  }

  function renderMetrics() {
    metricValues.forEach((metricValue) => {
      const key = metricValue.dataset.metric;
      if (!key) return;
      const value = Number(data.metrics[key]) || 0;
      metricValue.textContent = new Intl.NumberFormat('fr-FR').format(value);
    });

    totalDatasetsEl.textContent = new Intl.NumberFormat('fr-FR').format(
      data.metrics.peopleCount + data.metrics.phoneCount + data.metrics.emailCount,
    );

    keywordCountEl.textContent = data.keywords.length.toString();
    keywordsActiveCountEl.textContent = data.keywords.length.toString();

    if (data.lastUpdated) {
      const formatted = new Intl.DateTimeFormat('fr-FR', {
        dateStyle: 'medium',
        timeStyle: 'short',
      }).format(new Date(data.lastUpdated));
      lastUpdatedEl.textContent = formatted;
    } else {
      lastUpdatedEl.textContent = '—';
    }
  }

  metricForms.forEach((form) => {
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      const key = form.dataset.form;
      if (!key) return;
      const input = form.querySelector('[data-input]');
      const value = Number.parseInt(input.value, 10);

      if (Number.isNaN(value) || value < 0) {
        input.setCustomValidity('Veuillez renseigner un nombre positif.');
        input.reportValidity();
        return;
      }

      input.setCustomValidity('');
      data.metrics[key] = value;
      data.lastUpdated = new Date().toISOString();
      saveData();
      renderMetrics();
      input.value = '';
    });
  });

  keywordForm.addEventListener('submit', (event) => {
    event.preventDefault();
    const formData = new FormData(keywordForm);
    const name = formData.get('keyword-name').trim();
    const description = formData.get('keyword-description').trim();

    if (!name) {
      return;
    }

    const keyword = {
      id: generateId(),
      name,
      description,
      createdAt: new Date().toISOString(),
    };

    data.keywords.push(keyword);
    data.lastUpdated = new Date().toISOString();
    saveData();
    keywordForm.reset();
    keywordForm.querySelector('#keyword-name').focus();
    renderMetrics();
    renderKeywords();
  });

  function renderKeywords() {
    keywordList.innerHTML = '';

    if (data.keywords.length === 0) {
      keywordEmptyState.hidden = false;
      return;
    }

    keywordEmptyState.hidden = true;

    data.keywords
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
      .forEach((keyword) => {
        const item = keywordTemplate.content.firstElementChild.cloneNode(true);
        item.dataset.id = keyword.id;
        const titleEl = item.querySelector('.keyword-title');
        const descriptionEl = item.querySelector('.keyword-description');
        titleEl.textContent = keyword.name;
        descriptionEl.textContent = keyword.description || 'Aucune description renseignée.';
        descriptionEl.classList.toggle('keyword-description--empty', !keyword.description);

        item.querySelector('[data-action="edit"]').addEventListener('click', () => {
          startKeywordEdition(keyword.id);
        });

        item.querySelector('[data-action="delete"]').addEventListener('click', () => {
          deleteKeyword(keyword.id);
        });

        keywordList.appendChild(item);
      });
  }

  function startKeywordEdition(keywordId) {
    const keyword = data.keywords.find((item) => item.id === keywordId);
    if (!keyword) return;

    const listItem = keywordList.querySelector(`[data-id="${keywordId}"]`);
    if (!listItem) return;

    listItem.classList.add('editing');
    listItem.innerHTML = '';

    const form = document.createElement('form');
    form.className = 'keyword-edit-form';
    form.innerHTML = `
      <div class="form-row">
        <label for="edit-name-${keywordId}">Nom du mot clé *</label>
        <input id="edit-name-${keywordId}" name="name" type="text" required maxlength="80" value="${escapeHtml(
          keyword.name,
        )}" />
      </div>
      <div class="form-row">
        <label for="edit-description-${keywordId}">Description</label>
        <textarea id="edit-description-${keywordId}" name="description" maxlength="240" rows="3">${escapeHtml(
          keyword.description || '',
        )}</textarea>
      </div>
      <div class="keyword-edit-actions">
        <button type="button" class="keyword-edit-cancel">Annuler</button>
        <button type="submit" class="keyword-edit-save">Enregistrer</button>
      </div>
    `;

    form.addEventListener('submit', (event) => {
      event.preventDefault();
      const formData = new FormData(form);
      const name = formData.get('name').trim();
      const description = formData.get('description').trim();

      if (!name) {
        return;
      }

      keyword.name = name;
      keyword.description = description;
      data.lastUpdated = new Date().toISOString();
      saveData();
      renderMetrics();
      renderKeywords();
    });

    form.querySelector('.keyword-edit-cancel').addEventListener('click', () => {
      renderKeywords();
    });

    listItem.appendChild(form);
    const nameInput = form.querySelector('input[name="name"]');
    nameInput.focus();
    nameInput.setSelectionRange(nameInput.value.length, nameInput.value.length);
  }

  function deleteKeyword(keywordId) {
    data.keywords = data.keywords.filter((item) => item.id !== keywordId);
    data.lastUpdated = new Date().toISOString();
    saveData();
    renderMetrics();
    renderKeywords();
  }

  function generateId() {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) {
      return crypto.randomUUID();
    }
    return `kw-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function escapeHtml(value) {
    const text = document.createTextNode(value);
    const div = document.createElement('div');
    div.appendChild(text);
    return div.innerHTML;
  }
})();
