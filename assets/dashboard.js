import {
  clearActiveUser,
  loadActiveUser,
  loadDataForUser,
  saveDataForUser,
} from './storage.js';

document.addEventListener('DOMContentLoaded', () => {
  initializeDashboard().catch((error) => {
    console.error('Erreur lors de l\'initialisation du tableau de bord :', error);
  });
});

async function initializeDashboard() {
  const currentUser = loadActiveUser();
  if (!currentUser) {
    window.location.replace('index.html');
    return;
  }

  const currentUsernameEl = document.getElementById('current-username');
  const logoutButton = document.getElementById('logout-button');
  const pages = Array.from(document.querySelectorAll('.page'));
  const navButtons = Array.from(document.querySelectorAll('.nav-button'));
  const metricValues = Array.from(document.querySelectorAll('[data-metric]'));
  const metricForms = Array.from(document.querySelectorAll('.metric-form'));
  const totalDatasetsEl = document.getElementById('total-datasets');
  const keywordCountEl = document.getElementById('keywords-count');
  const keywordsActiveCountEl = document.getElementById('keywords-active-count');
  const lastUpdatedEl = document.getElementById('last-updated');
  const keywordForm = document.getElementById('keyword-form');
  const keywordList = document.getElementById('keyword-list');
  const keywordEmptyState = document.getElementById('keyword-empty-state');
  const keywordTemplate = document.getElementById('keyword-item-template');

  let data = loadDataForUser(currentUser);

  if (currentUsernameEl) {
    currentUsernameEl.textContent = currentUser;
  }

  logoutButton?.addEventListener('click', () => {
    clearActiveUser();
    window.location.replace('index.html');
  });

  navButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const target = button.dataset.target;
      if (target) {
        showPage(target);
      }
    });
  });

  metricForms.forEach((form) => {
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      const key = form.dataset.form;
      if (!key) {
        return;
      }
      const input = form.querySelector('[data-input]');
      if (!(input instanceof HTMLInputElement)) {
        return;
      }
      const value = Number.parseInt(input.value, 10);
      if (!Number.isFinite(value) || value < 0) {
        input.setCustomValidity('Veuillez renseigner un nombre positif.');
        input.reportValidity();
        return;
      }

      input.setCustomValidity('');
      data.metrics[key] = value;
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderMetrics();
      input.value = '';
    });
  });

  keywordForm?.addEventListener('submit', (event) => {
    event.preventDefault();
    const formData = new FormData(keywordForm);
    const name = (formData.get('keyword-name') || '').toString().trim();
    const description = (formData.get('keyword-description') || '').toString().trim();

    if (!name) {
      const nameInput = keywordForm.querySelector('#keyword-name');
      if (nameInput instanceof HTMLInputElement) {
        nameInput.focus();
      }
      return;
    }

    data.keywords.push({
      id: generateId(),
      name,
      description,
    });

    data.lastUpdated = new Date().toISOString();
    saveDataForUser(currentUser, data);
    keywordForm.reset();
    const nameInput = keywordForm.querySelector('#keyword-name');
    if (nameInput instanceof HTMLInputElement) {
      nameInput.focus();
    }
    renderMetrics();
    renderKeywords();
  });

  showPage('dashboard');
  renderMetrics();
  renderKeywords();

  function showPage(target) {
    navButtons.forEach((button) => {
      button.classList.toggle('active', button.dataset.target === target);
    });
    pages.forEach((page) => {
      page.classList.toggle('active', page.id === target);
    });
  }

  function renderMetrics() {
    const metrics = data?.metrics ?? {};

    metricValues.forEach((metricValue) => {
      const key = metricValue.dataset.metric;
      if (!key) {
        return;
      }
      const value = Number(metrics[key]) || 0;
      metricValue.textContent = new Intl.NumberFormat('fr-FR').format(value);
    });

    if (totalDatasetsEl) {
      const total =
        Number(metrics.peopleCount) + Number(metrics.phoneCount) + Number(metrics.emailCount);
      totalDatasetsEl.textContent = new Intl.NumberFormat('fr-FR').format(total || 0);
    }

    if (keywordCountEl) {
      keywordCountEl.textContent = data.keywords.length.toString();
    }

    if (keywordsActiveCountEl) {
      keywordsActiveCountEl.textContent = data.keywords.length.toString();
    }

    if (lastUpdatedEl) {
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
  }

  function renderKeywords() {
    if (!keywordList || !keywordEmptyState || !keywordTemplate) {
      return;
    }

    keywordList.innerHTML = '';

    if (!data || data.keywords.length === 0) {
      keywordEmptyState.hidden = false;
      return;
    }

    keywordEmptyState.hidden = true;

    const fragment = document.createDocumentFragment();

    data.keywords
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
      .forEach((keyword) => {
        const templateItem = keywordTemplate.content.firstElementChild;
        const listItem = templateItem
          ? templateItem.cloneNode(true)
          : createKeywordListItemFallback();

        listItem.classList.add('keyword-item');
        listItem.dataset.id = keyword.id;

        const titleEl = listItem.querySelector('.keyword-title');
        const descriptionEl = listItem.querySelector('.keyword-description');
        const editButton = listItem.querySelector('[data-action="edit"]');
        const deleteButton = listItem.querySelector('[data-action="delete"]');

        if (titleEl) {
          titleEl.textContent = keyword.name;
        }
        if (descriptionEl) {
          descriptionEl.textContent =
            keyword.description || 'Aucune description renseignée.';
          descriptionEl.classList.toggle('keyword-description--empty', !keyword.description);
        }

        editButton?.addEventListener('click', () => {
          startKeywordEdition(keyword.id);
        });

        deleteButton?.addEventListener('click', () => {
          deleteKeyword(keyword.id);
        });

        fragment.appendChild(listItem);
      });

    keywordList.appendChild(fragment);
  }

  function startKeywordEdition(keywordId) {
    const keyword = data.keywords.find((item) => item.id === keywordId);
    if (!keyword || !keywordList) {
      return;
    }

    const listItem = keywordList.querySelector(`[data-id="${keywordId}"]`);
    if (!listItem) {
      return;
    }

    listItem.classList.add('editing');
    listItem.innerHTML = '';

    const form = document.createElement('form');
    form.className = 'keyword-edit-form';

    const nameRow = document.createElement('div');
    nameRow.className = 'form-row';
    const nameLabel = document.createElement('label');
    const nameInput = document.createElement('input');
    const nameId = `edit-name-${keywordId}`;
    nameLabel.setAttribute('for', nameId);
    nameLabel.textContent = 'Nom du mot clé *';
    nameInput.id = nameId;
    nameInput.name = 'name';
    nameInput.type = 'text';
    nameInput.required = true;
    nameInput.maxLength = 80;
    nameInput.value = keyword.name;
    nameRow.append(nameLabel, nameInput);

    const descriptionRow = document.createElement('div');
    descriptionRow.className = 'form-row';
    const descriptionLabel = document.createElement('label');
    const descriptionInput = document.createElement('textarea');
    const descriptionId = `edit-description-${keywordId}`;
    descriptionLabel.setAttribute('for', descriptionId);
    descriptionLabel.textContent = 'Description';
    descriptionInput.id = descriptionId;
    descriptionInput.name = 'description';
    descriptionInput.rows = 3;
    descriptionInput.maxLength = 240;
    descriptionInput.value = keyword.description || '';
    descriptionRow.append(descriptionLabel, descriptionInput);

    const actionsRow = document.createElement('div');
    actionsRow.className = 'keyword-edit-actions';
    const cancelButton = document.createElement('button');
    cancelButton.type = 'button';
    cancelButton.className = 'keyword-edit-cancel';
    cancelButton.textContent = 'Annuler';
    const saveButton = document.createElement('button');
    saveButton.type = 'submit';
    saveButton.className = 'keyword-edit-save';
    saveButton.textContent = 'Enregistrer';
    actionsRow.append(cancelButton, saveButton);

    form.append(nameRow, descriptionRow, actionsRow);

    form.addEventListener('submit', (event) => {
      event.preventDefault();
      const name = nameInput.value.trim();
      const description = descriptionInput.value.trim();

      if (!name) {
        nameInput.focus();
        return;
      }

      keyword.name = name;
      keyword.description = description;
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderMetrics();
      renderKeywords();
    });

    cancelButton.addEventListener('click', () => {
      renderKeywords();
    });

    listItem.appendChild(form);
    nameInput.focus();
    nameInput.setSelectionRange(nameInput.value.length, nameInput.value.length);
  }

  function deleteKeyword(keywordId) {
    data.keywords = data.keywords.filter((item) => item.id !== keywordId);
    data.lastUpdated = new Date().toISOString();
    saveDataForUser(currentUser, data);
    renderMetrics();
    renderKeywords();
  }
}

function generateId() {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return `kw-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createKeywordListItemFallback() {
  const listItem = document.createElement('li');
  listItem.className = 'keyword-item';

  const keywordMain = document.createElement('div');
  keywordMain.className = 'keyword-main';
  const titleEl = document.createElement('h3');
  titleEl.className = 'keyword-title';
  const descriptionEl = document.createElement('p');
  descriptionEl.className = 'keyword-description';
  keywordMain.append(titleEl, descriptionEl);

  const actions = document.createElement('div');
  actions.className = 'keyword-actions';
  const editButton = document.createElement('button');
  editButton.type = 'button';
  editButton.className = 'keyword-button';
  editButton.dataset.action = 'edit';
  editButton.textContent = 'Modifier';
  const deleteButton = document.createElement('button');
  deleteButton.type = 'button';
  deleteButton.className = 'keyword-button keyword-button--danger';
  deleteButton.dataset.action = 'delete';
  deleteButton.textContent = 'Supprimer';
  actions.append(editButton, deleteButton);

  listItem.append(keywordMain, actions);
  return listItem;
}
