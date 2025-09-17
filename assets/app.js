(async function () {
  const USER_STORE_KEY = 'umanager-user-store';
  const ACTIVE_USER_KEY = 'umanager-active-user';
  const DATA_KEY_PREFIX = 'umanager-data-store:';
  const DEFAULT_ADMINS = [
    { username: 'admin1', password: 'emsal1', email: 'admin1@umanager.local' },
    { username: 'admin2', password: 'emsal2', email: 'admin2@umanager.local' },
    { username: 'admin3', password: 'emsal3', email: 'admin3@umanager.local' },
  ];

  const defaultData = {
    metrics: {
      peopleCount: 0,
      phoneCount: 0,
      emailCount: 0,
    },
    keywords: [],
    lastUpdated: null,
  };

  const authView = document.getElementById('auth-view');
  const loginSection = document.getElementById('login-section');
  const registerSection = document.getElementById('register-section');
  const loginForm = document.getElementById('login-form');
  const registerForm = document.getElementById('register-form');
  const loginError = document.getElementById('login-error');
  const registerError = document.getElementById('register-error');
  const showRegisterBtn = document.getElementById('show-register');
  const showLoginBtn = document.getElementById('show-login');
  const appContainer = document.querySelector('.app');
  const currentUsernameEl = document.getElementById('current-username');
  const logoutButton = document.getElementById('logout-button');

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

  let currentUser = null;
  let data = null;
  let appHandlersInitialized = false;

  await ensureDefaultAdmins();
  initializeAuth();
  initializeAppHandlers();

  const storedActiveUser = loadActiveUser();
  const userStore = loadUserStore();
  if (storedActiveUser && userStore.users[storedActiveUser]) {
    await activateUser(storedActiveUser);
  } else {
    showLoginView();
  }

  async function ensureDefaultAdmins() {
    const store = loadUserStore();
    let updated = false;

    for (const admin of DEFAULT_ADMINS) {
      if (!store.users[admin.username]) {
        const passwordHash = await hashPassword(admin.password);
        store.users[admin.username] = {
          email: admin.email,
          passwordHash,
        };
        updated = true;
      }
    }

    if (updated) {
      saveUserStore(store);
    }
  }

  function initializeAuth() {
    if (showRegisterBtn) {
      showRegisterBtn.addEventListener('click', () => {
        if (registerError) {
          registerError.textContent = '';
        }
        if (loginError) {
          loginError.textContent = '';
        }
        showRegisterView();
      });
    }

    if (showLoginBtn) {
      showLoginBtn.addEventListener('click', () => {
        if (registerError) {
          registerError.textContent = '';
        }
        if (loginError) {
          loginError.textContent = '';
        }
        showLoginView();
      });
    }

    if (logoutButton) {
      logoutButton.addEventListener('click', () => {
        currentUser = null;
        data = null;
        clearActiveUser();
        if (currentUsernameEl) {
          currentUsernameEl.textContent = '—';
        }
        if (appContainer) {
          appContainer.hidden = true;
        }
        showLoginView();
        if (loginForm) {
          loginForm.reset();
        }
        if (registerForm) {
          registerForm.reset();
        }
        if (loginError) {
          loginError.textContent = '';
        }
        if (registerError) {
          registerError.textContent = '';
        }
      });
    }

    if (loginForm) {
      loginForm.addEventListener('submit', async (event) => {
        event.preventDefault();
        if (loginError) {
          loginError.textContent = '';
        }

        const formData = new FormData(loginForm);
        const username = (formData.get('username') || '').toString().trim();
        const password = (formData.get('password') || '').toString();

        if (!username || !password) {
          if (loginError) {
            loginError.textContent = 'Veuillez renseigner vos identifiants.';
          }
          return;
        }

        const store = loadUserStore();
        const user = store.users[username];
        if (!user) {
          if (loginError) {
            loginError.textContent = 'Identifiant ou mot de passe invalide.';
          }
          return;
        }

        const passwordHash = await hashPassword(password);
        if (passwordHash !== user.passwordHash) {
          if (loginError) {
            loginError.textContent = 'Identifiant ou mot de passe invalide.';
          }
          return;
        }

        loginForm.reset();
        await activateUser(username);
      });
    }

    if (registerForm) {
      registerForm.addEventListener('submit', async (event) => {
        event.preventDefault();
        if (registerError) {
          registerError.textContent = '';
        }

        const formData = new FormData(registerForm);
        const username = (formData.get('username') || '').toString().trim();
        const email = (formData.get('email') || '').toString().trim();
        const password = (formData.get('password') || '').toString();
        const confirmPassword = (formData.get('password-confirm') || '')
          .toString();

        if (!username || !email || !password || !confirmPassword) {
          if (registerError) {
            registerError.textContent = 'Tous les champs sont obligatoires.';
          }
          return;
        }

        if (!isValidEmail(email)) {
          if (registerError) {
            registerError.textContent = 'Veuillez saisir une adresse mail valide.';
          }
          return;
        }

        if (password.length < 6) {
          if (registerError) {
            registerError.textContent =
              'Le mot de passe doit contenir au moins 6 caractères.';
          }
          return;
        }

        if (password !== confirmPassword) {
          if (registerError) {
            registerError.textContent =
              'La confirmation du mot de passe ne correspond pas.';
          }
          return;
        }

        const store = loadUserStore();
        if (store.users[username]) {
          if (registerError) {
            registerError.textContent = 'Cet identifiant est déjà utilisé.';
          }
          return;
        }

        const emailExists = Object.values(store.users).some((user) => {
          if (!user.email) {
            return false;
          }
          return user.email.toLowerCase() === email.toLowerCase();
        });

        if (emailExists) {
          if (registerError) {
            registerError.textContent =
              'Cette adresse mail est déjà associée à un compte.';
          }
          return;
        }

        const passwordHash = await hashPassword(password);
        store.users[username] = {
          email,
          passwordHash,
        };
        saveUserStore(store);
        registerForm.reset();
        await activateUser(username);
      });
    }
  }

  function initializeAppHandlers() {
    if (appHandlersInitialized) {
      return;
    }
    appHandlersInitialized = true;

    navButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const target = button.dataset.target;
        if (!target) {
          return;
        }
        showPage(target);
      });
    });

    metricForms.forEach((form) => {
      form.addEventListener('submit', (event) => {
        event.preventDefault();
        if (!currentUser || !data) {
          return;
        }
        const key = form.dataset.form;
        if (!key) {
          return;
        }
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

    if (keywordForm) {
      keywordForm.addEventListener('submit', (event) => {
        event.preventDefault();
        if (!currentUser || !data) {
          return;
        }
        const formData = new FormData(keywordForm);
        const name = (formData.get('keyword-name') || '').toString().trim();
        const description = (formData.get('keyword-description') || '')
          .toString()
          .trim();

        if (!name) {
          return;
        }

        const keyword = {
          id: generateId(),
          name,
          description,
        };

        data.keywords.push(keyword);
        data.lastUpdated = new Date().toISOString();
        saveData();
        keywordForm.reset();
        const keywordNameInput = keywordForm.querySelector('#keyword-name');
        if (keywordNameInput) {
          keywordNameInput.focus();
        }
        renderMetrics();
        renderKeywords();
      });
    }
  }

  async function activateUser(username) {
    currentUser = username;
    data = loadData();
    saveActiveUser(username);
    updateSidebarUser();
    if (loginError) {
      loginError.textContent = '';
    }
    if (registerError) {
      registerError.textContent = '';
    }
    if (authView) {
      authView.hidden = true;
    }
    if (appContainer) {
      appContainer.hidden = false;
    }
    showPage('dashboard');
    renderMetrics();
    renderKeywords();
  }

  function showLoginView() {
    if (authView) {
      authView.hidden = false;
    }
    if (loginSection) {
      loginSection.hidden = false;
    }
    if (registerSection) {
      registerSection.hidden = true;
    }
    if (loginError) {
      loginError.textContent = '';
    }
    window.requestAnimationFrame(() => {
      const usernameInput = document.getElementById('login-username');
      if (usernameInput) {
        usernameInput.focus();
      }
    });
  }

  function showRegisterView() {
    if (loginSection) {
      loginSection.hidden = true;
    }
    if (registerSection) {
      registerSection.hidden = false;
    }
    if (registerError) {
      registerError.textContent = '';
    }
    window.requestAnimationFrame(() => {
      const usernameInput = document.getElementById('register-username');
      if (usernameInput) {
        usernameInput.focus();
      }
    });
  }

  function renderMetrics() {
    if (!data) {
      metricValues.forEach((metricValue) => {
        metricValue.textContent = '0';
      });
      if (totalDatasetsEl) {
        totalDatasetsEl.textContent = '0';
      }
      if (keywordCountEl) {
        keywordCountEl.textContent = '0';
      }
      if (keywordsActiveCountEl) {
        keywordsActiveCountEl.textContent = '0';
      }
      if (lastUpdatedEl) {
        lastUpdatedEl.textContent = '—';
      }
      return;
    }

    metricValues.forEach((metricValue) => {
      const key = metricValue.dataset.metric;
      if (!key) {
        return;
      }
      const value = Number(data.metrics[key]) || 0;
      metricValue.textContent = new Intl.NumberFormat('fr-FR').format(value);
    });

    if (totalDatasetsEl) {
      totalDatasetsEl.textContent = new Intl.NumberFormat('fr-FR').format(
        data.metrics.peopleCount + data.metrics.phoneCount + data.metrics.emailCount,
      );
    }

    if (keywordCountEl) {
      keywordCountEl.textContent = data.keywords.length.toString();
    }

    if (keywordsActiveCountEl) {
      keywordsActiveCountEl.textContent = data.keywords.length.toString();
    }

    if (data.lastUpdated && lastUpdatedEl) {
      const formatted = new Intl.DateTimeFormat('fr-FR', {
        dateStyle: 'medium',
        timeStyle: 'short',
      }).format(new Date(data.lastUpdated));
      lastUpdatedEl.textContent = formatted;
    } else if (lastUpdatedEl) {
      lastUpdatedEl.textContent = '—';
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
        descriptionEl.classList.toggle(
          'keyword-description--empty',
          !keyword.description,
        );

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
    if (!data) {
      return;
    }
    const keyword = data.keywords.find((item) => item.id === keywordId);
    if (!keyword) {
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
      if (!data) {
        return;
      }
      const formData = new FormData(form);
      const name = (formData.get('name') || '').toString().trim();
      const description = (formData.get('description') || '').toString().trim();

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
    if (nameInput) {
      nameInput.focus();
      nameInput.setSelectionRange(nameInput.value.length, nameInput.value.length);
    }
  }

  function deleteKeyword(keywordId) {
    if (!data) {
      return;
    }
    data.keywords = data.keywords.filter((item) => item.id !== keywordId);
    data.lastUpdated = new Date().toISOString();
    saveData();
    renderMetrics();
    renderKeywords();
  }

  function showPage(target) {
    navButtons.forEach((button) => {
      button.classList.toggle('active', button.dataset.target === target);
    });
    pages.forEach((page) => {
      page.classList.toggle('active', page.id === target);
    });
  }

  function updateSidebarUser() {
    if (currentUsernameEl) {
      currentUsernameEl.textContent = currentUser || '—';
    }
  }

  function loadData() {
    if (!currentUser) {
      return cloneDefaultData();
    }

    const storageKey = `${DATA_KEY_PREFIX}${currentUser}`;
    try {
      const stored = window.localStorage.getItem(storageKey);
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

    return cloneDefaultData();
  }

  function saveData() {
    if (!currentUser || !data) {
      return;
    }

    const storageKey = `${DATA_KEY_PREFIX}${currentUser}`;
    try {
      window.localStorage.setItem(storageKey, JSON.stringify(data));
    } catch (error) {
      console.warn('Impossible de sauvegarder les données locales :', error);
    }
  }

  function cloneDefaultData() {
    return {
      metrics: { ...defaultData.metrics },
      keywords: [],
      lastUpdated: null,
    };
  }

  function loadUserStore() {
    try {
      const stored = window.localStorage.getItem(USER_STORE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        if (parsed && typeof parsed === 'object' && parsed.users) {
          const users =
            parsed.users && typeof parsed.users === 'object' ? parsed.users : {};
          return {
            users: { ...users },
          };
        }
      }
    } catch (error) {
      console.warn('Impossible de charger les comptes utilisateurs :', error);
    }
    return { users: {} };
  }

  function saveUserStore(store) {
    try {
      window.localStorage.setItem(USER_STORE_KEY, JSON.stringify(store));
    } catch (error) {
      console.warn('Impossible de sauvegarder les comptes utilisateurs :', error);
    }
  }

  function saveActiveUser(username) {
    try {
      window.localStorage.setItem(ACTIVE_USER_KEY, username);
    } catch (error) {
      console.warn("Impossible d'enregistrer l'utilisateur actif :", error);
    }
  }

  function loadActiveUser() {
    try {
      return window.localStorage.getItem(ACTIVE_USER_KEY);
    } catch (error) {
      console.warn("Impossible de charger l'utilisateur actif :", error);
      return null;
    }
  }

  function clearActiveUser() {
    try {
      window.localStorage.removeItem(ACTIVE_USER_KEY);
    } catch (error) {
      console.warn("Impossible de réinitialiser l'utilisateur actif :", error);
    }
  }

  function isValidEmail(value) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
  }

  async function hashPassword(password) {
    const encoder = new TextEncoder();
    const dataBuffer = encoder.encode(password);
    const hashBuffer = await crypto.subtle.digest('SHA-256', dataBuffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map((byte) => byte.toString(16).padStart(2, '0')).join('');
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
