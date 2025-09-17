(() => {
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

  document.addEventListener('DOMContentLoaded', () => {
    const isAuthPage = Boolean(
      document.getElementById('login-form') || document.getElementById('register-form'),
    );
    const isDashboardPage = Boolean(document.querySelector('.app'));

    if (isAuthPage) {
      initializeAuthPage().catch((error) => {
        console.error("Erreur lors de l'initialisation de la page de connexion :", error);
      });
    }

    if (isDashboardPage) {
      initializeDashboardPage().catch((error) => {
        console.error("Erreur lors de l'initialisation du tableau de bord :", error);
      });
    }
  });

  function navigateToDashboard() {
    window.location.replace('dashboard.html');
  }

  function navigateToLogin() {
    window.location.replace('index.html');
  }

  async function initializeAuthPage() {
    await ensureDefaultAdmins();

    const activeUser = loadActiveUser();
    if (activeUser) {
      navigateToDashboard();
      return;
    }

    const loginSection = document.getElementById('login-section');
    const registerSection = document.getElementById('register-section');
    const loginForm = document.getElementById('login-form');
    const registerForm = document.getElementById('register-form');
    const loginError = document.getElementById('login-error');
    const registerError = document.getElementById('register-error');
    const showRegisterBtn = document.getElementById('show-register');
    const showLoginBtn = document.getElementById('show-login');

    function setSectionVisibility(section, isVisible) {
      if (!section) {
        return;
      }

      if (isVisible) {
        section.hidden = false;
        section.removeAttribute('hidden');
        return;
      }

      section.hidden = true;
      if (!section.hasAttribute('hidden')) {
        section.setAttribute('hidden', '');
      }
    }

    function showLoginView() {
      setSectionVisibility(loginSection, true);
      setSectionVisibility(registerSection, false);
      if (loginError) {
        loginError.textContent = '';
      }
      window.requestAnimationFrame(() => {
        const loginUsername = document.getElementById('login-username');
        if (loginUsername && typeof loginUsername.focus === 'function') {
          loginUsername.focus();
        }
      });
    }

    function showRegisterView() {
      setSectionVisibility(loginSection, false);
      setSectionVisibility(registerSection, true);
      if (registerError) {
        registerError.textContent = '';
      }
      window.requestAnimationFrame(() => {
        const registerUsername = document.getElementById('register-username');
        if (registerUsername && typeof registerUsername.focus === 'function') {
          registerUsername.focus();
        }
      });
    }

    function displayLoginError(message) {
      if (loginError) {
        loginError.textContent = message;
      }
    }

    showLoginView();

    if (showRegisterBtn) {
      showRegisterBtn.addEventListener('click', () => {
        if (loginError) {
          loginError.textContent = '';
        }
        if (registerError) {
          registerError.textContent = '';
        }
        showRegisterView();
      });
    }

    if (showLoginBtn) {
      showLoginBtn.addEventListener('click', () => {
        if (loginError) {
          loginError.textContent = '';
        }
        if (registerError) {
          registerError.textContent = '';
        }
        showLoginView();
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
          displayLoginError('Veuillez renseigner vos identifiants.');
          return;
        }

        const store = loadUserStore();
        const user = store.users[username];

        if (!user) {
          displayLoginError("L'utilisateur n'est pas reconnu.");
          return;
        }

        let storedPasswordHash = user.passwordHash;

        if (!storedPasswordHash && typeof user.password === 'string') {
          if (user.password === password) {
            try {
              storedPasswordHash = await hashPassword(password);
            } catch (error) {
              console.error('Impossible de générer le hachage du mot de passe :', error);
              displayLoginError('Une erreur est survenue. Veuillez réessayer.');
              return;
            }
            user.passwordHash = storedPasswordHash;
            delete user.password;
            saveUserStore(store);
          } else {
            displayLoginError('Mot de passe incorrect.');
            return;
          }
        }

        if (!storedPasswordHash) {
          displayLoginError("L'utilisateur n'est pas reconnu.");
          return;
        }

        let passwordHash;
        try {
          passwordHash = await hashPassword(password);
        } catch (error) {
          console.error('Impossible de vérifier le mot de passe :', error);
          displayLoginError('Une erreur est survenue. Veuillez réessayer.');
          return;
        }

        if (passwordHash !== storedPasswordHash) {
          displayLoginError('Mot de passe incorrect.');
          return;
        }

        saveActiveUser(username);
        loginForm.reset();
        navigateToDashboard();
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
        const confirmPassword = (formData.get('password-confirm') || '').toString();

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
            registerError.textContent = 'Le mot de passe doit contenir au moins 6 caractères.';
          }
          return;
        }

        if (password !== confirmPassword) {
          if (registerError) {
            registerError.textContent = 'La confirmation du mot de passe ne correspond pas.';
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
            registerError.textContent = 'Cette adresse mail est déjà associée à un compte.';
          }
          return;
        }

        let passwordHash;
        try {
          passwordHash = await hashPassword(password);
        } catch (error) {
          console.error('Impossible de sécuriser le mot de passe :', error);
          if (registerError) {
            registerError.textContent =
              "Une erreur est survenue lors de la création du compte. Veuillez réessayer.";
          }
          return;
        }

        store.users[username] = {
          email,
          passwordHash,
        };

        saveUserStore(store);
        saveActiveUser(username);
        registerForm.reset();
        navigateToDashboard();
      });
    }
  }

  async function initializeDashboardPage() {
    const currentUser = loadActiveUser();
    if (!currentUser) {
      navigateToLogin();
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

    if (logoutButton) {
      logoutButton.addEventListener('click', () => {
        clearActiveUser();
        navigateToLogin();
      });
    }

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

    if (keywordForm) {
      keywordForm.addEventListener('submit', (event) => {
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
    }

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
      const metrics = data && typeof data === 'object' && data.metrics ? data.metrics : {};

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

          if (editButton) {
            editButton.addEventListener('click', () => {
              startKeywordEdition(keyword.id);
            });
          }

          if (deleteButton) {
            deleteButton.addEventListener('click', () => {
              deleteKeyword(keyword.id);
            });
          }

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

  async function ensureDefaultAdmins() {
    const store = loadUserStore();
    let updated = false;

    for (const admin of DEFAULT_ADMINS) {
      const existing = store.users[admin.username];
      if (!existing) {
        const passwordHash = await hashPassword(admin.password);
        store.users[admin.username] = {
          email: admin.email,
          passwordHash,
        };
        updated = true;
        continue;
      }

      if (!existing.passwordHash) {
        existing.passwordHash = await hashPassword(admin.password);
        updated = true;
      }

      if (!existing.email) {
        existing.email = admin.email;
        updated = true;
      }
    }

    if (updated) {
      saveUserStore(store);
    }
  }

  function loadUserStore() {
    try {
      const stored = window.localStorage.getItem(USER_STORE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        if (parsed && typeof parsed === 'object' && parsed.users) {
          const users = parsed.users && typeof parsed.users === 'object' ? parsed.users : {};
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

  function loadDataForUser(username) {
    if (!username) {
      return cloneDefaultData();
    }

    const storageKey = `${DATA_KEY_PREFIX}${username}`;
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

  function saveDataForUser(username, data) {
    if (!username) {
      return;
    }

    const storageKey = `${DATA_KEY_PREFIX}${username}`;
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

  async function hashPassword(password) {
    const message = `umanager::${password}`;
    const cryptoObj =
      typeof window !== 'undefined' && window.crypto && window.crypto.subtle ? window.crypto : null;

    if (cryptoObj && cryptoObj.subtle && typeof cryptoObj.subtle.digest === 'function') {
      try {
        const encoder = new TextEncoder();
        const dataBuffer = encoder.encode(message);
        const hashBuffer = await cryptoObj.subtle.digest('SHA-256', dataBuffer);
        return bufferToHex(hashBuffer);
      } catch (error) {
        console.warn('Échec du hachage via Web Crypto, utilisation du repli logiciel.', error);
      }
    }

    return sha256Sync(message);
  }

  function bufferToHex(buffer) {
    const hashArray = Array.from(new Uint8Array(buffer));
    return hashArray.map((byte) => byte.toString(16).padStart(2, '0')).join('');
  }

  function sha256Sync(message) {
    const k = new Uint32Array([
      0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1,
      0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3,
      0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786,
      0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
      0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147,
      0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13,
      0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b,
      0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
      0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a,
      0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208,
      0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
    ]);

    const encoder = new TextEncoder();
    const messageBytes = encoder.encode(message);
    const bitLength = BigInt(messageBytes.length) * 8n;
    const paddedLength = (((messageBytes.length + 9 + 63) >> 6) << 6);

    const buffer = new Uint8Array(paddedLength);
    buffer.set(messageBytes);
    buffer[messageBytes.length] = 0x80;

    const view = new DataView(buffer.buffer);
    const highBits = Number((bitLength >> 32n) & 0xffffffffn);
    const lowBits = Number(bitLength & 0xffffffffn);
    view.setUint32(paddedLength - 8, highBits);
    view.setUint32(paddedLength - 4, lowBits);

    let h0 = 0x6a09e667;
    let h1 = 0xbb67ae85;
    let h2 = 0x3c6ef372;
    let h3 = 0xa54ff53a;
    let h4 = 0x510e527f;
    let h5 = 0x9b05688c;
    let h6 = 0x1f83d9ab;
    let h7 = 0x5be0cd19;

    const w = new Uint32Array(64);

    for (let i = 0; i < paddedLength; i += 64) {
      for (let j = 0; j < 16; j += 1) {
        w[j] = view.getUint32(i + j * 4);
      }

      for (let j = 16; j < 64; j += 1) {
        const s0 = rightRotate(w[j - 15], 7) ^ rightRotate(w[j - 15], 18) ^ (w[j - 15] >>> 3);
        const s1 = rightRotate(w[j - 2], 17) ^ rightRotate(w[j - 2], 19) ^ (w[j - 2] >>> 10);
        w[j] = (w[j - 16] + s0 + w[j - 7] + s1) >>> 0;
      }

      let a = h0;
      let b = h1;
      let c = h2;
      let d = h3;
      let e = h4;
      let f = h5;
      let g = h6;
      let h = h7;

      for (let j = 0; j < 64; j += 1) {
        const S1 = rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25);
        const ch = (e & f) ^ (~e & g);
        const temp1 = (h + S1 + ch + k[j] + w[j]) >>> 0;
        const S0 = rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22);
        const maj = (a & b) ^ (a & c) ^ (b & c);
        const temp2 = (S0 + maj) >>> 0;

        h = g;
        g = f;
        f = e;
        e = (d + temp1) >>> 0;
        d = c;
        c = b;
        b = a;
        a = (temp1 + temp2) >>> 0;
      }

      h0 = (h0 + a) >>> 0;
      h1 = (h1 + b) >>> 0;
      h2 = (h2 + c) >>> 0;
      h3 = (h3 + d) >>> 0;
      h4 = (h4 + e) >>> 0;
      h5 = (h5 + f) >>> 0;
      h6 = (h6 + g) >>> 0;
      h7 = (h7 + h) >>> 0;
    }

    const hash = new Uint8Array(32);
    const hashView = new DataView(hash.buffer);
    hashView.setUint32(0, h0);
    hashView.setUint32(4, h1);
    hashView.setUint32(8, h2);
    hashView.setUint32(12, h3);
    hashView.setUint32(16, h4);
    hashView.setUint32(20, h5);
    hashView.setUint32(24, h6);
    hashView.setUint32(28, h7);

    return bufferToHex(hash.buffer);
  }

  function rightRotate(value, amount) {
    return (value >>> amount) | (value << (32 - amount));
  }

  function isValidEmail(value) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
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
})();
