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
    categories: [],
    keywords: [],
    contacts: [],
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
    const metricShares = Array.from(document.querySelectorAll('[data-metric-share]'));
    const coverageChartEl = document.getElementById('contact-coverage-chart');
    const coverageCountEls = {
      both: document.querySelector('[data-coverage-count="both"]'),
      phone: document.querySelector('[data-coverage-count="phone"]'),
      email: document.querySelector('[data-coverage-count="email"]'),
      none: document.querySelector('[data-coverage-count="none"]'),
    };
    const coveragePercentEls = {
      both: document.querySelector('[data-coverage-percent="both"]'),
      phone: document.querySelector('[data-coverage-percent="phone"]'),
      email: document.querySelector('[data-coverage-percent="email"]'),
      none: document.querySelector('[data-coverage-percent="none"]'),
    };
    const totalDatasetsEl = document.getElementById('total-datasets');
    const categoriesCountEl = document.getElementById('categories-count');
    const keywordsCountEl = document.getElementById('keywords-count');
    const categoriesActiveCountEl = document.getElementById('categories-active-count');
    const keywordsActiveCountEl = document.getElementById('keywords-active-count');
    const lastUpdatedEl = document.getElementById('last-updated');
    const categoryForm = document.getElementById('category-form');
    const categoryList = document.getElementById('category-list');
    const categoryEmptyState = document.getElementById('category-empty-state');
    const categoryTemplate = document.getElementById('category-item-template');
    const categoryTypeSelect = document.getElementById('category-type');
    const categoryOptionsRow = document.getElementById('category-options-row');
    const categoryOptionsInput = document.getElementById('category-options');
    const keywordForm = document.getElementById('keyword-form');
    const keywordList = document.getElementById('keyword-list');
    const keywordEmptyState = document.getElementById('keyword-empty-state');
    const keywordTemplate = document.getElementById('keyword-item-template');
    const contactsCountEl = document.getElementById('contacts-count');
    const contactForm = document.getElementById('contact-form');
    const contactCategoryFieldsContainer = document.getElementById('contact-category-fields');
    const contactCategoriesEmpty = document.getElementById('contact-categories-empty');
    const contactKeywordsContainer = document.getElementById('contact-keywords-container');
    const contactKeywordsEmpty = document.getElementById('contact-keywords-empty');
    const contactList = document.getElementById('contact-list');
    const contactEmptyState = document.getElementById('contact-empty-state');
    const contactSearchInput = document.getElementById('contact-search-input');
    const contactAdvancedSearchForm = document.getElementById('contact-advanced-search');
    const searchCategoryFieldsContainer = document.getElementById('search-category-fields');
    const searchCategoriesEmpty = document.getElementById('search-categories-empty');
    const searchKeywordsSelect = document.getElementById('search-keywords');
    const contactSearchCountEl = document.getElementById('contact-search-count');
    const contactTemplate = document.getElementById('contact-item-template');
    const contactSubmitButton = document.getElementById('contact-submit-button');
    const contactCancelEditButton = document.getElementById('contact-cancel-edit');
    const contactBackToSearchButton = document.getElementById('contact-back-to-search');
    const contactsAddTitle = document.getElementById('contacts-add-title');
    const contactsAddSubtitle = document.querySelector('#contacts-add .page-subtitle');
    const contactsAddTitleDefault = contactsAddTitle ? contactsAddTitle.textContent : '';
    const contactsAddSubtitleDefault = contactsAddSubtitle ? contactsAddSubtitle.textContent : '';

    const CATEGORY_TYPE_ORDER = ['text', 'number', 'date', 'list'];
    const CATEGORY_TYPES = new Set(CATEGORY_TYPE_ORDER);
    const CATEGORY_TYPE_LABELS = {
      text: 'Texte',
      number: 'Nombre',
      date: 'Date',
      list: 'Liste',
    };
    const EMAIL_KEYWORDS = ['mail', 'email', 'courriel', 'mel'];
    const PHONE_KEYWORDS = ['tel', 'telephone', 'mobile', 'portable', 'phone', 'gsm'];

    let data = loadDataForUser(currentUser);
    data = upgradeDataStructure(data);

    const numberFormatter = new Intl.NumberFormat('fr-FR');
    const percentFormatter = new Intl.NumberFormat('fr-FR', {
      style: 'percent',
      maximumFractionDigits: 1,
    });

    let contactSearchTerm = '';
    let advancedFilters = createEmptyAdvancedFilters();
    let contactEditReturnPage = 'contacts-search';
    let categoryDragAndDropInitialized = false;

    normalizeCategoryOrders();

    if (currentUsernameEl) {
      currentUsernameEl.textContent = currentUser;
    }

    if (logoutButton) {
      logoutButton.addEventListener('click', () => {
        clearActiveUser();
        navigateToLogin();
      });
    }

    if (categoryList && !categoryDragAndDropInitialized) {
      initializeCategoryDragAndDrop();
      categoryDragAndDropInitialized = true;
    }

    updateCategoryOptionsVisibility();
    if (categoryTypeSelect) {
      categoryTypeSelect.addEventListener('change', () => {
        updateCategoryOptionsVisibility();
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

    if (categoryForm) {
      categoryForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const formData = new FormData(categoryForm);
        const name = (formData.get('category-name') || '').toString().trim();
        const description = (formData.get('category-description') || '').toString().trim();
        let typeValue = (formData.get('category-type') || 'text').toString();
        if (!CATEGORY_TYPES.has(typeValue)) {
          typeValue = 'text';
        }

        let options = [];
        if (categoryOptionsInput instanceof HTMLTextAreaElement) {
          categoryOptionsInput.setCustomValidity('');
        }

        if (typeValue === 'list') {
          const rawOptions = (formData.get('category-options') || '').toString();
          options = parseCategoryOptions(rawOptions);
          if (categoryOptionsInput instanceof HTMLTextAreaElement) {
            if (options.length === 0) {
              categoryOptionsInput.setCustomValidity('Veuillez renseigner au moins une valeur.');
              categoryOptionsInput.reportValidity();
              return;
            }
            categoryOptionsInput.setCustomValidity('');
          }
        }

        if (!name) {
          const nameInput = categoryForm.querySelector('#category-name');
          if (nameInput instanceof HTMLInputElement) {
            nameInput.focus();
          }
          return;
        }

        data.categories.push({
          id: generateId('category'),
          name,
          description,
          type: typeValue,
          options,
          order: getNextCategoryOrderValue(),
        });

        normalizeCategoryOrders();

        data.lastUpdated = new Date().toISOString();
        saveDataForUser(currentUser, data);
        categoryForm.reset();
        updateCategoryOptionsVisibility();
        const nameInput = categoryForm.querySelector('#category-name');
        if (nameInput instanceof HTMLInputElement) {
          nameInput.focus();
        }
        renderMetrics();
        renderCategories();
      });

      categoryForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          if (categoryOptionsInput instanceof HTMLTextAreaElement) {
            categoryOptionsInput.setCustomValidity('');
          }
          updateCategoryOptionsVisibility();
        });
      });
    }

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
          id: generateId('keyword'),
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

    if (contactForm) {
      contactForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const formData = new FormData(contactForm);
        const notes = (formData.get('contact-notes') || '').toString().trim();
        const keywordIds = formData.getAll('contact-keywords').map((value) => value.toString());

        const categoryInputs = contactForm.querySelectorAll('[data-category-input]');
        const categoryValues = {};
        let hasInvalidCategoryValue = false;

        categoryInputs.forEach((element) => {
          if (
            !(
              element instanceof HTMLInputElement ||
              element instanceof HTMLTextAreaElement ||
              element instanceof HTMLSelectElement
            )
          ) {
            return;
          }

          element.setCustomValidity('');
          if (!element.checkValidity()) {
            element.reportValidity();
            hasInvalidCategoryValue = true;
            return;
          }

          const categoryId = element.dataset.categoryId || '';
          if (!categoryId) {
            return;
          }

          const type = element.dataset.categoryType || 'text';
          const rawValue = element.value != null ? element.value.toString().trim() : '';
          if (!rawValue) {
            return;
          }

          if (type === 'number') {
            const parsed = Number(element.value);
            if (Number.isFinite(parsed)) {
              categoryValues[categoryId] = parsed.toString();
            } else {
              element.setCustomValidity('Veuillez renseigner un nombre valide.');
              element.reportValidity();
              hasInvalidCategoryValue = true;
            }
            return;
          }

          categoryValues[categoryId] = element.value.toString().trim();
        });

        if (hasInvalidCategoryValue) {
          return;
        }

        const editingId = contactForm.dataset.editingId || '';
        const editingIndex = editingId
          ? data.contacts.findIndex((contact) => contact.id === editingId)
          : -1;

        const categoriesById = buildCategoryMap();
        const derivedName = buildDisplayNameFromCategories(categoryValues, categoriesById);
        const nowIso = new Date().toISOString();

        if (editingId && editingIndex !== -1) {
          const contactToUpdate = data.contacts[editingIndex];
          const previousName = getContactDisplayName(contactToUpdate, categoriesById);
          contactToUpdate.categoryValues = { ...categoryValues };
          contactToUpdate.keywords = keywordIds;
          contactToUpdate.notes = notes;
          if (derivedName) {
            contactToUpdate.fullName = derivedName;
            contactToUpdate.displayName = derivedName;
          } else if (previousName) {
            contactToUpdate.fullName = previousName;
            contactToUpdate.displayName = previousName;
          } else {
            contactToUpdate.fullName = 'Contact sans nom';
            contactToUpdate.displayName = 'Contact sans nom';
          }
          contactToUpdate.updatedAt = nowIso;
        } else {
          const displayName = derivedName || 'Contact sans nom';
          data.contacts.push({
            id: generateId('contact'),
            categoryValues: { ...categoryValues },
            keywords: keywordIds,
            notes,
            fullName: displayName,
            displayName,
            createdAt: nowIso,
            updatedAt: null,
          });
        }

        data.lastUpdated = new Date().toISOString();
        updateMetricsFromContacts();
        saveDataForUser(currentUser, data);
        resetContactForm();
        renderMetrics();
        renderContacts();
      });
    }

    if (contactCancelEditButton) {
      contactCancelEditButton.addEventListener('click', () => {
        resetContactForm();
      });
    }

    if (contactBackToSearchButton) {
      contactBackToSearchButton.addEventListener('click', () => {
        const targetPage = contactEditReturnPage || 'contacts-search';
        resetContactForm(false);
        showPage(targetPage);
        renderContacts();
      });
    }

    if (contactSearchInput) {
      contactSearchTerm = contactSearchInput.value || '';
      contactSearchInput.addEventListener('input', () => {
        contactSearchTerm = contactSearchInput.value;
        renderContacts();
      });
    }

    if (contactAdvancedSearchForm) {
      contactAdvancedSearchForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const formData = new FormData(contactAdvancedSearchForm);
        const categoryFilters = {};

        if (searchCategoryFieldsContainer) {
          const filterInputs = searchCategoryFieldsContainer.querySelectorAll(
            '[data-search-category-input]',
          );
          filterInputs.forEach((element) => {
            if (
              !(
                element instanceof HTMLInputElement ||
                element instanceof HTMLSelectElement ||
                element instanceof HTMLTextAreaElement
              )
            ) {
              return;
            }

            const categoryId = element.dataset.categoryId || '';
            if (!categoryId) {
              return;
            }

            const type = element.dataset.categoryType || 'text';
            const rawValue = element.value != null ? element.value.toString() : '';
            const trimmedValue = type === 'text' ? rawValue.trim() : rawValue;
            if (!trimmedValue) {
              return;
            }

            if (type === 'text') {
              categoryFilters[categoryId] = {
                type: 'text',
                rawValue: trimmedValue,
                normalizedValue: trimmedValue.toLowerCase(),
              };
            } else {
              categoryFilters[categoryId] = {
                type,
                rawValue: trimmedValue,
                normalizedValue: trimmedValue,
              };
            }
          });
        }

        const keywordFilters = formData.getAll('search-keywords').map((value) => value.toString());

        advancedFilters = {
          categories: categoryFilters,
          keywords: keywordFilters,
        };
        renderContacts();
      });

      contactAdvancedSearchForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          advancedFilters = createEmptyAdvancedFilters();
          renderSearchCategoryFields();
          renderContacts();
        });
      });
    }

    showPage('dashboard');
    renderMetrics();
    renderCategories();
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
      const contacts = Array.isArray(data.contacts) ? data.contacts : [];
      const totalContacts = contacts.length;
      const categoriesById = buildCategoryMap();
      const coverage = computeContactCoverage(contacts, categoriesById);

      metricValues.forEach((metricValue) => {
        const key = metricValue.dataset.metric;
        if (!key) {
          return;
        }
        const value = Number(metrics[key]) || 0;
        metricValue.textContent = numberFormatter.format(value);
      });

      metricShares.forEach((element) => {
        const key = element.dataset.metricShare;
        if (!key) {
          return;
        }

        let ratio = 0;
        if (totalContacts > 0) {
          if (key === 'phone') {
            ratio = (coverage.both + coverage.phoneOnly) / totalContacts;
          } else if (key === 'email') {
            ratio = (coverage.both + coverage.emailOnly) / totalContacts;
          }
        }

        const formattedShare = totalContacts > 0 ? percentFormatter.format(ratio) : '0 %';
        element.textContent = `${formattedShare} des contacts`;
      });

      updateCoverageVisualization(coverage, totalContacts);

      if (totalDatasetsEl) {
        const total =
          Number(metrics.peopleCount) + Number(metrics.phoneCount) + Number(metrics.emailCount);
        totalDatasetsEl.textContent = numberFormatter.format(total || 0);
      }

      if (categoriesCountEl) {
        categoriesCountEl.textContent = data.categories.length.toString();
      }

      if (keywordsCountEl) {
        keywordsCountEl.textContent = data.keywords.length.toString();
      }

      if (categoriesActiveCountEl) {
        categoriesActiveCountEl.textContent = data.categories.length.toString();
      }

      if (keywordsActiveCountEl) {
        keywordsActiveCountEl.textContent = data.keywords.length.toString();
      }

      if (contactsCountEl) {
        contactsCountEl.textContent = numberFormatter.format(totalContacts);
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

    function updateCoverageVisualization(coverage, totalContacts) {
      if (coverageChartEl) {
        if (!totalContacts) {
          coverageChartEl.style.background = 'var(--coverage-none)';
          coverageChartEl.setAttribute('aria-label', 'Aucune donnée de contact disponible.');
        } else {
          const segments = [
            { key: 'both', value: coverage.both, color: 'var(--coverage-both)' },
            { key: 'phone', value: coverage.phoneOnly, color: 'var(--coverage-phone)' },
            { key: 'email', value: coverage.emailOnly, color: 'var(--coverage-email)' },
            { key: 'none', value: coverage.none, color: 'var(--coverage-none)' },
          ];

          let currentAngle = 0;
          const gradientParts = [];

          segments.forEach((segment) => {
            if (!segment.value) {
              return;
            }
            const sweep = (segment.value / totalContacts) * 360;
            const start = currentAngle;
            const end = currentAngle + sweep;
            gradientParts.push(`${segment.color} ${start}deg ${end}deg`);
            currentAngle = end;
          });

          if (gradientParts.length === 0) {
            coverageChartEl.style.background = 'var(--coverage-none)';
          } else {
            coverageChartEl.style.background = `conic-gradient(${gradientParts.join(', ')})`;
          }

          const summaryParts = [
            formatCoveragePhrase('avec téléphone et e-mail', coverage.both),
            formatCoveragePhrase('avec uniquement téléphone', coverage.phoneOnly),
            formatCoveragePhrase('avec uniquement e-mail', coverage.emailOnly),
            formatCoveragePhrase('sans coordonnées', coverage.none),
          ].filter(Boolean);

          if (summaryParts.length > 0) {
            coverageChartEl.setAttribute(
              'aria-label',
              `Répartition des contacts : ${summaryParts.join(', ')}.`,
            );
          } else {
            coverageChartEl.setAttribute('aria-label', 'Répartition des contacts.');
          }
        }
      }

      Object.entries(coverageCountEls).forEach(([key, element]) => {
        if (!element) {
          return;
        }
        const value = Number(coverage[key]) || 0;
        element.textContent = formatContactCount(value);
      });

      Object.entries(coveragePercentEls).forEach(([key, element]) => {
        if (!element) {
          return;
        }
        const value = Number(coverage[key]) || 0;
        element.textContent = formatPercentValue(value, totalContacts);
      });
    }

    function formatCoveragePhrase(label, value) {
      if (!value) {
        return '';
      }
      return `${formatContactCount(value)} ${label}`;
    }

    function formatContactCount(value) {
      const absolute = Number(value) || 0;
      const label = absolute === 1 ? 'contact' : 'contacts';
      return `${numberFormatter.format(absolute)} ${label}`;
    }

    function formatPercentValue(value, total) {
      if (!total) {
        return '0 %';
      }
      const ratio = value / total;
      return percentFormatter.format(ratio);
    }

    function getCategoryOrderValue(category) {
      if (!category || typeof category.order !== 'number' || Number.isNaN(category.order)) {
        return Number.MAX_SAFE_INTEGER;
      }
      return category.order;
    }

    function sortCategoriesForDisplay(source = data.categories) {
      const list = Array.isArray(source)
        ? source.filter((item) => item && typeof item === 'object')
        : [];
      list.sort((a, b) => {
        const orderA = getCategoryOrderValue(a);
        const orderB = getCategoryOrderValue(b);
        if (orderA !== orderB) {
          return orderA - orderB;
        }
        const nameA = (a && a.name ? a.name : '').toString();
        const nameB = (b && b.name ? b.name : '').toString();
        return nameA.localeCompare(nameB, 'fr', { sensitivity: 'base' });
      });
      return list;
    }

    function getNextCategoryOrderValue() {
      if (!Array.isArray(data.categories) || data.categories.length === 0) {
        return 0;
      }
      return (
        data.categories.reduce((max, category) => {
          const value = getCategoryOrderValue(category);
          if (!Number.isFinite(value) || value === Number.MAX_SAFE_INTEGER) {
            return max;
          }
          return value > max ? value : max;
        }, -1) + 1
      );
    }

    function normalizeCategoryOrders() {
      if (!Array.isArray(data.categories)) {
        data.categories = [];
        return;
      }

      const sorted = sortCategoriesForDisplay(data.categories);
      sorted.forEach((category, index) => {
        category.order = index;
      });
      data.categories = sorted;
    }

    function buildCategoryMap(target = data) {
      const source = target && typeof target === 'object' ? target : data;
      return new Map(
        Array.isArray(source.categories)
          ? source.categories.map((category) => [category.id, category])
          : [],
      );
    }

    function buildDisplayNameFromCategories(categoryValues, categoriesById) {
      if (!categoryValues || typeof categoryValues !== 'object') {
        return '';
      }

      const categoriesSource = Array.isArray(data.categories) && data.categories.length > 0
        ? data.categories
        : Array.from(categoriesById instanceof Map ? categoriesById.values() : []);
      const orderedCategories = sortCategoriesForDisplay(categoriesSource);
      const collectedValues = [];

      orderedCategories.forEach((category) => {
        if (!category || !category.id || collectedValues.length >= 2) {
          return;
        }

        const rawValue = categoryValues[category.id];
        const valueString =
          rawValue === undefined || rawValue === null ? '' : rawValue.toString().trim();
        if (!valueString) {
          return;
        }

        collectedValues.push(valueString);
      });

      if (collectedValues.length > 0) {
        return collectedValues.join(' ');
      }

      const fallbackValue = Object.values(categoryValues).find((rawValue) => {
        const valueString =
          rawValue === undefined || rawValue === null ? '' : rawValue.toString().trim();
        return Boolean(valueString);
      });

      return fallbackValue ? fallbackValue.toString().trim() : '';
    }

    function getContactDisplayName(contact, categoriesById = buildCategoryMap()) {
      if (!contact || typeof contact !== 'object') {
        return 'Contact sans nom';
      }

      const categoryValues =
        contact.categoryValues && typeof contact.categoryValues === 'object'
          ? contact.categoryValues
          : {};
      const derivedName = buildDisplayNameFromCategories(categoryValues, categoriesById);
      if (derivedName) {
        return derivedName;
      }

      const directName = (contact.displayName || contact.fullName || '').toString().trim();
      if (directName) {
        return directName;
      }

      const legacyName = `${contact.firstName || ''} ${contact.usageName || ''}`.trim();
      if (legacyName) {
        return legacyName;
      }

      return 'Contact sans nom';
    }

    function normalizeLabel(value) {
      if (!value) {
        return '';
      }
      return value
        .toString()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase();
    }

    function isEmailValue(value) {
      if (!value) {
        return false;
      }
      const normalized = value.toString().trim();
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized);
    }

    function isPhoneValue(value) {
      if (!value) {
        return false;
      }
      const digits = value.toString().replace(/\D+/g, '');
      return digits.length >= 6;
    }

    function computeContactChannels(contact, categoriesById = buildCategoryMap()) {
      const emailSet = new Set();
      const phoneSet = new Set();

      const addEmail = (raw) => {
        const normalized = raw != null ? raw.toString().trim() : '';
        if (!normalized || !isEmailValue(normalized)) {
          return;
        }
        emailSet.add(normalized);
      };

      const addPhone = (raw) => {
        const normalized = raw != null ? raw.toString().trim() : '';
        if (!normalized || !isPhoneValue(normalized)) {
          return;
        }
        phoneSet.add(normalized);
      };

      if (contact && typeof contact === 'object') {
        addEmail(contact.email);
        addPhone(contact.mobile);
        addPhone(contact.landline);
        addPhone(contact.phone);
      }

      const categoryValues =
        contact && typeof contact === 'object' && contact.categoryValues && typeof contact.categoryValues === 'object'
          ? contact.categoryValues
          : {};

      Object.entries(categoryValues).forEach(([categoryId, rawValue]) => {
        const category = categoriesById.get(categoryId);
        if (!category) {
          return;
        }

        const value = rawValue === undefined || rawValue === null ? '' : rawValue.toString().trim();
        if (!value) {
          return;
        }

        const normalizedName = normalizeLabel(category.name || '');
        const nameSuggestsPhone = PHONE_KEYWORDS.some((keyword) => normalizedName.includes(keyword));
        const looksLikeEmail = isEmailValue(value);
        const looksLikePhone = isPhoneValue(value);

        if (looksLikeEmail) {
          addEmail(value);
          return;
        }

        if (looksLikePhone && nameSuggestsPhone) {
          addPhone(value);
        }
      });

      return {
        emails: Array.from(emailSet),
        phones: Array.from(phoneSet),
      };
    }

    function computeContactCoverage(contacts, categoriesById = buildCategoryMap()) {
      const coverage = {
        both: 0,
        phoneOnly: 0,
        emailOnly: 0,
        none: 0,
      };

      contacts.forEach((contact) => {
        const channels = computeContactChannels(contact, categoriesById);
        const hasPhone = channels.phones.length > 0;
        const hasEmail = channels.emails.length > 0;

        if (hasPhone && hasEmail) {
          coverage.both += 1;
        } else if (hasPhone) {
          coverage.phoneOnly += 1;
        } else if (hasEmail) {
          coverage.emailOnly += 1;
        } else {
          coverage.none += 1;
        }
      });

      return coverage;
    }

    function renderCategories() {
      if (!categoryList || !categoryEmptyState) {
        renderContactCategoryFields();
        renderSearchCategoryFields();
        renderContacts();
        return;
      }

      categoryList.innerHTML = '';

      const categories = sortCategoriesForDisplay();

      if (categories.length === 0) {
        categoryEmptyState.hidden = false;
        renderContactCategoryFields();
        renderSearchCategoryFields();
        renderContacts();
        return;
      }

      categoryEmptyState.hidden = true;

      const fragment = document.createDocumentFragment();

      categories.forEach((category) => {
        const templateItem = categoryTemplate ? categoryTemplate.content.firstElementChild : null;
        const listItem = templateItem
          ? templateItem.cloneNode(true)
          : createCategoryListItemFallback();

        listItem.classList.add('category-item');
        if (category.id) {
          listItem.dataset.id = category.id;
        }
        if (typeof category.order === 'number' && !Number.isNaN(category.order)) {
          listItem.dataset.order = category.order.toString();
        } else {
          delete listItem.dataset.order;
        }

        const dragHandle = listItem.querySelector('[data-drag-handle]');
        const titleEl = listItem.querySelector('.category-title');
        const descriptionEl = listItem.querySelector('.category-description');
        const metaEl = listItem.querySelector('.category-meta');
        const editButton = listItem.querySelector('[data-action="edit"]');
        const deleteButton = listItem.querySelector('[data-action="delete"]');

        if (dragHandle instanceof HTMLElement) {
          const categoryName = (category.name || '').toString();
          const dragLabel = categoryName
            ? `Réorganiser la catégorie ${categoryName}`
            : 'Réorganiser la catégorie';
          const dragTitle = categoryName
            ? `Déplacer la catégorie ${categoryName}`
            : 'Déplacer la catégorie';
          dragHandle.setAttribute('aria-label', dragLabel);
          dragHandle.setAttribute('title', dragTitle);
          dragHandle.setAttribute('draggable', 'true');
          dragHandle.draggable = true;
        }

        if (titleEl) {
          titleEl.textContent = category.name;
        }

        if (descriptionEl) {
          descriptionEl.textContent = category.description || 'Aucune description renseignée.';
          descriptionEl.classList.toggle('category-description--empty', !category.description);
        }

        if (metaEl) {
          const typeKey = CATEGORY_TYPES.has(category.type) ? category.type : 'text';
          const typeLabel = CATEGORY_TYPE_LABELS[typeKey] || CATEGORY_TYPE_LABELS.text;
          let metaText = `Type : ${typeLabel}`;
          if (typeKey === 'list') {
            const options = Array.isArray(category.options) ? category.options.filter((value) => value) : [];
            metaText += options.length > 0 ? ` · ${options.join(', ')}` : ' · Aucune valeur définie';
          }
          metaEl.textContent = metaText;
        }

        if (editButton) {
          editButton.addEventListener('click', () => {
            startCategoryEdition(category.id);
          });
        }

        if (deleteButton) {
          deleteButton.addEventListener('click', () => {
            deleteCategory(category.id);
          });
        }

        fragment.appendChild(listItem);
      });

      categoryList.appendChild(fragment);
      renderContactCategoryFields();
      renderSearchCategoryFields();
      renderContacts();
    }

    function initializeCategoryDragAndDrop() {
      if (!categoryList) {
        return;
      }

      let draggingItem = null;
      let previousOrder = [];

      categoryList.addEventListener('dragstart', (event) => {
        const handle =
          event.target instanceof HTMLElement ? event.target.closest('[data-drag-handle]') : null;
        if (!handle) {
          event.preventDefault();
          return;
        }

        const listItem = handle.closest('.category-item');
        if (!listItem) {
          event.preventDefault();
          return;
        }

        if (listItem.classList.contains('editing')) {
          event.preventDefault();
          return;
        }

        draggingItem = listItem;
        previousOrder = Array.from(categoryList.querySelectorAll('.category-item')).map(
          (item) => item.dataset.id || '',
        );
        listItem.classList.add('dragging');

        if (event.dataTransfer) {
          event.dataTransfer.effectAllowed = 'move';
          try {
            event.dataTransfer.setData('text/plain', listItem.dataset.id || '');
            event.dataTransfer.setDragImage(listItem, 16, 16);
          } catch (error) {
            // setDragImage peut échouer selon le navigateur, on ignore alors l'erreur.
          }
        }
      });

      categoryList.addEventListener('dragover', (event) => {
        if (!draggingItem) {
          return;
        }
        event.preventDefault();
        const afterElement = getDragAfterElement(categoryList, event.clientY);
        if (!afterElement) {
          categoryList.appendChild(draggingItem);
        } else if (afterElement !== draggingItem) {
          categoryList.insertBefore(draggingItem, afterElement);
        }
      });

      categoryList.addEventListener('drop', (event) => {
        if (draggingItem) {
          event.preventDefault();
        }
      });

      categoryList.addEventListener('dragend', () => {
        if (!draggingItem) {
          return;
        }
        draggingItem.classList.remove('dragging');
        draggingItem = null;
        persistCategoryOrderFromDom(previousOrder);
        previousOrder = [];
      });
    }

    function persistCategoryOrderFromDom(previousOrder = []) {
      if (!categoryList) {
        return;
      }

      const orderedIds = Array.from(categoryList.querySelectorAll('.category-item'))
        .map((item) => item.dataset.id || '')
        .filter((id) => Boolean(id));

      if (orderedIds.length === 0) {
        renderCategories();
        return;
      }

      const sanitizedPrevious = Array.isArray(previousOrder)
        ? previousOrder.filter((id) => Boolean(id))
        : [];
      const hasChanged =
        orderedIds.length !== sanitizedPrevious.length ||
        orderedIds.some((id, index) => id !== sanitizedPrevious[index]);

      if (!hasChanged) {
        renderCategories();
        return;
      }

      const categoriesById = new Map(
        Array.isArray(data.categories)
          ? data.categories.map((category) => [category.id, category])
          : [],
      );
      const reordered = [];

      orderedIds.forEach((categoryId) => {
        const category = categoriesById.get(categoryId);
        if (category) {
          reordered.push(category);
          categoriesById.delete(categoryId);
        }
      });

      categoriesById.forEach((category) => {
        reordered.push(category);
      });

      reordered.forEach((category, index) => {
        category.order = index;
      });

      data.categories = reordered;
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderMetrics();
      renderCategories();
    }

    function getDragAfterElement(container, clientY) {
      const items = Array.from(container.querySelectorAll('.category-item:not(.dragging)'));
      let closest = { offset: Number.NEGATIVE_INFINITY, element: null };

      items.forEach((child) => {
        const box = child.getBoundingClientRect();
        const offset = clientY - box.top - box.height / 2;
        if (offset < 0 && offset > closest.offset) {
          closest = { offset, element: child };
        }
      });

      return closest.element;
    }

    function renderKeywords() {
      if (!keywordList || !keywordEmptyState) {
        renderContactKeywordOptions();
        renderSearchKeywordOptions();
        renderContacts();
        return;
      }

      keywordList.innerHTML = '';

      const keywords = Array.isArray(data.keywords)
        ? data.keywords.slice().sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
        : [];

      if (keywords.length === 0) {
        keywordEmptyState.hidden = false;
        renderContactKeywordOptions();
        renderSearchKeywordOptions();
        renderContacts();
        return;
      }

      keywordEmptyState.hidden = true;

      const fragment = document.createDocumentFragment();

      keywords.forEach((keyword) => {
        const templateItem = keywordTemplate ? keywordTemplate.content.firstElementChild : null;
        const listItem = templateItem
          ? templateItem.cloneNode(true)
          : createKeywordListItemFallback();

        listItem.classList.add('keyword-item');
        if (keyword.id) {
          listItem.dataset.id = keyword.id;
        }

        const titleEl = listItem.querySelector('.keyword-title');
        const descriptionEl = listItem.querySelector('.keyword-description');
        const editButton = listItem.querySelector('[data-action="edit"]');
        const deleteButton = listItem.querySelector('[data-action="delete"]');

        if (titleEl) {
          titleEl.textContent = keyword.name;
        }

        if (descriptionEl) {
          descriptionEl.textContent = keyword.description || 'Aucune description renseignée.';
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
      renderContactKeywordOptions();
      renderSearchKeywordOptions();
      renderContacts();
    }

    function renderContactCategoryFields() {
      const categories = sortCategoriesForDisplay();

      if (!contactCategoryFieldsContainer) {
        if (contactCategoriesEmpty) {
          contactCategoriesEmpty.hidden = categories.length > 0;
        }
        return;
      }

      const previousValues = new Map();
      contactCategoryFieldsContainer.querySelectorAll('[data-category-input]').forEach((element) => {
        if (
          !element ||
          !(element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement || element instanceof HTMLSelectElement)
        ) {
          return;
        }
        const categoryId = element.dataset.categoryId || '';
        if (categoryId) {
          previousValues.set(categoryId, element.value);
        }
      });

      contactCategoryFieldsContainer.innerHTML = '';

      if (categories.length === 0) {
        if (contactCategoriesEmpty) {
          contactCategoriesEmpty.hidden = false;
        }
        return;
      }

      if (contactCategoriesEmpty) {
        contactCategoriesEmpty.hidden = true;
      }

      const editingId = contactForm && contactForm.dataset.editingId ? contactForm.dataset.editingId : '';
      let editingValues = null;
      if (editingId) {
        const editingContact = data.contacts.find((contact) => contact.id === editingId);
        if (editingContact && editingContact.categoryValues && typeof editingContact.categoryValues === 'object') {
          editingValues = editingContact.categoryValues;
        }
      }

      const fragment = document.createDocumentFragment();

      categories.forEach((category) => {
        const fieldWrapper = document.createElement('div');
        fieldWrapper.className = 'form-row';

        const fieldId = `contact-category-${category.id}`;
        const label = document.createElement('label');
        label.setAttribute('for', fieldId);
        label.textContent = category.name;

        const description = (category.description || '').toString().trim();
        let input;

        const baseType = CATEGORY_TYPES.has(category.type) ? category.type : 'text';

        if (baseType === 'number') {
          const numberInput = document.createElement('input');
          numberInput.type = 'number';
          numberInput.inputMode = 'decimal';
          numberInput.id = fieldId;
          numberInput.name = `contact-category-${category.id}`;
          if (description) {
            numberInput.placeholder = description;
          }
          input = numberInput;
        } else if (baseType === 'date') {
          const dateInput = document.createElement('input');
          dateInput.type = 'date';
          dateInput.id = fieldId;
          dateInput.name = `contact-category-${category.id}`;
          input = dateInput;
        } else if (baseType === 'list') {
          const select = document.createElement('select');
          select.id = fieldId;
          select.name = `contact-category-${category.id}`;

          const defaultOption = document.createElement('option');
          defaultOption.value = '';
          defaultOption.textContent = 'Sélectionnez une valeur';
          select.appendChild(defaultOption);

          const options = Array.isArray(category.options) ? category.options : [];
          options.forEach((optionValue) => {
            const option = document.createElement('option');
            option.value = optionValue;
            option.textContent = optionValue;
            select.appendChild(option);
          });
          input = select;
        } else {
          const textInput = document.createElement('input');
          textInput.type = 'text';
          textInput.id = fieldId;
          textInput.name = `contact-category-${category.id}`;
          textInput.maxLength = 200;
          if (description) {
            textInput.placeholder = description;
          }
          input = textInput;
        }

        input.dataset.categoryId = category.id || '';
        input.dataset.categoryType = baseType;
        input.setAttribute('data-category-input', 'true');

        let initialValue = '';
        if (category.id && previousValues.has(category.id)) {
          initialValue = previousValues.get(category.id) || '';
        } else if (category.id && editingValues && category.id in editingValues) {
          const rawValue = editingValues[category.id];
          initialValue = rawValue != null ? rawValue.toString() : '';
        }

        if (input instanceof HTMLSelectElement) {
          const options = Array.from(input.options).map((option) => option.value);
          if (options.includes(initialValue)) {
            input.value = initialValue;
          } else {
            input.value = '';
          }
        } else if (input instanceof HTMLInputElement || input instanceof HTMLTextAreaElement) {
          input.value = initialValue;
        }

        fieldWrapper.append(label, input);

        if (description && (baseType === 'list' || baseType === 'date')) {
          const hint = document.createElement('p');
          hint.className = 'form-hint';
          hint.textContent = description;
          fieldWrapper.appendChild(hint);
        }

        fragment.appendChild(fieldWrapper);
      });

      contactCategoryFieldsContainer.appendChild(fragment);
    }

    function renderContactKeywordOptions() {
      const keywords = Array.isArray(data.keywords)
        ? data.keywords.slice().sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
        : [];

      if (contactKeywordsContainer) {
        const previousSelection = new Set();
        contactKeywordsContainer
          .querySelectorAll('input[name="contact-keywords"]')
          .forEach((input) => {
            if (input instanceof HTMLInputElement && input.checked) {
              previousSelection.add(input.value);
            }
          });

        const editingContactId = contactForm && contactForm.dataset.editingId ? contactForm.dataset.editingId : '';
        const editingContact = editingContactId
          ? data.contacts.find((contact) => contact.id === editingContactId)
          : null;
        const editingKeywords = editingContact && Array.isArray(editingContact.keywords)
          ? new Set(editingContact.keywords)
          : new Set();

        contactKeywordsContainer.innerHTML = '';

        if (keywords.length === 0) {
          if (contactKeywordsEmpty) {
            contactKeywordsEmpty.hidden = false;
          }
        } else {
          if (contactKeywordsEmpty) {
            contactKeywordsEmpty.hidden = true;
          }

          const fragment = document.createDocumentFragment();

          keywords.forEach((keyword) => {
            const checkboxId = `contact-keyword-${keyword.id}`;
            const item = document.createElement('div');
            item.className = 'checkbox-item';

            const input = document.createElement('input');
            input.type = 'checkbox';
            input.id = checkboxId;
            input.name = 'contact-keywords';
            input.value = keyword.id || '';
            if (keyword.id && (previousSelection.has(keyword.id) || editingKeywords.has(keyword.id))) {
              input.checked = true;
            }

            const label = document.createElement('label');
            label.setAttribute('for', checkboxId);
            label.textContent = keyword.name;

            item.append(input, label);
            fragment.appendChild(item);
          });

          contactKeywordsContainer.appendChild(fragment);
        }
      } else if (contactKeywordsEmpty) {
        contactKeywordsEmpty.hidden = keywords.length > 0;
      }
    }

    function renderSearchCategoryFields() {
      if (!searchCategoryFieldsContainer) {
        if (
          advancedFilters.categories &&
          typeof advancedFilters.categories === 'object' &&
          Object.keys(advancedFilters.categories).length > 0
        ) {
          advancedFilters = { ...advancedFilters, categories: {} };
        }
        return;
      }

      const categories = sortCategoriesForDisplay();

      searchCategoryFieldsContainer.innerHTML = '';

      if (categories.length === 0) {
        if (searchCategoriesEmpty) {
          searchCategoriesEmpty.hidden = false;
        }
        return;
      }

      if (searchCategoriesEmpty) {
        searchCategoriesEmpty.hidden = true;
      }

      if (advancedFilters.categories && typeof advancedFilters.categories === 'object') {
        const allowedIds = new Set(categories.map((category) => category.id || ''));
        const filteredCategories = {};
        let hasChanges = false;
        Object.entries(advancedFilters.categories).forEach(([categoryId, filter]) => {
          if (allowedIds.has(categoryId)) {
            filteredCategories[categoryId] = filter;
          } else {
            hasChanges = true;
          }
        });

        if (hasChanges) {
          advancedFilters = { ...advancedFilters, categories: filteredCategories };
        }
      }

      const fragment = document.createDocumentFragment();

      categories.forEach((category) => {
        const fieldWrapper = document.createElement('div');
        fieldWrapper.className = 'form-row';

        const fieldId = `search-category-${category.id}`;
        const label = document.createElement('label');
        label.setAttribute('for', fieldId);
        label.textContent = category.name;

        let input;
        const description = (category.description || '').toString().trim();
        const baseType = CATEGORY_TYPES.has(category.type) ? category.type : 'text';

        if (baseType === 'number') {
          const numberInput = document.createElement('input');
          numberInput.type = 'number';
          numberInput.inputMode = 'decimal';
          numberInput.id = fieldId;
          numberInput.name = `search-category-${category.id}`;
          if (description) {
            numberInput.placeholder = description;
          }
          input = numberInput;
        } else if (baseType === 'date') {
          const dateInput = document.createElement('input');
          dateInput.type = 'date';
          dateInput.id = fieldId;
          dateInput.name = `search-category-${category.id}`;
          input = dateInput;
        } else if (baseType === 'list') {
          const select = document.createElement('select');
          select.id = fieldId;
          select.name = `search-category-${category.id}`;

          const defaultOption = document.createElement('option');
          defaultOption.value = '';
          defaultOption.textContent = 'Toutes les valeurs';
          select.appendChild(defaultOption);

          const options = Array.isArray(category.options) ? category.options : [];
          options.forEach((optionValue) => {
            const option = document.createElement('option');
            option.value = optionValue;
            option.textContent = optionValue;
            select.appendChild(option);
          });
          input = select;
        } else {
          const textInput = document.createElement('input');
          textInput.type = 'text';
          textInput.id = fieldId;
          textInput.name = `search-category-${category.id}`;
          textInput.maxLength = 200;
          if (description) {
            textInput.placeholder = description;
          }
          input = textInput;
        }

        input.dataset.categoryId = category.id || '';
        input.dataset.categoryType = baseType;
        input.setAttribute('data-search-category-input', 'true');

        const storedFilter =
          advancedFilters.categories && category.id && advancedFilters.categories[category.id]
            ? advancedFilters.categories[category.id]
            : null;
        if (storedFilter && typeof storedFilter === 'object') {
          const rawValue = storedFilter.rawValue != null ? storedFilter.rawValue.toString() : '';
          if (baseType === 'list') {
            input.value = rawValue;
          } else if (rawValue) {
            input.value = rawValue;
          }
        }

        fieldWrapper.append(label, input);
        fragment.appendChild(fieldWrapper);
      });

      searchCategoryFieldsContainer.appendChild(fragment);
    }

    function renderSearchKeywordOptions() {
      if (!searchKeywordsSelect) {
        if (Array.isArray(advancedFilters.keywords) && advancedFilters.keywords.length > 0) {
          advancedFilters = { ...advancedFilters, keywords: [] };
        }
        return;
      }

      const keywords = Array.isArray(data.keywords)
        ? data.keywords.slice().sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
        : [];

      const previousSelection = Array.isArray(advancedFilters.keywords)
        ? advancedFilters.keywords
        : [];
      const selectedValues = new Set(previousSelection);

      searchKeywordsSelect.innerHTML = '';

      keywords.forEach((keyword) => {
        const option = document.createElement('option');
        option.value = keyword.id || '';
        option.textContent = keyword.name;
        option.selected = selectedValues.has(keyword.id);
        searchKeywordsSelect.appendChild(option);
      });

      const allowedValues = new Set(keywords.map((keyword) => keyword.id || ''));
      const filteredSelection = previousSelection.filter((value) => allowedValues.has(value));

      if (filteredSelection.length !== previousSelection.length) {
        advancedFilters = { ...advancedFilters, keywords: filteredSelection };
      }
    }

    function renderContacts() {
      const contacts = Array.isArray(data.contacts) ? data.contacts.slice() : [];

      if (contactSearchCountEl) {
        contactSearchCountEl.textContent = '0';
      }

      if (!contactList || !contactEmptyState) {
        return;
      }

      contactList.innerHTML = '';

      const normalizedTerm = contactSearchTerm.trim().toLowerCase();
      const categoriesById = buildCategoryMap();
      const keywordsById = new Map(
        Array.isArray(data.keywords)
          ? data.keywords.map((keyword) => [keyword.id, keyword])
          : [],
      );

      const categoryFilterEntries =
        advancedFilters &&
        advancedFilters.categories &&
        typeof advancedFilters.categories === 'object'
          ? Object.entries(advancedFilters.categories)
          : [];

      const keywordFilters = Array.isArray(advancedFilters.keywords)
        ? advancedFilters.keywords.filter((value) => Boolean(value))
        : [];

      const hasAdvancedFilters = categoryFilterEntries.length > 0 || keywordFilters.length > 0;

      const filteredContacts = contacts
        .filter((contact) => {
          const categoryValues =
            contact && typeof contact === 'object' && contact.categoryValues && typeof contact.categoryValues === 'object'
              ? contact.categoryValues
              : {};
          const keywords = Array.isArray(contact.keywords) ? contact.keywords : [];
          for (const [categoryId, filter] of categoryFilterEntries) {
            if (!filter || typeof filter !== 'object') {
              continue;
            }
            const rawValue = categoryValues[categoryId];
            const valueString =
              rawValue === undefined || rawValue === null ? '' : rawValue.toString().trim();
            if (!valueString) {
              return false;
            }

            if (filter.type === 'text') {
              const normalizedFilter = (filter.normalizedValue || '').toString();
              if (!valueString.toLowerCase().includes(normalizedFilter)) {
                return false;
              }
            } else if (valueString !== filter.rawValue) {
              return false;
            }
          }

          if (keywordFilters.length > 0) {
            const keywordSet = new Set(keywords);
            const matchesKeywords = keywordFilters.every((keywordId) => keywordSet.has(keywordId));
            if (!matchesKeywords) {
              return false;
            }
          }

          if (!normalizedTerm) {
            return true;
          }

          const displayName = getContactDisplayName(contact, categoriesById);
          const channels = computeContactChannels(contact, categoriesById);

          const categoryEntries = Object.entries(categoryValues)
            .map(([categoryId, rawValue]) => {
              const category = categoriesById.get(categoryId);
              if (!category) {
                return null;
              }
              const value =
                rawValue === undefined || rawValue === null ? '' : rawValue.toString().trim();
              return value ? { name: category.name, value } : null;
            })
            .filter((entry) => Boolean(entry));

          const keywordNames = keywords
            .map((keywordId) => {
              const keyword = keywordsById.get(keywordId);
              return keyword ? keyword.name : '';
            })
            .filter(Boolean);

          const haystackParts = [
            displayName,
            contact.fullName || '',
            contact.firstName || '',
            contact.usageName || '',
            contact.birthName || '',
            contact.notes || '',
            contact.organization || '',
            contact.city || '',
            contact.customField || '',
            keywordNames.join(' '),
            categoryEntries.map((entry) => entry && entry.name).join(' '),
            categoryEntries.map((entry) => (entry ? entry.value : '')).join(' '),
            channels.emails.join(' '),
            channels.phones.join(' '),
          ];

          const haystack = haystackParts.join(' ').toLowerCase();
          return haystack.includes(normalizedTerm);
        })
        .sort((a, b) => {
          const nameA = getContactDisplayName(a, categoriesById).toString();
          const nameB = getContactDisplayName(b, categoriesById).toString();
          return nameA.localeCompare(nameB, 'fr', { sensitivity: 'base' });
        });

      if (contactSearchCountEl) {
        contactSearchCountEl.textContent = filteredContacts.length.toString();
      }

      if (filteredContacts.length === 0) {
        contactEmptyState.hidden = false;
        if (contacts.length === 0) {
          contactEmptyState.textContent = 'Ajoutez vos premiers contacts pour les retrouver ici.';
        } else if (normalizedTerm || hasAdvancedFilters) {
          contactEmptyState.textContent = 'Aucun contact ne correspond à vos critères de recherche.';
        } else {
          contactEmptyState.textContent = 'Aucun contact à afficher pour le moment.';
        }
        return;
      }

      contactEmptyState.hidden = true;

      const fragment = document.createDocumentFragment();
      const dateTimeFormatter = new Intl.DateTimeFormat('fr-FR', {
        dateStyle: 'medium',
        timeStyle: 'short',
      });

      filteredContacts.forEach((contact) => {
        const templateItem = contactTemplate && contactTemplate.content
          ? contactTemplate.content.firstElementChild
          : null;
        const listItem = templateItem
          ? templateItem.cloneNode(true)
          : createContactListItemFallback();

        listItem.classList.add('contact-item');
        if (contact.id) {
          listItem.dataset.id = contact.id;
        } else {
          delete listItem.dataset.id;
        }

        const nameEl = listItem.querySelector('.contact-name');
        if (nameEl) {
          const displayName = getContactDisplayName(contact, categoriesById) || 'Contact sans nom';
          nameEl.textContent = displayName;
        }

        const createdEl = listItem.querySelector('.contact-created');
        if (createdEl) {
          const createdValue = contact.createdAt ? new Date(contact.createdAt) : null;
          if (createdValue && !Number.isNaN(createdValue.getTime())) {
            createdEl.textContent = `Ajouté le ${dateTimeFormatter.format(createdValue)}`;
          } else {
            createdEl.textContent = '';
          }
        }

        const coordinatesEl = listItem.querySelector('.contact-coordinates');
        if (coordinatesEl) {
          const channels = computeContactChannels(contact, categoriesById);
          const coordinates = [...channels.phones, ...channels.emails];
          if (coordinates.length > 0) {
            coordinatesEl.textContent = coordinates.join(' · ');
            coordinatesEl.classList.remove('contact-coordinates--empty');
          } else {
            coordinatesEl.textContent = 'Aucune coordonnée renseignée.';
            coordinatesEl.classList.add('contact-coordinates--empty');
          }
        }

        const notesEl = listItem.querySelector('.contact-notes');
        if (notesEl) {
          const notesValue = (contact.notes || '').toString().trim();
          if (notesValue) {
            notesEl.textContent = notesValue;
            notesEl.classList.remove('contact-notes--empty');
          } else {
            notesEl.textContent = 'Aucune note supplémentaire.';
            notesEl.classList.add('contact-notes--empty');
          }
        }

        const categoriesContainer = listItem.querySelector('.contact-categories');
        if (categoriesContainer) {
          categoriesContainer.innerHTML = '';
          const categoryValues =
            contact && typeof contact === 'object' && contact.categoryValues && typeof contact.categoryValues === 'object'
              ? contact.categoryValues
              : {};

          const associatedCategories = Object.entries(categoryValues)
            .map(([categoryId, rawValue]) => {
              const category = categoriesById.get(categoryId);
              if (!category) {
                return null;
              }
              const valueString = rawValue === undefined || rawValue === null ? '' : rawValue.toString().trim();
              return {
                name: category.name,
                value: valueString,
                order: getCategoryOrderValue(category),
              };
            })
            .filter((entry) => entry && entry.value);

          if (associatedCategories.length > 0) {
            associatedCategories
              .sort((a, b) => {
                if (a.order !== b.order) {
                  return a.order - b.order;
                }
                return a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' });
              })
              .forEach((entry) => {
                if (!entry) {
                  return;
                }
                const chip = document.createElement('span');
                chip.className = 'category-chip';
                chip.textContent = `${entry.name} : ${entry.value}`;
                categoriesContainer.appendChild(chip);
              });
          } else {
            const chip = document.createElement('span');
            chip.className = 'category-chip category-chip--empty';
            chip.textContent = 'Aucune donnée personnalisée';
            categoriesContainer.appendChild(chip);
          }
        }

        const keywordsContainer = listItem.querySelector('.contact-keywords');
        if (keywordsContainer) {
          keywordsContainer.innerHTML = '';
          const associatedKeywords = Array.isArray(contact.keywords)
            ? contact.keywords
                .map((keywordId) => keywordsById.get(keywordId))
                .filter((keyword) => Boolean(keyword))
            : [];

          if (associatedKeywords.length > 0) {
            associatedKeywords
              .sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
              .forEach((keyword) => {
                const chip = document.createElement('span');
                chip.className = 'keyword-chip';
                chip.textContent = keyword.name;
                keywordsContainer.appendChild(chip);
              });
          } else {
            const chip = document.createElement('span');
            chip.className = 'keyword-chip keyword-chip--empty';
            chip.textContent = 'Sans mot clé';
            keywordsContainer.appendChild(chip);
          }
        }

        const editContactButton = listItem.querySelector('[data-action="edit-contact"]');
        if (editContactButton instanceof HTMLButtonElement) {
          if (contact.id) {
            editContactButton.disabled = false;
            editContactButton.addEventListener('click', () => {
              startContactEdition(contact.id);
            });
          } else {
            editContactButton.disabled = true;
          }
        }

        const deleteContactButton = listItem.querySelector('[data-action="delete-contact"]');
        if (deleteContactButton instanceof HTMLButtonElement) {
          if (contact.id) {
            deleteContactButton.disabled = false;
            deleteContactButton.addEventListener('click', () => {
              deleteContact(contact.id);
            });
          } else {
            deleteContactButton.disabled = true;
          }
        }

        fragment.appendChild(listItem);
      });

      contactList.appendChild(fragment);
    }

    function resetContactForm(shouldFocus = true) {
      if (!contactForm) {
        return;
      }

      contactForm.reset();
      exitContactEditMode();
      renderContactCategoryFields();

      if (contactKeywordsContainer) {
        contactKeywordsContainer
          .querySelectorAll('input[name="contact-keywords"]')
          .forEach((input) => {
            if (input instanceof HTMLInputElement) {
              input.checked = false;
            }
          });
      }

      if (shouldFocus) {
        const firstCategoryInput = contactCategoryFieldsContainer
          ? contactCategoryFieldsContainer.querySelector('[data-category-input]')
          : null;
        if (
          firstCategoryInput &&
          'focus' in firstCategoryInput &&
          typeof firstCategoryInput.focus === 'function'
        ) {
          firstCategoryInput.focus();
        } else {
          const notesField = contactForm.querySelector('#contact-notes');
          if (
            notesField &&
            'focus' in notesField &&
            typeof notesField.focus === 'function'
          ) {
            notesField.focus();
          }
        }
      }
    }

    function exitContactEditMode() {
      if (!contactForm) {
        return;
      }

      delete contactForm.dataset.editingId;
      if (contactSubmitButton) {
        contactSubmitButton.textContent = 'Ajouter le contact';
      }
      if (contactCancelEditButton) {
        contactCancelEditButton.hidden = true;
      }
      if (contactBackToSearchButton) {
        contactBackToSearchButton.hidden = true;
      }
      if (contactsAddTitle) {
        contactsAddTitle.textContent = contactsAddTitleDefault;
      }
      if (contactsAddSubtitle) {
        contactsAddSubtitle.textContent = contactsAddSubtitleDefault;
      }
      contactEditReturnPage = 'contacts-search';
    }

    function startContactEdition(contactId) {
      if (!contactForm) {
        return;
      }

      const contact = data.contacts.find((item) => item.id === contactId);
      if (!contact) {
        return;
      }

      const activePage = pages.find((page) => page.classList.contains('active'));
      contactEditReturnPage = activePage ? activePage.id : 'contacts-search';

      showPage('contacts-add');
      contactForm.reset();
      contactForm.dataset.editingId = contactId;
      if (contactSubmitButton) {
        contactSubmitButton.textContent = 'Enregistrer les modifications';
      }
      if (contactCancelEditButton) {
        contactCancelEditButton.hidden = false;
      }
      if (contactBackToSearchButton) {
        contactBackToSearchButton.hidden = false;
      }
      if (contactsAddTitle) {
        contactsAddTitle.textContent = 'Modifier un contact';
      }
      if (contactsAddSubtitle) {
        contactsAddSubtitle.textContent =
          'Mettez à jour les informations existantes et enregistrez vos changements.';
      }

      const assignValue = (selector, value) => {
        const element = contactForm.querySelector(selector);
        if (
          element instanceof HTMLInputElement ||
          element instanceof HTMLSelectElement ||
          element instanceof HTMLTextAreaElement
        ) {
          element.value = value != null ? value : '';
        }
      };

      assignValue('#contact-notes', contact.notes || '');

      renderContactCategoryFields();

      const keywordIds = Array.isArray(contact.keywords) ? new Set(contact.keywords) : new Set();
      if (contactKeywordsContainer) {
        contactKeywordsContainer
          .querySelectorAll('input[name="contact-keywords"]')
          .forEach((input) => {
            if (input instanceof HTMLInputElement) {
              input.checked = keywordIds.has(input.value);
            }
          });
      }

      const firstCategoryInput = contactCategoryFieldsContainer
        ? contactCategoryFieldsContainer.querySelector('[data-category-input]')
        : null;
      if (
        firstCategoryInput &&
        'focus' in firstCategoryInput &&
        typeof firstCategoryInput.focus === 'function'
      ) {
        firstCategoryInput.focus();
      } else {
        const notesField = contactForm.querySelector('#contact-notes');
        if (
          notesField &&
          'focus' in notesField &&
          typeof notesField.focus === 'function'
        ) {
          notesField.focus();
        }
      }
    }

    function deleteContact(contactId) {
      if (!contactId || !Array.isArray(data.contacts)) {
        return;
      }

      const contactIndex = data.contacts.findIndex((contact) => contact && contact.id === contactId);
      if (contactIndex === -1) {
        return;
      }

      const isEditingCurrentContact =
        contactForm && contactForm.dataset.editingId === contactId;

      data.contacts.splice(contactIndex, 1);
      data.lastUpdated = new Date().toISOString();
      updateMetricsFromContacts();
      saveDataForUser(currentUser, data);

      if (isEditingCurrentContact) {
        const targetPage = contactEditReturnPage || 'contacts-search';
        resetContactForm(false);
        showPage(targetPage);
      }

      renderMetrics();
      renderContacts();
    }

    function startCategoryEdition(categoryId) {
      const category = data.categories.find((item) => item.id === categoryId);
      if (!category || !categoryList) {
        return;
      }

      const listItem = categoryList.querySelector(`[data-id="${categoryId}"]`);
      if (!listItem) {
        return;
      }

      listItem.classList.add('editing');
      listItem.innerHTML = '';

      const form = document.createElement('form');
      form.className = 'category-edit-form';

      const nameRow = document.createElement('div');
      nameRow.className = 'form-row';
      const nameLabel = document.createElement('label');
      const nameInput = document.createElement('input');
      const nameId = `edit-category-name-${categoryId}`;
      nameLabel.setAttribute('for', nameId);
      nameLabel.textContent = 'Nom de la catégorie *';
      nameInput.id = nameId;
      nameInput.name = 'name';
      nameInput.type = 'text';
      nameInput.required = true;
      nameInput.maxLength = 80;
      nameInput.value = category.name;
      nameRow.append(nameLabel, nameInput);

      const descriptionRow = document.createElement('div');
      descriptionRow.className = 'form-row';
      const descriptionLabel = document.createElement('label');
      const descriptionInput = document.createElement('textarea');
      const descriptionId = `edit-category-description-${categoryId}`;
      descriptionLabel.setAttribute('for', descriptionId);
      descriptionLabel.textContent = 'Description';
      descriptionInput.id = descriptionId;
      descriptionInput.name = 'description';
      descriptionInput.rows = 3;
      descriptionInput.maxLength = 240;
      descriptionInput.value = category.description || '';
      descriptionRow.append(descriptionLabel, descriptionInput);

      const typeRow = document.createElement('div');
      typeRow.className = 'form-row';
      const typeLabel = document.createElement('label');
      const typeSelect = document.createElement('select');
      const typeId = `edit-category-type-${categoryId}`;
      typeLabel.setAttribute('for', typeId);
      typeLabel.textContent = 'Type de valeur *';
      typeSelect.id = typeId;
      typeSelect.name = 'type';
      CATEGORY_TYPE_ORDER.forEach((typeKey) => {
        const option = document.createElement('option');
        option.value = typeKey;
        option.textContent = CATEGORY_TYPE_LABELS[typeKey] || typeKey;
        if (typeKey === (CATEGORY_TYPES.has(category.type) ? category.type : 'text')) {
          option.selected = true;
        }
        typeSelect.appendChild(option);
      });
      typeRow.append(typeLabel, typeSelect);

      const optionsRow = document.createElement('div');
      optionsRow.className = 'form-row';
      const optionsLabel = document.createElement('label');
      const optionsTextarea = document.createElement('textarea');
      const optionsId = `edit-category-options-${categoryId}`;
      optionsLabel.setAttribute('for', optionsId);
      optionsLabel.textContent = 'Liste de valeurs';
      optionsTextarea.id = optionsId;
      optionsTextarea.name = 'options';
      optionsTextarea.rows = 3;
      optionsTextarea.maxLength = 500;
      optionsTextarea.value = Array.isArray(category.options) ? category.options.join('\n') : '';
      const optionsHint = document.createElement('p');
      optionsHint.className = 'form-hint';
      optionsHint.textContent = 'Saisissez une valeur par ligne.';
      optionsRow.append(optionsLabel, optionsTextarea, optionsHint);

      const toggleOptionsRow = () => {
        const shouldShow = typeSelect.value === 'list';
        if (shouldShow) {
          optionsRow.hidden = false;
          optionsRow.removeAttribute('hidden');
        } else {
          optionsRow.hidden = true;
          if (!optionsRow.hasAttribute('hidden')) {
            optionsRow.setAttribute('hidden', '');
          }
          optionsTextarea.setCustomValidity('');
        }
      };

      toggleOptionsRow();
      typeSelect.addEventListener('change', () => {
        toggleOptionsRow();
      });

      const actionsRow = document.createElement('div');
      actionsRow.className = 'category-edit-actions';
      const cancelButton = document.createElement('button');
      cancelButton.type = 'button';
      cancelButton.className = 'category-edit-cancel';
      cancelButton.textContent = 'Annuler';
      const saveButton = document.createElement('button');
      saveButton.type = 'submit';
      saveButton.className = 'category-edit-save';
      saveButton.textContent = 'Enregistrer';
      actionsRow.append(cancelButton, saveButton);

      form.append(nameRow, descriptionRow, typeRow, optionsRow, actionsRow);

      form.addEventListener('submit', (event) => {
        event.preventDefault();
        const name = nameInput.value.trim();
        const description = descriptionInput.value.trim();
        let typeValue = typeSelect.value;
        if (!CATEGORY_TYPES.has(typeValue)) {
          typeValue = 'text';
        }
        let optionValues = [];
        optionsTextarea.setCustomValidity('');
        if (typeValue === 'list') {
          optionValues = parseCategoryOptions(optionsTextarea.value || '');
          if (optionValues.length === 0) {
            optionsTextarea.setCustomValidity('Veuillez renseigner au moins une valeur.');
            optionsTextarea.reportValidity();
            return;
          }
        }

        if (!name) {
          nameInput.focus();
          return;
        }

        category.name = name;
        category.description = description;
        category.type = typeValue;
        category.options = optionValues;
        cleanupContactCategoryValues();
        data.lastUpdated = new Date().toISOString();
        saveDataForUser(currentUser, data);
        renderMetrics();
        renderCategories();
      });

      cancelButton.addEventListener('click', () => {
        renderCategories();
      });

      listItem.appendChild(form);
      nameInput.focus();
      nameInput.setSelectionRange(nameInput.value.length, nameInput.value.length);
    }

    function deleteCategory(categoryId) {
      data.categories = data.categories.filter((item) => item.id !== categoryId);
      normalizeCategoryOrders();
      if (Array.isArray(data.contacts)) {
        data.contacts.forEach((contact) => {
          if (contact && contact.categoryValues && typeof contact.categoryValues === 'object') {
            delete contact.categoryValues[categoryId];
          }
        });
      }
      cleanupContactCategoryValues();
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderMetrics();
      renderCategories();
    }

    function updateCategoryOptionsVisibility() {
      if (!categoryOptionsRow) {
        return;
      }

      const currentType = categoryTypeSelect ? categoryTypeSelect.value : 'text';
      if (currentType === 'list') {
        categoryOptionsRow.hidden = false;
        categoryOptionsRow.removeAttribute('hidden');
        return;
      }

      categoryOptionsRow.hidden = true;
      if (!categoryOptionsRow.hasAttribute('hidden')) {
        categoryOptionsRow.setAttribute('hidden', '');
      }
      if (categoryOptionsInput instanceof HTMLTextAreaElement) {
        categoryOptionsInput.setCustomValidity('');
        categoryOptionsInput.value = '';
      }
    }

    function parseCategoryOptions(rawValue) {
      if (!rawValue) {
        return [];
      }

      const values = rawValue
        .split(/\r?\n|,/)
        .map((value) => value.trim())
        .filter((value) => Boolean(value));

      return Array.from(new Set(values));
    }

    function updateMetricsFromContacts(target = data) {
      if (!target || typeof target !== 'object') {
        return;
      }

      if (!target.metrics || typeof target.metrics !== 'object') {
        target.metrics = { ...defaultData.metrics };
      }

      const contacts = Array.isArray(target.contacts) ? target.contacts : [];
      const categoriesById = buildCategoryMap(target);
      let emailCount = 0;
      let phoneCount = 0;

      contacts.forEach((contact) => {
        const channels = computeContactChannels(contact, categoriesById);
        emailCount += channels.emails.length;
        phoneCount += channels.phones.length;
      });

      target.metrics.peopleCount = contacts.length;
      target.metrics.emailCount = emailCount;
      target.metrics.phoneCount = phoneCount;
    }

    function cleanupContactCategoryValues(target = data) {
      if (!target || typeof target !== 'object') {
        return;
      }

      const categoriesById = new Map(
        Array.isArray(target.categories)
          ? target.categories.map((category) => [category.id, category])
          : [],
      );

      const contacts = Array.isArray(target.contacts) ? target.contacts : [];

      contacts.forEach((contact) => {
        if (!contact || typeof contact !== 'object') {
          return;
        }

        const sourceValues =
          contact.categoryValues && typeof contact.categoryValues === 'object'
            ? contact.categoryValues
            : {};
        const cleanedValues = {};

        Object.entries(sourceValues).forEach(([categoryId, rawValue]) => {
          const category = categoriesById.get(categoryId);
          if (!category) {
            return;
          }

          const baseType = CATEGORY_TYPES.has(category.type) ? category.type : 'text';
          if (rawValue === undefined || rawValue === null) {
            return;
          }

          const valueString = rawValue.toString().trim();
          if (!valueString) {
            return;
          }

          if (baseType === 'number') {
            const parsed = Number(valueString);
            if (!Number.isFinite(parsed)) {
              return;
            }
            cleanedValues[categoryId] = parsed.toString();
            return;
          }

          if (baseType === 'list') {
            const options = Array.isArray(category.options) ? category.options : [];
            if (!options.includes(valueString)) {
              return;
            }
          }

          cleanedValues[categoryId] = valueString;
        });

        contact.categoryValues = cleanedValues;
      });

      updateMetricsFromContacts(target);
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
      if (Array.isArray(data.contacts)) {
        data.contacts.forEach((contact) => {
          if (Array.isArray(contact.keywords)) {
            contact.keywords = contact.keywords.filter((item) => item !== keywordId);
          } else {
            contact.keywords = [];
          }
        });
      }
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderMetrics();
      renderKeywords();
    }

    function upgradeDataStructure(rawData) {
      const base = rawData && typeof rawData === 'object' ? rawData : {};

      if (!base.metrics || typeof base.metrics !== 'object') {
        base.metrics = { ...defaultData.metrics };
      } else {
        base.metrics = { ...defaultData.metrics, ...base.metrics };
      }

      if (!Array.isArray(base.categories)) {
        if (Array.isArray(base.keywords)) {
          base.categories = base.keywords.slice();
          base.keywords = [];
        } else {
          base.categories = [];
        }
      }

      base.categories = base.categories
        .filter((item) => item && typeof item === 'object')
        .map((category) => {
          const normalized = { ...category };
          const typeValue = CATEGORY_TYPES.has(normalized.type) ? normalized.type : 'text';
          normalized.type = typeValue;
          if (typeValue === 'list') {
            const optionValues = Array.isArray(normalized.options)
              ? normalized.options.map((value) => (value != null ? value.toString().trim() : ''))
              : [];
            normalized.options = Array.from(new Set(optionValues.filter((value) => Boolean(value))));
          } else {
            normalized.options = [];
          }
          const rawOrder = normalized.order;
          let parsedOrder = Number.NaN;
          if (typeof rawOrder === 'number') {
            parsedOrder = rawOrder;
          } else if (typeof rawOrder === 'string' && rawOrder.trim() !== '') {
            parsedOrder = Number(rawOrder);
          }
          normalized.order = Number.isFinite(parsedOrder) ? parsedOrder : Number.MAX_SAFE_INTEGER;
          return normalized;
        });

      base.categories = sortCategoriesForDisplay(base.categories);
      base.categories.forEach((category, index) => {
        category.order = index;
      });

      if (!Array.isArray(base.keywords)) {
        base.keywords = [];
      }

      if (!Array.isArray(base.contacts)) {
        base.contacts = [];
      }

      const categoriesById = buildCategoryMap(base);

      base.contacts.forEach((contact) => {
        if (!contact || typeof contact !== 'object') {
          return;
        }

        if (!Array.isArray(contact.keywords)) {
          contact.keywords = [];
        }

        const existingCategoryValues =
          contact.categoryValues && typeof contact.categoryValues === 'object'
            ? contact.categoryValues
            : {};
        const normalizedCategoryValues = {};

        Object.entries(existingCategoryValues).forEach(([categoryId, rawValue]) => {
          if (rawValue === undefined || rawValue === null) {
            return;
          }
          normalizedCategoryValues[categoryId] = rawValue;
        });

        if (Array.isArray(contact.categories)) {
          contact.categories.forEach((categoryId) => {
            if (categoryId && !(categoryId in normalizedCategoryValues)) {
              normalizedCategoryValues[categoryId] = 'Oui';
            }
          });
        }

        contact.categoryValues = normalizedCategoryValues;
        delete contact.categories;

        if (!('mobile' in contact) && typeof contact.phone === 'string') {
          contact.mobile = contact.phone;
        }

        if (!('landline' in contact)) {
          contact.landline = '';
        }

        if (!('phone' in contact) || typeof contact.phone !== 'string') {
          contact.phone = contact.mobile || contact.landline || '';
        }

        if (!contact.usageName && contact.fullName) {
          contact.usageName = contact.fullName;
        }

        if (!contact.firstName && contact.fullName) {
          const segments = contact.fullName.trim().split(/\s+/);
          if (segments.length > 1) {
            contact.firstName = segments.slice(0, segments.length - 1).join(' ');
          }
        }

        if (!contact.fullName) {
          const constructedName = `${contact.firstName || ''} ${contact.usageName || ''}`.trim();
          if (constructedName) {
            contact.fullName = constructedName;
          }
        }

        const derivedName = buildDisplayNameFromCategories(contact.categoryValues, categoriesById);
        const fallbackName = (contact.fullName || `${contact.firstName || ''} ${contact.usageName || ''}`)
          .toString()
          .trim();
        contact.displayName = derivedName || fallbackName || 'Contact sans nom';
      });

      cleanupContactCategoryValues(base);
      updateMetricsFromContacts(base);

      if (!('lastUpdated' in base)) {
        base.lastUpdated = null;
      }

      return base;
    }

    function createEmptyAdvancedFilters() {
      return {
        categories: {},
        keywords: [],
      };
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
          categories: Array.isArray(parsed.categories) ? parsed.categories : undefined,
          keywords: Array.isArray(parsed.keywords) ? parsed.keywords : [],
          contacts: Array.isArray(parsed.contacts) ? parsed.contacts : [],
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
      categories: [],
      keywords: [],
      contacts: [],
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

  function generateId(prefix = 'item') {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) {
      return crypto.randomUUID();
    }
    return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function createContactListItemFallback() {
    const listItem = document.createElement('li');
    listItem.className = 'contact-item';

    const header = document.createElement('div');
    header.className = 'contact-header';
    const nameEl = document.createElement('h3');
    nameEl.className = 'contact-name';
    const createdEl = document.createElement('span');
    createdEl.className = 'contact-created';
    header.append(nameEl, createdEl);

    const details = document.createElement('div');
    details.className = 'contact-details';
    const coordinatesEl = document.createElement('p');
    coordinatesEl.className = 'contact-coordinates contact-coordinates--empty';
    const notesEl = document.createElement('p');
    notesEl.className = 'contact-notes contact-notes--empty';
    details.append(coordinatesEl, notesEl);

    const categoriesContainer = document.createElement('div');
    categoriesContainer.className = 'contact-categories';
    const keywordsContainer = document.createElement('div');
    keywordsContainer.className = 'contact-keywords';
    const actionsContainer = document.createElement('div');
    actionsContainer.className = 'contact-actions';
    const editButton = document.createElement('button');
    editButton.type = 'button';
    editButton.className = 'contact-action-button';
    editButton.dataset.action = 'edit-contact';
    editButton.textContent = 'Modifier';
    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'contact-action-button contact-action-button--danger';
    deleteButton.dataset.action = 'delete-contact';
    deleteButton.textContent = 'Supprimer';
    actionsContainer.append(editButton, deleteButton);

    listItem.append(header, details, categoriesContainer, keywordsContainer, actionsContainer);
    return listItem;
  }

  function createCategoryListItemFallback() {
    const listItem = document.createElement('li');
    listItem.className = 'category-item';

    const dragHandle = document.createElement('button');
    dragHandle.type = 'button';
    dragHandle.className = 'category-drag-handle';
    dragHandle.dataset.dragHandle = 'true';
    dragHandle.setAttribute('aria-label', 'Réorganiser la catégorie');
    dragHandle.setAttribute('title', 'Déplacer la catégorie');
    dragHandle.draggable = true;
    dragHandle.textContent = '⋮⋮';

    const categoryMain = document.createElement('div');
    categoryMain.className = 'category-main';
    const titleEl = document.createElement('h3');
    titleEl.className = 'category-title';
    const descriptionEl = document.createElement('p');
    descriptionEl.className = 'category-description';
    const metaEl = document.createElement('p');
    metaEl.className = 'category-meta';
    categoryMain.append(titleEl, descriptionEl, metaEl);

    const actions = document.createElement('div');
    actions.className = 'category-actions';
    const editButton = document.createElement('button');
    editButton.type = 'button';
    editButton.className = 'category-button';
    editButton.dataset.action = 'edit';
    editButton.textContent = 'Modifier';
    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'category-button category-button--danger';
    deleteButton.dataset.action = 'delete';
    deleteButton.textContent = 'Supprimer';
    actions.append(editButton, deleteButton);

    listItem.append(dragHandle, categoryMain, actions);
    return listItem;
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
