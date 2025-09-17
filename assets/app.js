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
    const metricForms = Array.from(document.querySelectorAll('.metric-form'));
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
    const keywordForm = document.getElementById('keyword-form');
    const keywordList = document.getElementById('keyword-list');
    const keywordEmptyState = document.getElementById('keyword-empty-state');
    const keywordTemplate = document.getElementById('keyword-item-template');
    const contactsCountEl = document.getElementById('contacts-count');
    const contactForm = document.getElementById('contact-form');
    const contactCategoriesContainer = document.getElementById('contact-categories-container');
    const contactCategoriesEmpty = document.getElementById('contact-categories-empty');
    const contactKeywordsContainer = document.getElementById('contact-keywords-container');
    const contactKeywordsEmpty = document.getElementById('contact-keywords-empty');
    const contactList = document.getElementById('contact-list');
    const contactEmptyState = document.getElementById('contact-empty-state');
    const contactSearchInput = document.getElementById('contact-search-input');
    const contactAdvancedSearchForm = document.getElementById('contact-advanced-search');
    const searchCategorySelect = document.getElementById('search-category');
    const searchKeywordsSelect = document.getElementById('search-keywords');
    const contactSearchCountEl = document.getElementById('contact-search-count');
    const contactTemplate = document.getElementById('contact-item-template');

    let data = loadDataForUser(currentUser);
    data = upgradeDataStructure(data);

    let contactSearchTerm = '';
    let advancedFilters = createEmptyAdvancedFilters();

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

    if (categoryForm) {
      categoryForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const formData = new FormData(categoryForm);
        const name = (formData.get('category-name') || '').toString().trim();
        const description = (formData.get('category-description') || '').toString().trim();

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
        });

        data.lastUpdated = new Date().toISOString();
        saveDataForUser(currentUser, data);
        categoryForm.reset();
        const nameInput = categoryForm.querySelector('#category-name');
        if (nameInput instanceof HTMLInputElement) {
          nameInput.focus();
        }
        renderMetrics();
        renderCategories();
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
        const firstName = (formData.get('contact-first-name') || '').toString().trim();
        const usageName = (formData.get('contact-usage-name') || '').toString().trim();
        const birthName = (formData.get('contact-birth-name') || '').toString().trim();
        const gender = (formData.get('contact-gender') || '').toString();
        const ageRange = (formData.get('contact-age-range') || '').toString().trim();
        const email = (formData.get('contact-email') || '').toString().trim();
        const mobile = (formData.get('contact-mobile') || '').toString().trim();
        const landline = (formData.get('contact-landline') || '').toString().trim();
        const street = (formData.get('contact-street') || '').toString().trim();
        const city = (formData.get('contact-city') || '').toString().trim();
        const postalCode = (formData.get('contact-postal-code') || '').toString().trim();
        const country = (formData.get('contact-country') || '').toString().trim();
        const zone = (formData.get('contact-zone') || '').toString().trim();
        const campaignStatus = (formData.get('contact-campaign-status') || '').toString();
        const engagementLevel = (formData.get('contact-engagement-level') || '').toString();
        const archiveTheme = (formData.get('contact-archive-theme') || '').toString();
        const mandate = (formData.get('contact-mandate') || '').toString().trim();
        const profession = (formData.get('contact-profession') || '').toString().trim();
        const lastMembership = (formData.get('contact-last-membership') || '').toString().trim();
        const customField = (formData.get('contact-custom-field') || '').toString().trim();
        const organization = (formData.get('contact-organization') || '').toString().trim();
        const notes = (formData.get('contact-notes') || '').toString().trim();
        const categoryIds = formData
          .getAll('contact-categories')
          .map((value) => value.toString());
        const keywordIds = formData.getAll('contact-keywords').map((value) => value.toString());

        const firstNameInput = contactForm.querySelector('#contact-first-name');
        if (!firstName) {
          if (firstNameInput instanceof HTMLInputElement) {
            firstNameInput.focus();
          }
          return;
        }

        const usageNameInput = contactForm.querySelector('#contact-usage-name');
        if (!usageName) {
          if (usageNameInput instanceof HTMLInputElement) {
            usageNameInput.focus();
          }
          return;
        }

        const emailInput = contactForm.querySelector('#contact-email');
        if (emailInput instanceof HTMLInputElement) {
          emailInput.setCustomValidity('');
          if (email && !isValidEmail(email)) {
            emailInput.setCustomValidity('Veuillez renseigner une adresse e-mail valide.');
            emailInput.reportValidity();
            return;
          }
        }

        const fullName = `${firstName} ${usageName}`.trim();

        data.contacts.push({
          id: generateId('contact'),
          firstName,
          usageName,
          birthName,
          fullName: fullName || usageName || firstName,
          gender,
          ageRange,
          email,
          mobile,
          landline,
          phone: mobile || landline,
          street,
          city,
          postalCode,
          country,
          zone,
          campaignStatus,
          engagementLevel,
          archiveTheme,
          mandate,
          profession,
          lastMembership,
          customField,
          organization,
          notes,
          categories: categoryIds,
          keywords: keywordIds,
          createdAt: new Date().toISOString(),
        });

        data.lastUpdated = new Date().toISOString();
        saveDataForUser(currentUser, data);
        contactForm.reset();
        if (firstNameInput instanceof HTMLInputElement) {
          firstNameInput.focus();
        }
        renderMetrics();
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
        advancedFilters = {
          firstName: (formData.get('search-first-name') || '').toString().trim(),
          usageName: (formData.get('search-usage-name') || '').toString().trim(),
          birthName: (formData.get('search-birth-name') || '').toString().trim(),
          category: (formData.get('search-category') || '').toString(),
          gender: (formData.get('search-gender') || '').toString(),
          age: (formData.get('search-age') || '').toString().trim(),
          email: (formData.get('search-email') || '').toString().trim(),
          mobile: (formData.get('search-mobile') || '').toString().trim(),
          street: (formData.get('search-street') || '').toString().trim(),
          city: (formData.get('search-city') || '').toString().trim(),
          postalCode: (formData.get('search-postal-code') || '').toString().trim(),
          country: (formData.get('search-country') || '').toString().trim(),
          zone: (formData.get('search-zone') || '').toString().trim(),
          campaignStatus: (formData.get('search-campaign-status') || '').toString(),
          engagement: (formData.get('search-engagement') || '').toString(),
          archive: (formData.get('search-archive') || '').toString(),
          mandate: (formData.get('search-mandate') || '').toString().trim(),
          profession: (formData.get('search-profession') || '').toString().trim(),
          landline: (formData.get('search-landline') || '').toString().trim(),
          lastMembership: (formData.get('search-last-membership') || '').toString().trim(),
          custom: (formData.get('search-custom') || '').toString().trim(),
          organization: (formData.get('search-organization') || '').toString().trim(),
          keywords: formData.getAll('search-keywords').map((value) => value.toString()),
        };
        renderContacts();
      });

      contactAdvancedSearchForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          advancedFilters = createEmptyAdvancedFilters();
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
        const totalContacts = Array.isArray(data.contacts) ? data.contacts.length : 0;
        contactsCountEl.textContent = totalContacts.toString();
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

    function renderCategories() {
      if (!categoryList || !categoryEmptyState) {
        renderContactCategoryOptions();
        renderSearchCategoryOptions();
        renderContacts();
        return;
      }

      categoryList.innerHTML = '';

      const categories = Array.isArray(data.categories)
        ? data.categories.slice().sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
        : [];

      if (categories.length === 0) {
        categoryEmptyState.hidden = false;
        renderContactCategoryOptions();
        renderSearchCategoryOptions();
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

        const titleEl = listItem.querySelector('.category-title');
        const descriptionEl = listItem.querySelector('.category-description');
        const editButton = listItem.querySelector('[data-action="edit"]');
        const deleteButton = listItem.querySelector('[data-action="delete"]');

        if (titleEl) {
          titleEl.textContent = category.name;
        }

        if (descriptionEl) {
          descriptionEl.textContent = category.description || 'Aucune description renseignée.';
          descriptionEl.classList.toggle('category-description--empty', !category.description);
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
      renderContactCategoryOptions();
      renderSearchCategoryOptions();
      renderContacts();
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

    function renderContactCategoryOptions() {
      const categories = Array.isArray(data.categories)
        ? data.categories.slice().sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
        : [];

      if (contactCategoriesContainer) {
        contactCategoriesContainer.innerHTML = '';

        if (categories.length === 0) {
          if (contactCategoriesEmpty) {
            contactCategoriesEmpty.hidden = false;
          }
        } else {
          if (contactCategoriesEmpty) {
            contactCategoriesEmpty.hidden = true;
          }

          const fragment = document.createDocumentFragment();

          categories.forEach((category) => {
            const checkboxId = `contact-category-${category.id}`;
            const item = document.createElement('div');
            item.className = 'checkbox-item';

            const input = document.createElement('input');
            input.type = 'checkbox';
            input.id = checkboxId;
            input.name = 'contact-categories';
            input.value = category.id || '';

            const label = document.createElement('label');
            label.setAttribute('for', checkboxId);
            label.textContent = category.name;

            item.append(input, label);
            fragment.appendChild(item);
          });

          contactCategoriesContainer.appendChild(fragment);
        }
      } else if (contactCategoriesEmpty) {
        contactCategoriesEmpty.hidden = categories.length > 0;
      }
    }

    function renderContactKeywordOptions() {
      const keywords = Array.isArray(data.keywords)
        ? data.keywords.slice().sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
        : [];

      if (contactKeywordsContainer) {
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

    function renderSearchCategoryOptions() {
      if (!searchCategorySelect) {
        if (advancedFilters.category) {
          advancedFilters = { ...advancedFilters, category: '' };
        }
        return;
      }

      const categories = Array.isArray(data.categories)
        ? data.categories.slice().sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
        : [];

      const previousSelection = advancedFilters.category || '';
      const allowedValues = new Set(['', ...categories.map((category) => category.id || '')]);

      searchCategorySelect.innerHTML = '';

      const defaultOption = document.createElement('option');
      defaultOption.value = '';
      defaultOption.textContent = 'Toutes les catégories';
      searchCategorySelect.appendChild(defaultOption);

      categories.forEach((category) => {
        const option = document.createElement('option');
        option.value = category.id || '';
        option.textContent = category.name;
        searchCategorySelect.appendChild(option);
      });

      if (allowedValues.has(previousSelection)) {
        searchCategorySelect.value = previousSelection;
      } else {
        searchCategorySelect.value = '';
        if (previousSelection) {
          advancedFilters = { ...advancedFilters, category: '' };
        }
      }
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
      const categoriesById = new Map(
        Array.isArray(data.categories)
          ? data.categories.map((category) => [category.id, category])
          : [],
      );
      const keywordsById = new Map(
        Array.isArray(data.keywords)
          ? data.keywords.map((keyword) => [keyword.id, keyword])
          : [],
      );

      const normalizedTextFilters = {
        firstName: (advancedFilters.firstName || '').toLowerCase(),
        usageName: (advancedFilters.usageName || '').toLowerCase(),
        birthName: (advancedFilters.birthName || '').toLowerCase(),
        age: (advancedFilters.age || '').toLowerCase(),
        email: (advancedFilters.email || '').toLowerCase(),
        mobile: (advancedFilters.mobile || '').toLowerCase(),
        street: (advancedFilters.street || '').toLowerCase(),
        city: (advancedFilters.city || '').toLowerCase(),
        postalCode: (advancedFilters.postalCode || '').toLowerCase(),
        country: (advancedFilters.country || '').toLowerCase(),
        zone: (advancedFilters.zone || '').toLowerCase(),
        mandate: (advancedFilters.mandate || '').toLowerCase(),
        profession: (advancedFilters.profession || '').toLowerCase(),
        landline: (advancedFilters.landline || '').toLowerCase(),
        custom: (advancedFilters.custom || '').toLowerCase(),
        organization: (advancedFilters.organization || '').toLowerCase(),
      };

      const exactFilters = {
        gender: (advancedFilters.gender || '').toString(),
        campaignStatus: (advancedFilters.campaignStatus || '').toString(),
        engagement: (advancedFilters.engagement || '').toString(),
        archive: (advancedFilters.archive || '').toString(),
        lastMembership: (advancedFilters.lastMembership || '').toString(),
        category: (advancedFilters.category || '').toString(),
      };

      const keywordFilters = Array.isArray(advancedFilters.keywords)
        ? advancedFilters.keywords
        : [];

      const hasAdvancedFilters =
        Boolean(
          exactFilters.gender ||
            exactFilters.campaignStatus ||
            exactFilters.engagement ||
            exactFilters.archive ||
            exactFilters.lastMembership ||
            exactFilters.category,
        ) ||
        keywordFilters.length > 0 ||
        Object.values(normalizedTextFilters).some((value) => Boolean(value));

      const filteredContacts = contacts
        .filter((contact) => {
          const categories = Array.isArray(contact.categories) ? contact.categories : [];
          const keywords = Array.isArray(contact.keywords) ? contact.keywords : [];

          if (exactFilters.category && !categories.includes(exactFilters.category)) {
            return false;
          }

          if (keywordFilters.length > 0) {
            const keywordSet = new Set(keywords);
            const matchesKeywords = keywordFilters.every((keywordId) => keywordSet.has(keywordId));
            if (!matchesKeywords) {
              return false;
            }
          }

          const matchesExact = (value, filterValue) => {
            if (!filterValue) {
              return true;
            }
            return (value || '') === filterValue;
          };

          if (!matchesExact((contact.gender || '').toString(), exactFilters.gender)) {
            return false;
          }

          if (
            !matchesExact((contact.campaignStatus || '').toString(), exactFilters.campaignStatus)
          ) {
            return false;
          }

          if (!matchesExact((contact.engagementLevel || '').toString(), exactFilters.engagement)) {
            return false;
          }

          if (!matchesExact((contact.archiveTheme || '').toString(), exactFilters.archive)) {
            return false;
          }

          if (!matchesExact((contact.lastMembership || '').toString(), exactFilters.lastMembership)) {
            return false;
          }

          const matchesText = (value, filterValue) => {
            if (!filterValue) {
              return true;
            }
            return value.toString().toLowerCase().includes(filterValue);
          };

          if (!matchesText(contact.firstName || contact.fullName || '', normalizedTextFilters.firstName)) {
            return false;
          }

          if (!matchesText(contact.usageName || contact.fullName || '', normalizedTextFilters.usageName)) {
            return false;
          }

          if (!matchesText(contact.birthName || '', normalizedTextFilters.birthName)) {
            return false;
          }

          if (!matchesText(contact.ageRange || '', normalizedTextFilters.age)) {
            return false;
          }

          if (!matchesText(contact.email || '', normalizedTextFilters.email)) {
            return false;
          }

          if (!matchesText(contact.mobile || '', normalizedTextFilters.mobile)) {
            return false;
          }

          if (!matchesText(contact.street || '', normalizedTextFilters.street)) {
            return false;
          }

          if (!matchesText(contact.city || '', normalizedTextFilters.city)) {
            return false;
          }

          if (!matchesText(contact.postalCode || '', normalizedTextFilters.postalCode)) {
            return false;
          }

          if (!matchesText(contact.country || '', normalizedTextFilters.country)) {
            return false;
          }

          if (!matchesText(contact.zone || '', normalizedTextFilters.zone)) {
            return false;
          }

          if (!matchesText(contact.mandate || '', normalizedTextFilters.mandate)) {
            return false;
          }

          if (!matchesText(contact.profession || '', normalizedTextFilters.profession)) {
            return false;
          }

          if (!matchesText(contact.landline || '', normalizedTextFilters.landline)) {
            return false;
          }

          if (!matchesText(contact.customField || '', normalizedTextFilters.custom)) {
            return false;
          }

          if (!matchesText(contact.organization || '', normalizedTextFilters.organization)) {
            return false;
          }

          if (!normalizedTerm) {
            return true;
          }

          const categoryNames = categories
            .map((categoryId) => {
              const category = categoriesById.get(categoryId);
              return category ? category.name : '';
            })
            .filter(Boolean);
          const keywordNames = keywords
            .map((keywordId) => {
              const keyword = keywordsById.get(keywordId);
              return keyword ? keyword.name : '';
            })
            .filter(Boolean);

          const haystackParts = [
            contact.fullName || '',
            contact.firstName || '',
            contact.usageName || '',
            contact.birthName || '',
            contact.email || '',
            contact.mobile || '',
            contact.landline || '',
            contact.notes || '',
            contact.organization || '',
            contact.city || '',
            contact.customField || '',
            categoryNames.join(' '),
            keywordNames.join(' '),
          ];

          const haystack = haystackParts.join(' ').toLowerCase();
          return haystack.includes(normalizedTerm);
        })
        .sort((a, b) => {
          const nameA = (a.usageName || a.fullName || '').toString();
          const nameB = (b.usageName || b.fullName || '').toString();
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
          const displayName = (contact.fullName || `${contact.firstName || ''} ${contact.usageName || ''}`).trim();
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
          const coordinates = [];
          const mobileValue = (contact.mobile || '').toString().trim();
          const landlineValue = (contact.landline || '').toString().trim();
          const emailValue = (contact.email || '').toString().trim();
          if (mobileValue) {
            coordinates.push(mobileValue);
          }
          if (landlineValue) {
            coordinates.push(landlineValue);
          }
          if (emailValue) {
            coordinates.push(emailValue);
          }
          if (coordinates.length > 0) {
            coordinatesEl.textContent = coordinates.join(' · ');
            coordinatesEl.classList.remove('contact-coordinates--empty');
          } else {
            coordinatesEl.textContent = 'Aucun moyen de contact renseigné.';
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
          const associatedCategories = Array.isArray(contact.categories)
            ? contact.categories
                .map((categoryId) => categoriesById.get(categoryId))
                .filter((category) => Boolean(category))
            : [];

          if (associatedCategories.length > 0) {
            associatedCategories
              .sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
              .forEach((category) => {
                const chip = document.createElement('span');
                chip.className = 'category-chip';
                chip.textContent = category.name;
                categoriesContainer.appendChild(chip);
              });
          } else {
            const chip = document.createElement('span');
            chip.className = 'category-chip category-chip--empty';
            chip.textContent = 'Sans catégorie';
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

        fragment.appendChild(listItem);
      });

      contactList.appendChild(fragment);
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

      form.append(nameRow, descriptionRow, actionsRow);

      form.addEventListener('submit', (event) => {
        event.preventDefault();
        const name = nameInput.value.trim();
        const description = descriptionInput.value.trim();

        if (!name) {
          nameInput.focus();
          return;
        }

        category.name = name;
        category.description = description;
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
      if (Array.isArray(data.contacts)) {
        data.contacts.forEach((contact) => {
          if (Array.isArray(contact.categories)) {
            contact.categories = contact.categories.filter((item) => item !== categoryId);
          } else {
            contact.categories = [];
          }
        });
      }
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderMetrics();
      renderCategories();
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

      if (!Array.isArray(base.keywords)) {
        base.keywords = [];
      }

      if (!Array.isArray(base.contacts)) {
        base.contacts = [];
      }

      base.contacts.forEach((contact) => {
        if (!Array.isArray(contact.categories)) {
          if (Array.isArray(contact.keywords)) {
            contact.categories = contact.keywords.slice();
          } else {
            contact.categories = [];
          }
        }

        if (!Array.isArray(contact.keywords)) {
          contact.keywords = [];
        }

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
      });

      if (!('lastUpdated' in base)) {
        base.lastUpdated = null;
      }

      return base;
    }

    function createEmptyAdvancedFilters() {
      return {
        firstName: '',
        usageName: '',
        birthName: '',
        category: '',
        gender: '',
        age: '',
        email: '',
        mobile: '',
        street: '',
        city: '',
        postalCode: '',
        country: '',
        zone: '',
        campaignStatus: '',
        engagement: '',
        archive: '',
        mandate: '',
        profession: '',
        landline: '',
        lastMembership: '',
        custom: '',
        organization: '',
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

    listItem.append(header, details, categoriesContainer, keywordsContainer);
    return listItem;
  }

  function createCategoryListItemFallback() {
    const listItem = document.createElement('li');
    listItem.className = 'category-item';

    const categoryMain = document.createElement('div');
    categoryMain.className = 'category-main';
    const titleEl = document.createElement('h3');
    titleEl.className = 'category-title';
    const descriptionEl = document.createElement('p');
    descriptionEl.className = 'category-description';
    categoryMain.append(titleEl, descriptionEl);

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

    listItem.append(categoryMain, actions);
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
