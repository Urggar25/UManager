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
    events: [],
    tasks: [],
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
      phoneOnly: document.querySelector('[data-coverage-count="phone"]'),
      emailOnly: document.querySelector('[data-coverage-count="email"]'),
      none: document.querySelector('[data-coverage-count="none"]'),
    };
    const coveragePercentEls = {
      both: document.querySelector('[data-coverage-percent="both"]'),
      phoneOnly: document.querySelector('[data-coverage-percent="phone"]'),
      emailOnly: document.querySelector('[data-coverage-percent="email"]'),
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
    const contactPagination = document.getElementById('contact-pagination');
    const contactPaginationSummary = document.getElementById('contact-pagination-summary');
    const contactPaginationPageLabel = document.getElementById('contact-pagination-page');
    const contactPaginationPrevButton = document.getElementById('contact-pagination-prev');
    const contactPaginationNextButton = document.getElementById('contact-pagination-next');
    const contactResultsPerPageSelect = document.getElementById('contact-results-per-page');
    const contactTemplate = document.getElementById('contact-item-template');
    const contactSubmitButton = document.getElementById('contact-submit-button');
    const contactCancelEditButton = document.getElementById('contact-cancel-edit');
    const contactBackToSearchButton = document.getElementById('contact-back-to-search');
    const contactsAddTitle = document.getElementById('contacts-add-title');
    const contactsAddSubtitle = document.querySelector('#contacts-add .page-subtitle');
    const contactsAddTitleDefault = contactsAddTitle ? contactsAddTitle.textContent : '';
    const contactsAddSubtitleDefault = contactsAddSubtitle ? contactsAddSubtitle.textContent : '';
    const dashboardCalendarButton = document.getElementById('dashboard-calendar-button');
    const dashboardShortcutButtons = Array.from(
      document.querySelectorAll('[data-shortcut-target]'),
    );
    const calendarOverlay = document.getElementById('calendar-overlay');
    const calendarCloseButton = document.getElementById('calendar-close-button');
    const calendarGridEl = document.getElementById('calendar-grid');
    const calendarCurrentPeriodEl = document.getElementById('calendar-current-period');
    const calendarSelectedDateLabel = document.getElementById('calendar-selected-date-label');
    const calendarSelectedDateSummary = document.getElementById('calendar-selected-date-summary');
    const calendarEventEmptyState = document.getElementById('calendar-event-empty');
    const calendarEventList = document.getElementById('calendar-event-list');
    const calendarEventForm = document.getElementById('calendar-event-form');
    const calendarEventTitleInput = document.getElementById('calendar-event-title');
    const calendarEventDateInput = document.getElementById('calendar-event-date');
    const calendarEventTimeInput = document.getElementById('calendar-event-time');
    const calendarEventNotesInput = document.getElementById('calendar-event-notes');
    const calendarViewButtons = Array.from(document.querySelectorAll('[data-calendar-view]'));
    const calendarNavButtons = Array.from(document.querySelectorAll('[data-calendar-nav]'));
    const taskToggleButton = document.getElementById('dashboard-tasklist-button');
    const taskPanel = document.getElementById('dashboard-task-panel');
    const taskPanelDescription = document.getElementById('task-panel-description');
    const taskForm = document.getElementById('task-form');
    const taskList = document.getElementById('task-list');
    const taskEmptyState = document.getElementById('task-empty-state');
    const taskTitleInput = document.getElementById('task-title');
    const taskDueDateInput = document.getElementById('task-due-date');
    const taskColorInput = document.getElementById('task-color');
    const taskDescriptionInput = document.getElementById('task-description');
    const taskMemberSelect = document.getElementById('task-members');
    const taskCountBadge = document.getElementById('task-count-badge');

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
    const NAME_KEYWORDS = ['nom', 'prenom', 'name', 'usage', 'family'];
    const isoDatePattern = /^\d{4}-\d{2}-\d{2}$/;
    const isoTimePattern = /^\d{2}:\d{2}$/;
    const DEFAULT_TASK_COLOR = '#2563eb';
    const taskPanelDescriptionDefault = taskPanelDescription
      ? taskPanelDescription.textContent || ''
      : '';
    const taskPanelDescriptionDefaultTrimmed = taskPanelDescriptionDefault.trim();
    const taskPanelDescriptionNoMembers = taskPanelDescriptionDefaultTrimmed
      ? `${taskPanelDescriptionDefaultTrimmed} Ajoutez des comptes pour pouvoir attribuer des tâches.`
      : 'Ajoutez des comptes pour pouvoir attribuer des tâches.';
    const calendarMonthFormatter = new Intl.DateTimeFormat('fr-FR', {
      month: 'long',
      year: 'numeric',
    });
    const calendarLongDateFormatter = new Intl.DateTimeFormat('fr-FR', {
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric',
    });
    const calendarWeekRangeStartFormatter = new Intl.DateTimeFormat('fr-FR', {
      day: 'numeric',
      month: 'long',
    });
    const calendarWeekRangeEndFormatter = new Intl.DateTimeFormat('fr-FR', {
      day: 'numeric',
      month: 'long',
      year: 'numeric',
    });
    const calendarShortWeekdayFormatter = new Intl.DateTimeFormat('fr-FR', {
      weekday: 'short',
    });
    const taskDateFormatter = new Intl.DateTimeFormat('fr-FR', {
      dateStyle: 'medium',
    });
    const taskCommentDateFormatter = new Intl.DateTimeFormat('fr-FR', {
      dateStyle: 'medium',
      timeStyle: 'short',
    });

    let data = loadDataForUser(currentUser);
    data = upgradeDataStructure(data);
	
	if (typeof data.panelOwner !== 'string' || !data.panelOwner.trim()) {
	  data.panelOwner = currentUser;
	  saveDataForUser(currentUser, data);
	}

	if (!Array.isArray(data.teamMembers) || data.teamMembers.length === 0) {
	  data.teamMembers = [data.panelOwner];
	  saveDataForUser(currentUser, data);
	} else if (!data.teamMembers.includes(data.panelOwner)) {
	  data.teamMembers.unshift(data.panelOwner);
	  data.teamMembers = Array.from(new Set(data.teamMembers));
	  saveDataForUser(currentUser, data);
	}
	
	const navTeamBtn = document.getElementById('nav-team');
	const isOwner = currentUser === data.panelOwner;

	if (navTeamBtn instanceof HTMLElement) {
	  if (!isOwner) {
		navTeamBtn.remove(); // on supprime l’entrée du menu
	  }
	}

	const originalShowPage = showPage;
	showPage = function(pageId) {
	  if (pageId === 'team' && !isOwner) {
		originalShowPage('dashboard');
		return;
	  }
	  originalShowPage(pageId);
	};


	
	if (!Array.isArray(data.teamMembers) || data.teamMembers.length === 0) {
	  data.teamMembers = [currentUser];
	  saveDataForUser(currentUser, data);
	}

    const numberFormatter = new Intl.NumberFormat('fr-FR');
    const percentFormatter = new Intl.NumberFormat('fr-FR', {
      style: 'percent',
      maximumFractionDigits: 1,
    });

    const CONTACT_RESULTS_PER_PAGE_DEFAULT = 10;

    let contactSearchTerm = '';
    let advancedFilters = createEmptyAdvancedFilters();
    let contactEditReturnPage = 'contacts-search';
    let contactCurrentPage = 1;
    let contactResultsPerPage = CONTACT_RESULTS_PER_PAGE_DEFAULT;
    let categoryDragAndDropInitialized = false;
    let calendarViewMode = 'month';
    let calendarReferenceDate = startOfMonth(new Date());
    let calendarSelectedDate = startOfDay(new Date());
    let calendarLastFocusedElement = null;
    let calendarHasBeenOpened = false;
    let teamMembers = loadTeamMembers();
    let teamMembersById = new Map(teamMembers.map((member) => [member.username, member]));

    normalizeCategoryOrders();
    populateTaskMemberOptions();
    updateTaskPanelDescription();
    resetTaskFormDefaults();
    renderTasks();

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

    if (dashboardCalendarButton) {
      dashboardCalendarButton.addEventListener('click', () => {
        openCalendar();
      });
    }

    if (dashboardShortcutButtons.length > 0) {
      dashboardShortcutButtons.forEach((button) => {
        button.addEventListener('click', () => {
          const target = button.dataset.shortcutTarget;
          if (target) {
            showPage(target);
          }
        });
      });
    }

    if (calendarCloseButton) {
      calendarCloseButton.addEventListener('click', () => {
        closeCalendar();
      });
    }

    if (calendarOverlay) {
      calendarOverlay.addEventListener('click', (event) => {
        if (event.target === calendarOverlay) {
          closeCalendar();
        }
      });
    }

    if (taskToggleButton && taskPanel) {
      taskToggleButton.addEventListener('click', () => {
        if (isTaskPanelOpen()) {
          closeTaskPanel();
          return;
        }
        openTaskPanel();
      });
    }

    if (taskForm) {
	  taskForm.addEventListener('submit', (event) => {
		event.preventDefault();
		if (editingTaskId) {
		  applyTaskEditsFromForm();   // NOUVEAU : en mode édition
		} else {
		  createTaskFromForm();       // Existant : création
		}
	  });

	  taskForm.addEventListener('reset', () => {
		window.requestAnimationFrame(() => {
		  // En reset, on sort du mode édition si on y était
		  editingTaskId = '';
		  setTaskFormMode('create');
		  resetTaskFormDefaults();
		  if (taskTitleInput instanceof HTMLInputElement) taskTitleInput.focus();
		});
	  });
	}


    if (taskList) {
      taskList.addEventListener('click', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) {
          return;
        }
        if (target.dataset.action === 'delete-task') {
          const taskId = target.dataset.taskId || '';
          if (taskId) {
            deleteTask(taskId);
          }
        }
		if (target.dataset.action === 'edit-task') {
		  const taskId = target.dataset.taskId || '';
		  if (taskId) startEditTask(taskId);
		  return;
		}
      });

      taskList.addEventListener('submit', (event) => {
        const form = event.target;
        if (!(form instanceof HTMLFormElement) || !form.classList.contains('task-comment-form')) {
          return;
        }
        event.preventDefault();
        const taskId = form.dataset.taskId || '';
        if (!taskId) {
          return;
        }
        const textarea = form.querySelector('textarea');
        if (!(textarea instanceof HTMLTextAreaElement)) {
          return;
        }
        const content = textarea.value.trim();
        if (!content) {
          textarea.focus();
          if (typeof textarea.reportValidity === 'function') {
            textarea.reportValidity();
          }
          return;
        }
        addCommentToTask(taskId, content);
      });
    }

    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && isTaskPanelOpen()) {
        event.preventDefault();
        closeTaskPanel();
        if (taskToggleButton) {
          taskToggleButton.focus();
        }
      }
    });

    calendarViewButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const view = button.dataset.calendarView;
        if (!view || (view !== 'week' && view !== 'month')) {
          return;
        }
        if (calendarViewMode === view) {
          return;
        }
        calendarViewMode = view;
        if (view === 'month') {
          calendarReferenceDate = startOfMonth(calendarSelectedDate);
        } else {
          calendarReferenceDate = startOfWeek(calendarSelectedDate);
        }
        renderCalendar();
      });
    });

    calendarNavButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const direction = button.dataset.calendarNav;
        if (direction === 'prev') {
          shiftCalendar(-1);
        } else if (direction === 'next') {
          shiftCalendar(1);
        }
      });
    });

    if (calendarEventForm) {
      calendarEventForm.addEventListener('submit', (event) => {
        event.preventDefault();
        createCalendarEventFromForm();
      });

      calendarEventForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          if (calendarEventDateInput) {
            calendarEventDateInput.value = formatDateKey(calendarSelectedDate);
          }
          if (calendarEventTimeInput) {
            calendarEventTimeInput.value = '';
          }
          if (calendarEventNotesInput) {
            calendarEventNotesInput.value = '';
          }
        });
      });
    }

    if (calendarEventList) {
      calendarEventList.addEventListener('click', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) {
          return;
        }
        if (target.dataset.action === 'delete-calendar-event') {
          const eventId = target.dataset.eventId || '';
          if (eventId) {
            deleteCalendarEvent(eventId);
          }
        }
      });
    }

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
        contactCurrentPage = 1;
        renderContacts();
      });
    }

    if (contactResultsPerPageSelect instanceof HTMLSelectElement) {
      contactResultsPerPageSelect.value = contactResultsPerPage.toString();
      contactResultsPerPageSelect.addEventListener('change', () => {
        const parsed = Number.parseInt(contactResultsPerPageSelect.value, 10);
        if (Number.isFinite(parsed) && parsed > 0) {
          contactResultsPerPage = parsed;
        } else {
          contactResultsPerPage = CONTACT_RESULTS_PER_PAGE_DEFAULT;
        }
        contactCurrentPage = 1;
        renderContacts();
      });
    }

    if (contactPaginationPrevButton instanceof HTMLButtonElement) {
      contactPaginationPrevButton.addEventListener('click', () => {
        if (contactCurrentPage > 1) {
          contactCurrentPage -= 1;
          renderContacts();
        }
      });
    }

    if (contactPaginationNextButton instanceof HTMLButtonElement) {
      contactPaginationNextButton.addEventListener('click', () => {
        contactCurrentPage += 1;
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
        contactCurrentPage = 1;
        renderContacts();
      });

      contactAdvancedSearchForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          advancedFilters = createEmptyAdvancedFilters();
          renderSearchCategoryFields();
          contactCurrentPage = 1;
          renderContacts();
        });
      });
    }
	
	if (document.getElementById('team')) {
	  // Rendu initial
	  renderTeamPage();

	  // Submit "Ajouter un membre"
	  const addForm = document.getElementById('team-add-form');
	  if (addForm instanceof HTMLFormElement) {
		addForm.addEventListener('submit', (e) => {
		  e.preventDefault();
		  const formData = new FormData(addForm);
		  const username = (formData.get('username') || '').toString().trim();
		  if (username) {
			addTeamMember(username);
			addForm.reset();
			// recycle le rendu pour que le select se mette à jour
			renderTeamPage();
		  }
		});
	  }

	  // Click "Retirer" (robuste avec .closest)
	  const list = document.getElementById('team-list');
	  if (list) {
	    list.addEventListener('click', (e) => {
	  	  const target = e.target;
	  	  if (!(target instanceof Element)) return;

		  // On remonte jusqu'au bouton qui porte bien data-action="team-remove"
		  const btn = target.closest('button[data-action="team-remove"]');
		  if (!btn) return;

		  const username = btn.getAttribute('data-username') || '';
		  if (username) removeTeamMember(username);
	    });
	  }

	}

    const importApi = {
      getCategories: () =>
        sortCategoriesForDisplay().map((category) => ({
          id: category.id,
          name: category.name,
          type: category.type,
          options: Array.isArray(category.options) ? category.options.slice() : [],
        })),
      importContacts: (payload = {}) => {
        const rows = Array.isArray(payload.rows) ? payload.rows : [];
        const mapping =
          payload && typeof payload.mapping === 'object' ? payload.mapping : {};
        const skipHeader = Boolean(payload && payload.skipHeader);
        const fileName =
          payload && typeof payload.fileName === 'string' ? payload.fileName : '';
        return importContactsFromRows(rows, { mapping, skipHeader, fileName });
      },
    };

    window.UManager = window.UManager || {};
    window.UManager.importApi = importApi;

    showPage('dashboard');
    renderMetrics();
    renderCategories();
    renderKeywords();

    function openCalendar() {
      if (!calendarOverlay) {
        return;
      }

      if (!calendarHasBeenOpened) {
        calendarSelectedDate = startOfDay(new Date());
        calendarReferenceDate =
          calendarViewMode === 'month'
            ? startOfMonth(calendarSelectedDate)
            : startOfWeek(calendarSelectedDate);
        calendarHasBeenOpened = true;
      }

      calendarOverlay.classList.add('calendar-overlay--open');
      calendarOverlay.removeAttribute('hidden');
      calendarOverlay.setAttribute('aria-hidden', 'false');

      calendarLastFocusedElement =
        document.activeElement instanceof HTMLElement ? document.activeElement : null;

      renderCalendar();
      document.addEventListener('keydown', handleCalendarKeydown, true);

      if (calendarCloseButton) {
        calendarCloseButton.focus();
      }
    }

    function closeCalendar() {
      if (!calendarOverlay || !calendarOverlay.classList.contains('calendar-overlay--open')) {
        return;
      }
      calendarOverlay.classList.remove('calendar-overlay--open');
      calendarOverlay.setAttribute('aria-hidden', 'true');
      if (!calendarOverlay.hasAttribute('hidden')) {
        calendarOverlay.setAttribute('hidden', '');
      }
      document.removeEventListener('keydown', handleCalendarKeydown, true);
      if (calendarLastFocusedElement && typeof calendarLastFocusedElement.focus === 'function') {
        calendarLastFocusedElement.focus();
      }
    }

    function handleCalendarKeydown(event) {
      if (event.key === 'Escape') {
        event.preventDefault();
        closeCalendar();
      }
    }

    function renderCalendar() {
      if (!calendarGridEl) {
        return;
      }

      updateCalendarViewButtons();

      if (calendarCurrentPeriodEl) {
        if (calendarViewMode === 'month') {
          calendarCurrentPeriodEl.textContent = capitalizeLabel(
            calendarMonthFormatter.format(calendarReferenceDate),
          );
        } else {
          const start = startOfWeek(calendarReferenceDate);
          const end = addDays(start, 6);
          calendarCurrentPeriodEl.textContent = formatWeekRangeLabel(start, end);
        }
      }

      calendarGridEl.classList.toggle('calendar-grid--week', calendarViewMode === 'week');
      calendarGridEl.innerHTML = '';

      if (calendarViewMode === 'month') {
        renderMonthView();
      } else {
        renderWeekView();
      }

      updateSelectedDatePanel();
    }

    function updateCalendarViewButtons() {
      calendarViewButtons.forEach((button) => {
        const view = button.dataset.calendarView;
        const isActive = view === calendarViewMode;
        button.classList.toggle('active', isActive);
        button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
      });
    }

    function renderMonthView() {
      if (!calendarGridEl) {
        return;
      }
      const firstOfMonth = startOfMonth(calendarReferenceDate);
      const gridStart = startOfWeek(firstOfMonth);
      const fragment = document.createDocumentFragment();
      for (let index = 0; index < 42; index += 1) {
        const dayDate = addDays(gridStart, index);
        const isInCurrentPeriod =
          dayDate.getMonth() === calendarReferenceDate.getMonth() &&
          dayDate.getFullYear() === calendarReferenceDate.getFullYear();
        fragment.appendChild(
          createCalendarDayElement(dayDate, {
            isInCurrentPeriod,
            includeWeekdayLabel: index < 7,
          }),
        );
      }
      calendarGridEl.appendChild(fragment);
    }

    function renderWeekView() {
      if (!calendarGridEl) {
        return;
      }
      const weekStart = startOfWeek(calendarReferenceDate);
      const fragment = document.createDocumentFragment();
      for (let index = 0; index < 7; index += 1) {
        const dayDate = addDays(weekStart, index);
        fragment.appendChild(
          createCalendarDayElement(dayDate, {
            isInCurrentPeriod: true,
            includeWeekdayLabel: true,
          }),
        );
      }
      calendarGridEl.appendChild(fragment);
    }

    function createCalendarDayElement(date, options = {}) {
      const dayDate = startOfDay(date instanceof Date ? date : new Date(date));
      const dateKey = formatDateKey(dayDate);
      const isSelected = formatDateKey(calendarSelectedDate) === dateKey;
      const eventsForDay = getEventsForDate(dayDate);

      const element = document.createElement('div');
      element.className = 'calendar-day';
      if (!options.isInCurrentPeriod) {
        element.classList.add('calendar-day--outside');
      }
      if (isSelected) {
        element.classList.add('calendar-day--selected');
      }
      element.dataset.date = dateKey;
      element.setAttribute('role', 'button');
      element.tabIndex = 0;

      const header = document.createElement('div');
      header.className = 'calendar-day-header';

      const weekdayLabel = document.createElement('span');
      weekdayLabel.className = 'calendar-day-weekday';
      weekdayLabel.setAttribute('aria-hidden', 'true');
      if (options.includeWeekdayLabel) {
        weekdayLabel.textContent = capitalizeLabel(
          calendarShortWeekdayFormatter.format(dayDate),
        );
      } else {
        weekdayLabel.classList.add('calendar-day-weekday--empty');
        weekdayLabel.textContent = '\u00A0';
      }

      const dayNumber = document.createElement('span');
      dayNumber.className = 'calendar-day-number';
      dayNumber.textContent = dayDate.getDate().toString();
      dayNumber.setAttribute('aria-hidden', 'true');

      header.append(weekdayLabel, dayNumber);
      element.appendChild(header);

      if (eventsForDay.length > 0) {
        const list = document.createElement('ul');
        list.className = 'calendar-event-chips';
        eventsForDay.slice(0, 3).forEach((eventItem) => {
          const chip = document.createElement('li');
          chip.className = 'calendar-event-chip';
          chip.textContent = eventItem.time
            ? `${eventItem.time} · ${eventItem.title}`
            : eventItem.title;
          list.appendChild(chip);
        });
        if (eventsForDay.length > 3) {
          const moreChip = document.createElement('li');
          moreChip.className = 'calendar-event-chip calendar-event-chip--more';
          moreChip.textContent = `+${eventsForDay.length - 3} évènement(s)`;
          list.appendChild(moreChip);
        }
        element.appendChild(list);
      }

      const ariaLabelParts = [capitalizeLabel(calendarLongDateFormatter.format(dayDate))];
      if (eventsForDay.length === 0) {
        ariaLabelParts.push('Aucun évènement.');
      } else if (eventsForDay.length === 1) {
        ariaLabelParts.push('1 évènement.');
      } else {
        ariaLabelParts.push(`${eventsForDay.length} évènements.`);
      }
      element.setAttribute('aria-label', ariaLabelParts.join(' '));

      element.addEventListener('click', () => {
        selectCalendarDate(dayDate);
      });

      element.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          selectCalendarDate(dayDate);
        }
      });

      return element;
    }

    function selectCalendarDate(date) {
      calendarSelectedDate = startOfDay(date instanceof Date ? date : new Date(date));
      if (calendarViewMode === 'month') {
        if (
          calendarSelectedDate.getMonth() !== calendarReferenceDate.getMonth() ||
          calendarSelectedDate.getFullYear() !== calendarReferenceDate.getFullYear()
        ) {
          calendarReferenceDate = startOfMonth(calendarSelectedDate);
        }
      } else {
        calendarReferenceDate = startOfWeek(calendarSelectedDate);
      }
      renderCalendar();
    }

    function updateSelectedDatePanel() {
      const eventsForDay = getEventsForDate(calendarSelectedDate);

      if (calendarSelectedDateLabel) {
        calendarSelectedDateLabel.textContent = capitalizeLabel(
          calendarLongDateFormatter.format(calendarSelectedDate),
        );
      }

      if (calendarSelectedDateSummary) {
        if (eventsForDay.length === 0) {
          calendarSelectedDateSummary.textContent = 'Aucun évènement enregistré pour cette date.';
        } else if (eventsForDay.length === 1) {
          calendarSelectedDateSummary.textContent = '1 évènement planifié.';
        } else {
          calendarSelectedDateSummary.textContent = `${eventsForDay.length} évènements planifiés.`;
        }
      }

      if (calendarEventDateInput) {
        calendarEventDateInput.value = formatDateKey(calendarSelectedDate);
      }

      renderSelectedDateEvents(eventsForDay);
    }

    function renderSelectedDateEvents(eventsForDay) {
      if (!calendarEventList) {
        return;
      }

      calendarEventList.innerHTML = '';

      if (!Array.isArray(eventsForDay) || eventsForDay.length === 0) {
        if (calendarEventEmptyState) {
          calendarEventEmptyState.hidden = false;
        }
        return;
      }

      if (calendarEventEmptyState) {
        calendarEventEmptyState.hidden = true;
      }

      eventsForDay
        .slice()
        .sort(compareCalendarEvents)
        .forEach((eventItem) => {
          const listItem = document.createElement('li');
          listItem.className = 'calendar-event-item';
          listItem.dataset.eventId = eventItem.id;

          const header = document.createElement('div');
          header.className = 'calendar-event-item-header';

          const title = document.createElement('h4');
          title.className = 'calendar-event-title';
          title.textContent = eventItem.title;

          const time = document.createElement('span');
          time.className = 'calendar-event-time';
          time.textContent = eventItem.time ? eventItem.time : 'Toute la journée';

          header.append(title, time);
          listItem.appendChild(header);

          if (eventItem.notes) {
            const notes = document.createElement('p');
            notes.className = 'calendar-event-notes';
            notes.textContent = eventItem.notes;
            listItem.appendChild(notes);
          }

          const actions = document.createElement('div');
          actions.className = 'calendar-event-actions';
          const deleteButton = document.createElement('button');
          deleteButton.type = 'button';
          deleteButton.className = 'calendar-event-delete';
          deleteButton.dataset.action = 'delete-calendar-event';
          deleteButton.dataset.eventId = eventItem.id;
          deleteButton.textContent = 'Supprimer';
          actions.appendChild(deleteButton);
          listItem.appendChild(actions);

          calendarEventList.appendChild(listItem);
        });
    }

    function shiftCalendar(step) {
      if (!Number.isFinite(step) || step === 0) {
        return;
      }

      if (calendarViewMode === 'month') {
        const newReference = new Date(calendarReferenceDate);
        newReference.setMonth(newReference.getMonth() + step);
        calendarReferenceDate = startOfMonth(newReference);

        const newSelected = new Date(calendarSelectedDate);
        newSelected.setMonth(newSelected.getMonth() + step);
        calendarSelectedDate = startOfDay(newSelected);
      } else {
        calendarReferenceDate = startOfWeek(addDays(calendarReferenceDate, step * 7));
        calendarSelectedDate = startOfDay(addDays(calendarSelectedDate, step * 7));
      }

      renderCalendar();
    }

    function createCalendarEventFromForm() {
      if (!calendarEventForm || !calendarEventTitleInput || !calendarEventDateInput) {
        return;
      }

      const title = calendarEventTitleInput.value.trim();
      const dateValue = calendarEventDateInput.value;
      const timeValue = calendarEventTimeInput ? calendarEventTimeInput.value : '';
      const notesValue = calendarEventNotesInput ? calendarEventNotesInput.value.trim() : '';

      if (!title) {
        calendarEventTitleInput.focus();
        return;
      }

      if (!isValidDateKey(dateValue)) {
        calendarEventDateInput.focus();
        calendarEventDateInput.reportValidity();
        return;
      }

      const normalizedTime = timeValue && isoTimePattern.test(timeValue) ? timeValue : '';

      data.events.push({
        id: generateId('event'),
        title,
        date: dateValue,
        time: normalizedTime,
        notes: notesValue,
      });

      data.events.sort(compareCalendarEvents);
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);

      const parsedDate = parseDateInput(dateValue);
      if (parsedDate) {
        calendarSelectedDate = parsedDate;
        calendarReferenceDate =
          calendarViewMode === 'month' ? startOfMonth(parsedDate) : startOfWeek(parsedDate);
      }

      calendarEventForm.reset();
      calendarEventDateInput.value = dateValue;
      if (calendarEventTimeInput) {
        calendarEventTimeInput.value = normalizedTime;
      }
      if (calendarEventTitleInput) {
        calendarEventTitleInput.focus();
      }

      renderCalendar();
    }

    function deleteCalendarEvent(eventId) {
      const initialLength = data.events.length;
      data.events = data.events.filter((eventItem) => eventItem && eventItem.id !== eventId);
      if (data.events.length === initialLength) {
        return;
      }
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderCalendar();
    }
	
	function upsertCalendarEventForTask(task) {
	  const EVENT_TIME  = '20:00';
	  const EVENT_TITLE = `📝 Tâche · ${task.title}`;
	  if (!Array.isArray(data.events)) data.events = [];

	  // Pas de deadline → supprimer l’éventuel évènement lié
	  if (!task.dueDate) {
		if (task.calendarEventId) {
		  deleteCalendarEvent(task.calendarEventId);
		  task.calendarEventId = '';
		}
		return;
	  }

	  const existingId = (typeof task.calendarEventId === 'string' && task.calendarEventId) ? task.calendarEventId : '';
	  const existing   = existingId ? data.events.find(e => e && e.id === existingId) : null;

	  if (existing) {
		const updated = {
		  ...existing,
		  title: EVENT_TITLE,
		  date: task.dueDate,        // <-- nouvelle date
		  time: EVENT_TIME,
		  notes: task.description || '',
		};

		// Si la date a changé, on supprime puis on ré-ajoute (évite les caches internes)
		if (existing.date !== updated.date) {
		  data.events = data.events.filter(e => e && e.id !== existingId);
		  data.events.push(updated);
		} else {
		  // Sinon, remplacement immuable dans le tableau
		  data.events = data.events.map(e => (e && e.id === existingId) ? updated : e);
		}
	  } else {
		// Pas encore lié → on crée l’évènement
		const newId = generateId('event');
		data.events = [...data.events, {
		  id: newId,
		  title: EVENT_TITLE,
		  date: task.dueDate,
		  time: EVENT_TIME,
		  notes: task.description || '',
		}];
		task.calendarEventId = newId;
	  }

	  if (Array.isArray(data.events)) data.events.sort(compareCalendarEvents);
	}


    function getEventsForDate(date) {
      const dateKey = formatDateKey(date instanceof Date ? date : new Date(date));
      return data.events
        .filter((eventItem) => eventItem && eventItem.date === dateKey)
        .slice()
        .sort(compareCalendarEvents);
    }

    function formatWeekRangeLabel(start, end) {
      return `Semaine du ${capitalizeLabel(
        calendarWeekRangeStartFormatter.format(start),
      )} au ${capitalizeLabel(calendarWeekRangeEndFormatter.format(end))}`;
    }

    function startOfDay(date) {
      const reference = date instanceof Date ? new Date(date) : new Date();
      reference.setHours(0, 0, 0, 0);
      return reference;
    }

    function startOfMonth(date) {
      const reference = date instanceof Date ? date : new Date();
      return startOfDay(new Date(reference.getFullYear(), reference.getMonth(), 1));
    }

    function startOfWeek(date) {
      const reference = startOfDay(date instanceof Date ? date : new Date());
      const day = reference.getDay();
      const offset = (day + 6) % 7;
      reference.setDate(reference.getDate() - offset);
      return reference;
    }

    function addDays(date, amount) {
      const reference = date instanceof Date ? new Date(date) : new Date();
      reference.setDate(reference.getDate() + amount);
      return reference;
    }

    function formatDateKey(date) {
      if (!(date instanceof Date)) {
        return '';
      }
      const year = date.getFullYear();
      const month = `${date.getMonth() + 1}`.padStart(2, '0');
      const day = `${date.getDate()}`.padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    function isValidDateKey(value) {
      return typeof value === 'string' && isoDatePattern.test(value);
    }

    function parseDateInput(value) {
      if (!isValidDateKey(value)) {
        return null;
      }
      const [yearPart, monthPart, dayPart] = value.split('-');
      const year = Number.parseInt(yearPart, 10);
      const month = Number.parseInt(monthPart, 10);
      const day = Number.parseInt(dayPart, 10);
      if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
        return null;
      }
      return startOfDay(new Date(year, month - 1, day));
    }

    function parseDateTimeToParts(value) {
      if (typeof value !== 'string' || value.trim() === '') {
        return null;
      }
      const parsed = new Date(value);
      if (Number.isNaN(parsed.getTime())) {
        return null;
      }
      return {
        date: formatDateKey(parsed),
        time: `${String(parsed.getHours()).padStart(2, '0')}:${String(parsed.getMinutes()).padStart(
          2,
          '0',
        )}`,
      };
    }

    function capitalizeLabel(value) {
      if (typeof value !== 'string' || value.length === 0) {
        return '';
      }
      return value.charAt(0).toUpperCase() + value.slice(1);
    }

    function normalizeCalendarEvent(rawEvent) {
      const baseEvent = rawEvent && typeof rawEvent === 'object' ? rawEvent : {};

      let title = '';
      if (typeof baseEvent.title === 'string') {
        title = baseEvent.title.trim();
      } else if (typeof baseEvent.name === 'string') {
        title = baseEvent.name.trim();
      }
      if (!title) {
        title = 'Évènement';
      }

      let dateValue = '';
      if (typeof baseEvent.date === 'string' && isValidDateKey(baseEvent.date.trim())) {
        dateValue = baseEvent.date.trim();
      } else if (typeof baseEvent.day === 'string' && isValidDateKey(baseEvent.day.trim())) {
        dateValue = baseEvent.day.trim();
      }

      let timeValue = '';
      if (typeof baseEvent.time === 'string' && isoTimePattern.test(baseEvent.time.trim())) {
        timeValue = baseEvent.time.trim();
      } else if (typeof baseEvent.hour === 'string' && isoTimePattern.test(baseEvent.hour.trim())) {
        timeValue = baseEvent.hour.trim();
      }

      if ((!dateValue || !timeValue) && typeof baseEvent.datetime === 'string') {
        const parts = parseDateTimeToParts(baseEvent.datetime);
        if (parts) {
          if (!dateValue) {
            dateValue = parts.date;
          }
          if (!timeValue) {
            timeValue = parts.time;
          }
        }
      }

      if (!dateValue) {
        dateValue = formatDateKey(startOfDay(new Date()));
      }

      if (timeValue && !isoTimePattern.test(timeValue)) {
        timeValue = '';
      }

      let notes = '';
      if (typeof baseEvent.notes === 'string') {
        notes = baseEvent.notes.trim();
      } else if (typeof baseEvent.description === 'string') {
        notes = baseEvent.description.trim();
      }

      let id = '';
      if (typeof baseEvent.id === 'string' && baseEvent.id.trim()) {
        id = baseEvent.id.trim();
      }
      if (!id) {
        id = generateId('event');
      }

      return {
        id,
        title,
        date: dateValue,
        time: timeValue,
        notes,
      };
    }

    function compareCalendarEvents(a, b) {
      if (!a || !b) {
        return 0;
      }
      const dateComparison = a.date.localeCompare(b.date);
      if (dateComparison !== 0) {
        return dateComparison;
      }
      const timeA = a.time && isoTimePattern.test(a.time) ? a.time : '00:00';
      const timeB = b.time && isoTimePattern.test(b.time) ? b.time : '00:00';
      const timeComparison = timeA.localeCompare(timeB);
      if (timeComparison !== 0) {
        return timeComparison;
      }
      return a.title.localeCompare(b.title);
    }

    function showPage(target) {
      if (target !== 'dashboard') {
        closeCalendar();
      }
      navButtons.forEach((button) => {
        button.classList.toggle('active', button.dataset.target === target);
      });
      pages.forEach((page) => {
        page.classList.toggle('active', page.id === target);
      });
      document.dispatchEvent(
        new CustomEvent('umanager:page-changed', {
          detail: { pageId: target },
        }),
      );
    }

    function loadTeamMembers() {
	  const allowedUsernames = Array.isArray(data.teamMembers) && data.teamMembers.length > 0
		? data.teamMembers.map((u) => (typeof u === 'string' ? u.trim() : '')).filter(Boolean)
		: [currentUser];

	  const store = loadUserStore();
	  const usersObj = store && store.users && typeof store.users === 'object' ? store.users : {};

	  const members = allowedUsernames.map((username) => {
		const details = usersObj[username] && typeof usersObj[username] === 'object' ? usersObj[username] : {};
		const email = typeof details.email === 'string' ? details.email : '';
		return { username, email };
	  });

	  return members.sort((a, b) => a.username.localeCompare(b.username, 'fr', { sensitivity: 'base' }));
	}


    function populateTaskMemberOptions() {
      if (!(taskMemberSelect instanceof HTMLSelectElement)) {
        return;
      }

      taskMemberSelect.innerHTML = '';

      if (!Array.isArray(teamMembers) || teamMembers.length === 0) {
        taskMemberSelect.disabled = true;
        const option = document.createElement('option');
        option.value = '';
        option.textContent = 'Aucun membre disponible';
        option.disabled = true;
        taskMemberSelect.appendChild(option);
        return;
      }

      taskMemberSelect.disabled = false;
      const fragment = document.createDocumentFragment();

      teamMembers.forEach((member) => {
        const option = document.createElement('option');
        option.value = member.username;
        option.textContent = formatTeamMemberLabel(member);
        fragment.appendChild(option);
      });

      taskMemberSelect.appendChild(fragment);
    }
	
	function loadAllUsersFromStore() {
	  const store = loadUserStore();
	  const usersObj = store && store.users && typeof store.users === 'object' ? store.users : {};
	  // Retourne [{ username, email? }, ...]
	  return Object.keys(usersObj).map((u) => {
		const v = usersObj[u] || {};
		return { username: u, email: typeof v.email === 'string' ? v.email : '' };
	  }).sort((a, b) => a.username.localeCompare(b.username, 'fr', { sensitivity: 'base' }));
	}

	function getAvailableUsersForTeam() {
	  const all = loadAllUsersFromStore();
	  const team = new Set(Array.isArray(data.teamMembers) ? data.teamMembers : []);
	  return all.filter(({ username }) => !team.has(username));
	}

	function rebuildTeamCaches() {
	  // Recalcule la liste et la map depuis la source de vérité (data.teamMembers + store)
	  teamMembers = loadTeamMembers();
	  teamMembersById = new Map(teamMembers.map((m) => [m.username, m]));

	  // Répercuter partout
	  populateTaskMemberOptions(); // met à jour le <select> des tâches
	  sanitizeTasksAgainstTeam();  // purge les assignations hors équipe si besoin
	  updateTaskPanelDescription();
	}


	function sanitizeTasksAgainstTeam() {
	  if (!Array.isArray(data.tasks)) return;
	  let changed = false;
	  data.tasks.forEach((t) => {
		if (!t || !Array.isArray(t.assignedMembers)) return;
		const filtered = t.assignedMembers.filter((u) => teamMembersById.has(u));
		if (filtered.length !== t.assignedMembers.length) {
		  t.assignedMembers = filtered;
		  changed = true;
		}
	  });
	  if (changed) {
		data.tasks = data.tasks.map((it) => normalizeTask(it)).sort(compareTasks);
		data.lastUpdated = new Date().toISOString();
		saveDataForUser(currentUser, data);
		renderTasks();
	  }
	}


    function updateTaskPanelDescription() {
      if (!taskPanelDescription) {
        return;
      }

      if (!Array.isArray(teamMembers) || teamMembers.length === 0) {
        taskPanelDescription.textContent = taskPanelDescriptionNoMembers;
        return;
      }

      taskPanelDescription.textContent =
        taskPanelDescriptionDefaultTrimmed || taskPanelDescriptionDefault || '';
    }

    function resetTaskFormDefaults() {
      if (taskColorInput instanceof HTMLInputElement) {
        taskColorInput.value = DEFAULT_TASK_COLOR;
      }

      if (taskDueDateInput instanceof HTMLInputElement) {
        taskDueDateInput.setCustomValidity('');
      }

      if (taskMemberSelect instanceof HTMLSelectElement && !taskMemberSelect.disabled) {
        Array.from(taskMemberSelect.options).forEach((option) => {
          option.selected = false;
        });
      }

      if (taskDescriptionInput instanceof HTMLTextAreaElement) {
        taskDescriptionInput.value = '';
      }
    }

    function isTaskPanelOpen() {
      return Boolean(taskPanel && !taskPanel.hidden);
    }

    function openTaskPanel() {
      if (!taskPanel) {
        return;
      }

      taskPanel.hidden = false;
      taskPanel.removeAttribute('hidden');
      if (taskToggleButton) {
        taskToggleButton.setAttribute('aria-expanded', 'true');
      }

      window.requestAnimationFrame(() => {
        if (taskTitleInput instanceof HTMLInputElement) {
          taskTitleInput.focus();
        }
      });
    }

    function closeTaskPanel() {
      if (!taskPanel) {
        return;
      }

      taskPanel.hidden = true;
      if (!taskPanel.hasAttribute('hidden')) {
        taskPanel.setAttribute('hidden', '');
      }
      if (taskToggleButton) {
        taskToggleButton.setAttribute('aria-expanded', 'false');
      }

      if (taskForm) {
        taskForm.reset();
      }
      resetTaskFormDefaults();
    }

    function createTaskFromForm() {
      if (!taskForm || !(taskTitleInput instanceof HTMLInputElement)) {
        return;
      }

      const formData = new FormData(taskForm);
      const title = (formData.get('task-title') || '').toString().trim();
      if (!title) {
        taskTitleInput.focus();
        return;
      }

      let dueDate = '';
      const dueDateRaw = (formData.get('task-due-date') || '').toString().trim();
      if (taskDueDateInput instanceof HTMLInputElement) {
        taskDueDateInput.setCustomValidity('');
      }
      if (dueDateRaw) {
        if (isValidDateKey(dueDateRaw)) {
          dueDate = dueDateRaw;
        } else if (taskDueDateInput instanceof HTMLInputElement) {
          taskDueDateInput.setCustomValidity('Veuillez sélectionner une date valide.');
          taskDueDateInput.reportValidity();
          taskDueDateInput.focus();
          return;
        }
      }

      const colorValue = normalizeTaskColor(
        (formData.get('task-color') || DEFAULT_TASK_COLOR).toString(),
      );
      const descriptionValue = (formData.get('task-description') || '').toString().trim();

      const members = [];
      if (taskMemberSelect instanceof HTMLSelectElement && !taskMemberSelect.disabled) {
        Array.from(taskMemberSelect.selectedOptions).forEach((option) => {
          if (option && option.value) {
            members.push(option.value);
          }
        });
      }
	  
	  const allowedMembers = members.filter((u) => teamMembersById.has(u));

      const newTask = normalizeTask({
        id: generateId('task'),
        title,
        dueDate,
        color: colorValue,
        description: descriptionValue,
        assignedMembers: allowedMembers,
        createdAt: new Date().toISOString(),
        createdBy: currentUser,
        comments: [],
      });
	  
	  if (newTask.dueDate) {
	    const eventId = generateId('event');
	    // Titre d’évènement explicite ; ajuste si tu veux quelque chose de plus court
	    const eventTitle = `📝 Tâche · ${newTask.title}`;

	    data.events.push({
		  id: eventId,
		  title: eventTitle,
		  date: newTask.dueDate,
		  time: '20:00',           // <- à 20h
		  notes: newTask.description || '',
	    });

	    // Lier l’évènement à la tâche
	    newTask.calendarEventId = eventId;
  
	    // Garder le tri et l’état cohérents dans le calendrier
	    data.events.sort(compareCalendarEvents);
	  }

      data.tasks.push(newTask);
      data.tasks = data.tasks.map((item) => normalizeTask(item)).sort(compareTasks);
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      taskForm.reset();
      resetTaskFormDefaults();
      renderTasks();

      window.requestAnimationFrame(() => {
        taskTitleInput.focus();
      });
    }

    function renderTasks() {
      if (!taskList) {
        return;
      }

      const tasks = Array.isArray(data.tasks) ? data.tasks.slice() : [];
      tasks.sort(compareTasks);
      if (Array.isArray(data.tasks)) {
        data.tasks = tasks.slice();
      }

      taskList.innerHTML = '';

      if (taskEmptyState) {
        if (tasks.length === 0) {
          taskEmptyState.hidden = false;
          taskEmptyState.removeAttribute('hidden');
        } else {
          taskEmptyState.hidden = true;
          if (!taskEmptyState.hasAttribute('hidden')) {
            taskEmptyState.setAttribute('hidden', '');
          }
        }
      }

      updateTaskCountDisplay(tasks.length);

      if (tasks.length === 0) {
        return;
      }

      const fragment = document.createDocumentFragment();

      tasks.forEach((task) => {
        if (!task || typeof task !== 'object') {
          return;
        }

        const listItem = document.createElement('li');
        listItem.className = 'task-item';
        listItem.dataset.id = task.id || '';

        const taskColor = normalizeTaskColor(task.color);
        listItem.style.setProperty('--task-color', taskColor);
        listItem.style.setProperty('--task-color-soft', hexToRgba(taskColor, 0.12));

        const header = document.createElement('div');
        header.className = 'task-item-header';

        const titleEl = document.createElement('h3');
        titleEl.className = 'task-title';
        titleEl.textContent = task.title || 'Nouvelle tâche';

        const metaContainer = document.createElement('div');
        metaContainer.className = 'task-meta';

        if (task.dueDate) {
          const dueDate = parseDateInput(task.dueDate);
          const dueDateEl = document.createElement('span');
          dueDateEl.className = 'task-due-date';
          const formattedDueDate =
            dueDate instanceof Date && !Number.isNaN(dueDate.getTime())
              ? capitalizeLabel(taskDateFormatter.format(dueDate))
              : task.dueDate;
          dueDateEl.textContent = `Échéance : ${formattedDueDate}`;

          if (dueDate && dueDate.getTime() < startOfDay(new Date()).getTime()) {
            dueDateEl.classList.add('task-due-date--overdue');
          }

          metaContainer.appendChild(dueDateEl);
        }

        if (Array.isArray(task.assignedMembers) && task.assignedMembers.length > 0) {
          const membersContainer = document.createElement('div');
          membersContainer.className = 'task-members';

          task.assignedMembers.forEach((memberId) => {
            if (!memberId) {
              return;
            }
            const chip = document.createElement('span');
            chip.className = 'task-member-chip';
            const member = teamMembersById.get(memberId);
            chip.textContent = member ? formatTeamMemberLabel(member) : memberId;
            membersContainer.appendChild(chip);
          });

          metaContainer.appendChild(membersContainer);
        }

        const actions = document.createElement('div');
        actions.className = 'task-actions';
		
		const editButton = document.createElement('button');
		editButton.type = 'button';
		editButton.className = 'task-action-button';
		editButton.dataset.action = 'edit-task';
		editButton.dataset.taskId = task.id || '';
		editButton.textContent = 'Modifier';
		actions.appendChild(editButton);

        const deleteButton = document.createElement('button');
        deleteButton.type = 'button';
        deleteButton.className = 'task-action-button task-action-button--danger';
        deleteButton.dataset.action = 'delete-task';
        deleteButton.dataset.taskId = task.id || '';
        deleteButton.textContent = 'Supprimer';

        actions.appendChild(deleteButton);

        header.appendChild(titleEl);
        if (metaContainer.childElementCount > 0) {
          header.appendChild(metaContainer);
        }
        header.appendChild(actions);

        const descriptionEl = document.createElement('p');
        descriptionEl.className = 'task-description';
        if (task.description) {
          descriptionEl.textContent = task.description;
        } else {
          descriptionEl.textContent = 'Aucune description pour cette tâche.';
          descriptionEl.classList.add('task-description--empty');
        }

        const commentsSection = document.createElement('div');
        commentsSection.className = 'task-comments';

        const commentsTitle = document.createElement('h4');
        commentsTitle.className = 'task-comments-title';
        commentsTitle.textContent = 'Commentaires';
        commentsSection.appendChild(commentsTitle);

        const commentsList = document.createElement('ul');
        commentsList.className = 'task-comment-list';

        if (Array.isArray(task.comments) && task.comments.length > 0) {
          task.comments
            .slice()
            .sort((a, b) => {
              const dateA = typeof a.createdAt === 'string' ? a.createdAt : '';
              const dateB = typeof b.createdAt === 'string' ? b.createdAt : '';
              return dateA.localeCompare(dateB);
            })
            .forEach((comment) => {
              if (!comment || typeof comment !== 'object' || !comment.content) {
                return;
              }
              const commentItem = document.createElement('li');
              commentItem.className = 'task-comment';
              const meta = document.createElement('span');
              meta.className = 'task-comment-meta';
              const authorLabel = formatMemberName(comment.author);
              let dateLabel = '';
              if (comment.createdAt && !Number.isNaN(new Date(comment.createdAt).getTime())) {
                dateLabel = capitalizeLabel(
                  taskCommentDateFormatter.format(new Date(comment.createdAt)),
                );
              }
              meta.textContent = dateLabel ? `${authorLabel} · ${dateLabel}` : authorLabel;
              const content = document.createElement('p');
              content.className = 'task-comment-content';
              content.textContent = comment.content;
              commentItem.append(meta, content);
              commentsList.appendChild(commentItem);
            });
        } else {
          const emptyMessage = document.createElement('p');
          emptyMessage.className = 'task-comment-empty';
          emptyMessage.textContent = 'Aucun commentaire pour le moment.';
          commentsSection.appendChild(emptyMessage);
        }

        commentsSection.appendChild(commentsList);

        const commentForm = document.createElement('form');
        commentForm.className = 'task-comment-form';
        commentForm.dataset.taskId = task.id || '';

        const commentLabel = document.createElement('label');
        commentLabel.className = 'sr-only';
        const commentFieldId = `task-comment-${task.id}`;
        commentLabel.setAttribute('for', commentFieldId);
        commentLabel.textContent = 'Ajouter un commentaire';

        const commentTextarea = document.createElement('textarea');
        commentTextarea.id = commentFieldId;
        commentTextarea.name = 'comment';
        commentTextarea.required = true;
        commentTextarea.placeholder = 'Écrire un nouveau commentaire';

        const commentSubmit = document.createElement('button');
        commentSubmit.type = 'submit';
        commentSubmit.className = 'primary-button task-comment-submit';
        commentSubmit.textContent = 'Publier';

        commentForm.append(commentLabel, commentTextarea, commentSubmit);
        commentsSection.appendChild(commentForm);

        listItem.append(header, descriptionEl, commentsSection);
        fragment.appendChild(listItem);
      });

      taskList.appendChild(fragment);
    }
	
	function renderTeamPage() {
	  // Remplir infos haut de page
	  const ownerEl = document.getElementById('team-owner-name');
	  if (ownerEl) ownerEl.textContent = data.panelOwner || '—';

	  const listEl = document.getElementById('team-list');
	  const emptyEl = document.getElementById('team-empty');
	  const countEl = document.getElementById('team-count');

	  if (!(listEl instanceof HTMLElement)) return;

	  // Contenu liste
	  const members = Array.isArray(data.teamMembers) ? data.teamMembers.slice() : [];
	  listEl.innerHTML = '';

	  if (!members.length) {
		if (emptyEl) emptyEl.hidden = false;
		if (countEl) countEl.textContent = '0';
	  } else {
		if (emptyEl) emptyEl.hidden = true;
		if (countEl) countEl.textContent = String(members.length);

		members.forEach((username) => {
		  const li = document.createElement('li');
		  li.className = 'category-item';

		  const main = document.createElement('div');
		  main.className = 'category-main';

		  const title = document.createElement('h3');
		  title.className = 'category-title';
		  title.textContent = username;

		  const desc = document.createElement('p');
		  desc.className = 'category-description';
		  desc.textContent = (username === data.panelOwner)
			? 'Fondateur du panel'
			: 'Membre';

		  main.appendChild(title);
		  main.appendChild(desc);

		  const actions = document.createElement('div');
		  actions.className = 'category-actions';

		  // Bouton retirer (sauf fondateur)
		  if (username !== data.panelOwner) {
			const removeBtn = document.createElement('button');
			removeBtn.type = 'button';
			removeBtn.className = 'category-button category-button--danger';
			removeBtn.textContent = 'Retirer';
			removeBtn.setAttribute('data-action', 'team-remove');
			removeBtn.setAttribute('data-username', username);
			actions.appendChild(removeBtn);
		  }


		  li.appendChild(main);
		  li.appendChild(actions);
		  listEl.appendChild(li);
		});
	  }

	  // Remplir le select "Ajouter un membre"
	  const select = document.getElementById('team-user-select');
	  if (select instanceof HTMLSelectElement) {
		const candidates = getAvailableUsersForTeam();
		select.innerHTML = '';
		if (candidates.length === 0) {
		  const opt = document.createElement('option');
		  opt.value = '';
		  opt.disabled = true;
		  opt.selected = true;
		  opt.textContent = 'Aucun utilisateur disponible';
		  select.appendChild(opt);
		  select.disabled = true;
		} else {
		  candidates.forEach(({ username, email }) => {
			const opt = document.createElement('option');
			opt.value = username;
			opt.textContent = email ? `${username} (${email})` : username;
			select.appendChild(opt);
		  });
		  select.disabled = false;
		}
	  }
	}
	
	function addTeamMember(username) {
	  if (!username) return;
	  if (!Array.isArray(data.teamMembers)) data.teamMembers = [];
	  if (!data.teamMembers.includes(username)) {
		data.teamMembers.push(username);
		data.teamMembers = Array.from(new Set(data.teamMembers));
		data.lastUpdated = new Date().toISOString();
		saveDataForUser(currentUser, data);
		rebuildTeamCaches();
		renderTeamPage();
	  }
	}

	function removeTeamMember(username) {
	  if (!username) return;
	  if (username === data.panelOwner) return; // Sécurité
	  if (!Array.isArray(data.teamMembers)) return;

	  const before = data.teamMembers.length;
	  data.teamMembers = data.teamMembers.filter((u) => u !== username);

	  if (data.teamMembers.length !== before) {
		data.lastUpdated = new Date().toISOString();
		saveDataForUser(currentUser, data);
		rebuildTeamCaches();
		renderTeamPage();
	  }
	}

    function updateTaskCountDisplay(count) {
      let badgeLabel = '';
      if (count === 0) {
        badgeLabel = 'Aucune tâche';
      } else if (count === 1) {
        badgeLabel = '1 tâche';
      } else {
        badgeLabel = `${count} tâches`;
      }

      if (taskCountBadge) {
        taskCountBadge.textContent = badgeLabel;
      }

      if (taskToggleButton) {
        const ariaLabel =
          count === 0 ? 'Liste des tâches (aucune tâche)' : `Liste des tâches (${badgeLabel})`;
        taskToggleButton.setAttribute('aria-label', ariaLabel);
        taskToggleButton.title = ariaLabel;
      }
    }

    function deleteTask(taskId) {
	  if (!taskId) return;

	  // Chercher la tâche pour savoir si elle a un évènement lié
	  const existingTask = Array.isArray(data.tasks)
		? data.tasks.find((t) => t && t.id === taskId)
		: null;

	  // Supprimer la tâche
	  const initialLength = Array.isArray(data.tasks) ? data.tasks.length : 0;
	  data.tasks = Array.isArray(data.tasks)
		? data.tasks.filter((task) => task && task.id !== taskId)
		: [];

	  if (data.tasks.length === initialLength) return;

	  // S’il y a un évènement lié, on le supprime aussi
	  if (existingTask && typeof existingTask.calendarEventId === 'string' && existingTask.calendarEventId) {
		deleteCalendarEvent(existingTask.calendarEventId);
		// deleteCalendarEvent() fait déjà save + renderCalendar()
	  }

	  data.lastUpdated = new Date().toISOString();
	  saveDataForUser(currentUser, data);
	  renderTasks();
	}
	
	// État local : id de la tâche en cours d’édition
	let editingTaskId = '';

	function setTaskFormMode(mode /* 'create' | 'edit' */) {
	  if (!taskForm) return;
	  const submitBtn = taskForm.querySelector('button[type="submit"]');
	  const resetBtn  = taskForm.querySelector('button[type="reset"]');

	  taskForm.dataset.mode = mode;
	  if (submitBtn instanceof HTMLButtonElement) {
		submitBtn.textContent = mode === 'edit' ? 'Enregistrer' : 'Ajouter la tâche';
	  }
	  if (resetBtn instanceof HTMLButtonElement) {
		resetBtn.textContent = mode === 'edit' ? 'Annuler la modification' : 'Réinitialiser';
	  }
	}

	function startEditTask(taskId) {
	  if (!Array.isArray(data.tasks)) return;
	  const t = data.tasks.find((x) => x && x.id === taskId);
	  if (!t) return;

	  // Ouvre la page "Tâches" si besoin
	  showPage && showPage('tasks');

	  // Remplit le formulaire
	  if (taskTitleInput instanceof HTMLInputElement) taskTitleInput.value = t.title || '';
	  if (taskDueDateInput instanceof HTMLInputElement) taskDueDateInput.value = t.dueDate || '';
	  if (taskColorInput instanceof HTMLInputElement) taskColorInput.value = normalizeTaskColor(t.color || DEFAULT_TASK_COLOR);
	  if (taskDescriptionInput instanceof HTMLTextAreaElement) taskDescriptionInput.value = t.description || '';

	  if (taskMemberSelect instanceof HTMLSelectElement && !taskMemberSelect.disabled) {
		Array.from(taskMemberSelect.options).forEach((opt) => {
		  opt.selected = Array.isArray(t.assignedMembers) ? t.assignedMembers.includes(opt.value) : false;
		});
	  }

	  editingTaskId = t.id;
	  setTaskFormMode('edit');

	  // Focus UX
	  window.requestAnimationFrame(() => {
		taskTitleInput && taskTitleInput.focus();
		taskTitleInput && taskTitleInput.setSelectionRange(taskTitleInput.value.length, taskTitleInput.value.length);
	  });
	}

	function applyTaskEditsFromForm() {
	  if (!taskForm || !editingTaskId) return;
	  if (!(taskTitleInput instanceof HTMLInputElement)) return;

	  const formData = new FormData(taskForm);
	  const title = (formData.get('task-title') || '').toString().trim();
	  if (!title) { taskTitleInput.focus(); return; }

	  // Date
	  let dueDate = '';
	  const dueDateRaw = (formData.get('task-due-date') || '').toString().trim();
	  if (taskDueDateInput instanceof HTMLInputElement) taskDueDateInput.setCustomValidity('');
	  if (dueDateRaw) {
		if (isValidDateKey(dueDateRaw)) {
		  dueDate = dueDateRaw; // attendu: YYYY-MM-DD
		} else if (taskDueDateInput instanceof HTMLInputElement) {
		  taskDueDateInput.setCustomValidity('Veuillez sélectionner une date valide.');
		  taskDueDateInput.reportValidity();
		  taskDueDateInput.focus();
		  return;
		}
	  }

	  const colorValue       = normalizeTaskColor((formData.get('task-color') || DEFAULT_TASK_COLOR).toString());
	  const descriptionValue = (formData.get('task-description') || '').toString().trim();

	  const members = [];
	  if (taskMemberSelect instanceof HTMLSelectElement && !taskMemberSelect.disabled) {
		Array.from(taskMemberSelect.selectedOptions).forEach((option) => {
		  if (option && option.value) members.push(option.value);
		});
	  }

	  // Récupérer & mettre à jour la tâche
	  const task = Array.isArray(data.tasks) ? data.tasks.find((x) => x && x.id === editingTaskId) : null;
	  if (!task) return;

	  task.title           = title;
	  task.dueDate         = dueDate;               // peut devenir ''
	  task.color           = colorValue;
	  task.description     = descriptionValue;
	  task.assignedMembers = members;

	  // >>> Synchronisation calendrier (IMMUTABLE + prise en charge du changement de date)
	  upsertCalendarEventForTask(task);

	  // Sauvegarde + rendu
	  data.tasks = data.tasks.map((it) => normalizeTask(it)).sort(compareTasks);
	  data.lastUpdated = new Date().toISOString();
	  saveDataForUser(currentUser, data);

	  if (typeof renderCalendar === 'function') renderCalendar();

	  // Reset mode édition
	  editingTaskId = '';
	  taskForm.reset();
	  resetTaskFormDefaults();
	  setTaskFormMode('create');
	  renderTasks();
	}



    function addCommentToTask(taskId, rawContent) {
      if (!taskId || typeof rawContent !== 'string') {
        return;
      }

      const content = rawContent.trim();
      if (!content) {
        return;
      }

      if (!Array.isArray(data.tasks)) {
        return;
      }

      const task = data.tasks.find((item) => item && item.id === taskId);
      if (!task) {
        return;
      }

      if (!Array.isArray(task.comments)) {
        task.comments = [];
      }

      const comment = normalizeTaskComment({
        id: generateId('comment'),
        author: currentUser,
        content,
        createdAt: new Date().toISOString(),
      });

      if (!comment) {
        return;
      }

      task.comments.push(comment);
      task.comments = task.comments
        .map((item) => normalizeTaskComment(item))
        .filter((item) => item && item.content);

      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderTasks();

      window.requestAnimationFrame(() => {
        const textarea = document.getElementById(`task-comment-${taskId}`);
        if (textarea instanceof HTMLTextAreaElement) {
          textarea.value = '';
          textarea.focus();
        }
      });
    }

    function formatTeamMemberLabel(member) {
      if (!member || typeof member !== 'object') {
        return '';
      }

      if (member.email) {
        return `${member.username} (${member.email})`;
      }

      return member.username;
    }

    function formatMemberName(username) {
      if (typeof username !== 'string' || username.trim() === '') {
        return 'Utilisateur';
      }
      const normalized = username.trim();
      const member = teamMembersById.get(normalized);
      if (member) {
        return formatTeamMemberLabel(member);
      }
      return normalized;
    }

    function notifyDataChanged(section, detail = {}) {
      if (!section) {
        return;
      }

      const eventDetail = { section, ...detail };
      document.dispatchEvent(new CustomEvent('umanager:data-changed', { detail: eventDetail }));
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

    function normalizeComparableValue(value) {
      if (value === undefined || value === null) {
        return '';
      }
      return value
        .toString()
        .trim()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
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
        const currentValue = element.value != null ? element.value.toString() : '';
        if (categoryId && currentValue !== '') {
          previousValues.set(categoryId, currentValue);
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

      const totalResults = filteredContacts.length;

      if (contactSearchCountEl) {
        contactSearchCountEl.textContent = totalResults.toString();
      }

      if (contactPagination) {
        contactPagination.hidden = totalResults === 0;
      }

      if (totalResults === 0) {
        contactEmptyState.hidden = false;
        if (contacts.length === 0) {
          contactEmptyState.textContent = 'Ajoutez vos premiers contacts pour les retrouver ici.';
        } else if (normalizedTerm || hasAdvancedFilters) {
          contactEmptyState.textContent = 'Aucun contact ne correspond à vos critères de recherche.';
        } else {
          contactEmptyState.textContent = 'Aucun contact à afficher pour le moment.';
        }
        if (contactPaginationSummary) {
          contactPaginationSummary.textContent = 'Affichage 0 sur 0 contacts';
        }
        if (contactPaginationPageLabel) {
          contactPaginationPageLabel.textContent = 'Page 1 / 1';
        }
        if (contactPaginationPrevButton instanceof HTMLButtonElement) {
          contactPaginationPrevButton.disabled = true;
        }
        if (contactPaginationNextButton instanceof HTMLButtonElement) {
          contactPaginationNextButton.disabled = true;
        }
        contactCurrentPage = 1;
        return;
      }

      contactEmptyState.hidden = true;

      const normalizedPerPage = Number.isFinite(contactResultsPerPage)
        ? Math.max(1, Math.floor(contactResultsPerPage))
        : CONTACT_RESULTS_PER_PAGE_DEFAULT;
      if (normalizedPerPage !== contactResultsPerPage) {
        contactResultsPerPage = normalizedPerPage;
      }

      const totalPages = Math.max(1, Math.ceil(totalResults / contactResultsPerPage));
      if (contactCurrentPage > totalPages) {
        contactCurrentPage = totalPages;
      } else if (contactCurrentPage < 1) {
        contactCurrentPage = 1;
      }

      let startIndex = (contactCurrentPage - 1) * contactResultsPerPage;
      let endIndex = Math.min(startIndex + contactResultsPerPage, totalResults);
      let pageContacts = filteredContacts.slice(startIndex, endIndex);

      if (pageContacts.length === 0) {
        contactCurrentPage = 1;
        startIndex = 0;
        endIndex = Math.min(contactResultsPerPage, totalResults);
        pageContacts = filteredContacts.slice(startIndex, endIndex);
      }

      const startNumber = startIndex + 1;
      const endNumber = startIndex + pageContacts.length;

      if (contactPaginationSummary) {
        const summaryStart = numberFormatter.format(startNumber);
        const summaryEnd = numberFormatter.format(endNumber);
        const totalLabel = numberFormatter.format(totalResults);
        const suffix = totalResults === 1 ? 'contact' : 'contacts';
        contactPaginationSummary.textContent = `Affichage ${summaryStart} – ${summaryEnd} sur ${totalLabel} ${suffix}`;
      }

      if (contactPaginationPageLabel) {
        contactPaginationPageLabel.textContent = `Page ${contactCurrentPage} / ${totalPages}`;
      }

      if (contactPaginationPrevButton instanceof HTMLButtonElement) {
        contactPaginationPrevButton.disabled = contactCurrentPage <= 1;
      }

      if (contactPaginationNextButton instanceof HTMLButtonElement) {
        contactPaginationNextButton.disabled = contactCurrentPage >= totalPages;
      }

      if (
        contactResultsPerPageSelect instanceof HTMLSelectElement &&
        contactResultsPerPageSelect.value !== contactResultsPerPage.toString()
      ) {
        contactResultsPerPageSelect.value = contactResultsPerPage.toString();
      }

      const fragment = document.createDocumentFragment();
      const dateTimeFormatter = new Intl.DateTimeFormat('fr-FR', {
        dateStyle: 'medium',
        timeStyle: 'short',
      });

      pageContacts.forEach((contact) => {
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

    function importContactsFromRows(rows, options = {}) {
      const safeRows = Array.isArray(rows) ? rows : [];
      const mappingSource =
        options && typeof options === 'object' && options.mapping ? options.mapping : {};
      const skipHeader = Boolean(options && options.skipHeader);
      const fileName =
        options && typeof options.fileName === 'string' ? options.fileName : '';

      const normalizeListValue = (raw) =>
        raw
          .toString()
          .trim()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .toLowerCase();

      const categoriesById = buildCategoryMap();
      if (!Array.isArray(data.contacts)) {
        data.contacts = [];
      }
      const normalizedMapping = Object.entries(mappingSource)
        .map(([categoryId, rawIndex]) => {
          const columnIndex = Number(rawIndex);
          if (!Number.isInteger(columnIndex) || columnIndex < 0) {
            return null;
          }

          const category = categoriesById.get(categoryId);
          if (!category) {
            return null;
          }

          const baseType = CATEGORY_TYPES.has(category.type) ? category.type : 'text';
          const listOptions =
            baseType === 'list' && Array.isArray(category.options)
              ? category.options
                  .map((option) => (option != null ? option.toString().trim() : ''))
                  .filter((option) => Boolean(option))
              : [];

          return {
            categoryId,
            columnIndex,
            name: category.name || '',
            type: baseType,
            listOptions: listOptions.map((value) => ({
              value,
              normalized: normalizeListValue(value),
            })),
          };
        })
        .filter((entry) => Boolean(entry));

      const enrichedMapping = normalizedMapping.map((mapping) => {
        const normalizedLabel = normalizeLabel(mapping.name || '');
        const isStrongIdentifier =
          EMAIL_KEYWORDS.some((keyword) => normalizedLabel.includes(keyword)) ||
          PHONE_KEYWORDS.some((keyword) => normalizedLabel.includes(keyword));
        const isNameField = NAME_KEYWORDS.some((keyword) => normalizedLabel.includes(keyword));
        return {
          ...mapping,
          normalizedLabel,
          isStrongIdentifier,
          isNameField,
        };
      });

      const startIndex = skipHeader ? 1 : 0;
      const totalRows = startIndex < safeRows.length ? safeRows.length - startIndex : 0;
      const errors = [];
      const errorRows = new Set();
      let importedCount = 0;
      let mergedCount = 0;
      let skippedEmptyCount = 0;
      let updatedContactsCount = 0;

      if (enrichedMapping.length === 0 || safeRows.length === 0 || totalRows === 0) {
        return {
          importedCount,
          mergedCount,
          skippedEmptyCount: totalRows,
          totalRows,
          errorCount: 0,
          errors,
          fileName,
        };
      }

      const existingContactsDetails = data.contacts.map((contact) => {
        const normalizedValues = new Map();
        if (contact && typeof contact === 'object' && contact.categoryValues) {
          Object.entries(contact.categoryValues).forEach(([categoryId, rawValue]) => {
            const normalized = normalizeComparableValue(rawValue);
            if (normalized) {
              normalizedValues.set(categoryId, normalized);
            }
          });
        }
        const normalizedDisplayName = normalizeComparableValue(
          getContactDisplayName(contact, categoriesById),
        );
        return { contact, normalizedValues, normalizedDisplayName };
      });

      const findMatchingContact = (normalizedValues, normalizedDerivedName) => {
        if (!(normalizedValues instanceof Map) || normalizedValues.size === 0) {
          return null;
        }

        let bestMatch = null;
        let bestScore = 0;

        existingContactsDetails.forEach((entry) => {
          if (!entry || !entry.contact) {
            return;
          }

          const existingValues = entry.normalizedValues;
          let strongMatches = 0;
          let regularMatches = 0;
          let nameFieldMatches = 0;
          let hasStrongConflict = false;

          enrichedMapping.forEach((mapping) => {
            const newValue = normalizedValues.get(mapping.categoryId);
            if (!newValue) {
              return;
            }
            const existingValue = existingValues.get(mapping.categoryId);
            if (!existingValue) {
              return;
            }
            if (existingValue === newValue) {
              if (mapping.isStrongIdentifier) {
                strongMatches += 1;
              } else {
                regularMatches += 1;
                if (mapping.isNameField) {
                  nameFieldMatches += 1;
                }
              }
            } else if (mapping.isStrongIdentifier) {
              hasStrongConflict = true;
            }
          });

          const totalMatches = strongMatches + regularMatches;
          let nameMatch = false;
          if (
            nameFieldMatches > 0 &&
            normalizedDerivedName &&
            entry.normalizedDisplayName &&
            entry.normalizedDisplayName === normalizedDerivedName
          ) {
            nameMatch = true;
          }

          if (hasStrongConflict) {
            return;
          }

          if (
            strongMatches === 0 &&
            totalMatches < 2 &&
            !(nameMatch && totalMatches >= 1)
          ) {
            return;
          }

          const score = strongMatches * 10 + regularMatches + (nameMatch ? 1 : 0);
          if (score > bestScore) {
            bestScore = score;
            bestMatch = entry;
          }
        });

        return bestMatch;
      };

      for (let rowIndex = startIndex; rowIndex < safeRows.length; rowIndex += 1) {
        const row = Array.isArray(safeRows[rowIndex]) ? safeRows[rowIndex] : [];
        const categoryValues = {};
        let hasValue = false;

        enrichedMapping.forEach((mapping) => {
          const cellValue = row[mapping.columnIndex];
          const valueString =
            cellValue === undefined || cellValue === null ? '' : cellValue.toString().trim();

          if (!valueString) {
            return;
          }

          if (mapping.type === 'number') {
            const normalizedNumber = valueString.replace(/\s+/g, '').replace(',', '.');
            const parsed = Number(normalizedNumber);
            if (!Number.isFinite(parsed)) {
              errors.push({
                row: rowIndex + 1,
                categoryId: mapping.categoryId,
                message: `La valeur « ${valueString} » n’est pas un nombre valide pour la catégorie « ${mapping.name} ».`,
              });
              errorRows.add(rowIndex + 1);
              return;
            }
            categoryValues[mapping.categoryId] = parsed.toString();
            hasValue = true;
            return;
          }

          if (mapping.type === 'list') {
            const normalizedValue = normalizeListValue(valueString);
            const matchedOption = mapping.listOptions.find(
              (option) => option.normalized === normalizedValue,
            );
            if (!matchedOption) {
              errors.push({
                row: rowIndex + 1,
                categoryId: mapping.categoryId,
                message: `La valeur « ${valueString} » ne correspond à aucune option de la catégorie « ${mapping.name} ».`,
              });
              errorRows.add(rowIndex + 1);
              return;
            }
            categoryValues[mapping.categoryId] = matchedOption.value;
            hasValue = true;
            return;
          }

          categoryValues[mapping.categoryId] = valueString;
          hasValue = true;
        });

        if (!hasValue) {
          skippedEmptyCount += 1;
          continue;
        }

        const normalizedRowValues = new Map();
        enrichedMapping.forEach((mapping) => {
          const rawValue = categoryValues[mapping.categoryId];
          if (rawValue === undefined || rawValue === null) {
            return;
          }
          const normalizedValue = normalizeComparableValue(rawValue);
          if (normalizedValue) {
            normalizedRowValues.set(mapping.categoryId, normalizedValue);
          }
        });

        const derivedName = buildDisplayNameFromCategories(categoryValues, categoriesById);
        const displayName = derivedName || 'Contact sans nom';
        const normalizedDerivedName = normalizeComparableValue(derivedName);
        const matchingEntry = findMatchingContact(normalizedRowValues, normalizedDerivedName);

        if (matchingEntry) {
          const contactToUpdate = matchingEntry.contact;
          if (contactToUpdate && typeof contactToUpdate === 'object') {
            if (!contactToUpdate.categoryValues || typeof contactToUpdate.categoryValues !== 'object') {
              contactToUpdate.categoryValues = {};
            }
            let contactModified = false;

            enrichedMapping.forEach((mapping) => {
              const newValue = categoryValues[mapping.categoryId];
              if (newValue === undefined || newValue === null || newValue === '') {
                return;
              }

              const existingValue = contactToUpdate.categoryValues[mapping.categoryId];
              if (
                existingValue === undefined ||
                existingValue === null ||
                existingValue === ''
              ) {
                contactToUpdate.categoryValues[mapping.categoryId] = newValue;
                matchingEntry.normalizedValues.set(
                  mapping.categoryId,
                  normalizeComparableValue(newValue),
                );
                contactModified = true;
                return;
              }

              const existingNormalized = normalizeComparableValue(existingValue);
              const newNormalized = normalizeComparableValue(newValue);
              if (existingNormalized === newNormalized) {
                return;
              }

              if (!mapping.isStrongIdentifier) {
                contactToUpdate.categoryValues[mapping.categoryId] = newValue;
                matchingEntry.normalizedValues.set(mapping.categoryId, newNormalized);
                contactModified = true;
              }
            });

            if (contactModified) {
              const updatedDerivedName = buildDisplayNameFromCategories(
                contactToUpdate.categoryValues,
                categoriesById,
              );
              if (updatedDerivedName) {
                contactToUpdate.fullName = updatedDerivedName;
                contactToUpdate.displayName = updatedDerivedName;
              }
              contactToUpdate.updatedAt = new Date().toISOString();
              matchingEntry.normalizedDisplayName = normalizeComparableValue(
                getContactDisplayName(contactToUpdate, categoriesById),
              );
              updatedContactsCount += 1;
            }
          }

        }

        if (matchingEntry) {
          mergedCount += 1;
          continue;
        }

        const nowIso = new Date().toISOString();
        const newContact = {
          id: generateId('contact'),
          categoryValues,
          keywords: [],
          notes: '',
          fullName: displayName,
          displayName,
          createdAt: nowIso,
          updatedAt: null,
        };
        data.contacts.push(newContact);
        importedCount += 1;
        existingContactsDetails.push({
          contact: newContact,
          normalizedValues: normalizedRowValues,
          normalizedDisplayName: normalizeComparableValue(displayName),
        });
      }

      const hasDataChanges = importedCount > 0 || updatedContactsCount > 0;
      if (hasDataChanges) {
        data.lastUpdated = new Date().toISOString();
        updateMetricsFromContacts();
        saveDataForUser(currentUser, data);
        renderMetrics();
        renderContacts();
      }

      if (importedCount > 0 || mergedCount > 0) {
        notifyDataChanged('contacts', {
          reason: 'import',
          importedCount,
          mergedCount,
          skippedEmptyCount,
          totalRows,
          errorCount: errorRows.size,
          fileName,
        });
      }

      return {
        importedCount,
        mergedCount,
        skippedEmptyCount,
        totalRows,
        errorCount: errorRows.size,
        errors,
        fileName,
      };
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

    function normalizeTaskColor(rawColor) {
      if (typeof rawColor !== 'string') {
        return DEFAULT_TASK_COLOR;
      }

      const value = rawColor.trim();
      const hexPattern = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;
      if (!hexPattern.test(value)) {
        return DEFAULT_TASK_COLOR;
      }

      if (value.length === 4) {
        const [, r, g, b] = value;
        return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
      }

      return value.length === 7 ? value.toLowerCase() : DEFAULT_TASK_COLOR;
    }

    function normalizeTask(rawTask) {
      const baseTask = rawTask && typeof rawTask === 'object' ? rawTask : {};

      let id = '';
      if (typeof baseTask.id === 'string' && baseTask.id.trim() !== '') {
        id = baseTask.id.trim();
      } else if (typeof baseTask.identifier === 'string' && baseTask.identifier.trim() !== '') {
        id = baseTask.identifier.trim();
      } else {
        id = generateId('task');
      }

      let title = '';
      if (typeof baseTask.title === 'string') {
        title = baseTask.title.trim();
      } else if (typeof baseTask.name === 'string') {
        title = baseTask.name.trim();
      }
      if (!title) {
        title = 'Tâche sans titre';
      }

      let dueDate = '';
      if (typeof baseTask.dueDate === 'string' && isValidDateKey(baseTask.dueDate.trim())) {
        dueDate = baseTask.dueDate.trim();
      } else if (
        typeof baseTask.deadline === 'string' &&
        isValidDateKey(baseTask.deadline.trim())
      ) {
        dueDate = baseTask.deadline.trim();
      }

      let description = '';
      if (typeof baseTask.description === 'string') {
        description = baseTask.description.trim();
      } else if (typeof baseTask.details === 'string') {
        description = baseTask.details.trim();
      }

      let createdAt = '';
      if (
        typeof baseTask.createdAt === 'string' &&
        !Number.isNaN(new Date(baseTask.createdAt).getTime())
      ) {
        createdAt = baseTask.createdAt;
      } else if (
        typeof baseTask.created === 'string' &&
        !Number.isNaN(new Date(baseTask.created).getTime())
      ) {
        createdAt = baseTask.created;
      } else {
        createdAt = new Date().toISOString();
      }

      let createdBy = '';
      if (typeof baseTask.createdBy === 'string') {
        createdBy = baseTask.createdBy.trim();
      } else if (typeof baseTask.owner === 'string') {
        createdBy = baseTask.owner.trim();
      }

      const assignedMembers = Array.isArray(baseTask.assignedMembers)
        ? Array.from(
            new Set(
              baseTask.assignedMembers
                .map((member) => (typeof member === 'string' ? member.trim() : ''))
                .filter(Boolean),
            ),
          )
        : [];

      const comments = Array.isArray(baseTask.comments)
        ? baseTask.comments
            .map((comment) => normalizeTaskComment(comment))
            .filter((comment) => comment && comment.content)
        : [];
		
	  let calendarEventId = '';
		if (typeof baseTask.calendarEventId === 'string' && baseTask.calendarEventId.trim() !== '') {
		  calendarEventId = baseTask.calendarEventId.trim();
		}

      return {
        id,
        title,
        dueDate,
        color: normalizeTaskColor(baseTask.color),
        description,
        assignedMembers,
        createdAt,
        createdBy,
        comments,
		calendarEventId,
      };
    }

    function normalizeTaskComment(rawComment) {
      const baseComment = rawComment && typeof rawComment === 'object' ? rawComment : {};

      let content = '';
      if (typeof baseComment.content === 'string') {
        content = baseComment.content.trim();
      } else if (typeof baseComment.message === 'string') {
        content = baseComment.message.trim();
      }

      if (!content) {
        return null;
      }

      let author = '';
      if (typeof baseComment.author === 'string') {
        author = baseComment.author.trim();
      } else if (typeof baseComment.user === 'string') {
        author = baseComment.user.trim();
      }

      let createdAt = '';
      if (
        typeof baseComment.createdAt === 'string' &&
        !Number.isNaN(new Date(baseComment.createdAt).getTime())
      ) {
        createdAt = baseComment.createdAt;
      } else if (
        typeof baseComment.date === 'string' &&
        !Number.isNaN(new Date(baseComment.date).getTime())
      ) {
        createdAt = baseComment.date;
      } else {
        createdAt = new Date().toISOString();
      }

      let id = '';
      if (typeof baseComment.id === 'string' && baseComment.id.trim() !== '') {
        id = baseComment.id.trim();
      } else {
        id = generateId('comment');
      }

      return {
        id,
        author,
        content,
        createdAt,
      };
    }

    function compareTasks(a, b) {
      if (!a || !b) {
        return 0;
      }

      const dateA =
        typeof a.dueDate === 'string' && isValidDateKey(a.dueDate) ? a.dueDate : '';
      const dateB =
        typeof b.dueDate === 'string' && isValidDateKey(b.dueDate) ? b.dueDate : '';

      if (dateA && dateB && dateA !== dateB) {
        return dateA.localeCompare(dateB);
      }

      if (dateA && !dateB) {
        return -1;
      }

      if (!dateA && dateB) {
        return 1;
      }

      const createdAtA = typeof a.createdAt === 'string' ? a.createdAt : '';
      const createdAtB = typeof b.createdAt === 'string' ? b.createdAt : '';

      if (createdAtA && createdAtB && createdAtA !== createdAtB) {
        return createdAtA.localeCompare(createdAtB);
      }

      const titleA = typeof a.title === 'string' ? a.title : '';
      const titleB = typeof b.title === 'string' ? b.title : '';
      return titleA.localeCompare(titleB, 'fr', { sensitivity: 'base' });
    }

    function hexToRgba(color, alpha = 0.16) {
      if (typeof color !== 'string') {
        return `rgba(37, 99, 235, ${alpha})`;
      }

      const normalized = color.trim().toLowerCase();
      const match = /^#([0-9a-f]{6})$/.exec(normalized);
      if (!match) {
        return `rgba(37, 99, 235, ${alpha})`;
      }

      const hex = match[1];
      const r = Number.parseInt(hex.slice(0, 2), 16);
      const g = Number.parseInt(hex.slice(2, 4), 16);
      const b = Number.parseInt(hex.slice(4, 6), 16);
      return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }

    function upgradeDataStructure(rawData) {
      const base = rawData && typeof rawData === 'object' ? rawData : {};
	  // -- PRÉSERVER / NORMALISER LE FONDATEUR ET L'ÉQUIPE -------------------------
	  base.panelOwner = (typeof base.panelOwner === 'string' && base.panelOwner.trim())
	    ? base.panelOwner.trim()
	    : (base.panelOwner || ''); // ne crée pas ici, juste préserve si présent

	  if (Array.isArray(base.teamMembers)) {
		base.teamMembers = Array.from(new Set(
		  base.teamMembers
		    .filter(u => typeof u === 'string')
		    .map(u => u.trim())
		    .filter(Boolean)
		));
	  } else {
		base.teamMembers = [];
	  }
		// ---------------------------------------------------------------------------


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

      if (!Array.isArray(base.events)) {
        base.events = [];
      } else {
        base.events = base.events
          .filter((item) => item && typeof item === 'object')
          .map((item) => normalizeCalendarEvent(item))
          .sort(compareCalendarEvents);
      }

      if (!Array.isArray(base.tasks)) {
        base.tasks = [];
      } else {
        base.tasks = base.tasks
          .filter((item) => item && typeof item === 'object')
          .map((item) => normalizeTask(item))
          .sort(compareTasks);
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
		  const parsed = JSON.parse(stored) || {};
		  // On préserve TOUT ce qui est présent, puis on normalise les blocs connus
		  const base = typeof parsed === 'object' ? parsed : {};

		  return {
			// préserve tous les champs déjà sauvés (dont panelOwner, teamMembers, etc.)
			...base,

			// normalisations sans perdre de données
			metrics: { ...defaultData.metrics, ...(base.metrics || {}) },
			categories: Array.isArray(base.categories) ? base.categories : [],
			keywords: Array.isArray(base.keywords) ? base.keywords : [],
			contacts: Array.isArray(base.contacts) ? base.contacts : [],
			events: Array.isArray(base.events) ? base.events : [],
			tasks: Array.isArray(base.tasks) ? base.tasks : [],
			lastUpdated: base.lastUpdated || null,

			// si jamais ces champs n’existaient pas encore
			panelOwner: typeof base.panelOwner === 'string' ? base.panelOwner : '',
			teamMembers: Array.isArray(base.teamMembers) ? base.teamMembers : [],
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
		events: [],
		tasks: [],
		lastUpdated: null,
		// Nouveaux champs persistés
		panelOwner: '',
		teamMembers: [],
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
