(() => {
  const USER_STORE_KEY = 'umanager-user-store';
  const ACTIVE_USER_KEY = 'umanager-active-user';
  const DATA_KEY_PREFIX = 'umanager-data-store:';
  const SUPER_ADMIN_EMAIL = 'fasquellesteven@gmail.com';
  const NEWS_STORE_KEY = 'umanager-news-store';
  const NEWS_DISPLAY_LIMIT = 5;
  const DEFAULT_SUBSCRIPTION_ID = 'none';
  const SUBSCRIPTION_PLAN_ORDER = ['none', 'essential', 'pro', 'ultimate'];
  const SUBSCRIPTION_PLANS = {
    none: {
      id: 'none',
      name: 'Aucun abonnement',
      priceLabel: '0 € / mois',
      contactLimit: 5,
      taskLimit: 5,
      features: ['5 contacts max', '5 tâches actives max'],
    },
    essential: {
      id: 'essential',
      name: 'Essentiel',
      priceLabel: '5 € / mois',
      contactLimit: 25,
      taskLimit: 10,
      features: ['25 contacts max', '10 tâches actives max'],
    },
    pro: {
      id: 'pro',
      name: 'Pro',
      priceLabel: '20 € / mois',
      contactLimit: 250,
      taskLimit: 50,
      features: ['250 contacts max', '50 tâches actives max', "Accès au tchat d'équipe"],
    },
    ultimate: {
      id: 'ultimate',
      name: 'Ultimate',
      priceLabel: '100 € / mois',
      contactLimit: 100000,
      taskLimit: 500,
      features: [
        '100 000 contacts max',
        '500 tâches actives max',
        "Accès au tchat d'équipe",
        'Accès au partage de document',
      ],
    },
  };

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
    taskCategories: [],
    tasks: [],
    teamChatMessages: [],
    savedSearches: [],
    emailTemplates: [],
    emailCampaigns: [],
    lastUpdated: null,
    subscription: DEFAULT_SUBSCRIPTION_ID,
    teams: [],
    currentTeamId: '',
  };

  const TEAM_CHAT_HISTORY_LIMIT = 200;
  const TEAM_CHAT_MAX_MESSAGE_LENGTH = 500;
  const TASK_MEMBER_NONE_VALUE = '__none__';

  const REQUIRED_CATEGORY_IDS = {
    firstName: 'contact-first-name',
    lastName: 'contact-last-name',
    birthDate: 'contact-birth-date',
    identifier: 'contact-identifier',
  };

  const REQUIRED_CONTACT_CATEGORIES = [
    { id: REQUIRED_CATEGORY_IDS.firstName, name: 'Prénom', type: 'text' },
    { id: REQUIRED_CATEGORY_IDS.lastName, name: 'Nom', type: 'text' },
    { id: REQUIRED_CATEGORY_IDS.birthDate, name: 'Date de naissance', type: 'date' },
    { id: REQUIRED_CATEGORY_IDS.identifier, name: 'Identifiant', type: 'text' },
  ];

  const REQUIRED_CATEGORY_ID_SET = new Set(
    REQUIRED_CONTACT_CATEGORIES.map((category) => category.id),
  );

  const CONTACT_TYPE_OPTIONS = [
    { value: 'person', label: 'Personne', badgeClass: 'contact-type-badge--person' },
    { value: 'company', label: 'Entreprise', badgeClass: 'contact-type-badge--company' },
    { value: 'association', label: 'Association', badgeClass: 'contact-type-badge--association' },
    { value: 'institution', label: 'Institution', badgeClass: 'contact-type-badge--institution' },
  ];
  const CONTACT_TYPE_DEFAULT = 'person';
  const CONTACT_TYPE_MAP = new Map(
    CONTACT_TYPE_OPTIONS.map((option) => [option.value, option]),
  );

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
    const topbarButtons = Array.from(document.querySelectorAll('.topbar-button'));
    const contextNav = document.getElementById('context-nav');
    const contextNavList = document.getElementById('context-nav-list');
    const contextEmptyState = document.getElementById('context-empty-state');
    const contextEmptyDefaultText = contextEmptyState
      ? contextEmptyState.textContent || ''
      : '';
    const summaryNavigateButtons = Array.from(document.querySelectorAll('[data-navigate]'));
    const taskSummaryActiveEl = document.getElementById('task-summary-active');
    const taskSummaryArchivedEl = document.getElementById('task-summary-archived');
    const taskStatusButtons = Array.from(document.querySelectorAll('.task-status-button'));
    const taskStatusCountEls = {
      active: document.querySelector('[data-task-status-count="active"]'),
      archived: document.querySelector('[data-task-status-count="archived"]'),
    };
    const homeCurrentTeamEl = document.getElementById('home-current-team');
    const homeTeamsList = document.getElementById('home-teams-list');
    const homeTeamsEmpty = document.getElementById('home-teams-empty');
    const homeTeamForm = document.getElementById('home-team-form');
    const homeTeamNameInput = document.getElementById('home-team-name');
    const homeTeamRoleInput = document.getElementById('home-team-role');
    const homeTeamFeedback = document.getElementById('home-team-feedback');
    const homeSubscriptionNameEl = document.getElementById('home-subscription-name');
    const homeSubscriptionPriceEl = document.getElementById('home-subscription-price');
    const homeSubscriptionContactLimitEl = document.getElementById('home-subscription-contact-limit');
    const homeSubscriptionTaskLimitEl = document.getElementById('home-subscription-task-limit');
    const homeSubscriptionSelect = document.getElementById('home-subscription-select');
    const homeSubscriptionFeaturesList = document.getElementById('home-subscription-features');
    const homeDirectoryLimitEl = document.getElementById('home-directory-limit');
    const homeDirectoryProgressBar = document.getElementById('home-directory-progress-bar');
    const homeDirectoryUsageEl = document.getElementById('home-directory-usage');
    const homeNewsList = document.getElementById('home-news-list');
    const homeNewsEmpty = document.getElementById('home-news-empty');
    const homeNewsUpdatedEl = document.getElementById('home-news-updated');
    const topbarHomeButton = document.getElementById('topbar-home');
    const topbarTeamButton = document.getElementById('topbar-team');
    const topbarAdministrationButton = document.getElementById('topbar-administration');
    const adminArticleForm = document.getElementById('admin-article-form');
    const adminArticleFeedback = document.getElementById('admin-article-feedback');
    const adminNewsList = document.getElementById('admin-news-list');
    const adminNewsEmpty = document.getElementById('admin-news-empty');
    const adminImpersonateForm = document.getElementById('admin-impersonate-form');
    const adminImpersonateSelect = document.getElementById('admin-impersonate-select');
    const adminImpersonateFeedback = document.getElementById('admin-impersonate-feedback');
    const homeNewsUpdatedDefault = homeNewsUpdatedEl ? homeNewsUpdatedEl.textContent || '' : '';
    const adminArticleFeedbackDefault =
      adminArticleFeedback ? adminArticleFeedback.textContent || '' : '';
    const adminImpersonateFeedbackDefault =
      adminImpersonateFeedback ? adminImpersonateFeedback.textContent || '' : '';

    if (homeTeamFeedback) {
      homeTeamFeedback.textContent = '';
      homeTeamFeedback.hidden = true;
      homeTeamFeedback.setAttribute('hidden', '');
    }

    if (adminArticleFeedback) {
      adminArticleFeedback.textContent = '';
      adminArticleFeedback.hidden = true;
      adminArticleFeedback.setAttribute('hidden', '');
    }

    if (adminImpersonateFeedback) {
      adminImpersonateFeedback.textContent = '';
      adminImpersonateFeedback.hidden = true;
      adminImpersonateFeedback.setAttribute('hidden', '');
    }

    let userStore = loadUserStore();
    const currentUserRecord =
      userStore &&
      typeof userStore === 'object' &&
      userStore.users &&
      typeof userStore.users === 'object'
        ? userStore.users[currentUser]
        : null;
    const currentUserEmail =
      currentUserRecord && typeof currentUserRecord.email === 'string'
        ? currentUserRecord.email.trim()
        : '';
    const normalizedAdminEmail = SUPER_ADMIN_EMAIL.toLowerCase();
    const isSuperAdmin = currentUserEmail.toLowerCase() === normalizedAdminEmail;

    const MODULE_CONFIG = {
      home: [{ id: 'home-overview', label: "Vue d'ensemble" }],
      tasks: [
        { id: 'tasks-summary', label: 'Page de récapitulatif' },
        { id: 'tasks-organization', label: 'Organisation des tâches' },
        { id: 'tasks-create', label: 'Créer une tâche' },
        { id: 'tasks-list', label: 'Recherche de tâches' },
        { id: 'team-chat', label: 'Tchat interne' },
      ],
      directory: [
        { id: 'directory-home', label: 'Accueil du répertoire' },
        { id: 'categories', label: 'Gestion des catégories' },
        { id: 'keywords', label: 'Gestion des mots clés' },
        { id: 'contacts-add', label: 'Ajouter des contacts' },
        { id: 'contacts-import', label: 'Importer des contacts' },
        { id: 'contacts-search', label: 'Rechercher des contacts' },
      ],
      communication: [
        { id: 'email-templates', label: 'Créer un modèle mail' },
        { id: 'email-campaigns', label: 'Envoyer une campagne de mail' },
      ],
      team: [{ id: 'team', label: "Gestion de l'équipe" }],
    };

    if (isSuperAdmin) {
      MODULE_CONFIG.administration = [
        { id: 'administration', label: 'Administration générale' },
      ];
    }

    const pageModuleMap = new Map();
    Object.entries(MODULE_CONFIG).forEach(([moduleId, entries]) => {
      entries.forEach((entry) => {
        if (entry && entry.id) {
          pageModuleMap.set(entry.id, moduleId);
        }
      });
    });

    if (!isSuperAdmin) {
      if (
        topbarAdministrationButton instanceof HTMLElement &&
        topbarAdministrationButton.parentElement
      ) {
        topbarAdministrationButton.parentElement.removeChild(topbarAdministrationButton);
      }
      pageModuleMap.delete('administration');
      const adminIndex = topbarButtons.indexOf(topbarAdministrationButton);
      if (adminIndex >= 0) {
        topbarButtons.splice(adminIndex, 1);
      }
    }

    const TASK_STATUS_VALUES = new Set(['active', 'archived']);
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
    const coverageHighlightEl = document.getElementById('contact-coverage-highlight');
    const coverageSummaryEl = document.getElementById('contact-coverage-summary');
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
    const keywordStatsOverlay = document.getElementById('keyword-stats-overlay');
    const keywordStatsDialog = document.getElementById('keyword-stats-dialog');
    const keywordStatsTitle = document.getElementById('keyword-stats-title');
    const keywordStatsPercentEl = document.getElementById('keyword-stats-percent');
    const keywordStatsAssignedCountEl = document.getElementById('keyword-stats-assigned-count');
    const keywordStatsAssignedPercentEl = document.getElementById('keyword-stats-assigned-percent');
    const keywordStatsRemainingCountEl = document.getElementById('keyword-stats-remaining-count');
    const keywordStatsRemainingPercentEl = document.getElementById('keyword-stats-remaining-percent');
    const keywordStatsTotalEl = document.getElementById('keyword-stats-total');
    const keywordStatsSummaryEl = document.getElementById('keyword-stats-summary');
    const keywordStatsChart = document.getElementById('keyword-stats-chart');
    const keywordStatsCloseButtons = Array.from(
      document.querySelectorAll('[data-close-keyword-stats]'),
    ).filter((element) => element instanceof HTMLButtonElement);
    const contactsCountEl = document.getElementById('contacts-count');
    const contactForm = document.getElementById('contact-form');
    const contactTypeSelect = document.getElementById('contact-type');
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
    const searchKeywordsContainer = document.getElementById('search-keywords');
    const searchKeywordModeInputs = Array.from(
      document.querySelectorAll('input[name="search-keyword-mode"]'),
    ).filter((element) => element instanceof HTMLInputElement);
    const contactSearchCountEl = document.getElementById('contact-search-count');
    const contactSaveSearchButton = document.getElementById('contact-save-search');
    const contactSavedSearchSelect = document.getElementById('contact-saved-searches');
    const contactSaveSearchFeedback = document.getElementById('contact-save-search-feedback');
    const savedSearchCountEl = document.getElementById('saved-search-count');
    const savedSearchList = document.getElementById('saved-search-list');
    const savedSearchEmptyState = document.getElementById('saved-search-empty');
    const emailTemplateForm = document.getElementById('email-template-form');
    const emailTemplateBlocksContainer = document.getElementById('email-template-blocks');
    const emailTemplateAddParagraphButton = document.getElementById('email-template-add-paragraph');
    const emailTemplateAddImageButton = document.getElementById('email-template-add-image');
    const emailTemplateAddButtonButton = document.getElementById('email-template-add-button');
    const emailTemplatePreview = document.getElementById('email-template-preview');
    const emailTemplateFeedback = document.getElementById('email-template-feedback');
    const emailTemplateList = document.getElementById('email-template-list');
    const emailTemplateEmptyState = document.getElementById('email-template-empty');
    const emailTemplateNameInput = document.getElementById('email-template-name');
    const emailTemplateSubjectInput = document.getElementById('email-template-subject');
    const emailTemplateBlocksEmptyText = emailTemplateBlocksContainer
      ? (emailTemplateBlocksContainer.textContent || '').trim() ||
        'Ajoutez un bloc pour commencer la composition du message.'
      : 'Ajoutez un bloc pour commencer la composition du message.';
    const campaignSavedSearchSelect = document.getElementById('campaign-saved-search-select');
    const campaignTemplateSelect = document.getElementById('campaign-template-select');
    const emailCampaignForm = document.getElementById('email-campaign-form');
    const emailCampaignFeedback = document.getElementById('email-campaign-feedback');
    const campaignSummary = document.getElementById('campaign-summary');
    const campaignRecipientList = document.getElementById('campaign-recipient-list');
    const campaignEmailPreview = document.getElementById('campaign-email-preview');
    const campaignHistoryList = document.getElementById('campaign-history-list');
    const campaignHistoryEmpty = document.getElementById('campaign-history-empty');
    const emailCampaignSubjectInput = document.getElementById('email-campaign-subject');
    const campaignSummaryDefaultText = campaignSummary
      ? campaignSummary.textContent || ''
      : 'Sélectionnez une recherche sauvegardée pour afficher les contacts concernés.';
    const campaignEmailPreviewEmptyHtml = campaignEmailPreview
      ? campaignEmailPreview.innerHTML
      : '<p class="empty-state">Choisissez un modèle pour afficher un aperçu.</p>';
    const contactSelectAllButton = document.getElementById('contact-select-all');
    const contactDeleteSelectedButton = document.getElementById('contact-delete-selected');
    const contactBulkKeywordSelect = document.getElementById('contact-bulk-keyword');
    const contactBulkKeywordButton = document.getElementById('contact-add-keyword');
    const contactSelectedCountEl = document.getElementById('contact-selected-count');
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
    const taskEmptyStateDefaultText = taskEmptyState
      ? taskEmptyState.textContent || ''
      : '';
    const taskCategoryForm = document.getElementById('task-category-form');
    const taskCategoryNameInput = document.getElementById('task-category-name');
    const taskCategoryColorInput = document.getElementById('task-category-color');
    const taskCategoryNav = document.getElementById('task-category-nav');
    const taskCategoryIndicator = document.getElementById('task-category-indicator');
    const taskCategoryList = document.getElementById('task-category-list');
    const taskCategoryEmpty = document.getElementById('task-category-empty');
    const taskCategorySelect = document.getElementById('task-category');
    const taskAttachmentInput = document.getElementById('task-attachment');
    const taskAttachmentPreview = document.getElementById('task-attachment-preview');
    const taskAttachmentPreviewName = document.getElementById('task-attachment-preview-name');
    const taskAttachmentPreviewMeta = document.getElementById('task-attachment-preview-meta');
    const taskAttachmentPreviewOpen = document.getElementById('task-attachment-preview-open');
    const taskAttachmentRemoveButton = document.getElementById('task-attachment-remove');
    const taskAttachmentRemovedMessage = document.getElementById('task-attachment-removed-message');
    const teamChatList = document.getElementById('team-chat-messages');
    const teamChatEmptyState = document.getElementById('team-chat-empty');
    const teamChatForm = document.getElementById('team-chat-form');
    const teamChatInput = document.getElementById('team-chat-input');

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
    const DEFAULT_TASK_CATEGORY_COLOR = '#4f46e5';
    const TASK_CATEGORY_FILTER_ALL = 'all';
    const TASK_CATEGORY_FILTER_UNCATEGORIZED = 'uncategorized';
    const MAX_TASK_ATTACHMENT_SIZE = 2 * 1024 * 1024;
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
    const newsDateFormatter = new Intl.DateTimeFormat('fr-FR', {
      dateStyle: 'medium',
      timeStyle: 'short',
    });

    let data = loadDataForUser(currentUser);
    data = upgradeDataStructure(data);
    let newsData = loadNewsStore();

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

    const isOwner = currentUser === data.panelOwner;

    if (!isOwner) {
      if (topbarTeamButton instanceof HTMLElement && topbarTeamButton.parentElement) {
        topbarTeamButton.parentElement.removeChild(topbarTeamButton);
      }
      delete MODULE_CONFIG.team;
      pageModuleMap.delete('team');
      const teamIndex = topbarButtons.indexOf(topbarTeamButton);
      if (teamIndex >= 0) {
        topbarButtons.splice(teamIndex, 1);
      }
    }

    if (!Array.isArray(data.teamMembers) || data.teamMembers.length === 0) {
      data.teamMembers = [currentUser];
      saveDataForUser(currentUser, data);
    }

    const numberFormatter = new Intl.NumberFormat('fr-FR');
    const percentFormatter = new Intl.NumberFormat('fr-FR', {
      style: 'percent',
      maximumFractionDigits: 1,
    });
    const fileSizeFormatter = new Intl.NumberFormat('fr-FR', {
      maximumFractionDigits: 1,
      minimumFractionDigits: 0,
    });
    const savedSearchDateFormatter = new Intl.DateTimeFormat('fr-FR', {
      dateStyle: 'medium',
      timeStyle: 'short',
    });
    const campaignDateFormatter = new Intl.DateTimeFormat('fr-FR', {
      dateStyle: 'medium',
      timeStyle: 'short',
    });

    const CONTACT_RESULTS_PER_PAGE_DEFAULT = 10;
    const KEYWORD_FILTER_MODE_ALL = 'all';
    const KEYWORD_FILTER_MODE_ANY = 'any';

    let contactSearchTerm = '';
    let advancedFilters = createEmptyAdvancedFilters();
    let activeSavedSearchId = '';
    let contactEditReturnPage = 'contacts-search';
    let contactCurrentPage = 1;
    let contactResultsPerPage = CONTACT_RESULTS_PER_PAGE_DEFAULT;
    let selectedContactIds = new Set();
    let lastContactSearchResultIds = [];
    let categoryDragAndDropInitialized = false;
    let calendarViewMode = 'month';
    let calendarReferenceDate = startOfMonth(new Date());
    let calendarSelectedDate = startOfDay(new Date());
    let calendarLastFocusedElement = null;
    let calendarHasBeenOpened = false;
    let teamMembers = loadTeamMembers();
    let teamMembersById = new Map(teamMembers.map((member) => [member.username, member]));
    let currentModuleId = '';
    let currentPageId = '';
    let taskCategoriesById = new Map();
    let taskCategoryFilter = TASK_CATEGORY_FILTER_ALL;
    let taskStatusFilter = 'active';
    let editingTaskExistingAttachment = null;
    let removeAttachmentOnSubmit = false;
    let taskCategoryIndicatorFrame = 0;
    let emailTemplateDraftBlocks = [];
    let emailTemplateEditingId = '';
    let campaignRecipients = [];
    let campaignSubjectEdited = false;
    let campaignActiveTemplateId = '';
    let campaignSubjectTemplateId = '';
    let keywordStatsActiveId = '';
    let keywordStatsPreviousFocus = null;

    renderHomeTeams();
    renderHomeSubscription();
    renderHomeDirectoryLimit();
    renderHomeNews();
    if (isSuperAdmin) {
      renderAdminNewsList();
      renderAdminUserOptions();
    }

    normalizeCategoryOrders();
    populateTaskMemberOptions();
    updateTaskPanelDescription();
    renderTaskCategories();
    resetTaskCategoryFormDefaults();
    resetTaskFormDefaults();
    renderTasks();
    renderTeamChat();

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

    if (homeTeamForm instanceof HTMLFormElement) {
      homeTeamForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const nameValue = homeTeamNameInput ? homeTeamNameInput.value : '';
        const roleValue = homeTeamRoleInput ? homeTeamRoleInput.value : '';
        const added = addTeamMembership(nameValue, roleValue);
        if (added) {
          homeTeamForm.reset();
          if (homeTeamNameInput instanceof HTMLInputElement) {
            homeTeamNameInput.focus();
          }
        }
      });
    }

    if (homeTeamsList) {
      homeTeamsList.addEventListener('click', (event) => {
        const target =
          event.target instanceof HTMLElement
            ? event.target.closest('button[data-action]')
            : null;
        if (!(target instanceof HTMLButtonElement)) {
          return;
        }

        const action = target.dataset.action || '';
        const teamId = target.dataset.teamId || '';
        if (!teamId) {
          return;
        }

        if (action === 'home-team-set-current') {
          setCurrentTeam(teamId);
        } else if (action === 'home-team-remove') {
          removeTeamMembership(teamId);
        }
      });
    }

    if (homeSubscriptionSelect instanceof HTMLSelectElement) {
      homeSubscriptionSelect.addEventListener('change', (event) => {
        const target = event.target instanceof HTMLSelectElement ? event.target : homeSubscriptionSelect;
        const selectedPlan = getSubscriptionPlan(target.value);
        if (!selectedPlan || selectedPlan.id === data.subscription) {
          return;
        }
        data.subscription = selectedPlan.id;
        data.lastUpdated = new Date().toISOString();
        saveDataForUser(currentUser, data);
        renderHomeSubscription();
        renderHomeDirectoryLimit();
      });
    }

    if (adminArticleForm instanceof HTMLFormElement && isSuperAdmin) {
      adminArticleForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const formData = new FormData(adminArticleForm);
        const title = (formData.get('title') || '').toString().trim();
        const summary = (formData.get('summary') || '').toString().trim();
        const content = (formData.get('content') || '').toString().trim();

        if (!title || !summary) {
          setAdminArticleFeedback('Veuillez renseigner le titre et le résumé.', 'error');
          return;
        }

        addNewsArticle({
          id: generateId('news'),
          title,
          summary,
          content,
          createdAt: new Date().toISOString(),
          author: currentUser,
        });

        adminArticleForm.reset();
        setAdminArticleFeedback('La nouveauté a été publiée.', 'success');
      });
    }

    if (adminNewsList && isSuperAdmin) {
      adminNewsList.addEventListener('click', (event) => {
        const target =
          event.target instanceof HTMLElement
            ? event.target.closest('button[data-action="delete-news"]')
            : null;
        if (!(target instanceof HTMLButtonElement)) {
          return;
        }

        const articleId = target.dataset.newsId || '';
        if (!articleId) {
          return;
        }

        const removed = deleteNewsArticle(articleId);
        if (removed) {
          setAdminArticleFeedback('La nouveauté a été supprimée.', 'success');
        } else {
          setAdminArticleFeedback("Impossible de supprimer cette nouveauté.", 'error');
        }
      });
    }

    if (adminImpersonateForm instanceof HTMLFormElement && isSuperAdmin) {
      adminImpersonateForm.addEventListener('submit', (event) => {
        event.preventDefault();
        if (!(adminImpersonateSelect instanceof HTMLSelectElement)) {
          return;
        }
        userStore = loadUserStore();
        const selectedUsername = adminImpersonateSelect.value;
        if (!selectedUsername) {
          setAdminImpersonateFeedback('Sélectionnez un utilisateur.', 'error');
          return;
        }
        if (selectedUsername === currentUser) {
          setAdminImpersonateFeedback('Vous êtes déjà connecté avec ce compte.', 'error');
          return;
        }
        const usersObj =
          userStore && userStore.users && typeof userStore.users === 'object'
            ? userStore.users
            : {};
        if (!usersObj[selectedUsername]) {
          setAdminImpersonateFeedback('Utilisateur introuvable.', 'error');
          return;
        }

        saveActiveUser(selectedUsername);
        setAdminImpersonateFeedback(`Connexion en cours en tant que ${selectedUsername}…`, 'success');
        window.location.reload();
      });
    }

    document.addEventListener('umanager:data-changed', (event) => {
      const detail = event && typeof event === 'object' ? event.detail : null;
      if (detail && detail.section === 'contacts') {
        renderHomeDirectoryLimit();
      }
    });

    document.addEventListener('umanager:page-changed', (event) => {
      const detail = event && typeof event === 'object' ? event.detail : null;
      if (detail && detail.pageId === 'administration' && isSuperAdmin) {
        renderAdminUserOptions();
      }
    });

    topbarButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const moduleId = button.dataset.topTarget || '';
        if (!moduleId) {
          return;
        }
        activateModule(moduleId);
      });
    });

    if (contextNavList) {
      contextNavList.addEventListener('click', (event) => {
        const target =
          event.target instanceof HTMLElement
            ? event.target.closest('.context-button')
            : null;
        if (!(target instanceof HTMLButtonElement)) {
          return;
        }
        const pageId = target.dataset.contextTarget || '';
        if (pageId) {
          activatePage(pageId);
        }
      });
    }

    summaryNavigateButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const targetPage = button.dataset.navigate || '';
        if (targetPage) {
          showPage(targetPage);
        }
      });
    });

    taskStatusButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const status = button.dataset.taskStatus || '';
        setTaskStatusFilter(status);
      });
    });

    updateTaskStatusButtons();

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

    if (taskCategoryForm) {
      taskCategoryForm.addEventListener('submit', (event) => {
        event.preventDefault();
        createTaskCategoryFromForm();
      });

      taskCategoryForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          resetTaskCategoryFormDefaults();
          if (taskCategoryNameInput instanceof HTMLInputElement) {
            taskCategoryNameInput.focus();
          }
        });
      });
    }

    if (taskCategoryNav) {
      taskCategoryNav.addEventListener('click', (event) => {
        const button =
          event.target instanceof HTMLElement
            ? event.target.closest('[data-task-category-filter]')
            : null;
        if (button instanceof HTMLButtonElement) {
          const filterId = button.dataset.taskCategoryFilter || TASK_CATEGORY_FILTER_ALL;
          setTaskCategoryFilter(filterId);
        }
      });

      taskCategoryNav.addEventListener('scroll', () => {
        scheduleTaskCategoryIndicatorUpdate();
      });
    }

    window.addEventListener('resize', () => {
      scheduleTaskCategoryIndicatorUpdate();
    });

    if (taskCategoryList) {
      taskCategoryList.addEventListener('click', (event) => {
        const button =
          event.target instanceof HTMLElement
            ? event.target.closest('[data-task-category-action]')
            : null;
        if (!button) {
          return;
        }

        if (button.dataset.taskCategoryAction === 'delete') {
          const categoryId = button.dataset.taskCategoryId || '';
          if (categoryId) {
            deleteTaskCategory(categoryId);
          }
        }
      });
    }

    if (taskAttachmentInput) {
      taskAttachmentInput.addEventListener('change', () => {
        handleTaskAttachmentChange();
      });
    }

    if (taskAttachmentRemoveButton) {
      taskAttachmentRemoveButton.addEventListener('click', () => {
        handleTaskAttachmentRemove();
      });
    }

    if (taskMemberSelect instanceof HTMLSelectElement) {
      taskMemberSelect.addEventListener('change', () => {
        if (taskMemberSelect.disabled) {
          return;
        }

        const options = Array.from(taskMemberSelect.options || []);
        const noneOption = options.find((option) => option.value === TASK_MEMBER_NONE_VALUE);
        if (!noneOption) {
          return;
        }

        const hasAssignments = options.some(
          (option) => option.value !== TASK_MEMBER_NONE_VALUE && option.selected,
        );

        if (hasAssignments) {
          noneOption.selected = false;
        } else {
          noneOption.selected = true;
        }
      });
    }

    if (taskForm) {
      taskForm.addEventListener('submit', async (event) => {
        event.preventDefault();
        try {
          if (editingTaskId) {
            await applyTaskEditsFromForm();
          } else {
            await createTaskFromForm();
          }
        } catch (error) {
          console.error('Erreur lors de la sauvegarde de la tâche :', error);
        }
      });

      taskForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          editingTaskId = '';
          setTaskFormMode('create');
          resetTaskFormDefaults();
          if (taskTitleInput instanceof HTMLInputElement) {
            taskTitleInput.focus();
          }
        });
      });
    }

    if (teamChatForm instanceof HTMLFormElement) {
      teamChatForm.addEventListener('submit', (event) => {
        event.preventDefault();

        if (!(teamChatInput instanceof HTMLTextAreaElement)) {
          return;
        }

        const rawContent = teamChatInput.value.trim();
        if (!rawContent) {
          teamChatInput.focus();
          return;
        }

        const limitedContent = rawContent.slice(0, TEAM_CHAT_MAX_MESSAGE_LENGTH);
        const message = normalizeTeamChatMessage({
          id: generateId('chat'),
          author: currentUser,
          content: limitedContent,
          createdAt: new Date().toISOString(),
        });

        if (!message) {
          return;
        }

        if (!Array.isArray(data.teamChatMessages)) {
          data.teamChatMessages = [];
        }

        data.teamChatMessages.push(message);
        if (data.teamChatMessages.length > TEAM_CHAT_HISTORY_LIMIT) {
          data.teamChatMessages = data.teamChatMessages.slice(-TEAM_CHAT_HISTORY_LIMIT);
        }

        data.lastUpdated = new Date().toISOString();
        saveDataForUser(currentUser, data);
        renderTeamChat();
        notifyDataChanged('team-chat', { type: 'message-created', messageId: message.id });

        teamChatForm.reset();
        window.requestAnimationFrame(() => {
          if (teamChatInput instanceof HTMLTextAreaElement) {
            teamChatInput.focus();
          }
        });
      });
    }


    if (taskList) {
      taskList.addEventListener('click', (event) => {
        const target =
          event.target instanceof HTMLElement
            ? event.target.closest('[data-action]')
            : null;
        if (!(target instanceof HTMLElement)) {
          return;
        }

        const action = target.dataset.action || '';
        const taskId = target.dataset.taskId || '';

        switch (action) {
          case 'delete-task':
            if (taskId) {
              deleteTask(taskId);
            }
            break;
          case 'edit-task':
            if (taskId) {
              startEditTask(taskId);
            }
            break;
          case 'archive-task':
            if (taskId) {
              archiveTask(taskId);
            }
            break;
          case 'restore-task':
            if (taskId) {
              restoreTask(taskId);
            }
            break;
          default:
            break;
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

    keywordStatsCloseButtons.forEach((button) => {
      button.addEventListener('click', () => {
        closeKeywordStats();
      });
    });

    if (keywordStatsOverlay) {
      keywordStatsOverlay.addEventListener('click', (event) => {
        if (event.target === keywordStatsOverlay) {
          closeKeywordStats();
        }
      });
    }

    document.addEventListener('keydown', handleKeywordStatsKeydown);
    
    if (emailTemplateAddParagraphButton instanceof HTMLButtonElement) {
      emailTemplateAddParagraphButton.addEventListener('click', () => {
        addEmailTemplateBlock('paragraph');
      });
    }

    if (emailTemplateAddImageButton instanceof HTMLButtonElement) {
      emailTemplateAddImageButton.addEventListener('click', () => {
        addEmailTemplateBlock('image');
      });
    }

    if (emailTemplateAddButtonButton instanceof HTMLButtonElement) {
      emailTemplateAddButtonButton.addEventListener('click', () => {
        addEmailTemplateBlock('button');
      });
    }

    if (emailTemplateBlocksContainer) {
      emailTemplateBlocksContainer.addEventListener('input', handleEmailTemplateBlockInput);
      emailTemplateBlocksContainer.addEventListener('change', handleEmailTemplateBlockInput);
      emailTemplateBlocksContainer.addEventListener('click', handleEmailTemplateBlockAction);
    }

    if (emailTemplateForm) {
      emailTemplateForm.addEventListener('submit', (event) => {
        event.preventDefault();
        if (!emailTemplateNameInput || !emailTemplateSubjectInput) {
          return;
        }

        const name = emailTemplateNameInput.value.trim();
        const subject = emailTemplateSubjectInput.value.trim();

        if (!name) {
          setEmailTemplateFeedback('Veuillez renseigner un nom pour le modèle.', 'error');
          emailTemplateNameInput.focus();
          return;
        }

        if (!subject) {
          setEmailTemplateFeedback("Veuillez renseigner l'objet du mail.", 'error');
          emailTemplateSubjectInput.focus();
          return;
        }

        if (!Array.isArray(emailTemplateDraftBlocks) || emailTemplateDraftBlocks.length === 0) {
          setEmailTemplateFeedback('Ajoutez au moins un bloc de contenu au modèle.', 'error');
          return;
        }

        for (const block of emailTemplateDraftBlocks) {
          if (block.type === 'image' && !block.data.url) {
            setEmailTemplateFeedback("Chaque bloc image doit comporter l'URL de l'image.", 'error');
            focusEmailTemplateBlock(block.id, "input[data-block-field='url']");
            return;
          }
          if (block.type === 'button') {
            if (!block.data.label) {
              setEmailTemplateFeedback('Chaque bouton doit avoir un libellé.', 'error');
              focusEmailTemplateBlock(block.id, "input[data-block-field='label']");
              return;
            }
            if (!block.data.url) {
              setEmailTemplateFeedback('Chaque bouton doit pointer vers une adresse web.', 'error');
              focusEmailTemplateBlock(block.id, "input[data-block-field='url']");
              return;
            }
          }
        }

        const normalizedBlocks = emailTemplateDraftBlocks.map((block) =>
          cloneEmailTemplateBlock(block, { preserveId: true }),
        );

        const now = new Date().toISOString();
        if (!Array.isArray(data.emailTemplates)) {
          data.emailTemplates = [];
        }

        let templateRecord = null;
        let operation = 'created';

        if (emailTemplateEditingId) {
          templateRecord = data.emailTemplates.find((item) => item && item.id === emailTemplateEditingId) || null;
          if (templateRecord) {
            templateRecord.name = name;
            templateRecord.subject = subject;
            templateRecord.blocks = normalizedBlocks;
            templateRecord.updatedAt = now;
            if (!templateRecord.createdAt) {
              templateRecord.createdAt = now;
            }
            operation = 'updated';
          }
        }

        if (!templateRecord) {
          templateRecord = {
            id: generateId('email-template'),
            name,
            subject,
            blocks: normalizedBlocks,
            createdAt: now,
            updatedAt: now,
          };
          data.emailTemplates.push(templateRecord);
        }

        emailTemplateEditingId = '';
        data.lastUpdated = now;
        saveDataForUser(currentUser, data);
        renderEmailTemplateList();
        renderCampaignTemplateOptions();

        const feedbackMessage =
          operation === 'updated'
            ? `Modèle « ${name} » mis à jour.`
            : `Modèle « ${name} » enregistré.`;
        notifyDataChanged('email-templates', {
          type: operation,
          templateId: templateRecord.id,
        });
        resetEmailTemplateForm(true, true);
        setEmailTemplateFeedback(feedbackMessage, 'success');
      });

      emailTemplateForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          resetEmailTemplateForm();
        });
      });
    }

    if (emailTemplateList) {
      emailTemplateList.addEventListener('click', (event) => {
        const button = event.target instanceof HTMLElement ? event.target.closest('button[data-action]') : null;
        if (!(button instanceof HTMLButtonElement)) {
          return;
        }
        const templateId = button.dataset.templateId || '';
        if (!templateId) {
          return;
        }
        if (button.dataset.action === 'edit-template') {
          loadEmailTemplateForEditing(templateId);
        } else if (button.dataset.action === 'duplicate-template') {
          duplicateEmailTemplate(templateId);
        } else if (button.dataset.action === 'delete-template') {
          deleteEmailTemplate(templateId);
        }
      });
    }

    if (campaignTemplateSelect instanceof HTMLSelectElement) {
      campaignTemplateSelect.addEventListener('change', () => {
        campaignSubjectEdited = false;
        updateCampaignTemplatePreview();
      });
    }

    if (campaignSavedSearchSelect instanceof HTMLSelectElement) {
      campaignSavedSearchSelect.addEventListener('change', () => {
        updateCampaignRecipients();
      });
    }

    if (emailCampaignSubjectInput instanceof HTMLInputElement) {
      emailCampaignSubjectInput.addEventListener('input', () => {
        const currentValue = emailCampaignSubjectInput.value.trim();
        if (!campaignSubjectTemplateId) {
          campaignSubjectEdited = currentValue.length > 0;
          return;
        }
        const template = getTemplateById(campaignSubjectTemplateId);
        const templateSubject = template && template.subject ? template.subject.trim() : '';
        campaignSubjectEdited = currentValue !== templateSubject;
      });
    }

    let campaignResetKeepFeedback = false;

    if (emailCampaignForm) {
      emailCampaignForm.addEventListener('submit', (event) => {
        event.preventDefault();
        if (!(campaignTemplateSelect instanceof HTMLSelectElement)) {
          return;
        }
        if (!(campaignSavedSearchSelect instanceof HTMLSelectElement)) {
          return;
        }

        const templateId = campaignTemplateSelect.value;
        if (!templateId) {
          setEmailCampaignFeedback('Sélectionnez un modèle de mail à envoyer.', 'error');
          campaignTemplateSelect.focus();
          return;
        }
        const template = getTemplateById(templateId);
        if (!template) {
          setEmailCampaignFeedback('Le modèle choisi est introuvable.', 'error');
          return;
        }

        const savedSearchId = campaignSavedSearchSelect.value;
        if (!savedSearchId) {
          setEmailCampaignFeedback('Sélectionnez une recherche sauvegardée.', 'error');
          campaignSavedSearchSelect.focus();
          return;
        }

        const { recipients, savedSearch } = getRecipientsForSavedSearch(savedSearchId);
        if (recipients.length === 0) {
          setEmailCampaignFeedback(
            "Aucun contact avec adresse mail n'est disponible dans cette recherche.",
            'error',
          );
          return;
        }

        const formData = new FormData(emailCampaignForm);
        const name = (formData.get('email-campaign-name') || '').toString().trim();
        if (!name) {
          setEmailCampaignFeedback('Donnez un nom à la campagne pour la retrouver facilement.', 'error');
          const nameInput = emailCampaignForm.querySelector('#email-campaign-name');
          if (nameInput instanceof HTMLInputElement) {
            nameInput.focus();
          }
          return;
        }

        const senderEmail = (formData.get('email-campaign-sender-email') || '').toString().trim();
        if (!isValidEmail(senderEmail)) {
          setEmailCampaignFeedback("Renseignez une adresse mail d'expéditeur valide.", 'error');
          const senderInput = emailCampaignForm.querySelector('#email-campaign-sender-email');
          if (senderInput instanceof HTMLInputElement) {
            senderInput.focus();
          }
          return;
        }

        const senderName = (formData.get('email-campaign-sender-name') || '').toString().trim();
        const subject = (formData.get('email-campaign-subject') || '').toString().trim();
        if (!subject) {
          setEmailCampaignFeedback("L'objet du mail est obligatoire.", 'error');
          if (emailCampaignSubjectInput) {
            emailCampaignSubjectInput.focus();
          }
          return;
        }

        const now = new Date().toISOString();
        if (!Array.isArray(data.emailCampaigns)) {
          data.emailCampaigns = [];
        }

        const campaignRecord = {
          id: generateId('email-campaign'),
          name,
          templateId: template.id || templateId,
          templateName: template.name || 'Modèle sans titre',
          savedSearchId,
          savedSearchName: savedSearch ? savedSearch.name : '',
          senderName,
          senderEmail,
          subject,
          sentAt: now,
          recipientCount: recipients.length,
          recipients: recipients.map((recipient) => ({ ...recipient })),
          templateSnapshot: Array.isArray(template.blocks)
            ? template.blocks.map((block) => cloneEmailTemplateBlock(block, { preserveId: true }))
            : [],
        };

        data.emailCampaigns.push(campaignRecord);
        data.lastUpdated = now;
        saveDataForUser(currentUser, data);
        renderCampaignHistory();
        const sendStatus = dispatchCampaignSend(campaignRecord, template, {
          senderName,
          senderEmail,
        });
        if (sendStatus === 'skipped') {
          setEmailCampaignFeedback(`Campagne « ${name} » enregistrée.`, 'success');
        }
        notifyDataChanged('email-campaigns', {
          type: 'created',
          campaignId: campaignRecord.id,
        });
        campaignResetKeepFeedback = true;
        if (emailCampaignForm) {
          emailCampaignForm.reset();
        }
      });

      emailCampaignForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          resetCampaignForm(true, { keepFeedback: campaignResetKeepFeedback });
          campaignResetKeepFeedback = false;
        });
      });
    }

    if (contactForm) {
      contactForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const formData = new FormData(contactForm);
        const contactType = normalizeContactType(
          (formData.get('contact-type') || '').toString(),
        );
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

        const identifierResult = ensureContactIdentifier(categoryValues);
        if (!identifierResult) {
          const birthDateInput = contactForm.querySelector(
            `[data-category-id="${REQUIRED_CATEGORY_IDS.birthDate}"]`,
          );
          if (
            birthDateInput &&
            'reportValidity' in birthDateInput &&
            typeof birthDateInput.reportValidity === 'function'
          ) {
            birthDateInput.reportValidity();
          }
          if (birthDateInput && 'focus' in birthDateInput && typeof birthDateInput.focus === 'function') {
            birthDateInput.focus();
          }
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
          contactToUpdate.type = contactType;
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
            type: contactType,
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
        setActiveSavedSearchId('');
        clearContactSaveSearchFeedback();
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
        const rawKeywordMode = formData.get('search-keyword-mode') || KEYWORD_FILTER_MODE_ALL;
        const keywordMode =
          rawKeywordMode === KEYWORD_FILTER_MODE_ANY
            ? KEYWORD_FILTER_MODE_ANY
            : KEYWORD_FILTER_MODE_ALL;

        advancedFilters = {
          categories: categoryFilters,
          keywords: keywordFilters,
          keywordMode,
        };
        setActiveSavedSearchId('');
        clearContactSaveSearchFeedback();
        contactCurrentPage = 1;
        renderContacts();
      });

      contactAdvancedSearchForm.addEventListener('reset', () => {
        window.requestAnimationFrame(() => {
          advancedFilters = createEmptyAdvancedFilters();
          renderSearchCategoryFields();
          renderSearchKeywordOptions();
          setActiveSavedSearchId('');
          clearContactSaveSearchFeedback();
          contactCurrentPage = 1;
          renderContacts();
        });
      });
    }

    if (contactSaveSearchButton instanceof HTMLButtonElement) {
      contactSaveSearchButton.addEventListener('click', () => {
        handleSaveCurrentSearch();
      });
    }

    if (contactSavedSearchSelect instanceof HTMLSelectElement) {
      contactSavedSearchSelect.addEventListener('change', () => {
        const selectedId = contactSavedSearchSelect.value || '';
        if (!selectedId) {
          setActiveSavedSearchId('');
          clearContactSaveSearchFeedback();
          return;
        }

        const applied = applySavedSearchById(selectedId, {
          navigateToSearchPage: false,
          focusSearchInput: false,
          announce: true,
        });
        if (!applied) {
          contactSavedSearchSelect.value = '';
        }
      });
    }

    if (savedSearchList) {
      savedSearchList.addEventListener('click', (event) => {
        const target =
          event.target instanceof HTMLElement ? event.target.closest('[data-action]') : null;
        if (!(target instanceof HTMLElement)) {
          return;
        }

        const savedSearchId = target.dataset.savedSearchId || '';
        if (!savedSearchId) {
          return;
        }

        if (target.dataset.action === 'apply-saved-search') {
          const applied = applySavedSearchById(savedSearchId, {
            navigateToSearchPage: true,
            focusSearchInput: true,
            announce: true,
          });
          if (!applied && contactSaveSearchFeedback) {
            contactSaveSearchFeedback.textContent =
              "Impossible d'appliquer cette recherche sauvegardée.";
          }
        } else if (target.dataset.action === 'delete-saved-search') {
          deleteSavedSearch(savedSearchId);
        }
      });
    }

    if (contactSelectAllButton instanceof HTMLButtonElement) {
      contactSelectAllButton.addEventListener('click', () => {
        selectedContactIds = new Set(lastContactSearchResultIds);
        updateSelectedContactsUI();
        refreshContactSelectionCheckboxes();
      });
    }

    if (contactDeleteSelectedButton instanceof HTMLButtonElement) {
      contactDeleteSelectedButton.addEventListener('click', () => {
        if (selectedContactIds.size === 0) {
          return;
        }
        const count = selectedContactIds.size;
        const confirmationMessage =
          count === 1
            ? 'Supprimer le contact sélectionné ?'
            : `Supprimer les ${numberFormatter.format(count)} contacts sélectionnés ?`;
        if (!window.confirm(confirmationMessage)) {
          return;
        }
        removeContactsByIds(Array.from(selectedContactIds));
      });
    }

    if (contactBulkKeywordSelect instanceof HTMLSelectElement) {
      contactBulkKeywordSelect.addEventListener('change', () => {
        updateSelectedContactsUI();
      });
    }

    if (contactBulkKeywordButton instanceof HTMLButtonElement) {
      contactBulkKeywordButton.addEventListener('click', () => {
        if (
          selectedContactIds.size === 0 ||
          !(contactBulkKeywordSelect instanceof HTMLSelectElement)
        ) {
          return;
        }
        const keywordId = contactBulkKeywordSelect.value;
        if (!keywordId) {
          return;
        }
        addKeywordToContacts(Array.from(selectedContactIds), keywordId);
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
        const autoMerge = Boolean(payload && payload.autoMerge);
        return importContactsFromRows(rows, { mapping, skipHeader, fileName, autoMerge });
      },
    };

    window.UManager = window.UManager || {};
    window.UManager.importApi = importApi;

    activateModule('home', 'home-overview');
    renderMetrics();
    renderCategories();
    renderKeywords();
    renderEmailTemplateBlocks();
    renderEmailTemplateList();
    renderCampaignTemplateOptions();
    renderEmailTemplatePreview();
    renderCampaignHistory();
    updateCampaignRecipients();

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
      const moduleToRestore =
        currentModuleId && currentModuleId !== 'calendar' ? currentModuleId : 'tasks';
      renderContextNavigationForModule(moduleToRestore);
      setActiveTopbar(moduleToRestore);
      if (calendarLastFocusedElement && typeof calendarLastFocusedElement.focus === 'function') {
        calendarLastFocusedElement.focus();
      }
      calendarLastFocusedElement = null;
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

    function setActiveTopbar(moduleId) {
      topbarButtons.forEach((button) => {
        const isActive = button.dataset.topTarget === moduleId;
        button.classList.toggle('active', isActive);
        if (isActive) {
          button.setAttribute('aria-current', 'page');
        } else {
          button.removeAttribute('aria-current');
        }
      });
    }

    function renderContextNavigationForModule(moduleId) {
      if (!contextNavList) {
        return;
      }

      contextNavList.innerHTML = '';
      const entries = MODULE_CONFIG[moduleId] || [];

      if (entries.length === 0) {
        if (contextNav) {
          contextNav.hidden = true;
          if (!contextNav.hasAttribute('hidden')) {
            contextNav.setAttribute('hidden', '');
          }
        }
        if (contextEmptyState) {
          contextEmptyState.hidden = false;
          contextEmptyState.removeAttribute('hidden');
          contextEmptyState.textContent =
            contextEmptyDefaultText || 'Aucune option disponible pour ce module.';
        }
        return;
      }

      if (contextNav) {
        contextNav.hidden = false;
        contextNav.removeAttribute('hidden');
      }
      if (contextEmptyState) {
        contextEmptyState.hidden = true;
        if (!contextEmptyState.hasAttribute('hidden')) {
          contextEmptyState.setAttribute('hidden', '');
        }
        contextEmptyState.textContent = contextEmptyDefaultText;
      }

      const fragment = document.createDocumentFragment();
      entries.forEach((entry) => {
        if (!entry || !entry.id) {
          return;
        }
        const listItem = document.createElement('li');
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'context-button';
        button.dataset.contextTarget = entry.id;
        button.textContent = entry.label;
        listItem.appendChild(button);
        fragment.appendChild(listItem);
      });
      contextNavList.appendChild(fragment);
      setActiveContextButton(currentPageId);
    }

    function setActiveContextButton(pageId) {
      if (!contextNavList) {
        return;
      }
      const buttons = Array.from(contextNavList.querySelectorAll('.context-button'));
      buttons.forEach((button) => {
        const isActive =
          button instanceof HTMLButtonElement && button.dataset.contextTarget === pageId;
        button.classList.toggle('active', isActive);
        if (isActive) {
          button.setAttribute('aria-current', 'page');
        } else {
          button.removeAttribute('aria-current');
        }
      });
    }

    function activateModule(moduleId, requestedPageId) {
      if (!moduleId) {
        return;
      }

      if (moduleId === 'calendar') {
        setActiveTopbar(moduleId);
        if (contextNav) {
          contextNav.hidden = true;
          if (!contextNav.hasAttribute('hidden')) {
            contextNav.setAttribute('hidden', '');
          }
        }
        if (contextEmptyState) {
          contextEmptyState.hidden = false;
          contextEmptyState.textContent = 'Le calendrier ne propose pas de menu dédié.';
        }
        openCalendar();
        return;
      }

      if (moduleId === 'administration' && !isSuperAdmin) {
        activateModule('home', 'home-overview');
        return;
      }

      if (moduleId === 'team' && !isOwner) {
        activateModule('home', 'home-overview');
        return;
      }

      closeCalendar();

      if (moduleId !== currentModuleId) {
        currentModuleId = moduleId;
        renderContextNavigationForModule(moduleId);
      } else if (!contextNavList || contextNavList.childElementCount === 0) {
        renderContextNavigationForModule(moduleId);
      }

      setActiveTopbar(moduleId);

      const entries = MODULE_CONFIG[moduleId] || [];
      const defaultPageId = entries.length > 0 ? entries[0].id : '';
      const nextPageId =
        requestedPageId && pageModuleMap.get(requestedPageId) === moduleId
          ? requestedPageId
          : defaultPageId;

      if (nextPageId) {
        activatePage(nextPageId);
        return;
      }

      pages.forEach((page) => page.classList.remove('active'));
      currentPageId = '';
      setActiveContextButton('');
    }

    function activatePage(pageId) {
      if (!pageId) {
        return;
      }

      if (pageId === 'team' && !isOwner) {
        activateModule('home', 'home-overview');
        return;
      }

      if (pageId === 'administration' && !isSuperAdmin) {
        activateModule('home', 'home-overview');
        return;
      }

      closeCalendar();

      const moduleId = pageModuleMap.get(pageId);
      if (moduleId && moduleId !== currentModuleId) {
        activateModule(moduleId, pageId);
        return;
      }

      let hasMatch = false;
      pages.forEach((page) => {
        const isActive = page.id === pageId;
        if (isActive) {
          hasMatch = true;
        }
        page.classList.toggle('active', isActive);
      });

      if (!hasMatch) {
        return;
      }

      currentPageId = pageId;
      setActiveContextButton(pageId);
      document.dispatchEvent(
        new CustomEvent('umanager:page-changed', {
          detail: { pageId },
        }),
      );
    }

    function showPage(target) {
      activatePage(target);
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

      const previouslySelected = taskMemberSelect.disabled
        ? []
        : Array.from(taskMemberSelect.selectedOptions || []).map((option) => option.value);

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

      const noneOption = document.createElement('option');
      noneOption.value = TASK_MEMBER_NONE_VALUE;
      noneOption.textContent = 'Personne';
      fragment.appendChild(noneOption);

      teamMembers.forEach((member) => {
        const option = document.createElement('option');
        option.value = member.username;
        option.textContent = formatTeamMemberLabel(member);
        fragment.appendChild(option);
      });

      taskMemberSelect.appendChild(fragment);

      const sanitizedSelection = Array.isArray(previouslySelected)
        ? previouslySelected.filter(
            (value) =>
              typeof value === 'string' &&
              value !== TASK_MEMBER_NONE_VALUE &&
              teamMembersById.has(value),
          )
        : [];

      setTaskMemberSelection(sanitizedSelection);
    }

    function setTaskMemberSelection(memberIds) {
      if (!(taskMemberSelect instanceof HTMLSelectElement) || taskMemberSelect.disabled) {
        return;
      }

      const normalizedIds = Array.isArray(memberIds)
        ? memberIds
            .map((value) => (typeof value === 'string' ? value.trim() : ''))
            .filter((value) => value && teamMembersById.has(value))
        : [];

      const hasAssignments = normalizedIds.length > 0;

      Array.from(taskMemberSelect.options).forEach((option) => {
        if (!(option instanceof HTMLOptionElement)) {
          return;
        }

        if (option.value === TASK_MEMBER_NONE_VALUE) {
          option.selected = !hasAssignments;
        } else {
          option.selected = normalizedIds.includes(option.value);
        }
      });
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
        renderTeamChat();
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
        setTaskMemberSelection([]);
      }

      if (taskDescriptionInput instanceof HTMLTextAreaElement) {
        taskDescriptionInput.value = '';
      }

      resetTaskAttachmentState();

      if (taskCategorySelect instanceof HTMLSelectElement) {
        const defaultCategoryId =
          taskCategoryFilter !== TASK_CATEGORY_FILTER_ALL &&
          taskCategoryFilter !== TASK_CATEGORY_FILTER_UNCATEGORIZED &&
          taskCategoriesById.has(taskCategoryFilter)
            ? taskCategoryFilter
            : '';
        setTaskFormCategory(defaultCategoryId);
      }
    }

    function setTaskFormCategory(categoryId) {
      if (!(taskCategorySelect instanceof HTMLSelectElement)) {
        return;
      }

      const sanitized =
        typeof categoryId === 'string' && taskCategoriesById.has(categoryId)
          ? categoryId
          : '';
      taskCategorySelect.value = sanitized;
    }

    function populateTaskCategorySelect(categories) {
      if (!(taskCategorySelect instanceof HTMLSelectElement)) {
        return;
      }

      const previousValue = taskCategorySelect.value;

      taskCategorySelect.innerHTML = '';

      const defaultOption = document.createElement('option');
      defaultOption.value = '';
      defaultOption.textContent = 'Sans catégorie';
      taskCategorySelect.appendChild(defaultOption);

      categories.forEach((category) => {
        if (!category || typeof category.id !== 'string' || !category.id) {
          return;
        }

        const option = document.createElement('option');
        option.value = category.id;
        option.textContent = category.name;
        taskCategorySelect.appendChild(option);
      });

      if (previousValue && taskCategoriesById.has(previousValue)) {
        taskCategorySelect.value = previousValue;
      } else {
        taskCategorySelect.value = '';
      }
    }

    function getSortedTaskCategories() {
      if (!Array.isArray(data.taskCategories)) {
        data.taskCategories = [];
      }

      return data.taskCategories
        .filter((item) => item && typeof item === 'object')
        .map((item) => normalizeTaskCategory(item))
        .sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }));
    }

    function renderTaskCategories() {
      const categories = getSortedTaskCategories();
      data.taskCategories = categories.slice();
      taskCategoriesById = new Map(categories.map((category) => [category.id, category]));

      populateTaskCategorySelect(categories);
      renderTaskCategoryNav(categories);
      renderTaskCategoryList(categories);

      if (
        taskCategoryFilter !== TASK_CATEGORY_FILTER_ALL &&
        taskCategoryFilter !== TASK_CATEGORY_FILTER_UNCATEGORIZED &&
        !taskCategoriesById.has(taskCategoryFilter)
      ) {
        taskCategoryFilter = TASK_CATEGORY_FILTER_ALL;
      }

      updateTaskCategoryNavState();
      scheduleTaskCategoryIndicatorUpdate();
    }

    function renderTaskCategoryNav(categories) {
      if (!taskCategoryNav) {
        return;
      }

      const indicator = taskCategoryIndicator;
      taskCategoryNav.innerHTML = '';

      const counts = buildTaskCounts(
        filterTasksByStatus(Array.isArray(data.tasks) ? data.tasks : []),
      );

      const navItems = [
        { id: TASK_CATEGORY_FILTER_ALL, label: 'Toutes', icon: '✨' },
        {
          id: TASK_CATEGORY_FILTER_UNCATEGORIZED,
          label: 'Sans catégorie',
          icon: '📂',
        },
        ...categories.map((category) => ({
          id: category.id,
          label: category.name,
          icon: '⬤',
          color: category.color,
        })),
      ];

      navItems.forEach((item) => {
        if (!item || !item.id) {
          return;
        }

        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'task-category-chip';
        button.dataset.taskCategoryFilter = item.id;
        button.setAttribute('aria-pressed', 'false');

        const icon = document.createElement('span');
        icon.className = 'task-category-chip-icon';
        icon.textContent = item.icon;
        if (item.color) {
          icon.style.color = item.color;
        }

        const label = document.createElement('span');
        label.className = 'task-category-chip-label';
        label.textContent = item.label;

        const count = document.createElement('span');
        count.className = 'task-category-chip-count';
        count.dataset.taskCategoryCount = item.id;
        count.textContent = formatTaskCountLabel(counts.get(item.id) || 0);

        button.appendChild(icon);
        button.appendChild(label);
        button.appendChild(count);

        taskCategoryNav.appendChild(button);
      });

      if (indicator) {
        taskCategoryNav.appendChild(indicator);
      }

      updateTaskCategoryNavState();
      updateTaskCategoryCountBadges(counts);
    }

    function renderTaskCategoryList(categories) {
      if (!taskCategoryList) {
        return;
      }

      taskCategoryList.innerHTML = '';

      if (!Array.isArray(categories) || categories.length === 0) {
        if (taskCategoryEmpty) {
          taskCategoryEmpty.hidden = false;
          taskCategoryEmpty.removeAttribute('hidden');
        }
        return;
      }

      if (taskCategoryEmpty) {
        taskCategoryEmpty.hidden = true;
        if (!taskCategoryEmpty.hasAttribute('hidden')) {
          taskCategoryEmpty.setAttribute('hidden', '');
        }
      }

      const counts = buildTaskCounts(Array.isArray(data.tasks) ? data.tasks : []);

      categories.forEach((category) => {
        if (!category || typeof category.id !== 'string') {
          return;
        }

        const item = document.createElement('li');
        item.className = 'task-category-item';
        item.style.setProperty('--task-category-color', category.color);

        const content = document.createElement('div');
        content.className = 'task-category-item-content';

        const dot = document.createElement('span');
        dot.className = 'task-category-item-dot';
        dot.setAttribute('aria-hidden', 'true');
        content.appendChild(dot);

        const nameSpan = document.createElement('span');
        nameSpan.textContent = category.name;
        content.appendChild(nameSpan);

        const actions = document.createElement('div');
        actions.className = 'task-category-item-actions';

        const countSpan = document.createElement('span');
        countSpan.className = 'task-category-item-count';
        countSpan.dataset.taskCategoryListCount = category.id;
        countSpan.textContent = formatTaskCountLabel(counts.get(category.id) || 0);

        const deleteButton = document.createElement('button');
        deleteButton.type = 'button';
        deleteButton.className = 'task-category-item-delete';
        deleteButton.dataset.taskCategoryAction = 'delete';
        deleteButton.dataset.taskCategoryId = category.id;
        deleteButton.textContent = 'Supprimer';

        actions.appendChild(countSpan);
        actions.appendChild(deleteButton);

        item.appendChild(content);
        item.appendChild(actions);

        taskCategoryList.appendChild(item);
      });

      updateTaskCategoryCountBadges(counts);
    }

    function updateTaskCategoryNavState() {
      if (!taskCategoryNav) {
        return;
      }

      const buttons = Array.from(
        taskCategoryNav.querySelectorAll('[data-task-category-filter]'),
      );
      buttons.forEach((button) => {
        if (!(button instanceof HTMLButtonElement)) {
          return;
        }

        const isActive = button.dataset.taskCategoryFilter === taskCategoryFilter;
        button.classList.toggle('active', isActive);
        button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
      });

      scheduleTaskCategoryIndicatorUpdate();
    }

    function updateTaskCategoryIndicatorPosition() {
      if (!taskCategoryNav || !taskCategoryIndicator) {
        return;
      }

      const activeButton = taskCategoryNav.querySelector(
        '[data-task-category-filter].active',
      );

      if (!(activeButton instanceof HTMLElement)) {
        taskCategoryIndicator.style.opacity = '0';
        return;
      }

      const left = activeButton.offsetLeft - taskCategoryNav.scrollLeft;
      const width = activeButton.offsetWidth;

      taskCategoryIndicator.style.transform = `translateX(${left}px)`;
      taskCategoryIndicator.style.width = `${width}px`;
      taskCategoryIndicator.style.opacity = '1';
    }

    function scheduleTaskCategoryIndicatorUpdate() {
      if (!taskCategoryIndicator) {
        return;
      }

      if (taskCategoryIndicatorFrame) {
        window.cancelAnimationFrame(taskCategoryIndicatorFrame);
      }

      taskCategoryIndicatorFrame = window.requestAnimationFrame(() => {
        taskCategoryIndicatorFrame = 0;
        updateTaskCategoryIndicatorPosition();
      });
    }

    function setTaskStatusFilter(rawValue) {
      const normalized =
        typeof rawValue === 'string' ? rawValue.trim().toLowerCase() : '';
      const nextStatus = TASK_STATUS_VALUES.has(normalized) ? normalized : 'active';
      if (nextStatus === taskStatusFilter) {
        return;
      }

      taskStatusFilter = nextStatus;
      updateTaskStatusButtons();
      renderTasks();
    }

    function updateTaskStatusButtons() {
      taskStatusButtons.forEach((button) => {
        const status = button.dataset.taskStatus || '';
        const isActive = status === taskStatusFilter;
        button.classList.toggle('active', isActive);
        button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
      });
    }

    function setTaskCategoryFilter(rawValue) {
      let value = typeof rawValue === 'string' ? rawValue : TASK_CATEGORY_FILTER_ALL;

      if (
        value !== TASK_CATEGORY_FILTER_ALL &&
        value !== TASK_CATEGORY_FILTER_UNCATEGORIZED &&
        !taskCategoriesById.has(value)
      ) {
        value = TASK_CATEGORY_FILTER_ALL;
      }

      if (taskCategoryFilter === value) {
        updateTaskCategoryNavState();
        renderTasks();
        return;
      }

      taskCategoryFilter = value;
      updateTaskCategoryNavState();
      renderTasks();

      if (
        !editingTaskId &&
        value !== TASK_CATEGORY_FILTER_ALL &&
        value !== TASK_CATEGORY_FILTER_UNCATEGORIZED
      ) {
        setTaskFormCategory(value);
      }
    }

    function buildTaskCounts(tasks) {
      const counts = new Map();
      const source = Array.isArray(tasks) ? tasks : [];

      counts.set(TASK_CATEGORY_FILTER_ALL, source.length);

      let uncategorized = 0;

      source.forEach((task) => {
        if (!task || typeof task !== 'object') {
          return;
        }

        const categoryId =
          typeof task.categoryId === 'string' && task.categoryId ? task.categoryId : '';
        if (!categoryId) {
          uncategorized += 1;
          return;
        }

        const current = counts.get(categoryId) || 0;
        counts.set(categoryId, current + 1);
      });

      counts.set(TASK_CATEGORY_FILTER_UNCATEGORIZED, uncategorized);

      return counts;
    }

    function updateTaskCategoryCountBadges(counts) {
      if (taskCategoryNav) {
        const navCountElements = taskCategoryNav.querySelectorAll('[data-task-category-count]');
        navCountElements.forEach((element) => {
          const key = element.dataset.taskCategoryCount || '';
          element.textContent = formatTaskCountLabel(counts.get(key) || 0);
        });
      }

      if (taskCategoryList) {
        const listCountElements = taskCategoryList.querySelectorAll(
          '[data-task-category-list-count]',
        );
        listCountElements.forEach((element) => {
          const key = element.dataset.taskCategoryListCount || '';
          element.textContent = formatTaskCountLabel(counts.get(key) || 0);
        });
      }
    }

    function formatTaskCountLabel(count) {
      const absolute = Number.isFinite(count) ? Math.max(0, Math.trunc(count)) : 0;
      const label = absolute > 1 ? 'tâches' : 'tâche';
      return `${numberFormatter.format(absolute)} ${label}`;
    }

    function resetTaskCategoryFormDefaults() {
      if (taskCategoryColorInput instanceof HTMLInputElement) {
        taskCategoryColorInput.value = DEFAULT_TASK_CATEGORY_COLOR;
      }

      if (taskCategoryNameInput instanceof HTMLInputElement) {
        taskCategoryNameInput.setCustomValidity('');
      }
    }

    function createTaskCategoryFromForm() {
      if (!taskCategoryForm) {
        return;
      }

      const formData = new FormData(taskCategoryForm);
      const name = (formData.get('task-category-name') || '').toString().trim();
      const colorValue = normalizeTaskCategoryColor(
        (formData.get('task-category-color') || DEFAULT_TASK_CATEGORY_COLOR).toString(),
      );

      if (taskCategoryNameInput instanceof HTMLInputElement) {
        taskCategoryNameInput.setCustomValidity('');
      }

      if (!name) {
        if (taskCategoryNameInput instanceof HTMLInputElement) {
          taskCategoryNameInput.focus();
        }
        return;
      }

      const normalizedName = name.toLowerCase();
      const duplicate = Array.from(taskCategoriesById.values()).some(
        (category) => category.name.toLowerCase() === normalizedName,
      );

      if (duplicate) {
        if (taskCategoryNameInput instanceof HTMLInputElement) {
          taskCategoryNameInput.setCustomValidity('Une catégorie porte déjà ce nom.');
          taskCategoryNameInput.reportValidity();
          taskCategoryNameInput.focus();
        }
        return;
      }

      const newCategory = normalizeTaskCategory({
        id: generateId('task-category'),
        name,
        color: colorValue,
        createdAt: new Date().toISOString(),
      });

      if (!Array.isArray(data.taskCategories)) {
        data.taskCategories = [];
      }

      data.taskCategories.push(newCategory);
      data.taskCategories = getSortedTaskCategories();

      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);

      renderTaskCategories();
      setTaskCategoryFilter(newCategory.id);

      taskCategoryForm.reset();

      window.requestAnimationFrame(() => {
        if (taskCategoryNameInput instanceof HTMLInputElement) {
          taskCategoryNameInput.focus();
        }
      });
    }

    function deleteTaskCategory(categoryId) {
      if (!categoryId) {
        return;
      }

      if (!Array.isArray(data.taskCategories)) {
        return;
      }

      const category = data.taskCategories.find(
        (item) => item && typeof item.id === 'string' && item.id === categoryId,
      );

      if (!category) {
        return;
      }

      const confirmationMessage = `Supprimer la catégorie « ${category.name} » ? Les tâches associées seront déplacées dans « Sans catégorie ».`;
      if (typeof window !== 'undefined' && !window.confirm(confirmationMessage)) {
        return;
      }

      data.taskCategories = data.taskCategories.filter(
        (item) => item && typeof item.id === 'string' && item.id !== categoryId,
      );

      if (Array.isArray(data.tasks)) {
        data.tasks.forEach((task) => {
          if (task && task.categoryId === categoryId) {
            task.categoryId = '';
          }
        });

        data.tasks = data.tasks.map((task) => normalizeTask(task)).sort(compareTasks);
      }

      if (taskCategoryFilter === categoryId) {
        taskCategoryFilter = TASK_CATEGORY_FILTER_ALL;
      }

      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);

      renderTaskCategories();
      renderTasks();
    }

    function handleTaskAttachmentChange() {
      if (!(taskAttachmentInput instanceof HTMLInputElement)) {
        return;
      }

      const file =
        taskAttachmentInput.files && taskAttachmentInput.files.length > 0
          ? taskAttachmentInput.files[0]
          : null;

      taskAttachmentInput.setCustomValidity('');

      if (!file) {
        removeAttachmentOnSubmit = false;
        updateTaskAttachmentPreview();
        return;
      }

      if (file.size > MAX_TASK_ATTACHMENT_SIZE) {
        const message = `Le document dépasse la taille maximale autorisée (${formatFileSize(
          MAX_TASK_ATTACHMENT_SIZE,
        )}).`;
        taskAttachmentInput.value = '';
        taskAttachmentInput.setCustomValidity(message);
        taskAttachmentInput.reportValidity();
        removeAttachmentOnSubmit = false;
        updateTaskAttachmentPreview();
        return;
      }

      removeAttachmentOnSubmit = false;
      updateTaskAttachmentPreview();
    }

    function handleTaskAttachmentRemove() {
      if (!(taskAttachmentInput instanceof HTMLInputElement)) {
        return;
      }

      const hasSelectedFile = Boolean(
        taskAttachmentInput.files && taskAttachmentInput.files.length > 0,
      );

      if (hasSelectedFile) {
        taskAttachmentInput.value = '';
        taskAttachmentInput.setCustomValidity('');
        removeAttachmentOnSubmit = false;
        updateTaskAttachmentPreview();
        return;
      }

      if (editingTaskExistingAttachment) {
        removeAttachmentOnSubmit = true;
        taskAttachmentInput.value = '';
        taskAttachmentInput.setCustomValidity('');
        updateTaskAttachmentPreview();
      }
    }

    function resetTaskAttachmentState() {
      if (taskAttachmentInput instanceof HTMLInputElement) {
        taskAttachmentInput.value = '';
        taskAttachmentInput.setCustomValidity('');
      }

      editingTaskExistingAttachment = null;
      removeAttachmentOnSubmit = false;
      updateTaskAttachmentPreview();
    }

    function updateTaskAttachmentPreview() {
      if (!taskAttachmentPreview) {
        return;
      }

      const selectedFile =
        taskAttachmentInput instanceof HTMLInputElement && taskAttachmentInput.files
          ? taskAttachmentInput.files[0] || null
          : null;

      const hasSelectedFile = Boolean(selectedFile);
      const hasExistingAttachment = Boolean(editingTaskExistingAttachment);
      const showRemovalMessage = removeAttachmentOnSubmit && hasExistingAttachment;

      if (taskAttachmentRemovedMessage) {
        if (showRemovalMessage) {
          const attachmentName =
            editingTaskExistingAttachment && editingTaskExistingAttachment.name
              ? ` « ${editingTaskExistingAttachment.name} »`
              : '';
          taskAttachmentRemovedMessage.textContent = `Le document${attachmentName} sera supprimé lors de l’enregistrement.`;
          taskAttachmentRemovedMessage.hidden = false;
          taskAttachmentRemovedMessage.removeAttribute('hidden');
        } else {
          taskAttachmentRemovedMessage.hidden = true;
          if (!taskAttachmentRemovedMessage.hasAttribute('hidden')) {
            taskAttachmentRemovedMessage.setAttribute('hidden', '');
          }
        }
      }

      if (showRemovalMessage || (!hasSelectedFile && !hasExistingAttachment)) {
        taskAttachmentPreview.hidden = true;
        if (!taskAttachmentPreview.hasAttribute('hidden')) {
          taskAttachmentPreview.setAttribute('hidden', '');
        }

        if (taskAttachmentPreviewOpen instanceof HTMLAnchorElement) {
          taskAttachmentPreviewOpen.hidden = true;
          taskAttachmentPreviewOpen.href = '#';
          taskAttachmentPreviewOpen.removeAttribute('download');
        }

        return;
      }

      taskAttachmentPreview.hidden = false;
      taskAttachmentPreview.removeAttribute('hidden');

      if (hasSelectedFile) {
        if (taskAttachmentPreviewName) {
          taskAttachmentPreviewName.textContent = selectedFile.name || 'Document joint';
        }

        if (taskAttachmentPreviewMeta) {
          taskAttachmentPreviewMeta.textContent = formatFileSize(selectedFile.size);
        }

        if (taskAttachmentPreviewOpen instanceof HTMLAnchorElement) {
          taskAttachmentPreviewOpen.hidden = true;
          taskAttachmentPreviewOpen.href = '#';
          taskAttachmentPreviewOpen.removeAttribute('download');
        }

        return;
      }

      const attachment = editingTaskExistingAttachment;
      if (attachment) {
        if (taskAttachmentPreviewName) {
          taskAttachmentPreviewName.textContent = attachment.name || 'Document joint';
        }

        if (taskAttachmentPreviewMeta) {
          const sizeLabel = attachment.size ? formatFileSize(Number(attachment.size)) : '';
          taskAttachmentPreviewMeta.textContent = sizeLabel || 'Document existant';
        }

        if (taskAttachmentPreviewOpen instanceof HTMLAnchorElement) {
          const href =
            (typeof attachment.dataUrl === 'string' && attachment.dataUrl) ||
            (typeof attachment.url === 'string' && attachment.url) ||
            '';
          if (href) {
            taskAttachmentPreviewOpen.hidden = false;
            taskAttachmentPreviewOpen.href = href;
            if (attachment.name && attachment.dataUrl) {
              taskAttachmentPreviewOpen.setAttribute('download', attachment.name);
            } else {
              taskAttachmentPreviewOpen.removeAttribute('download');
            }
          } else {
            taskAttachmentPreviewOpen.hidden = true;
            taskAttachmentPreviewOpen.href = '#';
            taskAttachmentPreviewOpen.removeAttribute('download');
          }
        }
      }
    }

    function formatFileSize(bytes) {
      if (!Number.isFinite(bytes) || bytes <= 0) {
        return '';
      }

      const units = ['octet', 'Ko', 'Mo', 'Go'];
      let value = bytes;
      let unitIndex = 0;

      while (value >= 1024 && unitIndex < units.length - 1) {
        value /= 1024;
        unitIndex += 1;
      }

      const formattedValue =
        unitIndex === 0
          ? numberFormatter.format(Math.round(value))
          : fileSizeFormatter.format(value);

      if (unitIndex === 0) {
        const plural = Math.round(value) > 1 ? 's' : '';
        return `${formattedValue} octet${plural}`;
      }

      return `${formattedValue} ${units[unitIndex]}`;
    }

    function buildAttachmentFromFile(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = () => {
          const result = typeof reader.result === 'string' ? reader.result : '';
          if (!result) {
            reject(new Error('Impossible de lire le document joint.'));
            return;
          }

          resolve(
            normalizeTaskAttachment({
              name: file.name,
              size: file.size,
              type: file.type,
              dataUrl: result,
              uploadedAt: new Date().toISOString(),
            }),
          );
        };

        reader.onerror = () => {
          reject(reader.error || new Error('Impossible de lire le document joint.'));
        };

        reader.readAsDataURL(file);
      });
    }

    async function createTaskFromForm() {
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
          if (option && option.value && option.value !== TASK_MEMBER_NONE_VALUE) {
            members.push(option.value);
          }
        });
      }

      const allowedMembers = members.filter((u) => teamMembersById.has(u));

      const categoryIdRaw = (formData.get('task-category') || '').toString().trim();
      const categoryId =
        categoryIdRaw && taskCategoriesById.has(categoryIdRaw) ? categoryIdRaw : '';

      let attachment = null;
      if (taskAttachmentInput instanceof HTMLInputElement) {
        const file =
          taskAttachmentInput.files && taskAttachmentInput.files.length > 0
            ? taskAttachmentInput.files[0]
            : null;

        taskAttachmentInput.setCustomValidity('');

        if (file) {
          if (file.size > MAX_TASK_ATTACHMENT_SIZE) {
            const message = `Le document dépasse la taille maximale autorisée (${formatFileSize(
              MAX_TASK_ATTACHMENT_SIZE,
            )}).`;
            taskAttachmentInput.setCustomValidity(message);
            taskAttachmentInput.reportValidity();
            return;
          }

          attachment = await buildAttachmentFromFile(file);
        }
      }

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
        categoryId,
        attachment,
        isArchived: false,
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

    function filterTasksByStatus(tasks, status = taskStatusFilter) {
      const normalized = TASK_STATUS_VALUES.has(status) ? status : 'active';
      if (!Array.isArray(tasks)) {
        return [];
      }

      return tasks.filter((task) => {
        if (!task || typeof task !== 'object') {
          return false;
        }

        const archived = Boolean(task.isArchived);
        return normalized === 'archived' ? archived : !archived;
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

      const tasksForStatus = filterTasksByStatus(tasks);
      const counts = buildTaskCounts(tasksForStatus);
      updateTaskCategoryCountBadges(counts);

      const categories = Array.from(taskCategoriesById.values());
      const boardColumns = [
        {
          id: TASK_CATEGORY_FILTER_UNCATEGORIZED,
          name: 'Sans catégorie',
          color: '#94a3b8',
          tasks: [],
        },
        ...categories.map((category) => ({
          id: category.id,
          name: category.name,
          color: category.color,
          tasks: [],
        })),
      ];

      const columnsById = new Map(boardColumns.map((column) => [column.id, column]));

      tasksForStatus.forEach((task) => {
        if (!task || typeof task !== 'object') {
          return;
        }

        const rawCategoryId =
          typeof task.categoryId === 'string' && task.categoryId ? task.categoryId : '';
        const columnId = columnsById.has(rawCategoryId)
          ? rawCategoryId
          : TASK_CATEGORY_FILTER_UNCATEGORIZED;
        const column = columnsById.get(columnId);
        if (column) {
          column.tasks.push(task);
        }
      });

      let visibleColumns;
      if (taskCategoryFilter === TASK_CATEGORY_FILTER_ALL) {
        visibleColumns = boardColumns;
      } else {
        const targetColumn = columnsById.get(taskCategoryFilter);
        visibleColumns = targetColumn ? [targetColumn] : boardColumns;
      }

      if (!Array.isArray(visibleColumns) || visibleColumns.length === 0) {
        visibleColumns = boardColumns;
      }

      const visibleTaskCount = visibleColumns.reduce((total, column) => {
        if (!column || !Array.isArray(column.tasks)) {
          return total;
        }
        return total + column.tasks.length;
      }, 0);

      updateTaskCountDisplay(tasksForStatus.length, visibleTaskCount);

      taskList.innerHTML = '';

      const isFilteredView = taskCategoryFilter !== TASK_CATEGORY_FILTER_ALL;
      if (taskEmptyState) {
        if (visibleTaskCount === 0) {
          taskEmptyState.hidden = false;
          taskEmptyState.removeAttribute('hidden');
          if (taskStatusFilter === 'archived') {
            taskEmptyState.textContent = isFilteredView
              ? 'Aucune carte archivée dans cette colonne.'
              : 'Aucune carte archivée pour le moment.';
          } else {
            const defaultMessage =
              taskEmptyStateDefaultText ||
              'Aucune carte n’est disponible pour le moment. Ajoutez votre première tâche depuis l’onglet «\u00a0Créer une tâche\u00a0».';
            taskEmptyState.textContent = isFilteredView
              ? 'Aucune carte dans cette colonne pour le moment.'
              : defaultMessage;
          }
        } else {
          taskEmptyState.hidden = true;
          if (!taskEmptyState.hasAttribute('hidden')) {
            taskEmptyState.setAttribute('hidden', '');
          }
          taskEmptyState.textContent =
            taskStatusFilter === 'archived'
              ? 'Aucune carte archivée pour le moment.'
              : taskEmptyStateDefaultText;
        }
      }

      const fragment = document.createDocumentFragment();
      const isArchivedView = taskStatusFilter === 'archived';

      visibleColumns.forEach((column) => {
        if (!column) {
          return;
        }

        const isFocused = isFilteredView && column.id === taskCategoryFilter;
        const columnElement = createTaskColumn(column, counts, {
          isFocused,
          isArchivedView,
        });
        if (columnElement) {
          fragment.appendChild(columnElement);
        }
      });

      taskList.appendChild(fragment);
      updateTaskSummaryCounts();
      scheduleTaskCategoryIndicatorUpdate();
    }

    function createTaskColumn(column, counts, options = {}) {
      if (!column || typeof column !== 'object') {
        return null;
      }

      const { isFocused = false, isArchivedView = false } = options;
      const columnId =
        typeof column.id === 'string' && column.id ? column.id : TASK_CATEGORY_FILTER_UNCATEGORIZED;
      const columnLabel =
        typeof column.name === 'string' && column.name.trim()
          ? column.name.trim()
          : columnId === TASK_CATEGORY_FILTER_UNCATEGORIZED
          ? 'Sans catégorie'
          : 'Catégorie';

      const columnElement = document.createElement('section');
      columnElement.className = 'task-board-column';
      columnElement.dataset.taskCategoryId =
        columnId === TASK_CATEGORY_FILTER_UNCATEGORIZED ? '' : columnId;
      columnElement.setAttribute('role', 'region');

      const headerId = `task-board-column-${columnId}`;
      columnElement.setAttribute('aria-labelledby', headerId);

      if (isFocused) {
        columnElement.classList.add('task-board-column--focused');
      }

      const baseColor =
        columnId === TASK_CATEGORY_FILTER_UNCATEGORIZED
          ? '#94a3b8'
          : column.color || DEFAULT_TASK_CATEGORY_COLOR;
      columnElement.style.setProperty(
        '--task-board-color',
        normalizeTaskCategoryColor(baseColor),
      );

      const header = document.createElement('header');
      header.className = 'task-board-column-header';

      const title = document.createElement('h3');
      title.className = 'task-board-column-title';
      title.id = headerId;

      const dot = document.createElement('span');
      dot.className = 'task-board-column-dot';
      dot.setAttribute('aria-hidden', 'true');

      const titleLabel = document.createElement('span');
      titleLabel.textContent = columnLabel;

      title.append(dot, titleLabel);

      const countLabel = document.createElement('span');
      countLabel.className = 'task-board-column-count';
      countLabel.textContent = formatTaskCountLabel(counts.get(columnId) || 0);

      header.append(title, countLabel);

      const body = document.createElement('div');
      body.className = 'task-board-column-body';

      const tasksInColumn = Array.isArray(column.tasks) ? column.tasks : [];
      if (tasksInColumn.length === 0) {
        const emptyMessage = document.createElement('p');
        emptyMessage.className = 'task-board-column-empty';
        if (isArchivedView) {
          emptyMessage.textContent = 'Aucune carte archivée dans cette colonne.';
        } else if (columnId === TASK_CATEGORY_FILTER_UNCATEGORIZED) {
          emptyMessage.textContent = 'Aucune tâche sans catégorie pour le moment.';
        } else {
          emptyMessage.textContent = 'Aucune tâche dans cette catégorie pour le moment.';
        }
        body.appendChild(emptyMessage);
      } else {
        const cardsFragment = document.createDocumentFragment();
        tasksInColumn.forEach((task) => {
          const card = createTaskCard(task);
          if (card) {
            cardsFragment.appendChild(card);
          }
        });
        body.appendChild(cardsFragment);
      }

      columnElement.append(header, body);
      return columnElement;
    }

    function createTaskCard(task) {
      if (!task || typeof task !== 'object') {
        return null;
      }

      const card = document.createElement('article');
      card.className = 'task-item';
      card.dataset.id = task.id || '';
      card.dataset.categoryId = task.categoryId || '';

      const isArchived = Boolean(task.isArchived);
      if (isArchived) {
        card.classList.add('task-item--archived');
      }

      const isAssignedToMe =
        Array.isArray(task.assignedMembers) && task.assignedMembers.includes(currentUser);
      if (isAssignedToMe) {
        card.classList.add('task-item--assigned-me');
      } else {
        const taskColor = normalizeTaskColor(task.color);
        card.style.setProperty('--task-color', taskColor);
        card.style.setProperty('--task-color-soft', hexToRgba(taskColor, 0.12));
      }

      const header = document.createElement('div');
      header.className = 'task-item-header';

      const titleEl = document.createElement('h3');
      titleEl.className = 'task-title';

      if (isAssignedToMe) {
        const assignedBadge = document.createElement('span');
        assignedBadge.className = 'task-assigned-badge';
        assignedBadge.setAttribute('title', 'Tâche qui vous est attribuée');
        assignedBadge.setAttribute('aria-label', 'Tâche qui vous est attribuée');
        assignedBadge.textContent = '⭐';
        titleEl.appendChild(assignedBadge);
      }

      const titleText = document.createElement('span');
      titleText.className = 'task-title-text';
      titleText.textContent = task.title || 'Nouvelle tâche';
      titleEl.appendChild(titleText);

      const metaContainer = document.createElement('div');
      metaContainer.className = 'task-meta';

      if (isArchived) {
        const statusBadge = document.createElement('span');
        statusBadge.className = 'task-status-badge';
        statusBadge.textContent = 'Archivée';
        metaContainer.appendChild(statusBadge);
      }

      const category =
        typeof task.categoryId === 'string' && task.categoryId
          ? taskCategoriesById.get(task.categoryId)
          : null;
      const categoryChip = document.createElement('span');
      categoryChip.className = 'task-item-category';
      if (category) {
        categoryChip.textContent = category.name;
        categoryChip.style.setProperty('--task-category-color', category.color);
      } else {
        categoryChip.textContent = 'Sans catégorie';
        categoryChip.classList.add('task-item-category--default');
      }
      metaContainer.appendChild(categoryChip);

      if (task.dueDate) {
        const dueDate = parseDateInput(task.dueDate);
        const dueDateEl = document.createElement('span');
        dueDateEl.className = 'task-due-date';
        const formattedDueDate =
          dueDate instanceof Date && !Number.isNaN(dueDate.getTime())
            ? capitalizeLabel(taskDateFormatter.format(dueDate))
            : task.dueDate;
        dueDateEl.textContent = `Échéance : ${formattedDueDate}`;

        if (dueDate && dueDate.getTime() < startOfDay(new Date()).getTime() && !isArchived) {
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

      const archiveButton = document.createElement('button');
      archiveButton.type = 'button';
      archiveButton.className = 'task-action-button';
      archiveButton.dataset.action = isArchived ? 'restore-task' : 'archive-task';
      archiveButton.dataset.taskId = task.id || '';
      archiveButton.textContent = isArchived ? 'Restaurer' : 'Archiver';
      actions.appendChild(archiveButton);

      const deleteButton = document.createElement('button');
      deleteButton.type = 'button';
      deleteButton.className = 'task-action-button task-action-button--danger';
      deleteButton.dataset.action = 'delete-task';
      deleteButton.dataset.taskId = task.id || '';
      deleteButton.textContent = 'Supprimer';

      actions.appendChild(deleteButton);

      header.append(titleEl, metaContainer, actions);

      const descriptionEl = document.createElement('p');
      descriptionEl.className = 'task-description';
      if (task.description) {
        descriptionEl.textContent = task.description;
      } else {
        descriptionEl.textContent = 'Aucune description pour cette tâche.';
        descriptionEl.classList.add('task-description--empty');
      }

      let attachmentNode = null;
      if (task.attachment && typeof task.attachment === 'object') {
        const attachmentHref =
          (typeof task.attachment.dataUrl === 'string' && task.attachment.dataUrl) ||
          (typeof task.attachment.url === 'string' && task.attachment.url) ||
          '';

        if (attachmentHref) {
          attachmentNode = document.createElement('div');
          attachmentNode.className = 'task-attachment';

          const attachmentLink = document.createElement('a');
          attachmentLink.className = 'task-attachment-link';
          attachmentLink.href = attachmentHref;
          attachmentLink.target = '_blank';
          attachmentLink.rel = 'noopener';

          if (
            task.attachment.name &&
            typeof task.attachment.dataUrl === 'string' &&
            task.attachment.dataUrl
          ) {
            attachmentLink.setAttribute('download', task.attachment.name);
          } else {
            attachmentLink.removeAttribute('download');
          }

          const attachmentIcon = document.createElement('span');
          attachmentIcon.className = 'task-attachment-link-icon';
          attachmentIcon.setAttribute('aria-hidden', 'true');
          attachmentIcon.textContent = '📎';

          const attachmentLabel = document.createElement('span');
          attachmentLabel.textContent = task.attachment.name || 'Document joint';

          attachmentLink.append(attachmentIcon, attachmentLabel);
          attachmentNode.appendChild(attachmentLink);

          const rawSize = Number(task.attachment.size);
          const sizeLabel = Number.isFinite(rawSize) && rawSize > 0 ? formatFileSize(rawSize) : '';
          if (sizeLabel) {
            const attachmentMeta = document.createElement('span');
            attachmentMeta.className = 'task-attachment-meta';
            attachmentMeta.textContent = sizeLabel;
            attachmentNode.appendChild(attachmentMeta);
          }
        }
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

      const nodesToAppend = [header, descriptionEl];
      if (attachmentNode) {
        nodesToAppend.push(attachmentNode);
      }
      nodesToAppend.push(commentsSection);
      card.append(...nodesToAppend);

      return card;
    }

    function renderTeamChat() {
      if (!teamChatList) {
        return;
      }

      const messages = Array.isArray(data.teamChatMessages)
        ? data.teamChatMessages
            .filter((message) => message && typeof message === 'object' && message.content)
            .slice()
        : [];

      messages.sort((a, b) => {
        const dateA = typeof a.createdAt === 'string' ? a.createdAt : '';
        const dateB = typeof b.createdAt === 'string' ? b.createdAt : '';
        return dateA.localeCompare(dateB);
      });

      teamChatList.innerHTML = '';

      if (teamChatEmptyState) {
        if (messages.length === 0) {
          teamChatEmptyState.hidden = false;
          teamChatEmptyState.removeAttribute('hidden');
        } else {
          teamChatEmptyState.hidden = true;
          if (!teamChatEmptyState.hasAttribute('hidden')) {
            teamChatEmptyState.setAttribute('hidden', '');
          }
        }
      }

      if (messages.length === 0) {
        return;
      }

      const fragment = document.createDocumentFragment();

      messages.forEach((message) => {
        if (!message || typeof message !== 'object' || !message.content) {
          return;
        }

        const item = document.createElement('li');
        item.className = 'team-chat-message';
        item.setAttribute('role', 'listitem');

        const meta = document.createElement('div');
        meta.className = 'team-chat-message-meta';

        const authorLabel = formatMemberName(message.author);
        let dateLabel = '';
        if (message.createdAt && !Number.isNaN(new Date(message.createdAt).getTime())) {
          dateLabel = capitalizeLabel(
            taskCommentDateFormatter.format(new Date(message.createdAt)),
          );
        }
        meta.textContent = dateLabel ? `${authorLabel} · ${dateLabel}` : authorLabel;

        const content = document.createElement('p');
        content.className = 'team-chat-message-content';
        content.textContent = message.content;

        item.append(meta, content);
        fragment.appendChild(item);
      });

      teamChatList.appendChild(fragment);

      window.requestAnimationFrame(() => {
        if (teamChatList instanceof HTMLElement) {
          teamChatList.scrollTop = teamChatList.scrollHeight;
        }
      });
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

    function updateTaskCountDisplay(totalCount, visibleCount = totalCount) {
      const totalLabel = formatTaskCountLabel(totalCount);
      const visibleLabel = formatTaskCountLabel(visibleCount);

      const badgeLabel =
        visibleCount !== totalCount ? `${visibleLabel} · ${totalLabel}` : totalLabel;

      if (taskCountBadge) {
        taskCountBadge.textContent = badgeLabel;
      }

    }

    function updateTaskSummaryCounts() {
      const tasks = Array.isArray(data.tasks) ? data.tasks : [];
      let activeCount = 0;
      let archivedCount = 0;

      tasks.forEach((task) => {
        if (!task || typeof task !== 'object') {
          return;
        }

        if (task.isArchived) {
          archivedCount += 1;
        } else {
          activeCount += 1;
        }
      });

      if (taskSummaryActiveEl) {
        taskSummaryActiveEl.textContent = numberFormatter.format(activeCount);
      }
      if (taskSummaryArchivedEl) {
        taskSummaryArchivedEl.textContent = numberFormatter.format(archivedCount);
      }
      if (taskStatusCountEls.active instanceof HTMLElement) {
        taskStatusCountEls.active.textContent = numberFormatter.format(activeCount);
      }
      if (taskStatusCountEls.archived instanceof HTMLElement) {
        taskStatusCountEls.archived.textContent = numberFormatter.format(archivedCount);
      }
    }

    function archiveTask(taskId) {
      if (!taskId || !Array.isArray(data.tasks)) {
        return;
      }

      const task = data.tasks.find((item) => item && item.id === taskId);
      if (!task || task.isArchived) {
        return;
      }

      task.isArchived = true;
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderTasks();
    }

    function restoreTask(taskId) {
      if (!taskId || !Array.isArray(data.tasks)) {
        return;
      }

      const task = data.tasks.find((item) => item && item.id === taskId);
      if (!task || !task.isArchived) {
        return;
      }

      task.isArchived = false;
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderTasks();
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

          // Ouvre la page « Créer une tâche » pour faciliter la modification
          showPage && showPage('tasks-create');

      // Remplit le formulaire
      if (taskTitleInput instanceof HTMLInputElement) taskTitleInput.value = t.title || '';
      if (taskDueDateInput instanceof HTMLInputElement) taskDueDateInput.value = t.dueDate || '';
      if (taskColorInput instanceof HTMLInputElement) taskColorInput.value = normalizeTaskColor(t.color || DEFAULT_TASK_COLOR);
      if (taskDescriptionInput instanceof HTMLTextAreaElement) taskDescriptionInput.value = t.description || '';

        if (taskMemberSelect instanceof HTMLSelectElement && !taskMemberSelect.disabled) {
          const assignedMembers = Array.isArray(t.assignedMembers) ? t.assignedMembers.slice() : [];
          setTaskMemberSelection(assignedMembers);
        }

          if (taskCategorySelect instanceof HTMLSelectElement) {
            setTaskFormCategory(t.categoryId);
          }

          if (taskAttachmentInput instanceof HTMLInputElement) {
            taskAttachmentInput.value = '';
            taskAttachmentInput.setCustomValidity('');
          }

          editingTaskExistingAttachment =
            t.attachment && typeof t.attachment === 'object' ? { ...t.attachment } : null;
          removeAttachmentOnSubmit = false;
          updateTaskAttachmentPreview();

        editingTaskId = t.id;
        setTaskFormMode('edit');

      // Focus UX
      window.requestAnimationFrame(() => {
        taskTitleInput && taskTitleInput.focus();
        taskTitleInput && taskTitleInput.setSelectionRange(taskTitleInput.value.length, taskTitleInput.value.length);
      });
    }

    async function applyTaskEditsFromForm() {
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
              if (option && option.value && option.value !== TASK_MEMBER_NONE_VALUE) {
                members.push(option.value);
              }
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
      const categoryIdRaw = (formData.get('task-category') || '').toString().trim();
      task.categoryId =
        categoryIdRaw && taskCategoriesById.has(categoryIdRaw) ? categoryIdRaw : '';

      if (taskAttachmentInput instanceof HTMLInputElement) {
        const file =
          taskAttachmentInput.files && taskAttachmentInput.files.length > 0
            ? taskAttachmentInput.files[0]
            : null;

        taskAttachmentInput.setCustomValidity('');

        if (file) {
          if (file.size > MAX_TASK_ATTACHMENT_SIZE) {
            const message = `Le document dépasse la taille maximale autorisée (${formatFileSize(
              MAX_TASK_ATTACHMENT_SIZE,
            )}).`;
            taskAttachmentInput.setCustomValidity(message);
            taskAttachmentInput.reportValidity();
            return;
          }

          task.attachment = await buildAttachmentFromFile(file);
          removeAttachmentOnSubmit = false;
        } else if (removeAttachmentOnSubmit) {
          task.attachment = null;
        }
      }

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

    function getEmailApi() {
      if (typeof window === 'undefined') {
        return null;
      }

      if (window.UManager && window.UManager.emailApi) {
        return window.UManager.emailApi;
      }

      return null;
    }

    function getSubscriptionPlan(planId) {
      const key = typeof planId === 'string' ? planId.trim().toLowerCase() : '';
      if (SUBSCRIPTION_PLANS[key]) {
        return SUBSCRIPTION_PLANS[key];
      }
      return SUBSCRIPTION_PLANS[DEFAULT_SUBSCRIPTION_ID];
    }

    function setHomeTeamFeedback(message, type = '') {
      if (!(homeTeamFeedback instanceof HTMLElement)) {
        return;
      }

      homeTeamFeedback.classList.remove('form-feedback--error', 'form-feedback--success');

      if (!message) {
        homeTeamFeedback.textContent = '';
        homeTeamFeedback.hidden = true;
        homeTeamFeedback.setAttribute('hidden', '');
        return;
      }

      homeTeamFeedback.textContent = message;
      homeTeamFeedback.hidden = false;
      homeTeamFeedback.removeAttribute('hidden');

      if (type === 'error') {
        homeTeamFeedback.classList.add('form-feedback--error');
      } else if (type === 'success') {
        homeTeamFeedback.classList.add('form-feedback--success');
      }
    }

    function renderHomeTeams() {
      const teams = Array.isArray(data.teams) ? data.teams.slice() : [];
      const currentTeamIdValue = typeof data.currentTeamId === 'string' ? data.currentTeamId : '';

      if (homeCurrentTeamEl) {
        const currentTeam = teams.find((team) => team && team.id === currentTeamIdValue);
        homeCurrentTeamEl.textContent = currentTeam ? currentTeam.name : '—';
      }

      if (!(homeTeamsList instanceof HTMLElement)) {
        if (homeTeamsEmpty) {
          if (teams.length === 0) {
            homeTeamsEmpty.hidden = false;
            homeTeamsEmpty.removeAttribute('hidden');
          } else {
            homeTeamsEmpty.hidden = true;
            if (!homeTeamsEmpty.hasAttribute('hidden')) {
              homeTeamsEmpty.setAttribute('hidden', '');
            }
          }
        }
        return;
      }

      homeTeamsList.innerHTML = '';

      if (teams.length === 0) {
        if (homeTeamsEmpty) {
          homeTeamsEmpty.hidden = false;
          homeTeamsEmpty.removeAttribute('hidden');
        }
        return;
      }

      if (homeTeamsEmpty) {
        homeTeamsEmpty.hidden = true;
        if (!homeTeamsEmpty.hasAttribute('hidden')) {
          homeTeamsEmpty.setAttribute('hidden', '');
        }
      }

      const fragment = document.createDocumentFragment();
      teams.forEach((team) => {
        if (!team || typeof team !== 'object') {
          return;
        }

        const item = document.createElement('li');
        item.className = 'home-team-item';

        const main = document.createElement('div');
        main.className = 'home-team-main';

        const nameEl = document.createElement('p');
        nameEl.className = 'home-team-name';
        nameEl.textContent = team.name || 'Équipe';
        main.appendChild(nameEl);

        if (team.role) {
          const roleEl = document.createElement('p');
          roleEl.className = 'home-team-role';
          roleEl.textContent = `Rôle : ${team.role}`;
          main.appendChild(roleEl);
        }

        const actions = document.createElement('div');
        actions.className = 'home-team-actions';

        if (team.id === currentTeamIdValue) {
          const tag = document.createElement('span');
          tag.className = 'home-team-tag';
          tag.textContent = 'Équipe actuelle';
          actions.appendChild(tag);
        } else {
          const setButton = document.createElement('button');
          setButton.type = 'button';
          setButton.className = 'home-team-action';
          setButton.dataset.action = 'home-team-set-current';
          setButton.dataset.teamId = team.id;
          setButton.textContent = 'Définir comme équipe actuelle';
          actions.appendChild(setButton);
        }

        if (teams.length > 1) {
          const removeButton = document.createElement('button');
          removeButton.type = 'button';
          removeButton.className = 'home-team-action home-team-action--danger';
          removeButton.dataset.action = 'home-team-remove';
          removeButton.dataset.teamId = team.id;
          removeButton.textContent = 'Retirer de mes équipes';
          actions.appendChild(removeButton);
        }

        item.append(main, actions);
        fragment.appendChild(item);
      });

      homeTeamsList.appendChild(fragment);
    }

    function addTeamMembership(name, role) {
      const normalizedName = typeof name === 'string' ? name.trim() : '';
      const normalizedRole = typeof role === 'string' ? role.trim() : '';

      if (!normalizedName) {
        setHomeTeamFeedback("Veuillez indiquer le nom de l'équipe.", 'error');
        return false;
      }

      const teams = Array.isArray(data.teams) ? data.teams.slice() : [];
      const alreadyExists = teams.some(
        (team) =>
          team &&
          typeof team === 'object' &&
          typeof team.name === 'string' &&
          team.name.trim().toLowerCase() === normalizedName.toLowerCase(),
      );

      if (alreadyExists) {
        setHomeTeamFeedback('Cette équipe est déjà enregistrée.', 'error');
        return false;
      }

      const newTeam = {
        id: generateId('user-team'),
        name: normalizedName,
        role: normalizedRole || 'Membre',
      };

      teams.push(newTeam);
      data.teams = teams;

      if (!data.currentTeamId) {
        data.currentTeamId = newTeam.id;
      }

      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderHomeTeams();
      setHomeTeamFeedback(`L'équipe « ${newTeam.name} » a été ajoutée.`, 'success');
      return true;
    }

    function setCurrentTeam(teamId) {
      if (!Array.isArray(data.teams)) {
        setHomeTeamFeedback("Aucune équipe n'est disponible.", 'error');
        return false;
      }

      const target = data.teams.find((team) => team && team.id === teamId);
      if (!target) {
        setHomeTeamFeedback("Équipe introuvable.", 'error');
        return false;
      }

      if (data.currentTeamId === teamId) {
        setHomeTeamFeedback(`« ${target.name} » est déjà votre équipe active.`, 'error');
        return false;
      }

      data.currentTeamId = teamId;
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderHomeTeams();
      setHomeTeamFeedback(`« ${target.name} » est désormais votre équipe active.`, 'success');
      return true;
    }

    function removeTeamMembership(teamId) {
      if (!Array.isArray(data.teams)) {
        setHomeTeamFeedback("Aucune équipe à retirer.", 'error');
        return false;
      }

      if (data.teams.length <= 1) {
        setHomeTeamFeedback('Vous devez conserver au moins une équipe.', 'error');
        return false;
      }

      const previousLength = data.teams.length;
      data.teams = data.teams.filter((team) => team && team.id !== teamId);

      if (data.teams.length === previousLength) {
        setHomeTeamFeedback("Équipe introuvable.", 'error');
        return false;
      }

      if (!data.teams.some((team) => team && team.id === data.currentTeamId)) {
        data.currentTeamId = data.teams[0] ? data.teams[0].id : '';
      }

      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      renderHomeTeams();
      setHomeTeamFeedback('Équipe retirée avec succès.', 'success');
      return true;
    }

    function renderHomeSubscription() {
      const plan = getSubscriptionPlan(data.subscription);

      if (homeSubscriptionNameEl) {
        homeSubscriptionNameEl.textContent = plan.name;
      }

      if (homeSubscriptionPriceEl) {
        homeSubscriptionPriceEl.textContent = plan.priceLabel;
      }

      if (homeSubscriptionContactLimitEl) {
        homeSubscriptionContactLimitEl.textContent = numberFormatter.format(
          Number.isFinite(plan.contactLimit) ? plan.contactLimit : 0,
        );
      }

      if (homeSubscriptionTaskLimitEl) {
        homeSubscriptionTaskLimitEl.textContent = numberFormatter.format(
          Number.isFinite(plan.taskLimit) ? plan.taskLimit : 0,
        );
      }

      if (homeSubscriptionSelect instanceof HTMLSelectElement) {
        const selectedValue = plan.id;
        const fragment = document.createDocumentFragment();

        SUBSCRIPTION_PLAN_ORDER.forEach((planId) => {
          const optionPlan = SUBSCRIPTION_PLANS[planId];
          if (!optionPlan) {
            return;
          }
          const option = document.createElement('option');
          option.value = optionPlan.id;
          option.textContent = `${optionPlan.name} — ${optionPlan.priceLabel}`;
          fragment.appendChild(option);
        });

        homeSubscriptionSelect.innerHTML = '';
        homeSubscriptionSelect.appendChild(fragment);
        homeSubscriptionSelect.value = selectedValue;
      }

      if (homeSubscriptionFeaturesList) {
        homeSubscriptionFeaturesList.innerHTML = '';
        const features = Array.isArray(plan.features) ? plan.features : [];
        const fragment = document.createDocumentFragment();

        if (features.length === 0) {
          const placeholder = document.createElement('li');
          placeholder.textContent = 'Aucune caractéristique supplémentaire.';
          fragment.appendChild(placeholder);
        } else {
          features.forEach((feature) => {
            const item = document.createElement('li');
            item.textContent = feature;
            fragment.appendChild(item);
          });
        }

        homeSubscriptionFeaturesList.appendChild(fragment);
      }
    }

    function renderHomeDirectoryLimit() {
      const plan = getSubscriptionPlan(data.subscription);
      const limit = Number.isFinite(plan.contactLimit) ? Number(plan.contactLimit) : 0;
      const totalContacts = Array.isArray(data.contacts) ? data.contacts.length : 0;

      if (homeDirectoryLimitEl) {
        const limitText =
          limit > 0
            ? `${numberFormatter.format(limit)} contact${limit > 1 ? 's' : ''} max`
            : 'Aucune limite configurée';
        homeDirectoryLimitEl.textContent = `Limite actuelle : ${limitText}`;
      }

      if (homeDirectoryUsageEl) {
        if (limit > 0) {
          const remaining = Math.max(limit - totalContacts, 0);
          const remainingText =
            remaining > 0
              ? `${numberFormatter.format(remaining)} restant${remaining > 1 ? 's' : ''}`
              : 'Limite atteinte';
          homeDirectoryUsageEl.textContent = `${numberFormatter.format(totalContacts)} contact${
            totalContacts > 1 ? 's' : ''
          } enregistrés · ${remainingText}`;
        } else {
          homeDirectoryUsageEl.textContent = `${numberFormatter.format(totalContacts)} contact${
            totalContacts > 1 ? 's' : ''
          } enregistrés`;
        }
      }

      if (homeDirectoryProgressBar instanceof HTMLElement) {
        const effectiveLimit = limit > 0 ? limit : Math.max(totalContacts, 1);
        const ratio = limit > 0 ? Math.min(totalContacts / limit, 1) : 0;
        const percent = Math.max(0, Math.min(ratio, 1)) * 100;
        homeDirectoryProgressBar.style.width = `${percent}%`;
        homeDirectoryProgressBar.setAttribute('aria-valuemin', '0');
        homeDirectoryProgressBar.setAttribute('aria-valuemax', effectiveLimit.toString());
        homeDirectoryProgressBar.setAttribute(
          'aria-valuenow',
          Math.min(totalContacts, effectiveLimit).toString(),
        );
        const label = limit > 0
          ? `${numberFormatter.format(totalContacts)} sur ${numberFormatter.format(limit)} contacts`
          : `${numberFormatter.format(totalContacts)} contacts enregistrés`;
        homeDirectoryProgressBar.setAttribute('aria-label', `Utilisation du répertoire : ${label}`);
      }
    }

    function renderHomeNews() {
      if (!(homeNewsList instanceof HTMLElement)) {
        return;
      }

      const articles = Array.isArray(newsData.articles) ? newsData.articles.slice() : [];
      articles.sort((a, b) => {
        const dateA = typeof a.createdAt === 'string' ? a.createdAt : '';
        const dateB = typeof b.createdAt === 'string' ? b.createdAt : '';
        return dateB.localeCompare(dateA);
      });

      homeNewsList.innerHTML = '';

      if (articles.length === 0) {
        if (homeNewsEmpty) {
          homeNewsEmpty.hidden = false;
          homeNewsEmpty.removeAttribute('hidden');
        }
        if (homeNewsUpdatedEl) {
          homeNewsUpdatedEl.textContent = homeNewsUpdatedDefault || '';
        }
        return;
      }

      if (homeNewsEmpty) {
        homeNewsEmpty.hidden = true;
        if (!homeNewsEmpty.hasAttribute('hidden')) {
          homeNewsEmpty.setAttribute('hidden', '');
        }
      }

      const fragment = document.createDocumentFragment();
      articles.slice(0, NEWS_DISPLAY_LIMIT).forEach((article) => {
        if (!article || typeof article !== 'object') {
          return;
        }

        const item = document.createElement('li');
        item.className = 'home-news-item';

        const titleEl = document.createElement('h3');
        titleEl.className = 'home-news-title';
        titleEl.textContent = article.title || 'Annonce';
        item.appendChild(titleEl);

        const metaParts = [];
        if (article.createdAt) {
          const createdDate = new Date(article.createdAt);
          if (!Number.isNaN(createdDate.getTime())) {
            metaParts.push(newsDateFormatter.format(createdDate));
          }
        }
        if (article.author) {
          metaParts.push(`par ${article.author}`);
        }
        if (metaParts.length > 0) {
          const metaEl = document.createElement('p');
          metaEl.className = 'home-news-meta';
          metaEl.textContent = metaParts.join(' · ');
          item.appendChild(metaEl);
        }

        if (article.summary) {
          const summaryEl = document.createElement('p');
          summaryEl.className = 'home-news-summary';
          summaryEl.textContent = article.summary;
          item.appendChild(summaryEl);
        }

        if (article.content) {
          const contentEl = document.createElement('p');
          contentEl.className = 'home-news-content';
          contentEl.textContent = article.content;
          item.appendChild(contentEl);
        }

        fragment.appendChild(item);
      });

      homeNewsList.appendChild(fragment);

      if (homeNewsUpdatedEl) {
        const latest = articles[0];
        if (latest && latest.createdAt) {
          const latestDate = new Date(latest.createdAt);
          if (!Number.isNaN(latestDate.getTime())) {
            homeNewsUpdatedEl.textContent = `Dernière mise à jour : ${newsDateFormatter.format(
              latestDate,
            )}`;
          } else {
            homeNewsUpdatedEl.textContent = homeNewsUpdatedDefault || '';
          }
        } else {
          homeNewsUpdatedEl.textContent = homeNewsUpdatedDefault || '';
        }
      }
    }

    function setAdminArticleFeedback(message, type = '') {
      if (!(adminArticleFeedback instanceof HTMLElement)) {
        return;
      }

      adminArticleFeedback.classList.remove('form-feedback--error', 'form-feedback--success');

      if (!message) {
        adminArticleFeedback.textContent = '';
        adminArticleFeedback.hidden = true;
        adminArticleFeedback.setAttribute('hidden', '');
        return;
      }

      adminArticleFeedback.textContent = message;
      adminArticleFeedback.hidden = false;
      adminArticleFeedback.removeAttribute('hidden');

      if (type === 'error') {
        adminArticleFeedback.classList.add('form-feedback--error');
      } else if (type === 'success') {
        adminArticleFeedback.classList.add('form-feedback--success');
      }
    }

    function setAdminImpersonateFeedback(message, type = '') {
      if (!(adminImpersonateFeedback instanceof HTMLElement)) {
        return;
      }

      adminImpersonateFeedback.classList.remove('form-feedback--error', 'form-feedback--success');

      if (!message) {
        adminImpersonateFeedback.textContent = '';
        adminImpersonateFeedback.hidden = true;
        adminImpersonateFeedback.setAttribute('hidden', '');
        return;
      }

      adminImpersonateFeedback.textContent = message;
      adminImpersonateFeedback.hidden = false;
      adminImpersonateFeedback.removeAttribute('hidden');

      if (type === 'error') {
        adminImpersonateFeedback.classList.add('form-feedback--error');
      } else if (type === 'success') {
        adminImpersonateFeedback.classList.add('form-feedback--success');
      }
    }

    function addNewsArticle(article) {
      const normalized = normalizeNewsArticle(article);
      if (!normalized) {
        return null;
      }

      if (!newsData || typeof newsData !== 'object') {
        newsData = { articles: [] };
      }

      if (!Array.isArray(newsData.articles)) {
        newsData.articles = [];
      }

      newsData.articles = newsData.articles.filter(
        (existing) => existing && existing.id !== normalized.id,
      );
      newsData.articles.unshift(normalized);
      newsData.articles.sort((a, b) => {
        const dateA = typeof a.createdAt === 'string' ? a.createdAt : '';
        const dateB = typeof b.createdAt === 'string' ? b.createdAt : '';
        return dateB.localeCompare(dateA);
      });
      saveNewsStore(newsData);
      renderHomeNews();
      renderAdminNewsList();
      return normalized;
    }

    function deleteNewsArticle(articleId) {
      if (!newsData || !Array.isArray(newsData.articles)) {
        return false;
      }

      const previousLength = newsData.articles.length;
      newsData.articles = newsData.articles.filter(
        (article) => article && article.id !== articleId,
      );

      if (newsData.articles.length === previousLength) {
        return false;
      }

      saveNewsStore(newsData);
      renderHomeNews();
      renderAdminNewsList();
      return true;
    }

    function renderAdminNewsList() {
      if (!(adminNewsList instanceof HTMLElement)) {
        return;
      }

      const articles = Array.isArray(newsData.articles) ? newsData.articles.slice() : [];
      articles.sort((a, b) => {
        const dateA = typeof a.createdAt === 'string' ? a.createdAt : '';
        const dateB = typeof b.createdAt === 'string' ? b.createdAt : '';
        return dateB.localeCompare(dateA);
      });

      adminNewsList.innerHTML = '';

      if (articles.length === 0) {
        if (adminNewsEmpty) {
          adminNewsEmpty.hidden = false;
          adminNewsEmpty.removeAttribute('hidden');
        }
        return;
      }

      if (adminNewsEmpty) {
        adminNewsEmpty.hidden = true;
        if (!adminNewsEmpty.hasAttribute('hidden')) {
          adminNewsEmpty.setAttribute('hidden', '');
        }
      }

      const fragment = document.createDocumentFragment();
      articles.forEach((article) => {
        if (!article || typeof article !== 'object') {
          return;
        }

        const item = document.createElement('li');
        item.className = 'admin-news-item';

        const titleEl = document.createElement('h3');
        titleEl.className = 'admin-news-title';
        titleEl.textContent = article.title || 'Annonce';
        item.appendChild(titleEl);

        const metaParts = [];
        if (article.createdAt) {
          const createdDate = new Date(article.createdAt);
          if (!Number.isNaN(createdDate.getTime())) {
            metaParts.push(newsDateFormatter.format(createdDate));
          }
        }
        if (article.author) {
          metaParts.push(`par ${article.author}`);
        }
        if (metaParts.length > 0) {
          const metaEl = document.createElement('p');
          metaEl.className = 'admin-news-meta';
          metaEl.textContent = metaParts.join(' · ');
          item.appendChild(metaEl);
        }

        if (article.summary) {
          const summaryEl = document.createElement('p');
          summaryEl.className = 'admin-news-summary';
          summaryEl.textContent = article.summary;
          item.appendChild(summaryEl);
        }

        if (article.content) {
          const contentEl = document.createElement('p');
          contentEl.className = 'admin-news-content';
          contentEl.textContent = article.content;
          item.appendChild(contentEl);
        }

        const actions = document.createElement('div');
        actions.className = 'admin-news-actions';
        const deleteButton = document.createElement('button');
        deleteButton.type = 'button';
        deleteButton.className = 'ghost-button ghost-button--danger';
        deleteButton.dataset.action = 'delete-news';
        deleteButton.dataset.newsId = article.id;
        deleteButton.textContent = 'Supprimer';
        actions.appendChild(deleteButton);
        item.appendChild(actions);

        fragment.appendChild(item);
      });

      adminNewsList.appendChild(fragment);
    }

    function renderAdminUserOptions() {
      if (!(adminImpersonateSelect instanceof HTMLSelectElement)) {
        return;
      }

      userStore = loadUserStore();
      const usersObj =
        userStore && userStore.users && typeof userStore.users === 'object'
          ? userStore.users
          : {};

      const usernames = Object.keys(usersObj).sort((a, b) =>
        a.localeCompare(b, 'fr', { sensitivity: 'base' }),
      );

      adminImpersonateSelect.innerHTML = '';

      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = 'Sélectionnez un utilisateur';
      placeholder.disabled = true;
      placeholder.selected = true;
      adminImpersonateSelect.appendChild(placeholder);

      usernames.forEach((username) => {
        const details = usersObj[username];
        const email = details && typeof details.email === 'string' ? details.email : '';
        const option = document.createElement('option');
        option.value = username;
        option.textContent = email ? `${username} (${email})` : username;
        adminImpersonateSelect.appendChild(option);
      });
    }

    function renderMetrics() {
      const metrics = data && typeof data === 'object' && data.metrics ? data.metrics : {};
      const contacts = Array.isArray(data.contacts) ? data.contacts : [];
      const totalContacts = contacts.length;
      const categoriesById = buildCategoryMap();
      const coverage = computeContactCoverage(contacts, categoriesById);
      const contactsWithChannel = coverage.both + coverage.phoneOnly + coverage.emailOnly;

      if (coverageHighlightEl) {
        const highlightRatio = totalContacts > 0 ? contactsWithChannel / totalContacts : 0;
        coverageHighlightEl.textContent = totalContacts > 0
          ? percentFormatter.format(highlightRatio)
          : '0 %';
      }

      if (coverageSummaryEl) {
        let summaryText = '';
        if (totalContacts === 0) {
          summaryText = "Aucun contact n'a encore été enregistré.";
        } else if (contactsWithChannel === 0) {
          summaryText = "Aucun moyen de contact n'a encore été détecté.";
        } else if (contactsWithChannel === totalContacts) {
          summaryText = 'Tous vos contacts disposent d\'au moins un canal de communication.';
        } else {
          const label = contactsWithChannel > 1 ? 'contacts' : 'contact';
          const verb = contactsWithChannel > 1 ? 'disposent' : 'dispose';
          summaryText = `${numberFormatter.format(contactsWithChannel)} ${label} ${verb} d\'au moins un canal.`;
        }
        coverageSummaryEl.textContent = summaryText;
      }

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

      renderHomeDirectoryLimit();

      refreshKeywordStatsIfOpen();
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

    function enforceRequiredContactCategories(target) {
      if (!target || typeof target !== 'object') {
        return;
      }

      if (!Array.isArray(target.categories)) {
        target.categories = [];
      }

      const idChanges = [];
      const normalizeName = (name) => normalizeLabel(name || '');

      REQUIRED_CONTACT_CATEGORIES.forEach((definition) => {
        let category = target.categories.find((item) => item && item.id === definition.id);
        if (!category) {
          const fallbackIndex = target.categories.findIndex(
            (item) => item && normalizeName(item.name) === normalizeName(definition.name),
          );
          if (fallbackIndex !== -1) {
            category = target.categories[fallbackIndex];
            const previousId = category.id;
            if (previousId && previousId !== definition.id) {
              idChanges.push({ oldId: previousId, newId: definition.id });
            }
            category.id = definition.id;
          } else {
            category = {
              id: definition.id,
              name: definition.name,
              description: '',
              type: definition.type,
              options: [],
              order: 0,
            };
            target.categories.push(category);
          }
        }
        category.name = definition.name;
        category.type = definition.type;
        category.options = [];
      });

      const requiredIds = new Set(REQUIRED_CONTACT_CATEGORIES.map((item) => item.id));
      const orderedRequired = REQUIRED_CONTACT_CATEGORIES.map((definition) =>
        target.categories.find((item) => item && item.id === definition.id),
      ).filter(Boolean);
      orderedRequired.forEach((category, index) => {
        category.order = index;
      });

      const remainingCategories = target.categories.filter(
        (category) => category && !requiredIds.has(category.id),
      );
      const sortedRemaining = sortCategoriesForDisplay(remainingCategories);
      sortedRemaining.forEach((category, index) => {
        category.order = REQUIRED_CONTACT_CATEGORIES.length + index;
      });

      target.categories = [...orderedRequired, ...sortedRemaining];

      if (!Array.isArray(target.contacts) || idChanges.length === 0) {
        return;
      }

      target.contacts.forEach((contact) => {
        if (!contact || typeof contact !== 'object') {
          return;
        }
        if (!contact.categoryValues || typeof contact.categoryValues !== 'object') {
          contact.categoryValues = {};
        }
        idChanges.forEach(({ oldId, newId }) => {
          if (!oldId || oldId === newId) {
            return;
          }
          if (
            contact.categoryValues[oldId] !== undefined &&
            contact.categoryValues[newId] === undefined
          ) {
            contact.categoryValues[newId] = contact.categoryValues[oldId];
          }
          delete contact.categoryValues[oldId];
        });
      });
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

    function escapeHtml(value) {
      if (value === undefined || value === null) {
        return '';
      }

      const text = value.toString();
      const escapeMap = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;',
      };

      return text.replace(/[&<>'"]/g, (char) => escapeMap[char] || char);
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

  function toTrimmedString(value) {
    if (value === undefined || value === null) {
      return '';
    }
    return value.toString().trim();
  }

  const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

  function normalizeIdentifierNameSegment(value) {
    const stringValue = toTrimmedString(value);
    if (!stringValue) {
      return '';
    }
    return stringValue
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .toLowerCase();
  }

  function parseDatePartsFromValue(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return {
        day: String(value.getUTCDate()).padStart(2, '0'),
        month: String(value.getUTCMonth() + 1).padStart(2, '0'),
        year: String(value.getUTCFullYear()).padStart(4, '0'),
      };
    }

    if (typeof value === 'number' && Number.isFinite(value)) {
      const excelDate = new Date(EXCEL_EPOCH_UTC + Math.floor(value) * 86400000);
      if (!Number.isNaN(excelDate.getTime())) {
        return {
          day: String(excelDate.getUTCDate()).padStart(2, '0'),
          month: String(excelDate.getUTCMonth() + 1).padStart(2, '0'),
          year: String(excelDate.getUTCFullYear()).padStart(4, '0'),
        };
      }
    }

    const stringValue = toTrimmedString(value);
    if (!stringValue) {
      return null;
    }

    const normalized = stringValue.replace(/[.]/g, '/');

    const identifierMatch = normalized.match(/^(\d{2})_(\d{2})_(\d{4})$/);
    if (identifierMatch) {
      return {
        day: identifierMatch[1],
        month: identifierMatch[2],
        year: identifierMatch[3],
      };
    }

    const isoMatch = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (isoMatch) {
      return {
        year: isoMatch[1],
        month: isoMatch[2].padStart(2, '0'),
        day: isoMatch[3].padStart(2, '0'),
      };
    }

    const frMatch = normalized.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (frMatch) {
      const day = frMatch[1].padStart(2, '0');
      const month = frMatch[2].padStart(2, '0');
      let yearNumber = Number(frMatch[3]);
      if (!Number.isFinite(yearNumber)) {
        return null;
      }
      if (frMatch[3].length === 2) {
        yearNumber += yearNumber >= 50 ? 1900 : 2000;
      }
      const year = yearNumber.toString().padStart(4, '0');
      return { day, month, year };
    }

    const compactMatch = normalized.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (compactMatch) {
      return {
        year: compactMatch[1],
        month: compactMatch[2],
        day: compactMatch[3],
      };
    }

    const parsed = new Date(normalized);
    if (!Number.isNaN(parsed.getTime())) {
      return {
        day: String(parsed.getUTCDate()).padStart(2, '0'),
        month: String(parsed.getUTCMonth() + 1).padStart(2, '0'),
        year: String(parsed.getUTCFullYear()).padStart(4, '0'),
      };
    }

    return null;
  }

  function buildContactIdentifierData(firstName, lastName, birthDate) {
    const normalizedFirst = normalizeIdentifierNameSegment(firstName);
    const normalizedLast = normalizeIdentifierNameSegment(lastName);
    const dateParts = parseDatePartsFromValue(birthDate);
    if (!normalizedFirst || !normalizedLast || !dateParts) {
      return null;
    }
    return {
      identifier: `${normalizedFirst}_${normalizedLast}_${dateParts.day}_${dateParts.month}_${dateParts.year}`,
      normalizedDate: `${dateParts.year}-${dateParts.month}-${dateParts.day}`,
    };
  }

  function ensureContactIdentifier(categoryValues, fallback = {}) {
    if (!categoryValues || typeof categoryValues !== 'object') {
      return null;
    }

    const firstNameValue = toTrimmedString(
      categoryValues[REQUIRED_CATEGORY_IDS.firstName] ?? fallback.firstName,
    );
    const lastNameValue = toTrimmedString(
      categoryValues[REQUIRED_CATEGORY_IDS.lastName] ?? fallback.lastName,
    );
    const rawBirthDate =
      categoryValues[REQUIRED_CATEGORY_IDS.birthDate] ?? fallback.birthDate ?? '';
    const birthDateValue =
      typeof rawBirthDate === 'number' || rawBirthDate instanceof Date
        ? rawBirthDate
        : toTrimmedString(rawBirthDate);

    const identifierData = buildContactIdentifierData(
      firstNameValue,
      lastNameValue,
      birthDateValue,
    );

    if (!identifierData) {
      return null;
    }

    categoryValues[REQUIRED_CATEGORY_IDS.firstName] = firstNameValue;
    categoryValues[REQUIRED_CATEGORY_IDS.lastName] = lastNameValue;
    categoryValues[REQUIRED_CATEGORY_IDS.birthDate] = identifierData.normalizedDate;
    categoryValues[REQUIRED_CATEGORY_IDS.identifier] = identifierData.identifier;

    return identifierData;
  }

  function formatDateForInput(value) {
    const parts = parseDatePartsFromValue(value);
    if (!parts) {
      return '';
    }
    return `${parts.year}-${parts.month}-${parts.day}`;
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
        renderSavedSearchOptions();
        renderSavedSearchList();
        return;
      }

      categoryList.innerHTML = '';

      const categories = sortCategoriesForDisplay();

      if (categories.length === 0) {
        categoryEmptyState.hidden = false;
        renderContactCategoryFields();
        renderSearchCategoryFields();
        renderContacts();
        renderSavedSearchOptions();
        renderSavedSearchList();
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
        const isProtectedCategory = REQUIRED_CATEGORY_ID_SET.has(category.id);

        if (dragHandle instanceof HTMLElement) {
          if (isProtectedCategory) {
            dragHandle.removeAttribute('aria-label');
            dragHandle.removeAttribute('title');
            dragHandle.removeAttribute('draggable');
            dragHandle.draggable = false;
            dragHandle.setAttribute('aria-hidden', 'true');
          } else {
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
          if (isProtectedCategory) {
            editButton.remove();
          } else {
            editButton.addEventListener('click', () => {
              startCategoryEdition(category.id);
            });
          }
        }

        if (deleteButton) {
          if (isProtectedCategory) {
            deleteButton.remove();
          } else {
            deleteButton.addEventListener('click', () => {
              deleteCategory(category.id);
            });
          }
        }

        fragment.appendChild(listItem);
      });

      categoryList.appendChild(fragment);
      renderContactCategoryFields();
      renderSearchCategoryFields();
      renderContacts();
      renderSavedSearchOptions();
      renderSavedSearchList();
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
        renderBulkKeywordOptions();
        renderContacts();
        renderSavedSearchOptions();
        renderSavedSearchList();
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
        renderBulkKeywordOptions();
        renderContacts();
        renderSavedSearchOptions();
        renderSavedSearchList();
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
        const statsButton = listItem.querySelector('[data-action="stats"]');

        if (titleEl) {
          titleEl.textContent = keyword.name;
        }

        if (descriptionEl) {
          descriptionEl.textContent = keyword.description || 'Aucune description renseignée.';
          descriptionEl.classList.toggle('keyword-description--empty', !keyword.description);
        }

        if (statsButton) {
          statsButton.addEventListener('click', () => {
            openKeywordStats(keyword.id);
          });
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
      renderBulkKeywordOptions();
      renderContacts();
      renderSavedSearchOptions();
      renderSavedSearchList();
      refreshKeywordStatsIfOpen();
    }

    function openKeywordStats(keywordId) {
      if (!keywordId) {
        return;
      }

      if (!updateKeywordStatsView(keywordId)) {
        return;
      }

      keywordStatsActiveId = keywordId;
      keywordStatsPreviousFocus =
        document.activeElement instanceof HTMLElement ? document.activeElement : null;

      if (keywordStatsOverlay) {
        keywordStatsOverlay.hidden = false;
        keywordStatsOverlay.setAttribute('aria-hidden', 'false');
      }

      if (keywordStatsDialog) {
        keywordStatsDialog.setAttribute('tabindex', '-1');
      }

      const focusTarget =
        keywordStatsCloseButtons.length > 0 ? keywordStatsCloseButtons[0] : keywordStatsDialog;
      if (focusTarget && typeof focusTarget.focus === 'function') {
        focusTarget.focus();
      }
    }

    function closeKeywordStats() {
      if (!keywordStatsOverlay) {
        return;
      }

      if (keywordStatsOverlay.hidden) {
        return;
      }

      keywordStatsOverlay.hidden = true;
      keywordStatsOverlay.setAttribute('aria-hidden', 'true');
      keywordStatsActiveId = '';

      if (keywordStatsPreviousFocus && typeof keywordStatsPreviousFocus.focus === 'function') {
        keywordStatsPreviousFocus.focus();
      }
      keywordStatsPreviousFocus = null;
    }

    function updateKeywordStatsView(keywordId) {
      if (!keywordId) {
        return false;
      }

      const keywords = Array.isArray(data.keywords) ? data.keywords : [];
      const keyword = keywords.find((item) => item && item.id === keywordId);
      if (!keyword) {
        return false;
      }

      const contacts = Array.isArray(data.contacts) ? data.contacts : [];
      let assignedCount = 0;
      contacts.forEach((contact) => {
        const assigned = Array.isArray(contact.keywords) ? contact.keywords : [];
        if (assigned.includes(keywordId)) {
          assignedCount += 1;
        }
      });

      const total = contacts.length;
      const remaining = Math.max(0, total - assignedCount);
      const assignedRatio = total > 0 ? assignedCount / total : 0;
      const remainingRatio = total > 0 ? remaining / total : 0;

      const keywordName = keyword.name || 'Mot clé';

      if (keywordStatsTitle) {
        keywordStatsTitle.textContent = `Statistiques — ${keywordName}`;
      }

      if (keywordStatsPercentEl) {
        keywordStatsPercentEl.textContent = total > 0
          ? percentFormatter.format(assignedRatio)
          : '0 %';
      }

      if (keywordStatsAssignedCountEl) {
        keywordStatsAssignedCountEl.textContent = numberFormatter.format(assignedCount);
      }

      if (keywordStatsAssignedPercentEl) {
        keywordStatsAssignedPercentEl.textContent = total > 0
          ? percentFormatter.format(assignedRatio)
          : '0 %';
      }

      if (keywordStatsRemainingCountEl) {
        keywordStatsRemainingCountEl.textContent = numberFormatter.format(remaining);
      }

      if (keywordStatsRemainingPercentEl) {
        keywordStatsRemainingPercentEl.textContent = total > 0
          ? percentFormatter.format(remainingRatio)
          : '0 %';
      }

      if (keywordStatsTotalEl) {
        keywordStatsTotalEl.textContent = numberFormatter.format(total);
      }

      if (keywordStatsSummaryEl) {
        let summaryText = '';
        if (total === 0) {
          summaryText = "Aucun contact n'a encore été enregistré.";
        } else if (assignedCount === 0) {
          summaryText = "Ce mot clé n'est associé à aucun contact pour le moment.";
        } else {
          summaryText = `${formatContactCount(assignedCount)} sur ${formatContactCount(total)} comportent ce mot clé.`;
        }
        keywordStatsSummaryEl.textContent = summaryText;
      }

      if (keywordStatsChart) {
        if (total > 0) {
          const angle = Math.max(0, Math.min(360, assignedRatio * 360));
          keywordStatsChart.style.background = `conic-gradient(var(--color-primary) 0deg ${angle}deg, rgba(148, 163, 184, 0.35) ${angle}deg 360deg)`;
          keywordStatsChart.classList.remove('keyword-stats-chart--empty');
          keywordStatsChart.setAttribute(
            'aria-label',
            `${numberFormatter.format(assignedCount)} destinataire(s) sur ${numberFormatter.format(total)} portent le mot clé ${keywordName}.`,
          );
        } else {
          keywordStatsChart.style.background = 'var(--coverage-none)';
          keywordStatsChart.classList.add('keyword-stats-chart--empty');
          keywordStatsChart.setAttribute(
            'aria-label',
            "Répartition du mot clé indisponible : aucun contact enregistré.",
          );
        }
      }

      return true;
    }

    function refreshKeywordStatsIfOpen() {
      if (!keywordStatsActiveId) {
        return;
      }

      if (keywordStatsOverlay && keywordStatsOverlay.hidden) {
        return;
      }

      updateKeywordStatsView(keywordStatsActiveId);
    }

    function handleKeywordStatsKeydown(event) {
      if (event.key !== 'Escape') {
        return;
      }

      if (keywordStatsOverlay && keywordStatsOverlay.hidden) {
        return;
      }

      event.preventDefault();
      closeKeywordStats();
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

        if (category.id === REQUIRED_CATEGORY_IDS.firstName || category.id === REQUIRED_CATEGORY_IDS.lastName) {
          if (input instanceof HTMLInputElement) {
            input.required = true;
          }
        }

        if (category.id === REQUIRED_CATEGORY_IDS.birthDate) {
          if (input instanceof HTMLInputElement) {
            input.required = true;
          }
        }

        if (category.id === REQUIRED_CATEGORY_IDS.identifier) {
          if (input instanceof HTMLInputElement) {
            input.readOnly = true;
            input.placeholder = input.placeholder || 'Identifiant généré automatiquement';
          }
        }

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
          const resolvedValue =
            baseType === 'date' && input instanceof HTMLInputElement
              ? formatDateForInput(initialValue)
              : initialValue;
          input.value = resolvedValue;
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
      if (!searchKeywordsContainer) {
        if (Array.isArray(advancedFilters.keywords) && advancedFilters.keywords.length > 0) {
          advancedFilters = { ...advancedFilters, keywords: [] };
        }
        if (advancedFilters.keywordMode !== KEYWORD_FILTER_MODE_ALL) {
          advancedFilters = { ...advancedFilters, keywordMode: KEYWORD_FILTER_MODE_ALL };
        }
        return;
      }

      const keywords = Array.isArray(data.keywords)
        ? data.keywords
            .slice()
            .sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
        : [];

      const previousSelection = Array.isArray(advancedFilters.keywords)
        ? advancedFilters.keywords
        : [];
      const selectedValues = new Set(previousSelection);

      searchKeywordsContainer.innerHTML = '';
      searchKeywordsContainer.classList.toggle('checkbox-grid', keywords.length > 0);

      if (keywords.length === 0) {
        const message = document.createElement('p');
        message.className = 'advanced-keywords-empty';
        message.textContent =
          'Aucun mot clé disponible pour le moment. Ajoutez vos mots clés dans l’onglet « Mots clés ».';
        searchKeywordsContainer.appendChild(message);
      } else {
        const fragment = document.createDocumentFragment();
        keywords.forEach((keyword) => {
          if (!keyword || typeof keyword !== 'object' || !keyword.id) {
            return;
          }

          const keywordId = keyword.id.toString();
          if (!keywordId) {
            return;
          }

          const checkboxId = `search-keyword-${keywordId}`;
          const wrapper = document.createElement('div');
          wrapper.className = 'checkbox-item';

          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.id = checkboxId;
          checkbox.name = 'search-keywords';
          checkbox.value = keywordId;
          checkbox.checked = selectedValues.has(keywordId);

          const label = document.createElement('label');
          label.setAttribute('for', checkboxId);
          label.textContent = keyword.name || 'Mot clé sans nom';

          wrapper.append(checkbox, label);
          fragment.appendChild(wrapper);
        });
        searchKeywordsContainer.appendChild(fragment);
      }

      const allowedValues = new Set(
        Array.isArray(keywords)
          ? keywords
            .map((keyword) => (keyword && keyword.id ? keyword.id.toString() : ''))
            .filter((value) => Boolean(value))
          : []
      );
      const filteredSelection = previousSelection.filter((value) => allowedValues.has(value));

      if (filteredSelection.length !== previousSelection.length) {
        advancedFilters = { ...advancedFilters, keywords: filteredSelection };
      }

      updateSearchKeywordModeUI(keywords.length);
    }

    function updateSearchKeywordModeUI(keywordCount) {
      const shouldForceAll = typeof keywordCount === 'number' && keywordCount === 0;
      let normalizedMode = KEYWORD_FILTER_MODE_ALL;
      if (!shouldForceAll && advancedFilters.keywordMode === KEYWORD_FILTER_MODE_ANY) {
        normalizedMode = KEYWORD_FILTER_MODE_ANY;
      }

      let matched = false;
      if (!searchKeywordModeInputs || searchKeywordModeInputs.length === 0) {
        return; // rien à traiter
      }

      searchKeywordModeInputs.forEach((input) => {
        if (!(input instanceof HTMLInputElement)) {
          return;
        }

        input.disabled = shouldForceAll;

        if (input.value === normalizedMode) {
          input.checked = true;
          matched = true;
        } else if (input.checked) {
          input.checked = false;
        }
      });


      if (!matched) {
        const fallback = searchKeywordModeInputs.find(
          (input) => input instanceof HTMLInputElement && input.value === KEYWORD_FILTER_MODE_ALL,
        );
        if (fallback) {
          fallback.checked = true;
        }
      }

      if (normalizedMode !== advancedFilters.keywordMode) {
        advancedFilters = { ...(advancedFilters || {}), keywordMode: normalizedMode };
      }
    }

    function handleSaveCurrentSearch() {
      if (!Array.isArray(data.savedSearches)) {
        data.savedSearches = [];
      }

      const activeSavedSearch = activeSavedSearchId
        ? data.savedSearches.find((search) => search && search.id === activeSavedSearchId)
        : null;
      let suggestedName = activeSavedSearch && activeSavedSearch.name ? activeSavedSearch.name : '';

      if (!suggestedName) {
        const trimmedTerm = (contactSearchTerm || '').toString().trim();
        if (trimmedTerm) {
          suggestedName = trimmedTerm;
        }
      }

      if (!suggestedName) {
        suggestedName = 'Nouvelle recherche';
      }

      const enteredName = window.prompt('Nom de la recherche sauvegardée', suggestedName);
      if (enteredName === null) {
        return;
      }

      const normalizedName = enteredName.trim();
      if (!normalizedName) {
        if (contactSaveSearchFeedback) {
          contactSaveSearchFeedback.textContent = 'Veuillez saisir un nom pour enregistrer la recherche.';
        }
        return;
      }

      const now = new Date().toISOString();
      const searchTerm = (contactSearchTerm || '').toString().trim();
      const filtersSnapshot = cloneAdvancedFilters(advancedFilters);
      let savedSearch = activeSavedSearch || null;

      if (!savedSearch) {
        const lowerName = normalizedName.toLowerCase();
        savedSearch = data.savedSearches.find(
          (item) => item && typeof item.name === 'string' && item.name.toLowerCase() === lowerName,
        ) || null;
      }

      let feedbackMessage = '';
      let eventType = 'updated';

      if (savedSearch) {
        savedSearch.name = normalizedName;
        savedSearch.searchTerm = searchTerm;
        savedSearch.advancedFilters = filtersSnapshot;
        if (!savedSearch.createdAt) {
          savedSearch.createdAt = now;
        }
        savedSearch.updatedAt = now;
        setActiveSavedSearchId(savedSearch.id || '');
        feedbackMessage = `Recherche « ${normalizedName} » mise à jour.`;
      } else {
        const newSavedSearch = {
          id: generateId('saved-search'),
          name: normalizedName,
          searchTerm,
          advancedFilters: filtersSnapshot,
          createdAt: now,
          updatedAt: now,
        };
        data.savedSearches.push(newSavedSearch);
        setActiveSavedSearchId(newSavedSearch.id);
        feedbackMessage = `Recherche « ${normalizedName} » enregistrée.`;
        eventType = 'created';
        savedSearch = newSavedSearch;
      }

      sortSavedSearchesInPlace();
      data.lastUpdated = now;
      saveDataForUser(currentUser, data);
      renderSavedSearchOptions();
      renderSavedSearchList();
      if (contactSaveSearchFeedback) {
        contactSaveSearchFeedback.textContent = feedbackMessage;
      }
      notifyDataChanged('saved-searches', {
        type: eventType,
        savedSearchId: savedSearch ? savedSearch.id : '',
      });
    }

    function applySavedSearchById(savedSearchId, options = {}) {
      if (!savedSearchId || !Array.isArray(data.savedSearches)) {
        return false;
      }

      const savedSearch = data.savedSearches.find((item) => item && item.id === savedSearchId);
      if (!savedSearch) {
        return false;
      }

      return applySavedSearch(savedSearch, options);
    }

    function applySavedSearch(savedSearch, options = {}) {
      if (!savedSearch || typeof savedSearch !== 'object') {
        return false;
      }

      const {
        navigateToSearchPage = false,
        focusSearchInput = false,
        announce = false,
      } = options;

      const clonedFilters = cloneAdvancedFilters(savedSearch.advancedFilters);
      advancedFilters = clonedFilters;

      const searchTerm = typeof savedSearch.searchTerm === 'string' ? savedSearch.searchTerm.trim() : '';
      contactSearchTerm = searchTerm;

      if (navigateToSearchPage) {
        showPage('contacts-search');
      }

      if (contactSearchInput) {
        contactSearchInput.value = searchTerm;
        if (focusSearchInput && typeof contactSearchInput.focus === 'function') {
          contactSearchInput.focus();
        }
      }

      renderSearchCategoryFields();
      renderSearchKeywordOptions();
      contactCurrentPage = 1;
      renderContacts();
      setActiveSavedSearchId(savedSearch.id || '');

      if (announce) {
        if (contactSaveSearchFeedback) {
          contactSaveSearchFeedback.textContent = `Recherche « ${savedSearch.name || 'sans nom'} » appliquée.`;
        }
      } else {
        clearContactSaveSearchFeedback();
      }

      return true;
    }

    function deleteSavedSearch(savedSearchId) {
      if (!savedSearchId || !Array.isArray(data.savedSearches)) {
        return;
      }

      const index = data.savedSearches.findIndex((item) => item && item.id === savedSearchId);
      if (index === -1) {
        return;
      }

      const target = data.savedSearches[index];
      const confirmationMessage = target && target.name
        ? `Supprimer la recherche « ${target.name} » ?`
        : 'Supprimer cette recherche sauvegardée ?';

      if (!window.confirm(confirmationMessage)) {
        return;
      }

      data.savedSearches.splice(index, 1);
      const now = new Date().toISOString();
      sortSavedSearchesInPlace();
      data.lastUpdated = now;
      saveDataForUser(currentUser, data);

      if (activeSavedSearchId === savedSearchId) {
        setActiveSavedSearchId('');
      }

      renderSavedSearchOptions();
      renderSavedSearchList();
      if (contactSaveSearchFeedback) {
        contactSaveSearchFeedback.textContent = 'La recherche sauvegardée a été supprimée.';
      }

      notifyDataChanged('saved-searches', { type: 'deleted', savedSearchId });
    }

    function renderSavedSearchOptions() {
      const savedSearches = Array.isArray(data.savedSearches) ? data.savedSearches.slice() : [];
      savedSearches.sort((a, b) => {
        const diff = getSavedSearchTimestamp(b) - getSavedSearchTimestamp(a);
        if (diff !== 0) {
          return diff;
        }
        const nameA = a && typeof a.name === 'string' ? a.name : '';
        const nameB = b && typeof b.name === 'string' ? b.name : '';
        return nameA.localeCompare(nameB, 'fr', { sensitivity: 'base' });
      });

      if (contactSavedSearchSelect instanceof HTMLSelectElement) {
        const previousValue = contactSavedSearchSelect.value;
        contactSavedSearchSelect.innerHTML = '';

        const placeholder = document.createElement('option');
        placeholder.value = '';
        placeholder.textContent =
          savedSearches.length === 0
            ? 'Aucune recherche disponible'
            : 'Aucune recherche sélectionnée';
        contactSavedSearchSelect.appendChild(placeholder);

        savedSearches.forEach((savedSearch) => {
          const option = document.createElement('option');
          option.value = savedSearch.id || '';
          option.textContent = savedSearch.name || 'Recherche sans nom';
          contactSavedSearchSelect.appendChild(option);
        });

        contactSavedSearchSelect.disabled = savedSearches.length === 0;

        if (!activeSavedSearchId && previousValue && previousValue !== '') {
          const stillAvailable = savedSearches.some((search) => search.id === previousValue);
          if (stillAvailable) {
            contactSavedSearchSelect.value = previousValue;
          }
        }

        updateSavedSearchSelectionUI();
      }

      if (campaignSavedSearchSelect instanceof HTMLSelectElement) {
        const previousValue = campaignSavedSearchSelect.value;
        campaignSavedSearchSelect.innerHTML = '';

        const placeholder = document.createElement('option');
        placeholder.value = '';
        const hasChoices = savedSearches.length > 0;
        placeholder.textContent = hasChoices
          ? 'Sélectionnez une recherche'
          : 'Aucune recherche disponible';
        placeholder.disabled = hasChoices;
        placeholder.selected = true;
        campaignSavedSearchSelect.appendChild(placeholder);

        savedSearches.forEach((savedSearch) => {
          const option = document.createElement('option');
          option.value = savedSearch.id || '';
          option.textContent = savedSearch.name || 'Recherche sans nom';
          campaignSavedSearchSelect.appendChild(option);
        });

        campaignSavedSearchSelect.disabled = !hasChoices;

        if (hasChoices) {
          const stillAvailable = savedSearches.some((search) => search.id === previousValue);
          if (stillAvailable) {
            campaignSavedSearchSelect.value = previousValue;
          }
        }
      }
      
      updateCampaignRecipients();
    }

    function renderSavedSearchList() {
      const savedSearches = Array.isArray(data.savedSearches) ? data.savedSearches.slice() : [];
      savedSearches.sort((a, b) => {
        const diff = getSavedSearchTimestamp(b) - getSavedSearchTimestamp(a);
        if (diff !== 0) {
          return diff;
        }
        const nameA = a && typeof a.name === 'string' ? a.name : '';
        const nameB = b && typeof b.name === 'string' ? b.name : '';
        return nameA.localeCompare(nameB, 'fr', { sensitivity: 'base' });
      });

      if (savedSearchCountEl) {
        savedSearchCountEl.textContent = numberFormatter.format(savedSearches.length);
      }

      if (!savedSearchList) {
        if (savedSearchEmptyState) {
          savedSearchEmptyState.hidden = savedSearches.length > 0;
        }
        return;
      }

      savedSearchList.innerHTML = '';

      if (savedSearches.length === 0) {
        if (savedSearchEmptyState) {
          savedSearchEmptyState.hidden = false;
        }
        return;
      }

      if (savedSearchEmptyState) {
        savedSearchEmptyState.hidden = true;
      }

      const fragment = document.createDocumentFragment();
      savedSearches.forEach((savedSearch) => {
        const listItem = document.createElement('li');
        listItem.className = 'saved-search-item';
        if (savedSearch.id) {
          listItem.dataset.savedSearchId = savedSearch.id;
        }

        const header = document.createElement('div');
        header.className = 'saved-search-header';

        const title = document.createElement('h3');
        title.className = 'saved-search-name';
        title.textContent = savedSearch.name || 'Recherche sans nom';
        header.appendChild(title);

        const metaText = buildSavedSearchMeta(savedSearch);
        if (metaText) {
          const meta = document.createElement('p');
          meta.className = 'saved-search-meta';
          meta.textContent = metaText;
          header.appendChild(meta);
        }

        listItem.appendChild(header);

        const summaryText = formatSavedSearchSummary(savedSearch);
        if (summaryText) {
          const summary = document.createElement('p');
          summary.className = 'saved-search-summary';
          summary.textContent = summaryText;
          listItem.appendChild(summary);
        }

        const actions = document.createElement('div');
        actions.className = 'saved-search-actions';

        const applyButton = document.createElement('button');
        applyButton.type = 'button';
        applyButton.className = 'secondary-button';
        applyButton.dataset.action = 'apply-saved-search';
        applyButton.dataset.savedSearchId = savedSearch.id || '';
        applyButton.textContent = 'Appliquer dans la recherche';
        actions.appendChild(applyButton);

        const deleteButton = document.createElement('button');
        deleteButton.type = 'button';
        deleteButton.className = 'secondary-button saved-search-delete';
        deleteButton.dataset.action = 'delete-saved-search';
        deleteButton.dataset.savedSearchId = savedSearch.id || '';
        deleteButton.textContent = 'Supprimer';
        actions.appendChild(deleteButton);

        listItem.appendChild(actions);
        fragment.appendChild(listItem);
      });

      savedSearchList.appendChild(fragment);
    }

    function buildSavedSearchMeta(savedSearch) {
      if (!savedSearch || typeof savedSearch !== 'object') {
        return '';
      }

      const parts = [];
      const createdLabel = formatSavedSearchDate(savedSearch.createdAt);
      if (createdLabel) {
        parts.push(`Créée le ${createdLabel}`);
      }

      const updatedLabel = formatSavedSearchDate(savedSearch.updatedAt);
      if (updatedLabel && updatedLabel !== createdLabel) {
        parts.push(`Modifiée le ${updatedLabel}`);
      }

      return parts.join(' · ');
    }

    function formatSavedSearchDate(value) {
      if (typeof value !== 'string' || !value) {
        return '';
      }
      const parsed = Date.parse(value);
      if (Number.isNaN(parsed)) {
        return '';
      }
      return savedSearchDateFormatter.format(new Date(parsed));
    }

    function formatSavedSearchSummary(savedSearch) {
      if (!savedSearch || typeof savedSearch !== 'object') {
        return '';
      }

      const parts = [];
      const searchTerm = typeof savedSearch.searchTerm === 'string' ? savedSearch.searchTerm.trim() : '';
      if (searchTerm) {
        parts.push(`Terme libre : « ${searchTerm} »`);
      }

      const categoriesSummary = [];
      const categoriesById = buildCategoryMap();
      const categoryFilters =
        savedSearch.advancedFilters && savedSearch.advancedFilters.categories
          ? savedSearch.advancedFilters.categories
          : {};

      Object.entries(categoryFilters).forEach(([categoryId, filter]) => {
        if (!filter || typeof filter !== 'object') {
          return;
        }

        const category = categoriesById.get(categoryId);
        const label = category ? category.name : 'Catégorie supprimée';
        const rawValue = filter.rawValue != null ? filter.rawValue.toString() : '';
        if (!rawValue) {
          return;
        }
        categoriesSummary.push(`${label} = ${rawValue}`);
      });

      if (categoriesSummary.length > 0) {
        parts.push(`Catégories : ${categoriesSummary.join(', ')}`);
      }

      const keywordIds =
        savedSearch.advancedFilters && Array.isArray(savedSearch.advancedFilters.keywords)
          ? savedSearch.advancedFilters.keywords
          : [];
      if (keywordIds.length > 0) {
        const keywordsById = new Map(
          Array.isArray(data.keywords) ? data.keywords.map((keyword) => [keyword.id, keyword]) : [],
        );
        const keywordNames = keywordIds
          .map((keywordId) => {
            const keyword = keywordsById.get(keywordId);
            return keyword ? keyword.name : 'Mot clé supprimé';
          })
          .filter((value) => Boolean(value));

        if (keywordNames.length > 0) {
          const keywordModeLabel =
          savedSearch?.advancedFilters?.keywordMode === KEYWORD_FILTER_MODE_ANY
            ? 'OU'
            : 'ET';
          parts.push(`Mots clés (${keywordModeLabel}) : ${keywordNames.join(', ')}`);
        }
      }

      if (parts.length === 0) {
        return 'Aucun filtre spécifique. Tous les contacts seront affichés.';
      }

      return parts.join(' · ');
    }

    function setActiveSavedSearchId(nextId) {
      const normalized = typeof nextId === 'string' ? nextId : '';
      if (activeSavedSearchId === normalized) {
        updateSavedSearchSelectionUI();
        return;
      }
      activeSavedSearchId = normalized;
      updateSavedSearchSelectionUI();
    }

    function updateSavedSearchSelectionUI() {
      if (!(contactSavedSearchSelect instanceof HTMLSelectElement)) {
        return;
      }

      if (!activeSavedSearchId) {
        contactSavedSearchSelect.value = '';
        return;
      }

      const hasOption = Array.from(contactSavedSearchSelect.options).some(
        (option) => option.value === activeSavedSearchId,
      );

      contactSavedSearchSelect.value = hasOption ? activeSavedSearchId : '';
    }

    function clearContactSaveSearchFeedback() {
      if (contactSaveSearchFeedback) {
        contactSaveSearchFeedback.textContent = '';
      }
    }

    const EMAIL_TEMPLATE_BLOCK_DEFAULTS = {
      paragraph: { text: '', align: 'left' },
      image: { url: '', alt: '', link: '', caption: '' },
      button: { label: '', url: '', variant: 'primary' },
    };
    const EMAIL_TEMPLATE_ALIGN_VALUES = new Set(['left', 'center', 'right']);
    const EMAIL_TEMPLATE_BUTTON_VARIANTS = new Set(['primary', 'outline']);

    function addEmailTemplateBlock(type, initialData = {}) {
      const block = createEmailTemplateBlock(type, initialData);
      emailTemplateDraftBlocks.push(block);
      renderEmailTemplateBlocks();
      renderEmailTemplatePreview();
      focusEmailTemplateBlock(block.id);
    }

    function createEmailTemplateBlock(type, initialData = {}, options = {}) {
      const normalizedType = type === 'image' || type === 'button' ? type : 'paragraph';
      const preserveId = Boolean(options && options.preserveId);
      const source =
        initialData &&
        typeof initialData === 'object' &&
        initialData.data &&
        typeof initialData.data === 'object'
          ? initialData.data
          : initialData;
      const data = normalizeEmailTemplateBlockData(normalizedType, source);
      let blockId = '';
      if (
        preserveId &&
        initialData &&
        typeof initialData.id === 'string' &&
        initialData.id.trim()
      ) {
        blockId = initialData.id.trim();
      } else if (options && typeof options.id === 'string' && options.id.trim()) {
        blockId = options.id.trim();
      } else {
        blockId = generateId('email-block');
      }
      
      return { id: blockId, type: normalizedType, data };

    }

    function normalizeEmailTemplateBlockData(type, rawData = {}) {
      const defaults =
        EMAIL_TEMPLATE_BLOCK_DEFAULTS[type] || EMAIL_TEMPLATE_BLOCK_DEFAULTS.paragraph;
      const data = { ...defaults };

      if (type === 'paragraph') {
        if (typeof rawData.text === 'string') {
          data.text = rawData.text;
        }
        if (
          typeof rawData.align === 'string' &&
          EMAIL_TEMPLATE_ALIGN_VALUES.has(rawData.align)
        ) {
          data.align = rawData.align;
        }
        if (!EMAIL_TEMPLATE_ALIGN_VALUES.has(data.align)) {
          data.align = 'left';
        }

      } else if (type === 'image') {
        if (typeof rawData.url === 'string') data.url = rawData.url.trim();
        if (typeof rawData.alt === 'string') data.alt = rawData.alt.trim();
        if (typeof rawData.link === 'string') data.link = rawData.link.trim();
        if (typeof rawData.caption === 'string') data.caption = rawData.caption.trim();

      } else if (type === 'button') {
        if (typeof rawData.label === 'string') data.label = rawData.label.trim();
        if (typeof rawData.url === 'string') data.url = rawData.url.trim();
        if (
          typeof rawData.variant === 'string' &&
          EMAIL_TEMPLATE_BUTTON_VARIANTS.has(rawData.variant)
        ) {
          data.variant = rawData.variant;
        }
        if (!EMAIL_TEMPLATE_BUTTON_VARIANTS.has(data.variant)) {
          data.variant = 'primary';
        }
      }

      return data;
    }


    function cloneEmailTemplateBlock(block, options = {}) {
      if (!block || typeof block !== 'object') {
        return createEmailTemplateBlock('paragraph', {});
      }
      return createEmailTemplateBlock(block.type, block, options);
    }

    function renderEmailTemplateBlocks() {
      if (!emailTemplateBlocksContainer) {
        return;
      }
      emailTemplateBlocksContainer.innerHTML = '';

      if (!Array.isArray(emailTemplateDraftBlocks) || emailTemplateDraftBlocks.length === 0) {
        const emptyState = document.createElement('p');
        emptyState.className = 'empty-state';
        emptyState.textContent = emailTemplateBlocksEmptyText;
        emailTemplateBlocksContainer.appendChild(emptyState);
        return;
      }
      const fragment = document.createDocumentFragment();
      emailTemplateDraftBlocks.forEach((block, index) => {
        const element = createEmailTemplateBlockElement(block, index);
        fragment.appendChild(element);
      });
      emailTemplateBlocksContainer.appendChild(fragment);
    }

    function createEmailTemplateBlockElement(block, index) {
      const wrapper = document.createElement('div');
      wrapper.className = 'email-template-block';
      wrapper.dataset.blockId = block.id;
      wrapper.dataset.blockType = block.type;

      const header = document.createElement('div');
      header.className = 'email-template-block-header';
      const title = document.createElement('h3');
      title.textContent = `Bloc ${index + 1} · ${getEmailBlockLabel(block.type)}`;
      header.appendChild(title);

      const controls = document.createElement('div');
      controls.className = 'email-template-block-controls';
      controls.appendChild(createBlockControlButton('Monter', 'move-up', index === 0));
      controls.appendChild(
        createBlockControlButton('Descendre', 'move-down', index === emailTemplateDraftBlocks.length - 1),
      );
      controls.appendChild(createBlockControlButton('Dupliquer', 'duplicate'));
      controls.appendChild(createBlockControlButton('Supprimer', 'remove', false, true));
      header.appendChild(controls);
      wrapper.appendChild(header);

      if (block.type === 'paragraph') {
        wrapper.appendChild(
          createTextareaRow(`${block.id}-text`, 'Contenu du paragraphe', block.data.text, 'text'),
        );

        const alignRow = document.createElement('div');
        alignRow.className = 'form-row';
        const alignLabel = document.createElement('label');
        alignLabel.setAttribute('for', `${block.id}-align`);
        alignLabel.textContent = 'Alignement';
        alignRow.appendChild(alignLabel);
        const alignSelect = document.createElement('select');
        alignSelect.id = `${block.id}-align`;
        alignSelect.dataset.blockField = 'align';
        EMAIL_TEMPLATE_ALIGN_VALUES.forEach((value) => {
          const option = document.createElement('option');
          option.value = value;
          option.textContent =
            value === 'left' ? 'Aligné à gauche' : value === 'center' ? 'Centré' : 'Aligné à droite';
          if (value === block.data.align) {
            option.selected = true;
          }
          alignSelect.appendChild(option);
        });
        alignRow.appendChild(alignSelect);
        wrapper.appendChild(alignRow);
        return wrapper;
      }

      if (block.type === 'image') {
        wrapper.appendChild(
          createInputRow(`${block.id}-url`, "URL de l'image *", block.data.url, 'url', 'url', true),
        );
        wrapper.appendChild(
          createInputRow(`${block.id}-alt`, 'Texte alternatif', block.data.alt, 'text', 'alt'),
        );
        wrapper.appendChild(
          createInputRow(`${block.id}-link`, "Lien sur l'image", block.data.link, 'url', 'link'),
        );
        wrapper.appendChild(
          createInputRow(`${block.id}-caption`, 'Légende', block.data.caption, 'text', 'caption'),
        );
        return wrapper;
      }

      wrapper.appendChild(
        createInputRow(`${block.id}-label`, 'Texte du bouton *', block.data.label, 'text', 'label', true),
      );
      wrapper.appendChild(
        createInputRow(`${block.id}-button-url`, 'Lien cible *', block.data.url, 'url', 'url', true),
      );
      const variantRow = document.createElement('div');
      variantRow.className = 'form-row';
      const variantLabel = document.createElement('label');
      variantLabel.setAttribute('for', `${block.id}-variant`);
      variantLabel.textContent = 'Style du bouton';
      variantRow.appendChild(variantLabel);
      const variantSelect = document.createElement('select');
      variantSelect.id = `${block.id}-variant`;
      variantSelect.dataset.blockField = 'variant';
      const variants = [
        { value: 'primary', label: 'Principal' },
        { value: 'outline', label: 'Contour' },
      ];
      variants.forEach((optionConfig) => {
        const option = document.createElement('option');
        option.value = optionConfig.value;
        option.textContent = optionConfig.label;
        if (optionConfig.value === block.data.variant) {
          option.selected = true;
        }
        variantSelect.appendChild(option);
      });
      variantRow.appendChild(variantSelect);
      wrapper.appendChild(variantRow);
      return wrapper;
    }

    function createBlockControlButton(label, action, disabled = false, danger = false) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = danger ? 'ghost-button ghost-button--danger' : 'ghost-button';
      button.dataset.action = action;
      button.textContent = label;
      button.disabled = disabled;
      return button;
    }

    function createTextareaRow(id, labelText, value, field) {
      const row = document.createElement('div');
      row.className = 'form-row';
      const label = document.createElement('label');
      label.setAttribute('for', id);
      label.textContent = labelText;
      row.appendChild(label);
      const textarea = document.createElement('textarea');
      textarea.id = id;
      textarea.dataset.blockField = field;
      textarea.value = value || '';
      row.appendChild(textarea);
      return row;
    }

    function createInputRow(id, labelText, value, type, field, required = false) {
      const row = document.createElement('div');
      row.className = 'form-row';
      const label = document.createElement('label');
      label.setAttribute('for', id);
      label.textContent = labelText;
      row.appendChild(label);
      const input = document.createElement('input');
      input.id = id;
      input.type = type;
      input.dataset.blockField = field;
      input.value = value || '';
      if (required) {
        input.required = true;
      }
      row.appendChild(input);
      return row;
    }

    function getEmailBlockLabel(type) {
      if (type === 'image') {
        return 'Image';
      }
      if (type === 'button') {
        return "Bouton d'appel à l'action";
      }
      return 'Paragraphe';
    }

    function handleEmailTemplateBlockInput(event) {
      const target = event.target;
      if (
        !(
          target instanceof HTMLInputElement ||
          target instanceof HTMLTextAreaElement ||
          target instanceof HTMLSelectElement
        )
      ) {
        return;
      }

      const field = target.dataset.blockField;
      if (!field) {
        return;
      }

      const blockElement = target.closest('.email-template-block');
      if (!blockElement) {
        return;
      }

      const blockId = blockElement.dataset.blockId || '';
      const block = emailTemplateDraftBlocks.find((item) => item && item.id === blockId);
      if (!block) {
        return;
      }

      if (block.type === 'paragraph') {
        if (field === 'align') {
          const alignValue = target.value.trim();
          block.data.align = EMAIL_TEMPLATE_ALIGN_VALUES.has(alignValue) ? alignValue : 'left';
          target.value = block.data.align;
        } else if (field === 'text') {
          block.data.text = target.value;
        }
      } else if (block.type === 'image') {
        if (field in block.data) {
          block.data[field] = target.value.trim();
        }
      } else if (block.type === 'button') {
        if (field === 'variant') {
          const variantValue = target.value.trim();
          block.data.variant = EMAIL_TEMPLATE_BUTTON_VARIANTS.has(variantValue)
            ? variantValue
            : 'primary';
          target.value = block.data.variant;
        } else if (field in block.data) {
          block.data[field] = target.value.trim();
        }
      }

      renderEmailTemplatePreview();
    }

    function handleEmailTemplateBlockAction(event) {
      const button = event.target instanceof HTMLElement ? event.target.closest('button[data-action]') : null;
      if (!(button instanceof HTMLButtonElement)) {
        return;
      }

      const blockElement = button.closest('.email-template-block');
      if (!blockElement) {
        return;
      }
      const blockId = blockElement.dataset.blockId || '';
      const action = button.dataset.action || '';
      const index = emailTemplateDraftBlocks.findIndex((item) => item && item.id === blockId);
      if (index === -1) {
        return;
      }

      if (action === 'move-up' && index > 0) {
        const [block] = emailTemplateDraftBlocks.splice(index, 1);
        emailTemplateDraftBlocks.splice(index - 1, 0, block);
        renderEmailTemplateBlocks();
        renderEmailTemplatePreview();
        focusEmailTemplateBlock(block.id);
        return;
      }

      if (action === 'move-down' && index < emailTemplateDraftBlocks.length - 1) {
        const [block] = emailTemplateDraftBlocks.splice(index, 1);
        emailTemplateDraftBlocks.splice(index + 1, 0, block);
        renderEmailTemplateBlocks();
        renderEmailTemplatePreview();
        focusEmailTemplateBlock(block.id);
        return;
      }

      if (action === 'duplicate') {
        const clone = cloneEmailTemplateBlock(emailTemplateDraftBlocks[index], { preserveId: false });
        emailTemplateDraftBlocks.splice(index + 1, 0, clone);
        renderEmailTemplateBlocks();
        renderEmailTemplatePreview();
        focusEmailTemplateBlock(clone.id);
        return;
      }

      if (action === 'remove') {
        emailTemplateDraftBlocks.splice(index, 1);
        renderEmailTemplateBlocks();
        renderEmailTemplatePreview();
      }
    }

    function focusEmailTemplateBlock(blockId, selector = '[data-block-field]') {
      if (!blockId || !emailTemplateBlocksContainer) {
        return;
      }
      const escapedId = escapeSelector(blockId);
      window.requestAnimationFrame(() => {
        const target = emailTemplateBlocksContainer.querySelector(
          `.email-template-block[data-block-id="${escapedId}"] ${selector}`,
        );
        if (target && typeof target.focus === 'function') {
          target.focus();
        }
      });
    }

    function renderEmailTemplatePreview() {
      renderTemplatePreview(emailTemplateDraftBlocks, emailTemplatePreview);
    }

    function renderTemplatePreview(blocks, container) {
      if (!container) {
        return;
      }

      container.innerHTML = '';

      const validBlocks = Array.isArray(blocks) ? blocks.filter((block) => block && block.type) : [];
      if (validBlocks.length === 0) {
        container.innerHTML =
          '<p class="empty-state">Ajoutez des blocs pour visualiser le rendu du mail.</p>';
        return;
      }

      const fragment = document.createDocumentFragment();
      validBlocks.forEach((block) => {
        let element = null;
        if (block.type === 'paragraph') {
          element = createParagraphPreview(block);
        } else if (block.type === 'image') {
          element = createImagePreview(block);
        } else if (block.type === 'button') {
          element = createButtonPreview(block);
        }
        if (element) {
          fragment.appendChild(element);
        }
      });
      container.appendChild(fragment);
    }

    function createParagraphPreview(block) {
      const paragraph = document.createElement('p');
      if (block.data.align && EMAIL_TEMPLATE_ALIGN_VALUES.has(block.data.align)) {
        paragraph.style.textAlign = block.data.align;
      }

      const text = typeof block.data.text === 'string' ? block.data.text : '';
      const lines = text.split(/\r?\n/);
      lines.forEach((line, index) => {
        paragraph.appendChild(createLinkifiedFragment(line));
        if (index < lines.length - 1) {
          paragraph.appendChild(document.createElement('br'));
        }
      });
      return paragraph;
    }

    function createImagePreview(block) {
      const figure = document.createElement('figure');
      const imageUrl = block.data.url || '';
      if (!imageUrl) {
        const placeholder = document.createElement('p');
        placeholder.textContent = 'Image non renseignée';
        figure.appendChild(placeholder);
        return figure;
      }

      const image = document.createElement('img');
      image.src = imageUrl;
      image.alt = block.data.alt || '';

      let imageContainer = image;
      if (block.data.link) {
        const link = document.createElement('a');
        link.href = block.data.link;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.appendChild(image);
        imageContainer = link;
      }
      figure.appendChild(imageContainer);

      if (block.data.caption) {
        const caption = document.createElement('figcaption');
        caption.textContent = block.data.caption;
        figure.appendChild(caption);
      }
      return figure;
    }

    function createButtonPreview(block) {
      const container = document.createElement('p');
      container.style.textAlign = 'center';
      const button = document.createElement('a');
      button.className = 'email-preview-button';
      if (block.data.variant === 'outline') {
        button.classList.add('email-preview-button--outline');
      }
      button.href = block.data.url || '#';
      button.target = '_blank';
      button.rel = 'noopener noreferrer';
      button.textContent = block.data.label || "Appel à l'action";
      container.appendChild(button);
      return container;
    }

    function createLinkifiedFragment(text) {
      const fragment = document.createDocumentFragment();
      if (!text) {
        fragment.appendChild(document.createTextNode(''));
        return fragment;
      }
      const urlPattern = /(https?:\/\/[^\s]+)/gi;
      let lastIndex = 0;
      let match = urlPattern.exec(text);
      while (match) {
        if (match.index > lastIndex) {
          fragment.appendChild(document.createTextNode(text.slice(lastIndex, match.index)));
        }
        const url = match[0];
        const link = document.createElement('a');
        link.href = url;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.textContent = url;
        fragment.appendChild(link);
        lastIndex = match.index + url.length;
        match = urlPattern.exec(text);
      }
      if (lastIndex < text.length) {
        fragment.appendChild(document.createTextNode(text.slice(lastIndex)));
      }
      return fragment;
    }

    function buildEmailHtmlFromBlocks(blocks, context = {}) {
      const safeBlocks = Array.isArray(blocks)
        ? blocks.filter((block) => block && typeof block === 'object')
        : [];

      const parts = [];

      safeBlocks.forEach((block) => {
        if (block.type === 'paragraph') {
          const rawText = block.data && typeof block.data.text === 'string' ? block.data.text : '';
          const align = block.data && typeof block.data.align === 'string' ? block.data.align : 'left';
          const safeAlign = EMAIL_TEMPLATE_ALIGN_VALUES.has(align) ? align : 'left';
          const lines = rawText.split(/\r?\n/).map((line) => escapeHtml(line));
          const textContent = lines.join('<br />') || '&nbsp;';
          parts.push(
            `<p style="margin:0;font-size:16px;line-height:1.6;text-align:${safeAlign};">${textContent}</p>`,
          );
        } else if (block.type === 'image') {
          const imageUrl = block.data && block.data.url ? block.data.url.trim() : '';
          if (!imageUrl) {
            return;
          }
          const alt = block.data && block.data.alt ? block.data.alt.trim() : '';
          const caption = block.data && block.data.caption ? block.data.caption.trim() : '';
          const link = block.data && block.data.link ? block.data.link.trim() : '';
          const escapedImage = `<img src="${escapeHtml(imageUrl)}" alt="${escapeHtml(alt)}" style="max-width:100%;height:auto;border-radius:8px;" />`;
          const imageContent = link
            ? `<a href="${escapeHtml(link)}" style="color:#2563eb;text-decoration:none;">${escapedImage}</a>`
            : escapedImage;
          let figureHtml = `<figure style="margin:0;">${imageContent}`;
          if (caption) {
            figureHtml += `<figcaption style="margin-top:8px;font-size:14px;color:#475569;">${escapeHtml(
              caption,
            )}</figcaption>`;
          }
          figureHtml += '</figure>';
          parts.push(figureHtml);
        } else if (block.type === 'button') {
          const label = block.data && block.data.label ? block.data.label.trim() : '';
          const url = block.data && block.data.url ? block.data.url.trim() : '';
          if (!label) {
            return;
          }
          const variant = block.data && block.data.variant ? block.data.variant : 'primary';
          const isOutline = variant === 'outline';
          const background = isOutline ? '#ffffff' : '#2563eb';
          const color = isOutline ? '#2563eb' : '#ffffff';
          const border = isOutline ? '1px solid #2563eb' : '1px solid #2563eb';
          const href = url ? escapeHtml(url) : '#';
          parts.push(
            `<p style="margin:0;text-align:center;"><a href="${href}" style="display:inline-block;padding:12px 22px;border-radius:999px;font-weight:600;background:${background};color:${color};border:${border};text-decoration:none;">${escapeHtml(
              label,
            )}</a></p>`,
          );
        }
      });

      const subject = typeof context.subject === 'string' ? context.subject.trim() : '';
      const documentTitle = subject ? escapeHtml(subject) : 'Campagne UManager';
      const headerHtml = subject
        ? `<h1 style="margin:0 0 16px;font-size:22px;font-weight:700;">${documentTitle}</h1>`
        : '';
      const bodyContent = parts.length > 0 ? parts.join('') : '<p style="margin:0;">&nbsp;</p>';

      return `<!doctype html><html lang="fr"><head><meta charset="utf-8" /><title>${documentTitle}</title></head><body style="margin:0;padding:32px;background-color:#f4f6fb;font-family:'Noto Sans','Segoe UI',Arial,sans-serif;color:#0f172a;"><div style="max-width:640px;margin:0 auto;background:#ffffff;border-radius:12px;padding:32px 28px;box-shadow:0 24px 48px rgba(15,23,42,0.08);">${headerHtml}<div style="display:flex;flex-direction:column;gap:18px;">${bodyContent}</div></div><p style="margin:24px 0 0;text-align:center;font-size:12px;color:#94a3b8;">Message envoyé depuis UManager</p></body></html>`;
    }

    function buildEmailTextFromBlocks(blocks, context = {}) {
      const safeBlocks = Array.isArray(blocks)
        ? blocks.filter((block) => block && typeof block === 'object')
        : [];

      const sections = [];

      safeBlocks.forEach((block) => {
        if (block.type === 'paragraph') {
          const text = block.data && typeof block.data.text === 'string' ? block.data.text.trim() : '';
          if (text) {
            sections.push(text);
          }
        } else if (block.type === 'image') {
          const caption = block.data && block.data.caption ? block.data.caption.trim() : '';
          const alt = block.data && block.data.alt ? block.data.alt.trim() : '';
          const url = block.data && block.data.url ? block.data.url.trim() : '';
          const lines = [];
          if (caption) {
            lines.push(caption);
          } else if (alt) {
            lines.push(alt);
          }
          if (url) {
            lines.push(url);
          }
          if (lines.length > 0) {
            sections.push(lines.join('\n'));
          }
        } else if (block.type === 'button') {
          const label = block.data && block.data.label ? block.data.label.trim() : '';
          const url = block.data && block.data.url ? block.data.url.trim() : '';
          if (label && url) {
            sections.push(`${label}\n${url}`);
          } else if (label || url) {
            sections.push(label || url);
          }
        }
      });

      const subject = typeof context.subject === 'string' ? context.subject.trim() : '';
      const content = subject ? [subject, ...sections] : sections.slice();
      const result = content.join('\n\n').trim();
      return result || subject;
    }

    function setEmailTemplateFeedback(message, status = '') {
      if (!emailTemplateFeedback) {
        return;
      }
      emailTemplateFeedback.textContent = message;
      emailTemplateFeedback.classList.remove('form-feedback--error', 'form-feedback--success');
      if (status === 'error') {
        emailTemplateFeedback.classList.add('form-feedback--error');
      } else if (status === 'success') {
        emailTemplateFeedback.classList.add('form-feedback--success');
      }
    }

    function resetEmailTemplateForm(shouldFocus = true, keepFeedback = false) {
      if (emailTemplateForm) {
        emailTemplateForm.reset();
        delete emailTemplateForm.dataset.editingId;
      }
      emailTemplateEditingId = '';
      emailTemplateDraftBlocks = [];
      if (!keepFeedback) {
        setEmailTemplateFeedback('');
      }
      renderEmailTemplateBlocks();
      renderEmailTemplatePreview();
      if (shouldFocus && emailTemplateNameInput && typeof emailTemplateNameInput.focus === 'function') {
        emailTemplateNameInput.focus();
      }
    }

    function loadEmailTemplateForEditing(templateId) {
      const template = getTemplateById(templateId);
      if (!template) {
        return;
      }

      emailTemplateEditingId = template.id || '';
      if (emailTemplateForm) {
        emailTemplateForm.dataset.editingId = template.id || '';
      }
      if (emailTemplateNameInput) {
        emailTemplateNameInput.value = template.name || '';
      }
      if (emailTemplateSubjectInput) {
        emailTemplateSubjectInput.value = template.subject || '';
      }
      emailTemplateDraftBlocks = Array.isArray(template.blocks)
        ? template.blocks.map((block) => cloneEmailTemplateBlock(block, { preserveId: true }))
        : [];
      renderEmailTemplateBlocks();
      renderEmailTemplatePreview();
      setEmailTemplateFeedback(
        `Modification du modèle « ${template.name || 'Sans titre'} ».`,
        '',
      );
      if (emailTemplateNameInput && typeof emailTemplateNameInput.focus === 'function') {
        emailTemplateNameInput.focus();
      }
    }

    function duplicateEmailTemplate(templateId) {
      const template = getTemplateById(templateId);
      if (!template) {
        return;
      }
      const now = new Date().toISOString();
      const copy = {
        id: generateId('email-template'),
        name: `${template.name || 'Modèle sans titre'} (copie)`,
        subject: template.subject || '',
        blocks: Array.isArray(template.blocks)
          ? template.blocks.map((block) => cloneEmailTemplateBlock(block, { preserveId: false }))
          : [],
        createdAt: now,
        updatedAt: now,
      };
      if (!Array.isArray(data.emailTemplates)) {
        data.emailTemplates = [];
      }
      data.emailTemplates.push(copy);
      data.lastUpdated = now;
      saveDataForUser(currentUser, data);
      renderEmailTemplateList();
      renderCampaignTemplateOptions();
      setEmailTemplateFeedback(`Modèle « ${template.name || 'Sans titre'} » dupliqué.`, 'success');
      notifyDataChanged('email-templates', { type: 'duplicated', templateId: copy.id });
    }

    function deleteEmailTemplate(templateId) {
      if (!Array.isArray(data.emailTemplates)) {
        return;
      }
      const index = data.emailTemplates.findIndex((item) => item && item.id === templateId);
      if (index === -1) {
        return;
      }
      const [removed] = data.emailTemplates.splice(index, 1);
      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      if (emailTemplateEditingId === templateId) {
        resetEmailTemplateForm(false);
      }
      if (
        campaignTemplateSelect instanceof HTMLSelectElement &&
        campaignTemplateSelect.value === templateId
      ) {
        campaignTemplateSelect.value = '';
        campaignActiveTemplateId = '';
        updateCampaignTemplatePreview();
      }
      renderEmailTemplateList();
      renderCampaignTemplateOptions();
      setEmailTemplateFeedback(
        `Modèle « ${(removed && removed.name) || 'Sans titre'} » supprimé.`,
        'success',
      );
      notifyDataChanged('email-templates', { type: 'deleted', templateId });
    }

    function renderEmailTemplateList() {
      const templates = getSortedEmailTemplates();
      if (emailTemplateEmptyState) {
        emailTemplateEmptyState.hidden = templates.length > 0;
      }
      if (!emailTemplateList) {
        return;
      }
      emailTemplateList.innerHTML = '';
      if (templates.length === 0) {
        return;
      }
      const fragment = document.createDocumentFragment();
      templates.forEach((template) => {
        const listItem = document.createElement('li');
        listItem.className = 'email-template-card';
        listItem.dataset.templateId = template.id || '';

        const header = document.createElement('div');
        header.className = 'email-template-card-header';
        const title = document.createElement('h3');
        title.textContent = template.name || 'Modèle sans titre';
        header.appendChild(title);
        const meta = document.createElement('p');
        meta.className = 'email-template-card-meta';
        meta.textContent = buildEmailTemplateMeta(template);
        header.appendChild(meta);
        listItem.appendChild(header);

        const preview = document.createElement('div');
        preview.className = 'email-template-card-preview';
        preview.textContent = summarizeTemplateBlocks(template.blocks);
        listItem.appendChild(preview);

        const actions = document.createElement('div');
        actions.className = 'email-template-card-actions';
        actions.appendChild(createTemplateActionButton('Modifier', 'edit-template', template.id));
        actions.appendChild(createTemplateActionButton('Dupliquer', 'duplicate-template', template.id));
        actions.appendChild(
          createTemplateActionButton('Supprimer', 'delete-template', template.id, true),
        );
        listItem.appendChild(actions);

        fragment.appendChild(listItem);
      });
      emailTemplateList.appendChild(fragment);
    }

    function createTemplateActionButton(label, action, templateId, danger = false) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = danger ? 'secondary-button ghost-button--danger' : 'secondary-button';
      button.dataset.action = action;
      button.dataset.templateId = templateId || '';
      button.textContent = label;
      return button;
    }

    function summarizeTemplateBlocks(blocks) {
      if (!Array.isArray(blocks) || blocks.length === 0) {
        return 'Aucun contenu pour le moment.';
      }
      const descriptions = [];
      blocks.forEach((block) => {
        if (!block) {
          return;
        }
        if (block.type === 'paragraph') {
          const text = block.data && typeof block.data.text === 'string' ? block.data.text.trim() : '';
          if (text) {
            descriptions.push(text.replace(/\s+/g, ' ').slice(0, 120));
          }
          return;
        }
        if (block.type === 'image') {
          descriptions.push('Bloc image');
          return;
        }
        if (block.type === 'button') {
          const label = block.data && typeof block.data.label === 'string' ? block.data.label.trim() : '';
          descriptions.push(`Bouton : ${label || 'Sans libellé'}`);
        }
      });
      if (descriptions.length === 0) {
        return 'Aucun contenu pour le moment.';
      }
      return descriptions.join(' · ');
    }

    function buildEmailTemplateMeta(template) {
      const parts = [];
      if (template.subject) {
        parts.push(`Objet : ${template.subject}`);
      }
      const blockCount = Array.isArray(template.blocks) ? template.blocks.length : 0;
      if (blockCount === 1) {
        parts.push('1 bloc');
      } else if (blockCount > 1) {
        parts.push(`${blockCount} blocs`);
      }
      const timestamp = getTemplateTimestamp(template);
      if (timestamp) {
        parts.push(`Mis à jour le ${campaignDateFormatter.format(new Date(timestamp))}`);
      }
      return parts.join(' · ');
    }

    function getSortedEmailTemplates() {
      const templates = Array.isArray(data.emailTemplates) ? data.emailTemplates.slice() : [];
      templates.sort((a, b) => {
        const diff = getTemplateTimestamp(b) - getTemplateTimestamp(a);
        if (diff !== 0) {
          return diff;
        }
        const nameA = a && typeof a.name === 'string' ? a.name : '';
        const nameB = b && typeof b.name === 'string' ? b.name : '';
        return nameA.localeCompare(nameB, 'fr', { sensitivity: 'base' });
      });
      return templates;
    }

    function getTemplateTimestamp(template) {
      if (!template || typeof template !== 'object') {
        return 0;
      }
      const updated = Date.parse(template.updatedAt || '');
      if (!Number.isNaN(updated)) {
        return updated;
      }
      const created = Date.parse(template.createdAt || '');
      if (!Number.isNaN(created)) {
        return created;
      }
      return 0;
    }

    function getTemplateById(templateId) {
      if (!templateId || !Array.isArray(data.emailTemplates)) {
        return null;
      }
      return data.emailTemplates.find((template) => template && template.id === templateId) || null;
    }

    function renderCampaignTemplateOptions() {
      if (!(campaignTemplateSelect instanceof HTMLSelectElement)) {
        return;
      }
      const templates = getSortedEmailTemplates();
      const previousValue = campaignTemplateSelect.value || '';
      campaignTemplateSelect.innerHTML = '';
      const placeholder = document.createElement('option');
      const hasTemplates = templates.length > 0;
      placeholder.value = '';
      placeholder.textContent = hasTemplates ? 'Sélectionnez un modèle' : 'Aucun modèle disponible';
      placeholder.disabled = hasTemplates;
      placeholder.selected = true;
      campaignTemplateSelect.appendChild(placeholder);
      templates.forEach((template) => {
        const option = document.createElement('option');
        option.value = template.id || '';
        option.textContent = template.name || 'Modèle sans titre';
        campaignTemplateSelect.appendChild(option);
      });
      campaignTemplateSelect.disabled = !hasTemplates;
      if (hasTemplates && templates.some((template) => template.id === previousValue)) {
        campaignTemplateSelect.value = previousValue;
        campaignActiveTemplateId = previousValue;
      } else {
        campaignTemplateSelect.value = '';
        campaignActiveTemplateId = '';
      }
      updateCampaignTemplatePreview();
    }

    function setEmailCampaignFeedback(message, status = '') {
      if (!emailCampaignFeedback) {
        return;
      }
      emailCampaignFeedback.textContent = message;
      emailCampaignFeedback.classList.remove('form-feedback--error', 'form-feedback--success');
      if (status === 'error') {
        emailCampaignFeedback.classList.add('form-feedback--error');
      } else if (status === 'success') {
        emailCampaignFeedback.classList.add('form-feedback--success');
      }
    }

    function dispatchCampaignSend(campaignRecord, template, sender) {
      const api = getEmailApi();
      if (!api || typeof api.sendCampaign !== 'function') {
        return 'skipped';
      }

      try {
        const blocks = Array.isArray(template && template.blocks) ? template.blocks : [];
        const html = buildEmailHtmlFromBlocks(blocks, { subject: campaignRecord.subject });
        const text = buildEmailTextFromBlocks(blocks, { subject: campaignRecord.subject });

        const senderEmailValue =
          sender && typeof sender.senderEmail === 'string'
            ? sender.senderEmail
            : sender && typeof sender.email === 'string'
            ? sender.email
            : '';
        const senderNameValue =
          sender && typeof sender.senderName === 'string'
            ? sender.senderName
            : sender && typeof sender.name === 'string'
            ? sender.name
            : '';

        const payload = {
          campaignId: campaignRecord.id,
          name: campaignRecord.name,
          subject: campaignRecord.subject,
          sender: {
            email: senderEmailValue,
            name: senderNameValue,
          },
          recipients: Array.isArray(campaignRecord.recipients)
            ? campaignRecord.recipients.map((recipient) => ({
                email: recipient.email,
                name: recipient.name,
                contactId: recipient.contactId,
              }))
            : [],
          html,
          text,
        };

        setEmailCampaignFeedback(
          `Campagne « ${campaignRecord.name} » enregistrée. Envoi en cours…`,
          'success',
        );

        Promise.resolve(api.sendCampaign(payload))
          .then(() => {
            const recipientCount = Number.isFinite(campaignRecord.recipientCount)
              ? campaignRecord.recipientCount
              : payload.recipients.length;
            const label = recipientCount > 1 ? 'destinataires' : 'destinataire';
            setEmailCampaignFeedback(
              `Campagne « ${campaignRecord.name} » envoyée à ${numberFormatter.format(recipientCount)} ${label}.`,
              'success',
            );
          })
          .catch((error) => {
            console.error("Erreur lors de l'envoi de la campagne :", error);
            setEmailCampaignFeedback(
              "La campagne a été enregistrée mais l'envoi a échoué. Vérifiez la configuration du serveur mail.",
              'error',
            );
          });

        return 'pending';
      } catch (error) {
        console.error("Impossible de préparer l'envoi de la campagne :", error);
        setEmailCampaignFeedback(
          "La campagne a été enregistrée mais l'envoi automatique n'a pas pu être préparé.",
          'error',
        );
        return 'error';
      }
    }

    function updateCampaignTemplatePreview() {
      const templateId =
        campaignTemplateSelect instanceof HTMLSelectElement ? campaignTemplateSelect.value : '';
      const template = getTemplateById(templateId);
      if (!template) {
        campaignActiveTemplateId = '';
        if (campaignEmailPreview) {
          campaignEmailPreview.innerHTML = campaignEmailPreviewEmptyHtml;
        }
        if (!campaignSubjectEdited && emailCampaignSubjectInput) {
          emailCampaignSubjectInput.value = '';
        }
        campaignSubjectTemplateId = '';
        return;
      }

      campaignActiveTemplateId = template.id || '';
      renderTemplatePreview(template.blocks, campaignEmailPreview);
      if (emailCampaignSubjectInput) {
        const templateSubject = template.subject || '';
        if (!campaignSubjectEdited) {
          emailCampaignSubjectInput.value = templateSubject;
        }
        campaignSubjectTemplateId = template.id || '';
        if (emailCampaignSubjectInput.value.trim() === templateSubject.trim()) {
          campaignSubjectEdited = false;
        }
      }
    }

    function updateCampaignRecipients() {
      if (!campaignSummary) {
        return;
      }
      const savedSearchId =
        campaignSavedSearchSelect instanceof HTMLSelectElement ? campaignSavedSearchSelect.value : '';
      if (!savedSearchId) {
        campaignRecipients = [];
        campaignSummary.textContent = campaignSummaryDefaultText;
        if (campaignRecipientList) {
          campaignRecipientList.innerHTML = '';
        }
        return;
      }

      const { recipients, savedSearch } = getRecipientsForSavedSearch(savedSearchId);
      campaignRecipients = recipients;

      if (campaignRecipientList) {
        campaignRecipientList.innerHTML = '';
        if (recipients.length > 0) {
          const fragment = document.createDocumentFragment();
          recipients.forEach((recipient) => {
            const item = document.createElement('li');
            item.className = 'campaign-recipient-item';
            const emailLine = document.createElement('p');
            emailLine.className = 'campaign-recipient-email';
            emailLine.textContent = recipient.email;
            item.appendChild(emailLine);
            const nameLine = document.createElement('p');
            nameLine.className = 'campaign-recipient-name';
            nameLine.textContent = recipient.name || 'Contact';
            item.appendChild(nameLine);
            fragment.appendChild(item);
          });
          campaignRecipientList.appendChild(fragment);
        }
      }

      if (recipients.length === 0) {
        campaignSummary.textContent =
          "Aucun contact muni d'une adresse mail dans cette recherche sauvegardée.";
      } else {
        const suffix = recipients.length > 1 ? 'destinataires' : 'destinataire';
        const searchName = savedSearch ? savedSearch.name : '';
        campaignSummary.textContent = `${numberFormatter.format(
          recipients.length,
        )} ${suffix} recevront ce message via la recherche « ${searchName} ».`;
      }
    }

    function getRecipientsForSavedSearch(savedSearchId) {
      if (!savedSearchId || !Array.isArray(data.savedSearches)) {
        return { recipients: [], savedSearch: null };
      }
      const savedSearch =
        data.savedSearches.find((item) => item && item.id === savedSearchId) || null;
      if (!savedSearch) {
        return { recipients: [], savedSearch: null };
      }

      const categoriesById = buildCategoryMap();
      const keywordsById = new Map(
        Array.isArray(data.keywords) ? data.keywords.map((keyword) => [keyword.id, keyword]) : [],
      );
      const { matches } = filterContactsUsingSearch(
        Array.isArray(data.contacts) ? data.contacts.slice() : [],
        savedSearch.searchTerm || '',
        cloneAdvancedFilters(savedSearch.advancedFilters || {}),
        categoriesById,
        keywordsById,
      );

      const emailSet = new Set();
      const recipients = [];
      matches.forEach((contact) => {
        const channels = computeContactChannels(contact, categoriesById);
        channels.emails.forEach((email) => {
          if (!email || emailSet.has(email)) {
            return;
          }
          emailSet.add(email);
          recipients.push({
            contactId: contact.id || '',
            email,
            name: getContactDisplayName(contact, categoriesById),
          });
        });
      });
      return { recipients, savedSearch };
    }

    function renderCampaignHistory() {
      if (campaignHistoryEmpty) {
        campaignHistoryEmpty.hidden = Array.isArray(data.emailCampaigns) && data.emailCampaigns.length > 0;
      }
      if (!campaignHistoryList) {
        return;
      }
      campaignHistoryList.innerHTML = '';
      const campaigns = Array.isArray(data.emailCampaigns) ? data.emailCampaigns.slice() : [];
      if (campaigns.length === 0) {
        return;
      }
      campaigns.sort((a, b) => getCampaignTimestamp(b) - getCampaignTimestamp(a));
      const fragment = document.createDocumentFragment();
      campaigns.forEach((campaign) => {
        const item = document.createElement('li');
        item.className = 'campaign-history-item';
        const title = document.createElement('h3');
        title.className = 'campaign-history-title';
        title.textContent = campaign.name || 'Campagne sans nom';
        item.appendChild(title);
        const meta = document.createElement('p');
        meta.className = 'campaign-history-meta';
        meta.textContent = buildCampaignHistoryMeta(campaign);
        item.appendChild(meta);
        const summary = document.createElement('p');
        summary.className = 'campaign-history-summary';
        summary.textContent = `Objet : ${campaign.subject || '—'} · Modèle : ${
          campaign.templateName || 'Sans titre'
        }`;
        item.appendChild(summary);
        fragment.appendChild(item);
      });
      campaignHistoryList.appendChild(fragment);
    }

    function buildCampaignHistoryMeta(campaign) {
      const parts = [];
      const timestamp = getCampaignTimestamp(campaign);
      if (timestamp) {
        parts.push(`Envoyée le ${campaignDateFormatter.format(new Date(timestamp))}`);
      }
      const count = Number.isFinite(campaign.recipientCount) ? campaign.recipientCount : 0;
      if (count > 0) {
        const suffix = count > 1 ? 'destinataires' : 'destinataire';
        parts.push(`${numberFormatter.format(count)} ${suffix}`);
      } else {
        parts.push('Aucun destinataire');
      }
      if (campaign.senderEmail) {
        parts.push(`Expéditeur : ${campaign.senderEmail}`);
      }
      return parts.join(' · ');
    }

    function getCampaignTimestamp(campaign) {
      if (!campaign || typeof campaign !== 'object') {
        return 0;
      }
      const sentAt = Date.parse(campaign.sentAt || '');
      if (!Number.isNaN(sentAt)) {
        return sentAt;
      }
      const createdAt = Date.parse(campaign.createdAt || '');
      if (!Number.isNaN(createdAt)) {
        return createdAt;
      }
      return 0;
    }

    function resetCampaignForm(shouldFocus = true, options = {}) {
      const keepFeedback = Boolean(options && options.keepFeedback);
      campaignRecipients = [];
      campaignSubjectEdited = false;
      campaignActiveTemplateId = '';
      campaignSubjectTemplateId = '';
      if (campaignTemplateSelect instanceof HTMLSelectElement) {
        campaignTemplateSelect.value = '';
      }
      if (campaignSavedSearchSelect instanceof HTMLSelectElement) {
        campaignSavedSearchSelect.value = '';
      }
      if (campaignSummary) {
        campaignSummary.textContent = campaignSummaryDefaultText;
      }
      if (campaignRecipientList) {
        campaignRecipientList.innerHTML = '';
      }
      if (campaignEmailPreview) {
        campaignEmailPreview.innerHTML = campaignEmailPreviewEmptyHtml;
      }
      if (!keepFeedback) {
        setEmailCampaignFeedback('');
      }
      if (shouldFocus && campaignTemplateSelect instanceof HTMLSelectElement) {
        campaignTemplateSelect.focus();
      }
    }

    function escapeSelector(value) {
      if (typeof value !== 'string') {
        return '';
      }
      if (typeof CSS !== 'undefined' && typeof CSS.escape === 'function') {
        return CSS.escape(value);
      }
      return value.replace(/(["\\])/g, '\\$1');
    }

    function sortSavedSearchesInPlace() {
      if (!Array.isArray(data.savedSearches)) {
        data.savedSearches = [];
        return;
      }

      data.savedSearches.sort((a, b) => {
        const diff = getSavedSearchTimestamp(b) - getSavedSearchTimestamp(a);
        if (diff !== 0) {
          return diff;
        }
        const nameA = a && typeof a.name === 'string' ? a.name : '';
        const nameB = b && typeof b.name === 'string' ? b.name : '';
        return nameA.localeCompare(nameB, 'fr', { sensitivity: 'base' });
      });
    }

    function renderBulkKeywordOptions() {
      if (!(contactBulkKeywordSelect instanceof HTMLSelectElement)) {
        return;
      }

      const keywords = Array.isArray(data.keywords)
        ? data.keywords
            .slice()
            .sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }))
        : [];

      const previousValue = contactBulkKeywordSelect.value || '';
      contactBulkKeywordSelect.innerHTML = '';

      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = keywords.length === 0
        ? 'Aucun mot clé disponible'
        : 'Choisir un mot clé';
      contactBulkKeywordSelect.appendChild(placeholder);

      let restored = false;
      keywords.forEach((keyword) => {
        const option = document.createElement('option');
        option.value = keyword.id || '';
        option.textContent = keyword.name;
        if (keyword.id && keyword.id === previousValue) {
          option.selected = true;
          restored = true;
        }
        contactBulkKeywordSelect.appendChild(option);
      });

      if (!restored) {
        contactBulkKeywordSelect.value = '';
      }

      contactBulkKeywordSelect.disabled = keywords.length === 0;
      updateSelectedContactsUI();
    }

    function filterContactsUsingSearch(
      contacts,
      searchTerm,
      filters,
      categoriesById,
      keywordsById,
    ) {
      const normalizedTerm = typeof searchTerm === 'string' ? searchTerm.trim().toLowerCase() : '';
      const categoryFilterEntries =
        filters && filters.categories && typeof filters.categories === 'object'
          ? Object.entries(filters.categories).filter(([, filter]) => filter && typeof filter === 'object')
          : [];
      const keywordFilters = Array.isArray(filters && filters.keywords)
        ? filters.keywords.filter((value) => Boolean(value))
        : [];
      const keywordMode =
        filters && filters.keywordMode === KEYWORD_FILTER_MODE_ANY
          ? KEYWORD_FILTER_MODE_ANY
          : KEYWORD_FILTER_MODE_ALL;
      const hasAdvancedFilters = categoryFilterEntries.length > 0 || keywordFilters.length > 0;

      const matches = contacts.filter((contact) =>
        contactMatchesSearchCriteria(
          contact,
          normalizedTerm,
          categoryFilterEntries,
          keywordFilters,
          keywordMode,
          categoriesById,
          keywordsById,
        ),
      );

      matches.sort((a, b) => {
        const nameA = getContactDisplayName(a, categoriesById).toString();
        const nameB = getContactDisplayName(b, categoriesById).toString();
        return nameA.localeCompare(nameB, 'fr', { sensitivity: 'base' });
      });

      return { matches, normalizedTerm, hasAdvancedFilters };
    }

    function contactMatchesSearchCriteria(
      contact,
      normalizedTerm,
      categoryFilterEntries,
      keywordFilters,
      keywordMode,
      categoriesById,
      keywordsById,
    ) {
      if (!contact || typeof contact !== 'object') {
        return false;
      }

      const categoryValues =
        contact.categoryValues && typeof contact.categoryValues === 'object'
          ? contact.categoryValues
          : {};
      const keywords = Array.isArray(contact.keywords) ? contact.keywords : [];

      for (const [categoryId, filter] of categoryFilterEntries) {
        if (!filter || typeof filter !== 'object') {
          continue;
        }

        const rawValue = categoryValues[categoryId];
        const valueString = rawValue === undefined || rawValue === null ? '' : rawValue.toString().trim();
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
        const matchesKeywords =
          keywordMode === KEYWORD_FILTER_MODE_ANY
            ? keywordFilters.some((keywordId) => keywordSet.has(keywordId))
            : keywordFilters.every((keywordId) => keywordSet.has(keywordId));
        if (!matchesKeywords) {
          return false;
        }
      }

      if (!normalizedTerm) {
        return true;
      }

      const normalizedType = normalizeContactType(contact.type);
      const typeConfig =
        CONTACT_TYPE_MAP.get(normalizedType) || CONTACT_TYPE_MAP.get(CONTACT_TYPE_DEFAULT);
      const displayName = getContactDisplayName(contact, categoriesById);
      const channels = computeContactChannels(contact, categoriesById);

      const categoryEntries = Object.entries(categoryValues)
        .map(([categoryId, rawValue]) => {
          const category = categoriesById.get(categoryId);
          if (!category) {
            return null;
          }
          const value = rawValue === undefined || rawValue === null ? '' : rawValue.toString().trim();
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
        typeConfig ? typeConfig.label : '',
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
    }

    function normalizeContactType(rawType) {
      if (typeof rawType === 'string') {
        const normalized = rawType.trim().toLowerCase();
        if (CONTACT_TYPE_MAP.has(normalized)) {
          return normalized;
        }
      }
      return CONTACT_TYPE_DEFAULT;
    }

    function updateSelectedContactsUI() {
      const selectedCount = selectedContactIds.size;
      if (contactSelectedCountEl) {
        if (selectedCount === 0) {
          contactSelectedCountEl.textContent = 'Aucun contact sélectionné';
        } else {
          const suffix = selectedCount > 1 ? 's' : '';
          contactSelectedCountEl.textContent = `${numberFormatter.format(
            selectedCount,
          )} contact${suffix} sélectionné${suffix}`;
        }
      }

      if (contactDeleteSelectedButton instanceof HTMLButtonElement) {
        contactDeleteSelectedButton.disabled = selectedCount === 0;
      }

      const keywordSelectEnabled =
        contactBulkKeywordSelect instanceof HTMLSelectElement && !contactBulkKeywordSelect.disabled;
      if (contactBulkKeywordButton instanceof HTMLButtonElement) {
        const hasKeywordSelection =
          keywordSelectEnabled && contactBulkKeywordSelect && contactBulkKeywordSelect.value;
        contactBulkKeywordButton.disabled = selectedCount === 0 || !hasKeywordSelection;
      }

      if (contactSelectAllButton instanceof HTMLButtonElement) {
        contactSelectAllButton.disabled = lastContactSearchResultIds.length === 0;
      }
    }

    function refreshContactSelectionCheckboxes() {
      if (!contactList) {
        return;
      }
      contactList.querySelectorAll('.contact-select-checkbox').forEach((element) => {
        if (!(element instanceof HTMLInputElement)) {
          return;
        }
        const contactId = element.dataset.contactId || '';
        element.checked = contactId ? selectedContactIds.has(contactId) : false;
      });
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

      const validContactIds = new Set(
        contacts
          .map((contact) => (contact && contact.id ? contact.id : ''))
          .filter((id) => Boolean(id)),
      );
      let selectionChanged = false;
      selectedContactIds.forEach((id) => {
        if (!validContactIds.has(id)) {
          selectedContactIds.delete(id);
          selectionChanged = true;
        }
      });
      if (selectionChanged) {
        updateSelectedContactsUI();
      }

      const categoriesById = buildCategoryMap();
      const keywordsById = new Map(
        Array.isArray(data.keywords)
          ? data.keywords.map((keyword) => [keyword.id, keyword])
          : [],
      );

      const {
        matches: filteredContacts,
        normalizedTerm,
        hasAdvancedFilters,
      } = filterContactsUsingSearch(contacts, contactSearchTerm, advancedFilters, categoriesById, keywordsById);

      lastContactSearchResultIds = filteredContacts
        .map((contact) => (contact && contact.id ? contact.id : ''))
        .filter((id) => Boolean(id));

      const totalResults = filteredContacts.length;

      if (contactSearchCountEl) {
        contactSearchCountEl.textContent = totalResults.toString();
      }

      updateSelectedContactsUI();

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

        const selectCheckbox = listItem.querySelector('.contact-select-checkbox');
        if (selectCheckbox instanceof HTMLInputElement) {
          if (contact.id) {
            selectCheckbox.disabled = false;
            selectCheckbox.dataset.contactId = contact.id;
            selectCheckbox.checked = selectedContactIds.has(contact.id);
            selectCheckbox.addEventListener('change', () => {
              if (!contact.id) {
                return;
              }
              if (selectCheckbox.checked) {
                selectedContactIds.add(contact.id);
              } else {
                selectedContactIds.delete(contact.id);
              }
              updateSelectedContactsUI();
            });
          } else {
            selectCheckbox.checked = false;
            selectCheckbox.disabled = true;
            selectCheckbox.removeAttribute('data-contact-id');
          }
        }

        const nameEl = listItem.querySelector('.contact-name');
        if (nameEl) {
          const displayName = getContactDisplayName(contact, categoriesById) || 'Contact sans nom';
          nameEl.textContent = displayName;
        }

        const typeBadgeEl = listItem.querySelector('.contact-type-badge');
        if (typeBadgeEl) {
          const normalizedType = normalizeContactType(contact.type);
          const typeConfig =
            CONTACT_TYPE_MAP.get(normalizedType) || CONTACT_TYPE_MAP.get(CONTACT_TYPE_DEFAULT);
          typeBadgeEl.textContent = typeConfig ? typeConfig.label : '';
          typeBadgeEl.className = 'contact-type-badge';
          if (typeConfig && typeConfig.badgeClass) {
            typeBadgeEl.classList.add(typeConfig.badgeClass);
          }
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
      updateSelectedContactsUI();
    }

    function resetContactForm(shouldFocus = true) {
      if (!contactForm) {
        return;
      }

      contactForm.reset();
      if (contactTypeSelect instanceof HTMLSelectElement) {
        contactTypeSelect.value = CONTACT_TYPE_DEFAULT;
      }
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
      if (contactTypeSelect instanceof HTMLSelectElement) {
        contactTypeSelect.value = normalizeContactType(contact.type);
      }
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

    function removeContactsByIds(contactIds = [], options = {}) {
      if (!Array.isArray(data.contacts) || data.contacts.length === 0) {
        return 0;
      }

      const normalizedIds = Array.isArray(contactIds)
        ? contactIds.filter((value) => typeof value === 'string' && value)
        : [];
      if (normalizedIds.length === 0) {
        return 0;
      }

      const idSet = new Set(normalizedIds);
      const initialLength = data.contacts.length;
      const isEditingRemoved =
        contactForm && contactForm.dataset.editingId && idSet.has(contactForm.dataset.editingId);

      data.contacts = data.contacts.filter((contact) => !contact || !idSet.has(contact.id));
      const removedCount = initialLength - data.contacts.length;
      if (removedCount === 0) {
        return 0;
      }

      data.lastUpdated = new Date().toISOString();
      updateMetricsFromContacts();
      saveDataForUser(currentUser, data);

      if (isEditingRemoved) {
        const targetPage = contactEditReturnPage || 'contacts-search';
        resetContactForm(false);
        showPage(targetPage);
      }

      selectedContactIds = new Set([...selectedContactIds].filter((id) => !idSet.has(id)));
      const reasonFromOptions =
        options && typeof options.reason === 'string' && options.reason.trim()
          ? options.reason.trim()
          : null;
      const reason = reasonFromOptions || (idSet.size > 1 ? 'bulk-delete' : 'delete');
      notifyDataChanged('contacts', { reason, removedCount });
      renderMetrics();
      renderContacts();
      return removedCount;
    }

    function addKeywordToContacts(contactIds = [], keywordId) {
      if (!keywordId) {
        return 0;
      }

      const normalizedIds = Array.isArray(contactIds)
        ? contactIds.filter((value) => typeof value === 'string' && value)
        : [];
      if (normalizedIds.length === 0) {
        return 0;
      }

      const idSet = new Set(normalizedIds);
      const contacts = Array.isArray(data.contacts) ? data.contacts : [];
      let updatedCount = 0;

      contacts.forEach((contact) => {
        if (!contact || !contact.id || !idSet.has(contact.id)) {
          return;
        }
        if (!Array.isArray(contact.keywords)) {
          contact.keywords = [];
        }
        if (!contact.keywords.includes(keywordId)) {
          contact.keywords.push(keywordId);
          updatedCount += 1;
        }
      });

      if (updatedCount === 0) {
        return 0;
      }

      data.lastUpdated = new Date().toISOString();
      saveDataForUser(currentUser, data);
      notifyDataChanged('contacts', {
        reason: 'bulk-add-keyword',
        keywordId,
        affectedCount: updatedCount,
      });
      renderMetrics();
      renderContacts();
      return updatedCount;
    }

    function deleteContact(contactId) {
      if (!contactId || !Array.isArray(data.contacts)) {
        return;
      }

      removeContactsByIds([contactId], { reason: 'delete' });
    }

    function startCategoryEdition(categoryId) {
      if (REQUIRED_CATEGORY_ID_SET.has(categoryId)) {
        return;
      }
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
      if (REQUIRED_CATEGORY_ID_SET.has(categoryId)) {
        return;
      }
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

      const startIndex = skipHeader ? 1 : 0;
      const totalRows = startIndex < safeRows.length ? safeRows.length - startIndex : 0;
      const errors = [];
      const errorRows = new Set();
      let importedCount = 0;
      let skippedEmptyCount = 0;
      let updatedContactsCount = 0;

      if (normalizedMapping.length === 0 || safeRows.length === 0 || totalRows === 0) {
        return {
          importedCount,
          mergedCount: 0,
          skippedEmptyCount: totalRows,
          totalRows,
          errorCount: 0,
          errors,
          fileName,
          duplicatesDetected: 0,
          autoMergeApplied: false,
          updatedCount: 0,
        };
      }

      const requiredImportCategoryIds = [
        REQUIRED_CATEGORY_IDS.firstName,
        REQUIRED_CATEGORY_IDS.lastName,
        REQUIRED_CATEGORY_IDS.birthDate,
      ];

      const missingRequiredMappings = requiredImportCategoryIds.filter(
        (requiredId) => !normalizedMapping.some((mapping) => mapping.categoryId === requiredId),
      );

      if (missingRequiredMappings.length > 0) {
        const missingLabels = missingRequiredMappings.map((categoryId) => {
          const category = categoriesById.get(categoryId);
          return category && category.name ? `« ${category.name} »` : 'une catégorie requise';
        });
        errors.push({
          row: 0,
          categoryId: '',
          message: `Les catégories obligatoires ${missingLabels.join(', ')} doivent être associées à une colonne pour importer des contacts.`,
        });
        return {
          importedCount,
          mergedCount: 0,
          skippedEmptyCount: totalRows,
          totalRows,
          errorCount: errors.length,
          errors,
          fileName,
          duplicatesDetected: 0,
          autoMergeApplied: false,
          updatedCount: 0,
        };
      }

      const existingContactsByIdentifier = new Map();
      data.contacts.forEach((contact) => {
        if (!contact || typeof contact !== 'object') {
          return;
        }
        if (!contact.categoryValues || typeof contact.categoryValues !== 'object') {
          contact.categoryValues = {};
        }
        const existingIdentifier = contact.categoryValues[REQUIRED_CATEGORY_IDS.identifier];
        if (existingIdentifier) {
          existingContactsByIdentifier.set(existingIdentifier, contact);
          return;
        }
        const identifierData = ensureContactIdentifier(contact.categoryValues);
        if (identifierData) {
          existingContactsByIdentifier.set(identifierData.identifier, contact);
        }
      });

      for (let rowIndex = startIndex; rowIndex < safeRows.length; rowIndex += 1) {
        const row = Array.isArray(safeRows[rowIndex]) ? safeRows[rowIndex] : [];
        const categoryValues = {};
        const rawValues = new Map();
        let hasNonEmptyValue = false;

        normalizedMapping.forEach((mapping) => {
          const cellValue = row[mapping.columnIndex];
          const valueString =
            cellValue === undefined || cellValue === null ? '' : cellValue.toString().trim();
          rawValues.set(mapping.categoryId, valueString);

          if (!valueString) {
            return;
          }

          hasNonEmptyValue = true;

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
            return;
          }

          categoryValues[mapping.categoryId] = valueString;
        });

        if (!hasNonEmptyValue) {
          skippedEmptyCount += 1;
          continue;
        }

        const firstNameValue =
          categoryValues[REQUIRED_CATEGORY_IDS.firstName] ??
          rawValues.get(REQUIRED_CATEGORY_IDS.firstName) ??
          '';
        const lastNameValue =
          categoryValues[REQUIRED_CATEGORY_IDS.lastName] ??
          rawValues.get(REQUIRED_CATEGORY_IDS.lastName) ??
          '';
        const birthDateValue =
          categoryValues[REQUIRED_CATEGORY_IDS.birthDate] ??
          rawValues.get(REQUIRED_CATEGORY_IDS.birthDate) ??
          '';

        const identifierData = buildContactIdentifierData(
          firstNameValue,
          lastNameValue,
          birthDateValue,
        );
        if (!identifierData) {
          errors.push({
            row: rowIndex + 1,
            categoryId: REQUIRED_CATEGORY_IDS.identifier,
            message: 'Impossible de calculer l’identifiant : vérifiez le prénom, le nom et la date de naissance.',
          });
          errorRows.add(rowIndex + 1);
          continue;
        }

        categoryValues[REQUIRED_CATEGORY_IDS.firstName] = toTrimmedString(firstNameValue);
        categoryValues[REQUIRED_CATEGORY_IDS.lastName] = toTrimmedString(lastNameValue);
        categoryValues[REQUIRED_CATEGORY_IDS.birthDate] = identifierData.normalizedDate;
        categoryValues[REQUIRED_CATEGORY_IDS.identifier] = identifierData.identifier;

        const existingContact = existingContactsByIdentifier.get(identifierData.identifier);
        if (existingContact) {
          const targetValues =
            existingContact.categoryValues && typeof existingContact.categoryValues === 'object'
              ? existingContact.categoryValues
              : (existingContact.categoryValues = {});
          let contactModified = false;

          normalizedMapping.forEach((mapping) => {
            if (REQUIRED_CATEGORY_ID_SET.has(mapping.categoryId)) {
              return;
            }
            const rawValue = rawValues.get(mapping.categoryId) || '';
            if (!rawValue) {
              return;
            }
            const newValue =
              categoryValues[mapping.categoryId] !== undefined
                ? categoryValues[mapping.categoryId]
                : rawValue;
            const previousValue = targetValues[mapping.categoryId] ?? '';
            if (previousValue.toString() !== newValue.toString()) {
              targetValues[mapping.categoryId] = newValue;
              contactModified = true;
            }
          });

          if (contactModified) {
            const updatedDerivedName = buildDisplayNameFromCategories(
              targetValues,
              categoriesById,
            );
            if (updatedDerivedName) {
              existingContact.fullName = updatedDerivedName;
              existingContact.displayName = updatedDerivedName;
            }
            existingContact.updatedAt = new Date().toISOString();
            updatedContactsCount += 1;
          }
          continue;
        }

        const nowIso = new Date().toISOString();
        const displayName =
          buildDisplayNameFromCategories(categoryValues, categoriesById) || 'Contact sans nom';
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
        existingContactsByIdentifier.set(identifierData.identifier, newContact);
        importedCount += 1;
      }

      const hasDataChanges = importedCount > 0 || updatedContactsCount > 0;
      if (hasDataChanges) {
        data.lastUpdated = new Date().toISOString();
        updateMetricsFromContacts();
        saveDataForUser(currentUser, data);
        renderMetrics();
        renderContacts();
      }

      if (importedCount > 0 || updatedContactsCount > 0) {
        notifyDataChanged('contacts', {
          reason: 'import',
          importedCount,
          updatedCount: updatedContactsCount,
          skippedEmptyCount,
          totalRows,
          errorCount: errorRows.size,
          fileName,
        });
      }

      return {
        importedCount,
        mergedCount: 0,
        skippedEmptyCount,
        totalRows,
        errorCount: errorRows.size,
        errors,
        fileName,
        duplicatesDetected: 0,
        autoMergeApplied: false,
        updatedCount: updatedContactsCount,
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

      let categoryId = '';
      if (typeof baseTask.categoryId === 'string' && baseTask.categoryId.trim()) {
        categoryId = baseTask.categoryId.trim();
      } else if (typeof baseTask.category === 'string' && baseTask.category.trim()) {
        categoryId = baseTask.category.trim();
      }

      const attachment = normalizeTaskAttachment(baseTask.attachment);

      let calendarEventId = '';
      if (
        typeof baseTask.calendarEventId === 'string' &&
        baseTask.calendarEventId.trim() !== ''
      ) {
        calendarEventId = baseTask.calendarEventId.trim();
      }

      let isArchived = false;
      if (typeof baseTask.isArchived === 'boolean') {
        isArchived = baseTask.isArchived;
      } else if (typeof baseTask.status === 'string') {
        isArchived = baseTask.status.trim().toLowerCase() === 'archived';
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
        categoryId,
        attachment,
        calendarEventId,
        isArchived,
      };
    }

    function normalizeTaskAttachment(rawAttachment) {
      const baseAttachment =
        rawAttachment && typeof rawAttachment === 'object' ? rawAttachment : {};

      let dataUrl = '';
      const dataUrlCandidates = [
        typeof baseAttachment.dataUrl === 'string' ? baseAttachment.dataUrl.trim() : '',
        typeof baseAttachment.data === 'string' ? baseAttachment.data.trim() : '',
      ];

      for (const candidate of dataUrlCandidates) {
        if (candidate && candidate.startsWith('data:')) {
          dataUrl = candidate;
          break;
        }
      }

      let url = '';
      const urlCandidates = [
        typeof baseAttachment.url === 'string' ? baseAttachment.url.trim() : '',
        typeof baseAttachment.href === 'string' ? baseAttachment.href.trim() : '',
        typeof baseAttachment.link === 'string' ? baseAttachment.link.trim() : '',
      ];

      for (const candidate of urlCandidates) {
        if (candidate) {
          url = candidate;
          break;
        }
      }

      if (!dataUrl && !url) {
        return null;
      }

      let name = '';
      if (typeof baseAttachment.name === 'string') {
        name = baseAttachment.name.trim();
      }

      let type = '';
      if (typeof baseAttachment.type === 'string') {
        type = baseAttachment.type.trim();
      }

      let size = 0;
      if (typeof baseAttachment.size === 'number' && Number.isFinite(baseAttachment.size)) {
        size = baseAttachment.size;
      } else if (typeof baseAttachment.size === 'string' && baseAttachment.size.trim()) {
        const parsed = Number(baseAttachment.size);
        if (Number.isFinite(parsed)) {
          size = parsed;
        }
      }

      let uploadedAt = '';
      if (
        typeof baseAttachment.uploadedAt === 'string' &&
        !Number.isNaN(new Date(baseAttachment.uploadedAt).getTime())
      ) {
        uploadedAt = baseAttachment.uploadedAt;
      } else {
        uploadedAt = new Date().toISOString();
      }

      return {
        name,
        type,
        size,
        dataUrl,
        url,
        uploadedAt,
      };
    }

    function normalizeTaskCategory(rawCategory) {
      const baseCategory = rawCategory && typeof rawCategory === 'object' ? rawCategory : {};

      let id = '';
      if (typeof baseCategory.id === 'string' && baseCategory.id.trim()) {
        id = baseCategory.id.trim();
      } else if (typeof baseCategory.identifier === 'string' && baseCategory.identifier.trim()) {
        id = baseCategory.identifier.trim();
      } else {
        id = generateId('task-category');
      }

      let name = '';
      if (typeof baseCategory.name === 'string') {
        name = baseCategory.name.trim();
      } else if (typeof baseCategory.label === 'string') {
        name = baseCategory.label.trim();
      }
      if (!name) {
        name = 'Catégorie';
      }

      const color = normalizeTaskCategoryColor(baseCategory.color);

      let createdAt = '';
      if (
        typeof baseCategory.createdAt === 'string' &&
        !Number.isNaN(new Date(baseCategory.createdAt).getTime())
      ) {
        createdAt = baseCategory.createdAt;
      } else {
        createdAt = new Date().toISOString();
      }

      return {
        id,
        name,
        color,
        createdAt,
      };
    }

    function normalizeTaskCategoryColor(rawColor) {
      if (typeof rawColor === 'string' && rawColor.trim()) {
        return normalizeTaskColor(rawColor);
      }

      return DEFAULT_TASK_CATEGORY_COLOR;
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

    function normalizeTeamChatMessage(rawMessage) {
      const baseMessage = rawMessage && typeof rawMessage === 'object' ? rawMessage : {};

      const content =
        typeof baseMessage.content === 'string'
          ? baseMessage.content.trim().slice(0, TEAM_CHAT_MAX_MESSAGE_LENGTH)
          : '';

      if (!content) {
        return null;
      }

      const author =
        typeof baseMessage.author === 'string' && baseMessage.author.trim()
          ? baseMessage.author.trim()
          : '';

      let createdAt = '';
      if (
        typeof baseMessage.createdAt === 'string' &&
        !Number.isNaN(new Date(baseMessage.createdAt).getTime())
      ) {
        createdAt = new Date(baseMessage.createdAt).toISOString();
      } else if (
        typeof baseMessage.timestamp === 'string' &&
        !Number.isNaN(new Date(baseMessage.timestamp).getTime())
      ) {
        createdAt = new Date(baseMessage.timestamp).toISOString();
      } else {
        createdAt = new Date().toISOString();
      }

      const id =
        typeof baseMessage.id === 'string' && baseMessage.id.trim()
          ? baseMessage.id.trim()
          : generateId('chat');

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
      const normalizedSubscriptionKey =
        typeof base.subscription === 'string' ? base.subscription.trim().toLowerCase() : '';
      const subscriptionPlan =
        normalizedSubscriptionKey && SUBSCRIPTION_PLANS[normalizedSubscriptionKey]
          ? SUBSCRIPTION_PLANS[normalizedSubscriptionKey]
          : null;
      base.subscription = subscriptionPlan ? subscriptionPlan.id : DEFAULT_SUBSCRIPTION_ID;

      const normalizedTeams = [];
      const seenTeamIds = new Set();
      if (Array.isArray(base.teams)) {
        base.teams.forEach((entry) => {
          const normalized = normalizeUserTeam(entry);
          if (!normalized) {
            return;
          }

          let candidateId = normalized.id;
          while (!candidateId || seenTeamIds.has(candidateId)) {
            candidateId = generateId('user-team');
          }

          normalized.id = candidateId;
          seenTeamIds.add(candidateId);
          normalizedTeams.push(normalized);
        });
      }

      if (normalizedTeams.length === 0) {
        normalizedTeams.push(createDefaultUserTeam());
      }

      base.teams = normalizedTeams;

      const candidateCurrentTeamId =
        typeof base.currentTeamId === 'string' ? base.currentTeamId.trim() : '';
      const hasCurrentTeam = normalizedTeams.some((team) => team.id === candidateCurrentTeamId);
      base.currentTeamId = hasCurrentTeam ? candidateCurrentTeamId : normalizedTeams[0].id;
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

      if (!Array.isArray(base.keywords)) {
        base.keywords = [];
      }

      if (!Array.isArray(base.contacts)) {
        base.contacts = [];
      }

      enforceRequiredContactCategories(base);
      base.categories = sortCategoriesForDisplay(base.categories);
      base.categories.forEach((category, index) => {
        category.order = index;
      });

      if (!Array.isArray(base.events)) {
        base.events = [];
      } else {
        base.events = base.events
          .filter((item) => item && typeof item === 'object')
          .map((item) => normalizeCalendarEvent(item))
          .sort(compareCalendarEvents);
      }

      if (!Array.isArray(base.taskCategories)) {
        base.taskCategories = [];
      } else {
        const normalizedTaskCategories = base.taskCategories
          .filter((item) => item && typeof item === 'object')
          .map((item) => normalizeTaskCategory(item))
          .filter((item) => item && item.id);

        const uniqueTaskCategories = new Map();
        normalizedTaskCategories.forEach((category) => {
          if (category && category.id && !uniqueTaskCategories.has(category.id)) {
            uniqueTaskCategories.set(category.id, category);
          }
        });

        base.taskCategories = Array.from(uniqueTaskCategories.values()).sort((a, b) =>
          a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }),
        );
      }

      const taskCategoryIdSet = new Set(
        Array.isArray(base.taskCategories)
          ? base.taskCategories.map((category) => category.id)
          : [],
      );

      if (!Array.isArray(base.tasks)) {
        base.tasks = [];
      } else {
        base.tasks = base.tasks
          .filter((item) => item && typeof item === 'object')
          .map((item) => normalizeTask(item))
          .sort(compareTasks);

        base.tasks.forEach((task) => {
          if (!task || typeof task !== 'object') {
            return;
          }

          if (typeof task.categoryId !== 'string' || !taskCategoryIdSet.has(task.categoryId)) {
            task.categoryId = '';
          }

          task.isArchived = Boolean(task.isArchived);

          if (!task.attachment || typeof task.attachment !== 'object') {
            task.attachment = null;
            return;
          }

          const hasDataUrl =
            typeof task.attachment.dataUrl === 'string' && task.attachment.dataUrl;
          const hasUrl = typeof task.attachment.url === 'string' && task.attachment.url;
          if (!hasDataUrl && !hasUrl) {
            task.attachment = null;
          }
        });
      }

      if (!Array.isArray(base.teamChatMessages)) {
        base.teamChatMessages = [];
      } else {
        base.teamChatMessages = base.teamChatMessages
          .map((item) => normalizeTeamChatMessage(item))
          .filter((item) => item && item.content)
          .slice(-TEAM_CHAT_HISTORY_LIMIT);
      }

      const categoriesById = buildCategoryMap(base);

      if (!Array.isArray(base.savedSearches)) {
        base.savedSearches = [];
      } else {
        const normalizedSavedSearches = [];
        const seenIds = new Set();

        base.savedSearches.forEach((item) => {
          const normalized = normalizeSavedSearch(item, categoriesById);
          if (!normalized) {
            return;
          }

          let identifier = normalized.id;
          while (!identifier || seenIds.has(identifier)) {
            identifier = generateId('saved-search');
            normalized.id = identifier;
          }

          seenIds.add(identifier);
          normalizedSavedSearches.push(normalized);
        });

        normalizedSavedSearches.sort((a, b) => {
          const diff = getSavedSearchTimestamp(b) - getSavedSearchTimestamp(a);
          if (diff !== 0) {
            return diff;
          }
          return a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' });
        });

        base.savedSearches = normalizedSavedSearches;
      }
      
      const safeMap = (arr, normalizer, label) => {
        if (!Array.isArray(arr)) return [];
        const isFn = typeof normalizer === 'function';
        return arr.map((item, idx) => {
          try {
            return isFn ? normalizer(item) : item;
          } catch (e) {
            console.error(`[${label}] normalize failed at index ${idx}:`, item, e);
            return null;
          }
        }).filter(Boolean);
      };

      base.emailTemplates = safeMap(
        base.emailTemplates,
        (typeof normalizeEmailTemplateRecord === 'function' ? normalizeEmailTemplateRecord : null),
        'emailTemplates'
      );

      base.emailCampaigns = safeMap(
        base.emailCampaigns,
        (typeof normalizeEmailCampaignRecord === 'function' ? normalizeEmailCampaignRecord : null),
        'emailCampaigns'
      );


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

        ensureContactIdentifier(normalizedCategoryValues);
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

    function normalizeUserTeam(rawTeam) {
      if (!rawTeam || typeof rawTeam !== 'object') {
        return null;
      }

      const normalizedId =
        typeof rawTeam.id === 'string' && rawTeam.id.trim() ? rawTeam.id.trim() : '';
      const name =
        getFirstNonEmptyString([rawTeam.name, rawTeam.label, rawTeam.title]) ||
        'Équipe sans nom';
      const role =
        getFirstNonEmptyString([rawTeam.role, rawTeam.position, rawTeam.function]) ||
        'Membre';

      return {
        id: normalizedId,
        name,
        role,
      };
    }

    function createDefaultUserTeam() {
      return {
        id: generateId('user-team'),
        name: 'Équipe principale',
        role: 'Membre',
      };
    }

    function createEmptyAdvancedFilters() {
      return {
        categories: {},
        keywords: [],
        keywordMode: (window.KEYWORD_FILTER_MODE_ALL || 'all'),
      };
    }

    function cloneAdvancedFilters(source) {
      const result = createEmptyAdvancedFilters();
      if (!source || typeof source !== 'object') {
        return result;
      }

      const rawCategories =
        source.categories && typeof source.categories === 'object' ? source.categories : {};
      Object.entries(rawCategories).forEach(([categoryId, filter]) => {
        if (!categoryId || !filter || typeof filter !== 'object') {
          return;
        }

        const type = CATEGORY_TYPES.has(filter.type) ? filter.type : 'text';
        const rawValueSource = filter.rawValue != null ? filter.rawValue.toString() : '';
        const rawValue = type === 'text' ? rawValueSource.trim() : rawValueSource;
        if (!rawValue) {
          return;
        }

        let normalizedValue = '';
        if (typeof filter.normalizedValue === 'string' && filter.normalizedValue) {
          normalizedValue = filter.normalizedValue;
        } else if (type === 'text') {
          normalizedValue = rawValue.toLowerCase();
        } else {
          normalizedValue = rawValue;
        }

        result.categories[categoryId] = {
          type,
          rawValue,
          normalizedValue,
        };
      });

      const rawKeywords = Array.isArray(source.keywords) ? source.keywords : [];
      const ANY = window.KEYWORD_FILTER_MODE_ANY || 'any';
      const ALL = window.KEYWORD_FILTER_MODE_ALL || 'all';

      // keywords
      const items = Array.isArray(rawKeywords) ? rawKeywords : [];
      const keywordSet = new Set();
      items.forEach((v) => {
        if (typeof v === 'string') {
          const t = v.trim();
          if (t) keywordSet.add(t);
        }
      });
      result.keywords = Array.from(keywordSet);

      // mode
      const mode = (source && typeof source.keywordMode === 'string') ? source.keywordMode : ALL;
      result.keywordMode = (mode === ANY) ? ANY : ALL;

      return result;
    }
    
    function getFirstNonEmptyString(candidates) {
      if (!Array.isArray(candidates)) {
        return '';
      }
      for (const candidate of candidates) {
        if (typeof candidate === 'string') {
          const value = candidate.trim();
          if (value) {
            return value;
          }
        }
        if (typeof candidate === 'number' && Number.isFinite(candidate)) {
          const value = candidate.toString();
          if (value) {
            return value;
          }
        }
      }
      return '';
    }

    function selectValidIsoDate(...candidates) {
      const values = Array.isArray(candidates) ? candidates : [];
      for (const candidate of values) {
        if (candidate instanceof Date && !Number.isNaN(candidate.getTime())) {
          return candidate.toISOString();
        }
        if (typeof candidate === 'number' && Number.isFinite(candidate)) {
          const date = new Date(candidate);
          if (!Number.isNaN(date.getTime())) {
            return date.toISOString();
          }
        }
        if (typeof candidate === 'string') {
          const value = candidate.trim();
          if (value && !Number.isNaN(Date.parse(value))) {
            return value;
          }
        }
      }
      return '';
    }

    function ensureUniqueEmailBlockId(existingId, seenIds) {
      const usedIds = seenIds instanceof Set ? seenIds : new Set();
      let candidateId = typeof existingId === 'string' ? existingId.trim() : '';
      if (!candidateId) {
        candidateId = generateId('email-block');
      }
      while (usedIds.has(candidateId)) {
        candidateId = generateId('email-block');
      }
      usedIds.add(candidateId);
      return candidateId;
    }

    function normalizeEmailTemplateRecord(rawTemplate) {
      if (!rawTemplate || typeof rawTemplate !== 'object') {
        return null;
      }

      const nowIso = new Date().toISOString();
      const normalized = {
        id: '',
        name: '',
        subject: '',
        blocks: [],
        createdAt: '',
        updatedAt: '',
      };

      normalized.id =
        getFirstNonEmptyString([rawTemplate.id, rawTemplate.templateId, rawTemplate.identifier]) ||
        generateId('email-template');

      normalized.name =
        getFirstNonEmptyString([rawTemplate.name, rawTemplate.title, rawTemplate.label]) ||
        'Modèle sans titre';

      normalized.subject = getFirstNonEmptyString([
        rawTemplate.subject,
        rawTemplate.subjectLine,
        rawTemplate.subject_line,
        rawTemplate.title,
      ]);

      const blocksSource = Array.isArray(rawTemplate.blocks)
        ? rawTemplate.blocks
        : Array.isArray(rawTemplate.content)
        ? rawTemplate.content
        : [];

      const seenBlockIds = new Set();
      blocksSource.forEach((block) => {
        if (block == null) {
          return;
        }

        let normalizedBlock = null;
        if (typeof block === 'string') {
          normalizedBlock = createEmailTemplateBlock('paragraph', { text: block });
        } else if (typeof block === 'object') {
          normalizedBlock = cloneEmailTemplateBlock(block, { preserveId: true });
        }

        if (!normalizedBlock || typeof normalizedBlock !== 'object') {
          return;
        }

        normalizedBlock.id = ensureUniqueEmailBlockId(normalizedBlock.id, seenBlockIds);
        normalized.blocks.push(normalizedBlock);
      });

      normalized.createdAt = selectValidIsoDate(
        rawTemplate.createdAt,
        rawTemplate.created_at,
        rawTemplate.dateCreated,
        rawTemplate.createdOn,
      );
      if (!normalized.createdAt) {
        normalized.createdAt = nowIso;
      }

      normalized.updatedAt = selectValidIsoDate(
        rawTemplate.updatedAt,
        rawTemplate.updated_at,
        rawTemplate.dateUpdated,
        rawTemplate.modifiedAt,
        rawTemplate.lastUpdated,
      );
      if (!normalized.updatedAt) {
        normalized.updatedAt = normalized.createdAt;
      }

      return normalized;
    }

    function normalizeEmailCampaignRecord(rawCampaign) {
      if (!rawCampaign || typeof rawCampaign !== 'object') {
        return null;
      }

      const nowIso = new Date().toISOString();
      const normalized = {
        id: '',
        name: '',
        templateId: '',
        templateName: '',
        savedSearchId: '',
        savedSearchName: '',
        senderName: '',
        senderEmail: '',
        subject: '',
        sentAt: '',
        createdAt: '',
        recipientCount: 0,
        recipients: [],
        templateSnapshot: [],
      };

      normalized.id =
        getFirstNonEmptyString([rawCampaign.id, rawCampaign.campaignId, rawCampaign.identifier]) ||
        generateId('email-campaign');

      normalized.name =
        getFirstNonEmptyString([rawCampaign.name, rawCampaign.title, rawCampaign.label]) ||
        'Campagne sans nom';

      normalized.templateId = getFirstNonEmptyString([
        rawCampaign.templateId,
        rawCampaign.template_id,
        rawCampaign.template && rawCampaign.template.id,
      ]);

      normalized.templateName = getFirstNonEmptyString([
        rawCampaign.templateName,
        rawCampaign.template && rawCampaign.template.name,
      ]);
      if (!normalized.templateName && normalized.templateId) {
        normalized.templateName = 'Modèle sans titre';
      }

      normalized.savedSearchId = getFirstNonEmptyString([
        rawCampaign.savedSearchId,
        rawCampaign.saved_search_id,
        rawCampaign.savedSearch && rawCampaign.savedSearch.id,
      ]);

      normalized.savedSearchName = getFirstNonEmptyString([
        rawCampaign.savedSearchName,
        rawCampaign.savedSearch && rawCampaign.savedSearch.name,
      ]);
      if (!normalized.savedSearchName && normalized.savedSearchId) {
        normalized.savedSearchName = 'Recherche sauvegardée';
      }

      normalized.senderName = getFirstNonEmptyString([
        rawCampaign.senderName,
        rawCampaign.sender_name,
        rawCampaign.fromName,
      ]);

      const senderEmailCandidate = getFirstNonEmptyString([
        rawCampaign.senderEmail,
        rawCampaign.sender_email,
        rawCampaign.from,
        rawCampaign.fromEmail,
      ]);
      normalized.senderEmail = isValidEmail(senderEmailCandidate) ? senderEmailCandidate : '';

      normalized.subject = getFirstNonEmptyString([
        rawCampaign.subject,
        rawCampaign.subjectLine,
        rawCampaign.subject_line,
        rawCampaign.title,
      ]);

      normalized.sentAt = selectValidIsoDate(
        rawCampaign.sentAt,
        rawCampaign.sent_at,
        rawCampaign.dispatchedAt,
      );
      if (!normalized.sentAt) {
        normalized.sentAt = nowIso;
      }

      normalized.createdAt = selectValidIsoDate(
        rawCampaign.createdAt,
        rawCampaign.created_at,
        rawCampaign.queuedAt,
        normalized.sentAt,
      );
      if (!normalized.createdAt) {
        normalized.createdAt = normalized.sentAt;
      }

      const rawRecipients = Array.isArray(rawCampaign.recipients)
        ? rawCampaign.recipients
        : Array.isArray(rawCampaign.contacts)
        ? rawCampaign.contacts
        : [];

      const recipients = [];
      const seenEmails = new Set();

      rawRecipients.forEach((entry) => {
        if (entry == null) {
          return;
        }

        if (typeof entry === 'string') {
          const email = entry.trim();
          const normalizedEmail = isValidEmail(email) ? email : '';
          if (!normalizedEmail) {
            return;
          }
          const emailKey = normalizedEmail.toLowerCase();
          if (seenEmails.has(emailKey)) {
            return;
          }
          seenEmails.add(emailKey);
          recipients.push({ contactId: '', email: normalizedEmail, name: '' });
          return;
        }

        if (typeof entry !== 'object') {
          return;
        }

        const email = getFirstNonEmptyString([entry.email, entry.mail, entry.address]);
        const normalizedEmail = isValidEmail(email) ? email : '';
        if (!normalizedEmail) {
          return;
        }
        const emailKey = normalizedEmail.toLowerCase();
        if (seenEmails.has(emailKey)) {
          return;
        }
        seenEmails.add(emailKey);

        const contactId = getFirstNonEmptyString([entry.contactId, entry.contact_id, entry.id]);
        const name = getFirstNonEmptyString([entry.name, entry.fullName, entry.label]);

        recipients.push({ contactId, email: normalizedEmail, name });
      });

      normalized.recipients = recipients;

      let recipientCount = Number(rawCampaign.recipientCount);
      if (!Number.isFinite(recipientCount) || recipientCount < 0) {
        const altCount = Number(rawCampaign.recipientsCount);
        if (Number.isFinite(altCount) && altCount >= 0) {
          recipientCount = altCount;
        } else {
          recipientCount = recipients.length;
        }
      }
      recipientCount = Math.max(0, Math.round(recipientCount));
      normalized.recipientCount = recipientCount;

      const snapshotSource = Array.isArray(rawCampaign.templateSnapshot)
        ? rawCampaign.templateSnapshot
        : Array.isArray(rawCampaign.template && rawCampaign.template.blocks)
        ? rawCampaign.template.blocks
        : Array.isArray(rawCampaign.blocks)
        ? rawCampaign.blocks
        : [];

      const snapshot = [];
      const snapshotIds = new Set();
      snapshotSource.forEach((block) => {
        if (block == null) {
          return;
        }

        let normalizedBlock = null;
        if (typeof block === 'string') {
          normalizedBlock = createEmailTemplateBlock('paragraph', { text: block });
        } else if (typeof block === 'object') {
          normalizedBlock = cloneEmailTemplateBlock(block, { preserveId: true });
        }

        if (!normalizedBlock || typeof normalizedBlock !== 'object') {
          return;
        }

        normalizedBlock.id = ensureUniqueEmailBlockId(normalizedBlock.id, snapshotIds);
        snapshot.push(normalizedBlock);
      });

      normalized.templateSnapshot = snapshot;

      return normalized;
    }

    function normalizeSavedSearch(rawSavedSearch, categoriesById = buildCategoryMap()) {
      if (!rawSavedSearch || typeof rawSavedSearch !== 'object') {
        return null;
      }

      const normalized = {
        id: '',
        name: '',
        searchTerm: '',
        advancedFilters: createEmptyAdvancedFilters(),
        createdAt: '',
        updatedAt: '',
      };

      if (typeof rawSavedSearch.id === 'string' && rawSavedSearch.id.trim()) {
        normalized.id = rawSavedSearch.id.trim();
      } else if (typeof rawSavedSearch.identifier === 'string' && rawSavedSearch.identifier.trim()) {
        normalized.id = rawSavedSearch.identifier.trim();
      } else {
        normalized.id = generateId('saved-search');
      }

      if (typeof rawSavedSearch.name === 'string' && rawSavedSearch.name.trim()) {
        normalized.name = rawSavedSearch.name.trim();
      } else if (typeof rawSavedSearch.label === 'string' && rawSavedSearch.label.trim()) {
        normalized.name = rawSavedSearch.label.trim();
      } else {
        normalized.name = 'Recherche sans nom';
      }

      const termCandidates = [rawSavedSearch.searchTerm, rawSavedSearch.term, rawSavedSearch.query];
      for (const candidate of termCandidates) {
        if (typeof candidate === 'string' && candidate.trim()) {
          normalized.searchTerm = candidate.trim();
          break;
        }
      }

      const filtersSource =
        rawSavedSearch.advancedFilters && typeof rawSavedSearch.advancedFilters === 'object'
          ? rawSavedSearch.advancedFilters
          : rawSavedSearch.filters;
      const clonedFilters = cloneAdvancedFilters(filtersSource || {});

      if (categoriesById instanceof Map) {
        Object.keys(clonedFilters.categories).forEach((categoryId) => {
          if (!categoriesById.has(categoryId)) {
            delete clonedFilters.categories[categoryId];
          }
        });
      }

      normalized.advancedFilters = clonedFilters;

      if (
        typeof rawSavedSearch.createdAt === 'string' &&
        !Number.isNaN(Date.parse(rawSavedSearch.createdAt))
      ) {
        normalized.createdAt = rawSavedSearch.createdAt;
      } else {
        normalized.createdAt = new Date().toISOString();
      }

      if (
        typeof rawSavedSearch.updatedAt === 'string' &&
        !Number.isNaN(Date.parse(rawSavedSearch.updatedAt))
      ) {
        normalized.updatedAt = rawSavedSearch.updatedAt;
      } else {
        normalized.updatedAt = normalized.createdAt;
      }

      return normalized;
    }

    function getSavedSearchTimestamp(savedSearch) {
      if (!savedSearch || typeof savedSearch !== 'object') {
        return 0;
      }

      const updated = Date.parse(savedSearch.updatedAt || '');
      if (!Number.isNaN(updated)) {
        return updated;
      }

      const created = Date.parse(savedSearch.createdAt || '');
      if (!Number.isNaN(created)) {
        return created;
      }

      return 0;
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

  function loadNewsStore() {
    try {
      const stored = window.localStorage.getItem(NEWS_STORE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        if (parsed && typeof parsed === 'object') {
          const articles = Array.isArray(parsed.articles) ? parsed.articles : [];
          const normalizedArticles = articles
            .map((article) => normalizeNewsArticle(article))
            .filter(Boolean)
            .sort((a, b) => {
              const dateA = typeof a.createdAt === 'string' ? a.createdAt : '';
              const dateB = typeof b.createdAt === 'string' ? b.createdAt : '';
              return dateB.localeCompare(dateA);
            });

          return { articles: normalizedArticles };
        }
      }
    } catch (error) {
      console.warn('Impossible de charger les nouveautés :', error);
    }

    return { articles: [] };
  }

  function saveNewsStore(store) {
    if (!store || typeof store !== 'object') {
      return;
    }

    const articles = Array.isArray(store.articles) ? store.articles : [];
    const payload = {
      articles: articles
        .map((article) => normalizeNewsArticle(article))
        .filter(Boolean)
        .sort((a, b) => {
          const dateA = typeof a.createdAt === 'string' ? a.createdAt : '';
          const dateB = typeof b.createdAt === 'string' ? b.createdAt : '';
          return dateB.localeCompare(dateA);
        }),
    };

    try {
      window.localStorage.setItem(NEWS_STORE_KEY, JSON.stringify(payload));
    } catch (error) {
      console.warn('Impossible de sauvegarder les nouveautés :', error);
    }
  }

  function normalizeNewsArticle(rawArticle) {
    if (!rawArticle || typeof rawArticle !== 'object') {
      return null;
    }

    const normalizedId =
      typeof rawArticle.id === 'string' && rawArticle.id.trim()
        ? rawArticle.id.trim()
        : '';

    const normalized = {
      id: normalizedId || generateId('news'),
      title: getFirstNonEmptyString([rawArticle.title, rawArticle.name]) || 'Annonce',
      summary:
        getFirstNonEmptyString([rawArticle.summary, rawArticle.resume, rawArticle.description]) ||
        '',
      content:
        typeof rawArticle.content === 'string' ? rawArticle.content.trim() : '',
      createdAt:
        selectValidIsoDate(rawArticle.createdAt, rawArticle.created_at) ||
        new Date().toISOString(),
      author:
        getFirstNonEmptyString([rawArticle.author, rawArticle.createdBy, rawArticle.owner]) || '',
    };

    return normalized;
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
                        taskCategories: Array.isArray(base.taskCategories) ? base.taskCategories : [],
                        tasks: Array.isArray(base.tasks) ? base.tasks : [],
                        teamChatMessages: Array.isArray(base.teamChatMessages)
                          ? base.teamChatMessages
                          : [],
                        savedSearches: Array.isArray(base.savedSearches)
                          ? base.savedSearches
                          : [],
                        emailTemplates: Array.isArray(base.emailTemplates)
                          ? base.emailTemplates
                          : [],
                        emailCampaigns: Array.isArray(base.emailCampaigns)
                          ? base.emailCampaigns
                          : [],
                        lastUpdated: base.lastUpdated || null,

                        // si jamais ces champs n’existaient pas encore
                        panelOwner: typeof base.panelOwner === 'string' ? base.panelOwner : '',
                        teamMembers: Array.isArray(base.teamMembers) ? base.teamMembers : [],
                        subscription:
                          typeof base.subscription === 'string'
                            ? base.subscription
                            : DEFAULT_SUBSCRIPTION_ID,
                        teams: Array.isArray(base.teams) ? base.teams : [],
                        currentTeamId:
                          typeof base.currentTeamId === 'string' ? base.currentTeamId : '',
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
        taskCategories: [],
        tasks: [],
        teamChatMessages: [],
        savedSearches: [],
        emailTemplates: [],
        emailCampaigns: [],
        lastUpdated: null,
        // Nouveaux champs persistés
        panelOwner: '',
        teamMembers: [],
        subscription: DEFAULT_SUBSCRIPTION_ID,
        teams: [],
        currentTeamId: '',
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

    const headerText = document.createElement('div');
    headerText.className = 'contact-header-text';

    const headerMain = document.createElement('div');
    headerMain.className = 'contact-header-main';

    const nameEl = document.createElement('h3');
    nameEl.className = 'contact-name';
    const typeBadge = document.createElement('span');
    typeBadge.className = 'contact-type-badge';
    headerMain.append(nameEl, typeBadge);

    const createdEl = document.createElement('span');
    createdEl.className = 'contact-created';

    headerText.append(headerMain, createdEl);

    const selectLabel = document.createElement('label');
    selectLabel.className = 'contact-select';
    const selectInput = document.createElement('input');
    selectInput.type = 'checkbox';
    selectInput.className = 'contact-select-checkbox';
    const srLabel = document.createElement('span');
    srLabel.className = 'sr-only';
    srLabel.textContent = 'Sélectionner ce contact';
    selectLabel.append(selectInput, srLabel);

    header.append(headerText, selectLabel);

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
    const statsButton = document.createElement('button');
    statsButton.type = 'button';
    statsButton.className = 'keyword-button keyword-button--stats';
    statsButton.dataset.action = 'stats';
    statsButton.textContent = 'Statistiques';
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
    actions.append(statsButton, editButton, deleteButton);

    listItem.append(keywordMain, actions);
    return listItem;
  }
})();