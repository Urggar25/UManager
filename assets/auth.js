import {
  ensureDefaultAdmins,
  hashPassword,
  loadActiveUser,
  loadUserStore,
  saveActiveUser,
  saveUserStore,
} from './storage.js';

document.addEventListener('DOMContentLoaded', () => {
  initializeAuth().catch((error) => {
    console.error('Erreur lors de l\'initialisation de la page de connexion :', error);
  });
});

function navigateToDashboard() {
  window.location.replace('dashboard.html');
}

async function initializeAuth() {
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
      document.getElementById('login-username')?.focus();
    });
  }

  function showRegisterView() {
    setSectionVisibility(loginSection, false);
    setSectionVisibility(registerSection, true);
    if (registerError) {
      registerError.textContent = '';
    }
    window.requestAnimationFrame(() => {
      document.getElementById('register-username')?.focus();
    });
  }

  function displayLoginError(message) {
    if (loginError) {
      loginError.textContent = message;
    }
  }

  showLoginView();

  showRegisterBtn?.addEventListener('click', () => {
    loginError && (loginError.textContent = '');
    registerError && (registerError.textContent = '');
    showRegisterView();
  });

  showLoginBtn?.addEventListener('click', () => {
    loginError && (loginError.textContent = '');
    registerError && (registerError.textContent = '');
    showLoginView();
  });

  loginForm?.addEventListener('submit', async (event) => {
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

  registerForm?.addEventListener('submit', async (event) => {
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

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}
