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

export async function ensureDefaultAdmins() {
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

export function loadUserStore() {
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

export function saveUserStore(store) {
  try {
    window.localStorage.setItem(USER_STORE_KEY, JSON.stringify(store));
  } catch (error) {
    console.warn('Impossible de sauvegarder les comptes utilisateurs :', error);
  }
}

export function saveActiveUser(username) {
  try {
    window.localStorage.setItem(ACTIVE_USER_KEY, username);
  } catch (error) {
    console.warn("Impossible d'enregistrer l'utilisateur actif :", error);
  }
}

export function loadActiveUser() {
  try {
    return window.localStorage.getItem(ACTIVE_USER_KEY);
  } catch (error) {
    console.warn("Impossible de charger l'utilisateur actif :", error);
    return null;
  }
}

export function clearActiveUser() {
  try {
    window.localStorage.removeItem(ACTIVE_USER_KEY);
  } catch (error) {
    console.warn("Impossible de réinitialiser l'utilisateur actif :", error);
  }
}

export function loadDataForUser(username) {
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

export function saveDataForUser(username, data) {
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

export function cloneDefaultData() {
  return {
    metrics: { ...defaultData.metrics },
    keywords: [],
    lastUpdated: null,
  };
}

export async function hashPassword(password) {
  const message = `umanager::${password}`;
  const cryptoObj =
    typeof window !== 'undefined' && window.crypto && window.crypto.subtle
      ? window.crypto
      : null;

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
  view.setUint32(buffer.length - 8, highBits, false);
  view.setUint32(buffer.length - 4, lowBits, false);

  const hash = new Uint32Array([
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
    0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
  ]);

  const w = new Uint32Array(64);

  for (let offset = 0; offset < buffer.length; offset += 64) {
    for (let i = 0; i < 16; i++) {
      w[i] = view.getUint32(offset + i * 4, false);
    }

    for (let i = 16; i < 64; i++) {
      const s0 = rightRotate(w[i - 15], 7) ^ rightRotate(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rightRotate(w[i - 2], 17) ^ rightRotate(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
    }

    let a = hash[0];
    let b = hash[1];
    let c = hash[2];
    let d = hash[3];
    let e = hash[4];
    let f = hash[5];
    let g = hash[6];
    let h = hash[7];

    for (let i = 0; i < 64; i++) {
      const S1 = rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (h + S1 + ch + k[i] + w[i]) >>> 0;
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

    hash[0] = (hash[0] + a) >>> 0;
    hash[1] = (hash[1] + b) >>> 0;
    hash[2] = (hash[2] + c) >>> 0;
    hash[3] = (hash[3] + d) >>> 0;
    hash[4] = (hash[4] + e) >>> 0;
    hash[5] = (hash[5] + f) >>> 0;
    hash[6] = (hash[6] + g) >>> 0;
    hash[7] = (hash[7] + h) >>> 0;
  }

  return Array.from(hash)
    .map((value) => value.toString(16).padStart(8, '0'))
    .join('');
}

function rightRotate(value, amount) {
  return (value >>> amount) | (value << (32 - amount));
}
