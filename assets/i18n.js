(function () {
  const SUPPORTED = new Set(['en', 'ru']);
  const DEFAULT_FALLBACK = 'en';
  const textNodeKeys = new WeakMap();
  const localeCache = new Map();
  let language = detectLanguage();
  let messages = {};

  function detectLanguage() {
    try {
      const stored = window.localStorage.getItem('lang');
      if (SUPPORTED.has(stored)) return stored;
    } catch (error) {
      // Ignore storage access issues in private or locked-down contexts.
    }

    const navLanguages = Array.isArray(navigator.languages) && navigator.languages.length
      ? navigator.languages
      : [navigator.language || ''];
    return navLanguages.some((item) => String(item).toLowerCase().includes('ru')) ? 'ru' : DEFAULT_FALLBACK;
  }

  function interpolate(value, params) {
    if (typeof value !== 'string') return value;
    return value.replace(/\{(\w+)\}/g, (_, key) => (
      Object.prototype.hasOwnProperty.call(params || {}, key) ? String(params[key]) : `{${key}}`
    ));
  }

  function lookup(key) {
    if (!key || !messages) return undefined;
    if (Object.prototype.hasOwnProperty.call(messages, key)) return messages[key];

    const dotIndex = key.indexOf('.');
    if (dotIndex > 0) {
      const namespace = key.slice(0, dotIndex);
      const rest = key.slice(dotIndex + 1);
      const group = messages[namespace];
      if (group && Object.prototype.hasOwnProperty.call(group, rest)) return group[rest];
    }

    return key.split('.').reduce((node, part) => {
      if (node && Object.prototype.hasOwnProperty.call(node, part)) return node[part];
      return undefined;
    }, messages);
  }

  function t(key, params = {}, fallback = key) {
    const value = lookup(key);
    if (value == null) return interpolate(fallback, params);
    return interpolate(value, params);
  }

  async function loadLocale(lang) {
    if (localeCache.has(lang)) return localeCache.get(lang);
    const response = await fetch(`./locales/${lang}.json`, { cache: 'no-cache' });
    if (!response.ok) {
      throw new Error(`Unable to load locale ${lang}: HTTP ${response.status}`);
    }
    const data = await response.json();
    localeCache.set(lang, data);
    return data;
  }

  function applyTextNode(node) {
    const raw = node.nodeValue || '';
    const trimmed = raw.trim();
    if (!trimmed) return;

    let sourceKey = textNodeKeys.get(node);
    if (!sourceKey) {
      if (!messages.text || !Object.prototype.hasOwnProperty.call(messages.text, trimmed)) return;
      sourceKey = trimmed;
      textNodeKeys.set(node, sourceKey);
    }

    const translated = messages.text?.[sourceKey] ?? sourceKey;
    const prefix = raw.match(/^\s*/)?.[0] || '';
    const suffix = raw.match(/\s*$/)?.[0] || '';
    node.nodeValue = `${prefix}${translated}${suffix}`;
  }

  function applyDataI18n(root) {
    const nodes = root.querySelectorAll?.('[data-i18n]') || [];
    nodes.forEach((node) => {
      node.textContent = t(node.dataset.i18n, {}, node.textContent);
    });
  }

  function walkText(root) {
    const walker = document.createTreeWalker(
      root,
      NodeFilter.SHOW_TEXT,
      {
        acceptNode(node) {
          const parent = node.parentElement;
          if (!parent) return NodeFilter.FILTER_REJECT;
          if (['SCRIPT', 'STYLE', 'TEXTAREA'].includes(parent.tagName)) return NodeFilter.FILTER_REJECT;
          return NodeFilter.FILTER_ACCEPT;
        },
      },
    );
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(applyTextNode);
  }

  function updateLanguageToggle() {
    const button = document.querySelector('[data-testid="language-toggle"]');
    if (!button) return;
    button.textContent = t('ui.languageToggle', {}, language === 'ru' ? 'EN' : 'RU');
    button.setAttribute('aria-label', t('ui.languageToggleLabel', {}, 'Toggle language'));
    button.setAttribute('title', t('ui.languageToggleLabel', {}, 'Toggle language'));
  }

  function apply(root = document) {
    document.documentElement.lang = language;
    document.title = t('meta.title', {}, document.title);
    applyDataI18n(root);
    walkText(root.body || root);
    updateLanguageToggle();
  }

  async function setLanguage(nextLanguage) {
    const normalized = SUPPORTED.has(nextLanguage) ? nextLanguage : DEFAULT_FALLBACK;
    language = normalized;
    messages = await loadLocale(language);
    try {
      window.localStorage.setItem('lang', language);
    } catch (error) {
      // Non-persistent language switching is still fine when storage is unavailable.
    }
    apply(document);
    window.dispatchEvent(new CustomEvent('app:i18n:change', { detail: { language } }));
    return language;
  }

  function getLanguage() {
    return language;
  }

  function toggleLanguage() {
    return setLanguage(language === 'ru' ? 'en' : 'ru');
  }

  const ready = loadLocale(language)
    .then((data) => {
      messages = data;
      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => apply(document), { once: true });
      } else {
        apply(document);
      }
      return language;
    })
    .catch((error) => {
      console.error(error);
      return language;
    });

  window.AppI18n = { t, getLanguage, setLanguage, toggleLanguage, apply, ready };

  document.addEventListener('click', (event) => {
    const button = event.target.closest?.('[data-testid="language-toggle"]');
    if (!button) return;
    toggleLanguage();
  });
}());
