(() => {
  const LANG_KEY = "xlance_static_lang";
  const THEME_KEY = "xlance_static_theme";
  const root = document.documentElement;

  function getStored(key) {
    try {
      return localStorage.getItem(key);
    } catch {
      return null;
    }
  }

  function setStored(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch {
      // Ignore storage failures; the current page can still use the detected value.
    }
  }

  function normalizeLang(value) {
    if (!value) return "";
    return String(value).toLowerCase().startsWith("en") ? "en" : "zh";
  }

  function detectLang() {
    const saved = normalizeLang(getStored(LANG_KEY));
    if (saved) return saved;
    const languages = navigator.languages && navigator.languages.length ? navigator.languages : [navigator.language || ""];
    const matched = languages.map((item) => String(item).toLowerCase()).find((item) => item.startsWith("zh") || item.startsWith("en"));
    const lang = matched && matched.startsWith("en") ? "en" : "zh";
    setStored(LANG_KEY, lang);
    return lang;
  }

  function normalizeTheme(value) {
    return value === "dark" || value === "light" ? value : "";
  }

  function detectTheme() {
    const saved = normalizeTheme(getStored(THEME_KEY));
    if (saved) return saved;
    const theme = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
    setStored(THEME_KEY, theme);
    return theme;
  }

  const lang = detectLang();
  root.lang = lang === "zh" ? "zh-CN" : "en";
  root.dataset.lang = lang;
  root.dataset.theme = detectTheme();
})();
