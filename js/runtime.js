(() => {
  const DATA = window.XLANCE_SITE;
  const LANG_KEY = "xlance_static_lang";
  const THEME_KEY = "xlance_static_theme";
  const HOME_NEWS_LIMIT = 6;
  const THEME_ORDER = ["dark", "light"];
  const GENERATED_JSON_PATHS = {
    members: "assets/data/members.json",
    papers: "assets/data/papers.json",
    news: "assets/data/news.json"
  };

  function canonicalizeHtmlPath() {
    if (!/^https?:$/.test(window.location.protocol)) return;
    const { pathname, search, hash } = window.location;
    if (!pathname.endsWith(".html")) return;
    const cleanPath = pathname.endsWith("/index.html")
      ? pathname.slice(0, -"index.html".length) || "/"
      : pathname.slice(0, -".html".length);
    window.history.replaceState(null, "", `${cleanPath}${search}${hash}`);
  }

  canonicalizeHtmlPath();

  const UI = {
    zh: {
      langName: "中文",
      switchTo: "EN",
      homeCtaJoin: "加入我们",
      visit: "查看原文",
      open: "查看详情",
      paperLink: "论文链接",
      projectLink: "项目主页",
      watchVideo: "观看视频",
      directions: "研究方向",
      latest: "实验室动态",
      papers: "论文",
      representativeWorks: "代表工作",
      paperList: "论文列表",
      paperFiles: "附件",
      people: "成员",
      projects: "项目",
      repoCatalog: "项目仓库",
      faculty: "教师",
      students: "在校生",
      alumni: "校友",
      postdoc: "博士后",
      phd: "博士",
      master: "硕士",
      undergrad: "本科生",
      xmart: "Xmart",
      xmartSessions: "Xmart 期数",
      about: "实验室简介",
      credo: "实验室信条",
      frontierLecture: "前沿讲座",
      youthForum: "学生论坛",
      joinUs: "加入我们",
      recruiting: "招募信息",
      contact: "联系方式",
      contactPerson: "茅懿琼老师",
      emailSubject: "邮件标题注明：",
      resume: "简历需包括：",
      recruitingText: "实验室长期面向本科生进行招生，欢迎对通用大语言模型、AI for Science、基于语音、NLP 和 CV 的交互式智能有强烈兴趣的大二、大三学生。期待具备优秀理论与工程实践能力，立志成为人工智能行业杰出工程师和科学家的同学加入我们。",
      resumeText: "基本信息、自我能力、项目经历、后续深造计划、成绩单、联系方式。",
      source: "内容来源",
      all: "全部",
      joined: "加入时间",
      memberNo: "编号",
      office: "地址",
      email: "邮箱",
      researchLabel: "研究方向",
      themeTitle: "切换明暗模式",
      themeLight: "浅色",
      themeDark: "暗色",
      followWechat: "关注实验室公众号",
      uniqueVisitorCount: "独立访客",
      pageViewCount: "本页访问",
      siteVisitCount: "全站访问",
      visitorCountLoading: "加载中",
      footer: ""
    },
    en: {
      langName: "English",
      switchTo: "中文",
      homeCtaJoin: "Join Us",
      visit: "Read original",
      open: "View details",
      paperLink: "Paper link",
      projectLink: "Project page",
      watchVideo: "Watch video",
      directions: "Research Areas",
      latest: "Lab News",
      papers: "Papers",
      representativeWorks: "Featured Works",
      paperList: "Paper List",
      paperFiles: "Files",
      people: "Members",
      projects: "Projects",
      repoCatalog: "Project Repositories",
      faculty: "Faculty",
      students: "Students",
      alumni: "Alumni",
      postdoc: "Postdocs",
      phd: "PhD",
      master: "Master",
      undergrad: "Undergraduates",
      xmart: "Xmart",
      xmartSessions: "Xmart Sessions",
      about: "About X-LANCE",
      credo: "Lab Credo",
      frontierLecture: "Frontier Lecture",
      youthForum: "Youth Forum",
      joinUs: "Join Us",
      recruiting: "Recruitment",
      contact: "Contact",
      contactPerson: "Yiqiong Mao",
      emailSubject: "Email subject:",
      resume: "Please include:",
      recruitingText: "We welcome second- and third-year undergraduate students who are strongly interested in general large language models, AI for Science, and interactive intelligence based on speech, NLP, and CV. We seek students with strong theoretical and engineering abilities who aspire to become outstanding engineers and scientists in AI.",
      resumeText: "Basic information, strengths, project experience, future study plan, transcript, and contact information.",
      source: "Content Source",
      all: "All",
      joined: "Joined",
      memberNo: "ID",
      office: "Address",
      email: "Email",
      researchLabel: "Research Areas",
      themeTitle: "Switch color mode",
      themeLight: "Light",
      themeDark: "Dark",
      followWechat: "Follow us on WeChat",
      uniqueVisitorCount: "Visitors",
      pageViewCount: "Page Views",
      siteVisitCount: "Total Visits",
      visitorCountLoading: "Loading",
      footer: ""
    }
  };

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function tr(value, lang) {
    if (value && typeof value === "object" && !Array.isArray(value)) {
      return value[lang] ?? value.zh ?? value.en ?? "";
    }
    return value ?? "";
  }

  function trList(value, lang) {
    const result = tr(value, lang);
    if (Array.isArray(result)) return result;
    return result ? [result] : [];
  }

  const IMAGE_PREFIX = "assets/img/";
  const WEBP_IMAGE_PREFIX = "assets/img-webp/";
  const RASTER_IMAGE_RE = /\.(?:png|jpe?g)$/i;
  const HIGH_RES_CACHE_KEY = "xlance_high_res_images_v1";
  const HIGH_RES_CACHE_LIMIT = 500;
  const WEBP_CONCURRENCY = 6;
  const HIGH_RES_CONCURRENCY = 2;
  const IMAGE_PLACEHOLDER_SRC = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";
  const VERCOUNT_SCRIPT_SRC = "https://events.vercount.one/js";
  const VERCOUNT_FIELDS = [
    { key: "pagePv", id: "vercount_value_page_pv", containerId: "vercount_container_page_pv", labelKey: "pageViewCount" },
    { key: "sitePv", id: "vercount_value_site_pv", containerId: "vercount_container_site_pv", labelKey: "siteVisitCount" },
    { key: "siteUv", id: "vercount_value_site_uv", containerId: "vercount_container_site_uv", labelKey: "uniqueVisitorCount" }
  ];
  let highResImageCache;
  let progressiveImageBatch = 0;
  let vercountLoaded = false;
  let vercountValues = {};
  let vercountObservers = [];

  function optimizedImageSrc(src) {
    const value = String(src || "");
    if (!value.startsWith(IMAGE_PREFIX) || !RASTER_IMAGE_RE.test(value)) return value;
    return value.replace(IMAGE_PREFIX, WEBP_IMAGE_PREFIX).replace(RASTER_IMAGE_RE, ".webp");
  }

  function isOptimizableImage(src) {
    const value = String(src || "");
    return value.startsWith(IMAGE_PREFIX) && RASTER_IMAGE_RE.test(value);
  }

  function absoluteImageSrc(src) {
    try {
      return new URL(String(src || ""), document.baseURI).href;
    } catch (_error) {
      return String(src || "");
    }
  }

  function sameImageSrc(left, right) {
    return absoluteImageSrc(left) === absoluteImageSrc(right);
  }

  function highResCache() {
    if (highResImageCache) return highResImageCache;
    try {
      const saved = JSON.parse(localStorage.getItem(HIGH_RES_CACHE_KEY) || "[]");
      highResImageCache = new Set(Array.isArray(saved) ? saved : []);
    } catch (_error) {
      highResImageCache = new Set();
    }
    return highResImageCache;
  }

  function markHighResCached(src) {
    const key = absoluteImageSrc(src);
    const cache = highResCache();
    cache.delete(key);
    cache.add(key);
    try {
      const items = Array.from(cache).slice(-HIGH_RES_CACHE_LIMIT);
      highResImageCache = new Set(items);
      localStorage.setItem(HIGH_RES_CACHE_KEY, JSON.stringify(items));
    } catch (_error) {
      // Private browsing or quota limits should not affect image rendering.
    }
  }

  function jsArg(value) {
    return escapeHtml(JSON.stringify(String(value || "")));
  }

  function attrsWithClass(attrs, extraClass) {
    const value = String(attrs || "").trim();
    if (!value) return `class="${extraClass}"`;
    if (/\bclass\s*=/.test(value)) {
      return value.replace(/\bclass=(["'])(.*?)\1/, (_match, quote, existing) => `class=${quote}${existing} ${extraClass}${quote}`);
    }
    return `${value} class="${extraClass}"`;
  }

  function eagerImageAttrs(attrs) {
    const value = String(attrs || "").trim();
    if (!value) return 'loading="eager"';
    if (/\bloading\s*=/.test(value)) {
      return value.replace(/\bloading=(["'])(.*?)\1/, (_match, quote) => `loading=${quote}eager${quote}`);
    }
    return `${value} loading="eager"`;
  }

  function imageTag(src, alt = "", attrs = "", fallbackSrc = "") {
    const original = String(src || "");
    if (!isOptimizableImage(original)) {
      const attrText = attrs ? ` ${attrs}` : "";
      return `<img src="${escapeHtml(original)}" alt="${escapeHtml(alt)}"${attrText} />`;
    }
    const optimized = optimizedImageSrc(original);
    const attrText = attrsWithClass(eagerImageAttrs(attrs), "progressive-image is-preview");
    return `<img src="${IMAGE_PLACEHOLDER_SRC}" alt="${escapeHtml(alt)}" ${attrText} data-webp-src="${escapeHtml(optimized)}" data-original-src="${escapeHtml(original)}" data-fallback-src="${escapeHtml(fallbackSrc)}" onload="window.xlanceImageLoaded(this)" onerror="window.xlanceImageFallback(this)" />`;
  }

  window.xlanceImageLoaded = function xlanceImageLoaded(image) {
    const original = image.dataset.originalSrc || "";
    const current = image.getAttribute("src") || "";
    if (original && sameImageSrc(current, original)) {
      markHighResCached(original);
      image.classList.remove("is-preview", "is-upgrading");
      image.classList.add("is-highres");
    }
  };

  window.xlanceImageFallback = function xlanceImageFallback(image) {
    const original = image.dataset.originalSrc || "";
    const webp = image.dataset.webpSrc || "";
    const fallback = image.dataset.fallbackSrc || "";
    const current = image.getAttribute("src") || "";

    if (webp && !sameImageSrc(current, webp) && image.dataset.webpFallbackTried !== "1") {
      image.dataset.webpFallbackTried = "1";
      image.src = webp;
      return;
    }

    if (original && !sameImageSrc(current, original) && image.dataset.originalFallbackTried !== "1") {
      image.dataset.originalFallbackTried = "1";
      image.src = original;
      return;
    }

    if (fallback && image.dataset.webpFinalFallbackTried !== "1") {
      const optimizedFallback = optimizedImageSrc(fallback);
      image.dataset.webpFinalFallbackTried = "1";
      if (optimizedFallback && current !== optimizedFallback) {
        image.src = optimizedFallback;
        return;
      }
    }

    if (fallback && image.dataset.originalFinalFallbackTried !== "1" && current !== fallback) {
      image.dataset.originalFinalFallbackTried = "1";
      image.src = fallback;
      return;
    }

    image.onerror = null;
  };

  function settleHighResLoad(done) {
    if (typeof done === "function") done();
  }

  function waitForCurrentImage(image, done) {
    if (image.complete) {
      settleHighResLoad(done);
      return;
    }
    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      image.removeEventListener("load", finish);
      image.removeEventListener("error", finish);
      settleHighResLoad(done);
    };
    image.addEventListener("load", finish);
    image.addEventListener("error", finish);
  }

  function upgradeProgressiveImage(image, done) {
    const original = image.dataset.originalSrc || "";
    if (!original || image.dataset.highResRequested === "1") {
      settleHighResLoad(done);
      return;
    }
    if (sameImageSrc(image.getAttribute("src") || "", original)) {
      if (image.complete) {
        if (image.naturalWidth > 0) {
          markHighResCached(original);
          image.classList.remove("is-preview", "is-upgrading");
          image.classList.add("is-highres");
        }
        settleHighResLoad(done);
      } else {
        waitForCurrentImage(image, () => {
          if (image.complete && image.naturalWidth > 0) {
            markHighResCached(original);
            image.classList.remove("is-preview", "is-upgrading");
            image.classList.add("is-highres");
          }
          settleHighResLoad(done);
        });
      }
      return;
    }

    image.dataset.highResRequested = "1";
    const highRes = new Image();
    highRes.decoding = "async";
    highRes.onload = () => {
      markHighResCached(original);
      const decoded = highRes.decode ? highRes.decode().catch(() => {}) : Promise.resolve();
      decoded.then(() => {
        if (!image.isConnected || sameImageSrc(image.getAttribute("src") || "", original)) {
          settleHighResLoad(done);
          return;
        }
        image.classList.add("is-upgrading");
        window.setTimeout(() => {
          if (!image.isConnected) {
            settleHighResLoad(done);
            return;
          }
          image.src = original;
          requestAnimationFrame(() => {
            image.classList.remove("is-preview", "is-upgrading");
            image.classList.add("is-highres");
            settleHighResLoad(done);
          });
        }, 140);
      });
    };
    highRes.onerror = () => {
      image.dataset.highResRequested = "0";
      settleHighResLoad(done);
    };
    highRes.src = original;
  }

  function loadWebpImage(image, done) {
    const webp = image.dataset.webpSrc || "";
    const original = image.dataset.originalSrc || "";
    const current = image.getAttribute("src") || "";
    if (!webp || !image.isConnected || sameImageSrc(current, webp) || sameImageSrc(current, original)) {
      settleHighResLoad(done);
      return;
    }

    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      image.removeEventListener("load", finish);
      image.removeEventListener("error", finish);
      settleHighResLoad(done);
    };

    image.addEventListener("load", finish);
    image.addEventListener("error", finish);
    image.dataset.webpRequested = "1";
    image.src = webp;
    if (sameImageSrc(image.getAttribute("src") || "", webp) && image.complete) finish();
  }

  function orderedImagesForViewport(images) {
    const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
    const current = [];
    const below = [];
    const above = [];
    images.forEach((image) => {
      const rect = image.getBoundingClientRect();
      if (rect.bottom >= 0 && rect.top <= viewportHeight) {
        current.push(image);
      } else if (rect.top > viewportHeight) {
        below.push(image);
      } else {
        above.push(image);
      }
    });
    return current.concat(below, above);
  }

  function runProgressiveImageStage(batch, images, concurrency, loadImage, done) {
    let index = 0;
    let active = 0;
    let finished = false;

    const pump = () => {
      if (batch !== progressiveImageBatch || finished) return;
      while (active < concurrency && index < images.length) {
        const image = images[index];
        index += 1;
        if (!image.isConnected) continue;
        active += 1;
        let settled = false;
        loadImage(image, () => {
          if (settled) return;
          settled = true;
          active = Math.max(0, active - 1);
          pump();
        });
      }

      if (active === 0 && index >= images.length) {
        finished = true;
        done();
      }
    };

    pump();
  }

  function startProgressiveImageStages(batch, images) {
    const ordered = orderedImagesForViewport(images);
    runProgressiveImageStage(batch, ordered, WEBP_CONCURRENCY, loadWebpImage, () => {
      runProgressiveImageStage(batch, ordered, HIGH_RES_CONCURRENCY, upgradeProgressiveImage, () => {});
    });
  }

  function scheduleProgressiveImageUpgrades(root = document) {
    const images = Array.from(root.querySelectorAll("img[data-original-src][data-webp-src]"));
    progressiveImageBatch += 1;
    const batch = progressiveImageBatch;
    images.forEach((image) => {
      if (sameImageSrc(image.getAttribute("src") || "", image.dataset.originalSrc || "")) {
        markHighResCached(image.dataset.originalSrc);
        image.classList.add("is-highres");
      }
    });
    requestAnimationFrame(() => {
      if (batch !== progressiveImageBatch) return;
      startProgressiveImageStages(batch, images);
    });
  }

  function applyMembersData(data) {
    if (data && typeof data === "object") window.XLANCE_MEMBERS = data;
  }

  function applyPapersData(data) {
    if (data && typeof data === "object") window.XLANCE_FULL_PAPERS = data;
  }

  function applyNewsData(data) {
    if (Array.isArray(data) && data.length) DATA.news = data;
  }

  async function loadGeneratedData() {
    applyMembersData(window.XLANCE_MEMBERS);
    applyPapersData(window.XLANCE_FULL_PAPERS);
    if (!window.fetch || window.location.protocol === "file:") return;
    try {
      const cacheBust = `?v=${Date.now()}`;
      const [membersResponse, papersResponse, newsResponse] = await Promise.all([
        fetch(`${GENERATED_JSON_PATHS.members}${cacheBust}`, { cache: "no-store" }),
        fetch(`${GENERATED_JSON_PATHS.papers}${cacheBust}`, { cache: "no-store" }),
        fetch(`${GENERATED_JSON_PATHS.news}${cacheBust}`, { cache: "no-store" })
      ]);
      if (membersResponse.ok) applyMembersData(await membersResponse.json());
      if (papersResponse.ok) applyPapersData(await papersResponse.json());
      if (newsResponse.ok) applyNewsData(await newsResponse.json());
    } catch (_error) {
      // data.js keeps a small fallback news list for direct file opening.
    }
  }

  function normalizeLang(value) {
    return value && value.toLowerCase().startsWith("en") ? "en" : "zh";
  }

  function detectLang() {
    const saved = localStorage.getItem(LANG_KEY);
    if (saved) return normalizeLang(saved);
    const languages = navigator.languages && navigator.languages.length ? navigator.languages : [navigator.language || ""];
    const matched = languages.map((item) => String(item).toLowerCase()).find((item) => item.startsWith("zh") || item.startsWith("en"));
    return matched && matched.startsWith("en") ? "en" : "zh";
  }

  function normalizeTheme(value) {
    return THEME_ORDER.includes(value) ? value : "";
  }

  function detectTheme() {
    const saved = normalizeTheme(localStorage.getItem(THEME_KEY));
    if (saved) return saved;
    const theme = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
    localStorage.setItem(THEME_KEY, theme);
    return theme;
  }

  function setLang(next) {
    localStorage.setItem(LANG_KEY, next);
    document.documentElement.lang = next === "zh" ? "zh-CN" : "en";
    document.documentElement.dataset.lang = next;
  }

  function setTheme(next) {
    const theme = normalizeTheme(next) || detectTheme();
    localStorage.setItem(THEME_KEY, theme);
    document.documentElement.dataset.theme = theme;
    currentTheme = theme;
  }

  function nextTheme(theme) {
    const index = THEME_ORDER.indexOf(normalizeTheme(theme));
    return THEME_ORDER[(index + 1) % THEME_ORDER.length];
  }

  function themeName(theme, lang) {
    const labels = {
      light: UI[lang].themeLight,
      dark: UI[lang].themeDark
    };
    return labels[normalizeTheme(theme)];
  }

  function themeIcon(theme) {
    if (normalizeTheme(theme) === "dark") {
      return `
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path d="M20 15.4A8.2 8.2 0 0 1 8.6 4a7.4 7.4 0 1 0 11.4 11.4Z"></path>
        </svg>
      `;
    }
    return `
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <circle cx="12" cy="12" r="4"></circle>
        <path d="M12 2v2.2M12 19.8V22M4.9 4.9l1.6 1.6M17.5 17.5l1.6 1.6M2 12h2.2M19.8 12H22M4.9 19.1l1.6-1.6M17.5 6.5l1.6-1.6"></path>
      </svg>
    `;
  }

  function pageId() {
    return document.body.dataset.page || "home";
  }

  function navActiveId(active) {
    if (active === "faculty" || active === "students") return "members";
    if (active === "resources" || active === "media") return "projects";
    return active;
  }

  function syncHeaderHeight() {
    const header = document.querySelector(".site-header");
    if (!header) return;
    document.documentElement.style.setProperty("--site-header-height", `${header.offsetHeight}px`);
  }

  function syncActiveNav() {
    const active = document.querySelector(".site-nav a.active");
    if (!active) return;
    active.scrollIntoView({ block: "nearest", inline: "center" });
  }

  function scrollToHashTarget() {
    if (!window.location.hash) return;
    const id = decodeURIComponent(window.location.hash.slice(1));
    const target = document.getElementById(id);
    if (!target) return;
    const headerHeight = document.querySelector(".site-header")?.offsetHeight || 70;
    const toc = target.classList.contains("paper-year-section") && window.matchMedia("(max-width: 860px)").matches
      ? document.querySelector(".paper-year-toc")
      : null;
    const offset = headerHeight + (toc ? toc.offsetHeight + 12 : 16);
    const top = Math.max(0, target.getBoundingClientRect().top + window.scrollY - offset);
    const root = document.documentElement;
    const previousBehavior = root.style.scrollBehavior;
    root.style.scrollBehavior = "auto";
    window.scrollTo({ top, left: 0, behavior: "auto" });
    root.style.scrollBehavior = previousBehavior;
  }

  function typeLabel(type, lang) {
    const labels = {
      paper: { zh: "论文", en: "Paper" },
      honor: { zh: "荣誉", en: "Honor" },
      result: { zh: "成果", en: "Result" },
      talk: { zh: "观点", en: "Talk" },
      community: { zh: "团队", en: "Community" },
      speech: { zh: "语音", en: "Speech" },
      multimodal: { zh: "多模态", en: "Multimodal" },
      science: { zh: "科学智能", en: "AI4Science" },
      agents: { zh: "智能体", en: "Agents" },
      nlp: { zh: "语言", en: "NLP" }
    };
    return tr(labels[type] || type, lang);
  }

  function renderHeader(active, lang, theme) {
    const activeId = navActiveId(active);
    const nav = DATA.nav.map((item) => `
      <a href="${escapeHtml(item.href)}" class="${item.id === activeId ? "active" : ""}">${escapeHtml(tr(item.label, lang))}</a>
    `).join("");
    const themeText = themeName(theme, lang);
    const header = document.getElementById("site-header");
    header.innerHTML = `
      <header class="site-header">
        <div class="wrap nav-bar">
          <a class="brand" href="/" aria-label="X-LANCE home">
            ${imageTag(logoAsset(lang), "X-LANCE")}
            <span>
              <strong>${escapeHtml(DATA.lab.shortName)}</strong>
              <small>${escapeHtml(tr(DATA.lab.subtitle, lang))}</small>
            </span>
          </a>
          <nav class="site-nav" aria-label="Primary navigation">${nav}</nav>
          <div class="header-actions">
            <button class="theme-switch" type="button" data-theme-toggle aria-label="${escapeHtml(`${UI[lang].themeTitle}: ${themeText}`)}" title="${escapeHtml(`${UI[lang].themeTitle}: ${themeText}`)}">${themeIcon(theme)}</button>
            <button class="lang-switch" type="button" data-lang-toggle>${UI[lang].switchTo}</button>
            <a class="lang-switch login-switch" href="https://x-lance.sjtu.edu.cn/docs/login">${lang === "zh" ? "登录" : "Login"}</a>
          </div>
        </div>
      </header>
    `;
  }

  function socialIcon(kind) {
    const icons = {
      github: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 2.5a9.5 9.5 0 0 0-3 18.5c.48.09.65-.2.65-.46v-1.7c-2.65.58-3.2-1.12-3.2-1.12-.43-1.1-1.05-1.4-1.05-1.4-.86-.58.07-.57.07-.57.95.07 1.45.97 1.45.97.85 1.45 2.22 1.03 2.76.79.09-.61.33-1.03.6-1.27-2.12-.24-4.35-1.06-4.35-4.72 0-1.04.37-1.9.98-2.57-.1-.24-.43-1.21.09-2.52 0 0 .8-.26 2.62.98A9.1 9.1 0 0 1 12 7.1c.82 0 1.65.11 2.42.33 1.82-1.24 2.61-.98 2.61-.98.53 1.31.2 2.28.1 2.52.6.67.97 1.53.97 2.57 0 3.67-2.24 4.48-4.37 4.71.34.3.65.9.65 1.82v2.7c0 .27.17.56.66.46A9.5 9.5 0 0 0 12 2.5Z" fill="currentColor"/></svg>',
      huggingface: '<img class="footer-brand-img footer-brand-img-huggingface" src="assets/img/logo/huggingface_logo-noborder.svg" alt="" aria-hidden="true" />',
      bilibili: '<img class="footer-brand-img footer-brand-img-bilibili" src="assets/img/logo/bilibili-favicon.ico" alt="" aria-hidden="true" />',
      external: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 4h6v6M20 4l-9 9M19 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1h5" fill="none"/></svg>'
    };
    return icons[kind] || icons.external;
  }

  function cacheVercountValue(field) {
    const value = document.getElementById(field.id)?.textContent?.trim();
    if (value && value !== UI.zh.visitorCountLoading && value !== UI.en.visitorCountLoading) {
      vercountValues[field.key] = value;
    }
  }

  function observeVercountValues() {
    vercountObservers.forEach((observer) => observer.disconnect());
    vercountObservers = [];
    VERCOUNT_FIELDS.forEach((field) => {
      const valueNode = document.getElementById(field.id);
      if (!valueNode) return;
      cacheVercountValue(field);
      const observer = new MutationObserver(() => cacheVercountValue(field));
      observer.observe(valueNode, { childList: true, characterData: true, subtree: true });
      vercountObservers.push(observer);
    });
  }

  function loadVercount() {
    if (vercountLoaded) return;
    vercountLoaded = true;
    const script = document.createElement("script");
    script.defer = true;
    script.src = VERCOUNT_SCRIPT_SRC;
    document.head.appendChild(script);
  }

  function renderFooter(lang) {
    const fallbackLinks = [{ href: DATA.lab.github, kind: "github", label: { zh: "GitHub", en: "GitHub" } }];
    const links = Array.isArray(DATA.lab.links) && DATA.lab.links.length ? DATA.lab.links : fallbackLinks;
    const footer = document.getElementById("site-footer");
    footer.innerHTML = `
      <footer class="site-footer">
        <div class="footer-aerial">
          ${imageTag(DATA.lab.assets.footerAerial, "", 'class="footer-aerial-image" loading="lazy" decoding="async"')}
          <div class="wrap footer-content">
            <div class="footer-copy">
              <h2>${lang === "zh" ? "上海交通大学X-LANCE实验室" : "X-LANCE Lab, SJTU"}</h2>
              <p>${escapeHtml(DATA.lab.email)} · ${escapeHtml(tr(DATA.lab.location, lang))}</p>
              ${DATA.lab.dlc ? `<a class="footer-dlc" href="${escapeHtml(DATA.lab.dlc.href)}" target="_blank" rel="noopener">${socialIcon("external")}<span>${escapeHtml(tr(DATA.lab.dlc.label, lang))}</span></a>` : ""}
              <div class="footer-links">
                ${links.map((link) => {
                  const href = link.href || link.url || "";
                  const label = tr(link.label, lang);
                  return `<a href="${escapeHtml(href)}"${href.startsWith("http") ? " target=\"_blank\" rel=\"noopener\"" : ""}>${socialIcon(link.kind)}<span>${escapeHtml(label)}</span></a>`;
                }).join("")}
              </div>
            </div>
            <div class="footer-qr">
              ${imageTag(DATA.lab.assets.wechat, UI[lang].followWechat, 'loading="lazy" decoding="async"')}
              <span>${escapeHtml(UI[lang].followWechat)}</span>
            </div>
            <div class="footer-visitor" aria-label="${escapeHtml(lang === "zh" ? "访问统计" : "Visit statistics")}">
              ${VERCOUNT_FIELDS.map((field) => `
                <div class="footer-visitor-item" id="${field.containerId}">
                  <span>${escapeHtml(UI[lang][field.labelKey])}</span>
                  <strong id="${field.id}">${escapeHtml(vercountValues[field.key] || UI[lang].visitorCountLoading)}</strong>
                </div>
              `).join("")}
            </div>
          </div>
        </div>
      </footer>
    `;
    observeVercountValues();
    loadVercount();
  }

  function pageHeroImage(page = pageId()) {
    return DATA.lab.pageHeroes?.[page] || DATA.lab.assets.hero;
  }

  function logoAsset(lang) {
    return lang === "en"
      ? DATA.lab.assets.logoEn || DATA.lab.assets.logo || DATA.lab.assets.logoName
      : DATA.lab.assets.logoZh || DATA.lab.assets.logo || DATA.lab.assets.logoName;
  }

  function heroLogoAsset(lang) {
    return lang === "en"
      ? DATA.lab.assets.heroLogoEn || logoAsset(lang)
      : DATA.lab.assets.heroLogoZh || logoAsset(lang);
  }

  function renderHeroTitle(lang) {
    if (lang === "zh") {
      return `
        <h1 class="hero-title hero-title-zh">
          <span class="hero-title-school">上海交通大学</span>
          <span class="hero-title-main">X-LANCE 实验室</span>
          <span class="hero-title-full">跨媒体语言智能实验室</span>
        </h1>
      `;
    }
    return `
      <h1 class="hero-title hero-title-en">
        <span class="hero-title-school">Shanghai Jiao Tong University</span>
        <span class="hero-title-main">X-LANCE Lab</span>
        <span class="hero-title-full">Cross-Media Language Intelligence Laboratory</span>
      </h1>
    `;
  }

  function renderHero(lang) {
    const heroImage = pageHeroImage("home");
    const stats = homeStats(lang).map((item) => `
      <a class="stat" href="${escapeHtml(item.href)}" data-stat-key="${escapeHtml(item.key)}">
        <strong><span class="stat-number" data-stat-target="${escapeHtml(item.value)}">0</span>${item.plus ? '<b aria-hidden="true">+</b>' : ""}</strong>
        <span>${escapeHtml(tr(item.label, lang))}</span>
      </a>
    `).join("");
    return `
      <section class="hero-section">
        <div class="hero-copy">
          ${imageTag(heroLogoAsset(lang), "X-LANCE", 'class="hero-logo"')}
          ${renderHeroTitle(lang)}
          <p class="lead">${escapeHtml(tr(DATA.lab.intro, lang))}</p>
          <div class="hero-actions">
            <a class="button primary" href="/join">${escapeHtml(UI[lang].homeCtaJoin)}</a>
          </div>
        </div>
        <div class="hero-media">
          ${imageTag(heroImage, "X-LANCE", 'decoding="async"')}
        </div>
      </section>
      <section class="stats-strip">${stats}</section>
    `;
  }

  function renderPageHero(lang, title, description = "") {
    const heroImage = pageHeroImage();
    return `
      <section class="page-hero">
        ${imageTag(heroImage, "", 'class="page-hero-image" decoding="async"')}
        <div class="page-hero-content wrap">
          <h1>${escapeHtml(title)}</h1>
          ${description ? `<p>${escapeHtml(description)}</p>` : ""}
        </div>
      </section>
    `;
  }

  function numericStatFallback(key, fallback = 0) {
    const item = (DATA.stats || []).find((stat) => stat.key === key);
    const value = Number.parseInt(item?.value, 10);
    return Number.isFinite(value) ? value : fallback;
  }

  function statLabel(key, fallback) {
    const item = (DATA.stats || []).find((stat) => stat.key === key);
    return item?.label || fallback;
  }

  function studentCount() {
    return (memberData().students || []).reduce((total, section) => total + (section.members || []).length, 0);
  }

  function paperPublicationCount() {
    const total = (fullPaperData().years || []).reduce((sum, year) => sum + Number(year.count || 0), 0);
    return total || numericStatFallback("papers", DATA.papers.length);
  }

  function homeStats(lang) {
    const people = memberData();
    return [
      { key: "research", value: DATA.research.length, label: statLabel("research", { zh: "个研究方向", en: "Research Areas" }), href: "#research" },
      { key: "faculty", value: (people.faculty || []).length || numericStatFallback("faculty", 7), label: statLabel("faculty", { zh: "位老师", en: "Faculty" }), href: "/members#members-faculty" },
      { key: "students", value: studentCount() || numericStatFallback("students", 130), label: statLabel("students", { zh: "位学生", en: "Students" }), href: "/members#members-postdoc", plus: true },
      { key: "alumni", value: alumniLists().total || numericStatFallback("alumni", 166), label: statLabel("alumni", { zh: "位校友", en: "Alumni" }), href: "/members#members-alumni", plus: true },
      { key: "papers", value: paperPublicationCount(), label: statLabel("papers", { zh: "篇论文", en: "Papers" }), href: "/papers", plus: true },
      { key: "xmart", value: DATA.videos.length, label: statLabel("xmart", { zh: "场Xmart论坛", en: "Xmart Lectures" }), href: "/videos" }
    ].map((item) => ({ ...item, label: tr(item.label, lang) }));
  }

  function formatStatNumber(value) {
    const locale = currentLang === "zh" ? "zh-CN" : "en-US";
    return new Intl.NumberFormat(locale).format(value);
  }

  function animateStats() {
    const numbers = Array.from(document.querySelectorAll(".stat-number[data-stat-target]"));
    if (!numbers.length) return;
    const reducedMotion = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    numbers.forEach((node, index) => {
      const target = Number.parseInt(node.dataset.statTarget, 10) || 0;
      if (reducedMotion || target <= 0) {
        node.textContent = formatStatNumber(target);
        return;
      }
      const steps = 16 + index * 2;
      let step = 0;
      let previous = 0;
      node.classList.add("is-jumping");
      const timer = window.setInterval(() => {
        step += 1;
        const progress = step / steps;
        if (step >= steps) {
          window.clearInterval(timer);
          node.textContent = formatStatNumber(target);
          node.classList.remove("is-jumping");
          return;
        }
        const eased = 1 - Math.pow(1 - progress, 2.8);
        const next = Math.max(previous, Math.min(target, Math.round(target * eased)));
        previous = next;
        node.textContent = formatStatNumber(next);
      }, 48 + index * 8);
    });
  }

  function cardTags(tags) {
    return (tags || []).map((tag) => `<span class="tag">${escapeHtml(tag)}</span>`).join("");
  }

  function renderResearchCards(lang, limit, showTags = true) {
    return DATA.research.slice(0, limit || DATA.research.length).map((item, index) => `
      <article class="info-card research-card">
        <span class="research-index" aria-hidden="true">${String(index + 1).padStart(2, "0")}</span>
        <div class="research-card-content">
          ${showTags ? `<div class="card-tags">${cardTags(item.tags)}</div>` : ""}
          <h3>${escapeHtml(tr(item.title, lang))}</h3>
          <p>${escapeHtml(tr(item.summary, lang))}</p>
        </div>
      </article>
    `).join("");
  }

  function renderHighlightCards(lang, limit) {
    return DATA.highlights.slice(0, limit || DATA.highlights.length).map((item) => `
      <article class="highlight-card">
        <div class="meta-line"><span>${escapeHtml(tr(item.category, lang))}</span><time>${escapeHtml(item.date)}</time></div>
        <h3>${escapeHtml(tr(item.title, lang))}</h3>
        <p>${escapeHtml(tr(item.summary, lang))}</p>
        <a href="${escapeHtml(item.link)}" target="_blank" rel="noopener">${escapeHtml(UI[lang].visit)}</a>
      </article>
    `).join("");
  }

  function renderFocus(lang) {
    return DATA.focusTopics.map((item) => `
      <article class="focus-item">
        <h3>${escapeHtml(tr(item.title, lang))}</h3>
        <p>${escapeHtml(tr(item.body, lang))}</p>
      </article>
    `).join("");
  }

  function repoByName(name) {
    return DATA.repos.find((repo) => repo.name === name);
  }

  function renderWorkAreas(lang, limit) {
    return DATA.workAreas.slice(0, limit || DATA.workAreas.length).map((area) => {
      const chips = area.repos.map((name) => {
        const repo = repoByName(name);
        const label = repo ? repo.name : name;
        return repo
          ? `<a href="${escapeHtml(repo.url)}" target="_blank" rel="noopener">${escapeHtml(label)}</a>`
          : `<span>${escapeHtml(label)}</span>`;
      }).join("");
      return `
        <article class="work-card">
          <div class="work-card-top">
            <span>${escapeHtml(area.repos.length)}</span>
            <h3>${escapeHtml(tr(area.title, lang))}</h3>
          </div>
          <p>${escapeHtml(tr(area.body, lang))}</p>
          <div class="repo-chips">${chips}</div>
        </article>
      `;
    }).join("");
  }

  function renderNewsList(lang, limit) {
    return DATA.news.slice(0, limit || DATA.news.length).map((item) => {
      const summary = tr(item.summary, lang);
      const links = Array.isArray(item.links) && item.links.length
        ? item.links
        : (item.link ? [{ url: item.link, label: { zh: UI.zh.visit, en: UI.en.visit } }] : []);
      return `
        <article class="news-item">
          <div class="date-block">
            <time>${escapeHtml(item.date)}</time>
            <span>${escapeHtml(typeLabel(item.type, lang))}</span>
          </div>
          <div>
            <h3>${escapeHtml(tr(item.title, lang))}</h3>
            ${summary ? `<p>${escapeHtml(summary)}</p>` : ""}
            ${links.length ? `<div class="news-links">${links.map((link) => `<a href="${escapeHtml(link.url)}" target="_blank" rel="noopener">${escapeHtml(tr(link.label, lang) || UI[lang].visit)}</a>`).join("")}</div>` : ""}
          </div>
        </article>
      `;
    }).join("");
  }

  function renderMultilineText(value) {
    return escapeHtml(value).replace(/\n/g, "<br />");
  }

  function renderAboutHome(lang) {
    const about = DATA.about || {};
    const paragraphs = tr(about.paragraphs, lang);
    const list = Array.isArray(paragraphs) ? paragraphs : [];
    return `
      <section class="wrap section home-about" id="about">
        <div class="section-head reveal-block">
          <h2>${escapeHtml(tr(about.title, lang) || UI[lang].about)}</h2>
        </div>
        <div class="about-layout reveal-block">
          <div class="about-panel">
            ${list.map((paragraph) => `<p>${escapeHtml(paragraph)}</p>`).join("")}
          </div>
          <aside class="home-about-qr">
            ${imageTag(DATA.lab.assets.wechat, UI[lang].followWechat, 'loading="lazy" decoding="async"')}
            <span>${escapeHtml(UI[lang].followWechat)}</span>
          </aside>
        </div>
      </section>
    `;
  }

  function renderCredoHome(lang) {
    const credo = DATA.credo || {};
    return `
      <section class="wrap section home-credo" id="credo">
        <div class="section-head reveal-block">
          <h2>${escapeHtml(tr(credo.title, lang) || UI[lang].credo)}</h2>
        </div>
        <div class="about-panel credo-panel reveal-block">
          <blockquote>${renderMultilineText(tr(credo.body, lang))}</blockquote>
        </div>
      </section>
    `;
  }

  function renderPaperCards(lang, list) {
    return list.map((paper) => `
      <article class="paper-card" data-topic="${escapeHtml(paper.topic)}">
        <div class="paper-topline">
          <span>${escapeHtml(paper.year)}</span>
          <span>${escapeHtml(paper.venue)}</span>
          <span>${escapeHtml(typeLabel(paper.topic, lang))}</span>
        </div>
        <h3>${escapeHtml(paper.title)}</h3>
        <p class="authors">${escapeHtml(paper.authors)}</p>
        <p>${escapeHtml(tr(paper.why, lang))}</p>
        <a href="${escapeHtml(paper.link)}" target="_blank" rel="noopener">${escapeHtml(UI[lang].paperLink)}</a>
      </article>
    `).join("");
  }

  function fullPaperData() {
    return window.XLANCE_FULL_PAPERS || { years: [], typeLabels: {} };
  }

  function paperTypeLabel(type, lang) {
    const data = fullPaperData();
    return tr((data.typeLabels || {})[type] || type, lang);
  }

  function paperYearId(year) {
    return `papers-year-${String(year).replace(/[^0-9A-Za-z_-]/g, "")}`;
  }

  function withoutLegacyFileLabels(value) {
    return String(value || "")
      .replace(/(?:[\w\u4e00-\u9fff-]+)?\.(?:pdf|pptx?)(?=(?:\s|[,;]|$))/gi, "")
      .replace(/\s+([,.;:])/g, "$1")
      .replace(/([.,])\1+/g, "$1")
      .replace(/\s{2,}/g, " ")
      .trim();
  }

  function normalizedPaperEntry(entry, year) {
    let citation = withoutLegacyFileLabels(entry.citation);
    const files = (entry.files || []).map((file) => {
      const value = String(file || "");
      const glued = value.match(/^(\d{4})\.(.+\.(?:pdf|pptx?))$/i);
      if (!glued) return value;
      if (!citation.includes(glued[1])) {
        citation = `${citation.replace(/[,\s]*$/, "")}, ${glued[1]}.`;
      }
      return glued[2];
    });
    if (year && !citation.includes(String(year))) {
      citation = `${citation.replace(/[,\s]*$/, "")}, ${year}.`;
    }
    return {
      title: entry.title || "",
      authors: entry.authors || "",
      venueShort: entry.venueShort || entry.venue_short || "",
      citation,
      files,
      award: entry.award || ""
    };
  }

  function renderPaperLinks(links, lang) {
    const clean = (links || []).filter((link) => String(link?.url || "").startsWith("http"));
    if (!clean.length) return "";
    return `
      <div class="paper-links" aria-label="${escapeHtml(UI[lang].paperLink)}">
        ${clean.map((link) => `<a href="${escapeHtml(link.url)}" target="_blank" rel="noopener">${escapeHtml(tr(link.label, lang))}</a>`).join("")}
      </div>
    `;
  }

  function renderFullPaperItem(entry, lang, year) {
    const normalized = normalizedPaperEntry(entry, year);
    const venue = normalized.venueShort || paperTypeLabel(entry.type, lang);
    return `
      <article class="full-paper-item">
        <div class="paper-topline">
          <span>${escapeHtml(year)}</span>
          <span>${escapeHtml(venue)}</span>
          <span>${escapeHtml(paperTypeLabel(entry.type, lang))}</span>
        </div>
        ${normalized.title ? `<h3>${escapeHtml(normalized.title)}</h3>` : ""}
        ${normalized.authors ? `<p class="authors">${escapeHtml(normalized.authors)}</p>` : ""}
        ${normalized.citation ? `<p class="paper-citation">${escapeHtml(normalized.citation)}</p>` : ""}
        ${normalized.award ? `<p class="paper-award">${escapeHtml(normalized.award)}</p>` : ""}
        ${renderPaperLinks(entry.links, lang)}
      </article>
    `;
  }

  function renderPaperYearToc(lang) {
    const years = fullPaperData().years || [];
    return `
      <aside class="paper-year-toc" aria-label="${escapeHtml(UI[lang].paperList)}">
        ${years.map((year) => `
          <a href="#${escapeHtml(paperYearId(year.year))}">
            <span>${escapeHtml(year.year)}</span>
            <strong>${escapeHtml(year.count)}</strong>
          </a>
        `).join("")}
      </aside>
    `;
  }

  function renderFullPaperYears(lang) {
    const data = fullPaperData();
    const typeOrder = ["journal", "conference", "book"];
    return (data.years || []).map((year) => {
      const groups = typeOrder
        .map((type) => ({
          type,
          items: (year.items || []).filter((item) => item.type === type)
        }))
        .filter((group) => group.items.length);
      return `
        <section class="paper-year-section" id="${escapeHtml(paperYearId(year.year))}">
          <div class="member-section-head paper-year-head">
            <h2>${escapeHtml(year.year)}</h2>
            <span>${escapeHtml(year.count)}</span>
          </div>
          ${groups.map((group) => `
            <div class="paper-type-group">
              <h3>${escapeHtml(paperTypeLabel(group.type, lang))}</h3>
              <div class="full-paper-list">${group.items.map((item) => renderFullPaperItem(item, lang, year.year)).join("")}</div>
            </div>
          `).join("")}
        </section>
      `;
    }).join("");
  }

  function memberData() {
    return window.XLANCE_MEMBERS || { faculty: [], students: [], alumni: [] };
  }

  function joinedYear(person) {
    const year = Number.parseInt(person.joined, 10);
    return Number.isFinite(year) ? year : Number.POSITIVE_INFINITY;
  }

  function nameCompare(a, b) {
    return String(tr(a.name, "zh") || tr(a.name, "en")).localeCompare(
      String(tr(b.name, "zh") || tr(b.name, "en")),
      "zh-Hans-CN-u-co-pinyin",
      { numeric: true, sensitivity: "base" }
    );
  }

  function currentFaculty() {
    const faculty = [...(memberData().faculty || [])];
    const pinned = faculty.find((person) => tr(person.name, "zh") === "申丽萍");
    const last = faculty.find((person) => tr(person.name, "zh") === "茅懿琼");
    const sorted = faculty
      .filter((person) => person !== pinned && person !== last)
      .sort((a, b) => joinedYear(a) - joinedYear(b) || nameCompare(a, b));
    if (last) sorted.push(last);
    if (!pinned) return sorted;
    const first = sorted.shift();
    return first ? [first, pinned, ...sorted] : [pinned];
  }

  function personKey(person) {
    return person.id
      ? `id:${person.id}`
      : `name:${tr(person.name, "zh")}|${tr(person.name, "en")}`;
  }

  function alumniIdNumber(person) {
    const match = String(person.id || "").match(/^\s*(\d+)/);
    return match ? Number.parseInt(match[1], 10) : Number.POSITIVE_INFINITY;
  }

  function compareAlumni(a, b) {
    const aNumber = alumniIdNumber(a);
    const bNumber = alumniIdNumber(b);
    const aHasId = Number.isFinite(aNumber);
    const bHasId = Number.isFinite(bNumber);
    if (aHasId && bHasId && aNumber !== bNumber) return aNumber - bNumber;
    if (aHasId !== bHasId) return aHasId ? -1 : 1;
    if (a.id && b.id && a.id !== b.id) {
      return String(a.id).localeCompare(String(b.id), "en", { numeric: true, sensitivity: "base" });
    }
    return nameCompare(a, b);
  }

  function alumniLists() {
    const people = memberData();
    const featured = people.featuredAlumni || [];
    const featuredKeys = new Set(featured.map(personKey));
    const ordinary = (people.alumni || [])
      .filter((person) => !featuredKeys.has(personKey(person)))
      .sort(compareAlumni);
    return {
      featured,
      ordinary,
      total: featured.length + ordinary.length
    };
  }

  function renderHomepageMemberPreview(lang) {
    return currentFaculty().slice(0, 4).map((person) => `
      <article class="person-card compact-person">
        ${imageTag(person.image, tr(person.name, lang), 'loading="lazy" decoding="async"', "assets/img/members/octocat-placeholder.png")}
        <div>
          <h3>${escapeHtml(tr(person.name, lang))}</h3>
          <p class="role">${escapeHtml(tr(person.role, lang))}</p>
        </div>
      </article>
    `).join("");
  }

  function renderMemberCard(person, lang, options = {}) {
    const link = person.link || person.website || "";
    const bio = options.bio ? tr(person.bio, lang) : "";
    const facultyDetails = options.bio ? renderFacultyDetails(person, lang) : "";
    const facultyResearch = options.bio ? renderFacultyResearch(person, lang) : "";
    const title = renderPersonTitle(person, lang);
    const name = escapeHtml(tr(person.name, lang));
    const linkedName = link ? `<a href="${escapeHtml(link)}" target="_blank" rel="noopener">${name}</a>` : name;
    const facultyId = options.bio && person.id ? ` <span class="faculty-name-id">${escapeHtml(person.id)}</span>` : "";
    return `
      <article class="member-card ${options.bio ? "faculty-member" : ""}">
        ${imageTag(person.image, tr(person.name, lang), 'loading="lazy" decoding="async"', "assets/img/members/octocat-placeholder.png")}
        <div>
          <h3>${linkedName}${facultyId}</h3>
          ${title ? `<p class="role">${title}</p>` : ""}
          ${!options.bio && person.id ? `<p class="member-id">${escapeHtml(person.id)}</p>` : ""}
          ${!options.bio && person.email ? `<p class="member-id">${escapeHtml(person.email)}</p>` : ""}
          ${facultyDetails}
          ${facultyResearch}
          ${bio ? `<p class="faculty-bio">${escapeHtml(bio)}</p>` : ""}
        </div>
      </article>
    `;
  }

  function renderPersonTitle(person, lang) {
    const title = tr(person.title, lang);
    if (!title) return "";
    const link = person.titleLink;
    const label = link ? tr(link.label, lang) : "";
    if (!link || !link.url || !label) return escapeHtml(title);
    const index = title.indexOf(label);
    if (index < 0) return escapeHtml(title);
    return [
      escapeHtml(title.slice(0, index)),
      `<a href="${escapeHtml(link.url)}" target="_blank" rel="noopener">${escapeHtml(label)}</a>`,
      escapeHtml(title.slice(index + label.length))
    ].join("");
  }

  function renderFacultyDetails(person, lang) {
    const office = tr(person.office, lang);
    const joined = tr(person.joinedRange, lang) || (person.joined ? (lang === "zh" ? `${person.joined}年-至今` : `${person.joined}-present`) : "");
    const rows = [
      joined ? [UI[lang].joined, escapeHtml(joined)] : null,
      office ? [UI[lang].office, escapeHtml(office)] : null,
      person.email ? [UI[lang].email, `<a href="mailto:${escapeHtml(person.email)}">${escapeHtml(person.email)}</a>`] : null
    ].filter(Boolean);
    if (!rows.length) return "";
    return `
      <dl class="faculty-details">
        ${rows.map(([label, value]) => `<div><dt>${escapeHtml(label)}</dt><dd>${value}</dd></div>`).join("")}
      </dl>
    `;
  }

  function renderFacultyResearch(person, lang) {
    const items = trList(person.research, lang);
    if (!items.length) return "";
    const label = tr(person.researchLabel, lang) || UI[lang].researchLabel;
    return `
      <div class="faculty-research">
        <strong>${escapeHtml(label)}</strong>
        <div class="card-tags faculty-tags">${items.map((item) => `<span class="tag">${escapeHtml(item)}</span>`).join("")}</div>
      </div>
    `;
  }

  function renderMemberGrid(members, lang) {
    return members.map((person) => renderMemberCard(person, lang)).join("");
  }

  function renderStudentSections(lang) {
    const people = memberData();
    return (people.students || []).map((section) => `
      <section class="member-section" id="members-${escapeHtml(section.key)}">
        <div class="member-section-head">
          <h2>${escapeHtml(tr(section.title, lang))}</h2>
          <span>${section.members.length}</span>
        </div>
        <div class="member-grid">${renderMemberGrid(section.members, lang)}</div>
      </section>
    `).join("");
  }

  function renderMemberToc(lang) {
    const people = memberData();
    const alumni = alumniLists();
    const items = [
      { href: "#members-faculty", label: UI[lang].faculty, count: currentFaculty().length },
      ...(people.students || []).map((section) => ({
        href: `#members-${section.key}`,
        label: tr(section.title, lang),
        count: section.members.length
      })),
      { href: "#members-alumni", label: UI[lang].alumni, count: alumni.total }
    ];
    return `
      <aside class="member-toc" aria-label="${escapeHtml(UI[lang].people)}">
        ${items.map((item) => `
          <a href="${escapeHtml(item.href)}">
            <span>${escapeHtml(item.label)}</span>
            <strong>${escapeHtml(item.count)}</strong>
          </a>
        `).join("")}
      </aside>
    `;
  }

  function renderRepoCard(repo, lang) {
    return `
      <article class="repo-card">
        <div>
          <div class="repo-card-top">
            <span class="tag">${escapeHtml(typeLabel(repo.topic, lang))}</span>
            ${repo.meta ? `<span class="repo-meta">${escapeHtml(repo.meta)}</span>` : ""}
          </div>
          <h3>${escapeHtml(repo.name)}</h3>
          <p>${escapeHtml(tr(repo.desc, lang))}</p>
        </div>
        <a href="${escapeHtml(repo.url)}" target="_blank" rel="noopener">${escapeHtml(UI[lang].projectLink)}</a>
      </article>
    `;
  }

  function renderRepos(lang, list = DATA.repos) {
    return list.map((repo) => renderRepoCard(repo, lang)).join("");
  }

  function renderRepoGroups(lang) {
    return DATA.workAreas.map((area) => {
      const repos = DATA.repos.filter((repo) => repo.group === area.key);
      return `
        <section class="repo-section">
          <div class="member-section-head repo-section-head">
            <div>
              <h2>${escapeHtml(tr(area.title, lang))}</h2>
            </div>
            <span>${repos.length}</span>
          </div>
          <div class="repo-grid">${renderRepos(lang, repos)}</div>
        </section>
      `;
    }).join("");
  }

  function isFrontierLecture(video) {
    const series = String(tr(video.series, "zh") || tr(video.series, "en")).toLowerCase();
    return series.includes("前沿") || series.includes("frontier");
  }

  function xmartSeriesLabel(video, lang) {
    return isFrontierLecture(video) ? UI[lang].frontierLecture : UI[lang].youthForum;
  }

  function renderVideos(lang, list = DATA.videos) {
    return list.map((video) => `
      <a class="video-card" href="${escapeHtml(video.url)}" target="_blank" rel="noopener" aria-label="${escapeHtml(`${UI[lang].watchVideo}: ${tr(video.title, lang)}`)}">
        <div class="video-thumb">
          ${imageTag(video.cover, "", 'loading="lazy" decoding="async"')}
          <span>${escapeHtml(UI[lang].watchVideo)}</span>
        </div>
        <div>
          <div class="paper-topline"><span>${escapeHtml(video.date || video.year)}</span><span>${escapeHtml(xmartSeriesLabel(video, lang))}</span></div>
          <h3>${escapeHtml(tr(video.title, lang))}</h3>
        </div>
      </a>
    `).join("");
  }

  function renderHome(lang) {
    return `
      ${renderHero(lang)}
      ${renderAboutHome(lang)}
      ${renderCredoHome(lang)}
      <section class="wrap section research-overview" id="research">
        <div class="section-head reveal-block">
          <h2>${escapeHtml(UI[lang].directions)}</h2>
        </div>
        <div class="research-grid reveal-block">${renderResearchCards(lang, undefined, false)}</div>
      </section>
      <section class="section-band muted" id="news">
        <div class="wrap reveal-block">
          <div class="section-head">
            <h2>${escapeHtml(UI[lang].latest)}</h2>
          </div>
          <div class="news-list compact">${renderNewsList(lang, HOME_NEWS_LIMIT)}</div>
        </div>
      </section>
    `;
  }

  function renderResearchPage(lang) {
    return `
      ${renderPageHero(lang, UI[lang].directions, tr(DATA.lab.intro, lang))}
      <section class="wrap section">
        <div class="card-grid wide">${renderResearchCards(lang)}</div>
      </section>
    `;
  }

  function renderNewsPage(lang) {
    return `
      ${renderPageHero(lang, UI[lang].latest)}
      <section class="wrap section">
        <div class="news-list">${renderNewsList(lang)}</div>
      </section>
    `;
  }

  function renderPapersPage(lang) {
    const topics = [UI[lang].all, ...Array.from(new Set(DATA.papers.map((paper) => typeLabel(paper.topic, lang))))];
    const buttons = topics.map((label, idx) => `<button class="${idx === 0 ? "active" : ""}" data-paper-filter="${idx === 0 ? "all" : escapeHtml(label)}">${escapeHtml(label)}</button>`).join("");
    const full = fullPaperData();
    const fullCount = (full.years || []).reduce((total, year) => total + Number(year.count || 0), 0);
    return `
      ${renderPageHero(lang, UI[lang].papers)}
      <section class="wrap section">
        <div class="section-head">
          <h2>${escapeHtml(UI[lang].representativeWorks)}</h2>
        </div>
        <div class="filter-bar">${buttons}</div>
        <div class="paper-list">${renderPaperCards(lang, DATA.papers)}</div>
      </section>
      <section class="section-band paper-archive-band">
        <div class="wrap">
          <div class="member-section-head paper-archive-head">
            <h2>${escapeHtml(UI[lang].paperList)}</h2>
            <span>${escapeHtml(fullCount)}</span>
          </div>
          <div class="papers-layout">
            ${renderPaperYearToc(lang)}
            <div class="paper-year-sections">${renderFullPaperYears(lang)}</div>
          </div>
        </div>
      </section>
    `;
  }

  function renderFacultyPage(lang) {
    return `
      ${renderPageHero(lang, UI[lang].faculty)}
      <section class="wrap section">
        <div class="member-grid faculty-grid">${currentFaculty().map((person) => renderMemberCard(person, lang, { bio: true })).join("")}</div>
      </section>
    `;
  }

  function renderStudentsPage(lang) {
    return `
      ${renderPageHero(lang, UI[lang].students)}
      <section class="wrap section member-sections">
        ${renderStudentSections(lang)}
      </section>
    `;
  }

  function renderAlumniSection(lang) {
    const alumni = alumniLists();
    return `
      <section class="member-section" id="members-alumni">
        <div class="member-section-head">
          <h2>${escapeHtml(UI[lang].alumni)}</h2>
          <span>${alumni.total}</span>
        </div>
        ${alumni.featured.length ? `<div class="member-grid faculty-grid alumni-featured">${alumni.featured.map((person) => renderMemberCard(person, lang, { bio: true })).join("")}</div>` : ""}
        <div class="member-grid alumni-grid">${renderMemberGrid(alumni.ordinary, lang)}</div>
      </section>
    `;
  }

  function renderMembersPage(lang) {
    const faculty = currentFaculty();
    return `
      ${renderPageHero(lang, UI[lang].people)}
      <section class="wrap section members-layout">
        ${renderMemberToc(lang)}
        <div class="member-sections">
          <section class="member-section" id="members-faculty">
            <div class="member-section-head">
              <h2>${escapeHtml(UI[lang].faculty)}</h2>
              <span>${faculty.length}</span>
            </div>
            <div class="member-grid faculty-grid">${faculty.map((person) => renderMemberCard(person, lang, { bio: true })).join("")}</div>
          </section>
          ${renderStudentSections(lang)}
          ${renderAlumniSection(lang)}
        </div>
      </section>
    `;
  }

  function renderProjectsPage(lang) {
    return `
      ${renderPageHero(lang, UI[lang].projects)}
      <section class="wrap section">
        <div class="work-grid">${renderWorkAreas(lang)}</div>
      </section>
      <section class="section-band resource-band">
        <div class="wrap repo-sections">${renderRepoGroups(lang)}</div>
      </section>
    `;
  }

  function renderXmartPage(lang) {
    const groups = [
      { key: "frontier", title: UI[lang].frontierLecture, videos: DATA.videos.filter(isFrontierLecture) },
      { key: "youth", title: UI[lang].youthForum, videos: DATA.videos.filter((video) => !isFrontierLecture(video)) }
    ];
    return `
      ${renderPageHero(lang, UI[lang].xmart)}
      <section class="wrap section xmart-sections">
        ${groups.map((group) => `
          <section class="xmart-section" id="xmart-${escapeHtml(group.key)}">
            <div class="member-section-head repo-section-head">
              <h2>${escapeHtml(group.title)}</h2>
              <span>${group.videos.length}</span>
            </div>
            <div class="video-grid">${renderVideos(lang, group.videos)}</div>
          </section>
        `).join("")}
      </section>
    `;
  }

  function renderJoinPage(lang) {
    const emailSubject = lang === "zh"
      ? "【实验室招新】+ 姓名 + 学校 + 专业 + 年级"
      : "[X-LANCE Recruitment] + Name + School + Major + Year";
    return `
      ${renderPageHero(lang, UI[lang].joinUs)}
      <section class="wrap section join-layout">
        <article class="join-panel">
          <h2>${escapeHtml(UI[lang].recruiting)}</h2>
          <p>${escapeHtml(UI[lang].recruitingText)}</p>
        </article>
        <aside class="join-panel join-contact">
          <h2>${escapeHtml(UI[lang].contact)}</h2>
          <p><strong>${escapeHtml(UI[lang].contactPerson)}</strong></p>
          <p class="join-email">${escapeHtml(DATA.lab.email)}</p>
          <p><strong>${escapeHtml(UI[lang].emailSubject)}</strong><br />${escapeHtml(emailSubject)}</p>
          <p><strong>${escapeHtml(UI[lang].resume)}</strong><br />${escapeHtml(UI[lang].resumeText)}</p>
        </aside>
      </section>
    `;
  }

  function renderSourcePage(lang) {
    return `
      ${renderPageHero(lang, UI[lang].source)}
    `;
  }

  function bindPaperFilter() {
    const bar = document.querySelector(".filter-bar");
    if (!bar) return;
    bar.addEventListener("click", (event) => {
      const button = event.target.closest("button");
      if (!button) return;
      const filter = button.dataset.paperFilter;
      document.querySelectorAll(".filter-bar button").forEach((item) => item.classList.toggle("active", item === button));
      document.querySelectorAll(".paper-card").forEach((card) => {
        const label = typeLabel(card.dataset.topic, currentLang);
        card.hidden = filter !== "all" && label !== filter;
      });
    });
  }

  let currentLang = detectLang();
  let currentTheme = detectTheme();

  function render() {
    setLang(currentLang);
    setTheme(currentTheme);
    const active = pageId();
    const activeId = navActiveId(active);
    renderHeader(active, currentLang, currentTheme);
    syncHeaderHeight();
    requestAnimationFrame(syncActiveNav);
    renderFooter(currentLang);
    const renderers = {
      home: renderHome,
      research: renderResearchPage,
      news: renderNewsPage,
      papers: renderPapersPage,
      faculty: renderMembersPage,
      students: renderMembersPage,
      members: renderMembersPage,
      projects: renderProjectsPage,
      resources: renderProjectsPage,
      xmart: renderXmartPage,
      media: renderXmartPage,
      join: renderJoinPage,
      source: renderSourcePage
    };
    const app = document.getElementById("app");
    app.innerHTML = (renderers[active] || renderHome)(currentLang);
    document.title = `${tr((DATA.nav.find((item) => item.id === activeId) || {}).label, currentLang) || DATA.lab.shortName} - ${DATA.lab.shortName}`;
    document.querySelector("[data-lang-toggle]").addEventListener("click", () => {
      currentLang = currentLang === "zh" ? "en" : "zh";
      render();
    });
    document.querySelector("[data-theme-toggle]").addEventListener("click", () => {
      currentTheme = nextTheme(currentTheme);
      render();
    });
    bindPaperFilter();
    scheduleProgressiveImageUpgrades(document);
    requestAnimationFrame(animateStats);
    requestAnimationFrame(() => {
      syncHeaderHeight();
      scrollToHashTarget();
    });
    setTimeout(scrollToHashTarget, 80);
  }

  window.addEventListener("resize", syncHeaderHeight);
  window.addEventListener("hashchange", () => {
    requestAnimationFrame(scrollToHashTarget);
    setTimeout(scrollToHashTarget, 80);
  });
  document.addEventListener("DOMContentLoaded", () => {
    loadGeneratedData().finally(render);
  });
})();
