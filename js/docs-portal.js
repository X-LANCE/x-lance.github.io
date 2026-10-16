(() => {
  let user = null, app = null, loading = false, dirty = false, notify = () => {}, lastPath = null, revision = 0, authRequest = 0;
  const route = () => new URL(location.href).searchParams.get('docs');
  const safe = path => typeof path === 'string' && /^\/(?!\/)[^\r\n]*$/.test(path) ? path : '/';
  const href = path => '/?docs=' + encodeURIComponent(safe(path));
  const leave = () => !dirty || confirm(document.documentElement.lang.startsWith('zh') ? '还有未保存的修改，确认离开？' : 'You have unsaved changes. Leave this page?');
  function auth(next, broadcast = false) {
    if (JSON.stringify(user) === JSON.stringify(next)) return;
    user = next;
    if (broadcast) localStorage.setItem('xlance_auth_changed', String(Date.now()));
    notify();
  }
  async function refresh() {
    const request = ++authRequest;
    try {
      const r = await fetch('/docs/api/v1/auth/me/', { credentials: 'same-origin', cache: 'no-store' });
      if (!r.ok) return;
      const data = await r.json(); if (request !== authRequest) return; auth(data.user); app?.update({ authVersion: ++revision });
    } catch { /* Public static pages remain usable when the session service is unavailable. */ }
  }
  function go(path) {
    if (!leave()) return;
    dirty = false; history.pushState(null, '', path === null ? '/' : href(path));
    if (app && path !== null) { lastPath = safe(path); app.navigate(lastPath); }
    notify(); window.scrollTo(0, 0);
  }
  async function login(username, password) {
    const fallback = document.documentElement.lang.startsWith('zh') ? '登录失败，请稍后重试。' : 'Sign-in failed. Please try again.';
    const identity = await fetch('/docs/api/v1/auth/me/', { credentials: 'same-origin', cache: 'no-store' });
    if (!identity.ok) throw new Error(fallback);
    const session = await identity.json();
    const response = await fetch('/docs/api/v1/auth/login/', {
      method: 'POST', credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', 'X-CSRFToken': session.csrf_token },
      body: JSON.stringify({ username, password })
    });
    const result = await response.json().catch(() => null);
    if (!response.ok || !result?.user) {
      const message = result?.error?.message || result?.error || result?.detail;
      throw new Error(typeof message === 'string' ? message : fallback);
    }
    ++authRequest;
    auth(result.user, true);
    app?.update({ authVersion: ++revision });
    go('/');
  }
  async function logout() {
    if (!leave()) return;
    try {
      const r = await fetch('/docs/api/v1/auth/me/', { credentials: 'same-origin', cache: 'no-store' });
      if (!r.ok) throw new Error();
      const data = await r.json();
      const result = await fetch('/docs/api/v1/auth/logout/', { method: 'POST', credentials: 'same-origin', headers: { 'X-CSRFToken': data.csrf_token } });
      if (!result.ok) throw new Error();
      ++authRequest; dirty = false; auth(null, true); go(null);
    } catch { alert(document.documentElement.lang.startsWith('zh') ? '退出失败，请重试。' : 'Sign-out failed. Please retry.'); }
  }
  function sync(lang, theme) {
    if (route() === null) { app?.unmount(); app = null; lastPath = null; dirty = false; return; }
    const options = { lang, theme, authVersion: revision, path: safe(route()),
      onAuth: next => { ++authRequest; auth(next, true); }, onDirty: value => { dirty = value; },
      onNavigate: next => {
        lastPath = next;
        if (route() !== next) { history[route() === '/login' ? 'replaceState' : 'pushState'](null, '', href(next)); window.scrollTo(0, 0); notify(); }
      }
    };
    if (app) { app.update(options); return; }
    if (!document.getElementById('docs-portal')) document.getElementById('app').innerHTML = '<div id="docs-portal"><p class="portal-loading" role="status">' + (lang === 'zh' ? '正在加载…' : 'Loading…') + '</p></div>';
    if (loading) return;
    loading = true;
    const host = document.getElementById('docs-portal');
    import('/docs/assets/portal.js?v=20261016-integrated').then(module => {
      if (!host.isConnected || route() === null) return;
      host.replaceChildren(); options.path = safe(route());
      options.lang = document.documentElement.lang.startsWith('zh') ? 'zh' : 'en'; options.theme = document.documentElement.dataset.theme;
      lastPath = options.path; app = module.mount(host, options);
    }).catch(() => { if (host.isConnected) host.innerHTML = '<p class="portal-loading">' + (lang === 'zh' ? '加载失败，请刷新页面重试。' : 'Loading failed. Please refresh to retry.') + '</p>'; }).finally(() => { loading = false; if (!host.isConnected && route() !== null && !app) sync(document.documentElement.lang.startsWith('zh') ? 'zh' : 'en', document.documentElement.dataset.theme); });
  }
  function init(callback) {
    notify = callback;
    document.addEventListener('click', event => {
      const link = event.composedPath().find(node => node instanceof HTMLAnchorElement);
      if (link && dirty && !link.getAttribute('href')?.startsWith('#')) {
        if (!leave()) { event.preventDefault(); event.stopImmediatePropagation(); return; }
        dirty = false;
      }
      if (event.composedPath().some(node => node instanceof HTMLElement && node.hasAttribute('data-site-logout'))) { event.preventDefault(); logout(); return; }
      if (link?.hasAttribute('data-docs-route') && !event.ctrlKey && !event.metaKey && !event.shiftKey && event.button === 0) { event.preventDefault(); go(link.dataset.docsRoute); }
    }, true);
    window.addEventListener('popstate', () => {
      if (!leave()) { history.pushState(null, '', lastPath === null ? '/' : href(lastPath)); return; }
      dirty = false;
      if (app && route() !== null) { lastPath = safe(route()); app.navigate(lastPath); }
      notify();
    });
    window.addEventListener('focus', refresh);
    window.addEventListener('storage', event => {
      if (event.key === 'xlance_auth_changed') refresh();
      if (event.key === 'xlance_static_lang' || event.key === 'xlance_static_theme') window.dispatchEvent(new Event('xlance-preferences'));
    });
    refresh();
  }
  window.XLANCE_PORTAL = { init, sync, route, href, login, get user() { return user; } };
})();
