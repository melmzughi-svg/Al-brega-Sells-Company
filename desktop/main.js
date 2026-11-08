'use strict';
/* Desktop shell for the sales system. It opens the live system (GitHub Pages) in its own window, so the app is
   always the latest version with no reinstall, and the service worker from the site keeps it working offline.
   Firebase sign-in works because the page still runs on its real, authorised origin. */
const { app, BrowserWindow, Menu, shell, session, dialog } = require('electron');
const path = require('node:path');
const fs = require('node:fs');

// BREGA_APP_URL is honoured only when running from source (npm start), for local testing; the installed app ignores it.
const APP_URL = (!app.isPackaged && process.env.BREGA_APP_URL) || 'https://melmzughi-svg.github.io/Al-brega-Sells-Company/';
const APP_ORIGIN = new URL(APP_URL).origin;
const APP_PATH = new URL(APP_URL).pathname;
// Hosts allowed to open inside the app (Google / Firebase sign-in popups). Everything else opens in the default browser.
const AUTH_HOSTS = ['brega-sales.firebaseapp.com', 'accounts.google.com'];
const ICON = path.join(__dirname, 'build', 'icon.png');

app.setAppUserModelId('ly.brega.sales.desktop');
if (!app.requestSingleInstanceLock()) { app.quit(); }

let win = null;
const stateFile = () => path.join(app.getPath('userData'), 'window-state.json');
function readState() { try { return JSON.parse(fs.readFileSync(stateFile(), 'utf8')); } catch (e) { return { width: 1360, height: 860, maximized: true }; } }
function saveState() {
  if (!win || win.isDestroyed()) return;
  const b = win.getNormalBounds();
  try { fs.writeFileSync(stateFile(), JSON.stringify({ ...b, maximized: win.isMaximized() })); } catch (e) { /* ignore */ }
}

function isAppUrl(u) { try { const x = new URL(u); return x.origin === APP_ORIGIN && x.pathname.startsWith(APP_PATH); } catch (e) { return false; } }
function isAuthUrl(u) { try { return AUTH_HOSTS.includes(new URL(u).hostname); } catch (e) { return false; } }
function openOutside(u) { try { const p = new URL(u).protocol; if (['https:', 'http:', 'mailto:', 'tel:'].includes(p)) shell.openExternal(u); } catch (e) { /* ignore */ } }

function childOptions() {
  return { icon: ICON, autoHideMenuBar: true, webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true } };
}

function wireContents(wc) {
  // window.open: blank windows (print / PDF export) and sign-in popups stay in the app; other links go to the browser.
  wc.setWindowOpenHandler(({ url }) => {
    if (!url || url === 'about:blank' || isAppUrl(url) || isAuthUrl(url)) return { action: 'allow', overrideBrowserWindowOptions: childOptions() };
    openOutside(url);
    return { action: 'deny' };
  });
  wc.on('will-navigate', (e, url) => {
    if (isAppUrl(url) || isAuthUrl(url) || url.startsWith('file:')) return;
    e.preventDefault(); openOutside(url);
  });
}

function showOffline() {
  if (win && !win.isDestroyed()) win.loadFile(path.join(__dirname, 'offline.html'));
}

function createWindow() {
  const s = readState();
  win = new BrowserWindow({
    x: s.x, y: s.y, width: s.width || 1360, height: s.height || 860, minWidth: 900, minHeight: 600,
    title: 'منظومة المبيعات - البريقة', icon: ICON, backgroundColor: '#0b3168', show: false, autoHideMenuBar: true,
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: false }
  });
  if (s.maximized) win.maximize();
  win.once('ready-to-show', () => win.show());
  ['resize', 'move', 'close'].forEach(ev => win.on(ev, saveState));
  wireContents(win.webContents);
  win.webContents.on('did-fail-load', (e, code, desc, url, isMain) => {
    // -3 = aborted (normal during redirects). Only fall back when the main page itself could not load (no network and no cache yet).
    if (isMain && code !== -3 && isAppUrl(url)) showOffline();
  });
  win.loadURL(APP_URL);
}

function buildMenu() {
  const t = [
    { label: 'المنظومة', submenu: [
      { label: 'الصفحة الرئيسية', accelerator: 'Alt+Home', click: () => win && win.loadURL(APP_URL) },
      { label: 'إعادة التحميل', role: 'reload' },
      { label: 'إعادة التحميل الكامل', role: 'forceReload' },
      { type: 'separator' },
      { label: 'طباعة', accelerator: 'CmdOrCtrl+P', click: () => win && win.webContents.print() },
      { type: 'separator' },
      { label: 'خروج', role: 'quit' }
    ] },
    { label: 'تحرير', submenu: [
      { label: 'تراجع', role: 'undo' }, { label: 'إعادة', role: 'redo' }, { type: 'separator' },
      { label: 'قص', role: 'cut' }, { label: 'نسخ', role: 'copy' }, { label: 'لصق', role: 'paste' }, { label: 'تحديد الكل', role: 'selectAll' }
    ] },
    { label: 'عرض', submenu: [
      { label: 'تكبير', role: 'zoomIn' }, { label: 'تصغير', role: 'zoomOut' }, { label: 'الحجم الطبيعي', role: 'resetZoom' },
      { type: 'separator' }, { label: 'ملء الشاشة', role: 'togglefullscreen' },
      { type: 'separator' }, { label: 'أدوات المطوّر', role: 'toggleDevTools' }
    ] },
    { label: 'مساعدة', submenu: [
      { label: 'فتح المنظومة في المتصفح', click: () => shell.openExternal(APP_URL) },
      { label: 'حول البرنامج', click: () => dialog.showMessageBox(win, {
        type: 'info', title: 'حول البرنامج', message: 'منظومة الإدارة العامة للمبيعات والشؤون التجارية',
        detail: `شركة البريقة لتسويق النفط\nنسخة سطح المكتب ${app.getVersion()}\nالعنوان: ${APP_URL}\nتُحدَّث المنظومة تلقائياً من الإنترنت وتعمل دون اتصال بعد أول تشغيل.`
      }) }
    ] }
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(t));
}

app.on('second-instance', () => { if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });

app.whenReady().then(() => {
  // Google refuses sign-in from user agents that name Electron; present a plain Chrome user agent instead.
  const ua = session.defaultSession.getUserAgent()
    .replace(/ Electron\/\S+/, '')
    .replace(/\(KHTML, like Gecko\) .*?Chrome\//, '(KHTML, like Gecko) Chrome/');
  session.defaultSession.setUserAgent(ua);
  app.userAgentFallback = ua;
  // Camera (document scan) and notifications only for the system itself; everything else is refused.
  session.defaultSession.setPermissionRequestHandler((wc, permission, cb, details) => {
    const ok = ['media', 'notifications', 'clipboard-sanitized-write', 'fullscreen'].includes(permission);
    cb(ok && isAppUrl(details.requestingUrl || wc.getURL()));
  });
  session.defaultSession.setPermissionCheckHandler((wc, permission, origin) =>
    ['media', 'notifications', 'clipboard-sanitized-write', 'fullscreen'].includes(permission) && origin === APP_ORIGIN);
  app.on('web-contents-created', (e, wc) => wireContents(wc));
  buildMenu();
  createWindow();
});

app.on('window-all-closed', () => app.quit());
