// ==UserScript==
// @name         WP Admin Tools Panel — User DB / Order CHB / LOG
// @namespace    https://github.com/blackowl0192/script_folder
// @version      1.9.2
// @description  Единая панель для WP Admin: добавление юзера в БД, редактирование ордера ЧБ, редактирование ЛОГ
// @author       Black Owl
// @match        *://*/wp-admin/*
// @updateURL    https://raw.githubusercontent.com/blackowl0192/script_folder/main/wp-admin-tools.user.js
// @downloadURL  https://raw.githubusercontent.com/blackowl0192/script_folder/main/wp-admin-tools.user.js
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  /************************************************************
   * 0. ДОСТУПНЫЕ САЙТЫ
   * Скрипт запускается только на этих доменах и только внутри /wp-admin/
   ************************************************************/
  const ALLOWED_DOMAINS = [
    'easycourses.pro',
    'fastedu.pro',
    'courseplanet.shop',
    'helpersm.com',
    'smmfable.com',
    'onlysmm.pro',
    'zonefl.com',
    'freeskillup.com',
    'skilt.net',
    'mysundates.com',
    'feeldate.com',
    'prime2key.com',
    'playcodestore.com',
    'digitplaystore.com',
    'datelya.com',
    'edulecturers.com',
    'skillsmm.com',
    'learnway.shop',
    'studixio.com',
    'keymplace.com',
    'cs2void.com',
    'dota2void.com',
    'hottestcourse.com',
    'dotachamps.com',
    'cs2champs.com',
    'cs2money.shop',
    'yourlibrarium.com',
    'bestsocials.online',
    'profcourse.com',
    'playfun.pro',
    'cs2leader.com',
    'dota2leader.com',
    'dota2money.shop',
    'googolsell.com',
    'steambuy.pro',
    'cs2ultra.com',
    'dota2ultra.com',
    'megadota2.com',
    'megacs2.com',
    'cs2story.com',
    'dota2story.com',
    'cs2ultima.com',
    'cs2skin.pro',
    'dota2skin.pro'
  ];

  const currentHost = window.location.hostname.replace(/^www\./, '');
  const currentPath = window.location.pathname;

  if (!currentPath.includes('/wp-admin/')) return;
  if (!ALLOWED_DOMAINS.includes(currentHost)) return;

  /************************************************************
   * 1. ОБЩИЕ НАСТРОЙКИ ПАНЕЛИ
   ************************************************************/
  const APP_ID = 'bo-wp-admin-tools';
  const STORAGE_KEY = 'bo_wp_admin_tools_panel_state_v2';
  const USER_AUTOFILL_DONE_KEY = 'bo_crm_user_autofill_done';

  let alreadyAutoFilledUser = false;
  let alreadyRedirectedToEdit = false;

  function loadState() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {};
    } catch (e) {
      return {};
    }
  }

  function saveState(state) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      ...loadState(),
      ...state
    }));
  }

  const savedState = loadState();

  /************************************************************
   * 2. ОБЩИЕ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
   ************************************************************/
  function qs(selector, root = document) {
    return root.querySelector(selector);
  }

  function qsa(selector, root = document) {
    return Array.from(root.querySelectorAll(selector));
  }

  function setValue(selector, value) {
    const el = qs(selector);

    if (!el || value === undefined || value === null || value === '') {
      return false;
    }

    el.focus();
    el.value = value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new Event('blur', { bubbles: true }));

    return true;
  }

  function flashButton(btn, text = '✅ Готово', delay = 1500) {
    const oldText = btn.textContent;
    btn.textContent = text;
    setTimeout(() => {
      btn.textContent = oldText;
    }, delay);
  }

  function showStatus(text, type = 'info') {
    const status = qs('#bo-status');
    if (!status) return;

    status.textContent = text;
    status.className = `bo-status ${type}`;
  }


  /************************************************************
   * 2.1 СОХРАНЕНИЕ ДАННЫХ ПОЛЕЙ И НОРМАЛИЗАЦИЯ ДАТЫ
   ************************************************************/
  const FORM_STORAGE_KEY = 'bo_wp_admin_tools_form_values_v1';

  function loadFormValues() {
    try {
      return JSON.parse(localStorage.getItem(FORM_STORAGE_KEY)) || {};
    } catch (e) {
      return {};
    }
  }

  function saveFormValue(fieldId, value) {
    localStorage.setItem(FORM_STORAGE_KEY, JSON.stringify({
      ...loadFormValues(),
      [fieldId]: value
    }));
  }

  function restoreFormValues() {
    const values = loadFormValues();
    Object.keys(values).forEach(id => {
      const el = qs(`#${id}`, app);
      if (el && values[id] !== undefined && values[id] !== null) {
        el.value = values[id];
      }
    });
  }

  function bindPersistentField(id) {
    const el = qs(`#${id}`, app);
    if (!el) return;

    el.addEventListener('input', () => saveFormValue(id, el.value));
    el.addEventListener('change', () => saveFormValue(id, el.value));
  }

  /**
   * Приводит дату к формату DD.MM.YYYY HH:MM:SS.
   * Понимает варианты с лишними пробелами, запятыми, тире, слешами,
   * одинарными цифрами и ISO-подобный формат YYYY-MM-DD HH:MM:SS.
   */
  function normalizeDateTime(rawValue) {
    if (!rawValue) return null;

    let value = String(rawValue)
      .trim()
      .replace(/\s+/g, ' ')
      .replace(/[，,]+/g, ' ')
      .replace(/[\/\\]/g, '.')
      .replace(/\s*[–—-]\s*/g, '.')
      .replace(/\s*T\s*/i, ' ')
      .trim();

    // DD.MM.YYYY HH:MM[:SS]
    let match = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
    if (match) {
      const [, dd, mm, yyyy, hh, min, ss = '00'] = match;
      return `${dd.padStart(2, '0')}.${mm.padStart(2, '0')}.${yyyy} ${hh.padStart(2, '0')}:${min.padStart(2, '0')}:${ss.padStart(2, '0')}`;
    }

    // YYYY.MM.DD HH:MM[:SS]
    match = value.match(/^(\d{4})\.(\d{1,2})\.(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
    if (match) {
      const [, yyyy, mm, dd, hh, min, ss = '00'] = match;
      return `${dd.padStart(2, '0')}.${mm.padStart(2, '0')}.${yyyy} ${hh.padStart(2, '0')}:${min.padStart(2, '0')}:${ss.padStart(2, '0')}`;
    }

    return null;
  }

  /************************************************************
   * 3. CSS — ДИЗАЙН ПАНЕЛИ
   ************************************************************/
  const style = document.createElement('style');
  style.textContent = `
    #${APP_ID} {
      position: fixed;
      top: ${savedState.top || 80}px;
      left: ${savedState.left || 80}px;
      width: ${savedState.width || 560}px;
      height: ${savedState.height || 500}px;
      min-width: 390px;
      min-height: 280px;
      z-index: 999999;
      background: #111827;
      color: #f9fafb;
      border: 1px solid #374151;
      border-radius: 14px;
      box-shadow: 0 14px 42px rgba(0,0,0,0.38);
      font-family: Arial, sans-serif;
      resize: both;
      overflow: hidden;
      box-sizing: border-box;
    }

    #${APP_ID}.bo-collapsed {
      height: 48px !important;
      min-height: 48px;
      resize: none;
    }

    #${APP_ID}.bo-fullscreen {
      top: 0 !important;
      left: 0 !important;
      width: 100vw !important;
      height: 100vh !important;
      min-width: 0;
      min-height: 0;
      border-radius: 0;
      resize: none;
    }

    #${APP_ID} * {
      box-sizing: border-box;
    }

    .bo-header {
      height: 48px;
      background: linear-gradient(135deg, #1f2937, #111827);
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 12px 0 16px;
      cursor: move;
      border-bottom: 1px solid #374151;
      user-select: none;
    }

    .bo-title {
      font-size: 14px;
      font-weight: 700;
      letter-spacing: .2px;
    }

    .bo-header-buttons {
      display: flex;
      gap: 6px;
    }

    .bo-icon-btn {
      width: 28px;
      height: 28px;
      border: 1px solid #4b5563;
      background: #1f2937;
      color: #f9fafb;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      line-height: 1;
    }

    .bo-icon-btn:hover {
      background: #374151;
    }

    .bo-body {
      height: calc(100% - 48px);
      display: flex;
      flex-direction: column;
    }

    .bo-tabs {
      display: flex;
      gap: 6px;
      padding: 10px;
      background: #111827;
      border-bottom: 1px solid #374151;
    }

    .bo-tab {
      flex: 1;
      min-height: 38px;
      padding: 8px;
      border: 1px solid #374151;
      background: #1f2937;
      color: #d1d5db;
      border-radius: 10px;
      cursor: pointer;
      font-size: 13px;
      font-weight: 700;
      line-height: 1.2;
    }

    .bo-tab:hover {
      background: #374151;
    }

    .bo-tab.active {
      background: #2563eb;
      border-color: #3b82f6;
      color: white;
    }

    .bo-content {
      flex: 1;
      overflow: auto;
      padding: 14px;
      background: #0f172a;
    }

    .bo-panel {
      display: none;
    }

    .bo-panel.active {
      display: block;
    }

    .bo-section {
      background: #111827;
      border: 1px solid #374151;
      border-radius: 12px;
      padding: 14px;
      margin-bottom: 12px;
    }

    .bo-section-title {
      font-size: 14px;
      font-weight: 700;
      margin-bottom: 10px;
      color: #f3f4f6;
    }

    .bo-note {
      font-size: 13px;
      color: #9ca3af;
      line-height: 1.45;
      margin-bottom: 10px;
    }

    .bo-grid-2 {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }

    .bo-btn-row {
      display: flex;
      gap: 8px;
      margin-top: 8px;
    }

    .bo-btn {
      padding: 9px 12px;
      border: 0;
      border-radius: 10px;
      background: #2563eb;
      color: #fff;
      cursor: pointer;
      font-weight: 700;
      font-size: 13px;
    }

    .bo-btn:hover {
      background: #1d4ed8;
    }

    .bo-btn.green {
      background: #16a34a;
    }

    .bo-btn.green:hover {
      background: #15803d;
    }

    .bo-btn.purple {
      background: #7c3aed;
    }

    .bo-btn.purple:hover {
      background: #6d28d9;
    }

    .bo-btn.gray {
      background: #374151;
    }

    .bo-btn.gray:hover {
      background: #4b5563;
    }

    .bo-input,
    .bo-textarea {
      width: 100%;
      background: #020617;
      color: #f9fafb;
      border: 1px solid #374151;
      border-radius: 10px;
      padding: 9px 10px;
      margin-bottom: 10px;
      outline: none;
      font-size: 13px;
    }

    .bo-textarea {
      min-height: 130px;
      resize: vertical;
      font-family: Consolas, monospace;
    }

    .bo-label {
      display: block;
      font-size: 12px;
      color: #d1d5db;
      margin-bottom: 5px;
      font-weight: 700;
    }

    .bo-status {
      font-size: 12px;
      padding: 8px 10px;
      border-radius: 10px;
      margin: 0 14px 12px 14px;
      background: #1f2937;
      color: #d1d5db;
      border: 1px solid #374151;
    }

    .bo-status.ok {
      color: #bbf7d0;
      border-color: #166534;
    }

    .bo-status.warn {
      color: #fde68a;
      border-color: #92400e;
    }

    .bo-status.error {
      color: #fecaca;
      border-color: #991b1b;
    }
  `;
  document.head.appendChild(style);

  /************************************************************
   * 4. HTML — ПАНЕЛЬ И ВКЛАДКИ
   ************************************************************/
  const oldApp = qs(`#${APP_ID}`);
  if (oldApp) oldApp.remove();

  const app = document.createElement('div');
  app.id = APP_ID;

  app.innerHTML = `
    <div class="bo-header">
      <div class="bo-title">WP Admin Tools</div>
      <div class="bo-header-buttons">
        <button class="bo-icon-btn" id="bo-fullscreen-btn" title="Развернуть на весь экран">⛶</button>
        <button class="bo-icon-btn" id="bo-collapse-btn" title="Свернуть">—</button>
        <button class="bo-icon-btn" id="bo-close-btn" title="Закрыть">×</button>
      </div>
    </div>

    <div class="bo-body">
      <div class="bo-tabs">
        <button class="bo-tab active" data-tab="user-db">Добавление юзера в БД</button>
        <button class="bo-tab" data-tab="order-chb">Редактирование ордера ЧБ</button>
        <button class="bo-tab" data-tab="log-edit">Редактирование ЛОГ</button>
      </div>

      <div id="bo-status" class="bo-status">Готово. Домен разрешён: ${currentHost}</div>

      <div class="bo-content">
        <div class="bo-panel active" data-panel="user-db">
          <div class="bo-section">
            <div class="bo-section-title">Добавление юзера в БД</div>
            <div class="bo-note">
              Работает на страницах <b>user-new.php</b> и <b>user-edit.php</b>. Автоматическая подстановка отключена: данные вставляются из буфера только по кнопке или вручную.
            </div>

            <label class="bo-label">Данные пользователя</label>
            <textarea id="bo-user-data" class="bo-textarea" placeholder="ID юзера\nid923985978\nИмя\nMark Ivanova\nПочта\nmark.ivanova@testuserdb.com\nПароль\nFKwVtMNp7\nАдрес\n...\nГород\n...\nПосткод\n...\nСтрана\n...\nТелефон\n..."></textarea>

            <div class="bo-btn-row">
              <button class="bo-btn gray" id="bo-user-read-clipboard">Вставить из буфера</button>
              <button class="bo-btn green" id="bo-user-fill">Заполнить</button>
            </div>
          </div>
        </div>

        <div class="bo-panel" data-panel="order-chb">
          <div class="bo-section">
            <div class="bo-section-title">Товары ордера</div>
            <div class="bo-note">
              Можно вставлять напрямую из Excel: <b>Полная цена</b> + <b>Кол-во лекции</b> — для сайтов с курсами. Либо для сайтов с играми: <b>Расчетная цена</b> + <b>Кол-во игр</b>.
            </div>

            <label class="bo-label">Данные товаров из Excel или вручную</label>
            <textarea id="bo-order-items-data" class="bo-textarea" placeholder="Полная цена&#9;Кол-во лекции
126,00&#9;35,00
134,00&#9;25,00
90,00&#9;25,00"></textarea>

            <button class="bo-btn green" id="bo-order-items-apply" style="width:100%;">Изменить товары</button>
            <button class="bo-btn" id="bo-order-create" style="width:100%;margin-top:8px;">Создать ордер</button>
          </div>

          <div class="bo-section">
            <div class="bo-section-title">Редактирование ордера ЧБ</div>
            <div class="bo-note">
              Заполняет дату ордера, Transaction ID, ставит статус Completed и выполняет очистку ЧБ.
            </div>

            <label class="bo-label">Дата ордера</label>
            <input id="bo-order-date" class="bo-input" placeholder="21.10.2025 21:00:53">

            <label class="bo-label">Transaction ID</label>
            <input id="bo-order-tx" class="bo-input" placeholder="Transaction ID">

            <div class="bo-btn-row">
              <button class="bo-btn green" id="bo-order-apply">Применить</button>
              <button class="bo-btn" id="bo-order-update">Update</button>
            </div>

            <button class="bo-btn purple" id="bo-order-chb-fix" style="width:100%;margin-top:10px;">Изменить ЧБ</button>
          </div>
        </div>

        <div class="bo-panel" data-panel="log-edit">
          <div class="bo-section">
            <div class="bo-section-title">Редактирование ЛОГ</div>
            <div class="bo-note">
              Меняет IP, дату, дублирует строки и распределяет время по диапазону.
            </div>

            <label class="bo-label">IP</label>
            <input id="bo-log-ip" class="bo-input" placeholder="192.168.1.1">
            <button class="bo-btn" id="bo-log-change-ip" style="width:100%;margin-bottom:10px;">Изменить IP</button>

            <label class="bo-label">Дата</label>
            <input id="bo-log-date" class="bo-input" placeholder="16.10.2025 20:26:18">
            <button class="bo-btn" id="bo-log-change-date" style="width:100%;margin-bottom:10px;">Изменить дату</button>

            <label class="bo-label">Количество строк для дублирования</label>
            <input id="bo-log-duplicate-count" class="bo-input" type="number" value="16" min="1">
            <button class="bo-btn green" id="bo-log-duplicate" style="width:100%;margin-bottom:10px;">Clear & Duplicate</button>

            <div class="bo-grid-2">
              <div>
                <label class="bo-label">Время от, 24ч</label>
                <input id="bo-log-time-from" class="bo-input" type="text" value="10:15:00" placeholder="10:15:00">
              </div>
              <div>
                <label class="bo-label">Время до, 24ч</label>
                <input id="bo-log-time-to" class="bo-input" type="text" value="12:20:00" placeholder="12:20:00">
              </div>
            </div>

            <label class="bo-label">Created — время из полной даты</label>
            <input id="bo-log-created-datetime" class="bo-input" type="text" placeholder="02.11.2025 22:03:58">
            <div class="bo-note" style="margin-bottom:10px;">
              Можно вставить полную дату. Скрипт удалит лишние пробелы, отбросит дату и возьмёт только время.
            </div>

            <button class="bo-btn purple" id="bo-log-time-range" style="width:100%;">Применить диапазон</button>
          </div>
        </div>
      </div>
    </div>
  `;

  document.body.appendChild(app);

  /************************************************************
   * 5. ПЕРЕКЛЮЧЕНИЕ ВКЛАДОК
   ************************************************************/
  const tabs = qsa('.bo-tab', app);
  const panels = qsa('.bo-panel', app);

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const target = tab.dataset.tab;

      tabs.forEach(t => t.classList.remove('active'));
      panels.forEach(p => p.classList.remove('active'));

      tab.classList.add('active');
      qs(`[data-panel="${target}"]`, app).classList.add('active');
      saveState({ activeTab: target });
    });
  });

  if (savedState.activeTab) {
    const savedTab = qs(`[data-tab="${savedState.activeTab}"]`, app);
    if (savedTab) savedTab.click();
  }

  /************************************************************
   * 6. ПЕРЕТАСКИВАНИЕ ПАНЕЛИ
   ************************************************************/
  const header = qs('.bo-header', app);
  let isDragging = false;
  let offsetX = 0;
  let offsetY = 0;

  header.addEventListener('mousedown', (e) => {
    if (e.target.tagName === 'BUTTON') return;
    if (app.classList.contains('bo-fullscreen')) return;

    isDragging = true;
    offsetX = e.clientX - app.offsetLeft;
    offsetY = e.clientY - app.offsetTop;
    document.body.style.userSelect = 'none';
  });

  document.addEventListener('mousemove', (e) => {
    if (!isDragging) return;

    const left = Math.max(0, e.clientX - offsetX);
    const top = Math.max(0, e.clientY - offsetY);

    app.style.left = `${left}px`;
    app.style.top = `${top}px`;
  });

  document.addEventListener('mouseup', () => {
    if (!isDragging) return;

    isDragging = false;
    document.body.style.userSelect = '';

    saveState({
      left: app.offsetLeft,
      top: app.offsetTop
    });
  });

  /************************************************************
   * 7. РЕСАЙЗ + СВЕРНУТЬ / ЗАКРЫТЬ
   ************************************************************/
  const resizeObserver = new ResizeObserver(() => {
    if (app.classList.contains('bo-collapsed')) return;
    if (app.classList.contains('bo-fullscreen')) return;

    saveState({
      width: app.offsetWidth,
      height: app.offsetHeight
    });
  });
  resizeObserver.observe(app);

  const fullscreenBtn = qs('#bo-fullscreen-btn', app);
  const collapseBtn = qs('#bo-collapse-btn', app);
  const closeBtn = qs('#bo-close-btn', app);

  fullscreenBtn.addEventListener('click', () => {
    if (app.classList.contains('bo-collapsed')) {
      app.classList.remove('bo-collapsed');
      collapseBtn.textContent = '—';
      saveState({ collapsed: false });
    }

    app.classList.toggle('bo-fullscreen');
    const fullscreen = app.classList.contains('bo-fullscreen');
    fullscreenBtn.textContent = fullscreen ? '❐' : '⛶';
    fullscreenBtn.title = fullscreen ? 'Вернуть прежний размер' : 'Развернуть на весь экран';
    saveState({ fullscreen });
  });

  collapseBtn.addEventListener('click', () => {
    app.classList.toggle('bo-collapsed');
    const collapsed = app.classList.contains('bo-collapsed');
    collapseBtn.textContent = collapsed ? '+' : '—';
    saveState({ collapsed });
  });

  closeBtn.addEventListener('click', () => {
    app.remove();
  });

  if (savedState.collapsed) {
    app.classList.add('bo-collapsed');
    collapseBtn.textContent = '+';
  }

  if (savedState.fullscreen && !savedState.collapsed) {
    app.classList.add('bo-fullscreen');
    fullscreenBtn.textContent = '❐';
    fullscreenBtn.title = 'Вернуть прежний размер';
  }

  /************************************************************
   * 8. ВКЛАДКА 1 — ДОБАВЛЕНИЕ ЮЗЕРА В БД
   ************************************************************/
  function parseUserData(data) {
    const lines = String(data || '')
      .trim()
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);

    const result = {};

    for (let i = 0; i < lines.length; i++) {
      const key = lines[i].toLowerCase();
      const value = lines[i + 1];
      if (!value) continue;

      if (key.includes('id юзера')) {
        result.id = value;
        i++;
      } else if (key.includes('имя')) {
        result.name = value;
        i++;
      } else if (key.includes('почта')) {
        result.email = value;
        i++;
      } else if (key.includes('пароль')) {
        result.password = value;
        i++;
      } else if (key.includes('адрес')) {
        result.address = value;
        i++;
      } else if (key.includes('город')) {
        result.city = value;
        i++;
      } else if (key.includes('пост код') || key.includes('посткод')) {
        result.postcode = value;
        i++;
      } else if (key.includes('страна')) {
        result.state = value;
        result.country = value;
        i++;
      } else if (key.includes('телефон')) {
        result.phone = value;
        i++;
      }
    }

    return result;
  }

  function isValidClipboardUserData(text) {
    return Boolean(
      text &&
      text.includes('ID юзера') &&
      text.includes('Имя') &&
      text.includes('Почта') &&
      text.includes('Пароль')
    );
  }

  function fillUserForm(rawData) {
    const isUserNewPage = currentPath === '/wp-admin/user-new.php';
    const isUserEditPage = currentPath === '/wp-admin/user-edit.php';

    if (!isUserNewPage && !isUserEditPage) {
      showStatus('Вкладка юзера работает только на user-new.php или user-edit.php', 'warn');
      return false;
    }

    if (!rawData || !rawData.trim()) {
      showStatus('Нет данных пользователя для заполнения', 'warn');
      return false;
    }

    const userData = parseUserData(rawData);

    if (!userData.id && !userData.email && !userData.password) {
      showStatus('Данные не похожи на данные пользователя', 'error');
      return false;
    }

    const fullName = userData.name || '';
    const nameParts = fullName.split(' ').filter(Boolean);
    const firstName = nameParts[0] || '';
    const lastName = nameParts.slice(1).join(' ') || '';

    setValue('input[name="user_login"]', userData.id);
    setValue('input[name="email"]', userData.email);
    setValue('input[name="first_name"]', firstName);
    setValue('input[name="last_name"]', lastName);
    setValue('input[name="pass1"]', userData.password);

    setValue('input[name="billing_first_name"]', firstName);
    setValue('input[name="billing_last_name"]', lastName);
    setValue('input[name="billing_email"]', userData.email);
    setValue('input[name="billing_address_1"]', userData.address);
    setValue('input[name="billing_city"]', userData.city);
    setValue('input[name="billing_postcode"]', userData.postcode);
    setValue('input[name="billing_state"]', userData.state);
    setValue('input[name="billing_phone"]', userData.phone);

    setValue('input[name="shipping_first_name"]', firstName);
    setValue('input[name="shipping_last_name"]', lastName);
    setValue('input[name="shipping_address_1"]', userData.address);
    setValue('input[name="shipping_city"]', userData.city);
    setValue('input[name="shipping_postcode"]', userData.postcode);
    setValue('input[name="shipping_state"]', userData.state);

    const notificationCheckbox = qs('input[name="send_user_notification"]');
    if (notificationCheckbox) {
      notificationCheckbox.checked = false;
      notificationCheckbox.dispatchEvent(new Event('change', { bubbles: true }));
    }

    setTimeout(() => {
      const copyBillingBtn = qs('#copy_billing');
      if (copyBillingBtn) copyBillingBtn.click();
    }, 300);

    if (isUserEditPage) {
      sessionStorage.setItem(USER_AUTOFILL_DONE_KEY, '1');
    }

    showStatus('Форма пользователя заполнена', 'ok');
    return true;
  }

  async function readUserClipboardToTextarea() {
    const textarea = qs('#bo-user-data', app);

    try {
      const clipboardText = await navigator.clipboard.readText();

      if (!clipboardText) {
        showStatus('Буфер обмена пустой', 'warn');
        return false;
      }

      textarea.value = clipboardText;
      saveFormValue('bo-user-data', clipboardText);
      showStatus('Данные вставлены из буфера', 'ok');
      return true;
    } catch (err) {
      showStatus('Не удалось прочитать буфер обмена. Вставь данные вручную.', 'error');
      console.warn(err);
      return false;
    }
  }

  async function autoFillUserFromClipboard() {
    const isUserNewPage = currentPath === '/wp-admin/user-new.php';
    const isUserEditPage = currentPath === '/wp-admin/user-edit.php';

    if (!isUserNewPage && !isUserEditPage) return;
    if (alreadyAutoFilledUser) return;

    if (isUserEditPage && sessionStorage.getItem(USER_AUTOFILL_DONE_KEY) === '1') {
      showStatus('Edit user уже был заполнен. Повторный автозапуск остановлен.', 'warn');
      return;
    }

    try {
      const clipboardText = await navigator.clipboard.readText();

      if (!isValidClipboardUserData(clipboardText)) return;

      qs('#bo-user-data', app).value = clipboardText;

      if (fillUserForm(clipboardText)) {
        alreadyAutoFilledUser = true;
      }
    } catch (err) {
      console.warn('Не удалось автоматически прочитать буфер обмена:', err);
    }
  }

  function autoOpenEditUserAfterCreate() {
    const isUserNewPage = currentPath === '/wp-admin/user-new.php';
    if (!isUserNewPage || alreadyRedirectedToEdit) return;

    const editUserLink = qsa('a').find(a => {
      const text = a.textContent.trim().toLowerCase();
      return text.includes('edit user') && a.href && a.href.includes('user-edit.php');
    });

    if (!editUserLink) return;

    alreadyRedirectedToEdit = true;
    showStatus('Найдена ссылка Edit user. Выполняю переход.', 'ok');

    setTimeout(() => {
      window.location.href = editUserLink.href;
    }, 700);
  }

  function initUserTab() {
    const isUserNewPage = currentPath === '/wp-admin/user-new.php';
    const isUserEditPage = currentPath === '/wp-admin/user-edit.php';

    if (isUserNewPage) {
      sessionStorage.removeItem(USER_AUTOFILL_DONE_KEY);
    }

    qs('#bo-user-read-clipboard', app).addEventListener('click', async (e) => {
      const ok = await readUserClipboardToTextarea();
      if (ok) flashButton(e.currentTarget, '✅ Вставлено');
    });

    qs('#bo-user-fill', app).addEventListener('click', (e) => {
      const rawData = qs('#bo-user-data', app).value;
      if (fillUserForm(rawData)) flashButton(e.currentTarget, '✅ Заполнено');
    });

    // Автоматическое чтение буфера и автоподстановка отключены.
    // Заполнение запускается только вручную кнопками во вкладке.

    setTimeout(autoOpenEditUserAfterCreate, 1000);
    setTimeout(autoOpenEditUserAfterCreate, 2500);
    setTimeout(autoOpenEditUserAfterCreate, 4500);
  }

  /************************************************************
   * 9. ВКЛАДКА 2 — РЕДАКТИРОВАНИЕ ОРДЕРА ЧБ
   ************************************************************/
  function applyOrderData() {
    const dateInput = qs('#bo-order-date', app);
    const txInput = qs('#bo-order-tx', app);
    const normalizedDate = normalizeDateTime(dateInput.value);

    if (!normalizedDate) {
      showStatus('Формат даты должен быть DD.MM.YYYY HH:MM:SS', 'error');
      return false;
    }

    dateInput.value = normalizedDate;
    saveFormValue('bo-order-date', normalizedDate);

    const match = normalizedDate.match(/^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
    const [, d, m, y, h, min, s] = match;
    const formattedDate = `${y}-${m}-${d}`;

    setValue('[name="order_date"]', formattedDate);
    setValue('[name="order_date_hour"]', String(Number(h)));
    setValue('[name="order_date_minute"]', String(Number(min)));
    setValue('[name="order_date_second"]', String(Number(s)));

    if (txInput.value.trim()) {
      setValue('#_transaction_id', txInput.value.trim());
    }

    const status = qs('#order_status');
    if (status) {
      status.value = 'wc-completed';
      status.dispatchEvent(new Event('change', { bubbles: true }));
    }

    const originalConfirm = window.confirm;
    window.confirm = () => true;
    qs('.billing-same-as-shipping')?.click();

    setTimeout(() => {
      window.confirm = originalConfirm;
    }, 500);

    showStatus('Данные ордера применены', 'ok');
    return true;
  }

  function updateOrder() {
    const updateBtn = qs('button[name="save"][value="Update"]');
    if (!updateBtn) {
      showStatus('Кнопка Update не найдена', 'warn');
      return false;
    }

    updateBtn.click();
    showStatus('Нажата кнопка Update', 'ok');
    return true;
  }

  /**
   * Нажимает WooCommerce Create для создания нового ордера.
   * Используется отдельной кнопкой "Создать ордер" во вкладке ордера ЧБ.
   */
  function createOrder() {
    const createBtn = qs('button.save_order[name="save"][value="Create"], button.save_order[value="Create"]');

    if (!createBtn) {
      showStatus('Кнопка Create не найдена', 'warn');
      return false;
    }

    createBtn.click();
    showStatus('Нажата кнопка Create', 'ok');
    return true;
  }

  /**
   * Генерирует случайный 10-значный номер заказа.
   * Первая цифра 1-9, чтобы номер не начинался с нуля.
   */
  function generateTenDigitOrderNumber() {
    return String(Math.floor(1000000000 + Math.random() * 9000000000));
  }

  /**
   * Если в заголовке WooCommerce номер заказа короче 9 цифр,
   * заменяет его на случайный номер из 10 цифр.
   *
   * Пример заголовка:
   * <h2 class="woocommerce-order-data__heading">Order #27335 details</h2>
   */
  function fixShortOrderHeadingNumber() {
    const heading = qs('h2.woocommerce-order-data__heading');

    if (!heading) {
      showStatus('Заголовок номера ордера не найден', 'warn');
      return false;
    }

    const originalText = heading.textContent || '';
    const match = originalText.match(/#\s*(\d+)/);

    if (!match) {
      showStatus('Номер ордера в заголовке не найден', 'warn');
      return false;
    }

    const currentNumber = match[1];

    // Если номер уже 9 цифр или больше — ничего не меняем.
    if (currentNumber.length >= 9) {
      showStatus(`Номер ордера не изменён: ${currentNumber}`, 'info');
      return true;
    }

    const newNumber = generateTenDigitOrderNumber();

    heading.textContent = originalText.replace(/#\s*\d+/, `#${newNumber}`);

    showStatus(`Номер ордера заменён: ${currentNumber} → ${newNumber}`, 'ok');
    return true;
  }

  function fixChbOrderView() {
    // 0. Если номер ордера в заголовке короче 9 цифр — заменить на 10-значный
    fixShortOrderHeadingNumber();

    // 1. Удалить скидки
    qsa('.wc-order-item-discount').forEach(e => e.remove());

    // 2. Удалить строки оплаты с via Payment / via Card
    qsa('.description').forEach(el => {
      const text = el.textContent.replace(/\s+/g, ' ').trim();
      if (/\bvia\b/i.test(text) && /(payment|card)/i.test(text)) {
        el.remove();
      }
    });

    // 3. Очистка мета ордера
    const meta = qs('.woocommerce-order-data__meta.order_number');
    if (meta) {
      let text = meta.textContent;
      text = text.replace(/\.\s*Paid on.*$/i, '');
      text = text.replace(/Customer IP.*$/i, '');
      text = text.replace('Payment with card', 'Payment by card');
      meta.textContent = text.trim();
    }

    // 4. Очистка email/ID клиента в select2
    const clearBtn = qs('.select2-selection__clear');
    if (clearBtn && clearBtn.nextSibling) {
      clearBtn.nextSibling.textContent = clearBtn.nextSibling.textContent.replace(/#\d+\s–\s/, '');
    }

    showStatus('ЧБ-очистка ордера выполнена', 'ok');
    return true;
  }


  /**
   * Нормализует цену для WooCommerce в формат 10,74.
   * Убирает пробелы, символы валют, заменяет точку на запятую.
   */
  function normalizeOrderItemPrice(value) {
    if (value === undefined || value === null) return null;

    let cleaned = String(value)
      .trim()
      .replace(/\s+/g, '')
      .replace(/[€$£₽]/g, '')
      .replace('.', ',');

    if (!/^\d+(,\d{1,2})?$/.test(cleaned)) {
      return null;
    }

    if (!cleaned.includes(',')) {
      cleaned += ',00';
    }

    const [whole, cents = '00'] = cleaned.split(',');
    return `${whole},${cents.padEnd(2, '0')}`;
  }

  /**
   * Устанавливает значение в input и обновляет связанные data-атрибуты WooCommerce.
   */
  function setOrderItemInputValue(input, value) {
    if (!input) return false;

    input.focus();
    input.value = value;
    input.setAttribute('value', value);

    if (input.dataset.qty !== undefined) input.dataset.qty = value;
    if (input.dataset.subtotal !== undefined) input.dataset.subtotal = value;
    if (input.dataset.total !== undefined) input.dataset.total = value;

    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    input.dispatchEvent(new Event('blur', { bubbles: true }));

    return true;
  }

  /**
   * Проверяет, похоже ли значение на цену/число из Excel.
   */
  function isNumericLike(value) {
    const cleaned = String(value || '')
      .trim()
      .replace(/\s+/g, '')
      .replace(/[€$]/g, '')
      .replace(',', '.');

    return /^\d+(\.\d+)?$/.test(cleaned);
  }

  /**
   * Нормализует количество лекций/товаров из Excel.
   * Пример: "35,00" -> "35".
   */
  function normalizeOrderItemQty(value) {
    if (value === undefined || value === null) return null;

    const cleaned = String(value)
      .trim()
      .replace(/\s+/g, '')
      .replace(',', '.');

    if (!/^\d+(\.\d+)?$/.test(cleaned)) return null;

    const num = Number(cleaned);
    if (!Number.isFinite(num) || num < 0) return null;

    // WooCommerce quantity должен быть целым числом.
    return String(Math.round(num));
  }

  /**
   * Парсит данные товаров из textarea.
   * Поддерживает 3 формата:
   * 1) Excel: Полная цена<TAB>Кол-во лекции
   * 2) Ручной: кол-во | сумма
   * 3) Старый: две строки на товар: количество, затем сумма.
   */
  function parseOrderItemsData(rawText) {
    const lines = String(rawText || '')
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);

    if (!lines.length) {
      return { error: 'Вставьте данные товаров.' };
    }

    const items = [];
    const hasInlineSeparator = lines.some(line => /[|;\t]/.test(line));

    if (hasInlineSeparator) {
      let mode = 'manual_qty_total';
      let startIndex = 0;

      const firstParts = lines[0]
        .split(/[|;\t]/)
        .map(part => part.trim())
        .filter(Boolean);

      const firstLineLower = lines[0].toLowerCase();

      // Excel-вставка с заголовком: Полная цена | Кол-во лекции.
      if (
        firstLineLower.includes('полная цена') ||
        firstLineLower.includes('кол-во') ||
        firstLineLower.includes('лекци') ||
        firstLineLower.includes('lessons') ||
        firstLineLower.includes('price')
      ) {
        mode = 'excel_price_qty';
        startIndex = 1;
      }

      // Если заголовка нет, но первая колонка похожа на цену, а вторая на количество,
      // считаем это Excel-форматом: цена TAB количество.
      if (startIndex === 0 && firstParts.length >= 2) {
        const first = firstParts[0];
        const second = firstParts[1];
        const firstNum = Number(String(first).replace(',', '.'));
        const secondNum = Number(String(second).replace(',', '.'));

        if (
          isNumericLike(first) &&
          isNumericLike(second) &&
          firstNum > secondNum &&
          String(first).includes(',')
        ) {
          mode = 'excel_price_qty';
        }
      }

      for (let i = startIndex; i < lines.length; i++) {
        const parts = lines[i]
          .split(/[|;\t]/)
          .map(part => part.trim())
          .filter(Boolean);

        if (parts.length < 2) {
          return { error: `Ошибка в строке ${i + 1}. Нужно минимум 2 значения.` };
        }

        let qty;
        let total;

        if (mode === 'excel_price_qty') {
          // Excel: Полная цена | Кол-во лекции
          total = normalizeOrderItemPrice(parts[0]);
          qty = normalizeOrderItemQty(parts[1]);
        } else {
          // Ручной формат: Кол-во | Сумма
          qty = normalizeOrderItemQty(parts[0]);
          total = normalizeOrderItemPrice(parts[1]);
        }

        if (!qty) {
          return { error: `Ошибка количества в строке ${i + 1}. Значение: ${mode === 'excel_price_qty' ? parts[1] : parts[0]}` };
        }

        if (!total) {
          return { error: `Ошибка суммы в строке ${i + 1}. Значение: ${mode === 'excel_price_qty' ? parts[0] : parts[1]}` };
        }

        items.push({ qty, total });
      }

      return { items };
    }

    // Старый формат: две строки на товар: количество, сумма.
    if (lines.length % 2 !== 0) {
      return { error: 'Неверный формат. Вставьте данные из Excel или используйте: кол-во | сумма.' };
    }

    for (let i = 0; i < lines.length; i += 2) {
      const qty = normalizeOrderItemQty(lines[i]);
      const total = normalizeOrderItemPrice(lines[i + 1]);

      if (!qty) {
        return { error: `Ошибка количества в строке ${i + 1}: ${lines[i]}` };
      }

      if (!total) {
        return { error: `Ошибка суммы в строке ${i + 2}: ${lines[i + 1]}` };
      }

      items.push({ qty, total });
    }

    return { items };
  }

  /**
   * Массово меняет товары в WooCommerce order items сверху вниз.
   * Меняет Quantity, Before discount и Total.
   */
  function applyOrderItemsFromText(rawText) {
    const parsed = parseOrderItemsData(rawText);

    if (parsed.error) {
      showStatus(parsed.error, 'error');
      return false;
    }

    const rows = qsa('#order_line_items tr.item');

    if (!rows.length) {
      showStatus('Товары в ордере не найдены.', 'error');
      return false;
    }

    const items = parsed.items;

    if (items.length > rows.length) {
      showStatus(`Вставлено ${items.length} строк, а товаров в ордере только ${rows.length}.`, 'error');
      return false;
    }

    let updated = 0;

    items.forEach((item, index) => {
      const row = rows[index];
      const orderItemId = row.dataset.order_item_id;

      const qtyInput = row.querySelector(`input[name="order_item_qty[${orderItemId}]"]`);
      const subtotalInput = row.querySelector(`input[name="line_subtotal[${orderItemId}]"]`);
      const totalInput = row.querySelector(`input[name="line_total[${orderItemId}]"]`);

      const okQty = setOrderItemInputValue(qtyInput, item.qty);
      const okSubtotal = setOrderItemInputValue(subtotalInput, item.total);
      const okTotal = setOrderItemInputValue(totalInput, item.total);

      if (okQty || okSubtotal || okTotal) updated++;
    });

    showStatus(`Товары обновлены: ${updated} из ${items.length}. Запускаю Recalculate...`, 'ok');
    return true;
  }

  /**
   * Нажимает кнопку WooCommerce Recalculate после изменения товаров.
   * Сделано с небольшой задержкой, чтобы WooCommerce успел получить input/change события.
   */
  function clickOrderRecalculate(afterDone) {
    const recalculateBtn = qs('button.calculate-action');

    if (!recalculateBtn) {
      showStatus('Кнопка Recalculate не найдена. Проверьте товары и нажмите Update/Create вручную.', 'warn');
      if (typeof afterDone === 'function') afterDone(false);
      return false;
    }

    showStatus('Запущен Recalculate...', 'info');

    setTimeout(() => {
      recalculateBtn.click();

      let checks = 0;
      const maxChecks = 40;

      const timer = setInterval(() => {
        checks++;

        const loading =
          qs('.blockUI') ||
          qs('.woocommerce-layout__loading') ||
          qs('.spinner.is-active') ||
          recalculateBtn.disabled;

        if (!loading || checks >= maxChecks) {
          clearInterval(timer);
          showStatus('Recalculate выполнен. Теперь можно нажать Update или Создать ордер.', 'ok');
          if (typeof afterDone === 'function') afterDone(true);
        }
      }, 300);
    }, 300);

    return true;
  }

  function initOrderTab() {
    qs('#bo-order-apply', app).addEventListener('click', (e) => {
      if (applyOrderData()) flashButton(e.currentTarget, '✅');
    });

    qs('#bo-order-update', app).addEventListener('click', (e) => {
      if (updateOrder()) flashButton(e.currentTarget, '✅');
    });

    qs('#bo-order-chb-fix', app).addEventListener('click', (e) => {
      if (fixChbOrderView()) flashButton(e.currentTarget, '✅');
    });

    qs('#bo-order-items-apply', app).addEventListener('click', (e) => {
      const btn = e.currentTarget;
      const rawText = qs('#bo-order-items-data', app).value;

      if (applyOrderItemsFromText(rawText)) {
        btn.textContent = 'Recalculate...';
        clickOrderRecalculate(() => {
          flashButton(btn, '✅ Готово', 2000);
        });
      }
    });

    qs('#bo-order-create', app).addEventListener('click', (e) => {
      if (createOrder()) flashButton(e.currentTarget, '✅ Create');
    });
  }

  /************************************************************
   * 10. ВКЛАДКА 3 — РЕДАКТИРОВАНИЕ ЛОГ
   ************************************************************/
  function timeToSeconds(t) {
    const parts = String(t).split(':').map(Number);
    const h = parts[0] || 0;
    const m = parts[1] || 0;
    const sec = parts[2] || 0;
    return h * 3600 + m * 60 + sec;
  }

  function secondsToTime(s) {
    return [
      Math.floor(s / 3600),
      Math.floor((s % 3600) / 60),
      s % 60
    ].map(v => String(v).padStart(2, '0')).join(':');
  }

  /**
   * Нормализует время в 24-часовой формат HH:MM:SS.
   * Принимает:
   * - 2130       -> 21:30:00
   * - 213058     -> 21:30:58
   * - 21:30      -> 21:30:00
   * - 21:30:58   -> 21:30:58
   * - 02.11.2025 22:03:58 -> 22:03:58
   * - строки с лишними пробелами.
   */
  function normalizeTime24(rawValue) {
    if (!rawValue) return null;

    const value = String(rawValue)
      .trim()
      .replace(/\s+/g, ' ')
      .replace(/[,，]+/g, ' ')
      .trim();

    // Если есть дата + время, берём последнее найденное время.
    const matches = value.match(/(\d{1,2})\s*:\s*(\d{1,2})(?:\s*:\s*(\d{1,2}))?/g);

    let h;
    let m;
    let sec;

    if (matches && matches.length) {
      const time = matches[matches.length - 1].replace(/\s+/g, '');
      const parts = time.split(':').map(Number);
      h = parts[0];
      m = parts[1];
      sec = parts.length > 2 ? parts[2] : 0;
    } else {
      // Короткий ввод без двоеточий: 2130 -> 21:30:00, 213058 -> 21:30:58.
      const digits = value.replace(/\D/g, '');

      if (digits.length === 3) {
        h = Number(digits.slice(0, 1));
        m = Number(digits.slice(1, 3));
        sec = 0;
      } else if (digits.length === 4) {
        h = Number(digits.slice(0, 2));
        m = Number(digits.slice(2, 4));
        sec = 0;
      } else if (digits.length === 5) {
        h = Number(digits.slice(0, 1));
        m = Number(digits.slice(1, 3));
        sec = Number(digits.slice(3, 5));
      } else if (digits.length === 6) {
        h = Number(digits.slice(0, 2));
        m = Number(digits.slice(2, 4));
        sec = Number(digits.slice(4, 6));
      } else {
        return null;
      }
    }

    if (
      Number.isNaN(h) || Number.isNaN(m) || Number.isNaN(sec) ||
      h < 0 || h > 23 ||
      m < 0 || m > 59 ||
      sec < 0 || sec > 59
    ) {
      return null;
    }

    return [h, m, sec].map(v => String(v).padStart(2, '0')).join(':');
  }

  function normalizeLogTimeInput(inputId) {
    const input = qs(`#${inputId}`, app);
    if (!input) return null;

    const normalized = normalizeTime24(input.value);
    if (!normalized) return null;

    input.value = normalized;
    saveFormValue(inputId, normalized);
    return normalized;
  }

  function updateDateCell(val) {
    qsa('td.crtd.column-crtd[data-colname="Date"]').forEach(td => {
      const d = new Date(val);
      if (isNaN(d)) return;

      td.innerHTML =
        d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }) + '<br>' +
        d.toLocaleTimeString('en-GB', {
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
          hour12: false
        });
    });
  }

  function updateLogIP(ip) {
    qsa('td.scip.column-scip[data-colname="IP"]').forEach(td => {
      const a = qs('a', td);
      if (!a) return;

      a.textContent = ip;
      a.href = `https://whatismyipaddress.com/ip/${ip}?utm_source=plugin&utm_medium=referral&utm_campaign=wsal`;
    });
  }

  function getDatePartFromDateCell(td) {
    if (!td) return '';
    return (td.innerHTML || '').split('<br>')[0].trim();
  }

  function setLogRowTime(row, timeValue) {
    const dateCell = qs('td.crtd.column-crtd', row);
    if (!dateCell) return false;

    const datePart = getDatePartFromDateCell(dateCell) || new Date().toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });

    dateCell.innerHTML = `${datePart}<br>${timeValue}`;
    return true;
  }

  function setLogRowEvent(row, eventName) {
    const eventCell = qs('.event_type', row);
    if (eventCell) {
      eventCell.textContent = eventName;
      return true;
    }

    const fallback = qsa('td', row).find(td => /Viewed|Modified|Created/i.test(td.textContent || ''));
    if (fallback) {
      fallback.textContent = eventName;
      return true;
    }

    return false;
  }

  function getLogRowTimeSeconds(row) {
    const dateCell = qs('td.crtd.column-crtd', row);
    if (!dateCell) return -1;

    const html = dateCell.innerHTML || '';
    const timePart = html.split('<br>')[1] || dateCell.textContent || '';
    const normalized = normalizeTime24(timePart);
    if (!normalized) return -1;

    return timeToSeconds(normalized);
  }

  function sortLogRowsByTimeDesc(tableBody) {
    const rows = qsa('tr', tableBody);
    rows
      .sort((a, b) => getLogRowTimeSeconds(b) - getLogRowTimeSeconds(a))
      .forEach(row => tableBody.appendChild(row));
  }

  function addCreatedLogRowAtTime(createdTime) {
    if (!createdTime) return true;

    const tableBody = qs('table tbody');
    if (!tableBody) {
      showStatus('Таблица ЛОГ не найдена для Created', 'warn');
      return false;
    }

    const template = qs('tr', tableBody);
    if (!template) {
      showStatus('Нет строки-шаблона для Created', 'warn');
      return false;
    }

    const clone = template.cloneNode(true);
    setLogRowTime(clone, createdTime);
    setLogRowEvent(clone, 'Created');
    tableBody.appendChild(clone);
    sortLogRowsByTimeDesc(tableBody);

    return true;
  }

  function applyIncreasingTimes(minT, maxT, createdT = '') {
    const minTime = normalizeTime24(minT);
    const maxTime = normalizeTime24(maxT);
    const createdTime = createdT ? normalizeTime24(createdT) : '';

    if (!minTime || !maxTime) {
      showStatus('Введите время ОТ и ДО в формате HH:MM или HH:MM:SS', 'error');
      return false;
    }

    const min = timeToSeconds(minTime);
    const max = timeToSeconds(maxTime);

    if (max <= min) {
      showStatus('Время ДО должно быть больше времени ОТ', 'error');
      return false;
    }

    if (createdT && !createdTime) {
      showStatus('Created-время должно быть в формате DD.MM.YYYY HH:MM:SS или HH:MM:SS', 'error');
      return false;
    }

    if (createdTime) {
      const createdSeconds = timeToSeconds(createdTime);
      if (createdSeconds < min || createdSeconds > max) {
        showStatus('Created-время должно попадать в интервал ОТ–ДО', 'error');
        return false;
      }
    }

    const rows = qsa('td.crtd.column-crtd').reverse();
    if (!rows.length) {
      showStatus('Строки даты в таблице ЛОГ не найдены', 'warn');
      return false;
    }

    const step = Math.floor((max - min) / Math.max(rows.length - 1, 1));
    let last = min;

    rows.forEach((td, i) => {
      const [date] = td.innerHTML.split('<br>');
      let candidate = min + step * i + Math.floor(Math.random() * 31);

      if (candidate <= last) candidate = last + 1;
      if (candidate > max) candidate = max;

      last = candidate;
      td.innerHTML = `${date}<br>${secondsToTime(candidate)}`;
    });

    if (createdTime) {
      addCreatedLogRowAtTime(createdTime);
    }

    showStatus(createdTime ? 'Диапазон применён, Created добавлен' : 'Диапазон времени применён', 'ok');
    return true;
  }

  function getRandomEventType() {
    const r = Math.random() * 100;

    // Created больше не участвует в общей автогенерации.
    // Created создаётся только отдельной строкой через поле "Created — время из полной даты".
    return r < 70 ? 'Viewed' : 'Modified';
  }

  function changeLogIP() {
    const ip = qs('#bo-log-ip', app).value.trim();

    if (!/^(\d{1,3}\.){3}\d{1,3}$/.test(ip)) {
      showStatus('Введите корректный IPv4', 'error');
      return false;
    }

    updateLogIP(ip);
    showStatus('IP изменён', 'ok');
    return true;
  }

  function changeLogDate() {
    const dateInput = qs('#bo-log-date', app);
    const normalizedDate = normalizeDateTime(dateInput.value);

    if (!normalizedDate) {
      showStatus('Введите дату в формате DD.MM.YYYY HH:MM:SS', 'error');
      return false;
    }

    dateInput.value = normalizedDate;
    saveFormValue('bo-log-date', normalizedDate);

    const match = normalizedDate.match(/^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
    const [, dd, mm, yyyy, hh, min, ss] = match;
    updateDateCell(`${yyyy}-${mm}-${dd}T${hh}:${min}:${ss}`);
    showStatus('Дата в ЛОГ изменена', 'ok');
    return true;
  }

  function duplicateLogRows() {
    const count = parseInt(qs('#bo-log-duplicate-count', app).value, 10);

    if (isNaN(count) || count <= 0) {
      showStatus('Введите корректное количество строк', 'error');
      return false;
    }

    const tableBody = qs('table tbody');
    if (!tableBody) {
      showStatus('Таблица ЛОГ не найдена', 'warn');
      return false;
    }

    const rows = qsa('tr', tableBody);
    if (!rows.length) {
      showStatus('В таблице нет строк для копирования', 'warn');
      return false;
    }

    const template = rows[0];
    tableBody.innerHTML = '';

    for (let i = 0; i < count; i++) {
      const clone = template.cloneNode(true);
      const eventCell = qs('.event_type', clone);
      if (eventCell) eventCell.textContent = getRandomEventType();
      tableBody.appendChild(clone);
    }

    showStatus(`Создано строк: ${count}`, 'ok');
    return true;
  }

  function initLogTab() {
    qs('#bo-log-change-ip', app).addEventListener('click', (e) => {
      if (changeLogIP()) flashButton(e.currentTarget, '✅ Готово');
    });

    qs('#bo-log-change-date', app).addEventListener('click', (e) => {
      if (changeLogDate()) flashButton(e.currentTarget, '✅ Готово');
    });

    qs('#bo-log-duplicate', app).addEventListener('click', (e) => {
      if (duplicateLogRows()) flashButton(e.currentTarget, '✅ Готово');
    });

    qs('#bo-log-time-range', app).addEventListener('click', (e) => {
      const from = normalizeLogTimeInput('bo-log-time-from');
      const to = normalizeLogTimeInput('bo-log-time-to');
      const createdInput = qs('#bo-log-created-datetime', app);
      const created = createdInput && createdInput.value.trim()
        ? normalizeLogTimeInput('bo-log-created-datetime')
        : '';

      if (!from || !to) {
        showStatus('Введите время ОТ и ДО в формате HH:MM:SS', 'error');
        return;
      }

      if (createdInput && createdInput.value.trim() && !created) {
        showStatus('Created-время указано неверно', 'error');
        return;
      }

      if (applyIncreasingTimes(from, to, created)) flashButton(e.currentTarget, '✅ Готово');
    });
  }

  /************************************************************
   * 11. ЗАПУСК
   ************************************************************/
  function init() {
    restoreFormValues();

    [
      'bo-user-data',
      'bo-order-date',
      'bo-order-tx',
      'bo-order-items-data',
      'bo-log-ip',
      'bo-log-date',
      'bo-log-duplicate-count',
      'bo-log-time-from',
      'bo-log-time-to',
      'bo-log-created-datetime'
    ].forEach(bindPersistentField);

    initUserTab();
    initOrderTab();
    initLogTab();
    console.log('[WP Admin Tools] Панель запущена:', currentHost);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
