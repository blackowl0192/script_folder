// ==UserScript==
// @name         SteamBuy — Excel → Trash (автоочистка поиска, подвижная панель)
// @namespace    https://steambuy.pro/
// @version      1.3.1
// @description  Поиск только по ID из Excel; проверка названия перед Trash; цикл, отчёт, подвижная панель и обновление через GitHub.
// @homepageURL  https://github.com/blackowl0192/script_folder/blob/main/steambuy_excel_trash_v1_2.user.js
// @updateURL    https://raw.githubusercontent.com/blackowl0192/script_folder/main/steambuy_excel_trash_v1_2.user.js
// @downloadURL  https://raw.githubusercontent.com/blackowl0192/script_folder/main/steambuy_excel_trash_v1_2.user.js
// @match        https://steambuy.pro/wp-admin/*
// @noframes
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_getValue
// @grant        GM_setValue
// @run-at       document-idle
// ==/UserScript==

(() => {
  'use strict';

  // Ключ намеренно прежний: сохранённая очередь из v1.1 доступна после обновления.
  const KEY = 'steamBuyExcelTrash_v1';
  const UI_KEY = 'steamBuyExcelTrash_panel_v12';
  const LIST_URL = 'https://steambuy.pro/wp-admin/edit.php?post_type=product&post_status=all&lang=ru';
  const MAX_PAGES = 30;
  const VERIFY_PAGES = 12;
  const STEP_DELAY = 650;
  const OK = 'Удалено (проверено)';
  const WAIT = 'Ожидает';
  let imported = null;
  let state = GM_getValue(KEY, null);
  // Очереди старых версий тоже переключаются на поиск по ID.
  if (state && state.searchBy !== 'id') {
    state.searchBy = 'id';
    GM_setValue(KEY, state);
  }
  let scheduled = false;
  let busy = false;

  // WordPress может перенаправить пользователя после Trash на post.php,
  // edit.php без post_type, либо на другую страницу в wp-admin. Сохраняем
  // очередь в Tampermonkey и возвращаемся к списку товаров. В версии 1.2
  // Trash выполняется фоновым запросом, поэтому такие переходы редки.
  const params = new URLSearchParams(location.search);
  const isProductPage = location.pathname.endsWith('/edit.php') && params.get('post_type') === 'product';
  if (!isProductPage) {
    if (state?.running) {
      const timeKey = 'steamBuyExcelTrash_redirectTime';
      const oldTime = Number(sessionStorage.getItem(timeKey)) || 0;
      if (Date.now() - oldTime > 1500) {
        sessionStorage.setItem(timeKey, String(Date.now()));
        location.replace(LIST_URL);
      }
    }
    return;
  }

  // В поиск WordPress вводится исключительно ID из Excel. Название — только для проверки перед Trash.
  function queryFor(item) {
    return item.id;
  }

  const norm = value => String(value ?? '')
    .normalize('NFKC')
    .toLocaleLowerCase()
    .replace(/[\u2122\u00ae\u00a9]/g, '')
    .replace(/[\u2018\u2019`]/g, "'")
    .replace(/[\u2010-\u2015\u2212]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

  const clean = value => String(value ?? '').trim();
  const $ = (selector, parent = document) => parent.querySelector(selector);
  const $$ = (selector, parent = document) => Array.from(parent.querySelectorAll(selector));

  function save() {
    GM_setValue(KEY, state);
    renderStatus();
  }

  function delayStep() {
    if (scheduled || !state?.running) return;
    scheduled = true;
    setTimeout(() => {
      scheduled = false;
      runStep();
    }, STEP_DELAY);
  }

  function finish() {
    state.running = false;
    state.phase = 'idle';
    state.finishedAt = new Date().toLocaleString('ru-RU');
    save();
    setMessage('Обработка закончена. Скачайте Excel с результатами.');
  }

  function advance() {
    state.index += 1;
    state.phase = 'search';
    state.page = 0;
    state.candidates = [];
    state.previewNames = [];
    state.resultCount = 0;
    state.pendingPostId = '';
    state.searchAttempts = 0;
    save();
    if (state.index >= state.items.length) finish();
    else {
      const next = state.items[state.index];
      setMessage(`Следующая игра: ${next.name} (${next.id}). Запускаю новый поиск...`);
      delayStep();
    }
  }

  function mark(status, note = '') {
    const item = state.items[state.index];
    if (!item) return finish();
    item.status = status;
    item.note = note;
    advance();
  }

  function setMessage(message) {
    const el = $('#sbp-message');
    if (el) el.textContent = message;
  }

  function buildSearchUrl(query, status = 'all') {
    const url = new URL(LIST_URL);
    url.searchParams.set('s', String(query).trim());
    url.searchParams.set('post_status', status);
    url.searchParams.set('paged', '1');
    return url;
  }

  function searchFor(query, nextPhase) {
    const field = $('#post-search-input');
    const submit = $('#search-submit');
    if (!field || !submit) {
      state.running = false;
      save();
      setMessage('Не найдено поле поиска или кнопка Search products. Обработка остановлена.');
      return;
    }

    const form = field.form || $('#posts-filter');
    if (!form) {
      state.running = false;
      save();
      setMessage('Не найдена форма списка товаров. Обработка остановлена.');
      return;
    }

    const attempts = (state.searchAttempts || 0) + 1;
    state.searchAttempts = attempts;
    if (attempts > 3) {
      return mark('Ошибка поиска', 'Не удалось выполнить новый поиск после 3 попыток.');
    }

    for (const name of ['product_cat', 'product_type', 'stock_status']) {
      const control = form.elements.namedItem(name);
      if (control && 'value' in control) control.value = '';
    }
    const paged = form.elements.namedItem('paged');
    if (paged && 'value' in paged) paged.value = '1';
    const ensureFormField = (name, value) => {
      let element = form.elements.namedItem(name);
      if (!element) {
        element = document.createElement('input');
        element.type = 'hidden';
        element.name = name;
        form.appendChild(element);
      }
      if ('value' in element) element.value = value;
    };
    ensureFormField('lang', 'ru');
    ensureFormField('post_type', 'product');
    ensureFormField('post_status', 'all');
    form.action = LIST_URL;

    // ВАЖНО: отдельно и полностью удаляем ПРЕДЫДУЩИЙ текст, затем
    // отдельно вписываем НОВУЮ игру. Срабатывают события для WP/плагинов.
    field.focus();
    field.value = '';
    field.dispatchEvent(new Event('input', { bubbles: true }));
    field.dispatchEvent(new Event('change', { bubbles: true }));
    field.value = String(query).trim();
    field.dispatchEvent(new Event('input', { bubbles: true }));
    field.dispatchEvent(new Event('change', { bubbles: true }));

    state.phase = nextPhase;
    state.page = 0;
    save();
    setMessage(`Поиск ${state.index + 1}/${state.items.length}: ${field.value}`);

    // Нажимаем именно Search products. При AJAX-перехвате или сломанном
    // submit подстраховываемся принудительным переходом по точному URL.
    const queryUrl = buildSearchUrl(query).href;
    const job = state;
    const jobIndex = state.index;
    submit.click();
    setTimeout(() => {
      if (state !== job || !state?.running || state.index !== jobIndex || state.phase !== nextPhase) return;
      // Если скрипт ещё работает на этой же странице, submit не вызвал
      // обычную загрузку. Переходим по URL с НОВЫМ s=..., без старого s.
      location.assign(queryUrl);
    }, 1700);
  }

  function currentSearchIs(query) {
    const search = new URLSearchParams(location.search).get('s');
    return search !== null && search.trim() === String(query).trim();
  }

  function rowsOnPage() {
    return $$('#the-list > tr[id^="post-"]').filter(row => /^post-\d+$/.test(row.id));
  }

  function rowName(row) {
    return clean(($('a.row-title', row) || $('.column-name strong a', row) || $('.column-name a', row))?.textContent);
  }

  function nextPageUrl() {
    const link = $('a.next-page:not(.disabled)', $('.tablenav-pages') || document);
    if (!link || !link.href) return null;
    try {
      const url = new URL(link.href, location.href);
      if (url.origin !== location.origin || !url.pathname.endsWith('/edit.php')) return null;
      return url.href;
    } catch { return null; }
  }

  function goToNextPageIfNeeded() {
    const link = nextPageUrl();
    if (!link) return false;
    if (state.page >= MAX_PAGES) return false;
    save();
    location.assign(link);
    return true;
  }

  function validTrashLink(candidate) {
    try {
      const url = new URL(candidate.trashUrl);
      return url.origin === location.origin &&
        url.pathname === '/wp-admin/post.php' &&
        url.searchParams.get('action') === 'trash' &&
        url.searchParams.get('post') === candidate.postId &&
        Boolean(url.searchParams.get('_wpnonce'));
    } catch { return false; }
  }

  function handleSearchResults(item) {
    if (!currentSearchIs(queryFor(item))) return searchFor(queryFor(item), 'result');
    state.searchAttempts = 0; // Сброс после подтверждённой загрузки нового s=...
    const rows = rowsOnPage();
    state.page += 1;
    state.resultCount += rows.length;

    for (const row of rows) {
      const title = rowName(row);
      if (!title) continue;
      if (state.previewNames.length < 5) state.previewNames.push(title);
      if (norm(title) !== norm(item.name)) continue;
      const trash = $('span.trash a.trash-product, span.trash a.submitdelete, span.trash a', row);
      const postId = row.id.replace('post-', '');
      state.candidates.push({ postId, title, trashUrl: trash?.href || '' });
    }
    save();

    if (state.candidates.length > 1) {
      return mark('Несколько совпадений', 'Найдено несколько товаров с тем же названием; автоматическое удаление запрещено.');
    }
    if (state.page < MAX_PAGES && goToNextPageIfNeeded()) return;
    if (state.page >= MAX_PAGES && nextPageUrl()) {
      return mark('Проверить вручную', `Результатов больше ${MAX_PAGES} страниц. Удаление не выполнялось.`);
    }

    if (state.candidates.length === 0) {
      const note = state.resultCount
        ? `По ID ${item.id} найдено ${state.resultCount} строк, но ни у одной нет точного совпадения названия с Excel. Примеры: ${state.previewNames.join(' | ')}`
        : `По ID ${item.id} товары не найдены.`;
      return mark(state.resultCount ? 'Название не совпало' : 'Не найдено', note);
    }

    const candidate = state.candidates[0];
    if (!validTrashLink(candidate)) {
      return mark('Нет ссылки Trash', 'Название совпало, но безопасная ссылка на перемещение в корзину не найдена.');
    }

    // ТОЛЬКО точное совпадение. Не переходим на post.php: при обычном
    // переходе WordPress мог прерывать выполнение скрипта после 1-й игры.
    // Выполняем Trash с теми же cookies в фоне и остаёмся на странице поиска.
    item.foundName = candidate.title;
    item.wpPostId = candidate.postId;
    item.status = 'Отправлен в Trash, проверяем';
    state.pendingPostId = candidate.postId;
    state.phase = 'deleting';
    save();
    trashAndVerify(candidate, item);
  }

  async function verifyAbsent(query, postId) {
    let nextUrl = buildSearchUrl(query).href;
    for (let page = 0; page < VERIFY_PAGES; page++) {
      const response = await fetch(nextUrl, {
        credentials: 'same-origin', redirect: 'follow', cache: 'no-store'
      });
      if (!response.ok) throw new Error(`Проверка списка: HTTP ${response.status}`);
      if (new URL(response.url).pathname.endsWith('/wp-login.php')) {
        throw new Error('Сессия WordPress завершилась.');
      }
      const doc = new DOMParser().parseFromString(await response.text(), 'text/html');
      if (!doc.querySelector('#the-list')) {
        throw new Error('WordPress не вернул таблицу товаров при проверке.');
      }
      if (doc.getElementById(`post-${postId}`)) return false;
      const next = doc.querySelector('.tablenav-pages a.next-page:not(.disabled)');
      if (!next?.getAttribute('href')) return true;
      const url = new URL(next.getAttribute('href'), nextUrl);
      if (url.origin !== location.origin || !url.pathname.endsWith('/edit.php')) {
        throw new Error('Некорректная ссылка следующей страницы проверки.');
      }
      nextUrl = url.href;
    }
    return null; // Слишком много страниц, безопаснее проверить вручную.
  }

  async function trashAndVerify(candidate, item) {
    const job = state;
    const jobIndex = state.index;
    busy = true;
    try {
      setMessage(`Trash: ${item.name} (WP ${candidate.postId}). Проверяю результат...`);
      const response = await fetch(candidate.trashUrl, {
        method: 'GET', credentials: 'same-origin', redirect: 'follow', cache: 'no-store'
      });
      if (!response.ok) throw new Error(`Trash: HTTP ${response.status}`);
      const url = new URL(response.url);
      if (url.origin !== location.origin || url.pathname.endsWith('/wp-login.php')) {
        throw new Error('Не удалось подтвердить авторизацию в WordPress.');
      }
      // Прочитать тело, чтобы запрос WordPress полностью завершился.
      const html = await response.text();
      if (/id=["']loginform["']/i.test(html) || /class=["']wp-die-message["']/i.test(html)) {
        throw new Error('WordPress вернул страницу входа/ошибки.');
      }
      if (state !== job || state.index !== jobIndex) return;

      const stillPresent = await verifyAbsent(queryFor(item), candidate.postId);
      if (state !== job || state.index !== jobIndex) return;
      if (stillPresent === false) {
        mark('Удаление не подтверждено', 'Товар с тем же WP post ID всё ещё отображается в списке «Все».');
      } else if (stillPresent === null) {
        mark('Проверить вручную', 'Проверка после Trash превысила лимит страниц.');
      } else {
        mark(OK, 'После Trash товар не найден среди активных товаров.');
      }
    } catch (error) {
      if (state === job && state.index === jobIndex) {
        mark('Проверить вручную', `Не удалось однозначно подтвердить Trash: ${error.message}. Повторно не удалялся.`);
      }
    } finally {
      busy = false;
      // Сразу берём СЛЕДУЮЩУЮ игру. Поле будет очищено в searchFor().
      if (state?.running && state.phase === 'search') delayStep();
    }
  }

  async function verifyInterrupted(item) {
    // После обновления старого скрипта или случайного закрытия вкладки:
    // не нажимаем Trash повторно — сначала проверяем уже сделанное действие.
    if (busy) return;
    const job = state;
    const jobIndex = state.index;
    busy = true;
    try {
      const result = await verifyAbsent(queryFor(item), state.pendingPostId);
      if (state !== job || state.index !== jobIndex) return;
      if (result === true) mark(OK, 'Проверено после возобновления работы.');
      else if (result === false) mark('Удаление не подтверждено', 'Товар найден после перезагрузки вкладки.');
      else mark('Проверить вручную', 'Проверка прерванного удаления не завершена.');
    } catch (e) {
      if (state === job && state.index === jobIndex) {
        mark('Проверить вручную', `После перезагрузки не удалось проверить удаление: ${e.message}`);
      }
    } finally {
      busy = false;
      if (state?.running && state.phase === 'search') delayStep();
    }
  }

  function runStep() {
    if (!state?.running || busy) return;
    const item = state.items[state.index];
    if (!item) return finish();
    if (item.status === 'Пропущено' || item.status === 'Дубликат в Excel' || item.status === 'Конфликт ID в Excel') {
      return advance();
    }
    switch (state.phase) {
      case 'search': return searchFor(queryFor(item), 'result');
      case 'result': return handleSearchResults(item);
      case 'deleting': // Не повторять Trash после перезагрузки страницы.
      case 'verify-search': // Обработка незавершённого задания старой версии.
      case 'verify-result': return verifyInterrupted(item);
      default:
        state.running = false;
        save();
        setMessage('Неизвестный этап обработки; выполнение приостановлено.');
    }
  }

  // Панель управления в wp-admin.
  const panel = document.createElement('aside');
  panel.id = 'sbp-panel';
  panel.innerHTML = `
    <style>
      #sbp-panel { position:fixed; z-index:999999; top:65px; right:16px; width:390px;
        height:560px; min-width:280px; min-height:75px;
        max-width:calc(100vw - 12px); max-height:calc(100vh - 12px);
        overflow:auto; resize:both; background:#fff; color:#1d2327;
        border:1px solid #ccd0d4; border-radius:10px; box-shadow:0 8px 28px #0004;
        padding:0; font:13px/1.45 Arial,sans-serif; }
      #sbp-panel * { box-sizing:border-box; }
      #sbp-panel #sbp-handle { position:sticky; top:0; z-index:2; display:flex;
        justify-content:space-between; align-items:center; gap:8px; padding:11px 12px;
        background:#eef5fb; border-bottom:1px solid #ccd0d4;
        cursor:move; user-select:none; touch-action:none; }
      #sbp-panel #sbp-handle b { font-size:14px; }
      #sbp-panel #sbp-handle button { min-width:28px; padding:3px 6px; background:#fff; }
      #sbp-panel #sbp-body { padding:12px 14px 15px; }
      #sbp-panel.sbp-collapsed { height:46px !important; min-height:46px;
        resize:none; overflow:hidden; }
      #sbp-panel.sbp-collapsed #sbp-body { display:none; }
      #sbp-panel.sbp-dragging { opacity:.92; }
      #sbp-panel label { display:block; font-weight:600; margin:8px 0 3px; }
      #sbp-panel input[type=file], #sbp-panel input[type=number],
      #sbp-panel select { display:block; width:100%; max-width:none; min-height:32px; }
      #sbp-panel .sbp-cols { display:grid; grid-template-columns:1fr 1fr; gap:10px; }
      #sbp-panel button { border:1px solid #2271b1; color:#2271b1; background:white;
        cursor:pointer; padding:7px 9px; border-radius:5px; font-weight:600; }
      #sbp-panel button.primary { color:white; background:#2271b1; }
      #sbp-panel button.danger { color:#b32d2e; border-color:#b32d2e; }
      #sbp-panel button:disabled { opacity:.45; cursor:not-allowed; }
      #sbp-panel .sbp-actions { display:flex; flex-wrap:wrap; gap:7px; margin-top:12px; }
      #sbp-panel #sbp-preview { white-space:pre-wrap; overflow-wrap:anywhere; max-height:90px;
        overflow:auto; border:1px solid #ddd; background:#f6f7f7; border-radius:5px;
        padding:8px; font:12px/1.4 monospace; }
      #sbp-panel .sbp-subtle { color:#646970; font-size:12px; }
      #sbp-panel .sbp-github { margin:7px 0 12px; display:flex; flex-wrap:wrap; gap:6px 12px; align-items:center; }
      #sbp-panel .sbp-github a { color:#2271b1; text-decoration:underline; font-weight:600; }
      #sbp-panel .sbp-github a:hover { color:#135e96; }
      #sbp-panel #sbp-message { margin-top:9px; color:#135e96; overflow-wrap:anywhere; }
      #sbp-panel #sbp-status { margin-top:10px; white-space:pre-line; background:#f0f6fc; padding:9px; border-radius:5px; }
    </style>
    <div id="sbp-handle" title="Зажмите мышью и переместите панель">
      <b>⠿ SteamBuy · Excel → Trash</b>
      <span><button id="sbp-home" type="button" title="Вернуть панель в угол">↺</button>
      <button id="sbp-collapse" type="button" title="Свернуть / развернуть">−</button></span>
    </div>
    <div id="sbp-body">
    <div class="sbp-subtle">v1.3.1 · Поиск только по ID из Excel. Перетащите панель за заголовок.</div>
    <div class="sbp-github">
      <span class="sbp-subtle">Автообновления через Tampermonkey</span>
      <a href="https://github.com/blackowl0192/script_folder/edit/main/steambuy_excel_trash_v1_2.user.js" target="_blank" rel="noopener noreferrer" title="Открыть редактор файла в GitHub">✎ Редактировать в GitHub</a>
      <a href="https://github.com/blackowl0192/script_folder/blob/main/steambuy_excel_trash_v1_2.user.js" target="_blank" rel="noopener noreferrer">Посмотреть код</a>
    </div>
    <div class="sbp-subtle">Удаление только при точном совпадении названия. Товары попадают в корзину.</div>
    <label for="sbp-file">Excel-файл (.xlsx, .xls, .csv)</label>
    <input id="sbp-file" type="file" accept=".xlsx,.xls,.csv" />
    <label for="sbp-sheet">Лист</label><select id="sbp-sheet"></select>
    <div class="sbp-cols">
      <div><label for="sbp-name-col">Колонка названия</label><select id="sbp-name-col"></select></div>
      <div><label for="sbp-id-col">Колонка ID</label><select id="sbp-id-col"></select></div>
    </div>
    <div class="sbp-subtle">В поле Search products вводится <b>ID из Excel</b>. Название игры используется только для сверки перед переносом в Trash.</div>
    <label for="sbp-start-row">Первая строка с игрой (1 или 2)</label>
    <input id="sbp-start-row" type="number" min="1" value="2" />
    <label>Предпросмотр</label><div id="sbp-preview">Выберите файл для загрузки.</div>
    <div class="sbp-actions">
      <button id="sbp-start" class="primary" disabled>Начать перенос</button>
      <button id="sbp-toggle" disabled>Пауза</button>
      <button id="sbp-export" disabled>Скачать Excel</button>
      <button id="sbp-reset" class="danger">Сбросить</button>
    </div>
    <div id="sbp-status"></div><div id="sbp-message"></div>
    <div class="sbp-subtle" style="margin-top:9px">Автоматически: очистка поля → следующий ID → Search products → сверка названия → Trash. При перезагрузке прогресс сохраняется.</div>
    <div class="sbp-subtle" style="margin-top:7px">При редактировании кода в GitHub увеличивайте @version (например, до 1.3.2), чтобы Tampermonkey обнаружил обновление.</div>
    </div>
  `;
  document.body.appendChild(panel);

  // Перетаскивание мышью + изменяемый размер + сохранение позиции.
  const drag = $('#sbp-handle');
  const collapsedButton = $('#sbp-collapse');
  let layout = GM_getValue(UI_KEY, {}) || {};
  function persistLayout() {
    if (panel.classList.contains('sbp-collapsed')) return;
    const r = panel.getBoundingClientRect();
    layout = { ...layout, left: r.left, top: r.top, width: r.width, height: r.height, collapsed: false };
    GM_setValue(UI_KEY, layout);
  }
  function applyLayout() {
    if (Number.isFinite(layout.width)) panel.style.width = `${layout.width}px`;
    if (Number.isFinite(layout.height)) panel.style.height = `${layout.height}px`;
    if (Number.isFinite(layout.left)) {
      panel.style.left = `${Math.max(0, Math.min(layout.left, innerWidth - panel.getBoundingClientRect().width - 6))}px`;
      panel.style.right = 'auto';
    }
    if (Number.isFinite(layout.top)) panel.style.top = `${Math.max(0, Math.min(layout.top, innerHeight - panel.getBoundingClientRect().height - 6))}px`;
    if (layout.collapsed) {
      panel.classList.add('sbp-collapsed');
      collapsedButton.textContent = '+';
    }
  }
  applyLayout();
  let dragging = null;
  drag.addEventListener('pointerdown', event => {
    if (event.target.closest('button') || event.button !== 0) return;
    const r = panel.getBoundingClientRect();
    dragging = { x: event.clientX, y: event.clientY, left: r.left, top: r.top };
    panel.style.left = `${r.left}px`;
    panel.style.right = 'auto';
    panel.classList.add('sbp-dragging');
    drag.setPointerCapture(event.pointerId);
    event.preventDefault();
  });
  drag.addEventListener('pointermove', event => {
    if (!dragging) return;
    const left = Math.max(0, Math.min(innerWidth - panel.getBoundingClientRect().width - 6, dragging.left + event.clientX - dragging.x));
    const top = Math.max(0, Math.min(innerHeight - panel.getBoundingClientRect().height - 6, dragging.top + event.clientY - dragging.y));
    panel.style.left = `${left}px`;
    panel.style.top = `${top}px`;
  });
  function endDrag() {
    if (!dragging) return;
    dragging = null;
    panel.classList.remove('sbp-dragging');
    const r = panel.getBoundingClientRect();
    layout = { ...layout, left: r.left, top: r.top };
    GM_setValue(UI_KEY, layout);
  }
  drag.addEventListener('pointerup', endDrag);
  drag.addEventListener('pointercancel', endDrag);
  collapsedButton.addEventListener('click', () => {
    const nowCollapsed = !panel.classList.contains('sbp-collapsed');
    if (nowCollapsed) persistLayout();
    panel.classList.toggle('sbp-collapsed', nowCollapsed);
    collapsedButton.textContent = nowCollapsed ? '+' : '−';
    layout.collapsed = nowCollapsed;
    if (!nowCollapsed && Number.isFinite(layout.height)) panel.style.height = `${layout.height}px`;
    GM_setValue(UI_KEY, layout);
  });
  $('#sbp-home').addEventListener('click', () => {
    layout = { ...layout, left: innerWidth - Math.min(390, innerWidth - 12) - 16, top: 65, width: 390, height: 560 };
    applyLayout();
    GM_setValue(UI_KEY, layout);
  });
  if (typeof ResizeObserver !== 'undefined') {
    let resizeTimer = null;
    new ResizeObserver(() => {
      if (panel.classList.contains('sbp-collapsed')) return;
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(persistLayout, 500);
    }).observe(panel);
  }

  const fileInput = $('#sbp-file');
  const sheetSelect = $('#sbp-sheet');
  const nameSelect = $('#sbp-name-col');
  const idSelect = $('#sbp-id-col');
  const startRowInput = $('#sbp-start-row');
  const preview = $('#sbp-preview');

  function selectedRows() {
    return imported?.sheets?.[sheetSelect.value] || [];
  }

  function colLabel(index) {
    let n = index + 1;
    let s = '';
    while (n > 0) { n -= 1; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
    return s;
  }

  function refreshColumns() {
    const rows = selectedRows();
    const width = Math.max(2, Math.min(70, ...rows.slice(0, 100).map(row => row.length)));
    const top = rows[0] || [];
    nameSelect.replaceChildren();
    idSelect.replaceChildren();
    for (let i = 0; i < width; i++) {
      const label = `${colLabel(i)} · ${clean(top[i]).slice(0, 28) || '—'}`;
      for (const parent of [nameSelect, idSelect]) {
        const opt = document.createElement('option');
        opt.value = String(i);
        opt.textContent = label;
        parent.appendChild(opt);
      }
    }
    const header = top.map(cell => norm(cell));
    const maybeName = header.findIndex(x => /^(названи[ея]|игра|game|product name|title|name)$/.test(x));
    const maybeId = header.findIndex(x => /^(id|id товара|product id|артикул|sku)$/.test(x));
    nameSelect.value = String(maybeName >= 0 ? maybeName : 0);
    idSelect.value = String(maybeId >= 0 ? maybeId : 1);
    startRowInput.value = (maybeName >= 0 || maybeId >= 0) ? '2' : '1';
    updatePreview();
  }

  function updatePreview() {
    const rows = selectedRows();
    if (!imported) return;
    const start = Math.max(1, Number(startRowInput.value) || 1);
    const nameCol = Number(nameSelect.value);
    const idCol = Number(idSelect.value);
    const sample = rows.slice(start - 1).filter(row => clean(row[nameCol]) || clean(row[idCol])).slice(0, 4);
    preview.textContent = sample.length
      ? sample.map(row => `${clean(row[nameCol])}  →  ${clean(row[idCol])}`).join('\n')
      : 'Нет заполненных строк. Проверьте номера колонок.';
    $('#sbp-start').disabled = !sample.length || Boolean(state?.running);
  }

  fileInput.addEventListener('change', async () => {
    if (state?.running) return setMessage('Сначала остановите текущую обработку.');
    const file = fileInput.files?.[0];
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array' });
      imported = { fileName: file.name, sheets: {} };
      for (const name of wb.SheetNames) {
        imported.sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], {
          header: 1, raw: false, defval: '', blankrows: true
        });
      }
      sheetSelect.replaceChildren();
      for (const name of wb.SheetNames) {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        sheetSelect.appendChild(option);
      }
      refreshColumns();
      setMessage(`Загружен файл ${file.name}. Проверьте соответствие колонок.`);
    } catch (error) {
      imported = null;
      $('#sbp-start').disabled = true;
      setMessage(`Не удалось открыть Excel: ${error.message}`);
    }
  });
  sheetSelect.addEventListener('change', refreshColumns);
  [nameSelect, idSelect, startRowInput].forEach(el => el.addEventListener('change', updatePreview));

  $('#sbp-start').addEventListener('click', () => {
    if (!imported) return;
    const rows = selectedRows();
    const startRow = Math.max(1, Math.floor(Number(startRowInput.value) || 1));
    const nameCol = Number(nameSelect.value);
    const idCol = Number(idSelect.value);
    if (nameCol === idCol) return setMessage('Название и ID должны быть в разных колонках.');

    const items = [];
    rows.forEach((row, index) => {
      if (index < startRow - 1) return;
      const name = clean(row[nameCol]);
      const id = clean(row[idCol]);
      if (!name && !id) return;
      items.push({ sourceRow: index + 1, name, id, status: !name || !id ? 'Пропущено' : WAIT,
        note: !name || !id ? 'Пустой ID или название.' : '', foundName: '', wpPostId: '' });
    });
    if (!items.length) return setMessage('Не найдено ни одной игры для обработки.');

    const byId = new Map();
    for (const item of items) {
      if (!item.id || !item.name) continue;
      if (!byId.has(item.id)) byId.set(item.id, []);
      byId.get(item.id).push(item);
    }
    for (const group of byId.values()) {
      const distinctNames = new Set(group.map(item => norm(item.name)));
      if (distinctNames.size > 1) {
        group.forEach(item => { item.status = 'Конфликт ID в Excel'; item.note = 'Один ID прикреплён к разным названиям.'; });
      } else {
        group.slice(1).forEach(item => { item.status = 'Дубликат в Excel'; item.note = 'Повтор того же ID и названия в файле.'; });
      }
    }
    const ready = items.filter(item => item.status === WAIT).length;
    if (!ready) return setMessage('Нет записей для поиска: проверьте ID и названия.');
    const searchDescription = 'по ID из Excel';
    if (!confirm(`Файл: ${imported.fileName}\nИгр в таблице: ${items.length}\nК поиску и переносу в Trash: ${ready}\nСпособ поиска: ${searchDescription}\n\nПосле каждого Trash скрипт автоматически проверит результат и перейдёт к следующей игре. Удаляется только товар с точно совпавшим названием.\n\nПродолжить?`)) return;

    state = {
      fileName: imported.fileName, sheetName: sheetSelect.value,
      originalRows: rows, startRow, nameCol, idCol, searchBy: 'id',
      items, index: 0, running: true, phase: 'search',
      page: 0, candidates: [], previewNames: [], resultCount: 0, pendingPostId: '',
      searchAttempts: 0,
      startedAt: new Date().toLocaleString('ru-RU')
    };
    save();
    setMessage('Обработка запущена. Не закрывайте вкладку.');
    delayStep();
  });

  $('#sbp-toggle').addEventListener('click', () => {
    if (!state?.items?.length) return;
    state.running = !state.running;
    save();
    setMessage(state.running ? 'Продолжаем обработку.' : 'Пауза. Можно скачать текущий Excel-отчёт.');
    if (state.running) delayStep();
  });

  $('#sbp-reset').addEventListener('click', () => {
    if (busy) return setMessage('Сначала дождитесь ответа на текущий запрос Trash.');
    if (!confirm('Удалить сохранённый прогресс обработки? Товары, ранее отправленные в Trash, не восстановятся.')) return;
    state = null;
    GM_setValue(KEY, null);
    renderStatus();
    updatePreview();
    setMessage('Прогресс сброшен. Можно начать новый файл.');
  });

  $('#sbp-export').addEventListener('click', () => {
    if (!state?.items?.length) return;
    try {
      const original = state.originalRows.map(row => [...row]);
      const shift = state.startRow === 1 ? 1 : 0;
      if (shift) original.unshift([]);
      const maxCols = Math.max(2, ...original.map(row => row.length));
      const headers = ['Статус обработки', 'Название на сайте', 'WP post ID', 'Комментарий'];
      if (!original[0]) original[0] = [];
      for (let j = 0; j < headers.length; j++) original[0][maxCols + j] = headers[j];

      for (const item of state.items) {
        const rowIndex = item.sourceRow - 1 + shift;
        const row = original[rowIndex] || (original[rowIndex] = []);
        row[maxCols] = item.status;
        row[maxCols + 1] = item.foundName;
        row[maxCols + 2] = item.wpPostId;
        row[maxCols + 3] = item.note;
      }
      const wb = XLSX.utils.book_new();
      const output = XLSX.utils.aoa_to_sheet(original);
      output['!cols'] = [
        ...Array.from({ length: maxCols }, (_, i) => ({ wch: i === state.nameCol ? 54 : i === state.idCol ? 20 : 20 })),
        { wch: 28 }, { wch: 54 }, { wch: 17 }, { wch: 90 }
      ];
      output['!autofilter'] = { ref: `A1:${colLabel(maxCols + 3)}${Math.max(2, original.length)}` };
      XLSX.utils.book_append_sheet(wb, output, 'Список и статусы');
      const problems = state.items.filter(item => item.status !== OK).map(item => [
        item.sourceRow, item.id, item.name, item.status, item.foundName, item.wpPostId, item.note
      ]);
      const missed = XLSX.utils.aoa_to_sheet([
        ['Строка Excel', 'ID', 'Название игры', 'Статус', 'Найдено на сайте', 'WP post ID', 'Комментарий'],
        ...problems
      ]);
      missed['!cols'] = [{ wch: 14 }, { wch: 20 }, { wch: 50 }, { wch: 27 }, { wch: 50 }, { wch: 16 }, { wch: 90 }];
      XLSX.utils.book_append_sheet(wb, missed, 'Проблемные ID');
      const base = state.fileName.replace(/\.(xlsx|xls|csv)$/i, '').replace(/[\\/:*?"<>|]/g, '_');
      XLSX.writeFile(wb, `${base}_результат_Trash.xlsx`);
      setMessage('Excel-отчёт сохранён.');
    } catch (error) {
      setMessage(`Ошибка сохранения Excel: ${error.message}`);
    }
  });

  function renderStatus() {
    const button = $('#sbp-toggle');
    const exportButton = $('#sbp-export');
    const startButton = $('#sbp-start');
    const status = $('#sbp-status');
    if (!status) return;
    const hasData = Boolean(state?.items?.length);
    button.disabled = !hasData || (!state.running && state.index >= state.items.length);
    button.textContent = state?.running ? 'Пауза' : 'Продолжить';
    exportButton.disabled = !hasData;
    startButton.disabled = !imported || Boolean(state?.running);
    fileInput.disabled = Boolean(state?.running);
    if (!hasData) { status.textContent = 'Нет активной задачи.'; return; }
    const counts = {};
    state.items.forEach(item => { counts[item.status] = (counts[item.status] || 0) + 1; });
    const processed = state.items.filter(item => item.status !== WAIT && item.status !== 'Отправлен в Trash, проверяем').length;
    const current = state.items[state.index];
    status.textContent = [
      `${state.running ? '▶ Обработка' : 'Ⅱ Пауза / готово'} · ${processed}/${state.items.length}`,
      `Поиск: только по ID из Excel`,
      `Текущая игра: ${current ? `${current.name} [${current.id}]` : '—'}`,
      ...Object.entries(counts).map(([key, count]) => `${key}: ${count}`)
    ].join('\n');
  }

  renderStatus();
  if (state?.running) {
    const item = state.items?.[state.index];
    if (item) setMessage(`Автопродолжение: ${item.name} (${item.id}). Этап: ${state.phase}.`);
    delayStep();
  }
})();
