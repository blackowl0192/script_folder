// ==UserScript==
// @name         SteamBuy — Product Trash Manager
// @namespace    https://steambuy.pro/
// @version      1.4.0
// @description  Безопасное удаление товара по WordPress ID, тест одного товара, Excel-очередь, отчёт, автообновление из GitHub.
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
    setMessage('Задача завершена. Результат доступен в отчёте XLSX.');
  }

  function advance() {
    state.index += 1;
    state.phase = 'lookup';
    state.page = 0;
    state.candidates = [];
    state.previewNames = [];
    state.resultCount = 0;
    state.pendingPostId = '';
    state.searchAttempts = 0;
    save();
    if (state.mode === 'single' || state.index >= state.items.length) finish();
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

  function setMessage(message){
    const el=$('#sbp-message');if(el) el.textContent=message;
    if(state){
      if(!Array.isArray(state.logs))state.logs=[];
      if(state.logs.at(-1)?.message!==message){
        state.logs.push({time:new Date().toLocaleTimeString('ru-RU'),message});
        if(state.logs.length>35) state.logs.splice(0,state.logs.length-35);
        GM_setValue(KEY,state);
      }
    }
    const log=$('#sbp-log');
    if(log)log.textContent=(state?.logs||[]).map(x=>`${x.time}  ${x.message}`).join('\n');
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

  async function lookupById(item) {
    if (busy) return;
    const job=state, jobIndex=state.index, id=clean(item.id);
    if (!/^[1-9]\d*$/.test(id)) return mark('Некорректный ID','Нужен положительный числовой WordPress product ID, не SKU.');
    busy=true;
    try {
      setMessage(`Проверка карточки товара по ID ${id}...`);
      const url=new URL('/wp-admin/post.php',location.origin);
      url.searchParams.set('post',id);
      url.searchParams.set('action','edit');
      const response=await fetch(url.href,{credentials:'same-origin',redirect:'follow',cache:'no-store'});
      if(!response.ok) throw new Error(`HTTP ${response.status}`);
      if(new URL(response.url).pathname.endsWith('/wp-login.php')) throw new Error('Сессия WordPress истекла.');
      const doc=new DOMParser().parseFromString(await response.text(),'text/html');
      if(doc.querySelector('#loginform,.wp-die-message')) throw new Error('Нет доступа к этому ID или WordPress вернул ошибку.');
      const form=doc.querySelector('form#post');
      const actualId=clean(doc.querySelector('#post_ID,input[name="post_ID"]')?.value);
      const type=clean(doc.querySelector('#post_type,input[name="post_type"]')?.value);
      const name=clean(doc.querySelector('#title,input[name="post_title"],textarea[name="post_title"]')?.value);
      const status=clean(doc.querySelector('#post_status,input[name="post_status"],#original_post_status')?.value);
      if(state!==job||state.index!==jobIndex||!state.running) return;
      if(!form||actualId!==id) return mark('ID не подтверждён',`Карточка с WordPress ID ${id} не открылась или ID отличается.`);
      if(type!=='product') return mark('Не товар',`Для ID ${id} обнаружен тип ${type||'неизвестен'}. Удаление запрещено.`);
      if(status==='trash') return mark('Уже в корзине',`Товар ID ${id} уже в Trash.`);
      if(!name) return mark('Название недоступно',`Карточка ID ${id} не предоставила название для безопасной проверки.`);
      if(item.name && norm(name)!==norm(item.name)) return mark('Название не совпало',`По ID ${id} сайт возвращает «${name}», а в Excel «${item.name}».`);
      const link=doc.querySelector('#delete-action a.submitdelete,#delete-action a');
      const candidate={postId:id,title:name,trashUrl:link?.href||''};
      if(!validTrashLink(candidate)) return mark('Нет ссылки Trash',`Карточка товара ${id} не содержит действительную ссылку удаления с nonce.`);
      if(state.mode==='single'&&!confirm(`ПРОВЕРЬТЕ ТОВАР:\nWordPress ID: ${id}\nНазвание: ${name}\n\nПереместить ТОЛЬКО ЭТОТ ОДИН товар в корзину?`))
        return mark('Отменено пользователем','Тест остановлен до удаления.');
      if(!state.running) return;
      item.foundName=name;
      item.name=item.name||name;
      item.wpPostId=id;
      item.status='Отправлен в Trash, проверяем';
      state.pendingPostId=id;
      state.phase='deleting';
      save();
      busy=false;
      return await trashAndVerify(candidate,item);
    } catch(e) {
      if(state===job&&state.index===jobIndex) mark('Ошибка проверки ID',String(e.message||e));
    } finally {
      busy=false;
      if(state?.running&&state.phase==='lookup') delayStep();
    }
  }

  async function verifyTrashStatus(postId) {
    const url=new URL('/wp-admin/post.php',location.origin);
    url.searchParams.set('post',String(postId));
    url.searchParams.set('action','edit');
    const response=await fetch(url.href,{credentials:'same-origin',redirect:'follow',cache:'no-store'});
    if(!response.ok) return null;
    if(new URL(response.url).pathname.endsWith('/wp-login.php')) throw new Error('Сессия WordPress истекла.');
    const doc=new DOMParser().parseFromString(await response.text(),'text/html');
    const actualId=clean(doc.querySelector('#post_ID,input[name="post_ID"]')?.value);
    if(actualId!==String(postId)) return null;
    const status=clean(doc.querySelector('#post_status,input[name="post_status"],#original_post_status,input[name="original_post_status"]')?.value);
    if(status==='trash') return true;
    if(status) return false;
    return null;
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
      const postId = row.id.replace('post-', '');
      if (postId !== clean(item.id)) continue; // Даже в старой фазе поиска нужен точный WordPress ID.
      const trash = $('span.trash a.trash-product, span.trash a.submitdelete, span.trash a', row);
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

      const trashed = await verifyTrashStatus(candidate.postId);
      if (state !== job || state.index !== jobIndex) return;
      if (trashed === true) mark(OK, 'Карточка WordPress подтвердила статус trash.');
      else if (trashed === false) mark('Удаление не подтверждено', 'После запроса статус товара не стал trash.');
      else mark('Проверить вручную', 'Не удалось подтвердить статус trash по карточке товара. Повторное удаление не выполнялось.');
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
      const result = await verifyTrashStatus(state.pendingPostId);
      if (state !== job || state.index !== jobIndex) return;
      if (result === true) mark(OK, 'Статус trash проверен после возобновления.');
      else if (result === false) mark('Удаление не подтверждено', 'После перезагрузки статус товара не trash.');
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
      case 'search': // Совместимость с очередью v1.3.1.
      case 'lookup': return lookupById(item);
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
      #sbp-panel {position:fixed;z-index:999999;right:16px;top:65px;width:438px;height:680px;min-width:330px;min-height:52px;max-width:calc(100vw - 12px);max-height:calc(100vh - 12px);resize:both;overflow:auto;background:#f8fafc;color:#152235;border:1px solid #ced8e5;border-radius:15px;box-shadow:0 18px 54px rgba(22,36,58,.27);font:13px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif}
      #sbp-panel * {box-sizing:border-box}
      #sbp-panel #sbp-handle {position:sticky;top:0;z-index:2;display:flex;justify-content:space-between;align-items:center;gap:10px;padding:12px 14px;background:linear-gradient(120deg,#12243d,#22476f);color:#fff;cursor:move;user-select:none;touch-action:none}
      #sbp-panel #sbp-handle b {font-size:15px;letter-spacing:.2px}
      #sbp-panel #sbp-handle small {font-size:11px;opacity:.72;font-weight:500}
      #sbp-panel #sbp-handle button {min-width:31px;padding:5px;background:#ffffff18;color:#fff;border-color:#ffffff55}
      #sbp-panel #sbp-body {padding:13px 14px 18px}
      #sbp-panel.sbp-collapsed {height:50px!important;min-height:50px;resize:none;overflow:hidden}
      #sbp-panel.sbp-collapsed #sbp-body {display:none}
      #sbp-panel.sbp-dragging {opacity:.93}
      #sbp-panel .sbp-card {background:#fff;border:1px solid #dce5f0;border-radius:11px;margin:0 0 11px;padding:12px;box-shadow:0 2px 6px #15223508}
      #sbp-panel .sbp-card h3 {margin:0 0 7px;font-size:13px;font-weight:750;color:#142d48}
      #sbp-panel .sbp-card p {margin:5px 0 8px}
      #sbp-panel .sbp-subtle {color:#697889;font-size:11.5px}
      #sbp-panel label {display:block;font-weight:600;margin:9px 0 4px;color:#344760}
      #sbp-panel input[type=file],#sbp-panel input[type=number],#sbp-panel input[type=text],#sbp-panel select {display:block;width:100%;max-width:none;min-height:35px;padding:6px 9px;border:1px solid #cad7e6;border-radius:7px;background:#fff;color:#16283e;box-shadow:none}
      #sbp-panel input:focus-visible,#sbp-panel select:focus-visible,#sbp-panel button:focus-visible {outline:2px solid #2684c7;outline-offset:2px}
      #sbp-panel .sbp-cols {display:grid;grid-template-columns:1fr 1fr;gap:9px}
      #sbp-panel button {border:1px solid #b7cce0;background:#fff;color:#255379;cursor:pointer;padding:8px 10px;border-radius:7px;font-weight:650;transition:background .12s,border-color .12s}
      #sbp-panel button:hover:not(:disabled) {background:#eaf3fd;border-color:#588bb6}
      #sbp-panel button.primary {background:#176db1;color:#fff;border-color:#176db1}
      #sbp-panel button.primary:hover:not(:disabled) {background:#10568e}
      #sbp-panel button.danger {color:#bd3339;border-color:#e5b7bd}
      #sbp-panel button.danger:hover:not(:disabled) {background:#fff1f1}
      #sbp-panel button:disabled {opacity:.48;cursor:not-allowed}
      #sbp-panel .sbp-actions {display:flex;flex-wrap:wrap;gap:7px;margin-top:10px}
      #sbp-panel .sbp-actions button {flex:1 1 auto}
      #sbp-panel #sbp-preview {white-space:pre-wrap;overflow-wrap:anywhere;max-height:110px;overflow:auto;background:#f3f7fc;border:1px solid #e0e7f0;border-radius:7px;padding:9px;font:12px/1.6 Consolas,monospace}
      #sbp-panel .sbp-gh {display:flex;justify-content:space-between;gap:6px;align-items:center;margin:0 0 10px;padding:0 2px}
      #sbp-panel .sbp-gh a {font-weight:600;color:#176db1;text-decoration:none}
      #sbp-panel .sbp-gh a:hover {text-decoration:underline}
      #sbp-panel .sbp-badge {display:inline-block;background:#d8f2e6;color:#176547;padding:3px 7px;border-radius:7px;font-size:11px;font-weight:750}
      #sbp-panel #sbp-progress {height:8px;overflow:hidden;background:#e7eef8;border-radius:50px;margin:10px 0}
      #sbp-panel #sbp-progress-fill {height:100%;width:0;background:linear-gradient(90deg,#1979b5,#22b6a0);transition:width .25s}
      #sbp-panel #sbp-stats {font-weight:700;color:#1f3c5d}
      #sbp-panel #sbp-status {white-space:pre-line;overflow-wrap:anywhere;border-top:1px solid #e2e9f2;padding-top:8px;color:#445a72;font-size:12px}
      #sbp-panel #sbp-message {margin-top:8px;color:#134b74;font-weight:650;overflow-wrap:anywhere}
      #sbp-panel #sbp-log {white-space:pre-wrap;overflow-wrap:anywhere;max-height:115px;overflow:auto;color:#627386;font:11px/1.55 Consolas,monospace}
      #sbp-panel details summary {cursor:pointer;font-weight:650;color:#355878}
      @media(max-width:600px){#sbp-panel {width:min(438px,calc(100vw - 12px));right:6px;top:50px}#sbp-panel .sbp-cols {grid-template-columns:1fr}}
    </style>
    <div id="sbp-handle" title="Перемещение панели мышью">
      <div><b>SteamBuy · Product Trash</b><div><small>v1.4.0 · поиск по WordPress ID</small></div></div>
      <span><button id="sbp-home" type="button" title="Сбросить расположение">↺</button> <button id="sbp-collapse" type="button" title="Свернуть/развернуть">−</button></span>
    </div>
    <div id="sbp-body">
      <div class="sbp-gh"><span class="sbp-badge">GitHub · автообновления</span><a href="https://github.com/blackowl0192/script_folder/blob/main/steambuy_excel_trash_v1_2.user.js" target="_blank" rel="noopener noreferrer">Исходный код</a></div>
      <section class="sbp-card">
        <h3>01 · Тест одного товара</h3>
        <p class="sbp-subtle">Проверка по внутреннему ID WooCommerce, показ найденного названия и дополнительное подтверждение. Excel не требуется.</p>
        <div class="sbp-cols"><div><label for="sbp-single-id">ID товара *</label><input id="sbp-single-id" type="text" inputmode="numeric" placeholder="Например, 142" autocomplete="off"></div><div><label for="sbp-single-name">Название (необязательно)</label><input id="sbp-single-name" type="text" placeholder="Сверить с сайтом" autocomplete="off"></div></div>
        <div class="sbp-actions"><button id="sbp-test-manual" type="button" class="primary">Проверить и удалить 1 товар</button></div>
      </section>
      <section class="sbp-card">
        <h3>02 · Массовая обработка Excel</h3>
        <p class="sbp-subtle">ID из файла трактуется как WordPress post ID товара, не как SKU. При несовпадении ID, типа или названия удаление не происходит.</p>
        <label for="sbp-file">Excel / CSV</label><input id="sbp-file" type="file" accept=".xlsx,.xls,.csv">
        <label for="sbp-sheet">Лист</label><select id="sbp-sheet"></select>
        <div class="sbp-cols"><div><label for="sbp-id-col">Колонка ID</label><select id="sbp-id-col"></select></div><div><label for="sbp-name-col">Название игры</label><select id="sbp-name-col"></select></div></div>
        <label for="sbp-start-row">Начальная строка (учитывая заголовок)</label><input id="sbp-start-row" type="number" min="1" value="2">
        <label>Первые строки</label><div id="sbp-preview">Выберите файл.</div>
        <div class="sbp-actions"><button id="sbp-test-first" type="button" disabled>Тест: 1 из Excel</button><button id="sbp-start" type="button" class="primary" disabled>Обработать весь Excel</button></div>
      </section>
      <section class="sbp-card">
        <h3>03 · Управление и отчёт</h3>
        <div class="sbp-actions"><button id="sbp-toggle" type="button" disabled>Пауза</button><button id="sbp-export" type="button" disabled>Скачать отчёт XLSX</button><button id="sbp-reset" type="button" class="danger">Сбросить</button></div>
        <div id="sbp-stats">Нет задачи</div><div id="sbp-progress"><div id="sbp-progress-fill"></div></div>
        <div id="sbp-status">Загрузите Excel или проверьте один ID.</div>
        <div id="sbp-message" role="status" aria-live="polite"></div>
      </section>
      <section class="sbp-card"><details><summary>Последние действия и безопасность</summary><p class="sbp-subtle">Отправка в корзину, не безвозвратное удаление. После сомнительного ответа повторный запрос Trash автоматически не отправляется. Прогресс хранится в Tampermonkey.</p><div id="sbp-log"></div></details></section>
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
    layout = { ...layout, left: Math.max(0, innerWidth - Math.min(438, innerWidth - 12) - 16), top: 65, width: 438, height: 680, collapsed: false };
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
    const maybeId = header.findIndex(x => /^(id|id товара|product id|wp id|wordpress id)$/.test(x));
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
    $('#sbp-test-first').disabled = !sample.length || Boolean(state?.running);
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

  function startImportedQueue(single) {
    if(!imported||busy||state?.running) return;
    const rows=selectedRows(),startRow=Math.max(1,Math.floor(Number(startRowInput.value)||1));
    const nameCol=Number(nameSelect.value),idCol=Number(idSelect.value);
    if(nameCol===idCol) return setMessage('Колонки ID и названия должны различаться.');
    const items=[];
    rows.forEach((row,index)=>{
      if(index<startRow-1) return;
      const name=clean(row[nameCol]),id=clean(row[idCol]);
      if(!name&&!id) return;
      const invalid=!/^[1-9]\d*$/.test(id),missing=!name;
      items.push({sourceRow:index+1,name,id,status:missing?'Пропущено':invalid?'Некорректный ID':WAIT,
        note:missing?'Не указано название.':invalid?'ID должен быть числовым WordPress post ID.':'',foundName:'',wpPostId:''});
    });
    if(!items.length) return setMessage('Не найдено строк с товарами.');
    const byId=new Map();
    for(const item of items){
      if(item.status!==WAIT) continue;
      if(!byId.has(item.id)) byId.set(item.id,[]);
      byId.get(item.id).push(item);
    }
    for(const group of byId.values()){
      const names=new Set(group.map(x=>norm(x.name)));
      if(names.size>1) group.forEach(x=>{x.status='Конфликт ID в Excel';x.note='Один и тот же ID сопоставлен нескольким названиям.';});
      else group.slice(1).forEach(x=>{x.status='Дубликат в Excel';x.note='Повтор ID и названия.';});
    }
    const ready=items.filter(x=>x.status===WAIT);
    if(!ready.length) return setMessage('Нет корректных строк с числовым ID и названием.');
    if(state?.items?.length&&state.index<state.items.length&&!confirm('Предыдущий прогресс будет заменён. Продолжить?')) return;
    if(!confirm(single?
      `ТЕСТ: только 1 товар из Excel.\nID: ${ready[0].id}\nНазвание: ${ready[0].name}\n\nПеред удалением будет дополнительная проверка и подтверждение. Продолжить?`:
      `Файл: ${imported.fileName}\nВсего строк: ${items.length}\nК обработке: ${ready.length}\n\nТовары перемещаются в корзину по WordPress ID при точном совпадении названия. Начать?`)) return;
    state={
      fileName:imported.fileName,sheetName:sheetSelect.value,originalRows:rows,
      startRow,nameCol,idCol,searchBy:'id',mode:single?'single':'batch',
      items:single?[ready[0]]:items,index:0,running:true,phase:'lookup',
      page:0,candidates:[],previewNames:[],resultCount:0,pendingPostId:'',searchAttempts:0,
      startedAt:new Date().toLocaleString('ru-RU'),logs:[]
    };
    save();
    setMessage(single?`Тест одного товара: ID ${ready[0].id}.`:`Обработка ${ready.length} товаров началась.`);
    delayStep();
  }

  function startManualTest() {
    if(busy||state?.running) return setMessage('Сначала приостановите действующую задачу.');
    const id=clean($('#sbp-single-id')?.value);
    const name=clean($('#sbp-single-name')?.value);
    if(!/^[1-9]\d*$/.test(id)) return setMessage('Введите числовой WordPress ID товара (не SKU).');
    if(state?.items?.length&&state.index<state.items.length&&!confirm('Предыдущий прогресс будет заменён. Продолжить?')) return;
    const item={sourceRow:2,id,name,status:WAIT,note:'',foundName:'',wpPostId:''};
    state={
      fileName:`test_product_${id}.xlsx`,sheetName:'Один товар',
      originalRows:[['ID','Название'],[id,name]],startRow:2,nameCol:1,idCol:0,
      searchBy:'id',mode:'single',items:[item],index:0,running:true,phase:'lookup',
      page:0,candidates:[],previewNames:[],resultCount:0,pendingPostId:'',
      searchAttempts:0,startedAt:new Date().toLocaleString('ru-RU'),logs:[]
    };
    save();
    setMessage(`Тестовая проверка товара ${id}. Сначала прочитаем карточку, затем запросим подтверждение.`);
    delayStep();
  }

  $('#sbp-start').addEventListener('click', () => startImportedQueue(false));
  $('#sbp-test-first').addEventListener('click', () => startImportedQueue(true));
  $('#sbp-test-manual').addEventListener('click', startManualTest);

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

  function renderStatus(){
    const button=$('#sbp-toggle'),exportButton=$('#sbp-export'),startButton=$('#sbp-start');
    const testFirst=$('#sbp-test-first'),testManual=$('#sbp-test-manual');
    const status=$('#sbp-status');if(!status) return;
    const hasData=Boolean(state?.items?.length);
    button.disabled=!hasData||state.phase==='idle'||state.index>=state.items.length;
    button.textContent=state?.running?'Пауза':'Продолжить';
    exportButton.disabled=!hasData;
    startButton.disabled=!imported||Boolean(state?.running)||busy;
    testFirst.disabled=!imported||Boolean(state?.running)||busy;
    testManual.disabled=Boolean(state?.running)||busy;
    fileInput.disabled=Boolean(state?.running);
    const fill=$('#sbp-progress-fill'),stats=$('#sbp-stats'),logs=$('#sbp-log');
    if(!hasData){
      status.textContent='Нет активной задачи.';
      stats.textContent='Нет задачи';
      fill.style.width='0%';
      if(logs) logs.textContent='';
      return;
    }
    const counts={};
    state.items.forEach(item=>{counts[item.status]=(counts[item.status]||0)+1;});
    const processed=state.items.filter(item=>item.status!==WAIT&&item.status!=='Отправлен в Trash, проверяем').length;
    fill.style.width=`${Math.round(100*processed/state.items.length)}%`;
    stats.textContent=`Обработано ${processed}/${state.items.length} · в корзине ${counts[OK]||0}`;
    const current=state.items[state.index];
    status.textContent=[
      `${state.running?'Выполняется':state.phase==='idle'?'Завершено':'Приостановлено'} · ${state.mode==='single'?'один товар':'весь Excel'}`,
      `Поиск: только по WordPress ID товара`,
      `Текущий: ${current?`${current.name||'Без названия'} [${current.id}]`:'—'}`,
      ...Object.entries(counts).map(([s,n])=>`${s}: ${n}`)
    ].join('\n');
    if(logs) logs.textContent=(state.logs||[]).map(x=>`${x.time}  ${x.message}`).join('\n');
  }

  renderStatus();
  if (state?.running) {
    const item = state.items?.[state.index];
    if (item) setMessage(`Автопродолжение: ${item.name} (${item.id}). Этап: ${state.phase}.`);
    delayStep();
  }
})();
