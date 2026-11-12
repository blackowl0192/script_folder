// ==UserScript==
// @name         Product Manager — Trash / Private
// @namespace    https://steambuy.pro/
// @version      2.2.4
// @description  Excel product_link → выбрать действие Trash или Private → открыть товар → применить → проверить; 7 сайтов.
// @homepageURL  https://github.com/blackowl0192/script_folder/blob/main/steambuy_excel_trash_v1_2.user.js
// @updateURL    https://raw.githubusercontent.com/blackowl0192/script_folder/main/steambuy_excel_trash_v1_2.user.js
// @downloadURL  https://raw.githubusercontent.com/blackowl0192/script_folder/main/steambuy_excel_trash_v1_2.user.js
// @match        https://keymplace.com/wp-admin/*
// @match        https://keymplace.com/wp-login.php*
// @match        https://playfun.pro/wp-admin/*
// @match        https://playfun.pro/wp-login.php*
// @match        https://prime2key.com/wp-admin/*
// @match        https://prime2key.com/wp-login.php*
// @match        https://digitplaystore.com/wp-admin/*
// @match        https://digitplaystore.com/wp-login.php*
// @match        https://googolsell.com/wp-admin/*
// @match        https://googolsell.com/wp-login.php*
// @match        https://steambuy.pro/wp-admin/*
// @match        https://steambuy.pro/wp-login.php*
// @match        https://ggmarket.pro/wp-admin/*
// @match        https://ggmarket.pro/wp-login.php*
// @noframes
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @run-at       document-idle
// ==/UserScript==

(() => {
  'use strict';
  // Keep v3 queues, but require explicit action selection after upgrading to v2.2.
  const VERSION = '2.2.4';
  const DOMAIN = location.hostname;
  const PREFIX = 'sb-product-trash-v3:' + DOMAIN;
  const OLD_LEASE = 'sb-product-trash-v2:' + DOMAIN + ':lease';
  const KEY = PREFIX + ':state';
  const LEASE = PREFIX + ':lease';
  const LAYOUT = PREFIX + ':layout';
  const TAB = PREFIX + ':tab-id';
  const LIST = location.origin + '/wp-admin/edit.php?post_type=product';
  const MAX_LOGS = 1500;
  const PAGESIZE = 25;
  const OK = 'done';
  const TERMINAL = new Set(['done', 'not_found', 'failed', 'unconfirmed', 'skipped']);
  const $ = (s, p = root) => p.querySelector(s);
  const $$ = (s, p = root) => Array.from(p.querySelectorAll(s));
  const clean = x => String(x == null ? '' : x).trim();
  const idValid = x => /^[1-9]\d*$/.test(clean(x));
  const esc = x => clean(x).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const normalized = x => clean(x).normalize('NFKC').toLowerCase().replace(/\s+/g, ' ');
  const now = () => new Date().toISOString();
  const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
  const tabId = sessionStorage.getItem(TAB) || (Date.now().toString(36) + Math.random().toString(36).slice(2));
  sessionStorage.setItem(TAB, tabId);
  const defaults = () => ({
    v: 3, schema: 2, action: '', filename: '', sheet: '', rows: [], logs: [], index: -1,
    phase: 'idle', running: false, paused: false, mode: 'batch', runCount: 0,
    owner: '', pending: null,
    settings: { delay: 2, stopOnError: false }, updated: '', started: ''
  });
  let state = GM_getValue(KEY, null) || defaults();
  if (state.v !== 3 || !Array.isArray(state.rows)) state = defaults();
  state.settings = Object.assign({delay: 2}, state.settings || {});
  // Older v2.1 runs were delete-only. Pause on upgrade rather than silently continuing Trash.
  if(state.schema !== 2){
    const currentRow=state.rows[state.index];
    const trashPending=!!(state.pending && currentRow &&
      ['verify','verifying','deleting'].includes(state.phase));
    state.running=false;state.paused=!!state.rows.length;
    state.action=trashPending?'trash':'';
    if(trashPending) currentRow.action='trash';
    else state.started='';
    state.schema=2;
    state.logs.push({time:new Date().toISOString(),level:'warn',id:currentRow?.id||'',
      row:currentRow?.excelRow||'',message:'Обновление v2.2: автоматическая работа приостановлена. Выберите действие перед продолжением.'});
    GM_setValue(KEY,state);
  }
  let root, panel, book = null, fileName = '';
  let activeTab = 'queue', filter = 'all', search = '', page = 0;
  let busy = false, moving = false, timer = null, heartbeat = null, expanded = false, dragging = null;
  const labels = {pending:'Ожидает', opening:'Открытие', deleting:'Удаляется',
    saving:'Сохранение', verifying:'Проверка',done:'Выполнено',
    not_found:'Не найден', failed:'Ошибка обработки', unconfirmed:'Не подтверждено', skipped:'Пропущен'};
  const ACTION_LABELS = {trash:'Переместить в Trash', private:'Сделать приватным'};
  const actionName = action => ACTION_LABELS[action] || 'Не выбрано';
  const rowLabel = r => r.status==='done'
    ? (r.action==='private'?'Приватный':'Удалён') : (labels[r.status]||r.status);
  const statuses = {done:'done', failed:'bad', not_found:'bad', unconfirmed:'uncertain', skipped:'neutral',
    opening:'work', deleting:'work', saving:'work', verifying:'work', pending:'neutral'};
  function persist() {
    state.updated = now();
    GM_setValue(KEY, state);
    // A panel rendering failure must not block processing the product queue.
    try { render(); }
    catch (e) {
      const message='Ошибка обновления панели: '+(e?.message||String(e));
      console.error('[Product Manager]',message,e);
      const last=state.logs[state.logs.length-1];
      if(!last||last.message!==message){
        state.logs.push({time:now(),level:'error',id:current()?.id||'',
          row:current()?.excelRow||'',message});
        if(state.logs.length>MAX_LOGS)state.logs.splice(0,state.logs.length-MAX_LOGS);
        GM_setValue(KEY,state);
      }
    }
  }
  function log(level, message, row) {
    state.logs.push({time:now(),level, id:row ? row.id : '', row:row ? row.excelRow : '',message});
    if (state.logs.length > MAX_LOGS) state.logs.splice(0,state.logs.length - MAX_LOGS);
    persist();
  }
  function toast(message) {
    if (!panel) return;
    $('.notice')?.remove();
    const box = document.createElement('div');
    box.className = 'notice';
    box.textContent = message;
    const x = document.createElement('button');
    x.textContent = '×';
    x.title = 'Закрыть';
    x.onclick = () => box.remove();
    box.prepend(x);
    panel.appendChild(box);
    setTimeout(() => box.remove(), 14000);
  }
  function ownLease() {
    const lease = GM_getValue(LEASE, null);
    return lease && lease.id === tabId && lease.until > Date.now();
  }
  function claim() {
    const previous = GM_getValue(OLD_LEASE, null);
    if(previous && previous.until > Date.now()) {
      log('warn','Старт заблокирован: предыдущая версия скрипта ещё работает в другой вкладке.');
      toast('Другая вкладка ещё работает на старой версии скрипта. Остановите её, прежде чем начинать.');
      return false;
    }
    const lease = GM_getValue(LEASE, null);
    if (lease && lease.until > Date.now() && lease.id !== tabId) {
      log('warn','Старт заблокирован: очередь управляется другой вкладкой этого сайта.');
      toast('Обработка уже запущена в другой вкладке этого сайта. Сначала остановите её там.');
      return false;
    }
    GM_setValue(LEASE, {id:tabId, until:Date.now() + 25000});
    const good = ownLease();
    if (good) startHeartbeat();
    return good;
  }
  function startHeartbeat() {
    clearInterval(heartbeat);
    heartbeat = setInterval(() => {
      if (!state.running || !ownLease()) { clearInterval(heartbeat); return; }
      GM_setValue(LEASE, {id:tabId,until:Date.now() + 25000});
    },4000);
  }
  function release() {
    clearInterval(heartbeat);
    if (ownLease()) GM_setValue(LEASE, null);
  }
  function readState() {
    const fresh = GM_getValue(KEY, null);
    if (fresh?.v === 3 && Array.isArray(fresh.rows)) state = fresh;
  }
  function canEdit() {
    const lease = GM_getValue(LEASE,null);
    if (busy || (state.running && (state.owner===tabId || (lease && lease.until>Date.now())))) {
      toast('Сначала приостановите обработку.'); return false;
    }
    if (lease && lease.until > Date.now() && lease.id !== tabId) {
      toast('Очередь используется другой вкладкой.'); return false;
    }
    if (state.running && state.owner!==tabId) {
      state.running=false;state.paused=true;persist();
    }
    return true;
  }
  function move(url, row) {
    if (moving) return;
    const u = new URL(url, location.href);
    if (u.origin !== location.origin || !u.pathname.startsWith('/wp-admin/')) {
      throw new Error('Заблокирован переход на сторонний URL.');
    }
    moving = true;
    state.phase='open';
    if (row) {
      row.status='opening';
      row.message='Переходим по product_link: '+u.href;
    }
    log('info','Открываем редактор по прямой ссылке: '+u.href,row);
    persist();
    // Direct top-frame navigation; preserve ALL WPML query parameters (lang, trid, update_needed...).
    window.location.assign(u.href);
    setTimeout(() => {
      if (!moving || document.visibilityState === 'hidden') return;
      moving = false;
      state.running = false;
      state.paused = true;
      persist(); release();
      log('error','Переход по product_link не завершился. Адрес: '+u.href,row);
      toast('Переход не завершился. Обработка остановлена; проверьте страницу и нажмите «Продолжить».');
    }, 18000);
  }
  function current() { return state.rows[state.index] || null; }
  function nextPending() { return state.rows.findIndex(r => r.status === 'pending'); }
  function schedule(ms = 80) {
    clearTimeout(timer);
    if (state.running && !busy && !moving) timer = setTimeout(()=>{
      try{runStep();}
      catch(e){
        const row=current();
        const message='Неожиданная ошибка скрипта: '+(e?.message||String(e));
        if(row&&!TERMINAL.has(row.status))finalize(row,'failed',message);
        else {log('error',message,row);state.running=false;state.paused=true;persist();release();}
      }
    },ms);
  }
  function pause(reason = 'Пауза по запросу пользователя') {
    clearTimeout(timer);
    state.running = false;
    state.paused = true;
    log('info', reason, current());
    release();
  }
  function stop() { pause('Остановлено пользователем. Сохранённые результаты не потеряны.'); }
  function finalize(row, status, message) {
    if (!row) return;
    if(!row.action && ['trash','private'].includes(state.action))row.action=state.action;
    row.status = status;
    row.message = message;
    row.finished = now();
    state.pending = null;
    state.phase = 'idle';
    state.index = -1;
    state.runCount++;
    log(status === 'done' ? 'ok' : (status==='not_found'||status==='unconfirmed'?'warn':'error'), '['+actionName(row.action)+'] '+message, row);
    if (state.mode === 'one' && state.runCount >= 1) {
      state.running = false;
      state.paused = true;
      log('info','Тест одного товара завершён. Остальная очередь не затронута.');
      release();
      return;
    }
    if (nextPending() < 0) {
      state.running = false;
      state.paused = false;
      log('ok','Обработка файла закончена. Можно скачать отчёт.');
      release();
    } else {
      schedule(Math.max(600,Number(state.settings.delay) * 1000));
    }
  }
  function failFatal(message,row) {
    if (row && !TERMINAL.has(row.status)) {
      row.status='failed';row.message=message;row.finished=now();
    }
    state.running=false;state.paused=true;state.index=-1;state.phase='idle';state.pending=null;
    log('error','Обработка остановлена: ' + message,row);
    release();toast(message);
  }

  // A product_link must point to the editor of a product on the CURRENT site.
  function parseProductLink(text) {
    const raw=clean(text);
    if(!raw)throw new Error('Пустая ссылка product_link.');
    let u;
    try{u=new URL(raw,location.origin);}catch{throw new Error('Неверный URL product_link.');}
    if(u.origin!==location.origin||u.protocol!=='https:'||u.username||u.password)throw new Error('Ссылка относится к другому сайту или использует не HTTPS.');
    if(u.pathname!=='/wp-admin/post.php')throw new Error('Нужна ссылка вида /wp-admin/post.php?post=ID&action=edit.');
    const id=u.searchParams.get('post');
    if(!idValid(id))throw new Error('В product_link нет числового параметра post=ID.');
    if(![null,'edit'].includes(u.searchParams.get('action')))throw new Error('Ссылка должна открывать редактор (action=edit), а не другое действие.');
    u.searchParams.set('action','edit');u.hash='';
    return {id,url:u.href};
  }
  function isProductEditor(row) {
    const u=new URL(location.href);
    return u.origin===location.origin&&u.pathname==='/wp-admin/post.php'&&
      [null,'edit'].includes(u.searchParams.get('action'))&&u.searchParams.get('post')===row.id;
  }
  function safeTrashLink(link,row) {
    if(!link)return null;
    let u;
    try{u=new URL(link.getAttribute('href'),location.href);}catch{return null;}
    if(u.origin!==location.origin||u.username||u.password||u.pathname!=='/wp-admin/post.php'||
      u.searchParams.get('post')!==row.id||
      u.searchParams.get('action')!=='trash'||!u.searchParams.get('_wpnonce'))return null;
    return u.href;
  }
  // WordPress Classic Editor: Edit visibility -> Private -> OK -> Update.
  function makePrivate(row,status) {
    if(status==='private')
      return finalize(row,'done','Товар уже приватный. Update не требуется.');
    const edit=document.querySelector('a.edit-visibility[href="#visibility"]');
    const radio=document.querySelector('#visibility-radio-private');
    const ok=document.querySelector('a.save-post-visibility[href="#visibility"]');
    const update=document.querySelector('#publish');
    const missing=[];
    if(!edit)missing.push('Edit visibility');
    if(!radio)missing.push('#visibility-radio-private');
    if(!ok)missing.push('OK / save-post-visibility');
    if(!update)missing.push('Update / #publish');
    if(missing.length)return finalize(row,'failed',
      'Редактор открыт, но не найдены элементы приватности: '+missing.join(', ')+'.');
    if(update.disabled)return finalize(row,'failed','Кнопка Update недоступна (disabled).');
    try{
      log('info','ID '+row.id+': открываем Edit visibility.',row);
      edit.click();
      radio.click(); // Triggers native input/change listeners for Classic Editor.
      if(!radio.checked)throw new Error('Не удалось выбрать radio Private.');
      ok.click();
      if(!radio.checked)throw new Error('После OK переключатель Private не выбран.');
    }catch(e){
      return finalize(row,'failed','Не удалось выбрать Private → OK: '+(e.message||e));
    }
    // Persist BEFORE Update: after a page reload, only verification is permitted.
    state.phase='verify-private';
    state.pending={id:row.id,action:'private',at:now()};
    row.status='saving';
    row.message='Private выбран и подтверждён через OK; выполняется Update.';
    log('info','ID '+row.id+': выбран Private, подтверждено OK. Нажимаем Update.',row);
    try{
      update.click();
    }catch(e){
      return finalize(row,'failed','Не удалось нажать Update: '+(e.message||e)+
        '. Повторный Update автоматически не выполняется.');
    }
    // If an on-page handler suppressed navigation, verify server-side without submitting again.
    setTimeout(()=>{
      if(state.running&&ownLease()&&current()?.id===row.id&&
        state.phase==='verify-private'&&!busy&&!moving)verifyPrivate(row);
    },4500);
  }
  async function verifyPrivate(row) {
    if(busy)return;
    busy=true;
    row.status='verifying';
    row.message='Проверка статуса Private после Update (без повторного сохранения).';
    persist();
    try{
      // Fetch a fresh WordPress editor response; current page may still contain unsaved changes.
      for(let attempt=1;attempt<=3;attempt++){
        const doc=await fetchAdmin(row.productLink);
        const error=doc.querySelector('.wp-die-message');
        if(error)throw new Error('WordPress вернул ошибку: '+clean(error.textContent).slice(0,180));
        const id=clean(doc.querySelector('#post_ID,input[name="post_ID"]')?.value);
        const type=clean(doc.querySelector('#post_type,input[name="post_type"]')?.value)||
          (doc.body?.classList.contains('post-type-product')?'product':'');
        const savedStatus=clean(doc.querySelector('#post_status,input[name="post_status"]')?.value)||
          clean(doc.querySelector('#original_post_status,input[name="original_post_status"]')?.value);
        if(id!==row.id || type!=='product'){
          throw new Error('Не удалось удостовериться в ID/типе товара после Update. Получено ID='+
            (id||'—')+', тип='+(type||'—')+'.');
        }
        if(savedStatus==='private'){
          return finalize(row,'done','Товар ID '+row.id+' успешно сохранён как Private.');
        }
        if(attempt<3){
          log('warn','ID '+row.id+': статус после Update = '+(savedStatus||'неизвестен')+
            '. Повторная проверка без нажатия Update.',row);
          await delay(1100);
        }else{
          return finalize(row,'failed','Private не подтверждён после Update. Текущий статус: '+
            (savedStatus||'неизвестен')+'. Повторный Update автоматически не выполняется.');
        }
      }
    }catch(e){
      finalize(row,'failed','Ошибка проверки приватности ID '+row.id+': '+(e.message||e)+
        '. Повторного Update не будет.');
      if(/сессия|войти|авторизац/i.test(e.message||'')){
        state.running=false;state.paused=true;persist();release();
      }
    }finally{
      busy=false;
      if(state.running)schedule(Math.max(700,Number(state.settings.delay)*1000));
    }
  }
  function openProduct(row) {
    // Action is fixed for the row once processing starts, regardless of UI changes.
    if(!['trash','private'].includes(row.action)) row.action=state.action;
    if(!['trash','private'].includes(row.action))return failFatal('Выберите действие перед запуском.',row);
    if(!isProductEditor(row)){
      row.openAttempts=(row.openAttempts||0)+1;
      if(row.openAttempts>2)return finalize(row,'not_found',
        'По ссылке дважды не удалось открыть редактор товара. Возможна переадресация или отсутствие товара.');
      row.status='opening';row.message='Открываем ссылку товара.';
      state.phase='open';
      log('info','Переход к product_link: '+row.productLink,row);
      return move(row.productLink,row);
    }
    const err=document.querySelector('.wp-die-message,#loginform');
    if(err){
      const message=clean(err.textContent).slice(0,200);
      if(/log.?in|войти|авторизац/i.test(message))return failFatal('Требуется вход в WordPress: '+message,row);
      return finalize(row,'not_found','По ссылке не открылся редактор: '+(message||'ошибка WordPress'));
    }
    const form=document.querySelector('form#post');
    const hiddenId=clean(document.querySelector('#post_ID,input[name="post_ID"]')?.value);
    const bodyId=(document.body.className.match(/\bpostid-(\d+)\b/)||[])[1]||'';
    const actualId=hiddenId||bodyId;
    log('info','Карточка загружена: URL='+location.href+
      '; post_ID='+(hiddenId||'—')+'; body ID='+(bodyId||'—')+
      '; form#post='+(form?'да':'нет')+'; режим='+actionName(row.action)+'.',row);
    const type=clean(document.querySelector('#post_type,input[name="post_type"]')?.value)||
      (document.body.classList.contains('post-type-product')?'product':'');
    const status=clean(document.querySelector('#post_status,input[name="post_status"]')?.value);
    if(hiddenId&&hiddenId!==row.id)return finalize(row,'failed',
      'В редакторе обнаружен post_ID '+hiddenId+', а в Excel указан '+row.id+'.');
    if(bodyId&&bodyId!==row.id)return finalize(row,'failed',
      'WordPress body сообщает ID '+bodyId+', а в Excel указан '+row.id+'.');
    // Some editor variants omit form#post or post_ID. The matched URL and a
    // nonce-protected #delete-action link (checked below) still identify the item.
    if(!actualId && !form && !document.querySelector('#delete-action'))
      return finalize(row,'not_found','По ссылке не открылась карточка товара WordPress.');
    if(type!=='product')return finalize(row,'failed',
      'Карточка ID '+row.id+': не удалось подтвердить тип product (найден: '+(type||'неизвестно')+').');
    row.wpPostId=actualId||row.id;
    row.foundName=clean(document.querySelector('#title,input[name="post_title"]')?.value)||
      clean(document.querySelector('.editor-post-title__input')?.textContent)||'';
    if(status==='trash') {
      if(row.action==='trash')return finalize(row,'done','Товар уже находится в корзине.');
      return finalize(row,'failed','Нельзя сделать приватным товар, который уже находится в корзине.');
    }
    if(row.action==='private')return makePrivate(row,status);
    const anchor=document.querySelector('#delete-action a.submitdelete.deletion')||
      document.querySelector('#delete-action a.submitdelete');
    const link=safeTrashLink(anchor,row);
    if(!link) {
      const raw=anchor?.getAttribute('href')||'';
      let reason='Не найдена ссылка #delete-action → Move to Trash.';
      if(anchor && raw) {
        try{
          const ref=new URL(raw,location.href),target=ref.searchParams.get('post');
          if(target!==row.id)
            reason='Ссылка редактора относится к товару ID '+row.id+
              ', но Move to Trash ведёт на ID '+(target||'неизвестно')+
              '. Удаление другого товара заблокировано.';
          else reason='Найдена кнопка Move to Trash, но её URL некорректен или отсутствует _wpnonce.';
        }catch{reason='Некорректный URL кнопки Move to Trash.';}
      }
      return finalize(row,'failed',reason);
    }
    // A persisted verify checkpoint ensures a reload never clicks Move to Trash a second time.
    state.phase='verify';state.pending={id:row.id,action:'trash',href:link,at:now()};
    row.status='deleting';row.message='Отправляем Move to Trash и ждём перенаправление WordPress.';
    log('info','ID '+row.id+': кнопка #delete-action → Move to Trash найдена. Открываем её WP URL.',row);
    // WordPress button is a nonce-protected GET link. Navigating to its href is
    // equivalent to clicking the link, but avoids JS handlers swallowing the click.
    // State is saved BEFORE the request, so reloading never submits Trash again.
    moving=true;
    try {window.location.assign(link);}
    catch(error){
      moving=false;
      return finalize(row,'failed','Ошибка перехода по Move to Trash: '+(error.message||error));
    }
    // If navigation is unexpectedly blocked, verify the outcome instead of retrying.
    setTimeout(()=>{
      if(state.running&&ownLease()&&current()?.id===row.id&&state.phase==='verify'&&!busy){
        moving=false;
        log('warn','После Trash страница не изменилась. Проверяем только ответ на текущей странице.',row);
        verifyTrash(row);
      }
    },4000);
  }
  function successRedirect(row) {
    const u=new URL(location.href);
    if(u.origin!==location.origin||u.pathname!=='/wp-admin/edit.php'||
      u.searchParams.get('trashed')!=='1')return false;
    const ids=u.searchParams.get('ids');
    // WordPress and WPML may omit post_type and ids from a normal redirect.
    // Explicitly mismatched IDs are NEVER accepted.
    return !ids || ids.split(/[,\s]+/).includes(row.id);
  }
  async function fetchAdmin(url) {
    const response=await fetch(url,{credentials:'same-origin',redirect:'follow',cache:'no-store'});
    if(!response.ok && response.status!==403)throw new Error('HTTP '+response.status);
    const destination=new URL(response.url);
    if(destination.origin!==location.origin||destination.pathname.endsWith('/wp-login.php'))
      throw new Error('Истекла сессия WordPress.');
    const doc=new DOMParser().parseFromString(await response.text(),'text/html');
    if(doc.querySelector('#loginform'))throw new Error('Требуется войти в WordPress.');
    return doc;
  }
  function verifyTrash(row) {
    if(busy)return;
    busy=true;
    row.status='verifying';
    row.message='Читаем подтверждение WordPress после Move to Trash.';
    persist();
    try {
      const u=new URL(location.href);
      const ids=u.searchParams.get('ids');
      const wrongId=ids&&u.searchParams.get('trashed')==='1'&&
        !ids.split(/[,\s]+/).includes(row.id);
      if(wrongId){
        return finalize(row,'failed','После Trash WordPress указал другой ID ('+ids+
          '), а ожидался '+row.id+'. Повторное удаление запрещено.');
      }
      const messages=Array.from(document.querySelectorAll(
        '#wpbody-content .notice, #wpbody-content .updated, #wpbody-content #message, #wpbody-content .error, .wp-die-message'
      )).map(node=>clean(node.textContent)).filter(Boolean);
      const wpml=messages.filter(t=>
        /you are trying to edit a translation that is currently in the process of being added using wpml/i.test(t)
      );
      // WPML warns about translation management, not whether Trash succeeded.
      for(const warning of wpml)
        log('warn','Предупреждение WPML (не ошибка Trash): '+warning.slice(0,220),row);
      if(successRedirect(row)) {
        return finalize(row,'done','WordPress подтвердил Move to Trash (trashed=1). ID '+row.id+'.');
      }
      const positive=messages.find(t=>
        /(?:moved|sent|put).{0,90}trash|trash.{0,90}(?:moved|sent|put)|перемещен.{0,60}корзин|отправлен.{0,60}корзин/i.test(t)
      );
      const listPage=u.origin===location.origin&&u.pathname==='/wp-admin/edit.php'&&
        [null,'product'].includes(u.searchParams.get('post_type'));
      if(listPage&&positive)
        return finalize(row,'done','WordPress подтвердил Trash уведомлением: '+positive.slice(0,170));
      // After a successful WPML hook the normal trashed=1 parameters can be lost.
      // Navigating back to the products list after the nonce-protected Trash URL
      // is treated as completion, with a clear caveat in the operator log.
      const realErrors=messages.filter(t=>!wpml.includes(t)&&
        /(?:cannot|could not|not allowed|permission denied|invalid nonce|link you followed has expired|failed to|не удалось|нет прав|ошибка удален)/i.test(t)
      );
      if(listPage && wpml.length && !realErrors.length) {
        log('warn','После Move to Trash открыт список товаров без trashed=1. WPML мог удалить параметры подтверждения; статус принят по перенаправлению, без проверки корзины.',row);
        return finalize(row,'done','Переход на список товаров после Move to Trash выполнен. WordPress не показал стандартное подтверждение.');
      }
      const liveStatus=clean(document.querySelector('#post_status,input[name="post_status"]')?.value);
      const liveId=clean(document.querySelector('#post_ID,input[name="post_ID"]')?.value);
      if(liveId===row.id && liveStatus==='trash')
        return finalize(row,'done','Страница WordPress показывает статус trash для ID '+row.id+'.');
      if(realErrors.length)
        return finalize(row,'failed','WordPress сообщил об ошибке удаления: '+realErrors[0].slice(0,200));
      if(document.querySelector('#loginform'))
        return finalize(row,'unconfirmed','Сессия WordPress завершилась после запроса Trash; результат неизвестен.');
      return finalize(row,'unconfirmed',
        'Команда Move to Trash отправлена, но ответ WordPress не содержит явного подтверждения. '+
        (wpml.length?'Показано предупреждение WPML, которое не означает ошибку удаления. ':'')+
        'Результат требует ручной проверки. URL: '+location.href);
    }catch(e) {
      finalize(row,'unconfirmed','Не удалось прочитать ответ WordPress после Trash: '+
        (e.message||e)+'. Не повторяем удаление; результат требует проверки.');
    }finally {
      busy=false;
      if(state.running)schedule(Math.max(600,Number(state.settings.delay)*1000));
    }
  }
  function runStep() {
    if(!state.running||busy||moving)return;
    if(!ownLease()){
      state.running=false;state.paused=true;persist();toast('Очередь остановлена: управление у другой вкладки.');
      return;
    }
    if(state.index<0||!current()){
      state.index=nextPending();
      if(state.index<0){state.running=false;state.phase='idle';log('ok','Очередь закончена.');release();return;}
      state.phase='open';persist();
    }
    const row=current();
    if(!row||TERMINAL.has(row.status)){
      state.index=-1;state.phase='idle';persist();return schedule();
    }
    if(['verify','verifying','deleting','verify-trash'].includes(state.phase)){
      log('warn','Возобновление: проверка предыдущего Trash. Второго нажатия не будет.',row);
      return verifyTrash(row);
    }
    if(['verify-private','saving'].includes(state.phase)){
      log('warn','Возобновление: проверяем сохранённую приватность без повторного Update.',row);
      return verifyPrivate(row);
    }
    state.phase='open';openProduct(row);
  }
  function start(one=false) {
    if(busy||moving)return toast('Дождитесь завершения текущего действия.');
    readState();
    if(!state.rows.length){log('error','Запуск невозможен: очередь product_link пуста.');return toast('Сначала загрузите Excel со столбцом product_link.');}
    const pendingAction=state.pending?.action || (state.pending && ['verify','verifying','deleting'].includes(state.phase)?'trash':'');
    const selected=clean(state.action);
    if(!['trash','private'].includes(selected)&&!pendingAction){log('error','Запуск невозможен: оператор не выбрал Trash или Private.');return toast('Сначала выберите действие: Trash или Private.');}
    if(!claim())return;
    const ongoing=state.index>=0&&!!current()&&!TERMINAL.has(current().status);
    if(!ongoing){state.index=nextPending();state.phase='open';}
    if(state.index<0){release();return toast('Все ссылки обработаны. Загрузите новый Excel.');}
    const row=current();
    if(pendingAction && !row.action)row.action=pendingAction;
    if(row.status==='pending'&&state.phase==='open'){
      const count=state.rows.filter(r=>r.status==='pending').length;
      const action=actionName(state.action);
      if(!confirm(one?
        'ТЕСТ: обработать один товар?\nДействие: '+action+'\n'+row.productLink :
        'Запустить обработку товаров?\nСайт: '+DOMAIN+'\nДействие: '+action+
        '\nОсталось ссылок: '+count)){
        release();return;
      }
    }
    state.mode=one?'one':'batch';
    state.runCount=0;
    if(!row.action)row.action=state.action;
    state.running=true;state.paused=false;state.owner=tabId;
    if(!state.started)state.started=now();
    log('info',(one?'Тест одной ссылки':'Массовая обработка product_link')+' запущен. Действие: '+actionName(row.action)+'.',row);
    schedule();
  }

  function reset() {
    if (!canEdit()) return;
    if (!confirm('Сбросить импорт, результаты и журнал для ' + DOMAIN +
      '?\nТовары, уже помещённые в корзину, не восстановятся.')) return;
    release();
    state=defaults();book=null;fileName='';
    $('#file').value='';
    $('#filename').textContent='Файл не выбран';
    $('#mapping').classList.add('hidden');
    persist();
    setTab('import');
    toast('Очередь и журнал очищены.');
  }
  function sheetRows() {
    const name = book?.SheetNames[Number($('#sheet').value)];
    return name ? XLSX.utils.sheet_to_json(book.Sheets[name],
      {header:1,raw:false,defval:'',blankrows:true}) : [];
  }
  function linkColumn() {
    return (sheetRows()[0]||[]).findIndex(x=>normalized(x).replace(/[\s\-]+/g,'_')==='product_link');
  }
  function linkCell(data,r,c) {
    const shown=clean(data[r]?.[c]);
    const sheet=book.Sheets[book.SheetNames[Number($('#sheet').value)]];
    const target=clean(sheet[XLSX.utils.encode_cell({r,c})]?.l?.Target);
    return /^https?:\/\//i.test(shown)||shown.startsWith('/wp-admin/')?shown:(target||shown);
  }
  function fillMapping() {
    if(!book)return;
    const col=linkColumn();
    $('#linkColumn').textContent=col<0?'Колонка product_link не найдена в первой строке.':
      'Найдена колонка product_link (№ '+(col+1)+').';
    $('#load').disabled=col<0;
    const rows=sheetRows();
    $('#preview').textContent=col<0?'В первой строке должен быть заголовок product_link.':
      rows.slice(1).map((r,i)=>clean(linkCell(rows,i+1,col))).filter(Boolean).slice(0,5).join('\n')||
      'Нет ссылок.';
  }
  function preview(){fillMapping();}

  async function readExcel(file) {
    if (!file || !canEdit()) return;
    try {
      if (typeof XLSX==='undefined')throw new Error('Библиотека XLSX недоступна.');
      if (file.size>25*1024*1024)throw new Error('Максимальный размер файла — 25 МБ.');
      book=XLSX.read(await file.arrayBuffer(),{type:'array'});
      fileName=file.name;
      $('#filename').textContent=file.name + ' · ' + book.SheetNames.length + ' листов';
      $('#sheet').innerHTML=book.SheetNames.map((name,i)=>'<option value="'+i+'">'+esc(name)+'</option>').join('');
      $('#mapping').classList.remove('hidden');
      fillMapping();
      setTab('import');
    } catch (e) {log('error','Не удалось прочитать Excel: '+e.message);toast('Не удалось прочитать Excel: '+e.message);}
  }
  function importQueue(){
    if(!book||!canEdit())return;
    const data=sheetRows(),col=linkColumn();
    if(col<0)return toast('В первой строке Excel нужен столбец product_link.');
    const seen=new Set(),rows=[];
    for(let i=1;i<data.length;i++){
      const raw=clean(linkCell(data,i,col));
      if(!raw)continue;
      const row={excelRow:i+1,productLink:raw,id:'',name:'',status:'pending',
        message:'',foundName:'',wpPostId:'',finished:''};
      try{
        const p=parseProductLink(raw);
        row.productLink=p.url;row.id=p.id;
        if(seen.has(row.id))throw new Error('Дубликат product_link с ID '+row.id+'.');
        seen.add(row.id);
      }catch(e){row.status='skipped';row.message=e.message;}
      rows.push(row);
    }
    if(!rows.length)return toast('В product_link нет ссылок.');
    const good=rows.filter(r=>r.status==='pending').length;
    if(!good)return toast('Нет корректных ссылок на редактор товаров этого сайта.');
    if(state.rows.length&&!confirm('Заменить прежнюю очередь для '+DOMAIN+'? Вы можете сначала скачать отчёт.'))return;
    const chosenAction=state.action;
    state=defaults();state.action=chosenAction;state.filename=fileName;state.sheet=book.SheetNames[Number($('#sheet').value)];state.rows=rows;
    for(const r of rows){
      if(r.status!=='skipped')continue;
      state.logs.push({time:now(),level:'warn',id:r.id,row:r.excelRow,
        message:'Пропуск ссылки '+r.productLink+': '+r.message});
    }
    if(state.logs.length>MAX_LOGS)state.logs.splice(0,state.logs.length-MAX_LOGS);
    log('info','Импорт '+fileName+': '+rows.length+' ссылок, к обработке '+good+', пропущено '+(rows.length-good)+'. Действие: '+actionName(state.action)+'.');
    setTab('queue');
  }
  function report(){
    if(!state.rows.length)return toast('Очередь пуста.');
    if(typeof XLSX==='undefined')return toast('Библиотека XLSX недоступна.');
    const cols=['Строка Excel','product_link','ID из ссылки','Название на сайте','Действие','Статус','Комментарий','Завершено'];
    const line=r=>[r.excelRow,r.productLink,r.id,r.foundName,actionName(r.action||state.action),rowLabel(r),r.message,r.finished];
    const wb=XLSX.utils.book_new();
    const add=(name,rows)=>{
      const ws=XLSX.utils.aoa_to_sheet([cols,...rows.map(line)]);
      ws['!cols']=[{wch:13},{wch:100},{wch:18},{wch:50},{wch:30},{wch:24},{wch:105},{wch:25}];
      XLSX.utils.book_append_sheet(wb,ws,name);
    };
    add('Все товары',state.rows);
    add('Удалено',state.rows.filter(r=>r.status==='done'&&r.action!=='private'));
    add('Приватные',state.rows.filter(r=>r.status==='done'&&r.action==='private'));
    add('Не найдено',state.rows.filter(r=>r.status==='not_found'));
    add('Не подтверждено',state.rows.filter(r=>r.status==='unconfirmed'));
    add('Ошибки обработки',state.rows.filter(r=>r.status==='failed'));
    add('Пропущено',state.rows.filter(r=>r.status==='skipped'));
    const ws=XLSX.utils.aoa_to_sheet([['Время','Уровень','ID','Строка','Сообщение'],
      ...state.logs.map(x=>[x.time,x.level,x.id,x.row,x.message])]);
    ws['!cols']=[{wch:26},{wch:14},{wch:20},{wch:14},{wch:100}];
    XLSX.utils.book_append_sheet(wb,ws,'Журнал');
    const safe=(state.filename||'products').replace(/\.(xlsx|xls|csv)$/i,'').replace(/[\\/:*?"<>|]/g,'_');
    XLSX.writeFile(wb,safe+'_'+DOMAIN+'_product_manager_result.xlsx');
    log('info','Отчёт по ссылкам product_link скачан.');
  }
  function downloadLog() {
    const text=state.logs.map(x=>[new Date(x.time).toLocaleString('ru-RU'),x.level,
      'ID '+(x.id||'—'),x.message].join(' | ')).join('\n');
    const blob=new Blob([text],{type:'text/plain;charset=utf-8'});
    const url=URL.createObjectURL(blob),a=document.createElement('a');
    a.href=url;a.download='trash_log_'+DOMAIN+'.txt';a.click();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
  }
  const CSS = [
    ':host{all:initial;font:13px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:#182438;color-scheme:light}',
    '*{box-sizing:border-box}button,input,select{font:inherit}button{cursor:pointer;border:1px solid #d9e1eb;border-radius:9px;padding:8px 12px;background:#fff;color:#26354c;font-weight:600}button:hover{background:#f0f4fa}button:disabled{opacity:.42;cursor:default}button:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid #4f7df3;outline-offset:2px}',
    '.panel{position:fixed;right:22px;bottom:20px;width:710px;height:650px;max-width:calc(100vw - 24px);max-height:calc(100vh - 40px);background:white;box-shadow:0 18px 70px #0f233b40;border:1px solid #dae2ee;border-radius:16px;display:flex;flex-direction:column;overflow:hidden;resize:both;pointer-events:auto}',
    '.panel.full{inset:24px!important;width:auto!important;height:auto!important;max-width:none;max-height:none;resize:none}',
    '.head{background:#14263e;color:white;padding:16px 18px;display:flex;align-items:center;gap:12px;cursor:move;touch-action:none;user-select:none}.logo{height:36px;width:36px;background:#4275ed;display:grid;place-items:center;border-radius:10px;font-size:18px;font-weight:800}.head b{font-size:16px;letter-spacing:.2px}.head small{color:#b1c2d9;display:block;font-size:11px}.head .spacer{flex:1}.icon{background:#ffffff13;color:#fff;border-color:#ffffff22;padding:4px 9px}',
    '.body{display:flex;flex-direction:column;min-height:0;flex:1}.stats{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;padding:14px 16px 10px}.card{border:1px solid #e4eaf2;border-radius:10px;padding:9px 12px;background:#fafbfd}.card b{display:block;font-size:21px;line-height:1.25}.card span{font-size:11px;color:#68778c}.card.done b{color:#16835e}.card.bad b{color:#bd4f39}',
    '.progress{height:5px;background:#eaf0f8;margin:0 16px;border-radius:5px;overflow:hidden}.progress i{height:100%;display:block;background:#4076ed;transition:width .3s}.current{padding:10px 16px;display:flex;gap:8px;align-items:center;color:#586b83;min-height:42px}.dot{height:8px;width:8px;background:#93a1b3;border-radius:50%;flex:none}.dot.run{background:#3974ee;box-shadow:0 0 0 4px #e5edff}.current span:last-child{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
    '.operation{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin:0 16px 12px;padding:11px 12px;background:#f4f7fc;border:1px solid #dce6f3;border-radius:10px}.operation label{font-weight:750;font-size:12px;color:#31445c}.operation select{flex:1;min-width:200px}.operation small{flex-basis:100%;color:#6b7b92;font-size:11px}.actions{display:flex;gap:7px;flex-wrap:wrap;padding:0 16px 12px}.primary{background:#3468df!important;border-color:#3468df!important;color:white}.danger{color:#b04439;border-color:#eabcb6;background:#fff7f6}.tabs{display:flex;gap:18px;overflow:auto;flex-shrink:0;padding:0 16px;border-bottom:1px solid #e4eaf2}.tabs button{flex-shrink:0;border:0;border-radius:0;padding:10px 0;color:#7a889a;background:none;font-size:12px}.tabs button.on{color:#265cce;border-bottom:2px solid #3468df}',
    '.view{padding:12px 16px;min-height:0;flex:1;overflow:auto}.hidden{display:none!important}.tools{display:flex;gap:8px;margin-bottom:10px}input,select{border:1px solid #dce3ed;border-radius:8px;background:#fff;padding:8px;color:#22324a;min-width:0}input[type=search]{flex:1}table{border-collapse:collapse;width:100%;table-layout:fixed;font-size:12px}th{text-align:left;color:#7c899b;font-size:10px;text-transform:uppercase;letter-spacing:.5px;padding:8px 6px;background:#f7f9fc;position:sticky;top:0}td{padding:10px 6px;border-bottom:1px solid #edf1f6;vertical-align:top}tr.currentrow{background:#eff5ff}td .title{font-weight:600;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}td small{color:#8290a2;display:block;font-size:10px;overflow-wrap:anywhere}.badge{font-size:10px;padding:3px 7px;border-radius:6px;display:inline-block;white-space:nowrap;background:#eef2f7;color:#62738a}.badge.done{background:#e4f5ec;color:#147c55}.badge.bad{background:#fff0e8;color:#b24b28}.badge.uncertain{background:#fff6dc;color:#956400}.badge.work{background:#e7efff;color:#2b62ce}',
    '.foot{padding:10px 16px;border-top:1px solid #e4eaf2;display:flex;align-items:center;gap:8px;background:#fafbfd;font-size:11px;color:#7a889a}.foot span:first-child{flex:1}.foot button{padding:4px 8px;font-size:11px}.empty{text-align:center;padding:28px 16px;color:#77869b}.empty b{display:block;color:#31435c;font-size:16px;margin-bottom:6px}.drop{border:2px dashed #c4d2e9;border-radius:12px;text-align:center;padding:20px;margin:4px 0 12px;background:#f7faff}.drop.over{border-color:#3468df;background:#eaf1ff}.drop label{cursor:pointer;color:#2b63d4;font-weight:700}.drop small{display:block;margin-top:5px;color:#7b8a9e}.mapping{display:grid;grid-template-columns:120px 1fr;gap:8px;align-items:center}.mapping label{color:#61728a}.hint{font-size:12px;color:#6d7e94;margin:10px 0}.logline{display:grid;grid-template-columns:80px 80px 1fr;gap:8px;padding:8px 0;border-bottom:1px solid #edf1f6;font-size:11px;overflow-wrap:anywhere}.logline.error{color:#a83f2f}.logline.ok{color:#147c55}.logline.warn{color:#955e0d}',
    '.pill{position:fixed;right:22px;bottom:20px;border-radius:14px;background:#14263e;color:#fff;padding:12px 18px;box-shadow:0 8px 30px #0f233b40;pointer-events:auto}.pill:hover{background:#24466b;color:white}.notice{position:absolute;bottom:54px;left:16px;right:16px;z-index:8;padding:12px;background:#fff7e7;border:1px solid #e3cdab;border-radius:10px;color:#75551f;box-shadow:0 6px 20px #0002}.notice button{float:right;padding:0 6px;margin-right:5px}.help h3{font-size:13px;color:#263d5c}.help p,.help li{font-size:12px;color:#61728a}.help ol{padding-left:20px}.setting{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid #edf1f6}.setting input[type=number]{width:80px}.setting small{display:block;color:#7a899c}',
    '@media(max-width:550px){.panel{right:8px;bottom:8px;max-width:calc(100vw - 16px)}.stats{gap:4px}.card{padding:7px}.tabs{gap:12px}.actions button{font-size:11px;padding:7px}}'
  ].join('');
  function mount() {
    const host=document.createElement('div');
    host.id='sb-product-trash-app';
    host.style.cssText='position:fixed;z-index:2147483646;inset:0;pointer-events:none';
    document.documentElement.append(host);
    root=host.attachShadow({mode:'open'});
    root.innerHTML='<style>'+CSS+'</style>'+
      '<button class="pill hidden" id="pill">Product Manager · открыть</button>'+
      '<section class="panel" id="panel"><header class="head" id="drag">'+
      '<div class="logo">P</div><div><b>Product Manager</b><small>'+esc(DOMAIN)+' · Trash / Private · v'+VERSION+'</small></div>'+
      '<span class="spacer"></span><button class="icon" id="full" title="На весь экран">⛶</button>'+
      '<button class="icon" id="min" title="Свернуть">—</button></header>'+
      '<div class="body"><div class="stats">'+
      '<div class="card"><b id="total">0</b><span>Товаров в файле</span></div>'+
      '<div class="card done"><b id="done">0</b><span id="doneLabel">Выполнено</span></div>'+
      '<div class="card bad"><b id="missing">0</b><span>Не найдено</span></div>'+
      '<div class="card bad"><b id="errors">0</b><span>Ошибки</span></div>'+
      '<div class="card"><b id="unconfirmed">0</b><span>Без подтверждения</span></div></div>'+
      '<div class="progress"><i id="bar"></i></div>'+
      '<div class="current"><i class="dot" id="dot"></i><span id="current">Загрузите Excel с товарами</span></div>'+
      '<div class="operation"><label for="operation">Действие с товарами</label>'+
      '<select id="operation"><option value="">Выберите действие...</option>'+
      '<option value="trash">Удалить товар (Move to Trash)</option>'+
      '<option value="private">Сделать товар приватным (Private + Update)</option></select>'+
      '<small>Сначала выберите действие, затем нажмите «Тест: 1 товар» или «Начать». После запуска действие нельзя изменить без новой загрузки очереди.</small></div>'+
      '<div class="actions"><button class="primary" id="start">▶ Начать</button><button id="one">Тест: 1 товар</button>'+
      '<button id="pause">Ⅱ Пауза</button><button class="danger" id="stop">■ Стоп</button>'+
      '<button id="report">↓ Отчёт Excel</button></div>'+
      '<nav class="tabs">'+
      '<button data-tab="queue" class="on">Очередь</button>'+
      '<button data-tab="not_found">Не найдены</button>'+
      '<button data-tab="failed">Ошибки обработки</button>'+ 
      '<button data-tab="unconfirmed">Не подтверждено</button>'+
      '<button data-tab="log">Журнал</button>'+
      '<button data-tab="import">Импорт Excel</button>'+
      '<button data-tab="help">Как работать</button></nav>'+
      '<div class="view" id="queue"><div class="tools">'+
      '<input type="search" id="search" placeholder="Фильтр очереди: ссылка, ID, название…">'+
      '<select id="filter"><option value="all">Все статусы</option><option value="pending">В очереди</option>'+
      '<option value="done">Выполнены</option><option value="skipped">Пропущено</option></select></div><div id="rows"></div></div>'+
      '<div class="view hidden" id="not_found"><div id="missingRows"></div></div>'+
      '<div class="view hidden" id="failed"><div id="failedRows"></div></div>'+
      '<div class="view hidden" id="unconfirmed"><div id="unconfirmedRows"></div></div>'+
      '<div class="view hidden" id="log"><div class="tools"><button id="downloadLog">↓ Скачать журнал TXT</button><span class="hint">Последние 1500 событий</span></div><div id="logs"></div></div>'+
      '<div class="view hidden" id="import"><div class="tools"><button class="danger" id="reset">Сбросить очередь и журнал</button></div>'+
      '<div class="drop" id="drop"><label for="file">Выбрать Excel / CSV</label><input type="file" id="file" accept=".xlsx,.xls,.csv" hidden>'+
      '<small>Или перетащите файл сюда</small><small id="filename">Файл не выбран</small></div>'+
      '<div class="hidden" id="mapping"><div class="mapping"><label for="sheet">Лист Excel</label><select id="sheet"></select>'+ 
      '<label>Столбец</label><strong id="linkColumn">—</strong></div>'+
      '<p class="hint">Обязательный столбец: product_link. В нём укажите ссылки вида https://'+esc(DOMAIN)+'/wp-admin/post.php?post=3287&amp;action=edit. Другие столбцы не нужны. ID автоматически извлекается из ссылки.</p>'+
      '<pre id="preview" style="white-space:pre-wrap;overflow-wrap:anywhere;background:#f6f8fc;padding:10px;border-radius:9px"></pre>'+
      '<button class="primary" id="load">Загрузить в очередь</button></div>'+
      '<p class="hint">Сначала выберите действие над товарами в верхней панели. При ошибке отдельного товара скрипт фиксирует причину в журнале и продолжает работу. Очереди сайтов независимы.</p>'+
      '<div class="setting"><div>Пауза между товарами<small>Секунды, 1–30</small></div>'+
      '<input id="delay" type="number" min="1" max="30" value="2"></div>'+
      '<p class="hint">Отсутствующая ссылка или ошибка сохранения не останавливают очередь. Если сессия WordPress истекла, скрипт ставится на паузу до повторного входа.</p></div>'+
      '<div class="view help hidden" id="help">'+
      '<h3>Выбор действия и импорт</h3><ol>'+
      '<li>Подготовьте Excel с одним столбцом product_link, содержащим ссылки редактора товара на этом сайте.</li>'+
      '<li>Загрузите Excel. Вверху панели выберите «Удалить товар (Move to Trash)» или «Сделать приватным».</li>'+
      '<li>Нажмите «Тест: 1 товар» для пробного применения выбранного действия к одному товару.</li>'+
      '<li>После теста нажмите «Начать» для обработки остальных ссылок.</li>'+
      '</ol><h3>Режим: Удалить товар</h3><p>Открывается product_link → Move to Trash → перенаправление WordPress → запись результата → следующая ссылка. Без дополнительных запросов к карточке и страницам корзины.</p>'+
      '<h3>Режим: Сделать приватным</h3><p>Открывается product_link → Edit visibility → Private → OK → Update → серверное подтверждение статуса private. Если статус уже private, обновление пропускается.</p>'+
      '<h3>Ошибки и результаты</h3><p>Предупреждения WPML о переводе не считаются ошибкой Trash. Успех подтверждается перенаправлением или уведомлением WordPress. Если подтверждения нет — «Не подтверждено», а не ложная «Ошибка удаления». Без дополнительных запросов к корзине.</p>'+
      '<p>При истечении авторизации скрипт останавливает обработку до входа. Повторный Trash/Update после перезагрузки не выполняется.</p>'+
      '<p>Выбранное действие закрепляется за очередью после запуска. Для другого действия создайте новую очередь из Excel.</p></div>'+
      '<footer class="foot"><span id="footer">Очередь сохраняется локально в Tampermonkey</span>'+
      '<button id="prev" title="Предыдущая страница">←</button><span id="pages">1 / 1</span>'+
      '<button id="next" title="Следующая страница">→</button></footer></div></section>';
    panel=$('#panel');
    const layout=GM_getValue(LAYOUT,null);
    if (layout && typeof layout==='object') {
      for (const p of ['left','top','width','height']) {
        if (Number.isFinite(layout[p])) panel.style[p]=layout[p]+'px';
      }
      if (Number.isFinite(layout.left)) panel.style.right='auto';
      if (Number.isFinite(layout.top)) panel.style.bottom='auto';
    }
    $('#min').onclick=()=>{
      panel.classList.add('hidden');$('#pill').classList.remove('hidden');
      GM_setValue(PREFIX+':minimized',true);
    };
    const show=()=>{
      panel.classList.remove('hidden');$('#pill').classList.add('hidden');
      GM_setValue(PREFIX+':minimized',false);
      render();
    };
    $('#pill').onclick=show;
    if (typeof GM_registerMenuCommand==='function')GM_registerMenuCommand('Открыть Product Trash Manager',show);
    $('#full').onclick=()=>{
      expanded=!expanded;panel.classList.toggle('full',expanded);
      if (!expanded) {panel.style.left='';panel.style.top='';panel.style.right='22px';panel.style.bottom='20px';}
    };
    $('#drag').addEventListener('pointerdown',e=>{
      if (e.target.closest('button')||expanded||e.button!==0)return;
      const r=panel.getBoundingClientRect();
      dragging={dx:e.clientX-r.left,dy:e.clientY-r.top};
      e.currentTarget.setPointerCapture(e.pointerId);
    });
    $('#drag').addEventListener('pointermove',e=>{
      if(!dragging)return;
      panel.style.left=Math.max(0,Math.min(innerWidth-panel.offsetWidth,e.clientX-dragging.dx))+'px';
      panel.style.top=Math.max(0,Math.min(innerHeight-panel.offsetHeight,e.clientY-dragging.dy))+'px';
      panel.style.right='auto';panel.style.bottom='auto';
    });
    $('#drag').addEventListener('pointerup',()=>dragging=null);
    $('#drag').addEventListener('pointercancel',()=>dragging=null);
    const layoutSave=()=>{
      if(expanded||panel.classList.contains('hidden'))return;
      const r=panel.getBoundingClientRect();
      GM_setValue(LAYOUT,{left:r.left,top:r.top,width:r.width,height:r.height});
    };
    $('#drag').addEventListener('pointerup',layoutSave);
    if (typeof ResizeObserver!=='undefined') {
      let timeout;
      new ResizeObserver(()=>{clearTimeout(timeout);timeout=setTimeout(layoutSave,600);}).observe(panel);
    }
    $$('[data-tab]').forEach(b=>b.onclick=()=>setTab(b.dataset.tab));
    $('#operation').value=state.action||'';
    $('#operation').onchange=e=>{
      if(state.running||busy||state.started){
        e.target.value=state.action||'';
        toast('Для смены действия создайте новую очередь. Во время выполнения менять действие запрещено.');
        return;
      }
      if(!['','trash','private'].includes(e.target.value))return;
      state.action=e.target.value;
      log('info','Оператор выбрал действие: '+actionName(state.action)+'.');
    };
    $('#start').onclick=()=>start(false);
    $('#one').onclick=()=>start(true);
    $('#pause').onclick=()=>pause();
    $('#stop').onclick=stop;
    $('#report').onclick=report;
    $('#downloadLog').onclick=downloadLog;
    $('#reset').onclick=reset;
    $('#prev').onclick=()=>{page=Math.max(0,page-1);renderRows();};
    $('#next').onclick=()=>{page++;renderRows();};
    $('#search').oninput=e=>{search=e.target.value;page=0;renderRows();};
    $('#filter').onchange=e=>{filter=e.target.value;page=0;renderRows();};
    $('#file').onchange=e=>readExcel(e.target.files?.[0]);
    $('#sheet').onchange=fillMapping;
    
    $('#load').onclick=importQueue;
    const drop=$('#drop');
    drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('over');});
    drop.addEventListener('dragleave',()=>drop.classList.remove('over'));
    drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('over');readExcel(e.dataTransfer.files[0]);});
    $('#delay').value=state.settings.delay;
    $('#delay').onchange=e=>{
      if(!canEdit()){e.target.value=state.settings.delay;return;}
      state.settings.delay=Math.min(30,Math.max(1,Number(e.target.value)||2));
      e.target.value=state.settings.delay;persist();
    };
    if (GM_getValue(PREFIX+':minimized',false)) {
      panel.classList.add('hidden');$('#pill').classList.remove('hidden');
    }
    if(!state.rows.length)setTab('import');
    else render();
  }
  function setTab(tab) {
    activeTab=tab;page=0;
    for(const name of ['queue','not_found','failed','unconfirmed','log','import','help'])
      $('#'+name).classList.toggle('hidden',name!==tab);
    $$('[data-tab]').forEach(b=>b.classList.toggle('on',b.dataset.tab===tab));
    render();
  }
  function subset() {
    if (activeTab==='not_found') return state.rows.filter(x=>x.status==='not_found');
    if (activeTab==='failed') return state.rows.filter(x=>x.status==='failed');
    if (activeTab==='unconfirmed') return state.rows.filter(x=>x.status==='unconfirmed');
    return state.rows.filter(x=>{
      const matches=filter==='all'||(filter==='pending' ? ['pending','opening','deleting','saving','verifying'].includes(x.status) : x.status===filter);
      return matches && normalized(x.id+' '+x.productLink+' '+x.foundName).includes(normalized(search));
    });
  }
  function renderRows() {
    if(!root)return;
    if(!['queue','not_found','failed','unconfirmed'].includes(activeTab))return;
    const rows=subset();
    const pages=Math.max(1,Math.ceil(rows.length/PAGESIZE));
    page=Math.min(page,pages-1);
    $('#pages').textContent=(page+1)+' / '+pages;
    $('#prev').disabled=page===0;$('#next').disabled=page>=pages-1;
    const target=activeTab==='queue' ? $('#rows') : activeTab==='not_found' ? $('#missingRows') : activeTab==='unconfirmed' ? $('#unconfirmedRows') : $('#failedRows');
    if(!rows.length){
      target.innerHTML='<div class="empty"><b>Список пуст</b>'+
        (activeTab==='queue'?'Импортируйте Excel или измените фильтр.':'Здесь пока нет записей.')+'</div>';
      return;
    }
    target.innerHTML='<table><thead><tr><th style="width:50px">Строка</th><th>Игра / ID</th>'+
      '<th style="width:112px">Статус</th></tr></thead><tbody>'+
      rows.slice(page*PAGESIZE,page*PAGESIZE+PAGESIZE).map(r=>
        '<tr class="'+(current()===r?'currentrow':'')+'"><td>'+r.excelRow+'</td>'+
        '<td><span class="title" title="'+esc(r.foundName||r.productLink||'—')+'">'+esc(r.foundName||('Товар ID '+(r.id||'—')))+
        '</span><small>'+esc(r.productLink||'')+'</small><small>'+esc(r.message||'')+'</small></td>'+
        '<td><span class="badge '+(statuses[r.status]||'neutral')+'">'+esc(rowLabel(r))+
        '</span></td></tr>'
      ).join('')+'</tbody></table>';
  }
  function renderLogs() {
    $('#logs').innerHTML=state.logs.slice().reverse().map(x=>
      '<div class="logline '+esc(x.level)+'"><time>'+esc(new Date(x.time).toLocaleTimeString('ru-RU'))+
      '</time><span>'+esc(x.id?'ID '+x.id:'Система')+'</span><span>'+esc(x.message)+'</span></div>'
    ).join('')||'<div class="empty">Журнал пока пуст</div>';
  }
  function render() {
    if(!root)return;
    const count=s=>state.rows.filter(r=>r.status===s).length;
    const done=count('done'),missing=count('not_found'),error=count('failed'),unconfirmed=count('unconfirmed');
    const pending=count('pending')+count('opening')+count('deleting')+count('saving')+count('verifying');
    $('#total').textContent=state.rows.length;
    $('#done').textContent=done;
    $('#doneLabel').textContent=state.action==='private'?'Приватных':'Выполнено';
    $('#missing').textContent=missing;
    $('#errors').textContent=error;
    $('#unconfirmed').textContent=unconfirmed;
    $('#bar').style.width=(state.rows.length?
      Math.round(100*(done+missing+error+unconfirmed+count('skipped'))/state.rows.length):0)+'%';
    const row=current();
    $('#current').textContent=state.running?
      'Обработка ссылки ID '+(row?.id||'—')+' · '+(labels[row?.status]||'Открытие'):
      state.paused?'Пауза · нажмите «Начать» для продолжения':
      state.rows.length?(state.filename+' · ожидают: '+pending):
      'Загрузите Excel с product_link';
    $('#dot').classList.toggle('run',state.running);
    const hasPending=nextPending()>=0||!!(row&&!TERMINAL.has(row.status));
    const lease=GM_getValue(LEASE,null);
    const lockedElsewhere=!!(lease && lease.until>Date.now() && lease.id!==tabId);
    const activeHere=state.running && state.owner===tabId;
    const recoverable=state.running && state.owner!==tabId && !lockedElsewhere;
    $('#start').textContent=recoverable?'▶ Возобновить':state.paused?'▶ Продолжить':'▶ Начать';
    $('#start').disabled=activeHere||lockedElsewhere||busy||!hasPending;
    $('#one').disabled=activeHere||lockedElsewhere||busy||!hasPending;
    $('#operation').value=state.action||'';
    $('#operation').disabled=activeHere||lockedElsewhere||busy||!!state.started;
    if(!['trash','private'].includes(state.action)){
      $('#start').disabled=true;$('#one').disabled=true;
    }
    $('#pause').disabled=!activeHere;
    $('#stop').disabled=!activeHere;
    $('#report').disabled=!state.rows.length;
    $('#load').disabled=state.running||busy||!book||linkColumn()<0;
    $('#reset').disabled=state.running||busy;
    $('#footer').textContent=state.updated?'Сохранено · '+new Date(state.updated).toLocaleTimeString('ru-RU'):
      'Очередь сохраняется автоматически';
    $$('[data-tab]').forEach(b=>{
      const n=b.dataset.tab;
      if(n==='not_found')b.textContent='Не найдены ('+missing+')';
      if(n==='failed')b.textContent='Ошибки обработки ('+error+')';
      if(n==='unconfirmed')b.textContent='Не подтверждено ('+unconfirmed+')';
    });
    if(['queue','not_found','failed','unconfirmed'].includes(activeTab))renderRows();
    if(activeTab==='log')renderLogs();
    $('.foot #prev').classList.toggle('hidden',!['queue','not_found','failed','unconfirmed'].includes(activeTab));
    $('.foot #next').classList.toggle('hidden',!['queue','not_found','failed','unconfirmed'].includes(activeTab));
    $('.foot #pages').classList.toggle('hidden',!['queue','not_found','failed','unconfirmed'].includes(activeTab));
  }
  // Do not resume unattended after a WordPress login redirect.
  if(location.pathname.endsWith('/wp-login.php')) {
    if(state.running) {
      state.running=false;state.paused=true;
      log('error','Сессия WordPress закончилась. Войдите снова и нажмите «Продолжить» вручную.');
      release();
    }
    return;
  }
  mount();
  if(state.running && state.owner === tabId) {
    if(claim()) {
      log('info','Возобновление после перехода или перезагрузки вкладки.');
      schedule(250);
    } else {
      state.running=false;state.paused=true;persist();
    }
  }
})();