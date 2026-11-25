// ==UserScript==
// @name         steam-site-editor
// @namespace    https://steambuy.pro/
// @version      2.6.1
// @description  WooCommerce Excel: Trash / Private / Price, быстрый режим 1-10 параллельных запросов без загрузки картинок, журнал.
// @homepageURL  https://github.com/blackowl0192/script_folder/blob/main/steam-site-editor.user.js
// @updateURL    https://raw.githubusercontent.com/blackowl0192/script_folder/main/steam-site-editor.user.js
// @downloadURL  https://raw.githubusercontent.com/blackowl0192/script_folder/main/steam-site-editor.user.js
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
  const VERSION = '2.6.1';
  const DOMAIN = location.hostname;
  const PREFIX = 'sb-product-trash-v3:' + DOMAIN;
  const OLD_LEASE = 'sb-product-trash-v2:' + DOMAIN + ':lease';
  const KEY = PREFIX + ':state';
  const LEASE = PREFIX + ':lease';
  const LAYOUT = PREFIX + ':layout';
  const TAB = PREFIX + ':tab-id';
  const PRICE_KEY = PREFIX + ':prices:state';
  const PRICE_LEASE = PREFIX + ':prices:lease';
  const LIST = location.origin + '/wp-admin/edit.php?post_type=product';
  const MAX_LOGS = 1500;
  const PAGESIZE = 25;
  const OK = 'done';
  const TERMINAL = new Set(['done', 'unchanged', 'not_found', 'failed', 'unconfirmed', 'skipped']);
  const $ = (s, p = root) => p.querySelector(s);
  const $$ = (s, p = root) => Array.from(p.querySelectorAll(s));
  const clean = x => String(x == null ? '' : x).trim();
  const idValid = x => /^[1-9]\d*$/.test(clean(x));
  const esc = x => clean(x).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const normalized = x => clean(x).normalize('NFKC').toLowerCase().replace(/\s+/g, ' ');
  const now = () => new Date().toISOString();
  const tabId = sessionStorage.getItem(TAB) || (Date.now().toString(36) + Math.random().toString(36).slice(2));
  sessionStorage.setItem(TAB, tabId);
  const PAGE_INSTANCE = Date.now().toString(36)+'-'+Math.random().toString(36).slice(2);
  const defaults = () => ({
    v: 3, schema: 2, action: '', filename: '', sheet: '', rows: [], logs: [], index: -1,
    phase: 'idle', running: false, paused: false, mode: 'batch', runCount: 0,
    owner: '', pending: null,
    settings: { delay: 1, mode: 'compatible', workers: 5 }, updated: '', started: ''
  });
  let state = GM_getValue(KEY, null) || defaults();
  if (state.v !== 3 || !Array.isArray(state.rows)) state = defaults();
  state.settings = Object.assign({delay: 1,mode:'compatible',workers:5}, state.settings || {});
  if(!['compatible','fast'].includes(state.settings.mode))state.settings.mode='compatible';
  state.settings.workers=Math.max(1,Math.min(10,Number(state.settings.workers)||5));
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
  let root, panel, book = null, fileName = '', excelColumn = -1, priceColumn = -1;
  const savedTab=sessionStorage.getItem(PREFIX+':ui:tab');
  let activeTab=['queue','log','import','help'].includes(savedTab)?savedTab:'queue';
  let filter=sessionStorage.getItem(PREFIX+':ui:filter')||'all',search='',page=0;
  let logSearch=sessionStorage.getItem(PREFIX+':ui:logSearch')||'';
  let logLevel=sessionStorage.getItem(PREFIX+':ui:logLevel')||'all';
  let logLimit=Number(sessionStorage.getItem(PREFIX+':ui:logLimit'))||100;
  let busy = false, moving = false, timer = null, heartbeat = null, expanded = false, dragging = null;
  const fastActive=new Map();
  let fastSaveTimer=null,fastPaintTimer=null,fastLimit=5,fastErrors=0,fastSerial=0,fastPausedNotice=false;
  const labels = {pending:'Ожидает', opening:'Открытие', deleting:'Удаляется',
    saving:'Сохранение', verifying:'Проверка',done:'Выполнено',
    not_found:'Не найден', failed:'Ошибка обработки', unconfirmed:'Не подтверждено', skipped:'Пропущен'};
  const ACTION_LABELS = {trash:'Переместить в Trash', private:'Сделать приватным', price:'Редактировать Regular price'};
  const actionName = action => ACTION_LABELS[action] || 'Не выбрано';
  const rowLabel = r => r.status==='done'
    ? (r.action==='private'?'Приватный':r.action==='price'?'Цена изменена':'Удалён')
    : r.status==='unchanged'?'Без изменений':(labels[r.status]||r.status);
  const statuses = {done:'done',unchanged:'neutral', failed:'bad', not_found:'bad', unconfirmed:'uncertain', skipped:'neutral',
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
    const former=GM_getValue(PRICE_LEASE,null);
    if(former&&former.until>Date.now()){
      toast('Старая версия редактора цен ещё выполняется. Остановите её и обновите вкладку.');
      return false;
    }
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
      if ((!state.running&&!fastActive.size) || !ownLease()) { clearInterval(heartbeat); return; }
      GM_setValue(LEASE, {id:tabId,until:Date.now() + 25000});
    },4000);
  }
  function release() {
    if(fastActive.size)return; // In-flight POST/Trash must finish before releasing the lease.
    clearInterval(heartbeat);
    if (ownLease()) GM_setValue(LEASE, null);
  }
  function readState() {
    const fresh = GM_getValue(KEY, null);
    if (fresh?.v === 3 && Array.isArray(fresh.rows)) state = fresh;
  }
  function canEdit() {
    const lease = GM_getValue(LEASE,null);
    if (busy || fastActive.size || (state.running && (state.owner===tabId || (lease && lease.until>Date.now())))) {
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
    log('info', reason+(fastActive.size?' Активные запросы ('+fastActive.size+') завершатся без запуска новых.':''), current());
    release();
  }
  function stop() { pause('Остановлено пользователем. Сохранённые результаты не потеряны.'); }
  function finalize(row, status, message) {
    if (!row) return;
    if(!row.action && ['trash','private','price'].includes(state.action))row.action=state.action;
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
      schedule(Math.max(500,Number(state.settings.delay) * 1000));
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
    // This is a freshly loaded editor page: radio visibility is authoritative
    // even when WPML keeps the hidden #post_status field at "publish".
    const before=privateVisibility(document);
    if(before.id===row.id && (before.private || status==='private'))
      return finalize(row,'done','Товар уже приватный (Visibility: Private). Update не требуется.');
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
    state.pending={id:row.id,action:'private',pageInstance:PAGE_INSTANCE,at:now()};
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
  // Classic Editor / WPML: post_status can remain "publish" for a Private
  // product. Prefer the saved Visibility controls over the hidden status field.
  function privateVisibility(doc) {
    const get=selector=>doc.querySelector(selector);
    const id=clean(get('#post_ID,input[name="post_ID"]')?.value)||
      (clean(doc.body?.className).match(/\bpostid-(\d+)\b/)||[])[1]||'';
    const type=clean(get('#post_type,input[name="post_type"]')?.value)||
      (doc.body?.classList?.contains('post-type-product')?'product':'');
    const status=clean(get('#post_status,input[name="post_status"]')?.value);
    const originalStatus=clean(get('#original_post_status,input[name="original_post_status"]')?.value);
    const radio=get('#visibility-radio-private');
    const publicRadio=get('#visibility-radio-public');
    const passwordRadio=get('#visibility-radio-password');
    const display=clean(get('#post-visibility-display')?.textContent);
    const radioPrivate=radio?.checked===true;
    const publicSelected=publicRadio?.checked===true || passwordRadio?.checked===true;
    const displayPrivate=/(?:\bprivate\b|приватн|личн(?:ое|ая)|закрыт)/i.test(display);
    const statusPrivate=status==='private' || originalStatus==='private';
    const isPrivate=!publicSelected && (statusPrivate || radioPrivate ||
      (displayPrivate && !radio));
    return {id,type,status,originalStatus,radioPrivate,display,isPrivate};
  }
  async function verifyPrivate(row) {
    if(busy)return;
    busy=true;
    row.status='verifying';
    row.message='Быстрая проверка Visibility: Private после Update.';
    persist();
    try {
      // A new userscript instance means WordPress has completed navigation:
      // inspect the already loaded editor instead of fetching the same page again.
      const fresh=!!(state.pending?.pageInstance &&
        state.pending.pageInstance!==PAGE_INSTANCE);
      const sameEditor=isProductEditor(row);
      const usePage=fresh && sameEditor;
      const doc=usePage?document:await fetchAdmin(row.productLink);
      const source=usePage?'загруженная после Update страница':'один запрос к карточке';
      if(doc.querySelector('#loginform'))throw new Error('Требуется войти в WordPress.');
      const info=privateVisibility(doc);
      if(info.id!==row.id || info.type!=='product')
        throw new Error('Проверка ID/типа не пройдена: ожидался product '+row.id+
          ', получены ID='+(info.id||'—')+', тип='+(info.type||'—')+'.');
      log('info','Private: '+source+'; ID '+row.id+
        '; visibility-radio-private='+(info.radioPrivate?'checked':'unchecked')+
        '; display='+(info.display||'—')+
        '; post_status='+(info.status||'—')+'.',row);
      if(info.isPrivate)
        return finalize(row,'done','Visibility: Private подтверждён после Update. ID '+row.id+'.');
      // Never resend Update. If the saved page does not confirm Private,
      // classify as unconfirmed instead of misreporting WPML "publish" as failure.
      return finalize(row,'unconfirmed',
        'После Update приватность не подтверждена: Visibility='+(info.display||'—')+
        ', post_status='+(info.status||'—')+'. Повторный Update не выполняется.');
    }catch(e) {
      const message=String(e?.message||e);
      finalize(row,/сессия|войти|авторизац/i.test(message)?'failed':'unconfirmed',
        'Проверка приватности ID '+row.id+': '+message+
        '. Повторный Update не выполняется.');
      if(/сессия|войти|авторизац/i.test(message)){
        state.running=false;state.paused=true;persist();release();
      }
    }finally {
      busy=false;
      if(state.running)schedule(Math.max(500,Number(state.settings.delay)*1000));
    }
  }
  // Common action: set WooCommerce Regular price from the Excel row.
  function applyPrice(row) {
    const input=document.querySelector('#_regular_price');
    const sale=document.querySelector('#_sale_price');
    const update=document.querySelector('#publish');
    if(!input || !sale || !update)
      return finalize(row,'failed','Не найдено поле Regular price, Sale price или кнопка Update.');
    if(input.name!=='_regular_price'||sale.name!=='_sale_price')
      return finalize(row,'failed','Поля цен не совпадают с ожидаемыми WooCommerce _regular_price / _sale_price.');
    if(input.disabled||sale.disabled||update.disabled)
      return finalize(row,'failed','Поле Regular price, Sale price или Update недоступно (disabled).');
    const label=clean(update.value||update.textContent);
    if(!/^(?:update|обновить)$/i.test(label))
      return finalize(row,'failed','Ожидалась кнопка Update, но найдена «'+(label||'неизвестно')+'».');
    if(!Number.isSafeInteger(row.cents)||!row.price)
      return finalize(row,'failed','Для товара ID '+row.id+' отсутствует корректная цена из Excel.');
    row.before=clean(input.value);
    row.saleBefore=clean(sale.value);
    let regularMatches=false;
    try{regularMatches=parseRegularPrice(row.before).cents===row.cents;}
    catch{/* An empty or unusual old price can still be replaced. */}
    if(regularMatches && !row.saleBefore){
      row.after=row.before;row.saleAfter='';
      return finalize(row,'unchanged','Regular price уже '+row.price+
        ', Sale price пустая. Update не требуется.');
    }
    log('info','ID '+row.id+': Regular price '+(row.before||'(пусто)')+
      ' → '+row.price+'; Sale price '+(row.saleBefore||'(пусто)')+' → (пусто).',row);
    try {
      const setter=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')?.set;
      if(setter)setter.call(input,row.price);
      else input.value=row.price;
      input.dispatchEvent(new Event('input',{bubbles:true}));
      input.dispatchEvent(new Event('change',{bubbles:true}));
      if(setter)setter.call(sale,'');
      else sale.value='';
      sale.dispatchEvent(new Event('input',{bubbles:true}));
      sale.dispatchEvent(new Event('change',{bubbles:true}));
      if(parseRegularPrice(input.value).cents!==row.cents)
        throw new Error('Regular price после ввода содержит '+clean(input.value)+', ожидалось '+row.price+'.');
      if(clean(sale.value)!=='')
        throw new Error('Sale price не очистилась перед Update: '+clean(sale.value)+'.');
    }catch(e){
      return finalize(row,'failed','Не удалось ввести цену: '+String(e?.message||e));
    }
    // Save checkpoint BEFORE Update, do not automatically submit twice.
    state.phase='verify-price';
    state.pending={id:row.id,action:'price',pageInstance:PAGE_INSTANCE,at:now()};
    row.status='saving';row.message='Regular price установлена, Sale price очищена; нажимаем Update.';
    log('info','ID '+row.id+': нажимаем Update. Regular '+row.price+'; Sale price пустая.',row);
    try{update.click();}
    catch(e){
      return finalize(row,'unconfirmed','Update вызвал ошибку: '+String(e?.message||e)+
        '. Повторно Update не нажимаем.');
    }
    setTimeout(()=>{
      if(state.running&&ownLease()&&current()?.id===row.id&&
        state.phase==='verify-price'&&!busy&&!moving)verifyPrice(row);
    },4000);
  }
  async function verifyPrice(row) {
    if(busy)return;
    busy=true;row.status='verifying';row.message='Проверяем сохранённую цену после Update.';
    persist();
    try{
      const usePage=!!(state.pending?.pageInstance &&
        state.pending.pageInstance!==PAGE_INSTANCE && isProductEditor(row));
      const doc=usePage?document:await fetchAdmin(row.productLink);
      if(doc.querySelector('#loginform'))throw new Error('Требуется войти в WordPress.');
      const id=clean(doc.querySelector('#post_ID,input[name="post_ID"]')?.value);
      const type=clean(doc.querySelector('#post_type,input[name="post_type"]')?.value)||
        (doc.body?.classList.contains('post-type-product')?'product':'');
      if(id!==row.id||type!=='product')
        throw new Error('После Update открыта другая карточка (ID '+id+', тип '+type+').');
      const input=doc.querySelector('#_regular_price');
      const sale=doc.querySelector('#_sale_price');
      if(!input||!sale)throw new Error('Не найдены сохранённые поля Regular price / Sale price.');
      row.after=clean(input.value);
      row.saleAfter=clean(sale.value);
      let saved;
      try{saved=parseRegularPrice(row.after);}catch{saved=null;}
      log('info','ID '+row.id+': после Update Regular price='+(row.after||'(пусто)')+
        ' (ожидалась '+row.price+'); Sale price='+(row.saleAfter||'(пусто)')+
        ' (ожидалась пустая); источник: '+(usePage?'загруженный редактор':'один контрольный запрос')+'.',row);
      if(saved&&saved.cents===row.cents&&row.saleAfter==='')
        return finalize(row,'done','Regular price: '+(row.before||'(пусто)')+
          ' → '+row.after+'; Sale price очищена'+
          (row.saleBefore?' (была '+row.saleBefore+')':'')+'.');
      return finalize(row,'unconfirmed','После Update цены не соответствуют заданию: Regular price ожидалась '+
        row.price+', получено '+(row.after||'(пусто)')+
        '; Sale price должна быть пустой, получено '+(row.saleAfter||'(пусто)')+
        '. Повторный Update не выполняется.');
    }catch(e){
      const msg=String(e?.message||e);
      finalize(row,'unconfirmed','Не удалось подтвердить цену после Update: '+msg+
        '. Повторного Update не будет.');
      if(/сессия|войти|авторизац/i.test(msg)){
        state.running=false;state.paused=true;persist();release();
      }
    }finally{
      busy=false;
      if(state.running)schedule(Math.max(500,Number(state.settings.delay)*1000));
    }
  }
  function openProduct(row) {
    // Action is fixed for the row once processing starts, regardless of UI changes.
    if(!['trash','private','price'].includes(row.action)) row.action=state.action;
    if(!['trash','private','price'].includes(row.action))return failFatal('Выберите действие перед запуском.',row);
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
    if(row.action==='price')return applyPrice(row);
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
      if(state.running)schedule(Math.max(500,Number(state.settings.delay)*1000));
    }
  }
  function runStep() {
    if(state.settings.mode==='fast')return fastPump();
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
    if(state.phase==='verify-price'||(state.phase==='saving'&&row.action==='price')){
      log('info','Возобновление: проверяем сохранение цены без повторного Update.',row);
      return verifyPrice(row);
    }
    if(['verify-private','saving'].includes(state.phase)){
      log('warn','Возобновление: проверяем сохранённую приватность без повторного Update.',row);
      return verifyPrivate(row);
    }
    state.phase='open';openProduct(row);
  }
  function start(one=false) {
    if(busy||moving||fastActive.size)return toast('Дождитесь завершения текущих запросов.');
    if(state.settings.mode==='fast')return fastStart(one);
    readState();
    if(!state.rows.length){log('error','Запуск невозможен: очередь product_link пуста.');return toast('Сначала загрузите Excel со столбцом product_link.');}
    const pendingAction=state.pending?.action || (state.pending && ['verify','verifying','deleting'].includes(state.phase)?'trash':'');
    const selected=clean(state.action);
    if(!['trash','private','price'].includes(selected)&&!pendingAction){log('error','Запуск невозможен: действие не выбрано.');return toast('Выберите действие с товарами.');}
    if(state.rows.some(r=>r.status==='pending'&&r.action&&r.action!==selected))
      return toast('Вы изменили действие после импорта. Загрузите Excel в очередь заново.');
    if(selected==='price' && state.rows.some(r=>r.status==='pending' &&
      (!r.price || !Number.isSafeInteger(r.cents))))
      return toast('В очереди нет подготовленных цен. Выберите «Редактировать цену» и импортируйте Excel заново.');
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
    const previousSettings={...state.settings};
    state=defaults();state.settings=Object.assign(state.settings,previousSettings);
    book=null;fileName='';excelColumn=-1;priceColumn=-1;
    $('#file').value='';
    $('#filename').textContent='Файл не выбран';
    $('#mapping').classList.add('hidden');
    persist();
    setTab('import');
    toast('Очередь и журнал очищены.');
  }
  // Preserve queues imported through the former standalone Prices tab (v2.4).
  function migrateLegacyPrices(){
    if(!canEdit())return;
    const lease=GM_getValue(PRICE_LEASE,null);
    if(lease&&lease.until>Date.now())
      return toast('Старая очередь цен ещё запущена в другой вкладке. Остановите её перед переносом.');
    const old=GM_getValue(PRICE_KEY,null);
    if(!old||!Array.isArray(old.rows)||!old.rows.length)
      return toast('Сохранённая очередь цен версии 2.4 не найдена.');
    if(!confirm('Перенести '+old.rows.length+' записей из старой очереди цен?\n'+
      'Текущая очередь Product Manager будет заменена, журнал сохранится в Excel перед переносом только при ручном экспорте.'))
      return;
    const delay=Number(old.delay)||0.5;
    const rows=old.rows.map(r=>{
      const item={...r,action:'price'};
      if(item.status==='updated')item.status='done';
      if(!TERMINAL.has(item.status)&&item.status!=='saving'&&item.status!=='verifying')
        item.status='pending';
      return item;
    });
    const previous=state;
    const migrated=defaults();
    migrated.action='price';migrated.rows=rows;
    migrated.filename=old.filename||'prices-v2.4';
    migrated.sheet=old.sheet||'';
    migrated.logs=(old.logs||[]).slice(-MAX_LOGS);
    migrated.settings.delay=Math.min(30,Math.max(0.5,delay));
    migrated.running=false;migrated.paused=true;
    migrated.mode='batch';migrated.owner='';migrated.started=now();
    const idx=Number.isInteger(old.index)?old.index:-1;
    if(idx>=0&&idx<rows.length&&!TERMINAL.has(rows[idx].status)){
      migrated.index=idx;
      if(old.phase==='verify-price'&&old.pending?.id===rows[idx].id){
        migrated.phase='verify-price';migrated.pending={...old.pending,action:'price'};
      }else if(['saving','verifying'].includes(rows[idx].status)){
        rows[idx].status='unconfirmed';
        rows[idx].message='Перенос старой очереди: сохранение могло завершиться. Не повторяем Update без проверки.';
        migrated.index=-1;
      }else{
        migrated.phase='open';
      }
    }
    if(previous.rows.length)
      migrated.logs.push({time:now(),level:'warn',id:'',row:'',
        message:'Ветка очереди Product Manager была заменена при подтверждённом переносе старых цен.'});
    state=migrated;
    old.running=false;old.paused=true;
    GM_setValue(PRICE_KEY,old);
    log('info','Старая очередь цен перенесена в общий режим действий: '+rows.length+' строк. Операция поставлена на паузу.');
    setTab('queue');
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
    $('#priceColumnWrap').classList.toggle('hidden',state.action!=='price');
    $('#priceColumn').classList.toggle('hidden',state.action!=='price');
    if(!book){
      excelColumn=-1;priceColumn=-1;
      $('#load').disabled=true;
      return;
    }
    const rows=sheetRows(),head=rows[0]||[];
    const col=head.findIndex(x=>normalized(x).replace(/[\s\-]+/g,'_')==='product_link');
    excelColumn=col;
    $('#linkColumn').textContent=col<0?'Колонка product_link не найдена.':
      'Найдена колонка product_link (№ '+(col+1)+').';
    const select=$('#priceColumn'),previous=select.value;
    select.innerHTML='<option value="">Выберите столбец цены...</option>'+
      head.map((h,i)=>'<option value="'+i+'">'+esc(h||'Колонка '+(i+1))+'</option>').join('');
    const normalizedHeaders=head.map(x=>normalized(x).replace(/[\s\-]+/g,'_'));
    const names=['price','_regular_price','regular_price','цена','цена_игры','стоимость','цена_товара','price_€','regular_price_(€)'];
    const suggested=normalizedHeaders.findIndex(h=>names.includes(h));
    let selected=previous!=='' && head[Number(previous)]!==undefined ? Number(previous) :
      suggested>=0?suggested:(head.length===2&&col>=0?1-col:-1);
    if(selected===col)selected=-1;
    select.value=selected>=0?String(selected):'';
    priceColumn=selected;
    const isPrice=state.action==='price';
    $('#load').disabled=col<0||(isPrice&&selected<0);
    $('#preview').textContent=col<0?'В Excel требуется заголовок product_link.':
      rows.slice(1,6).map((r,i)=>'Строка '+(i+2)+': '+
        (isPrice?'цена '+clean(r[selected])+' · ':'')+clean(linkCell(rows,i+1,col)))
        .filter(Boolean).join('\n')||'Нет ссылок.';
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
    const isPrice=state.action==='price',chosenPrice=$('#priceColumn').value;
    const pc=isPrice && chosenPrice!==''?Number(chosenPrice):-1;
    if(isPrice&&(pc<0||pc===col||!Number.isInteger(pc)))
      return toast('Для редактирования цены укажите отдельный столбец цены.');
    if(!['trash','private','price'].includes(state.action))
      return toast('Сначала выберите действие с товарами.');
    const seen=new Set(),rows=[];
    for(let i=1;i<data.length;i++){
      const raw=clean(linkCell(data,i,col));
      if(!raw && !(isPrice&&clean(data[i]?.[pc])))continue;
      const row={excelRow:i+1,productLink:raw,id:'',name:'',status:'pending',
        message:'',foundName:'',wpPostId:'',finished:'',action:state.action,
        rawPrice:isPrice?clean(data[i]?.[pc]):'',price:'',cents:null,before:'',after:''};
      try{
        const p=parseProductLink(raw);
        row.productLink=p.url;row.id=p.id;
        if(isPrice){
          const result=parseRegularPrice(row.rawPrice);
          row.price=result.value;row.cents=result.cents;
        }
        if(seen.has(row.id))throw new Error('Дубликат product_link с ID '+row.id+'.');
        seen.add(row.id);
      }catch(e){row.status='skipped';row.message=e.message;}
      rows.push(row);
    }
    if(!rows.length)return toast('В product_link нет ссылок.');
    const good=rows.filter(r=>r.status==='pending').length;
    if(!good)return toast('Нет корректных ссылок на редактор товаров этого сайта.');
    if(state.rows.length&&!confirm('Заменить прежнюю очередь для '+DOMAIN+'? Вы можете сначала скачать отчёт.'))return;
    const chosenAction=state.action, chosenSettings={...state.settings};
    state=defaults();state.action=chosenAction;
    state.settings=Object.assign(state.settings,chosenSettings);state.filename=fileName;state.sheet=book.SheetNames[Number($('#sheet').value)];state.rows=rows;
    for(const r of rows){
      if(r.status!=='skipped')continue;
      state.logs.push({time:now(),level:'warn',id:r.id,row:r.excelRow,
        message:'Пропуск ссылки '+r.productLink+': '+r.message});
    }
    if(state.logs.length>MAX_LOGS)state.logs.splice(0,state.logs.length-MAX_LOGS);
    log('info','Импорт '+fileName+': '+rows.length+' ссылок, к обработке '+good+', пропущено '+(rows.length-good)+'. Действие: '+actionName(state.action)+'.');
    filter='all';$('#filter').value='all';
    sessionStorage.setItem(PREFIX+':ui:filter','all');
    setTab('queue');
  }
  function report(){
    if(!state.rows.length)return toast('Очередь пуста.');
    if(typeof XLSX==='undefined')return toast('Библиотека XLSX недоступна.');
    const cols=['Строка Excel','product_link','ID из ссылки','Название на сайте','Действие',
      'Цена из Excel','Нормализованная цена','Regular до','Regular после',
      'Sale до','Sale после','Режим','Поток','Время (с)','Статус','Комментарий','Завершено'];
    const line=r=>[r.excelRow,r.productLink,r.id,r.foundName,actionName(r.action||state.action),
      r.rawPrice||'',r.price||'',r.before||'',r.after||'',
      r.saleBefore||'',r.saleAfter??'',
      r.fastWorker?'Быстрый':'Совместимый',r.fastWorker||'',
      r.elapsedMs!=null?Number((r.elapsedMs/1000).toFixed(2)):'',rowLabel(r),r.message,r.finished];
    const wb=XLSX.utils.book_new();
    const add=(name,rows)=>{
      const ws=XLSX.utils.aoa_to_sheet([cols,...rows.map(line)]);
      ws['!cols']=[{wch:13},{wch:95},{wch:18},{wch:40},{wch:25},{wch:20},{wch:20},{wch:16},{wch:16},{wch:16},{wch:16},{wch:15},{wch:10},{wch:13},{wch:22},{wch:95},{wch:25}];
      XLSX.utils.book_append_sheet(wb,ws,name);
    };
    add('Все товары',state.rows);
    add('Удалено',state.rows.filter(r=>r.status==='done'&&r.action==='trash'));
    add('Цены изменены',state.rows.filter(r=>r.status==='done'&&r.action==='price'));
    add('Без изменений',state.rows.filter(r=>r.status==='unchanged'));
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
    a.href=url;a.download='product_manager_log_'+DOMAIN+'.txt';a.click();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
  }

  // Shared price input normalization for the unified Price action.
  function parseRegularPrice(input){
    const raw=clean(input).replace(/[\u00A0\u202F\u2009]/g,' ');
    if(!raw)throw new Error('Цена отсутствует.');
    let number=raw.replace(/\b(?:EUR|USD|GBP|RUB|UAH|PLN|CHF|CAD|AUD|TRY|KZT|euro|euros)\b/gi,'')
      .replace(/[€$£₽₴₺₸¥₹₪]/g,'').replace(/\s+/g,'');
    if(!/^\d[\d,.]*$/.test(number))throw new Error('Недопустимый формат цены: '+raw.slice(0,90));
    const dots=(number.match(/\./g)||[]).length,commas=(number.match(/,/g)||[]).length;
    let whole=number,fract='';
    if(dots&&commas){
      const pos=Math.max(number.lastIndexOf('.'),number.lastIndexOf(','));
      whole=number.slice(0,pos).replace(/[.,]/g,'');
      fract=number.slice(pos+1);
      if(!/^\d{1,2}$/.test(fract))throw new Error('Ожидается максимум 2 цифры после запятой.');
    }else if(dots||commas){
      const separator=commas?',':'.';
      const pieces=number.split(separator);
      const tail=pieces[pieces.length-1];
      if(pieces.length===2&&tail.length<=2&&tail.length>=1){
        whole=pieces[0];fract=tail;
      }else if(pieces.length>=2&&pieces.slice(1).every(p=>p.length===3)&&
        pieces[0].length>=1&&pieces[0].length<=3){
        whole=pieces.join('');
      }else if(pieces.length>2&&/^\d{1,2}$/.test(tail)&&
        pieces.slice(1,-1).every(p=>p.length===3)){
        whole=pieces.slice(0,-1).join('');fract=tail;
      }else throw new Error('Неоднозначное число: '+raw.slice(0,90));
    }
    if(!/^\d+$/.test(whole)||fract&&!/^\d{1,2}$/.test(fract))
      throw new Error('Некорректная цена.');
    whole=whole.replace(/^0+(?=\d)/,'');
    const cents=Number(whole)*100+Number(fract.padEnd(2,'0')||'0');
    if(!Number.isSafeInteger(cents))throw new Error('Слишком большая цена.');
    return {value:whole+(fract?','+fract.padEnd(2,'0'):''),cents};
  }


  // === Fast mode: bounded same-origin fetch workers; no page navigation or image loading. ===
  function fastSave(critical=false) {
    if(critical){
      clearTimeout(fastSaveTimer);fastSaveTimer=null;
      state.updated=now();GM_setValue(KEY,state);
    }else if(!fastSaveTimer){
      fastSaveTimer=setTimeout(()=>{
        fastSaveTimer=null;state.updated=now();GM_setValue(KEY,state);
      },160);
    }
    if(!fastPaintTimer){
      fastPaintTimer=setTimeout(()=>{
        fastPaintTimer=null;
        try{render();}catch(e){console.error('[Product Manager: fast UI]',e);}
      },250);
    }
  }
  function fastLog(level,message,row,critical=false){
    state.logs.push({time:now(),level,id:row?.id||'',row:row?.excelRow||'',message});
    if(state.logs.length>MAX_LOGS)state.logs.splice(0,state.logs.length-MAX_LOGS);
    fastSave(critical);
  }
  function fastReadyToWrite(row) {
    return state.running && state.owner===tabId && ownLease() &&
      fastActive.has(row.id);
  }
  function fastNoWriteOnPause(row) {
    if(fastReadyToWrite(row))return false;
    row.status='pending';row.fastStage='';row.message='Пауза до сохранения: товар не изменён.';
    fastLog('info','ID '+row.id+': остановлено до отправки изменений.',row,true);
    return true;
  }
  function fastFinish(row,status,message,elapsed=0){
    if(TERMINAL.has(row.status))return;
    row.status=status;row.message=message;row.fastStage='';
    row.finished=now();row.elapsedMs=Math.round(elapsed);
    state.runCount++;
    fastLog(status==='done'||status==='unchanged'?'ok':
      status==='failed'?'error':'warn',
      '['+actionName(row.action)+'] '+message+
      (elapsed?' ('+(elapsed/1000).toFixed(2)+' с)':''),row);
    if(state.mode==='one'&&state.runCount>=1){
      state.running=false;state.paused=true;
      fastLog('info','Тест одного товара завершён. Остальная очередь не затронута.',row,true);
    }
  }
  function fastUrlAllowed(raw){
    const u=new URL(raw,location.href);
    if(u.origin!==location.origin||!u.pathname.startsWith('/wp-admin/')||
      u.username||u.password)throw new Error('Переход за пределы текущего wp-admin заблокирован.');
    return u;
  }
  function fastSanitizeMedia(html){
    const rawBlocks=/(<textarea\b[^>]*>[\s\S]*?<\/textarea\s*>|<script\b[^>]*>[\s\S]*?<\/script\s*>|<style\b[^>]*>[\s\S]*?<\/style\s*>|<template\b[^>]*>[\s\S]*?<\/template\s*>)/gi;
    return String(html).split(rawBlocks).map(part=>{
      if(/^<(?:textarea|script|style|template)\b/i.test(part))return part;
      return part.replace(/<img\b[^>]*>/gi,'')
        .replace(/<source\b[^>]*>/gi,'')
        .replace(/<iframe\b[^>]*>[\s\S]*?<\/iframe\s*>/gi,'')
        .replace(/<video\b[^>]*>[\s\S]*?<\/video\s*>/gi,'');
    }).join('');
  }
  async function fastRequest(url,options={}) {
    const u=fastUrlAllowed(url);
    const controller=new AbortController();
    const timeout=setTimeout(()=>controller.abort(),30000);
    let response;
    try{
      response=await fetch(u.href,{credentials:'same-origin',redirect:'follow',
        cache:'no-store',signal:controller.signal,...options});
      const redirected=new URL(response.url||u.href,location.href);
      if(redirected.origin!==location.origin||redirected.pathname.endsWith('/wp-login.php'))
        throw new Error('Требуется авторизация WordPress (redirect за пределы wp-admin).');
      const final=fastUrlAllowed(redirected.href);
      if(!response.ok){
        const ex=new Error('HTTP '+response.status+' при запросе '+u.pathname);
        ex.httpStatus=response.status;throw ex;
      }
      const mime=response.headers?.get('content-type')||'';
      if(mime&&!/(html|text\/plain|application\/xhtml)/i.test(mime))
        throw new Error('Неожиданный тип ответа: '+mime);
      const html=await response.text();
      // Strip resource elements BEFORE parsing, but DO NOT alter markup in
      // <textarea> (product description), scripts, styles or <template>.
      const doc=new DOMParser().parseFromString(fastSanitizeMedia(html),'text/html');
      if(doc.querySelector('#loginform'))
        throw new Error('Требуется войти в WordPress.');
      const die=doc.querySelector('.wp-die-message');
      if(die)throw new Error('WordPress: '+clean(die.textContent).slice(0,200));
      return {doc,url:final,response};
    }finally{clearTimeout(timeout);}
  }
  function fastEditor(row,doc) {
    const form=doc.querySelector('form#post');
    const id=clean(doc.querySelector('#post_ID,input[name="post_ID"]')?.value);
    const type=clean(doc.querySelector('#post_type,input[name="post_type"]')?.value)||
      (doc.body?.classList.contains('post-type-product')?'product':'');
    if(id!==row.id||type!=='product')
      throw new Error('ID или тип товара не совпадает: ожидался product '+row.id+
        ', получено '+(type||'—')+' '+(id||'—')+'.');
    if(!form)throw new Error('На странице отсутствует форма form#post.');
    const editpost=form.querySelector('input[name="action"][value="editpost"]');
    const nonce=form.querySelector('input[name="_wpnonce"]');
    const button=form.querySelector('#publish');
    const label=clean(button?.value||button?.textContent);
    if(!editpost||!nonce?.value||!button||button.disabled||
      !/^(?:update|обновить)$/i.test(label))
      throw new Error('Не подтверждены безопасная форма editpost, nonce или кнопка Update.');
    const action=fastUrlAllowed(new URL(form.getAttribute('action')||'post.php',row.productLink).href);
    if(action.pathname!=='/wp-admin/post.php')
      throw new Error('Отправка формы вне /wp-admin/post.php запрещена.');
    if(form.getAttribute('method')?.toLowerCase()!=='post')
      throw new Error('Форма не использует метод POST.');
    return {form,action:action.href,button};
  }
  function fastPayload(row,editor){
    const data=new FormData(editor.form);
    if(clean(data.get('post_ID'))!==row.id||
      clean(data.get('action'))!=='editpost'||!clean(data.get('_wpnonce')))
      throw new Error('Поля ID, action или nonce в форме не совпадают.');
    data.set('save',clean(editor.button.value)||'Update');
    return data;
  }
  function fastPrepareSubmission(row){
    row.fastStage='submitted';row.status='saving';
    row.message='Запрос на сохранение отправляется; повторно не отправлять.';
    fastLog('info','ID '+row.id+': этап сохранения, установлен контроль от повторного Update/Trash.',row,true);
  }
  function fastResultDoc(row,doc,kind){
    const id=clean(doc.querySelector('#post_ID,input[name="post_ID"]')?.value);
    const type=clean(doc.querySelector('#post_type,input[name="post_type"]')?.value)||
      (doc.body?.classList.contains('post-type-product')?'product':'');
    if(id!==row.id||type!=='product')return {confirmed:false,reason:'нет редактора с ожидаемым ID'};
    if(kind==='price'){
      const regular=doc.querySelector('#_regular_price'),sale=doc.querySelector('#_sale_price');
      row.after=clean(regular?.value);
      row.saleAfter=clean(sale?.value);
      if(!regular||!sale)
        return {confirmed:false,reason:'не найдены сохранённые поля Regular price / Sale price'};
      let regularMatches=false;
      try{regularMatches=parseRegularPrice(row.after).cents===row.cents;}
      catch{/* field is empty or malformed */}
      if(regularMatches&&row.saleAfter==='')
        return {confirmed:true,detail:'Regular price: '+row.after+'; Sale price пустая'};
      return {confirmed:false,reason:'Regular price ожидалась '+row.price+', получено '+(row.after||'пусто')+
        '; Sale price должна быть пустой, получено '+(row.saleAfter||'пусто')};
    }
    const info=privateVisibility(doc);
    return info.id===row.id&&info.type==='product'&&info.isPrivate?
      {confirmed:true,detail:'Visibility Private'}:
      {confirmed:false,reason:'сохранённая Visibility не подтверждена'};
  }
  async function fastVerifySaved(row,postResult,kind){
    let result=fastResultDoc(row,postResult.doc,kind);
    if(result.confirmed)return result;
    // The POST response may be a short/partial admin reply; exactly one read-only fallback.
    const verify=await fastRequest(row.productLink);
    result=fastResultDoc(row,verify.doc,kind);
    return result;
  }
  function fastTrashEvidence(row,result){
    const u=result.url,ids=u.searchParams.get('ids'),trashed=u.searchParams.get('trashed');
    if(ids&&ids.split(/[,\s]+/).some(id=>id&&id!==row.id))
      return {confirmed:false,reason:'в URL перенаправления указан другой ID: '+ids};
    if(u.pathname==='/wp-admin/edit.php'&&trashed==='1'&&(!ids||ids.split(/[,\s]+/).includes(row.id)))
      return {confirmed:true,reason:'WordPress trashed=1'+(ids?' ID '+ids:'')};
    const notices=Array.from(result.doc.querySelectorAll('.notice-success,.updated,#message'))
      .map(x=>clean(x.textContent)).filter(Boolean);
    const evidence=notices.find(t=>/(?:moved|sent|put).{0,90}trash|trash.{0,90}(?:moved|sent|put)|перемещен.{0,60}корзин|отправлен.{0,60}корзин/i.test(t));
    if(u.pathname==='/wp-admin/edit.php'&&evidence)
      return {confirmed:true,reason:evidence.slice(0,150)};
    return {confirmed:false,reason:'нет подтверждения Trash в ответе WordPress'};
  }
  async function fastProcess(row,worker) {
    const started=performance.now();
    let stage='read';
    try{
      row.fastWorker=worker;row.fastStage='loading';row.status='opening';
      row.message='Поток '+worker+': получаем HTML без картинок.';
      fastLog('info','Поток '+worker+' · ID '+row.id+': GET редактора, без отображения страницы и загрузки изображений.',row);
      const opened=await fastRequest(row.productLink);
      const doc=opened.doc;
      // Check exact URL/product identity before touching any product.
      const hiddenId=clean(doc.querySelector('#post_ID,input[name="post_ID"]')?.value);
      const type=clean(doc.querySelector('#post_type,input[name="post_type"]')?.value)||
        (doc.body?.classList.contains('post-type-product')?'product':'');
      if(hiddenId!==row.id||type!=='product')
        throw new Error('Неверная карточка товара: '+type+' '+hiddenId+' (ожидался '+row.id+').');
      row.foundName=clean(doc.querySelector('#title,input[name="post_title"]')?.value)||'';
      if(row.action==='trash'){
        const anchor=doc.querySelector('#delete-action a.submitdelete.deletion')||
          doc.querySelector('#delete-action a.submitdelete');
        const link=safeTrashLink(anchor,row);
        if(!link)throw new Error('Неверная или отсутствующая ссылка Trash с ID и nonce.');
        if(fastNoWriteOnPause(row))return;
        fastPrepareSubmission(row);
        stage='write';
        const res=await fastRequest(link);
        const confirmation=fastTrashEvidence(row,res);
        fastFinish(row,confirmation.confirmed?'done':'unconfirmed',
          confirmation.confirmed?'Товар ID '+row.id+' перемещён в Trash. '+confirmation.reason:
          'Trash отправлен, но результат не подтверждён: '+confirmation.reason,
          performance.now()-started);
        return;
      }
      const editor=fastEditor(row,doc);
      const payload=fastPayload(row,editor);
      if(row.action==='price'){
        const field=editor.form.querySelector('#_regular_price');
        const sale=editor.form.querySelector('#_sale_price');
        if(!field||field.disabled||field.name!=='_regular_price')
          throw new Error('Отсутствует редактируемое поле _regular_price.');
        if(!sale||sale.disabled||sale.name!=='_sale_price')
          throw new Error('Отсутствует редактируемое поле _sale_price; изменение цены остановлено.');
        if(!Number.isSafeInteger(row.cents)||!row.price)
          throw new Error('В Excel нет корректной цены.');
        row.before=clean(field.value);
        row.saleBefore=clean(sale.value);
        let regularMatches=false;
        try{regularMatches=parseRegularPrice(row.before).cents===row.cents;}
        catch{/* allow replacing a blank existing price */}
        if(regularMatches&&!row.saleBefore){
          row.after=row.before;row.saleAfter='';
          return fastFinish(row,'unchanged','Regular price уже '+row.before+
            ', Sale price пустая. Update не требуется.',performance.now()-started);
        }
        payload.set('_regular_price',row.price);
        payload.set('_sale_price','');
        fastLog('info','Поток '+worker+' · ID '+row.id+': Regular price '+
          (row.before||'(пусто)')+' → '+row.price+
          '; Sale price '+(row.saleBefore||'(пусто)')+' → (пусто).',row);
      }else if(row.action==='private'){
        const info=privateVisibility(doc);
        if(info.isPrivate)
          return fastFinish(row,'done','Товар уже Private; Update не требуется.',performance.now()-started);
        if(!editor.form.querySelector('#visibility-radio-private'))
          throw new Error('Нет переключателя Visibility Private; быстрый POST отменён.');
        payload.set('visibility','private');
        payload.set('post_status','private');
        payload.set('post_password','');
        fastLog('info','Поток '+worker+' · ID '+row.id+': меняем Visibility на Private.',row);
      }else throw new Error('Неизвестное действие '+row.action+'.');
      if(fastNoWriteOnPause(row))return;
      fastPrepareSubmission(row);
      stage='write';
      const res=await fastRequest(editor.action,{method:'POST',body:payload});
      row.fastStage='verifying';stage='verify';
      fastLog('info','ID '+row.id+': Update отправлен; сверяем результат.',row);
      const confirmation=await fastVerifySaved(row,res,row.action);
      if(confirmation.confirmed)
        fastFinish(row,'done',row.action==='price'?
          'Regular price сохранена: '+(row.before||'(пусто)')+' → '+
            (row.after||row.price)+'; Sale price очищена'+
            (row.saleBefore?' (была '+row.saleBefore+')':'')+'.':
          'Visibility Private подтверждена после Update.',
          performance.now()-started);
      else fastFinish(row,'unconfirmed','Update выполнен, но сохранение не подтверждено: '+
        confirmation.reason+'. Повторный Update не выполняется.',performance.now()-started);
    }catch(e){
      const message=String(e?.message||e),submitted=stage!=='read';
      const code=e?.httpStatus||0;
      if([429,502,503,504].includes(code)){
        fastErrors++;
        fastLimit=Math.max(1,Math.floor(fastLimit/2));
        fastLog('warn','Сервер вернул HTTP '+code+'. Параллельность снижена до '+fastLimit+'.',row);
      }
      fastFinish(row,submitted?'unconfirmed':'failed',
        (submitted?'Результат операции неизвестен. Повторная запись запрещена: ':
          'Операция не начата: ')+message,performance.now()-started);
      if(/авторизац|войти|wp-login\.php|требуется войти/i.test(message)||[401,403].includes(code)){
        state.running=false;state.paused=true;
        fastLog('error','Сессия или права WordPress требуют проверки. Новые задачи остановлены.',row,true);
      }
    }
  }
  function fastPump(){
    if(!state.running||state.settings.mode!=='fast')return;
    if(!ownLease()){
      state.running=false;state.paused=true;
      fastLog('error','Очередь остановлена: потеряно управление вкладкой.',null,true);
      return;
    }
    const cap=state.mode==='one'?1:Math.max(1,fastLimit);
    while(state.running&&fastActive.size<cap){
      const row=state.rows.find(r=>r.status==='pending');
      if(!row)break;
      row.action=row.action||state.action;
      const worker=++fastSerial;
      // Reserve synchronously to prevent two workers receiving the same ID.
      row.status='opening';row.fastStage='loading';row.fastWorker=worker;
      fastActive.set(row.id,true);
      fastSave(); // GET reservation is safe to restart after a reload; save batches.
      Promise.resolve().then(()=>fastProcess(row,worker))
        .catch(e=>{if(!TERMINAL.has(row.status))
          fastFinish(row,'unconfirmed','Неожиданная ошибка потока: '+String(e?.message||e));})
        .finally(()=>{
          fastActive.delete(row.id);
          if(!state.running){if(!fastActive.size){fastSave(true);release();}return;}
          if(state.mode==='one'&&state.runCount>=1){
            state.running=false;state.paused=true;
            fastSave(true);if(!fastActive.size)release();return;
          }
          if(!state.rows.some(r=>r.status==='pending')&&!fastActive.size){
            state.running=false;state.paused=false;
            fastLog('ok','Быстрая обработка файла завершена. Можно скачать Excel-отчёт.',null,true);
            release();return;
          }
          fastPump();
        });
    }
    fastSave();
  }
  function fastStart(one=false){
    if(fastActive.size||busy||moving)return toast('Дождитесь завершения текущих запросов.');
    readState();
    if(!state.rows.length)return toast('Сначала импортируйте Excel.');
    if(!['trash','private','price'].includes(state.action))
      return toast('Выберите действие с товарами.');
    const ready=state.rows.filter(r=>r.status==='pending');
    if(!ready.length)return toast('Нет ожидающих товаров.');
    if(ready.some(r=>r.action&&r.action!==state.action))
      return toast('Действие в очереди отличается от выбранного. Импортируйте Excel снова.');
    if(state.action==='price'&&ready.some(r=>!r.price||!Number.isSafeInteger(r.cents)))
      return toast('В очереди есть товары без корректной цены.');
    if(!claim())return;
    if(!confirm(one?
      'Быстрый тест одного товара? Запросы будут отправлены прямо в wp-admin.\n'+
        'ID '+ready[0].id+' · '+actionName(state.action):
      'Запустить быструю обработку '+ready.length+' товаров?\n'+
        'Параллельность: '+state.settings.workers+
        '\nДействие: '+actionName(state.action)+
        '\nКарточки и изображения не открываются. Сначала рекомендуется тест одного товара.')){
      release();return;
    }
    fastErrors=0;fastLimit=Math.min(10,Math.max(1,Number(state.settings.workers)||5));
    state.mode=one?'one':'batch';state.runCount=0;
    state.running=true;state.paused=false;state.owner=tabId;
    state.phase='fast';state.index=-1;
    if(!state.started)state.started=now();
    fastLog('info','Быстрый режим: '+(one?'тест 1 товара':'массовая обработка')+
      '; потоков '+(one?1:fastLimit)+'; '+actionName(state.action)+
      '. Редактор не отображается, изображения не скачиваются.',null,true);
    fastPump();
  }
  // On reload, all POST/Trash sent before the reload are NEVER submitted again.
  function fastRecoverAfterReload(){
    if(state.settings.mode!=='fast')return;
    let changed=false;
    for(const row of state.rows){
      if(TERMINAL.has(row.status))continue;
      if(row.fastStage==='submitted'||row.fastStage==='verifying'||
        row.status==='saving'||row.status==='verifying'){
        row.status='unconfirmed';row.finished=now();
        row.message='Страница перезагружена во время сохранения. Результат требует проверки; повторной отправки нет.';
        row.fastStage='';changed=true;
      }else if(row.fastStage==='loading'||row.status==='opening'){
        row.status='pending';row.fastStage='';changed=true;
      }
    }
    // One-item test must never silently continue with the next row after a reload.
    if(state.mode==='one'&&state.running){
      state.running=false;state.paused=true;state.runCount=Math.max(1,state.runCount);
      changed=true;
    }
    if(changed){
      state.index=-1;state.phase='fast';
      state.logs.push({time:now(),level:'warn',id:'',row:'',
        message:'Восстановление очереди после перезагрузки: операции с отправленным Update/Trash отмечены «Не подтверждено», без повторной отправки.'});
      fastSave(true);
    }
  }

  const CSS = [
    ':host{all:initial;font:13px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:#182438;color-scheme:light}',
    '*{box-sizing:border-box}button,input,select{font:inherit}button{cursor:pointer;border:1px solid #d9e1eb;border-radius:9px;padding:8px 12px;background:#fff;color:#26354c;font-weight:600}button:hover{background:#f0f4fa}button:disabled{opacity:.42;cursor:default}button:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid #4f7df3;outline-offset:2px}',
    '.panel{position:fixed;right:22px;bottom:20px;width:710px;height:650px;max-width:calc(100vw - 24px);max-height:calc(100vh - 40px);background:white;box-shadow:0 18px 70px #0f233b40;border:1px solid #dae2ee;border-radius:16px;display:flex;flex-direction:column;overflow:hidden;resize:both;pointer-events:auto}',
    '.panel.full{inset:24px!important;width:auto!important;height:auto!important;max-width:none;max-height:none;resize:none}',
    '.head{background:#14263e;color:white;padding:16px 18px;display:flex;align-items:center;gap:12px;cursor:move;touch-action:none;user-select:none}.logo{height:36px;width:36px;background:#4275ed;display:grid;place-items:center;border-radius:10px;font-size:18px;font-weight:800}.head b{font-size:16px;letter-spacing:.2px}.head small{color:#b1c2d9;display:block;font-size:11px}.head .spacer{flex:1}.icon{background:#ffffff13;color:#fff;border-color:#ffffff22;padding:4px 9px}',
    '.body{display:flex;flex-direction:column;min-height:0;flex:1}.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;padding:14px 16px 10px}.card{border:1px solid #e4eaf2;border-radius:10px;padding:9px 12px;background:#fafbfd}.card b{display:block;font-size:21px;line-height:1.25}.card span{font-size:11px;color:#68778c}.card.done b{color:#16835e}.card.bad b{color:#bd4f39}',
    '.progress{height:5px;background:#eaf0f8;margin:0 16px;border-radius:5px;overflow:hidden}.progress i{height:100%;display:block;background:#4076ed;transition:width .3s}.current{padding:10px 16px;display:flex;gap:8px;align-items:center;color:#586b83;min-height:42px}.dot{height:8px;width:8px;background:#93a1b3;border-radius:50%;flex:none}.dot.run{background:#3974ee;box-shadow:0 0 0 4px #e5edff}.current span:last-child{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
    '.operation{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin:0 16px 12px;padding:11px 12px;background:#f4f7fc;border:1px solid #dce6f3;border-radius:10px}.operation label{font-weight:750;font-size:12px;color:#31445c}.operation select{flex:1;min-width:200px}.operation small{flex-basis:100%;color:#6b7b92;font-size:11px}.actions{display:flex;gap:7px;flex-wrap:wrap;padding:0 16px 12px}.primary{background:#3468df!important;border-color:#3468df!important;color:white}.danger{color:#b04439;border-color:#eabcb6;background:#fff7f6}.tabs{display:flex;gap:18px;overflow:auto;flex-shrink:0;padding:0 16px;border-bottom:1px solid #e4eaf2}.tabs button{flex-shrink:0;border:0;border-radius:0;padding:10px 0;color:#7a889a;background:none;font-size:12px}.tabs button.on{color:#265cce;border-bottom:2px solid #3468df}',
    '.view{padding:12px 16px;min-height:0;flex:1;overflow:auto}.hidden{display:none!important}.tools{display:flex;gap:8px;margin-bottom:10px}input,select{border:1px solid #dce3ed;border-radius:8px;background:#fff;padding:8px;color:#22324a;min-width:0}input[type=search]{flex:1}table{border-collapse:collapse;width:100%;table-layout:fixed;font-size:12px}th{text-align:left;color:#7c899b;font-size:10px;text-transform:uppercase;letter-spacing:.5px;padding:8px 6px;background:#f7f9fc;position:sticky;top:0}td{padding:10px 6px;border-bottom:1px solid #edf1f6;vertical-align:top}tr.currentrow{background:#eff5ff}td .title{font-weight:600;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}td small{color:#8290a2;display:block;font-size:10px;overflow-wrap:anywhere}.badge{font-size:10px;padding:3px 7px;border-radius:6px;display:inline-block;white-space:nowrap;background:#eef2f7;color:#62738a}.badge.done{background:#e4f5ec;color:#147c55}.badge.bad{background:#fff0e8;color:#b24b28}.badge.uncertain{background:#fff6dc;color:#956400}.badge.work{background:#e7efff;color:#2b62ce}',
    '.foot{padding:10px 16px;border-top:1px solid #e4eaf2;display:flex;align-items:center;gap:8px;background:#fafbfd;font-size:11px;color:#7a889a}.foot span:first-child{flex:1}.foot button{padding:4px 8px;font-size:11px}.empty{text-align:center;padding:28px 16px;color:#77869b}.empty b{display:block;color:#31435c;font-size:16px;margin-bottom:6px}.drop{border:2px dashed #c4d2e9;border-radius:12px;text-align:center;padding:20px;margin:4px 0 12px;background:#f7faff}.drop.over{border-color:#3468df;background:#eaf1ff}.drop label{cursor:pointer;color:#2b63d4;font-weight:700}.drop small{display:block;margin-top:5px;color:#7b8a9e}.mapping{display:grid;grid-template-columns:120px 1fr;gap:8px;align-items:center}.mapping label{color:#61728a}.hint{font-size:12px;color:#6d7e94;margin:10px 0}.logline{display:grid;grid-template-columns:80px 80px 1fr;gap:8px;padding:8px 0;border-bottom:1px solid #edf1f6;font-size:11px;overflow-wrap:anywhere}.logline.error{color:#a83f2f}.logline.ok{color:#147c55}.logline.warn{color:#955e0d}',
    '.pill{position:fixed;right:22px;bottom:20px;border-radius:14px;background:#14263e;color:#fff;padding:12px 18px;box-shadow:0 8px 30px #0f233b40;pointer-events:auto}.pill:hover{background:#24466b;color:white}.notice{position:absolute;bottom:54px;left:16px;right:16px;z-index:8;padding:12px;background:#fff7e7;border:1px solid #e3cdab;border-radius:10px;color:#75551f;box-shadow:0 6px 20px #0002}.notice button{float:right;padding:0 6px;margin-right:5px}.help h3{font-size:13px;color:#263d5c}.help p,.help li{font-size:12px;color:#61728a}.help ol{padding-left:20px}.setting{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid #edf1f6}.setting input[type=number]{width:80px}.setting small{display:block;color:#7a899c}',
    '.logfilters{display:grid;grid-template-columns:minmax(0,1fr) 160px 120px;gap:8px}.logsummary{justify-content:space-between;align-items:center;color:#64758c;font-size:11px}.loglist{display:grid;gap:5px}.logline{grid-template-columns:85px 145px 1fr!important;gap:8px;background:#f9fbff;border:1px solid #e5ebf4;border-radius:9px;padding:10px 12px!important;color:#324960;font-size:12px!important}.logline>span:last-child{line-height:1.55;overflow-wrap:anywhere;white-space:pre-wrap}.logline time{color:#657a95;font-variant-numeric:tabular-nums}.logtitle{display:block;color:#728096;overflow-wrap:anywhere;font-size:10px;line-height:1.4}.logline.warn{border-left:3px solid #c8942a}.logline.error{border-left:3px solid #b55346}.logline.ok{border-left:3px solid #4c9a73}@media(max-width:550px){.logfilters{grid-template-columns:1fr 1fr}.logfilters input{grid-column:1/-1}.logline{grid-template-columns:70px 100px 1fr!important;padding:8px!important}}',
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
      '<div class="logo">P</div><div><b>Product Manager</b><small>'+esc(DOMAIN)+' · Trash / Private / Price · v'+VERSION+'</small></div>'+
      '<span class="spacer"></span><button class="icon" id="full" title="На весь экран">⛶</button>'+
      '<button class="icon" id="min" title="Свернуть">—</button></header>'+
      '<div class="body"><div class="stats">'+
      '<div class="card"><b id="total">0</b><span>Товаров в файле</span></div>'+
      '<div class="card done"><b id="done">0</b><span id="doneLabel">Выполнено</span></div>'+
      
      
      '<div class="card"><b id="issues">0</b><span>Требуют внимания</span></div></div>'+
      '<div class="progress"><i id="bar"></i></div>'+
      '<div class="current"><i class="dot" id="dot"></i><span id="current">Загрузите Excel с товарами</span></div>'+
      '<div class="operation"><label for="operation">Действие с товарами</label>'+
      '<select id="operation"><option value="">Выберите действие...</option>'+
      '<option value="trash">Удалить товар (Move to Trash)</option>'+
      '<option value="private">Сделать товар приватным (Private + Update)</option>'+ 
      '<option value="price">Редактировать цену (Regular price + Update)</option></select>'+
      '<small>Сначала выберите действие, затем нажмите «Тест: 1 товар» или «Начать». После запуска действие нельзя изменить без новой загрузки очереди.</small></div>'+
      '<div class="actions"><button class="primary" id="start">▶ Начать</button><button id="one">Тест: 1 товар</button>'+
      '<button id="pause">Ⅱ Пауза</button><button class="danger" id="stop">■ Стоп</button>'+
      '<button id="report">↓ Отчёт Excel</button></div>'+
      '<nav class="tabs">'+
      '<button data-tab="queue" class="on">Очередь</button>'+
      
      
      
      '<button data-tab="log">Журнал</button>'+
      '<button data-tab="import">Импорт</button>'+ 
      
      '<button data-tab="help">Как работать</button></nav>'+
      '<div class="view" id="queue"><div class="tools">'+
      '<input type="search" id="search" placeholder="Фильтр очереди: ссылка, ID, название…">'+
      '<select id="filter"><option value="all">Все статусы</option><option value="pending">Ожидают</option>'+
      '<option value="done">Выполнены</option><option value="unchanged">Цена без изменений</option>'+
      '<option value="not_found">Не найдены</option><option value="failed">Ошибки обработки</option>'+
      '<option value="unconfirmed">Не подтверждено</option><option value="skipped">Пропущено</option>'+
      '</select></div><div id="rows"></div></div>'+
      
      
      
      '<div class="view hidden" id="log"><div class="tools logfilters">'+
      '<input id="logSearch" type="search" placeholder="Поиск: ID, товар или текст ошибки">'+
      '<select id="logLevel"><option value="all">Все события</option>'+
      '<option value="error">Ошибки</option><option value="warn">Предупреждения</option>'+
      '<option value="ok">Успешно</option><option value="info">Действия</option></select>'+
      '<select id="logLimit"><option value="50">50 событий</option>'+
      '<option value="100" selected>100 событий</option><option value="200">200 событий</option></select>'+
      '</div><div class="tools logsummary"><span id="logFound">Журнал пуст</span>'+
      '<button id="downloadLog">↓ Скачать весь журнал TXT</button></div>'+
      '<div id="logs" class="loglist"></div></div>'+
      '<div class="view hidden" id="import"><div class="tools">'+
      '<button class="danger" id="reset">Сбросить очередь и журнал</button>'+
      '<button class="hidden" id="legacyPrices">Восстановить очередь цен из v2.4</button></div>'+
      '<div class="drop" id="drop"><label for="file">Выбрать Excel / CSV</label><input type="file" id="file" accept=".xlsx,.xls,.csv" hidden>'+
      '<small>Или перетащите файл сюда</small><small id="filename">Файл не выбран</small></div>'+
      '<div class="hidden" id="mapping"><div class="mapping"><label for="sheet">Лист Excel</label><select id="sheet"></select>'+ 
      '<label>Столбец URL</label><strong id="linkColumn">—</strong>'+
      '<label id="priceColumnWrap" class="hidden" for="priceColumn">Столбец цены</label>'+
      '<select class="hidden" id="priceColumn"><option value="">Выберите столбец цены...</option></select></div>'+
      '<p class="hint">Excel: product_link обязателен для всех действий. Для изменения цены дополнительно нужен столбец price, regular_price или «цена». Пример: €18,60 → 18,60. Столбец можно выбрать вручную.</p>'+
      '<pre id="preview" style="white-space:pre-wrap;overflow-wrap:anywhere;background:#f6f8fc;padding:10px;border-radius:9px"></pre>'+
      '<button class="primary" id="load">Загрузить в очередь</button></div>'+
      '<p class="hint">Сначала выберите действие над товарами в верхней панели. При ошибке отдельного товара скрипт фиксирует причину в журнале и продолжает работу. Очереди сайтов независимы.</p>'+
      '<div class="setting"><div>Режим обработки<small>Быстрый режим получает HTML без отображения карточек и изображений. Сначала протестируйте 1 товар.</small></div>'+
      '<select id="processingMode"><option value="compatible">Совместимый — открывать карточки</option>'+
      '<option value="fast">Быстрый — параллельные запросы</option></select></div>'+
      '<div class="setting" id="fastWorkersWrap"><div>Параллельные запросы<small>1–10, рекомендуем тест с 1, затем 3–5</small></div>'+
      '<input id="fastWorkers" type="number" min="1" max="10" step="1" value="5"></div>'+
      '<div class="hint" id="fastHint">Быстрый режим не загружает изображения редактора; операции выполняются без перезагрузки вкладки. WPML и WooCommerce требуют теста перед массовым запуском.</div>'+
      '<div class="setting"><div>Пауза между товарами<small>Только совместимый режим; секунды, 0.5–30</small></div>'+
      '<input id="delay" type="number" min="0.5" max="30" step="0.5" value="1"></div>'+
      '<p class="hint">Отсутствующая ссылка или ошибка сохранения не останавливают очередь. Если сессия WordPress истекла, скрипт ставится на паузу до повторного входа.</p></div>'+
      '<div class="view help hidden" id="help">'+
      '<h3>Выбор действия и импорт</h3><ol>'+
      '<li>Подготовьте Excel со столбцом product_link. Для изменения цены добавьте price либо «цена».</li>'+
      '<li>В списке «Действие с товарами» выберите Trash, Private либо «Редактировать цену». Затем импортируйте Excel.</li>'+
      '<li>Нажмите «Тест: 1 товар» для пробного применения выбранного действия к одному товару.</li>'+
      '<li>После теста нажмите «Начать» для обработки остальных ссылок.</li>'+
      '</ol><h3>Режим: Удалить товар</h3><p>Открывается product_link → Move to Trash → перенаправление WordPress → запись результата → следующая ссылка. Без дополнительных запросов к карточке и страницам корзины.</p>'+
      '<h3>Режим: Сделать приватным</h3><p>Открывается product_link → Edit visibility → Private → OK → Update → проверка выбранного Visibility на заново загруженной странице. Обычно без дополнительных запросов. Если обновления страницы не было — одна проверка через сервер. Повторно Update не нажимается.</p>'+
      '<h3>Редактирование цены</h3><p>Открывается product_link → проверяется ID → Regular price из Excel, Sale price очищается → Update → проверяются оба сохранённых поля. Даже если Regular price уже совпадает, непустая Sale price будет очищена через Update. Если обе цены соответствуют заданию, Update пропускается. Повторный Update автоматически не выполняется.</p>'+
      '<h3>Быстрый режим 1–10 запросов</h3><p>В «Импорте» выберите быстрый режим и количество параллельных запросов. Товары запрашиваются как HTML, изображения удаляются до разбора, интерфейс WordPress не загружается. Update/Trash выполняются один раз с проверкой ID и nonce; при отсутствии подтверждения — «Не подтверждено», автоматических повторов нет. В быстром режиме кнопка «Пауза» прекращает выдачу новых товаров, уже отправленные запросы завершаются.</p>'+
      '<h3>Очередь и журнал</h3><p>Все результаты, включая «Не найдены», «Ошибки обработки» и «Не подтверждено», выбираются выпадающим фильтром во вкладке «Очередь». Во вкладке «Журнал» можно искать по ID и фильтровать сообщения по типу. Весь журнал доступен в TXT, результаты — в Excel.</p>'+
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
      if(!['','trash','private','price'].includes(e.target.value))return;
      state.action=e.target.value;
      log('info','Оператор выбрал действие: '+actionName(state.action)+'.');
      fillMapping();
    };
    $('#start').onclick=()=>start(false);
    $('#one').onclick=()=>start(true);
    $('#pause').onclick=()=>pause();
    $('#stop').onclick=stop;
    $('#report').onclick=report;
    $('#downloadLog').onclick=downloadLog;
    $('#reset').onclick=reset;
    $('#legacyPrices').onclick=migrateLegacyPrices;
    $('#prev').onclick=()=>{page=Math.max(0,page-1);renderRows();};
    $('#next').onclick=()=>{page++;renderRows();};
    $('#search').oninput=e=>{search=e.target.value;page=0;renderRows();};
    $('#filter').onchange=e=>{filter=e.target.value;sessionStorage.setItem(PREFIX+':ui:filter',filter);page=0;renderRows();};
    $('#file').onchange=e=>readExcel(e.target.files?.[0]);
    $('#sheet').onchange=fillMapping;
    $('#priceColumn').onchange=e=>{
      priceColumn=e.target.value===''?-1:Number(e.target.value);
      fillMapping();
    };
    $('#logSearch').oninput=e=>{logSearch=e.target.value;sessionStorage.setItem(PREFIX+':ui:logSearch',logSearch);renderLogs();};
    $('#logLevel').onchange=e=>{logLevel=e.target.value;sessionStorage.setItem(PREFIX+':ui:logLevel',logLevel);renderLogs();};
    $('#logLimit').onchange=e=>{logLimit=Number(e.target.value)||100;sessionStorage.setItem(PREFIX+':ui:logLimit',String(logLimit));renderLogs();};
    
    $('#load').onclick=importQueue;
    const drop=$('#drop');
    drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('over');});
    drop.addEventListener('dragleave',()=>drop.classList.remove('over'));
    drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('over');readExcel(e.dataTransfer.files[0]);});
    $('#filter').value=filter;
    $('#logSearch').value=logSearch;
    $('#logLevel').value=logLevel;
    $('#logLimit').value=String(logLimit);
    $('#processingMode').value=state.settings.mode;
    $('#fastWorkers').value=state.settings.workers;
    $('#processingMode').onchange=e=>{
      if(!canEdit()){e.target.value=state.settings.mode;return;}
      if(state.started){e.target.value=state.settings.mode;return toast('Для смены режима после старта загрузите новую очередь.');}
      state.settings.mode=e.target.value==='fast'?'fast':'compatible';
      log('info','Режим обработки: '+(state.settings.mode==='fast'?'быстрый параллельный':'совместимый')+'.');
      render();
    };
    $('#fastWorkers').onchange=e=>{
      if(!canEdit()){e.target.value=state.settings.workers;return;}
      state.settings.workers=Math.max(1,Math.min(10,Math.round(Number(e.target.value)||5)));
      e.target.value=state.settings.workers;
      log('info','Количество параллельных запросов: '+state.settings.workers+'.');
    };
    $('#delay').value=state.settings.delay;
    $('#delay').onchange=e=>{
      if(!canEdit()){e.target.value=state.settings.delay;return;}
      state.settings.delay=Math.round(Math.min(30,Math.max(0.5,Number(e.target.value)||1))*2)/2;
      e.target.value=state.settings.delay;persist();
    };
    if (GM_getValue(PREFIX+':minimized',false)) {
      panel.classList.add('hidden');$('#pill').classList.remove('hidden');
    }
    if(!state.rows.length)setTab('import');
    else setTab(activeTab);
  }
  function setTab(tab) {
    activeTab=tab;sessionStorage.setItem(PREFIX+':ui:tab',tab);page=0;
    for(const name of ['queue','log','import','help'])
      $('#'+name).classList.toggle('hidden',name!==tab);
    $$('[data-tab]').forEach(b=>b.classList.toggle('on',b.dataset.tab===tab));
    render();
  }
  function subset() {
    return state.rows.filter(r=>{
      const matches=filter==='all'||(filter==='pending'?
        ['pending','opening','deleting','saving','verifying'].includes(r.status) : r.status===filter);
      return matches && normalized(r.id+' '+r.productLink+' '+r.foundName+
        ' '+r.price+' '+r.before+' '+r.after).includes(normalized(search));
    });
  }
  function renderRows() {
    if(!root)return;
    if(activeTab!=='queue')return;
    const rows=subset();
    const pages=Math.max(1,Math.ceil(rows.length/PAGESIZE));
    page=Math.min(page,pages-1);
    $('#pages').textContent=(page+1)+' / '+pages;
    $('#prev').disabled=page===0;$('#next').disabled=page>=pages-1;
    const target=$('#rows');
    if(!rows.length){
      target.innerHTML='<div class="empty"><b>Список пуст</b>'+
        'Импортируйте Excel или измените фильтр.</div>';
      return;
    }
    target.innerHTML='<table><thead><tr><th style="width:50px">Строка</th><th>Игра / ID</th>'+
      '<th style="width:112px">Статус</th></tr></thead><tbody>'+
      rows.slice(page*PAGESIZE,page*PAGESIZE+PAGESIZE).map(r=>
        '<tr class="'+(current()===r?'currentrow':'')+'"><td>'+r.excelRow+'</td>'+
        '<td><span class="title" title="'+esc(r.foundName||r.productLink||'—')+'">'+esc(r.foundName||('Товар ID '+(r.id||'—')))+
        '</span><small>'+esc(r.productLink||'')+'</small>'+
        (r.action==='price'?'<small>Цена: '+esc(r.price||r.rawPrice||'—')+
          ' · До: '+esc(r.before||'—')+' · После: '+esc(r.after||'—')+
          '</small><small>Sale price: '+esc(r.saleBefore||'пусто')+
          ' → '+esc(r.saleAfter??'—')+'</small>':'')+
        '<small>'+esc(r.message||'')+'</small></td>'+
        '<td><span class="badge '+(statuses[r.status]||'neutral')+'">'+esc(rowLabel(r))+
        '</span></td></tr>'
      ).join('')+'</tbody></table>';
  }
  function renderLogs() {
    const query=normalized(logSearch),names=new Map(
      state.rows.filter(r=>r.id).map(r=>[r.id,r.foundName||'']));
    const found=state.logs.filter(x=>(logLevel==='all'||x.level===logLevel) &&
      normalized(x.id+' '+x.row+' '+x.message+' '+(names.get(x.id)||'')).includes(query));
    const latest=found.slice(-logLimit).reverse();
    $('#logFound').textContent='Найдено: '+found.length+' из '+state.logs.length+
      ' · показано '+latest.length;
    $('#logs').innerHTML=latest.map(x=>
      '<div class="logline '+esc(x.level)+'"><time>'+esc(new Date(x.time).toLocaleTimeString('ru-RU'))+
      '</time><span>'+esc(x.id?'ID '+x.id:'Система')+
      (names.get(x.id)?'<small class="logtitle">'+esc(names.get(x.id))+'</small>':'')+
      '</span><span>'+esc(x.message)+'</span></div>'
    ).join('')||'<div class="empty">Нет сообщений по выбранному фильтру.</div>';
  }
  function render() {
    if(!root)return;
    const counts=Object.create(null);
    for(const r of state.rows)counts[r.status]=(counts[r.status]||0)+1;
    const count=s=>counts[s]||0;
    const done=count('done')+count('unchanged'),missing=count('not_found'),error=count('failed'),unconfirmed=count('unconfirmed');
    const pending=count('pending')+count('opening')+count('deleting')+count('saving')+count('verifying');
    $('#total').textContent=state.rows.length;
    $('#done').textContent=done;
    $('#doneLabel').textContent=state.action==='private'?'Приватных':state.action==='price'?'Цены обработаны':'Выполнено';
    $('#issues').textContent=missing+error+unconfirmed;
    $('#bar').style.width=(state.rows.length?
      Math.round(100*(done+missing+error+unconfirmed+count('skipped'))/state.rows.length):0)+'%';
    const row=current();
    $('#current').textContent=state.running&&state.settings.mode==='fast'?
      'Быстрый режим · активные запросы '+fastActive.size+'/'+
      (state.mode==='one'?1:fastLimit)+' · выполнено '+done+' из '+state.rows.length:
      state.running?
      'Обработка ссылки ID '+(row?.id||'—')+' · '+(labels[row?.status]||'Открытие'):
      state.paused?'Пауза · нажмите «Начать» для продолжения':
      state.rows.length?(state.filename+' · ожидают: '+pending):
      'Загрузите Excel с product_link';
    $('#dot').classList.toggle('run',state.running);
    const hasPending=nextPending()>=0||!!(row&&!TERMINAL.has(row.status))||fastActive.size>0;
    const lease=GM_getValue(LEASE,null);
    const lockedElsewhere=!!(lease && lease.until>Date.now() && lease.id!==tabId);
    const activeHere=state.running && state.owner===tabId;
    const recoverable=state.running && state.owner!==tabId && !lockedElsewhere;
    $('#start').textContent=recoverable?'▶ Возобновить':state.paused?'▶ Продолжить':'▶ Начать';
    $('#start').disabled=activeHere||lockedElsewhere||busy||fastActive.size>0||!hasPending;
    $('#one').disabled=activeHere||lockedElsewhere||busy||fastActive.size>0||!hasPending;
    $('#operation').value=state.action||'';
    $('#operation').disabled=activeHere||lockedElsewhere||busy||fastActive.size>0||!!state.started;
    if(!['trash','private','price'].includes(state.action)){
      $('#start').disabled=true;$('#one').disabled=true;
    }
    $('#pause').disabled=!activeHere;
    $('#stop').disabled=!activeHere;
    $('#report').disabled=!state.rows.length;
    $('#load').disabled=state.running||busy||fastActive.size>0||!book||excelColumn<0||(state.action==='price'&&priceColumn<0);
    $('#reset').disabled=state.running||busy||fastActive.size>0;
    $('#legacyPrices').disabled=state.running||busy||fastActive.size>0;
    const oldPrice=GM_getValue(PRICE_KEY,null);
    $('#legacyPrices').classList.toggle('hidden',
      !oldPrice||!Array.isArray(oldPrice.rows)||!oldPrice.rows.length);
    $('#priceColumn').disabled=state.running||busy||fastActive.size>0||!book;
    $('#processingMode').disabled=state.running||busy||fastActive.size>0||!!state.started;
    $('#fastWorkers').disabled=state.running||busy||fastActive.size>0;
    $('#fastWorkersWrap').classList.toggle('hidden',state.settings.mode!=='fast');
    $('#delay').disabled=state.settings.mode==='fast'||state.running||busy;
    $('#footer').textContent=state.updated?'Сохранено · '+new Date(state.updated).toLocaleTimeString('ru-RU'):
      'Очередь сохраняется автоматически';
    const labelsForFilter={
      all:'Все статусы ('+state.rows.length+')',
      pending:'Ожидают ('+pending+')',done:'Выполнены ('+count('done')+')',
      unchanged:'Цена без изменений ('+count('unchanged')+')',
      not_found:'Не найдены ('+missing+')',failed:'Ошибки обработки ('+error+')',
      unconfirmed:'Не подтверждено ('+unconfirmed+')',skipped:'Пропущено ('+count('skipped')+')'
    };
    $$('#filter option').forEach(opt=>{
      if(labelsForFilter[opt.value])opt.textContent=labelsForFilter[opt.value];
    });
    if(activeTab==='queue')renderRows();
    if(activeTab==='log')renderLogs();
    $('.foot #prev').classList.toggle('hidden',activeTab!=='queue');
    $('.foot #next').classList.toggle('hidden',activeTab!=='queue');
    $('.foot #pages').classList.toggle('hidden',activeTab!=='queue');
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
  if(state.settings.mode==='fast')fastRecoverAfterReload();
  if(state.running && state.owner === tabId) {
    if(claim()) {
      if(state.settings.mode==='fast'){
        fastLimit=Math.min(10,Math.max(1,Number(state.settings.workers)||5));
        fastLog('warn','Восстановление быстрого режима после обновления страницы.',null,true);
        fastPump();
      }else{
        log('info','Возобновление после перехода или перезагрузки вкладки.');
        schedule(250);
      }
    } else {
      state.running=false;state.paused=true;persist();
    }
  }
})();