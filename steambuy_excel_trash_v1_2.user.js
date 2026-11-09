// ==UserScript==
// @name         SteamBuy — Product Trash Manager
// @namespace    https://steambuy.pro/
// @version      2.1.0
// @description  Excel product_link → открыть редактор → Move to Trash → следующий; лог и отчёты, 7 сайтов.
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
  // Version 2 uses a new queue key: old running 1.x jobs cannot trigger deletions after upgrade.
  const VERSION = '2.1.0';
  const DOMAIN = location.hostname;
  const PREFIX = 'sb-product-trash-v3:' + DOMAIN;
  const OLD_LEASE = 'sb-product-trash-v2:' + DOMAIN + ':lease';
  const KEY = PREFIX + ':state';
  const LEASE = PREFIX + ':lease';
  const LAYOUT = PREFIX + ':layout';
  const TAB = PREFIX + ':tab-id';
  const LIST = location.origin + '/wp-admin/edit.php?post_type=product';
  const MAX_VERIFY_PAGES = 15;
  const MAX_LOGS = 1500;
  const PAGESIZE = 25;
  const OK = 'done';
  const TERMINAL = new Set(['done', 'not_found', 'failed', 'skipped']);
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
    v: 3, filename: '', sheet: '', rows: [], logs: [], index: -1,
    phase: 'idle', running: false, paused: false, mode: 'batch', runCount: 0,
    owner: '', pending: null,
    settings: { delay: 2, stopOnError: false }, updated: '', started: ''
  });
  let state = GM_getValue(KEY, null) || defaults();
  if (state.v !== 3 || !Array.isArray(state.rows)) state = defaults();
  state.settings = Object.assign({delay: 2, stopOnError: false}, state.settings || {});
  let root, panel, book = null, fileName = '';
  let activeTab = 'queue', filter = 'all', search = '', page = 0;
  let busy = false, moving = false, timer = null, heartbeat = null, expanded = false, dragging = null;
  const labels = {pending:'Ожидает', opening:'Открытие', deleting:'Удаляется', verifying:'Проверка',
    done:'Удалён', not_found:'Не найден', failed:'Ошибка удаления', skipped:'Пропущен'};
  const statuses = {done:'done', failed:'bad', not_found:'bad', skipped:'neutral',
    opening:'work', deleting:'work', verifying:'work', pending:'neutral'};
  function persist() {
    state.updated = now();
    GM_setValue(KEY, state);
    render();
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
      toast('Другая вкладка ещё работает на старой версии скрипта. Остановите её, прежде чем начинать.');
      return false;
    }
    const lease = GM_getValue(LEASE, null);
    if (lease && lease.until > Date.now() && lease.id !== tabId) {
      toast('Удаление уже запущено в другой вкладке этого сайта. Сначала остановите его там.');
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
    persist();
    if (row) log('info', 'Переход: ' + u.pathname + u.search, row);
    if (u.href === location.href) location.reload();
    else location.assign(u.href);
    setTimeout(() => {
      if (!moving || document.visibilityState === 'hidden') return;
      moving = false;
      state.running = false;
      state.paused = true;
      persist(); release();
      toast('Переход не завершился. Обработка остановлена; проверьте страницу и нажмите «Продолжить».');
    }, 18000);
  }
  function current() { return state.rows[state.index] || null; }
  function nextPending() { return state.rows.findIndex(r => r.status === 'pending'); }
  function schedule(ms = 80) {
    clearTimeout(timer);
    if (state.running && !busy && !moving) timer = setTimeout(runStep,ms);
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
    row.status = status;
    row.message = message;
    row.finished = now();
    state.pending = null;
    state.phase = 'idle';
    state.index = -1;
    state.runCount++;
    log(status === 'done' ? 'ok' : 'error', message, row);
    if (state.mode === 'one' && state.runCount >= 1) {
      state.running = false;
      state.paused = true;
      log('info','Тест одного товара завершён. Остальная очередь не затронута.');
      release();
      return;
    }
    if (status !== 'done' && status !== 'not_found' && state.settings.stopOnError) {
      state.running = false;
      state.paused = true;
      log('error','Обработка остановлена согласно настройке «Пауза при ошибке».',row);
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
      u.searchParams.get('action')==='edit'&&u.searchParams.get('post')===row.id;
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
  function openProduct(row) {
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
    const actualId=clean(document.querySelector('#post_ID,input[name="post_ID"]')?.value);
    const type=clean(document.querySelector('#post_type,input[name="post_type"]')?.value)||
      (document.body.classList.contains('post-type-product')?'product':'');
    const status=clean(document.querySelector('#post_status,input[name="post_status"]')?.value);
    if(!form||!actualId)return finalize(row,'not_found','Карточка не найдена: нет формы редактирования товара или post_ID.');
    if(actualId!==row.id)return finalize(row,'failed','WordPress открыл другой ID: '+actualId+' (ожидался '+row.id+').');
    if(type!=='product')return finalize(row,'failed','ID '+row.id+' не является товаром (тип: '+(type||'неизвестно')+').');
    row.wpPostId=actualId;
    row.foundName=clean(document.querySelector('#title,input[name="post_title"]')?.value)||
      clean(document.querySelector('.editor-post-title__input')?.textContent)||'';
    if(status==='trash')return finalize(row,'done','Товар уже находится в корзине.');
    const anchor=document.querySelector('#delete-action a.submitdelete.deletion')||
      document.querySelector('#delete-action a.submitdelete');
    const link=safeTrashLink(anchor,row);
    if(!link)return finalize(row,'failed','Товар найден, но кнопка #delete-action → Move to Trash с проверяемым nonce отсутствует.');
    // A persisted verify checkpoint ensures a reload never clicks Move to Trash a second time.
    state.phase='verify';state.pending={id:row.id,href:link,at:now()};
    row.status='deleting';row.message='Нажимаем Move to Trash, затем проверяем.';
    log('info','Товар '+(row.foundName||'без названия')+' (ID '+row.id+'). Клик по Move to Trash.',row);
    anchor.click();
    // WordPress normally navigates. If an interceptor blocks navigation, verify without re-clicking.
    setTimeout(()=>{
      if(state.running&&ownLease()&&current()?.id===row.id&&state.phase==='verify'&&!busy&&!moving)
        verifyTrash(row);
    },4000);
  }
  function successRedirect(row) {
    const u=new URL(location.href);
    if(u.origin!==location.origin||!u.pathname.endsWith('/wp-admin/edit.php')||
      u.searchParams.get('trashed')!=='1')return false;
    return (u.searchParams.get('ids')||'').split(/[,\s]+/).includes(row.id);
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
  async function verifyTrash(row) {
    if(busy)return;
    busy=true;row.status='verifying';row.message='Проверка без повторного нажатия Trash.';
    persist();
    try {
      if(successRedirect(row)){
        finalize(row,'done','WordPress подтвердил Trash через redirect (trashed=1, ids='+row.id+').');
        return;
      }
      // Direct edit-page check does not use product search.
      try{
        const doc=await fetchAdmin(row.productLink);
        const die=clean(doc.querySelector('.wp-die-message')?.textContent);
        if(die&&/trash|корзин/i.test(die)){
          finalize(row,'done','WordPress сообщает, что товар в корзине.');
          return;
        }
        const editId=clean(doc.querySelector('#post_ID,input[name="post_ID"]')?.value);
        const status=clean(doc.querySelector('#post_status,input[name="post_status"]')?.value);
        if(editId===row.id&&status==='trash'){
          finalize(row,'done','Проверен статус товара: trash.');
          return;
        }
      }catch(e){
        log('warn','Проверка карточки: '+e.message,row);
        if(/сессия|войти/i.test(e.message))throw e;
      }
      // Fallback enumerates the Trash table; never invokes the WordPress search field.
      for(let p=1;p<=MAX_VERIFY_PAGES;p++){
        const u=new URL('/wp-admin/edit.php',location.origin);
        u.searchParams.set('post_type','product');
        u.searchParams.set('post_status','trash');
        u.searchParams.set('paged',String(p));
        const doc=await fetchAdmin(u.href);
        if(!doc.querySelector('#the-list'))throw new Error('Нет таблицы товаров в корзине.');
        if(doc.getElementById('post-'+row.id)){
          finalize(row,'done','Товар '+row.id+' найден в корзине, страница '+p+'.');
          return;
        }
        if(!doc.querySelector('.tablenav-pages a.next-page:not(.disabled)'))break;
        if(p===MAX_VERIFY_PAGES)throw new Error('Лимит просмотра корзины: '+MAX_VERIFY_PAGES+' страниц.');
      }
      finalize(row,'failed','После Move to Trash товар не подтверждён в корзине. Автоматический повтор запрещён.');
    }catch(e){
      finalize(row,'failed','Не удалось проверить удаление: '+e.message+'. Повторный Trash не выполняется.');
      if(/сессия|войти/i.test(e.message)){
        state.running=false;state.paused=true;persist();release();
      }
    }finally{
      busy=false;
      if(state.running)schedule(Math.max(700,Number(state.settings.delay)*1000));
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
    if(['verify','verifying','deleting'].includes(state.phase)){
      log('warn','Возобновление: проверка предыдущего Trash. Второго нажатия не будет.',row);
      return verifyTrash(row);
    }
    state.phase='open';openProduct(row);
  }
  function start(one=false) {
    if(busy||moving)return toast('Дождитесь завершения текущего действия.');
    readState();
    if(!state.rows.length)return toast('Сначала загрузите Excel со столбцом product_link.');
    if(!claim())return;
    const ongoing=state.index>=0&&!!current()&&!TERMINAL.has(current().status);
    if(!ongoing){state.index=nextPending();state.phase='open';}
    if(state.index<0){release();return toast('Все ссылки обработаны. Загрузите новый Excel.');}
    const row=current();
    if(row.status==='pending'&&state.phase==='open'){
      const count=state.rows.filter(r=>r.status==='pending').length;
      if(!confirm(one?
        'ТЕСТ: переместить ровно один товар в Trash?\n'+row.productLink+
        '\n\nЭто реальное перемещение в корзину.' :
        'Начать массовое перемещение товаров в Trash?\nСайт: '+DOMAIN+
        '\nОсталось ссылок: '+count)){
        release();return;
      }
    }
    state.mode=one?'one':'batch';
    state.runCount=0;state.running=true;state.paused=false;state.owner=tabId;
    if(!state.started)state.started=now();
    log('info',(one?'Тест одной ссылки':'Массовая обработка product_link')+' запущен.',row);
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
    } catch (e) {toast('Не удалось прочитать Excel: '+e.message);}
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
    state=defaults();state.filename=fileName;state.sheet=book.SheetNames[Number($('#sheet').value)];state.rows=rows;
    log('info','Импорт '+fileName+': '+rows.length+' ссылок, к обработке '+good+', пропущено '+(rows.length-good)+'.');
    setTab('queue');
  }
  function report(){
    if(!state.rows.length)return toast('Очередь пуста.');
    if(typeof XLSX==='undefined')return toast('Библиотека XLSX недоступна.');
    const cols=['Строка Excel','product_link','ID из ссылки','Название на сайте','Статус','Комментарий','Завершено'];
    const line=r=>[r.excelRow,r.productLink,r.id,r.foundName,labels[r.status]||r.status,r.message,r.finished];
    const wb=XLSX.utils.book_new();
    const add=(name,rows)=>{
      const ws=XLSX.utils.aoa_to_sheet([cols,...rows.map(line)]);
      ws['!cols']=[{wch:13},{wch:100},{wch:18},{wch:50},{wch:24},{wch:105},{wch:25}];
      XLSX.utils.book_append_sheet(wb,ws,name);
    };
    add('Все товары',state.rows);
    add('Удалено',state.rows.filter(r=>r.status==='done'));
    add('Не найдено',state.rows.filter(r=>r.status==='not_found'));
    add('Ошибки удаления',state.rows.filter(r=>r.status==='failed'));
    add('Пропущено',state.rows.filter(r=>r.status==='skipped'));
    const ws=XLSX.utils.aoa_to_sheet([['Время','Уровень','ID','Строка','Сообщение'],
      ...state.logs.map(x=>[x.time,x.level,x.id,x.row,x.message])]);
    ws['!cols']=[{wch:26},{wch:14},{wch:20},{wch:14},{wch:100}];
    XLSX.utils.book_append_sheet(wb,ws,'Журнал');
    const safe=(state.filename||'products').replace(/\.(xlsx|xls|csv)$/i,'').replace(/[\\/:*?"<>|]/g,'_');
    XLSX.writeFile(wb,safe+'_'+DOMAIN+'_trash_result.xlsx');
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
    '.body{display:flex;flex-direction:column;min-height:0;flex:1}.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;padding:14px 16px 10px}.card{border:1px solid #e4eaf2;border-radius:10px;padding:9px 12px;background:#fafbfd}.card b{display:block;font-size:21px;line-height:1.25}.card span{font-size:11px;color:#68778c}.card.done b{color:#16835e}.card.bad b{color:#bd4f39}',
    '.progress{height:5px;background:#eaf0f8;margin:0 16px;border-radius:5px;overflow:hidden}.progress i{height:100%;display:block;background:#4076ed;transition:width .3s}.current{padding:10px 16px;display:flex;gap:8px;align-items:center;color:#586b83;min-height:42px}.dot{height:8px;width:8px;background:#93a1b3;border-radius:50%;flex:none}.dot.run{background:#3974ee;box-shadow:0 0 0 4px #e5edff}.current span:last-child{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
    '.actions{display:flex;gap:7px;flex-wrap:wrap;padding:0 16px 12px}.primary{background:#3468df!important;border-color:#3468df!important;color:white}.danger{color:#b04439;border-color:#eabcb6;background:#fff7f6}.tabs{display:flex;gap:18px;overflow:auto;flex-shrink:0;padding:0 16px;border-bottom:1px solid #e4eaf2}.tabs button{flex-shrink:0;border:0;border-radius:0;padding:10px 0;color:#7a889a;background:none;font-size:12px}.tabs button.on{color:#265cce;border-bottom:2px solid #3468df}',
    '.view{padding:12px 16px;min-height:0;flex:1;overflow:auto}.hidden{display:none!important}.tools{display:flex;gap:8px;margin-bottom:10px}input,select{border:1px solid #dce3ed;border-radius:8px;background:#fff;padding:8px;color:#22324a;min-width:0}input[type=search]{flex:1}table{border-collapse:collapse;width:100%;table-layout:fixed;font-size:12px}th{text-align:left;color:#7c899b;font-size:10px;text-transform:uppercase;letter-spacing:.5px;padding:8px 6px;background:#f7f9fc;position:sticky;top:0}td{padding:10px 6px;border-bottom:1px solid #edf1f6;vertical-align:top}tr.currentrow{background:#eff5ff}td .title{font-weight:600;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}td small{color:#8290a2;display:block;font-size:10px;overflow-wrap:anywhere}.badge{font-size:10px;padding:3px 7px;border-radius:6px;display:inline-block;white-space:nowrap;background:#eef2f7;color:#62738a}.badge.done{background:#e4f5ec;color:#147c55}.badge.bad{background:#fff0e8;color:#b24b28}.badge.work{background:#e7efff;color:#2b62ce}',
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
      '<button class="pill hidden" id="pill">Product Trash · открыть</button>'+
      '<section class="panel" id="panel"><header class="head" id="drag">'+
      '<div class="logo">T</div><div><b>Product Trash Manager</b><small>'+esc(DOMAIN)+' · WooCommerce · v'+VERSION+'</small></div>'+
      '<span class="spacer"></span><button class="icon" id="full" title="На весь экран">⛶</button>'+
      '<button class="icon" id="min" title="Свернуть">—</button></header>'+
      '<div class="body"><div class="stats">'+
      '<div class="card"><b id="total">0</b><span>Товаров в файле</span></div>'+
      '<div class="card done"><b id="done">0</b><span>Удалено</span></div>'+
      '<div class="card bad"><b id="missing">0</b><span>Не найдено</span></div>'+
      '<div class="card bad"><b id="errors">0</b><span>Ошибки</span></div></div>'+
      '<div class="progress"><i id="bar"></i></div>'+
      '<div class="current"><i class="dot" id="dot"></i><span id="current">Загрузите Excel с товарами</span></div>'+
      '<div class="actions"><button class="primary" id="start">▶ Начать</button><button id="one">Тест: 1 товар</button>'+
      '<button id="pause">Ⅱ Пауза</button><button class="danger" id="stop">■ Стоп</button>'+
      '<button id="report">↓ Отчёт Excel</button></div>'+
      '<nav class="tabs">'+
      '<button data-tab="queue" class="on">Очередь</button>'+
      '<button data-tab="not_found">Не найдены</button>'+
      '<button data-tab="failed">Ошибки удаления</button>'+
      '<button data-tab="log">Журнал</button>'+
      '<button data-tab="import">Импорт Excel</button>'+
      '<button data-tab="help">Как работать</button></nav>'+
      '<div class="view" id="queue"><div class="tools">'+
      '<input type="search" id="search" placeholder="Фильтр очереди: ссылка, ID, название…">'+
      '<select id="filter"><option value="all">Все статусы</option><option value="pending">В очереди</option>'+
      '<option value="done">Удалены</option><option value="skipped">Пропущено</option></select></div><div id="rows"></div></div>'+
      '<div class="view hidden" id="not_found"><div id="missingRows"></div></div>'+
      '<div class="view hidden" id="failed"><div id="failedRows"></div></div>'+
      '<div class="view hidden" id="log"><div class="tools"><button id="downloadLog">↓ Скачать журнал TXT</button><span class="hint">Последние 1500 событий</span></div><div id="logs"></div></div>'+
      '<div class="view hidden" id="import"><div class="tools"><button class="danger" id="reset">Сбросить очередь и журнал</button></div>'+
      '<div class="drop" id="drop"><label for="file">Выбрать Excel / CSV</label><input type="file" id="file" accept=".xlsx,.xls,.csv" hidden>'+
      '<small>Или перетащите файл сюда</small><small id="filename">Файл не выбран</small></div>'+
      '<div class="hidden" id="mapping"><div class="mapping"><label for="sheet">Лист Excel</label><select id="sheet"></select>'+ 
      '<label>Столбец</label><strong id="linkColumn">—</strong></div>'+
      '<p class="hint">Обязательный столбец: product_link. В нём укажите ссылки вида https://'+esc(DOMAIN)+'/wp-admin/post.php?post=3287&amp;action=edit. Другие столбцы не нужны. ID автоматически извлекается из ссылки.</p>'+
      '<pre id="preview" style="white-space:pre-wrap;overflow-wrap:anywhere;background:#f6f8fc;padding:10px;border-radius:9px"></pre>'+
      '<button class="primary" id="load">Загрузить в очередь</button></div>'+
      '<p class="hint">Каждый сайт хранит собственную очередь. Повторные ссылки на товар пропускаются; чужие домены блокируются.</p>'+
      '<div class="setting"><div>Пауза между товарами<small>Секунды, 1–30</small></div>'+
      '<input id="delay" type="number" min="1" max="30" value="2"></div>'+
      '<label class="setting"><input id="stopOnError" type="checkbox"><div>Останавливаться при ошибке удаления<small>Не найденные товары всегда записываются отдельно, очередь продолжается.</small></div></label></div>'+
      '<div class="view help hidden" id="help">'+
      '<h3>Как удалить товары по product_link</h3><ol>'+
      '<li>В Excel подготовьте один обязательный столбец <b>product_link</b>. Каждая строка — ссылка редактора товара этого сайта.</li>'+
      '<li>В разделе «Импорт Excel» выберите файл и лист, затем нажмите «Загрузить в очередь».</li>'+
      '<li>Нажмите «Тест: 1 товар», чтобы реально переместить в корзину только один товар из списка.</li>'+
      '<li>Нажмите «Начать» для остальных. Скрипт по очереди открывает product_link; поиск ID в списке товаров больше не используется.</li>'+
      '<li>В карточке сверяется ID из ссылки и тип записи product, затем нажимается кнопка #delete-action → Move to Trash.</li>'+
      '<li>После перехода проверяется результат. При перезагрузке ссылка Trash повторно не нажимается.</li>'+
      '<li>Не открывшиеся карточки фиксируются в «Не найдены». Отсутствующая кнопка или неподтверждённое удаление — в «Ошибки удаления».</li>'+
      '<li>После работы скачайте Excel-отчёт или TXT-журнал.</li>'+
      '</ol><h3>Безопасность</h3><p>Принимаются только редакторские ссылки того же домена. Некорректные URL и дубли пропускаются. Очереди сайтов независимы. Trash — перемещение в корзину, а не безвозвратное удаление.</p>'+
      '<p>Версии обновляются Tampermonkey через GitHub.</p></div>'+
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
    $('#stopOnError').checked=!!state.settings.stopOnError;
    $('#delay').onchange=e=>{
      if(!canEdit()){e.target.value=state.settings.delay;return;}
      state.settings.delay=Math.min(30,Math.max(1,Number(e.target.value)||2));
      e.target.value=state.settings.delay;persist();
    };
    $('#stopOnError').onchange=e=>{
      if(!canEdit()){e.target.checked=state.settings.stopOnError;return;}
      state.settings.stopOnError=e.target.checked;persist();
    };
    if (GM_getValue(PREFIX+':minimized',false)) {
      panel.classList.add('hidden');$('#pill').classList.remove('hidden');
    }
    if(!state.rows.length)setTab('import');
    else render();
  }
  function setTab(tab) {
    activeTab=tab;page=0;
    for(const name of ['queue','not_found','failed','log','import','help'])
      $('#'+name).classList.toggle('hidden',name!==tab);
    $$('[data-tab]').forEach(b=>b.classList.toggle('on',b.dataset.tab===tab));
    render();
  }
  function subset() {
    if (activeTab==='not_found') return state.rows.filter(x=>x.status==='not_found');
    if (activeTab==='failed') return state.rows.filter(x=>x.status==='failed');
    return state.rows.filter(x=>{
      const matches=filter==='all'||(filter==='pending' ? ['pending','opening','deleting','verifying'].includes(x.status) : x.status===filter);
      return matches && normalized(x.id+' '+x.productLink+' '+x.foundName).includes(normalized(search));
    });
  }
  function renderRows() {
    if(!root)return;
    if(!['queue','not_found','failed'].includes(activeTab))return;
    const rows=subset();
    const pages=Math.max(1,Math.ceil(rows.length/PAGESIZE));
    page=Math.min(page,pages-1);
    $('#pages').textContent=(page+1)+' / '+pages;
    $('#prev').disabled=page===0;$('#next').disabled=page>=pages-1;
    const target=activeTab==='queue' ? $('#rows') : activeTab==='not_found' ? $('#missingRows') : $('#failedRows');
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
        '<td><span class="badge '+(statuses[r.status]||'neutral')+'">'+(labels[r.status]||esc(r.status))+
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
    const done=count('done'),missing=count('not_found'),error=count('failed');
    const pending=count('pending')+count('opening')+count('deleting')+count('verifying');
    $('#total').textContent=state.rows.length;
    $('#done').textContent=done;
    $('#missing').textContent=missing;
    $('#errors').textContent=error;
    $('#bar').style.width=(state.rows.length?
      Math.round(100*(done+missing+error+count('skipped'))/state.rows.length):0)+'%';
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
      if(n==='failed')b.textContent='Ошибки удаления ('+error+')';
    });
    if(['queue','not_found','failed'].includes(activeTab))renderRows();
    if(activeTab==='log')renderLogs();
    $('.foot #prev').classList.toggle('hidden',!['queue','not_found','failed'].includes(activeTab));
    $('.foot #next').classList.toggle('hidden',!['queue','not_found','failed'].includes(activeTab));
    $('.foot #pages').classList.toggle('hidden',!['queue','not_found','failed'].includes(activeTab));
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