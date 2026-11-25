// ==UserScript==
// @name         Редактор курсов для SMM курсов
// @namespace    blackowl0192.woocommerce.course.products
// @version      1.5.2
// @description  Единый менеджер WooCommerce: массовое редактирование курсов из Excel (название, описание, цена, уроки, обложка) и дат публикации по ID. Очереди, отдельная панель управления, проверка, журнал и отчёты.
// @match        https://smmmaster.net/wp-admin/*
// @match        https://smmmaster.net/wp-login.php*
// @match        https://www.smmmaster.net/wp-admin/*
// @match        https://www.smmmaster.net/wp-login.php*
// @match        https://hottestcourse.com/wp-admin/*
// @match        https://hottestcourse.com/wp-login.php*
// @match        https://www.hottestcourse.com/wp-admin/*
// @match        https://www.hottestcourse.com/wp-login.php*
// @match        https://learnmarket.net/wp-admin/*
// @match        https://learnmarket.net/wp-login.php*
// @match        https://www.learnmarket.net/wp-admin/*
// @match        https://www.learnmarket.net/wp-login.php*
// @match        https://profcourse.com/wp-admin/*
// @match        https://profcourse.com/wp-login.php*
// @match        https://www.profcourse.com/wp-admin/*
// @match        https://www.profcourse.com/wp-login.php*
// @match        https://bestsocials.online/wp-admin/*
// @match        https://bestsocials.online/wp-login.php*
// @match        https://www.bestsocials.online/wp-admin/*
// @match        https://www.bestsocials.online/wp-login.php*
// @updateURL   https://raw.githubusercontent.com/blackowl0192/script_folder/main/smm-courses-editor.user.js
// @downloadURL https://raw.githubusercontent.com/blackowl0192/script_folder/main/smm-courses-editor.user.js
// @run-at       document-idle
// @noframes
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_addValueChangeListener
// @grant        GM_getTab
// @grant        GM_saveTab
// @grant        GM_registerMenuCommand
// @grant        unsafeWindow
// ==/UserScript==

/*
 * Independent script: does not replace the Tutor LMS Courses & Tutors Manager.
 * Import XLSX/XLS/CSV columns: ID | Name | Description | Regular price | lessons | img.
 * Excel ID is the WooCommerce product post ID, img is the 14-digit MEDIA FILE NAME,
 * NOT the internal WordPress attachment ID. Images must already be in Media Library.
 * Queue is shuffled once on import. Original Excel row numbers stay in all reports.
 * No Update is submitted unless field checks pass; saved data is re-read on a fresh page.
 * Use a single browser tab. Excel and queue stay in local Tampermonkey storage.
 */

(() => {
  'use strict';
  const SCRIPT_VERSION='1.5.2';
  const HOSTS = new Set(['smmmaster.net', 'hottestcourse.com', 'learnmarket.net', 'profcourse.com', 'bestsocials.online']);
  const SITE = location.hostname.replace(/^www\./, '');
  if (location.protocol !== 'https:' || !HOSTS.has(SITE)) return;
  const ORIGIN = location.origin;
  // A small emergency launcher remains visible if the main UI fails to initialize.
  // It is removed after the full panel has mounted successfully.
  const bootLauncher=document.createElement('button');
  bootLauncher.id='bwo-smm-manager-boot';
  bootLauncher.textContent='SMM Manager · v'+SCRIPT_VERSION+' · Открыть';
  bootLauncher.style.cssText='all:initial!important;box-sizing:border-box!important;position:fixed!important;right:16px!important;bottom:16px!important;z-index:2147483647!important;padding:12px 16px!important;border-radius:12px!important;background:#153d77!important;color:#fff!important;font:600 13px system-ui,sans-serif!important;cursor:pointer!important;box-shadow:0 5px 20px #001a3f75!important';
  document.documentElement.append(bootLauncher);
  function fatalLauncher(err){
    const message=String(err?.message||err||'Неизвестная ошибка');
    console.error('[SMM Course Manager] Сбой панели:',err);
    bootLauncher.textContent='SMM Manager v'+SCRIPT_VERSION+' · Ошибка запуска';
    bootLauncher.title='Версия '+SCRIPT_VERSION+' — '+message;
    bootLauncher.onclick=()=>{
      if(panel&&root){panel.classList.remove('hidden');root.getElementById('pill')?.classList.add('hidden');}
      else alert('Не удалось открыть панель SMM Manager v'+SCRIPT_VERSION+'.\n'+message+'\nПроверьте ошибки в консоли браузера (F12).');
    };
  }
  bootLauncher.onclick=()=>{if(panel&&root){panel.classList.remove('hidden');root.getElementById('pill')?.classList.add('hidden');}};
  // Excel is optional at startup: an unavailable CDN must never hide the manager UI.
  const XLSX_URL='https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
  let sheetLib=null, sheetLoading=null;
  function getXLSX(){
    return sheetLib || (typeof unsafeWindow!=='undefined' && unsafeWindow.XLSX) || null;
  }
  function ensureXLSX(){
    if(getXLSX())return Promise.resolve(getXLSX());
    if(sheetLoading)return sheetLoading;
    sheetLoading=new Promise((resolve,reject)=>{
      const script=document.createElement('script');
      let completed=false;
      const timer=setTimeout(()=>finish(new Error('Не удалось загрузить SheetJS за 25 секунд. Проверьте подключение к cdn.sheetjs.com или используйте CSV.')),25000);
      function finish(err){
        if(completed)return;
        completed=true;clearTimeout(timer);
        if(err)reject(err);
        else {
          const lib=getXLSX();
          if(lib){sheetLib=lib;resolve(lib);}
          else reject(new Error('SheetJS загрузился, но XLSX недоступен. Можно импортировать CSV.'));
        }
      }
      script.src=XLSX_URL;script.async=true;script.crossOrigin='anonymous';
      script.onload=()=>finish(null);
      script.onerror=()=>finish(new Error('Не удалось загрузить Excel-библиотеку. Можно использовать CSV.'));
      (document.head||document.documentElement).append(script);
    }).catch(err=>{sheetLoading=null;throw err;});
    return sheetLoading;
  }
  const SMM_PARAMS = new URLSearchParams(location.search);
  const IS_DASHBOARD = SMM_PARAMS.get('smm_manager') === 'dashboard';
  const COURSE_WORKER_SESSION = 'smm.manager.course.worker.' + SITE;
  const WORKER_PREFIX = 'smm-manager-courses-' + SITE + '-w';
  const workerRequested = Number(SMM_PARAMS.get('smm_worker') || 0);
  const validWorker = n => Number.isInteger(n) && n >= 1 && n <= 5;
  if(SMM_PARAMS.get('smm_manager')==='course-worker' && validWorker(workerRequested)){
    window.name=WORKER_PREFIX+workerRequested;
    sessionStorage.setItem(COURSE_WORKER_SESSION,String(workerRequested));
  }
  const namedWorker = window.name.startsWith(WORKER_PREFIX) ? Number(window.name.slice(WORKER_PREFIX.length)) : 0;
  const MASTER_KEY = `bwo.wc.course.product.v1.${SITE}`;
  const rememberedWorker = sessionStorage.getItem(COURSE_WORKER_SESSION)===String(namedWorker);
  // WordPress may redirect admin from example.com to www.example.com, where sessionStorage
  // is a separate origin. Re-identify the named tab only if this exact worker is active.
  const candidatePool = validWorker(namedWorker) ? GM_getValue(MASTER_KEY+'.pool',null) : null;
  const candidateShard = validWorker(namedWorker) ? GM_getValue(MASTER_KEY+'.worker.'+namedWorker,null) : null;
  const matchingActiveWorker = candidatePool?.workers?.includes(namedWorker) &&
    candidateShard?.run===candidatePool.run && candidateShard.running===true;
  const WORKER_ID = !IS_DASHBOARD && validWorker(namedWorker) &&
    (rememberedWorker||matchingActiveWorker) ? namedWorker : 0;
  if(WORKER_ID && !rememberedWorker)sessionStorage.setItem(COURSE_WORKER_SESSION,String(WORKER_ID));
  const IS_COURSE_WORKER = WORKER_ID>0;
  const POOL_KEY = MASTER_KEY + '.pool';
  const STOP_ALL_KEY = MASTER_KEY + '.emergency-stop';
  const workerKey = id => MASTER_KEY+'.worker.'+id;
  const KEY = IS_COURSE_WORKER ? workerKey(WORKER_ID) : MASTER_KEY;
  const LEASE = `${KEY}.lease`;
  const TICKET = `${KEY}.nav`;
  // A stable identity that survives WordPress POST/redirect for each named worker tab.
  const WORKER_TAB_SESSION = `${KEY}.tab-identity`;
  const COMMAND = `${KEY}.command`;
  const MODE_KEY = `${MASTER_KEY}.ui.mode`;
  const ALIAS = {id:['id','postid','productid'],name:['name','title','productname'],description:['description','content','info'],price:['regularprice','price'],lessons:['lessons','lesson'],img:['img','image','photo']};
  const STATUSES = {pending:'В очереди',opening:'Открывается',filling:'Заполняется',saving:'Сохраняется',verifying:'Проверяется',done:'Выполнено',review:'Проверить',error:'Ошибка',skipped:'Пропущено'};
  const ACTIVE = new Set(['opening','filling','saving','verifying']);
  const sleep = ms => new Promise(resolve=>setTimeout(resolve,ms));
  const uid = () => crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`;
  const DOC = uid();
  const error = (message,fatal=false) => Object.assign(new Error(message),{fatal});
  const STOP = Symbol('stopped');
  const esc = value => String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const empty = () => ({version:1,rows:[],logs:[],filename:'',run:'',index:-1,phase:'idle',running:false,paused:false,owner:'',single:false,stopAfterVerify:false,settings:{delay:3,stopOnError:true,workers:1},updated:''});
  let state = GM_getValue(KEY,null)||empty();

  // Every worker has an independent state key; the dashboard only merges read-only snapshots.
  function currentPool(){return GM_getValue(POOL_KEY,null);}
  function emergencyStopFor(run){
    const stop=GM_getValue(STOP_ALL_KEY,null);
    return !!(run && stop && stop.run===run);
  }
  function activeLease(id){const l=GM_getValue(workerKey(id)+'.lease',null);return !!(l && l.expires>Date.now());}
  function poolBusy(pool=currentPool()){
    return !!(pool?.workers?.length && (pool.workers.some(activeLease) ||
      pool.workers.some(id=>{
        const x=GM_getValue(workerKey(id),null);
        if(x?.run!==pool.run)return false;
        const recentlyActive=Date.now()-Date.parse(x.updated||0)<120000;
        return (Date.now()-pool.createdAt<20000 || recentlyActive && x.running) &&
          x.rows?.some(r=>r.status==='pending'||ACTIVE.has(r.status));
      })));
  }
  function mergedCourseState(){
    if(IS_COURSE_WORKER)return GM_getValue(KEY,empty());
    const base=GM_getValue(MASTER_KEY,empty());
    const pool=currentPool();
    if(!pool?.workers?.length)return base;
    const rows=base.rows.map(r=>({...r}));
    const positions=new Map(rows.map((r,i)=>[r.key,i]));
    const logs=[...(base.logs||[])],workers=[];
    for(const id of pool.workers){
      const sub=GM_getValue(workerKey(id),null);
      if(!sub||sub.run!==pool.run)continue;
      for(const r of sub.rows||[]){
        const ix=positions.get(r.key);
        if(ix!==undefined)rows[ix]={...r,worker:id};
      }
      logs.push(...(sub.logs||[]).map(x=>({...x,worker:id})));
      workers.push({id,live:activeLease(id),running:!!sub.running,
        done:sub.rows.filter(r=>r.status==='done').length,
        total:sub.rows.length,paused:!!sub.paused});
    }
    logs.sort((a,b)=>String(a.time||'').localeCompare(String(b.time||'')));
    return {...base,rows,logs:logs.slice(-2000),running:workers.some(w=>w.live),
      paused:workers.some(w=>w.paused),poolWorkers:workers,
      poolStarting:Date.now()-pool.createdAt<20000 && workers.some(w=>!w.live&&w.total&&!w.done)};
  }
  function consolidatePool(recover=false){
    const pool=currentPool();
    if(!pool)return GM_getValue(MASTER_KEY,empty());
    // Manual recovery is permitted ONLY after all worker leases have expired,
    // with a 20-second startup grace. It must not interrupt a live POST.
    if(poolBusy(pool) && !(recover && !pool.workers.some(activeLease) &&
      Date.now()-pool.createdAt>=20000))return null;
    const result=mergedCourseState();
    for(const row of result.rows){
      if(['saving','verifying'].includes(row.status)){
        if(row.expected && row.saveAt){
          row.status='verifying';
          row.message='Предыдущий Update мог сохраниться. При продолжении проверим результат с новой страницы без повторной отправки.';
        }else{
          row.status='review';
          row.message='Сохранение не подтверждено и нет достаточных данных для сверки. Проверьте вручную; повторный Update запрещён.';
        }
      }else if(['opening','filling'].includes(row.status)){
        row.status='pending';
        row.message='Прервано до отправки Update, возвращено в очередь.';
      }
      delete row.worker;
    }
    delete result.poolWorkers;delete result.poolStarting;
    result.running=false;result.paused=true;result.index=-1;result.phase='idle';result.owner='';
    result.stopAfterVerify=false;result.updated=new Date().toISOString();
    GM_setValue(MASTER_KEY,result);
    GM_setValue(POOL_KEY,null);
    return result;
  }
  let tabId,root,panel,workbook=null,importFile='',working=false,leaving=false,halted=false,single=false,leaseTimer;
  let postedFromThisDocument=false;
  let tab='import',page=0,filter='all',search='',maximized=false;
  const $ = id=>root.getElementById(id);

  function cleanHeader(value){return String(value??'').toLowerCase().replace(/[\s_\-]+/g,'').trim();}
  function productURL(id){return `${ORIGIN}/wp-admin/post.php?post=${id}&action=edit`;}
  function productIDFromURL(value){
    let url;try{url=new URL(String(value),ORIGIN);}catch{return '';}
    return url.origin===ORIGIN&&url.pathname==='/wp-admin/post.php'&&url.searchParams.get('action')==='edit'&&/^\d+$/.test(url.searchParams.get('post')||'')?url.searchParams.get('post'):'';
  }
  function normalizePrice(value){
    // WooCommerce expects an unformatted decimal, not a currency string.
    let s=String(value??'').trim();
    if(!s)throw error('Regular price не заполнена.');
    s=s.replace(/\p{Sc}/gu,'').replace(/\b(?:EUR|USD|GBP|RON|MDL|RUB|CAD|AUD)\b/gi,'')
      .replace(/[\s\u00a0\u202f\u2007\u2009]/g,'');
    if(!s||!/^[0-9.,]+$/.test(s))throw error(`Некорректная цена: ${JSON.stringify(value)}.`);
    const dots=(s.match(/\./g)||[]).length,commas=(s.match(/,/g)||[]).length;
    if(dots && commas){
      const dec=s.lastIndexOf('.')>s.lastIndexOf(',')?'.':',';
      const grp=dec==='.'?',':'.';const groups=s.split(dec);
      if(groups.length!==2||!/^\d{1,2}$/.test(groups[1])||!new RegExp('^\\d{1,3}(?:\\'+grp+'\\d{3})+$').test(groups[0]))throw error('Неоднозначные разделители цены.');
      s=groups[0].split(grp).join('')+'.'+groups[1];
    }else{
      if(!/^\d+(?:[.,]\d{1,2})?$/.test(s))throw error('Цена должна содержать максимум две цифры после запятой. Разделители тысяч без десятичной части не поддерживаются.');
      s=s.replace(',','.');
    }
    const parts=s.split('.');if(parts[0].length>9)throw error('Слишком большая цена.');
    return `${String(Number(parts[0]))}.${(parts[1]||'').padEnd(2,'0')}`;
  }
  // Keep imported prices in normalized dot form for comparisons/storage.
  // These WooCommerce shops require a decimal comma in the editable price field.
  function wooPriceForInput(value){
    return normalizePrice(value).replace('.',',');
  }
  function normalizeLessons(value){
    const s=String(value??'').trim();
    if(!/^\d+(?:\.0+)?$/.test(s)||!Number.isSafeInteger(Number(s))||Number(s)<1||Number(s)>100000)
      throw error('lessons: необходимо положительное целое число.');
    return String(Number(s));
  }
  function normalizeImg(value){
    const s=String(value??'').trim();
    if(!/^\d{14}$/.test(s))throw error('img: требуется 14 цифр — имя файла в Media Library, не attachment ID. В Excel укажите текстовый формат.');
    return s;
  }
  function slugFromTitle(name){
    const map={а:'a',б:'b',в:'v',г:'g',д:'d',е:'e',ё:'e',ж:'zh',з:'z',и:'i',й:'y',к:'k',л:'l',м:'m',н:'n',о:'o',п:'p',р:'r',с:'s',т:'t',у:'u',ф:'f',х:'kh',ц:'ts',ч:'ch',ш:'sh',щ:'shch',ъ:'',ы:'y',ь:'',э:'e',ю:'yu',я:'ya',і:'i',ї:'yi',є:'ye'};
    return String(name).trim().toLowerCase().replace(/[а-яёіїє]/g,c=>map[c]).normalize('NFKD')
      .replace(/[\u0300-\u036f]/g,'').replace(/&/g,' and ').replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
  }
  function validateHTML(input){
    const html=String(input??'');const doc=new DOMParser().parseFromString(html,'text/html');
    if(doc.querySelector('script,iframe,object,embed,form,base,meta,link'))throw error('Description содержит запрещённые HTML-элементы (script/iframe/form и др.).');
    for(const node of doc.body.querySelectorAll('*'))for(const a of [...node.attributes]){
      if(/^on/i.test(a.name)||(/^(href|src|action|xlink:href)$/i.test(a.name)&&/^\s*(?:javascript|vbscript|data):/i.test(a.value)))throw error('Description содержит небезопасный HTML-атрибут.');
    }
    if(!html.trim())throw error('Description не заполнено.');
    return html;
  }
  function richSignature(html){
    const doc=new DOMParser().parseFromString(String(html??''),'text/html');
    const walk=node=>{
      if(node.nodeType===3)return node.nodeValue.replace(/\u00a0/g,' ').replace(/\s+/g,' ');
      if(node.nodeType!==1)return '';
      const tag=({b:'strong',i:'em'}[node.tagName.toLowerCase()]||node.tagName.toLowerCase());
      // Content is audited by text and structural tags, but UI-only style changes are ignored.
      if(['style'].includes(tag))return '';
      const href=tag==='a'?`(${node.getAttribute('href')||''})`:'';
      return `<${tag}${href}>${[...node.childNodes].map(walk).join('').trim()}</${tag}>`;
    };
    return [...doc.body.childNodes].map(walk).join('').replace(/>\s+</g,'><').trim();
  }
  function visible(el){return !!(el?.isConnected&&el.getClientRects().length&&getComputedStyle(el).visibility!=='hidden');}
  function one(selector,scope=document,required=true){
    const found=[...scope.querySelectorAll(selector)].filter(visible);
    if(found.length===1)return found[0];
    if(required)throw error(`Ожидался один элемент ${selector}, найдено ${found.length}.`);
    return null;
  }
  function authCheck(){if(location.pathname.includes('wp-login.php')||document.querySelector('#loginform'))throw error('Необходимо авторизоваться в WordPress.',true);}
  function persist(){state.updated=new Date().toISOString();GM_setValue(KEY,state);}
  function log(level,message,row){
    state.logs.push({time:new Date().toISOString(),level,row:row?.excelRow||'',id:row?.id||'',stage:state.phase,message:String(message).slice(0,1800)});
    if(state.logs.length>2000)state.logs.splice(0,state.logs.length-2000);
    persist();render();
  }
  function stage(r,status,message){guard(status==='saving'||status==='verifying');r.status=status;state.phase=status;r.message=message;log('info',message,r);}
  function owned(){const lease=GM_getValue(LEASE,null);return !!(lease&&lease.tab===tabId&&lease.expires>Date.now());}
  function guard(allowStopped=false){
    authCheck();
    if(IS_COURSE_WORKER && currentPool()?.run!==state.run)
      throw error('Этот рабочий поток принадлежит старой очереди и остановлен.',true);
    // Cross-tab emergency barrier. Reads GM storage synchronously before EVERY
    // destructive action, even if a worker missed a command during navigation.
    if(IS_COURSE_WORKER && emergencyStopFor(state.run))halted=true;
    if(!owned())throw error('Другая вкладка управляет очередью либо завершилась аренда.',true);
    if(halted&&!allowStopped)throw STOP;
  }
  async function waitFor(fn,label,timeout=25000,allowStopped=false){
    const started=Date.now();
    while(Date.now()-started<timeout){guard(allowStopped);const result=fn();if(result)return result;await sleep(200);}
    throw error(`Время ожидания истекло: ${label}`);
  }
  function nativeValue(el,value){
    const prop=el.tagName==='TEXTAREA'?unsafeWindow.HTMLTextAreaElement.prototype:unsafeWindow.HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(prop,'value').set.call(el,String(value));
    el.dispatchEvent(new unsafeWindow.Event('input',{bubbles:true}));
    el.dispatchEvent(new unsafeWindow.Event('change',{bubbles:true}));
  }
  async function fill(selector,value){
    guard();const el=one(selector);if(el.disabled||el.readOnly)throw error(`Поле ${selector} заблокировано.`);
    el.focus();nativeValue(el,value);await sleep(150);guard();
    if(el.value!==String(value))throw error(`Поле ${selector} сбросилось.`);
  }
  function readSlug(){
    // Classic WordPress abbreviates the permalink after OK. Read the full value.
    const input=document.querySelector('#new-post-slug');
    if(input&&visible(input))return input.value.trim();
    const full=document.querySelector('#editable-post-name-full');
    if(full?.textContent?.trim())return full.textContent.trim();
    const hidden=document.querySelector('#post_name');
    if(hidden?.value?.trim())return hidden.value.trim();
    const editable=document.querySelector('#editable-post-name');
    const displayed=editable?.textContent?.trim()||'';
    // Ellipsis is never sufficient to verify the entire slug.
    return /[\u2026]|\.\.\./.test(displayed)?'':displayed;
  }
  async function writeSlug(slug){
    if(readSlug()===slug)return;
    const edit=await waitFor(()=>one('#edit-slug-buttons .edit-slug',document,false),'кнопка Edit permalink',12000);
    guard();edit.click();
    await waitFor(()=>one('#new-post-slug',document,false),'редактор slug',12000);
    await fill('#new-post-slug',slug);
    const save=one('#edit-slug-buttons .save');guard();save.click();
    try{
      await waitFor(()=>readSlug()===slug&&!one('#edit-slug-buttons .save',document,false),'применение slug',20000);
    }catch(ex){
      if(!/Время ожидания истекло/.test(ex.message||''))throw ex;
      const full=document.querySelector('#editable-post-name-full')?.textContent?.trim()||'';
      const hidden=document.querySelector('#post_name')?.value?.trim()||'';
      const display=document.querySelector('#editable-post-name')?.textContent?.trim()||'';
      throw error('WordPress не подтвердил slug. Ожидалось: '+slug+'; полный slug: '+(full||'—')+'; post_name: '+(hidden||'—')+'; краткое отображение: '+(display||'—')+'. Update не выполнен.');
    }
    log('info','Slug подтверждён: '+slug,state.rows[state.index]);
  }
  function tiny(){return unsafeWindow.tinymce?.get('content');}
  async function readDescription(allowStopped=false){
    const ed=tiny();if(ed&&ed.initialized&&!ed.isHidden())return ed.getContent();
    const el=document.querySelector('#content');
    if(el&&visible(el))return el.value;
    const text=document.querySelector('#content-html');
    if(text&&visible(text)){
      guard(allowStopped);text.click();await waitFor(()=>visible(document.querySelector('#content')),'Text редактор',12000,allowStopped);
      return document.querySelector('#content').value;
    }
    if(el)return el.value;
    throw error('Не найден редактор Description (#content).');
  }
  async function writeDescription(raw){
    const ta=document.querySelector('#content');if(!ta)throw error('Не найдено поле #content.');
    const text=one('#content-html');guard();text.click();
    await waitFor(()=>visible(document.querySelector('#content')),'переключение Description → Text',12000);
    await fill('#content',raw);
    // WordPress imports text into TinyMCE when switching to Visual.
    const visual=one('#content-tmce');guard();visual.click();
    const ed=await waitFor(()=>{const e=tiny();return e?.initialized&&!e.isHidden()?e:null;},'переключение Description → Visual',25000);
    let html='',transformationLogged=false;
    for(let i=0;i<3;i++){
      guard();html=ed.getContent();
      if(!html.trim())throw error('Visual очистил Description. Update не выполнен.');
      // TinyMCE legitimately normalizes text/HTML when switching from Text to Visual:
      // paragraphs, whitespace, entities, and inline formatting can be rewritten.
      // Treat the editor's final nonempty HTML as the value to save and verify.
      if(!transformationLogged && richSignature(html)!==richSignature(raw)){
        log('info','Visual преобразовал Description; итоговый HTML редактора будет использован для сохранения и проверки.',state.rows[state.index]);
        transformationLogged=true;
      }
      ed.setDirty?.(true);ed.save();
      ed.fire('change');
      await sleep(250);
      const current=tiny();if(current?.initialized&&!current.isHidden()&&richSignature(current.getContent())===richSignature(html))break;
      if(i===2)throw error('Visual не удерживает новое описание. Update не выполнен.');
    }
    return html;
  }
  function imageStem(value){
    if(!value)return '';
    try{
      return decodeURIComponent(new URL(String(value),ORIGIN).pathname.split('/').pop()||'')
        .replace(/\.(?:jpe?g|png|webp|gif|avif)$/i,'')
        .replace(/(?:-(?:\d+x\d+|scaled|rotated))+$/gi,'').toLowerCase();
    }catch{return '';}
  }
  function imageMatches(value,img){return !!img&&imageStem(value)===img.toLowerCase();}
  function thumbnailID(){return String(document.querySelector('input#_thumbnail_id,input[name="_thumbnail_id"]')?.value||'').trim();}
  function imagePreview(){
    const rootDiv=document.querySelector('#postimagediv');if(!rootDiv)return '';
    const images=[...rootDiv.querySelectorAll('img[src]')].filter(el=>!el.closest('.spinner'));
    if(images.length===1)return images[0].getAttribute('src')||'';
    return images.find(el=>el.closest('#set-post-thumbnail'))?.getAttribute('src')||'';
  }
  function findPhotoItem(item,img){
    const image=item.querySelector('img[src]');
    const fileStem=imageStem(image?.getAttribute('src'));
    const filename=item.querySelector('.filename')?.textContent?.trim()||'';
    const exactTitle=(item.getAttribute('aria-label')||item.getAttribute('title')||'').trim();
    return fileStem===img||imageStem(filename)===img||exactTitle===img;
  }
  async function writeImage(img,row){
    guard();
    if(imageMatches(imagePreview(),img)){
      const existing=thumbnailID();log('info',`Обложка ${img} уже установлена. Повторный выбор не требуется (attachment ID=${existing||'не указан'}).`,row);
      return {img,attachmentId:existing};
    }
    const remove=one('#remove-post-thumbnail',document,false);
    if(remove){log('info','Удаляем старое изображение из формы товара.',row);guard();remove.click();}
    const link=await waitFor(()=>one('#set-post-thumbnail',document,false),'Set product image',15000);
    log('info',`Открываем Media Library; ищем имя файла img=${img}.`,row);
    guard();link.click();
    const modal=await waitFor(()=>one('.media-modal',document,false),'Media Library',18000);
    let search=one('#media-search-input',modal,false);
    if(!search){
      const library=[...modal.querySelectorAll('.media-router button,.media-router a,.media-menu button')]
        .find(el=>visible(el)&&/media library|медиатека|библиотека файлов/i.test(el.textContent));
      if(library){guard();library.click();}
      search=await waitFor(()=>one('#media-search-input',modal,false),'поиск Media Library',18000);
    }
    search.focus();nativeValue(search,img);
    search.dispatchEvent(new unsafeWindow.KeyboardEvent('keyup',{bubbles:true,key:'Enter'}));
    await sleep(900);
    let stable='',since=0;
    const selected=await waitFor(()=>{
      if(one('#media-search-input',modal,false)?.value!==img)return null;
      if([...modal.querySelectorAll('.spinner.is-active,.media-progress-bar')].some(visible)){since=0;return null;}
      const items=[...modal.querySelectorAll('.attachments .attachment')].filter(visible);
      const matches=items.filter(el=>findPhotoItem(el,img));
      if(matches.length>1)throw error(`Найдено несколько точных файлов img=${img}. Update не выполнен.`);
      if(matches.length!==1){since=0;return null;}
      const el=matches[0],signature=(el.getAttribute('data-id')||'')+'|'+(el.querySelector('img')?.src||'');
      if(signature!==stable||!since){stable=signature;since=Date.now();return null;}
      return Date.now()-since>800?el:null;
    },`точное изображение img=${img}`,35000);
    const attachment=selected.getAttribute('data-id')||'';
    if(!/^\d+$/.test(attachment))throw error('Медиатека не вернула внутренний attachment ID выбранной фотографии.');
    if(selected.getAttribute('aria-checked')!=='true'){
      const preview=selected.querySelector('.attachment-preview');if(!preview)throw error('У фотографии нет кнопки выбора .attachment-preview.');
      guard();preview.click();
    }
    await waitFor(()=>selected.getAttribute('aria-checked')==='true','выбор файла',12000);
    const selectedItems=[...modal.querySelectorAll('.attachment[aria-checked="true"]')];
    if(selectedItems.length!==1||selectedItems[0]!==selected)throw error('Медиатека выбрала несколько фотографий либо не тот файл.');
    const set=one('.media-button-select',modal);
    if(set.disabled||set.getAttribute('aria-disabled')==='true')throw error('Set product image недоступна.');
    log('info',`Выбрано ${img}; WordPress attachment ID=${attachment}. Нажимаем Set product image.`,row);
    guard();set.click();
    await waitFor(()=>!visible(modal),'закрытие Media Library',20000);
    let confirmed='';
    try{
      confirmed=await waitFor(()=>{
        const id=thumbnailID(),src=imagePreview();
        if(id&&id!=='0'&&id!=='-1'&&id!==attachment)return null;
        if(id===attachment&&(!src||imageMatches(src,img)))return 'id='+attachment;
        if(!id&&imageMatches(src,img))return 'preview='+src;
        return null;
      },'подтверждение обложки товара',30000);
    }catch(e){
      if(!/Время ожидания истекло/.test(e.message||''))throw e;
      throw error(`Новая обложка не подтверждена. img=${img}, attachment=${attachment}, текущее _thumbnail_id=${thumbnailID()}, preview=${imagePreview()}. Update не выполнен.`);
    }
    log('ok',`Обложка ${img} выбрана и подтверждена (${confirmed}).`,row);
    return {img,attachmentId:attachment};
  }
  function lessonsField(){
    const byID=one('#acf-field_6502c6f7841f9',document,false);
    if(byID)return byID;
    const byName=one('input[name="acf[field_6502c6f7841f9]"]',document,false);
    if(byName)return byName;
    // Other sites can have a different ACF field key for the same "lessons" field.
    const candidates=[...document.querySelectorAll('.acf-field[data-name="lessons"] input[type="number"],.acf-field[data-name="lessons"] input[type="text"]')].filter(visible);
    if(candidates.length===1)return candidates[0];
    throw error(`Поле ACF lessons не найдено однозначно (вариантов: ${candidates.length}).`);
  }
  function readDownloadable(){return !!one('#_downloadable',document,false)?.checked;}
  async function disableDownloadable(){
    guard();const el=one('#_downloadable');
    if(el.checked){el.click();await sleep(220);guard();}
    if(el.checked)throw error('Не удалось снять флажок Downloadable.');
  }
  async function editorReady(allowStopped=false){
    await waitFor(()=>one('input#title[name="post_title"]',document,false)
      &&one('#content-html',document,false)
      &&one('#_regular_price',document,false)
      &&one('#_downloadable',document,false)
      &&one('#publish',document,false)
      &&(()=>{try{return !!lessonsField();}catch{return false;}})(),
      'поля WooCommerce и ACF',50000,allowStopped);
    await sleep(400);guard(allowStopped);
  }
  async function snapshot(allowStopped=false){
    return {name:one('input#title[name="post_title"]').value,slug:readSlug(),description:await readDescription(allowStopped),
      price:one('#_regular_price').value,lessons:lessonsField().value,
      downloadable:readDownloadable(),image:imagePreview(),attachmentId:thumbnailID()};
  }

  // Read-only verification from WordPress HTML, without reloading the worker tab
  // or submitting a second Update. All fields must be present before marking done.
  async function fetchSavedSnapshot(row){
    const requested=new URL(productURL(row.id));
    const response=await fetch(requested.href,{
      method:'GET',credentials:'same-origin',redirect:'follow',cache:'no-store',
      signal:AbortSignal.timeout(45000)
    });
    const target=new URL(response.url||requested.href,ORIGIN);
    if(target.protocol!=='https:'||target.hostname.replace(/^www\./,'')!==SITE)
      throw error('Проверка перенаправлена на другой сайт. Товар не подтверждён.');
    if(!response.ok||/\/wp-login\.php$/.test(target.pathname))
      throw error('Не получена страница товара: HTTP '+response.status+'.');
    const doc=new DOMParser().parseFromString(await response.text(),'text/html');
    if(doc.querySelector('#loginform'))throw error('Для проверки требуется заново войти в WordPress.');
    const form=doc.querySelector('form#post');
    if(!form)throw error('WordPress не вернул форму редактирования товара.');
    const formId=String(form.querySelector('[name="post_ID"]')?.value||'').trim();
    if(formId!==row.id)throw error('В ответе другой товар: '+formId+' вместо '+row.id+'.');
    const required=(selector,label)=>{
      const input=form.querySelector(selector)||doc.querySelector(selector);
      if(!input)throw error('В ответе отсутствует поле '+label+'; не могу подтвердить сохранение.');
      return input;
    };
    const slugValues=[
      form.querySelector('#post_name')?.value?.trim(),
      doc.querySelector('#editable-post-name-full')?.textContent?.trim(),
      doc.querySelector('#new-post-slug')?.value?.trim(),
      doc.querySelector('#editable-post-name')?.textContent?.trim()
    ];
    const slug=slugValues.find(x=>x&&!/[\u2026]|\.\.\./.test(x))||'';
    if(!slug)throw error('WordPress не вернул полный slug для проверки.');
    const lessons=form.querySelector('#acf-field_6502c6f7841f9') ||
      form.querySelector('input[name="acf[field_6502c6f7841f9]"]') ||
      form.querySelector('.acf-field[data-name="lessons"] input[type="number"],.acf-field[data-name="lessons"] input[type="text"]');
    if(!lessons)throw error('WordPress не вернул поле ACF lessons.');
    const imageBlock=doc.querySelector('#postimagediv');
    const image=imageBlock?.querySelector('#set-post-thumbnail img[src],img[src]')?.getAttribute('src')||'';
    const thumbnail=form.querySelector('input#_thumbnail_id,input[name="_thumbnail_id"]')||
      doc.querySelector('input#_thumbnail_id,input[name="_thumbnail_id"]');
    const attachmentId=String(thumbnail?.value||'').trim();
    if(row.expected?.img&&!attachmentId&&!image)
      throw error('Не удалось прочитать обложку из сохранённой страницы.');
    const actual={
      name:required('input#title[name="post_title"]','Name').value,
      slug,
      description:required('textarea#content[name="content"]','Description').value,
      price:required('#_regular_price','Regular price').value,
      lessons:lessons.value,
      downloadable:!!form.querySelector('#_downloadable')?.checked,
      image,attachmentId
    };
    log('info','Свежая страница товара получена: HTTP '+response.status+
      ', ID '+row.id+', attachment '+(attachmentId||'по изображению')+'.',row);
    return actual;
  }
  async function verifySavedRow(row){
    if(!row.expected||!row.saveAt)
      throw error('Нет данных прежнего Update. Повторное сохранение запрещено.');
    stage(row,'verifying','Сверяю сохранённый товар через GET WordPress, без перезагрузки и без Update.');
    let lastError='';
    for(let attempt=1;attempt<=2;attempt++){
      guard(true);
      try{
        const actual=await fetchSavedSnapshot(row);
        row.actual=actual;row.checked=new Date().toISOString();persist();
        const diff=differences(actual,row.expected);
        if(!diff.length){
          row.status='done';row.message='Все поля подтверждены с сервера после Update';
          log('ok',row.message,row);
          return true;
        }
        lastError='Несовпадения после Update: '+diff.join(', ');
      }catch(ex){lastError=ex.message||String(ex);}
      log('warn','Проверка '+attempt+'/2: '+lastError,row);
      if(attempt<2)await sleep(1200);
    }
    row.status='review';row.message=lastError+'; повторный Update не выполнялся.';
    log('error',row.message,row);
    pauseOtherCourseWorkers();
    return false;
  }
  function differences(actual,expected){
    const diff=[];
    if(actual.name!==expected.name)diff.push('Name');
    if(actual.slug!==expected.slug)diff.push('slug');
    if(richSignature(actual.description)!==richSignature(expected.description)){
      const normalizedText=h=>new DOMParser().parseFromString(String(h??''),'text/html').body.textContent.replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();
      if(normalizedText(actual.description)!==normalizedText(expected.description))diff.push('Description / HTML');
      else{
        // If markup has been significantly modified, do not silently accept structural loss.
        const tags=h=>[...new DOMParser().parseFromString(String(h),'text/html').body.querySelectorAll('table,thead,tbody,tr,td,th,ul,ol,li,h1,h2,h3,h4,h5,h6,strong,a')].map(x=>x.tagName).join(',');
        if(tags(actual.description)!==tags(expected.description))diff.push('Description structure');
      }
    }
    try{if(normalizePrice(actual.price)!==normalizePrice(expected.price))diff.push('Regular price');}catch{diff.push('Regular price');}
    if(String(Number(actual.lessons))!==String(Number(expected.lessons)))diff.push('lessons');
    if(actual.downloadable)diff.push('Downloadable');
    if(expected.img){
      const exactPreview=imageMatches(actual.image,expected.img);
      const idMatch=!!(expected.attachmentId&&actual.attachmentId===expected.attachmentId);
      if(!exactPreview&&!idMatch)diff.push('Product image');
    }
    return diff;
  }
  async function acquire(){
    const l=GM_getValue(LEASE,null);
    if(l&&l.expires>Date.now()&&l.tab!==tabId)throw error('Очередь запущена в другой вкладке. Остановите её там или дождитесь освобождения.',true);
    GM_setValue(LEASE,{tab:tabId,expires:Date.now()+20000});await sleep(170);
    if(!owned())throw error('Не удалось получить управление очередью.',true);
    clearInterval(leaseTimer);
    leaseTimer=setInterval(()=>{
      if(!owned()){halted=true;clearInterval(leaseTimer);return;}
      GM_setValue(LEASE,{tab:tabId,expires:Date.now()+20000});
    },4000);
  }
  function release(){clearInterval(leaseTimer);if(owned())GM_setValue(LEASE,null);}
  function nav(url,kind='open'){
    guard(true);
    const r=state.rows[state.index];
    sessionStorage.setItem(TICKET,JSON.stringify({tab:tabId,run:state.run,id:r?.id||'',kind,at:Date.now()}));
    if(IS_COURSE_WORKER)window.name=WORKER_PREFIX+WORKER_ID;
    persist();leaving=true;
    const resolved=new URL(url,ORIGIN);
    // A shop can canonically redirect between the www and non-www admin host.
    if(resolved.protocol!==location.protocol || resolved.hostname.replace(/^www\./,'')!==SITE)
      throw error('Запрещён переход на другой сайт.',true);
    resolved.host=location.host;
    if(kind==='verify'||postedFromThisDocument){
      // We have already sent this page's WordPress form via fetch.
      // Classic Editor's pending-changes guard would otherwise block
      // the next course even though Update was submitted and checked.
      // Disable only the core WordPress namespace, not unrelated handlers.
      try{unsafeWindow.jQuery?.(unsafeWindow).off('beforeunload.edit-post');}catch(ex){
        console.warn('[SMM Course Manager] WP beforeunload cleanup:',ex);
      }
    }
    // Browser navigation can be cancelled by plugin/editor beforeunload listeners.
    // Report the block instead of silently leaving the worker marked as running.
    setTimeout(()=>{
      if(!leaving||!state.running)return;
      const current=state.rows[state.index];
      state.running=false;state.paused=true;state.stopAfterVerify=true;
      leaving=false;persist();release();working=false;
      log('error','Переход '+kind+' не завершён за 18 секунд. Возможен диалог «Покинуть страницу» или блокировка браузера. '+ 
        'Товар '+(current?.id||'')+' оставлен на проверку; автоматический Update повторно не отправляется.',current);
    },18000);
    log('info','Переход '+kind+' в той же рабочей вкладке: '+resolved.pathname+resolved.search,r);
    if(resolved.href===location.href)location.reload();else location.assign(resolved.href);
  }

  // Submit WordPress's existing Classic Editor form without relying on a browser
  // navigation from button.click(). One POST only, then a fresh GET/DOM verification.
  // Uncertain responses are NEVER retried automatically.
  async function submitWooUpdate(form,button,row){
    if(form.method.toLowerCase()!=='post')throw error('Форма WordPress не использует POST. Сохранение запрещено.');
    if(button.type!=='submit')throw error('Кнопка #publish не является submit. Сохранение запрещено.');
    const action=new URL(form.getAttribute('action')||form.action||'post.php',location.href);
    if(action.origin!==ORIGIN || !/\/wp-admin\/post\.php$/.test(action.pathname))
      throw error('Форма WordPress отправляется на неожиданный адрес: '+action.href+'. Сохранение запрещено.');
    const payload=new FormData(form,button);
    if(button.name && !payload.has(button.name))payload.append(button.name,button.value||'Update');
    if(String(payload.get('post_ID')||'')!==row.id)
      throw error('POST form содержит посторонний post_ID. Сохранение запрещено.');
    if(String(payload.get('action')||'')!=='editpost')
      throw error('POST form содержит неожиданный action (ожидалось editpost). Сохранение запрещено.');
    if(!String(payload.get('_wpnonce')||''))
      throw error('В форме отсутствует nonce WordPress. Сохранение запрещено.');
    if(payload.get('post_title')!==row.name)
      throw error('Перед POST форма изменила название. Сохранение запрещено.');
    const methodName=button.name||'(без имени)';
    log('info','Подготовлен WordPress POST: action=editpost, post_ID='+row.id+', кнопка '+methodName+'. Отправляю один запрос.',row);
    const started=Date.now();
    let response;
    try{
      response=await fetch(action.href,{
        method:'POST',body:payload,credentials:'same-origin',
        redirect:'follow',cache:'no-store',signal:AbortSignal.timeout(60000)
      });
    }catch(ex){
      throw error('Результат WordPress POST неизвестен: '+(ex.message||String(ex))+
        '. Автоматическая повторная отправка запрещена.');
    }
    const finished=new URL(response.url||action.href,ORIGIN);
    const inSameSite=finished.protocol==='https:' && finished.hostname.replace(/^www\./,'')===SITE;
    const elapsed=((Date.now()-started)/1000).toFixed(1);
    log(response.ok?'info':'error',
      'WordPress POST: HTTP '+response.status+', '+elapsed+' сек, redirect='+
      (response.redirected?'да':'нет')+', url='+finished.pathname+finished.search,row);
    if(!inSameSite || !response.ok)
      throw error('WordPress POST вернул неподтверждённый ответ (HTTP '+response.status+'). Повторный Update запрещён.');
    if(/\/wp-login\.php$/.test(finished.pathname))
      throw error('Сессия WordPress истекла при POST. Повторный Update запрещён.');
    // A 200 error page can be returned even after a redirect; inspect the response.
    const html=await response.text();
    const documentResult=new DOMParser().parseFromString(html,'text/html');
    if(documentResult.querySelector('#loginform'))
      throw error('После POST WordPress потребовал повторный вход. Товар направлен на проверку.');
    const wpError=documentResult.querySelector('body.wp-die-message,#wpbody-content > .error,#wpbody-content > .notice-error');
    if(wpError?.textContent?.trim()){
      throw error('WordPress сообщил об ошибке после POST: '+wpError.textContent.trim().slice(0,250));
    }
    // Do not infer success from HTTP 200; only the fresh-page field comparison
    // in run() may mark a course as done.
    return {status:response.status,redirected:response.redirected,url:finished.href,seconds:elapsed};
  }

  async function between(){
    const limit=Date.now()+state.settings.delay*1000;
    while(Date.now()<limit){guard();await sleep(Math.min(250,limit-Date.now()));}
  }
  function stop(pause=false){
    halted=true;state.paused=pause;state.running=false;
    if(['saving','verifying'].includes(state.rows[state.index]?.status))state.stopAfterVerify=true;
    persist();render();
    if(!working){release();return;}
    log('info',`${pause?'Пауза':'Стоп'} запрошен. Новое сохранение не начнётся; отправленный Update будет проверен.`,state.rows[state.index]);
  }
  async function run(){
    if(working)return;
    working=true;leaving=false;
    try{
      while(true){
        guard();
        let i=state.index;
        if(i<0||!state.rows[i]||!ACTIVE.has(state.rows[i].status)){
          i=state.rows.findIndex(x=>['saving','verifying'].includes(x.status)&&x.expected);
          if(i<0)i=state.rows.findIndex(x=>x.status==='pending');
        }
        if(i<0){state.running=false;state.index=-1;state.phase='idle';persist();log('ok','Очередь завершена. Скачайте отчёт.');break;}
        state.index=i;const r=state.rows[i];
        try{
          if(['saving','verifying'].includes(r.status)&&!r.expected){
            r.status='review';
            r.message='Нет данных для проверки прежнего Update. Повторная отправка заблокирована.';
            state.running=false;state.paused=true;state.phase='idle';persist();
            log('error',r.message,r);pauseOtherCourseWorkers();break;
          }
          // Check a previously submitted Update BEFORE deciding to navigate
          // to an editor page. Verification is GET-only and page-independent.
          if(r.expected&&['saving','verifying'].includes(r.status)){
            await verifySavedRow(r);
            const stopNow=state.stopAfterVerify||single||halted||
              (r.status==='review'&&state.settings.stopOnError);
            state.stopAfterVerify=false;state.phase='idle';state.index=-1;persist();
            if(stopNow){state.running=false;persist();break;}
            await between();continue;
          }
          if(productIDFromURL(location.href)!==r.id){
            stage(r,'opening','Открываем товар WooCommerce');
            nav(r.url,'open');return;
          }
          await editorReady(false);
          stage(r,'filling','Считываем товар, заполняем данные');
          r.before=await snapshot();persist();
          await fill('input#title[name="post_title"]',r.name);
          await writeSlug(r.slug);
          await fill('#_regular_price',wooPriceForInput(r.price));
          log('info','Regular price заполнена в формате сайта: '+wooPriceForInput(r.price),r);
          const l=lessonsField();if(l.disabled||l.readOnly)throw error('ACF lessons заблокировано.');
          guard();l.focus();nativeValue(l,r.lessons);await sleep(150);
          if(String(Number(lessonsField().value))!==r.lessons)throw error('ACF lessons не приняло новое значение.');
          await disableDownloadable();
          const selected=await writeImage(r.img,r);
          // Description is written last to avoid WooCommerce/ACF rerenders resetting TinyMCE.
          const normalized=await writeDescription(r.description);
          guard();
          const ed=tiny();if(ed?.initialized&&!ed.isHidden())ed.save();
          r.expected={name:r.name,slug:r.slug,description:normalized,price:r.price,lessons:r.lessons,downloadable:false,img:r.img,attachmentId:selected.attachmentId};
          const beforeSave=await snapshot();
          const mismatch=differences(beforeSave,r.expected);
          if(mismatch.length)throw error('Поля не прошли проверку перед Update: '+mismatch.join(', ')+'.');
          const button=one('#publish');if(button.disabled)throw error('Update недоступна.');
          const form=button.form||document.querySelector('form#post');
          if(!form)throw error('Форма WordPress #post не найдена.');
          if(!form.contains(button))throw error('Кнопка Update не принадлежит форме товара.');
          // Explicitly synchronize the final Visual content with the submitted textarea.
          // TinyMCE can serialize the same markup differently when Text/Visual is toggled.
          // WordPress submits textarea#content, not the visible iframe document.
          const contentField=form.querySelector('textarea#content[name="content"]');
          if(!contentField||contentField.disabled)throw error('POST form не содержит доступное поле Description #content. Update не выполнен.');
          if(!ed?.initialized||ed.isHidden())throw error('Visual недоступен перед Update. Сохранение остановлено.');
          const visualHTML=ed.getContent();
          if(!visualHTML.trim())throw error('Visual не содержит Description перед Update.');
          if(richSignature(visualHTML)!==richSignature(normalized))
            throw error('Visual изменил Description после проверки заполнения. Update не выполнен.');
          ed.save();
          if(contentField.value!==visualHTML){
            // Set only the POST field, without triggering a second TinyMCE re-render.
            const setter=Object.getOwnPropertyDescriptor(unsafeWindow.HTMLTextAreaElement.prototype,'value').set;
            setter.call(contentField,visualHTML);
            log('info','Description синхронизировано: содержимое Visual перенесено в POST textarea.',r);
          }
          await sleep(200);guard();
          const body=new FormData(form);
          if(body.get('post_title')!==r.name)throw error('POST form потеряла название. Update не выполнен.');
          if(body.getAll('content').length!==1)throw error('POST form должна содержать ровно одно поле Description. Update не выполнен.');
          if(String(body.get('content')??'')!==visualHTML)
            throw error('POST form не удерживает Description после синхронизации с Visual. Update не выполнен.');
          const fieldName=lessonsField().name;
          if(!fieldName||String(body.get(fieldName))!==r.lessons)throw error('POST form потеряла lessons. Update не выполнен.');
          if(body.has('_downloadable'))throw error('POST form всё ещё содержит Downloadable. Update не выполнен.');
          if(String(body.get('_regular_price')||'')!==wooPriceForInput(r.price))throw error('POST form содержит неправильную Regular price (ожидалась десятичная запятая). Update не выполнен.');
          // Write a pre-POST checkpoint before any network operation. A crash/timeout
          // leaves the row in saving, where only verification is allowed.
          stage(r,'saving','Сохраняем через контролируемый POST WordPress, затем проверяем на новой странице');
          r.saveAt=Date.now();r.saveDoc=DOC;state.stopAfterVerify=single;persist();
          guard();
          await submitWooUpdate(form,button,r);
          postedFromThisDocument=true;
          // Keeping the manager alive during POST avoids the lost-page event
          // that blocked processing after a native button.click() navigation.
          if(halted)state.stopAfterVerify=true;
          log('info','POST принят. Проверяю товар с сервера через GET без перезагрузки рабочей вкладки.',r);
          await verifySavedRow(r);
          const stopAfter=state.stopAfterVerify||single||halted||
            (r.status==='review'&&state.settings.stopOnError);
          state.stopAfterVerify=false;state.phase='idle';state.index=-1;persist();
          if(stopAfter){state.running=false;persist();break;}
          await between();continue;
        }catch(ex){
          if(ex===STOP)throw ex;
          r.status=r.saveAt&&r.expected?'review':'error';r.message=ex.message||String(ex);
          state.phase='idle';state.index=-1;persist();log('error',r.message,r);
          pauseOtherCourseWorkers();
          // A failed row may leave unsaved edits: stop instead of triggering a risky navigation.
          if(ex.fatal||halted||single||state.settings.stopOnError){state.running=false;persist();break;}
          // For non-stopping mode, navigate to the NEXT pending item from a clean page.
          // The failed row may have uncommitted edits, never re-submit it implicitly.
          const next=state.rows.findIndex(x=>x.status==='pending');
          if(next<0){state.running=false;persist();break;}
          state.index=next;state.phase='opening';persist();
          log('warn','Пропускаем ошибочный товар и открываем следующую запись. На текущем товаре Update не выполнялся.',r);
          nav(state.rows[next].url,'recover');return;
        }
      }
    }catch(ex){if(ex!==STOP)log('error',ex.message||String(ex));state.running=false;persist();}
    finally{if(!leaving){working=false;state.running=false;persist();release();render();}}
  }
  async function start(oneOnly=false,automatic=false){
    if(working){notice('Очередь уже выполняется.');return;}
    if(!IS_COURSE_WORKER){notice('Запускайте очередь из кнопки «Начать» в панели управления.');return;}
    state=GM_getValue(KEY,empty());
    const activePool=currentPool();
    if(!activePool || activePool.run!==state.run || !activePool.workers.includes(WORKER_ID)){
      notice('Рабочая вкладка не принадлежит активному запуску.');return;
    }
    const pendingCommand=GM_getValue(COMMAND,null);
    if(emergencyStopFor(state.run)){
      state.running=false;state.paused=true;persist();
      log('warn','Аварийный стоп всех потоков активен. Новые действия и Update запрещены.');
      return;
    }
    if(pendingCommand?.run===state.run && ['stop','pause'].includes(pendingCommand.kind)){
      state.running=false;state.paused=true;persist();
      log('warn','Рабочий поток остановлен командой из панели до продолжения работы.');
      return;
    }
    if(['running'].includes(GM_getValue('bwo.smm.dates.v2.'+SITE,null)?.status) || GM_getValue('bwo.smm.dates.v2.'+SITE,null)?.phase==='saving'){
      log('error','Редактор дат выполняет задачи на этом сайте. Одновременные изменения заблокированы.');
      return;
    }
    try{
      authCheck();
      if(currentPool()?.run!==state.run)throw error('Рабочий поток устарел и не может изменять товары.',true);
      await acquire();halted=false;
      single=automatic?!!state.single:oneOnly;
      if(!automatic){state.paused=false;state.single=oneOnly;}
      state.owner=tabId;state.running=true;
      // Stalled saving/verifying rows are checked by GET in run().
      // Never navigate or repeat Update just to verify a saved product.
      persist();render();await run();
    }catch(ex){state.running=false;persist();release();log('error',ex.message||String(ex));}
  }
  function editable(){
    if(working || IS_COURSE_WORKER){notice('Остановите обработку и используйте основную панель управления.');return false;}
    if(poolBusy()){notice('Есть активные рабочие вкладки курсов. Сначала дождитесь их завершения.');return false;}
    const legacy=GM_getValue(MASTER_KEY+'.lease',null);
    if(legacy && legacy.expires>Date.now()){notice('Старая очередь ещё выполняется.');return false;}
    state=consolidatePool()||GM_getValue(MASTER_KEY,empty());
    return true;
  }

  async function verifyStalledFromDashboard(){
    if(IS_COURSE_WORKER||working)return;
    const pool=currentPool();
    if(pool?.workers?.some(activeLease)){
      notice('Нельзя проверять зависшие товары, пока рабочие вкладки ещё активны. Остановите их и дождитесь освобождения блокировок.');return;
    }
    if(pool&&Date.now()-pool.createdAt<20000){
      notice('Рабочие потоки только запущены. Повторите проверку после завершения или остановки.');return;
    }
    const candidates=mergedCourseState().rows.filter(r=>r.expected&&r.saveAt&&
      ['saving','verifying','review'].includes(r.status));
    if(!candidates.length){notice('Нет товаров с отправленным Update для повторной проверки.');return;}
    if(!confirm('Проверить сохранение товара ID '+candidates[0].id+
      ' без повторного Update? Старые рабочие вкладки должны быть закрыты.'))return;
    const base=consolidatePool(true);
    if(!base){notice('Очередь ещё используется рабочим потоком.');return;}
    state=base;
    const row=state.rows.find(r=>r.expected&&r.saveAt&&
      ['saving','verifying','review'].includes(r.status));
    if(!row){notice('Не найден товар для проверки.');return;}
    working=true;state.index=state.rows.indexOf(row);state.phase='verifying';
    row.status='verifying';row.message='Ручная проверка после зависшего Update; повторный POST не отправляется.';
    persist();render();
    log('info','Начинаю восстановление зависшего товара '+row.id+' через GET WordPress без Update.',row);
    let problem='';
    try{
      for(let attempt=1;attempt<=2;attempt++){
        try{
          const actual=await fetchSavedSnapshot(row);
          row.actual=actual;row.checked=new Date().toISOString();persist();
          const differencesList=differences(actual,row.expected);
          if(!differencesList.length){
            row.status='done';row.message='Сохранение подтверждено после восстановления';
            log('ok',row.message,row);problem='';break;
          }
          problem='Не совпали поля: '+differencesList.join(', ');
        }catch(ex){problem=ex.message||String(ex);}
        log('warn','Восстановление, проверка '+attempt+'/2: '+problem,row);
        if(attempt<2)await sleep(1200);
      }
      if(problem){
        row.status='review';row.message=problem+'; Update не отправлялся.';
        log('error',row.message,row);
      }
    }finally{
      state.index=-1;state.phase='idle';state.paused=true;state.running=false;
      persist();working=false;render();
    }
  }
  function shuffle(rows){for(let i=rows.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[rows[i],rows[j]]=[rows[j],rows[i]];}return rows;}
  function sheetRows(){return getXLSX().utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[Number($('sheet').value)]],{header:1,defval:'',raw:false,blankrows:true});}
  async function readFile(file){
    if(!file||!editable())return;
    try{
      const xlsx=await ensureXLSX();
      if(file.size>25*1024*1024)throw error('Файл больше 25 МБ.');
      workbook=xlsx.read(await file.arrayBuffer(),{type:'array',cellDates:false});importFile=file.name;
      $('filename').textContent=file.name;
      $('sheet').innerHTML=workbook.SheetNames.map((name,i)=>`<option value="${i}">${esc(name)}</option>`).join('');
      $('mapping').classList.remove('hidden');setMapping();switchTab('import');
    }catch(ex){notice(ex.message||String(ex));}
  }
  function setMapping(){
    if(!workbook)return;
    const header=sheetRows()[0]||[];
    for(const [key,aliases] of Object.entries(ALIAS)){
      const el=$(`map-${key}`);
      el.innerHTML='<option value="">Выберите колонку…</option>'+header.map((h,i)=>`<option value="${i}">${esc(h||`Колонка ${i+1}`)}</option>`).join('');
      const found=header.findIndex(h=>aliases.includes(cleanHeader(h)));el.value=found<0?'':String(found);
    }
  }
  function loadQueue(){
    if(!workbook||!editable())return;
    try{
      const names=Object.keys(ALIAS),maps=names.map(name=>$(`map-${name}`).value);
      if(maps.some(v=>v==='')||new Set(maps).size!==names.length)throw error('Сопоставьте шесть разных колонок: ID, Name, Description, Regular price, lessons, img.');
      const [idIndex,nameIndex,descriptionIndex,priceIndex,lessonsIndex,imgIndex]=maps.map(Number);
      const rows=[],seen=new Set(),data=sheetRows();
      for(let i=1;i<data.length;i++){
        const line=data[i];if(line.every(cell=>!String(cell).trim()))continue;
        const r={key:uid(),excelRow:i+1,id:String(line[idIndex]??'').trim(),name:String(line[nameIndex]??'').trim(),description:String(line[descriptionIndex]??''),rawPrice:String(line[priceIndex]??''),rawLessons:String(line[lessonsIndex]??''),rawImg:String(line[imgIndex]??'').trim(),status:'pending',message:''};
        try{
          if(!/^[1-9]\d*$/.test(r.id))throw error('ID должен быть числовым WordPress post ID.');
          if(seen.has(r.id))throw error('Повторный ID товара в Excel.');
          seen.add(r.id);
          if(!r.name)throw error('Name не заполнено.');
          r.slug=slugFromTitle(r.name);if(!r.slug||r.slug.length>180)throw error('Не удалось получить корректный slug (до 180 символов) из Name.');
          r.description=validateHTML(r.description);
          r.price=normalizePrice(r.rawPrice);
          r.lessons=normalizeLessons(r.rawLessons);
          r.img=normalizeImg(r.rawImg);
          r.url=productURL(r.id);
        }catch(ex){r.status='error';r.invalid=true;r.message=ex.message||String(ex);}
        rows.push(r);
      }
      if(!rows.length)throw error('В Excel нет данных.');
      if(state.rows.length&&!confirm('Заменить текущую очередь? Для сохранения истории сначала скачайте отчёт.'))return;
      const previous=state.settings;state=empty();state.settings=previous;state.filename=importFile;state.rows=shuffle(rows);
      page=0;persist();switchTab('queue');log('info',`Импортировано ${rows.length} товаров. Очередь перемешана. Ошибочных строк: ${rows.filter(r=>r.invalid).length}. Номера строк Excel сохранены.`);
    }catch(ex){notice(ex.message||String(ex));}
  }
  function download(payload,name,mime){
    const url=URL.createObjectURL(new Blob([payload],{type:mime}));const a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),20000);
  }
  function csv(rows){return '\uFEFF'+rows.map(row=>row.map(cell=>{let s=String(cell??'');if(/^\s*[=+@-]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';}).join(';')).join('\r\n');}
  function exportReport(){
    if(!IS_COURSE_WORKER)state=mergedCourseState();
    const header=['Excel Row','Product ID','Name','Edit URL','Description','Regular price','lessons','img','Status','Message','Checked At','Original name','Original slug','Original price','Original lessons','Original image','Saved image','Saved WordPress attachment ID','Excel price (source)'];
    const rows=state.rows.map(r=>[r.excelRow,r.id,r.name,r.url,r.description,r.price,r.lessons,r.img,STATUSES[r.status],r.message,r.checked,r.before?.name,r.before?.slug,r.before?.price,r.before?.lessons,r.before?.image,r.actual?.image,r.actual?.attachmentId,r.rawPrice]);
    download(csv([header,...rows]),`${SITE}-woo-products-report.csv`,'text/csv;charset=utf-8');
  }
  function exportLog(){if(!IS_COURSE_WORKER)state=mergedCourseState();download(csv([['Time','Level','Excel Row','Product ID','Stage','Message'],...state.logs.map(x=>[x.time,x.level,x.row,x.id,x.stage,x.message])]),`${SITE}-woo-products-log.csv`,'text/csv;charset=utf-8');}
  function notice(message){
    if(!panel)return;panel.querySelector('.wc-error')?.remove();const el=document.createElement('div');el.className='wc-error';const close=document.createElement('button');close.textContent='×';close.onclick=()=>el.remove();el.append(close,document.createTextNode(String(message)));panel.append(el);
  }
  function resetAll(){
    if(!editable())return;
    if(!confirm(`Очистить очередь, журнал и настройки для ${SITE}? Изменения в WooCommerce сохранятся.`))return;
    release();state=empty();persist();workbook=null;importFile='';page=0;filter='all';search='';
    $('filename').textContent='Данные Excel обрабатываются локально';$('sheet').innerHTML='';$('file').value='';$('mapping').classList.add('hidden');
    for(const k of Object.keys(ALIAS))$(`map-${k}`).innerHTML='';
    $('search').value='';$('filter').value='all';panel.querySelector('.wc-detail')?.remove();panel.querySelector('.wc-error')?.remove();switchTab('import');
  }
  function retryErrors(){
    if(!editable())return;let count=0;
    for(const r of state.rows){if(r.status==='error'&&!r.invalid){r.status='pending';r.message='Повторная попытка';delete r.expected;delete r.saveAt;delete r.saveDoc;count++;}}
    persist();log('info',`Возвращено ошибок в очередь: ${count}`);
  }
  function detail(r){
    if(!r)return;panel.querySelector('.wc-detail')?.remove();
    const d=document.createElement('div');d.className='wc-detail';
    d.innerHTML=`<header><b>Строка ${r.excelRow} · товар ID ${esc(r.id)}</b><button data-close>Закрыть</button></header>
      <p><span class="tag ${esc(r.status)}">${STATUSES[r.status]}</span> ${esc(r.message)}</p><p><strong>${esc(r.name)}</strong></p>
      <p>Slug: <code>${esc(r.slug||'—')}</code></p><p>Цена: ${esc(r.price)} · Уроков: ${esc(r.lessons)} · img: ${esc(r.img)}</p>
      <p><a href="${esc(r.url||'#')}" target="_blank" rel="noopener noreferrer">Открыть товар</a></p><h4>HTML Description</h4><pre>${esc(r.description)}</pre>
      <div class="actions"><button data-recheck>Проверить без Update</button><button data-reset>Вернуть в очередь</button><button data-skip>Пропустить</button></div>`;
    d.querySelector('[data-close]').onclick=()=>d.remove();
    d.querySelector('[data-recheck]').disabled=!r.expected||!r.saveAt;
    d.querySelector('[data-recheck]').onclick=()=>{
      if(!editable())return;const row=state.rows.find(x=>x.key===r.key);
      row.status='verifying';row.message='Повторная проверка без Update';row.saveDoc=DOC;state.index=-1;state.phase='idle';persist();d.remove();launchCourse(true);
    };
    d.querySelector('[data-reset]').onclick=()=>{
      if(!editable())return;const row=state.rows.find(x=>x.key===r.key);
      if(row.invalid){notice('Исправьте Excel и импортируйте его заново.');return;}
      if(['done','review','saving','verifying'].includes(row.status)&&!confirm('Повторно редактировать и сохранять товар? Если Update уже был отправлен, сначала используйте «Проверить без Update».'))return;
      row.status='pending';row.message='Возвращено в очередь вручную';delete row.expected;delete row.saveAt;delete row.saveDoc;
      if(state.index===state.rows.indexOf(row)){state.index=-1;state.phase='idle';}
      persist();render();d.remove();
    };
    d.querySelector('[data-skip]').onclick=()=>{
      if(!editable())return;const row=state.rows.find(x=>x.key===r.key);row.status='skipped';row.message='Пропущено вручную';if(state.index===state.rows.indexOf(row)){state.index=-1;state.phase='idle';}persist();render();d.remove();
    };
    panel.append(d);
  }
  function switchTab(value){
    tab=value;page=0;
    for(const name of ['queue','import','log','settings','help'])
      $(name).classList.toggle('hidden',name==='queue'?!['queue','processed'].includes(value):name!==value);
    root.querySelectorAll('[data-tab]').forEach(b=>b.classList.toggle('active',b.dataset.tab===value));render();
  }
  function renderRows(){
    if(!root)return;
    const doneTab=tab==='processed';
    const list=state.rows.filter(r=>(doneTab?['done','review'].includes(r.status):!['done','review'].includes(r.status))
      &&(filter==='all'||(filter==='errors'?['error','review'].includes(r.status):r.status===filter))
      &&`${r.name} ${r.id} ${r.img}`.toLowerCase().includes(search.toLowerCase()));
    const pages=Math.max(1,Math.ceil(list.length/25));page=Math.min(page,pages-1);
    $('pages').textContent=`${page+1} / ${pages}`;$('prev').disabled=page===0;$('next').disabled=page>=pages-1;
    if(!list.length){$('rows').innerHTML='<div class="empty"><b>Нет товаров</b>Импортируйте Excel или измените фильтр.</div>';return;}
    $('rows').innerHTML=`<table><thead><tr><th style="width:48px">Excel</th><th>Товар</th><th style="width:68px">Цена</th><th style="width:96px">Статус</th><th style="width:42px"></th></tr></thead><tbody>`+
      list.slice(page*25,page*25+25).map(r=>`<tr class="${state.rows[state.index]===r?'current-row':''}"><td>${r.excelRow}</td><td><strong title="${esc(r.name)}">${esc(r.name)}</strong><small>ID ${esc(r.id)} · img ${esc(r.img)} · ${esc(r.message||'')}</small></td><td>${esc(r.price??'—')}</td><td><span class="tag ${esc(r.status)}">${STATUSES[r.status]}</span></td><td><button data-detail="${esc(r.key)}">⋯</button></td></tr>`).join('')+'</tbody></table>';
    $('rows').querySelectorAll('[data-detail]').forEach(el=>el.onclick=()=>detail(state.rows.find(r=>r.key===el.dataset.detail)));
  }
  function renderLog(){if(root)$('logs').innerHTML=state.logs.slice().reverse().map(x=>`<div class="log ${esc(x.level)}"><time>${new Date(x.time).toLocaleTimeString()}</time><span>${x.row?'№ '+x.row:'Система'}</span><span>${esc(x.message)}</span></div>`).join('')||'<div class="empty">Журнал пуст</div>';}
  function render(){
    if(!root)return;
    if(!IS_COURSE_WORKER && !working)state=mergedCourseState();
    const count=statuses=>state.rows.filter(r=>statuses.includes(r.status)).length;
    $('total').textContent=state.rows.length;$('done').textContent=count(['done']);$('pending').textContent=count(['pending',...ACTIVE]);$('failed').textContent=count(['error','review']);
    $('bar').style.width=(state.rows.length?100*count(['done','skipped'])/state.rows.length:0)+'%';
    const current=state.rows[state.index];
    const ws=state.poolWorkers||[];
    $('current').textContent=ws.length?('Рабочие вкладки: '+ws.filter(x=>x.live).length+'/'+ws.length+' · '+ws.map(x=>'№'+x.id+': '+x.done+'/'+x.total+(x.live?' работает':'')).join(' · ')):
      (working||state.running)?(halted?'Остановка после текущей операции':`${STATUSES[current?.status]||'Работа'} · ${current?.name||current?.id||''}`):
      state.paused?'На паузе — нажмите Продолжить':state.rows.length?`${state.filename} · ${SITE}`:'Загрузите Excel с товарами';
    const hasWork=state.rows.some(x=>x.status==='pending'||ACTIVE.has(x.status));
    const isActive=working||!!state.running||poolBusy();
    $('start').disabled=isActive||!hasWork;$('one').disabled=isActive||!hasWork;$('stop').disabled=!isActive;$('pause').disabled=!isActive||halted;
    const pool=currentPool();
    const hasUnverified=state.rows.some(r=>r.expected&&r.saveAt&&['saving','verifying','review'].includes(r.status));
    $('verifyStalled').disabled=working||IS_COURSE_WORKER||!hasUnverified||
      !!pool?.workers?.some(activeLease)||!!(pool&&Date.now()-pool.createdAt<20000);
    $('start').textContent=state.paused||state.rows.some(x=>ACTIVE.has(x.status))?'▶ Продолжить':'▶ Начать';
    $('retry').disabled=isActive||!state.rows.some(x=>x.status==='error'&&!x.invalid);
    $('load').disabled=isActive;$('reset').disabled=isActive;
    $('delay').disabled=isActive;$('stopOnError').disabled=isActive;$('workers').disabled=isActive;
    $('delay').value=state.settings.delay;$('stopOnError').checked=!!state.settings.stopOnError;
    $('workers').value=Math.max(1,Math.min(5,Number(state.settings.workers)||1));
    $('updated').textContent=state.updated?`Сохранено ${new Date(state.updated).toLocaleTimeString()} · ${SITE}`:`Очередь хранится отдельно для ${SITE}`;
    root.querySelector('[data-tab="processed"]').textContent=`Обработанные (${count(['done','review'])})`;
    if(['queue','processed'].includes(tab))renderRows();if(tab==='log')renderLog();
  }

  // Register additional functions here without changing the shared panel or queue engines.
  const MODULES = [];
  function registerModule(module){
    if(!module?.id||!module?.title||!module?.container||MODULES.some(m=>m.id===module.id))
      throw new Error('Некорректный или повторяющийся модуль SMM');
    MODULES.push(module);
  }
  registerModule({id:'courses',title:'Редактор курсов',container:'course-content'});
  registerModule({id:'dates',title:'Редактор дат',container:'date-content',
    mount:()=>initializeDateManager(root,{SITE,ORIGIN})});
  function selectModule(name){
    if(!root)return;
    const chosen=MODULES.some(m=>m.id===name)?name:MODULES[0].id;
    MODULES.forEach(m=>$(m.container).classList.toggle('hidden',m.id!==chosen));
    root.querySelectorAll('[data-module]').forEach(b=>b.classList.toggle('active',b.dataset.module===chosen));
    GM_setValue(MODE_KEY,chosen);
  }
  function openDashboard(){
    const url=ORIGIN+'/wp-admin/index.php?smm_manager=dashboard';
    const w=window.open(url,'smm-manager-dashboard-'+SITE);
    if(!w)notice('Браузер заблокировал новую вкладку. Разрешите всплывающие окна для сайта.');
  }
  function launchCourse(oneOnly){
    if(working || IS_COURSE_WORKER){notice('Запустите очередь из основной панели.');return;}
    const dates=GM_getValue('bwo.smm.dates.v2.'+SITE,null);
    if(dates?.status==='running'||dates?.phase==='saving'){notice('Редактор дат работает на этом сайте. Сначала остановите его.');return;}
    const oldPool=currentPool();
    const stoppedPool=!!oldPool && emergencyStopFor(oldPool.run);
    const canRecoverStopped=stoppedPool && !oldPool.workers.some(activeLease) &&
      Date.now()-oldPool.createdAt>=20000;
    if(poolBusy() && !canRecoverStopped){
      notice(stoppedPool?
        'Аварийный стоп отправлен, но ещё есть активные вкладки или блокировки. Дождитесь их освобождения.' :
        'Рабочие вкладки курсов уже активны.');
      return;
    }
    const legacy=GM_getValue(MASTER_KEY+'.lease',null);
    if(legacy?.expires>Date.now()){notice('Предыдущий обработчик курсов ещё активен.');return;}
    const base=consolidatePool(canRecoverStopped);
    if(!base){notice('Рабочие потоки ещё завершают задачи.');return;}
    const uncertain=base.rows.filter(r=>['saving','verifying'].includes(r.status)&&!r.expected);
    if(uncertain.length){
      for(const row of uncertain){
        row.status='review';
        row.message='Предыдущая операция сохранения не имеет данных для проверки. Проверьте товар вручную.';
      }
      GM_setValue(MASTER_KEY,base);
    }
    // A previous Update with expected values is verified only; it is never submitted again.
    const jobs=[...base.rows.filter(r=>['saving','verifying'].includes(r.status)&&r.expected),
      ...base.rows.filter(r=>r.status==='pending'||['opening','filling'].includes(r.status))];
    if(!jobs.length){notice('Нет курсов в очереди. Для товара со статусом «Проверить» используйте проверку без Update.');return;}
    const picked=oneOnly?jobs.slice(0,1):jobs;
    const amount=oneOnly?1:Math.min(5,picked.length,Math.max(1,Math.round(Number(base.settings?.workers)||1)));
    const pool={run:uid(),workers:Array.from({length:amount},(_,i)=>i+1),createdAt:Date.now(),oneOnly};
    const buckets=pool.workers.map(()=>[]);
    picked.forEach((row,i)=>buckets[i%amount].push({...row,status:['opening','filling'].includes(row.status)?'pending':row.status}));
    for(let i=0;i<amount;i++){
      const shard=empty();
      shard.run=pool.run;shard.filename=base.filename;shard.single=oneOnly;
      shard.settings={...empty().settings,...base.settings};shard.rows=buckets[i];
      GM_setValue(workerKey(i+1),shard);
      GM_setValue(workerKey(i+1)+'.lease',null);
      GM_setValue(workerKey(i+1)+'.command',null);
    }
    // A new run is only possible through this explicit Start action after
    // all old leases are gone and their queue has been consolidated.
    GM_setValue(STOP_ALL_KEY,null);
    GM_setValue(POOL_KEY,pool);
    let opened=0;
    // Only N named tabs are opened per launch. All subsequent courses use location.assign in those tabs.
    for(let i=0;i<amount;i++){
      const url=new URL(productURL(buckets[i][0].id));
      url.searchParams.set('smm_manager','course-worker');
      url.searchParams.set('smm_worker',String(i+1));
      url.searchParams.set('smm_run',pool.run);
      url.searchParams.set('smm_one',oneOnly?'1':'0');
      if(window.open(url.href,WORKER_PREFIX+(i+1)))opened++;
    }
    state=mergedCourseState();render();
    if(opened<amount)notice('Открыто '+opened+' из '+amount+' вкладок. Разрешите всплывающие окна для сайта.');
  }
  function pauseOtherCourseWorkers(){
    if(!IS_COURSE_WORKER||!state.settings.stopOnError)return;
    const pool=currentPool();
    if(!pool||pool.run!==state.run)return;
    for(const id of pool.workers){
      if(id===WORKER_ID)continue;
      GM_setValue(workerKey(id)+'.command',{kind:'pause',run:pool.run,at:Date.now(),id:uid()});
    }
  }
  function emergencyStopAllCourseWorkers(){
    const pool=currentPool();
    if(!pool?.workers?.length){
      notice('Нет активного запуска курсов. Останавливать нечего.');return;
    }
    const now=Date.now();
    const existing=GM_getValue(STOP_ALL_KEY,null);
    if(existing?.run===pool.run){
      notice('Аварийная остановка уже запрошена. Новые действия заблокированы.');return;
    }
    // Store the durable barrier FIRST. A worker which missed a change event
    // still sees it in guard() and at every new page startup.
    GM_setValue(STOP_ALL_KEY,{run:pool.run,at:now,reason:'emergency',site:SITE});
    GM_setValue(POOL_KEY,{...pool,emergencyStoppedAt:now});
    for(const id of pool.workers){
      // Backward compatible "stop" command for workers already running old code.
      GM_setValue(workerKey(id)+'.command',{
        kind:'stop',run:pool.run,id:uid(),at:now,emergency:true
      });
    }
    if(IS_COURSE_WORKER && state.run===pool.run)stop(false);
    if(root)render();
    notice('АВАРИЙНЫЙ СТОП: запрет новых действий отправлен всем '+pool.workers.length+
      ' потокам. Уже отправленный Update не отменяется. Результаты очереди сохранены.');
  }
  // Always accessible from Tampermonkey, including on a worker page when the panel
  // is hidden or the page navigation did not finish.
  GM_registerMenuCommand('АВАРИЙНЫЙ СТОП ВСЕХ ПОТОКОВ · Курсы',emergencyStopAllCourseWorkers);
  function stopCourseFromPanel(pause){
    if(IS_COURSE_WORKER){stop(pause);return;}
    const pool=currentPool();
    if(!pool){notice('Очередь курсов сейчас не выполняется.');return;}
    let active=0;
    for(const id of pool.workers){
      if(activeLease(id))active++;
      // Persist the command even for a tab that is still loading or navigating.
      GM_setValue(workerKey(id)+'.command',{kind:pause?'pause':'stop',run:pool.run,id:uid(),at:Date.now()});
    }
    notice((pause?'Пауза':'Стоп')+' отправлена для '+pool.workers.length+' потоков; активных: '+active+'.');
  }

  const css=`
    .module-nav{display:flex;gap:8px;padding:9px 14px;background:#eef3fb;border-bottom:1px solid #d6e1ef;align-items:center}
    .module-actions{display:flex;gap:8px;flex-wrap:wrap}
    .module-nav button{font-size:12px;padding:8px 12px;background:#fff}
    .module-nav button.active{background:#2e64ce;color:#fff;border-color:#2e64ce}
    .module-nav .spacer{flex:1}
    .panel.dashboard{inset:0!important;top:0!important;right:0!important;bottom:0!important;left:0!important;width:100vw!important;height:100vh!important;max-height:none!important;max-width:none!important;border-radius:0!important;resize:none!important}

    :host{all:initial;font:13px/1.5 -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#182438;color-scheme:light}
    *{box-sizing:border-box}button,input,select{font:inherit}button{background:#fff;color:#26354c;border:1px solid #d9e1eb;border-radius:9px;padding:8px 12px;cursor:pointer;font-weight:600}button:hover{background:#edf3fa;color:#182438}button:disabled{opacity:.44;cursor:default}button:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid #4374e9;outline-offset:2px}
    .panel{position:fixed;right:18px;bottom:18px;width:760px;height:680px;max-width:calc(100vw - 24px);max-height:calc(100vh - 36px);display:flex;flex-direction:column;border-radius:16px;box-shadow:0 18px 75px #12243a55;background:white;border:1px solid #d6dfed;overflow:hidden;resize:both}
    .panel.full{inset:20px;width:auto;height:auto;max-height:none;max-width:none;resize:none}
    .head{background:#14263e;color:#fff;padding:14px 16px;display:flex;align-items:center;gap:10px;cursor:move;touch-action:none}.logo{width:36px;height:36px;border-radius:10px;background:#3d74e5;display:grid;place-items:center;font-weight:800;font-size:18px}.head b{font-size:16px}.head small{display:block;color:#aec0d8;font-size:11px}.head .space{flex:1}.head button{color:#fff;background:#ffffff18;border-color:#ffffff20;padding:5px 10px}.head button:hover{color:#fff;background:#ffffff35}
    .body{display:flex;flex-direction:column;min-height:0;flex:1}.stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;padding:13px 15px 9px}.stat{border:1px solid #e4e9f2;background:#fafcfe;border-radius:10px;padding:8px 10px}.stat b{display:block;font-size:21px;line-height:1.25}.stat small{font-size:11px;color:#718198}.stat.success b{color:#17855a}.stat.bad b{color:#b64e3d}.progress{height:5px;background:#e9f0fa;margin:0 15px;border-radius:6px;overflow:hidden}.progress i{display:block;background:#3c74e9;height:100%;transition:width .3s}
    .current{padding:10px 15px;color:#5c718a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.actions{display:flex;gap:7px;flex-wrap:wrap;padding:0 15px 10px}.primary{background:#346cdf!important;border-color:#346cdf!important;color:white!important}.primary:hover{background:#2459c8!important}.danger{background:#fff7f5;color:#a3433c;border-color:#edc2bc}
    .tabs{display:flex;overflow-x:auto;flex-shrink:0;border-bottom:1px solid #e6ecf3;padding:0 15px;gap:16px}.tabs button{border:0;border-bottom:2px solid transparent;border-radius:0;background:transparent;color:#718198;padding:10px 1px;font-size:12px;white-space:nowrap}.tabs button.active{color:#2661cd;border-bottom-color:#346cdf}
    .view{padding:12px 15px;min-height:0;flex:1;overflow:auto}.hidden{display:none!important}.tools{display:flex;gap:8px;margin-bottom:10px}input,select{min-width:0;border:1px solid #dce3ed;border-radius:8px;background:white;color:#26364c;padding:8px}input[type=search]{flex:1}table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:12px}th{position:sticky;top:0;background:#f6f8fc;text-align:left;color:#8390a2;font-size:10px;text-transform:uppercase;letter-spacing:.4px;padding:8px 6px}td{padding:9px 6px;border-bottom:1px solid #edf1f6;vertical-align:top}td strong{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}td small{display:block;font-size:10px;color:#8694a7;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}td button{padding:2px 7px}.current-row{background:#edf4ff}
    .tag{border-radius:6px;background:#edf1f7;color:#5f718a;padding:3px 7px;font-size:10px;display:inline-block;white-space:nowrap}.tag.done{color:#147d50;background:#e3f6ec}.tag.error,.tag.review{color:#ad4b27;background:#fff0e7}.tag.saving,.tag.opening,.tag.filling,.tag.verifying{color:#2861c8;background:#e7efff}
    .drop{border:2px dashed #c3d2e9;background:#f6faff;border-radius:12px;padding:18px;text-align:center;margin-bottom:12px}.drop label{font-weight:700;color:#2e62d7;cursor:pointer}.drop small{display:block;color:#788aa2;font-size:11px;margin-top:6px}.drop.over{background:#e8f1ff;border-color:#346cdf}.mapping{display:grid;grid-template-columns:150px minmax(0,1fr);gap:8px;align-items:center}.mapping label{color:#63758c}.hint{font-size:12px;color:#6b7b92;line-height:1.5}.notice{padding:10px 12px;margin:10px 0;background:#fff7e8;border:1px solid #e9d4a8;border-radius:9px;color:#80622f;font-size:12px}
    .setting{display:flex;gap:12px;align-items:center;padding:12px 0;border-bottom:1px solid #edf1f6}.setting div{flex:1}.setting small{display:block;color:#78879c}.setting input[type=number]{width:80px}.foot{border-top:1px solid #e6edf5;background:#fafbfd;padding:8px 15px;display:flex;gap:10px;align-items:center;color:#7d8b9d;font-size:11px}.foot span:first-child{flex:1}.foot button{padding:4px 8px}
    .log{display:grid;grid-template-columns:66px 48px 1fr;gap:8px;font-size:11px;padding:8px 0;border-bottom:1px solid #edf1f6;overflow-wrap:anywhere}.log time{color:#8c98a8}.log.error{color:#a63e30}.log.ok{color:#167b53}.empty{text-align:center;color:#8390a0;padding:32px 16px}.empty b{display:block;color:#263a55;font-size:15px}.help h3{font-size:13px;color:#2c405c}.help p,.help li{font-size:12px;color:#62748c}.help ol{padding-left:22px}
    button.pill{position:fixed;right:18px;bottom:18px;color:#fff!important;background:#14263e!important;border:1px solid #365875!important;border-radius:13px;padding:12px 18px;box-shadow:0 8px 30px #0b234044;font-weight:700}button.pill:hover,button.pill:focus-visible{color:#fff!important;background:#294e79!important;border-color:#5e88b4!important}
    .wc-error{position:absolute;bottom:52px;left:15px;right:15px;z-index:9;background:#fff1ef;border:1px solid #f0c1b8;color:#9f3e34;padding:12px;border-radius:10px;box-shadow:0 8px 25px #0002}.wc-error button{float:right;padding:1px 6px}.wc-detail{position:absolute;inset:60px 12px 12px;z-index:6;background:#fff;border:1px solid #d4e0ef;box-shadow:0 10px 35px #0003;border-radius:12px;padding:15px;overflow:auto}.wc-detail header{display:flex;justify-content:space-between}.wc-detail pre{white-space:pre-wrap;overflow-wrap:anywhere;font:11px/1.5 monospace;background:#f6f8fc;padding:10px;border-radius:8px}.wc-detail a{color:#2863da}.wc-detail .actions{padding:0}
    @media(max-width:540px){.panel{right:7px;bottom:7px;width:calc(100vw - 14px);height:min(650px,calc(100vh - 14px))}.stats{gap:4px}.stat{padding:6px}.actions button{font-size:11px;padding:6px}.tabs{gap:11px}}
  `;
function initializeDateManager(root, context) {
  'use strict';
  const { SITE, ORIGIN } = context;
  if (window.__SMM_MANAGER_DATES__) return;
  window.__SMM_MANAGER_DATES__ = true;
  const params = new URLSearchParams(location.search);
  const DATE_WORKER_SESSION = 'smm.manager.date.worker.' + SITE;
  const workerURL = ORIGIN + '/wp-admin/edit.php?post_type=product&smm_manager=date-worker';
  const workerName = 'smm-manager-dates-' + SITE;
  if (params.get('smm_manager') === 'date-worker') {
    window.name = workerName;
    sessionStorage.setItem(DATE_WORKER_SESSION, '1');
  }
  const IS_DATE_WORKER = params.get('smm_manager') !== 'dashboard' &&
    window.name === workerName && sessionStorage.getItem(DATE_WORKER_SESSION) === '1';

  const STORAGE = 'bwo.smm.dates.v2.' + SITE;
  const LEGACY_STORAGE = 'smmmaster_dates_v1_state';
  const DEFAULT_DATE_TO = new Date().toISOString().slice(0,10);
  const DEFAULT_DATE_FROM = DEFAULT_DATE_TO.slice(0,4) + '-01-01';
  const MAX_STORED_LOG = 60;
  const $ = (sel, scope = root) => scope.querySelector(sel);
  const $$ = (sel, scope = document) => Array.from(scope.querySelectorAll(sel));
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  let busy = false;
  let rowsFromFile = [];
  let activeFileName = '';
  let activeSheet = '';

  const emptyState = () => ({
    ids: [], index: 0, status: 'idle', phase: 'search',
    records: [], notes: [], lastMessage: 'Загрузите файл с ID',
    batchRemaining: null, dateCandidate: null, source: '',
    missStreak: 0, delayMs: 350, fastMode: false, prefetch: 2, dryRun: false,
    dateFrom: DEFAULT_DATE_FROM, dateTo: DEFAULT_DATE_TO,
    lastChanged: Date.now()
  });
  let state;
  let initial = GM_getValue(STORAGE, null);
  if (!initial && SITE === 'smmmaster.net') {
    try {
      const legacy = JSON.parse(localStorage.getItem(LEGACY_STORAGE) || 'null');
      if (legacy && Array.isArray(legacy.ids)) {
        initial = { ...legacy, status:'paused', phase: legacy.phase === 'saving' ? 'saving' : 'search' };
        GM_setValue(STORAGE, initial);
      }
    } catch (e) { console.warn('[SMM dates] Migration skipped:',e); }
  }
  state = { ...emptyState(), ...(initial && typeof initial === 'object' ? initial : {}) };
  if (!initial) GM_setValue(STORAGE, state);
  if (!Array.isArray(state.ids)) state.ids = [];
  if (!Array.isArray(state.records)) state.records = [];
  if (!Array.isArray(state.notes)) state.notes = [];

  const save = () => {
    state.lastChanged = Date.now();
    try { GM_setValue(STORAGE, state); }
    catch (e) { console.error('[SMM] Не удалось сохранить очередь:', e); }
  };
  const count = (status) => state.records.filter(r => r.status === status).length;
  const pct = () => state.ids.length ? Math.round(100 * state.index / state.ids.length) : 0;
  const setMessage = (msg) => { state.lastMessage = msg; save(); draw(); };
  const log = (msg, severity = 'info') => {
    const hhmmss = new Date().toLocaleTimeString('ru-RU');
    state.notes.unshift({ time: hhmmss, msg, severity });
    state.notes.length = Math.min(state.notes.length, MAX_STORED_LOG);
    state.lastMessage = msg;
    save(); draw();
  };
  function statusLine(msg, severity = 'info') {
    console[severity === 'error' ? 'error' : 'log']('[SMM]', msg);
    if ($('#smm-date-tool')) log(msg, severity);
  }
  const pad = v => String(v).padStart(2, '0');
  const isMidnight = date => date?.hh === '00' && date?.mn === '00';
  function parseDateBoundary(value) {
    if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return null;
    const [year, month, day] = value.split('-').map(Number);
    const ms = Date.UTC(year, month - 1, day);
    const d = new Date(ms);
    if (d.getUTCFullYear() !== year || d.getUTCMonth() !== month - 1 || d.getUTCDate() !== day) return null;
    return ms;
  }
  function dateRange(from = state.dateFrom, to = state.dateTo) {
    const start = parseDateBoundary(from), end = parseDateBoundary(to);
    if (start === null || end === null || start > end) return null;
    return { start, end: end + 86400000 - 1000 };
  }
  const dateLabel = value => value.split('-').reverse().join('.');
  const randomDate = () => {
    const range = dateRange();
    if (!range) throw new Error('Некорректный диапазон дат. Укажите начало и конец периода.');
    let vals;
    do {
      const n = Math.floor(range.start / 1000) + Math.floor(Math.random() * (Math.floor(range.end / 1000) - Math.floor(range.start / 1000) + 1));
      const d = new Date(n * 1000);
      vals = { aa: String(d.getUTCFullYear()), mm: pad(d.getUTCMonth()+1), jj: pad(d.getUTCDate()),
        hh: pad(d.getUTCHours()), mn: pad(d.getUTCMinutes()), ss: pad(d.getUTCSeconds()) };
    } while (isMidnight(vals)); // Никогда не выставлять 00:00 (даже 00:00:59)
    vals.text = `${vals.aa}-${vals.mm}-${vals.jj} ${vals.hh}:${vals.mn}:${vals.ss}`;
    return vals;
  };
  const unique = ids => [...new Set(ids.map(x => String(x).trim()).filter(x => /^\d{1,12}$/.test(x)))];
  function toId(raw) {
    const s = String(raw ?? '').trim();
    if (/^\d{1,12}$/.test(s)) return s;
    let m = s.match(/^(\d{4,12})\s+\S/); // 204665 Crab God
    if (m) return m[1];
    m = s.match(/(?:^|\s)(\d{4,12})\s*$/); // Crab God 204665
    return m ? m[1] : '';
  }
  function chooseBestColumn(rows) {
    const head = (rows[0] || []).map(x => String(x ?? '').trim().toLowerCase());
    const exact = head.findIndex(x => /^(id|id товара|product\s*id|ид|айди|идентификатор)$/.test(x));
    if (exact >= 0) return exact;
    let best = 0, top = -1;
    for (let c = 0; c < Math.min(30, Math.max(...rows.slice(0, 10).map(r => r.length), 1)); c++) {
      const vals = rows.slice(0, 300).map(r => r[c]);
      const n = vals.filter(v => /^\d{4,12}$/.test(String(v ?? '').trim())).length;
      if (n > top) { top = n; best = c; }
    }
    return best;
  }
  function idsFromColumn(rows, c) {
    const result = [];
    for (const row of rows) {
      if (!row || c >= row.length) continue;
      const id = toId(row[c]);
      if (id) result.push(id);
    }
    return unique(result);
  }

  const style = document.createElement('style');
  style.textContent = `
  #smm-date-tool, #smm-date-tool *{box-sizing:border-box}
  #smm-date-tool{--bg:#10192b;--panel:#18243a;--line:#30415b;--txt:#edf4ff;--muted:#9fb0cb;--blue:#4e9dff;--green:#60d6a9;--red:#ff7b83;--yellow:#ffc869;position:fixed;z-index:2147483600;top:78px;right:22px;width:420px;min-width:340px;min-height:115px;max-width:96vw;max-height:92vh;resize:both;overflow:hidden;border:1px solid #47617c;border-radius:16px;background:var(--bg);box-shadow:0 16px 45px #0a152bc4;color:var(--txt);font:13px/1.45 system-ui,-apple-system,'Segoe UI',Arial,sans-serif;direction:ltr}
  #smm-date-tool .head{padding:13px 15px;background:linear-gradient(110deg,#263c5e,#16253e);border-bottom:1px solid #34506e;display:flex;align-items:center;gap:11px;cursor:move;user-select:none;touch-action:none}
  #smm-date-tool .badge{height:32px;width:32px;border-radius:10px;display:grid;place-items:center;background:#3476e244;color:white;font-size:19px}
  #smm-date-tool .head strong{font-size:14px;color:#fff;display:block}
  #smm-date-tool .head small{color:#b3c6e5;display:block;font-size:11px}
  #smm-date-tool .head .spacer{flex:1}
  #smm-date-tool button{font:inherit;color:var(--txt);cursor:pointer;border:1px solid #3e5575;background:#243653;border-radius:9px;padding:8px 10px;transition:filter .12s}
  #smm-date-tool button:hover{filter:brightness(1.22)}
  #smm-date-tool button:disabled{opacity:.4;cursor:not-allowed}
  #smm-date-tool .iconbtn{padding:4px 8px;font-size:17px;background:#253c59}
  #smm-date-tool .main{padding:14px;overflow:auto;max-height:calc(92vh - 61px);height:calc(100% - 59px);min-height:240px}
  #smm-date-tool.collapsed{height:59px!important;min-height:59px!important;resize:none}
  #smm-date-tool.collapsed .main{display:none}
  #smm-date-tool label.caption{display:block;color:#aabdda;font-size:11px;text-transform:uppercase;letter-spacing:.6px;font-weight:700;margin:12px 0 6px}
  #smm-date-tool textarea,#smm-date-tool input[type=number],#smm-date-tool input[type=date],#smm-date-tool select{width:100%;background:#121d30;border:1px solid #3a506e;border-radius:9px;color:white;padding:9px;font:inherit}
  #smm-date-tool textarea{height:65px;min-height:50px;resize:vertical}
  #smm-date-tool input[type=date]{min-width:0;color-scheme:dark}
  #smm-date-tool .date-range{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:10px}
  #smm-date-tool .date-range label{display:block;color:var(--muted);font-size:11px}
  #smm-date-tool input[type=file]{position:absolute;left:-10000px;width:1px;height:1px}
  #smm-date-tool .row{display:flex;gap:7px;align-items:center;flex-wrap:wrap}
  #smm-date-tool .row>*{flex:1 1 auto}
  #smm-date-tool .primary{background:#3479eb;border-color:#3479eb;color:#fff;font-weight:700}
  #smm-date-tool .green{background:#167454;border-color:#26896a;color:#fff}
  #smm-date-tool .warn{background:#694524;border-color:#947046;color:#ffe9c7}
  #smm-date-tool .danger{background:#4b2936;border-color:#754353;color:#ffbec8}
  #smm-date-tool .mini{font-size:11px;padding:5px 8px}
  #smm-date-tool .sub{color:var(--muted);font-size:11px;margin-top:6px}
  #smm-date-tool .stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:6px;margin:11px 0}
  #smm-date-tool .stat{background:#19283f;border:1px solid #2f415a;border-radius:10px;padding:8px 3px;text-align:center}
  #smm-date-tool .stat strong{display:block;font-size:17px;font-variant-numeric:tabular-nums}
  #smm-date-tool .stat span{font-size:10px;color:var(--muted)}
  #smm-date-tool .bar{height:8px;background:#0b1322;border-radius:9px;overflow:hidden;border:1px solid #2c3a54}
  #smm-date-tool .fill{height:100%;background:linear-gradient(90deg,#2c75e9,#68bcf8);width:0%;transition:width .2s}
  #smm-date-tool .pills{display:flex;gap:6px;align-items:center;justify-content:space-between;font-size:11px;color:var(--muted)}
  #smm-date-tool .state{border-radius:8px;padding:8px 10px;border:1px solid #304463;background:#14243b;color:#d6e6ff;min-height:35px;margin:10px 0;overflow-wrap:anywhere}
  #smm-date-tool .options{display:flex;align-items:center;gap:8px;font-size:12px;color:#b5c6df}
  #smm-date-tool input[type=checkbox]{accent-color:#62b0ff}
  #smm-date-tool .log{margin-top:6px;height:145px;overflow:auto;padding:8px;border-radius:9px;background:#090f1b;border:1px solid #2a3c54;font-size:11px;line-height:1.6}
  #smm-date-tool .log div{border-bottom:1px dotted #203248;padding:3px 0;overflow-wrap:anywhere}
  #smm-date-tool .log .error{color:#ff9b99}
  #smm-date-tool .log .ok{color:#89e0ba}
  #smm-date-tool .log .info{color:#b4c8e2}
  #smm-date-tool hr{border:0;border-top:1px solid #2f425b;margin:13px 0}
  #smm-date-tool details{border:1px solid #2c3f58;border-radius:9px;padding:8px;margin:10px 0}
  #smm-date-tool summary{cursor:pointer;color:#c2d2eb}
  #smm-date-tool{position:static;inset:auto;width:100%;min-width:0;min-height:0;max-width:none;max-height:none;height:100%;resize:none;border:0;border-radius:0;box-shadow:none;background:#fff;color:#1f314c;font:13px/1.5 system-ui,sans-serif}
  #smm-date-tool .head{display:none}
  #smm-date-tool .main{height:100%;min-height:0;max-height:none;padding:0 15px 15px;overflow:auto}
  #smm-date-tool .date-nav{padding:0;margin:0 0 10px;border-bottom:1px solid #e5eafa}
  #smm-date-tool .date-nav button{background:none;border:0;border-bottom:2px solid transparent;color:#53687f;padding:10px}
  #smm-date-tool .date-nav button.active{border-bottom-color:#346cdf;color:#245fc6}
  #smm-date-tool .stat{background:#f7f9fd;border-color:#dde5f1}
  #smm-date-tool .stat strong{color:#1e3b66}
  #smm-date-tool .stat span,#smm-date-tool .pills,#smm-date-tool .sub,#smm-date-tool label.caption{color:#657a93}
  #smm-date-tool button{background:#fff;color:#304563;border-color:#d3deeb}
  #smm-date-tool button.primary{background:#346cdf;color:#fff;border-color:#346cdf}
  #smm-date-tool button.green{background:#16835a;color:#fff}
  #smm-date-tool button.danger{background:#fff1ef;color:#9a3c33}
  #smm-date-tool button.warn{background:#fff4e3;color:#7e5321}
  #smm-date-tool textarea,#smm-date-tool input[type=number],#smm-date-tool input[type=date],#smm-date-tool select{background:#fff;border-color:#d4deec;color:#243951;color-scheme:light}
  #smm-date-tool .state{color:#344f6e;background:#f5f8fe;border-color:#d5e1f1}
  #smm-date-tool .log{background:#f5f8fe;border-color:#d5e1f1}
  #smm-date-tool .log .info{color:#526b88} #smm-date-tool .log .ok{color:#147d50} #smm-date-tool .log .error{color:#a63e30}
  #smm-date-tool .options{color:#4c6380} #smm-date-tool summary{color:#3b6097}
  #smm-date-processed{padding:8px 0;overflow-x:auto} #smm-date-processed table{width:100%}
  #smm-date-help{padding:8px 0;color:#475d76;line-height:1.7} #smm-date-help li{margin:7px 0}
  `;
  root.append(style);

  const panel = document.createElement('section');
  panel.id = 'smm-date-tool';
  panel.innerHTML = `
    <div class="head" id="smm-grip">
      <div class="badge">◷</div>
      <div><strong>Редактор дат товаров</strong><small>${SITE} · Quick Edit · v1.2.0</small></div>
      <div class="spacer"></div><button class="iconbtn" id="smm-collapse" title="Свернуть/развернуть" type="button">−</button>
    </div>
    <div class="main">
      <nav class="tabs date-nav"><button data-dtab="work" class="active">Редактор дат</button><button data-dtab="processed">Обработанные</button><button data-dtab="help">Как работать</button></nav>
      <div id="smm-date-work">
      <div class="row"><button id="smm-import" type="button" class="primary">⬆ Загрузить Excel / CSV / TXT</button><input id="smm-file" type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" /></div>
      <div id="smm-file-info" class="sub">Данные сохраняются в этом браузере между перезагрузками.</div>
      <div id="smm-column-wrap" style="display:none"><label class="caption">Столбец с ID</label><select id="smm-column"></select><button id="smm-accept-column" class="mini green" type="button" style="margin-top:7px">Применить этот столбец</button></div>
      <details><summary>Или вставить ID текстом</summary><textarea id="smm-paste" placeholder="285067\n204665\n278600"></textarea><button id="smm-add-pasted" class="mini" type="button">Загрузить ID из поля</button></details>
      <div class="stats"><div class="stat"><strong id="smm-total">0</strong><span>Всего ID</span></div><div class="stat"><strong id="smm-ok">0</strong><span>Готово</span></div><div class="stat"><strong id="smm-missing">0</strong><span>Не найдены</span></div><div class="stat"><strong id="smm-error">0</strong><span>Ошибки</span></div></div>
      <div class="pills"><span id="smm-progress-label">0 / 0</span><span id="smm-progress-pct">0%</span></div><div class="bar"><div class="fill" id="smm-fill"></div></div>
      <div class="state" id="smm-state">Загрузите ID для начала</div>
      <label class="caption">Диапазон случайных дат (включительно)</label>
      <div class="date-range"><label>От<input id="smm-date-from" type="date" /></label><label>До<input id="smm-date-to" type="date" /></label></div>
      <div id="smm-date-range-info" class="sub"></div>
      <div class="options"><label style="display:flex;align-items:center;gap:5px"><input type="checkbox" id="smm-dry" /> Проверка поиска без изменения дат</label></div>
      <div class="row" style="margin-top:10px"><button id="smm-test" class="warn" type="button">Тест: 3 ID</button><button id="smm-start" class="primary" type="button">▶ Старт / продолжить</button><button id="smm-pause" type="button">Ⅱ Пауза</button></div>
      <div class="row" style="margin-top:7px"><button id="smm-report" type="button" class="green">↓ Отчет XLSX</button><button id="smm-missing-csv" type="button">↓ Нет ID (CSV)</button><button id="smm-reset" type="button" class="danger">↺ Сброс</button></div>
      <label class="caption">История операций</label><div class="log" id="smm-log" aria-live="polite"></div>
      <div class="options" style="margin-top:9px"><label>Пауза между ID (мс) <input id="smm-delay" type="number" min="200" max="15000" step="50" style="width:100px" /></label></div>
      <div class="options"><label><input id="smm-fast" type="checkbox" /> Ускоренный поиск (включать после теста 3 ID; есть резервный режим)</label></div>
      <div class="options"><label>Параллельных запросов поиска (1–5) <input id="smm-prefetch" type="number" min="1" max="5" step="1" style="width:70px" /></label></div>
      <div class="sub">Случайная дата внутри диапазона; 00:00 исключено. При 5 последовательных отсутствующих ID — автопауза. Отключите старую отдельную версию редактора дат.</div>
      </div>
      <div id="smm-date-processed" class="hidden"><h3>Обработанные товары</h3><div id="smm-result-table"></div></div>
      <div id="smm-date-help" class="hidden"><h3>Как работать с редактором дат</h3>
        <ol><li>Выберите «2. Редактор дат» на нужном сайте.</li>
        <li>Импортируйте Excel, CSV, TXT с ID товаров или вставьте ID вручную. Выберите колонку при импорте.</li>
        <li>Задайте диапазон дат: обе границы включены, время 00:00 исключено.</li>
        <li>Начните с «Тест: 3 ID». Режим «Проверка поиска» не меняет товары.</li>
        <li>Нажмите «Старт». Задания выполняются в отдельной рабочей вкладке, исходная панель остаётся на месте.</li>
        <li>Пауза запрещает переход к следующему товару, но не отменяет уже отправленный Update.</li>
        <li>Выгрузите XLSX-отчёт или CSV ненайденных ID.</li></ol>
        <p>Изменения выполняются через стандартный WordPress Quick Edit. Ускоренный поиск поддерживает 1–5 параллельных GET-запросов, но сохраняет даты последовательно через Quick Edit. Если этот способ не сработал, используется обычный поиск. При неопределённом ответе Update скрипт останавливается.</p>
      </div>
    </div>`;
  root.querySelector('#date-content').append(panel);

  // Separate module UI uses the shared manager window.
  const dateTabs = panel.querySelectorAll('[data-dtab]');
  const switchDateTab = (target) => {
    for (const t of ['work','processed','help']) $('#smm-date-' + t).classList.toggle('hidden',t !== target);
    dateTabs.forEach(b => b.classList.toggle('active',b.dataset.dtab === target));
    if (target === 'processed') draw();
  };
  dateTabs.forEach(b => b.addEventListener('click',() => switchDateTab(b.dataset.dtab)));
  GM_addValueChangeListener(STORAGE, (key, previous, value, remote) => {
    if (!remote || !value || typeof value !== 'object') return;
    state = { ...emptyState(), ...value };
    draw();
  });

  function draw() {
    if (!panel.isConnected) return;
    $('#smm-total').textContent = state.ids.length;
    $('#smm-ok').textContent = count('updated') + count('found');
    $('#smm-missing').textContent = count('missing');
    $('#smm-error').textContent = count('error');
    $('#smm-progress-label').textContent = `${state.index} / ${state.ids.length}`;
    $('#smm-progress-pct').textContent = `${pct()}%`;
    $('#smm-fill').style.width = `${pct()}%`;
    const m = {running:'● Работает',paused:'Ⅱ Пауза',idle:'Ожидание',done:'✓ Завершено'}[state.status] || state.status;
    const current = state.ids[state.index];
    $('#smm-state').textContent = `${m}${current ? ' · ID ' + current : ''}\n${state.lastMessage}`;
    $('#smm-dry').checked = !!state.dryRun;
    $('#smm-dry').disabled = state.status === 'running';
    $('#smm-delay').value = String(state.delayMs);
    $('#smm-delay').disabled = state.status === 'running';
    $('#smm-fast').checked = !!state.fastMode;
    $('#smm-fast').disabled = state.status === 'running';
    $('#smm-prefetch').value = String(Math.max(1,Math.min(5,Number(state.prefetch)||2)));
    $('#smm-prefetch').disabled = state.status === 'running' || !state.fastMode;
    const resultTable = $('#smm-result-table');
    if (resultTable) {
      const e = val => String(val??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
      resultTable.innerHTML = state.records.length ? '<table><thead><tr><th>ID</th><th>Статус</th><th>Дата</th><th>Информация</th></tr></thead><tbody>' +
        state.records.slice(-150).reverse().map(r=>'<tr><td>'+e(r.id)+'</td><td>'+e(r.status)+'</td><td>'+e(r.date)+'</td><td>'+e(r.detail)+'</td></tr>').join('')+'</tbody></table>' : '<p>Обработанных записей пока нет.</p>';
    }
    const rangeLocked = state.status === 'running';
    $('#smm-date-from').disabled = rangeLocked;
    $('#smm-date-to').disabled = rangeLocked;
    const range = dateRange();
    $('#smm-date-range-info').textContent = range
      ? `Период: ${dateLabel(state.dateFrom)} — ${dateLabel(state.dateTo)} · обе даты включены`
      : 'Ошибка: начальная дата не должна быть позже конечной.';
    $('#smm-pause').disabled = state.status !== 'running';
    $('#smm-test').disabled = !state.ids.length || state.status === 'running';
    $('#smm-start').disabled = !state.ids.length || state.status === 'running' || state.index >= state.ids.length;
    $('#smm-log').replaceChildren(...state.notes.slice(0, 35).map(n => {
      const item = document.createElement('div');
      item.className = n.severity;
      item.textContent = `[${n.time}] ${n.msg}`;
      return item;
    }));
    if (state.source) $('#smm-file-info').textContent = `${state.source} · ${state.ids.length} уникальных ID`;
  }

  function storeImported(ids, source) {
    if (state.status === 'running') { log('Сначала поставьте скрипт на паузу, затем заменяйте список ID', 'error'); return; }
    searchCache.clear();
    ids = unique(ids);
    if (!ids.length) { log('Не удалось найти числовые ID в выбранном столбце', 'error'); return; }
    if (state.records.length && !confirm('Заменить текущий список и сбросить историю? Перед заменой скачайте отчет, если он нужен.')) return;
    const { dateFrom, dateTo, delayMs, fastMode, prefetch } = state;
    state = { ...emptyState(), dateFrom, dateTo, delayMs, fastMode, prefetch };
    state.ids = ids;
    state.source = source;
    state.lastMessage = `Загружено ${ids.length} уникальных ID. Нажмите «Тест: 3 ID» или «Старт».`;
    save(); draw();
    log(`Загружен список: ${source}. ID: ${ids.length}`);
  }
  function showColumnPicker(rows, name, sheet) {
    rowsFromFile = rows;
    activeFileName = name;
    activeSheet = sheet;
    const max = Math.min(30, Math.max(...rows.slice(0, 100).map(r => r.length), 1));
    const sel = $('#smm-column');
    sel.replaceChildren();
    for (let i=0;i<max;i++) {
      const opt = document.createElement('option');
      const examples = rows.slice(0, 25).map(r => String(r[i] ?? '').trim()).filter(Boolean).slice(0, 3).join(' · ');
      opt.value = String(i);
      opt.textContent = `Столбец ${i+1} (${String.fromCharCode(65+i)}) — ${examples.slice(0, 80)}`;
      sel.append(opt);
    }
    sel.value = String(chooseBestColumn(rows));
    $('#smm-column-wrap').style.display = '';
    $('#smm-file-info').textContent = `${name} · лист: ${sheet}. Выберите столбец с ID.`;
    log(`Файл ${name} прочитан: ${rows.length} строк; выбран столбец ${Number(sel.value)+1}`);
  }
  async function onFileChange(file) {
    if (!file) return;
    try {
      let rows, sheetName = 'Текст';
      if (/\.(xlsx|xls|csv|tsv)$/i.test(file.name)) {
        const xlsx = await ensureXLSX();
        const bytes = await file.arrayBuffer();
        const wb = xlsx.read(bytes, {type:'array'});
        sheetName = wb.SheetNames[0];
        rows = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], {header:1, raw:false, defval:''});
      } else {
        const txt = await file.text();
        rows = txt.split(/\r?\n/).map(s => s.split(/[;,\t]/));
      }
      showColumnPicker(rows, file.name, sheetName);
    } catch (e) { log('Ошибка чтения файла: ' + e.message, 'error'); }
  }

  function finish(status, detail = '', date = '') {
    const id = state.ids[state.index];
    if (!id) return;
    state.records.push({id, status, date, detail});
    state.index++;
    state.phase = 'search';
    state.dateCandidate = null;
    if (status === 'missing') state.missStreak++;
    else state.missStreak = 0;
    if (state.batchRemaining != null) state.batchRemaining--;
    const symbols = {updated:'✓',found:'✓',missing:'?',error:'!'};
    log(`${symbols[status]} ID ${id}: ${detail || status}${date ? ' · ' + date : ''}`, status === 'error' || status === 'missing' ? 'error' : 'ok');
    if (state.index >= state.ids.length) {
      state.status = 'done'; state.lastMessage = `Обработаны все ${state.ids.length} ID.`;
      log('Очередь завершена. Можно выгрузить отчет.', 'ok');
    } else if (state.batchRemaining === 0) {
      state.status = 'paused'; state.lastMessage = 'Тестовая партия завершена. Проверьте результат перед продолжением.';
      log('Пауза после тестовых ID');
    } else if (state.missStreak >= 5) {
      state.status = 'paused'; state.lastMessage = 'Пять ID подряд не найдены. Проверьте, поддерживает ли поиск WordPress поиск по ID.';
      log('Автопауза: 5 последовательных ненайденных ID', 'error');
    }
    save(); draw();
  }

  async function waitFor(getResult, timeoutMs, every = 100) {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      const val = getResult();
      if (val) return val;
      await sleep(every);
    }
    throw new Error('Превышено время ожидания ответа страницы');
  }
  function setInputValue(el, val) {
    if (!el) throw new Error('Не найдено поле даты');
    el.value = val;
    el.dispatchEvent(new Event('input', {bubbles:true}));
    el.dispatchEvent(new Event('change', {bubbles:true}));
  }
  function ensureEditorDate(editor, date) {
    const fieldset = $('fieldset.inline-edit-date', editor);
    if (!fieldset) throw new Error('В Quick Edit нет поля даты');
    for (const f of ['mm','jj','aa','hh','mn']) setInputValue($(`[name="${f}"]`, fieldset), date[f]);
    const sec = $('[name="ss"]', fieldset) || $('[name="ss"]', editor);
    if (sec) setInputValue(sec, date.ss);
    // WordPress itself includes edit_date=true in its AJAX Quick Edit request.
  }
  function verifyPostDate(id, expected) {
    const holder = document.getElementById('inline_' + id);
    if (!holder) return null;
    const read = f => $(`.${f}`, holder)?.textContent.trim() || null;
    const vals = Object.fromEntries(['aa','mm','jj','hh','mn'].map(k => [k, read(k)]));
    if (Object.values(vals).some(v => v == null)) return null;
    return ['aa','mm','jj','hh','mn'].every(k => Number(vals[k]) === Number(expected[k]));
  }
  async function processFound(id, row) {
    if (state.dryRun) { finish('found', 'найден, дата не менялась'); return; }
    const editLink = $('.row-actions .editinline', row) || $('.editinline', row);
    if (!editLink) { finish('error', 'у товара нет кнопки Quick Edit'); return; }
    const activeEditor = document.getElementById('edit-' + id);
    if (activeEditor) $('.cancel', activeEditor)?.click();
    state.phase = 'editing'; save();
    editLink.click();
    const editor = await waitFor(() => document.getElementById('edit-' + id), 6000);
    if (state.status !== 'running') { log('Пауза до нажатия Update'); return; }
    // Проверяем также дату, сохранённую предыдущей версией скрипта.
    const date = state.dateCandidate && !isMidnight(state.dateCandidate) ? state.dateCandidate : randomDate();
    state.dateCandidate = date;
    ensureEditorDate(editor, date);
    save(); draw();
    await sleep(220);
    if (state.status !== 'running') { log('Пауза до нажатия Update'); return; }
    const saveButton = $('.inline-edit-save .button-primary.save', editor) || $('button.save', editor);
    if (!saveButton) throw new Error('Не найдена кнопка Update в Quick Edit');
    // Mark pending BEFORE submitting. A reload during AJAX will auto-pause rather than silently re-submit.
    state.phase = 'saving'; save();
    saveButton.click();
    const response = await waitFor(() => {
      const currentEditor = document.getElementById('edit-' + id);
      if (!currentEditor) {
        const updatedRow = document.getElementById('post-' + id);
        return updatedRow ? {type:'ok'} : {type:'error',msg:'Редактор закрылся, но строка товара не появилась'};
      }
      const visibleError = $('.notice-error:not(.hidden) .error, .notice-error:not(.hidden) p.error', currentEditor);
      if (visibleError && visibleError.textContent.trim()) return {type:'error',msg:visibleError.textContent.trim()};
      return null;
    }, 35000, 170);
    if (response.type === 'error') {
      state.phase = 'search'; save();
      finish('error', 'Update: ' + response.msg, date.text);
      return;
    }
    const verified = verifyPostDate(id, date);
    if (verified !== true) {
      state.status = 'paused'; state.phase = 'saving';
      log(`ID ${id}: Update отправлен, но ${verified === false?'дата отличается':'сервер не вернул дату для проверки'}. Проверьте товар вручную; автоматизация остановлена.`, 'error');
      save(); return;
    }
    finish('updated', 'дата сохранена и проверена', date.text);
  }
  async function searchId(id) {
    const field = document.getElementById('post-search-input');
    const submit = document.getElementById('search-submit');
    if (!field || !submit) throw new Error('Не найдена строка поиска #post-search-input / #search-submit');
    // WordPress keeps a hidden Quick Edit template inside the GET form.
    // Disable its controls so post_title, prices and nonce do not leak into search URLs.
    document.querySelectorAll('#inline-edit input, #inline-edit select, #inline-edit textarea, #bulk-edit input, #bulk-edit select, #bulk-edit textarea').forEach(el => {el.disabled = true;});
    // Remove previous query. The search form navigates to the server's result page.
    setInputValue(field, '');
    setInputValue(field, id);
    state.phase = 'search'; save();
    log(`Поиск ID ${id} (${state.index + 1} из ${state.ids.length})`);
    submit.click();
    // If a plugin prevents the search form from navigating, do not silently stall.
    setTimeout(() => {
      if (state.status === 'running' && state.ids[state.index] === id &&
          (new URLSearchParams(location.search).get('s') || '') !== id) {
        state.status = 'paused';
        log(`Поиск ID ${id}: форма не перешла на страницу результатов. Проверьте форму Search products.`, 'error');
      }
    }, 8000);
  }

  // Parallelize read-only searches, never the Quick Edit Update itself.
  const searchCache = new Map();
  async function fetchSearchPage(id) {
    const url = ORIGIN + '/wp-admin/edit.php?post_type=product&s=' + encodeURIComponent(id);
    const response = await fetch(url,{credentials:'same-origin',cache:'no-store',redirect:'follow',signal:AbortSignal.timeout(18000)});
    if (!response.ok || response.url.includes('wp-login.php')) throw new Error('Нет доступа к поиску товаров');
    return (await response.text()).replace(/<img\b[^>]*>/gi,'');
  }
  function prefetchPages() {
    if (!state.fastMode) return;
    const limit = Math.max(1,Math.min(5,Number(state.prefetch)||2));
    const upcoming = state.ids.slice(state.index,state.index+limit);
    for (const key of searchCache.keys()) if (!upcoming.includes(key)) searchCache.delete(key);
    for (const id of upcoming) {
      if (!searchCache.has(id)) searchCache.set(id,fetchSearchPage(id).then(html=>({html}),error=>({error})));
    }
  }
  async function getSearchPage(id) {
    prefetchPages();
    const task = searchCache.get(id) || fetchSearchPage(id).then(html=>({html}),error=>({error}));
    const result = await task;
    searchCache.delete(id);
    if (result.error) throw result.error;
    return result.html;
  }
  // Fast path: fetch just the product-list HTML; never load its images/scripts.
  // Quick Edit / Update is still performed by native WordPress JavaScript.
  async function fastSearchAndEdit(id) {
    if (location.pathname !== '/wp-admin/edit.php' ||
        new URLSearchParams(location.search).get('post_type') !== 'product') return false;
    const list = document.querySelector('#the-list');
    if (!list || !unsafeWindow.inlineEditPost || !document.querySelector('#inline-edit')) return false;
    try {
      const html = await getSearchPage(id);
      const parsed = new DOMParser().parseFromString(html,'text/html');
      if (!parsed.querySelector('#the-list') || parsed.querySelector('#loginform')) throw new Error('Неполный HTML списка WordPress');
      const row = parsed.getElementById('post-' + id);
      if (!row) {
        finish('missing','точный ID не найден (ускоренный поиск)');
        return true;
      }
      if (!row.querySelector('.editinline')) return false;
      const imported = document.importNode(row,true);
      // Do not remove WordPress's hidden Quick Edit/Bulk Edit templates.
      const templates = [...list.querySelectorAll('#inline-edit,#bulk-edit')];
      list.replaceChildren(imported,...templates);
      const inlineMeta = parsed.getElementById('inline_' + id);
      if (inlineMeta && !document.getElementById('inline_' + id)) imported.append(document.importNode(inlineMeta,true));
      log('ID '+id+': строка товара найдена через fetch без полной загрузки страницы');
      await processFound(id, imported);
      return true;
    } catch (e) {
      if (state.phase === 'saving') {
        state.status = 'paused'; save();
        log('Update мог выполниться для ID '+id+'. Повторный запрос заблокирован: '+e.message,'error');
        return true;
      }
      log('Ускоренный поиск недоступен; обычный поиск WordPress: '+e.message);
      return false;
    }
  }
  async function tick() {
    if (busy || state.status !== 'running') return;
    busy = true;
    try {
      while (state.status === 'running' && state.index < state.ids.length) {
        if (unsafeWindow.__SMM_RANDOM_DATES__ || window.__SMM_RANDOM_DATES__) {
          state.status = 'paused'; save();
          log('Найден активный старый скрипт редактора дат. Отключите smmmaster_random_dates.user.js в Tampermonkey.', 'error');
          break;
        }
        const id = state.ids[state.index];
        const visibleId = new URLSearchParams(location.search).get('s') || '';
        if (visibleId !== id) {
          await sleep(Math.min(1000, state.delayMs));
          if (state.status !== 'running') break;
          if (state.fastMode && await fastSearchAndEdit(id)) {
            if (state.status !== 'running') break;
            await sleep(state.delayMs);
            continue;
          }
          await searchId(id);
          return; // Browser navigates; the script auto-resumes on the results page.
        }
        setMessage(`ID ${id}: ищу точное совпадение и открываю Quick Edit...`);
        const row = document.getElementById('post-' + id);
        if (!row) {
          finish('missing', 'точный ID отсутствует в результатах поиска');
        } else {
          try { await processFound(id, row); }
          catch (e) {
            // Saving may have reached the server: don't automatically retry it.
            if (state.phase === 'saving') {
              state.status = 'paused';
              state.lastMessage = `ID ${id}: ответ сервера не получен. Проверьте дату вручную. ${e.message}`;
              log(state.lastMessage, 'error'); save(); break;
            }
            finish('error', e.message);
          }
        }
        if (state.status !== 'running') break;
        await sleep(state.delayMs);
      }
    } catch (e) {
      state.status = 'paused';
      log(`Пауза из-за ошибки: ${e.message}`, 'error'); save();
    } finally { busy = false; draw(); }
  }
  function start(limit = null) {
    if (unsafeWindow.__SMM_RANDOM_DATES__ || window.__SMM_RANDOM_DATES__) {
      log('Обнаружен старый отдельный smmmaster_random_dates.user.js. Отключите его в Tampermonkey, чтобы избежать двойной обработки.', 'error');
      return;
    }
    const courseKey = 'bwo.wc.course.product.v1.' + SITE;
    const courseState = GM_getValue(courseKey, null);
    const courseLease = GM_getValue(courseKey + '.lease', null);
    const coursePool = GM_getValue(courseKey + '.pool', null);
    const liveCoursePool = coursePool?.workers?.some(id => {
      const key = courseKey + '.worker.' + id;
      const lease = GM_getValue(key + '.lease',null);
      const shard = GM_getValue(key,null);
      const live = lease?.expires > Date.now();
      const recentlySaving = shard?.run === coursePool.run && shard?.running &&
        Date.now()-Date.parse(shard.updated||0)<120000;
      return live || recentlySaving;
    });
    if (courseState?.running || (courseLease && courseLease.expires > Date.now()) ||
        liveCoursePool || (coursePool && Date.now()-coursePool.createdAt < 20000)) {
      log('Редактор курсов сейчас выполняет обновления на этом сайте. Дождитесь завершения.', 'error'); return;
    }
    if (!state.ids.length) { log('Сначала загрузите список ID', 'error'); return; }
    if (state.index >= state.ids.length) { log('Все ID обработаны. Сбросьте или загрузите новый список.'); return; }
    if (!dateRange()) { log('Укажите корректный диапазон дат: дата «От» не позже даты «До».', 'error'); return; }
    if (state.phase === 'saving') {
      if (!confirm('Последний Update мог выполниться, но ответ не был получен. Проверьте дату текущего товара на сайте. Повторно попробовать этот ID?')) return;
      state.phase = 'search'; state.dateCandidate = null;
    }
    state.batchRemaining = limit;
    state.status = 'running';
    state.dryRun = $('#smm-dry').checked;
    save(); draw();
    log(`${limit ? 'Тестовая партия: '+limit+' ID' : 'Запуск обработки'}. ${state.dryRun ? 'Только поиск; никаких обновлений' : 'Изменение дат включено: '+dateLabel(state.dateFrom)+' — '+dateLabel(state.dateTo)}`);
    if (IS_DATE_WORKER) {
      tick();
    } else {
      const w = window.open(workerURL, workerName);
      if (!w) {
        state.status = 'paused'; save();
        log('Рабочая вкладка заблокирована браузером. Разрешите всплывающие окна.', 'error');
      }
    }
  }
  function pause() {
    state.status = 'paused';
    state.lastMessage = state.phase === 'saving' ? 'Ожидаем ответа Update; следующий ID автоматически не начнётся.' : 'Остановлено пользователем';
    save(); draw(); log('Пауза по команде пользователя');
  }
  const safeForExcel = value => {
    let s = String(value ?? '');
    if (/^[\s\t\r\n]*[=+@-]/.test(s)) s = "'" + s;
    return s;
  };
  function exportCsv(rows, filename) {
    const csv = '\ufeff' + rows.map(row => row.map(v => `"${safeForExcel(v).replace(/"/g,'""')}"`).join(';')).join('\r\n');
    download(new Blob([csv], {type:'text/csv;charset=utf-8'}), filename);
  }
  function download(blob, filename) {
    const obj = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = obj; a.download = filename;
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(obj), 10000);
  }
  function reportRows(selected = state.records) {
    const labels = {updated:'Обновлено',found:'Найдено, без изменений',missing:'Не найдено',error:'Ошибка'};
    return [['ID','Статус','Новая дата (время сайта)','Детали'], ...selected.map(r => [r.id,labels[r.status] || r.status,r.date || '',r.detail || ''])];
  }
  async function reportExcel() {
    let xlsx;
    try{xlsx=await ensureXLSX();}
    catch(e){
      log('Не загружена Excel-библиотека. Отчет сохранен в CSV: '+e.message, 'error');
      exportCsv(reportRows(), 'smmmaster_report.csv'); return;
    }
    const wb = xlsx.utils.book_new();
    for (const [sheet, subset] of [
      ['Все результаты',state.records], ['Обновлены',state.records.filter(r => r.status === 'updated')],
      ['Не найдены',state.records.filter(r => r.status === 'missing')],
      ['Ошибки',state.records.filter(r => r.status === 'error')]
    ]) {
      const ws = xlsx.utils.aoa_to_sheet(reportRows(subset));
      ws['!cols'] = [{wch:15},{wch:28},{wch:24},{wch:65}];
      xlsx.utils.book_append_sheet(wb, ws, sheet);
    }
    xlsx.writeFile(wb, 'smmmaster_dates_report.xlsx');
    log('Excel-отчёт выгружен.');
  }

  $('#smm-import').addEventListener('click', () => $('#smm-file').click());
  $('#smm-file').addEventListener('change', e => onFileChange(e.target.files[0]));
  $('#smm-accept-column').addEventListener('click', () => {
    const n = Number($('#smm-column').value);
    storeImported(idsFromColumn(rowsFromFile, n), `${activeFileName} · ${activeSheet} · столбец ${n+1}`);
    $('#smm-column-wrap').style.display = 'none';
  });
  $('#smm-add-pasted').addEventListener('click', () => {
    const chunks = $('#smm-paste').value.split(/[\n\r\t;,]+/);
    storeImported(chunks.map(toId).filter(Boolean), 'Вставленный текст');
  });
  for (const [field, key] of [['smm-date-from', 'dateFrom'], ['smm-date-to', 'dateTo']]) {
    const input = $('#' + field);
    input.value = state[key];
    input.addEventListener('change', () => {
      if (state.status === 'running') { input.value = state[key]; return; }
      const candidateFrom = key === 'dateFrom' ? input.value : state.dateFrom;
      const candidateTo = key === 'dateTo' ? input.value : state.dateTo;
      if (!dateRange(candidateFrom, candidateTo)) {
        input.value = state[key];
        log('Неверный диапазон: проверьте обе даты. Начало не может быть позже конца.', 'error');
        return;
      }
      state[key] = input.value;
      state.dateCandidate = null;
      save(); draw();
      log('Диапазон обновлён: ' + dateLabel(state.dateFrom) + ' — ' + dateLabel(state.dateTo));
    });
  }
  $('#smm-test').addEventListener('click', () => start(3));
  $('#smm-start').addEventListener('click', () => start(null));
  $('#smm-pause').addEventListener('click', pause);
  $('#smm-report').addEventListener('click', reportExcel);
  $('#smm-missing-csv').addEventListener('click', () => {
    const missing = state.records.filter(r => r.status === 'missing');
    exportCsv(reportRows(missing), 'smmmaster_not_found_ID.csv');
    log(`Экспортировано ненайденных ID: ${missing.length}`);
  });
  $('#smm-reset').addEventListener('click', () => {
    if (state.status === 'running') { log('Сначала поставьте скрипт на паузу', 'error'); return; }
    if (!confirm('Удалить загруженную очередь и историю результатов в этом браузере?')) return;
    const { dateFrom, dateTo, delayMs, fastMode, prefetch } = state;
    state = { ...emptyState(), dateFrom, dateTo, delayMs, fastMode, prefetch }; save(); draw();
  });
  $('#smm-dry').addEventListener('change', e => { state.dryRun = e.target.checked; save(); });
  $('#smm-delay').addEventListener('change', e => {
    if (state.status === 'running') { e.target.value=state.delayMs; return; }
    state.delayMs=Math.max(200,Math.min(15000,Number(e.target.value)||350)); save(); draw();
  });
  $('#smm-fast').addEventListener('change', e => {
    if (state.status === 'running') { e.target.checked=!!state.fastMode; return; }
    state.fastMode=e.target.checked; save(); draw();
  });
  $('#smm-prefetch').addEventListener('change', e => {
    if (state.status === 'running') { e.target.value=state.prefetch; return; }
    state.prefetch=Math.max(1,Math.min(5,Math.round(Number(e.target.value)||2))); save(); draw();
  });
  if (state.status === 'running' && state.phase === 'saving' && IS_DATE_WORKER) {
    state.status = 'paused';
    state.lastMessage = 'Страница перезагрузилась во время Update. Проверьте текущий ID вручную до продолжения.';
    log(state.lastMessage, 'error');
  }
  draw();
  if (state.status === 'running' && IS_DATE_WORKER &&
      location.pathname === '/wp-admin/edit.php' &&
      new URLSearchParams(location.search).get('post_type') === 'product') {
    log(`Рабочая вкладка: продолжение с ID ${state.ids[state.index]}`);
    setTimeout(tick, 350);
  }
}
  function mount(){
    if(typeof css!=='string'||!css.includes('.panel{')){
      throw new Error('Не загружены стили панели (css). Версия: '+SCRIPT_VERSION);
    }
    const host=document.createElement('div');host.id='bwo-woo-course-manager';host.style.cssText='position:fixed;inset:0;pointer-events:none;z-index:2147483646';document.documentElement.append(host);
    root=host.attachShadow({mode:'open'});
    root.innerHTML=`<style>${css}:host>section,:host>button{pointer-events:auto}</style>
      <button class="pill hidden" id="pill">SMM · открыть менеджер</button>
      <section class="panel" id="panel"><header class="head" id="drag"><div class="logo">W</div><div><b>SMM Course Manager · Курсы и даты</b><small>${esc(SITE)} · v1.5.2</small></div><span class="space"></span><button id="dashboard" title="Открыть отдельную вкладку">↗ Отдельная вкладка</button><button id="full" title="На весь экран">⛶</button><button id="min" title="Свернуть">—</button></header>
      <nav class="module-nav"><div id="module-buttons" class="module-actions"></div><span class="spacer"></span><span class="hint">${esc(SITE)}</span></nav>
      <div class="body" id="course-content"><div class="stats"><div class="stat"><b id="total">0</b><small>Всего товаров</small></div><div class="stat success"><b id="done">0</b><small>Готово</small></div><div class="stat"><b id="pending">0</b><small>В очереди / работе</small></div><div class="stat bad"><b id="failed">0</b><small>Ошибки / проверка</small></div></div>
      <div class="progress"><i id="bar"></i></div><div class="current" id="current">Загрузите Excel с товарами</div>
      <div class="actions"><button class="primary" id="start">▶ Начать</button><button id="one">Один товар</button><button id="verifyStalled">Проверить зависшие (GET)</button><button id="pause">Ⅱ Пауза</button><button class="danger" id="stop">■ Стоп</button><button id="retry">Повторить ошибки</button><button id="report">↓ Отчёт</button></div>
      <nav class="tabs"><button data-tab="queue">Очередь</button><button data-tab="processed">Обработанные</button><button data-tab="import" class="active">Импорт Excel</button><button data-tab="log">Журнал</button><button data-tab="settings">Настройки</button><button data-tab="help">Как работать</button></nav>
      <div class="view hidden" id="queue"><div class="tools"><input type="search" id="search" placeholder="Название, ID или img…"><select id="filter"><option value="all">Все статусы</option><option value="pending">В очереди</option><option value="error">Ошибка</option><option value="skipped">Пропущено</option></select></div><div id="rows"></div></div>
      <div class="view" id="import"><div class="tools"><button class="danger" id="reset">Очистить очередь и историю</button></div><div class="drop" id="drop"><label for="file">Выбрать Excel-файл</label><input id="file" type="file" accept=".xlsx,.xls,.csv" hidden><small>Или перетащите сюда · .xlsx / .xls / .csv</small><small id="filename">Данные Excel обрабатываются локально</small></div>
      <div id="mapping" class="hidden"><div class="mapping"><label>ID товара</label><select id="map-id"></select><label>Name</label><select id="map-name"></select><label>Description (HTML)</label><select id="map-description"></select><label>Regular price</label><select id="map-price"></select><label>lessons</label><select id="map-lessons"></select><label>img (14 цифр)</label><select id="map-img"></select><label>Лист</label><select id="sheet"></select></div><p class="hint">При импорте цена 3.42€ распознаётся как 3.42, а в поле WooCommerce записывается 3,42. Колонка img — имя загруженного изображения из 14 цифр, НЕ WordPress attachment ID. Очередь перемешивается автоматически.</p><button class="primary" id="load">Загрузить в очередь</button></div>
      <div class="notice">Строки проверяются до запуска. Новая загрузка заменяет текущую очередь; скачайте отчёт, если он нужен. Обрабатываются только товары сайта <b>${esc(SITE)}</b>.</div></div>
      <div class="view hidden" id="log"><div class="tools"><button id="downloadLog">↓ Скачать журнал</button><span class="hint">Последние 2000 событий</span></div><div id="logs"></div></div>
      <div class="view hidden" id="settings"><label class="setting"><div>Параллельных потоков<small>От 1 до 5 рабочих вкладок. Каждая вкладка обрабатывает много курсов по очереди — новые вкладки не создаются для каждого товара.</small></div><input id="workers" type="number" min="1" max="5" step="1" value="1"></label><label class="setting"><div>Пауза между товарами<small>В секундах после проверки результата</small></div><input id="delay" type="number" min="1" max="60" value="3"></label><label class="setting"><input id="stopOnError" type="checkbox" checked><div>Останавливаться при ошибке<small>Рекомендуется: ошибка может оставить несохранённые изменения в редакторе</small></div></label><div class="setting"><div>Перемешивание Excel<small>Однократно при каждом импорте; исходные номера строк сохраняются</small></div><span class="tag done">Всегда</span></div><div class="setting"><div>Downloadable<small>Флажок снимается у каждого товара</small></div><span class="tag done">Всегда</span></div><button id="backup">↓ Резервная копия JSON</button></div>
      <div class="view help hidden" id="help"><h3>1. Подготовьте Excel</h3><p>Первая строка: <b>ID, Name, Description, Regular price, lessons, img</b>. ID — существующий post ID товара. Description — HTML, как в редакторе WordPress. Regular price может быть с € или запятой. lessons — целое положительное число. img — 14-значное имя уже загруженной фотографии в Media Library (не внутренний attachment ID).</p>
      <h3>2. Импортируйте товары</h3><ol><li>Войдите в <b>${esc(SITE)}</b> как администратор.</li><li>Откройте «Импорт Excel», выберите файл, лист и шесть колонок.</li><li>Нажмите «Загрузить в очередь». Курсы будут перемешаны; номера исходных строк останутся в отчёте.</li><li>Нажмите «Один товар» и проверьте сохранённые поля на сайте; после этого нажмите «Начать» для остальных.</li></ol>
      <h3>3. Как работают потоки</h3><p>На вкладке «Настройки» выберите от 1 до 5 потоков. Один поток использует одну рабочую вкладку и последовательно обрабатывает все назначенные курсы. Например, для 400 товаров и 5 потоков создаётся пять вкладок примерно по 80 товаров. Все следующие товары открываются переходом в тех же вкладках. Не закрывайте рабочие вкладки до завершения.</p>
      <h3>4. Что делает скрипт</h3><p>Открывает <code>/wp-admin/post.php?post=ID&amp;action=edit</code>. Заполняет Name и slug, переключает редактор Description в Text, вставляет HTML и возвращается в Visual. Меняет Regular price и ACF lessons, снимает Downloadable, удаляет старую обложку, ищет точную фотографию по img и нажимает Set product image. После проверки полей нажимает Update и перезагружает товар для сверки.</p>
      <h3>5. Пауза, ошибки и безопасность</h3><p>Не редактируйте страницу вручную во время работы. Пауза останавливает новые действия, но не отменяет отправленный Update. После сохранения статус «Выполнено» ставится только после сверки на новой странице. Если данные не совпали — «Проверить», повторный Update автоматически не выполняется. Для такого товара доступна проверка без Update через «⋯».</p><p>Если произошла ошибка, исправьте причину, затем «Повторить ошибки». Неверные строки Excel исправляются только повторным импортом. Скачивайте отчёт и журнал. Очереди и настройки независимы для каждого сайта: <b>smmmaster.net, hottestcourse.com, learnmarket.net, profcourse.com, bestsocials.online</b>.</p>
      <h3>Ограничения</h3><p>Скрипт рассчитан на Classic WordPress / WooCommerce и указанные элементы ACF. В отдельных темах или версиях админки разметка может отличаться. При отсутствии нужных полей скрипт останавливается до Update.</p></div>
      <footer class="foot"><span id="updated">Очередь хранится локально</span><button id="prev">←</button><span id="pages" style="flex:none">1 / 1</span><button id="next">→</button></footer></div><div class="body hidden" id="date-content"></div></section>`;
    panel=$('panel');
    const on=(id,ev,fn)=>$(id).addEventListener(ev,fn);
    on('start','click',()=>launchCourse(false));on('one','click',()=>launchCourse(true));
    on('verifyStalled','click',()=>verifyStalledFromDashboard().catch(ex=>notice(ex.message||String(ex))));
    on('pause','click',()=>stopCourseFromPanel(true));on('stop','click',()=>stopCourseFromPanel(false));
    on('retry','click',retryErrors);on('report','click',exportReport);on('downloadLog','click',exportLog);on('reset','click',resetAll);
    on('load','click',loadQueue);on('backup','click',()=>download(JSON.stringify(IS_COURSE_WORKER?state:mergedCourseState(),null,2),`${SITE}-woo-products-backup.json`,'application/json'));
    on('file','change',e=>readFile(e.target.files[0]));on('sheet','change',setMapping);
    const drop=$('drop');drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('over');});drop.addEventListener('dragleave',()=>drop.classList.remove('over'));drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('over');readFile(e.dataTransfer.files[0]);});
    root.querySelectorAll('[data-tab]').forEach(b=>b.addEventListener('click',()=>switchTab(b.dataset.tab)));
    on('search','input',e=>{search=e.target.value;page=0;renderRows();});on('filter','change',e=>{filter=e.target.value;page=0;renderRows();});
    on('prev','click',()=>{page=Math.max(0,page-1);renderRows();});on('next','click',()=>{page++;renderRows();});
    on('delay','change',e=>{if(!editable())return;state.settings.delay=Math.min(60,Math.max(1,Number(e.target.value)||3));persist();render();});
    on('workers','change',e=>{if(!editable())return;state.settings.workers=Math.max(1,Math.min(5,Math.round(Number(e.target.value)||1)));persist();render();});
    on('stopOnError','change',e=>{if(!editable())return;state.settings.stopOnError=e.target.checked;persist();render();});
    on('min','click',()=>{panel.classList.add('hidden');$('pill').classList.remove('hidden');});
    const show=()=>{panel.classList.remove('hidden');$('pill').classList.add('hidden');};
    on('pill','click',show);
    on('dashboard','click',openDashboard);
    MODULES.forEach((m,i)=>{
      const button=document.createElement('button');
      button.type='button';button.dataset.module=m.id;button.textContent=(i+1)+'. '+m.title;
      $('module-buttons').append(button);
      button.addEventListener('click',()=>selectModule(m.id));
      GM_registerMenuCommand((i+1)+'. '+m.title,()=>{show();selectModule(m.id);});
    });
    GM_registerMenuCommand('Открыть менеджер в отдельной вкладке',openDashboard);
    on('full','click',()=>{maximized=!maximized;panel.classList.toggle('full',maximized);panel.style.top='';panel.style.left='';panel.style.right='';panel.style.bottom='';});
    let drag=null;
    on('drag','pointerdown',e=>{if(e.target.closest('button')||maximized)return;const rect=panel.getBoundingClientRect();drag={x:e.clientX-rect.left,y:e.clientY-rect.top};e.currentTarget.setPointerCapture(e.pointerId);});
    on('drag','pointermove',e=>{if(!drag)return;panel.style.left=Math.max(0,Math.min(innerWidth-panel.offsetWidth,e.clientX-drag.x))+'px';panel.style.top=Math.max(0,Math.min(innerHeight-panel.offsetHeight,e.clientY-drag.y))+'px';panel.style.right='auto';panel.style.bottom='auto';});
    on('drag','pointerup',()=>drag=null);
    MODULES.forEach(m=>m.mount?.());
    if(IS_DASHBOARD){
      panel.classList.add('dashboard');
      document.documentElement.style.overflow='hidden';
      $('dashboard').textContent='Панель открыта отдельно';
      $('dashboard').disabled=true;
    } else if(IS_COURSE_WORKER || SMM_PARAMS.get('smm_manager')==='date-worker') {
      // Workers use the compact launcher while editing, but retain a visible way to show the full panel.
      panel.classList.add('hidden');$('pill').classList.remove('hidden');
      $('pill').textContent='SMM Manager · Открыть панель';
    }
    selectModule(GM_getValue(MODE_KEY,'courses'));
    render();if(state.rows.length)switchTab('queue');
    if(!IS_COURSE_WORKER)setInterval(()=>{if(!working&&root)render();},3000);
  }
  function captureErrors(){
    const seen=new Map();function record(text){if(!working)return;if(/Video loading failed/i.test(text))return;const at=seen.get(text)||0;if(Date.now()-at<30000)return;seen.set(text,Date.now());log('warn',`Предупреждение страницы: ${text}`,state.rows[state.index]);}
    window.addEventListener('error',e=>{if(e.message)record(e.message);},true);
    window.addEventListener('unhandledrejection',e=>record(String(e.reason?.message||e.reason||'Unhandled rejection')));
  }
  let bootStarted=false;
  async function bootManager(data={}){
    if(bootStarted)return;
    bootStarted=true;
    try{
    const tabData=data&&typeof data==='object'?data:{};
    let ticket;try{ticket=JSON.parse(sessionStorage.getItem(TICKET)||'null');}catch{ticket=null;}
    // Upgrading from versions before 1.3.9: the navigation ticket contains the
    // previous tab identity. Reuse it so a live lease continues on the same tab.
    const savedWorkerId=IS_COURSE_WORKER?sessionStorage.getItem(WORKER_TAB_SESSION):null;
    const ticketTab=IS_COURSE_WORKER && ticket?.run===state.run && ticket?.tab ? ticket.tab : null;
    tabId=savedWorkerId||ticketTab||tabData.wcCourseProductTabId||uid();
    if(IS_COURSE_WORKER)sessionStorage.setItem(WORKER_TAB_SESSION,tabId);
    tabData.wcCourseProductTabId=tabId;
    try{GM_saveTab(tabData);}catch(err){console.warn('[SMM Course Manager] GM_saveTab:',err);}
    sessionStorage.removeItem(TICKET);
    mount();captureErrors();
    if(IS_COURSE_WORKER){
      GM_addValueChangeListener(COMMAND,(name,oldValue,newValue,remote)=>{
        if(remote&&working&&newValue?.run===state.run&&['pause','stop'].includes(newValue.kind))
          stop(newValue.kind==='pause');
      });
    } else {
      const refresh=()=>{if(!working){state=mergedCourseState();render();}};
      GM_addValueChangeListener(MASTER_KEY,(name,oldValue,newValue,remote)=>{if(remote)refresh();});
      GM_addValueChangeListener(POOL_KEY,(name,oldValue,newValue,remote)=>{if(remote)refresh();});
      for(let i=1;i<=5;i++)
        GM_addValueChangeListener(workerKey(i),(name,oldValue,newValue,remote)=>{if(remote)refresh();});
    }
    const row=state.rows[state.index];
    // An old ticket can survive if the same named tab is reused for a new run.
    // A fresh worker launch URL must not be blocked by a stale ticket.
    if(IS_COURSE_WORKER&&SMM_PARAMS.get('smm_manager')==='course-worker'&&
      (!ticket || ticket.run!==state.run || ticket.id!==row?.id)){
      if(SMM_PARAMS.get('smm_run')===state.run){
        await start(SMM_PARAMS.get('smm_one')==='1');
      } else notice('Ссылка рабочего процесса устарела — задача не запущена.');
    }else if(IS_COURSE_WORKER&&ticket&&row&&ticket.tab===tabId&&ticket.run===state.run&&ticket.id===row.id&&Date.now()-ticket.at<180000){
      log('info','Новая страница в том же рабочем потоке: продолжаем очередь.',row);
      await start(state.single,true);
    }else if(IS_COURSE_WORKER&&state.running&&row&&
      ['saving','verifying'].includes(row.status)&&row.expected&&
      productIDFromURL(location.href)===row.id&&currentPool()?.run===state.run){
      // Ticket can be lost across WP redirects; verify only, never re-submit Update.
      log('warn','Билет перехода не восстановился. Проверяю сохранение на открытой странице без повторного Update.',row);
      await start(state.single,true);
    }else if(IS_COURSE_WORKER&&state.running){
      state.running=false;state.paused=true;persist();
      log('warn','Автопереход не подтверждён. Поток приостановлен; дальнейшие Update не выполняются.',row);
    }
    bootLauncher.remove();
    }catch(err){fatalLauncher(err);}
  }
  try{
    if(typeof GM_getTab==='function')GM_getTab(data=>{bootManager(data).catch(fatalLauncher);});
    else bootManager({}).catch(fatalLauncher);
  }catch(err){bootManager({}).catch(fatalLauncher);}
  // Use the same boot fallback on workers: GM_getTab callback can fail after WP redirects.
  // Worker identity and the ticket survive via sessionStorage.
  setTimeout(()=>{if(!bootStarted)bootManager({}).catch(fatalLauncher);},3500);
})();