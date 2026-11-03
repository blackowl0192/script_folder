// ==UserScript==
// @name         Tutor LMS — Courses & Tutors Manager (Multi-site)
// @namespace    yourcourse.local.tools
// @version      1.2.3
// @description  Excel → Tutor LMS на 6 сайтах: курсы, преподаватели, фото, Bio, проверки, отдельные очереди по доменам.
// @match        https://yourcourse.shop/wp-admin/*
// @match        https://yourcourse.shop/wp-login.php*
// @match        https://yourlibrarium.com/wp-admin/*
// @match        https://yourlibrarium.com/wp-login.php*
// @match        https://learnmarket.net/wp-admin/*
// @match        https://learnmarket.net/wp-login.php*
// @match        https://bestseminar.net/wp-admin/*
// @match        https://bestseminar.net/wp-login.php*
// @match        https://educationwebs.com/wp-admin/*
// @match        https://educationwebs.com/wp-login.php*
// @match        https://easycourses.pro/wp-admin/*
// @match        https://easycourses.pro/wp-login.php*
// @updateURL    https://raw.githubusercontent.com/blackowl0192/script_folder/main/course-tutor-manager.user.js
// @downloadURL  https://raw.githubusercontent.com/blackowl0192/script_folder/main/course-tutor-manager.user.js
// @run-at       document-idle
// @noframes
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_getTab
// @grant        GM_saveTab
// @grant        GM_registerMenuCommand
// @grant        unsafeWindow
// ==/UserScript==

/* Установка: Tampermonkey → заменить исходный скрипт целиком → Ctrl+S.
 * Режим «Курсы»: импортируйте Excel (очередь перемешивается автоматически) и нажмите «Один курс»
 * либо «Начать». Режим «Преподаватели»: Excel с url_edit, login, name, bio, img.
 * На странице /wp-admin/users.php?role=tutor_instructor выбирайте вкладку «Преподаватели».
 * Один скрипт поддерживает 6 доменов. Загружайте Excel со ссылками только текущего сайта.
 * Очереди и настройки для каждого домена независимы, ваши данные yourcourse.shop сохранены.
 * Инструкции доступны во вкладке «Как работать» отдельно для курсов и преподавателей.
 * Description содержит HTML. Пустые поля пропускаются; Sale Price очищается всегда.
 * «Стоп» не отменяет уже отправленный Update. После сохранения страница обновляется:
 * только повторное чтение сохранённых данных даёт статус «Выполнено».
 * Состояние хранится в Tampermonkey локально. Не удаляйте скрипт до экспорта отчёта.
 * При новом запуске браузера продолжение вручную. Обе очереди запускаются по очереди в одной вкладке.
 * Данные Excel не отправляются сторонним сервисам. SheetJS загружается при установке.
 */

(() => {
  'use strict';
  const ORIGIN = location.origin;
  const SUPPORTED_HOSTS = new Set(['yourcourse.shop','yourlibrarium.com','learnmarket.net',
    'bestseminar.net','educationwebs.com','easycourses.pro']);
  if(location.protocol!=='https:'||!SUPPORTED_HOSTS.has(location.hostname))return;
  // Keep original yourcourse.shop keys so upgrading from 1.1.5 does not erase queues.
  // GM storage is shared across @match sites; other domains MUST use namespaced keys.
  const SITE_SUFFIX = location.hostname==='yourcourse.shop'?'':'.'+location.hostname;
  const KEY = 'yc.course.manager.v1'+SITE_SUFFIX, LEASE = KEY + '.lease', TICKET = KEY + '.ticket';
  const EXPORT_PREFIX = location.hostname.replace(/[^a-z0-9.-]/gi,'_')+'-';
  const LABEL = {pending:'В очереди', opening:'Открывается', filling:'Заполняется', saving:'Сохраняется', verifying:'Проверяется', done:'Выполнено', error:'Ошибка', review:'Проверить', skipped:'Пропущено'};
  const ACTIVE = new Set(['opening','filling','saving','verifying']);
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const uid = () => crypto.randomUUID();
  const DOCUMENT_ID = uid();
  const esc = s => String(s ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fail = (message, fatal=false) => Object.assign(new Error(message),{fatal});
  const STOP = Symbol('stop');
  const empty = () => ({version:1, rows:[], logs:[], running:false, owner:null, index:-1, phase:'idle', filename:'', settings:{delay:3, stopOnError:false}, updated:''});
  let state = GM_getValue(KEY, null) || empty();
  let tabId, busy=false, leaving=false, halted=false, single=false, leaseTimer, panel, root;
  let page=0, filter='all', search='', activeTab='queue', book=null, importName='', expanded=false;

  function price(value) {
    // Excel may display price as "3.42€", "€ 3,42" or "3.42 EUR".
    // Keep a decimal string with a dot and exactly two fraction digits for Tutor LMS.
    let s=String(value ?? '').trim();
    if(!s) return null;
    s=s.replace(/\p{Sc}/gu,'').replace(/\b(?:EUR|USD|GBP|MDL|RON|RUB)\b/gi,'')
       .replace(/[\s\u00a0\u202f\u2007\u2009]/g,'');
    if(!/^\d+(?:[.,]\d{1,2})?$/.test(s)) {
      throw fail('Некорректная цена: допустимо число с максимум двумя десятичными знаками (например 3.42€, 3,42 € или €3.42).');
    }
    const [a,b='']=s.replace(',','.').split('.');
    if(a.length>9) throw fail('Слишком большая цена.');
    return `${String(Number(a))}.${b.padEnd(2,'0')}`;
  }
  function slug(value) {
    const map={а:'a',б:'b',в:'v',г:'g',д:'d',е:'e',ё:'e',ж:'zh',з:'z',и:'i',й:'y',к:'k',л:'l',м:'m',н:'n',о:'o',п:'p',р:'r',с:'s',т:'t',у:'u',ф:'f',х:'kh',ц:'ts',ч:'ch',ш:'sh',щ:'shch',ъ:'',ы:'y',ь:'',э:'e',ю:'yu',я:'ya',і:'i',ї:'yi',є:'ye'};
    return String(value).toLowerCase().replace(/[а-яёіїє]/g,c=>map[c]).normalize('NFKD').replace(/[\u0300-\u036f]/g,'').replace(/&/g,' and ').replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
  }
  function courseURL(value) {
    let u; try {u=new URL(String(value).trim());} catch {throw fail('Некорректный Course URL.');}
    if(u.origin!==ORIGIN || u.pathname!=='/wp-admin/admin.php' || u.searchParams.get('page')!=='create-course' || !/^\d+$/.test(u.searchParams.get('course_id')||'')) throw fail('Course URL должен вести на текущий сайт '+location.hostname+'/wp-admin/admin.php?page=create-course&course_id=…; ссылки с другого сайта запрещены.');
    return {id:u.searchParams.get('course_id'), url:`${ORIGIN}/wp-admin/admin.php?page=create-course&course_id=${u.searchParams.get('course_id')}`};
  }
  function normalizeHeader(s){return String(s??'').toLowerCase().replace(/[\s_\-]+/g,'');}
  function cleanHTML(h) {
    const d=new DOMParser().parseFromString(String(h),'text/html');
    if(d.querySelector('script,iframe,object,embed,form,base,meta,link')) throw fail('Description содержит активные HTML-элементы. Удалите script/iframe/form и подобные теги.');
    for(const e of d.body.querySelectorAll('*')) for(const a of e.attributes) {
      if(/^on/i.test(a.name)||(/^(href|src|action|xlink:href)$/i.test(a.name)&&/^\s*(javascript|vbscript|data):/i.test(a.value))) throw fail('Description содержит обработчик события или небезопасный адрес.');
    }
    return String(h);
  }
  function htmlSignature(h) {
    const d=new DOMParser().parseFromString(String(h),'text/html');
    const walk=n=>{
      if(n.nodeType===3)return n.nodeValue.replace(/\u00a0/g,' ').replace(/\s+/g,' ');
      if(n.nodeType!==1)return '';
      let tag=n.tagName.toLowerCase(); if(tag==='b')tag='strong';if(tag==='i')tag='em';
      const attrs=[...n.attributes].filter(a=>!/^data-mce-|^id$/.test(a.name)).map(a=>[a.name,a.value]).sort((a,b)=>a[0].localeCompare(b[0]));
      return `<${tag}${attrs.map(([k,v])=>` ${k}=${JSON.stringify(v)}`).join('')}>${[...n.childNodes].map(walk).join('').trim()}</${tag}>`;
    };
    let v=[...d.body.childNodes].map(walk).join('').trim();
    if(!d.body.children.length && v)v=`<p>${v}</p>`;
    return v.replace(/>\s+</g,'><');
  }
  function textSignature(h){return new DOMParser().parseFromString(h,'text/html').body.textContent.replace(/\s+/g,' ').trim();}
  function visible(e){return !!(e && e.isConnected && e.getClientRects().length && getComputedStyle(e).visibility!=='hidden');}
  function one(selector,scope=document,required=true){const a=[...scope.querySelectorAll(selector)].filter(visible);if(a.length===1)return a[0];if(required)throw fail(`Ожидался один элемент ${selector}, найдено: ${a.length}`);return null;}
  function authCheck(){if(location.pathname.includes('wp-login.php')||document.querySelector('#loginform'))throw fail('Требуется вход администратора. Войдите и нажмите «Продолжить».',true);}
  function owned(){const l=GM_getValue(LEASE,null);return l && l.tab===tabId && l.expires>Date.now();}
  function guard(allowStop=false){authCheck();if(!owned())throw fail('Управление очередью перешло другой вкладке.',true);if(halted&&!allowStop)throw STOP;}
  function persist(){state.updated=new Date().toISOString();GM_setValue(KEY,state);}
  function log(level,message,row=null){state.logs.push({time:new Date().toISOString(),level,row:row?.excelRow||'',course:row?.id||'',stage:state.phase,message:String(message).slice(0,1600)});if(state.logs.length>1500)state.logs.splice(0,state.logs.length-1500);persist();render();}
  function stage(r,status,message){guard(status==='saving'||status==='verifying');r.status=status;state.phase=status;r.message=message;log('info',message,r);}
  async function waitFor(fn,label,timeout=25000,allowStop=false){const start=Date.now();while(Date.now()-start<timeout){guard(allowStop);const v=fn();if(v)return v;await sleep(250);}throw fail(`Время ожидания истекло: ${label}`);}
  function nativeValue(e,value){
    const proto=e.tagName==='TEXTAREA'?unsafeWindow.HTMLTextAreaElement.prototype:unsafeWindow.HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto,'value').set.call(e,String(value));
    e.dispatchEvent(new unsafeWindow.Event('input',{bubbles:true}));e.dispatchEvent(new unsafeWindow.Event('change',{bubbles:true}));e.dispatchEvent(new unsafeWindow.Event('blur',{bubbles:true}));
  }
  async function field(selector,value){guard();const e=one(selector);e.focus();nativeValue(e,value);await sleep(250);guard();if(one(selector).value!==String(value))throw fail(`Поле не приняло значение: ${selector}`);}
  function editorRef(){const a=[...document.querySelectorAll('textarea[data-cy="tutor-tinymce"]')];if(a.length!==1)throw fail(`Найдено редакторов описания: ${a.length}; нужен один.`);return a[0];}
  function editorButton(id,mode){return [...document.querySelectorAll(`button.switch-${mode}`)].find(b=>b.dataset.wpEditorId===id&&visible(b));}
  function tiny(id){return unsafeWindow.tinymce?.get(id);}
  async function readDescription(allowStop=false){
    const ta=editorRef();const ed=tiny(ta.id);
    if(ed && !ed.isHidden())return ed.getContent();
    if(visible(ta))return ta.value;
    const b=editorButton(ta.id,'html');if(!b)throw fail('Не найдена кнопка Text для Description.');
    guard(allowStop);b.click();await waitFor(()=>visible(ta),'режим Text',10000,allowStop);return ta.value;
  }
  async function writeDescription(value){
    let ta=editorRef();const id=ta.id;
    const textButton=editorButton(id,'html');if(!textButton)throw fail('Не найдена кнопка Text.');
    guard();textButton.click();await waitFor(()=>visible(ta),'текстовое поле Description');
    ta=editorRef();nativeValue(ta,value);await sleep(300);guard();
    const vb=editorButton(id,'tmce');if(!vb)throw fail('Не найдена кнопка Visual.');vb.click();
    await waitFor(()=>{const e=tiny(editorRef().id);return e&&!e.isHidden()&&e.initialized?e:null;},'Visual / TinyMCE',20000);
    // Notify the visual editor's own integration. Do not dispatch a second native
    // textarea change after Visual: it can race the controlled form's state.
    for(let attempt=0;attempt<2;attempt++){
      guard();const ed=tiny(editorRef().id);
      if(!ed||ed.isHidden())throw fail('Visual перестал быть активным. Update не нажат.');
      ed.setContent(value,{format:'html'});
      const html=ed.getContent();
      if(htmlSignature(html)!==htmlSignature(value))throw fail(descriptionMismatch(value,html,'Visual изменил разметку'));
      ed.setDirty?.(true);
      ed.fire('input');ed.fire('change',{level:{content:html},lastLevel:null});ed.fire('keyup',{key:'',keyCode:0});
      ed.save();ed.fire('blur');
      // Re-acquire the editor after the host framework has processed its events.
      let stable=true;
      for(let n=0;n<4;n++){
        await sleep(300);guard();const current=tiny(editorRef().id);
        if(!current||current.isHidden()||htmlSignature(current.getContent())!==htmlSignature(html)){stable=false;break;}
      }
      if(stable)return html;
      log('warn','Описание изменилось после синхронизации Visual. Повторяем синхронизацию один раз.',state.rows[state.index]);
    }
    throw fail(descriptionMismatch(value,await readDescription(),'Описание сбрасывается после Visual'));
  }
  function descriptionMismatch(expected,actual,label){
    const a=textSignature(expected),b=textSignature(actual);
    let at=0;while(at<Math.min(a.length,b.length)&&a[at]===b[at])at++;
    return `${label}. ${a===b?'Текст совпадает, отличается HTML-разметка.':`Текст различается: ожидалось ${a.length} символов, получено ${b.length}; первое различие на позиции ${at+1}.`} Update не нажат.`;
  }
  function publicLink(){return one('a[data-cy="course-slug"]',document,false);}
  function linkSlug(){const a=publicLink();if(!a)return '';try{return decodeURIComponent(new URL(a.href).pathname.split('/').filter(Boolean).pop()||'');}catch{return '';}}
  function editSlugButton(){
    const a=publicLink();if(!a)throw fail('Не найдена ссылка Course URL.');
    let p=a.parentElement;
    for(let depth=0;p&&depth<5;depth++,p=p.parentElement){
      const buttons=[...p.querySelectorAll('button')].filter(b=>visible(b)&&b.querySelector('svg')&&!b.matches('[data-cy="dropdown-trigger"],[data-cy="course-builder-submit-button"]'));
      if(buttons.length===1)return buttons[0];
      // Match the supplied pencil geometry only within the URL block.
      const pencils=buttons.filter(b=>[...b.querySelectorAll('path')].some(path=>(path.getAttribute('d')||'').startsWith('M15.9288 4.20041')));
      if(pencils.length===1)return pencils[0];
      if(buttons.length>1)throw fail('Несколько кнопок рядом с URL. Не удалось однозначно найти карандаш.');
    }
    throw fail('Не найдена кнопка редактирования URL.');
  }
  function slugSave(input){let p=input.parentElement;for(let i=0;p&&i<6;i++,p=p.parentElement){const bs=[...p.querySelectorAll('button')].filter(b=>visible(b)&&/^(Save|Сохранить)$/i.test(b.textContent.trim()));if(bs.length===1)return bs[0];if(bs.length>1)break;}throw fail('Не найдена однозначная кнопка Save рядом с post_name.');}
  async function writeSlug(value){
    if(linkSlug()===value)return;
    let input=one('input[name="post_name"]',document,false);
    if(!input){guard();editSlugButton().click();input=await waitFor(()=>one('input[name="post_name"]',document,false),'поле slug');}
    await field('input[name="post_name"]',value);guard();slugSave(input).click();
    await waitFor(()=>!one('input[name="post_name"]',document,false)&&linkSlug()===value,'сохранение slug',20000);
  }
  function imageKey(src){
    if(!src)return '';try{const u=new URL(src,ORIGIN);return u.origin+decodeURIComponent(u.pathname).replace(/-\d+x\d+(?=\.[^.]+$)/,'');}catch{return String(src);}
  }
  // Compare real image identity rather than exact thumbnail/CDN URL.
  function courseImageStem(src){
    if(!src)return '';
    try{
      const name=decodeURIComponent(String(src).split(/[?#]/)[0].split('/').pop()||'');
      return name.replace(/\.(?:jpe?g|png|webp|gif|avif)$/i,'')
        .replace(/(?:-(?:\d+x\d+|scaled|rotated))+$/gi,'').toLowerCase();
    }catch{return '';}
  }
  function sameCourseImage(actual,expected){
    const a=courseImageStem(actual),b=courseImageStem(expected);
    return !!(a&&b&&a===b);
  }
  function imagePreview(preferredKey=''){
    const all=[...document.querySelectorAll('[data-cy="media-preview"]')];
    const marked=all.filter(e=>e.querySelector('[data-cy="replace-media"]'));
    if(marked.length>1)throw fail('Несколько блоков обложки. Нельзя однозначно выбрать изображение.');
    if(marked.length===1)return marked[0].querySelector('img')?.getAttribute('src')||'';
    // Tutor LMS may replace the preview block and omit the button temporarily.
    const withImages=all.filter(e=>e.querySelector('img[src]'));
    if(withImages.length===1)return withImages[0].querySelector('img').getAttribute('src')||'';
    if(preferredKey){
      const matches=withImages.filter(e=>courseImageStem(e.querySelector('img').getAttribute('src'))===String(preferredKey).toLowerCase());
      if(matches.length===1)return matches[0].querySelector('img').getAttribute('src')||'';
    }
    return '';
  }
  function imageMatch(item,key){
    let filename='';try{filename=imageKey(item.querySelector('img')?.src).split('/').pop().replace(/\.[^.]+$/,'');}catch{}
    return (item.getAttribute('aria-label')||'').trim()===key||filename===key;
  }
  async function writeImage(key,r){
    guard();const old=one('.media-modal',document,false);
    if(old){one('.media-modal-close',old).click();await waitFor(()=>!one('.media-modal',document,false),'закрытие старого окна медиатеки');}
    const buttons=[...document.querySelectorAll('[data-cy="media-preview"] [data-cy="replace-media"]')];
    if(buttons.length!==1)throw fail('Не найдена единственная кнопка Replace Image. Update не нажат.');
    log('info',`Меняем картинку: поиск img=${key}`,r);guard();buttons[0].click();
    const modal=await waitFor(()=>one('.media-modal',document,false),'окно медиатеки');
    let input=one('#media-search-input',modal,false);
    if(!input){const tabs=[...modal.querySelectorAll('.media-router button,.media-router a')].filter(e=>visible(e)&&/^(Media Library|Медиафайлы|Библиотека файлов)$/i.test(e.textContent.trim()));if(tabs.length===1)tabs[0].click();input=await waitFor(()=>one('#media-search-input',modal,false),'поиск медиатеки');}
    input.focus();nativeValue(input,key);input.dispatchEvent(new unsafeWindow.KeyboardEvent('keyup',{bubbles:true,key:'Enter'}));
    await sleep(1200);guard();let stable='',since=0;
    const item=await waitFor(()=>{
      if(one('#media-search-input',modal,false)?.value!==key)return null;
      if([...modal.querySelectorAll('.spinner.is-active,.media-progress-bar')].some(visible)){since=0;return null;}
      const items=[...modal.querySelectorAll('.attachments .attachment')].filter(visible),matches=items.filter(e=>imageMatch(e,key));
      if(matches.length>1)throw fail(`Найдено несколько картинок img=${key}. Update не нажат.`);
      if(items.length!==1||matches.length!==1){since=0;return null;}
      const signature=items[0].getAttribute('data-id')+'|'+items[0].querySelector('img')?.src;
      if(signature!==stable||!since){stable=signature;since=Date.now();return null;}
      return Date.now()-since>=800?items[0]:null;
    },`единственная картинка img=${key}: не найдена или поиск неоднозначен`,30000);
    guard();if(item.getAttribute('aria-checked')!=='true'){const preview=item.querySelector('.attachment-preview');if(!preview)throw fail('Не найдено превью картинки.');preview.click();}
    await waitFor(()=>item.getAttribute('aria-checked')==='true','выбор картинки');
    if(modal.querySelectorAll('.attachments .attachment[aria-checked="true"]').length!==1)throw fail('Выбрано несколько файлов. Update не нажат.');
    const expected=imageKey(item.querySelector('img')?.src);if(!expected)throw fail('Не удалось прочитать адрес картинки.');
    const select=one('.media-button-select',modal);if(select.disabled)throw fail('Select недоступна.');guard();select.click();
    await waitFor(()=>!visible(modal),'закрытие медиатеки');
    let previewSrc='';
    try{
      previewSrc=await waitFor(()=>{
        const src=imagePreview(key);
        return sameCourseImage(src,expected)?src:null;
      },'новая обложка',30000);
    }catch(e){
      if(!/Время ожидания истекло/.test(e.message||''))throw e;
      const previews=[...document.querySelectorAll('[data-cy="media-preview"] img')]
        .map(img=>img.getAttribute('src')||'').slice(0,5);
      throw fail('Выбор фото выполнен, но новая обложка не подтверждена. img='+key+'; ожидалось: '+expected+'; превью: '+JSON.stringify(previews)+'. Update не нажат.');
    }
    if(imageKey(previewSrc)!==expected)
      log('info','Превью использует другой размер, формат или CDN: '+previewSrc+'. Имя изображения совпадает с img='+key+'.',r);
    log('info','Картинка '+key+' подтверждена по имени файла. Заполняем остальные поля.',r);
    return expected;
  }
  async function ready(allowStop=false){
    await waitFor(()=>one('input[name="post_title"]',document,false)&&one('input[name="course_price"]',document,false)&&one('input[name="course_sale_price"]',document,false)&&one('button[data-cy="course-builder-submit-button"]',document,false),'поля Tutor LMS',45000,allowStop);
    await sleep(700);guard(allowStop);
  }
  function currentCourse(){try{return courseURL(location.href).id;}catch{return '';}}
  async function snapshot(allowStop=false){return {title:one('input[name="post_title"]').value,price:one('input[name="course_price"]').value,sale:one('input[name="course_sale_price"]').value,html:await readDescription(allowStop),slug:linkSlug(),url:publicLink()?.href||'',image:imageKey(imagePreview())};}
  function differences(actual,expected,persisted=false){const errors=[];
    if(expected.image!==undefined&&!sameCourseImage(actual.image,expected.image))errors.push('Изображение курса');
    if(actual.title!==expected.title)errors.push('Title');
    const sale=String(actual.sale??'').trim();
    if(sale!==''&&!(persisted&&state.settings.acceptZeroSale&&/^0+(?:[.,]0+)?$/.test(sale)))errors.push(`Sale Price не пустое (значение: ${JSON.stringify(sale)})`);
    try{if(price(actual.price)!==price(expected.price))errors.push('Regular Price');}catch{errors.push('Regular Price');}
    if(htmlSignature(actual.html)!==htmlSignature(expected.html))errors.push('Description / HTML');
    if(actual.slug!==expected.slug)errors.push('slug');return errors;
  }
  function navigate(url){
    guard(true);
    const r=state.rows[state.index];
    if(r?.status==='verifying')r.verifyFromDocument=DOCUMENT_ID;
    persist();
    sessionStorage.setItem(TICKET,JSON.stringify({tab:tabId,at:Date.now(),url,run:state.run}));
    leaving=true;
    // If a beforeunload prompt or navigation failure prevents leaving, unlock the UI.
    let departed=false;
    window.addEventListener('pagehide',()=>{departed=true;},{once:true});
    setTimeout(()=>{
      if(departed||!leaving||!owned())return;
      leaving=false;busy=false;halted=true;state.running=false;sessionStorage.removeItem(TICKET);
      log('error','Переход не завершился за 20 секунд. Если браузер спрашивает о выходе со страницы, подтвердите переход. Иначе обновите страницу вручную и нажмите «Продолжить».',r);
      release();render();
    },20000);
    if(new URL(url,location.href).href===location.href)location.reload();
    else location.assign(url);
  }
  function release(){clearInterval(leaseTimer);if(owned())GM_setValue(LEASE,null);}
  async function acquire(){
    const old=GM_getValue(LEASE,null);
    if(old&&old.expires>Date.now()&&old.tab!==tabId)throw fail('Очередь работает в другой вкладке. Остановите её там или подождите 20 секунд после закрытия.');
    GM_setValue(LEASE,{tab:tabId,expires:Date.now()+20000});await sleep(180);
    if(!owned())throw fail('Не удалось получить управление очередью.');
    leaseTimer=setInterval(()=>{if(!owned()){halted=true;clearInterval(leaseTimer);return;}GM_setValue(LEASE,{tab:tabId,expires:Date.now()+20000});},4000);
  }
  async function betweenCourses(){
    const end=Date.now()+state.settings.delay*1000;
    while(Date.now()<end){guard();await sleep(Math.min(200,end-Date.now()));}
    guard();
  }
  function stop(asPause=false){
    state.paused=asPause;
    if(state.owner!==tabId || !busy){state.running=false;halted=true;persist();release();render();return;}
    halted=true;state.running=false;persist();render();
    log('info',`${asPause?'Пауза':'Стоп'} запрошена. Новые действия прекращаются; уже отправленный Update будет проверен. Следующий курс не откроется.`);
  }
  async function run(){
    if(busy)return;busy=true;leaving=false;
    try{
      while(true){
        guard();
        let i=state.index;
        if(i<0||!state.rows[i]||!ACTIVE.has(state.rows[i].status))i=state.rows.findIndex(r=>r.status==='pending');
        if(i<0){state.running=false;state.phase='idle';state.index=-1;persist();log('info','Очередь завершена. Можно скачать отчёт.');break;}
        state.index=i;const r=state.rows[i];
        try{
          if(currentCourse()!==r.id){stage(r,'opening','Переход к редактору');navigate(r.url);return;}
          log('info','Ожидаем поля редактора (до 45 секунд)',r);
          await ready(state.phase==='verifying'||state.phase==='saving');
          if(r.expected && ['saving','verifying'].includes(r.status)){
            stage(r,'verifying','Читаем значения после перезагрузки');
            const actual=await snapshot(true);
            if(r.slug && actual.slug.startsWith(r.expected.slug+'-') && /^\d+$/.test(actual.slug.slice(r.expected.slug.length+1))){
              log('warn',`WordPress добавил суффикс к slug: ${actual.slug}. Фактический адрес сохранён в отчёте.`,r);
              r.expected.slug=actual.slug;
            }
            const mismatch=differences(actual,r.expected,true);
            r.actualSale=actual.sale;r.actualImage=actual.image;
            if(String(actual.sale).trim()!==''&&state.settings.acceptZeroSale&&/^0+(?:[.,]0+)?$/.test(String(actual.sale).trim()))log('warn',`Сайт вернул Sale Price = ${actual.sale}. Ноль принят по включённой настройке; перед Update поле очищалось.`,r);
            r.actualURL=actual.url;r.checked=new Date().toISOString();
            if(mismatch.length){r.status='review';r.message=`После загрузки не совпали: ${mismatch.join(', ')}. Повторное сохранение не выполнялось.`;log('error',r.message,r);}
            else{r.status='done';r.message='Сохранение подтверждено повторной загрузкой';log('ok',r.message,r);}
            state.phase='idle';state.index=-1;persist();
            if(single||halted||(r.status==='review'&&state.settings.stopOnError)){state.running=false;persist();break;}
            await betweenCourses();continue;
          }
          stage(r,'filling','Проверка курса и заполнение полей');
          r.before=await snapshot();persist();
          let selectedImage;if(r.img)selectedImage=await writeImage(r.img,r);
          if(r.title)await field('input[name="post_title"]',r.title);
          if(r.slug)await writeSlug(r.slug);
          await field('input[name="course_sale_price"]','');
          if(r.price!==null)await field('input[name="course_price"]',r.price);
          // Description is last: price/title updates can re-render a controlled editor.
          await sleep(500);guard();
          let normalized=r.before.html;
          if(r.description){log('info','Цены заполнены. Синхронизируем Description: Text → Visual',r);normalized=await writeDescription(r.description);}
          guard();
          r.expected={title:r.title||r.before.title,slug:r.slug||r.before.slug,html:normalized,price:r.price??r.before.price,sale:'',...(r.img?{image:selectedImage}:{})};
          const filled=await snapshot(), mismatch=differences(filled,r.expected);
          if(mismatch.length)throw fail(`Проверка заполнения: ${mismatch.join(', ')}. ${mismatch.includes('Description / HTML')?descriptionMismatch(r.expected.html,filled.html,'Повторное чтение Description'):''}`);
          guard();
          if(!r.description){
            const ta=editorRef();
            if(visible(ta)){
              const vb=editorButton(ta.id,'tmce');if(!vb)throw fail('Не найдена кнопка Visual.');vb.click();
              const ed=await waitFor(()=>{const e=tiny(ta.id);return e&&!e.isHidden()&&e.initialized?e:null;},'Visual перед Update',20000);
              if(htmlSignature(ed.getContent())!==htmlSignature(r.expected.html))throw fail('Visual изменил существующее описание. Update не нажат.');
            }
          }
          guard();
          // snapshot does not switch a working Visual editor back to Text.
          const button=one('button[data-cy="course-builder-submit-button"]');
          if(button.disabled||button.getAttribute('aria-disabled')==='true')throw fail('Кнопка Update недоступна.');
          stage(r,'saving','Отправляем Update');r.saveAt=Date.now();r.saveDocument=DOCUMENT_ID;persist();
          button.click();
          // Wait for busy state to finish, or use a conservative delay when the site exposes none.
          let sawBusy=false, idleSince=0;const started=Date.now();
          while(Date.now()-started<45000){
            guard(true);const b=one('button[data-cy="course-builder-submit-button"]',document,false);
            const working=!b||b.disabled||b.getAttribute('aria-busy')==='true'||/updating|saving|сохранение/i.test(b.textContent);
            if(working){sawBusy=true;idleSince=0;}else if(!idleSince)idleSince=Date.now();
            if(sawBusy&&idleSince&&Date.now()-idleSince>1500&&Date.now()-started>3000)break;
            if(!sawBusy&&Date.now()-started>12000)break;
            await sleep(300);
          }
          if(sawBusy&&!idleSince)throw fail('Update не завершился за 45 секунд. Результат требует проверки.');
          stage(r,'verifying','Перезагрузка для проверки сохранения');
          state.stopAfterVerify=halted||single;persist();
          navigate(r.url);return;
        }catch(e){
          if(e===STOP)throw e;
          r.status=r.expected&&r.saveAt?'review':'error';r.message=e.message||String(e);state.phase='idle';state.index=-1;persist();log('error',r.message,r);
          if(e.fatal||halted||single||state.settings.stopOnError){state.running=false;persist();break;}
          await betweenCourses();
        }
      }
    }catch(e){if(e!==STOP)log('error',e.message||String(e));state.running=false;persist();}
    finally{if(!leaving){busy=false;state.running=false;persist();release();render();}}
  }
  async function start(oneOnly=false,automatic=false){
    if(tBusy){message('Очередь преподавателей уже работает. Остановите её перед запуском курсов.');return;}
    if(busy){message('Уже выполняется: '+(state.rows[state.index]?.message||state.phase)+'. Для остановки нажмите «Стоп».');return;}
    state=GM_getValue(KEY,empty());
    try{
      authCheck();await acquire();halted=false;single=oneOnly;
      if(!automatic)state.paused=false;
      state.owner=tabId;state.run=automatic?state.run:uid();state.running=true;
      if(state.stopAfterVerify){single=true;state.stopAfterVerify=false;}
      state.single=single;
      // A resumed interrupted save is verified on a fresh page before any new mutation.
      const r=state.rows[state.index];
      if(!automatic&&r&&['saving','verifying'].includes(r.status)&&r.expected){
        const previousDocument=r.verifyFromDocument||r.saveDocument;
        state.phase='verifying';r.status='verifying';
        if(!previousDocument||previousDocument===DOCUMENT_ID||currentCourse()!==r.id){
          busy=true;log('info','Возобновление: загружаем курс заново для проверки, без повторного Update',r);
          persist();navigate(r.url);return;
        }
        log('info','Возобновление на новой странице: проверяем сохранённые поля',r);
      }
      state.single=single;persist();render();await run();
    }catch(e){state.running=false;persist();release();log('error',e.message||String(e));}
  }

  const css=`
    :host{all:initial;font:13px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:#182438;color-scheme:light}
    *{box-sizing:border-box}button,input,select{font:inherit}button{cursor:pointer;border:1px solid #d9e1eb;border-radius:9px;padding:8px 12px;background:#fff;color:#26354c;font-weight:600}button:hover{background:#f0f4fa}button:disabled{opacity:.42;cursor:default}button:focus-visible,input:focus,select:focus{outline:2px solid #4f7df3;outline-offset:2px}
    .panel{position:fixed;right:22px;bottom:20px;width:710px;height:650px;max-width:calc(100vw - 24px);max-height:calc(100vh - 40px);background:#fff;box-shadow:0 18px 70px #0f233b40;border:1px solid #dae2ee;border-radius:16px;display:flex;flex-direction:column;overflow:hidden;resize:both}
    .panel.full{inset:24px;width:auto;height:auto;max-width:none;max-height:none;resize:none}.head{background:#14263e;color:white;padding:16px 18px;display:flex;align-items:center;gap:12px;cursor:move;touch-action:none}.logo{height:36px;width:36px;background:#4275ed;display:grid;place-items:center;border-radius:10px;font-size:18px;font-weight:800}.head b{font-size:16px;letter-spacing:.2px}.head small{color:#b1c2d9;display:block;font-size:11px}.head .spacer{flex:1}.icon{background:#ffffff13;color:#fff;border-color:#ffffff22;padding:4px 9px}
    .topmodes{display:flex;flex-shrink:0;padding:0 16px;background:#f5f7fb;border-bottom:1px solid #dce5f2;gap:10px}.topmodes button{flex:1;border:none;border-bottom:3px solid transparent;border-radius:0;background:transparent;padding:12px 8px;color:#64758e}.topmodes button.on{color:#2458ce;border-bottom-color:#3269e8;background:#eaf1ff}.body{display:flex;flex-direction:column;min-height:0;flex:1}.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;padding:14px 16px 10px}.card{border:1px solid #e4eaf2;border-radius:10px;padding:9px 12px;background:#fafbfd}.card b{display:block;font-size:21px;line-height:1.25}.card span{font-size:11px;color:#68778c}.card.done b{color:#16835e}.card.bad b{color:#bd4f39}
    .progress{height:5px;background:#eaf0f8;margin:0 16px;border-radius:5px;overflow:hidden}.progress i{height:100%;display:block;background:#4076ed;transition:width .3s}.current{padding:10px 16px;display:flex;gap:8px;align-items:center;color:#586b83;min-height:42px}.dot{height:8px;width:8px;background:#93a1b3;border-radius:50%;flex:none}.dot.run{background:#3974ee;box-shadow:0 0 0 4px #e5edff}.current span:last-child{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .actions{display:flex;gap:7px;flex-wrap:wrap;padding:0 16px 12px}.primary{background:#3468df!important;border-color:#3468df!important;color:white}.danger{color:#b04439;border-color:#eabcb6;background:#fff7f6}.tabs{display:flex;overflow-x:auto;flex-shrink:0;padding:0 16px;border-bottom:1px solid #e4eaf2;gap:18px}.tabs button{flex-shrink:0;border:0;border-radius:0;padding:10px 0;color:#7a889a;background:none;font-size:12px}.tabs button.on{color:#265cce;border-bottom:2px solid #3468df}
    .view{padding:12px 16px;min-height:0;flex:1;overflow:auto}.hidden{display:none!important}.tools{display:flex;gap:8px;margin-bottom:10px}input,select{border:1px solid #dce3ed;border-radius:8px;background:#fff;padding:8px;color:#22324a;min-width:0}input[type=search]{flex:1}table{border-collapse:collapse;width:100%;table-layout:fixed;font-size:12px}th{text-align:left;color:#7c899b;font-size:10px;text-transform:uppercase;letter-spacing:.5px;padding:8px 6px;background:#f7f9fc;position:sticky;top:0}td{padding:10px 6px;border-bottom:1px solid #edf1f6;vertical-align:top}tr.currentrow{background:#eff5ff}td .title{font-weight:600;line-height:1.4;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}td small{color:#8290a2;display:block;font-size:10px}td button{padding:3px 7px;font-size:11px}.badge{font-size:10px;padding:3px 7px;border-radius:6px;display:inline-block;white-space:nowrap;background:#eef2f7;color:#62738a}.badge.done{background:#e4f5ec;color:#147c55}.badge.error,.badge.review{background:#fff0e8;color:#b24b28}.badge.filling,.badge.opening,.badge.saving,.badge.verifying{background:#e7efff;color:#2b62ce}
    .foot{padding:10px 16px;border-top:1px solid #e4eaf2;display:flex;align-items:center;gap:8px;background:#fafbfd;font-size:11px;color:#7a889a}.foot span{flex:1}.foot button{padding:4px 8px;font-size:11px}.empty{text-align:center;padding:28px 16px;color:#77869b}.empty b{display:block;color:#31435c;font-size:16px;margin-bottom:6px}.drop{border:2px dashed #c4d2e9;border-radius:12px;text-align:center;padding:20px;margin:4px 0 12px;background:#f7faff}.drop.over{border-color:#3468df;background:#eaf1ff}.drop label{cursor:pointer;color:#2b63d4;font-weight:700}.drop small{display:block;margin-top:5px;color:#7b8a9e}.mapping{display:grid;grid-template-columns:120px 1fr;gap:8px;align-items:center}.mapping label{color:#61728a}.hint{font-size:12px;color:#6d7e94;margin:10px 0}.notice{padding:10px 12px;background:#fff7e8;border:1px solid #f0d8aa;border-radius:9px;color:#835e24;font-size:12px;margin-bottom:10px}.setting{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid #edf1f6}.setting input[type=number]{width:80px}.setting div{flex:1}.setting small{display:block;color:#7a899c}.logline{display:grid;grid-template-columns:65px 50px 1fr;gap:8px;padding:8px 0;border-bottom:1px solid #edf1f6;font-size:11px;overflow-wrap:anywhere}.logline time{color:#8a97a8}.logline.error{color:#a83f2f}.logline.ok{color:#147c55}.help h3{font-size:13px;color:#263d5c}.help p,.help li{color:#61728a;font-size:12px}.help ol{padding-left:20px}.pill{position:fixed;right:22px;bottom:20px;border-radius:14px;background:#14263e;color:#fff;padding:12px 18px;box-shadow:0 8px 30px #0f233b40}button.pill:hover,button.pill:focus-visible{background:#24466b;color:#fff;border-color:#5680ae}.detail{position:absolute;inset:70px 12px 12px;background:white;z-index:4;border:1px solid #ccd8e9;box-shadow:0 10px 40px #14263e33;border-radius:12px;padding:16px;overflow:auto}.detail header{display:flex;justify-content:space-between;gap:10px}.detail pre{white-space:pre-wrap;overflow-wrap:anywhere;background:#f5f7fb;padding:10px;border-radius:8px;font:11px/1.6 monospace}.detail a{color:#2965dc;overflow-wrap:anywhere}.errorbox{position:absolute;bottom:54px;left:16px;right:16px;z-index:8;padding:12px;background:#fff0ee;border:1px solid #e8b9af;border-radius:10px;color:#a13e2e;box-shadow:0 6px 20px #0002}.errorbox button{float:right;padding:0 6px}.mono{font-family:monospace}
    @media(max-width:550px){.panel{right:8px;bottom:8px;max-width:calc(100vw - 16px)}.actions button{font-size:11px;padding:7px}.stats{gap:4px}.card{padding:7px}.tabs{gap:12px}}
  `;
  function mount(){
    const host=document.createElement('div');host.id='yc-course-manager';host.style.cssText='position:fixed;z-index:2147483646;inset:0;pointer-events:none';document.documentElement.append(host);
    root=host.attachShadow({mode:'open'});
    root.innerHTML=`<style>${css}:host>section,:host>button{pointer-events:auto}</style><button class="pill hidden" id="pill">Курсы · открыть панель</button><section class="panel" id="panel">
      <header class="head" id="drag"><div class="logo">C</div><div><b>Course Manager</b><small>${esc(location.hostname)} · TUTOR LMS · v1.2.3</small></div><span class="spacer"></span><button class="icon" id="full" title="На весь экран">⛶</button><button class="icon" id="min" title="Свернуть">—</button></header>
      <div class="body"><div class="stats"><div class="card"><b id="total">0</b><span>Курсов в файле</span></div><div class="card done"><b id="done">0</b><span>Проверено и готово</span></div><div class="card"><b id="pending">0</b><span>В очереди / работе</span></div><div class="card bad"><b id="errors">0</b><span>Ошибки / проверка</span></div></div>
      <div class="progress"><i id="bar"></i></div><div class="current"><i class="dot" id="dot"></i><span id="current">Загрузите файл с курсами</span></div>
      <div class="actions"><button class="primary" id="start">▶ Начать</button><button id="one">Один курс</button><button id="pause">Ⅱ Пауза</button><button class="danger" id="stop">■ Стоп</button><button id="retry">Повторить ошибки</button><button id="report">↓ Отчёт</button></div>
      <nav class="tabs"><button data-tab="queue" class="on">Очередь</button><button data-tab="processed">Обработанные</button><button data-tab="import">Импорт Excel</button><button data-tab="log">Журнал</button><button data-tab="settings">Настройки</button><button data-tab="help">Как работать</button></nav>
      <div class="view" id="queue"><div class="tools"><input type="search" id="search" placeholder="Название или ID курса…"><select id="filter"><option value="all">Все статусы</option><option value="pending">В очереди</option><option value="done">Выполнено</option><option value="error">Ошибки и проверка</option><option value="skipped">Пропущено</option></select></div><div id="rows"></div></div>
      <div class="view hidden" id="import"><div class="tools"><button class="danger" id="resetAll">Очистить форму и всю историю</button></div><div class="drop" id="drop"><label for="file">Выбрать Excel-файл</label><input id="file" type="file" accept=".xlsx,.xls,.csv" hidden><small>Или перетащите сюда · .xlsx / .xls / .csv</small><small id="filename">Данные читаются локально в браузере</small></div><div id="mapping" class="hidden"><div class="mapping"><label>Лист</label><select id="sheet"></select><label>Название</label><select id="mapTitle"></select><label>Ссылка редактора</label><select id="mapURL"></select><label>HTML-описание</label><select id="mapDescription"></select><label>Обычная цена</label><select id="mapPrice"></select><label>Картинка img (необязательно)</label><select id="mapImg"></select></div><p class="hint">Ссылка должна вести в wp-admin с course_id. Цены 3.42€, €3,42 или 3.42 EUR станут 3.42; в WordPress передаётся только число. Пустые значения не заменяют существующие поля. img необязателен: цифры в текстовом формате Excel. Картинка заменяется первой. После импорта курсы автоматически перемешиваются; исходные номера строк Excel сохраняются.</p><button class="primary" id="load">Загрузить в очередь</button></div><p class="hint">Новая загрузка заменяет текущую очередь. Перед этим можно скачать отчёт. Заголовки должны находиться в первой строке.</p></div>
      <div class="view hidden" id="log"><div class="tools"><button id="downloadLog">↓ Скачать журнал</button><span class="hint">Последние 1500 событий</span></div><div id="logs"></div></div>
      <div class="view hidden" id="settings"><label class="setting"><div>Пауза между курсами<small>Секунды после проверки сохранения</small></div><input id="delay" type="number" min="1" max="60" value="3"></label><label class="setting"><input id="stopOnError" type="checkbox"><div>Останавливаться при любой ошибке<small>Потеря авторизации всегда останавливает очередь</small></div></label><label class="setting"><input id="acceptZeroSale" type="checkbox"><div>При проверке принимать Sale Price = 0<small>Только если сайт возвращает ноль вместо пустого поля. Перед Update поле всё равно очищается.</small></div></label><div class="setting"><div>Slug из Title<small>Генерируется автоматически, если Title заполнен</small></div><span class="badge done">Включено</span></div><div class="setting"><div>Очищать Sale Price<small>Для каждого обрабатываемого курса, включая значение 0</small></div><span class="badge done">Всегда</span></div><p class="hint">После Update страница перезагружается. «Выполнено» появляется только после сверки сохранённых полей. По умолчанию требуется пустое Sale Price. При необходимости разрешите нулевое значение переключателем выше. Ненулевая скидка всегда считается расхождением.</p><button id="backup">↓ Резервная копия JSON</button></div>
      <div class="view help hidden" id="help"><h3>Курсы — как работать</h3><ol><li>Откройте админ-панель сайта <strong>${esc(location.hostname)}</strong> и войдите с правами на редактирование курсов.</li><li>На вкладке «Курсы» откройте «Импорт Excel», выберите файл и лист. Сопоставьте Title, Course URL, Description, Regular Price; img необязателен.</li><li>Убедитесь, что Course URL каждой строки указывает на <strong>${esc(location.hostname)}</strong> и содержит <code>page=create-course&amp;course_id=…</code>. Чужие домены отклоняются.</li><li>Нажмите «Загрузить в очередь»: курсы перемешаются случайным образом, но номер исходной строки Excel сохранится в журнале и отчёте. Затем нажмите «Один курс» для проверки первой случайной записи. Если результат правильный, нажмите «Начать» для остальных.</li><li>Скрипт обновит заголовок, slug, цену, HTML и, если указан img, обложку; Sale Price очищается. После Update он перезагрузит редактор и сверит сохранение.</li><li>В конце скачайте «Отчёт»; подробности ошибок доступны через «⋯» и «Журнал».</li></ol><h3>Пауза и ошибки</h3><p>Используйте «Пауза» или «Стоп» для прекращения новых действий. Уже отправленное сохранение может завершиться; после возобновления оно сначала проверяется. Не редактируйте запись вручную во время автоматизации. При ошибках используйте «Повторить ошибки» только после проверки причины.</p><h3>Описание и изображения</h3><p>Description содержит HTML. Изображение img — числовое имя файла в медиатеке (не WordPress attachment ID). Пустые значения не перезаписывают существующие поля, кроме Sale Price. Изменение slug само по себе не создаёт редирект со старого адреса.</p><h3>Несколько сайтов</h3><p>На каждом домене своя очередь, журнал и настройки. Для работы с другим сайтом сначала откройте его wp-admin и импортируйте отдельный Excel с адресами именно этого сайта. Данные и история ${esc(location.hostname)} не будут смешаны с другими сайтами.</p></div>
      <footer class="foot"><span id="footer">Очередь сохраняется автоматически</span><button id="prev">←</button><span id="pages" style="flex:none">1 / 1</span><button id="next">→</button></footer></div></section>`;
    panel=root.getElementById('panel');
    tMount();
    const on=(id,event,fn)=>root.getElementById(id).addEventListener(event,fn);
    on('start','click',()=>start(false));on('one','click',()=>start(true));on('stop','click',()=>stop(false));on('pause','click',()=>stop(true));on('resetAll','click',resetAll);
    on('min','click',()=>{panel.classList.add('hidden');root.getElementById('pill').classList.remove('hidden');});
    const show=()=>{panel.classList.remove('hidden');root.getElementById('pill').classList.add('hidden');};
    on('pill','click',show);GM_registerMenuCommand('Открыть Course Manager',show);
    on('full','click',()=>{expanded=!expanded;panel.classList.toggle('full',expanded);panel.style.left='';panel.style.top='';});
    root.querySelectorAll('[data-tab]').forEach(b=>b.addEventListener('click',()=>tab(b.dataset.tab)));
    on('search','input',e=>{search=e.target.value;page=0;renderRows();});on('filter','change',e=>{filter=e.target.value;page=0;renderRows();});
    on('prev','click',()=>{page=Math.max(0,page-1);renderRows();});on('next','click',()=>{page++;renderRows();});
    on('file','change',e=>readFile(e.target.files[0]));on('sheet','change',mapping);
    const drop=root.getElementById('drop');drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('over');});drop.addEventListener('dragleave',()=>drop.classList.remove('over'));drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('over');readFile(e.dataTransfer.files[0]);});
    on('load','click',loadQueue);on('report','click',report);on('downloadLog','click',downloadLog);
    on('backup','click',()=>download(JSON.stringify(state,null,2),EXPORT_PREFIX+'course-manager-backup.json','application/json'));
    on('retry','click',()=>{if(!editable())return;let n=0;state.rows.forEach(r=>{if(r.status==='error'&&!r.invalid){r.status='pending';r.message='Повторная попытка';delete r.expected;delete r.saveAt;n++;}});persist();log('info',`Возвращено в очередь: ${n}`);});
    on('delay','change',e=>{if(!editable())return;state.settings.delay=Math.min(60,Math.max(1,Number(e.target.value)||3));e.target.value=state.settings.delay;persist();});
    on('acceptZeroSale','change',e=>{if(!editable())return;state.settings.acceptZeroSale=e.target.checked;persist();});
    on('stopOnError','change',e=>{if(!editable())return;state.settings.stopOnError=e.target.checked;persist();});
    let drag=null;
    on('drag','pointerdown',e=>{if(e.target.closest('button')||expanded)return;const r=panel.getBoundingClientRect();drag={x:e.clientX-r.left,y:e.clientY-r.top};e.currentTarget.setPointerCapture(e.pointerId);});
    on('drag','pointermove',e=>{if(!drag)return;panel.style.left=Math.max(0,Math.min(innerWidth-panel.offsetWidth,e.clientX-drag.x))+'px';panel.style.top=Math.max(0,Math.min(innerHeight-panel.offsetHeight,e.clientY-drag.y))+'px';panel.style.right='auto';panel.style.bottom='auto';});on('drag','pointerup',()=>drag=null);
    render();if(!state.rows.length)tab('import');
  }
  function tab(name){activeTab=name;page=0;filter='all';root.getElementById('filter').value='all';for(const n of ['queue','import','log','settings','help'])root.getElementById(n).classList.toggle('hidden',n==='queue'?!['queue','processed'].includes(name):n!==name);root.querySelectorAll('[data-tab]').forEach(b=>b.classList.toggle('on',b.dataset.tab===name));render();}
  function message(s){root.querySelector('.errorbox')?.remove();const e=document.createElement('div');e.className='errorbox';const close=document.createElement('button');close.textContent='×';close.onclick=()=>e.remove();e.append(close,document.createTextNode(s));panel.append(e);}
  function editable(){const l=GM_getValue(LEASE,null);if(busy||(l&&l.expires>Date.now())){message('Сначала остановите обработку и дождитесь завершения текущего действия.');return false;}state=GM_getValue(KEY,empty());return true;}
  function resetAll(){
    if(!editable())return;
    if(!confirm('Полностью очистить приложение? Будут удалены загруженный файл, очередь, обработанные курсы, история, журнал и настройки. Изменения на сайте останутся.'))return;
    halted=true;busy=false;leaving=false;single=false;release();
    sessionStorage.removeItem(TICKET);GM_setValue(LEASE,null);
    state=empty();GM_setValue(KEY,state);
    book=null;importName='';page=0;filter='all';search='';
    root.getElementById('file').value='';root.getElementById('filename').textContent='Данные читаются локально в браузере';
    root.getElementById('sheet').innerHTML='';
    for(const name of ['Title','URL','Description','Price','Img'])root.getElementById('map'+name).innerHTML='';
    root.getElementById('mapping').classList.add('hidden');
    root.getElementById('search').value='';root.getElementById('filter').value='all';
    root.querySelector('.detail')?.remove();root.querySelector('.errorbox')?.remove();
    tab('import');renderRows();renderLogs();
  }
  function render(){if(!root)return;const count=s=>state.rows.filter(r=>s.includes(r.status)).length;
    root.getElementById('total').textContent=state.rows.length;root.getElementById('done').textContent=count(['done']);root.getElementById('pending').textContent=count(['pending',...ACTIVE]);root.getElementById('errors').textContent=count(['error','review']);
    root.getElementById('bar').style.width=(state.rows.length?100*count(['done','skipped'])/state.rows.length:0)+'%';
    const r=state.rows[state.index];root.getElementById('current').textContent=busy?(halted?'Останавливаемся после текущего действия…':`${LABEL[r?.status]||'Запуск'} · ${r?.title||r?.id||'очередь'}`):(state.paused?'На паузе · нажмите «Продолжить»':r&&ACTIVE.has(r.status)?'Приостановлено · продолжение с сохранённого этапа':state.rows.length?`${state.filename} · готово к запуску`:'Загрузите файл с курсами');
    root.getElementById('dot').classList.toggle('run',busy);root.getElementById('start').textContent=(state.paused||state.rows.some(r=>r.status!=='pending'&&!r.invalid))?'▶ Продолжить':'▶ Начать';
    const can=state.rows.some(r=>r.status==='pending'||ACTIVE.has(r.status));root.getElementById('start').disabled=busy||!can;root.getElementById('one').disabled=busy||!can;root.getElementById('stop').disabled=!busy;root.getElementById('pause').disabled=!busy||halted;root.getElementById('resetAll').disabled=busy;root.getElementById('acceptZeroSale').disabled=busy;root.getElementById('acceptZeroSale').checked=!!state.settings.acceptZeroSale;root.querySelector('[data-tab=processed]').textContent='Обработанные ('+count(['done','review'])+')';root.getElementById('retry').disabled=busy||!state.rows.some(r=>r.status==='error'&&!r.invalid);root.getElementById('load').disabled=busy;
    root.getElementById('delay').disabled=busy;root.getElementById('stopOnError').disabled=busy;root.getElementById('delay').value=state.settings.delay;root.getElementById('stopOnError').checked=state.settings.stopOnError;
    root.getElementById('footer').textContent=state.updated?'Сохранено локально · '+new Date(state.updated).toLocaleTimeString():'Очередь сохраняется автоматически';
    if(activeTab==='queue'||activeTab==='processed')renderRows();if(activeTab==='log')renderLogs();
  }
  function renderRows(){
    const rows=state.rows.filter(r=>(activeTab==='processed'?['done','review'].includes(r.status):!['done','review'].includes(r.status))&&(filter==='all'||(filter==='error'?['error','review'].includes(r.status):r.status===filter))&&`${r.title} ${r.id}`.toLowerCase().includes(search.toLowerCase()));
    const pages=Math.max(1,Math.ceil(rows.length/25));page=Math.min(page,pages-1);root.getElementById('pages').textContent=`${page+1} / ${pages}`;root.getElementById('prev').disabled=page===0;root.getElementById('next').disabled=page>=pages-1;
    const target=root.getElementById('rows');if(!rows.length){target.innerHTML='<div class="empty"><b>Здесь пока нет курсов</b>Импортируйте Excel или измените фильтр.</div>';return;}
    target.innerHTML=`<table><thead><tr><th style="width:45px">№</th><th>Курс</th><th style="width:65px">Цена</th><th style="width:105px">Статус</th><th style="width:35px"></th></tr></thead><tbody>${rows.slice(page*25,page*25+25).map(r=>`<tr class="${state.rows[state.index]===r?'currentrow':''}"><td>${r.excelRow}</td><td><span class="title" title="${esc(r.title)}">${esc(r.title||'Название без изменений')}</span><small>ID ${esc(r.id||'—')} · ${esc(r.message||r.slug||'')}</small></td><td>${esc(r.price??'—')}</td><td><span class="badge ${r.status}">${LABEL[r.status]}</span></td><td><button data-detail="${r.key}" title="Подробности">⋯</button></td></tr>`).join('')}</tbody></table>`;
    target.querySelectorAll('[data-detail]').forEach(b=>b.onclick=()=>detail(state.rows.find(r=>r.key===b.dataset.detail)));
  }
  function detail(r){root.querySelector('.detail')?.remove();const d=document.createElement('div');d.className='detail';d.innerHTML=`<header><b>Строка ${r.excelRow} · ID ${esc(r.id)}</b><button data-close>Закрыть</button></header><p><span class="badge ${r.status}">${LABEL[r.status]}</span> ${esc(r.message)}</p><p><strong>${esc(r.title||'Название не изменяется')}</strong></p><p>Цена: ${esc(r.price??'не изменяется')} · Sale Price: очистить</p><p>Новый slug: <span class="mono">${esc(r.slug||'не изменяется')}</span></p><p><a href="${esc(r.url||'#')}" target="_blank" rel="noopener noreferrer">Открыть редактор в другой вкладке</a></p>${r.actualURL?`<p>Фактический URL: ${esc(r.actualURL)}</p>`:''}<p>Картинка img: ${esc(r.img||'без изменений')}</p><h4>Description</h4><pre>${esc(r.description||'Не изменяется')}</pre><div class="actions"><button data-verify>Проверить сохранение повторно</button><button data-skip>Пропустить</button><button data-reset>Вернуть в очередь</button></div>`;
    d.querySelector('[data-close]').onclick=()=>d.remove();
    d.querySelector('[data-verify]').disabled=!r.expected||!r.saveAt;
    d.querySelector('[data-verify]').onclick=()=>{if(!editable())return;const row=state.rows.find(x=>x.key===r.key);row.status='verifying';row.verifyFromDocument=DOCUMENT_ID;state.index=state.rows.indexOf(row);state.phase='verifying';persist();d.remove();start(true);};
    d.querySelector('[data-skip]').onclick=()=>{if(!editable())return;const row=state.rows.find(x=>x.key===r.key);row.status='skipped';if(state.index===state.rows.indexOf(row)){state.index=-1;state.phase='idle';}persist();render();d.remove();};
    d.querySelector('[data-reset]').onclick=()=>{if(!editable())return;const row=state.rows.find(x=>x.key===r.key);if(row.invalid){message('Исправьте данные в Excel и импортируйте файл заново.');return;}if(['done','review','saving','verifying'].includes(row.status)&&!confirm('Повторно заполнить и сохранить этот курс? Для неопределённого сохранения сначала рекомендуется «Продолжить».'))return;row.status='pending';delete row.expected;delete row.saveAt;row.message='Возвращено вручную';persist();render();d.remove();};panel.append(d);
  }
  function renderLogs(){root.getElementById('logs').innerHTML=state.logs.slice().reverse().map(l=>`<div class="logline ${esc(l.level)}"><time>${new Date(l.time).toLocaleTimeString()}</time><span>${l.row?'№ '+l.row:'Система'}</span><span>${esc(l.message)}</span></div>`).join('')||'<div class="empty">События обработки появятся здесь</div>';}
  async function readFile(file){
    if(!file||!editable())return;
    try{if(typeof XLSX==='undefined')throw fail('Не загрузилась библиотека Excel. Проверьте @require в Tampermonkey.');if(file.size>25*1024*1024)throw fail('Файл больше 25 МБ. Разделите его на части.');book=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});importName=file.name;root.getElementById('filename').textContent=file.name;root.getElementById('sheet').innerHTML=book.SheetNames.map((n,i)=>`<option value="${i}">${esc(n)}</option>`).join('');root.getElementById('mapping').classList.remove('hidden');mapping();tab('import');}catch(e){message(e.message);}
  }
  function sheetData(){return XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[Number(root.getElementById('sheet').value)]],{header:1,defval:'',raw:false,blankrows:true});}
  function mapping(){const header=sheetData()[0]||[];const aliases={Title:['title'],URL:['courseurl','itempagelink'],Description:['description','info'],Price:['regularprice','price'],Img:['img']};for(const [name,keys] of Object.entries(aliases)){const s=root.getElementById('map'+name);s.innerHTML='<option value="">Выберите колонку…</option>'+header.map((h,i)=>`<option value="${i}">${esc(h||`Колонка ${i+1}`)}</option>`).join('');let idx=-1;for(const key of keys){idx=header.findIndex(h=>normalizeHeader(h)===key);if(idx>=0)break;}s.value=idx<0?'':String(idx);}}
  // Shuffle only new course imports. The persisted queue keeps its order after reload/resume.
  // excelRow remains unchanged, so logs and exports always refer to the original Excel row.
  function shuffleCourseQueue(rows){
    for(let i=rows.length-1;i>0;i--){
      const j=Math.floor(Math.random()*(i+1));
      [rows[i],rows[j]]=[rows[j],rows[i]];
    }
    return rows;
  }
  function loadQueue(){
    if(!book||!editable())return;
    try{
      const maps=['Title','URL','Description','Price'].map(n=>root.getElementById('map'+n).value);if(maps.some(v=>v==='')||new Set(maps).size!==4)throw fail('Выберите четыре разные колонки.');
      const imgMap=root.getElementById('mapImg').value;if(imgMap!==''&&maps.includes(imgMap))throw fail('Колонка img должна отличаться от остальных.');
      const [ti,ui,di,pi]=maps.map(Number), data=sheetData(), seen=new Set(), rows=[];
      for(let i=1;i<data.length;i++){
        const line=data[i];if(line.every(v=>!String(v).trim()))continue;
        const r={key:uid(),excelRow:i+1,title:String(line[ti]??'').trim(),sourceURL:String(line[ui]??'').trim(),description:String(line[di]??''),rawPrice:String(line[pi]??''),img:imgMap===''?'':String(line[Number(imgMap)]??'').trim(),status:'pending',message:'',price:null,id:'',url:'',slug:''};
        try{if(r.img&&!/^\d+$/.test(r.img))throw fail('img: только цифры без пробелов и научной записи. Задайте текстовый формат Excel.');const u=courseURL(r.sourceURL);r.id=u.id;r.url=u.url;if(seen.has(r.id))throw fail('Повторный course_id в файле.');seen.add(r.id);r.price=price(r.rawPrice);if(r.title){r.slug=slug(r.title);if(!r.slug)throw fail('Не удалось сформировать slug из Title.');if(r.slug.length>180)throw fail('Slug длиннее 180 символов. Сократите Title.');}if(r.description.trim())cleanHTML(r.description);else r.description='';}
        catch(e){r.status='error';r.invalid=true;r.message=e.message;}rows.push(r);
      }
      if(!rows.length)throw fail('На выбранном листе нет данных.');
      const slugs=new Map();for(const r of rows){if(r.slug&&!r.invalid){if(slugs.has(r.slug)){r.status='error';r.invalid=true;r.message=`Одинаковый slug со строкой ${slugs.get(r.slug)}. Измените Title.`;}else slugs.set(r.slug,r.excelRow);}}
      if(state.rows.length&&!confirm('Заменить текущую очередь новым файлом? Сначала скачайте отчёт, если он нужен.'))return;
      const settings=state.settings;shuffleCourseQueue(rows);state=empty();state.settings=settings;state.filename=importName;state.rows=rows;page=0;persist();tab('queue');log('info',`Импортировано ${rows.length} строк, очередь случайно перемешана. Номера строк Excel сохранены. Ошибок входных данных: ${rows.filter(r=>r.invalid).length}.`);
    }catch(e){message(e.message);}
  }
  function download(text,name,type){const url=URL.createObjectURL(new Blob([text],{type})),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
  function csv(rows){return '\uFEFF'+rows.map(r=>r.map(v=>{let s=String(v??'');if(/^[\s]*[=+@-]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';}).join(';')).join('\r\n');}
  function report(){download(csv([['Excel Row','Course ID','Title','Course URL','Description','Regular Price','Status','Message','Actual URL','Checked At','Original Title','Original Slug','Original Regular Price','Original Sale Price','Original Description','Saved Sale Price','img','Saved Image','Original Image','Excel Price (source)'],...state.rows.map(r=>[r.excelRow,r.id,r.title,r.sourceURL,r.description,r.price??'',LABEL[r.status],r.message,r.actualURL,r.checked,r.before?.title,r.before?.slug,r.before?.price,r.before?.sale,r.before?.html,r.actualSale,r.img,r.actualImage,r.before?.image,r.rawPrice])]),EXPORT_PREFIX+'course-manager-report.csv','text/csv;charset=utf-8');}
  function downloadLog(){download(csv([['Time','Level','Excel Row','Course ID','Stage','Message'],...state.logs.map(l=>[l.time,l.level,l.row,l.course,l.stage,l.message])]),EXPORT_PREFIX+'course-manager-log.csv','text/csv;charset=utf-8');}
  function captureErrors(){
    const seen=new Map();
    const record=s=>{if(!busy||!owned())return;const last=seen.get(s)||0;if(Date.now()-last<60000)return;seen.set(s,Date.now());log('warn',/Video loading failed/i.test(s)?`Предупреждение видеоплеера (не останавливает очередь): ${s}`:`Ошибка страницы: ${s}`,state.rows[state.index]);};
    window.addEventListener('error',e=>{if(e.message)record(e.message);else if(e.target?.src&&/codesample|tinymce/i.test(e.target.src))record('Не загружен '+e.target.src);},true);
    window.addEventListener('unhandledrejection',e=>record(String(e.reason?.message||e.reason||'Unhandled promise rejection')));
    let last='';const observer=new MutationObserver(()=>{if(!busy)return;const items=[...document.querySelectorAll('[role="alert"],.mce-notification,.tox-notification,.tutor-alert,.notice-error')].filter(visible);const text=items.map(e=>e.textContent.trim()).filter(Boolean).join(' | ');if(text&&text!==last){last=text;record(text);}});observer.observe(document.body,{subtree:true,childList:true});
  }
  // === Independent Tutor Instructor workflow; uses the same browser-tab lease as Courses. ===
  const TUTOR_KEY = 'yc.course.manager.tutors.v1'+SITE_SUFFIX;
  const TUTOR_TICKET = TUTOR_KEY + '.ticket';
  const MODE_KEY = 'yc.course.manager.selected-mode'+SITE_SUFFIX;
  const tEmpty = () => ({version:1,rows:[],logs:[],running:false,index:-1,owner:null,run:'',phase:'idle',paused:false,single:false,filename:'',settings:{delay:3,stopOnError:false},updated:''});
  let tState = GM_getValue(TUTOR_KEY,null) || tEmpty();
  let tBook=null, tImportName='', tTab='import', tPage=0, tFilter='all', tSearch='';
  let tBusy=false, tHalted=false, tLeaving=false, tSingle=false;
  let mainMode = GM_getValue(MODE_KEY, 'courses');

  function tURL(value){
    let u;
    try{u=new URL(String(value).trim());}catch{throw fail('Некорректный url_edit.');}
    const id=u.searchParams.get('user_id');
    if(u.origin!==ORIGIN || u.pathname!=='/wp-admin/user-edit.php' || !/^[1-9]\d*$/.test(id||''))
      throw fail('url_edit должен вести на текущий сайт '+location.hostname+'/wp-admin/user-edit.php?user_id=123; ссылки с другого сайта запрещены.');
    return {id,url:`${ORIGIN}/wp-admin/user-edit.php?user_id=${id}`};
  }
  function tCurrentID(){try{return tURL(location.href).id;}catch{return '';}}
  function tSplitName(s){
    const parts=String(s??'').trim().replace(/\s+/g,' ').split(' ').filter(Boolean);
    if(parts.length<2)throw fail('В name требуется имя и фамилия, разделённые пробелом.');
    return {first:parts[0],last:parts.slice(1).join(' '),display:parts.join(' ')};
  }
  function tBioHTML(s){
    const v=String(s??'').trim();if(!v)return '';
    if(/<\/?[a-z][\w-]*(?:\s[^>]*)?>/i.test(v))return cleanHTML(v);
    return v.split(/\n\s*\n/).map(part=>`<p>${esc(part).replace(/\n/g,'<br>')}</p>`).join('');
  }
  function tBioSig(s){
    const d=new DOMParser().parseFromString(String(s??''),'text/html');
    return (d.body.textContent||'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();
  }
  // WordPress may change <p>/<br> and typographic punctuation while saving,
  // even when the instructor's visible Bio is unchanged. Preserve word content;
  // do not treat arbitrary substrings or similar-length text as equal.
  function tBioSemanticText(html, typography=false){
    const doc=new DOMParser().parseFromString(String(html??''),'text/html');
    const blocks=new Set(['P','DIV','LI','UL','OL','H1','H2','H3','H4','H5','H6',
      'BLOCKQUOTE','SECTION','ARTICLE','PRE','TABLE','TR','TD','TH']);
    let out='';
    function walk(node){
      if(node.nodeType===3){out+=node.nodeValue||'';return;}
      if(node.nodeType!==1)return;
      const tag=node.tagName.toUpperCase();
      if(tag==='BR'){out+=' ';return;}
      if(blocks.has(tag))out+=' ';
      for(const child of node.childNodes)walk(child);
      if(blocks.has(tag))out+=' ';
    }
    for(const child of doc.body.childNodes)walk(child);
    let text=out.normalize('NFC')
      .replace(/[\u00a0\u1680\u2000-\u200a\u2028\u2029\u202f\u205f\u3000]/g,' ')
      .replace(/[\u00ad\u200b\u2060\ufeff]/g,'')
      .replace(/\s+/g,' ').trim();
    if(typography)text=text.replace(/[\u2018\u2019\u02bc]/g,"'")
      .replace(/[\u201c\u201d]/g,'"')
      .replace(/[\u2010-\u2015]/g,'-')
      .replace(/\u2026/g,'...');
    return text;
  }
  function tBioCompare(actual,expected){
    if(tBioSig(actual)===tBioSig(expected))return 'exact';
    if(tBioSemanticText(actual)===tBioSemanticText(expected))return 'format';
    if(tBioSemanticText(actual,true)===tBioSemanticText(expected,true))return 'typography';
    return '';
  }
  function tDiff(actual, expected){
    const out=[];
    if(actual.login!==expected.login)out.push('Login');
    if(actual.first!==expected.first)out.push('First Name');
    if(actual.last!==expected.last)out.push('Last Name');
    if(actual.display!==expected.display)out.push('Display name');
    if(expected.bio!==undefined && !tBioCompare(actual.bio,expected.bio))out.push('Profile Bio');
    if(expected.img!==undefined && String(actual.img)!==String(expected.img))out.push('Profile Photo ID');
    return out;
  }
  function tPersist(){tState.updated=new Date().toISOString();GM_setValue(TUTOR_KEY,tState);}
  function tLog(level,msg,row=null){
    tState.logs.push({time:new Date().toISOString(),level,row:row?.excelRow||'',login:row?.login||'',stage:tState.phase,message:String(msg).slice(0,1600)});
    if(tState.logs.length>1500)tState.logs.splice(0,tState.logs.length-1500);
    tPersist();tRender();
  }
  function tGuard(allowStop=false){authCheck();if(!owned())throw fail('Очередь управляется другой вкладкой.',true);if(tHalted&&!allowStop)throw STOP;}
  async function tWaitFor(fn,label,timeout=25000,allowStop=false){
    const begin=Date.now();
    while(Date.now()-begin<timeout){tGuard(allowStop);const x=fn();if(x)return x;await sleep(250);}
    throw fail('Истекло ожидание: '+label);
  }
  function tStage(r,s,message){tGuard(s==='saving'||s==='verifying');r.status=s;tState.phase=s;r.message=message;tLog('info',message,r);}
  function tFieldNode(id){
    const items=[...document.querySelectorAll('#'+id)];
    if(items.length!==1)throw fail(`Не найдено однозначное поле #${id}.`);
    return items[0];
  }
  async function tSetField(id,value){
    tGuard();const el=tFieldNode(id);if(!visible(el))throw fail('Поле скрыто: #'+id);
    el.focus();nativeValue(el,value);await sleep(180);tGuard();
    if(tFieldNode(id).value!==String(value))throw fail('Поле отклонило новое значение: #'+id);
  }
  async function tSelectDisplay(name){
    tGuard();const sel=tFieldNode('display_name');
    if(sel.tagName!=='SELECT')throw fail('Display name не является списком.');
    // WordPress generates options from the old first/last name on initial render.
    let option=[...sel.options].find(o=>o.value===name);
    if(!option){option=document.createElement('option');option.value=name;option.textContent=name;sel.append(option);}
    sel.value=name;
    sel.dispatchEvent(new unsafeWindow.Event('input',{bubbles:true}));
    sel.dispatchEvent(new unsafeWindow.Event('change',{bubbles:true}));
    await sleep(250);tGuard();
    if(tFieldNode('display_name').value!==name)throw fail('Display name сброшен скриптом сайта. Update не выполнен.');
  }
  function tBioIframeBody(){
    // TinyMCE renders its editable document inside a same-origin iframe.
    // This is the actual <body id="tinymce" contenteditable="true"> supplied
    // by Tutor LMS, not the WordPress admin document body.
    const iframe=document.getElementById('_tutor_profile_bio_ifr');
    try{
      const body=iframe?.contentDocument?.body;
      return body?.id==='tinymce'&&body.isContentEditable ? body : null;
    }catch{return null;}
  }
  function tActiveBioEditor(){
    const editor=tiny('_tutor_profile_bio');
    return editor?.initialized&&!editor.isHidden() ? editor : null;
  }
  function tReadBio(){
    const area=tFieldNode('_tutor_profile_bio');
    if(visible(area))return area.value;
    const editor=tActiveBioEditor();
    if(editor)return editor.getContent();
    // Some Tutor LMS pages render the iframe before tinymce.get() is exposed.
    return tBioIframeBody()?.innerHTML ?? area.value;
  }
  function tPersistedBioSources(){
    const sources=[];
    const area=tFieldNode('_tutor_profile_bio');
    sources.push({name:'textarea',html:area.value});
    const editor=tActiveBioEditor();
    if(editor){
      try{sources.push({name:'TinyMCE',html:editor.getContent()});}catch{}
    }
    const body=tBioIframeBody();
    if(body)sources.push({name:'iframe',html:body.innerHTML});
    return sources;
  }
  function tPersistedBioCheck(expected){
    const sources=tPersistedBioSources();
    for(const source of sources){
      const method=tBioCompare(source.html,expected);
      if(method)return {ok:true,source:source.name,method,html:source.html};
    }
    const target=tBioSemanticText(expected);
    const info=sources.map(source=>{
      const actual=tBioSemanticText(source.html);
      let at=0;while(at<Math.min(actual.length,target.length)&&actual[at]===target[at])at++;
      const position=actual===target?0:at+1;
      return `${source.name}=${actual.length}/${target.length}, позиция=${position}`;
    }).join('; ');
    return {ok:false,info};
  }
  async function tVerifyPersistedBio(expected){
    let last;
    // Tutor LMS initializes TinyMCE asynchronously; a single early snapshot
    // may reflect a stale visual editor, while the saved textarea is correct.
    for(let i=0;i<20;i++){
      tGuard(true);
      last=tPersistedBioCheck(expected);
      if(last.ok)return last;
      if(i<19)await sleep(250);
    }
    return last;
  }
  function tBioDebug(expected){
    const area=tFieldNode('_tutor_profile_bio');
    const editor=tActiveBioEditor();
    const body=tBioIframeBody();
    const visual=editor?.getContent() ?? (!visible(area)?body?.innerHTML:undefined);
    const actualPost=tBioSig(area.value),actualVisual=visual===undefined?null:tBioSig(visual);
    const diffAt=(actual)=>{
      if(actual===null)return null;
      let n=0;while(n<Math.min(actual.length,expected.length)&&actual[n]===expected[n])n++;
      return n===actual.length&&n===expected.length?0:n+1;
    };
    const form=area.form;
    const formBio=form?new unsafeWindow.FormData(form).get('_tutor_profile_bio'):null;
    return {
      postOK:actualPost===expected,
      visualOK:actualVisual===null||actualVisual===expected,
      formOK:formBio!==null&&tBioSig(formBio)===expected,
      editorActive:!!editor,
      iframePresent:!!body,
      textareaVisible:visible(area),
      expectedLength:expected.length,
      postLength:actualPost.length,
      visualLength:actualVisual?.length??null,
      postFirstDifference:diffAt(actualPost),
      visualFirstDifference:diffAt(actualVisual)
    };
  }
  function tBioFailureDetails(check){
    return `POST=${check.postOK}(${check.postLength}/${check.expectedLength}, позиция ${check.postFirstDifference}), `+
      `Visual=${check.visualOK}(${check.visualLength??'—'}/${check.expectedLength}, позиция ${check.visualFirstDifference??'—'}), `+
      `FormData=${check.formOK}, Text=${check.textareaVisible}`;
  }
  function tSetBioTextarea(html){
    // Set the value native WordPress will POST. Do not dispatch synthetic
    // change/blur events: Tutor may use them to restore the previous content.
    const area=tFieldNode('_tutor_profile_bio');
    const setter=Object.getOwnPropertyDescriptor(unsafeWindow.HTMLTextAreaElement.prototype,'value')?.set;
    if(setter)setter.call(area,String(html));else area.value=String(html);
  }
  async function tSwitchBioToText(){
    const id='_tutor_profile_bio';
    const area=tFieldNode(id);
    if(visible(area)&&!tActiveBioEditor())return;
    const wrap=document.getElementById('wp-'+id+'-wrap');
    const buttons=[...(wrap?.querySelectorAll('button.switch-html')||[])].filter(visible);
    if(buttons.length>1)throw fail('Найдено несколько кнопок Text для Profile Bio.');
    tGuard();
    if(buttons.length===1)buttons[0].click();
    else if(typeof unsafeWindow.switchEditors?.go==='function')unsafeWindow.switchEditors.go(id,'html');
    else throw fail('Не удалось найти штатную вкладку Text редактора Profile Bio.');
    await tWaitFor(()=>visible(tFieldNode(id))&&!tActiveBioEditor(),
      'переключение Profile Bio в режим Text',12000);
  }
  async function tWriteBio(html,row){
    const expected=tBioSig(html);
    await tWaitFor(()=>visible(tFieldNode('_tutor_profile_bio'))||tActiveBioEditor()||tBioIframeBody(),
      'редактор Profile Bio',25000);
    // Editing the iframe directly and calling TinyMCE.setContent caused both
    // the editor and the submitted textarea to revert on this Tutor LMS page.
    // Switch to WordPress Text mode, update the actual form control and KEEP
    // Text mode until the form is submitted; never call editor.save() here.
    await tSwitchBioToText();
    tLog('info','Profile Bio: WordPress переключён в Text. Заменяем старый текст в POST textarea.',row);
    const area=tFieldNode('_tutor_profile_bio');
    area.focus();
    tSetBioTextarea(html);
    let last=tBioDebug(expected);
    for(let i=0;i<7;i++){
      await sleep(250);tGuard();
      last=tBioDebug(expected);
      if(!last.textareaVisible||last.editorActive||!last.postOK||!last.formOK)break;
    }
    if(last.textareaVisible&&!last.editorActive&&last.postOK&&last.formOK){
      tLog('info',`Profile Bio установлен: Text=true, POST=true, FormData=true, символов=${expected.length}.`,row);
      return;
    }
    // Diagnostics contain only lengths and mismatch positions, NEVER Bio content.
    throw fail(`Profile Bio не удерживается в режиме Text: ${tBioFailureDetails(last)}. Update не выполнен.`);
  }
  function tPhotoBlock(){
    const inputs=[...document.querySelectorAll('.tutor-video-poster-wrap input[name="_tutor_profile_photo"]')];
    if(inputs.length!==1)throw fail('Не найдено единственное поле ID фото преподавателя (_tutor_profile_photo).');
    return inputs[0].closest('.tutor-video-poster-wrap');
  }
  function tPhotoID(){
    const block=tPhotoBlock();
    return block.querySelector('input[name="_tutor_profile_photo"]').value.trim();
  }
  function tSnapshot(){
    return {
      login:tFieldNode('user_login').value.trim(),
      first:tFieldNode('first_name').value.trim(),
      last:tFieldNode('last_name').value.trim(),
      display:tFieldNode('display_name').value,
      bio:tReadBio(),img:tPhotoID()
    };
  }
  async function tReady(allowStop=false){
    await tWaitFor(()=>{
      return document.querySelector('#user_login[disabled]')&&document.querySelector('#first_name')&&
        document.querySelector('#last_name')&&document.querySelector('#display_name')&&
        document.querySelector('#_tutor_profile_bio')&&
        document.querySelector('.tutor-video-poster-wrap input[name="_tutor_profile_photo"]')&&
        document.querySelector('input#submit[value="Update User"]');
    },'поля редактирования преподавателя',35000,allowStop);
    await sleep(350);tGuard(allowStop);
  }
  function tModal(){return [...document.querySelectorAll('.media-modal')].filter(visible).at(-1)||null;}
  function tMediaButton(modal){
    // Tutor LMS calls the confirm button "Upload media", not "Select".
    // Never click the button until the exact image is selected.
    const primary=[...modal.querySelectorAll('.media-toolbar-primary button.media-button-select')].filter(visible);
    if(primary.length>1)throw fail('Найдено несколько кнопок Upload media в одном окне.');
    if(primary.length===1)return primary[0].disabled?null:primary[0];
    const fallback=[...modal.querySelectorAll('button.media-button-select')].filter(visible);
    if(fallback.length>1)throw fail('Несколько кнопок подтверждения фото.');
    return fallback.length===1&&!fallback[0].disabled?fallback[0]:null;
  }
  function tPhotoFileStem(urlOrFilename){
    let text=String(urlOrFilename??'').trim();
    if(!text)return '';
    try{text=decodeURIComponent(new URL(text,ORIGIN).pathname.split('/').pop()||'');}
    catch{text=text.split(/[\\/]/).pop().split('?')[0];}
    // Media Library displays generated thumbnails such as 46020317267655-150x150.webp.
    // These are NOT WordPress attachment IDs; the number is a file-name stem.
    return text.replace(/\.(?:webp|avif|jpe?g|png|gif|bmp|tiff?)$/i,'')
      .replace(/-\d+x\d+$/,'').trim();
  }
  function tPhotoItemMatches(item,searchName){
    const wanted=tPhotoFileStem(searchName);
    if(!wanted)return false;
    const img=item.querySelector('.attachment-preview img, .thumbnail img');
    for(const src of [img?.getAttribute('src'),img?.getAttribute('data-src'),
      item.getAttribute('data-filename'),item.getAttribute('aria-label')]){
      if(src&&tPhotoFileStem(src)===wanted)return true;
    }
    return false;
  }
  async function tFindPhotoInLibrary(modal,searchName){
    let stable='',since=0;
    return tWaitFor(()=>{
      const input=modal.querySelector('#media-search-input');
      if(!input||input.value!==searchName){stable='';since=0;return null;}
      const spinners=[...modal.querySelectorAll('.spinner.is-active,.media-progress-bar')].filter(visible);
      if(spinners.length){stable='';since=0;return null;}
      const items=[...modal.querySelectorAll('.attachments .attachment')].filter(visible);
      const matches=items.filter(x=>tPhotoItemMatches(x,searchName));
      if(matches.length>1)throw fail('По имени файла '+searchName+' найдено несколько изображений. Выбор отменён.');
      if(matches.length!==1){stable='';since=0;return null;}
      // Allow Media Library's asynchronous search and thumbnail rendering to settle.
      const found=matches[0];
      const signature=(found.getAttribute('data-id')||'')+'|'+(found.querySelector('img')?.getAttribute('src')||'');
      if(signature!==stable||!since){stable=signature;since=Date.now();return null;}
      return Date.now()-since>=1000?found:null;
    },'изображение по имени файла '+searchName+' (проверьте наличие в Media Library)',30000);
  }
  async function tWritePhoto(imgName,row){
    tGuard();const block=tPhotoBlock();
    const key=String(imgName).trim();
    const upload=[...block.querySelectorAll('button.tutor_video_poster_upload_btn')].filter(visible);
    if(upload.length!==1)throw fail('Найдена неоднозначная кнопка Upload для Profile Photo.');
    if(tModal())throw fail('Открыто другое окно медиатеки. Закройте его перед запуском.');
    tLog('info','Открываем Media Library. Ищем имя изображения '+key+' (не attachment ID).',row);
    upload[0].click();
    const modal=await tWaitFor(()=>tModal(),'окно Media Library',16000);
    let searchInput=one('#media-search-input',modal,false);
    if(!searchInput){
      const browse=modal.querySelector('#menu-item-browse')||[...modal.querySelectorAll('.media-menu-item')].find(x=>/Media Library|Медиафайлы/i.test(x.textContent));
      if(!browse)throw fail('Не найдена вкладка Media Library.');
      browse.click();
      searchInput=await tWaitFor(()=>one('#media-search-input',modal,false),'поиск Media Library',12000);
    }
    searchInput.focus();nativeValue(searchInput,key);
    searchInput.dispatchEvent(new unsafeWindow.KeyboardEvent('keyup',{bubbles:true,key:'Enter'}));
    const item=await tFindPhotoInLibrary(modal,key);
    const attachmentID=String(item.getAttribute('data-id')||'').trim();
    if(!/^[1-9]\d*$/.test(attachmentID))
      throw fail('Изображение '+key+' найдено, но его настоящий WordPress attachment ID отсутствует. Update не выполнен.');
    const preview=item.querySelector('.attachment-preview.js--select-attachment')||item.querySelector('.attachment-preview');
    if(!preview)throw fail('Найдена картинка, но нет элемента .attachment-preview для выбора.');
    tGuard();
    if(item.getAttribute('aria-checked')!=='true')preview.click();
    await tWaitFor(()=>item.getAttribute('aria-checked')==='true','выбор картинки '+key,10000);
    const selected=[...modal.querySelectorAll('.attachments .attachment[aria-checked="true"]')].filter(visible);
    if(selected.length!==1||selected[0]!==item)
      throw fail('В медиатеке выделено не одно нужное изображение. Update не выполнен.');
    tLog('info','Изображение '+key+' выбрано кликом по превью; WordPress attachment ID = '+attachmentID,row);
    const button=await tWaitFor(()=>tMediaButton(modal),'активная кнопка Upload media',12000);
    tGuard();button.click();
    await tWaitFor(()=>!visible(modal),'закрытие Media Library после Upload media',12000);
    await tWaitFor(()=>tPhotoID()===attachmentID,
      'сохранение реального attachment ID '+attachmentID+' в _tutor_profile_photo',12000);
    tLog('info','Фото '+key+' установлено; поле _tutor_profile_photo = '+attachmentID+'.',row);
    return attachmentID;
  }
  function tTicket(kind,url){sessionStorage.setItem(TUTOR_TICKET,JSON.stringify({tab:tabId,at:Date.now(),run:tState.run,kind,url}));}
  function tNavigate(url){
    tGuard(true);tPersist();tTicket('nav',url);tLeaving=true;
    let unloaded=false;window.addEventListener('pagehide',()=>{unloaded=true;},{once:true});
    setTimeout(()=>{
      if(unloaded||!tLeaving||!owned())return;
      tLeaving=false;tBusy=false;tHalted=true;tState.running=false;sessionStorage.removeItem(TUTOR_TICKET);
      tLog('error','Не завершился переход. Проверьте диалог браузера и нажмите «Продолжить».',tState.rows[tState.index]);
      release();tRender();
    },20000);
    if(new URL(url).href===location.href)location.reload();else location.assign(url);
  }
  function tStop(pause=false){
    tState.paused=pause;
    const current=tState.rows[tState.index];
    if(current&&['saving','verifying'].includes(current.status))tState.stopAfterVerify=true;
    tHalted=true;tState.running=false;
    if(!tBusy)release();
    tPersist();tLog('info',pause?'Пауза запрошена.':'Остановка запрошена.');
  }
  async function tDelay(){
    const until=Date.now()+tState.settings.delay*1000;
    while(Date.now()<until){tGuard();await sleep(Math.min(250,until-Date.now()));}
  }
  async function tSubmit(row){
    tGuard();
    const button=tFieldNode('submit');
    if(button.tagName!=='INPUT'||button.type!=='submit'||button.value!=='Update User'||button.disabled)
      throw fail('Кнопка Update User отсутствует или недоступна.');
    const form=button.form;
    if(!form||!form.contains(tFieldNode('user_login')))throw fail('Update User относится к другой форме.');
    // Validate the bytes that would be submitted by the real WordPress form.
    // In Text mode editor.save() would overwrite the textarea with stale
    // iframe content. DO NOT call it when the editor is hidden.
    if(row.expected?.bio!==undefined){
      const expected=tBioSig(row.expected.bio);
      const area=tFieldNode('_tutor_profile_bio');
      if(area.form!==form)throw fail('Profile Bio не принадлежит форме Update User. Сохранение отменено.');
      if(!visible(area)||tActiveBioEditor())
        throw fail('Profile Bio переключён из Text обратно в Visual перед Update. Сохранение отменено.');
      const check=tBioDebug(expected);
      if(!check.postOK||!check.formOK)
        throw fail(`Перед Update User Profile Bio различается: ${tBioFailureDetails(check)}. Сохранение отменено.`);
    }
    tStage(row,'saving','Отправляем Update User; затем проверяем данные на новой странице');
    row.saveDocument=DOCUMENT_ID;row.saveAt=Date.now();tPersist();
    tTicket('post',row.url);tLeaving=true;
    let unloaded=false;window.addEventListener('pagehide',()=>{unloaded=true;},{once:true});
    button.click();
    const end=Date.now()+18000;
    while(Date.now()<end){if(unloaded)return;await sleep(300);}
    // A POST might have reached the server; never click Update twice automatically.
    tLeaving=false;sessionStorage.removeItem(TUTOR_TICKET);
    throw fail('Update User не привёл к загрузке новой страницы. Сохранение неопределённо; проверьте вручную.');
  }
  async function tRun(){
    if(tBusy)return;tBusy=true;tLeaving=false;
    try{
      while(true){
        tGuard();let i=tState.index;
        if(i<0||!tState.rows[i]||!ACTIVE.has(tState.rows[i].status))i=tState.rows.findIndex(r=>r.status==='pending');
        if(i<0){tState.running=false;tState.index=-1;tState.phase='idle';tPersist();tLog('ok','Очередь преподавателей завершена.');break;}
        tState.index=i;const row=tState.rows[i];
        try{
          if(tCurrentID()!==row.id){
            if(!row.expected||!['saving','verifying'].includes(row.status))tStage(row,'opening','Переход к профилю преподавателя');
            else tLog('info','Сохранение было отправлено; возвращаемся к профилю для проверки без повторного Update.',row);
            tNavigate(row.url);return;
          }
          await tReady(['saving','verifying'].includes(row.status));
          // Never edit a profile when the disabled WordPress login differs from Excel.
          if(tFieldNode('user_login').value.trim()!==row.login)throw fail(`Несовпадение login: Excel=${row.login}; профиль=${tFieldNode('user_login').value}. Сохранение запрещено.`);
          if(row.expected&&['saving','verifying'].includes(row.status)){
            if(row.saveDocument===DOCUMENT_ID){tStage(row,'verifying','Для проверки необходима новая загрузка редактора');tNavigate(row.url);return;}
            tStage(row,'verifying','Сверяем результат Update User после перезагрузки');
            const result=tSnapshot();
            const bioCheck=row.expected.bio===undefined?null:await tVerifyPersistedBio(row.expected.bio);
            // Keep actual saved content for the audit/report; never substitute
            // expected text to force a success result.
            if(bioCheck?.ok)result.bio=bioCheck.html;
            const errors=tDiff(result,row.expected);
            row.checked=new Date().toISOString();row.after={...result};
            if(bioCheck?.ok)tLog('info',`Profile Bio после сохранения: совпало (${bioCheck.source}, ${bioCheck.method}).`,row);
            else if(bioCheck)tLog('warn','Profile Bio после сохранения различается: '+bioCheck.info,row);
            if(errors.length){row.status='review';row.message='После сохранения не совпали: '+errors.join(', ');tLog('error',row.message,row);}
            else{row.status='done';row.message='Имя, публичное имя, Bio и ID фото проверены после загрузки';tLog('ok',row.message,row);}
            tState.phase='idle';tState.index=-1;tPersist();
            if(tSingle||tHalted||(row.status==='review'&&tState.settings.stopOnError)){tState.running=false;tPersist();break;}
            await tDelay();continue;
          }
          tStage(row,'filling','Сверяем исходный профиль и заполняем поля');
          row.before=tSnapshot();tPersist();
          await tSetField('first_name',row.first);
          await tSetField('last_name',row.last);
          await tSelectDisplay(row.display);
          const selectedPhotoID=row.img?await tWritePhoto(row.img,row):null;
          if(selectedPhotoID)row.selectedPhotoID=selectedPhotoID;
          // Set Bio last: WP Media Library may re-render TinyMCE on close.
          if(row.bio)await tWriteBio(row.bioHTML,row);
          // Blank bio/img leaves existing profile data untouched.
          const target={login:row.login,first:row.first,last:row.last,display:row.display,
            ...(row.bio?{bio:row.bioHTML}:{}),...(row.img?{img:selectedPhotoID}:{})};
          const current=tSnapshot(), errors=tDiff(current,target);
          if(errors.length)throw fail('Перед Update поля отличаются от Excel: '+errors.join(', '));
          row.expected=target;tPersist();
          await tSubmit(row);
          return; // Normally pagehide stops the execution before reaching this line.
        }catch(err){
          if(err===STOP)throw err;
          row.status=row.saveAt&&row.expected?'review':'error';row.message=err.message||String(err);
          tState.index=-1;tState.phase='idle';tPersist();tLog('error',row.message,row);
          if(err.fatal||tHalted||tSingle||tState.settings.stopOnError){tState.running=false;tPersist();break;}
          await tDelay();
        }
      }
    }catch(err){if(err!==STOP)tLog('error',err.message||String(err));tState.running=false;tPersist();}
    finally{if(!tLeaving){tBusy=false;tState.running=false;tPersist();release();tRender();}}
  }
  async function tStart(oneOnly=false,automatic=false){
    if(busy||tBusy){tMessage('Уже выполняется другая очередь.');return;}
    tState=GM_getValue(TUTOR_KEY,tEmpty());
    try{
      authCheck();await acquire();tHalted=false;tSingle=oneOnly||!!tState.stopAfterVerify;
      tState.stopAfterVerify=false;
      tState.paused=false;tState.owner=tabId;tState.run=automatic?tState.run:uid();
      tState.running=true;tState.single=tSingle;tPersist();tSetMode('tutors');
      await tRun();
    }catch(err){tState.running=false;tPersist();release();tMessage(err.message||String(err));tLog('error',err.message||String(err));}
  }

  function tSetMode(mode){
    if(mode!=='courses'&&mode!=='tutors')return;
    mainMode=mode;GM_setValue(MODE_KEY,mode);
    root.getElementById('courseBody')?.classList.toggle('hidden',mode!=='courses');
    root.getElementById('tutorBody')?.classList.toggle('hidden',mode!=='tutors');
    root.querySelectorAll('[data-main-mode]').forEach(b=>b.classList.toggle('on',b.dataset.mainMode===mode));
    if(mode==='tutors')tRender();else render();
  }
  function tMessage(s){message(s);}
  function tEditable(){
    const l=GM_getValue(LEASE,null);
    if(busy||tBusy||(l&&l.expires>Date.now())){tMessage('Сначала остановите очередь и дождитесь освобождения вкладки.');return false;}
    tState=GM_getValue(TUTOR_KEY,tEmpty());return true;
  }
  function tTabSet(tab){
    tTab=tab;tPage=0;
    for(const name of ['queue','import','log','settings','help'])root.getElementById('t-'+name)?.classList.toggle('hidden',tab!==name);
    root.querySelectorAll('[data-tutor-tab]').forEach(b=>b.classList.toggle('on',b.dataset.tutorTab===tab));
    tRender();
  }
  function tMount(){
    const courseBody=panel.querySelector('.body');courseBody.id='courseBody';
    const mainNav=document.createElement('nav');mainNav.className='topmodes';
    mainNav.innerHTML='<button data-main-mode="courses">Курсы</button><button data-main-mode="tutors">Преподаватели</button>';
    courseBody.before(mainNav);
    const tutor=document.createElement('div');tutor.id='tutorBody';tutor.className='body hidden';
    tutor.innerHTML=`<div class="stats"><div class="card"><b id="t-total">0</b><span>Преподавателей</span></div><div class="card done"><b id="t-done">0</b><span>Выполнено</span></div><div class="card"><b id="t-pending">0</b><span>В очереди</span></div><div class="card bad"><b id="t-errors">0</b><span>Ошибки / сверка</span></div></div>
      <div class="progress"><i id="t-bar"></i></div><div class="current"><i class="dot" id="t-dot"></i><span id="t-current">Загрузите Excel с преподавателями</span></div>
      <div class="actions"><button class="primary" id="t-start">▶ Начать</button><button id="t-one">Один профиль</button><button id="t-pause">Ⅱ Пауза</button><button class="danger" id="t-stop">■ Стоп</button><button id="t-retry">Повторить ошибки</button><button id="t-report">↓ Отчёт</button></div>
      <nav class="tabs"><button class="on" data-tutor-tab="queue">Очередь</button><button data-tutor-tab="import">Импорт Excel</button><button data-tutor-tab="log">Журнал</button><button data-tutor-tab="settings">Настройки</button><button data-tutor-tab="help">Как работать</button></nav>
      <div class="view" id="t-queue"><div class="tools"><input id="t-search" type="search" placeholder="Имя, login, ID…"><select id="t-filter"><option value="all">Все статусы</option><option value="pending">В очереди</option><option value="done">Выполнено</option><option value="error">Ошибки и проверка</option></select></div><div id="t-rows"></div></div>
      <div class="view hidden" id="t-import"><div class="tools"><button class="danger" id="t-reset">Очистить очередь преподавателей</button></div><div class="drop" id="t-drop"><label for="t-file">Выбрать Excel-файл</label><input id="t-file" type="file" accept=".xlsx,.xls,.csv" hidden><small>Или перетащите сюда .xlsx / .xls / .csv</small><small id="t-filename">Только локальное чтение файла</small></div><div id="t-mapping" class="hidden"><div class="mapping"><label>Лист</label><select id="t-sheet"></select><label>url_edit</label><select id="t-mapURL"></select><label>login</label><select id="t-mapLogin"></select><label>name</label><select id="t-mapName"></select><label>bio</label><select id="t-mapBio"></select><label>img</label><select id="t-mapImg"></select></div><p class="hint">name: сначала имя, затем фамилия. Bio поддерживает текст и HTML. img — числовая часть имени файла фотографии (например 46020317267655.webp), НЕ внутренний ID WordPress. Пустые bio/img не заменяют существующие значения. Login проверяется до изменений.</p><button class="primary" id="t-load">Загрузить в очередь</button></div><p class="hint">Все 5 столбцов обязательны, значения bio и img могут быть пустыми. После загрузки начните с кнопки «Один профиль».</p></div>
      <div class="view hidden" id="t-log"><div class="tools"><button id="t-download-log">↓ Скачать журнал</button></div><div id="t-logs"></div></div>
      <div class="view hidden" id="t-settings"><label class="setting"><div>Пауза между преподавателями<small>Секунды после подтверждённого сохранения</small></div><input id="t-delay" type="number" min="1" max="60" value="3"></label><label class="setting"><input id="t-stop-errors" type="checkbox"><div>Останавливать очередь при ошибке<small>Неопределённое сохранение отмечается «Проверить».</small></div></label><p class="hint">URL профиля вида /wp-admin/user-edit.php?user_id=ID. Публичное имя формируется из name. Фото ищется по имени файла из img; внутренний attachment ID WordPress считывается из медиатеки и сверяется после перезагрузки.</p><button id="t-backup">↓ Резервная копия JSON</button></div>
      <div class="view help hidden" id="t-help"><h3>Преподаватели — как работать</h3><ol><li>Откройте <strong>${esc(location.hostname)}</strong>/wp-admin/users.php?role=tutor_instructor и войдите под администратором. Перейдите в раздел «Преподаватели».</li><li>Во вкладке «Импорт Excel» выберите файл с пятью столбцами: <code>url_edit</code>, <code>login</code>, <code>name</code>, <code>bio</code>, <code>img</code>. Выберите нужный лист и проверьте сопоставление колонок.</li><li><code>url_edit</code> должен вести на <strong>${esc(location.hostname)}</strong>/wp-admin/user-edit.php?user_id=ID. Ссылки на другие сайты отклоняются. <code>login</code> должен точно совпадать с Username профиля.</li><li><code>name</code> — имя и фамилия через пробел (первое слово → First Name, остальная часть → Last Name). Display name устанавливается полным именем. <code>bio</code> — новый текст/HTML описания. <code>img</code> — числовое имя файла фотографии, например 46020317267655 для 46020317267655.webp; это не attachment ID.</li><li>Нажмите «Загрузить в очередь» → «Один профиль». Убедитесь, что имя, Bio и фото на сайте правильные, затем нажмите «Начать» для остальных записей.</li><li>Скрипт проверяет login, меняет имя и фото, записывает Bio через WordPress Text, нажимает Update User и после перезагрузки проверяет поля. Пустые значения bio/img оставляют существующие данные.</li><li>В конце скачайте «Отчёт». Ошибки откройте кнопкой «⋯» в строке или в «Журнале»; повторите только после проверки причины.</li></ol><h3>Пауза, повтор и проверка</h3><p>«Пауза» и «Стоп» не отменяют уже отправленный Update User. При продолжении скрипт проверяет результат до нового сохранения. Статус «Проверить» требует проверки на самом сайте; если расхождение только в формате Bio, после личной проверки доступно «Подтвердить Bio вручную».</p><h3>Несколько сайтов</h3><p>Очередь преподавателей сайта <strong>${esc(location.hostname)}</strong> хранится отдельно от очереди курсов и отдельно от всех остальных доменов. Запускайте импорт для каждого сайта в его собственной админ-панели, без переходов между сайтами во время обработки.</p></div>
      <footer class="foot"><span id="t-footer">Очередь хранится в Tampermonkey</span><button id="t-prev">←</button><span id="t-pages" style="flex:none">1 / 1</span><button id="t-next">→</button></footer>`;
    panel.append(tutor);
    root.querySelectorAll('[data-main-mode]').forEach(b=>b.addEventListener('click',()=>tSetMode(b.dataset.mainMode)));
    root.querySelectorAll('[data-tutor-tab]').forEach(b=>b.addEventListener('click',()=>tTabSet(b.dataset.tutorTab)));
    const on=(id,ev,fn)=>root.getElementById(id).addEventListener(ev,fn);
    on('t-start','click',()=>tStart(false));on('t-one','click',()=>tStart(true));
    on('t-pause','click',()=>tStop(true));on('t-stop','click',()=>tStop(false));
    on('t-search','input',e=>{tSearch=e.target.value;tPage=0;tRenderRows();});
    on('t-filter','change',e=>{tFilter=e.target.value;tPage=0;tRenderRows();});
    on('t-prev','click',()=>{tPage=Math.max(0,tPage-1);tRenderRows();});
    on('t-next','click',()=>{tPage++;tRenderRows();});
    on('t-file','change',e=>tReadFile(e.target.files[0]));on('t-sheet','change',tMapping);
    const drop=root.getElementById('t-drop');
    drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('over');});
    drop.addEventListener('dragleave',()=>drop.classList.remove('over'));
    drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('over');tReadFile(e.dataTransfer.files[0]);});
    on('t-load','click',tLoadQueue);on('t-report','click',tReport);on('t-download-log','click',tDownloadLog);
    on('t-backup','click',()=>download(JSON.stringify(tState,null,2),EXPORT_PREFIX+'tutor-manager-backup.json','application/json'));
    on('t-delay','change',e=>{if(!tEditable())return;tState.settings.delay=Math.min(60,Math.max(1,Number(e.target.value)||3));tPersist();tRender();});
    on('t-stop-errors','change',e=>{if(!tEditable())return;tState.settings.stopOnError=e.target.checked;tPersist();tRender();});
    on('t-retry','click',()=>{if(!tEditable())return;let n=0;tState.rows.forEach(r=>{if(r.status==='error'&&!r.invalid){r.status='pending';r.message='Повторная попытка';delete r.expected;delete r.saveAt;n++;}});tLog('info',`Возвращено в очередь: ${n}`);});
    on('t-reset','click',()=>{
      if(!tEditable()||!confirm('Удалить только очередь преподавателей, историю и журнал? Данные курсов останутся.'))return;
      tState=tEmpty();tPersist();tBook=null;tImportName='';tPage=0;tSearch='';tFilter='all';
      root.getElementById('t-file').value='';root.getElementById('t-mapping').classList.add('hidden');
      root.getElementById('t-filename').textContent='Только локальное чтение файла';tTabSet('import');
    });
    if(!tState.rows.length)tTab='import';
    tTabSet(tTab);
    // Prefer the teacher screen when opened directly from Tutor LMS users listing.
    if(location.pathname==='/wp-admin/users.php'&&new URLSearchParams(location.search).get('role')==='tutor_instructor')mainMode='tutors';
    if(tCurrentID() && (tState.running||mainMode==='tutors'))mainMode='tutors';
    tSetMode(mainMode);
  }
  function tRender(){
    if(!root?.getElementById('t-total'))return;
    const count=statuses=>tState.rows.filter(r=>statuses.includes(r.status)).length;
    root.getElementById('t-total').textContent=tState.rows.length;
    root.getElementById('t-done').textContent=count(['done']);
    root.getElementById('t-pending').textContent=count(['pending',...ACTIVE]);
    root.getElementById('t-errors').textContent=count(['error','review']);
    root.getElementById('t-bar').style.width=(tState.rows.length?100*count(['done','skipped'])/tState.rows.length:0)+'%';
    const r=tState.rows[tState.index];
    root.getElementById('t-current').textContent=tBusy?`${tHalted?'Остановка':LABEL[r?.status]||'Обработка'} · ${r?.name||r?.login||''}`:(tState.paused?'Пауза · нажмите «Продолжить»':tState.rows.length?`${tState.filename} · ${tState.rows.length} профилей`:'Загрузите Excel с преподавателями');
    root.getElementById('t-dot').classList.toggle('run',tBusy);
    root.getElementById('t-start').textContent=tState.paused||tState.rows.some(r=>ACTIVE.has(r.status))?'▶ Продолжить':'▶ Начать';
    const can=tState.rows.some(r=>r.status==='pending'||ACTIVE.has(r.status));
    root.getElementById('t-start').disabled=tBusy||busy||!can;
    root.getElementById('t-one').disabled=tBusy||busy||!can;
    root.getElementById('t-stop').disabled=!tBusy;
    root.getElementById('t-pause').disabled=!tBusy||tHalted;
    root.getElementById('t-retry').disabled=tBusy||busy||!tState.rows.some(r=>r.status==='error'&&!r.invalid);
    root.getElementById('t-reset').disabled=tBusy||busy;
    root.getElementById('t-load').disabled=tBusy||busy;
    root.getElementById('t-delay').disabled=tBusy||busy;
    root.getElementById('t-stop-errors').disabled=tBusy||busy;
    root.getElementById('t-delay').value=tState.settings.delay;
    root.getElementById('t-stop-errors').checked=!!tState.settings.stopOnError;
    root.getElementById('t-footer').textContent=tState.updated?'Сохранено локально · '+new Date(tState.updated).toLocaleTimeString():'Очередь хранится в Tampermonkey';
    if(tTab==='queue')tRenderRows();if(tTab==='log')tRenderLog();
  }
  function tRenderRows(){
    const rows=tState.rows.filter(r=>(tFilter==='all'||(tFilter==='error'?['error','review'].includes(r.status):r.status===tFilter))&&`${r.name} ${r.login} ${r.id}`.toLowerCase().includes(tSearch.toLowerCase()));
    const pages=Math.max(1,Math.ceil(rows.length/25));tPage=Math.min(tPage,pages-1);
    root.getElementById('t-pages').textContent=`${tPage+1} / ${pages}`;
    root.getElementById('t-prev').disabled=tPage===0;root.getElementById('t-next').disabled=tPage>=pages-1;
    const el=root.getElementById('t-rows');
    if(!rows.length){el.innerHTML='<div class="empty"><b>Нет строк</b>Импортируйте Excel или измените фильтр.</div>';return;}
    el.innerHTML=`<table><thead><tr><th style="width:45px">№</th><th>Преподаватель</th><th style="width:80px">img</th><th style="width:110px">Статус</th><th style="width:35px"></th></tr></thead><tbody>${rows.slice(tPage*25,tPage*25+25).map(r=>`<tr class="${tState.rows[tState.index]===r?'currentrow':''}"><td>${r.excelRow}</td><td><span class="title" title="${esc(r.name)}">${esc(r.name)}</span><small>${esc(r.login)} · ID ${esc(r.id||'—')} · ${esc(r.message)}</small></td><td>${esc(r.img||'—')}</td><td><span class="badge ${r.status}">${LABEL[r.status]}</span></td><td><button data-t-detail="${r.key}">⋯</button></td></tr>`).join('')}</tbody></table>`;
    el.querySelectorAll('[data-t-detail]').forEach(b=>b.onclick=()=>tDetail(tState.rows.find(r=>r.key===b.dataset.tDetail)));
  }
  function tDetail(r){
    root.querySelector('.detail')?.remove();const d=document.createElement('div');d.className='detail';
    d.innerHTML=`<header><b>Строка ${r.excelRow} · ${esc(r.login)}</b><button data-close>Закрыть</button></header><p><span class="badge ${r.status}">${LABEL[r.status]}</span> ${esc(r.message)}</p><p><strong>${esc(r.name)}</strong></p><p>Имя: ${esc(r.first)} · Фамилия: ${esc(r.last)}</p><p>Display name: ${esc(r.display)}</p><p>Фото (имя файла): ${esc(r.img||'не изменяется')} · attachment ID: ${esc(r.selectedPhotoID||'ещё не выбран')}</p><p><a href="${esc(r.url||'#')}" target="_blank" rel="noopener noreferrer">Открыть профиль</a></p><h4>Profile Bio</h4><pre>${esc(r.bio||'не изменяется')}</pre><div class="actions"><button data-verify>Повторить проверку</button><button data-confirm-bio class="hidden">Подтвердить Bio вручную</button><button data-skip>Пропустить</button><button data-reset>Вернуть в очередь</button></div>`;
    d.querySelector('[data-close]').onclick=()=>d.remove();
    d.querySelector('[data-verify]').disabled=!r.expected||!r.saveAt;
    d.querySelector('[data-verify]').onclick=()=>{if(!tEditable())return;const row=tState.rows.find(x=>x.key===r.key);row.status='verifying';row.saveDocument=DOCUMENT_ID;tState.index=tState.rows.indexOf(row);tState.phase='verifying';tPersist();d.remove();tStart(true);};
    const manualBioOnly=r.status==='review'&&!!r.expected&&!!r.saveAt&&!!r.after&&
      r.message.includes('Profile Bio')&&tDiff({...r.after,bio:r.expected.bio},r.expected).length===0;
    const confirmBio=d.querySelector('[data-confirm-bio]');
    if(manualBioOnly)confirmBio.classList.remove('hidden');
    confirmBio.onclick=()=>{
      if(!tEditable())return;
      const row=tState.rows.find(x=>x.key===r.key);
      if(!row||row.status!=='review'||!row.expected||!row.after||
        tDiff({...row.after,bio:row.expected.bio},row.expected).length!==0){tMessage('Статус изменился или имеются другие расхождения.');return;}
      if(!confirm('Подтверждаете, что лично открыли профиль на сайте и проверили, что Bio правильный? Скрипт не будет повторно нажимать Update User.'))return;
      row.status='done';row.manualBioVerified=true;row.checked=new Date().toISOString();
      row.message='Bio подтверждено вручную после проверки на сайте (без повторного сохранения)';
      tLog('warn',row.message,row);tPersist();tRender();d.remove();
    };
    d.querySelector('[data-skip]').onclick=()=>{if(!tEditable())return;const row=tState.rows.find(x=>x.key===r.key);row.status='skipped';if(tState.index===tState.rows.indexOf(row))tState.index=-1;tPersist();tRender();d.remove();};
    d.querySelector('[data-reset]').onclick=()=>{if(!tEditable())return;const row=tState.rows.find(x=>x.key===r.key);if(row.invalid){tMessage('Исправьте Excel и импортируйте заново.');return;}if(['done','review','saving','verifying'].includes(row.status)&&!confirm('Повторно заполнить и отправить Update User?'))return;row.status='pending';delete row.expected;delete row.saveAt;row.message='Возвращено в очередь';tPersist();tRender();d.remove();};
    panel.append(d);
  }
  function tRenderLog(){root.getElementById('t-logs').innerHTML=tState.logs.slice().reverse().map(l=>`<div class="logline ${esc(l.level)}"><time>${new Date(l.time).toLocaleTimeString()}</time><span>${l.row?'№ '+l.row:'Система'}</span><span>${esc(l.message)}</span></div>`).join('')||'<div class="empty">Журнал пока пуст</div>';}
  async function tReadFile(file){
    if(!file||!tEditable())return;
    try{
      if(typeof XLSX==='undefined')throw fail('Библиотека Excel (SheetJS) не загружена.');
      if(file.size>25*1024*1024)throw fail('Максимальный размер Excel — 25 МБ.');
      tBook=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});tImportName=file.name;
      root.getElementById('t-filename').textContent=file.name;
      root.getElementById('t-sheet').innerHTML=tBook.SheetNames.map((n,i)=>`<option value="${i}">${esc(n)}</option>`).join('');
      root.getElementById('t-mapping').classList.remove('hidden');tMapping();tTabSet('import');
    }catch(err){tMessage(err.message||String(err));}
  }
  function tSheetData(){return XLSX.utils.sheet_to_json(tBook.Sheets[tBook.SheetNames[Number(root.getElementById('t-sheet').value)]],{header:1,defval:'',raw:false,blankrows:true});}
  function tMapping(){
    const headers=tSheetData()[0]||[];
    for(const [name,alias] of Object.entries({URL:'urledit',Login:'login',Name:'name',Bio:'bio',Img:'img'})){
      const sel=root.getElementById('t-map'+name);
      sel.innerHTML='<option value="">Выберите колонку…</option>'+headers.map((h,i)=>`<option value="${i}">${esc(h||`Колонка ${i+1}`)}</option>`).join('');
      const i=headers.findIndex(h=>normalizeHeader(h)===alias);sel.value=i<0?'':String(i);
    }
  }
  function tLoadQueue(){
    if(!tBook||!tEditable())return;
    try{
      const names=['URL','Login','Name','Bio','Img'],cols=names.map(n=>root.getElementById('t-map'+n).value);
      if(cols.some(x=>x==='')||new Set(cols).size!==5)throw fail('Выберите 5 разных столбцов: url_edit, login, name, bio, img.');
      const [urlCol,loginCol,nameCol,bioCol,imgCol]=cols.map(Number),data=tSheetData(),seen=new Set(),rows=[];
      for(let i=1;i<data.length;i++){
        const line=data[i];if(line.every(v=>!String(v).trim()))continue;
        const row={key:uid(),excelRow:i+1,sourceURL:String(line[urlCol]??'').trim(),login:String(line[loginCol]??'').trim(),name:String(line[nameCol]??'').trim(),bio:String(line[bioCol]??'').trim(),img:String(line[imgCol]??'').trim(),status:'pending',message:'',id:'',url:'',first:'',last:'',display:'',bioHTML:''};
        try{
          const u=tURL(row.sourceURL);row.id=u.id;row.url=u.url;
          if(seen.has(row.id))throw fail('Повторный user_id '+row.id+' в файле.');seen.add(row.id);
          if(!row.login)throw fail('Пустой login.');
          Object.assign(row,tSplitName(row.name));
          if(row.img&&!/^[1-9]\d*$/.test(row.img))throw fail('img должен содержать числовое имя изображения (например 46020317267655), не attachment ID WordPress. Избегайте научной записи Excel.');
          row.bioHTML=tBioHTML(row.bio);
        }catch(err){row.status='error';row.invalid=true;row.message=err.message||String(err);}
        rows.push(row);
      }
      if(!rows.length)throw fail('На листе нет строк для обработки.');
      if(tState.rows.length&&!confirm('Заменить текущую очередь преподавателей? При необходимости скачайте отчёт.'))return;
      const settings=tState.settings;tState=tEmpty();tState.settings=settings;tState.rows=rows;tState.filename=tImportName;
      tPage=0;tPersist();tTabSet('queue');tLog('info',`Импортировано ${rows.length} профилей. Ошибок формата: ${rows.filter(r=>r.invalid).length}.`);
    }catch(err){tMessage(err.message||String(err));}
  }
  function tReport(){
    download(csv([['Excel Row','User ID','Login','Name','First Name','Last Name','Display Name','Bio','img filename','Selected attachment ID','url_edit','Status','Message','Checked At','Original First','Original Last','Original Display Name','Original Bio','Original Photo ID','Saved First','Saved Last','Saved Display Name','Saved Bio','Saved Photo ID'],
      ...tState.rows.map(r=>[r.excelRow,r.id,r.login,r.name,r.first,r.last,r.display,r.bio,r.img,r.selectedPhotoID,r.sourceURL,LABEL[r.status],r.message,r.checked,r.before?.first,r.before?.last,r.before?.display,r.before?.bio,r.before?.img,r.after?.first,r.after?.last,r.after?.display,r.after?.bio,r.after?.img])]),EXPORT_PREFIX+'tutor-manager-report.csv','text/csv;charset=utf-8');
  }
  function tDownloadLog(){download(csv([['Time','Level','Excel Row','Login','Stage','Message'],...tState.logs.map(l=>[l.time,l.level,l.row,l.login,l.stage,l.message])]),EXPORT_PREFIX+'tutor-manager-log.csv','text/csv;charset=utf-8');}
  async function tResumeOnLoad(){
    let ticket=null;try{ticket=JSON.parse(sessionStorage.getItem(TUTOR_TICKET)||'null');}catch{}
    sessionStorage.removeItem(TUTOR_TICKET);
    const row=tState.rows[tState.index];
    const mustVerify=!!row&&['saving','verifying'].includes(row.status)&&!!row.expected&&!!tState.stopAfterVerify;
    if(ticket?.tab===tabId && ticket.run===tState.run && Date.now()-ticket.at<120000 &&
       (tState.running||mustVerify) && row && location.origin===ORIGIN){
      tLog('info','Новая страница загружена: продолжаем очередь преподавателей.',tState.rows[tState.index]);
      await tStart(tState.single,true);
    }else if(tState.running){
      const l=GM_getValue(LEASE,null);
      if(!l||l.expires<Date.now()||l.tab===tabId){tState.running=false;tPersist();tLog('info','Автопереход не подтверждён. Нажмите «Продолжить».',tState.rows[tState.index]);}
    }
  }

  // Test hook is available only in a Node VM, never on the web page.
  if(typeof module!=='undefined'&&module.exports){module.exports={ORIGIN,SITE_SUFFIX,KEY,TUTOR_KEY,MODE_KEY,LEASE,courseURL,price,slug,normalizeHeader,csv,htmlSignature,cleanHTML,differences,tURL,tSplitName,tBioHTML,tDiff,tReadBio,tWriteBio,tBioIframeBody,tBioDebug,tSetBioTextarea,tActiveBioEditor,tSwitchBioToText,tBioFailureDetails,tBioSemanticText,tBioCompare,tPersistedBioCheck,tVerifyPersistedBio};return;}
  GM_getTab(async info=>{
    tabId=info.ycTabId||uid();info.ycTabId=tabId;GM_saveTab(info);
    let ticket=null;try{ticket=JSON.parse(sessionStorage.getItem(TICKET)||'null');}catch{}sessionStorage.removeItem(TICKET);
    mount();captureErrors();
    let sameCourse=false;try{sameCourse=!!ticket&&courseURL(ticket.url).id===currentCourse();}catch{}
    if(ticket&&ticket.tab===tabId&&ticket.run===state.run&&Date.now()-ticket.at<90000&&sameCourse){
      log('info','Страница загружена: продолжаем сохранённый этап',state.rows[state.index]);
      await start(state.single,true);
    }else if(state.running){
      const l=GM_getValue(LEASE,null);if(!l||l.expires<Date.now()||l.tab===tabId){state.running=false;persist();log('info','Автопереход не подтверждён. Нажмите «Продолжить» для возобновления.',state.rows[state.index]);}
    }
    if(!busy)await tResumeOnLoad();
    if(location.pathname.includes('wp-login.php')){
      state.running=false;tState.running=false;tPersist();release();log('error','Требуется вход администратора. После входа откройте wp-admin и нажмите «Продолжить».',state.rows[state.index]);
      message('Войдите администратором, затем откройте wp-admin и нажмите «Продолжить».');
    }
  });
})();
