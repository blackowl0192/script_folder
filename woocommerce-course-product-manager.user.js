// ==UserScript==
// @name         WooCommerce — Course Product Manager (HottestCourse / SMMMaster)
// @namespace    blackowl0192.woocommerce.course.products
// @version      1.0.4
// @description  Excel → WooCommerce products: title, permalink, HTML, regular price, ACF lessons, downloadable off, product image. Per-site queue, verification and logs.
// @match        https://hottestcourse.com/wp-admin/*
// @match        https://hottestcourse.com/wp-login.php*
// @match        https://smmmaster.net/wp-admin/*
// @match        https://smmmaster.net/wp-login.php*
// @updateURL   https://raw.githubusercontent.com/blackowl0192/script_folder/main/woocommerce-course-product-manager.user.js
// @downloadURL https://raw.githubusercontent.com/blackowl0192/script_folder/main/woocommerce-course-product-manager.user.js
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
  const HOSTS = new Set(['hottestcourse.com', 'smmmaster.net']);
  if (location.protocol !== 'https:' || !HOSTS.has(location.hostname)) return;
  const ORIGIN = location.origin;
  const SITE = location.hostname;
  const KEY = `bwo.wc.course.product.v1.${SITE}`;
  const LEASE = `${KEY}.lease`;
  const TICKET = `${KEY}.nav`;
  const ALIAS = {id:['id','postid','productid'],name:['name','title','productname'],description:['description','content','info'],price:['regularprice','price'],lessons:['lessons','lesson'],img:['img','image','photo']};
  const STATUSES = {pending:'В очереди',opening:'Открывается',filling:'Заполняется',saving:'Сохраняется',verifying:'Проверяется',done:'Выполнено',review:'Проверить',error:'Ошибка',skipped:'Пропущено'};
  const ACTIVE = new Set(['opening','filling','saving','verifying']);
  const sleep = ms => new Promise(resolve=>setTimeout(resolve,ms));
  const uid = () => crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`;
  const DOC = uid();
  const error = (message,fatal=false) => Object.assign(new Error(message),{fatal});
  const STOP = Symbol('stopped');
  const esc = value => String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const empty = () => ({version:1,rows:[],logs:[],filename:'',run:'',index:-1,phase:'idle',running:false,paused:false,owner:'',single:false,stopAfterVerify:false,settings:{delay:3,stopOnError:true},updated:''});
  let state = GM_getValue(KEY,null)||empty();
  let tabId,root,panel,workbook=null,importFile='',working=false,leaving=false,halted=false,single=false,leaseTimer;
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
  function guard(allowStopped=false){authCheck();if(!owned())throw error('Другая вкладка управляет очередью либо завершилась аренда.',true);if(halted&&!allowStopped)throw STOP;}
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
    persist();leaving=true;
    const resolved=new URL(url,ORIGIN);
    if(resolved.origin!==ORIGIN)throw error('Запрещён переход на другой сайт.',true);
    if(resolved.href===location.href)location.reload();else location.assign(resolved.href);
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
        if(i<0||!state.rows[i]||!ACTIVE.has(state.rows[i].status))i=state.rows.findIndex(x=>x.status==='pending');
        if(i<0){state.running=false;state.index=-1;state.phase='idle';persist();log('ok','Очередь завершена. Скачайте отчёт.');break;}
        state.index=i;const r=state.rows[i];
        try{
          if(productIDFromURL(location.href)!==r.id){stage(r,'opening','Открываем товар WooCommerce');nav(r.url,'open');return;}
          await editorReady(['saving','verifying'].includes(r.status));
          if(r.expected&&['saving','verifying'].includes(r.status)){
            // A previous update may have finished. Never submit a second update during verification.
            if(r.saveDoc===DOC){stage(r,'verifying','Загружаем страницу заново, чтобы проверить результат Update');nav(r.url,'verify');return;}
            stage(r,'verifying','Проверяем сохранение на новой странице');
            const actual=await snapshot(true);
            r.actual=actual;r.checked=new Date().toISOString();
            const diff=differences(actual,r.expected);
            if(diff.length){r.status='review';r.message=`После сохранения не совпали: ${diff.join(', ')}. Повторный Update не отправлен.`;log('error',r.message,r);}
            else{r.status='done';r.message='Все значения подтверждены после перезагрузки';log('ok',r.message,r);}
            const stopNow=state.stopAfterVerify||single||halted||(r.status==='review'&&state.settings.stopOnError);
            state.stopAfterVerify=false;state.phase='idle';state.index=-1;persist();
            if(stopNow){state.running=false;persist();break;}
            await between();continue;
          }
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
          stage(r,'saving','Нажимаем Update; затем проверяем сохранение на новой странице');
          r.saveAt=Date.now();r.saveDoc=DOC;state.stopAfterVerify=single;persist();
          sessionStorage.setItem(TICKET,JSON.stringify({tab:tabId,run:state.run,id:r.id,kind:'verify',at:Date.now()}));
          guard();button.click();
          // Classic WordPress normally navigates on submit. If it does not, fetch a fresh
          // document; a failed save will be marked 'review' without duplicate submissions.
          await sleep(14000);
          if(halted)state.stopAfterVerify=true;
          if(!leaving){leaving=true;persist();location.assign(r.url);}
          return;
        }catch(ex){
          if(ex===STOP)throw ex;
          r.status=r.saveAt&&r.expected?'review':'error';r.message=ex.message||String(ex);
          state.phase='idle';state.index=-1;persist();log('error',r.message,r);
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
    state=GM_getValue(KEY,empty());
    try{
      authCheck();await acquire();halted=false;
      single=automatic?!!state.single:oneOnly;
      if(!automatic){state.run=uid();state.paused=false;state.single=oneOnly;}
      state.owner=tabId;state.running=true;
      const r=state.rows[state.index];
      if(r&&['saving','verifying'].includes(r.status)&&r.expected&&r.saveDoc===DOC){
        state.phase='verifying';r.status='verifying';persist();nav(r.url,'verify');return;
      }
      persist();render();await run();
    }catch(ex){state.running=false;persist();release();log('error',ex.message||String(ex));}
  }
  function editable(){
    if(working){notice('Сначала остановите обработку.');return false;}
    const lease=GM_getValue(LEASE,null);
    if(lease&&lease.expires>Date.now()&&lease.tab!==tabId){notice('Другая вкладка выполняет очередь.');return false;}
    state=GM_getValue(KEY,empty());return true;
  }
  function shuffle(rows){for(let i=rows.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[rows[i],rows[j]]=[rows[j],rows[i]];}return rows;}
  function sheetRows(){return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[Number($('sheet').value)]],{header:1,defval:'',raw:false,blankrows:true});}
  async function readFile(file){
    if(!file||!editable())return;
    try{
      if(typeof XLSX==='undefined')throw error('Не загрузилась библиотека Excel (SheetJS).');
      if(file.size>25*1024*1024)throw error('Файл больше 25 МБ.');
      workbook=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});importFile=file.name;
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
    const header=['Excel Row','Product ID','Name','Edit URL','Description','Regular price','lessons','img','Status','Message','Checked At','Original name','Original slug','Original price','Original lessons','Original image','Saved image','Saved WordPress attachment ID','Excel price (source)'];
    const rows=state.rows.map(r=>[r.excelRow,r.id,r.name,r.url,r.description,r.price,r.lessons,r.img,STATUSES[r.status],r.message,r.checked,r.before?.name,r.before?.slug,r.before?.price,r.before?.lessons,r.before?.image,r.actual?.image,r.actual?.attachmentId,r.rawPrice]);
    download(csv([header,...rows]),`${SITE}-woo-products-report.csv`,'text/csv;charset=utf-8');
  }
  function exportLog(){download(csv([['Time','Level','Excel Row','Product ID','Stage','Message'],...state.logs.map(x=>[x.time,x.level,x.row,x.id,x.stage,x.message])]),`${SITE}-woo-products-log.csv`,'text/csv;charset=utf-8');}
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
      row.status='verifying';row.message='Повторная проверка без Update';row.saveDoc=DOC;state.index=state.rows.indexOf(row);state.phase='verifying';persist();d.remove();start(true);
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
    const count=statuses=>state.rows.filter(r=>statuses.includes(r.status)).length;
    $('total').textContent=state.rows.length;$('done').textContent=count(['done']);$('pending').textContent=count(['pending',...ACTIVE]);$('failed').textContent=count(['error','review']);
    $('bar').style.width=(state.rows.length?100*count(['done','skipped'])/state.rows.length:0)+'%';
    const current=state.rows[state.index];
    $('current').textContent=working?(halted?'Остановка после текущей операции':`${STATUSES[current?.status]||'Работа'} · ${current?.name||current?.id||''}`):
      state.paused?'На паузе — нажмите Продолжить':state.rows.length?`${state.filename} · ${SITE}`:'Загрузите Excel с товарами';
    const hasWork=state.rows.some(x=>x.status==='pending'||ACTIVE.has(x.status));
    $('start').disabled=working||!hasWork;$('one').disabled=working||!hasWork;$('stop').disabled=!working;$('pause').disabled=!working||halted;
    $('start').textContent=state.paused||state.rows.some(x=>ACTIVE.has(x.status))?'▶ Продолжить':'▶ Начать';
    $('retry').disabled=working||!state.rows.some(x=>x.status==='error'&&!x.invalid);
    $('load').disabled=working;$('reset').disabled=working;
    $('delay').disabled=working;$('stopOnError').disabled=working;
    $('delay').value=state.settings.delay;$('stopOnError').checked=!!state.settings.stopOnError;
    $('updated').textContent=state.updated?`Сохранено ${new Date(state.updated).toLocaleTimeString()} · ${SITE}`:`Очередь хранится отдельно для ${SITE}`;
    root.querySelector('[data-tab="processed"]').textContent=`Обработанные (${count(['done','review'])})`;
    if(['queue','processed'].includes(tab))renderRows();if(tab==='log')renderLog();
  }
  const css=`
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
  function mount(){
    const host=document.createElement('div');host.id='bwo-woo-course-manager';host.style.cssText='position:fixed;inset:0;pointer-events:none;z-index:2147483646';document.documentElement.append(host);
    root=host.attachShadow({mode:'open'});
    root.innerHTML=`<style>${css}:host>section,:host>button{pointer-events:auto}</style>
      <button class="pill hidden" id="pill">Товары · открыть панель</button>
      <section class="panel" id="panel"><header class="head" id="drag"><div class="logo">W</div><div><b>WooCommerce · Курсы</b><small>${esc(SITE)} · v1.0.0</small></div><span class="space"></span><button id="full" title="На весь экран">⛶</button><button id="min" title="Свернуть">—</button></header>
      <div class="body"><div class="stats"><div class="stat"><b id="total">0</b><small>Всего товаров</small></div><div class="stat success"><b id="done">0</b><small>Готово</small></div><div class="stat"><b id="pending">0</b><small>В очереди / работе</small></div><div class="stat bad"><b id="failed">0</b><small>Ошибки / проверка</small></div></div>
      <div class="progress"><i id="bar"></i></div><div class="current" id="current">Загрузите Excel с товарами</div>
      <div class="actions"><button class="primary" id="start">▶ Начать</button><button id="one">Один товар</button><button id="pause">Ⅱ Пауза</button><button class="danger" id="stop">■ Стоп</button><button id="retry">Повторить ошибки</button><button id="report">↓ Отчёт</button></div>
      <nav class="tabs"><button data-tab="queue">Очередь</button><button data-tab="processed">Обработанные</button><button data-tab="import" class="active">Импорт Excel</button><button data-tab="log">Журнал</button><button data-tab="settings">Настройки</button><button data-tab="help">Как работать</button></nav>
      <div class="view hidden" id="queue"><div class="tools"><input type="search" id="search" placeholder="Название, ID или img…"><select id="filter"><option value="all">Все статусы</option><option value="pending">В очереди</option><option value="error">Ошибка</option><option value="skipped">Пропущено</option></select></div><div id="rows"></div></div>
      <div class="view" id="import"><div class="tools"><button class="danger" id="reset">Очистить очередь и историю</button></div><div class="drop" id="drop"><label for="file">Выбрать Excel-файл</label><input id="file" type="file" accept=".xlsx,.xls,.csv" hidden><small>Или перетащите сюда · .xlsx / .xls / .csv</small><small id="filename">Данные Excel обрабатываются локально</small></div>
      <div id="mapping" class="hidden"><div class="mapping"><label>ID товара</label><select id="map-id"></select><label>Name</label><select id="map-name"></select><label>Description (HTML)</label><select id="map-description"></select><label>Regular price</label><select id="map-price"></select><label>lessons</label><select id="map-lessons"></select><label>img (14 цифр)</label><select id="map-img"></select><label>Лист</label><select id="sheet"></select></div><p class="hint">При импорте цена 3.42€ распознаётся как 3.42, а в поле WooCommerce записывается 3,42. Колонка img — имя загруженного изображения из 14 цифр, НЕ WordPress attachment ID. Очередь перемешивается автоматически.</p><button class="primary" id="load">Загрузить в очередь</button></div>
      <div class="notice">Строки проверяются до запуска. Новая загрузка заменяет текущую очередь; скачайте отчёт, если он нужен. Обрабатываются только товары сайта <b>${esc(SITE)}</b>.</div></div>
      <div class="view hidden" id="log"><div class="tools"><button id="downloadLog">↓ Скачать журнал</button><span class="hint">Последние 2000 событий</span></div><div id="logs"></div></div>
      <div class="view hidden" id="settings"><label class="setting"><div>Пауза между товарами<small>В секундах после проверки результата</small></div><input id="delay" type="number" min="1" max="60" value="3"></label><label class="setting"><input id="stopOnError" type="checkbox" checked><div>Останавливаться при ошибке<small>Рекомендуется: ошибка может оставить несохранённые изменения в редакторе</small></div></label><div class="setting"><div>Перемешивание Excel<small>Однократно при каждом импорте; исходные номера строк сохраняются</small></div><span class="tag done">Всегда</span></div><div class="setting"><div>Downloadable<small>Флажок снимается у каждого товара</small></div><span class="tag done">Всегда</span></div><button id="backup">↓ Резервная копия JSON</button></div>
      <div class="view help hidden" id="help"><h3>1. Подготовьте Excel</h3><p>Первая строка: <b>ID, Name, Description, Regular price, lessons, img</b>. ID — существующий post ID товара. Description — HTML, как в редакторе WordPress. Regular price может быть с € или запятой. lessons — целое положительное число. img — 14-значное имя уже загруженной фотографии в Media Library (не внутренний attachment ID).</p>
      <h3>2. Импортируйте товары</h3><ol><li>Войдите в <b>${esc(SITE)}</b> как администратор.</li><li>Откройте «Импорт Excel», выберите файл, лист и шесть колонок.</li><li>Нажмите «Загрузить в очередь». Курсы будут перемешаны; номера исходных строк останутся в отчёте.</li><li>Нажмите «Один товар» и проверьте сохранённые поля на сайте; после этого нажмите «Начать» для остальных.</li></ol>
      <h3>3. Что делает скрипт</h3><p>Открывает <code>/wp-admin/post.php?post=ID&amp;action=edit</code>. Заполняет Name и slug, переключает редактор Description в Text, вставляет HTML и возвращается в Visual. Меняет Regular price и ACF lessons, снимает Downloadable, удаляет старую обложку, ищет точную фотографию по img и нажимает Set product image. После проверки полей нажимает Update и перезагружает товар для сверки.</p>
      <h3>4. Пауза, ошибки и безопасность</h3><p>Не редактируйте страницу вручную во время работы. Пауза останавливает новые действия, но не отменяет отправленный Update. После сохранения статус «Выполнено» ставится только после сверки на новой странице. Если данные не совпали — «Проверить», повторный Update автоматически не выполняется. Для такого товара доступна проверка без Update через «⋯».</p><p>Если произошла ошибка, исправьте причину, затем «Повторить ошибки». Неверные строки Excel исправляются только повторным импортом. Скачивайте отчёт и журнал. Сайты <b>hottestcourse.com</b> и <b>smmmaster.net</b> имеют независимые очереди и настройки.</p>
      <h3>Ограничения</h3><p>Скрипт рассчитан на Classic WordPress / WooCommerce и указанные элементы ACF. В отдельных темах или версиях админки разметка может отличаться. При отсутствии нужных полей скрипт останавливается до Update.</p></div>
      <footer class="foot"><span id="updated">Очередь хранится локально</span><button id="prev">←</button><span id="pages" style="flex:none">1 / 1</span><button id="next">→</button></footer></div></section>`;
    panel=$('panel');
    const on=(id,ev,fn)=>$(id).addEventListener(ev,fn);
    on('start','click',()=>start(false));on('one','click',()=>start(true));on('pause','click',()=>stop(true));on('stop','click',()=>stop(false));
    on('retry','click',retryErrors);on('report','click',exportReport);on('downloadLog','click',exportLog);on('reset','click',resetAll);
    on('load','click',loadQueue);on('backup','click',()=>download(JSON.stringify(state,null,2),`${SITE}-woo-products-backup.json`,'application/json'));
    on('file','change',e=>readFile(e.target.files[0]));on('sheet','change',setMapping);
    const drop=$('drop');drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('over');});drop.addEventListener('dragleave',()=>drop.classList.remove('over'));drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('over');readFile(e.dataTransfer.files[0]);});
    root.querySelectorAll('[data-tab]').forEach(b=>b.addEventListener('click',()=>switchTab(b.dataset.tab)));
    on('search','input',e=>{search=e.target.value;page=0;renderRows();});on('filter','change',e=>{filter=e.target.value;page=0;renderRows();});
    on('prev','click',()=>{page=Math.max(0,page-1);renderRows();});on('next','click',()=>{page++;renderRows();});
    on('delay','change',e=>{if(!editable())return;state.settings.delay=Math.min(60,Math.max(1,Number(e.target.value)||3));persist();render();});
    on('stopOnError','change',e=>{if(!editable())return;state.settings.stopOnError=e.target.checked;persist();render();});
    on('min','click',()=>{panel.classList.add('hidden');$('pill').classList.remove('hidden');});
    const show=()=>{panel.classList.remove('hidden');$('pill').classList.add('hidden');};
    on('pill','click',show);GM_registerMenuCommand('Открыть WooCommerce Course Manager',show);
    on('full','click',()=>{maximized=!maximized;panel.classList.toggle('full',maximized);panel.style.top='';panel.style.left='';panel.style.right='';panel.style.bottom='';});
    let drag=null;
    on('drag','pointerdown',e=>{if(e.target.closest('button')||maximized)return;const rect=panel.getBoundingClientRect();drag={x:e.clientX-rect.left,y:e.clientY-rect.top};e.currentTarget.setPointerCapture(e.pointerId);});
    on('drag','pointermove',e=>{if(!drag)return;panel.style.left=Math.max(0,Math.min(innerWidth-panel.offsetWidth,e.clientX-drag.x))+'px';panel.style.top=Math.max(0,Math.min(innerHeight-panel.offsetHeight,e.clientY-drag.y))+'px';panel.style.right='auto';panel.style.bottom='auto';});
    on('drag','pointerup',()=>drag=null);
    render();if(state.rows.length)switchTab('queue');
  }
  function captureErrors(){
    const seen=new Map();function record(text){if(!working)return;if(/Video loading failed/i.test(text))return;const at=seen.get(text)||0;if(Date.now()-at<30000)return;seen.set(text,Date.now());log('warn',`Предупреждение страницы: ${text}`,state.rows[state.index]);}
    window.addEventListener('error',e=>{if(e.message)record(e.message);},true);
    window.addEventListener('unhandledrejection',e=>record(String(e.reason?.message||e.reason||'Unhandled rejection')));
  }
  GM_getTab(async data=>{
    tabId=data.wcCourseProductTabId||uid();data.wcCourseProductTabId=tabId;GM_saveTab(data);
    let ticket;try{ticket=JSON.parse(sessionStorage.getItem(TICKET)||'null');}catch{ticket=null;}
    sessionStorage.removeItem(TICKET);
    mount();captureErrors();
    const row=state.rows[state.index];
    if(ticket&&row&&ticket.tab===tabId&&ticket.run===state.run&&ticket.id===row.id&&Date.now()-ticket.at<180000){
      log('info','Новая страница загружена: продолжаем очередь.',row);
      await start(state.single,true);
    }else if(state.running){
      state.running=false;state.paused=true;persist();log('warn','Автопереход не подтверждён. Нажмите «Продолжить» для возобновления.',row);
    }
  });
})();