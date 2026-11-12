// ==UserScript==
// @name         SMMmaster — массовое изменение дат товаров по ID
// @namespace    https://smmmaster.net/
// @version      1.2.0
// @description  Читает ID из Excel/CSV/TXT, находит товар в WooCommerce, меняет дату через Quick Edit → Update; журнал, пауза и отчет.
// @match        https://smmmaster.net/wp-admin/edit.php*
// @match        https://www.smmmaster.net/wp-admin/edit.php*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        none
// @run-at       document-idle
// @noframes
// ==/UserScript==

(() => {
  'use strict';
  if (new URLSearchParams(location.search).get('post_type') !== 'product') return;
  if (window.__SMM_RANDOM_DATES__) return;
  window.__SMM_RANDOM_DATES__ = true;

  const STORAGE = 'smmmaster_dates_v1_state';
  const LAYOUT = 'smmmaster_dates_v1_layout';
  const FROM = Date.UTC(2026, 0, 1, 0, 0, 0);
  const TO = Date.UTC(2026, 10, 10, 23, 59, 59);
  const MAX_STORED_LOG = 60;
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  let busy = false;
  let rowsFromFile = [];
  let activeFileName = '';
  let activeSheet = '';

  const emptyState = () => ({
    ids: [], index: 0, status: 'idle', phase: 'search',
    records: [], notes: [], lastMessage: 'Загрузите файл с ID',
    batchRemaining: null, dateCandidate: null, source: '',
    missStreak: 0, delayMs: 1300, dryRun: false,
    lastChanged: Date.now()
  });
  let state;
  try { state = { ...emptyState(), ...JSON.parse(localStorage.getItem(STORAGE) || '{}') }; }
  catch { state = emptyState(); }
  if (!Array.isArray(state.ids)) state.ids = [];
  if (!Array.isArray(state.records)) state.records = [];
  if (!Array.isArray(state.notes)) state.notes = [];

  const save = () => {
    state.lastChanged = Date.now();
    try { localStorage.setItem(STORAGE, JSON.stringify(state)); }
    catch (e) { console.error('[SMM] Не удалось сохранить очередь в localStorage:', e); }
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
  const randomDate = () => {
    let vals;
    do {
      const n = Math.floor(FROM / 1000) + Math.floor(Math.random() * (Math.floor(TO / 1000) - Math.floor(FROM / 1000) + 1));
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
  #smm-date-tool textarea,#smm-date-tool input[type=number],#smm-date-tool select{width:100%;background:#121d30;border:1px solid #3a506e;border-radius:9px;color:white;padding:9px;font:inherit}
  #smm-date-tool textarea{height:65px;min-height:50px;resize:vertical}
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
  `;
  document.head.append(style);

  const panel = document.createElement('section');
  panel.id = 'smm-date-tool';
  panel.innerHTML = `
    <div class="head" id="smm-grip">
      <div class="badge">◷</div>
      <div><strong>SMMmaster · Даты товаров</strong><small>Quick Edit · без удаления · v1.2</small></div>
      <div class="spacer"></div><button class="iconbtn" id="smm-collapse" title="Свернуть/развернуть" type="button">−</button>
    </div>
    <div class="main">
      <div class="row"><button id="smm-import" type="button" class="primary">⬆ Загрузить Excel / CSV / TXT</button><input id="smm-file" type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" /></div>
      <div id="smm-file-info" class="sub">Данные сохраняются в этом браузере между перезагрузками.</div>
      <div id="smm-column-wrap" style="display:none"><label class="caption">Столбец с ID</label><select id="smm-column"></select><button id="smm-accept-column" class="mini green" type="button" style="margin-top:7px">Применить этот столбец</button></div>
      <details><summary>Или вставить ID текстом</summary><textarea id="smm-paste" placeholder="285067\n204665\n278600"></textarea><button id="smm-add-pasted" class="mini" type="button">Загрузить ID из поля</button></details>
      <div class="stats"><div class="stat"><strong id="smm-total">0</strong><span>Всего ID</span></div><div class="stat"><strong id="smm-ok">0</strong><span>Готово</span></div><div class="stat"><strong id="smm-missing">0</strong><span>Не найдены</span></div><div class="stat"><strong id="smm-error">0</strong><span>Ошибки</span></div></div>
      <div class="pills"><span id="smm-progress-label">0 / 0</span><span id="smm-progress-pct">0%</span></div><div class="bar"><div class="fill" id="smm-fill"></div></div>
      <div class="state" id="smm-state">Загрузите ID для начала</div>
      <div class="options"><label style="display:flex;align-items:center;gap:5px"><input type="checkbox" id="smm-dry" /> Проверка поиска без изменения дат</label></div>
      <div class="row" style="margin-top:10px"><button id="smm-test" class="warn" type="button">Тест: 3 ID</button><button id="smm-start" class="primary" type="button">▶ Старт / продолжить</button><button id="smm-pause" type="button">Ⅱ Пауза</button></div>
      <div class="row" style="margin-top:7px"><button id="smm-report" type="button" class="green">↓ Отчет XLSX</button><button id="smm-missing-csv" type="button">↓ Нет ID (CSV)</button><button id="smm-reset" type="button" class="danger">↺ Сброс</button></div>
      <label class="caption">История операций</label><div class="log" id="smm-log" aria-live="polite"></div>
      <div class="sub">Дата выбирается случайно с 01.01.2026 по 10.11.2026; время 00:00 исключено. Товары не удаляются. При 5 последовательных отсутствующих ID — автопауза.</div>
    </div>`;
  document.body.append(panel);

  let layout = {};
  try { layout = JSON.parse(localStorage.getItem(LAYOUT) || '{}'); } catch {}
  if (Number.isFinite(layout.left) && Number.isFinite(layout.top)) {
    panel.style.left = Math.max(0, Math.min(layout.left, window.innerWidth - 70)) + 'px';
    panel.style.top = Math.max(0, Math.min(layout.top, window.innerHeight - 55)) + 'px';
    panel.style.right = 'auto';
  }
  if (Number.isFinite(layout.width)) panel.style.width = Math.max(340, layout.width) + 'px';
  if (Number.isFinite(layout.height)) panel.style.height = Math.max(440, layout.height) + 'px';
  if (layout.collapsed) panel.classList.add('collapsed');
  $('#smm-collapse').textContent = panel.classList.contains('collapsed') ? '+' : '−';

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
    ids = unique(ids);
    if (!ids.length) { log('Не удалось найти числовые ID в выбранном столбце', 'error'); return; }
    if (state.records.length && !confirm('Заменить текущий список и сбросить историю? Перед заменой скачайте отчет, если он нужен.')) return;
    state = emptyState();
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
      if (/\.(xlsx|xls|csv|tsv)$/i.test(file.name) && typeof XLSX !== 'undefined') {
        const bytes = await file.arrayBuffer();
        const wb = XLSX.read(bytes, {type:'array'});
        sheetName = wb.SheetNames[0];
        rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1, raw:false, defval:''});
      } else if (/\.(xlsx|xls)$/i.test(file.name)) {
        throw new Error('Библиотека Excel не загрузилась. Проверьте интернет или сохраните файл как CSV.');
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
    if (verified === false) {
      state.status = 'paused'; state.phase = 'saving';
      log(`ID ${id}: ответ Update получен, но дата в таблице отличается. Проверьте товар вручную; автоматизация остановлена.`, 'error');
      save(); return;
    }
    finish('updated', verified === true ? 'дата сохранена и проверена' : 'Update подтвержден WordPress', date.text);
  }
  async function searchId(id) {
    const field = document.getElementById('post-search-input');
    const submit = document.getElementById('search-submit');
    if (!field || !submit) throw new Error('Не найдена строка поиска #post-search-input / #search-submit');
    // WordPress keeps a hidden Quick Edit template inside the GET form.
    // Disable its controls so post_title, prices and nonce do not leak into search URLs.
    $$('#inline-edit input, #inline-edit select, #inline-edit textarea, #bulk-edit input, #bulk-edit select, #bulk-edit textarea').forEach(el => {el.disabled = true;});
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
  async function tick() {
    if (busy || state.status !== 'running') return;
    busy = true;
    try {
      while (state.status === 'running' && state.index < state.ids.length) {
        const id = state.ids[state.index];
        const visibleId = new URLSearchParams(location.search).get('s') || '';
        if (visibleId !== id) {
          await sleep(Math.min(1000, state.delayMs));
          if (state.status !== 'running') break;
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
    if (!state.ids.length) { log('Сначала загрузите список ID', 'error'); return; }
    if (state.index >= state.ids.length) { log('Все ID обработаны. Сбросьте или загрузите новый список.'); return; }
    if (state.phase === 'saving') {
      if (!confirm('Последний Update мог выполниться, но ответ не был получен. Проверьте дату текущего товара на сайте. Повторно попробовать этот ID?')) return;
      state.phase = 'search'; state.dateCandidate = null;
    }
    state.batchRemaining = limit;
    state.status = 'running';
    state.dryRun = $('#smm-dry').checked;
    save(); draw();
    log(`${limit ? 'Тестовая партия: '+limit+' ID' : 'Запуск обработки'}. ${state.dryRun ? 'Только поиск; никаких обновлений' : 'Изменение дат включено'}`);
    tick();
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
  function reportExcel() {
    if (typeof XLSX === 'undefined') {
      log('Не загружена Excel-библиотека. Отчет сохранен в CSV.', 'error');
      exportCsv(reportRows(), 'smmmaster_report.csv'); return;
    }
    const wb = XLSX.utils.book_new();
    for (const [sheet, subset] of [
      ['Все результаты',state.records], ['Обновлены',state.records.filter(r => r.status === 'updated')],
      ['Не найдены',state.records.filter(r => r.status === 'missing')],
      ['Ошибки',state.records.filter(r => r.status === 'error')]
    ]) {
      const ws = XLSX.utils.aoa_to_sheet(reportRows(subset));
      ws['!cols'] = [{wch:15},{wch:28},{wch:24},{wch:65}];
      XLSX.utils.book_append_sheet(wb, ws, sheet);
    }
    XLSX.writeFile(wb, 'smmmaster_dates_report.xlsx');
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
    state = emptyState(); save(); draw();
  });
  $('#smm-dry').addEventListener('change', e => { state.dryRun = e.target.checked; save(); });
  $('#smm-collapse').addEventListener('click', () => {
    panel.classList.toggle('collapsed');
    $('#smm-collapse').textContent = panel.classList.contains('collapsed') ? '+' : '−';
    persistLayout();
  });
  function persistLayout() {
    const r = panel.getBoundingClientRect();
    localStorage.setItem(LAYOUT, JSON.stringify({left:Math.round(r.left),top:Math.round(r.top),width:Math.round(r.width),height:Math.round(r.height),collapsed:panel.classList.contains('collapsed')}));
  }
  (() => {
    const grip = $('#smm-grip');
    let drag = null;
    grip.addEventListener('pointerdown', (e) => {
      if (e.target.closest('button')) return;
      if (e.button !== 0) return;
      const r = panel.getBoundingClientRect();
      drag = {pointerId:e.pointerId, dx:e.clientX-r.left,dy:e.clientY-r.top};
      grip.setPointerCapture(e.pointerId);
    });
    grip.addEventListener('pointermove', e => {
      if (!drag || e.pointerId !== drag.pointerId) return;
      const left = Math.max(0, Math.min(window.innerWidth - 65, e.clientX-drag.dx));
      const top = Math.max(0, Math.min(window.innerHeight - 55, e.clientY-drag.dy));
      panel.style.left = left + 'px'; panel.style.top = top + 'px'; panel.style.right = 'auto';
    });
    grip.addEventListener('pointerup', e => { if (drag && drag.pointerId === e.pointerId) {drag=null;persistLayout();} });
    grip.addEventListener('pointercancel', () => {drag=null;persistLayout();});
    if (typeof ResizeObserver !== 'undefined') {
      let t;
      new ResizeObserver(() => { clearTimeout(t); t=setTimeout(persistLayout,350); }).observe(panel);
    }
  })();

  if (state.status === 'running' && state.phase === 'saving') {
    state.status = 'paused';
    state.lastMessage = 'Страница перезагрузилась во время Update. Проверьте текущий ID вручную до продолжения.';
    log(state.lastMessage, 'error');
  }
  draw();
  if (state.status === 'running') {
    log(`Возобновление после загрузки страницы: ID ${state.ids[state.index]}`);
    setTimeout(tick, 650);
  }
})();