// ==UserScript==
// @name         YourCourse — Excel → Tutor LMS Enrollments
// @namespace    https://yourcourse.shop/
// @version      1.4.0
// @description  Excel → курсы; первый найденный курс → Select; выбор 20–60 студентов; перетаскиваемая, сворачиваемая панель; Okay → следующий курс
// @author       YourCourse
// @match        https://yourcourse.shop/wp-admin/admin.php*
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @run-at       document-idle
// @grant        none
// ==/UserScript==

/*
  ВАЖНО
  - Работает только на странице Tutor LMS Enrollments в авторизованной вкладке.
  - На каждой странице списка (обычно 10 строк) нажимает «Name»,
    затем стрелку вправо, пока не выбрано 20–60 человек.
    На последней странице снимает лишние галочки при необходимости.
  - Перед Add проверяет отмеченные строки. При несовпадении
    счётчиков останавливается без Enroll Now.
  - Если структура таблицы / пагинация неизвестна, приостанавливается, НЕ нажимая Enroll Now.
  - Тест: обрабатывает только один курс и НЕ нажимает Enroll Now.
  - Курс: выбирает ПЕРВЫЙ из результатов поиска (даже если их несколько),
    выполняет имитацию наведения на карточку/цену и кликает Select.
  - Панель можно перетаскивать, менять размер, сворачивать; положение сохраняется.
  - После Enroll Now обрабатывает итоговое окно с кнопкой Okay,
    затем кликает Enroll Students, чтобы открыть следующий курс.
  - При ошибке/непонятном ответе сайта останавливается, не создавая дубль.
  - Данные Excel (только названия курсов) хранятся в sessionStorage этой вкладки.
*/

(() => {
  'use strict';
  const BASE = 'https://yourcourse.shop/wp-admin/admin.php?page=enrollments&action=add_new';
  const KEY = 'yourcourse_excel_enrollment_v1';
  const PANEL_KEY = 'yourcourse_enroll_panel_position_v13';
  const S = (s, scope = document) => scope.querySelector(s);
  const SA = (s, scope = document) => Array.from(scope.querySelectorAll(s));
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const norm = s => String(s || '').normalize('NFKC').replace(/\s+/g, ' ').trim().toLocaleLowerCase();
  const visible = el => !!el && el.isConnected && el.getClientRects().length > 0 &&
    getComputedStyle(el).visibility !== 'hidden' && getComputedStyle(el).display !== 'none';
  const byText = (scope, sel, text) => SA(sel, scope).find(e => visible(e) && norm(e.textContent) === norm(text));
  const rand = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
  const shuffle = arr => {
    const out = arr.slice();
    for (let i = out.length - 1; i > 0; i--) {
      const j = rand(0, i);
      [out[i], out[j]] = [out[j], out[i]];
    }
    return out;
  };
  const initial = () => ({
    titles: [], fileName: '', index: 0, done: [], status: 'idle', phase: 'idle',
    test: false, target: 0, selected: 0, log: [], manualReason: ''
  });
  let state;
  try { state = Object.assign(initial(), JSON.parse(sessionStorage.getItem(KEY) || '{}')); }
  catch { state = initial(); }
  let workbook = null;
  let jobId = 0;
  let active = false;
  let courseWasClicked = false;
  let ui = {};
  const isEnrollPage = new URL(location.href).searchParams.get('page') === 'enrollments';
  if (!isEnrollPage) return;

  function save() {
    // Never auto-repeat an uncertain submission on reload.
    sessionStorage.setItem(KEY, JSON.stringify(state));
  }
  function log(message) {
    const t = new Date().toLocaleTimeString('ru-RU');
    state.log.push(`[${t}] ${message}`);
    state.log = state.log.slice(-90);
    save(); render();
    console.log('[YourCourse]', message);
  }
  function status(status, reason = '') {
    state.status = status; state.manualReason = reason; save(); render();
  }
  function isAddNew() { return new URL(location.href).searchParams.get('action') === 'add_new'; }
  function current() { return state.titles[state.index] || ''; }
  function isStopped(id) { return id !== jobId || state.status === 'stopped'; }
  async function checkpoint(id) {
    if (isStopped(id)) throw new Error('STOP');
    while (state.status === 'paused' || state.status === 'manual_students') {
      await sleep(180);
      if (isStopped(id)) throw new Error('STOP');
    }
  }
  async function delay(ms, id) {
    let left = ms;
    while (left > 0) {
      await checkpoint(id);
      const step = Math.min(left, 150);
      await sleep(step); left -= step;
    }
  }
  async function waitFor(fn, id, timeout = 12000, interval = 180) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      await checkpoint(id);
      const item = fn();
      if (item) return item;
      await sleep(interval);
    }
    return null;
  }
  function click(el, name) {
    if (!visible(el) || el.disabled) throw new Error(`Недоступна кнопка «${name}»`);
    el.click();
  }
  function setReactValue(el, value) {
    // The native setter is required for React-controlled inputs.
    const proto = Object.getPrototypeOf(el);
    const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set ||
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
    setter.call(el, value);
    el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: value }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }
  function dataButton(cy, type) {
    return SA(`button[data-cy="${cy}"]${type ? `[type="${type}"]` : ''}`)
      .find(el => visible(el) && !el.disabled);
  }
  function courseSelectButtons() {
    // The site's Select button can be hidden until hovering the price/card.
    return SA('button[data-cy="select-course"]')
      .filter(b => !ui.root?.contains(b) && b.isConnected);
  }
  function isCoursePrice(el) {
    if (!el || !visible(el)) return false;
    if (el.matches('[data-cy*="price"],[class*="price"],[class*="Price"]')) return true;
    return /(?:\$|€|£|₽|₹|\b(?:free|usd|eur|price|mdl|lei)\b|\d[.,]\d{2})/i.test(el.textContent || '');
  }
  function courseCardFor(el, title = '') {
    const exactName = norm(title);
    const isSelect = !!el?.matches?.('button[data-cy="select-course"]');
    let item = el;
    let best = null;
    for (let depth = 0; item && item !== document.body && depth < 9; depth++, item = item.parentElement) {
      if (!visible(item)) continue;
      const t = norm(item.textContent);
      // Avoid treating the full dialog or whole page as a single card.
      if (t.length > 1200) break;
      if ((isSelect || exactName && t.includes(exactName)) &&
          (item.querySelector('button[data-cy="select-course"]') ||
           item.querySelector('[class*="price"], [class*="Price"], [data-cy*="price"]') ||
           /(?:\$|€|£|₽|₹|\bfree\b)/i.test(t))) best = item;
      if (item.matches('tr,li,[role="row"],[data-cy*="course"],[class*="course-card"],[class*="courseCard"]')) {
        if (isSelect || !exactName || t.includes(exactName)) return item;
      }
    }
    return best;
  }
  function hoverLikeMouse(el) {
    if (!el || !visible(el)) return;
    try { el.scrollIntoView({block:'nearest', inline:'nearest', behavior:'instant'}); }
    catch { el.scrollIntoView({block:'nearest'}); }
    const r = el.getBoundingClientRect();
    const x = Math.round(r.left + Math.min(r.width * 0.5, Math.max(1,r.width-1)));
    const y = Math.round(r.top + Math.min(r.height * 0.5, Math.max(1,r.height-1)));
    const opts = {bubbles:true, cancelable:true, composed:true, clientX:x, clientY:y, relatedTarget:null};
    // React uses bubbling mouseover to synthesize onMouseEnter.
    if (typeof PointerEvent === 'function') {
      el.dispatchEvent(new PointerEvent('pointerover', { ...opts, pointerType:'mouse', isPrimary:true }));
      el.dispatchEvent(new PointerEvent('pointerenter', { ...opts, bubbles:false, pointerType:'mouse', isPrimary:true }));
      el.dispatchEvent(new PointerEvent('pointermove', { ...opts, pointerType:'mouse', isPrimary:true }));
    }
    el.dispatchEvent(new MouseEvent('mouseover', opts));
    el.dispatchEvent(new MouseEvent('mouseenter', {...opts, bubbles:false}));
    el.dispatchEvent(new MouseEvent('mousemove', opts));
  }
  // Hover over course card and price, then expose CSS-hover-only button without
  // moving the real OS mouse (not possible from a userscript).
  function revealCourseOptions(title, searchInput) {
    const root = searchInput.closest('[role="dialog"], [aria-modal="true"]') || document;
    let cards = [];
    for (const button of courseSelectButtons()) {
      const card = courseCardFor(button, title);
      if (card && visible(card) && !cards.includes(card)) cards.push(card);
    }
    if (!cards.length) {
      // Some React UIs create Select only AFTER hovering a card. Try the exact title.
      const titleEls = SA('h2,h3,h4,h5,span,p,div,strong', root)
        .filter(el => visible(el) && norm(el.textContent) === norm(title) &&
          !Array.from(el.children).some(ch => norm(ch.textContent) === norm(title)));
      for (const el of titleEls.slice(0,10)) {
        const card = courseCardFor(el, title) || el.closest('li,tr,[role="row"]') || el.parentElement?.parentElement;
        if (card && visible(card) && !cards.includes(card)) cards.push(card);
      }
    }
    // Enumerate all visible cards in the returned search results, not only
    // the one whose title exactly matches the Excel text. Otherwise a UI that
    // renders buttons on hover could appear to have only one result.
    const priceNodes = SA('[data-cy*="price"],[class*="price"],[class*="Price"]', root)
      .filter(el => visible(el) && (el.textContent || '').length < 90);
    for (const price of priceNodes.slice(0,80)) {
      const card = price.closest('tr,li,[role="row"],[class*="course-card"],[class*="courseCard"]') ||
        price.parentElement?.parentElement;
      if (card && card !== root && visible(card) && !cards.includes(card)) cards.push(card);
    }
    // When price has no useful class, look for compact currency labels.
    if (cards.length <= 1) {
      for (const el of SA('span,strong,p', root).filter(e => visible(e) && e.children.length===0 &&
        (e.textContent || '').length < 55 && isCoursePrice(e)).slice(0,80)) {
        const card = el.closest('tr,li,[role="row"],[class*="course-card"],[class*="courseCard"]') ||
          el.parentElement?.parentElement;
        if (card && card !== root && visible(card) && !cards.includes(card)) cards.push(card);
      }
    }
    for (const card of cards) {
      hoverLikeMouse(card);
      const priceEls = SA('[data-cy*="price"],[class*="price"],[class*="Price"],span,div,strong', card)
        .filter(el => visible(el) && el.children.length < 3 && isCoursePrice(el));
      // Prefer a real price element and NOT the whole result card.
      const price = priceEls.sort((a,b) => a.textContent.length - b.textContent.length)[0];
      if (price) hoverLikeMouse(price);
    }
    // Only manipulate buttons whose own course card is visible; don't expose
    // hidden/stale results elsewhere on the page.
    const buttons = courseSelectButtons().filter(b => {
      if (visible(b)) return true;
      const card = courseCardFor(b, title);
      if (!card || !visible(card)) return false;
      for (let ancestor = b; ancestor && ancestor !== card; ancestor = ancestor.parentElement) {
        const cs = getComputedStyle(ancestor);
        if (cs.display === 'none') ancestor.style.setProperty('display','inline-flex','important');
        if (cs.visibility === 'hidden') ancestor.style.setProperty('visibility','visible','important');
        if (cs.opacity === '0') ancestor.style.setProperty('opacity','1','important');
        ancestor.style.setProperty('pointer-events','auto','important');
      }
      return visible(b);
    });
    return buttons;
  }
  // Find search on the recently opened Select Course modal, including portals
  // that do not advertise role="dialog". Do not depend on generated CSS classes or UUID ids.
  function findCourseSearch() {
    const candidates = SA('input[name="search"], input[placeholder="Search..."], input[placeholder*="Search"]')
      .filter(el => visible(el) && !ui.root?.contains(el));
    return candidates.find(el => el.closest('[role="dialog"], [aria-modal="true"]')) ||
      candidates.find(el => norm(el.getAttribute('placeholder')).includes('search')) ||
      candidates[0] || null;
  }
  async function searchCourse(name, id) {
    state.phase = 'course_search'; save();
    const open = await waitFor(() => dataButton('select-course-plan'), id, 15000);
    if (!open) throw new Error('Не найдена кнопка Select Course.');
    click(open, 'Select Course');
    const input = await waitFor(findCourseSearch, id, 12000);
    if (!input) throw new Error('Не найден поиск курсов.');
    input.focus();
    setReactValue(input, name);
    log(`Ищу курс: ${name}`);
    // Give React's debounce and the network search time to replace stale results.
    await delay(1900, id);
    const start = Date.now();
    let lastSignature = null;
    let stableSince = Date.now();
    let buttons = [];
    while (Date.now() - start < 11000) {
      await checkpoint(id);
      buttons = revealCourseOptions(name, input);
      const signature = buttons.map(b => norm(courseCardFor(b, name)?.textContent ||
        b.closest('tr,[role="row"],li')?.textContent || b.parentElement?.parentElement?.textContent || b.textContent))
        .join('|') + `#${buttons.length}`;
      if (signature !== lastSignature) {
        lastSignature = signature;
        stableSince = Date.now();
      }
      if (Date.now() - stableSince >= 1050 && Date.now() - start >= 1250) break;
      await sleep(160);
    }
    buttons = revealCourseOptions(name, input);
    if (buttons.length >= 1) {
      // Always choose the FIRST search result in document order, even when
      // the search yields 2, 3 or more courses. Hover price/card as before.
      const btn = buttons[0];
      const card = courseCardFor(btn, name);
      if (card) hoverLikeMouse(card);
      await delay(180, id);
      const liveButton = courseSelectButtons().find(b => b === btn || (card && card.contains(b)));
      if (!liveButton || !visible(liveButton)) throw new Error('Кнопка Select не появилась после наведения на цену.');
      click(liveButton, 'Select');
      const ready = await waitFor(() => dataButton('add-students', 'button'), id, 9500);
      if (!ready) throw new Error('Нажата Select, но кнопка Add Students не появилась. Проверьте выбор курса.');
      log(`Найдено курсов: ${buttons.length}. Автоматически выбран первый результат (наведение на цену → Select).`);
      await delay(400, id);
      return;
    }
    const why = 'Курсы не найдены (ни одной кнопки Select).';
    courseWasClicked = false;
    status('manual_course', `${why} Выберите курс вручную кнопкой Select.`);
    log(`Нужен ручной выбор курса: ${why}`);
    while (true) {
      if (isStopped(id)) throw new Error('STOP');
      if (courseWasClicked) {
        await sleep(500);
        const add = dataButton('add-students', 'button');
        if (add) {
          courseWasClicked = false;
          status('running');
          log('Курс выбран вручную, продолжаю.');
          return;
        }
      }
      await sleep(190);
    }
  }

  function labelForCheckbox(box) {
    if (!box) return null;
    const wrapping = box.closest('label');
    if (wrapping) return wrapping;
    if (box.id) return SA('label[for]').find(e => e.getAttribute('for') === box.id) || null;
    return null;
  }
  function usableCheckbox(box) {
    // Some Tutor LMS / React styles hide the native input but leave its label clickable.
    return box && box.isConnected && !box.disabled &&
      (visible(box) || visible(labelForCheckbox(box)));
  }
  function setCheckbox(box, checked, description) {
    if (!usableCheckbox(box)) throw new Error(`Не найдена доступная галочка: ${description}`);
    if (box.checked === checked) return;
    const label = labelForCheckbox(box);
    if (visible(box)) box.click();
    else if (visible(label)) label.click();
    if (box.isConnected && box.checked !== checked) {
      throw new Error(`Сайт не изменил галочку: ${description}`);
    }
  }
  function addStudentsDialog() {
    const add = SA('button[data-cy="add-students"][type="submit"]')
      .find(b => visible(b) && norm(b.textContent).includes('add'));
    if (!add) return null;
    const dialog = add.closest('[role="dialog"], [aria-modal="true"], .tutor-modal, .tutor-modal-container');
    if (dialog) return dialog;
    let p = add.parentElement;
    while (p && p !== document.body) {
      if (SA('input[type="checkbox"]', p).filter(usableCheckbox).length >= 2) return p;
      p = p.parentElement;
    }
    return null; // Do not search arbitrary checkboxes outside the Add Students popup.
  }
  function nameCheckbox(scope) {
    const boxes = SA('input[type="checkbox"]', scope).filter(usableCheckbox);
    return boxes.find(box => norm(box.getAttribute('label')) === 'name') ||
      boxes.find(box => norm(labelForCheckbox(box)?.textContent) === 'name' ||
        !!labelForCheckbox(box)?.querySelector('[title="Name"]')) || null;
  }
  function studentCheckBoxes(scope) {
    const nameBox = nameCheckbox(scope);
    const boxes = SA('input[type="checkbox"]', scope).filter(box => {
      if (!usableCheckbox(box) || box === nameBox) return false;
      if (box.closest('thead, [role="columnheader"], header')) return false;
      const description = norm(box.getAttribute('aria-label') || box.getAttribute('label'));
      if (/^(name|select all|all|выбрать всех)$/.test(description)) return false;
      if (/^(name|select all|all)$/.test(norm(labelForCheckbox(box)?.textContent))) return false;
      return true;
    });
    // Prefer explicit data rows, excluding possible filter-checkbox controls.
    const inRows = boxes.filter(box => box.closest('tbody tr, [role="rowgroup"] [role="row"], [data-row-key]'));
    const chosen = inRows.length ? inRows : boxes;
    // One selection checkbox per student row.
    const rowSeen = new Set();
    return chosen.filter(box => {
      const row = box.closest('tr, [role="row"], [data-row-key]');
      if (row && rowSeen.has(row)) return false;
      if (row) rowSeen.add(row);
      return true;
    });
  }
  function studentKey(box) {
    const row = box.closest('tr, [role="row"], [data-row-key]') || box.parentElement?.parentElement;
    const rowID = row?.getAttribute('data-row-key') || row?.getAttribute('data-id') || '';
    const value = box.value && box.value !== 'on' ? box.value : '';
    const content = norm(row?.textContent);
    // A UUID checkbox id is the last resort when the row has no display label.
    return `${rowID}|${value}|${content || box.id || ''}`;
  }
  function reportedStudentCount(scope) {
    // Some implementations show a selection count in an Add button or badge.
    // We only use explicit 'selected' wording, never the number of search results.
    const near = SA('button, [role="status"], [data-cy]', scope)
      .map(e => (e.textContent || '').trim()).filter(s => s.length < 100);
    for (const t of near) {
      const m = t.match(/(?:^|\s)(\d{1,6})\s+(?:students?\s+)?selected\b|\bselected\s*[:(]?\s*(\d{1,6})\b/i);
      if (m) return Number(m[1] || m[2]);
    }
    return null;
  }
  function nextStudentPage(scope) {
    const options = SA('button, a', scope).filter(el => visible(el) &&
      !el.disabled && el.getAttribute('aria-disabled') !== 'true');
    // Точная стрелка вправо из окна Add Students (её CSS-класс может измениться).
    // Не путаем со стрелкой влево / закрытием модального окна.
    const exactRightArrow = el => !!el.querySelector(
      'svg path[d*="M15 30.5L25 20.5L15 10.5"]');
    return options.find(exactRightArrow) || options.find(el => {
      const text = norm(el.textContent);
      const aria = norm(el.getAttribute('aria-label') || el.getAttribute('title'));
      const cy = norm(el.getAttribute('data-cy'));
      return /^(next|next page|следующая|далее|›|»|→)$/.test(text) ||
        /^(next|next page|следующая страница)$/.test(aria) ||
        /^(next-page|pagination-next)$/.test(cy);
    }) || null;
  }
  // Для выбора студентов ориентируемся на отдельные чекбоксы в видимых строках,
  // а не на состояние одного общего чекбокса, которое может быть промежуточным.
  function pageSelection(scope) {
    const rows = studentCheckBoxes(scope);
    const keys = rows.map(studentKey);
    if (new Set(keys).size !== keys.length) {
      throw new Error('На странице найдены повторяющиеся идентификаторы студентов — автоматический выбор остановлен.');
    }
    return { rows, keys };
  }
  async function choosePageStudents(scope, chosenKeys, id) {
    let header = nameCheckbox(scope);
    if (!header) throw new Error('Не найдена галочка «Name» в окне Add Students.');
    const {rows, keys} = pageSelection(scope);
    if (!rows.length) throw new Error('На странице нет галочек отдельных студентов.');
    if (rows.some(box => box.checked)) {
      throw new Error('На новой странице уже есть отмеченные студенты. Нужно проверить, не выбирает ли «Name» сразу весь список.');
    }
    // По запросу оператора кликаем Name, чтобы выбрать текущие 10 строк одним действием.
    setCheckbox(header, true, 'Name: выбрать студентов на этой странице');
    await delay(220, id);
    scope = addStudentsDialog() || scope;
    const afterAll = pageSelection(scope);
    if (afterAll.keys.join('|') !== keys.join('|')) {
      throw new Error('Список студентов изменился после клика Name — остановлено.');
    }
    if (afterAll.rows.some(box => !box.checked)) {
      throw new Error('Name не выбрала все видимые строки; возможно, сайт изменил механизм массового выбора.');
    }
    // Последняя страница: снимаем лишние галочки, чтобы итог был РОВНО 20–60.
    for (const key of afterAll.keys) {
      await checkpoint(id);
      if (!chosenKeys.has(key)) {
        // React может перерисовать checkbox после каждого клика; берём актуальную строку.
        const live = pageSelection(addStudentsDialog() || scope);
        const index = live.keys.indexOf(key);
        if (index < 0) throw new Error('Студент исчез из списка до снятия галочки.');
        setCheckbox(live.rows[index], false, `снять лишнего студента ${key.slice(0,40)}`);
        await delay(55, id);
      }
    }
    scope = addStudentsDialog() || scope;
    const verified = pageSelection(scope);
    if (verified.keys.join('|') !== keys.join('|')) {
      throw new Error('Список студентов изменился при уточнении выбора.');
    }
    for (let i = 0; i < verified.rows.length; i++) {
      if (verified.rows[i].checked !== chosenKeys.has(verified.keys[i])) {
        throw new Error('Состояние галочек после выбора не совпадает с планом.');
      }
    }
    return chosenKeys.size;
  }
  async function selectStudents(id) {
    state.phase = 'student_dialog'; save();
    const opener = await waitFor(() => dataButton('add-students', 'button'), id, 14000);
    if (!opener) throw new Error('Не найдена кнопка Add Students.');
    click(opener, 'Add Students');
    let scope = await waitFor(addStudentsDialog, id, 13000);
    if (!scope) throw new Error('Не открылось окно выбора студентов.');
    state.phase = 'select_students';
    state.target = rand(20, 60);
    state.selected = 0; save(); render();
    log(`Для курса выбрано случайное количество: ${state.target} студентов.`);
    const selectedKeys = new Set();
    let pages = 0;
    while (selectedKeys.size < state.target) {
      await checkpoint(id);
      scope = addStudentsDialog() || scope;
      const {rows, keys} = pageSelection(scope);
      if (!rows.length) throw new Error('Не найдены строки студентов — нужен HTML таблицы выбора.');
      if (keys.some(key => selectedKeys.has(key))) {
        throw new Error('Следующая страница повторяет уже выбранных студентов; повторный выбор заблокирован.');
      }
      const remaining = state.target - selectedKeys.size;
      const count = Math.min(remaining, rows.length);
      const chosen = new Set(shuffle(keys).slice(0, count));
      await choosePageStudents(scope, chosen, id);
      for (const key of chosen) selectedKeys.add(key);
      state.selected = selectedKeys.size; save(); render();
      const report = reportedStudentCount(addStudentsDialog() || scope);
      if (report != null && report !== state.selected) {
        throw new Error(`Сайт показывает ${report} выбранных, а скрипт — ${state.selected}. Возможно, Name выбрала все страницы. Add заблокирована.`);
      }
      log(`Страница ${pages+1}: Name → ${count} студентов; всего ${state.selected}/${state.target}.`);
      if (selectedKeys.size === state.target) break;
      if (++pages >= 35) throw new Error('Лимит: 35 страниц студентов.');
      const next = nextStudentPage(scope);
      const before = keys.join('|');
      if (!next) {
        status('manual_students', `Выбрано ${state.selected}/${state.target}. Не найдена стрелка вправо. Перейдите на следующую страницу студентов вручную и нажмите «Продолжить».`);
        log('Не нашёл стрелку вправо; ожидаю ручного перехода.');
        while (state.status === 'manual_students') {
          if (isStopped(id)) throw new Error('STOP');
          await sleep(250);
        }
        const after = pageSelection(addStudentsDialog() || scope).keys.join('|');
        if (before === after) throw new Error('Страница студентов после ручного перехода не изменилась.');
        continue;
      }
      // Автоматически нажимаем именно на кнопку со стрелкой вправо из HTML пользователя.
      click(next, 'Стрелка вправо — следующие 10 студентов');
      const changed = await waitFor(() => {
        const d = addStudentsDialog() || scope;
        const after = pageSelection(d).keys.join('|');
        return after && after !== before;
      }, id, 12000);
      if (!changed) throw new Error('После нажатия стрелки вправо новые 10 студентов не появились.');
      log('Нажата стрелка вправо — открыта следующая страница студентов.');
      await delay(200, id);
    }
    if (state.selected !== state.target) throw new Error('Количество выбранных студентов не совпало с целевым.');
    const report = reportedStudentCount(addStudentsDialog() || scope);
    if (report != null && report !== state.target) {
      throw new Error(`Сайт показывает ${report} студентов вместо ${state.target}; Add заблокирована.`);
    }
    log(`Все страницы обработаны; выбрано ${state.target} студентов.`);
    const add = SA('button[data-cy="add-students"][type="submit"]')
      .find(b => visible(b) && !b.disabled && norm(b.textContent).includes('add'));
    if (!add) throw new Error('Не найдена кнопка Add.');
    click(add, 'Add');
    await delay(600, id);
    const closed = await waitFor(() => !visible(add) || !add.isConnected, id, 9500);
    if (!closed) throw new Error('Окно Add Students не закрылось; Enroll Now не нажата.');
    log('Add нажата — список студентов подготовлен.');
    // В некоторых версиях Tutor LMS появляется промежуточное окно Okay уже после Add.
    // Не подтверждаем сообщения об ошибке.
    const addOkay = okayButton();
    if (addOkay) {
      const m = dialogTextForOkay(addOkay);
      if (isFailureText(m)) throw new Error(`После Add появилась ошибка: ${m.slice(0,220)}`);
      click(addOkay, 'Okay после Add');
      await delay(300, id);
      log('Закрыто промежуточное окно Okay после Add.');
    }
  }
  function isFailureText(text) {
    return /\b(error|failed|failure|unsuccessful|not enrolled|cannot enroll|could not|ошибка|не удалось|не зачислен)\b/i.test(text || '');
  }
  function okayButton() {
    return SA('button[type="button"], button:not([type])').find(b =>
      visible(b) && !ui.root?.contains(b) && !b.disabled && norm(b.textContent) === 'okay');
  }
  function dialogTextForOkay(btn) {
    if (!btn) return '';
    const scope = btn.closest('[role="dialog"], [aria-modal="true"], .tutor-modal, [class*="modal"], [class*="dialog"]');
    return norm((scope || btn.parentElement?.parentElement || btn).textContent).slice(0,2000);
  }
  function successMessage() {
    const alerts = SA('[role="alert"], [role="status"], [class*="toast"], [class*="notice"], [class*="success"], [class*="notification"]')
      .filter(el => visible(el) && !ui.root?.contains(el));
    for (const el of alerts) {
      const t = norm(el.textContent);
      if (isFailureText(t)) return {failed: true, text: t};
      if (/\b(successfully enrolled|enrolled successfully|enrollment successful|successfully added|success|успешно зачислен|зачисление успешно)\b/i.test(t)) {
        if (/\b(0 successful|0 enrolled|zero students)\b/.test(t)) return {failed:true,text:t};
        return {ok:true,text:t};
      }
    }
    return null;
  }
  async function submitEnrollment(id) {
    if (state.test) {
      state.phase = 'test_done';
      status('test_done', `Тест завершён: курс найден, ${state.selected} студентов подготовлено, Add нажата. Enroll Now НЕ нажималась.`);
      log('ТЕСТ: остановлено перед Enroll Now. Зачисления НЕ отправлялись.');
      return false;
    }
    const enroll = await waitFor(() => dataButton('enroll-now'), id, 14000);
    if (!enroll) throw new Error('Не найдена активная кнопка Enroll Now.');
    const previousOkay = okayButton();
    state.phase = 'submitted'; save();
    log('Нажимаю Enroll Now; жду подтверждение сайта и кнопку Okay.');
    click(enroll, 'Enroll Now');
    const outcome = await waitFor(() => {
      const msg = successMessage();
      if (msg?.failed) return msg;
      const okay = okayButton();
      if (okay && okay !== previousOkay) {
        const dialogText = dialogTextForOkay(okay);
        if (isFailureText(dialogText)) return {failed:true,text:dialogText};
        // Окно Okay появилось именно ПОСЛЕ отправки Enroll Now.
        // Это ожидаемое подтверждение сайта. Обрабатываем его автоматически.
        return {ok:true,okay,text:dialogText || 'Confirmation Okay'};
      }
      return msg?.ok ? msg : null;
    }, id, 19000);
    if (outcome?.failed) {
      status('needs_confirmation', `После Enroll Now сайт сообщил об ошибке: ${outcome.text.slice(0,230)}. Проверьте вручную.`);
      log('Ошибка при зачислении — переход к следующему курсу заблокирован.');
      return false;
    }
    if (!outcome?.ok) {
      status('needs_confirmation', 'Не найдено подтверждение после Enroll Now. Проверьте зачисление вручную, чтобы не создать дубль.');
      log('Нет подтверждения зачисления, работа приостановлена.');
      return false;
    }
    const okay = outcome.okay || await waitFor(() => {
      const found = okayButton();
      return found && found !== previousOkay ? found : null;
    }, id, 3300);
    if (okay) {
      const msg = dialogTextForOkay(okay);
      if (isFailureText(msg)) {
        status('needs_confirmation', `Окно Okay содержит ошибку: ${msg.slice(0,230)}`);
        log('Остановка: окно Okay содержит сообщение об ошибке.');
        return false;
      }
      // Save a durable success marker BEFORE clicking Okay: the click may navigate immediately.
      // The next page must advance the index instead of repeating this enrollment.
      state.phase = 'confirmed_success'; save();
      click(okay, 'Okay после зачисления');
      log('После Enroll Now автоматически нажата кнопка Okay.');
      await delay(450, id);
    } else {
      state.phase = 'confirmed_success'; save();
      log('Получено отдельное сообщение об успехе, но кнопка Okay не появилась.');
    }
    log(`Зачисление подтверждено: ${String(outcome.text || '').slice(0,130)}`);
    return true;
  }
  function findEnrollStudentsLink() {
    const links = SA('a[href]').filter(a => visible(a) && !ui.root?.contains(a));
    return links.find(a => {
      try {
        const url = new URL(a.href, location.href);
        return url.origin === location.origin && url.pathname.endsWith('/wp-admin/admin.php') &&
          url.searchParams.get('page') === 'enrollments' &&
          url.searchParams.get('action') === 'add_new' &&
          norm(a.textContent) === 'enroll students';
      } catch { return false; }
    }) || null;
  }
  async function openNextEnrollment() {
    if (state.status !== 'running') return;
    const until = Date.now() + 8000;
    while (Date.now() < until) {
      if (state.status !== 'running') return;
      const link = findEnrollStudentsLink();
      if (link) {
        log('Нажимаю Enroll Students для следующего курса.');
        click(link, 'Enroll Students');
        return; // Новая страница восстановит список из sessionStorage.
      }
      await sleep(250);
    }
    // Резервный вариант для темы Tutor LMS, в которой ссылка отсутствует.
    log('Ссылка Enroll Students на странице не появилась; открываю ту же форму по её адресу.');
    location.assign(BASE);
  }
  function completeAndNext(manual = false) {
    const title = current();
    if (!title) return;
    state.done.push({title, count:state.selected, time:new Date().toISOString(), manual});
    state.index++;
    state.phase = 'idle'; state.target = 0; state.selected = 0;
    state.manualReason = ''; state.status = state.index < state.titles.length ? 'running' : 'done';
    save(); render();
    log(`Готово: ${title}. Обработано ${state.index} из ${state.titles.length}.`);
    if (state.status === 'running') void openNextEnrollment();
  }
  async function run() {
    if (active || state.status !== 'running') return;
    if (state.index >= state.titles.length) { status('done', 'Все курсы обработаны.'); return; }
    if (!isAddNew()) { await openNextEnrollment(); return; }
    const id = ++jobId;
    active = true;
    try {
      log(`${state.test ? 'ТЕСТ' : 'Работа'}: ${state.index + 1}/${state.titles.length} — ${current()}`);
      await searchCourse(current(), id);
      await checkpoint(id);
      await selectStudents(id);
      await checkpoint(id);
      const ok = await submitEnrollment(id);
      if (ok) completeAndNext();
    } catch (e) {
      if (e.message !== 'STOP') {
        // After an irreversible submission, never retry automatically.
        status(state.phase === 'submitted' ? 'needs_confirmation' : 'paused', e.message);
        log(`Остановка: ${e.message}`);
      }
    } finally {
      active = false; render();
    }
  }
  function start(test = false) {
    if (active || !state.titles.length) { log('Сначала загрузите Excel и выберите названия курсов.'); return; }
    if (state.phase === 'submitted' || state.status === 'needs_confirmation') {
      log('Нужно проверить отправленное зачисление, чтобы не создать дубли.'); return;
    }
    // Always reload the clean add-new form before Start/Test.
    state.test = test;
    state.phase = 'idle'; state.target = 0; state.selected = 0;
    state.manualReason = ''; status('running');
    // A test leaves the form staged. Always open a clean form when starting again.
    location.assign(BASE);
  }
  function resume() {
    if (state.status === 'needs_confirmation') return;
    if (state.status === 'manual_students') {
      status('running'); return;
    }
    if (state.status === 'manual_course') {
      log('Выберите конкретный курс кнопкой Select в окне поиска.'); return;
    }
    if (state.status !== 'paused') return;
    if (active) status('running');
    else {
      // A paused interrupted operation is restarted on a clean form.
      state.phase = 'idle'; status('running'); location.assign(BASE);
    }
  }
  function stop() {
    jobId++;
    status('stopped', 'Остановлено оператором. Результаты предыдущих курсов сохранены в этой вкладке.');
    log('STOP: новые клики не выполняются.');
  }
  function reset() {
    if (!confirm('Сбросить список курсов, прогресс и журнал в этой вкладке?')) return;
    jobId++;
    state = initial(); workbook = null; save(); render();
    ui.file.value = ''; ui.sheets.innerHTML = '<option value="0">Лист 1</option>';
    log('Список и прогресс очищены.');
  }
  async function onFile(file) {
    if (!file) return;
    try {
      if (!/\.xlsx$/i.test(file.name)) throw new Error('Сохраните таблицу в формате .xlsx (старый .xls не поддерживается).');
      if (typeof ExcelJS === 'undefined') throw new Error('Библиотека ExcelJS не загрузилась. Проверьте интернет и Tampermonkey.');
      const w = new ExcelJS.Workbook();
      await w.xlsx.load(await file.arrayBuffer());
      if (!w.worksheets.length) throw new Error('В Excel нет листов.');
      workbook = w;
      ui.sheets.innerHTML = '';
      w.worksheets.forEach((sh, i) => {
        const op = document.createElement('option'); op.value = String(i); op.textContent = `${i + 1}. ${sh.name}`;
        ui.sheets.append(op);
      });
      state.fileName = file.name;
      buildTitles();
    } catch (e) { log(`Ошибка загрузки: ${e.message}`); alert(e.message); }
  }
  function columnToNumber(col) {
    const raw = String(col || '').trim().toUpperCase();
    if (/^[1-9]\d*$/.test(raw)) return Number(raw);
    if (!/^[A-Z]{1,3}$/.test(raw)) throw new Error('Укажите столбец: A, B, C ... либо номер.');
    let n = 0;
    for (const c of raw) n = n * 26 + (c.charCodeAt(0) - 64);
    return n;
  }
  function cellToText(cell) {
    const v = cell.value;
    if (v == null) return '';
    if (typeof v === 'string' || typeof v === 'number') return String(v).trim();
    if (v.richText) return v.richText.map(p => p.text || '').join('').trim();
    if (v.text) return String(v.text).trim();
    if (v.result != null) return String(v.result).trim();
    return String(cell.text || '').trim();
  }
  function buildTitles() {
    if (!workbook) return;
    if (active || state.status === 'running') { log('Приостановите работу, прежде чем менять Excel.'); return; }
    try {
      const sheet = workbook.worksheets[Number(ui.sheets.value)];
      if (!sheet) throw new Error('Лист не найден.');
      const col = columnToNumber(ui.column.value);
      const first = Math.max(1, Number(ui.firstRow.value) || 1);
      const names = [];
      const seen = new Set();
      let repeated = 0;
      for (let r = first; r <= sheet.rowCount; r++) {
        const title = cellToText(sheet.getRow(r).getCell(col));
        if (!title) continue;
        if (r === first && /^(course title|course name|title|название курса|курс|название)$/i.test(title)) continue;
        if (ui.dedup.checked && seen.has(norm(title))) { repeated++; continue; }
        seen.add(norm(title)); names.push(title);
      }
      if (!names.length) throw new Error('В этом столбце не найдено названий курсов.');
      jobId++;
      state.titles = names;
      state.index = 0; state.done = [];
      state.status = 'idle'; state.phase = 'idle'; state.test = false;
      state.target = 0; state.selected = 0;
      state.manualReason = '';
      log(`Загружено курсов: ${names.length}${repeated ? ` (пропущено дублей: ${repeated})` : ''}.`);
    } catch (e) { log(`Ошибка списка: ${e.message}`); }
  }
  function render() {
    if (!ui.root) return;
    ui.state.textContent = ({ idle:'Готов к старту', running:'Выполняется', paused:'Пауза',
      manual_course:'Нужен выбор курса', manual_students:'Нужна смена страницы',
      needs_confirmation:'Проверка зачисления', test_done:'Тест завершён',
      done:'Все курсы обработаны', stopped:'Остановлено' })[state.status] || state.status;
    ui.progress.textContent = `${state.index} / ${state.titles.length}`;
    ui.current.textContent = current() || '—';
    ui.students.textContent = state.target ? `${state.selected} / ${state.target}` : '—';
    ui.reason.textContent = state.manualReason || '';
    ui.preview.textContent = state.titles.length ?
      `Файл: ${state.fileName || 'сохранённый список'}\nПервые курсы: ${state.titles.slice(0, 3).join(' | ')}` :
      'Загрузите файл .xlsx. Названия курсов — по одному в строке.';
    ui.logs.textContent = state.log.slice(-14).join('\n');
    ui.logs.scrollTop = ui.logs.scrollHeight;
    ui.start.disabled = active || !state.titles.length || state.status === 'running' || state.status === 'needs_confirmation';
    ui.test.disabled = ui.start.disabled;
    ui.pause.textContent = state.status === 'paused' || state.status === 'manual_students' ? '▶ Продолжить' : 'Ⅱ Пауза';
    ui.pause.disabled = !['running','paused','manual_students'].includes(state.status);
    ui.confirm.hidden = state.status !== 'needs_confirmation';
    ui.skip.hidden = !['paused','manual_course','manual_students','stopped'].includes(state.status);
  }
  function initializePanelPosition(root) {
    const header = root.querySelector('#yc-header');
    const collapse = root.querySelector('#yc-collapse');
    const resetPos = root.querySelector('#yc-reset-pos');
    let data = {};
    try { data = JSON.parse(localStorage.getItem(PANEL_KEY) || '{}'); } catch {}
    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi,v));
    function place(left, top) {
      const box = root.getBoundingClientRect();
      root.style.left = `${Math.round(clamp(left, 0, Math.max(0,innerWidth - box.width)))}px`;
      root.style.top = `${Math.round(clamp(top, 0, Math.max(0,innerHeight - Math.min(box.height,innerHeight - 30))))}px`;
      root.style.right = 'auto'; root.style.bottom = 'auto';
    }
    function persist() {
      const r = root.getBoundingClientRect();
      localStorage.setItem(PANEL_KEY, JSON.stringify({x:r.left,y:r.top,collapsed:root.classList.contains('yc-collapsed'),width:r.width}));
    }
    if (Number.isFinite(data.width)) root.style.width = `${clamp(data.width,290,Math.max(290,innerWidth - 12))}px`;
    if (data.collapsed) {
      root.classList.add('yc-collapsed'); collapse.textContent='+';
      collapse.title='Развернуть панель';
    }
    if (Number.isFinite(data.x) && Number.isFinite(data.y)) place(data.x, data.y);
    collapse.addEventListener('click', event => {
      event.stopPropagation();
      const now = root.classList.toggle('yc-collapsed');
      collapse.textContent = now ? '+' : '−';
      collapse.title = now ? 'Развернуть панель' : 'Свернуть панель';
      const r = root.getBoundingClientRect();
      place(r.left,r.top); persist();
    });
    resetPos.addEventListener('click', event => {
      event.stopPropagation();
      localStorage.removeItem(PANEL_KEY);
      root.style.left='auto'; root.style.top='auto'; root.style.right='16px'; root.style.bottom='16px';
      root.style.width='360px'; root.classList.remove('yc-collapsed');
      collapse.textContent='−';collapse.title='Свернуть панель';
    });
    let drag = null;
    header.addEventListener('pointerdown', event => {
      if (event.button !== 0 || event.target.closest('button')) return;
      const r = root.getBoundingClientRect();
      drag = {id:event.pointerId,x:event.clientX,y:event.clientY,left:r.left,top:r.top};
      header.setPointerCapture(event.pointerId);
      event.preventDefault();
    });
    header.addEventListener('pointermove', event => {
      if (!drag || drag.id !== event.pointerId) return;
      place(drag.left + event.clientX - drag.x, drag.top + event.clientY - drag.y);
    });
    const stopDrag = event => {
      if (!drag || drag.id !== event.pointerId) return;
      drag = null;
      if (header.hasPointerCapture(event.pointerId)) header.releasePointerCapture(event.pointerId);
      persist();
    };
    header.addEventListener('pointerup', stopDrag);
    header.addEventListener('pointercancel', stopDrag);
    // Native resize: persist width and keep the header reachable after viewport changes.
    if (typeof ResizeObserver === 'function') {
      let previousWidth = root.getBoundingClientRect().width;
      new ResizeObserver(entries => {
        const width = entries[0]?.contentRect.width;
        if (width && Math.abs(width - previousWidth) > 1) {
          previousWidth = width;
          if (!drag) persist();
        }
      }).observe(root);
    }
    window.addEventListener('resize', () => {
      const r = root.getBoundingClientRect();
      if (root.style.left) { place(r.left,r.top); persist(); }
    });
  }

  function buildPanel() {
    const style = document.createElement('style');
    style.textContent = `
#yc-enroll-panel {position:fixed;right:16px;bottom:16px;z-index:2147483647;width:360px;min-width:290px;max-width:calc(100vw - 12px);max-height:92vh;overflow:auto;resize:both;background:#111827;color:#e5e7eb;border:1px solid #374151;border-radius:12px;box-shadow:0 18px 60px #0008;font:13px/1.4 Arial,sans-serif;padding:14px;box-sizing:border-box}
#yc-enroll-panel * {box-sizing:border-box}
#yc-enroll-panel .yc-row{display:flex;gap:7px;margin:8px 0;align-items:center;flex-wrap:wrap}
#yc-enroll-panel label {font-size:12px;display:block;color:#d1d5db}
#yc-enroll-panel input:not([type=checkbox]),#yc-enroll-panel select{background:#fff;color:#111827;border:1px solid #9ca3af;border-radius:5px;padding:6px;min-width:0}
#yc-enroll-panel input[type=file]{max-width:100%;font-size:12px}
#yc-enroll-panel button{background:#2563eb;color:white;border:0;border-radius:6px;padding:8px 10px;cursor:pointer;font:bold 12px Arial}
#yc-enroll-panel button:disabled{opacity:.4;cursor:not-allowed}
#yc-enroll-panel button.yc-secondary{background:#4b5563}
#yc-enroll-panel button.yc-stop{background:#b91c1c}
#yc-enroll-panel .yc-note{white-space:pre-wrap;font-size:11px;color:#cbd5e1;margin:7px 0;line-height:1.5;overflow-wrap:anywhere}
#yc-enroll-panel .yc-danger{color:#fbbf24;font-size:12px;white-space:pre-wrap;overflow-wrap:anywhere}
#yc-enroll-panel .yc-stats{display:grid;grid-template-columns:1fr 1fr;gap:5px;background:#1f2937;border-radius:6px;padding:8px;margin:8px 0}
#yc-enroll-panel pre{white-space:pre-wrap;overflow-wrap:anywhere;background:#030712;padding:8px;max-height:145px;overflow:auto;font-size:11px;border-radius:5px}
#yc-enroll-panel h3{margin:0;font-size:15px;color:white;line-height:1.3}
#yc-enroll-panel .yc-header{display:flex;align-items:center;gap:7px;margin:-14px -14px 12px;padding:12px 10px 12px 14px;cursor:grab;background:#1f2937;border-radius:11px 11px 0 0;user-select:none;touch-action:none}
#yc-enroll-panel .yc-header:active{cursor:grabbing}
#yc-enroll-panel .yc-header .yc-grip{color:#94a3b8;font-size:17px;letter-spacing:-2px}
#yc-enroll-panel .yc-header h3{flex:1;min-width:0}
#yc-enroll-panel .yc-header button{padding:5px 8px;background:#374151;font-size:13px;flex:none;touch-action:auto}
#yc-enroll-panel.yc-collapsed{height:auto!important;min-height:0;resize:none;overflow:hidden}
#yc-enroll-panel.yc-collapsed .yc-header{margin-bottom:-14px}
#yc-enroll-panel.yc-collapsed .yc-panel-body{display:none}
#yc-enroll-panel .yc-panel-body{min-height:0}

#yc-enroll-panel .yc-small{width:67px}
`;
    document.head.append(style);
    const root = document.createElement('section');
    root.id = 'yc-enroll-panel';
    root.innerHTML = `
<div class="yc-header" id="yc-header" title="Перетащите панель за этот заголовок">
<span class="yc-grip" aria-hidden="true">⠿</span><h3>📚 YourCourse · Excel</h3>
<button type="button" id="yc-reset-pos" title="Вернуть панель вниз справа">⌖</button>
<button type="button" id="yc-collapse" title="Свернуть или развернуть панель">−</button></div>
<div class="yc-panel-body" id="yc-panel-body">
<label>Excel-файл (.xlsx) <input id="yc-file" type="file" accept=".xlsx"></label>
<div class="yc-row"><label>Лист<br><select id="yc-sheet"><option value="0">Лист 1</option></select></label>
<label>Столбец<br><input class="yc-small" id="yc-column" value="A" title="A, B, C ..."></label>
<label>С строки<br><input class="yc-small" id="yc-row" type="number" min="1" value="1"></label></div>
<label><input id="yc-dedup" type="checkbox" checked> Пропускать одинаковые названия</label>
<div class="yc-row"><button id="yc-read" class="yc-secondary">↻ Прочитать столбец</button><button id="yc-reset" class="yc-secondary">Сброс</button></div>
<div id="yc-preview" class="yc-note"></div>
<div class="yc-stats"><div>Статус: <b id="yc-status"></b></div><div>Готово: <b id="yc-progress"></b></div>
<div style="grid-column:span 2">Текущий курс: <b id="yc-current"></b></div>
<div>Студентов: <b id="yc-students"></b></div></div>
<div class="yc-row"><button id="yc-start">▶ Старт</button><button id="yc-test" class="yc-secondary">🧪 Тест (1 курс)</button>
<button id="yc-pause" class="yc-secondary">Ⅱ Пауза</button><button id="yc-stop" class="yc-stop">■ Стоп</button></div>
<div id="yc-reason" class="yc-danger"></div>
<div class="yc-row"><button id="yc-confirm" class="yc-secondary" hidden>✓ Проверил успех → далее</button>
<button id="yc-skip" class="yc-secondary" hidden>Пропустить курс</button></div>
<label>Журнал</label><pre id="yc-logs"></pre>
<div class="yc-note">ТЕСТ: выбирает курс и студентов, нажимает Add, но НЕ нажимает Enroll Now. Если курс один — Select автоматически. На каждой странице студентов нажимает Name (обычно 10 человек), затем стрелку вправо. При 20–60 выбранных нажимает Add; после реального Enroll Now — Okay → Enroll Students. Панель перемещается за заголовок и сворачивается кнопкой −.</div></div>`;
    document.body.append(root);
    initializePanelPosition(root);
    const get = id => S('#' + id, root);
    ui = { root, file:get('yc-file'), sheets:get('yc-sheet'), column:get('yc-column'), firstRow:get('yc-row'),
      dedup:get('yc-dedup'), preview:get('yc-preview'), state:get('yc-status'), progress:get('yc-progress'),
      current:get('yc-current'), students:get('yc-students'), reason:get('yc-reason'), logs:get('yc-logs'),
      start:get('yc-start'), test:get('yc-test'), pause:get('yc-pause'), confirm:get('yc-confirm'), skip:get('yc-skip') };
    ui.file.addEventListener('change', () => onFile(ui.file.files[0]));
    get('yc-read').addEventListener('click', buildTitles);
    get('yc-reset').addEventListener('click', reset);
    ui.start.addEventListener('click', () => start(false));
    ui.test.addEventListener('click', () => start(true));
    ui.pause.addEventListener('click', () => {
      if (state.status === 'running') status('paused', 'Остановлено перед следующим действием.');
      else resume();
    });
    get('yc-stop').addEventListener('click', stop);
    ui.confirm.addEventListener('click', () => {
      if (state.status === 'needs_confirmation' && confirm('Вы проверили, что студенты действительно зачислены в этот курс?')) completeAndNext(true);
    });
    ui.skip.addEventListener('click', () => {
      if (!confirm(`Пропустить курс «${current()}» без зачисления?`)) return;
      jobId++;
      state.done.push({title:current(), skipped:true,time:new Date().toISOString()});
      state.index++; state.phase = 'idle'; state.status = state.index >= state.titles.length ? 'done' : 'running';
      state.manualReason = ''; state.target = 0; state.selected = 0;
      save(); render();
      if (state.status === 'running') location.assign(BASE);
    });
    // Automatic handoff after the operator clicks the site's Select (not the panel).
    document.addEventListener('click', event => {
      if (state.status === 'manual_course' && event.isTrusted &&
          event.target instanceof Element && event.target.closest('button[data-cy="select-course"]')) {
        courseWasClicked = true;
      }
    }, true);
    render();
  }

  function restore() {
    // Do not automatically retry a submission after navigation / reload.
    if (state.phase === 'submitted' && state.status === 'running') {
      status('needs_confirmation', 'Страница была обновлена после Enroll Now. Проверьте зачисление перед продолжением.');
    }
    if (state.status === 'manual_course' || state.status === 'manual_students') {
      status('paused', 'После перезагрузки ручной выбор нужно повторить на чистой форме.');
    }
    buildPanel();
    if (state.phase === 'confirmed_success' && state.status === 'running') {
      log('После подтверждённого зачисления страница обновилась; продолжаю со следующего курса без повтора.');
      completeAndNext();
      return;
    }
    if (state.status === 'running') run();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', restore, {once:true});
  else restore();
})();
