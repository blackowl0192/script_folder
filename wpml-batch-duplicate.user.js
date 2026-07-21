// ==UserScript==
// @name         WPML Batch Duplicate Products from Excel
// @namespace    local.wpml.batch
// @version      1.0.1
// @description  Скрипт нужен для дублирование Стим игр из RUS в ENG. Загружает Excel с колонками URL и Name, последовательно дублирует English-перевод, нажимает Update и переходит к следующему URL.
// @match        *://*/wp-admin/edit.php?post_type=product&lang=ru*
// @updateURL    https://raw.githubusercontent.com/blackowl0192/script_folder/main/wpml-batch-duplicate.user.js
// @downloadURL  https://raw.githubusercontent.com/blackowl0192/script_folder/main/wpml-batch-duplicate.user.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
    'use strict';

    const KEY = 'wpml_batch_duplicate_state_v1';

    const DEFAULT_STATE = {
        items: [],
        index: 0,
        status: 'idle', // idle | running | paused | waiting_after_update | done | error
        intervalSec: 10,
        actionDelayMs: 1800,
        currentStage: '',
        lastMessage: 'Ожидание загрузки Excel',
        nextAt: 0,
        logs: []
    };

    let state = loadState();
    let processingLock = false;
    let countdownTimer = null;

    function loadState() {
        const saved = GM_getValue(KEY, null);
        return saved ? { ...DEFAULT_STATE, ...saved } : { ...DEFAULT_STATE };
    }

    function saveState(patch = {}) {
        state = { ...state, ...patch };
        GM_setValue(KEY, state);
        render();
    }

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function normalizeUrl(url) {
        try {
            return new URL(String(url).trim(), location.href).href;
        } catch {
            return String(url || '').trim();
        }
    }

    function samePage(a, b) {
        try {
            const ua = new URL(a);
            const ub = new URL(b);
            return ua.origin === ub.origin &&
                   ua.pathname === ub.pathname &&
                   ua.search === ub.search;
        } catch {
            return a === b;
        }
    }

    function currentItem() {
        return state.items[state.index] || null;
    }

    function findEnglishDuplicateCheckbox() {
        const rows = [...document.querySelectorAll('tr')];

        for (const row of rows) {
            const text = (row.innerText || '').replace(/\s+/g, ' ').trim().toLowerCase();
            if (!text.includes('english')) continue;

            const checkbox = row.querySelector(
                'input[type="checkbox"][name="icl_dupes[]"][value="en"], ' +
                'input[type="checkbox"][name="icl_dupes[]"][title="Create duplicate"]'
            );

            if (checkbox) return checkbox;
        }

        // Запасной вариант.
        return document.querySelector('input[type="checkbox"][name="icl_dupes[]"][value="en"]');
    }

    function findDuplicateButton() {
        return document.querySelector('#icl_make_duplicates');
    }

    function findUpdateButton() {
        return document.querySelector(
            'input#publish[type="submit"][name="save"], ' +
            'input#publish[type="submit"][value="Update"], ' +
            'button#publish'
        );
    }


    function addLog(type, message, item = currentItem(), stage = state.currentStage) {
        const entry = {
            time: new Date().toLocaleString(),
            type,
            name: item?.name || '(без Name)',
            url: item?.url || '',
            stage: stage || '',
            message: String(message || '')
        };

        const logs = Array.isArray(state.logs) ? [...state.logs, entry] : [entry];
        state.logs = logs;
        GM_setValue(KEY, state);
        render();
    }

    function goToItem(index, reason = 'Ручное переключение') {
        if (!state.items.length) return;

        const safeIndex = Math.max(0, Math.min(index, state.items.length - 1));
        clearInterval(countdownTimer);

        saveState({
            index: safeIndex,
            status: 'paused',
            currentStage: reason,
            lastMessage: `Перехожу к игре: ${state.items[safeIndex].name || state.items[safeIndex].url}`,
            nextAt: 0
        });

        location.href = state.items[safeIndex].url;
    }

    async function processCurrentPage() {
        if (processingLock || state.status !== 'running') return;

        const item = currentItem();
        if (!item) {
            finishWork();
            return;
        }

        // Если мы еще не на нужной странице — переходим.
        if (!samePage(location.href, item.url)) {
            saveState({
                currentStage: 'Переход к странице',
                lastMessage: `Открываю: ${item.name || item.url}`
            });
            location.href = item.url;
            return;
        }

        processingLock = true;

        try {
            saveState({
                currentStage: '1/3 — Дублирование',
                lastMessage: 'Ищу строку English и ставлю галочку Duplicate'
            });

            const checkbox = await waitForElement(findEnglishDuplicateCheckbox, 15000);
            if (!checkbox) throw new Error('Не найдена галочка Duplicate для English');

            if (!checkbox.checked) {
                checkbox.click();
                checkbox.dispatchEvent(new Event('change', { bubbles: true }));
            }

            await sleep(state.actionDelayMs);

            if (state.status === 'paused') return;

            const duplicateBtn = await waitForElement(findDuplicateButton, 10000);
            if (!duplicateBtn) throw new Error('Не найдена кнопка Duplicate (#icl_make_duplicates)');

            saveState({
                currentStage: '2/3 — Duplicate',
                lastMessage: 'Нажимаю кнопку Duplicate'
            });

            duplicateBtn.click();

            await sleep(state.actionDelayMs);

            if (state.status === 'paused') return;

            const updateBtn = await waitForElement(findUpdateButton, 10000);
            if (!updateBtn) throw new Error('Не найдена кнопка Update (#publish)');

            saveState({
                status: 'waiting_after_update',
                currentStage: '3/3 — Update',
                lastMessage: `Нажимаю Update. Следующий переход через ${state.intervalSec} сек.`,
                nextAt: Date.now() + state.intervalSec * 1000
            });

            addLog('success', 'Duplicate выполнен, нажата кнопка Update', item, 'Update');

            // Сохраняем состояние ДО клика: Update обычно перезагружает страницу.
            updateBtn.click();

            // Если страница не перезагрузилась, таймер все равно продолжит работу.
            startWaitingCountdown();

        } catch (err) {
            console.error('[WPML Batch]', err);
            addLog('error', err.message || String(err), item, state.currentStage);
            saveState({
                status: 'error',
                currentStage: 'Ошибка',
                lastMessage: err.message || String(err)
            });
        } finally {
            processingLock = false;
        }
    }

    function startWaitingCountdown() {
        clearInterval(countdownTimer);

        const tick = () => {
            if (state.status === 'paused') {
                render();
                return;
            }

            if (state.status !== 'waiting_after_update') {
                clearInterval(countdownTimer);
                return;
            }

            const left = Math.max(0, Math.ceil((state.nextAt - Date.now()) / 1000));

            if (left <= 0) {
                clearInterval(countdownTimer);
                goToNextItem();
                return;
            }

            state.lastMessage = `Ожидание перед следующей игрой: ${left} сек.`;
            GM_setValue(KEY, state);
            render();
        };

        tick();
        countdownTimer = setInterval(tick, 500);
    }

    function goToNextItem() {
        const nextIndex = state.index + 1;

        if (nextIndex >= state.items.length) {
            finishWork();
            return;
        }

        saveState({
            index: nextIndex,
            status: 'running',
            currentStage: 'Переход к следующей игре',
            lastMessage: `Открываю ${state.items[nextIndex].name || state.items[nextIndex].url}`,
            nextAt: 0
        });

        location.href = state.items[nextIndex].url;
    }

    function finishWork() {
        saveState({
            status: 'done',
            currentStage: 'Готово',
            lastMessage: `Обработка завершена. Обработано: ${state.items.length}`,
            nextAt: 0
        });
    }

    function waitForElement(getter, timeoutMs) {
        return new Promise(resolve => {
            const started = Date.now();

            const timer = setInterval(() => {
                const el = getter();
                if (el) {
                    clearInterval(timer);
                    resolve(el);
                    return;
                }

                if (Date.now() - started >= timeoutMs) {
                    clearInterval(timer);
                    resolve(null);
                }
            }, 300);
        });
    }

    function parseExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = e => {
                try {
                    const workbook = XLSX.read(e.target.result, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(firstSheet, {
                        defval: '',
                        raw: false
                    });

                    if (!rows.length) {
                        throw new Error('Excel-файл пустой');
                    }

                    const findKey = (row, target) =>
                        Object.keys(row).find(k => String(k).trim().toLowerCase() === target.toLowerCase());

                    const urlKey = findKey(rows[0], 'URL');
                    const nameKey = findKey(rows[0], 'Name');

                    if (!urlKey) {
                        throw new Error('Не найден столбец URL');
                    }

                    const items = rows
                        .map(row => ({
                            url: normalizeUrl(row[urlKey]),
                            name: nameKey ? String(row[nameKey] || '').trim() : ''
                        }))
                        .filter(item => item.url);

                    if (!items.length) {
                        throw new Error('В столбце URL нет ссылок');
                    }

                    resolve(items);
                } catch (err) {
                    reject(err);
                }
            };

            reader.onerror = () => reject(new Error('Не удалось прочитать Excel-файл'));
            reader.readAsArrayBuffer(file);
        });
    }

    // ---------- UI ----------

    const panel = document.createElement('div');
    panel.id = 'wpml-batch-panel';
    panel.innerHTML = `
        <div class="wpmlb-header">
            <strong>WPML Batch Duplicate</strong>
            <button id="wpmlb-collapse" title="Свернуть">—</button>
        </div>

        <div id="wpmlb-body">
            <label class="wpmlb-label">Excel-файл (.xlsx / .xls)</label>
            <input id="wpmlb-file" type="file" accept=".xlsx,.xls" />

            <label class="wpmlb-label">Интервал между ссылками, сек.</label>
            <input id="wpmlb-interval" type="number" min="0" step="1" value="${state.intervalSec}" />

            <div class="wpmlb-buttons">
                <button id="wpmlb-start" class="primary">Старт</button>
                <button id="wpmlb-pause">Пауза</button>
                <button id="wpmlb-resume">Продолжить</button>
                <button id="wpmlb-reset" class="danger">Сброс</button>
            </div>

            <div class="wpmlb-nav">
                <button id="wpmlb-prev">← Предыдущая игра</button>
                <button id="wpmlb-next">Следующая игра →</button>
            </div>

            <div class="wpmlb-status">
                <div><span>Статус:</span> <b id="wpmlb-status-text"></b></div>
                <div><span>Этап:</span> <b id="wpmlb-stage"></b></div>
                <div><span>Игра:</span> <b id="wpmlb-game"></b></div>
                <div><span>Прогресс:</span> <b id="wpmlb-progress"></b></div>
                <div class="wpmlb-message" id="wpmlb-message"></div>
            </div>

            <div class="wpmlb-log-wrap">
                <div class="wpmlb-log-head">
                    <strong>Лог действий и ошибок</strong>
                    <button id="wpmlb-clear-log">Очистить лог</button>
                </div>
                <div id="wpmlb-log"></div>
            </div>
        </div>
    `;

    document.documentElement.appendChild(panel);

    const style = document.createElement('style');
    style.textContent = `
        #wpml-batch-panel {
            position: fixed;
            top: 80px;
            right: 20px;
            z-index: 999999;
            width: 360px;
            min-width: 300px;
            max-width: 600px;
            background: #fff;
            border: 1px solid #b9c0c8;
            border-radius: 10px;
            box-shadow: 0 10px 30px rgba(0,0,0,.22);
            color: #1d2327;
            font-family: Arial, sans-serif;
            font-size: 13px;
            resize: both;
            overflow: auto;
        }

        #wpml-batch-panel .wpmlb-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 12px;
            background: #1d2327;
            color: #fff;
            cursor: move;
            user-select: none;
        }

        #wpml-batch-panel .wpmlb-header button {
            border: 0;
            background: transparent;
            color: #fff;
            font-size: 20px;
            cursor: pointer;
        }

        #wpmlb-body {
            padding: 12px;
        }

        .wpmlb-label {
            display: block;
            margin: 9px 0 5px;
            font-weight: 600;
        }

        #wpmlb-file,
        #wpmlb-interval {
            width: 100%;
            box-sizing: border-box;
        }

        .wpmlb-buttons {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 7px;
            margin-top: 12px;
        }

        .wpmlb-buttons button {
            min-height: 34px;
            cursor: pointer;
            border: 1px solid #8c8f94;
            border-radius: 5px;
            background: #f6f7f7;
        }

        .wpmlb-buttons .primary {
            background: #2271b1;
            border-color: #2271b1;
            color: white;
        }

        .wpmlb-buttons .danger {
            color: #b32d2e;
        }

        .wpmlb-nav {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 7px;
            margin-top: 8px;
        }

        .wpmlb-nav button {
            min-height: 34px;
            cursor: pointer;
            border: 1px solid #8c8f94;
            border-radius: 5px;
            background: #fff;
        }

        .wpmlb-status {
            margin-top: 12px;
            padding: 10px;
            background: #f6f7f7;
            border-radius: 6px;
            line-height: 1.6;
        }

        .wpmlb-status span {
            color: #646970;
        }

        .wpmlb-message {
            margin-top: 8px;
            padding-top: 8px;
            border-top: 1px solid #dcdcde;
            overflow-wrap: anywhere;
        }

        .wpmlb-log-wrap {
            margin-top: 12px;
            border-top: 1px solid #dcdcde;
            padding-top: 10px;
        }

        .wpmlb-log-head {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 8px;
            margin-bottom: 7px;
        }

        .wpmlb-log-head button {
            cursor: pointer;
            border: 1px solid #8c8f94;
            border-radius: 4px;
            background: #f6f7f7;
            font-size: 11px;
            padding: 4px 7px;
        }

        #wpmlb-log {
            max-height: 220px;
            overflow: auto;
            background: #111827;
            color: #e5e7eb;
            border-radius: 6px;
            padding: 8px;
            font-family: Consolas, monospace;
            font-size: 11px;
            line-height: 1.45;
        }

        .wpmlb-log-entry {
            padding: 6px 0;
            border-bottom: 1px solid rgba(255,255,255,.12);
        }

        .wpmlb-log-entry:last-child {
            border-bottom: 0;
        }

        .wpmlb-log-error {
            color: #fca5a5;
        }

        .wpmlb-log-success {
            color: #86efac;
        }
    `;
    document.head.appendChild(style);

    const $ = id => document.getElementById(id);

    function render() {
        const item = currentItem();

        const statusLabels = {
            idle: 'Ожидание',
            running: 'Выполняется',
            paused: 'Пауза',
            waiting_after_update: 'Ожидание',
            done: 'Завершено',
            error: 'Ошибка'
        };

        $('wpmlb-status-text').textContent = statusLabels[state.status] || state.status;
        $('wpmlb-stage').textContent = state.currentStage || '—';
        $('wpmlb-game').textContent = item ? (item.name || '(без Name)') : '—';
        $('wpmlb-progress').textContent = state.items.length
            ? `${Math.min(state.index + 1, state.items.length)} / ${state.items.length}`
            : '0 / 0';
        $('wpmlb-message').textContent = state.lastMessage || '';
        $('wpmlb-interval').value = state.intervalSec;

        $('wpmlb-pause').disabled = !['running', 'waiting_after_update'].includes(state.status);
        $('wpmlb-resume').disabled = state.status !== 'paused';

        $('wpmlb-prev').disabled = !state.items.length || state.index <= 0;
        $('wpmlb-next').disabled = !state.items.length || state.index >= state.items.length - 1;

        const logEl = $('wpmlb-log');
        const logs = Array.isArray(state.logs) ? state.logs : [];

        if (!logs.length) {
            logEl.innerHTML = '<div>Лог пока пуст.</div>';
        } else {
            logEl.innerHTML = logs.slice().reverse().map(entry => {
                const cls = entry.type === 'error' ? 'wpmlb-log-error' : 'wpmlb-log-success';
                const label = entry.type === 'error' ? 'ОШИБКА' : 'OK';
                return `
                    <div class="wpmlb-log-entry ${cls}">
                        <div><b>[${label}] ${entry.time}</b></div>
                        <div>Игра: ${escapeHtml(entry.name)}</div>
                        <div>Этап: ${escapeHtml(entry.stage)}</div>
                        <div>${escapeHtml(entry.message)}</div>
                        ${entry.url ? `<div>${escapeHtml(entry.url)}</div>` : ''}
                    </div>
                `;
            }).join('');
        }
    }

    function escapeHtml(value) {
        return String(value ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    $('wpmlb-file').addEventListener('change', async e => {
        const file = e.target.files?.[0];
        if (!file) return;

        try {
            const items = await parseExcel(file);
            saveState({
                items,
                index: 0,
                status: 'idle',
                currentStage: 'Файл загружен',
                lastMessage: `Загружено строк: ${items.length}`
            });
        } catch (err) {
            saveState({
                status: 'error',
                currentStage: 'Ошибка Excel',
                lastMessage: err.message || String(err)
            });
        }
    });

    $('wpmlb-interval').addEventListener('change', e => {
        const sec = Math.max(0, Number(e.target.value) || 0);
        saveState({ intervalSec: sec });
    });

    $('wpmlb-start').addEventListener('click', () => {
        if (!state.items.length) {
            saveState({
                status: 'error',
                currentStage: 'Нет данных',
                lastMessage: 'Сначала загрузите Excel с колонками URL и Name'
            });
            return;
        }

        const sec = Math.max(0, Number($('wpmlb-interval').value) || 0);
        saveState({
            index: state.status === 'done' ? 0 : state.index,
            status: 'running',
            intervalSec: sec,
            currentStage: 'Запуск',
            lastMessage: 'Запускаю обработку'
        });

        processCurrentPage();
    });

    $('wpmlb-pause').addEventListener('click', () => {
        saveState({
            status: 'paused',
            currentStage: 'Пауза',
            lastMessage: 'Обработка приостановлена'
        });
    });

    $('wpmlb-resume').addEventListener('click', () => {
        const item = currentItem();
        if (!item) return;

        // Если пауза была во время ожидания после Update, продолжаем с новым интервалом.
        // Иначе повторно запускаем сценарий на текущем URL.
        const wasWaiting = state.nextAt > 0;

        if (wasWaiting) {
            saveState({
                status: 'waiting_after_update',
                nextAt: Date.now() + state.intervalSec * 1000,
                currentStage: 'Ожидание после Update',
                lastMessage: `Продолжаю. Следующий переход через ${state.intervalSec} сек.`
            });
            startWaitingCountdown();
        } else {
            saveState({
                status: 'running',
                currentStage: 'Продолжение',
                lastMessage: 'Продолжаю обработку текущей игры'
            });
            processCurrentPage();
        }
    });


    $('wpmlb-prev').addEventListener('click', () => {
        if (state.index > 0) {
            goToItem(state.index - 1, 'Переход к предыдущей игре');
        }
    });

    $('wpmlb-next').addEventListener('click', () => {
        if (state.index < state.items.length - 1) {
            goToItem(state.index + 1, 'Переход к следующей игре');
        }
    });

    $('wpmlb-clear-log').addEventListener('click', () => {
        saveState({ logs: [] });
    });

    $('wpmlb-reset').addEventListener('click', () => {
        clearInterval(countdownTimer);
        GM_deleteValue(KEY);
        state = { ...DEFAULT_STATE };
        $('wpmlb-file').value = '';
        render();
    });

    $('wpmlb-collapse').addEventListener('click', () => {
        const body = $('wpmlb-body');
        const hidden = body.style.display === 'none';
        body.style.display = hidden ? '' : 'none';
        $('wpmlb-collapse').textContent = hidden ? '—' : '+';
    });

    // Перетаскивание панели.
    (function makeDraggable() {
        const header = panel.querySelector('.wpmlb-header');
        let dragging = false;
        let offsetX = 0;
        let offsetY = 0;

        header.addEventListener('mousedown', e => {
            if (e.target.tagName === 'BUTTON') return;
            dragging = true;
            const rect = panel.getBoundingClientRect();
            offsetX = e.clientX - rect.left;
            offsetY = e.clientY - rect.top;
            panel.style.right = 'auto';
        });

        document.addEventListener('mousemove', e => {
            if (!dragging) return;
            panel.style.left = `${Math.max(0, e.clientX - offsetX)}px`;
            panel.style.top = `${Math.max(0, e.clientY - offsetY)}px`;
        });

        document.addEventListener('mouseup', () => {
            dragging = false;
        });
    })();

    render();

    // Восстановление автоматической работы после перехода/Update.
    if (state.status === 'running') {
        setTimeout(processCurrentPage, 800);
    } else if (state.status === 'waiting_after_update') {
        startWaitingCountdown();
    }
})();
