// ==UserScript==
// @name         G2A Парсер
// @namespace    https://g2a.com/
// @version      1.5.1
// @description  Loads an Excel list of G2A links, collects Basic info attributes, and exports a finished Excel report.
// @match        https://www.g2a.com/*
// @match        https://g2a.com/*
// @updateURL    https://raw.githubusercontent.com/blackowl0192/script_folder/main/g2a-parser.user.js
// @downloadURL  https://raw.githubusercontent.com/blackowl0192/script_folder/main/g2a-parser.user.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @grant        GM_addStyle
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

(() => {
    'use strict';

    const STORAGE = {
        items: 'g2a_collector_items',
        results: 'g2a_collector_results',
        currentIndex: 'g2a_collector_current_index',
        running: 'g2a_collector_running',
        delay: 'g2a_collector_delay',
        lastError: 'g2a_collector_last_error',
        currentUrl: 'g2a_collector_current_url',
    };

    const OUTPUT_HEADERS = [
        'game_name',
        'price',
        'categories',
        'developer',
        'publisher',
        'features',
        'platform',
        'region',
        'release_date',
        'restrictions',
        'description_html',
        'source_url',
    ];

    const LABEL_ALIASES = new Map([
        ['genre', 'categories'],
        ['genres', 'categories'],
        ['developer', 'developer'],
        ['developers', 'developer'],
        ['publisher', 'publisher'],
        ['publishers', 'publisher'],
        ['game mode', 'gameMode'],
        ['game modes', 'gameMode'],
        ['player perspective', 'perspective'],
        ['player perspectives', 'perspective'],
        ['theme', 'theme'],
        ['themes', 'theme'],
        ['features', 'featuresRaw'],
        ['platform', 'platform'],
        ['region', 'region'],
        ['release date', 'releaseDate'],
        ['age restrictions', 'restrictions'],
        ['age restriction', 'restrictions'],
        ['age rating', 'restrictions'],
    ]);

    const state = {
        items: GM_getValue(STORAGE.items, []),
        results: GM_getValue(STORAGE.results, []),
        currentIndex: Number(GM_getValue(STORAGE.currentIndex, 0)),
        running: Boolean(GM_getValue(STORAGE.running, false)),
        delay: Number(GM_getValue(STORAGE.delay, 8)),
        lastError: String(GM_getValue(STORAGE.lastError, '')),
        countdownTimer: null,
        collectTimer: null,
    };

    function cleanText(value) {
        return String(value ?? '')
            .replace(/\u00a0/g, ' ')
            .replace(/\s+/g, ' ')
            .trim()
            .replace(/^[\s:|–—-]+|[\s:|–—-]+$/g, '');
    }

    function normalizeLabel(value) {
        return cleanText(value).toLowerCase().replace(/:$/, '');
    }

    function uniqueJoin(values) {
        const seen = new Set();
        const output = [];

        for (const value of values) {
            for (const part of cleanText(value).split(/\s*[;,|]\s*/)) {
                const normalized = part.toLowerCase();
                if (part && !seen.has(normalized)) {
                    seen.add(normalized);
                    output.push(part);
                }
            }
        }

        return output.join('; ');
    }

    function normalizeUrl(url) {
        try {
            const parsed = new URL(url);
            parsed.search = '';
            parsed.hash = '';
            return parsed.href;
        } catch {
            return cleanText(url);
        }
    }

    function saveState() {
        GM_setValue(STORAGE.items, state.items);
        GM_setValue(STORAGE.results, state.results);
        GM_setValue(STORAGE.currentIndex, state.currentIndex);
        GM_setValue(STORAGE.running, state.running);
        GM_setValue(STORAGE.delay, state.delay);
        GM_setValue(STORAGE.lastError, state.lastError);
    }

    function currentItem() {
        return state.items[state.currentIndex] || null;
    }

    function getHostPanel() {
        return document.getElementById('g2a-collector-panel');
    }

    function updatePanel() {
        const panel = getHostPanel();
        if (!panel) return;

        const item = currentItem();
        panel.querySelector('[data-role="progress"]').textContent =
            `${Math.min(state.currentIndex, state.items.length)} / ${state.items.length}`;

        panel.querySelector('[data-role="current"]').textContent =
            item ? item.name : '—';

        panel.querySelector('[data-role="status"]').textContent =
            state.running ? 'Работает' : 'Пауза';

        panel.querySelector('[data-role="error"]').textContent =
            state.lastError || '—';

        const delayInput = panel.querySelector('[data-role="delay"]');
        if (document.activeElement !== delayInput) {
            delayInput.value = String(state.delay);
        }

        panel.querySelector('[data-action="start"]').disabled =
            state.items.length === 0 || state.running;

        panel.querySelector('[data-action="pause"]').disabled =
            !state.running;

        panel.querySelector('[data-action="download"]').disabled =
            state.results.length === 0;
    }

    function enablePanelDrag(panel) {
        const handle = panel.querySelector('.g2a-drag-handle');
        if (!handle) return;

        let dragging = false;
        let startX = 0;
        let startY = 0;
        let startLeft = 0;
        let startTop = 0;

        handle.addEventListener('mousedown', (event) => {
            if (event.button !== 0) return;

            const rect = panel.getBoundingClientRect();
            dragging = true;
            startX = event.clientX;
            startY = event.clientY;
            startLeft = rect.left;
            startTop = rect.top;

            panel.style.left = `${rect.left}px`;
            panel.style.top = `${rect.top}px`;
            panel.style.right = 'auto';
            panel.style.bottom = 'auto';

            document.body.style.userSelect = 'none';
            event.preventDefault();
        });

        document.addEventListener('mousemove', (event) => {
            if (!dragging) return;

            const maxLeft = Math.max(0, window.innerWidth - panel.offsetWidth);
            const maxTop = Math.max(0, window.innerHeight - 40);

            const left = Math.min(maxLeft, Math.max(0, startLeft + event.clientX - startX));
            const top = Math.min(maxTop, Math.max(0, startTop + event.clientY - startY));

            panel.style.left = `${left}px`;
            panel.style.top = `${top}px`;
        });

        document.addEventListener('mouseup', () => {
            if (!dragging) return;
            dragging = false;
            document.body.style.userSelect = '';
        });
    }

    function addPanel() {
        if (getHostPanel()) return;

        const panel = document.createElement('div');
        panel.id = 'g2a-collector-panel';
        panel.innerHTML = `
            <div class="g2a-title g2a-drag-handle" title="Перетащите панель">G2A Парсер <span class="g2a-drag-mark">⋮⋮</span></div>

            <label class="g2a-file-row">
                <span>Загрузить Excel</span>
                <input data-role="file" type="file" accept=".xlsx,.xls,.csv">
            </label>

            <div class="g2a-grid">
                <div>Обработано</div><strong data-role="progress">0 / 0</strong>
                <div>Текущая игра</div><strong data-role="current">—</strong>
                <div>Статус</div><strong data-role="status">Paused</strong>
                <div>Переход через</div>
                <div class="g2a-delay-wrap">
                    <input data-role="delay" type="number" min="2" max="600" step="1">
                    <span>сек.</span>
                </div>
                <div>До перехода</div><strong data-role="countdown">—</strong>
                <div>Последняя ошибка</div><strong data-role="error">—</strong>
            </div>

            <div class="g2a-buttons">
                <button data-action="start">Начать</button>
                <button data-action="pause">Пауза</button>
                <button data-action="retry">Повторить текущую</button>
                <button data-action="skip">Пропустить</button>
                <button data-action="download">Скачать Excel</button>
                <button data-action="clear" class="danger">Очистить прогресс</button>
            </div>
        `;

        document.documentElement.appendChild(panel);
        enablePanelDrag(panel);

        panel.querySelector('[data-role="file"]').addEventListener('change', handleFileUpload);
        panel.querySelector('[data-role="delay"]').addEventListener('change', handleDelayChange);
        panel.querySelector('[data-action="start"]').addEventListener('click', startCollector);
        panel.querySelector('[data-action="pause"]').addEventListener('click', pauseCollector);
        panel.querySelector('[data-action="retry"]').addEventListener('click', retryCurrent);
        panel.querySelector('[data-action="skip"]').addEventListener('click', skipCurrent);
        panel.querySelector('[data-action="download"]').addEventListener('click', downloadReport);
        panel.querySelector('[data-action="clear"]').addEventListener('click', clearProgress);

        updatePanel();
    }

    async function handleFileUpload(event) {
        const file = event.target.files?.[0];
        if (!file) return;

        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet, {
                defval: '',
                raw: false,
            });

            const parsed = rows
                .map((row) => {
                    const normalized = {};
                    for (const [key, value] of Object.entries(row)) {
                        normalized[normalizeLabel(key)] = value;
                    }

                    const name = cleanText(
                        normalized.name ||
                        normalized['game name'] ||
                        normalized.game_name
                    );

                    const link = normalizeUrl(
                        normalized.link ||
                        normalized.url ||
                        normalized.source_url ||
                        normalized['source url']
                    );

                    return {
                        name,
                        price: cleanText(normalized.price),
                        link,
                    };
                })
                .filter((item) => item.name && item.link);

            if (!parsed.length) {
                throw new Error('No valid rows found. Required columns: Name and link.');
            }

            state.items = parsed;
            state.results = [];
            state.currentIndex = 0;
            state.running = false;
            state.lastError = '';
            saveState();
            updatePanel();
            alert(`Loaded ${parsed.length} games.`);
        } catch (error) {
            state.lastError = error.message || String(error);
            saveState();
            updatePanel();
        }
    }

    function handleDelayChange(event) {
        const value = Number(event.target.value);
        state.delay = Number.isFinite(value)
            ? Math.min(600, Math.max(2, Math.round(value)))
            : 8;

        saveState();
        updatePanel();
    }

    function startCollector() {
        if (!state.items.length) {
            alert('Load the Excel file first.');
            return;
        }

        if (state.currentIndex >= state.items.length) {
            alert('All rows are already processed.');
            return;
        }

        state.running = true;
        state.lastError = '';
        saveState();
        updatePanel();
        processCurrentPage();
    }

    function pauseCollector() {
        state.running = false;
        clearTimers();
        saveState();
        updatePanel();
    }

    function retryCurrent() {
        clearTimers();
        state.running = true;
        state.lastError = '';
        saveState();
        updatePanel();

        const item = currentItem();
        if (!item) return;

        if (normalizeUrl(location.href) !== item.link) {
            location.href = item.link;
            return;
        }

        processCurrentPage();
    }

    function skipCurrent() {
        clearTimers();

        const item = currentItem();
        if (item) {
            state.results.push({
                game_name: normalizeGameName(item.name),
                price: item.price || '',
                categories: '',
                developer: '',
                publisher: '',
                features: '',
                platform: inferPlatform(item.link),
                region: inferRegion(item.link),
                release_date: '',
                restrictions: '',
                description_html: '',
                source_url: item.link,
                status: 'skipped',
            });

            state.currentIndex += 1;
        }

        saveState();
        updatePanel();

        if (state.running) {
            scheduleNextNavigation();
        }
    }

    function clearProgress() {
        if (!confirm('Delete the imported list and all collected results?')) return;

        clearTimers();

        for (const key of Object.values(STORAGE)) {
            GM_deleteValue(key);
        }

        state.items = [];
        state.results = [];
        state.currentIndex = 0;
        state.running = false;
        state.delay = 8;
        state.lastError = '';
        updatePanel();
    }

    function clearTimers() {
        if (state.countdownTimer) {
            clearInterval(state.countdownTimer);
            state.countdownTimer = null;
        }

        if (state.collectTimer) {
            clearTimeout(state.collectTimer);
            state.collectTimer = null;
        }

        const panel = getHostPanel();
        panel?.querySelector('[data-role="countdown"]')?.replaceChildren('—');
    }

    function isCurrentPageExpected() {
        const item = currentItem();
        if (!item) return false;
        return normalizeUrl(location.href) === item.link;
    }

    function processCurrentPage() {
        if (!state.running) return;

        const item = currentItem();

        if (!item) {
            finishCollector();
            return;
        }

        if (!isCurrentPageExpected()) {
            GM_setValue(STORAGE.currentUrl, item.link);
            location.href = item.link;
            return;
        }

        waitForPageData()
            .then(() => collectCurrentProduct(item))
            .then((result) => {
                const existingIndex = state.results.findIndex(
                    (entry) => entry.source_url === result.source_url
                );

                if (existingIndex >= 0) {
                    state.results[existingIndex] = result;
                } else {
                    state.results.push(result);
                }

                state.currentIndex += 1;
                state.lastError = '';
                saveState();
                updatePanel();

                if (state.currentIndex >= state.items.length) {
                    finishCollector();
                } else {
                    scheduleNextNavigation();
                }
            })
            .catch((error) => {
                state.lastError = error.message || String(error);
                saveState();
                updatePanel();
                state.running = false;
                saveState();
                alert(
                    'The page could not be processed. Wait until the product page is fully loaded, then press Retry current. If G2A shows a verification page, complete it manually first.'
                );
            });
    }

    async function waitForPageData() {
        const timeoutMs = 60000;
        const started = Date.now();

        while (Date.now() - started < timeoutMs) {
            const trigger = document.querySelector(
                '#accordion-trigger-basic-info, button[aria-controls="accordion-content-basic-info"]'
            );

            const content = document.querySelector(
                '#accordion-content-basic-info, [id="accordion-content-basic-info"]'
            );

            const bodyText = cleanText(document.body?.innerText).toLowerCase();

            if (
                trigger ||
                content ||
                bodyText.includes('basic info') ||
                bodyText.includes('product information')
            ) {
                await sleep(1800);
                return;
            }

            await sleep(500);
        }

        throw new Error(
            'The G2A Basic info accordion was not found within 60 seconds.'
        );
    }

    function sleep(ms) {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }

    function scheduleNextNavigation() {
        clearTimers();

        const next = currentItem();
        if (!next) {
            finishCollector();
            return;
        }

        let remaining = state.delay;
        const countdown = getHostPanel()?.querySelector('[data-role="countdown"]');

        if (countdown) countdown.textContent = `${remaining} сек.`;

        state.countdownTimer = setInterval(() => {
            remaining -= 1;

            if (countdown) countdown.textContent = `${Math.max(0, remaining)} сек.`;

            if (remaining <= 0) {
                clearTimers();

                if (!state.running) return;

                GM_setValue(STORAGE.currentUrl, next.link);
                location.href = next.link;
            }
        }, 1000);
    }

    function finishCollector() {
        clearTimers();
        state.running = false;
        saveState();
        updatePanel();
        alert('All games are processed. Use Download Excel.');
    }

    function findBasicInfoRoot() {
        const exactContent = document.querySelector(
            '#accordion-content-basic-info, [id="accordion-content-basic-info"]'
        );

        if (exactContent) {
            return exactContent;
        }

        const trigger = document.querySelector(
            '#accordion-trigger-basic-info, button[aria-controls="accordion-content-basic-info"]'
        );

        if (trigger) {
            const controlledId = trigger.getAttribute('aria-controls');

            if (controlledId) {
                const controlled = document.getElementById(controlledId);
                if (controlled) return controlled;
            }

            const accordionItem =
                trigger.closest('[data-state], section, article, li, div');

            if (accordionItem) {
                return accordionItem;
            }
        }

        const acceptedHeadings = new Set([
            'basic info',
            'basic information',
            'product information',
            'product details',
            'details',
        ]);

        const elements = [...document.querySelectorAll(
            'h1,h2,h3,h4,h5,h6,button,[role="button"],div,span,p'
        )];

        const heading = elements.find((element) => {
            const label = normalizeLabel(element.innerText);
            return acceptedHeadings.has(label);
        });

        if (heading) {
            let node = heading;

            for (let depth = 0; depth < 9 && node; depth += 1) {
                const nodeText = cleanText(node.innerText);

                if (nodeText.length >= 50 && nodeText.length <= 20000) {
                    return node;
                }

                node = node.parentElement;
            }
        }

        return document.querySelector('main') || document.body;
    }

    function clickBasicInfoIfNeeded() {
        const exactTrigger = document.querySelector(
            '#accordion-trigger-basic-info, button[aria-controls="accordion-content-basic-info"]'
        );

        if (exactTrigger) {
            try {
                exactTrigger.scrollIntoView({
                    block: 'center',
                    behavior: 'instant',
                });

                if (exactTrigger.getAttribute('aria-expanded') !== 'true') {
                    exactTrigger.click();
                }

                return;
            } catch {
                // Continue with the fallback selectors below.
            }
        }

        const acceptedHeadings = new Set([
            'basic info',
            'basic information',
            'product information',
            'product details',
        ]);

        const candidates = [...document.querySelectorAll(
            'button,[role="button"],summary,h2,h3,h4,div'
        )];

        const target = candidates.find((element) =>
            acceptedHeadings.has(normalizeLabel(element.innerText))
        );

        if (!target) return;

        try {
            target.scrollIntoView({
                block: 'center',
                behavior: 'instant',
            });

            if (target.getAttribute('aria-expanded') !== 'true') {
                target.click();
            }
        } catch {
            // The section may already be expanded or not clickable.
        }
    }

    function parseDomPairs(root) {
        const result = {};

        const elements = [...root.querySelectorAll(
            'dt,dd,th,td,strong,b,span,div,p'
        )];

        for (const element of elements) {
            const label = normalizeLabel(element.innerText);
            const mapped = LABEL_ALIASES.get(label);

            if (!mapped) continue;

            const candidates = [
                element.nextElementSibling,
                element.parentElement?.querySelector('dd'),
                element.parentElement?.children?.[1],
            ].filter(Boolean);

            for (const candidate of candidates) {
                const value = cleanText(candidate.innerText);

                if (
                    value &&
                    normalizeLabel(value) !== label &&
                    value.length < 2000
                ) {
                    result[mapped] ||= value;
                    break;
                }
            }
        }

        return result;
    }

    function parseTextPairs(root) {
        const lines = root.innerText
            .split(/\r?\n/)
            .map(cleanText)
            .filter(Boolean);

        const result = {};
        const sectionHeadings = new Set([
            'basic info',
            'basic information',
            'product information',
            'product details',
        ]);

        const headingIndex = lines.findIndex((line) =>
            sectionHeadings.has(normalizeLabel(line))
        );

        const startIndex = headingIndex >= 0 ? headingIndex + 1 : 0;

        let index = startIndex;

        while (index < lines.length) {
            const line = lines[index];

            const inline = line.match(/^([^:]{2,40}):\s*(.+)$/);
            if (inline) {
                const mapped = LABEL_ALIASES.get(normalizeLabel(inline[1]));
                if (mapped) {
                    result[mapped] ||= cleanText(inline[2]);
                    index += 1;
                    continue;
                }
            }

            const mapped = LABEL_ALIASES.get(normalizeLabel(line));

            if (!mapped) {
                index += 1;
                continue;
            }

            const values = [];
            let next = index + 1;

            while (next < lines.length) {
                const nextLabel = normalizeLabel(lines[next]);
                if (LABEL_ALIASES.has(nextLabel)) break;
                if (/^(reviews|offers|product description|system requirements)$/i.test(nextLabel)) break;

                values.push(lines[next]);
                next += 1;
            }

            if (values.length) {
                result[mapped] ||= uniqueJoin(values);
            }

            index = Math.max(next, index + 1);
        }

        return result;
    }

    function extractStructuredData() {
        const result = {};

        const scripts = [
            ...document.querySelectorAll(
                'script[type="application/ld+json"], script#__NEXT_DATA__, script[type="application/json"]'
            ),
        ];

        const visit = (value, depth = 0) => {
            if (depth > 10 || value == null) return;

            if (Array.isArray(value)) {
                for (const entry of value) visit(entry, depth + 1);
                return;
            }

            if (typeof value !== 'object') return;

            for (const [rawKey, rawValue] of Object.entries(value)) {
                const key = normalizeLabel(
                    rawKey.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/_/g, ' ')
                );

                const mapped = LABEL_ALIASES.get(key);

                if (
                    mapped &&
                    (typeof rawValue === 'string' || typeof rawValue === 'number')
                ) {
                    result[mapped] ||= cleanText(rawValue);
                }

                if (
                    key === 'name' &&
                    typeof rawValue === 'string' &&
                    !result.structuredName
                ) {
                    result.structuredName = cleanText(rawValue);
                }

                visit(rawValue, depth + 1);
            }
        };

        for (const script of scripts) {
            try {
                const parsed = JSON.parse(script.textContent || '');
                visit(parsed);
            } catch {
                // Ignore scripts that do not contain valid JSON.
            }
        }

        return result;
    }

    function inferPlatform(url) {
        const slug = url.toLowerCase();
        return slug.includes('-pc-') ? 'PC' : '';
    }

    function inferRegion(url) {
        const slug = url.toLowerCase();

        if (slug.includes('-global-')) return 'Global';
        if (slug.includes('-europe-') || slug.includes('-eu-')) return 'Europe';
        if (slug.includes('-row-')) return 'ROW';
        if (slug.includes('-us-')) return 'United States';
        if (slug.includes('-uk-')) return 'United Kingdom';

        return '';
    }

    function makeAbsoluteUrl(value) {
        const url = cleanText(value);
        if (!url) return '';

        try {
            return new URL(url, location.href).href;
        } catch {
            return url;
        }
    }

    function sanitizeDescriptionHtml(article) {
        if (!article) return '';

        const clone = article.cloneNode(true);

        // Remove executable / non-content elements.
        clone.querySelectorAll(
            'script, style, iframe, object, embed, form, button, input, textarea, select'
        ).forEach((node) => node.remove());

        // Keep paragraph/list/heading/image structure, but remove JS event handlers.
        clone.querySelectorAll('*').forEach((element) => {
            for (const attr of [...element.attributes]) {
                if (/^on/i.test(attr.name)) {
                    element.removeAttribute(attr.name);
                }
            }

            if (element.tagName === 'IMG') {
                const src =
                    element.getAttribute('src') ||
                    element.getAttribute('data-src') ||
                    element.getAttribute('data-lazy-src');

                if (src) {
                    element.setAttribute('src', makeAbsoluteUrl(src));
                }

                // Remove srcset to avoid importing responsive variants.
                element.removeAttribute('srcset');
                element.removeAttribute('data-src');
                element.removeAttribute('data-lazy-src');
                element.removeAttribute('loading');
            }

            if (element.tagName === 'A') {
                const href = element.getAttribute('href');
                if (href) {
                    element.setAttribute('href', makeAbsoluteUrl(href));
                }
            }

            // Remove site-specific visual classes/ids; keep semantic HTML only.
            element.removeAttribute('class');
            element.removeAttribute('id');
            element.removeAttribute('style');
            element.removeAttribute('data-testid');
            element.removeAttribute('data-open');
            element.removeAttribute('aria-hidden');
        });

        return clone.innerHTML.trim();
    }

    function extractDescriptionHtml() {
        const trigger = document.querySelector(
            '#accordion-trigger-description, button[aria-controls="accordion-content-description"]'
        );

        if (trigger && trigger.getAttribute('aria-expanded') !== 'true') {
            try {
                trigger.scrollIntoView({
                    block: 'center',
                    behavior: 'instant',
                });
                trigger.click();
            } catch {
                // Continue; content may already be present in DOM.
            }
        }

        const content = document.querySelector(
            '#accordion-content-description'
        );

        const article =
            content?.querySelector(
                'article[data-testid="description-content"], article.prose'
            ) ||
            document.querySelector(
                'article[data-testid="description-content"], [data-testid="description-container"] article'
            );

        return sanitizeDescriptionHtml(article);
    }

    function normalizeGameName(value) {
        let name = cleanText(value);

        // Main requested rule:
        // ARC Raiders (PC) - Steam Account - GLOBAL -> ARC Raiders
        const pcIndex = name.search(/\s*\(PC\)/i);
        if (pcIndex >= 0) {
            name = name.slice(0, pcIndex);
        }

        return cleanText(name);
    }

    function normalizeDate(value) {
        const text = cleanText(value);
        if (!text) return '';

        // YYYY-MM-DD -> DD.MM.YYYY
        const iso = text.match(/\b(\d{4})-(\d{2})-(\d{2})\b/);
        if (iso) return `${iso[3]}.${iso[2]}.${iso[1]}`;

        // DD.MM.YYYY / DD/MM/YYYY -> DD.MM.YYYY
        const numeric = text.match(/\b(\d{1,2})[./](\d{1,2})[./](\d{4})\b/);
        if (numeric) {
            return `${numeric[1].padStart(2, '0')}.${numeric[2].padStart(2, '0')}.${numeric[3]}`;
        }

        // October 30, 2025 -> 30.10.2025
        const months = {
            january: '01',
            february: '02',
            march: '03',
            april: '04',
            may: '05',
            june: '06',
            july: '07',
            august: '08',
            september: '09',
            october: '10',
            november: '11',
            december: '12',
        };

        const englishDate = text.match(
            /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})\b/i
        );

        if (englishDate) {
            const month = months[englishDate[1].toLowerCase()];
            const day = englishDate[2].padStart(2, '0');
            const year = englishDate[3];
            return `${day}.${month}.${year}`;
        }

        return text;
    }

    function getPageTitle() {
        const heading = document.querySelector('h1');
        return cleanText(heading?.innerText) || cleanText(document.title.split('|')[0]);
    }

    async function collectCurrentProduct(item) {
        clickBasicInfoIfNeeded();

        const trigger = document.querySelector(
            '#accordion-trigger-basic-info, button[aria-controls="accordion-content-basic-info"]'
        );

        if (trigger) {
            const started = Date.now();

            while (
                trigger.getAttribute('aria-expanded') !== 'true' &&
                Date.now() - started < 5000
            ) {
                await sleep(250);
            }
        }

        await sleep(700);

        const root = findBasicInfoRoot();
        const domData = root ? parseDomPairs(root) : {};
        const textData = root ? parseTextPairs(root) : {};
        const structuredData = extractStructuredData();

        // DOM values have priority, followed by visible text and structured JSON.
        const data = {
            ...structuredData,
            ...textData,
            ...domData,
        };

        const descriptionHtml = extractDescriptionHtml();
        await sleep(500);

        const result = {
            game_name: normalizeGameName(
                getPageTitle() ||
                cleanText(structuredData.structuredName) ||
                item.name
            ),
            price: item.price || '',
            categories: cleanText(data.categories),
            developer: cleanText(data.developer),
            publisher: cleanText(data.publisher),
            features: cleanText(data.gameMode),
            platform: cleanText(data.platform) || inferPlatform(item.link),
            region: cleanText(data.region) || inferRegion(item.link),
            release_date: normalizeDate(data.releaseDate),
            restrictions: cleanText(data.restrictions),
            description_html: descriptionHtml,
            source_url: item.link,
            status: 'collected',
        };

        const meaningfulFields = [
            result.categories,
            result.developer,
            result.publisher,
            result.features,
            result.release_date,
            result.restrictions,
        ].filter(Boolean).length;

        // Do not stop the entire queue when G2A omits some fields.
        // The row is saved as partial and can be reviewed later.
        if (meaningfulFields < 2) {
            result.status = 'partial';
            state.lastError =
                'Only part of the attributes was detected on this page. The row was saved.';
        }

        return result;
    }

    function downloadReport() {
        if (!state.results.length) {
            alert('No collected rows yet.');
            return;
        }

        const rows = state.results.map((result) => {
            const row = {};
            for (const header of OUTPUT_HEADERS) {
                row[header] = result[header] || '';
            }
            return row;
        });

        const sheet = XLSX.utils.json_to_sheet(rows, {
            header: OUTPUT_HEADERS,
        });

        sheet['!cols'] = [
            { wch: 38 }, // game_name
            { wch: 12 }, // price
            { wch: 32 }, // categories
            { wch: 28 }, // developer
            { wch: 28 }, // publisher
            { wch: 32 }, // features = Game Mode only
            { wch: 14 }, // platform
            { wch: 16 }, // region
            { wch: 18 }, // release_date
            { wch: 28 }, // restrictions
            { wch: 90 }, // description_html
            { wch: 70 }, // source_url
        ];

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, sheet, 'Games filled');
        XLSX.writeFile(workbook, 'games_attributes_all.xlsx');
    }

    GM_addStyle(`
        #g2a-collector-panel {
            position: fixed;
            z-index: 2147483647;
            top: 16px;
            right: 16px;
            width: 360px;
            min-width: 320px;
            min-height: 260px;
            max-width: calc(100vw - 20px);
            max-height: calc(100vh - 20px);
            resize: both;
            overflow: auto;
            box-sizing: border-box;
            padding: 14px;
            border: 1px solid #273448;
            border-radius: 12px;
            background: #111827;
            color: #f8fafc;
            box-shadow: 0 12px 32px rgba(0, 0, 0, .32);
            font-family: Arial, sans-serif;
            font-size: 13px;
            line-height: 1.4;
        }

        #g2a-collector-panel * {
            box-sizing: border-box;
        }

        #g2a-collector-panel .g2a-title {
            margin-bottom: 12px;
            font-size: 16px;
            font-weight: 700;
        }

        #g2a-collector-panel .g2a-drag-handle {
            display: flex;
            align-items: center;
            justify-content: space-between;
            cursor: move;
            user-select: none;
        }

        #g2a-collector-panel .g2a-drag-mark {
            font-size: 18px;
            opacity: .65;
            letter-spacing: -3px;
        }

        #g2a-collector-panel .g2a-file-row {
            display: grid;
            gap: 6px;
            margin-bottom: 12px;
        }

        #g2a-collector-panel input {
            width: 100%;
            min-width: 0;
            padding: 7px 8px;
            border: 1px solid #475569;
            border-radius: 7px;
            background: #0f172a;
            color: #f8fafc;
        }

        #g2a-collector-panel .g2a-grid {
            display: grid;
            grid-template-columns: 120px minmax(0, 1fr);
            gap: 7px 10px;
            align-items: start;
            margin-bottom: 12px;
        }

        #g2a-collector-panel .g2a-grid strong {
            overflow-wrap: anywhere;
            font-weight: 600;
        }

        #g2a-collector-panel .g2a-delay-wrap {
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 6px;
            align-items: center;
        }

        #g2a-collector-panel .g2a-buttons {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
        }

        #g2a-collector-panel button {
            min-height: 34px;
            padding: 7px 9px;
            border: 0;
            border-radius: 7px;
            background: #2563eb;
            color: #fff;
            cursor: pointer;
            font-weight: 600;
        }

        #g2a-collector-panel button:hover:not(:disabled) {
            filter: brightness(1.08);
        }

        #g2a-collector-panel button:disabled {
            cursor: not-allowed;
            opacity: .45;
        }

        #g2a-collector-panel button.danger {
            background: #b91c1c;
        }
    `);

    addPanel();

    if (state.running) {
        window.addEventListener('load', () => {
            state.collectTimer = setTimeout(processCurrentPage, 1500);
        });
    }
})();
