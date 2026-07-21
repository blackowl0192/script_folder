// ==UserScript==
// @name         Prime2Key Excel Product Importer
// @namespace    https://prime2key.com/
// @version      2.8.2
// @description  Импорт товаров из Excel в WordPress/WooCommerce, отдельная панель и автоматическое добавление/раскрытие атрибутов.
// @match        https://prime2key.com/wp-admin/*
// @updateURL    https://raw.githubusercontent.com/blackowl0192/script_folder/main/prime2key-product-importer.user.js
// @downloadURL  https://raw.githubusercontent.com/blackowl0192/script_folder/main/prime2key-product-importer.user.js
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const STORAGE_KEY = 'prime2key_excel_importer_v2';

    // Структура файла games_attributes_all.xlsx:
    // game_name, price, categories, developer, publisher,
    // features, region, release_date, description_html, source_url
    const ATTRIBUTES = [
        { label: 'Разработчик', taxonomy: 'pa_developer', key: 'developer' },
        { label: 'Особенности', taxonomy: 'pa_features', key: 'features' },
        { label: 'Язык', taxonomy: 'pa_language', key: '__language', defaultValue: 'Английский' },
        { label: 'Платформа', taxonomy: 'pa_platform', key: '__platform', defaultValue: 'PC' },
        { label: 'Издатель', taxonomy: 'pa_publisher', key: 'publisher' },
        { label: 'Регион', taxonomy: 'pa_region', key: 'region' },
        { label: 'Дата выпуска', taxonomy: 'pa_release-date', key: 'release_date' },
        { label: 'Возрастной рейтинг', taxonomy: 'pa_restrictions', key: '__restrictions', defaultValue: 'Без ограничений' }
    ];

    let state = loadState();
    let panelWindow = null;

    function loadState() {
        try {
            return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {
                rows: [],
                currentIndex: 0,
                fileName: ''
            };
        } catch {
            return { rows: [], currentIndex: 0, fileName: '' };
        }
    }

    function saveState() {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    }


    function normalizeRow(row) {
        return {
            game_name: String(row.game_name ?? '').trim(),
            price: String(row.price ?? '').trim(),
            categories: String(row.categories ?? '').trim(),
            developer: String(row.developer ?? '').trim(),
            publisher: String(row.publisher ?? '').trim(),
            features: String(row.features ?? '').trim(),
            region: String(row.region ?? '').trim(),
            release_date: String(row.release_date ?? '').trim(),
            description_html: String(row.description_html ?? '').trim(),
            source_url: String(row.source_url ?? '').trim()
        };
    }

    function normalizePrice(value) {
        return String(value ?? '')
            .replace(/\s/g, '')
            .replace(/[€$£]/g, '')
            .replace(',', '.')
            .trim();
    }

    function escapeHtml(value) {
        return String(value ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    async function copyText(text, button) {
        const value = String(text ?? '');

        try {
            // Кнопка находится в отдельном popup-окне.
            // Используем document именно этого окна, а не основной вкладки.
            const targetDocument =
                button?.ownerDocument ||
                getPanelDocument() ||
                document;

            const targetWindow =
                targetDocument.defaultView ||
                panelWindow ||
                window;

            let copied = false;

            // Вариант 1: современный Clipboard API.
            try {
                if (
                    targetWindow.navigator?.clipboard &&
                    targetWindow.isSecureContext
                ) {
                    await targetWindow.navigator.clipboard.writeText(value);
                    copied = true;
                }
            } catch (clipboardError) {
                console.warn(
                    '[Prime2Key Importer] Clipboard API недоступен:',
                    clipboardError
                );
            }

            // Вариант 2: execCommand в том же popup-окне.
            if (!copied) {
                const textarea =
                    targetDocument.createElement('textarea');

                textarea.value = value;
                textarea.setAttribute('readonly', '');

                textarea.style.position = 'fixed';
                textarea.style.top = '0';
                textarea.style.left = '0';
                textarea.style.width = '2px';
                textarea.style.height = '2px';
                textarea.style.opacity = '0.01';
                textarea.style.pointerEvents = 'none';

                targetDocument.body.appendChild(textarea);

                textarea.focus();
                textarea.select();
                textarea.setSelectionRange(0, textarea.value.length);

                copied =
                    targetDocument.execCommand &&
                    targetDocument.execCommand('copy');

                textarea.remove();
            }

            // Вариант 3: временный input в основной вкладке.
            if (!copied) {
                const textarea =
                    document.createElement('textarea');

                textarea.value = value;
                textarea.setAttribute('readonly', '');
                textarea.style.position = 'fixed';
                textarea.style.left = '-9999px';

                document.body.appendChild(textarea);

                window.focus();
                textarea.focus();
                textarea.select();
                textarea.setSelectionRange(0, textarea.value.length);

                copied =
                    document.execCommand &&
                    document.execCommand('copy');

                textarea.remove();

                if (panelWindow && !panelWindow.closed) {
                    panelWindow.focus();
                }
            }

            if (!copied) {
                throw new Error(
                    'Браузер не разрешил доступ к буферу обмена.'
                );
            }

            if (button) {
                const oldText = button.textContent;

                button.textContent = 'Скопировано';

                setTimeout(() => {
                    if (button && button.isConnected) {
                        button.textContent = oldText;
                    }
                }, 900);
            }

        } catch (error) {
            console.error(
                '[Prime2Key Importer] Copy error:',
                error
            );

            alert(
                'Не удалось скопировать значение.\n\n' +
                'Ошибка: ' +
                (error?.message || String(error))
            );
        }
    }

    function setNativeValue(element, value) {
        if (!element) {
            return false;
        }

        const nextValue = value ?? '';

        try {
            const prototype = Object.getPrototypeOf(element);
            const descriptor =
                Object.getOwnPropertyDescriptor(prototype, 'value');

            if (descriptor && typeof descriptor.set === 'function') {
                descriptor.set.call(element, nextValue);
            } else {
                element.value = nextValue;
            }
        } catch (error) {
            element.value = nextValue;
        }

        element.dispatchEvent(
            new Event('input', {
                bubbles: true
            })
        );

        element.dispatchEvent(
            new Event('change', {
                bubbles: true
            })
        );

        return true;
    }

    function fillBasicFields(data) {
        const result = {
            title: false,
            content: false,
            price: false,
            soldIndividually: false
        };

        const title =
            document.querySelector(
                '#title, input[name="post_title"]'
            );

        result.title =
            setNativeValue(
                title,
                data.game_name || ''
            );

        const content =
            document.querySelector(
                '#content, textarea[name="content"]'
            );

        result.content =
            setNativeValue(
                content,
                data.description_html || ''
            );

        // WordPress Classic Editor / TinyMCE.
        try {
            if (
                window.tinymce &&
                window.tinymce.get('content')
            ) {
                const editor =
                    window.tinymce.get('content');

                editor.setContent(
                    data.description_html || ''
                );

                editor.save();
                result.content = true;
            }
        } catch (error) {
            console.warn(
                '[Prime2Key Importer] TinyMCE error:',
                error
            );
        }

        const price =
            document.querySelector(
                '#_regular_price, input[name="_regular_price"]'
            );

        result.price =
            setNativeValue(
                price,
                normalizePrice(data.price)
            );

        const sold =
            document.querySelector(
                '#_sold_individually, input[name="_sold_individually"]'
            );

        if (sold) {
            if (!sold.checked) {
                sold.checked = true;

                sold.dispatchEvent(
                    new Event('input', {
                        bubbles: true
                    })
                );

                sold.dispatchEvent(
                    new Event('change', {
                        bubbles: true
                    })
                );
            }

            result.soldIndividually = true;
        }

        const missing = [];

        if (!result.title) {
            missing.push('Название (#title)');
        }

        if (!result.content) {
            missing.push('Описание (#content)');
        }

        if (!result.price) {
            missing.push('Regular price (#_regular_price)');
        }

        if (!result.soldIndividually) {
            missing.push('Sold individually (#_sold_individually)');
        }

        if (missing.length) {
            console.warn(
                '[Prime2Key Importer] Не найдены поля:',
                missing
            );

            alert(
                'Часть основных полей не найдена на странице:\n\n' +
                missing.join('\n') +
                '\n\nОстальные найденные поля были заполнены.'
            );
        }

        return result;
    }

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function getAttributesTabLink() {
        return document.querySelector(
            '.product_data_tabs .attribute_options a, ' +
            '.product_data_tabs li.attribute_options a, ' +
            'li.attribute_options a, ' +
            'a[href="#product_attributes"]'
        );
    }

    function getAttributeSearchSelect() {
        return document.querySelector(
            '#product_attributes select.wc-attribute-search, ' +
            'select.wc-attribute-search'
        );
    }

    function findAttributeRowByTaxonomy(taxonomy) {
        return document.querySelector(
            `#product_attributes .woocommerce_attribute[data-taxonomy="${escapeCssValue(taxonomy)}"]`
        );
    }

    function isVisible(element) {
        if (!element) return false;
        const style = window.getComputedStyle(element);
        return style.display !== 'none' && style.visibility !== 'hidden';
    }

    async function ensureAttributesPanelOpen() {
        let panel = document.querySelector('#product_attributes');

        if (!panel || !isVisible(panel)) {
            const tabLink = getAttributesTabLink();
            if (tabLink) {
                tabLink.click();
                await sleep(500);
            }
        }

        const started = Date.now();

        while (Date.now() - started < 8000) {
            panel = document.querySelector('#product_attributes');
            const select = getAttributeSearchSelect();

            if (panel && select) {
                return { panel, select };
            }

            await sleep(200);
        }

        throw new Error(
            'Не удалось открыть вкладку «Атрибуты» или найти список Add existing.'
        );
    }

    function escapeCssValue(value) {
        if (window.CSS && typeof window.CSS.escape === 'function') {
            return window.CSS.escape(value);
        }
        return String(value).replace(/["\\]/g, '\\$&');
    }

    async function waitForAttributeRow(taxonomy, timeout = 10000) {
        const started = Date.now();

        while (Date.now() - started < timeout) {
            const row = findAttributeRowByTaxonomy(taxonomy);
            if (row) return row;
            await sleep(200);
        }

        return null;
    }

    async function openAddExistingDropdown() {
        const select = getAttributeSearchSelect();
        if (!select) {
            throw new Error('Не найден список Add existing.');
        }

        const container = select.nextElementSibling;
        const selection = container?.querySelector('.select2-selection--single');

        if (!selection) {
            throw new Error('Не найден интерфейс Select2 для Add existing.');
        }

        selection.dispatchEvent(new MouseEvent('mousedown', {
            bubbles: true,
            cancelable: true,
            view: window
        }));

        selection.dispatchEvent(new MouseEvent('mouseup', {
            bubbles: true,
            cancelable: true,
            view: window
        }));

        selection.click();

        await sleep(250);
    }

    async function chooseAttributeFromDropdown(label) {
        await openAddExistingDropdown();

        const started = Date.now();

        while (Date.now() - started < 5000) {
            const options = Array.from(document.querySelectorAll(
                '.select2-container--open .select2-results__option[role="option"], ' +
                '.select2-dropdown .select2-results__option[role="option"]'
            ));

            const target = options.find(el =>
                String(el.textContent || '').trim().toLowerCase() ===
                String(label || '').trim().toLowerCase()
            );

            if (target) {
                target.dispatchEvent(new MouseEvent('mouseup', {
                    bubbles: true,
                    cancelable: true,
                    view: window
                }));
                target.click();
                return true;
            }

            await sleep(150);
        }

        return false;
    }

    async function addAttribute(config) {
        const existing = findAttributeRowByTaxonomy(config.taxonomy);
        if (existing) return existing;

        const ok = await chooseAttributeFromDropdown(config.label);

        if (!ok) {
            throw new Error(
                `Не удалось выбрать атрибут «${config.label}» в Add existing.`
            );
        }

        const row = await waitForAttributeRow(config.taxonomy, 10000);

        if (!row) {
            throw new Error(
                `После выбора не появился блок атрибута «${config.label}».`
            );
        }

        return row;
    }

    function splitValues(value) {
        return String(value ?? '')
            .split(/[,;|]/)
            .map(v => v.trim())
            .filter(Boolean);
    }

    async function expandAttributeRow(row) {
        if (!row) return;

        const content = row.querySelector('.woocommerce_attribute_data');
        const isOpen =
            row.classList.contains('open') ||
            (content && window.getComputedStyle(content).display !== 'none');

        if (isOpen) return;

        const header = row.querySelector('h3');
        if (header) {
            header.click();
            await sleep(250);
        }
    }

    function getValueSelect2Selection(row) {
        return row.querySelector(
            '.select2-selection--multiple'
        );
    }

    function getCreateValueButton(row) {
        return row.querySelector(
            'button.add_new_attribute, .add_new_attribute'
        );
    }

    async function openAttributeValueDropdown(row) {
        await expandAttributeRow(row);

        const selection = getValueSelect2Selection(row);

        if (!selection) {
            throw new Error(
                'Не найдено поле Select values внутри атрибута.'
            );
        }

        selection.dispatchEvent(new MouseEvent('mousedown', {
            bubbles: true,
            cancelable: true,
            view: window
        }));

        selection.dispatchEvent(new MouseEvent('mouseup', {
            bubbles: true,
            cancelable: true,
            view: window
        }));

        selection.click();

        await sleep(250);

        const search = document.querySelector(
            '.select2-container--open .select2-search__field'
        );

        if (!search) {
            throw new Error(
                'Не найдено поле ввода Select2 для значения атрибута.'
            );
        }

        return search;
    }

    async function searchAndSelectExistingValue(row, value) {
        const search = await openAttributeValueDropdown(row);

        search.focus();
        search.value = value;

        search.dispatchEvent(new Event('input', { bubbles: true }));
        search.dispatchEvent(new KeyboardEvent('keyup', {
            bubbles: true,
            key: value.slice(-1) || 'a'
        }));

        const started = Date.now();

        while (Date.now() - started < 5000) {
            const options = Array.from(document.querySelectorAll(
                '.select2-container--open .select2-results__option[role="option"]'
            ));

            const target = options.find(el => {
                const txt = String(el.textContent || '').trim().toLowerCase();
                return (
                    el.getAttribute('aria-disabled') !== 'true' &&
                    txt === String(value).trim().toLowerCase()
                );
            });

            if (target) {
                target.dispatchEvent(new MouseEvent('mouseup', {
                    bubbles: true,
                    cancelable: true,
                    view: window
                }));
                target.click();
                await sleep(300);
                return true;
            }

            await sleep(200);
        }

        document.dispatchEvent(new KeyboardEvent('keydown', {
            bubbles: true,
            key: 'Escape',
            code: 'Escape',
            keyCode: 27
        }));

        return false;
    }

    async function createNewValueViaPrompt(row, value) {
        const button = getCreateValueButton(row);

        if (!button) {
            throw new Error(
                'Не найдена кнопка Create value.'
            );
        }

        const originalPrompt = window.prompt;

        try {
            window.prompt = function () {
                return value;
            };

            button.click();
            await sleep(700);
        } finally {
            window.prompt = originalPrompt;
        }

        return true;
    }

    async function fillSingleAttributeValue(row, value) {
        if (!value) return true;

        let selected = await searchAndSelectExistingValue(row, value);

        if (selected) {
            return true;
        }

        await createNewValueViaPrompt(row, value);

        // После Create value WooCommerce обычно автоматически добавляет термин
        // в текущий атрибут. Если нет — пробуем найти и выбрать его ещё раз.
        await sleep(800);

        selected = await searchAndSelectExistingValue(row, value);

        // Даже если повторный поиск не сработал, термин мог уже быть добавлен
        // автоматически после Create value, поэтому не считаем это фатальной ошибкой.
        return true;
    }

    async function fillAttributeRow(row, rawValue) {
        const values = splitValues(rawValue);

        for (const value of values) {
            await fillSingleAttributeValue(row, value);
            await sleep(300);
        }
    }

    function normalizeReleaseDate(value) {
        const raw = String(value ?? '').trim();

        if (!raw) {
            return '';
        }

        let match = raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
        if (match) {
            const day = match[1].padStart(2, '0');
            const month = match[2].padStart(2, '0');
            const year = match[3];
            return `${day}.${month}.${year}`;
        }

        match = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
        if (match) {
            const year = match[1];
            const month = match[2].padStart(2, '0');
            const day = match[3].padStart(2, '0');
            return `${day}.${month}.${year}`;
        }

        match = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (match) {
            const day = match[1].padStart(2, '0');
            const month = match[2].padStart(2, '0');
            const year = match[3];
            return `${day}.${month}.${year}`;
        }

        const parsed = new Date(raw);

        if (!Number.isNaN(parsed.getTime())) {
            const day = String(parsed.getDate()).padStart(2, '0');
            const month = String(parsed.getMonth() + 1).padStart(2, '0');
            const year = String(parsed.getFullYear());
            return `${day}.${month}.${year}`;
        }

        return raw;
    }

    async function addAllAttributesOnly() {
        const button = panelById('p2k-attributes');
        const oldText = button ? button.textContent : '';
        const valuesByTaxonomy = {
            'pa_language': 'Английский',
            'pa_platform': 'PC',
            'pa_region': 'Global',
            'pa_restrictions': 'Без ограничений'
        };

        try {
            if (button) {
                button.disabled = true;
                button.textContent = 'Добавление атрибутов...';
            }

            await ensureAttributesPanelOpen();

            for (let i = 0; i < ATTRIBUTES.length; i++) {
                const config = ATTRIBUTES[i];

                if (button) {
                    button.textContent =
                        `Атрибут ${i + 1}/${ATTRIBUTES.length}: ${config.label}`;
                }

                const row = await addAttribute(config);
                await expandAttributeRow(row);

                const value = valuesByTaxonomy[config.taxonomy] || '';

                if (value) {
                    if (button) {
                        button.textContent = `${config.label} → ${value}`;
                    }

                    await fillSingleAttributeValue(row, value);
                }

                await sleep(350);
            }

            setStatus('Атрибуты добавлены и заполнены');

        } catch (error) {
            console.error('[Prime2Key Importer] Attribute UI error:', error);
            setStatus(`Ошибка атрибутов: ${error.message}`);
            alert(
                'Ошибка автоматического добавления атрибутов:\n\n' +
                error.message +
                '\n\nПодробности также записаны в Console (F12).'
            );
        } finally {
            if (button) {
                button.disabled = false;
                button.textContent = oldText || 'Добавить атрибуты';
            }
        }
    }

    const DUPLICATE_AFTER_PUBLISH_KEY =
        'prime2key_duplicate_after_publish_v1';

    function publishAndDuplicate() {
        const publishButton =
            document.querySelector(
                'input#publish[type="submit"][name="publish"][value="Publish"]'
            );

        if (!publishButton) {
            alert(
                'Не найдена кнопка Publish:\n\n' +
                'input#publish[name="publish"][value="Publish"]'
            );
            return;
        }

        // Сохраняем флаг, потому что Publish перезагрузит страницу.
        sessionStorage.setItem(
            DUPLICATE_AFTER_PUBLISH_KEY,
            '1'
        );

        publishButton.click();
    }

    async function continueDuplicateAfterPublish() {
        if (
            sessionStorage.getItem(
                DUPLICATE_AFTER_PUBLISH_KEY
            ) !== '1'
        ) {
            return;
        }

        // После Publish страница перезагрузилась.
        // По требованию ждём 2 секунды перед дублированием.
        await sleep(2000);

        const started = Date.now();

        while (Date.now() - started < 10000) {
            const checkbox =
                document.querySelector(
                    'input[type="checkbox"][name="icl_dupes[]"][value="en"]'
                );

            const duplicateButton =
                document.querySelector(
                    'input#icl_make_duplicates[type="button"][value="Duplicate"]'
                );

            if (checkbox && duplicateButton) {
                if (!checkbox.checked) {
                    checkbox.checked = true;

                    checkbox.dispatchEvent(
                        new Event('input', {
                            bubbles: true
                        })
                    );

                    checkbox.dispatchEvent(
                        new Event('change', {
                            bubbles: true
                        })
                    );
                }

                // Небольшая пауза, чтобы WPML обработал установку checkbox.
                await sleep(250);

                duplicateButton.click();

                // После Duplicate ждём ровно 1 секунду.
                await sleep(1000);

                const updateButton =
                    document.querySelector(
                        'input#publish[type="submit"][name="save"][value="Update"]'
                    );

                if (!updateButton) {
                    console.warn(
                        '[Prime2Key Importer] Кнопка Update не найдена.'
                    );

                    alert(
                        'Дублирование выполнено, но не найдена кнопка Update:\n\n' +
                        'input#publish[name="save"][value="Update"]'
                    );

                    return;
                }

                // Убираем флаг перед финальным Update,
                // чтобы после следующей перезагрузки сценарий не повторился.
                sessionStorage.removeItem(
                    DUPLICATE_AFTER_PUBLISH_KEY
                );

                updateButton.click();
                return;
            }

            await sleep(300);
        }

        alert(
            'После Publish не удалось выполнить дублирование.\n\n' +
            'Не найдены checkbox icl_dupes[]=en или кнопка #icl_make_duplicates.'
        );
    }

    function getPanelDocument() {
        if (!panelWindow || panelWindow.closed) return null;
        return panelWindow.document;
    }

    function panelById(id) {
        return getPanelDocument()?.getElementById(id) || null;
    }

    function ensureLauncherButton() {
        if (document.getElementById('p2k-open-panel')) return;

        const button = document.createElement('button');
        button.id = 'p2k-open-panel';
        button.type = 'button';
        button.textContent = 'Prime2Key Importer';
        button.title = 'Открыть панель Excel Importer';
        button.style.cssText = [
            'position:fixed',
            'right:18px',
            'bottom:18px',
            'z-index:999999',
            'padding:9px 13px',
            'border:1px solid #388bfd',
            'border-radius:8px',
            'background:#1f6feb',
            'color:#fff',
            'font:600 12px -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif',
            'cursor:pointer',
            'box-shadow:0 6px 20px rgba(0,0,0,.25)'
        ].join(';');

        button.addEventListener('click', () => {
            openPanelWindow(true);
        });

        document.body.appendChild(button);
    }

    function openPanelWindow(focus = true) {
        if (panelWindow && !panelWindow.closed) {
            if (focus) panelWindow.focus();
            render();
            return panelWindow;
        }

        panelWindow = window.open(
            '',
            'Prime2KeyExcelImporter',
            'width=560,height=720,resizable=yes,scrollbars=yes'
        );

        if (!panelWindow) {
            alert(
                'Браузер заблокировал отдельное окно панели.\n\n' +
                'Разрешите всплывающие окна для prime2key.com и нажмите кнопку ' +
                '«Prime2Key Importer» в правом нижнем углу.'
            );
            return null;
        }

        const doc = panelWindow.document;
        doc.open();
        doc.write(`<!doctype html>
<html lang="ru">
<head>
    <meta charset="utf-8">
    <title>Prime2Key Excel Importer</title>
</head>
<body>
    <div id="p2k-excel-importer">
        <div class="p2k-header">
            <div>
                <div class="p2k-title">Prime2Key Excel Importer</div>
                <div class="p2k-subtitle" id="p2k-status">Excel не загружен</div>
            </div>
        </div>

        <div class="p2k-body">
            <div class="p2k-upload-row">
                <label class="p2k-file-btn">
                    Выбрать Excel
                    <input type="file" id="p2k-file" accept=".xlsx,.xls,.csv" hidden>
                </label>
                <button type="button" class="p2k-secondary" id="p2k-clear">Очистить</button>
            </div>

            <div class="p2k-nav">
                <button type="button" class="p2k-secondary p2k-arrow" id="p2k-prev">←</button>
                <select id="p2k-row-select"></select>
                <button type="button" class="p2k-secondary p2k-arrow" id="p2k-next">→</button>
            </div>

            <div class="p2k-actions">
                <button type="button" class="p2k-primary" id="p2k-fill">Заполнить основные поля</button>
                <button type="button" class="p2k-primary p2k-blue" id="p2k-attributes">Добавить атрибуты</button>
            </div>

            <div class="p2k-actions">
                <button type="button" class="p2k-primary p2k-publish" id="p2k-publish-duplicate">
                    Опубликовать и Дублировать
                </button>
            </div>

            <div class="p2k-hint">
                Окно можно перемещать и изменять его размер стандартными средствами браузера.
                Кнопки работают с открытой страницей редактирования товара Prime2Key.
            </div>

            <div id="p2k-data"></div>
        </div>
    </div>
</body>
</html>`);
        doc.close();

        injectStyles(doc);
        bindEvents(doc);
        render();

        panelWindow.addEventListener('beforeunload', () => {
            panelWindow = null;
        });

        if (focus) panelWindow.focus();
        return panelWindow;
    }

    function createPanel() {
        ensureLauncherButton();

        // Пытаемся открыть окно автоматически.
        // Если браузер блокирует popup без пользовательского действия,
        // пользователь может открыть его кнопкой в правом нижнем углу.
        openPanelWindow(false);
    }

    function injectStyles(doc) {
        const style = doc.createElement('style');
        style.textContent = `
            html, body {
                margin: 0;
                min-height: 100%;
                background: #161b22;
                color: #f0f3f6;
                font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
            }

            * { box-sizing: border-box; }

            #p2k-excel-importer {
                width: 100%;
                min-height: 100vh;
                background: #161b22;
            }

            .p2k-header {
                position: sticky;
                top: 0;
                z-index: 10;
                padding: 13px 14px;
                background: #0d1117;
                border-bottom: 1px solid #30363d;
            }

            .p2k-title {
                font-size: 14px;
                font-weight: 700;
            }

            .p2k-subtitle {
                margin-top: 3px;
                font-size: 11px;
                color: #8b949e;
            }

            .p2k-body {
                padding: 12px;
            }

            .p2k-upload-row,
            .p2k-nav,
            .p2k-actions {
                display: flex;
                gap: 8px;
                margin-bottom: 10px;
            }

            #p2k-row-select {
                flex: 1;
                min-width: 0;
                background: #0d1117;
                color: #f0f3f6;
                border: 1px solid #30363d;
                border-radius: 7px;
                padding: 7px 8px;
            }

            .p2k-file-btn,
            .p2k-primary,
            .p2k-secondary,
            .p2k-copy {
                cursor: pointer;
                border-radius: 7px;
                font-size: 12px;
            }

            .p2k-file-btn,
            .p2k-primary {
                background: #238636;
                color: #fff;
                border: 1px solid #2ea043;
                padding: 8px 11px;
                text-align: center;
            }

            .p2k-file-btn {
                flex: 1;
            }

            .p2k-primary {
                flex: 1;
                font-weight: 700;
            }

            .p2k-blue {
                background: #1f6feb;
                border-color: #388bfd;
            }

            .p2k-publish {
                background: #8250df;
                border-color: #a475f9;
            }

            .p2k-primary:disabled {
                opacity: .6;
                cursor: wait;
            }

            .p2k-secondary,
            .p2k-copy {
                background: #21262d;
                color: #f0f3f6;
                border: 1px solid #30363d;
                padding: 7px 10px;
            }

            .p2k-arrow {
                min-width: 38px;
            }

            .p2k-game-link {
                display: block;
                margin: 10px 0;
                padding: 8px 9px;
                color: #58a6ff;
                background: #0d1117;
                border: 1px solid #1f6feb;
                border-radius: 7px;
                word-break: break-all;
                text-decoration: none;
                font-size: 11px;
                line-height: 1.4;
            }

            .p2k-hint {
                margin: 8px 0 10px;
                padding: 7px 9px;
                background: #21262d;
                border-radius: 7px;
                color: #8b949e;
                font-size: 10px;
                line-height: 1.35;
            }

            .p2k-compact-table {
                width: 100%;
                border-collapse: collapse;
                table-layout: fixed;
                background: #0d1117;
                border: 1px solid #30363d;
            }

            .p2k-compact-table th,
            .p2k-compact-table td {
                padding: 6px 7px;
                border-bottom: 1px solid #30363d;
                vertical-align: top;
                text-align: left;
                font-size: 11px;
                line-height: 1.35;
            }

            .p2k-compact-table tr:last-child th,
            .p2k-compact-table tr:last-child td {
                border-bottom: 0;
            }

            .p2k-compact-table th {
                width: 31%;
                color: #8b949e;
                font-weight: 600;
            }

            .p2k-compact-table td {
                color: #f0f3f6;
                word-break: break-word;
            }

            .p2k-table-cell {
                display: flex;
                align-items: flex-start;
                gap: 6px;
            }

            .p2k-table-text {
                flex: 1;
                min-width: 0;
                white-space: pre-wrap;
            }

            .p2k-table-copy {
                flex: 0 0 auto;
                padding: 2px 5px;
                font-size: 9px;
            }

            .p2k-empty {
                color: #6e7681;
                font-style: italic;
            }
        `;
        doc.head.appendChild(style);
    }

    function bindEvents(doc) {
        doc.getElementById('p2k-file').addEventListener('change', handleFile);

        doc.getElementById('p2k-fill').addEventListener('click', () => {
            const data = getCurrentData();

            if (!data) {
                setStatus('Сначала загрузите Excel и выберите игру');
                return;
            }

            try {
                const result = fillBasicFields(data);

                const filledCount =
                    Object.values(result).filter(Boolean).length;

                setStatus(
                    `Основные поля: ${filledCount}/4 • ${data.game_name}`
                );

                // Возвращаем фокус на основную вкладку, чтобы сразу видеть результат.
                try {
                    window.focus();
                } catch (error) {
                    // Не критично.
                }

            } catch (error) {
                console.error(
                    '[Prime2Key Importer] Basic fields error:',
                    error
                );

                setStatus('Ошибка заполнения основных полей');

                alert(
                    'Ошибка заполнения основных полей:\n\n' +
                    (error?.message || String(error))
                );
            }
        });

        doc.getElementById('p2k-attributes').addEventListener('click', async () => {
            await addAllAttributesOnly();
        });

        doc.getElementById('p2k-publish-duplicate').addEventListener('click', () => {
            publishAndDuplicate();
        });

        doc.getElementById('p2k-prev').addEventListener('click', () => {
            changeRow(-1);
        });

        doc.getElementById('p2k-next').addEventListener('click', () => {
            changeRow(1);
        });

        doc.getElementById('p2k-row-select').addEventListener('change', event => {
            state.currentIndex = Number(event.target.value) || 0;
            saveState();
            render();
        });

        doc.getElementById('p2k-clear').addEventListener('click', () => {
            state = {
                rows: [],
                currentIndex: 0,
                fileName: ''
            };
            saveState();
            render();
        });
    }

    async function ensureXLSX() {
        // В зависимости от режима Tampermonkey библиотека из @require
        // может быть доступна через window.XLSX или напрямую как XLSX.
        if (window.XLSX) {
            return window.XLSX;
        }

        if (typeof XLSX !== 'undefined') {
            return XLSX;
        }

        // Резервный вариант: загружаем SheetJS прямо в основную вкладку.
        const existingScript = document.querySelector(
            'script[data-p2k-xlsx-loader="1"]'
        );

        if (!existingScript) {
            const script = document.createElement('script');
            script.src =
                'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
            script.async = true;
            script.dataset.p2kXlsxLoader = '1';
            document.head.appendChild(script);
        }

        const started = Date.now();

        while (Date.now() - started < 10000) {
            if (window.XLSX) {
                return window.XLSX;
            }

            await sleep(100);
        }

        throw new Error(
            'Библиотека XLSX не загрузилась. Проверьте доступ к cdn.jsdelivr.net.'
        );
    }

    async function handleFile(event) {
        const file = event.target.files?.[0];
        if (!file) return;

        setStatus(`Чтение файла: ${file.name}...`);

        try {
            const XLSXLib = await ensureXLSX();
            const buffer = await file.arrayBuffer();

            const workbook = XLSXLib.read(
                buffer,
                {
                    type: 'array',
                    cellDates: true
                }
            );

            if (!workbook.SheetNames.length) {
                throw new Error('В Excel-файле не найдено ни одного листа.');
            }

            const firstSheet =
                workbook.Sheets[workbook.SheetNames[0]];

            const rows =
                XLSXLib.utils.sheet_to_json(
                    firstSheet,
                    {
                        defval: '',
                        raw: false
                    }
                );

            const normalizedRows =
                rows
                    .map(normalizeRow)
                    .filter(row => row.game_name);

            if (!normalizedRows.length) {
                const columns =
                    rows.length
                        ? Object.keys(rows[0]).join(', ')
                        : 'нет данных';

                throw new Error(
                    'Не найдено ни одной строки с колонкой game_name. ' +
                    'Найденные колонки: ' +
                    columns
                );
            }

            state.rows = normalizedRows;
            state.currentIndex = 0;
            state.fileName = file.name;

            saveState();
            render();

            setStatus(
                `${file.name} • загружено ${state.rows.length} игр`
            );

        } catch (error) {
            console.error(
                '[Prime2Key Importer] Excel parse error:',
                error
            );

            setStatus('Ошибка чтения Excel');

            alert(
                'Не удалось прочитать Excel-файл.\n\n' +
                (error?.message || String(error))
            );
        } finally {
            // Позволяет повторно выбрать тот же файл.
            event.target.value = '';
        }
    }

    function changeRow(delta) {
        if (!state.rows.length) return;
        state.currentIndex = Math.max(0, Math.min(state.rows.length - 1, state.currentIndex + delta));
        saveState();
        render();
    }

    function getCurrentData() {
        return state.rows[state.currentIndex] || null;
    }

    function setStatus(text) {
        const element = panelById('p2k-status');
        if (element) element.textContent = text;
    }


    function render() {
        const select = panelById('p2k-row-select');
        const dataContainer = panelById('p2k-data');
        if (!select || !dataContainer) return;

        if (!state.rows.length) {
            select.innerHTML = '<option value="0">Нет данных</option>';
            dataContainer.innerHTML =
                '<div class="p2k-hint">Загрузите .xlsx, .xls или .csv файл.</div>';
            setStatus('Excel не загружен');
            return;
        }

        select.innerHTML = state.rows.map((row, index) => `
            <option value="${index}" ${index === state.currentIndex ? 'selected' : ''}>
                ${index + 1}. ${escapeHtml(row.game_name || `Строка ${index + 1}`)}
            </option>
        `).join('');

        const data = getCurrentData();

        const fields = [
            ['categories', 'Categories'],
            ['developer', 'Developer'],
            ['features', 'Features'],
            ['publisher', 'Publisher'],
            ['release_date', 'Release date'],
            ['__platform', 'Платформа', 'PC']
        ];

        const linkHtml = data.source_url
            ? `<a class="p2k-game-link" href="${escapeHtml(data.source_url)}" target="_blank" rel="noopener noreferrer">
                    Открыть источник игры ↗<br>${escapeHtml(data.source_url)}
               </a>`
            : `<div class="p2k-game-link p2k-empty">source_url отсутствует</div>`;

        const tableHtml = `
            <table class="p2k-compact-table">
                <tbody>
                    ${fields.map(([key, label, fixedValue]) => {
                        const value = fixedValue || data[key] || '';
                        return `
                            <tr>
                                <th>${escapeHtml(label)}</th>
                                <td>
                                    <div class="p2k-table-cell">
                                        <span class="p2k-table-text ${value ? '' : 'p2k-empty'}">
                                            ${value ? escapeHtml(value) : 'Нет данных'}
                                        </span>
                                        <button
                                            type="button"
                                            class="p2k-copy p2k-table-copy"
                                            data-copy-value="${escapeHtml(value)}"
                                        >
                                            Копировать
                                        </button>
                                    </div>
                                </td>
                            </tr>`;
                    }).join('')}
                </tbody>
            </table>`;

        dataContainer.innerHTML = linkHtml + tableHtml;

        dataContainer.querySelectorAll('.p2k-copy').forEach(button => {
            button.addEventListener('click', () => copyText(button.dataset.copyValue || '', button));
        });

        setStatus(`${state.fileName || 'Excel'} • ${state.currentIndex + 1}/${state.rows.length}`);
    }

    createPanel();

    continueDuplicateAfterPublish().catch(error => {
        console.error(
            '[Prime2Key Importer] Duplicate continuation error:',
            error
        );
    });
})();
