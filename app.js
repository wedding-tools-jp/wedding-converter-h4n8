import {
    loadPdf,
    renderPdfPageToCanvas,
    renderPdfPagePreviewDataUrl,
    fileToDataUrl,
    fileToCanvas,
} from './pdf-utils.js';
import { buildPpt } from './ppt-builder.js';
import { buildIdsPackage, makePptFilename } from './ids-csv.js';
import { pickUsbDirectory, writeIdsPackage, isSupported as fsIsSupported, UserCancelled } from './filesystem.js';

const WEEKEND_RANGE_MONTHS = 3;
const DOW_LABELS = ['日', '月', '火', '水', '木', '金', '土'];
const MAX_FILES = 2;
const MAX_PAGES_PER_PDF = 50;
const DEFAULT_DISPLAY_SECONDS = 30;
const WIN_RESERVED = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$/i;

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const selectBtn = document.getElementById('selectBtn');
const fileList = document.getElementById('fileList');
const fileListItems = document.getElementById('fileListItems');
const convertBtn = document.getElementById('convertBtn');
const clearBtn = document.getElementById('clearBtn');
const progress = document.getElementById('progress');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const result = document.getElementById('result');
const resetBtn = document.getElementById('resetBtn');
const dateSelect = document.getElementById('dateSelect');
const dateCustom = document.getElementById('dateCustom');
const dateIconBtn = document.getElementById('dateIconBtn');
const customerName = document.getElementById('customerName');
const filenamePreview = document.getElementById('filenamePreview');
const previewEmpty = document.getElementById('previewEmpty');
const previewGrid = document.getElementById('previewGrid');

let selectedFiles = [];

initDateSelector();
checkBrowserSupport();

selectBtn.addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => { handleFiles(Array.from(e.target.files)); fileInput.value = ''; });
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    handleFiles(Array.from(e.dataTransfer.files));
});

clearBtn.addEventListener('click', () => {
    selectedFiles = [];
    customerName.value = '';
    renderFileList();
    renderPreviews();
    updatePreview();
});

resetBtn.addEventListener('click', () => {
    selectedFiles = [];
    customerName.value = '';
    renderFileList();
    renderPreviews();
    updatePreview();
    result.style.display = 'none';
    dropZone.style.display = 'block';
});

convertBtn.addEventListener('click', exportToUsb);
customerName.addEventListener('input', updatePreview);

dateSelect.addEventListener('change', () => {
    dateCustom.value = dateSelect.value;
    updatePreview();
});

dateIconBtn.addEventListener('click', () => {
    if (typeof dateCustom.showPicker === 'function') {
        try { dateCustom.showPicker(); return; } catch (e) { /* fall through */ }
    }
    dateCustom.focus();
    dateCustom.click();
});

dateCustom.addEventListener('change', () => {
    const matchOption = Array.from(dateSelect.options).find(o => o.value === dateCustom.value);
    dateSelect.value = matchOption ? dateCustom.value : '';
    updatePreview();
});

function checkBrowserSupport() {
    if (!fsIsSupported()) {
        alert('このブラウザは USB への直接書き出しに対応していません。\nChrome または Edge をご使用ください。');
        convertBtn.disabled = true;
    }
}

function initDateSelector() {
    const options = generateWeekendOptions();
    const defaultDate = formatYMD(getDefaultDate());
    options.forEach(opt => {
        const option = document.createElement('option');
        option.value = opt.value;
        option.textContent = opt.label;
        if (opt.value === defaultDate) option.selected = true;
        dateSelect.appendChild(option);
    });
    dateCustom.value = defaultDate;
    updatePreview();
}

function getDefaultDate() {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dow = today.getDay();
    if (dow === 0 || dow === 6) return today;
    const daysUntilSat = (6 - dow + 7) % 7;
    const next = new Date(today);
    next.setDate(today.getDate() + daysUntilSat);
    return next;
}

function generateWeekendOptions() {
    const start = new Date();
    start.setHours(0, 0, 0, 0);
    const end = new Date(start);
    end.setMonth(start.getMonth() + WEEKEND_RANGE_MONTHS);
    const options = [];
    const cursor = new Date(start);
    while (cursor <= end) {
        const dow = cursor.getDay();
        if (dow === 0 || dow === 6) {
            options.push({ value: formatYMD(cursor), label: formatLabel(cursor) });
        }
        cursor.setDate(cursor.getDate() + 1);
    }
    return options;
}

function formatYMD(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function formatLabel(d) {
    return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')} (${DOW_LABELS[d.getDay()]})`;
}

function getCurrentDateStr() {
    return dateCustom.value || dateSelect.value || '';
}

function parseDate(ymd) {
    if (!ymd) return null;
    const [y, m, d] = ymd.split('-').map(Number);
    return new Date(y, m - 1, d);
}

function sanitizeName(name) {
    let s = name.replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').trim();
    if (WIN_RESERVED.test(s)) s = '_' + s;
    return s;
}

function buildLabel() {
    const dateStr = getCurrentDateStr();
    const ymd = dateStr.replace(/-/g, '');
    const rawName = customerName.value.trim().replace(/家$/, '');
    const name = sanitizeName(rawName);
    if (!ymd) return '';
    if (!name) return `${ymd}`;
    return `${ymd}_${name}家`;
}

function updatePreview() {
    const label = buildLabel();
    const hasName = customerName.value.trim().replace(/家$/, '').length > 0;
    const hasSanitized = sanitizeName(customerName.value.trim().replace(/家$/, '')).length > 0;
    const hasDate = getCurrentDateStr().length > 0;
    filenamePreview.textContent = label || '—';
    convertBtn.disabled = !(hasName && hasSanitized && hasDate && selectedFiles.length > 0 && fsIsSupported());
}

function handleFiles(files) {
    const accepted = files.filter(f => {
        const ext = f.name.toLowerCase().split('.').pop();
        return ['pdf', 'jpg', 'jpeg', 'png'].includes(ext);
    });
    if (accepted.length === 0) {
        alert('対応していないファイル形式です。PDF・JPEG・PNG のみ対応しています。');
        return;
    }
    if (accepted.length < files.length) {
        alert(`${files.length - accepted.length} 個のファイルは対応していない形式のためスキップされました。`);
    }
    const remaining = MAX_FILES - selectedFiles.length;
    if (remaining <= 0) {
        alert(`同時に処理できるファイルは ${MAX_FILES} 個までです。`);
        return;
    }
    const toAdd = accepted.slice(0, remaining);
    if (accepted.length > remaining) {
        alert(`同時に処理できるファイルは ${MAX_FILES} 個までです。最初の ${toAdd.length} 個だけ追加しました。`);
    }
    selectedFiles = [...selectedFiles, ...toAdd];
    renderFileList();
    renderPreviews();
    updatePreview();
}

function renderFileList() {
    if (selectedFiles.length === 0) { fileList.style.display = 'none'; return; }
    fileList.style.display = 'block';
    fileListItems.innerHTML = '';
    selectedFiles.forEach((file, index) => {
        const ext = file.name.toLowerCase().split('.').pop();
        const isPdf = ext === 'pdf';
        const li = document.createElement('li');
        li.innerHTML = `
            <div class="file-icon ${isPdf ? 'pdf' : 'img'}">${isPdf ? 'PDF' : escapeHtml(ext.toUpperCase())}</div>
            <div class="file-name">${escapeHtml(file.name)}</div>
            <div class="file-size">${formatSize(file.size)}</div>
            <button class="file-remove" data-index="${index}" title="削除">×</button>
        `;
        fileListItems.appendChild(li);
    });
    fileListItems.querySelectorAll('.file-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.target.dataset.index);
            selectedFiles.splice(idx, 1);
            renderFileList();
            renderPreviews();
            updatePreview();
        });
    });
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1024 / 1024).toFixed(1) + ' MB';
}

function updateProgress(current, total, label) {
    const pct = total === 0 ? 0 : Math.round((current / total) * 100);
    progressFill.style.width = pct + '%';
    progressText.textContent = `${label}（${current} / ${total}）`;
}

// Open every PDF once, validate page counts, then render in a single pass.
// Returns { slides: Canvas[], cleanup: () => Promise<void> } so PDF.js workers can be released.
async function collectSlideCanvases() {
    const opened = [];
    let totalPages = 0;
    for (const file of selectedFiles) {
        const ext = file.name.toLowerCase().split('.').pop();
        if (ext === 'pdf') {
            const pdf = await loadPdf(file);
            if (pdf.numPages > MAX_PAGES_PER_PDF) {
                await Promise.all(opened.filter(o => o.pdf).map(o => o.pdf.destroy().catch(() => {})));
                await pdf.destroy().catch(() => {});
                const e = new Error(`PDF のページ数が多すぎます（${pdf.numPages}ページ / 上限 ${MAX_PAGES_PER_PDF}ページ）: ${file.name}`);
                e.name = 'TooManyPagesError';
                throw e;
            }
            opened.push({ file, pdf });
            totalPages += pdf.numPages;
        } else {
            opened.push({ file, pdf: null });
            totalPages += 1;
        }
    }

    const slides = [];
    let processed = 0;
    try {
        for (const o of opened) {
            if (o.pdf) {
                for (let p = 1; p <= o.pdf.numPages; p++) {
                    updateProgress(processed, totalPages, 'ページを変換中');
                    slides.push(await renderPdfPageToCanvas(o.pdf, p));
                    processed++;
                }
            } else {
                updateProgress(processed, totalPages, 'ページを変換中');
                slides.push(await fileToCanvas(o.file));
                processed++;
            }
        }
    } finally {
        await Promise.all(opened.filter(o => o.pdf).map(o => o.pdf.destroy().catch(() => {})));
    }
    return slides;
}

async function exportToUsb() {
    if (selectedFiles.length === 0) return;
    const dateStr = getCurrentDateStr();
    const date = parseDate(dateStr);
    if (!date) { alert('日付を選択してください。'); return; }
    const label = buildLabel();
    if (!label) { alert('客名を入力してください。'); return; }

    // 1. Ask the user to pick the USB directory FIRST, while we still have the user gesture.
    let usbHandle;
    try {
        usbHandle = await pickUsbDirectory();
    } catch (err) {
        if (err instanceof UserCancelled) return;
        console.error(err);
        alert('書き出し先の選択に失敗しました。');
        return;
    }
    const confirmed = confirm(
        `「${usbHandle.name}」に書き出します。\n\n` +
        `本当にこのフォルダで間違いありませんか？\n` +
        `（USB ドライブ以外を選んでいる場合は「キャンセル」を押してください）`
    );
    if (!confirmed) return;

    fileList.style.display = 'none';
    dropZone.style.display = 'none';
    progress.style.display = 'block';
    progressFill.style.width = '0%';
    progressText.textContent = '画像を準備中...';

    try {
        const canvases = await collectSlideCanvases();
        const total = canvases.length;
        const pptFiles = [];
        for (let i = 0; i < total; i++) {
            updateProgress(i, total, 'PowerPointを生成中');
            const filename = makePptFilename(date, i, total);
            const bytes = await buildPpt(canvases[i]);
            pptFiles.push({ filename, bytes, displaySeconds: DEFAULT_DISPLAY_SECONDS });
            // free canvas memory
            canvases[i].width = 0;
            canvases[i].height = 0;
        }

        progressText.textContent = 'パッケージを構築中...';
        const tree = buildIdsPackage({
            pptFiles,
            scheduleStartDate: date,
            timestamp: new Date(),
        });

        progressText.textContent = 'USB に書き込み中...';
        await writeIdsPackage(usbHandle, tree, ({ done, total, label: itemLabel }) => {
            updateProgress(done, total, 'USB に書き込み中');
        });

        progress.style.display = 'none';
        result.style.display = 'block';
        selectedFiles = [];
        customerName.value = '';
        renderPreviews();
        updatePreview();
    } catch (err) {
        console.error(err);
        let msg = '書き出し中にエラーが発生しました。';
        if (err && err.name === 'QuotaExceededError') {
            msg = 'USB の空き容量が不足しています。空きを確保してから再度お試しください。';
        } else if (err && err.name === 'NotAllowedError') {
            msg = 'USB への書き込みが許可されませんでした。書き込み権限のあるドライブを選択してください。';
        } else if (err && err.name === 'NoModificationAllowedError') {
            msg = 'USB が書き込み禁止になっています。ロックスイッチや別ファイルでの使用を確認してください。';
        } else if (err && err.name === 'TooManyPagesError') {
            msg = err.message;
        }
        alert(msg);
        progress.style.display = 'none';
        dropZone.style.display = 'block';
        renderFileList();
    }
}

async function renderPreviews() {
    if (selectedFiles.length === 0) {
        previewEmpty.style.display = 'flex';
        previewGrid.style.display = 'none';
        previewGrid.innerHTML = '';
        previewGrid.classList.remove('cols-2');
        return;
    }
    previewEmpty.style.display = 'none';
    previewGrid.style.display = 'grid';
    previewGrid.classList.toggle('cols-2', selectedFiles.length === 2);
    previewGrid.innerHTML = '';
    for (const file of selectedFiles) {
        const card = document.createElement('div');
        card.className = 'preview-card';
        card.innerHTML = `
            <div class="preview-card-img"><div class="preview-loading">読み込み中...</div></div>
            <div class="preview-card-meta">
                <span class="preview-card-name">${escapeHtml(file.name)}</span>
                <span class="preview-card-pages"></span>
            </div>
        `;
        previewGrid.appendChild(card);
        renderPreviewIntoCard(file, card).catch(err => {
            console.error(err);
            card.querySelector('.preview-card-img').innerHTML = '<div class="preview-loading">プレビュー失敗</div>';
        });
    }
}

async function renderPreviewIntoCard(file, card) {
    const ext = file.name.toLowerCase().split('.').pop();
    const imgWrap = card.querySelector('.preview-card-img');
    const pagesLabel = card.querySelector('.preview-card-pages');
    if (ext === 'pdf') {
        const pdf = await loadPdf(file);
        try {
            pagesLabel.textContent = `${pdf.numPages}ページ`;
            const dataUrl = await renderPdfPagePreviewDataUrl(pdf, 1);
            const img = new Image();
            img.src = dataUrl;
            imgWrap.innerHTML = '';
            imgWrap.appendChild(img);
        } finally {
            await pdf.destroy().catch(() => {});
        }
    } else {
        const dataUrl = await fileToDataUrl(file);
        pagesLabel.textContent = '画像';
        const img = new Image();
        img.src = dataUrl;
        imgWrap.innerHTML = '';
        imgWrap.appendChild(img);
    }
}
