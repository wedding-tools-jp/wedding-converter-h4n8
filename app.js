import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.worker.min.mjs';

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;
const PDF_RENDER_SCALE = 2.0;
const BG_COLOR = '000000';
const WEEKEND_RANGE_MONTHS = 3;
const DOW_LABELS = ['日', '月', '火', '水', '木', '金', '土'];

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
const dateOtherBtn = document.getElementById('dateOtherBtn');
const dateOtherText = dateOtherBtn.querySelector('.date-other-text');
const customerName = document.getElementById('customerName');
const filenamePreview = document.getElementById('filenamePreview');

let selectedFiles = [];

initDateSelector();

selectBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
});

dropZone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => {
    handleFiles(Array.from(e.target.files));
    fileInput.value = '';
});

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    handleFiles(Array.from(e.dataTransfer.files));
});

clearBtn.addEventListener('click', () => {
    selectedFiles = [];
    customerName.value = '';
    renderFileList();
    updatePreview();
});

resetBtn.addEventListener('click', () => {
    selectedFiles = [];
    customerName.value = '';
    renderFileList();
    updatePreview();
    result.style.display = 'none';
    dropZone.style.display = 'block';
});

convertBtn.addEventListener('click', convertToPptx);
customerName.addEventListener('input', updatePreview);

dateSelect.addEventListener('change', () => {
    dateCustom.value = dateSelect.value;
    resetOtherButton();
    updatePreview();
});

dateOtherBtn.addEventListener('click', () => {
    if (typeof dateCustom.showPicker === 'function') {
        try {
            dateCustom.showPicker();
            return;
        } catch (e) {
            // fall through
        }
    }
    dateCustom.focus();
    dateCustom.click();
});

dateCustom.addEventListener('change', () => {
    const matchOption = Array.from(dateSelect.options).find(o => o.value === dateCustom.value);
    if (matchOption) {
        dateSelect.value = dateCustom.value;
        resetOtherButton();
    } else {
        dateSelect.value = '';
        showPickedOnButton(dateCustom.value);
    }
    updatePreview();
});

function resetOtherButton() {
    dateOtherText.textContent = 'その他の日付を選択';
    dateOtherBtn.classList.remove('has-date');
}

function showPickedOnButton(ymdStr) {
    if (!ymdStr) {
        resetOtherButton();
        return;
    }
    const [y, m, d] = ymdStr.split('-').map(Number);
    const date = new Date(y, m - 1, d);
    dateOtherText.textContent = formatLabel(date);
    dateOtherBtn.classList.add('has-date');
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
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

function formatLabel(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}/${m}/${day} (${DOW_LABELS[d.getDay()]})`;
}

function getCurrentDateStr() {
    return dateCustom.value || dateSelect.value || '';
}

function sanitizeName(name) {
    return name.replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').trim();
}

function buildFilename() {
    const dateStr = getCurrentDateStr();
    const ymd = dateStr.replace(/-/g, '');
    const rawName = customerName.value.trim().replace(/家$/, '');
    const name = sanitizeName(rawName);
    if (!ymd) return '';
    if (!name) return `${ymd}.pptx`;
    return `${ymd}_${name}家.pptx`;
}

function updatePreview() {
    const fname = buildFilename();
    const hasName = customerName.value.trim().length > 0;
    const hasDate = getCurrentDateStr().length > 0;
    filenamePreview.textContent = fname || '—';
    convertBtn.disabled = !(hasName && hasDate && selectedFiles.length > 0);
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

    selectedFiles = [...selectedFiles, ...accepted];
    renderFileList();
    updatePreview();
}

function renderFileList() {
    if (selectedFiles.length === 0) {
        fileList.style.display = 'none';
        return;
    }

    fileList.style.display = 'block';
    fileListItems.innerHTML = '';

    selectedFiles.forEach((file, index) => {
        const ext = file.name.toLowerCase().split('.').pop();
        const isPdf = ext === 'pdf';
        const li = document.createElement('li');
        li.innerHTML = `
            <div class="file-icon ${isPdf ? 'pdf' : 'img'}">${isPdf ? 'PDF' : ext.toUpperCase()}</div>
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

async function convertToPptx() {
    if (selectedFiles.length === 0) return;
    const filename = buildFilename();
    if (!filename) {
        alert('日付と客名を入力してください。');
        return;
    }

    fileList.style.display = 'none';
    dropZone.style.display = 'none';
    progress.style.display = 'block';
    progressFill.style.width = '0%';
    progressText.textContent = '画像を準備中...';

    try {
        const slideImages = [];

        let totalPages = 0;
        for (const file of selectedFiles) {
            const ext = file.name.toLowerCase().split('.').pop();
            if (ext === 'pdf') {
                const pdf = await loadPdf(file);
                totalPages += pdf.numPages;
                slideImages.push({ type: 'pdf', pdf, file });
            } else {
                totalPages += 1;
                slideImages.push({ type: 'image', file });
            }
        }

        let processed = 0;
        const renderedSlides = [];

        for (const item of slideImages) {
            if (item.type === 'pdf') {
                for (let p = 1; p <= item.pdf.numPages; p++) {
                    updateProgress(processed, totalPages, 'ページを変換中');
                    const dataUrl = await renderPdfPage(item.pdf, p);
                    const dims = await getImageDimensions(dataUrl);
                    renderedSlides.push({ dataUrl, dims });
                    processed++;
                }
            } else {
                updateProgress(processed, totalPages, 'ページを変換中');
                const dataUrl = await fileToDataUrl(item.file);
                const dims = await getImageDimensions(dataUrl);
                renderedSlides.push({ dataUrl, dims });
                processed++;
            }
        }

        updateProgress(processed, totalPages, 'PowerPointを生成中');

        const pptx = new PptxGenJS();
        pptx.defineLayout({ name: 'CUSTOM_16_9', width: SLIDE_W, height: SLIDE_H });
        pptx.layout = 'CUSTOM_16_9';

        for (const slide of renderedSlides) {
            const s = pptx.addSlide();
            s.background = { color: BG_COLOR };
            const fit = calculateFit(slide.dims.width, slide.dims.height);
            s.addImage({
                data: slide.dataUrl,
                x: fit.x,
                y: fit.y,
                w: fit.w,
                h: fit.h
            });
        }

        await pptx.writeFile({ fileName: filename });

        progress.style.display = 'none';
        result.style.display = 'block';
        selectedFiles = [];
        customerName.value = '';
        updatePreview();
    } catch (err) {
        console.error(err);
        alert('変換中にエラーが発生しました：\n' + (err.message || err));
        progress.style.display = 'none';
        dropZone.style.display = 'block';
        renderFileList();
    }
}

function calculateFit(imgW, imgH) {
    const slideRatio = SLIDE_W / SLIDE_H;
    const imgRatio = imgW / imgH;

    let w, h, x, y;

    if (imgRatio > slideRatio) {
        w = SLIDE_W;
        h = SLIDE_W / imgRatio;
        x = 0;
        y = (SLIDE_H - h) / 2;
    } else {
        h = SLIDE_H;
        w = SLIDE_H * imgRatio;
        x = (SLIDE_W - w) / 2;
        y = 0;
    }

    return { x, y, w, h };
}

function fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error('ファイル読み込みエラー: ' + file.name));
        reader.readAsDataURL(file);
    });
}

function getImageDimensions(dataUrl) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
        img.onerror = () => reject(new Error('画像の読み込みに失敗しました'));
        img.src = dataUrl;
    });
}

async function loadPdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
    return await loadingTask.promise;
}

async function renderPdfPage(pdf, pageNum) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale: PDF_RENDER_SCALE });

    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');

    await page.render({ canvasContext: ctx, viewport }).promise;

    return canvas.toDataURL('image/jpeg', 0.92);
}
