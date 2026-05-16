// PDF/image rendering utilities. Pure I/O helpers extracted from app.js.

import * as pdfjsLib from './assets/vendor/pdfjs/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = './assets/vendor/pdfjs/pdf.worker.min.mjs';

const PDF_RENDER_SCALE = 2.0;
const PREVIEW_RENDER_SCALE = 1.2;

export async function loadPdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    return await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
}

// Render a PDF page to a canvas at high resolution (used for final output).
export async function renderPdfPageToCanvas(pdf, pageNum) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale: PDF_RENDER_SCALE });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport }).promise;
    return canvas;
}

// Render at lower resolution for the on-screen preview.
export async function renderPdfPagePreviewDataUrl(pdf, pageNum) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale: PREVIEW_RENDER_SCALE });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport }).promise;
    return canvas.toDataURL('image/jpeg', 0.85);
}

export function fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error('ファイル読み込みエラー: ' + file.name));
        reader.readAsDataURL(file);
    });
}

export function dataUrlToImage(dataUrl) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => reject(new Error('画像の読み込みに失敗しました'));
        img.src = dataUrl;
    });
}

export async function imageToCanvas(img) {
    const canvas = document.createElement('canvas');
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0);
    return canvas;
}

export async function fileToCanvas(file) {
    const url = await fileToDataUrl(file);
    const img = await dataUrlToImage(url);
    return await imageToCanvas(img);
}
