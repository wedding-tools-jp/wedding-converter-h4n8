// Build a .ppt by byte-substituting the embedded JPEG in a bundled template.
// See TEMPLATE_SPEC.md for the byte layout assumptions.

const TEMPLATE_URL = 'assets/template.ppt';
const JPEG_OFFSET = 537;
const JPEG_SLOT_SIZE = 224068;
const TEMPLATE_TOTAL_SIZE = 337920;

let cachedTemplate = null;

async function loadTemplate() {
    if (cachedTemplate) return cachedTemplate;
    const res = await fetch(TEMPLATE_URL);
    if (!res.ok) throw new Error(`テンプレート読み込み失敗: ${res.status}`);
    const buf = new Uint8Array(await res.arrayBuffer());
    if (buf.length !== TEMPLATE_TOTAL_SIZE) {
        throw new Error(`テンプレートのサイズが想定外: ${buf.length} (期待: ${TEMPLATE_TOTAL_SIZE})`);
    }
    cachedTemplate = buf;
    return buf;
}

// Re-encode a canvas to JPEG that fits within targetBytes (with at least 4 bytes headroom
// so COM-marker padding is always possible), then pad to exactly targetBytes.
// A COM segment is minimum 4 bytes, so we need either need==0 or need>=4.
async function encodeJpegExactSize(canvas, targetBytes) {
    const ceiling = targetBytes - 4; // leave room for at least one COM segment
    let lo = 0.30, hi = 0.95;
    let blob = await canvasToBlob(canvas, hi);
    if (blob.size > ceiling) {
        for (let i = 0; i < 10; i++) {
            const q = (lo + hi) / 2;
            blob = await canvasToBlob(canvas, q);
            if (blob.size > ceiling) hi = q; else lo = q;
        }
        blob = await canvasToBlob(canvas, lo);
        if (blob.size > ceiling) {
            throw new Error(`JPEGエンコード失敗: 最低品質でも ${blob.size} バイト (上限 ${ceiling})`);
        }
    }
    let bytes = new Uint8Array(await blob.arrayBuffer());
    // Final safety: if size landed in the un-paddable gap (targetBytes-3..targetBytes-1),
    // try slightly lower quality.
    let attempt = 0;
    while (bytes.length > targetBytes - 4 && bytes.length !== targetBytes && attempt < 5) {
        lo = Math.max(0.10, lo - 0.05);
        blob = await canvasToBlob(canvas, lo);
        bytes = new Uint8Array(await blob.arrayBuffer());
        attempt++;
    }
    return padJpegToExactSize(bytes, targetBytes);
}

function canvasToBlob(canvas, quality) {
    return new Promise((resolve, reject) => {
        canvas.toBlob(b => b ? resolve(b) : reject(new Error('toBlob失敗')), 'image/jpeg', quality);
    });
}

// JPEG layout: SOI(2) + segments... + EOI(2). We insert a COM segment (FF FE LEN_HI LEN_LO data)
// right after SOI to pad to exact size. Max COM segment size is 65533 bytes of data.
function padJpegToExactSize(jpeg, targetBytes) {
    if (jpeg.length === targetBytes) return jpeg;
    if (jpeg.length > targetBytes) {
        throw new Error(`JPEGがターゲットより大きい: ${jpeg.length} > ${targetBytes}`);
    }
    if (jpeg[0] !== 0xFF || jpeg[1] !== 0xD8) {
        throw new Error('JPEG SOI マーカーが見つからない');
    }

    const need = targetBytes - jpeg.length;
    // Each COM segment is marker(2)+length(2)+data, minimum 4 bytes total.
    // So `need` must be 0 (handled above) or >= 4. encodeJpegExactSize() guarantees this.
    if (need < 4) {
        throw new Error(`パディング不可能サイズ: 残${need}バイト (4バイト未満)`);
    }
    const segments = [];
    let remaining = need;
    while (remaining > 0) {
        // max segLen = 65535 → total bytes = 65537
        const totalSegBytes = Math.min(remaining, 65537);
        if (totalSegBytes < 4) {
            // last fragment of 1-3 bytes left — shrink previous to make this fit as 4-byte segment.
            // encodeJpegExactSize() guarantees need >= 4, so segments[] is never empty here, but guard regardless.
            if (segments.length === 0) {
                throw new Error(`パディング不可能: 残${totalSegBytes}バイトで先行セグメントなし`);
            }
            const lastSeg = segments[segments.length - 1];
            lastSeg.size -= (4 - totalSegBytes);
            segments.push({ size: 4 });
            break;
        }
        segments.push({ size: totalSegBytes });
        remaining -= totalSegBytes;
    }

    const padTotal = segments.reduce((s, x) => s + x.size, 0);
    const out = new Uint8Array(jpeg.length + padTotal);
    out[0] = 0xFF; out[1] = 0xD8;
    let pos = 2;
    for (const seg of segments) {
        out[pos++] = 0xFF;
        out[pos++] = 0xFE;
        const segLen = seg.size - 2; // length field value includes itself
        out[pos++] = (segLen >> 8) & 0xFF;
        out[pos++] = segLen & 0xFF;
        pos += seg.size - 4; // zeros (Uint8Array default)
    }
    out.set(jpeg.subarray(2), pos);
    if (out.length !== targetBytes) {
        throw new Error(`パディング後サイズ不一致: ${out.length} != ${targetBytes}`);
    }
    return out;
}

// Render a source image/canvas into a fixed 1684x1190 JPEG (matching template image dimensions),
// letterboxed on black to preserve aspect ratio.
const TARGET_IMG_W = 1684;
const TARGET_IMG_H = 1190;

async function renderToTemplateSizedJpeg(sourceCanvas) {
    const canvas = document.createElement('canvas');
    canvas.width = TARGET_IMG_W;
    canvas.height = TARGET_IMG_H;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#000000';
    ctx.fillRect(0, 0, TARGET_IMG_W, TARGET_IMG_H);

    const srcW = sourceCanvas.width, srcH = sourceCanvas.height;
    const srcRatio = srcW / srcH;
    const dstRatio = TARGET_IMG_W / TARGET_IMG_H;
    let dw, dh, dx, dy;
    if (srcRatio > dstRatio) {
        dw = TARGET_IMG_W;
        dh = TARGET_IMG_W / srcRatio;
        dx = 0;
        dy = (TARGET_IMG_H - dh) / 2;
    } else {
        dh = TARGET_IMG_H;
        dw = TARGET_IMG_H * srcRatio;
        dx = (TARGET_IMG_W - dw) / 2;
        dy = 0;
    }
    ctx.drawImage(sourceCanvas, dx, dy, dw, dh);
    return await encodeJpegExactSize(canvas, JPEG_SLOT_SIZE);
}

// Build a .ppt Uint8Array from a single source canvas.
export async function buildPpt(sourceCanvas) {
    const tpl = await loadTemplate();
    const jpeg = await renderToTemplateSizedJpeg(sourceCanvas);
    if (jpeg.length !== JPEG_SLOT_SIZE) {
        throw new Error(`JPEG slot size mismatch: ${jpeg.length}`);
    }
    const out = new Uint8Array(TEMPLATE_TOTAL_SIZE);
    out.set(tpl.subarray(0, JPEG_OFFSET), 0);
    out.set(jpeg, JPEG_OFFSET);
    out.set(tpl.subarray(JPEG_OFFSET + JPEG_SLOT_SIZE), JPEG_OFFSET + JPEG_SLOT_SIZE);
    return out;
}
