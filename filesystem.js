// File System Access API wrapper for writing the IDS Package to a USB drive.

const PACKAGE_ROOT = 'IDS Package';

export class UserCancelled extends Error {
    constructor() { super('ユーザーがキャンセルしました'); this.name = 'UserCancelled'; }
}

export function isSupported() {
    return typeof window !== 'undefined' && typeof window.showDirectoryPicker === 'function';
}

// Prompt for a directory (USB) and return the handle.
export async function pickUsbDirectory() {
    if (!isSupported()) {
        throw new Error('このブラウザは File System Access API に対応していません。Chrome / Edge を使用してください。');
    }
    try {
        return await window.showDirectoryPicker({ mode: 'readwrite', id: 'ids-usb-target' });
    } catch (err) {
        if (err && err.name === 'AbortError') throw new UserCancelled();
        throw err;
    }
}

// Recursively get or create a directory under root.
async function ensureDir(rootHandle, segments) {
    let h = rootHandle;
    for (const seg of segments) {
        h = await h.getDirectoryHandle(seg, { create: true });
    }
    return h;
}

// Remove the package root if it already exists (overwrite behavior).
async function removeExisting(rootHandle, name) {
    try {
        await rootHandle.removeEntry(name, { recursive: true });
    } catch (err) {
        if (err && err.name === 'NotFoundError') return;
        throw err;
    }
}

// Write the IDS Package tree under {selectedDir}/IDS Package/.
// files: [{ path: string[], bytes: Uint8Array }, ...]
// onProgress: ({ done, total, label }) optional callback.
export async function writeIdsPackage(rootHandle, files, onProgress) {
    await removeExisting(rootHandle, PACKAGE_ROOT);
    const pkgRoot = await rootHandle.getDirectoryHandle(PACKAGE_ROOT, { create: true });

    let done = 0;
    const total = files.length;
    for (const f of files) {
        if (!f.path || f.path.length === 0) throw new Error('empty path');
        const dirs = f.path.slice(0, -1);
        const name = f.path[f.path.length - 1];
        const dirHandle = dirs.length === 0 ? pkgRoot : await ensureDir(pkgRoot, dirs);
        const fileHandle = await dirHandle.getFileHandle(name, { create: true });
        const writable = await fileHandle.createWritable();
        try {
            if (f.bytes && f.bytes.length > 0) {
                await writable.write(f.bytes);
            }
        } finally {
            await writable.close();
        }
        done++;
        if (onProgress) onProgress({ done, total, label: f.path.join('/') });
    }
    return { packageRootName: PACKAGE_ROOT, written: done };
}
