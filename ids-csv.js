// Generate the IDS Package CSV files for e-Signage Lite.
// All CSVs use CRLF line endings (including trailing CRLF) and ASCII content.
// Field formats are reverse-engineered from a working sample; see TEMPLATE_SPEC.md.

const CRLF = '\r\n';
const SCHEDULE_DAYS = 366;

function pad2(n) { return String(n).padStart(2, '0'); }

function formatDateSlash(d) {
    return `${d.getFullYear()}/${pad2(d.getMonth() + 1)}/${pad2(d.getDate())}`;
}

function formatDateTimeSlash(d) {
    return `${formatDateSlash(d)} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
}

// Filename inside contents/ for the .ppt: "M.D.ppt" (no leading zeros) for single,
// "M.D-N.ppt" for multi-slide.
export function makePptFilename(date, index, total) {
    const base = `${date.getMonth() + 1}.${date.getDate()}`;
    if (total <= 1) return `${base}.ppt`;
    return `${base}-${index + 1}.ppt`;
}

function packageIds() {
    return `Schedule.csv${CRLF}`;
}

function scheduleCsv(startDate) {
    const lines = ['e-Signage Lite', 'false'];
    const cursor = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    for (let i = 0; i < SCHEDULE_DAYS; i++) {
        const flag = i === 0 ? 'true' : 'false';
        lines.push(`${formatDateSlash(cursor)},TimeTable.csv,${flag}`);
        cursor.setDate(cursor.getDate() + 1);
    }
    return lines.join(CRLF) + CRLF;
}

function timeTableCsv() {
    const lines = [
        'e-Signage Lite',
        '0xFFFFFF',
        '00:00:00,23:59:59,0,Program_0000_2400,0,',
    ];
    return lines.join(CRLF) + CRLF;
}

function programCsv() {
    const lines = [
        'e-Signage Lite Program 00:00-24:00',
        '0xFFFFFF,0,0,1920,1080,00:00:00,0',
        '0,playlist_0.csv,0,0,1920,1080,0x000000,false,0,true',
        '1,playlist_1.csv,0,0,1920,1080,0x000000,false,0,true',
        '2,playlist_2.csv,0,0,1,1,0x000000,false,0,false',
        '3,playlist_3.csv,0,0,1,1,0x000000,false,0,false',
        '4,playlist_4.csv,0,0,1,1,0x000000,false,0,false',
        '5,playlist_5.csv,0,0,1,1,0x000000,false,0,false',
        '6,playlist_6.csv,0,0,1,1,0x000000,false,0,false',
        '7,playlist_7.csv,0,0,0,0,0x000000,false,0,true',
        '8,playlist_8.csv,0,0,1,1,0x000000,false,0,false',
    ];
    return lines.join(CRLF) + CRLF;
}

// playlist_1.csv: one reference per content
function playlist1Csv(contentCount) {
    const lines = [];
    for (let i = 1; i <= contentCount; i++) {
        const idx = String(i).padStart(3, '0');
        lines.push(`information\\\\contents_1_${idx}.csv`);
    }
    return lines.join(CRLF) + CRLF;
}

// information/contents_1_NNN.csv: reference to the ppt + playback parameters
function contentCsv(pptFilename, displaySeconds) {
    const dispStr = secondsToHMS(displaySeconds);
    const lines = [
        `..\\\\contents\\\\${pptFilename}`,
        `1,${dispStr},,false,0x000000,0,,0,0,0,100,-2147483647,2000,1,-1,-1,-1,-1,11`,
    ];
    return lines.join(CRLF) + CRLF;
}

function secondsToHMS(sec) {
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    const s = sec % 60;
    return `${pad2(h)}:${pad2(m)}:${pad2(s)}`;
}

// Catalog.csv: one line per ppt file
function catalogCsv(pptEntries, timestamp) {
    const tsStr = formatDateTimeSlash(timestamp);
    const lines = pptEntries.map(e =>
        `PW_DPC01,${e.filename},${tsStr},${e.size},${e.filename}`
    );
    return lines.join(CRLF) + CRLF;
}

const EMPTY = new Uint8Array(0);

function encodeText(s) {
    return new TextEncoder().encode(s);
}

// Build the complete IDS Package file tree.
//
// args:
//   pptFiles: [{ filename: '5.16.ppt', bytes: Uint8Array, displaySeconds: 30 }, ...]
//   scheduleStartDate: Date (typically today; today gets `true`, rest `false`)
//   timestamp: Date (used for Catalog.csv timestamps)
//
// returns: array of { path: string[], bytes: Uint8Array }
export function buildIdsPackage({ pptFiles, scheduleStartDate, timestamp }) {
    if (!pptFiles || pptFiles.length === 0) throw new Error('pptFiles is empty');
    const files = [];

    files.push({ path: ['Package.ids'], bytes: encodeText(packageIds()) });
    files.push({ path: ['Schedule.csv'], bytes: encodeText(scheduleCsv(scheduleStartDate)) });
    files.push({ path: ['TimeTable.csv'], bytes: encodeText(timeTableCsv()) });

    files.push({ path: ['Program_0000_2400', 'program.csv'], bytes: encodeText(programCsv()) });
    // playlist_0 and 2..8 are present but empty in the sample
    for (const idx of [0, 2, 3, 4, 5, 6, 7, 8]) {
        files.push({ path: ['Program_0000_2400', `playlist_${idx}.csv`], bytes: EMPTY });
    }
    files.push({ path: ['Program_0000_2400', 'playlist_1.csv'], bytes: encodeText(playlist1Csv(pptFiles.length)) });

    pptFiles.forEach((p, i) => {
        const idx = String(i + 1).padStart(3, '0');
        files.push({
            path: ['Program_0000_2400', 'information', `contents_1_${idx}.csv`],
            bytes: encodeText(contentCsv(p.filename, p.displaySeconds)),
        });
    });

    pptFiles.forEach(p => {
        files.push({ path: ['contents', p.filename], bytes: p.bytes });
    });
    const catalogEntries = pptFiles.map(p => ({ filename: p.filename, size: p.bytes.length }));
    files.push({ path: ['contents', 'Catalog.csv'], bytes: encodeText(catalogCsv(catalogEntries, timestamp)) });

    return files;
}
