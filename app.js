/* File Converter Tools - Main JS */

pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// ---------------------------------------------------
// TAB / PAGE NAVIGATION
// ---------------------------------------------------
const pageTitles = {
    'pdf-to-csv-converter': 'Free Hindi PDF to CSV Converter Online | FileShift',
    'image-to-text-ocr': 'Free Image to Text Converter - OCR Online | FileShift',
    'jpg-to-png-converter': 'Free JPG to PNG Converter Online | FileShift'
};

const pageDescriptions = {
    'pdf-to-csv-converter': 'Convert Hindi PDF to CSV online for free. Works with Kruti Dev, DevLys, Chanakya, Mangal fonts. OCR-based extraction, no uploads, runs in your browser.',
    'image-to-text-ocr': 'Free online image to text converter with OCR. Extract text from images in 50+ languages including Hindi, English, Bengali, Tamil, Telugu. Download as .doc file.',
    'jpg-to-png-converter': 'Convert JPG to PNG online for free. Lossless quality, transparency support, batch conversion. No software needed, runs in your browser.'
};

function switchTab(tabId) {
    document.querySelectorAll('.nav-tool').forEach(l => {
        l.classList.remove('active');
        l.setAttribute('aria-selected', 'false');
    });
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    const activeLink = document.querySelector(`.nav-tool[data-tab="${tabId}"]`);
    if (activeLink) {
        activeLink.classList.add('active');
        activeLink.setAttribute('aria-selected', 'true');
    }
    const page = document.getElementById(tabId);
    if (page) page.classList.add('active');
    // Update page title & meta description for SEO
    if (pageTitles[tabId]) document.title = pageTitles[tabId];
    if (pageDescriptions[tabId]) {
        const meta = document.querySelector('meta[name="description"]');
        if (meta) meta.setAttribute('content', pageDescriptions[tabId]);
    }
}

document.querySelectorAll('.nav-tool').forEach(link => {
    link.addEventListener('click', e => {
        e.preventDefault();
        const tab = link.dataset.tab;
        switchTab(tab);
        history.pushState(null, '', '#' + tab);
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
});

// Footer tool links
document.querySelectorAll('.footer-tool-link').forEach(link => {
    link.addEventListener('click', e => {
        e.preventDefault();
        const tab = link.dataset.tab;
        switchTab(tab);
        history.pushState(null, '', '#' + tab);
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
});

// Handle direct URL hash (e.g. fileshift.tools/#image-to-text-ocr)
window.addEventListener('DOMContentLoaded', () => {
    const hash = location.hash.replace('#', '');
    if (hash && document.getElementById(hash)) {
        switchTab(hash);
    }
});
window.addEventListener('popstate', () => {
    const hash = location.hash.replace('#', '');
    if (hash && document.getElementById(hash)) switchTab(hash);
});

// ---------------------------------------------------
// HELPERS
// ---------------------------------------------------
function formatBytes(b) {
    if (!b) return '0 B';
    const k = 1024;
    const s = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(b) / Math.log(k));
    return parseFloat((b / Math.pow(k, i)).toFixed(1)) + ' ' + s[i];
}

function showFileInfo(id, file, onRemove) {
    const el = document.getElementById(id);
    el.innerHTML = `
        <i class="fas fa-file"></i>
        <span class="file-name">${file.name}</span>
        <span class="file-size">${formatBytes(file.size)}</span>
        <button class="file-remove" title="Remove"><i class="fas fa-times"></i></button>
    `;
    el.classList.remove('hidden');
    el.querySelector('.file-remove').onclick = () => { el.classList.add('hidden'); if (onRemove) onRemove(); };
}

function progress(fillId, statusId, pct, msg) {
    const wrap = document.getElementById(fillId).closest('.progress-wrap');
    wrap.classList.remove('hidden');
    document.getElementById(fillId).style.width = pct + '%';
    document.getElementById(statusId).textContent = msg;
}

function hideProgress(fillId) {
    document.getElementById(fillId).closest('.progress-wrap').classList.add('hidden');
}

function setupDrop(zoneId, inputId, handler) {
    const zone = document.getElementById(zoneId);
    const inp = document.getElementById(inputId);

    zone.addEventListener('click', e => { if (e.target.tagName !== 'BUTTON') inp.click(); });

    ['dragenter', 'dragover'].forEach(ev =>
        zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add('dragover'); })
    );
    ['dragleave', 'drop'].forEach(ev =>
        zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.remove('dragover'); })
    );
    zone.addEventListener('drop', e => { if (e.dataTransfer.files.length) handler(e.dataTransfer.files); });
    inp.addEventListener('change', () => { if (inp.files.length) handler(inp.files); });
}

// ---------------------------------------------------
// 1) PDF TO CSV (Hindi OCR approach)
//    PDF → render pages as images → OCR with Hindi → CSV
// ---------------------------------------------------
let csvData = '';

setupDrop('pdf-drop-zone', 'pdf-input', files => handlePdf(files[0]));

async function handlePdf(file) {
    if (!file || file.type !== 'application/pdf') { alert('Please pick a PDF file.'); return; }

    showFileInfo('pdf-file-info', file, resetPdf);
    progress('pdf-progress-fill', 'pdf-status', 5, 'Loading PDF...');
    document.getElementById('pdf-preview').classList.add('hidden');

    try {
        const buf = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
        const total = pdf.numPages;
        const allText = [];

        for (let i = 1; i <= total; i++) {
            // Step 1: Render PDF page to canvas image
            progress('pdf-progress-fill', 'pdf-status',
                Math.round((i / total) * 40),
                `Rendering page ${i} of ${total} as image...`
            );

            const page = await pdf.getPage(i);
            const scale = 2; // 2x for better OCR quality
            const viewport = page.getViewport({ scale });
            const canvas = document.createElement('canvas');
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            const ctx = canvas.getContext('2d');

            await page.render({ canvasContext: ctx, viewport }).promise;

            // Step 2: Run OCR on the rendered image
            progress('pdf-progress-fill', 'pdf-status',
                40 + Math.round((i / total) * 50),
                `OCR scanning page ${i} of ${total} (Hindi + English)...`
            );

            const imgData = canvas.toDataURL('image/png');
            const result = await Tesseract.recognize(imgData, 'hin+eng', {
                logger: m => {
                    if (m.status === 'recognizing text') {
                        const pagePct = Math.round(m.progress * 100);
                        const overallPct = 40 + Math.round(((i - 1 + m.progress) / total) * 50);
                        progress('pdf-progress-fill', 'pdf-status', overallPct,
                            `OCR page ${i}/${total} - ${pagePct}%`);
                    }
                }
            });

            // Collect recognized text lines
            const lines = result.data.text.split('\n').map(l => l.trim()).filter(Boolean);
            lines.forEach(line => allText.push(line));

            // Clean up canvas
            canvas.width = 0;
            canvas.height = 0;
        }

        if (!allText.length) {
            progress('pdf-progress-fill', 'pdf-status', 100, 'No text could be recognized in this PDF.');
            return;
        }

        // Step 3: Build rows for CSV
        // Split each line by common separators (tabs, multiple spaces, pipes)
        const rows = allText.map(line => {
            // Try tab first, then multiple spaces, then keep as single column
            if (line.includes('\t')) return line.split('\t').map(c => c.trim());
            const parts = line.split(/\s{2,}/).map(c => c.trim()).filter(Boolean);
            return parts.length > 1 ? parts : [line];
        });

        const maxCols = Math.max(...rows.map(r => r.length));
        rows.forEach(r => { while (r.length < maxCols) r.push(''); });

        csvData = rows.map(r => r.map(c => `"${c.replace(/"/g, '""')}"`).join(',')).join('\n');

        // Table preview
        const table = document.getElementById('pdf-table');
        table.innerHTML = '';
        rows.slice(0, 80).forEach((row, idx) => {
            const tr = document.createElement('tr');
            row.forEach(c => {
                const el = document.createElement(idx === 0 ? 'th' : 'td');
                el.textContent = c;
                tr.appendChild(el);
            });
            table.appendChild(tr);
        });
        if (rows.length > 80) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = maxCols;
            td.style.textAlign = 'center';
            td.style.color = '#999';
            td.textContent = `+ ${rows.length - 80} more rows`;
            tr.appendChild(td);
            table.appendChild(tr);
        }

        progress('pdf-progress-fill', 'pdf-status', 100, `Done - ${rows.length} rows extracted from ${total} pages.`);
        document.getElementById('pdf-preview').classList.remove('hidden');

    } catch (err) {
        console.error(err);
        progress('pdf-progress-fill', 'pdf-status', 100, 'Something went wrong: ' + err.message);
    }
}

document.getElementById('pdf-download-btn').onclick = () => {
    if (!csvData) return;
    // UTF-8 BOM so Excel opens Hindi/Unicode text properly
    const bom = '\uFEFF';
    saveAs(new Blob([bom + csvData], { type: 'text/csv;charset=utf-8;' }), 'converted.csv');
};

function resetPdf() {
    csvData = '';
    document.getElementById('pdf-input').value = '';
    document.getElementById('pdf-preview').classList.add('hidden');
    hideProgress('pdf-progress-fill');
}

// ---------------------------------------------------
// 2) IMAGE TO TEXT (OCR)
// ---------------------------------------------------
let ocrText = '';

setupDrop('img-drop-zone', 'img-input', files => handleOcr(files[0]));

async function handleOcr(file) {
    if (!file || !file.type.startsWith('image/')) { alert('Please pick an image file.'); return; }

    showFileInfo('img-file-info', file, resetOcr);
    progress('ocr-progress-fill', 'ocr-status', 5, 'Starting OCR engine...');
    document.getElementById('img-preview').classList.add('hidden');

    document.getElementById('ocr-image-preview').src = URL.createObjectURL(file);

    try {
        // Get selected languages from checkboxes
        const checkedLangs = document.querySelectorAll('input[name="ocr-lang"]:checked');
        const selectedLangs = Array.from(checkedLangs).map(cb => cb.value);
        const langs = selectedLangs.length ? selectedLangs.join('+') : 'hin+eng';

        const langNames = Array.from(checkedLangs).map(cb => cb.parentElement.querySelector('.lang-chip').textContent);
        progress('ocr-progress-fill', 'ocr-status', 8,
            `Loading language data for: ${langNames.join(', ')}...`);

        const result = await Tesseract.recognize(file, langs, {
            logger: m => {
                if (m.status === 'loading tesseract core') {
                    progress('ocr-progress-fill', 'ocr-status', 10, 'Loading OCR engine...');
                } else if (m.status === 'initializing tesseract') {
                    progress('ocr-progress-fill', 'ocr-status', 15, 'Initializing...');
                } else if (m.status === 'loading language traineddata') {
                    progress('ocr-progress-fill', 'ocr-status', 20, 'Downloading language data...');
                } else if (m.status === 'recognizing text') {
                    const pct = 30 + Math.round(m.progress * 70);
                    progress('ocr-progress-fill', 'ocr-status', pct, `Recognizing text... ${pct}%`);
                }
            }
        });

        ocrText = result.data.text;
        const conf = Math.round(result.data.confidence);
        const words = ocrText.split(/\s+/).filter(Boolean).length;

        document.getElementById('ocr-text-output').value = ocrText;
        document.getElementById('ocr-stats').innerHTML = `
            <span><i class="fas fa-chart-bar"></i> Confidence: ${conf}%</span>
            <span><i class="fas fa-font"></i> Words: ${words}</span>
            <span><i class="fas fa-text-width"></i> Characters: ${ocrText.length}</span>
        `;

        progress('ocr-progress-fill', 'ocr-status', 100, 'Done!');
        document.getElementById('img-preview').classList.remove('hidden');

    } catch (err) {
        console.error(err);
        progress('ocr-progress-fill', 'ocr-status', 100, 'Something went wrong: ' + err.message);
    }
}

document.getElementById('ocr-download-btn').onclick = () => {
    if (!ocrText) return;
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>Extracted Text</title>
<style>body{font-family:Calibri,sans-serif;font-size:12pt;line-height:1.6;padding:40px}</style>
</head><body>${ocrText.replace(/\n/g, '<br>')}</body></html>`;
    saveAs(new Blob(['\ufeff' + html], { type: 'application/msword' }), 'extracted-text.doc');
};

function resetOcr() {
    ocrText = '';
    document.getElementById('img-input').value = '';
    document.getElementById('img-preview').classList.add('hidden');
    hideProgress('ocr-progress-fill');
}

// ---------------------------------------------------
// 3) JPG TO PNG
// ---------------------------------------------------
let pngBlobs = [];

setupDrop('jpg-drop-zone', 'jpg-input', handleJpg);

async function handleJpg(fileList) {
    const files = Array.from(fileList).filter(f => f.type === 'image/jpeg');
    if (!files.length) { alert('Please pick JPG/JPEG files.'); return; }

    const info = files.length === 1
        ? files[0]
        : { name: `${files.length} JPG files`, size: files.reduce((a, f) => a + f.size, 0) };
    showFileInfo('jpg-file-info', info, resetJpg);

    progress('jpg-progress-fill', 'jpg-status', 5, 'Converting...');
    document.getElementById('jpg-preview').classList.add('hidden');

    pngBlobs = [];
    const grid = document.getElementById('jpg-results');
    grid.innerHTML = '';

    for (let i = 0; i < files.length; i++) {
        progress('jpg-progress-fill', 'jpg-status',
            Math.round(((i + 1) / files.length) * 95),
            `Converting ${i + 1} of ${files.length}...`
        );

        const blob = await jpgToPng(files[i]);
        const name = files[i].name.replace(/\.jpe?g$/i, '.png');
        pngBlobs.push({ name, blob });

        const card = document.createElement('div');
        card.className = 'img-card';
        const url = URL.createObjectURL(blob);
        card.innerHTML = `
            <img src="${url}" alt="${name}">
            <div class="img-card-foot">
                <span class="img-name" title="${name}">${name}</span>
                <button class="dl-one" title="Download"><i class="fas fa-download"></i></button>
            </div>
        `;
        card.querySelector('.dl-one').onclick = () => saveAs(blob, name);
        grid.appendChild(card);
    }

    progress('jpg-progress-fill', 'jpg-status', 100,
        `Done - ${files.length} image${files.length > 1 ? 's' : ''} converted.`);
    document.getElementById('jpg-preview').classList.remove('hidden');
}

function jpgToPng(file) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const c = document.createElement('canvas');
            c.width = img.naturalWidth;
            c.height = img.naturalHeight;
            c.getContext('2d').drawImage(img, 0, 0);
            c.toBlob(b => b ? resolve(b) : reject(new Error('Conversion failed')), 'image/png');
            URL.revokeObjectURL(img.src);
        };
        img.onerror = () => reject(new Error('Could not load image'));
        img.src = URL.createObjectURL(file);
    });
}

document.getElementById('jpg-download-all-btn').onclick = async () => {
    if (!pngBlobs.length) return;
    if (pngBlobs.length === 1) { saveAs(pngBlobs[0].blob, pngBlobs[0].name); return; }
    const zip = new JSZip();
    pngBlobs.forEach(p => zip.file(p.name, p.blob));
    const z = await zip.generateAsync({ type: 'blob' });
    saveAs(z, 'converted-images.zip');
};

function resetJpg() {
    pngBlobs = [];
    document.getElementById('jpg-input').value = '';
    document.getElementById('jpg-preview').classList.add('hidden');
    document.getElementById('jpg-results').innerHTML = '';
    hideProgress('jpg-progress-fill');
}
