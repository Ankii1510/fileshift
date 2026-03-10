/* File Converter Tools - Main JS */

pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// ---------------------------------------------------
// TAB / PAGE NAVIGATION
// ---------------------------------------------------
const pageTitles = {
    'pdf-to-csv-converter': 'Free PDF to CSV Converter Online | FileFlipr',
    'image-to-text-ocr': 'Free Image to Text Converter - OCR Online | FileFlipr',
    'jpg-to-png-converter': 'Free JPG to PNG Converter Online | FileFlipr',
    'pdf-to-jpg-converter': 'Free PDF to JPG Converter Online | FileFlipr',
    'jpg-to-pdf-converter': 'Free JPG to PDF Converter Online | FileFlipr',
    'pdf-to-word-converter': 'Free PDF to Word Converter Online | FileFlipr',
    'word-to-pdf-converter': 'Free Word to PDF Converter Online | FileFlipr',
    'csv-to-excel-converter': 'Free CSV to Excel Converter Online | FileFlipr',
    'pdf-to-excel-converter': 'Free PDF to Excel Converter Online | FileFlipr',
    'excel-to-pdf-converter': 'Free Excel to PDF Converter Online | FileFlipr'
};

const pageDescriptions = {
    'pdf-to-csv-converter': 'Convert PDF to CSV online for free. OCR-based extraction supports all languages and fonts. No uploads, runs in your browser.',
    'image-to-text-ocr': 'Free online image to text converter with OCR. Extract text from images in 50+ languages including Hindi, English, Bengali, Tamil, Telugu. Download as .doc file.',
    'jpg-to-png-converter': 'Convert JPG to PNG online for free. Lossless quality, transparency support, batch conversion. No software needed, runs in your browser.',
    'pdf-to-jpg-converter': 'Convert PDF to JPG online for free. Each page becomes a high-quality JPG image. Download individually or as ZIP. Runs in your browser.',
    'jpg-to-pdf-converter': 'Convert JPG to PDF online for free. Combine multiple JPG images into a single PDF document. No uploads, runs in your browser.',
    'pdf-to-word-converter': 'Convert PDF to Word online for free. Extract text from PDF and download as editable .doc file. Runs in your browser.',
    'word-to-pdf-converter': 'Convert Word to PDF online for free. Upload .docx file and get a clean PDF document. Runs in your browser.',
    'csv-to-excel-converter': 'Convert CSV to Excel online for free. Upload CSV and download a proper .xlsx spreadsheet. Runs in your browser.',
    'pdf-to-excel-converter': 'Convert PDF to Excel online for free. Extract tables from PDF using OCR and download as .xlsx. Runs in your browser.',
    'excel-to-pdf-converter': 'Convert Excel to PDF online for free. Upload .xlsx and get a clean PDF with formatted tables. Runs in your browser.'
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

// Handle direct URL hash (e.g. fileflipr.com/#image-to-text-ocr)
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
    const safeName = document.createElement('span');
    safeName.textContent = file.name;
    el.innerHTML = '';
    const icon = document.createElement('i');
    icon.className = 'fas fa-file';
    const nameSpan = document.createElement('span');
    nameSpan.className = 'file-name';
    nameSpan.textContent = file.name;
    const sizeSpan = document.createElement('span');
    sizeSpan.className = 'file-size';
    sizeSpan.textContent = formatBytes(file.size);
    const removeBtn = document.createElement('button');
    removeBtn.className = 'file-remove';
    removeBtn.title = 'Remove';
    removeBtn.innerHTML = '<i class="fas fa-times"></i>';
    removeBtn.onclick = () => { el.classList.add('hidden'); if (onRemove) onRemove(); };
    el.appendChild(icon);
    el.appendChild(nameSpan);
    el.appendChild(sizeSpan);
    el.appendChild(removeBtn);
    el.classList.remove('hidden');
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
//    PDF â†’ render pages as images â†’ OCR with Hindi â†’ CSV
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

// ---------------------------------------------------
// 4) PDF TO JPG
// ---------------------------------------------------
let p2jBlobs = [];

setupDrop('p2j-drop-zone', 'p2j-input', files => handleP2j(files[0]));

async function handleP2j(file) {
    if (!file || file.type !== 'application/pdf') { alert('Please pick a PDF file.'); return; }

    showFileInfo('p2j-file-info', file, resetP2j);
    progress('p2j-progress-fill', 'p2j-status', 5, 'Loading PDF...');
    document.getElementById('p2j-preview').classList.add('hidden');

    try {
        const buf = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
        const total = pdf.numPages;

        p2jBlobs = [];
        const grid = document.getElementById('p2j-results');
        grid.innerHTML = '';

        for (let i = 1; i <= total; i++) {
            progress('p2j-progress-fill', 'p2j-status',
                Math.round((i / total) * 90),
                `Rendering page ${i} of ${total}...`
            );

            const page = await pdf.getPage(i);
            const scale = 2;
            const viewport = page.getViewport({ scale });
            const canvas = document.createElement('canvas');
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            const ctx = canvas.getContext('2d');
            await page.render({ canvasContext: ctx, viewport }).promise;

            const blob = await new Promise(resolve =>
                canvas.toBlob(b => resolve(b), 'image/jpeg', 0.92)
            );
            const name = `page-${i}.jpg`;
            p2jBlobs.push({ name, blob });

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

            canvas.width = 0;
            canvas.height = 0;
        }

        progress('p2j-progress-fill', 'p2j-status', 100,
            `Done - ${total} page${total > 1 ? 's' : ''} converted.`);
        document.getElementById('p2j-preview').classList.remove('hidden');

    } catch (err) {
        console.error(err);
        progress('p2j-progress-fill', 'p2j-status', 100, 'Something went wrong: ' + err.message);
    }
}

document.getElementById('p2j-download-all-btn').onclick = async () => {
    if (!p2jBlobs.length) return;
    if (p2jBlobs.length === 1) { saveAs(p2jBlobs[0].blob, p2jBlobs[0].name); return; }
    const zip = new JSZip();
    p2jBlobs.forEach(p => zip.file(p.name, p.blob));
    const z = await zip.generateAsync({ type: 'blob' });
    saveAs(z, 'pdf-pages.zip');
};

function resetP2j() {
    p2jBlobs = [];
    document.getElementById('p2j-input').value = '';
    document.getElementById('p2j-preview').classList.add('hidden');
    document.getElementById('p2j-results').innerHTML = '';
    hideProgress('p2j-progress-fill');
}

// ---------------------------------------------------
// 5) JPG TO PDF
// ---------------------------------------------------
let j2pFiles = [];

setupDrop('j2p-drop-zone', 'j2p-input', handleJ2p);

async function handleJ2p(fileList) {
    const files = Array.from(fileList).filter(f => f.type === 'image/jpeg');
    if (!files.length) { alert('Please pick JPG/JPEG files.'); return; }

    const info = files.length === 1
        ? files[0]
        : { name: `${files.length} JPG files`, size: files.reduce((a, f) => a + f.size, 0) };
    showFileInfo('j2p-file-info', info, resetJ2p);

    progress('j2p-progress-fill', 'j2p-status', 5, 'Reading images...');
    document.getElementById('j2p-preview').classList.add('hidden');

    j2pFiles = files;
    const grid = document.getElementById('j2p-results');
    grid.innerHTML = '';

    for (let i = 0; i < files.length; i++) {
        progress('j2p-progress-fill', 'j2p-status',
            Math.round(((i + 1) / files.length) * 90),
            `Loading image ${i + 1} of ${files.length}...`
        );

        const url = URL.createObjectURL(files[i]);
        const card = document.createElement('div');
        card.className = 'img-card';
        card.innerHTML = `
            <img src="${url}" alt="Page ${i + 1}">
            <div class="img-card-foot">
                <span class="img-name" title="${files[i].name}">Page ${i + 1}</span>
            </div>
        `;
        grid.appendChild(card);
    }

    progress('j2p-progress-fill', 'j2p-status', 100,
        `${files.length} image${files.length > 1 ? 's' : ''} ready. Click Download PDF.`);
    document.getElementById('j2p-preview').classList.remove('hidden');
}

document.getElementById('j2p-download-btn').onclick = async () => {
    if (!j2pFiles.length) return;

    progress('j2p-progress-fill', 'j2p-status', 10, 'Building PDF...');
    const { jsPDF } = window.jspdf;

    let pdf = null;

    for (let i = 0; i < j2pFiles.length; i++) {
        progress('j2p-progress-fill', 'j2p-status',
            10 + Math.round(((i + 1) / j2pFiles.length) * 85),
            `Adding page ${i + 1} of ${j2pFiles.length}...`
        );

        const imgData = await readFileAsDataURL(j2pFiles[i]);
        const dims = await getImageDimensions(imgData);

        const orientation = dims.width > dims.height ? 'landscape' : 'portrait';
        const pageWidth = orientation === 'landscape' ? 297 : 210;
        const pageHeight = orientation === 'landscape' ? 210 : 297;

        // Fit image to page with margins
        const margin = 0;
        const maxW = pageWidth - margin * 2;
        const maxH = pageHeight - margin * 2;
        const ratio = Math.min(maxW / dims.width, maxH / dims.height);
        const w = dims.width * ratio;
        const h = dims.height * ratio;
        const x = (pageWidth - w) / 2;
        const y = (pageHeight - h) / 2;

        if (i === 0) {
            pdf = new jsPDF({ orientation, unit: 'mm', format: 'a4' });
        } else {
            pdf.addPage('a4', orientation);
        }
        pdf.addImage(imgData, 'JPEG', x, y, w, h);
    }

    progress('j2p-progress-fill', 'j2p-status', 100, 'Done!');
    pdf.save('combined.pdf');
};

function readFileAsDataURL(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error('Could not read file'));
        reader.readAsDataURL(file);
    });
}

function getImageDimensions(dataUrl) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
        img.onerror = () => reject(new Error('Could not load image'));
        img.src = dataUrl;
    });
}

function resetJ2p() {
    j2pFiles = [];
    document.getElementById('j2p-input').value = '';
    document.getElementById('j2p-preview').classList.add('hidden');
    document.getElementById('j2p-results').innerHTML = '';
    hideProgress('j2p-progress-fill');
}

// ---------------------------------------------------
// 6) PDF TO WORD
// ---------------------------------------------------
let p2wText = '';

setupDrop('p2w-drop-zone', 'p2w-input', files => handleP2w(files[0]));

async function handleP2w(file) {
    if (!file || file.type !== 'application/pdf') { alert('Please pick a PDF file.'); return; }

    showFileInfo('p2w-file-info', file, resetP2w);
    progress('p2w-progress-fill', 'p2w-status', 5, 'Loading PDF...');
    document.getElementById('p2w-preview').classList.add('hidden');

    try {
        const buf = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
        const total = pdf.numPages;
        const allText = [];

        // Try text layer extraction first
        let hasTextLayer = false;
        for (let i = 1; i <= total; i++) {
            progress('p2w-progress-fill', 'p2w-status',
                Math.round((i / total) * 50),
                `Extracting text from page ${i} of ${total}...`
            );
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            const pageText = content.items.map(item => item.str).join(' ').trim();
            if (pageText) { hasTextLayer = true; allText.push(pageText); }
        }

        // If no text layer, fall back to OCR
        if (!hasTextLayer) {
            allText.length = 0;
            for (let i = 1; i <= total; i++) {
                progress('p2w-progress-fill', 'p2w-status',
                    Math.round((i / total) * 40),
                    `Rendering page ${i} of ${total} for OCR...`
                );
                const page = await pdf.getPage(i);
                const scale = 2;
                const viewport = page.getViewport({ scale });
                const canvas = document.createElement('canvas');
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;

                progress('p2w-progress-fill', 'p2w-status',
                    40 + Math.round((i / total) * 55),
                    `OCR scanning page ${i} of ${total}...`
                );
                const imgData = canvas.toDataURL('image/png');
                const result = await Tesseract.recognize(imgData, 'hin+eng');
                const lines = result.data.text.trim();
                if (lines) allText.push(lines);
                canvas.width = 0; canvas.height = 0;
            }
        }

        if (!allText.length) {
            progress('p2w-progress-fill', 'p2w-status', 100, 'No text could be extracted from this PDF.');
            return;
        }

        p2wText = allText.join('\n\n');
        const preview = p2wText.length > 3000 ? p2wText.substring(0, 3000) + '\n\n... (preview truncated)' : p2wText;
        document.getElementById('p2w-text-preview').textContent = preview;
        progress('p2w-progress-fill', 'p2w-status', 100, `Done - ${total} pages extracted.`);
        document.getElementById('p2w-preview').classList.remove('hidden');

    } catch (err) {
        console.error(err);
        progress('p2w-progress-fill', 'p2w-status', 100, 'Something went wrong: ' + err.message);
    }
}

document.getElementById('p2w-download-btn').onclick = () => {
    if (!p2wText) return;
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>Converted Document</title>
<style>body{font-family:Calibri,sans-serif;font-size:12pt;line-height:1.8;padding:40px 60px}p{margin-bottom:12px}</style>
</head><body>${p2wText.split('\n').map(line => '<p>' + line.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') + '</p>').join('')}</body></html>`;
    saveAs(new Blob(['\ufeff' + html], { type: 'application/msword' }), 'converted.doc');
};

function resetP2w() {
    p2wText = '';
    document.getElementById('p2w-input').value = '';
    document.getElementById('p2w-preview').classList.add('hidden');
    hideProgress('p2w-progress-fill');
}

// ---------------------------------------------------
// 7) WORD TO PDF
// ---------------------------------------------------
let w2pHtml = '';

setupDrop('w2p-drop-zone', 'w2p-input', files => handleW2p(files[0]));

async function handleW2p(file) {
    if (!file) { alert('Please pick a Word file.'); return; }
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (ext !== '.docx' && ext !== '.doc') { alert('Please pick a .docx file.'); return; }

    showFileInfo('w2p-file-info', file, resetW2p);
    progress('w2p-progress-fill', 'w2p-status', 10, 'Reading Word document...');
    document.getElementById('w2p-preview').classList.add('hidden');

    try {
        const buf = await file.arrayBuffer();
        progress('w2p-progress-fill', 'w2p-status', 40, 'Converting to viewable format...');

        const result = await mammoth.convertToHtml({ arrayBuffer: buf });
        w2pHtml = result.value;

        const previewEl = document.getElementById('w2p-html-preview');
        previewEl.innerHTML = w2pHtml;

        progress('w2p-progress-fill', 'w2p-status', 100, 'Done! Preview shown below.');
        document.getElementById('w2p-preview').classList.remove('hidden');

    } catch (err) {
        console.error(err);
        progress('w2p-progress-fill', 'w2p-status', 100, 'Something went wrong: ' + err.message);
    }
}

document.getElementById('w2p-download-btn').onclick = async () => {
    if (!w2pHtml) return;
    progress('w2p-progress-fill', 'w2p-status', 10, 'Building PDF...');

    const container = document.getElementById('w2p-html-preview');
    const origWidth = container.style.width;
    container.style.width = '700px';

    try {
        const canvas = await html2canvas(container, { scale: 2, useCORS: true, windowWidth: 700 });
        container.style.width = origWidth;

        const { jsPDF } = window.jspdf;
        const pageWidth = 210;
        const pageHeight = 297;
        const margin = 10;
        const contentWidth = pageWidth - margin * 2;
        const imgHeight = (canvas.height * contentWidth) / canvas.width;
        const contentHeight = pageHeight - margin * 2;

        const pdf = new jsPDF('p', 'mm', 'a4');
        let y = 0;
        let page = 0;

        while (y < imgHeight) {
            if (page > 0) pdf.addPage();

            const sourceY = (y / imgHeight) * canvas.height;
            const sourceH = Math.min((contentHeight / imgHeight) * canvas.height, canvas.height - sourceY);
            const destH = (sourceH / canvas.height) * imgHeight;

            const pageCanvas = document.createElement('canvas');
            pageCanvas.width = canvas.width;
            pageCanvas.height = Math.ceil(sourceH);
            const ctx = pageCanvas.getContext('2d');
            ctx.drawImage(canvas, 0, sourceY, canvas.width, sourceH, 0, 0, canvas.width, sourceH);

            const pageImg = pageCanvas.toDataURL('image/jpeg', 0.95);
            pdf.addImage(pageImg, 'JPEG', margin, margin, contentWidth, destH);

            y += contentHeight;
            page++;
        }

        progress('w2p-progress-fill', 'w2p-status', 100, 'Done!');
        pdf.save('converted.pdf');
    } catch (err) {
        container.style.width = origWidth;
        console.error(err);
        progress('w2p-progress-fill', 'w2p-status', 100, 'Something went wrong: ' + err.message);
    }
};

function resetW2p() {
    w2pHtml = '';
    document.getElementById('w2p-input').value = '';
    document.getElementById('w2p-preview').classList.add('hidden');
    document.getElementById('w2p-html-preview').innerHTML = '';
    hideProgress('w2p-progress-fill');
}

// ---------------------------------------------------
// 8) CSV TO EXCEL
// ---------------------------------------------------
let c2eRows = [];

setupDrop('c2e-drop-zone', 'c2e-input', files => handleC2e(files[0]));

function parseCSV(text) {
    const rows = [];
    let current = [];
    let cell = '';
    let inQuote = false;

    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (inQuote) {
            if (ch === '"') {
                if (text[i + 1] === '"') { cell += '"'; i++; }
                else inQuote = false;
            } else {
                cell += ch;
            }
        } else {
            if (ch === '"') { inQuote = true; }
            else if (ch === ',' || ch === '\t') { current.push(cell.trim()); cell = ''; }
            else if (ch === '\n' || (ch === '\r' && text[i + 1] === '\n')) {
                if (ch === '\r') i++;
                current.push(cell.trim());
                cell = '';
                if (current.some(c => c)) rows.push(current);
                current = [];
            } else {
                cell += ch;
            }
        }
    }
    current.push(cell.trim());
    if (current.some(c => c)) rows.push(current);
    return rows;
}

async function handleC2e(file) {
    if (!file) { alert('Please pick a CSV file.'); return; }

    showFileInfo('c2e-file-info', file, resetC2e);
    progress('c2e-progress-fill', 'c2e-status', 10, 'Reading CSV...');
    document.getElementById('c2e-preview').classList.add('hidden');

    try {
        const text = await file.text();
        progress('c2e-progress-fill', 'c2e-status', 40, 'Parsing data...');

        const rows = parseCSV(text);
        if (!rows.length) {
            progress('c2e-progress-fill', 'c2e-status', 100, 'No data found in CSV file.');
            return;
        }

        // Normalize column count
        const maxCols = Math.max(...rows.map(r => r.length));
        rows.forEach(r => { while (r.length < maxCols) r.push(''); });

        // Preview table
        const table = document.getElementById('c2e-table');
        table.innerHTML = '';
        rows.slice(0, 50).forEach((row, idx) => {
            const tr = document.createElement('tr');
            row.forEach(c => {
                const el = document.createElement(idx === 0 ? 'th' : 'td');
                el.textContent = c;
                tr.appendChild(el);
            });
            table.appendChild(tr);
        });
        if (rows.length > 50) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = maxCols;
            td.style.textAlign = 'center';
            td.style.color = '#999';
            td.textContent = `+ ${rows.length - 50} more rows`;
            tr.appendChild(td);
            table.appendChild(tr);
        }

        c2eRows = rows;
        progress('c2e-progress-fill', 'c2e-status', 100, `Done - ${rows.length} rows found.`);
        document.getElementById('c2e-preview').classList.remove('hidden');

    } catch (err) {
        console.error(err);
        progress('c2e-progress-fill', 'c2e-status', 100, 'Something went wrong: ' + err.message);
    }
}

document.getElementById('c2e-download-btn').onclick = () => {
    if (!c2eRows.length) return;
    const ws = XLSX.utils.aoa_to_sheet(c2eRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'converted.xlsx');
};

function resetC2e() {
    c2eRows = [];
    document.getElementById('c2e-input').value = '';
    document.getElementById('c2e-preview').classList.add('hidden');
    document.getElementById('c2e-table').innerHTML = '';
    hideProgress('c2e-progress-fill');
}

// ---------------------------------------------------
// 9) PDF TO EXCEL
// ---------------------------------------------------
let p2eRows = [];

setupDrop('p2e-drop-zone', 'p2e-input', files => handleP2e(files[0]));

async function handleP2e(file) {
    if (!file || file.type !== 'application/pdf') { alert('Please pick a PDF file.'); return; }

    showFileInfo('p2e-file-info', file, resetP2e);
    progress('p2e-progress-fill', 'p2e-status', 5, 'Loading PDF...');
    document.getElementById('p2e-preview').classList.add('hidden');

    try {
        const buf = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
        const total = pdf.numPages;
        const allText = [];

        for (let i = 1; i <= total; i++) {
            progress('p2e-progress-fill', 'p2e-status',
                Math.round((i / total) * 40),
                `Rendering page ${i} of ${total} as image...`
            );

            const page = await pdf.getPage(i);
            const scale = 2;
            const viewport = page.getViewport({ scale });
            const canvas = document.createElement('canvas');
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;

            progress('p2e-progress-fill', 'p2e-status',
                40 + Math.round((i / total) * 50),
                `OCR scanning page ${i} of ${total}...`
            );

            const imgData = canvas.toDataURL('image/png');
            const result = await Tesseract.recognize(imgData, 'hin+eng', {
                logger: m => {
                    if (m.status === 'recognizing text') {
                        const overallPct = 40 + Math.round(((i - 1 + m.progress) / total) * 50);
                        progress('p2e-progress-fill', 'p2e-status', overallPct,
                            `OCR page ${i}/${total} - ${Math.round(m.progress * 100)}%`);
                    }
                }
            });

            const lines = result.data.text.split('\n').map(l => l.trim()).filter(Boolean);
            lines.forEach(line => allText.push(line));
            canvas.width = 0; canvas.height = 0;
        }

        if (!allText.length) {
            progress('p2e-progress-fill', 'p2e-status', 100, 'No text could be recognized in this PDF.');
            return;
        }

        // Build rows
        const rows = allText.map(line => {
            if (line.includes('\t')) return line.split('\t').map(c => c.trim());
            const parts = line.split(/\s{2,}/).map(c => c.trim()).filter(Boolean);
            return parts.length > 1 ? parts : [line];
        });

        const maxCols = Math.max(...rows.map(r => r.length));
        rows.forEach(r => { while (r.length < maxCols) r.push(''); });

        // Table preview
        const table = document.getElementById('p2e-table');
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

        p2eRows = rows;
        progress('p2e-progress-fill', 'p2e-status', 100, `Done - ${rows.length} rows extracted from ${total} pages.`);
        document.getElementById('p2e-preview').classList.remove('hidden');

    } catch (err) {
        console.error(err);
        progress('p2e-progress-fill', 'p2e-status', 100, 'Something went wrong: ' + err.message);
    }
}

document.getElementById('p2e-download-btn').onclick = () => {
    if (!p2eRows.length) return;
    const ws = XLSX.utils.aoa_to_sheet(p2eRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'converted.xlsx');
};

function resetP2e() {
    p2eRows = [];
    document.getElementById('p2e-input').value = '';
    document.getElementById('p2e-preview').classList.add('hidden');
    document.getElementById('p2e-table').innerHTML = '';
    hideProgress('p2e-progress-fill');
}

// ---------------------------------------------------
// 10) EXCEL TO PDF
// ---------------------------------------------------
let e2pData = [];

setupDrop('e2p-drop-zone', 'e2p-input', files => handleE2p(files[0]));

async function handleE2p(file) {
    if (!file) { alert('Please pick an Excel file.'); return; }

    showFileInfo('e2p-file-info', file, resetE2p);
    progress('e2p-progress-fill', 'e2p-status', 10, 'Reading Excel file...');
    document.getElementById('e2p-preview').classList.add('hidden');

    try {
        const buf = await file.arrayBuffer();
        progress('e2p-progress-fill', 'e2p-status', 50, 'Parsing spreadsheet...');

        const wb = XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        if (!data.length) {
            progress('e2p-progress-fill', 'e2p-status', 100, 'No data found in Excel file.');
            return;
        }

        // Normalize: ensure all rows are arrays with same column count
        const maxCols = Math.max(...data.map(r => (Array.isArray(r) ? r : [r]).length));
        const normalized = data.map(r => {
            const arr = Array.isArray(r) ? r : [r];
            while (arr.length < maxCols) arr.push('');
            return arr.map(c => (c != null ? String(c) : ''));
        });

        // Preview table
        const table = document.getElementById('e2p-table');
        table.innerHTML = '';
        normalized.slice(0, 50).forEach((row, idx) => {
            const tr = document.createElement('tr');
            row.forEach(c => {
                const el = document.createElement(idx === 0 ? 'th' : 'td');
                el.textContent = c;
                tr.appendChild(el);
            });
            table.appendChild(tr);
        });
        if (normalized.length > 50) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = maxCols;
            td.style.textAlign = 'center';
            td.style.color = '#999';
            td.textContent = `+ ${normalized.length - 50} more rows`;
            tr.appendChild(td);
            table.appendChild(tr);
        }

        e2pData = normalized;
        progress('e2p-progress-fill', 'e2p-status', 100,
            `Done - ${data.length} rows, ${wb.SheetNames.length} sheet(s).`);
        document.getElementById('e2p-preview').classList.remove('hidden');

    } catch (err) {
        console.error(err);
        progress('e2p-progress-fill', 'e2p-status', 100, 'Something went wrong: ' + err.message);
    }
}

document.getElementById('e2p-download-btn').onclick = () => {
    if (!e2pData.length) return;
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('l', 'mm', 'a4');

    const head = e2pData.length > 0 ? [e2pData[0]] : [];
    const body = e2pData.slice(1);

    pdf.autoTable({
        head: head,
        body: body,
        startY: 10,
        styles: { fontSize: 8, cellPadding: 2 },
        headStyles: { fillColor: [41, 128, 185] },
        theme: 'grid'
    });

    pdf.save('converted.pdf');
};

function resetE2p() {
    e2pData = [];
    document.getElementById('e2p-input').value = '';
    document.getElementById('e2p-preview').classList.add('hidden');
    document.getElementById('e2p-table').innerHTML = '';
    hideProgress('e2p-progress-fill');
}

