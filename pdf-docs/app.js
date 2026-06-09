const fileInput = document.getElementById('pdfUpload');
const dropZone = document.getElementById('dropZone');
const convertBtn = document.getElementById('convertBtn');
const downloadLink = document.getElementById('downloadLink');
const messageEl = document.getElementById('message');
const originalSizeEl = document.getElementById('originalSize');
const pageCountEl = document.getElementById('pageCount');
const supportText = document.getElementById('supportText');

let currentFile = null;
let currentPdfBuffer = null;
const MAX_FILE_BYTES = 50 * 1024 * 1024; // 50 MB

function setMessage(text, isError = false) {
    messageEl.textContent = text;
    messageEl.style.color = isError ? '#c0392b' : '#555';
}

function formatBytes(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function updateFileDisplay(file) {
    originalSizeEl.textContent = formatBytes(file.size);
    pageCountEl.textContent = '…';
    convertBtn.disabled = true;
    downloadLink.style.display = 'none';
}

function resetState() {
    convertBtn.disabled = true;
    downloadLink.style.display = 'none';
    currentPdfBuffer = null;
    originalSizeEl.textContent = '—';
    pageCountEl.textContent = '—';
    setMessage('Choose a PDF and click Convert to create a Word document.');
}

async function handleFile(file) {
    if (!file) return;
    const fileName = file.name || '';
    const isPdfMime = file.type && file.type.toLowerCase() === 'application/pdf';
    const isPdfExt = fileName.toLowerCase().endsWith('.pdf');

    if (!isPdfMime && !isPdfExt) {
        resetState();
        setMessage('Please select a valid PDF file.', true);
        return;
    }

    if (file.size > MAX_FILE_BYTES) {
        resetState();
        setMessage('PDF is too large. Please upload a file smaller than 50 MB.', true);
        return;
    }

    currentFile = file;
    updateFileDisplay(file);
    setMessage('Reading PDF metadata…');

    try {
        const arrayBuffer = await file.arrayBuffer();
        currentPdfBuffer = arrayBuffer;
        const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
        const pdf = await loadingTask.promise;
        pageCountEl.textContent = pdf.numPages;
        convertBtn.disabled = false;
        setMessage('PDF loaded. Ready to convert.');
        showToast('PDF loaded');
    } catch (error) {
        console.error('PDF load error:', error);
        resetState();
        setMessage('Unable to read the PDF. Please upload a valid PDF file.', true);
    }
}

fileInput.addEventListener('change', (event) => {
    handleFile(event.target.files && event.target.files[0]);
});

['dragenter', 'dragover'].forEach(evt => {
    dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
});

['dragleave', 'drop'].forEach(evt => {
    dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
    });
});

dropZone.addEventListener('drop', (event) => {
    const file = event.dataTransfer.files && event.dataTransfer.files[0];
    handleFile(file);
});

convertBtn.addEventListener('click', async () => {
    if (!currentFile) {
        setMessage('No PDF selected. Please upload a PDF file first.', true);
        return;
    }

    if (!currentPdfBuffer) {
        setMessage('PDF is not fully loaded yet. Please wait.', true);
        return;
    }

    setMessage('Converting PDF to Word. This may take a moment…');
    convertBtn.disabled = true;
    downloadLink.style.display = 'none';

    try {
        const arrayBuffer = currentPdfBuffer;
        const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
        const pdf = await loadingTask.promise;
        const docChildren = [];

        for (let pageIndex = 1; pageIndex <= pdf.numPages; pageIndex += 1) {
            const page = await pdf.getPage(pageIndex);
            const textContent = await page.getTextContent({ normalizeWhitespace: true });
            const textItems = textContent.items.map(item => item.str || '').filter(Boolean);
            const pageLines = textItems.join(' ').split(/\r?\n|\u2028|\u2029/).map(line => line.trim()).filter(Boolean);

            docChildren.push(new docx.Paragraph({
                text: `Page ${pageIndex}`,
                heading: docx.HeadingLevel.HEADING_3,
            }));

            if (pageLines.length === 0) {
                docChildren.push(new docx.Paragraph({
                    text: 'No extractable text found on this page.',
                    italics: true,
                }));
            } else {
                pageLines.forEach((line) => {
                    docChildren.push(new docx.Paragraph({ text: line }));
                });
            }

            if (pageIndex < pdf.numPages) {
                docChildren.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
            }
        }

        const doc = new docx.Document({
            sections: [{ properties: {}, children: docChildren }],
        });

        const blob = await docx.Packer.toBlob(doc);
        if (!blob || blob.size === 0) {
            throw new Error('Generated Word file is empty.');
        }

        if (downloadLink.href) {
            URL.revokeObjectURL(downloadLink.href);
        }

        const url = URL.createObjectURL(blob);
        downloadLink.href = url;
        downloadLink.download = currentFile.name.replace(/\.pdf$/i, '') + '.docx';
        downloadLink.style.display = 'inline-flex';
        downloadLink.textContent = `Download Word (${formatBytes(blob.size)})`;
        setMessage('Conversion complete. Download your Word file below.');
    } catch (error) {
        console.error('Conversion error:', error);
        setMessage('Conversion failed. Please try another PDF or refresh the page.', true);
        supportText.textContent = 'If the issue persists, contact pndkpl007@gmail.com for support.';
    } finally {
        convertBtn.disabled = false;
    }
});

const toastEl = document.getElementById('toast');
let toastTimer = null;

function showToast(text) {
    if (!toastEl) return;
    toastEl.textContent = text;
    toastEl.classList.add('visible');
    clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => {
        toastEl.classList.remove('visible');
    }, 2100);
}

function initPage() {
    if (typeof pdfjsLib === 'undefined' || typeof pdfjsLib.getDocument !== 'function') {
        setMessage('PDF library failed to load. Please refresh the page.', true);
        convertBtn.disabled = true;
        return;
    }
    if (typeof docx === 'undefined' || typeof docx.Document !== 'function' || typeof docx.Packer !== 'object') {
        setMessage('Word export library failed to load. Please refresh the page.', true);
        convertBtn.disabled = true;
        return;
    }

    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
    resetState();
}

window.addEventListener('DOMContentLoaded', initPage);
