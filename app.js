function formatBytes(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

const MIN_TARGET = 100 * 1024; // 100 KB

const fileInput = document.getElementById('imageUpload');
const originalSizeEl = document.getElementById('originalSize');
const targetSizeEl = document.getElementById('targetSize');
const sizeSlider = document.getElementById('sizeSlider');
const compressBtn = document.getElementById('compressBtn');
const downloadLink = document.getElementById('downloadLink');
const messageEl = document.getElementById('message');
const dropZone = document.getElementById('dropZone');

let currentFile = null;

function setMessage(text, isError = false) {
    messageEl.textContent = text;
    messageEl.style.color = isError ? '#c0392b' : '#555';
}

function updateSliderRange(size) {
    const max = Math.max(size, MIN_TARGET);
    sizeSlider.min = MIN_TARGET;
    sizeSlider.max = max;
    sizeSlider.value = Math.min(max, Math.floor(size * 0.5));
    sizeSlider.disabled = false;
    compressBtn.disabled = false;
    targetSizeEl.textContent = formatBytes(Number(sizeSlider.value));
}

fileInput.addEventListener('change', (e) => {
    if (!e.target.files || !e.target.files[0]) return;
    handleFile(e.target.files[0]);
});

['dragenter', 'dragover'].forEach(evt => {
    dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropZone.style.backgroundColor = '#f4f6ff';
    });
});
['dragleave', 'drop'].forEach(evt => {
    dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropZone.style.backgroundColor = '';
    });
});

dropZone.addEventListener('drop', (e) => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
});

sizeSlider.addEventListener('input', () => {
    targetSizeEl.textContent = formatBytes(Number(sizeSlider.value));
});

compressBtn.addEventListener('click', async () => {
    if (!currentFile) return;
    const target = Math.max(MIN_TARGET, Number(sizeSlider.value));
    setMessage('Compressing…');
    try {
        const blob = await compressImageToTarget(currentFile, target);
        if (!blob) {
            setMessage('Could not compress to target size. Try a larger target.', true);
            return;
        }
        const url = URL.createObjectURL(blob);
        downloadLink.href = url;
        downloadLink.download = (currentFile.name || 'compressed')
            .replace(/\.[^.]+$/, '') + '.jpg';
        downloadLink.style.display = 'inline-flex';
        downloadLink.textContent = `Download (${formatBytes(blob.size)})`;
        setMessage('Compression complete.');
    } catch (err) {
        setMessage('Compression failed: ' + err.message, true);
    }
});

function handleFile(file) {
    currentFile = file;
    originalSizeEl.textContent = formatBytes(file.size);
    downloadLink.style.display = 'none';
    if (!file.type.startsWith('image/')) {
        setMessage('Please select an image file.', true);
        sizeSlider.disabled = true;
        compressBtn.disabled = true;
        return;
    }
    if (file.size <= MIN_TARGET) {
        setMessage('Image is already smaller than 100 KB; no compression needed.');
        sizeSlider.disabled = true;
        compressBtn.disabled = true;
        targetSizeEl.textContent = formatBytes(file.size);
        return;
    }
    setMessage('Ready. Choose a target size (min 100 KB).');
    updateSliderRange(file.size);
}

async function compressImageToTarget(file, targetBytes) {
    // Load image into canvas
    const img = await loadImageFromFile(file);
    const canvas = document.createElement('canvas');
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0);

    // We'll convert to JPEG and binary-search quality to reach target size
    let minQ = 0.05, maxQ = 0.95;
    let bestBlob = null;

    for (let i = 0; i < 8; i++) { // limit iterations
        const q = (minQ + maxQ) / 2;
        const dataUrl = canvas.toDataURL('image/jpeg', q);
        const blob = dataURLToBlob(dataUrl);
        if (blob.size <= targetBytes) {
            bestBlob = blob;
            // try lower quality to reduce further (but don't go below 0.05)
            maxQ = q;
        } else {
            // too large, reduce quality
            minQ = q;
        }
        // small break if already very close
        if (bestBlob && Math.abs(bestBlob.size - targetBytes) < 1024) break;
    }

    // If bestBlob still larger than target, return null
    if (!bestBlob) return null;
    // Ensure at least MIN_TARGET
    if (bestBlob.size < MIN_TARGET) return bestBlob;
    return bestBlob;
}

function loadImageFromFile(file) {
    return new Promise((resolve, reject) => {
        const url = URL.createObjectURL(file);
        const img = new Image();
        img.onload = () => {
            URL.revokeObjectURL(url);
            resolve(img);
        };
        img.onerror = (e) => reject(new Error('Failed to load image'));
        img.src = url;
    });
}

function dataURLToBlob(dataurl) {
    const arr = dataurl.split(',');
    const mime = arr[0].match(/:(.*?);/)[1];
    const bstr = atob(arr[1]);
    let n = bstr.length;
    const u8arr = new Uint8Array(n);
    while (n--) {
        u8arr[n] = bstr.charCodeAt(n);
    }
    return new Blob([u8arr], { type: mime });
}
