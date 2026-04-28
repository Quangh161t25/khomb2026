const CONFIG = {
    spreadsheetId: "1cnA33cHHMhcOSaXa9l4Jeu6qw8QnXlUnEU4Bqtkj9wo",
    authSheetName: "DSNV",
    udctSheetName: "UD_CT",
    sanphamSheetName: "DS_SP_PM",
    dsSpCtSheetName: "DS_SP_CT",
    inventorySheetName: "TON_KHO",
    dhctSheetName: "DH_CT",
    hhbhSheetName: "HH_BH",
    hhNvDienSheetName: "HH_NV_DIEN",
    imgbbApiKey: "1bad1429a242d7040fda3f2cfddb3a25",
    serviceAccountEmail: "test-gia-ason@api-test-sheet-161.iam.gserviceaccount.com",
    privateKey: "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC3NN84hLTkQPZd\nLj7niXZTICq7nHsuTn3J6r2Paq12m70/lYSmrwh1i0EStr9bO19QM8cevGlslwGr\nWSVOLJlc6+w1HGPKvRXtA41kYV9MYIvpzIPQtkFE7Hxq71QyBARcv39Lfzze6Ioj\n3G8VBvAKFLAnCUr97GHRv+KbCTFxPZupd3PEB+xS5ZUlzdBCEZvDid3iXaaEJJ+l\nTd1apAGQHjtnDTLOkiTa8zf7X5ebALwnI9MziOdN8VyprHXGhkachPbKyrG0QwEs\n2jtiI6Y5ULsBPjNefoavH8MKU5DEAT9h0fZ7KfsKYVMDuXqmEKBs0D3B4Z6aDZQW\nwT2dDRZDAgMBAAECggEAEIuVoSzZVuFhaz1GI9ji0IacjvO50cIq7M8Zrj4/F756\nEw6PIhKENafAb7U4INm2AnzUMO8CqL9Jpxs85qUM3W4JysSByqLUiRW2184amIyb\nj7jCXfLBTQn8AbHgrUepl5d/vBmFYMgon/mqjbNiGDb4FZgEQSkie5o6fi/dWp5d\nNahbZl+WTOB/znhAfKh/zferHNxldR/ERmwOubZUerkqysWiBigc3ovpLSUof9ur\nz3hNPPp0CKQjF40xuQc6FYTHUHMLuMvp78PXuc/mYqQmZ8VOGhU+faGtZ4m+QJly\ndF5dS8U5cwKEF+ptuAUiWSahn6INb9yKn3+FcsW0UQKBgQDb8N4eWFvbgpRo/vxo\nwBN2u2TWubj6clcrq/1a+VR0njC28Can0ogJHhrFhPxVs5D/rugs3HlbyAXJFptY\nV0DZPCwBxGU5P5RbGjXWWEUXjp4ISKQD8WKfVlXNr79TqLdOg2NZBYQAi06Cpo/T\nPV9l7LSG2Tj/9WdvD7W2wvrpaQKBgQDVPjpJN6xh7+sHtSU0mjKvrqigpHbuSQ/o\nXpUaWSIpJffm5QpFPAOcTT5mHZCyllicJQIrfPSY+sH8n+sF03CUqVkV4Q2UqfOf\npFaLDB4P6SQ8iesZyF4VKFrj/cAvRJmp0e5W/DRnFkoEp+8c+nrru2+Dzm9kb7Uq\n0CiltqYAywKBgBtcfrV1to+7Ue0x84KwintV2rifyDRX7yI+tjkQFYKgf1zyyUxN\nc6D2vsvdvGqI+TvlrXqPPwW8/4NBrbeyux2LT8o0fYc+sp0WyKXOu2Gv21caelUH\nPYam/eultn6Y2Z0J2V0kw4Qx0GWOhQv5cZnDdb3k3iNxixmU8b03ynEpAoGBAKEA\n7O0fNe50QRZ+tOq0ihSPYQ55XrqnO3WNBDLynZJH8pbI1CpWF7vJrpVXOUs9rQWo\nA61mGR/wJMtiywaJEHWOL48PbzuR3jno0NcHfSMyOoPi9jlvSWncIFQH4TVPLF5F\n/Rh8L+ytrZE6YpWUoX6e9KGmGgDRPw5mQGpuL4RlAoGADe9n080SXlsUk4nHVjUz\nEfv7EBoBkgOpqb9T1foRfJl46NxmmTOYV3iGIhjwcDskEg284k4iq/gH6EEFyEBc\nVz13jzB1nBgjfezFesVQz7bA/+Wik6HZtxAxVg38BKMt+Q1tYw9wOjbGPqOn++VC\nsR2Sh8e3h3Knd6j1tceRIFU=\n-----END PRIVATE KEY-----\n",
    tokenUrl: "https://oauth2.googleapis.com/token",
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
};

let isUploading = false;

// Chạy ngay khi DOM vừa tải để tránh chớp màn hình đăng nhập
document.addEventListener('DOMContentLoaded', () => {
    if (localStorage.getItem('erp_current_user')) {
        const ls = document.getElementById('loginScreen');
        const ma = document.getElementById('mainApp');
        if (ls) ls.classList.add('hidden');
        if (ma) ma.classList.remove('hidden');
    }
});

let isLoggedIn = false;
let usersData = [];
let currentUser = null;
let accessToken = null;
let tokenExpiry = 0;
let udctData = [];
let sanphamData = [];
let dsSpCtData = [];
let reportData = [];
let mergedChart = null;
let upmisaData = [];
let inventoryData = [];
let dhctData = [];
let hangHoanData = [];
let filteredHangHoanData = [];
let hhShopDienData = [];
let filteredHHShopDienData = [];
let hhBhMvdSet = new Set();
let currentEditRowIndex = -1;
let currentHangHoanEditIndex = -1;
let hhDrawerMode = 'edit';
let currentHHShopRowIndex = -1;
let continuousScanRunning = false;
let continuousScanStream = null;
let continuousScanLastValue = '';
let continuousScanLastAt = 0;
let currentDrawerMode = 'udct';
let udctAutoSaveTimer = null;
let suppressUDCTAutoSave = false;

function escapeHtml(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function toYMD(input) {
    if (!input) return '';
    const v = String(input).trim().split(' ')[0];
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(v)) {
        const [d, m, y] = v.split('/');
        return `${y}-${m}-${d}`;
    }
    return '';
}

function fillHangHoanFilterOptions() {
    const fillSelect = (id, values, label) => {
        const el = document.getElementById(id);
        if (!el) return;
        el.innerHTML = `<option value="">${label}</option>` + values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
    };
    const renderKhoButtons = (values, currentVal) => {
        const container = document.getElementById('filterHHKhoButtons');
        if (!container) return;
        let html = `<button onclick="setHHKhoFilter('')" class="px-3 py-1.5 text-[11px] rounded-lg font-bold transition-all duration-200 ${currentVal === '' ? 'bg-white text-primary shadow-sm' : 'text-slate-500 hover:text-slate-700'}">Tất cả</button>`;
        values.forEach(v => {
            html += `<button onclick="setHHKhoFilter('${escapeHtml(v)}')" class="px-3 py-1.5 text-[11px] rounded-lg font-bold transition-all duration-200 ${currentVal === v ? 'bg-white text-primary shadow-sm' : 'text-slate-500 hover:text-slate-700'}">${escapeHtml(v)}</button>`;
        });
        container.innerHTML = html;
    };
    const khoList = [...new Set(hangHoanData.map(i => i.kho).filter(Boolean))].sort();
    const gianList = [...new Set(hangHoanData.map(i => i.ma_gian).filter(Boolean))].sort();
    const currentKho = document.getElementById('filterHHKho')?.value || '';
    renderKhoButtons(khoList, currentKho);
    fillSelect('filterHHMaGian', gianList, 'Tất cả gian hàng');
}

function setHHKhoFilter(value) {
    const el = document.getElementById('filterHHKho');
    if (el) el.value = value;
    fillHangHoanFilterOptions();
    filterHangHoanData();
}

function setHangHoanToday() {
    const fromInput = document.getElementById('filterHHFrom');
    const toInput = document.getElementById('filterHHTo');
    if (!fromInput || !toInput) return;
    const today = new Date();
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    const todayStr = `${y}-${m}-${d}`;
    fromInput.value = todayStr;
    toInput.value = todayStr;
    filterHangHoanData();
}

function changeHangHoanDate(which, step) {
    const input = document.getElementById(which === 'from' ? 'filterHHFrom' : 'filterHHTo');
    if (!input) return;
    if (!input.value) {
        setHangHoanToday();
        return;
    }
    const parts = input.value.split('-');
    if (parts.length !== 3) return;
    const dt = new Date(parts[0], parts[1] - 1, parts[2]);
    dt.setDate(dt.getDate() + step);
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    input.value = `${y}-${m}-${d}`;
    filterHangHoanData();
}

function openImagePreview(url) {
    if (!url) return;
    const overlay = document.getElementById('imagePreviewOverlay');
    const image = document.getElementById('imagePreviewContent');
    if (!overlay || !image) return;
    image.src = url;
    overlay.classList.remove('hidden');
}

function closeImagePreview() {
    const overlay = document.getElementById('imagePreviewOverlay');
    const image = document.getElementById('imagePreviewContent');
    if (!overlay || !image) return;
    overlay.classList.add('hidden');
    image.src = '';
}

async function loadDsSpCtData() {
    try {
        const data = await fetchSheetData(CONFIG.dsSpCtSheetName);
        if (!data || data.length <= 1) {
            dsSpCtData = [];
            return;
        }
        const headers = data[0].map(h => (h || '').toString().toLowerCase().trim());
        const findIdx = (names) => {
            for (const name of names) {
                const idx = headers.indexOf(name.toLowerCase());
                if (idx !== -1) return idx;
            }
            return -1;
        };
        const idxIdSp = findIdx(['id_sp', 'sku', 'id sp']);
        const idxIdSpCt = findIdx(['id_sp_ct', 'sku_ct', 'id sp ct']);
        const idxTen = findIdx(['ten', 'tên', 'ten_sp', 'tên sản phẩm']);
        const idxGiaNhap = findIdx(['gia_nhap', 'giá nhập', 'gianhap', 'gia_mua', 'gia']);

        dsSpCtData = data.slice(1).map(row => ({
            id_sp: (idxIdSp !== -1 ? row[idxIdSp] : row[3] || '').toString().trim(),
            id_sp_ct: (idxIdSpCt !== -1 ? row[idxIdSpCt] : row[2] || '').toString().trim(),
            ten: (idxTen !== -1 ? row[idxTen] : row[4] || '').toString().trim(),
            gia_nhap: (idxGiaNhap !== -1 ? row[idxGiaNhap] : 0)
        })).filter(item => item.id_sp || item.id_sp_ct || item.ten);
        populateHhFormOptions();
    } catch (error) {
        console.error('Load DS_SP_CT error:', error);
        dsSpCtData = [];
    }
}

function renderHhKhoButtons(value) {
    const current = (value || '').toString().trim().toUpperCase() || 'KHO';
    const buttons = document.querySelectorAll('#hhEditKhoButtons button');
    buttons.forEach(btn => {
        const active = btn.textContent.trim().toUpperCase() === current;
        btn.classList.toggle('bg-slate-100', active);
        btn.classList.toggle('bg-white', !active);
    });
}

function setHhKho(value) {
    const normalized = (value || 'KHO').toString().trim().toUpperCase();
    document.getElementById('hhEditKho').value = normalized;
    renderHhKhoButtons(normalized);
}

function populateHhFormOptions() {
    const maGianList = document.getElementById('hhMaGianList');
    const skuList = document.getElementById('hhSkuList');
    if (maGianList) {
        const uniqueMaGian = [...new Set(hangHoanData.map(i => (i.ma_gian || '').toString().trim()).filter(Boolean))].sort();
        maGianList.innerHTML = uniqueMaGian.map(v => `<option value="${escapeHtml(v)}">`).join('');
    }
    if (skuList) {
        const uniqueSku = [...new Set(dsSpCtData.map(i => (i.id_sp || '').toString().trim()).filter(Boolean))].sort();
        skuList.innerHTML = uniqueSku.map(v => `<option value="${escapeHtml(v)}">`).join('');
    }
    handleHhSkuChange();

    // HH SHOP ĐIỀN suggestions
    const hhShopMaGianList = document.getElementById('hhShopMaGianList');
    const hhShopSkuList = document.getElementById('hhShopSkuList');
    if (hhShopMaGianList) {
        const uniqueMaGian = [...new Set([
            ...hangHoanData.map(i => (i.ma_gian || '').toString().trim()),
            ...hhShopDienData.map(i => (i.ma_gian || '').toString().trim()),
            ...udctData.map(i => (i.ma_gian || '').toString().trim())
        ].filter(Boolean))].sort();
        hhShopMaGianList.innerHTML = uniqueMaGian.map(v => `<option value="${escapeHtml(v)}">`).join('');
    }
    if (hhShopSkuList) {
        const uniqueSku = [...new Set([
            ...dsSpCtData.map(i => (i.id_sp || '').toString().trim()),
            ...hhShopDienData.map(i => (i.sku || '').toString().trim())
        ].filter(Boolean))].sort();
        hhShopSkuList.innerHTML = uniqueSku.map(v => `<option value="${escapeHtml(v)}">`).join('');
    }
}

function handleHhSkuChange() {
    const sku = (document.getElementById('hhEditSKU')?.value || '').toString().trim();
    const skuCtList = document.getElementById('hhSkuCtList');
    const skuCtBtns = document.getElementById('hhSkuCtButtons');

    if (skuCtList) {
        const options = [...new Set(dsSpCtData
            .filter(item => ((item.id_sp_ct || '').substring(0, 4).toUpperCase() === sku.toUpperCase()))
            .map(item => item.id_sp_ct)
            .filter(Boolean))];

        skuCtList.innerHTML = options.map(v => `<option value="${escapeHtml(v)}">`).join('');

        // Hiển thị dạng nút chọn nhanh
        if (skuCtBtns) {
            if (sku && options.length > 0) {
                skuCtBtns.innerHTML = options.map(v => `
                    <button type="button" onclick="setHhSkuCt('${escapeHtml(v)}')" 
                        class="px-3 py-1.5 rounded-lg border border-blue-200 bg-blue-50 text-blue-700 text-[11px] font-bold hover:bg-blue-100 transition-all">
                        ${escapeHtml(v)}
                    </button>
                `).join('');
            } else {
                skuCtBtns.innerHTML = '';
            }
        }
    }
    const currentSkuCt = (document.getElementById('hhEditSKUCT')?.value || '').toString().trim();
    if (currentSkuCt && currentSkuCt.substring(0, 4).toUpperCase() !== sku.toUpperCase()) {
        document.getElementById('hhEditSKUCT').value = '';
        if (skuCtBtns) skuCtBtns.innerHTML = '';
    }
}

function setHhSkuCt(value) {
    const input = document.getElementById('hhEditSKUCT');
    if (input) {
        input.value = value;
        handleHhSkuCtChange();
    }
    // Xóa các nút sau khi chọn
    const skuCtBtns = document.getElementById('hhSkuCtButtons');
    if (skuCtBtns) skuCtBtns.innerHTML = '';
}

function handleHhSkuCtChange() {
    const skuCt = (document.getElementById('hhEditSKUCT')?.value || '').toString().trim();
    if (!skuCt) return;
    const match = dsSpCtData.find(item => (item.id_sp_ct || '').toString() === skuCt);
    if (!match) return;
    if (!document.getElementById('hhEditSKU').value) {
        document.getElementById('hhEditSKU').value = match.id_sp || skuCt.substring(0, 4);
    }
    document.getElementById('hhEditTenSP').value = match.ten || document.getElementById('hhEditTenSP').value;
    handleHhSkuChange();
}

async function scanQrToInput(inputId) {
    const inputEl = document.getElementById(inputId);
    if (!inputEl) return;

    const overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:9999;display:flex;flex-direction:column;align-items:center;justify-content:center;';

    const video = document.createElement('video');
    video.style.cssText = 'max-width:90vw;max-height:70vh;border-radius:12px;border:3px solid #fff;box-shadow:0 0 20px rgba(0,0,0,0.5);';

    const closeBtn = document.createElement('button');
    closeBtn.innerHTML = 'Hủy / Tắt Camera';
    closeBtn.style.cssText = 'margin-top:20px;padding:12px 24px;background:#ef4444;color:#fff;border:none;border-radius:30px;font-weight:bold;font-size:14px;cursor:pointer;box-shadow:0 10px 15px -3px rgba(0,0,0,0.1); transition: all 0.2s;';
    closeBtn.onmouseover = () => closeBtn.style.transform = 'scale(1.05)';
    closeBtn.onmouseout = () => closeBtn.style.transform = 'scale(1)';

    overlay.appendChild(video);
    overlay.appendChild(closeBtn);
    document.body.appendChild(overlay);

    let stream;
    let isStopped = false;

    closeBtn.onclick = () => {
        isStopped = true;
        if (stream) stream.getTracks().forEach(t => t.stop());
        overlay.remove();
    };

    if (!('BarcodeDetector' in window) || !navigator.mediaDevices?.getUserMedia) {
        const manual = prompt('Thiết bị không hỗ trợ quét QR tự động. Dán mã tại đây:');
        if (manual) {
            inputEl.value = manual.trim();
            inputEl.dispatchEvent(new Event('input', { bubbles: true }));
            inputEl.dispatchEvent(new Event('change', { bubbles: true }));
        }
        overlay.remove();
        return;
    }

    try {
        const detector = new BarcodeDetector({ formats: ['qr_code', 'code_128', 'ean_13'] });
        stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
        video.srcObject = stream;
        video.setAttribute('playsinline', 'true');
        await video.play();

        while (!isStopped) {
            const barcodes = await detector.detect(video);
            if (barcodes.length) {
                inputEl.value = (barcodes[0].rawValue || '').trim();
                inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                inputEl.dispatchEvent(new Event('change', { bubbles: true }));
                break;
            }
            await new Promise(r => setTimeout(r, 150));
        }
    } catch (err) {
        console.error('QR scan error:', err);
        const manual = prompt('Lỗi camera / Không quét được. Dán mã tại đây:');
        if (manual) {
            inputEl.value = manual.trim();
            inputEl.dispatchEvent(new Event('input', { bubbles: true }));
            inputEl.dispatchEvent(new Event('change', { bubbles: true }));
        }
    } finally {
        if (stream) stream.getTracks().forEach(t => t.stop());
        if (overlay.parentNode) overlay.remove();
    }
}

function handleHhMvdInputChange(val) {
    const noticeEl = document.getElementById('hhMvdDuplicateNotice');
    if (!noticeEl) return;
    const mvd = (val || '').toString().trim();
    if (!mvd) {
        noticeEl.classList.add('hidden');
        return;
    }
    // Tìm trong dữ liệu hàng hoàn hiện có (loại trừ chính nó nếu đang edit)
    const duplicate = hangHoanData.find(item => {
        const isDuplicateMvd = (item.mvd || '').toString().trim() === mvd || (item.mvd_2 || '').toString().trim() === mvd;
        if (!isDuplicateMvd) return false;

        if (hhDrawerMode === 'edit' && currentHangHoanEditIndex !== -1) {
            const currentItem = hangHoanData[currentHangHoanEditIndex];
            return item !== currentItem;
        }
        return true;
    });

    if (duplicate) {
        noticeEl.textContent = `⚠️ MVD điền vào ngày ${duplicate.ngay_nhan || '?'}`;
        noticeEl.classList.remove('hidden');
    } else {
        noticeEl.classList.add('hidden');
    }
}

async function scanQrForHhMvd() {
    await scanQrToInput('hhEditMVD');
}

async function scanQrForHhMvd2() {
    await scanQrToInput('hhEditMVD2');
}

async function uploadImageHh(input, index) {
    const file = input.files[0];
    if (!file) return;

    // Chặn tự động lưu trong khi đang upload
    isUploading = true;

    const statusLabel = document.getElementById('saveStatus');
    if (statusLabel) {
        statusLabel.textContent = "Đang tải ảnh lên ImgBB...";
        statusLabel.style.display = 'block';
    }

    try {
        const formData = new FormData();
        formData.append('image', file);

        // 1. Gửi ảnh lên ImgBB
        const response = await fetch(`https://api.imgbb.com/1/upload?key=${CONFIG.imgbbApiKey}`, {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            const directUrl = result.data.url;

            // 2. Cập nhật link vào ô Input tương ứng
            const targetInput = document.getElementById(`hhEditAnh${index}`);
            if (targetInput) {
                targetInput.value = directUrl;
            }

            // 3. Cập nhật preview và lưu dữ liệu
            refreshHhImagePreviews();
            isUploading = false;
            await saveHhDetail();
        } else {
            alert("Lỗi tải ảnh lên ImgBB: " + (result.error?.message || "Không xác định"));
            isUploading = false;
            refreshHhImagePreviews();
        }
    } catch (err) {
        console.error("Upload error:", err);
        alert("Lỗi kết nối khi tải ảnh: " + err.message);
        isUploading = false;
        refreshHhImagePreviews();
    } finally {
        if (statusLabel) {
            statusLabel.style.display = 'none';
        }
        input.value = ""; // Reset input file
    }
}

function refreshHhImagePreviews() {
    // Luôn quét qua 3 ô tiềm năng để đảm bảo đồng bộ
    for (let i = 1; i <= 3; i++) {
        const urlEl = document.getElementById(`hhEditAnh${i}`);
        const preview = document.getElementById(`hhImagePreview${i}`);
        if (!urlEl || !preview) continue;
        const url = urlEl.value;
        if (url) {
            preview.innerHTML = `
                <img src="${url}" class="w-full h-full object-cover">
                <button onclick="event.stopPropagation(); removeHhImage(${i})" class="absolute top-0 right-0 p-1.5 bg-rose-500 text-white rounded-bl-lg hover:bg-rose-600 transition-all shadow-md">
                    <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12" />
                    </svg>
                </button>
            `;
            preview.classList.remove('border-dashed');
            preview.classList.add('border-solid', 'border-primary/20');
        } else {
            preview.innerHTML = `
                <div class="text-center p-2">
                    <svg class="w-8 h-8 mx-auto text-slate-300 group-hover:text-primary transition-colors"
                        fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                            d="M3 9a2 2 0 012-2h.93a2 2 0 001.664-.89l.812-1.22A2 2 0 0110.07 4h3.86a2 2 0 011.664.89l.812 1.22A2 2 0 0018.07 7H19a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2V9z" />
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                            d="M15 13a3 3 0 11-6 0 3 3 0 016 0z" />
                    </svg>
                    <span class="text-[10px] text-slate-400 block mt-1 font-bold">Chạm để tải ảnh</span>
                </div>
            `;
            preview.classList.add('border-dashed');
            preview.classList.remove('border-solid', 'border-primary/20');
        }
    }
}

function removeHhImage(index) {
    const el = document.getElementById(`hhEditAnh${index}`);
    if (el) el.value = '';
    refreshHhImagePreviews();
}

async function appendHangHoanQuickByMvd(mvdRaw) {
    const mvd = (mvdRaw || '').toString().trim();
    if (!mvd) return false;
    const token = await getAccessToken();
    if (!token) return false;
    const today = new Date().toISOString().split('T')[0];
    const appendValues = [[
        `${Date.now()}`,
        today,
        mvd,
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '1',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        'KHO',
        '',
        '',
        '',
        '',
        ''
    ]];
    const appendUrl = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.hhbhSheetName}!A:Z:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
    const appendResp = await fetch(appendUrl, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: appendValues })
    });
    return appendResp.ok;
}

function stopContinuousMvdScan() {
    continuousScanRunning = false;
    if (continuousScanStream) {
        continuousScanStream.getTracks().forEach(t => t.stop());
        continuousScanStream = null;
    }
    const root = document.getElementById('continuousScanOverlay');
    if (root) root.remove();
}

async function startContinuousMvdScan() {
    if (continuousScanRunning) return;
    if (!('BarcodeDetector' in window) || !navigator.mediaDevices?.getUserMedia) {
        alert('Thiết bị chưa hỗ trợ quét QR liên tục tự động.');
        return;
    }
    const overlay = document.createElement('div');
    overlay.id = 'continuousScanOverlay';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:10000;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;padding:16px;';
    overlay.innerHTML = `
                <div style="color:#fff;font-weight:700;font-size:18px;">Quét liên tục MVD</div>
                <div id="continuousScanStatus" style="color:#cbd5e1;font-size:13px;">Đang khởi động camera...</div>
                <video id="continuousScanVideo" style="max-width:94vw;max-height:72vh;border-radius:10px;border:2px solid rgba(255,255,255,.55);"></video>
                <div style="display:flex;gap:8px;">
                    <button id="continuousStopBtn" style="padding:8px 14px;border-radius:8px;border:none;background:#ef4444;color:#fff;font-weight:700;cursor:pointer;">Dừng quét</button>
                </div>
            `;
    document.body.appendChild(overlay);
    const statusEl = document.getElementById('continuousScanStatus');
    const video = document.getElementById('continuousScanVideo');
    document.getElementById('continuousStopBtn')?.addEventListener('click', stopContinuousMvdScan);
    continuousScanRunning = true;
    continuousScanLastValue = '';
    continuousScanLastAt = 0;
    try {
        const detector = new BarcodeDetector({ formats: ['qr_code'] });
        continuousScanStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
        video.srcObject = continuousScanStream;
        video.setAttribute('playsinline', 'true');
        await video.play();
        let okCount = 0;
        while (continuousScanRunning) {
            const codes = await detector.detect(video);
            if (codes.length) {
                const rawValue = (codes[0].rawValue || '').trim();
                const now = Date.now();
                // Chống lưu trùng liên tiếp do camera giữ khung hình.
                if (rawValue && (rawValue !== continuousScanLastValue || (now - continuousScanLastAt) > 2000)) {
                    continuousScanLastValue = rawValue;
                    continuousScanLastAt = now;
                    statusEl.textContent = `Đang lưu MVD: ${rawValue} ...`;
                    const ok = await appendHangHoanQuickByMvd(rawValue);
                    if (ok) {
                        okCount += 1;
                        statusEl.textContent = `Đã lưu ${okCount} mã. MVD mới nhất: ${rawValue}`;
                    } else {
                        statusEl.textContent = `Lưu lỗi với mã: ${rawValue}. Tiếp tục quét...`;
                    }
                }
            }
            await new Promise(r => setTimeout(r, 160));
        }
    } catch (error) {
        console.error('Continuous scan error:', error);
        alert('Không thể khởi tạo quét liên tục MVD.');
    } finally {
        stopContinuousMvdScan();
        // Làm mới bảng để nhìn thấy dữ liệu vừa quét.
        await fetchHangHoanData();
    }
}

async function startHHSearchQrScan() {
    if (!('BarcodeDetector' in window) || !navigator.mediaDevices?.getUserMedia) {
        alert('Thiết bị chưa hỗ trợ quét QR/Barcode.');
        return;
    }
    const overlay = document.createElement('div');
    overlay.id = 'qrSearchOverlay';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.9);z-index:99999;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:20px;';
    overlay.innerHTML = `
        <div style="color:#fff;font-weight:bold;font-size:18px;">QUÉT MVD ĐỂ TÌM KIẾM</div>
        <video id="qrSearchVideo" style="width:90vw;max-height:65vh;border:3px solid #fff;border-radius:16px;object-fit:cover;"></video>
        <button id="qrSearchCloseBtn" style="padding:12px 30px;background:#ef4444;color:#fff;border-radius:12px;border:none;font-weight:bold;font-size:16px;box-shadow:0 10px 15px -3px rgba(0,0,0,0.1);">ĐÓNG</button>
    `;
    document.body.appendChild(overlay);
    const video = document.getElementById('qrSearchVideo');
    let running = true;
    let stream = null;

    document.getElementById('qrSearchCloseBtn').onclick = () => {
        running = false;
        if (stream) stream.getTracks().forEach(t => t.stop());
        overlay.remove();
    };

    try {
        const detector = new BarcodeDetector({ formats: ['qr_code', 'code_128', 'ean_13', 'code_39'] });
        stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
        video.srcObject = stream;
        video.setAttribute('playsinline', 'true');
        await video.play();

        while (running) {
            const codes = await detector.detect(video);
            if (codes.length) {
                const val = (codes[0].rawValue || '').trim();
                if (val) {
                    const searchInput = document.getElementById('filterHHSearch');
                    if (searchInput) {
                        searchInput.value = val;
                        filterHangHoanData();
                        // Thông báo ngắn gọn
                        statusNotify('Đã nhập mã: ' + val, 'success');
                    }
                    running = false;
                    break;
                }
            }
            await new Promise(r => setTimeout(r, 150));
        }
    } catch (err) {
        console.error(err);
        alert('Lỗi camera: ' + err.message);
    } finally {
        if (stream) stream.getTracks().forEach(t => t.stop());
        overlay.remove();
    }
}

async function fetchHangHoanData() {
    const loadingOverlay = document.getElementById('loadingOverlay');
    const tbody = document.getElementById('hangHoanTableBody');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');
    if (tbody) tbody.innerHTML = `<tr><td colspan="11" class="text-center py-8 text-slate-500">Đang tải dữ liệu...</td></tr>`;

    try {
        const token = await getAccessToken();
        if (!token) return;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.hhbhSheetName}!A:Z`;
        const response = await fetch(url, { headers: { "Authorization": `Bearer ${token}` } });
        const result = await response.json();

        if (result.values && result.values.length > 1) {
            hangHoanData = result.values.slice(1).map((row, idx) => ({
                rowIndex: idx + 2,
                id: row[0] || '',
                ngay_nhan: toYMD(row[1] || ''),
                mvd: row[2] || '',
                mvd_2: row[3] || '',
                ma_gian: row[4] || '',
                anh_1: row[5] || '',
                anh_2: row[6] || '',
                anh_3: row[7] || '',
                ngay_xly: row[8] || '',
                sku: row[9] || '',
                sku_ct: row[10] || '',
                slg: row[11] || '',
                ten_sp: row[12] || '',
                ghi_chu: row[13] || '',
                tinh_trang: row[14] || '',
                trang_thai: row[15] || '',
                sku_slg: row[16] || '',
                id_nv: row[17] || '',
                udt: row[18] || '',
                mvd_gian: row[19] || '',
                kho: row[20] || '',
                lb3: row[21] || '',
                id_dh: row[22] || '',
                id_dh_ct: row[23] || '',
                stt: row[24] || '',
                danh_dau: row[25] || ''
            })).sort((a, b) => {
                const da = toYMD(a.ngay_nhan);
                const db = toYMD(b.ngay_nhan);
                return db.localeCompare(da);
            });
            setHangHoanToday();
            fillHangHoanFilterOptions();
            populateHhFormOptions();
            filterHangHoanData();
        } else {
            hangHoanData = [];
            filteredHangHoanData = [];
            if (tbody) tbody.innerHTML = `<tr><td colspan="11" class="text-center py-8 text-slate-500">Không có dữ liệu hàng hoàn.</td></tr>`;
        }
    } catch (err) {
        console.error("Lỗi tải HH_BH:", err);
        if (tbody) tbody.innerHTML = `<tr><td colspan="11" class="text-center py-8 text-slate-500">Không thể tải dữ liệu từ HH_BH.</td></tr>`;
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

function filterHangHoanData() {
    const fFrom = document.getElementById('filterHHFrom')?.value || '';
    const fTo = document.getElementById('filterHHTo')?.value || '';
    const fKho = document.getElementById('filterHHKho')?.value || '';
    const fGian = document.getElementById('filterHHMaGian')?.value || '';
    const search = (document.getElementById('filterHHSearch')?.value || '').toLowerCase().trim();

    filteredHangHoanData = hangHoanData.filter(item => {
        const ngay = toYMD(item.ngay_nhan);
        if (fFrom && ngay < fFrom) return false;
        if (fTo && ngay > fTo) return false;
        if (fKho && item.kho !== fKho) return false;
        if (fGian && item.ma_gian !== fGian) return false;
        if (search) {
            const rowText = `${item.mvd || ''} ${item.mvd_2 || ''} ${item.ma_gian || ''} ${item.sku || ''} ${item.sku_ct || ''} ${item.ten_sp || ''} ${item.tinh_trang || ''}`.toLowerCase();
            if (!rowText.includes(search)) return false;
        }
        return true;
    }).sort((a, b) => toYMD(b.ngay_nhan).localeCompare(toYMD(a.ngay_nhan)));
    renderHangHoanTable();
}

function openHhDetail(index) {
    const item = filteredHangHoanData[index];
    const actualIndex = hangHoanData.indexOf(item);
    if (!item || actualIndex < 0) return;
    hhDrawerMode = 'edit';
    currentHangHoanEditIndex = actualIndex;
    const isKinhDoanh = currentUser && currentUser.role === 'kinhdoanh';
    document.getElementById('hhDrawerTitle').textContent = 'Chi tiết hàng hoàn';
    document.getElementById('hhSaveButton').textContent = 'Lưu thay đổi';
    document.getElementById('hhDrawerRowId').textContent = `Row ID: ${item.id || item.ngay_nhan || '-'}`;
    document.getElementById('hhEditMVD').value = item.mvd || '';
    document.getElementById('hhEditMVD2').value = item.mvd_2 || '';
    document.getElementById('hhEditMaGian').value = item.ma_gian || '';
    document.getElementById('hhEditAnh1').value = item.anh_1 || '';
    document.getElementById('hhEditAnh2').value = item.anh_2 || '';
    document.getElementById('hhEditAnh3').value = item.anh_3 || '';
    document.getElementById('hhEditSKU').value = item.sku || '';
    document.getElementById('hhEditSKUCT').value = item.sku_ct || '';
    document.getElementById('hhEditSLG').value = item.slg || '';
    document.getElementById('hhEditTinhTrang').value = item.tinh_trang || '';
    document.getElementById('hhEditTenSP').value = item.ten_sp || '';
    document.getElementById('hhEditKho').value = item.kho || '';
    populateHhFormOptions();
    renderHhKhoButtons(item.kho || 'KHO');
    refreshHhImagePreviews();
    const noticeEl = document.getElementById('hhMvdDuplicateNotice');
    if (noticeEl) noticeEl.classList.add('hidden');
    ['hhEditMVD', 'hhEditMaGian', 'hhEditSKU', 'hhEditSKUCT', 'hhEditSLG', 'hhEditTinhTrang', 'hhEditTenSP', 'hhEditKho'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.disabled = isKinhDoanh;
    });
    document.querySelectorAll('#hhEditKhoButtons button').forEach(btn => btn.disabled = isKinhDoanh);
    const footer = document.querySelector('#hhDrawer .pt-4.border-t.border-slate-200.flex.gap-3');
    if (footer) footer.style.display = isKinhDoanh ? 'none' : '';
    document.getElementById('hhDrawerOverlay').classList.remove('hidden');
    document.getElementById('hhDrawer').classList.add('open');
}

function openNewHangHoanDrawer() {
    if (currentUser && currentUser.role === 'kinhdoanh') {
        alert('Tài khoản KINHDOANH không được thêm mới Dữ liệu Hàng hoàn.');
        return;
    }
    hhDrawerMode = 'create';
    currentHangHoanEditIndex = -1;
    document.getElementById('hhDrawerTitle').textContent = 'Thêm sản phẩm hàng hoàn';
    document.getElementById('hhSaveButton').textContent = 'Thêm mới';
    document.getElementById('hhDrawerRowId').textContent = `Row ID: NEW-${Date.now()}`;
    document.getElementById('hhEditMVD').value = '';
    document.getElementById('hhEditMVD2').value = '';
    document.getElementById('hhEditMaGian').value = '';
    document.getElementById('hhEditAnh1').value = '';
    document.getElementById('hhEditAnh2').value = '';
    document.getElementById('hhEditAnh3').value = '';
    document.getElementById('hhEditSKU').value = '';
    document.getElementById('hhEditSKUCT').value = '';
    document.getElementById('hhEditSLG').value = '1';
    document.getElementById('hhEditTinhTrang').value = '';
    document.getElementById('hhEditTenSP').value = '';
    document.getElementById('hhEditKho').value = 'KHO';
    populateHhFormOptions();
    renderHhKhoButtons('KHO');
    refreshHhImagePreviews();
    const noticeEl = document.getElementById('hhMvdDuplicateNotice');
    if (noticeEl) noticeEl.classList.add('hidden');
    ['hhEditMVD', 'hhEditMaGian', 'hhEditSKU', 'hhEditSKUCT', 'hhEditSLG', 'hhEditTinhTrang', 'hhEditTenSP', 'hhEditKho'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.disabled = false;
    });
    document.querySelectorAll('#hhEditKhoButtons button').forEach(btn => btn.disabled = false);
    const footer = document.querySelector('#hhDrawer .pt-4.border-t.border-slate-200.flex.gap-3');
    if (footer) footer.style.display = '';
    document.getElementById('hhDrawerOverlay').classList.remove('hidden');
    document.getElementById('hhDrawer').classList.add('open');

    // Tự động focus vào ô MVD
    setTimeout(() => {
        const mvdInput = document.getElementById('hhEditMVD');
        if (mvdInput) mvdInput.focus();
    }, 400);
}

function closeHhDetailDrawer() {
    document.getElementById('hhDrawerOverlay').classList.add('hidden');
    document.getElementById('hhDrawer').classList.remove('open');
    currentHangHoanEditIndex = -1;
    hhDrawerMode = 'edit';
}

async function saveHhDetail() {
    if (currentUser && currentUser.role === 'kinhdoanh') {
        alert('Tài khoản KINHDOANH không được sửa Dữ liệu Hàng hoàn.');
        return;
    }
    if (isUploading) {
        console.warn('Đang upload ảnh, vui lòng đợi...');
        return;
    }
    const isCreateMode = hhDrawerMode === 'create';
    if (!isCreateMode && currentHangHoanEditIndex === -1) return;
    const item = !isCreateMode ? hangHoanData[currentHangHoanEditIndex] : null;
    if (!isCreateMode && !item) return;
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');
    try {
        const token = await getAccessToken();
        const newData = {
            mvd: document.getElementById('hhEditMVD').value,
            mvd_2: document.getElementById('hhEditMVD2').value,
            ma_gian: document.getElementById('hhEditMaGian').value,
            sku: document.getElementById('hhEditSKU').value,
            sku_ct: document.getElementById('hhEditSKUCT').value,
            slg: document.getElementById('hhEditSLG').value,
            tinh_trang: document.getElementById('hhEditTinhTrang').value,
            ten_sp: document.getElementById('hhEditTenSP').value,
            kho: document.getElementById('hhEditKho').value,
            anh_1: document.getElementById('hhEditAnh1').value,
            anh_2: document.getElementById('hhEditAnh2').value,
            anh_3: document.getElementById('hhEditAnh3').value
        };
        if (newData.sku_ct && !newData.ten_sp) {
            const matchedSp = dsSpCtData.find(i => (i.id_sp_ct || '') === newData.sku_ct);
            if (matchedSp?.ten) newData.ten_sp = matchedSp.ten;
        }
        const skuCatalog = new Set(dsSpCtData.map(i => (i.id_sp || '').trim()).filter(Boolean));
        if (newData.sku && skuCatalog.size && !skuCatalog.has(newData.sku.trim())) {
            alert('SKU không tồn tại trong DS_SP_CT (cột id_sp).');
            return;
        }
        // Bỏ chặn lưu khi thiếu SKU/SKU_CT theo yêu cầu người dùng
        if (isCreateMode) {
            const today = new Date().toISOString().split('T')[0];
            const mvd = (newData.mvd || '').trim();
            const mvd2 = (newData.mvd_2 || '').trim();
            const maGian = (newData.ma_gian || '').trim();
            const appendValues = [[
                `${Date.now()}`,
                today,
                mvd,
                mvd2,
                maGian,
                newData.anh_1 || '',
                newData.anh_2 || '',
                newData.anh_3 || '',
                '', // anh_4
                newData.sku || '',
                newData.sku_ct || '',
                newData.slg || '1',
                newData.ten_sp || '',
                '',
                newData.tinh_trang || '',
                '',
                '',
                '',
                '',
                mvd && maGian ? `${mvd}-${maGian}` : '',
                newData.kho || '',
                '',
                '',
                '',
                ''
            ]];
            const appendUrl = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.hhbhSheetName}!A:Z:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
            const appendResp = await fetch(appendUrl, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ values: appendValues })
            });
            if (!appendResp.ok) {
                const errText = await appendResp.text();
                console.error('Create HH error:', errText);
                alert('Lỗi khi thêm mới dữ liệu Hàng hoàn.');
                return;
            }
            await fetchHangHoanData();
            closeHhDetailDrawer();
            return;
        }
        const rowIndex = item.rowIndex || (hangHoanData.indexOf(item) + 2);
        const batchUpdates = [
            { range: `${CONFIG.hhbhSheetName}!C${rowIndex}`, values: [[newData.mvd]] },
            { range: `${CONFIG.hhbhSheetName}!D${rowIndex}`, values: [[newData.mvd_2]] },
            { range: `${CONFIG.hhbhSheetName}!E${rowIndex}`, values: [[newData.ma_gian]] },
            { range: `${CONFIG.hhbhSheetName}!F${rowIndex}`, values: [[newData.anh_1]] },
            { range: `${CONFIG.hhbhSheetName}!G${rowIndex}`, values: [[newData.anh_2]] },
            { range: `${CONFIG.hhbhSheetName}!H${rowIndex}`, values: [[newData.anh_3]] },
            { range: `${CONFIG.hhbhSheetName}!J${rowIndex}`, values: [[newData.sku]] },
            { range: `${CONFIG.hhbhSheetName}!K${rowIndex}`, values: [[newData.sku_ct]] },
            { range: `${CONFIG.hhbhSheetName}!L${rowIndex}`, values: [[newData.slg]] },
            { range: `${CONFIG.hhbhSheetName}!M${rowIndex}`, values: [[newData.ten_sp]] },
            { range: `${CONFIG.hhbhSheetName}!O${rowIndex}`, values: [[newData.tinh_trang]] },
            { range: `${CONFIG.hhbhSheetName}!T${rowIndex}`, values: [[(newData.mvd && newData.ma_gian) ? `${newData.mvd}-${newData.ma_gian}` : '']] },
            { range: `${CONFIG.hhbhSheetName}!U${rowIndex}`, values: [[newData.kho]] }
        ];
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
        const resp = await fetch(url, { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data: batchUpdates }) });
        if (resp.ok) {
            Object.assign(item, newData, { rowIndex });
            filterHangHoanData();
            closeHhDetailDrawer();
        } else {
            const errText = await resp.text();
            console.error('Save HH error:', errText);
            alert('Lỗi khi lưu dữ liệu Hàng hoàn vào Google Sheet.');
        }
    } catch (err) {
        console.error("Save HH Detail Error:", err);
        alert('Đã xảy ra lỗi khi lưu Hàng hoàn: ' + (err.message || 'Lỗi không xác định'));
    } finally {
        loadingOverlay.classList.add('hidden');
    }
}

function buildSkuTongMap(data) {
    const byMvd = new Map();
    const sourceData = (typeof hangHoanData !== 'undefined' && hangHoanData.length) ? hangHoanData : data;

    sourceData.forEach(item => {
        const mvd = (item.mvd || '').toString().trim();
        const ngay = (item.ngay_nhan || '').toString().trim();
        if (!mvd || !ngay) return;

        // Key kết hợp Ngày nhận và MVD
        const key = `${ngay}|${mvd}`;

        if (!byMvd.has(key)) {
            byMvd.set(key, { ma_gian: '', items: [] });
        }
        const bucket = byMvd.get(key);
        if (!bucket.ma_gian && item.ma_gian) bucket.ma_gian = (item.ma_gian || '').toString().trim();

        const skuVal = (item.sku || '').toString().trim();
        if (!skuVal) return;
        const slgVal = (item.slg || '0').toString().trim();
        bucket.items.push(`${skuVal} x ${slgVal}`);
    });

    const result = new Map();
    byMvd.forEach((bucket, key) => {
        const skuTong = bucket.items.join(' + ');
        result.set(key, { ma_gian: bucket.ma_gian || '', skuTong });
    });
    return result;
}

function getSkuTongForItem(item, skuTongMap) {
    const mvd = (item.mvd || '').toString().trim();
    const ngay = (item.ngay_nhan || '').toString().trim();
    if (!mvd || !ngay) return '';
    const key = `${ngay}|${mvd}`;
    return skuTongMap.get(key)?.skuTong || '';
}

function getMaGianForItem(item, skuTongMap) {
    const mvd = (item.mvd || '').toString().trim();
    const ngay = (item.ngay_nhan || '').toString().trim();
    if (!mvd || !ngay) return item.ma_gian || '';
    const key = `${ngay}|${mvd}`;
    return skuTongMap.get(key)?.ma_gian || item.ma_gian || '';
}

function renderHangHoanTable() {
    const tbody = document.getElementById('hangHoanTableBody');
    const stats = document.getElementById('hangHoanStats');
    if (!tbody) return;
    if (stats) stats.textContent = `Số đơn: ${filteredHangHoanData.length.toLocaleString('vi-VN')}`;
    if (!filteredHangHoanData.length) {
        tbody.innerHTML = '<tr><td colspan="11" class="text-center py-8 text-slate-500">Không có dữ liệu phù hợp bộ lọc.</td></tr>';
        return;
    }

    const skuTongMap = buildSkuTongMap(filteredHangHoanData);

    tbody.innerHTML = filteredHangHoanData.map((item, index) => {
        const imgUrl = (item.anh_3 || '').trim();
        const imgHtml = imgUrl
            ? `<img src="${escapeHtml(imgUrl)}" alt="Ảnh hàng hoàn" class="w-12 h-12 object-cover rounded border border-slate-200 cursor-pointer hover:opacity-80" onclick="openImagePreview('${escapeHtml(imgUrl)}')">`
            : '<span class="text-xs text-slate-400">Không có</span>';
        const skuTong = getSkuTongForItem(item, skuTongMap);
        const maGian = getMaGianForItem(item, skuTongMap);
        const displayNgay = formatYmdToDmy(item.ngay_nhan) || item.ngay_nhan;
        return `
                    <tr class="border-b border-slate-100 hover:bg-slate-50 transition-colors cursor-pointer" ondblclick="openHhDetail(${index})">
                        <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(displayNgay)}</td>
                        <td class="px-3 py-2 text-sm font-medium text-slate-900">${escapeHtml(item.mvd)}</td>
                        <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.mvd_2)}</td>
                        <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(maGian)}</td>
                        <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.sku)}</td>
                        <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.sku_ct)}</td>
                        <td class="px-3 py-2 text-sm text-right font-semibold text-slate-900">${(parseFloat(item.slg) || 0).toLocaleString('vi-VN')}</td>
                        <td class="px-3 py-2 text-sm text-slate-700 max-w-[240px] truncate" title="${escapeHtml(item.ten_sp)}">${escapeHtml(item.ten_sp)}</td>
                        <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.kho)}</td>
                        <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.tinh_trang)}</td>
                        <td class="px-3 py-2 text-sm text-slate-700 max-w-[260px] truncate" title="${escapeHtml(skuTong)}">${escapeHtml(skuTong || '-')}</td>
                        <td class="px-3 py-2 text-sm">${imgHtml}</td>
                    </tr>
                `;
    }).join('');
}

function exportHangHoanSummaryToExcel() {
    if (!filteredHangHoanData || !filteredHangHoanData.length) {
        alert('Không có dữ liệu hợp lệ để xuất!');
        return;
    }

    const skuTongMap = buildSkuTongMap(filteredHangHoanData);

    // Lấy danh sách các cặp (Ngày | MVD) duy nhất từ dữ liệu đang lọc
    const uniqueKeys = [...new Set(filteredHangHoanData
        .map(item => `${(item.ngay_nhan || '').toString().trim()}|${(item.mvd || '').toString().trim()}`)
        .filter(k => k.split('|')[1]))];

    const headers = ['Ngày nhận', 'MVD', 'MVD 2', 'Mã gian', 'SKU tổng'];
    const excelData = [headers, ...uniqueKeys.map(key => {
        const [ngay, mvd] = key.split('|');
        const info = skuTongMap.get(key) || { ma_gian: '', skuTong: '' };
        const displayNgay = formatYmdToDmy(ngay) || ngay;
        // Tìm dòng đầu tiên khớp với cặp Ngày-MVD để lấy MVD 2
        const firstRow = filteredHangHoanData.find(item =>
            (item.ngay_nhan || '').toString().trim() === ngay &&
            (item.mvd || '').toString().trim() === mvd
        ) || {};
        return [displayNgay, mvd, firstRow.mvd_2 || '', info.ma_gian || '', info.skuTong || ''];
    })];

    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'MVD_SKU_Tong');
    const filterDate = document.getElementById('filterHHNgayNhan')?.value || 'TatCa';
    const now = new Date();
    const timeStr = now.getHours().toString().padStart(2, '0') + 'h' + now.getMinutes().toString().padStart(2, '0');
    XLSX.writeFile(wb, `MVD_SKU_Tong_${filterDate}_${timeStr}.xlsx`);
}

function exportHangHoanToExcel() {
    if (!filteredHangHoanData || !filteredHangHoanData.length) {
        alert('Không có dữ liệu hợp lệ để xuất!');
        return;
    }

    const headers = ['Ngày nhận', 'MVD', 'MVD 2', 'Mã gian', 'SKU', 'SKU CT', 'SKU tổng', 'SLG', 'Tên SP', 'Tình trạng', 'Kho', 'Ảnh 1', 'Ảnh 2', 'Ảnh 3', 'Ngày xử lý', 'Ghi chú', 'Trạng thái', 'SKU-SLG', 'ID NV', 'UDT', 'MVD-Gian', 'LB3', 'ID ĐH', 'ID ĐH CT', 'STT', 'Đánh dấu'];
    const skuTongMap = buildSkuTongMap(filteredHangHoanData);
    const excelData = [headers, ...filteredHangHoanData.map(item => [
        formatYmdToDmy(item.ngay_nhan) || item.ngay_nhan, item.mvd || '', item.mvd_2 || '', item.ma_gian || '', item.sku || '', item.sku_ct || '', getSkuTongForItem(item, skuTongMap) || '', item.slg || '', item.ten_sp || '', item.tinh_trang || '', item.kho || '',
        item.anh_1 || '', item.anh_2 || '', item.anh_3 || '', item.ngay_xly || '', item.ghi_chu || '', item.trang_thai || '', item.sku_slg || '', item.id_nv || '', item.udt || '',
        item.mvd_gian || '', item.lb3 || '', item.id_dh || '', item.id_dh_ct || '', item.stt || '', item.danh_dau || ''
    ])];

    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'HH_BH');
    const filterDate = document.getElementById('filterHHNgayNhan')?.value || 'TatCa';
    const now = new Date();
    const timeStr = now.getHours().toString().padStart(2, '0') + 'h' + now.getMinutes().toString().padStart(2, '0');
    XLSX.writeFile(wb, `HH_BH_${filterDate}_${timeStr}.xlsx`);
}

function exportHangHoanToMisa() {
    if (!filteredHangHoanData || !filteredHangHoanData.length) {
        alert('Không có dữ liệu hợp lệ để xuất MISA!');
        return;
    }

    const headers = ['Hiển thị trên sổ', 'Hình thức bán hàng', 'Phương thức thanh toán', 'Kiêm phiếu xuất kho', 'Lập kèm hóa đơn', 'Đã lập hóa đơn', 'Ngày hạch toán (*)', 'Ngày chứng từ (*)', 'Số chứng từ (*)', 'Số phiếu xuất', 'Lý do xuất', 'Số hóa đơn', 'Ngày hóa đơn', 'Mã đơn hàng', 'Mã thống kê', 'Mã khách hàng', 'Tên khách hàng', 'Địa chỉ', 'Mã số thuế', 'Diễn giải', 'Nộp vào TK', 'NV bán hàng', 'Mã hàng (*)', 'Tên hàng', 'Hàng khuyến mại', 'TK Tiền/Chi phí/Nợ (*)', 'TK Doanh thu/Có (*)', 'ĐVT', 'Số lượng', 'Đơn giá sau thuế', 'Đơn giá', 'Thành tiền', 'Tỷ lệ CK (%)', 'Tiền chiết khấu', 'TK chiết khấu', 'Giá tính thuế XK', '% thuế XK', 'Tiền thuế XK', 'TK thuế XK', '% thuế GTGT', 'Tiền thuế GTGT', 'TK thuế GTGT', 'HH không TH trên tờ khai thuế GTGT', 'Kho', 'TK giá vốn', 'TK Kho', 'Đơn giá vốn', 'Tiền vốn', 'Hàng hóa giữ hộ/bán hộ'];

    const formatDateVN = (input) => {
        if (!input) return '';
        const raw = String(input).split(' ')[0];
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) return raw;
        if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
            const [y, m, d] = raw.split('-');
            return `${d}/${m}/${y}`;
        }
        const dt = new Date(input);
        if (isNaN(dt.getTime())) return '';
        const d = String(dt.getDate()).padStart(2, '0');
        const m = String(dt.getMonth() + 1).padStart(2, '0');
        const y = dt.getFullYear();
        return `${d}/${m}/${y}`;
    };

    const formatCertDate = (input) => {
        const vn = formatDateVN(input);
        if (!vn) return '';
        const [d, m, y] = vn.split('/');
        return `${d}${m}${String(y).slice(-2)}`;
    };

    const getMisaValue = (item, colName) => {
        const ngayVN = formatDateVN(item.ngay_nhan);
        const certDate = formatCertDate(item.ngay_nhan);

        if (colName === 'Hiển thị trên sổ') return '0';
        if (colName === 'Hình thức bán hàng') return '0';
        if (colName === 'Phương thức thanh toán') return '0';
        if (colName === 'Kiêm phiếu xuất kho') return '1';
        if (colName === 'Lập kèm hóa đơn') return '0';
        if (colName === 'Đã lập hóa đơn') return '0';
        if (colName === 'Ngày hạch toán (*)' || colName === 'Ngày chứng từ (*)') return ngayVN;
        if (colName === 'Số chứng từ (*)' || colName === 'Số phiếu xuất') return item.mvd ? `HH-${item.mvd}-${certDate}` : '';
        if (colName === 'Lý do xuất') return 'Hàng hoàn';
        if (colName === 'Mã đơn hàng') return item.id_dh || '';
        if (colName === 'Mã thống kê') return item.ma_gian || '';
        if (colName === 'Mã khách hàng') return item.ma_gian || '';
        if (colName === 'Tên khách hàng') return item.ma_gian || '';
        if (colName === 'Diễn giải') return `${item.kho || ''} HÀNG HOÀN NGÀY ${ngayVN}`.trim();
        if (colName === 'Nộp vào TK') return '';
        if (colName === 'NV bán hàng') return '';
        if (colName === 'Mã hàng (*)') return item.sku_ct || item.sku || '';
        if (colName === 'Tên hàng') return item.ten_sp || '';
        if (colName === 'Hàng khuyến mại') return '';
        if (colName === 'TK Tiền/Chi phí/Nợ (*)') return '131';
        if (colName === 'TK Doanh thu/Có (*)') return '5111';
        if (colName === 'ĐVT') return 'Cái';
        if (colName === 'Số lượng') return item.slg || '';
        if (colName === 'Đơn giá sau thuế') return '';
        if (colName === 'Đơn giá') return '';
        if (colName === 'Thành tiền') return '';
        if (colName === 'Tỷ lệ CK (%)') return '';
        if (colName === 'Tiền chiết khấu') return '';
        if (colName === 'TK chiết khấu') return '';
        if (colName === 'Giá tính thuế XK') return '';
        if (colName === '% thuế XK') return '';
        if (colName === 'Tiền thuế XK') return '';
        if (colName === 'TK thuế XK') return '';
        if (colName === '% thuế GTGT') return '';
        if (colName === 'Tiền thuế GTGT') return '';
        if (colName === 'TK thuế GTGT') return '';
        if (colName === 'HH không TH trên tờ khai thuế GTGT') return '';
        if (colName === 'Kho') return item.kho || '';
        if (colName === 'TK giá vốn') return '632';
        if (colName === 'TK Kho') return '1561';
        if (colName === 'Đơn giá vốn') return '';
        if (colName === 'Tiền vốn') return '';
        if (colName === 'Hàng hóa giữ hộ/bán hộ') return '';
        return '';
    };

    const rows = filteredHangHoanData.map(item => headers.map(h => getMisaValue(item, h)));
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'MISA');
    const filterDate = document.getElementById('filterHHNgayNhan')?.value || 'TatCa';
    const now = new Date();
    const timeStr = now.getHours().toString().padStart(2, '0') + 'h' + now.getMinutes().toString().padStart(2, '0');
    XLSX.writeFile(wb, `MISA_HH_BH_${filterDate}_${timeStr}.xlsx`);
}

function toggleSidebar() {
    if (window.innerWidth <= 1024) {
        toggleMobileSidebar();
        return;
    }
    const sidebar = document.getElementById('sidebar');
    const icon = document.getElementById('sidebarToggleIcon');
    const isCollapsed = sidebar.classList.contains('sidebar-collapsed');
    if (isCollapsed) {
        sidebar.classList.remove('sidebar-collapsed', 'w-16');
        sidebar.classList.add('w-64');
        icon.innerHTML = '<path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 19l-7-7 7-7m8 14l-7-7 7-7" />';
    } else {
        sidebar.classList.add('sidebar-collapsed', 'w-16');
        sidebar.classList.remove('w-64');
        icon.innerHTML = '<path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 5l7 7-7 7M5 5l7 7-7 7" />';
    }
}

function getTodayYmd() {
    return new Date().toISOString().split('T')[0];
}

function formatYmdToDmy(ymd) {
    if (!ymd || !/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return '';
    const [y, m, d] = ymd.split('-');
    return `${d}/${m}/${y}`;
}

function parseDmyToYmd(dmy) {
    if (!dmy) return '';
    const v = String(dmy).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(v)) {
        const [d, m, y] = v.split('/');
        return `${y}-${m}-${d}`;
    }
    return '';
}

function getCurrentWeekRangeYmd() {
    const now = new Date();
    const day = now.getDay() || 7; // Monday = 1, Sunday = 7
    const start = new Date(now);
    start.setHours(0, 0, 0, 0);
    start.setDate(now.getDate() - day + 1);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    const toYmd = (date) => date.toISOString().split('T')[0];
    return { from: toYmd(start), to: toYmd(end) };
}

function getUdctSummaryByMvd(mvd) {
    const key = (mvd || '').toString().trim();
    if (!key) return { mdh: '', ma_gian: '', sku: '' };
    const rows = udctData.filter(i => (i.mvd || '').toString().trim() === key);
    const mdh = [...new Set(rows.map(i => (i.mdh || '').toString().trim()).filter(Boolean))].join(', ');
    const maGian = [...new Set(rows.map(i => (i.ma_gian || '').toString().trim()).filter(Boolean))].join(', ');
    const sku = [...new Set(rows.map(i => (i.id_sp || '').toString().trim()).filter(Boolean))].join(', ');
    return { mdh, ma_gian: maGian, sku };
}

function refreshHHShopAutoFields() {
    const mvd = (document.getElementById('hhShopEditMVD')?.value || '').toString().trim();
    const hoanTra = (document.getElementById('hhShopEditHoanTra')?.value || '').toString().trim();
    const mvdInfo = getUdctSummaryByMvd(mvd);
    const hasLinkedMvd = !!(mvd && (mvdInfo.mdh || mvdInfo.ma_gian || mvdInfo.sku));
    const currentMdh = (document.getElementById('hhShopEditMDH')?.value || '').toString().trim();
    const currentMaGian = (document.getElementById('hhShopEditMaGian')?.value || '').toString().trim();
    const currentSku = (document.getElementById('hhShopEditSKU')?.value || '').toString().trim();
    if (hasLinkedMvd) {
        if (!currentMdh) document.getElementById('hhShopEditMDH').value = (mvdInfo.mdh || '').trim();
        if (!currentMaGian) document.getElementById('hhShopEditMaGian').value = (mvdInfo.ma_gian || '').trim();
        if (!currentSku) document.getElementById('hhShopEditSKU').value = (mvdInfo.sku || '').trim();
        document.getElementById('hhShopDrawerRowId').textContent = `✅ ${mvd}`;
    } else {
        document.getElementById('hhShopDrawerRowId').textContent = mvd ? `⚠️ ${mvd}` : 'NEW';
    }
    const mvdTraEl = document.getElementById('hhShopEditMVDTra');
    if (mvdTraEl && !mvdTraEl.dataset.manual) {
        mvdTraEl.value = hoanTra === 'hoàn' ? mvd : '';
    }
    document.getElementById('hhShopEditSKUTra').value = (document.getElementById('hhShopEditSKU')?.value || '').toString().trim();
    const rawDate = document.getElementById('hhShopEditNgayTraRaw').value || getTodayYmd();
    document.getElementById('hhShopEditNgayTraRaw').value = rawDate;
    document.getElementById('hhShopEditNgayTra').value = formatYmdToDmy(rawDate);
    setHHShopButtonGroup('hhShopHoanTraButtons', hoanTra || 'hoàn');
}

function saveHHShopFilterState() {
    const from = document.getElementById('hhShopNgayTraFrom')?.value || '';
    const to = document.getElementById('hhShopNgayTraTo')?.value || '';
    const search = document.getElementById('hhShopSearchMvd')?.value || '';
    const xacNhan = document.getElementById('hhShopXacNhanFilter')?.value || '';
    const daNhan = document.getElementById('hhShopDaNhanFilterButtons')?.dataset.value || '';
    localStorage.setItem('hhShopDienFilterState', JSON.stringify({ from, to, search, xacNhan, daNhan }));
}

function loadHHShopFilterState() {
    try {
        return JSON.parse(localStorage.getItem('hhShopDienFilterState') || '{}') || {};
    } catch {
        return {};
    }
}

function applyHHShopFilterState() {
    const state = loadHHShopFilterState();
    const fromEl = document.getElementById('hhShopNgayTraFrom');
    const toEl = document.getElementById('hhShopNgayTraTo');
    const searchEl = document.getElementById('hhShopSearchMvd');
    const xacNhanButtons = document.getElementById('hhShopXacNhanFilterButtons');
    const daNhanButtons = document.getElementById('hhShopDaNhanFilterButtons');
    if (fromEl && state.from !== undefined) fromEl.value = state.from || '';
    if (toEl && state.to !== undefined) toEl.value = state.to || '';
    if (searchEl && state.search !== undefined) searchEl.value = state.search || '';
    if (xacNhanButtons && state.xacNhan !== undefined) xacNhanButtons.dataset.value = state.xacNhan || '';
    if (daNhanButtons && state.daNhan !== undefined) daNhanButtons.dataset.value = state.daNhan || '';
}

function setHHShopDateFilter(from, to) {
    const fromEl = document.getElementById('hhShopNgayTraFrom');
    const toEl = document.getElementById('hhShopNgayTraTo');
    if (fromEl) fromEl.value = from || '';
    if (toEl) toEl.value = to || '';
    saveHHShopFilterState();
    renderHHShopDienTable();
}

function setHHShopNgayTraFilterToday() {
    const today = getTodayYmd();
    setHHShopDateFilter(today, today);
}

function setHHShopNgayTraFilterThisWeek() {
    const { from, to } = getCurrentWeekRangeYmd();
    setHHShopDateFilter(from, to);
}

function clearHHShopNgayTraFilter() {
    setHHShopDateFilter('', '');
}

function syncHHShopButtonFilter(root, value) {
    if (!root) return;
    const nextValue = value || '';
    root.dataset.value = nextValue;
    root.querySelectorAll('button').forEach(btn => {
        const active = (btn.dataset.value || '') === nextValue;
        btn.classList.toggle('bg-slate-100', active);
        btn.classList.toggle('bg-white', !active);
        btn.classList.toggle('text-primary', active && nextValue === '');
    });
}

function setHHShopButtonFilter(containerId, value) {
    const root = document.getElementById(containerId);
    syncHHShopButtonFilter(root, value);
    saveHHShopFilterState();
    renderHHShopDienTable();
}

function setHHShopXacNhanFilter(value) {
    setHHShopButtonFilter('hhShopXacNhanFilterButtons', value);
}

function setHHShopDaNhanFilter(value) {
    setHHShopButtonFilter('hhShopDaNhanFilterButtons', value);
}

function markHHShopMvdTraManual(isManual) {
    const el = document.getElementById('hhShopEditMVDTra');
    if (el) el.dataset.manual = isManual ? '1' : '';
}

function scheduleHHShopAutoSave() {
    if (currentHHShopRowIndex < 0) return; // Không tự động lưu khi đang thêm mới
    clearTimeout(window.hhShopAutoSaveTimer);
    window.hhShopAutoSaveTimer = setTimeout(() => {
        saveHHShopDien();
    }, 300);
}

function syncHHShopButtonGroup(root, activeValue) {
    if (!root) return;
    const current = (activeValue || '').toString().trim().toLowerCase();
    root.querySelectorAll('button').forEach(btn => {
        const val = (btn.dataset.value || btn.textContent || '').toString().trim().toLowerCase();
        const active = current === val;
        btn.classList.toggle('bg-slate-100', active);
        btn.classList.toggle('bg-white', !active);
        btn.classList.toggle('ring-2', active);
        btn.classList.toggle('ring-primary/20', active);
    });
}

function setHHShopButtonGroup(containerId, activeValue) {
    syncHHShopButtonGroup(document.getElementById(containerId), activeValue);
}

async function setHHShopXacNhan(value) {
    const nextValue = value || '';
    document.getElementById('hhShopEditXacNhan').value = nextValue;
    setHHShopButtonGroup('hhShopXacNhanButtons', nextValue || 'trống');
    if (currentHHShopRowIndex < 0) return;
    const item = hhShopDienData[currentHHShopRowIndex];
    if (!item) return;
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');
    try {
        const token = await getAccessToken();
        if (!token) return;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
        const body = {
            valueInputOption: 'USER_ENTERED',
            data: [
                { range: `${CONFIG.hhNvDienSheetName}!K${item.rowIndex}`, values: [[nextValue]] }
            ]
        };
        const resp = await fetch(url, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        if (!resp.ok) throw new Error(await resp.text());
        item.xac_nhan = nextValue;
        renderHHShopDienTable();
    } catch (error) {
        console.error(error);
        alert('Lỗi khi cập nhật xác nhận.');
    } finally {
        loadingOverlay.classList.add('hidden');
    }
}

async function deleteHHShopDien() {
    if (currentHHShopRowIndex < 0) return alert('Không xác định được dòng cần xóa.');
    const item = hhShopDienData[currentHHShopRowIndex];
    if (!item) return alert('Không xác định được dòng cần xóa.');
    if (!confirm(`Xóa dòng HH SHOP ĐIỀN này? (Row ${item.rowIndex})`)) return;
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');
    try {
        const token = await getAccessToken();
        if (!token) return;
        const sheetId = await fetchSheetMeta(CONFIG.hhNvDienSheetName, token);
        if (sheetId === null || sheetId === undefined) throw new Error('Không lấy được sheetId');
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}:batchUpdate`;
        const body = {
            requests: [{
                deleteDimension: {
                    range: {
                        sheetId,
                        dimension: 'ROWS',
                        startIndex: item.rowIndex - 1,
                        endIndex: item.rowIndex
                    }
                }
            }]
        };
        const resp = await fetch(url, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        if (!resp.ok) {
            console.error('Delete HH SHOP error:', await resp.text());
            alert('Lỗi khi xóa dòng HH SHOP ĐIỀN.');
            return;
        }
        closeHHShopDrawer();
        await fetchHHShopDienData();
    } catch (error) {
        console.error(error);
        alert('Có lỗi khi xóa HH SHOP ĐIỀN.');
    } finally {
        loadingOverlay.classList.add('hidden');
    }
}

async function fetchSheetMeta(sheetName, token) {
    const resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}?fields=sheets(properties(sheetId,title))`, {
        headers: { 'Authorization': `Bearer ${token}` }
    });
    const data = await resp.json();
    const sheet = (data.sheets || []).find(s => s.properties?.title === sheetName);
    return sheet?.properties?.sheetId ?? null;
}

function setHHShopHoanTra(value) {
    document.getElementById('hhShopEditHoanTra').value = value;
    const mvdTraEl = document.getElementById('hhShopEditMVDTra');
    if (mvdTraEl) mvdTraEl.dataset.manual = '';
    setHHShopButtonGroup('hhShopHoanTraButtons', value);
    refreshHHShopAutoFields();
}



function handleHHShopMvdChange() {
    refreshHHShopAutoFields();
}

function setHHShopNgayTraRaw(ymdValue) {
    const ymd = ymdValue || getTodayYmd();
    const rawEl = document.getElementById('hhShopEditNgayTraRaw');
    const dmyEl = document.getElementById('hhShopEditNgayTra');
    const picker = document.getElementById('hhShopNgayTraPicker');
    if (rawEl) rawEl.value = ymd;
    if (dmyEl) dmyEl.value = formatYmdToDmy(ymd);
    if (picker) picker.value = ymd;
}

function openHHShopNgayTraPicker() {
    const picker = document.getElementById('hhShopNgayTraPicker');
    if (!picker) return;
    picker.value = document.getElementById('hhShopEditNgayTraRaw').value || getTodayYmd();
    picker.style.position = 'fixed';
    picker.style.left = '0';
    picker.style.top = '0';
    picker.style.zIndex = '99999';
    picker.showPicker?.();
    if (!picker.showPicker) picker.click();
}

function getTodayYmd() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

function parseDmyToYmd(dmy) {
    if (!dmy) return '';
    const parts = String(dmy).trim().split('/');
    if (parts.length !== 3) return '';
    return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
}

// Duplicate formatYmdToDmy logic removed here to use the consolidated version at top

function shiftDateValueByDays(currentValue, deltaDays) {
    const current = currentValue || getTodayYmd();
    const dt = new Date(`${current}T00:00:00`);
    dt.setDate(dt.getDate() + deltaDays);
    const yyyy = dt.getFullYear();
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const dd = String(dt.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

function changeHHShopNgayTra(step) {
    const currentRaw = document.getElementById('hhShopEditNgayTraRaw')?.value || getTodayYmd();
    setHHShopNgayTraRaw(shiftDateValueByDays(currentRaw, step));
    return false;
}

function shiftHHFilterDate(id, delta) {
    const el = document.getElementById(id);
    if (!el) return;
    el.value = shiftDateValueByDays(el.value, delta);
    filterHangHoanData();
}

function shiftHHShopNgayTra(which, step) {
    const inputId = which === 'from' ? 'hhShopNgayTraFrom' : 'hhShopNgayTraTo';
    const input = document.getElementById(inputId);
    if (!input) return false;
    input.value = shiftDateValueByDays(input.value || getTodayYmd(), step);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    return false;
}

function openHHShopNewDrawer() {
    currentHHShopRowIndex = -1;
    document.getElementById('hhShopEditMVD').value = '';
    document.getElementById('hhShopEditMDH').value = '';
    document.getElementById('hhShopEditMaGian').value = '';
    document.getElementById('hhShopEditSKU').value = '';
    document.getElementById('hhShopEditHoanTra').value = 'hoàn';
    setHHShopNgayTraRaw(getTodayYmd());
    const mvdTraEl = document.getElementById('hhShopEditMVDTra');
    if (mvdTraEl) {
        mvdTraEl.value = '';
        mvdTraEl.dataset.manual = '';
    }
    document.getElementById('hhShopEditSKUTra').value = '';
    document.getElementById('hhShopEditGhiChu').value = '';
    document.getElementById('hhShopEditXacNhan').value = '';
    setHHShopButtonGroup('hhShopHoanTraButtons', 'hoàn');
    setHHShopButtonGroup('hhShopXacNhanButtons', 'trống');
    refreshHHShopAutoFields();
    document.getElementById('hhShopDrawerOverlay').classList.remove('hidden');
    document.getElementById('hhShopDrawer').classList.add('open');
    setTimeout(() => document.getElementById('hhShopEditMVD')?.focus(), 30);
}

function closeHHShopDrawer() {
    document.getElementById('hhShopDrawerOverlay').classList.add('hidden');
    document.getElementById('hhShopDrawer').classList.remove('open');
    const delBtn = document.getElementById('hhShopDeleteButton');
    if (delBtn) delBtn.classList.add('hidden');
}

async function saveHHShopDien() {
    const mvd = (document.getElementById('hhShopEditMVD').value || '').toString().trim();
    if (!mvd) return alert('Vui lòng nhập MVD.');
    const mdh = document.getElementById('hhShopEditMDH').value || '';
    const maGian = document.getElementById('hhShopEditMaGian').value || '';
    const sku = document.getElementById('hhShopEditSKU').value || '';
    const hoanTra = document.getElementById('hhShopEditHoanTra').value || 'hoàn';
    const ngayTraRaw = document.getElementById('hhShopEditNgayTraRaw').value || parseDmyToYmd(document.getElementById('hhShopEditNgayTra').value) || getTodayYmd();
    const ngayTra = formatYmdToDmy(ngayTraRaw);
    const mvdTraEl = document.getElementById('hhShopEditMVDTra');
    const mvdTraManual = (mvdTraEl?.dataset.manual || '') === '1';
    const mvdTra = mvdTraManual ? (mvdTraEl?.value || '').trim() : (hoanTra === 'hoàn' ? mvd : '');
    const skuTra = document.getElementById('hhShopEditSKUTra').value || sku;
    const rawGhiChu = (document.getElementById('hhShopEditGhiChu').value || '').trim();
    const ghiChu = rawGhiChu || `[${hoanTra.toUpperCase()}] MVD ${mvd}${mdh ? ` | MDH: ${mdh}` : ''}`;
    const xacNhan = (document.getElementById('hhShopEditXacNhan').value || '').toString().trim();

    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');
    try {
        const token = await getAccessToken();
        if (!token) return;
        const values = [[`${Date.now()}`, mvd, mdh, maGian, sku, hoanTra, ngayTra, mvdTra, skuTra, ghiChu, xacNhan, ""]];
        const isUpdate = currentHHShopRowIndex >= 0 && hhShopDienData[currentHHShopRowIndex];
        const url = isUpdate
            ? `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.hhNvDienSheetName}!A${hhShopDienData[currentHHShopRowIndex].rowIndex}:L${hhShopDienData[currentHHShopRowIndex].rowIndex}?valueInputOption=USER_ENTERED`
            : `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.hhNvDienSheetName}!A:L:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
        const method = isUpdate ? 'PUT' : 'POST';
        const body = isUpdate ? { values } : { values };
        const resp = await fetch(url, {
            method,
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        if (!resp.ok) {
            const errText = await resp.text();
            console.error('Save HH SHOP ĐIỀN error:', errText);
            alert('Lỗi khi lưu HH SHOP ĐIỀN.');
            return;
        }
        closeHHShopDrawer();
        await fetchHHShopDienData();
    } catch (error) {
        console.error(error);
        alert('Có lỗi khi lưu HH SHOP ĐIỀN.');
    } finally {
        loadingOverlay.classList.add('hidden');
    }
}

function openHHShopDetail(rowIndex) {
    const idx = hhShopDienData.findIndex(i => Number(i.rowIndex) === Number(rowIndex));
    const item = idx >= 0 ? hhShopDienData[idx] : null;
    if (!item) return;
    currentHHShopRowIndex = idx;
    document.getElementById('hhShopEditMVD').value = item.mvd || '';
    document.getElementById('hhShopEditMDH').value = item.mdh || '';
    document.getElementById('hhShopEditMaGian').value = item.ma_gian || '';
    document.getElementById('hhShopEditSKU').value = item.sku || '';
    document.getElementById('hhShopEditHoanTra').value = item.hoan_tra || 'hoàn';
    document.getElementById('hhShopEditNgayTraRaw').value = parseDmyToYmd(item.ngay_tra) || getTodayYmd();
    document.getElementById('hhShopEditNgayTra').value = item.ngay_tra || formatYmdToDmy(getTodayYmd());
    const mvdTraEl = document.getElementById('hhShopEditMVDTra');
    if (mvdTraEl) {
        mvdTraEl.value = item.mvd_tra || '';
        mvdTraEl.dataset.manual = item.mvd_tra ? '1' : '';
    }
    document.getElementById('hhShopEditSKUTra').value = item.sku_tra || '';
    document.getElementById('hhShopEditGhiChu').value = item.ghi_chu || '';
    document.getElementById('hhShopDrawerRowId').textContent = `Row ID: ${item.rowIndex}`;
    document.getElementById('hhShopDrawerOverlay').classList.remove('hidden');
    document.getElementById('hhShopDrawer').classList.add('open');
    const delBtn = document.getElementById('hhShopDeleteButton');
    if (delBtn) delBtn.classList.remove('hidden');
    refreshHHShopAutoFields();
}

function closeHHShopDrawer() {
    document.getElementById('hhShopDrawerOverlay').classList.add('hidden');
    document.getElementById('hhShopDrawer').classList.remove('open');
}

async function fetchHHShopDienData() {
    const tbody = document.getElementById('hhShopTableBody');
    if (tbody) tbody.innerHTML = '<tr><td colspan="11" class="text-center py-8 text-slate-500">Đang tải dữ liệu...</td></tr>';
    try {
        const token = await getAccessToken();
        if (!token) return;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.hhNvDienSheetName}!A:L`;
        const response = await fetch(url, { headers: { "Authorization": `Bearer ${token}` } });
        const result = await response.json();
        const hhbhData = await fetchSheetData(CONFIG.hhbhSheetName);
        hhBhMvdSet = new Set((hhbhData || [])
            .slice(1)
            .flatMap(row => [row[2], row[3]]
                .map(v => (v || '').toString().trim())
                .filter(Boolean)));
        if (result.values && result.values.length > 1) {
            hhShopDienData = result.values.slice(1).map((row, idx) => ({
                rowIndex: idx + 2,
                id: row[0] || '',
                mvd: row[1] || '',
                mdh: row[2] || '',
                ma_gian: row[3] || '',
                sku: row[4] || '',
                hoan_tra: row[5] || '',
                ngay_tra: row[6] || '',
                mvd_tra: row[7] || '',
                sku_tra: row[8] || '',
                ghi_chu: row[9] || '',
                xac_nhan: row[10] || '',
                udt: row[11] || ''
            }));
        } else {
            hhShopDienData = [];
        }
        applyHHShopFilterState();
        renderHHShopDienTable();
    } catch (error) {
        console.error('Load HH SHOP ĐIỀN error:', error);
        if (tbody) tbody.innerHTML = '<tr><td colspan="11" class="text-center py-8 text-slate-500">Không thể tải dữ liệu HH_NV_DIEN.</td></tr>';
    }
}

function renderHHShopDienTable() {
    const tbody = document.getElementById('hhShopTableBody');
    const stats = document.getElementById('hhShopStats');
    if (!tbody) return;
    const search = (document.getElementById('hhShopSearchMvd')?.value || '').toLowerCase().trim();
    const fromYmd = document.getElementById('hhShopNgayTraFrom')?.value || '';
    const toYmd = document.getElementById('hhShopNgayTraTo')?.value || '';
    const xacNhanFilter = (document.getElementById('hhShopXacNhanFilterButtons')?.dataset.value || '').toLowerCase().trim();
    const daNhanFilter = (document.getElementById('hhShopDaNhanFilterButtons')?.dataset.value || '').toLowerCase().trim();
    saveHHShopFilterState();
    filteredHHShopDienData = hhShopDienData
        .slice()
        .sort((a, b) => {
            const ay = parseDmyToYmd(a.ngay_tra) || '';
            const by = parseDmyToYmd(b.ngay_tra) || '';
            if (ay !== by) return by.localeCompare(ay);
            return Number(b.rowIndex || 0) - Number(a.rowIndex || 0);
        })
        .filter(item => {
            if (search) {
                const matchSearch = [
                    item.mvd,
                    item.mdh,
                    item.ma_gian,
                    item.sku,
                    item.hoan_tra,
                    item.ngay_tra,
                    item.mvd_tra,
                    item.sku_tra,
                    item.ghi_chu,
                    item.xac_nhan
                ].some(value => (value || '').toString().toLowerCase().includes(search));
                if (!matchSearch) return false;
            }
            const ngayTraYmd = parseDmyToYmd(item.ngay_tra);
            if (fromYmd && (!ngayTraYmd || ngayTraYmd < fromYmd)) return false;
            if (toYmd && (!ngayTraYmd || ngayTraYmd > toYmd)) return false;
            if (xacNhanFilter) {
                const itemXacNhan = (item.xac_nhan || '').toString().toLowerCase().trim();
                const normalized = xacNhanFilter === 'trống' ? '' : xacNhanFilter;
                if (xacNhanFilter === 'trống') {
                    if (itemXacNhan) return false;
                } else if (itemXacNhan !== normalized) {
                    return false;
                }
            }
            if (daNhanFilter) {
                const isCoDon = hhBhMvdSet.has((item.mvd || '').toString().trim()) || hhBhMvdSet.has((item.mvd_tra || '').toString().trim());
                if (daNhanFilter === 'có đơn' && !isCoDon) return false;
                if (daNhanFilter === 'trống' && isCoDon) return false;
            }
            return true;
        });
    if (stats) stats.textContent = `Số dòng: ${filteredHHShopDienData.length.toLocaleString('vi-VN')}`;
    if (!filteredHHShopDienData.length) {
        tbody.innerHTML = '<tr><td colspan="11" class="text-center py-8 text-slate-500">Không có dữ liệu.</td></tr>';
        return;
    }
    tbody.innerHTML = filteredHHShopDienData.map(item => {
        const isCoDon = hhBhMvdSet.has((item.mvd || '').toString().trim()) || hhBhMvdSet.has((item.mvd_tra || '').toString().trim());
        return `
                <tr ondblclick="openHHShopDetail(${item.rowIndex})" class="border-b border-slate-100 hover:bg-slate-50 cursor-pointer">
                    <td class="px-3 py-2 text-sm text-slate-900 font-medium">${item.mvd ? `✅ ${escapeHtml(item.mvd)}` : '-'}</td>
                    <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.mdh || '-')}</td>
                    <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.ma_gian || '-')}</td>
                    <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.sku || '-')}</td>
                    <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.hoan_tra || '-')}</td>
                    <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.ngay_tra || '-')}</td>
                    <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.mvd_tra || '-')}</td>
                    <td class="px-3 py-2 text-sm text-slate-700">${escapeHtml(item.sku_tra || '-')}</td>
                    <td class="px-3 py-2 text-sm text-slate-700 max-w-[220px] truncate" title="${escapeHtml(item.ghi_chu || '')}">${escapeHtml(item.ghi_chu || '-')}</td>
                    <td class="px-3 py-2 text-sm text-slate-700">
                        <button type="button" onclick="event.stopPropagation(); setHHShopXacNhan('${(item.xac_nhan || '').replace(/'/g, "\\'")}')" class="px-3 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold ${item.xac_nhan ? 'bg-slate-100' : 'bg-white'}">${escapeHtml(item.xac_nhan || 'Trống')}</button>
                    </td>
                    <td class="px-3 py-2 text-sm text-slate-700">${isCoDon ? 'có đơn' : ''}</td>
                </tr>`;
    }).join('');
}

function syncMobileSidebar(isOpen) {
    const sidebar = document.getElementById('sidebar');
    const backdrop = document.getElementById('mobileSidebarBackdrop');
    if (!sidebar || !backdrop) return;
    sidebar.classList.toggle('mobile-open', !!isOpen);
    backdrop.classList.toggle('hidden', !isOpen);
    document.body.style.overflow = isOpen ? 'hidden' : '';
}

function openMobileSidebar() {
    syncMobileSidebar(true);
}

function closeMobileSidebar() {
    syncMobileSidebar(false);
}

function toggleMobileSidebar() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    syncMobileSidebar(!sidebar.classList.contains('mobile-open'));
}

function openDetailDrawer(originalIndex) {
    currentDrawerMode = 'udct';
    currentEditRowIndex = originalIndex;
    currentHangHoanEditIndex = -1;
    const item = udctData[originalIndex];
    if (!item) return;
    const isKinhDoanh = currentUser && currentUser.role === 'kinhdoanh';
    suppressUDCTAutoSave = true;
    document.getElementById('drawerRowId').textContent = `Row ID: ${item.rowIndex}`;
    document.getElementById('drawerTenSP').textContent = item.ten_sp || 'N/A';
    document.getElementById('editSoLuong').value = item.so_luong || '';
    document.getElementById('editDonGia').value = item.don_gia_1 || '';
    document.getElementById('editIdSP').value = item.id_sp || '';
    document.getElementById('editIdSPCT').value = item.id_sp_ct || '';
    document.getElementById('editTinhTrang').value = item.tinh_trang || 'Chờ xác nhận';
    document.getElementById('editTrangThai').value = item.trang_thai || '';
    renderEditTrangThaiButtons(item.trang_thai || '');
    document.getElementById('editGhiChu').value = item.ghi_chu || '';

    ['editSoLuong', 'editDonGia', 'editIdSP', 'editIdSPCT'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.disabled = isKinhDoanh;
    });
    const ghiChuEl = document.getElementById('editGhiChu');
    if (ghiChuEl) ghiChuEl.disabled = false; // KINHDOANH luôn được sửa ghi chú

    const ttButtons = document.getElementById('editTrangThaiButtons');
    if (ttButtons) ttButtons.style.pointerEvents = isKinhDoanh ? 'none' : '';

    const footer = document.querySelector('#detailDrawer .pt-4.border-t.border-slate-200.flex.gap-3');
    if (footer) footer.style.display = isKinhDoanh ? 'none' : '';

    const kdFooter = document.getElementById('kdGhiChuFooter');
    if (kdFooter) kdFooter.classList.toggle('hidden', !isKinhDoanh);

    handleIdSPChange();
    handleIdSPCTChange();
    if (!document.getElementById('editIdSPCT').value) {
        document.getElementById('drawerTenSP').textContent = item.ten_sp || 'N/A';
    }
    suppressUDCTAutoSave = false;

    document.getElementById('detailDrawerOverlay').classList.remove('hidden');
    document.getElementById('detailDrawer').classList.add('open');
}

async function saveKDGhiChu() {
    if (currentEditRowIndex === -1) return;
    const item = udctData[currentEditRowIndex];
    if (!item) return;
    const newGhiChu = document.getElementById('editGhiChu').value;
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');
    try {
        const token = await getAccessToken();
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
        const resp = await fetch(url, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({
                valueInputOption: 'USER_ENTERED',
                data: [{ range: `${CONFIG.udctSheetName}!AA${item.rowIndex}`, values: [[newGhiChu]] }]
            })
        });
        if (resp.ok) {
            item.ghi_chu = newGhiChu;
            renderUDCTTable();
            closeDetailDrawer();
        } else {
            alert('Lỗi khi lưu ghi chú.');
        }
    } catch (err) {
        alert('Lỗi: ' + err.message);
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

function closeDetailDrawer() {
    document.getElementById('detailDrawerOverlay').classList.add('hidden');
    document.getElementById('detailDrawer').classList.remove('open');
    currentEditRowIndex = -1;
    currentDrawerMode = 'udct';
    clearTimeout(udctAutoSaveTimer);
}

async function saveRowDetail(showLoading = false) {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (showLoading) loadingOverlay.classList.remove('hidden');

    try {
        const token = await getAccessToken();
        if (currentDrawerMode === 'hh') {
            if (currentHangHoanEditIndex === -1) return;
            const item = hangHoanData[currentHangHoanEditIndex];
            if (!item) return;

            const newData = {
                mvd: document.getElementById('editIdSP')?.value || '',
                ma_gian: document.getElementById('editIdSPCT')?.value || '',
                sku: document.getElementById('editSoLuong')?.value || '',
                sku_ct: document.getElementById('editDonGia')?.value || '',
                slg: document.getElementById('editTinhTrang')?.value || '',
                tinh_trang: document.getElementById('editTrangThai')?.value || '',
                kho: document.getElementById('editGhiChu')?.value || ''
            };

            const batchUpdates = [
                { range: `${CONFIG.hhbhSheetName}!C${item.rowIndex}`, values: [[newData.mvd]] },
                { range: `${CONFIG.hhbhSheetName}!D${item.rowIndex}`, values: [[newData.ma_gian]] },
                { range: `${CONFIG.hhbhSheetName}!H${item.rowIndex}`, values: [[newData.sku]] },
                { range: `${CONFIG.hhbhSheetName}!I${item.rowIndex}`, values: [[newData.sku_ct]] },
                { range: `${CONFIG.hhbhSheetName}!J${item.rowIndex}`, values: [[newData.slg]] },
                { range: `${CONFIG.hhbhSheetName}!N${item.rowIndex}`, values: [[newData.tinh_trang]] },
                { range: `${CONFIG.hhbhSheetName}!T${item.rowIndex}`, values: [[newData.kho]] }
            ];

            const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
            const resp = await fetch(url, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data: batchUpdates })
            });

            if (resp.ok) {
                Object.assign(item, {
                    mvd: newData.mvd,
                    ma_gian: newData.ma_gian,
                    sku: newData.sku,
                    sku_ct: newData.sku_ct,
                    slg: newData.slg,
                    tinh_trang: newData.tinh_trang,
                    kho: newData.kho
                });
                filterHangHoanData();
                closeDetailDrawer();
                alert('Cập nhật Hàng hoàn thành công!');
            } else {
                console.error('Save error:', await resp.text());
                alert('Lỗi khi lưu dữ liệu Hàng hoàn vào Google Sheet.');
            }
            return;
        }

        if (currentEditRowIndex === -1) return;
        const item = udctData[currentEditRowIndex];
        const newData = {
            so_luong: document.getElementById('editSoLuong').value,
            don_gia_1: document.getElementById('editDonGia').value,
            id_sp: document.getElementById('editIdSP').value,
            id_sp_ct: document.getElementById('editIdSPCT').value,
            tinh_trang: document.getElementById('editTinhTrang').value,
            trang_thai: document.getElementById('editTrangThai').value,
            ghi_chu: document.getElementById('editGhiChu').value
        };
        const nextSlgXuat = (newData.trang_thai || '').toLowerCase().includes('hủy') ? 0 : (newData.so_luong || 0);
        // so_luong=O, id_sp=P, id_sp_ct=Q, tinh_trang=X, trang_thai=Y, slg_xuat=S, ghi_chu=AA, don_gia=AE
        const batchUpdates = [
            { range: `${CONFIG.udctSheetName}!O${item.rowIndex}`, values: [[newData.so_luong]] },
            { range: `${CONFIG.udctSheetName}!P${item.rowIndex}`, values: [[newData.id_sp]] },
            { range: `${CONFIG.udctSheetName}!Q${item.rowIndex}`, values: [[newData.id_sp_ct]] },
            { range: `${CONFIG.udctSheetName}!S${item.rowIndex}`, values: [[nextSlgXuat]] },
            { range: `${CONFIG.udctSheetName}!X${item.rowIndex}`, values: [[newData.tinh_trang]] },
            { range: `${CONFIG.udctSheetName}!Y${item.rowIndex}`, values: [[newData.trang_thai]] },
            { range: `${CONFIG.udctSheetName}!AA${item.rowIndex}`, values: [[newData.ghi_chu]] },
            { range: `${CONFIG.udctSheetName}!AE${item.rowIndex}`, values: [[newData.don_gia_1]] }
        ];
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
        const resp = await fetch(url, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data: batchUpdates })
        });
        if (resp.ok) {
            Object.assign(item, newData);
            item.slg_xuat = nextSlgXuat;
            renderUDCTTable();
        } else {
            console.error('Save error:', await resp.text());
            alert('Lỗi khi lưu dữ liệu vào Google Sheet.');
        }
    } catch (err) {
        console.error('Save error:', err);
        alert('Đã xảy ra lỗi khi lưu.');
    } finally {
        if (showLoading) loadingOverlay.classList.add('hidden');
    }
}

function handleIdSPChange() {
    const idSpVal = document.getElementById('editIdSP').value.trim();
    const idSpCtList = document.getElementById('idSpCtList');
    const buttonContainer = document.getElementById('editIdSPCTButtons');
    if (!idSpCtList) return;

    let filteredMã = sanphamData.map(sp => sp.sku_con || '');
    if (idSpVal) {
        const searchVal = idSpVal.toLowerCase();
        filteredMã = filteredMã.filter(ma => (ma || '').toString().toLowerCase().startsWith(searchVal));
    }
    const uniqueMã = [...new Set(filteredMã)].filter(Boolean);

    // Update Datalist
    idSpCtList.innerHTML = uniqueMã.map(ma => `<option value="${ma}">`).join('');

    // Update Buttons
    if (buttonContainer) {
        if (idSpVal && uniqueMã.length > 0) {
            buttonContainer.innerHTML = uniqueMã.slice(0, 15).map(ma => `
                <button onclick="selectIdSPCTSuggestion('${ma}')" 
                        class="px-2 py-1 bg-blue-50 text-[10px] font-bold text-blue-600 rounded border border-blue-100 hover:bg-blue-100 transition-all">
                    ${ma}
                </button>
            `).join('');
        } else {
            buttonContainer.innerHTML = '';
        }
    }

    // Clear current ID SP CT if it doesn't match the new prefix
    const currentCt = document.getElementById('editIdSPCT').value.trim();
    if (currentCt && idSpVal && !currentCt.startsWith(idSpVal)) {
        document.getElementById('editIdSPCT').value = '';
        document.getElementById('drawerTenSP').textContent = '';
    }
}

function selectIdSPCTSuggestion(val) {
    const input = document.getElementById('editIdSPCT');
    if (input) {
        input.value = val;
        handleIdSPCTChange();
        if (typeof scheduleUDCTAutoSave === 'function') scheduleUDCTAutoSave();
    }
}

function handleIdSPCTChange() {
    const currentCt = document.getElementById('editIdSPCT').value.trim().toLowerCase();
    if (!currentCt) return;

    const sp = sanphamData.find(item => (item.sku_con || '').toString().toLowerCase() === currentCt);
    if (sp) {
        document.getElementById('drawerTenSP').textContent = sp.ten_sp || 'N/A';
    }
}

function populateSPLists() {
    const idSpList = document.getElementById('idSpList');
    if (idSpList && sanphamData) {
        const uniqueIdSPs = [...new Set(sanphamData.map(sp => (sp.sku_con || '').substring(0, 4)))].filter(Boolean);
        idSpList.innerHTML = uniqueIdSPs.map(id => `<option value="${id}">`).join('');
    }
}

function generateRandomString(length) {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    for (let i = 0; i < length; i++) {
        result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
}

function generateId(parts) {
    return parts.join(' | ');
}

let tokenRequestPromise = null;
async function getAccessToken() {
    if (accessToken && Date.now() < tokenExpiry - 300000) return accessToken;
    if (tokenRequestPromise) return tokenRequestPromise;

    tokenRequestPromise = (async () => {
        try {
            const header = { alg: "RS256", typ: "JWT" };
            const now = Math.floor(Date.now() / 1000);
            const payload = {
                iss: CONFIG.serviceAccountEmail,
                scope: CONFIG.scopes.join(" "),
                aud: CONFIG.tokenUrl,
                exp: now + 3600,
                iat: now
            };
            const sJWT = KJUR.jws.JWS.sign("RS256", JSON.stringify(header), JSON.stringify(payload), CONFIG.privateKey);

            try {
                const response = await fetch(CONFIG.tokenUrl, {
                    method: "POST",
                    headers: { "Content-Type": "application/x-www-form-urlencoded" },
                    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${sJWT}`
                });
                const data = await response.json();
                if (data.error) {
                    console.error("Token error:", data);
                    throw new Error(data.error_description || "Failed to get token");
                }
                accessToken = data.access_token;
                tokenExpiry = Date.now() + (data.expires_in * 1000);
                return accessToken;
            } catch (err) {
                console.error("Token fetch error:", err);
                alert("Không thể xác thực với Google API: " + err.message);
            }
        } finally {
            tokenRequestPromise = null;
        }
    })();
    return tokenRequestPromise;
}

async function fetchSheetData(sheetName) {
    try {
        const token = await getAccessToken();
        if (!token) return [];
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!A1:AF10000`;
        const resp = await fetch(url, { headers: { "Authorization": `Bearer ${token}` } });
        if (!resp.ok) {
            console.error(`Fetch ${sheetName} failed:`, resp.status);
            return [];
        }
        const data = await resp.json();
        return data.values || [];
    } catch (err) {
        console.error(`Fetch ${sheetName} error:`, err);
        return [];
    }
}

async function fetchAuthData() {
    const data = await fetchSheetData(CONFIG.authSheetName);
    if (!data || data.length <= 1) return [];

    const headers = data[0].map(h => (h || '').toString().trim().toLowerCase());
    const getIndex = (name) => headers.indexOf(name.toLowerCase());
    const idxId = getIndex('id');
    const idxHoten = getIndex('hoten');
    const idxQuyen = getIndex('quyen');
    const idxMatKhau = getIndex('mat_khau');
    const idxTinhTrang = getIndex('tinhtrang');

    usersData = data.slice(1).map(row => ({
        id: (idxId !== -1 ? row[idxId] : row[0] || '').toString().trim(),
        name: (idxHoten !== -1 ? row[idxHoten] : row[1] || '').toString().trim(),
        role: (idxQuyen !== -1 ? row[idxQuyen] : '').toString().trim().toLowerCase() || 'user',
        password: (idxMatKhau !== -1 ? row[idxMatKhau] : '').toString(),
        tinhTrang: (idxTinhTrang !== -1 ? row[idxTinhTrang] : '').toString().trim(),
        raw: row
    })).filter(user => user.id && user.password);

    return usersData;
}

async function clearSheetData(sheetName) {
    try {
        const token = await getAccessToken();
        if (!token) return false;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!A2:AF10000:clear`;
        const resp = await fetch(url, {
            method: "POST",
            headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" }
        });
        if (!resp.ok) {
            const errorText = await resp.text();
            console.error("Lỗi khi xóa dữ liệu:", errorText);
            return false;
        }
        console.log(`Đã xóa dữ liệu cũ trong ${sheetName}`);
        return true;
    } catch (err) {
        console.error("Lỗi clearSheetData:", err);
        return false;
    }
}

async function appendSheetData(sheetName, values) {
    try {
        const token = await getAccessToken();
        if (!token) return false;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!A:A:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
        const resp = await fetch(url, {
            method: "POST",
            headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ values: values, majorDimension: "ROWS" })
        });
        if (!resp.ok) {
            const errorText = await resp.text();
            console.error("Lỗi khi ghi dữ liệu:", errorText);
            return false;
        }
        const result = await resp.json();
        console.log("Ghi dữ liệu thành công:", result);
        return true;
    } catch (err) {
        console.error("Lỗi appendSheetData:", err);
        return false;
    }
}

async function updateSheetCell(sheetName, rowIndex, colIndex, value) {
    try {
        const token = await getAccessToken();
        if (!token) return false;

        // Convert colIndex (1-based) to Letter (A=1, B=2...)
        let colLetter = "";
        let temp = colIndex;
        while (temp > 0) {
            let mod = (temp - 1) % 26;
            colLetter = String.fromCharCode(65 + mod) + colLetter;
            temp = Math.floor((temp - mod) / 26);
        }

        const range = `${sheetName}!${colLetter}${rowIndex}`;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${range}?valueInputOption=USER_ENTERED`;

        const resp = await fetch(url, {
            method: "PUT",
            headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ values: [[value]] })
        });

        return resp.ok;
    } catch (err) {
        console.error("Lỗi updateSheetCell:", err);
        return false;
    }
}

async function updateSheetValue(sheetName, range, value) {
    try {
        const token = await getAccessToken();
        if (!token) return false;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!${range}?valueInputOption=USER_ENTERED`;
        const resp = await fetch(url, {
            method: "PUT",
            headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ values: [[value]] })
        });
        return resp.ok;
    } catch (err) {
        console.error("Lỗi updateSheetValue:", err);
        return false;
    }
}

// fetchAuthData removed as login is disabled


function generateSkeletonRows(columnsCount, rowsCount = 10) {
    let html = '';
    for (let i = 0; i < rowsCount; i++) {
        let cols = '';
        for (let j = 0; j < columnsCount; j++) {
            cols += `<td class="px-4 py-4"><div class="h-4 bg-slate-200 animate-pulse rounded w-3/4"></div></td>`;
        }
        html += `<tr class="border-b border-slate-100 bg-white">${cols}</tr>`;
    }
    return html;
}

function saveFiltersToCache() {
    const filters = {
        filterUDCTFrom: document.getElementById('filterUDCTFrom')?.value || '',
        filterUDCTTo: document.getElementById('filterUDCTTo')?.value || '',
        filterUDCTSan: document.getElementById('filterUDCTSan')?.value || '',
        filterUDCTKhungH: document.getElementById('filterUDCTKhungH')?.value || '',
        filterUDCTTrangThai: document.getElementById('filterUDCTTrangThai')?.value || '',
        filterUDCTMaGian: document.getElementById('filterUDCTMaGian')?.value || '',
        fromDate: document.getElementById('fromDate')?.value || '',
        toDate: document.getElementById('toDate')?.value || '',
        filterMaGian: document.getElementById('filterMaGian')?.value || ''
    };
    localStorage.setItem('erp_filters', JSON.stringify(filters));
}

function loadFiltersFromCache() {
    const filtersStr = localStorage.getItem('erp_filters');
    if (filtersStr) {
        try {
            const filters = JSON.parse(filtersStr);
            Object.keys(filters).forEach(id => {
                const el = document.getElementById(id);
                if (el) el.value = filters[id];
            });
        } catch (e) { }
    } else {
        const today = new Date();
        const d = String(today.getDate()).padStart(2, '0');
        const m = String(today.getMonth() + 1).padStart(2, '0');
        const y = today.getFullYear();
        const strDate = `${y}-${m}-${d}`;
        if (document.getElementById('fromDate')) document.getElementById('fromDate').value = strDate;
        if (document.getElementById('toDate')) document.getElementById('toDate').value = strDate;
    }
}

async function loadUDCTData() {
    document.getElementById('donhangTableBody').innerHTML = generateSkeletonRows(18, 5);
    try {
        const data = await fetchSheetData(CONFIG.udctSheetName);
        if (data.length <= 1) {
            document.getElementById('donhangTableBody').innerHTML = '<tr><td colspan="18" class="text-center py-8 text-slate-500">Không có dữ liệu</td></tr>';
            return;
        }

        udctData = data.slice(1).map((row, idx) => ({
            rowIndex: idx + 2,
            ngay: row[4] || '',
            san: row[8] || '',
            khung_h: row[9] || '',
            ma_gian: row[10] || '',
            mvd: row[11] || '',
            mdh: row[12] || '',
            sku_shop_up: row[13] || '',
            so_luong: row[14] || '',
            id_sp: row[15] || '',
            id_sp_ct: row[16] || '',
            ten_sp: row[17] || '',
            slg_xuat: row[18] || '',
            don_gia_1: row[30] || '',
            tinh_trang: row[23] || '',
            trang_thai: row[24] || '',
            ghi_chu: row[26] || '',
            mien: row[7] || ''
        }));

        populateUDCTFilters();

        // Helper to parse date for sorting
        const toIsoDate = (str) => {
            if (!str) return '';
            const s = str.split(' ')[0];
            if (s.includes('/')) {
                const [d, m, y] = s.split('/');
                return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
            }
            return s;
        };

        // Sắp xếp Đơn chi tiết theo ngày giảm dần (mới nhất lên đầu)
        udctData.sort((a, b) => {
            const da = toIsoDate(a.ngay);
            const db = toIsoDate(b.ngay);
            return db.localeCompare(da);
        });

        setUDCTQuickDate('today'); // Sets default date to today and calls filterUDCTTable
        buildUpmisaData();
    } catch (error) {
        console.error("Load UDCT error:", error);
    }
}

function normalizeSanLabel(v) {
    return (v || '')
        .toString()
        .trim()
        .replace(/\d+$/, '')
        .trim();
}



function populateUDCTFilters() {
    const sans = [...new Set(udctData.map(i => normalizeSanLabel(i.san)))].filter(Boolean).sort();
    const khungHs = [...new Set(udctData.map(i => i.khung_h))].filter(Boolean).sort((a, b) => {
        const numA = parseInt(a) || 0;
        const numB = parseInt(b) || 0;
        return numA - numB;
    });
    const maGians = [...new Set(udctData.map(i => i.ma_gian))].filter(Boolean).sort();
    const idSps = [...new Set(udctData.map(i => i.id_sp))].filter(Boolean).sort();
    const idSpCts = [...new Set(udctData.map(i => i.id_sp_ct))].filter(Boolean).sort();

    const khSelection = document.getElementById('filterUDCTKhungH')?.value || '';
    const mgSelection = document.getElementById('filterUDCTMaGian')?.value || '';
    const sanSelection = document.getElementById('filterUDCTSan')?.value || '';
    const ttSelection = document.getElementById('filterUDCTTrangThai')?.value || '';
    const selectedStatuses = new Set((ttSelection || '').split('||').map(normalizeTrangThai).filter(Boolean));

    const commonFiltered = getUDCTBaseFilteredForStatusCounts();

    const fillSelect = (id, list, placeholder) => {
        const select = document.getElementById(id);
        if (!select) return;
        const currentVal = select.value;
        select.innerHTML = `<option value="">${placeholder}</option>` +
            list.map(v => `<option value="${v}">${v}</option>`).join('');
        select.value = currentVal;
    };

    const renderBtns = (id, options, currentVal, countField) => {
        const container = document.getElementById(id + 'Buttons');
        if (!container) return;
        const counts = options.reduce((acc, opt) => {
            acc[opt] = commonFiltered.filter(item => {
                if (countField === 'san') {
                    if (normalizeSanLabel(item.san) !== opt) return false;
                    if (khSelection && item.khung_h !== khSelection) return false;
                    if (mgSelection && item.ma_gian !== mgSelection) return false;
                    if (selectedStatuses.size > 0 && !selectedStatuses.has(normalizeTrangThai(item.trang_thai))) return false;
                } else if (countField === 'khung_h') {
                    if (item.khung_h !== opt) return false;
                    if (sanSelection && normalizeSanLabel(item.san) !== sanSelection) return false;
                    if (mgSelection && item.ma_gian !== mgSelection) return false;
                    if (selectedStatuses.size > 0 && !selectedStatuses.has(normalizeTrangThai(item.trang_thai))) return false;
                }
                return true;
            }).length;
            return acc;
        }, {});

        let html = `<button onclick="setUDCTBtnFilter('${id}', '')" class="px-2 py-1.5 text-[11px] rounded-lg font-bold transition-all duration-200 ${currentVal === '' ? 'bg-white text-primary shadow-sm' : 'text-slate-500 hover:text-slate-600'}">Tất cả</button>`;
        options.forEach(opt => {
            const badgeHtml = `<span class="ml-1 inline-flex items-center justify-center min-w-[22px] h-5 px-1.5 text-[10px] font-bold leading-none text-red-600 bg-white border border-red-200 shadow-sm rounded-md">${counts[opt] || 0}</span>`;
            html += `<button onclick="setUDCTBtnFilter('${id}', '${opt}')" data-opt="${opt}" class="px-2 py-1.5 text-[11px] rounded-lg font-bold transition-all duration-200 flex items-center ${currentVal === opt ? 'bg-white text-primary shadow-sm' : 'text-slate-500 hover:text-slate-600'}">${opt} ${badgeHtml}</button>`;
        });
        container.innerHTML = html;
    };

    const renderMultiStatusBtns = (id, options, currentVal) => {
        const container = document.getElementById(id + 'Buttons');
        if (!container) return;

        const counts = options.reduce((acc, opt) => {
            acc[opt] = commonFiltered.filter(item => {
                if (normalizeTrangThai(item.trang_thai) !== normalizeTrangThai(opt)) return false;
                if (sanSelection && normalizeSanLabel(item.san) !== sanSelection) return false;
                if (khSelection && item.khung_h !== khSelection) return false;
                if (mgSelection && item.ma_gian !== mgSelection) return false;
                return true;
            }).length;
            return acc;
        }, {});

        const selected = new Set((currentVal || '').split('||').map(normalizeTrangThai).filter(Boolean));
        let html = `<button onclick="setUDCTStatusMultiFilter('')" class="px-2 py-1.5 text-[11px] rounded-lg font-bold transition-all duration-200 ${selected.size === 0 ? 'bg-white text-primary shadow-sm' : 'text-slate-500 hover:text-slate-600'}">Tất cả</button>`;
        options.forEach(opt => {
            const c = counts[opt] || 0;
            const active = selected.has(normalizeTrangThai(opt));
            const badgeHtml = `<span class="ml-1 inline-flex items-center justify-center min-w-[22px] h-5 px-1.5 text-[10px] font-bold leading-none text-red-600 bg-white border border-red-200 shadow-sm rounded-md">${c}</span>`;
            html += `<button onclick="setUDCTStatusMultiFilter('${opt}')" class="px-2 py-1.5 text-[11px] rounded-lg font-bold transition-all duration-200 flex items-center ${active ? 'bg-white text-primary shadow-sm' : 'text-slate-500 hover:text-slate-600'}">${opt} ${badgeHtml}</button>`;
        });
        container.innerHTML = html;
    };

    const sf = document.getElementById('filterUDCTSan');
    if (sf) renderBtns('filterUDCTSan', sans, sf.value, 'san');

    const khf = document.getElementById('filterUDCTKhungH');
    if (khf) renderBtns('filterUDCTKhungH', khungHs, khf.value, 'khung_h');

    const ttf = document.getElementById('filterUDCTTrangThai');
    if (ttf) renderMultiStatusBtns('filterUDCTTrangThai', udctTrangThaiOptions, ttf.value);

    fillSelect('filterUDCTMaGian', maGians, 'Tất cả Mã gian');

    const mgList = document.getElementById('maGianList');
    if (mgList) {
        mgList.innerHTML = maGians.map(v => `<option value="${v}">`).join('');
    }
    const idSpList = document.getElementById('udctIdSpList');
    if (idSpList) {
        idSpList.innerHTML = idSps.map(v => `<option value="${v}">`).join('');
    }
    const idSpCtList = document.getElementById('udctIdSpCtList');
    if (idSpCtList) {
        idSpCtList.innerHTML = idSpCts.map(v => `<option value="${v}">`).join('');
    }

    // Update Tab Badges
    const counts = {
        duplicate: 0,
        notes: commonFiltered.filter(i => (i.ghi_chu || '').toString().trim()).length,
        noSku: commonFiltered.filter(i => !(i.id_sp_ct || '').toString().trim()).length
    };
    const mvdCounts = {};
    commonFiltered.forEach(i => { if (i.mvd && i.mvd !== '-') mvdCounts[i.mvd] = (mvdCounts[i.mvd] || 0) + 1; });
    counts.duplicate = commonFiltered.filter(i => i.mvd && i.mvd !== '-' && mvdCounts[i.mvd] > 1).length;

    if (document.getElementById('udctDuplicateCountBadge')) document.getElementById('udctDuplicateCountBadge').textContent = counts.duplicate;
    if (document.getElementById('udctNotesCountBadge')) document.getElementById('udctNotesCountBadge').textContent = counts.notes;
    if (document.getElementById('udctNoSkuCountBadge')) document.getElementById('udctNoSkuCountBadge').textContent = counts.noSku;
}

function setUDCTBtnFilter(id, val) {
    const el = document.getElementById(id);
    if (el) el.value = val;
    // Gọi filterUDCTTable sẽ tự động gọi luôn populateUDCTFilters
    filterUDCTTable();
}

function setUDCTStatusMultiFilter(val) {
    const el = document.getElementById('filterUDCTTrangThai');
    if (!el) return;
    const selected = new Set((el.value || '').split('||').map(normalizeTrangThai).filter(Boolean));
    const normalizedVal = normalizeTrangThai(val);
    if (!val) {
        selected.clear();
    } else if (selected.has(normalizedVal)) {
        selected.delete(normalizedVal);
    } else {
        selected.add(normalizedVal);
    }
    el.value = Array.from(selected).join('||');
    // Gọi filterUDCTTable sẽ tự động gọi luôn populateUDCTFilters
    filterUDCTTable();
}

function renderEditTrangThaiButtons(currentVal = '') {
    const container = document.getElementById('editTrangThaiButtons');
    if (!container) return;
    const options = [
        { value: '', label: 'Để trống' },
        { value: '1 THAY THẾ', label: '1 THAY THẾ' },
        { value: '2 HỦY', label: '2 HỦY' },
        { value: '3 HÊT HÀNG', label: '3 HÊT HÀNG' },
        { value: '4 MAI GỌI', label: '4 MAI GỌI' }
    ];
    container.innerHTML = options.map(opt => {
        const active = currentVal === opt.value;
        return `<button type="button" onclick="setEditTrangThai('${opt.value}')" class="px-3 py-1.5 text-xs rounded-lg font-bold border transition-all duration-200 ${active ? 'bg-primary text-white border-primary shadow-sm' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}">${opt.label}</button>`;
    }).join('');
}

function setEditTrangThai(value) {
    const input = document.getElementById('editTrangThai');
    if (input) input.value = value;
    renderEditTrangThaiButtons(value);
    scheduleUDCTAutoSave();
}

function scheduleUDCTAutoSave() {
    if (suppressUDCTAutoSave) return;
    if (currentDrawerMode !== 'udct' || currentEditRowIndex === -1) return;
    clearTimeout(udctAutoSaveTimer);
    udctAutoSaveTimer = setTimeout(() => {
        saveRowDetail(false);
    }, 450);
}

let filteredUDCT = [];
let udctCurrentPage = 1;
const uitItemsPerPage = 500; // tăng lên 500 dòng 
let udctQuickStatusTab = 'all'; // 'all' or 'cancelled' or 'duplicate' or 'notes'
const udctSelectedRows = new Set();
const udctTrangThaiOptions = ['1 THAY THẾ', '2 HỦY', '3 HÊT HÀNG', '4 MAI GỌI'];

function normalizeTrangThai(v) {
    return (v || '').toString().trim().toUpperCase();
}

async function saveUDCTMainInline(rowIndex, field, value) {
    const item = udctData.find(i => Number(i.rowIndex) === Number(rowIndex));
    if (!item) return;

    // Chuẩn hóa giá trị
    const val = value.trim();
    item[field] = val;

    try {
        const token = await getAccessToken();
        const updates = [];

        if (field === 'id_sp_ct') {
            // Cập nhật mã (Cột Q - Index 16)
            updates.push({
                range: `${CONFIG.udctSheetName}!Q${rowIndex}`,
                values: [[val]]
            });

            // Tìm thông tin sản phẩm để cập nhật nốt Tên và Giá
            if (sanphamData && sanphamData.length > 0) {
                const sp = sanphamData.find(s => (s.sku_con || '').toString().toLowerCase() === val.toLowerCase());
                if (sp) {
                    item.ten_sp = sp.ten_sp;
                    item.id_sp = (sp.sku_con || '').substring(0, 4); // Cập nhật cả ID SP (Cột P)
                    item.don_gia_1 = sp.gia_ban;

                    updates.push({ range: `${CONFIG.udctSheetName}!R${rowIndex}`, values: [[item.ten_sp]] });
                    updates.push({ range: `${CONFIG.udctSheetName}!P${rowIndex}`, values: [[item.id_sp]] });
                    updates.push({ range: `${CONFIG.udctSheetName}!AE${rowIndex}`, values: [[item.don_gia_1]] });
                }
            }
        }

        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
        const resp = await fetch(url, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data: updates })
        });

        if (resp.ok) {
            renderUDCTTable();
        } else {
            console.error("Save Main Inline Error:", await resp.text());
            alert('Lỗi khi lưu dữ liệu trực tiếp.');
        }
    } catch (err) {
        console.error("Save Main Inline Exception:", err);
        alert('Có lỗi khi lưu: ' + err.message);
    }
}

function getUDCTBaseFilteredForStatusCounts() {
    const from = document.getElementById('filterUDCTFrom')?.value || '';
    const to = document.getElementById('filterUDCTTo')?.value || '';
    const san = document.getElementById('filterUDCTSan')?.value || '';
    const kh = document.getElementById('filterUDCTKhungH')?.value || '';
    const mg = document.getElementById('filterUDCTMaGian')?.value || '';
    const idSp = document.getElementById('filterUDCTIdSp')?.value || '';
    const idSpCt = document.getElementById('filterUDCTIdSpCt')?.value || '';
    const search = (document.getElementById('filterUDCTSearch')?.value || '').toLowerCase();

    return udctData.filter(item => {
        const rawDate = item.ngay ? item.ngay.split(' ')[0] : '';
        let itemDate = rawDate;
        if (rawDate.includes('/')) {
            const [d, m, y] = rawDate.split('/');
            itemDate = `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
        }
        if (from && itemDate < from) return false;
        if (to && itemDate > to) return false;
        if (san && normalizeSanLabel(item.san) !== san) return false;
        if (kh && item.khung_h !== kh) return false;
        if (mg && item.ma_gian !== mg) return false;
        if (idSp && item.id_sp !== idSp) return false;
        if (idSpCt && item.id_sp_ct !== idSpCt) return false;
        if (udctQuickStatusTab === 'cancelled') {
            if (!(item.trang_thai || '').toLowerCase().includes('hủy') && !(item.tinh_trang || '').toLowerCase().includes('hủy')) {
                return false;
            }
        }
        if (udctQuickStatusTab === 'notes') {
            if (!(item.ghi_chu || '').toString().trim()) return false;
        }
        if (search) {
            const searchTerms = search.split(',').map(s => s.trim()).filter(Boolean);
            const rowText = `${item.mvd} ${item.mdh} ${item.ten_sp} ${item.id_sp_ct}`.toLowerCase();
            if (!searchTerms.some(term => rowText.includes(term))) return false;
        }
        return true;
    });
}

function setUDCTQuickTab(tab) {
    udctQuickStatusTab = tab;
    const tabs = {
        all: document.getElementById('udctTabAll'),
        cancelled: document.getElementById('udctTabCancelled'),
        duplicate: document.getElementById('udctTabDuplicate'),
        notes: document.getElementById('udctTabNotes'),
        noSku: document.getElementById('udctTabNoSku')
    };

    Object.keys(tabs).forEach(k => {
        if (!tabs[k]) return;
        if (k === tab) {
            tabs[k].className = "px-4 py-2 text-sm font-semibold rounded-t-lg border-b-2 border-primary text-primary transition-colors";
        } else {
            tabs[k].className = "px-4 py-2 text-sm font-semibold rounded-t-lg border-b-2 border-transparent text-slate-500 hover:text-slate-700 transition-colors";
        }
    });
    filterUDCTTable();
}

function setUDCTQuickDate(type) {
    const fromInput = document.getElementById('filterUDCTFrom');
    const toInput = document.getElementById('filterUDCTTo');
    const today = new Date();
    let fromDate, toDate;

    const format = (d) => {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    };

    if (type === 'today') {
        fromDate = format(today);
        toDate = format(today);
    } else if (type === 'thisWeek') {
        const day = today.getDay();
        const diff = today.getDate() - day + (day === 0 ? -6 : 1);
        const firstDay = new Date(today.setDate(diff));
        fromDate = format(firstDay);

        const lastDay = new Date(firstDay);
        lastDay.setDate(firstDay.getDate() + 6);
        toDate = format(lastDay);
    } else if (type === 'thisMonth') {
        const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
        const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);
        fromDate = format(firstDay);
        toDate = format(lastDay);
    }

    fromInput.value = fromDate;
    toInput.value = toDate;
    filterUDCTTable();
}

function isUDCTTrung(item, mvdMap = {}, mdhMap = {}) {
    const mvdDup = item.mvd && mvdMap[item.mvd]?.size > 1;
    const mdhDup = item.mdh && mdhMap[item.mdh]?.size > 1;
    return Boolean(mvdDup || mdhDup);
}

function filterUDCTTable() {
    const from = document.getElementById('filterUDCTFrom').value;
    const to = document.getElementById('filterUDCTTo').value;
    const san = document.getElementById('filterUDCTSan').value;
    const kh = document.getElementById('filterUDCTKhungH').value;
    const mg = document.getElementById('filterUDCTMaGian').value;
    const idSp = document.getElementById('filterUDCTIdSp')?.value || '';
    const idSpCt = document.getElementById('filterUDCTIdSpCt')?.value || '';
    const tt = document.getElementById('filterUDCTTrangThai').value;
    const selectedStatuses = new Set((tt || '').split('||').map(normalizeTrangThai).filter(Boolean));
    const searchInput = document.getElementById('filterUDCTSearch');
    const search = searchInput ? searchInput.value.toLowerCase() : '';

    const base = udctData.filter(item => {
        const rawDate = item.ngay ? item.ngay.split(' ')[0] : '';
        let itemDate = rawDate;
        if (rawDate.includes('/')) {
            const [d, m, y] = rawDate.split('/');
            itemDate = `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
        }
        if (from && itemDate < from) return false;
        if (to && itemDate > to) return false;
        if (san && normalizeSanLabel(item.san) !== san) return false;
        if (kh && item.khung_h !== kh) return false;
        if (mg && item.ma_gian !== mg) return false;
        if (idSp && item.id_sp !== idSp) return false;
        if (idSpCt && item.id_sp_ct !== idSpCt) return false;
        if (selectedStatuses.size > 0 && !selectedStatuses.has(normalizeTrangThai(item.trang_thai))) return false;

        if (search) {
            const searchTerms = search.split(',').map(s => s.trim()).filter(Boolean);
            const rowText = `${item.mvd} ${item.mdh} ${item.ten_sp} ${item.id_sp_ct}`.toLowerCase();
            const isMatch = searchTerms.some(term => rowText.includes(term));
            if (!isMatch) return false;
        }
        return true;
    });

    filteredUDCT = udctQuickStatusTab === 'cancelled'
        ? base.filter(item => (item.trang_thai || '').toLowerCase().includes('hủy') || (item.tinh_trang || '').toLowerCase().includes('hủy'))
        : udctQuickStatusTab === 'notes'
            ? base.filter(item => (item.ghi_chu || '').toString().trim())
            : udctQuickStatusTab === 'noSku'
                ? base.filter(item => !(item.id_sp_ct || '').toString().trim())
                : udctQuickStatusTab === 'duplicate'
                    ? (() => {
                        const mvdCounts = {};
                        base.forEach(i => { if (i.mvd && i.mvd !== '-') mvdCounts[i.mvd] = (mvdCounts[i.mvd] || 0) + 1; });
                        return base.filter(i => i.mvd && i.mvd !== '-' && mvdCounts[i.mvd] > 1);
                    })()
                    : base;

    // Cập nhật badge Tổng dòng và Unique MVD
    const totalRowsBadge = document.getElementById('udctTotalRowsBadge');
    const uniqueMVDBadge = document.getElementById('udctUniqueMVDBadge');
    if (totalRowsBadge) totalRowsBadge.textContent = filteredUDCT.length.toLocaleString('vi-VN');
    if (uniqueMVDBadge) {
        const uniqueMVDs = new Set(filteredUDCT.map(i => i.mvd).filter(v => v && v !== '-'));
        uniqueMVDBadge.textContent = uniqueMVDs.size.toLocaleString('vi-VN');
    }

    saveFiltersToCache();
    udctCurrentPage = 1;

    // Cập nhật lại logic đếm số trên các thẻ button (Sàn, Khung H, ...)
    populateUDCTFilters();

    renderUDCTTable();
}

function changeUDCTDate(id, delta) {
    const input = document.getElementById(id);
    if (!input.value) return;
    const d = new Date(input.value);
    d.setDate(d.getDate() + delta);
    input.value = d.toISOString().split('T')[0];
    filterUDCTTable();
}

function changeReportDate(id, delta) {
    const input = document.getElementById(id);
    if (!input.value) return;
    const d = new Date(input.value);
    d.setDate(d.getDate() + delta);
    input.value = d.toISOString().split('T')[0];
    autoFilterReport();
}

function changeUDCTPage(step) {
    const totalPages = Math.ceil(filteredUDCT.length / uitItemsPerPage);
    udctCurrentPage += step;
    if (udctCurrentPage < 1) udctCurrentPage = 1;
    if (udctCurrentPage > totalPages) udctCurrentPage = totalPages;
    renderUDCTTable();
}

function renderUDCTTable() {
    const tbody = document.getElementById('donhangTableBody');
    if (!tbody) return;

    const baseFilteredForCounts = getUDCTBaseFilteredForStatusCounts();

    const mvdMap = {};
    const mdhMap = {};
    udctData.forEach(d => {
        const datePart = (d.ngay || '').split(' ')[0];
        const timeKey = `${datePart}|${d.khung_h}`;
        if (d.mvd && d.mvd !== '-' && d.mvd !== '') {
            if (!mvdMap[d.mvd]) mvdMap[d.mvd] = new Set();
            mvdMap[d.mvd].add(timeKey);
        }
        if (d.mdh && d.mdh !== '-' && d.mdh !== '') {
            if (!mdhMap[d.mdh]) mdhMap[d.mdh] = new Set();
            mdhMap[d.mdh].add(timeKey);
        }
    });
    const duplicateList = baseFilteredForCounts.filter(item => isUDCTTrung(item, mvdMap, mdhMap));
    window.__udctDuplicateCount = duplicateList.length;
    if (udctQuickStatusTab === 'duplicate') {
        filteredUDCT = duplicateList;
    }
    const duplicateBadge = document.getElementById('udctDuplicateCountBadge');
    if (duplicateBadge) duplicateBadge.textContent = duplicateList.length.toLocaleString('vi-VN');
    const notesCount = baseFilteredForCounts.filter(item => (item.ghi_chu || '').toString().trim()).length;
    const notesBadge = document.getElementById('udctNotesCountBadge');
    if (notesBadge) notesBadge.textContent = notesCount.toLocaleString('vi-VN');

    if (filteredUDCT.length === 0) {
        tbody.innerHTML = '<tr><td colspan="18" class="text-center py-8 text-slate-500">Không tìm thấy dữ liệu phù hợp</td></tr>';
        document.getElementById('udctPageInfo').innerHTML = `Đang hiển thị <span class="font-medium text-slate-900">0-0</span> trong số <span class="font-medium text-slate-900">0</span> đơn hàng`;
        document.getElementById('udctPrevPage').disabled = true;
        document.getElementById('udctNextPage').disabled = true;
        const selectAll = document.getElementById('udctSelectAll');
        if (selectAll) selectAll.checked = false;
        return;
    }

    const totalItems = filteredUDCT.length;
    const totalPages = Math.ceil(totalItems / uitItemsPerPage);
    if (udctCurrentPage > totalPages) udctCurrentPage = totalPages;

    const startIndex = (udctCurrentPage - 1) * uitItemsPerPage;
    const endIndex = Math.min(startIndex + uitItemsPerPage, totalItems);
    const pageData = filteredUDCT.slice(startIndex, endIndex);

    document.getElementById('udctPageInfo').innerHTML = `Đang hiển thị <span class="font-medium text-slate-900">${startIndex + 1}-${endIndex}</span> trong số <span class="font-medium text-slate-900">${totalItems}</span> đơn hàng`;
    document.getElementById('udctPrevPage').disabled = udctCurrentPage === 1;
    document.getElementById('udctNextPage').disabled = udctCurrentPage === totalPages;

    tbody.innerHTML = pageData.map((item) => `
                <tr ondblclick="openDetailDrawer(${udctData.indexOf(item)})" class="border-b border-slate-100 hover:bg-slate-50 cursor-pointer group">

                    <td class="px-3 py-2 text-[13px] text-slate-900">${item.ngay || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-900">${item.san || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-900">${item.khung_h || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-900 font-medium">${item.ma_gian || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-900">${item.mvd || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-900 font-medium text-blue-600">${item.mdh || '-'}</td>
                    <td class="px-3 py-2 text-[13px]">
                        ${((item.mvd && mvdMap[item.mvd]?.size > 1) || (item.mdh && mdhMap[item.mdh]?.size > 1))
            ? '<span class="text-rose-600 font-bold bg-rose-50 px-2 py-0.5 rounded border border-rose-200">TRÙNG</span>'
            : '<span class="text-slate-400">-</span>'}
                    </td>
                    <td class="px-3 py-2 text-[13px] text-slate-500 font-mono">${item.sku_shop_up || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-900">${item.so_luong || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-500">${item.id_sp || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-500 hover:bg-white transition-colors" onclick="event.stopPropagation()">
                        <input type="text" value="${item.id_sp_ct || ''}" 
                               onchange="saveUDCTMainInline('${item.rowIndex}', 'id_sp_ct', this.value)"
                               class="w-full bg-transparent border-none focus:ring-1 focus:ring-blue-400 rounded px-1 outline-none font-bold text-blue-600 h-8">
                        ${(() => {
            const idSp = item.id_sp || '';
            const isMissing = !item.id_sp_ct || item.id_sp_ct === '-' || item.id_sp_ct === '';
            if (!isMissing || !idSp || !sanphamData) return '';
            const matches = sanphamData.filter(s =>
                (s.sku_con || '').toString().startsWith(idSp) &&
                (s.sku_con || '').toString().length > 5
            ).slice(0, 3);
            return ''; // Bỏ gợi ý
        })()}
                    </td>
                    <td class="px-3 py-2 text-[13px] text-slate-900 max-w-[6rem] truncate">${item.ten_sp || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-900 font-semibold">${item.slg_xuat || '-'}</td>
                    <td class="px-3 py-2 text-[13px] text-slate-900">${parseFloat(item.don_gia_1 || 0).toLocaleString('vi-VN')}</td>
                    <td class="px-3 py-2 text-[13px]">
                        <span class="px-2 py-0.5 rounded-full text-[11px] font-medium ${item.trang_thai === '1 THAY THẾ' ? 'bg-green-100 text-green-700' : (item.trang_thai === '2 HỦY' ? 'bg-red-100 text-red-700' : (item.trang_thai === '3 HÊT HÀNG' ? 'bg-amber-100 text-amber-700' : (item.trang_thai === '4 MAI GỌI' ? 'bg-blue-100 text-blue-700' : 'bg-slate-100 text-slate-700')))}">
                            ${item.trang_thai || '-'}
                        </span>
                    </td>
                    <td class="px-3 py-2 text-[13px] text-slate-500 italic truncate max-w-[150px]">${item.ghi_chu || '-'}</td>
                    ${currentUser && currentUser.role === 'kinhdoanh' ? '<td class="px-3 py-2 text-[13px] text-slate-400">-</td>' : `
                    <td class="px-3 py-2 text-[13px] w-44">
                        <button onclick="event.stopPropagation(); updateUDCTPrice(${udctData.indexOf(item)})" class="p-1 px-2 text-blue-600 hover:bg-blue-50 rounded-lg border border-blue-100" title="Up đơn giá">
                            <span class="text-[10px]">Up giá</span>
                        </button>
                        <button onclick="event.stopPropagation(); quickCancelUDCT(${udctData.indexOf(item)})" class="p-1 px-2 ml-1 text-red-600 hover:bg-red-50 rounded-lg border border-red-100" title="Hủy nhanh">
                            <span class="text-[10px]">Hủy nhanh</span>
                        </button>
                        <button onclick="event.stopPropagation(); quickSetUDCTStatus(${udctData.indexOf(item)}, '3 HÊT HÀNG')" class="p-1 px-2 ml-1 text-amber-600 hover:bg-amber-50 rounded-lg border border-amber-100" title="3 HÊT HÀNG">
                            <span class="text-[10px]">3 HÊT HÀNG</span>
                        </button>
                        <button onclick="event.stopPropagation(); quickSetUDCTStatus(${udctData.indexOf(item)}, '4 MAI GỌI')" class="p-1 px-2 ml-1 text-sky-600 hover:bg-sky-50 rounded-lg border border-sky-100" title="4 MAI GỌI">
                            <span class="text-[10px]">4 MAI GỌI</span>
                        </button>
                    </td>`}
                </tr>
            `).join('');

}

function copyUniqueMVD() {
    if (!filteredUDCT || filteredUDCT.length === 0) {
        alert("Không có dữ liệu để copy!");
        return;
    }

    // Chỉ copy trong trang hiện tại (lọc theo trang)
    const startIndex = (udctCurrentPage - 1) * uitItemsPerPage;
    const endIndex = Math.min(startIndex + uitItemsPerPage, filteredUDCT.length);
    const pageData = filteredUDCT.slice(startIndex, endIndex);

    const mvds = pageData
        .map(item => (item.mvd || '').trim())
        .filter(mvd => mvd && mvd !== '-');
    const uniqueMVDs = [...new Set(mvds)];

    if (uniqueMVDs.length === 0) {
        alert("Không tìm thấy MVD hợp lệ!");
        return;
    }

    const textToCopy = uniqueMVDs.join('\n');
    const textArea = document.createElement("textarea");
    textArea.value = textToCopy;
    textArea.style.position = "fixed";
    textArea.style.left = "-9999px";
    textArea.style.top = "0";
    document.body.appendChild(textArea);
    textArea.select();
    try {
        const successful = document.execCommand('copy');
        if (successful) {
            alert(`Đã copy ${uniqueMVDs.length} MVD duy nhất!`);
        }
    } catch (err) {
        console.error('Lỗi khi copy:', err);
        alert('Lỗi khi sao chép!');
    }
    document.body.removeChild(textArea);
}

function toggleUDCTRowSelection(rowIndex, checked) {
    if (checked) udctSelectedRows.add(rowIndex);
    else udctSelectedRows.delete(rowIndex);
}

function toggleSelectAllUDCT(source) {
    const startIndex = (udctCurrentPage - 1) * uitItemsPerPage;
    const endIndex = Math.min(startIndex + uitItemsPerPage, filteredUDCT.length);
    const pageData = filteredUDCT.slice(startIndex, endIndex);
    pageData.forEach(item => {
        if (source.checked) udctSelectedRows.add(item.rowIndex);
        else udctSelectedRows.delete(item.rowIndex);
    });
    renderUDCTTable();
}

async function quickCancelUDCT(originalIndex) {
    const item = udctData[originalIndex];
    if (!item) return;
    const success = await updateSheetValue(CONFIG.udctSheetName, `Y${item.rowIndex}`, '2 HỦY');
    if (!success) return alert(`Lỗi khi cập nhật trạng thái hủy cho dòng ${item.rowIndex}`);
    item.trang_thai = '2 HỦY';
    item.slg_xuat = 0;
    await updateSheetValue(CONFIG.udctSheetName, `S${item.rowIndex}`, 0);
    renderUDCTTable();
}

async function quickSetUDCTStatus(originalIndex, statusValue) {
    const item = udctData[originalIndex];
    if (!item) return;
    const success = await updateSheetValue(CONFIG.udctSheetName, `Y${item.rowIndex}`, statusValue);
    if (!success) return alert(`Lỗi khi cập nhật trạng thái cho dòng ${item.rowIndex}`);
    item.trang_thai = statusValue;
    if ((statusValue || '').toString().toLowerCase().includes('hủy')) {
        item.slg_xuat = 0;
        await updateSheetValue(CONFIG.udctSheetName, `S${item.rowIndex}`, 0);
    } else {
        item.slg_xuat = item.so_luong;
        await updateSheetValue(CONFIG.udctSheetName, `S${item.rowIndex}`, item.slg_xuat);
    }
    renderUDCTTable();
}

async function updateSheetValue(sheetName, range, value) {
    try {
        const token = await getAccessToken();
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!${range}?valueInputOption=USER_ENTERED`;
        const resp = await fetch(url, {
            method: 'PUT',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ values: [[value]] })
        });
        return resp.ok;
    } catch (err) {
        console.error("Update cell error:", err);
        return false;
    }
}

async function updateUDCTPrice(originalIndex, silent = false) {
    const item = udctData[originalIndex];
    if (!item) return;

    const searchId = (item.id_sp_ct || item.id_sp || "").toString().trim().toLowerCase();
    const searchSku3 = (item.id_sp || "").toString().trim().toLowerCase();
    const searchFallback = (item.id_sp_ct || "").toString().trim().toLowerCase();
    const sp = sanphamData.find(s => {
        const sku3 = (s.sku_3 || "").toString().trim().toLowerCase();
        const skuCon = (s.sku_con || "").toString().trim().toLowerCase();
        return sku3 === searchId || skuCon === searchId || sku3 === searchSku3 || sku3 === searchFallback || skuCon === searchSku3 || skuCon === searchFallback;
    });

    if (!sp) {
        const sku = item.id_sp_ct || item.id_sp || "N/A";
        console.warn(`[PriceUpdate] Không tìm thấy SKU match cho: "${searchId}" (Row ${item.rowIndex})`);
        if (!silent) alert(`LỖI: Không tìm thấy SKU "${sku}" trong cột "Mã" của sheet Sản phẩm PM.`);
        return false;
    }

    const newPrice = sp.gia_ban - sp.gia_dong_goi;
    item.don_gia_1 = newPrice;

    if (!(item.trang_thai || '').toLowerCase().includes('hủy')) {
        item.slg_xuat = item.so_luong;
        await updateSheetValue(CONFIG.udctSheetName, `S${item.rowIndex}`, item.slg_xuat);
    }

    // Sheet index is rowIndex, column AE is index 31 (1-based for range)
    const success = await updateSheetValue(CONFIG.udctSheetName, `AE${item.rowIndex}`, newPrice);

    if (success) {
        if (!silent) {
            const row = udctData.find(d => d.rowIndex === item.rowIndex);
            if (row && !(row.trang_thai || '').toLowerCase().includes('hủy')) {
                row.slg_xuat = row.so_luong;
            }
            renderUDCTTable();
            console.log(`Updated row ${item.rowIndex} price to ${newPrice}`);
        } else if (suggestionType === 'row_add') {
            const idInput = document.getElementById('row_add_id_sp_ct');
            const tenInput = document.getElementById('row_add_ten');
            const giaInput = document.getElementById('row_add_gia_nhap');
            if (idInput) idInput.value = id;
            if (tenInput) tenInput.value = ten;
            if (giaInput) giaInput.value = gia || 0;
            // Optional: auto focus sl
            const slInput = document.getElementById('row_add_sl');
            if (slInput) slInput.focus();
        } else {
            if (!silent) alert(`Lỗi khi cập nhật Google Sheet cho dòng ${item.rowIndex}`);
        }
        return success;
    }
}

async function updateAllPricesBatch() {
    if (!filteredUDCT.length) return alert("Không có dữ liệu đang hiển thị để cập nhật.");
    if (!confirm(`Bạn có chắc muốn cập nhật đơn giá cho ${filteredUDCT.length} dòng đang hiển thị?`)) return;

    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');

    try {
        const token = await getAccessToken();
        const updates = [];

        for (const item of filteredUDCT) {
            const searchId = (item.id_sp_ct || item.id_sp || "").toString().trim().toLowerCase();
            const searchSku3 = (item.id_sp || "").toString().trim().toLowerCase();
            const searchFallback = (item.id_sp_ct || "").toString().trim().toLowerCase();
            const sp = sanphamData.find(s => {
                const sku3 = (s.sku_3 || "").toString().trim().toLowerCase();
                const skuCon = (s.sku_con || "").toString().trim().toLowerCase();
                return sku3 === searchId || skuCon === searchId || sku3 === searchSku3 || sku3 === searchFallback || skuCon === searchSku3 || skuCon === searchFallback;
            });

            if (sp) {
                const newPrice = sp.gia_ban - sp.gia_dong_goi;
                item.don_gia_1 = newPrice;
                updates.push({
                    range: `${CONFIG.udctSheetName}!AE${item.rowIndex}`,
                    values: [[newPrice]]
                });

                if (!(item.trang_thai || '').toLowerCase().includes('hủy')) {
                    item.slg_xuat = item.so_luong;
                    updates.push({
                        range: `${CONFIG.udctSheetName}!S${item.rowIndex}`,
                        values: [[item.slg_xuat]]
                    });
                }
            }
        }

        if (updates.length === 0) {
            alert("Không tìm thấy SKU phù hợp để cập nhật.");
            loadingOverlay.classList.add('hidden');
            return;
        }

        // Chunking để đảm bảo ổn định (500 dòng mỗi batch)
        const chunkSize = 500;
        let successCount = 0;
        for (let i = 0; i < updates.length; i += chunkSize) {
            const chunk = updates.slice(i, i + chunkSize);
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
            const body = {
                valueInputOption: 'USER_ENTERED',
                data: chunk
            };
            const resp = await fetch(url, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
                body: JSON.stringify(body)
            });
            if (resp.ok) successCount += chunk.filter(c => c.range.includes('AE')).length;
        }

        renderUDCTTable();
        alert(`Đã cập nhật đơn giá cho ${successCount} đơn hàng thành công!`);
    } catch (err) {
        console.error("Batch price update error:", err);
        alert("Có lỗi xảy ra khi cập nhật giá hàng loạt.");
    } finally {
        loadingOverlay.classList.add('hidden');
    }
}

async function batchUpdateUDCTStatus(statusValue) {
    if (!filteredUDCT.length) return alert("Không có dữ liệu đang hiển thị để cập nhật.");
    const actionName = statusValue === '2 HỦY' ? 'HỦY NHANH' : (statusValue === '4 MAI GỌI' ? 'MAI GỌI' : 'BỎ TRẠNG THÁI');
    if (!confirm(`Bạn có chắc muốn ${actionName} toàn bộ ${filteredUDCT.length} đơn hàng đang hiển thị?`)) return;

    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');

    try {
        const token = await getAccessToken();
        const updates = [];

        filteredUDCT.forEach(item => {
            const isHuy = (statusValue || '').toString().toLowerCase().includes('hủy');
            item.trang_thai = statusValue;
            if (isHuy) {
                item.slg_xuat = 0;
            } else {
                item.slg_xuat = item.so_luong;
            }

            updates.push({
                range: `${CONFIG.udctSheetName}!Y${item.rowIndex}`,
                values: [[statusValue]]
            });
            updates.push({
                range: `${CONFIG.udctSheetName}!S${item.rowIndex}`,
                values: [[item.slg_xuat]]
            });
        });

        // Chunking để đảm bảo ổn định
        const chunkSize = 400;
        for (let i = 0; i < updates.length; i += chunkSize) {
            const chunk = updates.slice(i, i + chunkSize);
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
            const resp = await fetch(url, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    valueInputOption: 'USER_ENTERED',
                    data: chunk
                })
            });
            if (!resp.ok) console.error("Batch status update error:", await resp.text());
        }

        renderUDCTTable();
        alert(`Đã ${actionName} ${filteredUDCT.length} đơn hàng thành công!`);
    } catch (err) {
        console.error("Batch status update error:", err);
        alert("Có lỗi xảy ra: " + err.message);
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

let sanphamCurrentPage = 1;
const SP_PER_PAGE = 100;

async function loadSanphamData() {
    const tbody = document.getElementById('sanphamTableBody');
    if (tbody) tbody.innerHTML = generateSkeletonRows(7, 10);

    const data = await fetchSheetData(CONFIG.sanphamSheetName);
    if (data && data.length > 0) {
        const headers = data[0].map(h => (h || "").toString().toLowerCase().trim());

        const findIdx = (names) => {
            for (const name of names) {
                const idx = headers.indexOf(name.toLowerCase());
                if (idx !== -1) return idx;
            }
            return -1;
        };

        const idxSkuCon = findIdx(['mã', 'sku_con', 'sku con']);
        const idxGiaBan = findIdx(['giá bán lẻ', 'gia_ban', 'giá bán']);
        const idxGiaDongGoi = findIdx(['giá đón gói', 'gia_dong_goi', 'giá đóng gói']);
        const idxTen = findIdx(['tên', 'ten_sp', 'tên sản phẩm']);

        sanphamData = data.slice(1).map((row, idx) => ({
            rowIndex: idx + 2,
            sku_con: idxSkuCon !== -1 ? row[idxSkuCon] : (row[5] || ""),
            gia_ban: parseFloat(idxGiaBan !== -1 ? row[idxGiaBan] : row[9]) || 0,
            gia_dong_goi: parseFloat(idxGiaDongGoi !== -1 ? row[idxGiaDongGoi] : row[11]) || 0,
            ten_sp: idxTen !== -1 ? row[idxTen] : (row[1] || ""),
            mapping: row
        })).filter(item => item.mapping.some(cell => String(cell || '').trim() !== ''));

        renderSanphamTable(1);
        populateSPLists();
    } else {
        if (tbody) tbody.innerHTML = '<tr><td colspan="32" class="text-center py-8 text-slate-500">Không có dữ liệu</td></tr>';
    }
}

function renderSanphamTable(page = 1) {
    saveFiltersToCache();
    const tbody = document.getElementById('sanphamTableBody');
    if (!tbody) return;

    const search = (document.getElementById('spFilterSearch')?.value || '').toLowerCase().trim();

    let filtered = sanphamData;
    if (search) filtered = filtered.filter(item =>
        item.mapping.some(c => (c || '').toString().toLowerCase().includes(search))
    );

    const totalRows = filtered.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / SP_PER_PAGE));
    sanphamCurrentPage = Math.min(page, totalPages);

    const start = (sanphamCurrentPage - 1) * SP_PER_PAGE;
    const paginated = filtered.slice(start, start + SP_PER_PAGE);

    document.getElementById('spStatsText').textContent = `Hiển thị: ${paginated.length}/${totalRows} sản phẩm (Tổng: ${sanphamData.length})`;
    document.getElementById('spPageInfo').textContent = `Trang ${sanphamCurrentPage} / ${totalPages}`;
    document.getElementById('spPrevBtn').disabled = sanphamCurrentPage <= 1;
    document.getElementById('spNextBtn').disabled = sanphamCurrentPage >= totalPages;

    if (!paginated.length) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center py-8 text-slate-500">Không tìm thấy sản phẩm phù hợp</td></tr>';
        return;
    }

    tbody.innerHTML = paginated.map(item => `
                <tr class="border-b border-slate-100 hover:bg-slate-50">
                    <td class="px-2 py-1.5 text-xs text-slate-900 max-w-[240px] truncate">${item.mapping[1] || '-'}</td>
                    <td class="px-2 py-1.5 text-xs text-slate-900 max-w-[200px] truncate">${item.mapping[2] || '-'}</td>
                    <td class="px-2 py-1.5 text-xs text-slate-900 whitespace-nowrap">${item.mapping[5] || '-'}</td>
                    <td class="px-2 py-1.5 text-xs text-slate-900 whitespace-nowrap">${item.mapping[7] || '-'}</td>
                    <td class="px-2 py-1.5 text-xs text-slate-900 whitespace-nowrap">${item.mapping[9] || '-'}</td>
                    <td class="px-2 py-1.5 text-xs text-slate-900 whitespace-nowrap">${item.mapping[11] || '-'}</td>
                    <td class="px-2 py-1.5 text-xs text-slate-900 whitespace-nowrap">${item.mapping[31] || '-'}</td>
                </tr>
            `).join('');
}

function changeSanphamPage(delta) {
    renderSanphamTable(sanphamCurrentPage + delta);
}

// Hàm xây dựng dữ liệu UPMISA từ udctData
function buildUpmisaData() {
    if (!udctData.length) return;

    upmisaData = [];

    const extraColumns = [
        { name: "Hiển thị trên sổ", default: "0" }, { name: "Hình thức bán hàng", default: "0" }, { name: "Phương thức thanh toán", default: "0" },
        { name: "Kiêm phiếu xuất kho", default: "1" }, { name: "Lập kèm hóa đơn", default: "0" }, { name: "Đã lập hóa đơn", default: "0" },
        { name: "Ngày hạch toán (*)", default: "" }, { name: "Ngày chứng từ (*)", default: "" }, { name: "Số chứng từ (*)", default: "" },
        { name: "Số phiếu xuất", default: "" }, { name: "Lý do xuất", default: "" }, { name: "Số hóa đơn", default: "" },
        { name: "Ngày hóa đơn", default: "" }, { name: "Mã đơn hàng", default: "" }, { name: "Mã thống kê", default: "" },
        { name: "Mã khách hàng", default: "" }, { name: "Tên khách hàng", default: "" }, { name: "Địa chỉ", default: "" },
        { name: "Mã số thuế", default: "" }, { name: "Diễn giải", default: "" }, { name: "Nộp vào TK", default: "" },
        { name: "NV bán hàng", default: "" }, { name: "Mã hàng (*)", default: "" }, { name: "Tên hàng", default: "" },
        { name: "Hàng khuyến mại", default: "" }, { name: "TK Tiền/Chi phí/Nợ (*)", default: "131" }, { name: "TK Doanh thu/Có (*)", default: "5111" },
        { name: "ĐVT", default: "" }, { name: "Số lượng", default: "" }, { name: "Đơn giá sau thuế", default: "" },
        { name: "Đơn giá", default: "" }, { name: "Thành tiền", default: "" }, { name: "Tỷ lệ CK (%)", default: "" },
        { name: "Tiền chiết khấu", default: "" }, { name: "TK chiết khấu", default: "" }, { name: "Giá tính thuế XK", default: "" },
        { name: "% thuế XK", default: "" }, { name: "Tiền thuế XK", default: "" }, { name: "TK thuế XK", default: "" },
        { name: "% thuế GTGT", default: "" }, { name: "Tiền thuế GTGT", default: "" }, { name: "TK thuế GTGT", default: "" },
        { name: "HH không TH trên tờ khai thuế GTGT", default: "" }, { name: "Kho", default: "KMN" }, { name: "TK giá vốn", default: "632" },
        { name: "TK Kho", default: "1561" }, { name: "Đơn giá vốn", default: "" }, { name: "Tiền vốn", default: "" }, { name: "Hàng hóa giữ hộ/bán hộ", default: "" }
    ];

    const formatDateShortFromStr = (dateStr) => {
        if (!dateStr) return "";
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return "";
        const d = date.getDate().toString().padStart(2, '0');
        const m = (date.getMonth() + 1).toString().padStart(2, '0');
        const y = date.getFullYear().toString().slice(-2);
        return `${d}${m}${y}`;
    };

    const getCellValue = (item, colName) => {
        const ngayStr = item.ngay ? item.ngay.split(' ')[0] : '';
        const formatNgayVN = (s) => {
            if (!s) return '';
            const parts = s.split('-');
            if (parts.length !== 3) return s;
            const y = parts[0];
            const m = parts[1].padStart(2, '0');
            const d = parts[2].padStart(2, '0');
            return `${d}/${m}/${y}`;
        };
        const ngayVN = formatNgayVN(ngayStr);

        if (colName === "Ngày hạch toán (*)" || colName === "Ngày chứng từ (*)") return ngayVN;

        if (colName === "Số chứng từ (*)" || colName === "Số phiếu xuất") {
            let prefix = "";
            const san = (item.san || "").toLowerCase().trim();
            const kh = (item.khung_h || "").toUpperCase().trim();

            if (san === "shopee") prefix = (kh === "10H" ? "N " : "") + "SPE";
            else if (san === "lazada") prefix = "LDZ";
            else if (san === "best") prefix = "BEST";
            else if (san === "tiktok") prefix = "TT";
            else if (san === "đơn ngoài") prefix = "XDN";

            const mg = (item.ma_gian || "").trim();
            const nf = formatDateShortFromStr(ngayStr);
            const suffix = (item.mien || "MN").toUpperCase();
            return prefix && mg && nf ? `${prefix}-${mg}-${nf}.${suffix}` : "";
        }

        if (colName === "Mã đơn hàng") return (item.mdh && item.mvd) ? `${item.mdh}/${item.mvd}` : (item.mdh || item.mvd || "");
        if (colName === "Mã khách hàng") return item.ma_gian || "";
        if (colName === "Diễn giải") {
            const san = (item.san || "").trim();
            const kh = (item.khung_h || "").toUpperCase().trim();
            const mien = (item.mien || "").trim();
            const p = (san === "ĐƠN NGOÀI") ? " " : (kh === "10H" ? " NOW " : " ");
            return `${mien}${p}${san} NGÀY ${ngayVN}`;
        }

        if (colName === "Mã hàng (*)") return item.id_sp || "";
        if (colName === "Số lượng") return item.slg_xuat || "";
        if (colName === "Đơn giá") return (parseFloat(item.don_gia_1) || 0).toString();
        if (colName === "Thành tiền") return ((parseFloat(item.don_gia_1) || 0) * (parseFloat(item.slg_xuat) || 0)).toString();
        if (colName === "Kho") return "K" + (item.mien || "").toUpperCase();

        const ex = extraColumns.find(c => c.name === colName);
        return ex ? ex.default : "";
    };

    upmisaData = [];

    // Sắp xếp udctData theo ngày giảm dần (mới nhất lên đầu) trước khi build
    const sortedUdct = [...udctData].sort((a, b) => {
        const da = a.ngay ? new Date(a.ngay) : new Date(0);
        const db = b.ngay ? new Date(b.ngay) : new Date(0);
        return db - da;
    });

    sortedUdct.forEach(item => {
        if (item.trang_thai && item.trang_thai.toUpperCase().includes("HỦY")) return;

        const upmisaRow = extraColumns.map(col => getCellValue(item, col.name));
        upmisaData.push(upmisaRow);
    });

    renderUpmisaTable();
}

let upmisaCurrentPage = 1;
let filteredUpmisa = [];

function renderUpmisaTable(page = 1) {
    const tbody = document.getElementById('upmisaTableBody');
    const searchVal = document.getElementById('upmisaSearchInput').value.toLowerCase();
    const dateVal = document.getElementById('upmisaDateFilter').value;
    const rowsPerPage = parseInt(document.getElementById('upmisaRowsPerPage').value);

    // 1. Filtering
    let filtered = upmisaData;
    if (dateVal) {
        const [y, m, d] = dateVal.split('-');
        const dateVN = `${d}/${m}/${y}`;
        filtered = filtered.filter(row => row[6] === dateVN);
    }
    if (searchVal) {
        filtered = filtered.filter(row =>
            row.some(cell => cell.toString().toLowerCase().includes(searchVal))
        );
    }

    // 2. Sorting (Date DESC -> Thành tiền ASC)
    filtered.sort((a, b) => {
        const parseDate = (dstr) => {
            if (!dstr) return '';
            const parts = dstr.split('/');
            if (parts.length === 3) return `${parts[2]}${parts[1].padStart(2, '0')}${parts[0].padStart(2, '0')}`;
            return dstr;
        };
        const dateA = parseDate(a[6] || '');
        const dateB = parseDate(b[6] || '');
        if (dateA !== dateB) return dateB.localeCompare(dateA);

        const ttA = parseFloat((a[31] || '').toString().replace(/,/g, '')) || 0;
        const ttB = parseFloat((b[31] || '').toString().replace(/,/g, '')) || 0;
        return ttA - ttB;
    });
    filteredUpmisa = filtered;

    // 3. Pagination
    const totalRows = filteredUpmisa.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / rowsPerPage));
    upmisaCurrentPage = Math.min(page, totalPages);

    const start = (upmisaCurrentPage - 1) * rowsPerPage;
    const end = start + rowsPerPage;
    const paginated = filteredUpmisa.slice(start, end);

    // Update Stats
    document.getElementById('upmisaStatsText').textContent = `Hiển thị: ${paginated.length}/${totalRows} dòng (Tổng: ${upmisaData.length})`;
    document.getElementById('upmisaPageInfo').textContent = `Trang ${upmisaCurrentPage} / ${totalPages}`;
    document.getElementById('upmisaPrevBtn').disabled = upmisaCurrentPage <= 1;
    document.getElementById('upmisaNextBtn').disabled = upmisaCurrentPage >= totalPages;

    if (!paginated.length) {
        tbody.innerHTML = '<tr><td colspan="50" class="text-center py-8 text-slate-500">Không có dữ liệu phù hợp</td></tr>';
        return;
    }

    tbody.innerHTML = paginated.map(row => `
                <tr class="border-b border-slate-100 hover:bg-slate-50">
                    ${row.map(cell => `<td class="px-3 py-2 text-[13px] text-slate-900">${cell || '-'}</td>`).join('')}
                </tr>
            `).join('');
}

function changeUpmisaPage(delta) {
    renderUpmisaTable(upmisaCurrentPage + delta);
}

function refreshUpmisaData() {
    buildUpmisaData();
    alert('Đã làm mới dữ liệu UPMISA!');
}

function exportUpmisaToExcel() {
    if (!filteredUpmisa || !filteredUpmisa.length) {
        alert('Không có dữ liệu hợp lệ để xuất!');
        return;
    }

    const headers = [
        'Hiển thị trên sổ', 'Hình thức bán hàng', 'Phương thức thanh toán', 'Kiêm phiếu xuất kho',
        'Lập kèm hóa đơn', 'Đã lập hóa đơn', 'Ngày hạch toán (*)', 'Ngày chứng từ (*)', 'Số chứng từ (*)',
        'Số phiếu xuất', 'Lý do xuất', 'Số hóa đơn', 'Ngày hóa đơn', 'Mã đơn hàng', 'Mã thống kê',
        'Mã khách hàng', 'Tên khách hàng', 'Địa chỉ', 'Mã số thuế', 'Diễn giải', 'Nộp vào TK',
        'NV bán hàng', 'Mã hàng (*)', 'Tên hàng', 'Hàng khuyến mại', 'TK Tiền/Chi phí/Nợ (*)',
        'TK Doanh thu/Có (*)', 'ĐVT', 'Số lượng', 'Đơn giá sau thuế', 'Đơn giá', 'Thành tiền',
        'Tỷ lệ CK (%)', 'Tiền chiết khấu', 'TK chiết khấu', 'Giá tính thuế XK', '% thuế XK',
        'Tiền thuế XK', 'TK thuế XK', '% thuế GTGT', 'Tiền thuế GTGT', 'TK thuế GTGT',
        'HH không TH trên tờ khai thuế GTGT', 'Kho', 'TK giá vốn', 'TK Kho', 'Đơn giá vốn',
        'Tiền vốn', 'Hàng hóa giữ hộ/bán hộ'
    ];

    const excelData = [headers, ...filteredUpmisa];
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'UPMISA');
    const filterDate = document.getElementById('upmisaDateFilter')?.value || 'TatCa';
    const now = new Date();
    const timeStr = now.getHours().toString().padStart(2, '0') + 'h' + now.getMinutes().toString().padStart(2, '0');
    XLSX.writeFile(wb, `UPMISA_${filterDate}_${timeStr}.xlsx`);
}

function setQuickDate(type) {
    const fromInput = document.getElementById('fromDate');
    const toInput = document.getElementById('toDate');
    const today = new Date();
    let fromDate, toDate;

    const format = (d) => {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    };

    if (type === 'today') {
        fromDate = format(today);
        toDate = format(today);
    } else if (type === 'thisWeek') {
        const day = today.getDay(); // 0 is Sunday
        const diff = today.getDate() - day + (day === 0 ? -6 : 1); // Adjust for Monday start
        const firstDay = new Date(today.setDate(diff));
        fromDate = format(firstDay);

        // Reset today and set to Sunday
        const lastDay = new Date(firstDay);
        lastDay.setDate(firstDay.getDate() + 6);
        toDate = format(lastDay);
    } else if (type === 'thisMonth') {
        const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
        const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);
        fromDate = format(firstDay);
        toDate = format(lastDay);
    }

    fromInput.value = fromDate;
    toInput.value = toDate;
    autoFilterReport(); // Tự động lọc sau khi đổi ngày
}

let filterTimeout;
function autoFilterReport() {
    saveFiltersToCache();
    filterReport();
}

async function filterReport() {
    const fromDate = document.getElementById('fromDate').value;
    const toDate = document.getElementById('toDate').value;
    const filterMaGian = document.getElementById('filterMaGian').value.trim().toLowerCase();
    const filterIdSp = document.getElementById('filterReportIdSp').value.trim().toLowerCase();

    if (!fromDate || !toDate) {
        alert('Vui lòng chọn đầy đủ ngày bắt đầu và ngày kết thúc!');
        return;
    }

    if (!udctData || udctData.length === 0) {
        // Tự động thử tải dữ liệu nếu chưa có
        const loadingOverlay = document.getElementById('loadingOverlay');
        if (loadingOverlay) loadingOverlay.classList.remove('hidden');
        try {
            await loadUDCTData();
        } catch (err) {
            console.error("Auto load UDCT failed:", err);
        } finally {
            if (loadingOverlay) loadingOverlay.classList.add('hidden');
        }
    }

    if (!udctData || udctData.length === 0) {
        alert('Dữ liệu đang được tải hoặc chưa có. Vui lòng đợi trong giây lát hoặc kiểm tra kết nối.');
        return;
    }

    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');

    try {
        // Helper: chuyển dd/mm/yyyy hoặc yyyy-mm-dd → yyyy-mm-dd để so sánh
        const toIsoDate = (str) => {
            if (!str) return '';
            const s = str.split(' ')[0]; // bỏ phần giờ
            if (s.includes('/')) {
                // dd/mm/yyyy
                const [d, m, y] = s.split('/');
                return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
            }
            return s; // đã là yyyy-mm-dd
        };

        // Sử dụng udctData đã load sẵn thay vì fetch lại từ Sheets
        const filtered = udctData.filter(item => {
            const ngayIso = toIsoDate(item.ngay);
            const matchDate = ngayIso >= fromDate && ngayIso <= toDate;
            const matchMaGian = !filterMaGian || (item.ma_gian || '').toLowerCase().includes(filterMaGian);
            const idSpValue = (item.id_sp_ct || item.id_sp || '').toString().toLowerCase();
            const matchIdSp = !filterIdSp || idSpValue.includes(filterIdSp);

            return matchDate && matchMaGian && matchIdSp;
        }).map(item => ({
            ngay: item.ngay,
            ngayDon: toIsoDate(item.ngay),
            san: item.san || 'Khác',
            ma_gian: item.ma_gian || 'N/A',
            id_sp: item.id_sp || '',
            id_sp_ct: item.id_sp_ct || item.id_sp || '',
            mvd: item.mvd || '',
            mdh: item.mdh || '',
            ten_sp: item.ten_sp || '',
            slg_xuat: parseFloat(item.slg_xuat) || 0,
            don_gia: parseFloat(item.don_gia_1) || 0
        }));

        reportData = filtered;
        const totalOrders = filtered.length;

        let totalRevenue = 0;
        const sanStats = {};
        const magianStats = {};
        const idspStats = {};
        const dailyStats = {};

        for (const item of filtered) {
            const rev = item.don_gia * item.slg_xuat;
            totalRevenue += rev;

            if (!sanStats[item.san]) sanStats[item.san] = { so_don: 0, doanh_thu: 0 };
            sanStats[item.san].so_don++;
            sanStats[item.san].doanh_thu += rev;

            if (!magianStats[item.ma_gian]) magianStats[item.ma_gian] = { so_don: 0, doanh_thu: 0 };
            magianStats[item.ma_gian].so_don++;
            magianStats[item.ma_gian].doanh_thu += rev;

            const idKey = item.id_sp_ct || 'N/A';
            if (!idspStats[idKey]) idspStats[idKey] = { ten_sp: item.ten_sp || '', slg: 0, doanh_thu: 0 };
            idspStats[idKey].slg += item.slg_xuat;
            idspStats[idKey].doanh_thu += rev;

            if (!dailyStats[item.ngayDon]) dailyStats[item.ngayDon] = { so_don: 0, doanh_thu: 0 };
            dailyStats[item.ngayDon].so_don++;
            dailyStats[item.ngayDon].doanh_thu += rev;
        }

        document.getElementById('totalOrders').textContent = totalOrders.toLocaleString('vi-VN');
        document.getElementById('totalRevenue').textContent = totalRevenue.toLocaleString('vi-VN');
        document.getElementById('totalSan').textContent = Object.keys(sanStats).length;

        // Render bảng mã gian
        document.getElementById('magianTableBody').innerHTML = Object.entries(magianStats)
            .sort((a, b) => b[1].doanh_thu - a[1].doanh_thu)
            .map(([mg, s]) => `
                        <tr class="border-b border-slate-100 hover:bg-slate-50">
                            <td class="px-4 py-3 text-sm font-medium text-slate-900">${mg}</td>
                            <td class="px-4 py-3 text-sm text-slate-700">${s.so_don.toLocaleString('vi-VN')}</td>
                            <td class="px-4 py-3 text-sm text-slate-700">${s.doanh_thu.toLocaleString('vi-VN')}</td>
                        </tr>`).join('') || '<tr><td colspan="3" class="text-center py-4 text-slate-400">Không có dữ liệu</td></tr>';

        document.getElementById('idspTableBody').innerHTML = Object.entries(idspStats)
            .sort((a, b) => b[1].doanh_thu - a[1].doanh_thu)
            .map(([idsp, s]) => `
                        <tr class="border-b border-slate-100 hover:bg-slate-50">
                            <td class="px-4 py-3 text-sm font-medium text-primary">${idsp}</td>
                            <td class="px-4 py-3 text-sm text-slate-700 max-w-[160px] truncate" title="${s.ten_sp}">${s.ten_sp || '-'}</td>
                            <td class="px-4 py-3 text-sm text-slate-700">${s.slg.toLocaleString('vi-VN')}</td>
                            <td class="px-4 py-3 text-sm text-slate-700">${s.doanh_thu.toLocaleString('vi-VN')}</td>
                        </tr>`).join('') || '<tr><td colspan="4" class="text-center py-4 text-slate-400">Không có dữ liệu</td></tr>';

        let textSummary = `BÁO CÁO NGÀY: ${fromDate} đến ${toDate}\n`;
        textSummary += `====================================\n`;
        textSummary += `TỔNG SỐ ĐƠN: ${totalOrders}\n`;
        textSummary += `TỔNG DOANH THU: ${totalRevenue.toLocaleString('vi-VN')}\n\n`;
        textSummary += `\n`;
        Object.entries(sanStats).forEach(([san, stats]) => {
            textSummary += `- Sàn ${san}: ${stats.so_don} đơn - ${stats.doanh_thu.toLocaleString('vi-VN')}\n`;
        });
        if (filterMaGian) textSummary += `\nLỌC THEO MÃ GIAN: ${filterMaGian}\n`;
        document.getElementById('textReportArea').textContent = textSummary;

        renderCharts(dailyStats);

        window.__detailRows = filtered;
        renderDetailTable();

    } catch (error) {
        console.error('Filter error:', error);
        alert('Có lỗi xảy ra khi lọc dữ liệu!');
    } finally {
        loadingOverlay.classList.add('hidden');
    }
}

const detailSortState = {
    key: 'ngay',
    dir: 'desc'
};

function normalizeDetailValue(item, key) {
    if (key === 'ngay') return item.ngay || '';
    if (key === 'san') return item.san || '';
    if (key === 'mvd') return item.mvd || '';
    if (key === 'mdh') return item.mdh || '';
    if (key === 'ten_sp') return item.ten_sp || '';
    if (key === 'id_sp') return item.id_sp || '';
    if (key === 'slg_xuat') return parseFloat(item.slg_xuat) || 0;
    if (key === 'don_gia') return parseFloat(item.don_gia) || 0;
    if (key === 'thanh_tien') return (parseFloat(item.don_gia) || 0) * (parseFloat(item.slg_xuat) || 0);
    return '';
}

function updateDetailSortIndicators() {
    const keys = ['Ngay', 'San', 'Mvd', 'Mdh', 'TenSp', 'IdSp', 'SlgXuat', 'DonGia', 'ThanhTien'];
    keys.forEach(k => {
        const el = document.getElementById(`detailSort${k}`);
        if (el) el.textContent = '↕';
    });
    const activeKeyMap = {
        ngay: 'Ngay', san: 'San', mvd: 'Mvd', mdh: 'Mdh', ten_sp: 'TenSp', id_sp: 'IdSp', slg_xuat: 'SlgXuat', don_gia: 'DonGia', thanh_tien: 'ThanhTien'
    };
    const activeEl = document.getElementById(`detailSort${activeKeyMap[detailSortState.key]}`);
    if (activeEl) activeEl.textContent = detailSortState.dir === 'asc' ? '↑' : '↓';
}

function sortDetailTable(key) {
    if (detailSortState.key === key) {
        detailSortState.dir = detailSortState.dir === 'asc' ? 'desc' : 'asc';
    } else {
        detailSortState.key = key;
        detailSortState.dir = key === 'ngay' ? 'desc' : 'asc';
    }
    renderDetailTable();
}

function renderDetailTable() {
    const detailBody = document.getElementById('detailTableBody');
    if (!detailBody) return;
    const rows = (window.__detailRows || []).slice();
    rows.sort((a, b) => {
        const av = normalizeDetailValue(a, detailSortState.key);
        const bv = normalizeDetailValue(b, detailSortState.key);
        let cmp = 0;
        if (typeof av === 'number' && typeof bv === 'number') cmp = av - bv;
        else cmp = String(av).localeCompare(String(bv), 'vi', { numeric: true, sensitivity: 'base' });
        return detailSortState.dir === 'asc' ? cmp : -cmp;
    });
    if (rows.length === 0) {
        detailBody.innerHTML = '<tr><td colspan="9" class="text-center py-8 text-slate-500">Không có dữ liệu</td></tr>';
    } else {
        detailBody.innerHTML = rows.map(item => `
                    <tr class="border-b border-slate-100 hover:bg-slate-50">
                        <td class="px-4 py-3 text-sm text-slate-900">${item.ngay || '-'}</td>
                        <td class="px-4 py-3 text-sm text-slate-900">${item.san || '-'}</td>
                        <td class="px-4 py-3 text-sm text-slate-900">${item.mvd || '-'}</td>
                        <td class="px-4 py-3 text-sm text-slate-900">${item.mdh || '-'}</td>
                        <td class="px-4 py-3 text-sm text-slate-900">${item.ten_sp || '-'}</td>
                        <td class="px-4 py-3 text-sm text-slate-900">${item.id_sp || '-'}</td>
                        <td class="px-4 py-3 text-sm text-slate-900">${item.slg_xuat}</td>
                        <td class="px-4 py-3 text-sm text-slate-900">${parseFloat(item.don_gia || 0).toLocaleString('vi-VN')}</td>
                        <td class="px-4 py-3 text-sm text-slate-900 font-medium">${((parseFloat(item.don_gia) || 0) * (parseFloat(item.slg_xuat) || 0)).toLocaleString('vi-VN')}</td>
                    </tr>
                `).join('');
    }
    updateDetailSortIndicators();
}

function renderCharts(dailyStats) {
    const labels = Object.keys(dailyStats).sort();
    const revenueData = labels.map(l => dailyStats[l].doanh_thu);
    const ordersData = labels.map(l => dailyStats[l].so_don);

    if (mergedChart) mergedChart.destroy();

    const ctx = document.getElementById('mergedChart').getContext('2d');
    mergedChart = new Chart(ctx, {
        data: {
            labels: labels,
            datasets: [
                {
                    type: 'bar',
                    label: 'Doanh thu (đ)',
                    data: revenueData,
                    backgroundColor: 'rgba(37, 99, 235, 0.7)',
                    borderColor: '#2563eb',
                    borderWidth: 1,
                    yAxisID: 'yRev',
                    order: 2
                },
                {
                    type: 'line',
                    label: 'Số đơn (đơn)',
                    data: ordersData,
                    borderColor: '#ef4444',
                    backgroundColor: '#ef4444',
                    borderWidth: 3,
                    pointRadius: 4,
                    pointHoverRadius: 6,
                    fill: false,
                    tension: 0.3,
                    yAxisID: 'yOrders',
                    order: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: { position: 'top', align: 'end' },
                tooltip: { backgroundColor: 'rgba(0,0,0,0.8)', padding: 12 }
            },
            scales: {
                x: { grid: { display: false } },
                yRev: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    title: { display: true, text: 'Doanh thu (vnđ)', color: '#2563eb' },
                    grid: { color: '#f1f5f9' },
                    beginAtZero: true
                },
                yOrders: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    title: { display: true, text: 'Số lượng đơn', color: '#ef4444' },
                    grid: { drawOnChartArea: false },
                    beginAtZero: true
                }
            }
        }
    });
}

function copyTextReport() {
    const content = document.getElementById('textReportArea').textContent;

    const textArea = document.createElement("textarea");
    textArea.value = content;

    textArea.style.top = "0";
    textArea.style.left = "0";
    textArea.style.position = "fixed";
    textArea.style.opacity = "0";

    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();

    try {
        const successful = document.execCommand('copy');
        if (successful) {
            alert('Đã sao chép báo cáo vào bộ nhớ tạm!');
        } else {
            alert('Lỗi khi sao chép!');
        }
    } catch (err) {
        console.error('Copy error:', err);
        alert('Lỗi khi sao chép!');
    }

    document.body.removeChild(textArea);
}

function exportReportToExcel() {
    if (reportData.length === 0) {
        alert('Không có dữ liệu để xuất! Vui lòng chọn ngày và lọc trước.');
        return;
    }

    const excelData = [
        ['BÁO CÁO ĐƠN HÀNG CHI TIẾT'],
        [`Từ ngày: ${document.getElementById('fromDate').value} đến ${document.getElementById('toDate').value}`],
        [`Mã gian lọc: ${document.getElementById('filterMaGian').value || 'Toàn bộ'}`],
        [],
        ['1. THỐNG KÊ TỔNG QUAN'],
        ['Tổng số đơn', document.getElementById('totalOrders').textContent],
        ['Tổng doanh thu', document.getElementById('totalRevenue').textContent],
        ['Số sàn', document.getElementById('totalSan').textContent],
        [],
        ['2. THỐNG KÊ THEO MÃ GIAN'],
        ['Mã gian', 'Số đơn', 'Doanh thu']
    ];

    const magianRows = document.querySelectorAll('#magianTableBody tr');
    magianRows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if (cells.length === 3) {
            excelData.push([cells[0].textContent, cells[1].textContent, cells[2].textContent]);
        }
    });

    excelData.push([], ['3. CHI TIẾT ĐƠN HÀNG']);
    excelData.push(['Ngày', 'Sàn', 'MVD', 'MDH', 'Tên SP', 'SLG xuất', 'Đơn giá', 'Thành tiền']);

    const detailRows = document.querySelectorAll('#detailTableBody tr');
    detailRows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if (cells.length === 8) {
            excelData.push(Array.from(cells).map(c => c.textContent));
        }
    });

    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'BaoCao');
    XLSX.writeFile(wb, `BaoCao_ERP_${document.getElementById('fromDate').value}_${document.getElementById('toDate').value}.xlsx`);
}

function exportIdSPExcel() {
    const tbody = document.getElementById('idspTableBody');
    if (!tbody || tbody.innerHTML.includes('Không lấy được dữ liệu') || tbody.innerHTML.includes('Chọn ngày')) {
        alert('Không có dữ liệu để xuất! Vui lòng chọn ngày và lọc báo cáo trước.');
        return;
    }

    const excelData = [['id_sp_ct', 'SL xuất']];
    const rows = tbody.querySelectorAll('tr');

    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if (cells.length === 4) { // ID SP, Tên SP, SL xuất, Doanh thu
            excelData.push([
                cells[0].textContent.trim(),
                cells[2].textContent.replace(/,/g, '').replace(/\./g, '').trim() // Remove formatting
            ]);
        }
    });

    if (excelData.length <= 1) {
        alert('Không có dữ liệu hợp lệ để xuất!');
        return;
    }

    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'ThongKe_ID_SP');
    XLSX.writeFile(wb, `ThongKe_ID_SP_${document.getElementById('fromDate').value}_${document.getElementById('toDate').value}.xlsx`);
}

function parseFileName(fileName) {
    const parts = fileName.replace('.xlsx', '').replace('.xls', '').split('_');
    let mienRaw = parts[2] || '';
    let sanRaw = parts[3] || '';
    const ngay = parts[4] || '';
    let khung_h = parts[6] || '';

    // Cắt ngắn Khung H chỉ lấy tới chữ "H"
    if (khung_h.toUpperCase().includes('H')) {
        khung_h = khung_h.toUpperCase().split('H')[0] + 'H';
    }

    // Mapping Miền (Region)
    let mien = mienRaw;
    const mienUpper = mienRaw.toUpperCase();
    if (mienUpper === 'HCM') mien = 'MN';
    else if (mienUpper === 'HN') mien = 'MB';

    // Mapping Sàn (Platform)
    let san = sanRaw;
    const sanUpper = sanRaw.toUpperCase();
    if (sanUpper === 'SHP') san = 'shopee';
    else if (sanUpper === 'DN') san = 'đơn ngoài';
    else if (sanUpper === 'BEST') san = 'best';
    else if (sanUpper === 'TT') san = 'tiktok';

    return { mien, san, ngay, khung_h };
}

async function handleExcelUploadDonhang(files) {
    if (currentUser && ['demo', 'kinhdoanh'].includes(currentUser.role)) {
        alert('Tài khoản này không được phép tải Excel cho Đơn chi tiết.');
        return;
    }
    if (!files || files.length === 0) return;
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');

    try {
        const appendValues = [];
        let filesProcessed = 0;

        for (const file of files) {
            const fileNameInfo = parseFileName(file.name);
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const sheetName = workbook.SheetNames.find(name => name.includes('Tổng đơn') || name === 'Tổng đơn 1');
            const firstSheet = sheetName ? workbook.Sheets[sheetName] : workbook.Sheets[workbook.SheetNames[0]];
            const excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

            if (!excelData || excelData.length < 2) {
                console.warn(`File ${file.name} không có dữ liệu!`);
                continue;
            }

            const rows = excelData.slice(1);
            for (const row of rows) {
                const san = fileNameInfo.san;
                const kh = (fileNameInfo.khung_h || "").toUpperCase();

                const c_b = (row[1] || '').toString();
                const c_c = (row[2] || '').toString();
                let c_f = fileNameInfo.ngay;
                // Chuyển định dạng ngày từ yyyy-mm-dd sang dd/mm/yyyy
                if (c_f && c_f.includes('-')) {
                    const dParts = c_f.split('-');
                    if (dParts.length === 3 && dParts[0].length === 4) {
                        c_f = `${dParts[2]}/${dParts[1]}/${dParts[0]}`;
                    }
                }
                const c_i = (row[8] || '').toString(); // Tên SP
                const c_j = (row[9] || '').toString(); // Mã SKU phân loại (sku_shop_up)

                let c_d = ''; // MDH
                let c_g = ''; // MVD
                let c_p = ''; // Số lượng

                if (san === 'shopee') {
                    if (kh === '10H') {
                        c_d = (row[3] || '').toString();
                        c_g = (row[3] || '').toString();
                        c_p = (row[15] || '').toString();
                    } else {
                        c_d = (row[3] || '').toString();
                        c_g = (row[6] || '').toString();
                        c_p = (row[15] || '').toString();
                    }
                } else if (san === 'tiktok' || san === 'best' || san === 'đơn ngoài') {
                    c_d = (row[6] || '').toString();
                    c_g = (row[6] || '').toString();
                    c_p = (row[10] || '').toString();
                } else {
                    c_d = (row[3] || '').toString();
                    c_g = (row[6] || '').toString();
                    c_p = (row[15] || '').toString();
                }

                if (!c_d && !c_g) continue;

                const sku_shop_up = c_j; // Lấy nguyên cột J
                const id_sp = c_c;
                let id_sp_ct = '';
                if (c_j.substring(0, 4) === c_c) {
                    id_sp_ct = c_j.substring(0, 10);
                }

                const random8 = generateRandomString(8);
                const id_parts = [c_f, 'xuất', 'hằng ngày', fileNameInfo.mien, fileNameInfo.san, fileNameInfo.khung_h, c_b, c_d, c_g, sku_shop_up, c_p, random8];
                const id_up_don_parts = [c_f, 'xuất', 'hằng ngày', fileNameInfo.mien, fileNameInfo.san, fileNameInfo.khung_h, c_b, c_d, c_g];
                const id_dh_parts = [c_f, 'xuất', 'hằng ngày', fileNameInfo.mien];
                const id_dh_ct_parts = [c_f, 'xuất', 'hằng ngày', fileNameInfo.mien, id_sp];

                const newId = generateId(id_parts).toUpperCase();
                const id_up_don = generateId(id_up_don_parts).toUpperCase();
                const id_dh = generateId(id_dh_parts).toUpperCase();
                const id_dh_ct = generateId(id_dh_ct_parts).toUpperCase();

                appendValues.push([
                    newId, id_up_don, id_dh, id_dh_ct,
                    c_f.toUpperCase(), 'XUẤT', 'HẰNG NGÀY', fileNameInfo.mien.toUpperCase(),
                    fileNameInfo.san, fileNameInfo.khung_h, c_b, c_g, c_d,
                    sku_shop_up, c_p, id_sp, id_sp_ct,
                    c_i, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                    '', '', '', '', '', '', '', '', '', '', '', ''
                ]);
            }
            filesProcessed++;
        }

        if (appendValues.length === 0) {
            alert("Không có dữ liệu hợp lệ để import!");
            return;
        }

        console.log(`Đang nạp ${appendValues.length} dòng từ ${filesProcessed} file vào ${CONFIG.udctSheetName}...`);
        const success = await appendSheetData(CONFIG.udctSheetName, appendValues);

        if (success) {
            alert(`Import thành công ${appendValues.length} dòng dữ liệu từ ${filesProcessed} file!`);
            await loadUDCTData();
        } else {
            alert("Import thất bại! Vui lòng kiểm tra console hoặc quyền truy cập Sheet.");
        }
    } catch (error) {
        console.error("Excel upload error:", error);
        alert("Có lỗi xảy ra khi xử lý file Excel!");
    } finally {
        loadingOverlay.classList.add('hidden');
        document.getElementById('excelUploadDonhang').value = '';
    }
}

async function handleExcelUpload(files) {
    if (currentUser && ['demo', 'kinhdoanh'].includes(currentUser.role)) {
        alert('Tài khoản này không được phép tải Excel lên.');
        return;
    }
    if (!files || files.length === 0) return;
    const file = files[0];
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

        if (!excelData || excelData.length < 2) {
            alert("File Excel không có dữ liệu!");
            return;
        }

        const rows = excelData.slice(1);
        const sheetData = [];

        for (const row of rows) {
            const newRow = [];
            for (let i = 0; i < 32; i++) {
                let value = '';
                if (i < row.length && row[i] !== undefined && row[i] !== null) {
                    value = row[i].toString();
                }
                newRow.push(value);
            }
            sheetData.push(newRow);
        }

        if (sheetData.length === 0) {
            alert("Không có dữ liệu để import!");
            return;
        }

        console.log(`Đang xử lý ${CONFIG.sanphamSheetName}: Xóa cũ và nạp mới...`);
        const cleared = await clearSheetData(CONFIG.sanphamSheetName);
        if (!cleared) {
            if (!confirm("Không thể xóa dữ liệu cũ, bạn có muốn tiếp tục ghi đè không?")) {
                return;
            }
        }

        const success = await appendSheetData(CONFIG.sanphamSheetName, sheetData);

        if (success) {
            alert(`Import thành công ${sheetData.length} dòng dữ liệu vào ${CONFIG.sanphamSheetName}!`);
            await loadSanphamData();
        } else {
            alert("Import thất bại! Vui lòng kiểm tra console để biết thêm chi tiết.");
        }
    } catch (error) {
        console.error("Excel upload error:", error);
        alert("Có lỗi xảy ra khi xử lý file Excel: " + error.message);
    } finally {
        loadingOverlay.classList.add('hidden');
        document.getElementById('excelUpload').value = '';
    }
}

function setupDragAndDrop() {
    const dropZone = document.getElementById('dropZone');
    if (!dropZone) return;

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0 && (files[0].name.endsWith('.xlsx') || files[0].name.endsWith('.xls'))) {
            handleExcelUpload(files);
        } else {
            alert("Vui lòng kéo thả file Excel (.xlsx hoặc .xls)");
        }
    });
}

function setupDragAndDropDonhang() {
    const dropZone = document.getElementById('dropZoneDonhang');
    if (!dropZone) return;

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0 && (files[0].name.endsWith('.xlsx') || files[0].name.endsWith('.xls'))) {
            handleExcelUploadDonhang(files);
        } else {
            alert("Vui lòng kéo thả file Excel (.xlsx hoặc .xls)");
        }
    });
}

function switchModule(module) {
    const homeModule = document.getElementById('moduleHome');
    const donhangModule = document.getElementById('moduleDonhang');
    const sanphamModule = document.getElementById('moduleSanpham');
    const baocaoModule = document.getElementById('moduleBaocao');
    const upmisaModule = document.getElementById('moduleUpmisa');
    const inventoryModule = document.getElementById('moduleInventory');
    const hangHoanModule = document.getElementById('moduleHangHoan');
    const hhShopDienModule = document.getElementById('moduleHHShopDien');
    const pageTitle = document.getElementById('pageTitle');
    const sidebarHome = document.getElementById('sidebarHome');
    const sidebarDonhang = document.getElementById('sidebarDonhang');
    const sidebarSanpham = document.getElementById('sidebarSanpham');
    const sidebarBaocao = document.getElementById('sidebarBaocao');
    const sidebarUpmisa = document.getElementById('sidebarUpmisa');
    const sidebarInventory = document.getElementById('sidebarInventory');
    const sidebarDHCT = document.getElementById('sidebarDHCT');
    const sidebarUniqueDHCT = document.getElementById('sidebarUniqueDHCT');
    const sidebarHangHoan = document.getElementById('sidebarHangHoan');
    const sidebarHHShopDien = document.getElementById('sidebarHHShopDien');

    homeModule.style.display = 'none';
    donhangModule.style.display = 'none';
    sanphamModule.style.display = 'none';
    baocaoModule.style.display = 'none';
    upmisaModule.style.display = 'none';
    inventoryModule.style.display = 'none';
    if (hangHoanModule) hangHoanModule.style.display = 'none';
    if (hhShopDienModule) hhShopDienModule.style.display = 'none';
    const dhctModule = document.getElementById('moduleDHCT');
    if (dhctModule) dhctModule.style.display = 'none';
    const uniqueDHCTModule = document.getElementById('moduleUniqueDHCT');
    if (uniqueDHCTModule) uniqueDHCTModule.style.display = 'none';

    const resetSidebar = () => {
        [sidebarHome, sidebarDonhang, sidebarSanpham, sidebarBaocao, sidebarUpmisa, sidebarInventory, sidebarDHCT, sidebarUniqueDHCT, sidebarHangHoan, sidebarHHShopDien].forEach(s => {
            if (s) {
                s.classList.remove('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
                s.classList.add('text-slate-600');
            }
        });
    };

    if (module === 'home') {
        homeModule.style.display = 'block';
        pageTitle.textContent = 'Trang chủ';
        resetSidebar();
        sidebarHome.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
        sidebarHome.classList.remove('text-slate-600');
    } else if (module === 'donhang') {
        donhangModule.style.display = 'flex';
        pageTitle.textContent = 'UP Đơn chi tiết';
        resetSidebar();
        sidebarDonhang.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
        sidebarDonhang.classList.remove('text-slate-600');
    } else if (module === 'sanpham') {
        sanphamModule.style.display = 'flex';
        pageTitle.textContent = 'Sản phẩm phần mềm';
        resetSidebar();
        sidebarSanpham.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
        sidebarSanpham.classList.remove('text-slate-600');
        setTimeout(() => setupDragAndDrop(), 100);
    } else if (module === 'baocao') {
        baocaoModule.style.display = 'flex';
        pageTitle.textContent = 'Báo cáo';
        resetSidebar();
        sidebarBaocao.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
        sidebarBaocao.classList.remove('text-slate-600');
        // Không tự động reset ngày khi chuyển module để giữ bộ lọc local
        if (!document.getElementById('fromDate').value || !document.getElementById('toDate').value) {
            setQuickDate('today');
        } else {
            autoFilterReport();
        }
    } else if (module === 'upmisa') {
        upmisaModule.style.display = 'flex';
        pageTitle.textContent = 'UPMISA';
        resetSidebar();
        sidebarUpmisa.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
        sidebarUpmisa.classList.remove('text-slate-600');

        const today = new Date();
        const d = String(today.getDate()).padStart(2, '0');
        const m = String(today.getMonth() + 1).padStart(2, '0');
        const y = today.getFullYear();
        document.getElementById('upmisaDateFilter').value = `${y}-${m}-${d}`;
        buildUpmisaData();
    } else if (module === 'inventory') {
        inventoryModule.style.display = 'flex';
        pageTitle.textContent = 'Quản lý Tồn Kho';
        resetSidebar();
        sidebarInventory.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
        sidebarInventory.classList.remove('text-slate-600');
        fetchInventoryData();
    } else if (module === 'dh_ct') {
        if (dhctModule) dhctModule.style.display = 'flex';
        pageTitle.textContent = 'Dữ liệu DH Chi Tiết';
        resetSidebar();
        if (sidebarDHCT) {
            sidebarDHCT.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
            sidebarDHCT.classList.remove('text-slate-600');
        }
        fetchDHCTData();
    } else if (module === 'unique_dh_ct') {
        if (uniqueDHCTModule) uniqueDHCTModule.style.display = 'flex';
        pageTitle.textContent = 'Đơn hàng trên DH Chi Tiết';
        resetSidebar();
        if (sidebarUniqueDHCT) {
            sidebarUniqueDHCT.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
            sidebarUniqueDHCT.classList.remove('text-slate-600');
        }
        fetchDHCTData();
    } else if (module === 'hang_hoan') {
        if (hangHoanModule) hangHoanModule.style.display = 'flex';
        pageTitle.textContent = 'Dữ liệu Hàng hoàn';
        resetSidebar();
        if (sidebarHangHoan) {
            sidebarHangHoan.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
            sidebarHangHoan.classList.remove('text-slate-600');
        }
        fetchHangHoanData();
    } else if (module === 'hh_shop_dien') {
        if (hhShopDienModule) hhShopDienModule.style.display = 'flex';
        pageTitle.textContent = 'HH SHOP ĐIỀN';
        resetSidebar();
        if (sidebarHHShopDien) {
            sidebarHHShopDien.classList.add('active', 'bg-blue-50', 'text-primary', 'border-r-2', 'border-primary');
            sidebarHHShopDien.classList.remove('text-slate-600');
        }
        fetchHHShopDienData();
    }
    if (window.innerWidth <= 1024) closeMobileSidebar();
}

// Logic module DH_CT
async function fetchDHCTData() {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');

    try {
        const token = await getAccessToken();
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.dhctSheetName}!A:P`;

        const response = await fetch(url, {
            headers: { "Authorization": `Bearer ${token}` }
        });

        if (!response.ok) {
            const errData = await response.json();
            throw new Error(errData.error?.message || `HTTP ${response.status}`);
        }

        const result = await response.json();
        if (result.values && result.values.length > 0) {
            dhctData = result.values.slice(1).map((row, idx) => ({
                rowIndex: idx + 2,
                id_dh_ct: (row[0] || '').toString().trim(),
                id_dh: row[1], // Cột B
                ngay: row[2],  // Cột C
                truong: row[3], // Cột D
                ncc: row[4],    // Cột E
                ghi_chu: row[5], // Cột F
                kho: row[6],     // Cột G
                id_sp_ct: row[7], // Cột H
                ten: row[10],     // Cột K
                so_luong: row[11], // Cột L
                gia_nhap: row[12]  // Cột M
            }));
            renderDHCTTable();
            renderUniqueDHCTTable();
        } else {
            document.getElementById('dhctTableBody').innerHTML = '<tr><td colspan="9" class="text-center py-8 text-slate-500">Không có dữ liệu.</td></tr>';
        }
    } catch (err) {
        console.error("Lỗi tải DH_CT:", err);
        alert("Không thể tải dữ liệu từ sheet DH_CT: " + err.message);
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

function renderDHCTTable() {
    const searchInput = document.getElementById('filterDHCTSearch');
    const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';
    const tableBody = document.getElementById('dhctTableBody');
    if (!tableBody) return;

    const filtered = dhctData.filter(item =>
        (item.id_dh && item.id_dh.toString().toLowerCase().includes(searchTerm)) ||
        (item.ten && item.ten.toLowerCase().includes(searchTerm)) ||
        (item.id_sp_ct && item.id_sp_ct.toString().toLowerCase().includes(searchTerm))
    ).sort((a, b) => {
        const da = toYMD(a.ngay);
        const db = toYMD(b.ngay);
        return db.localeCompare(da);
    });

    if (filtered.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="9" class="text-center py-8 text-slate-500">Không tìm thấy kết quả phù hợp.</td></tr>';
        return;
    }

    tableBody.innerHTML = filtered.map(item => `
                <tr class="border-b border-slate-100 hover:bg-slate-50 transition-colors">
                    <td class="px-4 py-3 text-sm font-medium text-blue-600">${item.id_dh || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-600">${item.ngay || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-600">${item.truong || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-600">${item.ncc || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-600 max-w-[200px] truncate" title="${item.ghi_chu || ''}">${item.ghi_chu || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-600">${item.kho || ''}</td>
                    <td class="px-4 py-3 text-sm font-semibold text-slate-900">${item.id_sp_ct || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-700">${item.ten || ''}</td>
                    <td class="px-4 py-3 text-sm text-right font-bold text-slate-900">${(parseFloat(item.so_luong) || 0).toLocaleString('vi-VN')}</td>
                </tr>
            `).join('');
}

let currentUDCTMasterKey = null;
let currentUDCTSelectedItem = null;

function renderUniqueDHCTTable() {
    const tbody = document.getElementById('udctMasterList');
    if (!tbody) return;

    if (!dhctData || dhctData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="2" class="text-center py-10 text-slate-400 italic">Chưa có dữ liệu DH_CT</td></tr>';
        return;
    }

    const grouped = {};
    dhctData.forEach(item => {
        const ngay = (item.ngay || '').toString().trim().split(' ')[0];
        const truong = (item.truong || '').toString().trim();
        const ncc = (item.ncc || '').toString().trim();
        const key = `${ngay}|${truong}|${ncc}`;
        if (!grouped[key]) {
            grouped[key] = { ngay, truong, ncc, count: 0 };
        }
        grouped[key].count++;
    });

    const sortedList = Object.values(grouped).sort((a, b) => {
        const da = toYMD(a.ngay);
        const db = toYMD(b.ngay);
        return db.localeCompare(da) || a.truong.localeCompare(b.truong);
    });

    // Lấy ngày hôm nay định dạng DD/MM/YYYY để check ẩn hiện nút copy
    const now = new Date();
    const d = String(now.getDate()).padStart(2, '0');
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const y = now.getFullYear();
    const todayStr = `${d}/${m}/${y}`;

    let html = '';
    let currentDate = null;

    sortedList.forEach(item => {
        const key = `${item.ngay}|${item.truong}|${item.ncc}`;
        const isKinhDoanh = currentUser && currentUser.role === 'kinhdoanh';
        if (item.ngay !== currentDate) {
            const countForDate = sortedList.filter(x => x.ngay === item.ngay).length;
            html += `
                <tr class="bg-slate-50 border-y border-slate-200">
                    <td colspan="2" class="px-2 py-1.5 text-[10px] font-bold text-slate-500 flex items-center gap-1.5">
                        <span class="text-slate-700">${item.ngay}</span> 
                        <span class="bg-slate-200 px-1 rounded text-[9px] min-w-[14px] text-center text-slate-600">${countForDate}</span>
                    </td>
                </tr>
            `;
            currentDate = item.ngay;
        }

        const isActive = currentUDCTMasterKey === key;

        // Kiểm tra xem đơn hàng này (truong & ncc) đã có trong ngày hôm nay chưa
        const existsToday = sortedList.some(x => x.ngay.includes(todayStr) && x.truong === item.truong && x.ncc === item.ncc);
        const isNotToday = !item.ngay.includes(todayStr);

        html += `
            <tr onclick="selectUDCTMasterRow('${key.replace(/'/g, "\\'")}')" 
                class="border-b border-slate-100 cursor-pointer transition-colors ${isActive ? 'bg-blue-50 text-blue-700' : 'hover:bg-slate-50 text-slate-600'} group">
                <td class="px-3 py-2.5 font-bold uppercase leading-tight text-[10px] w-20">${item.truong}</td>
                <td class="px-3 py-2.5 uppercase leading-tight text-[10px] font-medium flex items-center justify-between">
                    <span>${item.ncc}</span>
                    ${(isNotToday && !existsToday && !isKinhDoanh) ? `
                    <button onclick="event.stopPropagation(); copyDHCTGroup('${key.replace(/'/g, "\\'")}')" 
                            class="p-1 bg-blue-50 text-blue-600 rounded hover:bg-blue-100 transition-colors opacity-0 group-hover:opacity-100" 
                            title="Copy sang hôm nay">
                        <svg class="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7v8a2 2 0 002 2h6M8 7V5a2 2 0 012-2h4.586a1 1 0 01.707.293l4.414 4.414a1 1 0 01.293.707V15a2 2 0 01-2 2h-2M8 7H6a2 2 0 00-2 2v10a2 2 0 002 2h8a2 2 0 002-2v-2" />
                        </svg>
                    </button>` : ''}
                </td>
            </tr>
        `;
    });

    tbody.innerHTML = html;

    // Auto select first row if none selected
    if (!currentUDCTMasterKey && sortedList.length > 0) {
        selectUDCTMasterRow(`${sortedList[0].ngay}|${sortedList[0].truong}|${sortedList[0].ncc}`);
    }
}

async function copyDHCTGroup(key) {
    if (!confirm('Bạn muốn copy đơn hàng này sang ngày hôm nay?')) return;

    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');

    try {
        const parts = key.split('|');
        const oldNgay = parts[0];
        const truong = parts[1];
        const ncc = parts[2];

        // Tìm các dòng thuộc nhóm này
        const itemsToCopy = dhctData.filter(item =>
            (item.ngay || '').toString().trim().split(' ')[0] === oldNgay &&
            (item.truong || '').toString().trim() === truong &&
            (item.ncc || '').toString().trim() === ncc
        );

        if (itemsToCopy.length === 0) {
            alert('Không tìm thấy dữ liệu để copy.');
            return;
        }

        const now = new Date();
        const d = String(now.getDate()).padStart(2, '0');
        const m = String(now.getMonth() + 1).padStart(2, '0');
        const y = now.getFullYear();
        const todayStr = `${d}/${m}/${y}`;

        const orderIdMap = {};

        const newRows = itemsToCopy.map((item, idx) => {
            const key = `${todayStr}|${truong}|${ncc}`;
            if (!orderIdMap[key]) {
                orderIdMap[key] = 'DH-' + (Date.now() + idx).toString().slice(-4);
            }
            const id_dh_ct = 'CT-' + (Date.now() + idx).toString().slice(-6);
            const id_dh = orderIdMap[key];
            const sl = parseFloat(item.so_luong) || 0;
            const gia = parseFloat(item.gia_nhap) || 0;

            // Cấu trúc (16 cột): [id_dh_ct, id_dh, ngay, truong, ncc, ghi_chu, kho, id_sp_ct, _, _, ten, sl, gia, thanh_tien, _, _]
            return [
                id_dh_ct, id_dh, todayStr, truong, ncc, item.ghi_chu || '', item.kho || 'KHO', item.id_sp_ct || '',
                '', '', item.ten || '', sl, gia, sl * gia, '', ''
            ];
        });

        const success = await appendSheetData(CONFIG.dhctSheetName, newRows);
        if (success) {
            await fetchDHCTData(); // Reload từ Sheet
            const newKey = `${todayStr}|${truong}|${ncc}`;
            currentUDCTMasterKey = newKey;
            renderUniqueDHCTTable();
            selectUDCTMasterRow(newKey);
            alert(`Đã copy ${itemsToCopy.length} dòng sang ngày ${todayStr}`);
        }
    } catch (err) {
        console.error('Copy DHCT Group Error:', err);
        alert('Lỗi khi copy: ' + err.message);
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

function getTopUDCTIdSpCts() {
    if (!dhctData) return [];
    const counts = {};
    dhctData.forEach(item => {
        const id = (item.id_sp_ct || '').toString().trim();
        if (id) counts[id] = (counts[id] || 0) + 1;
    });
    return Object.entries(counts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(e => e[0]);
}

function quickFillIdSpCt(id) {
    const input = document.getElementById('row_add_id_sp_ct');
    if (input) {
        input.value = id;
        // Tự động tìm thông tin sản phẩm và điền nốt các cột khác
        if (dsSpCtData && dsSpCtData.length > 0) {
            const searchId = id.toString().trim().toLowerCase();
            const match = dsSpCtData.find(m => (m.id_sp_ct || '').toString().trim().toLowerCase() === searchId);
            if (match) {
                selectSpCtSuggestion(match.id_sp_ct, (match.ten || '').replace(/'/g, "\\'"), match.gia_nhap || 0);
            } else {
                handleIdSpCtInput(input, 'row_add');
            }
        } else {
            handleIdSpCtInput(input, 'row_add');
        }
    }
}

function selectUDCTMasterRow(key) {
    currentUDCTMasterKey = key;

    // Update active UI state for master list
    document.querySelectorAll('#udctMasterList tr').forEach(tr => {
        if (tr.getAttribute('onclick')?.includes(`'${key}'`)) {
            tr.classList.add('bg-blue-50', 'text-blue-700');
            tr.classList.remove('text-slate-600');
        } else if (!tr.classList.contains('bg-slate-50')) {
            tr.classList.remove('bg-blue-50', 'text-blue-700');
            tr.classList.add('text-slate-600');
        }
    });

    const parts = key.split('|');
    renderUDCTSubDetails(parts[0], parts[1], parts[2]);
}

function renderUDCTSubDetails(ngay, truong, ncc) {
    const tbody = document.getElementById('udctSubDetailList');
    if (!tbody) return;

    const filtered = dhctData.filter(item =>
        (item.ngay || '').toString().trim().split(' ')[0] === ngay &&
        (item.truong || '').toString().trim() === truong &&
        (item.ncc || '').toString().trim() === ncc
    );
    const isKinhDoanh = currentUser && currentUser.role === 'kinhdoanh';

    let html = '';

    // Render existing items
    filtered.forEach(item => {
        const isSelected = currentUDCTSelectedItem && currentUDCTSelectedItem.rowIndex === item.rowIndex;
        const sl = parseFloat(item.so_luong) || 0;
        const gia = parseFloat(item.gia_nhap) || 0;
        const thanhTien = sl * gia;

        html += `
            <tr data-id="${item.id_dh_ct}" 
                class="border-b border-slate-100 hover:bg-slate-50 transition-colors ${isSelected ? 'bg-blue-50' : ''}">
                <td class="p-0 text-center">
                    <button onclick="${isKinhDoanh ? '' : `toggleUDCTRowKho('${item.id_dh_ct}', '${item.kho || 'KHO'}')`}"
                            ${isKinhDoanh ? 'style="pointer-events:none"' : ''}
                            class="w-11 h-7 text-[9px] font-bold rounded border-none transition-all ${(item.kho || 'KHO') === 'BH' ? 'bg-amber-100 text-amber-700' : 'bg-blue-100 text-blue-700'}">
                        ${item.kho || 'KHO'}
                    </button>
                </td>
                <td class="p-0 w-24">
                    <input type="text" value="${item.id_sp_ct || ''}" 
                           ${isKinhDoanh ? 'disabled' : ''}
                           oninput="handleIdSpCtInput(this, 'row')"
                           onchange="saveUDCTRowInline('${item.id_dh_ct}', 'id_sp_ct', this.value)"
                           autocomplete="off"
                           class="w-full h-10 px-2 bg-transparent text-[11px] font-bold text-slate-900 border-none outline-none focus:ring-1 focus:ring-blue-400 rounded uppercase disabled:opacity-80">
                </td>
                <td class="p-0">
                    <input type="text" value="${item.ten || ''}" 
                           ${isKinhDoanh ? 'disabled' : ''}
                           onchange="saveUDCTRowInline('${item.id_dh_ct}', 'ten', this.value)"
                           class="w-full h-10 px-2 bg-transparent text-[11px] text-slate-600 border-none outline-none focus:ring-1 focus:ring-blue-400 rounded disabled:opacity-80">
                </td>
                <td class="p-0 w-20">
                    <input type="number" value="${sl}" 
                           ${isKinhDoanh ? 'disabled' : ''}
                           onchange="saveUDCTRowInline('${item.id_dh_ct}', 'so_luong', this.value)"
                           class="w-full h-10 px-2 bg-transparent text-[11px] font-bold text-slate-900 border-none outline-none focus:ring-1 focus:ring-blue-400 rounded text-right disabled:opacity-80">
                </td>
                <td class="p-0 w-24">
                    <input type="number" value="${gia}" 
                           ${isKinhDoanh ? 'disabled' : ''}
                           onchange="saveUDCTRowInline('${item.id_dh_ct}', 'gia_nhap', this.value)"
                           class="w-full h-10 px-2 bg-transparent text-[11px] text-slate-600 border-none outline-none focus:ring-1 focus:ring-blue-400 rounded text-right disabled:opacity-80">
                </td>
                <td class="px-2 py-2.5 text-right font-bold text-rose-600 text-[11px] w-32 relative">
                    <div class="flex items-center justify-end gap-2">
                        <span>${thanhTien.toLocaleString()}</span>
                        ${isKinhDoanh ? '' : `
                        <button onclick="event.stopPropagation(); currentUDCTSelectedItem = dhctData.find(x => x.id_dh_ct === '${item.id_dh_ct}'); deleteUDCTDetail()" 
                                class="text-rose-400 hover:text-rose-600 p-1 transition-all" title="Xóa dòng này">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                        </button>`}
                    </div>
                </td>
            </tr>
        `;
    });

    // Persistent "Add New" row
    if (!isKinhDoanh) {
        html += `
            <tr class="bg-blue-50/20 hover:bg-blue-50 transition-colors border-t border-dashed border-blue-300">
                <td class="p-0 text-center">
                    <button id="row_add_kho_btn" onclick="this.textContent = (this.textContent.trim() === 'KHO' ? 'BH' : 'KHO'); this.classList.toggle('bg-amber-100'); this.classList.toggle('text-amber-700'); this.classList.toggle('bg-blue-100'); this.classList.toggle('text-blue-700');" 
                            class="w-11 h-8 text-[9px] font-bold rounded bg-blue-100 text-blue-700 transition-all">
                        KHO
                    </button>
                </td>
                <td class="p-0 w-24 align-top">
                    <input type="text" id="row_add_id_sp_ct" placeholder="Nhập mã sp..." 
                        oninput="handleIdSpCtInput(this, 'row_add')"
                        onkeydown="if(event.key==='Enter') saveUDCTRowNew()"
                        autocomplete="off"
                        class="w-full h-11 px-2 bg-transparent text-[11px] font-bold text-blue-600 border-none outline-none focus:ring-1 focus:ring-blue-400 rounded uppercase placeholder:text-blue-300">
                    <div class="flex flex-wrap gap-1 px-1 pb-1 overflow-x-auto max-h-20 no-scrollbar">
                    </div>
                </td>
                <td class="p-0">
                    <input type="text" id="row_add_ten" placeholder="Tên sản phẩm..."
                        onkeydown="if(event.key==='Enter') saveUDCTRowNew()"
                        class="w-full h-11 px-2 bg-transparent text-[11px] text-slate-500 border-none outline-none focus:ring-1 focus:ring-blue-400 rounded">
                </td>
                <td class="p-0">
                    <input type="number" id="row_add_sl" value="1"
                        onkeydown="if(event.key==='Enter') saveUDCTRowNew()"
                        class="w-full h-11 px-2 bg-transparent text-[11px] font-bold text-slate-900 border-none outline-none focus:ring-1 focus:ring-blue-400 rounded text-right">
                </td>
                <td class="p-0">
                    <input type="number" id="row_add_gia_nhap" value="0"
                        onkeydown="if(event.key==='Enter') saveUDCTRowNew()"
                        class="w-full h-11 px-2 bg-transparent text-[11px] text-slate-600 border-none outline-none focus:ring-1 focus:ring-blue-400 rounded text-right">
                </td>
                <td class="p-0 text-center">
                </td>
            </tr>
        `;
    }
    tbody.innerHTML = html;
}

async function saveUDCTRowNew() {
    if (!currentUDCTMasterKey) {
        alert('Vui lòng chọn đơn hàng bên trái trước!');
        return;
    }

    const id_sp_ct = document.getElementById('row_add_id_sp_ct').value.trim();
    const ten = document.getElementById('row_add_ten').value.trim();
    const sl = parseFloat(document.getElementById('row_add_sl').value) || 0;
    const gia = parseFloat(document.getElementById('row_add_gia_nhap').value) || 0;
    const kho = document.getElementById('row_add_kho_btn')?.textContent?.trim() || 'KHO';

    if (!id_sp_ct) {
        alert('Vui lòng nhập ID Sản phẩm!');
        document.getElementById('row_add_id_sp_ct').focus();
        return;
    }

    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');

    try {
        const parts = currentUDCTMasterKey.split('|');
        const ngay = parts[0];
        const truong = parts[1];
        const ncc = parts[2];

        const id_dh_ct = 'CT-' + Date.now().toString().slice(-6);
        const id_dh = 'DH-' + Date.now().toString().slice(-4);

        const newRow = [
            id_dh_ct, id_dh, ngay, truong, ncc, '', kho, id_sp_ct,
            '', '', ten, sl, gia, sl * gia, '', ''
        ];

        const success = await appendSheetData(CONFIG.dhctSheetName, [newRow]);
        if (success) {
            await fetchDHCTData(); // Reload from Sheet
            selectUDCTMasterRow(currentUDCTMasterKey);
            // Autofocus back to SKU input for next item
            setTimeout(() => {
                const skuInput = document.getElementById('row_add_id_sp_ct');
                if (skuInput) skuInput.focus();
            }, 100);
        }
    } catch (err) {
        console.error('Save New Inline Row Error:', err);
        alert('Lỗi: ' + err.message);
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

async function handleExcelUploadUniqueDHCT(files) {
    if (currentUser && currentUser.role === 'kinhdoanh') {
        alert('Tài khoản KINHDOANH không được phép tải Excel.');
        return;
    }
    if (!files || files.length === 0) return;
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');

    try {
        const file = files[0];
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { cellDates: true });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

        if (!excelData || excelData.length < 2) {
            alert("File không có dữ liệu!");
            return;
        }

        const headers = excelData[0].map(h => (h || '').toString().toLowerCase().trim());
        const findCol = (names) => headers.findIndex(h => names.some(n => h.includes(n)));

        const colIdx = {
            ngay: findCol(['ngay', 'ngày', 'date']),
            truong: findCol(['truong', 'trường', 'platform']),
            ncc: findCol(['ncc', 'nha cung cap', 'nhà cung cấp']),
            kho: findCol(['kho', 'warehouse']),
            id_sp_ct: findCol(['id_sp_ct', 'sku', 'mã sp', 'ma sp', 'id sp ct']),
            ten: findCol(['ten', 'tên', 'product', 'tên sản phẩm']),
            sl: findCol(['so_luong', 'sl', 'số lượng', 'quantity']),
            gia: findCol(['gia_nhap', 'gia', 'giá', 'giá nhập']),
            ghi_chu: findCol(['ghi_chu', 'ghi chú', 'note'])
        };

        const rows = excelData.slice(1);
        const appendValues = [];
        const now = Date.now();

        let defNgay = '', defTruong = '', defNcc = '';
        if (currentUDCTMasterKey) {
            const parts = currentUDCTMasterKey.split('|');
            defNgay = parts[0];
            defTruong = parts[1];
            defNcc = parts[2];
        } else {
            const d = new Date();
            defNgay = `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
        }

        const orderIdMap = {};
        rows.forEach((row, idx) => {
            const id_sp_ct = colIdx.id_sp_ct !== -1 ? (row[colIdx.id_sp_ct] || '').toString().trim() : '';
            if (!id_sp_ct) return;

            let rawNgay = colIdx.ngay !== -1 ? row[colIdx.ngay] : null;
            let ngay = defNgay;
            if (rawNgay) {
                let val = rawNgay;
                // Nếu là chuỗi số thì chuyển sang số
                if (typeof val === 'string' && val.trim() !== "" && !isNaN(val) && !val.includes('/') && !val.includes('-')) {
                    val = parseFloat(val);
                }

                if (typeof val === 'number') {
                    // Xử lý số serial date từ Excel
                    try {
                        const d = XLSX.SSF.parse_date_code(val);
                        ngay = `${String(d.d).padStart(2, '0')}/${String(d.m).padStart(2, '0')}/${d.y}`;
                    } catch (e) {
                        const dt = new Date(Math.round((val - 25569) * 86400 * 1000));
                        ngay = `${String(dt.getDate()).padStart(2, '0')}/${String(dt.getMonth() + 1).padStart(2, '0')}/${dt.getFullYear()}`;
                    }
                } else if (val instanceof Date) {
                    ngay = `${String(val.getDate()).padStart(2, '0')}/${String(val.getMonth() + 1).padStart(2, '0')}/${val.getFullYear()}`;
                } else {
                    ngay = String(val).trim();
                    if (ngay.includes('-')) {
                        const ps = ngay.split('-');
                        if (ps.length === 3 && ps[0].length === 4) {
                            ngay = `${ps[2]}/${ps[1]}/${ps[0]}`;
                        } else {
                            ngay = ngay.replace(/-/g, '/');
                        }
                    }
                }
            }

            const truong = colIdx.truong !== -1 && row[colIdx.truong] ? (row[colIdx.truong] || '').toString().trim() : defTruong;
            const ncc = colIdx.ncc !== -1 && row[colIdx.ncc] ? (row[colIdx.ncc] || '').toString().trim() : defNcc;

            if (!truong || !ncc) return;

            const kho = colIdx.kho !== -1 && row[colIdx.kho] ? (row[colIdx.kho] || '').toString().trim() : 'KHO';
            let ten = colIdx.ten !== -1 ? (row[colIdx.ten] || '').toString().trim() : '';
            const sl = colIdx.sl !== -1 ? parseFloat(row[colIdx.sl]) || 0 : 1;
            let gia = colIdx.gia !== -1 ? parseFloat(row[colIdx.gia]) || 0 : 0;
            const ghi_chu = colIdx.ghi_chu !== -1 ? (row[colIdx.ghi_chu] || '').toString().trim() : '';

            // Auto-fill Product Info if missing
            if (dsSpCtData && (!ten || !gia)) {
                const sp = dsSpCtData.find(s => (s.id_sp_ct || '').toLowerCase() === id_sp_ct.toLowerCase());
                if (sp) {
                    if (!ten) ten = sp.ten || '';
                    if (!gia) gia = parseFloat(sp.gia_nhap) || 0;
                }
            }

            const key = `${ngay}|${truong}|${ncc}`;
            if (!orderIdMap[key]) {
                orderIdMap[key] = 'DH-' + (now + idx).toString().slice(-4);
            }

            const id_dh_ct = 'CT-' + (now + idx).toString().slice(-6);
            const id_dh = orderIdMap[key];

            // Cấu trúc 16 cột: [id_dh_ct(0), id_dh(1), ngay(2), truong(3), ncc(4), ghi_chu(5), kho(6), id_sp_ct(7), _, _, ten(10), sl(11), gia(12), thanh_tien(13), _, _]
            appendValues.push([
                id_dh_ct, id_dh, ngay, truong, ncc, ghi_chu, kho, id_sp_ct,
                '', '', ten, sl, gia, sl * gia, '', ''
            ]);
        });

        if (appendValues.length === 0) {
            alert("Không tìm thấy dữ liệu hợp lệ (Yêu cầu ít nhất có cột SKU/ID SP CT)");
            return;
        }

        const success = await appendSheetData(CONFIG.dhctSheetName, appendValues);
        if (success) {
            alert(`Đã tải lên thành công ${appendValues.length} sản phẩm!`);
            await fetchDHCTData();
            if (currentUDCTMasterKey) selectUDCTMasterRow(currentUDCTMasterKey);
            else renderUniqueDHCTTable();
        } else {
            alert("Lỗi khi kết nối với Google Sheet.");
        }
    } catch (err) {
        console.error("Excel Upload Sub-Detail Error:", err);
        alert("Lỗi xử lý file Excel: " + err.message);
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
        document.getElementById('excelUploadUniqueDHCT').value = '';
    }
}

function downloadUniqueDHCTTemplate() {
    const headers = [['Ngày (DD/MM/YYYY)', 'Trường', 'NCC', 'Kho', 'ID SP CT', 'Số lượng', 'Giá nhập', 'Tên sản phẩm', 'Ghi chú']];
    const ws = XLSX.utils.aoa_to_sheet(headers);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template_DHCT');
    XLSX.writeFile(wb, 'Mau_Import_DHCT.xlsx');
}

function exportUniqueDHCTToExcel() {
    if (!dhctData || dhctData.length === 0) return alert('Không có dữ liệu để xuất!');

    // Nếu đang chọn một nhóm đơn hàng, chỉ xuất nhóm đó
    let dataToExport = dhctData;
    if (currentUDCTMasterKey) {
        const parts = currentUDCTMasterKey.split('|');
        dataToExport = dhctData.filter(item =>
            (item.ngay || '').toString().trim().split(' ')[0] === parts[0] &&
            (item.truong || '').toString().trim() === parts[1] &&
            (item.ncc || '').toString().trim() === parts[2]
        );
    }

    const headers = ['ID DH CT', 'ID DH', 'Ngày', 'Trường', 'NCC', 'Ghi chú', 'Kho', 'ID SP CT', 'Tên sản phẩm', 'Số lượng', 'Giá nhập', 'Thành tiền'];
    const rows = dataToExport.map(item => [
        item.id_dh_ct, item.id_dh, item.ngay, item.truong, item.ncc, item.ghi_chu,
        item.kho, item.id_sp_ct, item.ten, item.so_luong, item.gia_nhap, (item.so_luong * item.gia_nhap)
    ]);

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'DS_DHCT');
    const fileName = currentUDCTMasterKey ? `DHCT_${currentUDCTMasterKey.replace(/\|/g, '_')}.xlsx` : `Tat_Ca_DHCT.xlsx`;
    XLSX.writeFile(wb, fileName);
}

async function toggleUDCTRowKho(id_dh_ct, currentKho) {
    const nextKho = currentKho === 'KHO' ? 'BH' : 'KHO';
    await saveUDCTRowInline(id_dh_ct, 'kho', nextKho);
    // UI update will happen automatically if saveUDCTRowInline re-renders, 
    // but we can also manually flip its class if needed for immediate feel.
    selectUDCTMasterRow(currentUDCTMasterKey);
}

async function saveUDCTRowInline(id_dh_ct, field, value) {
    const item = dhctData.find(x => x.id_dh_ct === String(id_dh_ct));
    if (!item) return;

    // Update local data
    item[field] = value;

    // Auto-update Details panel if this row is currently selected
    // Auto-update Details panel if it exists (for backward compatibility or other modules)
    if (currentUDCTSelectedItem?.id_dh_ct === id_dh_ct) {
        const inputId = `detail_${field}`;
        const inputEl = document.getElementById(inputId);
        if (inputEl) inputEl.value = value;

        // Re-calculate thanh_tien in details if present
        const ttEl = document.getElementById('detail_thanh_tien');
        if (ttEl) {
            const sl = parseFloat(item.so_luong) || 0;
            const gia = parseFloat(item.gia_nhap) || 0;
            ttEl.textContent = (sl * gia).toLocaleString();
        }
    }

    // Save to server
    try {
        const rowIndex = item.rowIndex;
        // Mapping fields to columns (1-based): kho->G(7), id_sp_ct->H(8), ten->K(11), so_luong->L(12), gia_nhap->M(13)
        const colMap = {
            'kho': 7,
            'id_sp_ct': 8,
            'ten': 11,
            'so_luong': 12,
            'gia_nhap': 13
        };
        const col = colMap[field];
        if (col) {
            await updateSheetCell(CONFIG.dhctSheetName, rowIndex, col, value);
            // Re-render only if needed, but since it's an input, the value is already there.
            // Just update the row's Thanh Tien label
            const sl = parseFloat(item.so_luong) || 0;
            const gia = parseFloat(item.gia_nhap) || 0;
            const tr = document.querySelector(`#udctSubDetailList tr[data-id="${id_dh_ct}"]`);
            if (tr) {
                const ttCell = tr.querySelector('td:last-child');
                if (ttCell) ttCell.textContent = (sl * gia).toLocaleString();
            }
        }
    } catch (err) {
        console.error('Inline save failed:', err);
    }
}

function prepareAddUDCTSubDetail() {
    if (!currentUDCTMasterKey) {
        alert('Vui lòng chọn đơn hàng bên trái trước!');
        return;
    }

    currentUDCTSelectedItem = null; // Clear selection for add mode
    const parts = currentUDCTMasterKey.split('|');

    const content = document.getElementById('udctDetailContent');
    const empty = document.getElementById('udctDetailEmpty');
    const addBtn = document.getElementById('btnSaveNewUDCTItem');

    if (content) content.classList.remove('hidden');
    if (empty) empty.classList.add('hidden');
    if (addBtn) addBtn.classList.remove('hidden'); // Show "Add New" button

    const idDhEl = document.getElementById('detail_id_dh');
    if (idDhEl) idDhEl.textContent = 'Mới...';
    const khoEl = document.getElementById('detail_kho');
    if (khoEl) khoEl.textContent = 'Tự động';

    const idSpEl = document.getElementById('detail_id_sp_ct');
    if (idSpEl) idSpEl.value = '';
    const tenEl = document.getElementById('detail_ten');
    if (tenEl) tenEl.value = '';
    const slEl = document.getElementById('detail_sl');
    if (slEl) slEl.value = 1;
    const giaEl = document.getElementById('detail_gia_nhap');
    if (giaEl) giaEl.value = 0;
    const ttEl = document.getElementById('detail_thanh_tien');
    if (ttEl) ttEl.textContent = '0';

    // Highlight the add row in table (optional redraw)
    renderUDCTSubDetails(parts[0], parts[1], parts[2]);
}

async function saveUDCTDetail() {
    if (!currentUDCTSelectedItem) return; // Không lưu nếu đang ở chế độ thêm mới

    const id_sp_ct = document.getElementById('detail_id_sp_ct')?.value?.trim();
    const ten = document.getElementById('detail_ten')?.value?.trim();
    const sl = parseFloat(document.getElementById('detail_sl')?.value) || 0;
    const gia = parseFloat(document.getElementById('detail_gia_nhap')?.value) || 0;

    // Update local data
    currentUDCTSelectedItem.id_sp_ct = id_sp_ct;
    currentUDCTSelectedItem.ten = ten;
    currentUDCTSelectedItem.so_luong = sl;
    currentUDCTSelectedItem.gia_nhap = gia;

    // Update UI
    const ttEl = document.getElementById('detail_thanh_tien');
    if (ttEl) ttEl.textContent = (sl * gia).toLocaleString();

    if (currentUDCTMasterKey) {
        const p = currentUDCTMasterKey.split('|');
        renderUDCTSubDetails(p[0], p[1], p[2]);
    }

    // Save to server
    try {
        const rowIndex = currentUDCTSelectedItem.rowIndex;
        // Map: H(8), K(11), L(12), M(13) -> index-wise they are: H(7), K(10), L(11), M(12)
        const updates = [
            { row: rowIndex, col: 8, value: id_sp_ct }, // Cột H
            { row: rowIndex, col: 11, value: ten },    // Cột K
            { row: rowIndex, col: 12, value: sl },     // Cột L
            { row: rowIndex, col: 13, value: gia }      // Cột M
        ];

        for (const up of updates) {
            await updateSheetCell(CONFIG.dhctSheetName, up.row, up.col, up.value);
        }
    } catch (err) {
        console.error('Save detail failed:', err);
    }
}

async function addNewUDCTSubDetail() {
    if (!currentUDCTMasterKey) return;

    const parts = currentUDCTMasterKey.split('|');
    const ngay = parts[0];
    const truong = parts[1];
    const ncc = parts[2];

    const id_sp_ct = document.getElementById('detail_id_sp_ct').value.trim();
    const ten = document.getElementById('detail_ten').value.trim();
    const sl = parseFloat(document.getElementById('detail_sl').value) || 0;
    const gia = parseFloat(document.getElementById('detail_gia_nhap').value) || 0;

    if (!id_sp_ct) {
        alert('Vui lòng nhập ID Sản phẩm!');
        return;
    }

    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');

    try {
        const id_dh_ct = 'CT-' + Date.now().toString().slice(-6);
        const id_dh = 'DH-' + Date.now().toString().slice(-4);

        const newRow = [
            id_dh_ct, id_dh, ngay, truong, ncc, '', '', id_sp_ct,
            '', '', ten, sl, gia, sl * gia, '', ''
        ];

        const success = await appendSheetData(CONFIG.dhctSheetName, [newRow]);
        if (success) {
            await fetchDHCTData(); // Reload all
            // Re-select master row to refresh UI
            selectUDCTMasterRow(currentUDCTMasterKey);
        }
    } catch (err) {
        alert('Lỗi thêm sản phẩm: ' + err.message);
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

async function deleteUDCTDetail() {
    if (!currentUDCTSelectedItem) return;
    if (!confirm('Bạn có chắc chắn muốn xóa sản phẩm này khỏi đơn hàng?')) return;

    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');

    try {
        const token = await getAccessToken();
        const sheetId = await fetchSheetMeta(CONFIG.dhctSheetName, token);
        if (sheetId === null) throw new Error('Không lấy được sheetId');

        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}:batchUpdate`;
        const body = {
            requests: [{
                deleteDimension: {
                    range: {
                        sheetId,
                        dimension: 'ROWS',
                        startIndex: currentUDCTSelectedItem.rowIndex - 1,
                        endIndex: currentUDCTSelectedItem.rowIndex
                    }
                }
            }]
        };

        const resp = await fetch(url, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });

        if (resp.ok) {
            dhctData = dhctData.filter(x => x.id_dh_ct !== currentUDCTSelectedItem.id_dh_ct);
            // Re-render
            const masterKey = currentUDCTMasterKey;
            renderUniqueDHCTTable();
            if (masterKey) selectUDCTMasterRow(masterKey);
        } else {
            alert('Lỗi khi xóa dòng trên Google Sheets.');
        }
    } catch (err) {
        console.error('Delete UDCT Detail Error:', err);
        alert('Có lỗi xảy ra khi xóa.');
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

function adjustUDCTDetailSL(delta) {
    const input = document.getElementById('detail_sl');
    if (!input) return;
    let val = parseFloat(input.value) || 0;
    val += delta;
    if (val < 0) val = 0;
    input.value = val;

    const price = currentUDCTSelectedItem ? (parseFloat(currentUDCTSelectedItem.don_gia) || 0) : 0;
    document.getElementById('detail_thanh_tien').textContent = (price * val).toLocaleString();
}

// Giữ lại các hàm cũ đề phòng link ngoài gọi
function openUniqueDHCTModal(ngay, truong, ncc) {
    const key = `${ngay}|${truong}|${ncc}`;
    switchModule('unique_dh_ct');
    selectUDCTMasterRow(key);
}

function closeUniqueDHCTModal() {
    if (!overlay) return;

    document.getElementById('uniqueDHCTModalTitle').textContent = `${ngay} | ${truong} | ${ncc}`;

    const tbody = document.getElementById('uniqueDHCTDetailTableBody');

    const filtered = dhctData.filter(item =>
        (item.ngay ? item.ngay.toString().trim() : '') === ngay &&
        (item.truong ? item.truong.toString().trim() : '') === truong &&
        (item.ncc ? item.ncc.toString().trim() : '') === ncc
    ).sort((a, b) => (b.id_dh_ct || '').localeCompare(a.id_dh_ct || ''));

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center py-6 text-slate-500">Không có dữ liệu chi tiết</td></tr>';
    } else {
        tbody.innerHTML = filtered.map(item => `
                    <tr class="border-b border-slate-100 hover:bg-slate-50">
                        <td class="px-3 py-2 text-sm text-blue-600 font-medium">${item.id_dh || ''}</td>
                        <td class="px-3 py-2 text-sm font-semibold text-slate-900">${item.id_sp_ct || ''}</td>
                        <td class="px-3 py-2 text-sm text-slate-700 max-w-[200px] truncate" title="${item.ten || ''}">${item.ten || ''}</td>
                        <td class="px-3 py-2 text-sm text-right font-bold">${(parseFloat(item.so_luong) || 0).toLocaleString('vi-VN')}</td>
                        <td class="px-3 py-2 text-sm text-slate-500">${item.kho || ''}</td>
                        <td class="px-3 py-2 text-sm text-slate-500 max-w-[150px] truncate" title="${item.ghi_chu || ''}">${item.ghi_chu || ''}</td>
                    </tr>
                `).join('');
    }

    const modal = document.getElementById('uniqueDHCTDetailModal');

    overlay.classList.remove('hidden');
    // Allow slight delay for animation if needed
    setTimeout(() => {
        if (modal) modal.classList.remove('translate-x-full');
    }, 10);
}

function closeUniqueDHCTModal() {
    const overlay = document.getElementById('uniqueDHCTDetailOverlay');
    const modal = document.getElementById('uniqueDHCTDetailModal');

    if (modal) {
        modal.classList.add('translate-x-full');
    }

    if (overlay) {
        setTimeout(() => {
            overlay.classList.add('hidden');
        }, 300);
    }
}

let inventoryCurrentPage = 1;
const INVENTORY_PER_PAGE = 100;
let filteredInventoryData = [];

// Logic module Tồn Kho
async function fetchInventoryData() {
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');

    try {
        const token = await getAccessToken();
        // Tên sheet là TON_KHO - mảng từ A đến K (11 cột)
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.inventorySheetName}!A:K`;

        const response = await fetch(url, {
            headers: { "Authorization": `Bearer ${token}` }
        });

        const result = await response.json();
        if (result.values && result.values.length > 0) {
            // Lưu dữ liệu (bỏ dòng tiêu đề đầu tiên)
            inventoryData = result.values.slice(1);
            inventoryCurrentPage = 1;
            filterInventory();
        } else {
            document.getElementById('inventoryTableBody').innerHTML = '<tr><td colspan="7" class="text-center py-8 text-slate-500">Không có dữ liệu tồn kho.</td></tr>';
        }
    } catch (err) {
        console.error("Lỗi tải tồn kho:", err);
        alert("Không thể tải dữ liệu từ sheet TON_KHO. Hãy kiểm tra tên sheet.");
    } finally {
        loadingOverlay.classList.add('hidden');
    }
}

function renderInventory() {
    const tbody = document.getElementById('inventoryTableBody');
    if (!tbody) return;

    const totalRows = filteredInventoryData.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / INVENTORY_PER_PAGE));
    inventoryCurrentPage = Math.min(inventoryCurrentPage, totalPages);
    if (inventoryCurrentPage < 1) inventoryCurrentPage = 1;

    const start = (inventoryCurrentPage - 1) * INVENTORY_PER_PAGE;
    const paginated = filteredInventoryData.slice(start, start + INVENTORY_PER_PAGE);

    document.getElementById('inventoryStatsText').textContent = `Hiển thị: ${paginated.length}/${totalRows} sản phẩm (Tổng: ${inventoryData.length})`;
    document.getElementById('inventoryPageInfo').textContent = `Trang ${inventoryCurrentPage} / ${totalPages}`;
    document.getElementById('inventoryPrevBtn').disabled = inventoryCurrentPage <= 1;
    document.getElementById('inventoryNextBtn').disabled = inventoryCurrentPage >= totalPages;

    if (paginated.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center py-8 text-slate-500">Không tìm thấy dữ liệu phù hợp.</td></tr>';
        return;
    }

    tbody.innerHTML = paginated.map(row => {
        // Ánh xạ cột dựa trên cấu trúc: id(0), kho(1), id_sp_ct(2), id_sp(3), ten_sp(4), kt1(5), kt2(6), ton_dau(7), nhap(8), xuat(9), ton_cuoi(10)
        const ton_dau = parseFloat(row[7]) || 0;
        const nhap = parseFloat(row[8]) || 0;
        const xuat = parseFloat(row[9]) || 0;
        const ton_cuoi = parseFloat(row[10]) || 0;

        return `
                    <tr class="border-b border-slate-100 hover:bg-slate-50 transition-colors">
                        <td class="px-3 py-2 text-sm text-slate-700">${row[1] || ''}</td>
                        <td class="px-3 py-2 text-sm font-semibold text-slate-900">${row[2] || ''}</td>
                        <td class="px-3 py-2 text-sm text-slate-700">${row[4] || ''}</td>
                        <td class="px-3 py-2 text-sm text-right text-slate-700">${ton_dau.toLocaleString('vi-VN')}</td>
                        <td class="px-3 py-2 text-sm text-right text-emerald-600 font-medium">+${nhap.toLocaleString('vi-VN')}</td>
                        <td class="px-3 py-2 text-sm text-right text-rose-500 font-medium">-${xuat.toLocaleString('vi-VN')}</td>
                        <td class="px-3 py-2 text-sm text-right font-bold text-slate-900">${ton_cuoi.toLocaleString('vi-VN')}</td>
                    </tr>
                `;
    }).join('');
}

function filterInventory() {
    const searchTerm = document.getElementById('inventorySearch').value.toLowerCase().trim();

    filteredInventoryData = inventoryData.filter(row => {
        const idSpCt = (row[2] || '').toLowerCase();
        const tenSp = (row[4] || '').toLowerCase();
        // Tìm theo id_sp_ct hoặc (tên + id_sp_ct)
        return idSpCt.includes(searchTerm) || (tenSp + " " + idSpCt).includes(searchTerm);
    });

    inventoryCurrentPage = 1;
    renderInventory();
}

function changeInventoryPage(delta) {
    inventoryCurrentPage += delta;
    renderInventory();
}

function changeUpmisaDate(step) {
    const dateInput = document.getElementById('upmisaDateFilter');
    if (!dateInput.value) {
        const today = new Date();
        const y = today.getFullYear();
        const m = String(today.getMonth() + 1).padStart(2, '0');
        const d = String(today.getDate()).padStart(2, '0');
        dateInput.value = `${y}-${m}-${d}`;
    } else {
        const parts = dateInput.value.split('-');
        if (parts.length === 3) {
            const currentDate = new Date(parts[0], parts[1] - 1, parts[2]);
            currentDate.setDate(currentDate.getDate() + step);
            const y = currentDate.getFullYear();
            const m = String(currentDate.getMonth() + 1).padStart(2, '0');
            const d = String(currentDate.getDate()).padStart(2, '0');
            dateInput.value = `${y}-${m}-${d}`;
        }
    }
    renderUpmisaTable(1);
}

function applyRoleUI(role) {
    const isDemo = role === 'demo';
    const isKinhDoanh = role === 'kinhdoanh';
    const isRestricted = isDemo || isKinhDoanh;
    const hideIds = ['sidebarUpmisa', 'sidebarBaocao', 'sidebarInventory', 'sidebarDHCT'];
    hideIds.forEach(hid => {
        const el = document.getElementById(hid);
        if (el) el.style.display = isRestricted ? 'none' : '';
    });

    const hangHoanModule = document.getElementById('sidebarHangHoan');
    if (hangHoanModule) hangHoanModule.style.display = isDemo ? 'none' : '';
    const hhShopDienModule = document.getElementById('sidebarHHShopDien');
    if (hhShopDienModule) hhShopDienModule.style.display = isDemo ? 'none' : '';

    if (isKinhDoanh) {
        const hhActions = [
            "button[onclick='exportHangHoanSummaryToExcel()']",
            "button[onclick='exportHangHoanToExcel()']",
            "button[onclick='exportHangHoanToMisa()']"
        ];
        document.querySelectorAll(hhActions.join(', ')).forEach(b => b.style.display = 'none');
    }

    ['excelUploadDonhang', 'excelUpload'].forEach(upId => {
        const excelUp = document.getElementById(upId);
        if (excelUp && excelUp.parentElement) {
            excelUp.parentElement.style.display = isRestricted ? 'none' : '';
        }
    });

    document.querySelectorAll("button[onclick='updateAllPricesBatch()'], button[onclick^='batchUpdateUDCTStatus']").forEach(b => {
        b.style.display = isRestricted ? 'none' : '';
    });

    document.querySelectorAll("button[onclick='exportHangHoanSummaryToExcel()'], button[onclick='exportHangHoanToExcel()'], button[onclick='exportHangHoanToMisa()'], button[onclick='exportReportToExcel()'], button[onclick='exportIdSPExcel()'], button[onclick='exportUpmisaToExcel()']").forEach(b => {
        b.style.display = isKinhDoanh ? 'none' : '';
    });

    const actionHeader = Array.from(document.querySelectorAll('th')).find(th => th.textContent.trim() === 'Action');
    if (actionHeader) actionHeader.style.display = isKinhDoanh ? 'none' : '';

    document.querySelectorAll('#donhangTableBody td:nth-child(16), #donhangTableBody th:nth-child(16)').forEach(el => {
        el.style.display = isKinhDoanh ? 'none' : '';
    });

    const donhangExcelBtn = document.querySelector("label[for='excelUploadDonhang']");
    if (donhangExcelBtn) donhangExcelBtn.style.display = isRestricted ? 'none' : '';

    const donhangRefreshBtn = document.querySelector("button[onclick='loadUDCTData()']");
    if (donhangRefreshBtn) donhangRefreshBtn.style.display = isRestricted ? 'none' : '';

    // Hide Add DH button in moduleUniqueDHCT
    const addUniqueDHBtn = document.querySelector("button[onclick='openAddDHCTModal()']");
    if (addUniqueDHBtn) addUniqueDHBtn.style.display = isKinhDoanh ? 'none' : '';
}

async function handleLogin() {
    const id = document.getElementById('loginId')?.value.trim();
    const password = document.getElementById('loginPassword')?.value;
    const errorEl = document.getElementById('loginError');
    if (errorEl) {
        errorEl.classList.add('hidden');
        errorEl.textContent = '';
    }

    if (!id || !password) {
        if (errorEl) {
            errorEl.textContent = 'Vui lòng nhập đầy đủ tài khoản và mật khẩu.';
            errorEl.classList.remove('hidden');
        }
        return;
    }

    document.getElementById('loginLoading').classList.remove('hidden');
    try {
        await fetchAuthData();
        const user = usersData.find(u => u.id === id && u.password === password);
        if (!user) {
            if (errorEl) {
                errorEl.textContent = 'Sai tài khoản hoặc mật khẩu.';
                errorEl.classList.remove('hidden');
            }
            return;
        }

        if (user.tinhTrang && /nghỉ|khóa|block|inactive/i.test(user.tinhTrang)) {
            if (errorEl) {
                errorEl.textContent = 'Tài khoản đang bị khóa hoặc ngưng hoạt động.';
                errorEl.classList.remove('hidden');
            }
            return;
        }

        isLoggedIn = true;
        currentUser = {
            id: user.id,
            name: user.name || user.id,
            role: (user.role || 'user').toString().trim().toLowerCase(),
            password: user.password
        };
        localStorage.setItem('erp_last_user_id', user.id);
        localStorage.setItem('erp_current_user', JSON.stringify(currentUser));
        updateUserProfileUI();
        document.getElementById('loginScreen').classList.add('hidden');
        document.getElementById('mainApp').classList.remove('hidden');
        document.getElementById('welcomeUserName').textContent = currentUser.name;
        applyRoleUI(currentUser.role);

        await loadUDCTData();
        await loadSanphamData();
        await loadDsSpCtData();

        const urlParams = new URLSearchParams(window.location.search);
        const link = urlParams.get('link');
        const moduleMapping = {
            'don_chi_tiet': 'donhang',
            'san_pham': 'sanpham',
            'bao_cao': 'baocao',
            'upmisa': 'upmisa',
            'ton_kho': 'inventory',
            'dh_ct': 'dh_ct',
            'hang_hoan': 'hang_hoan',
            'hh_shop_dien': 'hh_shop_dien'
        };
        if (link && moduleMapping[link]) switchModule(moduleMapping[link]);
    } catch (err) {
        console.error('Login error:', err);
        if (errorEl) {
            errorEl.textContent = 'Không thể xác thực tài khoản. Vui lòng thử lại.';
            errorEl.classList.remove('hidden');
        }
    } finally {
        document.getElementById('loginLoading').classList.add('hidden');
    }
}

function handleLogout() {
    if (confirm("Bạn có muốn đăng xuất không?")) {
        localStorage.removeItem('erp_current_user');
        isLoggedIn = false;
        currentUser = null;
        window.location.reload();
    }
}

function updateUserProfileUI() {
    if (!currentUser) return;
    const sidebarUserName = document.getElementById('sidebarUserName');
    const userInitial = document.getElementById('userInitial');
    const roleLabel = document.querySelector('.sidebar-text .text-xs.text-slate-500');
    if (sidebarUserName) sidebarUserName.textContent = currentUser.name;
    if (userInitial && currentUser.name) userInitial.textContent = currentUser.name.charAt(0).toUpperCase();
    if (roleLabel) roleLabel.textContent = currentUser.role === 'kinhdoanh' ? 'Kinh doanh' : (currentUser.role || 'Nhân viên');
}

window.onload = async () => {
    const loginScreen = document.getElementById('loginScreen');
    const mainApp = document.getElementById('mainApp');
    const lastId = localStorage.getItem('erp_last_user_id');
    if (lastId && document.getElementById('loginId')) document.getElementById('loginId').value = lastId;

    try {
        const savedUser = localStorage.getItem('erp_current_user');
        if (savedUser) {
            const parsedUser = JSON.parse(savedUser);
            if (parsedUser && parsedUser.id) {
                // Immediate UI update for smooth transition
                if (loginScreen) loginScreen.classList.add('hidden');
                if (mainApp) mainApp.classList.remove('hidden');
                isLoggedIn = true;
                currentUser = parsedUser;
                if (document.getElementById('welcomeUserName')) document.getElementById('welcomeUserName').textContent = currentUser.name;
                updateUserProfileUI();
                applyRoleUI(currentUser.role);

                // Khởi động tải dữ liệu ngay từ đầu để tránh lỗi "Chưa có dữ liệu"
                loadUDCTData();
                loadSanphamData();
                loadDsSpCtData();
                fetchDHCTData(); // Thêm luôn DH_CT mới


                // Chuyển nhanh qua module đang yêu cầu (nếu có)
                const urlParams = new URLSearchParams(window.location.search);
                const link = urlParams.get('link');
                const moduleMapping = {
                    'don_chi_tiet': 'donhang',
                    'san_pham': 'sanpham',
                    'bao_cao': 'baocao',
                    'upmisa': 'upmisa',
                    'ton_kho': 'inventory',
                    'dh_ct': 'dh_ct',
                    'hang_hoan': 'hang_hoan',
                    'hh_shop_dien': 'hh_shop_dien'
                };
                if (link && moduleMapping[link]) {
                    if (typeof switchModule === 'function') switchModule(moduleMapping[link]);
                }

                // Fetch data in background
                await fetchAuthData();
                const user = usersData.find(u => u.id === parsedUser.id && u.password === parsedUser.password);
                if (user && !(user.tinhTrang && /nghỉ|khóa|block|inactive/i.test(user.tinhTrang))) {
                    currentUser = {
                        id: user.id,
                        name: user.name || user.id,
                        role: (user.role || 'user').toString().trim().toLowerCase(),
                        password: user.password
                    };
                    updateUserProfileUI();
                    applyRoleUI(currentUser.role);
                    await loadUDCTData();
                    await loadSanphamData();
                    await loadDsSpCtData();
                } else {
                    handleLogout(); // Session invalid
                }
            } else {
                if (loginScreen) loginScreen.classList.remove('hidden');
                if (mainApp) mainApp.classList.add('hidden');
            }
        } else {
            if (loginScreen) loginScreen.classList.remove('hidden');
            if (mainApp) mainApp.classList.add('hidden');
        }
    } catch (e) {
        console.warn('Auto login failed:', e);
        if (loginScreen) loginScreen.classList.remove('hidden');
        if (mainApp) mainApp.classList.add('hidden');
    }

    if (window.innerWidth > 1024) closeMobileSidebar();
};

window.addEventListener('resize', () => {
    if (window.innerWidth > 1024) closeMobileSidebar();
});

// LOGIC THÊM ĐƠN HÀNG MỚI (MODAL)
function openAddDHCTModal() {
    const modal = document.getElementById('addDHCTModal');
    if (modal) {
        modal.classList.remove('hidden');
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('addDHCTNgay').value = today;
        document.getElementById('addDHCTTruong').value = '';
        document.getElementById('addDHCTNcc').value = '';

        // Set default Type to NHẬP
        const nhapBtn = Array.from(document.querySelectorAll('#addDHCTTruongBtns button')).find(b => b.textContent.trim() === 'NHẬP');
        if (nhapBtn) {
            selectAddDHCTTruong(nhapBtn, 'NHẬP');
        } else {
            document.getElementById('addDHCTTruong').value = 'NHẬP';
        }

        // Suggesst NCC buttons from existing data
        const nccSet = new Set();
        dhctData.forEach(item => { if (item.ncc) nccSet.add(item.ncc.trim()); });
        const distinctNccs = Array.from(nccSet).sort().slice(0, 20); // Top 20 alphabetic
        const nccBox = document.getElementById('addDHCTNccSuggestions');
        if (nccBox) {
            nccBox.innerHTML = distinctNccs.map(ncc => `
                <button onclick="document.getElementById('addDHCTNcc').value='${ncc.replace(/'/g, "\\'")}'"
                    class="px-2 py-1 bg-slate-50 border border-slate-200 hover:border-primary hover:text-primary rounded-lg text-[10px] font-medium text-slate-500 transition-all">
                    ${ncc}
                </button>
            `).join('');
        }
    }
}

function selectAddDHCTTruong(btn, value) {
    document.getElementById('addDHCTTruong').value = value;
    // UI Feedback
    document.querySelectorAll('#addDHCTTruongBtns button').forEach(b => {
        b.classList.remove('bg-emerald-500', 'bg-rose-500', 'text-white', 'border-transparent');
        b.classList.add('text-slate-600', 'border-slate-200');
    });

    if (value === 'NHẬP') {
        btn.classList.add('bg-emerald-500', 'text-white', 'border-transparent');
        btn.classList.remove('text-slate-600', 'border-slate-200');
    } else {
        btn.classList.add('bg-rose-500', 'text-white', 'border-transparent');
        btn.classList.remove('text-slate-600', 'border-slate-200');
    }
}

function closeAddDHCTModal() {
    const modal = document.getElementById('addDHCTModal');
    if (modal) modal.classList.add('hidden');
}

async function saveNewDHCT() {
    const ngay = document.getElementById('addDHCTNgay').value;
    const truong = document.getElementById('addDHCTTruong').value.trim();
    const ncc = document.getElementById('addDHCTNcc').value.trim();

    if (!ngay || !truong || !ncc) {
        alert('Vui lòng nhập đầy đủ Ngày, Trường và NCC!');
        return;
    }

    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');

    try {
        const id_dh_ct = 'CT-' + Date.now().toString().slice(-6);
        const id_dh = 'DH-' + Date.now().toString().slice(-4);

        const [y, m, d] = ngay.split('-');
        const dateFormatted = `${d}/${m}/${y}`;

        // Header (12 cột): id_dh_ct(0), id_dh(1), ngay(2), truong(3), ncc(4)
        const newRow = [
            id_dh_ct,
            id_dh,
            dateFormatted,
            truong,
            ncc,
            '', '', '', '', '', '', '', '', '', '', ''
        ];

        const success = await appendSheetData(CONFIG.dhctSheetName, [newRow]);
        if (success) {
            closeAddDHCTModal();
            // Load lại dữ liệu để Master List cập nhật
            await fetchDHCTData();
        } else {
            throw new Error('Không thể ghi dữ liệu vào Google Sheets.');
        }
    } catch (err) {
        console.error('Save New DHCT Error:', err);
        alert('Có lỗi xảy ra khi lưu: ' + err.message);
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

// --- LOGIC GỢI Ý ID_SP_CT ---
let suggestionTarget = null;
let suggestionType = 'row';

function handleIdSpCtInput(input, type) {
    suggestionTarget = input;
    suggestionType = type;
    const val = input.value.trim().toLowerCase();
    const sugBox = document.getElementById('spctSuggestions');

    if (!val || !dsSpCtData || dsSpCtData.length === 0) {
        if (sugBox) sugBox.classList.add('hidden');
        return;
    }

    const matches = dsSpCtData.filter(item =>
        (item.id_sp_ct || '').toLowerCase().includes(val) ||
        (item.ten || '').toLowerCase().includes(val)
    ).slice(0, 30);

    if (matches.length === 0) {
        if (sugBox) sugBox.classList.add('hidden');
        return;
    }

    const rect = input.getBoundingClientRect();
    sugBox.style.left = `${rect.left}px`;
    sugBox.style.top = `${rect.bottom + window.scrollY}px`;
    sugBox.style.width = `${Math.max(rect.width, 320)}px`;
    sugBox.classList.remove('hidden');

    sugBox.innerHTML = matches.map(m => `
        <div onclick="selectSpCtSuggestion('${m.id_sp_ct}', '${(m.ten || '').replace(/'/g, "\\'")}', '${m.gia_nhap || 0}')" 
             class="px-3 py-2 hover:bg-blue-50 cursor-pointer border-b border-slate-50 last:border-0 group transition-colors">
            <div class="font-bold text-primary group-hover:text-blue-700">${m.id_sp_ct}</div>
            <div class="text-[10px] text-slate-500 truncate">${m.ten}</div>
            <div class="text-[9px] text-slate-400 italic">Giá: ${(parseFloat(m.gia_nhap) || 0).toLocaleString()}</div>
        </div>
    `).join('');
}

function selectSpCtSuggestion(id, ten, gia) {
    if (!suggestionTarget) return;

    suggestionTarget.value = id;
    const sugBox = document.getElementById('spctSuggestions');
    if (sugBox) sugBox.classList.add('hidden');

    if (suggestionType === 'detail') {
        const tenInput = document.getElementById('detail_ten');
        const giaInput = document.getElementById('detail_gia_nhap');
        if (tenInput) tenInput.value = ten;
        if (giaInput) giaInput.value = gia || 0;
        // Trigger auto-save
        saveUDCTDetail();
    } else if (suggestionType === 'row_add') {
        const tenInput = document.getElementById('row_add_ten');
        const giaInput = document.getElementById('row_add_gia_nhap');
        if (tenInput) tenInput.value = ten;
        if (giaInput) giaInput.value = gia || 0;
        // Auto-save immediately for rapid entry
        saveUDCTRowNew();
    } else {
        // Row mode (existing item)
        const tr = suggestionTarget.closest('tr');
        if (tr) {
            const idRow = tr.getAttribute('data-id');
            const tenInput = tr.querySelector('input[onchange*="\'ten\'"]');
            const giaInput = tr.querySelector('input[onchange*="\'gia_nhap\'"]');
            if (tenInput) tenInput.value = ten;
            if (giaInput) giaInput.value = gia || 0;

            // Trigger inline save for ALL fields
            saveUDCTRowInline(idRow, 'id_sp_ct', id);
            saveUDCTRowInline(idRow, 'ten', ten);
            saveUDCTRowInline(idRow, 'gia_nhap', gia || 0);
        }
    }
}

document.addEventListener('mousedown', (e) => {
    const sugBox = document.getElementById('spctSuggestions');
    if (sugBox && !sugBox.contains(e.target) && e.target !== suggestionTarget) {
        sugBox.classList.add('hidden');
    }
});
