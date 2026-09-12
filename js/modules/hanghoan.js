// hanghoan - Module Pattern (IIFE)
(function () {
function formatDateTimeNow() {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, '0');
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const y = now.getFullYear();
    const h = String(now.getHours()).padStart(2, '0');
    const min = String(now.getMinutes()).padStart(2, '0');
    const s = String(now.getSeconds()).padStart(2, '0');
    return `${d}/${m}/${y} ${h}:${min}:${s}`;
}

function getEditorDisplayName() {
    if (typeof currentUser !== 'undefined' && currentUser) {
        if (currentUser.name && currentUser.id && currentUser.name !== currentUser.id) {
            return `${currentUser.name} (${currentUser.id})`;
        }
        return currentUser.name || currentUser.id || currentUser.role || 'User';
    }
    return 'User';
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

function stepHhEditSLG(step) {
    const input = document.getElementById('hhEditSLG');
    if (!input) return;
    let currentVal = parseInt(input.value, 10);
    if (isNaN(currentVal) || currentVal < 1) currentVal = 1;
    currentVal = Math.max(1, currentVal + step);
    input.value = currentVal;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
}

function stepHhEditNgayNhan(step) {
    const input = document.getElementById('hhEditNgayNhan');
    if (!input) return;
    if (!input.value) {
        const now = new Date();
        const y = now.getFullYear();
        const m = String(now.getMonth() + 1).padStart(2, '0');
        const d = String(now.getDate()).padStart(2, '0');
        input.value = `${y}-${m}-${d}`;
        return;
    }
    const parts = input.value.split('-');
    if (parts.length !== 3) return;
    const dt = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    dt.setDate(dt.getDate() + step);
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    input.value = `${y}-${m}-${d}`;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
}

function shiftHHFilterDate(idOrWhich, delta) {
    let input = document.getElementById(idOrWhich);
    if (!input) {
        if (idOrWhich === 'from') input = document.getElementById('filterHHFrom');
        else if (idOrWhich === 'to') input = document.getElementById('filterHHTo');
    }
    if (!input) return;

    if (!input.value) {
        const today = new Date();
        const y = today.getFullYear();
        const m = String(today.getMonth() + 1).padStart(2, '0');
        const d = String(today.getDate()).padStart(2, '0');
        input.value = `${y}-${m}-${d}`;
    }

    const parts = input.value.split('-');
    if (parts.length !== 3) return;
    const dt = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    dt.setDate(dt.getDate() + delta);
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    input.value = `${y}-${m}-${d}`;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    filterHangHoanData();
}

function changeHangHoanDate(which, step) {
    shiftHHFilterDate(which, step);
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


function renderHhKhoButtons(value) {
    const current = (value || '').toString().trim().toUpperCase() || 'KHO';
    const buttons = document.querySelectorAll('#hhEditKhoButtons button');
    buttons.forEach(btn => {
        const active = btn.textContent.trim().toUpperCase() === current;
        btn.classList.toggle('bg-white', active);
        btn.classList.toggle('text-slate-800', active);
        btn.classList.toggle('shadow-sm', active);
        btn.classList.toggle('text-slate-500', !active);
    });
}

function setHhKho(value) {
    const normalized = (value || 'KHO').toString().trim().toUpperCase();
    document.getElementById('hhEditKho').value = normalized;
    renderHhKhoButtons(normalized);
}

function populateHhFormOptions() {
    const maGianList = document.getElementById('hhMaGianList');
    if (maGianList) {
        const uniqueMaGian = [...new Set(hangHoanData.map(i => (i.ma_gian || '').toString().trim()).filter(Boolean))].sort();
        maGianList.innerHTML = uniqueMaGian.map(v => `<option value="${escapeHtml(v)}">`).join('');
    }
}

function ensureHhCatalogLoaded(callback) {
    if (sanphamData && sanphamData.length > 0) {
        if (typeof callback === 'function') callback();
        return;
    }

    if (!hhCatalogLoadPromise) {
        hhCatalogLoadPromise = typeof loadSanphamData === 'function' ? loadSanphamData() : Promise.resolve().finally(() => {
            hhCatalogLoadPromise = null;
        });
    }

    hhCatalogLoadPromise.then(() => {
        if (typeof callback === 'function') callback();
    }).catch(error => {
        console.error('HH catalog reload error:', error);
    });
}

function getHhSkuMatches(value, limit = 50) {
    const search = (value || '').toString().trim().toLowerCase();
    const seen = new Set();
    const matches = [];

    sanphamData.forEach(item => {
        const sku = (item.id_sp || '').toString().trim();
        if (!sku || seen.has(sku)) return;

        const skuLower = sku.toLowerCase();
        const nameLower = (item.ten_sp || '').toString().toLowerCase();
        if (!search || skuLower.includes(search) || nameLower.includes(search)) {
            seen.add(sku);
            matches.push({ sku, ten: item.ten_sp || '' });
        }
    });

    return matches.sort((a, b) => a.sku.localeCompare(b.sku)).slice(0, limit);
}

function renderHhSkuSuggestions(forceShow = false) { return;
    const input = document.getElementById('hhEditSKU');
    const sugBox = document.getElementById('hhSkuSuggestions');
    if (!input || !sugBox) return;

    const value = input.value || '';
    if ((!sanphamData || sanphamData.length === 0) && (forceShow || value.trim())) {
        sugBox.innerHTML = '<div class="suggestion-item"><span class="item-name">Đang tải danh mục SKU...</span></div>';
        sugBox.classList.remove('hidden');
        ensureHhCatalogLoaded(() => renderHhSkuSuggestions(forceShow));
        return;
    }

    const suggestions = getHhSkuMatches(value);

    if ((forceShow || value.trim()) && suggestions.length > 0) {
        sugBox.innerHTML = suggestions.map(item => `
            <div class="suggestion-item" onclick="setHhSku('${escapeHtml(item.sku)}')">
                <span class="item-code">${escapeHtml(item.sku)}</span>
                <span class="item-name">${escapeHtml(item.ten_sp)}</span>
            </div>
        `).join('');
        sugBox.classList.remove('hidden');
    } else {
        sugBox.innerHTML = '';
        sugBox.classList.add('hidden');
    }
}

let hhUdctLoadPromise = null;
let hhCurrentMvd2Suggestions = [];
let hhCurrentMdhSuggestions = [];

function ensureHhUdctLoaded(callback) {
    if (udctData && udctData.length > 0) {
        if (typeof callback === 'function') callback();
        return;
    }
    if (!hhUdctLoadPromise) {
        hhUdctLoadPromise = typeof loadUDCTData === 'function' ? loadUDCTData(true) : Promise.resolve();
    }
    hhUdctLoadPromise.then(() => {
        hhUdctLoadPromise = null;
        if (typeof callback === 'function') callback();
    }).catch(err => {
        hhUdctLoadPromise = null;
        console.error('HH UDCT load error:', err);
    });
}

function renderUdctSuggestionItem(item, idx, type) {
    const fnName = type === 'mvd2' ? 'selectHhUdctForMvd2ByIndex' : 'selectHhUdctForMdhByIndex';
    const mainCode = type === 'mvd2' ? (item.mvd || item.mdh || '-') : (item.mdh || item.mvd || '-');
    const subCode = type === 'mvd2' ? (item.mdh ? `MDH: ${escapeHtml(item.mdh)}` : '') : (item.mvd ? `MVD: ${escapeHtml(item.mvd)}` : '');
    const slg = item.slg_xuat || item.so_luong || '1';
    const skuDisplay = item.id_sp_ct || item.id_sp || '';

    return `
        <div class="px-3 py-2 border-b border-slate-100 last:border-0 hover:bg-blue-50/70 active:bg-blue-100/70 cursor-pointer transition-colors"
            onmousedown="event.preventDefault(); ${fnName}(${idx});"
            onclick="${fnName}(${idx});">
            <div class="flex items-center justify-between gap-2">
                <div class="flex items-center gap-1.5 min-w-0">
                    <span class="font-bold text-slate-900 text-xs truncate">${escapeHtml(mainCode)}</span>
                    ${subCode ? `<span class="text-[10px] text-blue-600 bg-blue-50 border border-blue-200/60 px-1.5 py-0.5 rounded font-medium shrink-0">${subCode}</span>` : ''}
                </div>
                <span class="text-[11px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200/60 px-1.5 py-0.5 rounded shrink-0">SL: ${escapeHtml(slg)}</span>
            </div>
            <div class="flex items-center gap-1.5 text-[11px] text-slate-500 mt-1 truncate">
                ${item.ma_gian ? `<span class="font-semibold text-slate-700 shrink-0">${escapeHtml(item.ma_gian)}</span>` : ''}
                ${item.ma_gian && skuDisplay ? `<span class="text-slate-300">•</span>` : ''}
                ${skuDisplay ? `<span class="text-indigo-600 font-medium shrink-0">${escapeHtml(skuDisplay)}</span>` : ''}
                ${item.ten_sp ? `<span class="text-slate-300">•</span><span class="truncate text-slate-600" title="${escapeHtml(item.ten_sp)}">${escapeHtml(item.ten_sp)}</span>` : ''}
            </div>
        </div>
    `;
}

function applyUdctItemToHangHoanForm(item, currentField) {
    if (!item) return;

    const isKinhDoanh = currentUser && currentUser.role === 'kinhdoanh';
    if (hhDrawerMode === 'edit' && isKinhDoanh) {
        if (currentField === 'mvd2') {
            const mvd2Input = document.getElementById('hhEditMVD2');
            if (mvd2Input) mvd2Input.value = item.mvd || '';
            const mdhInput = document.getElementById('hhEditMDH');
            if (mdhInput && !mdhInput.value && item.mdh) mdhInput.value = item.mdh;
        } else {
            const mdhInput = document.getElementById('hhEditMDH');
            if (mdhInput) mdhInput.value = item.mdh || '';
            const mvd2Input = document.getElementById('hhEditMVD2');
            if (mvd2Input && !mvd2Input.value && item.mvd) mvd2Input.value = item.mvd;
        }
        const mvd2Sug = document.getElementById('hhMvd2Suggestions');
        if (mvd2Sug) mvd2Sug.classList.add('hidden');
        const mdhSug = document.getElementById('hhMdhSuggestions');
        if (mdhSug) mdhSug.classList.add('hidden');
        return;
    }

    if (currentField === 'mvd2') {
        const mvd2Input = document.getElementById('hhEditMVD2');
        if (mvd2Input) mvd2Input.value = item.mvd || '';
        const mdhInput = document.getElementById('hhEditMDH');
        if (mdhInput && !mdhInput.value && item.mdh) mdhInput.value = item.mdh;
    } else {
        const mdhInput = document.getElementById('hhEditMDH');
        if (mdhInput) mdhInput.value = item.mdh || '';
        const mvd2Input = document.getElementById('hhEditMVD2');
        if (mvd2Input && !mvd2Input.value && item.mvd) mvd2Input.value = item.mvd;
    }

    const maGianInput = document.getElementById('hhEditMaGian');
    if (maGianInput && item.ma_gian) {
        maGianInput.value = item.ma_gian;
    }

    const skuCt = (item.id_sp_ct || item.sku_shop_up || '').toString().trim();
    const skuCtInput = document.getElementById('hhEditSKUCT');
    if (skuCtInput) {
        skuCtInput.value = skuCt;
    }

    let skuVal = (item.id_sp || '').toString().trim();
    let tenSpVal = item.ten_sp || '';
    if (skuCt && typeof sanphamData !== 'undefined' && sanphamData.length) {
        const matchedSp = sanphamData.find(i => (i.sku_con || '').toString().trim().toUpperCase() === skuCt.toUpperCase());
        if (matchedSp) {
            skuVal = matchedSp.id_sp || matchedSp.sku_con.substring(0, 4) || skuVal;
            tenSpVal = matchedSp.ten_sp || matchedSp.ten || tenSpVal;
        }
    }

    const skuInput = document.getElementById('hhEditSKU');
    if (skuInput) {
        skuInput.value = skuVal;
    }

    const slgInput = document.getElementById('hhEditSLG');
    if (slgInput) {
        slgInput.value = item.slg_xuat || item.so_luong || '1';
    }

    const tenSpInput = document.getElementById('hhEditTenSP');
    if (tenSpInput && tenSpVal) {
        tenSpInput.value = tenSpVal;
    }

    const mvd2Sug = document.getElementById('hhMvd2Suggestions');
    if (mvd2Sug) mvd2Sug.classList.add('hidden');
    const mdhSug = document.getElementById('hhMdhSuggestions');
    if (mdhSug) mdhSug.classList.add('hidden');
    const skuCtSug = document.getElementById('hhSkuCtSuggestions');
    if (skuCtSug) skuCtSug.classList.add('hidden');
}

function selectHhUdctForMvd2ByIndex(index) {
    const item = hhCurrentMvd2Suggestions[index];
    applyUdctItemToHangHoanForm(item, 'mvd2');
}

function selectHhUdctForMdhByIndex(index) {
    const item = hhCurrentMdhSuggestions[index];
    applyUdctItemToHangHoanForm(item, 'mdh');
}

function setHhMvd2(value) {
    const input = document.getElementById('hhEditMVD2');
    if (input) {
        input.value = value;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }
    const sugBox = document.getElementById('hhMvd2Suggestions');
    if (sugBox) sugBox.classList.add('hidden');
}

function handleHhMvd2Change(forceShow = false) {
    const input = document.getElementById('hhEditMVD2');
    if (!input) return;
    const val = input.value.trim().toLowerCase();
    const sugBox = document.getElementById('hhMvd2Suggestions');
    const maGianCurrent = (document.getElementById('hhEditMaGian')?.value || '').trim().toLowerCase();

    if ((!udctData || udctData.length === 0) && (val || forceShow)) {
        if (sugBox) {
            sugBox.innerHTML = '<div class="px-3 py-2 text-xs text-slate-500">Đang tải dữ liệu UD_CT...</div>';
            sugBox.classList.remove('hidden');
        }
        ensureHhUdctLoaded(() => handleHhMvd2Change(forceShow));
        return;
    }

    if (sugBox) {
        if (val.length >= 1 || forceShow) {
            const seen = new Set();
            const matchesWithMaGian = [];
            const otherMatches = [];

            for (const item of udctData) {
                const mvd = (item.mvd || '').toString().trim();
                const mdh = (item.mdh || '').toString().trim();
                const maGian = (item.ma_gian || '').toString().trim();
                const skuCt = (item.id_sp_ct || '').toString().trim();
                const tenSp = (item.ten_sp || '').toString().trim();
                const key = `${mvd}|${skuCt}|${mdh}`.toUpperCase();
                if ((!mvd && !mdh) || seen.has(key)) continue;

                const isMatchText = !val || mvd.toLowerCase().includes(val) || mdh.toLowerCase().includes(val) || maGian.toLowerCase().includes(val) || skuCt.toLowerCase().includes(val) || tenSp.toLowerCase().includes(val);
                if (isMatchText) {
                    seen.add(key);
                    if (maGianCurrent && maGian.toLowerCase() === maGianCurrent) {
                        matchesWithMaGian.push(item);
                    } else {
                        otherMatches.push(item);
                    }
                }
            }

            let matches = [];
            if (maGianCurrent) {
                matchesWithMaGian.sort((a, b) => (b.mvd || b.mdh || '').toString().localeCompare((a.mvd || a.mdh || '').toString()));
                matches = matchesWithMaGian;
                if (matches.length === 0 && val) {
                    otherMatches.sort((a, b) => (b.mvd || b.mdh || '').toString().localeCompare((a.mvd || a.mdh || '').toString()));
                    matches = otherMatches;
                }
            } else {
                otherMatches.sort((a, b) => (b.mvd || b.mdh || '').toString().localeCompare((a.mvd || a.mdh || '').toString()));
                matches = otherMatches;
            }

            hhCurrentMvd2Suggestions = matches.slice(0, 15);

            if (hhCurrentMvd2Suggestions.length > 0) {
                sugBox.innerHTML = hhCurrentMvd2Suggestions.map((item, idx) => renderUdctSuggestionItem(item, idx, 'mvd2')).join('');
                sugBox.classList.remove('hidden');
            } else {
                sugBox.innerHTML = '<div class="px-3 py-2 text-xs text-slate-400 italic">Không tìm thấy MVD' + (maGianCurrent ? ` thuộc gian [${maGianCurrent.toUpperCase()}]` : '') + '</div>';
                sugBox.classList.remove('hidden');
            }
        } else {
            sugBox.innerHTML = '';
            sugBox.classList.add('hidden');
        }
    }
}

function setHhMdh(value) {
    const input = document.getElementById('hhEditMDH');
    if (input) {
        input.value = value;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }
    const sugBox = document.getElementById('hhMdhSuggestions');
    if (sugBox) sugBox.classList.add('hidden');
}

function handleHhMdhChange(forceShow = false) {
    const input = document.getElementById('hhEditMDH');
    if (!input) return;
    const val = input.value.trim().toLowerCase();
    const sugBox = document.getElementById('hhMdhSuggestions');
    const maGianCurrent = (document.getElementById('hhEditMaGian')?.value || '').trim().toLowerCase();

    if ((!udctData || udctData.length === 0) && (val || forceShow)) {
        if (sugBox) {
            sugBox.innerHTML = '<div class="px-3 py-2 text-xs text-slate-500">Đang tải dữ liệu UD_CT...</div>';
            sugBox.classList.remove('hidden');
        }
        ensureHhUdctLoaded(() => handleHhMdhChange(forceShow));
        return;
    }

    if (sugBox) {
        if (val.length >= 1 || forceShow) {
            const seen = new Set();
            const matchesWithMaGian = [];
            const otherMatches = [];

            for (const item of udctData) {
                const mdh = (item.mdh || '').toString().trim();
                const mvd = (item.mvd || '').toString().trim();
                const maGian = (item.ma_gian || '').toString().trim();
                const skuCt = (item.id_sp_ct || '').toString().trim();
                const tenSp = (item.ten_sp || '').toString().trim();
                const key = `${mdh}|${skuCt}|${mvd}`.toUpperCase();
                if ((!mdh && !mvd) || seen.has(key)) continue;

                const isMatchText = !val || mdh.toLowerCase().includes(val) || mvd.toLowerCase().includes(val) || maGian.toLowerCase().includes(val) || skuCt.toLowerCase().includes(val) || tenSp.toLowerCase().includes(val);
                if (isMatchText) {
                    seen.add(key);
                    if (maGianCurrent && maGian.toLowerCase() === maGianCurrent) {
                        matchesWithMaGian.push(item);
                    } else {
                        otherMatches.push(item);
                    }
                }
            }

            let matches = [];
            if (maGianCurrent) {
                matchesWithMaGian.sort((a, b) => (b.mdh || b.mvd || '').toString().localeCompare((a.mdh || a.mvd || '').toString()));
                matches = matchesWithMaGian;
                if (matches.length === 0 && val) {
                    otherMatches.sort((a, b) => (b.mdh || b.mvd || '').toString().localeCompare((a.mdh || a.mvd || '').toString()));
                    matches = otherMatches;
                }
            } else {
                otherMatches.sort((a, b) => (b.mdh || b.mvd || '').toString().localeCompare((a.mdh || a.mvd || '').toString()));
                matches = otherMatches;
            }

            hhCurrentMdhSuggestions = matches.slice(0, 15);

            if (hhCurrentMdhSuggestions.length > 0) {
                sugBox.innerHTML = hhCurrentMdhSuggestions.map((item, idx) => renderUdctSuggestionItem(item, idx, 'mdh')).join('');
                sugBox.classList.remove('hidden');
            } else {
                sugBox.innerHTML = '<div class="px-3 py-2 text-xs text-slate-400 italic">Không tìm thấy đơn hàng' + (maGianCurrent ? ` thuộc gian [${maGianCurrent.toUpperCase()}]` : '') + '</div>';
                sugBox.classList.remove('hidden');
            }
        } else {
            sugBox.innerHTML = '';
            sugBox.classList.add('hidden');
        }
    }
}

function setHhSku(value) {
    const input = document.getElementById('hhEditSKU');
    if (!input) return;

    input.value = value;
    const sugBox = document.getElementById('hhSkuSuggestions');
    if (sugBox) sugBox.classList.add('hidden');

    const skuCtInput = document.getElementById('hhEditSKUCT');
    if (skuCtInput) skuCtInput.value = '';
    handleHhSkuChange();

    if (skuCtInput) {
        skuCtInput.focus();
        handleHhSkuCtChange();
    }
}

function handleHhSkuChange() {
    const sku = (document.getElementById('hhEditSKU')?.value || '').toString().trim().toUpperCase();
    const sugBox = document.getElementById('hhSkuCtSuggestions');
    const skuCtBtns = document.getElementById('hhSkuCtButtons');
    renderHhSkuSuggestions(false);

    if ((!sanphamData || sanphamData.length === 0) && sku) {
        ensureHhCatalogLoaded(handleHhSkuChange);
        return;
    }

    if (sugBox) {
        const filteredItems = sanphamData.filter(item => {
            const code = (item.sku_con || '').toString().trim();
            if (code.length <= 5) return false;
            return (item.id_sp || '').toUpperCase() === sku || code.toUpperCase().startsWith(sku);
        }).sort((a, b) => (b.sku_con || '').toString().localeCompare((a.sku_con || '').toString()));

        if (sku && filteredItems.length > 0) {
            sugBox.innerHTML = filteredItems.map(item => `
                <div class="suggestion-item" onmousedown="event.preventDefault(); setHhSkuCt('${escapeHtml(item.sku_con)}')" onclick="setHhSkuCt('${escapeHtml(item.sku_con)}')">
                    <span class="item-code">${escapeHtml(item.sku_con)}</span>
                    <span class="item-name">${escapeHtml(item.ten_sp)}</span>
                </div>
            `).join('');
        } else {
            sugBox.innerHTML = '';
            sugBox.classList.add('hidden');
        }

        if (skuCtBtns) {
            const uniqueIds = [...new Set(filteredItems.map(i => i.sku_con))].slice(0, 10);
            if (sku && uniqueIds.length > 0) {
                skuCtBtns.innerHTML = uniqueIds.map(v => `
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
}

function setHhSkuCt(value) {
    const input = document.getElementById('hhEditSKUCT');
    if (input) {
        input.value = value;
        handleHhSkuCtChange();
    }
    const sugBox = document.getElementById('hhSkuCtSuggestions');
    if (sugBox) sugBox.classList.add('hidden');
    const skuCtBtns = document.getElementById('hhSkuCtButtons');
    if (skuCtBtns) skuCtBtns.innerHTML = '';
}

function handleHhSkuCtChange(forceShow = false) {
    const skuCtInput = document.getElementById('hhEditSKUCT');
    if (!skuCtInput) return;

    const val = skuCtInput.value.trim();
    const sku = (document.getElementById('hhEditSKU')?.value || '').toString().trim().toLowerCase();
    const sugBox = document.getElementById('hhSkuCtSuggestions');

    if ((!sanphamData || sanphamData.length === 0) && (val || sku || forceShow)) {
        if (sugBox) {
            sugBox.innerHTML = '<div class="suggestion-item"><span class="item-name">Đang tải danh mục SKU CT...</span></div>';
            sugBox.classList.remove('hidden');
        }
        ensureHhCatalogLoaded(() => handleHhSkuCtChange(forceShow));
        return;
    }

    if (sugBox) {
        if (val.length >= 1 || sku || forceShow) {
            const search = val.toLowerCase();
            const suggestions = sanphamData.filter(item => {
                const code = (item.sku_con || '').toString().trim();
                if (code.length <= 5) return false;
                if (!search) {
                    return sku
                        ? (item.id_sp || '').toLowerCase() === sku ||
                        code.toLowerCase().startsWith(sku)
                        : true;
                }
                return code.toLowerCase().includes(search) ||
                    (item.ten_sp || '').toLowerCase().includes(search);
            }).sort((a, b) => (b.sku_con || '').toString().localeCompare((a.sku_con || '').toString())).slice(0, 50);

            if (suggestions.length > 0) {
                sugBox.innerHTML = suggestions.map(item => `
                    <div class="suggestion-item" onmousedown="event.preventDefault(); setHhSkuCt('${escapeHtml(item.sku_con)}')" onclick="setHhSkuCt('${escapeHtml(item.sku_con)}')">
                        <span class="item-code">${escapeHtml(item.sku_con)}</span>
                        <span class="item-name">${escapeHtml(item.ten_sp)}</span>
                    </div>
                `).join('');
                sugBox.classList.remove('hidden');
            } else {
                sugBox.innerHTML = '';
                sugBox.classList.add('hidden');
            }
        } else {
            sugBox.innerHTML = '';
            sugBox.classList.add('hidden');
        }
    }

    if (!val) {
        const skuInput = document.getElementById('hhEditSKU');
        if (skuInput) skuInput.value = '';
        return;
    }

    const match = sanphamData.find(item => (item.sku_con || '').toString().toUpperCase() === val.toUpperCase());
    if (match) {
        const skuInput = document.getElementById('hhEditSKU');
        if (skuInput) skuInput.value = match.id_sp || match.sku_con.substring(0, 4);
        const tenSpInput = document.getElementById('hhEditTenSP');
        if (tenSpInput) tenSpInput.value = match.ten_sp || tenSpInput.value;
    }
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
    const mvd = (val || '').toString().trim();
    if (!mvd) {
        if (noticeEl) noticeEl.classList.add('hidden');
        return;
    }

    const udctMatch = udctData.find(item => (item.mvd || '').toString().trim() === mvd);
    if (udctMatch && hhDrawerMode === 'create') {
        const maGianEl = document.getElementById('hhEditMaGian');
        const skuEl = document.getElementById('hhEditSKU');
        const skuCtEl = document.getElementById('hhEditSKUCT');
        const slgEl = document.getElementById('hhEditSLG');
        const tenSpEl = document.getElementById('hhEditTenSP');

        if (maGianEl && !maGianEl.value) maGianEl.value = (udctMatch.ma_gian || '').toString().toUpperCase();
        if (skuCtEl && !skuCtEl.value) {
            skuCtEl.value = (udctMatch.id_sp_ct || '').toString().toUpperCase();
            if (typeof handleHhSkuCtChange === 'function') handleHhSkuCtChange();
        }
        if (skuEl && !skuEl.value) skuEl.value = (udctMatch.id_sp || '').toString().toUpperCase();
        if (slgEl && (!slgEl.value || slgEl.value === '1')) slgEl.value = udctMatch.slg_xuat || '';
        if (tenSpEl && !tenSpEl.value) tenSpEl.value = udctMatch.ten_sp || '';
        const mdhEl = document.getElementById('hhEditMDH');
        if (mdhEl && !mdhEl.value) mdhEl.value = (udctMatch.mdh || '').toString().trim();
        const ngayEl = document.getElementById('hhEditNgayNhan');
        if (ngayEl && !ngayEl.value && udctMatch.ngay) ngayEl.value = toYMD(udctMatch.ngay);
    }

    if (!noticeEl) return;
    const duplicate = hangHoanData.find(item => {
        const isDuplicateMvd = (item.mvd || '').toString().trim() === mvd || (item.mvd_2 || '').toString().trim() === mvd;
        if (!isDuplicateMvd) return false;
        if (hhDrawerMode === 'edit' && currentHangHoanEditIndex !== -1) {
            return item !== hangHoanData[currentHangHoanEditIndex];
        }
        return true;
    });

    if (duplicate) {
        noticeEl.textContent = `MVD da co trong Hang hoan ngay ${duplicate.ngay_nhan || '?'}`;
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
    isUploading = true;
    const statusLabel = document.getElementById('saveStatus');
    if (statusLabel) {
        statusLabel.textContent = 'Dang tai anh len ImgBB...';
        statusLabel.style.display = 'block';
    }

    try {
        const formData = new FormData();
        formData.append('image', file);
        const response = await fetch(`https://api.imgbb.com/1/upload?key=${CONFIG.imgbbApiKey}`, {
            method: 'POST',
            body: formData
        });
        const result = await response.json();
        if (result.success) {
            const targetInput = document.getElementById(`hhEditAnh${index}`);
            if (targetInput) targetInput.value = result.data.url;
            refreshHhImagePreviews();
    const historySection = document.getElementById('hhHistorySection');
    const historyContent = document.getElementById('hhHistoryContent');
    const lastUpdateBadge = document.getElementById('hhLastUpdateBadge');
    if (historySection) historySection.classList.remove('hidden');
    if (historyContent) {
        if (item.ghi_chu && item.ghi_chu.trim()) {
            historyContent.textContent = item.ghi_chu.trim();
        } else {
            historyContent.textContent = 'Chưa có ghi chú / lịch sử chỉnh sửa';
        }
    }
    if (lastUpdateBadge) {
        if (item.id_nv || item.udt) {
            lastUpdateBadge.textContent = `${item.id_nv || 'NV'} • ${item.udt || ''}`;
            lastUpdateBadge.title = `Người cập nhật: ${item.id_nv || ''} lúc ${item.udt || ''}`;
        } else {
            lastUpdateBadge.textContent = '';
        }
    }
            isUploading = false;
            await saveHhDetail();
        } else {
            alert('Loi tai anh len ImgBB: ' + (result.error?.message || 'Khong xac dinh'));
            isUploading = false;
            refreshHhImagePreviews();
        }
    } catch (err) {
        console.error('Upload error:', err);
        alert('Loi ket noi khi tai anh: ' + err.message);
        isUploading = false;
        refreshHhImagePreviews();
    } finally {
        if (statusLabel) statusLabel.style.display = 'none';
        input.value = '';
    }
}

function refreshHhImagePreviews() {
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
                    <svg class="w-8 h-8 mx-auto text-slate-300 group-hover:text-primary transition-colors" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 9a2 2 0 012-2h.93a2 2 0 001.664-.89l.812-1.22A2 2 0 0110.07 4h3.86a2 2 0 011.664.89l.812 1.22A2 2 0 0018.07 7H19a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2V9z" />
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 13a3 3 0 11-6 0 3 3 0 016 0z" />
                    </svg>
                    <span class="text-[10px] text-slate-400 block mt-1 font-bold">Cham de tai anh</span>
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

    let mvd2 = '';
    let maGian = '';
    let sku = '';
    let skuCt = '';
    let slg = '1';
    let tenSp = '';
    let mdh = '';

    const udctMatch = udctData.find(item => (item.mvd || '').toString().trim() === mvd);
    if (udctMatch) {
        maGian = (udctMatch.ma_gian || '').toString().toUpperCase();
        skuCt = (udctMatch.id_sp_ct || '').toString().toUpperCase();
        sku = skuCt ? (udctMatch.id_sp || '').toString().toUpperCase() : '';
        slg = udctMatch.slg_xuat || '1';
        tenSp = udctMatch.ten_sp || '';
        mdh = (udctMatch.mdh || '').toString().trim();
    }

    if (skuCt) {
        const matchedSp = sanphamData.find(i => (i.sku_con || '').toString().trim().toUpperCase() === skuCt.toString().trim().toUpperCase());
        if (matchedSp) {
            sku = matchedSp.id_sp || matchedSp.sku_con.substring(0, 4) || sku;
            if (!tenSp) tenSp = matchedSp.ten_sp || matchedSp.ten || '';
        }
    }

    const nowStr = formatDateTimeNow();
    const editor = getEditorDisplayName();
    const initGhiChu = `[${nowStr} - ${editor}] Quét nhanh MVD`;
    const appendValues = [[
        `${Date.now()}`,
        today,
        mvd,
        mvd2,
        maGian,
        '',
        '',
        '',
        '',
        sku,
        skuCt,
        slg,
        tenSp,
        initGhiChu,
        '',
        '',
        '',
        editor,
        nowStr,
        (mvd && maGian) ? `${mvd}-${maGian}` : '',
        'KHO',
        '',
        '',
        '',
        '',
        ''
    ]];
    const appendUrl = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.hhbhSheetName}!A:A:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
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
            })).sort((a, b) => b.rowIndex - a.rowIndex);
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
            const rowText = `${item.mvd || ''} ${item.mvd_2 || ''} ${item.ma_gian || ''} ${item.sku || ''} ${item.sku_ct || ''} ${item.ten_sp || ''} ${item.tinh_trang || ''} ${item.id_dh || ''}`.toLowerCase();
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
    document.getElementById('hhSaveButton').textContent = isKinhDoanh ? 'Lưu MVD 2 & MDH' : 'Lưu thay đổi';
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
    const ngayNhanInput = document.getElementById('hhEditNgayNhan');
    if (ngayNhanInput) ngayNhanInput.value = item.ngay_nhan || '';
    const mdhInput = document.getElementById('hhEditMDH');
    if (mdhInput) mdhInput.value = item.id_dh || '';
    populateHhFormOptions();
    renderHhKhoButtons(item.kho || 'KHO');
    refreshHhImagePreviews();
    const historySection = document.getElementById('hhHistorySection');
    const historyContent = document.getElementById('hhHistoryContent');
    const lastUpdateBadge = document.getElementById('hhLastUpdateBadge');
    if (historySection) historySection.classList.add('hidden');
    if (historyContent) historyContent.textContent = '';
    if (lastUpdateBadge) lastUpdateBadge.textContent = '';
    const noticeEl = document.getElementById('hhMvdDuplicateNotice');
    if (noticeEl) noticeEl.classList.add('hidden');
    ['hhEditMVD', 'hhEditMaGian', 'hhEditSKU', 'hhEditSKUCT', 'hhEditSLG', 'hhEditTinhTrang', 'hhEditTenSP', 'hhEditKho', 'hhEditNgayNhan'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.disabled = isKinhDoanh;
            const parent = el.closest('.group') || el.parentElement;
            if (parent) {
                const clearBtn = parent.querySelector('button[onclick*="clearHhInput"]');
                if (clearBtn) {
                    clearBtn.disabled = isKinhDoanh;
                    clearBtn.style.display = isKinhDoanh ? 'none' : '';
                }
            }
        }
    });
    ['hhEditMVD2', 'hhEditMDH'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.disabled = false;
            const parent = el.closest('.group') || el.parentElement;
            if (parent) {
                const clearBtn = parent.querySelector('button[onclick*="clearHhInput"]');
                if (clearBtn) {
                    clearBtn.disabled = false;
                    clearBtn.style.display = '';
                }
            }
        }
    });
    const qrMvdBtn = document.querySelector('button[onclick*="scanQrForHhMvd()"]');
    if (qrMvdBtn) {
        qrMvdBtn.disabled = isKinhDoanh;
        qrMvdBtn.style.opacity = isKinhDoanh ? '0.5' : '1';
        qrMvdBtn.style.pointerEvents = isKinhDoanh ? 'none' : '';
    }
    document.querySelectorAll('button[onclick*="stepHhEditSLG"], button[onclick*="stepHhEditNgayNhan"]').forEach(btn => {
        btn.disabled = isKinhDoanh;
        btn.style.opacity = isKinhDoanh ? '0.5' : '1';
        btn.style.pointerEvents = isKinhDoanh ? 'none' : '';
    });
    document.querySelectorAll('#hhEditKhoButtons button').forEach(btn => btn.disabled = isKinhDoanh);
    const copyBtn = document.getElementById('hhCopyButton');
    if (copyBtn) copyBtn.style.display = isKinhDoanh ? 'none' : 'inline-flex';
    const deleteBtn = document.getElementById('hhDeleteButton');
    if (deleteBtn) {
        deleteBtn.style.display = isKinhDoanh ? 'none' : 'inline-block';
        deleteBtn.classList.toggle('hidden', isKinhDoanh);
    }
    const footer = document.getElementById('hhDrawerFooter') || document.querySelector('#hhDrawer .pt-4.border-t.border-slate-200.flex.gap-3');
    if (footer) footer.style.display = '';
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
    const defaultDate = document.getElementById('filterHHTo')?.value || document.getElementById('filterHHFrom')?.value || new Date().toISOString().split('T')[0];
    const ngayNhanInput = document.getElementById('hhEditNgayNhan');
    if (ngayNhanInput) ngayNhanInput.value = defaultDate;
    const mdhInput = document.getElementById('hhEditMDH');
    if (mdhInput) mdhInput.value = '';
    populateHhFormOptions();
    renderHhKhoButtons('KHO');
    refreshHhImagePreviews();
    const noticeEl = document.getElementById('hhMvdDuplicateNotice');
    if (noticeEl) noticeEl.classList.add('hidden');
    ['hhEditMVD', 'hhEditMaGian', 'hhEditSKU', 'hhEditSKUCT', 'hhEditSLG', 'hhEditTinhTrang', 'hhEditTenSP', 'hhEditKho', 'hhEditNgayNhan', 'hhEditMDH'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.disabled = false;
            const parent = el.closest('.group') || el.parentElement;
            if (parent) {
                const clearBtn = parent.querySelector('button[onclick*="clearHhInput"]');
                if (clearBtn) {
                    clearBtn.disabled = false;
                    clearBtn.style.display = '';
                }
            }
        }
    });
    const qrMvdBtn = document.querySelector('button[onclick*="scanQrForHhMvd()"]');
    if (qrMvdBtn) {
        qrMvdBtn.disabled = false;
        qrMvdBtn.style.opacity = '1';
        qrMvdBtn.style.pointerEvents = '';
    }
    document.querySelectorAll('button[onclick*="stepHhEditSLG"], button[onclick*="stepHhEditNgayNhan"]').forEach(btn => {
        btn.disabled = false;
        btn.style.opacity = '1';
        btn.style.pointerEvents = '';
    });
    document.querySelectorAll('#hhEditKhoButtons button').forEach(btn => btn.disabled = false);
    const copyBtn = document.getElementById('hhCopyButton');
    if (copyBtn) copyBtn.style.display = 'none';
    const deleteBtn = document.getElementById('hhDeleteButton');
    if (deleteBtn) {
        deleteBtn.style.display = 'none';
        deleteBtn.classList.add('hidden');
    }
    const footer = document.getElementById('hhDrawerFooter') || document.querySelector('#hhDrawer .pt-4.border-t.border-slate-200.flex.gap-3');
    if (footer) footer.style.display = '';
    document.getElementById('hhDrawerOverlay').classList.remove('hidden');
    document.getElementById('hhDrawer').classList.add('open');

    setTimeout(() => {
        const mvdInput = document.getElementById('hhEditMVD');
        if (mvdInput) mvdInput.focus();
    }, 400);
}

function copyCurrentHangHoan() {
    if (currentUser && currentUser.role === 'kinhdoanh') {
        alert('Tài khoản KINHDOANH không được thêm mới Dữ liệu Hàng hoàn.');
        return;
    }
    hhDrawerMode = 'create';
    currentHangHoanEditIndex = -1;
    document.getElementById('hhDrawerTitle').textContent = 'Sao chép hàng hoàn';
    document.getElementById('hhSaveButton').textContent = 'Thêm mới';
    document.getElementById('hhDrawerRowId').textContent = `Row ID: NEW-${Date.now()}`;
    
    const skuCtInput = document.getElementById('hhEditSKUCT');
    const skuInput = document.getElementById('hhEditSKU');
    if (skuCtInput) skuCtInput.value = '';
    if (skuInput) skuInput.value = '';

    ['hhEditMVD', 'hhEditMaGian', 'hhEditSKU', 'hhEditSKUCT', 'hhEditSLG', 'hhEditTinhTrang', 'hhEditTenSP', 'hhEditKho', 'hhEditNgayNhan', 'hhEditMDH'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.disabled = false;
            const parent = el.closest('.group') || el.parentElement;
            if (parent) {
                const clearBtn = parent.querySelector('button[onclick*="clearHhInput"]');
                if (clearBtn) {
                    clearBtn.disabled = false;
                    clearBtn.style.display = '';
                }
            }
        }
    });
    const qrMvdBtn = document.querySelector('button[onclick*="scanQrForHhMvd()"]');
    if (qrMvdBtn) {
        qrMvdBtn.disabled = false;
        qrMvdBtn.style.opacity = '1';
        qrMvdBtn.style.pointerEvents = '';
    }
    document.querySelectorAll('button[onclick*="stepHhEditSLG"], button[onclick*="stepHhEditNgayNhan"]').forEach(btn => {
        btn.disabled = false;
        btn.style.opacity = '1';
        btn.style.pointerEvents = '';
    });
    document.querySelectorAll('#hhEditKhoButtons button').forEach(btn => btn.disabled = false);

    const copyBtn = document.getElementById('hhCopyButton');
    if (copyBtn) copyBtn.style.display = 'none';
    const deleteBtn = document.getElementById('hhDeleteButton');
    if (deleteBtn) {
        deleteBtn.style.display = 'none';
        deleteBtn.classList.add('hidden');
    }

    const noticeEl = document.getElementById('hhMvdDuplicateNotice');
    if (noticeEl) noticeEl.classList.add('hidden');

    const skuCtBtns = document.getElementById('hhSkuCtButtons');
    if (skuCtBtns) skuCtBtns.innerHTML = '';
    const sugBox = document.getElementById('hhSkuCtSuggestions');
    if (sugBox) sugBox.classList.add('hidden');

    if (typeof showToast === 'function') {
        showToast('Đã sao chép đơn hàng hoàn. SKU CT được để trống.', 'info');
    }
}

function closeHhDetailDrawer() {
    document.getElementById('hhDrawerOverlay').classList.add('hidden');
    document.getElementById('hhDrawer').classList.remove('open');
    const mvd2Sug = document.getElementById('hhMvd2Suggestions');
    if (mvd2Sug) mvd2Sug.classList.add('hidden');
    const mdhSug = document.getElementById('hhMdhSuggestions');
    if (mdhSug) mdhSug.classList.add('hidden');
    const skuSug = document.getElementById('hhSkuSuggestions');
    if (skuSug) skuSug.classList.add('hidden');
    const skuCtSug = document.getElementById('hhSkuCtSuggestions');
    if (skuCtSug) skuCtSug.classList.add('hidden');
    const deleteBtn = document.getElementById('hhDeleteButton');
    if (deleteBtn) {
        deleteBtn.style.display = 'none';
        deleteBtn.classList.add('hidden');
    }
    currentHangHoanEditIndex = -1;
    hhDrawerMode = 'edit';
}

async function deleteCurrentHangHoan() {
    if (hhDrawerMode === 'create' || currentHangHoanEditIndex === -1) {
        return;
    }
    const isKinhDoanh = currentUser && currentUser.role === 'kinhdoanh';
    if (isKinhDoanh) {
        alert('Tài khoản KINHDOANH không có quyền xóa dữ liệu Hàng hoàn.');
        return;
    }
    const item = hangHoanData[currentHangHoanEditIndex];
    if (!item) {
        if (typeof showToast === 'function') showToast('Không xác định được dòng cần xóa.', 'error');
        else alert('Không xác định được dòng cần xóa.');
        return;
    }
    const rowIndex = item.rowIndex || (currentHangHoanEditIndex + 2);
    const confirmMsg = `Bạn có chắc chắn muốn xóa dòng Hàng hoàn này?\n- MVD: ${item.mvd || item.mvd_2 || '-'}\n- SKU CT: ${item.sku_ct || '-'}\n- Số lượng: ${item.slg || '1'}`;
    if (!confirm(confirmMsg)) return;

    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');

    try {
        const token = await getAccessToken();
        if (!token) {
            alert('Không thể xác thực Google Sheets. Vui lòng đăng nhập lại.');
            return;
        }

        let sheetId = null;
        if (typeof fetchSheetMeta === 'function') {
            sheetId = await fetchSheetMeta(CONFIG.hhbhSheetName, token);
        }
        if (sheetId === null || sheetId === undefined) {
            const metaResp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}?fields=sheets(properties(sheetId,title))`, {
                headers: { 'Authorization': `Bearer ${token}` }
            });
            const metaData = await metaResp.json();
            const sheet = (metaData.sheets || []).find(s => s.properties?.title === CONFIG.hhbhSheetName);
            sheetId = sheet?.properties?.sheetId ?? null;
        }

        if (sheetId === null || sheetId === undefined) {
            throw new Error(`Không tìm thấy sheetId của ${CONFIG.hhbhSheetName}`);
        }

        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}:batchUpdate`;
        const body = {
            requests: [{
                deleteDimension: {
                    range: {
                        sheetId: sheetId,
                        dimension: 'ROWS',
                        startIndex: rowIndex - 1,
                        endIndex: rowIndex
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
            const errText = await resp.text();
            console.error('Delete HH error:', errText);
            alert('Lỗi khi xóa dòng Hàng hoàn trên Google Sheet.');
            return;
        }

        closeHhDetailDrawer();
        await fetchHangHoanData();
        if (typeof showToast === 'function') {
            showToast('Đã xóa dòng Hàng hoàn thành công!', 'success');
        }
    } catch (err) {
        console.error('Delete HH error:', err);
        alert('Có lỗi khi xóa Hàng hoàn: ' + (err.message || 'Lỗi không xác định'));
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
}

async function saveHhDetail() {
    const isKinhDoanh = currentUser && currentUser.role === 'kinhdoanh';
    if (isKinhDoanh) {
        if (hhDrawerMode === 'create' || currentHangHoanEditIndex === -1) {
            alert('Tài khoản KINHDOANH chỉ được sửa MVD 2 và MDH trên dòng Hàng hoàn đã có.');
            return;
        }
        const item = hangHoanData[currentHangHoanEditIndex];
        if (!item) return;
        const loadingOverlay = document.getElementById('loadingOverlay');
        loadingOverlay.classList.remove('hidden');
        try {
            const token = await getAccessToken();
            const rowIndex = item.rowIndex || (hangHoanData.indexOf(item) + 2);
            const mvd2 = (document.getElementById('hhEditMVD2')?.value || '').trim();
            const mdh = (document.getElementById('hhEditMDH')?.value || '').trim();

            const oldMvd2 = (item.mvd_2 || '').trim();
            const oldMdh = (item.id_dh || '').trim();

            const diffs = [];
            if (mvd2 !== oldMvd2) {
                diffs.push(`MVD 2: "${oldMvd2}" ➔ "${mvd2}"`);
            }
            if (mdh !== oldMdh) {
                diffs.push(`MDH: "${oldMdh}" ➔ "${mdh}"`);
            }

            const nowStr = formatDateTimeNow();
            const editor = getEditorDisplayName();

            let newGhiChu = item.ghi_chu || '';
            if (diffs.length > 0) {
                const logEntry = `[${nowStr} - ${editor}] Sửa: ${diffs.join(', ')}`;
                newGhiChu = newGhiChu.trim() ? `${logEntry}\n${newGhiChu.trim()}` : logEntry;
            }

            const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
            const resp = await fetch(url, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    valueInputOption: 'USER_ENTERED',
                    data: [
                        { range: `${CONFIG.hhbhSheetName}!D${rowIndex}`, values: [[mvd2]] },
                        { range: `${CONFIG.hhbhSheetName}!N${rowIndex}`, values: [[newGhiChu]] },
                        { range: `${CONFIG.hhbhSheetName}!R${rowIndex}`, values: [[editor]] },
                        { range: `${CONFIG.hhbhSheetName}!S${rowIndex}`, values: [[nowStr]] },
                        { range: `${CONFIG.hhbhSheetName}!W${rowIndex}`, values: [[mdh]] }
                    ]
                })
            });
            if (resp.ok) {
                item.mvd_2 = mvd2;
                item.id_dh = mdh;
                item.ghi_chu = newGhiChu;
                item.id_nv = editor;
                item.udt = nowStr;
                filterHangHoanData();
                closeHhDetailDrawer();
                showToast('Đã lưu MVD 2 & MDH và ghi nhận lịch sử thành công!', 'success');
            } else {
                const errText = await resp.text();
                console.error('Save HH MVD 2 & MDH error:', errText);
                alert('Lỗi khi lưu MVD 2 & MDH.');
            }
        } catch (err) {
            console.error('Save HH MVD 2 & MDH Error:', err);
            alert('Đã xảy ra lỗi khi lưu MVD 2 & MDH: ' + (err.message || 'Lỗi không xác định'));
        } finally {
            loadingOverlay.classList.add('hidden');
        }
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
        const today = new Date().toISOString().split('T')[0];
        const newData = {
            ngay_nhan: document.getElementById('hhEditNgayNhan')?.value || today,
            mvd: (document.getElementById('hhEditMVD').value || '').trim(),
            mvd_2: (document.getElementById('hhEditMVD2').value || '').trim(),
            ma_gian: (document.getElementById('hhEditMaGian').value || '').trim(),
            sku: (document.getElementById('hhEditSKU').value || '').trim(),
            sku_ct: (document.getElementById('hhEditSKUCT').value || '').trim(),
            slg: (document.getElementById('hhEditSLG').value || '1').trim(),
            tinh_trang: (document.getElementById('hhEditTinhTrang').value || '').trim(),
            ten_sp: (document.getElementById('hhEditTenSP').value || '').trim(),
            kho: (document.getElementById('hhEditKho').value || 'KHO').trim(),
            anh_1: (document.getElementById('hhEditAnh1').value || '').trim(),
            anh_2: (document.getElementById('hhEditAnh2').value || '').trim(),
            anh_3: (document.getElementById('hhEditAnh3').value || '').trim(),
            id_dh: (document.getElementById('hhEditMDH')?.value || '').trim()
        };
        if (newData.sku_ct) {
            const matchedSp = sanphamData.find(i => (i.sku_con || '').toString().trim().toUpperCase() === newData.sku_ct.toString().trim().toUpperCase());
            if (matchedSp) {
                newData.sku = matchedSp.id_sp || matchedSp.sku_con.substring(0, 4) || newData.sku;
                if (!newData.ten_sp) newData.ten_sp = matchedSp.ten_sp || matchedSp.ten || '';
                const skuInput = document.getElementById('hhEditSKU');
                if (skuInput) skuInput.value = newData.sku;
            }
        } else {
            newData.sku = '';
            const skuInput = document.getElementById('hhEditSKU');
            if (skuInput) skuInput.value = '';
        }

        const nowStr = formatDateTimeNow();
        const editor = getEditorDisplayName();

        if (isCreateMode) {
            const mvd = (newData.mvd || '').trim();
            const mvd2 = (newData.mvd_2 || '').trim();
            const maGian = (newData.ma_gian || '').trim();
            const initGhiChu = `[${nowStr} - ${editor}] Tạo mới`;
            const appendValues = [[
                `${Date.now()}`,
                newData.ngay_nhan || today,
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
                initGhiChu, // ghi_chu
                newData.tinh_trang || '',
                '', // trang_thai
                '', // sku_slg
                editor, // id_nv
                nowStr, // udt
                mvd && maGian ? `${mvd}-${maGian}` : '',
                newData.kho || '',
                '', // lb3
                newData.id_dh || '', // id_dh / MDH
                '', // id_dh_ct
                '', // stt
                ''  // danh_dau
            ]];
            const appendUrl = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.hhbhSheetName}!A:A:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
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
            showToast('Đã thêm mới và ghi nhận lịch sử thành công!', 'success');
            return;
        }

        const oldNgayNhan = (item.ngay_nhan || '').trim();
        const oldMvd = (item.mvd || '').trim();
        const oldMvd2 = (item.mvd_2 || '').trim();
        const oldMaGian = (item.ma_gian || '').trim();
        const oldSku = (item.sku || '').trim();
        const oldSkuCt = (item.sku_ct || '').trim();
        const oldSlg = (item.slg || '').toString().trim();
        const oldTinhTrang = (item.tinh_trang || '').trim();
        const oldTenSp = (item.ten_sp || '').trim();
        const oldKho = (item.kho || '').trim();
        const oldMdh = (item.id_dh || '').trim();
        const oldAnh1 = (item.anh_1 || '').trim();
        const oldAnh2 = (item.anh_2 || '').trim();
        const oldAnh3 = (item.anh_3 || '').trim();

        const diffs = [];
        if (newData.ngay_nhan !== oldNgayNhan) diffs.push(`Ngày nhận: "${oldNgayNhan}" ➔ "${newData.ngay_nhan}"`);
        if (newData.mvd !== oldMvd) diffs.push(`MVD: "${oldMvd}" ➔ "${newData.mvd}"`);
        if (newData.mvd_2 !== oldMvd2) diffs.push(`MVD 2: "${oldMvd2}" ➔ "${newData.mvd_2}"`);
        if (newData.ma_gian !== oldMaGian) diffs.push(`Mã gian: "${oldMaGian}" ➔ "${newData.ma_gian}"`);
        if (newData.sku !== oldSku) diffs.push(`SKU: "${oldSku}" ➔ "${newData.sku}"`);
        if (newData.sku_ct !== oldSkuCt) diffs.push(`SKU CT: "${oldSkuCt}" ➔ "${newData.sku_ct}"`);
        if (newData.slg !== oldSlg) diffs.push(`SLG: "${oldSlg}" ➔ "${newData.slg}"`);
        if (newData.tinh_trang !== oldTinhTrang) diffs.push(`Tình trạng: "${oldTinhTrang}" ➔ "${newData.tinh_trang}"`);
        if (newData.ten_sp !== oldTenSp) diffs.push(`Tên SP: "${oldTenSp}" ➔ "${newData.ten_sp}"`);
        if (newData.kho !== oldKho) diffs.push(`Kho: "${oldKho}" ➔ "${newData.kho}"`);
        if (newData.id_dh !== oldMdh) diffs.push(`MDH: "${oldMdh}" ➔ "${newData.id_dh}"`);
        if (newData.anh_1 !== oldAnh1 || newData.anh_2 !== oldAnh2 || newData.anh_3 !== oldAnh3) {
            diffs.push('Ảnh: Cập nhật ảnh');
        }

        let newGhiChu = item.ghi_chu || '';
        if (diffs.length > 0) {
            const logEntry = `[${nowStr} - ${editor}] Sửa: ${diffs.join(', ')}`;
            newGhiChu = newGhiChu.trim() ? `${logEntry}\n${newGhiChu.trim()}` : logEntry;
        }

        const rowIndex = item.rowIndex || (hangHoanData.indexOf(item) + 2);
        const batchUpdates = [
            { range: `${CONFIG.hhbhSheetName}!B${rowIndex}`, values: [[newData.ngay_nhan]] },
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
            { range: `${CONFIG.hhbhSheetName}!N${rowIndex}`, values: [[newGhiChu]] },
            { range: `${CONFIG.hhbhSheetName}!O${rowIndex}`, values: [[newData.tinh_trang]] },
            { range: `${CONFIG.hhbhSheetName}!R${rowIndex}`, values: [[editor]] },
            { range: `${CONFIG.hhbhSheetName}!S${rowIndex}`, values: [[nowStr]] },
            { range: `${CONFIG.hhbhSheetName}!T${rowIndex}`, values: [[(newData.mvd && newData.ma_gian) ? `${newData.mvd}-${newData.ma_gian}` : '']] },
            { range: `${CONFIG.hhbhSheetName}!U${rowIndex}`, values: [[newData.kho]] },
            { range: `${CONFIG.hhbhSheetName}!W${rowIndex}`, values: [[newData.id_dh]] }
        ];
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values:batchUpdate`;
        const resp = await fetch(url, { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data: batchUpdates }) });
        if (resp.ok) {
            Object.assign(item, newData, { ghi_chu: newGhiChu, id_nv: editor, udt: nowStr, rowIndex });
            filterHangHoanData();
            closeHhDetailDrawer();
            showToast('Đã lưu thay đổi và ghi nhận lịch sử thành công!', 'success');
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
                    <tr class="border-b border-slate-100 hover:bg-slate-50 transition-colors cursor-pointer" ondblclick="openHhDetail(${index})" title="${escapeHtml(item.ghi_chu ? `Lịch sử / Ghi chú:\n${item.ghi_chu}` : '')}">
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


    Object.assign(window.AppModules = window.AppModules || {}, { ['hanghoan']: true });
    window.fillHangHoanFilterOptions = fillHangHoanFilterOptions;
    window.setHHKhoFilter = setHHKhoFilter;
    window.setHangHoanToday = setHangHoanToday;
    window.shiftHHFilterDate = shiftHHFilterDate;
    window.changeHangHoanDate = changeHangHoanDate;
    window.stepHhEditNgayNhan = stepHhEditNgayNhan;
    window.stepHhEditSLG = stepHhEditSLG;
    window.selectHhUdctForMvd2ByIndex = selectHhUdctForMvd2ByIndex;
    window.selectHhUdctForMdhByIndex = selectHhUdctForMdhByIndex;
    window.setHhMvd2 = setHhMvd2;
    window.handleHhMvd2Change = handleHhMvd2Change;
    window.setHhMdh = setHhMdh;
    window.handleHhMdhChange = handleHhMdhChange;
document.addEventListener('click', (e) => {
    if (!e.target.closest('#hhEditMVD2') && !e.target.closest('#hhMvd2Suggestions')) {
        const box = document.getElementById('hhMvd2Suggestions');
        if (box) box.classList.add('hidden');
    }
    if (!e.target.closest('#hhEditMDH') && !e.target.closest('#hhMdhSuggestions')) {
        const box = document.getElementById('hhMdhSuggestions');
        if (box) box.classList.add('hidden');
    }
});

    window.openImagePreview = openImagePreview;
    window.closeImagePreview = closeImagePreview;
        window.renderHhKhoButtons = renderHhKhoButtons;
    window.setHhKho = setHhKho;
    window.populateHhFormOptions = populateHhFormOptions;
    window.ensureHhCatalogLoaded = ensureHhCatalogLoaded;
    window.getHhSkuMatches = getHhSkuMatches;
    window.renderHhSkuSuggestions = renderHhSkuSuggestions;
    window.setHhSku = setHhSku;
    window.handleHhSkuChange = handleHhSkuChange;
    window.setHhSkuCt = setHhSkuCt;
    window.handleHhSkuCtChange = handleHhSkuCtChange;
    window.scanQrToInput = scanQrToInput;
    window.handleHhMvdInputChange = handleHhMvdInputChange;
    window.scanQrForHhMvd = scanQrForHhMvd;
    window.scanQrForHhMvd2 = scanQrForHhMvd2;
    window.uploadImageHh = uploadImageHh;
    window.refreshHhImagePreviews = refreshHhImagePreviews;
    window.removeHhImage = removeHhImage;
    window.appendHangHoanQuickByMvd = appendHangHoanQuickByMvd;
    window.stopContinuousMvdScan = stopContinuousMvdScan;
    window.startContinuousMvdScan = startContinuousMvdScan;
    window.startHHSearchQrScan = startHHSearchQrScan;
    window.fetchHangHoanData = fetchHangHoanData;
    window.filterHangHoanData = filterHangHoanData;
    window.openHhDetail = openHhDetail;
    window.openNewHangHoanDrawer = openNewHangHoanDrawer;
    window.copyCurrentHangHoan = copyCurrentHangHoan;
    window.closeHhDetailDrawer = closeHhDetailDrawer;
    window.deleteCurrentHangHoan = deleteCurrentHangHoan;
    window.saveHhDetail = saveHhDetail;
    window.buildSkuTongMap = buildSkuTongMap;
    window.getSkuTongForItem = getSkuTongForItem;
    window.getMaGianForItem = getMaGianForItem;
    window.renderHangHoanTable = renderHangHoanTable;
    window.exportHangHoanSummaryToExcel = exportHangHoanSummaryToExcel;
    window.exportHangHoanToExcel = exportHangHoanToExcel;
    window.exportHangHoanToMisa = exportHangHoanToMisa;
})();
