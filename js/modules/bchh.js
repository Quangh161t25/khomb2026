// bchh - Module Pattern (IIFE)
(function () {
// BÁO CÁO HÀNG HOÀN LOGIC
let bchhFilteredData = [];

function setBCHHQuickDate(type) {
    const today = new Date();
    const toDate = document.getElementById('bcHHToDate');
    const fromDate = document.getElementById('bcHHFromDate');

    if (type === 'today') {
        const d = String(today.getDate()).padStart(2, '0');
        const m = String(today.getMonth() + 1).padStart(2, '0');
        const y = today.getFullYear();
        const dateStr = `${y}-${m}-${d}`;
        fromDate.value = dateStr;
        toDate.value = dateStr;
    } else if (type === 'thisWeek') {
        const day = today.getDay();
        const diff = today.getDate() - day + (day === 0 ? -6 : 1);
        const startOfWeek = new Date(today.setDate(diff));

        fromDate.value = `${startOfWeek.getFullYear()}-${String(startOfWeek.getMonth() + 1).padStart(2, '0')}-${String(startOfWeek.getDate()).padStart(2, '0')}`;

        const now = new Date();
        toDate.value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    } else if (type === 'thisMonth') {
        const y = today.getFullYear();
        const m = String(today.getMonth() + 1).padStart(2, '0');
        fromDate.value = `${y}-${m}-01`;

        const now = new Date();
        toDate.value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    }

    filterBCHHData();
}

function changeBCHHDate(id, direction) {
    const input = document.getElementById(id);
    if (!input || !input.value) return;
    const d = new Date(input.value);
    d.setDate(d.getDate() + direction);
    input.value = d.toISOString().split('T')[0];
    filterBCHHData();
}

function filterBCHHParams() {
    filterBCHHData();
}

function filterBCHHData() {
    if (!hangHoanData || hangHoanData.length === 0) return;

    const fFrom = document.getElementById('bcHHFromDate').value;
    const fTo = document.getElementById('bcHHToDate').value;
    const fGian = (document.getElementById('filterBCHHMaGian').value || '').trim().toLowerCase();

    // Populate Mã gian datalist
    const maGianSet = new Set();
    hangHoanData.forEach(item => {
        if (item.ma_gian) maGianSet.add(item.ma_gian.trim());
    });
    const maGianList = document.getElementById('bcHHMaGianList');
    if (maGianList) {
        maGianList.innerHTML = Array.from(maGianSet).sort().map(mg => `<option value="${mg.replace(/"/g, '&quot;')}">`).join('');
    }

    bchhFilteredData = hangHoanData.filter(item => {
        const ngay = toYMD(item.ngay_nhan);
        if (fFrom && ngay < fFrom) return false;
        if (fTo && ngay > fTo) return false;
        if (fGian && !(item.ma_gian || '').toLowerCase().includes(fGian)) return false;
        return true;
    });

    renderBCHHStats();
}

function renderBCHHStats() {
    document.getElementById('bchhTotalOrders').textContent = bchhFilteredData.length.toLocaleString('vi-VN');
    const totalQty = bchhFilteredData.reduce((sum, item) => sum + (parseFloat(item.slg) || 0), 0);
    document.getElementById('bchhTotalQuantity').textContent = totalQty.toLocaleString('vi-VN');

    const byMaGian = {};
    const byTinhTrang = {};
    const bySku = {};
    const byKho = {};
    const byNgay = {};

    bchhFilteredData.forEach(item => {
        const mg = item.ma_gian || 'Trống';
        const tt = item.tinh_trang || 'Trống';
        const sku = item.sku || 'Trống';
        const kho = item.kho || 'Trống';
        const q = parseFloat(item.slg) || 0;

        if (!byMaGian[mg]) byMaGian[mg] = { don: 0, sp: 0 };
        byMaGian[mg].don += 1;
        byMaGian[mg].sp += q;

        if (!byTinhTrang[tt]) byTinhTrang[tt] = { don: 0, sp: 0 };
        byTinhTrang[tt].don += 1;
        byTinhTrang[tt].sp += q;

        if (!bySku[sku]) bySku[sku] = { don: 0, sp: 0 };
        bySku[sku].don += 1;
        bySku[sku].sp += q;

        if (!byKho[kho]) byKho[kho] = { don: 0, sp: 0 };
        byKho[kho].don += 1;
        byKho[kho].sp += q;

        const ngay = formatYmdToDmy(item.ngay_nhan) || item.ngay_nhan || 'Trống';
        if (!byNgay[ngay]) byNgay[ngay] = { don: 0, sp: 0, sortKey: item.ngay_nhan || '' };
        byNgay[ngay].don += 1;
        byNgay[ngay].sp += q;
    });

    const tbMaGian = document.getElementById('bchhMaGianTableBody');
    if (!Object.keys(byMaGian).length) {
        tbMaGian.innerHTML = '<tr><td colspan="3" class="text-center py-4 text-slate-500">Không có dữ liệu</td></tr>';
    } else {
        tbMaGian.innerHTML = Object.entries(byMaGian).sort((a, b) => b[1].don - a[1].don).map(([k, v]) => `
            <tr class="border-b border-slate-100 hover:bg-slate-50">
                <td class="px-4 py-2 text-sm font-medium text-slate-900">${escapeHtml(k)}</td>
                <td class="px-4 py-2 text-sm text-slate-700">${v.don.toLocaleString('vi-VN')}</td>
                <td class="px-4 py-2 text-sm text-slate-700">${v.sp.toLocaleString('vi-VN')}</td>
            </tr>
        `).join('');
    }

    const tbTinhTrang = document.getElementById('bchhTinhTrangTableBody');
    if (!Object.keys(byTinhTrang).length) {
        tbTinhTrang.innerHTML = '<tr><td colspan="3" class="text-center py-4 text-slate-500">Không có dữ liệu</td></tr>';
    } else {
        tbTinhTrang.innerHTML = Object.entries(byTinhTrang).sort((a, b) => b[1].don - a[1].don).map(([k, v]) => `
            <tr class="border-b border-slate-100 hover:bg-slate-50">
                <td class="px-4 py-2 text-sm font-medium text-slate-900">${escapeHtml(k)}</td>
                <td class="px-4 py-2 text-sm text-slate-700">${v.don.toLocaleString('vi-VN')}</td>
                <td class="px-4 py-2 text-sm text-slate-700">${v.sp.toLocaleString('vi-VN')}</td>
            </tr>
        `).join('');
    }

    const tbSku = document.getElementById('bchhSkuTableBody');
    if (tbSku) {
        if (!Object.keys(bySku).length) {
            tbSku.innerHTML = '<tr><td colspan="3" class="text-center py-4 text-slate-500">Không có dữ liệu</td></tr>';
        } else {
            tbSku.innerHTML = Object.entries(bySku).sort((a, b) => b[1].don - a[1].don).map(([k, v]) => `
                <tr class="border-b border-slate-100 hover:bg-slate-50">
                    <td class="px-4 py-2 text-sm font-medium text-slate-900">${escapeHtml(k)}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${v.don.toLocaleString('vi-VN')}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${v.sp.toLocaleString('vi-VN')}</td>
                </tr>
            `).join('');
        }
    }

    const tbKho = document.getElementById('bchhKhoTableBody');
    if (tbKho) {
        if (!Object.keys(byKho).length) {
            tbKho.innerHTML = '<tr><td colspan="3" class="text-center py-4 text-slate-500">Không có dữ liệu</td></tr>';
        } else {
            tbKho.innerHTML = Object.entries(byKho).sort((a, b) => b[1].don - a[1].don).map(([k, v]) => `
                <tr class="border-b border-slate-100 hover:bg-slate-50">
                    <td class="px-4 py-2 text-sm font-medium text-slate-900">${escapeHtml(k)}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${v.don.toLocaleString('vi-VN')}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${v.sp.toLocaleString('vi-VN')}</td>
                </tr>
            `).join('');
        }
    }

    const tbNgay = document.getElementById('bchhNgayTableBody');
    if (tbNgay) {
        if (!Object.keys(byNgay).length) {
            tbNgay.innerHTML = '<tr><td colspan="3" class="text-center py-4 text-slate-500">Không có dữ liệu</td></tr>';
        } else {
            tbNgay.innerHTML = Object.entries(byNgay).sort((a, b) => b[1].sortKey.localeCompare(a[1].sortKey)).map(([k, v]) => `
                <tr class="border-b border-slate-100 hover:bg-slate-50">
                    <td class="px-4 py-2 text-sm font-medium text-slate-900">${escapeHtml(k)}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${v.don.toLocaleString('vi-VN')}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${v.sp.toLocaleString('vi-VN')}</td>
                </tr>
            `).join('');
        }
    }

    const tbDetails = document.getElementById('bchhDetailsTableBody');
    if (tbDetails) {
        if (!bchhFilteredData.length) {
            tbDetails.innerHTML = '<tr><td colspan="8" class="text-center py-4 text-slate-500">Không có dữ liệu</td></tr>';
        } else {
            tbDetails.innerHTML = bchhFilteredData.map(item => `
                <tr class="hover:bg-slate-50 transition-colors">
                    <td class="px-4 py-2 text-sm text-slate-700 whitespace-nowrap">${escapeHtml(formatYmdToDmy(item.ngay_nhan) || item.ngay_nhan)}</td>
                    <td class="px-4 py-2 text-sm font-medium text-slate-900">${escapeHtml(item.mvd || '')}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${escapeHtml(item.ma_gian || '')}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${escapeHtml(item.sku || '')}</td>
                    <td class="px-4 py-2 text-sm text-right font-semibold text-slate-900">${(parseFloat(item.slg) || 0).toLocaleString('vi-VN')}</td>
                    <td class="px-4 py-2 text-sm text-slate-700 truncate max-w-[200px]" title="${escapeHtml(item.ten_sp || '')}">${escapeHtml(item.ten_sp || '')}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${escapeHtml(item.tinh_trang || '')}</td>
                    <td class="px-4 py-2 text-sm text-slate-700">${escapeHtml(item.kho || '')}</td>
                </tr>
            `).join('');
        }
    }
}

function exportBCHHToExcel() {
    if (!bchhFilteredData || bchhFilteredData.length === 0) {
        alert('Không có dữ liệu để xuất!');
        return;
    }

    const headers = ['Ngày nhận', 'MVD', 'Mã gian', 'SKU', 'SLG', 'Tên SP', 'Tình trạng', 'Kho', 'Trạng thái'];
    const excelData = [headers, ...bchhFilteredData.map(item => [
        formatYmdToDmy(item.ngay_nhan) || item.ngay_nhan,
        item.mvd || '',
        item.ma_gian || '',
        item.sku || '',
        item.slg || '',
        item.ten_sp || '',
        item.tinh_trang || '',
        item.kho || '',
        item.trạng_thái || ''
    ])];

    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'BC_Hang_Hoan');

    const fFrom = document.getElementById('bcHHFromDate').value || 'ToanBo';
    const fTo = document.getElementById('bcHHToDate').value || '';
    const now = new Date();
    const timeStr = now.getHours().toString().padStart(2, '0') + 'h' + now.getMinutes().toString().padStart(2, '0');

    let fName = `BCHH_${fFrom}`;
    if (fTo) fName += `_den_${fTo}`;
    fName += `_${timeStr}.xlsx`;

    XLSX.writeFile(wb, fName);
}



    Object.assign(window.AppModules = window.AppModules || {}, { ['bchh']: true });
    window.setBCHHQuickDate = setBCHHQuickDate;
    window.changeBCHHDate = changeBCHHDate;
    window.filterBCHHParams = filterBCHHParams;
    window.filterBCHHData = filterBCHHData;
    window.renderBCHHStats = renderBCHHStats;
    window.exportBCHHToExcel = exportBCHHToExcel;
})();
