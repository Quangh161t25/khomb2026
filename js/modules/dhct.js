// dhct - Module Pattern (IIFE)
(function () {
// Logic module DH_CT
async function fetchDHCTData(silent = false) {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (!silent && loadingOverlay) loadingOverlay.classList.remove('hidden');

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
                kho: row[5],     // Cột F
                id_sp_ct: row[6], // Cột G
                id_sp: row[7], // Cột H
                ten: row[8],     // Cột I
                so_luong: row[9], // Cột J
                gia_nhap: row[10],  // Cột K
                thanh_tien_nhap: row[11], // Cột L
                so_luong_2: row[12], // Cột M
                id_ton_kho: row[13], // Cột N
                xac_nhan: row[14] // Cột O
            }));
            if (silent) {
                // Cập nhật bảng nếu không đang mở drawer
                const isDHCTModule = pageTitle.textContent === 'Dữ liệu DH Chi Tiết' || pageTitle.textContent === 'Đơn hàng trên DH Chi Tiết' || pageTitle.textContent === 'Đơn hàng chi tiết';
                const isDHTongModule = pageTitle.textContent === 'Đơn hàng';
                if (isDHCTModule) {
                    renderDHCTTable();
                    renderUniqueDHCTTable();
                }
                if (isDHTongModule) {
                    renderDonhangTongTable();
                }
            } else {
                renderDHCTTable();
                renderUniqueDHCTTable();
                renderDonhangTongTable();
            }
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
    
    const fromDateInput = document.getElementById('filterDHCTFromDate');
    const toDateInput = document.getElementById('filterDHCTToDate');
    const nccInput = document.getElementById('filterDHCTNcc');
    
    const fromDate = fromDateInput && fromDateInput.value ? fromDateInput.value : '';
    const toDate = toDateInput && toDateInput.value ? toDateInput.value : '';
    const nccSearch = nccInput ? nccInput.value.toLowerCase().trim() : '';

    const tableBody = document.getElementById('dhctTableBody');
    if (!tableBody) return;

    const filtered = dhctData.filter(item => {
        const matchSearch = (item.id_dh && item.id_dh.toString().toLowerCase().includes(searchTerm)) ||
            (item.ten && item.ten.toLowerCase().includes(searchTerm)) ||
            (item.id_sp_ct && item.id_sp_ct.toString().toLowerCase().includes(searchTerm));
            
        const itemYMD = toYMD(item.ngay);
        const matchFromDate = !fromDate || itemYMD >= fromDate;
        const matchToDate = !toDate || itemYMD <= toDate;
        const matchNcc = !nccSearch || (item.ncc && item.ncc.toLowerCase().includes(nccSearch));
        
        return matchSearch && matchFromDate && matchToDate && matchNcc;
    }).sort((a, b) => {
        // 1. Ngày lớn tới bé
        const da = toYMD(a.ngay);
        const db = toYMD(b.ngay);
        if (da !== db) return db.localeCompare(da);
        
        // 2. Trường nhập trước xuất sau ("NHẬP" < "XUẤT")
        const truongA = a.truong || '';
        const truongB = b.truong || '';
        if (truongA !== truongB) {
            if (truongA === 'NHẬP') return -1;
            if (truongB === 'NHẬP') return 1;
            return truongA.localeCompare(truongB);
        }
        
        // 3. NCC a-z
        const nccA = a.ncc || '';
        const nccB = b.ncc || '';
        if (nccA !== nccB) return nccA.localeCompare(nccB);
        
        // 4. ID SP CT a-z
        const idSpCtA = a.id_sp_ct || '';
        const idSpCtB = b.id_sp_ct || '';
        return idSpCtA.localeCompare(idSpCtB);
    });

    if (filtered.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="10" class="text-center py-8 text-slate-500">Không tìm thấy kết quả phù hợp.</td></tr>';
        return;
    }

    tableBody.innerHTML = filtered.map(item => `
                <tr class="border-b border-slate-100 hover:bg-slate-50 transition-colors cursor-pointer" ondblclick="window.openDhctModal('${item.id_dh}')" title="Nhấn đúp để sửa hoặc thêm SP">
                    <td class="px-4 py-3 text-sm text-slate-600">${item.ngay || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-600">${item.truong || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-600">${item.ncc || ''}</td>
                    <td class="px-4 py-3 text-sm font-semibold text-slate-900">${item.id_sp_ct || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-600">${item.id_sp || ''}</td>
                    <td class="px-4 py-3 text-sm text-slate-700">${item.ten || ''}</td>
                    <td class="px-4 py-3 text-sm text-right font-bold text-slate-900">${(parseFloat(item.so_luong) || 0).toLocaleString('vi-VN')}</td>
                    <td class="px-4 py-3 text-sm text-right text-slate-600">${(parseFloat(item.gia_nhap) || 0).toLocaleString('vi-VN')} đ</td>
                    <td class="px-4 py-3 text-sm text-right font-bold text-primary">${(parseFloat(item.thanh_tien_nhap) || 0).toLocaleString('vi-VN')} đ</td>
                    <td class="px-4 py-3 text-sm font-medium">
                        ${item.id_sp_ct ? `<button onclick="event.stopPropagation(); window.toggleConfirmRow('${item.id_dh_ct}', '${item.xac_nhan}')" class="px-2 py-1 rounded text-[11px] font-bold transition-colors border ${item.xac_nhan === 'ĐÃ XÁC NHẬN' ? 'border-green-600 text-green-600 hover:bg-green-50' : 'border-amber-500 text-amber-500 hover:bg-amber-50'}">${item.xac_nhan === 'ĐÃ XÁC NHẬN' ? 'ĐÃ XÁC NHẬN' : 'CHỜ XÁC NHẬN'}</button>` : `<span class="text-slate-400 italic text-[11px]">Không có SP CT</span>`}
                    </td>
                </tr>
            `).join('');
}

window.renderDonhangTongTable = function() {
    const searchInput = document.getElementById('filterDHTongSearch');
    const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';
    
    const fromDateInput = document.getElementById('filterDHTongFromDate');
    const toDateInput = document.getElementById('filterDHTongToDate');
    const nccInput = document.getElementById('filterDHTongNcc');
    
    const fromDate = fromDateInput && fromDateInput.value ? fromDateInput.value : '';
    const toDate = toDateInput && toDateInput.value ? toDateInput.value : '';
    const nccSearch = nccInput ? nccInput.value.toLowerCase().trim() : '';

    const tableBody = document.getElementById('dhTongTableBody');
    if (!tableBody) return;

    if (!dhctData || dhctData.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="6" class="text-center py-8 text-slate-500">Chưa tải dữ liệu...</td></tr>';
        return;
    }

    // Group by id_dh
    const grouped = {};
    dhctData.forEach(item => {
        if (!item.id_dh) return;
        
        const itemYMD = toYMD(item.ngay);
        const matchFromDate = !fromDate || itemYMD >= fromDate;
        const matchToDate = !toDate || itemYMD <= toDate;
        const matchNcc = !nccSearch || (item.ncc && item.ncc.toLowerCase().includes(nccSearch));
        const matchSearch = !searchTerm || item.id_dh.toString().toLowerCase().includes(searchTerm);
        
        if (matchFromDate && matchToDate && matchNcc && matchSearch) {
            if (!grouped[item.id_dh]) {
                grouped[item.id_dh] = {
                    id_dh: item.id_dh,
                    ngay: item.ngay,
                    truong: item.truong,
                    ncc: item.ncc,
                    tong_tien: 0,
                    xac_nhan: 'ĐÃ XÁC NHẬN'
                };
            }
            grouped[item.id_dh].tong_tien += (parseFloat(item.thanh_tien_nhap) || 0);
            if (item.xac_nhan !== 'ĐÃ XÁC NHẬN') {
                grouped[item.id_dh].xac_nhan = 'CHỜ XÁC NHẬN';
            }
        }
    });

    const result = Object.values(grouped).sort((a, b) => {
        const da = toYMD(a.ngay);
        const db = toYMD(b.ngay);
        if (da !== db) return db.localeCompare(da);
        return a.id_dh.localeCompare(b.id_dh);
    });

    if (result.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="6" class="text-center py-8 text-slate-500">Không tìm thấy kết quả phù hợp.</td></tr>';
        return;
    }

    tableBody.innerHTML = result.map(item => `
        <tr class="border-b border-slate-100 hover:bg-slate-50 transition-colors cursor-pointer" ondblclick="window.openDhctModal('${item.id_dh}')" title="Nhấn đúp để xem chi tiết">
            <td class="px-4 py-3 text-sm font-semibold text-slate-900">${item.id_dh || ''}</td>
            <td class="px-4 py-3 text-sm text-slate-600">${item.ngay || ''}</td>
            <td class="px-4 py-3 text-sm text-slate-600">${item.truong || ''}</td>
            <td class="px-4 py-3 text-sm text-slate-600">${item.ncc || ''}</td>
            <td class="px-4 py-3 text-sm text-right font-bold text-primary">${item.tong_tien.toLocaleString('vi-VN')} đ</td>
            <td class="px-4 py-3 text-sm text-center font-medium">
                <button onclick="event.stopPropagation(); window.toggleConfirmOrder('${item.id_dh}', '${item.xac_nhan}')" class="px-3 py-1.5 rounded-lg text-xs font-bold transition-colors border ${item.xac_nhan === 'ĐÃ XÁC NHẬN' ? 'bg-green-50 border-green-600 text-green-600 hover:bg-green-100' : 'bg-amber-50 border-amber-500 text-amber-600 hover:bg-amber-100'}">${item.xac_nhan === 'ĐÃ XÁC NHẬN' ? 'ĐÃ XÁC NHẬN' : 'CHỜ XÁC NHẬN'}</button>
            </td>
        </tr>
    `).join('');
};

window.toggleConfirmRow = async function(id_dh_ct, currentStatus) {
    if (!id_dh_ct) return;
    const newStatus = currentStatus === 'ĐÃ XÁC NHẬN' ? 'CHỜ XÁC NHẬN' : 'ĐÃ XÁC NHẬN';
    const row = dhctData.find(r => r.id_dh_ct === id_dh_ct);
    if (!row) return;

    // Optimistic UI update
    row.xac_nhan = newStatus;
    if (document.getElementById('moduleDhct').style.display !== 'none') {
        renderDHCTTable();
    }
    if (document.getElementById('moduleDonhangTong').style.display !== 'none') {
        renderDonhangTongTable();
    }

    // Background API call
    window.updateSheetCell(CONFIG.dhctSheetName, row.rowIndex, 15, newStatus).catch(err => {
        console.error("Lỗi khi cập nhật trạng thái:", err);
        // Revert on error
        row.xac_nhan = currentStatus;
        if (document.getElementById('moduleDhct').style.display !== 'none') renderDHCTTable();
        if (document.getElementById('moduleDonhangTong').style.display !== 'none') renderDonhangTongTable();
        alert("Lỗi khi cập nhật trạng thái: " + err.message);
    });
};

window.toggleConfirmOrder = async function(id_dh, currentStatus) {
    if (!id_dh) return;
    const newStatus = currentStatus === 'ĐÃ XÁC NHẬN' ? 'CHỜ XÁC NHẬN' : 'ĐÃ XÁC NHẬN';
    
    const rowsToUpdate = dhctData.filter(r => r.id_dh === id_dh && r.id_sp_ct);
    if (rowsToUpdate.length === 0) return;

    // Optimistic UI update
    const previousStatuses = rowsToUpdate.map(r => r.xac_nhan);
    rowsToUpdate.forEach(r => r.xac_nhan = newStatus);

    if (document.getElementById('moduleDhct').style.display !== 'none') {
        renderDHCTTable();
    }
    if (document.getElementById('moduleDonhangTong').style.display !== 'none') {
        renderDonhangTongTable();
    }

    // Background API calls in parallel
    Promise.all(rowsToUpdate.map(row => 
        window.updateSheetCell(CONFIG.dhctSheetName, row.rowIndex, 15, newStatus)
    )).catch(err => {
        console.error("Lỗi khi cập nhật trạng thái đơn hàng:", err);
        // Revert on error
        rowsToUpdate.forEach((r, idx) => r.xac_nhan = previousStatuses[idx]);
        if (document.getElementById('moduleDhct').style.display !== 'none') renderDHCTTable();
        if (document.getElementById('moduleDonhangTong').style.display !== 'none') renderDonhangTongTable();
        alert("Lỗi khi cập nhật trạng thái đơn hàng: " + err.message);
    });
};

    Object.assign(window.AppModules = window.AppModules || {}, { ['dhct']: true });
    window.fetchDHCTData = fetchDHCTData;
    window.renderDHCTTable = renderDHCTTable;
})();
