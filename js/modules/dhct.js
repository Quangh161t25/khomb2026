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
                ghi_chu: row[5], // Cột F
                kho: row[6],     // Cột G
                id_sp_ct: row[7], // Cột H
                ten: row[10],     // Cột K
                so_luong: row[11], // Cột L
                gia_nhap: row[12]  // Cột M
            }));
            if (silent) {
                // Cập nhật bảng nếu không đang mở drawer
                const isDHCTModule = pageTitle.textContent === 'Dữ liệu DH Chi Tiết' || pageTitle.textContent === 'Đơn hàng trên DH Chi Tiết';
                if (isDHCTModule) {
                    renderDHCTTable();
                    renderUniqueDHCTTable();
                }
            } else {
                renderDHCTTable();
                renderUniqueDHCTTable();
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


    Object.assign(window.AppModules = window.AppModules || {}, { ['dhct']: true });
    window.fetchDHCTData = fetchDHCTData;
    window.renderDHCTTable = renderDHCTTable;
})();
