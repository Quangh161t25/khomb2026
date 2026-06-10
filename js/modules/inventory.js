// inventory - Module Pattern (IIFE)
(function () {
let inventoryCurrentPage = 1;
const INVENTORY_PER_PAGE = 100;
let filteredInventoryData = [];

// Logic module Tồn Kho
async function fetchInventoryData() {
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.classList.remove('hidden');

    try {
        const token = await getAccessToken();
        const urlInventory = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.inventorySheetName}!A:K`;
        const urlDHCT = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${CONFIG.dhctSheetName}!A:P`;

        const [respInv, respDHCT] = await Promise.all([
            fetch(urlInventory, { headers: { "Authorization": `Bearer ${token}` } }),
            fetch(urlDHCT, { headers: { "Authorization": `Bearer ${token}` } })
        ]);

        const resultInv = await respInv.json();
        const resultDHCT = await respDHCT.json();

        // Xử lý Ledger (DH_CT) - Sử dụng đúng cấu trúc chuẩn của dhctData
        if (resultDHCT.values && resultDHCT.values.length > 0) {
            dhctData = resultDHCT.values.slice(1).map((row, idx) => ({
                rowIndex: idx + 2,
                id_dh_ct: (row[0] || '').toString().trim(),
                id_dh: row[1],
                ngay: row[2],
                truong: (row[3] || '').toString().trim(),
                ncc: row[4],
                ghi_chu: row[5],
                kho: (row[6] || 'KHO').toString().trim(),
                id_sp_ct: (row[7] || '').toString().trim(),
                ten: row[10],
                so_luong: parseFloat(row[11]) || 0,
                gia_nhap: parseFloat(row[12]) || 0
            }));
        }

        if (resultInv.values && resultInv.values.length > 0) {
            inventoryData = resultInv.values.slice(1);
            inventoryCurrentPage = 1;
            filterInventory();
        } else {
            document.getElementById('inventoryTableBody').innerHTML = '<tr><td colspan="7" class="text-center py-8 text-slate-500">Không có dữ liệu tồn kho.</td></tr>';
        }
    } catch (err) {
        console.error("Lỗi tải tồn kho:", err);
        alert("Không thể tải dữ liệu tồn kho.");
    } finally {
        loadingOverlay.classList.add('hidden');
    }
}

function renderInventory() {
    const tbody = document.getElementById('inventoryTableBody');
    if (!tbody) return;

    // Tổng hợp Nhập/Xuất từ dhctData (Sử dụng cấu trúc so_luong chuẩn)
    const ledgerSums = {};
    if (dhctData && dhctData.length > 0) {
        dhctData.forEach(item => {
            const khoKey = (item.kho || 'KHO').trim().toUpperCase();
            const spKey = (item.id_sp_ct || '').trim().toUpperCase();
            const key = `${khoKey}_${spKey}`;
            const loai = (item.truong || '').trim().toUpperCase();

            if (!ledgerSums[key]) ledgerSums[key] = { nhap: 0, xuat: 0 };
            const sl = parseFloat(item.so_luong) || 0;

            if (loai === 'NHẬP') ledgerSums[key].nhap += sl;
            else if (loai === 'XUẤT') ledgerSums[key].xuat += sl;
        });
    }

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
        const khoName = (row[1] || 'KHO').toString().trim().toUpperCase();
        const idSpCt = (row[2] || '').toString().trim().toUpperCase();
        const key = `${khoName}_${idSpCt}`;

        const ton_dau = parseFloat(row[7]) || 0;
        const nhap = ledgerSums[key]?.nhap || 0;
        const xuat = ledgerSums[key]?.xuat || 0;
        const ton_cuoi = ton_dau + nhap - xuat;

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


    Object.assign(window.AppModules = window.AppModules || {}, { ['inventory']: true });
    window.fetchInventoryData = fetchInventoryData;
    window.renderInventory = renderInventory;
    window.filterInventory = filterInventory;
    window.changeInventoryPage = changeInventoryPage;
})();
