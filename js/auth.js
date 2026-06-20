// auth - Module Pattern (IIFE)
(function () {
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

    document.querySelectorAll("button[onclick='updateAllPricesBatch()'], button[onclick^='batchUpdateUDCTStatus'], button[onclick='batchRefreshSelectedRows()']").forEach(b => {
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

    const loadingEl = document.getElementById('loginLoading');
    const loginBtn = document.querySelector("button[onclick='handleLogin()']");
    if (loadingEl) loadingEl.classList.remove('hidden');
    if (loginBtn) { loginBtn.disabled = true; loginBtn.textContent = 'Đang xác thực...'; }

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

        let nextUrl = 'index.html';
        const urlParams = new URLSearchParams(window.location.search);
        const link = urlParams.get('link');
        if (link) {
            nextUrl += '?link=' + link;
        }
        window.location.href = nextUrl;
    } catch (err) {
        console.error('Login error:', err);
        if (errorEl) {
            errorEl.textContent = 'Không thể xác thực tài khoản. Vui lòng thử lại.';
            errorEl.classList.remove('hidden');
        }
    } finally {
        if (loadingEl) loadingEl.classList.add('hidden');
        if (loginBtn) { loginBtn.disabled = false; loginBtn.textContent = 'Đăng nhập'; }
    }
}

function handleLogout() {
    if (confirm("Bạn có muốn đăng xuất không?")) {
        localStorage.removeItem('erp_current_user');
        isLoggedIn = false;
        currentUser = null;
        window.location.href = 'login.html';
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

// ═══════════════════════════════════════════════════════════════
// OFFLINE SYNC: Hàng đợi thao tác khi mất mạng
// ═══════════════════════════════════════════════════════════════
function getOfflineQueue() {
    try { return JSON.parse(localStorage.getItem('erp_offline_queue') || '[]'); } catch { return []; }
}
function saveOfflineQueue(q) {
    localStorage.setItem('erp_offline_queue', JSON.stringify(q));
}
function enqueueOfflineOp(op) {
    const q = getOfflineQueue();
    op.ts = Date.now();
    q.push(op);
    saveOfflineQueue(q);
    showOfflineBadge(q.length);
    console.log('[Offline] Đã lưu thao tác vào hàng đợi:', op);
}
window.enqueueOfflineOp = enqueueOfflineOp;

function showOfflineBadge(count) {
    let badge = document.getElementById('offlineSyncBadge');
    if (!badge) {
        badge = document.createElement('div');
        badge.id = 'offlineSyncBadge';
        badge.style.cssText = 'position:fixed;bottom:16px;left:50%;transform:translateX(-50%);z-index:9999;background:#f59e0b;color:#fff;padding:6px 16px;border-radius:999px;font-size:12px;font-weight:700;box-shadow:0 2px 12px rgba(0,0,0,0.2);pointer-events:none;transition:opacity 0.3s';
        document.body.appendChild(badge);
    }
    if (count > 0) {
        badge.textContent = `⏳ ${count} thao tác chờ đồng bộ`;
        badge.style.opacity = '1';
        badge.style.display = 'block';
    } else {
        badge.textContent = '✅ Đã đồng bộ xong';
        badge.style.background = '#10b981';
        badge.style.opacity = '1';
        setTimeout(() => { badge.style.opacity = '0'; setTimeout(() => { badge.style.display = 'none'; badge.style.background = '#f59e0b'; }, 300); }, 2000);
    }
}

async function flushOfflineQueue() {
    const q = getOfflineQueue();
    if (q.length === 0) return;
    if (!navigator.onLine) return;

    console.log(`[Offline] Đang đồng bộ ${q.length} thao tác...`);
    showToast(`Đang đồng bộ ${q.length} thao tác offline...`, 'info');

    const failed = [];
    for (const op of q) {
        try {
            let ok = false;
            if (op.type === 'updateCell') {
                ok = await updateSheetCell(op.sheetName, op.rowIndex, op.colIndex, op.value);
            } else if (op.type === 'appendRow') {
                ok = await appendSheetData(op.sheetName, op.values);
            } else if (op.type === 'updateRow') {
                ok = await updateSheetRow(op.sheetName, op.rowIndex, op.rowData);
            }
            if (!ok) failed.push(op);
        } catch (e) {
            console.error('[Offline] Lỗi khi đồng bộ:', e, op);
            failed.push(op);
        }
    }

    saveOfflineQueue(failed);
    showOfflineBadge(failed.length);

    if (failed.length === 0) {
        showToast('Đồng bộ hoàn tất!', 'success');
        // Refresh data sau khi đồng bộ
        if (isLoggedIn) {
            loadUDCTData(true);
            fetchDHCTData(true);
        }
    } else {
        showToast(`${failed.length} thao tác chưa đồng bộ được, sẽ thử lại sau.`, 'warning');
    }
}
window.flushOfflineQueue = flushOfflineQueue;

// Lắng nghe khi có mạng trở lại
window.addEventListener('online', () => {
    console.log('[Network] Đã có mạng, đồng bộ...');
    showToast('Đã kết nối mạng. Đang đồng bộ dữ liệu...', 'info');
    setTimeout(() => flushOfflineQueue(), 1500);
});

window.addEventListener('offline', () => {
    console.log('[Network] Mất mạng, chuyển sang offline mode');
    showToast('Mất kết nối mạng. Thao tác sẽ được lưu và đồng bộ khi có mạng.', 'warning');
});

// ═══════════════════════════════════════════════════════════════
// KHỞI ĐỘNG ỨNG DỤNG
// ═══════════════════════════════════════════════════════════════
window.onload = async () => {
    const isLoginPage = window.location.pathname.endsWith('login.html');
    const loginScreen = document.getElementById('loginScreen');
    const mainApp = document.getElementById('mainApp');
    const lastId = localStorage.getItem('erp_last_user_id');
    if (lastId && document.getElementById('loginId')) document.getElementById('loginId').value = lastId;

    // Kiểm tra hàng đợi offline còn tồn đọng
    const offlineQ = getOfflineQueue();
    if (offlineQ.length > 0) showOfflineBadge(offlineQ.length);

    try {
        const savedUser = localStorage.getItem('erp_current_user');
        if (savedUser) {
            const parsedUser = JSON.parse(savedUser);
            if (parsedUser && parsedUser.id) {
                if (isLoginPage) {
                    window.location.href = 'index.html';
                    return;
                }
                // Hiện app ngay lập tức (không chờ mạng)
                if (loginScreen) loginScreen.classList.add('hidden');
                if (mainApp) mainApp.classList.remove('hidden');
                isLoggedIn = true;
                currentUser = parsedUser;
                if (document.getElementById('welcomeUserName')) document.getElementById('welcomeUserName').textContent = currentUser.name;
                updateUserProfileUI();
                applyRoleUI(currentUser.role);

                // Khởi động tự động làm mới
                startAutoRefresh();

                // Tải dữ liệu ban đầu (hoạt động cả offline qua cache)
                loadUDCTData();
                loadSanphamData();
                fetchDHCTData();

                // Chuyển module theo URL nếu có
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

                // Xác minh session trong nền (CHỈ logout nếu có mạng VÀ user không hợp lệ)
                try {
                    if (navigator.onLine) {
                        await fetchAuthData();
                        const user = usersData.find(u => u.id === parsedUser.id && u.password === parsedUser.password);
                        if (user && !(user.tinhTrang && /nghỉ|khóa|block|inactive/i.test(user.tinhTrang))) {
                            currentUser = {
                                id: user.id,
                                name: user.name || user.id,
                                role: (user.role || 'user').toString().trim().toLowerCase(),
                                password: user.password
                            };
                            localStorage.setItem('erp_current_user', JSON.stringify(currentUser));
                            updateUserProfileUI();
                            applyRoleUI(currentUser.role);
                            await loadUDCTData();
                            await loadSanphamData();
                        } else if (user === undefined) {
                            // Không tìm thấy user → tài khoản bị xóa
                            handleLogout();
                        }
                        // Nếu usersData rỗng (mạng lỗi) → giữ nguyên session
                    }
                } catch (bgErr) {
                    // Lỗi mạng khi xác minh ngầm → KHÔNG logout, tiếp tục offline
                    console.warn('[Auth] Xác minh ngầm thất bại (offline?):', bgErr.message);
                }

            } else {
                if (!isLoginPage) window.location.href = 'login.html';
                if (loginScreen) loginScreen.classList.remove('hidden');
                if (mainApp) mainApp.classList.add('hidden');
            }
        } else {
            if (!isLoginPage) window.location.href = 'login.html';
            if (loginScreen) loginScreen.classList.remove('hidden');
            if (mainApp) mainApp.classList.add('hidden');
        }
    } catch (e) {
        console.error("Lỗi khởi tạo:", e);
        // Nếu lỗi nhưng CÓ savedUser → tiếp tục offline thay vì redirect login
        const savedUser = localStorage.getItem('erp_current_user');
        if (!isLoginPage && !savedUser) {
            window.location.href = 'login.html';
        } else if (!isLoginPage && savedUser) {
            // Giữ nguyên, đã hiện app ở trên
            console.warn('[Auth] Lỗi nhưng có cached user, tiếp tục offline');
        }
    }

    if (window.innerWidth > 1024) closeMobileSidebar();
};

function startAutoRefresh() {
    console.log("Auto refresh started (2 mins)");
    setInterval(async () => {
        if (isLoggedIn && navigator.onLine) {
            console.log("Running auto background refresh...");
            await loadUDCTData(true);
            await fetchDHCTData(true);
        }
    }, 120000);
}

window.addEventListener('resize', () => {
    if (window.innerWidth > 1024) closeMobileSidebar();
});

    Object.assign(window.AppModules = window.AppModules || {}, { ['auth']: true });
    window.applyRoleUI = applyRoleUI;
    window.handleLogin = handleLogin;
    window.handleLogout = handleLogout;
    window.updateUserProfileUI = updateUserProfileUI;
    window.startAutoRefresh = startAutoRefresh;
    window.enqueueOfflineOp = enqueueOfflineOp;
    window.flushOfflineQueue = flushOfflineQueue;
})();
