// api - Module Pattern (IIFE)
(function () {
// tokenRequestPromise is declared in js/state.js

// ═══════════════════════════════════════════════════════════════
// CACHE helpers (read cache cho offline)
// ═══════════════════════════════════════════════════════════════
const CACHE_PREFIX = 'erp_cache_';
const CACHE_TTL = 24 * 60 * 60 * 1000; // 24h

function cacheSet(key, data) {
    try {
        localStorage.setItem(CACHE_PREFIX + key, JSON.stringify({ ts: Date.now(), data }));
    } catch (e) {
        // localStorage đầy → bỏ qua
    }
}

function cacheGet(key) {
    try {
        const raw = localStorage.getItem(CACHE_PREFIX + key);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        if (Date.now() - parsed.ts > CACHE_TTL) return null;
        return parsed.data;
    } catch { return null; }
}

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
                throw err;
            }
        } finally {
            tokenRequestPromise = null;
        }
    })();
    return tokenRequestPromise;
}

async function fetchSheetData(sheetName) {
    // Thử lấy từ mạng, fallback về cache nếu offline
    if (!navigator.onLine) {
        const cached = cacheGet('sheet_' + sheetName);
        if (cached) {
            console.log(`[Cache] Dùng cache offline cho ${sheetName}`);
            return cached;
        }
        return [];
    }
    try {
        const token = await getAccessToken();
        if (!token) {
            const cached = cacheGet('sheet_' + sheetName);
            return cached || [];
        }
        const range = typeof getSheetRange === 'function'
            ? getSheetRange(sheetName, 'read')
            : `A1:AF${sheetName === CONFIG.udctSheetName ? 100000 : 10000}`;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!${range}`;
        const resp = await fetch(url, { headers: { "Authorization": `Bearer ${token}` } });
        if (!resp.ok) {
            console.error(`Fetch ${sheetName} failed:`, resp.status);
            const cached = cacheGet('sheet_' + sheetName);
            return cached || [];
        }
        const data = await resp.json();
        const values = data.values || [];
        // Lưu vào cache
        if (values.length > 0) cacheSet('sheet_' + sheetName, values);
        return values;
    } catch (err) {
        console.error(`Fetch ${sheetName} error:`, err);
        const cached = cacheGet('sheet_' + sheetName);
        if (cached) {
            console.log(`[Cache] Fallback cache cho ${sheetName} do lỗi mạng`);
            return cached;
        }
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
    if (!navigator.onLine) {
        console.warn('[Offline] clearSheetData bị bỏ qua khi offline');
        return false;
    }
    try {
        const token = await getAccessToken();
        if (!token) return false;
        const range = typeof getSheetRange === 'function'
            ? getSheetRange(sheetName, 'clear')
            : `A2:AF${sheetName === CONFIG.udctSheetName ? 100000 : 10000}`;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!${range}:clear`;
        const resp = await fetch(url, {
            method: "POST",
            headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" }
        });
        if (!resp.ok) {
            const errorText = await resp.text();
            console.error("Lỗi khi xóa dữ liệu:", errorText);
            return false;
        }
        return true;
    } catch (err) {
        console.error("Lỗi clearSheetData:", err);
        return false;
    }
}

async function appendSheetData(sheetName, values) {
    if (!navigator.onLine) {
        // Đưa vào hàng đợi offline
        if (typeof enqueueOfflineOp === 'function') {
            enqueueOfflineOp({ type: 'appendRow', sheetName, values });
        }
        return 'offline';
    }
    try {
        const token = await getAccessToken();
        if (!token) return false;
        const range = typeof getSheetRange === 'function' ? getSheetRange(sheetName, 'append') : 'A:A';
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!${range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
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
        return true;
    } catch (err) {
        console.error("Lỗi appendSheetData:", err);
        if (typeof enqueueOfflineOp === 'function') {
            enqueueOfflineOp({ type: 'appendRow', sheetName, values });
            return 'offline';
        }
        return false;
    }
}

async function updateSheetCell(sheetName, rowIndex, colIndex, value) {
    if (!navigator.onLine) {
        if (typeof enqueueOfflineOp === 'function') {
            enqueueOfflineOp({ type: 'updateCell', sheetName, rowIndex, colIndex, value });
        }
        return 'offline';
    }
    try {
        const token = await getAccessToken();
        if (!token) return false;

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

        if (!resp.ok) {
            if (typeof enqueueOfflineOp === 'function') {
                enqueueOfflineOp({ type: 'updateCell', sheetName, rowIndex, colIndex, value });
                return 'offline';
            }
        }
        return resp.ok;
    } catch (err) {
        console.error("Lỗi updateSheetCell:", err);
        if (typeof enqueueOfflineOp === 'function') {
            enqueueOfflineOp({ type: 'updateCell', sheetName, rowIndex, colIndex, value });
            return 'offline';
        }
        return false;
    }
}

async function updateSheetValue(sheetName, range, value) {
    if (!navigator.onLine) {
        if (typeof enqueueOfflineOp === 'function') {
            enqueueOfflineOp({ type: 'updateCell', sheetName, rowIndex: range, colIndex: 1, value });
        }
        return 'offline';
    }
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

async function updateSheetRow(sheetName, rowIndex, rowDataArray) {
    if (!navigator.onLine) {
        if (typeof enqueueOfflineOp === 'function') {
            enqueueOfflineOp({ type: 'updateRow', sheetName, rowIndex, rowData: rowDataArray });
        }
        return 'offline';
    }
    try {
        const token = await getAccessToken();
        if (!token) return false;
        
        const range = `${sheetName}!A${rowIndex}`;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${range}?valueInputOption=USER_ENTERED`;
        
        const resp = await fetch(url, {
            method: "PUT",
            headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ values: [rowDataArray] })
        });
        
        if (!resp.ok) {
            const errorText = await resp.text();
            console.error("Lỗi khi cập nhật dòng:", errorText);
            if (typeof enqueueOfflineOp === 'function') {
                enqueueOfflineOp({ type: 'updateRow', sheetName, rowIndex, rowData: rowDataArray });
                return 'offline';
            }
            return false;
        }
        return true;
    } catch (err) {
        console.error("Lỗi updateSheetRow:", err);
        if (typeof enqueueOfflineOp === 'function') {
            enqueueOfflineOp({ type: 'updateRow', sheetName, rowIndex, rowData: rowDataArray });
            return 'offline';
        }
        return false;
    }
}

    Object.assign(window.AppModules = window.AppModules || {}, { ['api']: true });
    window.getAccessToken = getAccessToken;
    window.fetchSheetData = fetchSheetData;
    window.fetchAuthData = fetchAuthData;
    window.clearSheetData = clearSheetData;
    window.appendSheetData = appendSheetData;
    window.updateSheetCell = updateSheetCell;
    window.updateSheetValue = updateSheetValue;
    window.updateSheetRow = updateSheetRow;
    window.cacheSet = cacheSet;
    window.cacheGet = cacheGet;
})();
