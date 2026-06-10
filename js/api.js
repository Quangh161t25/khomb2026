// api - Module Pattern (IIFE)
(function () {
// tokenRequestPromise is declared in js/state.js
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
        const rowLimit = sheetName === CONFIG.udctSheetName ? 100000 : 10000;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!A1:AF${rowLimit}`;
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
        const rowLimit = sheetName === CONFIG.udctSheetName ? 100000 : 10000;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.spreadsheetId}/values/${sheetName}!A2:AF${rowLimit}:clear`;
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



    Object.assign(window.AppModules = window.AppModules || {}, { ['api']: true });
    window.getAccessToken = getAccessToken;
    window.fetchSheetData = fetchSheetData;
    window.fetchAuthData = fetchAuthData;
    window.clearSheetData = clearSheetData;
    window.appendSheetData = appendSheetData;
    window.updateSheetCell = updateSheetCell;
    window.updateSheetValue = updateSheetValue;
})();
