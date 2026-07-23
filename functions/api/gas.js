// ⚠️ GAS 代理層：/api/gas
// 前端不再直接打 script.google.com，而是打這個路徑，由這層 Function 代替前端去呼叫
// 真正的 Google Apps Script 網址。這樣做兩件事原本在純前端做不到的事：
//
// 1. 真正的 GAS 網址不再出現在使用者看得到的網頁原始碼裡（存在 Function 的環境變數）。
// 2. 可以用 Cloudflare KV 記錄「這個 IP 密碼打錯幾次」，超過門檻就直接擋掉，
//    不用再去打 GAS，這是 GAS 本身做不到的 —— Apps Script 的 Web App 觸發器
//    拿不到穩定的來源 IP，沒辦法自己做這件事。
//
// ── 部署前要做的事 ──────────────────────────────────────────────
// A. Cloudflare Pages 專案 > Settings > Environment variables
//    新增一個變數：GAS_WEB_APP_URL = 你們原本 WEB_APP_URL 那個完整網址
//    （builds 用「Production」和「Preview」都要各設一次）
//
// B. Cloudflare Pages 專案 > Settings > Functions > KV namespace bindings
//    建立一個 KV namespace（例如取名 admin-rate-limit），變數名稱綁定為 RATE_LIMIT_KV
//    （Workers & Pages > KV 頁面可以直接建立 namespace，再回來綁定）
// ────────────────────────────────────────────────────────────────

const MAX_FAILED_ATTEMPTS = 10;          // 同一 IP 在下面的時間窗內最多可以打錯幾次密碼
const WINDOW_SECONDS = 15 * 60;          // 15 分鐘的滑動時間窗
const BLOCK_SECONDS = 15 * 60;           // 超過門檻後，封鎖這個 IP 多久

// ⚠️ 已知限制，先寫清楚，之後如果覺得不夠再調整：
// 1. 你們沒有帳號系統，只有一組共用密碼，所以這裡只能鎖 IP，鎖不了「帳號」。
//    店裡 WiFi 底下所有手機共用一個對外 IP，極端情況下一個人連續打錯密碼，
//    會連累同店其他人也被鎖 15 分鐘。權衡下這是內部工具可接受的代價——
//    真要做到「只鎖那個人」，得先做帳號系統，成本不成比例。
// 2. Cloudflare KV 是最終一致性，不是原子計數。如果有人短時間內平行送出
//    大量請求（例如寫程式一次送 20 個），讀到的計數可能還沒更新，導致
//    超過門檻才真正生效。這在「內部後台」的威脅情境下風險可接受；
//    如果之後想補到滴水不漏，要接 Cloudflare Turnstile 做人機驗證，
//    是額外一塊工程，需要的話再說。

function getClientIp(request) {
    // Cloudflare 會自動附加這個 header，是目前這一層能拿到的、比較可靠的來源 IP，
    // 這是 GAS 直接被前端呼叫時完全拿不到的資訊。
    return request.headers.get('CF-Connecting-IP') || 'unknown';
}

async function isBlocked(kv, ip) {
    if (!kv) return false; // 沒設定 KV 就跳過限流，不要讓忘記設定變成整個系統打不開
    const raw = await kv.get(`block:${ip}`);
    return raw !== null;
}

async function recordFailure(kv, ip) {
    if (!kv) return;
    const key = `fail:${ip}`;
    const raw = await kv.get(key);
    const count = raw ? parseInt(raw, 10) + 1 : 1;
    await kv.put(key, String(count), { expirationTtl: WINDOW_SECONDS });
    if (count >= MAX_FAILED_ATTEMPTS) {
        await kv.put(`block:${ip}`, '1', { expirationTtl: BLOCK_SECONDS });
    }
}

async function clearFailures(kv, ip) {
    if (!kv) return;
    await kv.delete(`fail:${ip}`);
}

function jsonResponse(obj, status) {
    return new Response(JSON.stringify(obj), {
        status: status || 200,
        headers: { 'Content-Type': 'application/json; charset=utf-8' }
    });
}

export async function onRequest(context) {
    const { request, env } = context;
    const kv = env.RATE_LIMIT_KV;
    const gasUrl = env.GAS_WEB_APP_URL;

    if (!gasUrl) {
        // 部署時忘記設定環境變數，回明確錯誤，不要讓它安靜地失敗
        return jsonResponse({ status: 'error', message: '伺服器尚未設定 GAS_WEB_APP_URL，請聯絡管理員' }, 500);
    }

    const ip = getClientIp(request);

    if (await isBlocked(kv, ip)) {
        return jsonResponse({ status: 'error', message: '嘗試次數過多，請稍後再試（約 15 分鐘後解除）' }, 429);
    }

    const incomingUrl = new URL(request.url);

    try {
        let gasResp;
        if (request.method === 'POST') {
            const body = await request.text();
            gasResp = await fetch(gasUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body
            });
        } else {
            const target = gasUrl + incomingUrl.search; // 把原本的 ?action=...&password=... 原封不動轉發
            gasResp = await fetch(target, { method: 'GET' });
        }

        const text = await gasResp.text();
        let data;
        try { data = JSON.parse(text); } catch (e) { data = null; }

        // 只計「密碼錯誤」這種失敗，不要把其他業務錯誤（例如查無資料）也算進暴力破解次數。
        // 你們 Code.gs 那邊密碼錯誤時的 message 目前都帶「密碼」兩個字，用這個保守判斷。
        const looksLikeAuthFailure = data && data.status !== 'success' &&
            typeof data.message === 'string' && data.message.includes('密碼');

        if (looksLikeAuthFailure) {
            await recordFailure(kv, ip);
        } else if (data && data.status === 'success') {
            await clearFailures(kv, ip);
        }

        return new Response(text, {
            status: gasResp.status,
            headers: { 'Content-Type': 'application/json; charset=utf-8' }
        });
    } catch (err) {
        return jsonResponse({ status: 'error', message: '代理伺服器連線失敗：' + err.message }, 502);
    }
}
