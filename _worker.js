// ⚠️ Workers with Assets 架構：main = "_worker.js"，這個檔案是唯一的進入點，
// functions/ 資料夾不會被執行，所以裝置門禁、密碼限流、安全 headers 全部整合在這裡。
//
// 需要的 binding / 變數（不要寫進 wrangler.toml 的 [vars]，那是明文、會進 git）：
// - env.GAS_WEB_APP_URL  → 用 Secret 設定（Dashboard: Settings > Variables > 選 "Secret" 類型；
//                           或本機用 `npx wrangler secret put GAS_WEB_APP_URL`）
// - env.RATE_LIMIT_KV    → 用 wrangler.toml 的 [[kv_namespaces]] 綁定
//
// ★★★ 2026-07-23：暫時拿掉「不是手機」「不是LINE內建瀏覽器」這兩道門禁，
// 原因：客人在LINE App裡點LIFF連結卻一直被導去Google，一直查不出是LIFF
// Endpoint URL設定、部署環境、還是UA判斷式本身的問題，先把這兩道檢查拿掉
// 讓系統能正常運作、不擋到任何人，之後查出根本原因、確認沒問題了，
// 建議再把這兩道檢查加回來（否則任何人都能用電腦瀏覽器直接開這個網址，
// 繞過LINE身份驗證的前提，不是嚴重漏洞但失去了原本這道防線的用意）。

const MOBILE_UA_REGEX = /Mobi|Android|iPhone|iPad|iPod|Windows Phone|BlackBerry/i;
const MAX_FAILED_ATTEMPTS = 10;
const WINDOW_SECONDS = 15 * 60;
const BLOCK_SECONDS = 15 * 60;

function jsonResponse(obj, status) {
    return new Response(JSON.stringify(obj), {
        status: status || 200,
        headers: { 'Content-Type': 'application/json; charset=utf-8' }
    });
}

function withSecurityHeaders(response) {
    const newHeaders = new Headers(response.headers);
    newHeaders.set('X-Frame-Options', 'DENY');
    newHeaders.set('Cache-Control', 'no-store, no-cache, must-revalidate');
    // ⚠️ 2026-07-21 拿掉 Content-Security-Policy：客人點餐頁面的 LINE 分享功能
    // （liff.shareTargetPicker）需要跟 LINE 官方網域互動，CSP 限制頁面只能跟
    // 自己網域溝通，會讓分享按鈕點下去沒反應。這個問題之前在 index.html 自己的
    // 程式碼歷史裡就出現過一次、也修過一次（拿掉 CSP），這次是在 _worker.js
    // 加安全 headers 時不小心重新引入了同樣的問題，教訓：CSP 對這類會跟外部
    // 服務（LINE、金流等）互動的頁面風險較高，加之前要先確認會不會擋到功能。
    return new Response(response.body, {
        status: response.status,
        statusText: response.statusText,
        headers: newHeaders
    });
}

function getIp(request) {
    return request.headers.get('CF-Connecting-IP') || 'unknown';
}

// ★ 2026-07-21 新增：限流改用「已驗證的 LINE 身份」而不是 IP，原因是同店
// WiFi 底下所有裝置共用一個對外 IP，用 IP 限流會連累同店其他人。但這裡
// 不能直接信任前端說「我是誰」——前端傳來的身份宣稱可以被偽造，如果直接
// 採信，等於形同虛設，有心人只要每次假裝不同身份就能繞過限流。
// 所以這裡由 Cloudflare 自己直接打 LINE 官方的驗證 API，獨立確認這個
// idToken 真的是 LINE 核發的、而且真的核發給這個 LIFF 用的，不透過 Google
// Apps Script那一層，也不相信前端的任何說法。驗證失敗（沒帶 idToken、
// token 過期、不是這個LIFF核發的…）就安全退回用 IP 限流，不會因為這層
// 附加機制故障就讓密碼登入整個用不了。
async function verifyLineIdTokenWorker(idToken, channelId) {
    if (!idToken || !channelId) return null;
    try {
        const resp = await fetch('https://api.line.me/oauth2/v2.1/verify', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: `id_token=${encodeURIComponent(idToken)}&client_id=${encodeURIComponent(channelId)}`
        });
        if (resp.status !== 200) return null;
        const data = await resp.json();
        return data.sub || null;
    } catch (err) {
        return null;
    }
}

async function getRateLimitKey(request, env) {
    // 先試著從請求裡拿 idToken（GET 帶在網址參數，POST 帶在 JSON body 裡），
    // 驗證成功就用「line:真實UserID」當 key，驗證不到才退回用 IP。
    const incomingUrl = new URL(request.url);
    let idToken = incomingUrl.searchParams.get('idToken') || '';
    if (!idToken && request.method === 'POST') {
        try {
            const cloned = request.clone();
            const bodyText = await cloned.text();
            const bodyJson = JSON.parse(bodyText);
            idToken = bodyJson.idToken || '';
        } catch (e) { /* 不是JSON或沒有idToken欄位，忽略 */ }
    }
    if (idToken && env.LINE_CHANNEL_ID) {
        const verifiedUserId = await verifyLineIdTokenWorker(idToken, env.LINE_CHANNEL_ID);
        if (verifiedUserId) return 'line:' + verifiedUserId;
    }
    return 'ip:' + getIp(request);
}

async function isBlocked(kv, ip) {
    if (!kv) return false;
    return (await kv.get(`block:${ip}`)) !== null;
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

async function handleApi(request, env) {
    const kv = env.RATE_LIMIT_KV;
    const gasUrl = env.GAS_WEB_APP_URL;
    const rateLimitKey = await getRateLimitKey(request, env);

    if (!gasUrl) {
        return jsonResponse({ status: 'error', message: '伺服器尚未設定 GAS_WEB_APP_URL，請聯絡管理員' }, 500);
    }

    if (await isBlocked(kv, rateLimitKey)) {
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
            gasResp = await fetch(gasUrl + incomingUrl.search, { method: 'GET' });
        }

        const text = await gasResp.text();
        let data;
        try { data = JSON.parse(text); } catch (e) { data = null; }

        const looksLikeAuthFailure = data && data.status !== 'success' &&
            typeof data.message === 'string' &&
            (data.message.includes('密碼') || data.message.includes('登入') || data.message.includes('權限'));

        if (looksLikeAuthFailure) {
            await recordFailure(kv, rateLimitKey);
        } else if (data && data.status === 'success') {
            await clearFailures(kv, rateLimitKey);
        }

        return new Response(text, {
            status: gasResp.status,
            headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }
        });
    } catch (err) {
        return jsonResponse({ status: 'error', message: '代理伺服器連線失敗：' + err.message }, 502);
    }
}

export default {
    async fetch(request, env) {
        const url = new URL(request.url);

        // /api/ 底下走代理 + 限流邏輯；不做 UA 擋（密碼本身是驗證），
        // 但一樣套用安全 headers
        if (url.pathname.startsWith('/api/')) {
            const resp = await handleApi(request, env);
            return withSecurityHeaders(resp);
        }

        // ★★★ 2026-07-23：原本這裡有「不是手機格式」「不是LINE內建瀏覽器」
        // 兩道檢查，沒過就直接302導去Google。現在暫時拿掉，任何裝置、任何
        // 瀏覽器都能直接看到頁面（含 /admin，kiosk 金鑰通道也還在，繼續有效，
        // 只是現在就算沒帶kioskKey也不會被擋）。
        //
        // 拿掉之後的影響：這兩道原本是「裝置層級」的第一道防線，拿掉後不代表
        // 系統不安全——真正保護資料的是後面 idToken 驗證、後台密碼、
        // LINE 身份白名單這些機制，都還在正常運作，沒有被動到。只是現在
        // 電腦瀏覽器也能直接開到點餐頁面/後台登入畫面（但沒有密碼/身份
        // 一樣進不去實際功能），跟之前「电脑打开直接跳Google、看起来像
        // 网址不存在」的隐蔽效果不一样了。

        // 通過檢查（或是合法的 kiosk 電腦），交給 Assets 服務靜態檔案
        // 路由規則：
        //   /        → index.html（客人點餐頁面，Cloudflare 預設就會找 index.html，不用改寫）
        //   /admin   → admin_hub_basic.html（後台管理，要手動改寫路徑）
        if (env.ASSETS) {
            let assetRequest = request;
            if (url.pathname === '/admin' || url.pathname === '/admin/') {
                const rewrittenUrl = new URL(request.url);
                rewrittenUrl.pathname = '/admin_hub_basic.html';
                assetRequest = new Request(rewrittenUrl.toString(), request);
            }
            const resp = await env.ASSETS.fetch(assetRequest);
            return withSecurityHeaders(resp);
        }

        return withSecurityHeaders(new Response('Worker live, /api/ ready', { status: 200 }));
    }
};
