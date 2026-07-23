// ⚠️ Workers with Assets 架構：main = "_worker.js"，這個檔案是唯一的進入點，
// functions/ 資料夾不會被執行，所以裝置門禁、密碼限流、安全 headers 全部整合在這裡。
//
// 需要的 binding / 變數（不要寫進 wrangler.toml 的 [vars]，那是明文、會進 git）：
// - env.GAS_WEB_APP_URL  → 用 Secret 設定（Dashboard: Settings > Variables > 選 "Secret" 類型；
//                           或本機用 `npx wrangler secret put GAS_WEB_APP_URL`）
// - env.RATE_LIMIT_KV    → 用 wrangler.toml 的 [[kv_namespaces]] 綁定

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

function mobileBlockedResponse() {
    // ⚠️ 2026-07-21 改成直接跳轉 Google，不顯示「僅限手機瀏覽」這種提示文字。
    // 原本的提示文字等於告訴對方「這裡確實有東西、只是被擋住」，反而引誘
    // 好奇的人繼續嘗試繞過；改成跳轉 Google，電腦打開這個網址時，
    // 看起來就像網址打錯、或這裡什麼都沒有，不會讓人知道背後其實有系統。
    return new Response(null, {
        status: 302,
        headers: { 'Location': 'https://www.google.com' }
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
        const ua = request.headers.get('User-Agent') || '';
        const isMobile = MOBILE_UA_REGEX.test(ua);

        // ★ 2026-07-22 新增：裝置門禁改成可以透過網頁開關設定，不用再叫
        // 開發者改程式碼、重新部署才能調整。用RATE_LIMIT_KV存一個開關值
        // （這個KV本來就有綁定，借來用不用另外多開一個KV namespace）。
        // 目前預設「解除限制」（沒設定過的話，KV裡讀不到值，視為關閉），
        // 電腦、一般瀏覽器都能直接開，不會被擋轉去Google。
        // ⚠️ 這段一定要放在下面「/api/ 開頭都轉去GAS代理」的判斷之前，
        // 不然/api/device-restriction這個路徑本身也是/api/開頭，會先被
        // 那條規則攔截、永遠執行不到這裡。
        if (url.pathname === '/api/device-restriction') {
            const kv = env.RATE_LIMIT_KV;
            if (request.method === 'GET') {
                const enabled = kv ? (await kv.get('device_restriction_enabled')) === 'true' : false;
                return withSecurityHeaders(jsonResponse({ status: 'success', enabled }));
            }
            if (request.method === 'POST') {
                // ★ 2026-07-22 修正：改用大家平常在用的後台共用密碼驗證，不要求
                // 額外記一組kioskKey——kioskKey是給kiosk電腦另一種用途的，
                // 一般後台使用者不會知道。做法是把密碼轉發給GAS那邊既有的
                // 密碼驗證邏輯確認一次，驗證通過才真的執行開關，不在Worker
                // 這裡自己另外存一份密碼、維護兩套密碼邏輯。
                const bodyUrl = new URL(request.url);
                const passwordForToggle = bodyUrl.searchParams.get('password') || '';
                if (!env.GAS_WEB_APP_URL) {
                    return withSecurityHeaders(jsonResponse({ status: 'error', message: '伺服器尚未設定完成' }, 500));
                }
                try {
                    const verifyResp = await fetch(env.GAS_WEB_APP_URL + '?action=getFeatureToggles&password=' + encodeURIComponent(passwordForToggle));
                    const verifyData = await verifyResp.json();
                    if (verifyData.status !== 'success') {
                        return withSecurityHeaders(jsonResponse({ status: 'error', message: '密碼錯誤' }, 403));
                    }
                } catch (verifyErr) {
                    return withSecurityHeaders(jsonResponse({ status: 'error', message: '驗證密碼時發生錯誤：' + verifyErr.message }, 500));
                }
                const enabledParam = bodyUrl.searchParams.get('enabled') === 'true';
                if (kv) await kv.put('device_restriction_enabled', String(enabledParam));
                return withSecurityHeaders(jsonResponse({ status: 'success', enabled: enabledParam }));
            }
            return withSecurityHeaders(jsonResponse({ status: 'error', message: '不支援的方法' }, 405));
        }

        // /api/ 底下走代理 + 限流邏輯；不做 UA 擋（密碼本身是驗證），
        // 但一樣套用安全 headers
        if (url.pathname.startsWith('/api/')) {
            const resp = await handleApi(request, env);
            return withSecurityHeaders(resp);
        }

        // ★ 2026-07-21 新增：後台改用 Windows 封閉式電腦（kiosk 環境）操作，
        // Windows 上的 Chrome 不管有沒有開 kiosk 模式，User-Agent 都不會是
        // 手機格式、也不會有 LINE 標記，照原本規則會被直接擋在外面。
        // 這裡開一條「例外通道」：只有 /admin 路徑、而且網址帶著正確金鑰的
        // 請求，才能跳過手機/LINE 檢查——不是把規則整個放寬，只有知道這組
        // 金鑰的這台特定電腦能用，其他人一樣被擋。金鑰放在 kiosk 電腦的
        // 啟動捷徑網址裡，不會出現在一般人看得到的地方。
        const kioskKey = url.searchParams.get('kioskKey') || '';
        const isKioskAdmin = url.pathname.startsWith('/admin') &&
            !!env.ADMIN_KIOSK_KEY && kioskKey === env.ADMIN_KIOSK_KEY;

        // ★ 2026-07-22 修正：裝置門禁現在是可設定的開關，不是寫死一定要
        // 檢查。預設關閉（沒設定過就是關閉），讀KV裡的值決定要不要真的
        // 執行下面手機/LINE的檢查。
        const kvForCheck = env.RATE_LIMIT_KV;
        const deviceRestrictionEnabled = kvForCheck ? (await kvForCheck.get('device_restriction_enabled')) === 'true' : false;

        if (!isKioskAdmin && deviceRestrictionEnabled) {
            // 其他所有路徑（含首頁、tool_*.html 靜態檔案）：先過裝置門禁
            if (!isMobile) {
                return withSecurityHeaders(mobileBlockedResponse());
            }

            // 所有路徑（含 /admin，除了上面的 kiosk 例外）都要求一定要從
            // LINE 內建瀏覽器打開，在 Cloudflare 這一層就擋掉，不是「送出
            // 程式碼後才靠 JavaScript 判斷再踢人」。
            if (!/\bLine\//i.test(ua)) {
                return withSecurityHeaders(mobileBlockedResponse());
            }
        }

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
