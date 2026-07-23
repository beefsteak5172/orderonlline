// ⚠️⚠️⚠️ 這是「診斷用」的臨時版本，不是正式版！⚠️⚠️⚠️
// 用途：原本卡在 mobile/LINE 檢查會直接 302 跳轉去 Google，看不到任何線索。
// 這個版本改成把偵測結果印在畫面上，方便找出「一開就跳去 Google」到底是
// 卡在 isMobile 判斷，還是卡在 Line/ 判斷，還是 UA 字串本身長得跟預期不同。
//
// 使用方式：
// 1. 用這個檔案「暫時取代」你現有的 _worker.js，重新部署一次
// 2. 手機在 LINE App 裡點一次 LIFF 連結，把看到的畫面截圖給我
// 3. 診斷完看完結果後，記得把原本正常版本的 _worker.js 換回去、重新部署，
//    不要讓這個診斷版本留在正式環境（它會把 User-Agent 資訊暴露給任何訪客看）

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
    return new Response(response.body, {
        status: response.status,
        statusText: response.statusText,
        headers: newHeaders
    });
}

// ★ 診斷用：原本這裡是 mobileBlockedResponse()，直接 302 去 Google，
// 現在改成回傳一個顯示偵測結果的 HTML 頁面，不會再跳轉。
function debugBlockedResponse(reason, ua, isMobile, hasLineMarker) {
    const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>診斷模式</title>
<style>body{font-family:-apple-system,sans-serif;padding:20px;line-height:1.8;background:#111;color:#eee;}
h1{color:#f66;font-size:20px;}
.box{background:#222;padding:16px;border-radius:8px;margin:12px 0;word-break:break-all;}
.label{color:#8cf;font-weight:bold;}
.yes{color:#6f6;} .no{color:#f66;}
</style></head><body>
<h1>🔍 診斷模式：這次請求被擋下來了</h1>
<div class="box"><span class="label">被擋的原因：</span><br>${reason}</div>
<div class="box"><span class="label">isMobile 判斷結果：</span> <span class="${isMobile ? 'yes' : 'no'}">${isMobile}</span></div>
<div class="box"><span class="label">是否偵測到 Line/ 標記：</span> <span class="${hasLineMarker ? 'yes' : 'no'}">${hasLineMarker}</span></div>
<div class="box"><span class="label">完整 User-Agent 字串：</span><br>${ua}</div>
<p style="color:#999;font-size:13px;">把這個畫面截圖給開發者即可，這是暫時的診斷頁面，不是正式頁面。</p>
</body></html>`;
    return new Response(html, {
        status: 200,
        headers: { 'Content-Type': 'text/html; charset=utf-8' }
    });
}

function getIp(request) {
    return request.headers.get('CF-Connecting-IP') || 'unknown';
}

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
        const hasLineMarker = /\bLine\//i.test(ua);

        if (url.pathname.startsWith('/api/')) {
            const resp = await handleApi(request, env);
            return withSecurityHeaders(resp);
        }

        const kioskKey = url.searchParams.get('kioskKey') || '';
        const isKioskAdmin = url.pathname.startsWith('/admin') &&
            !!env.ADMIN_KIOSK_KEY && kioskKey === env.ADMIN_KIOSK_KEY;

        if (!isKioskAdmin) {
            // ★ 診斷模式：不是手機格式 → 顯示診斷頁面（原本是直接跳轉 Google）
            if (!isMobile) {
                return withSecurityHeaders(debugBlockedResponse('isMobile 判斷失敗（User-Agent 不像手機格式）', ua, isMobile, hasLineMarker));
            }

            // ★ 診斷模式：沒有 Line/ 標記 → 顯示診斷頁面（原本是直接跳轉 Google）
            if (!hasLineMarker) {
                return withSecurityHeaders(debugBlockedResponse('沒有偵測到 Line/ 標記（判斷不是從 LINE 內建瀏覽器打開的）', ua, isMobile, hasLineMarker));
            }
        }

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
