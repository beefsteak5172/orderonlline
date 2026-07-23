// ⚠️ Cloudflare Pages Functions 邊緣層裝置門禁
// 放在 functions/_middleware.js，Cloudflare 會在請求「進到你的 HTML/資源之前」先跑這段。
// 判定不是手機，直接回 403，連 admin_hub_basic.html 的原始碼都不會被送到對方瀏覽器裡，
// 跟前端 JS 那層（一定要先把整份 HTML 下載下來才能執行）是完全不同層級的防禦。
//
// 這一關防的是「一般人隨手用電腦瀏覽器打開網址」，UA 一看就不是手機，直接被邊緣節點擋掉。
// 防不住「刻意偽造 UA」的人（curl -A、瀏覽器擴充套件、devtools 換 UA 都能繞過字串比對），
// 那一類對手要靠前端那四層（devtools 尺寸偵測／滑鼠軌跡／字型指紋／持續監控）繼續擋。
// 兩層疊加：邊緣層擋掉「沒動手腳的大多數」，前端層擋掉「稍微動了手腳、但沒深入研究的人」。

const MOBILE_UA_REGEX = /Mobi|Android|iPhone|iPad|iPod|Windows Phone|BlackBerry/i;

// ⚠️ 2026-07-20 新增：安全 headers，套用在所有回應（包含放行的正常頁面）上
// - X-Frame-Options: DENY → 防止後台被其他網站用 <iframe> 嵌入做釣魚
// - Cache-Control: no-store → 防止手機瀏覽器把已登入頁面／密碼快取下來，
//   離開網頁後被同一支手機的其他使用者（或遺失手機時）翻出快取內容看到
// - Content-Security-Policy → 限制頁面只能載入自己網域的資源，降低被注入外部腳本的風險
//   （這裡先給寬鬆一點的設定，因為你們頁面裡有 inline <script>，之後想收緊要另外處理 nonce）
function withSecurityHeaders(response) {
    const newHeaders = new Headers(response.headers);
    newHeaders.set('X-Frame-Options', 'DENY');
    newHeaders.set('Cache-Control', 'no-store, no-cache, must-revalidate');
    newHeaders.set(
        'Content-Security-Policy',
        "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data: https:; connect-src 'self'; frame-ancestors 'none';"
    );
    return new Response(response.body, {
        status: response.status,
        statusText: response.statusText,
        headers: newHeaders
    });
}

export async function onRequest(context) {
    const { request, next } = context;
    const url = new URL(request.url);

    // /api/ 底下是給前端 fetch 呼叫用的代理端點，不是给人直接瀏覽的頁面，
    // 這裡不做 UA 擋，交給 functions/api/gas.js 自己的邏輯處理
    // （密碼本身就是一層驗證，UA 擋在這裡意義不大，還可能誤擋合法的 fetch）
    if (url.pathname.startsWith('/api/')) {
        const resp = await next();
        return withSecurityHeaders(resp);
    }

    const ua = request.headers.get('User-Agent') || '';
    const isMobile = MOBILE_UA_REGEX.test(ua);

    if (!isMobile) {
        return withSecurityHeaders(new Response(
            `<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
            <body style="background:#0a0a0a;color:#fff;display:flex;justify-content:center;align-items:center;height:100vh;font-family:sans-serif;text-align:center;margin:0;">
                <div><h1 style="font-size:20px;">僅限手機瀏覽</h1><p style="color:#888;font-size:14px;">請使用手機開啟此連結</p></div>
            </body></html>`,
            {
                status: 403,
                headers: { 'Content-Type': 'text/html; charset=utf-8' }
            }
        ));
    }

    const resp = await next();
    return withSecurityHeaders(resp);
}
