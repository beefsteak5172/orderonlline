// ★ Cloudflare Pages版本的LINE守門邏輯，取代原本Netlify的
// netlify/edge-functions/line-gate.js。這份用的是Cloudflare Pages
// Functions的官方語法（onRequest + context.next()），跟Netlify的
// Edge Function語法不一樣，不能直接複製過來，這份是重寫過的。
//
// 只守首頁這條路（跟Netlify那版的path="/"設定邏輯一致），
// 其他頁面（例如訂單看板、菜單管理）不受這個限制——但這幾份本來就不
// 部署到這個網域，是店家自己在本機電腦開的，且各自有獨立的後台密碼
// 保護，不依賴LINE身份，設計上本來就不需要LINE守門這一層。
//
// 判斷依據：LINE的內建瀏覽器（LIFF容器）打開網頁時，User-Agent字串裡
// 固定會帶"Line/x.x.x"這個標記，這是LINE App本身加上去的。不是從LINE
// 進來的請求，直接redirect到Google，完全不會把index.html的內容送出去。
export async function onRequest(context) {
  const { request, next } = context;
  const url = new URL(request.url);

  // ★ 統一轉小寫再比對，避免/INDEX.HTML、/Index.html這種大小寫變化版本
  // 因為完全比對（===）比不到，導致LINE守門檢查被整個跳過
  const pathLower = url.pathname.toLowerCase();
  const isRootPath = pathLower === '/' || pathLower === '/index.html';

  if (isRootPath) {
    const userAgent = request.headers.get('user-agent') || '';
    const isFromLine = /Line/i.test(userAgent);

    if (!isFromLine) {
      return Response.redirect('https://www.google.com', 302);
    }
  }

  return next();
}
