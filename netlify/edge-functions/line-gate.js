// ★ LINE來源守門：在檔案內容送出去「之前」就攔截，不是內容送到對方手上才轉走。
// 這是Netlify Edge Function，跑在Netlify的邊緣節點，比你的index.html本身還早一步執行。
//
// 判斷依據：LINE的內建瀏覽器（LIFF容器）打開網頁時，User-Agent字串裡固定會帶
// "Line/x.x.x" 這個標記，這是LINE App本身加上去的，不是我們能改的東西。
// 不是從LINE進來的請求（User-Agent沒有這個標記），直接redirect到別的地方，
// 完全不會把index.html的內容送給對方。
//
// ★ 老實留一句話在這裡，之後回來看要記得：
// 這道關卡防的是「隨手打開網址的人」跟「連User-Agent都懶得改的簡單工具」，
// 防不住「特意把curl的User-Agent字串改成LINE那種格式」的人——這種偽造成本
// 低到任何懂一點技術的人都做得到。但店家評估過，即使偽裝成功，對方拿到的
// 東西跟真實客人打開App看到的完全一樣（沒有分級內容），這道關卡的意義是
// 阻擋「隨手/自動化」的存取行為，不是阻擋「有心人」，這點跟店家已經對齊過。
const LINE_UA_PATTERN = /Line/i;

export default async (request, context) => {
  const userAgent = request.headers.get('user-agent') || '';
  const isFromLine = LINE_UA_PATTERN.test(userAgent);

  if (!isFromLine) {
    return Response.redirect('https://www.google.com', 302);
  }

  return context.next();
};

export const config = {
  path: '/*'
};
