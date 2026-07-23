// functions/_middleware.js
export async function onRequest(context){
  const { request, next } = context;
  const url = new URL(request.url);
  const ua = request.headers.get('User-Agent') || '';
  const cf = request.cf || {};
  if(/\.(css|js|png|jpg|jpeg|gif|webp|svg|ico|woff2|mp4|json)$/i.test(url.pathname)){
    return next();
  }
  if(url.pathname.startsWith('/api/')){
    const res = await next();
    res.headers.set('Cache-Control','no-store');
    return res;
  }
  const isMobileCF = cf.device ? (cf.device.type === 'mobile') : false;
  const isMobileUA = /Mobi|Android|iPhone|iPad|iPod|Mobile|Windows Phone/i.test(ua);
  if(!isMobileCF && !isMobileUA){
    return new Response(`<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>僅限手機</title></head><body style="margin:0;background:#0a0a0a;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;font-family:sans-serif;text-align:center"><div><div style="font-size:48px">📱</div><h2>此頁面僅限手機瀏覽</h2><p style="opacity:.6;font-size:13px">請用 iPhone / Android 開啟</p></div></body></html>`, {status:403, headers:{'Content-Type':'text/html; charset=utf-8','Cache-Control':'no-store'}});
  }
  const res = await next();
  res.headers.set('X-Frame-Options','DENY');
  res.headers.set('X-Content-Type-Options','nosniff');
  res.headers.set('Referrer-Policy','no-referrer');
  res.headers.set('Cache-Control','no-store');
  return res;
}
