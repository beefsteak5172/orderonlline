/**
 * 咱的台雞店 線上訂餐系統 - Apps Script 後端
 * ------------------------------------------------------
 * ★ 本版已整合 LINE ID Token 身份驗證（防止偽造/冒用LINE UserID）
 *   新增：LINE_CHANNEL_ID 常數、verifyLineIdToken() 函式
 *   修改：handleSubmitOrder / handleGetMyOrders / handleGetBalance / handleSubmitTopupRequest
 *         開頭都會先驗證 idToken，通過後才用「驗證過的真實UserID」執行邏輯，
 *         不再相信前端自己傳來的 lineUserId 字串。
 * ★★ 2026-07-14 bugfix：修正 doGet 路由判斷式的 bug（送單請求原本被誤判成開網址）。
 * ★★★ 2026-07-15 防濫用強化：
 *   1. 取餐時段時效性檢查：pickupDate/pickupTime已過去或不到1小時內，直接拒收
 *      （擋舊分頁/舊快取留著的orderData在很久之後才被重送）
 *   2. 頻率限制：同一個（驗證過的）LINE UserID，1分鐘內最多送出3張訂單
 *      （用CacheService存時間戳陣列，擋機器人/連點狂洗試算表配額）
 *   3. 簡易來源token驗證：前端送單時要附上APP_SHARED_SECRET，核對不符直接拒絕
 *      （GAS doGet拿不到真正的Origin/Referer header，這層只能擋隨手亂打API的，
 *      擋不住認真想繞過的人——真正的防線還是idToken驗證跟後端重新核算金額）
 * ★★★★ 2026-07-18 安全審計修正（本次一次補齊5項）：
 *   1. verifyAdminPassword()：tryLock()失敗（搶不到鎖）時，原本會直接放行執行
 *      讀取/比對/寫回邏輯，等於鎖形同虛設，高併發下密碼鎖定次數一樣可以被繞過。
 *      改成搶不到鎖就直接回錯誤，不落地執行。
 *   2. handleSubmitOrderCore()第一階段庫存核算：同樣的tryLock()沒判斷成功與否的
 *      問題，改成搶不到鎖就直接擋單，不要在沒鎖保護的情況下核算庫存
 *      （最終複查那關本來就有做對，這裡補齊讓兩處邏輯一致）。
 *   3. 訂單編號一律由伺服器端generateOrderNumber()產生，不再信任客戶端傳來的
 *      orderData.orderNumber，避免有心人蓄意帶入跟別人相同的訂單編號，
 *      干擾退款/儲值扣款比對邏輯。
 *   4. sanitizeForSheet()原本漏蓋orderData.items跟paymentMethod兩個欄位，
 *      這兩個雖然正常情況下是前端自己組出來的，但API本身沒有嚴格驗證呼叫者
 *      一定是自家前端，補上這兩處的公式注入防護。
 *   5. handleSubmitTopupRequest／handleSubmitPaymentReport原本只驗證身份，
 *      沒有頻率限制，同一個真實帳號可以無限狂送，狂洗LINE verify API配額
 *      跟試算表寫入配額。抽出通用的checkRateLimit()，這兩支API也掛上限制。
 * ★★★★★ 2026-07-20 資料拆分：後台「營業日報表」相關資料（固定成本設定、
 *   成本類別設定、每日營運記錄、採購品項資料庫、採購單記錄、本地系統日報）
 *   全數搬到獨立的第二份試算表（REPORT_SPREADSHEET_ID），跟客人點餐相關的
 *   工作表（訂單資料、用戶資料、菜單…等）分開放，原因：
 *   1. 工作表數量太多，同一份試算表底部的分頁列擠成一長排，不好操作
 *   2. 兩種資料的存取頻率、開放對象本來就不同——點餐相關的表每次客人
 *      下單/查詢都在讀寫，營業報表只有店家自己月結/日結會看，混在一起
 *      沒有必要，分開後也方便未來把B表的存取權限收得更緊（不給前台API
 *      以外的人碰）
 *   做法：getOrCreateSheet()新增第三個參數可以指定要開哪一份試算表；
 *   所有原本操作這幾張表的函式，一律改成指定REPORT_SPREADSHEET_ID；
 *   setupAllSheets()只再初始化A表（點餐用）的工作表，B表（報表用）改由
 *   新增的setupReportSheets()負責初始化，兩支互不影響、各自獨立執行。
 * ★★★★★★ 2026-07-21：拿掉「缺匯款截圖」的管理員備註提示（missingScreenshot
 *   那段），因為已經不需要這個舊版前端相容流程了。
 * ------------------------------------------------------
 * 對應：咱的台雞店_線上訂餐.html（前端表單）
 *       Google 試算表（A：點餐主表）：https://docs.google.com/spreadsheets/d/1wn-9Yswm0JFtm-AEMlC7mC1P6KWoGPAg3FA8GMcYzx8/edit
 *       Google 試算表（B：營業日報表）：https://docs.google.com/spreadsheets/d/1J2ul-2p3dHHN38IaWH3yt7bl91nNBUOkBk2bYyuGG40/edit
 *
 * 這是「獨立 Apps Script 專案」（不是從試算表裡「擴充功能」打開的），
 * 所以程式碼用 SpreadsheetApp.openById() 指定寫入哪一份試算表，
 * 不依賴 getActiveSpreadsheet()。
 *
 * 使用方式：
 * 1. 在 script.google.com 開啟你這個獨立專案
 * 2. 把這個檔案內容整段貼進去，取代原本的內容
 * 3. 把下面 APP_SHARED_SECRET 改成你自己的隨機字串（跟前端 index.html 的
 *    APP_TOKEN 常數要一致，否則所有訂單都會被擋掉）
 * 4. 點選「部署」→「管理部署作業」→ 編輯現有部署 → 版本選「新版本」→ 部署
 *    （如果是第一次部署才選「新增部署作業」）
 * 5. 部署後得到的網址貼到前端 HTML 的 WEB_APP_URL
 * 6. Apps Script編輯器上方函式下拉選單，先執行一次 setupAllSheets()（建立
 *    A表的工作表），再執行一次 setupReportSheets()（建立B表的工作表）
 *
 * ------------------------------------------------------
 * 注意：本檔案支援兩種進入方式
 *  - doGet 無參數（開啟網頁）→ 渲染 index.html（Kiosk / 前端頁面）
 *  - doGet 有 action 參數 → 走 API 路由
 *  - doGet 有 orderData 參數（沒有 action）→ 走送訂單 API 路由
 * ------------------------------------------------------
 */

// ────────────────────────────────────────────
// 設定區：依你實際狀況調整
// ────────────────────────────────────────────
const SPREADSHEET_ID = '1C0J_4ot1IEJ_INpJeredfZYBRNDHn7JWy31GQu6ocvg';

// ★ 2026-07-20 新增：營業日報表系統獨立搬到這一份試算表，跟上面的點餐主表
// 分開存放，理由詳見檔案最上方的說明區塊。
const REPORT_SPREADSHEET_ID = '1J2ul-2p3dHHN38IaWH3yt7bl91nNBUOkBk2bYyuGG40';

// ★ 你的 LINE Channel ID（在 LINE Developers Console → 該Channel → Basic settings 裡找到）
// 用來驗證前端送來的 idToken 真的是核發給「你這個LIFF」的，不是別人隨便拿一個token套過來
const LINE_CHANNEL_ID = '2010768353';

// ★ 安全修正：APP_SHARED_SECRET改成存在PropertiesService，不再直接寫死在
// 程式碼裡。原因：程式碼原始碼很容易被複製、貼給別人看（今天就發生過，
// 你把整份Code.gs貼給別的AI分析，裡面寫死的密鑰就這樣一起外流了），
// PropertiesService的值不會顯示在程式碼文字裡，複製貼上原始碼不會連帶
// 洩漏這個值。第一次使用前，要先執行下面的setAppSharedSecret()一次性
// 設定好，前端index.html的APP_TOKEN常數也要改成同一組字串，兩邊要一致。
function getAppSharedSecret() {
  return PropertiesService.getScriptProperties().getProperty('APP_SHARED_SECRET') || '';
}

function setAppSharedSecret() {
  const YOUR_SECRET_HERE = '9d855a020376cc770519248ca8e90164ec7c3e7415393ed1';
  PropertiesService.getScriptProperties().setProperty('APP_SHARED_SECRET', YOUR_SECRET_HERE);
  Logger.log('✅ APP_SHARED_SECRET 已儲存');
}

const SHEET_NAME_ORDERS = '訂單資料';
const SHEET_NAME_USERS  = '用戶資料';
const SHEET_NAME_MENU   = '菜單';
const SHEET_NAME_ADDONS = '加購項目';
const SHEET_NAME_OPTION_SPECS = '選項規格';
const SHEET_NAME_PROMO_RULES = '優惠規則';
const SHEET_NAME_ANNOUNCEMENTS = '促銷公告';
const SHEET_NAME_TOPUP_TIERS = '儲值方案';
const SHEET_NAME_TOPUP_LEDGER = '儲值記錄';
const SHEET_NAME_PAYMENT_REPORTS = '付款回報';
const SHEET_NAME_DISCLAIMER = '免責聲明同意記錄';

// ────────────────────────────────────────────
// ★ 營業日報表系統（採購/成本/物料/損益分析）專用工作表——這一整組是
// 從「營業日報表_完整版.html」串接過來的，原本那份是連你自己電腦上的
// 本機伺服器（127.0.0.1:8765），現在改成全部存進Google試算表。
// 這些都是後台管理用的資料，不對客人開放，所有API都要求後台密碼。
// ★ 2026-07-20：這一組工作表全部改存在 REPORT_SPREADSHEET_ID（獨立的
// 第二份試算表），不再跟上面點餐用的工作表放在同一份，理由詳見檔案最上方。
// ────────────────────────────────────────────
const SHEET_NAME_FIXED_COSTS = '固定成本設定';
const SHEET_NAME_COST_CATEGORIES = '成本類別設定';
const SHEET_NAME_DAILY_RECORDS = '每日營運記錄';
const SHEET_NAME_PURCHASE_ITEMS = '採購品項資料庫';
const SHEET_NAME_PURCHASE_ORDERS = '採購單記錄';
const SHEET_NAME_LOCAL_DAILY_REPORT = '本地系統日報';

// ────────────────────────────────────────────
// LINE Messaging API 推播設定（訂單完成時通知客人）
// ────────────────────────────────────────────
function setLineChannelAccessToken() {
  const YOUR_TOKEN_HERE = '請貼上你從LINE Developers Console拿到的Channel access token';
  PropertiesService.getScriptProperties().setProperty('LINE_CHANNEL_ACCESS_TOKEN', YOUR_TOKEN_HERE);
  Logger.log('✅ LINE Channel Access Token 已儲存，訂單完成通知功能已可使用');
}

// ★ 新增（2026-07-19）：店家自己的LINE UserID，設定好之後，每次有新訂單
// 進來，系統會主動推播一則LINE訊息通知店家，不用一直盯著訂單看板螢幕。
// 取得方式：店家自己用LINE加自己的LINE官方帳號好友，正常走一次LIFF流程
// （例如打開點餐頁面），這個過程系統就會驗證過他的身份、知道他的UserID——
// 可以請店家在「訂單資料」表或「用戶資料」表找到自己那筆，複製LINE_UserID
// 那一串貼進這裡。可以設定多組（用逗號分隔），例如老闆+店長都想收到通知。
function setShopNotifyLineUserIds() {
  const YOUR_LINE_USER_IDS_HERE = '請貼上店家自己的LINE UserID，多組用逗號分隔';
  PropertiesService.getScriptProperties().setProperty('SHOP_NOTIFY_LINE_USER_IDS', YOUR_LINE_USER_IDS_HERE);
  Logger.log('✅ 店家通知用的LINE UserID已儲存，新訂單進來時會開始推播');
}

function getShopNotifyLineUserIds() {
  const raw = PropertiesService.getScriptProperties().getProperty('SHOP_NOTIFY_LINE_USER_IDS') || '';
  if (!raw || raw.indexOf('請貼上') !== -1) return [];
  return raw.split(',').map(s => s.trim()).filter(Boolean);
}

// ★ 新訂單推播給店家：每一組設定好的UserID都推一次，單一組失敗不影響其他組，
// 也不影響訂單本身的送出結果（推播只是錦上添花，不是訂單成立的必要條件）。
function notifyShopNewOrder(orderNumber, takeNumber, customerName, total, pickupTime) {
  try {
    if (!isFeatureEnabled('lineNotifyNewOrder')) return; // ★ 2026-07-22 新增：功能開關，關閉時完全不推播
    const ids = getShopNotifyLineUserIds();
    if (ids.length === 0) return;
    const msg = `🔔 新訂單！\n取餐號：${takeNumber || '?'}\n訂單編號：${orderNumber}\n客戶：${customerName || '（未填）'}\n金額：$${total}\n取餐時間：${pickupTime || ''}`;
    ids.forEach(id => {
      const result = sendLinePushMessage(id, msg);
      if (!result.ok) {
        Logger.log(`[新訂單通知] 推播給 ${id} 失敗：${result.reason}`);
      }
    });
  } catch (err) {
    Logger.log('[新訂單通知] 處理失敗（不影響訂單本身）：' + err.message);
  }
}

function getLineChannelAccessToken() {
  return PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN') || '';
}

function extractLineUserId(rawField) {
  const s = String(rawField || '').trim();
  if (!s) return '';
  const parts = s.split('｜');
  return parts.length > 1 ? parts[1].trim() : s;
}

function sendLinePushMessage(lineUserId, text) {
  const token = getLineChannelAccessToken();
  if (!token || token.indexOf('請貼上') !== -1) {
    return { ok: false, reason: '尚未設定 LINE_CHANNEL_ACCESS_TOKEN，請執行 setLineChannelAccessToken() 完成設定' };
  }
  if (!lineUserId) {
    return { ok: false, reason: '缺少LINE UserID（這筆訂單可能不是透過LIFF下單，沒有留下LINE身份）' };
  }
  try {
    const resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({
        to: lineUserId,
        messages: [{ type: 'text', text: text }]
      }),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) return { ok: true };
    return { ok: false, reason: `HTTP ${code}：${resp.getContentText()}` };
  } catch (err) {
    return { ok: false, reason: err.message };
  }
}

// ────────────────────────────────────────────
// ★ 驗證LIFF送來的ID Token是否為LINE官方真正核發、且沒被竄改。
// 通過後回傳data.sub，這才是真正、無法偽造的LINE UserID，
// 所有需要身份判斷的API都改用這個，不再相信前端自己傳來的lineUserId字串。
// ────────────────────────────────────────────
function verifyLineIdToken(idToken) {
  if (!idToken) return { ok: false, reason: '缺少身份憑證，請透過LINE重新進入頁面' };
  try {
    const resp = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: { id_token: idToken, client_id: LINE_CHANNEL_ID },
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() !== 200) {
      return { ok: false, reason: data.error_description || '身份憑證驗證失敗，請重新從LINE進入' };
    }
    return { ok: true, lineUserId: data.sub, lineDisplayName: data.name || '' };
  } catch (err) {
    return { ok: false, reason: '身份驗證過程發生錯誤：' + err.message };
  }
}

// ────────────────────────────────────────────
// ★ 新增：把「已驗證的LINE身份」記進「用戶資料」表，如果還沒記過的話。
// 這支函式刻意設計成任何地方呼叫都安全：
//   1. 沒有lineUserId就直接跳過，不報錯
//   2. 內部包try/catch，寫入失敗只記Logger.log，絕不往外拋錯——這永遠是
//      「順手記錄」的附加動作，不該讓它拖垮呼叫它的那支API的主要功能
//   3. 用完整掃描「用戶資料」表第一欄比對，同一個人不會被記錄兩次
// ────────────────────────────────────────────
function registerUserIfNew(lineUserId, lineDisplayName) {
  if (!lineUserId) return;
  try {
    const sheet = getOrCreateSheet(SHEET_NAME_USERS, ['LINE User ID', 'LINE Display Name', '註冊時間']);
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      if (data.some(row => row[0] === lineUserId)) return; // 已經記過了，不重複寫
    }
    sheet.appendRow([
      lineUserId,
      sanitizeForSheet(lineDisplayName || ''),
      Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss')
    ]);
  } catch (err) {
    Logger.log('[用戶資料] registerUserIfNew 寫入失敗：' + err.message);
  }
}

// ────────────────────────────────────────────
// ★ 簡易來源驗證：核對前端送來的 appToken 是否等於後端設定的共用密鑰。
// 只用在「會寫入資料/消耗配額」的公開端點（目前是送訂單），
// 讀取類的公開API（getMenu/getAnnouncements等）不受影響，維持原本的開放存取。
// ────────────────────────────────────────────
function checkAppToken(rawToken) {
  const token = (rawToken || '').trim();
  const secret = getAppSharedSecret();
  if (!secret) {
    // 店家還沒執行 setAppSharedSecret() 設定密鑰，先不擋，但這代表這層防護目前形同虛設
    return { ok: true };
  }
  if (token !== secret) {
    return { ok: false, reason: '請求來源無法驗證，請重新整理頁面後再試一次' };
  }
  return { ok: true };
}

// ────────────────────────────────────────────
// ★ 浮水印陷阱：跟前端的網域鎖是同一組概念的後端版本，用途不同——
// 前端網域鎖是「擋」，這裡是「記錄」，用來對付「有人把網域鎖那段程式碼
// 刪掉繞過去」的情況：只要他複製走的檔案裡，showVersionInfo()那次請求
// 沒有一併被改掉，這裡就會記到一筆非自家網域的來源，事後可以查。
// 跟前端網域鎖用同一份白名單，這裡也要留意 script.google.com／
// *.googleusercontent.com 是自己人（GAS原生渲染這份頁面時走的網域），
// 不能誤記成可疑來源。
// ────────────────────────────────────────────
const ALLOWED_CLIENT_HOSTS = ['cline4.anonymousbeefsteak.workers.dev', 'localhost', '127.0.0.1', 'script.google.com'];

function isAllowedClientHost(host) {
  if (!host) return true; // 沒帶這個參數（例如舊版前端還沒更新）不記錄，避免誤判
  if (ALLOWED_CLIENT_HOSTS.indexOf(host) !== -1) return true;
  if (/\.googleusercontent\.com$/.test(host)) return true;
  return false;
}

function logSuspiciousClientHost(e) {
  if (!e || !e.parameter) return;
  const clientHost = (e.parameter.clientHost || '').trim();
  if (isAllowedClientHost(clientHost)) return;

  const sheet = getOrCreateSheet('來源監控', ['時間', '來源網域', '備註']);
  sheet.appendRow([
    Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
    sanitizeForSheet(clientHost),
    '偵測到非自家網域打這支API，可能是程式碼被複製到其他網站上使用'
  ]);
}

// ────────────────────────────────────────────
// ★ 取餐時段時效性檢查：擋掉「舊分頁/舊快取留著的舊orderData」在很久之後
// 才被送出，或客人卡在結帳頁面很久才按送出，這時候原本算好的取餐時段可能
// 早就過去了。pickupTime格式固定是「HH:MM-HH:MM」，取開始時間來比較。
// ────────────────────────────────────────────
function parsePickupStartDateTime(pickupDate, pickupTime) {
  if (!pickupDate || !pickupTime) return null;
  const startPart = String(pickupTime).split('-')[0].trim();
  const iso = `${pickupDate}T${startPart}:00+08:00`;
  const d = new Date(iso);
  if (isNaN(d.getTime())) return null;
  return d;
}

// ★ 這個門檻一定要比前端 PICKUP_LEAD_HOURS_START（=60分鐘）小，留出緩衝空間。
// 原因：前端算「取餐時段起點」是在客人打開結帳表單那一刻算好的（現在+60分鐘），
// 但客人填姓名/電話/勾選儲值餘額/備註這些動作需要時間，等真正按下送出，
// 後端收到請求時，「起點」跟「現在」的差距已經比60分鐘小了——如果這裡也設60，
// 等於幾乎每一筆正常訂單都會被擋下來，不是防呆，是防到自己人。
// 45分鐘代表：只要客人在15分鐘內完成結帳表單就不會被擋，這對正常填單速度綽綽有餘，
// 同時還是能擋掉「舊分頁擺著超過15分鐘才想到要送出」這種真正過期的情況。
const PICKUP_MIN_LEAD_MINUTES_SERVER = 45;

function checkPickupWindowValid(pickupDate, pickupTime) {
  const pickupStart = parsePickupStartDateTime(pickupDate, pickupTime);
  if (!pickupStart) {
    return { ok: false, reason: '取餐時段格式錯誤，請重新整理頁面再送出一次' };
  }
  const minutesUntilPickup = (pickupStart.getTime() - Date.now()) / 60000;
  if (minutesUntilPickup < PICKUP_MIN_LEAD_MINUTES_SERVER) {
    return { ok: false, reason: '這個取餐時段已經過期或太接近現在時間，請重新整理頁面，取得最新的取餐時段後再送出' };
  }
  return { ok: true };
}

// ────────────────────────────────────────────
// ★ 通用頻率限制：用CacheService存最近操作時間戳陣列，過期自動清除。
// actionKey用來區分不同操作各自獨立計數（下單、儲值申請、付款回報...），
// 同一個人在「下單」被限流，不影響他在「付款回報」的額度，反之亦然。
// ────────────────────────────────────────────
function checkRateLimit(lineUserId, actionKey, maxCount, windowMs, tooFrequentMessage) {
  if (!lineUserId) return { ok: true };
  const cache = CacheService.getScriptCache();
  const key = `rate_${actionKey}_` + lineUserId;
  let timestamps = [];
  const raw = cache.get(key);
  if (raw) {
    try { timestamps = JSON.parse(raw); } catch (e) { timestamps = []; }
  }
  const now = Date.now();
  timestamps = timestamps.filter(t => now - t < windowMs);
  if (timestamps.length >= maxCount) {
    return { ok: false, reason: tooFrequentMessage || '操作太頻繁，請稍等一下再試' };
  }
  timestamps.push(now);
  // CacheService.put的過期秒數上限是21600秒(6小時)，這裡窗口都遠小於此，
  // 多留30秒緩衝即可，不需要額外處理上限問題
  cache.put(key, JSON.stringify(timestamps), Math.ceil(windowMs / 1000) + 30);
  return { ok: true };
}

// ────────────────────────────────────────────
// ★ 頻率限制：同一個（驗證過的）LINE UserID，1分鐘內最多送出3張訂單，
// 擋機器人或手滑連點狂洗試算表配額。
// 沒有驗證過身份（沒有idToken）的請求無法用這個機制限流——但前端現在
// 強制要求透過LIFF登入才能進頁面，正常流程一定會帶idToken，這裡先只管
// 有身份的情況即可。
// ────────────────────────────────────────────
const ORDER_RATE_LIMIT_MAX = 3;
const ORDER_RATE_LIMIT_WINDOW_MS = 60000;

function checkOrderRateLimit(lineUserId) {
  return checkRateLimit(
    lineUserId,
    'order',
    ORDER_RATE_LIMIT_MAX,
    ORDER_RATE_LIMIT_WINDOW_MS,
    `下單太頻繁，同一身份1分鐘內最多送出${ORDER_RATE_LIMIT_MAX}張訂單，請稍等一下再試`
  );
}

// ★ 儲值申請、付款回報都不像下單那樣本來就有「購物流程」的天然節奏，
// 濫用起來對LINE verify API跟試算表寫入配額的消耗一樣大，比照下單的
// 額度設定（1分鐘3次），各自獨立計數。
const TOPUP_RATE_LIMIT_MAX = 3;
const TOPUP_RATE_LIMIT_WINDOW_MS = 60000;

function checkTopupRateLimit(lineUserId) {
  return checkRateLimit(
    lineUserId,
    'topup',
    TOPUP_RATE_LIMIT_MAX,
    TOPUP_RATE_LIMIT_WINDOW_MS,
    `申請太頻繁，同一身份1分鐘內最多送出${TOPUP_RATE_LIMIT_MAX}次儲值申請，請稍等一下再試`
  );
}

const PAYMENT_REPORT_RATE_LIMIT_MAX = 3;
const PAYMENT_REPORT_RATE_LIMIT_WINDOW_MS = 60000;

function checkPaymentReportRateLimit(lineUserId) {
  return checkRateLimit(
    lineUserId,
    'paymentReport',
    PAYMENT_REPORT_RATE_LIMIT_MAX,
    PAYMENT_REPORT_RATE_LIMIT_WINDOW_MS,
    `回報太頻繁，同一身份1分鐘內最多送出${PAYMENT_REPORT_RATE_LIMIT_MAX}次付款回報，請稍等一下再試`
  );
}

// ★ 安全修正：後台密碼同樣改成存在PropertiesService，理由跟上面的
// APP_SHARED_SECRET一樣——不要讓密碼明文出現在程式碼裡。
function getAdminPasswordValue() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
}

function setAdminPassword() {
  const YOUR_PASSWORD_HERE = '1234';
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', YOUR_PASSWORD_HERE);
  Logger.log('✅ 後台密碼已儲存');
}

// ★ 安全考量調整：這個鎖定機制的失敗次數存在PropertiesService，是全域計數，
// 不分請求來源（GAS平台本身沒辦法可靠拿到真實IP去區分「是誰在打」），代表
// 任何人（不管是真的想暴力猜密碼，還是單純誤打）打錯5次，你自己也會被鎖住，
// 這是這個機制天生的攻擊面，沒有乾淨的修法。這裡把次數放寬、鎖定時間縮短，
// 降低「自己被鎖住」的痛苦跟機率，同時8次嘗試對真正想暴力猜密碼的人來說
// 還是完全不夠用，防護力沒有實質下降。真的被鎖住，還是可以執行
// resetLoginLockout()手動解鎖。
const LOGIN_MAX_ATTEMPTS = 8;
const LOGIN_LOCKOUT_MINUTES = 5;

function resetLoginLockout() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('login_fail_count');
  props.deleteProperty('login_last_fail_time');
  Logger.log('✅ 密碼鎖定已解除，可以重新登入了');
}

function verifyAdminPassword(password) {
  // ★ 安全修正：原本這裡是「讀取→判斷→寫回」沒有加鎖，平行送出多筆猜密碼
  // 請求時，每個請求都讀到同一個舊的失敗次數、各自算完+1再各自寫回，
  // 會導致失敗次數卡住、永遠到不了鎖定門檻，等於鎖定機制可以被平行請求
  // 繞過。加上LockService，確保同一時間只有一個請求能讀取+更新這個計數，
  // 跟儲值扣款那邊用的是同一套機制。
  //
  // ★★ 2026-07-18 修正：原本這裡拿不到鎖（tryLock逾時）時，gotLock會是
  // false，但底下的讀取/比對/寫回邏輯完全沒有判斷gotLock就直接執行，等於
  // 鎖形同虛設——高併發密碼嘗試時，多個請求還是可能同時讀到舊的失敗次數，
  // 上面提到要修的race condition在「搶不到鎖」這個分支完全沒被堵住。
  // 改成：搶不到鎖就直接回錯誤，不落地執行任何讀取/比對/寫回，寧可讓使用者
  // 重新整理再試一次，也不要讓鎖定機制在忙線時失去保護力。
  const lock = LockService.getScriptLock();
  let gotLock = false;
  try {
    gotLock = lock.tryLock(5000);
  } catch (lockErr) {
    gotLock = false;
  }

  if (!gotLock) {
    return { ok: false, message: '系統忙線中，請稍後再試一次' };
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const failCountRaw = props.getProperty('login_fail_count');
    const lastFailRaw = props.getProperty('login_last_fail_time');
    const failCount = failCountRaw ? parseInt(failCountRaw, 10) : 0;
    const lastFailTime = lastFailRaw ? parseInt(lastFailRaw, 10) : 0;
    const now = Date.now();

    if (failCount >= LOGIN_MAX_ATTEMPTS) {
      const minutesSinceLastFail = (now - lastFailTime) / 60000;
      if (minutesSinceLastFail < LOGIN_LOCKOUT_MINUTES) {
        const remain = Math.ceil(LOGIN_LOCKOUT_MINUTES - minutesSinceLastFail);
        return { ok: false, message: `密碼錯誤次數過多，帳號已暫時鎖定，請 ${remain} 分鐘後再試` };
      }
      props.deleteProperty('login_fail_count');
      props.deleteProperty('login_last_fail_time');
    }

    const adminPasswordValue = getAdminPasswordValue();

    // ★ 安全防呆：如果店家還沒執行過 setAdminPassword() 設定密碼，
    // adminPasswordValue會是空字串。這裡故意「一律拒絕」而不是放行，
    // 避免萬一有人送出空字串當密碼，跟「還沒設定」的空字串比對相符，
    // 變成不用密碼就能登入後台。
    if (!adminPasswordValue) {
      return { ok: false, message: '後台密碼尚未設定，請先在 Apps Script 執行一次 setAdminPassword() 完成設定' };
    }

    if (password === adminPasswordValue) {
      props.deleteProperty('login_fail_count');
      props.deleteProperty('login_last_fail_time');
      return { ok: true };
    }

    const newCount = failCount + 1;
    props.setProperty('login_fail_count', String(newCount));
    props.setProperty('login_last_fail_time', String(now));
    const remaining = LOGIN_MAX_ATTEMPTS - newCount;
    if (remaining > 0) {
      return { ok: false, message: `密碼錯誤，還剩 ${remaining} 次機會，超過將暫時鎖定 ${LOGIN_LOCKOUT_MINUTES} 分鐘` };
    }
    return { ok: false, message: `密碼錯誤次數過多，帳號已暫時鎖定 ${LOGIN_LOCKOUT_MINUTES} 分鐘` };
  } finally {
    lock.releaseLock();
  }
}

// ────────────────────────────────────────────
// ★ 新增（2026-07-21）：後台專用的 LINE 身份白名單。跟客人點餐頁面的
// idToken 驗證是同一套機制（verifyLineIdToken），差別在於這裡驗證完之後
// 還要多比對一次「這個人是不是在允許清單裡」，不是任何 LINE 帳號都能通過。
// 這是後台的第一道關卡，通過之後才會看到密碼登入畫面——兩道關卡疊加，
// 不是取代密碼，密碼還是要繼續設定。
// ────────────────────────────────────────────
function getAdminAllowedLineUserIds() {
  const raw = PropertiesService.getScriptProperties().getProperty('ADMIN_ALLOWED_LINE_USER_IDS') || '';
  if (!raw || raw.indexOf('請貼上') !== -1) return [];
  return raw.split(',').map(s => s.trim()).filter(Boolean);
}

// ★ 2026-07-22 新增：員工帳號分權限。格式擴充成「LINE_UserID:角色」，例如
// 「U123abc:owner,U456def:staff」，角色沒寫的話預設當owner（維持舊資料
// 相容，之前設定過的白名單不會因為升級這個功能就突然變成權限受限）。
// owner可以看到全部功能；staff看不到「菜單編輯」「系統設定」這兩個
// 比較敏感、容易改壞事情的分頁，其他（訂單處理、儲值確認等日常操作）
// 都還是能用。這是「介面層的整理」，讓員工介面單純一點、不容易誤觸，
// 不是後端強制擋權限的資安等級防護——後端的密碼驗證還是同一組共用密碼，
// 這點要跟店家說清楚，不要誤以為員工帳號完全進不去那些功能。
function getAdminRoleForLineUserId(lineUserId) {
  const raw = PropertiesService.getScriptProperties().getProperty('ADMIN_ALLOWED_LINE_USER_IDS') || '';
  if (!raw || raw.indexOf('請貼上') !== -1) return 'owner';
  const entries = raw.split(',').map(s => s.trim()).filter(Boolean);
  for (const entry of entries) {
    const parts = entry.split(':');
    if (parts[0] === lineUserId) {
      return parts[1] === 'staff' ? 'staff' : 'owner';
    }
  }
  return 'owner';
}

// ★ 執行這支一次性設定允許進後台的 LINE UserID，多組用逗號分隔
// （例如老闆+店長各自的LINE UserID都想放行）。可以加角色標記，例如
// 'Uxxx:owner,Uyyy:staff'，沒加標記的話預設是owner。
function setAdminAllowedLineUserIds() {
  const YOUR_LINE_USER_IDS_HERE = 'U677cec06ea014f2b8f47e72e4b7d0709';
  PropertiesService.getScriptProperties().setProperty('ADMIN_ALLOWED_LINE_USER_IDS', YOUR_LINE_USER_IDS_HERE);
  Logger.log('✅ 後台允許的 LINE UserID 清單已儲存');
}

function handleVerifyAdminLineId(e) {
  try {
    const idToken = e.parameter.idToken || '';
    const verify = verifyLineIdToken(idToken);
    if (!verify.ok) {
      return jsonResponse({ status: 'success', allowed: false, reason: verify.reason });
    }
    const allowedIds = getAdminAllowedLineUserIds();
    if (allowedIds.length === 0) {
      // ★ 還沒執行過 setAdminAllowedLineUserIds() 設定白名單，先不擋，
      // 避免忘記設定就把自己也鎖在後台外面。這代表這層防護目前形同虛設，
      // 要真的生效一定要記得執行上面那支設定函式。
      return jsonResponse({ status: 'success', allowed: true, note: '尚未設定白名單，暫不限制', role: 'owner' });
    }
    // ★ 2026-07-22 修正：原本用exact indexOf比對，格式擴充成「id:role」後，
    // 要先把角色標記拆掉，只拿UserID部分來比對，不然帶了角色標記的白名單
    // 反而會比對不到、把自己人擋在外面
    const idsOnly = allowedIds.map(entry => entry.split(':')[0]);
    const allowed = idsOnly.indexOf(verify.lineUserId) !== -1;
    const role = allowed ? getAdminRoleForLineUserId(verify.lineUserId) : null;
    return jsonResponse({ status: 'success', allowed: allowed, role: role });
  } catch (err) {
    return errResponse('驗證後台身份失敗：', err);
  }
}


// 無 action 且無 orderData → 渲染網頁介面（Kiosk / 前端頁面）
// 有 action 參數 → 走 API 路由
// 有 orderData 參數（送訂單，本來就不帶 action）→ 走送訂單路由
// ────────────────────────────────────────────
function doGet(e) {
  // ★ 修正：原本這裡只看「有沒有 action」，導致送訂單的請求（本來就不帶
  // action 參數）被誤判成「客人直接開網址」，回傳整包HTML而不是JSON，
  // 前端 response.json() 解析失敗才會顯示「網路連線問題」。
  // 現在改成同時看有沒有 orderData，送單請求才能正確落到下面的
  // handleSubmitOrder(e)，而不是被攔在這裡渲染頁面。
  if (!e || !e.parameter || (!e.parameter.action && !e.parameter.orderData)) {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('咱的台雞店 - 線上訂餐')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const action = e.parameter.action;

  // ★ 安全審計修正：原本這裡沒有先判斷「有沒有帶action」，而是一路比對到底、
  // 比對不到任何已知action就落到 handleSubmitOrder(e)。這代表只要帶著
  // orderData 參數，不管action打什麼字（甚至亂打一個不存在的字串），都會被
  // 當成送訂單處理，跑一趟完整的驗證鏈，還會回傳「缺少orderData參數」這種
  // 誤導性錯誤訊息。改成明確判斷：沒帶action才是送單，有帶但對不到任何
  // 已知action，直接回傳「不支援的操作」，不落底。不影響任何正常流程
  // （送單本來就不帶action，22個合法action全部照舊）。
  if (!action) {
    return handleSubmitOrder(e);
  }

  if (action === 'getMenu') return handleGetMenu(e);
  if (action === 'getStoreStatus') return handleGetStoreStatus(e);
  if (action === 'setStoreStatus') return handleSetStoreStatus(e);
  if (action === 'getFeatureToggles') return handleGetFeatureToggles(e);
  if (action === 'getCustomerNotes') return handleGetCustomerNotes(e);
  if (action === 'getClosureSchedule') return handleGetClosureSchedule(e);
  if (action === 'getRefundRecords') return handleGetRefundRecords(e);
  if (action === 'getSupplierPriceCompare') return handleGetSupplierPriceCompare(e);
  if (action === 'getSystemHealth') return handleGetSystemHealth(e);
  if (action === 'getPrepList') return handleGetPrepList(e);
  if (action === 'getSalesHeatmap') return handleGetSalesHeatmap(e);
  if (action === 'getRepeatCustomerRate') return handleGetRepeatCustomerRate(e);
  if (action === 'getMyMemberTier') return handleGetMyMemberTier(e);
  if (action === 'getCancelReasonStats') return handleGetCancelReasonStats(e);
  if (action === 'getMyUsualPickupTime') return handleGetMyUsualPickupTime(e);
  if (action === 'getShiftNotes') return handleGetShiftNotes(e);
  if (action === 'getClockRecords') return handleGetClockRecords(e);
  if (action === 'getMenuPresets') return handleGetMenuPresets(e);
  if (action === 'exportOrdersCsv') return handleExportOrdersCsv(e);
  if (action === 'verifyAdminLineId') return handleVerifyAdminLineId(e);
  if (action === 'updateMenu') return handleUpdateMenu(e);
  if (action === 'updateAddons') return handleUpdateAddons(e);
  if (action === 'updateOptionSpecs') return handleUpdateOptionSpecs(e);
  if (action === 'getOrders') return handleGetOrders(e);
  if (action === 'getSalesSummary') return handleGetSalesSummary(e);
  if (action === 'updateOrderStatus') return handleUpdateOrderStatus(e);
  if (action === 'getCustomerInfo') return handleGetCustomerInfo(e);
  if (action === 'getMyOrders') return handleGetMyOrders(e);
  if (action === 'cancelMyOrder') return handleCancelMyOrder(e);
  if (action === 'getPromoRules') return handleGetPromoRules();
  if (action === 'updatePromoRules') return handleUpdatePromoRules(e);
  if (action === 'getAnnouncements') return handleGetAnnouncements();
  if (action === 'updateAnnouncements') return handleUpdateAnnouncements(e);
  if (action === 'getTopupTiers') return handleGetTopupTiers();
  if (action === 'updateTopupTiers') return handleUpdateTopupTiers(e);
  if (action === 'getBalance') return handleGetBalance(e);
  if (action === 'checkDisclaimerAgreed') return handleCheckDisclaimerAgreed(e);
  if (action === 'agreeDisclaimer') return handleAgreeDisclaimer(e);
  if (action === 'submitTopupRequest') return handleSubmitTopupRequest(e);
  if (action === 'getTopupRequests') return handleGetTopupRequests(e);
  if (action === 'confirmTopupRequest') return handleConfirmTopupRequest(e);
  if (action === 'submitPaymentReport') return handleSubmitPaymentReport(e);
  if (action === 'getPaymentReports') return handleGetPaymentReports(e);
  if (action === 'confirmPaymentReport') return handleConfirmPaymentReport(e);
  if (action === 'getFixedCosts') return handleGetFixedCosts(e);
  if (action === 'updateFixedCosts') return handleUpdateFixedCosts(e);
  if (action === 'getCostReport') return handleGetCostReport(e);
  if (action === 'getRecentPurchaseItemNames') return handleGetRecentPurchaseItemNames(e);
  if (action === 'getCostCategories') return handleGetCostCategories(e);
  if (action === 'getDailyRecord') return handleGetDailyRecord(e);
  if (action === 'getRevenueByPeriod') return handleGetRevenueByPeriod(e);
  if (action === 'getPurchaseItems') return handleGetPurchaseItems(e);
  if (action === 'getMonthlyCategoryTotals') return handleGetMonthlyCategoryTotals(e);

  // ★ 帶了action參數，但對不到任何已知路由（打錯字/亂打），直接回傳錯誤，
  // 不再落到送單流程
  return jsonResponse({ status: 'error', message: '不支援的操作' });
}

// ────────────────────────────────────────────
// 訂單送出（★ 已整合 idToken 身份驗證 ＋ 時效性檢查 ＋ 頻率限制 ＋ 來源token）
// ────────────────────────────────────────────
// ────────────────────────────────────────────
// ★ 匯款截圖儲存：base64圖片資料解碼後存進Google Drive一個專屬資料夾，
// 回傳檔案的檢視連結，寫進試算表對應那筆訂單的欄位。
// ────────────────────────────────────────────
// ────────────────────────────────────────────
// ★ 自動備份：每天把整份試算表複製一份存進Drive，保留最近30天，
// 避免無限累積佔空間。這個函式本身不會自動執行，要靠下面的
// setupBackupTrigger()設定每日觸發器才會真的按時跑。
// ★ 2026-07-20：拆表之後，A表（點餐主表）跟B表（營業日報表）各自都要
// 備份，這支函式改成同時備份兩份試算表，避免只備份到A表、B表反而沒人管。
// ────────────────────────────────────────────
function backupSpreadsheet() {
  const folder = getOrCreateDriveFolder('咱的台雞店_試算表備份');
  const timestamp = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd_HHmm');

  try {
    DriveApp.getFileById(SPREADSHEET_ID).makeCopy(`備份_點餐主表_${timestamp}`, folder);
    Logger.log('[備份] 成功建立點餐主表備份：備份_點餐主表_' + timestamp);
  } catch (err) {
    Logger.log('[備份] 點餐主表備份失敗：' + err.message);
  }

  try {
    DriveApp.getFileById(REPORT_SPREADSHEET_ID).makeCopy(`備份_營業日報表_${timestamp}`, folder);
    Logger.log('[備份] 成功建立營業日報表備份：備份_營業日報表_' + timestamp);
  } catch (err) {
    Logger.log('[備份] 營業日報表備份失敗：' + err.message);
  }

  // 只留最近30天的備份，舊的自動清掉（兩份表的備份都放在同一個資料夾，一起清）
  try {
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 30);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getDateCreated() < cutoff) {
        f.setTrashed(true);
      }
    }
  } catch (err) {
    Logger.log('[備份] 清理舊備份失敗：' + err.message);
  }
}

function setupBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'backupSpreadsheet') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('backupSpreadsheet')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
  Logger.log('[備份] 已設定每日凌晨3點自動備份，設定成功');
}

function getOrCreateDriveFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// ★ 通用圖片儲存函式：base64圖片資料解碼後存進指定的Drive資料夾，回傳
// 「可以直接當<img src>來源顯示」的網址（uc?export=view格式，不是
// file.getUrl()那種網頁檢視連結）。匯款截圖、菜單照片都共用這支，
// 不用各寫一份重複邏輯。
function saveImageToDrive(base64DataUrl, folderName, filenamePrefix) {
  if (!base64DataUrl) return '';
  try {
    const matches = base64DataUrl.match(/^data:(image\/[a-zA-Z]+);base64,(.+)$/);
    if (!matches) return '';
    const mimeType = matches[1];
    const base64Data = matches[2];
    const bytes = Utilities.base64Decode(base64Data);
    const ext = mimeType.split('/')[1] || 'jpg';
    const blob = Utilities.newBlob(bytes, mimeType, `${filenamePrefix}_${Date.now()}.${ext}`);
    const folder = getOrCreateDriveFolder(folderName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return `https://drive.google.com/uc?export=view&id=${file.getId()}`;
  } catch (err) {
    Logger.log(`[圖片儲存-${folderName}] 儲存失敗：` + err.message);
    return '';
  }
}

// 相容舊呼叫名稱，內部改用通用版本
function saveScreenshotToDrive(base64DataUrl, orderNumber) {
  return saveImageToDrive(base64DataUrl, '咱的台雞店_匯款截圖', `${orderNumber}_匯款截圖`);
}

// ────────────────────────────────────────────
// ★ 菜單照片上傳：菜單管理後台用，客人端看不到這支API。
// 走doPost，因為圖片資料塞不進GET網址長度限制。
// ────────────────────────────────────────────
function handleUploadMenuImage(body) {
  try {
    const password = body.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) {
      return jsonResponse({ status: 'error', message: authCheck.message });
    }

    const imageBase64 = body.imageBase64 || '';
    if (!imageBase64) {
      return jsonResponse({ status: 'error', message: '缺少圖片資料' });
    }

    const url = saveImageToDrive(imageBase64, '咱的台雞店_菜單照片', '菜單照片');
    if (!url) {
      return jsonResponse({ status: 'error', message: '圖片上傳失敗，請重新選擇圖片再試一次（檔案可能太大或格式不支援）' });
    }

    return jsonResponse({ status: 'success', url: url });
  } catch (err) {
    return errResponse('上傳失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 訂單送出核心邏輯：doGet（舊版GET相容，無截圖功能）跟doPost（新版，
// 支援截圖上傳）兩個進入點都會呼叫這支共用函式，業務邏輯只寫一份，
// 避免兩邊各寫一次、之後改東西漏改其中一邊。
//
// 參數說明：
//   orderData        - 訂單內容物件（跟以前一樣的格式）
//   rawIdToken       - LINE idToken字串
//   appToken         - 來源驗證token字串
//   screenshotBase64 - 匯款截圖的base64圖片字串（data:image/xxx;base64,....），
//                      沒有截圖就傳空字串。只有doPost這條路才可能有值，
//                      doGet這條舊路由永遠傳空字串（GET網址帶不動圖片資料）。
// ────────────────────────────────────────────
// ★ 防公式注入：Google試算表規則是儲存格內容以=、+、-、@開頭會被當成公式執行，
// 不是純文字。客人填的姓名、備註這些欄位直接寫進試算表前，先過這一關——
// 開頭如果是這幾個公式觸發符號，前面加一個單引號，讓Sheets把它當純文字處理，
// 不會被誤判成公式執行（例如有人打「=HYPERLINK(...)」想在你看訂單時塞一個
// 連結誘騙你點擊，經過這個處理後會變成純文字顯示，不會被執行）。
function sanitizeForSheet(value) {
  const str = String(value == null ? '' : value);
  if (/^[=+\-@]/.test(str)) {
    return "'" + str;
  }
  return str;
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：功能開關系統。這是走向「以後能變成公版、多店共用」
// 的地基——現階段只服務一家店，開關存在單一份 PropertiesService 裡；
// 以後真的要做多店版本，這裡的判斷邏輯改成「讀這家店專屬的開關設定」
// 就好，呼叫這些函式的其他程式碼完全不用改。
// 新功能預設全部開啟，除非店家自己去後台關掉。
// ────────────────────────────────────────────
const DEFAULT_FEATURE_TOGGLES = {
  repeatOrder: true,           // 客人「再訂一次」按鈕
  lineNotifyNewOrder: false,   // LINE推播新訂單提醒（預設關閉，因為需要額外設定LINE Messaging API金鑰，沒設定金鑰硬開會一直失敗）
  customerNote: true,          // 熟客標記/備註
  scheduledClosure: true,      // 排程公休（週期性/未來日期）
  autoDelistOnStockout: true,  // 進貨數量歸零，對應菜單自動下架
  // ★ 2026-07-22 第二批新增
  itemProfitMargin: true,      // 每項品項的實際毛利率
  supplierPriceCompare: true,  // 供應商價格比較
  refundRecord: true,          // 正式退款/補償紀錄
  exportOrdersCsv: true,       // 訂單匯出成CSV/Excel
  systemHealthDashboard: true, // 系統健康度儀表板
  // ★ 2026-07-22 第三批新增
  customerRating: true,        // 客人評分/意見回饋
  busyTimeNotice: true,        // 忙碌時段提示（給客人看）
  pickupNotifyCustomer: false, // 取餐通知主動推播給客人（預設關閉，跟lineNotifyNewOrder一樣需要LINE_CHANNEL_ACCESS_TOKEN，沒設定金鑰開了也不會出錯）
  soldOutMessage: true,        // 限量品項售完提示「明天再來」
  closureCountdown: true,      // 公休倒數提示（客人提前看到）
  // ★ 2026-07-22 第四批新增
  pickupTimeSort: true,        // 出餐順序建議（依取餐時間排序）
  prepList: true,               // 備料清單自動產生
  kitchenDisplayMode: true,    // 廚房專用大螢幕模式
  orderEditLock: true,         // 多人同時處理訂單的狀態鎖定
  prepTimeStats: true,         // 出餐時間統計
  // ★ 2026-07-22 第五批新增
  birthdayDiscount: true,      // 生日優惠
  referralReward: true,        // 推薦好友獎勵
  firstOrderDiscount: true,    // 首次消費折扣
  thresholdGift: true,         // 滿額贈品（其實促銷系統裡THRESHOLD_GIFT類型早就支援，這裡只是補開關）
  timeLimitedDeal: true,        // 限時搶購
  // ★ 2026-07-22 第六批新增
  salesHeatmap: true,          // 星期幾/時段銷售熱力圖
  repeatCustomerRate: true,    // 回購率追蹤
  cancelReasonStats: true,     // 客人取消訂單原因統計
  memberTier: true,            // 簡易會員等級/累積消費徽章
  rememberPickupTime: true,     // 客人常用取餐時段記憶
  // ★ 2026-07-22 第七批新增
  staffPermissions: true,      // 員工帳號分權限
  shiftHandoverNotes: true,    // 交接班記事本
  staffClockIn: true,          // 員工出勤打卡
  menuPresets: true            // 菜單預設集（涵蓋複製昨天菜單、季節性菜單切換）
};

function getFeatureToggles() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('FEATURE_TOGGLES');
  if (!raw) return { ...DEFAULT_FEATURE_TOGGLES };
  try {
    // 用預設值當底，讀到的設定蓋上去——這樣以後新增功能、加新的預設開關，
    // 舊的既有設定不會因為JSON裡沒有這個新key就整個壞掉
    return { ...DEFAULT_FEATURE_TOGGLES, ...JSON.parse(raw) };
  } catch (e) {
    return { ...DEFAULT_FEATURE_TOGGLES };
  }
}

function isFeatureEnabled(key) {
  const toggles = getFeatureToggles();
  return toggles[key] !== false; // 沒有這個key時，預設當作開啟，避免新功能漏掉判斷式而被誤關
}

function handleGetFeatureToggles(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  try {
    return jsonResponse({ status: 'success', toggles: getFeatureToggles() });
  } catch (err) {
    return errResponse('讀取功能開關失敗：', err);
  }
}

function handleSetFeatureToggles(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  try {
    const props = PropertiesService.getScriptProperties();
    const current = getFeatureToggles();
    const updated = { ...current, ...(body.toggles || {}) };
    props.setProperty('FEATURE_TOGGLES', JSON.stringify(updated));
    logOperatorAction('更新功能開關', '', body.operator);
    return jsonResponse({ status: 'success', toggles: updated });
  } catch (err) {
    return errResponse('更新功能開關失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：店休/暫停接單功能
// 用 PropertiesService 存店家目前是「營業中」還是「公休中」，不用另外
// 開一個試算表分頁，這種單一開關值用 PropertiesService 最輕量。
//
// ★ 2026-07-22 擴充：加入「排程公休」——除了手動立即切換的開關，現在
// 也支援「每週固定公休的星期幾」（例如每週一）跟「預先排定的日期區間」
// （例如過年連續休好幾天），不用每次都臨時想到才手動點。判斷順序是：
// 先看有沒有落在排定的日期區間或每週固定公休日裡，只要符合任何一項，
// 直接算公休，不用管手動開關當時是什麼狀態；都沒有符合，才看手動開關。
// ────────────────────────────────────────────
function getStoreOpenStatus() {
  const props = PropertiesService.getScriptProperties();
  const isOpen = props.getProperty('STORE_IS_OPEN');
  const manualIsOpen = isOpen === null ? true : isOpen === 'true';
  const message = props.getProperty('STORE_CLOSED_MESSAGE') || '今日公休，請見諒';

  if (isFeatureEnabled('scheduledClosure')) {
    const scheduleCheck = checkScheduledClosure();
    if (scheduleCheck.isClosed) {
      return { isOpen: false, message: scheduleCheck.message || message };
    }
  }

  return { isOpen: manualIsOpen, message };
}

// ★ 檢查「今天」是不是落在排程公休範圍內（每週固定公休日、或預先排定的
// 日期區間），回傳{isClosed, message}
function checkScheduledClosure() {
  const props = PropertiesService.getScriptProperties();
  const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');
  const todayWeekday = new Date().getDay(); // 0=週日, 1=週一...6=週六，用GMT+8的話建議改用formatDate拿weekday比較保險，但這裡先簡化

  // 每週固定公休日：存一個陣列，例如 [1] 代表每週一公休
  try {
    const weeklyRaw = props.getProperty('WEEKLY_CLOSED_DAYS');
    const weeklyClosedDays = weeklyRaw ? JSON.parse(weeklyRaw) : [];
    const weekdayNum = parseInt(Utilities.formatDate(new Date(), 'GMT+8', 'u'), 10) % 7; // 1=週一...7=週日 轉成 0=週日...6=週六
    if (weeklyClosedDays.indexOf(weekdayNum) !== -1) {
      return { isClosed: true, message: '今日固定公休，請見諒' };
    }
  } catch (e) { /* 解析失敗當作沒設定，不擋營業 */ }

  // 預先排定的日期區間：存一個陣列，每筆是 {from, to, message}
  try {
    const rangesRaw = props.getProperty('CLOSURE_DATE_RANGES');
    const ranges = rangesRaw ? JSON.parse(rangesRaw) : [];
    for (const range of ranges) {
      if (todayStr >= range.from && todayStr <= range.to) {
        return { isClosed: true, message: range.message || '公休期間，請見諒' };
      }
    }
  } catch (e) { /* 解析失敗當作沒設定，不擋營業 */ }

  return { isClosed: false };
}

function handleGetStoreStatus(e) {
  try {
    // ★ 2026-07-22 新增：順便帶出「再訂一次」功能的開關狀態，客人點餐頁面
    // 本來就會呼叫這支API查公休狀態，搭便車一起帶回來，不用多開一支API、
    // 多一次網路請求
    const extra = { repeatOrderEnabled: isFeatureEnabled('repeatOrder') };

    // ★ 2026-07-22 新增：忙碌時段提示——算「待處理訂單」（未匯款/已匯款/
    // 準備中，還沒完成也還沒取消）的數量，超過門檻就提示客人「目前較忙」。
    // 門檻先寫死10筆，這個數字對小吃店來說算合理的「開始塞車」的量。
    if (isFeatureEnabled('busyTimeNotice')) {
      try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
        let pendingCount = 0;
        if (sheet) {
          const lastRow = sheet.getLastRow();
          if (lastRow >= 2) {
            const statuses = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
            const pendingStatuses = ['未匯款', '已匯款', '準備中'];
            pendingCount = statuses.filter(row => pendingStatuses.indexOf(row[0]) !== -1).length;
          }
        }
        const BUSY_THRESHOLD = 10;
        extra.isBusy = pendingCount >= BUSY_THRESHOLD;
        extra.pendingOrderCount = pendingCount;
      } catch (busyErr) { extra.isBusy = false; }
    }

    // ★ 2026-07-22 新增：公休倒數提示——如果未來3天內有排定的公休（每週
    // 固定公休日或預先排定的區間），提前告訴客人，不要當天才通知
    if (isFeatureEnabled('closureCountdown')) {
      try {
        extra.upcomingClosure = findUpcomingClosure(3);
      } catch (closureErr) { extra.upcomingClosure = null; }
    }

    return jsonResponse({ status: 'success', ...getStoreOpenStatus(), ...extra });
  } catch (err) {
    return errResponse('取得營業狀態失敗：', err);
  }
}

// ★ 找出未來withinDays天內，最近一個會公休的日期（不含今天，今天公不公休
// 已經在isOpen裡反映了，這裡專門找「還沒到、但快到了」的公休）
function findUpcomingClosure(withinDays) {
  const props = PropertiesService.getScriptProperties();
  const weeklyRaw = props.getProperty('WEEKLY_CLOSED_DAYS');
  const weeklyClosedDays = weeklyRaw ? JSON.parse(weeklyRaw) : [];
  const rangesRaw = props.getProperty('CLOSURE_DATE_RANGES');
  const ranges = rangesRaw ? JSON.parse(rangesRaw) : [];

  for (let i = 1; i <= withinDays; i++) {
    const checkDate = new Date();
    checkDate.setDate(checkDate.getDate() + i);
    const checkDateStr = Utilities.formatDate(checkDate, 'GMT+8', 'yyyy-MM-dd');
    const weekdayNum = parseInt(Utilities.formatDate(checkDate, 'GMT+8', 'u'), 10) % 7;

    if (weeklyClosedDays.indexOf(weekdayNum) !== -1) {
      return { date: checkDateStr, message: '固定公休' };
    }
    for (const range of ranges) {
      if (checkDateStr >= range.from && checkDateStr <= range.to) {
        return { date: checkDateStr, message: range.message || '公休' };
      }
    }
  }
  return null;
}

function handleSetStoreStatus(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) {
    return jsonResponse({ status: 'error', message: authCheck.message });
  }
  try {
    const props = PropertiesService.getScriptProperties();
    const isOpen = e.parameter.isOpen === 'true';
    props.setProperty('STORE_IS_OPEN', String(isOpen));
    if (e.parameter.message !== undefined) {
      props.setProperty('STORE_CLOSED_MESSAGE', e.parameter.message);
    }
    logOperatorAction(isOpen ? '恢復營業' : '設定公休', '', e.parameter.operator);
    return jsonResponse({ status: 'success', ...getStoreOpenStatus() });
  } catch (err) {
    return errResponse('設定營業狀態失敗：', err);
  }
}

// ★ 排程公休：查詢目前設定
function handleGetClosureSchedule(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  try {
    const props = PropertiesService.getScriptProperties();
    const weeklyRaw = props.getProperty('WEEKLY_CLOSED_DAYS');
    const rangesRaw = props.getProperty('CLOSURE_DATE_RANGES');
    return jsonResponse({
      status: 'success',
      weeklyClosedDays: weeklyRaw ? JSON.parse(weeklyRaw) : [],
      dateRanges: rangesRaw ? JSON.parse(rangesRaw) : []
    });
  } catch (err) {
    return errResponse('讀取排程公休設定失敗：', err);
  }
}

// ★ 排程公休：更新設定
function handleSetClosureSchedule(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  try {
    const props = PropertiesService.getScriptProperties();
    const weeklyClosedDays = Array.isArray(body.weeklyClosedDays) ? body.weeklyClosedDays : [];
    const dateRanges = Array.isArray(body.dateRanges) ? body.dateRanges : [];
    props.setProperty('WEEKLY_CLOSED_DAYS', JSON.stringify(weeklyClosedDays));
    props.setProperty('CLOSURE_DATE_RANGES', JSON.stringify(dateRanges));
    logOperatorAction('更新排程公休設定', '', body.operator);
    return jsonResponse({ status: 'success' });
  } catch (err) {
    return errResponse('更新排程公休設定失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：生日優惠。客人要先設定過生日月份（獨立一張分頁記錄），
// 生日當月第一次下單才會給折扣，同一年不能重複領取。折扣金額用
// PropertiesService存一個可調整的數字，預設$30，店家可以自己調整
// （目前沒有另外做設定介面，需要調整的話直接改BIRTHDAY_DISCOUNT_AMOUNT
// 這個PropertiesService的值，或請開發者協助調整程式碼裡的預設值）。
// ────────────────────────────────────────────
const SHEET_NAME_CUSTOMER_BIRTHDAYS = '顧客生日';

function checkBirthdayDiscount(lineUserId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_CUSTOMER_BIRTHDAYS);
  if (!sheet) return { eligible: false };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { eligible: false };

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // LINE_UserID, 生日月份, 上次領取年度
  const currentMonth = parseInt(Utilities.formatDate(new Date(), 'GMT+8', 'M'), 10);
  const currentYear = parseInt(Utilities.formatDate(new Date(), 'GMT+8', 'yyyy'), 10);

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== lineUserId) continue;
    const birthdayMonth = Number(data[i][1]);
    if (birthdayMonth !== currentMonth) return { eligible: false };
    const lastClaimedYear = Number(data[i][2]) || 0;
    if (lastClaimedYear >= currentYear) return { eligible: false }; // 今年已經領過了

    const props = PropertiesService.getScriptProperties();
    const amount = Number(props.getProperty('BIRTHDAY_DISCOUNT_AMOUNT')) || 30;

    // ★ 標記這一年已經領過，避免同一個生日月份重複下單重複折扣
    sheet.getRange(i + 2, 3).setValue(currentYear);
    return { eligible: true, discount: amount, label: `🎂 生日優惠(-$${amount})` };
  }
  return { eligible: false };
}

function handleSetCustomerBirthday(body) {
  if (!isFeatureEnabled('birthdayDiscount')) return jsonResponse({ status: 'error', message: '此功能未啟用' });
  try {
    const idVerify = verifyLineIdToken(body.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;

    const month = parseInt(body.month, 10);
    if (!month || month < 1 || month > 12) return jsonResponse({ status: 'error', message: '生日月份格式不正確' });

    const sheet = getOrCreateSheet(SHEET_NAME_CUSTOMER_BIRTHDAYS, ['LINE_UserID', '生日月份', '上次領取年度'], SPREADSHEET_ID);
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === lineUserId) { targetRow = i + 2; break; }
      }
    }
    if (targetRow !== -1) {
      sheet.getRange(targetRow, 2).setValue(month);
    } else {
      sheet.appendRow([lineUserId, month, 0]);
    }
    return jsonResponse({ status: 'success' });
  } catch (err) {
    return errResponse('設定生日失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：首次消費折扣。判斷這個LINE UserID在訂單資料裡，
// 有沒有出現過「非取消狀態」的訂單，沒有的話代表這是第一筆，給折扣。
// ────────────────────────────────────────────
function checkFirstOrderDiscount(lineUserId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!sheet) return { eligible: true, discount: getFirstOrderDiscountAmount(), label: buildFirstOrderLabel() };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { eligible: true, discount: getFirstOrderDiscountAmount(), label: buildFirstOrderLabel() };

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const hasPriorOrder = data.some(row => row[0] !== '' && row[1] !== '已取消' && String(row[5] || '').indexOf(lineUserId) !== -1);
  if (hasPriorOrder) return { eligible: false };

  return { eligible: true, discount: getFirstOrderDiscountAmount(), label: buildFirstOrderLabel() };
}
function getFirstOrderDiscountAmount() {
  return Number(PropertiesService.getScriptProperties().getProperty('FIRST_ORDER_DISCOUNT_AMOUNT')) || 20;
}
function buildFirstOrderLabel() {
  return `🎉 首次消費優惠(-$${getFirstOrderDiscountAmount()})`;
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：推薦好友獎勵。客人分享自己的專屬連結（帶著自己的
// LINE UserID當推薦碼）給朋友，朋友透過這個連結進來、送出「第一筆」
// 訂單時，雙方都會拿到儲值金獎勵——直接寫進既有的儲值帳本（跟儲值/
// 扣款用同一張表），不用另外做一套獎勵系統。
// ────────────────────────────────────────────
function processReferralReward(referralCode, refereeLineUserId, refereeName, orderNumber) {
  if (!isFeatureEnabled('referralReward')) return;
  if (!referralCode || !refereeLineUserId) return;
  if (referralCode === refereeLineUserId) return; // 不能推薦自己

  try {
    // 確認這真的是「第一筆」訂單才給獎勵（避免重複下單重複領獎勵）
    const firstOrderCheck = checkFirstOrderDiscount(refereeLineUserId);
    // checkFirstOrderDiscount此時已經寫入這筆新訂單了，所以要用「這筆訂單
    // 是不是這個人有紀錄以來的第一筆」來判斷，用訂單資料表本身查會更準確：
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const priorOrderCount = data.filter(row =>
      row[0] !== '' && row[0] !== orderNumber && row[1] !== '已取消' && String(row[5] || '').indexOf(refereeLineUserId) !== -1
    ).length;
    if (priorOrderCount > 0) return; // 不是第一筆，不給推薦獎勵

    const props = PropertiesService.getScriptProperties();
    const rewardAmount = Number(props.getProperty('REFERRAL_REWARD_AMOUNT')) || 30;

    const ledgerSheet = getOrCreateSheet(SHEET_NAME_TOPUP_LEDGER, [
      '記錄ID', '時間', 'LINE_UserID', 'LINE_DisplayName', '類型', '金額', '狀態',
      '對應訂單編號', '顧客姓名', '聯絡電話', '管理員備註'
    ]);
    const now = new Date();
    // 推薦人拿獎勵
    ledgerSheet.appendRow([
      `referral-${Date.now()}-r`, Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
      referralCode, '', '推薦獎勵', rewardAmount, '已確認', orderNumber, '', '',
      `推薦${refereeName || '朋友'}首次消費，獲得獎勵`
    ]);
    // 被推薦人（新客人）也拿獎勵
    ledgerSheet.appendRow([
      `referral-${Date.now()}-e`, Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
      refereeLineUserId, refereeName || '', '推薦獎勵', rewardAmount, '已確認', orderNumber, '', '',
      `透過好友推薦加入，獲得獎勵`
    ]);
  } catch (err) {
    Logger.log('[推薦好友獎勵] 處理失敗（不影響訂單本身）：' + err.message);
  }
}

function handleSubmitOrderCore(params) {
  const orderData = params.orderData;
  const rawIdToken = params.rawIdToken || '';
  const appTokenValue = params.appToken || '';
  const screenshotBase64 = params.screenshotBase64 || '';

  try {
    if (!orderData) {
      return jsonResponse({ status: 'error', message: '缺少 orderData 參數' });
    }

    // ★ 2026-07-22 新增：公休期間，後端直接拒絕送單，不是只靠前端畫面
    // 擋著——就算有人繞過前端畫面直接打API，公休時一樣送不進來。
    const storeStatus = getStoreOpenStatus();
    if (!storeStatus.isOpen) {
      return jsonResponse({ status: 'error', message: storeStatus.message });
    }

    // ★★ 2026-07-18 安全修正：訂單編號一律由伺服器端產生，不採信客戶端傳來
    // 的orderData.orderNumber。原本是「orderData.orderNumber || generateOrderNumber()」，
    // 代表只要客戶端故意帶一個orderNumber欄位，系統就會照單全收——如果剛好
    // 帶了跟別人相同（或猜測）的訂單編號，會干擾後面「取消退款」、「儲值扣款
    // 記錄比對對應訂單編號」這些依賴orderNumber唯一性的邏輯。訂單編號本來
    // 就該是伺服器內部生成的識別碼，不該讓客戶端有置喙空間，這裡在最一開始
    // 就覆蓋掉，後面全部邏輯統一使用這個值。
    orderData.orderNumber = generateOrderNumber();

    // ★ 來源驗證：沒帶對的 appToken 直接拒絕
    const tokenCheck = checkAppToken(appTokenValue);
    if (!tokenCheck.ok) {
      return jsonResponse({ status: 'error', message: tokenCheck.reason });
    }

    // ★ 安全修正：idToken從「選填」改成「強制要求」。原本設計是「有帶才驗證，
    // 沒帶就跳過」，這造成一個漏洞——完全不帶idToken送單，verifiedLineUserId
    // 會是空字串，讓下面的頻率限制直接放行不擋，等於只要知道APP_TOKEN，
    // 誰都能不受限制狂送訂單。既然系統其他地方（加好友檢查、免責聲明）都已經
    // 要求客人一定要有合法LIFF身份才能走到結帳這步，這裡改成一致的強制要求，
    // 沒有idToken或驗證失敗，一律拒收，不會影響正常透過LIFF點餐的客人。
    if (!rawIdToken) {
      return jsonResponse({ status: 'error', message: '缺少身份憑證，請透過LINE重新整理頁面後再試一次' });
    }
    const verify = verifyLineIdToken(rawIdToken);
    if (!verify.ok) {
      return jsonResponse({ status: 'error', message: verify.reason });
    }
    const verifiedLineUserId = verify.lineUserId;

    // ★ 頻率限制：同一個驗證過的身份，1分鐘內最多送出3張訂單
    const rateCheck = checkOrderRateLimit(verifiedLineUserId);
    if (!rateCheck.ok) {
      return jsonResponse({ status: 'error', message: rateCheck.reason });
    }

    const missing = [];
    if (!orderData.customerName)  missing.push('顧客姓名');
    if (!orderData.customerPhone) missing.push('聯絡電話');
    if (!orderData.pickupDate)    missing.push('取餐日期');
    if (!orderData.pickupTime)    missing.push('取餐時間');
    if (!orderData.items)         missing.push('訂單內容');

    if (missing.length > 0) {
      return jsonResponse({ status: 'error', message: `缺少必填欄位：${missing.join('、')}` });
    }

    const phoneOk = /^09\d{8}$/.test(orderData.customerPhone);
    if (!phoneOk) {
      return jsonResponse({ status: 'error', message: '電話格式錯誤，須為09開頭的10碼手機號碼' });
    }

    // ★ 取餐時段時效性檢查：已過去或不到1小時內，直接拒收，避免舊分頁/舊快取
    // 留著的舊orderData在很久之後才被重送，造成客人跟店家對取餐時間認知不一致
    const pickupCheck = checkPickupWindowValid(orderData.pickupDate, orderData.pickupTime);
    if (!pickupCheck.ok) {
      return jsonResponse({ status: 'error', message: pickupCheck.reason });
    }

    // ★ 安全庫存強制核算：這裡才是真正防超賣的關卡，不是前端顯示用的那個。
    // 只有取餐日期是「今天」的訂單才需要檢查（庫存限量是以取餐日為單位），
    // 只檢查有設定安全庫存（>0）的品項，沒設定的品項不限量、略過不檢查。
    //
    // ★ 安全修正：加上LockService保護「讀取已售數量→判斷」這段，避免跟
    // 密碼鎖定那次一樣的競爭條件——兩個客人同時搶同一款限量商品的最後1份時，
    // 原本會同時讀到「還沒賣完」，兩邊都能通過檢查，變成超賣。
    // 這裡鎖住的是「檢查」本身，只在購物車真的有限量商品時才鎖，
    // 不影響其他訂單的處理速度。
    //
    // ★★ 2026-07-18 安全修正：原本tryLock()失敗（gotStockLock為false）時，
    // 底下讀取soldTodayMap、逐一比對庫存的邏輯完全沒有判斷gotStockLock就
    // 直接執行，等於鎖形同虛設——高併發搶最後一份限量商品時，「搶不到鎖」
    // 那個分支完全沒被保護到，一樣可能造成多筆請求同時判定「還有庫存」。
    // 改成搶不到鎖就直接擋單，請客人稍後再試，跟下面「最終複查」那段的
    // 寫法保持一致（那段本來就有正確判斷gotFinalLock）。
    let stockCheckPassed = true;
    try {
      const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');
      if (orderData.pickupDate === todayStr) {
        let cartForStock = [];
        try { cartForStock = JSON.parse(orderData.cartSummary || '[]'); } catch (parseErr) { cartForStock = []; }

        if (cartForStock.length > 0) {
          const ssStock = SpreadsheetApp.openById(SPREADSHEET_ID);
          const menuItemsForStock = handleGetMenuRaw(ssStock);
          const stockLimitMap = {};
          menuItemsForStock.forEach(it => {
            if (Number(it.safetyStock) > 0) stockLimitMap[it.name] = Number(it.safetyStock);
          });

          const cartHasLimitedItem = cartForStock.some(ci => stockLimitMap[ci.name]);

          if (cartHasLimitedItem) {
            const stockLock = LockService.getScriptLock();
            let gotStockLock = false;
            try {
              gotStockLock = stockLock.tryLock(8000);
            } catch (lockErr) {
              gotStockLock = false;
            }
            if (!gotStockLock) {
              return jsonResponse({ status: 'error', message: '目前搶購人數較多，系統忙線中，請稍後再試一次' });
            }
            try {
              const soldTodayMap = computeSoldTodayByItemName();
              for (const ci of cartForStock) {
                const limit = stockLimitMap[ci.name];
                if (!limit) continue;
                const alreadySold = soldTodayMap[ci.name] || 0;
                const wantQty = Number(ci.quantity) || 0;
                if (alreadySold + wantQty > limit) {
                  const remain = Math.max(0, limit - alreadySold);
                  return jsonResponse({ status: 'error', message: `「${ci.name}」今日剩餘 ${remain} 份，您選的數量超過剩餘庫存，請重新調整數量後再試一次` });
                }
              }
            } finally {
              stockLock.releaseLock();
            }
          }
        }
      }
    } catch (stockErr) {
      Logger.log('[庫存核算] 下單核算失敗，本次不擋單：' + stockErr.message);
      // ★ 核算過程本身出錯（例如試算表暫時忙線），選擇不擋單而不是誤擋真客人，
      // 跟頻率限制、加好友檢查那些防呆機制同一個保守原則
    }

    const priceCheck = recomputeOrderTotal(orderData.cartSummary);
    if (priceCheck.blocked) {
      return jsonResponse({ status: 'error', message: priceCheck.reason });
    }
    const serverTotal = priceCheck.total;
    const priceMismatch = priceCheck.ok && Math.abs(serverTotal - Number(orderData.total || 0)) > 0.5;

    let promoDiscount = 0;
    let promoLabelsText = '';
    try {
      let cartForPromo = [];
      try { cartForPromo = JSON.parse(orderData.cartSummary || '[]'); } catch (parseErr) { cartForPromo = []; }
      const ssPromo = SpreadsheetApp.openById(SPREADSHEET_ID);
      const menuPriceMapPromo = {};
      handleGetMenuRaw(ssPromo).forEach(it => { menuPriceMapPromo[it.name] = Number(it.price) || 0; });
      const cartItemsForPromo = cartForPromo
        .filter(c => menuPriceMapPromo.hasOwnProperty(c.name))
        .map(c => ({ name: c.name, quantity: Number(c.quantity) || 0, unitPrice: menuPriceMapPromo[c.name] }));
      const promoRules = getPromoRulesList();
      const promoResult = calcPromoDiscount(cartItemsForPromo, promoRules);
      promoDiscount = promoResult.discount;
      if (promoResult.labels.length > 0) {
        promoLabelsText = promoResult.labels.map(l => `${l.text}(-$${l.amount})`).join('、');
      }
    } catch (promoErr) {
      Logger.log('[促銷折扣] 核算失敗，本次訂單不套用促銷折扣：' + promoErr.message);
    }

    // ★ 2026-07-22 新增：生日優惠、首次消費折扣。這兩個都是「跟客人身份
    // 有關」的折扣，不是促銷規則引擎（那個只看購物車內容，不知道是誰在
    // 買），所以獨立算，跟促銷折扣「取其中優惠較大的一個」，不會疊加
    // ——避免多重折扣疊加到金額異常，這是保守但安全的設計。
    let specialDiscount = 0;
    let specialLabel = '';
    if (verifiedLineUserId) {
      try {
        const birthdayResult = isFeatureEnabled('birthdayDiscount') ? checkBirthdayDiscount(verifiedLineUserId) : null;
        if (birthdayResult && birthdayResult.eligible) {
          specialDiscount = birthdayResult.discount;
          specialLabel = birthdayResult.label;
        }
      } catch (bdErr) { Logger.log('[生日優惠] 核算失敗：' + bdErr.message); }

      if (specialDiscount === 0 && isFeatureEnabled('firstOrderDiscount')) {
        try {
          const firstOrderResult = checkFirstOrderDiscount(verifiedLineUserId);
          if (firstOrderResult && firstOrderResult.eligible) {
            specialDiscount = firstOrderResult.discount;
            specialLabel = firstOrderResult.label;
          }
        } catch (foErr) { Logger.log('[首次消費折扣] 核算失敗：' + foErr.message); }
      }
    }
    // 跟促銷折扣取較大的那個，不疊加
    if (specialDiscount > promoDiscount) {
      promoDiscount = specialDiscount;
      promoLabelsText = specialLabel;
    }

    const rawTotal = priceCheck.ok ? serverTotal : Number(orderData.total || 0);

    // ★ 使用儲值餘額付款：一定要有「驗證過」的LINE身份才允許扣款，
    // 沒有idToken或驗證失敗，verifiedLineUserId會是空字串，這裡自然就不會觸發餘額付款
    let paidByBalance = false;
    let balanceDeducted = 0;
    const wantUseBalance = orderData.useBalance === true || orderData.useBalance === 'true';
    if (wantUseBalance && verifiedLineUserId) {
      const lock = LockService.getScriptLock();
      let gotLock = false;
      try {
        gotLock = lock.tryLock(5000);
        if (gotLock) {
          const currentBalance = computeBalance(verifiedLineUserId);
          if (currentBalance >= rawTotal && rawTotal > 0) {
            paidByBalance = true;
            balanceDeducted = rawTotal;
          }
        }
      } finally {
        if (gotLock) lock.releaseLock();
      }
    }

    if (paidByBalance) {
      promoDiscount = 0;
      promoLabelsText = '';
    }
    const finalTotal = paidByBalance ? 0 : Math.max(0, rawTotal - promoDiscount);

    const orderNumber = orderData.orderNumber;
    const now = new Date();

    // ★ 截圖存進Drive，拿到連結才繼續往下寫入試算表
    let screenshotUrl = '';
    if (!paidByBalance && screenshotBase64) {
      screenshotUrl = saveScreenshotToDrive(screenshotBase64, orderNumber);
      if (!screenshotUrl) {
        return jsonResponse({ status: 'error', message: '匯款截圖上傳失敗，請重新選擇圖片再試一次（檔案可能太大或格式不支援）' });
      }
    }

    // ★★ 安全修正：庫存核算跟真正寫入訂單之間，隔著金額核算、促銷計算、
    // 儲值扣款判斷、截圖上傳（最慢的一步）這些動作，前面那次庫存檢查
    // 通過之後，到這裡才真正「算已售出」的這段時間，理論上還是有機會
    // 被別的並行請求插隊超賣。這裡在真正寫入之前，重新加鎖、複查一次
    // 「當下」的庫存夠不夠，不夠就在這裡擋下來，不要照樣寫入造成超賣。
    try {
      const todayStrFinal = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');
      if (orderData.pickupDate === todayStrFinal) {
        let cartForFinalCheck = [];
        try { cartForFinalCheck = JSON.parse(orderData.cartSummary || '[]'); } catch (parseErr) { cartForFinalCheck = []; }
        if (cartForFinalCheck.length > 0) {
          const ssFinal = SpreadsheetApp.openById(SPREADSHEET_ID);
          const menuItemsFinal = handleGetMenuRaw(ssFinal);
          const stockLimitMapFinal = {};
          menuItemsFinal.forEach(it => {
            if (Number(it.safetyStock) > 0) stockLimitMapFinal[it.name] = Number(it.safetyStock);
          });
          const cartHasLimitedItemFinal = cartForFinalCheck.some(ci => stockLimitMapFinal[ci.name]);
          if (cartHasLimitedItemFinal) {
            const finalStockLock = LockService.getScriptLock();
            let gotFinalLock = false;
            try {
              gotFinalLock = finalStockLock.tryLock(8000);
              if (gotFinalLock) {
                const soldTodayMapFinal = computeSoldTodayByItemName();
                for (const ci of cartForFinalCheck) {
                  const limitFinal = stockLimitMapFinal[ci.name];
                  if (!limitFinal) continue;
                  const alreadySoldFinal = soldTodayMapFinal[ci.name] || 0;
                  const wantQtyFinal = Number(ci.quantity) || 0;
                  if (alreadySoldFinal + wantQtyFinal > limitFinal) {
                    const remainFinal = Math.max(0, limitFinal - alreadySoldFinal);
                    return jsonResponse({ status: 'error', message: `「${ci.name}」剛剛被其他人買完了，目前剩餘 ${remainFinal} 份，請重新調整數量後再試一次` });
                  }
                }
              }
            } finally {
              if (gotFinalLock) finalStockLock.releaseLock();
            }
          }
        }
      }
    } catch (finalStockErr) {
      Logger.log('[庫存核算-最終複查] 失敗，本次不擋單：' + finalStockErr.message);
    }

    const sheet = getOrCreateSheet(SHEET_NAME_ORDERS, [
      '訂單編號', '訂單狀態', '創建時間',
      '顧客姓名', '手機號碼', '顧客LINE_UserID',
      '訂單內容', '總金額',
      '取餐日期', '取餐時間',
      '付款方式', '外送備註', '訂單備註',
      '匯款截圖確認', '管理員備註', '最後更新時間',
      '購物車明細JSON', '匯款截圖網址'
    ]);

    let adminNote = priceMismatch
      ? `⚠️金額異常：客戶端送出$${orderData.total}，後端依目前菜單價格核算為$${serverTotal}，請人工核對後再確認收款`
      : (priceCheck.ok ? '' : `（無法核對金額：${priceCheck.reason}，請人工核對總金額）`);
    if (promoLabelsText) {
      adminNote = (adminNote ? adminNote + '｜' : '') + `🏷️ 促銷優惠：${promoLabelsText}`;
    }
    if (paidByBalance) {
      adminNote = (adminNote ? adminNote + '｜' : '') + `💰 已使用儲值餘額付款$${balanceDeducted}，系統自動核帳，免匯款截圖`;
    } else if (wantUseBalance) {
      adminNote = (adminNote ? adminNote + '｜' : '') + `⚠️客戶想用儲值餘額付款但餘額不足或身份未驗證，已忽略，仍走一般匯款流程`;
    }

    sheet.appendRow([
      orderNumber,
      paidByBalance ? '已匯款' : '未匯款',
      Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
      sanitizeForSheet(orderData.customerName),
      sanitizeForSheet(orderData.customerPhone),
      orderData.customerLineName ? `${sanitizeForSheet(orderData.customerLineName)}｜${verifiedLineUserId || ''}` : (verifiedLineUserId || ''),
      // ★ 2026-07-18 安全修正：items跟paymentMethod原本沒有經過sanitizeForSheet，
      // 雖然正常情況下這兩個欄位是前端自己組出來的（不是客人自由輸入的文字框），
      // 但API本身並沒有嚴格驗證呼叫者一定是自家前端，有心人直接打API帶惡意
      // 字串一樣寫得進來，這裡補上防護，跟其他寫入欄位一致。
      sanitizeForSheet(orderData.items),
      finalTotal,
      orderData.pickupDate,
      orderData.pickupTime,
      sanitizeForSheet(paidByBalance ? '儲值餘額' : (orderData.paymentMethod || '匯款')),
      sanitizeForSheet(orderData.deliveryNote || ''),
      sanitizeForSheet(orderData.orderNotes || ''),
      paidByBalance ? '已確認' : '',
      adminNote,
      Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
      orderData.cartSummary || '[]',
      screenshotUrl
    ]);

    if (paidByBalance) {
      // ★ 安全修正：訂單本身跟儲值扣款記錄是兩個分開的appendRow動作，
      // 如果訂單寫入成功、扣款記錄寫入剛好失敗（例如試算表短暫忙線），
      // 原本會讓整支函式往外拋錯，客人看到「錯誤」以為沒送成功，但試算表
      // 其實已經留了一筆「已匯款」的訂單，扣款記錄卻沒寫進去——客人可能
      // 不知道自己其實有一筆成立的訂單，店家也不知道要對帳。改成把扣款記錄
      // 寫入包進獨立的try/catch，失敗的話不讓整支函式報錯，訂單照樣算
      // 送出成功，但在管理員備註標記警告，店家自己去核對這筆是否真的扣款。
      //
      // ★★ 安全修正：之前判斷「餘額夠不夠」的鎖，只包住「讀取＋標記」這段，
      // 鎖在這裡就放開了，中間隔著訂單寫入、截圖上傳這些會花時間的動作，
      // 真正的扣款記錄寫入是在鎖外面才發生。同一個人開兩個分頁/兩台裝置
      // 幾乎同時送單，理論上兩邊都可能判定「餘額足夠」，各自扣同一筆餘額，
      // 變成雙重扣款。這裡在真正寫入扣款記錄之前，重新加鎖、複查一次
      // 「當下」的餘額夠不夠付這一筆——如果這段時間餘額已經被別的並行請求
      // 扣光了，就不要真的寫扣款記錄，改成標記警告讓店家人工核對，
      // 不會讓客人的餘額被扣成負的。
      let ledgerWriteOk = false;
      const ledgerLock = LockService.getScriptLock();
      let gotLedgerLock = false;
      try {
        gotLedgerLock = ledgerLock.tryLock(8000);
      } catch (lockErr) {
        gotLedgerLock = false;
      }
      try {
        if (gotLedgerLock) {
          const balanceNow = computeBalance(verifiedLineUserId);
          if (balanceNow >= balanceDeducted) {
            const ledgerSheet = getOrCreateSheet(SHEET_NAME_TOPUP_LEDGER, [
              '記錄ID', '時間', 'LINE_UserID', 'LINE_DisplayName', '類型', '金額', '狀態',
              '對應訂單編號', '顧客姓名', '聯絡電話', '管理員備註'
            ]);
            ledgerSheet.appendRow([
              `deduct-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
              Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
              verifiedLineUserId || '',
              sanitizeForSheet(orderData.customerLineName || ''),
              '訂單扣款',
              -balanceDeducted,
              '已確認',
              orderNumber,
              sanitizeForSheet(orderData.customerName || ''),
              sanitizeForSheet(orderData.customerPhone || ''),
              `訂單${orderNumber}扣款`
            ]);
            ledgerWriteOk = true;
          }
        }
      } catch (ledgerErr) {
        Logger.log(`[儲值扣款] 訂單${orderNumber}的扣款記錄寫入失敗，請人工核對：` + ledgerErr.message);
      } finally {
        if (gotLedgerLock) ledgerLock.releaseLock();
      }
      if (!ledgerWriteOk) {
        try {
          sheet.getRange(sheet.getLastRow(), 15).setValue(
            `⚠️扣款記錄未寫入（可能是同時有其他請求搶到餘額，或系統忙線），請人工核對這筆訂單的儲值餘額是否已正確扣除`
          );
        } catch (noteErr) { /* 連標記警告都失敗，至少前面Logger.log還留了紀錄 */ }
      }
    }

    const takeNumber = getTakeNumberForOrder(orderNumber);

    // ★ 新增：訂單成立後，主動推播通知店家自己的LINE，不用一直盯著看板螢幕。
    // 這步失敗完全不影響訂單本身（notifyShopNewOrder內部自己有try/catch兜底）
    notifyShopNewOrder(orderNumber, takeNumber, orderData.customerName, finalTotal, orderData.pickupTime);

    // ★ 2026-07-22 新增：處理推薦好友獎勵，失敗不影響訂單本身
    if (orderData.referralCode && verifiedLineUserId) {
      processReferralReward(orderData.referralCode, verifiedLineUserId, orderData.customerLineName, orderNumber);
    }

    return jsonResponse({
      status: 'success',
      orderNumber: orderNumber,
      takeNumber: takeNumber,
      finalTotal: finalTotal,
      promoDiscount: promoDiscount,
      promoLabelsText: promoLabelsText,
      paidByBalance: paidByBalance,
      balanceDeducted: balanceDeducted,
      screenshotUrl: screenshotUrl
    });

  } catch (err) {
    return errResponse('伺服器錯誤：', err);
  }
}

// ★ 舊版GET路徑：保留給還沒更新到最新版的舊前端相容用，沒有截圖上傳能力，
// 匯款訂單（非儲值付款）一律會被handleSubmitOrderCore裡的強制截圖檢查擋下來，
// 這是刻意設計，逼前端非改用doPost不可才能正常送出匯款訂單。
function handleSubmitOrder(e) {
  try {
    const raw = e.parameter.orderData;
    if (!raw) {
      return jsonResponse({ status: 'error', message: '缺少 orderData 參數' });
    }
    return handleSubmitOrderCore({
      orderData: JSON.parse(raw),
      rawIdToken: e.parameter.idToken || '',
      appToken: e.parameter.appToken || '',
      screenshotBase64: ''
    });
  } catch (err) {
    return errResponse('伺服器錯誤：', err);
  }
}

// ★ 新版POST路徑：支援夾帶匯款截圖（base64圖片資料），這是主要的送單管道。
// 前端要用fetch(url, {method:'POST', body: JSON.stringify({...})})呼叫，
// Content-Type建議用'text/plain;charset=utf-8'（不要用application/json），
// 這是GAS的已知眉角——用application/json瀏覽器會先發一個OPTIONS
// 預檢請求（CORS preflight），GAS的doPost不會正確回應這個預檢，
// 導致整個請求失敗；改用text/plain可以繞開這個問題，GAS這邊一樣能正常
// 用JSON.parse(e.postData.contents)解析內容，不影響資料本身。
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: '缺少請求內容' });
    }
    const body = JSON.parse(e.postData.contents);

    // ★ 安全審計修正：跟doGet同一種問題——原本只判斷body.action ===
    // 'uploadMenuImage'這一種情況，除此之外任何值（打錯字/留空/以後新增
    // 功能但漏寫判斷式）都會落到handleSubmitOrderCore，跟送單請求混在一起。
    // 改成：有帶body.action就必須對到已知值，對不到直接回傳「不支援的操作」，
    // 不落底；沒有帶body.action才是送單路由，且先檢查orderData是否存在。
    if (body.action) {
      if (body.action === 'uploadMenuImage') {
        return handleUploadMenuImage(body);
      }
      // ★ 2026-07-21 補上：菜單編輯存檔用的三個動作，之前漏掉導致
      // 「不支援的操作」，見上方 handleUpdateMenuPost 等函式的說明。
      if (body.action === 'updateMenu') return handleUpdateMenuPost(body);
      if (body.action === 'updateAddons') return handleUpdateAddonsPost(body);
      if (body.action === 'updateOptionSpecs') return handleUpdateOptionSpecsPost(body);
      // ★ 營業日報表系統：寫入類動作（可能資料量大，走POST body不走網址參數）
      if (body.action === 'updateCostCategories') return handleUpdateCostCategories(body);
      if (body.action === 'updateDailyRecord') return handleUpdateDailyRecord(body);
      if (body.action === 'updatePurchaseItems') return handleUpdatePurchaseItems(body);
      if (body.action === 'addPurchaseItem') return handleAddPurchaseItem(body);
      if (body.action === 'submitPurchaseOrder') return handleSubmitPurchaseOrder(body);
      if (body.action === 'receiveLocalDailyReport') return handleReceiveLocalDailyReport(body);
      if (body.action === 'updateFixedCosts') return handleUpdateFixedCosts(body);
      if (body.action === 'setFeatureToggles') return handleSetFeatureToggles(body);
      if (body.action === 'setCustomerNote') return handleSetCustomerNote(body);
      if (body.action === 'setClosureSchedule') return handleSetClosureSchedule(body);
      if (body.action === 'addRefundRecord') return handleAddRefundRecord(body);
      if (body.action === 'submitOrderRating') return handleSubmitOrderRating(body);
      if (body.action === 'setCustomerBirthday') return handleSetCustomerBirthday(body);
      if (body.action === 'addShiftNote') return handleAddShiftNote(body);
      if (body.action === 'clockInOut') return handleClockInOut(body);
      if (body.action === 'saveMenuPreset') return handleSaveMenuPreset(body);
      if (body.action === 'applyMenuPreset') return handleApplyMenuPreset(body);
      return jsonResponse({ status: 'error', message: '不支援的操作' });
    }

    if (!body.orderData) {
      return jsonResponse({ status: 'error', message: '缺少 orderData 參數' });
    }

    return handleSubmitOrderCore({
      orderData: body.orderData,
      rawIdToken: body.idToken || '',
      appToken: body.appToken || '',
      screenshotBase64: body.paymentScreenshotBase64 || ''
    });
  } catch (err) {
    return errResponse('請求格式錯誤：', err);
  }
}

// 菜單讀取
// ────────────────────────────────────────────
// ★ 統計「今天」（依取餐日期）每個品項已經賣出多少份，用來核算安全庫存。
// 只掃「未取消」的訂單，取消的訂單不算數。這支函式會被handleGetMenu
// （顯示用）跟handleSubmitOrderCore（下單當下的強制核算，防超賣）兩處呼叫。
function computeSoldTodayByItemName() {
  const map = {};
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return map;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;

    const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
    const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');

    data.forEach(row => {
      if (row[0] === '') return; // 空列跳過
      const status = row[1];
      if (status === '已取消') return;
      const pickupDate = row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+8', 'yyyy-MM-dd') : String(row[8]);
      if (pickupDate !== todayStr) return;

      let cart = [];
      try { cart = JSON.parse(row[16] || '[]'); } catch (parseErr) { return; }
      cart.forEach(ci => {
        if (!ci || !ci.name) return;
        map[ci.name] = (map[ci.name] || 0) + (Number(ci.quantity) || 0);
      });
    });
  } catch (err) {
    Logger.log('[庫存核算] 統計失敗：' + err.message);
  }
  return map;
}

function handleGetMenu(e) {
  try {
    // ★ 浮水印記錄：只有偵測到不是自己網域打來的請求才記一筆，不擋、不影響
    // 任何人正常使用（包括你自己開發時用localhost測試），純粹留一筆紀錄，
    // 事後想查是誰在別的網域上用你的系統時翻「來源監控」工作表就有線索。
    try { logSuspiciousClientHost(e); } catch (logErr) { /* 記錄失敗不影響菜單讀取 */ }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_MENU);
    if (!sheet) {
      return jsonResponse({ status: 'error', message: '找不到「菜單」工作表，請先執行 setupAllSheets' });
    }

    const lastRow = sheet.getLastRow();
    let items = [];
    if (lastRow >= 2) {
      // ★ 2026-07-23 擴充：讀取範圍從11欄擴大到14欄，加入廚房縮寫/描述/份量
      const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
      items = data
        .filter(row => row[1] !== '')
        .map((row, i) => ({
          id: i,
          category: row[0] || '未分類',
          name: row[1],
          price: row[2],
          icon: row[3],
          status: row[4],
          image: row[5],
          options: row[6] || '',
          cost: Number(row[7]) || 0,
          safetyStock: Number(row[8]) || 0,
          canBeCombo: String(row[9] || '').trim().toUpperCase() === 'Y',
          comboUpgradePrice: Number(row[10]) || 0,
          shortName: row[11] || '',
          description: row[12] || '',
          weight: row[13] || ''
        }));
    }

    // ★ 即時庫存核算：只有設定過安全庫存的品項才需要算，沒設定的（安全庫存=0）
    // 略過不算，省下不必要的運算，因為多數店家的多數品項應該都是不限量供應。
    const itemsWithLimit = items.filter(it => it.safetyStock > 0);
    if (itemsWithLimit.length > 0) {
      const soldTodayMap = computeSoldTodayByItemName();
      items.forEach(it => {
        if (it.safetyStock > 0) {
          it.soldToday = soldTodayMap[it.name] || 0;
        }
      });
    }

    const addons = getAddonsList(ss);
    const optionSpecs = getOptionSpecsList(ss);

    return jsonResponse({ status: 'success', items: items, addons: addons, optionSpecs: optionSpecs, soldOutMessageEnabled: isFeatureEnabled('soldOutMessage') });

  } catch (err) {
    return errResponse('讀取菜單失敗：', err);
  }
}

// ★ 新增（2026-07-19）：判斷現在是「早餐時段」還是「正常時段」，套餐的
// 預設配餐/飲料依這個時段自動切換。分界：11:30之前算早餐時段，跟
// 成本報表既有的早餐時段定義一致（00:00-11:29），只是這裡簡化成
// 二分法（早餐/正常），不像成本報表細分四個時段。
function getCurrentComboPeriod() {
  const now = new Date();
  const hhmm = Utilities.formatDate(now, 'GMT+8', 'HH:mm');
  const parts = hhmm.split(':');
  const minutes = Number(parts[0]) * 60 + Number(parts[1]);
  return minutes < (11 * 60 + 30) ? 'breakfast' : 'normal';
}

// ★ 解析「選項清單」欄位，格式是「名稱:價差,名稱:價差」，例如
// 「黃金薯條:0,生菜沙拉:10,玉米濃湯:5」。如果沒有冒號（舊格式相容），
// 價差當作0處理，不會讓舊資料壞掉。
function parseOptionChoices(raw) {
  return String(raw || '').split(',').map(s => s.trim()).filter(Boolean).map(entry => {
    const parts = entry.split(':');
    return {
      name: parts[0].trim(),
      priceDiff: parts.length > 1 ? (Number(parts[1]) || 0) : 0
    };
  });
}

function getOptionSpecsList(ss) {
  const sheet = ss.getSheetByName(SHEET_NAME_OPTION_SPECS);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const period = getCurrentComboPeriod();
  return data
    .filter(row => row[0] !== '')
    .map((row, i) => {
      const breakfastDefault = row[4] || '';
      const normalDefault = row[5] || '';
      return {
        id: i,
        name: row[0],
        choices: row[1] || '',
        choiceList: parseOptionChoices(row[1]),
        limit: Number(row[2]) || 1,
        required: Number(row[3]) || 0,
        defaultChoice: period === 'breakfast' ? (breakfastDefault || normalDefault) : (normalDefault || breakfastDefault)
      };
    });
}

function getAddonsList(ss) {
  const sheet = ss.getSheetByName(SHEET_NAME_ADDONS);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return data
    .filter(row => row[0] !== '')
    .map((row, i) => ({
      id: i,
      name: row[0],
      price: row[1],
      status: row[2] || '供應中'
    }));
}

// ★ 2026-07-21 新增：POST 版本，供 doPost 呼叫。前端 tool_menu.html 存整個
// 菜單時，三個動作（updateMenu/updateAddons/updateOptionSpecs）都是用 POST
// 送出的（因為品項多、中文字多，網址長度容易爆掉），但 doPost 的路由表
// 一直沒有接住這三個動作名稱，全部落到「不支援的操作」——這是修這個問題。
// body.items/body.addons/body.specs 從 JSON body 解析出來，本來就已經是
// 陣列，不像 doGet 那條路收到的是要再 JSON.parse() 一次的字串。
function handleUpdateMenuPost(body) {
  try {
    const password = body.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

    const items = body.items;
    if (!Array.isArray(items) || items.length === 0) {
      return jsonResponse({ status: 'error', message: '菜單內容不可為空' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_MENU);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME_MENU);

    sheet.clear();
    // ★ 2026-07-23 新增：廚房縮寫、描述、份量三個欄位，加在最後面（不動
    // 前面既有欄位的順序），避免影響其他直接用欄位編號存取的程式碼
    // （例如autoDelistMenuItemsByName用第2欄=名稱、第5欄=狀態）
    const headers = ['分類', '名稱', '價格', '圖標', '狀態', '圖片', '客製選項', '成本', '安全庫存(每日限量，0或空白=不限量)', '可升級套餐(Y/N)', '套餐加購價', '廚房縮寫', '描述', '份量'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    const rows = items.map(it => [
      it.category || '未分類', it.name || '', Number(it.price) || 0, it.icon || '',
      it.status || '供應中', it.image || '', it.options || '', Number(it.cost) || 0,
      Number(it.safetyStock) || '', it.canBeCombo ? 'Y' : '', Number(it.comboUpgradePrice) || 0,
      it.shortName || '', it.description || '', it.weight || ''
    ]);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    logOperatorAction('儲存菜單', `共 ${rows.length} 項`, body.operator);
    return jsonResponse({ status: 'success', message: `菜單已更新，共 ${rows.length} 項` });
  } catch (err) {
    return errResponse('更新菜單失敗：', err);
  }
}

function handleUpdateAddonsPost(body) {
  try {
    const password = body.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

    const addons = body.addons;
    if (!Array.isArray(addons)) return jsonResponse({ status: 'error', message: 'addons 格式錯誤' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_ADDONS);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME_ADDONS);

    sheet.clear();
    const headers = ['名稱', '價格', '狀態'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    if (addons.length > 0) {
      const rows = addons.map(a => [a.name || '', Number(a.price) || 0, a.status || '供應中']);
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    logOperatorAction('儲存加購項目', `共 ${addons.length} 項`, body.operator);
    return jsonResponse({ status: 'success', message: `加購項目已更新，共 ${addons.length} 項` });
  } catch (err) {
    return errResponse('更新加購項目失敗：', err);
  }
}

function handleUpdateOptionSpecsPost(body) {
  try {
    const password = body.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

    const specs = body.specs;
    if (!Array.isArray(specs)) return jsonResponse({ status: 'error', message: 'specs 格式錯誤' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_OPTION_SPECS);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME_OPTION_SPECS);

    sheet.clear();
    const headers = ['名稱', '選項清單', '上限', '必選', '早餐時段預設值', '正常時段預設值'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    if (specs.length > 0) {
      const rows = specs.map(s => [
        s.name || '', s.choices || '', Number(s.limit) || 1, Number(s.required) || 0,
        s.breakfastDefault || '', s.normalDefault || ''
      ]);
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    logOperatorAction('儲存客製選項', `共 ${specs.length} 組`, body.operator);
    return jsonResponse({ status: 'success', message: `選項規格已更新，共 ${specs.length} 組` });
  } catch (err) {
    return errResponse('更新選項規格失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：菜單預設集。把「現在的整份菜單」存成一個有名字的
// 預設集（例如「冬季菜單」「夏季菜單」），之後可以一鍵套用回來，涵蓋
// 兩種用途：(1) 複製昨天/之前的菜單設定 (2) 季節性菜單快速切換——
// 兩者其實是同一套機制，差別只在店家自己怎麼命名、什麼時候套用。
// ────────────────────────────────────────────
const SHEET_NAME_MENU_PRESETS = '菜單預設集';

function handleGetMenuPresets(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('menuPresets')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_MENU_PRESETS);
    if (!sheet) return jsonResponse({ status: 'success', presets: [] });
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', presets: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // 名稱, 儲存時間（不把整份菜單JSON也傳回來，列表不需要那麼大的資料量）
    const presets = data
      .filter(row => row[0] !== '')
      .map(row => ({
        name: row[0],
        savedAt: row[1] instanceof Date ? Utilities.formatDate(row[1], 'GMT+8', 'yyyy-MM-dd HH:mm') : String(row[1])
      }));
    return jsonResponse({ status: 'success', presets: presets });
  } catch (err) {
    return errResponse('讀取菜單預設集失敗：', err);
  }
}

function handleSaveMenuPreset(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('menuPresets')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const presetName = sanitizeForSheet((body.name || '').trim());
    if (!presetName) return jsonResponse({ status: 'error', message: '請輸入預設集名稱' });

    // 直接拿現在的完整菜單（用既有的讀取邏輯，不用自己重寫一次）
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const currentItems = handleGetMenuRaw(ss);
    const itemsJson = JSON.stringify(currentItems);

    const sheet = getOrCreateSheet(SHEET_NAME_MENU_PRESETS, ['名稱', '儲存時間', '菜單內容JSON'], SPREADSHEET_ID);
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= 2) {
      const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < names.length; i++) {
        if (names[i][0] === presetName) { targetRow = i + 2; break; }
      }
    }
    if (targetRow !== -1) {
      // 同名預設集，直接覆蓋，不留舊版本（避免同名越存越多筆搞混）
      sheet.getRange(targetRow, 2, 1, 2).setValues([[new Date(), itemsJson]]);
    } else {
      sheet.appendRow([presetName, new Date(), itemsJson]);
    }
    logOperatorAction('儲存菜單預設集「' + presetName + '」', '', body.operator);
    return jsonResponse({ status: 'success' });
  } catch (err) {
    return errResponse('儲存菜單預設集失敗：', err);
  }
}

function handleApplyMenuPreset(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('menuPresets')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const presetName = (body.name || '').trim();
    if (!presetName) return jsonResponse({ status: 'error', message: '缺少預設集名稱' });

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_MENU_PRESETS);
    if (!sheet) return jsonResponse({ status: 'error', message: '找不到這個預設集' });
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'error', message: '找不到這個預設集' });

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const matched = data.find(row => row[0] === presetName);
    if (!matched) return jsonResponse({ status: 'error', message: '找不到這個預設集' });

    const items = JSON.parse(matched[2]);
    // ★ 直接重用handleUpdateMenu同一套寫入邏輯，用組出一個假的e.parameter
    // 物件呼叫它，不要複製貼上整段寫入試算表的程式碼、維護兩份一樣的邏輯
    const fakeEvent = { parameter: { password: password, items: JSON.stringify(items) } };
    const result = handleUpdateMenu(fakeEvent);
    logOperatorAction('套用菜單預設集「' + presetName + '」', '', body.operator);
    return result;
  } catch (err) {
    return errResponse('套用菜單預設集失敗：', err);
  }
}

function handleUpdateMenu(e) {
  try {
    const password = e.parameter.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) {
      return jsonResponse({ status: 'error', message: authCheck.message });
    }

    const raw = e.parameter.items;
    if (!raw) {
      return jsonResponse({ status: 'error', message: '缺少 items 參數' });
    }

    const items = JSON.parse(raw);
    if (!Array.isArray(items) || items.length === 0) {
      return jsonResponse({ status: 'error', message: '菜單內容不可為空' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_MENU);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME_MENU);
    }

    sheet.clear();
    // ★ 2026-07-23 新增：廚房縮寫、描述、份量三個欄位，加在最後面
    const headers = ['分類', '名稱', '價格', '圖標', '狀態', '圖片', '客製選項', '成本', '安全庫存(每日限量，0或空白=不限量)', '可升級套餐(Y/N)', '套餐加購價', '廚房縮寫', '描述', '份量'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    const rows = items.map(it => [
      it.category || '未分類',
      it.name || '',
      Number(it.price) || 0,
      it.icon || '',
      it.status || '供應中',
      it.image || '',
      it.options || '',
      Number(it.cost) || 0,
      Number(it.safetyStock) || '',
      it.canBeCombo ? 'Y' : '',
      Number(it.comboUpgradePrice) || 0,
      it.shortName || '',
      it.description || '',
      it.weight || ''
    ]);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    return jsonResponse({ status: 'success', message: `菜單已更新，共 ${rows.length} 項` });

  } catch (err) {
    return errResponse('更新菜單失敗：', err);
  }
}

function handleUpdateAddons(e) {
  try {
    const password = e.parameter.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) {
      return jsonResponse({ status: 'error', message: authCheck.message });
    }

    const raw = e.parameter.addons;
    if (!raw) {
      return jsonResponse({ status: 'error', message: '缺少 addons 參數' });
    }

    const addons = JSON.parse(raw);
    if (!Array.isArray(addons)) {
      return jsonResponse({ status: 'error', message: 'addons 格式錯誤' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_ADDONS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME_ADDONS);
    }

    sheet.clear();
    const headers = ['名稱', '價格', '狀態'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    if (addons.length > 0) {
      const rows = addons.map(a => [a.name || '', Number(a.price) || 0, a.status || '供應中']);
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    return jsonResponse({ status: 'success', message: `加購項目已更新，共 ${addons.length} 項` });

  } catch (err) {
    return errResponse('更新加購項目失敗：', err);
  }
}

function handleUpdateOptionSpecs(e) {
  try {
    const password = e.parameter.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) {
      return jsonResponse({ status: 'error', message: authCheck.message });
    }

    const raw = e.parameter.specs;
    if (!raw) {
      return jsonResponse({ status: 'error', message: '缺少 specs 參數' });
    }

    const specs = JSON.parse(raw);
    if (!Array.isArray(specs)) {
      return jsonResponse({ status: 'error', message: 'specs 格式錯誤' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_OPTION_SPECS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME_OPTION_SPECS);
    }

    sheet.clear();
    const headers = ['名稱', '選項清單', '上限', '必選', '早餐時段預設值', '正常時段預設值'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    if (specs.length > 0) {
      const rows = specs.map(s => [
        s.name || '',
        s.choices || '',
        Number(s.limit) || 1,
        Number(s.required) || 0,
        s.breakfastDefault || '',
        s.normalDefault || ''
      ]);
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    return jsonResponse({ status: 'success', message: `選項規格已更新，共 ${specs.length} 組` });

  } catch (err) {
    return errResponse('更新選項規格失敗：', err);
  }
}

function formatPhoneFromSheet(val) {
  let s = String(val == null ? '' : val).trim();
  if (/^9\d{8}$/.test(s)) {
    s = '0' + s;
  }
  return s;
}

function lookupCustomerInfoByLineUserId(lineUserId) {
  if (!lineUserId) return { customerName: '', customerPhone: '' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!sheet) return { customerName: '', customerPhone: '' };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { customerName: '', customerPhone: '' };
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const cellVal = String(row[5] || '');
    if (cellVal && cellVal.indexOf(lineUserId) !== -1) {
      return { customerName: row[3] || '', customerPhone: formatPhoneFromSheet(row[4]) };
    }
  }
  return { customerName: '', customerPhone: '' };
}

// ★ 老客戶自動帶入資料：這支不涉及金錢/隱私外流風險（只回傳「這個LINE帳號自己」的資料），
// 維持原本用lineUserId直接查即可，不強制要求idToken，避免影響頁面載入體驗；
// 如果要更嚴謹，也可以比照下面幾支API加上驗證，做法完全相同。
function handleGetCustomerInfo(e) {
  try {
    // ★ 安全修正：這支API原本只信任網址參數傳來的lineUserId，沒有驗證身份，
    // 代表只要知道任何一組LINE UserID字串，任何人都能查到那個人的姓名+電話，
    // 完全不用證明自己就是那個人。改成強制要求idToken，驗證過後只用「驗證出來
    // 的身份」去查，不採信客人自己填的lineUserId參數——這樣客人只查得到自己的
    // 資料，查不到別人的。
    const rawIdToken = e.parameter.idToken || '';
    if (!rawIdToken) {
      return jsonResponse({ status: 'success', found: false });
    }
    const idVerify = verifyLineIdToken(rawIdToken);
    if (!idVerify.ok) {
      return jsonResponse({ status: 'success', found: false });
    }
    const lineUserId = idVerify.lineUserId;
    const info = lookupCustomerInfoByLineUserId(lineUserId);
    if (info.customerName || info.customerPhone) {
      return jsonResponse({ status: 'success', found: true, customerName: info.customerName, customerPhone: info.customerPhone });
    }
    return jsonResponse({ status: 'success', found: false });

  } catch (err) {
    return errResponse('查詢客戶資料失敗：', err);
  }
}

function computeQueueNumbersForOrders(allRows) {
  const byDate = {};
  allRows.forEach(o => {
    if (o.status === '已取消') return;
    const key = o.pickupDate || '';
    if (!byDate[key]) byDate[key] = [];
    byDate[key].push(o);
  });
  const numberMap = {};
  Object.keys(byDate).forEach(key => {
    const list = byDate[key];
    list.sort((a, b) => String(a.createdAt || '').localeCompare(String(b.createdAt || '')));
    list.forEach((o, i) => { numberMap[o.orderNumber] = i + 1; });
  });
  return numberMap;
}

function getTakeNumberForOrder(orderNumber) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return null;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const allRowsForQueue = data
      .filter(row => row[0] !== '')
      .map(row => ({
        orderNumber: row[0],
        status: row[1],
        createdAt: row[2] instanceof Date ? Utilities.formatDate(row[2], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[2]),
        pickupDate: row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+8', 'yyyy-MM-dd') : String(row[8])
      }));
    const queueNumbers = computeQueueNumbersForOrders(allRowsForQueue);
    return queueNumbers[orderNumber] || null;
  } catch (err) {
    return null;
  }
}

// ────────────────────────────────────────────
// ★ 付款回報：客人不上傳截圖，改成自己填「取餐號＋電話後三碼＋金額」申報付款。
// 這是「客人自己申報」不是「證明付款」，最後還是要店家自己核對銀行明細，
// 但至少省掉截圖上傳這條容易出包的路，改用簡單文字表單，穩定很多。
//
// 自動比對邏輯：拿客人填的取餐號、電話後三碼，去試算表裡找「取餐號算出來
// 相符、電話後三碼也相符」的訂單，優先比對狀態還是「未匯款」的（已經確認過
// 的訂單不需要再比對）。比對得到就記錄對應的訂單編號/客人/金額，方便店家
// 一眼看出這筆回報是哪張訂單；比對不到也照樣收下，只是標記「未比對到」，
// 由店家自己在看板上手動處理，不會因為比對失敗就擋住客人送出申報。
// ────────────────────────────────────────────
function findMatchingOrderForReport(reportedTakeNumber, phoneLast3, reportedAmount) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return null;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

    const rows = data
      .filter(row => row[0] !== '')
      .map(row => ({
        orderNumber: row[0],
        status: row[1],
        createdAt: row[2] instanceof Date ? Utilities.formatDate(row[2], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[2]),
        customerName: row[3],
        customerPhone: formatPhoneFromSheet(row[4]),
        lineUserIdField: row[5], // ★ 原始字串（可能是「顯示名稱｜UserID」組合格式），供身份核對用
        total: row[7],
        pickupDate: row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+8', 'yyyy-MM-dd') : String(row[8])
      }));

    const queueNumbers = computeQueueNumbersForOrders(rows);

    // ★ 取餐號、電話後三碼、金額三個都要對，才算比對成功（老闆明確要求）。
    // 優先選「未匯款」狀態的，避免已經確認過的舊訂單干擾比對結果。
    const candidates = rows.filter(r => {
      const qn = queueNumbers[r.orderNumber];
      const phoneOk = String(r.customerPhone || '').slice(-3) === String(phoneLast3 || '').trim();
      const amountOk = Number(r.total) === Number(reportedAmount);
      return qn && String(qn) === String(reportedTakeNumber).trim() && phoneOk && amountOk;
    });

    if (candidates.length === 0) return null;

    const unpaidFirst = candidates.find(r => r.status === '未匯款');
    return unpaidFirst || candidates[0];
  } catch (err) {
    Logger.log('[付款回報比對] 失敗：' + err.message);
    return null;
  }
}

function handleSubmitPaymentReport(e) {
  try {
    // ★ 安全強化：跟送單、儲值申請一樣，改成強制要求idToken，驗證這個人
    // 真的是誰。由於這個彈窗只會在客人已經走完LIFF入口流程（加好友檢查、
    // 免責聲明）之後才會出現，這時候idToken理論上一定拿得到，改成強制
    // 不會誤傷正常客人。
    const rawIdToken = e.parameter.idToken || '';
    if (!rawIdToken) {
      return jsonResponse({ status: 'error', message: '缺少身份憑證，請透過LINE重新整理頁面後再試一次' });
    }
    const idVerify = verifyLineIdToken(rawIdToken);
    if (!idVerify.ok) {
      return jsonResponse({ status: 'error', message: idVerify.reason });
    }
    const verifiedLineUserId = idVerify.lineUserId;

    // ★ 2026-07-18 安全修正：原本這支API只驗證身份，沒有頻率限制，代表
    // 同一個帳號可以無限次狂送回報，每一次都會打一次LINE verify API、
    // appendRow寫入試算表，有心人可以拿這支洗爆你的每日配額。比照下單的
    // 做法補上限流，1分鐘最多3次，正常客人回報一次付款絕對用不到這麼多次。
    const reportRateCheck = checkPaymentReportRateLimit(verifiedLineUserId);
    if (!reportRateCheck.ok) {
      return jsonResponse({ status: 'error', message: reportRateCheck.reason });
    }

    const takeNumber = (e.parameter.takeNumber || '').trim();
    const phoneLast3 = (e.parameter.phoneLast3 || '').trim();
    const reportedAmount = Number(e.parameter.amount || 0);
    const bankLast3 = (e.parameter.bankLast3 || '').trim(); // ★ 純記錄用，不參與比對，但必填——店家要靠這個去銀行帳戶核對錢真的有沒有到

    const echoValue = (v) => {
      const s = String(v == null ? '' : v).trim();
      if (!s) return '（空白）';
      return s.length > 20 ? s.slice(0, 20) + '…' : s;
    };

    const fieldErrors = [];
    if (!takeNumber) {
      fieldErrors.push('【取餐號碼】未填寫');
    }
    if (!phoneLast3) {
      fieldErrors.push('【電話後三碼】未填寫');
    } else if (!/^\d{3}$/.test(phoneLast3)) {
      fieldErrors.push(`【電話後三碼】格式錯誤，您輸入的是「${echoValue(phoneLast3)}」，須為3碼數字（例如：789）`);
    }
    if (!e.parameter.amount) {
      fieldErrors.push('【餐點總金額】未填寫');
    } else if (!reportedAmount || reportedAmount <= 0) {
      fieldErrors.push(`【餐點總金額】格式錯誤，您輸入的是「${echoValue(e.parameter.amount)}」，須為大於0的數字`);
    }
    if (!bankLast3) {
      fieldErrors.push('【銀行帳號末三碼】未填寫');
    } else if (!/^\d{3}$/.test(bankLast3)) {
      fieldErrors.push(`【銀行帳號末三碼】格式錯誤，您輸入的是「${echoValue(bankLast3)}」，須為3碼數字（例如：456）`);
    }

    if (fieldErrors.length > 0) {
      return jsonResponse({ status: 'error', message: `送出失敗，請修正以下欄位：\n${fieldErrors.join('\n')}` });
    }

    const matched = findMatchingOrderForReport(takeNumber, phoneLast3, reportedAmount);
    if (!matched) {
      return jsonResponse({ status: 'error', message: '查無對應訂單，請確認取餐號、電話後三碼、金額是否正確（可到「查詢我的訂單」核對）' });
    }

    const matchedLineUserId = extractLineUserId(matched.lineUserIdField);
    if (matchedLineUserId !== verifiedLineUserId) {
      return jsonResponse({ status: 'error', message: '查無對應訂單，請確認取餐號、電話後三碼、金額是否正確（可到「查詢我的訂單」核對）' });
    }

    const now = new Date();
    const sheet = getOrCreateSheet(SHEET_NAME_PAYMENT_REPORTS, [
      '回報時間', '客人填的取餐號', '客人填的電話後三碼', '客人回報金額', '銀行末三碼（不比對，店家自行核對）',
      '比對狀態', '對應訂單編號', '對應客人姓名', '對應訂單金額', '店家已確認'
    ]);

    sheet.appendRow([
      Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
      takeNumber,
      phoneLast3,
      reportedAmount,
      bankLast3,
      '已比對到訂單',
      matched.orderNumber,
      matched.customerName,
      matched.total,
      ''
    ]);

    return jsonResponse({
      status: 'success',
      matched: true,
      message: '回報成功，已自動比對到您的訂單，店家確認收款後就會開始準備餐點'
    });
  } catch (err) {
    return errResponse('回報失敗：', err);
  }
}

function handleGetPaymentReports(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_PAYMENT_REPORTS);
    if (!sheet) return jsonResponse({ status: 'success', reports: [] });
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', reports: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const reports = data
      .filter(row => row[0] !== '')
      .map((row, i) => ({
        rowNumber: i + 2,
        reportedAt: row[0] instanceof Date ? Utilities.formatDate(row[0], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[0]),
        takeNumber: row[1],
        phoneLast3: row[2],
        reportedAmount: row[3],
        bankLast3: row[4],
        matchStatus: row[5],
        matchedOrderNumber: row[6],
        matchedCustomerName: row[7],
        matchedOrderTotal: row[8],
        confirmed: row[9]
      }))
      // ★ 按回報時間排序，方便跟手機銀行的轉帳明細（同樣是時間排序）由上往下對照
      .sort((a, b) => String(a.reportedAt).localeCompare(String(b.reportedAt)));

    return jsonResponse({ status: 'success', reports: reports });
  } catch (err) {
    return errResponse('讀取付款回報失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 一鍵確認收款：店家自己去銀行帳戶核對過（用銀行末三碼比對錢真的有沒有
// 到），確認沒問題才按這顆按鈕。這一步刻意保留人工判斷、不自動化——
// 系統比對「取餐號＋電話＋金額」只能證明客人知道這些資訊，沒辦法證明
// 錢真的進到店家帳戶，這道人工核對是防止有人不用真的付款、單靠看畫面
// 資訊就讓系統自動放行出餐的最後一道關卡。
// ────────────────────────────────────────────
function handleConfirmPaymentReport(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const rowNumber = parseInt(e.parameter.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) {
      return jsonResponse({ status: 'error', message: '缺少有效的 rowNumber' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const reportSheet = ss.getSheetByName(SHEET_NAME_PAYMENT_REPORTS);
    if (!reportSheet) return jsonResponse({ status: 'error', message: '找不到付款回報工作表' });

    const rowData = reportSheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
    const matchedOrderNumber = rowData[6]; // 對應訂單編號那一欄
    if (!matchedOrderNumber) {
      return jsonResponse({ status: 'error', message: '這筆回報沒有對應到訂單，無法自動確認' });
    }

    // 找到對應訂單，把狀態改成已匯款
    const orderSheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!orderSheet) return jsonResponse({ status: 'error', message: '找不到訂單資料工作表' });
    const lastRow = orderSheet.getLastRow();
    const orderData = orderSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let orderRowNumber = -1;
    for (let i = 0; i < orderData.length; i++) {
      if (orderData[i][0] === matchedOrderNumber) { orderRowNumber = i + 2; break; }
    }
    if (orderRowNumber === -1) {
      return jsonResponse({ status: 'error', message: `找不到訂單編號 ${matchedOrderNumber} 對應的資料列` });
    }

    orderSheet.getRange(orderRowNumber, 2).setValue('已匯款'); // 訂單狀態欄
    orderSheet.getRange(orderRowNumber, 14).setValue('已確認'); // 匯款截圖確認欄（沿用既有欄位）
    orderSheet.getRange(orderRowNumber, 16).setValue(
      Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss')
    ); // 最後更新時間欄

    // 付款回報那一列標記已確認
    reportSheet.getRange(rowNumber, 10).setValue(
      Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss') + ' 已確認'
    );

    logOperatorAction('確認收款', matchedOrderNumber, e.parameter.operator);
    return jsonResponse({ status: 'success', message: `已將訂單 ${matchedOrderNumber} 標記為已匯款` });
  } catch (err) {
    return errResponse('確認收款失敗：', err);
  }
}

// ────────────────────────────────────────────
// 客人查詢自己的訂單（★ 已整合 idToken 身份驗證）
// ────────────────────────────────────────────
// ────────────────────────────────────────────
// ★ 2026-07-22 新增：客人自己取消訂單。有嚴格限制，不是任何時候都能取消：
//   1. 一定要驗證LINE身份，只能取消「自己下的」訂單，不能用訂單編號亂猜取消別人的
//   2. 訂單狀態一定要還是「未匯款」才能自己取消——已經匯款、已經在準備、
//      已經完成的訂單，代表店家可能已經開始處理，這時候要取消要走
//      「聯絡店家」，不讓客人自己按一按就取消，避免店家白做工
//   3. 一定要在下單後 CANCEL_WINDOW_MINUTES 分鐘之內，超過時間一樣要
//      聯絡店家處理，不是無限期都能自己取消
// ────────────────────────────────────────────
const CANCEL_WINDOW_MINUTES = 10;

// ★ 2026-07-22 新增：查詢取消原因統計，依原因分組計數
function handleGetCancelReasonStats(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('cancelReasonStats')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('取消原因統計');
    if (!sheet) return jsonResponse({ status: 'success', stats: [] });
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', stats: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const countMap = {};
    data.forEach(row => {
      const reason = row[2] || '（未填寫原因）';
      countMap[reason] = (countMap[reason] || 0) + 1;
    });
    const stats = Object.keys(countMap)
      .map(reason => ({ reason, count: countMap[reason] }))
      .sort((a, b) => b.count - a.count);

    return jsonResponse({ status: 'success', stats: stats, totalCount: data.length });
  } catch (err) {
    return errResponse('讀取取消原因統計失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：客人常用取餐時段記憶。查這個客人過去的訂單，找出
// 最常選的取餐時間，點餐表單載入時可以拿來預填，減少客人選擇的步驟。
// ────────────────────────────────────────────
function handleGetMyUsualPickupTime(e) {
  if (!isFeatureEnabled('rememberPickupTime')) return jsonResponse({ status: 'error', message: '此功能未啟用' });
  try {
    const idVerify = verifyLineIdToken(e.parameter.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return jsonResponse({ status: 'success', usualTime: null });
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', usualTime: null });

    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const timeCount = {};
    data.forEach(row => {
      if (row[0] === '' || row[1] === '已取消') return;
      if (String(row[5] || '').indexOf(lineUserId) === -1) return;
      const timeVal = row[9] instanceof Date ? Utilities.formatDate(row[9], 'GMT+8', 'HH:mm') : String(row[9]);
      if (!timeVal) return;
      timeCount[timeVal] = (timeCount[timeVal] || 0) + 1;
    });

    const sorted = Object.keys(timeCount).sort((a, b) => timeCount[b] - timeCount[a]);
    return jsonResponse({ status: 'success', usualTime: sorted.length > 0 ? sorted[0] : null });
  } catch (err) {
    return errResponse('讀取常用取餐時段失敗：', err);
  }
}

function handleCancelMyOrder(e) {
  try {
    const idVerify = verifyLineIdToken(e.parameter.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;

    const orderNumber = (e.parameter.orderNumber || '').trim();
    if (!orderNumber) return jsonResponse({ status: 'error', message: '缺少訂單編號' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return jsonResponse({ status: 'error', message: '找不到訂單資料' });

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] !== orderNumber) continue;

      // 身份確認：這筆訂單的客人資訊欄位裡，要包含目前登入者的LINE UserID
      if (String(row[5] || '').indexOf(lineUserId) === -1) {
        return jsonResponse({ status: 'error', message: '這不是您的訂單，無法取消' });
      }

      const currentStatus = row[1];
      if (currentStatus === '已取消') {
        return jsonResponse({ status: 'error', message: '這筆訂單已經是取消狀態' });
      }
      if (currentStatus !== '未匯款') {
        return jsonResponse({ status: 'error', message: '這筆訂單已經進入處理流程，無法自行取消，請直接聯絡店家' });
      }

      const createdAt = row[2] instanceof Date ? row[2] : new Date(row[2]);
      const minutesSinceOrder = (new Date() - createdAt) / 60000;
      if (minutesSinceOrder > CANCEL_WINDOW_MINUTES) {
        return jsonResponse({ status: 'error', message: `已超過下單後 ${CANCEL_WINDOW_MINUTES} 分鐘，無法自行取消，請直接聯絡店家處理` });
      }

      const rowNumber = i + 2;
      sheet.getRange(rowNumber, 2).setValue('已取消');
      logOperatorAction('客人自行取消訂單', orderNumber, '客人本人（LINE身份已驗證）');

      // ★ 2026-07-22 新增：記錄取消原因（客人選填），累積起來能看出問題
      // 出在哪個環節。獨立一張分頁記錄，不影響訂單本身的資料結構。
      if (isFeatureEnabled('cancelReasonStats') && e.parameter.cancelReason) {
        try {
          const reasonSheet = getOrCreateSheet('取消原因統計', ['時間', '訂單編號', '原因'], SPREADSHEET_ID);
          reasonSheet.appendRow([new Date(), orderNumber, sanitizeForSheet(e.parameter.cancelReason)]);
        } catch (reasonErr) { Logger.log('[取消原因] 記錄失敗（不影響取消本身）：' + reasonErr.message); }
      }

      return jsonResponse({ status: 'success', message: '已成功取消訂單' });
    }

    return jsonResponse({ status: 'error', message: '找不到這筆訂單' });
  } catch (err) {
    return errResponse('取消訂單失敗：', err);
  }
}

function handleGetMyOrders(e) {
  try {
    const idVerify = verifyLineIdToken(e.parameter.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;
    // ★ 日期篩選：客人可以指定「取餐日期」查詢，格式 yyyy-MM-dd，空字串代表不篩選
    // （不篩選時維持原本「近期10筆」的行為；有篩選時放寬到50筆，因為客人明確
    // 指定了某一天，不該被「整體最近10筆」這個上限卡住看不到那天的訂單）
    const filterDate = (e.parameter.filterDate || '').trim();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return jsonResponse({ status: 'success', orders: [] });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', orders: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();

    const allRowsForQueue = data
      .filter(row => row[0] !== '')
      .map(row => ({
        orderNumber: row[0],
        status: row[1],
        createdAt: row[2] instanceof Date ? Utilities.formatDate(row[2], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[2]),
        pickupDate: row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+8', 'yyyy-MM-dd') : String(row[8])
      }));
    const queueNumbers = computeQueueNumbersForOrders(allRowsForQueue);

    let orders = data
      .filter(row => row[0] !== '' && String(row[5] || '').indexOf(lineUserId) !== -1)
      .map(row => {
        const orderNumber = row[0];
        const status = row[1];
        return {
          orderNumber: orderNumber,
          takeNumber: status === '已取消' ? null : (queueNumbers[orderNumber] || null),
          status: status,
          createdAt: row[2] instanceof Date ? Utilities.formatDate(row[2], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[2]),
          items: row[6],
          total: row[7],
          pickupDate: row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+8', 'yyyy-MM-dd') : String(row[8]),
          pickupTime: row[9] instanceof Date ? Utilities.formatDate(row[9], 'GMT+8', 'HH:mm') : String(row[9]),
          cartSummary: row[16] || '[]' // ★ 2026-07-22 新增：給「再訂一次」功能用的購物車明細JSON
        };
      })
      .reverse();

    if (filterDate) {
      orders = orders.filter(o => o.pickupDate === filterDate);
    }
    orders = orders.slice(0, filterDate ? 50 : 10);

    return jsonResponse({ status: 'success', orders: orders, customerRatingEnabled: isFeatureEnabled('customerRating') });

  } catch (err) {
    return errResponse('查詢訂單失敗：', err);
  }
}

function getPromoRulesList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_PROMO_RULES);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  return data
    .filter(row => row[0] !== '')
    .map(row => {
      let params = {};
      try { params = JSON.parse(row[4] || '{}'); } catch (e) { params = {}; }
      const itemNames = String(row[5] || '').split(',').map(s => s.trim()).filter(Boolean);
      return {
        id: row[0],
        name: row[1] || '',
        type: row[2] || '',
        enabled: row[3] === true || row[3] === 'TRUE' || row[3] === '啟用',
        params: params,
        itemNames: itemNames
      };
    });
}

function handleGetPromoRules() {
  try {
    return jsonResponse({ status: 'success', rules: getPromoRulesList() });
  } catch (err) {
    return errResponse('讀取優惠規則失敗：', err);
  }
}

function handleUpdatePromoRules(e) {
  try {
    const password = e.parameter.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) {
      return jsonResponse({ status: 'error', message: authCheck.message });
    }

    const raw = e.parameter.rules;
    if (!raw) {
      return jsonResponse({ status: 'error', message: '缺少 rules 參數' });
    }
    const rules = JSON.parse(raw);
    if (!Array.isArray(rules)) {
      return jsonResponse({ status: 'error', message: 'rules 格式錯誤' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_PROMO_RULES);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME_PROMO_RULES);

    sheet.clear();
    const headers = ['規則ID', '名稱', '類型', '啟用', '參數JSON', '指定品項'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    if (rules.length > 0) {
      const rows = rules.map(r => [
        r.id || `rule-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        r.name || '',
        r.type || '',
        r.enabled ? '啟用' : '停用',
        JSON.stringify(r.params || {}),
        (r.itemNames || []).join(',')
      ]);
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    return jsonResponse({ status: 'success', message: `優惠規則已更新，共 ${rules.length} 條` });
  } catch (err) {
    return errResponse('更新優惠規則失敗：', err);
  }
}

function handleGetAnnouncements() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ANNOUNCEMENTS);
    if (!sheet) return jsonResponse({ status: 'success', announcements: [] });
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', announcements: [] });
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const announcements = data
      .filter(row => row[0] !== '' && row[1] !== '')
      .map(row => ({
        id: row[0],
        content: row[1] || '',
        enabled: row[2] === true || row[2] === 'TRUE' || row[2] === '啟用',
        displayType: row[3] || '跑馬燈'
      }));
    return jsonResponse({ status: 'success', announcements: announcements });
  } catch (err) {
    return errResponse('讀取促銷公告失敗：', err);
  }
}

function handleUpdateAnnouncements(e) {
  try {
    const password = e.parameter.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) {
      return jsonResponse({ status: 'error', message: authCheck.message });
    }

    const raw = e.parameter.announcements;
    if (!raw) {
      return jsonResponse({ status: 'error', message: '缺少 announcements 參數' });
    }
    const announcements = JSON.parse(raw);
    if (!Array.isArray(announcements)) {
      return jsonResponse({ status: 'error', message: 'announcements 格式錯誤' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_ANNOUNCEMENTS);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME_ANNOUNCEMENTS);

    sheet.clear();
    const headers = ['公告ID', '內容', '啟用', '顯示方式'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    if (announcements.length > 0) {
      const rows = announcements.map(a => [
        a.id || `ann-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        a.content || '',
        a.enabled ? '啟用' : '停用',
        a.displayType || '跑馬燈'
      ]);
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    return jsonResponse({ status: 'success', message: `促銷公告已更新，共 ${announcements.length} 則` });
  } catch (err) {
    return errResponse('更新促銷公告失敗：', err);
  }
}

function getTopupTiersList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_TOPUP_TIERS);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return data
    .filter(row => row[0] !== '')
    .map(row => ({
      id: row[0],
      payAmount: Number(row[1]) || 0,
      creditAmount: Number(row[2]) || 0,
      enabled: row[3] === true || row[3] === 'TRUE' || row[3] === '啟用'
    }));
}

function handleGetTopupTiers() {
  try {
    const tiers = getTopupTiersList().filter(t => t.enabled);
    return jsonResponse({ status: 'success', tiers: tiers });
  } catch (err) {
    return errResponse('讀取儲值方案失敗：', err);
  }
}

function handleUpdateTopupTiers(e) {
  try {
    const password = e.parameter.password || '';
    const authCheck = verifyAdminPassword(password);
    if (!authCheck.ok) {
      return jsonResponse({ status: 'error', message: authCheck.message });
    }
    const raw = e.parameter.tiers;
    if (!raw) return jsonResponse({ status: 'error', message: '缺少 tiers 參數' });
    const tiers = JSON.parse(raw);
    if (!Array.isArray(tiers)) return jsonResponse({ status: 'error', message: 'tiers 格式錯誤' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_TOPUP_TIERS);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME_TOPUP_TIERS);

    sheet.clear();
    const headers = ['方案ID', '儲值金額', '到帳金額', '啟用'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    if (tiers.length > 0) {
      const rows = tiers.map(t => [
        t.id || `tier-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        Number(t.payAmount) || 0,
        Number(t.creditAmount) || 0,
        t.enabled ? '啟用' : '停用'
      ]);
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    return jsonResponse({ status: 'success', message: `儲值方案已更新，共 ${tiers.length} 組` });
  } catch (err) {
    return errResponse('更新儲值方案失敗：', err);
  }
}

function computeBalance(lineUserId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_TOPUP_LEDGER);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  let balance = 0;
  data.forEach(row => {
    const rowLineUserId = String(row[2] || '');
    const status = row[6];
    const amount = Number(row[5]) || 0;
    if (rowLineUserId === lineUserId && status === '已確認') {
      balance += amount;
    }
  });
  return balance;
}

// ────────────────────────────────────────────
// 客人查詢自己的儲值餘額（★ 已整合 idToken 身份驗證）
// ────────────────────────────────────────────
function handleGetBalance(e) {
  try {
    const idVerify = verifyLineIdToken(e.parameter.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;
    return jsonResponse({ status: 'success', balance: computeBalance(lineUserId) });
  } catch (err) {
    return errResponse('查詢餘額失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 食品責任免責聲明：第一次點餐的客人要先同意過才能點餐，同意過一次
// 就記住，之後不用每次都跳出來——跟儲值規範（每次都要重新同意）刻意
// 設計成不同行為，因為這個是「客人身份」層級的一次性確認，不是每次
// 交易都要重新確認的東西。用idToken驗證出來的lineUserId當唯一識別，
// 存進試算表，換裝置、清瀏覽器快取都查得到「這個人同意過了」。
// ────────────────────────────────────────────
function handleCheckDisclaimerAgreed(e) {
  try {
    const idVerify = verifyLineIdToken(e.parameter.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;

    // ★ 最早的「用戶資料」寫入點：開頁面就記，不等點餐或同意聲明
    registerUserIfNew(lineUserId, idVerify.lineDisplayName);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_DISCLAIMER);
    if (!sheet) return jsonResponse({ status: 'success', agreed: false });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', agreed: false });

    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const agreed = data.some(row => row[0] === lineUserId);
    return jsonResponse({ status: 'success', agreed: agreed });
  } catch (err) {
    return errResponse('查詢同意狀態失敗：', err);
  }
}

function handleAgreeDisclaimer(e) {
  try {
    const idVerify = verifyLineIdToken(e.parameter.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;

    // ★ 保險：就算客人不知為何跳過了checkDisclaimerAgreed直接打這支，
    // 這裡也一樣會確保「用戶資料」記到人（registerUserIfNew內部本來就
    // 有防重複記錄的判斷，重複呼叫是安全的）
    registerUserIfNew(lineUserId, idVerify.lineDisplayName);

    const sheet = getOrCreateSheet(SHEET_NAME_DISCLAIMER, ['LINE_UserID', '同意時間']);

    // ★ 防重複記錄：如果已經同意過，不要再新增一筆一樣的，直接回傳成功就好
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      if (data.some(row => row[0] === lineUserId)) {
        return jsonResponse({ status: 'success', message: '已經同意過了' });
      }
    }

    sheet.appendRow([lineUserId, Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss')]);
    return jsonResponse({ status: 'success', message: '已記錄同意' });
  } catch (err) {
    return errResponse('記錄同意失敗：', err);
  }
}

// ────────────────────────────────────────────
// 客人提出儲值申請（★ 已整合 idToken 身份驗證）
// ────────────────────────────────────────────
function handleSubmitTopupRequest(e) {
  try {
    const idVerify = verifyLineIdToken(e.parameter.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;

    // ★ 2026-07-18 安全修正：原本這支API只驗證身份，沒有頻率限制，
    // 同一個帳號可以無限次狂送儲值申請，每一次都消耗一次LINE verify API
    // 呼叫額度＋一次試算表寫入，有心人可以拿這支洗爆你的每日配額。
    // 比照下單的做法補上限流，1分鐘最多3次。
    const topupRateCheck = checkTopupRateLimit(lineUserId);
    if (!topupRateCheck.ok) {
      return jsonResponse({ status: 'error', message: topupRateCheck.reason });
    }

    const lineDisplayName = (e.parameter.lineDisplayName || idVerify.lineDisplayName || '').trim();
    const payAmount = Number(e.parameter.payAmount || 0);

    if (payAmount <= 0) return jsonResponse({ status: 'error', message: '儲值金額不正確' });

    const tiers = getTopupTiersList().filter(t => t.enabled);
    const matchedTier = tiers.find(t => t.payAmount === payAmount);
    if (!matchedTier) {
      return jsonResponse({ status: 'error', message: '找不到對應的儲值方案，可能方案已異動，請重新整理頁面再試' });
    }
    const creditAmount = matchedTier.creditAmount;

    const historicalInfo = lookupCustomerInfoByLineUserId(lineUserId);
    const customerName = historicalInfo.customerName || lineDisplayName || '（未知）';
    const customerPhone = historicalInfo.customerPhone || '';

    const recordId = `topup-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
    const now = new Date();
    const sheet = getOrCreateSheet(SHEET_NAME_TOPUP_LEDGER, [
      '記錄ID', '時間', 'LINE_UserID', 'LINE_DisplayName', '類型', '金額', '狀態',
      '對應訂單編號', '顧客姓名', '聯絡電話', '管理員備註'
    ]);
    sheet.appendRow([
      recordId,
      Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
      lineUserId,
      sanitizeForSheet(lineDisplayName),
      '儲值申請',
      creditAmount,
      '待確認',
      '',
      sanitizeForSheet(customerName),
      sanitizeForSheet(customerPhone),
      `客人儲值$${payAmount}，方案到帳$${creditAmount}，待店家核對匯款截圖後確認`
    ]);

    return jsonResponse({ status: 'success', message: '儲值申請已送出，請匯款後將截圖傳給店家確認', recordId: recordId, payAmount: payAmount, creditAmount: creditAmount });
  } catch (err) {
    return errResponse('儲值申請失敗：', err);
  }
}

function handleGetTopupRequests(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_TOPUP_LEDGER);
    if (!sheet) return jsonResponse({ status: 'success', records: [] });
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', records: [] });
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const records = data
      .filter(row => row[0] !== '')
      .map((row, i) => ({
        rowNumber: i + 2,
        recordId: row[0],
        createdAt: row[1] instanceof Date ? Utilities.formatDate(row[1], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[1]),
        lineUserId: row[2],
        lineDisplayName: row[3],
        type: row[4],
        amount: row[5],
        status: row[6],
        orderNumber: row[7],
        customerName: row[8],
        customerPhone: formatPhoneFromSheet(row[9]),
        adminNote: row[10]
      }))
      .reverse();
    return jsonResponse({ status: 'success', records: records });
  } catch (err) {
    return errResponse('讀取儲值記錄失敗：', err);
  }
}

function handleConfirmTopupRequest(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const rowNumber = parseInt(e.parameter.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) return jsonResponse({ status: 'error', message: '缺少有效的 rowNumber' });
    const approve = e.parameter.approve !== 'false';

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_TOPUP_LEDGER);
    if (!sheet) return jsonResponse({ status: 'error', message: '找不到儲值記錄工作表' });

    sheet.getRange(rowNumber, 7).setValue(approve ? '已確認' : '已取消');
    logOperatorAction(approve ? '確認儲值入帳' : '駁回儲值申請', 'row ' + rowNumber, e.parameter.operator);
    return jsonResponse({ status: 'success', message: approve ? '已確認入帳' : '已駁回這筆儲值申請' });
  } catch (err) {
    return errResponse('確認儲值失敗：', err);
  }
}

function calcPromoDiscount(cartItems, promoRules) {
  const rawSubtotal = Math.round(cartItems.reduce((s, ci) => s + (ci.unitPrice * ci.quantity), 0));
  if (rawSubtotal <= 0 || !promoRules || promoRules.length === 0) {
    return { discount: 0, labels: [], rawSubtotal: rawSubtotal };
  }

  let bestDiscount = 0;
  let bestLabel = null;

  promoRules.forEach(rule => {
    if (!rule.enabled) return;

    // ★ 2026-07-22 新增：限時搶購——規則的params裡可以額外加上「幾點到
    // 幾點」的時段限制（params.timeWindow = {start:'14:00', end:'17:00'}），
    // 只有在這個時段內下單才會套用，不在時段內就當作這條規則不存在，跳過。
    // 沒有設timeWindow的規則不受影響，維持全天候都適用。存在params裡面
    // 是因為params本來就是自由格式的JSON欄位，不用另外改動試算表結構。
    if (isFeatureEnabled('timeLimitedDeal') && rule.params && rule.params.timeWindow && rule.params.timeWindow.start && rule.params.timeWindow.end) {
      const nowTimeStr = Utilities.formatDate(new Date(), 'GMT+8', 'HH:mm');
      if (nowTimeStr < rule.params.timeWindow.start || nowTimeStr > rule.params.timeWindow.end) return;
    }

    const hasItemFilter = rule.itemNames && rule.itemNames.length > 0;
    const eligible = cartItems.filter(ci => !hasItemFilter || rule.itemNames.includes(ci.name));
    const eligibleSubtotal = Math.round(eligible.reduce((s, ci) => s + (ci.unitPrice * ci.quantity), 0));
    const totalQty = eligible.reduce((s, ci) => s + ci.quantity, 0);
    const allPrices = eligible.flatMap(ci => Array(ci.quantity).fill(ci.unitPrice)).sort((a, b) => a - b);

    let ruleDiscount = 0;
    let ruleLabel = '';
    const p = rule.params || {};

    if (rule.type === 'BUYXGETY') {
      if (eligible.length === 0) return;
      const buyX = parseInt(p.buyX || 5, 10);
      const getY = parseInt(p.getY || 1, 10);
      const freeCount = Math.floor(totalQty / (buyX + getY)) * getY;
      if (freeCount > 0) {
        for (let i = 0; i < freeCount; i++) ruleDiscount += allPrices[i] || 0;
        ruleLabel = rule.name || `買${buyX}送${getY}`;
      }
    } else if (rule.type === 'NTH_DISCOUNT') {
      if (eligible.length === 0) return;
      const nth = parseInt(p.nth || 2, 10);
      const rate = parseFloat(p.rate || 0.5);
      const count = Math.floor(totalQty / nth);
      if (count > 0) {
        for (let i = 0; i < count; i++) {
          ruleDiscount += Math.round((allPrices[nth - 1 + i * nth] || 0) * (1 - rate));
        }
        ruleLabel = rule.name || `第${nth}件${Math.round((1 - rate) * 10)}折`;
      }
    } else if (rule.type === 'PERCENT_OFF') {
      const threshold = Number(p.threshold || 0);
      const rate = parseFloat(p.rate || 0.9);
      const base = hasItemFilter ? eligibleSubtotal : rawSubtotal;
      if (base >= threshold && (!hasItemFilter || eligible.length > 0)) {
        ruleDiscount = Math.round(base * (1 - rate));
        ruleLabel = rule.name || (threshold > 0 ? `滿${threshold}打${Math.round(rate * 10)}折` : `打${Math.round(rate * 10)}折`);
      }
    } else if (rule.type === 'THRESHOLD_GIFT') {
      const threshold = Number(p.threshold || 0);
      const giftValue = Number(p.giftValue || 0);
      if (rawSubtotal >= threshold && threshold > 0 && giftValue > 0) {
        ruleDiscount = giftValue;
        const giftName = p.giftItemName || '好禮';
        const giftQty = parseInt(p.giftQty || 1, 10);
        ruleLabel = rule.name || `滿${threshold}送${giftName}x${giftQty}`;
      }
    }

    if (ruleDiscount > bestDiscount) {
      bestDiscount = ruleDiscount;
      bestLabel = { text: ruleLabel, amount: ruleDiscount, ruleId: rule.id };
    }
  });

  const finalDiscount = Math.min(Math.round(bestDiscount), rawSubtotal);
  return {
    discount: finalDiscount < 0 ? 0 : finalDiscount,
    labels: (finalDiscount < 0 || !bestLabel) ? [] : [bestLabel],
    rawSubtotal: rawSubtotal
  };
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：簡易營業報表。統計指定日期（預設今天）的訂單，
// 算出總營收、訂單數、品項銷售排行——「已取消」的訂單不計入統計，
// 避免虛報營收。這裡用「下單時間（createdAt）」當統計基準，不是取餐
// 時間，比較符合「今天做了多少生意」這個直覺。
// ────────────────────────────────────────────
// ────────────────────────────────────────────
// ★ 2026-07-22 新增：備料清單自動產生。把「今天待處理」（未匯款/已匯款/
// 準備中，還沒完成也沒取消）的訂單，購物車明細裡的品項全部加總，算出
// 「今天總共要準備幾份雞排、幾份薯條」，不用自己一筆一筆數。
// ────────────────────────────────────────────
function handleGetPrepList(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('prepList')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return jsonResponse({ status: 'success', items: [] });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', items: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
    const pendingStatuses = ['未匯款', '已匯款', '準備中'];
    const itemQtyMap = {};
    let orderCount = 0;

    data.forEach(row => {
      if (row[0] === '') return;
      if (pendingStatuses.indexOf(row[1]) === -1) return;
      orderCount++;
      try {
        const cart = JSON.parse(row[16] || '[]');
        cart.forEach(entry => {
          const name = entry.name || '未知品項';
          const qty = Number(entry.quantity) || 1;
          itemQtyMap[name] = (itemQtyMap[name] || 0) + qty;
        });
      } catch (parseErr) { /* 這筆訂單購物車格式異常，跳過，不影響其他筆 */ }
    });

    const items = Object.keys(itemQtyMap)
      .map(name => ({ name, quantity: itemQtyMap[name] }))
      .sort((a, b) => b.quantity - a.quantity);

    return jsonResponse({ status: 'success', items: items, orderCount: orderCount });
  } catch (err) {
    return errResponse('讀取備料清單失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：星期幾/時段銷售熱力圖。統計最近90天內，每個「星期幾
// x 幾點」的訂單數量，用顏色深淺看出哪個時段最多客人，幫助排班、備料。
// 用「下單時間」而不是「取餐時間」統計，比較能反映「客人習慣什麼時候
// 上門/下單」這個規律。
// ────────────────────────────────────────────
// ────────────────────────────────────────────
// ★ 2026-07-22 新增：回購率追蹤。統計「歷史上有過訂單的不重複客人」裡，
// 有多少比例是「訂購過2次以上」的回頭客，用來看客人黏著度好不好。
// ────────────────────────────────────────────
function handleGetRepeatCustomerRate(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('repeatCustomerRate')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return jsonResponse({ status: 'success', totalCustomers: 0, repeatCustomers: 0, repeatRate: 0 });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', totalCustomers: 0, repeatCustomers: 0, repeatRate: 0 });

    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const orderCountByCustomer = {};
    data.forEach(row => {
      if (row[0] === '' || row[1] === '已取消') return;
      const lineUserId = extractLineUserId(row[5]);
      if (!lineUserId) return;
      orderCountByCustomer[lineUserId] = (orderCountByCustomer[lineUserId] || 0) + 1;
    });

    const totalCustomers = Object.keys(orderCountByCustomer).length;
    const repeatCustomers = Object.values(orderCountByCustomer).filter(c => c >= 2).length;
    const repeatRate = totalCustomers > 0 ? Math.round((repeatCustomers / totalCustomers) * 100) : 0;

    return jsonResponse({ status: 'success', totalCustomers, repeatCustomers, repeatRate });
  } catch (err) {
    return errResponse('讀取回購率失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：簡易會員等級/累積消費徽章。依照這位客人歷史上
// （非取消訂單）的累積消費金額，給一個等級稱號，門檻先寫死幾個常見級距，
// 之後如果要調整，直接改MEMBER_TIERS這個陣列就好。
// ────────────────────────────────────────────
const MEMBER_TIERS = [
  { min: 5000, name: '🏆 金牌會員', color: '#d97706' },
  { min: 2000, name: '🥈 銀牌會員', color: '#64748b' },
  { min: 500, name: '🥉 銅牌會員', color: '#92400e' },
  { min: 0, name: '🌱 新朋友', color: '#16a34a' }
];

function computeMemberTier(totalSpent) {
  for (const tier of MEMBER_TIERS) {
    if (totalSpent >= tier.min) return tier;
  }
  return MEMBER_TIERS[MEMBER_TIERS.length - 1];
}

function handleGetMyMemberTier(e) {
  if (!isFeatureEnabled('memberTier')) return jsonResponse({ status: 'error', message: '此功能未啟用' });
  try {
    const idVerify = verifyLineIdToken(e.parameter.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    let totalSpent = 0;
    let orderCount = 0;
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
        data.forEach(row => {
          if (row[0] === '' || row[1] === '已取消') return;
          if (String(row[5] || '').indexOf(lineUserId) === -1) return;
          totalSpent += Number(row[7]) || 0;
          orderCount++;
        });
      }
    }
    const tier = computeMemberTier(totalSpent);
    const nextTier = MEMBER_TIERS.filter(t => t.min > totalSpent).sort((a, b) => a.min - b.min)[0];

    return jsonResponse({
      status: 'success',
      totalSpent, orderCount,
      tierName: tier.name, tierColor: tier.color,
      nextTierName: nextTier ? nextTier.name : null,
      amountToNextTier: nextTier ? nextTier.min - totalSpent : 0
    });
  } catch (err) {
    return errResponse('讀取會員等級失敗：', err);
  }
}

function handleGetSalesHeatmap(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('salesHeatmap')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const RECENT_DAYS = 90;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - RECENT_DAYS);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    // grid[星期幾(0=日...6=六)][小時(0-23)] = 訂單數
    const grid = Array.from({ length: 7 }, () => Array(24).fill(0));

    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // 訂單編號, 狀態, 建立時間
        data.forEach(row => {
          if (row[0] === '' || row[1] === '已取消') return;
          const createdAt = row[2] instanceof Date ? row[2] : new Date(row[2]);
          if (isNaN(createdAt) || createdAt < cutoff) return;
          const weekday = parseInt(Utilities.formatDate(createdAt, 'GMT+8', 'u'), 10) % 7; // 轉成0=週日...6=週六
          const hour = parseInt(Utilities.formatDate(createdAt, 'GMT+8', 'H'), 10);
          grid[weekday][hour]++;
        });
      }
    }

    return jsonResponse({ status: 'success', grid: grid, days: RECENT_DAYS });
  } catch (err) {
    return errResponse('讀取銷售熱力圖失敗：', err);
  }
}

function handleGetSalesSummary(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) {
    return jsonResponse({ status: 'error', message: authCheck.message });
  }
  try {
    const targetDate = (e.parameter.date || '').trim() || Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return jsonResponse({ status: 'success', date: targetDate, orderCount: 0, totalRevenue: 0, itemRanking: [] });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', date: targetDate, orderCount: 0, totalRevenue: 0, itemRanking: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();

    let orderCount = 0;
    let totalRevenue = 0;
    const itemQtyMap = {};

    data.forEach(row => {
      if (row[0] === '') return;
      const status = row[1];
      if (status === '已取消') return; // 取消的訂單不計入營收統計

      const createdAt = row[2] instanceof Date ? row[2] : new Date(row[2]);
      const createdDateStr = Utilities.formatDate(createdAt, 'GMT+8', 'yyyy-MM-dd');
      if (createdDateStr !== targetDate) return;

      orderCount++;
      totalRevenue += Number(row[7]) || 0;

      try {
        const cart = JSON.parse(row[16] || '[]');
        cart.forEach(entry => {
          const name = entry.name || '未知品項';
          const qty = Number(entry.quantity) || 1;
          itemQtyMap[name] = (itemQtyMap[name] || 0) + qty;
        });
      } catch (parseErr) { /* 這筆訂單的購物車格式有問題，跳過品項統計，不影響營收加總 */ }
    });

    const itemRanking = Object.keys(itemQtyMap)
      .map(name => ({ name, quantity: itemQtyMap[name] }))
      .sort((a, b) => b.quantity - a.quantity)
      .slice(0, 10); // 只取賣最好的前10名，避免品項太多列一大串

    return jsonResponse({
      status: 'success',
      date: targetDate,
      orderCount,
      totalRevenue,
      itemRanking
    });
  } catch (err) {
    return errResponse('讀取營業報表失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：熟客標記/備註。用「客人電話號碼」當識別鍵（訂單本來
// 就一定會有電話，比LINE UserID更直覺，店員在訂單看板上看到的也是電話），
// 存在一張新的分頁「顧客備註」，一支電話一列，之後每次改備註直接覆蓋
// 掉那一列，不用保留歷史版本（備註本來就是「目前最新的認知」，不需要
// 版本紀錄）。
// ────────────────────────────────────────────
const SHEET_NAME_CUSTOMER_NOTES = '顧客備註';

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：交接班記事本。簡單的共用留言板，早班寫給晚班的
// 小提醒（例如「今天雞肉少一批，要注意」），不用另外開LINE群組講。
// ────────────────────────────────────────────
const SHEET_NAME_SHIFT_NOTES = '交接班記事';

function handleGetShiftNotes(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('shiftHandoverNotes')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_SHIFT_NOTES);
    if (!sheet) return jsonResponse({ status: 'success', notes: [] });
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', notes: [] });

    // 只回傳最近30筆，太多會很長，舊的參考價值也不高
    const startRow = Math.max(2, lastRow - 29);
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
    const notes = data
      .filter(row => row[0] !== '')
      .map(row => ({
        time: row[0] instanceof Date ? Utilities.formatDate(row[0], 'GMT+8', 'yyyy-MM-dd HH:mm') : String(row[0]),
        author: row[1],
        note: row[2]
      }))
      .reverse();
    return jsonResponse({ status: 'success', notes: notes });
  } catch (err) {
    return errResponse('讀取交接班記事失敗：', err);
  }
}

function handleAddShiftNote(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('shiftHandoverNotes')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const note = sanitizeForSheet(body.note || '').trim();
    if (!note) return jsonResponse({ status: 'error', message: '請輸入內容' });
    const sheet = getOrCreateSheet(SHEET_NAME_SHIFT_NOTES, ['時間', '寫的人', '內容'], SPREADSHEET_ID);
    sheet.appendRow([new Date(), sanitizeForSheet(body.operator || '未知'), note]);
    return jsonResponse({ status: 'success' });
  } catch (err) {
    return errResponse('新增交接班記事失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：員工出勤打卡。簡單的上下班打卡記錄，跟訂餐系統
// 整合在同一個入口，不用另外裝打卡APP。用「操作人員」名稱識別是誰
// （跟其他功能的operator欄位同一套機制），不做複雜的排班比對。
// ────────────────────────────────────────────
const SHEET_NAME_CLOCK_RECORDS = '員工打卡紀錄';

function handleClockInOut(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('staffClockIn')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const type = body.type === 'out' ? '下班' : '上班';
    const operator = sanitizeForSheet(body.operator || '未知');
    const sheet = getOrCreateSheet(SHEET_NAME_CLOCK_RECORDS, ['時間', '員工', '類型'], SPREADSHEET_ID);
    sheet.appendRow([new Date(), operator, type]);
    return jsonResponse({ status: 'success', message: `已記錄${type}打卡` });
  } catch (err) {
    return errResponse('打卡失敗：', err);
  }
}

function handleGetClockRecords(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('staffClockIn')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_CLOCK_RECORDS);
    if (!sheet) return jsonResponse({ status: 'success', records: [] });
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', records: [] });

    const startRow = Math.max(2, lastRow - 49); // 最近50筆
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
    const records = data
      .filter(row => row[0] !== '')
      .map(row => ({
        time: row[0] instanceof Date ? Utilities.formatDate(row[0], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[0]),
        staff: row[1],
        type: row[2]
      }))
      .reverse();
    return jsonResponse({ status: 'success', records: records });
  } catch (err) {
    return errResponse('讀取打卡紀錄失敗：', err);
  }
}

function handleGetCustomerNotes(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('customerNote')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_CUSTOMER_NOTES);
    if (!sheet) return jsonResponse({ status: 'success', notes: {} });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', notes: {} });

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // 電話, 備註
    const notes = {};
    data.forEach(row => {
      if (row[0]) notes[String(row[0]).trim()] = row[1] || '';
    });
    return jsonResponse({ status: 'success', notes: notes });
  } catch (err) {
    return errResponse('讀取顧客備註失敗：', err);
  }
}

function handleSetCustomerNote(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('customerNote')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const phone = (body.phone || '').trim();
    const note = sanitizeForSheet(body.note || '');
    if (!phone) return jsonResponse({ status: 'error', message: '缺少電話號碼' });

    const sheet = getOrCreateSheet(SHEET_NAME_CUSTOMER_NOTES, ['電話', '備註'], SPREADSHEET_ID);
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= 2) {
      const phones = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < phones.length; i++) {
        if (String(phones[i][0]).trim() === phone) { targetRow = i + 2; break; }
      }
    }
    if (note === '') {
      // 備註清空，直接刪掉這一列，不留空白列
      if (targetRow !== -1) sheet.deleteRow(targetRow);
    } else if (targetRow !== -1) {
      sheet.getRange(targetRow, 2).setValue(note);
    } else {
      sheet.appendRow([phone, note]);
    }
    logOperatorAction('更新顧客備註', phone, body.operator);
    return jsonResponse({ status: 'success' });
  } catch (err) {
    return errResponse('更新顧客備註失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：正式退款/補償紀錄。獨立一張分頁記錄每一筆退款/補償，
// 不再只是寫在訂單備註裡——之後可以回頭統計「這個月退了幾次、都是什麼
// 原因」，找出問題出在哪個環節。
// ────────────────────────────────────────────
const SHEET_NAME_REFUNDS = '退款補償紀錄';

function handleGetRefundRecords(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('refundRecord')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_REFUNDS);
    if (!sheet) return jsonResponse({ status: 'success', records: [] });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', records: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const records = data
      .filter(row => row[0] !== '')
      .map(row => ({
        time: row[0] instanceof Date ? Utilities.formatDate(row[0], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[0]),
        orderNumber: row[1],
        amount: Number(row[2]) || 0,
        reason: row[3],
        operator: row[4],
        note: row[5]
      }))
      .reverse();
    return jsonResponse({ status: 'success', records: records });
  } catch (err) {
    return errResponse('讀取退款紀錄失敗：', err);
  }
}

function handleAddRefundRecord(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('refundRecord')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const orderNumber = (body.orderNumber || '').trim();
    const amount = Number(body.amount) || 0;
    const reason = sanitizeForSheet(body.reason || '');
    if (amount <= 0) return jsonResponse({ status: 'error', message: '退款金額要大於0' });
    if (!reason) return jsonResponse({ status: 'error', message: '請填寫退款原因' });

    const sheet = getOrCreateSheet(SHEET_NAME_REFUNDS, ['時間', '訂單編號', '金額', '原因', '操作人員', '備註'], SPREADSHEET_ID);
    sheet.appendRow([
      new Date(),
      orderNumber,
      amount,
      reason,
      sanitizeForSheet(body.operator || ''),
      sanitizeForSheet(body.note || '')
    ]);
    logOperatorAction('登記退款/補償', `${orderNumber}／$${amount}／${reason}`, body.operator);
    return jsonResponse({ status: 'success' });
  } catch (err) {
    return errResponse('登記退款失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：供應商價格比較。同一個品項，如果最近曾經跟不同
// 廠商叫過貨，列出「上次跟誰買、多少錢」，方便比較挑便宜的。用「最近
// 60天內」的紀錄，太久以前的價格參考價值不大（食材價格會波動）。
// ────────────────────────────────────────────
function handleGetSupplierPriceCompare(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('supplierPriceCompare')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const RECENT_DAYS = 60;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - RECENT_DAYS);

    const reportSs = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    const sheet = reportSs.getSheetByName(SHEET_NAME_DAILY_RECORDS);
    if (!sheet) return jsonResponse({ status: 'success', comparisons: [] });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', comparisons: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    // itemSupplierRecords: { 品名: [{supplier, amount, date}, ...] }
    const itemSupplierRecords = {};

    data.forEach(row => {
      const dateVal = row[0];
      const dateObj = dateVal instanceof Date ? dateVal : new Date(dateVal);
      if (isNaN(dateObj) || dateObj < cutoff) return;
      try {
        const entries = JSON.parse(row[1] || '[]');
        entries.forEach(en => {
          if (en.category !== '食材成本') return;
          const name = (en.item || '').trim();
          const supplier = (en.supplier || '').trim();
          if (!name || !supplier) return; // 沒填廠商的紀錄，沒辦法拿來比較，跳過
          if (!itemSupplierRecords[name]) itemSupplierRecords[name] = [];
          itemSupplierRecords[name].push({
            supplier: supplier,
            amount: Number(en.amount) || 0,
            date: Utilities.formatDate(dateObj, 'GMT+8', 'yyyy-MM-dd')
          });
        });
      } catch (parseErr) { /* 跳過格式異常的那一天 */ }
    });

    // 只列出「同一個品項，曾經跟2家以上不同廠商叫過貨」的，才有比較的意義
    const comparisons = [];
    Object.keys(itemSupplierRecords).forEach(name => {
      const records = itemSupplierRecords[name];
      const supplierSet = new Set(records.map(r => r.supplier));
      if (supplierSet.size < 2) return;

      // 每家廠商，取「最近一次」的金額當代表
      const bySupplierLatest = {};
      records.forEach(r => {
        if (!bySupplierLatest[r.supplier] || r.date > bySupplierLatest[r.supplier].date) {
          bySupplierLatest[r.supplier] = r;
        }
      });
      const supplierPrices = Object.values(bySupplierLatest).sort((a, b) => a.amount - b.amount);
      comparisons.push({ item: name, suppliers: supplierPrices });
    });

    return jsonResponse({ status: 'success', comparisons: comparisons });
  } catch (err) {
    return errResponse('讀取供應商比價失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：系統健康度儀表板。一個畫面總覽幾個關鍵狀態，不用
// 東翻西找才知道系統有沒有問題。
// ────────────────────────────────────────────
function handleGetSystemHealth(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('systemHealthDashboard')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  const health = {
    scriptVersion: SCRIPT_VERSION,
    checkedAt: Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss')
  };

  // 1. 每日自動備份有沒有設定觸發器
  try {
    health.backupTriggerActive = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'backupSpreadsheet');
  } catch (err) { health.backupTriggerActive = null; }

  // 2. 最近一次備份是什麼時候
  try {
    const folder = getOrCreateDriveFolder('咱的台雞店_試算表備份');
    let latestDate = null;
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const created = f.getDateCreated();
      if (!latestDate || created > latestDate) latestDate = created;
    }
    health.lastBackupTime = latestDate ? Utilities.formatDate(latestDate, 'GMT+8', 'yyyy-MM-dd HH:mm') : null;
    health.lastBackupDaysAgo = latestDate ? Math.floor((new Date() - latestDate) / 86400000) : null;
  } catch (err) {
    health.lastBackupTime = null;
    health.lastBackupDaysAgo = null;
  }

  // 3. 今日訂單概況：有沒有訂單卡在「未匯款」太久（超過2小時還沒處理，
  // 可能是客人忘了付款、或店家漏看，兩種情況都值得提醒關注一下）
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    let stuckUnpaidCount = 0;
    let todayOrderCount = 0;
    const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
        const twoHoursAgo = new Date(Date.now() - 2 * 60 * 60 * 1000);
        data.forEach(row => {
          if (row[0] === '') return;
          const createdAt = row[2] instanceof Date ? row[2] : new Date(row[2]);
          const createdDateStr = Utilities.formatDate(createdAt, 'GMT+8', 'yyyy-MM-dd');
          if (createdDateStr === todayStr) todayOrderCount++;
          if (row[1] === '未匯款' && createdAt < twoHoursAgo) stuckUnpaidCount++;
        });
      }
    }
    health.todayOrderCount = todayOrderCount;
    health.stuckUnpaidCount = stuckUnpaidCount;
  } catch (err) {
    health.todayOrderCount = null;
    health.stuckUnpaidCount = null;
  }

  // 4. 目前的營業狀態（跟公休設定放在同一個儀表板，一目了然）
  try {
    const storeStatus = getStoreOpenStatus();
    health.storeIsOpen = storeStatus.isOpen;
  } catch (err) { health.storeIsOpen = null; }

  // 5. ★ 2026-07-22 新增：出餐時間統計——今天已完成的訂單，平均花多久時間
  // （從下單到標記完成），抓「今天」是因為時效性最有參考價值，太久以前的
  // 資料拿來比較意義不大
  if (isFeatureEnabled('prepTimeStats')) {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
      const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');
      let totalMinutes = 0;
      let completedCount = 0;
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
          const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
          data.forEach(row => {
            if (row[0] === '' || row[1] !== '已完成') return;
            const createdAt = row[2] instanceof Date ? row[2] : new Date(row[2]);
            const updatedAt = row[15] instanceof Date ? row[15] : new Date(row[15]);
            const createdDateStr = Utilities.formatDate(createdAt, 'GMT+8', 'yyyy-MM-dd');
            if (createdDateStr !== todayStr) return;
            if (isNaN(updatedAt) || updatedAt < createdAt) return; // 資料異常，跳過避免算出負數
            totalMinutes += (updatedAt - createdAt) / 60000;
            completedCount++;
          });
        }
      }
      health.avgPrepMinutes = completedCount > 0 ? Math.round(totalMinutes / completedCount) : null;
      health.completedTodayCount = completedCount;
    } catch (err) {
      health.avgPrepMinutes = null;
      health.completedTodayCount = null;
    }
  }

  return jsonResponse({ status: 'success', health: health });
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：訂單匯出成CSV，選一段期間，把訂單明細整理成表格
// 格式，方便自己另外整理或給會計看。CSV可以直接用Excel打開，不用另外
// 產生真正的.xlsx二進位檔案（GAS沒有內建的簡單方式產生.xlsx，CSV是
// 最實際、最不容易出錯的做法，Excel/Google試算表都能直接開）。
// ────────────────────────────────────────────
function handleExportOrdersCsv(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });
  if (!isFeatureEnabled('exportOrdersCsv')) return jsonResponse({ status: 'error', message: '此功能未啟用' });

  try {
    const fromStr = (e.parameter.from || '').trim();
    const toStr = (e.parameter.to || '').trim();
    if (!fromStr || !toStr) return jsonResponse({ status: 'error', message: '缺少 from/to 參數' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return jsonResponse({ status: 'success', csv: '' });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', csv: '' });

    const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    const rows = [['訂單編號', '狀態', '下單時間', '姓名', '電話', '內容', '金額', '取餐日期', '取餐時間', '付款方式']];

    data.forEach(row => {
      if (row[0] === '') return;
      const createdAt = row[2] instanceof Date ? row[2] : new Date(row[2]);
      const dateStr = Utilities.formatDate(createdAt, 'GMT+8', 'yyyy-MM-dd');
      if (dateStr < fromStr || dateStr > toStr) return;
      rows.push([
        row[0], row[1],
        Utilities.formatDate(createdAt, 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
        row[3], row[4], row[6], row[7],
        row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+8', 'yyyy-MM-dd') : row[8],
        row[9] instanceof Date ? Utilities.formatDate(row[9], 'GMT+8', 'HH:mm') : row[9],
        row[10]
      ]);
    });

    // 轉成CSV文字：每個欄位用雙引號包起來、內部雙引號跳脫，避免內容裡有逗號/換行搞亂欄位
    const csv = rows.map(r => r.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\r\n');
    return jsonResponse({ status: 'success', csv: csv, rowCount: rows.length - 1 });
  } catch (err) {
    return errResponse('匯出訂單失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 2026-07-22 新增：客人評分/意見回饋。取餐完成後，客人可以給這筆訂單
// 打星星評分、寫幾句話。獨立一張分頁記錄，用「訂單編號」防止同一筆
// 訂單重複評分（客人只能評一次，避免洗評分）。
// ────────────────────────────────────────────
const SHEET_NAME_RATINGS = '顧客評分';

function handleSubmitOrderRating(body) {
  if (!isFeatureEnabled('customerRating')) return jsonResponse({ status: 'error', message: '此功能未啟用' });
  try {
    const idVerify = verifyLineIdToken(body.idToken || '');
    if (!idVerify.ok) return jsonResponse({ status: 'error', message: idVerify.reason });
    const lineUserId = idVerify.lineUserId;

    const orderNumber = (body.orderNumber || '').trim();
    const stars = Number(body.stars) || 0;
    if (!orderNumber) return jsonResponse({ status: 'error', message: '缺少訂單編號' });
    if (stars < 1 || stars > 5) return jsonResponse({ status: 'error', message: '評分要在1-5顆星之間' });

    // 確認這筆訂單真的是這個人的、而且已經完成，才允許評分
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ordersSheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!ordersSheet) return jsonResponse({ status: 'error', message: '找不到訂單資料' });
    const lastRow = ordersSheet.getLastRow();
    const data = ordersSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const matched = data.find(row => row[0] === orderNumber && String(row[5] || '').indexOf(lineUserId) !== -1);
    if (!matched) return jsonResponse({ status: 'error', message: '找不到這筆訂單，或不是您本人的訂單' });
    if (matched[1] !== '已完成') return jsonResponse({ status: 'error', message: '這筆訂單還沒完成，無法評分' });

    const ratingSheet = getOrCreateSheet(SHEET_NAME_RATINGS, ['時間', '訂單編號', 'LINE_UserID', '星數', '意見'], SPREADSHEET_ID);
    const ratingLastRow = ratingSheet.getLastRow();
    if (ratingLastRow >= 2) {
      const existing = ratingSheet.getRange(2, 2, ratingLastRow - 1, 1).getValues();
      if (existing.some(row => row[0] === orderNumber)) {
        return jsonResponse({ status: 'error', message: '這筆訂單已經評分過了' });
      }
    }

    ratingSheet.appendRow([new Date(), orderNumber, lineUserId, stars, sanitizeForSheet(body.comment || '')]);
    return jsonResponse({ status: 'success' });
  } catch (err) {
    return errResponse('送出評分失敗：', err);
  }
}

function handleGetOrders(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) {
    return jsonResponse({ status: 'error', message: authCheck.message });
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return jsonResponse({ status: 'success', orders: [] });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', orders: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    const orders = data
      .filter(row => row[0] !== '')
      .map((row, i) => ({
        rowNumber: i + 2,
        orderNumber: row[0],
        status: row[1],
        createdAt: row[2] instanceof Date ? Utilities.formatDate(row[2], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[2]),
        customerName: row[3],
        customerPhone: formatPhoneFromSheet(row[4]),
        items: row[6],
        total: row[7],
        pickupDate: row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+8', 'yyyy-MM-dd') : String(row[8]),
        pickupTime: row[9] instanceof Date ? Utilities.formatDate(row[9], 'GMT+8', 'HH:mm') : String(row[9]),
        paymentMethod: row[10],
        orderNotes: row[12],
        paymentConfirmed: row[13],
        adminNote: row[14],
        updatedAt: row[15] instanceof Date ? Utilities.formatDate(row[15], 'GMT+8', 'yyyy-MM-dd HH:mm:ss') : String(row[15] || ''), // ★ 2026-07-22 新增：多人鎖定機制用
        // ★ 補上：本地日報表系統(data_core.py)需要這個欄位才能算出品項銷售排行，
        // 資料本來getRange就有讀進來(row[16])，只是組裝回傳物件時漏掉沒放進去
        cartSummary: row[16] || '[]',
        screenshotUrl: row[17] || ''
      }))
      .reverse();

    return jsonResponse({ status: 'success', orders: orders, pickupTimeSortEnabled: isFeatureEnabled('pickupTimeSort') });
  } catch (err) {
    return errResponse('讀取訂單失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 新增（2026-07-21）：操作紀錄。記錄「誰、什麼時候、對哪筆資料做了
// 什麼動作」，是資安建議裡「每一筆會改資料的操作都要留紀錄」的第一步
// 實作。operatorName 來自前端登入時抓到的 LINE 顯示名稱（順手利用已經
// 驗證過的 LIFF 身份，沒有另外做一套帳號系統），沒有值時記成
//「（未知操作者）」，不影響功能、只是查不到是誰。
// 這支函式故意「絕不拋錯」——記錄失敗只寫進 Logger.log，不能因為記錄
// 這個附加動作失敗，就讓真正的業務操作（例如改訂單狀態）也跟著失敗。
// ────────────────────────────────────────────
function logOperatorAction(actionType, targetDescription, operatorName) {
  try {
    const sheet = getOrCreateSheet('操作紀錄', ['時間', '操作人', '動作', '對象']);
    sheet.appendRow([
      Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
      sanitizeForSheet(operatorName || '（未知操作者）'),
      sanitizeForSheet(actionType || ''),
      sanitizeForSheet(targetDescription || '')
    ]);
  } catch (err) {
    Logger.log('[操作紀錄] 寫入失敗（不影響原本操作）：' + err.message);
  }
}

function handleUpdateOrderStatus(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) {
    return jsonResponse({ status: 'error', message: authCheck.message });
  }

  try {
    const rowNumber = parseInt(e.parameter.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) {
      return jsonResponse({ status: 'error', message: '缺少有效的 rowNumber' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return jsonResponse({ status: 'error', message: '找不到訂單資料工作表' });

    // ★ 2026-07-22 新增：多人同時處理訂單的狀態鎖定。用「最後更新時間」
    // 當版本號比對——前端讀取訂單列表時，會帶著當時看到的「最後更新時間」
    // 一起送過來；如果現在試算表上的「最後更新時間」跟前端帶來的不一樣，
    // 代表在這中間，已經有別人（或自己另一個分頁）改過這筆訂單了，這次
    // 更新直接擋下來，不要覆蓋掉別人的操作，請操作的人重新整理再試一次。
    // 沒有帶expectedUpdatedAt參數（例如舊版前端還沒更新）就不做這個檢查，
    // 維持原本行為，避免因為前端版本沒同步而完全卡死。
    if (isFeatureEnabled('orderEditLock') && e.parameter.expectedUpdatedAt) {
      const currentUpdatedAt = sheet.getRange(rowNumber, 16).getValue();
      const currentUpdatedAtStr = currentUpdatedAt instanceof Date
        ? Utilities.formatDate(currentUpdatedAt, 'GMT+8', 'yyyy-MM-dd HH:mm:ss')
        : String(currentUpdatedAt);
      if (currentUpdatedAtStr !== e.parameter.expectedUpdatedAt) {
        return jsonResponse({
          status: 'error',
          message: '這筆訂單剛剛已經被其他人更新過了，請重新整理畫面確認最新狀態後再操作',
          conflict: true
        });
      }
    }

    const newStatus = e.parameter.status;
    if (newStatus) {
      sheet.getRange(rowNumber, 2).setValue(newStatus);
      // ★ 訂單編號在第1欄，一併記進紀錄，方便日後查是哪一筆訂單被改的
      const orderNumberForLog = sheet.getRange(rowNumber, 1).getValue();
      logOperatorAction('更新訂單狀態為「' + newStatus + '」', String(orderNumberForLog), e.parameter.operator);

      // ★ 2026-07-22 新增：訂單完成時，主動推播LINE通知客人來取餐，不用
      // 客人自己一直開頁面查詢。這步失敗完全不影響狀態更新本身。
      if (newStatus === '已完成' && isFeatureEnabled('pickupNotifyCustomer')) {
        try {
          const orderRowForNotify = sheet.getRange(rowNumber, 1, 1, 6).getValues()[0];
          const lineUserIdForNotify = extractLineUserId(orderRowForNotify[5]);
          if (lineUserIdForNotify) {
            const msg = `🎉 您的餐點已經準備好囉！\n訂單編號：${orderRowForNotify[0]}\n可以來取餐了～`;
            sendLinePushMessage(lineUserIdForNotify, msg);
          }
        } catch (notifyErr) {
          Logger.log('[取餐通知] 推播給客人失敗（不影響訂單狀態）：' + notifyErr.message);
        }
      }
    }
    if (e.parameter.paymentConfirmed) {
      sheet.getRange(rowNumber, 14).setValue(e.parameter.paymentConfirmed);
    }
    if (e.parameter.adminNote !== undefined) {
      sheet.getRange(rowNumber, 15).setValue(e.parameter.adminNote);
    }
    sheet.getRange(rowNumber, 16).setValue(
      Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss')
    );

    // ★ 業務邏輯修正：訂單取消時，如果當初是用儲值餘額付款，要自動把錢退回去，
    // 不然客人的錢就這樣憑空消失，系統不會自動退，之前完全沒有這段邏輯。
    // 用「有沒有已經退過款」判斷，避免店家不小心把狀態切來切去、觸發重複退款。
    if (newStatus === '已取消') {
      try {
        const orderRowFull = sheet.getRange(rowNumber, 1, 1, 18).getValues()[0];
        const orderNumberForRefund = orderRowFull[0];
        const paymentMethodForRefund = orderRowFull[10];
        const orderTotalForRefund = Number(orderRowFull[7]) || 0;
        const lineUserIdForRefund = extractLineUserId(orderRowFull[5]);

        if (paymentMethodForRefund === '儲值餘額' && orderTotalForRefund > 0 && lineUserIdForRefund) {
          const ledgerSheet = getOrCreateSheet(SHEET_NAME_TOPUP_LEDGER, [
            '記錄ID', '時間', 'LINE_UserID', 'LINE_DisplayName', '類型', '金額', '狀態',
            '對應訂單編號', '顧客姓名', '聯絡電話', '管理員備註'
          ]);
          const ledgerLastRow = ledgerSheet.getLastRow();
          let alreadyRefunded = false;
          if (ledgerLastRow >= 2) {
            const ledgerData = ledgerSheet.getRange(2, 1, ledgerLastRow - 1, 8).getValues();
            alreadyRefunded = ledgerData.some(row => row[4] === '訂單取消退款' && row[7] === orderNumberForRefund);
          }
          if (!alreadyRefunded) {
            ledgerSheet.appendRow([
              `refund-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
              Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
              lineUserIdForRefund,
              '',
              '訂單取消退款',
              orderTotalForRefund,
              '已確認',
              orderNumberForRefund,
              '',
              '',
              `訂單${orderNumberForRefund}取消，自動退回儲值餘額$${orderTotalForRefund}`
            ]);
          }
        }
      } catch (refundErr) {
        Logger.log('[取消退款] 處理失敗，請人工核對餘額：' + refundErr.message);
        // ★ 就算退款處理失敗，也不要讓整個「取消訂單」這個動作失敗卡住，
        // 訂單狀態還是要能正常改成已取消，退款失敗只記錄下來，之後人工補救
      }
    }

    let notified = null;
    if (newStatus === '已完成') {
      const rowData = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
      const orderNumber = rowData[0];
      const lineUserId = extractLineUserId(rowData[5]);
      const pickupTime = rowData[9] instanceof Date ? Utilities.formatDate(rowData[9], 'GMT+8', 'HH:mm') : String(rowData[9]);
      if (lineUserId) {
        const msg = `🍗 咱的台雞店\n您的訂單 ${orderNumber} 已經完成囉！\n📍歡迎現在前來取餐～\n（取餐時間預約：${pickupTime}）`;
        const result = sendLinePushMessage(lineUserId, msg);
        notified = result.ok;
        if (!result.ok) {
          Logger.log(`[LINE推播] 訂單${orderNumber}通知失敗：${result.reason}`);
        }
      }
    }

    return jsonResponse({ status: 'success', message: '訂單已更新', notified: notified });
  } catch (err) {
    return errResponse('更新訂單失敗：', err);
  }
}

function recomputeOrderTotal(cartSummaryJson) {
  const MAX_QTY_PER_ITEM = 10;

  if (!cartSummaryJson) {
    return { ok: false, total: 0, reason: '前端未提供購物車明細，無法核算' };
  }
  let cart;
  try {
    cart = JSON.parse(cartSummaryJson);
    if (!Array.isArray(cart)) throw new Error('格式錯誤');
  } catch (e) {
    return { ok: false, total: 0, reason: '購物車明細格式錯誤' };
  }

  const overLimit = cart.filter(entry => (Number(entry.quantity) || 0) > MAX_QTY_PER_ITEM);
  if (overLimit.length > 0) {
    return { ok: false, total: 0, reason: `品項數量超過單筆上限(${MAX_QTY_PER_ITEM})：${overLimit.map(e => e.name).join('、')}`, blocked: true };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const menuItems = handleGetMenuRaw(ss);
    const addonItems = getAddonsList(ss);
    const optionSpecs = getOptionSpecsList(ss);

    const menuByName = {};
    menuItems.forEach(it => { menuByName[it.name] = it; });
    const addonPriceByName = {};
    addonItems.forEach(a => { addonPriceByName[a.name] = Number(a.price) || 0; });
    const optionSpecByName = {};
    optionSpecs.forEach(spec => { optionSpecByName[spec.name] = spec; });

    let total = 0;
    const unknownItems = [];
    // ★ 這裡是套餐邏輯真正被驗證的地方——不管前端算出來的金額是多少，
    // 後端一律用自己查到的菜單價格＋套餐加購價＋選項價差重新算一次，
    // 完全不採信客戶端送來的任何價格數字。

    cart.forEach(entry => {
      const qty = Number(entry.quantity) || 1;
      const item = menuByName[entry.name];
      if (!item) {
        unknownItems.push(entry.name);
        return;
      }

      if (entry.isCombo) {
        // ★ 安全檢查：這個品項在菜單上根本沒開放升級套餐，卻宣稱是套餐，
        // 直接當成「品項對不到菜單」處理，擋下這筆訂單，不採信客戶端
        // 自己說「這是套餐」這件事。
        if (!item.canBeCombo) {
          unknownItems.push(entry.name + '（未開放套餐升級）');
          return;
        }
        let lineTotal = (Number(item.price) || 0) + (Number(item.comboUpgradePrice) || 0);

        const combo = entry.comboChoices || {};
        Object.keys(combo).forEach(groupName => {
          const spec = optionSpecByName[groupName];
          const chosenName = combo[groupName];
          if (!spec) {
            unknownItems.push(`套餐選項群組不存在：${groupName}`);
            return;
          }
          const matched = spec.choiceList.find(c => c.name === chosenName);
          if (!matched) {
            unknownItems.push(`套餐選項不存在：${groupName}-${chosenName}`);
            return;
          }
          lineTotal += matched.priceDiff;
        });

        total += lineTotal * qty;
      } else {
        total += (Number(item.price) || 0) * qty;
      }

      (entry.addons || []).forEach(addonName => {
        if (addonPriceByName.hasOwnProperty(addonName)) {
          total += addonPriceByName[addonName];
        } else {
          unknownItems.push('加購:' + addonName);
        }
      });
    });

    if (unknownItems.length > 0) {
      return { ok: false, total: Math.round(total), reason: `品項對不到菜單（可能菜單已更動）：${unknownItems.join('、')}` };
    }
    return { ok: true, total: Math.round(total) };

  } catch (err) {
    return { ok: false, total: 0, reason: '核算過程發生錯誤：' + err.message };
  }
}

function handleGetMenuRaw(ss) {
  const sheet = ss.getSheetByName(SHEET_NAME_MENU);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // ★ 2026-07-23 擴充：讀取範圍從11欄擴大到14欄，把廚房縮寫/描述/份量也
  // 一併讀進來——這支函式除了給成本比對用，也被handleSaveMenuPreset拿去
  // 存快照，沒讀到這三個新欄位的話，套用預設集時會把這些資料弄丟
  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  return data
    .filter(row => row[1] !== '')
    .map(row => ({
      category: row[0],
      name: row[1],
      price: row[2],
      icon: row[3],
      status: row[4],
      image: row[5],
      options: row[6],
      cost: Number(row[7]) || 0,
      safetyStock: Number(row[8]) || 0,
      canBeCombo: String(row[9] || '').trim().toUpperCase() === 'Y',
      comboUpgradePrice: Number(row[10]) || 0,
      shortName: row[11] || '',
      description: row[12] || '',
      weight: row[13] || ''
    }));
}

// ────────────────────────────────────────────
// ★ 新增（2026-07-19）：簡易版成本報表
// 設計原則：不新增「進貨記錄」這種要店家逐筆登記的東西，食材成本直接用
// 「菜單」表本來就有的「成本」欄位（店家設定菜單時填一次），乘以當天賣出
// 的數量加總即可，操作負擔最低。固定成本則是店家每月填一次總額，系統自動
// 依當月天數分攤到每一天。前端只負責顯示，所有運算都在這裡完成。
//
// ★ 2026-07-20：這一整組「成本報表」相關函式，資料來源改成
// REPORT_SPREADSHEET_ID（獨立的營業日報表試算表）。要注意的是，成本報表
// 需要「訂單資料/菜單」（在A表點餐主表）跟「固定成本/每日營運記錄」
// （在B表營業日報表）兩邊的資料一起運算，所以 handleGetCostReport() 內部
// 會同時開兩個ss（ss指向A表、reportSs指向B表），其餘只碰B表資料的函式，
// 一律固定開 REPORT_SPREADSHEET_ID。
// ────────────────────────────────────────────

function getFixedCostsList() {
  const ss = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_FIXED_COSTS);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return data
    .filter(row => row[0] !== '')
    .map(row => ({ name: row[0], amount: Number(row[1]) || 0 }));
}

function getCostCategoriesData() {
  const ss = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_COST_CATEGORIES);
  if (!sheet) return { expenseCategories: [], supplierCategories: [] };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { expenseCategories: [], supplierCategories: [] };
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const expenseCategories = [];
  const supplierCategories = [];
  data.forEach(row => {
    if (!row[1]) return;
    const item = { key: row[1], name: row[2] || '' };
    if (row[0] === 'expense') expenseCategories.push(item);
    else if (row[0] === 'supplier') supplierCategories.push(item);
  });
  return { expenseCategories, supplierCategories };
}

function handleGetCostCategories(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const data = getCostCategoriesData();
    return jsonResponse({ status: 'success', expenseCategories: data.expenseCategories, supplierCategories: data.supplierCategories });
  } catch (err) {
    return errResponse('讀取成本類別失敗：', err);
  }
}

function handleUpdateCostCategories(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const expenseCategories = Array.isArray(body.expenseCategories) ? body.expenseCategories : [];
    const supplierCategories = Array.isArray(body.supplierCategories) ? body.supplierCategories : [];

    const sheet = getOrCreateSheet(SHEET_NAME_COST_CATEGORIES, ['類型', 'Key', '名稱'], REPORT_SPREADSHEET_ID);
    sheet.clear();
    sheet.getRange(1, 1, 1, 3).setValues([['類型', 'Key', '名稱']]);
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    const rows = [];
    expenseCategories.forEach(c => rows.push(['expense', c.key || '', sanitizeForSheet(c.name || '')]));
    supplierCategories.forEach(c => rows.push(['supplier', c.key || '', sanitizeForSheet(c.name || '')]));
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }

    return jsonResponse({ status: 'success', message: '成本類別已更新' });
  } catch (err) {
    return errResponse('更新成本類別失敗：', err);
  }
}

function getDailyRecordForDate(ss, dateStr) {
  const sheet = ss.getSheetByName(SHEET_NAME_DAILY_RECORDS);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  for (let i = 0; i < data.length; i++) {
    const rowDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'GMT+8', 'yyyy-MM-dd') : String(data[i][0]);
    if (rowDate === dateStr) {
      let entries = [], revenue = {};
      try { entries = JSON.parse(data[i][1] || '[]'); } catch (e1) {}
      try { revenue = JSON.parse(data[i][2] || '{}'); } catch (e2) {}
      return { date: dateStr, entries, revenue, note: data[i][3] || '' };
    }
  }
  return null;
}

// ★ 新增：給一個月份(YYYY-MM)，掃這個月「1號到指定日期」所有已存的每日記錄，
// 依分類（expenses的每個key、purchases的每個key）加總，回傳累計金額，
// 供前端「累計」欄位自動顯示用，不用店家自己手動累加。
function handleGetMonthlyCategoryTotals(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const uptoDate = (e.parameter.uptoDate || '').trim();
    if (!uptoDate) return jsonResponse({ status: 'error', message: '缺少 uptoDate 參數' });

    const monthStr = uptoDate.slice(0, 7);
    const ss = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_DAILY_RECORDS);

    // ★ 改成依「科別」（固定/非固定/食材成本）加總累計金額，配合entries
    // 自由列表的新格式——廠商/項目是自由文字、數量不固定，沒辦法再依
    // 固定key加總，改成依這3個科別分類即可
    const categoryTotals = { '固定': 0, '非固定': 0, '食材成本': 0 };
    let revenueTotal = 0;

    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
        data.forEach(row => {
          const rowDate = row[0] instanceof Date ? Utilities.formatDate(row[0], 'GMT+8', 'yyyy-MM-dd') : String(row[0]);
          if (rowDate.slice(0, 7) !== monthStr || rowDate > uptoDate) return;

          let entries = [], revenue = {};
          try { entries = JSON.parse(row[1] || '[]'); } catch (e1) {}
          try { revenue = JSON.parse(row[2] || '{}'); } catch (e2) {}

          entries.forEach(en => {
            if (categoryTotals.hasOwnProperty(en.category)) {
              categoryTotals[en.category] += Number(en.amount) || 0;
            }
          });
          Object.values(revenue).forEach(r => {
            revenueTotal += (Number(r.card) || 0) + (Number(r.cash) || 0) + (Number(r.other) || 0);
          });
        });
      }
    }

    return jsonResponse({ status: 'success', categoryTotals: categoryTotals, revenueTotal: revenueTotal });
  } catch (err) {
    return errResponse('讀取月累計失敗：', err);
  }
}

// ★ 新增（2026-07-19）：依「取餐時間」把當天訂單分類到早餐/午餐/晚餐/宵夜，
// 時段分界跟套餐系統的時段判斷一致（早餐00:00-11:29、午餐11:30-17:29、
// 晚餐17:30-21:29、宵夜21:30-23:59），供「更簡版日報表」自動帶入營收使用，
// 不用店家自己重複輸入系統本來就有的訂單金額。
function classifyPickupTimePeriod(pickupTime) {
  const startTime = String(pickupTime || '').split('-')[0].trim();
  const parts = startTime.split(':');
  if (parts.length < 2) return null;
  const h = Number(parts[0]);
  const m = Number(parts[1]);
  if (isNaN(h) || isNaN(m)) return null;
  const minutes = h * 60 + m;
  if (minutes < 11 * 60 + 30) return 'breakfast';
  if (minutes < 17 * 60 + 30) return 'lunch';
  if (minutes < 21 * 60 + 30) return 'dinner';
  return 'supper';
}

// ────────────────────────────────────────────
// ★ 新增（2026-07-19）：採購單明細——詳細登錄模式，跟「更簡版」的自由列表
// 並存，兩種都可以用。這套需要「品項資料庫」支援代號自動帶入品名/單價，
// 跟自由列表最大的差異在這裡。
// ────────────────────────────────────────────

function getPurchaseItemsList() {
  const ss = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_PURCHASE_ITEMS);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  return data
    .filter(row => row[0] !== '')
    .map(row => ({
      code: row[0],
      name: row[1] || '',
      category: row[2] || '',
      unit: row[3] || '',
      refPrice: Number(row[4]) || 0
    }));
}

function handleGetPurchaseItems(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    return jsonResponse({ status: 'success', items: getPurchaseItemsList() });
  } catch (err) {
    return errResponse('讀取品項資料庫失敗：', err);
  }
}

// ★ F2快速新增品項：只需要代號/品名/類別，單位/參考單價可留空之後補
function handleAddPurchaseItem(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const code = (body.code || '').trim();
    const name = (body.name || '').trim();
    if (!code || !name) return jsonResponse({ status: 'error', message: '代號跟品名為必填' });

    const sheet = getOrCreateSheet(SHEET_NAME_PURCHASE_ITEMS, ['代號', '品名', '類別', '單位', '參考單價'], REPORT_SPREADSHEET_ID);
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const codes = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < codes.length; i++) {
        if (String(codes[i][0]) === code) {
          return jsonResponse({ status: 'error', message: `代號「${code}」已經存在，請用其他代號` });
        }
      }
    }

    sheet.appendRow([
      code,
      sanitizeForSheet(name),
      sanitizeForSheet(body.category || ''),
      sanitizeForSheet(body.unit || ''),
      Number(body.refPrice) || 0
    ]);

    return jsonResponse({ status: 'success', message: '品項已新增' });
  } catch (err) {
    return errResponse('新增品項失敗：', err);
  }
}

function handleUpdatePurchaseItems(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const items = Array.isArray(body.items) ? body.items : [];
    const sheet = getOrCreateSheet(SHEET_NAME_PURCHASE_ITEMS, ['代號', '品名', '類別', '單位', '參考單價'], REPORT_SPREADSHEET_ID);
    sheet.clear();
    sheet.getRange(1, 1, 1, 5).setValues([['代號', '品名', '類別', '單位', '參考單價']]);
    sheet.getRange(1, 1, 1, 5)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    if (items.length > 0) {
      const rows = items.map(it => [
        it.code || '', sanitizeForSheet(it.name || ''), sanitizeForSheet(it.category || ''),
        sanitizeForSheet(it.unit || ''), Number(it.refPrice) || 0
      ]);
      sheet.getRange(2, 1, rows.length, 5).setValues(rows);
    }

    return jsonResponse({ status: 'success', message: '品項資料庫已更新' });
  } catch (err) {
    return errResponse('更新品項資料庫失敗：', err);
  }
}

// ★ 送出採購單：把整張單（進貨日期/供應商/付款方式/備註+多個品項列）存進
// 「採購單記錄」表留存查核，同時「自動同步」把每個品項列，各自轉成一筆
// category='食材成本'的entry，附加進「每日營運記錄」對應日期的entries裡——
// 這樣送出採購單之後，成本報表就會自動反映這筆採購，不用再手動去更簡版
// 那邊重複登記一次同樣的金額。
// ★ 2026-07-20：採購單記錄、每日營運記錄現在都在REPORT_SPREADSHEET_ID，
// 這支函式全程只需要開一份試算表(reportSs)即可，不用再碰A表。
function handleSubmitPurchaseOrder(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const purchaseDate = (body.purchaseDate || '').trim();
    const supplier = (body.supplier || '').trim();
    const paymentMethod = (body.paymentMethod || '').trim();
    const note = (body.note || '').trim();
    const lines = Array.isArray(body.lines) ? body.lines : [];

    if (!purchaseDate) return jsonResponse({ status: 'error', message: '缺少進貨日期' });
    if (lines.length === 0) return jsonResponse({ status: 'error', message: '至少要有一個品項' });

    // 1. 存進採購單記錄表（留存查核用，一列一個品項）
    const ordersSheet = getOrCreateSheet(SHEET_NAME_PURCHASE_ORDERS, [
      '進貨日期', '供應商', '付款方式', '整單備註', '品項代號', '品項名稱', '類別', '數量', '單位', '單價', '小計', '單據號碼'
    ], REPORT_SPREADSHEET_ID);
    const orderRows = lines.map(line => [
      purchaseDate,
      sanitizeForSheet(supplier),
      sanitizeForSheet(paymentMethod),
      sanitizeForSheet(note),
      line.code || '',
      sanitizeForSheet(line.name || ''),
      sanitizeForSheet(line.category || ''),
      Number(line.qty) || 0,
      sanitizeForSheet(line.unit || ''),
      Number(line.price) || 0,
      Number(line.subtotal) || 0,
      sanitizeForSheet(line.receiptNo || '')
    ]);
    ordersSheet.getRange(ordersSheet.getLastRow() + 1, 1, orderRows.length, 12).setValues(orderRows);

    // 2. 自動同步進每日營運記錄——每個品項列轉成一筆食材成本entry
    const reportSs = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    const existingRecord = getDailyRecordForDate(reportSs, purchaseDate) || { entries: [], revenue: {}, note: '' };
    const newEntries = lines.map(line => ({
      category: '食材成本',
      supplier: supplier,
      item: line.name || '',
      amount: Number(line.subtotal) || 0
    }));
    const mergedEntries = (existingRecord.entries || []).concat(newEntries);

    const dailySheet = getOrCreateSheet(SHEET_NAME_DAILY_RECORDS, ['日期', '明細記錄(JSON)', '實收入(JSON)', '備註'], REPORT_SPREADSHEET_ID);
    const lastRow = dailySheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= 2) {
      const dates = dailySheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        const rowDate = dates[i][0] instanceof Date ? Utilities.formatDate(dates[i][0], 'GMT+8', 'yyyy-MM-dd') : String(dates[i][0]);
        if (rowDate === purchaseDate) { targetRow = i + 2; break; }
      }
    }
    const rowValues = [purchaseDate, JSON.stringify(mergedEntries), JSON.stringify(existingRecord.revenue || {}), sanitizeForSheet(existingRecord.note || '')];
    if (targetRow > 0) {
      dailySheet.getRange(targetRow, 1, 1, 4).setValues([rowValues]);
    } else {
      dailySheet.appendRow(rowValues);
    }

    return jsonResponse({ status: 'success', message: `採購單已送出，共 ${lines.length} 個品項，已自動同步進 ${purchaseDate} 的成本記錄` });
  } catch (err) {
    return errResponse('送出採購單失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 新增（2026-07-19）：接收本地系統（invoice_server_lpt.py的
// /daily_report_upload）推播過來的每日報表。本地系統早就做好了推播的
// 一半（_post_to_gas函式），只是Code.gs這邊一直沒有接收端，這裡補上。
//
// 驗證用的是APP_SHARED_SECRET（機器對機器信任，本地系統本來就要跟
// invoice.ini裡設定的密碼一致），不是verifyAdminPassword()那組給人登入
// 網頁用的密碼——這是自動化推播，不該要求人工登入密碼。
// ★ 2026-07-20：接收到的資料改寫進REPORT_SPREADSHEET_ID的「本地系統日報」
// 工作表，跟其他營業報表資料放在一起。
// ────────────────────────────────────────────
function handleReceiveLocalDailyReport(body) {
  const token = (body.token || '').trim();
  const secret = getAppSharedSecret();
  if (!secret || token !== secret) {
    return jsonResponse({ status: 'error', message: '驗證失敗，token不符或尚未設定APP_SHARED_SECRET' });
  }

  try {
    const dateStr = (body.date || '').trim();
    if (!dateStr) return jsonResponse({ status: 'error', message: '缺少 date 參數' });

    const sheet = getOrCreateSheet(SHEET_NAME_LOCAL_DAILY_REPORT, [
      '日期', '店名', '含稅營收', '未稅營收', '稅金', '訂單數', '有效訂單數', '取消訂單數',
      '平均客單價', '食材成本', '固定成本', '非固定成本', '總成本', '毛利', '淨利', '成本模式',
      '品項排行JSON', '推播時間', '收到時間'
    ], REPORT_SPREADSHEET_ID);

    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= 2) {
      const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        const rowDate = dates[i][0] instanceof Date ? Utilities.formatDate(dates[i][0], 'GMT+8', 'yyyy-MM-dd') : String(dates[i][0]);
        if (rowDate === dateStr) { targetRow = i + 2; break; }
      }
    }

    const rowValues = [
      dateStr,
      sanitizeForSheet(body.store || ''),
      Number(body.withTax) || 0,
      Number(body.pretax) || 0,
      Number(body.tax) || 0,
      Number(body.orderCount) || 0,
      Number(body.validCount) || 0,
      Number(body.cancelCount) || 0,
      Number(body.avgOrderValue) || 0,
      Number(body.foodCost) || 0,
      Number(body.fixedCost) || 0,
      Number(body.nonFoodCost) || 0,
      Number(body.totalCost) || 0,
      Number(body.grossProfit) || 0,
      Number(body.netProfit) || 0,
      sanitizeForSheet(body.costMode || ''),
      JSON.stringify(body.items || []),
      sanitizeForSheet(body.generatedAt || ''),
      Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss')
    ];

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return jsonResponse({ status: 'success', message: `本地系統日報（${dateStr}）已接收` });
  } catch (err) {
    return errResponse('接收本地日報失敗：', err);
  }
}

function handleGetRevenueByPeriod(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const dateStr = (e.parameter.date || '').trim();
    if (!dateStr) return jsonResponse({ status: 'error', message: '缺少 date 參數' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    const totals = { breakfast: 0, lunch: 0, dinner: 0, supper: 0 };

    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
        data.forEach(row => {
          if (row[0] === '') return;
          if (row[1] === '已取消') return;
          const pickupDate = row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+8', 'yyyy-MM-dd') : String(row[8]);
          if (pickupDate !== dateStr) return;
          const pickupTime = row[9] instanceof Date ? Utilities.formatDate(row[9], 'GMT+8', 'HH:mm') : String(row[9]);
          const period = classifyPickupTimePeriod(pickupTime);
          if (!period) return;
          totals[period] += Number(row[7]) || 0;
        });
      }
    }

    Object.keys(totals).forEach(k => { totals[k] = Math.round(totals[k]); });
    return jsonResponse({ status: 'success', totals: totals });
  } catch (err) {
    return errResponse('讀取營收分時段資料失敗：', err);
  }
}

function handleGetDailyRecord(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const dateStr = (e.parameter.date || '').trim();
    if (!dateStr) return jsonResponse({ status: 'error', message: '缺少 date 參數' });

    const ss = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    const record = getDailyRecordForDate(ss, dateStr);
    if (record) {
      return jsonResponse({ status: 'success', found: true, record: record });
    }
    return jsonResponse({ status: 'success', found: false, record: { entries: [], revenue: {}, note: '' } });
  } catch (err) {
    return errResponse('讀取每日記錄失敗：', err);
  }
}

function handleUpdateDailyRecord(body) {
  const password = body.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const dateStr = (body.date || '').trim();
    if (!dateStr) return jsonResponse({ status: 'error', message: '缺少 date 參數' });

    const rawEntries = Array.isArray(body.entries) ? body.entries : [];
    const validCategories = ['固定', '非固定', '食材成本'];
    const entries = rawEntries
      .filter(en => en && validCategories.indexOf(en.category) !== -1)
      .map(en => ({
        category: en.category,
        supplier: sanitizeForSheet(en.supplier || ''),
        item: sanitizeForSheet(en.item || ''),
        amount: Number(en.amount) || 0,
        stockout: !!en.stockout // ★ 2026-07-22 新增：這筆食材是否標記「缺貨」
      }));

    // ★ 2026-07-22 新增：進貨數量歸零/標記缺貨，自動下架對應菜單品項。
    // 用「品名完全相符」去比對菜單品項名稱（例如進貨登記「雞胸肉」，菜單
    // 裡剛好也有一個品項叫「雞胸肉」，就會被自動下架）——這是最簡單、
    // 不用額外維護「食材對應哪些菜單品項」這種複雜設定表的做法，但也因此
    // 只能抓「品名完全一樣」的情況，如果菜單品項名稱寫法不完全一樣
    // （例如「炸雞胸」跟進貨的「雞胸肉」），不會自動連動，還是要手動下架。
    let autoDelistedItems = [];
    if (isFeatureEnabled('autoDelistOnStockout')) {
      const stockoutItemNames = entries
        .filter(en => en.category === '食材成本' && (en.stockout || en.amount === 0))
        .map(en => en.item)
        .filter(Boolean);
      if (stockoutItemNames.length > 0) {
        autoDelistedItems = autoDelistMenuItemsByName(stockoutItemNames, body.operator || '');
      }
    }

    const sheet = getOrCreateSheet(SHEET_NAME_DAILY_RECORDS, ['日期', '明細記錄(JSON)', '實收入(JSON)', '備註'], REPORT_SPREADSHEET_ID);
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= 2) {
      const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        const rowDate = dates[i][0] instanceof Date ? Utilities.formatDate(dates[i][0], 'GMT+8', 'yyyy-MM-dd') : String(dates[i][0]);
        if (rowDate === dateStr) { targetRow = i + 2; break; }
      }
    }

    const rowValues = [
      dateStr,
      JSON.stringify(entries),
      JSON.stringify(body.revenue || {}),
      sanitizeForSheet(body.note || '')
    ];

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, 4).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    logOperatorAction('儲存每日進貨記錄', dateStr, body.operator);
    const delistNote = autoDelistedItems.length > 0
      ? `（已自動下架：${autoDelistedItems.join('、')}）`
      : '';
    return jsonResponse({ status: 'success', message: '本日記錄已儲存' + delistNote, autoDelistedItems });
  } catch (err) {
    return errResponse('儲存每日記錄失敗：', err);
  }
}

// ★ 2026-07-22 新增：依品名比對，把菜單裡完全同名的品項自動改成「已下架」。
// 回傳實際被下架的品項名稱清單（給前端顯示提示用）。
// 欄位位置對照 handleGetMenu 確認過：A=分類, B=品名, C=價格, D=圖示,
// E=狀態, F=照片, G=選項, H=成本, I=安全庫存, J=可升級套餐, K=升級價格
function autoDelistMenuItemsByName(itemNames, operator) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_MENU);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const nameSet = new Set(itemNames.map(n => n.trim()));
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const delisted = [];

    for (let i = 0; i < data.length; i++) {
      const name = String(data[i][1] || '').trim(); // B欄=品名
      if (nameSet.has(name)) {
        const rowNumber = i + 2;
        const currentStatus = data[i][4]; // E欄=狀態
        if (currentStatus !== '已下架') {
          sheet.getRange(rowNumber, 5).setValue('已下架'); // E欄
          delisted.push(name);
        }
      }
    }
    if (delisted.length > 0) {
      logOperatorAction('進貨缺貨自動下架菜單品項', delisted.join('、'), operator || '系統自動');
    }
    return delisted;
  } catch (err) {
    console.error('自動下架菜單品項失敗：', err.message);
    return [];
  }
}

function handleGetFixedCosts(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    return jsonResponse({ status: 'success', items: getFixedCostsList() });
  } catch (err) {
    return errResponse('讀取固定成本失敗：', err);
  }
}

function handleUpdateFixedCosts(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const raw = e.parameter.items;
    if (!raw) return jsonResponse({ status: 'error', message: '缺少 items 參數' });
    const items = JSON.parse(raw);
    if (!Array.isArray(items)) return jsonResponse({ status: 'error', message: 'items 格式錯誤' });

    const sheet = getOrCreateSheet(SHEET_NAME_FIXED_COSTS, ['項目名稱', '金額'], REPORT_SPREADSHEET_ID);
    sheet.clear();
    sheet.getRange(1, 1, 1, 2).setValues([['項目名稱', '金額']]);
    sheet.getRange(1, 1, 1, 2)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    if (items.length > 0) {
      const rows = items.map(it => [sanitizeForSheet(it.name || ''), Number(it.amount) || 0]);
      sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }

    return jsonResponse({ status: 'success', message: `固定成本已更新，共 ${items.length} 項` });
  } catch (err) {
    return errResponse('更新固定成本失敗：', err);
  }
}

function getDaysInMonth(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  return new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
}

// ★ 算單一天的：營收、食材成本、賣出品項明細。掃「訂單資料」表，
// 只挑取餐日期符合、且未取消的訂單。
function computeDayCostData(dateStr, ss, menuCostMap) {
  const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  let revenue = 0;
  let foodCost = 0;
  const itemQtyMap = {};

  if (!sheet) return { revenue, foodCost, itemQtyMap };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { revenue, foodCost, itemQtyMap };

  const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
  data.forEach(row => {
    if (row[0] === '') return;
    const status = row[1];
    if (status === '已取消') return;
    const pickupDate = row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+8', 'yyyy-MM-dd') : String(row[8]);
    if (pickupDate !== dateStr) return;

    revenue += Number(row[7]) || 0;

    let cart = [];
    try { cart = JSON.parse(row[16] || '[]'); } catch (parseErr) { return; }
    cart.forEach(ci => {
      if (!ci || !ci.name) return;
      const qty = Number(ci.quantity) || 0;
      itemQtyMap[ci.name] = (itemQtyMap[ci.name] || 0) + qty;
      const unitCost = menuCostMap[ci.name] || 0;
      foodCost += unitCost * qty;
    });
  });

  return { revenue: Math.round(revenue), foodCost: Math.round(foodCost), itemQtyMap };
}

// ★ 成本報表主要API：給日期區間，回傳每天的營收/成本/利潤明細（供趨勢圖、
// 今日/昨日對比使用）＋整個區間的品項銷售排行（供熱銷/滯銷榜使用）
// ★ 2026-07-20：這支函式同時需要A表（訂單/菜單）跟B表（固定成本/每日
// 手動登錄記錄）的資料，改成同時開兩份試算表：ss指向A表，reportSs指向B表。
// ────────────────────────────────────────────
// ★ 2026-07-22 新增：取得最近進貨用過的品名清單，讓進貨登錄畫面能自動
// 建議「之前打過的品名」，不用每次都重新手打「雞胸肉」「洋蔥」這些
// 常買的東西，手機打字比較省力。掃描「每日營運記錄」整張表（通常一天
// 一列，資料量不大），只取最近RECENT_DAYS天內的紀錄，避免掃到太久
// 以前、可能已經不再進貨的品項。
// ────────────────────────────────────────────
function handleGetRecentPurchaseItemNames(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const RECENT_DAYS = 60;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - RECENT_DAYS);

    const reportSs = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    const sheet = reportSs.getSheetByName(SHEET_NAME_DAILY_RECORDS);
    if (!sheet) return jsonResponse({ status: 'success', itemNames: [], itemSupplierMap: {} });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', itemNames: [], itemSupplierMap: {} });

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // 日期, 明細記錄(JSON)
    const nameCount = {};
    // ★ 2026-07-22 新增：記錄「這個品名，最近一次是跟哪家廠商叫的」，
    // 讓前端選品名時可以自動帶出廠商，不用每次都重選。用「最近一次」
    // 而不是「最常見」，因為同一樣食材偶爾會換廠商叫貨（例如原廠商
    // 缺貨），最近一次的紀錄比較符合「現在通常跟誰叫」的實際狀況。
    const itemSupplierMap = {};
    const itemLastDate = {};

    data.forEach(row => {
      const dateVal = row[0];
      const dateObj = dateVal instanceof Date ? dateVal : new Date(dateVal);
      if (isNaN(dateObj) || dateObj < cutoff) return;

      try {
        const entries = JSON.parse(row[1] || '[]');
        entries.forEach(en => {
          if (en.category !== '食材成本') return;
          const name = (en.item || '').trim();
          if (!name) return;
          nameCount[name] = (nameCount[name] || 0) + 1;

          const supplier = (en.supplier || '').trim();
          if (supplier && (!itemLastDate[name] || dateObj > itemLastDate[name])) {
            itemLastDate[name] = dateObj;
            itemSupplierMap[name] = supplier;
          }
        });
      } catch (parseErr) { /* 這天的JSON格式有問題，跳過，不影響其他天 */ }
    });

    // 依出現次數排序，常買的排前面，最多回傳50個，避免建議清單太長
    const itemNames = Object.keys(nameCount)
      .sort((a, b) => nameCount[b] - nameCount[a])
      .slice(0, 50);

    return jsonResponse({ status: 'success', itemNames: itemNames, itemSupplierMap: itemSupplierMap });
  } catch (err) {
    return errResponse('讀取進貨品名建議失敗：', err);
  }
}

function handleGetCostReport(e) {
  const password = e.parameter.password || '';
  const authCheck = verifyAdminPassword(password);
  if (!authCheck.ok) return jsonResponse({ status: 'error', message: authCheck.message });

  try {
    const dateFrom = (e.parameter.from || '').trim();
    const dateTo = (e.parameter.to || '').trim();
    if (!dateFrom || !dateTo) {
      return jsonResponse({ status: 'error', message: '缺少 from / to 日期參數' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const reportSs = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    const menuItems = handleGetMenuRaw(ss);
    const menuCostMap = {};
    menuItems.forEach(it => { menuCostMap[it.name] = it.cost; });

    const fixedCosts = getFixedCostsList();
    const monthlyFixedTotal = fixedCosts.reduce((s, f) => s + f.amount, 0);

    const daily = [];
    const totalItemQty = {};

    let d = new Date(dateFrom + 'T00:00:00');
    const end = new Date(dateTo + 'T00:00:00');
    // ★ 區間安全上限：最多算93天（約3個月），避免網址誤帶超大區間時
    // 迴圈跑太久超過GAS執行時間限制
    let safetyCounter = 0;
    while (d <= end && safetyCounter < 93) {
      const dateStr = Utilities.formatDate(d, 'GMT+8', 'yyyy-MM-dd');
      const dayData = computeDayCostData(dateStr, ss, menuCostMap);
      const daysInMo = getDaysInMonth(dateStr);
      const fixedToday = daysInMo > 0 ? Math.round(monthlyFixedTotal / daysInMo) : 0;

      const manualRecord = getDailyRecordForDate(reportSs, dateStr);
      let revenue = dayData.revenue;
      let foodCost = dayData.foodCost;
      let fixedCostForDay = fixedToday;

      if (manualRecord) {
        const manualRevenue = Object.values(manualRecord.revenue || {})
          .reduce((s, r) => s + (Number(r.card) || 0) + (Number(r.cash) || 0) + (Number(r.other) || 0), 0);
        const entries = manualRecord.entries || [];
        const manualFoodCost = entries.filter(en => en.category === '食材成本').reduce((s, en) => s + (Number(en.amount) || 0), 0);
        const manualOtherCost = entries.filter(en => en.category === '固定' || en.category === '非固定').reduce((s, en) => s + (Number(en.amount) || 0), 0);
        revenue = manualRevenue;
        foodCost = manualFoodCost;
        fixedCostForDay = manualOtherCost;
      }

      const profit = revenue - foodCost - fixedCostForDay;

      daily.push({
        date: dateStr,
        revenue: revenue,
        foodCost: foodCost,
        fixedCost: fixedCostForDay,
        profit: profit,
        hasManualRecord: !!manualRecord
      });

      Object.keys(dayData.itemQtyMap).forEach(name => {
        totalItemQty[name] = (totalItemQty[name] || 0) + dayData.itemQtyMap[name];
      });

      d.setDate(d.getDate() + 1);
      safetyCounter++;
    }

    // 品項排行：依銷售數量由高到低排序，同時附上該品項的價格供前端算營收貢獻
    const priceMap = {};
    menuItems.forEach(it => { priceMap[it.name] = it.price; });
    const itemRanking = Object.keys(totalItemQty)
      .map(name => ({ name: name, qty: totalItemQty[name], revenue: Math.round((priceMap[name] || 0) * totalItemQty[name]) }))
      .sort((a, b) => b.qty - a.qty);

    return jsonResponse({
      status: 'success',
      daily: daily,
      itemRanking: itemRanking,
      monthlyFixedTotal: monthlyFixedTotal
    });
  } catch (err) {
    return errResponse('讀取成本報表失敗：', err);
  }
}

// ────────────────────────────────────────────
// ★ 新增（2026-07-19）：新客戶上線前的健檢清單。手動在Apps Script編輯器
// 選這支函式執行，看下面「執行紀錄」的完整報告，確認每一項該設定的東西
// 都設定好了，不用重演今天一路debug過的「以為設定好結果沒有」的狀況。
// 這支函式不會動任何資料，純粹讀取現況、印出報告，執行是安全的。
// ★ 2026-07-20：健檢範圍擴大到同時檢查A表（點餐主表）跟B表（營業日報表）
// 兩份試算表是否都能正常開啟、必要工作表是否都存在。
// ────────────────────────────────────────────
function runHealthCheck() {
  const lines = [];
  let passCount = 0, failCount = 0, warnCount = 0;

  function check(label, ok, detail) {
    if (ok) { passCount++; lines.push(`✅ ${label}`); }
    else { failCount++; lines.push(`❌ ${label}｜${detail || ''}`); }
  }
  function warn(label, detail) {
    warnCount++; lines.push(`⚠️ ${label}｜${detail || ''}`);
  }

  lines.push('══════ 新客戶上線健檢報告 ══════');
  lines.push('');

  // 1. 兩份試算表本身能不能正常打開
  let ss = null;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    check('點餐主表（A表）可以正常開啟', true);
  } catch (err) {
    check('點餐主表（A表）可以正常開啟', false, 'SPREADSHEET_ID設定錯誤或沒有存取權限：' + err.message);
  }

  let reportSs = null;
  try {
    reportSs = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    check('營業日報表（B表）可以正常開啟', true);
  } catch (err) {
    check('營業日報表（B表）可以正常開啟', false, 'REPORT_SPREADSHEET_ID設定錯誤或沒有存取權限：' + err.message);
  }

  // 2. 三個機密設定（PropertiesService）
  const secret = getAppSharedSecret();
  check('APP_SHARED_SECRET 已設定', !!secret, '請執行 setAppSharedSecret() 完成設定');

  const adminPw = getAdminPasswordValue();
  check('後台密碼（ADMIN_PASSWORD）已設定', !!adminPw, '請執行 setAdminPassword() 完成設定');

  const lineToken = getLineChannelAccessToken();
  const lineTokenOk = !!lineToken && lineToken.indexOf('請貼上') === -1;
  if (lineTokenOk) { check('LINE Channel Access Token 已設定', true); }
  else { warn('LINE Channel Access Token 尚未設定', '訂單完成通知、店家新訂單通知都不會運作，請執行 setLineChannelAccessToken()（如果不需要推播通知可以先略過）'); }

  const shopNotifyIds = getShopNotifyLineUserIds();
  if (shopNotifyIds.length > 0) { check(`店家新訂單通知已設定（共${shopNotifyIds.length}組LINE UserID）`, true); }
  else { warn('店家新訂單通知尚未設定', '新訂單不會主動推播給店家，請執行 setShopNotifyLineUserIds()（非必要功能，可以先略過）'); }

  // 3. LINE_CHANNEL_ID有沒有像是還沒改過的樣子
  check('LINE_CHANNEL_ID 看起來已填寫', !!LINE_CHANNEL_ID && LINE_CHANNEL_ID.length > 5, '請確認頂部LINE_CHANNEL_ID常數已經改成這個客戶自己的Channel ID');

  // 4. 逐一檢查A表必要工作表是否存在
  if (ss) {
    const requiredSheets = [
      SHEET_NAME_ORDERS, SHEET_NAME_USERS, SHEET_NAME_MENU, SHEET_NAME_ADDONS,
      SHEET_NAME_OPTION_SPECS, SHEET_NAME_PROMO_RULES, SHEET_NAME_ANNOUNCEMENTS,
      SHEET_NAME_TOPUP_TIERS, SHEET_NAME_TOPUP_LEDGER, SHEET_NAME_PAYMENT_REPORTS,
      SHEET_NAME_DISCLAIMER
    ];
    requiredSheets.forEach(name => {
      check(`A表工作表「${name}」存在`, !!ss.getSheetByName(name), '請執行 setupAllSheets() 建立A表所有必要工作表');
    });

    // 5. 菜單裡有沒有至少一項供應中的品項（避免客人打開頁面看到空菜單）
    try {
      const menuItems = handleGetMenuRaw(ss);
      check('菜單至少有1個品項', menuItems.length > 0, '請至少新增1個菜單品項，不然點餐頁面會是空的');
    } catch (err) {
      warn('無法檢查菜單內容', err.message);
    }
  }

  // 5.5 逐一檢查B表必要工作表是否存在
  if (reportSs) {
    const requiredReportSheets = [
      SHEET_NAME_FIXED_COSTS, SHEET_NAME_COST_CATEGORIES, SHEET_NAME_DAILY_RECORDS, SHEET_NAME_PURCHASE_ITEMS
    ];
    requiredReportSheets.forEach(name => {
      check(`B表工作表「${name}」存在`, !!reportSs.getSheetByName(name), '請執行 setupReportSheets() 建立B表所有必要工作表');
    });
  }

  // 6. 每日自動備份的觸發器有沒有設定
  try {
    const hasBackupTrigger = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'backupSpreadsheet');
    if (hasBackupTrigger) { check('每日自動備份已設定', true); }
    else { warn('每日自動備份尚未設定', '請執行 setupBackupTrigger()，避免資料萬一出問題時求助無門（非必要但強烈建議）'); }
  } catch (err) {
    warn('無法檢查備份觸發器', err.message);
  }

  lines.push('');
  lines.push('══════════════════════════');
  lines.push(`總結：${passCount} 項通過｜${failCount} 項失敗（必須修正才能上線）｜${warnCount} 項警告（建議但非必要）`);
  if (failCount > 0) {
    lines.push('⚠️ 有「❌失敗」項目，代表系統還沒準備好正式上線，請照上面的指示逐一修正。');
  } else if (warnCount > 0) {
    lines.push('✅ 沒有必須修正的項目，可以上線；上面「⚠️警告」的部分是加分項，有空再補。');
  } else {
    lines.push('✅ 全部檢查通過，可以正式上線！');
  }

  const report = lines.join('\n');
  Logger.log(report);
  return report;
}

function generateOrderNumber() {
  const now = new Date();
  const datePart = Utilities.formatDate(now, 'GMT+8', 'yyyyMMdd');
  // ★ 安全修正：原本只用4碼隨機亂數（1000~9999，只有9000種可能），生意量大
  // 起來（同一天100~150張訂單左右）用生日悖論算，撞號機率就超過一半。
  // 加上毫秒級時間戳的末3碼一起組合，等於大幅擴增可能的組合數，
  // 實務上兩筆訂單同一毫秒又抽到同樣亂數的機率趨近於零。
  const msPart = String(now.getTime()).slice(-3);
  const rand = Math.floor(Math.random() * 9000 + 1000);
  return `ORD-${datePart}-${rand}${msPart}`;
}

// ★ 2026-07-20：新增第三個參數 spreadsheetId（可省略，預設仍是A表的
// SPREADSHEET_ID），讓B表（營業日報表）相關函式也能共用同一支
// getOrCreateSheet()，不用另外重寫一份幾乎一樣的邏輯。
function getOrCreateSheet(name, headers, spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId || SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

const SCRIPT_VERSION = 'v85-2026-07-23-menu-extra-fields';
const EXPECTED_FRONTEND_VERSION = 'v97-2026-07-23-menu-extra-fields';

// ★ 安全修正：統一的錯誤回應函式，取代原本27處「直接把err.message丟給
// 客人看」的寫法。真正的錯誤細節（可能包含試算表內部欄位名稱、變數名稱、
// GAS內部錯誤訊息這些不該讓外部看到的東西）只寫進Logger.log()，
// 只有你自己在Apps Script執行紀錄裡看得到；回傳給客人的訊息維持原本的
// 中文提示前綴，但拿掉技術細節，改成通用的「請稍後再試」。
function errResponse(prefixRaw, err) {
  const prefix = String(prefixRaw || '').replace(/[：:]+$/, ''); // 去掉尾端可能帶的冒號，避免跟後面組合時重複
  Logger.log(`[錯誤] ${prefix}：` + (err && err.message ? err.message : String(err)));
  return jsonResponse({ status: 'error', message: `${prefix}，請稍後再試，如持續發生請聯絡店家` });
}

function jsonResponse(obj) {
  obj.scriptVersion = SCRIPT_VERSION;
  obj.expectedFrontendVersion = EXPECTED_FRONTEND_VERSION;
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ★ 2026-07-20：這支函式只負責初始化A表（點餐主表）的工作表，原本這裡
// 也一併建立的固定成本設定／成本類別設定／每日營運記錄／採購品項資料庫
// 四張表，已經移到下面新增的 setupReportSheets()，改為建立在B表
// （REPORT_SPREADSHEET_ID）。兩支函式互不影響，各自獨立執行即可。
function setupAllSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let ordersSheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet(SHEET_NAME_ORDERS);
  }
  const orderHeaders = [
    '訂單編號', '訂單狀態', '創建時間',
    '顧客姓名', '手機號碼', '顧客LINE_UserID',
    '訂單內容', '總金額',
    '取餐日期', '取餐時間',
    '付款方式', '外送備註', '訂單備註',
    '匯款截圖確認', '管理員備註', '最後更新時間',
    '購物車明細JSON', '匯款截圖網址'
  ];
  ordersSheet.getRange(1, 1, 1, orderHeaders.length).setValues([orderHeaders]);
  ordersSheet.getRange(1, 1, 1, orderHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  ordersSheet.setFrozenRows(1);

  ordersSheet.getRange('E2:E100000').setNumberFormat('@');

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未匯款', '已匯款', '準備中', '已完成', '已取消'], true)
    .setAllowInvalid(true)
    .build();
  ordersSheet.getRange('B2:B100000').setDataValidation(statusRule);

  const confirmRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未確認', '已確認'], true)
    .setAllowInvalid(true)
    .build();
  ordersSheet.getRange('N2:N100000').setDataValidation(confirmRule);

  const orderWidths = [110,90,130, 80,90,150, 220,70, 90,80, 60,190,150, 90,150,130, 260,220];
  orderWidths.forEach((w, i) => ordersSheet.setColumnWidth(i + 1, w));

  let promoSheet = ss.getSheetByName(SHEET_NAME_PROMO_RULES);
  if (!promoSheet) {
    promoSheet = ss.insertSheet(SHEET_NAME_PROMO_RULES);
  }
  promoSheet.clear();
  const promoHeaders = ['規則ID', '名稱', '類型', '啟用', '參數JSON', '指定品項'];
  const promoItems = [
    ['rule-demo-1', '買5送1', 'BUYXGETY', '停用', '{"buyX":5,"getY":1}', ''],
    ['rule-demo-2', '第二件半價', 'NTH_DISCOUNT', '停用', '{"nth":2,"rate":0.5}', ''],
    ['rule-demo-3', '滿300打9折', 'PERCENT_OFF', '停用', '{"threshold":300,"rate":0.9}', ''],
    ['rule-demo-4', '滿300送脆皮雞塊', 'THRESHOLD_GIFT', '停用', '{"threshold":300,"giftItemName":"⓪ 脆皮雞塊","giftQty":1,"giftValue":75}', '']
  ];
  promoSheet.getRange(1, 1, 1, promoHeaders.length).setValues([promoHeaders]);
  promoSheet.getRange(1, 1, 1, promoHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  promoSheet.getRange(2, 1, promoItems.length, promoHeaders.length).setValues(promoItems);
  promoSheet.setFrozenRows(1);
  const promoWidths = [130, 160, 140, 70, 280, 200];
  promoWidths.forEach((w, i) => promoSheet.setColumnWidth(i + 1, w));

  let announceSheet = ss.getSheetByName(SHEET_NAME_ANNOUNCEMENTS);
  if (!announceSheet) {
    announceSheet = ss.insertSheet(SHEET_NAME_ANNOUNCEMENTS);
  }
  announceSheet.clear();
  const announceHeaders = ['公告ID', '內容', '啟用', '顯示方式'];
  const announceItems = [
    ['ann-demo-1', '🎉 本週優惠：第二件半價，數量有限！', '停用', '跑馬燈'],
    ['ann-demo-2', '🍗 招牌脆皮雞塊限時加碼中，快來嚐鮮！', '停用', '小視窗']
  ];
  announceSheet.getRange(1, 1, 1, announceHeaders.length).setValues([announceHeaders]);
  announceSheet.getRange(1, 1, 1, announceHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  announceSheet.getRange(2, 1, announceItems.length, announceHeaders.length).setValues(announceItems);
  announceSheet.setFrozenRows(1);
  const announceWidths = [130, 400, 70, 100];
  announceWidths.forEach((w, i) => announceSheet.setColumnWidth(i + 1, w));

  let tierSheet = ss.getSheetByName(SHEET_NAME_TOPUP_TIERS);
  if (!tierSheet) {
    tierSheet = ss.insertSheet(SHEET_NAME_TOPUP_TIERS);
  }
  tierSheet.clear();
  const tierHeaders = ['方案ID', '儲值金額', '到帳金額', '啟用'];
  const tierItems = [
    ['tier-demo-1', 1000, 1300, '停用'],
    ['tier-demo-2', 1500, 2000, '停用']
  ];
  tierSheet.getRange(1, 1, 1, tierHeaders.length).setValues([tierHeaders]);
  tierSheet.getRange(1, 1, 1, tierHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  tierSheet.getRange(2, 1, tierItems.length, tierHeaders.length).setValues(tierItems);
  tierSheet.setFrozenRows(1);
  const tierWidths = [140, 100, 100, 70];
  tierWidths.forEach((w, i) => tierSheet.setColumnWidth(i + 1, w));

  let ledgerSheet = ss.getSheetByName(SHEET_NAME_TOPUP_LEDGER);
  if (!ledgerSheet) {
    ledgerSheet = ss.insertSheet(SHEET_NAME_TOPUP_LEDGER);
  }
  const ledgerHeaders = ['記錄ID', '時間', 'LINE_UserID', 'LINE_DisplayName', '類型', '金額', '狀態', '對應訂單編號', '顧客姓名', '聯絡電話', '管理員備註'];
  ledgerSheet.getRange(1, 1, 1, ledgerHeaders.length).setValues([ledgerHeaders]);
  ledgerSheet.getRange(1, 1, 1, ledgerHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  ledgerSheet.setFrozenRows(1);
  const ledgerWidths = [160, 130, 250, 150, 90, 80, 80, 150, 90, 110, 260];
  ledgerWidths.forEach((w, i) => ledgerSheet.setColumnWidth(i + 1, w));

  let usersSheet = ss.getSheetByName(SHEET_NAME_USERS);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(SHEET_NAME_USERS);
  }
  const userHeaders = ['LINE User ID', 'LINE Display Name', '註冊時間'];
  usersSheet.getRange(1, 1, 1, userHeaders.length).setValues([userHeaders]);
  usersSheet.getRange(1, 1, 1, userHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  usersSheet.setFrozenRows(1);

  let menuSheet = ss.getSheetByName(SHEET_NAME_MENU);
  if (!menuSheet) {
    menuSheet = ss.insertSheet(SHEET_NAME_MENU);
  }
  menuSheet.clear();

  const menuHeaders = ['分類', '名稱', '價格', '圖標', '狀態', '圖片', '客製選項', '成本', '安全庫存(每日限量，0或空白=不限量)', '可升級套餐(Y/N)', '套餐加購價'];
  const menuItems = [
    ['招牌', '⓪ 脆皮雞塊', 75, '🍗', '供應中', '', '', 0, '', '', 0],
    ['甜點', '① 巧克力熔岩佐冰淇淋', 75, '🍫', '供應中', '', '', 0, '', '', 0],
    ['甜點', '② 格子鬆餅佐冰淇淋', 75, '🧇', '供應中', '', '', 0, '', '', 0],
    ['主食', '③ 任選牛／魚／雞肉粥', 75, '🍲', '供應中', '', '牛肉,魚,雞肉', 0, '', '', 0],
    ['甜點', '④ 蜜糖吐司佐冰淇淋', 75, '🍞', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑤ 炸花生湯圓佐冰淇淋', 75, '🍡', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑥ 美式鬆餅佐冰淇淋', 75, '🥞', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑦ 阿爾薩斯蘋果牛角', 75, '🥐', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑧ 焦糖布丁佐冰淇淋', 75, '🍮', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑨ 紫米紅豆露佐冰淇淋', 75, '🍧', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑩ 蓮紅棗蜜蘋甜湯', 75, '🥣', '供應中', '', '', 0, '', '', 0],
    ['主食', '⑪ 任選涼麵．飲料或湯', 75, '🍜', '供應中', '', '飲料,湯', 0, '', '', 0],
    ['主食', '⑫ 脆薯．飲料或湯', 75, '🍟', '供應中', '', '飲料,湯', 0, '', '', 0],
    ['甜點', '⑬ 義式奶酪佐冰淇淋', 75, '🍨', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑭ 波士頓花生冰淇淋堡', 75, '🥯', '供應中', '', '', 0, '', '', 0],
    ['主食', '⑮ 蒜香拌義大利天使麵', 75, '🍝', '供應中', '', '', 0, '', '', 0],
    ['主食', '⑯「哈」也不是鍋貼', 75, '🥟', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑰ 雞蛋糕佐冰淇淋', 75, '🧁', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑱ 法式布蕾佐冰淇淋', 75, '🍮', '供應中', '', '', 0, '', '', 0],
    ['沙拉', '⑲ 水果鮮蔬沙拉', 75, '🥗', '供應中', '', '', 0, '', '', 0],
    ['甜點', '⑳ 鮮爆泡芙佐冰淇淋', 75, '🧁', '供應中', '', '', 0, '', '', 0]
  ];

  menuSheet.getRange(1, 1, 1, menuHeaders.length).setValues([menuHeaders]);
  menuSheet.getRange(1, 1, 1, menuHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  menuSheet.getRange(2, 1, menuItems.length, menuHeaders.length).setValues(menuItems);
  menuSheet.setFrozenRows(1);
  const menuWidths = [80, 220, 60, 50, 80, 220, 180, 70, 100];
  menuWidths.forEach((w, i) => menuSheet.setColumnWidth(i + 1, w));

  let addonsSheet = ss.getSheetByName(SHEET_NAME_ADDONS);
  if (!addonsSheet) {
    addonsSheet = ss.insertSheet(SHEET_NAME_ADDONS);
  }
  addonsSheet.clear();
  const addonsHeaders = ['名稱', '價格', '狀態'];
  const addonsItems = [
    ['加蛋', 10, '供應中'],
    ['加起司', 15, '供應中'],
    ['多醬', 5, '供應中']
  ];
  addonsSheet.getRange(1, 1, 1, addonsHeaders.length).setValues([addonsHeaders]);
  addonsSheet.getRange(1, 1, 1, addonsHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  addonsSheet.getRange(2, 1, addonsItems.length, addonsHeaders.length).setValues(addonsItems);
  addonsSheet.setFrozenRows(1);
  const addonsWidths = [180, 80, 100];
  addonsWidths.forEach((w, i) => addonsSheet.setColumnWidth(i + 1, w));

  let specsSheet = ss.getSheetByName(SHEET_NAME_OPTION_SPECS);
  if (!specsSheet) {
    specsSheet = ss.insertSheet(SHEET_NAME_OPTION_SPECS);
  }
  specsSheet.clear();
  const specsHeaders = ['名稱', '選項清單', '上限', '必選', '早餐時段預設值', '正常時段預設值'];
  const specsItems = [
    ['蛋白質', '牛肉:0,魚:0,雞肉:0', 1, 1, '', ''],
    ['湯或飲料', '飲料:0,湯:0', 1, 1, '', ''],
    ['甜度', '正常:0,少糖:0,無糖:0', 1, 0, '', '']
  ];
  specsSheet.getRange(1, 1, 1, specsHeaders.length).setValues([specsHeaders]);
  specsSheet.getRange(1, 1, 1, specsHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  specsSheet.getRange(2, 1, specsItems.length, specsHeaders.length).setValues(specsItems);
  specsSheet.setFrozenRows(1);
  const specsWidths = [140, 260, 60, 60];
  specsWidths.forEach((w, i) => specsSheet.setColumnWidth(i + 1, w));

  Logger.log('✅ A表（點餐主表）工作表已初始化完成，共 9 張工作表。若尚未初始化B表（營業日報表），請接著執行 setupReportSheets()');
}

function setupReportSheets() {
  const ss = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);

  let fixedCostsSheet = ss.getSheetByName(SHEET_NAME_FIXED_COSTS);
  if (!fixedCostsSheet) {
    fixedCostsSheet = ss.insertSheet(SHEET_NAME_FIXED_COSTS);
  }
  fixedCostsSheet.clear();
  const fixedCostsHeaders = ['項目名稱', '金額'];
  fixedCostsSheet.getRange(1, 1, 1, fixedCostsHeaders.length).setValues([fixedCostsHeaders]);
  fixedCostsSheet.getRange(1, 1, 1, fixedCostsHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  fixedCostsSheet.setFrozenRows(1);
  const fixedCostsWidths = [200, 120];
  fixedCostsWidths.forEach((w, i) => fixedCostsSheet.setColumnWidth(i + 1, w));

  let costCategoriesSheet = ss.getSheetByName(SHEET_NAME_COST_CATEGORIES);
  if (!costCategoriesSheet) {
    costCategoriesSheet = ss.insertSheet(SHEET_NAME_COST_CATEGORIES);
  }
  costCategoriesSheet.clear();
  const costCategoriesHeaders = ['類型', 'Key', '名稱'];
  costCategoriesSheet.getRange(1, 1, 1, costCategoriesHeaders.length).setValues([costCategoriesHeaders]);
  costCategoriesSheet.getRange(1, 1, 1, costCategoriesHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  costCategoriesSheet.setFrozenRows(1);
  [180, 200, 200].forEach((w, i) => costCategoriesSheet.setColumnWidth(i + 1, w));

  let dailyRecordsSheet = ss.getSheetByName(SHEET_NAME_DAILY_RECORDS);
  if (!dailyRecordsSheet) {
    dailyRecordsSheet = ss.insertSheet(SHEET_NAME_DAILY_RECORDS);
  }
  dailyRecordsSheet.clear();
  const dailyRecordsHeaders = ['日期', '明細記錄(JSON)', '實收入(JSON)', '備註'];
  dailyRecordsSheet.getRange(1, 1, 1, dailyRecordsHeaders.length).setValues([dailyRecordsHeaders]);
  dailyRecordsSheet.getRange(1, 1, 1, dailyRecordsHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  dailyRecordsSheet.setFrozenRows(1);
  [110, 350, 260, 200].forEach((w, i) => dailyRecordsSheet.setColumnWidth(i + 1, w));

  let purchaseItemsSheet = ss.getSheetByName(SHEET_NAME_PURCHASE_ITEMS);
  if (!purchaseItemsSheet) {
    purchaseItemsSheet = ss.insertSheet(SHEET_NAME_PURCHASE_ITEMS);
  }
  purchaseItemsSheet.clear();
  const purchaseItemsHeaders = ['代號', '品名', '類別', '單位', '參考單價'];
  purchaseItemsSheet.getRange(1, 1, 1, purchaseItemsHeaders.length).setValues([purchaseItemsHeaders]);
  purchaseItemsSheet.getRange(1, 1, 1, purchaseItemsHeaders.length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2980B9')
    .setHorizontalAlignment('center');
  purchaseItemsSheet.setFrozenRows(1);
  [100, 200, 120, 80, 120].forEach((w, i) => purchaseItemsSheet.setColumnWidth(i + 1, w));

  Logger.log('✅ B表（營業日報表）工作表已初始化完成，共 4 張工作表（「採購單記錄」「本地系統日報」會在第一次真的送出資料時自動建立，不用先手動建）');
}

const EXPIRE_HOURS = 3;

function expireOldOrders() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const STATUS_COL = 2;
    const CREATED_COL = 3;
    const UPDATED_COL = 16;

    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const now = new Date();
    let expiredCount = 0;

    data.forEach((row, i) => {
      const status = row[STATUS_COL - 1];
      const createdStr = row[CREATED_COL - 1];
      if (status !== '未匯款' || !createdStr) return;

      const created = new Date(String(createdStr).replace(' ', 'T'));
      if (isNaN(created.getTime())) return;

      const hoursPassed = (now.getTime() - created.getTime()) / (1000 * 60 * 60);
      if (hoursPassed >= EXPIRE_HOURS) {
        const rowIndex = i + 2;
        sheet.getRange(rowIndex, STATUS_COL).setValue('已逾期');
        sheet.getRange(rowIndex, UPDATED_COL).setValue(
          Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss')
        );
        expiredCount++;
      }
    });

    Logger.log(`✅ expireOldOrders 執行完畢，共標記 ${expiredCount} 筆逾期訂單`);
  } catch (err) {
    Logger.log('expireOldOrders 錯誤: ' + err.message);
  }
}

function onEditInstallable(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME_ORDERS) return;

    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (row === 1) return;

    const COL_CONFIRM = 14;
    const COL_STATUS  = 2;
    const COL_UPDATED = 16;

    if (col === COL_CONFIRM) {
      const val = e.range.getValue();
      if (val === '已確認') {
        sheet.getRange(row, COL_STATUS).setValue('已匯款');
      }
      sheet.getRange(row, COL_UPDATED).setValue(
        Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss')
      );
    }
  } catch (err) {
    Logger.log('onEditInstallable 錯誤: ' + err.message);
  }
}
