/** 幾 A 幾 B 版本三：代號識別 + GAS + Google Sheets。
 * 代號不是身分驗證。知道代號的人能查看與操作該代號的紀錄。
 * 答案與 A/B 判斷都在後端；未猜對的答案不會送到玩家畫面，
 * 只有「除錯模式」(action: reveal) 會依「目前代號」取回該局原本存在雲端的答案。
 */
const SETTINGS = Object.freeze({sheetName:'遊戲存檔', maxRounds:50, maxGuesses:100});
const HEADERS = ['使用者代號','局ID','局數','狀態','開始時間','更新時間','猜測次數','猜測歷程','完整存檔JSON'];

// 在編輯器選 setup 執行。可重複初始化，不清空資料；網址只寫入執行記錄，不回傳前端。
function setup() {
  locked_(() => {
    const props = PropertiesService.getScriptProperties();
    let id = props.getProperty('GUESS_BOOK_ID');
    if (!id) {
      const book = SpreadsheetApp.create('數位科技概論｜幾A幾B版本三存檔');
      book.setSpreadsheetTimeZone('Asia/Taipei');
      id = book.getId();
      props.setProperty('GUESS_BOOK_ID', id);
    }
    const book = SpreadsheetApp.openById(id);
    let sheet = book.getSheetByName(SETTINGS.sheetName);
    if (!sheet) { sheet = book.insertSheet(SETTINGS.sheetName); sheet.appendRow(HEADERS); sheet.setFrozenRows(1); }
    checkHeaders_(sheet);
    console.log('請保留此試算表網址：' + book.getUrl());
    console.log('接著部署為網頁應用程式；執行身分選「我」，勿改成存取者。');
  });
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('幾A幾B版本三｜雲端紀錄')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function fail_(message) { const error = new Error(message); error.isPublic = true; throw error; }
function locked_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return callback(); } finally { try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); } }
}
function normalizeCode_(value) {
  if (typeof value !== 'string') fail_('請輸入使用者代號。');
  const code = value.trim().toLowerCase();
  if (!/^[a-z0-9_-]{3,24}$/.test(code)) fail_('代號需為 3～24 個英文字母、數字、底線或連字號。');
  return code;
}
function id_(value) {
  if (typeof value !== 'string' || !/^[a-zA-Z0-9_-]{12,80}$/.test(value)) fail_('操作識別碼不正確，請重新載入網頁。');
  return value;
}
// 猜測一律用「字串」處理，才不會遺失開頭的 0。
function guessValue_(value) {
  if (typeof value !== 'string') fail_('請輸入四位數字（可以 0 開頭）。');
  const text = value.trim();
  if (!/^\d{4}$/.test(text)) fail_('請輸入剛好四位數字（可以 0 開頭）。');
  if (new Set(text.split('')).size !== 4) fail_('四個數字不能重複。');
  return text;
}
// A：數字與位置都對；B：數字對、位置不同。答案與猜測都沒有重複數字，所以不會重複計算。
function score_(guess, answer) {
  let a = 0, b = 0;
  for (let i = 0; i < 4; i++) {
    if (guess[i] === answer[i]) a++;
    else if (answer.indexOf(guess[i]) >= 0) b++;
  }
  return {a: a, b: b};
}
function randomAnswer_() {
  const digits = ['0','1','2','3','4','5','6','7','8','9'];
  for (let i = digits.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = digits[i]; digits[i] = digits[j]; digits[j] = t;
  }
  return digits.slice(0, 4).join('');
}
function sheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('GUESS_BOOK_ID');
  if (!id) fail_('尚未初始化。請製作者在 GAS 編輯器先執行 setup。');
  const sheet = SpreadsheetApp.openById(id).getSheetByName(SETTINGS.sheetName);
  if (!sheet) fail_('找不到遊戲存檔分頁，請製作者檢查初始化。');
  checkHeaders_(sheet); return sheet;
}
function checkHeaders_(sheet) {
  const actual = sheet.getRange(1,1,1,HEADERS.length).getValues()[0];
  if (actual.join('|') !== HEADERS.join('|')) fail_('存檔欄位與程式不符，請製作者確認表頭；不要自行改動欄位。');
}
function readPlayer_(sheet,code) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,HEADERS.length).getValues()
    .map((values,index) => ({values:values,row:index+2})).filter(record => record.values[0] === code)
    .map(record => {
      let game;
      try { game = JSON.parse(record.values[8]); } catch (_) { fail_('此代號有損壞的存檔，請製作者檢查試算表。'); }
      if (game.player !== code || game.id !== record.values[1] || !Array.isArray(game.guesses)) fail_('存檔不一致，請製作者檢查資料。');
      return {row:record.row,game:game};
    });
}
function write_(sheet,row,game) {
  const label = {playing:'未完成',won:'已猜對'}[game.status];
  const history = game.guesses.map(g => g.order + '. ' + g.value + ' → ' + g.a + 'A' + g.b + 'B').join('\n');
  const json = JSON.stringify(game);
  if (json.length > 40000) fail_('此局資料太大，請開新局。');
  // 先把整列設成純文字格式，再寫入，四位數字的開頭 0 才不會被試算表吃掉。
  sheet.getRange(row,1,1,HEADERS.length).setNumberFormat('@').setValues([[
    game.player,game.id,game.number,label,game.startedAt,game.updatedAt,game.guesses.length,history,json
  ]]);
  SpreadsheetApp.flush();
}
// 不把尚未猜對的答案送到玩家瀏覽器；後端負責判斷。
function publicGame_(game) {
  return {id:game.id,number:game.number,status:game.status,revision:game.revision,
    startedAt:game.startedAt,updatedAt:game.updatedAt,endedAt:game.endedAt,
    answer:game.status === 'won' ? game.answer : null,
    guesses:game.guesses.map(g=>({order:g.order,value:g.value,a:g.a,b:g.b,at:g.at}))};
}
function snapshot_(code,records,selectedId,message) {
  return {ok:true,code:code,message:message,selectedId:selectedId || null,
    rounds:records.map(r=>publicGame_(r.game)).sort((a,b)=>a.number-b.number),
    limits:{rounds:SETTINGS.maxRounds,guesses:SETTINGS.maxGuesses}};
}
// 單一公開入口：每次都依代號重新篩選，不能只憑局 ID 取別人的局。
function api(request) {
  try {
    if (!request || typeof request !== 'object') fail_('請求格式錯誤。');
    const code = normalizeCode_(request.code);
    if (['load','new','guess','reveal'].indexOf(request.action) < 0) fail_('未知的操作。');
    return locked_(() => {
      const sheet = sheet_(), records = readPlayer_(sheet,code);
      if (request.action === 'load') return snapshot_(code,records,null,records.length ? '已讀回此代號的雲端紀錄。' : '此代號沒有紀錄，可以開啟第一局。');
      if (request.action === 'reveal') {
        // 除錯模式：只回傳「此代號」名下、指定局原本存在雲端的答案；不產生新答案、不改任何資料。
        const rid = id_(request.gameId);
        const rec = records.find(r=>r.game.id === rid);
        if (!rec) fail_('這個代號找不到指定的局，請重新讀取。');
        return {ok:true,code:code,gameId:rid,answer:rec.game.answer,message:'除錯模式：已從雲端存檔取得這一局的答案。'};
      }
      if (request.action === 'new') {
        const requestId = id_(request.requestId);
        const existing = records.find(r=>r.game.id === requestId);
        if (existing) return snapshot_(code,records,existing.game.id,'這一局已建立；重試沒有再新增一局。');
        if (records.length >= SETTINGS.maxRounds) fail_('此代號已達課堂範例的 50 局上限；舊紀錄會保留。');
        const now = new Date().toISOString();
        const game = {id:requestId,player:code,number:records.reduce((n,r)=>Math.max(n,r.game.number),0)+1,
          answer:randomAnswer_(),status:'playing',revision:0,startedAt:now,updatedAt:now,endedAt:null,guesses:[]};
        const row = sheet.getLastRow()+1;
        write_(sheet,row,game); records.push({row:row,game:game});
        return snapshot_(code,records,game.id,'新局已建立並存到雲端。');
      }
      const requestId = id_(request.requestId), gameId = id_(request.gameId), value = guessValue_(request.value);
      const record = records.find(r=>r.game.id === gameId);
      if (!record) fail_('這個代號找不到指定的局，請重新讀取。');
      const game = record.game;
      const duplicate = game.guesses.find(g=>g.requestId === requestId);
      if (duplicate) {
        if (duplicate.value !== value) fail_('同一操作識別碼不能改送不同數字，請重新讀取。');
        return snapshot_(code,records,game.id,'這次猜測已存過；重試沒有重複記錄。');
      }
      if (!Number.isInteger(request.revision) || request.revision !== game.revision) {
        const snap = snapshot_(code,records,game.id,'此局在另一個分頁或裝置已更新；已讀回最新資料，請再確認後猜測。');
        snap.conflict = true; return snap;
      }
      if (game.status !== 'playing') fail_('這一局已猜對，請開新局。');
      if (game.guesses.length >= SETTINGS.maxGuesses) fail_('此局已達 100 次猜測上限，請開新局；紀錄會保留。');
      const r = score_(value,game.answer), now = new Date().toISOString();
      game.guesses.push({requestId:requestId,order:game.guesses.length+1,value:value,a:r.a,b:r.b,at:now});
      game.revision++; game.updatedAt = now;
      if (r.a === 4) { game.status = 'won'; game.endedAt = now; }
      write_(sheet,record.row,game);
      return snapshot_(code,records,game.id,value + '：' + r.a + 'A' + r.b + 'B。' + (r.a === 4 ? '猜對了！' : '') + '已儲存到雲端。');
    });
  } catch (error) {
    if (error.isPublic) return {ok:false,error:error.message};
    console.error(error.stack || error.message);
    return {ok:false,uncertain:true,error:'雲端操作未確認。請重試同一操作，或重新讀取；製作者可在 GAS 執行項目查看錯誤。'};
  }
}
