/**
 * 社員意見箱 データ保存用 Google Apps Script
 *
 * スプレッドシートに「意見」「コメント」「投票済ログ」「社員名簿」「設定」「操作履歴」の各シートを作り、
 * 意見箱の画面（index.html）からの操作を受け付けて読み書きする。
 *
 * 匿名性のための設計：
 *  - 意見・コメントには投稿者を一切記録しない。
 *  - 「誰が投票したか」は投票済ログ（社員番号・意見ID・日付）にだけ記録し、賛成/反対の中身は記録しない。
 *  - 「賛成・反対の数」は投票期間中はスクリプトのプロパティに保存し、締切後にだけ意見シートへ書き出す。
 *    （シートの変更履歴から「このログの行と同時に賛成が1増えた」と突き合わせられないようにするため）
 *
 * 同時操作への対策：
 *  - 画面からは「意見12に賛成」のような操作だけを送り、全データの上書きはしない。
 *  - 書き込みはすべて LockService で1件ずつ順番に処理する。
 *
 * セットアップ手順は同じフォルダの README.md を参照。
 */

const SHEETS = {
  opinions: "意見",
  comments: "コメント",
  voted: "投票済ログ",
  employees: "社員名簿",
  settings: "設定",
  log: "操作履歴"
};

// 各シートの列（順番を変えないこと）
const COLS = {
  opinions: [
    ["id", "ID"], ["status", "状態"], ["category", "カテゴリ"], ["article", "条文"], ["title", "タイトル"],
    ["current", "現状"], ["problem", "困っていること"], ["proposal", "提案"], ["effect", "期待できる効果"],
    ["postedAt", "投稿日"], ["deadline", "投票期限"],
    ["yes", "賛成（確定）"], ["no", "反対（確定）"], ["voters", "投票者数（確定）"],
    ["submittedAt", "提出日"], ["answerDue", "回答期限"],
    ["result", "会社回答"], ["reason", "回答理由"], ["answeredAt", "回答日"],
    ["planDate", "実施予定日"], ["doneAt", "実施日"], ["hidden", "非表示"]
  ],
  comments: [["id", "ID"], ["opinionId", "意見ID"], ["stance", "立場"], ["body", "本文"], ["postedAt", "投稿日"], ["hidden", "非表示"]],
  voted: [["empNo", "社員番号"], ["opinionId", "意見ID"], ["date", "投票日"]],
  employees: [
    ["empNo", "社員ID"], ["name", "氏名"], ["dept", "部署"], ["email", "メールアドレス"],
    ["active", "在籍（○/×）"], ["pwHash", "パスワード（ハッシュ）"], ["salt", "ソルト"], ["registeredAt", "登録日"]
  ],
  settings: [["key", "項目"], ["value", "値"], ["note", "説明"]],
  log: [["at", "日時"], ["actor", "操作者"], ["action", "操作"], ["detail", "内容"]]
};

const DEFAULT_SETTINGS = [
  ["投票期間（日）", "14", "新しい意見の投票期限を、投稿日から何日後にするか"],
  ["回答期限（日）", "30", "提出した日から、会社の回答期限を何日後にするか"],
  ["提出先", "佐野ケーブルテレビ 御中", "意見書（PDF）の宛先"],
  ["提出者名義", "社員意見箱 運営", "意見書（PDF）の提出者欄"]
];

const PW_MIN = 8;
const PW_ROUNDS = 1000;               // パスワードのハッシュを重ねる回数（総当たり対策）
const SESSION_DAYS = 90;              // ログインしたままでいられる日数
const MAX_REGISTER_PER_HOUR = 30;     // 1時間あたりの新規登録の上限（いたずら対策）

const CATEGORIES = ["就業規則", "勤務・シフト", "業務改善", "設備・環境", "福利厚生", "その他"];
const STANCES = ["賛成の立場", "反対の立場", "質問・その他"];
const RESULTS = ["可決", "一部可決", "否決", "継続検討"];

const STATUS = { open: "投票中", skipped: "見送り", submitted: "提出済", answered: "回答済", done: "実施済" };

const LIMITS = { title: 60, field: 1000, article: 30, comment: 500, reason: 2000 };
const MAX_FAILURES = 10;        // この回数ログインに失敗すると
const LOCK_SECONDS = 10 * 60;   // この秒数、そのアカウントでのログインを受け付けない
const TZ = "Asia/Tokyo";
const CODE_CHARS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // 見間違えやすい 0/O・1/I を除く

// ===== スプレッドシートのメニュー =====

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("意見箱")
    .addItem("初期設定（シートを作成）", "setupSheets")
    .addSeparator()
    .addItem("運営パスワードを設定", "setAdminPassword")
    .addItem("会社用パスワードを設定", "setCompanyPassword")
    .addToUi();
}

function setupSheets() {
  ensureSheets_();
  SpreadsheetApp.getUi().alert("シートを作成しました。社員は意見箱の画面から自分で登録します。");
}

function setAdminPassword() { promptPassword_("ADMIN_PW", "運営パスワードを設定", "意見箱の「運営」ログインに使うパスワードを入力してください。"); }
function setCompanyPassword() { promptPassword_("COMPANY_PW", "会社用パスワードを設定", "意見箱の「会社」ログインに使うパスワードを入力してください。"); }

function promptPassword_(prop, title, message) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const pw = res.getResponseText().trim();
  if (pw.length < 6) { ui.alert("6文字以上で設定してください。設定しませんでした。"); return; }
  PropertiesService.getScriptProperties().setProperty(prop, sha256Hex_(pw));
  ui.alert("設定しました。");
}

// ===== Web API =====

// GET：パラメータ p があれば、POST と同じ操作として処理する。
// スマホ（Safari）では POST の途中で Google に転送されて送信内容が消えることがあるため、画面側はその場合に GET で送り直す。
// p は操作内容の JSON を UTF-8 → Base64（URL用）にしたもの。
function doGet(e) {
  const p = e && e.parameter && e.parameter.p;
  if (!p) return json_({ ok: true, app: "ikenbako" });
  let contents;
  try {
    contents = Utilities.newBlob(Utilities.base64DecodeWebSafe(p)).getDataAsString("UTF-8");
  } catch (err) {
    return json_({ ok: false, error: "送信内容を読み取れませんでした。" });
  }
  return handle_(contents);
}

// ブラウザからは CORS のプリフライトを避けるため Content-Type: text/plain で JSON を送る。
function doPost(e) {
  return handle_(e && e.postData ? e.postData.contents : "");
}

let CURRENT_RID = ""; // 処理中の操作の番号（送り直しによる二重処理を防ぐ）

function handle_(contents) {
  try {
    const req = JSON.parse(contents);
    CURRENT_RID = String(req.rid || "").slice(0, 40);
    ensureSheets_();
    if (req.action === "register") return json_(register_(req));
    const user = authenticate_(req.auth || {});
    return json_(dispatch_(req, user));
  } catch (err) {
    return json_({ ok: false, error: err && err.userMessage ? err.userMessage : "エラーが発生しました：" + err, code: err && err.code });
  }
}

function dispatch_(req, user) {
  const a = req.action;
  if (a === "login" || a === "list") return listPayload_(user);
  if (a === "logout") { if (user.sessionKey) PropertiesService.getScriptProperties().deleteProperty(user.sessionKey); return { ok: true }; }

  if (a === "post") { need_(user, "emp"); return withLock_(() => postOpinion_(req), user); }
  if (a === "vote") { need_(user, "emp"); return withLock_(() => vote_(user, req), user); }
  if (a === "comment") { need_(user, "emp"); return withLock_(() => addComment_(req), user); }
  if (a === "changePassword") {
    need_(user, "emp");
    const payload = withLock_(() => changePassword_(user, req), user);
    payload.token = createSession_(findUser_(user.empNo)); // この端末はログインしたままにする
    return payload;
  }

  if (a === "setDeadline") { need_(user, "admin"); return withLock_(() => setDeadline_(req), user); }
  if (a === "setHidden") { need_(user, "admin"); return withLock_(() => setHidden_(req), user); }
  if (a === "setStatus") { need_(user, "admin"); return withLock_(() => setStatus_(req), user); }
  if (a === "submit") { need_(user, "admin"); return withLock_(() => submit_(req), user); }
  if (a === "deleteUser") { need_(user, "admin"); return withLock_(() => deleteUser_(req), user); }
  if (a === "resetPassword") {
    need_(user, "admin");
    let temp;
    const payload = withLock_(() => { temp = resetPassword_(req); }, user);
    payload.tempPassword = temp;
    return payload;
  }

  if (a === "answer") { need_(user, "company"); return withLock_(() => answer_(req), user); }

  throw userError_("不明な操作です。");
}

// 書き込みはロックを取って1件ずつ処理し、処理後の最新一覧を返す。
function withLock_(fn, user) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw userError_("混み合っています。少し待ってからもう一度お試しください。");
  let message;
  const cache = CacheService.getScriptCache();
  const ridKey = CURRENT_RID ? "rid_" + CURRENT_RID : "";
  try {
    if (ridKey && cache.get(ridKey)) {
      message = cache.get(ridKey); // 同じ操作がすでに処理済み（送り直し）なので、もう一度は実行しない
    } else {
      message = fn();
      SpreadsheetApp.flush();
      if (ridKey) cache.put(ridKey, message || "処理しました。", 600);
    }
  } finally {
    lock.releaseLock();
  }
  const payload = listPayload_(user);
  if (message) payload.message = message;
  return payload;
}

// ===== 社員の新規登録 =====

function register_(req) {
  const name = text_(req.name, 30, "氏名", true);
  const dept = text_(req.dept, 30, "部署", false);
  const email = normEmail_(req.email);
  const pw = String(req.password || "");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 100) throw userError_("メールアドレスが正しくありません。");
  checkPassword_(pw);

  const cache = CacheService.getScriptCache();
  const count = Number(cache.get("register_count") || 0);
  if (count >= MAX_REGISTER_PER_HOUR) throw userError_("登録が集中しています。しばらくしてからもう一度お試しください。");

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw userError_("混み合っています。少し待ってからもう一度お試しください。");
  let user;
  try {
    const emps = readTable_("employees");
    const key = normName_(name);
    if (emps.some(e => e.name && normName_(e.name) === key)) {
      throw userError_("同じ氏名の方がすでに登録されています。心当たりがない場合は運営に問い合わせてください。");
    }
    if (emps.some(e => normEmail_(e.email) === email)) {
      throw userError_("このメールアドレスはすでに登録されています。");
    }
    const props = PropertiesService.getScriptProperties();
    // 社員IDは削除後も使い回さない（投票済ログが別の人に引き継がれないように）
    const seq = Math.max(Number(props.getProperty("USER_SEQ") || 0),
      emps.reduce((m, e) => Math.max(m, Number(String(e.empNo).replace(/\D/g, "")) || 0), 0)) + 1;
    props.setProperty("USER_SEQ", String(seq));
    const salt = Utilities.getUuid();
    user = {
      empNo: "U" + ("000" + seq).slice(-4), name: name, dept: dept, email: email, active: "○",
      pwHash: hashPw_(pw, salt), salt: salt, registeredAt: today_()
    };
    appendRow_("employees", user);
    appendLog_("（新規登録）", "アカウント登録", user.empNo + " " + name);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  cache.put("register_count", String(count + 1), 3600);

  const payload = listPayload_({ role: "emp", empNo: user.empNo, name: user.name });
  payload.token = createSession_(user);
  payload.message = "登録しました。";
  return payload;
}

function changePassword_(user, req) {
  const emp = findUser_(user.empNo);
  if (hashPw_(String(req.current || ""), emp.salt) !== emp.pwHash) throw userError_("現在のパスワードが違います。");
  const pw = String(req.password || "");
  checkPassword_(pw);
  const salt = Utilities.getUuid();
  setCells_("employees", emp._row, { pwHash: hashPw_(pw, salt), salt: salt });
  return "パスワードを変更しました。ほかの端末では再ログインが必要です。";
}

// ===== 運営による社員管理 =====

// アカウントを削除する（退職・重複登録・なりすましなど）。投票済ログは残るが、投票率の分母からは外れる。
function deleteUser_(req) {
  const emp = findUser_(req.empNo);
  SpreadsheetApp.getActive().getSheetByName(SHEETS.employees).deleteRow(emp._row);
  appendLog_("運営", "アカウント削除", emp.empNo + " " + emp.name);
  return emp.name + " さんのアカウントを削除しました。";
}

// パスワードを忘れた社員に仮パスワードを発行する（運営が本人に伝える）
function resetPassword_(req) {
  const emp = findUser_(req.empNo);
  const temp = randomCode_();
  const salt = Utilities.getUuid();
  setCells_("employees", emp._row, { pwHash: hashPw_(temp, salt), salt: salt });
  appendLog_("運営", "仮パスワード発行", emp.name);
  return temp;
}

function findUser_(empNo) {
  const emp = readTable_("employees").find(e => String(e.empNo) === String(empNo));
  if (!emp) throw userError_("社員が見つかりません。");
  return emp;
}

// ===== 認証 =====

function authenticate_(auth) {
  const props = PropertiesService.getScriptProperties();
  const role = auth.role;

  // ログイン済みの社員（トークン）
  if (role === "emp" && auth.token) {
    const key = sessionKey_(auth.token);
    const raw = props.getProperty(key);
    const s = raw ? JSON.parse(raw) : null;
    if (!s || s.e < Date.now()) throw userError_("ログインの有効期限が切れました。もう一度ログインしてください。", "SESSION");
    const emp = readTable_("employees").find(e => String(e.empNo) === s.u);
    if (!emp || !isActive_(emp) || emp.pwHash.slice(0, 16) !== s.v) {
      props.deleteProperty(key);
      throw userError_("もう一度ログインしてください。", "SESSION");
    }
    return { role: "emp", empNo: String(emp.empNo), name: emp.name, sessionKey: key };
  }

  const id = role === "emp" ? normEmail_(auth.email) : role;
  const failKey = "fail_" + role + "_" + String(id || "").slice(0, 60);
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILURES) throw userError_("ログインの失敗が続いたため、10分間ログインできません。");
  const fail = msg => {
    cache.put(failKey, String(fails + 1), LOCK_SECONDS);
    throw userError_(msg);
  };

  // 社員のログイン（メールアドレス＋パスワード）→ トークンを発行
  if (role === "emp") {
    const pw = String(auth.password || "");
    if (!id || !pw) fail("メールアドレスとパスワードを入力してください。");
    const emp = readTable_("employees").find(e => normEmail_(e.email) === id);
    if (!emp || !emp.salt || hashPw_(pw, emp.salt) !== emp.pwHash) fail("メールアドレスまたはパスワードが違います。");
    if (!isActive_(emp)) throw userError_("このアカウントは現在利用できません。運営に問い合わせてください。");
    cache.remove(failKey);
    const user = { role: "emp", empNo: String(emp.empNo), name: emp.name };
    user.newToken = createSession_(emp);
    return user;
  }

  if (role === "admin" || role === "company") {
    const stored = props.getProperty(role === "admin" ? "ADMIN_PW" : "COMPANY_PW");
    if (!stored) throw userError_((role === "admin" ? "運営" : "会社用") + "パスワードがまだ設定されていません。スプレッドシートのメニューから設定してください。");
    if (sha256Hex_(String(auth.code || "")) !== stored) fail("パスワードが違います。");
    return { role: role };
  }
  throw userError_("ログインしてください。", "SESSION");
}

function need_(user, role) {
  if (user.role !== role) throw userError_("この操作は行えません。");
}

// ログイン状態（トークン）はスクリプトのプロパティに保存する。トークンそのものではなくハッシュ値をキーにする。
function createSession_(emp) {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  // 期限切れのログイン状態を掃除
  props.getKeys().forEach(k => {
    if (k.indexOf("S_") !== 0) return;
    try { if (JSON.parse(props.getProperty(k)).e < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
  const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  props.setProperty(sessionKey_(token), JSON.stringify({ u: String(emp.empNo), v: String(emp.pwHash).slice(0, 16), e: now + SESSION_DAYS * 86400000 }));
  return token;
}

function sessionKey_(token) { return "S_" + sha256Hex_(String(token)).slice(0, 40); }

function hashPw_(pw, salt) {
  let h = sha256Hex_(salt + ":" + pw);
  for (let i = 0; i < PW_ROUNDS; i++) h = sha256Hex_(h + salt);
  return h;
}

function checkPassword_(pw) {
  if (pw.length < PW_MIN) throw userError_("パスワードは" + PW_MIN + "文字以上にしてください。");
  if (pw.length > 100) throw userError_("パスワードが長すぎます。");
}

function normEmail_(v) { return String(v || "").trim().toLowerCase(); }

// 氏名の重複判定用：全角・半角をそろえ、空白をすべて取り除く（「吉沢 大将」と「吉沢大将」を同一とみなす）
function normName_(v) { return String(v || "").normalize("NFKC").replace(/\s+/g, ""); }

// ===== 一覧 =====

function listPayload_(user) {
  finalizeClosed_();
  const today = today_();
  const opinions = readTable_("opinions");
  const comments = readTable_("comments");
  const voted = readTable_("voted");
  const employees = readTable_("employees").filter(e => e.empNo && isActive_(e));
  const props = PropertiesService.getScriptProperties();
  const isAdmin = user.role === "admin";

  const votersBy = {};
  voted.forEach(v => { const k = String(v.opinionId); votersBy[k] = (votersBy[k] || 0) + 1; });

  const commentsBy = {};
  comments.forEach(c => {
    if (isTrue_(c.hidden) && !isAdmin) return;
    const k = String(c.opinionId);
    (commentsBy[k] = commentsBy[k] || []).push({
      id: Number(c.id), stance: c.stance, body: c.body, postedAt: c.postedAt, hidden: isTrue_(c.hidden)
    });
  });

  const list = opinions
    .filter(o => o.id !== "" && (isAdmin || !isTrue_(o.hidden)))
    .map(o => {
      const id = String(o.id);
      const finalized = o.yes !== "";
      const t = finalized ? { y: Number(o.yes), n: Number(o.no) } : readTally_(props, id);
      return {
        id: Number(o.id), status: o.status, category: o.category, article: o.article, title: o.title,
        current: o.current, problem: o.problem, proposal: o.proposal, effect: o.effect,
        postedAt: o.postedAt, deadline: o.deadline,
        closed: o.status !== STATUS.open || o.deadline < today,
        yes: t.y, no: t.n, voters: finalized ? Number(o.voters) : (votersBy[id] || 0),
        submittedAt: o.submittedAt, answerDue: o.answerDue,
        result: o.result, reason: o.reason, answeredAt: o.answeredAt,
        planDate: o.planDate, doneAt: o.doneAt,
        hidden: isTrue_(o.hidden),
        comments: commentsBy[id] || []
      };
    });

  const payload = {
    ok: true,
    role: user.role,
    name: user.name || "",
    today: today,
    activeCount: employees.length,
    settings: publicSettings_(),
    categories: CATEGORIES, stances: STANCES, results: RESULTS,
    opinions: list
  };

  if (user.newToken) payload.token = user.newToken;
  if (user.role === "emp") {
    payload.myVoted = voted.filter(v => String(v.empNo).trim() === user.empNo).map(v => Number(v.opinionId));
  }

  // 未投票者一覧は運営だけに返す（会社には個人ごとの投票状況を見せない）
  if (isAdmin) {
    const openIds = list.filter(o => o.status === STATUS.open && !o.closed && !o.hidden).map(o => o.id);
    const votedSet = {};
    voted.forEach(v => { votedSet[String(v.empNo).trim() + "|" + v.opinionId] = true; });
    payload.nonVoters = employees.map(e => {
      const missing = openIds.filter(id => !votedSet[String(e.empNo).trim() + "|" + id]);
      return { empNo: String(e.empNo), name: e.name, dept: e.dept, missing: missing };
    }).filter(e => e.missing.length);
    payload.users = readTable_("employees").filter(e => e.empNo).map(e => ({
      empNo: String(e.empNo), name: e.name, dept: e.dept, email: e.email, active: isActive_(e), registeredAt: e.registeredAt
    }));
  }
  return payload;
}

// 投票期限を過ぎた・投票中でなくなった意見の賛否を、プロパティからシートへ確定させる。
function finalizeClosed_() {
  const today = today_();
  const pending = readTable_("opinions").filter(o => o.id !== "" && o.yes === "" && (o.status !== STATUS.open || o.deadline < today));
  if (!pending.length) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; // 次の読み込み時に再試行する
  try {
    readTable_("opinions")
      .filter(o => o.id !== "" && o.yes === "" && (o.status !== STATUS.open || o.deadline < today))
      .forEach(o => finalizeOne_(o));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function finalizeOne_(o) {
  const props = PropertiesService.getScriptProperties();
  const t = readTally_(props, String(o.id));
  const voters = readTable_("voted").filter(v => String(v.opinionId) === String(o.id)).length;
  setCells_("opinions", o._row, { yes: t.y, no: t.n, voters: voters });
}

// ===== 社員の操作 =====

function postOpinion_(req) {
  const category = oneOf_(req.category, CATEGORIES, "カテゴリ", req.categoryIdx);
  const title = text_(req.title, LIMITS.title, "タイトル", true);
  const proposal = text_(req.proposal, LIMITS.field, "提案", true);
  const current = text_(req.current, LIMITS.field, "現状", false);
  const problem = text_(req.problem, LIMITS.field, "困っていること", false);
  const effect = text_(req.effect, LIMITS.field, "期待できる効果", false);
  const article = text_(req.article, LIMITS.article, "条文", false);

  const opinions = readTable_("opinions");
  const id = opinions.reduce((m, o) => Math.max(m, Number(o.id) || 0), 0) + 1;
  const days = Number(getSetting_("投票期間（日）")) || 14;
  const row = {
    id: id, status: STATUS.open, category: category, article: article, title: title,
    current: current, problem: problem, proposal: proposal, effect: effect,
    postedAt: today_(), deadline: addDays_(today_(), days), hidden: ""
  };
  appendRow_("opinions", row);
  // 投稿者は記録しない
  appendLog_("（匿名）", "意見投稿", "No." + id + " " + title);
  return "意見を投稿しました（No." + id + "）。";
}

function vote_(user, req) {
  const id = Number(req.id);
  const side = req.side === "yes" ? "y" : req.side === "no" ? "n" : "";
  if (!side) throw userError_("賛成か反対を選んでください。");
  const o = findOpinion_(id);
  if (isTrue_(o.hidden) || o.status !== STATUS.open) throw userError_("この意見は投票を受け付けていません。");
  if (o.deadline < today_()) throw userError_("投票期限を過ぎています。");
  const already = readTable_("voted").some(v => String(v.empNo).trim() === user.empNo && Number(v.opinionId) === id);
  if (already) throw userError_("この意見にはすでに投票しています。");

  // 投票済ログには賛否を書かない。賛否の数はプロパティにだけ加算する。
  appendRow_("voted", { empNo: user.empNo, opinionId: id, date: today_() });
  const props = PropertiesService.getScriptProperties();
  const t = readTally_(props, String(id));
  t[side] += 1;
  props.setProperty("T_" + id, JSON.stringify(t));
  return "投票しました。";
}

function addComment_(req) {
  const id = Number(req.id);
  const o = findOpinion_(id);
  if (isTrue_(o.hidden)) throw userError_("この意見にはコメントできません。");
  const stance = oneOf_(req.stance, STANCES, "立場", req.stanceIdx);
  const body = text_(req.body, LIMITS.comment, "コメント", true);
  const comments = readTable_("comments");
  const cid = comments.reduce((m, c) => Math.max(m, Number(c.id) || 0), 0) + 1;
  appendRow_("comments", { id: cid, opinionId: id, stance: stance, body: body, postedAt: today_(), hidden: "" });
  appendLog_("（匿名）", "コメント投稿", "No." + id);
  return "コメントしました。";
}

// ===== 運営の操作 =====

function setDeadline_(req) {
  const o = findOpinion_(Number(req.id));
  const d = String(req.deadline || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw userError_("期限の日付が正しくありません。");
  if (o.status !== STATUS.open) throw userError_("投票中の意見だけ期限を変更できます。");
  if (o.yes !== "") throw userError_("締切済みの意見は期限を変更できません（集計が確定しています）。");
  setCells_("opinions", o._row, { deadline: d });
  appendLog_("運営", "投票期限変更", "No." + o.id + " → " + d);
  return "投票期限を " + d + " に変更しました。";
}

function setHidden_(req) {
  const hidden = req.hidden ? "○" : "";
  if (req.type === "comment") {
    const c = readTable_("comments").find(x => Number(x.id) === Number(req.id));
    if (!c) throw userError_("コメントが見つかりません。");
    setCells_("comments", c._row, { hidden: hidden });
    appendLog_("運営", req.hidden ? "コメント非表示" : "コメント再表示", "コメントID " + c.id + "（意見No." + c.opinionId + "）");
  } else {
    const o = findOpinion_(Number(req.id));
    setCells_("opinions", o._row, { hidden: hidden });
    appendLog_("運営", req.hidden ? "意見非表示" : "意見再表示", "No." + o.id);
  }
  return req.hidden ? "非表示にしました。" : "再表示しました。";
}

// 見送り・実施済などの状態変更
function setStatus_(req) {
  const o = findOpinion_(Number(req.id));
  const to = req.status;
  if (to === STATUS.skipped) {
    if (o.status !== STATUS.open) throw userError_("投票中・締切の意見だけ見送りにできます。");
    setCells_("opinions", o._row, { status: STATUS.skipped });
    finalizeOne_(findOpinion_(o.id));
  } else if (to === STATUS.done) {
    if (o.status !== STATUS.answered) throw userError_("会社の回答が済んだ意見だけ実施済にできます。");
    setCells_("opinions", o._row, { status: STATUS.done, doneAt: today_() });
  } else if (to === STATUS.answered) {
    if (o.status !== STATUS.done) throw userError_("この変更はできません。");
    setCells_("opinions", o._row, { status: STATUS.answered, doneAt: "" });
  } else {
    throw userError_("この変更はできません。");
  }
  appendLog_("運営", "状態変更", "No." + o.id + " → " + to);
  return "「" + to + "」にしました。";
}

// 意見書として提出：状態を提出済にし、集計を確定させる。
function submit_(req) {
  const ids = (req.ids || []).map(Number).filter(Boolean);
  if (!ids.length) throw userError_("提出する意見を選んでください。");
  const days = Number(getSetting_("回答期限（日）")) || 30;
  const today = today_();
  ids.forEach(id => {
    const o = findOpinion_(id);
    if (o.status !== STATUS.open && o.status !== STATUS.skipped) throw userError_("No." + id + " はすでに提出済みです。");
    setCells_("opinions", o._row, { status: STATUS.submitted, submittedAt: today, answerDue: addDays_(today, days) });
    finalizeOne_(findOpinion_(id));
  });
  appendLog_("運営", "意見書提出", ids.map(i => "No." + i).join("、"));
  return ids.length + " 件を提出済にしました。";
}

// ===== 会社の操作 =====

function answer_(req) {
  const o = findOpinion_(Number(req.id));
  if ([STATUS.submitted, STATUS.answered].indexOf(o.status) < 0) throw userError_("提出済みの意見だけ回答できます。");
  const result = oneOf_(req.result, RESULTS, "回答", req.resultIdx);
  const reason = text_(req.reason, LIMITS.reason, "理由", true);
  const planDate = String(req.planDate || "");
  if (planDate && !/^\d{4}-\d{2}-\d{2}$/.test(planDate)) throw userError_("実施予定日が正しくありません。");
  setCells_("opinions", o._row, { status: STATUS.answered, result: result, reason: reason, answeredAt: today_(), planDate: planDate });
  appendLog_("会社", "回答", "No." + o.id + " " + result);
  return "回答を登録しました。";
}

// ===== シート読み書き =====

function ensureSheets_() {
  const ss = SpreadsheetApp.getActive();
  // 旧版（投票コード方式）の社員名簿は「社員名簿（旧）」に退避して作り直す
  const old = ss.getSheetByName(SHEETS.employees);
  if (old && old.getRange(1, 5).getValue() === "投票コード（ハッシュ）") old.setName(SHEETS.employees + "（旧）");
  Object.keys(SHEETS).forEach(key => {
    let sh = ss.getSheetByName(SHEETS[key]);
    if (sh) return;
    sh = ss.insertSheet(SHEETS[key]);
    const headers = COLS[key].map(c => c[1]);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#e8eef7");
    sh.setFrozenRows(1);
    if (key === "settings") {
      sh.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    }
    if (key === "employees") {
      sh.hideColumns(colIndex_("employees", "pwHash"), 2); // ハッシュ・ソルト列は隠す
    }
  });
}

function readTable_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS[key]);
  const cols = COLS[key];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, cols.length).getValues();
  return values.map((r, i) => {
    const o = { _row: i + 2 };
    cols.forEach((c, j) => { o[c[0]] = cellStr_(r[j]); });
    return o;
  });
}

function appendRow_(key, obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS[key]);
  sh.appendRow(COLS[key].map(c => toCell_(obj[c[0]])));
}

function setCells_(key, row, obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS[key]);
  Object.keys(obj).forEach(k => {
    sh.getRange(row, colIndex_(key, k)).setValue(toCell_(obj[k]));
  });
}

function colIndex_(key, name) {
  const i = COLS[key].findIndex(c => c[0] === name);
  if (i < 0) throw new Error("unknown column " + name);
  return i + 1;
}

function findOpinion_(id) {
  const o = readTable_("opinions").find(x => Number(x.id) === Number(id));
  if (!o) throw userError_("意見が見つかりません。");
  return o;
}

// 文字は先頭に ' を付けて書き込む（日付・数式に勝手に変換されないように。"=..." の数式実行も防ぐ）
function toCell_(v) {
  if (v === undefined || v === null || v === "") return "";
  if (typeof v === "number") return v;
  return "'" + String(v);
}

function txt_(v) { return toCell_(v); }

function cellStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  return v === null || v === undefined ? "" : String(v);
}

function appendLog_(actor, action, detail) {
  appendRow_("log", { at: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm"), actor: actor, action: action, detail: detail });
}

function getSetting_(name) {
  const s = readTable_("settings").find(r => r.key === name);
  return s ? s.value : "";
}

function publicSettings_() {
  const o = {};
  readTable_("settings").forEach(r => { if (r.key) o[r.key] = r.value; });
  return o;
}

// ===== ユーティリティ =====

function readTally_(props, id) {
  const raw = props.getProperty("T_" + id);
  if (!raw) return { y: 0, n: 0 };
  try { const t = JSON.parse(raw); return { y: Number(t.y) || 0, n: Number(t.n) || 0 }; }
  catch (e) { return { y: 0, n: 0 }; }
}

function text_(v, max, label, required) {
  const s = String(v === undefined || v === null ? "" : v).replace(/\r\n?/g, "\n").trim();
  if (required && !s) throw userError_(label + "を入力してください。");
  if (s.length > max) throw userError_(label + "は" + max + "文字以内で入力してください。");
  return s;
}

// 選択肢は「何番目か」（idx）で受け取る。名前で届いた場合も、空白や全角半角の違いを無視して照合する。
function oneOf_(v, list, label, idx) {
  const i = Number(idx);
  if (idx !== undefined && idx !== null && idx !== "" && Number.isInteger(i) && i >= 0 && i < list.length) return list[i];
  const key = normName_(v);
  const hit = list.find(x => normName_(x) === key);
  if (hit) return hit;
  throw userError_(label + "を選んでください。（受け取った値：" + String(v === undefined ? "なし" : v).slice(0, 30) + "）");
}

function isActive_(e) { return String(e.active).trim() !== "×"; }
function isTrue_(v) { return String(v).trim() !== "" && String(v).trim() !== "FALSE"; }

function today_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd"); }

function addDays_(ymd, days) {
  const p = ymd.split("-").map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + days));
  return Utilities.formatDate(d, "UTC", "yyyy-MM-dd");
}

function randomCode_() {
  // UUID（乱数）から8文字のコードを作る
  const hex = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  let code = "";
  for (let i = 0; i < 8; i++) code += CODE_CHARS[parseInt(hex.substr(i * 4, 4), 16) % CODE_CHARS.length];
  return code;
}

function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

function userError_(msg, code) { const e = new Error(msg); e.userMessage = msg; e.code = code; return e; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
