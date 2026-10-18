
/* =====================================================
   OC CHECK-IN
   GitHub Pages + Google Sheets API
   Google Apps Script (Code.gs) 不使用
===================================================== */

// Google Cloud Consoleで作成したOAuthクライアントIDに変更する
const GOOGLE_CLIENT_ID = "974919799974-d866j4db02d8ifum919h51det026maf7.apps.googleusercontent.com";

const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const DISCOVERY_DOC =
  "https://sheets.googleapis.com/$discovery/rest?version=v4";

const $ = (id) => document.getElementById(id);

let tokenClient = null;
let apiReady = false;
let accessToken = null;
let participants = [];
let currentSpreadsheetId = "";
let currentSheetTitle = "";
let headerMap = {};
let statusColumnIndex = -1;
let toastTimer = null;

const FIELD_ALIASES = {
  school: [
    "中学校名", "中学校", "学校名", "在籍中学校"
  ],
  grade: [
    "学年", "中学校の学年", "現在の学年"
  ],
  name: [
    "中学生のお名前", "お名前", "氏名", "名前", "参加者氏名"
  ],
  furigana: [
    "フリガナ", "ふりがな", "フリガナ（全角）", "氏名フリガナ"
  ],
  attendance: [
    "オーキャン参加回数", "OC参加回数", "参加回数",
    "オープンキャンパス参加回数", "何回目"
  ],
  session: [
    "第1部or第2部", "第1部／第2部", "参加時間",
    "参加部", "希望時間", "第1部・第2部"
  ],
  lesson: [
    "体験授業", "体験授業希望", "希望する体験授業",
    "体験授業の希望"
  ],
  club: [
    "クラブ体験等参加希望", "クラブ体験", "部活動体験",
    "クラブ体験等", "参加希望"
  ],
  status: [
    "来校状況"
  ]
};

function normalize(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/[\s　・（）()「」【】［］]/g, "")
    .toLowerCase();
}

function showToast(message) {
  const toast = $("toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("show"), 3200);
}

function setMessage(message, type = "") {
  const el = $("sheetStatus");
  el.textContent = message;
  el.className = "message" + (type ? ` ${type}` : "");
}

function setLoginState(loggedIn, email = "") {
  $("loginBtn").classList.toggle("hidden", loggedIn);
  $("logoutBtn").classList.toggle("hidden", !loggedIn);
  $("loginStatus").textContent = loggedIn
    ? `${email || "Googleログイン済み"}`
    : "未ログイン";
}

function getSpreadsheetId(url) {
  const match = String(url).trim().match(
    /docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/
  );
  if (!match) {
    throw new Error("スプレッドシートのURLを確認してください。");
  }
  return match[1];
}

function quoteSheetTitle(title) {
  return "'" + title.replace(/'/g, "''") + "'";
}

async function initializeGoogle() {
  try {
    await new Promise((resolve, reject) => {
      const started = Date.now();

      function check() {
        if (window.gapi && window.google?.accounts?.oauth2) {
          resolve();
        } else if (Date.now() - started > 20000) {
          reject(new Error("Google認証ライブラリを読み込めませんでした。"));
        } else {
          setTimeout(check, 200);
        }
      }
      check();
    });

    await new Promise((resolve) => gapi.load("client", resolve));

    await gapi.client.init({
      discoveryDocs: [DISCOVERY_DOC]
    });

    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: SHEETS_SCOPE,
      callback: async (response) => {
        if (response.error) {
          setMessage("Googleログインに失敗しました。", "error");
          return;
        }

        accessToken = response.access_token;
        gapi.client.setToken({ access_token: accessToken });
        apiReady = true;

        setLoginState(true);
        showToast("Googleに接続しました。");

        const savedUrl = $("sheetUrl").value.trim();
        if (savedUrl) {
          await loadSpreadsheet();
        }
      }
    });

    if (
      !GOOGLE_CLIENT_ID ||
      GOOGLE_CLIENT_ID.startsWith("ここに")
    ) {
      $("loginBtn").disabled = true;
      setMessage(
        "Google CloudでOAuthクライアントIDを作成し、script.jsに設定してください。"
      );
      return;
    }

    $("loginBtn").disabled = false;
    setMessage("Googleにログインして回答シートを読み込んでください。");
  } catch (error) {
    console.error(error);
    setMessage(error.message || "初期化に失敗しました。", "error");
  }
}

function login() {
  if (!tokenClient) {
    showToast("Google認証の準備中です。少し待ってください。");
    return;
  }

  tokenClient.requestAccessToken({
    prompt: accessToken ? "" : "consent"
  });
}

function logout() {
  if (accessToken && window.google?.accounts?.oauth2) {
    google.accounts.oauth2.revoke(accessToken);
  }

  accessToken = null;
  apiReady = false;
  participants = [];
  currentSpreadsheetId = "";
  currentSheetTitle = "";
  headerMap = {};
  statusColumnIndex = -1;

  if (window.gapi?.client) {
    gapi.client.setToken(null);
  }

  setLoginState(false);
  renderResults();
  updateStats();
  setMessage("ログアウトしました。", "success");
}

function findHeaderIndex(headers, aliases) {
  const normalizedHeaders = headers.map(normalize);

  // 完全一致を優先
  for (const alias of aliases) {
    const index = normalizedHeaders.indexOf(normalize(alias));
    if (index !== -1) return index;
  }

  // 一部一致も試す
  for (const alias of aliases) {
    const target = normalize(alias);
    if (target.length < 3) continue;

    const index = normalizedHeaders.findIndex(
      (header) => header.includes(target) || target.includes(header)
    );

    if (index !== -1) return index;
  }

  return -1;
}

function resolveHeaders(headers) {
  const result = {};

  for (const [field, aliases] of Object.entries(FIELD_ALIASES)) {
    result[field] = findHeaderIndex(headers, aliases);
  }

  return result;
}

async function loadSpreadsheet() {
  if (!apiReady || !accessToken) {
    showToast("先にGoogleでログインしてください。");
    return;
  }

  const url = $("sheetUrl").value.trim();

  if (!url) {
    setMessage("スプレッドシートURLを入力してください。", "error");
    return;
  }

  const button = $("loadSheetBtn");
  button.disabled = true;
  button.textContent = "読み込み中…";
  setMessage("回答データを読み込んでいます…");

  try {
    const spreadsheetId = getSpreadsheetId(url);

    // スプレッドシートのタブ一覧を取得
    const metadata = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId,
      fields: "spreadsheetId,properties.title,sheets.properties"
    });

    const sheets = metadata.result.sheets || [];
    if (!sheets.length) {
      throw new Error("スプレッドシートにシートがありません。");
    }

    // 「フォームの回答」が含まれるタブを優先
    const selectedSheet =
      sheets.find((s) =>
        /フォームの回答|form responses/i.test(s.properties.title)
      ) || sheets[0];

    const sheetTitle = selectedSheet.properties.title;
    const range = `${quoteSheetTitle(sheetTitle)}!A1:ZZ`;

    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId,
      range
    });

    const rows = response.result.values || [];

    if (rows.length < 1) {
      throw new Error("シートに回答データがありません。");
    }

    const headers = rows[0].map((value) => String(value ?? "").trim());
    headerMap = resolveHeaders(headers);

    if (headerMap.name === -1) {
      throw new Error(
        "名前の列が見つかりません。1行目の質問名を確認してください。"
      );
    }

    if (headerMap.status === -1) {
      throw new Error(
        "回答シートの1行目の末尾に「来校状況」という列見出しを追加してください。その後、再読み込みしてください。"
      );
    }

    statusColumnIndex = headerMap.status;

    participants = rows.slice(1).map((row, index) => {
      const get = (field) => {
        const column = headerMap[field];
        return column >= 0 ? String(row[column] ?? "").trim() : "";
      };

      return {
        rowNumber: index + 2,
        school: get("school"),
        grade: get("grade"),
        name: get("name"),
        furigana: get("furigana"),
        attendance: get("attendance"),
        session: get("session"),
        lesson: get("lesson"),
        club: get("club"),
        status: get("status") || "未到校"
      };
    }).filter((person) => person.name);

    currentSpreadsheetId = spreadsheetId;
    currentSheetTitle = sheetTitle;

    $("eventName").value =
      $("eventName").value.trim() || metadata.result.properties.title || "";

    updateStats();
    renderResults();

    setMessage(
      `${sheetTitle} を読み込みました（${participants.length}名）。`,
      "success"
    );
    showToast("回答データを読み込みました。");
  } catch (error) {
    console.error(error);
    setMessage(
      error?.result?.error?.message || error.message ||
        "読み込みに失敗しました。",
      "error"
    );
  } finally {
    button.disabled = false;
    button.textContent = "回答を読み込む";
  }
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[char]);
}

function detail(label, value, full = false) {
  return `
    <div class="detail ${full ? "full" : ""}">
      <dt>${escapeHtml(label)}</dt>
      <dd>${escapeHtml(value || "未回答")}</dd>
    </div>`;
}

function renderCard(person, index) {
  const arrived = person.status === "来校済み";
  const waiting = !arrived;

  return `
    <article class="participant-card">
      <div class="participant-top">
        <div>
          <h3 class="participant-name">${escapeHtml(person.name)}</h3>
          <p class="furigana">${escapeHtml(person.furigana || "フリガナ未登録")}</p>
          <p class="school-name">${escapeHtml(person.school || "中学校名未登録")}</p>
        </div>
        <span class="session-badge">${escapeHtml(person.session || "部未設定")}</span>
      </div>

      <dl class="details-grid">
        ${detail("学年", person.grade)}
        ${detail("OC参加回数", person.attendance)}
        ${detail("体験授業", person.lesson, true)}
        ${detail("クラブ体験等参加希望", person.club, true)}
      </dl>

      <div class="arrival-row">
        <span class="arrival-label">
          現在の受付状況：<strong>${escapeHtml(person.status || "未到校")}</strong>
        </span>
        <div class="arrival-actions">
          <button
            class="arrival-btn ${waiting ? "selected-waiting" : ""}"
            data-index="${index}" data-status="未到校"
            ${waiting ? "disabled" : ""}>
            未到校
          </button>
          <button
            class="arrival-btn ${arrived ? "selected-arrived" : ""}"
            data-index="${index}" data-status="来校済み"
            ${arrived ? "disabled" : ""}>
            来校済み
          </button>
        </div>
      </div>
    </article>`;
}

function renderResults() {
  const query = normalize($("searchInput").value);
  const results = $("results");

  const matched = participants
    .map((person, index) => ({ person, index }))
    .filter(({ person }) => {
      if (!query) return true;

      const searchable = normalize([
        person.school,
        person.name,
        person.furigana
      ].join(" "));

      return searchable.includes(query);
    });

  $("resultCount").textContent = participants.length
    ? `${matched.length}名 / 全${participants.length}名`
    : "回答を読み込んでください";

  if (!participants.length) {
    results.innerHTML = `
      <div class="empty-state">
        <span class="empty-icon">⌕</span>
        <h3>回答データがありません</h3>
        <p>Googleログイン後に回答シートを読み込んでください。</p>
      </div>`;
    return;
  }

  if (!matched.length) {
    results.innerHTML = `
      <div class="empty-state">
        <span class="empty-icon">⌕</span>
        <h3>該当する参加者がいません</h3>
        <p>中学校名・名前・フリガナを確認してください。</p>
      </div>`;
    return;
  }

  results.innerHTML = matched
    .map(({ person, index }) => renderCard(person, index))
    .join("");
}

function updateStats() {
  const total = participants.length;
  const arrived = participants.filter(
    (person) => person.status === "来校済み"
  ).length;

  $("totalCount").textContent = total;
  $("arrivedCount").textContent = arrived;
  $("waitingCount").textContent = total - arrived;
}

async function updateArrival(index, status) {
  if (!apiReady || !accessToken) {
    showToast("Googleにログインしてください。");
    return;
  }

  const person = participants[index];
  if (!person || statusColumnIndex < 0) return;

  const buttons = document.querySelectorAll(
    `.arrival-btn[data-index="${index}"]`
  );
  buttons.forEach((button) => button.disabled = true);

  try {
    const columnLetter = indexToColumn(statusColumnIndex);
    const cellRange =
      `${quoteSheetTitle(currentSheetTitle)}!${columnLetter}${person.rowNumber}`;

    await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: currentSpreadsheetId,
      range: cellRange,
      valueInputOption: "RAW",
      resource: { values: [[status]] }
    });

    person.status = status;
    updateStats();
    renderResults();
    showToast(`${person.name} さん：${status} を保存しました。`);
  } catch (error) {
    console.error(error);
    buttons.forEach((button) => button.disabled = false);
    showToast(
      error?.result?.error?.message ||
      "保存できませんでした。編集権限を確認してください。"
    );
  }
}

function indexToColumn(index) {
  let n = index + 1;
  let column = "";

  while (n > 0) {
    const remainder = (n - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    n = Math.floor((n - 1) / 26);
  }

  return column;
}

function clearSettings() {
  $("eventName").value = "";
  $("sheetUrl").value = "";
  $("searchInput").value = "";

  participants = [];
  currentSpreadsheetId = "";
  currentSheetTitle = "";
  headerMap = {};
  statusColumnIndex = -1;

  updateStats();
  renderResults();
  setMessage("設定をクリアしました。");
}

$("loginBtn").addEventListener("click", login);
$("logoutBtn").addEventListener("click", logout);
$("loadSheetBtn").addEventListener("click", loadSpreadsheet);
$("refreshBtn").addEventListener("click", loadSpreadsheet);
$("clearSettingsBtn").addEventListener("click", clearSettings);

$("searchInput").addEventListener("input", renderResults);

$("clearSearchBtn").addEventListener("click", () => {
  $("searchInput").value = "";
  renderResults();
  $("searchInput").focus();
});

$("results").addEventListener("click", (event) => {
  const button = event.target.closest(".arrival-btn");
  if (!button) return;

  updateArrival(
    Number(button.dataset.index),
    button.dataset.status
  );
});

setLoginState(false);
updateStats();
renderResults();
initializeGoogle();
