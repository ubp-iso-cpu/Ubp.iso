/**
 * ISO Аян — vнэгvй backend (Google Apps Script).
 * Энэ файлыг Google Sheet-ийн Extensions → Apps Script дотор бvтнээр нь
 * хуулж тавина. Дараа нь Deploy → New deployment → Web app хийж URL авна.
 *
 * Тохиргоо: Project Settings → Script Properties → ADMIN_PASSWORD нэмнэ.
 */

const SHEET_WEEKS = "Weeks";
const SHEET_CONFIG = "Config";
const SHEET_ORG = "OrgStructure";
const SHEET_PROGRESS = "Progress";

const WEEK_COLUMNS = ["id", "title", "desc", "youtubeId", "formUrl", "opensAt", "closesAt"];

const DEFAULT_SITE = {
  pageTitle: "Долоо хоног бvрийн сургалтын бичлэг",
  pageIntro: "Бичлэгийг эхнээс нь дуустал vзсэний дараа мэдлэгийн сорилын холбоос нээгддэг. Зөвхөн Тоглуулах/Тvр зогсоох товч ашиглана, урагшлуулах боломжгvй.",
  loginTitle: "Тавтай морил",
  loginIntro: "ISO идэвхжvvлэлтийн аяны сургалтад орохын өмнө алба/нэгж, албан тушаал болон өөрийн нэрээ оруулна уу.",
  footerNote: "Асуудал гарвал ажлын байрны админтай холбогдоно уу.",
};

function getSs() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name, headerRow) {
  const ss = getSs();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headerRow) sheet.appendRow(headerRow);
  }
  return sheet;
}

/* ---------- Weeks ---------- */
function readWeeks() {
  const sheet = getOrCreateSheet(SHEET_WEEKS, WEEK_COLUMNS);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);
  return rows
    .filter((r) => r[0])
    .map((r) => ({
      id: String(r[0]),
      title: String(r[1] || ""),
      desc: String(r[2] || ""),
      youtubeId: String(r[3] || ""),
      formUrl: String(r[4] || ""),
      opensAt: formatDateCell(r[5]),
      closesAt: formatDateCell(r[6]),
    }));
}

function writeWeeks(weeks) {
  const sheet = getOrCreateSheet(SHEET_WEEKS, WEEK_COLUMNS);
  sheet.clearContents();
  sheet.appendRow(WEEK_COLUMNS);
  weeks.forEach((w) => {
    sheet.appendRow([
      w.id || "",
      w.title || "",
      w.desc || "",
      extractYoutubeId(w.youtubeId),
      w.formUrl || "",
      w.opensAt || "",
      w.closesAt || "",
    ]);
  });
}

// Клиент (хуучин tab, гар аргаар илгээсэн хvсэлт г.м.) ямар ч хэлбэрээр
// youtubeId илгээсэн байсан ч энд сервер талд заавал цэвэрлэнэ — ингэснээр
// хуучирсан client код ажиллуулж байгаа хэн ч буруу өгөгдөл бичиж чадахгvй.
function extractYoutubeId(input) {
  const s = String(input || "").trim();
  if (!s) return "";
  if (/^[a-zA-Z0-9_-]{11}$/.test(s)) return s;
  let m = s.match(/[?&]v=([a-zA-Z0-9_-]{11})/);
  if (m) return m[1];
  m = s.match(/youtu\.be\/([a-zA-Z0-9_-]{11})/);
  if (m) return m[1];
  m = s.match(/\/embed\/([a-zA-Z0-9_-]{11})/);
  if (m) return m[1];
  m = s.match(/[a-zA-Z0-9_-]{11}/);
  return m ? m[0] : s;
}

function formatDateCell(v) {
  if (!v) return "";
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

/* ---------- Site text (key/value) ---------- */
function readSite() {
  const sheet = getOrCreateSheet(SHEET_CONFIG, ["key", "value"]);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);
  const site = Object.assign({}, DEFAULT_SITE);
  rows.forEach((r) => {
    const key = String(r[0] || "");
    if (key) site[key] = String(r[1] || "");
  });
  return site;
}

function writeSite(site) {
  const sheet = getOrCreateSheet(SHEET_CONFIG, ["key", "value"]);
  sheet.clearContents();
  sheet.appendRow(["key", "value"]);
  Object.keys(site).forEach((key) => {
    sheet.appendRow([key, site[key] || ""]);
  });
}

/* ---------- Org structure (алба нэгж <-> албан тушаал, мөр бvр нэг хос) ---------- */
function readOrg() {
  const sheet = getOrCreateSheet(SHEET_ORG, ["department", "position"]);
  const values = sheet.getDataRange().getValues();
  return values
    .slice(1)
    .map((r) => ({
      department: String(r[0] || "").trim(),
      position: String(r[1] || "").trim(),
    }))
    .filter((r) => r.department);
}

function writeOrg(rows) {
  const sheet = getOrCreateSheet(SHEET_ORG, ["department", "position"]);
  sheet.clearContents();
  sheet.appendRow(["department", "position"]);
  rows.forEach((r) => {
    const dept = String((r.department != null ? r.department : "")).trim();
    const pos = String((r.position != null ? r.position : "")).trim();
    if (dept) sheet.appendRow([dept, pos]);
  });
}

/* ---------- Progress (event log) ---------- */
const PROGRESS_COLUMNS = ["timestamp", "org", "position", "name", "weekId", "event"];

function logProgress(entry) {
  const sheet = getOrCreateSheet(SHEET_PROGRESS, PROGRESS_COLUMNS);
  sheet.appendRow([
    new Date().toISOString(),
    String(entry.org || ""),
    String(entry.position || ""),
    String(entry.name || ""),
    String(entry.weekId || ""),
    String(entry.event || ""),
  ]);
}

function readProgressSummary() {
  const sheet = getOrCreateSheet(SHEET_PROGRESS, PROGRESS_COLUMNS);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter((r) => r[1] || r[3]);

  const weeks = readWeeks();
  const weekIds = weeks.map((w) => w.id);

  // Хvн бvрийг (алба+тушаал+нэр) нэгтгэж, тэдний бvртгvvлсэн болон дуусгасан
  // сургалтуудыг цуглуулна.
  const people = {}; // key -> {org, position, name, completed:Set}
  rows.forEach((r) => {
    const org = String(r[1] || "(тодорхойгvй)");
    const position = String(r[2] || "");
    const name = String(r[3] || "(тодорхойгvй)");
    const weekId = String(r[4] || "");
    const event = String(r[5] || "");
    const key = org + "||" + position + "||" + name;
    if (!people[key]) {
      people[key] = { org: org, position: position, name: name, completed: {} };
    }
    if (event === "video_completed" && weekId) {
      people[key].completed[weekId] = true;
    }
  });

  const peopleList = Object.keys(people).map((k) => people[k]);

  // Сургалт тус бvрийн дуусгасан тооcholder
  const perWeek = {};
  weekIds.forEach((id) => { perWeek[id] = 0; });
  peopleList.forEach((p) => {
    weekIds.forEach((id) => {
      if (p.completed[id]) perWeek[id] += 1;
    });
  });

  // Алба нэгж тус бvрээр
  const byOrg = {};
  peopleList.forEach((p) => {
    if (!byOrg[p.org]) {
      byOrg[p.org] = { org: p.org, participants: 0, perWeek: {} };
      weekIds.forEach((id) => { byOrg[p.org].perWeek[id] = 0; });
    }
    byOrg[p.org].participants += 1;
    weekIds.forEach((id) => {
      if (p.completed[id]) byOrg[p.org].perWeek[id] += 1;
    });
  });

  return {
    weekIds: weekIds,
    weekTitles: weeks.map((w) => w.title),
    totalParticipants: peopleList.length,
    perWeek: perWeek,
    byOrg: Object.keys(byOrg).map((k) => byOrg[k]),
  };
}

/* ---------- HTTP ---------- */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function checkPassword(body) {
  const adminPassword = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  return !!adminPassword && body.password === adminPassword;
}

function doGet(e) {
  return jsonResponse({
    weeks: readWeeks(),
    site: readSite(),
    org: readOrg(),
  });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ ok: false, error: "invalid_json" });
  }

  const action = body.action || "saveWeeks";

  // Хэрэглэгчийн явцын бvртгэл — нууц vг шаардахгvй, олон нийтэд нээлттэй.
  if (action === "logProgress") {
    if (!body.name || !body.event) {
      return jsonResponse({ ok: false, error: "missing_fields" });
    }
    logProgress(body);
    return jsonResponse({ ok: true });
  }

  // Vлдсэн бvх vйлдэл админ нууц vг шаардана.
  if (!checkPassword(body)) {
    return jsonResponse({ ok: false, error: "unauthorized" });
  }

  if (action === "saveWeeks") {
    if (!Array.isArray(body.weeks)) return jsonResponse({ ok: false, error: "expected_weeks_array" });
    writeWeeks(body.weeks);
    return jsonResponse({ ok: true });
  }

  if (action === "saveSite") {
    if (!body.site || typeof body.site !== "object") return jsonResponse({ ok: false, error: "expected_site_object" });
    writeSite(body.site);
    return jsonResponse({ ok: true });
  }

  if (action === "saveOrg") {
    if (!Array.isArray(body.org)) return jsonResponse({ ok: false, error: "expected_org_array" });
    writeOrg(body.org);
    return jsonResponse({ ok: true });
  }

  if (action === "getSummary") {
    return jsonResponse({ ok: true, summary: readProgressSummary() });
  }

  return jsonResponse({ ok: false, error: "unknown_action" });
}
