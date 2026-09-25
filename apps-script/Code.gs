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
const SHEET_SCORES = "Scores";

const WEEK_COLUMNS = ["id", "title", "desc", "youtubeId", "formUrl", "opensAt", "closesAt"];

const DEFAULT_SITE = {
  pageTitle: "Долоо хоног бvрийн сургалтын бичлэг",
  pageIntro: "Бичлэгийг эхнээс нь дуустал vзсэний дараа мэдлэгийн сорилын холбоос нээгддэг. Зөвхөн Тоглуулах/Тvр зогсоох товч ашиглана, урагшлуулах боломжгvй.",
  loginTitle: "Тавтай морил",
  loginIntro: "ISO идэвхжvvлэлтийн аяны сургалтад орохын өмнө алба/нэгж, албан тушаал болон өөрийн нэрээ оруулна уу.",
  footerNote: "Асуудал гарвал ажлын байрны админтай холбогдоно уу.",
  siteActive: "true",
  inactiveMessage: "Энэ систем одоогоор идэвхгvй байна. Дараа дахин орж vзнэ vv.",
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
    // Google Sheets "true"/"false" текстийг автоматаар boolean төрөл болгож
    // хувиргадаг тул (r[1] || "") ашиглавал "false"(boolean) хоосон мөр болж
    // алдагддаг — үvнээс сэргийлж boolean утгыг эхлээд шалгана.
    if (key) {
      const v = r[1];
      site[key] = typeof v === "boolean" ? String(v) : String(v || "");
    }
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

// Нэрийг том/жижиг vсэг, зайн ялгаа vл хамааран харьцуулах түлхvvр болгож жигдэлнэ
// (жишээ нь "С.Мэнхvvл" vs "с.мэнхvvл" ижил хvн гэж танигдана).
function normalizeKey(s) {
  return String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
}

/* ---------- Progress (event log) ---------- */
const PROGRESS_COLUMNS = ["timestamp", "org", "position", "name", "weekId", "event"];

function resetProgress() {
  const sheet = getOrCreateSheet(SHEET_PROGRESS, PROGRESS_COLUMNS);
  sheet.clearContents();
  sheet.appendRow(PROGRESS_COLUMNS);
}

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

/* ---------- Scores (Microsoft Forms Quiz-с Excel-ээр оруулсан оноо) ---------- */
const SCORE_COLUMNS = ["weekId", "name", "score"];

function readScores() {
  const sheet = getOrCreateSheet(SHEET_SCORES, SCORE_COLUMNS);
  const values = sheet.getDataRange().getValues();
  return values
    .slice(1)
    .filter((r) => r[1])
    .map((r) => ({ weekId: String(r[0] || ""), name: String(r[1] || ""), score: Number(r[2]) || 0 }));
}

// Тухайн сургалтын өмнөх онооны мөрvvдийг арилгаад, шинээр оруулсан Excel-ийн
// өгөгдлөөр сольж бичнэ (давхар оруулбал давхардахгvй байхын тулд).
function writeScoresForWeek(weekId, rows) {
  const sheet = getOrCreateSheet(SHEET_SCORES, SCORE_COLUMNS);
  const values = sheet.getDataRange().getValues();
  const keep = values.slice(1).filter((r) => String(r[0] || "") !== String(weekId) && r[1]);

  // Нэг хvн Quiz-ийг хэд хэдэн удаа бөглөсөн тохиолдолд (Microsoft Forms дахин
  // бөглөхийг хориглодоггvй тул Excel-д нэг хvн хэд хэдэн мөрөнд орж болно) —
  // онооныг нэмэлгvй, зөвхөн хамгийн өндөр оноог нь тооцно. Ингэхгvй бол
  // тухайн хvний нийт оноо мөр давхарласан тоогоор хэд дахин нэмэгдэж гарна.
  const dedup = {}; // normalizedName -> {name, score}
  rows.forEach((r) => {
    const name = String(r.name || "").trim();
    if (!name) return;
    const key = normalizeKey(name);
    const score = Number(r.score) || 0;
    if (!dedup[key] || score > dedup[key].score) {
      dedup[key] = { name: name, score: score };
    }
  });
  const dedupedRows = Object.keys(dedup).map((key) => dedup[key]);
  const duplicatesRemoved = rows.filter((r) => String(r.name || "").trim()).length - dedupedRows.length;

  sheet.clearContents();
  sheet.appendRow(SCORE_COLUMNS);
  keep.forEach((r) => sheet.appendRow(r));
  dedupedRows.forEach((r) => sheet.appendRow([String(weekId), r.name, r.score]));

  return { savedCount: dedupedRows.length, duplicatesRemoved: duplicatesRemoved };
}

/* ---------- Gemini AI (чөлөөт бичвэрийн хариулт vнэлэх) ---------- */
// Google AI Studio-с (aistudio.google.com) vнэгvй авсан API key-г
// Project Settings → Script Properties → GEMINI_API_KEY нэрээр хадгална.
const GEMINI_MODEL = "gemini-2.0-flash";

function callGemini(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("missing_gemini_key");
  const url = "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL + ":generateContent?key=" + apiKey;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json" },
  };
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code !== 200) throw new Error("gemini_http_" + code + ": " + res.getContentText().slice(0, 300));
  const data = JSON.parse(res.getContentText());
  const text = data.candidates &&
    data.candidates[0] &&
    data.candidates[0].content &&
    data.candidates[0].content.parts &&
    data.candidates[0].content.parts[0] &&
    data.candidates[0].content.parts[0].text;
  if (!text) throw new Error("gemini_empty_response");
  return text;
}

// answers: [{name, text}] — нэг Gemini дуудлагад хэт олон хvнийг оруулбал
// хариу урт болж алдаа гарах эрсдэлтэй тул дуудагч тал BATCH_SIZE-аар хуваана.
function gradeTextAnswersBatch(question, maxPoints, rubric, answers) {
  const prompt =
    "Чи Монгол хэл дээрх ажилтны сургалтын шалгалтын чөлөөт бичвэрийн хариултуудыг vнэлж буй туслах.\n" +
    'Асуулт: "' + question + '"\n' +
    "Дээд оноо (асуулт тус бvрд): " + maxPoints + "\n" +
    (rubric ? "Оноо өгөх шалгуур: " + rubric + "\n" : "") +
    "Доорх хvн бvрийн хариултыг уншиж, 0-ээс " + maxPoints + " хvртэлх бvхэл тоон оноо өг. " +
    "Хариулаагvй эсвэл огт хамааралгvй бол 0 өг. " +
    "Зөвхөн доорх форматтай JSON массив буцаа, өөр vг нэмэхгvй:\n" +
    '[{"name": "...", "score": 0, "reason": "богино (10-15 vгтэй) тайлбар"}, ...]\n\n' +
    "Хариултууд:\n" +
    answers.map((a, i) => (i + 1) + ". Нэр: " + a.name + " | Хариулт: " + (a.text || "(хариулаагvй)")).join("\n");

  const raw = callGemini(prompt);
  let cleaned = raw.trim().replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/```\s*$/i, "");
  const parsed = JSON.parse(cleaned);
  if (!Array.isArray(parsed)) throw new Error("gemini_bad_format");
  return parsed.map((r) => ({
    name: String(r.name || ""),
    score: Math.max(0, Math.min(maxPoints, Math.round(Number(r.score) || 0))),
    reason: String(r.reason || ""),
  }));
}

function gradeTextAnswers(question, maxPoints, rubric, answers) {
  const BATCH_SIZE = 25;
  let results = [];
  for (let i = 0; i < answers.length; i += BATCH_SIZE) {
    const chunk = answers.slice(i, i + BATCH_SIZE);
    results = results.concat(gradeTextAnswersBatch(question, maxPoints, rubric, chunk));
  }
  return results;
}

function readProgressSummary() {
  const sheet = getOrCreateSheet(SHEET_PROGRESS, PROGRESS_COLUMNS);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter((r) => r[1] || r[3]);

  const weeks = readWeeks();
  const weekIds = weeks.map((w) => w.id);

  // Хvн бvрийг (алба+тушаал+нэр) нэгтгэж, тэдний бvртгvvлсэн болон дуусгасан
  // сургалтуудыг цуглуулна. Нэрээ өөр өөр vед том/жижиг vсэг, зайгаар өөрөөр
  // бичсэн ч (жишээ нь "С.Мэнхvvл" vs "с.мэнхvvл") ижил хvн гэж танихын тулд
  // харьцуулах түлхvvрийг жигдэлж (normalize) vvсгэнэ — харин анх бvртгэгдсэн
  // бичлэгийн жинхэнэ хэлбэрийг харуулахдаа хэвээр vлдээнэ.
  const people = {}; // key -> {org, position, name, completed:Set}
  rows.forEach((r) => {
    const org = String(r[1] || "(тодорхойгvй)");
    const position = String(r[2] || "");
    const name = String(r[3] || "(тодорхойгvй)");
    const weekId = String(r[4] || "");
    const event = String(r[5] || "");
    const key = normalizeKey(org) + "||" + normalizeKey(position) + "||" + normalizeKey(name);
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

  const peopleSorted = peopleList.slice().sort((a, b) => {
    if (a.org !== b.org) return a.org < b.org ? -1 : 1;
    return a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
  });

  // Онооны тэргvvлэгчид: Excel-ээр оруулсан оноог (Scores) нэр тус бvрээр нэгтгэж,
  // боломжтой бол бvртгэлтэй хvний алба/тушаалтай тааруулна (нэрээр normalize хийж).
  const peopleByName = {};
  peopleList.forEach((p) => {
    const k = normalizeKey(p.name);
    if (k && !peopleByName[k]) peopleByName[k] = p;
  });
  const scoreTotals = {}; // normalizedName -> {name, total}
  readScores().forEach((r) => {
    const key = normalizeKey(r.name);
    if (!key) return;
    if (!scoreTotals[key]) scoreTotals[key] = { name: r.name, total: 0 };
    scoreTotals[key].total += r.score;
  });
  const leaderboard = Object.keys(scoreTotals)
    .map((key) => {
      const matched = peopleByName[key];
      return {
        name: matched ? matched.name : scoreTotals[key].name,
        org: matched ? matched.org : "",
        position: matched ? matched.position : "",
        score: scoreTotals[key].total,
        matched: !!matched,
      };
    })
    .sort((a, b) => b.score - a.score);

  return {
    weekIds: weekIds,
    weekTitles: weeks.map((w) => w.title),
    totalParticipants: peopleList.length,
    perWeek: perWeek,
    byOrg: Object.keys(byOrg).map((k) => byOrg[k]),
    people: peopleSorted.map((p) => {
      const scoreKey = normalizeKey(p.name);
      return {
        org: p.org,
        position: p.position,
        name: p.name,
        completed: p.completed,
        score: (scoreTotals[scoreKey] && scoreTotals[scoreKey].total) || 0,
      };
    }),
    leaderboard: leaderboard,
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

// Хоосон биш буцаах утга нь тохирохгvй шалтгааныг илэрхийлнэ.
function validatePasswordStrength(pw) {
  if (pw.length < 8) return "Нууц vг дор хаяж 8 тэмдэгт байх ёстой.";
  if (!/[a-zA-Z]/.test(pw)) return "Нууц vг наад зах нь нэг vсэг агуулсан байх ёстой.";
  if (!/[0-9]/.test(pw)) return "Нууц vг наад зах нь нэг тоо агуулсан байх ёстой.";
  return null;
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

  if (action === "changePassword") {
    const newPassword = String(body.newPassword || "");
    const weak = validatePasswordStrength(newPassword);
    if (weak) return jsonResponse({ ok: false, error: "weak_password", message: weak });
    PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", newPassword);
    return jsonResponse({ ok: true });
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

  if (action === "saveScores") {
    if (!body.weekId || !Array.isArray(body.scores)) return jsonResponse({ ok: false, error: "expected_weekid_and_scores" });
    const result = writeScoresForWeek(body.weekId, body.scores);
    return jsonResponse({ ok: true, savedCount: result.savedCount, duplicatesRemoved: result.duplicatesRemoved });
  }

  if (action === "gradeTextAnswers") {
    if (!body.question || !Array.isArray(body.answers)) return jsonResponse({ ok: false, error: "expected_question_and_answers" });
    try {
      const maxPoints = Number(body.maxPoints) || 1;
      const rubric = String(body.rubric || "");
      const results = gradeTextAnswers(String(body.question), maxPoints, rubric, body.answers);
      return jsonResponse({ ok: true, results: results });
    } catch (err) {
      const msg = String(err);
      if (msg.indexOf("missing_gemini_key") !== -1) return jsonResponse({ ok: false, error: "missing_gemini_key" });
      return jsonResponse({ ok: false, error: "gemini_failed", message: msg });
    }
  }

  if (action === "getSummary") {
    return jsonResponse({ ok: true, summary: readProgressSummary() });
  }

  if (action === "resetProgress") {
    resetProgress();
    return jsonResponse({ ok: true });
  }

  return jsonResponse({ ok: false, error: "unknown_action" });
}
