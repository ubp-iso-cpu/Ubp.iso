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

// "=", "+", "-", "@"-оор эхэлсэн текстийг Sheets томьёо гэж тайлбарладаг (жишээ нь
// нийтийн logProgress-оор ирсэн =IMPORTXML(...) өгөгдлийг гадагш алдуулна). Урд нь "'"
// залгаж зөвхөн текст болгоно — уншихад "'" харагдахгvй.
function safeCell(v) {
  const s = String(v == null ? "" : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// Мөр бvрт appendRow хийхийн оронд нэг setValues-ээр бичнэ: хурдан бөгөөд эхлээд бичээд
// дараа нь илvvдсэн мөрийг цэвэрлэдэг тул алдаа гарвал хуучин өгөгдөл бvтэн vлдэнэ.
function replaceSheetRows(sheet, header, rows) {
  const all = [header].concat(rows);
  sheet.getRange(1, 1, all.length, header.length).setValues(all);
  const last = sheet.getLastRow();
  if (last > all.length) {
    sheet.getRange(all.length + 1, 1, last - all.length, Math.max(header.length, sheet.getLastColumn())).clearContent();
  }
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
  replaceSheetRows(sheet, WEEK_COLUMNS, weeks.map((w) => [
    safeCell(w.id),
    safeCell(w.title),
    safeCell(w.desc),
    extractYoutubeId(w.youtubeId),
    safeCell(w.formUrl),
    safeCell(w.opensAt),
    safeCell(w.closesAt),
  ]));
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
  // ID олдохгvй бол оролтыг хэвээр нь хадгалахгvй (ажилтны хуудсанд HTML болж орох эрсдэлтэй).
  return m ? m[0] : "";
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
  replaceSheetRows(sheet, ["key", "value"], Object.keys(site).map((key) => [safeCell(key), safeCell(site[key])]));
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
  const out = [];
  rows.forEach((r) => {
    const dept = String((r.department != null ? r.department : "")).trim();
    const pos = String((r.position != null ? r.position : "")).trim();
    if (dept) out.push([safeCell(dept), safeCell(pos)]);
  });
  replaceSheetRows(sheet, ["department", "position"], out);
}

// Латин vсгийг ижил харагддаг кирилл vсэг болгоно (гар сольж бичихэд элбэг: латин "C" ба
// кирилл "С"). Латин "v" кирилл нэрэнд зөвхөн "ү"-гийн оронд бичигддэг.
const HOMOGLYPHS = { a: "а", b: "в", c: "с", e: "е", h: "н", k: "к", m: "м", o: "о", p: "р", t: "т", x: "х", y: "у", v: "ү", "ё": "е" };

// Зөвхөн харьцуулахад зориулсан түлхvvр (Sheet-ийн өгөгдлийг өөрчлөхгvй): том/жижиг vсэг,
// давхар зай, цэг/зураасны орчмын зай, латин/кирилл ижил vсэг, харагдахгvй тэмдэгтийг
// тооцохгvй — "С. Мөнхзул", "С.Мөнхзул", "с.мөнхзул", латин "C.Мөнхзул" нэг хvн болно.
function normalizeKey(s) {
  return String(s || "")
    .normalize("NFC")
    .replace(/[​-‍﻿]/g, "")
    .toLowerCase()
    .replace(/[a-zё]/g, (ch) => HOMOGLYPHS[ch] || ch)
    .replace(/[·․‧]/g, ".")
    .replace(/\s*([.,\-])\s*/g, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

const CYR2LAT = { "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo", "ж": "j", "з": "z", "и": "i", "й": "i",
  "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "ө": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ү": "u",
  "ф": "f", "х": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "i", "ы": "i", "ь": "i", "э": "e", "ю": "yu", "я": "ya" };

function hasCyrillic(s) {
  return /[Ѐ-ӿ]/.test(s);
}

function isLatinName(s) {
  return !hasCyrillic(s) && /[a-z]/i.test(s);
}

// Латинаар бичсэн нэрийг ("Munkhsuld") кирилл нэртэй ("Мөнхсүлд") дуудлагаар нь тулгах
// "араг яс": галиглалын түгээмэл ялгааг (kh/h/x, ө/ү/o/u, ts/c, давхар эгшиг, зураас, зай,
// ь→i) тооцохгvй. Хэт өргөн тул зөвхөн латин ↔ кирилл тулгахад хэрэглэнэ.
function translitKey(name) {
  let s = String(name || "").normalize("NFC").replace(/[​-‍﻿]/g, "").toLowerCase();
  const cyr = hasCyrillic(s);
  s = s.replace(/[Ѐ-ӿ]/g, (ch) => (ch in CYR2LAT ? CYR2LAT[ch] : ch));
  if (cyr) s = s.replace(/v/g, "u"); // кирилл нэр доторх латин "v" нь "ү"-гийн орлуулга
  return s
    .replace(/[·․‧]/g, ".")
    .replace(/[\s\-]+/g, "")
    .replace(/kh|x/g, "h")
    .replace(/zh/g, "j")
    .replace(/oe|ue|[öü]/g, "u")
    .replace(/o/g, "u")
    .replace(/c(?!h)/g, "ts")
    .replace(/w/g, "v")
    .replace(/y/g, "i")
    .replace(/([aeiu])\1+/g, "$1");
}

// Ижил алба+тушаалд латин нэртэй бичлэгийг дуудлага нь таарсан цорын ганц кирилл
// бичлэгтэй (эсвэл бусад латин хувилбартай) нэгтгэнэ. Нэг араг ястай кирилл бичлэг 2+
// байвал аль нь болохыг мэдэх боломжгvй тул нэгтгэхгvй.
function mergeLatinSpellings(people) {
  const buckets = {};
  Object.keys(people).forEach((key) => {
    const p = people[key];
    const b = normalizeKey(p.org) + "||" + normalizeKey(p.position) + "||" + translitKey(Object.keys(p.variants)[0]);
    (buckets[b] = buckets[b] || []).push(key);
  });
  Object.keys(buckets).forEach((b) => {
    const keys = buckets[b];
    if (keys.length < 2) return;
    const latin = keys.filter((k) => Object.keys(people[k].variants).every(isLatinName));
    const cyrillic = keys.filter((k) => latin.indexOf(k) === -1);
    if (!latin.length || cyrillic.length > 1) return;
    const target = people[cyrillic.length ? cyrillic[0] : latin[0]];
    keys.forEach((k) => {
      const src = people[k];
      if (src === target) return;
      Object.keys(src.variants).forEach((n) => { target.variants[n] = (target.variants[n] || 0) + src.variants[n]; });
      Object.keys(src.completed).forEach((w) => { target.completed[w] = true; });
      delete people[k];
    });
  });
}

/* ---------- Progress (event log) ---------- */
const PROGRESS_COLUMNS = ["timestamp", "org", "position", "name", "weekId", "event"];

function resetProgress() {
  const sheet = getOrCreateSheet(SHEET_PROGRESS, PROGRESS_COLUMNS);
  sheet.clearContents();
  sheet.appendRow(PROGRESS_COLUMNS);
}

// Ажилтны хуудас зөвхөн эдгээр vйл явдлыг илгээдэг — нийтийн endpoint тул бусдыг хvлээж авахгvй.
const PROGRESS_EVENTS = ["login", "video_completed"];

function logProgress(entry) {
  const sheet = getOrCreateSheet(SHEET_PROGRESS, PROGRESS_COLUMNS);
  const clip = (v) => safeCell(String(v == null ? "" : v).slice(0, 200));
  sheet.appendRow([
    new Date().toISOString(),
    clip(entry.org),
    clip(entry.position),
    clip(entry.name),
    clip(entry.weekId),
    clip(entry.event),
  ]);
}

/* ---------- config.json-г GitHub Pages-д шууд нийтлэх ---------- */
// GitHub-ийн cron нь "5 минут тутам" гэж тохируулсан ч бодит байдалд 15–20 минут
// тутам л ажилладаг тул админы өөрчлөлт (жишээ нь сайт хаах) ажилтнуудад хоцорч
// хvрдэг. Иймд хадгалах бvрт config.json-г GitHub-д шууд commit хийж ~1–2 минутад
// тусгана; cron workflow нь зөвхөн нөөц болж vлдэнэ.
// Script Properties → GITHUB_TOKEN: зөвхөн энэ repo-д "Contents: Read and write"
// эрхтэй fine-grained token. Тохируулаагvй бол зөвхөн cron-оор (хоцорч) шинэчлэгдэнэ.
const GITHUB_REPO = "ubp-iso-cpu/Ubp.iso";
const GITHUB_BRANCH = "main";
const GITHUB_CONFIG_PATH = "config.json";

// doGet болон GitHub-д нийтлэх config.json хоёулаа яг ижил агуулгатай байх ёстой
// (эс тэгвээс cron болон шууд нийтлэлт ээлжлэн commit хийж "эргэлдэнэ").
// rev нь админ цонх хуучирсан эсэхийг шалгахад хэрэглэгдэнэ (ажилтны хуудас үл тооно).
function buildPublicBundle() {
  return { weeks: readWeeks(), site: readSite(), org: readOrg(), rev: readRevisions() };
}

/* ---------- Зэрэг засварлалтаас хамгаалах (revision) ---------- */
// Хоёр цонх/админ нэг хэсгийг зэрэг засвал сvvлд хадгалсан нь нөгөөгийнхийг
// чимээгvй дарж устгадаг байсан. Хэсэг бvрийн хувилбарын дугаарыг хадгалж,
// хуучирсан цонхоос хадгалахыг татгалзана.
const REV_KEYS = { weeks: "REV_WEEKS", site: "REV_SITE", org: "REV_ORG" };

function readRevisions() {
  const p = PropertiesService.getScriptProperties();
  return {
    weeks: Number(p.getProperty(REV_KEYS.weeks)) || 0,
    site: Number(p.getProperty(REV_KEYS.site)) || 0,
    org: Number(p.getProperty(REV_KEYS.org)) || 0,
  };
}

// Шалгах–бичих–дугаар өсгөх–нийтлэх дөрвийг lock дотор хийнэ. Агуулга бодитоор
// өөрчлөгдсөн vед л дугаарыг өсгөнө (өөрчлөлтгvй хадгалалт бусад цонхыг хуучруулахгvй).
// Нийтлэлтийг lock дотор хийснээр зэрэг хадгалалтын хуучин snapshot шинийг дарахгvй.
// baseRev-гvй хvсэлт (хуучирсан admin.html) бусдын засварыг дарах эрсдэлтэй тул татгалзана.
function saveSection(section, baseRev, readFn, writeFn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const current = Number(props.getProperty(REV_KEYS[section])) || 0;
    if (baseRev === undefined || baseRev === null) return { error: "outdated_client", rev: current };
    if (Number(baseRev) !== current) return { error: "conflict", rev: current };
    const before = JSON.stringify(readFn());
    writeFn();
    const changed = JSON.stringify(readFn()) !== before;
    const rev = changed ? current + 1 : current;
    if (changed) props.setProperty(REV_KEYS[section], String(rev));
    return { rev: rev, publish: publishConfigToGitHub() };
  } finally {
    lock.releaseLock();
  }
}

function githubErrorMessage(res) {
  try {
    return String(JSON.parse(res.getContentText()).message || "");
  } catch (e) {
    return res.getContentText().slice(0, 200);
  }
}

// Буцаах утга: { status: "published" | "unchanged" | "no_token" | "error", code, detail }.
// Нийтлэлт амжилтгvй болсон ч Sheet-д хадгалалт аль хэдийн хийгдсэн тул алдаа
// шидэхгvй — cron нөөцөөр хожуу ч гэсэн шинэчлэгдэнэ. Web app-ийн лог Executions-д
// ихэвчлэн харагддаггvй тул шалтгааныг (code/detail) админд шууд буцаана.
function publishConfigToGitHub() {
  // Хуулж тавихад санамсаргvй орсон зай/мөр шилжилтийг арилгана.
  const token = String(PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN") || "").trim();
  if (!token) return { status: "no_token" };
  const json = JSON.stringify(buildPublicBundle());
  const api = "https://api.github.com/repos/" + GITHUB_REPO + "/contents/" + GITHUB_CONFIG_PATH;
  const headers = {
    Authorization: "Bearer " + token,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
  };
  try {
    for (let attempt = 1; attempt <= 2; attempt++) {
      const getRes = UrlFetchApp.fetch(api + "?ref=" + GITHUB_BRANCH, { headers: headers, muteHttpExceptions: true });
      const getCode = getRes.getResponseCode();
      let sha = null;
      if (getCode === 200) {
        const cur = JSON.parse(getRes.getContentText());
        sha = cur.sha;
        const currentText = Utilities.newBlob(
          Utilities.base64Decode(String(cur.content || "").replace(/\s/g, ""))
        ).getDataAsString("UTF-8");
        if (currentText === json) return { status: "unchanged" };
      } else if (getCode !== 404) {
        return { status: "error", code: getCode, detail: "GET " + getCode + ": " + githubErrorMessage(getRes) };
      }
      const payload = {
        message: "Publish config.json from admin save",
        content: Utilities.base64Encode(json, Utilities.Charset.UTF_8),
        branch: GITHUB_BRANCH,
      };
      if (sha) payload.sha = sha;
      const putRes = UrlFetchApp.fetch(api, {
        method: "put",
        contentType: "application/json",
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });
      const putCode = putRes.getResponseCode();
      if (putCode === 200 || putCode === 201) return { status: "published" };
      // 409/422: хооронд нь cron workflow commit хийж sha хуучирсан — дахин оролдоно
      if (putCode !== 409 && putCode !== 422) {
        return { status: "error", code: putCode, detail: "PUT " + putCode + ": " + githubErrorMessage(putRes) };
      }
    }
    return { status: "error", code: 409, detail: "PUT 409: config.json зэрэг өөрчлөгдөж байна, дахин хадгална уу" };
  } catch (err) {
    return { status: "error", code: 0, detail: String(err) };
  }
}

// saveSection-ийн vр дvнг (нийтлэлтийн төлөвийн хамт) админд буцаана.
function savedResponse(saved) {
  if (saved.error) return jsonResponse({ ok: false, error: saved.error, rev: saved.rev });
  const p = saved.publish;
  return jsonResponse({ ok: true, rev: saved.rev, publish: p.status, publishCode: p.code || 0, publishDetail: p.detail || "" });
}

function readProgressSummary() {
  const sheet = getOrCreateSheet(SHEET_PROGRESS, PROGRESS_COLUMNS);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter((r) => r[1] || r[3]);

  const weeks = readWeeks();
  const weekIds = weeks.map((w) => w.id);

  // Хvн бvрийг (алба+тушаал+нэр) нэгтгэж, тэдний бvртгvvлсэн болон дуусгасан
  // сургалтуудыг цуглуулна. Нэрээ өөр өөр vед өөрөөр бичсэн ч (normalizeKey-г харна)
  // нэг хvн гэж танина. Зөвхөн ижил алба+тушаалтай бол нэгтгэнэ — өөр албаны ижил
  // нэртэй хvмvvс ихэвчлэн өөр хvн байдаг.
  const people = {}; // key -> {org, position, completed, variants: {бичигдсэн нэр: тоо}}
  rows.forEach((r) => {
    const org = String(r[1] || "(тодорхойгvй)");
    const position = String(r[2] || "");
    const name = String(r[3] || "(тодорхойгvй)").trim().replace(/\s+/g, " ");
    const weekId = String(r[4] || "");
    const event = String(r[5] || "");
    const key = normalizeKey(org) + "||" + normalizeKey(position) + "||" + normalizeKey(name);
    if (!people[key]) {
      people[key] = { org: org, position: position, completed: {}, variants: {} };
    }
    people[key].variants[name] = (people[key].variants[name] || 0) + 1;
    if (event === "video_completed" && weekId) {
      people[key].completed[weekId] = true;
    }
  });

  mergeLatinSpellings(people);

  // Хамгийн олон удаа бичигдсэн хувилбарыг (кирилл хувилбар байвал түүнээс) vндсэн нэр
  // болгоно; нэгтгэсэн бол бvх хувилбарыг (aliases) буцааж, админ шалгах боломжтой болгоно.
  const peopleList = Object.keys(people).map((k) => {
    const p = people[k];
    const names = Object.keys(p.variants);
    const pool = names.some(hasCyrillic) ? names.filter(hasCyrillic) : names;
    const name = pool.reduce((best, n) => (p.variants[n] > p.variants[best] ? n : best), pool[0]);
    const person = { org: p.org, position: p.position, name: name, completed: p.completed };
    if (names.length > 1) person.aliases = names;
    return person;
  });

  // Сургалт тус бvрийн дуусгасан тоо
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

  return {
    weekIds: weekIds,
    weekTitles: weeks.map((w) => w.title),
    totalParticipants: peopleList.length,
    perWeek: perWeek,
    byOrg: Object.keys(byOrg).map((k) => byOrg[k]),
    people: peopleSorted,
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

// Web app-ийн URL ажилтны хуудсаар дамжин нийтэд ил тул нууц vгийг хязгааргvй таах
// боломжтой байсан. 15 минутад 20-оос олон буруу оролдлого гарвал тvр хаана.
const LOGIN_FAIL_KEY = "pw_fail_count";
const LOGIN_FAIL_LIMIT = 20;
const LOGIN_FAIL_WINDOW_SEC = 900;

function tooManyFailedLogins() {
  return Number(CacheService.getScriptCache().get(LOGIN_FAIL_KEY) || 0) >= LOGIN_FAIL_LIMIT;
}

function recordFailedLogin() {
  const cache = CacheService.getScriptCache();
  cache.put(LOGIN_FAIL_KEY, String(Number(cache.get(LOGIN_FAIL_KEY) || 0) + 1), LOGIN_FAIL_WINDOW_SEC);
}

// Хоосон биш буцаах утга нь тохирохгvй шалтгааныг илэрхийлнэ.
function validatePasswordStrength(pw) {
  if (pw.length < 8) return "Нууц vг дор хаяж 8 тэмдэгт байх ёстой.";
  if (!/[a-zA-Z]/.test(pw)) return "Нууц vг наад зах нь нэг vсэг агуулсан байх ёстой.";
  if (!/[0-9]/.test(pw)) return "Нууц vг наад зах нь нэг тоо агуулсан байх ёстой.";
  return null;
}

function doGet(e) {
  return jsonResponse(buildPublicBundle());
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
    if (!body.name || PROGRESS_EVENTS.indexOf(body.event) === -1) {
      return jsonResponse({ ok: false, error: "missing_fields" });
    }
    logProgress(body);
    return jsonResponse({ ok: true });
  }

  // Vлдсэн бvх vйлдэл админ нууц vг шаардана.
  if (tooManyFailedLogins()) {
    return jsonResponse({ ok: false, error: "too_many_attempts" });
  }
  if (!checkPassword(body)) {
    recordFailedLogin();
    return jsonResponse({ ok: false, error: "unauthorized" });
  }

  if (action === "changePassword") {
    const newPassword = String(body.newPassword || "");
    const weak = validatePasswordStrength(newPassword);
    if (weak) return jsonResponse({ ok: false, error: "weak_password", message: weak });
    PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", newPassword);
    return jsonResponse({ ok: true });
  }

  // Нэвтрэх vед нууц vгийг юу ч бичихгvйгээр шалгана.
  if (action === "verifyPassword") {
    return jsonResponse({ ok: true, rev: readRevisions() });
  }

  if (action === "saveWeeks") {
    if (!Array.isArray(body.weeks)) return jsonResponse({ ok: false, error: "expected_weeks_array" });
    return savedResponse(saveSection("weeks", body.baseRev, readWeeks, () => writeWeeks(body.weeks)));
  }

  if (action === "saveSite") {
    if (!body.site || typeof body.site !== "object") return jsonResponse({ ok: false, error: "expected_site_object" });
    return savedResponse(saveSection("site", body.baseRev, readSite, () => writeSite(body.site)));
  }

  if (action === "saveOrg") {
    if (!Array.isArray(body.org)) return jsonResponse({ ok: false, error: "expected_org_array" });
    return savedResponse(saveSection("org", body.baseRev, readOrg, () => writeOrg(body.org)));
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
