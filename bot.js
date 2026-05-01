import { chromium } from "playwright";
import XLSX from "xlsx";
import fs from "fs/promises";
import path from "path";

const BASE_URL = "https://portal.pua.edu.eg";
const LOGIN_URL = `${BASE_URL}/SelfService/Home.aspx`;
const GRADEBOOK_URL = `${BASE_URL}/SelfService/CourseManager/Gradebook.aspx?view=4`;

function cleanText(value) {
  return String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseCenterCourseText(raw, year, term, session) {
  const text = cleanText(raw);
  const parts = text.split("/").map(cleanText);
  return {
    raw: text,
    year,
    term,
    session,
    courseCode: parts[0] || "",
    activityType: parts[1] || "",
    section: parts[2] || "",
  };
}

function toAbsoluteUrl(href) {
  if (!href) return "";
  if (href.startsWith("http")) return href;
  return new URL(href, BASE_URL).toString();
}

function extractSectionId(href) {
  const match = String(href || "").match(/sectionid=(\d+)/i);
  return match ? match[1] : "";
}

function buildAttendancePageUrl(sectionId) {
  return `${BASE_URL}/SelfService/CourseManager/Gradebook.aspx?view=4&sectionid=${sectionId}`;
}

function buildAttendanceDownloadUrl(sectionId) {
  return `${BASE_URL}/SelfService/CourseManager/AttendanceDownload.aspx?sectionid=${sectionId}`;
}

function excelSerialToDate(serial) {
  const n = Number(serial);
  if (!Number.isFinite(n)) return null;
  if (n < 30000 || n > 60000) return null;
  const utcDays = Math.floor(n - 25569);
  const utcValue = utcDays * 86400;
  const date = new Date(utcValue * 1000);
  if (Number.isNaN(date.getTime())) return null;
  return date;
}

function twoDigitYearToFourDigit(yearText) {
  const n = Number(yearText);
  if (String(yearText).length === 2) return n >= 70 ? 1900 + n : 2000 + n;
  return n;
}

function normalizeExcelDateHeader(value) {
  if (value === null || value === undefined || value === "") return null;
  let dateObj = null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) dateObj = value;
  if (!dateObj && typeof value === "number") dateObj = excelSerialToDate(value);

  if (dateObj) {
    const yyyy = dateObj.getFullYear();
    const mm = String(dateObj.getMonth() + 1).padStart(2, "0");
    const dd = String(dateObj.getDate()).padStart(2, "0");
    return { isoDate: `${yyyy}-${mm}-${dd}`, displayDate: `${dd}/${mm}/${yyyy}`, raw: String(value) };
  }

  const text = cleanText(value);
  let m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    let a = Number(m[1]);
    let b = Number(m[2]);
    const year = twoDigitYearToFourDigit(m[3]);
    let month = a;
    let day = b;
    if (a > 12 && b <= 12) {
      day = a;
      month = b;
    }
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const mm = String(month).padStart(2, "0");
      const dd = String(day).padStart(2, "0");
      return { isoDate: `${year}-${mm}-${dd}`, displayDate: `${dd}/${mm}/${year}`, raw: text };
    }
  }

  const months = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3,
    apr: 4, april: 4, may: 5, jun: 6, june: 6, jul: 7, july: 7,
    aug: 8, august: 8, sep: 9, sept: 9, september: 9,
    oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12,
  };

  m = text.match(/^(\d{1,2})[-\s]+([A-Za-z]+)[-\s,]+(\d{2,4})/);
  if (m) {
    const day = Number(m[1]);
    const month = months[m[2].toLowerCase()];
    const year = twoDigitYearToFourDigit(m[3]);
    if (month && day >= 1 && day <= 31) {
      const mm = String(month).padStart(2, "0");
      const dd = String(day).padStart(2, "0");
      return { isoDate: `${year}-${mm}-${dd}`, displayDate: `${dd}/${mm}/${year}`, raw: text };
    }
  }

  m = text.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{2,4})/);
  if (m) {
    const month = months[m[1].toLowerCase()];
    const day = Number(m[2]);
    const year = twoDigitYearToFourDigit(m[3]);
    if (month && day >= 1 && day <= 31) {
      const mm = String(month).padStart(2, "0");
      const dd = String(day).padStart(2, "0");
      return { isoDate: `${year}-${mm}-${dd}`, displayDate: `${dd}/${mm}/${year}`, raw: text };
    }
  }
  return null;
}

function isoToPowerCampusDateParts(isoDate) {
  const m = String(isoDate || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return { year: m[1], month: String(Number(m[2])), day: String(Number(m[3])) };
}

function pcDateTextToIso(month, day, year) {
  return `${year}-${String(Number(month)).padStart(2, "0")}-${String(Number(day)).padStart(2, "0")}`;
}

function meetingIdFromRawDate(raw) {
  const m = cleanText(raw).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})/);
  if (!m) return "";
  const month = String(Number(m[1]));
  const day = String(Number(m[2]));
  const year = m[3];
  const hour = String(Number(m[4])).padStart(2, "0");
  const minute = m[5].padStart(2, "0");
  const second = m[6].padStart(2, "0");
  return `${day}${month}${year}${hour}${minute}${second}`;
}

function mapStatus(value) {
  const text = cleanText(value).toLowerCase();
  if (!text || text === "-") return "-";
  if (text.includes("present")) return "present";
  if (text.includes("absent") && !text.includes("exec")) return "absent";
  if (text.includes("exec")) return "excused_absence";
  if (text.includes("medical")) return "medical_excuse";
  if (text.includes("late")) return "late_registration";
  if (text.includes("mid")) return "mid_term";
  return cleanText(value);
}

function targetStatusText(localStatus) {
  const status = cleanText(localStatus).toLowerCase();
  if (status === "present") return "Present";
  if (status === "absent") return "Absent";
  if (status === "excused" || status === "excused_absence") return "Execused Absence";
  // AttendQR late is an attendance mark, not Late Registration in PowerCampus.
  // Mark it Present for PowerCampus unless you later decide on a different mapping.
  if (status === "late") return "Present";
  return status.includes("present") ? "Present" : "Absent";
}

export async function openBrowser() {
  return chromium.launch({ headless: process.env.HEADLESS !== "false" });
}

async function findFirstExisting(page, selectors) {
  for (const selector of selectors) {
    const loc = page.locator(selector).first();
    if (await loc.count()) return loc;
  }
  return null;
}

export async function login(page, username, password, targetUrl = LOGIN_URL) {
  await page.goto(targetUrl, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(1000);

  let bodyText = await page.locator("body").innerText().catch(() => "");
  let passwordFieldCount = await page.locator('input[type="password"]').count();

  const alreadyLoggedIn =
    /Welcome/i.test(bodyText) || /Log Out/i.test(bodyText) ||
    /Grading\/Attendance/i.test(bodyText) || /My Profile/i.test(bodyText) ||
    (/Course Sections/i.test(bodyText) && passwordFieldCount === 0);

  if (alreadyLoggedIn) {
    return { ok: true, title: await page.title().catch(() => ""), url: page.url(), welcomeText: (bodyText.match(/Welcome\s*\([^)]+\)/i) || [""])[0] };
  }

  const usernameInput = await findFirstExisting(page, [
    'input[type="text"]', 'input[name*="UserName" i]', 'input[id*="UserName" i]', 'input[name*="username" i]', 'input[id*="username" i]',
  ]);
  const passwordInput = await findFirstExisting(page, [
    'input[type="password"]', 'input[name*="Password" i]', 'input[id*="Password" i]', 'input[name*="password" i]', 'input[id*="password" i]',
  ]);

  if (!usernameInput || !passwordInput) {
    throw new Error(`Could not find username/password fields. URL="${page.url()}" Text="${bodyText.slice(0, 500)}"`);
  }

  await usernameInput.fill(username);
  await passwordInput.fill(password);

  let submit = page.locator('input[type="submit"]').first();
  if (!(await submit.count())) submit = page.getByRole("button", { name: /log|sign/i }).first();
  if (!(await submit.count())) throw new Error("Could not find login button.");

  await Promise.all([page.waitForLoadState("domcontentloaded").catch(() => null), submit.click()]);
  await page.waitForTimeout(3500);

  bodyText = await page.locator("body").innerText().catch(() => "");
  const title = await page.title().catch(() => "");
  const url = page.url();
  passwordFieldCount = await page.locator('input[type="password"]').count();

  const looksLoggedIn =
    /Welcome/i.test(bodyText) || /Log Out/i.test(bodyText) || /Classes/i.test(bodyText) ||
    /Grading\/Attendance/i.test(bodyText) || /My Profile/i.test(bodyText) ||
    /Course Sections/i.test(bodyText) ||
    (passwordFieldCount === 0 && /SelfService/i.test(url) && !/Login\.aspx/i.test(url));

  if (!looksLoggedIn) {
    throw new Error(`Login may have failed. Title="${title}" URL="${url}" Text="${bodyText.slice(0, 700)}"`);
  }

  return { ok: true, title, url, welcomeText: (bodyText.match(/Welcome\s*\([^)]+\)/i) || [""])[0] };
}

export async function testLogin({ username, password }) {
  const browser = await openBrowser();
  const page = await browser.newPage();
  try { return await login(page, username, password, GRADEBOOK_URL); }
  finally { await browser.close(); }
}

async function scrapeCenterCourses(page) {
  return await page.evaluate(() => {
    const clean = (v) => String(v || "").replace(/\s+/g, " ").trim();
    const bodyText = document.body.innerText || "";
    const year = (bodyText.match(/\b(202[0-9]|203[0-9])\b/) || ["", ""])[1];
    const term = (bodyText.match(/\b(Spring|Fall|Summer|Winter)\b/i) || ["", ""])[1];
    const session = (bodyText.match(/\bSession\s+\d+\b/i) || [""])[0];
    const allLinks = [...document.querySelectorAll("a")]
      .map((a) => ({ text: clean(a.innerText || a.textContent || ""), href: a.getAttribute("href") || "", absoluteHref: a.href || "" }))
      .filter((x) => x.text && /Gradebook\.aspx\?.*sectionid=\d+/i.test(x.absoluteHref));
    const centerLinks = allLinks.filter((x) => /^LC\s*\d+\s*\/\s*[^/]+\s*\/\s*[^/]+$/i.test(x.text));
    const seen = new Set();
    const unique = [];
    for (const link of centerLinks) {
      const sectionIdMatch = String(link.absoluteHref || "").match(/sectionid=(\d+)/i);
      const sectionId = sectionIdMatch ? sectionIdMatch[1] : link.absoluteHref;
      if (!seen.has(sectionId)) {
        seen.add(sectionId);
        link.absoluteHref = `https://portal.pua.edu.eg/SelfService/CourseManager/Gradebook.aspx?sectionid=${sectionId}`;
        unique.push(link);
      }
    }
    return { year, term, session, courses: unique, debugCountAllLinks: allLinks.length, debugCountCenterLinks: unique.length, debugUrl: window.location.href, debugTitle: document.title, debugTextStart: bodyText.slice(0, 1200) };
  });
}

export async function discoverCoursesFromCenter(page) {
  let data = await scrapeCenterCourses(page);
  if (!data.courses.length) {
    await page.goto(GRADEBOOK_URL, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(1500);
    data = await scrapeCenterCourses(page);
  }
  if (!data.courses.length) {
    const gradingLink = page.locator("a").filter({ hasText: /Grading\/Attendance/i }).first();
    if (await gradingLink.count()) {
      await Promise.all([page.waitForLoadState("domcontentloaded").catch(() => null), gradingLink.click()]);
      await page.waitForTimeout(1500);
      data = await scrapeCenterCourses(page);
    }
  }
  const courses = data.courses.map((link) => {
    const parsed = parseCenterCourseText(link.text, data.year, data.term, data.session);
    const href = toAbsoluteUrl(link.absoluteHref || link.href);
    const sectionId = extractSectionId(href);
    return { ...parsed, href, sectionId, attendancePageUrl: buildAttendancePageUrl(sectionId), attendanceDownloadUrl: buildAttendanceDownloadUrl(sectionId) };
  });
  return { year: data.year, term: data.term, session: data.session, courseCount: courses.length, courses, debugCountAllLinks: data.debugCountAllLinks, debugCountCenterLinks: data.debugCountCenterLinks, debugUrl: data.debugUrl, debugTitle: data.debugTitle, debugTextStart: data.debugTextStart };
}

async function chooseMicrosoftExcelAndDownload(page) {
  await page.waitForTimeout(1000);
  const selects = page.locator("select");
  const count = await selects.count();
  if (count < 1) throw new Error("Could not find Download Format dropdown.");
  const select = selects.first();
  const options = await select.locator("option").evaluateAll((ops) => ops.map((o) => ({ text: o.textContent?.trim() || "", value: o.getAttribute("value") || "" })));
  const excelOption = options.find((o) => /microsoft excel/i.test(o.text)) || options.find((o) => /excel/i.test(o.text));
  if (!excelOption) throw new Error("Could not find Microsoft Excel option. Found options: " + options.map((o) => o.text).join(", "));
  await select.selectOption(excelOption.value);
  const submitCandidates = ['input[type="submit"]', "button", 'input[type="button"]'];
  let clickTarget = null;
  for (const selector of submitCandidates) {
    const loc = page.locator(selector).first();
    if (await loc.count()) { clickTarget = loc; break; }
  }
  if (!clickTarget) throw new Error("Could not find download button.");
  const [download] = await Promise.all([page.waitForEvent("download", { timeout: 30000 }), clickTarget.click()]);
  return download;
}

function parseAttendanceWorkbook(filePath, courseMeta) {
  const workbook = XLSX.readFile(filePath, { cellDates: true, raw: true });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  if (!rows.length) throw new Error("Downloaded .xls file is empty.");

  let headerRowIndex = -1, idCol = -1, nameCol = -1;
  for (let r = 0; r < Math.min(rows.length, 40); r++) {
    const row = rows[r].map(cleanText);
    const possibleId = row.findIndex((c) => /^(people\s*id|student\s*id|id)$/i.test(c));
    const possibleName = row.findIndex((c) => /^(name|student\s*name)$/i.test(c));
    if (possibleId >= 0 && possibleName >= 0) { headerRowIndex = r; idCol = possibleId; nameCol = possibleName; break; }
  }
  if (headerRowIndex < 0) throw new Error("Could not find header row with People ID and Name.");

  const dateColumnsMap = new Map();
  const scanStart = Math.max(0, headerRowIndex - 3);
  const scanEnd = Math.min(rows.length - 1, headerRowIndex + 3);
  for (let r = scanStart; r <= scanEnd; r++) {
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      if (c === idCol || c === nameCol) continue;
      const normalized = normalizeExcelDateHeader(row[c]);
      if (normalized && !dateColumnsMap.has(c)) dateColumnsMap.set(c, { columnIndex: c, ...normalized, headerRowUsed: r });
    }
  }
  const dateColumns = [...dateColumnsMap.values()].sort((a, b) => a.columnIndex - b.columnIndex);

  const studentsMap = new Map();
  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r];
    const studentId = cleanText(row[idCol]);
    if (!/^\d{6,}$/.test(studentId)) continue;
    const name = cleanText(row[nameCol]);
    if (!name) continue;
    if (!studentsMap.has(studentId)) studentsMap.set(studentId, { studentId, name, records: [] });
    const student = studentsMap.get(studentId);
    for (const dateCol of dateColumns) {
      const rawStatus = cleanText(row[dateCol.columnIndex]);
      student.records.push({ isoDate: dateCol.isoDate, displayDate: dateCol.displayDate, rawStatus, status: mapStatus(rawStatus) });
    }
  }

  return { ...courseMeta, dates: dateColumns.map((d) => ({ isoDate: d.isoDate, displayDate: d.displayDate, raw: d.raw, columnIndex: d.columnIndex, headerRowUsed: d.headerRowUsed })), dateCount: dateColumns.length, students: [...studentsMap.values()], studentCount: studentsMap.size, sheetName: firstSheetName, headerRowIndex, idCol, nameCol, debugFirstRows: rows.slice(0, 5).map((row) => row.map(cleanText)) };
}

export async function importCourseFromDownload(page, course) {
  const sectionId = course.sectionId || extractSectionId(course.href);
  if (!sectionId) throw new Error(`Missing sectionId for course ${course.raw}`);
  const downloadUrl = buildAttendanceDownloadUrl(sectionId);
  await page.goto(downloadUrl, { waitUntil: "domcontentloaded" });
  const download = await chooseMicrosoftExcelAndDownload(page);
  const downloadedPath = await download.path();
  if (!downloadedPath) throw new Error("Download path is empty. Could not access downloaded .xls.");
  const suggested = download.suggestedFilename();
  const tempDir = path.join(process.cwd(), "downloads");
  await fs.mkdir(tempDir, { recursive: true });
  const safeName = suggested || `attendance-${sectionId}.xls`;
  const savedPath = path.join(tempDir, `${sectionId}-${Date.now()}-${safeName}`);
  await fs.copyFile(downloadedPath, savedPath);
  return parseAttendanceWorkbook(savedPath, { ...course, sectionId, attendancePageUrl: buildAttendancePageUrl(sectionId), attendanceDownloadUrl: downloadUrl, downloadedFile: suggested, downloadedPath: savedPath });
}

export async function importAllFromPowerCampus({ username, password }) {
  const browser = await openBrowser();
  const page = await browser.newPage({ acceptDownloads: true });
  try {
    await login(page, username, password, GRADEBOOK_URL);
    const discovery = await discoverCoursesFromCenter(page);
    const courses = [], errors = [];
    for (const course of discovery.courses) {
      try { courses.push(await importCourseFromDownload(page, course)); }
      catch (err) { errors.push({ course: course.raw, sectionId: course.sectionId, error: err.message }); }
    }
    return { ok: errors.length === 0, source: "powercampus-download-xls", discoveredCourseCount: discovery.courseCount, importedCourseCount: courses.length, errors, courses };
  } finally { await browser.close(); }
}

export async function discoverCoursesOnly({ username, password }) {
  const browser = await openBrowser();
  const page = await browser.newPage();
  try { await login(page, username, password, GRADEBOOK_URL); return await discoverCoursesFromCenter(page); }
  finally { await browser.close(); }
}

async function resolveMeetingPage(page, attendanceUrl, isoDate) {
  await page.goto(attendanceUrl || GRADEBOOK_URL, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(1500);

  if (/meetingId=\d+/i.test(page.url())) {
    return { url: page.url(), method: "already-on-meeting-page" };
  }

  const sectionId = extractSectionId(page.url()) || extractSectionId(attendanceUrl);
  if (!sectionId) throw new Error("Could not determine PowerCampus sectionId from attendance URL.");

  const targetParts = isoToPowerCampusDateParts(isoDate);
  if (!targetParts) throw new Error(`Invalid local date for PowerCampus push: ${isoDate}`);
  const targetText = `${targetParts.month}/${targetParts.day}/${targetParts.year}`;

  const linkResult = await page.evaluate((targetTextArg) => {
    const clean = (v) => String(v || "").replace(/\s+/g, " ").trim();
    const links = [...document.querySelectorAll("a")].map((a) => ({ text: clean(a.innerText || a.textContent || ""), href: a.href || "" }));
    const direct = links.find((x) => x.text.startsWith(targetTextArg) && /meetingId=\d+/i.test(x.href));
    return direct || null;
  }, targetText);

  if (linkResult?.href) {
    await page.goto(linkResult.href, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(1500);
    return { url: page.url(), method: "clicked-existing-date-link", raw: linkResult.text };
  }

  const rawDate = await page.evaluate((targetTextArg) => {
    const text = document.body.innerText || "";
    const escaped = targetTextArg.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(`${escaped}\\s+\\d{1,2}:\\d{2}:\\d{2}\\s+-\\s+\\d{1,2}:\\d{2}:\\d{2}`);
    const m = text.match(re);
    return m ? m[0] : "";
  }, targetText);

  if (rawDate) {
    const meetingId = meetingIdFromRawDate(rawDate);
    if (meetingId) {
      const constructedUrl = `${BASE_URL}/SelfService/CourseManager/Gradebook.aspx?sectionid=${sectionId}&view=4&meetingId=${meetingId}`;
      await page.goto(constructedUrl, { waitUntil: "domcontentloaded" });
      await page.waitForTimeout(1500);
      return { url: page.url(), method: "constructed-meeting-url", raw: rawDate, meetingId };
    }
  }

  throw new Error(`Could not find PowerCampus attendance date ${targetText} on the Attendance page.`);
}

async function readPowerCampusAttendanceRows(page) {
  return await page.evaluate(() => {
    const clean = (v) => String(v || "").replace(/\s+/g, " ").trim();
    const statusWords = ["Absent", "Execused Absence", "Excused Absence", "Late Registration", "Medical Execuse", "Medical Excuse", "Mid Term", "Present"];
    const rows = [];
    const selects = [...document.querySelectorAll("select")].filter((select) => /ddlbDailyAttendance/i.test(select.id || select.name || ""));

    for (const select of selects) {
      const tr = select.closest("tr");
      const text = clean(tr ? tr.innerText : select.parentElement?.innerText || "");
      const studentMatch = text.match(/\b\d{6,}\b/);
      if (!studentMatch) continue;
      const studentId = studentMatch[0];
      const options = [...select.options].map((o) => ({ value: o.value, text: clean(o.textContent || "") }));
      const selectedOption = options.find((o) => o.value === select.value) || null;
      let name = text.replace(studentId, " ");
      for (const word of statusWords) name = name.replace(new RegExp(word, "ig"), " ");
      name = clean(name.replace(/Apply Status/ig, " ").replace(/[-]+/g, " "));
      rows.push({ studentId, name, selectId: clean(select.id), selectName: clean(select.name), currentValue: select.value, currentStatus: selectedOption?.text || "", availableStatuses: options });
    }
    return rows;
  });
}

function makePushPlan(pcRows, localRecords) {
  const pcByStudentId = new Map(pcRows.map((row) => [String(row.studentId), row]));
  const localByStudentId = new Map((localRecords || []).map((record) => [String(record.studentId), record]));
  const matched = [];
  const missingOnPowerCampus = [];

  for (const record of localRecords || []) {
    const pc = pcByStudentId.get(String(record.studentId));
    if (!pc) {
      missingOnPowerCampus.push(record);
      continue;
    }
    const targetText = targetStatusText(record.status);
    const targetOption = pc.availableStatuses.find((o) => cleanText(o.text).toLowerCase() === targetText.toLowerCase());
    matched.push({
      studentId: record.studentId,
      name: record.name || pc.name,
      localStatus: record.status,
      targetPowerCampusStatus: targetText,
      targetValue: targetOption?.value || "",
      currentPowerCampusStatus: pc.currentStatus,
      currentValue: pc.currentValue,
      selectId: pc.selectId,
      selectName: pc.selectName,
      needsChange: targetOption ? String(targetOption.value) !== String(pc.currentValue) : false,
      error: targetOption ? "" : `PowerCampus option not found: ${targetText}`,
    });
  }

  const extraOnPowerCampus = pcRows.filter((row) => !localByStudentId.has(String(row.studentId))).map((row) => ({ studentId: row.studentId, name: row.name, currentStatus: row.currentStatus }));
  const errors = matched.filter((m) => m.error);
  const changes = matched.filter((m) => !m.error && m.needsChange);
  return { matched, changes, missingOnPowerCampus, extraOnPowerCampus, errors };
}

async function applyPushPlan(page, changes) {
  const result = await page.evaluate((changesArg) => {
    const clean = (v) => String(v || "").replace(/\s+/g, " ").trim();
    const out = [];
    const attendanceSelects = [...document.querySelectorAll("select")].filter((select) => /ddlbDailyAttendance/i.test(select.id || select.name || ""));

    for (const change of changesArg) {
      let select = null;

      if (change.selectId) {
        select = document.getElementById(clean(change.selectId)) || document.getElementById(change.selectId);
      }

      if (!select && change.selectName && window.CSS?.escape) {
        select = document.querySelector(`select[name="${CSS.escape(clean(change.selectName))}"]`) || document.querySelector(`select[name="${CSS.escape(change.selectName)}"]`);
      }

      if (!select) {
        select = attendanceSelects.find((candidate) => {
          const tr = candidate.closest("tr");
          const text = clean(tr ? tr.innerText : candidate.parentElement?.innerText || "");
          return text.includes(String(change.studentId));
        }) || null;
      }

      if (!select) {
        out.push({ studentId: change.studentId, ok: false, error: "Could not find status dropdown" });
        continue;
      }

      select.value = String(change.targetValue);
      select.dispatchEvent(new Event("input", { bubbles: true }));
      select.dispatchEvent(new Event("change", { bubbles: true }));
      out.push({ studentId: change.studentId, ok: true, targetValue: change.targetValue, targetPowerCampusStatus: change.targetPowerCampusStatus });
    }
    return out;
  }, changes);
  return result;
}

async function clickSave(page) {
  const clicked = await page.evaluate(() => {
    const candidates = [...document.querySelectorAll('input[type="submit"], input[type="button"], button')];
    const target = candidates.find((el) => /save|submit|update/i.test(el.value || el.textContent || el.getAttribute("title") || ""));
    if (target) { target.click(); return true; }
    return false;
  });
  if (!clicked) throw new Error("Could not find Save/Submit button on PowerCampus attendance page.");
  await page.waitForTimeout(2500);
  return { clicked: true, url: page.url(), title: await page.title().catch(() => "") };
}

export async function dryRunPushAttendance({ username, password, attendanceUrl, records, localSummary }) {
  const browser = await openBrowser();
  const page = await browser.newPage();
  try {
    await login(page, username, password, attendanceUrl || GRADEBOOK_URL);
    const meeting = await resolveMeetingPage(page, attendanceUrl, localSummary?.date);
    const pcRows = await readPowerCampusAttendanceRows(page);
    const plan = makePushPlan(pcRows, records || []);
    return {
      ok: plan.errors.length === 0,
      mode: "dry-run",
      submitted: false,
      meeting,
      localSummary,
      powerCampusSummary: { editableRows: pcRows.length },
      matchedCount: plan.matched.length,
      changeCount: plan.changes.length,
      missingOnPowerCampusCount: plan.missingOnPowerCampus.length,
      extraOnPowerCampusCount: plan.extraOnPowerCampus.length,
      errors: plan.errors,
      changes: plan.changes,
      missingOnPowerCampus: plan.missingOnPowerCampus,
      extraOnPowerCampus: plan.extraOnPowerCampus,
    };
  } finally { await browser.close(); }
}

export async function submitPushAttendance({ username, password, attendanceUrl, records, localSummary }) {
  if (process.env.ALLOW_POWER_CAMPUS_SUBMIT !== "true") {
    const dryRun = await dryRunPushAttendance({ username, password, attendanceUrl, records, localSummary });
    return { ...dryRun, ok: false, mode: "submit-blocked", submitted: false, error: "Real submit is locked. Set ALLOW_POWER_CAMPUS_SUBMIT=true in the bot .env only after dry-run is perfect." };
  }

  const browser = await openBrowser();
  const page = await browser.newPage();
  try {
    await login(page, username, password, attendanceUrl || GRADEBOOK_URL);
    const meeting = await resolveMeetingPage(page, attendanceUrl, localSummary?.date);
    const pcRows = await readPowerCampusAttendanceRows(page);
    const plan = makePushPlan(pcRows, records || []);
    if (plan.errors.length) {
      return { ok: false, mode: "submit", submitted: false, meeting, localSummary, errors: plan.errors, changes: plan.changes, missingOnPowerCampus: plan.missingOnPowerCampus };
    }
    const applied = await applyPushPlan(page, plan.changes);
    const failed = applied.filter((x) => !x.ok);
    if (failed.length) {
      return { ok: false, mode: "submit", submitted: false, meeting, localSummary, applied, failed };
    }
    const save = await clickSave(page);
    return { ok: true, mode: "submit", submitted: true, meeting, localSummary, appliedCount: applied.length, applied, save };
  } finally { await browser.close(); }
}
