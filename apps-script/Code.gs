/** @OnlyCurrentDoc */
// Ocean Charter Communities: the small backend that lives inside the Google Sheet.
// It saves new groups and broken-link reports, and hands the website the group
// list WITHOUT anyone's email.

const CATEGORIES = [
  "Learning differences",
  "Grades & classes",
  "Culture & language",
  "Family life",
  "Interests & hobbies",
  "Festivals & volunteering"
];
const GROUP_HEADERS = ["id", "added", "name", "category", "description", "link", "first_name", "email"];
const REPORT_HEADERS = ["reported", "group_id", "group_name", "link"];

// Run once from the editor: creates the Groups and Reports tabs with headers.
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [["Groups", GROUP_HEADERS], ["Reports", REPORT_HEADERS]].forEach(([name, headers]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sh.setFrozenRows(1);
  });
  const extra = ss.getSheetByName("Sheet1");
  if (extra && ss.getSheets().length > 2) ss.deleteSheet(extra);
}

// Website reads the group list.
function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Groups");
  const rows = sh.getDataRange().getValues().slice(1);
  const groups = rows
    .filter(r => r[0] && r[2])
    .map(r => ({
      id: String(r[0]),
      added: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      name: String(r[2]),
      cat: String(r[3]),
      desc: String(r[4]),
      link: String(r[5]),
      who: String(r[6])
      // email (r[7]) is never sent to the website
    }));
  return json({ ok: true, groups });
}

// Website sends a new group or a broken-link report.
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: "Bad request" });
  }
  if (data.website) return json({ ok: true }); // hidden spam-trap field was filled in by a bot

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (data.action === "add") return addGroup(data);
    if (data.action === "report") return addReport(data);
    return json({ ok: false, error: "Unknown action" });
  } finally {
    lock.releaseLock();
  }
}

function addGroup(d) {
  const g = {
    name: clean(d.name, 60),
    cat: clean(d.cat, 40),
    desc: clean(d.desc, 200),
    link: clean(d.link, 300),
    who: clean(d.who, 30),
    email: clean(d.email, 120)
  };
  if (!g.name || !g.desc || !g.who) return json({ ok: false, error: "Please fill in every field." });
  if (CATEGORIES.indexOf(g.cat) === -1) return json({ ok: false, error: "Please choose a category." });
  if (!/^https:\/\/\S+\.\S+$/.test(g.link)) return json({ ok: false, error: "Please paste a full link starting with https://" });
  if (!/^\S+@\S+\.\S+$/.test(g.email)) return json({ ok: false, error: "Please add a valid email." });

  const id = Utilities.getUuid().slice(0, 8);
  const added = new Date();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Groups")
    .appendRow([id, added, safe(g.name), g.cat, safe(g.desc), safe(g.link), safe(g.who), safe(g.email)]);

  notify("New group: " + g.name,
    g.who + " (" + g.email + ") added a group.\n\n" +
    g.name + "\n" + g.cat + "\n" + g.desc + "\n" + g.link + "\n\n" +
    "To remove it, delete its row in the sheet:\n" + SpreadsheetApp.getActiveSpreadsheet().getUrl());

  return json({ ok: true, group: { id, added: added.toISOString(), name: g.name, cat: g.cat, desc: g.desc, link: g.link, who: g.who } });
}

function addReport(d) {
  const id = clean(d.id, 20);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Groups");
  const row = sh.getDataRange().getValues().slice(1).find(r => String(r[0]) === id);
  if (!row) return json({ ok: false, error: "Group not found" });

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reports")
    .appendRow([new Date(), id, safe(String(row[2])), safe(String(row[5]))]);

  notify("Broken link reported: " + row[2],
    "Someone reported that this group's link may be broken.\n\n" +
    row[2] + "\n" + row[5] + "\n" +
    "Started by " + row[6] + " (" + row[7] + ")\n\n" +
    SpreadsheetApp.getActiveSpreadsheet().getUrl());

  return json({ ok: true });
}

// Emails the sheet's owner (you). Free limit is 100 a day.
function notify(subject, body) {
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "[OC Communities] " + subject, body);
  } catch (err) {
    // A failed notice should never block a parent from adding a group.
  }
}

function clean(v, max) {
  return String(v == null ? "" : v).replace(/\s+/g, " ").trim().slice(0, max);
}

// Stops text like "=IMPORTXML(...)" from running as a spreadsheet formula.
function safe(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
