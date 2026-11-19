/* =====================================================================
   Best India Group CRM - DAILY REPORTS (server)
   New file. Existing code is not changed.
   Sheet "Daily_Reports" in the CRM spreadsheet: one row per staff member per date.
   The app fills the report from the day's work; staff add their own lines and
   tomorrow's plan, save it here and send it to their WhatsApp group.
   - Staff can read and write only their own reports.
   - Managers can read the reports of the staff who report to them.
   - Admins can read everyone's.
   ===================================================================== */
const REP_SHEET = "Daily_Reports";
const REP_H = ["Date", "Staff ID", "Staff Name", "Activities", "Plan", "Report Text", "Saved At", "Sent At", "Meta"];

function repNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); }
function repToday_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
function repStr_(v) { if (v === null || v === undefined) return ""; if (Object.prototype.toString.call(v) === "[object Date]") return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); return String(v); }
function repDay_(v) { const s = repStr_(v).trim(); return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : s; }
function repCfg_() { try { return JSON.parse(readAll_().cfg || "{}"); } catch (e) { return {}; } }
function repName_(cfg, u) { if (u === "admin") return cfg.ownerName || "Admin"; const s = (cfg.staff || []).find(x => x && x.id === u); return s ? s.name : u; }
function repTeam_(cfg, me) {
  if (me.admin) return null; // null = everyone
  const staff = cfg.staff || [], m = staff.find(s => s && s.id === me.u), ids = [me.u];
  if (m && m.role === "mgr") staff.forEach(s => { if (s && [].concat(s.rep || []).indexOf(me.u) >= 0) ids.push(s.id); });
  return ids;
}
function repSheet_() {
  const sh = sheet_(REP_SHEET);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, REP_H.length).setValues([REP_H]).setFontWeight("bold").setBackground("#0F458B").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), REP_H.length).setNumberFormat("@");
    try { sh.setColumnWidth(4, 420); sh.setColumnWidth(5, 320); sh.setColumnWidth(6, 420); sh.hideColumns(9); } catch (e) {}
  }
  return sh;
}
function repRows_() {
  const sh = repSheet_(), last = sh.getLastRow(); if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, REP_H.length).getValues(), out = [];
  vals.forEach((r, i) => { if (!r[0] && !r[1]) return; out.push({ r: i + 2, d: repDay_(r[0]), u: repStr_(r[1]), n: repStr_(r[2]), a: repStr_(r[3]), p: repStr_(r[4]), x: repStr_(r[5]), s: repStr_(r[6]), w: repStr_(r[7]), m: repStr_(r[8]) }); });
  return out;
}
function repLines_(arr) { return (arr || []).map(x => String(x || "").replace(/\s+/g, " ").trim()).filter(Boolean).slice(0, 60).map(x => x.slice(0, 500)); }
function repOut_(o) { return { d: o.d, u: o.u, n: o.n, a: o.a ? o.a.split("\n").map(x => x.replace(/^\d+\.\s*/, "")) : [], p: o.p ? o.p.split("\n").map(x => x.replace(/^\d+\.\s*/, "")) : [], x: o.x, s: o.s, w: o.w, m: o.m }; }

/* Save (or update) my report for a date. o = {t, d, act:[...], plan:[...], txt, meta, sent:0|1} */
function crmRepSave(o) {
  const me = auth_(o && o.t), u = me.u;
  if (!u || u === "admin") throw new Error("Daily reports are for staff accounts");
  const d = String(o.d || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error("Wrong date");
  if (d > repToday_()) throw new Error("A report cannot be for a future date");
  const act = repLines_(o.act), plan = repLines_(o.plan);
  if (!act.length) throw new Error("Add at least one work activity");
  const cfg = repCfg_(), lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error("Busy - please try again in a few seconds");
  try {
    const sh = repSheet_(), R = repRows_(), old = R.find(x => x.d === d && x.u === u), now = repNow_();
    const row = {
      d, u, n: repName_(cfg, u),
      a: act.map((x, i) => (i + 1) + ". " + x).join("\n"),
      p: plan.map((x, i) => (i + 1) + ". " + x).join("\n"),
      x: String(o.txt || "").slice(0, 20000),
      s: now,
      w: o.sent ? now : (old ? old.w : ""),
      m: String(o.meta || "").slice(0, 4000)
    };
    const vals = [[row.d, row.u, row.n, row.a, row.p, row.x, row.s, row.w, row.m]];
    const r = old ? old.r : Math.max(sh.getLastRow(), 1) + 1, rg = sh.getRange(r, 1, 1, REP_H.length);
    rg.setNumberFormat("@"); rg.setValues(vals);
    return repOut_(row);
  } finally { lock.releaseLock(); }
}

/* Read reports. o = {t, from, to, u(optional staff id)} - at most 62 days at a time. */
function crmRepGet(o) {
  const me = auth_(o && o.t), cfg = repCfg_(), team = repTeam_(cfg, me);
  const to = /^\d{4}-\d{2}-\d{2}$/.test(String(o.to || "")) ? String(o.to) : repToday_();
  let from = /^\d{4}-\d{2}-\d{2}$/.test(String(o.from || "")) ? String(o.from) : to;
  const f = new Date(from + "T00:00:00"), t = new Date(to + "T00:00:00");
  if ((t - f) / 864e5 > 62) { f.setTime(t.getTime() - 62 * 864e5); from = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"); }
  return repRows_().filter(x => x.d >= from && x.d <= to && (!o.u || x.u === o.u) && (!team || team.indexOf(x.u) >= 0)).map(repOut_);
}
