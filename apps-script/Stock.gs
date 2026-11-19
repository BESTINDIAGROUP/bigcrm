// Best India Group CRM - STOCK & STORE (server).  Version s1
// A separate file in the same Apps Script project. It uses the CRM's existing spreadsheet (ss_ / sheet_),
// sign-in check (auth_), lock and settings - it does not change any existing function or sheet.
// Sheets used (created on first use, in the SAME spreadsheet): Stock_Products, Stock_Categories,
// Stock_Movements, Stock_Settings, Stock_Audit, Stock_Snapshot (hidden). Nothing else is read or written.
//
// ONE-TIME SETUP: run setupStockSync() from the editor (press Run, then Allow). It creates the sheets,
// protects them so only the owner can edit them directly, and installs the change watcher that records
// direct Google Sheet edits in the audit trail.

const STK = { P: "Stock_Products", C: "Stock_Categories", M: "Stock_Movements", S: "Stock_Settings", A: "Stock_Audit", N: "Stock_Snapshot" };
const STK_H = {
  P: ["Product ID", "Serial No", "Product Name", "Product Code", "Category", "Subcategory", "Specification", "Description", "Unit", "Rate", "Quantity", "Min Level", "Stock Value", "Status", "Updated At", "Updated By"],
  C: ["ID", "Category", "Subcategory", "Status"],
  M: ["Movement ID", "Date", "Type", "Product ID", "Product Name", "Quantity", "Balance After", "Store", "Project ID", "Project", "Sub-project", "Furnishing Status", "Remarks", "Staff", "User ID", "Recorded At", "Source"],
  S: ["Key", "Value"],
  A: ["Audit ID", "Date & Time", "Staff", "User ID", "Source", "Action", "Product ID", "Product Name", "Field", "Previous Value", "New Value", "Remarks", "Check"],
  N: ["Kind", "ID", "Data"]
};
const STK_NUM = { "Serial No": 1, "Rate": 1, "Quantity": 1, "Min Level": 1, "Stock Value": 1, "Balance After": 1 };
const STK_TYPES = { "Stock In": 1, "Return": 1, "Stock Issued": -1, "Damage / Loss": -1, "Adjustment": 0 };
const STK_PF = ["Serial No", "Product Name", "Product Code", "Category", "Subcategory", "Specification", "Description", "Unit", "Rate", "Min Level", "Status"];
const STK_UNITS = ["Nos", "Sets", "Boxes", "Packs", "Kg", "Grams", "Metres", "Litres", "Rolls", "Pairs", "Pieces", "Units", "Assets"];
const STK_CATS = [["Chemistry", ["Glassware", "Chemicals", "Apparatus"]], ["Physics", ["Electronics", "Measuring Instruments", "Laboratory Apparatus"]], ["Mathematics", ["Geometry Instruments", "Mathematical Models"]], ["Biology", ["Microscopes", "Slides", "Dissection Instruments"]], ["IT Products", ["Computers", "Printers", "Accessories"]], ["Other", []]];
const STK_STATUS = ["Issued", "Delivered to school", "Installed / Furnished"];

/* ---------------- small helpers ---------------- */
function stkNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); }
function stkToday_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
function stkStr_(v) { if (v === null || v === undefined) return ""; if (Object.prototype.toString.call(v) === "[object Date]") return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd"); return String(v).trim(); }
function stkNum_(v) { if (v === "" || v === null || v === undefined) return null; const n = typeof v === "number" ? v : Number(String(v).replace(/[₹,\s]/g, "")); return isFinite(n) ? n : NaN; }
function stkKey_(s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, ""); }
function stkAN_(v) { return STK_H.A.slice(0, 12).map(h => (v[h] === null || v[h] === undefined ? "" : String(v[h])).trim()).join("|"); }
function stkHash_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s))).slice(0, 16); }
function stkSettings_() {
  if (STK_MEMO._S) return STK_MEMO._S;
  const out = {}; stkRows_("S").forEach(r => { try { out[r.v.Key] = JSON.parse(r.v.Value); } catch (e) { out[r.v.Key] = r.v.Value; } });
  out.keepers = out.keepers || []; out.units = out.units || STK_UNITS.slice(); out.statuses = out.statuses || STK_STATUS.slice();
  out.stores = out.stores || [{ id: "main", n: "Central Store" }]; STK_MEMO._S = out; return out;
}
function stkSetSetting_(k, v) {
  delete STK_MEMO._S; STK_MEMO._meta = 1; const sh = stkSheet_("S"), R = stkRows_("S"), r = R.find(x => x.v.Key === k), val = JSON.stringify(v);
  if (r) sh.getRange(r.r, 2).setValue(val); else stkAppend_("S", { Key: k, Value: val });
}
function stkWho_(u) {
  if (!u) return "";
  let d = {}; try { d = JSON.parse(readAll_().cfg || "{}"); } catch (e) {}
  if (u === "admin") return d.ownerName || "Admin";
  const s = (d.staff || []).find(x => x && x.id === u); return s ? s.name : u;
}
function stkPerm_(me) { const st = stkSettings_(); return { admin: !!me.admin, keeper: !!me.admin || st.keepers.indexOf(me.u) >= 0 }; }

/* ---------------- sheets ---------------- */
var STK_MEMO = {};
function stkSheet_(k) {
  if (STK_MEMO[k]) return STK_MEMO[k];
  const sh = sheet_(STK[k]), H = STK_H[k]; STK_MEMO[k] = sh;
  if (sh.getLastRow() < 1 || String(sh.getRange(1, 1).getValue()) !== H[0]) {
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight("bold").setBackground("#0F458B").setFontColor("#ffffff");
      sh.setFrozenRows(1);
      H.forEach((h, i) => { if (!STK_NUM[h]) sh.getRange(2, i + 1, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat("@"); });
      if (k === "N") try { sh.hideSheet(); } catch (e) {}
    }
  }
  return sh;
}
function stkRows_(k, from, to) {
  const sh = stkSheet_(k), H = STK_H[k], last = sh.getLastRow(); if (last < 2) return [];
  const f0 = Math.max(2, from || 2), t0 = Math.min(last, to || last); if (t0 < f0) return [];
  const head = STK_MEMO["_h" + k] || (STK_MEMO["_h" + k] = sh.getRange(1, 1, 1, Math.max(H.length, sh.getLastColumn())).getValues()[0].map(stkStr_));
  const idx = H.map(h => head.indexOf(h)), vals = sh.getRange(f0, 1, t0 - f0 + 1, head.length).getValues(), out = [];
  vals.forEach((row, i) => { i += f0 - 2; if (row.every(c => c === "" || c === null)) return; const v = {}; H.forEach((h, j) => { const c = idx[j] < 0 ? "" : row[idx[j]]; v[h] = STK_NUM[h] ? stkNum_(c) : stkStr_(c); }); out.push({ r: i + 2, v }); });
  return out;
}
function stkRowArr_(k, v) { return STK_H[k].map(h => { const x = v[h]; if (STK_NUM[h]) return x === null || x === undefined || x === "" || (typeof x === "number" && !isFinite(x)) ? "" : x; return x === null || x === undefined ? "" : String(x); }); }
function stkAppend_(k, v) {
  if (k === "P" && v["Product ID"]) stkChg_(v["Product ID"]);
  const sh = stkSheet_(k), r = Math.max(sh.getLastRow(), 1) + 1, H = STK_H[k];
  const rg = sh.getRange(r, 1, 1, H.length); rg.setNumberFormats([H.map(h => STK_NUM[h] ? "General" : "@")]); rg.setValues([stkRowArr_(k, v)]); return r;
}
function stkChg_(id) { (STK_MEMO._chg = STK_MEMO._chg || {})[id] = 1; }
function stkWrite_(k, r, v) { if (k === "P" && v["Product ID"]) stkChg_(v["Product ID"]); if (k === "C" || k === "S") STK_MEMO._meta = 1; const H = STK_H[k], rg = stkSheet_(k).getRange(r, 1, 1, H.length); rg.setNumberFormats([H.map(h => STK_NUM[h] ? "General" : "@")]); rg.setValues([stkRowArr_(k, v)]); }
function stkProtect_() {
  const me = Session.getEffectiveUser();
  ["P", "C", "M", "S", "A", "N"].forEach(k => {
    const sh = stkSheet_(k), ex = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const p = ex.length ? ex[0] : sh.protect();
    p.setDescription("Stock & Store - only the owner may edit directly. Changes are recorded in Stock_Audit.");
    try { p.addEditor(me); p.removeEditors(p.getEditors().filter(x => x.getEmail() !== me.getEmail())); if (p.canDomainEdit()) p.setDomainEdit(false); } catch (e) {}
  });
}

/* ---------------- audit (append-only, chained check) ---------------- */
function stkAudit_(L) {
  if (!L.length) return;
  const sh = stkSheet_("A"), last = sh.getLastRow(); let prev = last >= 2 ? stkStr_(sh.getRange(last, 13).getValue()) : "";
  let n = +(props_().getProperty("STK_NA") || 0); const rows = [];
  L.forEach(a => {
    n++; const v = { "Audit ID": "A-" + String(n).padStart(6, "0"), "Date & Time": a.t || stkNow_(), Staff: a.who || "", "User ID": a.u || "", Source: a.src || "App", Action: a.act, "Product ID": a.pid || "", "Product Name": a.pn || "", Field: a.f || "", "Previous Value": a.o === undefined || a.o === null ? "" : String(a.o), "New Value": a.n === undefined || a.n === null ? "" : String(a.n), Remarks: a.rm || "" };
    v.Check = stkHash_(prev + "|" + stkAN_(v)); prev = v.Check; rows.push(stkRowArr_("A", v));
  });
  props_().setProperty("STK_NA", String(n));
  const rg = sh.getRange(Math.max(last, 1) + 1, 1, rows.length, STK_H.A.length); rg.setNumberFormat("@"); rg.setValues(rows);
}
function stkAuditCheck_(R) { let prev = ""; for (let i = 0; i < R.length; i++) { const v = R[i].v, h = stkHash_(prev + "|" + stkAN_(v)); if (h !== v.Check) return { ok: false, row: R[i].r }; prev = v.Check; } return { ok: true }; }

/* ---------------- snapshot (hidden) used to detect direct Google Sheet edits ---------------- */
function stkSnapRead_() {
  const o = { P: {}, X: {}, C: {}, M: {} }, at = {}, mc = {};
  stkRows_("N").forEach(r => { try { const d = JSON.parse(r.v.Data); at[r.v.Kind + "|" + r.v.ID] = { r: r.r, d: r.v.Data };
    if (r.v.Kind === "MC") Object.keys(d).forEach(id => { o.M[id] = d[id]; mc[id] = r.v.ID; });
    else { o[r.v.Kind] = o[r.v.Kind] || {}; o[r.v.Kind][r.v.ID] = d; } } catch (e) {} });
  STK_MEMO._NAT = at; STK_MEMO._MC = mc; return o;
}
function stkSnapWrite_(o) {
  const sh = stkSheet_("N"), at = STK_MEMO._NAT || {}, mc = STK_MEMO._MC || {}, seen = {}, add = [], rows = {};
  const put = (k, id, obj) => { const key = k + "|" + id, d = JSON.stringify(obj); seen[key] = 1; const was = at[key]; if (!was) add.push([k, id, d]); else if (was.d !== d) { const rg = sh.getRange(was.r, 1, 1, 3); rg.setNumberFormat("@"); rg.setValues([[k, id, d]]); } };
  ["P", "X", "C"].forEach(k => Object.keys(o[k] || {}).forEach(id => put(k, id, o[k][id])));
  const ch = {}, size = {}; Object.keys(mc).forEach(id => { size[mc[id]] = (size[mc[id]] || 0) + 1; });
  let cur = Object.keys(size).sort((a, b) => +a - +b).pop() || "1";
  Object.keys(o.M || {}).forEach(id => { let c = mc[id]; if (!c) { if ((size[cur] || 0) >= 250) cur = String(+cur + 1); c = cur; size[c] = (size[c] || 0) + 1; mc[id] = c; } (ch[c] = ch[c] || {})[id] = o.M[id]; });
  Object.keys(ch).forEach(c => put("MC", c, ch[c]));
  Object.keys(at).forEach(key => { if (!seen[key]) sh.getRange(at[key].r, 1, 1, 3).clearContent(); });
  if (add.length) { const r0 = Math.max(sh.getLastRow(), 1) + 1, rg = sh.getRange(r0, 1, add.length, 3); rg.setNumberFormat("@"); rg.setValues(add); }
  STK_MEMO._NAT = null;
}
/* ---------------- the core: read everything, absorb direct sheet edits, keep quantities true ---------------- */
// src: "App" or "Google Sheet"; who: staff name or e-mail for sheet edits.
function stkSync_(src, who, quiet) {
  const t = stkNow_(), au = [], flags = [], sheetWho = who || "Google Sheet edit (user not provided by Google)";
  const S = stkSettings_(), snap = stkSnapRead_(), first = !Object.keys(snap.P).length && !Object.keys(snap.C).length && !Object.keys(snap.M).length, nd = first || !!quiet;
  let np = +(props_().getProperty("STK_NP") || 0), nc = +(props_().getProperty("STK_NC") || 0), nm = +(props_().getProperty("STK_NM") || 0);
  /* categories */
  const C = stkRows_("C"), cSeen = {};
  C.forEach(r => { const v = r.v; let ch = false;
    if (!v.Category) { if (v.ID || v.Subcategory) flags.push("Stock_Categories row " + r.r + ": category name is empty."); return; }
    if (!v.ID || cSeen[v.ID]) { nc++; v.ID = (v.Subcategory ? "S-" : "C-") + String(nc).padStart(4, "0"); ch = true; if (!first) au.push({ act: "Category Changed", f: v.Subcategory ? "Subcategory added" : "Category added", n: v.Category + (v.Subcategory ? " › " + v.Subcategory : ""), src: "Google Sheet", who: sheetWho }); }
    if (!v.Status) { v.Status = "Active"; ch = true; }
    cSeen[v.ID] = 1; if (ch) stkWrite_("C", r.r, v);
    const o = snap.C[v.ID]; if (!o || ["Category", "Subcategory", "Status"].some(f => (o[f] || "") !== (v[f] || ""))) STK_MEMO._meta = 1; if (o && !nd) ["Category", "Subcategory", "Status"].forEach(f => { if ((o[f] || "") !== (v[f] || "")) au.push({ act: "Category Changed", f, o: o[f], n: v[f], src: "Google Sheet", who: sheetWho, rm: "Edited directly in Google Sheet" }); });
  });
  Object.keys(snap.C).forEach(id => { if (!cSeen[id] && !nd) au.push({ act: "Category Changed", f: "Deleted in Google Sheet", o: snap.C[id].Category + (snap.C[id].Subcategory ? " › " + snap.C[id].Subcategory : ""), src: "Google Sheet", who: sheetWho }); });
  const catNames = {}; C.forEach(r => { if (r.v.Category && r.v.Status !== "Deleted") { catNames[stkKey_(r.v.Category)] = 1; if (r.v.Subcategory) catNames[stkKey_(r.v.Category) + "›" + stkKey_(r.v.Subcategory)] = 1; } });
  /* products */
  const P = stkRows_("P"), pSeen = {}, byKey = {}; let maxSl = 0;
  P.forEach(r => { if (r.v["Serial No"] > maxSl) maxSl = r.v["Serial No"]; if (r.v["Product ID"] && r.v["Status"] !== "Deleted" && r.v["Status"] !== "Duplicate?") byKey[stkKey_(r.v["Product Name"]) + "|" + stkKey_(r.v.Specification)] = r.v["Product ID"]; });
  const prodById = {};
  P.forEach(r => {
    const v = r.v; let ch = false;
    if (!v["Product Name"]) { flags.push("Stock_Products row " + r.r + ": Product Name is empty."); if (v["Product ID"]) prodById[v["Product ID"]] = r; return; }
    if (v["Product ID"] && pSeen[v["Product ID"]]) { flags.push("Stock_Products row " + r.r + ": Product ID " + v["Product ID"] + " was repeated (copied row) – it was treated as a new row."); v["Product ID"] = ""; ch = true; }
    if (!v["Product ID"]) {
      const k = stkKey_(v["Product Name"]) + "|" + stkKey_(v.Specification);
      if (byKey[k]) { if (v.Status !== "Duplicate?") { v.Status = "Duplicate?"; stkWrite_("P", r.r, v); } flags.push("Stock_Products row " + r.r + ": \"" + v["Product Name"] + "\" duplicates " + byKey[k] + " – not added. Delete the row or change the name / specification."); return; }
      if (v.Status === "Duplicate?") v.Status = "";
      np++; v["Product ID"] = "P-" + String(np).padStart(6, "0"); if (!(v["Serial No"] > 0)) { maxSl++; v["Serial No"] = maxSl; }
      v.Status = v.Status || "Active"; v["Updated At"] = t; v["Updated By"] = src === "App" ? v["Updated By"] : sheetWho; byKey[k] = v["Product ID"]; ch = true;
      if (!first) au.push({ act: "Added", pid: v["Product ID"], pn: v["Product Name"], src: "Google Sheet", who: sheetWho, rm: "Row added directly in Google Sheet" });
    }
    if (!v.Status) { v.Status = "Active"; ch = true; }
    ["Rate", "Min Level", "Quantity"].forEach(f => { if (typeof v[f] === "number" && !isFinite(v[f])) { flags.push("Stock_Products row " + r.r + ": " + f + " is not a number."); v[f] = null; } });
    if (v.Category && !catNames[stkKey_(v.Category)]) flags.push("\"" + v["Product Name"] + "\": category \"" + v.Category + "\" is not in Stock settings.");
    else if (v.Category && v.Subcategory && !catNames[stkKey_(v.Category) + "›" + stkKey_(v.Subcategory)]) flags.push("\"" + v["Product Name"] + "\": subcategory \"" + v.Subcategory + "\" is not under " + v.Category + ".");
    pSeen[v["Product ID"]] = 1; prodById[v["Product ID"]] = r; r.ch = ch;
    const o = snap.P[v["Product ID"]];
    if (o) STK_PF.forEach(f => { const a = o[f] === null || o[f] === undefined ? "" : String(o[f]), b = v[f] === null || v[f] === undefined ? "" : String(v[f]); if (a !== b) stkChg_(v["Product ID"]); if (a !== b && !nd) au.push({ act: f === "Status" ? (b === "Deleted" ? "Deleted" : a === "Deleted" ? "Restored" : "Edited") : f === "Category" || f === "Subcategory" ? "Category Changed" : "Edited", pid: v["Product ID"], pn: v["Product Name"], f, o: a, n: b, src: "Google Sheet", who: sheetWho, rm: "Edited directly in Google Sheet" }); });
  });
  const X = snap.X || {};
  Object.keys(snap.P).forEach(id => { if (!pSeen[id] && !prodById[id] && !nd) { const o = snap.P[id]; X[id] = Object.assign({}, o, { "Product ID": id, deletedAt: t, deletedBy: sheetWho, how: "Row deleted in Google Sheet" }); stkChg_(id); au.push({ act: "Deleted", pid: id, pn: o["Product Name"], f: "Row deleted in Google Sheet", o: stkDesc_(o), src: "Google Sheet", who: sheetWho }); } });
  /* movements */
  const M = stkRows_("M"), mSeen = {}, bal = {}, has = {};
  M.sort((a, b) => (a.v.Date + a.v["Recorded At"]).localeCompare(b.v.Date + b.v["Recorded At"]) || a.r - b.r).forEach(r => {
    const v = r.v; let ch = false;
    if (!v["Movement ID"] || mSeen[v["Movement ID"]]) { nm++; v["Movement ID"] = "M-S" + String(nm).padStart(6, "0"); v.Source = "Google Sheet"; v.Staff = v.Staff || sheetWho; v["Recorded At"] = v["Recorded At"] || t; ch = true; if (!first) au.push({ act: stkActOf_(v.Type), pid: v["Product ID"], pn: v["Product Name"], f: "Quantity", n: v.Quantity, src: "Google Sheet", who: sheetWho, rm: "Movement row added directly in Google Sheet" }); }
    mSeen[v["Movement ID"]] = 1;
    const pr = prodById[v["Product ID"]] || (v["Product ID"] ? null : P.find(p => p.v["Product Name"] && stkKey_(p.v["Product Name"]) === stkKey_(v["Product Name"])));
    if (pr && !v["Product ID"]) { v["Product ID"] = pr.v["Product ID"]; ch = true; }
    const sign = STK_TYPES[v.Type], q = v.Quantity;
    if (sign === undefined) flags.push("Stock_Movements row " + r.r + ": unknown type \"" + v.Type + "\" – ignored.");
    else if (typeof q !== "number" || !isFinite(q) || (sign !== 0 && q <= 0)) flags.push("Stock_Movements row " + r.r + ": quantity is not valid – ignored.");
    else if (!pr && !X[v["Product ID"]]) flags.push("Stock_Movements row " + r.r + ": product " + (v["Product ID"] || v["Product Name"]) + " not found.");
    else { const id = v["Product ID"]; bal[id] = (bal[id] || 0) + (sign === 0 ? q : sign * q); has[id] = 1; if (v["Balance After"] !== bal[id]) { v["Balance After"] = bal[id]; ch = true; } }
    if (ch) stkWrite_("M", r.r, v);
    const h = stkHash_(STK_H.M.filter(x => x !== "Balance After" && x !== "Furnishing Status").map(x => v[x]).join("|")), o = snap.M[v["Movement ID"]];
    if (o && o.h !== h && !nd) au.push({ act: "Edited", pid: v["Product ID"], pn: v["Product Name"], f: "Movement " + v["Movement ID"], o: o.d, n: stkMovDesc_(v), src: "Google Sheet", who: sheetWho, rm: "Movement edited directly in Google Sheet" });
    if (!o || o.h !== h || (o.s || "") !== (v["Furnishing Status"] || "")) STK_MEMO._meta = 1;
    if (o && (o.s || "") !== (v["Furnishing Status"] || "") && !nd) au.push({ act: "Status Updated", pid: v["Product ID"], pn: v["Product Name"], f: "Furnishing status (" + v["Movement ID"] + ")", o: o.s, n: v["Furnishing Status"], src: "Google Sheet", who: sheetWho });
    r.h = h;
  });
  Object.keys(snap.M).forEach(id => { if (!mSeen[id] && !nd) au.push({ act: "Deleted", f: "Movement " + id, o: snap.M[id].d, src: "Google Sheet", who: sheetWho, rm: "Movement row deleted directly in Google Sheet" }); });
  /* quantities: the balance always equals the movements. A quantity typed in the sheet becomes an Adjustment. */
  P.forEach(r => {
    const v = r.v; if (!v["Product ID"] || v.Status === "Duplicate?") return;
    const id = v["Product ID"], b = has[id] ? bal[id] : null, q = v.Quantity, o = snap.P[id], oq = o ? (o.Quantity === undefined ? null : o.Quantity) : null;
    if (q !== b) {
      if (typeof q === "number" && isFinite(q) && q >= 0 && (o ? oq === b : !has[id])) {
        const d = q - (b || 0); nm++; const mv = { "Movement ID": "M-S" + String(nm).padStart(6, "0"), Date: stkToday_(), Type: "Adjustment", "Product ID": id, "Product Name": v["Product Name"], Quantity: d, "Balance After": q, Store: "main", Remarks: "Quantity edited directly in Google Sheet", Staff: sheetWho, "User ID": "", "Recorded At": t, Source: "Google Sheet" };
        if (d !== 0 || b === null) { stkAppend_("M", mv); au.push({ act: "Stock Adjusted", pid: id, pn: v["Product Name"], f: "Quantity", o: b === null ? "" : b, n: q, src: "Google Sheet", who: sheetWho, rm: "Quantity edited directly in Google Sheet" }); }
        has[id] = 1; bal[id] = q;
      } else if (typeof q === "number" && q < 0) { flags.push("\"" + v["Product Name"] + "\": a negative quantity was typed in the sheet – put back to " + (b === null ? "blank" : b) + "."); v.Quantity = b; r.ch = true; }
      else { v.Quantity = b; r.ch = true; }
    }
    const sv = typeof v.Quantity === "number" && typeof v.Rate === "number" ? Math.round(v.Quantity * v.Rate * 100) / 100 : null;
    if (sv !== v["Stock Value"]) { v["Stock Value"] = sv; r.ch = true; }
    if (r.ch) stkWrite_("P", r.r, v);
  });
  props_().setProperty("STK_NP", String(np)); props_().setProperty("STK_NC", String(nc)); props_().setProperty("STK_NM", String(nm));
  stkAudit_(au);
  /* new snapshot */
  const ns = { P: {}, X: X, C: {}, M: {} };
  P.forEach(r => { const v = r.v; if (v["Product ID"] && v.Status !== "Duplicate?") { const o = {}; STK_PF.concat(["Quantity"]).forEach(f => o[f] = v[f] === undefined ? null : v[f]); ns.P[v["Product ID"]] = o; if (X[v["Product ID"]]) delete X[v["Product ID"]]; } });
  C.forEach(r => { if (r.v.ID) ns.C[r.v.ID] = { Category: r.v.Category, Subcategory: r.v.Subcategory, Status: r.v.Status }; });
  M.forEach(r => { if (r.v["Movement ID"]) ns.M[r.v["Movement ID"]] = { h: r.h, d: stkMovDesc_(r.v), s: r.v["Furnishing Status"] || "" }; });
  stkSnapWrite_(ns);
  props_().setProperty("STK_FULLAT", String(Date.now())); props_().setProperty("STK_FLAGS", JSON.stringify(flags.slice(0, 200)));
  const mids = {}; M.forEach(r => { if (r.v["Movement ID"]) mids[r.v["Movement ID"]] = 1; });
  return { P, C, M, X, S, flags, mids, snap: ns, full: true };
}
function stkDesc_(o) { return ["Serial " + (o["Serial No"] || ""), o["Product Name"], o.Specification, o["Product Code"], o.Category + (o.Subcategory ? " › " + o.Subcategory : ""), "Unit " + (o.Unit || ""), "Rate " + (o.Rate === null || o.Rate === undefined ? "" : o.Rate), "Qty " + (o.Quantity === null || o.Quantity === undefined ? "" : o.Quantity)].filter(x => x && String(x).trim()).join(" · "); }
function stkMovDesc_(v) { return [v.Date, v.Type, v.Quantity, v["Product Name"], v.Project, v.Remarks].filter(x => x !== "" && x !== null && x !== undefined).join(" · "); }
function stkActOf_(t) { return t === "Stock In" ? "Stock Added" : t === "Return" ? "Stock Returned" : t === "Stock Issued" ? "Stock Issued" : t === "Damage / Loss" ? "Damage / Loss" : "Stock Adjusted"; }

/* ---------------- fast path: normal loads and saves read only what they need ---------------- */
function stkNeedFull_() { const a = +(props_().getProperty("STK_FULLAT") || 0); return !a || Date.now() - a > 10 * 60 * 1000; }
function stkLoad_() {
  const S = stkSettings_(), snap = stkSnapRead_(), P = stkRows_("P"), C = stkRows_("C"), sh = stkSheet_("M"), last = sh.getLastRow(), mids = {};
  let M = []; if (last >= 2) { sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => { if (r[0]) mids[String(r[0])] = 1; }); M = stkRows_("M", last - 799, last); }
  let flags = []; try { flags = JSON.parse(props_().getProperty("STK_FLAGS") || "[]"); } catch (e) {}
  return { P, C, M, X: snap.X || {}, S, flags, mids, snap, full: false };
}
function stkData_() { return stkNeedFull_() ? stkSync_("App") : stkLoad_(); }
/* change log: phones that are up to date receive only the products that changed */
function stkSeq_() { return +(props_().getProperty("STK_SEQ") || 0); }
function stkLogChanges_() {
  const ids = Object.keys(STK_MEMO._chg || {}); if (!ids.length && !STK_MEMO._meta) return stkSeq_();
  const seq = stkSeq_() + 1; let L = []; try { L = JSON.parse(props_().getProperty("STK_LOG") || "[]"); } catch (e) {}
  L.push(Object.assign({ s: seq, p: ids.length > 400 ? "*" : ids }, STK_MEMO._meta ? { m: 1 } : {})); while (JSON.stringify(L).length > 7000 && L.length > 1) L.shift();
  props_().setProperty("STK_LOG", JSON.stringify(L)); props_().setProperty("STK_SEQ", String(seq)); STK_MEMO._chg = {}; STK_MEMO._meta = 0; return seq;
}
function stkRevBump_() { stkLogChanges_(); }
// have = change number the phone already has; mt = newest "Recorded At" it already has.
function stkModel_(me, D, have, mt) {
  const pm = stkPerm_(me), seq = stkSeq_(), H = STK_H.P.filter(h => pm.admin || (h !== "Rate" && h !== "Stock Value")), MH = STK_H.M;
  let pick = null, meta = true; have = +have || 0;
  if (have > 0 && have <= seq) { let L = []; try { L = JSON.parse(props_().getProperty("STK_LOG") || "[]"); } catch (e) {}
    const need = L.filter(x => x.s > have);
    if (have === seq) { pick = {}; meta = false; }
    else if (need.length && need[0].s === have + 1 && !need.some(x => x.p === "*")) { pick = {}; need.forEach(x => x.p.forEach(id => pick[id] = 1)); meta = need.some(x => x.m); } }
  const out = { seq, na: +(props_().getProperty("STK_NA") || 0), delta: !!pick, H, MH, me: { u: me.u, admin: pm.admin, keeper: pm.keeper, name: stkWho_(me.u) }, t: stkNow_() };
  out.P = D.P.filter(r => r.v["Product ID"] && (!pick || pick[r.v["Product ID"]])).map(r => H.map(h => { const x = r.v[h]; return x === null || x === undefined ? "" : x; }));
  out.gone = pick ? Object.keys(pick).filter(id => !D.P.some(r => r.v["Product ID"] === id)) : [];
  if (!pick || Object.keys(pick).length || meta) {
    let M = D.M.slice().sort((a, b) => (b.v.Date + b.v["Recorded At"]).localeCompare(a.v.Date + a.v["Recorded At"]) || b.r - a.r);
    if (pick && !meta && mt) { M = D.M.filter(r => String(r.v["Recorded At"]) > mt); out.mAdd = true; }
    out.M = M.slice(0, 600).map(r => MH.map(h => { const x = r.v[h]; return x === null || x === undefined ? "" : x; }));
  }
  if (!pick || meta) { out.C = D.C.map(r => r.v); out.X = pm.admin ? Object.keys(D.X || {}).map(id => D.X[id]) : []; out.S = { units: D.S.units, statuses: D.S.statuses, stores: D.S.stores, keepers: D.S.keepers }; out.flags = pm.keeper ? D.flags : []; }
  return out;
}
/* after an app save: refresh the hidden snapshot only for what this save touched */
function stkSnapPatch_(D) {
  const ns = D.snap; ns.P = ns.P || {}; ns.M = ns.M || {}; ns.X = D.X || ns.X || {};
  Object.keys(STK_MEMO._chg || {}).forEach(id => { const r = D.P.find(x => x.v["Product ID"] === id); if (r && r.v.Status !== "Duplicate?") { const o = {}; STK_PF.concat(["Quantity"]).forEach(f => o[f] = r.v[f] === undefined ? null : r.v[f]); ns.P[id] = o; delete ns.X[id]; } });
  (STK_MEMO._newM || []).forEach(v => { ns.M[v["Movement ID"]] = { h: stkHash_(STK_H.M.filter(x => x !== "Balance After" && x !== "Furnishing Status").map(x => v[x]).join("|")), d: stkMovDesc_(v), s: v["Furnishing Status"] || "" }; });
  (STK_MEMO._stM || []).forEach(v => { if (ns.M[v["Movement ID"]]) ns.M[v["Movement ID"]].s = v["Furnishing Status"] || ""; });
  if (STK_MEMO._meta) { ns.C = {}; D.C.forEach(r => { if (r.v.ID) ns.C[r.v.ID] = { Category: r.v.Category, Subcategory: r.v.Subcategory, Status: r.v.Status }; }); }
  stkSnapWrite_(ns);
}
function stkRev_() { return String(stkSeq_()); }

/* ================= calls from the app ================= */
// Cheap "anything new?" check (no sheet is read). Every 10 minutes the first caller also runs the full check.
function crmStockRev(t) {
  auth_(t);
  if (stkNeedFull_()) { const lock = LockService.getScriptLock(); if (lock.tryLock(100)) { try { if (stkNeedFull_()) { stkSync_("App"); stkLogChanges_(); } } finally { lock.releaseLock(); } } }
  return stkRev_();
}
// have = the change number the phone already has; only changes since then are sent.
function crmStockGet(t, have, mt) {
  if (t && typeof t === "object") { have = t.have; mt = t.mt; t = t.t; }
  const me = auth_(t), lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error("Stock is busy - trying again in a few seconds");
  try { const D = stkData_(); stkLogChanges_(); return stkModel_(me, D, have, mt); } finally { lock.releaseLock(); }
}

function crmStockAudit(t) {
  const me = auth_(t), pm = stkPerm_(me); if (!pm.keeper) throw new Error("Only the admin and store keepers can see stock activity");
  const R = stkRows_("A"), chk = stkAuditCheck_(R);
  let L = R.map(r => r.v); if (!pm.admin) L = L.filter(v => v.Field !== "Rate" && v.Field !== "Stock Value").map(v => { const o = Object.assign({}, v); if (/Rate/.test(o["Previous Value"] + o["New Value"])) { o["Previous Value"] = ""; o["New Value"] = ""; } return o; });
  return { L: L.slice(-3000).reverse(), ok: chk.ok, badRow: chk.row || 0, full: pm.admin };
}

// One save of any kind. o = {t, have, op:{id, k, ...}}. Every save is checked, locked, written once and audited.
function crmStockSave(o) {
  const me = auth_(o && o.t), op = (o && o.op) || {}, pm = stkPerm_(me), who = stkWho_(me.u);
  if (!op.id || !op.k) throw new Error("Missing request");
  const cache = CacheService.getScriptCache(), ck = "stkop_" + op.id, done = cache.get(ck);
  if (done) return JSON.parse(done);                                   // the same request again (double tap / retry)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error("Stock is busy - please try again in a few seconds");
  try {
    if (cache.get(ck)) return JSON.parse(cache.get(ck));
    const D = stkData_(); stkLogChanges_();
    const t0 = stkNow_(), au = [], need = (ok, msg) => { if (!ok) throw new Error(msg || "You are not allowed to do this"); };
    const P = D.P, find = id => P.find(r => r.v["Product ID"] === id), A = x => au.push(Object.assign({ who, u: me.u, src: "App", t: t0 }, x));
    const addP = v => { const r = stkAppend_("P", v); P.push({ r, v }); return r; };
    let res = { ok: true };
    if (op.k === "move") {
      need(pm.keeper, "Only the admin or a store keeper can record stock");
      if (D.mids[op.mid]) { /* already saved (retry) */ }
      else {
        const r = find(op.pid); need(r && r.v.Status === "Active", "This product is not available"); const sign = STK_TYPES[op.type]; need(sign !== undefined, "Unknown movement type");
        const q = Number(op.qty), hadQ = typeof r.v.Quantity === "number", b = hadQ ? r.v.Quantity : 0;
        need(isFinite(q) && (sign === 0 ? q >= 0 : q > 0), "Enter a valid quantity");
        if ((op.type === "Damage / Loss" || op.type === "Adjustment") && !String(op.rm || "").trim()) need(false, "Please write the reason");
        const d = sign === 0 ? q - b : sign * q; if (b + d < 0) need(false, "Only " + b + " " + (r.v.Unit || "") + " in stock - you cannot " + (op.type === "Stock Issued" ? "issue" : "remove") + " " + q + ".");
        const date = /^\d{4}-\d{2}-\d{2}$/.test(op.date || "") ? op.date : stkToday_(); need(date <= stkToday_(), "The date cannot be in the future");
        const mv = { "Movement ID": op.mid, Date: date, Type: op.type, "Product ID": op.pid, "Product Name": r.v["Product Name"], Quantity: sign === 0 ? d : q, "Balance After": b + d, Store: op.store || "main", "Project ID": op.lid || "", Project: op.ln || "", "Sub-project": op.sn || "", "Furnishing Status": op.type === "Stock Issued" ? (op.fs || D.S.statuses[0] || "") : "", Remarks: op.rm || "", Staff: who, "User ID": me.u, "Recorded At": t0, Source: "App" };
        const mr = stkAppend_("M", mv); D.M.push({ r: mr, v: mv }); D.mids[op.mid] = 1; (STK_MEMO._newM = STK_MEMO._newM || []).push(mv);
        r.v.Quantity = b + d; r.v["Stock Value"] = typeof r.v.Rate === "number" ? Math.round(r.v.Quantity * r.v.Rate * 100) / 100 : null; r.v["Updated At"] = t0; r.v["Updated By"] = who; stkWrite_("P", r.r, r.v);
        A({ act: stkActOf_(op.type), pid: op.pid, pn: r.v["Product Name"], f: "Quantity", o: hadQ ? b : "", n: b + d, rm: [op.type + " " + (sign === 0 ? "(counted " + q + ")" : q), op.ln, op.sn, op.rm].filter(Boolean).join(" · ") });
        res.balance = b + d;
      }
    } else if (op.k === "status") {
      need(pm.keeper); let r = D.M.find(x => x.v["Movement ID"] === op.mid);
      if (!r && D.mids[op.mid]) { const all = stkRows_("M"); r = all.find(x => x.v["Movement ID"] === op.mid); }
      need(r, "Transaction not found"); const o0 = r.v["Furnishing Status"];
      if (o0 !== op.fs) { r.v["Furnishing Status"] = op.fs; stkWrite_("M", r.r, r.v); (STK_MEMO._stM = STK_MEMO._stM || []).push(r.v); STK_MEMO._meta = 1; A({ act: "Status Updated", pid: r.v["Product ID"], pn: r.v["Product Name"], f: "Furnishing status (" + op.mid + ")", o: o0, n: op.fs, rm: op.rm || r.v.Project || "" }); }
    } else if (op.k === "padd") {
      need(pm.keeper); const v = stkClean_(op.v, pm), k = stkKey_(v["Product Name"]) + "|" + stkKey_(v.Specification); need(v["Product Name"], "Product Name is required");
      const dup = P.find(r => r.v["Product ID"] && r.v.Status !== "Duplicate?" && stkKey_(r.v["Product Name"]) + "|" + stkKey_(r.v.Specification) === k);
      if (dup) need(false, "DUPLICATE:" + dup.v["Product ID"] + ":This product already exists (" + dup.v["Product Name"] + (dup.v.Specification ? ", " + dup.v.Specification : "") + (dup.v.Status === "Deleted" ? " – deleted, ask the admin to restore it" : "") + ").");
      let np = +(props_().getProperty("STK_NP") || 0) + 1; props_().setProperty("STK_NP", String(np));
      const sl = P.reduce((m, r) => Math.max(m, r.v["Serial No"] || 0), 0) + 1;
      Object.assign(v, { "Product ID": "P-" + String(np).padStart(6, "0"), "Serial No": sl, Quantity: null, "Stock Value": null, Status: "Active", "Updated At": t0, "Updated By": who });
      addP(v); A({ act: "Added", pid: v["Product ID"], pn: v["Product Name"], n: stkDesc_(v), rm: op.rm || "" }); res.pid = v["Product ID"];
    } else if (op.k === "pedit") {
      need(pm.keeper); const r = find(op.pid); need(r, "Product not found"); need(r.v.Status !== "Deleted", "This product is deleted");
      const cur = {}; STK_PF.forEach(f => cur[f] = r.v[f] === null || r.v[f] === undefined ? "" : String(r.v[f]));
      if (op.base) Object.keys(op.base).forEach(f => { if (f === "Rate" && !pm.admin) return; if ((cur[f] || "") !== (op.base[f] === null || op.base[f] === undefined ? "" : String(op.base[f]))) need(false, "CONFLICT:This product was changed by someone else or in Google Sheets. It has been reloaded – please check and save again."); });
      const v = stkClean_(op.v, pm); need(v["Product Name"], "Product Name is required");
      const k = stkKey_(v["Product Name"]) + "|" + stkKey_(v.Specification), dup = P.find(x => x !== r && x.v["Product ID"] && x.v.Status !== "Duplicate?" && stkKey_(x.v["Product Name"]) + "|" + stkKey_(x.v.Specification) === k);
      if (dup) need(false, "DUPLICATE:" + dup.v["Product ID"] + ":Another product already has this name and specification (" + dup.v["Product ID"] + ").");
      let any = false; Object.keys(v).forEach(f => { if (f === "Rate" && !pm.admin) return; const a = r.v[f] === null || r.v[f] === undefined ? "" : String(r.v[f]), b = v[f] === null || v[f] === undefined ? "" : String(v[f]); if (a !== b) { any = true; A({ act: f === "Category" || f === "Subcategory" ? "Category Changed" : "Edited", pid: op.pid, pn: v["Product Name"], f, o: a, n: b, rm: op.rm || "" }); r.v[f] = v[f]; } });
      if (any) { r.v["Stock Value"] = typeof r.v.Quantity === "number" && typeof r.v.Rate === "number" ? Math.round(r.v.Quantity * r.v.Rate * 100) / 100 : null; r.v["Updated At"] = t0; r.v["Updated By"] = who; stkWrite_("P", r.r, r.v); }
    } else if (op.k === "pdel" || op.k === "prestore") {
      need(pm.admin, "Only the admin can delete or restore products"); const r = find(op.pid);
      if (op.k === "pdel") { need(r && r.v.Status !== "Deleted", "Product not found"); need(String(op.rm || "").trim(), "Please write why it is deleted"); r.v.Status = "Deleted"; r.v["Updated At"] = t0; r.v["Updated By"] = who; stkWrite_("P", r.r, r.v); A({ act: "Deleted", pid: op.pid, pn: r.v["Product Name"], f: "Status", o: stkDesc_(r.v), n: "Deleted", rm: op.rm }); }
      else if (r) { need(r.v.Status === "Deleted", "This product is not deleted"); const k = stkKey_(r.v["Product Name"]) + "|" + stkKey_(r.v.Specification); need(!P.some(x => x !== r && x.v.Status === "Active" && stkKey_(x.v["Product Name"]) + "|" + stkKey_(x.v.Specification) === k), "An active product with the same name and specification exists - rename one first."); r.v.Status = "Active"; r.v["Updated At"] = t0; r.v["Updated By"] = who; stkWrite_("P", r.r, r.v); A({ act: "Restored", pid: op.pid, pn: r.v["Product Name"], f: "Status", o: "Deleted", n: "Active", rm: op.rm || "" }); }
      else { const x = (D.X || {})[op.pid]; need(x, "Deleted product not found"); const v = {}; STK_H.P.forEach(h => v[h] = x[h] === undefined ? "" : x[h]); v["Product ID"] = op.pid; v.Status = "Active"; v["Updated At"] = t0; v["Updated By"] = who;
        const all = stkRows_("M").filter(m => m.v["Product ID"] === op.pid && STK_TYPES[m.v.Type] !== undefined && typeof m.v.Quantity === "number"); v.Quantity = all.length ? all.reduce((a, m) => a + (STK_TYPES[m.v.Type] === 0 ? m.v.Quantity : STK_TYPES[m.v.Type] * m.v.Quantity), 0) : null;
        v["Stock Value"] = typeof v.Quantity === "number" && typeof v.Rate === "number" ? Math.round(v.Quantity * v.Rate * 100) / 100 : null; addP(v); delete D.X[op.pid]; A({ act: "Restored", pid: op.pid, pn: v["Product Name"], f: "Row restored", o: "Deleted in Google Sheet", n: "Active", rm: op.rm || "" }); }
    } else if (op.k === "cat") {
      need(pm.admin, "Only the admin can change categories"); const C = D.C, nm = String(op.n || "").trim(), cat = String(op.cat || "").trim();
      if (op.a === "add") { need(nm, "Type a name"); need(!C.some(r => r.v.Status !== "Deleted" && stkKey_(r.v.Category) === stkKey_(op.sub ? cat : nm) && stkKey_(r.v.Subcategory) === (op.sub ? stkKey_(nm) : "")), "It already exists"); let nc = +(props_().getProperty("STK_NC") || 0) + 1; props_().setProperty("STK_NC", String(nc)); const v = { ID: (op.sub ? "S-" : "C-") + String(nc).padStart(4, "0"), Category: op.sub ? cat : nm, Subcategory: op.sub ? nm : "", Status: "Active" }; C.push({ r: stkAppend_("C", v), v }); STK_MEMO._meta = 1; A({ act: "Category Changed", f: op.sub ? "Subcategory added" : "Category added", n: op.sub ? cat + " › " + nm : nm, rm: op.rm || "" }); }
      else if (op.a === "ren") {
        const r = C.find(x => x.v.ID === op.cid); need(r && nm, "Not found"); const isSub = !!r.v.Subcategory, old = isSub ? r.v.Subcategory : r.v.Category;
        need(!C.some(x => x !== r && x.v.Status !== "Deleted" && (isSub ? stkKey_(x.v.Category) === stkKey_(r.v.Category) && stkKey_(x.v.Subcategory) === stkKey_(nm) : !x.v.Subcategory && stkKey_(x.v.Category) === stkKey_(nm))), "That name is already used");
        if (isSub) { r.v.Subcategory = nm; stkWrite_("C", r.r, r.v); P.forEach(p => { if (stkKey_(p.v.Category) === stkKey_(r.v.Category) && p.v.Subcategory === old) { p.v.Subcategory = nm; stkWrite_("P", p.r, p.v); } }); }
        else { C.forEach(x => { if (x.v.Category === old) { x.v.Category = nm; stkWrite_("C", x.r, x.v); } }); P.forEach(p => { if (p.v.Category === old) { p.v.Category = nm; stkWrite_("P", p.r, p.v); } }); }
        A({ act: "Category Changed", f: isSub ? "Subcategory renamed" : "Category renamed", o: old, n: nm, rm: op.rm || "" });
      } else if (op.a === "del") {
        const r = C.find(x => x.v.ID === op.cid); need(r, "Not found"); const isSub = !!r.v.Subcategory, nm0 = isSub ? r.v.Category + " › " + r.v.Subcategory : r.v.Category; let moved = 0;
        if (isSub) { P.forEach(p => { if (p.v.Category === r.v.Category && p.v.Subcategory === r.v.Subcategory) { p.v.Subcategory = ""; stkWrite_("P", p.r, p.v); moved++; } }); r.v.Status = "Deleted"; stkWrite_("C", r.r, r.v); }
        else { const to = String(op.to || "").trim(); const users = P.filter(p => p.v.Category === r.v.Category); if (users.length) need(to && to !== r.v.Category && C.some(x => !x.v.Subcategory && x.v.Category === to && x.v.Status !== "Deleted"), "Choose the category to move its " + users.length + " product(s) to"); users.forEach(p => { p.v.Category = to; p.v.Subcategory = ""; stkWrite_("P", p.r, p.v); moved++; }); C.forEach(x => { if (x.v.Category === r.v.Category) { x.v.Status = "Deleted"; stkWrite_("C", x.r, x.v); } }); }
        A({ act: "Category Changed", f: isSub ? "Subcategory deleted" : "Category deleted", o: nm0, n: moved ? moved + " product(s) moved" + (op.to ? " to " + op.to : " (subcategory cleared)") : "", rm: op.rm || "" });
      }
    } else if (op.k === "settings") {
      need(pm.admin, "Only the admin can change stock settings"); const S0 = D.S;
      ["keepers", "units", "statuses"].forEach(k => { if (op[k] && JSON.stringify(op[k]) !== JSON.stringify(S0[k])) { stkSetSetting_(k, op[k]); A({ act: "Settings Changed", f: k === "keepers" ? "Store keepers" : k === "units" ? "Units" : "Furnishing statuses", o: k === "keepers" ? (S0[k] || []).map(stkWho_).join(", ") : (S0[k] || []).join(", "), n: k === "keepers" ? op[k].map(stkWho_).join(", ") : op[k].join(", ") }); } });
      D.S = stkSettings_();
    } else if (op.k === "import") {
      need(pm.admin, "Only the admin can import products"); const L = op.list || []; let added = 0, skipped = 0, np = +(props_().getProperty("STK_NP") || 0), sl = P.reduce((m, r) => Math.max(m, r.v["Serial No"] || 0), 0);
      const keys = {}; P.forEach(r => { if (r.v["Product ID"]) keys[stkKey_(r.v["Product Name"]) + "|" + stkKey_(r.v.Specification)] = 1; });
      const rows = [], vs = [];
      L.forEach(x => { const v = stkClean_(x, pm), k = stkKey_(v["Product Name"]) + "|" + stkKey_(v.Specification); if (!v["Product Name"] || keys[k]) { skipped++; return; } keys[k] = 1; np++; sl++; Object.assign(v, { "Product ID": "P-" + String(np).padStart(6, "0"), "Serial No": sl, Quantity: null, "Stock Value": null, Status: "Active", "Updated At": t0, "Updated By": who }); rows.push(stkRowArr_("P", v)); vs.push(v); added++; stkChg_(v["Product ID"]); A({ act: "Added", pid: v["Product ID"], pn: v["Product Name"], n: stkDesc_(v), rm: "Imported" + (op.src ? " from " + op.src : "") }); });
      props_().setProperty("STK_NP", String(np));
      if (rows.length) { const sh = stkSheet_("P"), r0 = Math.max(sh.getLastRow(), 1) + 1, rg = sh.getRange(r0, 1, rows.length, STK_H.P.length); rg.setNumberFormats(rows.map(() => STK_H.P.map(h => STK_NUM[h] ? "General" : "@"))); rg.setValues(rows); vs.forEach((v, i) => P.push({ r: r0 + i, v })); }
      res.added = added; res.skipped = skipped;
    } else need(false, "Unknown request");
    stkAudit_(au); stkSnapPatch_(D); stkLogChanges_();
    res.data = stkModel_(me, D, o.have, o.mt);
    try { cache.put(ck, JSON.stringify({ ok: true, again: true, balance: res.balance, pid: res.pid, added: res.added, skipped: res.skipped }), 21600); } catch (e) {}
    return res;
  } finally { lock.releaseLock(); }
}
function stkClean_(x, pm) {
  x = x || {}; const v = {}, s = k => String(x[k] === undefined || x[k] === null ? "" : x[k]).trim();
  ["Product Name", "Product Code", "Category", "Subcategory", "Specification", "Description", "Unit"].forEach(k => v[k] = s(k));
  const ml = stkNum_(x["Min Level"]); v["Min Level"] = ml === null || !isFinite(ml) || ml < 0 ? null : ml;
  if (pm.admin) { const rt = stkNum_(x.Rate); v.Rate = rt === null || !isFinite(rt) || rt < 0 ? null : rt; }
  return v;
}

/* ================= one-time setup and the change watcher ================= */
function setupStockSync() {
  ["P", "C", "M", "S", "A", "N"].forEach(stkSheet_);
  if (!stkRows_("C").length) { let nc = 0; STK_CATS.forEach(([c, subs]) => { nc++; stkAppend_("C", { ID: "C-" + String(nc).padStart(4, "0"), Category: c, Subcategory: "", Status: "Active" }); subs.forEach(s => { nc++; stkAppend_("C", { ID: "S-" + String(nc).padStart(4, "0"), Category: c, Subcategory: s, Status: "Active" }); }); }); props_().setProperty("STK_NC", String(nc)); }
  const S = stkSettings_(); if (!stkRows_("S").length) { stkSetSetting_("units", S.units); stkSetSetting_("statuses", S.statuses); stkSetSetting_("stores", S.stores); stkSetSetting_("keepers", []); }
  stkProtect_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "stkOnChange").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("stkOnChange").forSpreadsheet(ss_()).onChange().create();
  const lock = LockService.getScriptLock(); lock.waitLock(30000); try { stkSync_("App"); stkLogChanges_(); } finally { lock.releaseLock(); }
  Logger.log("Stock & Store is ready. Sheets: " + Object.keys(STK).map(k => STK[k]).join(", ") + ". Direct edits in these sheets are now recorded in Stock_Audit.");
}
// Runs by itself when someone edits the spreadsheet directly. Only reacts to the Stock sheets.
function stkOnChange(e) {
  let who = ""; try { who = (e && e.user && e.user.getEmail && e.user.getEmail()) || ""; } catch (x) {}
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return;
  try { stkSync_("Google Sheet", who); stkLogChanges_(); } finally { lock.releaseLock(); }
}
