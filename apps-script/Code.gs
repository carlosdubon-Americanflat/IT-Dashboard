// ═══════════════════════════════════════════════════════════════════════════
// American Flat — IT SaaS Dashboard: Google Apps Script Data Feed
// ───────────────────────────────────────────────────────────────────────────
// REVISION 2026-06-03 (reconciliation fix):
//   1) 2026 YTD KPI (ytd26) now = sum of per-app col F on "IT Binder Details",
//      so the headline KPI MATCHES the app table, dept charts, and top-5.
//   2) monthly25 is now read LIVE from "Softwares and Subscriptions - 2025"
//      instead of a hardcoded array, so the 2025 trend reflects the sheet.
//   3) readMonthly2026 now reads ALL month columns dynamically (and drops the
//      trailing row-total) instead of being hard-capped at 4 months, so the
//      monthly chart + projection pick up new months automatically.
//   4) Projection logic (avg26 / proj26) is otherwise unchanged — a run-rate
//      from the "Monthly Spend" tab, smoother than annualizing lumpy Q1
//      annual prepayments.
//   5) Each app now also returns admin (col H), notes (col I, holds seat
//      pricing) and users (col J, seat count) for display on the dashboard.
// ═══════════════════════════════════════════════════════════════════════════

function doGet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = buildDashboardData(ss);
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: e.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function buildDashboardData(ss) {
  const apps      = readApps(ss);
  const monthly26 = readMonthly2026(ss);
  const actMonths = monthly26.filter(v => v > 0).length || 1;
  const s25ytdMap = readS25YTD(ss, actMonths);
  const monthly25 = readMonthly2025(ss);          // ← NEW: live, replaces hardcoded array
  const data = computeAll(apps, monthly26, monthly25, s25ytdMap);
  data.latestMonth = readLatestMonthByApp(ss);    // ← NEW: per-app breakdown of the latest month
  data.monthlyByApp = readMonthlyByApp(ss);       // ← NEW 2026-10-01: per-app, every month
  return data;
}

// ── Read per-app spend for the latest month with data, from "Monthly Spend" ──
//   Returns { label: 'May', apps: [{name, value}, ...] } sorted high→low.
function readLatestMonthByApp(ss) {
  const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const sheet  = ss.getSheetByName('Monthly Spend');
  const values = sheet.getDataRange().getValues();

  // Locate the Grand Total row to learn which month columns have data.
  let gt = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().indexOf('Grand Total') !== -1) { gt = i; break; }
  }
  if (gt < 0) return { label: '', apps: [] };

  // Month columns start at index 2 (col C); drop the trailing row-total column.
  const grow = values[gt];
  const cols = [];
  for (let c = 2; c < grow.length; c++) {
    const v = grow[c];
    if (v === '' || v === null || v === undefined) continue;
    cols.push(c);
  }
  if (cols.length >= 2) {
    const last = parseAmt(grow[cols[cols.length - 1]]);
    const rest = cols.slice(0, -1).reduce(function (s, c) { return s + parseAmt(grow[c]); }, 0);
    if (Math.abs(last - rest) < 1) cols.pop();
  }

  // Latest month column whose Grand Total is > 0.
  let monthCol = -1;
  cols.forEach(function (c) { if (parseAmt(grow[c]) > 0) monthCol = c; });
  if (monthCol < 0) return { label: '', apps: [] };
  const monthIdx = monthCol - 2;

  const apps = [];
  for (let i = 0; i < values.length; i++) {
    if (i === gt) continue;
    const grp  = String(values[i][0] || '').trim();
    const name = String(values[i][1] || '').trim();
    if (!name) continue;
    if (grp.indexOf('Total') !== -1 || name.indexOf('Total') !== -1) continue;
    if (name === '*End of Records') break;
    const val = parseAmt(values[i][monthCol]);
    if (val > 0) apps.push({ name: name, value: round2(val) });
  }
  apps.sort(function (a, b) { return b.value - a.value; });
  return { label: MONTHS[monthIdx] || ('Month ' + (monthIdx + 1)), apps: apps };
}

// ── Per-app spend for every month, from the "Monthly Spend" pivot ──────────
//   REVISION 2026-10-01: feeds the dashboard's Monthly Breakdown table.
//   Returns { months: ['Jan',...], rows: [{dept, name, values:[...]}] } where
//   dept is carried down from col A (the pivot only prints it on a group's
//   first row) and only months with a Grand Total > 0 are included.
function readMonthlyByApp(ss) {
  const sheet  = ss.getSheetByName('Monthly Spend');
  const values = sheet.getDataRange().getValues();
  if (!values.length) return { months: [], rows: [] };

  // Month columns = header cells from col C until "Total Spend" / blank.
  const header = values[0];
  const monthCols = [];
  for (let c = 2; c < header.length; c++) {
    const h = String(header[c] || '').trim();
    if (!h || h.toLowerCase().indexOf('total') !== -1) break;
    monthCols.push({ col: c, label: h.replace(/^SUM of\s+/i, '').slice(0, 3) });
  }

  let gt = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().indexOf('Grand Total') !== -1) { gt = i; break; }
  }
  const keep = monthCols.filter(function (m) { return gt < 0 || parseAmt(values[gt][m.col]) > 0; });

  const rows = [];
  let dept = '';
  for (let i = 1; i < values.length; i++) {
    if (i === gt) break;
    const grp  = String(values[i][0] || '').trim();
    const name = String(values[i][1] || '').trim();
    if (grp && grp.indexOf('Total') === -1) dept = grp;
    if (!name || grp.indexOf('Total') !== -1 || name.indexOf('Total') !== -1) continue;
    rows.push({
      dept: dept === 'Ops' ? 'Operations' : dept,
      name: name,
      values: keep.map(function (m) { return round2(parseAmt(values[i][m.col])); })
    });
  }
  return { months: keep.map(function (m) { return m.label; }), rows: rows };
}

// ── Read app list from "IT Binder Details" ───────────────────────────────
function readApps(ss) {
  const sheet  = ss.getSheetByName('IT Binder Details');
  const values = sheet.getDataRange().getValues();
  const validStatuses = new Set(['Active', 'Reduced', 'Cancelled']);
  const apps = [];

  for (let i = 1; i < values.length; i++) {
    const row    = values[i];
    const name   = String(row[0] || '').trim();
    const dept   = String(row[1] || '').trim();
    const type   = String(row[2] || '').trim();
    const renewal= String(row[3] || '').trim();
    const status = String(row[6] || '').trim();
    const admin  = String(row[7] || '').trim();   // Col H — Admin User Access
    const notes  = String(row[8] || '').trim();   // Col I — notes (holds seat pricing)
    const users  = String(row[9] || '').trim();   // Col J — Users / seat count

    if (!name || !validStatuses.has(status)) continue;
    if (name.toLowerCase().includes('grand total')) continue;
    if (name.toLowerCase().includes('department'))  continue;
    if (name === '*End of Records') break;

    apps.push({
      name, dept, type, renewal,
      s25: parseAmt(row[4]),
      s26: parseAmt(row[5]),
      status, admin, notes, users
    });
  }
  return apps;
}

// ── Read per-app Jan–actMonths 2025 spend from monthly sheet ─────────────
function readS25YTD(ss, actMonths) {
  const sheet = ss.getSheetByName('Softwares and Subscriptions - 2025');
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const row  = values[i];
    const name = String(row[4] || '').trim();   // Column E
    if (!name) continue;
    if (name.toLowerCase().includes('grand total')) continue;
    if (name.toLowerCase().includes('department'))  continue;
    if (name === '*End of Records') break;

    // Columns K–V = indices 10–21 (Jan–Dec)
    let ytd = 0;
    for (let m = 0; m < actMonths; m++) {
      ytd += parseAmt(row[10 + m]);
    }
    map[name] = round2((map[name] || 0) + ytd);
  }
  return map;
}

// ── NEW: Read full-year 2025 monthly totals from "Softwares and Subscriptions - 2025" ──
//   Sums each month (cols K–V = indices 10–21, Jan–Dec) across every app row.
//   Replaces the old hardcoded monthly25 array so the 2025 trend tracks the sheet.
function readMonthly2025(ss) {
  const sheet = ss.getSheetByName('Softwares and Subscriptions - 2025');
  if (!sheet) return [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];

  const values = sheet.getDataRange().getValues();
  const totals = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];

  for (let i = 1; i < values.length; i++) {
    const row  = values[i];
    const name = String(row[4] || '').trim();   // Column E
    if (!name) continue;
    if (name.toLowerCase().includes('grand total')) continue;
    if (name.toLowerCase().includes('department'))  continue;
    if (name === '*End of Records') break;

    for (let m = 0; m < 12; m++) {
      totals[m] = round2(totals[m] + parseAmt(row[10 + m]));
    }
  }
  return totals;
}

// ── Read monthly 2026 grand totals from "Monthly Spend" ─────────────────
//   Reads ALL month columns dynamically (col C onward) and auto-drops the
//   trailing row-total column, so it adapts as new months are added.
function readMonthly2026(ss) {
  const sheet  = ss.getSheetByName('Monthly Spend');
  const values = sheet.getDataRange().getValues();

  for (let i = 0; i < values.length; i++) {
    const label = String(values[i][0] || '').trim();
    if (label.includes('Grand Total')) {
      const row = values[i];
      // Collect every populated numeric cell from column C (index 2) onward.
      const nums = [];
      for (let c = 2; c < row.length; c++) {
        const cell = row[c];
        if (cell === '' || cell === null || cell === undefined) continue;
        nums.push(parseAmt(cell));
      }
      // The last cell is the row TOTAL (≈ sum of the months). Drop it if so.
      if (nums.length >= 2) {
        const last    = nums[nums.length - 1];
        const sumRest = nums.slice(0, -1).reduce((s, v) => s + v, 0);
        if (Math.abs(last - sumRest) < 1) nums.pop();
      }
      return nums;
    }
  }
  return [];
}

// ── Compute all dashboard metrics ────────────────────────────────────────
function computeAll(apps, monthly26, monthly25, s25ytdMap) {
  const actMonths   = monthly26.filter(v => v > 0).length || 1;
  const total25     = apps.reduce((s, a) => s + a.s25, 0);

  // 2026 YTD = complete per-app cumulative spend (matches the app table & dept charts).
  const ytd26       = apps.reduce((s, a) => s + a.s26, 0);   // ← CHANGED (was Σ monthly26)

  // Projection run-rate stays on the smoother "Monthly Spend" detail.
  const monthlyYtd  = monthly26.reduce((s, v) => s + v, 0);
  const avg26       = monthlyYtd / actMonths;
  const proj26      = avg26 * 12;
  const savedVs25   = total25 - proj26;

  const activeCount    = apps.filter(a => a.status === 'Active').length;
  const reducedCount   = apps.filter(a => a.status === 'Reduced').length;
  const cancelledCount = apps.filter(a => a.status === 'Cancelled').length;

  const dept25 = {}, dept26 = {};
  apps.forEach(a => {
    const d = a.dept || 'Other';
    dept25[d] = (dept25[d] || 0) + a.s25;
    dept26[d] = (dept26[d] || 0) + a.s26;
  });

  const top10_25 = [...apps]
    .sort((a, b) => b.s25 - a.s25)
    .slice(0, 10)
    .map(a => ({ name: a.name, value: round2(a.s25) }));

  const top5_26 = [...apps]
    .filter(a => a.s26 > 0)
    .sort((a, b) => b.s26 - a.s26)
    .slice(0, 5)
    .map(a => ({ name: a.name, value: round2(a.s26) }));

  return {
    generatedAt: new Date().toISOString(),
    kpis: {
      total25:       round2(total25),
      ytd26:         round2(ytd26),
      proj26:        round2(proj26),
      avg26:         round2(avg26),
      savedVs25:     round2(savedVs25),
      activeCount,
      reducedCount,
      cancelledCount
    },
    monthly26,
    monthly25,                                        // ← CHANGED (was hardcoded array)
    dept25,
    dept26,
    top10_25,
    top5_26,
    apps: apps
      .sort((a, b) => b.s25 - a.s25)
      .map(a => ({
        name: a.name, dept: a.dept, renewal: a.renewal,
        s25:    round2(a.s25),
        s26:    round2(a.s26),
        s25ytd: round2(s25ytdMap[a.name] || 0),
        status: a.status,
        admin:  a.admin,            // Admin User Access (col H)
        notes:  a.notes,            // notes / seat pricing (col I)
        users:  a.users             // seat count (col J)
      }))
  };
}

function parseAmt(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  return parseFloat(String(v).replace(/[$,\s]/g, '')) || 0;
}

function round2(n) { return Math.round(n * 100) / 100; }