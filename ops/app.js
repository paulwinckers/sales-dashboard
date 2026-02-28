const CONFIG = {
  SHEET_API_URL: "https://script.google.com/macros/s/AKfycbxemOcHaO8jJL2JNvr6G3INrHOSahH3-1QYcsrb5IV19DG77lPUPtDkco_s9r8RFwmI/exec", // /exec
  API_KEY: "",
  SHEET_TAB_NAME: "OpsDailyLog",
  CALENDAR_URL: "../data/calendar.csv",
  TARGETS_URL: "../data/targets.csv"
};

/* ---------------- helpers ---------------- */
const el = (id) => document.getElementById(id);

const safeNum = (v) => {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
};
const fmt1 = (n) => (Number.isFinite(n) ? n.toFixed(1) : "—");
const fmt0 = (n) => (Number.isFinite(n) ? String(Math.round(n)) : "—");

function toast(msg){
  const t = el("toast");
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(()=>t.classList.remove("show"), 2200);
}

function ymd(d){
  const pad = (x)=> String(x).padStart(2,"0");
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
}

function monthKeyFromDate(d){
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
}

function monthStart(selYmd){
  const d=new Date(selYmd+"T00:00:00");
  return ymd(new Date(d.getFullYear(), d.getMonth(), 1));
}

function yearStart(selYmd){
  const d=new Date(selYmd+"T00:00:00");
  return `${d.getFullYear()}-01-01`;
}
function yearEnd(selYmd){
  const d=new Date(selYmd+"T00:00:00");
  return `${d.getFullYear()}-12-31`;
}

function setBoxStatus(node, state){
  node.classList.remove("good","warn","bad");
  if(state) node.classList.add(state);
}

function statusByPct(pct){
  if(pct==null) return null;
  if(pct<0.95) return "bad";
  if(pct<1.0) return "warn";
  return "good";
}

function tagByPct(pct){
  // prior rule: red if more than 5% short, yellow if within 5%, green if hit/exceed
  if(pct==null) return { cls:"", txt:"No budget" };
  if(pct < 0.95) return { cls:"bad", txt:"Off Target" };
  if(pct < 1.0)  return { cls:"warn", txt:"Close" };
  return { cls:"good", txt:"On Target" };
}

function escapeHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function hasValue(v){
  return v !== null && v !== undefined && String(v).trim() !== "";
}

function isOffDay(dayType){
  const dt = String(dayType || "").trim().toLowerCase();
  return dt === "weekend" || dt === "stat";
}

/* ---------------- CSV parsing ---------------- */
function parseDelimited(text){
  const lines = text.split(/\r?\n/).filter(l=>l.trim().length);
  if(!lines.length) return { headers:[], rows:[] };
  const delim = lines[0].includes("\t") ? "\t" : ",";
  const splitLine = (line)=>{
    const out=[]; let cur="", inQ=false;
    for(let i=0;i<line.length;i++){
      const ch=line[i];
      if(ch === '"'){
        if(inQ && line[i+1] === '"'){ cur+='"'; i++; }
        else inQ=!inQ;
      } else if(ch===delim && !inQ){
        out.push(cur); cur="";
      } else cur += ch;
    }
    out.push(cur);
    return out.map(x=>x.trim());
  };
  const headers = splitLine(lines[0]).map(h=>h.replace(/\uFEFF/g,""));
  const rows = [];
  for(let i=1;i<lines.length;i++) rows.push(splitLine(lines[i]));
  return { headers, rows };
}

/* ---------------- targets.csv fallback ---------------- */
let targets = {};
let targetsLoaded = false;

function toMonthKey(raw){
  const s=(raw??"").toString().trim();
  if(!s) return null;

  const monthMap={jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};

  let m=s.match(/^(\d{2})[-\s]?([A-Za-z]{3,})\.?$/);  // 26-Jan
  if(m){
    const yy=String(m[1]).padStart(2,"0");
    const mon=m[2].slice(0,3).toLowerCase();
    const mm=monthMap[mon];
    if(mm) return `20${yy}-${mm}`;
  }

  m=s.match(/^([A-Za-z]{3,})[-\s]?(\d{2})$/);        // Jan-26
  if(m){
    const mon=m[1].slice(0,3).toLowerCase();
    const yy=String(m[2]).padStart(2,"0");
    const mm=monthMap[mon];
    if(mm) return `20${yy}-${mm}`;
  }

  const m2=s.match(/^(\d{4})-(\d{2})$/);
  if(m2) return `${m2[1]}-${m2[2]}`;

  const d=new Date(s);
  if(!isNaN(d)) return monthKeyFromDate(d);

  return null;
}

function parseTargets(text){
  const { headers, rows } = parseDelimited(text);
  const lower = headers.map(h=>h.toLowerCase());

  const monthIdx = lower.findIndex(h => h==="month" || h.includes("month"));
  const constIdx = lower.findIndex(h => h.includes("construction") && h.includes("hours"));
  const maintIdx = lower.findIndex(h => h.includes("maintenance") && h.includes("hours"));
  if(monthIdx < 0 || (constIdx < 0 && maintIdx < 0)) return {};

  const parseHours=(v)=>{
    const s=(v??"").toString().replaceAll(",","").trim();
    if(!s || s==="-" ) return 0;
    const n=Number(s);
    return Number.isFinite(n) ? n : 0;
  };

  const out = {};
  for(const row of rows){
    const mk = toMonthKey(row[monthIdx]);
    if(!mk) continue;
    out[mk] = {
      maintenanceMonthly: maintIdx>=0 ? parseHours(row[maintIdx]) : 0,
      constructionMonthly: constIdx>=0 ? parseHours(row[constIdx]) : 0
    };
  }
  return out;
}

/* ---------------- calendar.csv ---------------- */
let calendar = {};
let calendarLoaded = false;

function parseCalendar(text){
  const { headers, rows } = parseDelimited(text);
  const lower = headers.map(h => h.toLowerCase());

  const dateIdx = lower.findIndex(h => h === "date" || h.includes("date"));
  const maintIdx = lower.findIndex(h => h.includes("maintenancehours") || (h.includes("maintenance") && h.includes("hours")));
  const constIdx = lower.findIndex(h => h.includes("constructionhours") || (h.includes("construction") && h.includes("hours")));
  if (dateIdx < 0 || (maintIdx < 0 && constIdx < 0)) return {};

  const toDateKey = (v) => {
    const s = (v ?? "").toString().trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    const d = new Date(s);
    if (!isNaN(d)) return ymd(d);
    return null;
  };

  const toNum = (v) => {
    const s = (v ?? "").toString().replaceAll(",", "").trim();
    if (!s || s === "-") return 0;
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  };

  const out = {};
  for (const row of rows){
    const key = toDateKey(row[dateIdx]);
    if (!key) continue;
    out[key] = {
      maintTarget: maintIdx >= 0 ? toNum(row[maintIdx]) : 0,
      constTarget: constIdx >= 0 ? toNum(row[constIdx]) : 0
    };
  }
  return out;
}

function getCalendarTargetForDate(dateKey){
  const r = calendar[dateKey];
  if (!r) return { maint: 0, cons: 0, total: 0, has: false };
  const maint = safeNum(r.maintTarget);
  const cons  = safeNum(r.constTarget);
  return { maint, cons, total: maint + cons, has: true };
}

function sumCalendarTargetsByDivision(startYmd, endYmd){
  let maint = 0, cons = 0;
  for (const [k, r] of Object.entries(calendar)){
    if (k < startYmd || k > endYmd) continue;
    maint += safeNum(r.maintTarget);
    cons  += safeNum(r.constTarget);
  }
  return { maint, cons, total: maint + cons };
}

/* ---------------- daily from Google Sheet ---------------- */
let daily = {};
let sheetLoadedCount = 0;

function getDayRec(dateKey){
  return daily[dateKey] || {
    date: dateKey,
    dayType:"",
    targetMaint:"",
    targetConst:"",
    actualMaint:"",
    actualConst:"",
    missedTickets:"",
    safetyIncidents:"",
    notes:""
  };
}

async function loadDailyFromGoogle(rangeStart, rangeEnd){
  if(!CONFIG.SHEET_API_URL || CONFIG.SHEET_API_URL.includes("PASTE_")){
    throw new Error("Missing Apps Script URL in CONFIG.SHEET_API_URL");
  }

  const u = new URL(CONFIG.SHEET_API_URL);
  u.searchParams.set("tab", CONFIG.SHEET_TAB_NAME);
  if(rangeStart) u.searchParams.set("start", rangeStart);
  if(rangeEnd) u.searchParams.set("end", rangeEnd);
  if(CONFIG.API_KEY) u.searchParams.set("key", CONFIG.API_KEY);

  const res = await fetch(u.toString(), { cache: "no-store" });
  if(!res.ok) throw new Error(`Google endpoint HTTP ${res.status}`);
  const data = await res.json();

  const map = {};
  const rows = Array.isArray(data.rows) ? data.rows : [];

  for(const r of rows){
    const date = String(r.Date || "").slice(0,10);
    if(!date.match(/^\d{4}-\d{2}-\d{2}$/)) continue;

    map[date] = {
      date,
      dayType: r.DayType ?? "",
      targetMaint: r.TargetMaint ?? "",
      targetConst: r.TargetConst ?? "",
      actualMaint: r.ActualMaint ?? "",
      actualConst: r.ActualConst ?? "",
      missedTickets: r.MissedTickets ?? "",
      safetyIncidents: r.SafetyIncidents ?? "",
      notes: r.Notes ?? ""
    };
  }

  daily = map;
  sheetLoadedCount = rows.length;
}

/* ---------------- budget logic ---------------- */
function computeFallbackDailyTargets(dateKey){
  const d = new Date(dateKey+"T00:00:00");
  const mk = monthKeyFromDate(d);
  const t = targets[mk] || { maintenanceMonthly:0, constructionMonthly:0 };

  // weekday-only fallback (rarely used since sheet has daily targets)
  const year = d.getFullYear();
  const month = d.getMonth();
  let wd = 0;
  const it = new Date(year, month, 1);
  while(it.getMonth()===month){
    const day = it.getDay();
    if(day!==0 && day!==6) wd++;
    it.setDate(it.getDate()+1);
  }

  const dm = wd>0 ? safeNum(t.maintenanceMonthly)/wd : 0;
  const dc = wd>0 ? safeNum(t.constructionMonthly)/wd : 0;
  return { maint: dm, cons: dc, total: dm+dc, source:"targets.csv avg" };
}

function getDayBudget(dateKey){
  const rec = getDayRec(dateKey);

  // Sheet daily targets (explicit, including 0)
  const sheetHasMaint = hasValue(rec.targetMaint);
  const sheetHasConst = hasValue(rec.targetConst);
  if(sheetHasMaint || sheetHasConst){
    const m = safeNum(rec.targetMaint);
    const c = safeNum(rec.targetConst);
    return { maint:m, cons:c, total:m+c, source:"Sheet" };
  }

  // Calendar fallback
  if(calendarLoaded){
    const c = getCalendarTargetForDate(dateKey);
    if(c.has) return { maint:c.maint, cons:c.cons, total:c.total, source:"Calendar" };
  }

  // targets.csv avg fallback
  return computeFallbackDailyTargets(dateKey);
}

function sumBudgetForRange(startYmd, endYmd){
  // Prefer summing sheet targets if any exist
  let anySheet = false;
  let sm=0, sc=0;

  for(let d=new Date(startYmd+"T00:00:00"); d<=new Date(endYmd+"T00:00:00"); d.setDate(d.getDate()+1)){
    const k = ymd(d);
    const rec = getDayRec(k);
    const sheetHasMaint = hasValue(rec.targetMaint);
    const sheetHasConst = hasValue(rec.targetConst);
    if(sheetHasMaint || sheetHasConst){
      anySheet = true;
      sm += safeNum(rec.targetMaint);
      sc += safeNum(rec.targetConst);
    }
  }
  if(anySheet) return { maint:sm, cons:sc, total:sm+sc, source:"Sheet" };

  if(calendarLoaded){
    const c = sumCalendarTargetsByDivision(startYmd, endYmd);
    return { ...c, source:"Calendar" };
  }

  // fallback: weekday averages
  let maint=0, cons=0;
  for(let d=new Date(startYmd+"T00:00:00"); d<=new Date(endYmd+"T00:00:00"); d.setDate(d.getDate()+1)){
    const day = d.getDay();
    if(day===0 || day===6) continue;
    const fb = computeFallbackDailyTargets(ymd(d));
    maint += fb.maint;
    cons  += fb.cons;
  }
  return { maint, cons, total: maint+cons, source:"targets.csv avg" };
}

function sumActualsForRange(startYmd, endYmd){
  let maint=0, cons=0;
  for(const [k, rec] of Object.entries(daily)){
    if(k < startYmd || k > endYmd) continue;
    maint += safeNum(rec.actualMaint);
    cons  += safeNum(rec.actualConst);
  }
  return { maint, cons, total: maint+cons };
}

/* ---------------- render log (workdays by default) ---------------- */
function renderLog(asOfYmd){
  const range = el("logRange").value;      // "14" | "30" | "90" | "ytd"
  const includeOff = el("includeOffDays").checked;

  const tbody = el("logTbody");
  tbody.innerHTML = "";

  const startYmd = (range === "ytd") ? yearStart(asOfYmd) : null;
  const needed = (range === "ytd") ? Infinity : parseInt(range, 10);

  let shown = 0;

  for(let d=new Date(asOfYmd+"T00:00:00"); ; d.setDate(d.getDate()-1)){
    const key = ymd(d);

    if(range === "ytd" && key < startYmd) break;

    const rec = getDayRec(key);

    // Default: skip Weekend & Stat
    if(!includeOff && isOffDay(rec.dayType)) continue;

    const b = getDayBudget(key);

    const aM = safeNum(rec.actualMaint);
    const aC = safeNum(rec.actualConst);
    const aT = aM + aC;

    const bM = safeNum(b.maint);
    const bC = safeNum(b.cons);
    const bT = bM + bC;

    const pct = bT>0 ? (aT/bT) : null;
    const tag = tagByPct(pct);

    const deltaHrs = aT - bT;
    const deltaPct = (bT>0) ? ((aT/bT)-1) : null;

    const tickets = safeNum(rec.missedTickets);
    const safety  = safeNum(rec.safetyIncidents);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="mono">${key}</td>
      <td><span class="tag ${tag.cls}">${tag.txt}</span></td>
      <td>${escapeHtml(rec.dayType ?? "")}</td>
      <td class="mono">${fmt1(bM)}</td>
      <td class="mono">${fmt1(bC)}</td>
      <td class="mono">${fmt1(bT)}</td>
      <td class="mono">${fmt1(aM)}</td>
      <td class="mono">${fmt1(aC)}</td>
      <td class="mono">${fmt1(aT)}</td>
      <td class="mono">${deltaHrs>=0?"+":""}${deltaHrs.toFixed(1)}</td>
      <td class="mono">${deltaPct==null ? "—" : `${(deltaPct*100).toFixed(1)}%`}</td>
      <td class="mono">${fmt0(tickets)}</td>
      <td class="mono">${fmt0(safety)}</td>
      <td>${escapeHtml(rec.notes ?? "")}</td>
    `;
    tbody.appendChild(tr);

    shown++;
    if(shown >= needed) break;
  }

  if(shown === 0){
    tbody.innerHTML = `<tr><td colspan="14" class="hint">No rows to display for this range/filter.</td></tr>`;
  }
}

/* ---------------- render summaries ---------------- */
function renderSummaries(asOfYmd){
  const view = el("kpiView").value;

  const mStart = monthStart(asOfYmd);
  const mBudget = sumBudgetForRange(mStart, asOfYmd);
  const mActual = sumActualsForRange(mStart, asOfYmd);

  const yStart = yearStart(asOfYmd);
  const yBudget = sumBudgetForRange(yStart, asOfYmd);
  const yActual = sumActualsForRange(yStart, asOfYmd);

  el("mtdMaintLine").textContent = `${fmt1(mActual.maint)} / ${fmt1(mBudget.maint)}`;
  el("mtdConstLine").textContent = `${fmt1(mActual.cons)} / ${fmt1(mBudget.cons)}`;
  el("ytdMaintLine").textContent = `${fmt1(yActual.maint)} / ${fmt1(yBudget.maint)}`;
  el("ytdConstLine").textContent = `${fmt1(yActual.cons)} / ${fmt1(yBudget.cons)}`;

  const pick = (obj, key) => key==="maint" ? obj.maint : key==="const" ? obj.cons : obj.total;

  // MTD
  const mA = pick(mActual, view);
  const mB = pick(mBudget, view);
  const mPct = mB>0 ? mA/mB : null;
  const mVar = mA - mB;

  setBoxStatus(el("mtdStatus"), statusByPct(mPct));
  el("mtdHeadline").textContent = (mB>0) ? `${fmt1(mA)} / ${fmt1(mB)} hrs` : "No budget";
  el("mtdBudget").textContent = fmt1(mB);
  el("mtdActual").textContent = fmt1(mA);
  el("mtdVar").textContent = (mB>0||mA>0) ? `${mVar>=0?"+":""}${mVar.toFixed(1)}` : "—";
  el("mtdPct").textContent = (mPct!=null) ? `${(mPct*100).toFixed(1)}%` : "—";
  el("mtdBar").style.width = `${Math.min(100, Math.max(0, (mPct||0)*100))}%`;

  // YTD
  const yA = pick(yActual, view);
  const yB = pick(yBudget, view);
  const yPct = yB>0 ? yA/yB : null;
  const yVar = yA - yB;

  setBoxStatus(el("ytdStatus"), statusByPct(yPct));
  el("ytdHeadline").textContent = (yB>0) ? `${fmt1(yA)} / ${fmt1(yB)} hrs` : "No budget";
  el("ytdBudget").textContent = fmt1(yB);
  el("ytdActual").textContent = fmt1(yA);
  el("ytdVar").textContent = (yB>0||yA>0) ? `${yVar>=0?"+":""}${yVar.toFixed(1)}` : "—";
  el("ytdPct").textContent = (yPct!=null) ? `${(yPct*100).toFixed(1)}%` : "—";
  el("ytdBar").style.width = `${Math.min(100, Math.max(0, (yPct||0)*100))}%`;
}

/* ---------------- loads ---------------- */
async function fetchText(url){
  const res = await fetch(url, { cache: "no-store" });
  if(!res.ok) throw new Error(`${url} HTTP ${res.status}`);
  return res.text();
}
async function loadCalendarFromRepo(){
  const text = await fetchText(CONFIG.CALENDAR_URL);
  calendar = parseCalendar(text);
  calendarLoaded = Object.keys(calendar).length > 0;
}
async function loadTargetsFromRepo(){
  const text = await fetchText(CONFIG.TARGETS_URL);
  targets = parseTargets(text);
  targetsLoaded = Object.keys(targets).length > 0;
}

/* ---------------- refresh all ---------------- */
async function refreshAll(){
  // As-of = yesterday
  const now = new Date();
  const asOf = new Date(now.getFullYear(), now.getMonth(), now.getDate()-1);
  const asOfYmd = ymd(asOf);

  el("asOfLabel").textContent = asOfYmd;

  try { await loadCalendarFromRepo(); } catch(e){ console.warn(e); calendarLoaded=false; }
  try { await loadTargetsFromRepo(); } catch(e){ console.warn(e); targetsLoaded=false; }

  el("targetsLabel").textContent = targetsLoaded ? "Yes" : "No";
  el("calendarLabel").textContent = calendarLoaded ? "Yes" : "No";

  // Always load full year so prior months are available
  try{
    await loadDailyFromGoogle(yearStart(asOfYmd), yearEnd(asOfYmd));
    el("sheetStatus").textContent = `Loaded (${sheetLoadedCount})`;
  } catch(err){
    console.error(err);
    el("sheetStatus").textContent = "Offline";
    toast("Google Sheet fetch failed. Check Apps Script /exec URL.");
  }

  renderLog(asOfYmd);
  renderSummaries(asOfYmd);
}

/* ---------------- init ---------------- */
(function init(){
  el("btnRefresh").addEventListener("click", refreshAll);

  el("kpiView").addEventListener("change", ()=>{
    const now = new Date();
    const asOf = new Date(now.getFullYear(), now.getMonth(), now.getDate()-1);
    renderSummaries(ymd(asOf));
  });

  el("logRange").addEventListener("change", ()=>{
    const now = new Date();
    const asOf = new Date(now.getFullYear(), now.getMonth(), now.getDate()-1);
    renderLog(ymd(asOf));
  });

  el("includeOffDays").addEventListener("change", ()=>{
    const now = new Date();
    const asOf = new Date(now.getFullYear(), now.getMonth(), now.getDate()-1);
    renderLog(ymd(asOf));
  });

  refreshAll();
})();
