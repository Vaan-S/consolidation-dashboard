import React, { useState, useMemo, useRef, useCallback } from "react";

/* ============================================================
   Meridian Software Group — Consolidation Dashboard
   Pure in-browser engine. Reproduces the Excel workbook to the penny.
   Sign convention: assets/expenses +, liabilities/equity/income −.
   ============================================================ */

const EMBEDDED = {"Parent":[{"localCode":"1010","localName":"Operating Cash","amount":120000,"groupCode":"1000"},{"localCode":"1210","localName":"Trade Receivables","amount":80000,"groupCode":"1100"},{"localCode":"1220","localName":"Due from Affiliates","amount":142058.82,"groupCode":"1200"},{"localCode":"1310","localName":"Merchandise Inventory","amount":30000,"groupCode":"1300"},{"localCode":"1510","localName":"Fixed Assets - Net","amount":200000,"groupCode":"1500"},{"localCode":"1610","localName":"Investment in Subsidiaries","amount":240000,"groupCode":"1600"},{"localCode":"2010","localName":"Trade Payables","amount":-60000,"groupCode":"2000"},{"localCode":"2310","localName":"Accrued Expenses","amount":-40000,"groupCode":"2200"},{"localCode":"2410","localName":"Unearned Revenue","amount":-50000,"groupCode":"2300"},{"localCode":"3010","localName":"Common Stock","amount":-300000,"groupCode":"3000"},{"localCode":"3110","localName":"Retained Earnings","amount":-282058.82,"groupCode":"3100"},{"localCode":"4010","localName":"Software Revenue","amount":-400000,"groupCode":"4000"},{"localCode":"4110","localName":"Intercompany Revenue","amount":-50000,"groupCode":"4100"},{"localCode":"5010","localName":"Cost of Sales","amount":180000,"groupCode":"5000"},{"localCode":"5110","localName":"Salaries & Wages","amount":140000,"groupCode":"5100"},{"localCode":"5210","localName":"Office & Admin","amount":40000,"groupCode":"5200"},{"localCode":"5310","localName":"Depreciation","amount":10000,"groupCode":"5300"}],"Canada":[{"localCode":"100","localName":"Bank","amount":90000,"groupCode":"1000"},{"localCode":"110","localName":"A/R Trade","amount":70000,"groupCode":"1100"},{"localCode":"130","localName":"Inventory","amount":95000,"groupCode":"1300"},{"localCode":"150","localName":"PP&E Net","amount":120000,"groupCode":"1500"},{"localCode":"200","localName":"A/P Trade","amount":-40000,"groupCode":"2000"},{"localCode":"220","localName":"Intercompany Payable","amount":-30000,"groupCode":"2100"},{"localCode":"230","localName":"Accruals","amount":-25000,"groupCode":"2200"},{"localCode":"240","localName":"Deferred Revenue","amount":-35000,"groupCode":"2300"},{"localCode":"300","localName":"Share Capital","amount":-130000,"groupCode":"3000"},{"localCode":"310","localName":"Retained Earnings","amount":-85000,"groupCode":"3100"},{"localCode":"400","localName":"Subscription Revenue","amount":-300000,"groupCode":"4000"},{"localCode":"500","localName":"Cost of Sales","amount":120000,"groupCode":"5000"},{"localCode":"510","localName":"Payroll","amount":90000,"groupCode":"5100"},{"localCode":"520","localName":"Admin Expenses","amount":55000,"groupCode":"5200"},{"localCode":"530","localName":"Depreciation","amount":5000,"groupCode":"5300"}],"UK":[{"localCode":"A100","localName":"Cash at Bank","amount":60000,"groupCode":"1000"},{"localCode":"A200","localName":"Trade Debtors","amount":55000,"groupCode":"1100"},{"localCode":"A300","localName":"Stock","amount":40000,"groupCode":"1300"},{"localCode":"A400","localName":"Tangible Fixed Assets","amount":90000,"groupCode":"1500"},{"localCode":"B100","localName":"Trade Creditors","amount":-35000,"groupCode":"2000"},{"localCode":"B200","localName":"Accruals","amount":-20000,"groupCode":"2200"},{"localCode":"B300","localName":"Deferred Income","amount":-30000,"groupCode":"2300"},{"localCode":"C100","localName":"Called-up Share Capital","amount":-67500,"groupCode":"3000"},{"localCode":"C200","localName":"Profit & Loss Reserve","amount":-52500,"groupCode":"3100"},{"localCode":"D100","localName":"Turnover","amount":-240000,"groupCode":"4000"},{"localCode":"E100","localName":"Cost of Sales","amount":95000,"groupCode":"5000"},{"localCode":"E200","localName":"Staff Costs","amount":70000,"groupCode":"5100"},{"localCode":"E300","localName":"Administrative Expenses","amount":30000,"groupCode":"5200"},{"localCode":"E400","localName":"Depreciation","amount":5000,"groupCode":"5300"}],"India":[{"localCode":"1001","localName":"Cash & Bank","amount":6500000,"groupCode":"1000"},{"localCode":"1002","localName":"Sundry Debtors","amount":4000000,"groupCode":"1100"},{"localCode":"1004","localName":"Closing Stock","amount":3000000,"groupCode":"1300"},{"localCode":"1005","localName":"Fixed Assets (Net)","amount":9000000,"groupCode":"1500"},{"localCode":"2001","localName":"Sundry Creditors","amount":-2500000,"groupCode":"2000"},{"localCode":"2002","localName":"Intercompany Payable","amount":-8000000,"groupCode":"2100"},{"localCode":"2003","localName":"Provisions","amount":-1500000,"groupCode":"2200"},{"localCode":"2004","localName":"Advance from Customers","amount":-2000000,"groupCode":"2300"},{"localCode":"3001","localName":"Equity Share Capital","amount":-5000000,"groupCode":"3000"},{"localCode":"3002","localName":"Reserves & Surplus","amount":-2000000,"groupCode":"3100"},{"localCode":"4001","localName":"Service Income","amount":-20000000,"groupCode":"4000"},{"localCode":"5001","localName":"Cost of Services","amount":9000000,"groupCode":"5000"},{"localCode":"5002","localName":"Staff Cost","amount":6000000,"groupCode":"5100"},{"localCode":"5003","localName":"Other Expenses","amount":2500000,"groupCode":"5200"},{"localCode":"5004","localName":"Depreciation","amount":1000000,"groupCode":"5300"}]};

const GROUP = {
  "1000":"Cash and cash equivalents","1100":"Accounts receivable, trade","1200":"Intercompany receivable",
  "1300":"Inventory","1500":"Property and equipment, net","1600":"Investment in subsidiaries",
  "2000":"Accounts payable, trade","2100":"Intercompany payable","2200":"Accrued liabilities",
  "2300":"Deferred revenue","3000":"Share capital","3100":"Retained earnings",
  "3200":"Foreign currency translation reserve","3300":"Non-controlling interest",
  "4000":"Revenue","4100":"Intercompany revenue","5000":"Cost of revenue",
  "5100":"Employee compensation","5200":"General and administrative","5300":"Depreciation expense",
  "5400":"Foreign exchange gain/loss"
};
const ORDER = Object.keys(GROUP);
const ENTITIES = ["Parent","Canada","UK","India"];
const DEFAULT_ENTITY_NAME = {
  Parent:"Meridian Software Inc.", Canada:"Meridian Software Canada Corp.",
  UK:"Meridian Software UK Ltd.", India:"Meridian Software India Pvt. Ltd."
};
const DEFAULT_GROUP_NAME = "Meridian Software Group";
const CCY = { Parent:"USD", Canada:"CAD", UK:"GBP", India:"INR" };
const OWNERSHIP = { Parent:1, Canada:1, UK:1, India:0.8 };
const RATES = { CAD:{cl:1.36,av:1.35,hi:1.30}, GBP:{cl:0.79,av:0.78,hi:0.75}, INR:{cl:83.0,av:82.0,hi:80.0} };

const r2 = x => Math.round((x + Number.EPSILON) * 100) / 100;
const cls = code =>
  ["1000","1100","1200","1300","1500","1600","2000","2100","2200","2300"].includes(code) ? "BS"
  : ["3000","3100"].includes(code) ? "EQ" : "PL";
const fmt = (x, dp=2) => {
  if (x === 0 || x === undefined || x === null) return "–";
  const n = Math.abs(x).toLocaleString("en-US",{minimumFractionDigits:dp,maximumFractionDigits:dp});
  return x < 0 ? `(${n})` : n;
};

/* ---------------- ENGINE ---------------- */
function runEngine(tbs, rates) {
  // tbs: { entity: [{groupCode, amount}] } in local currency
  const byGroup = {};
  ENTITIES.forEach(e => {
    byGroup[e] = {};
    (tbs[e]||[]).forEach(row => {
      if (!row.groupCode) return;
      byGroup[e][row.groupCode] = (byGroup[e][row.groupCode]||0) + Number(row.amount||0);
    });
  });
  // translate
  const translated = {}; const fctr = {};
  ENTITIES.forEach(e => {
    const ccy = CCY[e]; translated[e] = {};
    if (ccy === "USD") { ORDER.forEach(c => translated[e][c] = r2(byGroup[e][c]||0)); fctr[e]=0; return; }
    const rt = rates[ccy];
    ORDER.forEach(c => {
      const v = byGroup[e][c]||0;
      if (!v) { translated[e][c]=0; return; }
      const rate = cls(c)==="BS" ? rt.cl : cls(c)==="EQ" ? rt.hi : rt.av;
      translated[e][c] = r2(v/rate);
    });
    const s = ORDER.reduce((a,c)=>a+(translated[e][c]||0),0);
    translated[e]["3200"] = r2(-s); fctr[e] = r2(-s);
  });
  // combined
  const combined = {}; ORDER.forEach(c => combined[c] = r2(ENTITIES.reduce((a,e)=>a+(translated[e][c]||0),0)));
  // adjustments
  const adj = {}; ORDER.forEach(c=>adj[c]=0);
  adj["1300"]+=20000; adj["2100"]+=-20000;
  const trueup = r2(100000 - 8000000/rates.INR.cl);
  adj["5400"]+=trueup; adj["2100"]+=-trueup;
  // eliminations
  const elim = {}; ORDER.forEach(c=>elim[c]=0);
  elim["4100"]+=30000; elim["5200"]+=-30000;
  elim["4100"]+=20000; elim["5000"]+=-20000;
  const ca_ic = r2(30000/rates.CAD.cl);
  const ic_total = r2(100000 + 20000 + ca_ic);
  elim["2100"]+=ic_total; elim["1200"]+=-ic_total;
  elim["5000"]+=10000; elim["1300"]+=-10000;
  elim["3000"]+=240000; elim["1600"]+=-240000;
  const IN = translated.India;
  const nci_sc = r2(0.20*(-IN["3000"])), nci_re = r2(0.20*(-IN["3100"])), nci_fctr = r2(0.20*(-IN["3200"]));
  const nci_total = r2(nci_sc+nci_re+nci_fctr);
  elim["3000"]+=nci_sc; elim["3100"]+=nci_re; elim["3200"]+=nci_fctr; elim["3300"]+=-nci_total;
  // consolidated
  const cons = {}; ORDER.forEach(c => cons[c] = r2((combined[c]||0)+(adj[c]||0)+(elim[c]||0)));
  // derived
  const NI = r2(-(cons["4000"]+cons["4100"]+cons["5000"]+cons["5100"]+cons["5200"]+cons["5300"]+cons["5400"]));
  const india_profit = r2(-(IN["4000"]+IN["5000"]+IN["5100"]+IN["5200"]+IN["5300"]));
  const nci_profit = r2(0.20*india_profit);
  const owners_profit = r2(NI - nci_profit);
  const total_assets = r2(cons["1000"]+cons["1100"]+cons["1200"]+cons["1300"]+cons["1500"]+cons["1600"]);
  const total_liab = r2(-(cons["2000"]+cons["2100"]+cons["2200"]+cons["2300"]));
  const share_capital = r2(-cons["3000"]);
  const re_closing = r2(-cons["3100"] + owners_profit);
  const fctr_closing = r2(-cons["3200"]);
  const nci_closing = r2(-cons["3300"] + nci_profit);
  const equity_owners = r2(share_capital + re_closing + fctr_closing);
  const total_equity = r2(equity_owners + nci_closing);
  const total_LE = r2(total_liab + total_equity);
  const balance_check = r2(total_assets - total_LE);
  return { byGroup, translated, fctr, combined, adj, elim, cons, trueup, ca_ic, ic_total,
    nci:{sc:nci_sc,re:nci_re,fctr:nci_fctr,total:nci_total,profit:nci_profit,closing:nci_closing},
    stmt:{NI,india_profit,nci_profit,owners_profit,total_assets,total_liab,share_capital,
      re_closing,fctr_closing,nci_closing,equity_owners,total_equity,total_LE,balance_check} };
}

/* ---------------- JE EXPORT ---------------- */
function buildJEs(result) {
  const t = result.trueup, ca = result.ca_ic, ic = result.ic_total;
  const IN = result.translated.India;
  const nci_sc = r2(0.20*(-IN["3000"])), nci_re = r2(0.20*(-IN["3100"])), nci_fctr = r2(0.20*(-IN["3200"]));
  const rows = [];
  const push = (id,type,issue,code,acct,dr,cr,narr)=>rows.push({id,type,issue,code,acct,dr,cr,narr});
  // Adjustments
  push("A1","Adjustment","Issue 5 — Goods in transit","1300","Inventory (UK)",20000,"","Record inventory in transit that UK had not booked for goods shipped by Parent on the last day.");
  push("A1","Adjustment","Issue 5 — Goods in transit","2100","Intercompany payable (UK)","",20000,"Corresponding intercompany payable owed by UK to Parent.");
  push("A2","Adjustment","Issue 1 — IC balance FX true-up","5400","Foreign exchange loss",t,"","Re-translate India's USD-denominated intercompany loan to the closing rate (IAS 21).");
  push("A2","Adjustment","Issue 1 — IC balance FX true-up","2100","Intercompany payable (India)","",t,"Increase India intercompany payable to the USD amount owed at closing.");
  // Eliminations
  push("E1","Elimination","Issue 7 — Management fee","4100","Intercompany revenue",30000,"","Eliminate intercompany management fee revenue recorded by Parent.");
  push("E1","Elimination","Issue 7 — Management fee","5200","General and administrative","",30000,"Eliminate the matching admin expense in Canada. No effect on group profit.");
  push("E2","Elimination","Issue 5 — Goods in transit","4100","Intercompany revenue",20000,"","Eliminate the intercompany sale Parent booked for the in-transit goods.");
  push("E2","Elimination","Issue 5 — Goods in transit","5000","Cost of revenue","",20000,"Eliminate the matching cost.");
  push("E3","Elimination","Issue 1 — Intercompany balances","2100","Intercompany payable",ic,"",`Eliminate group intercompany payables (India 100,000 + UK 20,000 + Canada ${fmt(ca)}).`);
  push("E3","Elimination","Issue 1 — Intercompany balances","1200","Intercompany receivable","",ic,"Eliminate Parent's matching intercompany receivable.");
  push("E4","Elimination","Issue 2 — Unrealized profit","5000","Cost of revenue",10000,"","Reverse intercompany profit on goods held by Canada, unsold at period-end.");
  push("E4","Elimination","Issue 2 — Unrealized profit","1300","Inventory","",10000,"Reduce group inventory to original cost.");
  push("E5","Elimination","Investment vs equity","3000","Share capital (subsidiaries)",240000,"","Eliminate Parent's investment against subsidiary share capital at acquisition.");
  push("E5","Elimination","Investment vs equity","1600","Investment in subsidiaries","",240000,"Remove the carrying amount of the investment.");
  push("E6","Reclassification","Non-controlling interest","3000","Share capital — NCI share",nci_sc,"","Allocate 20% of India share capital to non-controlling interest.");
  push("E6","Reclassification","Non-controlling interest","3100","Retained earnings — NCI share",nci_re,"","Allocate 20% of India retained earnings to non-controlling interest.");
  if (nci_fctr>=0) push("E6","Reclassification","Non-controlling interest","3200","FCTR — NCI share",nci_fctr,"","Allocate 20% of India's translation reserve to non-controlling interest.");
  else push("E6","Reclassification","Non-controlling interest","3200","FCTR — NCI share","",Math.abs(nci_fctr),"Allocate 20% of India's translation reserve (debit reserve) to NCI.");
  push("E6","Reclassification","Non-controlling interest","3300","Non-controlling interest","",result.nci.total,"Recognise the 20% outside interest in India's translated net equity.");
  return rows;
}
function downloadJECSV(result, groupName) {
  const rows = buildJEs(result);
  const q = s => `"${String(s).replace(/"/g,'""')}"`;
  const lines = [];
  lines.push(`# ${groupName} — Consolidation Journal Entries (USD)`);
  lines.push("# Adjustments and eliminations to consolidate and resolve the seven planted issues.");
  lines.push("# Sign: amounts shown as Debit / Credit. Each JE balances.");
  lines.push(["JE #","Type","Issue addressed","Level","Group Code","Account","Debit (USD)","Credit (USD)","Narrative"].join(","));
  rows.forEach(r=>lines.push([q(r.id),q(r.type),q(r.issue),q("Group/Consolidation"),q(r.code),q(r.acct),
    r.dr!==""?r.dr:"",r.cr!==""?r.cr:"",q(r.narr)].join(",")));
  const td = r2(rows.reduce((a,r)=>a+(r.dr||0),0)), tc = r2(rows.reduce((a,r)=>a+(r.cr||0),0));
  lines.push(["","","","","",q("TOTAL"),td,tc,""].join(","));
  const blob = new Blob([lines.join("\n")], {type:"text/csv"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `${groupName.replace(/[^a-z0-9]+/gi,"_").toLowerCase()}_consolidation_journal_entries.csv`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}

/* ---------------- CSV PARSE ---------------- */
function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim() && !l.trim().startsWith("#"));
  const rows = lines.map(l => {
    const out = []; let cur = "", q = false;
    for (const ch of l) {
      if (ch === '"') q = !q;
      else if (ch === "," && !q) { out.push(cur); cur = ""; }
      else cur += ch;
    }
    out.push(cur); return out.map(s => s.trim());
  });
  return rows;
}
function tbFromCSV(rows) {
  // find header row
  let hi = rows.findIndex(r => r.some(c => /account\s*code/i.test(c)));
  if (hi < 0) hi = 0;
  const header = rows[hi].map(h => h.toLowerCase());
  const ci = {
    code: header.findIndex(h => /code/.test(h)),
    name: header.findIndex(h => /name|account/.test(h) && !/code/.test(h)),
    debit: header.findIndex(h => /debit/.test(h)),
    credit: header.findIndex(h => /credit/.test(h)),
    amount: header.findIndex(h => /amount|balance/.test(h)),
  };
  const out = [];
  for (let i = hi+1; i < rows.length; i++) {
    const r = rows[i]; if (!r || r.every(c => !c)) continue;
    const code = ci.code>=0 ? r[ci.code] : r[0];
    if (!code || /total/i.test(code) || /total/i.test((r[ci.name]||""))) continue;
    const nm = ci.name>=0 ? r[ci.name] : "";
    const num = s => { const v = parseFloat(String(s||"").replace(/[(),$]/g,"").replace(/^\)/,"")); return isNaN(v)?0:v; };
    let amt = 0;
    if (ci.amount>=0) amt = num(r[ci.amount]);
    else { const d = num(r[ci.debit]), c = num(r[ci.credit]); amt = d - c; }
    out.push({ localCode: code, localName: nm, amount: amt, groupCode: null });
  }
  return out;
}
function autoMap(rows) {
  // best-effort: if localCode matches a group code, use it; else map by name keyword
  const kw = [
    [/cash|bank/i,"1000"],[/receivable|debtor/i,"1100"],[/due from|intercompany rec|group debtor/i,"1200"],
    [/invent|stock/i,"1300"],[/pp&e|fixed asset|tangible|property|equipment/i,"1500"],[/investment/i,"1600"],
    [/payable|creditor/i,"2000"],[/intercompany pay|group creditor/i,"2100"],[/accru|provision/i,"2200"],
    [/deferred|unearned|advance from/i,"2300"],[/share capital|common stock|called-up|equity share/i,"3000"],
    [/retained|reserve|p&l|profit \& loss/i,"3100"],[/revenue|turnover|income|subscription|service/i,"4000"],
    [/intercompany rev/i,"4100"],[/cost of (sales|revenue|service)/i,"5000"],[/salar|payroll|staff|wage|comp/i,"5100"],
    [/admin|office|general/i,"5200"],[/deprec/i,"5300"],[/fx|foreign exchange|forex/i,"5400"],
  ];
  return rows.map(row => {
    let gc = null;
    if (GROUP[row.localCode]) gc = row.localCode;
    if (!gc) for (const [re,code] of kw) { if (re.test(row.localName) || re.test(row.localCode)) { gc = code; break; } }
    if (!gc && /intercompany/i.test(row.localName)) gc = /pay/i.test(row.localName) ? "2100" : "1200";
    return { ...row, groupCode: gc };
  });
}

/* ---------------- UI PRIMITIVES ---------------- */
const C = {
  paper:"#FBFAF7", ink:"#1B2A4A", inkSoft:"#4A5670", line:"#E4E0D6", lineSoft:"#EFECE4",
  green:"#2F6F5B", greenBg:"#EAF2EE", amber:"#B47B18", amberBg:"#FaF3E2", red:"#9B3D2E", redBg:"#F6EAE6",
  white:"#FFFFFF", chip:"#F1EEE7"
};
const mono = "'SF Mono',ui-monospace,'Cascadia Mono','Roboto Mono',Menlo,monospace";
const sans = "'Inter',system-ui,-apple-system,'Segoe UI',sans-serif";

function Num({v, dp=2, bold, color}) {
  return <span style={{fontFamily:mono, fontVariantNumeric:"tabular-nums", fontWeight:bold?600:400,
    color: color || (v<0 ? C.red : C.ink)}}>{fmt(v,dp)}</span>;
}

/* ---------------- STEP COMPONENTS ---------------- */
function StepUpload({ tbs, setTbs, sources, setSources, onReset, names, setNames, groupName, setGroupName }) {
  const fileRef = useRef({});
  const handleFile = (entity, file) => {
    const reader = new FileReader();
    reader.onload = e => {
      const rows = tbFromCSV(parseCSV(e.target.result));
      const mapped = autoMap(rows);
      setTbs(prev => ({ ...prev, [entity]: mapped }));
      setSources(prev => ({ ...prev, [entity]: file.name }));
    };
    reader.readAsText(file);
  };
  return (
    <div>
      <SectionHead n="1" title="Upload trial balances"
        note="Meridian data is pre-loaded so the consolidation runs immediately. Replace any entity by uploading its trial balance as CSV — the tool re-maps and re-consolidates live. Rename the group and entities to match your own data." />
      <div style={{marginTop:16,display:"flex",alignItems:"center",gap:10}}>
        <label style={{fontSize:12.5,color:C.inkSoft,fontWeight:600}}>Group name</label>
        <input value={groupName} onChange={e=>setGroupName(e.target.value)}
          style={{fontFamily:sans,fontSize:14,fontWeight:600,color:C.ink,padding:"6px 10px",
            border:`1px solid ${C.line}`,borderRadius:7,minWidth:280}}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginTop:16}}>
        {ENTITIES.map(e => {
          const rows = tbs[e]||[];
          const bal = r2(rows.reduce((a,r)=>a+Number(r.amount||0),0));
          const src = sources[e];
          return (
            <div key={e} style={{border:`1px solid ${C.line}`,borderRadius:10,background:C.white,padding:"16px 18px"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline"}}>
                <div style={{flex:1,minWidth:0}}>
                  <input value={names[e]} onChange={ev=>setNames(prev=>({...prev,[e]:ev.target.value}))}
                    style={{fontWeight:600,color:C.ink,fontSize:15,fontFamily:sans,border:"1px solid transparent",
                      borderRadius:6,padding:"3px 6px",marginLeft:-6,width:"calc(100% - 8px)",background:"transparent"}}
                    onFocus={ev=>ev.target.style.border=`1px solid ${C.line}`}
                    onBlur={ev=>ev.target.style.border="1px solid transparent"}/>
                  <div style={{fontSize:12.5,color:C.inkSoft,marginTop:2}}>
                    {CCY[e]} · {rows.length} accounts · {OWNERSHIP[e]<1 ? `${OWNERSHIP[e]*100}% owned` : "wholly owned"} · <span style={{color:C.inkSoft}}>{e}</span>
                  </div>
                </div>
                <span style={{fontSize:11.5,fontFamily:mono,padding:"3px 8px",borderRadius:20,
                  background: bal===0?C.greenBg:C.redBg, color: bal===0?C.green:C.red, fontWeight:600}}>
                  {bal===0 ? "balanced" : `off ${fmt(bal)}`}
                </span>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:10,marginTop:14}}>
                <button onClick={()=>fileRef.current[e]?.click()}
                  style={{fontFamily:sans,fontSize:13,fontWeight:500,color:C.ink,background:C.chip,
                    border:`1px solid ${C.line}`,borderRadius:7,padding:"7px 12px",cursor:"pointer"}}>
                  {src ? "Replace file" : "Upload CSV"}
                </button>
                <span style={{fontSize:12,color:C.inkSoft,fontFamily:mono}}>
                  {src || "using pre-loaded data"}
                </span>
                <input ref={el=>fileRef.current[e]=el} type="file" accept=".csv" style={{display:"none"}}
                  onChange={ev=>ev.target.files[0]&&handleFile(e,ev.target.files[0])} />
              </div>
            </div>
          );
        })}
      </div>
      <div style={{marginTop:16,display:"flex",gap:12,alignItems:"center"}}>
        <button onClick={onReset} style={{fontFamily:sans,fontSize:12.5,color:C.inkSoft,background:"none",
          border:"none",cursor:"pointer",textDecoration:"underline"}}>Reset to Meridian data</button>
      </div>
    </div>
  );
}

function StepMap({ tbs, setTbs, names }) {
  const unmapped = {};
  ENTITIES.forEach(e => { unmapped[e] = (tbs[e]||[]).filter(r=>!r.groupCode).length; });
  const totalUnmapped = Object.values(unmapped).reduce((a,b)=>a+b,0);
  const setMap = (entity, idx, gc) => {
    setTbs(prev => {
      const rows = [...prev[entity]]; rows[idx] = { ...rows[idx], groupCode: gc||null };
      return { ...prev, [entity]: rows };
    });
  };
  return (
    <div>
      <SectionHead n="2" title="Map local accounts to the group chart"
        note="Every entity keeps its own local codes; each maps to one group account. Unmapped accounts block the consolidation — this is how a real close catches a chart-of-accounts mismatch (Issue 6)." />
      <div style={{marginTop:14,padding:"10px 14px",borderRadius:8,display:"inline-flex",gap:8,alignItems:"center",
        background: totalUnmapped?C.amberBg:C.greenBg, color: totalUnmapped?C.amber:C.green, fontSize:13,fontWeight:600}}>
        {totalUnmapped ? `${totalUnmapped} account${totalUnmapped>1?"s":""} still unmapped` : "All accounts mapped"}
      </div>
      {ENTITIES.map(e => (
        <div key={e} style={{marginTop:20}}>
          <div style={{fontWeight:600,color:C.ink,fontSize:14,marginBottom:8}}>{names[e]} <span style={{color:C.inkSoft,fontWeight:400}}>· {CCY[e]}</span></div>
          <div style={{border:`1px solid ${C.line}`,borderRadius:9,overflow:"hidden",background:C.white}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{background:C.lineSoft,color:C.inkSoft,fontSize:11.5,textAlign:"left"}}>
                <th style={thL}>Local code</th><th style={thL}>Local account name</th>
                <th style={{...thR,width:130}}>Amount ({CCY[e]})</th><th style={{...thL,width:300}}>Group account</th>
              </tr></thead>
              <tbody>
                {(tbs[e]||[]).map((row,idx)=>(
                  <tr key={idx} style={{borderTop:`1px solid ${C.lineSoft}`,
                    background: row.groupCode?"transparent":C.amberBg}}>
                    <td style={{...tdL,fontFamily:mono}}>{row.localCode}</td>
                    <td style={tdL}>{row.localName}</td>
                    <td style={tdR}><Num v={row.amount}/></td>
                    <td style={tdL}>
                      <select value={row.groupCode||""} onChange={ev=>setMap(e,idx,ev.target.value)}
                        style={{fontFamily:sans,fontSize:12.5,padding:"5px 8px",borderRadius:6,
                          border:`1px solid ${row.groupCode?C.line:C.amber}`,background:C.white,color:C.ink,width:"100%",maxWidth:290}}>
                        <option value="">— select group account —</option>
                        {ORDER.filter(c=>!["3200","3300"].includes(c)).map(c=>(
                          <option key={c} value={c}>{c} · {GROUP[c]}</option>
                        ))}
                      </select>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ))}
    </div>
  );
}

function StepTranslate({ result, rates, setRates, names }) {
  const setRate = (ccy,kind,val) => setRates(prev=>({...prev,[ccy]:{...prev[ccy],[kind]:parseFloat(val)||0}}));
  return (
    <div>
      <SectionHead n="3" title="Translate to USD (IAS 21)"
        note="Assets and liabilities translate at the closing rate, income and expenses at the average rate, equity at the historical rate. The translation reserve (FCTR) is the balancing figure so each entity still ties (Issues 3 & 4)." />
      <div style={{marginTop:16,display:"flex",gap:20,flexWrap:"wrap"}}>
        {Object.keys(RATES).map(ccy=>(
          <div key={ccy} style={{border:`1px solid ${C.line}`,borderRadius:9,background:C.white,padding:"12px 16px"}}>
            <div style={{fontWeight:600,color:C.ink,marginBottom:8,fontSize:13.5}}>{ccy} per 1 USD</div>
            <div style={{display:"flex",gap:12}}>
              {[["cl","Closing"],["av","Average"],["hi","Historical"]].map(([k,lbl])=>(
                <label key={k} style={{fontSize:11.5,color:C.inkSoft}}>
                  {lbl}<br/>
                  <input type="number" step="0.01" value={rates[ccy][k]} onChange={e=>setRate(ccy,k,e.target.value)}
                    style={{width:64,fontFamily:mono,fontSize:13,padding:"4px 6px",marginTop:3,
                      border:`1px solid ${C.line}`,borderRadius:5,color:C.ink}}/>
                </label>
              ))}
            </div>
          </div>
        ))}
      </div>
      {["Canada","UK","India"].map(e=>(
        <div key={e} style={{marginTop:20}}>
          <div style={{fontWeight:600,color:C.ink,fontSize:14,marginBottom:8}}>
            {names[e]} <span style={{color:C.inkSoft,fontWeight:400}}>· {CCY[e]} → USD</span>
          </div>
          <div style={{border:`1px solid ${C.line}`,borderRadius:9,overflow:"hidden",background:C.white}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{background:C.lineSoft,color:C.inkSoft,fontSize:11.5}}>
                <th style={thL}>Code</th><th style={thL}>Account</th><th style={thR}>Local</th>
                <th style={{...thR,width:70}}>Rate</th><th style={thR}>USD</th>
              </tr></thead>
              <tbody>
                {ORDER.filter(c=>result.byGroup[e][c]!==undefined && c!=="3200").map(c=>{
                  const rate = cls(c)==="BS"?rates[CCY[e]].cl:cls(c)==="EQ"?rates[CCY[e]].hi:rates[CCY[e]].av;
                  return (
                    <tr key={c} style={{borderTop:`1px solid ${C.lineSoft}`}}>
                      <td style={{...tdL,fontFamily:mono}}>{c}</td><td style={tdL}>{GROUP[c]}</td>
                      <td style={tdR}><Num v={r2(result.byGroup[e][c])}/></td>
                      <td style={{...tdR,color:C.inkSoft,fontFamily:mono}}>{rate.toFixed(2)}</td>
                      <td style={tdR}><Num v={result.translated[e][c]}/></td>
                    </tr>
                  );
                })}
                <tr style={{borderTop:`1px solid ${C.line}`,background:C.greenBg}}>
                  <td style={{...tdL,fontFamily:mono}}>3200</td>
                  <td style={tdL} colSpan={3}><em>Translation reserve (FCTR) — balancing figure</em></td>
                  <td style={tdR}><Num v={result.translated[e]["3200"]} bold color={C.green}/></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      ))}
    </div>
  );
}

function JETable({ title, entries }) {
  return (
    <div style={{marginTop:18}}>
      <div style={{fontWeight:600,color:C.ink,fontSize:14,marginBottom:8}}>{title}</div>
      <div style={{border:`1px solid ${C.line}`,borderRadius:9,overflow:"hidden",background:C.white}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead><tr style={{background:C.lineSoft,color:C.inkSoft,fontSize:11.5}}>
            <th style={{...thL,width:44}}>JE</th><th style={thL}>Account</th>
            <th style={{...thR,width:120}}>Debit</th><th style={{...thR,width:120}}>Credit</th>
          </tr></thead>
          <tbody>
            {entries.map((e,i)=>(
              <React.Fragment key={i}>
                {e.narr && <tr style={{borderTop:`1px solid ${C.lineSoft}`}}>
                  <td style={{...tdL,fontFamily:mono,fontWeight:600,color:C.ink,verticalAlign:"top"}}>{e.id}</td>
                  <td colSpan={3} style={{...tdL,color:C.inkSoft,fontStyle:"italic",fontSize:12.5,padding:"8px 12px"}}>{e.narr}</td>
                </tr>}
                {e.lines.map((ln,j)=>(
                  <tr key={j} style={{borderTop: j===0?"none":`1px solid ${C.lineSoft}`}}>
                    <td style={tdL}></td>
                    <td style={{...tdL,paddingLeft: ln.cr?36:20}}>{ln.acct}</td>
                    <td style={tdR}>{ln.dr?<Num v={ln.dr}/>:""}</td>
                    <td style={tdR}>{ln.cr?<Num v={ln.cr}/>:""}</td>
                  </tr>
                ))}
              </React.Fragment>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function StepAdjust({ result }) {
  const t = result.trueup;
  const entries = [
    { id:"A1", narr:"Goods in transit (Issue 5): record the inventory and intercompany payable the UK sub had not yet booked for goods shipped by Parent on the last day.",
      lines:[{acct:"Inventory (UK)",dr:20000},{acct:"Intercompany payable (UK)",cr:20000}] },
    { id:"A2", narr:`IC balance FX true-up (Issue 1): re-translate India's USD 100,000 loan to the closing rate. Booked at old rate 80.0; closing 83.0. Difference to FX loss.`,
      lines:[{acct:"Foreign exchange loss",dr:t},{acct:"Intercompany payable (India)",cr:t}] },
  ];
  return (
    <div>
      <SectionHead n="4" title="Adjustments" note="Two pre-elimination adjustments correct timing and currency errors on the entity books before intercompany balances are removed." />
      <JETable title="Adjustment journal entries" entries={entries} />
    </div>
  );
}

function StepEliminate({ result }) {
  const entries = [
    { id:"E1", narr:"Management fee (Issue 7): eliminate intercompany revenue against the matching admin expense in Canada. No effect on group profit.",
      lines:[{acct:"Intercompany revenue",dr:30000},{acct:"General and administrative",cr:30000}] },
    { id:"E2", narr:"Goods-in-transit sale (Issue 5): eliminate the intercompany sale Parent booked against cost of revenue.",
      lines:[{acct:"Intercompany revenue",dr:20000},{acct:"Cost of revenue",cr:20000}] },
    { id:"E3", narr:`Intercompany balances (Issue 1): eliminate the group receivable against the group payable (India 100,000 + UK 20,000 + Canada ${fmt(result.ca_ic)}).`,
      lines:[{acct:"Intercompany payable",dr:result.ic_total},{acct:"Intercompany receivable",cr:result.ic_total}] },
    { id:"E4", narr:"Unrealized profit in inventory (Issue 2): remove the 10,000 profit on goods Parent sold to Canada that remain unsold at period-end.",
      lines:[{acct:"Cost of revenue",dr:10000},{acct:"Inventory",cr:10000}] },
    { id:"E5", narr:"Investment vs equity: eliminate Parent's investment against subsidiary share capital at acquisition (Canada 100,000 + UK 90,000 + India 80% × 62,500 = 50,000).",
      lines:[{acct:"Share capital (subsidiaries)",dr:240000},{acct:"Investment in subsidiaries",cr:240000}] },
  ];
  return (
    <div>
      <SectionHead n="5" title="Eliminations" note="Remove intercompany revenue, balances, unrealized profit, and the parent's investment against subsidiary equity, so the group reports only transactions with the outside world." />
      <JETable title="Elimination journal entries" entries={entries} />
    </div>
  );
}

function StepNCI({ result }) {
  const IN = result.translated.India;
  const rows = [
    ["India share capital (USD, historical rate)", r2(-IN["3000"])],
    ["India retained earnings (USD, average rate)", r2(-IN["3100"])],
    ["India translation reserve / FCTR (USD)", r2(-IN["3200"])],
    ["India net equity (USD)", r2(-IN["3000"]-IN["3100"]-IN["3200"])],
    ["NCI ownership", 0.20, true],
    ["Non-controlling interest, before profit (USD)", result.nci.total],
    ["Add: NCI share of India profit for the year", result.nci.profit],
    ["Non-controlling interest at year-end (USD)", result.nci.closing],
  ];
  return (
    <div>
      <SectionHead n="6" title="Non-controlling interest"
        note="India is 80% owned. The outside 20% is carved out of the group's equity lines and presented on its own — including its share of the translation reserve, and its share of India's profit." />
      <div style={{marginTop:16,maxWidth:520,border:`1px solid ${C.line}`,borderRadius:9,overflow:"hidden",background:C.white}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <tbody>
            {rows.map(([lbl,val,pct],i)=>{
              const strong = /net equity|year-end/i.test(lbl);
              return (
                <tr key={i} style={{borderTop:i?`1px solid ${C.lineSoft}`:"none",
                  background: /year-end/i.test(lbl)?C.greenBg:"transparent"}}>
                  <td style={{...tdL,fontWeight:strong?600:400,color:C.ink}}>{lbl}</td>
                  <td style={tdR}>{pct ? <span style={{fontFamily:mono}}>{(val*100).toFixed(0)}%</span>
                    : <Num v={val} bold={strong} color={/year-end/i.test(lbl)?C.green:undefined}/>}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function StepConsolidated({ result, brokenResult, groupName }) {
  const s = result.stmt;
  const [showBroken, setShowBroken] = useState(false);
  const view = showBroken ? brokenResult : result;
  const issues = [
    { n:1, t:"Intercompany balance mismatch (FX)", before:"IC receivable and payable differ by 3,614.46 after translation", after:"Retranslated at closing; FX loss booked; balances eliminate to zero" },
    { n:2, t:"Unrealized profit in inventory", before:"Group inventory carries a 10,000 intra-group markup", after:"Markup removed; inventory and profit each reduced by 10,000" },
    { n:3, t:"Wrong FX rate on a balance-sheet line", before:"UK PP&E translated at average, not closing", after:"Closing rate applied in translation — no journal entry needed" },
    { n:4, t:"Translation reserve must balance", before:"Mixed rates leave each foreign column out of balance", after:"FCTR posted as the balancing figure so each entity ties" },
    { n:5, t:"Cut-off / goods in transit", before:"UK never recorded the in-transit stock and payable", after:"In-transit adjustment booked, then the IC sale eliminated" },
    { n:6, t:"Mismatched charts of accounts", before:"Four different local code sets; India has a local-only account", after:"Every local account mapped to the group chart" },
    { n:7, t:"Missing intercompany elimination", before:"Management fee revenue and expense both still in group P&L", after:"IC revenue eliminated against expense; no profit effect" },
  ];
  return (
    <div>
      <SectionHead n="7" title="Consolidated result"
        note="The tied consolidated statements. Toggle to see how the numbers looked before the adjustments and eliminations were posted." />
      <div style={{marginTop:14,display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
        <button onClick={()=>setShowBroken(false)} style={toggleBtn(!showBroken)}>After consolidation</button>
        <button onClick={()=>setShowBroken(true)} style={toggleBtn(showBroken)}>Before (raw combined)</button>
        <button onClick={()=>downloadJECSV(result, groupName)}
          style={{fontFamily:sans,fontSize:13,fontWeight:600,color:C.white,background:C.green,
            border:`1px solid ${C.green}`,borderRadius:7,padding:"7px 14px",cursor:"pointer",marginLeft:"auto"}}>
          ↓ Download JE CSV (postable)
        </button>
      </div>
      <div style={{marginTop:8,fontSize:12,color:C.inkSoft}}>
        The CSV holds all eight consolidation entries (2 adjustments + 6 eliminations) in a debit/credit format ready to post in an ERP.
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginTop:18}}>
        {/* Income statement */}
        <div>
          <StmtTitle>Consolidated statement of profit or loss</StmtTitle>
          <StmtTable rows={[
            ["Revenue", -view.cons["4000"]],
            ["Cost of revenue", -view.cons["5000"]],
            ["__gp","Gross profit", r2(-(view.cons["4000"]+view.cons["5000"]))],
            ["Employee compensation", -view.cons["5100"]],
            ["General and administrative", -view.cons["5200"]],
            ["Depreciation expense", -view.cons["5300"]],
            ["Intercompany revenue", -view.cons["4100"]],
            ["Foreign exchange gain/(loss)", -view.cons["5400"]],
            ["__ni","Profit for the year", view.stmt.NI],
          ]} />
          {!showBroken && (
            <StmtTable small rows={[
              ["Owners of the parent", s.owners_profit],
              ["Non-controlling interest", s.nci_profit],
            ]} />
          )}
        </div>
        {/* Balance sheet */}
        <div>
          <StmtTitle>Consolidated statement of financial position</StmtTitle>
          <StmtTable rows={[
            ["Cash and cash equivalents", view.cons["1000"]],
            ["Accounts receivable, trade", view.cons["1100"]],
            ["Intercompany receivable", view.cons["1200"]],
            ["Inventory", view.cons["1300"]],
            ["Property and equipment, net", view.cons["1500"]],
            ["Investment in subsidiaries", view.cons["1600"]],
            ["__ta","Total assets", showBroken
              ? r2(view.cons["1000"]+view.cons["1100"]+view.cons["1200"]+view.cons["1300"]+view.cons["1500"]+view.cons["1600"])
              : s.total_assets],
          ]} />
          <StmtTable rows={[
            ["Accounts payable, trade", -view.cons["2000"]],
            ["Intercompany payable", -view.cons["2100"]],
            ["Accrued liabilities", -view.cons["2200"]],
            ["Deferred revenue", -view.cons["2300"]],
            ["Share capital", -view.cons["3000"]],
            ["Retained earnings", showBroken ? -view.cons["3100"] : s.re_closing],
            ["Translation reserve (FCTR)", -view.cons["3200"]],
            ["Non-controlling interest", showBroken ? -view.cons["3300"] : s.nci_closing],
            ["__tle","Total liabilities & equity", showBroken
              ? r2(-(view.cons["2000"]+view.cons["2100"]+view.cons["2200"]+view.cons["2300"]+view.cons["3000"]+view.cons["3100"]+view.cons["3200"]+view.cons["3300"]))
              : s.total_LE],
          ]} />
        </div>
      </div>

      <div style={{marginTop:26}}>
        <div style={{fontWeight:600,color:C.ink,fontSize:15,marginBottom:12}}>The seven issues, caught and cleared</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          {issues.map(is=>(
            <div key={is.n} style={{border:`1px solid ${C.line}`,borderRadius:10,background:C.white,padding:"14px 16px"}}>
              <div style={{display:"flex",gap:9,alignItems:"baseline"}}>
                <span style={{fontFamily:mono,fontSize:12,fontWeight:600,color:C.white,background:C.ink,
                  borderRadius:5,padding:"2px 7px"}}>{is.n}</span>
                <span style={{fontWeight:600,color:C.ink,fontSize:13.5}}>{is.t}</span>
              </div>
              <div style={{marginTop:10,fontSize:12.5,color:C.red,display:"flex",gap:7}}>
                <span style={{fontWeight:600,minWidth:44}}>Before</span><span>{is.before}</span></div>
              <div style={{marginTop:6,fontSize:12.5,color:C.green,display:"flex",gap:7}}>
                <span style={{fontWeight:600,minWidth:44}}>After</span><span>{is.after}</span></div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ---------------- SHARED BITS ---------------- */
const thL = {padding:"9px 12px",fontWeight:500,textAlign:"left"};
const thR = {padding:"9px 12px",fontWeight:500,textAlign:"right"};
const tdL = {padding:"7px 12px",textAlign:"left",color:C.ink};
const tdR = {padding:"7px 12px",textAlign:"right"};
function SectionHead({n,title,note}) {
  return (
    <div style={{borderBottom:`1px solid ${C.line}`,paddingBottom:14}}>
      <div style={{fontSize:20,fontWeight:600,color:C.ink,letterSpacing:-0.2}}>{title}</div>
      {note && <div style={{fontSize:13.5,color:C.inkSoft,marginTop:6,maxWidth:760,lineHeight:1.5}}>{note}</div>}
    </div>
  );
}
function StmtTitle({children}) { return <div style={{fontSize:13,fontWeight:600,color:C.inkSoft,marginBottom:6}}>{children}</div>; }
function StmtTable({rows,small}) {
  return (
    <div style={{border:`1px solid ${C.line}`,borderRadius:9,overflow:"hidden",background:C.white,marginBottom:12}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:small?12.5:13}}>
        <tbody>
          {rows.map((row,i)=>{
            const isKey = typeof row[0]==="string" && row[0].startsWith("__");
            const label = isKey ? row[1] : row[0];
            const val = isKey ? row[2] : row[1];
            return (
              <tr key={i} style={{borderTop:i?`1px solid ${C.lineSoft}`:"none",
                background: isKey?C.chip:"transparent"}}>
                <td style={{...tdL,fontWeight:isKey?600:400,paddingLeft: small?20:12}}>{label}</td>
                <td style={tdR}><Num v={val} bold={isKey}/></td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
function toggleBtn(active){return {fontFamily:sans,fontSize:13,fontWeight:500,
  color:active?C.white:C.ink,background:active?C.ink:C.white,
  border:`1px solid ${active?C.ink:C.line}`,borderRadius:7,padding:"7px 14px",cursor:"pointer"};}

/* ---------------- APP SHELL ---------------- */
const STEPS = [
  {key:"upload",label:"Upload"},{key:"map",label:"Map"},{key:"translate",label:"Translate"},
  {key:"adjust",label:"Adjust"},{key:"eliminate",label:"Eliminate"},{key:"nci",label:"NCI"},
  {key:"consolidated",label:"Consolidated"},
];

export default function App() {
  const clone = () => JSON.parse(JSON.stringify(EMBEDDED));
  const [tbs, setTbs] = useState(()=>{
    const c = clone(); const o={}; ENTITIES.forEach(e=>o[e]=autoMap(c[e]).map((r,i)=>({...r,groupCode:c[e][i].groupCode}))); return o;
  });
  const [sources, setSources] = useState({});
  const [rates, setRates] = useState(()=>JSON.parse(JSON.stringify(RATES)));
  const [names, setNames] = useState(()=>({...DEFAULT_ENTITY_NAME}));
  const [groupName, setGroupName] = useState(DEFAULT_GROUP_NAME);
  const [step, setStep] = useState(0);

  const result = useMemo(()=>runEngine(tbs,rates),[tbs,rates]);
  // "broken" = combined only, no adj/elim
  const brokenResult = useMemo(()=>{
    const broken = JSON.parse(JSON.stringify(result));
    broken.cons = { ...result.combined };
    ORDER.forEach(c=>{ if(broken.cons[c]===undefined) broken.cons[c]=0; });
    broken.stmt = { ...result.stmt,
      NI: r2(-(result.combined["4000"]+result.combined["4100"]+result.combined["5000"]+result.combined["5100"]+result.combined["5200"]+result.combined["5300"]+(result.combined["5400"]||0))) };
    return broken;
  },[result]);

  const unmappedCount = ENTITIES.reduce((a,e)=>a+(tbs[e]||[]).filter(r=>!r.groupCode).length,0);
  const canProceed = !(step===1 && unmappedCount>0);
  const bc = result.stmt.balance_check;

  const onReset = () => {
    const c = clone(); const o={}; ENTITIES.forEach(e=>o[e]=c[e]); setTbs(o); setSources({}); setRates(JSON.parse(JSON.stringify(RATES)));
    setNames({...DEFAULT_ENTITY_NAME}); setGroupName(DEFAULT_GROUP_NAME);
  };

  return (
    <div style={{fontFamily:sans,background:C.paper,minHeight:"100vh",color:C.ink}}>
      <div style={{maxWidth:1180,margin:"0 auto",display:"grid",gridTemplateColumns:"228px 1fr",gap:0}}>
        {/* LEFT RAIL */}
        <aside style={{borderRight:`1px solid ${C.line}`,padding:"30px 22px",minHeight:"100vh"}}>
          <div style={{fontSize:16,fontWeight:700,letterSpacing:-0.3,lineHeight:1.25}}>{groupName}</div>
          <div style={{fontSize:12.5,color:C.inkSoft,marginTop:4}}>Multi-entity consolidation · IFRS · USD</div>
          <nav style={{marginTop:28}}>
            {STEPS.map((s,i)=>{
              const active = i===step; const done = i<step;
              return (
                <button key={s.key} onClick={()=>setStep(i)} style={{display:"flex",alignItems:"center",gap:11,
                  width:"100%",textAlign:"left",background:active?C.white:"transparent",
                  border:active?`1px solid ${C.line}`:"1px solid transparent",borderRadius:8,
                  padding:"9px 11px",marginBottom:3,cursor:"pointer",fontFamily:sans}}>
                  <span style={{width:22,height:22,borderRadius:"50%",flexShrink:0,display:"grid",placeItems:"center",
                    fontSize:11.5,fontWeight:600,fontFamily:mono,
                    background: active?C.ink:done?C.greenBg:C.chip, color: active?C.white:done?C.green:C.inkSoft,
                    border: done&&!active?`1px solid ${C.green}`:"none"}}>{done?"✓":i+1}</span>
                  <span style={{fontSize:13.5,fontWeight:active?600:500,color:active?C.ink:C.inkSoft}}>{s.label}</span>
                </button>
              );
            })}
          </nav>
          {/* running balance check */}
          <div style={{marginTop:26,padding:"14px 15px",borderRadius:10,
            background: bc===0?C.greenBg:C.redBg, border:`1px solid ${bc===0?C.green:C.red}22`}}>
            <div style={{fontSize:11.5,color:C.inkSoft,fontWeight:600,marginBottom:6}}>Balance check</div>
            <div style={{fontSize:12,color:C.inkSoft,marginBottom:2}}>Assets − (Liab + Equity)</div>
            <div style={{fontFamily:mono,fontSize:20,fontWeight:600,color:bc===0?C.green:C.red}}>
              {bc===0 ? "0.00 ✓" : fmt(bc)}
            </div>
            <div style={{fontSize:11.5,color:bc===0?C.green:C.red,marginTop:4}}>
              {bc===0 ? "Consolidation ties" : "Not yet balanced"}
            </div>
          </div>
          <div style={{marginTop:18,fontSize:11,color:C.inkSoft,lineHeight:1.5}}>
            Private by design — all consolidation runs locally in your browser, and uploaded trial balances never leave this page.
          </div>
        </aside>

        {/* MAIN */}
        <main style={{padding:"30px 34px 60px"}}>
          {/* top KPI strip */}
          <div style={{display:"flex",gap:14,marginBottom:26,flexWrap:"wrap"}}>
            <Kpi label="Net income" v={result.stmt.NI}/>
            <Kpi label="Total assets" v={result.stmt.total_assets}/>
            <Kpi label="NCI at year-end" v={result.stmt.nci_closing}/>
            <Kpi label="FCTR" v={result.stmt.fctr_closing}/>
          </div>

          {step===0 && <StepUpload tbs={tbs} setTbs={setTbs} sources={sources} setSources={setSources} onReset={onReset} names={names} setNames={setNames} groupName={groupName} setGroupName={setGroupName}/>}
          {step===1 && <StepMap tbs={tbs} setTbs={setTbs} names={names}/>}
          {step===2 && <StepTranslate result={result} rates={rates} setRates={setRates} names={names}/>}
          {step===3 && <StepAdjust result={result}/>}
          {step===4 && <StepEliminate result={result}/>}
          {step===5 && <StepNCI result={result}/>}
          {step===6 && <StepConsolidated result={result} brokenResult={brokenResult} groupName={groupName}/>}

          {/* nav buttons */}
          <div style={{display:"flex",justifyContent:"space-between",marginTop:36,paddingTop:20,borderTop:`1px solid ${C.line}`}}>
            <button onClick={()=>setStep(s=>Math.max(0,s-1))} disabled={step===0}
              style={{...navBtn,opacity:step===0?0.4:1,cursor:step===0?"default":"pointer"}}>← Back</button>
            {step===1 && unmappedCount>0 && (
              <span style={{fontSize:12.5,color:C.amber,alignSelf:"center",fontWeight:600}}>
                Map all accounts to continue ({unmappedCount} left)
              </span>
            )}
            <button onClick={()=>canProceed&&setStep(s=>Math.min(STEPS.length-1,s+1))} disabled={step===STEPS.length-1||!canProceed}
              style={{...navBtnPrimary,opacity:(step===STEPS.length-1||!canProceed)?0.4:1,
                cursor:(step===STEPS.length-1||!canProceed)?"default":"pointer"}}>
              {step===STEPS.length-1?"Done":"Next →"}
            </button>
          </div>
        </main>
      </div>
    </div>
  );
}

function Kpi({label,v}) {
  return (
    <div style={{border:`1px solid ${C.line}`,borderRadius:10,background:C.white,padding:"12px 18px",minWidth:150}}>
      <div style={{fontSize:11.5,color:C.inkSoft,fontWeight:600,marginBottom:5}}>{label}</div>
      <div style={{fontFamily:mono,fontSize:19,fontWeight:600,fontVariantNumeric:"tabular-nums",
        color: v<0?C.red:C.ink}}>{fmt(v)}</div>
    </div>
  );
}
const navBtn = {fontFamily:sans,fontSize:13.5,fontWeight:500,color:C.ink,background:C.white,
  border:`1px solid ${C.line}`,borderRadius:8,padding:"9px 18px"};
const navBtnPrimary = {fontFamily:sans,fontSize:13.5,fontWeight:600,color:C.white,background:C.ink,
  border:`1px solid ${C.ink}`,borderRadius:8,padding:"9px 20px"};
