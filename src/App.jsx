import React, { useState, useMemo } from "react";
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from "recharts";

// ── CONFIG ────────────────────────────────────────────────────────────────────
const MAKE_WEBHOOK = "https://hook.us2.make.com/ixy5r7umily7fz8rq8big679w3vphlzo";
const GID = { forecast:"0", anb:"1814362881", assumptions:"1862177871", invoicing:"1851109536" };
const WEEKS_PER_PAGE = 13;
const S = { bg:"#f8fafc",border:"#e2e8f0",green:"#16a34a",red:"#dc2626",blue:"#2563eb",yellow:"#b45309",muted:"#64748b",purple:"#7c3aed",gBg:"#f0fdf4",rBg:"#fef2f2",bBg:"#eff6ff",yBg:"#fffbeb",pBg:"#f5f3ff" };

const WEEKS = [
  {s:"12/28/25",e:"1/2/26",fc:false},{s:"1/3/26",e:"1/9/26",fc:false},{s:"1/10/26",e:"1/16/26",fc:false},
  {s:"1/17/26",e:"1/23/26",fc:false},{s:"1/24/26",e:"1/30/26",fc:false},{s:"1/31/26",e:"2/6/26",fc:false},
  {s:"2/7/26",e:"2/13/26",fc:false},{s:"2/14/26",e:"2/20/26",fc:false},{s:"2/21/26",e:"2/27/26",fc:false},
  {s:"2/28/26",e:"3/6/26",fc:false},
  {s:"3/7/26",e:"3/13/26",fc:true},{s:"3/14/26",e:"3/20/26",fc:true},{s:"3/21/26",e:"3/27/26",fc:true},
  {s:"3/28/26",e:"4/3/26",fc:true},{s:"4/4/26",e:"4/10/26",fc:true},{s:"4/11/26",e:"4/17/26",fc:true},
  {s:"4/18/26",e:"4/24/26",fc:true},{s:"4/25/26",e:"5/1/26",fc:true},{s:"5/2/26",e:"5/8/26",fc:true},
  {s:"5/9/26",e:"5/15/26",fc:true},{s:"5/16/26",e:"5/22/26",fc:true},{s:"5/23/26",e:"5/29/26",fc:true},
];
const NW = WEEKS.length;
const LAST_ACTUAL = WEEKS.map((w,i)=>({...w,i})).filter(w=>!w.fc).pop().i;

// ── HELPERS ───────────────────────────────────────────────────────────────────
const fmt = v => { if(!v&&v!==0) return "-"; const a=Math.abs(v); return v<0?`(${a.toLocaleString("en-US",{maximumFractionDigits:0})})`:`${a.toLocaleString("en-US",{maximumFractionDigits:0})}`; };
const fmtD = v => { if(!v&&v!==0) return "-"; return v<0?`▼ ${Math.abs(v).toLocaleString("en-US",{maximumFractionDigits:0})}`:`▲ ${v.toLocaleString("en-US",{maximumFractionDigits:0})}`; };
const parseAmt = s => { if(!s&&s!==0) return null; const n=parseFloat(String(s).replace(/[$,\s()]/g,"")); return isNaN(n)?null:(String(s).includes("(")?-n:n); };
const pad = (a,l) => { const r=[...a]; while(r.length<l) r.push(null); return r.slice(0,l); };

// Parse date string including "Sun Mar 08 2026 00:00:00 GMT-0300 (...)" format
const parseDate = str => {
  if(!str) return null;
  const m = str.match(/([A-Za-z]+)\s+(\d{1,2})\s+(\d{4})/);
  if(m) { const d=new Date(`${m[1]} ${m[2]} ${m[3]}`); return isNaN(d)?null:d; }
  const d = new Date(str); return isNaN(d)?null:d;
};

const dateToWeekIdx = str => {
  const d = parseDate(str); if(!d) return -1;
  const y=d.getFullYear(),mo=d.getMonth(),dy=d.getDate();
  for(let i=0;i<WEEKS.length;i++){
    const s=new Date(WEEKS[i].s),e=new Date(WEEKS[i].e);
    const sy=s.getFullYear(),sm=s.getMonth(),sd=s.getDate();
    const ey=e.getFullYear(),em=e.getMonth(),ed=e.getDate();
    const ok1=(y>sy)||(y===sy&&mo>sm)||(y===sy&&mo===sm&&dy>=sd);
    const ok2=(y<ey)||(y===ey&&mo<em)||(y===ey&&mo===em&&dy<=ed);
    if(ok1&&ok2) return i;
  }
  return -1;
};

// CSV split respecting quoted fields
const csvSplit = line => {
  const cols=[]; let c="",q=false;
  for(let i=0;i<=line.length;i++){
    const ch=i<line.length?line[i]:"";
    if(ch==='"'){q=!q;continue;}
    if((ch===","||i===line.length)&&!q){cols.push(c.trim());c="";}
    else c+=ch;
  }
  return cols;
};

const parseTable = text => {
  const lines=text.trim().split("\n").filter(l=>l.trim());
  if(lines.length<2) return [];
  const sep=lines[0].includes("\t")?"\t":",";
  const hdrs=lines[0].split(sep).map(h=>h.trim().replace(/"/g,""));
  return lines.slice(1).map(line=>{
    const cols=sep==="\t"?line.split("\t").map(c=>c.trim()):csvSplit(line);
    const o={}; hdrs.forEach((h,i)=>{o[h]=cols[i]||"";}); return o;
  });
};

// ── MAPPING ───────────────────────────────────────────────────────────────────
const BANK_RULES = [
  {m:["5/3 bank","stripe","cherry","repeatmd"],l:"Injectable/Skin Income"},
  {m:["trueaesthetics","true aesthetics"],l:"TrueAesthetics"},
  {m:["nitra"],l:"Nitra"},{m:["vantiv"],l:"Bank & Merchant Fees"},
  {m:["gusto","dd sweep"],l:"Wages"},{m:["tax sweep"],l:"Payroll Taxes"},
  {m:["vestwell"],l:"Benefits"},{m:["venmo","trilogy medical"],l:"Misc- TBD"},
  {m:["holland commercial","the den"],l:"Rent"},{m:["jpmorgan chase"],l:"Car Loan"},
  {m:["delta dental","insur prem"],l:"Insurance"},{m:["xcel energy"],l:"Utilities"},
  {m:["co dept revenue taxpayment"],l:"Taxes, Licenses & Fees"},
  {m:["easy track","easytrack"],l:"Payroll Fees"},
  {m:["aspire galderma","galderma"],l:"Rebates"},
  {m:["advs ed serv","collegeinvest"],l:"Personal/Distributions"},
  {m:["capital one"],l:"Capital One"},{m:["amex","american express"],l:"Amex"},
];
const INV_RULES = [
  {m:["tox membership","true tox","tox 1 member"],l:"Tox Membership"},
  {m:["beauty hive membership","bh membership"],l:"BH Membership"},
  {m:["weight loss membership","weightloss membership","restorative dose"],l:"Weightloss Membership"},
  {m:["gift card","card decline fee","deposit redeemed"],l:"Gift Cards/Deposit Redeemed"},
];
const STAFF_MAP = {"taylor campbell":"Taylor Campbell","emily kurtz":"Emily Kurtz","leah barr":"Leah Barr","rico alvarado":"Rico Alvarado"};
const STAFF_RATES = {"Emily Kurtz":3000,"Taylor Campbell":4800,"Leah Barr":1000,"Rico Alvarado":450};
const TOX_PRICE=175, WL_PRICE=455;

const IN_MAP = {"Injectable/Skin Income":"inj","Rebates":"reb","Gift Cards/Deposits Purchased":"gft","Gift Cards/Deposit Redeemed":"grd","Taylor Campbell":"tc","Emily Kurtz":"ek","Leah Barr":"lb","Rico Alvarado":"ra","BH Membership":"bhm","Tox Membership":"tox","Weightloss Membership":"wlm","Skincare Products":"skn"};
const OUT_MAP = {"TrueAesthetics":"ta","Nitra":"nit","Capital One":"cap","Amex":"amx","Wages":"wag","Payroll Taxes":"pt","Payroll Fees":"pf","Benefits":"ben","Rent":"rnt","Utilities":"utl","Bank & Merchant Fees":"bkf","Insurance":"ins","Car Loan":"cl","Taxes, Licenses & Fees":"tax","Misc- TBD":"misc","Personal/Distributions":"dist","Dues & Subscriptions":"dues","Repairs & Maintenance":"rm","Image First":"imgf","OKC Location":"okc"};

const matchRule = (rules,...f) => { const c=f.join(" ").toLowerCase(); for(const r of rules){if(r.m.some(m=>c.includes(m)))return r.l;} return null; };
const classifyBank = row => {
  if((row["Type"]||"").trim()) return {line:row["Type"].trim(),conf:"mapped"};
  const l=matchRule(BANK_RULES,row["Name"]||"",row["Memo/Description"]||"",row["Split"]||"");
  return l?{line:l,conf:"auto"}:{line:null,conf:"unknown"};
};
const classifyInv = row => {
  if((row["Cash Inflows"]||"").trim()) return {line:row["Cash Inflows"].trim(),conf:"mapped"};
  const prod=(row["Service/Product"]||"").toLowerCase(), staff=(row["Staff"]||"").toLowerCase().trim();
  const tax=parseAmt(row["Tax"]), total=parseAmt(row["Total Due"]);
  if(!total) return {line:null,conf:"skip"};
  const pl=matchRule(INV_RULES,prod); if(pl) return {line:pl,conf:"auto"};
  if(tax&&Math.abs(tax)>0) return {line:"Skincare Products",conf:"auto"};
  if(STAFF_MAP[staff]) return {line:STAFF_MAP[staff],conf:"auto"};
  return {line:null,conf:"unknown"};
};

// ── FORECAST PARSERS ──────────────────────────────────────────────────────────
const parseFcExpenses = text => {
  const totals={};
  Object.values(OUT_MAP).forEach(id=>{totals[id]=pad([],NW);});
  text.split("\n").forEach((line,li)=>{
    if(li===0||!line.trim()) return;
    const cols=csvSplit(line);
    if(cols.length<5) return;
    // cols: Name=0, Account=1, Date=2, Amount=3, Map=4
    const mapRaw=(cols[4]||"").replace(/^"|"$/g,"").trim();
    const amt=Math.abs(parseFloat((cols[3]||"").replace(/[^0-9.]/g,""))||0);
    const dateStr=(cols[2]||"").trim();
    if(!mapRaw||!amt) return;
    const outKey=Object.keys(OUT_MAP).find(k=>k.trim().toLowerCase()===mapRaw.toLowerCase());
    if(!outKey) return;
    const wi=dateToWeekIdx(dateStr);
    if(wi>=0&&wi<NW&&WEEKS[wi].fc) totals[OUT_MAP[outKey]][wi]=(totals[OUT_MAP[outKey]][wi]||0)+amt;
  });
  Object.keys(totals).forEach(id=>{totals[id]=totals[id].map(v=>v||null);});
  return totals;
};

const parseFcInflows = text => {
  const lines=text.trim().split("\n").filter(l=>l.trim());
  if(lines.length<2) return {};
  const sep=lines[0].includes("\t")?"\t":",";
  const rawHdrs=lines[0].split(sep).map(h=>h.trim().replace(/"/g,""));
  const wkHdrs=rawHdrs.slice(2);
  const hdrWi=wkHdrs.map(h=>{ const d=parseDate(h); if(!d) return -1; for(let i=0;i<WEEKS.length;i++){const s=new Date(WEEKS[i].s),e=new Date(WEEKS[i].e);const y=d.getFullYear(),mo=d.getMonth(),dy=d.getDate();const sy=s.getFullYear(),sm=s.getMonth(),sd=s.getDate(),ey=e.getFullYear(),em=e.getMonth(),ed=e.getDate();if((y>sy||(y===sy&&mo>sm)||(y===sy&&mo===sm&&dy>=sd))&&(y<ey||(y===ey&&mo<em)||(y===ey&&mo===em&&dy<=ed)))return i;}return -1;});
  const res={}; Object.values(IN_MAP).forEach(id=>{res[id]=pad([],NW);});
  const toxAcc=pad([],NW), wlAcc=pad([],NW);
  let sec=null;
  lines.slice(1).forEach(line=>{
    const cols=sep==="\t"?line.split("\t").map(c=>c.trim()):csvSplit(line);
    const label=(cols[0]||"").trim(), rpd=parseAmt(cols[1]), lbl=label.toLowerCase();
    if(lbl.includes("tox membership")&&!rpd){sec="tox";return;}
    if((lbl.includes("weightloss")||lbl.includes("weight loss"))&&!rpd){sec="wl";return;}
    if(!lbl) return;
    const sk=Object.keys(STAFF_MAP).find(k=>lbl===k);
    if(sk&&rpd){
      const rate=STAFF_RATES[STAFF_MAP[sk]]||rpd, id=IN_MAP[STAFF_MAP[sk]];
      wkHdrs.forEach((_,hi)=>{ const wi=hdrWi[hi]; if(wi<0||!WEEKS[wi].fc) return; const days=parseAmt(cols[hi+2]); if(!days) return; res[id][wi]=(res[id][wi]||0)+Math.round(days*rate); res["inj"][wi]=(res["inj"][wi]||0)+Math.round(days*rate); });
      sec=null; return;
    }
    if(sec==="tox"&&(lbl==="new"||lbl.includes("re-occur"))){
      wkHdrs.forEach((_,hi)=>{ const wi=hdrWi[hi]; if(wi<0||!WEEKS[wi].fc) return; const q=parseAmt(cols[hi+2]); if(q) toxAcc[wi]=(toxAcc[wi]||0)+q; }); return;
    }
    if(sec==="wl"&&(lbl==="new"||lbl.includes("re-occur"))){
      wkHdrs.forEach((_,hi)=>{ const wi=hdrWi[hi]; if(wi<0||!WEEKS[wi].fc) return; const q=parseAmt(cols[hi+2]); if(q) wlAcc[wi]=(wlAcc[wi]||0)+q; }); return;
    }
  });
  toxAcc.forEach((q,i)=>{ if(q) res["tox"][i]=Math.round(q*TOX_PRICE); });
  wlAcc.forEach((q,i)=>{ if(q) res["wlm"][i]=Math.round(q*WL_PRICE); });
  Object.keys(res).forEach(id=>{res[id]=res[id].map(v=>v||null);});
  return res;
};

// ── INITIAL DATA ──────────────────────────────────────────────────────────────
const INIT_IN=[
  {id:"inj",label:"Injectable/Skin Income",      v:pad([11464,20275,27483,22791,28862,31773,23413,21447,21104,28483],NW)},
  {id:"tc", label:"  Taylor Campbell",            v:pad([11464,6566,11163,11896,16005,15788,13591,5518,7257,14240],NW)},
  {id:"ek", label:"  Emily Kurtz",                v:pad([null,13709,16320,10895,12857,15985,9822,15929,13847,14243],NW)},
  {id:"lb", label:"  Leah Barr",                  v:pad([2060,4430,3026,1660,680,2378,3452,2323,3875,9270],NW)},
  {id:"ra", label:"  Rico Alvarado",              v:pad([null,1375,1365,1205,2290,615,650,1525,1870,1825],NW)},
  {id:"bhm",label:"BH Membership",               v:pad([18000,null,null,12100,394,17075,null,275,null,17250],NW)},
  {id:"tox",label:"Tox Membership",              v:pad([608,3955,34503,5123,1067,3066,1896,35963,1472,null],NW)},
  {id:"wlm",label:"Weightloss Membership",       v:pad([9775,2050,3855,null,null,11275,1341,4550,880,9480],NW)},
  {id:"gft",label:"Gift Cards/Deposits Purchased",v:pad([400,1375,525,2050,855,425,215,275,400,1025],NW)},
  {id:"skn",label:"Skincare Products",           v:pad([5452,6163,4357,2720,7054,10039,4569,5159,6653,6311],NW)},
  {id:"reb",label:"Rebates",                     v:pad([null,500,null,null,2300,null,null,null,null,null],NW)},
  {id:"grd",label:"Gift Cards/Deposit Redeemed", v:pad([-7038,-11624,-5296,-11025,-10940,-6590,-6412,-6652,-1870,-12390],NW)},
  {id:"inv",label:"Invoice to Cash Timing",      v:pad([-32995,21650,-35380,38514,-13693,-221,-1460,-2848,-868,-2239],NW)},
  {id:"oth",label:"Other Cash Inflows",          v:pad([],NW)},
];

// ── REEMPLAZAR INIT_OUT ───────────────────────────────────────────────────────
const INIT_OUT=[
  {sec:"Inventory",rows:[
    {id:"ta", label:"TrueAesthetics",  v:pad([7154,null,null,3408,22357,7135,3913,null,34819,13423],NW)},
    {id:"nit",label:"Nitra",           v:pad([8000,10000,null,10000,10000,10000,10000,10000,10000,10000],NW)},
    {id:"cap",label:"Capital One",     v:pad([null,null,null,14597,null,null,null,6522,null,null],NW)},
    {id:"amx",label:"Amex",            v:pad([null,null,9799,null,null,null,6859,null,null,null],NW)},
  ]},
  {sec:"Personnel",rows:[
    {id:"wag",label:"Wages",           v:pad([null,15293,null,13169,null,17152,null,12174,null,16694],NW)},
    {id:"pt", label:"Payroll Taxes",   v:pad([null,8091,null,16296,null,10437,null,6117,9614,10027],NW)},
    {id:"pf", label:"Payroll Fees",    v:pad([null,250,null,null,null,null,200,null,null,25],NW)},
    {id:"ben",label:"Benefits",        v:pad([null,null,5238,null,35,null,4837,null,null,35],NW)},
  ]},
  {sec:"Facilities",rows:[
    {id:"rnt",label:"Rent",            v:pad([null,5630,null,null,null,5626,null,null,null,5630],NW)},
    {id:"utl",label:"Utilities",       v:pad([null,504,null,null,null,551,null,null,null,null],NW)},
  ]},
  {sec:"Other Expenses",rows:[
    {id:"bkf",label:"Bank & Merchant Fees",    v:pad([null,7531,null,null,null,6974,null,null,null,7398],NW)},
    {id:"ins",label:"Insurance",               v:pad([null,353,1638,null,null,353,1638,null,null,353],NW)},
    {id:"cl", label:"Car Loan",                v:pad([null,1783,null,null,null,1783,null,null,null,1783],NW)},
    {id:"tax",label:"Taxes & Licenses",        v:pad([null,null,null,1501,167,null,32,894,3044,576],NW)},
    {id:"misc",label:"Misc - TBD",             v:pad([60,309,78,920,450,60,815,157,632,836],NW)},
    {id:"dist",label:"Personal/Distributions", v:pad([1133,8165,382,null,90,8303,88,null,null,8303],NW)},
    {id:"dues",label:"Dues & Subscriptions",   v:pad([null,null,null,null,null,null,null,null,null,null],NW)},
    {id:"rm",  label:"Repairs & Maintenance",  v:pad([965,null,965,null,965,null,null,965,null,null],NW)},
    {id:"imgf",label:"Image First",            v:pad([130,null,null,null,null,null,90,null,null,null],NW)},
    {id:"okc", label:"OKC Location",           v:pad([2300,null,null,null,null,null,null,null,null,null],NW)},
  ]},
];
const BB = pad([49856,33794,21705,20611,36427,21939,17545,17547,50978,27790],NW);
const ALL_LINES = [...INIT_IN.map(r=>r.label.trim()),...INIT_OUT.flatMap(s=>s.rows.map(r=>r.label))];

// ── FETCH ─────────────────────────────────────────────────────────────────────
const fetchSheet = async gid => {
  const res = await fetch(`${MAKE_WEBHOOK}?gid=${gid}`,{method:"GET",mode:"cors"});
  if(!res.ok) throw new Error(`HTTP ${res.status}`);
  const text = await res.text();
  if(!text||text.length<5) throw new Error("Empty response");
  return text;
};

// ── SUBCOMPONENTS ─────────────────────────────────────────────────────────────
function KPI({label,value,sub,color,bg}) {
  return React.createElement("div",{style:{background:bg,border:"1px solid #e2e8f0",borderRadius:10,padding:"10px 14px",flex:1,minWidth:130}},
    React.createElement("div",{style:{fontSize:10,color:"#64748b",marginBottom:3}},label),
    React.createElement("div",{style:{fontSize:20,fontWeight:700,color}},value),
    sub&&React.createElement("div",{style:{fontSize:10,color:"#64748b",marginTop:2}},sub)
  );
}

// ── APP ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab,setTab]=useState("table");
  const [inflows,setInflows]=useState(INIT_IN);
  const [outflows,setOutflows]=useState(INIT_OUT);
  const [fcSnapshot,setFcSnapshot]=useState(null);
  const [page,setPage]=useState(0);
  const [collapsed,setCollapsed]=useState({});
  const [editCell,setEditCell]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [syncState,setSyncState]=useState({anb:null,inv:null,fc:null});
  const [syncMsg,setSyncMsg]=useState(null);
  const [unkModal,setUnkModal]=useState(false);
  const [pending,setPending]=useState({type:null,rows:[],wi:null,unk:[]});
  const [manMap,setManMap]=useState({});
  const [pasteModal,setPasteModal]=useState(false);
  const [pasteExp,setPasteExp]=useState("");
  const [pasteInf,setPasteInf]=useState("");

  const infTot=useMemo(()=>Array.from({length:NW},(_,i)=>inflows.reduce((s,r)=>s+(r.v[i]||0),0)),[inflows]);
  const outTot=useMemo(()=>Array.from({length:NW},(_,i)=>outflows.flatMap(s=>s.rows).reduce((s,r)=>s+(r.v[i]||0),0)),[outflows]);
  const beginBal=useMemo(()=>{const b=[...BB];for(let i=1;i<NW;i++)b[i]=b[i-1]+infTot[i-1]-outTot[i-1];return b;},[infTot,outTot]);
  const endBal=useMemo(()=>Array.from({length:NW},(_,i)=>beginBal[i]+infTot[i]-outTot[i]),[beginBal,infTot,outTot]);
  const totalPages=Math.ceil(NW/WEEKS_PER_PAGE);
  const pageWeeks=WEEKS.slice(page*WEEKS_PER_PAGE,(page+1)*WEEKS_PER_PAGE).map((w,j)=>({...w,i:page*WEEKS_PER_PAGE+j}));
  const chartData=WEEKS.map((w,i)=>({name:w.s,Inflows:infTot[i],Outflows:outTot[i],Balance:endBal[i]}));

  // Apply classified rows to cash flow
  const applyRows = (rows,amtKey,wi) => {
    const totals={};
    rows.forEach(r=>{ if(!r._line) return; const a=parseAmt(r[amtKey]); if(!a) return; totals[r._line]=(totals[r._line]||0)+a; });
    setInflows(prev=>prev.map(r=>{ const k=Object.keys(IN_MAP).find(k=>IN_MAP[k]===r.id); if(!k||totals[k]===undefined)return r; return{...r,v:r.v.map((v,i)=>i===wi?(totals[k]):v)}; }));
    setOutflows(prev=>prev.map(s=>({...s,rows:s.rows.map(r=>{ const k=Object.keys(OUT_MAP).find(k=>OUT_MAP[k]===r.id); if(!k||totals[k]===undefined)return r; return{...r,v:r.v.map((v,i)=>i===wi?Math.abs(totals[k]):v)}; })})));
  };

  // Apply forecast maps
  const applyFc = (expMap,infMap) => {
    setOutflows(prev=>prev.map(s=>({...s,rows:s.rows.map(r=>{ const nv=expMap[r.id]; if(!nv)return r; return{...r,v:r.v.map((old,i)=>WEEKS[i].fc?(nv[i]??null):old)}; })})));
    setInflows(prev=>prev.map(r=>{ const nv=infMap[r.id]; if(!nv)return r; return{...r,v:r.v.map((old,i)=>WEEKS[i].fc?(nv[i]??null):old)}; }));
    const snap={inflows:{},outflows:{}};
    Object.keys(infMap).forEach(id=>{snap.inflows[id]=[...infMap[id]];});
    Object.keys(expMap).forEach(id=>{snap.outflows[id]=[...expMap[id]];});
    setFcSnapshot(snap);
    const ec=Object.values(expMap).flat().filter(Boolean).length;
    const ic=Object.values(infMap).flat().filter(Boolean).length;
    setSyncState(s=>({...s,fc:"done"}));
    setSyncMsg(`✓ Forecasts synced — ${ec} expense values, ${ic} inflow values`);
  };

  const syncANB = async () => {
    setSyncState(s=>({...s,anb:"loading"}));setSyncMsg(null);
    try {
      const text=await fetchSheet(GID.anb);
      const rows=parseTable(text);
      const wi=LAST_ACTUAL;
      const ws=new Date(WEEKS[wi].s),we=new Date(WEEKS[wi].e);we.setHours(23,59,59);
      const wRows=rows.filter(r=>{const d=parseDate(r["Date"]||"");return d&&d>=ws&&d<=we;});
      const cl=wRows.map(r=>{const res=classifyBank(r);return{...r,_line:res.line,_conf:res.conf};});
      const unk=cl.filter(r=>r._conf==="unknown"&&parseAmt(r["Amount"]));
      if(unk.length){setPending({type:"anb",rows:cl,wi,unk});setUnkModal(true);}
      else{applyRows(cl,"Amount",wi);setSyncState(s=>({...s,anb:"done"}));setSyncMsg(`✓ ANB synced to ${WEEKS[wi].s}–${WEEKS[wi].e}`);}
    }catch(e){setSyncState(s=>({...s,anb:"error"}));setSyncMsg(`ANB error: ${e.message}`);}
  };

  const syncInv = async () => {
    setSyncState(s=>({...s,inv:"loading"}));setSyncMsg(null);
    try {
      const text=await fetchSheet(GID.invoicing);
      const rows=parseTable(text);
      const wi=LAST_ACTUAL;
      const ws=new Date(WEEKS[wi].s),we=new Date(WEEKS[wi].e);we.setHours(23,59,59);
      const wRows=rows.filter(r=>{const d=parseDate(r["Date"]||"");return d&&d>=ws&&d<=we;});
      const cl=wRows.map(r=>{const res=classifyInv(r);return{...r,_line:res.line,_conf:res.conf};});
      const unk=cl.filter(r=>r._conf==="unknown"&&parseAmt(r["Total Due"]));
      if(unk.length){setPending({type:"inv",rows:cl,wi,unk});setUnkModal(true);}
      else{applyRows(cl,"Total Due",wi);setSyncState(s=>({...s,inv:"done"}));setSyncMsg(`✓ Invoicing synced to ${WEEKS[wi].s}–${WEEKS[wi].e}`);}
    }catch(e){setSyncState(s=>({...s,inv:"error"}));setSyncMsg(`Invoicing error: ${e.message}`);}
  };

  const syncFc = async () => {
    setSyncState(s=>({...s,fc:"loading"}));setSyncMsg(null);
    try {
      const [expText,infText]=await Promise.all([fetchSheet(GID.forecast),fetchSheet(GID.assumptions)]);
      applyFc(parseFcExpenses(expText),parseFcInflows(infText));
    }catch(e){setSyncState(s=>({...s,fc:"error"}));setSyncMsg(`Forecast error: ${e.message}`);}
  };

  const syncFcPaste = () => {
    setSyncState(s=>({...s,fc:"loading"}));setSyncMsg(null);
    applyFc(parseFcExpenses(pasteExp),parseFcInflows(pasteInf));
    setPasteModal(false);
  };

  const confirmUnk = () => {
    const final=pending.rows.map((r,i)=>({...r,_line:manMap[i]||r._line}));
    applyRows(final,pending.type==="anb"?"Amount":"Total Due",pending.wi);
    setSyncState(s=>({...s,[pending.type]:"done"}));
    setSyncMsg(`✓ ${pending.type==="anb"?"ANB":"Invoicing"} synced with manual mappings`);
    setUnkModal(false);setManMap({});
  };

  const commitEdit=(sec,rowId,wi,val)=>{
    const num=val===""?null:parseFloat(val)||null;
    if(sec==="in")setInflows(prev=>prev.map(r=>r.id===rowId?{...r,v:r.v.map((v,i)=>i===wi?num:v)}:r));
    else setOutflows(prev=>prev.map(s=>({...s,rows:s.rows.map(r=>r.id===rowId?{...r,v:r.v.map((v,i)=>i===wi?num:v)}:r)})));
    setEditCell(null);
  };

  const handleExport=()=>{
    const hdr=["Line",...WEEKS.map(w=>`${w.s}-${w.e}`)].join(",");
    const csv=[hdr,["Beginning Balance",...beginBal].join(","),...inflows.map(r=>[r.label,...r.v.map(v=>v??0)].join(",")),...outflows.flatMap(s=>s.rows.map(r=>[r.label,...r.v.map(v=>v??0)].join(","))),["Ending Balance",...endBal].join(",")].join("\n");
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));a.download="RI_cashflow.csv";a.click();
  };

  // ── STYLES ──
  const th=fc=>({padding:"4px 8px",textAlign:"right",minWidth:90,fontSize:11,fontWeight:600,background:fc?"#fef9ec":"#eff6ff",borderLeft:"1px solid #e2e8f0",color:fc?"#b45309":"#1d4ed8",whiteSpace:"nowrap"});
  const td=fc=>({padding:"3px 8px",textAlign:"right",fontSize:12,background:fc?"#fffdf5":"#f8fbff",borderLeft:"1px solid #e2e8f0",whiteSpace:"nowrap",cursor:"pointer",color:"#1e293b"});
  const lbl=(sub,bold,color,bg)=>({padding:`3px 8px 3px ${sub?"20px":"8px"}`,fontSize:12,fontWeight:bold?700:400,color:color||(sub?"#64748b":"#1e293b"),minWidth:200,position:"sticky",left:0,background:bg||"#fff",zIndex:2,borderRight:"2px solid #e2e8f0",whiteSpace:"nowrap"});
  const sHdr=(bg,col)=>({padding:"5px 8px",fontSize:11,fontWeight:700,color:col||"#1d4ed8",textTransform:"uppercase",letterSpacing:1,position:"sticky",left:0,background:bg||"#eff6ff",zIndex:2,borderRight:"2px solid #e2e8f0"});
  const subS={padding:"4px 8px 4px 12px",fontSize:10,fontWeight:700,color:"#64748b",position:"sticky",left:0,background:"#f8fafc",zIndex:2,borderRight:"2px solid #e2e8f0",cursor:"pointer"};
  const btnS=(active,col)=>({padding:"6px 14px",borderRadius:7,border:"1px solid #e2e8f0",cursor:"pointer",fontSize:12,fontWeight:600,background:active?(col||"#2563eb"):"#f1f5f9",color:active?"#fff":"#475569"});
  const inp={padding:"8px 12px",borderRadius:8,border:"1px solid #e2e8f0",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box"};
  const syncBtn=(label,key,onClick,color)=>{
    const st=syncState[key];
    return React.createElement("button",{onClick,disabled:st==="loading",style:{display:"flex",alignItems:"center",gap:5,padding:"6px 14px",borderRadius:7,border:"1px solid #e2e8f0",cursor:st==="loading"?"default":"pointer",fontSize:12,fontWeight:600,background:st==="done"?"#dcfce7":st==="loading"?"#f1f5f9":"#fff",color:st==="done"?S.green:st==="loading"?S.muted:(color||S.blue),transition:"all .2s"}},
      st==="loading"?"⟳ Syncing...":st==="done"?`✓ ${label}`:`⟳ ${label}`
    );
  };

  const renderCell=(section,row,wi)=>{
    const v=row.v[wi], isEd=editCell&&editCell.sec===section&&editCell.id===row.id&&editCell.wi===wi;
    const fc=WEEKS[wi]&&WEEKS[wi].fc, color=v<0?"#b91c1c":v>0?"#1e293b":"#94a3b8";
    if(isEd) return React.createElement("td",{key:wi,style:{...td(fc),color}},
      React.createElement("input",{autoFocus:true,value:editVal,onChange:e=>setEditVal(e.target.value),
        onBlur:()=>commitEdit(section,row.id,wi,editVal),onKeyDown:e=>{if(e.key==="Enter")commitEdit(section,row.id,wi,editVal);},
        style:{width:68,background:"#eff6ff",border:"1px solid #2563eb",borderRadius:4,color:"#1e293b",fontSize:11,padding:"1px 4px",textAlign:"right"}})
    );
    return React.createElement("td",{key:wi,style:{...td(fc),color},onClick:()=>{setEditCell({sec:section,id:row.id,wi});setEditVal(v===null?"":String(v));}},fmt(v));
  };

  const el=(t,p,...c)=>React.createElement(t,p,...c.flat().filter(v=>v!==false&&v!=null&&v!==undefined));
  const modal=(content)=>el("div",{style:{position:"fixed",inset:0,background:"rgba(0,0,0,0.28)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}},
    el("div",{style:{background:"#fff",borderRadius:14,padding:24,width:600,maxWidth:"96vw",maxHeight:"88vh",overflowY:"auto",boxShadow:"0 8px 32px rgba(0,0,0,0.14)"},...{}},content)
  );

  // Variance rows
  const varWeeks=WEEKS.map((w,i)=>({...w,i})).filter(w=>!w.fc).slice(-6);
  const varRows=useMemo(()=>{
    if(!fcSnapshot) return [];
    const rows=[];
    inflows.forEach(row=>{
      const fc=fcSnapshot.inflows[row.id]; if(!fc) return;
      if(!varWeeks.some(w=>row.v[w.i]||fc[w.i])) return;
      rows.push({label:row.label,sub:row.label.startsWith("  "),type:"in",act:row.v,fc});
    });
    outflows.flatMap(s=>s.rows).forEach(row=>{
      const fc=fcSnapshot.outflows[row.id]; if(!fc) return;
      if(!varWeeks.some(w=>row.v[w.i]||fc[w.i])) return;
      rows.push({label:row.label,sub:false,type:"out",act:row.v,fc});
    });
    return rows;
  },[fcSnapshot,inflows,outflows]);

  return el("div",{style:{fontFamily:"system-ui,sans-serif",background:S.bg,minHeight:"100vh",color:"#1e293b"}},

    // Header
    el("div",{style:{padding:"10px 16px",borderBottom:"1px solid #e2e8f0",background:"#fff",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}},
      el("div",null,
        el("div",{style:{fontSize:10,color:S.muted,textTransform:"uppercase",letterSpacing:1}},"Restorative Injectables — Denver"),
        el("div",{style:{fontSize:17,fontWeight:700}},"Cash Flow Worksheet")
      ),
      el("div",{style:{display:"flex",gap:6,flexWrap:"wrap"}},
        el("button",{onClick:()=>setTab("table"),style:btnS(tab==="table")},"📋 Cash Flow"),
        el("button",{onClick:()=>setTab("variance"),style:btnS(tab==="variance",S.purple)},"📊 Forecast vs Actual"),
        el("button",{onClick:()=>setTab("chart"),style:btnS(tab==="chart")},"📈 Chart"),
        el("button",{onClick:handleExport,style:{...btnS(false),...{color:S.green,borderColor:"#bbf7d0",background:S.gBg}}},"⬇ Export CSV")
      )
    ),

    // Sync bar
    el("div",{style:{padding:"8px 16px",background:"#fff",borderBottom:"1px solid #e2e8f0",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}},
      syncBtn("Sync ANB (Outflows)","anb",syncANB,S.red),
      syncBtn("Sync Invoicing (Inflows)","inv",syncInv,S.green),
      syncBtn("Sync Forecasts","fc",syncFc,S.purple),
      el("button",{onClick:()=>setPasteModal(true),style:{...btnS(false),...{color:S.purple,borderColor:"#ddd6fe",fontSize:11}}},"📋 Paste Forecasts"),
      el("div",{style:{fontSize:11,color:S.muted,marginLeft:4}},"Last actual week · Forecast weeks"),
      syncMsg&&el("div",{style:{fontSize:11,color:syncMsg.includes("error")||syncMsg.includes("Error")?S.red:S.green,marginLeft:"auto"}},syncMsg)
    ),

    // KPIs
    el("div",{style:{display:"flex",gap:10,padding:"10px 16px",flexWrap:"wrap",background:"#f8fafc",borderBottom:"1px solid #e2e8f0"}},
      el(KPI,{label:`Balance — ${WEEKS[LAST_ACTUAL].s}`,value:fmt(endBal[LAST_ACTUAL]),color:"#1d4ed8",bg:S.bBg}),
      el(KPI,{label:"Inflows (last actual week)",value:fmt(infTot[LAST_ACTUAL]),color:S.green,bg:S.gBg}),
      el(KPI,{label:"Outflows (last actual week)",value:fmt(outTot[LAST_ACTUAL]),color:S.red,bg:S.rBg}),
      el(KPI,{label:"Projected end balance",value:fmt(endBal[NW-1]),sub:`(${WEEKS[NW-1].e})`,color:S.yellow,bg:S.yBg})
    ),

    // Paste Forecast Modal
    pasteModal&&modal(
      el("div",null,
        el("div",{style:{fontWeight:700,fontSize:15,marginBottom:4}},"📋 Paste Forecast Data"),
        el("div",{style:{fontSize:12,color:S.muted,marginBottom:14}},"Open each tab in Google Sheets → Ctrl+A → Ctrl+C → paste below."),
        el("div",{style:{marginBottom:12}},
          el("div",{style:{fontSize:12,fontWeight:700,color:S.red,marginBottom:4}},"📉 Expenses tab (Name | Account | Date | Amount | Map)"),
          el("textarea",{value:pasteExp,onChange:e=>setPasteExp(e.target.value),rows:6,placeholder:"Paste forecast expenses here...",style:{width:"100%",padding:8,borderRadius:8,border:"1px solid #fecaca",fontSize:11,boxSizing:"border-box",fontFamily:"monospace",resize:"vertical",outline:"none"}})
        ),
        el("div",{style:{marginBottom:16}},
          el("div",{style:{fontSize:12,fontWeight:700,color:S.purple,marginBottom:4}},"📈 Assumptions tab"),
          el("textarea",{value:pasteInf,onChange:e=>setPasteInf(e.target.value),rows:6,placeholder:"Paste assumptions here...",style:{width:"100%",padding:8,borderRadius:8,border:"1px solid #ddd6fe",fontSize:11,boxSizing:"border-box",fontFamily:"monospace",resize:"vertical",outline:"none"}})
        ),
        el("div",{style:{display:"flex",gap:8,justifyContent:"flex-end"}},
          el("button",{onClick:()=>setPasteModal(false),style:{padding:"8px 16px",borderRadius:8,border:"1px solid #e2e8f0",background:"#f1f5f9",cursor:"pointer",fontSize:13,color:"#475569"}},"Cancel"),
          el("button",{onClick:syncFcPaste,style:{padding:"8px 20px",borderRadius:8,border:"none",background:S.purple,cursor:"pointer",fontSize:13,fontWeight:600,color:"#fff"}},"Apply Forecasts")
        )
      )
    ),

    // Unknown rows modal
    unkModal&&modal(
      el("div",null,
        el("div",{style:{fontWeight:700,fontSize:15,marginBottom:4,color:"#b45309"}},`⚠ ${pending.unk.length} rows need manual mapping`),
        el("div",{style:{fontSize:12,color:S.muted,marginBottom:14}},"Assign a line to each row then click Apply."),
        el("div",{style:{overflowX:"auto"}},
          el("table",{style:{borderCollapse:"collapse",width:"100%",fontSize:12}},
            el("thead",null,el("tr",{style:{background:"#fffbeb"}},
              ...["Date","Name","Description","Amount","→ Line"].map(h=>el("th",{key:h,style:{padding:"6px 10px",textAlign:"left",borderBottom:"1px solid #e2e8f0",fontWeight:600,color:"#78350f"}},h))
            )),
            el("tbody",null,...pending.unk.map((row,idx)=>{
              const amt=parseAmt(row["Amount"]||row["Total Due"]);
              const oi=pending.rows.indexOf(row);
              return el("tr",{key:idx,style:{borderBottom:"1px solid #f1f5f9"}},
                el("td",{style:{padding:"5px 10px",color:S.muted}},row["Date"]),
                el("td",{style:{padding:"5px 10px",fontWeight:500}},(row["Name"]||row["Client Name"]||"—").slice(0,22)),
                el("td",{style:{padding:"5px 10px",color:S.muted,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}},(row["Memo/Description"]||row["Service/Product"]||"—").slice(0,40)),
                el("td",{style:{padding:"5px 10px",fontWeight:600,color:amt<0?"#b91c1c":"#16a34a",whiteSpace:"nowrap"}},fmt(amt)),
                el("td",{style:{padding:"5px 10px"}},
                  el("select",{value:manMap[oi]||"",onChange:e=>setManMap(m=>({...m,[oi]:e.target.value})),style:{width:"100%",padding:"4px 8px",borderRadius:6,border:"1px solid #e2e8f0",fontSize:12,background:"#fff"}},
                    el("option",{value:""},"— Skip —"),
                    ...ALL_LINES.map(l=>el("option",{key:l,value:l},l))
                  )
                )
              );
            }))
          )
        ),
        el("div",{style:{display:"flex",gap:8,justifyContent:"flex-end",marginTop:16}},
          el("button",{onClick:()=>{setUnkModal(false);setManMap({});setSyncState(s=>({...s,[pending.type]:null}));},style:{padding:"8px 16px",borderRadius:8,border:"1px solid #e2e8f0",background:"#f1f5f9",cursor:"pointer",fontSize:13,color:"#475569"}},"Cancel"),
          el("button",{onClick:confirmUnk,style:{padding:"8px 20px",borderRadius:8,border:"none",background:S.blue,cursor:"pointer",fontSize:13,fontWeight:600,color:"#fff"}},"Apply Mappings")
        )
      )
    ),

    // ── VARIANCE TAB ──
    tab==="variance"&&el("div",{style:{padding:16}},
      !fcSnapshot
        ?el("div",{style:{background:"#fff",borderRadius:12,border:"1px solid #e2e8f0",padding:32,textAlign:"center",color:S.muted}},
            el("div",{style:{fontSize:32,marginBottom:8}},"🔮"),
            el("div",{style:{fontWeight:600,marginBottom:4}},"No forecast data yet"),
            el("div",{style:{fontSize:12}},"Click Sync Forecasts or Paste Forecasts to load forecast data.")
          )
        :el("div",null,
            el("div",{style:{fontWeight:700,fontSize:15,marginBottom:12}},"Forecast vs Actual — Last 6 Weeks"),
            el("div",{style:{overflowX:"auto"}},
              el("table",{style:{borderCollapse:"collapse",minWidth:"100%",fontSize:12}},
                el("thead",null,el("tr",{style:{background:"#f8fafc"}},
                  el("th",{style:{...lbl(false,false,S.muted),fontSize:11,padding:"6px 8px",fontWeight:600}},"Line"),
                  ...varWeeks.flatMap(w=>[
                    el("th",{key:`f${w.i}`,style:{padding:"4px 8px",textAlign:"right",minWidth:78,fontSize:10,fontWeight:600,color:S.purple,background:S.pBg,borderLeft:"1px solid #e2e8f0",whiteSpace:"nowrap"}},`Fcst ${w.s}`),
                    el("th",{key:`a${w.i}`,style:{padding:"4px 8px",textAlign:"right",minWidth:78,fontSize:10,fontWeight:600,color:"#1d4ed8",background:S.bBg,borderLeft:"1px solid #e2e8f0",whiteSpace:"nowrap"}},`Act ${w.s}`),
                    el("th",{key:`d${w.i}`,style:{padding:"4px 8px",textAlign:"right",minWidth:65,fontSize:10,fontWeight:600,color:S.muted,background:"#f8fafc",borderLeft:"1px solid #e2e8f0",whiteSpace:"nowrap"}},"Δ"),
                  ])
                )),
                el("tbody",null,...varRows.map((row,ri)=>
                  el("tr",{key:ri,style:{borderBottom:"1px solid #f1f5f9",background:row.sub?"#fafafa":"#fff"}},
                    el("td",{style:lbl(row.sub,false)},row.label),
                    ...varWeeks.flatMap(w=>{
                      const fc=row.fc[w.i]||null, ac=row.act[w.i]||null;
                      const delta=(ac||0)-(fc||0);
                      const good=row.type==="in"?delta>=0:delta<=0;
                      const dc=delta===0||(!ac&&!fc)?S.muted:good?S.green:S.red;
                      return [
                        el("td",{key:`f${w.i}`,style:{padding:"3px 8px",textAlign:"right",background:S.pBg,borderLeft:"1px solid #e2e8f0",color:S.purple}},fmt(fc)),
                        el("td",{key:`a${w.i}`,style:{padding:"3px 8px",textAlign:"right",background:S.bBg,borderLeft:"1px solid #e2e8f0",color:"#1d4ed8",fontWeight:500}},fmt(ac)),
                        el("td",{key:`d${w.i}`,style:{padding:"3px 8px",textAlign:"right",background:"#f8fafc",borderLeft:"1px solid #e2e8f0",color:dc,fontWeight:600,fontSize:11}},fmtD(delta)),
                      ];
                    })
                  )
                ))
              )
            )
          )
    ),

    // ── CHART TAB ──
    tab==="chart"&&el("div",{style:{padding:16}},
      el("div",{style:{background:"#fff",borderRadius:12,border:"1px solid #e2e8f0",padding:16,marginBottom:12}},
        el("div",{style:{fontWeight:600,marginBottom:10,fontSize:13}},"Balance · Inflows · Outflows"),
        el(ResponsiveContainer,{width:"100%",height:220},
          el(LineChart,{data:chartData},el(CartesianGrid,{strokeDasharray:"3 3",stroke:"#e2e8f0"}),el(XAxis,{dataKey:"name",tick:{fontSize:9,fill:S.muted},interval:2}),el(YAxis,{tickFormatter:v=>`$${(v/1000).toFixed(0)}k`,tick:{fontSize:10,fill:S.muted}}),el(Tooltip,{formatter:v=>`$${Number(v).toLocaleString("en-US")}`,contentStyle:{background:"#fff",border:"1px solid #e2e8f0",borderRadius:8,fontSize:11}}),el(Legend,{wrapperStyle:{fontSize:11}}),el(Line,{type:"monotone",dataKey:"Balance",stroke:S.blue,strokeWidth:2,dot:{r:2}}),el(Line,{type:"monotone",dataKey:"Inflows",stroke:S.green,strokeWidth:1.5,dot:false}),el(Line,{type:"monotone",dataKey:"Outflows",stroke:S.red,strokeWidth:1.5,dot:false}))
        )
      ),
      el("div",{style:{background:"#fff",borderRadius:12,border:"1px solid #e2e8f0",padding:16}},
        el("div",{style:{fontWeight:600,marginBottom:10,fontSize:13}},"Inflows vs Outflows by Week"),
        el(ResponsiveContainer,{width:"100%",height:180},
          el(BarChart,{data:chartData},el(CartesianGrid,{strokeDasharray:"3 3",stroke:"#e2e8f0"}),el(XAxis,{dataKey:"name",tick:{fontSize:9,fill:S.muted},interval:2}),el(YAxis,{tickFormatter:v=>`$${(v/1000).toFixed(0)}k`,tick:{fontSize:10,fill:S.muted}}),el(Tooltip,{formatter:v=>`$${Number(v).toLocaleString("en-US")}`,contentStyle:{background:"#fff",border:"1px solid #e2e8f0",borderRadius:8,fontSize:11}}),el(Legend,{wrapperStyle:{fontSize:11}}),el(Bar,{dataKey:"Inflows",fill:S.green,radius:[3,3,0,0]}),el(Bar,{dataKey:"Outflows",fill:S.red,radius:[3,3,0,0]}))
        )
      )
    ),

    // ── TABLE TAB ──
    tab==="table"&&el("div",null,
      el("div",{style:{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 16px",background:"#fff",borderBottom:"1px solid #e2e8f0",flexWrap:"wrap",gap:6}},
        el("span",{style:{fontSize:11,color:S.muted}},"Click any cell to edit · Amber = Forecast"),
        el("div",{style:{display:"flex",gap:4,alignItems:"center"}},
          el("button",{onClick:()=>setPage(p=>Math.max(0,p-1)),disabled:page===0,style:{padding:"4px 12px",borderRadius:6,border:"1px solid #e2e8f0",cursor:page===0?"default":"pointer",fontSize:13,background:"#f1f5f9",color:page===0?"#cbd5e1":"#1e293b"}},"‹"),
          el("span",{style:{fontSize:11,color:S.muted,lineHeight:"28px"}},`Page ${page+1}/${totalPages}`),
          el("button",{onClick:()=>setPage(p=>Math.min(totalPages-1,p+1)),disabled:page===totalPages-1,style:{padding:"4px 12px",borderRadius:6,border:"1px solid #e2e8f0",cursor:page===totalPages-1?"default":"pointer",fontSize:13,background:"#f1f5f9",color:page===totalPages-1?"#cbd5e1":"#1e293b"}},"›")
        )
      ),
      el("div",{style:{overflowX:"auto",paddingBottom:24}},
        el("table",{style:{borderCollapse:"collapse",minWidth:"100%"}},
          el("thead",null,el("tr",{style:{background:"#f8fafc"}},
            el("th",{style:{...lbl(false,false,S.muted),fontSize:11,padding:"6px 8px",fontWeight:600}},"Line"),
            ...pageWeeks.map(w=>el("th",{key:w.i,style:th(w.fc)},el("div",{style:{fontSize:9}},w.fc?"Forecast":"Actual"),el("div",{style:{fontSize:9,fontWeight:400,color:"#94a3b8"}},w.s),el("div",{style:{fontSize:9,fontWeight:400,color:"#94a3b8"}},w.e)))
          )),
          el("tbody",null,
            ...[["Beginning Balance",beginBal,"#1d4ed8","#eff6ff"],["Ending Balance",endBal,"#1d4ed8","#eff6ff"]].map(([label,vals,c,bg])=>
              el("tr",{key:label,style:{background:bg,borderTop:"2px solid #e2e8f0"}},el("td",{style:lbl(false,true,c,bg)},label),...pageWeeks.map(w=>el("td",{key:w.i,style:{...td(w.fc),fontWeight:700,color:c,background:bg}},fmt(vals[w.i]))))
            ),
            el("tr",{key:"ih"},el("td",{colSpan:pageWeeks.length+1,style:sHdr("#eff6ff","#1d4ed8")},"Cash Inflows")),
            ...inflows.map(row=>el("tr",{key:row.id,style:{borderBottom:"1px solid #f1f5f9"}},el("td",{style:lbl(row.label.startsWith("  "),false)},row.label),...pageWeeks.map(w=>renderCell("in",row,w.i)))),
            el("tr",{key:"it",style:{background:S.gBg,borderTop:"2px solid #e2e8f0"}},el("td",{style:lbl(false,true,S.green,S.gBg)},"Total Cash Inflows"),...pageWeeks.map(w=>el("td",{key:w.i,style:{...td(w.fc),fontWeight:700,color:S.green,background:S.gBg}},fmt(infTot[w.i])))),
            el("tr",{key:"oh"},el("td",{colSpan:pageWeeks.length+1,style:sHdr(S.rBg,S.red)},"Cash Outflows")),
            ...outflows.flatMap((sec,si)=>{
              const isC=collapsed[sec.sec];
              const hdr=el("tr",{key:`sh${si}`,style:{background:"#f8fafc",cursor:"pointer"},onClick:()=>setCollapsed(c=>({...c,[sec.sec]:!c[sec.sec]}))},el("td",{colSpan:pageWeeks.length+1,style:subS},`${isC?"▶":"▼"} ${sec.sec}`));
              if(isC) return [hdr];
              return [hdr,...sec.rows.map(row=>el("tr",{key:row.id,style:{borderBottom:"1px solid #f1f5f9"}},el("td",{style:lbl(false,false)},row.label),...pageWeeks.map(w=>renderCell("out",row,w.i))))];
            }),
            el("tr",{key:"ot",style:{background:S.rBg,borderTop:"2px solid #e2e8f0"}},el("td",{style:lbl(false,true,S.red,S.rBg)},"Total Cash Outflows"),...pageWeeks.map(w=>el("td",{key:w.i,style:{...td(w.fc),fontWeight:700,color:S.red,background:S.rBg}},fmt(outTot[w.i])))),
            el("tr",{key:"chg",style:{background:"#f8fafc",borderTop:"1px solid #e2e8f0"}},el("td",{style:lbl(false,true,S.purple,"#f8fafc")},"Change in Cash"),...pageWeeks.map(w=>{const v=infTot[w.i]-outTot[w.i];return el("td",{key:w.i,style:{...td(w.fc),fontWeight:700,color:v>=0?S.green:S.red,background:"#f8fafc"}},fmt(v));}))
          )
        )
      )
    )
  );
}
