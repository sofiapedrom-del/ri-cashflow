import React, { useState, useMemo } from "react";
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from "recharts";

// ── CONFIG ────────────────────────────────────────────────────────────────────
const MAKE_WEBHOOK = "https://hook.us2.make.com/ixy5r7umily7fz8rq8big679w3vphlzo";
const GID = { forecast:"0", anb:"1814362881", assumptions:"1862177871", invoicing:"1851109536" };
const WEEKS_PER_PAGE = 13;
const S = { bg:"#f8fafc",border:"#e2e8f0",green:"#16a34a",red:"#dc2626",blue:"#2563eb",yellow:"#b45309",muted:"#64748b",purple:"#7c3aed",gBg:"#f0fdf4",rBg:"#fef2f2",bBg:"#eff6ff",yBg:"#fffbeb",pBg:"#f5f3ff" };

const WEEKS_INIT = [
  {s:"12/28/25",e:"1/2/26",fc:false},{s:"1/3/26",e:"1/9/26",fc:false},{s:"1/10/26",e:"1/16/26",fc:false},
  {s:"1/17/26",e:"1/23/26",fc:false},{s:"1/24/26",e:"1/30/26",fc:false},{s:"1/31/26",e:"2/6/26",fc:false},
  {s:"2/7/26",e:"2/13/26",fc:false},{s:"2/14/26",e:"2/20/26",fc:false},{s:"2/21/26",e:"2/27/26",fc:false},
  {s:"2/28/26",e:"3/6/26",fc:false},
  {s:"3/7/26",e:"3/13/26",fc:true},{s:"3/14/26",e:"3/20/26",fc:true},{s:"3/21/26",e:"3/27/26",fc:true},
  {s:"3/28/26",e:"4/3/26",fc:true},{s:"4/4/26",e:"4/10/26",fc:true},{s:"4/11/26",e:"4/17/26",fc:true},
  {s:"4/18/26",e:"4/24/26",fc:true},{s:"4/25/26",e:"5/1/26",fc:true},{s:"5/2/26",e:"5/8/26",fc:true},
  {s:"5/9/26",e:"5/15/26",fc:true},{s:"5/16/26",e:"5/22/26",fc:true},{s:"5/23/26",e:"5/29/26",fc:true},
];
const NW = WEEKS_INIT.length;

// ── HELPERS ───────────────────────────────────────────────────────────────────
const fmt = v => { if(!v&&v!==0) return "-"; const a=Math.abs(v); return v<0?`(${a.toLocaleString("en-US",{maximumFractionDigits:0})})`:`${a.toLocaleString("en-US",{maximumFractionDigits:0})}`; };
const fmtD = v => { if(!v&&v!==0) return "-"; return v<0?`▼ ${Math.abs(v).toLocaleString("en-US",{maximumFractionDigits:0})}`:`▲ ${v.toLocaleString("en-US",{maximumFractionDigits:0})}`; };
const parseAmt = s => { if(!s&&s!==0) return null; const n=parseFloat(String(s).replace(/[$,\s()]/g,"")); return isNaN(n)?null:(String(s).includes("(")?-n:n); };
const pad = (a,l) => { const r=[...a]; while(r.length<l) r.push(null); return r.slice(0,l); };
const parseDate = str => { if(!str) return null; const m=str.match(/([A-Za-z]+)\s+(\d{1,2})\s+(\d{4})/); if(m){const d=new Date(`${m[1]} ${m[2]} ${m[3]}`);return isNaN(d)?null:d;} const d=new Date(str);return isNaN(d)?null:d; };
const dateToWeekIdx = (str,weeks) => { const d=parseDate(str);if(!d)return -1; const y=d.getFullYear(),mo=d.getMonth(),dy=d.getDate(); for(let i=0;i<weeks.length;i++){const s=new Date(weeks[i].s),e=new Date(weeks[i].e);const sy=s.getFullYear(),sm=s.getMonth(),sd=s.getDate(),ey=e.getFullYear(),em=e.getMonth(),ed=e.getDate();const ok1=(y>sy)||(y===sy&&mo>sm)||(y===sy&&mo===sm&&dy>=sd);const ok2=(y<ey)||(y===ey&&mo<em)||(y===ey&&mo===em&&dy<=ed);if(ok1&&ok2)return i;}return -1; };
const csvSplit = line => { const cols=[];let c="",q=false;for(let i=0;i<=line.length;i++){const ch=i<line.length?line[i]:"";if(ch==='"'){q=!q;continue;}if((ch===","||i===line.length)&&!q){cols.push(c.trim());c="";}else c+=ch;}return cols; };
const parseTable = text => { const lines=text.trim().split("\n").filter(l=>l.trim());if(lines.length<2)return[];const sep=lines[0].includes("\t")?"\t":",";const hdrs=lines[0].split(sep).map(h=>h.trim().replace(/"/g,""));return lines.slice(1).map(line=>{const cols=sep==="\t"?line.split("\t").map(c=>c.trim()):csvSplit(line);const o={};hdrs.forEach((h,i)=>{o[h]=cols[i]||"";});return o;}); };

// ── DEFAULT MAPPING RULES ─────────────────────────────────────────────────────
const DEFAULT_BANK_RULES = [
  {id:"b1",keywords:"5/3 bank,stripe,cherry,repeatmd",line:"Injectable/Skin Income"},
  {id:"b2",keywords:"trueaesthetics,true aesthetics",line:"TrueAesthetics"},
  {id:"b3",keywords:"nitra",line:"Nitra"},
  {id:"b4",keywords:"vantiv",line:"Bank & Merchant Fees"},
  {id:"b5",keywords:"gusto,dd sweep",line:"Wages"},
  {id:"b6",keywords:"tax sweep",line:"Payroll Taxes"},
  {id:"b7",keywords:"vestwell",line:"Benefits"},
  {id:"b8",keywords:"venmo,trilogy medical",line:"Misc- TBD"},
  {id:"b9",keywords:"holland commercial,the den",line:"Rent"},
  {id:"b10",keywords:"jpmorgan chase",line:"Car Loan"},
  {id:"b11",keywords:"delta dental,insur prem",line:"Insurance"},
  {id:"b12",keywords:"xcel energy",line:"Utilities"},
  {id:"b13",keywords:"co dept revenue taxpayment",line:"Taxes, Licenses & Fees"},
  {id:"b14",keywords:"easy track,easytrack",line:"Payroll Fees"},
  {id:"b15",keywords:"aspire galderma,galderma",line:"Rebates"},
  {id:"b16",keywords:"advs ed serv,collegeinvest",line:"Personal/Distributions"},
  {id:"b17",keywords:"capital one",line:"Capital One"},
  {id:"b18",keywords:"amex,american express",line:"Amex"},
];
const DEFAULT_INV_RULES = [
  {id:"i1",keywords:"tox membership,true tox,tox 1 member",line:"Tox Membership"},
  {id:"i2",keywords:"beauty hive membership,bh membership",line:"BH Membership"},
  {id:"i3",keywords:"weight loss membership,weightloss membership,restorative dose",line:"Weightloss Membership"},
  {id:"i4",keywords:"gift card,card decline fee,deposit redeemed",line:"Gift Cards/Deposit Redeemed"},
];

const IN_MAP = {"Injectable/Skin Income":"inj","Rebates":"reb","Gift Cards/Deposits Purchased":"gft","Gift Cards/Deposit Redeemed":"grd","Taylor Campbell":"tc","Emily Kurtz":"ek","Leah Barr":"lb","Rico Alvarado":"ra","BH Membership":"bhm","Tox Membership":"tox","Weightloss Membership":"wlm","Skincare Products":"skn"};
const OUT_MAP = {"TrueAesthetics":"ta","Nitra":"nit","Capital One":"cap","Amex":"amx","Wages":"wag","Payroll Taxes":"pt","Payroll Fees":"pf","Benefits":"ben","Rent":"rnt","Utilities":"utl","Bank & Merchant Fees":"bkf","Insurance":"ins","Car Loan":"cl","Taxes, Licenses & Fees":"tax","Misc- TBD":"misc","Personal/Distributions":"dist","Dues & Subscriptions":"dues","Repairs & Maintenance":"rm","Image First":"imgf","OKC Location":"okc"};
const STAFF_MAP = {"taylor campbell":"Taylor Campbell","emily kurtz":"Emily Kurtz","leah barr":"Leah Barr","rico alvarado":"Rico Alvarado"};
const STAFF_RATES = {"Emily Kurtz":3000,"Taylor Campbell":4800,"Leah Barr":1000,"Rico Alvarado":450};
const TOX_PRICE=175, WL_PRICE=455;
const ALL_LINES = [...Object.keys(IN_MAP),...Object.keys(OUT_MAP)];

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
    {id:"dues",label:"Dues & Subscriptions",   v:pad([],NW)},
    {id:"rm",  label:"Repairs & Maintenance",  v:pad([965,null,965,null,965,null,null,965,null,null],NW)},
    {id:"imgf",label:"Image First",            v:pad([130,null,null,null,null,null,90,null,null,null],NW)},
    {id:"okc", label:"OKC Location",           v:pad([2300,null,null,null,null,null,null,null,null,null],NW)},
  ]},
];
const BB_INIT = pad([49856,33794,21705,20611,36427,21939,17545,17547,50978,27790],NW);

// ── FORECAST PARSERS ──────────────────────────────────────────────────────────
const parseFcExpenses = (text,weeks) => {
  const totals={};Object.values(OUT_MAP).forEach(id=>{totals[id]=pad([],NW);});
  text.split("\n").forEach((line,li)=>{if(li===0||!line.trim())return;const cols=csvSplit(line);if(cols.length<5)return;const mapRaw=(cols[4]||"").replace(/^"|"$/g,"").trim();const amt=Math.abs(parseFloat((cols[3]||"").replace(/[^0-9.]/g,""))||0);const dateStr=(cols[2]||"").trim();if(!mapRaw||!amt)return;const outKey=Object.keys(OUT_MAP).find(k=>k.trim().toLowerCase()===mapRaw.toLowerCase());if(!outKey)return;const wi=dateToWeekIdx(dateStr,weeks);if(wi>=0&&wi<NW&&weeks[wi].fc)totals[OUT_MAP[outKey]][wi]=(totals[OUT_MAP[outKey]][wi]||0)+amt;});
  Object.keys(totals).forEach(id=>{totals[id]=totals[id].map(v=>v||null);});return totals;
};
const parseFcInflows = (text,weeks) => {
  const lines=text.trim().split("\n").filter(l=>l.trim());if(lines.length<2)return{};
  const sep=lines[0].includes("\t")?"\t":",";const rawHdrs=lines[0].split(sep).map(h=>h.trim().replace(/"/g,""));const wkHdrs=rawHdrs.slice(2);
  const hdrWi=wkHdrs.map(h=>{const d=parseDate(h);if(!d)return -1;return dateToWeekIdx(h,weeks);});
  const res={};Object.values(IN_MAP).forEach(id=>{res[id]=pad([],NW);});
  const toxAcc=pad([],NW),wlAcc=pad([],NW);let sec=null;
  lines.slice(1).forEach(line=>{const cols=sep==="\t"?line.split("\t").map(c=>c.trim()):csvSplit(line);const label=(cols[0]||"").trim(),rpd=parseAmt(cols[1]),lbl=label.toLowerCase();if(lbl.includes("tox membership")&&!rpd){sec="tox";return;}if((lbl.includes("weightloss")||lbl.includes("weight loss"))&&!rpd){sec="wl";return;}if(!lbl)return;const sk=Object.keys(STAFF_MAP).find(k=>lbl===k);if(sk&&rpd){const rate=STAFF_RATES[STAFF_MAP[sk]]||rpd,id=IN_MAP[STAFF_MAP[sk]];wkHdrs.forEach((_,hi)=>{const wi=hdrWi[hi];if(wi<0||!weeks[wi].fc)return;const days=parseAmt(cols[hi+2]);if(!days)return;res[id][wi]=(res[id][wi]||0)+Math.round(days*rate);res["inj"][wi]=(res["inj"][wi]||0)+Math.round(days*rate);});sec=null;return;}if(sec==="tox"&&(lbl==="new"||lbl.includes("re-occur"))){wkHdrs.forEach((_,hi)=>{const wi=hdrWi[hi];if(wi<0||!weeks[wi].fc)return;const q=parseAmt(cols[hi+2]);if(q)toxAcc[wi]=(toxAcc[wi]||0)+q;});return;}if(sec==="wl"&&(lbl==="new"||lbl.includes("re-occur"))){wkHdrs.forEach((_,hi)=>{const wi=hdrWi[hi];if(wi<0||!weeks[wi].fc)return;const q=parseAmt(cols[hi+2]);if(q)wlAcc[wi]=(wlAcc[wi]||0)+q;});return;}});
  toxAcc.forEach((q,i)=>{if(q)res["tox"][i]=Math.round(q*TOX_PRICE);});wlAcc.forEach((q,i)=>{if(q)res["wlm"][i]=Math.round(q*WL_PRICE);});
  Object.keys(res).forEach(id=>{res[id]=res[id].map(v=>v||null);});return res;
};

// ── KPI ───────────────────────────────────────────────────────────────────────
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
  const [weeks,setWeeks]=useState(WEEKS_INIT);
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
  const [markActualModal,setMarkActualModal]=useState(false);
  const [pendingMarkWi,setPendingMarkWi]=useState(null);
  const [bankRules,setBankRules]=useState(DEFAULT_BANK_RULES);
  const [invRules,setInvRules]=useState(DEFAULT_INV_RULES);
  const [editRule,setEditRule]=useState(null); // {type,id,field,val}

  const lastActual = useMemo(()=>{const a=weeks.map((w,i)=>({...w,i})).filter(w=>!w.fc);return a.length?a[a.length-1].i:0;},[weeks]);

  // Build classifiers from editable rules
  const classifyBank = row => {
    if((row["Type"]||"").trim()) return {line:row["Type"].trim(),conf:"mapped"};
    const c=(row["Name"]||"").toLowerCase()+" "+(row["Memo/Description"]||"").toLowerCase()+" "+(row["Split"]||"").toLowerCase();
    for(const r of bankRules){const kws=r.keywords.split(",").map(k=>k.trim().toLowerCase());if(kws.some(k=>c.includes(k)))return{line:r.line,conf:"auto"};}
    return {line:null,conf:"unknown"};
  };
  const classifyInv = row => {
    if((row["Cash Inflows"]||"").trim()) return {line:row["Cash Inflows"].trim(),conf:"mapped"};
    const prod=(row["Service/Product"]||"").toLowerCase(),staff=(row["Staff"]||"").toLowerCase().trim();
    const tax=parseAmt(row["Tax"]),total=parseAmt(row["Total Due"]);
    if(!total) return {line:null,conf:"skip"};
    for(const r of invRules){const kws=r.keywords.split(",").map(k=>k.trim().toLowerCase());if(kws.some(k=>prod.includes(k)))return{line:r.line,conf:"auto"};}
    if(tax&&Math.abs(tax)>0) return {line:"Skincare Products",conf:"auto"};
    if(STAFF_MAP[staff]) return {line:STAFF_MAP[staff],conf:"auto"};
    return {line:null,conf:"unknown"};
  };

  const infTot=useMemo(()=>Array.from({length:NW},(_,i)=>inflows.reduce((s,r)=>s+(r.v[i]||0),0)),[inflows]);
  const outTot=useMemo(()=>Array.from({length:NW},(_,i)=>outflows.flatMap(s=>s.rows).reduce((s,r)=>s+(r.v[i]||0),0)),[outflows]);
  const beginBal=useMemo(()=>{const b=[...BB_INIT];for(let i=1;i<NW;i++)b[i]=b[i-1]+infTot[i-1]-outTot[i-1];return b;},[infTot,outTot]);
  const endBal=useMemo(()=>Array.from({length:NW},(_,i)=>beginBal[i]+infTot[i]-outTot[i]),[beginBal,infTot,outTot]);
  const totalPages=Math.ceil(NW/WEEKS_PER_PAGE);
  const pageWeeks=weeks.slice(page*WEEKS_PER_PAGE,(page+1)*WEEKS_PER_PAGE).map((w,j)=>({...w,i:page*WEEKS_PER_PAGE+j}));
  const chartData=weeks.map((w,i)=>({name:w.s,Inflows:infTot[i],Outflows:outTot[i],Balance:endBal[i]}));

  const applyRows = (rows,amtKey,wi) => {
    const totals={};
    rows.forEach(r=>{if(!r._line)return;const a=parseAmt(r[amtKey]);if(!a)return;totals[r._line]=(totals[r._line]||0)+a;});
    setInflows(prev=>prev.map(r=>{const k=Object.keys(IN_MAP).find(k=>IN_MAP[k]===r.id);if(!k||totals[k]===undefined)return r;return{...r,v:r.v.map((v,i)=>i===wi?(totals[k]):v)};}));
    setOutflows(prev=>prev.map(s=>({...s,rows:s.rows.map(r=>{const k=Object.keys(OUT_MAP).find(k=>OUT_MAP[k]===r.id);if(!k||totals[k]===undefined)return r;return{...r,v:r.v.map((v,i)=>i===wi?Math.abs(totals[k]):v)};})}))); 
  };

  const applyFc = (expMap,infMap) => {
    setOutflows(prev=>prev.map(s=>({...s,rows:s.rows.map(r=>{const nv=expMap[r.id];if(!nv)return r;return{...r,v:r.v.map((old,i)=>weeks[i].fc?(nv[i]??null):old)};})})));
    setInflows(prev=>prev.map(r=>{const nv=infMap[r.id];if(!nv)return r;return{...r,v:r.v.map((old,i)=>weeks[i].fc?(nv[i]??null):old)};}));
    const snap={inflows:{},outflows:{}};
    Object.keys(infMap).forEach(id=>{snap.inflows[id]=[...infMap[id]];});
    Object.keys(expMap).forEach(id=>{snap.outflows[id]=[...expMap[id]];});
    setFcSnapshot(snap);
    setSyncState(s=>({...s,fc:"done"}));
    setSyncMsg(`✓ Forecasts synced`);
  };

  const fetchSheet = async gid => {
    const res=await fetch(`${MAKE_WEBHOOK}?gid=${gid}`,{method:"GET",mode:"cors"});
    if(!res.ok)throw new Error(`HTTP ${res.status}`);
    const text=await res.text();
    if(!text||text.length<5)throw new Error("Empty response");
    return text;
  };

  const runSync = async (type) => {
    const key=type==="anb"?"anb":"inv";
    setSyncState(s=>({...s,[key]:"loading"}));setSyncMsg(null);
    try {
      const gid=type==="anb"?GID.anb:GID.invoicing;
      const text=await fetchSheet(gid);
      const rows=parseTable(text);
      const wi=lastActual+1; // next week to sync
      if(wi>=NW){setSyncMsg("No hay semanas de forecast para sincronizar.");setSyncState(s=>({...s,[key]:null}));return;}
      const ws=new Date(weeks[wi].s),we=new Date(weeks[wi].e);we.setHours(23,59,59);
      const wRows=rows.filter(r=>{const d=parseDate(r["Date"]||"");return d&&d>=ws&&d<=we;});
      const cl=wRows.map(r=>{const res=type==="anb"?classifyBank(r):classifyInv(r);return{...r,_line:res.line,_conf:res.conf};});
      const unk=cl.filter(r=>r._conf==="unknown"&&parseAmt(r[type==="anb"?"Amount":"Total Due"]));
      if(unk.length){setPending({type,rows:cl,wi,unk});setUnkModal(true);}
      else{applyRows(cl,type==="anb"?"Amount":"Total Due",wi);setPendingMarkWi(wi);setMarkActualModal(true);setSyncState(s=>({...s,[key]:"done"}));setSyncMsg(`✓ ${type==="anb"?"ANB":"Invoicing"} synced — ${weeks[wi].s}–${weeks[wi].e}`);}
    }catch(e){setSyncState(s=>({...s,[key]:"error"}));setSyncMsg(`Error: ${e.message}`);}
  };

  const syncFc = async () => {
    setSyncState(s=>({...s,fc:"loading"}));setSyncMsg(null);
    try {
      const [expText,infText]=await Promise.all([fetchSheet(GID.forecast),fetchSheet(GID.assumptions)]);
      applyFc(parseFcExpenses(expText,weeks),parseFcInflows(infText,weeks));
    }catch(e){setSyncState(s=>({...s,fc:"error"}));setSyncMsg(`Forecast error: ${e.message}`);}
  };

  const syncFcPaste = () => {
    setSyncState(s=>({...s,fc:"loading"}));setSyncMsg(null);
    applyFc(parseFcExpenses(pasteExp,weeks),parseFcInflows(pasteInf,weeks));
    setPasteModal(false);
  };

  const confirmUnk = () => {
    const final=pending.rows.map((r,i)=>({...r,_line:manMap[i]||r._line}));
    applyRows(final,pending.type==="anb"?"Amount":"Total Due",pending.wi);
    setSyncState(s=>({...s,[pending.type]:"done"}));
    setSyncMsg(`✓ ${pending.type==="anb"?"ANB":"Invoicing"} synced with manual mappings`);
    setUnkModal(false);setManMap({});
    setPendingMarkWi(pending.wi);setMarkActualModal(true);
  };

  const confirmMarkActual = () => {
    if(pendingMarkWi!==null){setWeeks(prev=>prev.map((w,i)=>i===pendingMarkWi?{...w,fc:false}:w));}
    setMarkActualModal(false);setPendingMarkWi(null);
  };

  const commitEdit=(sec,rowId,wi,val)=>{
    const num=val===""?null:parseFloat(val)||null;
    if(sec==="in")setInflows(prev=>prev.map(r=>r.id===rowId?{...r,v:r.v.map((v,i)=>i===wi?num:v)}:r));
    else setOutflows(prev=>prev.map(s=>({...s,rows:s.rows.map(r=>r.id===rowId?{...r,v:r.v.map((v,i)=>i===wi?num:v)}:r)})));
    setEditCell(null);
  };

  const handleExport=()=>{
    const hdr=["Line",...weeks.map(w=>`${w.s}-${w.e}`)].join(",");
    const csv=[hdr,["Beginning Balance",...beginBal].join(","),...inflows.map(r=>[r.label,...r.v.map(v=>v??0)].join(",")),...outflows.flatMap(s=>s.rows.map(r=>[r.label,...r.v.map(v=>v??0)].join(","))),["Ending Balance",...endBal].join(",")].join("\n");
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));a.download="RI_cashflow.csv";a.click();
  };

  // Rule editor helpers
  const updateRule=(type,id,field,val)=>{
    if(type==="bank") setBankRules(prev=>prev.map(r=>r.id===id?{...r,[field]:val}:r));
    else setInvRules(prev=>prev.map(r=>r.id===id?{...r,[field]:val}:r));
  };
  const addRule=(type)=>{
    const newId=`${type==="bank"?"b":"i"}${Date.now()}`;
    if(type==="bank") setBankRules(prev=>[...prev,{id:newId,keywords:"",line:""}]);
    else setInvRules(prev=>[...prev,{id:newId,keywords:"",line:""}]);
  };
  const deleteRule=(type,id)=>{
    if(type==="bank") setBankRules(prev=>prev.filter(r=>r.id!==id));
    else setInvRules(prev=>prev.filter(r=>r.id!==id));
  };

  // ── STYLES ──
  const th=fc=>({padding:"4px 8px",textAlign:"right",minWidth:90,fontSize:11,fontWeight:600,background:fc?"#fef9ec":"#eff6ff",borderLeft:"1px solid #e2e8f0",color:fc?"#b45309":"#1d4ed8",whiteSpace:"nowrap"});
  const td=fc=>({padding:"3px 8px",textAlign:"right",fontSize:12,background:fc?"#fffdf5":"#f8fbff",borderLeft:"1px solid #e2e8f0",whiteSpace:"nowrap",cursor:"pointer",color:"#1e293b"});
  const lbl=(sub,bold,color,bg)=>({padding:`3px 8px 3px ${sub?"20px":"8px"}`,fontSize:12,fontWeight:bold?700:400,color:color||(sub?"#64748b":"#1e293b"),minWidth:200,position:"sticky",left:0,background:bg||"#fff",zIndex:2,borderRight:"2px solid #e2e8f0",whiteSpace:"nowrap"});
  const sHdr=(bg,col)=>({padding:"5px 8px",fontSize:11,fontWeight:700,color:col||"#1d4ed8",textTransform:"uppercase",letterSpacing:1,position:"sticky",left:0,background:bg||"#eff6ff",zIndex:2,borderRight:"2px solid #e2e8f0"});
  const subS={padding:"4px 8px 4px 12px",fontSize:10,fontWeight:700,color:"#64748b",position:"sticky",left:0,background:"#f8fafc",zIndex:2,borderRight:"2px solid #e2e8f0",cursor:"pointer"};
  const btnS=(active,col)=>({padding:"6px 14px",borderRadius:7,border:"1px solid #e2e8f0",cursor:"pointer",fontSize:12,fontWeight:600,background:active?(col||"#2563eb"):"#f1f5f9",color:active?"#fff":"#475569"});
  const syncBtn=(label,key,onClick,color)=>{const st=syncState[key];return React.createElement("button",{onClick,disabled:st==="loading",style:{display:"flex",alignItems:"center",gap:5,padding:"6px 14px",borderRadius:7,border:"1px solid #e2e8f0",cursor:st==="loading"?"default":"pointer",fontSize:12,fontWeight:600,background:st==="done"?"#dcfce7":st==="loading"?"#f1f5f9":"#fff",color:st==="done"?S.green:st==="loading"?S.muted:(color||S.blue),transition:"all .2s"}},st==="loading"?"⟳ Syncing...":st==="done"?`✓ ${label}`:`⟳ ${label}`);};

  const renderCell=(section,row,wi)=>{
    const v=row.v[wi],isEd=editCell&&editCell.sec===section&&editCell.id===row.id&&editCell.wi===wi;
    const fc=weeks[wi]&&weeks[wi].fc,color=v<0?"#b91c1c":v>0?"#1e293b":"#94a3b8";
    if(isEd)return React.createElement("td",{key:wi,style:{...td(fc),color}},React.createElement("input",{autoFocus:true,value:editVal,onChange:e=>setEditVal(e.target.value),onBlur:()=>commitEdit(section,row.id,wi,editVal),onKeyDown:e=>{if(e.key==="Enter")commitEdit(section,row.id,wi,editVal);},style:{width:68,background:"#eff6ff",border:"1px solid #2563eb",borderRadius:4,color:"#1e293b",fontSize:11,padding:"1px 4px",textAlign:"right"}}));
    return React.createElement("td",{key:wi,style:{...td(fc),color},onClick:()=>{setEditCell({sec:section,id:row.id,wi});setEditVal(v===null?"":String(v));}},fmt(v));
  };

  const el=(t,p,...c)=>React.createElement(t,p,...c.flat().filter(v=>v!==false&&v!=null&&v!==undefined));
  const modal=(content)=>el("div",{style:{position:"fixed",inset:0,background:"rgba(0,0,0,0.28)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}},
    el("div",{style:{background:"#fff",borderRadius:14,padding:24,width:600,maxWidth:"96vw",maxHeight:"88vh",overflowY:"auto",boxShadow:"0 8px 32px rgba(0,0,0,0.14)"}},content)
  );

  const varWeeks=weeks.map((w,i)=>({...w,i})).filter(w=>!w.fc).slice(-6);
  const varRows=useMemo(()=>{
    if(!fcSnapshot)return[];
    const rows=[];
    inflows.forEach(row=>{const fc=fcSnapshot.inflows[row.id];if(!fc)return;if(!varWeeks.some(w=>row.v[w.i]||fc[w.i]))return;rows.push({label:row.label,sub:row.label.startsWith("  "),type:"in",act:row.v,fc});});
    outflows.flatMap(s=>s.rows).forEach(row=>{const fc=fcSnapshot.outflows[row.id];if(!fc)return;if(!varWeeks.some(w=>row.v[w.i]||fc[w.i]))return;rows.push({label:row.label,sub:false,type:"out",act:row.v,fc});});
    return rows;
  },[fcSnapshot,inflows,outflows,varWeeks]);

  // ── MAPPING RULES TAB ──
  const renderRulesTab=()=>{
    const ruleTable=(type,rules,title,color,bg)=>el("div",{style:{background:"#fff",borderRadius:12,border:"1px solid #e2e8f0",padding:16,marginBottom:16}},
      el("div",{style:{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}},
        el("div",{style:{fontWeight:700,fontSize:14,color}},[title]),
        el("button",{onClick:()=>addRule(type),style:{padding:"5px 12px",borderRadius:7,border:`1px solid ${color}`,background:bg,color,fontSize:12,fontWeight:600,cursor:"pointer"}},"+ Add Rule")
      ),
      el("div",{style:{overflowX:"auto"}},
        el("table",{style:{borderCollapse:"collapse",width:"100%",fontSize:12}},
          el("thead",null,el("tr",{style:{background:"#f8fafc"}},
            el("th",{style:{padding:"6px 10px",textAlign:"left",borderBottom:"1px solid #e2e8f0",fontWeight:600,color:S.muted,width:"45%"}},"Keywords (comma-separated)"),
            el("th",{style:{padding:"6px 10px",textAlign:"left",borderBottom:"1px solid #e2e8f0",fontWeight:600,color:S.muted,width:"45%"}},"→ Maps to Line"),
            el("th",{style:{padding:"6px 10px",borderBottom:"1px solid #e2e8f0",width:"10%"}})
          )),
          el("tbody",null,...rules.map(r=>
            el("tr",{key:r.id,style:{borderBottom:"1px solid #f1f5f9"}},
              el("td",{style:{padding:"4px 8px"}},
                el("input",{value:r.keywords,onChange:e=>updateRule(type,r.id,"keywords",e.target.value),style:{width:"100%",padding:"4px 8px",borderRadius:6,border:"1px solid #e2e8f0",fontSize:11,boxSizing:"border-box",fontFamily:"monospace"}})
              ),
              el("td",{style:{padding:"4px 8px"}},
                el("select",{value:r.line,onChange:e=>updateRule(type,r.id,"line",e.target.value),style:{width:"100%",padding:"4px 8px",borderRadius:6,border:"1px solid #e2e8f0",fontSize:12,background:"#fff"}},
                  el("option",{value:""},"— Select line —"),
                  ...ALL_LINES.map(l=>el("option",{key:l,value:l},l))
                )
              ),
              el("td",{style:{padding:"4px 8px",textAlign:"center"}},
                el("button",{onClick:()=>deleteRule(type,r.id),style:{background:"none",border:"none",cursor:"pointer",color:"#ef4444",fontSize:16,lineHeight:1}},"×")
              )
            )
          ))
        )
      )
    );
    return el("div",{style:{padding:16}},
      el("div",{style:{fontSize:13,color:S.muted,marginBottom:16}},"Estos mapeos se usan cuando sincronizás ANB o Invoicing. Las keywords son case-insensitive y se buscan dentro del nombre/descripción de cada fila."),
      ruleTable("bank",bankRules,"🏦 ANB — Expense Mapping Rules",S.red,"#fef2f2"),
      ruleTable("inv",invRules,"🧾 Invoicing — Inflow Mapping Rules",S.green,"#f0fdf4")
    );
  };

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
        el("button",{onClick:()=>setTab("rules"),style:btnS(tab==="rules","#0891b2")},"🗂 Mapping Rules"),
        el("button",{onClick:handleExport,style:{...btnS(false),...{color:S.green,borderColor:"#bbf7d0",background:S.gBg}}},"⬇ Export CSV")
      )
    ),

    // Sync bar
    el("div",{style:{padding:"8px 16px",background:"#fff",borderBottom:"1px solid #e2e8f0",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}},
      syncBtn("Sync ANB","anb",()=>runSync("anb"),S.red),
      syncBtn("Sync Invoicing","inv",()=>runSync("inv"),S.green),
      syncBtn("Sync Forecasts","fc",syncFc,S.purple),
      el("button",{onClick:()=>setPasteModal(true),style:{...btnS(false),...{color:S.purple,borderColor:"#ddd6fe",fontSize:11}}},"📋 Paste Forecasts"),
      el("div",{style:{fontSize:11,color:S.muted,marginLeft:4}},`Próxima semana a sincronizar: ${weeks[lastActual+1]?`${weeks[lastActual+1].s}–${weeks[lastActual+1].e}`:"—"}`),
      syncMsg&&el("div",{style:{fontSize:11,color:syncMsg.includes("error")||syncMsg.includes("Error")?S.red:S.green,marginLeft:"auto"}},syncMsg)
    ),

    // KPIs
    el("div",{style:{display:"flex",gap:10,padding:"10px 16px",flexWrap:"wrap",background:"#f8fafc",borderBottom:"1px solid #e2e8f0"}},
      el(KPI,{label:`Balance — ${weeks[lastActual].s}`,value:fmt(endBal[lastActual]),color:"#1d4ed8",bg:S.bBg}),
      el(KPI,{label:"Inflows (último actual)",value:fmt(infTot[lastActual]),color:S.green,bg:S.gBg}),
      el(KPI,{label:"Outflows (último actual)",value:fmt(outTot[lastActual]),color:S.red,bg:S.rBg}),
      el(KPI,{label:"Balance proyectado final",value:fmt(endBal[NW-1]),sub:`(${weeks[NW-1].e})`,color:S.yellow,bg:S.yBg})
    ),

    // ── MARK AS ACTUAL MODAL ──
    markActualModal&&modal(
      el("div",null,
        el("div",{style:{fontSize:24,marginBottom:8}},"✅"),
        el("div",{style:{fontWeight:700,fontSize:15,marginBottom:6}}),[`Sync completado — ${pendingMarkWi!==null?`${weeks[pendingMarkWi].s}–${weeks[pendingMarkWi].e}`:""}`],
        el("div",{style:{fontSize:13,color:S.muted,marginBottom:20}},"¿Querés marcar esta semana como Actual en lugar de Forecast?"),
        el("div",{style:{display:"flex",gap:8,justifyContent:"flex-end"}},
          el("button",{onClick:()=>{setMarkActualModal(false);setPendingMarkWi(null);},style:{padding:"8px 16px",borderRadius:8,border:"1px solid #e2e8f0",background:"#f1f5f9",cursor:"pointer",fontSize:13,color:"#475569"}},"Mantener como Forecast"),
          el("button",{onClick:confirmMarkActual,style:{padding:"8px 20px",borderRadius:8,border:"none",background:S.green,cursor:"pointer",fontSize:13,fontWeight:600,color:"#fff"}},"✓ Marcar como Actual")
        )
      )
    ),

    // Paste Modal
    pasteModal&&modal(
      el("div",null,
        el("div",{style:{fontWeight:700,fontSize:15,marginBottom:4}},"📋 Paste Forecast Data"),
        el("div",{style:{fontSize:12,color:S.muted,marginBottom:14}},"Abrí cada tab en Google Sheets → Ctrl+A → Ctrl+C → pegá abajo."),
        el("div",{style:{marginBottom:12}},
          el("div",{style:{fontSize:12,fontWeight:700,color:S.red,marginBottom:4}},"📉 Expenses tab"),
          el("textarea",{value:pasteExp,onChange:e=>setPasteExp(e.target.value),rows:6,style:{width:"100%",padding:8,borderRadius:8,border:"1px solid #fecaca",fontSize:11,boxSizing:"border-box",fontFamily:"monospace",resize:"vertical",outline:"none"}})
        ),
        el("div",{style:{marginBottom:16}},
          el("div",{style:{fontSize:12,fontWeight:700,color:S.purple,marginBottom:4}},"📈 Assumptions tab"),
          el("textarea",{value:pasteInf,onChange:e=>setPasteInf(e.target.value),rows:6,style:{width:"100%",padding:8,borderRadius:8,border:"1px solid #ddd6fe",fontSize:11,boxSizing:"border-box",fontFamily:"monospace",resize:"vertical",outline:"none"}})
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
        el("div",{style:{fontWeight:700,fontSize:15,marginBottom:4,color:"#b45309"}},`⚠ ${pending.unk.length} filas necesitan mapeo manual`),
        el("div",{style:{fontSize:12,color:S.muted,marginBottom:14}},"Asigná una línea a cada fila y hacé clic en Aplicar."),
        el("div",{style:{overflowX:"auto"}},
          el("table",{style:{borderCollapse:"collapse",width:"100%",fontSize:12}},
            el("thead",null,el("tr",{style:{background:"#fffbeb"}},
              ...["Fecha","Nombre","Descripción","Monto","→ Línea"].map(h=>el("th",{key:h,style:{padding:"6px 10px",textAlign:"left",borderBottom:"1px solid #e2e8f0",fontWeight:600,color:"#78350f"}},h))
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
                    el("option",{value:""},"— Ignorar —"),
                    ...ALL_LINES.map(l=>el("option",{key:l,value:l},l))
                  )
                )
              );
            }))
          )
        ),
        el("div",{style:{display:"flex",gap:8,justifyContent:"flex-end",marginTop:16}},
          el("button",{onClick:()=>{setUnkModal(false);setManMap({});setSyncState(s=>({...s,[pending.type]:null}));},style:{padding:"8px 16px",borderRadius:8,border:"1px solid #e2e8f0",background:"#f1f5f9",cursor:"pointer",fontSize:13,color:"#475569"}},"Cancelar"),
          el("button",{onClick:confirmUnk,style:{padding:"8px 20px",borderRadius:8,border:"none",background:S.blue,cursor:"pointer",fontSize:13,fontWeight:600,color:"#fff"}},"Aplicar Mapeos")
        )
      )
    ),

    // ── RULES TAB ──
    tab==="rules"&&renderRulesTab(),

    // ── VARIANCE TAB ──
    tab==="variance"&&el("div",{style:{padding:16}},
      !fcSnapshot
        ?el("div",{style:{background:"#fff",borderRadius:12,border:"1px solid #e2e8f0",padding:32,textAlign:"center",color:S.muted}},
            el("div",{style:{fontSize:32,marginBottom:8}},"🔮"),
            el("div",{style:{fontWeight:600,marginBottom:4}},"No hay datos de forecast"),
            el("div",{style:{fontSize:12}},"Hacé clic en Sync Forecasts o Paste Forecasts para cargar.")
          )
        :el("div",null,
            el("div",{style:{fontWeight:700,fontSize:15,marginBottom:12}},"Forecast vs Actual — Últimas semanas actuales"),
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
                      const fc=row.fc[w.i]||null,ac=row.act[w.i]||null;
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
        el("div",{style:{fontWeight:600,marginBottom:10,fontSize:13}},"Inflows vs Outflows por semana"),
        el(ResponsiveContainer,{width:"100%",height:180},
          el(BarChart,{data:chartData},el(CartesianGrid,{strokeDasharray:"3 3",stroke:"#e2e8f0"}),el(XAxis,{dataKey:"name",tick:{fontSize:9,fill:S.muted},interval:2}),el(YAxis,{tickFormatter:v=>`$${(v/1000).toFixed(0)}k`,tick:{fontSize:10,fill:S.muted}}),el(Tooltip,{formatter:v=>`$${Number(v).toLocaleString("en-US")}`,contentStyle:{background:"#fff",border:"1px solid #e2e8f0",borderRadius:8,fontSize:11}}),el(Legend,{wrapperStyle:{fontSize:11}}),el(Bar,{dataKey:"Inflows",fill:S.green,radius:[3,3,0,0]}),el(Bar,{dataKey:"Outflows",fill:S.red,radius:[3,3,0,0]}))
        )
      )
    ),

    // ── TABLE TAB ──
    tab==="table"&&el("div",null,
      el("div",{style:{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 16px",background:"#fff",borderBottom:"1px solid #e2e8f0",flexWrap:"wrap",gap:6}},
        el("span",{style:{fontSize:11,color:S.muted}},"Click en celda para editar · Amber = Forecast"),
        el("div",{style:{display:"flex",gap:4,alignItems:"center"}},
          el("button",{onClick:()=>setPage(p=>Math.max(0,p-1)),disabled:page===0,style:{padding:"4px 12px",borderRadius:6,border:"1px solid #e2e8f0",cursor:page===0?"default":"pointer",fontSize:13,background:"#f1f5f9",color:page===0?"#cbd5e1":"#1e293b"}},"‹"),
          el("span",{style:{fontSize:11,color:S.muted,lineHeight:"28px"}},`Página ${page+1}/${totalPages}`),
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
              if(isC)return[hdr];
              return[hdr,...sec.rows.map(row=>el("tr",{key:row.id,style:{borderBottom:"1px solid #f1f5f9"}},el("td",{style:lbl(false,false)},row.label),...pageWeeks.map(w=>renderCell("out",row,w.i))))];
            }),
            el("tr",{key:"ot",style:{background:S.rBg,borderTop:"2px solid #e2e8f0"}},el("td",{style:lbl(false,true,S.red,S.rBg)},"Total Cash Outflows"),...pageWeeks.map(w=>el("td",{key:w.i,style:{...td(w.fc),fontWeight:700,color:S.red,background:S.rBg}},fmt(outTot[w.i])))),
            el("tr",{key:"chg",style:{background:"#f8fafc",borderTop:"1px solid #e2e8f0"}},el("td",{style:lbl(false,true,S.purple,"#f8fafc")},"Change in Cash"),...pageWeeks.map(w=>{const v=infTot[w.i]-outTot[w.i];return el("td",{key:w.i,style:{...td(w.fc),fontWeight:700,color:v>=0?S.green:S.red,background:"#f8fafc"}},fmt(v));}))
          )
        )
      )
    )
  );
}
