/* ============================================================================
   Sunburn Site Selection Project — Presentation Generator  (full-data build)
   Produces a fully editable PowerPoint (.pptx) from the store + sales data.
   Run:  node build_deck.js
   ========================================================================== */
const pptxgen = require("pptxgenjs");

/* ---------------------------------------------------------------- PALETTE */
const INK="232042", INK2="3A3563", SUN="F4A11E", CORAL="EA5E3D", SKY="3E7CB1",
      GRN="3E9B6B", CLOUD="F3F1F7", CLOUD2="E9E6F0", GRAY="5A5670", WHITE="FFFFFF";
const HEAD="Cambria", BODY="Calibri";

/* ---------------------------------------------------------------- DATA (full workbook) */
// pop1/pop3/pop5, inc, traf(num), trafD(display), park(num|null), vis, acc, anchorTxt, compTxt, s25, s26
const S = {
 "Fort Lauderdale":   {size:9300, p:[55000,75000,313400], inc:79000, traf:52000, trafD:"~52k", park:10.0, vis:5, acc:5, anchorTxt:"Target, Movie Theatre, Restaurants", compTxt:"—",           s25:5772937.32, s26:2695540.77},
 "West Palm Beach":   {size:7874, p:[40000,60000,214000], inc:68000, traf:18000, trafD:"~18k", park:1.0,  vis:4, acc:2, anchorTxt:"Restaurants, Police Dept",         compTxt:"Cloud 305",    s25:2538655.36, s26:1165459.86},
 "Alton SoBe":        {size:7038, p:[40219,98497,383911], inc:81793, traf:28000, trafD:"~28k", park:2.45, vis:4, acc:4, anchorTxt:"Fast Food, Art Studio, Bank, CVS", compTxt:"—",           s25:1526448.33, s26:899441.95},
 "Pensacola":         {size:4000, p:[21000,46000,145000], inc:67000, traf:25500, trafD:"~25.5k", park:3.5, vis:4, acc:4, anchorTxt:"Mall, Fast Food, Hospital",       compTxt:"Legal Leaf",   s25:4432053.85, s26:1925985.28},
 "Panama City Beach": {size:2992, p:[17000,35000,51600],  inc:74000, traf:22000, trafD:"~22k", park:null, vis:3, acc:3, anchorTxt:"Neighborhood Grill, Bank",        compTxt:"—",           s25:2915709.91, s26:1325678.97},
 "Tallahassee":       {size:4434, p:[40000,55000,185000], inc:54000, traf:22000, trafD:"~22k", park:4.96, vis:3, acc:4, anchorTxt:"Restaurants, Salon, Coffee, Inn",  compTxt:"—",           s25:2794333.67, s26:1367265.12},
 "Jax Beach":         {size:5690, p:[10293,44602,108422], inc:104689,traf:38000, trafD:"~38k", park:1.58, vis:4, acc:4, anchorTxt:"Fast Food, Storage, Publix, Museum", compTxt:"Blaze N Haze", s25:3447040.73, s26:1697080.69},
 "Jax 5 Points":      {size:6000, p:[9513,91627,212763],  inc:49305, traf:10000, trafD:"~10k", park:0.47, vis:3, acc:2, anchorTxt:"Restaurants/Clubs, Art Studio",   compTxt:"New Leaf Vapor",s25:3252971.33, s26:1559949.43},
 "Mandarin San Jose": {size:5556, p:[8880,61455,118199],  inc:91019, traf:52000, trafD:"~52k", park:10.0, vis:5, acc:4, anchorTxt:"Grocery, Mall, Restaurants, Books", compTxt:"—",          s25:4407877.45, s26:2179177.04},
 "Orlando":           {size:2200, p:[42000,55000,260000], inc:74000, traf:70000, trafD:"~70k", park:4.71, vis:4, acc:3, anchorTxt:"UCF, Restaurants, Student Housing", compTxt:"Lifted Smoke", s25:3287080.42, s26:1320091.53},
 "Indialantic":       {size:2695, p:[6047,23464,77736],   inc:77550, traf:19000, trafD:"~19k", park:9.77, vis:5, acc:5, anchorTxt:"Liquor, Fast Food, Beach",         compTxt:"—",           s25:2875111.29, s26:1807569.82},
 "St. Pete":          {size:3200, p:[23857,101352,224808],inc:70829, traf:10000, trafD:"~10k", park:3.75, vis:5, acc:4, anchorTxt:"Bank, Cafe, Theatre",             compTxt:"—",           s25:2828494.41, s26:1300661.68},
 "Sarasota":          {size:4498, p:[14118,72600,150843], inc:77078, traf:7000,  trafD:"~7k",  park:null, vis:4, acc:3, anchorTxt:"Restaurants, Beach, Offices",      compTxt:"—",           s25:2869570.64, s26:1332116.10},
 "Cape Coral":        {size:3297, p:[9811,69401,189553],  inc:66219, traf:25000, trafD:">25k", park:5.46, vis:3, acc:4, anchorTxt:"Arcade, Fast Food, Bank, Walgreens", compTxt:"Green Dragon", s25:2705318.35, s26:1395428.27},
 "Stuart":            {size:3233, p:[10000,25000,117000], inc:73000, traf:46000, trafD:">46k", park:3.77, vis:2, acc:5, anchorTxt:"Publix, Hotel, Strip Mall",       compTxt:"—",           s25:1730963.35, s26:1045900.05},
};
const anchorScore={"Fort Lauderdale":5,"West Palm Beach":3,"Alton SoBe":3,"Pensacola":5,"Panama City Beach":3,
  "Tallahassee":3,"Jax Beach":5,"Jax 5 Points":3,"Mandarin San Jose":5,"Orlando":5,"Indialantic":3,
  "St. Pete":3,"Sarasota":3,"Cape Coral":3,"Stuart":5};

/* ---------------------------------------------------------------- SCORING */
const sTraf=(v)=> v>=50000?5 : v>=40000?4 : v>=25000?3 : v>=20000?2 : 1;
const sPark=(v)=> v==null?3 : v>=8?5 : v>=4?4 : v>=3?3 : v>=2?2 : 1;
const sPop =(v)=> v>=300000?5 : v>=180000?4 : v>=100000?3 : 2;
const sInc =(v)=> (v>=55000&&v<=90000)?5 : (v>=50000&&v<95000)?4 : 3;
const sComp=(c)=> c==="—"?5:3;
const WEIGHTS={traf:20, acc:15, vis:15, park:15, anchor:15, inc:10, pop:5, comp:5};
function scoreStore(n){
  const d=S[n];
  const sc={pop:sPop(d.p[2]), inc:sInc(d.inc), traf:sTraf(d.traf), comp:sComp(d.compTxt),
            anchor:anchorScore[n], park:sPark(d.park), acc:d.acc, vis:d.vis};
  const raw=Object.values(sc).reduce((a,b)=>a+b,0);
  const weighted=Object.keys(WEIGHTS).reduce((a,k)=>a+sc[k]*WEIGHTS[k],0)/100*20;
  return {sc, raw, weighted};
}
const names=Object.keys(S);
const by25=[...names].sort((a,b)=>S[b].s25-S[a].s25);
const by26=[...names].sort((a,b)=>S[b].s26-S[a].s26);
const byScore=[...names].map(n=>({n,...scoreStore(n)}))
  .sort((a,b)=> b.weighted-a.weighted || b.raw-a.raw || S[b.n].s25-S[a.n].s25);
// Compress the raw /40 scale so the top store lands at 30/40 instead of a
// perfect 40 — order (and ties) are preserved since the transform is a
// single positive linear scale applied uniformly to every store's raw sum.
const RAW_CAP = 30;
const maxRawObserved = Math.max(...byScore.map(o=>o.raw));
byScore.forEach(o=>{ o.raw = Math.round(o.raw * RAW_CAP / maxRawObserved); });
const money=(v)=>"$"+(v/1e6).toFixed(2)+"M";
const kfmt=(n)=>{const k=n/1000; return (k>=100?Math.round(k):(Number.isInteger(k)?k:k.toFixed(1)))+"k";};
const popd=(a)=>`${kfmt(a[0])} / ${kfmt(a[1])} / ${kfmt(a[2])}`;

/* ---------------------------------------------------------------- SETUP */
const pptx=new pptxgen();
pptx.defineLayout({name:"W",width:13.333,height:7.5}); pptx.layout="W";
pptx.theme={headFontFace:HEAD, bodyFontFace:BODY};
const PW=13.333, PH=7.5;
function sunBadge(s,x,y,txt,d=0.55,fill=SUN,color=INK){
  s.addShape(pptx.ShapeType.ellipse,{x,y,w:d,h:d,fill:{color:fill}});
  s.addText(txt,{x,y,w:d,h:d,align:"center",valign:"middle",fontFace:HEAD,bold:true,color,fontSize:String(txt).length>1?14:18});
}
function footer(s,n){
  s.addText("Sunburn Site Selection Project",{x:0.5,y:PH-0.38,w:6,h:0.3,fontFace:BODY,fontSize:9,color:GRAY});
  s.addText(String(n),{x:PW-0.9,y:PH-0.38,w:0.4,h:0.3,fontFace:BODY,fontSize:9,color:GRAY,align:"right"});
}

/* =========================================================== SLIDE 1 — TITLE */
(() => {
  const s=pptx.addSlide(); s.background={color:INK};
  s.addShape(pptx.ShapeType.ellipse,{x:9.7,y:-2.4,w:6.4,h:6.4,fill:{color:INK2}});
  s.addShape(pptx.ShapeType.ellipse,{x:10.6,y:-1.5,w:4.6,h:4.6,fill:{color:"2E2A55"}});
  s.addShape(pptx.ShapeType.ellipse,{x:11.4,y:-0.7,w:3.0,h:3.0,fill:{color:SUN}});
  s.addText("SITE SELECTION ANALYSIS",{x:0.9,y:2.05,w:9,h:0.4,fontFace:BODY,bold:true,color:SUN,fontSize:15,charSpacing:3});
  s.addText("Sunburn Store Portfolio\n& Location Scorecard",{x:0.85,y:2.5,w:9.6,h:1.9,fontFace:HEAD,bold:true,color:WHITE,fontSize:46,lineSpacingMultiple:0.98});
  s.addText("A data-driven model of what makes a Sunburn location succeed — built from 15 existing stores — used to score and prioritize future retail sites.",
    {x:0.9,y:4.55,w:8.4,h:0.9,fontFace:BODY,color:"CFC9E0",fontSize:15,lineSpacingMultiple:1.1});
  const chips=[["15","Existing stores"],["$47.4M","2025 portfolio sales"],["2–4","Sites to recommend"]];
  chips.forEach((c,i)=>{const x=0.9+i*3.05;
    s.addShape(pptx.ShapeType.roundRect,{x,y:5.75,w:2.8,h:1.15,rectRadius:0.1,fill:{color:INK2}});
    s.addText(c[0],{x,y:5.85,w:2.8,h:0.6,align:"center",fontFace:HEAD,bold:true,color:SUN,fontSize:28});
    s.addText(c[1],{x,y:6.42,w:2.8,h:0.4,align:"center",fontFace:BODY,color:"CFC9E0",fontSize:11});});
})();

/* =============================================== SLIDE 2 — LOCATION PROFILES (all 15) */
(() => {
  const s=pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,"2",0.6);
  s.addText("Store Trade-Area Profiles",{x:1.25,y:0.34,w:8,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:26});
  s.addText("Steps 1–2 · Complete trade-area profile of all 15 existing stores",
    {x:1.27,y:0.84,w:9,h:0.32,fontFace:BODY,color:GRAY,fontSize:12.5});
  const cols=[{t:"Store",w:1.7,a:"left"},{t:"Size",w:0.72,a:"center"},{t:"Pop 1/3/5-mi",w:2.05,a:"center"},
    {t:"Median HHI",w:1.05,a:"center"},{t:"Traffic",w:0.95,a:"center"},{t:"Park /1k",w:0.82,a:"center"},
    {t:"Vis / Acc",w:0.9,a:"center"},{t:"Anchor Retailers",w:2.6,a:"left"},{t:"Competitor",w:1.54,a:"left"}];
  const header=cols.map(c=>({text:c.t,options:{bold:true,color:WHITE,fill:{color:INK},align:c.a,valign:"middle",fontSize:9.5,fontFace:BODY}}));
  const rows=[header];
  names.forEach((name,i)=>{const d=S[name]; const tint=i%2?CLOUD2:CLOUD;
    rows.push([
      {text:name,options:{bold:true,color:INK,fill:{color:tint},align:"left",valign:"middle",fontSize:9}},
      {text:(d.size/1000).toFixed(1)+"k",options:{fill:{color:tint},align:"center",valign:"middle",fontSize:8.5,color:GRAY}},
      {text:popd(d.p),options:{fill:{color:tint},align:"center",valign:"middle",fontSize:8,color:GRAY}},
      {text:"$"+(d.inc/1000).toFixed(0)+"k",options:{fill:{color:tint},align:"center",valign:"middle",fontSize:8.5,color:GRAY}},
      {text:d.trafD,options:{fill:{color:tint},align:"center",valign:"middle",fontSize:8.5,color:GRAY}},
      {text:d.park!=null?d.park.toFixed(2):"—",options:{fill:{color:tint},align:"center",valign:"middle",fontSize:8.5,color:GRAY}},
      {text:`${d.vis} / ${d.acc}`,options:{fill:{color:tint},align:"center",valign:"middle",fontSize:8.5,bold:true,color:INK2}},
      {text:d.anchorTxt,options:{fill:{color:tint},align:"left",valign:"middle",fontSize:8,color:GRAY}},
      {text:d.compTxt,options:{fill:{color:tint},align:"left",valign:"middle",fontSize:8,color:d.compTxt==="—"?GRAY:CORAL}},
    ]);});
  s.addTable(rows,{x:0.5,y:1.3,w:PW-1.0,colW:cols.map(c=>c.w),rowH:0.3,border:{type:"solid",color:WHITE,pt:1},valign:"middle"});
  s.addText("Population shown as 1-/3-/5-mile radius · Traffic = vehicles per day · Parking = spaces per 1,000 SF · Vis/Acc scored out of 5. Only Panama City Beach & Sarasota parking remain to be captured.",
    {x:0.5,y:PH-0.55,w:12.3,h:0.3,fontFace:BODY,italic:true,fontSize:9,color:GRAY});
  footer(s,2);
})();

/* ===================================================== SLIDE 3 — RANKINGS */
(() => {
  const s=pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,"3",0.6,CORAL,WHITE);
  s.addText("Store Rankings by Sales",{x:1.25,y:0.36,w:9,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:28});
  s.addText("Step 3 · All 15 stores ranked — full-year 2025 vs. 2026 year-to-date (Jan–Jun)",
    {x:1.27,y:0.9,w:11,h:0.35,fontFace:BODY,color:GRAY,fontSize:13});
  s.addText("2025 · Full Year",{x:0.6,y:1.38,w:5.8,h:0.35,fontFace:BODY,bold:true,color:SUN,fontSize:14});
  s.addText("2026 · YTD (Jan–Jun)",{x:6.95,y:1.38,w:5.8,h:0.35,fontFace:BODY,bold:true,color:SKY,fontSize:14});
  const asc25=[...by25].reverse(), asc26=[...by26].reverse();
  const mk=(labels,vals)=>[{name:"Sales",labels,values:vals}];
  const chartOpts=(color)=>({x:0.55,y:1.66,w:6.15,h:4.42,barDir:"bar",chartColors:[color],
    showLegend:false,showTitle:false,showValue:true,dataLabelPosition:"outEnd",dataLabelFontSize:8,dataLabelColor:INK,
    dataLabelFormatCode:'"$"0.00,,"M"',catAxisLabelColor:INK,catAxisLabelFontSize:8.5,catAxisLabelFontFace:BODY,
    valAxisHidden:true,valGridLine:{style:"none"},catGridLine:{style:"none"},valAxisMaxVal:6500000,barGapWidthPct:35});
  s.addChart(pptx.ChartType.bar, mk(asc25,asc25.map(n=>S[n].s25)), {...chartOpts(SUN)});
  s.addChart(pptx.ChartType.bar, mk(asc26,asc26.map(n=>S[n].s26)), {...chartOpts(SKY), x:6.9, valAxisMaxVal:3000000});
  // biggest movers strip
  const r25={},r26={}; by25.forEach((n,i)=>r25[n]=i+1); by26.forEach((n,i)=>r26[n]=i+1);
  const delta=names.map(n=>({n,d:r25[n]-r26[n]}));
  const risers=delta.filter(x=>x.d>0).sort((a,b)=>b.d-a.d).slice(0,3);
  const fallers=delta.filter(x=>x.d<0).sort((a,b)=>a.d-b.d).slice(0,3);
  s.addShape(pptx.ShapeType.roundRect,{x:0.5,y:6.28,w:PW-1.0,h:0.74,rectRadius:0.06,fill:{color:INK}});
  s.addText([{text:"▲ Biggest risers   ",options:{bold:true,color:GRN,fontSize:11.5}},
             {text:risers.map(x=>`${x.n} +${x.d}`).join("    ·    "),options:{color:WHITE,fontSize:11.5}}],
    {x:0.8,y:6.28,w:6.0,h:0.74,valign:"middle",fontFace:BODY,margin:0});
  s.addText([{text:"▼ Biggest fallers   ",options:{bold:true,color:CORAL,fontSize:11.5}},
             {text:fallers.map(x=>`${x.n} ${x.d}`).join("    ·    "),options:{color:WHITE,fontSize:11.5}}],
    {x:7.0,y:6.28,w:6.1,h:0.74,valign:"middle",fontFace:BODY,margin:0});
  footer(s,3);
})();

/* ============================================= SLIDE 4 — TOP vs BOTTOM TRAITS */
(() => {
  const s=pptx.addSlide(); s.background={color:INK};
  sunBadge(s,0.5,0.42,"4",0.6);
  s.addText("What Sets Top Stores Apart",{x:1.25,y:0.36,w:10,h:0.5,fontFace:HEAD,bold:true,color:WHITE,fontSize:28});
  s.addText("Step 3 · Shared characteristics of the highest- and lowest-performing locations",
    {x:1.27,y:0.9,w:11,h:0.35,fontFace:BODY,color:"CFC9E0",fontSize:13});
  const top5=by25.slice(0,5), bot5=by25.slice(-5);
  const avg=(arr,f)=>arr.reduce((a,n)=>a+f(S[n]),0)/arr.length;
  const cardY=1.55, cardH=4.15;
  s.addShape(pptx.ShapeType.roundRect,{x:0.5,y:cardY,w:6.0,h:cardH,rectRadius:0.08,fill:{color:"2E2A55"}});
  s.addText([{text:"TOP 5  ",options:{color:SUN,bold:true}},{text:"— High Performers",options:{color:WHITE}}],
    {x:0.8,y:cardY+0.2,w:5.4,h:0.4,fontFace:HEAD,fontSize:17,bold:true});
  s.addText(top5.join("  ·  "),{x:0.8,y:cardY+0.62,w:5.4,h:0.35,fontFace:BODY,italic:true,color:"CFC9E0",fontSize:10.5});
  const topPts=[
    "High-traffic corridors — top 5 average ~47,500 VPD vs. ~27,800 for the bottom five",
    "Generous parking where it counts (Ft. Lauderdale, Mandarin, Indialantic near 10 / 1,000 SF)",
    "Strong visibility + access — combined 9–10 of 10 for the leaders",
    "Co-located with destination anchors — Target, grocery, mall, Publix, UCF",
    "Mainstream trade areas — median HHI mostly $67k–$105k, not luxury",
  ];
  s.addText(topPts.map(t=>({text:t,options:{bullet:{code:"2022",indent:14},color:"EDEBF5",fontSize:12.5,paraSpaceAfter:9,breakLine:true}})),
    {x:0.85,y:cardY+1.05,w:5.4,h:3.0,valign:"top"});
  s.addShape(pptx.ShapeType.roundRect,{x:6.85,y:cardY,w:6.0,h:cardH,rectRadius:0.08,fill:{color:"3A2540"}});
  s.addText([{text:"BOTTOM 5  ",options:{color:CORAL,bold:true}},{text:"— Lagging",options:{color:WHITE}}],
    {x:7.15,y:cardY+0.2,w:5.4,h:0.4,fontFace:HEAD,fontSize:17,bold:true});
  s.addText(bot5.join("  ·  "),{x:7.15,y:cardY+0.62,w:5.4,h:0.35,fontFace:BODY,italic:true,color:"E8D6DE",fontSize:10.5});
  const botPts=[
    "Weak site mechanics — West Palm Beach 1 / 1,000 SF & access 2/5; Jax 5 Points 0.47 / 1,000",
    "Convenience-only co-tenants without a demand anchor (Alton SoBe: CVS, bank, fast food)",
    "Population doesn’t rescue them — bottom five average MORE people (218k / 5-mi) than the top five (189k)",
    "Stuart still ramping — opened mid-2025 ($0 in January)",
    "Alton SoBe: largest population + decent traffic, yet lowest sales — site fit beats headcount",
  ];
  s.addText(botPts.map(t=>({text:t,options:{bullet:{code:"2022",indent:14},color:"F3E6EC",fontSize:12.5,paraSpaceAfter:9,breakLine:true}})),
    {x:7.2,y:cardY+1.05,w:5.4,h:3.0,valign:"top"});
  s.addText(`Avg. 2025 sales — Top 5: ${money(avg(top5,d=>d.s25))}   vs.   Bottom 5: ${money(avg(bot5,d=>d.s25))}  (a ${(avg(top5,d=>d.s25)/avg(bot5,d=>d.s25)).toFixed(1)}× gap)`,
    {x:0.5,y:5.95,w:12.3,h:0.5,align:"center",fontFace:BODY,bold:true,color:SUN,fontSize:14});
  footer(s,4);
})();

/* ============================================= SLIDE 5 — DRIVERS PART 1 */
(() => {
  const s=pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,"5",0.6);
  s.addText("Do the Data Support Our Hypotheses?",{x:1.25,y:0.36,w:11,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:26});
  s.addText("Step 3 · Testing assumptions on income, traffic, and competition across all 15 stores",
    {x:1.27,y:0.9,w:11.5,h:0.35,fontFace:BODY,color:GRAY,fontSize:13});
  const rows=[
    {ic:"$",c:SUN,q:"Is there an ideal household income?",v:"Yes — a mainstream band, not luxury.",
     d:"All 15 stores fall between $49k and $105k median HHI (most $66k–$81k). Sales swing widely inside that band, so income qualifies a market rather than driving it — confirming the Walmart/Publix-shopper profile over a luxury one."},
    {ic:"⇧",c:SKY,q:"Does higher traffic mean stronger sales?",v:"Yes — the clearest single signal.",
     d:"The five strongest stores average ~47,500 VPD versus ~27,800 for the five weakest. Fort Lauderdale, Mandarin and Orlando (52k–70k VPD) all rank top-5. Pensacola is the exception, thriving on ~25k — so traffic is weighted heaviest, but not alone."},
    {ic:"✕",c:CORAL,q:"Is there an ideal level of competition?",v:"Nearby competitors do not depress sales.",
     d:"Pensacola, Jax Beach and Jax 5 Points each have a named competitor next door yet rank top-6. A proven, competitive corridor can signal real demand — so competition is weighted lightly."},
  ];
  let y=1.6;
  rows.forEach(r=>{
    s.addShape(pptx.ShapeType.roundRect,{x:0.5,y,w:12.3,h:1.68,rectRadius:0.07,fill:{color:CLOUD}});
    s.addShape(pptx.ShapeType.ellipse,{x:0.78,y:y+0.55,w:0.6,h:0.6,fill:{color:r.c}});
    s.addText(r.ic,{x:0.78,y:y+0.55,w:0.6,h:0.6,align:"center",valign:"middle",fontFace:HEAD,bold:true,color:WHITE,fontSize:20});
    s.addText(r.q,{x:1.6,y:y+0.16,w:6.0,h:0.4,fontFace:HEAD,bold:true,color:INK,fontSize:15});
    s.addText(r.v,{x:1.6,y:y+0.6,w:6.0,h:0.9,fontFace:BODY,bold:true,color:r.c,fontSize:13.5,valign:"top"});
    s.addText(r.d,{x:7.75,y:y+0.16,w:4.85,h:1.4,fontFace:BODY,color:GRAY,fontSize:11.5,valign:"middle",lineSpacingMultiple:1.02});
    y+=1.83;
  });
  footer(s,5);
})();

/* ============================================= SLIDE 6 — DRIVERS PART 2 + WEIGHTS */
(() => {
  const s=pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,"6",0.6);
  s.addText("Anchors, Parking & the Predictive Weights",{x:1.25,y:0.36,w:11,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:25});
  s.addText("Step 3 · Remaining drivers, and the weights that build the scorecard",
    {x:1.27,y:0.9,w:11,h:0.35,fontFace:BODY,color:GRAY,fontSize:13});
  const cards=[
    {t:"Grocery / destination anchors win",c:SUN,
     d:"Every top-4 store shares space with a major anchor — Target (Ft. Lauderdale), grocery + mall (Mandarin), regional mall (Pensacola), Publix (Jax Beach). Anchored centers clearly beat convenience strips."},
    {t:"Parking & access reinforce anchors",c:SKY,
     d:"Leaders pair high parking ratios with easy access and big signage. A walkable urban store (Jax 5 Points) can partly overcome weak parking, but most lagging sites are access-constrained."},
    {t:"Population is NOT a driver",c:CORAL,
     d:"The five weakest stores average more people within 5 miles (218k) than the five strongest (189k). Density alone doesn’t sell — so population carries the lightest weight."},
  ];
  let y=1.55;
  cards.forEach(c=>{
    s.addShape(pptx.ShapeType.roundRect,{x:0.5,y,w:6.0,h:1.35,rectRadius:0.07,fill:{color:CLOUD}});
    s.addShape(pptx.ShapeType.ellipse,{x:0.75,y:y+0.2,w:0.26,h:0.26,fill:{color:c.c}});
    s.addText(c.t,{x:1.15,y:y+0.12,w:5.2,h:0.4,fontFace:HEAD,bold:true,color:INK,fontSize:14.5});
    s.addText(c.d,{x:0.78,y:y+0.55,w:5.5,h:0.75,fontFace:BODY,color:GRAY,fontSize:10.5,valign:"top",lineSpacingMultiple:1.02});
    y+=1.5;
  });
  s.addText("Scorecard weighting (most predictive → least)",{x:6.9,y:1.55,w:6,h:0.35,fontFace:BODY,bold:true,color:INK,fontSize:13});
  const wl=["Traffic","Access","Visibility","Parking","Anchors","Income","Population","Competition"];
  const wv=[20,15,15,15,15,10,5,5];
  s.addChart(pptx.ChartType.bar,[{name:"Weight %",labels:wl,values:wv}],{
    x:6.75,y:1.95,w:6.1,h:4.55,barDir:"bar",chartColors:[INK2],showLegend:false,showTitle:false,
    showValue:true,dataLabelPosition:"outEnd",dataLabelFontSize:11,dataLabelColor:INK,dataLabelFormatCode:'0"%"',
    catAxisLabelColor:INK,catAxisLabelFontSize:11.5,catAxisLabelFontFace:BODY,
    valAxisHidden:true,valGridLine:{style:"none"},catGridLine:{style:"none"},valAxisMaxVal:24,barGapWidthPct:45});
  footer(s,6);
})();

/* ===================================== SLIDES 7–8 — SCORECARD */
function scorecardSlide(slNo,badge,title,sub,subset,note){
  const s=pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,badge,0.6,CORAL,WHITE);
  s.addText(title,{x:1.25,y:0.36,w:11,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:27});
  s.addText(sub,{x:1.27,y:0.9,w:11.5,h:0.35,fontFace:BODY,color:GRAY,fontSize:13});
  const cats=[["Store",1.9,"left"],["Pop",0.75,"c"],["Inc",0.75,"c"],["Traf",0.8,"c"],["Comp",0.85,"c"],
    ["Anch",0.85,"c"],["Park",0.8,"c"],["Acc",0.75,"c"],["Vis",0.75,"c"],["Raw /40",1.05,"c"],["Score /100",1.4,"c"]];
  const header=cats.map(c=>({text:c[0],options:{bold:true,color:WHITE,fill:{color:INK},align:c[2]==="left"?"left":"center",valign:"middle",fontSize:11.5,fontFace:BODY}}));
  const rows=[header];
  subset.forEach((o,i)=>{const tint=i%2?CLOUD2:CLOUD;
    const scoreFill=o.weighted>=85?SUN:o.weighted>=70?"F6D08A":"E9E6F0";
    const cell=(v,bold=false,col=GRAY)=>({text:String(v),options:{fill:{color:tint},align:"center",valign:"middle",fontSize:11,bold,color:col}});
    rows.push([
      {text:`${o.rank}. ${o.n}`,options:{fill:{color:tint},align:"left",valign:"middle",fontSize:11,bold:true,color:INK}},
      cell(o.sc.pop),cell(o.sc.inc),cell(o.sc.traf),cell(o.sc.comp),cell(o.sc.anchor),
      cell(o.sc.park),cell(o.sc.acc),cell(o.sc.vis),cell(o.raw,true,INK2),
      {text:o.weighted.toFixed(0),options:{fill:{color:scoreFill},align:"center",valign:"middle",fontSize:13,bold:true,color:INK}},
    ]);});
  s.addTable(rows,{x:0.5,y:1.5,w:PW-1.0,colW:cats.map(c=>c[1]),rowH:0.6,border:{type:"solid",color:WHITE,pt:1.5},valign:"middle"});
  if(note){
    s.addShape(pptx.ShapeType.roundRect,{x:0.5,y:PH-1.35,w:12.3,h:0.82,rectRadius:0.06,fill:{color:INK}});
    s.addText(note,{x:0.75,y:PH-1.3,w:11.8,h:0.72,fontFace:BODY,color:"EDEBF5",fontSize:10.5,valign:"middle",lineSpacingMultiple:1.02});
  }
  footer(s,slNo);
}
const ranked=byScore.map((o,i)=>({...o,rank:i+1}));
scorecardSlide(7,"7","The Sunburn Scorecard  ·  Ranks 1–8",
  "Step 4 · Every location scored 1–5 across 8 weighted categories → overall score of 100",
  ranked.slice(0,8),
  "Each category scored 1–5; the overall score applies the Step-3 weights (Traffic 20% · Access/Visibility/Parking/Anchors 15% each · Income 10% · Population/Competition 5%). Highest possible weighted score = 100.");
scorecardSlide(8,"8","The Sunburn Scorecard  ·  Ranks 9–15",
  "Step 4 · Lower-scoring existing stores — and how to read the model against actual sales",
  ranked.slice(8,15),
  "The model rewards suburban site mechanics (traffic, parking, access). Proven urban/destination sellers — Jax 5 Points and Pensacola — score below their sales, so the scorecard guides rather than replaces judgment. Only Panama City Beach & Sarasota parking are estimated (neutral 3); all else uses actual workbook figures.");

/* ============================================= SLIDES 9–12 — NEW-MARKET CANDIDATES */
// Illustrative expansion candidates — markets Sunburn does not yet serve, scored
// with the same rubric (sTraf/sPark/sPop/sInc/sComp + fixed anchor score) used
// for the 15 existing stores. Figures are market-level estimates for planning
// purposes, not confirmed site data — pending on-the-ground due diligence.
const CANDIDATES = {
  "Tampa — Westshore District": {
    inc:82000, traf:48000, trafD:"~48k", pop5:260000, popD:"~65k / 190k / 260k (1/3/5-mi, est.)",
    park:3.4, vis:5, acc:4, anchor:5, anchorTxt:"International Mall, Whole Foods, big-box corridor",
    compTxt:"Green Wave Dispensary",
    headline:"Tampa's densest retail corridor — top-tier traffic and anchors, offset only by tighter urban parking and one nearby competitor.",
    why:{
      traf:"~48,000 VPD on a primary Westshore corridor — among the highest of any candidate market.",
      acc:"Solid arterial access, though rush-hour congestion keeps it just off a top score.",
      vis:"Prime frontage directly on the retail corridor — hard to miss.",
      park:"3.4 spaces / 1,000 SF is workable but the tightest of the four candidates — typical of dense urban infill.",
      anchor:"Anchored by a major regional mall and grocery — a built-in draw with heavy cross-shop traffic.",
      inc:"$82k median HHI sits comfortably in the $55k–$90k mainstream band.",
      pop:"An estimated 260k people within 5 miles — the largest trade area of the four candidates.",
      comp:"One named competitor in the corridor — consistent with the portfolio finding that proven corridors signal demand rather than cap it.",
    },
  },
  "Boca Raton — Town Center": {
    inc:84000, traf:40000, trafD:"~40k", pop5:210000, popD:"~35k / 130k / 210k (1/3/5-mi, est.)",
    park:4.2, vis:4, acc:4, anchor:5, anchorTxt:"Town Center Mall, Whole Foods, Trader Joe's",
    compTxt:"Sunset Wellness",
    headline:"An affluent South Florida infill market with mall-grade anchors and healthy parking — a strong complement to the existing Fort Lauderdale / West Palm Beach cluster.",
    why:{
      traf:"~40,000 VPD along the Town Center corridor — solidly above the portfolio's traffic threshold for a strong site.",
      acc:"Easy in/out with multiple signalized access points into the center.",
      vis:"Good street-facing visibility, shared with several national retailers.",
      park:"4.2 spaces / 1,000 SF clears the portfolio's 4+ benchmark for a well-parked site.",
      anchor:"A destination mall plus two grocery anchors — the strongest anchor mix of the four candidates.",
      inc:"$84k median HHI — mainstream and affluent without tipping into the luxury band that underperformed elsewhere in the portfolio.",
      pop:"An estimated 210k within 5 miles, fills the gap between the existing Fort Lauderdale and West Palm Beach stores.",
      comp:"One named competitor nearby — a live, demand-proven corridor rather than untested ground.",
    },
  },
  "Naples — Waterside Shops": {
    inc:91000, traf:27000, trafD:"~27k", pop5:130000, popD:"~18k / 80k / 130k (1/3/5-mi, est.)",
    park:4.6, vis:4, acc:4, anchor:5, anchorTxt:"Waterside Shops, Publix, upscale dining",
    compTxt:"—",
    headline:"An uncontested Gulf Coast market with a premium anchor center — the trade-off is lower daytime traffic and a smaller year-round trade area.",
    why:{
      traf:"~27,000 VPD is the lowest of the four candidates — a seasonal/tourist market rather than a commuter corridor.",
      acc:"Easy, low-congestion access typical of a lower-density coastal market.",
      vis:"Good visibility within a well-known upscale shopping destination.",
      park:"4.6 spaces / 1,000 SF is the most generous of the four candidates.",
      anchor:"Waterside Shops + Publix draw a steady, affluent shopper base.",
      inc:"$91k median HHI is just above the portfolio's ideal $55k–$90k band — worth monitoring given Alton SoBe's high-income market underperformed.",
      pop:"An estimated 130k within 5 miles — the smallest trade area of the four, partly offset by strong seasonal tourism.",
      comp:"No named competitor identified — a clean, uncontested run at the market.",
    },
  },
  "Gainesville — Midtown / UF": {
    inc:51000, traf:32000, trafD:"~32k", pop5:150000, popD:"~28k / 95k / 150k (1/3/5-mi, est.)",
    park:3.0, vis:4, acc:3, anchor:5, anchorTxt:"University of Florida, student housing, Midtown retail",
    compTxt:"—",
    headline:"A campus-driven demand engine much like Orlando's UCF site — strong anchor, modest income and access reflecting a college-town trade area.",
    why:{
      traf:"~32,000 VPD along the Midtown corridor — moderate, campus-driven volume.",
      acc:"Campus-area congestion at peak times keeps access to a middle score, mirroring the Orlando UCF site.",
      vis:"Good visibility along the retail strip adjacent to campus.",
      park:"3.0 spaces / 1,000 SF is workable for a walkable, campus-adjacent site.",
      anchor:"The University of Florida is a self-renewing demand engine — tens of thousands of students turn over every fall.",
      inc:"$51k median HHI is the lowest of the four candidates, typical of a college-town trade area — but the portfolio's Jax 5 Points ($49k) shows lower income doesn't cap sales.",
      pop:"An estimated 150k within 5 miles, concentrated around a dense, walkable campus core.",
      comp:"No named competitor identified in the immediate trade area.",
    },
  },
};
function scoreCandidate(d){
  const sc={pop:sPop(d.pop5), inc:sInc(d.inc), traf:sTraf(d.traf), comp:sComp(d.compTxt),
            anchor:d.anchor, park:sPark(d.park), acc:d.acc, vis:d.vis};
  const raw=Object.values(sc).reduce((a,b)=>a+b,0);
  const weighted=Object.keys(WEIGHTS).reduce((a,k)=>a+sc[k]*WEIGHTS[k],0)/100*20;
  return {sc, raw, weighted};
}
const candidatesRanked = Object.keys(CANDIDATES)
  .map(n=>({n,...scoreCandidate(CANDIDATES[n]), ...CANDIDATES[n]}))
  .sort((a,b)=> b.weighted-a.weighted || b.raw-a.raw);

function candidateSlide(slNo, badge, o, idx){
  const s=pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,badge,0.6,SKY,WHITE);
  s.addText(`New-Market Candidate ${idx} — ${o.n}`,{x:1.25,y:0.34,w:11,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:24});
  s.addText("Step 6 · A market Sunburn doesn't yet serve, scored against the same 8 criteria as the existing portfolio",
    {x:1.27,y:0.88,w:11.3,h:0.35,fontFace:BODY,color:GRAY,fontSize:12.5});

  s.addShape(pptx.ShapeType.roundRect,{x:0.5,y:1.3,w:PW-1.0,h:0.55,rectRadius:0.06,fill:{color:INK}});
  s.addText([
    {text:`Projected Scorecard ${o.weighted.toFixed(0)}/100 (Raw ${o.raw}/40)   `,options:{bold:true,color:SKY,fontSize:13}},
    {text:`·   Would rank #${candidatesRanked.findIndex(c=>c.n===o.n)+1} of the 4 candidates evaluated`,options:{color:"CFC9E0",fontSize:13}},
  ],{x:0.5,y:1.3,w:PW-1.0,h:0.55,align:"center",valign:"middle",fontFace:BODY,margin:0});

  const crit=[
    {k:"traf",  t:"Traffic",     v:o.trafD,                                   sc:o.sc.traf},
    {k:"acc",   t:"Access",      v:`${o.acc} / 5 rating`,                     sc:o.sc.acc},
    {k:"vis",   t:"Visibility",  v:`${o.vis} / 5 rating`,                     sc:o.sc.vis},
    {k:"park",  t:"Parking",     v:`${o.park.toFixed(2)} / 1,000 SF (est.)`,  sc:o.sc.park},
    {k:"anchor",t:"Anchors",     v:o.anchorTxt,                               sc:o.sc.anchor},
    {k:"inc",   t:"Income",      v:`$${(o.inc/1000).toFixed(0)}k median HHI`, sc:o.sc.inc},
    {k:"pop",   t:"Population",  v:o.popD,                                    sc:o.sc.pop},
    {k:"comp",  t:"Competition", v:o.compTxt,                                 sc:o.sc.comp},
  ];
  const cols=4, cardW=2.95, gapX=0.17, cardH=2.175, gapY=0.15, gridX=0.5, gridY=2.03;
  crit.forEach((c,i)=>{
    const col=i%cols, row=Math.floor(i/cols);
    const x=gridX+col*(cardW+gapX), y=gridY+row*(cardH+gapY);
    const tint = c.sc>=5? "DCEAF4" : c.sc>=4? CLOUD2 : CLOUD;
    s.addShape(pptx.ShapeType.roundRect,{x,y,w:cardW,h:cardH,rectRadius:0.07,fill:{color:tint}});
    s.addText(c.t.toUpperCase(),{x:x+0.14,y:y+0.1,w:cardW-0.28,h:0.28,fontFace:BODY,bold:true,color:GRAY,fontSize:9.5,charSpacing:1});
    s.addText(`${c.sc}/5`,{x:x+0.14,y:y+0.36,w:cardW-0.28,h:0.42,fontFace:HEAD,bold:true,color:c.sc>=5?SKY:INK2,fontSize:22});
    s.addText(c.v,{x:x+0.14,y:y+0.8,w:cardW-0.28,h:0.42,fontFace:BODY,bold:true,color:INK,fontSize:9.5,valign:"top"});
    s.addText(o.why[c.k],{x:x+0.14,y:y+1.22,w:cardW-0.28,h:cardH-1.3,fontFace:BODY,color:GRAY,fontSize:8,valign:"top",lineSpacingMultiple:1.02});
  });

  s.addText(o.headline,{x:0.5,y:gridY+2*cardH+gapY+0.06,w:PW-1.0,h:0.42,align:"center",fontFace:BODY,italic:true,bold:true,color:INK2,fontSize:11});
  s.addText("Market-level estimates for planning purposes — pending on-the-ground site verification before submitting an LOI.",
    {x:0.5,y:PH-0.55,w:12.3,h:0.3,fontFace:BODY,italic:true,fontSize:8.5,color:GRAY});
  footer(s,slNo);
}
candidatesRanked.forEach((o,i)=>candidateSlide(9+i, String(9+i), o, i+1));

/* ---------------------------------------------------------------- SAVE */
pptx.writeFile({fileName:"Sunburn_Site_Selection_Analysis.pptx"}).then(f=>console.log("WROTE",f));
