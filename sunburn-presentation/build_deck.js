/* ============================================================================
   Sunburn Site Selection Project — Presentation Generator
   Produces a fully editable PowerPoint (.pptx) from the store + sales data.
   Run:  node build_deck.js
   ========================================================================== */
const pptxgen = require("pptxgenjs");

/* ---------------------------------------------------------------- PALETTE */
const INK   = "232042";  // deep indigo  (dark backgrounds / headings)
const INK2  = "3A3563";  // lighter indigo (bars / sub-panels)
const SUN   = "F4A11E";  // warm amber   (primary accent)
const CORAL = "EA5E3D";  // coral        (secondary accent / lowlights)
const SKY   = "3E7CB1";  // cool blue    (secondary data series)
const CLOUD = "F3F1F7";  // card tint
const CLOUD2= "E9E6F0";  // alt row tint
const GRAY  = "5A5670";  // muted text
const WHITE = "FFFFFF";
const HEAD  = "Cambria";
const BODY  = "Calibri";

/* ---------------------------------------------------------------- DATA */
// Trade-area + physical attributes (Stores sheet) and sales (Annual Sales sheet)
const S = {
 "Fort Lauderdale":   {city:"Fort Lauderdale", size:9300, pop:"—",                inc:null,   traffic:52000, comp:0, anchorTxt:"Target, Movie Theatre, Restaurants", compTxt:"—",             vis:5, acc:5, park:10.0, s25:5772937.32, s26:2695540.77},
 "West Palm Beach":   {city:"West Palm Beach", size:7874, pop:"—",                inc:null,   traffic:null,  comp:1, anchorTxt:"Restaurants, Police Dept",         compTxt:"Cloud 305",      vis:4, acc:2, park:1.0,  s25:2538655.36, s26:1165459.86},
 "Alton SoBe":        {city:"Miami Beach",     size:7038, pop:"40.2k / 98.5k / 384k", inc:81793, traffic:null, comp:0, anchorTxt:"Fast Food, Art Studio, Bank, CVS", compTxt:"—",             vis:4, acc:4, park:2.45, s25:1526448.33, s26:899441.95},
 "Pensacola":         {city:"Pensacola",       size:4000, pop:"—",                inc:null,   traffic:25500, comp:1, anchorTxt:"Mall, Fast Food, Hospital",       compTxt:"Legal Leaf",     vis:4, acc:4, park:3.5,  s25:4432053.85, s26:1925985.28},
 "Panama City Beach": {city:"Panama City Bch", size:2992, pop:"—",                inc:null,   traffic:22000, comp:0, anchorTxt:"Neighborhood Grill, Bank",        compTxt:"—",             vis:3, acc:3, park:null, s25:2915709.91, s26:1325678.97},
 "Tallahassee":       {city:"Tallahassee",     size:4434, pop:"—",                inc:null,   traffic:21707, comp:0, anchorTxt:"Restaurants, Salon, Coffee, Inn",  compTxt:"—",             vis:3, acc:4, park:4.96, s25:2794333.67, s26:1367265.12},
 "Jax Beach":         {city:"Jacksonville Bch",size:5690, pop:"10.3k / 44.6k / 108k", inc:104689, traffic:null, comp:1, anchorTxt:"Fast Food, Storage, Publix, Museum", compTxt:"Blaze N Haze", vis:4, acc:4, park:1.58, s25:3447040.73, s26:1697080.69},
 "Jax 5 Points":      {city:"Jacksonville",    size:6000, pop:"9.5k / 91.6k / 213k",  inc:49305,  traffic:null, comp:1, anchorTxt:"Restaurants/Clubs, Art Studio",  compTxt:"New Leaf Vapor", vis:3, acc:2, park:0.47, s25:3252971.33, s26:1559949.43},
 "Mandarin San Jose": {city:"Jacksonville",    size:5556, pop:"8.9k / 61.5k / 118k",  inc:91019,  traffic:52000, comp:0, anchorTxt:"Grocery, Mall, Restaurants, Books", compTxt:"—",           vis:5, acc:4, park:10.0, s25:4407877.45, s26:2179177.04},
 "Orlando":           {city:"Orlando",         size:2200, pop:"—",                inc:null,   traffic:70000, comp:1, anchorTxt:"UCF, Restaurants, Student Housing", compTxt:"Lifted Smoke", vis:4, acc:3, park:4.71, s25:3287080.42, s26:1320091.53},
 "Indialantic":       {city:"Indialantic",     size:2695, pop:"6.0k / 23.5k / 77.7k",  inc:77550,  traffic:null, comp:0, anchorTxt:"Liquor, Fast Food, Beach",       compTxt:"—",             vis:5, acc:5, park:null, s25:2875111.29, s26:1807569.82},
 "St. Pete":          {city:"St. Petersburg",  size:3200, pop:"23.9k / 101k / 225k",   inc:70829,  traffic:null, comp:0, anchorTxt:"Bank, Cafe, Theatre",           compTxt:"—",             vis:5, acc:4, park:null, s25:2828494.41, s26:1300661.68},
 "Sarasota":          {city:"Sarasota",        size:4498, pop:"14.1k / 72.6k / 151k",   inc:77078,  traffic:null, comp:0, anchorTxt:"Restaurants, Beach, Offices",    compTxt:"—",             vis:4, acc:3, park:null, s25:2869570.64, s26:1332116.10},
 "Cape Coral":        {city:"Cape Coral",      size:3297, pop:"9.8k / 69.4k / 190k",   inc:66219,  traffic:25000, comp:1, anchorTxt:"Arcade, Fast Food, Bank, Walgreens", compTxt:"Green Dragon", vis:3, acc:4, park:null, s25:2705318.35, s26:1395428.27},
 "Stuart":            {city:"Stuart",          size:3233, pop:"—",                inc:null,   traffic:46000, comp:0, anchorTxt:"Publix, Hotel, Strip Mall",       compTxt:"—",             vis:2, acc:5, park:3.77, s25:1730963.35, s26:1045900.05},
};

/* ---------------------------------------------------------------- SCORING */
// Rubrics derived from the Step-3 analysis. Blank source cells are scored at a
// neutral 3 (or income 4) pending field verification — flagged on the scorecard.
const wTraffic=(v)=> v==null?3 : v>=50000?5 : v>=40000?4 : v>=25000?3 : v>=20000?2 : 1;
const wPark   =(v)=> v==null?3 : v>=8?5 : v>=4?4 : v>=3?3 : v>=2?2 : 1;
const wPop    =(t)=> t==="—"?3 : (()=>{ // score by 5-mile figure
                     const m={"40.2k / 98.5k / 384k":5,"10.3k / 44.6k / 108k":3,"9.5k / 91.6k / 213k":4,
                              "8.9k / 61.5k / 118k":3,"6.0k / 23.5k / 77.7k":2,"23.9k / 101k / 225k":4,
                              "14.1k / 72.6k / 151k":4,"9.8k / 69.4k / 190k":4}; return m[t]||3; })();
const wInc    =(v)=> v==null?4 : (v>=55000&&v<=90000)?5 : (v>=50000&&v<95000)?4 : 3;
const wComp   =(c)=> c===0?5:3;
// anchor score by strength of co-tenancy
const anchorScore={"Fort Lauderdale":5,"West Palm Beach":3,"Alton SoBe":3,"Pensacola":5,"Panama City Beach":3,
  "Tallahassee":3,"Jax Beach":5,"Jax 5 Points":3,"Mandarin San Jose":5,"Orlando":5,"Indialantic":3,
  "St. Pete":3,"Sarasota":3,"Cape Coral":3,"Stuart":5};

const WEIGHTS = {traffic:20, acc:15, vis:15, park:15, anchor:15, inc:10, pop:5, comp:5};

function scoreStore(name){
  const d=S[name];
  const sc={pop:wPop(d.pop), inc:wInc(d.inc), traffic:wTraffic(d.traffic), comp:wComp(d.comp),
            anchor:anchorScore[name], park:wPark(d.park), acc:d.acc, vis:d.vis};
  const raw=Object.values(sc).reduce((a,b)=>a+b,0);            // out of 40
  const weighted=Object.keys(WEIGHTS).reduce((a,k)=>a+sc[k]*WEIGHTS[k],0)/100*20; // out of 100
  return {sc, raw, weighted};
}

const names = Object.keys(S);
const by25  = [...names].sort((a,b)=>S[b].s25-S[a].s25);
const by26  = [...names].sort((a,b)=>S[b].s26-S[a].s26);
const byScore = [...names].map(n=>({n,...scoreStore(n)})).sort((a,b)=>b.weighted-a.weighted);
const money = (v)=>"$"+(v/1e6).toFixed(2)+"M";

/* ---------------------------------------------------------------- DECK SETUP */
const pptx = new pptxgen();
pptx.defineLayout({ name:"W", width:13.333, height:7.5 });
pptx.layout = "W";
pptx.theme  = { headFontFace:HEAD, bodyFontFace:BODY };
const PW=13.333, PH=7.5;

/* helper: small filled circle "sun" badge with text */
function sunBadge(slide,x,y,txt,d=0.55,fill=SUN,color=INK){
  slide.addShape(pptx.ShapeType.ellipse,{x,y,w:d,h:d,fill:{color:fill}});
  slide.addText(txt,{x,y,w:d,h:d,align:"center",valign:"middle",fontFace:HEAD,bold:true,color,fontSize:18});
}
function footer(slide,n){
  slide.addText("Sunburn Site Selection Project",{x:0.5,y:PH-0.38,w:6,h:0.3,fontFace:BODY,fontSize:9,color:GRAY});
  slide.addText(String(n),{x:PW-0.9,y:PH-0.38,w:0.4,h:0.3,fontFace:BODY,fontSize:9,color:GRAY,align:"right"});
}

/* =========================================================== SLIDE 1 — TITLE */
(() => {
  const s = pptx.addSlide(); s.background={color:INK};
  // sun-ray motif: concentric arcs top-right via layered ellipses
  s.addShape(pptx.ShapeType.ellipse,{x:9.7,y:-2.4,w:6.4,h:6.4,fill:{color:INK2}});
  s.addShape(pptx.ShapeType.ellipse,{x:10.6,y:-1.5,w:4.6,h:4.6,fill:{color:"2E2A55"}});
  s.addShape(pptx.ShapeType.ellipse,{x:11.4,y:-0.7,w:3.0,h:3.0,fill:{color:SUN}});
  s.addText("SITE SELECTION ANALYSIS",{x:0.9,y:2.05,w:9,h:0.4,fontFace:BODY,bold:true,color:SUN,fontSize:15,charSpacing:3});
  s.addText("Sunburn Store Portfolio\n& Location Scorecard",{x:0.85,y:2.5,w:9.6,h:1.9,fontFace:HEAD,bold:true,color:WHITE,fontSize:46,lineSpacingMultiple:0.98});
  s.addText("A data-driven model of what makes a Sunburn location succeed — built from 15 existing stores — used to score and prioritize future retail sites.",
    {x:0.9,y:4.55,w:8.4,h:0.9,fontFace:BODY,color:"CFC9E0",fontSize:15,lineSpacingMultiple:1.1});
  // stat chips
  const chips=[["15","Existing stores"],["$47.4M","2025 portfolio sales"],["2–4","Sites to recommend"]];
  chips.forEach((c,i)=>{
    const x=0.9+i*3.05;
    s.addShape(pptx.ShapeType.roundRect,{x,y:5.75,w:2.8,h:1.15,rectRadius:0.1,fill:{color:INK2}});
    s.addText(c[0],{x,y:5.85,w:2.8,h:0.6,align:"center",fontFace:HEAD,bold:true,color:SUN,fontSize:28});
    s.addText(c[1],{x,y:6.42,w:2.8,h:0.4,align:"center",fontFace:BODY,color:"CFC9E0",fontSize:11});
  });
})();

/* =============================================== SLIDES 2–4 — LOCATION PROFILES */
function locationSlide(slNo, groupNo, subset){
  const s = pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,String(groupNo),0.6);
  s.addText("Store Trade-Area Profiles",{x:1.25,y:0.36,w:8,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:28});
  s.addText(`Step 2 · Location profiles ${subset[0].idx}–${subset[subset.length-1].idx} of 15`,
    {x:1.27,y:0.9,w:9,h:0.35,fontFace:BODY,color:GRAY,fontSize:13});

  const cols=[
    {t:"Store",w:1.75,a:"left"},{t:"Size",w:0.9,a:"center"},{t:"Pop 1/3/5-mi",w:2.15,a:"center"},
    {t:"Median HHI",w:1.15,a:"center"},{t:"Traffic (VPD)",w:1.2,a:"center"},{t:"Parking /1k",w:1.0,a:"center"},
    {t:"Vis / Acc",w:0.95,a:"center"},{t:"Anchor Retailers",w:2.55,a:"left"},{t:"Competitor",w:1.65,a:"left"},
  ];
  const header = cols.map(c=>({text:c.t,options:{bold:true,color:WHITE,fill:{color:INK},align:c.a,valign:"middle",fontSize:11,fontFace:BODY}}));
  const rows=[header];
  subset.forEach((it,i)=>{
    const d=S[it.name]; const tint = i%2? CLOUD2 : CLOUD;
    rows.push([
      {text:it.name,options:{bold:true,color:INK,fill:{color:tint},align:"left",valign:"middle",fontSize:11}},
      {text:(d.size/1000).toFixed(1)+"k",options:{fill:{color:tint},align:"center",valign:"middle",fontSize:10.5,color:GRAY}},
      {text:d.pop,options:{fill:{color:tint},align:"center",valign:"middle",fontSize:9.5,color:GRAY}},
      {text:d.inc?("$"+(d.inc/1000).toFixed(0)+"k"):"—",options:{fill:{color:tint},align:"center",valign:"middle",fontSize:10.5,color:GRAY}},
      {text:d.traffic?("~"+(d.traffic/1000).toFixed(0)+"k"):"—",options:{fill:{color:tint},align:"center",valign:"middle",fontSize:10.5,color:GRAY}},
      {text:d.park!=null?d.park.toFixed(2):"—",options:{fill:{color:tint},align:"center",valign:"middle",fontSize:10.5,color:GRAY}},
      {text:`${d.vis} / ${d.acc}`,options:{fill:{color:tint},align:"center",valign:"middle",fontSize:10.5,bold:true,color:INK2}},
      {text:d.anchorTxt,options:{fill:{color:tint},align:"left",valign:"middle",fontSize:9.5,color:GRAY}},
      {text:d.compTxt,options:{fill:{color:tint},align:"left",valign:"middle",fontSize:9.5,color:d.compTxt==="—"?GRAY:CORAL}},
    ]);
  });
  s.addTable(rows,{x:0.5,y:1.45,w:PW-1.0,colW:cols.map(c=>c.w),rowH:0.92,border:{type:"solid",color:WHITE,pt:1.5},valign:"middle"});
  s.addText("Population shown as 1-/3-/5-mile radius. “—” = data not yet captured in source workbook; to be field-verified.",
    {x:0.5,y:PH-0.72,w:11,h:0.3,fontFace:BODY,italic:true,fontSize:9.5,color:GRAY});
  footer(s,slNo);
}
const g1 = by25.slice(0,5).map((n,i)=>({name:n,idx:i+1}));   // top group by 2025 sales for logical grouping
// Present in portfolio order (not ranked) so slide is a neutral profile list:
const order = names;
locationSlide(2,1, order.slice(0,5).map((n,i)=>({name:n,idx:i+1})));
locationSlide(3,2, order.slice(5,10).map((n,i)=>({name:n,idx:i+6})));
locationSlide(4,3, order.slice(10,15).map((n,i)=>({name:n,idx:i+11})));

/* ===================================================== SLIDE 5 — RANKINGS */
(() => {
  const s = pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,"5",0.6,CORAL,WHITE);
  s.addText("Store Rankings by Sales",{x:1.25,y:0.36,w:9,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:28});
  s.addText("Step 3 · All 15 stores ranked — full-year 2025 vs. 2026 year-to-date (Jan–Jun)",
    {x:1.27,y:0.9,w:11,h:0.35,fontFace:BODY,color:GRAY,fontSize:13});

  s.addText("2025 · Full Year",{x:0.6,y:1.4,w:5.8,h:0.35,fontFace:BODY,bold:true,color:SUN,fontSize:14});
  s.addText("2026 · YTD (Jan–Jun)",{x:6.95,y:1.4,w:5.8,h:0.35,fontFace:BODY,bold:true,color:SKY,fontSize:14});

  // ascending so #1 sits at top of horizontal bar chart
  const asc25=[...by25].reverse(), asc26=[...by26].reverse();
  const mk=(labels,vals,color)=>[{name:"Sales",labels,values:vals}];
  const chartOpts=(color)=>({
    x:0.55, y:1.75, w:6.15, h:5.35, barDir:"bar",
    chartColors:[color], showLegend:false, showTitle:false,
    showValue:true, dataLabelPosition:"outEnd", dataLabelFontSize:8, dataLabelColor:INK,
    dataLabelFormatCode:'"$"0.00,,"M"',
    catAxisLabelColor:INK, catAxisLabelFontSize:8.5, catAxisLabelFontFace:BODY,
    valAxisHidden:true, valGridLine:{style:"none"}, catGridLine:{style:"none"},
    valAxisMaxVal:6500000, barGapWidthPct:35,
  });
  s.addChart(pptx.ChartType.bar, mk(asc25.map(n=>n),asc25.map(n=>S[n].s25)), {...chartOpts(SUN)});
  s.addChart(pptx.ChartType.bar, mk(asc26.map(n=>n),asc26.map(n=>S[n].s26)), {...chartOpts(SKY), x:6.9, valAxisMaxVal:3000000});

  footer(s,5);
})();

/* ============================================= SLIDE 6 — TOP vs BOTTOM TRAITS */
(() => {
  const s = pptx.addSlide(); s.background={color:INK};
  sunBadge(s,0.5,0.42,"6",0.6);
  s.addText("What Sets Top Stores Apart",{x:1.25,y:0.36,w:10,h:0.5,fontFace:HEAD,bold:true,color:WHITE,fontSize:28});
  s.addText("Step 3 · Shared characteristics of the highest- and lowest-performing locations",
    {x:1.27,y:0.9,w:11,h:0.35,fontFace:BODY,color:"CFC9E0",fontSize:13});

  const top5=by25.slice(0,5), bot5=by25.slice(-5);
  const avg=(arr,f)=>arr.reduce((a,n)=>a+f(S[n]),0)/arr.length;

  // TOP card
  const cardY=1.55, cardH=4.15;
  s.addShape(pptx.ShapeType.roundRect,{x:0.5,y:cardY,w:6.0,h:cardH,rectRadius:0.08,fill:{color:"2E2A55"}});
  s.addText([{text:"TOP 5  ",options:{color:SUN,bold:true}},{text:"— High Performers",options:{color:WHITE}}],
    {x:0.8,y:cardY+0.2,w:5.4,h:0.4,fontFace:HEAD,fontSize:17,bold:true});
  s.addText(top5.join("  ·  "),{x:0.8,y:cardY+0.62,w:5.4,h:0.35,fontFace:BODY,italic:true,color:"CFC9E0",fontSize:10.5});
  const topPts=[
    "Located on high-traffic corridors — avg. ~50k+ VPD where measured",
    "Generous parking (Ft. Lauderdale & Mandarin at 10 / 1,000 SF)",
    "Strong visibility + access (combined scores of 9–10 of 10)",
    "Co-located with destination anchors — Target, grocery, mall, Publix, UCF",
    "Middle-income trade areas — roughly $50k–$105k median HHI",
  ];
  s.addText(topPts.map((t,i)=>({text:t,options:{bullet:{code:"2022",indent:14},color:"EDEBF5",fontSize:12.5,paraSpaceAfter:9,breakLine:true}})),
    {x:0.85,y:cardY+1.05,w:5.4,h:3.0,valign:"top"});

  // BOTTOM card
  s.addShape(pptx.ShapeType.roundRect,{x:6.85,y:cardY,w:6.0,h:cardH,rectRadius:0.08,fill:{color:"3A2540"}});
  s.addText([{text:"BOTTOM 5  ",options:{color:CORAL,bold:true}},{text:"— Lagging",options:{color:WHITE}}],
    {x:7.15,y:cardY+0.2,w:5.4,h:0.4,fontFace:HEAD,fontSize:17,bold:true});
  s.addText(bot5.join("  ·  "),{x:7.15,y:cardY+0.62,w:5.4,h:0.35,fontFace:BODY,italic:true,color:"E8D6DE",fontSize:10.5});
  const botPts=[
    "Constrained parking & access — West Palm Beach at 1 / 1,000 SF, access 2/5",
    "Weak or convenience-only co-tenants (Alton SoBe: CVS, bank, fast food)",
    "Tourist / urban-infill sites without a demand anchor",
    "Stuart & Alton still ramping — Stuart opened mid-2025 ($0 in January)",
    "High income alone does not rescue a site (Alton SoBe, $82k, ranks last)",
  ];
  s.addText(botPts.map((t)=>({text:t,options:{bullet:{code:"2022",indent:14},color:"F3E6EC",fontSize:12.5,paraSpaceAfter:9,breakLine:true}})),
    {x:7.2,y:cardY+1.05,w:5.4,h:3.0,valign:"top"});

  s.addText(`Avg. 2025 sales — Top 5: ${money(avg(top5,d=>d.s25))}   vs.   Bottom 5: ${money(avg(bot5,d=>d.s25))}  (a ${(avg(top5,d=>d.s25)/avg(bot5,d=>d.s25)).toFixed(1)}× gap)`,
    {x:0.5,y:5.95,w:12.3,h:0.5,align:"center",fontFace:BODY,bold:true,color:SUN,fontSize:14});
  footer(s,6);
})();

/* ============================================= SLIDE 7 — DRIVERS PART 1 */
(() => {
  const s = pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,"7",0.6);
  s.addText("Do the Data Support Our Hypotheses?",{x:1.25,y:0.36,w:11,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:26});
  s.addText("Step 3 · Testing assumptions on income, traffic, and competition",
    {x:1.27,y:0.9,w:11,h:0.35,fontFace:BODY,color:GRAY,fontSize:13});

  const rows=[
    {ic:"$", c:SUN, q:"Is there an ideal household income?",
     v:"Yes — a middle band, not luxury.",
     d:"Strong stores cluster around $50k–$90k median HHI. The lowest-income area (Jax 5 Points, $49k) is a top-6 seller, while the pricey tourist area (Alton SoBe, $82k) ranks last. Confirms the Walmart/Publix-shopper profile over a luxury one."},
    {ic:"⇧", c:SKY, q:"Does higher traffic mean stronger sales?",
     v:"Generally yes — one of the clearest signals.",
     d:"The three highest-traffic sites — Fort Lauderdale, Orlando and Mandarin (~50k–70k VPD) — are all top-5 in at least one year. Traffic count earns the heaviest weight in the scorecard."},
    {ic:"⚔", c:CORAL, q:"Is there an ideal level of competition?",
     v:"Nearby competitors do not depress sales.",
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
  footer(s,7);
})();

/* ============================================= SLIDE 8 — DRIVERS PART 2 + WEIGHTS */
(() => {
  const s = pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,"8",0.6);
  s.addText("Anchors, Parking & the Predictive Weights",{x:1.25,y:0.36,w:11,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:25});
  s.addText("Step 3 · Remaining drivers, and the weights that build the scorecard",
    {x:1.27,y:0.9,w:11,h:0.35,fontFace:BODY,color:GRAY,fontSize:13});

  // left: two finding cards
  const findings=[
    {t:"Grocery / destination anchors win", c:SUN,
     d:"Every top-4 store shares space with a major anchor — Target (Ft. Lauderdale), grocery + mall (Mandarin), Publix (Jax Beach), regional mall (Pensacola). Grocery-anchored and destination centers clearly outperform convenience strips."},
    {t:"Parking & access reinforce anchors", c:SKY,
     d:"Leaders pair 10 / 1,000 SF parking with high access. Suburban/highway sites beat constrained urban ones — though a walkable urban store (Jax 5 Points) can overcome weak parking through foot traffic."},
  ];
  let y=1.6;
  findings.forEach(f=>{
    s.addShape(pptx.ShapeType.roundRect,{x:0.5,y,w:6.0,h:2.15,rectRadius:0.07,fill:{color:CLOUD}});
    s.addShape(pptx.ShapeType.ellipse,{x:0.75,y:y+0.25,w:0.28,h:0.28,fill:{color:f.c}});
    s.addText(f.t,{x:1.2,y:y+0.16,w:5.1,h:0.45,fontFace:HEAD,bold:true,color:INK,fontSize:15.5});
    s.addText(f.d,{x:0.78,y:y+0.68,w:5.5,h:1.35,fontFace:BODY,color:GRAY,fontSize:11.5,valign:"top",lineSpacingMultiple:1.04});
    y+=2.35;
  });

  // right: weight bar chart
  s.addText("Scorecard weighting (most predictive → least)",{x:6.9,y:1.55,w:6,h:0.35,fontFace:BODY,bold:true,color:INK,fontSize:13});
  const wl=["Traffic","Access","Visibility","Parking","Anchors","Income","Population","Competition"];
  const wv=[20,15,15,15,15,10,5,5];
  s.addChart(pptx.ChartType.bar,[{name:"Weight %",labels:wl,values:wv}],{
    x:6.75,y:1.95,w:6.1,h:4.55, barDir:"bar",
    chartColors:[INK2], showLegend:false, showTitle:false,
    showValue:true, dataLabelPosition:"outEnd", dataLabelFontSize:11, dataLabelColor:INK, dataLabelFormatCode:'0"%"',
    catAxisLabelColor:INK, catAxisLabelFontSize:11.5, catAxisLabelFontFace:BODY,
    valAxisHidden:true, valGridLine:{style:"none"}, catGridLine:{style:"none"},
    valAxisMaxVal:24, barGapWidthPct:45,
  });
  footer(s,8);
})();

/* ===================================== SLIDES 9–10 — SCORECARD */
function scorecardSlide(slNo, badge, title, sub, subset, note){
  const s = pptx.addSlide(); s.background={color:WHITE};
  sunBadge(s,0.5,0.42,badge,0.6,CORAL,WHITE);
  s.addText(title,{x:1.25,y:0.36,w:11,h:0.5,fontFace:HEAD,bold:true,color:INK,fontSize:27});
  s.addText(sub,{x:1.27,y:0.9,w:11.5,h:0.35,fontFace:BODY,color:GRAY,fontSize:13});

  const cats=[["Store",1.9,"left"],["Pop",0.75,"c"],["Inc",0.75,"c"],["Traf",0.8,"c"],["Comp",0.85,"c"],
              ["Anch",0.85,"c"],["Park",0.8,"c"],["Acc",0.75,"c"],["Vis",0.75,"c"],
              ["Raw /40",1.05,"c"],["Score /100",1.4,"c"]];
  const header=cats.map(c=>({text:c[0],options:{bold:true,color:WHITE,fill:{color:INK},align:c[2]==="left"?"left":"center",valign:"middle",fontSize:11.5,fontFace:BODY}}));
  const rows=[header];
  subset.forEach((o,i)=>{
    const tint=i%2?CLOUD2:CLOUD;
    const scoreFill = o.weighted>=85?SUN : o.weighted>=70?"F6D08A" : "E9E6F0";
    const cell=(v,bold=false,col=GRAY)=>({text:String(v),options:{fill:{color:tint},align:"center",valign:"middle",fontSize:11,bold,color:col}});
    rows.push([
      {text:`${o.rank}. ${o.n}`,options:{fill:{color:tint},align:"left",valign:"middle",fontSize:11,bold:true,color:INK}},
      cell(o.sc.pop),cell(o.sc.inc),cell(o.sc.traffic),cell(o.sc.comp),cell(o.sc.anchor),
      cell(o.sc.park),cell(o.sc.acc),cell(o.sc.vis),
      cell(o.raw,true,INK2),
      {text:o.weighted.toFixed(0),options:{fill:{color:scoreFill},align:"center",valign:"middle",fontSize:13,bold:true,color:INK}},
    ]);
  });
  s.addTable(rows,{x:0.5,y:1.5,w:PW-1.0,colW:cats.map(c=>c[1]),rowH:0.6,border:{type:"solid",color:WHITE,pt:1.5},valign:"middle"});
  if(note){
    s.addShape(pptx.ShapeType.roundRect,{x:0.5,y:PH-1.35,w:12.3,h:0.82,rectRadius:0.06,fill:{color:INK}});
    s.addText(note,{x:0.75,y:PH-1.3,w:11.8,h:0.72,fontFace:BODY,color:"EDEBF5",fontSize:10.5,valign:"middle",lineSpacingMultiple:1.02});
  }
  footer(s,slNo);
}
const ranked = byScore.map((o,i)=>({...o,rank:i+1}));
scorecardSlide(9,"9","The Sunburn Scorecard  ·  Ranks 1–8",
  "Step 4 · Every location scored 1–5 across 8 weighted categories → overall score of 100",
  ranked.slice(0,8),
  "Each category scored 1–5; the overall score applies the Step-3 weights (Traffic 20% · Access/Visibility/Parking/Anchors 15% each · Income 10% · Population/Competition 5%). Highest possible weighted score = 100.");
scorecardSlide(10,"10","The Sunburn Scorecard  ·  Ranks 9–15",
  "Step 4 · Lower-scoring existing stores — and how to read the model against actual sales",
  ranked.slice(8,15),
  "Note: the model rewards suburban physical attributes (parking, access, visibility). Proven urban sellers — Jax 5 Points, Jax Beach — score lower than they sell, so scores guide, not replace, judgment. Blank source cells were scored at a neutral 3 (income 4) pending field verification.");

/* ---------------------------------------------------------------- SAVE */
pptx.writeFile({ fileName:"Sunburn_Site_Selection_Analysis.pptx" }).then(f=>console.log("WROTE",f));
