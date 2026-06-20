const path = require('path');
const GLOBAL = 'C:/Users/priya/AppData/Roaming/npm/node_modules';
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, TableOfContents, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak, VerticalAlign } = require(path.join(GLOBAL, 'docx'));
const fs = require('fs');

const PINK = "C2185B", AMBER = "B36B00", DARK = "1A1A1A", GREY = "666666",
  RED = "C0392B", GREEN = "1E8449", LIGHTP = "FCE4EC", LIGHTA = "FFF3E0",
  HEAD = "37474F", ROWALT = "F5F5F5";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const CW = 9360;

// ---------- helpers ----------
function H1(t){return new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun(t)]});}
function H2(t){return new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun(t)]});}
function H3(t){return new Paragraph({heading:HeadingLevel.HEADING_3,children:[new TextRun(t)]});}
function P(runs,opts={}){return new Paragraph(Object.assign({children:Array.isArray(runs)?runs:[new TextRun(runs)],spacing:{after:120}},opts));}
function bullet(text,lvl=0){return new Paragraph({numbering:{reference:"bul",level:lvl},spacing:{after:60},children:typeof text==='string'?[new TextRun(text)]:text});}
function num(text){return new Paragraph({numbering:{reference:"ord",level:0},spacing:{after:60},children:typeof text==='string'?[new TextRun(text)]:text});}
function R(t,o={}){return new TextRun(Object.assign({text:t},o));}

function cell(content, {w, fill, bold, color, align, size}={}){
  const runs = Array.isArray(content)?content:[new TextRun({text:String(content),bold:!!bold,color:color||undefined,size:size||20})];
  return new TableCell({
    borders, width:{size:w,type:WidthType.DXA},
    shading: fill?{fill,type:ShadingType.CLEAR}:undefined,
    margins:{top:60,bottom:60,left:110,right:110},
    verticalAlign:VerticalAlign.CENTER,
    children:[new Paragraph({alignment:align||AlignmentType.LEFT,children:runs})]
  });
}

function tableHeaderRow(cells, widths){
  return new TableRow({tableHeader:true,children:cells.map((c,i)=>cell(c,{w:widths[i],fill:HEAD,bold:true,color:"FFFFFF",align:AlignmentType.CENTER,size:18}))});
}

// ---------- data ----------
const videos=[
 ['3 May','Sunday',"Khela Ends... Or Begins? Bengal Elections '26",'47:12','15:34','33.0%','203.3K','65%','98.2%','0.4%','≈ same'],
 ['17 May','Sunday',"NEET Scam & Modi's Appeal",'48:43','16:58','34.9%','226.9K','64%','98.3%','0.3%','+28.9K'],
 ['24 May','Sunday',"CockRoach Party, Melody, NEET Protest",'47:52','15:58','33.4%','211.7K','64%','98.3%','0.7%','+13.7K'],
 ['31 May','Sunday',"Paper Leak, Vaibhav S., Karnataka Politics",'44:28','15:36','35.1%','209.4K','64%','99.1%','1.7%','+3.4K'],
 ['7 Jun','Sunday',"Inside CJP Protest, Malviya Nagar & More",'32:47','10:46','32.9%','165.4K','64%','98.1%','1.2%','≈ same'],
 ['14 Jun','Sunday',"370 Biryani, Modi record, Khan Sir Trouble?",'45:52','16:59','37.0%','221.5K','65%','97.9%','0.7%','+20.5K'],
 ['20 May','Midweek',"Melody Vs. Cockroach Janta Party",'19:37','9:10','46.8%','110.2K','70%','97.9%','0.6%','≈ same'],
 ['27 May','Midweek',"Media on Petrol, Ranveer Banned, CJP Blocked",'18:06','9:13','51.0%','90.3K','70%','98.7%','0.8%','-9.1K'],
 ['3 Jun','Midweek',"Anjana Vs. Khan Sir, CJP Protest, Delhi Fire",'24:07','10:47','44.8%','147.1K','71%','98.7%','1.1%','≈ same'],
];

// ---------- build content ----------
const kids=[];

// Title page
kids.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:1800,after:0},children:[R("CHANNEL PERFORMANCE REPORT",{bold:true,size:40,color:PINK})]}));
kids.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[R("@SundaySarthak",{bold:true,size:56,color:DARK})]}));
kids.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:40},children:[R("Why Viewers Don't Watch Fully — Retention, Engagement & Reach Analysis",{size:26,color:GREY,italics:true})]}));
kids.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:200,after:0},children:[R("Sunday Show  vs  Midweek",{bold:true,size:28,color:AMBER})]}));
kids.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:600},children:[R("Based on 9 YouTube Studio exports  ·  3 May – 14 June 2026",{size:20,color:GREY})]}));
kids.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:40},children:[R("Prepared: 20 June 2026",{size:20,color:GREY})]}));
kids.push(new Paragraph({children:[new PageBreak()]}));

// TOC
kids.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun("Table of Contents")]}));
kids.push(new TableOfContents("Table of Contents",{hyperlink:true,headingStyleRange:"1-2"}));
kids.push(new Paragraph({children:[new PageBreak()]}));

// 1. Executive summary
kids.push(H1("1. Executive Summary"));
kids.push(P([R("Your channel's content is clearly working: ",{}),R("like ratios sit around 98%, watch-time per video is strong (90K–227K hours), and most videos meet or beat their typical performance.",{bold:true})]));
kids.push(P([R("The real problem is not whether people watch — it is ",{}),R("how far into each video they get. ",{bold:true,color:RED}),R("On the Sunday Show, the average viewer watches only about ",{}),R("one-third (33–37%)",{bold:true,color:RED}),R(" of the video before leaving. The Midweek show holds viewers far better, at ",{}),R("45–51%.",{bold:true,color:GREEN})]));
kids.push(P([R("One-line diagnosis: ",{bold:true}),R("the content is good, but the Sunday Show is too long for how it is paced. A 45-minute single-desk format mechanically pushes \"percentage viewed\" down, a sharp drop in the first 30 seconds loses ~30% of viewers, and recurring mid-video dips bleed away the rest. The highest-impact fixes are the intro hook, tighter pacing, chapter structure, and a Shorts strategy.")]));

kids.push(H3("Headline numbers"));
const sumTbl=new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[3120,3120,3120],rows:[
  tableHeaderRow(["Metric","Sunday Show","Midweek"],[3120,3120,3120]),
  new TableRow({children:[cell("Episodes analysed",{w:3120}),cell("6",{w:3120,align:AlignmentType.CENTER}),cell("3",{w:3120,align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("Avg video length",{w:3120}),cell("44:32",{w:3120,align:AlignmentType.CENTER}),cell("20:37",{w:3120,align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("Avg view duration",{w:3120}),cell("15:18",{w:3120,align:AlignmentType.CENTER}),cell("9:43",{w:3120,align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("Avg % of video viewed",{w:3120}),cell([R("34.7%",{bold:true,color:RED})],{w:3120,align:AlignmentType.CENTER}),cell([R("47.5%",{bold:true,color:GREEN})],{w:3120,align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("Viewers still watching at 0:30",{w:3120}),cell("64.3%",{w:3120,align:AlignmentType.CENTER}),cell("70.3%",{w:3120,align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("Avg watch-time per video",{w:3120}),cell("206.4K hrs",{w:3120,align:AlignmentType.CENTER}),cell("115.9K hrs",{w:3120,align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("Avg like ratio",{w:3120}),cell("98.3%",{w:3120,align:AlignmentType.CENTER}),cell("98.4%",{w:3120,align:AlignmentType.CENTER})]}),
]});
kids.push(sumTbl);
kids.push(P("",{spacing:{after:40}}));

// Data scope note
kids.push(H3("Important note on data scope"));
kids.push(P([R("The 9 source PDFs are all from the YouTube Studio ",{}),R("Engagement",{bold:true}),R(" tab. They contain retention, watch-time, average view duration, likes and end-screen data. They do ",{}),R("not",{bold:true}),R(" contain the ",{}),R("Reach",{bold:true}),R(" tab (impressions, thumbnail click-through rate, traffic sources). Retention is therefore analysed from hard data; the thumbnail/reach section uses the visible thumbnails, available signals and best practices. For a precise CTR/reach diagnosis, export the Reach tab for these same videos.")]));
kids.push(new Paragraph({children:[new PageBreak()]}));

// 2. Problem statement
kids.push(H1("2. Problem Statement — Why Viewers Don't Watch Fully"));
kids.push(P("Five concrete, data-backed problems explain the drop-off:"));

kids.push(H2("Problem 1 — Length is fighting your retention percentage"));
kids.push(P([R("Evidence: ",{bold:true}),R("Sunday Shows run 44–48 min and retain 33–37%. Midweek runs 18–24 min and retains 45–51%.")]));
kids.push(P("Average view duration is actually similar in minutes (Sunday ≈ 15–17 min, Midweek ≈ 9–11 min). But because Sunday videos are about 2.5× longer, the percentage viewed collapses. YouTube's recommendation system rewards high and consistent percentage-viewed; a 33% average signals \"people don't finish this,\" which limits how aggressively the video is pushed."));

kids.push(H2("Problem 2 — The first 30 seconds leak about 30% of viewers"));
kids.push(P([R("Evidence: ",{bold:true}),R("\"viewers still watching at 0:30\" is 64–65% on Sunday and 70–71% on Midweek. Every single video flags \"Intro\" as a key retention moment — i.e. the steepest drop happens right at the start.")]));
kids.push(P("A steep intro cliff means the opening is not hooking hard enough — too much greeting and setup before the payoff. Anyone not hooked in 30 seconds never reaches the good parts."));

kids.push(H2("Problem 3 — A recurring mid-video \"Dip\""));
kids.push(P([R("Evidence: ",{bold:true}),R("8 of 9 videos flag a \"Dip\" key moment; the retention curve shows a clear sag mid-way.")]));
kids.push(P("There is a predictable point — often a slower segment, a transition, or a long monologue — where a chunk of the audience drops. This is the second-biggest leak after the intro."));

kids.push(H2("Problem 4 — Weak end-screens lose onward views"));
kids.push(P([R("Evidence: ",{bold:true}),R("End-screen click rate is below the 0.8% channel average on most videos — as low as 0.3–0.4% on 3rd and 17th May.")]));
kids.push(P("Even viewers who finish are not clicking through to the next video. That hurts session time and suggested-video reach, because YouTube promotes creators who keep viewers on the platform."));

kids.push(H2("Problem 5 — The 7th June Sunday Show is an early-warning outlier"));
kids.push(P([R("Evidence: ",{bold:true}),R("It is the shortest Sunday (32 min) yet has the lowest average view duration (10:46, below usual) and the lowest percentage-viewed band (32.9%).")]));
kids.push(P("Shorter did not automatically help here — which suggests the hook, thumbnail, title or topic mix that week underperformed. It is worth studying what was different."));
kids.push(new Paragraph({children:[new PageBreak()]}));

// 3. Retention analysis
kids.push(H1("3. Retention Analysis"));
kids.push(P("All videos share the same retention shape: a sharp drop in the first 30 seconds, then a long gradual decline with a mid-video dip. The differences are in how steep and how low the curve goes."));
kids.push(H3("The intro cliff (viewers remaining at 0:30)"));
const r30=new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[4680,2340,2340],rows:[
  tableHeaderRow(["Video","Type","Still watching @ 0:30"],[4680,2340,2340]),
  ...videos.map(v=>new TableRow({children:[cell(v[2],{w:4680,size:18}),cell(v[1],{w:2340,align:AlignmentType.CENTER,size:18,color:v[1]==='Sunday'?PINK:AMBER,bold:true}),cell(v[7],{w:2340,align:AlignmentType.CENTER})]}))
]});
kids.push(r30);
kids.push(P("",{spacing:{after:40}}));
kids.push(P([R("Reading: ",{bold:true}),R("Midweek consistently keeps ~6 percentage points more of its audience past the intro. Closing that gap on the Sunday Show is the single most valuable retention win available.")]));

kids.push(H3("What the curve shape tells us"));
kids.push(bullet("Cliff at 0:00–0:30: the opening is not delivering a strong enough reason to stay. This is where redesign pays off most."));
kids.push(bullet("Gradual slope after the intro: normal and healthy — viewers leave steadily as the video runs. Tighter editing flattens this slope."));
kids.push(bullet("Mid-video dip: a specific weak segment. Open the retention graph in Studio, find the exact timestamp, and watch what is on screen there."));
kids.push(bullet("Spikes / \"top moments\": positive signals — moments viewers re-watch or where retention rises. These are ideal Shorts clips."));
kids.push(new Paragraph({children:[new PageBreak()]}));

// 4. Engagement & % viewed
kids.push(H1("4. How Much of Each Video Gets Watched"));
kids.push(P("Average percentage viewed is the clearest signal in the dataset. Midweek is consistently higher because it is tighter."));
const pctTbl=new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[1400,1400,2000,1500,1530,1530],rows:[
  tableHeaderRow(["Date","Type","Length","Avg Dur","% Viewed","Watch hrs"],[1400,1400,2000,1500,1530,1530]),
  ...videos.map((v,i)=>new TableRow({children:[
    cell(v[0],{w:1400,size:18}),
    cell(v[1],{w:1400,align:AlignmentType.CENTER,size:18,color:v[1]==='Sunday'?PINK:AMBER,bold:true}),
    cell(v[3],{w:2000,align:AlignmentType.CENTER,size:18}),
    cell(v[4],{w:1500,align:AlignmentType.CENTER,size:18}),
    cell([R(v[5],{bold:true,color:parseFloat(v[5])>=45?GREEN:(parseFloat(v[5])<34?RED:AMBER),size:18})],{w:1530,align:AlignmentType.CENTER}),
    cell(v[6],{w:1530,align:AlignmentType.CENTER,size:18}),
  ]}))
]});
kids.push(pctTbl);
kids.push(P("",{spacing:{after:40}}));
kids.push(P([R("Best Sunday: ",{bold:true}),R("14 June (37.0%) — slightly higher because it had strong \"top moment\" and spike segments. "),R("Best overall: ",{bold:true}),R("27 May Midweek (51.0%). "),R("Weakest: ",{bold:true}),R("7 June Sunday (32.9%).")]));
kids.push(new Paragraph({children:[new PageBreak()]}));

// 5. Sunday vs Midweek
kids.push(H1("5. Sunday Show vs Midweek — Head to Head"));
kids.push(P("Two different formats with two different performance profiles. Midweek wins on retention efficiency; Sunday wins on total volume. They are different tools, not better-or-worse."));
const cmp=new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[3360,3000,3000],rows:[
  tableHeaderRow(["Dimension","Sunday Show","Midweek"],[3360,3000,3000]),
  new TableRow({children:[cell("Format feel",{w:3360,bold:true}),cell("Flagship, long, desk-based",{w:3000,size:18}),cell("Punchy, topical, fast",{w:3000,size:18})]}),
  new TableRow({children:[cell("Retention %",{w:3360,bold:true}),cell([R("34.7% — weak",{color:RED,size:18})],{w:3000}),cell([R("47.5% — strong",{color:GREEN,size:18})],{w:3000})]}),
  new TableRow({children:[cell("Total watch-time",{w:3360,bold:true}),cell([R("206K hrs — strong",{color:GREEN,size:18})],{w:3000}),cell([R("116K hrs — lower",{color:AMBER,size:18})],{w:3000})]}),
  new TableRow({children:[cell("Intro hold (@0:30)",{w:3360,bold:true}),cell("64.3%",{w:3000,size:18}),cell("70.3%",{w:3000,size:18})]}),
  new TableRow({children:[cell("Main strength",{w:3360,bold:true}),cell("Reach volume & authority",{w:3000,size:18}),cell("Holds viewers to the end",{w:3000,size:18})]}),
  new TableRow({children:[cell("Main weakness",{w:3360,bold:true}),cell("Low % viewed, steep intro drop",{w:3000,size:18}),cell("Lower total reach",{w:3000,size:18})]}),
]});
kids.push(cmp);
kids.push(P("",{spacing:{after:40}}));
kids.push(P([R("Strategic takeaway: ",{bold:true}),R("Midweek's 47% retention proves your audience will watch tightly-paced content to a high percentage. The opportunity is to import Midweek's pacing discipline into the Sunday Show — or to formally split the Sunday Show into clear segments so its retention behaves more like Midweek's.")]));
kids.push(new Paragraph({children:[new PageBreak()]}));

// 6. Thumbnails & reach
kids.push(H1("6. Thumbnails & Reach"));
kids.push(P([R("Data caveat: ",{bold:true,color:RED}),R("the source PDFs do not include the Reach tab, so impressions and true thumbnail click-through rate are not measured here. The notes below use the visible thumbnails, the like/end-screen signals and proven best practices.")]));
kids.push(H3("What your thumbnails do well"));
kids.push(bullet("Strong, consistent branding — bold show logos (SUNDAY SHOW / MIDWEEK) make episodes instantly recognisable."));
kids.push(bullet("Emotional faces plus recognisable public figures drive curiosity."));
kids.push(bullet("Punchy red keyword tags (\"BANNED\", \"SCAM EXPOSED\", \"ENCOUNTER\") create urgency."));
kids.push(bullet("Clear duration badge and clean framing."));
kids.push(H3("Thumbnail risks worth testing"));
kids.push(bullet("Visual clutter: three to four faces plus multiple text blocks can read as busy on a phone. Test a single-subject, one-big-idea version."));
kids.push(bullet("Text legibility at small size: make sure the main 2–3 words are readable at thumbnail size on mobile."));
kids.push(bullet("Title–thumbnail pairing: the thumbnail should add information the title does not repeat, to create a curiosity gap."));
kids.push(bullet("Use YouTube's \"Test & Compare\" thumbnail feature on the next three uploads."));
kids.push(H3("Reach signals we can read from this data"));
kids.push(bullet([R("End-screen CTR is mostly below the 0.8% channel average ",{}),R("(0.3–1.7%). 31 May (1.7%) is the best — copy what that end-screen did.",{})]));
kids.push(bullet("Watch-time is healthy (90K–227K hrs per video), so reach itself is solid — the constraint is depth-of-view, not initial clicks."));
kids.push(bullet("Shorts/remixes are barely used (most videos show \"nothing to show\"). Shorts are the single biggest untapped reach channel."));
kids.push(new Paragraph({children:[new PageBreak()]}));

// 7. Full data table
kids.push(H1("7. Full Data — All 9 Videos"));
const fw=[900,1050,2150,900,900,900,1020,720,820];
const headers=["Date","Type","Title","Length","Avg Dur","% View","Watch","@0:30","Likes"];
const full=new Table({width:{size:CW,type:WidthType.DXA},columnWidths:fw,rows:[
  tableHeaderRow(headers,fw),
  ...videos.map(v=>new TableRow({children:[
    cell(v[0],{w:fw[0],size:16}),
    cell(v[1],{w:fw[1],align:AlignmentType.CENTER,size:16,color:v[1]==='Sunday'?PINK:AMBER,bold:true}),
    cell(v[2],{w:fw[2],size:16}),
    cell(v[3],{w:fw[3],align:AlignmentType.CENTER,size:16}),
    cell(v[4],{w:fw[4],align:AlignmentType.CENTER,size:16}),
    cell([R(v[5],{bold:true,size:16,color:parseFloat(v[5])>=45?GREEN:(parseFloat(v[5])<34?RED:AMBER)})],{w:fw[5],align:AlignmentType.CENTER}),
    cell(v[6],{w:fw[6],align:AlignmentType.CENTER,size:16}),
    cell(v[7],{w:fw[7],align:AlignmentType.CENTER,size:16}),
    cell(v[8],{w:fw[8],align:AlignmentType.CENTER,size:16}),
  ]}))
]});
kids.push(full);
kids.push(P("",{spacing:{after:40}}));
kids.push(P([R("Channel reference averages: ",{bold:true}),R("Like ratio 98.6% · End-screen CTR 0.8%.")],{spacing:{after:120}}));
kids.push(new Paragraph({children:[new PageBreak()]}));

// 8. Solutions
kids.push(H1("8. Solutions — Prioritised Fix Plan"));
kids.push(P("Ordered by impact-to-effort. Items 1–3 are the highest-leverage."));

kids.push(H2("1. Rebuild the first 30 seconds (highest impact)"));
kids.push(num("Cold open: lead with the single most shocking or curious line/clip of the episode before any greeting or branding."));
kids.push(num("Move \"Hello dosto / welcome\" to after the hook (around 0:30–0:45)."));
kids.push(num("Tease the episode's three biggest stories in the first 20 seconds."));
kids.push(P([R("Goal: ",{bold:true}),R("push \"viewers at 0:30\" from 64% to 75%+ on Sunday Shows.")]));

kids.push(H2("2. Add explicit chapters/segments to Sunday Shows"));
kids.push(num("Add YouTube chapters (timestamps in the description) so a 45-minute video feels navigable."));
kids.push(num("Use on-screen segment cards (\"STORY 2 / 5\") to create momentum and reset attention."));
kids.push(num("Order stories strongest-first, not chronological — front-load the best content."));

kids.push(H2("3. Attack the mid-video \"Dip\""));
kids.push(num("Open the retention graph, find the exact dip timestamp, and watch what is on screen — usually a long monologue or slow transition."));
kids.push(num("Add a pattern interrupt every 4–6 minutes: B-roll, a clip, a graphic, a tone change, or a \"but here's the twist\" line."));
kids.push(num("Cut anything that does not advance a story."));

kids.push(H2("4. Test a shorter Sunday target (don't assume)"));
kids.push(num("Your Midweek data proves tighter = higher percentage viewed. Test trimming the Sunday Show toward ~35 minutes for 3–4 weeks."));
kids.push(num("Caveat from 7 June: shorter alone is not a fix if the hook is weak — pair length cuts with the intro rebuild."));

kids.push(H2("5. Fix end-screens to lift onward reach"));
kids.push(num("Always end-screen to your best-retaining related video (e.g. a Midweek episode), not a generic one."));
kids.push(num("Point to it verbally and visually in the last 20 seconds."));
kids.push(num("Study the 31 May end-screen (1.7% CTR) and replicate its layout and call-to-action."));

kids.push(H2("6. Turn every episode into 3–5 Shorts (biggest reach upside)"));
kids.push(num("Most videos show no Shorts/remixes — this is free reach being left on the table."));
kids.push(num("Clip the spikes and \"top moments\" (the data already labels them) into vertical Shorts that funnel to the full episode."));
kids.push(num("20 May's clips already pulled 666 remix views with zero effort — a deliberate strategy would multiply that."));

kids.push(H2("7. Run a thumbnail A/B test program"));
kids.push(num("Use \"Test & Compare\" on upcoming uploads: cluttered multi-face vs clean single-subject."));
kids.push(num("Ensure the 2–3 hero words are readable at phone size; keep a curiosity gap with the title."));
kids.push(num("Export the Reach tab so actual CTR can be measured and the loop closed."));

kids.push(new Paragraph({spacing:{before:300},children:[R("Bottom line: ",{bold:true,size:24}),R("Your audience likes the content and gives you strong watch-time. Win back the first 30 seconds, give long videos a navigable structure, kill the mid-video dip, and convert episodes into Shorts — and both retention percentage and reach should rise together.",{size:24})]}));

// ---------- document ----------
const doc=new Document({
  creator:"Channel Analytics",
  title:"SundaySarthak Channel Performance Report",
  styles:{
    default:{document:{run:{font:"Calibri",size:22}}},
    paragraphStyles:[
      {id:"Title",name:"Title",basedOn:"Normal",next:"Normal",run:{size:48,bold:true,font:"Calibri"},paragraph:{spacing:{after:240}}},
      {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:32,bold:true,font:"Calibri",color:PINK},paragraph:{spacing:{before:280,after:160},outlineLevel:0}},
      {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:26,bold:true,font:"Calibri",color:HEAD},paragraph:{spacing:{before:200,after:120},outlineLevel:1}},
      {id:"Heading3",name:"Heading 3",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:23,bold:true,font:"Calibri",color:AMBER},paragraph:{spacing:{before:160,after:80},outlineLevel:2}},
    ]
  },
  numbering:{config:[
    {reference:"bul",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:540,hanging:280}}}}]},
    {reference:"ord",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:540,hanging:280}}}}]},
  ]},
  sections:[{
    properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,border:{bottom:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD",space:2}},children:[R("@SundaySarthak — Performance Report",{size:16,color:GREY})]})]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[R("Page ",{size:16,color:GREY}),new TextRun({children:[PageNumber.CURRENT],size:16,color:GREY}),R(" of ",{size:16,color:GREY}),new TextRun({children:[PageNumber.TOTAL_PAGES],size:16,color:GREY})]})]})},
    children:kids
  }]
});

Packer.toBuffer(doc).then(buf=>{fs.writeFileSync("SundaySarthak_Report.docx",buf);console.log("WROTE SundaySarthak_Report.docx",buf.length,"bytes");});
