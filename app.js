const SAMPLE_REGIONS = [
  "National Capital Region","Cordillera Administrative Region","Region I","Region II","Region III","Region IV-A","MIMAROPA","Region V","Region VI","Region VII","Region VIII","Region IX","Region X","Region XI","Region XII","Region XIII","Bangsamoro","CARAGA","Negros Island Region","Philippines"
];

let rawData=[]; let charts={};
const fmt=new Intl.NumberFormat("en-PH");
const compact=new Intl.NumberFormat("en-PH",{notation:"compact",maximumFractionDigits:1});

function normalizeRows(rows){
  return rows.map(r=>{
    const year=Number(r.Year ?? r.year ?? r["YEAR"]);
    const population=Number(r.Population ?? r.population ?? r["POPULATION"]);
    const region=String(r.Region ?? r.region ?? r["REGION"] ?? "Unknown").trim();
    return {Year:year,Region:region,Population:population};
  }).filter(r=>Number.isFinite(r.Year)&&Number.isFinite(r.Population)&&r.Population>0&&r.Region);
}

async function loadDefaultData(){
  const url="population_philippines_annual_2000_2025.xlsx";
  try{
    const res=await fetch(url);
    if(res.ok){
      const buffer=await res.arrayBuffer();
      const wb=XLSX.read(buffer,{type:"array"});
      const sheetName=wb.SheetNames.find(n=>/annual/i.test(n))||wb.SheetNames[0];
      rawData=normalizeRows(XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{defval:""}));
    }
  }catch(e){console.warn("Workbook could not be loaded",e)}

  if(!rawData.length) rawData=makeDemoData();
  initialize();
}

function makeDemoData(){
  const populations=[
    ["National Capital Region",1174279],["Cordillera Administrative Region",630800],["Region I",3002500],["Region II",3172000],["Region III",10900000],["Region IV-A",12800000],["MIMAROPA",2718000],["Region V",5410000],["Region VI",7032000],["Region VII",6900000],["Region VIII",4110000],["Region IX",3480000],["Region X",4450000],["Region XI",4870000],["Region XII",4270000],["Region XIII",2580000],["Bangsamoro",3770000],["CARAGA",2360000],["Negros Island Region",4240000],["Philippines",76600000]
  ];
  const rows=[];
  for(let year=2000;year<=2025;year++){
    const t=(year-2000)/25;
    populations.forEach(([region,start])=>rows.push({Year:year,Region:region,Population:Math.round(start*(0.82+t*0.18)/(1+0.02*(region.startsWith("Region")?0.2:0))+start*0.0032*(year-2000))}));
  }
  return rows;
}

function initialize(){
  const years=[...new Set(rawData.map(d=>d.Year))].sort((a,b)=>a-b);
  const regions=[...new Set(rawData.map(d=>d.Region))].sort((a,b)=>a.localeCompare(b));
  const from=document.getElementById("from-year"),to=document.getElementById("to-year"),region=document.getElementById("region-filter");
  from.innerHTML=years.map(y=>`<option>${y}</option>`).join(""); to.innerHTML=years.map(y=>`<option>${y}</option>`).join("");
  region.innerHTML=`<option value="ALL">All regions</option>`+regions.map(r=>`<option value="${escapeHtml(r)}">${escapeHtml(r)}</option>`).join("");
  from.value=years[0];to.value=years.at(-1);region.value="ALL";
  bindNavigation();bindUpload();bindDownload();bindTable();updateDashboard();document.getElementById("loading").classList.add("hidden");
}

function escapeHtml(s){return String(s).replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]))}
function getFiltered(){
  const r=document.getElementById("region-filter").value;const from=+document.getElementById("from-year").value;const to=+document.getElementById("to-year").value;
  return rawData.filter(d=>d.Year>=from&&d.Year<=to&&(r==="ALL"||d.Region===r));
}

function updateDashboard(){
  const filtered=getFiltered();const region=document.getElementById("region-filter").value;
  const years=[...new Set(filtered.map(d=>d.Year))].sort((a,b)=>a-b);const latestYear=years.at(-1);const firstYear=years[0];
  const latestRows=filtered.filter(d=>d.Year===latestYear);const firstRows=filtered.filter(d=>d.Year===firstYear);
  let latestTotal=latestRows.reduce((s,d)=>s+d.Population,0);let firstTotal=firstRows.reduce((s,d)=>s+d.Population,0);
  document.getElementById("metric-population").textContent=compact.format(latestTotal);document.getElementById("metric-pop-year").textContent=`${latestYear} estimate`;
  document.getElementById("metric-regions").textContent=[...new Set(filtered.map(d=>d.Region))].length;
  const growth=firstTotal?((latestTotal-firstTotal)/firstTotal)*100:0;document.getElementById("metric-growth").textContent=`${growth>=0?"+":""}${growth.toFixed(1)}%`;document.getElementById("metric-growth-period").textContent=`${firstYear}–${latestYear}`;
  document.getElementById("metric-average").textContent=compact.format(filtered.length?filtered.reduce((s,d)=>s+d.Population,0)/filtered.length:0);
  document.getElementById("trend-title").textContent=region==="ALL"?"All regions combined":region;

  let trend=[];
  years.forEach(y=>trend.push({year:y,value:filtered.filter(d=>d.Year===y).reduce((s,d)=>s+d.Population,0)}));
  drawTrend(trend);

  const ranking=latestRows.filter(d=>d.Region!=="Philippines").sort((a,b)=>b.Population-a.Population).slice(0,8);document.getElementById("ranking-year").textContent=`Top regions · ${latestYear}`;
  const max=ranking[0]?.Population||1;document.getElementById("ranking-list").innerHTML=ranking.map((d,i)=>`<div class="rank-row"><span class="rank-num">${String(i+1).padStart(2,"0")}</span><span class="rank-name">${escapeHtml(d.Region)}</span><span class="rank-bar"><i style="width:${d.Population/max*100}%"></i></span></div>`).join("");

  const regionalLatest=latestRows.filter(d=>d.Region!=="Philippines").sort((a,b)=>b.Population-a.Population).slice(0,10);
  const growthRows=regionalLatest.map(d=>{const prior=filtered.find(x=>x.Region===d.Region&&x.Year===firstYear);return {Region:d.Region,Growth:prior?((d.Population-prior.Population)/prior.Population)*100:0}}).sort((a,b)=>b.Growth-a.Growth).slice(0,8);
  drawGrowth(growthRows);drawDistribution(regionalLatest);updateTable(filtered);
}

function drawTrend(data){const ctx=document.getElementById("trend-chart");charts.trend?.destroy();charts.trend=new Chart(ctx,{type:"line",data:{labels:data.map(d=>d.year),datasets:[{data:data.map(d=>d.value/1e6),borderColor:"#1b9a9a",backgroundColor:"rgba(27,154,154,.08)",fill:true,borderWidth:2,pointRadius:2,pointHoverRadius:4,tension:.35}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.y.toFixed(2)} million`}}},scales:{x:{grid:{display:false},ticks:{font:{size:9},color:"#929da6"}},y:{grid:{color:"#eef1f3"},ticks:{font:{size:9},color:"#929da6",callback:v=>`${v}M`}}}}});}
function drawGrowth(rows){const ctx=document.getElementById("growth-chart");charts.growth?.destroy();charts.growth=new Chart(ctx,{type:"bar",data:{labels:rows.map(d=>d.Region.replace("Region ","R.")),datasets:[{data:rows.map(d=>d.Growth),backgroundColor:"#8069cb",borderRadius:3,barPercentage:.65}]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.x.toFixed(1)}%`}}},scales:{x:{grid:{color:"#eef1f3"},ticks:{font:{size:8},callback:v=>`${v}%`}},y:{grid:{display:false},ticks:{font:{size:8}}}}}});}
function drawDistribution(rows){const ctx=document.getElementById("distribution-chart");charts.distribution?.destroy();charts.distribution=new Chart(ctx,{type:"bar",data:{labels:rows.map(d=>d.Region.replace("National Capital Region","NCR").replace("Region ","R.")),datasets:[{data:rows.map(d=>d.Population/1e6),backgroundColor:"#4a78dc",borderRadius:3,barPercentage:.7}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.y.toFixed(2)} million`}}},scales:{x:{grid:{display:false},ticks:{font:{size:7},maxRotation:45,minRotation:45}},y:{grid:{color:"#eef1f3"},ticks:{font:{size:8},callback:v=>`${v}M`}}}}});}

function bindNavigation(){document.querySelectorAll(".nav-item").forEach(btn=>btn.addEventListener("click",()=>{document.querySelectorAll(".nav-item").forEach(b=>b.classList.remove("active"));btn.classList.add("active");document.querySelectorAll(".view").forEach(v=>v.classList.remove("active"));document.getElementById(`${btn.dataset.view}-view`).classList.add("active");document.getElementById("page-title").textContent={overview:"Population Overview",regions:"Regional Profiles",models:"Machine Learning Models",data:"Data Explorer"}[btn.dataset.view];if(btn.dataset.view==="regions")renderRegions();}));document.getElementById("apply-filter").addEventListener("click",updateDashboard);}
function bindUpload(){document.getElementById("file-upload").addEventListener("change",async e=>{const file=e.target.files[0];if(!file)return;document.getElementById("loading").classList.remove("hidden");if(file.name.toLowerCase().endsWith(".csv")){const text=await file.text();rawData=normalizeRows(parseCSV(text));}else{const buffer=await file.arrayBuffer();const wb=XLSX.read(buffer,{type:"array"});const sheet=wb.SheetNames.find(n=>/annual/i.test(n))||wb.SheetNames[0];rawData=normalizeRows(XLSX.utils.sheet_to_json(wb.Sheets[sheet],{defval:""}));}initialize();document.getElementById("loading").classList.add("hidden");});}
function parseCSV(text){const lines=text.trim().split(/\r?\n/);const headers=lines.shift().split(",").map(h=>h.trim());return lines.map(line=>{const values=line.split(",");return Object.fromEntries(headers.map((h,i)=>[h,values[i]?.trim()]));});}
function bindDownload(){document.getElementById("download-csv").addEventListener("click",e=>{e.preventEvent?.();const filtered=getFiltered();const csv=["Year,Region,Population",...filtered.map(d=>`${d.Year},"${d.Region.replaceAll('"','""')}",${d.Population}`)].join("\n");const blob=new Blob([csv],{type:"text/csv"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="population_filtered.csv";a.click();URL.revokeObjectObjectURL(a.href);});}
function bindTable(){document.getElementById("table-search").addEventListener("input",updateTable);}
function updateTable(filtered){if(!filtered)filtered=getFiltered();const latest=Math.max(...filtered.map(d=>d.Year));const rows=filtered.filter(d=>d.Year===latest);const term=document.getElementById("table-search")?.value?.toLowerCase()||"";const visible=rows.filter(d=>d.Region.toLowerCase().includes(term)).sort((a,b)=>b.Population-a.Population);const total=visible.reduce((s,d)=>s+d.Population,0);document.getElementById("table-count").textContent=`${visible.length} records · ${latest}`;document.getElementById("data-table").innerHTML=visible.map(d=>`<tr><td>${d.Year}</td><td>${escapeHtml(d.Region)}</td><td>${fmt.format(d.Population)}</td><td>${total?(d.Population/total*100).toFixed(2):"0"}%</td></tr>`).join("");}
function renderRegions(){const latest=Math.max(...rawData.map(d=>d.Year));const rows=rawData.filter(d=>d.Year===latest).filter(d=>d.Region!=="Philippines").sort((a,b)=>b.Population-a.Population);document.getElementById("regions-grid").innerHTML=rows.map(d=>{const all=rawData.filter(x=>x.Region===d.Region);const first=all[0],growth=first?((d.Population-first.Population)/first.Population*100):0;return `<article class="region-card"><p>${escapeHtml(d.Region)}</p><strong>${compact.format(d.Population)}</strong><small>${growth>=0?"+":""}${growth.toFixed(1)}% since ${first?.Year||latest}</small></article>`}).join("");}

loadDefaultData();
