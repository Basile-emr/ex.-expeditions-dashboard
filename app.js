// ===== Global state
let RAW_ROWS = [];
let charts = { nb:null, vol:null, dens:null, poidsMoy:null, topMat:null, topIPO:null, pkg:null, dimType:null };

// ===== Utils
const num = (v) => {
  if (v == null) return NaN;
  const s = String(v).replace(',', '.').replace(/\s/g,'');
  const n = parseFloat(s);
  return isNaN(n) ? NaN : n;
};
const fmt = (n, digits=0) => isNaN(n) ? '–' : n.toLocaleString('fr-FR', { maximumFractionDigits: digits, minimumFractionDigits: digits });

// normalize header names (lowercase, no accents)
const normalize = (s) => String(s || '').toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g,'') // strip accents
  .replace(/\n/g,' ').replace(/\s+/g,' ').trim();

// Find column by keywords (array of alternatives)
function findCol(headers, alts){
  const H = headers.map(h => normalize(h));
  for (let i=0;i<headers.length;i++){
    for (const k of alts){ if (H[i].includes(k)) return { name: headers[i], idx:i }; }
  }
  return null;
}

// ===== Parsing Excel
function detectDimColumns(headers){
  // Try canonical names first
  const getIdx = (keys)=>{ const H=headers.map(h=> (h||'').toString()); const n=H.map(x=>x.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')); for(let i=0;i<n.length;i++){ for(const k of keys){ if(n[i].includes(k)) return i; } } return -1; };
  let iL=getIdx(['longueur','length']);
  let iW=getIdx(['largeur','width']);
  let iH=getIdx(['hauteur','height']);
  if(iL===-1 || iW===-1 || iH===-1){
    // fallback: any three headers containing '(cm)' at the end or unit hints
    const idxs=[];
    for(let i=0;i<headers.length;i++){ const t=(headers[i]||'').toString().toLowerCase(); if(/cm\)/.test(t) || /cm/.test(t)) idxs.push(i); }
    if(idxs.length>=3){ iL=idxs[0]; iW=idxs[1]; iH=idxs[2]; }
  }
  return {iL,iW,iH};
}

function parseSheetToRows(ws){
  const json = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  if (json.length === 0) return [];
  const headers = json[0].map(h => h==null? '': String(h));

  const colWeek = findCol(headers, ['semaine']);
  const colDate = findCol(headers, ['date']);
  const colMat  = findCol(headers, ['materiel','matériel']);
  const colIPO  = findCol(headers, ['ipo','so']);
  const colPBrut= findCol(headers, ['poids brut', 'brut (kg)', 'brut reel', 'brut re']);
  const colVol  = findCol(headers, ['volume (m3)','volume m3','volume']);
  const colL    = findCol(headers, ['longueur']);
  const coll    = findCol(headers, ['largeur']);
  const colH    = findCol(headers, ['hauteur']);
  const colType = findCol(headers, ['colisage','emballage','type']);

  const rows = [];
  for (let r=1;r<json.length;r++){
    const row = json[r];
    const semaineRaw = colWeek ? row[colWeek.idx] : null;
    const sw = parseInt(semaineRaw,10);
    if (isNaN(sw)) continue; // ignore header-like rows
    const rec = {
      semaine: sw,
      date:    colDate ? row[colDate.idx] : null,
      materiel:colMat  ? row[colMat.idx]  : null,
      ipo:     colIPO  ? row[colIPO.idx]  : null,
      pbrut:   num(colPBrut ? row[colPBrut.idx] : null),
      vol:     num(colVol ? row[colVol.idx] : null),
      L:       num(colL ? row[colL.idx] : null),
      W:       num(coll ? row[coll.idx] : null),
      H:       num(colH ? row[colH.idx] : null),
      type:    colType ? String(row[colType.idx]||'').trim() : null
    };
    // derive volume from dimensions if vol missing and dims available (cm → m³)
    if (isNaN(rec.vol) && !isNaN(rec.L) && !isNaN(rec.W) && !isNaN(rec.H)){
      rec.vol = (rec.L * rec.W * rec.H) / 1_000_000; // cm3 → m3
    }
    // derive type if missing
    if (!rec.type){ rec.type = classifyType(rec); }
    rows.push(rec);
  }
  return rows;
}

// Heuristic type classification when 'type/colisage' column is absent
function classifyType(r){
  // Use volume thresholds if volume available, else dimensional hints
  const v = r && !isNaN(r.vol) ? r.vol : NaN;
  if (!isNaN(v)){
    if (v <= 0.3) return 'S (≤0,3 m³)';
    if (v <= 0.9) return 'M (≤0,9 m³)';
    if (v <= 2.0) return 'L (≤2,0 m³)';
    return 'XL (>2,0 m³)';
  }
  // fallback: pallet-like dimensions
  if ((r.L>=120 || r.W>=80)) return 'Palette-like';
  if ((r.L>=100 || r.W>=60)) return 'Caisse M';
  return 'Colis/Carton';
}

function loadWorkbook(file){
  const reader = new FileReader();
  reader.onload = (e)=>{
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, {type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]]; // first sheet
    RAW_ROWS = parseSheetToRows(ws);
    document.getElementById('fileName').textContent = file.name + ' — ' + RAW_ROWS.length + ' lignes';
    buildFilters();
    renderAll();
  };
  reader.readAsArrayBuffer(file);
}

// ===== Filters
function buildFilters(){
  const weeks = [...new Set(RAW_ROWS.map(r=>r.semaine))].sort((a,b)=>a-b);
  const ipos  = [...new Set(RAW_ROWS.map(r=>String(r.ipo||'').trim()).filter(Boolean))].sort();
  fillSelect('semaineSelect', weeks.map(v=>({value:v, label:'S'+v})));  
  fillSelect('ipoSelect', ipos.map(v=>({value:v, label:v})));
}
function fillSelect(id, options){
  const el = document.getElementById(id);
  el.innerHTML = '';
  for (const opt of options){
    const o = document.createElement('option');
    o.value = String(opt.value);
    o.textContent = opt.label;
    el.appendChild(o);
  }
  el.addEventListener('change', ()=>renderAll());
}

function getFiltered(){
  const sSel = [...document.getElementById('semaineSelect').selectedOptions].map(o=>parseInt(o.value,10));
  const iSel = [...document.getElementById('ipoSelect').selectedOptions].map(o=>o.value);
  return RAW_ROWS.filter(r=>
    (sSel.length? sSel.includes(r.semaine): true)
    && (iSel.length? iSel.includes(String(r.ipo||'')) : true)
  );
}

function resetAll(){
  for (const id of ['semaineSelect','ipoSelect']){
    const el = document.getElementById(id);
    for (const o of el.options) o.selected = false;
  }
  renderAll();
}

// ===== KPI & aggregates
function computeKPIs(rows){
  const nb = rows.length;
  const vol = rows.reduce((s,r)=> s + (isNaN(r.vol)?0:r.vol), 0);
  const pbrut = rows.reduce((s,r)=> s + (isNaN(r.pbrut)?0:r.pbrut), 0);
  const dens = vol>0 ? (pbrut/vol) : NaN;
  const poidsMoy = nb>0 ? (pbrut/nb) : NaN;
  const volMoy   = nb>0 ? (vol/nb)   : NaN;
  const score = computeScore(rows, { nb, vol, dens });
  return { nb, vol, pbrut, dens, poidsMoy, volMoy, score };
}

function computeScore(rows, base){
  // Score 0..100 construit sur 4 composantes simples
  // 1) Intensité volume vs meilleur semaine (30 pts)
  const byW = groupByWeek(rows);
  const maxVol = byW.reduce((m,r)=>Math.max(m, r.vol), 0);
  const volPart = maxVol>0 ? Math.min(base.vol / maxVol, 1) * 30 : 0;
  // 2) Cadence nb colis vs meilleure semaine (30 pts)
  const maxNb = byW.reduce((m,r)=>Math.max(m, r.nb), 0);
  const nbPart = maxNb>0 ? Math.min(base.nb / maxNb, 1) * 30 : 0;
  // 3) Densité cible ~500 kg/m3 (30 pts, décroît avec l'écart relatif)
  const dens = base.dens || 0;
  const densPart = dens>0 ? (1 - Math.min(Math.abs(dens - 500)/500, 1)) * 30 : 0; // cible 500
  // 4) Variabilité (10 pts) – plus homogène = mieux
  const poids = rows.map(r=>isNaN(r.pbrut)?null:r.pbrut).filter(v=>v!=null);
  let varPart = 0;
  if (poids.length>5){
    const m = poids.reduce((s,x)=>s+x,0)/poids.length; const sd = Math.sqrt(poids.reduce((s,x)=>s+(x-m)*(x-m),0)/poids.length);
    const cv = m>0 ? sd/m : 1; // coefficient de variation
    varPart = (1 - Math.min(cv,1)) * 10;
  }
  const total = Math.round(volPart + nbPart + densPart + varPart);
  // Note explicative
  const note = `Vol:${Math.round(volPart)}/30 · Nb:${Math.round(nbPart)}/30 · Dens:${Math.round(densPart)}/30 · Var:${Math.round(varPart)}/10`;
  return { value: total, note };
}

function groupByWeek(rows){
  const map = new Map();
  for (const r of rows){
    if (!map.has(r.semaine)) map.set(r.semaine, { nb:0, vol:0, pbrut:0 });
    const a = map.get(r.semaine);
    a.nb += 1; a.vol += (isNaN(r.vol)?0:r.vol); a.pbrut += (isNaN(r.pbrut)?0:r.pbrut);
  }
  return [...map.entries()].map(([w,v])=>({ semaine:w, ...v })).sort((a,b)=>a.semaine-b.semaine);
}

function groupByType(rows){
  const map = new Map();
  for (const r of rows){
    const t = String(r.type||'Inconnu');
    if (!map.has(t)) map.set(t, { nb:0, L:0, W:0, H:0, c:0 });
    const a = map.get(t);
    a.nb += 1; if(!isNaN(r.L)) a.L += r.L; if(!isNaN(r.W)) a.W += r.W; if(!isNaN(r.H)) a.H += r.H; a.c += 1;
  }
  const arr = [];
  for (const [t,v] of map){
    arr.push({ type:t, nb:v.nb, Lm:(v.c? v.L/v.c : NaN), Wm:(v.c? v.W/v.c : NaN), Hm:(v.c? v.H/v.c : NaN) });
  }
  // sort by nb desc
  arr.sort((a,b)=>b.nb-a.nb);
  return arr;
}

function topCount(rows, key, n=10){
  const map = new Map();
  for (const r of rows){
    const k = String(r[key]||'').trim();
    if (!k) continue;
    map.set(k, (map.get(k)||0) + 1);
  }
  return [...map.entries()].map(([k,v])=>({label:k, value:v})).sort((a,b)=>b.value-a.value).slice(0,n);
}

// ===== Rendering
function renderKPIs(rows){
  const k = computeKPIs(rows);
  document.getElementById('kpiNbColis').textContent = fmt(k.nb);
  document.getElementById('kpiVol').textContent = fmt(k.vol, 2);
  document.getElementById('kpiPBrut').textContent = fmt(k.pbrut, 0);
  document.getElementById('kpiDensite').textContent = fmt(k.dens, 1);
  document.getElementById('kpiPoidsMoy').textContent = fmt(k.poidsMoy, 0);
  document.getElementById('kpiVolMoy').textContent = fmt(k.volMoy, 3);
  document.getElementById('kpiScore').textContent = isNaN(k.score.value)? '–' : k.score.value;
  document.getElementById('scoreNote').textContent = k.score.note;
}

function renderCharts(rows){
  const weekly = groupByWeek(rows);
  const labels = weekly.map(r=>'S'+r.semaine);
  const nbData = weekly.map(r=>r.nb);
  const volData= weekly.map(r=>r.vol);
  const densData = weekly.map(r=> r.vol>0 ? r.pbrut/r.vol : NaN);
  const poidsMoyData = weekly.map(r=> r.nb>0 ? r.pbrut/r.nb : NaN);

  // Destroy previous
  for (const k of Object.keys(charts)){
    if (charts[k]) { charts[k].destroy(); charts[k]=null; }
  }

  charts.nb = new Chart(document.getElementById('chartNb'), {
    type:'bar', data:{ labels, datasets:[{ label:'Colis', data: nbData, backgroundColor:'#60a5fa' }]},
    options:{ plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  charts.vol = new Chart(document.getElementById('chartVol'), {
    type:'line', data:{ labels, datasets:[{ label:'Volume (m³)', data: volData, borderColor:'#2dd4bf', backgroundColor:'rgba(45,212,191,.2)', tension:.3, fill:true }]},
    options:{ plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  charts.dens = new Chart(document.getElementById('chartDens'), {
    type:'line', data:{ labels, datasets:[{ label:'Densité (kg/m³)', data: densData, borderColor:'#f472b6', backgroundColor:'rgba(244,114,182,.2)', tension:.3, fill:true }]},
    options:{ plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  charts.poidsMoy = new Chart(document.getElementById('chartPoidsMoy'), {
    type:'bar', data:{ labels, datasets:[{ label:'kg / colis', data: poidsMoyData, backgroundColor:'#f59e0b' }]},
    options:{ plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  const topM = topCount(rows, 'materiel');
  charts.topMat = new Chart(document.getElementById('chartTopMat'), {
    type:'bar', data:{ labels: topM.map(x=>x.label), datasets:[{ label:'Colis', data: topM.map(x=>x.value), backgroundColor:'#22c55e' }]},
    options:{ indexAxis:'y', plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  const topI = topCount(rows, 'ipo');
  charts.topIPO = new Chart(document.getElementById('chartTopIPO'), {
    type:'bar', data:{ labels: topI.map(x=>x.label), datasets:[{ label:'Colis', data: topI.map(x=>x.value), backgroundColor:'#a78bfa' }]},
    options:{ indexAxis:'y', plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  // Packaging repartition
  const byType = groupByType(rows).slice(0,12);
  charts.pkg = new Chart(document.getElementById('chartPkg'), {
    type:'doughnut',
    data:{ labels: byType.map(x=>x.type), datasets:[{ label:'Colis', data: byType.map(x=>x.nb), backgroundColor:['#60a5fa','#2dd4bf','#f59e0b','#a78bfa','#ef4444','#22c55e','#eab308','#06b6d4','#f97316','#84cc16','#10b981','#38bdf8'] }]},
    options:{ plugins:{ legend:{ position:'right', labels:{ color:'#e5e7eb' }}, datalabels:{ color:'#001219', backgroundColor:'#e5e7eb', borderRadius:4, padding:4, formatter:(v,ctx)=> ctx.chart.data.labels[ctx.dataIndex] + ' ('+v+')' } } },
    plugins: [ChartDataLabels]
  });

  // Dimensions by type (average L/W/H)
  const L = byType.map(x=>x.Lm); const W = byType.map(x=>x.Wm); const H = byType.map(x=>x.Hm);
  charts.dimType = new Chart(document.getElementById('chartDimByType'), {
    type:'bar',
    data:{ labels: byType.map(x=>x.type), datasets:[
      { label:'Longueur (cm)', data:L, backgroundColor:'#60a5fa' },
      { label:'Largeur (cm)', data:W, backgroundColor:'#f59e0b' },
      { label:'Hauteur (cm)', data:H, backgroundColor:'#a78bfa' }
    ]},
    options:{ indexAxis:'y', plugins:{ legend:{ labels:{ color:'#e5e7eb' } } }, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });
}

function renderAll(){
  const sSel = [...document.getElementById('semaineSelect').selectedOptions].map(o=>parseInt(o.value,10));
  const iSel = [...document.getElementById('ipoSelect').selectedOptions].map(o=>o.value);
  const rows = RAW_ROWS.filter(r=> (sSel.length? sSel.includes(r.semaine): true) && (iSel.length? iSel.includes(String(r.ipo||'')) : true));
  renderKPIs(rows);
  renderCharts(rows);
}

// ===== Export filtered as CSV
function exportCSV(){
  const sSel = [...document.getElementById('semaineSelect').selectedOptions].map(o=>parseInt(o.value,10));
  const iSel = [...document.getElementById('ipoSelect').selectedOptions].map(o=>o.value);
  const rows = RAW_ROWS.filter(r=> (sSel.length? sSel.includes(r.semaine): true) && (iSel.length? iSel.includes(String(r.ipo||'')) : true));
  if (!rows.length) return alert('Aucune donnée à exporter.');
  const headers = ['semaine','date','materiel','ipo','pbrut','vol','type','L','W','H'];
  let csv = headers.join(';') + '\n';
  for (const r of rows){
    csv += [r.semaine, r.date || '', (r.materiel||'').toString().replaceAll(';',','), (r.ipo||'').toString().replaceAll(';',','), r.pbrut||'', r.vol||'', r.type||'', r.L||'', r.W||'', r.H||''].join(';') + '\n';
  }
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'export_filtre.csv'; a.click(); URL.revokeObjectURL(a.href);
}

// ===== Drag & Drop & Boot
function setupDropzone(){
  const dz = document.getElementById('dropzone');
  dz.addEventListener('dragover', e=>{ e.preventDefault(); dz.classList.add('drag'); });
  dz.addEventListener('dragleave', ()=> dz.classList.remove('drag'));
  dz.addEventListener('drop', e=>{ e.preventDefault(); dz.classList.remove('drag'); const f = e.dataTransfer.files[0]; if (f) { document.getElementById('fileInput').files = e.dataTransfer.files; document.getElementById('fileName').textContent = f.name; loadWorkbook(f); } });
}

window.addEventListener('DOMContentLoaded', ()=>{
  document.getElementById('fileInput').addEventListener('change', (e)=>{ const f = e.target.files[0]; if (f) { document.getElementById('fileName').textContent = f.name; loadWorkbook(f); } });
  document.getElementById('resetBtn').addEventListener('click', ()=>{ for (const id of ['semaineSelect','ipoSelect']){ const el = document.getElementById(id); for (const o of el.options) o.selected = false; } renderAll(); });
  document.getElementById('downloadBtn').addEventListener('click', exportCSV);
  setupDropzone();
});
