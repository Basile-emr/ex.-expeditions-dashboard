// ===== Global state
let RAW_ROWS = [];
let charts = { nb:null, vol:null, dens:null, poidsMoy:null, dimCaisse:null };

// ===== Utils
const num = (v) => {
  if (v == null) return NaN;
  const s = String(v).replace(',', '.').replace(/\s/g,'');
  const n = parseFloat(s);
  return isNaN(n) ? NaN : n;
};
const fmt = (n, digits=0) => isNaN(n) ? '–' : n.toLocaleString('fr-FR', { maximumFractionDigits: digits, minimumFractionDigits: digits });
const normalize = (s) => String(s||'').toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
  .replace(/\n/g,' ').replace(/\s+/g,' ').trim();
function findCol(headers, alts){
  const H = headers.map(h => normalize(h));
  for (let i=0;i<headers.length;i++){
    for (const k of alts){ if (H[i].includes(k)) return { name: headers[i], idx:i }; }
  }
  return null;
}

// ===== Parsing Excel
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
  const colW    = findCol(headers, ['largeur']);
  const colH    = findCol(headers, ['hauteur']);

  const rows = [];
  for (let r=1;r<json.length;r++){
    const row = json[r];
    const semaineRaw = colWeek ? row[colWeek.idx] : null;
    const sw = parseInt(semaineRaw,10);
    if (isNaN(sw)) continue; // ignore header-like rows
    const rec = {
      semaine: sw,
      date: colDate ? row[colDate.idx] : null,
      materiel: colMat ? row[colMat.idx] : null,
      ipo: colIPO ? row[colIPO.idx] : null,
      pbrut: num(colPBrut ? row[colPBrut.idx] : null),
      vol: num(colVol ? row[colVol.idx] : null),
      L: num(colL ? row[colL.idx] : null),
      W: num(colW ? row[colW.idx] : null),
      H: num(colH ? row[colH.idx] : null),
      dimKey: null
    };

    // derive volume from dimensions if vol missing and dims available (cm → m³)
    if (isNaN(rec.vol) && !isNaN(rec.L) && !isNaN(rec.W) && !isNaN(rec.H)){
      rec.vol = (rec.L * rec.W * rec.H) / 1_000_000;
    }
    // dimension key as L×W×H (rounded to cm)
    if (!isNaN(rec.L) && !isNaN(rec.W) && !isNaN(rec.H)){
      const Lr = Math.round(rec.L), Wr = Math.round(rec.W), Hr = Math.round(rec.H);
      rec.dimKey = `${Lr}×${Wr}×${Hr} cm`;
    }
    rows.push(rec);
  }
  return rows;
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
  const mats  = [...new Set(RAW_ROWS.map(r=>String(r.materiel||'').trim()).filter(Boolean))].sort();
  const ipos  = [...new Set(RAW_ROWS.map(r=>String(r.ipo||'').trim()).filter(Boolean))].sort();
  fillSelect('semaineSelect', weeks.map(v=>({value:v, label:'S'+v})));
  fillSelect('materielSelect', mats.map(v=>({value:v, label:v})));
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
  const mSel = [...document.getElementById('materielSelect').selectedOptions].map(o=>o.value);
  const iSel = [...document.getElementById('ipoSelect').selectedOptions].map(o=>o.value);
  return RAW_ROWS.filter(r=>
    (sSel.length? sSel.includes(r.semaine): true)
    && (mSel.length? mSel.includes(String(r.materiel||'')) : true)
    && (iSel.length? iSel.includes(String(r.ipo||'')) : true)
  );
}
function resetAll(){
  for (const id of ['semaineSelect','materielSelect','ipoSelect']){
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
  const volMoy = nb>0 ? (vol/nb) : NaN;
  const score = computeScore(rows, { nb, vol, dens });
  return { nb, vol, pbrut, dens, poidsMoy, volMoy, score };
}
function computeScore(rows, base){
  const byW = groupByWeek(rows);
  const maxVol = byW.reduce((m,r)=>Math.max(m, r.vol), 0);
  const volPart = maxVol>0 ? Math.min(base.vol / maxVol, 1) * 30 : 0;
  const maxNb = byW.reduce((m,r)=>Math.max(m, r.nb), 0);
  const nbPart = maxNb>0 ? Math.min(base.nb / maxNb, 1) * 30 : 0;
  const dens = base.dens || 0;
  const densPart = dens>0 ? (1 - Math.min(Math.abs(dens - 500)/500, 1)) * 30 : 0; // cible 500
  const poids = rows.map(r=>isNaN(r.pbrut)?null:r.pbrut).filter(v=>v!=null);
  let varPart = 0;
  if (poids.length>5){
    const m = poids.reduce((s,x)=>s+x,0)/poids.length; const sd = Math.sqrt(poids.reduce((s,x)=>s+(x-m)*(x-m),0)/poids.length);
    const cv = m>0 ? sd/m : 1; // coefficient de variation
    varPart = (1 - Math.min(cv,1)) * 10;
  }
  const total = Math.round(volPart + nbPart + densPart + varPart);
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
function groupByDim(rows){
  const map = new Map();
  for (const r of rows){
    const k = r.dimKey || '';
    if (!k) continue;
    map.set(k, (map.get(k) || 0) + 1);
  }
  return [...map.entries()].map(([label, value]) => ({label, value})).sort((a,b)=> b.value - a.value);
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
    type:'bar', data:{ labels, datasets:[{ label:'Caisses', data: nbData, backgroundColor:'#60a5fa' }]},
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
    type:'bar', data:{ labels, datasets:[{ label:'kg / caisse', data: poidsMoyData, backgroundColor:'#f59e0b' }]},
    options:{ plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  // Répartition des caisses par dimensions (top 12)
  const dimTop = groupByDim(rows).slice(0, 12);
  charts.dimCaisse = new Chart(document.getElementById('chartDimCaisse'), {
    type:'bar',
    data:{ labels: dimTop.map(x=>x.label), datasets:[{ label:'Caisses', data: dimTop.map(x=>x.value), backgroundColor:'#60a5fa' }] },
    options:{ indexAxis:'y', plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });
}
function renderAll(){
  const rows = getFiltered();
  renderKPIs(rows);
  renderCharts(rows);
}

// ===== Export filtered as CSV
function exportCSV(){
  const sSel = [...document.getElementById('semaineSelect').selectedOptions].map(o=>parseInt(o.value,10));
  const mSel = [...document.getElementById('materielSelect').selectedOptions].map(o=>o.value);
  const iSel = [...document.getElementById('ipoSelect').selectedOptions].map(o=>o.value);
  const rows = RAW_ROWS.filter(r=> (sSel.length? sSel.includes(r.semaine): true) && (mSel.length? mSel.includes(String(r.materiel||'')) : true) && (iSel.length? iSel.includes(String(r.ipo||'')) : true));
  if (!rows.length) return alert('Aucune donnée à exporter.');
  const headers = ['semaine','date','materiel','ipo','dim_caisse','pbrut','vol','L','W','H'];
  let csv = headers.join(';') + '\n';
  for (const r of rows){
    csv += [
      r.semaine,
      r.date || '',
      String(r.materiel||'').replaceAll(';',','),
      String(r.ipo||'').replaceAll(';',','),
      r.dimKey || '',
      isNaN(r.pbrut)?'':r.pbrut,
      isNaN(r.vol)?'':r.vol,
      isNaN(r.L)?'':r.L,
      isNaN(r.W)?'':r.W,
      isNaN(r.H)?'':r.H
    ].join(';') + '\n';
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
  document.getElementById('resetBtn').addEventListener('click', resetAll);
  document.getElementById('downloadBtn').addEventListener('click', exportCSV);
  setupDropzone();
});
