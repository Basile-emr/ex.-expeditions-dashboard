// ===== Global state
let RAW_ROWS = [];
let FILTERS = { semaines: new Set(), materiels: new Set(), ipos: new Set() };
let charts = { nb:null, vol:null, topMat:null, topIPO:null };

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

function parseSheetToRows(ws){
  const json = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  if (json.length === 0) return [];
  const headers = json[0].map(h => h==null? '': String(h));

  // columns of interest (robust to variants)
  const colWeek = findCol(headers, ['semaine']);
  const colDate = findCol(headers, ['date']);
  const colMat  = findCol(headers, ['materiel','matériel']);
  const colIPO  = findCol(headers, ['ipo','so']);
  const colPBrut= findCol(headers, ['poids brut', 'brut (kg)', 'brut reel', 'brut re']);
  const colVol  = findCol(headers, ['volume (m3)','volume m3','volume']);

  // map rows
  const rows = [];
  for (let r=1;r<json.length;r++){
    const row = json[r];
    const rec = {
      semaine: colWeek ? row[colWeek.idx] : null,
      date:    colDate ? row[colDate.idx] : null,
      materiel:colMat  ? row[colMat.idx]  : null,
      ipo:     colIPO  ? row[colIPO.idx]  : null,
      pbrut:   num(colPBrut ? row[colPBrut.idx] : null),
      vol:     num(colVol ? row[colVol.idx] : null)
    };
    // keep only numeric/semaine rows
    const sw = parseInt(rec.semaine,10);
    if (!isNaN(sw)) { rec.semaine = sw; rows.push(rec); }
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

// ===== KPIs & Charts
function computeKPIs(rows){
  const nb = rows.length;
  const vol = rows.reduce((s,r)=> s + (isNaN(r.vol)?0:r.vol), 0);
  const pbrut = rows.reduce((s,r)=> s + (isNaN(r.pbrut)?0:r.pbrut), 0);
  const dens = vol>0 ? (pbrut/vol) : NaN;
  return { nb, vol, pbrut, dens };
}

function groupByWeek(rows){
  const map = new Map();
  for (const r of rows){
    if (!map.has(r.semaine)) map.set(r.semaine, { nb:0, vol:0, pbrut:0 });
    const a = map.get(r.semaine);
    a.nb += 1; a.vol += (isNaN(r.vol)?0:r.vol); a.pbrut += (isNaN(r.pbrut)?0:r.pbrut);
  }
  const arr = [...map.entries()].map(([w,v])=>({ semaine:w, ...v })).sort((a,b)=>a.semaine-b.semaine);
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

function renderKPIs(rows){
  const k = computeKPIs(rows);
  document.getElementById('kpiNbColis').textContent = fmt(k.nb);
  document.getElementById('kpiVol').textContent = fmt(k.vol, 2);
  document.getElementById('kpiPBrut').textContent = fmt(k.pbrut, 0);
  document.getElementById('kpiDensite').textContent = fmt(k.dens, 1);
}

function renderCharts(rows){
  const weekly = groupByWeek(rows);
  const labels = weekly.map(r=>'S'+r.semaine);
  const nbData = weekly.map(r=>r.nb);
  const volData= weekly.map(r=>r.vol);

  // Destroy previous to avoid overlay
  for (const k of Object.keys(charts)){
    if (charts[k]) { charts[k].destroy(); charts[k]=null; }
  }

  charts.nb = new Chart(document.getElementById('chartNb'), {
    type:'bar',
    data:{ labels, datasets:[{ label:'Colis', data: nbData, backgroundColor:'#60a5fa' }]},
    options:{ plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  charts.vol = new Chart(document.getElementById('chartVol'), {
    type:'line',
    data:{ labels, datasets:[{ label:'Volume (m³)', data: volData, borderColor:'#2dd4bf', backgroundColor:'rgba(45,212,191,.2)', tension:.3, fill:true }]},
    options:{ plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  const topM = topCount(rows, 'materiel');
  charts.topMat = new Chart(document.getElementById('chartTopMat'), {
    type:'bar',
    data:{ labels: topM.map(x=>x.label), datasets:[{ label:'Colis', data: topM.map(x=>x.value), backgroundColor:'#f59e0b' }]},
    options:{ indexAxis:'y', plugins:{ legend:{ display:false }}, scales:{ x:{ ticks:{ color:'#cbd5e1'} }, y:{ ticks:{ color:'#cbd5e1'} } } }
  });

  const topI = topCount(rows, 'ipo');
  charts.topIPO = new Chart(document.getElementById('chartTopIPO'), {
    type:'bar',
    data:{ labels: topI.map(x=>x.label), datasets:[{ label:'Colis', data: topI.map(x=>x.value), backgroundColor:'#a78bfa' }]},
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
  const rows = getFiltered();
  if (!rows.length) return alert('Aucune donnée à exporter.');
  const headers = ['semaine','date','materiel','ipo','pbrut','vol'];
  let csv = headers.join(';') + '\n';
  for (const r of rows){
    csv += [r.semaine, r.date || '', (r.materiel||'').toString().replaceAll(';',','), (r.ipo||'').toString().replaceAll(';',','), r.pbrut||'', r.vol||''].join(';') + '\n';
  }
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'export_filtre.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

// ===== Drag & Drop
function setupDropzone(){
  const dz = document.getElementById('dropzone');
  dz.addEventListener('dragover', e=>{ e.preventDefault(); dz.classList.add('drag'); });
  dz.addEventListener('dragleave', ()=> dz.classList.remove('drag'));
  dz.addEventListener('drop', e=>{
    e.preventDefault(); dz.classList.remove('drag');
    const f = e.dataTransfer.files[0];
    if (f) { document.getElementById('fileInput').files = e.dataTransfer.files; document.getElementById('fileName').textContent = f.name; loadWorkbook(f); }
  });
}

// ===== Boot
window.addEventListener('DOMContentLoaded', ()=>{
  document.getElementById('fileInput').addEventListener('change', (e)=>{
    const f = e.target.files[0];
    if (f) { document.getElementById('fileName').textContent = f.name; loadWorkbook(f); }
  });
  document.getElementById('resetBtn').addEventListener('click', resetAll);
  document.getElementById('downloadBtn').addEventListener('click', exportCSV);
  setupDropzone();
});
