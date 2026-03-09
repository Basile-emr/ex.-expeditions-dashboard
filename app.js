// filepath: app.js
let RAW_ROWS = [];
let charts = {};
let filteredRows = [];

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const semaineSelect = document.getElementById('semaineSelect');
const ipoSelect = document.getElementById('ipoSelect');
const resetBtn = document.getElementById('resetBtn');
const downloadBtn = document.getElementById('downloadBtn');

dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.style.backgroundColor = '#2e5f99';
});

dropzone.addEventListener('dragleave', () => {
    dropzone.style.backgroundColor = '#2d2d2d';
});

dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.style.backgroundColor = '#2d2d2d';
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        loadExcel(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        loadExcel(e.target.files[0]);
    }
});

resetBtn.addEventListener('click', () => {
    Array.from(semaineSelect.options).forEach(opt => (opt.selected = false));
    Array.from(ipoSelect.options).forEach(opt => (opt.selected = false));
    applyFilters();
});

downloadBtn.addEventListener('click', exportCSV);

function loadExcel(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        parseExcel(data);
    };
    reader.readAsArrayBuffer(file);
}

function parseExcel(data) {
    if (data.length < 2) return;

    const headers = data[0];
    
    // Normaliser les en-têtes
    const normalizedHeaders = headers.map(h => normalizeHeader(h));

    RAW_ROWS = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const obj = {};

        // Mapper les en-têtes normalisés
        for (let j = 0; j < normalizedHeaders.length; j++) {
            const hdr = normalizedHeaders[j];
            if (hdr && hdr !== '' && hdr !== 'unknown') {
                obj[hdr] = row[j] !== undefined ? row[j] : '';
            }
        }

        const semaine = parseInt(obj.semaine);
        if (isNaN(semaine)) continue;

        obj.semaine = semaine;
        obj.date = obj.date || '';
        obj.ipo = obj.ipo || '';

        let pbrut = parseFloat(String(obj.pbrut || '').replace(',', '.'));
        obj.pbrut = isNaN(pbrut) ? 0 : pbrut;

        // Extraction des dimensions directement des colonnes F, G, H (indices 5, 6, 7)
        let L = null, W = null, H = null;

        if (row.length > 5) {
            L = parseFloat(String(row[5] || '').replace(',', '.'));
            if (isNaN(L)) L = null;
        }
        if (row.length > 6) {
            W = parseFloat(String(row[6] || '').replace(',', '.'));
            if (isNaN(W)) W = null;
        }
        if (row.length > 7) {
            H = parseFloat(String(row[7] || '').replace(',', '.'));
            if (isNaN(H)) H = null;
        }

        obj.L = L;
        obj.W = W;
        obj.H = H;

        let vol = parseFloat(String(obj.vol || '').replace(',', '.'));
        if (isNaN(vol)) vol = null;

        // Calcul du volume si absent mais dimensions présentes
        if (vol === null && L !== null && W !== null && H !== null) {
            vol = (L * W * H) / 1000000;
        }

        obj.vol = isNaN(vol) ? 0 : vol;

        // Construction de la clé dimension
        if (obj.L !== null && obj.W !== null && obj.H !== null) {
            const lRound = Math.round(obj.L);
            const wRound = Math.round(obj.W);
            const hRound = Math.round(obj.H);
            obj.dimKey = `${lRound}×${wRound}×${hRound} cm`;
        } else {
            obj.dimKey = '';
        }

        RAW_ROWS.push(obj);
    }

    populateFilters();
    applyFilters();
}

function normalizeHeader(h) {
    if (!h) return '';
    const s = String(h).toLowerCase().trim();
    
    if (s.includes('semaine')) return 'semaine';
    if (s.includes('date') && (s.includes('réception') || s.includes('reception'))) return 'date';
    if (s.includes('ipo') || s.includes('so')) return 'ipo';
    if (s.includes('poids') && (s.includes('brut') || s.includes('réel') || s.includes('reel'))) return 'pbrut';
    if (s.includes('volume') && (s.includes('m3') || s.includes('m³'))) return 'vol';
    
    return '';
}

function populateFilters() {
    const semaines = [...new Set(RAW_ROWS.map(r => r.semaine))].sort((a, b) => a - b);
    const ipos = [...new Set(RAW_ROWS.map(r => r.ipo).filter(x => x))].sort();

    semaineSelect.innerHTML = '';
    semaines.forEach(s => {
        const opt = document.createElement('option');
        opt.value = s;
        opt.textContent = `Semaine ${s}`;
        semaineSelect.appendChild(opt);
    });

    ipoSelect.innerHTML = '';
    ipos.forEach(ip => {
        const opt = document.createElement('option');
        opt.value = ip;
        opt.textContent = ip;
        ipoSelect.appendChild(opt);
    });
}

function applyFilters() {
    const selectedSemaines = Array.from(semaineSelect.selectedOptions).map(o => parseInt(o.value));
    const selectedIpos = Array.from(ipoSelect.selectedOptions).map(o => o.value);

    filteredRows = RAW_ROWS.filter(r => {
        const semMatch = selectedSemaines.length === 0 || selectedSemaines.includes(r.semaine);
        const ipoMatch = selectedIpos.length === 0 || selectedIpos.includes(r.ipo);
        return semMatch && ipoMatch;
    });

    updateKPIs();
    updateCharts();
}

function updateKPIs() {
    const nb = filteredRows.length;
    const totalVol = filteredRows.reduce((sum, r) => sum + (r.vol || 0), 0);
    const totalPbrut = filteredRows.reduce((sum, r) => sum + (r.pbrut || 0), 0);
    const dens = totalVol > 0 ? totalPbrut / totalVol : 0;
    const poidsMoy = nb > 0 ? totalPbrut / nb : 0;
    const volMoy = nb > 0 ? totalVol / nb : 0;

    document.getElementById('kpiNbCaisses').textContent = nb > 0 ? nb : '–';
    document.getElementById('kpiVol').textContent = nb > 0 ? totalVol.toFixed(2) : '–';
    document.getElementById('kpiPBrut').textContent = nb > 0 ? totalPbrut.toFixed(2) : '–';
    document.getElementById('kpiDensite').textContent = nb > 0 ? dens.toFixed(2) : '–';
    document.getElementById('kpiPoidsMoy').textContent = nb > 0 ? poidsMoy.toFixed(2) : '–';
    document.getElementById('kpiVolMoy').textContent = nb > 0 ? volMoy.toFixed(4) : '–';

    const score = calcScore(filteredRows);
    document.getElementById('kpiScore').textContent = score.total;
    document.getElementById('scoreNote').textContent = score.note;
}

function calcScore(rows) {
    if (rows.length === 0) {
        return { total: '–', note: '' };
    }

    const byWeek = {};
    rows.forEach(r => {
        if (!byWeek[r.semaine]) {
            byWeek[r.semaine] = { nb: 0, vol: 0, pbrut: 0 };
        }
        byWeek[r.semaine].nb += 1;
        byWeek[r.semaine].vol += r.vol || 0;
        byWeek[r.semaine].pbrut += r.pbrut || 0;
    });

    const weeks = Object.keys(byWeek).map(k => parseInt(k)).sort();
    const maxNb = Math.max(...weeks.map(w => byWeek[w].nb));
    const maxVol = Math.max(...weeks.map(w => byWeek[w].vol));

    let scoreVol = 0, scoreNb = 0, scoreDens = 0, scoreVar = 0;

    weeks.forEach(w => {
        const weekData = byWeek[w];
        const ratio = weekData.vol / maxVol;
        scoreVol = Math.max(scoreVol, Math.min(30, 30 * ratio));
    });

    weeks.forEach(w => {
        const weekData = byWeek[w];
        const ratio = weekData.nb / maxNb;
        scoreNb = Math.max(scoreNb, Math.min(30, 30 * ratio));
    });

    const totalVol = rows.reduce((s, r) => s + (r.vol || 0), 0);
    const totalPbrut = rows.reduce((s, r) => s + (r.pbrut || 0), 0);
    const densCible = 500;
    const densActuelle = totalVol > 0 ? totalPbrut / totalVol : 0;
    const ecart = densActuelle > 0 ? Math.abs(densActuelle - densCible) / densCible : 1;
    scoreDens = Math.max(0, 30 * (1 - Math.min(ecart, 1)));

    const pbruts = rows.filter(r => r.pbrut > 0).map(r => r.pbrut);
    if (pbruts.length >= 6) {
        const mean = pbruts.reduce((s, v) => s + v, 0) / pbruts.length;
        const variance = pbruts.reduce((s, v) => s + Math.pow(v - mean, 2), 0) / pbruts.length;
        const stdDev = Math.sqrt(variance);
        const cv = stdDev / mean;
        scoreVar = Math.max(0, 10 * (1 - Math.min(cv, 1)));
    }

    const total = Math.round(scoreVol + scoreNb + scoreDens + scoreVar);
    const note = `Vol:${Math.round(scoreVol)}/30 · Nb:${Math.round(scoreNb)}/30 · Dens:${Math.round(scoreDens)}/30 · Var:${Math.round(scoreVar)}/10`;

    return { total, note };
}

function updateCharts() {
    const byWeek = {};
    filteredRows.forEach(r => {
        if (!byWeek[r.semaine]) {
            byWeek[r.semaine] = { nb: 0, vol: 0, pbrut: 0, pbruts: [] };
        }
        byWeek[r.semaine].nb += 1;
        byWeek[r.semaine].vol += r.vol || 0;
        byWeek[r.semaine].pbrut += r.pbrut || 0;
        if (r.pbrut) byWeek[r.semaine].pbruts.push(r.pbrut);
    });

    const weeks = Object.keys(byWeek).map(k => parseInt(k)).sort((a, b) => a - b);
    const labels = weeks.map(w => `S${w}`);
    const nbData = weeks.map(w => byWeek[w].nb);
    const volData = weeks.map(w => byWeek[w].vol);
    const densData = weeks.map(w => {
        const v = byWeek[w].vol;
        return v > 0 ? byWeek[w].pbrut / v : 0;
    });
    const poidsMoyData = weeks.map(w => {
        const nb = byWeek[w].nb;
        return nb > 0 ? byWeek[w].pbrut / nb : 0;
    });

    destroyCharts();

    const ctx1 = document.getElementById('chartNb').getContext('2d');
    charts.nb = new Chart(ctx1, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Nb caisses',
                data: nbData,
                backgroundColor: '#4a9eff',
                borderColor: '#2e5f99',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: { legend: { display: false } },
            scales: {
                y: {
                    ticks: { color: '#e0e0e0' },
                    grid: { color: '#404040' }
                },
                x: {
                    ticks: { color: '#e0e0e0' },
                    grid: { color: '#404040' }
                }
            }
        }
    });

    const ctx2 = document.getElementById('chartVol').getContext('2d');
    charts.vol = new Chart(ctx2, {
        type: 'line',
        data: {
            labels,
            datasets: [{
                label: 'Volume (m³)',
                data: volData,
                borderColor: '#4ade80',
                backgroundColor: 'rgba(74, 222, 128, 0.1)',
                borderWidth: 2,
                tension: 0.4,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: { legend: { display: false } },
            scales: {
                y: {
                    ticks: { color: '#e0e0e0' },
                    grid: { color: '#404040' }
                },
                x: {
                    ticks: { color: '#e0e0e0' },
                    grid: { color: '#404040' }
                }
            }
        }
    });

    const ctx3 = document.getElementById('chartDens').getContext('2d');
    charts.dens = new Chart(ctx3, {
        type: 'line',
        data: {
            labels,
            datasets: [{
                label: 'Densité (kg/m³)',
                data: densData,
                borderColor: '#fbbf24',
                backgroundColor: 'rgba(251, 191, 36, 0.1)',
                borderWidth: 2,
                tension: 0.4,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: { legend: { display: false } },
            scales: {
                y: {
                    ticks: { color: '#e0e0e0' },
                    grid: { color: '#404040' }
                },
                x: {
                    ticks: { color: '#e0e0e0' },
                    grid: { color: '#404040' }
                }
            }
        }
    });

    const ctx4 = document.getElementById('chartPoidsMoy').getContext('2d');
    charts.poidsMoy = new Chart(ctx4, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Poids moyen / caisse (kg)',
                data: poidsMoyData,
                backgroundColor: '#f87171',
                borderColor: '#dc2626',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: { legend: { display: false } },
            scales: {
                y: {
                    ticks: { color: '#e0e0e0' },
                    grid: { color: '#404040' }
                },
                x: {
                    ticks: { color: '#e0e0e0' },
                    grid: { color: '#404040' }
                }
            }
        }
    });

    const dimCount = {};
    filteredRows.forEach(r => {
        if (r.dimKey) {
            dimCount[r.dimKey] = (dimCount[r.dimKey] || 0) + 1;
        }
    });

    const topDims = Object.entries(dimCount)
        .sort(([, a], [, b]) => b - a)
        .slice(0, 12)
        .map(([k, v]) => ({ dim: k, count: v }));

    const dimLabels = topDims.map(d => d.dim);
    const dimDataCounts = topDims.map(d => d.count);

    const ctx5 = document.getElementById('chartDimCaisse').getContext('2d');
    charts.dimCaisse = new Chart(ctx5, {
        type: 'bar',
        data: {
            labels: dimLabels,
            datasets: [{
                label: 'Nb caisses',
                data: dimDataCounts,
                backgroundColor: '#8b5cf6',
                borderColor: '#6d28d9',
                borderWidth: 1
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: true,
            plugins: { legend: { display: false } },
            scales: {
                y: {
                    ticks: { color: '#e0e0e0', font: { size: 10 } },
                    grid: { color: '#404040' }
                },
                x: {
                    ticks: { color: '#e0e0e0' },
                    grid: { color: '#404040' }
                }
            }
        }
    });
}

function destroyCharts() {
    Object.values(charts).forEach(chart => {
        if (chart) chart.destroy();
    });
    charts = {};
}

function exportCSV() {
    if (filteredRows.length === 0) {
        alert('Aucune donnée à exporter');
        return;
    }

    let csv = 'semaine;date;ipo;dim_caisse;pbrut;vol;L;W;H\n';
    filteredRows.forEach(r => {
        const dimCaisse = r.dimKey || '';
        const L = r.L !== null ? r.L : '';
        const W = r.W !== null ? r.W : '';
        const H = r.H !== null ? r.H : '';
        csv += `${r.semaine};${r.date};${r.ipo};${dimCaisse};${r.pbrut};${r.vol};${L};${W};${H}\n`;
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', 'export_caisses.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

semaineSelect.addEventListener('change', applyFilters);
ipoSelect.addEventListener('change', applyFilters);
