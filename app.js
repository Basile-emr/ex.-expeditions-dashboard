// ===== STATE & GLOBALS =====
let RAW_ROWS = [];
let CURRENT_HEADERS = [];
let charts = {};
let filteredRows = [];
let archives = JSON.parse(localStorage.getItem('archives')) || [];

// ===== TAB NAVIGATION =====
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        const tabName = btn.dataset.tab;
        showTab(tabName);
    });
});

function showTab(tabName) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    
    document.getElementById(`tab-${tabName}`).classList.add('active');
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

    if (tabName === 'archives') renderArchives();
    if (tabName === 'data') renderDataTable();
    if (tabName === 'dashboard') updateDashboard();
}

// ===== UPLOAD TAB =====
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');

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
    if (files.length > 0) loadExcel(files[0]);
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) loadExcel(e.target.files[0]);
});

function loadExcel(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        parseExcel(data);
        showUploadStatus(file.name);
    };
    reader.readAsArrayBuffer(file);
}

function parseExcel(data) {
    if (data.length < 4) return;

    const headerLine1 = data[0].map(h => String(h || '').toLowerCase().trim());
    const headerLine2 = data[1].map(h => String(h || '').toLowerCase().trim());
    
    const headers = headerLine1.map((h, idx) => {
        if (!h && headerLine2[idx]) return headerLine2[idx];
        return h;
    });

    CURRENT_HEADERS = headers;
    const normalizedHeaders = headers.map(h => normalizeHeader(h));

    RAW_ROWS = [];

    for (let i = 3; i < data.length; i++) {
        const row = data[i];
        
        const semaineRaw = row[0];
        const semaine = parseInt(semaineRaw);
        if (isNaN(semaine)) continue;

        const obj = {};
        obj.semaine = semaine;

        for (let j = 0; j < normalizedHeaders.length; j++) {
            const hdr = normalizedHeaders[j];
            if (hdr && hdr !== '') {
                obj[hdr] = row[j] !== undefined ? row[j] : '';
            }
        }
        // Garder aussi les données brutes originales
        for (let j = 0; j < row.length; j++) {
            if (!obj[`col_${j}`]) obj[`col_${j}`] = row[j] || '';
        }

        obj.date = obj.date || '';
        obj.ipo = obj.ipo || '';

        let pbrut = parseFloat(String(obj.pbrut || '').replace(',', '.'));
        obj.pbrut = isNaN(pbrut) ? 0 : pbrut;

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

        if (vol === null && L !== null && W !== null && H !== null) {
            vol = (L * W * H) / 1000000;
        }

        obj.vol = isNaN(vol) ? 0 : vol;

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

    populateDashboardFilters();
    showUploadStatus('Fichier chargé avec succès');
}

function normalizeHeader(h) {
    if (!h) return '';
    const s = String(h).toLowerCase().trim();
    
    if (s.includes('semaine')) return 'semaine';
    if (s.includes('date') && s.includes('réception')) return 'date';
    if (s.includes('ipo') || s.includes('so')) return 'ipo';
    if (s.includes('poids') && (s.includes('brut') || s.includes('réel') || s.includes('reel'))) return 'pbrut';
    if (s.includes('volume') && (s.includes('m3') || s.includes('m³'))) return 'vol';
    
    return '';
}

function showUploadStatus(message) {
    const status = document.getElementById('uploadStatus');
    const msg = document.getElementById('uploadMessage');
    const archiveBtn = document.getElementById('archiveBtn');
    
    msg.textContent = `✅ ${message} (${RAW_ROWS.length} lignes chargées)`;
    status.style.display = 'block';
    archiveBtn.style.display = 'inline-block';
    
    archiveBtn.onclick = () => archiveCurrentFile(message);
}

// ===== ARCHIVES TAB =====
document.getElementById('archiveBtn')?.addEventListener('click', archiveCurrentFile);

function archiveCurrentFile(fileName) {
    if (RAW_ROWS.length === 0) {
        alert('Aucune donnée à archiver');
        return;
    }

    const archive = {
        id: Date.now(),
        name: fileName || `Archive-${new Date().toLocaleDateString('fr-FR')}`,
        date: new Date().toLocaleString('fr-FR'),
        rows: RAW_ROWS,
        headers: CURRENT_HEADERS
    };

    archives.push(archive);
    localStorage.setItem('archives', JSON.stringify(archives));
    alert('✅ Fichier archivé avec succès');
    renderArchives();
}

function renderArchives() {
    const list = document.getElementById('archivesList');
    
    if (archives.length === 0) {
        list.innerHTML = '<p class="empty-message">Aucune archive pour le moment</p>';
        return;
    }

    list.innerHTML = archives.map(archive => `
        <div class="archive-card">
            <div class="archive-name">${archive.name}</div>
            <div class="archive-meta">📅 ${archive.date}</div>
            <div class="archive-meta">📊 ${archive.rows.length} lignes</div>
            <div class="archive-actions">
                <button class="btn btn-secondary btn-small" onclick="restoreArchive(${archive.id})">📂 Restaurer</button>
                <button class="btn btn-secondary btn-small" onclick="exportArchiveCSV(${archive.id})">💾 CSV</button>
                <button class="btn btn-danger btn-small" onclick="deleteArchive(${archive.id})">🗑️ Supprimer</button>
            </div>
        </div>
    `).join('');
}

function restoreArchive(id) {
    const archive = archives.find(a => a.id === id);
    if (!archive) return;
    
    RAW_ROWS = JSON.parse(JSON.stringify(archive.rows));
    CURRENT_HEADERS = archive.headers;
    populateDashboardFilters();
    showTab('data');
    alert(`✅ Archive "${archive.name}" restaurée`);
}

function deleteArchive(id) {
    if (!confirm('Confirmer la suppression ?')) return;
    archives = archives.filter(a => a.id !== id);
    localStorage.setItem('archives', JSON.stringify(archives));
    renderArchives();
}

function exportArchiveCSV(id) {
    const archive = archives.find(a => a.id === id);
    if (!archive) return;
    
    let csv = 'semaine;date;ipo;dim_caisse;pbrut;vol;L;W;H\n';
    archive.rows.forEach(r => {
        const dimCaisse = r.dimKey || '';
        const L = r.L !== null ? r.L : '';
        const W = r.W !== null ? r.W : '';
        const H = r.H !== null ? r.H : '';
        csv += `${r.semaine};${r.date};${r.ipo};${dimCaisse};${r.pbrut};${r.vol};${L};${W};${H}\n`;
    });

    downloadCSV(csv, `${archive.name}.csv`);
}

// ===== DATA TABLE TAB =====
document.getElementById('addRowBtn')?.addEventListener('click', addDataRow);
document.getElementById('deleteRowBtn')?.addEventListener('click', deleteSelectedRows);
document.getElementById('exportDataBtn')?.addEventListener('click', exportDataTableCSV);

function renderDataTable() {
    const container = document.getElementById('dataTableContainer');
    
    if (RAW_ROWS.length === 0) {
        container.innerHTML = '<p class="empty-message">Chargez un fichier d\'abord</p>';
        return;
    }

    const allKeys = [...new Set(RAW_ROWS.flatMap(r => Object.keys(r)))];
    const displayKeys = ['semaine', 'date', 'ipo', 'pbrut', 'vol', 'L', 'W', 'H', 'dimKey'].filter(k => allKeys.includes(k));

    let html = '<table><thead><tr><th><input type="checkbox" id="selectAll"></th>';
    displayKeys.forEach(key => {
        html += `<th>${key}</th>`;
    });
    html += '</tr></thead><tbody>';

    RAW_ROWS.forEach((row, idx) => {
        html += `<tr><td><input type="checkbox" class="row-checkbox" data-idx="${idx}"></td>`;
        displayKeys.forEach(key => {
            const val = row[key] !== null && row[key] !== undefined ? row[key] : '';
            const inputType = key === 'date' ? 'date' : (key === 'pbrut' || key === 'vol' || key === 'L' || key === 'W' || key === 'H' ? 'number' : 'text');
            html += `<td><input type="${inputType}" class="cell-input" data-idx="${idx}" data-key="${key}" value="${val}" onchange="updateCell(${idx}, '${key}', this.value)"></td>`;
        });
        html += `<td><button class="btn btn-danger btn-small" onclick="deleteRow(${idx})">❌</button></td></tr>`;
    });

    html += '</tbody></table>';
    container.innerHTML = html;

    document.getElementById('selectAll').addEventListener('change', function() {
        document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = this.checked);
    });
}

function updateCell(idx, key, value) {
    if (key === 'semaine' || key === 'pbrut' || key === 'vol' || key === 'L' || key === 'W' || key === 'H') {
        RAW_ROWS[idx][key] = isNaN(parseFloat(value)) ? 0 : parseFloat(value);
    } else {
        RAW_ROWS[idx][key] = value;
    }
    updateDashboard();
}

function addDataRow() {
    const newRow = {
        semaine: 1,
        date: '',
        ipo: '',
        pbrut: 0,
        vol: 0,
        L: null,
        W: null,
        H: null,
        dimKey: ''
    };
    RAW_ROWS.push(newRow);
    renderDataTable();
}

function deleteSelectedRows() {
    const selected = Array.from(document.querySelectorAll('.row-checkbox:checked')).map(cb => parseInt(cb.dataset.idx));
    if (selected.length === 0) {
        alert('Sélectionnez des lignes');
        return;
    }
    RAW_ROWS = RAW_ROWS.filter((_, idx) => !selected.includes(idx));
    renderDataTable();
}

function deleteRow(idx) {
    RAW_ROWS.splice(idx, 1);
    renderDataTable();
}

function exportDataTableCSV() {
    let csv = 'semaine;date;ipo;dim_caisse;pbrut;vol;L;W;H\n';
    RAW_ROWS.forEach(r => {
        const dimCaisse = r.dimKey || '';
        const L = r.L !== null ? r.L : '';
        const W = r.W !== null ? r.W : '';
        const H = r.H !== null ? r.H : '';
        csv += `${r.semaine};${r.date};${r.ipo};${dimCaisse};${r.pbrut};${r.vol};${L};${W};${H}\n`;
    });
    downloadCSV(csv, 'donnees_editees.csv');
}

// ===== DASHBOARD TAB =====
const dashSemSelect = document.getElementById('dashSemSelect');
const dashIpoSelect = document.getElementById('dashIpoSelect');
document.getElementById('dashResetBtn')?.addEventListener('click', resetDashFilters);
document.getElementById('dashExportBtn')?.addEventListener('click', exportDashboardCSV);

dashSemSelect?.addEventListener('change', updateDashboard);
dashIpoSelect?.addEventListener('change', updateDashboard);

function populateDashboardFilters() {
    const semaines = [...new Set(RAW_ROWS.map(r => r.semaine))].sort((a, b) => a - b);
    const ipos = [...new Set(RAW_ROWS.map(r => r.ipo).filter(x => x))].sort();

    dashSemSelect.innerHTML = '';
    semaines.forEach(s => {
        const opt = document.createElement('option');
        opt.value = s;
        opt.textContent = `Semaine ${s}`;
        dashSemSelect.appendChild(opt);
    });

    dashIpoSelect.innerHTML = '';
    ipos.forEach(ip => {
        const opt = document.createElement('option');
        opt.value = ip;
        opt.textContent = ip;
        dashIpoSelect.appendChild(opt);
    });

    updateDashboard();
}

function resetDashFilters() {
    Array.from(dashSemSelect.options).forEach(opt => (opt.selected = false));
    Array.from(dashIpoSelect.options).forEach(opt => (opt.selected = false));
    updateDashboard();
}

function updateDashboard() {
    const selectedSemaines = Array.from(dashSemSelect.selectedOptions).map(o => parseInt(o.value));
    const selectedIpos = Array.from(dashIpoSelect.selectedOptions).map(o => o.value);

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
    if (rows.length === 0) return { total: '–', note: '' };

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
        const ratio = byWeek[w].vol / maxVol;
        scoreVol = Math.max(scoreVol, Math.min(30, 30 * ratio));
    });

    weeks.forEach(w => {
        const ratio = byWeek[w].nb / maxNb;
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
            byWeek[r.semaine] = { nb: 0, vol: 0, pbrut: 0 };
        }
        byWeek[r.semaine].nb += 1;
        byWeek[r.semaine].vol += r.vol || 0;
        byWeek[r.semaine].pbrut += r.pbrut || 0;
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
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { ticks: { color: '#e0e0e0' }, grid: { color: '#404040' } },
                x: { ticks: { color: '#e0e0e0' }, grid: { color: '#404040' } }
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
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { ticks: { color: '#e0e0e0' }, grid: { color: '#404040' } },
                x: { ticks: { color: '#e0e0e0' }, grid: { color: '#404040' } }
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
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { ticks: { color: '#e0e0e0' }, grid: { color: '#404040' } },
                x: { ticks: { color: '#e0e0e0' }, grid: { color: '#404040' } }
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
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { ticks: { color: '#e0e0e0' }, grid: { color: '#404040' } },
                x: { ticks: { color: '#e0e0e0' }, grid: { color: '#404040' } }
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
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { ticks: { color: '#e0e0e0', font: { size: 10 } }, grid: { color: '#404040' } },
                x: { ticks: { color: '#e0e0e0' }, grid: { color: '#404040' } }
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

function exportDashboardCSV() {
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

    downloadCSV(csv, 'export_caisses_filtré.csv');
}

function downloadCSV(csv, filename) {
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
