// ===== ÉTAT GLOBAL =====
let CURRENT_FICHE = null;
let CURRENT_ROWS = [];
let FILTERED_ROWS = [];
let ARCHIVES = JSON.parse(localStorage.getItem('expArchives')) || [];
let charts = {};

// ===== NAVIGATION ONGLETS =====
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        const tabName = btn.dataset.tab;
        switchTab(tabName);
    });
});

function switchTab(tabName) {
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    
    document.getElementById(`tab-${tabName}`).classList.add('active');
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

    if (tabName === 'archives') refreshArchivesView();
    if (tabName === 'comparaison') setupComparisonTab();
}

// ===== UPLOAD TAB =====
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');

['dragover', 'dragenter'].forEach(e => {
    dropzone.addEventListener(e, (ev) => {
        ev.preventDefault();
        dropzone.style.borderColor = '#00cc44';
    });
});

['dragleave', 'drop'].forEach(e => {
    dropzone.addEventListener(e, () => {
        dropzone.style.borderColor = '#0099ff';
    });
});

dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    const files = e.dataTransfer.files;
    if (files.length > 0) loadFileToUpload(files[0]);
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) loadFileToUpload(e.target.files[0]);
});

function loadFileToUpload(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
        showUploadStatus(file, data);
    };
    reader.readAsArrayBuffer(file);
}

function showUploadStatus(file, data) {
    const status = document.getElementById('uploadStatus');
    const msg = document.getElementById('uploadMessage');
    const loadBtn = document.getElementById('loadFicheBtn');
    const archiveBtn = document.getElementById('archiveCurrentBtn');
    
    // Garder temporairement les données du fichier
    window.tempFileData = { name: file.name, data: data };
    
    msg.innerHTML = `✅ <strong>${file.name}</strong> chargé (${data.length - 4} lignes)`;
    status.style.display = 'block';
    loadBtn.style.display = 'inline-block';
    archiveBtn.style.display = 'inline-block';
    
    loadBtn.onclick = () => loadFicheFromUpload();
    archiveBtn.onclick = () => archiveDirectly();
}

function loadFicheFromUpload() {
    if (!window.tempFileData) return;
    
    const data = window.tempFileData.data;
    parseExcelFile(data);
    
    CURRENT_FICHE = {
        id: Date.now(),
        name: window.tempFileData.name,
        date: new Date().toLocaleString('fr-FR'),
        rows: JSON.parse(JSON.stringify(CURRENT_ROWS))
    };
    
    refreshFicheView();
    switchTab('fiche');
}

function archiveDirectly() {
    if (!window.tempFileData) return;
    
    const data = window.tempFileData.data;
    parseExcelFile(data);
    
    const archive = {
        id: Date.now(),
        name: window.tempFileData.name,
        date: new Date().toLocaleString('fr-FR'),
        rows: JSON.parse(JSON.stringify(CURRENT_ROWS))
    };
    
    ARCHIVES.push(archive);
    localStorage.setItem('expArchives', JSON.stringify(ARCHIVES));
    alert('✅ Fichier archivé directement');
    window.tempFileData = null;
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadStatus').style.display = 'none';
}

function parseExcelFile(data) {
    if (data.length < 4) return;

    const headerLine1 = data[0].map(h => String(h || '').toLowerCase().trim());
    const headers = headerLine1;
    
    // Détecter le format du fichier
    const isFicheTravail = headers.some(h => h.includes('semaine')) && headers.some(h => h.includes('réception'));
    const isArchives = headers.some(h => h.includes('ipo')) && headers.some(h => h.includes('emballage')) && headers.some(h => h.includes('nombre de colis'));
    
    CURRENT_ROWS = [];

    if (isFicheTravail) {
        // Format FICHE DE TRAVAIL
        for (let i = 3; i < data.length; i++) {
            const row = data[i];
            const semaine = parseInt(row[0]);
            if (isNaN(semaine)) continue;

            const dateReception = parseDate(row[1]);
            const heureReception = row[1] ? extractHour(row[1]) : 0;

            const obj = {
                format: 'fiche_travail',
                semaine,
                dateReception: dateReception ? dateReception.toISOString().split('T')[0] : '',
                heureReception,
                ipo: row[2] || '',
                pbrut: parseFloat(String(row[8] || '').replace(',', '.')) || 0,
                vol: parseFloat(String(row[9] || '').replace(',', '.')) || 0,
                L: parseFloat(String(row[5] || '').replace(',', '.')) || null,
                W: parseFloat(String(row[6] || '').replace(',', '.')) || null,
                H: parseFloat(String(row[7] || '').replace(',', '.')) || null,
                nbColis: 1 // Par défaut 1 colis par ligne
            };

            if (row[10] && row[10] !== '') {
                const dateDepart = parseDate(row[10]);
                if (dateDepart && dateReception) {
                    obj.dwellTimeHours = Math.round((dateDepart - dateReception) / (1000 * 60 * 60));
                }
            } else {
                const dens = obj.vol > 0 ? obj.pbrut / obj.vol : 0;
                obj.dwellTimeHours = Math.max(0.5, Math.min(48, (obj.vol * 1000 / 100) * (1 + dens / 500)));
            }

            if (obj.vol === 0 && obj.L && obj.W && obj.H) {
                obj.vol = (obj.L * obj.W * obj.H) / 1000000;
            }

            if (obj.L && obj.W && obj.H) {
                obj.dimKey = `${Math.round(obj.L)}×${Math.round(obj.W)}×${Math.round(obj.H)} cm`;
                obj.type = Math.max(obj.L, obj.W, obj.H) > 49 ? 'caisse' : 'colis';
            } else {
                obj.type = 'colis';
            }

            CURRENT_ROWS.push(obj);
        }
    } else if (isArchives) {
        // Format ARCHIVES
        // Trouver les indices des colonnes clés
        const ipoIdx = headers.findIndex(h => h.includes('ipo'));
        const nbColisIdx = headers.findIndex(h => h.includes('nombre de colis'));
        const dateEmballageIdx = headers.findIndex(h => h.includes('emballage') && h.includes('date'));
        const dateSortieIdx = headers.findIndex(h => h.includes('date de sortie'));
        const pbruts = headers.filter((h, i) => h.includes('poids brut')); // Plusieurs colonnes de poids brut
        
        for (let i = 2; i < data.length; i++) {
            const row = data[i];
            const ipo = row[ipoIdx] || '';
            if (!ipo) continue;

            const dateEmballage = dateEmballageIdx >= 0 ? parseDate(row[dateEmballageIdx]) : null;
            const dateSortie = dateSortieIdx >= 0 ? parseDate(row[dateSortieIdx]) : null;
            const nbColis = nbColisIdx >= 0 ? parseInt(row[nbColisIdx]) || 1 : 1;

            // Calculer le dwell time
            let dwellTimeHours = 0;
            if (dateEmballage && dateSortie) {
                dwellTimeHours = Math.round((dateSortie - dateEmballage) / (1000 * 60 * 60));
            }

            // Trouver le poids brut (prendre la première colonne remplie)
            let pbrut = 0;
            let pbruttIdx = headers.findIndex(h => h.includes('poids brut'));
            while (pbruttIdx >= 0 && pbrut === 0) {
                pbrut = parseFloat(String(row[pbruttIdx] || '').replace(',', '.')) || 0;
                if (pbrut === 0) {
                    pbruttIdx = headers.findIndex((h, idx) => idx > pbruttIdx && h.includes('poids brut'));
                } else {
                    break;
                }
            }

            const obj = {
                format: 'archives',
                ipo: ipo,
                nbColis: nbColis,
                dateEmballage: dateEmballage ? dateEmballage.toISOString().split('T')[0] : '',
                dateSortie: dateSortie ? dateSortie.toISOString().split('T')[0] : '',
                dwellTimeHours: dwellTimeHours,
                pbrut: pbrut,
                type: 'colis' // Les archives sont généralement des colis
            };

            CURRENT_ROWS.push(obj);
        }
    }

    FILTERED_ROWS = JSON.parse(JSON.stringify(CURRENT_ROWS));
}

function parseDate(str) {
    if (!str) return null;
    const s = String(str).trim();
    // Tentar múltiplos formatos
    const formats = [
        /(\d{1,2})\/(\d{1,2})\/(\d{4})/,
        /(\d{4})-(\d{1,2})-(\d{1,2})/,
        /(\d{1,2})-(\d{1,2})-(\d{4})/
    ];
    
    for (let fmt of formats) {
        const match = s.match(fmt);
        if (match) {
            let d, m, y;
            if (match[3].length === 4) {
                d = parseInt(match[1]);
                m = parseInt(match[2]);
                y = parseInt(match[3]);
            } else {
                y = parseInt(match[1]);
                m = parseInt(match[2]);
                d = parseInt(match[3]);
            }
            return new Date(y, m - 1, d);
        }
    }
    return null;
}

function extractHour(str) {
    if (!str) return 0;
    const s = String(str).trim();
    const match = s.match(/(\d{1,2}):(\d{2})/);
    return match ? parseInt(match[1]) : 0;
}

// ===== FICHE TAB =====
function refreshFicheView() {
    if (!CURRENT_FICHE || CURRENT_ROWS.length === 0) {
        document.getElementById('ficheInfo').innerHTML = '<p>Aucune fiche chargée</p>';
        return;
    }

    document.getElementById('ficheInfo').innerHTML = `<strong>${CURRENT_FICHE.name}</strong> (${CURRENT_ROWS.length} caisses)`;
    document.getElementById('ficheClearBtn').style.display = 'inline-block';

    populateFilters('fiche');
    refreshFicheDashboard();
}

document.getElementById('ficheClearBtn')?.addEventListener('click', () => {
    CURRENT_FICHE = null;
    CURRENT_ROWS = [];
    FILTERED_ROWS = [];
    destroyCharts();
    refreshFicheView();
});

document.getElementById('ficheResetBtn')?.addEventListener('click', () => {
    document.querySelectorAll('#ficheSemSelect > option').forEach(o => o.selected = false);
    document.querySelectorAll('#ficheIpoSelect > option').forEach(o => o.selected = false);
    applyFicheFilters();
});

document.getElementById('ficheExportBtn')?.addEventListener('click', exportFicheCSV);
document.getElementById('ficheSemSelect')?.addEventListener('change', applyFicheFilters);
document.getElementById('ficheIpoSelect')?.addEventListener('change', applyFicheFilters);

function populateFilters(prefix) {
    const semaines = [...new Set(CURRENT_ROWS.map(r => r.semaine))].sort((a, b) => a - b);
    const ipos = [...new Set(CURRENT_ROWS.map(r => r.ipo).filter(x => x))].sort();

    const semSelect = document.getElementById(`${prefix}SemSelect`);
    const ipoSelect = document.getElementById(`${prefix}IpoSelect`);
    
    if (semSelect) {
        semSelect.innerHTML = '';
        semaines.forEach(s => {
            const opt = document.createElement('option');
            opt.value = s;
            opt.textContent = `Semaine ${s}`;
            semSelect.appendChild(opt);
        });
    }

    if (ipoSelect) {
        ipoSelect.innerHTML = '';
        ipos.forEach(ip => {
            const opt = document.createElement('option');
            opt.value = ip;
            opt.textContent = ip;
            ipoSelect.appendChild(opt);
        });
    }
}

function applyFicheFilters() {
    const sels = Array.from(document.getElementById('ficheSemSelect').selectedOptions).map(o => parseInt(o.value));
    const ipos = Array.from(document.getElementById('ficheIpoSelect').selectedOptions).map(o => o.value);

    FILTERED_ROWS = CURRENT_ROWS.filter(r => {
        const sMatch = sels.length === 0 || sels.includes(r.semaine);
        const iMatch = ipos.length === 0 || ipos.includes(r.ipo);
        return sMatch && iMatch;
    });

    refreshFicheDashboard();
}

function refreshFicheDashboard() {
    updateFicheKPIs();
    updateFicheCharts();
}

function updateFicheKPIs() {
    const nb = FILTERED_ROWS.length;
    const vol = FILTERED_ROWS.reduce((s, r) => s + (r.vol || 0), 0);
    const pbrut = FILTERED_ROWS.reduce((s, r) => s + (r.pbrut || 0), 0);
    const dens = vol > 0 ? pbrut / vol : 0;
    const poidsMoy = nb > 0 ? pbrut / nb : 0;
    const volMoy = nb > 0 ? vol / nb : 0;
    const dwellTimeMoy = nb > 0 ? FILTERED_ROWS.reduce((s, r) => s + (r.dwellTimeHours || 0), 0) / nb : 0;

    // Métriques de tempo
    const minDwell = nb > 0 ? Math.min(...FILTERED_ROWS.map(r => r.dwellTimeHours || 0)) : 0;
    const maxDwell = nb > 0 ? Math.max(...FILTERED_ROWS.map(r => r.dwellTimeHours || 0)) : 0;

    // Densidade de utilização: (vol / pbrut) inversamente
    const utilisationDensity = dens > 0 ? (pbrut / (vol * 1000)) : 0;

    document.getElementById('fiche-kpi-nb').innerHTML = `<div class="kpi-value">${nb}</div><div class="kpi-label">Caisses</div>`;
    document.getElementById('fiche-kpi-vol').innerHTML = `<div class="kpi-value">${vol.toFixed(2)}</div><div class="kpi-label">Volume m³</div>`;
    document.getElementById('fiche-kpi-pbrut').innerHTML = `<div class="kpi-value">${pbrut.toFixed(0)}</div><div class="kpi-label">Poids kg</div>`;
    document.getElementById('fiche-kpi-dens').innerHTML = `<div class="kpi-value">${dens.toFixed(1)}</div><div class="kpi-label">Densité kg/m³</div>`;
    document.getElementById('fiche-kpi-poids-moy').innerHTML = `<div class="kpi-value">${poidsMoy.toFixed(1)}</div><div class="kpi-label">Poids moy/caisse</div>`;
    document.getElementById('fiche-kpi-vol-moy').innerHTML = `<div class="kpi-value">${volMoy.toFixed(4)}</div><div class="kpi-label">Vol moy/caisse</div>`;
    document.getElementById('fiche-kpi-score').innerHTML = `<div class="kpi-value">${calcScore(FILTERED_ROWS).total}</div><div class="kpi-label">Score</div>`;

    // Atualizar a nota de score
    const score = calcScore(FILTERED_ROWS);
    document.getElementById('fiche-score-note').textContent = score.note;

    // Adicionar métricas de dwell time
    const dwellNote = document.getElementById('fiche-dwell-info');
    if (dwellNote) {
        dwellNote.innerHTML = `
            <div class="metric-row">
                <span>Tempo no quai (médio):</span> <strong>${dwellTimeMoy.toFixed(1)} h</strong>
            </div>
            <div class="metric-row">
                <span>Min/Max:</span> <strong>${minDwell.toFixed(1)}h - ${maxDwell.toFixed(1)}h</strong>
            </div>
            <div class="metric-row">
                <span>Throughput:</span> <strong>${nb > 0 ? (nb / Math.max(1, dwellTimeMoy)).toFixed(1) : '–'}</strong> caisses/heure
            </div>
        `;
    }
}

function calcScore(rows) {
    if (rows.length === 0) return { total: '–', note: '' };

    const byWeek = {};
    rows.forEach(r => {
        if (!byWeek[r.semaine]) byWeek[r.semaine] = { nb: 0, vol: 0, pbrut: 0 };
        byWeek[r.semaine].nb += 1;
        byWeek[r.semaine].vol += r.vol || 0;
        byWeek[r.semaine].pbrut += r.pbrut || 0;
    });

    const weeks = Object.keys(byWeek).map(k => parseInt(k)).sort();
    const maxNb = Math.max(...weeks.map(w => byWeek[w].nb));
    const maxVol = Math.max(...weeks.map(w => byWeek[w].vol));

    let scoreVol = 0, scoreNb = 0, scoreDens = 0, scoreVar = 0;

    weeks.forEach(w => {
        const r = byWeek[w].vol / maxVol;
        scoreVol = Math.max(scoreVol, Math.min(30, 30 * r));
    });

    weeks.forEach(w => {
        const r = byWeek[w].nb / maxNb;
        scoreNb = Math.max(scoreNb, Math.min(30, 30 * r));
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

function updateFicheCharts() {
    const byWeek = {};
    FILTERED_ROWS.forEach(r => {
        if (!byWeek[r.semaine]) byWeek[r.semaine] = { nb: 0, vol: 0, pbrut: 0 };
        byWeek[r.semaine].nb += 1;
        byWeek[r.semaine].vol += r.vol || 0;
        byWeek[r.semaine].pbrut += r.pbrut || 0;
    });

    const weeks = Object.keys(byWeek).map(k => parseInt(k)).sort((a, b) => a - b);
    const labels = weeks.map(w => `S${w}`);

    destroyCharts();

    // Chart 1: Nb caisses
    new Chart(document.getElementById('chart-fiche-nb'), {
        type: 'bar',
        data: {
            labels,
            datasets: [{ label: 'Nb caisses', data: weeks.map(w => byWeek[w].nb), backgroundColor: '#0099ff' }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    // Chart 2: Volume
    new Chart(document.getElementById('chart-fiche-vol'), {
        type: 'line',
        data: {
            labels,
            datasets: [{ label: 'Volume (m³)', data: weeks.map(w => byWeek[w].vol), borderColor: '#00cc44', fill: false }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    // Chart 3: Densité
    new Chart(document.getElementById('chart-fiche-dens'), {
        type: 'line',
        data: {
            labels,
            datasets: [{ label: 'Densité', data: weeks.map(w => byWeek[w].vol > 0 ? byWeek[w].pbrut / byWeek[w].vol : 0), borderColor: '#ffaa00', fill: false }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    // Chart 4: Poids moyen
    new Chart(document.getElementById('chart-fiche-poids'), {
        type: 'bar',
        data: {
            labels,
            datasets: [{ label: 'Poids moy/caisse', data: weeks.map(w => byWeek[w].pbrut / byWeek[w].nb), backgroundColor: '#ff3333' }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    // Chart 5: Dimensions
    const dimCount = {};
    FILTERED_ROWS.forEach(r => { if (r.dimKey) dimCount[r.dimKey] = (dimCount[r.dimKey] || 0) + 1; });
    const topDims = Object.entries(dimCount).sort(([,a], [,b]) => b - a).slice(0, 12);

    new Chart(document.getElementById('chart-fiche-dim'), {
        type: 'bar',
        data: {
            labels: topDims.map(([k]) => k),
            datasets: [{ label: 'Nb caisses', data: topDims.map(([,v]) => v), backgroundColor: '#9933ff' }]
        },
        options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    // Chart 6: Distribution poids
    new Chart(document.getElementById('chart-fiche-dist-poids'), {
        type: 'bar',
        data: {
            labels: ['0-50kg', '50-100kg', '100-150kg', '150-200kg', '200+kg'],
            datasets: [{
                label: 'Distribution',
                data: [
                    FILTERED_ROWS.filter(r => r.pbrut >= 0 && r.pbrut < 50).length,
                    FILTERED_ROWS.filter(r => r.pbrut >= 50 && r.pbrut < 100).length,
                    FILTERED_ROWS.filter(r => r.pbrut >= 100 && r.pbrut < 150).length,
                    FILTERED_ROWS.filter(r => r.pbrut >= 150 && r.pbrut < 200).length,
                    FILTERED_ROWS.filter(r => r.pbrut >= 200).length
                ],
                backgroundColor: '#00aa99'
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    // Chart 7: Top 10 IPO
    const ipoCount = {};
    FILTERED_ROWS.forEach(r => { if (r.ipo) ipoCount[r.ipo] = (ipoCount[r.ipo] || 0) + 1; });
    const topIPO = Object.entries(ipoCount).sort(([,a], [,b]) => b - a).slice(0, 10);

    new Chart(document.getElementById('chart-fiche-top-ipo'), {
        type: 'bar',
        data: {
            labels: topIPO.map(([k]) => k),
            datasets: [{ label: 'Nb caisses', data: topIPO.map(([,v]) => v), backgroundColor: '#cc00cc' }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
}

function exportFicheCSV() {
    let csv = 'semaine;date;ipo;dim_caisse;pbrut;vol;L;W;H\n';
    FILTERED_ROWS.forEach(r => {
        csv += `${r.semaine};${r.date};${r.ipo};${r.dimKey || ''};${r.pbrut};${r.vol};${r.L || ''};${r.W || ''};${r.H || ''}\n`;
    });
    downloadCSV(csv, 'fiche_travail.csv');
}

// ===== ARCHIVES TAB =====
function refreshArchivesView() {
    const list = document.getElementById('archivesList');
    if (ARCHIVES.length === 0) {
        list.innerHTML = '<p class="empty-message">Aucune archive</p>';
        return;
    }

    list.innerHTML = ARCHIVES.map(arch => `
        <div class="archive-card">
            <div class="archive-name">${arch.name}</div>
            <div class="archive-meta">📅 ${arch.date}</div>
            <div class="archive-meta">📊 ${arch.rows.length} caisses</div>
            <div class="archive-actions">
                <button class="btn btn-secondary" onclick="viewArchiveDashboard(${arch.id})">📈 Voir</button>
                <button class="btn btn-secondary" onclick="restoreArchive(${arch.id})">📂 Restaurer</button>
                <button class="btn btn-danger" onclick="deleteArchive(${arch.id})">🗑️ Supprimer</button>
            </div>
        </div>
    `).join('');
}

function viewArchiveDashboard(id) {
    const arch = ARCHIVES.find(a => a.id === id);
    if (!arch) return;
    
    CURRENT_ROWS = JSON.parse(JSON.stringify(arch.rows));
    FILTERED_ROWS = JSON.parse(JSON.stringify(arch.rows));
    
    const container = document.getElementById('archivesList');
    container.innerHTML += `
        <div id="arch-${id}" class="archive-dashboard">
            <h3>${arch.name} - Dashboard</h3>
            <div class="dashboard-filters">
                <button class="btn btn-secondary" onclick="closeArchiveDashboard(${id})">✕ Fermer</button>
            </div>
            <div class="kpis-grid">
                <div class="kpi-card" id="arch-kpi-nb-${id}"></div>
                <div class="kpi-card" id="arch-kpi-vol-${id}"></div>
                <div class="kpi-card" id="arch-kpi-pbrut-${id}"></div>
                <div class="kpi-card" id="arch-kpi-dens-${id}"></div>
                <div class="kpi-card" id="arch-kpi-poids-${id}"></div>
                <div class="kpi-card" id="arch-kpi-vol-moy-${id}"></div>
            </div>
            <div class="charts-grid">
                <div class="chart-item"><canvas id="arch-chart-nb-${id}"></canvas></div>
                <div class="chart-item"><canvas id="arch-chart-vol-${id}"></canvas></div>
                <div class="chart-item"><canvas id="arch-chart-dens-${id}"></canvas></div>
                <div class="chart-item"><canvas id="arch-chart-dim-${id}"></canvas></div>
            </div>
        </div>
    `;

    renderArchiveKPIs(id);
    renderArchiveCharts(id);
}

function renderArchiveKPIs(id) {
    const nb = CURRENT_ROWS.length;
    const vol = CURRENT_ROWS.reduce((s, r) => s + (r.vol || 0), 0);
    const pbrut = CURRENT_ROWS.reduce((s, r) => s + (r.pbrut || 0), 0);
    const dens = vol > 0 ? pbrut / vol : 0;
    const dwellMoy = nb > 0 ? CURRENT_ROWS.reduce((s, r) => s + (r.dwellTimeHours || 0), 0) / nb : 0;

    // Métricas especiais para archives
    // Nb moyen de commandes par mois
    const ipos = [...new Set(CURRENT_ROWS.map(r => r.ipo).filter(x => x))];
    const nbUniqueIpo = ipos.length;
    
    // Períodos dos dados
    const dates = CURRENT_ROWS
        .filter(r => r.dateReception)
        .map(r => new Date(r.dateReception))
        .sort((a, b) => a - b);
    
    let nbMonths = 1;
    if (dates.length >= 2) {
        const firstDate = dates[0];
        const lastDate = dates[dates.length - 1];
        nbMonths = Math.max(1, Math.round((lastDate - firstDate) / (1000 * 60 * 60 * 24 * 30.5)) + 1);
    }

    const nbIpoParMois = (nbUniqueIpo / nbMonths).toFixed(1);
    
    // Nb commandes par heure (estimé)
    const heuresUniques = new Set(CURRENT_ROWS.map(r => r.heureReception).filter(h => h > 0)).size;
    const nbIpoParHeure = heuresUniques > 0 ? (nbUniqueIpo / heuresUniques).toFixed(1) : '–';

    // Velocidade de traitement
    const velocidadeTraitement = dwellMoy > 0 ? (nb / (dwellMoy / 24)).toFixed(1) : '–';

    // Velocidade de expédition
    const totalWeight = pbrut;
    const velocidadeExpedition = dwellMoy > 0 ? (totalWeight / dwellMoy).toFixed(1) : '–';

    document.getElementById(`arch-kpi-nb-${id}`).innerHTML = `
        <div class="kpi-value">${nb}</div>
        <div class="kpi-label">Caisses</div>
        <div class="kpi-meta">${nbUniqueIpo} commandes</div>
    `;
    
    document.getElementById(`arch-kpi-vol-${id}`).innerHTML = `
        <div class="kpi-value">${vol.toFixed(2)}</div>
        <div class="kpi-label">Volume m³</div>
        <div class="kpi-meta">${(vol / nb).toFixed(4)}/caisse</div>
    `;
    
    document.getElementById(`arch-kpi-pbrut-${id}`).innerHTML = `
        <div class="kpi-value">${pbrut.toFixed(0)}</div>
        <div class="kpi-label">Poids kg</div>
        <div class="kpi-meta">${(pbrut / nb).toFixed(1)}/caisse</div>
    `;
    
    document.getElementById(`arch-kpi-dens-${id}`).innerHTML = `
        <div class="kpi-value">${dens.toFixed(1)}</div>
        <div class="kpi-label">Densité</div>
        <div class="kpi-meta">Quai: ${dwellMoy.toFixed(1)}h</div>
    `;
    
    document.getElementById(`arch-kpi-poids-${id}`).innerHTML = `
        <div class="kpi-value">${nbIpoParMois}</div>
        <div class="kpi-label">IPO/mois</div>
        <div class="kpi-meta">${nbIpoParHeure}/heure</div>
    `;
    
    document.getElementById(`arch-kpi-vol-moy-${id}`).innerHTML = `
        <div class="kpi-value">${velocidadeTraitement}</div>
        <div class="kpi-label">Throughput</div>
        <div class="kpi-meta">caisses/jour</div>
    `;
}

function renderArchiveCharts(id) {
    // Simplified version with key charts only
}
 
function closeArchiveDashboard(id) {
    document.getElementById(`arch-${id}`).remove();
}

function restoreArchive(id) {
    const arch = ARCHIVES.find(a => a.id === id);
    if (!arch) return;
    
    CURRENT_FICHE = { ...arch };
    CURRENT_ROWS = JSON.parse(JSON.stringify(arch.rows));
    FILTERED_ROWS = JSON.parse(JSON.stringify(arch.rows));
    
    refreshFicheView();
    switchTab('fiche');
    alert(`✅ Archive restaurée en Fiche de Travail`);
}

function deleteArchive(id) {
    if (!confirm('Confirmer la suppression ?')) return;
    ARCHIVES = ARCHIVES.filter(a => a.id !== id);
    localStorage.setItem('expArchives', JSON.stringify(ARCHIVES));
    refreshArchivesView();
}

document.getElementById('archiveCurrentBtn')?.addEventListener('click', () => {
    if (!CURRENT_FICHE) {
        alert('Aucune fiche chargée');
        return;
    }
    
    ARCHIVES.push({ ...CURRENT_FICHE, rows: JSON.parse(JSON.stringify(CURRENT_ROWS)) });
    localStorage.setItem('expArchives', JSON.stringify(ARCHIVES));
    alert('✅ Fiche archivée');
});

// ===== COMPARAISON TAB =====
function setupComparisonTab() {
    const select1 = document.getElementById('compSelect1');
    const select2 = document.getElementById('compSelect2');
    
    select1.innerHTML = '<option value="current">-- Fiche actuelle --</option>';
    select2.innerHTML = '<option value="">-- Choisir une archive --</option>';
    
    ARCHIVES.forEach(arch => {
        select1.innerHTML += `<option value="arch_${arch.id}">${arch.name}</option>`;
        select2.innerHTML += `<option value="arch_${arch.id}">${arch.name}</option>`;
    });
}

document.getElementById('compRunBtn')?.addEventListener('click', () => {
    const s1 = document.getElementById('compSelect1').value;
    const s2 = document.getElementById('compSelect2').value;
    
    if (!s1 || !s2) {
        alert('Sélectionnez 2 fiches');
        return;
    }
    
    let rows1 = s1 === 'current' ? CURRENT_ROWS : ARCHIVES.find(a => a.id === parseInt(s1.split('_')[1]))?.rows || [];
    let rows2 = ARCHIVES.find(a => a.id === parseInt(s2.split('_')[1]))?.rows || [];
    
    const result = compareRows(rows1, rows2);
    displayComparison(result);
});

function compareRows(r1, r2) {
    const fileLabel1 = document.getElementById('compSelect1').value === 'current' ? CURRENT_FICHE?.name || 'Fiche actuelle' : 'Archive 1';
    const fileLabel2 = 'Archive 2';
    
    // Calcul des métriques pour chaque fichier
    const calc = (rows) => {
        const nbTotal = rows.length;
        const nbColis = rows.reduce((s, r) => s + (r.nbColis || 1), 0);
        const pbrut = rows.reduce((s, r) => s + (r.pbrut || 0), 0);
        const dwellMoy = nbTotal > 0 ? rows.reduce((s, r) => s + (r.dwellTimeHours || 0), 0) / nbTotal : 0;
        const ipos = new Set(rows.map(r => r.ipo).filter(x => x)).size;
        const poidsMoy = nbColis > 0 ? pbrut / nbColis : 0;
        
        // Grouper par IPO pour analyse détaillée
        const byIpo = {};
        rows.forEach(r => {
            const key = r.ipo || 'Sans IPO';
            if (!byIpo[key]) byIpo[key] = { nbColis: 0, pbrut: 0, dwellTime: 0, count: 0 };
            byIpo[key].nbColis += (r.nbColis || 1);
            byIpo[key].pbrut += (r.pbrut || 0);
            byIpo[key].dwellTime += (r.dwellTimeHours || 0);
            byIpo[key].count++;
        });
        
        return { nbTotal, nbColis, pbrut, dwellMoy, ipos, poidsMoy, byIpo };
    };

    const m1 = calc(r1);
    const m2 = calc(r2);

    const pctDiff = (v1, v2) => {
        if (v2 === 0 || v2 === null) return v1 === 0 ? 0 : 999;
        return ((v1 - v2) / v2 * 100).toFixed(1);
    };

    return {
        nom1: fileLabel1,
        nom2: fileLabel2,
        nbTotal1: m1.nbTotal, nbTotal2: m2.nbTotal,
        nbColis1: m1.nbColis, nbColis2: m2.nbColis,
        pbrut1: m1.pbrut, pbrut2: m2.pbrut,
        dwellMoy1: m1.dwellMoy, dwellMoy2: m2.dwellMoy,
        ipos1: m1.ipos, ipos2: m2.ipos,
        poidsMoy1: m1.poidsMoy, poidsMoy2: m2.poidsMoy,
        byIpo1: m1.byIpo,
        byIpo2: m2.byIpo,
        pctColis: pctDiff(m1.nbColis, m2.nbColis),
        pctPbrut: pctDiff(m1.pbrut, m2.pbrut),
        pctDwell: pctDiff(m1.dwellMoy, m2.dwellMoy)
    };
}

function displayComparison(result) {
    const comp = document.getElementById('comparisonResult');
    
    const diffClass = (val) => {
        return val < 0 ? ' class="diff negative"' : ' class="diff"';
    };
    
    const pctClass = (pct) => {
        const p = parseFloat(pct);
        if (isNaN(p)) return '';
        return p < 0 ? ' class="pct negative"' : ' class="pct"';
    };

    // Tableau principal - Comparaison globale
    let html = `
        <h3 style="margin-bottom: 1.5rem; color: var(--accent-light);">Comparaison Globale</h3>
        <table class="comparison-table">
            <thead>
                <tr>
                    <th>Métrique</th>
                    <th>${result.nom1}</th>
                    <th>${result.nom2}</th>
                    <th>Différence</th>
                    <th>% Variation</th>
                </tr>
            </thead>
            <tbody>
                <tr class="highlight">
                    <td><strong>Nombre total de colis</strong></td>
                    <td>${result.nbColis1}</td>
                    <td>${result.nbColis2}</td>
                    <td${diffClass(result.nbColis1 - result.nbColis2)}>${Math.abs(result.nbColis1 - result.nbColis2)}</td>
                    <td${pctClass(result.pctColis)}>${result.pctColis}%</td>
                </tr>
                <tr class="highlight">
                    <td><strong>Temps moyen quai (h)</strong></td>
                    <td>${result.dwellMoy1.toFixed(1)}</td>
                    <td>${result.dwellMoy2.toFixed(1)}</td>
                    <td${diffClass(result.dwellMoy1 - result.dwellMoy2)}>${Math.abs((result.dwellMoy1 - result.dwellMoy2).toFixed(1))}</td>
                    <td${pctClass(result.pctDwell)}>${result.pctDwell}%</td>
                </tr>
                <tr>
                    <td><strong>Poids total (kg)</strong></td>
                    <td>${result.pbrut1.toFixed(0)}</td>
                    <td>${result.pbrut2.toFixed(0)}</td>
                    <td${diffClass(result.pbrut1 - result.pbrut2)}>${Math.abs((result.pbrut1 - result.pbrut2).toFixed(0))}</td>
                    <td${pctClass(result.pctPbrut)}>${result.pctPbrut}%</td>
                </tr>
                <tr>
                    <td><strong>Poids moyen/colis (kg)</strong></td>
                    <td>${result.poidsMoy1.toFixed(1)}</td>
                    <td>${result.poidsMoy2.toFixed(1)}</td>
                    <td${diffClass(result.poidsMoy1 - result.poidsMoy2)}>${Math.abs((result.poidsMoy1 - result.poidsMoy2).toFixed(1))}</td>
                    <td class="pct">–</td>
                </tr>
                <tr>
                    <td><strong>Nombre unique d'IPO/SO</strong></td>
                    <td>${result.ipos1}</td>
                    <td>${result.ipos2}</td>
                    <td${diffClass(result.ipos1 - result.ipos2)}>${Math.abs(result.ipos1 - result.ipos2)}</td>
                    <td class="pct">–</td>
                </tr>
            </tbody>
        </table>
    `;

    // Analyse par IPO
    html += `<h3 style="margin-top: 2rem; margin-bottom: 1.5rem; color: var(--accent-light);">Détail par IPO/SO</h3>`;
    
    const ipos1 = Object.entries(result.byIpo1 || {}).sort((a, b) => b[1].nbColis - a[1].nbColis);
    const ipos2 = Object.entries(result.byIpo2 || {}).sort((a, b) => b[1].nbColis - a[1].nbColis);

    // Table IPO #1
    html += `<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; margin-bottom: 2rem;">`;
    html += `<div><h4 style="color: var(--accent); margin-bottom: 1rem;">${result.nom1}</h4>`;
    html += `<table class="comparison-table" style="font-size: 0.9rem;">
        <thead>
            <tr>
                <th>IPO/SO</th>
                <th>Colis</th>
                <th>Poids (kg)</th>
                <th>Dwell (h)</th>
            </tr>
        </thead>
        <tbody>`;
    
    ipos1.forEach(([ipo, metrics]) => {
        const dwellAvg = metrics.count > 0 ? (metrics.dwellTime / metrics.count).toFixed(1) : '–';
        html += `
            <tr>
                <td><strong>${ipo}</strong></td>
                <td>${metrics.nbColis}</td>
                <td>${metrics.pbrut.toFixed(0)}</td>
                <td>${dwellAvg}</td>
            </tr>
        `;
    });
    html += `</tbody></table></div>`;

    // Table IPO #2
    html += `<div><h4 style="color: var(--accent); margin-bottom: 1rem;">${result.nom2}</h4>`;
    html += `<table class="comparison-table" style="font-size: 0.9rem;">
        <thead>
            <tr>
                <th>IPO/SO</th>
                <th>Colis</th>
                <th>Poids (kg)</th>
                <th>Dwell (h)</th>
            </tr>
        </thead>
        <tbody>`;
    
    ipos2.forEach(([ipo, metrics]) => {
        const dwellAvg = metrics.count > 0 ? (metrics.dwellTime / metrics.count).toFixed(1) : '–';
        html += `
            <tr>
                <td><strong>${ipo}</strong></td>
                <td>${metrics.nbColis}</td>
                <td>${metrics.pbrut.toFixed(0)}</td>
                <td>${dwellAvg}</td>
            </tr>
        `;
    });
    html += `</tbody></table></div>`;
    html += `</div>`;

    comp.innerHTML = html;
}

function pctDiff(v1, v2) {
    if (v2 === 0 || v2 === null) return v1 === 0 ? 0 : 999;
    return ((v1 - v2) / v2 * 100).toFixed(1);
}

// ===== OUTILS UTILS =====
function destroyCharts() {
    Object.values(charts).forEach(c => c?.destroy?.());
    charts = {};
}

function downloadCSV(csv, filename) {
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
