// Firebase Configuration
const firebaseConfig = {
  apiKey: "AIzaSyBeWhxIYWeAliOgRLFd4YdPLFZYrszku2A",
  authDomain: "dashboard-eo.firebaseapp.com",
  projectId: "dashboard-eo",
  storageBucket: "dashboard-eo.firebasestorage.app",
  messagingSenderId: "572483276607",
  appId: "1:572483276607:web:75375e0b3a8992911b8d8a"
};

// Initialize Firebase
firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const storage = firebase.storage();

// Global State
let rawData = [];
let filteredData = [];
let originalJsonData = []; 
let charts = {};

// DOM Elements
const themeToggleCheckbox = document.getElementById('themeToggleCheckbox');
const excelUpload = document.getElementById('excelUpload');
const adminPanel = document.getElementById('adminPanel');
const loginBtn = document.getElementById('loginBtn');
const loginModal = document.getElementById('loginModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const doLoginBtn = document.getElementById('doLoginBtn');
const applyFiltersBtn = document.getElementById('applyFiltersBtn');
const clearFiltersBtn = document.getElementById('clearFiltersBtn');
const exportBtn = document.getElementById('exportBtn');

// Helper para parse de data DD-MM-YYYY HH:mm:ss ou YYYY-MM-DD
function parseDateBR(dateStr) {
    if(!dateStr) return null;
    if (dateStr instanceof Date) return dateStr;
    const parts = String(dateStr).split(/[- :\/T]/);
    if(parts.length >= 3) {
        if(parts[0].length === 4) return new Date(parts[0], parts[1]-1, parts[2]);
        return new Date(parts[2], parts[1]-1, parts[0]);
    }
    return new Date(dateStr);
}

function resolveStatusLogico(prazoStr, situacaoOriginal) {
    const sitLower = situacaoOriginal ? situacaoOriginal.toLowerCase() : '';
    if (!prazoStr || sitLower.includes('sem previs') || sitLower.includes('sem info')) return 'Sem prazo';
    if (sitLower.includes('entregue fora do prazo') || sitLower.includes('atras')) return 'Atrasado';
    if (sitLower === 'no prazo') return 'No prazo';

    const prazoDt = parseDateBR(prazoStr);
    const hoje = new Date();
    if(prazoDt) prazoDt.setHours(0,0,0,0);
    hoje.setHours(0,0,0,0);
    if (prazoDt < hoje) return 'Atrasado'; 
    else return 'No prazo'; 
}

// Theme Toggle
themeToggleCheckbox.addEventListener('change', (e) => {
    const html = document.documentElement;
    if (e.target.checked) html.setAttribute('data-theme', 'dark');
    else html.setAttribute('data-theme', 'light');
    updateChartsTheme();
});

// Auth Logic
auth.onAuthStateChanged(user => {
    if(user) {
        loginBtn.classList.add('hidden');
        adminPanel.classList.remove('hidden');
        document.getElementById('userEmail').textContent = user.email;
    } else {
        loginBtn.classList.remove('hidden');
        adminPanel.classList.add('hidden');
    }
});

loginBtn.addEventListener('click', () => {
    loginModal.classList.remove('hidden');
});
closeModalBtn.addEventListener('click', () => {
    loginModal.classList.add('hidden');
});
doLoginBtn.addEventListener('click', () => {
    const em = document.getElementById('loginEmail').value;
    const pw = document.getElementById('loginPassword').value;
    const errEl = document.getElementById('loginError');
    errEl.textContent = "Validando...";
    auth.signInWithEmailAndPassword(em, pw)
        .then(() => {
            loginModal.classList.add('hidden');
            errEl.textContent = "";
        })
        .catch(err => {
            errEl.textContent = "Erro: " + err.message;
        });
});
document.getElementById('logoutBtn').addEventListener('click', () => {
    auth.signOut();
});

// Load Initial Data from Storage
window.addEventListener('load', () => {
    storage.ref('registros.xlsx').getDownloadURL()
        .then(url => fetch(url))
        .then(res => res.blob())
        .then(blob => {
            const reader = new FileReader();
            reader.onload = (evt) => parseExcelData(evt.target.result);
            reader.readAsBinaryString(blob);
        })
        .catch(err => {
            console.log("Nenhum arquivo encontrado no Firebase ainda.");
        });
});

// Upload Excel
excelUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const btnLabel = document.querySelector('label[for="excelUpload"]');
    const oldText = btnLabel.textContent;
    btnLabel.textContent = "Enviando pra Nuvem...";
    
    storage.ref('registros.xlsx').put(file).then(() => {
        btnLabel.textContent = oldText;
        alert("Planilha sincronizada na Nuvem com sucesso!");
        
        const reader = new FileReader();
        reader.onload = (evt) => parseExcelData(evt.target.result);
        reader.readAsBinaryString(file);
    }).catch(err => {
        btnLabel.textContent = oldText;
        alert("Erro no upload: Você tem permissão no Firebase Rules?");
        console.error(err);
    });
});

function parseExcelData(binaryData) {
    const workbook = XLSX.read(binaryData, { type: 'binary' });
    const firstSheet = workbook.SheetNames[0];
    
    originalJsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { raw: false, defval: "" });
    
    rawData = originalJsonData.map((row, index) => {
        const getValFuzzy = (searchStr) => {
            const searchLower = searchStr.toLowerCase();
            for(let key in row) {
                if(key.toLowerCase().includes(searchLower)) return row[key];
            }
            return null;
        };

        const prazoStr = getValFuzzy('prazo entrega');
        const sitOriginal = getValFuzzy('situa') || getValFuzzy('status') || '';
        const situacaoReal = resolveStatusLogico(prazoStr, sitOriginal);

        return {
            _originalIndex: index,
            nfe: getValFuzzy('nf-e') || getValFuzzy('nota'),
            emissao: getValFuzzy('emiss'),
            destino: getValFuzzy('destino') || getValFuzzy('cidade'),
            uf: getValFuzzy('uf'),
            transportadora: getValFuzzy('transportadora') || getValFuzzy('transp'),
            docFrete: getValFuzzy('doc.frete') || getValFuzzy('ct-e'),
            serieCte: getValFuzzy('série ct-e') || getValFuzzy('serie'),
            prazoEntrega: prazoStr,
            situacaoOriginal: sitOriginal,
            situacao: situacaoReal
        };
    });

    filteredData = [...rawData];
    populateFilters();
    updateDashboard();
}

function populateFilters() {
    const transpSet = new Set();
    const mesSet = new Set();
    
    rawData.forEach(d => {
        if (d.transportadora) transpSet.add(d.transportadora);
        const dt = parseDateBR(d.emissao);
        if (dt && !isNaN(dt)) {
            const mesStr = `${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;
            mesSet.add(mesStr);
        }
    });
    
    const transpSelect = document.getElementById('filterTransportadora');
    transpSelect.innerHTML = '<option value="">Todas</option>';
    Array.from(transpSet).sort().forEach(t => {
        const opt = document.createElement('option');
        opt.value = t; opt.textContent = t; transpSelect.appendChild(opt);
    });

    const mesSelect = document.getElementById('filterMes');
    mesSelect.innerHTML = '<option value="">Todos os Meses</option>';
    Array.from(mesSet).sort((a,b) => {
        const [m1, y1] = a.split('/'); const [m2, y2] = b.split('/');
        return new Date(y2, m2-1) - new Date(y1, m1-1);
    }).forEach(m => {
        const opt = document.createElement('option');
        opt.value = m; opt.textContent = m; mesSelect.appendChild(opt);
    });
}

applyFiltersBtn.addEventListener('click', () => {
    const tFilter = document.getElementById('filterTransportadora').value;
    const sFilter = document.getElementById('filterSituacao').value;
    const cFilter = document.getElementById('filterSemCte').value;
    const nFilter = document.getElementById('filterNfe').value.trim();
    const mFilter = document.getElementById('filterMes').value;

    filteredData = rawData.filter(d => {
        let match = true;
        if (tFilter && d.transportadora !== tFilter) match = false;
        if (sFilter && !d.situacao.toLowerCase().includes(sFilter.toLowerCase())) match = false;
        if (nFilter && String(d.nfe) !== nFilter) match = false;
        if (cFilter === 'sim' && d.docFrete) match = false;
        if (mFilter) {
            const dt = parseDateBR(d.emissao);
            if (!dt || isNaN(dt)) match = false;
            else {
                const mesStr = `${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;
                if (mesStr !== mFilter) match = false;
            }
        }
        return match;
    });

    updateDashboard();
});

clearFiltersBtn.addEventListener('click', () => {
    document.getElementById('filterTransportadora').value = '';
    document.getElementById('filterSituacao').value = '';
    document.getElementById('filterSemCte').value = '';
    document.getElementById('filterNfe').value = '';
    document.getElementById('filterMes').value = '';
    filteredData = [...rawData];
    updateDashboard();
});

function updateDashboard() {
    updateKPIs();
    renderSLAChart();
    renderRegionsChart();
    renderBottlenecksChart();
    renderRanking();
}

function updateKPIs() {
    const totalNfe = filteredData.length;
    const totalCte = filteredData.filter(d => d.docFrete).length;
    const noPrazo = filteredData.filter(d => d.situacao === 'No prazo').length;
    const slaPercent = totalNfe > 0 ? ((noPrazo / totalNfe) * 100).toFixed(1) : 0;
    const semCte = totalNfe - totalCte;

    document.getElementById('kpiTotalNfe').textContent = totalNfe.toLocaleString();
    document.getElementById('kpiTotalCte').textContent = totalCte.toLocaleString();
    document.getElementById('kpiSla').textContent = `${slaPercent}%`;
    document.getElementById('kpiSemCte').textContent = semCte.toLocaleString();

    calculateMoMEvolution();
}

function formatEvo(elId, diff, isPercent, isGoodWhenUp) {
    const el = document.getElementById(elId);
    if (!el) return;
    if (diff === null) {
        el.textContent = '-'; el.className = 'kpi-evolution evo-neutral'; return;
    }
    
    const sign = diff > 0 ? '+' : '';
    const arrow = diff > 0 ? '↑' : (diff < 0 ? '↓' : '');
    const valStr = isPercent ? diff.toFixed(1) + '%' : diff.toLocaleString();
    
    el.textContent = `${arrow} ${sign}${valStr}`;
    
    if (diff === 0) el.className = 'kpi-evolution evo-neutral';
    else if (diff > 0) el.className = 'kpi-evolution ' + (isGoodWhenUp ? 'evo-good' : 'evo-bad');
    else el.className = 'kpi-evolution ' + (isGoodWhenUp ? 'evo-bad' : 'evo-good');
}

function calculateMoMEvolution() {
    const monthlyStats = {};
    const tFilter = document.getElementById('filterTransportadora').value;
    const sFilter = document.getElementById('filterSituacao').value;
    const cFilter = document.getElementById('filterSemCte').value;
    const nFilter = document.getElementById('filterNfe').value.trim();
    const mFilter = document.getElementById('filterMes').value;

    rawData.forEach(d => {
        if (tFilter && d.transportadora !== tFilter) return;
        if (sFilter && !d.situacao.toLowerCase().includes(sFilter.toLowerCase())) return;
        if (nFilter && String(d.nfe) !== nFilter) return;
        if (cFilter === 'sim' && d.docFrete) return;

        const dt = parseDateBR(d.emissao);
        if(!dt || isNaN(dt)) return;
        const monthKey = `${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;
        
        if(!monthlyStats[monthKey]) monthlyStats[monthKey] = { totalNfe: 0, totalCte: 0, noPrazo: 0 };
        monthlyStats[monthKey].totalNfe++;
        if (d.docFrete) monthlyStats[monthKey].totalCte++;
        if(d.situacao === 'No prazo') monthlyStats[monthKey].noPrazo++;
    });

    let currMonthStr = mFilter;
    const availableMonths = Object.keys(monthlyStats).sort((a,b) => {
        const [m1, y1] = a.split('/'); const [m2, y2] = b.split('/');
        return new Date(y1, m1-1) - new Date(y2, m2-1);
    });

    if (!currMonthStr && availableMonths.length > 0) currMonthStr = availableMonths[availableMonths.length - 1];

    if (currMonthStr) {
        const [cm, cy] = currMonthStr.split('/');
        let prevM = parseInt(cm) - 1;
        let prevY = parseInt(cy);
        if (prevM === 0) { prevM = 12; prevY -= 1; }
        const prevMonthStr = `${String(prevM).padStart(2,'0')}/${prevY}`;

        const curr = monthlyStats[currMonthStr] || { totalNfe: 0, totalCte: 0, noPrazo: 0 };
        const prev = monthlyStats[prevMonthStr] || { totalNfe: 0, totalCte: 0, noPrazo: 0 };

        const currSemCte = curr.totalNfe - curr.totalCte;
        const prevSemCte = prev.totalNfe - prev.totalCte;

        const currSla = curr.totalNfe > 0 ? (curr.noPrazo / curr.totalNfe) * 100 : 0;
        const prevSla = prev.totalNfe > 0 ? (prev.noPrazo / prev.totalNfe) * 100 : 0;
        
        formatEvo('kpiEvoNfe', curr.totalNfe - prev.totalNfe, false, true);
        formatEvo('kpiEvoCte', curr.totalCte - prev.totalCte, false, true);
        formatEvo('kpiEvoSla', currSla - prevSla, true, true);
        formatEvo('kpiEvoSemCte', currSemCte - prevSemCte, false, false);
    } else {
        ['kpiEvoNfe', 'kpiEvoCte', 'kpiEvoSla', 'kpiEvoSemCte'].forEach(id => formatEvo(id, null));
    }
}

function getChartColors() {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    return {
        text: isDark ? '#f8fafc' : '#1a202c',
        grid: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
        primary: isDark ? '#7c3aed' : '#5E17EB',
        success: '#10b981',
        warning: '#f59e0b',
        danger: '#ef4444'
    };
}

function updateChartsTheme() {
    if(filteredData.length > 0) updateDashboard();
}

function renderSLAChart() {
    const ctx = document.getElementById('slaChart').getContext('2d');
    const colors = getChartColors();
    
    let noPrazo = 0; let atrasado = 0; let semPrazo = 0;

    filteredData.forEach(d => {
        if (d.situacao === 'No prazo') noPrazo++;
        else if (d.situacao === 'Atrasado') atrasado++;
        else semPrazo++; 
    });

    if (charts.sla) charts.sla.destroy();

    charts.sla = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['No Prazo', 'Atrasado', 'Sem Prazo'],
            datasets: [{
                data: [noPrazo, atrasado, semPrazo],
                backgroundColor: [colors.success, colors.danger, colors.warning],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { labels: { color: colors.text } } }
        }
    });
}

function renderRegionsChart() {
    const ctx = document.getElementById('regionsChart').getContext('2d');
    const colors = getChartColors();
    
    const counts = {};
    filteredData.forEach(d => {
        if(d.destino && (d.situacao === 'Atrasado' || d.situacao === 'Sem prazo')) {
            counts[d.destino] = (counts[d.destino] || 0) + 1;
        }
    });

    const sorted = Object.entries(counts).sort((a,b) => b[1] - a[1]).slice(0, 10);
    
    if (charts.regions) charts.regions.destroy();

    charts.regions = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(i => i[0]),
            datasets: [{
                label: 'Atrasos / Sem Prazo',
                data: sorted.map(i => i[1]),
                backgroundColor: colors.danger,
                borderRadius: 4
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                y: { ticks: { color: colors.text }, grid: { color: colors.grid } },
                x: { ticks: { color: colors.text }, grid: { display: false } }
            },
            plugins: { legend: { display: false } }
        }
    });
}

function renderBottlenecksChart() {
    const ctx = document.getElementById('bottlenecksChart').getContext('2d');
    const colors = getChartColors();
    
    const stats = {};
    filteredData.forEach(d => {
        if(!d.transportadora) return;
        if(!stats[d.transportadora]) stats[d.transportadora] = { atrasos: 0, semPrazo: 0, noPrazo: 0 };
        if (d.situacao === 'Atrasado') stats[d.transportadora].atrasos++;
        if (d.situacao === 'Sem prazo') stats[d.transportadora].semPrazo++;
        if (d.situacao === 'No prazo') stats[d.transportadora].noPrazo++;
    });

    const sorted = Object.entries(stats)
        .sort((a,b) => (b[1].atrasos + b[1].semPrazo) - (a[1].atrasos + a[1].semPrazo))
        .slice(0, 15);

    if (charts.bottlenecks) charts.bottlenecks.destroy();

    charts.bottlenecks = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(i => i[0].substring(0, 20) + '...'),
            datasets: [
                { label: 'No Prazo', data: sorted.map(i => i[1].noPrazo), backgroundColor: colors.success },
                { label: 'Sem Prazo', data: sorted.map(i => i[1].semPrazo), backgroundColor: colors.warning },
                { label: 'Atrasadas', data: sorted.map(i => i[1].atrasos), backgroundColor: colors.danger }
            ]
        },
        options: {
            indexAxis: 'y', responsive: true, maintainAspectRatio: false,
            scales: {
                x: { stacked: true, ticks: { color: colors.text }, grid: { color: colors.grid } },
                y: { stacked: true, ticks: { color: colors.text }, grid: { display: false } }
            },
            plugins: { legend: { labels: { color: colors.text } } }
        }
    });
}

function renderRanking() {
    const tbody = document.getElementById('rankingBody');
    tbody.innerHTML = '';

    const stats = {};
    filteredData.forEach(d => {
        const t = d.transportadora || 'N/A';
        if(!stats[t]) stats[t] = { total: 0, atraso: 0, semPrazo: 0 };
        stats[t].total++;
        if (d.situacao === 'Atrasado') stats[t].atraso++;
        if (d.situacao === 'Sem prazo') stats[t].semPrazo++;
    });

    const ranking = Object.entries(stats).map(([name, data]) => {
        const percAtraso = (data.atraso / data.total) * 100;
        return { name, ...data, percAtraso };
    });

    ranking.sort((a, b) => b.percAtraso - a.percAtraso || b.atraso - a.atraso);

    ranking.forEach(r => {
        const tr = document.createElement('tr');
        let statusClass = 'status-good'; let statusText = 'Excelente';
        if (r.percAtraso > 15) { statusClass = 'status-critical'; statusText = 'Crítico'; tr.classList.add('worst-performer'); }
        else if (r.percAtraso > 5) { statusClass = 'status-warning'; statusText = 'Atenção'; }

        tr.innerHTML = `
            <td>${r.name}</td>
            <td>${r.total}</td>
            <td>${r.atraso}</td>
            <td>${r.semPrazo}</td>
            <td>${r.percAtraso.toFixed(1)}%</td>
            <td><span class="status-badge ${statusClass}">${statusText}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

exportBtn.addEventListener('click', () => {
    if(filteredData.length === 0) return alert('Não há dados para exportar.');
    const dataToExport = filteredData.map(d => originalJsonData[d._originalIndex]);
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Relatorio_Filtrado");
    XLSX.writeFile(wb, "relatorio_consolida.xlsx");
});
