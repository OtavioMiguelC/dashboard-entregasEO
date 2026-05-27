// Firebase Configuration
const firebaseConfig = {
  apiKey: "AIzaSyBeWhxIYWeAliOgRLFd4YdPLFZYrszku2A",
  authDomain: "dashboard-eo.firebaseapp.com",
  projectId: "dashboard-eo",
  storageBucket: "dashboard-eo.firebasestorage.app",
  messagingSenderId: "572483276607",
  appId: "1:572483276607:web:75375e0b3a8992911b8d8a",
  databaseURL: "https://dashboard-eo-default-rtdb.firebaseio.com/" 
};

// Initialize Firebase
firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.database();

// Global State
let rawData = [];
let filteredData = [];
let originalJsonData = []; // Agora será um Array de Arrays (Tabela Pura)
let charts = {};

// DOM Elements
const themeToggleCheckbox = document.getElementById('themeToggleCheckbox');
const excelUpload = document.getElementById('excelUpload');
const loginWall = document.getElementById('loginWall');
const appContainer = document.getElementById('appContainer');
const doLoginBtn = document.getElementById('doLoginBtn');
const uploadContainer = document.getElementById('uploadContainer');
const applyFiltersBtn = document.getElementById('applyFiltersBtn');
const clearFiltersBtn = document.getElementById('clearFiltersBtn');
const exportBtn = document.getElementById('exportBtn');

const adminReportsSection = document.getElementById('adminReportsSection');
const adminRankingBody = document.getElementById('adminRankingBody');
const adminExportBtn = document.getElementById('adminExportBtn');

const ADMIN_EMAIL = "otavio@oconsolida.com";

// Helper para parse de data
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

function resolveStatusLogico(prazoStr, dataEntregaStr, situacaoOriginal) {
    const sitLower = situacaoOriginal ? String(situacaoOriginal).toLowerCase() : '';
    
    const prazoDt = parseDateBR(prazoStr);
    const entregaDt = parseDateBR(dataEntregaStr);
    
    const temPrazo = prazoDt && !isNaN(prazoDt);
    const temEntrega = entregaDt && !isNaN(entregaDt);

    if (temEntrega) {
        if (!temPrazo) {
            return 'Entregue sem prazo';
        } else {
            entregaDt.setHours(0,0,0,0);
            prazoDt.setHours(0,0,0,0);
            if (entregaDt > prazoDt) return 'Atrasado';
            else return 'No prazo';
        }
    } else {
        if (!temPrazo) {
            return 'Sem prazo';
        } else {
            const hoje = new Date();
            hoje.setHours(0,0,0,0);
            prazoDt.setHours(0,0,0,0);
            if (prazoDt < hoje) return 'Atrasado'; 
            else return 'Aguardando'; 
        }
    }
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
        loginWall.classList.add('hidden');
        appContainer.classList.remove('hidden');
        document.getElementById('userEmail').textContent = user.email;
        
        const mainTabsContainer = document.getElementById('mainTabsContainer');
        
        if (user.email.toLowerCase() === ADMIN_EMAIL.toLowerCase()) {
            uploadContainer.classList.remove('hidden');
            adminReportsSection.classList.remove('hidden');
            if (mainTabsContainer) mainTabsContainer.classList.remove('hidden');
        } else {
            uploadContainer.classList.add('hidden');
            adminReportsSection.classList.add('hidden');
            if (mainTabsContainer) mainTabsContainer.classList.add('hidden');
        }
        
        loadDataFromRTDB();

    } else {
        loginWall.classList.remove('hidden');
        appContainer.classList.add('hidden');
    }
});

doLoginBtn.addEventListener('click', () => {
    const em = document.getElementById('loginEmail').value;
    const pw = document.getElementById('loginPassword').value;
    const errEl = document.getElementById('loginError');
    errEl.textContent = "Validando...";
    auth.signInWithEmailAndPassword(em, pw)
        .then(() => {
            errEl.textContent = "";
        })
        .catch(err => {
            errEl.textContent = "Erro: " + err.message;
        });
});

document.getElementById('logoutBtn').addEventListener('click', () => {
    auth.signOut();
});

// Load Initial Data from Database
function loadDataFromRTDB() {
    db.ref('relatorio_consolida').once('value')
        .then(snapshot => {
            const data = snapshot.val();
            if (data && Array.isArray(data) && data.length > 0) {
                originalJsonData = data;
                processDataEngine(); 
            } else {
                console.log("Banco de dados vazio ou formato inválido.");
            }
        })
        .catch(err => {
            console.error("Erro ao puxar dados do banco.", err);
        });
}

// Upload Excel -> Converter para Array of Arrays (AoA) -> Salvar
excelUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const btnLabel = document.querySelector('label[for="excelUpload"]');
    const oldText = btnLabel.textContent;
    btnLabel.textContent = "Lendo Excel...";
    
    const reader = new FileReader();
    reader.onload = (evt) => {
        try {
            const workbook = XLSX.read(evt.target.result, { type: 'binary' });
            const firstSheet = workbook.SheetNames[0];
            
            // Aqui é a mágica de fidelidade: lemos como matriz (Array de Arrays)
            // { header: 1 } retorna [ [col1, col2, col3], [val1, val2, val3], ... ]
            const rawAoA = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1, defval: "" });
            
            // Filtramos linhas vazias que o Excel as vezes cria
            const cleanedAoA = rawAoA.filter(row => row && row.length > 0 && row.some(cell => String(cell).trim() !== ""));

            if (cleanedAoA.length < 2) {
                alert("A planilha parece estar vazia ou sem cabeçalho.");
                btnLabel.textContent = oldText;
                return;
            }

            originalJsonData = cleanedAoA;
            
            btnLabel.textContent = "Salvando na Nuvem...";
            db.ref('relatorio_consolida').set(originalJsonData)
                .then(() => {
                    btnLabel.textContent = oldText;
                    alert("Dados salvos e prontos! A exportação será 100% idêntica ao original.");
                    processDataEngine();
                })
                .catch(err => {
                    btnLabel.textContent = oldText;
                    alert("Erro no banco de dados: " + err.message);
                });
        } catch (error) {
            btnLabel.textContent = oldText;
            alert("Erro ao ler planilha: " + error.message);
        }
    };
    reader.readAsBinaryString(file);
});

// Processa a Matriz (AoA) e alimenta os gráficos
function processDataEngine() {
    if (!originalJsonData || originalJsonData.length < 2) return;

    // A linha 0 sempre contém os cabeçalhos exatos originais.
    const headers = originalJsonData[0].map(h => String(h).toLowerCase().trim());

    const getIdx = (searchStr, excludeWords = []) => {
        const lowerSearch = searchStr.toLowerCase();
        // 1. Tenta correspondência exata
        const exact = headers.findIndex(h => h === lowerSearch);
        if (exact !== -1) return exact;
        
        // 2. Tenta includes, ignorando exclusões
        return headers.findIndex(h => {
            if (!h.includes(lowerSearch)) return false;
            for (let word of excludeWords) {
                if (h.includes(word)) return false;
            }
            return true;
        });
    };

    const idxPrazo = getIdx('prazo entrega');
    const idxSitOriginal = getIdx('situa') !== -1 ? getIdx('situa') : getIdx('status');
    const idxNfe = getIdx('nf-e') !== -1 ? getIdx('nf-e') : getIdx('nota');
    const idxEmissao = getIdx('emiss');
    const idxDestino = getIdx('destino') !== -1 ? getIdx('destino') : getIdx('cidade');
    const idxUf = getIdx('uf');
    const idxTransp = getIdx('transportadora') !== -1 ? getIdx('transportadora') : getIdx('transp');
    
    // Para CT-e, prioriza exato. Se for por includes, não pode ter 'série' nem 'serie' no nome
    let idxCte = getIdx('ct-e', ['série', 'serie']);
    if (idxCte === -1) idxCte = getIdx('doc.frete');
    if (idxCte === -1) idxCte = getIdx('frete');
    
    const idxSerieCte = getIdx('série ct-e') !== -1 ? getIdx('série ct-e') : getIdx('serie');
    const idxEntrega = getIdx('data entrega') !== -1 ? getIdx('data entrega') : getIdx('entrega');

    rawData = [];
    
    // Começa do índice 1 (pula o cabeçalho)
    for (let i = 1; i < originalJsonData.length; i++) {
        const row = originalJsonData[i];
        if (!row || row.length === 0) continue;

        const prazoStr = idxPrazo !== -1 ? row[idxPrazo] : null;
        const entregaStr = idxEntrega !== -1 ? row[idxEntrega] : null;
        const sitOriginal = idxSitOriginal !== -1 ? row[idxSitOriginal] : '';
        const situacaoReal = resolveStatusLogico(prazoStr, entregaStr, sitOriginal);

        rawData.push({
            _originalIndex: i, // Guarda qual é a linha verdadeira na Matriz
            nfe: idxNfe !== -1 ? row[idxNfe] : null,
            emissao: idxEmissao !== -1 ? row[idxEmissao] : null,
            destino: idxDestino !== -1 ? row[idxDestino] : null,
            uf: idxUf !== -1 ? row[idxUf] : null,
            transportadora: idxTransp !== -1 ? row[idxTransp] : null,
            docFrete: idxCte !== -1 ? row[idxCte] : null,
            serieCte: idxSerieCte !== -1 ? row[idxSerieCte] : null,
            prazoEntrega: prazoStr,
            dataEntrega: entregaStr,
            situacaoOriginal: sitOriginal,
            situacao: situacaoReal
        });
    }

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
        
        if (sFilter) {
            const isNoPrazo = sFilter.toLowerCase() === 'no prazo';
            if (isNoPrazo) {
                if (d.situacao !== 'No prazo' && d.situacao !== 'Aguardando') match = false;
            } else {
                if (d.situacao.toLowerCase() !== sFilter.toLowerCase()) match = false;
            }
        }
        
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
    updateDateRange();
    updateKPIs();
    renderSLAChart();
    renderRegionsChart();
    renderBottlenecksChart();
    renderRanking();
    renderAdminReports();
}

function updateDateRange() {
    const displayEl = document.getElementById('dateRangeDisplay');
    if (!displayEl) return;
    
    if (filteredData.length === 0) {
        displayEl.innerHTML = '<i class="fas fa-calendar-alt"></i> Sem dados de data';
        return;
    }
    
    let minDate = null;
    let maxDate = null;
    
    filteredData.forEach(d => {
        const dt = parseDateBR(d.emissao);
        if (dt && !isNaN(dt)) {
            if (!minDate || dt < minDate) minDate = dt;
            if (!maxDate || dt > maxDate) maxDate = dt;
        }
    });
    
    if (minDate && maxDate) {
        const fmt = (date) => `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth()+1).padStart(2, '0')}/${date.getFullYear()}`;
        displayEl.innerHTML = `<i class="fas fa-calendar-alt"></i> <strong>${fmt(minDate)}</strong> até <strong>${fmt(maxDate)}</strong>`;
    } else {
        displayEl.innerHTML = '<i class="fas fa-calendar-alt"></i> Data não disponível';
    }
}

function updateKPIs() {
    const totalNfe = filteredData.length;
    const totalCte = filteredData.filter(d => d.docFrete).length;
    const noPrazo = filteredData.filter(d => d.situacao === 'No prazo' || d.situacao === 'Aguardando').length;
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
        if(d.situacao === 'No prazo' || d.situacao === 'Aguardando') monthlyStats[monthKey].noPrazo++;
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
        danger: '#ef4444',
        info: '#3b82f6'
    };
}

function updateChartsTheme() {
    if(filteredData.length > 0) updateDashboard();
}

function renderSLAChart() {
    const ctx = document.getElementById('slaChart').getContext('2d');
    const colors = getChartColors();
    
    let noPrazo = 0; let atrasado = 0; let semPrazo = 0; let entSemPrazo = 0;

    filteredData.forEach(d => {
        if (d.situacao === 'No prazo' || d.situacao === 'Aguardando') noPrazo++;
        else if (d.situacao === 'Atrasado') atrasado++;
        else if (d.situacao === 'Entregue sem prazo') entSemPrazo++;
        else semPrazo++; 
    });

    if (charts.sla) charts.sla.destroy();

    charts.sla = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['No Prazo', 'Atrasado', 'Sem Prazo', 'Entregue sem Prazo'],
            datasets: [{
                data: [noPrazo, atrasado, semPrazo, entSemPrazo],
                backgroundColor: [colors.success, colors.danger, colors.warning, colors.info],
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
    
    const stats = {};
    filteredData.forEach(d => {
        if(!d.destino) return;
        if(d.situacao === 'Atrasado' || d.situacao === 'Sem prazo' || d.situacao === 'Entregue sem prazo') {
            if(!stats[d.destino]) stats[d.destino] = { atraso: 0, semPrazo: 0, entSemPrazo: 0 };
            if(d.situacao === 'Atrasado') stats[d.destino].atraso++;
            if(d.situacao === 'Sem prazo') stats[d.destino].semPrazo++;
            if(d.situacao === 'Entregue sem prazo') stats[d.destino].entSemPrazo++;
        }
    });

    const sorted = Object.entries(stats)
        .sort((a,b) => (b[1].atraso + b[1].semPrazo + b[1].entSemPrazo) - (a[1].atraso + a[1].semPrazo + a[1].entSemPrazo))
        .slice(0, 10);
    
    if (charts.regions) charts.regions.destroy();

    charts.regions = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(i => i[0].length > 15 ? i[0].substring(0, 15) + '...' : i[0]),
            datasets: [
                {
                    label: 'Atrasado',
                    data: sorted.map(i => i[1].atraso),
                    backgroundColor: colors.danger,
                    borderRadius: 2
                },
                {
                    label: 'Sem Prazo',
                    data: sorted.map(i => i[1].semPrazo),
                    backgroundColor: colors.warning,
                    borderRadius: 2
                },
                {
                    label: 'Ent. Sem Prazo',
                    data: sorted.map(i => i[1].entSemPrazo),
                    backgroundColor: colors.info,
                    borderRadius: 2
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                x: { stacked: true, ticks: { color: colors.text }, grid: { display: false } },
                y: { stacked: true, ticks: { color: colors.text }, grid: { color: colors.grid } }
            },
            plugins: { legend: { display: true, labels: { color: colors.text } } }
        }
    });
}

function renderBottlenecksChart() {
    const ctx = document.getElementById('bottlenecksChart').getContext('2d');
    const colors = getChartColors();
    
    const stats = {};
    filteredData.forEach(d => {
        if(!d.transportadora) return;
        if(!stats[d.transportadora]) stats[d.transportadora] = { atrasos: 0, semPrazo: 0, entSemPrazo: 0, noPrazo: 0 };
        if (d.situacao === 'Atrasado') stats[d.transportadora].atrasos++;
        if (d.situacao === 'Sem prazo') stats[d.transportadora].semPrazo++;
        if (d.situacao === 'Entregue sem prazo') stats[d.transportadora].entSemPrazo++;
        if (d.situacao === 'No prazo' || d.situacao === 'Aguardando') stats[d.transportadora].noPrazo++;
    });

    const sorted = Object.entries(stats)
        .sort((a,b) => (b[1].atrasos + b[1].semPrazo + b[1].entSemPrazo) - (a[1].atrasos + a[1].semPrazo + a[1].entSemPrazo))
        .slice(0, 15);

    if (charts.bottlenecks) charts.bottlenecks.destroy();

    charts.bottlenecks = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(i => i[0].substring(0, 20) + '...'),
            datasets: [
                { label: 'No Prazo', data: sorted.map(i => i[1].noPrazo), backgroundColor: colors.success },
                { label: 'Ent. Sem Prazo', data: sorted.map(i => i[1].entSemPrazo), backgroundColor: colors.info },
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
        if(!stats[t]) stats[t] = { total: 0, atraso: 0, semPrazo: 0, entSemPrazo: 0, noPrazo: 0 };
        stats[t].total++;
        if (d.situacao === 'Atrasado') stats[t].atraso++;
        else if (d.situacao === 'Sem prazo') stats[t].semPrazo++;
        else if (d.situacao === 'Entregue sem prazo') stats[t].entSemPrazo++;
        else if (d.situacao === 'No prazo' || d.situacao === 'Aguardando') stats[t].noPrazo++;
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
            <td>${r.noPrazo}</td>
            <td>${r.atraso}</td>
            <td>${r.semPrazo}</td>
            <td>${r.entSemPrazo}</td>
            <td>${r.percAtraso.toFixed(1)}%</td>
            <td><span class="status-badge ${statusClass}">${statusText}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

function renderAdminReports() {
    if (firebase.auth().currentUser && firebase.auth().currentUser.email.toLowerCase() !== ADMIN_EMAIL.toLowerCase()) return;
    
    adminRankingBody.innerHTML = '';
    
    // Agrupa notas Sem Prazo por Transportadora e Cidade
    const stats = {};
    filteredData.forEach(d => {
        if (d.situacao === 'Sem prazo' || d.situacao === 'Entregue sem prazo') {
            const key = `${d.transportadora || 'N/A'}|${d.destino || 'N/A'}`;
            if (!stats[key]) stats[key] = { transp: d.transportadora || 'N/A', cidade: d.destino || 'N/A', count: 0 };
            stats[key].count++;
        }
    });

    const ranking = Object.values(stats).sort((a, b) => b.count - a.count);

    ranking.forEach(r => {
        const tr = document.createElement('tr');
        let statusClass = 'status-warning';
        if (r.count > 10) statusClass = 'status-critical';
        
        tr.innerHTML = `
            <td>${r.transp}</td>
            <td>${r.cidade}</td>
            <td><strong>${r.count}</strong></td>
            <td><span class="status-badge ${statusClass}">Aguardando</span></td>
        `;
        adminRankingBody.appendChild(tr);
    });
}

// Exportação com fidelidade 100% (usando Array of Arrays)
exportBtn.addEventListener('click', () => {
    if(filteredData.length === 0) return alert('Não há dados para exportar.');
    
    // A linha 0 é o cabeçalho original com a exata ordem do Excel!
    const dataToExport = [originalJsonData[0]];
    
    // Adiciona as linhas correspondentes aos dados filtrados
    filteredData.forEach(d => {
        dataToExport.push(originalJsonData[d._originalIndex]);
    });
    
    // aoa_to_sheet reconstrói perfeitamente a planilha matriz
    const ws = XLSX.utils.aoa_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Relatorio_Filtrado");
    XLSX.writeFile(wb, "relatorio_consolida.xlsx");
});

adminExportBtn.addEventListener('click', () => {
    const semPrazoData = filteredData.filter(d => d.situacao === 'Sem prazo' || d.situacao === 'Entregue sem prazo');
    if(semPrazoData.length === 0) return alert('Não há notas "Sem prazo" ou "Entregue sem prazo" neste filtro.');
    
    // A linha 0 é o cabeçalho original com a exata ordem do Excel!
    const dataToExport = [originalJsonData[0]];
    
    semPrazoData.forEach(d => {
        dataToExport.push(originalJsonData[d._originalIndex]);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Notas_Sem_Prazo");
    XLSX.writeFile(wb, "relatorio_sem_prazo_consolida.xlsx");
});

// --- SISTEMA DE ABAS E CENTRAL DE E-MAILS ---
const tabBtnDashboard = document.getElementById('tabBtnDashboard');
const tabBtnEmailCenter = document.getElementById('tabBtnEmailCenter');
const dashboardSection = document.getElementById('dashboardSection');
const emailCenterSection = document.getElementById('emailCenterSection');

if (tabBtnDashboard && tabBtnEmailCenter) {
    tabBtnDashboard.addEventListener('click', () => {
        tabBtnDashboard.className = 'btn btn-primary';
        tabBtnEmailCenter.className = 'btn btn-secondary';
        dashboardSection.classList.remove('hidden');
        emailCenterSection.classList.add('hidden');
    });

    tabBtnEmailCenter.addEventListener('click', () => {
        tabBtnEmailCenter.className = 'btn btn-primary';
        tabBtnDashboard.className = 'btn btn-secondary';
        dashboardSection.classList.add('hidden');
        emailCenterSection.classList.remove('hidden');
        populateEmailTransp();
    });
}

function populateEmailTransp() {
    const emailSelectTransp = document.getElementById('emailSelectTransp');
    if(!emailSelectTransp) return;
    const currentVal = emailSelectTransp.value;
    
    const transpSet = new Set();
    rawData.forEach(d => {
        if (d.transportadora) transpSet.add(d.transportadora);
    });
    
    emailSelectTransp.innerHTML = '<option value="">Escolha...</option>';
    Array.from(transpSet).sort().forEach(t => {
        const opt = document.createElement('option');
        opt.value = t; opt.textContent = t;
        emailSelectTransp.appendChild(opt);
    });
    
    if (transpSet.has(currentVal)) {
        emailSelectTransp.value = currentVal;
    }
}

const btnGenerateEmail = document.getElementById('btnGenerateEmail');
const btnCopyAndOpen = document.getElementById('btnCopyAndOpen');
const emailHtmlContainer = document.getElementById('emailHtmlContainer');
const emailHelpText = document.getElementById('emailHelpText');

let currentGeneratedHtml = "";

if (btnGenerateEmail) {
    btnGenerateEmail.addEventListener('click', () => {
        const transp = document.getElementById('emailSelectTransp').value;
        if (!transp) return alert("Selecione uma transportadora primeiro.");
        
        const chkAtrasados = document.getElementById('chkAtrasados').checked;
        const chkSemPrazo = document.getElementById('chkSemPrazo').checked;
        const chkEntregueSemPrazo = document.getElementById('chkEntregueSemPrazo').checked;
        
        const allowedStatuses = [];
        if (chkAtrasados) allowedStatuses.push('Atrasado');
        if (chkSemPrazo) allowedStatuses.push('Sem prazo');
        if (chkEntregueSemPrazo) allowedStatuses.push('Entregue sem prazo');
        
        if (allowedStatuses.length === 0) return alert("Selecione pelo menos um status para cobrar.");

        const notasEmail = rawData.filter(d => d.transportadora === transp && allowedStatuses.includes(d.situacao));
        
        if (notasEmail.length === 0) {
            emailHtmlContainer.innerHTML = '<p style="color: red; text-align: center; margin-top: 50px;">Nenhuma nota encontrada para os filtros selecionados.</p>';
            btnCopyAndOpen.style.display = 'none';
            emailHelpText.style.display = 'none';
            return;
        }

        const countAtrasados = notasEmail.filter(n => n.situacao === 'Atrasado').length;
        const countSemPrazo = notasEmail.filter(n => n.situacao === 'Sem prazo').length;
        const countEntSemPrazo = notasEmail.filter(n => n.situacao === 'Entregue sem prazo').length;

        let html = `
        <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; max-width: 800px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; background-color: #ffffff;">
            
            <!-- Cabeçalho -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#5E17EB" style="background-color: #5E17EB; padding: 25px 30px;">
                <tr>
                    <td style="color: white;">
                        <h2 style="margin: 0; font-size: 22px; font-weight: 700;">Consolida — Relatório de Performance</h2>
                        <p style="margin: 5px 0 0 0; font-size: 14px; color: #e9d5ff;">Farol de entregas por status — ${transp}</p>
                    </td>
                    <td align="right">
                        <img src="https://raw.githubusercontent.com/OtavioMiguelC/dashboard-entregasEO/main/img/Consolida.png" alt="Consolida" style="height: 40px; background: white; padding: 8px 15px; border-radius: 6px; display: block;">
                    </td>
                </tr>
            </table>

            <!-- Corpo -->
            <div style="padding: 30px;">
                <h3 style="margin-top: 0; font-size: 18px; color: #2d3748;">Bom dia, equipe <strong>${transp}</strong>,</h3>
                <p style="font-size: 15px; color: #4a5568; line-height: 1.5; margin-bottom: 25px;">
                    Acompanhamos a performance das entregas usando um <strong>farol por cores</strong> para leitura executiva. 
                    Identificamos <strong>${notasEmail.length}</strong> ${notasEmail.length === 1 ? 'nota fiscal que requer' : 'notas fiscais que requerem'} atualização de status ou justificativa no sistema.
                </p>

                <!-- Alertas -->
                `;

        if (countAtrasados > 0) {
            html += `
                <div style="background-color: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; padding: 15px 20px; margin-bottom: 15px;">
                    <h4 style="margin: 0 0 8px 0; color: #dc2626; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px;">⚠️ Atenção Urgente</h4>
                    <p style="margin: 0; color: #991b1b; font-size: 15px;">
                        Há <strong>${countAtrasados} NFs em atraso</strong> aguardando retorno e ação imediata.
                    </p>
                </div>
            `;
        }

        if (countSemPrazo > 0) {
            html += `
                <div style="background-color: #fffbeb; border: 1px solid #fde68a; border-radius: 8px; padding: 15px 20px; margin-bottom: 15px;">
                    <h4 style="margin: 0 0 8px 0; color: #d97706; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px;">⏳ Prazos Faltando</h4>
                    <p style="margin: 0; color: #92400e; font-size: 15px;">
                        Há <strong>${countSemPrazo} entregas</strong> que ainda não possuem previsão registrada.
                    </p>
                </div>
            `;
        }
        
        if (countEntSemPrazo > 0) {
            html += `
                <div style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 15px 20px; margin-bottom: 15px;">
                    <h4 style="margin: 0 0 8px 0; color: #2563eb; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px;">📋 Validação Necessária</h4>
                    <p style="margin: 0; color: #1e40af; font-size: 15px;">
                        Há <strong>${countEntSemPrazo} entregas</strong> marcadas como concluídas, mas sem o registro original do prazo.
                    </p>
                </div>
            `;
        }

        html += `
                <!-- Tabela -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 30px; border-collapse: collapse; font-size: 13px;">
                    <thead>
                        <tr>
                            <th style="padding: 12px 8px; border-bottom: 2px solid #e2e8f0; text-align: left; color: #4a5568; font-weight: 600;">NF-e</th>
                            <th style="padding: 12px 8px; border-bottom: 2px solid #e2e8f0; text-align: left; color: #4a5568; font-weight: 600;">CT-e</th>
                            <th style="padding: 12px 8px; border-bottom: 2px solid #e2e8f0; text-align: left; color: #4a5568; font-weight: 600;">Destino</th>
                            <th style="padding: 12px 8px; border-bottom: 2px solid #e2e8f0; text-align: left; color: #4a5568; font-weight: 600;">Previsão</th>
                            <th style="padding: 12px 8px; border-bottom: 2px solid #e2e8f0; text-align: center; color: #4a5568; font-weight: 600;">Status</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        notasEmail.forEach(n => {
            let badgeColor = "#f59e0b";
            let badgeBg = "#fef3c7";
            let textColor = "#b45309";
            
            if (n.situacao === 'Atrasado') {
                badgeColor = "#ef4444"; badgeBg = "#fee2e2"; textColor = "#b91c1c";
            } else if (n.situacao === 'Entregue sem prazo') {
                badgeColor = "#3b82f6"; badgeBg = "#dbeafe"; textColor = "#1d4ed8";
            }
            
            const formatShortDate = (str) => {
                if (!str || str === '-') return '-';
                return String(str).split(' ')[0];
            };
            
            const prazoCurto = formatShortDate(n.prazoEntrega);
            
            html += `
                        <tr>
                            <td style="padding: 12px 8px; border-bottom: 1px solid #edf2f7; color: #1a202c; font-weight: 500;">${n.nfe || '-'}</td>
                            <td style="padding: 12px 8px; border-bottom: 1px solid #edf2f7; color: #4a5568;">${n.docFrete || '-'}</td>
                            <td style="padding: 12px 8px; border-bottom: 1px solid #edf2f7; color: #4a5568;">${n.destino || '-'}</td>
                            <td style="padding: 12px 8px; border-bottom: 1px solid #edf2f7; color: #4a5568;">${prazoCurto}</td>
                            <td style="padding: 12px 8px; border-bottom: 1px solid #edf2f7; text-align: center;">
                                <span style="background-color: ${badgeBg}; color: ${textColor}; border: 1px solid ${badgeColor}; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: bold; text-transform: uppercase;">
                                    ${n.situacao}
                                </span>
                            </td>
                        </tr>
            `;
        });
        
        html += `
                    </tbody>
                </table>
                
                <!-- Rodapé e Botão WhatsApp -->
                <div style="margin-top: 40px; text-align: center;">
                    <a href="https://wa.me/5547988411224" target="_blank" style="background-color: #25D366; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 15px; display: inline-block;">
                        💬 Fale Conosco no WhatsApp
                    </a>
                </div>

                <div style="margin-top: 35px; padding-top: 20px; border-top: 1px solid #e2e8f0; text-align: center;">
                    <p style="margin: 0; color: #4a5568; font-size: 14px;">Atenciosamente,</p>
                    <p style="margin: 5px 0; color: #1a202c; font-weight: 600; font-size: 15px;">Equipe Consolida</p>
                </div>
            </div>
        </div>
        `;
        
        currentGeneratedHtml = html;
        emailHtmlContainer.innerHTML = html;
        btnCopyAndOpen.style.display = 'block';
        emailHelpText.style.display = 'block';
    });
}

if (btnCopyAndOpen) {
    btnCopyAndOpen.addEventListener('click', async () => {
        try {
            const htmlBlob = new Blob([currentGeneratedHtml], { type: "text/html" });
            const plainBlob = new Blob(["Por favor, cole em um leitor de e-mails compatível com HTML."], { type: "text/plain" });
            const data = [new ClipboardItem({ "text/html": htmlBlob, "text/plain": plainBlob })];
            
            await navigator.clipboard.write(data);
            
            const destinatarios = document.getElementById('emailDestinos').value.trim();
            const transp = document.getElementById('emailSelectTransp').value;
            const assunto = encodeURIComponent(`Cobrança Logística - Notas Pendentes (${transp})`);
            
            window.location.href = `mailto:${destinatarios}?subject=${assunto}`;
            
            const oldText = btnCopyAndOpen.textContent;
            btnCopyAndOpen.textContent = "Copiado! Abra seu Outlook e cole (Ctrl+V)";
            setTimeout(() => { btnCopyAndOpen.textContent = oldText; }, 5000);
            
        } catch (err) {
            console.error("Erro ao copiar para clipboard: ", err);
            alert("Erro ao copiar a tabela HTML para a memória. O seu navegador pode ter bloqueado o recurso.");
        }
    });
}
