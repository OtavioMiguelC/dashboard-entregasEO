// Firebase Configuration Placeholder - Você vai substituir isso depois!
const firebaseConfig = {
    apiKey: "SUA_API_KEY",
    authDomain: "SEU_PROJETO.firebaseapp.com",
    projectId: "SEU_PROJETO",
    storageBucket: "SEU_PROJETO.appspot.com",
    messagingSenderId: "ID",
    appId: "APP_ID"
};

// Global State
let rawData = [];
let filteredData = [];
let charts = {};

// DOM Elements
const themeToggle = document.getElementById('themeToggle');
const excelUpload = document.getElementById('excelUpload');
const adminPanel = document.getElementById('adminPanel');
const loginBtn = document.getElementById('loginBtn');
const loginModal = document.getElementById('loginModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const doLoginBtn = document.getElementById('doLoginBtn');
const applyFiltersBtn = document.getElementById('applyFiltersBtn');
const clearFiltersBtn = document.getElementById('clearFiltersBtn');
const exportBtn = document.getElementById('exportBtn');

// Theme Toggle
themeToggle.addEventListener('click', () => {
    const html = document.documentElement;
    const isDark = html.getAttribute('data-theme') === 'dark';
    html.setAttribute('data-theme', isDark ? 'light' : 'dark');
    themeToggle.textContent = isDark ? '🌙' : '☀️';
    updateChartsTheme();
});

// Mock Login Toggle (Until Firebase is active)
loginBtn.addEventListener('click', () => {
    loginModal.classList.remove('hidden');
});
closeModalBtn.addEventListener('click', () => {
    loginModal.classList.add('hidden');
});
doLoginBtn.addEventListener('click', () => {
    // Simulando login
    loginModal.classList.add('hidden');
    loginBtn.classList.add('hidden');
    adminPanel.classList.remove('hidden');
    document.getElementById('userEmail').textContent = document.getElementById('loginEmail').value || 'admin@consolida.com';
});
document.getElementById('logoutBtn').addEventListener('click', () => {
    loginBtn.classList.remove('hidden');
    adminPanel.classList.add('hidden');
});

// Excel Parsing
excelUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheet = workbook.SheetNames[0];
        
        // Converte para JSON
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { raw: false });
        
        // Normaliza as chaves por causa de acentuação no Excel original
        rawData = jsonData.map(row => {
            const getVal = (possibleKeys) => {
                for(let key of possibleKeys) {
                    if(row[key] !== undefined) return row[key];
                }
                return null;
            };

            return {
                nfe: getVal(['NF-e', 'NFe', 'Nota']),
                emissao: getVal(['Emissão', 'Emissao']),
                destino: getVal(['Destino', 'Cidade']),
                uf: getVal(['UF Destino', 'UF']),
                transportadora: getVal(['Transportadora', 'Transp']),
                docFrete: getVal(['Doc.Frete', 'CT-e', 'CTe']),
                serieCte: getVal(['Série CT-e', 'Serie CTe']),
                atraso: parseFloat(getVal(['Atraso'])) || 0,
                situacao: getVal(['Situação', 'Situacao', 'Status']) || 'Desconhecida'
            };
        });

        filteredData = [...rawData];
        populateFilters();
        updateDashboard();
    };
    reader.readAsBinaryString(file);
});

function populateFilters() {
    const transpSet = new Set();
    rawData.forEach(d => {
        if (d.transportadora) transpSet.add(d.transportadora);
    });
    
    const transpSelect = document.getElementById('filterTransportadora');
    transpSelect.innerHTML = '<option value="">Todas</option>';
    Array.from(transpSet).sort().forEach(t => {
        const opt = document.createElement('option');
        opt.value = t;
        opt.textContent = t;
        transpSelect.appendChild(opt);
    });
}

// Filtering Logic
applyFiltersBtn.addEventListener('click', () => {
    const tFilter = document.getElementById('filterTransportadora').value;
    const sFilter = document.getElementById('filterSituacao').value;
    const cFilter = document.getElementById('filterSemCte').value;
    const nFilter = document.getElementById('filterNfe').value.trim();

    filteredData = rawData.filter(d => {
        let match = true;
        if (tFilter && d.transportadora !== tFilter) match = false;
        if (sFilter && !d.situacao.toLowerCase().includes(sFilter.toLowerCase())) match = false;
        if (nFilter && String(d.nfe) !== nFilter) match = false;
        
        // Verifica NF-e sem CT-e (Doc.Frete vazio)
        if (cFilter === 'sim' && d.docFrete) match = false;
        
        return match;
    });

    updateDashboard();
});

clearFiltersBtn.addEventListener('click', () => {
    document.getElementById('filterTransportadora').value = '';
    document.getElementById('filterSituacao').value = '';
    document.getElementById('filterSemCte').value = '';
    document.getElementById('filterNfe').value = '';
    filteredData = [...rawData];
    updateDashboard();
});

// Update Dashboard
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
    const noPrazo = filteredData.filter(d => d.situacao.toLowerCase().includes('no prazo')).length;
    const slaPercent = totalNfe > 0 ? ((noPrazo / totalNfe) * 100).toFixed(1) : 0;
    const semCte = totalNfe - totalCte;

    document.getElementById('kpiTotalNfe').textContent = totalNfe.toLocaleString();
    document.getElementById('kpiTotalCte').textContent = totalCte.toLocaleString();
    document.getElementById('kpiSla').textContent = `${slaPercent}%`;
    document.getElementById('kpiSemCte').textContent = semCte.toLocaleString();
}

// Colors based on theme
function getChartColors() {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    return {
        text: isDark ? '#f8fafc' : '#1a202c',
        grid: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
        primary: isDark ? '#3b82f6' : '#004481',
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
    
    let noPrazo = 0;
    let atrasado = 0;
    let semPrazo = 0;

    filteredData.forEach(d => {
        const s = d.situacao.toLowerCase();
        if (s.includes('no prazo')) noPrazo++;
        else if (s.includes('atrasa')) atrasado++;
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
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: colors.text } }
            }
        }
    });
}

function renderRegionsChart() {
    const ctx = document.getElementById('regionsChart').getContext('2d');
    const colors = getChartColors();
    
    const counts = {};
    filteredData.forEach(d => {
        if(d.destino) {
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
                label: 'Volume de Entregas',
                data: sorted.map(i => i[1]),
                backgroundColor: colors.primary,
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { ticks: { color: colors.text }, grid: { color: colors.grid } },
                x: { ticks: { color: colors.text }, grid: { display: false } }
            },
            plugins: {
                legend: { display: false }
            }
        }
    });
}

function renderBottlenecksChart() {
    const ctx = document.getElementById('bottlenecksChart').getContext('2d');
    const colors = getChartColors();
    
    const stats = {};
    filteredData.forEach(d => {
        if(!d.transportadora) return;
        if(!stats[d.transportadora]) {
            stats[d.transportadora] = { atrasos: 0, semPrazo: 0 };
        }
        const s = d.situacao.toLowerCase();
        if (s.includes('atrasa')) stats[d.transportadora].atrasos++;
        if (s.includes('sem prazo')) stats[d.transportadora].semPrazo++;
    });

    // Pega as top piores em quantidade absoluta de atrasos + sem prazo
    const sorted = Object.entries(stats)
        .sort((a,b) => (b[1].atrasos + b[1].semPrazo) - (a[1].atrasos + a[1].semPrazo))
        .slice(0, 15);

    if (charts.bottlenecks) charts.bottlenecks.destroy();

    charts.bottlenecks = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(i => i[0].substring(0, 20) + '...'), // Truncate name
            datasets: [
                {
                    label: 'Atrasadas',
                    data: sorted.map(i => i[1].atrasos),
                    backgroundColor: colors.danger,
                },
                {
                    label: 'Sem Prazo',
                    data: sorted.map(i => i[1].semPrazo),
                    backgroundColor: colors.warning,
                }
            ]
        },
        options: {
            indexAxis: 'y', // Horizontal bar chart
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true, ticks: { color: colors.text }, grid: { color: colors.grid } },
                y: { stacked: true, ticks: { color: colors.text }, grid: { display: false } }
            },
            plugins: {
                legend: { labels: { color: colors.text } }
            }
        }
    });
}

function renderRanking() {
    const tbody = document.getElementById('rankingBody');
    tbody.innerHTML = '';

    const stats = {};
    filteredData.forEach(d => {
        const t = d.transportadora || 'N/A';
        if(!stats[t]) {
            stats[t] = { total: 0, atraso: 0, semPrazo: 0 };
        }
        stats[t].total++;
        const s = d.situacao.toLowerCase();
        if (s.includes('atrasa')) stats[t].atraso++;
        if (s.includes('sem prazo')) stats[t].semPrazo++;
    });

    const ranking = Object.entries(stats).map(([name, data]) => {
        const percAtraso = (data.atraso / data.total) * 100;
        return { name, ...data, percAtraso };
    });

    // Sort: Piores primeiro (Maior % de atraso, desempate por qtde de atrasos)
    ranking.sort((a, b) => b.percAtraso - a.percAtraso || b.atraso - a.atraso);

    ranking.forEach(r => {
        const tr = document.createElement('tr');
        
        let statusClass = 'status-good';
        let statusText = 'Excelente';
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

// Export Report
exportBtn.addEventListener('click', () => {
    if(filteredData.length === 0) return alert('Não há dados para exportar.');
    
    // Convert JSON back to worksheet
    const ws = XLSX.utils.json_to_sheet(filteredData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Relatorio_Filtrado");
    
    // Save
    XLSX.writeFile(wb, "relatorio_consolida.xlsx");
});
