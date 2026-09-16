let rawData = [];
let filteredData = [];

let chartTopDuplicados = null;
let chartDonutProporcao = null;

function extractPartNumber(text) {
  if (!text || typeof text !== 'string') return "N/A";
  const cleaned = text.replace(/\b(TETRA PAK|ALFA LAVAL)\b/gi, '').trim();
  const tokens = cleaned.split(/\s+/);
  const digitTokens = tokens.filter(t => /\d/.test(t));
  return digitTokens.length > 0 ? digitTokens.join(" ") : "N/A";
}

document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById('excelFileInput');
  if (fileInput) {
    fileInput.addEventListener('change', handleFileUpload);
  }

  document.getElementById('searchInput').addEventListener('input', applyFilters);
  document.getElementById('filterComprador').addEventListener('change', applyFilters);
  document.getElementById('filterDuplicidade').addEventListener('change', applyFilters);
  document.getElementById('filterDivergente').addEventListener('change', applyFilters);
});

function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  document.getElementById('fileInfoLabel').innerText = `Base real (${file.name})`;

  const reader = new FileReader();
  reader.onload = function (evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames.includes('Sheet1') ? 'Sheet1' : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet);
    
    processDataset(json);
  };
  reader.readAsArrayBuffer(file);
}

function processDataset(json) {
  const parsed = json.map(row => {
    const textoBreve = String(row['Texto breve'] || row['Coluna1'] || '');
    const pn = extractPartNumber(textoBreve);
    let dtValidade = row['Fim da validade'];

    return {
      material: row['Material'] || 'N/A',
      textoBreve: textoBreve,
      partNumber: pn,
      comprador: row['Grupo de compradores'] || 'N/A',
      preco: parseFloat(row['Preço líquido']) || 0,
      validade: dtValidade ? String(dtValidade).substring(0, 10) : 'N/A'
    };
  });

  const pnFreqMap = {};
  const pnPricesMap = {};

  parsed.forEach(item => {
    if (item.partNumber !== 'N/A') {
      pnFreqMap[item.partNumber] = (pnFreqMap[item.partNumber] || 0) + 1;
      if (!pnPricesMap[item.partNumber]) pnPricesMap[item.partNumber] = new Set();
      pnPricesMap[item.partNumber].add(item.preco);
    }
  });

  rawData = parsed.map(item => ({
    ...item,
    frequencia: pnFreqMap[item.partNumber] || 1,
    temDivergenciaPreco: pnPricesMap[item.partNumber] ? pnPricesMap[item.partNumber].size > 1 : false
  }));

  populateFilterDropdowns();
  
  document.getElementById('mainContent').classList.remove('hidden');
  applyFilters();
}

function populateFilterDropdowns() {
  const compradores = [...new Set(rawData.map(d => d.comprador))].sort();
  const compSelect = document.getElementById('filterComprador');
  compSelect.innerHTML = '<option value="ALL">Todos os Compradores</option>';
  compradores.forEach(c => {
    compSelect.innerHTML += `<option value="${c}">${c}</option>`;
  });
}

function applyFilters() {
  const searchTerm = document.getElementById('searchInput').value.toLowerCase();
  const comprador = document.getElementById('filterComprador').value;
  const duplicidade = document.getElementById('filterDuplicidade').value;
  const divergente = document.getElementById('filterDivergente').value;

  filteredData = rawData.filter(item => {
    const matchSearch = item.textoBreve.toLowerCase().includes(searchTerm) ||
                        item.partNumber.toLowerCase().includes(searchTerm) ||
                        String(item.material).toLowerCase().includes(searchTerm);

    const matchComp = comprador === 'ALL' || item.comprador === comprador;

    let matchDup = true;
    if (duplicidade === 'UNICOS') matchDup = item.frequencia === 1;
    if (duplicidade === 'DUPLICADOS') matchDup = item.frequencia === 2;
    if (duplicidade === 'TRIPLICADOS') matchDup = item.frequencia >= 3;

    let matchDiv = true;
    if (divergente === 'SIM') matchDiv = item.temDivergenciaPreco === true;
    if (divergente === 'NAO') matchDiv = item.temDivergenciaPreco === false;

    return matchSearch && matchComp && matchDup && matchDiv;
  });

  updateDashboardUI();
}

function updateDashboardUI() {
  document.getElementById('filteredCountLabel').innerText = `${filteredData.length.toLocaleString('pt-BR')} de ${rawData.length.toLocaleString('pt-BR')} itens`;

  const totalItens = filteredData.length;
  const pnsUnicos = new Set(filteredData.filter(d => d.partNumber !== 'N/A').map(d => d.partNumber)).size;
  const duplicadosList = filteredData.filter(d => d.frequencia > 1 && d.partNumber !== 'N/A');
  const pnsDuplicadosUnicos = new Set(duplicadosList.map(d => d.partNumber)).size;
  const divergencias = new Set(filteredData.filter(d => d.temDivergenciaPreco).map(d => d.partNumber)).size;

  document.getElementById('kpiTotalItens').innerText = `${totalItens.toLocaleString('pt-BR')} ITENS`;
  document.getElementById('kpiTotalSub').innerText = `${totalItens.toLocaleString('pt-BR')} itens solicitados`;
  document.getElementById('kpiUnicos').innerText = `${pnsUnicos.toLocaleString('pt-BR')} PNs`;
  document.getElementById('kpiDuplicados').innerText = `${pnsDuplicadosUnicos.toLocaleString('pt-BR')} PNs`;
  document.getElementById('kpiDuplicadosSub').innerText = `Representam ${duplicadosList.length.toLocaleString('pt-BR')} itens`;
  document.getElementById('kpiDivergencias').innerText = `${divergencias} PNs`;

  renderCharts();
  renderTable();
}

function renderCharts() {
  const dupMap = {};
  filteredData.forEach(d => {
    if (d.frequencia > 1 && d.partNumber !== 'N/A') {
      dupMap[d.partNumber] = d.frequencia;
    }
  });
  const sortedDups = Object.entries(dupMap).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const dupLabels = sortedDups.map(d => d[0]);
  const dupValues = sortedDups.map(d => d[1]);

  if (chartTopDuplicados) chartTopDuplicados.destroy();
  const ctxDup = document.getElementById('chartTopDuplicados').getContext('2d');
  chartTopDuplicados = new Chart(ctxDup, {
    type: 'bar',
    data: {
      labels: dupLabels,
      datasets: [{
        data: dupValues,
        backgroundColor: '#10b981',
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: '#94a3b8', font: { size: 9 } } },
        y: { ticks: { color: '#94a3b8', stepSize: 1 } }
      }
    }
  });

  const unicosCount = filteredData.filter(d => d.frequencia === 1).length;
  const dup2Count = filteredData.filter(d => d.frequencia === 2).length;
  const dup3Count = filteredData.filter(d => d.frequencia >= 3).length;

  if (chartDonutProporcao) chartDonutProporcao.destroy();
  const ctxDonut = document.getElementById('chartDonutProporcao').getContext('2d');
  chartDonutProporcao = new Chart(ctxDonut, {
    type: 'doughnut',
    data: {
      labels: ['Únicos (1x)', 'Duplicados (2x)', 'Triplicados (3x+)'],
      datasets: [{
        data: [unicosCount, dup2Count, dup3Count],
        backgroundColor: ['#10b981', '#f59e0b', '#ef4444'],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8' } } },
      cutout: '70%'
    }
  });
}

function renderTable() {
  const tbody = document.getElementById('tableBody');
  tbody.innerHTML = '';

  const displayRows = filteredData.slice(0, 100);

  displayRows.forEach(item => {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-slate-800/50 transition border-b border-cardBorder/40";

    let badgeColor = "bg-slate-800 text-slate-400";
    if (item.frequencia === 2) badgeColor = "bg-amber-500/20 text-amber-400 border border-amber-500/30";
    if (item.frequencia >= 3) badgeColor = "bg-red-500/20 text-red-400 border border-red-500/30 font-bold";

    tr.innerHTML = `
      <td class="p-3 font-mono text-[11px] text-slate-400">${item.material}</td>
      <td class="p-3 font-sans text-xs text-slate-200">${item.textoBreve}</td>
      <td class="p-3 font-mono text-xs font-bold text-emeraldAccent">${item.partNumber}</td>
      <td class="p-3"><span class="px-2 py-0.5 rounded text-[10px] ${badgeColor}">${item.frequencia}x</span></td>
      <td class="p-3 text-xs text-slate-200">${item.preco ? item.preco.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : 'R$ 0,00'}</td>
      <td class="p-3 text-xs text-slate-400">${item.validade}</td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('tableStatus').innerText = `Exibindo ${displayRows.length.toLocaleString('pt-BR')} de ${filteredData.length.toLocaleString('pt-BR')} itens`;
}