let rawData = [];
let filteredData = [];

let chartTopDuplicados = null;
let chartDonutProporcao = null;
let chartValidade = null;
let chartFreqDistribucao = null;

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
  document.getElementById('filterCPP').addEventListener('change', applyFilters);
  document.getElementById('filterComprador').addEventListener('change', applyFilters);
  document.getElementById('filterDuplicidade').addEventListener('change', applyFilters);
});

function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  document.getElementById('fileInfoLabel').innerText = `Base carregada: (${file.name})`;

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
      cpp: String(row['CPP'] || row['Centro'] || row['Plant'] || 'N/A').trim(),
      textoBreve: textoBreve,
      partNumber: pn,
      comprador: String(row['Grupo de compradores'] || 'N/A').trim(),
      validade: dtValidade ? String(dtValidade).substring(0, 10) : 'N/A'
    };
  });

  const pnFreqMap = {};
  parsed.forEach(item => {
    if (item.partNumber !== 'N/A') {
      pnFreqMap[item.partNumber] = (pnFreqMap[item.partNumber] || 0) + 1;
    }
  });

  rawData = parsed.map(item => ({
    ...item,
    frequencia: pnFreqMap[item.partNumber] || 1
  }));

  populateFilterDropdowns();
  applyFilters();
}

function populateFilterDropdowns() {
  // Popula Compradores
  const compradores = [...new Set(rawData.map(d => d.comprador))].sort();
  const compSelect = document.getElementById('filterComprador');
  compSelect.innerHTML = '<option value="ALL">Todos os Compradores</option>';
  compradores.forEach(c => {
    compSelect.innerHTML += `<option value="${c}">${c}</option>`;
  });

  // Popula CPP
  const cpps = [...new Set(rawData.map(d => d.cpp))].sort();
  const cppSelect = document.getElementById('filterCPP');
  cppSelect.innerHTML = '<option value="ALL">Todos os CPPs</option>';
  cpps.forEach(cpp => {
    cppSelect.innerHTML += `<option value="${cpp}">CPP: ${cpp}</option>`;
  });
}

function applyFilters() {
  const searchTerm = document.getElementById('searchInput').value.toLowerCase();
  const cppVal = document.getElementById('filterCPP').value;
  const comprador = document.getElementById('filterComprador').value;
  const duplicidade = document.getElementById('filterDuplicidade').value;

  filteredData = rawData.filter(item => {
    const matchSearch = item.textoBreve.toLowerCase().includes(searchTerm) ||
                        item.partNumber.toLowerCase().includes(searchTerm) ||
                        String(item.material).toLowerCase().includes(searchTerm);

    const matchCPP = cppVal === 'ALL' || item.cpp === cppVal;
    const matchComp = comprador === 'ALL' || item.comprador === comprador;

    let matchDup = true;
    if (duplicidade === 'UNICOS') matchDup = item.frequencia === 1;
    if (duplicidade === 'DUPLICADOS') matchDup = item.frequencia === 2;
    if (duplicidade === 'TRIPLICADOS') matchDup = item.frequencia >= 3;

    return matchSearch && matchCPP && matchComp && matchDup;
  });

  updateDashboardUI();
}

function updateDashboardUI() {
  document.getElementById('filteredCountLabel').innerText = `${filteredData.length.toLocaleString('pt-BR')} de ${rawData.length.toLocaleString('pt-BR')} itens`;

  const totalItens = filteredData.length;
  const pnsUnicos = new Set(filteredData.filter(d => d.partNumber !== 'N/A').map(d => d.partNumber)).size;
  
  const duplicadosList = filteredData.filter(d => d.frequencia > 1 && d.partNumber !== 'N/A');
  const pnsDuplicadosUnicos = new Set(duplicadosList.map(d => d.partNumber)).size;

  let maxRep = 0;
  let maxPNs = [];
  filteredData.forEach(d => {
    if (d.frequencia > maxRep) {
      maxRep = d.frequencia;
      maxPNs = [d.partNumber];
    } else if (d.frequencia === maxRep && maxRep > 1 && !maxPNs.includes(d.partNumber)) {
      maxPNs.push(d.partNumber);
    }
  });

  document.getElementById('kpiTotalItens').innerText = `${totalItens.toLocaleString('pt-BR')} ITENS`;
  document.getElementById('kpiTotalSub').innerText = `${totalItens.toLocaleString('pt-BR')} linhas solicitadas`;
  document.getElementById('kpiUnicos').innerText = `${pnsUnicos.toLocaleString('pt-BR')} PNs`;
  document.getElementById('kpiDuplicados').innerText = `${pnsDuplicadosUnicos.toLocaleString('pt-BR')} PNs`;
  document.getElementById('kpiDuplicadosSub').innerText = `Representam ${duplicadosList.length.toLocaleString('pt-BR')} linhas`;
  
  document.getElementById('kpiMaxRepeticao').innerText = maxRep > 1 ? `${maxRep} VEZES` : '1 VEZ';
  document.getElementById('kpiMaxRepeticaoPNs').innerText = maxPNs.length > 0 ? `PNs: ${maxPNs.slice(0, 2).join(', ')}` : 'Sem repetições';

  renderCharts();
  renderTable();
}

function renderCharts() {
  // 1. Top 10 PNs Repetidos
  const dupMap = {};
  filteredData.forEach(d => {
    if (d.frequencia > 1 && d.partNumber !== 'N/A') {
      dupMap[d.partNumber] = d.frequencia;
    }
  });
  const sortedDups = Object.entries(dupMap).sort((a, b) => b[1] - a[1]).slice(0, 10);

  if (chartTopDuplicados) chartTopDuplicados.destroy();
  chartTopDuplicados = new Chart(document.getElementById('chartTopDuplicados').getContext('2d'), {
    type: 'bar',
    data: {
      labels: sortedDups.map(d => d[0]),
      datasets: [{ data: sortedDups.map(d => d[1]), backgroundColor: '#10b981', borderRadius: 4 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { x: { ticks: { color: '#94a3b8', font: { size: 9 } } }, y: { ticks: { color: '#94a3b8', stepSize: 1 } } }
    }
  });

  // 2. Proporção de Ocorrências
  const unicosCount = filteredData.filter(d => d.frequencia === 1).length;
  const dup2Count = filteredData.filter(d => d.frequencia === 2).length;
  const dup3Count = filteredData.filter(d => d.frequencia >= 3).length;

  if (chartDonutProporcao) chartDonutProporcao.destroy();
  chartDonutProporcao = new Chart(document.getElementById('chartDonutProporcao').getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: ['Únicos (1x)', 'Duplicados (2x)', 'Triplicados ou + (3x+)'],
      datasets: [{ data: [unicosCount, dup2Count, dup3Count], backgroundColor: ['#10b981', '#f59e0b', '#ef4444'], borderWidth: 0 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8' } } },
      cutout: '65%'
    }
  });

  // 3. Status de Validade dos Contratos (SUBSTITUIU GRUPO DE COMPRADORES)
  const valMap = {};
  filteredData.forEach(d => {
    const year = d.validade !== 'N/A' ? d.validade.substring(0, 4) : 'N/D';
    valMap[year] = (valMap[year] || 0) + 1;
  });
  const sortedVal = Object.entries(valMap).sort((a, b) => a[0].localeCompare(b[0]));

  if (chartValidade) chartValidade.destroy();
  chartValidade = new Chart(document.getElementById('chartValidade').getContext('2d'), {
    type: 'bar',
    data: {
      labels: sortedVal.map(v => v[0]),
      datasets: [{ data: sortedVal.map(v => v[1]), backgroundColor: '#3b82f6', borderRadius: 4 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { x: { ticks: { color: '#94a3b8', font: { size: 10 } } }, y: { ticks: { color: '#94a3b8' } } }
    }
  });

  // 4. Distribuição de Frequência de PNs
  const freqDist = { '1 Ocorrência': 0, '2 Ocorrências': 0, '3 Ocorrências': 0, '4+ Ocorrências': 0 };
  filteredData.forEach(d => {
    if (d.frequencia === 1) freqDist['1 Ocorrência']++;
    else if (d.frequencia === 2) freqDist['2 Ocorrências']++;
    else if (d.frequencia === 3) freqDist['3 Ocorrências']++;
    else freqDist['4+ Ocorrências']++;
  });

  if (chartFreqDistribucao) chartFreqDistribucao.destroy();
  chartFreqDistribucao = new Chart(document.getElementById('chartFreqDistribucao').getContext('2d'), {
    type: 'bar',
    data: {
      labels: Object.keys(freqDist),
      datasets: [{ data: Object.values(freqDist), backgroundColor: ['#10b981', '#f59e0b', '#ef4444', '#8b5cf6'], borderRadius: 4 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { x: { ticks: { color: '#94a3b8', font: { size: 10 } } }, y: { ticks: { color: '#94a3b8' } } }
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
      <td class="p-3 font-mono text-xs text-blue-400">${item.cpp}</td>
      <td class="p-3 font-sans text-xs text-slate-200">${item.textoBreve}</td>
      <td class="p-3 font-mono text-xs font-bold text-emeraldAccent">${item.partNumber}</td>
      <td class="p-3"><span class="px-2 py-0.5 rounded text-[10px] ${badgeColor}">${item.frequencia}x no contrato</span></td>
      <td class="p-3 text-xs text-slate-300">${item.comprador}</td>
      <td class="p-3 text-xs text-slate-400">${item.validade}</td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('tableStatus').innerText = `Exibindo ${displayRows.length.toLocaleString('pt-BR')} de ${filteredData.length.toLocaleString('pt-BR')} itens`;
}