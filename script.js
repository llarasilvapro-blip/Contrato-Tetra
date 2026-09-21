let rawData = [];
let filteredData = [];

// Instâncias dos gráficos (Chart.js)
let chartDonutProporcao = null;

// Converte datas seriais do Excel (ex: 46234) para DD/MM/AAAA
function formatExcelDate(excelDate) {
  if (!excelDate || excelDate === 'N/A') return 'N/A';
  
  if (typeof excelDate === 'string' && (excelDate.includes('/') || excelDate.includes('-'))) {
    const d = new Date(excelDate);
    if (!isNaN(d.getTime())) {
      return d.toLocaleDateString('pt-BR');
    }
    return excelDate;
  }

  const num = Number(excelDate);
  if (!isNaN(num) && num > 0) {
    const dateObj = XLSX.SSF.parse_date_code(num);
    if (dateObj) {
      const day = String(dateObj.d).padStart(2, '0');
      const month = String(dateObj.m).padStart(2, '0');
      const year = dateObj.y;
      return `${day}/${month}/${year}`;
    }
  }

  return String(excelDate);
}

// Converte texto em data para cálculos de KPI
function parseDateBR(dateStr) {
  if (!dateStr || dateStr === 'N/A') return null;
  const parts = dateStr.split('/');
  if (parts.length === 3) {
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  const d = new Date(dateStr);
  return isNaN(d.getTime()) ? null : d;
}

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
  document.getElementById('searchCPP').addEventListener('input', applyFilters);
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
    
    const rawCPP = row['CPP'] || row['Centro'] || row['Plant'] || row['Centro de Custo'] || row['Centro Lucro'] || 'N/A';
    const rawValidade = row['Fim da validade'] || row['Validade'] || row['Fim Validade'] || 'N/A';
    const rawInicio = row['Inicio validade'] || row['Inicio Validade'] || row['Data Inicio'] || 'N/A';

    return {
      material: row['Material'] || 'N/A',
      cpp: String(rawCPP).trim(),
      textoBreve: textoBreve,
      partNumber: pn,
      comprador: String(row['Grupo de compradores'] || 'N/A').trim(),
      validade: formatExcelDate(rawValidade),
      dataInicio: formatExcelDate(rawInicio),
      
      // Quantidades operacionais
      qtdContratada: Number(row['Qtd Contratada'] || row['Qtd Total'] || 0),
      qtdConsumida: Number(row['Qtd Consumida'] || row['Qtd Utilizada'] || 0),
      qtdDisponivel: Number(row['Qtd Disponivel'] || row['Saldo Qtd'] || 0)
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
  const compradores = [...new Set(rawData.map(d => d.comprador))].sort();
  const compSelect = document.getElementById('filterComprador');
  compSelect.innerHTML = '<option value="ALL">Todos os Compradores</option>';
  compradores.forEach(c => {
    compSelect.innerHTML += `<option value="${c}">${c}</option>`;
  });
}

function applyFilters() {
  const searchTerm = document.getElementById('searchInput').value.toLowerCase();
  const cppSearchTerm = document.getElementById('searchCPP').value.toLowerCase();
  const comprador = document.getElementById('filterComprador').value;
  const duplicidade = document.getElementById('filterDuplicidade').value;

  filteredData = rawData.filter(item => {
    const matchSearch = item.textoBreve.toLowerCase().includes(searchTerm) ||
                        item.partNumber.toLowerCase().includes(searchTerm) ||
                        String(item.material).toLowerCase().includes(searchTerm);

    const matchCPP = item.cpp.toLowerCase().includes(cppSearchTerm);
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

  // Atualização dos KPI's Centrais
  const totalItens = filteredData.length;
  const pnsUnicos = new Set(filteredData.filter(d => d.partNumber !== 'N/A').map(d => d.partNumber)).size;
  const duplicadosList = filteredData.filter(d => d.frequencia > 1 && d.partNumber !== 'N/A');
  const pnsDuplicadosUnicos = new Set(duplicadosList.map(d => d.partNumber)).size;

  document.getElementById('kpiTotalItens').innerText = `${totalItens.toLocaleString('pt-BR')} ITENS`;
  document.getElementById('kpiUnicos').innerText = `${pnsUnicos.toLocaleString('pt-BR')} PNs`;
  document.getElementById('kpiDuplicados').innerText = `${pnsDuplicadosUnicos.toLocaleString('pt-BR')} PNs`;

  // Renderizar componentes atualizados conforme os requisitos
  renderTabelaMateriaisRepetidos();
  renderKpiVencimento();
  renderMétricasContrato();
  renderCharts();
  renderTable();
}

// ----------------------------------------------------
// 1. TABELA DE MATERIAIS REPETIDOS (com pesquisa PN e CPP)
// ----------------------------------------------------
function renderTabelaMateriaisRepetidos() {
  const container = document.getElementById('tabelaMateriaisRepetidosBody');
  if (!container) return;
  container.innerHTML = '';

  const totalLinhas = filteredData.length;
  const mapPns = {};

  filteredData.forEach(d => {
    if (d.partNumber !== 'N/A') {
      if (!mapPns[d.partNumber]) {
        mapPns[d.partNumber] = { partNumber: d.partNumber, cpp: d.cpp, count: 0 };
      }
      mapPns[d.partNumber].count++;
    }
  });

  const arrayPNs = Object.values(mapPns)
    .filter(item => item.count > 1)
    .sort((a, b) => b.count - a.count);

  arrayPNs.forEach(item => {
    const pct = totalLinhas > 0 ? ((item.count / totalLinhas) * 100).toFixed(2) : 0;
    const tr = document.createElement('tr');
    tr.className = "border-b border-slate-700/50 hover:bg-slate-800/40";
    tr.innerHTML = `
      <td class="p-2 font-mono text-xs text-emerald-400">${item.partNumber}</td>
      <td class="p-2 font-mono text-xs text-blue-400">${item.cpp}</td>
      <td class="p-2 text-xs text-slate-200">${item.count.toLocaleString('pt-BR')}</td>
      <td class="p-2 text-xs font-semibold text-amber-400">${pct}%</td>
    `;
    container.appendChild(tr);
  });
}

// ----------------------------------------------------
// 2. QUADRANTE KPI DA DATA DE VENCIMENTO
// ----------------------------------------------------
function renderKpiVencimento() {
  const dates = filteredData
    .map(d => parseDateBR(d.validade))
    .filter(d => d !== null);

  if (dates.length === 0) {
    if (document.getElementById('kpiVencimentoDias')) {
      document.getElementById('kpiVencimentoDias').innerText = "N/D";
      document.getElementById('kpiVencimentoPct').innerText = "0%";
    }
    return;
  }

  // Obter a menor data de vencimento entre os dados filtrados
  const dataVencimentoMaisProxima = new Date(Math.min.apply(null, dates));
  const hoje = new Date();
  
  const diffTime = dataVencimentoMaisProxima - hoje;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  // Estimativa de vigência total (por omissão 365 dias ou calculada do início do contrato)
  const totalDiasContrato = 365;
  const pctRestante = Math.max(0, Math.min(100, ((diffDays / totalDiasContrato) * 100))).toFixed(1);
  const pctDecorrente = (100 - pctRestante).toFixed(1);

  if (document.getElementById('kpiVencimentoDias')) {
    document.getElementById('kpiVencimentoDias').innerText = `${diffDays} dias restantes`;
    document.getElementById('kpiVencimentoData').innerText = `Expira em: ${dataVencimentoMaisProxima.toLocaleDateString('pt-BR')}`;
    document.getElementById('kpiVencimentoPct').innerText = `${pctRestante}% restante (${pctDecorrente}% decorrido)`;
  }
}

// ----------------------------------------------------
// 3. QUADRANTE SALDO DO CONTRATO, DISPONÍVEL E CONSUMIDO
// ----------------------------------------------------
function renderMétricasContrato() {
  let totalConsumido = 0;
  let totalDisponivel = 0;
  let totalContratado = 0;

  filteredData.forEach(d => {
    totalConsumido += d.qtdConsumida;
    totalDisponivel += d.qtdDisponivel;
    totalContratado += d.qtdContratada;
  });

  if (totalContratado === 0) {
    totalContratado = totalConsumido + totalDisponivel; // Fallback caso não venha explicito
  }

  const saldoContrato = totalContratado - totalConsumido;

  const pctConsumido = totalContratado > 0 ? ((totalConsumido / totalContratado) * 100).toFixed(1) : 0;
  const pctSaldo = totalContratado > 0 ? ((saldoContrato / totalContratado) * 100).toFixed(1) : 0;
  const pctDisponivel = totalContratado > 0 ? ((totalDisponivel / totalContratado) * 100).toFixed(1) : 0;

  // Atualizar Elementos no HTML
  if (document.getElementById('metricConsumido')) {
    document.getElementById('metricConsumido').innerText = `${totalConsumido.toLocaleString('pt-BR')} (${pctConsumido}%)`;
    document.getElementById('metricSaldo').innerText = `${saldoContrato.toLocaleString('pt-BR')} (${pctSaldo}%)`;
    document.getElementById('metricDisponivel').innerText = `${totalDisponivel.toLocaleString('pt-BR')} (${pctDisponivel}%)`;
  }
}

function renderCharts() {
  // Gráfico Donut de Proporção de Duplicados
  const unicosCount = filteredData.filter(d => d.frequencia === 1).length;
  const dup2Count = filteredData.filter(d => d.frequencia === 2).length;
  const dup3Count = filteredData.filter(d => d.frequencia >= 3).length;

  const elChart = document.getElementById('chartDonutProporcao');
  if (elChart) {
    if (chartDonutProporcao) chartDonutProporcao.destroy();
    chartDonutProporcao = new Chart(elChart.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: ['Únicos (1x)', 'Duplicados (2x)', 'Triplicados ou + (3x+)'],
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
        cutout: '65%'
      }
    });
  }
}

function renderTable() {
  const tbody = document.getElementById('tableBody');
  if (!tbody) return;
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

  const statusEl = document.getElementById('tableStatus');
  if (statusEl) {
    statusEl.innerText = `Exibindo ${displayRows.length.toLocaleString('pt-BR')} de ${filteredData.length.toLocaleString('pt-BR')} itens`;
  }
}