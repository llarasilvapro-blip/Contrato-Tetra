let rawData = [];
let filteredData = [];

// Instância do Gráfico
let chartDonutProporcao = null;

// Helper seguro para atualizar o texto do DOM
function setElementText(id, text) {
  const el = document.getElementById(id);
  if (el) el.innerText = text;
}

// Converte datas seriais do Excel para DD/MM/AAAA
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
  if (!isNaN(num) && num > 0 && typeof XLSX !== 'undefined') {
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
  if (fileInput) fileInput.addEventListener('change', handleFileUpload);

  const searchInput = document.getElementById('searchInput');
  if (searchInput) searchInput.addEventListener('input', applyFilters);

  const searchCPP = document.getElementById('searchCPP');
  if (searchCPP) searchCPP.addEventListener('input', applyFilters);

  // Pesquisa dinâmica na tabela de Materiais Repetidos
  const searchRepetidosInput = document.getElementById('searchRepetidosInput');
  if (searchRepetidosInput) searchRepetidosInput.addEventListener('input', renderTabelaMateriaisRepetidos);

  const filterComprador = document.getElementById('filterComprador');
  if (filterComprador) filterComprador.addEventListener('change', applyFilters);

  const filterDuplicidade = document.getElementById('filterDuplicidade');
  if (filterDuplicidade) filterDuplicidade.addEventListener('change', applyFilters);
});

function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  setElementText('fileInfoLabel', `Base carregada: (${file.name})`);

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

    return {
      material: row['Material'] || 'N/A',
      cpp: String(rawCPP).trim(),
      textoBreve: textoBreve,
      partNumber: pn,
      comprador: String(row['Grupo de compradores'] || 'N/A').trim(),
      validade: formatExcelDate(rawValidade),
      
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
  const compSelect = document.getElementById('filterComprador');
  if (!compSelect) return;

  const compradores = [...new Set(rawData.map(d => d.comprador))].sort();
  compSelect.innerHTML = '<option value="ALL">Todos os Compradores</option>';
  compradores.forEach(c => {
    compSelect.innerHTML += `<option value="${c}">${c}</option>`;
  });
}

function applyFilters() {
  const searchInput = document.getElementById('searchInput');
  const searchCPP = document.getElementById('searchCPP');
  const filterComprador = document.getElementById('filterComprador');
  const filterDuplicidade = document.getElementById('filterDuplicidade');

  const searchTerm = searchInput ? searchInput.value.toLowerCase() : '';
  const cppSearchTerm = searchCPP ? searchCPP.value.toLowerCase() : '';
  const comprador = filterComprador ? filterComprador.value : 'ALL';
  const duplicidade = filterDuplicidade ? filterDuplicidade.value : 'ALL';

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
  setElementText('filteredCountLabel', `${filteredData.length.toLocaleString('pt-BR')} de ${rawData.length.toLocaleString('pt-BR')} itens`);

  // Métricas do Topo
  const totalItens = filteredData.length;
  const pnsUnicos = new Set(filteredData.filter(d => d.partNumber !== 'N/A').map(d => d.partNumber)).size;
  const duplicadosList = filteredData.filter(d => d.frequencia > 1 && d.partNumber !== 'N/A');
  const pnsDuplicadosUnicos = new Set(duplicadosList.map(d => d.partNumber)).size;

  let maxRep = 0;
  filteredData.forEach(d => { if (d.frequencia > maxRep) maxRep = d.frequencia; });

  setElementText('kpiTotalItens', `${totalItens.toLocaleString('pt-BR')} ITENS`);
  setElementText('kpiUnicos', `${pnsUnicos.toLocaleString('pt-BR')} PNs`);
  setElementText('kpiDuplicados', `${pnsDuplicadosUnicos.toLocaleString('pt-BR')} PNs`);
  setElementText('kpiMaxRepeticao', maxRep > 1 ? `${maxRep} VEZES` : '1 VEZ');

  // Renderizar Quadrantes
  renderTabelaMateriaisRepetidos();
  renderKpiVencimento();
  renderMétricasContrato();
  renderChartProporcao();
  renderTable();
}

// QUADRANTE 1: Tabela de Part Numbers Repetidos (CPP, Texto Breve, Part Number e Qtd)
function renderTabelaMateriaisRepetidos() {
  const container = document.getElementById('tabelaMateriaisRepetidosBody');
  if (!container) return;
  container.innerHTML = '';

  const searchRepetidosInput = document.getElementById('searchRepetidosInput');
  const term = searchRepetidosInput ? searchRepetidosInput.value.toLowerCase().trim() : '';

  const mapPns = {};

  filteredData.forEach(d => {
    if (d.partNumber !== 'N/A') {
      if (!mapPns[d.partNumber]) {
        mapPns[d.partNumber] = { 
          partNumber: d.partNumber, 
          cpp: d.cpp, 
          textoBreve: d.textoBreve, 
          count: 0 
        };
      }
      mapPns[d.partNumber].count++;
    }
  });

  let arrayPNs = Object.values(mapPns)
    .filter(item => item.count > 1)
    .sort((a, b) => b.count - a.count);

  if (term !== '') {
    arrayPNs = arrayPNs.filter(item => 
      item.partNumber.toLowerCase().includes(term) || 
      item.cpp.toLowerCase().includes(term) ||
      item.textoBreve.toLowerCase().includes(term)
    );
  }

  if (arrayPNs.length === 0) {
    container.innerHTML = `<tr><td colspan="4" class="p-3 text-center text-xs text-slate-500">Nenhum item repetido encontrado.</td></tr>`;
    return;
  }

  arrayPNs.forEach(item => {
    const tr = document.createElement('tr');
    tr.className = "border-b border-slate-700/50 hover:bg-slate-800/40";
    tr.innerHTML = `
      <td class="p-1.5 font-mono text-xs text-blue-400">${item.cpp}</td>
      <td class="p-1.5 text-xs text-slate-300 truncate max-w-[140px]" title="${item.textoBreve}">${item.textoBreve}</td>
      <td class="p-1.5 font-mono text-xs text-emerald-400">${item.partNumber}</td>
      <td class="p-1.5 text-xs font-bold text-amber-400 text-right">${item.count.toLocaleString('pt-BR')}</td>
    `;
    container.appendChild(tr);
  });
}

// QUADRANTE 2: Proporção (Donut Chart com Layout e Legenda Ajustados)
function renderChartProporcao() {
  const elChart = document.getElementById('chartDonutProporcao');
  if (!elChart || typeof Chart === 'undefined') return;

  const unicosCount = filteredData.filter(d => d.frequencia === 1).length;
  const dup2Count = filteredData.filter(d => d.frequencia === 2).length;
  const dup3Count = filteredData.filter(d => d.frequencia >= 3).length;

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
      layout: {
        padding: {
          top: 0,
          bottom: 5
        }
      },
      plugins: { 
        legend: { 
          position: 'bottom', 
          labels: { 
            color: '#94a3b8',
            font: { size: 10 },
            boxWidth: 12,
            padding: 8
          } 
        } 
      },
      cutout: '70%'
    }
  });
}

// QUADRANTE 3: Status de Validade
function renderKpiVencimento() {
  const dates = filteredData
    .map(d => parseDateBR(d.validade))
    .filter(d => d !== null);

  if (dates.length === 0) {
    setElementText('kpiVencimentoDias', 'N/D');
    setElementText('kpiVencimentoData', 'Expira em: N/D');
    setElementText('kpiVencimentoPct', '0% restante');
    return;
  }

  const dataVencimentoMaisProxima = new Date(Math.min.apply(null, dates));
  const hoje = new Date();
  
  const diffTime = dataVencimentoMaisProxima - hoje;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  const totalDiasContrato = 365;
  const pctRestante = Math.max(0, Math.min(100, ((diffDays / totalDiasContrato) * 100))).toFixed(1);
  const pctDecorrente = (100 - pctRestante).toFixed(1);

  setElementText('kpiVencimentoDias', `${diffDays} dias restantes`);
  setElementText('kpiVencimentoData', `Expira em: ${dataVencimentoMaisProxima.toLocaleDateString('pt-BR')}`);
  setElementText('kpiVencimentoPct', `${pctRestante}% restante (${pctDecorrente}% decorrido)`);
}

// QUADRANTE 4: Métricas de Saldo e Consumo
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
    totalContratado = totalConsumido + totalDisponivel;
  }

  const saldoContrato = totalContratado - totalConsumido;

  const pctConsumido = totalContratado > 0 ? ((totalConsumido / totalContratado) * 100).toFixed(1) : '0.0';
  const pctSaldo = totalContratado > 0 ? ((saldoContrato / totalContratado) * 100).toFixed(1) : '0.0';
  const pctDisponivel = totalContratado > 0 ? ((totalDisponivel / totalContratado) * 100).toFixed(1) : '0.0';

  setElementText('metricConsumido', `${totalConsumido.toLocaleString('pt-BR')} (${pctConsumido}%)`);
  setElementText('metricSaldo', `${saldoContrato.toLocaleString('pt-BR')} (${pctSaldo}%)`);
  setElementText('metricDisponivel', `${totalDisponivel.toLocaleString('pt-BR')} (${pctDisponivel}%)`);
}

// TABELA PRINCIPAL
function renderTable() {
  const tbody = document.getElementById('tableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const displayRows = filteredData.slice(0, 100);

  displayRows.forEach(item => {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-slate-800/50 transition border-b border-slate-800/40";

    let badgeColor = "bg-slate-800 text-slate-400";
    if (item.frequencia === 2) badgeColor = "bg-amber-500/20 text-amber-400 border border-amber-500/30";
    if (item.frequencia >= 3) badgeColor = "bg-red-500/20 text-red-400 border border-red-500/30 font-bold";

    tr.innerHTML = `
      <td class="p-3 font-mono text-[11px] text-slate-400">${item.material}</td>
      <td class="p-3 font-mono text-xs text-blue-400">${item.cpp}</td>
      <td class="p-3 font-sans text-xs text-slate-200">${item.textoBreve}</td>
      <td class="p-3 font-mono text-xs font-bold text-emerald-400">${item.partNumber}</td>
      <td class="p-3"><span class="px-2 py-0.5 rounded text-[10px] ${badgeColor}">${item.frequencia}x no contrato</span></td>
      <td class="p-3 text-xs text-slate-300">${item.comprador}</td>
      <td class="p-3 text-xs text-slate-400">${item.validade}</td>
    `;
    tbody.appendChild(tr);
  });

  setElementText('tableStatus', `Exibindo ${displayRows.length.toLocaleString('pt-BR')} de ${filteredData.length.toLocaleString('pt-BR')} itens`);
}