let rawData = [];
let filteredData = [];
let chartDonutProporcao = null;

function setElementText(id, text) {
  const el = document.getElementById(id);
  if (el) el.innerText = text;
}

function formatExcelDate(excelDate) {
  if (!excelDate || excelDate === 'N/A') return 'N/A';
  if (typeof excelDate === 'string' && (excelDate.includes('/') || excelDate.includes('-'))) {
    const d = new Date(excelDate);
    return !isNaN(d.getTime()) ? d.toLocaleDateString('pt-BR') : excelDate;
  }
  const num = Number(excelDate);
  if (!isNaN(num) && num > 0 && typeof XLSX !== 'undefined') {
    const dateObj = XLSX.SSF.parse_date_code(num);
    if (dateObj) {
      return `${String(dateObj.d).padStart(2, '0')}/${String(dateObj.m).padStart(2, '0')}/${dateObj.y}`;
    }
  }
  return String(excelDate);
}

function parseDateBR(dateStr) {
  if (!dateStr || dateStr === 'N/A') return null;
  const parts = dateStr.split('/');
  if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
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
    const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    processDataset(json);
  };
  reader.readAsArrayBuffer(file);
}

function processDataset(json) {
  const parsed = json.map(row => {
    const textoBreve = String(row['Texto breve'] || row['Coluna1'] || '');
    const pn = extractPartNumber(textoBreve);
    const matVal = row['Material'] || row['Centro'] || row['CPP'] || 'N/A';

    return {
      material: String(matVal).trim(),
      cpp: String(matVal).trim(),
      textoBreve: textoBreve,
      partNumber: pn,
      comprador: String(row['Grupo de compradores'] || row['Comprador'] || 'N/A').trim(),
      validade: formatExcelDate(row['Fim da validade'] || row['Validade'] || 'N/A'),
      condicaoPagamento: String(row['Condições pagamento'] || row['Condição de Pagamento'] || 'N/A').trim(),
      incoterms: String(row['Incoterms'] || 'N/A').trim()
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
  const filterComprador = document.getElementById('filterComprador');
  const filterDuplicidade = document.getElementById('filterDuplicidade');

  const searchTerm = searchInput ? searchInput.value.toLowerCase() : '';
  const comprador = filterComprador ? filterComprador.value : 'ALL';
  const duplicidade = filterDuplicidade ? filterDuplicidade.value : 'ALL';

  filteredData = rawData.filter(item => {
    const matchSearch = item.textoBreve.toLowerCase().includes(searchTerm) ||
                        item.partNumber.toLowerCase().includes(searchTerm) ||
                        item.material.toLowerCase().includes(searchTerm);

    const matchComp = comprador === 'ALL' || item.comprador === comprador;

    let matchDup = true;
    if (duplicidade === 'UNICOS') matchDup = item.frequencia === 1;
    if (duplicidade === 'DUPLICADOS') matchDup = item.frequencia === 2;
    if (duplicidade === 'TRIPLICADOS') matchDup = item.frequencia >= 3;

    return matchSearch && matchComp && matchDup;
  });

  updateDashboardUI();
}

function updateDashboardUI() {
  setElementText('filteredCountLabel', `${filteredData.length.toLocaleString('pt-BR')} de ${rawData.length.toLocaleString('pt-BR')} itens`);

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

  renderTabelaMateriaisRepetidos();
  renderKpiVencimento();
  renderCondicoesComerciais();
  renderChartProporcao();
  renderTable();
}

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
        mapPns[d.partNumber] = { partNumber: d.partNumber, material: d.material, textoBreve: d.textoBreve, count: 0 };
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
      item.material.toLowerCase().includes(term) ||
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
      <td class="p-1.5 font-mono text-xs text-blue-400">${item.material}</td>
      <td class="p-1.5 text-xs text-slate-300 truncate max-w-[120px]" title="${item.textoBreve}">${item.textoBreve}</td>
      <td class="p-1.5 font-mono text-xs text-emerald-400">${item.partNumber}</td>
      <td class="p-1.5 text-xs font-bold text-amber-400 text-right">${item.count.toLocaleString('pt-BR')}</td>
    `;
    container.appendChild(tr);
  });
}

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
        padding: { top: 10, bottom: 10 }
      },
      plugins: { 
        legend: { 
          position: 'bottom', 
          labels: { 
            color: '#94a3b8',
            font: { size: 10 },
            boxWidth: 10,
            padding: 8
          } 
        } 
      },
      cutout: '65%'
    }
  });
}

function renderKpiVencimento() {
  const dates = filteredData.map(d => parseDateBR(d.validade)).filter(d => d !== null);
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

// QUADRANTE 4: CONDIÇÕES COMERCIAIS
function renderCondicoesComerciais() {
  const condicoes = [...new Set(filteredData.map(d => d.condicaoPagamento).filter(v => v !== 'N/A'))];
  const fretes = [...new Set(filteredData.map(d => d.incoterms).filter(v => v !== 'N/A'))];

  // Exibe o Saldo Fixo do Contrato apenas se a planilha tiver sido carregada
  const saldoExibicao = rawData.length > 0 ? 'R$ 22.767.365,43' : 'R$ 0,00';

  setElementText('metricSaldoContrato', saldoExibicao);
  setElementText('metricCondicaoPagamento', condicoes.length > 0 ? condicoes.join(', ') : '--');
  setElementText('metricIncotermsFrete', fretes.length > 0 ? fretes.join(', ') : '--');
}

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
      <td class="p-3 font-mono text-[11px] text-blue-400">${item.material}</td>
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