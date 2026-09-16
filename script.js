let globalData = [];

// Função com Expressão Regular para extrair o Part Number do "Texto Breve"
function extractPartNumber(text) {
  if (!text || typeof text !== 'string') return "N/A";
  
  // Remove marcas comuns ao final da string para evitar falsos positivos
  const cleaned = text.replace(/\b(TETRA PAK|ALFA LAVAL)\b/gi, '').trim();
  
  // Extrai tokens contendo dígitos numericos
  const tokens = cleaned.split(/\s+/);
  const digitTokens = tokens.filter(token => /\d/.test(token));
  
  if (digitTokens.length > 0) {
    return digitTokens.join(" ");
  }
  
  return "N/A";
}

// Manipulação do carregamento do arquivo Excel via SheetJS
document.getElementById('excelFileInput').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Seleciona preferencialmente a aba "Sheet1" ou a primeira disponível
    const sheetName = workbook.SheetNames.includes('Sheet1') ? 'Sheet1' : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    const rawJson = XLSX.utils.sheet_to_json(worksheet);
    
    processData(rawJson);
  };
  reader.readAsArrayBuffer(file);
});

// Processamento dos dados e contagem de frequência de Part Numbers
function processData(rawJson) {
  // Extrai o Part Number
  const parsed = rawJson.map(row => {
    const textoBreve = row['Texto breve'] || row['Coluna1'] || '';
    const pn = extractPartNumber(textoBreve);
    return {
      material: row['Material'] || 'N/A',
      textoBreve: textoBreve,
      partNumber: pn,
      preco: row['Preço líquido'] || 0,
      validade: row['Fim da validade'] ? new Date(row['Fim da validade']).toLocaleDateString('pt-BR') : 'N/A'
    };
  });

  // Mapeia frequência dos Part Numbers
  const freqMap = {};
  parsed.forEach(item => {
    if (item.partNumber !== 'N/A') {
      freqMap[item.partNumber] = (freqMap[item.partNumber] || 0) + 1;
    }
  });

  // Atribui frequência individual
  globalData = parsed.map(item => ({
    ...item,
    frequencia: freqMap[item.partNumber] || 1
  }));

  updateMetrics();
  renderTable(globalData);

  // Exibe as seções na interface
  document.getElementById('metricsSection').classList.remove('hidden');
  document.getElementById('filterSection').classList.remove('hidden');
  document.getElementById('tableSection').classList.remove('hidden');
}

// Atualiza o painel de indicadores (KPIs)
function updateMetrics() {
  const total = globalData.length;
  const unicos = globalData.filter(d => d.frequencia === 1).length;
  const duplicados = globalData.filter(d => d.frequencia === 2).length;
  const triplicados = globalData.filter(d => d.frequencia >= 3).length;

  document.getElementById('metricTotal').innerText = total.toLocaleString('pt-BR');
  document.getElementById('metricUnicos').innerText = unicos.toLocaleString('pt-BR');
  document.getElementById('metricDuplicados').innerText = duplicados.toLocaleString('pt-BR');
  document.getElementById('metricTriplicados').innerText = triplicados.toLocaleString('pt-BR');
}

// Renderiza a tabela HTML com base nos dados filtrados
function renderTable(data) {
  const tbody = document.getElementById('tableBody');
  tbody.innerHTML = '';

  data.forEach(item => {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-slate-50 border-b border-slate-100";

    // Badge de Ocorrência
    let badgeClass = "bg-slate-100 text-slate-600";
    if (item.frequencia === 2) badgeClass = "bg-amber-100 text-amber-800 font-semibold";
    if (item.frequencia >= 3) badgeClass = "bg-red-100 text-red-800 font-bold";

    tr.innerHTML = `
      <td class="p-3 font-mono text-xs">${item.material}</td>
      <td class="p-3 font-medium text-slate-800">${item.textoBreve}</td>
      <td class="p-3 font-mono font-semibold text-blue-600">${item.partNumber}</td>
      <td class="p-3"><span class="px-2 py-1 rounded-full text-xs ${badgeClass}">${item.frequencia}x</span></td>
      <td class="p-3">${item.preco ? item.preco.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : 'R$ 0,00'}</td>
      <td class="p-3 text-xs text-slate-500">${item.validade}</td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('rowCount').innerText = `Exibindo ${data.length.toLocaleString('pt-BR')} registros`;
}

// Eventos de Busca e Filtro
document.getElementById('searchInput').addEventListener('input', applyFilters);
document.getElementById('freqFilter').addEventListener('change', applyFilters);

function applyFilters() {
  const searchTerm = document.getElementById('searchInput').value.toLowerCase();
  const freqValue = document.getElementById('freqFilter').value;

  const filtered = globalData.filter(item => {
    const matchesSearch = 
      item.textoBreve.toLowerCase().includes(searchTerm) ||
      item.partNumber.toLowerCase().includes(searchTerm) ||
      String(item.material).toLowerCase().includes(searchTerm);

    let matchesFreq = true;
    if (freqValue === 'UNICOS') matchesFreq = item.frequencia === 1;
    if (freqValue === 'DUPLICADOS') matchesFreq = item.frequencia === 2;
    if (freqValue === 'TRIPLICADOS') matchesFreq = item.frequencia >= 3;

    return matchesSearch && matchesFreq;
  });

  renderTable(filtered);
}