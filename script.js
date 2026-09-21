<div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-6">
  
  <!-- QUADRANTE 1: Tabela de Materiais Repetidos -->
  <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 h-64 flex flex-col">
    <h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Materiais Repetidos</h3>
    <div class="overflow-y-auto flex-1">
      <table class="w-full text-left border-collapse">
        <thead>
          <tr class="border-b border-slate-700 text-[11px] text-slate-400">
            <th class="p-1">Part Number</th>
            <th class="p-1">CPP</th>
            <th class="p-1">Qtd</th>
            <th class="p-1">% Total</th>
          </tr>
        </thead>
        <tbody id="tabelaMateriaisRepetidosBody">
          <!-- Preenchido via JS -->
        </tbody>
      </table>
    </div>
  </div>

  <!-- QUADRANTE 2: Proporção (Mantido) -->
  <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 h-64 flex flex-col">
    <h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Proporção de Ocorrências em Contrato</h3>
    <div class="flex-1 relative">
      <canvas id="chartDonutProporcao"></canvas>
    </div>
  </div>

  <!-- QUADRANTE 3: KPI Data de Vencimento -->
  <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 h-48 flex flex-col justify-center">
    <h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Status de Validade (KPI)</h3>
    <div class="mt-2">
      <div id="kpiVencimentoDias" class="text-3xl font-extrabold text-blue-400">-- dias restantes</div>
      <div id="kpiVencimentoData" class="text-sm text-slate-400 mt-1">Expira em: --/--/----</div>
      <div id="kpiVencimentoPct" class="text-xs font-semibold text-emerald-400 mt-2">--% restante</div>
    </div>
  </div>

  <!-- QUADRANTE 4: Saldo, Disponível e Consumido -->
  <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 h-48 flex flex-col justify-center">
    <h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Métricas de Saldo e Consumo</h3>
    <div class="grid grid-cols-3 gap-2 mt-2 text-center">
      <div class="bg-slate-800/60 p-2 rounded-lg">
        <span class="block text-[10px] text-slate-400 uppercase">Consumido</span>
        <span id="metricConsumido" class="text-sm font-bold text-amber-400">0 (0%)</span>
      </div>
      <div class="bg-slate-800/60 p-2 rounded-lg">
        <span class="block text-[10px] text-slate-400 uppercase">Saldo Contrato</span>
        <span id="metricSaldo" class="text-sm font-bold text-blue-400">0 (0%)</span>
      </div>
      <div class="bg-slate-800/60 p-2 rounded-lg">
        <span class="block text-[10px] text-slate-400 uppercase">Disponível</span>
        <span id="metricDisponivel" class="text-sm font-bold text-emerald-400">0 (0%)</span>
      </div>
    </div>
  </div>

</div>