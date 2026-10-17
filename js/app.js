/**
 * ============================================================================
 * APLICACIÓN CONTABLE ALTUS (iPhone PWA & Mac Desktop)
 * Conectada a Google Sheets + Motor NLP SAT Guatemala + Exportación Excel (.xlsx)
 * ============================================================================
 */

(function () {
  const STORAGE_KEY = 'altus_contabilidad_state_v2';
  const SHEETS_CONFIG_KEY = 'altus_sheets_webhook_config_v1';

  let state = loadState();
  let currentDraft = null;
  let currentQueryResult = null;
  let activeTab = 'asistente';
  let journalFilter = { search: '', month: 'todos', bank: 'todos', type: 'todos' };

  function loadState() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed && parsed.journal && parsed.journal.length >= 280) {
          return parsed;
        }
      }
    } catch (e) {
      console.warn('No se pudo cargar estado previo, usando datos auditados iniciales:', e);
    }
    return JSON.parse(JSON.stringify(window.ALTUS_INITIAL_DATA));
  }

  function saveState() {
    try {
      recalculateBalances();
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
      console.error('Error guardando estado:', e);
    }
  }

  function getSheetsConfig() {
    try {
      return JSON.parse(localStorage.getItem(SHEETS_CONFIG_KEY) || '{}');
    } catch (e) {
      return {};
    }
  }

  function saveSheetsConfig(cfg) {
    localStorage.setItem(SHEETS_CONFIG_KEY, JSON.stringify(cfg));
    updateSyncStatusUI();
  }

  function formatQ(num) {
    const val = Number(num || 0);
    return 'Q ' + val.toLocaleString('es-GT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function recalculateBalances() {
    let mon = 0;
    let aho = 0;
    for (const j of state.journal) {
      const bruto = Number(j.montoBruto || 0);
      const isAhorroBank = j.banco === 'BI-Ahorro';
      const isTraspasoAhorro = j.cuenta === '2102 - Provisiones Laborales' || j.tipo === 'Traspaso Ahorro';
      const origType = j.tipoOriginal || j.tipo;

      if (isAhorroBank) {
        aho = +(aho - bruto).toFixed(2);
      } else {
        if (origType === 'Gasto' || j.tipo === 'Gasto' || isTraspasoAhorro) {
          mon = +(mon - bruto).toFixed(2);
        } else {
          mon = +(mon + bruto).toFixed(2);
        }
        if (isTraspasoAhorro) {
          aho = +(aho + bruto).toFixed(2);
        }
      }
      j.saldoMonetaria = mon;
      j.saldoAhorro = aho;
    }
    state.metadata.saldoActualMonetaria = mon;
    state.metadata.saldoActualAhorro = aho;
    state.metadata.totalPatrimonioLiquido = +(mon + aho).toFixed(2);
  }

  // Push transaction to Google Sheets Webhook if configured
  async function syncEntryToGoogleSheets(entry) {
    const cfg = getSheetsConfig();
    if (!cfg.webhookUrl) {
      return { synced: false, reason: 'Modo Local Activo (Listo para vincular URL Apps Script)' };
    }
    try {
      const response = await fetch(cfg.webhookUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'appendEntry', entry })
      });
      cfg.lastSync = new Date().toISOString();
      saveSheetsConfig(cfg);
      return { synced: true };
    } catch (err) {
      console.warn('Error sincronizando con Google Sheets:', err);
      return { synced: false, reason: err.message };
    }
  }

  // Render Header Balances & Day 27 Alert
  function renderHeader() {
    recalculateBalances();
    const mon = state.metadata.saldoActualMonetaria;
    const aho = state.metadata.saldoActualAhorro;
    const total = state.metadata.totalPatrimonioLiquido;

    document.getElementById('header-saldo-monetaria').textContent = formatQ(mon);
    document.getElementById('header-saldo-ahorro').textContent = formatQ(aho);
    document.getElementById('header-saldo-total').textContent = formatQ(total);
    document.getElementById('header-entries-count').textContent = `${state.journal.length} asientos auditados`;

    // Calculate days until the 27th
    const now = new Date();
    const currentDay = now.getDate();
    let daysTo27 = 27 - currentDay;
    if (daysTo27 < 0) {
      const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 27);
      daysTo27 = Math.ceil((nextMonth - now) / (1000 * 60 * 60 * 24));
    }

    const bannerEl = document.getElementById('day27-alert-banner');
    if (bannerEl) {
      const badgeText = daysTo27 === 0
        ? '¡HOY ES DÍA 27! PAGO DE NÓMINA Y PROVISIONES'
        : `AVISO NÓMINA DÍA 27 (En ${daysTo27} días)`;
      document.getElementById('day27-countdown-badge').textContent = badgeText;
    }

    updateSyncStatusUI();
  }

  function updateSyncStatusUI() {
    const cfg = getSheetsConfig();
    const dot = document.getElementById('sheets-status-dot');
    const label = document.getElementById('sheets-status-label');
    if (!dot || !label) return;
    if (cfg.webhookUrl) {
      dot.className = 'w-2.5 h-2.5 rounded-full bg-emerald-400 animate-pulse';
      label.textContent = 'Google Sheets Conectado';
    } else {
      dot.className = 'w-2.5 h-2.5 rounded-full bg-amber-400';
      label.textContent = 'Base Auditada Activa (Conectar Sheets)';
    }
  }

  // Handle Natural Language Input
  function handleProcessNLP(customText) {
    const inputEl = document.getElementById('nlp-input');
    const text = customText !== undefined ? customText : (inputEl ? inputEl.value : '');
    if (!text || !text.trim()) return;
    if (inputEl && customText !== undefined) {
      inputEl.value = customText;
    }

    const parsed = window.AltusNLPEngine.parseNaturalInput(text, state);
    if (!parsed) return;

    if (parsed.kind === 'query_answer') {
      currentQueryResult = parsed;
      currentDraft = null;
    } else {
      currentDraft = parsed;
      currentQueryResult = null;
    }
    renderNLPPreview();
  }

  function renderNLPPreview() {
    const container = document.getElementById('nlp-preview-container');
    if (!container) return;

    if (currentQueryResult) {
      container.innerHTML = `
        <div class="bg-slate-800/95 border border-blue-500/40 rounded-2xl p-5 shadow-xl animate-fadeIn">
          <div class="flex items-center justify-between mb-2">
            <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-blue-500/20 text-blue-300 border border-blue-500/30">
              <span>✨ Respuesta Contable Instantánea</span>
            </span>
            <button onclick="window.AltusApp.clearPreview()" class="text-slate-400 hover:text-white text-xs">Cerrar ✕</button>
          </div>
          <h3 class="text-base font-bold text-white mb-1">${currentQueryResult.title}</h3>
          <p class="text-sm text-slate-300 mb-3">${currentQueryResult.answer}</p>
          <div class="space-y-1.5 bg-slate-900/80 rounded-xl p-3 border border-slate-700/70">
            ${currentQueryResult.highlights.map(h => `<div class="text-xs text-slate-200 font-medium">${h}</div>`).join('')}
          </div>
        </div>
      `;
      return;
    }

    if (!currentDraft) {
      container.innerHTML = '';
      return;
    }

    const accountOptions = state.config
      .filter(c => c.level === 3 || c.code === '1100' || c.code === '4200')
      .map(c => `<option value="${c.fullCode}" ${c.fullCode === currentDraft.cuenta ? 'selected' : ''}>${c.fullCode} (${c.ivaStatus})</option>`)
      .join('');

    const monthOptions = window.AltusNLPEngine.MONTHS
      .map(m => `<option value="${m}" ${m === currentDraft.mes ? 'selected' : ''}>${m.charAt(0).toUpperCase() + m.slice(1)}</option>`)
      .join('');

    const studentBadge = currentDraft.studentUpdate
      ? `<div class="mt-2 p-2.5 rounded-xl bg-emerald-950/60 border border-emerald-500/30 text-xs text-emerald-200 flex items-center justify-between">
           <div>
             <strong>Vínculo Familiar Detectado:</strong> ${currentDraft.studentUpdate.estudiante} (${currentDraft.studentUpdate.idFamilia})
             <span class="block text-emerald-300/80">Contacto: ${currentDraft.studentUpdate.responsable} / ${currentDraft.studentUpdate.contacto}</span>
           </div>
           <span class="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 font-semibold">Actualiza Matriz ${currentDraft.studentUpdate.mes.toUpperCase()}</span>
         </div>`
      : '';

    container.innerHTML = `
      <div class="bg-slate-800/95 border border-amber-500/50 rounded-2xl p-5 shadow-2xl animate-fadeIn">
        <div class="flex flex-wrap items-center justify-between gap-2 mb-3">
          <div class="flex items-center gap-2">
            <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-amber-500/20 text-amber-300 border border-amber-500/30">
              🛡️ Previsualización de Auditoría SAT
            </span>
            <span class="text-xs text-slate-400">${currentDraft.confidence}</span>
          </div>
          <button onclick="window.AltusApp.clearPreview()" class="text-xs text-slate-400 hover:text-white">Cancelar ✕</button>
        </div>

        ${studentBadge}

        <div class="grid grid-cols-1 md:grid-cols-3 gap-3 mt-3">
          <div class="md:col-span-2">
            <label class="block text-xs text-slate-400 mb-1">Descripción oficial en Libro Diario</label>
            <input id="draft-desc" type="text" value="${currentDraft.descripcion}"
              class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white focus:border-amber-400 focus:outline-none" />
          </div>
          <div>
            <label class="block text-xs text-slate-400 mb-1">Monto Bruto (Q)</label>
            <input id="draft-monto" type="number" step="0.01" value="${currentDraft.montoBruto}"
              onchange="window.AltusApp.recalcDraftFromInputs()"
              class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm font-bold text-amber-300 focus:border-amber-400 focus:outline-none" />
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-4 gap-3 mt-3">
          <div class="md:col-span-2">
            <label class="block text-xs text-slate-400 mb-1">Código y Partida Contable SAT</label>
            <select id="draft-cuenta" onchange="window.AltusApp.recalcDraftFromInputs()"
              class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white focus:border-amber-400 focus:outline-none">
              ${accountOptions}
            </select>
          </div>
          <div>
            <label class="block text-xs text-slate-400 mb-1">Cuenta Bancaria</label>
            <select id="draft-banco" onchange="window.AltusApp.recalcDraftFromInputs()"
              class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white focus:border-amber-400 focus:outline-none">
              <option value="BI- Monetaria" ${currentDraft.banco === 'BI- Monetaria' ? 'selected' : ''}>1101 - BI-Monetaria (Operativa)</option>
              <option value="BI-Ahorro" ${currentDraft.banco === 'BI-Ahorro' ? 'selected' : ''}>1102 - BI-Ahorro (Reservas/IGSS)</option>
            </select>
          </div>
          <div>
            <label class="block text-xs text-slate-400 mb-1">Mes Contable</label>
            <select id="draft-mes"
              class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white focus:border-amber-400 focus:outline-none">
              ${monthOptions}
            </select>
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-3 mt-3">
          <div>
            <label class="block text-xs text-slate-400 mb-1">Fecha Operación</label>
            <input id="draft-fecha" type="date" value="${currentDraft.fecha}"
              class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white" />
          </div>
          <div>
            <label class="block text-xs text-slate-400 mb-1">No. Documento / FEL / Boleta</label>
            <input id="draft-doc" type="text" value="${currentDraft.documento}"
              class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white" />
          </div>
          <div>
            <label class="block text-xs text-slate-400 mb-1">Desglose IVA SAT</label>
            <div class="bg-slate-900/90 border border-slate-700/80 rounded-xl px-3 py-2 text-xs text-slate-300">
              Neto: <strong>${formatQ(currentDraft.montoNeto)}</strong> | IVA: <strong>${formatQ(currentDraft.iva)}</strong>
            </div>
          </div>
        </div>

        <div class="mt-3 p-3 rounded-xl bg-slate-900/90 border border-slate-700/80 text-xs space-y-1">
          <div class="text-slate-400 font-semibold uppercase tracking-wider">Asiento de Partida Doble (NIIF / SAT):</div>
          <div class="text-emerald-300 font-mono">DEBE:  ${currentDraft.partidaDoble.debe}</div>
          <div class="text-sky-300 font-mono">HABER: ${currentDraft.partidaDoble.haber}</div>
          <div class="text-slate-400 pt-1 italic">${currentDraft.notaAuditoria}</div>
        </div>

        <div class="mt-4 flex flex-wrap items-center justify-end gap-3">
          <button onclick="window.AltusApp.clearPreview()"
            class="px-4 py-2.5 rounded-xl border border-slate-600 text-slate-300 hover:bg-slate-700 text-sm font-medium">
            Descartar
          </button>
          <button onclick="window.AltusApp.confirmDraftEntry()"
            class="px-5 py-2.5 rounded-xl bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-400 hover:to-amber-500 text-slate-950 font-bold text-sm shadow-lg flex items-center gap-2">
            <span>✅ Confirmar y Guardar en Libro & Sheets</span>
          </button>
        </div>
      </div>
    `;
  }

  function recalcDraftFromInputs() {
    if (!currentDraft) return;
    const monto = parseFloat(document.getElementById('draft-monto').value) || 0;
    const cuenta = document.getElementById('draft-cuenta').value;
    const banco = document.getElementById('draft-banco').value;
    const desc = document.getElementById('draft-desc').value;

    const accObj = state.config.find(c => c.fullCode === cuenta);
    const isAfecto = accObj && accObj.ivaStatus === 'Afecto';
    const iva = isAfecto ? +(monto - monto / 1.12).toFixed(2) : 0;
    const neto = +(monto - iva).toFixed(2);

    currentDraft.descripcion = desc;
    currentDraft.montoBruto = monto;
    currentDraft.cuenta = cuenta;
    currentDraft.banco = banco;
    currentDraft.iva = iva;
    currentDraft.montoNeto = neto;

    if (cuenta === '2102 - Provisiones Laborales') {
      currentDraft.tipo = 'Traspaso Ahorro';
      currentDraft.tipoOriginal = 'Gasto';
    } else if (banco === 'BI-Ahorro') {
      currentDraft.tipo = 'Pago Reserva Ahorro';
      currentDraft.tipoOriginal = 'Gasto';
    } else if (accObj) {
      currentDraft.tipo = accObj.type === 'Ingreso' ? 'Ingreso' : 'Gasto';
      currentDraft.tipoOriginal = currentDraft.tipo;
    }

    renderNLPPreview();
  }

  async function confirmDraftEntry() {
    if (!currentDraft) return;
    const desc = document.getElementById('draft-desc').value.trim();
    const monto = parseFloat(document.getElementById('draft-monto').value) || 0;
    const cuenta = document.getElementById('draft-cuenta').value;
    const banco = document.getElementById('draft-banco').value;
    const mes = document.getElementById('draft-mes').value;
    const fecha = document.getElementById('draft-fecha').value;
    const doc = document.getElementById('draft-doc').value.trim();

    const nextNum = state.journal.length + 1;
    const newEntry = {
      id: `ALT-26-${String(nextNum).padStart(3, '0')}`,
      row: nextNum + 1,
      fecha: fecha,
      mes: mes,
      descripcion: desc,
      cuentaOriginal: cuenta,
      cuenta: cuenta,
      tipo: currentDraft.tipo,
      tipoOriginal: currentDraft.tipoOriginal || (currentDraft.tipo === 'Ingreso' ? 'Ingreso' : 'Gasto'),
      montoBruto: monto,
      iva: currentDraft.iva,
      montoNeto: currentDraft.montoNeto,
      banco: banco,
      saldoMonetaria: 0,
      saldoAhorro: 0,
      documento: doc,
      centroCosto: currentDraft.centroCosto || '00-Administración',
      notaAuditoria: currentDraft.notaAuditoria,
      studentUpdate: currentDraft.studentUpdate || null
    };

    state.journal.push(newEntry);

    // If student tuition payment, update Colegiaturas matrix
    if (currentDraft.studentUpdate) {
      const stu = state.students.find(s => s.idFamilia === currentDraft.studentUpdate.idFamilia);
      if (stu) {
        stu.pagos[mes] = +(Number(stu.pagos[mes] || 0) + monto).toFixed(2);
        stu.totalPagado = Object.values(stu.pagos).reduce((a, b) => a + Number(b || 0), 0);
        stu.pendienteOctubre = Math.max(0, stu.pendienteOctubre - monto);
        stu.estado = stu.pendienteOctubre <= 1 ? 'Al Día' : 'Con Saldo Pendiente';
      }
    }

    saveState();
    const syncRes = await syncEntryToGoogleSheets(newEntry);

    currentDraft = null;
    const inputEl = document.getElementById('nlp-input');
    if (inputEl) inputEl.value = '';

    renderAll();
    showToast(`✅ Asiento ${newEntry.id} (${formatQ(monto)}) registrado en Libro Diario${syncRes.synced ? ' y sincronizado con Google Sheets' : ''}.`);
  }

  function clearPreview() {
    currentDraft = null;
    currentQueryResult = null;
    renderNLPPreview();
  }

  // Render Journal Table in Asistente Tab
  function renderJournalTable() {
    const tbody = document.getElementById('journal-table-body');
    if (!tbody) return;

    const search = journalFilter.search.toLowerCase().trim();
    const filtered = state.journal.slice().reverse().filter(j => {
      if (journalFilter.month !== 'todos' && j.mes !== journalFilter.month) return false;
      if (journalFilter.bank !== 'todos' && j.banco !== journalFilter.bank) return false;
      if (journalFilter.type !== 'todos' && j.tipo !== journalFilter.type) return false;
      if (search) {
        const hay = `${j.id} ${j.fecha} ${j.descripcion} ${j.cuenta} ${j.documento} ${j.notaAuditoria}`.toLowerCase();
        if (!hay.includes(search)) return false;
      }
      return true;
    });

    const countEl = document.getElementById('journal-filtered-count');
    if (countEl) countEl.textContent = `Mostrando ${Math.min(filtered.length, 60)} de ${filtered.length} registros`;

    tbody.innerHTML = filtered.slice(0, 60).map(j => {
      const isIncome = j.tipo === 'Ingreso' || j.tipo === 'Activo' || j.tipo === 'Pasivo';
      const isTransfer = j.tipo === 'Traspaso Ahorro';
      const isFromSavings = j.banco === 'BI-Ahorro';

      let badgeClass = 'bg-rose-500/15 text-rose-300 border-rose-500/30';
      let sign = '-';
      if (isIncome) {
        badgeClass = 'bg-emerald-500/15 text-emerald-300 border-emerald-500/30';
        sign = '+';
      } else if (isTransfer) {
        badgeClass = 'bg-amber-500/15 text-amber-300 border-amber-500/30';
        sign = '⇄';
      } else if (isFromSavings) {
        badgeClass = 'bg-purple-500/15 text-purple-300 border-purple-500/30';
        sign = '🛡️ -';
      }

      return `
        <tr class="border-b border-slate-800/80 hover:bg-slate-800/40 transition-colors">
          <td class="py-2.5 px-3 text-xs font-mono text-slate-400 whitespace-nowrap">${j.fecha}<span class="block text-[10px] text-slate-500">${j.id}</span></td>
          <td class="py-2.5 px-3 text-xs text-white font-medium">
            ${j.descripcion}
            ${j.notaAuditoria ? `<span class="block text-[11px] text-amber-300/80 mt-0.5">🔎 ${j.notaAuditoria}</span>` : ''}
          </td>
          <td class="py-2.5 px-3 text-xs text-slate-300 whitespace-nowrap">${j.cuenta}</td>
          <td class="py-2.5 px-3 text-xs whitespace-nowrap">
            <span class="px-2 py-0.5 rounded-full text-[11px] border ${j.banco === 'BI-Ahorro' ? 'bg-indigo-500/20 text-indigo-300 border-indigo-500/30' : 'bg-slate-700/60 text-slate-200 border-slate-600'}">
              ${j.banco}
            </span>
          </td>
          <td class="py-2.5 px-3 text-xs font-bold text-right whitespace-nowrap">
            <span class="px-2 py-0.5 rounded-lg border ${badgeClass}">${sign} ${formatQ(j.montoBruto)}</span>
          </td>
          <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-300 whitespace-nowrap">${formatQ(j.saldoMonetaria)}</td>
          <td class="py-2.5 px-3 text-xs font-mono text-right text-indigo-300 whitespace-nowrap">${formatQ(j.saldoAhorro)}</td>
        </tr>
      `;
    }).join('');
  }

  // Render Dashboard Charts & Top Expenses
  function renderDashboard() {
    const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre'];
    const monthlyStats = months.map(m => {
      const entries = state.journal.filter(j => j.mes === m);
      let ingresos = 0;
      let gastosOperativos = 0;
      let trasladosAhorro = 0;
      let pagosDesdeAhorro = 0;

      for (const e of entries) {
        const amt = Number(e.montoBruto || 0);
        if (e.tipo === 'Ingreso' || e.tipo === 'Pasivo') {
          ingresos += amt;
        } else if (e.tipo === 'Traspaso Ahorro' || e.cuenta === '2102 - Provisiones Laborales') {
          trasladosAhorro += amt;
        } else if (e.banco === 'BI-Ahorro' || e.tipo === 'Pago Reserva Ahorro') {
          pagosDesdeAhorro += amt;
        } else if (e.tipo === 'Gasto') {
          gastosOperativos += amt;
        }
      }
      return { mes: m, ingresos, gastosOperativos, trasladosAhorro, pagosDesdeAhorro };
    });

    const totalIngresos = monthlyStats.reduce((a, b) => a + b.ingresos, 0);
    const totalGastosOp = monthlyStats.reduce((a, b) => a + b.gastosOperativos, 0);
    const totalTraslados = monthlyStats.reduce((a, b) => a + b.trasladosAhorro, 0);
    const totalIvaCredito = state.journal
      .filter(j => j.tipo === 'Gasto')
      .reduce((a, b) => a + Number(b.iva || 0), 0);

    const kpiContainer = document.getElementById('dashboard-kpis');
    if (kpiContainer) {
      kpiContainer.innerHTML = `
        <div class="bg-slate-800/80 border border-slate-700/70 rounded-2xl p-4">
          <div class="text-xs text-slate-400">Ingresos Totales (Ene-Oct)</div>
          <div class="text-xl font-bold text-emerald-400 mt-1">${formatQ(totalIngresos)}</div>
          <div class="text-[11px] text-slate-400 mt-1">Colegiaturas, Inscripciones y Donaciones</div>
        </div>
        <div class="bg-slate-800/80 border border-slate-700/70 rounded-2xl p-4">
          <div class="text-xs text-slate-400">Gastos Operativos Directos</div>
          <div class="text-xl font-bold text-rose-400 mt-1">${formatQ(totalGastosOp)}</div>
          <div class="text-[11px] text-slate-400 mt-1">Sin duplicar pagos de BI-Ahorro</div>
        </div>
        <div class="bg-slate-800/80 border border-slate-700/70 rounded-2xl p-4">
          <div class="text-xs text-slate-400">Separado a Reserva Ahorro</div>
          <div class="text-xl font-bold text-indigo-400 mt-1">${formatQ(totalTraslados)}</div>
          <div class="text-[11px] text-slate-400 mt-1">Pagado desde Ahorro: ${formatQ(monthlyStats.reduce((a,b)=>a+b.pagosDesdeAhorro,0))}</div>
        </div>
        <div class="bg-slate-800/80 border border-slate-700/70 rounded-2xl p-4">
          <div class="text-xs text-slate-400">IVA Crédito Fiscal SAT (1104)</div>
          <div class="text-xl font-bold text-amber-400 mt-1">${formatQ(totalIvaCredito)}</div>
          <div class="text-[11px] text-slate-400 mt-1">Respaldado en compras afectas FEL</div>
        </div>
      `;
    }

    // Render Bar Chart Ingresos vs Egresos
    const maxVal = Math.max(...monthlyStats.map(s => Math.max(s.ingresos, s.gastosOperativos + s.trasladosAhorro)), 1);
    const chartEl = document.getElementById('income-vs-expense-chart');
    if (chartEl) {
      chartEl.innerHTML = monthlyStats.map(s => {
        const hIng = Math.max(6, Math.round((s.ingresos / maxVal) * 165));
        const hGas = Math.max(6, Math.round((s.gastosOperativos / maxVal) * 165));
        const hAho = Math.max(0, Math.round((s.trasladosAhorro / maxVal) * 165));
        const net = s.ingresos - s.gastosOperativos - s.trasladosAhorro;
        return `
          <div class="flex flex-col items-center flex-1 min-w-[54px] group relative">
            <div class="opacity-0 group-hover:opacity-100 transition-opacity absolute -top-16 z-20 bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-[10px] text-slate-200 whitespace-nowrap shadow-xl pointer-events-none">
              <strong class="uppercase text-amber-300">${s.mes}</strong><br/>
              Ingresos: ${formatQ(s.ingresos)}<br/>
              Gastos Op.: ${formatQ(s.gastosOperativos)}<br/>
              A Reserva Ahorro: ${formatQ(s.trasladosAhorro)}
            </div>
            <div class="h-44 flex items-end justify-center gap-1.5 w-full px-1 border-b border-slate-700/80">
              <div style="height:${hIng}px" class="w-3.5 md:w-4 rounded-t-md bg-gradient-to-t from-emerald-600 to-emerald-400 transition-all" title="Ingresos: ${formatQ(s.ingresos)}"></div>
              <div class="flex flex-col justify-end w-3.5 md:w-4">
                ${hAho > 0 ? `<div style="height:${hAho}px" class="w-full rounded-t-md bg-indigo-400" title="Traslado a Ahorro: ${formatQ(s.trasladosAhorro)}"></div>` : ''}
                <div style="height:${hGas}px" class="w-full ${hAho > 0 ? '' : 'rounded-t-md'} bg-gradient-to-t from-rose-600 to-rose-400" title="Gastos Operativos: ${formatQ(s.gastosOperativos)}"></div>
              </div>
            </div>
            <div class="text-[11px] font-medium text-slate-300 mt-2 capitalize">${s.mes.slice(0, 3)}</div>
            <div class="text-[10px] font-mono ${net >= 0 ? 'text-emerald-400' : 'text-rose-400'}">${net >= 0 ? '+' : ''}${Math.round(net / 1000)}k</div>
          </div>
        `;
      }).join('');
    }

    // Top Expense Categories
    const byExpenseAccount = {};
    for (const j of state.journal) {
      if (j.tipo === 'Ingreso' || j.tipo === 'Activo' || j.tipo === 'Pasivo') continue;
      // Avoid double counting: show real expenses + reserve transfers clearly labeled
      const acc = j.cuenta;
      byExpenseAccount[acc] = (byExpenseAccount[acc] || 0) + Number(j.montoBruto || 0);
    }

    const sortedExpenses = Object.entries(byExpenseAccount)
      .sort((a, b) => b[1] - a[1]);
    const grandTotalOut = sortedExpenses.reduce((a, b) => a + b[1], 0) || 1;

    const topExpensesEl = document.getElementById('top-expenses-list');
    if (topExpensesEl) {
      topExpensesEl.innerHTML = sortedExpenses.map(([acc, total], idx) => {
        const pct = ((total / grandTotalOut) * 100).toFixed(1);
        const isReserve = acc.includes('2102') || acc.includes('2103') || acc.includes('5104');
        return `
          <div class="p-3 rounded-xl bg-slate-900/70 border border-slate-800">
            <div class="flex items-center justify-between text-xs mb-1.5">
              <span class="font-semibold text-white">${idx + 1}. ${acc} ${isReserve ? '<span class="ml-1 text-[10px] px-1.5 py-0.5 rounded bg-indigo-500/20 text-indigo-300">Reserva / Ahorro</span>' : ''}</span>
              <span class="font-mono font-bold text-slate-200">${formatQ(total)} <span class="text-slate-400 font-normal">(${pct}%)</span></span>
            </div>
            <div class="w-full h-2 bg-slate-800 rounded-full overflow-hidden">
              <div style="width:${Math.max(3, pct)}%" class="h-full rounded-full ${isReserve ? 'bg-indigo-500' : 'bg-gradient-to-r from-amber-500 to-rose-500'}"></div>
            </div>
          </div>
        `;
      }).join('');
    }
  }

  // Render Colegiaturas Matrix
  function renderColegiaturas() {
    const container = document.getElementById('colegiaturas-table-body');
    if (!container) return;

    const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre'];

    container.innerHTML = state.students.map(s => {
      const badgeColor = s.estado === 'Al Día'
        ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30'
        : (s.estado === 'Retirado' ? 'bg-slate-700 text-slate-400 border-slate-600' : 'bg-amber-500/20 text-amber-300 border-amber-500/30');

      const monthCells = months.map(m => {
        const paid = Number(s.pagos[m] || 0);
        const startIdx = months.indexOf((s.mesInicio || 'Enero').toLowerCase());
        const mIdx = months.indexOf(m);
        if (s.estado === 'Retirado' || mIdx < startIdx) {
          return `<td class="py-2 px-2 text-center text-[11px] text-slate-600">—</td>`;
        }
        if (paid >= s.cuotaMensual * 0.95) {
          return `<td class="py-2 px-2 text-center"><span class="inline-block px-1.5 py-0.5 rounded bg-emerald-500/20 text-emerald-300 text-[11px] font-mono">Q${paid.toLocaleString()}</span></td>`;
        } else if (paid > 0) {
          return `<td class="py-2 px-2 text-center"><span class="inline-block px-1.5 py-0.5 rounded bg-amber-500/20 text-amber-300 text-[11px] font-mono">Q${paid.toLocaleString()}</span></td>`;
        } else {
          return `<td class="py-2 px-2 text-center">
            <button onclick="window.AltusApp.quickTuitionPay('${s.estudiante}', '${m}')"
              class="px-2 py-0.5 rounded bg-rose-500/20 hover:bg-rose-500/40 text-rose-300 border border-rose-500/30 text-[10px] font-semibold transition-colors"
              title="Registrar pago de ${m} para ${s.estudiante}">
              + Cobrar
            </button>
          </td>`;
        }
      }).join('');

      return `
        <tr class="border-b border-slate-800 hover:bg-slate-800/40">
          <td class="py-3 px-3 text-xs">
            <div class="font-bold text-white">${s.estudiante}</div>
            <div class="text-[11px] text-slate-400">${s.idFamilia} • ${s.grado}</div>
          </td>
          <td class="py-3 px-3 text-xs">
            <div class="text-slate-200 font-medium">${s.responsable}</div>
            ${s.contactoSecundario ? `<div class="text-[11px] text-amber-300 font-medium">👩‍👦 ${s.contactoSecundario}</div>` : ''}
            <div class="text-[10px] text-slate-400">${s.email} • NIT: ${s.nit}</div>
          </td>
          <td class="py-3 px-3 text-xs font-mono font-bold text-slate-200 text-right">${formatQ(s.cuotaMensual)}</td>
          ${monthCells}
          <td class="py-3 px-3 text-xs font-mono font-bold text-emerald-400 text-right">${formatQ(s.totalPagado)}</td>
          <td class="py-3 px-3 text-xs text-center">
            <span class="px-2.5 py-1 rounded-full text-[11px] font-semibold border ${badgeColor}">${s.estado}</span>
          </td>
        </tr>
      `;
    }).join('');
  }

  function quickTuitionPay(studentName, month) {
    switchTab('asistente');
    handleProcessNLP(`Ya pagó ${studentName} colegiatura de ${month}`);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // Render Nóminas & Savings Reserve Tab
  function renderNominas() {
    const staffBody = document.getElementById('nominas-staff-body');
    if (staffBody) {
      staffBody.innerHTML = state.staff.map(emp => `
        <tr class="border-b border-slate-800 hover:bg-slate-800/40">
          <td class="py-3 px-3 text-xs">
            <div class="font-bold text-white">${emp.nombre}</div>
            <div class="text-[11px] text-slate-400">${emp.puesto} • ${emp.cuentaBanco}</div>
          </td>
          <td class="py-3 px-3 text-xs font-mono text-right text-slate-200">${formatQ(emp.sueldoBase)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-slate-300">${formatQ(emp.bonificacion)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-rose-300">-${formatQ(emp.igssLaboral)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-rose-300">-${formatQ(emp.isrRetenido)}</td>
          <td class="py-3 px-3 text-xs font-mono font-bold text-right text-emerald-400">${formatQ(emp.liquidoRecibir)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-indigo-300">${formatQ(emp.igssPatronal)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-amber-300">${formatQ(emp.reservaBono14)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-amber-300">${formatQ(emp.reservaAguinaldo)}</td>
          <td class="py-3 px-3 text-xs text-center">
            <button onclick="window.AltusApp.triggerNLPFromButton('Pago sueldo ${emp.nombre}')"
              class="px-2.5 py-1 rounded-lg bg-emerald-500/20 hover:bg-emerald-500/30 text-emerald-300 border border-emerald-500/30 text-xs font-semibold">
              Pagar Día 27
            </button>
          </td>
        </tr>
      `).join('');
    }

    const obligationsEl = document.getElementById('day27-obligations-list');
    if (obligationsEl) {
      obligationsEl.innerHTML = state.day27Obligations.map(ob => `
        <div class="flex items-center justify-between p-3 rounded-xl bg-slate-900/80 border border-slate-800 hover:border-slate-700">
          <div>
            <div class="text-xs font-semibold text-white">${ob.concepto}</div>
            <div class="text-[11px] text-slate-400">${ob.cuenta} • <span class="text-amber-300">${ob.banco}</span></div>
          </div>
          <div class="flex items-center gap-3">
            <span class="text-xs font-mono font-bold text-slate-200">${formatQ(ob.monto)}</span>
            <button onclick="window.AltusApp.triggerNLPFromButton('${ob.concepto} Q${ob.monto}')"
              class="px-2.5 py-1 rounded-lg bg-slate-800 hover:bg-amber-500 hover:text-slate-950 text-amber-300 border border-amber-500/30 text-xs font-semibold transition-colors">
              Registrar
            </button>
          </div>
        </div>
      `).join('');
    }

    // Savings Account Movements Ledger
    const savingsMovements = state.journal.filter(j =>
      j.banco === 'BI-Ahorro' || j.cuenta === '2102 - Provisiones Laborales' || j.tipo === 'Traspaso Ahorro'
    );
    const savingsBody = document.getElementById('savings-ledger-body');
    if (savingsBody) {
      savingsBody.innerHTML = savingsMovements.slice().reverse().map(j => {
        const isDeposit = j.banco !== 'BI-Ahorro';
        return `
          <tr class="border-b border-slate-800/80">
            <td class="py-2.5 px-3 text-xs font-mono text-slate-400">${j.fecha}</td>
            <td class="py-2.5 px-3 text-xs text-white font-medium">${j.descripcion}</td>
            <td class="py-2.5 px-3 text-xs text-slate-300">${j.cuenta}</td>
            <td class="py-2.5 px-3 text-xs font-mono font-bold text-right ${isDeposit ? 'text-emerald-400' : 'text-rose-400'}">
              ${isDeposit ? '+ ' : '- '}${formatQ(j.montoBruto)}
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-indigo-300">${formatQ(j.saldoAhorro)}</td>
          </tr>
        `;
      }).join('');
    }
  }

  // Render Presupuesto 2027 & Conciliación
  function renderBudgetAndReconciliation() {
    const budgetBody = document.getElementById('budget-2027-body');
    if (budgetBody) {
      budgetBody.innerHTML = state.budget2027.partidas.map((p, idx) => {
        const ratio = p.anual2027 > 0 ? Math.min(100, ((p.real2026Oct / p.anual2027) * 100)).toFixed(1) : 0;
        const isIng = p.tipo === 'Ingreso';
        return `
          <tr class="border-b border-slate-800 hover:bg-slate-800/40">
            <td class="py-2.5 px-3 text-xs font-medium text-white">${p.codigo}</td>
            <td class="py-2.5 px-3 text-xs">
              <span class="px-2 py-0.5 rounded-full text-[11px] ${isIng ? 'bg-emerald-500/20 text-emerald-300' : 'bg-rose-500/20 text-rose-300'}">${p.tipo}</span>
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-300">${formatQ(p.mensualEne)}</td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-300">${formatQ(p.mensualFebDic)}</td>
            <td class="py-2.5 px-3 text-xs font-mono font-bold text-right text-amber-300">
              <input type="number" value="${p.anual2027}"
                onchange="window.AltusApp.updateBudget2027(${idx}, this.value)"
                class="w-28 bg-slate-900 border border-slate-700 rounded-lg px-2 py-1 text-right text-amber-300 font-mono text-xs" />
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-300">${formatQ(p.real2026Oct)}</td>
            <td class="py-2.5 px-3 text-xs text-right">
              <div class="flex items-center justify-end gap-2">
                <div class="w-16 h-1.5 bg-slate-800 rounded-full overflow-hidden">
                  <div style="width:${ratio}%" class="h-full ${isIng ? 'bg-emerald-400' : 'bg-amber-400'}"></div>
                </div>
                <span class="font-mono text-[11px] text-slate-400">${ratio}%</span>
              </div>
            </td>
          </tr>
        `;
      }).join('');
    }

    const recBody = document.getElementById('reconciliation-body');
    if (recBody) {
      recBody.innerHTML = state.reconciliations.map((r, idx) => {
        const hasBank = r.saldoEstadoCuenta !== null;
        return `
          <tr class="border-b border-slate-800 hover:bg-slate-800/40">
            <td class="py-2.5 px-3 text-xs font-bold capitalize text-white">${r.mes}</td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-200">${formatQ(r.saldoLibroMonetaria)}</td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-300">
              ${hasBank ? formatQ(r.saldoEstadoCuenta) : '<span class="text-slate-500">Pendiente</span>'}
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right ${r.diferencia === 0 ? 'text-emerald-400' : 'text-amber-400'}">
              ${hasBank ? formatQ(r.diferencia) : '—'}
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-indigo-300">${formatQ(r.saldoAhorro)}</td>
            <td class="py-2.5 px-3 text-xs text-center">
              <span class="px-2 py-0.5 rounded-full text-[11px] font-semibold ${hasBank ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30' : 'bg-slate-700 text-slate-300'}">
                ${hasBank ? '✅ Conciliado' : r.estadoAuditado}
              </span>
            </td>
            <td class="py-2.5 px-3 text-[11px] text-slate-400">${r.explicacionAuditoria || 'Sin diferencias'}</td>
          </tr>
        `;
      }).join('');
    }
  }

  function updateBudget2027(idx, newVal) {
    const val = parseFloat(newVal) || 0;
    if (state.budget2027.partidas[idx]) {
      state.budget2027.partidas[idx].anual2027 = val;
      state.budget2027.partidas[idx].mensualFebDic = +(val / 12).toFixed(2);
      saveState();
      renderBudgetAndReconciliation();
      showToast('Presupuesto 2027 actualizado.');
    }
  }

  function triggerNLPFromButton(promptText) {
    switchTab('asistente');
    handleProcessNLP(promptText);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function switchTab(tabId) {
    activeTab = tabId;
    document.querySelectorAll('.tab-section').forEach(el => {
      el.classList.toggle('hidden', el.id !== `tab-${tabId}`);
    });
    document.querySelectorAll('.nav-tab-btn').forEach(btn => {
      const isTarget = btn.getAttribute('data-tab') === tabId;
      if (isTarget) {
        btn.classList.add('bg-amber-500/20', 'text-amber-300', 'border-amber-500/40');
        btn.classList.remove('text-slate-400', 'border-transparent');
      } else {
        btn.classList.remove('bg-amber-500/20', 'text-amber-300', 'border-amber-500/40');
        btn.classList.add('text-slate-400', 'border-transparent');
      }
    });
  }

  // Voice Dictation (iPhone Safari & Mac Web Speech API)
  function startVoiceDictation() {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      showToast('Usa el micrófono del teclado de tu iPhone o escribe en la barra superior.');
      document.getElementById('nlp-input').focus();
      return;
    }
    const recognition = new SpeechRecognition();
    recognition.lang = 'es-GT';
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;

    const micBtn = document.getElementById('mic-btn');
    if (micBtn) micBtn.classList.add('animate-pulse', 'bg-rose-600', 'text-white');
    showToast('🎙️ Escuchando... Di por ejemplo: "Ya pagó Pamela"');

    recognition.onresult = function (event) {
      const transcript = event.results[0][0].transcript;
      document.getElementById('nlp-input').value = transcript;
      handleProcessNLP(transcript);
    };
    recognition.onerror = function () {
      if (micBtn) micBtn.classList.remove('animate-pulse', 'bg-rose-600', 'text-white');
    };
    recognition.onend = function () {
      if (micBtn) micBtn.classList.remove('animate-pulse', 'bg-rose-600', 'text-white');
    };
    recognition.start();
  }

  // Export Live Data to CSV / Excel compatible format
  function exportJournalCSV() {
    const headers = [
      'ID', 'Fecha', 'Mes', 'Descripcion', 'Cuenta_SAT', 'Tipo',
      'Monto_Bruto_Q', 'IVA_Q', 'Monto_Neto_Q', 'Banco',
      'Saldo_BI_Monetaria_Q', 'Saldo_BI_Ahorro_Q', 'Documento_FEL', 'Centro_Costo', 'Nota_Auditoria'
    ];
    const rows = state.journal.map(j => [
      j.id, j.fecha, j.mes, `"${(j.descripcion || '').replace(/"/g, '""')}"`,
      `"${j.cuenta}"`, j.tipo, j.montoBruto, j.iva, j.montoNeto, j.banco,
      j.saldoMonetaria, j.saldoAhorro, `"${j.documento || ''}"`, `"${j.centroCosto || ''}"`,
      `"${(j.notaAuditoria || '').replace(/"/g, '""')}"`
    ].join(','));

    const csvContent = '\uFEFF' + headers.join(',') + '\n' + rows.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `ALTUS_Diario_Auditado_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('📥 Libro Diario exportado correctamente.');
  }

  function openSheetsModal() {
    const modal = document.getElementById('sheets-modal');
    const cfg = getSheetsConfig();
    document.getElementById('webhook-url-input').value = cfg.webhookUrl || '';
    modal.classList.remove('hidden');
  }

  function closeSheetsModal() {
    document.getElementById('sheets-modal').classList.add('hidden');
  }

  function saveWebhookModal() {
    const url = document.getElementById('webhook-url-input').value.trim();
    saveSheetsConfig({ webhookUrl: url, updatedAt: new Date().toISOString() });
    closeSheetsModal();
    showToast(url ? '🔗 Conectado a Google Sheets en tiempo real.' : 'Modo local activado.');
  }

  function showToast(msg) {
    const toast = document.getElementById('app-toast');
    if (!toast) return;
    toast.textContent = msg;
    toast.classList.remove('hidden', 'opacity-0');
    setTimeout(() => {
      toast.classList.add('opacity-0');
      setTimeout(() => toast.classList.add('hidden'), 300);
    }, 3800);
  }

  function renderAll() {
    renderHeader();
    renderNLPPreview();
    renderJournalTable();
    renderDashboard();
    renderColegiaturas();
    renderNominas();
    renderBudgetAndReconciliation();
  }

  // Initialize Listeners
  window.addEventListener('DOMContentLoaded', () => {
    renderAll();

    const nlpForm = document.getElementById('nlp-form');
    if (nlpForm) {
      nlpForm.addEventListener('submit', e => {
        e.preventDefault();
        handleProcessNLP();
      });
    }

    const searchInput = document.getElementById('journal-search');
    if (searchInput) {
      searchInput.addEventListener('input', e => {
        journalFilter.search = e.target.value;
        renderJournalTable();
      });
    }

    const monthSelect = document.getElementById('journal-month-filter');
    if (monthSelect) {
      monthSelect.addEventListener('change', e => {
        journalFilter.month = e.target.value;
        renderJournalTable();
      });
    }

    const bankSelect = document.getElementById('journal-bank-filter');
    if (bankSelect) {
      bankSelect.addEventListener('change', e => {
        journalFilter.bank = e.target.value;
        renderJournalTable();
      });
    }
  });

  window.AltusApp = {
    handleProcessNLP,
    recalcDraftFromInputs,
    confirmDraftEntry,
    clearPreview,
    switchTab,
    quickTuitionPay,
    triggerNLPFromButton,
    updateBudget2027,
    startVoiceDictation,
    exportJournalCSV,
    openSheetsModal,
    closeSheetsModal,
    saveWebhookModal
  };
})();
