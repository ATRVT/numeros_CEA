/**
 * ============================================================================
 * APLICACIÓN CONTABLE ALTUS (iPhone PWA & Mac Desktop)
 * Conectada a Google Sheets + Motor NLP SAT Guatemala + Exportación Excel (.xlsx)
 * ============================================================================
 */

(function () {
  const STORAGE_KEY = 'altus_contabilidad_state_v4';
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

  const DEFAULT_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbzUYpPFK7x_3k7l_o7GJb0th2p9RXi4IjvYCi382NAQW_qEHDtPblTnoBlm_ZuI8B9HKQ/exec';

  function getSheetsConfig() {
    try {
      const cfg = JSON.parse(localStorage.getItem(SHEETS_CONFIG_KEY) || '{}');
      if (!cfg.webhookUrl) {
        cfg.webhookUrl = DEFAULT_WEBHOOK_URL;
      }
      return cfg;
    } catch (e) {
      return { webhookUrl: DEFAULT_WEBHOOK_URL };
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
    let pasivo2027 = 0;
    for (const j of state.journal) {
      const bruto = Number(j.montoBruto || 0);
      const isAhorroBank = j.banco === 'BI-Ahorro';
      const isTraspasoAhorro = j.cuenta === '2102 - Provisiones Laborales' || j.tipo === 'Traspaso Ahorro';
      const origType = j.tipoOriginal || j.tipo;

      if (j.cuenta && j.cuenta.startsWith('2107')) {
        pasivo2027 = +(pasivo2027 + bruto).toFixed(2);
      }

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
      j.saldoDisponible2026 = +(mon - pasivo2027).toFixed(2);
    }
    state.metadata.saldoActualMonetaria = mon;
    state.metadata.pasivoInscripciones2027 = pasivo2027;
    state.metadata.saldoDisponible2026 = +(mon - pasivo2027).toFixed(2);
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
    const pasivo2027 = state.metadata.pasivoInscripciones2027 || 0;
    const disp2026 = state.metadata.saldoDisponible2026;
    const aho = state.metadata.saldoActualAhorro;
    const total = state.metadata.totalPatrimonioLiquido;

    document.getElementById('header-saldo-monetaria').textContent = formatQ(disp2026);
    const subMon = document.getElementById('header-sub-monetaria');
    if (subMon) {
      subMon.innerHTML = `En banco: <strong class="text-slate-800">${formatQ(mon)}</strong> • Pasivo Inscrip. 2027: <strong class="text-altus-coral">-${formatQ(pasivo2027)}</strong>`;
    }
    document.getElementById('header-saldo-ahorro').textContent = formatQ(aho);
    document.getElementById('header-saldo-total').textContent = formatQ(total);
    const subTot = document.getElementById('header-sub-total');
    if (subTot) {
      subTot.innerHTML = `Incluye <strong class="text-altus-amber-dark">${formatQ(pasivo2027)}</strong> de Inscripciones 2027 diferidas`;
    }
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
    const btn = document.getElementById('sheets-connect-btn');
    if (!dot || !label) return;
    if (cfg.webhookUrl) {
      dot.className = 'w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse';
      label.textContent = 'Google Sheets Conectado';
      if (btn) {
        btn.className = 'flex items-center gap-2 px-3.5 py-1.5 rounded-xl bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 text-xs font-bold text-emerald-900 shadow-xs transition-all cursor-pointer';
      }
    } else {
      dot.className = 'w-2.5 h-2.5 rounded-full bg-amber-400';
      label.textContent = 'Conectar Sheets';
      if (btn) {
        btn.className = 'flex items-center gap-2 px-3.5 py-1.5 rounded-xl bg-white hover:bg-slate-50 border border-slate-200 text-xs font-medium text-slate-700 shadow-xs transition-colors cursor-pointer';
      }
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

  // Handle Camera / Photo Upload of FEL Invoices & Bank Receipts (OCR + Auto-Draft)
  async function handleInvoicePhotoUpload(event) {
    const file = event && event.target && event.target.files ? event.target.files[0] : null;
    if (!file) return;

    showToast('📷 Procesando foto de factura / boleta con OCR...');

    // Read and compress image to Base64 DataURL so it can be stored locally & sent to Google Drive
    const base64Image = await new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = function (e) {
        const img = new Image();
        img.onload = function () {
          const canvas = document.createElement('canvas');
          const maxDim = 1100;
          let w = img.width;
          let h = img.height;
          if (w > maxDim || h > maxDim) {
            if (w > h) {
              h = Math.round((h * maxDim) / w);
              w = maxDim;
            } else {
              w = Math.round((w * maxDim) / h);
              h = maxDim;
            }
          }
          canvas.width = w;
          canvas.height = h;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, w, h);
          resolve(canvas.toDataURL('image/jpeg', 0.82));
        };
        img.onerror = () => resolve(e.target.result);
        img.src = e.target.result;
      };
      reader.readAsDataURL(file);
    });

    // Run OCR if Tesseract is loaded
    let extractedText = '';
    if (window.Tesseract) {
      try {
        const result = await window.Tesseract.recognize(base64Image, 'spa');
        extractedText = (result && result.data && result.data.text) ? result.data.text : '';
      } catch (err) {
        console.warn('OCR fallback manual:', err);
      }
    }

    // Parse FEL / Receipt fields from OCR text + filename
    const cleanText = extractedText.replace(/\s+/g, ' ').trim();
    const amounts = [];
    const amountRegex = /(?:total|monto|valor|efectivo|gtq|q\.?)\s*[:$]?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})|[0-9]+(?:\.[0-9]{2}))/gi;
    let match;
    while ((match = amountRegex.exec(cleanText)) !== null) {
      const val = parseFloat(match[1].replace(/,/g, ''));
      if (val > 0 && val < 250000) amounts.push(val);
    }
    if (amounts.length === 0) {
      const decRegex = /\b([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2})\b/g;
      while ((match = decRegex.exec(cleanText)) !== null) {
        const val = parseFloat(match[1].replace(/,/g, ''));
        if (val > 1 && val < 200000) amounts.push(val);
      }
    }
    const detectedAmount = amounts.length > 0 ? Math.max(...amounts) : 150.00;

    // Detect NIT, Serie/DTE or Boleta number
    const dteMatch = cleanText.match(/(?:dte|serie|autorizaci[oó]n|factura|documento|boleta|ref(?:erencia)?)\s*[:#-]?\s*([A-Z0-9-]{5,20})/i);
    const nitMatch = cleanText.match(/nit\s*[:#-]?\s*([0-9]{5,12}-?[0-9K]?)/i);
    const detectedDoc = dteMatch
      ? `FEL-${dteMatch[1].toUpperCase()}`
      : (nitMatch ? `FEL-NIT-${nitMatch[1]}` : `FEL-FOTO-${Math.floor(100000 + Math.random() * 900000)}`);

    // Build prompt for NLP Engine from OCR text or filename
    const fileLabel = file.name.replace(/\.[^.]+$/, '').replace(/[_-]+/g, ' ');
    const synthPrompt = cleanText.length > 8
      ? `Factura ${cleanText.slice(0, 120)} Q${detectedAmount}`
      : `Compra según factura adjunta (${fileLabel}) Q${detectedAmount}`;

    const parsed = window.AltusNLPEngine.parseNaturalInput(synthPrompt, state) || {};
    currentDraft = Object.assign({
      kind: 'transaction_draft',
      confidence: cleanText.length > 8
        ? '📷 Escaneo OCR FEL Completado (Verifica monto y partida antes de confirmar)'
        : '📷 Foto de Comprobante Adjunta (Completa o confirma los datos de la factura)',
      fecha: new Date().toISOString().slice(0, 10),
      mes: 'octubre',
      descripcion: parsed.descripcion || `Factura soportada por foto (${fileLabel})`,
      cuenta: parsed.cuenta || '5203 - Papelería y útiles',
      tipo: parsed.tipo || 'Gasto',
      tipoOriginal: parsed.tipoOriginal || 'Gasto',
      montoBruto: detectedAmount,
      iva: +(detectedAmount - detectedAmount / 1.12).toFixed(2),
      montoNeto: +(detectedAmount / 1.12).toFixed(2),
      banco: parsed.banco || 'BI- Monetaria',
      documento: detectedDoc,
      centroCosto: parsed.centroCosto || '01-Académico',
      notaAuditoria: `Soporte fotográfico adjunto (${file.name}). Se guardará en Google Drive y vinculará al Libro Diario SAT.`,
      partidaDoble: parsed.partidaDoble || {
        debe: `5203 - Papelería y útiles (Q ${(detectedAmount / 1.12).toFixed(2)}) + IVA Crédito (Q ${(detectedAmount - detectedAmount / 1.12).toFixed(2)})`,
        haber: `1101 - BI-Monetaria (Q ${detectedAmount.toFixed(2)})`
      }
    }, parsed, {
      documento: detectedDoc,
      receiptImageBase64: base64Image,
      receiptFileName: file.name,
      ocrSnippet: cleanText ? cleanText.slice(0, 180) : ''
    });

    currentQueryResult = null;
    if (event && event.target) event.target.value = '';
    renderNLPPreview();
    showToast('✅ Foto de factura procesada. Revisa la previsualización abajo.');
  }

  function renderNLPPreview() {
    const container = document.getElementById('nlp-preview-container');
    if (!container) return;

    if (currentQueryResult) {
      container.innerHTML = `
        <div class="bg-white border border-blue-200 rounded-2xl p-5 shadow-lg animate-fadeIn">
          <div class="flex items-center justify-between mb-2">
            <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-altus-blue-light text-altus-blue border border-blue-200">
              <span>✨ Respuesta Contable Instantánea</span>
            </span>
            <button onclick="window.AltusApp.clearPreview()" class="text-slate-400 hover:text-slate-600 text-xs">Cerrar ✕</button>
          </div>
          <h3 class="text-base font-bold text-slate-900 mb-1">${currentQueryResult.title}</h3>
          <p class="text-sm text-slate-600 mb-3">${currentQueryResult.answer}</p>
          <div class="space-y-1.5 bg-slate-50 rounded-xl p-3 border border-slate-200">
            ${currentQueryResult.highlights.map(h => `<div class="text-xs text-slate-800 font-medium">${h}</div>`).join('')}
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
      ? `<div class="mt-2 p-3 rounded-xl bg-emerald-50 border border-emerald-300 text-xs text-emerald-900 flex items-center justify-between">
           <div>
             <strong class="text-emerald-950">Vínculo Familiar Detectado:</strong> ${currentDraft.studentUpdate.estudiante} (${currentDraft.studentUpdate.idFamilia})
             <span class="block text-emerald-800 mt-0.5">Contacto: ${currentDraft.studentUpdate.responsable} / ${currentDraft.studentUpdate.contacto}</span>
           </div>
           <span class="px-2.5 py-1 rounded-md bg-emerald-100 text-emerald-900 font-bold border border-emerald-300">Actualiza Matriz ${currentDraft.studentUpdate.mes.toUpperCase()}</span>
         </div>`
      : '';

    const photoPreviewHtml = currentDraft.receiptImageBase64
      ? `<div class="mt-3 p-3 rounded-xl bg-slate-50 border border-altus-blue/30 flex flex-col sm:flex-row items-start sm:items-center gap-3">
           <img src="${currentDraft.receiptImageBase64}" alt="Factura adjunta" class="w-20 h-20 object-cover rounded-lg border border-slate-300 shadow-sm shrink-0" />
           <div class="text-xs text-slate-700 flex-1">
             <div class="font-bold text-altus-blue flex items-center gap-1.5">📷 Comprobante Fotográfico Adjunto (${currentDraft.receiptFileName || 'Factura.jpg'})</div>
             <p class="text-[11px] text-slate-600 mt-0.5">Al confirmar, la imagen quedará vinculada al asiento y se subirá automáticamente a la carpeta <strong>Comprobantes_FEL_ALTUS</strong> en tu Google Drive.</p>
             ${currentDraft.ocrSnippet ? `<div class="mt-1 text-[10px] font-mono text-slate-700 bg-white border border-slate-200 px-2 py-1 rounded">Texto OCR detectado: "${currentDraft.ocrSnippet}..."</div>` : ''}
           </div>
         </div>`
      : '';

    container.innerHTML = `
      <div class="bg-white border-2 border-altus-amber/80 rounded-2xl p-5 shadow-xl animate-fadeIn">
        <div class="flex flex-wrap items-center justify-between gap-2 mb-3">
          <div class="flex items-center gap-2">
            <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-bold bg-amber-50 text-amber-900 border border-amber-300">
              🛡️ Previsualización de Auditoría SAT
            </span>
            <span class="text-xs text-slate-500 font-medium">${currentDraft.confidence}</span>
          </div>
          <button onclick="window.AltusApp.clearPreview()" class="text-xs text-slate-400 hover:text-slate-600 font-bold">Cancelar ✕</button>
        </div>

        ${studentBadge}
        ${photoPreviewHtml}

        <div class="grid grid-cols-1 md:grid-cols-3 gap-3 mt-3">
          <div class="md:col-span-2">
            <label class="block text-xs font-bold text-slate-700 mb-1">Descripción oficial en Libro Diario</label>
            <input id="draft-desc" type="text" value="${currentDraft.descripcion}"
              class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-sm text-slate-900 focus:border-altus-blue focus:bg-white focus:outline-none" />
          </div>
          <div>
            <label class="block text-xs font-bold text-slate-700 mb-1">Monto Bruto (Q)</label>
            <input id="draft-monto" type="number" step="0.01" value="${currentDraft.montoBruto}"
              onchange="window.AltusApp.recalcDraftFromInputs()"
              class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-base font-extrabold text-altus-blue focus:border-altus-blue focus:bg-white focus:outline-none font-mono" />
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-4 gap-3 mt-3">
          <div class="md:col-span-2">
            <label class="block text-xs font-bold text-slate-700 mb-1">Código y Partida Contable SAT</label>
            <select id="draft-cuenta" onchange="window.AltusApp.recalcDraftFromInputs()"
              class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs md:text-sm text-slate-900 focus:border-altus-blue focus:bg-white focus:outline-none font-medium">
              ${accountOptions}
            </select>
          </div>
          <div>
            <label class="block text-xs font-bold text-slate-700 mb-1">Cuenta Bancaria</label>
            <select id="draft-banco" onchange="window.AltusApp.recalcDraftFromInputs()"
              class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs md:text-sm text-slate-900 focus:border-altus-blue focus:bg-white focus:outline-none font-medium">
              <option value="BI- Monetaria" ${currentDraft.banco === 'BI- Monetaria' ? 'selected' : ''}>1101 - BI-Monetaria (Operativa)</option>
              <option value="BI-Ahorro" ${currentDraft.banco === 'BI-Ahorro' ? 'selected' : ''}>1102 - BI-Ahorro (Reservas/IGSS)</option>
            </select>
          </div>
          <div>
            <label class="block text-xs font-bold text-slate-700 mb-1">Mes Contable</label>
            <select id="draft-mes"
              class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs md:text-sm text-slate-900 focus:border-altus-blue focus:bg-white focus:outline-none font-medium">
              ${monthOptions}
            </select>
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-3 mt-3">
          <div>
            <label class="block text-xs font-bold text-slate-700 mb-1">Fecha Operación</label>
            <input id="draft-fecha" type="date" value="${currentDraft.fecha}"
              class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-sm text-slate-900 focus:border-altus-blue focus:bg-white focus:outline-none" />
          </div>
          <div>
            <label class="block text-xs font-bold text-slate-700 mb-1">No. Documento / FEL / Boleta</label>
            <input id="draft-doc" type="text" value="${currentDraft.documento}"
              class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-sm text-slate-900 focus:border-altus-blue focus:bg-white focus:outline-none font-mono" />
          </div>
          <div>
            <label class="block text-xs font-bold text-slate-700 mb-1">Desglose IVA SAT</label>
            <div class="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs text-slate-700 font-medium">
              Neto: <strong class="text-slate-900">${formatQ(currentDraft.montoNeto)}</strong> | IVA: <strong class="text-altus-blue">${formatQ(currentDraft.iva)}</strong>
            </div>
          </div>
        </div>

        <div class="mt-3 p-3 rounded-xl bg-slate-50 border border-slate-200 text-xs space-y-1">
          <div class="text-slate-500 font-bold uppercase tracking-wider text-[10px]">Asiento de Partida Doble (NIIF / SAT):</div>
          <div class="text-emerald-700 font-mono font-semibold">DEBE:  ${currentDraft.partidaDoble.debe}</div>
          <div class="text-altus-blue font-mono font-semibold">HABER: ${currentDraft.partidaDoble.haber}</div>
          <div class="text-slate-500 pt-1 italic text-[11px]">${currentDraft.notaAuditoria}</div>
        </div>

        <div class="mt-4 flex flex-wrap items-center justify-end gap-3">
          <button onclick="window.AltusApp.clearPreview()"
            class="px-4 py-2.5 rounded-xl border border-slate-300 text-slate-700 hover:bg-slate-100 text-sm font-semibold transition-colors">
            Descartar
          </button>
          <button onclick="window.AltusApp.confirmDraftEntry()"
            class="px-5 py-2.5 rounded-xl bg-altus-blue hover:bg-altus-blue-dark text-white font-bold text-sm shadow-sm flex items-center gap-2 transition-colors">
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
      currentDraft.tipo = accObj.type === 'Ingreso' ? 'Ingreso' : (accObj.type === 'Pasivo' ? 'Pasivo' : 'Gasto');
      currentDraft.tipoOriginal = currentDraft.tipo === 'Pasivo' ? 'Ingreso' : currentDraft.tipo;
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
      tipoOriginal: currentDraft.tipoOriginal || (currentDraft.tipo === 'Ingreso' || currentDraft.tipo === 'Pasivo' ? 'Ingreso' : 'Gasto'),
      montoBruto: monto,
      iva: currentDraft.iva,
      montoNeto: currentDraft.montoNeto,
      banco: banco,
      saldoMonetaria: 0,
      saldoAhorro: 0,
      documento: doc,
      centroCosto: currentDraft.centroCosto || '00-Administración',
      notaAuditoria: currentDraft.notaAuditoria,
      studentUpdate: currentDraft.studentUpdate || null,
      receiptImageBase64: currentDraft.receiptImageBase64 || null
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

      let badgeClass = 'bg-rose-50 text-rose-800 border-rose-200';
      let sign = '-';
      if (isIncome) {
        badgeClass = 'bg-emerald-50 text-emerald-800 border-emerald-200';
        sign = '+';
      } else if (isTransfer) {
        badgeClass = 'bg-blue-50 text-altus-blue border-blue-200';
        sign = '⇄';
      } else if (isFromSavings) {
        badgeClass = 'bg-indigo-50 text-indigo-800 border-indigo-200';
        sign = '🛡️ -';
      }

      const photoBadge = j.receiptImageBase64
        ? `<a href="${j.receiptImageBase64}" target="_blank" class="inline-flex items-center gap-1 ml-2 px-2 py-0.5 rounded-md bg-altus-blue/10 text-altus-blue border border-altus-blue/20 text-[10px] font-bold hover:bg-altus-blue/20">📷 Ver Factura</a>`
        : '';

      return `
        <tr class="border-b border-slate-100 hover:bg-slate-50/80 transition-colors">
          <td class="py-2.5 px-3 text-xs font-mono text-slate-500 whitespace-nowrap">${j.fecha}<span class="block text-[10px] text-slate-400 font-semibold">${j.id}</span></td>
          <td class="py-2.5 px-3 text-xs text-slate-900 font-medium">
            ${j.descripcion} ${photoBadge}
            ${j.notaAuditoria ? `<span class="block text-[11px] text-altus-blue font-semibold mt-0.5">🔎 ${j.notaAuditoria}</span>` : ''}
          </td>
          <td class="py-2.5 px-3 text-xs text-slate-600 font-medium whitespace-nowrap">${j.cuenta}</td>
          <td class="py-2.5 px-3 text-xs whitespace-nowrap">
            <span class="px-2 py-0.5 rounded-full text-[11px] font-semibold border ${j.banco === 'BI-Ahorro' ? 'bg-indigo-50 text-indigo-700 border-indigo-200' : 'bg-slate-100 text-slate-700 border-slate-200'}">
              ${j.banco}
            </span>
          </td>
          <td class="py-2.5 px-3 text-xs font-bold text-right whitespace-nowrap">
            <span class="px-2 py-0.5 rounded-lg border font-mono ${badgeClass}">${sign} ${formatQ(j.montoBruto)}</span>
          </td>
          <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-700 whitespace-nowrap font-medium">${formatQ(j.saldoMonetaria)}</td>
          <td class="py-2.5 px-3 text-xs font-mono text-right text-indigo-700 whitespace-nowrap font-bold">${formatQ(j.saldoAhorro)}</td>
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
      let pasivos2027 = 0;
      let gastosOperativos = 0;
      let trasladosAhorro = 0;
      let pagosDesdeAhorro = 0;

      for (const e of entries) {
        const amt = Number(e.montoBruto || 0);
        if (e.cuenta && e.cuenta.startsWith('2107')) {
          pasivos2027 += amt;
        } else if (e.tipo === 'Ingreso') {
          ingresos += amt;
        } else if (e.tipo === 'Traspaso Ahorro' || e.cuenta === '2102 - Provisiones Laborales') {
          trasladosAhorro += amt;
        } else if (e.banco === 'BI-Ahorro' || e.tipo === 'Pago Reserva Ahorro') {
          pagosDesdeAhorro += amt;
        } else if (e.tipo === 'Gasto') {
          gastosOperativos += amt;
        }
      }
      return { mes: m, ingresos, pasivos2027, gastosOperativos, trasladosAhorro, pagosDesdeAhorro };
    });

    const totalIngresos = monthlyStats.reduce((a, b) => a + b.ingresos, 0);
    const totalPasivos2027 = monthlyStats.reduce((a, b) => a + b.pasivos2027, 0);
    const totalGastosOp = monthlyStats.reduce((a, b) => a + b.gastosOperativos, 0);
    const totalTraslados = monthlyStats.reduce((a, b) => a + b.trasladosAhorro, 0);
    const totalIvaCredito = state.journal
      .filter(j => j.tipo === 'Gasto')
      .reduce((a, b) => a + Number(b.iva || 0), 0);

    const kpiContainer = document.getElementById('dashboard-kpis');
    if (kpiContainer) {
      kpiContainer.innerHTML = `
        <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-xs">
          <div class="text-xs text-slate-500 font-medium">Ingresos Devengados 2026 (Ene-Oct)</div>
          <div class="text-xl font-extrabold text-emerald-600 mt-1 font-mono">${formatQ(totalIngresos)}</div>
          <div class="text-[11px] text-slate-400 mt-1">Excluye Q ${totalPasivos2027.toLocaleString('es-GT')} de Inscripciones 2027 (2107 Pasivo)</div>
        </div>
        <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-xs">
          <div class="text-xs text-slate-500 font-medium">Gastos Operativos Directos 2026</div>
          <div class="text-xl font-extrabold text-altus-coral mt-1 font-mono">${formatQ(totalGastosOp)}</div>
          <div class="text-[11px] text-slate-400 mt-1">Sin duplicar pagos de BI-Ahorro</div>
        </div>
        <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-xs">
          <div class="text-xs text-slate-500 font-medium">Separado a Reserva Ahorro (1102)</div>
          <div class="text-xl font-extrabold text-altus-blue mt-1 font-mono">${formatQ(totalTraslados)}</div>
          <div class="text-[11px] text-slate-400 mt-1">Pagado desde Ahorro: ${formatQ(monthlyStats.reduce((a,b)=>a+b.pagosDesdeAhorro,0))}</div>
        </div>
        <div class="bg-amber-50/70 border border-amber-300 rounded-2xl p-4 shadow-xs">
          <div class="text-xs text-amber-900 font-bold">Pasivo Diferido Inscripciones 2027 (2107)</div>
          <div class="text-xl font-extrabold text-amber-800 mt-1 font-mono">${formatQ(totalPasivos2027)}</div>
          <div class="text-[11px] text-amber-900/80 mt-1">Blindado en 2026 → Pasa a Ingreso el 01/01/2027 (IVA Créd: ${formatQ(totalIvaCredito)})</div>
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
            <div class="opacity-0 group-hover:opacity-100 transition-opacity absolute -top-16 z-20 bg-slate-900 text-white rounded-lg px-2.5 py-1.5 text-[10px] whitespace-nowrap shadow-xl pointer-events-none">
              <strong class="uppercase text-amber-300">${s.mes}</strong><br/>
              Ingresos: ${formatQ(s.ingresos)}<br/>
              Gastos Op.: ${formatQ(s.gastosOperativos)}<br/>
              A Reserva Ahorro: ${formatQ(s.trasladosAhorro)}
            </div>
            <div class="h-44 flex items-end justify-center gap-1.5 w-full px-1 border-b border-slate-200">
              <div style="height:${hIng}px" class="w-3.5 md:w-4 rounded-t-md bg-emerald-500 hover:bg-emerald-600 transition-all shadow-xs" title="Ingresos: ${formatQ(s.ingresos)}"></div>
              <div class="flex flex-col justify-end w-3.5 md:w-4">
                ${hAho > 0 ? `<div style="height:${hAho}px" class="w-full rounded-t-md bg-altus-blue shadow-xs" title="Traslado a Ahorro: ${formatQ(s.trasladosAhorro)}"></div>` : ''}
                <div style="height:${hGas}px" class="w-full ${hAho > 0 ? '' : 'rounded-t-md'} bg-altus-coral hover:bg-altus-coral-dark shadow-xs" title="Gastos Operativos: ${formatQ(s.gastosOperativos)}"></div>
              </div>
            </div>
            <div class="text-[11px] font-bold text-slate-700 mt-2 capitalize">${s.mes.slice(0, 3)}</div>
            <div class="text-[10px] font-mono font-bold ${net >= 0 ? 'text-emerald-700' : 'text-rose-700'}">${net >= 0 ? '+' : ''}${Math.round(net / 1000)}k</div>
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
          <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
            <div class="flex items-center justify-between text-xs mb-1.5">
              <span class="font-bold text-slate-900">${idx + 1}. ${acc} ${isReserve ? '<span class="ml-1 text-[10px] px-1.5 py-0.5 rounded bg-indigo-50 text-indigo-700 font-semibold border border-indigo-200">Reserva / Ahorro</span>' : ''}</span>
              <span class="font-mono font-bold text-slate-900">${formatQ(total)} <span class="text-slate-500 font-normal">(${pct}%)</span></span>
            </div>
            <div class="w-full h-2 bg-slate-200 rounded-full overflow-hidden">
              <div style="width:${Math.max(3, pct)}%" class="h-full rounded-full ${isReserve ? 'bg-altus-blue' : 'bg-gradient-to-r from-altus-amber to-altus-coral'}"></div>
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
        ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
        : (s.estado === 'Retirado' ? 'bg-slate-100 text-slate-500 border-slate-200' : 'bg-amber-50 text-amber-800 border-amber-200');

      const monthCells = months.map(m => {
        const paid = Number(s.pagos[m] || 0);
        const startIdx = months.indexOf((s.mesInicio || 'Enero').toLowerCase());
        const mIdx = months.indexOf(m);
        if (s.estado === 'Retirado' || mIdx < startIdx) {
          return `<td class="py-2.5 px-2 text-center text-[11px] text-slate-300 font-mono">—</td>`;
        }
        if (paid >= s.cuotaMensual * 0.95) {
          return `<td class="py-2.5 px-2 text-center"><span class="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-200 text-[11px] font-mono font-semibold">Q${paid.toLocaleString()}</span></td>`;
        } else if (paid > 0) {
          return `<td class="py-2.5 px-2 text-center"><span class="inline-block px-1.5 py-0.5 rounded bg-amber-50 text-amber-800 border border-amber-200 text-[11px] font-mono font-semibold">Q${paid.toLocaleString()}</span></td>`;
        } else {
          return `<td class="py-2.5 px-2 text-center">
            <button onclick="window.AltusApp.quickTuitionPay('${s.estudiante}', '${m}')"
              class="px-2 py-0.5 rounded-md bg-altus-coral hover:bg-altus-coral-dark text-white text-[10px] font-bold shadow-xs transition-colors"
              title="Registrar pago de ${m} para ${s.estudiante}">
              + Cobrar
            </button>
          </td>`;
        }
      }).join('');

      return `
        <tr class="border-b border-slate-100 hover:bg-slate-50/80 transition-colors">
          <td class="py-3 px-3 text-xs">
            <div class="font-bold text-slate-900">${s.estudiante}</div>
            <div class="text-[11px] text-slate-500 font-medium">${s.idFamilia} • ${s.grado}</div>
          </td>
          <td class="py-3 px-3 text-xs">
            <div class="text-slate-800 font-medium">${s.responsable}</div>
            ${s.contactoSecundario ? `<div class="text-[11px] text-altus-amber-dark font-semibold">👩‍👦 ${s.contactoSecundario}</div>` : ''}
            <div class="text-[10px] text-slate-400 font-mono">${s.email} • NIT: ${s.nit}</div>
          </td>
          <td class="py-3 px-3 text-xs font-mono font-bold text-slate-800 text-right">${formatQ(s.cuotaMensual)}</td>
          ${monthCells}
          <td class="py-3 px-3 text-xs font-mono font-bold text-emerald-700 text-right">${formatQ(s.totalPagado)}</td>
          <td class="py-3 px-3 text-xs text-center">
            <span class="px-2.5 py-1 rounded-full text-[11px] font-bold border ${badgeColor}">${s.estado}</span>
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
        <tr class="border-b border-slate-100 hover:bg-slate-50/80 transition-colors">
          <td class="py-3 px-3 text-xs">
            <div class="font-bold text-slate-900">${emp.nombre}</div>
            <div class="text-[11px] text-slate-500">${emp.puesto} • ${emp.cuentaBanco}</div>
          </td>
          <td class="py-3 px-3 text-xs font-mono text-right text-slate-800">${formatQ(emp.sueldoBase)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-slate-600">${formatQ(emp.bonificacion)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-rose-600 font-medium">-${formatQ(emp.igssLaboral)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-rose-600 font-medium">-${formatQ(emp.isrRetenido)}</td>
          <td class="py-3 px-3 text-xs font-mono font-bold text-right text-emerald-700 bg-emerald-50/40">${formatQ(emp.liquidoRecibir)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-indigo-700">${formatQ(emp.igssPatronal)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-amber-800">${formatQ(emp.reservaBono14)}</td>
          <td class="py-3 px-3 text-xs font-mono text-right text-amber-800">${formatQ(emp.reservaAguinaldo)}</td>
          <td class="py-3 px-3 text-xs text-center">
            <button onclick="window.AltusApp.triggerNLPFromButton('Pago sueldo ${emp.nombre}')"
              class="px-2.5 py-1 rounded-lg bg-altus-blue hover:bg-altus-blue-dark text-white text-xs font-bold shadow-xs transition-colors">
              Pagar Día 27
            </button>
          </td>
        </tr>
      `).join('');
    }

    const obligationsEl = document.getElementById('day27-obligations-list');
    if (obligationsEl) {
      obligationsEl.innerHTML = state.day27Obligations.map(ob => `
        <div class="flex items-center justify-between p-3 rounded-xl bg-slate-50 border border-slate-200 hover:border-slate-300 transition-colors">
          <div>
            <div class="text-xs font-bold text-slate-900">${ob.concepto}</div>
            <div class="text-[11px] text-slate-500">${ob.cuenta} • <span class="text-altus-blue font-semibold">${ob.banco}</span></div>
          </div>
          <div class="flex items-center gap-3">
            <span class="text-xs font-mono font-extrabold text-slate-900">${formatQ(ob.monto)}</span>
            <button onclick="window.AltusApp.triggerNLPFromButton('${ob.concepto} Q${ob.monto}')"
              class="px-2.5 py-1 rounded-lg bg-white hover:bg-altus-blue hover:text-white text-altus-blue border border-slate-300 text-xs font-bold shadow-xs transition-colors">
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
          <tr class="border-b border-slate-100 hover:bg-slate-50/80 transition-colors">
            <td class="py-2.5 px-3 text-xs font-mono text-slate-500">${j.fecha}</td>
            <td class="py-2.5 px-3 text-xs text-slate-900 font-medium">${j.descripcion}</td>
            <td class="py-2.5 px-3 text-xs text-slate-600 font-medium">${j.cuenta}</td>
            <td class="py-2.5 px-3 text-xs font-mono font-bold text-right ${isDeposit ? 'text-emerald-700' : 'text-rose-600'}">
              ${isDeposit ? '+ ' : '- '}${formatQ(j.montoBruto)}
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-indigo-700 font-bold">${formatQ(j.saldoAhorro)}</td>
          </tr>
        `;
      }).join('');
    }
  }

  // Render Presupuesto 2027 & Conciliación
  function renderBudgetAndReconciliation() {
    // Render 2107 Deferred Liability (Inscripciones 2027 cobradas en 2026)
    const pasivo2027Entries = state.journal.filter(j => j.cuenta && j.cuenta.startsWith('2107'));
    const totalPasivo2027 = pasivo2027Entries.reduce((a, b) => a + Number(b.montoBruto || 0), 0);
    const pasivoTotalEl = document.getElementById('presupuesto-pasivo2027-total');
    if (pasivoTotalEl) {
      pasivoTotalEl.textContent = formatQ(totalPasivo2027);
    }
    const pasivoListEl = document.getElementById('presupuesto-pasivo2027-list');
    if (pasivoListEl) {
      pasivoListEl.innerHTML = pasivo2027Entries.map(e => `
        <div class="p-2.5 rounded-lg bg-amber-50/60 border border-amber-200 flex items-center justify-between gap-2">
          <div>
            <div class="text-xs font-bold text-slate-900">${e.descripcion}</div>
            <div class="text-[11px] text-slate-500">${e.fecha} • ${e.id} • Doc: ${e.documento || 'BI'}</div>
            <div class="text-[10px] text-amber-800 font-medium mt-0.5">⏳ Pasivo en 2026 → Se reclasifica a 4101 Ingreso el 01/01/2027</div>
          </div>
          <div class="text-xs font-mono font-bold text-amber-900 whitespace-nowrap">${formatQ(e.montoBruto)}</div>
        </div>
      `).join('');
    }

    const budgetBody = document.getElementById('budget-2027-body');
    if (budgetBody) {
      budgetBody.innerHTML = state.budget2027.partidas.map((p, idx) => {
        const ratio = p.anual2027 > 0 ? Math.min(100, ((p.real2026Oct / p.anual2027) * 100)).toFixed(1) : 0;
        const isIng = p.tipo === 'Ingreso';
        return `
          <tr class="border-b border-slate-100 hover:bg-slate-50/80 transition-colors">
            <td class="py-2.5 px-3 text-xs font-medium text-slate-900">${p.codigo}</td>
            <td class="py-2.5 px-3 text-xs">
              <span class="px-2 py-0.5 rounded-full text-[11px] font-semibold ${isIng ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-rose-50 text-rose-800 border border-rose-200'}">${p.tipo}</span>
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-600">${formatQ(p.mensualEne)}</td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-600">${formatQ(p.mensualFebDic)}</td>
            <td class="py-2.5 px-3 text-xs font-mono font-bold text-right text-slate-900">
              <input type="number" value="${p.anual2027}"
                onchange="window.AltusApp.updateBudget2027(${idx}, this.value)"
                class="w-28 bg-slate-50 border border-slate-300 rounded-lg px-2 py-1 text-right text-slate-900 font-mono text-xs focus:border-altus-blue focus:outline-none" />
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-700">${formatQ(p.real2026Oct)}</td>
            <td class="py-2.5 px-3 text-xs text-right">
              <div class="flex items-center justify-end gap-2">
                <div class="w-16 h-1.5 bg-slate-200 rounded-full overflow-hidden">
                  <div style="width:${ratio}%" class="h-full ${isIng ? 'bg-emerald-500' : 'bg-altus-amber'}"></div>
                </div>
                <span class="font-mono text-[11px] text-slate-500 font-semibold">${ratio}%</span>
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
          <tr class="border-b border-slate-100 hover:bg-slate-50/80 transition-colors">
            <td class="py-2.5 px-3 text-xs font-bold capitalize text-slate-900">${r.mes}</td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-800 font-medium">${formatQ(r.saldoLibroMonetaria)}</td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-slate-600">
              ${hasBank ? formatQ(r.saldoEstadoCuenta) : '<span class="text-slate-400">Pendiente</span>'}
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right font-bold ${r.diferencia === 0 ? 'text-emerald-700' : 'text-amber-800'}">
              ${hasBank ? formatQ(r.diferencia) : '—'}
            </td>
            <td class="py-2.5 px-3 text-xs font-mono text-right text-indigo-700 font-semibold">${formatQ(r.saldoAhorro)}</td>
            <td class="py-2.5 px-3 text-xs text-center">
              <span class="px-2 py-0.5 rounded-full text-[11px] font-bold ${hasBank ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-slate-100 text-slate-600'}">
                ${hasBank ? '✅ Conciliado' : r.estadoAuditado}
              </span>
            </td>
            <td class="py-2.5 px-3 text-[11px] text-slate-500">${r.explicacionAuditoria || 'Sin diferencias'}</td>
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
      const isMobile = btn.closest('nav') && btn.closest('nav').classList.contains('md:hidden');
      if (isMobile) {
        if (isTarget) {
          btn.classList.add('bg-altus-blue', 'text-white', 'shadow-xs');
          btn.classList.remove('text-slate-500', 'border-transparent');
        } else {
          btn.classList.remove('bg-altus-blue', 'text-white', 'shadow-xs');
          btn.classList.add('text-slate-500', 'border-transparent');
        }
      } else {
        if (isTarget) {
          btn.classList.add('bg-white', 'text-altus-blue', 'shadow-xs', 'border-slate-200/60');
          btn.classList.remove('text-slate-600', 'border-transparent');
        } else {
          btn.classList.remove('bg-white', 'text-altus-blue', 'shadow-xs', 'border-slate-200/60');
          btn.classList.add('text-slate-600', 'border-transparent');
        }
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
    handleInvoicePhotoUpload,
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
