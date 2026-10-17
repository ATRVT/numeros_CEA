/**
 * ============================================================================
 * MOTOR DE LENGUAJE NATURAL CONTABLE Y AUDITORÍA SAT (GUATEMALA)
 * Centro Educativo ALTUS - Reconocimiento de Familias, Cuentas SAT, IVA y Bancos
 * ============================================================================
 */

window.AltusNLPEngine = (function () {
  const MONTHS = [
    'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
  ];

  function normalize(str) {
    return String(str || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
  }

  function extractAmount(text) {
    // Exclude numbers that follow words like "bono", "dia", "mes", "boleta", "factura", "fel", "nit", "grado"
    const cleaned = text
      .replace(/bono\s*14/gi, 'BONO_CATORCE')
      .replace(/dia\s*27/gi, 'DIA_VEINTISIETE')
      .replace(/(?:boleta|factura|fel|doc|autorizacion|ref|no\.?|#)\s*[:#-]?\s*[a-z0-9-]+/gi, '');

    const qMatch = cleaned.match(/q\.?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]+)?|[0-9]+(?:\.[0-9]+)?)/i);
    if (qMatch) {
      return parseFloat(qMatch[1].replace(/,/g, ''));
    }
    const numMatches = cleaned.match(/\b([0-9]{1,3}(?:,[0-9]{3})+(?:\.[0-9]+)?|[0-9]{2,6}(?:\.[0-9]+)?)\b/g);
    if (numMatches && numMatches.length > 0) {
      for (const m of numMatches) {
        const val = parseFloat(m.replace(/,/g, ''));
        if (val !== 2026 && val !== 2027 && val > 0) {
          return val;
        }
      }
    }
    return null;
  }

  function extractDocRef(text) {
    const match = text.match(/(?:boleta|factura|fel|doc|comprobante|aut|ref|cheque)\s*(?:no\.?|#|:)?\s*([A-Z0-9-]{3,20})/i);
    return match ? match[1].toUpperCase() : '';
  }

  function extractMonth(text, fallbackMonth) {
    const norm = normalize(text);
    for (const m of MONTHS) {
      if (norm.includes(normalize(m))) {
        return m;
      }
    }
    return fallbackMonth || 'octubre';
  }

  function findStudent(text, students) {
    const normText = normalize(text);
    for (const s of students) {
      const candidates = [
        s.estudiante,
        s.responsable,
        s.contactoSecundario,
        ...(s.aliases || [])
      ].filter(Boolean);

      for (const alias of candidates) {
        const normAlias = normalize(alias);
        if (normAlias.length >= 3 && normText.includes(normAlias)) {
          return s;
        }
        // Check first name match if distinctive (e.g. Pamela, Mateo, Giulia, Andrei, Gunther, Rodrigo, Dulce, Marjorie, Eddy, Fernanda)
        const firstToken = normAlias.split(/\s+/)[0];
        if (firstToken.length >= 4 && new RegExp('\\b' + firstToken + '\\b', 'i').test(normText)) {
          return s;
        }
      }
    }
    return null;
  }

  function findPendingMonthForStudent(student, preferredMonth) {
    if (!student || !student.pagos) return preferredMonth || 'octubre';
    const startIdx = Math.max(0, MONTHS.indexOf(normalize(student.mesInicio || 'enero')));
    // If user explicitly mentioned a month in text, honor it; otherwise find earliest unpaid month up to current or next month
    if (preferredMonth) return preferredMonth;
    for (let i = startIdx; i <= 9; i++) { // Up to octubre (index 9)
      const m = MONTHS[i];
      if ((student.pagos[m] || 0) < student.cuotaMensual * 0.9) {
        return m;
      }
    }
    // If paid through octubre, suggest noviembre (or warn if octubre already paid)
    for (let i = 9; i < MONTHS.length; i++) {
      const m = MONTHS[i];
      if ((student.pagos[m] || 0) === 0) {
        return m;
      }
    }
    return 'octubre';
  }

  function parseNaturalInput(rawText, appData) {
    const text = String(rawText || '').trim();
    if (!text) return null;

    const norm = normalize(text);
    const todayStr = new Date().toISOString().slice(0, 10);
    const explicitMonth = MONTHS.find(m => norm.includes(normalize(m))) || null;
    const extractedAmount = extractAmount(text);
    const extractedDoc = extractDocRef(text);

    // 1. Check if it's a QUESTION / CONSULTA RÁPIDA
    const isQuery = norm.startsWith('?') || norm.startsWith('¿') ||
      /\b(quien|quienes|cuanto|cual|como va|estado de|pendiente|faltan|morosidad|saldo|resumen|grafica|top)\b/.test(norm) &&
      !/\b(pago|pagaron|ya pago|compre|pague|transferi|separar|trasladar|registrar)\b/.test(norm);

    if (isQuery) {
      return buildQueryResponse(text, norm, appData);
    }

    // 2. Check if it mentions a Student / Parent (e.g., "ya pagó Pamela")
    const matchedStudent = findStudent(text, appData.students || []);
    if (matchedStudent) {
      const isInscription = /\b(inscripcion|matricula)\b/.test(norm);
      const isMaterials = /\b(material|materiales|seguro|agenda|uniforme|cuenta ajena)\b/.test(norm);
      const targetMonth = explicitMonth || findPendingMonthForStudent(matchedStudent, null);
      const defaultAmt = isInscription
        ? (matchedStudent.inscripcionAcordada || 10300)
        : (matchedStudent.cuotaMensual || 9300);
      const monto = extractedAmount !== null ? extractedAmount : defaultAmt;

      let cuenta = '4102 - Colegiaturas Mensuales';
      let tipo = 'Ingreso';
      let desc = `Cobro mensualidad ${targetMonth} ${matchedStudent.estudiante}`;
      let partidaDoble = {
        debe: `1101 - Banco Operativo (BI-Monetaria): Q ${monto.toFixed(2)}`,
        haber: `4102 - Colegiaturas Mensuales (Exento IVA Art. 8): Q ${monto.toFixed(2)}`
      };

      if (isInscription) {
        cuenta = '4101 - Inscripciones';
        desc = `Cobro inscripción ${matchedStudent.estudiante} (${matchedStudent.idFamilia})`;
        partidaDoble.haber = `4101 - Inscripciones (Exento IVA Art. 8): Q ${monto.toFixed(2)}`;
      } else if (isMaterials) {
        cuenta = '2106 - Acreedores por Cuenta Ajena';
        tipo = 'Pasivo';
        desc = `Cobro cuenta ajena materiales/seguro ${matchedStudent.estudiante}`;
        partidaDoble.haber = `2106 - Acreedores por Cuenta Ajena: Q ${monto.toFixed(2)}`;
      }

      const alreadyPaidMonth = (matchedStudent.pagos && matchedStudent.pagos[targetMonth]) || 0;
      const notaSat = `Recibo Exento SAT (Centro Educativo Art. 8 Ley IVA) | Familia: ${matchedStudent.idFamilia} | Responsable: ${matchedStudent.responsable}${matchedStudent.contactoSecundario ? ' / ' + matchedStudent.contactoSecundario : ''} (NIT: ${matchedStudent.nit})`;

      return {
        kind: 'transaction',
        confidence: 'Alta (Alumno / Familia Identificada)',
        fecha: todayStr,
        mes: targetMonth,
        descripcion: desc,
        cuenta: cuenta,
        tipo: tipo,
        montoBruto: monto,
        iva: 0,
        montoNeto: monto,
        banco: 'BI- Monetaria',
        documento: extractedDoc || `FEL-EX-${Math.floor(100000 + Math.random() * 900000)}`,
        centroCosto: matchedStudent.grado.toLowerCase().includes('kinder') || matchedStudent.grado.toLowerCase().includes('pre')
          ? '01-Pre-Primaria'
          : '02-Primaria',
        notaAuditoria: notaSat,
        partidaDoble: partidaDoble,
        studentUpdate: (!isInscription && !isMaterials) ? {
          idFamilia: matchedStudent.idFamilia,
          estudiante: matchedStudent.estudiante,
          responsable: matchedStudent.responsable,
          contacto: matchedStudent.contactoSecundario || matchedStudent.email,
          mes: targetMonth,
          alreadyPaidMonth: alreadyPaidMonth
        } : null
      };
    }

    // 3. Check if it's the Monthly Transfer from Monetary to Savings Account (Monetaria -> Ahorro)
    if (/\b(separar|trasladar|transferir|pasar|apartar|provision|provisiones|reserva)\b/.test(norm) &&
        /\b(ahorro|provision|provisiones|reserva|laboral|laborales|mes)\b/.test(norm)) {
      const monto = extractedAmount !== null ? extractedAmount : 11551.09;
      const mes = explicitMonth || 'octubre';
      return {
        kind: 'transaction',
        confidence: 'Alta (Traspaso Reglado Monetaria -> Ahorro)',
        fecha: todayStr,
        mes: mes,
        descripcion: `Traslado mensual fondo provisiones laborales e IGSS (${mes})`,
        cuenta: '2102 - Provisiones Laborales',
        tipo: 'Traspaso Ahorro',
        tipoOriginal: 'Gasto',
        montoBruto: monto,
        iva: 0,
        montoNeto: monto,
        banco: 'BI- Monetaria',
        documento: extractedDoc || `TR-AHO-${todayStr.replace(/-/g, '')}`,
        centroCosto: '00-Administración',
        notaAuditoria: 'Traspaso interno BI-Monetaria (1101) -> BI-Ahorro (1102): Cubre IGSS (Q4,278.00) + Bono 14 (Q2,299.08) + Aguinaldo (Q2,299.08) + Indemnización (Q2,299.08) + ISR (Q375.85).',
        partidaDoble: {
          debe: `1102 - Banco Reservas y Provisiones (BI-Ahorro): +Q ${monto.toFixed(2)}`,
          haber: `1101 - Banco Operativo (BI-Monetaria): -Q ${monto.toFixed(2)}`
        }
      };
    }

    // 4. Check if it's IGSS Payment (Paid from Savings Account BI-Ahorro!)
    if (/\b(igss|iggs|planilla igss|seguro social)\b/.test(norm)) {
      const monto = extractedAmount !== null ? extractedAmount : 4278.00;
      const mes = explicitMonth || 'octubre';
      return {
        kind: 'transaction',
        confidence: 'Alta (Pago de IGSS desde Cuenta de Ahorro)',
        fecha: todayStr,
        mes: mes,
        descripcion: `Pago planilla IGSS patronal (10.67%) y laboral (4.83%) - ${mes}`,
        cuenta: '2103 - IGSS por Pagar',
        tipo: 'Pago Reserva Ahorro',
        tipoOriginal: 'Gasto',
        montoBruto: monto,
        iva: 0,
        montoNeto: monto,
        banco: 'BI-Ahorro',
        documento: extractedDoc || `IGSS-${todayStr.replace(/-/g, '')}`,
        centroCosto: '00-Administración',
        notaAuditoria: 'Pagado desde BI-Ahorro (1102). Liquida pasivo 2103 IGSS por Pagar sin duplicar gasto operativo.',
        partidaDoble: {
          debe: `2103 - IGSS por Pagar (Pasivo Laboral): Q ${monto.toFixed(2)}`,
          haber: `1102 - Banco Reservas y Provisiones (BI-Ahorro): -Q ${monto.toFixed(2)}`
        }
      };
    }

    // 5. Check if it's Bono 14 or Aguinaldo (Paid from Savings Account BI-Ahorro!)
    if (/\b(bono\s*14|bono catorce|aguinaldo)\b/.test(norm)) {
      const isAguinaldo = norm.includes('aguinaldo');
      const staffMatch = (appData.staff || []).find(emp => norm.includes(normalize(emp.nombre.split(' ')[0])));
      const monto = extractedAmount !== null ? extractedAmount : (staffMatch ? staffMatch.sueldoBase : 27600.00);
      const mes = explicitMonth || (isAguinaldo ? 'diciembre' : 'julio');
      const codigo = isAguinaldo ? '5105 - Aguinaldo' : '5104 - Bono 14';
      const decreto = isAguinaldo ? 'Decreto 76-78' : 'Decreto 42-92';
      const targetName = staffMatch ? staffMatch.nombre : 'Planilla Completa Docentes';

      return {
        kind: 'transaction',
        confidence: `Alta (Prestación de Ley ${decreto} desde BI-Ahorro)`,
        fecha: todayStr,
        mes: mes,
        descripcion: `Pago ${isAguinaldo ? 'Aguinaldo' : 'Bono 14'} - ${targetName}`,
        cuenta: codigo,
        tipo: 'Pago Reserva Ahorro',
        tipoOriginal: 'Gasto',
        montoBruto: monto,
        iva: 0,
        montoNeto: monto,
        banco: 'BI-Ahorro',
        documento: extractedDoc || `PREST-${todayStr.replace(/-/g, '')}`,
        centroCosto: '00-Administración',
        notaAuditoria: `Pagado desde BI-Ahorro (1102) conforme ${decreto}. Liquida provisión acumulada en 2102.`,
        partidaDoble: {
          debe: `2102 - Provisiones Laborales (${codigo}): Q ${monto.toFixed(2)}`,
          haber: `1102 - Banco Reservas y Provisiones (BI-Ahorro): -Q ${monto.toFixed(2)}`
        }
      };
    }

    // 6. Check if it's Individual Staff Salary or Full Payroll
    const staffMember = (appData.staff || []).find(emp => {
      const first = normalize(emp.nombre.split(' ')[0]);
      return norm.includes(first);
    });
    if (staffMember || /\b(sueldo|salario|nomina|planilla|quincena)\b/.test(norm)) {
      const mes = explicitMonth || 'octubre';
      if (staffMember) {
        const monto = extractedAmount !== null ? extractedAmount : staffMember.liquidoRecibir;
        return {
          kind: 'transaction',
          confidence: `Alta (Nómina Docente: ${staffMember.nombre})`,
          fecha: todayStr,
          mes: mes,
          descripcion: `Pago sueldo ${mes} ${staffMember.nombre} (${staffMember.puesto})`,
          cuenta: '5101 - Sueldos y Salarios',
          tipo: 'Gasto',
          montoBruto: monto,
          iva: 0,
          montoNeto: monto,
          banco: 'BI- Monetaria',
          documento: extractedDoc || `NOM-${todayStr.replace(/-/g, '')}`,
          centroCosto: staffMember.puesto === 'Directora' ? '00-Administración' : '01-Pre-Primaria / 02-Primaria',
          notaAuditoria: `Sueldo líquido Día 27 (Base Q${staffMember.sueldoBase} + Bonif. Q${staffMember.bonificacion} - IGSS Q${staffMember.igssLaboral} - ISR Q${staffMember.isrRetenido.toFixed(2)}) | Cuenta: ${staffMember.cuentaBanco}`,
          partidaDoble: {
            debe: `5101 - Sueldos y Salarios: Q ${monto.toFixed(2)}`,
            haber: `1101 - Banco Operativo (BI-Monetaria): -Q ${monto.toFixed(2)}`
          }
        };
      } else {
        const monto = extractedAmount !== null ? extractedAmount : 27141.07;
        return {
          kind: 'transaction',
          confidence: 'Alta (Pago Total Planilla Día 27)',
          fecha: todayStr,
          mes: mes,
          descripcion: `Pago planilla sueldos netos Día 27 (${mes}) - 5 colaboradoras`,
          cuenta: '5101 - Sueldos y Salarios',
          tipo: 'Gasto',
          montoBruto: monto,
          iva: 0,
          montoNeto: monto,
          banco: 'BI- Monetaria',
          documento: extractedDoc || `NOM-27-${mes.toUpperCase()}`,
          centroCosto: '00-Administración',
          notaAuditoria: 'Líquido total planilla Día 27 (Lucía, Jessica, Geovanna, Suly y Melissa). Exento de IVA.',
          partidaDoble: {
            debe: `5101 - Sueldos y Salarios: Q ${monto.toFixed(2)}`,
            haber: `1101 - Banco Operativo (BI-Monetaria): -Q ${monto.toFixed(2)}`
          }
        };
      }
    }

    // 7. Expense & Income Keyword Classification Rules (SAT Catalog)
    const rules = [
      { keywords: ['isr', 'declaraguate', 'retencion'], cuenta: '2105 - ISR Retenciones por Pagar', tipo: 'Gasto', afecto: false, defaultAmt: 375.85, descPrefix: 'Pago ISR Retenciones Empleados SAT' },
      { keywords: ['waybi', 'fundacion', 'donacion a'], cuenta: '5208 - Donaciones', tipo: 'Gasto', afecto: true, defaultAmt: 1700.00, descPrefix: 'Donación mensual Fundación Waybi' },
      { keywords: ['recibimos donacion', 'donacion recibida', 'aporte'], cuenta: '4203 - Donaciones recibidas', tipo: 'Ingreso', afecto: false, defaultAmt: 5000.00, descPrefix: 'Donación recibida para suministros' },
      { keywords: ['isa gonzalez', 'redes sociales', 'rrss', 'marketing', 'publicidad', 'facebook', 'instagram', 'anuncio'], cuenta: '5204 - Marketing y Publicidad', tipo: 'Gasto', afecto: true, defaultAmt: 1600.00, descPrefix: 'Gasto de Marketing y Publicidad' },
      { keywords: ['sergio', 'contador', 'lourdes', 'honorarios', 'factura profesional', 'asesoria'], cuenta: '5209 - Honorarios profesionales', tipo: 'Gasto', afecto: true, defaultAmt: 350.00, descPrefix: 'Pago de Honorarios Profesionales' },
      { keywords: ['alquiler', 'renta', 'arrendamiento', 'local'], cuenta: '5201 - Alquiler de Instalaciones', tipo: 'Gasto', afecto: true, defaultAmt: 4500.00, descPrefix: 'Pago Alquiler de Instalaciones' },
      { keywords: ['luz', 'agua', 'eegsa', 'empagua', 'energia', 'electrica'], cuenta: '5202 - Energía Eléctrica y Agua', tipo: 'Gasto', afecto: true, defaultAmt: 450.00, descPrefix: 'Pago Servicios Básicos (Energía/Agua)' },
      { keywords: ['internet', 'celular', 'telefono', 'tigo', 'claro', 'telecomunicaciones', 'wifi'], cuenta: '5212 - Telecomunicaciones e internet', tipo: 'Gasto', afecto: true, defaultAmt: 285.00, descPrefix: 'Pago Telecomunicaciones e Internet ALTUS' },
      { keywords: ['papeleria', 'utiles', 'suministros', 'arimany', 'platino', 'marcadores', 'hojas', 'impresion', 'material didactico', 'limpieza'], cuenta: '5203 - Papelería, útiles y suministros', tipo: 'Gasto', afecto: true, defaultAmt: 250.00, descPrefix: 'Compra Papelería, Útiles y Suministros' },
      { keywords: ['excursion', 'actividad', 'evento', 'celebracion', 'dia de la madre', 'dia del nino', 'convivio'], cuenta: '5205 - Actividades y excursiones', tipo: 'Gasto', afecto: true, defaultAmt: 300.00, descPrefix: 'Gasto Actividades y Excursiones Escolares' },
      { keywords: ['pupitre', 'computadora', 'escritorio', 'silla', 'pizarron', 'mobiliario', 'equipo', 'impresora'], cuenta: '1501 - Mobiliario y Equipo', tipo: 'Gasto', afecto: true, defaultAmt: 1500.00, descPrefix: 'Adquisición Mobiliario y Equipo Escolar' },
      { keywords: ['uniforme', 'libro', 'venta de'], cuenta: '4201 - Venta de Libros y Uniformes', tipo: 'Ingreso', afecto: true, defaultAmt: 400.00, descPrefix: 'Ingreso por Venta de Libros/Uniformes' },
      { keywords: ['marn', 'dideduc', 'mineduc', 'iusi', 'timbre', 'arbitrio', 'licencia'], cuenta: '5210 - Impuestos, tasas y arbitrios', tipo: 'Gasto', afecto: false, defaultAmt: 250.00, descPrefix: 'Pago Tasas/Trámites Gubernamentales' }
    ];

    for (const rule of rules) {
      if (rule.keywords.some(k => norm.includes(k))) {
        const monto = extractedAmount !== null ? extractedAmount : rule.defaultAmt;
        const mes = explicitMonth || 'octubre';
        const iva = rule.afecto ? +(monto - monto / 1.12).toFixed(2) : 0;
        const neto = +(monto - iva).toFixed(2);
        const usaAhorro = norm.includes('ahorro');
        return {
          kind: 'transaction',
          confidence: `Alta (Partida SAT ${rule.cuenta.split(' ')[0]})`,
          fecha: todayStr,
          mes: mes,
          descripcion: text.length > 6 ? text.charAt(0).toUpperCase() + text.slice(1) : rule.descPrefix,
          cuenta: rule.cuenta,
          tipo: usaAhorro ? 'Pago Reserva Ahorro' : rule.tipo,
          montoBruto: monto,
          iva: iva,
          montoNeto: neto,
          banco: usaAhorro ? 'BI-Ahorro' : 'BI- Monetaria',
          documento: extractedDoc || `FEL-${Math.floor(100000 + Math.random() * 900000)}`,
          centroCosto: '00-Administración',
          notaAuditoria: rule.afecto
            ? `Operación Afecta (12% IVA = Q ${iva.toFixed(2)} | Neto = Q ${neto.toFixed(2)}). Requiere Factura FEL con NIT de Centro Educativo ALTUS.`
            : `Operación Exenta de IVA según normativa SAT.`,
          partidaDoble: {
            debe: rule.tipo === 'Ingreso'
              ? `1101 - Banco Operativo (BI-Monetaria): +Q ${monto.toFixed(2)}`
              : `${rule.cuenta}: Q ${neto.toFixed(2)}${iva > 0 ? ` + 1104 IVA Crédito Fiscal: Q ${iva.toFixed(2)}` : ''}`,
            haber: rule.tipo === 'Ingreso'
              ? `${rule.cuenta}: Q ${neto.toFixed(2)}${iva > 0 ? ` + 2104 IVA Débito Fiscal: Q ${iva.toFixed(2)}` : ''}`
              : `${usaAhorro ? '1102 - BI-Ahorro' : '1101 - BI-Monetaria'}: -Q ${monto.toFixed(2)}`
          }
        };
      }
    }

    // 8. Fallback Generic Entry
    const fallbackMonto = extractedAmount !== null ? extractedAmount : 0;
    const isIncome = /\b(cobro|ingreso|abono|deposito|pago de)\b/.test(norm) && !/\b(pague|compre|gasto)\b/.test(norm);
    const fallbackCuenta = isIncome ? '4102 - Colegiaturas Mensuales' : '5203 - Papelería, útiles y suministros';
    const fallbackIva = isIncome ? 0 : +(fallbackMonto - fallbackMonto / 1.12).toFixed(2);
    const fallbackNeto = +(fallbackMonto - fallbackIva).toFixed(2);

    return {
      kind: 'transaction',
      confidence: 'Media (Clasificación Sugerida - Puedes ajustar partida)',
      fecha: todayStr,
      mes: explicitMonth || 'octubre',
      descripcion: text.charAt(0).toUpperCase() + text.slice(1),
      cuenta: fallbackCuenta,
      tipo: isIncome ? 'Ingreso' : 'Gasto',
      montoBruto: fallbackMonto,
      iva: fallbackIva,
      montoNeto: fallbackNeto,
      banco: norm.includes('ahorro') ? 'BI-Ahorro' : 'BI- Monetaria',
      documento: extractedDoc || `DOC-${todayStr.replace(/-/g, '')}`,
      centroCosto: '00-Administración',
      notaAuditoria: isIncome ? 'Ingreso educativo exento de IVA (Art. 8 Ley del IVA).' : 'Gasto operativo con desglose de IVA 12% FEL.',
      partidaDoble: {
        debe: isIncome ? `1101 - BI-Monetaria: +Q ${fallbackMonto.toFixed(2)}` : `${fallbackCuenta}: Q ${fallbackNeto.toFixed(2)}`,
        haber: isIncome ? `${fallbackCuenta}: Q ${fallbackMonto.toFixed(2)}` : `1101 - BI-Monetaria: -Q ${fallbackMonto.toFixed(2)}`
      }
    };
  }

  function buildQueryResponse(rawText, norm, appData) {
    const students = appData.students || [];
    const journal = appData.journal || [];
    const lastEntry = journal[journal.length - 1] || { saldoMonetaria: 0, saldoAhorro: 0 };

    // Student specific query
    const matchedStudent = findStudent(rawText, students);
    if (matchedStudent) {
      const paidMonths = MONTHS.filter(m => (matchedStudent.pagos[m] || 0) > 0);
      const unpaidMonths = MONTHS.slice(0, 10).filter(m => {
        const startIdx = MONTHS.indexOf(normalize(matchedStudent.mesInicio || 'enero'));
        return MONTHS.indexOf(m) >= startIdx && (matchedStudent.pagos[m] || 0) === 0;
      });
      return {
        kind: 'query_answer',
        title: `Estado de Cuenta: ${matchedStudent.estudiante} (${matchedStudent.idFamilia})`,
        answer: `Responsable: ${matchedStudent.responsable}${matchedStudent.contactoSecundario ? ' / ' + matchedStudent.contactoSecundario : ''} (${matchedStudent.email}). Cuota mensual: Q ${matchedStudent.cuotaMensual.toLocaleString('es-GT')}. Total pagado en 2026: Q ${matchedStudent.totalPagado.toLocaleString('es-GT')}.`,
        highlights: [
          `Meses pagados (${paidMonths.length}): ${paidMonths.map(m => m.charAt(0).toUpperCase() + m.slice(1)).join(', ')}`,
          unpaidMonths.length > 0
            ? `⚠️ Meses pendientes a Octubre: ${unpaidMonths.map(m => m.toUpperCase()).join(', ')} (Q ${matchedStudent.pendienteOctubre.toLocaleString('es-GT')})`
            : `✅ Solvente hasta Octubre 2026`
        ]
      };
    }

    // Pending tuition query
    if (/\b(quien|quienes|faltan|pendiente|mora|colegiatura|octubre)\b/.test(norm)) {
      const pendingOct = students.filter(s => s.estado !== 'Retirado' && (s.pagos.octubre || 0) === 0);
      const paidOct = students.filter(s => s.estado !== 'Retirado' && (s.pagos.octubre || 0) > 0);
      return {
        kind: 'query_answer',
        title: 'Control de Solvencia Colegiaturas (Octubre 2026)',
        answer: `En Octubre han pagado ${paidOct.length} alumnos (Q ${paidOct.reduce((a, b) => a + (b.pagos.octubre || 0), 0).toLocaleString('es-GT')}) y faltan ${pendingOct.length} familias por registrar pago:`,
        highlights: [
          ...pendingOct.map(s => `⏳ Pendiente Octubre: ${s.estudiante} (${s.responsable}) — Q ${s.cuotaMensual.toLocaleString('es-GT')}`),
          ...paidOct.map(s => `✅ Pagado Octubre: ${s.estudiante} — Q ${(s.pagos.octubre || 0).toLocaleString('es-GT')}`)
        ]
      };
    }

    // Balances / Savings query
    return {
      kind: 'query_answer',
      title: 'Resumen de Saldos Bancarios al Día',
      answer: `El patrimonio líquido total en cuentas de Banco Industrial es de Q ${(lastEntry.saldoMonetaria + lastEntry.saldoAhorro).toLocaleString('es-GT', { minimumFractionDigits: 2 })}.`,
      highlights: [
        `🏦 BI-Monetaria (1101 Operativa): Q ${lastEntry.saldoMonetaria.toLocaleString('es-GT', { minimumFractionDigits: 2 })}`,
        `🛡️ BI-Ahorro (1102 Reserva IGSS / Bono 14 / Aguinaldo): Q ${lastEntry.saldoAhorro.toLocaleString('es-GT', { minimumFractionDigits: 2 })}`,
        `📅 Próximo corte de Nómina Día 27: Q 27,141.07 sueldos netos + Q 11,551.09 traslado a Ahorro + Q 4,278.00 IGSS desde Ahorro.`
      ]
    };
  }

  return {
    parseNaturalInput,
    MONTHS
  };
})();
