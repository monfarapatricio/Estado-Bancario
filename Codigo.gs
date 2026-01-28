/** Public: servir el HTML si lo necesitás como WebApp (opcional) */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index') // si tu archivo se llama index.html
    .setTitle('Conciliación')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/************ CONFIG ************/
const HOJA_CP      = 'CP';
const HOJA_PAGOS   = 'Imp-CP';
const HOJA_RECURSOS= 'Recursos';
const HOJA_CONTRATOS = 'Contratos';

/************ OBTENER RESUMEN CP ************/
/**
 * Devuelve filas:
 * [cliente, cp, periodoStr, fechaCP, categoria, facturado, debito, abonado]
 */
function obtenerResumenPorCP() {
  const ss = SpreadsheetApp.getActive();
  const shCP = ss.getSheetByName(HOJA_CP);
  const shPagos = ss.getSheetByName(HOJA_PAGOS);
  if (!shCP || !shPagos) throw new Error('Faltan hojas "CP" y/o "Imp-CP".');

  // --- CP ---
  // Esperado: [CP | Cliente | Periodo | Fecha | Concepto | Descripción | Monto]
  const cpVals = getDataRows_(shCP);
  const idxCP = mapHeaders_(shCP, ['CP','Cliente','Periodo','Fecha','Concepto','Descripción','Monto']);

  /** deudas[cp] = {...} */
  const deudas = {};
  const referencias = {};

  function initDeuda_(cp, data) {
    if (deudas[cp]) return deudas[cp];
    const periodoDate = parsePeriodo_(data.periodoRaw) || data.fechaDoc || null;
    deudas[cp] = {
      cliente: data.cliente,
      periodoDate,
      periodoStr: formatPeriodo_(data.periodoRaw, periodoDate),
      fechaDoc: data.fechaDoc,
      facturado: 0,
      debito: 0,
      refactura: 0,
      conceptos: new Set(),
      descripciones: new Set()
    };
    return deudas[cp];
  }

  function extraerCpReferencia_(descripcion) {
    const match = String(descripcion || "").match(/CORRESPONDE AL CP\s*([^\s]+)/i);
    return match ? normalizeCP_(match[1]) : "";
  }

  cpVals.forEach(r => {
    const cp   = normalizeCP_(r[idxCP['CP']]);
    if (!cp) return;

    const cliente    = safeString_(r[idxCP['Cliente']]);
    const periodoRaw = r[idxCP['Periodo']];
    const fechaDoc   = toDate_(r[idxCP['Fecha']]); // Fecha CP
    const concepto   = safeString_(r[idxCP['Concepto']]);
    const desc       = safeString_(r[idxCP['Descripción']]);
    const montoNum   = toNumber_(r[idxCP['Monto']]);
    const conceptoNorm = concepto.trim().toUpperCase();

    const dataBase = { cliente, periodoRaw, fechaDoc };

    if (conceptoNorm === "DEBITO" || conceptoNorm === "REFACTURACION") {
      const cpRef = extraerCpReferencia_(desc);
      if (!cpRef) return;
      if (!referencias[cpRef]) referencias[cpRef] = { debito: 0, refactura: 0 };
      if (conceptoNorm === "DEBITO") referencias[cpRef].debito += Math.abs(montoNum);
      if (conceptoNorm === "REFACTURACION") referencias[cpRef].refactura += Math.abs(montoNum);
      initDeuda_(cpRef, dataBase);
      return;
    }

    const deuda = initDeuda_(cp, dataBase);
    deuda.conceptos.add(conceptoNorm);
    if (desc) deuda.descripciones.add(desc);
    deuda.facturado += Math.abs(montoNum);

    if (!deuda.periodoDate) {
      const p = parsePeriodo_(periodoRaw);
      if (p) {
        deuda.periodoDate = p;
        deuda.periodoStr  = formatPeriodo_(periodoRaw, p);
      } else if (fechaDoc) {
        deuda.periodoDate = fechaDoc;
        deuda.periodoStr  = formatDate_(fechaDoc);
      }
    }
    if (!deuda.fechaDoc && fechaDoc) deuda.fechaDoc = fechaDoc;
  });

  Object.keys(referencias).forEach(cpRef => {
    if (!deudas[cpRef]) return;
    deudas[cpRef].debito += referencias[cpRef].debito || 0;
    deudas[cpRef].refactura += referencias[cpRef].refactura || 0;
  });

  // --- PAGOS ---
  // Esperado: [Marca temporal | CP | Cliente | Nro de E recauda | Fecha de pago | Monto]
  const pagosVals = getDataRows_(shPagos);
  const idxPg = mapHeaders_(shPagos, ['Marca temporal','CP','Cliente','Nro de E recauda','Fecha de pago','Monto']);
  const pagosPorCP = {};
  const pagosAgrupados = agruparPagosPorRecibo_(pagosVals, idxPg);

  // Asignar el total del recibo a los CP indicados, descontando el saldo pendiente de cada CP
  pagosAgrupados.forEach(grupo => {
    let restante = grupo.montoTotal;
    grupo.cps.forEach(cp => {
      if (restante <= 0) return;
      const d = deudas[cp];
      if (!d) return;
      const abonadoActual = pagosPorCP[cp]?.abonado || 0;
    const saldoPend = Math.max(0, d.facturado + d.refactura - d.debito - abonadoActual);
      if (saldoPend <= 0) return;
      const aplicar = Math.min(restante, saldoPend);
      if (!pagosPorCP[cp]) pagosPorCP[cp] = { abonado: 0 };
      pagosPorCP[cp].abonado += aplicar;
      restante -= aplicar;
    });
  });

  function categorizarCP(d) {
    const textos = [
      ...Array.from(d.conceptos || []),
      ...Array.from(d.descripciones || [])
    ].filter(Boolean);
    const hayRefactura = (d.refactura || 0) > 0;
    const hayFuncionamiento = textos.some(t => /FUNCIONAMIENTO/i.test(t));
    if (hayFuncionamiento) return 'GASTO DE FUNCIONAMIENTO';
    if (hayRefactura) return 'REFACTURACIÓN';
    if (d.facturado === 0 && d.debito > 0) return 'Débitos';
    return 'GASTOS HOSPITALARIOS';
  }

  const out = [];
  Object.keys(deudas).forEach(cp => {
    const d = deudas[cp];
    if (!d) return;
    const abonado = pagosPorCP[cp]?.abonado || 0;
    const categoria = categorizarCP(d);
    out.push([
      d.cliente || '',
      cp,
      d.periodoStr || '',
      formatDate_(d.fechaDoc) || '',
      categoria,
      round2_(d.facturado),
      round2_(d.debito),
      round2_(d.refactura),
      round2_(abonado)
    ]);
  });

  return out;
}

/************ INGRESOS (CUT / PAGOS) ************/
function obtenerIngresos() {
  const ss = SpreadsheetApp.getActive();
  const shPagos = ss.getSheetByName(HOJA_PAGOS);
  if (!shPagos) throw new Error('Falta la hoja "Imp-CP".');

  const vals = getDataRows_(shPagos);
  const idx = mapHeaders_(shPagos, ['Marca temporal','CP','Cliente','Nro de E recauda','Fecha de pago','Monto']);

  const out = [];
  const pagosAgrupados = agruparPagosPorRecibo_(vals, idx);
  pagosAgrupados.forEach(grupo => {
    const cliente = safeString_(grupo.cliente);
    const fecha = grupo.fecha;
    const monto = grupo.montoTotal;
    out.push([cliente, formatDate_(fecha), round2_(monto)]);
  });
  return out;
}

/************ RECURSOS PROCESADOS (para front Recursos) ************/
function _mapClientes_() {
  const sh = SpreadsheetApp.getActive().getSheetByName("clientes");
  if (!sh) throw new Error('Falta hoja "clientes"');
  const vals = sh.getDataRange().getValues();
  const header = vals.shift();
  const idxId = header.indexOf("IDCliente");
  const idxCli = header.indexOf("Cliente");
  const map = new Map();
  vals.forEach(r => {
    const id = String(r[idxId]).trim();
    const nom = String(r[idxCli]).trim();
    if (id && nom) map.set(id, nom);
  });
  return map;
}

/**
 * Reglas de identificación:
 * - Sin ID y Detalle empieza con “Comprobante generado automáticamente...” => ID=2 (PAMI)
 * - Sin ID y no cumple => ID=9999 (SIN IDENTIFICAR)
 * - ID que no existe en “clientes” => ID=3 (ORTODONCIA)
 */
function obtenerRecursosProcesados() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_RECURSOS);
  if (!sh) throw new Error('Falta hoja "Recursos"');
  const vals = sh.getDataRange().getValues();
  const header = vals.shift();

  const idxCpte   = header.indexOf("Nro. Cpte");
  const idxFecha  = header.indexOf("Fecha");
  const idxId     = header.indexOf("IDCliente");
  const idxDet    = header.indexOf("Detalle");
  const idxPerc   = header.indexOf("Percibido");

  const cliMap = _mapClientes_();

  const NOMBRE_PAMI = "PAMI";
  const NOMBRE_SIN  = "SIN IDENTIFICAR";
  const NOMBRE_ORTO = "ORTODONCIA";
  const DET_PREFIX = "Comprobante generado automáticamente desde el Proceso de Conciliación Bancaria";

  const out = [];

  vals.forEach(r => {
    const nroCpte = r[idxCpte];
    const fechaRaw = r[idxFecha];
    const idRaw = String(r[idxId] ?? "").trim();
    const det = String(r[idxDet] ?? "").trim();
    const perc = Number(r[idxPerc] || 0);

    // fecha dd/mm/yyyy
    let fechaStr;
    if (fechaRaw instanceof Date) {
      const dd = String(fechaRaw.getDate()).padStart(2,"0");
      const mm = String(fechaRaw.getMonth()+1).padStart(2,"0");
      const yy = fechaRaw.getFullYear();
      fechaStr = `${dd}/${mm}/${yy}`;
    } else {
      fechaStr = String(fechaRaw);
    }

    let id = idRaw;
    let nombre;

    if (!id) {
      if (det.startsWith(DET_PREFIX)) { id = "2"; nombre = NOMBRE_PAMI; }
      else { id = "9999"; nombre = NOMBRE_SIN; }
    } else {
      if (cliMap.has(id)) nombre = cliMap.get(id);
      else { id = "3"; nombre = NOMBRE_ORTO; }
    }

    out.push([ id, nombre, fechaStr, String(nroCpte ?? ""), det, perc ]);
  });

  return out;
}

/************ RENT ROLL (ALQUILERES) – una fila por tramo ************/
/**
 * Hoja "Contratos" (una línea por tramo / renovación):
 * Encabezados aceptados (alias):
 *  - IDCliente:  ["IDCliente","ID Cliente","IdCliente","Id Cliente","ID"]
 *  - Cliente:    ["Cliente","Nombre cliente","Nombre"]
 *  - Inicio:     ["Inicio","Desde","Fecha inicio","Inicio contrato","Inicio de contrato"]
 *  - Fin:        ["Fin","Hasta","Fecha fin","Fin contrato","Fin de contrato"]
 *  - Importe:    ["Importe","Importe mensual","Monto","Alquiler"]
 *  - Obs:        ["Obs","Observaciones","Notas","Detalle"]
 *
 * params: {desde?: "yyyy-mm-dd", hasta?: "yyyy-mm-dd", includeOnlyIds?: string[]}
 *
 * Devuelve filas:
 * [id, cliente, "01/mm/yyyy", cDesde, cHasta, impMens, prorr, esperado, percibido, diferencia, obs]
 */
function calcularRentRollAlquileres(params) {
  const ss = SpreadsheetApp.getActive();
  const shC = ss.getSheetByName(typeof HOJA_CONTRATOS === 'string' ? HOJA_CONTRATOS : 'Contratos');
  const shR = ss.getSheetByName(typeof HOJA_RECURSOS  === 'string' ? HOJA_RECURSOS  : 'Recursos');
  if (!shC) throw new Error('Falta hoja "Contratos".');
  if (!shR) throw new Error('Falta hoja "Recursos".');

  // --- helpers de fechas/mes ---
  const startOfDay = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const mesInicio  = d => new Date(d.getFullYear(), d.getMonth(), 1);
  const mesFin     = d => new Date(d.getFullYear(), d.getMonth()+1, 0);
  const addMonths  = (d, n) => new Date(d.getFullYear(), d.getMonth()+n, 1);

  const hoy = toStartOfDay_(new Date());

  // --- parámetros de recorte opcionales ---
  let pDesde = params?.desde ? toDate_(params.desde) : null;
  let pHasta = params?.hasta ? toDate_(params.hasta) : null;
  if (pHasta && pHasta > hoy) pHasta = hoy; // nunca futuro
  const includeOnlyIds = Array.isArray(params?.includeOnlyIds) ? params.includeOnlyIds.map(String) : null;

  // --- CONTRATOS ---
  const idxC = mapHeaders_(shC, [
    'IDCliente','Cliente','Inicio','Fin','Importe Mensual','Prorrateo','Observación'
  ]);

  const segmentos = [];
  getDataRows_(shC).forEach(r => {
    const id   = safeString_(r[idxC['IDCliente']]);
    const cli  = safeString_(r[idxC['Cliente']]);
    const dIni = toDate_(r[idxC['Inicio']]);
    const dFin = toDate_(r[idxC['Fin']]);  // puede no venir -> abierto
    const imp  = toNumber_(r[idxC['Importe Mensual']]);
    const obs  = safeString_(r[idxC['Observación']]);

    if (!id || !cli || !dIni || !imp) return;
    if (includeOnlyIds && !includeOnlyIds.includes(id)) return;

    // recorte por parámetros
    let desde = dIni;
    let hasta = dFin || new Date(2100,0,1);

    if (pDesde && hasta < pDesde) return; // completamente antes
    if (pHasta && desde > pHasta) return; // completamente después
    if (pDesde && desde < pDesde) desde = pDesde;
    if (pHasta && hasta > pHasta) hasta = pHasta;

    // nunca futuro
    if (hasta > hoy) hasta = hoy;
    if (desde > hasta) return;

    segmentos.push({ id, cliente: cli, desde, hasta, importe: imp, obs });
  });

  if (!segmentos.length) return [];

  // --- RECURSOS (Percibido por id+mes) ---
  const idxR = mapHeaders_(shR, ['Nro. Cpte','Fecha','IDCliente','Detalle','Percibido']);
  const percPorMes = new Map(); // clave: "id|yyyy-mm" -> total
  getDataRows_(shR).forEach(r => {
    const id = String(r[idxR['IDCliente']] ?? '').trim();
    if (!id) return;
    if (includeOnlyIds && !includeOnlyIds.includes(id)) return;

    const f = toDate_(r[idxR['Fecha']]);
    if (!f) return;
    const f0 = startOfDay(f);
    if (f0 > hoy) return; // no futuro

    const ym = `${f0.getFullYear()}-${String(f0.getMonth()+1).padStart(2,'0')}`;
    const k  = `${id}|${ym}`;
    percPorMes.set(k, (percPorMes.get(k) || 0) + toNumber_(r[idxR['Percibido']]));
  });

  // --- EXPECTADO por mes (SIN prorrateo: siempre el importe completo del mes) ---
  const expMap = new Map(); // "id|yyyy-mm" -> { cliente, cDesde, cHasta, impMens, esperado, obs[] }

  segmentos.forEach(s => {
    let cur = mesInicio(s.desde);
    const limite = mesInicio(s.hasta); // último mes incluido

    while (cur <= limite) {
      const ym = `${cur.getFullYear()}-${String(cur.getMonth()+1).padStart(2,'0')}`;
      const k  = `${s.id}|${ym}`;
      if (!expMap.has(k)) {
        expMap.set(k, {
          cliente: s.cliente,
          cDesde: formatDate_(s.desde),
          cHasta: formatDate_(s.hasta),
          impMens: 0,
          esperado: 0,
          obs: []
        });
      }
      const it = expMap.get(k);
      it.impMens  = round2_(it.impMens + s.importe); // si tiene 2 contratos superpuestos en ese mes, suma
      it.esperado = round2_(it.esperado + s.importe);
      if (s.obs) it.obs.push(s.obs);

      cur = addMonths(cur, 1);
    }
  });

  // --- SALIDA ---
  const out = [];
  Array.from(expMap.entries())
    .sort((a,b) => {
      const [idA, ymA] = a[0].split('|');
      const [idB, ymB] = b[0].split('|');
      if (idA !== idB) return idA.localeCompare(idB);
      return ymA.localeCompare(ymB);
    })
    .forEach(([key, it]) => {
      const [id, ym] = key.split('|');
      const [yy, mm] = ym.split('-').map(Number);
      const perc = percPorMes.get(key) || 0;
      const dif  = round2_(Math.max(0, it.esperado - perc)); // saldo pendiente (no negativo)
      const mesStr = `01/${String(mm).padStart(2,'0')}/${yy}`; // para el front

      out.push([
        id,
        it.cliente,
        mesStr,
        it.cDesde,
        it.cHasta,
        round2_(it.impMens),
        false,                     // prorr (siempre false, no se usa prorrateo)
        round2_(it.esperado),      // esperado del mes (importe mensual completo)
        round2_(perc),             // percibido del mes
        dif,                       // saldo pendiente (>= 0)
        it.obs.join(' | ')
      ]);
    });

  return out;
}

/************ TEST RÁPIDO ************/
function test_calcularRentRollAlquileres(){
  const rows = calcularRentRollAlquileres({
    // opcional:
    // desde: '2024-01-01',
    // hasta: '2025-12-31',
    // includeOnlyIds: ['378477','168740']
  });
  Logger.log(rows.length + ' filas');
  Logger.log(JSON.stringify(rows.slice(0,10)));
}

/************ HELPERS ************/
function getDataRows_(sh) {
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];
  return vals.slice(1).filter(r => r.some(c => c !== '' && c !== null));
}

function agruparPagosPorRecibo_(pagosVals, idxPg) {
  const recibosMap = new Map();

  pagosVals.forEach(r => {
    const cp = normalizeCP_(r[idxPg['CP']]);
    const nroRec = safeString_(r[idxPg['Nro de E recauda']]);
    const cliente = safeString_(r[idxPg['Cliente']]);
    const fecha = toDate_(r[idxPg['Fecha de pago']]);
    const monto = toNumber_(r[idxPg['Monto']]);

    if (!nroRec) {
      if (!cp) return;
      const key = `CP:${cp}`;
      if (!recibosMap.has(key)) {
        recibosMap.set(key, { montoTotal: 0, cps: [cp], cliente, fecha });
      }
      const it = recibosMap.get(key);
      it.montoTotal += monto;
      if (!it.fecha && fecha) it.fecha = fecha;
      if (!it.cliente && cliente) it.cliente = cliente;
      return;
    }

    if (!recibosMap.has(nroRec)) {
      recibosMap.set(nroRec, { montoTotal: 0, cps: [], cliente, fecha });
    }
    const it = recibosMap.get(nroRec);
    it.montoTotal = Math.max(it.montoTotal, monto);
    if (cp && !it.cps.includes(cp)) it.cps.push(cp);
    if (!it.fecha && fecha) it.fecha = fecha;
    if (!it.cliente && cliente) it.cliente = cliente;
  });

  return Array.from(recibosMap.values());
}

function mapHeaders_(sh, expectedHeaders) {
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => (h+'').trim());
  const map = {};
  expectedHeaders.forEach(h => {
    const idx = headers.findIndex(x => normalizeLabel_(x) === normalizeLabel_(h));
    if (idx === -1) throw new Error(`No se encontró la columna "${h}" en la hoja "${sh.getName()}".`);
    map[h] = idx;
  });
  return map;
}

/**
 * Mapea encabezados con alias (para Contratos/Recursos).
 * spec: { clave: {aliases:[...], required:true|false} }
 */
function mapHeadersFlexible_(sh, spec) {
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => (h+'').trim());
  const norm = (s) => normalizeLabel_(s);
  const result = {};
  Object.keys(spec).forEach(key => {
    const { aliases, required } = spec[key];
    let idx = -1;
    for (const a of aliases) {
      const i = headers.findIndex(h => norm(h) === norm(a));
      if (i !== -1) { idx = i; break; }
    }
    if (required && idx === -1) {
      throw new Error(`No se encontró la columna requerida "${aliases[0]}" (alias: ${aliases.join(', ')}) en la hoja "${sh.getName()}".`);
    }
    result[key] = idx;
  });
  return result;
}

function normalizeLabel_(s) {
  return (s || '').toString().trim().toLowerCase()
    .replace(/\s+/g,' ')
    .replace(/[áä]/g,'a').replace(/[éë]/g,'e')
    .replace(/[íï]/g,'i').replace(/[óö]/g,'o')
    .replace(/[úü]/g,'u')
    .replace(/[^\w ]/g,'');
}
function normalizeCP_(v) {
  return (v === null || v === undefined) ? '' : (v + '').trim();
}
function safeString_(v) { return (v === null || v === undefined) ? '' : (v + '').trim(); }

function toNumber_(v) {
  const n = typeof v === 'number' ? v : parseFloat((v+'').replace(/\./g,'').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function toDate_(v) {
  if (v instanceof Date && !isNaN(v)) return toStartOfDay_(v);
  if (!v) return null;
  const s = (v+'').trim();

  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); // YYYY-MM-DD
  if (m) return toStartOfDay_(new Date(+m[1], +m[2]-1, +m[3]));

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);   // DD/MM/YYYY
  if (m) return toStartOfDay_(new Date(+m[3], +m[2]-1, +m[1]));

  m = s.match(/^(\d{1,2})\/(\d{4})$/);              // MM/YYYY
  if (m) return toStartOfDay_(new Date(+m[2], +m[1]-1, 1));

  m = s.match(/^(\d{4})-(\d{1,2})$/);               // YYYY-MM
  if (m) return toStartOfDay_(new Date(+m[1], +m[2]-1, 1));

  const d = new Date(s);
  return isNaN(d) ? null : toStartOfDay_(d);
}
function parsePeriodo_(v){ return toDate_(v); }
function toStartOfDay_(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

function diffDays_(d1, d2) {
  if (!d1 || !d2) return 0;
  const ms = toStartOfDay_(d2) - toStartOfDay_(d1);
  return Math.floor(ms / (1000*60*60*24));
}
function evaluarVencimiento_(periodoDate, hoy, diasUmbral) {
  if (!periodoDate) return { vencido:false, diasExcedidos:0 };
  const dias = diffDays_(periodoDate, hoy);
  const excedido = Math.max(0, dias - diasUmbral);
  return { vencido: dias > diasUmbral, diasExcedidos: excedido };
}

function formatDate_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}
function formatDateTime_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}
function formatPeriodo_(raw, parsedDate) {
  if (!raw && parsedDate) return formatDate_(parsedDate);
  const s = (raw || '').toString().trim();

  let m = s.match(/^(\d{1,2})\/(\d{4})$/); // MM/YYYY
  if (m) return '01/' + ('0'+m[1]).slice(-2) + '/' + m[2];

  m = s.match(/^(\d{4})-(\d{1,2})$/);      // YYYY-MM
  if (m) return '01/' + ('0'+m[2]).slice(-2) + '/' + m[1];

  const asDate = toDate_(s);
  return asDate ? formatDate_(asDate) : s;
}
function isDebito_(concepto) {
  const s = normalizeLabel_(concepto || '');
  const patrones = [
    /\bdebito(s)?\b/,
    /\bnota(s)? de debito\b/,
    /\bnd\b/,
    /\bdb\b/,
    /\breintegro(s)?\b/,
    /\bajuste(s)? (de )?debito\b/,
    /\bdebito medico\b/,
    /\bdevolucion(es)?\b/
  ];
  return patrones.some(re => re.test(s));
}
function round2_(x){ return Math.round((x + Number.EPSILON) * 100) / 100; }
