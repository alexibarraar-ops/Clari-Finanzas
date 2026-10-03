const CLARI = {
  SPREADSHEET_ID: '1uumwM1R2HXuvItxkj-MZ_XXxRg0qZSMAPEWELVLiNPE',
  SHEETS: {
    MOVIMIENTOS: 'MOVIMIENTOS',
    CATEGORIAS: 'CATEGORIAS',
    CUENTAS: 'CUENTAS',
    INSTRUMENTOS: 'INSTRUMENTOS',
    BENEFICIARIOS: 'BENEFICIARIOS',
    CONFIG: 'CONFIG',
    FRASES: 'FRASES'
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Clari Finanzas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getBootstrapData() {
  return {
    appUrl: ScriptApp.getService().getUrl(),
    config: getConfig_(),
    categorias: readObjects_(CLARI.SHEETS.CATEGORIAS).filter(r => isTrue_(r.ACTIVA)),
    cuentas: readObjects_(CLARI.SHEETS.CUENTAS).filter(r => isTrue_(r.ACTIVA)),
    instrumentos: readObjects_(CLARI.SHEETS.INSTRUMENTOS).filter(r => isTrue_(r.ACTIVO)),
    beneficiarios: readObjects_(CLARI.SHEETS.BENEFICIARIOS).filter(r => isTrue_(r.ACTIVO)),
    frase: getFraseDelDia_(),
    resumen: getResumenMes_(),
    ultimos: getUltimosMovimientos_(6)
  };
}

function guardarMovimiento(payload) {
  validarMovimiento_(payload);
  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires';
  const fecha = payload.fecha ? new Date(payload.fecha + 'T12:00:00') : now;
  const id = 'MOV-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  sh.appendRow([
    id,
    Utilities.formatDate(fecha, tz, 'dd/MM/yyyy'),
    Utilities.formatDate(now, tz, 'HH:mm:ss'),
    String(payload.tipo || '').toUpperCase(),
    Number(payload.monto),
    String(payload.concepto || '').trim(),
    String(payload.categoria || '').trim(),
    String(payload.cuenta || '').trim(),
    String(payload.medioPago || '').trim(),
    String(payload.instrumento || '').trim(),
    String(payload.tipoGasto || '').trim().toUpperCase(),
    String(payload.beneficiario || '').trim(),
    String(payload.nota || '').trim(),
    String(payload.origen || 'MANUAL').toUpperCase(),
    String(payload.textoOriginal || '').trim(),
    now,
    now
  ]);

  return { ok: true, id, resumen: getResumenMes_(), ultimos: getUltimosMovimientos_(6) };
}

function getConfig_() {
  const out = {};
  readObjects_(CLARI.SHEETS.CONFIG).forEach(r => { if (r.CLAVE) out[String(r.CLAVE)] = r.VALOR; });
  return out;
}

function getFraseDelDia_() {
  const frases = readObjects_(CLARI.SHEETS.FRASES).filter(r => isTrue_(r.ACTIVA) && r.FRASE);
  if (!frases.length) return 'Hoy también cuenta.';
  const dayKey = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires', 'yyyyMMdd'));
  return frases[dayKey % frases.length].FRASE;
}

function getUltimosMovimientos_(limite) {
  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];
  const count = Math.min(Number(limite || 6), lastRow - 1);
  const start = Math.max(2, lastRow - count + 1);
  return sh.getRange(start, 1, count, 17).getDisplayValues().reverse().map(r => ({
    id:r[0], fecha:r[1], hora:r[2], tipo:r[3],
    monto:Number(String(r[4]).replace(/\./g,'').replace(',','.'))||0,
    concepto:r[5], categoria:r[6], cuenta:r[7], medioPago:r[8],
    instrumento:r[9], tipoGasto:r[10], beneficiario:r[11], nota:r[12], origen:r[13]
  }));
}

function getResumenMes_() {
  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const lastRow = sh.getLastRow();
  let ingresos = 0, gastos = 0;
  if (lastRow > 1) {
    const now = new Date(), month = now.getMonth(), year = now.getFullYear();
    sh.getRange(2,1,lastRow-1,17).getValues().forEach(r => {
      const fecha = parseFecha_(r[1]);
      if (!fecha || fecha.getMonth() !== month || fecha.getFullYear() !== year) return;
      const tipo = String(r[3] || '').toUpperCase();
      const monto = Number(r[4] || 0);
      if (tipo === 'INGRESO') ingresos += monto;
      if (tipo === 'GASTO') gastos += monto;
    });
  }
  return { ingresos, gastos, disponible: ingresos - gastos };
}

function validarMovimiento_(p) {
  if (!p) throw new Error('No se recibió el movimiento.');
  if (!['GASTO','INGRESO'].includes(String(p.tipo || '').toUpperCase())) throw new Error('Tipo de movimiento inválido.');
  if (!p.monto || Number(p.monto) <= 0) throw new Error('Ingresá un monto válido.');
  if (!String(p.concepto || '').trim()) throw new Error('Ingresá un concepto.');
  if (!String(p.categoria || '').trim()) throw new Error('Elegí una categoría.');
  if (!String(p.cuenta || '').trim()) throw new Error('Elegí una cuenta o medio de pago.');
}

function readObjects_(sheetName) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1).filter(row => row.some(v => String(v).trim() !== '')).map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
}

function getSheet_(name) {
  const sh = SpreadsheetApp.openById(CLARI.SPREADSHEET_ID).getSheetByName(name);
  if (!sh) throw new Error('No existe la hoja: ' + name);
  return sh;
}

function isTrue_(v) { return String(v).toUpperCase() === 'TRUE'; }

function parseFecha_(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  const s = String(value || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]), 12,0,0);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
