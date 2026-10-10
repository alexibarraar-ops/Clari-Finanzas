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
    appUrl: 'https://alexibarraar-ops.github.io/Clari-Finanzas/',
    config: getConfig_(),
    categorias: readObjects_(CLARI.SHEETS.CATEGORIAS).filter(r => isTrue_(r.ACTIVA)),
    cuentas: readObjects_(CLARI.SHEETS.CUENTAS).filter(r => isTrue_(r.ACTIVA)),
    instrumentos: readObjects_(CLARI.SHEETS.INSTRUMENTOS).filter(r => isTrue_(r.ACTIVO)),
    beneficiarios: readObjects_(CLARI.SHEETS.BENEFICIARIOS).filter(r => isTrue_(r.ACTIVO)),
    frase: getFraseDelDia_(),
    resumen: getResumenMes_(),
    resumenDetallado: getResumenDetallado_(),
    saldos: getSaldosCuentas(),
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



function validarSaldosConfigurados_() {
  const s = getSaldosCuentas();
  if (!s.completas) throw new Error('Primero configurá el saldo actual de todas las cuentas.');
}

function guardarTransferencia(payload) {
  validarSaldosConfigurados_();
  validarTransferencia_(payload);

  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires';
  const fecha = payload.fecha ? new Date(payload.fecha + 'T12:00:00') : now;
  const editId = String(payload.id || '').trim();

  if (editId) {
    const row = buscarFilaMovimiento_(editId);
    if (!row) throw new Error('No encontré la transferencia a editar.');

    const tipoActual = String(sh.getRange(row, 4).getDisplayValue() || '').toUpperCase();
    if (tipoActual !== 'TRANSFERENCIA') throw new Error('El movimiento no es una transferencia.');

    const creadoEn = sh.getRange(row, 16).getValue() || now;
    const transferenciaId = sh.getRange(row, 19).getDisplayValue() || ('TRF-' + Utilities.getUuid().slice(0, 8).toUpperCase());

    sh.getRange(row, 1, 1, 19).setValues([[
      editId,
      Utilities.formatDate(fecha, tz, 'dd/MM/yyyy'),
      Utilities.formatDate(now, tz, 'HH:mm:ss'),
      'TRANSFERENCIA',
      Number(payload.monto),
      'Transferencia a ' + String(payload.cuentaDestino || '').trim(),
      'Transferencia entre cuentas',
      String(payload.cuentaOrigen || '').trim(),
      'TRANSFERENCIA',
      '',
      '',
      'Clari',
      String(payload.nota || '').trim(),
      String(payload.origen || 'MANUAL').toUpperCase(),
      String(payload.textoOriginal || '').trim(),
      creadoEn,
      now,
      String(payload.cuentaDestino || '').trim(),
      transferenciaId
    ]]);

    return { ok:true, id:editId, transferenciaId, resumen:getResumenMes_(), ultimos:getUltimosMovimientos_(6) };
  }

  const id = 'MOV-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  const transferenciaId = 'TRF-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  sh.appendRow([
    id,
    Utilities.formatDate(fecha, tz, 'dd/MM/yyyy'),
    Utilities.formatDate(now, tz, 'HH:mm:ss'),
    'TRANSFERENCIA',
    Number(payload.monto),
    'Transferencia a ' + String(payload.cuentaDestino || '').trim(),
    'Transferencia entre cuentas',
    String(payload.cuentaOrigen || '').trim(),
    'TRANSFERENCIA',
    '',
    '',
    'Clari',
    String(payload.nota || '').trim(),
    String(payload.origen || 'MANUAL').toUpperCase(),
    String(payload.textoOriginal || '').trim(),
    now,
    now,
    String(payload.cuentaDestino || '').trim(),
    transferenciaId
  ]);

  return { ok:true, id, transferenciaId, resumen:getResumenMes_(), ultimos:getUltimosMovimientos_(6) };
}


function guardarAjusteConciliacion(payload) {
  if (!payload) throw new Error('No se recibió el ajuste.');

  const monto = Number(payload.monto || 0);
  const cuenta = String(payload.cuenta || '').trim();
  const direccion = String(payload.direccion || '').toUpperCase();
  const nota = String(payload.nota || '').trim();

  if (!monto || monto <= 0) throw new Error('Ingresá un monto válido.');
  if (!cuenta) throw new Error('Elegí la cuenta.');
  if (!['SUMA','RESTA'].includes(direccion)) throw new Error('Dirección de ajuste inválida.');

  const cuentas = readObjects_(CLARI.SHEETS.CUENTAS)
    .filter(r => isTrue_(r.ACTIVA))
    .map(r => String(r.NOMBRE || '').trim());

  if (!cuentas.includes(cuenta)) throw new Error('La cuenta ya no está activa.');

  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires';
  const id = 'MOV-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  sh.appendRow([
    id,
    Utilities.formatDate(now, tz, 'dd/MM/yyyy'),
    Utilities.formatDate(now, tz, 'HH:mm:ss'),
    'AJUSTE',
    monto,
    direccion === 'SUMA' ? 'Ajuste de saldo positivo' : 'Ajuste de saldo negativo',
    'Ajuste de conciliación',
    cuenta,
    'AJUSTE',
    '',
    '',
    'Clari',
    nota || 'Ajuste por diferencia de conciliación',
    'MANUAL',
    '',
    now,
    now,
    '',
    ''
  ]);

  return {
    ok: true,
    id,
    resumen: getResumenMes_(),
    saldos: getSaldosCuentas(),
    ultimos: getUltimosMovimientos_(6)
  };
}

function validarTransferencia_(p) {
  if (!p) throw new Error('No se recibió la transferencia.');
  if (!p.monto || Number(p.monto) <= 0) throw new Error('Ingresá un monto válido.');

  const origen = String(p.cuentaOrigen || '').trim();
  const destino = String(p.cuentaDestino || '').trim();

  if (!origen) throw new Error('Elegí la cuenta de origen.');
  if (!destino) throw new Error('Elegí la cuenta de destino.');
  if (origen === destino) throw new Error('La cuenta de origen y destino deben ser distintas.');

  const cuentas = readObjects_(CLARI.SHEETS.CUENTAS)
    .filter(r => isTrue_(r.ACTIVA))
    .map(r => String(r.NOMBRE || '').trim());

  if (!cuentas.includes(origen) || !cuentas.includes(destino)) {
    throw new Error('Una de las cuentas ya no está activa.');
  }
}

function actualizarMovimiento(payload) {
  validarMovimiento_(payload);
  if (!payload.id) throw new Error('Falta el ID del movimiento.');

  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const row = buscarFilaMovimiento_(payload.id);
  if (!row) throw new Error('No encontré el movimiento a editar.');

  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires';
  const fecha = payload.fecha ? new Date(payload.fecha + 'T12:00:00') : now;

  const creadoEn = sh.getRange(row, 16).getValue();
  const origenActual = sh.getRange(row, 14).getDisplayValue();
  const textoOriginalActual = sh.getRange(row, 15).getDisplayValue();

  sh.getRange(row, 1, 1, 17).setValues([[
    String(payload.id),
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
    origenActual || String(payload.origen || 'MANUAL').toUpperCase(),
    textoOriginalActual || String(payload.textoOriginal || '').trim(),
    creadoEn || now,
    now
  ]]);

  return { ok: true, id: payload.id, resumen: getResumenMes_(), ultimos: getUltimosMovimientos_(6) };
}

function eliminarMovimiento(id) {
  if (!id) throw new Error('Falta el ID del movimiento.');

  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const row = buscarFilaMovimiento_(id);
  if (!row) throw new Error('No encontré el movimiento a eliminar.');

  sh.deleteRow(row);

  return { ok: true, id, resumen: getResumenMes_(), ultimos: getUltimosMovimientos_(6) };
}

function buscarFilaMovimiento_(id) {
  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return 0;

  const ids = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const target = String(id).trim();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === target) return i + 2;
  }
  return 0;
}


function getCategoriasAdmin() {
  return readObjects_(CLARI.SHEETS.CATEGORIAS)
    .filter(r => isTrue_(r.ACTIVA))
    .map(r => ({
      ID_CATEGORIA: String(r.ID_CATEGORIA || ''),
      NOMBRE: String(r.NOMBRE || ''),
      TIPO: String(r.TIPO || '').toUpperCase(),
      FIJO_VARIABLE_DEFAULT: String(r.FIJO_VARIABLE_DEFAULT || '').toUpperCase(),
      ACTIVA: String(r.ACTIVA || '')
    }))
    .sort((a,b) => {
      if (a.TIPO !== b.TIPO) return a.TIPO.localeCompare(b.TIPO);
      return a.NOMBRE.localeCompare(b.NOMBRE, 'es');
    });
}

function guardarCategoria(payload) {
  if (!payload) throw new Error('No se recibió la categoría.');

  const nombre = String(payload.nombre || '').trim();
  const tipo = String(payload.tipo || '').toUpperCase();
  const fijoVariable = tipo === 'GASTO'
    ? String(payload.fijoVariable || 'VARIABLE').toUpperCase()
    : '';

  if (!nombre) throw new Error('Ingresá un nombre para la categoría.');
  if (!['GASTO','INGRESO'].includes(tipo)) throw new Error('Tipo de categoría inválido.');
  if (tipo === 'GASTO' && !['FIJO','VARIABLE'].includes(fijoVariable)) {
    throw new Error('Elegí FIJO o VARIABLE.');
  }

  const sh = getSheet_(CLARI.SHEETS.CATEGORIAS);
  const lastRow = sh.getLastRow();
  const rows = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 5).getDisplayValues() : [];
  const id = String(payload.id || '').trim();
  const nombreNorm = normalizarTexto_(nombre);

  for (let i = 0; i < rows.length; i++) {
    const rowId = String(rows[i][0] || '').trim();
    const rowNombre = normalizarTexto_(rows[i][1]);
    const rowTipo = String(rows[i][2] || '').toUpperCase();
    const activa = isTrue_(rows[i][4]);

    if (activa && rowId !== id && rowTipo === tipo && rowNombre === nombreNorm) {
      throw new Error('Ya existe una categoría con ese nombre.');
    }
  }

  if (id) {
    const row = buscarFilaCategoria_(id);
    if (!row) throw new Error('No encontré la categoría a editar.');
    sh.getRange(row, 2, 1, 4).setValues([[
      nombre,
      tipo,
      fijoVariable,
      true
    ]]);
  } else {
    sh.appendRow([
      siguienteCategoriaId_(),
      nombre,
      tipo,
      fijoVariable,
      true
    ]);
  }

  return {
    ok: true,
    categorias: getCategoriasAdmin()
  };
}

function eliminarCategoria(id) {
  if (!id) throw new Error('Falta el ID de la categoría.');

  const row = buscarFilaCategoria_(id);
  if (!row) throw new Error('No encontré la categoría.');

  const sh = getSheet_(CLARI.SHEETS.CATEGORIAS);
  sh.getRange(row, 5).setValue(false);

  return {
    ok: true,
    categorias: getCategoriasAdmin()
  };
}

function buscarFilaCategoria_(id) {
  const sh = getSheet_(CLARI.SHEETS.CATEGORIAS);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return 0;

  const ids = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const target = String(id || '').trim();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === target) return i + 2;
  }
  return 0;
}

function siguienteCategoriaId_() {
  const sh = getSheet_(CLARI.SHEETS.CATEGORIAS);
  const lastRow = sh.getLastRow();
  let max = 0;

  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues().forEach(r => {
      const m = String(r[0] || '').match(/^CAT-(\d+)$/i);
      if (m) max = Math.max(max, Number(m[1]));
    });
  }

  return 'CAT-' + String(max + 1).padStart(3, '0');
}

function normalizarTexto_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}


function getSaldosCuentas() {
  const shCuentas = getSheet_(CLARI.SHEETS.CUENTAS);
  const lastCuentaRow = shCuentas.getLastRow();
  if (lastCuentaRow <= 1) return { cuentas: [], total: 0, completas: false };

  const raw = shCuentas.getRange(2, 1, lastCuentaRow - 1, 6).getValues();
  const cuentas = raw
    .filter(r => isTrue_(r[3]))
    .map(r => ({
      ID_CUENTA: String(r[0] || '').trim(),
      NOMBRE: String(r[1] || '').trim(),
      TIPO: String(r[2] || '').trim(),
      SALDO_BASE: Number(r[4] || 0),
      CORTE_EN: r[5]
    }));

  const shMov = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const lastRow = shMov.getLastRow();
  const movs = lastRow > 1 ? shMov.getRange(2, 1, lastRow - 1, 19).getValues() : [];

  const out = cuentas.map(c => {
    const nombre = c.NOMBRE;
    const saldoBase = Number(c.SALDO_BASE || 0);
    const corte = c.CORTE_EN instanceof Date && !isNaN(c.CORTE_EN) ? c.CORTE_EN : parseDateTime_(c.CORTE_EN);
    let delta = 0;

    if (corte) {
      movs.forEach(r => {
        const creadoEn = r[15] instanceof Date ? r[15] : parseDateTime_(r[15]);
        if (!creadoEn || creadoEn <= corte) return;

        const tipo = String(r[3] || '').toUpperCase();
        const monto = Number(r[4] || 0);
        const cuentaOrigen = String(r[7] || '').trim();
        const cuentaDestino = String(r[17] || '').trim();

        if (tipo === 'INGRESO' && cuentaOrigen === nombre) delta += monto;
        if (tipo === 'GASTO' && cuentaOrigen === nombre) delta -= monto;
        if (tipo === 'AJUSTE' && cuentaOrigen === nombre) {
          const concepto = String(r[5] || '').toLowerCase();
          if (concepto.includes('positivo')) delta += monto;
          else delta -= monto;
        }

        if (tipo === 'TRANSFERENCIA') {
          if (cuentaOrigen === nombre) delta -= monto;
          if (cuentaDestino === nombre) delta += monto;
        }
      });
    }

    return {
      id: c.ID_CUENTA,
      nombre,
      tipo: c.TIPO,
      configurada: !!corte,
      saldoBase,
      saldoActual: saldoBase + delta,
      corteEn: corte ? Utilities.formatDate(corte, Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm:ss') : ''
    };
  });

  return {
    cuentas: out,
    total: out.reduce((s,x) => s + Number(x.saldoActual || 0), 0),
    completas: out.length > 0 && out.every(x => x.configurada)
  };
}

function guardarSaldosBase(payload) {
  if (!payload || !Array.isArray(payload.cuentas)) throw new Error('No se recibieron los saldos.');

  const sh = getSheet_(CLARI.SHEETS.CUENTAS);
  const lastRow = sh.getLastRow();
  const activeRows = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 6).getValues() : [];
  const activeIds = activeRows
    .filter(r => isTrue_(r[3]))
    .map(r => String(r[0] || '').trim());

  const incoming = {};
  payload.cuentas.forEach(x => {
    const id = String(x.id || '').trim();
    if (!id) return;
    if (x.saldo === '' || x.saldo === null || typeof x.saldo === 'undefined') {
      throw new Error('Completá todos los saldos. Si una cuenta está vacía, ingresá 0.');
    }
    const saldo = Number(x.saldo);
    if (!isFinite(saldo)) throw new Error('Hay un saldo inválido.');
    incoming[id] = saldo;
  });

  activeIds.forEach(id => {
    if (!Object.prototype.hasOwnProperty.call(incoming, id)) {
      throw new Error('Completá el saldo de todas las cuentas activas.');
    }
  });

  const rowById = {};
  activeRows.forEach((r,i) => rowById[String(r[0] || '').trim()] = i + 2);
  const now = new Date();

  activeIds.forEach(id => {
    const row = rowById[id];
    sh.getRange(row, 5, 1, 2).setValues([[incoming[id], now]]);
  });

  return getSaldosCuentas();
}

function parseDateTime_(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  const s = String(value || '').trim();
  if (!s) return null;

  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]), Number(m[4]), Number(m[5]), Number(m[6] || 0));

  const d = new Date(s);
  return isNaN(d) ? null : d;
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


function getTodosMovimientos(limite) {
  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];

  const max = Math.min(Number(limite || 500), lastRow - 1);
  const start = Math.max(2, lastRow - max + 1);

  return sh.getRange(start, 1, max, 19).getDisplayValues().reverse().map(r => ({
    id:r[0], fecha:r[1], hora:r[2], tipo:r[3],
    monto:Number(String(r[4]).replace(/\./g,'').replace(',','.'))||0,
    concepto:r[5], categoria:r[6], cuenta:r[7], medioPago:r[8],
    instrumento:r[9], tipoGasto:r[10], beneficiario:r[11], nota:r[12], origen:r[13],
    cuentaDestino:r[17], transferenciaId:r[18]
  }));
}

function getUltimosMovimientos_(limite) {
  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];
  const count = Math.min(Number(limite || 6), lastRow - 1);
  const start = Math.max(2, lastRow - count + 1);
  return sh.getRange(start, 1, count, 19).getDisplayValues().reverse().map(r => ({
    id:r[0], fecha:r[1], hora:r[2], tipo:r[3],
    monto:Number(String(r[4]).replace(/\./g,'').replace(',','.'))||0,
    concepto:r[5], categoria:r[6], cuenta:r[7], medioPago:r[8],
    instrumento:r[9], tipoGasto:r[10], beneficiario:r[11], nota:r[12], origen:r[13],
    cuentaDestino:r[17], transferenciaId:r[18]
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


function getResumenDetallado() {
  return getResumenDetallado_();
}

function getResumenDetallado_() {
  const sh = getSheet_(CLARI.SHEETS.MOVIMIENTOS);
  const lastRow = sh.getLastRow();

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  const prev = new Date(currentYear, currentMonth - 1, 1);
  const prevMonth = prev.getMonth();
  const prevYear = prev.getFullYear();

  const categorias = {};
  const beneficiarios = {};
  const tipoGasto = { FIJO: 0, VARIABLE: 0 };

  let ingresos = 0;
  let gastos = 0;
  let gastosMesAnterior = 0;

  if (lastRow > 1) {
    const values = sh.getRange(2, 1, lastRow - 1, 17).getValues();

    values.forEach(r => {
      const fecha = parseFecha_(r[1]);
      if (!fecha) return;

      const tipo = String(r[3] || '').toUpperCase();
      const monto = Number(r[4] || 0);
      const categoria = String(r[6] || 'Otros');
      const tg = String(r[10] || '').toUpperCase();
      const beneficiario = String(r[11] || 'Clari');

      const isCurrent = fecha.getMonth() === currentMonth && fecha.getFullYear() === currentYear;
      const isPrev = fecha.getMonth() === prevMonth && fecha.getFullYear() === prevYear;

      if (isPrev && tipo === 'GASTO') {
        gastosMesAnterior += monto;
      }

      if (!isCurrent) return;

      if (tipo === 'INGRESO') {
        ingresos += monto;
        return;
      }

      if (tipo === 'GASTO') {
        gastos += monto;
        categorias[categoria] = (categorias[categoria] || 0) + monto;
        beneficiarios[beneficiario] = (beneficiarios[beneficiario] || 0) + monto;

        if (tg === 'FIJO' || tg === 'VARIABLE') {
          tipoGasto[tg] += monto;
        }
      }
    });
  }

  const categoriasOrdenadas = Object.keys(categorias)
    .map(nombre => ({ nombre, monto: categorias[nombre] }))
    .sort((a,b) => b.monto - a.monto);

  const beneficiariosOrdenados = Object.keys(beneficiarios)
    .map(nombre => ({ nombre, monto: beneficiarios[nombre] }))
    .sort((a,b) => b.monto - a.monto);

  let variacion = null;
  if (gastosMesAnterior > 0) {
    variacion = ((gastos - gastosMesAnterior) / gastosMesAnterior) * 100;
  }

  return {
    ingresos,
    gastos,
    disponible: ingresos - gastos,
    categorias: categoriasOrdenadas,
    beneficiarios: beneficiariosOrdenados,
    tipoGasto,
    gastosMesAnterior,
    variacionVsAnterior: variacion
  };
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
