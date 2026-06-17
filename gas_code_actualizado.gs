// ============================================================
// PALMA GRANDE S.A.S — Supervisión Calidad Cosecha
// Google Apps Script — Backend
// ============================================================

function autorizarScript() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Script OK. Sheet: ' + ss.getName());
  return { ok: true, msg: 'Script autorizado correctamente' };
}

const SHEET_NAME = 'Evaluaciones';

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  if (!e) e = {};
  const params = (e && e.parameter) ? e.parameter : {};

  let body = {};
  try {
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
  } catch(err) {}

  const action = body.action || params.action || '';

  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  let result;
  try {
    if (action === 'guardar') {
      result = guardarEvaluacion(body);
    } else if (action === 'listar') {
      result = listarEvaluaciones(params);
    } else if (action === 'guardarRuta') {
      result = guardarRuta(body);
    } else if (action === 'getRuta') {
      result = getRuta(params);
    } else if (action === 'resumen') {
      result = getResumen(params);
    } else if (action === 'generarPDF') {
      const p = body.params || {};
      result = generarPDFHistorial(p);
    } else if (action === 'generarPDFReporte') {
      const p = body.params || {};
      result = generarPDFReporte(p);
    } else if (action === 'generarPDFGerencia') {
      const p = body.params || {};
      result = generarPDFGerencia(p);
    } else if (action === 'getTendencias') {
      result = getTendencias(params);
    } else if (action === 'crearAnalisis') {
      result = crearHojasAnalisis();
    } else if (action === 'listarPersonal') {
      result = listarPersonal();
    } else if (action === 'togglePersonal') {
      result = togglePersonal(body.cod, body.activo);
    } else if (action === '') {
      result = { ok: true, status: 'GAS activo — Palma Grande Cosecha' };
    } else {
      result = { ok: false, error: 'Acción no reconocida: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.toString() };
  }

  output.setContent(JSON.stringify(result));
  return output;
}

// ── GUARDAR EVALUACIÓN ────────────────────────────────────────────────────────
function guardarEvaluacion(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  const HEADERS = [
    'ID', 'Fecha', 'Supervisor', 'Cosechero', 'Código',
    'Frente/Lote', 'Tipo Palma', 'Días Ciclo', 'Fuera Ciclo',
    'Racimos Maduros', 'Racimos Sobremaduro', 'Racimos Verdes', 'Racimos Podridos', 'Racimos Dejados',
    'Total Racimos', '% Verde',
    'Calificación Nota', 'Calificación Label',
    'Fruto Suelto', 'Daño Mecánico', 'Pedúnculo Largo', 'Disposición Hoja', 'Plato Enmalezado',
    'Prom Parámetros', 'Observaciones',
    'Lat', 'Lng', 'Precisión', 'Hora',
    'Firmó', 'Firma', 'Timestamp'
  ];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground('#3B6D11').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // Agregar columnas nuevas si la hoja ya existía de antes (sin lat/lng/hora, o sin firma)
    const COLUMNAS_NUEVAS = ['Lat', 'Lng', 'Precisión', 'Hora', 'Firmó', 'Firma'];
    let hdrsExist = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    COLUMNAS_NUEVAS.forEach(col => {
      if (!hdrsExist.includes(col)) {
        const nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol, 1, 1).setValue(col)
          .setBackground('#3B6D11').setFontColor('#fff').setFontWeight('bold');
        hdrsExist = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      }
    });
  }

  const id = data.id || ('EV-' + Utilities.formatDate(new Date(), 'America/Bogota', 'yyyyMMddHHmmss')
    + '-' + Math.random().toString(36).substr(2,4).toUpperCase());
  const ts = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');

  // Verificar duplicado
  if (data.id) {
    const existing = sheet.getDataRange().getValues();
    for (let i = 1; i < existing.length; i++) {
      if (existing[i][0] === data.id) return { ok: true, duplicado: true, id: data.id };
    }
  }

  // Leer headers actuales y armar fila dinámicamente
  const hdrsActuales = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const valores = {
    'ID':                 id,
    'Fecha':              data.fecha ? String(data.fecha) : '',
    'Supervisor':         data.supervisor || '',
    'Cosechero':          data.cosechero || '',
    'Código':             data.cosechCod || '',
    'Frente/Lote':        data.lote || '',
    'Tipo Palma':         data.palma || '',
    'Días Ciclo':         data.diasCiclo || 0,
    'Fuera Ciclo':        data.fueraCiclo ? 'SÍ' : 'NO',
    'Racimos Maduros':    data.racimos?.maduro || 0,
    'Racimos Sobremaduro':data.racimos?.sobremaduro || 0,
    'Racimos Verdes':     data.racimos?.verde || 0,
    'Racimos Podridos':   data.racimos?.podrido || 0,
    'Racimos Dejados':    data.racimos?.dejados || 0,
    'Total Racimos':      data.totalRacimos || 0,
    '% Verde':            data.pctVerde || 0,
    'Calificación Nota':  data.cal?.nota || '',
    'Calificación Label': data.cal?.label || '',
    'Fruto Suelto':       data.pts?.suelto || 0,
    'Daño Mecánico':      data.pts?.danio  || 0,
    'Pedúnculo Largo':    data.pts?.pedu   || 0,
    'Disposición Hoja':   data.pts?.hoja   || 0,
    'Plato Enmalezado':   data.pts?.plato  || 0,
    'Prom Parámetros':    data.prom || 0,
    'Observaciones':      data.obs || '',
    'Lat':                data.lat != null ? data.lat : '',
    'Lng':                data.lng != null ? data.lng : '',
    'Precisión':          data.acc != null ? data.acc : '',
    'Hora':               data.hora || '',
    'Firmó':              data.firma ? 'SÍ' : 'NO',
    'Firma':              data.firma || '',
    'Timestamp':          ts
  };

  const row = hdrsActuales.map(h => valores[h] !== undefined ? valores[h] : '');
  sheet.appendRow(row);

  // Forzar fecha como texto
  const idxFecha = hdrsActuales.indexOf('Fecha') + 1;
  if (idxFecha > 0 && data.fecha) {
    const fc = sheet.getRange(sheet.getLastRow(), idxFecha);
    fc.setNumberFormat('@STRING@');
    fc.setValue(String(data.fecha));
  }

  actualizarAnalisis();

  const lastRow  = sheet.getLastRow();
  const pctVerde = parseFloat(data.pctVerde) || 0;
  let bg = '#EAF3DE';
  if (pctVerde > 3 || data.fueraCiclo) bg = '#FCEBEB';
  else if (pctVerde > 1) bg = '#FAEEDA';
  sheet.getRange(lastRow, 1, 1, row.length).setBackground(bg);

  return { ok: true, id };
}

// ── LISTAR EVALUACIONES ───────────────────────────────────────────────────────
function listarEvaluaciones(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const NUMS    = ['Racimos Maduros','Racimos Sobremaduro','Racimos Verdes','Racimos Podridos','Racimos Dejados',
                   'Total Racimos','% Verde','Días Ciclo','Fruto Suelto','Daño Mecánico',
                   'Pedúnculo Largo','Disposición Hoja','Plato Enmalezado','Prom Parámetros',
                   'Lat','Lng','Precisión'];

  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h === 'Firma') return; // no enviar imagen al listar (muy pesada)
      if (h === 'Fecha') {
        if (row[i] instanceof Date) {
          obj[h] = Utilities.formatDate(row[i], 'America/Bogota', 'yyyy-MM-dd');
        } else {
          obj[h] = String(row[i]).trim().substring(0, 10);
        }
      } else if (NUMS.includes(h)) {
        obj[h] = (row[i] === '' || row[i] === null || row[i] === undefined) ? '' : (parseFloat(row[i]) || 0);
      } else if (h === 'Fuera Ciclo') {
        obj[h] = row[i] === true || String(row[i]).toUpperCase() === 'SÍ' ? 'SÍ' : 'NO';
      } else {
        obj[h] = row[i] !== null && row[i] !== undefined ? String(row[i]).trim() : '';
      }
    });
    return obj;
  }).filter(obj => obj['ID']);

  let result = data;
  if (params.fecha)      result = result.filter(r => r['Fecha'] === params.fecha);
  if (params.supervisor) result = result.filter(r => r['Supervisor'] === params.supervisor);

  return { ok: true, data: result };
}

// ── RESUMEN ───────────────────────────────────────────────────────────────────
function getResumen(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { ok: true, resumen: { total: 0, porSupervisor: {} } };
  }

  const rows = sheet.getDataRange().getValues();
  const hdrs = rows[0];
  const data = rows.slice(1).map(row => {
    const o = {};
    hdrs.forEach((h, i) => {
      if (h === 'Fecha' && row[i] instanceof Date) {
        o[h] = Utilities.formatDate(row[i], 'America/Bogota', 'yyyy-MM-dd');
      } else { o[h] = row[i]; }
    });
    return o;
  });

  const hoy     = params.fecha || Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd');
  const hoyData = data.filter(r => r['Fecha'] === hoy);

  const porSupervisor = {};
  ['Juan Carlos Oviedo', 'Brayan Pérez'].forEach(sup => {
    const evSup  = data.filter(r => r['Supervisor'] === sup);
    const evHoy  = hoyData.filter(r => r['Supervisor'] === sup);
    const racTot = evSup.reduce((a, r) => a + (parseFloat(r['Total Racimos'])||0), 0);
    const verdes = evSup.reduce((a, r) => a + (parseFloat(r['Racimos Verdes'])||0), 0);
    porSupervisor[sup] = {
      totalEvals:   evSup.length,
      evHoy:        evHoy.length,
      metaCumplida: evHoy.length >= 4,
      totalRacimos: racTot,
      pctVerde:     racTot > 0 ? Math.round(verdes/racTot*1000)/10 : 0,
      fueraCiclo:   evSup.filter(r => r['Fuera Ciclo'] === 'SÍ').length,
    };
  });

  return {
    ok: true,
    resumen: {
      total:      data.length,
      hoy:        hoyData.length,
      racimosHoy: hoyData.reduce((a, r) => a + (parseFloat(r['Total Racimos'])||0), 0),
      porSupervisor
    }
  };
}

// ── RUTAS GPS ─────────────────────────────────────────────────────────────────
const RUTAS_SHEET = 'Rutas';

function getRutaSheet(ss) {
  let sheet = ss.getSheetByName(RUTAS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RUTAS_SHEET);
    const headers = ['Fecha', 'Supervisor', 'Lat', 'Lng', 'Precisión', 'Hora', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#185FA5').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function guardarRuta(body) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = getRutaSheet(ss);
  const puntos = (body && body.puntos) ? body.puntos : [];
  if (!puntos.length) return { ok: true, insertados: 0 };

  const existentes = new Set();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 7, sheet.getLastRow()-1, 1)
      .getValues().forEach(r => existentes.add(String(r[0])));
  }

  const filas = puntos
    .filter(p => p.ts && !existentes.has(p.ts) && p.lat && p.lng)
    .map(p => [
      String(p.fecha || ''), String(p.sup || ''),
      parseFloat(p.lat), parseFloat(p.lng),
      parseInt(p.acc) || 0, String(p.hora || ''), String(p.ts)
    ]);

  if (filas.length > 0) {
    sheet.getRange(sheet.getLastRow()+1, 1, filas.length, 7).setValues(filas);
  }
  return { ok: true, insertados: filas.length };
}

function getRuta(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(RUTAS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, puntos: [] };

  const fecha = params.fecha || '';
  const sup   = params.sup   || '';
  const rows  = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues();
  const puntos = rows
    .filter(r => {
      const fOk = !fecha || String(r[0]) === fecha;
      const sOk = !sup   || String(r[1]) === sup;
      return fOk && sOk && r[2] && r[3];
    })
    .map(r => ({
      fecha: String(r[0]), sup: String(r[1]),
      lat: parseFloat(r[2]), lng: parseFloat(r[3]),
      acc: parseInt(r[4])||0, hora: String(r[5]), ts: String(r[6])
    }));

  return { ok: true, puntos };
}

// ── PERSONAL ──────────────────────────────────────────────────────────────────
const PERSONAL_SHEET = 'Personal';

function listarPersonal() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PERSONAL_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(PERSONAL_SHEET);
    const headers = ['Código', 'Nombre', 'Supervisor', 'SupIdx', 'Activo'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#3B6D11').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    const ejemplos = [
      ['C001','Pedro Ramírez',  'Juan Carlos Oviedo',0,true],
      ['C002','Luis Cárdenas',  'Juan Carlos Oviedo',0,true],
      ['C003','Miguel Torres',  'Juan Carlos Oviedo',0,true],
      ['C004','Andrés Salcedo', 'Juan Carlos Oviedo',0,true],
      ['C005','Jorge Patiño',   'Juan Carlos Oviedo',0,true],
      ['C006','Carlos Ruiz',    'Brayan Pérez',1,true],
      ['C007','Héctor Mora',    'Brayan Pérez',1,true],
      ['C008','Fabio Quintero', 'Brayan Pérez',1,true],
      ['C009','Edwin Castro',   'Brayan Pérez',1,true],
      ['C010','Nelson Vega',    'Brayan Pérez',1,true],
    ];
    sheet.getRange(2, 1, ejemplos.length, 5).setValues(ejemplos);
  }

  if (sheet.getLastRow() < 2) return { ok: true, data: [] };
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data    = rows.slice(1).map(row => {
    const obj = {}; headers.forEach((h,i) => obj[h] = row[i]);
    return {
      cod:    String(obj['Código']).trim(),
      nombre: String(obj['Nombre']).trim(),
      supIdx: parseInt(obj['SupIdx'])||0,
      activo: obj['Activo']===true || String(obj['Activo']).toLowerCase()==='true',
    };
  }).filter(p => p.cod);

  return { ok: true, data };
}

function togglePersonal(cod, nuevoEstado) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PERSONAL_SHEET);
  if (!sheet) return { ok: false, error: 'Hoja Personal no existe' };
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(cod).trim()) {
      sheet.getRange(i+1, 5).setValue(nuevoEstado);
      sheet.getRange(i+1, 1, 1, 5).setBackground(nuevoEstado ? '#EAF3DE' : '#F5F5F5');
      return { ok: true, cod, activo: nuevoEstado };
    }
  }
  return { ok: false, error: 'Cosechero no encontrado: ' + cod };
}

// ── ANÁLISIS ──────────────────────────────────────────────────────────────────
function crearHojasAnalisis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  crearResumenDiario(ss);
  crearResumenSemanal(ss);
  return { ok: true, msg: 'Hojas de análisis creadas/actualizadas' };
}

function actualizarAnalisis() {
  try { crearHojasAnalisis(); } catch(e) { Logger.log('Error análisis: ' + e); }
}

function crearResumenDiario(ss) {
  const HOJA = 'Resumen Diario';
  var sheet = ss.getSheetByName(HOJA);
  if (!sheet) sheet = ss.insertSheet(HOJA);
  sheet.clear();

  const src = ss.getSheetByName(SHEET_NAME);
  if (!src || src.getLastRow() < 2) { sheet.getRange('A1').setValue('Sin evaluaciones aún'); return; }

  const rows = src.getDataRange().getValues();
  const hdrs = rows[0];
  const data = rows.slice(1).map(r => { const o={}; hdrs.forEach((h,i)=>o[h]=r[i]); return o; });

  const porFecha = {};
  data.forEach(r => {
    const f = r['Fecha'] ? String(r['Fecha']).substring(0,10) : '';
    if (!f) return;
    if (!porFecha[f]) porFecha[f] = [];
    porFecha[f].push(r);
  });
  const fechas = Object.keys(porFecha).sort().reverse();

  const hdrsOut = ['Fecha','Evaluaciones','Racimos Total','% Verde','% Maduro','% Sobremaduro',
                   '% Podrido','JCO Evals','JCO % Verde','Brayan Evals','Brayan % Verde',
                   'Fuera Ciclo','Cal A','Cal B','Cal C+D'];
  sheet.getRange(1,1,1,hdrsOut.length).setValues([hdrsOut]);
  sheet.getRange(1,1,1,hdrsOut.length).setBackground('#3B6D11').setFontColor('#fff').setFontWeight('bold');

  const filas = fechas.map(f => {
    const ev = porFecha[f];
    const totRac  = ev.reduce((a,r)=>a+(parseFloat(r['Total Racimos'])||0),0);
    const verdes  = ev.reduce((a,r)=>a+(parseFloat(r['Racimos Verdes'])||0),0);
    const maduros = ev.reduce((a,r)=>a+(parseFloat(r['Racimos Maduros'])||0),0);
    const sobre   = ev.reduce((a,r)=>a+(parseFloat(r['Racimos Sobremaduro'])||0),0);
    const podrido = ev.reduce((a,r)=>a+(parseFloat(r['Racimos Podridos'])||0),0);
    const pctV=totRac>0?Math.round(verdes/totRac*1000)/10:0;
    const pctM=totRac>0?Math.round(maduros/totRac*1000)/10:0;
    const pctS=totRac>0?Math.round(sobre/totRac*1000)/10:0;
    const pctP=totRac>0?Math.round(podrido/totRac*1000)/10:0;
    const jco  = ev.filter(r=>r['Supervisor']==='Juan Carlos Oviedo');
    const bray = ev.filter(r=>r['Supervisor']==='Brayan Pérez');
    const rJco  = jco.reduce((a,r)=>a+(parseFloat(r['Total Racimos'])||0),0)||1;
    const rBray = bray.reduce((a,r)=>a+(parseFloat(r['Total Racimos'])||0),0)||1;
    const pvJco  = jco.length  ? Math.round(jco.reduce((a,r)=>a+(parseFloat(r['Racimos Verdes'])||0),0)/rJco*1000)/10:0;
    const pvBray = bray.length ? Math.round(bray.reduce((a,r)=>a+(parseFloat(r['Racimos Verdes'])||0),0)/rBray*1000)/10:0;
    const fuera=ev.filter(r=>r['Fuera Ciclo']==='SÍ').length;
    const calA=ev.filter(r=>r['Calificación Nota']==='A').length;
    const calB=ev.filter(r=>r['Calificación Nota']==='B').length;
    const calCD=ev.filter(r=>['C','D'].includes(r['Calificación Nota'])).length;
    return [f,ev.length,totRac,pctV,pctM,pctS,pctP,jco.length,pvJco,bray.length,pvBray,fuera,calA,calB,calCD];
  });

  if (filas.length>0) {
    sheet.getRange(2,1,filas.length,hdrsOut.length).setValues(filas);
    for(let i=0;i<filas.length;i++){
      const pv=filas[i][3];
      sheet.getRange(i+2,4).setBackground(pv>3?'#FCEBEB':'#EAF3DE').setFontColor(pv>3?'#A32D2D':'#3B6D11').setFontWeight('bold');
      if(filas[i][11]>0) sheet.getRange(i+2,12).setBackground('#FAEEDA').setFontColor('#854F0B');
    }
  }
  sheet.autoResizeColumns(1,hdrsOut.length);
  sheet.setFrozenRows(1);
}

function crearResumenSemanal(ss) {
  const HOJA = 'Resumen Semanal';
  var sheet = ss.getSheetByName(HOJA);
  if (!sheet) sheet = ss.insertSheet(HOJA);
  sheet.clear();

  const src = ss.getSheetByName(SHEET_NAME);
  if (!src || src.getLastRow() < 2) { sheet.getRange('A1').setValue('Sin evaluaciones aún'); return; }

  const rows = src.getDataRange().getValues();
  const hdrs = rows[0];
  const data = rows.slice(1).map(r => { const o={}; hdrs.forEach((h,i)=>o[h]=r[i]); return o; });

  function semanaISO(fechaStr) {
    const d=new Date(fechaStr); if(isNaN(d)) return 'N/A';
    const thu=new Date(d); thu.setDate(d.getDate()-((d.getDay()+6)%7)+3);
    const year=thu.getFullYear();
    const week=Math.ceil(((thu-new Date(year,0,4))/864e5+1)/7);
    return year+'-S'+String(week).padStart(2,'0');
  }

  const porSemana = {};
  data.forEach(r=>{
    const s=semanaISO(String(r['Fecha']).substring(0,10));
    if(!porSemana[s]) porSemana[s]=[];
    porSemana[s].push(r);
  });
  const semanas=Object.keys(porSemana).filter(s=>s!=='N/A').sort().reverse();

  const hdrsOut=['Semana','Días activos','Total Evals','Racimos Total','% Verde Sem.',
                 'Cosecheros eval.','Mejor cosechero','Peor cosechero',
                 'JCO Evals','Brayan Evals','Lotes fuera ciclo','Cal A %','Cal C+D %'];
  sheet.getRange(1,1,1,hdrsOut.length).setValues([hdrsOut]);
  sheet.getRange(1,1,1,hdrsOut.length).setBackground('#185FA5').setFontColor('#fff').setFontWeight('bold');

  const filas=semanas.map(s=>{
    const ev=porSemana[s];
    const totRac=ev.reduce((a,r)=>a+(parseFloat(r['Total Racimos'])||0),0);
    const verdes=ev.reduce((a,r)=>a+(parseFloat(r['Racimos Verdes'])||0),0);
    const pctV=totRac>0?Math.round(verdes/totRac*1000)/10:0;
    const dias=new Set(ev.map(r=>String(r['Fecha']).substring(0,10))).size;
    const porCos={};
    ev.forEach(r=>{
      const cod=r['Código'];
      if(!porCos[cod]) porCos[cod]={nombre:r['Cosechero'],rac:0,verde:0};
      porCos[cod].rac+=parseFloat(r['Total Racimos'])||0;
      porCos[cod].verde+=parseFloat(r['Racimos Verdes'])||0;
    });
    const cosArr=Object.values(porCos).filter(x=>x.rac>0)
      .map(x=>({...x,pct:Math.round(x.verde/x.rac*1000)/10})).sort((a,b)=>a.pct-b.pct);
    const mejor=cosArr.length>0?cosArr[0].nombre+' ('+cosArr[0].pct+'%)':'—';
    const peor=cosArr.length>1?cosArr[cosArr.length-1].nombre+' ('+cosArr[cosArr.length-1].pct+'%)':'—';
    const jco=ev.filter(r=>r['Supervisor']==='Juan Carlos Oviedo').length;
    const bray=ev.filter(r=>r['Supervisor']==='Brayan Pérez').length;
    const fuera=ev.filter(r=>r['Fuera Ciclo']==='SÍ').length;
    const calA=ev.length>0?Math.round(ev.filter(r=>r['Calificación Nota']==='A').length/ev.length*100):0;
    const calCD=ev.length>0?Math.round(ev.filter(r=>['C','D'].includes(r['Calificación Nota'])).length/ev.length*100):0;
    return [s,dias,ev.length,totRac,pctV,cosArr.length,mejor,peor,jco,bray,fuera,calA+'%',calCD+'%'];
  });

  if(filas.length>0){
    sheet.getRange(2,1,filas.length,hdrsOut.length).setValues(filas);
    for(let i=0;i<filas.length;i++){
      const pv=filas[i][4];
      sheet.getRange(i+2,5).setBackground(pv>3?'#FCEBEB':'#EAF3DE').setFontColor(pv>3?'#A32D2D':'#3B6D11').setFontWeight('bold');
      if(filas[i][10]>0) sheet.getRange(i+2,11).setBackground('#FAEEDA').setFontColor('#854F0B');
    }
  }
  sheet.autoResizeColumns(1,hdrsOut.length);
  sheet.setFrozenRows(1);
}

// ── GENERAR PDF HISTORIAL ─────────────────────────────────────────────────────
function generarPDFHistorial(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: 'Sin datos' };

  const desde=params.desde||'', hasta=params.hasta||'', supFilt=params.supervisor||'', calFilt=params.cal||'';
  const rows=sheet.getDataRange().getValues(), hdrs=rows[0];
  const NUMS=['Racimos Maduros','Racimos Sobremaduro','Racimos Verdes','Racimos Podridos','Racimos Dejados',
              'Total Racimos','% Verde','Días Ciclo','Fruto Suelto','Daño Mecánico',
              'Pedúnculo Largo','Disposición Hoja','Plato Enmalezado','Prom Parámetros'];

  let data=rows.slice(1).map(r=>{
    const o={};
    hdrs.forEach((h,i)=>{
      if(h==='Fecha') o[h]=r[i] instanceof Date?Utilities.formatDate(r[i],'America/Bogota','yyyy-MM-dd'):String(r[i]).trim().substring(0,10);
      else if(NUMS.includes(h)) o[h]=parseFloat(r[i])||0;
      else if(h==='Firma') return;
      else o[h]=r[i]!==null&&r[i]!==undefined?String(r[i]).trim():'';
    });
    return o;
  }).filter(o=>o['ID']);

  if(desde) data=data.filter(r=>r['Fecha']>=desde);
  if(hasta) data=data.filter(r=>r['Fecha']<=hasta);
  if(supFilt) data=data.filter(r=>r['Supervisor']===supFilt);
  if(calFilt) data=data.filter(r=>r['Calificación Nota']===calFilt);
  if(!data.length) return {ok:false,error:'Sin evaluaciones para los filtros seleccionados'};

  const totalRac=data.reduce((a,r)=>a+r['Total Racimos'],0);
  const totalVerde=data.reduce((a,r)=>a+r['Racimos Verdes'],0);
  const pctVG=totalRac>0?Math.round(totalVerde/totalRac*1000)/10:0;
  const calA=data.filter(r=>r['Calificación Nota']==='A').length;
  const calB=data.filter(r=>r['Calificación Nota']==='B').length;
  const calCD=data.filter(r=>['C','D'].includes(r['Calificación Nota'])).length;
  const pctColor=pctVG>3?'#A32D2D':'#3B6D11';
  const pctBg=pctVG>3?'#FCEBEB':'#EAF3DE';
  const supLabel=supFilt||'Todos los supervisores';

  const nivelParam=v=>{const n=Math.round(v);if(n>=5)return 'Excelente';if(n>=4)return 'Bueno';if(n>=3)return 'Aceptable';if(n>=2)return 'Deficiente';return 'Muy deficiente';};
  const estrellas=v=>{const n=Math.round(parseFloat(v)||0);let s='';for(let i=1;i<=5;i++)s+=`<span style="color:${i<=n?'#639922':'#ccc'};">&#9733;</span>`;return `${s} ${parseFloat(v).toFixed(1)} &mdash; ${nivelParam(v)}`;};

  const filas=data.map((r,i)=>{
    const nota=r['Calificación Nota']||'';
    const calBg=nota==='A'?'#EAF3DE':nota==='B'?'#FAEEDA':'#FCEBEB';
    const calCol=nota==='A'?'#3B6D11':nota==='B'?'#854F0B':'#A32D2D';
    const vCol=r['% Verde']>3?'#A32D2D':'#3B6D11';
    const rowBg=i%2===0?'#FFFFFF':'#F8F9F8';
    const fueraCiclo=r['Fuera Ciclo']==='SÍ';
    const dias=r['Días Ciclo'],max=r['Tipo Palma']==='hibrido'?22:13;
    const cicloColor=fueraCiclo?'#A32D2D':'#3B6D11';
    const cicloTxt=fueraCiclo?`+${dias-max}d fuera`:`${dias}d OK`;
    const firmoTag=r['Firmó']==='SÍ'?'<span style="font-size:8pt;color:#3B6D11;">✓ Firmó</span>':'<span style="font-size:8pt;color:#999;">Sin firma</span>';
    return `<tr style="background:${rowBg};border-bottom:1px solid #E8EDE8;">
      <td style="padding:10px 8px;vertical-align:top;width:80px;"><span style="font-size:9pt;color:#555;">${r['Fecha']}</span></td>
      <td style="padding:10px 8px;vertical-align:top;width:170px;">
        <div style="font-weight:bold;font-size:9pt;">${r['Cosechero']}</div>
        <div style="font-size:8pt;color:#777;">${r['Código']}</div>
        <div style="font-size:8pt;color:#555;">${r['Supervisor']}</div>
        <div style="font-size:8pt;color:#555;">${r['Frente/Lote']}</div>
        <div style="font-size:8pt;color:${cicloColor};">${cicloTxt}</div>
        <div style="margin-top:3px;">${firmoTag}</div>
      </td>
      <td style="padding:10px 8px;vertical-align:top;text-align:center;width:80px;">
        <div style="background:${calBg};border:1px solid ${calCol};border-radius:6px;padding:8px 4px;">
          <div style="font-size:18pt;font-weight:bold;color:${calCol};line-height:1;">${nota}</div>
          <div style="font-size:7pt;color:${calCol};margin-top:2px;">${r['Calificación Label']||''}</div>
          <div style="font-size:8pt;font-weight:bold;color:${vCol};margin-top:4px;">${r['% Verde'].toFixed(1)}% verde</div>
        </div>
      </td>
      <td style="padding:10px 8px;vertical-align:top;width:55px;text-align:center;">
        <div style="font-size:14pt;font-weight:bold;">${r['Total Racimos']}</div>
        <div style="font-size:7pt;color:#888;">racimos</div>
      </td>
      <td style="padding:10px 8px;vertical-align:top;width:140px;">
        <table style="width:100%;border-collapse:collapse;">
          <tr>
            <td style="width:50%;padding:2px;"><div style="background:#EAF3DE;border-radius:4px;padding:4px 6px;text-align:center;"><div style="font-size:11pt;font-weight:bold;color:#3B6D11;">${r['Racimos Maduros']}</div><div style="font-size:7pt;color:#3B6D11;">Maduro</div></div></td>
            <td style="width:50%;padding:2px;"><div style="background:#FAEEDA;border-radius:4px;padding:4px 6px;text-align:center;"><div style="font-size:11pt;font-weight:bold;color:#854F0B;">${r['Racimos Sobremaduro']}</div><div style="font-size:7pt;color:#854F0B;">Sobrem.</div></div></td>
          </tr>
          <tr>
            <td style="padding:2px;"><div style="background:#EEEDFE;border-radius:4px;padding:4px 6px;text-align:center;"><div style="font-size:11pt;font-weight:bold;color:#534AB7;">${r['Racimos Verdes']}</div><div style="font-size:7pt;color:#534AB7;">Verde</div></div></td>
            <td style="padding:2px;"><div style="background:#FCEBEB;border-radius:4px;padding:4px 6px;text-align:center;"><div style="font-size:11pt;font-weight:bold;color:#A32D2D;">${r['Racimos Podridos']}</div><div style="font-size:7pt;color:#A32D2D;">Podrido</div></div></td>
          </tr>
        </table>
      </td>
      <td style="padding:10px 8px;vertical-align:top;">
        <div style="font-size:8pt;margin-bottom:4px;"><b>Fruto suelto:</b><br>${estrellas(r['Fruto Suelto'])}</div>
        <div style="font-size:8pt;margin-bottom:4px;"><b>Daño mecánico:</b><br>${estrellas(r['Daño Mecánico'])}</div>
        <div style="font-size:8pt;margin-bottom:4px;"><b>Pedúnculo largo:</b><br>${estrellas(r['Pedúnculo Largo'])}</div>
        <div style="font-size:8pt;margin-bottom:4px;"><b>Disposición hoja:</b><br>${estrellas(r['Disposición Hoja'])}</div>
        <div style="font-size:8pt;"><b>Plato enmalezado:</b><br>${estrellas(r['Plato Enmalezado'])}</div>
        ${r['Observaciones']?`<div style="margin-top:6px;font-size:7pt;color:#888;font-style:italic;border-top:1px solid #eee;padding-top:4px;">Obs: ${r['Observaciones']}</div>`:''}
      </td>
    </tr>`;
  }).join('');

  const html=`<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>body{font-family:Arial,sans-serif;margin:0;padding:0;}table{border-collapse:collapse;}</style>
</head><body>
<table style="width:100%;background:#2D6011;margin-bottom:0;">
  <tr><td style="padding:16px 20px;">
    <div style="font-size:18pt;font-weight:800;color:#fff;letter-spacing:1px;">PALMAGRANDE S.A.S</div>
    <div style="font-size:10pt;color:rgba(255,255,255,.85);margin-top:3px;">Historial de evaluaciones de cosecha</div>
    <div style="font-size:9pt;color:rgba(255,255,255,.7);margin-top:3px;">${supLabel} &nbsp;&middot;&nbsp; ${desde} al ${hasta}</div>
  </td></tr>
</table>
<table style="width:100%;border-bottom:2px solid #2D6011;">
  <tr>
    <td style="width:25%;padding:12px 16px;text-align:center;border-right:1px solid #E0E8E0;"><div style="font-size:22pt;font-weight:800;">${data.length}</div><div style="font-size:9pt;color:#666;">Evaluaciones</div></td>
    <td style="width:25%;padding:12px 16px;text-align:center;border-right:1px solid #E0E8E0;"><div style="font-size:22pt;font-weight:800;">${totalRac}</div><div style="font-size:9pt;color:#666;">Total racimos</div></td>
    <td style="width:25%;padding:12px 16px;text-align:center;background:${pctBg};border-right:1px solid #E0E8E0;"><div style="font-size:22pt;font-weight:800;color:${pctColor};">${pctVG.toFixed(1)}%</div><div style="font-size:9pt;color:${pctColor};">% Verde global</div></td>
    <td style="width:25%;padding:12px 16px;text-align:center;"><div style="font-size:13pt;font-weight:800;"><span style="color:#3B6D11;">A:${calA}</span>&nbsp;<span style="color:#854F0B;">B:${calB}</span>&nbsp;<span style="color:#A32D2D;">C/D:${calCD}</span></div><div style="font-size:9pt;color:#666;margin-top:4px;">Calificaciones</div></td>
  </tr>
</table>
<table style="width:100%;">
  <thead><tr style="background:#3B6D11;">
    <th style="padding:8px;color:#fff;font-size:9pt;text-align:left;">Fecha</th>
    <th style="padding:8px;color:#fff;font-size:9pt;text-align:left;">Cosechero</th>
    <th style="padding:8px;color:#fff;font-size:9pt;text-align:center;">Calificación</th>
    <th style="padding:8px;color:#fff;font-size:9pt;text-align:center;">Total</th>
    <th style="padding:8px;color:#fff;font-size:9pt;text-align:center;">Racimos por madurez</th>
    <th style="padding:8px;color:#fff;font-size:9pt;text-align:left;">Parámetros de calidad</th>
  </tr></thead>
  <tbody>${filas}</tbody>
</table>
<table style="width:100%;margin-top:12px;border-top:1px solid #ddd;">
  <tr>
    <td style="padding:8px 12px;font-size:8pt;color:#999;">Palma Grande S.A.S — Sistema de supervisión de cosecha</td>
    <td style="padding:8px 12px;font-size:8pt;color:#999;text-align:right;">Generado: ${Utilities.formatDate(new Date(),'America/Bogota','dd/MM/yyyy HH:mm')}</td>
  </tr>
</table>
</body></html>`;

  const blob=Utilities.newBlob(html,'text/html','reporte.html');
  const pdfBlob=blob.getAs('application/pdf');
  return { ok:true, base64:Utilities.base64Encode(pdfBlob.getBytes()), filename:`Cosecha_${desde}_${hasta}.pdf` };
}

// ── GENERAR PDF REPORTE INDIVIDUAL/GRUPAL ─────────────────────────────────────
function generarPDFReporte(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: 'Sin datos' };

  const supFilt=params.supervisor||'', cosechCod=params.cosechero||'',
        periodo=params.periodo||'dia', fecha=params.fecha||'';
  const rows=sheet.getDataRange().getValues(), hdrs=rows[0];
  const NUMS=['Racimos Maduros','Racimos Sobremaduro','Racimos Verdes','Racimos Podridos','Racimos Dejados',
              'Total Racimos','% Verde','Días Ciclo','Fruto Suelto','Daño Mecánico',
              'Pedúnculo Largo','Disposición Hoja','Plato Enmalezado','Prom Parámetros'];

  let data=rows.slice(1).map(r=>{
    const o={};
    hdrs.forEach((h,i)=>{
      if(h==='Fecha') o[h]=r[i] instanceof Date?Utilities.formatDate(r[i],'America/Bogota','yyyy-MM-dd'):String(r[i]).trim().substring(0,10);
      else if(NUMS.includes(h)) o[h]=parseFloat(r[i])||0;
      else o[h]=r[i]!==null&&r[i]!==undefined?String(r[i]).trim():'';
    });
    return o;
  }).filter(o=>o['ID']);

  if(supFilt) data=data.filter(r=>r['Supervisor']===supFilt);
  if(periodo==='dia'&&fecha) {
    data=data.filter(r=>r['Fecha']===fecha);
  } else if(periodo==='semana'&&fecha) {
    const d=new Date(fecha);
    const lun=new Date(d); lun.setDate(d.getDate()-((d.getDay()+6)%7));
    const dom=new Date(lun); dom.setDate(lun.getDate()+6);
    const ini=Utilities.formatDate(lun,'America/Bogota','yyyy-MM-dd');
    const fin=Utilities.formatDate(dom,'America/Bogota','yyyy-MM-dd');
    data=data.filter(r=>r['Fecha']>=ini&&r['Fecha']<=fin);
  }

  const esIndividual=!!cosechCod;
  if(esIndividual) data=data.filter(r=>r['Código']===cosechCod);
  if(!data.length) return {ok:false,error:'Sin evaluaciones'};

  const periodoLabel=periodo==='dia'?`Fecha: ${fecha}`:`Semana del ${fecha}`;
  const html=esIndividual?buildHTMLIndividual(data,params,periodoLabel):buildHTMLGrupal(data,params,periodoLabel);
  const blob=Utilities.newBlob(html,'text/html','reporte.html');
  const pdfBlob=blob.getAs('application/pdf');
  const nombre=esIndividual?`Reporte_${data[0]['Cosechero'].split(' ')[0]}_${fecha}.pdf`:`Reporte_Grupo_${fecha}.pdf`;
  return {ok:true,base64:Utilities.base64Encode(pdfBlob.getBytes()),filename:nombre};
}

function buildHTMLIndividual(data, params, periodoLabel) {
  const nombre=data[0]['Cosechero'], cod=data[0]['Código'], sup=data[0]['Supervisor'];
  const firma=data[0]['Firma']||'', firmo=data[0]['Firmó']==='SÍ';
  const totRac=data.reduce((a,r)=>a+r['Total Racimos'],0);
  const verdes=data.reduce((a,r)=>a+r['Racimos Verdes'],0);
  const pctV=totRac>0?Math.round(verdes/totRac*1000)/10:0;
  const rM=data.reduce((a,r)=>a+r['Racimos Maduros'],0);
  const rS=data.reduce((a,r)=>a+r['Racimos Sobremaduro'],0);
  const rV=data.reduce((a,r)=>a+r['Racimos Verdes'],0);
  const rP=data.reduce((a,r)=>a+r['Racimos Podridos'],0);
  const rD=data.reduce((a,r)=>a+(r['Racimos Dejados']||0),0);
  const n=data.length||1;
  const pSuelto=(data.reduce((a,r)=>a+r['Fruto Suelto'],0)/n).toFixed(1);
  const pDanio=(data.reduce((a,r)=>a+r['Daño Mecánico'],0)/n).toFixed(1);
  const pPedu=(data.reduce((a,r)=>a+r['Pedúnculo Largo'],0)/n).toFixed(1);
  const pHoja=(data.reduce((a,r)=>a+r['Disposición Hoja'],0)/n).toFixed(1);
  const pPlato=(data.reduce((a,r)=>a+(r['Plato Enmalezado']||0),0)/n).toFixed(1);
  const obs=data.filter(r=>r['Observaciones']).map(r=>`<li>${r['Observaciones']}</li>`).join('');

  let nota;
  if(pctV<=1) nota='A'; else if(pctV<=3) nota='B'; else if(pctV<=5) nota='C'; else nota='D';
  const calBg=nota==='A'?'#EAF3DE':nota==='B'?'#FAEEDA':'#FCEBEB';
  const calCol=nota==='A'?'#3B6D11':nota==='B'?'#854F0B':'#A32D2D';
  const calLabel=nota==='A'?(pctV<=1?'Excelente':'Muy bueno'):nota==='B'?'Aceptable':nota==='C'?'Deficiente':'Crítico';

  const nivelParam=v=>{const vi=Math.round(parseFloat(v)||0);if(vi>=5)return 'Excelente';if(vi>=4)return 'Bueno';if(vi>=3)return 'Aceptable';if(vi>=2)return 'Deficiente';return 'Muy deficiente';};
  const estrellas=v=>{const vi=Math.round(parseFloat(v)||0);let s='';for(let i=1;i<=5;i++)s+=`<span style="color:${i<=vi?'#639922':'#ccc'};font-size:13pt;">&#9733;</span>`;return s+` <b>${parseFloat(v).toFixed(1)}</b> &mdash; ${nivelParam(v)}`;};

  const pctM=totRac>0?Math.round(rM/totRac*100):0;
  const pctS=totRac>0?Math.round(rS/totRac*100):0;
  const pctV2=totRac>0?Math.round(rV/totRac*100):0;
  const pctP=totRac>0?Math.round(rP/totRac*100):0;

  const firmaHTML=firmo&&firma
    ?`<table style="border:1px solid #ddd;border-radius:6px;margin-bottom:14px;"><tr><td style="padding:12px;text-align:center;">
        <div style="font-size:10pt;font-weight:bold;color:#555;margin-bottom:8px;text-transform:uppercase;letter-spacing:1px;">Firma del cosechero</div>
        <img src="${firma}" style="max-width:280px;max-height:110px;border:1px solid #eee;border-radius:4px;">
        <div style="font-size:8pt;color:#999;margin-top:6px;">Firma digital registrada en campo</div>
      </td></tr></table>`
    :`<table style="border:1px dashed #ccc;border-radius:6px;margin-bottom:14px;"><tr><td style="padding:12px;text-align:center;color:#999;">
        <div style="font-size:10pt;">Sin firma — el cosechero no firmó en esta evaluación</div>
      </td></tr></table>`;

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>body{font-family:Arial,sans-serif;font-size:11pt;padding:20px;color:#222;}
h2{font-size:10pt;font-weight:bold;color:#555;text-transform:uppercase;letter-spacing:1px;
   border-bottom:1px solid #ddd;padding-bottom:4px;margin:14px 0 8px;}
table{width:100%;border-collapse:collapse;}td,th{padding:6px 8px;vertical-align:middle;}</style>
</head><body>
<table style="background:#2d6011;border-radius:6px;margin-bottom:14px;"><tr><td style="padding:14px;color:#fff;">
  <div style="font-size:16pt;font-weight:bold;letter-spacing:1px;">PALMAGRANDE S.A.S</div>
  <div style="font-size:10pt;opacity:.85;">Reporte individual de cosecha</div>
  <div style="font-size:9pt;opacity:.7;">${periodoLabel} &nbsp;·&nbsp; ${data.length} evaluación${data.length>1?'es':''}</div>
</td></tr></table>
<table style="border:1px solid #ddd;border-radius:6px;margin-bottom:14px;"><tr>
  <td style="width:60px;text-align:center;">
    <div style="width:50px;height:50px;border-radius:25px;background:#EAF3DE;text-align:center;line-height:50px;font-size:16pt;font-weight:bold;color:#3B6D11;">
      ${nombre.split(' ').map(w=>w[0]).slice(0,2).join('')}
    </div>
  </td>
  <td><div style="font-size:14pt;font-weight:bold;">${nombre}</div>
      <div style="font-size:10pt;color:#666;">${cod} &nbsp;·&nbsp; ${sup}</div></td>
</tr></table>
<table style="background:${calBg};border:1px solid ${calCol};border-radius:6px;margin-bottom:14px;"><tr><td style="text-align:center;padding:14px;">
  <div style="font-size:48pt;font-weight:bold;color:${calCol};line-height:1;">${nota}</div>
  <div style="font-size:13pt;font-weight:bold;color:${calCol};">${calLabel}</div>
  <div style="font-size:10pt;color:${calCol};margin-top:4px;">${pctV.toFixed(1)}% de racimos verdes &nbsp;·&nbsp; Total: ${totRac} racimos</div>
</td></tr></table>
<h2>Racimos por madurez</h2>
<table style="margin-bottom:14px;"><tr>
  <td style="width:20%;padding:4px;"><table style="background:#EAF3DE;border-radius:6px;"><tr><td style="text-align:center;padding:10px;"><div style="font-size:22pt;font-weight:bold;color:#3B6D11;">${rM}</div><div style="font-size:9pt;color:#3B6D11;">${pctM}% Maduro</div></td></tr></table></td>
  <td style="width:20%;padding:4px;"><table style="background:#FAEEDA;border-radius:6px;"><tr><td style="text-align:center;padding:10px;"><div style="font-size:22pt;font-weight:bold;color:#854F0B;">${rS}</div><div style="font-size:9pt;color:#854F0B;">${pctS}% Sobremaduro</div></td></tr></table></td>
  <td style="width:20%;padding:4px;"><table style="background:#EEEDFE;border-radius:6px;"><tr><td style="text-align:center;padding:10px;"><div style="font-size:22pt;font-weight:bold;color:#534AB7;">${rV}</div><div style="font-size:9pt;color:#534AB7;">${pctV2}% Verde</div></td></tr></table></td>
  <td style="width:20%;padding:4px;"><table style="background:#FCEBEB;border-radius:6px;"><tr><td style="text-align:center;padding:10px;"><div style="font-size:22pt;font-weight:bold;color:#A32D2D;">${rP}</div><div style="font-size:9pt;color:#A32D2D;">${pctP}% Podrido</div></td></tr></table></td>
  <td style="width:20%;padding:4px;"><table style="background:#F0EBF8;border-radius:6px;"><tr><td style="text-align:center;padding:10px;"><div style="font-size:22pt;font-weight:bold;color:#5B2D9E;">${rD}</div><div style="font-size:9pt;color:#5B2D9E;">Dejados</div></td></tr></table></td>
</tr></table>
<h2>Parámetros de calidad del trabajo</h2>
<table style="border:1px solid #eee;border-radius:6px;margin-bottom:14px;">
  <tr style="border-bottom:1px solid #eee;"><td style="padding:8px;">Fruto suelto en suelo</td><td style="padding:8px;text-align:right;">${estrellas(pSuelto)}</td></tr>
  <tr style="border-bottom:1px solid #eee;"><td style="padding:8px;">Daño mecánico / cortes</td><td style="padding:8px;text-align:right;">${estrellas(pDanio)}</td></tr>
  <tr style="border-bottom:1px solid #eee;"><td style="padding:8px;">Pedúnculo largo</td><td style="padding:8px;text-align:right;">${estrellas(pPedu)}</td></tr>
  <tr style="border-bottom:1px solid #eee;"><td style="padding:8px;">Disposición de hoja</td><td style="padding:8px;text-align:right;">${estrellas(pHoja)}</td></tr>
  <tr><td style="padding:8px;">Plato enmalezado</td><td style="padding:8px;text-align:right;">${estrellas(pPlato)}</td></tr>
</table>
${obs?`<h2>Observaciones</h2><ul style="color:#666;font-size:10pt;padding-left:16px;margin-bottom:14px;">${obs}</ul>`:''}
${firmaHTML}
<table style="margin-top:14px;border-top:1px solid #ddd;"><tr>
  <td style="font-size:9pt;color:#999;">Palma Grande S.A.S — Supervisión de cosecha</td>
  <td style="font-size:9pt;color:#999;text-align:right;">Generado: ${Utilities.formatDate(new Date(),'America/Bogota','dd/MM/yyyy HH:mm')}</td>
</tr></table>
</body></html>`;
}

function buildHTMLGrupal(data, params, periodoLabel) {
  const sup=params.supervisor||'Todos los supervisores';
  const porCos={};
  data.forEach(r=>{
    const k=r['Código'];
    if(!porCos[k]) porCos[k]={nombre:r['Cosechero'],cod:k,evals:0,rac:0,verde:0};
    porCos[k].evals++; porCos[k].rac+=r['Total Racimos']; porCos[k].verde+=r['Racimos Verdes'];
  });
  const lista=Object.values(porCos).map(co=>({...co,pct:co.rac>0?Math.round(co.verde/co.rac*1000)/10:0})).sort((a,b)=>a.pct-b.pct);
  const filas=lista.map((co,i)=>{
    let nota; if(co.pct<=1) nota='A'; else if(co.pct<=3) nota='B'; else if(co.pct<=5) nota='C'; else nota='D';
    const calBg=nota==='A'?'#EAF3DE':nota==='B'?'#FAEEDA':'#FCEBEB';
    const calCol=nota==='A'?'#3B6D11':nota==='B'?'#854F0B':'#A32D2D';
    const vCol=co.pct>3?'#A32D2D':'#3B6D11';
    return `<tr style="background:${i%2===0?'#fff':'#f5f5f5'};">
      <td style="padding:8px;text-align:center;"><span style="background:${calBg};color:${calCol};padding:3px 10px;border-radius:4px;font-weight:bold;font-size:13pt;">${nota}</span></td>
      <td style="padding:8px;"><b>${co.nombre}</b><br><span style="font-size:9pt;color:#999;">${co.cod}</span></td>
      <td style="padding:8px;text-align:center;font-weight:bold;">${co.rac}</td>
      <td style="padding:8px;text-align:center;font-weight:bold;color:${vCol};">${co.pct.toFixed(1)}%</td>
      <td style="padding:8px;text-align:center;">${co.evals}</td>
    </tr>`;
  }).join('');

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>body{font-family:Arial,sans-serif;font-size:11pt;padding:20px;color:#222;}table{width:100%;border-collapse:collapse;}td,th{padding:6px 8px;}</style>
</head><body>
<table style="background:#2d6011;margin-bottom:14px;"><tr><td style="padding:14px;color:#fff;">
  <div style="font-size:16pt;font-weight:bold;letter-spacing:1px;">PALMAGRANDE S.A.S</div>
  <div style="font-size:10pt;opacity:.85;">Resumen grupal de cosecha</div>
  <div style="font-size:9pt;opacity:.7;">${sup} &nbsp;·&nbsp; ${periodoLabel} &nbsp;·&nbsp; ${data.length} evaluaciones &nbsp;·&nbsp; ${lista.length} cosecheros</div>
</td></tr></table>
<table style="border:1px solid #ddd;">
  <thead><tr style="background:#3B6D11;color:#fff;">
    <th style="padding:8px;text-align:center;width:60px;">Cal.</th>
    <th style="padding:8px;text-align:left;">Cosechero</th>
    <th style="padding:8px;text-align:center;">Racimos</th>
    <th style="padding:8px;text-align:center;">% Verde</th>
    <th style="padding:8px;text-align:center;">Evals.</th>
  </tr></thead>
  <tbody>${filas}</tbody>
</table>
<table style="margin-top:14px;border-top:1px solid #ddd;"><tr>
  <td style="font-size:9pt;color:#999;">Palma Grande S.A.S — Supervisión de cosecha</td>
  <td style="font-size:9pt;color:#999;text-align:right;">Generado: ${Utilities.formatDate(new Date(),'America/Bogota','dd/MM/yyyy HH:mm')}</td>
</tr></table>
</body></html>`;
}

// ── TENDENCIAS ────────────────────────────────────────────────────────────────
function getTendencias(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: { semanas:[], cosecheros:{}, supervisores:{} } };

  const rows=sheet.getDataRange().getValues(), hdrs=rows[0];
  const data=rows.slice(1).map(r=>{
    const o={};
    hdrs.forEach((h,i)=>{
      if(h==='Fecha') o[h]=r[i] instanceof Date?Utilities.formatDate(r[i],'America/Bogota','yyyy-MM-dd'):String(r[i]).trim().substring(0,10);
      else if(['Racimos Verdes','Total Racimos','% Verde'].includes(h)) o[h]=parseFloat(r[i])||0;
      else if(h==='Firma') return;
      else o[h]=r[i]!==null?String(r[i]).trim():'';
    });
    return o;
  }).filter(o=>o['ID']&&o['Fecha']);

  const getSemana=fecha=>{
    const d=new Date(fecha); if(isNaN(d)) return null;
    const thu=new Date(d); thu.setDate(d.getDate()-((d.getDay()+6)%7)+3);
    const year=thu.getFullYear();
    const week=Math.ceil(((thu-new Date(year,0,4))/864e5+1)/7);
    return year+'-S'+String(week).padStart(2,'0');
  };

  const porCos={}, porSup={};
  data.forEach(r=>{
    const sem=getSemana(r['Fecha']); if(!sem) return;
    const cod=r['Código']||r['Cosechero'], sup=r['Supervisor'];
    if(!porCos[cod]) porCos[cod]={nombre:r['Cosechero'],supervisor:sup,semanas:{}};
    if(!porCos[cod].semanas[sem]) porCos[cod].semanas[sem]={verde:0,total:0};
    porCos[cod].semanas[sem].verde+=r['Racimos Verdes'];
    porCos[cod].semanas[sem].total+=r['Total Racimos'];
    if(!porSup[sup]) porSup[sup]={semanas:{}};
    if(!porSup[sup].semanas[sem]) porSup[sup].semanas[sem]=0;
    porSup[sup].semanas[sem]+=1;
  });

  const semanas=[...new Set(data.map(r=>getSemana(r['Fecha'])).filter(Boolean))].sort();
  const cosResult={};
  Object.keys(porCos).forEach(cod=>{
    const c=porCos[cod];
    cosResult[cod]={
      nombre:c.nombre, supervisor:c.supervisor,
      semanas:semanas.map(s=>{const d=c.semanas[s];if(!d||d.total===0)return null;return Math.round(d.verde/d.total*1000)/10;})
    };
  });
  const supResult={};
  Object.keys(porSup).forEach(sup=>{
    supResult[sup]={semanas:semanas.map(s=>porSup[sup].semanas[s]||0)};
  });

  return {ok:true,data:{semanas,cosecheros:cosResult,supervisores:supResult}};
}

// ── REPORTE EJECUTIVO PARA GERENCIA ────────────────────────────────────────────
// Calificación: misma fórmula que usa la app (70% % racimos verdes + 30% calidad de parámetros)
function notaGerencia(pctVerde, prom) {
  const goodVerde   = Math.max(0, 100 - pctVerde * 20);
  const goodCalidad = (prom || 0) / 5 * 100;
  const score       = goodVerde * 0.7 + goodCalidad * 0.3;
  if (score >= 85) return 'A';
  if (score >= 60) return 'B';
  if (score >= 35) return 'C';
  return 'D';
}
const NOTA_COLOR = {
  A: {bg:'#EAF3DE', col:'#3B6D11', label:'Excelente'},
  B: {bg:'#FAEEDA', col:'#854F0B', label:'Aceptable'},
  C: {bg:'#F3E3D3', col:'#A35B0B', label:'Deficiente'},
  D: {bg:'#FCEBEB', col:'#A32D2D', label:'Crítico'}
};

function generarPDFGerencia(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { ok:false, error:'Sin datos' };

  const fechaIni = params.fechaIni || '', fechaFin = params.fechaFin || '';
  const rows = sheet.getDataRange().getValues(), hdrs = rows[0];
  const NUMS = ['Racimos Maduros','Racimos Sobremaduro','Racimos Verdes','Racimos Podridos','Racimos Dejados',
                'Total Racimos','% Verde','Días Ciclo','Fruto Suelto','Daño Mecánico',
                'Pedúnculo Largo','Disposición Hoja','Plato Enmalezado','Prom Parámetros'];

  let data = rows.slice(1).map(r => {
    const o = {};
    hdrs.forEach((h,i) => {
      if (h === 'Fecha') o[h] = r[i] instanceof Date ? Utilities.formatDate(r[i],'America/Bogota','yyyy-MM-dd') : String(r[i]).trim().substring(0,10);
      else if (NUMS.includes(h)) o[h] = parseFloat(r[i]) || 0;
      else o[h] = r[i] !== null && r[i] !== undefined ? String(r[i]).trim() : '';
    });
    return o;
  }).filter(o => o['ID']);

  if (fechaIni) data = data.filter(r => r['Fecha'] >= fechaIni);
  if (fechaFin) data = data.filter(r => r['Fecha'] <= fechaFin);
  if (!data.length) return { ok:false, error:'Sin evaluaciones en ese periodo' };

  const sum = (arr,k) => arr.reduce((a,r) => a + (r[k]||0), 0);
  const pctOf = (verde,total) => total > 0 ? Math.round(verde/total*1000)/10 : 0;

  // ── KPIs globales ──
  const totRacimos  = sum(data,'Total Racimos');
  const totVerde     = sum(data,'Racimos Verdes');
  const pctVerdeGlob = pctOf(totVerde, totRacimos);
  const lotes        = [...new Set(data.map(r => r['Frente/Lote']).filter(Boolean))];
  const supervisores = [...new Set(data.map(r => r['Supervisor']).filter(Boolean))];
  const cosecherosCD = [...new Set(data.filter(r => r['Calificación Nota']==='C'||r['Calificación Nota']==='D').map(r => r['Código']))];
  const lotesFC       = [...new Set(data.filter(r => r['Fuera Ciclo']==='SÍ').map(r => r['Frente/Lote']))];
  const conGPS        = data.filter(r => r['Lat'] !== '' && r['Lat'] != null).length;
  const pctGPS         = Math.round(conGPS/data.length*100);

  // ── Tendencia semanal ──
  const getSemana = fecha => {
    const d = new Date(fecha); if (isNaN(d)) return null;
    const thu = new Date(d); thu.setDate(d.getDate()-((d.getDay()+6)%7)+3);
    const year = thu.getFullYear();
    const week = Math.ceil(((thu-new Date(year,0,4))/864e5+1)/7);
    return year+'-S'+String(week).padStart(2,'0');
  };
  const porSemana = {};
  data.forEach(r => {
    const s = getSemana(r['Fecha']); if (!s) return;
    if (!porSemana[s]) porSemana[s] = {verde:0,total:0};
    porSemana[s].verde += r['Racimos Verdes']; porSemana[s].total += r['Total Racimos'];
  });
  const semanas = Object.keys(porSemana).sort().map(s => ({sem:s, pct:pctOf(porSemana[s].verde,porSemana[s].total)}));
  const tendenciaDelta = semanas.length >= 2 ? Math.round((semanas[0].pct - semanas[semanas.length-1].pct)*10)/10 : 0;

  // ── Ranking por supervisor ──
  const porSup = {};
  data.forEach(r => {
    const s = r['Supervisor']; if (!s) return;
    if (!porSup[s]) porSup[s] = {evals:0, verde:0, total:0, prom:0};
    porSup[s].evals++; porSup[s].verde += r['Racimos Verdes']; porSup[s].total += r['Total Racimos']; porSup[s].prom += (r['Prom Parámetros']||0);
  });
  const rankingSup = Object.keys(porSup).map(s => {
    const d = porSup[s], pct = pctOf(d.verde,d.total), prom = d.prom/d.evals;
    return { supervisor:s, evals:d.evals, pct, nota: notaGerencia(pct,prom) };
  }).sort((a,b) => a.pct - b.pct);

  // ── Cosecheros que requieren atención (>=2 evaluaciones en C/D) ──
  const porCos = {};
  data.forEach(r => {
    const cod = r['Código']; if (!cod) return;
    if (!porCos[cod]) porCos[cod] = {nombre:r['Cosechero'], supervisor:r['Supervisor'], evals:0, verde:0, total:0, cd:0};
    porCos[cod].evals++; porCos[cod].verde += r['Racimos Verdes']; porCos[cod].total += r['Total Racimos'];
    if (r['Calificación Nota']==='C'||r['Calificación Nota']==='D') porCos[cod].cd++;
  });
  const atencion = Object.values(porCos).filter(c => c.cd >= 2)
    .map(c => ({...c, pct: pctOf(c.verde,c.total)}))
    .sort((a,b) => b.cd/b.evals - a.cd/a.evals).slice(0,8);

  // ── Lotes con peor calidad ──
  const porLote = {};
  data.forEach(r => {
    const l = r['Frente/Lote']; if (!l) return;
    if (!porLote[l]) porLote[l] = {palma:r['Tipo Palma'], supervisor:r['Supervisor'], verde:0, total:0, diasMax:0};
    porLote[l].verde += r['Racimos Verdes']; porLote[l].total += r['Total Racimos'];
    porLote[l].diasMax = Math.max(porLote[l].diasMax, r['Días Ciclo']||0);
  });
  const rankingLotes = Object.keys(porLote).map(l => {
    const d = porLote[l], pct = pctOf(d.verde,d.total);
    return { lote:l, palma:d.palma, supervisor:d.supervisor, pct, nota: notaGerencia(pct,3) };
  }).sort((a,b) => b.pct - a.pct).slice(0,6);

  // ── Lotes fuera de ciclo ──
  const fueraCiclo = lotesFC.map(l => {
    const fila = data.find(r => r['Frente/Lote']===l && r['Fuera Ciclo']==='SÍ');
    return { lote:l, palma:fila['Tipo Palma'], dias:fila['Días Ciclo'], supervisor:fila['Supervisor'] };
  }).sort((a,b) => b.dias - a.dias).slice(0,8);

  // ── Promedio parámetros de calidad ──
  const n = data.length;
  const paramsCalidad = {
    'Fruto suelto en suelo':   sum(data,'Fruto Suelto')/n,
    'Daño mecánico / cortes':  sum(data,'Daño Mecánico')/n,
    'Pedúnculo largo':         sum(data,'Pedúnculo Largo')/n,
    'Disposición de hoja':     sum(data,'Disposición Hoja')/n,
    'Plato enmalezado':        sum(data,'Plato Enmalezado')/n
  };
  const peorParam = Object.keys(paramsCalidad).reduce((a,b) => paramsCalidad[a] < paramsCalidad[b] ? a : b);

  // ── Composición de racimos ──
  const rM = sum(data,'Racimos Maduros'), rS = sum(data,'Racimos Sobremaduro'),
        rV = sum(data,'Racimos Verdes'), rP = sum(data,'Racimos Podridos'), rD = sum(data,'Racimos Dejados');
  const totComp = rM+rS+rV+rP || 1;

  // ── Análisis narrativo automático ──
  let analisis = '';
  if (tendenciaDelta !== 0) {
    analisis += tendenciaDelta > 0
      ? `La calidad de cosecha mejoró durante el periodo: el % de fruta verde bajó ${Math.abs(tendenciaDelta)} puntos porcentuales entre la primera y la última semana, señal de que los correctivos aplicados están funcionando. `
      : `La calidad de cosecha empeoró durante el periodo: el % de fruta verde subió ${Math.abs(tendenciaDelta)} puntos porcentuales entre la primera y la última semana — requiere atención inmediata. `;
  }
  const peorSup = rankingSup[rankingSup.length-1];
  if (peorSup && (peorSup.nota==='C'||peorSup.nota==='D')) {
    analisis += `El equipo de ${peorSup.supervisor} es el que más necesita acompañamiento en calidad de corte (${peorSup.pct}% verde, nota ${peorSup.nota}). Esta nota mide % de fruta verde cosechada, no la cantidad de evaluaciones realizadas — un supervisor puede registrar más evaluaciones y aun así salir aquí si sus cosecheros tienen más fruta verde, eso no refleja menor esfuerzo, sino una oportunidad de reforzar técnica de corte en su frente. `;
  }
  if (atencion.length) {
    const top = atencion[0];
    analisis += `${top.nombre} (${top.supervisor}) es el cosechero con mayor reincidencia en nota C/D (${top.cd} de ${top.evals} evaluaciones) — se recomienda capacitación o acompañamiento directo en campo. `;
  }
  const loteRiesgo = rankingLotes.find(l => fueraCiclo.some(f => f.lote===l.lote));
  if (loteRiesgo) {
    analisis += `El ${loteRiesgo.lote} combina baja calidad (${loteRiesgo.pct}% verde, nota ${loteRiesgo.nota}) con corte atrasado, lo que sugiere un problema de programación de corte y no solo del cosechero. `;
  }
  analisis += `El parámetro más débil de la finca es "${peorParam}" (${paramsCalidad[peorParam].toFixed(1)}/5) — vale la pena revisar si falta personal de mantenimiento en los lotes con peor calificación.`;

  const periodoLabel = fechaIni === fechaFin ? fechaIni : `${fechaIni} al ${fechaFin}`;
  const html = buildHTMLGerencia({
    periodoLabel, totalEvals:data.length, totalLotes:lotes.length, totalSup:supervisores.length,
    pctVerdeGlob, cosecherosCD:cosecherosCD.length, lotesFC:lotesFC.length, pctGPS, tendenciaDelta,
    semanas, rankingSup, atencion, rankingLotes, fueraCiclo, params: paramsCalidad, peorParam,
    racimos:{M:Math.round(rM/totComp*100), S:Math.round(rS/totComp*100), V:Math.round(rV/totComp*100), P:Math.round(rP/totComp*100), D:rD},
    analisis
  });
  const blob = Utilities.newBlob(html,'text/html','reporte.html');
  const pdfBlob = blob.getAs('application/pdf');
  const nombre = `Reporte_Ejecutivo_Gerencia_${fechaIni}_${fechaFin}.pdf`;
  return { ok:true, base64: Utilities.base64Encode(pdfBlob.getBytes()), filename: nombre };
}

function buildHTMLGerencia(d) {
  const tag = (nota,extra) => `<span style="display:inline-block;padding:1px 7px;border-radius:8px;font-size:6.6pt;font-weight:bold;background:${NOTA_COLOR[nota].bg};color:${NOTA_COLOR[nota].col};">${nota}${extra?' · '+extra:''}</span>`;
  const bar = (pct,color) => `<div style="background:#eee;border-radius:4px;height:8px;width:100%;"><div style="height:8px;border-radius:4px;width:${Math.min(100,pct)}%;background:${color};"></div></div>`;

  const filasSup = d.rankingSup.map(s => `<tr><td>${s.supervisor}</td><td>${bar(Math.min(100,s.pct*16),NOTA_COLOR[s.nota].col)}</td><td style="text-align:center;">${tag(s.nota, s.pct+'%')}</td></tr>`).join('');

  const filasSemanas = d.semanas.length
    ? `<tr style="background:#f7f7f7;">${d.semanas.map(s=>`<td>${s.sem}</td>`).join('')}</tr><tr>${d.semanas.map(s=>`<td>${bar(s.pct*16,s.pct<=1?'#3B6D11':s.pct<=3?'#854F0B':s.pct<=5?'#A35B0B':'#A32D2D')}${s.pct}%</td>`).join('')}</tr>`
    : '<tr><td>Sin suficientes datos para mostrar tendencia</td></tr>';

  const filasParams = Object.keys(d.params).map(k => {
    const v = d.params[k]; const nivel = v>=4.5?'Bueno':v>=3.5?'Aceptable':'Débil';
    return `<tr><td style="width:46%;">${k}</td><td>${bar(v/5*100,'#3B6D11')}</td><td style="text-align:right;width:14%;">${v.toFixed(1)} ${nivel}</td></tr>`;
  }).join('');

  const filasAtencion = d.atencion.length ? d.atencion.map(c => {
    const nota = c.cd >= c.evals*0.6 ? 'D' : 'C';
    return `<tr><td>${c.nombre}</td><td>${c.supervisor}</td><td style="text-align:center;">${c.pct}%</td><td style="text-align:center;">${tag(nota)}</td><td style="text-align:center;">${c.cd}/${c.evals}</td></tr>`;
  }).join('') : '<tr><td colspan="5" style="color:#999;">Ningún cosechero con reincidencia en C/D</td></tr>';

  const filasLotes = d.rankingLotes.map(l => `<tr><td>${l.lote}</td><td>${l.palma}</td><td style="text-align:center;">${l.pct}%</td><td style="text-align:center;">${tag(l.nota)}</td><td>${l.supervisor}</td></tr>`).join('');

  const filasFC = d.fueraCiclo.length ? d.fueraCiclo.map(f => `<tr><td>${f.lote}</td><td>${f.palma}</td><td style="text-align:center;color:#A32D2D;">${f.dias}</td><td style="text-align:center;">${CICLO_MAX_GER[f.palma]||'-'}</td><td>${f.supervisor}</td></tr>`).join('') : '<tr><td colspan="5" style="color:#999;">Ningún lote fuera de ciclo</td></tr>';

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
*{box-sizing:border-box;}
body{font-family:Arial,sans-serif;font-size:7.6pt;padding:10px;color:#1f2924;}
h2{font-size:7.6pt;font-weight:bold;color:#2d6011;text-transform:uppercase;letter-spacing:.4px;border-bottom:1.5px solid #2d6011;padding-bottom:2px;margin:0 0 5px;}
table{width:100%;border-collapse:collapse;}
td,th{padding:2px 4px;vertical-align:middle;}
.kpi{border-radius:6px;text-align:center;padding:7px 4px;}
.box{border:1px solid #e3e3e3;border-radius:6px;padding:7px 8px;margin-bottom:7px;background:#fff;}
.col{vertical-align:top;padding:0 5px;}
</style></head><body>

<table style="background:linear-gradient(135deg,#2d6011,#3B8a1d);border-radius:7px;margin-bottom:9px;"><tr><td style="padding:10px 14px;color:#fff;">
  <table><tr>
    <td style="width:60%;"><div style="font-size:15pt;font-weight:bold;letter-spacing:.5px;">PALMA GRANDE S.A.S</div>
      <div style="font-size:8.5pt;opacity:.92;">Reporte Ejecutivo de Calidad de Cosecha</div></td>
    <td style="text-align:right;font-size:7.2pt;opacity:.85;">Periodo: ${d.periodoLabel}<br>${d.totalSup} supervisores · ${d.totalLotes} lotes · ${d.totalEvals} evaluaciones<br>${d.pctGPS}% con ubicación GPS verificada</td>
  </tr></table>
</td></tr></table>

<table style="margin-bottom:8px;"><tr>
  <td style="width:25%;padding:2px;"><table class="kpi" style="background:#EAF3DE;"><tr><td><div style="font-size:15pt;font-weight:bold;color:#3B6D11;">${d.pctVerdeGlob}%</div><div style="font-size:6.6pt;color:#3B6D11;">% Verde global</div></td></tr></table></td>
  <td style="width:25%;padding:2px;"><table class="kpi" style="background:#FAEEDA;"><tr><td><div style="font-size:15pt;font-weight:bold;color:#854F0B;">${d.cosecherosCD}</div><div style="font-size:6.6pt;color:#854F0B;">Cosecheros C/D</div></td></tr></table></td>
  <td style="width:25%;padding:2px;"><table class="kpi" style="background:#FCEBEB;"><tr><td><div style="font-size:15pt;font-weight:bold;color:#A32D2D;">${d.lotesFC}</div><div style="font-size:6.6pt;color:#A32D2D;">Lotes fuera de ciclo</div></td></tr></table></td>
  <td style="width:25%;padding:2px;"><table class="kpi" style="background:#EEEDFE;"><tr><td><div style="font-size:15pt;font-weight:bold;color:#534AB7;">${d.tendenciaDelta>0?'&#9660; ':d.tendenciaDelta<0?'&#9650; ':''}${Math.abs(d.tendenciaDelta)}pp</div><div style="font-size:6.6pt;color:#534AB7;">${d.tendenciaDelta>=0?'Mejora':'Empeoró'} en el periodo</div></td></tr></table></td>
</tr></table>

<div class="box" style="background:#F4F8F0;border-color:#cfe2bd;">
  <h2>Análisis y recomendaciones</h2>
  <div style="font-size:7.4pt;line-height:1.45;color:#2c3a2c;">${d.analisis}</div>
</div>

<table><tr>
<td class="col" style="width:50%;">

  <div class="box"><h2>Criterios de calificación</h2><table>
    <tr style="background:#f7f7f7;"><th style="text-align:left;">Nota</th><th style="text-align:left;">Significado</th><th style="text-align:left;">Acción</th></tr>
    <tr><td>${tag('A')}</td><td>Excelente / muy bueno</td><td>Mantener estándar</td></tr>
    <tr><td>${tag('B')}</td><td>Aceptable</td><td>Reforzar en próxima visita</td></tr>
    <tr><td>${tag('C')}</td><td>Deficiente</td><td>Llamado de atención + seguimiento</td></tr>
    <tr><td>${tag('D')}</td><td>Crítico</td><td>Capacitación obligatoria</td></tr>
  </table>
  <div style="font-size:6.4pt;color:#888;margin-top:4px;">Nota = 70% calidad de racimos (% verde) + 30% calidad de parámetros de campo.</div>
  </div>

  <div class="box"><h2>Calidad de cosecha por supervisor</h2><table>
    <tr style="background:#f7f7f7;"><th style="text-align:left;">Supervisor</th><th style="text-align:left;width:40%;">% Verde</th><th>Nota</th></tr>
    ${filasSup}
  </table>
  <div style="font-size:6.4pt;color:#888;margin-top:4px;">Ordenado por % de fruta verde, no por cantidad de evaluaciones realizadas.</div>
  </div>

  <div class="box"><h2>Tendencia semanal % verde global</h2><table>${filasSemanas}</table></div>

  <div class="box"><h2>Promedio parámetros de calidad (finca)</h2><table>${filasParams}</table></div>

</td>
<td class="col" style="width:50%;">

  <div class="box"><h2>Cosecheros que requieren atención (C/D recurrente)</h2><table>
    <tr style="background:#f7f7f7;"><th style="text-align:left;">Cosechero</th><th style="text-align:left;">Supervisor</th><th>%V</th><th>Nota</th><th>C/D</th></tr>
    ${filasAtencion}
  </table></div>

  <div class="box"><h2>Lotes con peor calidad de cosecha</h2><table>
    <tr style="background:#f7f7f7;"><th style="text-align:left;">Lote</th><th style="text-align:left;">Palma</th><th>%V</th><th>Nota</th><th style="text-align:left;">Supervisor</th></tr>
    ${filasLotes}
  </table></div>

  <div class="box"><h2>Lotes fuera de ciclo (corte atrasado)</h2><table>
    <tr style="background:#f7f7f7;"><th style="text-align:left;">Lote</th><th style="text-align:left;">Palma</th><th>Días</th><th>Máx.</th><th style="text-align:left;">Supervisor</th></tr>
    ${filasFC}
  </table></div>

  <div class="box"><h2>Composición de racimos (toda la finca)</h2><table><tr>
    <td style="width:20%;padding:2px;"><table style="background:#EAF3DE;border-radius:4px;"><tr><td style="text-align:center;padding:4px;"><div style="font-size:10pt;font-weight:bold;color:#3B6D11;">${d.racimos.M}%</div><div style="font-size:6.2pt;color:#3B6D11;">Maduro</div></td></tr></table></td>
    <td style="width:20%;padding:2px;"><table style="background:#FAEEDA;border-radius:4px;"><tr><td style="text-align:center;padding:4px;"><div style="font-size:10pt;font-weight:bold;color:#854F0B;">${d.racimos.S}%</div><div style="font-size:6.2pt;color:#854F0B;">Sobremad.</div></td></tr></table></td>
    <td style="width:20%;padding:2px;"><table style="background:#EEEDFE;border-radius:4px;"><tr><td style="text-align:center;padding:4px;"><div style="font-size:10pt;font-weight:bold;color:#534AB7;">${d.racimos.V}%</div><div style="font-size:6.2pt;color:#534AB7;">Verde</div></td></tr></table></td>
    <td style="width:20%;padding:2px;"><table style="background:#FCEBEB;border-radius:4px;"><tr><td style="text-align:center;padding:4px;"><div style="font-size:10pt;font-weight:bold;color:#A32D2D;">${d.racimos.P}%</div><div style="font-size:6.2pt;color:#A32D2D;">Podrido</div></td></tr></table></td>
    <td style="width:20%;padding:2px;"><table style="background:#F0EBF8;border-radius:4px;"><tr><td style="text-align:center;padding:4px;"><div style="font-size:10pt;font-weight:bold;color:#5B2D9E;">${d.racimos.D}</div><div style="font-size:6.2pt;color:#5B2D9E;">Dejados</div></td></tr></table></td>
  </tr></table></div>

</td>
</tr></table>

<table style="margin-top:3px;border-top:1px solid #ddd;"><tr>
  <td style="font-size:6.2pt;color:#999;">Palma Grande S.A.S — Supervisión de cosecha</td>
  <td style="font-size:6.2pt;color:#999;text-align:right;">Generado: ${Utilities.formatDate(new Date(),'America/Bogota','dd/MM/yyyy HH:mm')}</td>
</tr></table>

</body></html>`;
}

const CICLO_MAX_GER = { 'guineensis':13, 'hibrido':22 };
