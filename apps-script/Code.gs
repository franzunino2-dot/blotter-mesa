// ─── WEB APP: doGet devuelve todos los datos sin límite de CSV ───────────────
function doGet(e) {
  const data = getBlotterData();
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getBlotterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const patternGuion = /^\d{1,2}-\d{2}$/;
  const patternBarra = /^\d{1,2}\/\d{2}$/;
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;
  const HB_DESDE = new Date(2026, 3, 17);

  const result = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    const isGuion = patternGuion.test(name);
    const isBarra = patternBarra.test(name);
    if (!isGuion && !isBarra) return;

    const sep = isGuion ? '-' : '/';
    const parts = name.split(sep);
    const dia = parseInt(parts[0]);
    const mes = parseInt(parts[1]);
    const year = mes > currentMonth ? currentYear - 1 : currentYear;
    const fecha = pad2(dia) + '/' + pad2(mes) + '/' + year;
    const fechaDate = new Date(year, mes - 1, dia);

    const data = sheet.getDataRange().getValues();
    if (data.length < 5) return;

    // Leer comisión total bruta (col O fila 3)
    let comisionTotalBruta = 0;
    if (data.length >= 3) {
      for (let c = 0; c < data[2].length - 1; c++) {
        const label = normalizar(String(data[2][c]));
        if (label.includes('comision') || label.includes('comisión')) {
          const raw = data[2][c + 1];
          const val = parseFloat(String(raw).replace(/[^0-9.,-]/g, '').replace(',', '.')) || 0;
          if (val > 0) { comisionTotalBruta = val; break; }
        }
      }
      if (comisionTotalBruta === 0 && data[2].length >= 15) {
        const raw = data[2][14];
        comisionTotalBruta = parseFloat(String(raw).replace(/[^0-9.,-]/g, '').replace(',', '.')) || 0;
      }
    }

    // Buscar header
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
      const row = data[i].map(h => normalizar(String(h)));
      if (row.includes('especie') && row.includes('vn') && row.includes('monto')) {
        headerRow = i; break;
      }
    }
    if (headerRow === -1) return;

    const headers = data[headerRow].map(h => normalizar(String(h)));
    const idx = {
      mercado:   headers.indexOf('mercado'),
      op:        headers.indexOf('operacion'),
      especie:   headers.indexOf('especie'),
      moneda:    headers.indexOf('moneda'),
      precio:    headers.indexOf('precio'),
      vn:        headers.indexOf('vn'),
      monto:     headers.indexOf('monto'),
      comitente: headers.indexOf('comitente'),
      cliente:   headers.indexOf('cliente'),
      productor: headers.indexOf('productor'),
      plazo:     headers.indexOf('plazo'),
      pyl:       headers.findIndex(h => h.includes('pyl')),
    };
    const idxComProductor = headers.findIndex(h =>
      h.includes('productor $') || h.includes('productor$') ||
      (h.includes('productor') && h !== headers[idx.productor])
    );

    // Total día
    let totalDia = 0;
    for (let c = 0; c < data[0].length - 1; c++) {
      if (normalizar(String(data[0][c])) === 'total') {
        const val = parseFloat(String(data[0][c + 1]).replace(',', '.'));
        if (!isNaN(val)) { totalDia = val; break; }
      }
    }

    // Leer HOMEBROKER
    const hbRows = [];
    if (fechaDate >= HB_DESDE) {
      let enHB = false;
      const OTRAS_SECCIONES = ['pase', 'contado inmediato', 't+1', 'contado inmed'];
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const celdas = row.filter(c => c !== '').length;
        const rowJoined = normalizar(row.join(''));
        if (rowJoined.includes('homebroker') && celdas <= 3) {
          enHB = true; continue;
        }
        if (enHB) {
          const esOtraSeccion = OTRAS_SECCIONES.some(s => rowJoined.includes(s)) && celdas <= 3;
          if (esOtraSeccion) { enHB = false; continue; }
        }
        if (!enHB) continue;
        const productor = String(row[12] || '').trim();
        const comTotal = parseFloat(String(row[14] || '0').replace(/[^0-9.,-]/g, '').replace(',', '.')) || 0;
        const pylMesa  = parseFloat(String(row[15] || '0').replace(/[^0-9.,-]/g, '').replace(',', '.')) || 0;
        const comProd  = parseFloat(String(row[16] || '0').replace(/[^0-9.,-]/g, '').replace(',', '.')) || 0;
        if (!comTotal && !pylMesa && !comProd) continue;
        hbRows.push([productor, comTotal, pylMesa, comProd]);
      }
    }

    // Leer operaciones
    const SECCIONES = ['pase', 'contado inmediato', 't+1'];
    let seccionActual = 'pase';
    const rows = [];

    for (let i = headerRow + 1; i < data.length; i++) {
      const row = data[i];
      for (const sec of SECCIONES) {
        if (normalizar(row.join('')).includes(sec) && row.filter(c => c !== '').length <= 3) {
          seccionActual = sec; break;
        }
      }
      const op = String(row[idx.op] || '').trim().toUpperCase();
      if (!op) continue;
      const especie = String(row[idx.especie] || '').trim();
      if (!especie) continue;
      const vnRaw = row[idx.vn];
      const montoRaw = row[idx.monto];
      const vn = parseFloat(String(vnRaw).replace(',', '.')) || 0;
      const monto = parseFloat(String(montoRaw).replace(',', '.')) || 0;
      if (vn === 0 && monto === 0) continue;
      if (String(vnRaw).includes('DIV') || String(montoRaw).includes('DIV')) continue;
      const precioRaw = row[idx.precio];
      const precio = typeof precioRaw === 'number'
        ? Math.round(precioRaw * 10000) / 10000
        : parseFloat(String(precioRaw || '0').replace(/[^0-9.,-]/g, '').replace(',', '.')) || 0;
      const pylMesa = idx.pyl >= 0
        ? (parseFloat(String(row[idx.pyl] || '0').replace(',', '.')) || 0) : 0;
      const comProductor = idxComProductor >= 0
        ? (parseFloat(String(row[idxComProductor] || '0').replace(',', '.')) || 0) : 0;

      rows.push([
        seccionActual,
        String(row[idx.mercado] || '').trim().toUpperCase(),
        String(row[idx.plazo] || '').trim() || (seccionActual === 't+1' ? '1' : '0'),
        op,
        especie.toUpperCase(),
        String(row[idx.moneda] || '').trim(),
        precio, vn, monto,
        String(row[idx.comitente] || '').trim(),
        String(row[idx.cliente] || '').trim(),
        idx.productor >= 0 ? String(row[idx.productor] || '').trim() : '',
        pylMesa, comProductor
      ]);
    }

    if (rows.length > 0 || hbRows.length > 0) {
      result.push({
        fecha,
        ops: rows,
        totalDia,
        hb: hbRows,
        comisionTotalBruta
      });
    }
  });

  // Ordenar por fecha ascendente
  result.sort((a, b) => {
    const [da, ma, ya] = a.fecha.split('/').map(Number);
    const [db, mb, yb] = b.fecha.split('/').map(Number);
    return new Date(ya, ma-1, da) - new Date(yb, mb-1, db);
  });

  return result;
}

// Mantener getBlotterJSON para compatibilidad (escribe en _blotter_json)
function getBlotterJSON() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = getBlotterData();

  let outputSheet = ss.getSheetByName('_blotter_json');
  if (!outputSheet) outputSheet = ss.insertSheet('_blotter_json');
  outputSheet.clearContents();
  outputSheet.getRange('A1').setValue('fecha');
  outputSheet.getRange('B1').setValue('json');
  outputSheet.getRange('C1').setValue('totalDia');
  outputSheet.getRange('D1').setValue('actualizado');
  outputSheet.getRange('E1').setValue('homebroker');
  outputSheet.getRange('F1').setValue('comisionTotalBruta');

  data.forEach((d, i) => {
    const fila = i + 2;
    outputSheet.getRange(fila, 1).setValue(d.fecha);
    outputSheet.getRange(fila, 2).setValue(JSON.stringify(d.ops));
    outputSheet.getRange(fila, 3).setValue(d.totalDia);
    outputSheet.getRange(fila, 4).setValue(new Date().toLocaleString('es-AR'));
    outputSheet.getRange(fila, 5).setValue(JSON.stringify(d.hb));
    outputSheet.getRange(fila, 6).setValue(d.comisionTotalBruta);
  });
}

function pad2(n) { return String(n).padStart(2, '0'); }

function normalizar(s) {
  return s.trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function onEdit(e) {
  const sheetName = e.source.getActiveSheet().getName();
  if (/^\d{1,2}[-\/]\d{2}$/.test(sheetName)) {
    getBlotterJSON();
  }
}
