/**
 * Google Apps Script para sincronizar la Lista de Compra del Campamento (Versión Desglosada).
 * Pegar este código en "Extensiones > Apps Script" en tu hoja de cálculo de Google.
 * Después, implementarlo como "Aplicación web" (Deploy > New deployment > Web app).
 * Configurar el acceso como: "Cualquiera" (Anyone).
 */

function doGet(e) {
  var action = e.parameter.action;
  
  if (action === 'update_purchase') {
    return updatePurchase(e.parameter.material, e.parameter.status);
  }
  
  if (action === 'update_box') {
    return updateBox(e.parameter.material, e.parameter.box, e.parameter.quantity);
  }
  
  if (action === 'add_material') {
    return addMaterial(
      e.parameter.material,
      e.parameter.categoria,
      e.parameter.castores,
      e.parameter.lobatos,
      e.parameter.exploradores,
      e.parameter.pioneros,
      e.parameter.rutas,
      e.parameter.tipo,
      e.parameter.detalle
    );
  }
  
  // Por defecto, leer todos los datos
  return readAllData();
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = e.parameter;
  }
  
  if (data.action === 'update_purchase') {
    return updatePurchase(data.material, data.status);
  }
  
  if (data.action === 'update_box') {
    return updateBox(data.material, data.box, data.quantity);
  }
  
  if (data.action === 'add_material') {
    return addMaterial(
      data.material,
      data.categoria,
      data.castores,
      data.lobatos,
      data.exploradores,
      data.pioneros,
      data.rutas,
      data.tipo,
      data.detalle
    );
  }
  
  return readAllData();
}

// Retornar JSON con CORS habilitado
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── LEER DATOS ──
function readAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Leer Lista de la Compra
  var sheetLista = ss.getSheetByName("LISTA DE LA COMPRA");
  if (!sheetLista) {
    return jsonResponse({ error: "No se encontró la hoja 'LISTA DE LA COMPRA'" });
  }
  
  var rangeLista = sheetLista.getDataRange();
  var valuesLista = rangeLista.getValues();
  
  var materiales = [];
  for (var i = 1; i < valuesLista.length; i++) {
    var row = valuesLista[i];
    if (!row[0]) continue; // Omitir vacíos
    
    // Mapear el "quien lo pide" sumando qué columnas tienen cantidad > 0
    var quienPide = [];
    if (parseFloat(row[3]) > 0 || (row[3] && row[3].toString().startsWith("="))) quienPide.push("Castores");
    if (parseFloat(row[4]) > 0 || (row[4] && row[4].toString().startsWith("="))) quienPide.push("Lobatos");
    if (parseFloat(row[5]) > 0 || (row[5] && row[5].toString().startsWith("="))) quienPide.push("Exploradores");
    if (parseFloat(row[6]) > 0 || (row[6] && row[6].toString().startsWith("="))) quienPide.push("Pioneros");
    if (parseFloat(row[7]) > 0 || (row[7] && row[7].toString().startsWith("="))) quienPide.push("Rutas");
    
    materiales.push({
      material: row[0],
      categoria: row[1],
      cantidadRecomendada: row[2], // TOTAL (fórmula evaluada por Sheets)
      quien: quienPide.join(", ") || "Revisar",
      castores: row[3],
      lobatos: row[4],
      exploradores: row[5],
      pioneros: row[6],
      rutas: row[7],
      tipo: row[8],
      detalle: row[9],
      resumenCaja: row[10] || "",
      comprado: (row[11] || "").toString().trim().toUpperCase() === "SÍ" || (row[11] || "").toString().trim().toUpperCase() === "COMPRADO"
    });
  }
  
  // 2. Leer Reparto de Cajas
  var sheetCajas = ss.getSheetByName("REPARTO CAJAS");
  var repartoCajas = [];
  if (sheetCajas) {
    var rangeCajas = sheetCajas.getDataRange();
    var valuesCajas = rangeCajas.getValues();
    for (var j = 1; j < valuesCajas.length; j++) {
      var rowC = valuesCajas[j];
      if (!rowC[0]) continue;
      repartoCajas.push({
        material: rowC[0],
        caja: rowC[1],
        cantidad: rowC[2]
      });
    }
  }
  
  return jsonResponse({
    materiales: materiales,
    repartoCajas: repartoCajas
  });
}

// ── MARCAR COMO COMPRADO ──
function updatePurchase(material, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("LISTA DE LA COMPRA");
  if (!sheet) return jsonResponse({ error: "Hoja no encontrada" });
  
  var values = sheet.getDataRange().getValues();
  var found = false;
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][0].toString().trim() === material.toString().trim()) {
      var valueToSet = (status === "true" || status === true) ? "SÍ" : "";
      sheet.getRange(i + 1, 12).setValue(valueToSet); // Columna L (12)
      found = true;
      break;
    }
  }
  
  return jsonResponse({ success: found, material: material, comprado: status });
}

// ── REGISTRAR REPARTO EN CAJAS ──
function updateBox(material, boxName, qty) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCajas = ss.getSheetByName("REPARTO CAJAS");
  if (!sheetCajas) {
    sheetCajas = ss.insertSheet("REPARTO CAJAS");
    sheetCajas.appendRow(["MATERIAL", "CAJA", "CANTIDAD REPARTIDA"]);
  }
  
  var qtyVal = parseInt(qty, 10);
  if (isNaN(qtyVal)) qtyVal = 0;
  
  var rangeCajas = sheetCajas.getDataRange();
  var valuesCajas = rangeCajas.getValues();
  var foundRowIdx = -1;
  
  for (var i = 1; i < valuesCajas.length; i++) {
    var rowMat = valuesCajas[i][0].toString().trim();
    var rowBox = valuesCajas[i][1].toString().trim();
    if (rowMat === material.toString().trim() && rowBox === boxName.toString().trim()) {
      foundRowIdx = i + 1;
      break;
    }
  }
  
  if (qtyVal <= 0) {
    if (foundRowIdx !== -1) {
      sheetCajas.deleteRow(foundRowIdx);
    }
  } else {
    if (foundRowIdx !== -1) {
      sheetCajas.getRange(foundRowIdx, 3).setValue(qtyVal);
    } else {
      sheetCajas.appendRow([material, boxName, qtyVal]);
    }
  }
  
  // Re-calcular el resumen de cajas para la hoja principal
  updateCajaResumenPrincipal(ss, material);
  
  return jsonResponse({ success: true, material: material, caja: boxName, cantidad: qtyVal });
}

function updateCajaResumenPrincipal(ss, material) {
  var sheetLista = ss.getSheetByName("LISTA DE LA COMPRA");
  var sheetCajas = ss.getSheetByName("REPARTO CAJAS");
  if (!sheetLista || !sheetCajas) return;
  
  var valuesCajas = sheetCajas.getDataRange().getValues();
  var desgloses = [];
  for (var i = 1; i < valuesCajas.length; i++) {
    var rowMat = valuesCajas[i][0].toString().trim();
    if (rowMat === material.toString().trim()) {
      desgloses.push(valuesCajas[i][1] + " (" + valuesCajas[i][2] + ")");
    }
  }
  var resumenText = desgloses.join(", ");
  
  var valuesLista = sheetLista.getDataRange().getValues();
  for (var j = 1; j < valuesLista.length; j++) {
    if (valuesLista[j][0].toString().trim() === material.toString().trim()) {
      sheetLista.getRange(j + 1, 11).setValue(resumenText); // Columna K (11)
      break;
    }
  }
}

// ── AÑADIR NUEVO MATERIAL A LA LISTA ──
function addMaterial(material, categoria, castores, lobatos, exploradores, pioneros, rutas, tipo, detalle) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("LISTA DE LA COMPRA");
  if (!sheet) return jsonResponse({ error: "Hoja 'LISTA DE LA COMPRA' no encontrada" });
  
  // Comprobar duplicado
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().trim().toLowerCase() === material.toString().trim().toLowerCase()) {
      return jsonResponse({ error: "El material ya existe en la lista" });
    }
  }
  
  // Parsear cantidades
  var c = parseFloat(castores) || 0;
  var l = parseFloat(lobatos) || 0;
  var e = parseFloat(exploradores) || 0;
  var p = parseFloat(pioneros) || 0;
  var r = parseFloat(rutas) || 0;
  
  var nextRow = sheet.getLastRow() + 1;
  // Columna C: fórmula de suma automática de las unidades
  var formulaTotal = "=SUM(D" + nextRow + ":H" + nextRow + ")";
  
  sheet.appendRow([
    material,
    categoria,
    formulaTotal,
    c > 0 ? c : "",
    l > 0 ? l : "",
    e > 0 ? e : "",
    p > 0 ? p : "",
    r > 0 ? r : "",
    tipo,
    detalle,
    "", // Resumen cajas vacío
    ""  // Comprado vacío
  ]);
  
  return jsonResponse({ success: true, material: material });
}
