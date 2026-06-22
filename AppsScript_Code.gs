function summarizeServicios(jsonStr) {
  var s;
  try { s = JSON.parse(jsonStr || "{}") || {}; } catch (e) { return ""; }
  var parts = [];
  if (s.svc_pulido) parts.push("Pulido (" + (s.svc_pulido_base || "?") + ")");
  if (s.svc_platinado) parts.push("Platinado " + (s.svc_platinado_tipo || "") + " (" + (s.svc_platinado_base || "?") + ")");
  return parts.join(" + ");
}

function summarizeBitacora(jsonStr) {
  var arr;
  try { arr = JSON.parse(jsonStr || "[]"); } catch (e) { return ""; }
  if (!Array.isArray(arr)) return "";
  return arr.map(function (e) { return e.date + " " + e.time + " - " + e.text; }).join(" | ");
}

function doGet(e) {
  var sheet = SpreadsheetApp.openById('1H7wlH_q77FMRI1krbkJUSOqzDJqYGi8RB8qZ5PEr0vU').getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var leads = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      leads.push({
        id: rows[i][0], name: rows[i][1], company: rows[i][2],
        email: rows[i][3], phone: String(rows[i][4] || ""), source: rows[i][5],
        status: rows[i][6], notes: rows[i][7], date: rows[i][8],
        puesto: rows[i][9] || "", ciudad: rows[i][10] || "",
        followup: rows[i][11] || "", contactedDate: rows[i][12] || "",
        servicios: rows[i][13] || "", bitacora: rows[i][14] || "",
        archived: String(rows[i][17] || "").toUpperCase() === "TRUE"
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify(leads))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.openById('1H7wlH_q77FMRI1krbkJUSOqzDJqYGi8RB8qZ5PEr0vU').getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  if (data.action === 'add') {
    sheet.appendRow([
      data.id, data.nombre, data.empresa, data.correo,
      data.telefono, data.fuente, data.etapa, data.notas, data.fecha,
      data.puesto || "", data.ciudad || "", data.seguimiento || "",
      data.fechaContactado || "", data.servicios || "", data.bitacora || "",
      summarizeServicios(data.servicios), summarizeBitacora(data.bitacora),
      data.archivado || "FALSE"
    ]);
  }
  if (data.action === 'delete') {
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) { sheet.deleteRow(i + 1); break; }
    }
  }
  if (data.action === 'update') {
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 2, 1, 17).setValues([[
          data.nombre, data.empresa, data.correo, data.telefono, data.fuente,
          data.etapa, data.notas, data.fecha, data.puesto || "", data.ciudad || "",
          data.seguimiento || "", data.fechaContactado || "", data.servicios || "", data.bitacora || "",
          summarizeServicios(data.servicios), summarizeBitacora(data.bitacora),
          data.archivado || "FALSE"
        ]]);
        break;
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Ejecuta esta función UNA SOLA VEZ desde el editor de Apps Script (botón Ejecutar)
// para migrar los leads viejos que tenían los datos extra escondidos en Notas,
// y para llenar las columnas de texto legible (P y Q) en los leads ya existentes.
function migrarNotasViejas() {
  var sheet = SpreadsheetApp.openById('1H7wlH_q77FMRI1krbkJUSOqzDJqYGi8RB8qZ5PEr0vU').getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var notes = String(rows[i][7] || "");
    var servicios = String(rows[i][13] || "");
    var bitacora = String(rows[i][14] || "");
    var idx = notes.indexOf("[[SW:");
    if (idx >= 0) {
      var end = notes.indexOf("]]", idx);
      if (end >= 0) {
        var extra;
        try { extra = JSON.parse(notes.slice(idx + 5, end)); } catch (err) { extra = null; }
        if (extra) {
          var cleanNotes = notes.slice(0, idx).trim();
          servicios = JSON.stringify({
            svc_pulido: extra.svc_pulido || false, svc_pulido_base: extra.svc_pulido_base || "",
            svc_platinado: extra.svc_platinado || false, svc_platinado_tipo: extra.svc_platinado_tipo || "",
            svc_platinado_base: extra.svc_platinado_base || ""
          });
          bitacora = JSON.stringify(extra.log || []);
          sheet.getRange(i + 1, 8).setValue(cleanNotes);
          sheet.getRange(i + 1, 10, 1, 4).setValues([[
            extra.puesto || "", extra.ciudad || "", extra.followup || "", extra.contactedDate || ""
          ]]);
        }
      }
    }
    if (!rows[i][17]) sheet.getRange(i + 1, 18).setValue("FALSE");
    sheet.getRange(i + 1, 14, 1, 4).setValues([[
      servicios, bitacora, summarizeServicios(servicios), summarizeBitacora(bitacora)
    ]]);
  }
}
