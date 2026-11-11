/**
 * Enfoc — endpoint de sincronització amb Google Sheets
 *
 * 1. Crea (o obre) un full de càlcul i copia'n l'ID de la URL.
 * 2. Extensions > Apps Script, enganxa aquest codi i posa l'ID a SHEET_ID.
 * 3. Desplega > Nou desplegament > Aplicació web
 *      Executa com: jo
 *      Qui hi té accés: qualsevol
 * 4. Copia la URL /exec i enganxa-la als ajustos de l'app.
 *
 * Cada perfil (nom + PIN) escriu al seu propi full: POM_alex_1234
 */

var SHEET_ID = 'POSA_AQUI_LID_DEL_FULL';
var CAPCALERA = ['id','data','inici','fi','total_min','util_min','descans_min','extra_min','pomodoros','tema','notes','pujat'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action !== 'append') return json({ok:false, error:'acció desconeguda'});

    var sh = fullDelPerfil(body.key);
    var existents = {};
    var n = sh.getLastRow();
    if (n > 1) {
      sh.getRange(2, 1, n - 1, 1).getValues().forEach(function(r){ existents[r[0]] = true; });
    }

    var files = (body.records || [])
      .filter(function(r){ return r && r.id && !existents[r.id]; })
      .map(function(r){
        return [
          r.id, r.data, r.inici, r.fi,
          Math.round(r.total / 60), Math.round(r.util / 60),
          Math.round(r.descans / 60), Math.round(r.extra / 60),
          r.pomos, r.tema || '', r.notes || '', new Date()
        ];
      });

    if (files.length) sh.getRange(sh.getLastRow() + 1, 1, files.length, CAPCALERA.length).setValues(files);
    return json({ok:true, afegides:files.length});
  } catch (err) {
    return json({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    if (e.parameter.action !== 'list') return json({ok:false, error:'acció desconeguda'});
    var sh = fullDelPerfil(e.parameter.key);
    var n = sh.getLastRow();
    if (n < 2) return json({ok:true, records:[]});

    var tz = sh.getParent().getSpreadsheetTimeZone();
    var vals = sh.getRange(2, 1, n - 1, CAPCALERA.length).getValues();
    var recs = vals.map(function(v){
      return {
        id: v[0], data: txt(v[1], tz, 'yyyy-MM-dd'), inici: txt(v[2], tz, 'HH:mm'), fi: txt(v[3], tz, 'HH:mm'),
        total: v[4] * 60, util: v[5] * 60, descans: v[6] * 60, extra: v[7] * 60,
        pomos: v[8], tema: v[9], notes: v[10]
      };
    });
    return json({ok:true, records:recs});
  } catch (err) {
    return json({ok:false, error:String(err)});
  }
}

function fullDelPerfil(key) {
  if (!key || !/^[a-z0-9]+_\d{4}$/.test(key)) throw new Error('perfil no vàlid');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var nom = 'POM_' + key;
  var sh = ss.getSheetByName(nom);
  if (!sh) {
    sh = ss.insertSheet(nom);
    sh.appendRow(CAPCALERA);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CAPCALERA.length).setFontWeight('bold');
    sh.getRange('B:D').setNumberFormat('@');   // data, inici, fi: text pla
  }
  return sh;
}

// Si una cel·la ha estat convertida a data/hora, la tornem al text original.
function txt(v, tz, fmt) {
  return (v instanceof Date) ? Utilities.formatDate(v, tz, fmt) : String(v);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
