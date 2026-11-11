/**
 * Backend das anotações da Trilha Salário-Maternidade.
 * Fica dentro de uma Planilha Google (Extensões > Apps Script).
 * Abas criadas sozinhas: "Notas" e "Reacoes".
 */
var NOTAS = ['id','tipo','secao','trecho','texto','autor','autorNome','criadoEm','resolvida','resposta','respAutor'];
var REAC = ['id','item','emoji','autor','em'];

function aba_(nome, cols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nome);
  if (!sh) { sh = ss.insertSheet(nome); sh.appendRow(cols); sh.setFrozenRows(1); }
  return sh;
}
function ler_(nome, cols) {
  var v = aba_(nome, cols).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    var o = {};
    cols.forEach(function (c, j) { o[c] = v[i][j]; });
    out.push(o);
  }
  return out;
}
function linha_(sh, id) {
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return 0;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function txt_(s, max) { return String(s == null ? '' : s).slice(0, max); }

function doGet() {
  var notas = ler_('Notas', NOTAS).map(function (n) { n.resolvida = n.resolvida === true || n.resolvida === 'TRUE'; n.criadoEm = Number(n.criadoEm); return n; });
  notas.sort(function (a, b) { return b.criadoEm - a.criadoEm; });
  return json_({ ok: true, notas: notas, reacoes: ler_('Reacoes', REAC) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var b = JSON.parse(e.postData.contents);
    var sh, r;
    if (b.acao === 'add') {
      sh = aba_('Notas', NOTAS);
      var id = Utilities.getUuid();
      var n = b.nota || {};
      sh.appendRow([id, n.tipo === 'duvida' ? 'duvida' : 'anotacao', txt_(n.secao, 40), txt_(n.trecho, 600), txt_(n.texto, 4000),
        txt_(n.autor, 80), txt_(n.autorNome, 80), Date.now(), false, '', '']);
      return json_({ ok: true, id: id });
    }
    if (b.acao === 'update') {
      sh = aba_('Notas', NOTAS); r = linha_(sh, b.id);
      if (!r) return json_({ ok: false, erro: 'nao_encontrada' });
      ['resolvida', 'resposta', 'respAutor'].forEach(function (c) {
        if (c in b.patch) sh.getRange(r, NOTAS.indexOf(c) + 1).setValue(c === 'resolvida' ? !!b.patch[c] : txt_(b.patch[c], 4000));
      });
      return json_({ ok: true });
    }
    if (b.acao === 'delete') {
      sh = aba_('Notas', NOTAS); r = linha_(sh, b.id);
      if (r) sh.deleteRow(r);
      return json_({ ok: true });
    }
    if (b.acao === 'reacao') {
      sh = aba_('Reacoes', REAC);
      var rid = txt_(b.item, 40) + '__' + txt_(b.emoji, 20) + '__' + txt_(b.autor, 80);
      r = linha_(sh, rid);
      if (r) sh.deleteRow(r); else sh.appendRow([rid, txt_(b.item, 40), txt_(b.emoji, 20), txt_(b.autor, 80), Date.now()]);
      return json_({ ok: true });
    }
    return json_({ ok: false, erro: 'acao_desconhecida' });
  } finally {
    lock.releaseLock();
  }
}
