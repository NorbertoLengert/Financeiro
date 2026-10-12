/**
 * Data.gs
 * Leitura e escrita da planilha. Tudo que é "ler aba" e "gravar linha" fica aqui.
 */
var ABA = {
  lanc: 'Lancamentos', contas: 'Contas', cats: 'Categorias', nat: 'Naturezas',
  tipos: 'Tipos', saldos: 'Saldos', recs: 'Recorrencias', depara: 'DePara'
};

var CABECALHO = {
  Lancamentos: ['ID', 'Data', 'Descrição', 'Categoria', 'Meio Pagto', 'Natureza', 'Sinal', 'Destino', 'R$ Previsto', 'R$ Efetivo', 'Conferido', 'Recorrência', 'Ref', 'Ano', 'Mês'],
  Contas: ['Ordem', 'Nome', 'Tipo', 'Chave', 'Classe', 'Ativo', 'Saldo inicial', 'Data saldo inicial', 'Fatura de', 'Dia fechamento', 'Dia vencimento', 'Conta de pagamento'],
  Categorias: ['Ordem', 'Descrição', 'Natureza', 'Ativo'],
  Naturezas: ['Natureza', 'Sinal', 'Efeito'],
  Tipos: ['Código', 'Classe', 'Descrição'],
  Saldos: ['Data', 'Conta', 'Saldo Lancamentos', 'Saldo Instituicao', 'Diferenca', 'Chave'],
  Recorrencias: ['ID', 'Descrição', 'Meio Pagto', 'Categoria', 'Destino', 'R$ Previsto', 'Dia', 'Início', 'Nº Meses', 'Ativo'],
  DePara: ['Tipo', 'Valor no legado', 'Chave no app']
};

// ---------- planilha e conversões ----------
// Cada chamada ao serviço da planilha custa de dezenas a centenas de milissegundos.
// Por isso guardamos a planilha, as abas e o cadastro durante a execução.
var _memo = { ss: null, abas: {}, cad: null };
function ss_() {
  if (_memo.ss) return _memo.ss;
  var s = SpreadsheetApp.getActiveSpreadsheet();
  if (!s) {
    var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!id) throw new Error('Script não está ligado a uma planilha.');
    s = SpreadsheetApp.openById(id);
  }
  _memo.ss = s;
  return s;
}
function aba_(nome) {
  if (_memo.abas[nome]) return _memo.abas[nome];
  var sh = ss_().getSheetByName(nome);
  if (!sh) throw new Error('A aba "' + nome + '" não existe. Rode o menu Finanças > Configurar planilha.');
  _memo.abas[nome] = sh;
  return sh;
}
function tz_() { return ss_().getSpreadsheetTimeZone(); }
// O fuso do script é o mesmo da planilha (appsscript.json), então basta ler a data em JavaScript puro.
// Antes isso chamava o serviço da planilha uma vez por linha, o que deixava tudo lento.
function dataIso_(v) {
  if (v instanceof Date) return v.getFullYear() + '-' + Fin.pad(v.getMonth() + 1) + '-' + Fin.pad(v.getDate());
  return v == null ? '' : String(v);
}
function mesIso_(v) {
  if (v instanceof Date) return v.getFullYear() + '-' + Fin.pad(v.getMonth() + 1);
  var s = String(v || '');
  return /^\d{4}-\d{2}/.test(s) ? s.slice(0, 7) : '';
}
function isoData_(s) {
  if (!s) return '';
  var p = Fin.parseIso(s);
  return new Date(p.y, p.m - 1, p.d, 12, 0, 0);
}
function isoMes_(ym) {
  if (!ym) return '';
  return new Date(+ym.slice(0, 4), +ym.slice(5, 7) - 1, 1, 12, 0, 0);
}
function hoje_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function numOuNulo_(v) { return (v === '' || v == null) ? null : Fin.parseNum(v); }
function vazio_(v) { return v == null ? '' : v; }
function sim_(v) { return v === true || String(v).toUpperCase() === 'S' || String(v).toUpperCase() === 'TRUE'; }
function ativo_(v) { return v === '' || v == null || sim_(v); }

function lerAba_(nome) {
  var sh = aba_(nome), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, CABECALHO[nome].length).getValues();
}

function comLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function novoId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

// ---------- cadastros ----------
var CAD_CHAVE = 'cadastros_v3';
function limparCache_() {
  _memo.cad = null;
  try { CacheService.getScriptCache().remove(CAD_CHAVE); } catch (e) {}
}
// Cadastros mudam pouco: ficam 10 minutos no cache do Google. Qualquer gravação feita pelo app,
// pelos importadores ou à mão na planilha (gatilho onEdit) limpa o cache.
function lerCadastros_() {
  if (_memo.cad) return _memo.cad;
  try {
    var c = CacheService.getScriptCache().get(CAD_CHAVE);
    if (c) { _memo.cad = JSON.parse(c); return _memo.cad; }
  } catch (e) {}
  var cad = lerCadastrosPlanilha_();
  try { CacheService.getScriptCache().put(CAD_CHAVE, JSON.stringify(cad), 600); } catch (e) {}
  _memo.cad = cad;
  return cad;
}
function lerCadastrosPlanilha_() {
  var nat = {};
  lerAba_(ABA.nat).forEach(function (r) {
    if (!r[0]) return;
    var s = String(r[1]).trim();
    nat[r[0]] = { nome: r[0], sinal: (s === '+' || s === '1') ? 1 : -1, efeito: String(r[2] || 'neutro').trim() };
  });
  var tipos = {};
  lerAba_(ABA.tipos).forEach(function (r) { if (r[0]) tipos[r[0]] = { codigo: r[0], classe: r[1], desc: r[2] }; });

  var contas = [];
  lerAba_(ABA.contas).forEach(function (r) {
    if (!r[1] && !r[3]) return;
    var tipo = String(r[2] || '').trim(), nome = String(r[1] || '').trim();
    contas.push({
      ordem: Number(r[0]) || 0, nome: nome, tipo: tipo,
      chave: String(r[3] || (tipo + ' - ' + nome)).trim(),
      classe: String(r[4] || (tipos[tipo] ? tipos[tipo].classe : 'outro')).trim(),
      ativo: ativo_(r[5]), saldoInicial: numOuNulo_(r[6]), dataSaldoInicial: dataIso_(r[7]),
      faturaDe: String(r[8] || '').trim(), diaFech: Number(r[9]) || 0, diaVenc: Number(r[10]) || 0,
      contaPgto: String(r[11] || '').trim()
    });
  });
  contas.sort(function (a, b) { return a.ordem - b.ordem; });
  var contasMap = {};
  contas.forEach(function (c) { contasMap[c.chave] = c; });

  var categorias = {}, listaCats = [];
  lerAba_(ABA.cats).forEach(function (r) {
    if (!r[1]) return;
    var c = { ordem: Number(r[0]) || 0, descricao: String(r[1]).trim(), natureza: String(r[2]).trim(), ativo: ativo_(r[3]) };
    categorias[c.descricao] = c; listaCats.push(c);
  });
  var deParaMeio = {}, deParaFatura = [], aliasCat = {}, aliasNat = {};
  lerAba_(ABA.depara).forEach(function (r) {
    if (r[0] === 'meio') deParaMeio[String(r[1]).trim()] = String(r[2]).trim();
    else if (r[0] === 'fatura') deParaFatura.push([String(r[1]), String(r[2]).trim()]);
    else if (r[0] === 'categoria') aliasCat[Fin.norm(r[1])] = String(r[2]).trim();
    else if (r[0] === 'natureza') aliasNat[Fin.norm(r[1])] = String(r[2]).trim();
  });
  return { nat: nat, tipos: tipos, contas: contas, contasMap: contasMap, categorias: categorias, listaCats: listaCats, deParaMeio: deParaMeio, deParaFatura: deParaFatura, aliasCat: aliasCat, aliasNat: aliasNat };
}

// ---------- lançamentos ----------
function lancDeLinha_(r) {
  return {
    id: String(r[0]), data: dataIso_(r[1]), desc: String(r[2]), categoria: String(r[3]), meio: String(r[4]),
    natureza: String(r[5]), destino: String(r[7] || ''), previsto: numOuNulo_(r[8]), efetivo: numOuNulo_(r[9]),
    conferido: sim_(r[10]), recId: String(r[11] || ''), ref: String(r[12] || '')
  };
}
function lancParaLinha_(l, nat) {
  var s = Fin.sinalDe(nat, l.natureza);
  var ano = +l.data.slice(0, 4), mes = +l.data.slice(5, 7);
  return [l.id, isoData_(l.data), l.desc, l.categoria, l.meio, l.natureza, s > 0 ? '+' : '-', l.destino || '',
    l.previsto == null ? '' : l.previsto, l.efetivo == null ? '' : l.efetivo, l.conferido ? 'S' : '',
    l.recId || '', l.ref || '', ano, mes];
}
function lerLancs_() {
  var out = [];
  lerAba_(ABA.lanc).forEach(function (r) { if (r[0] !== '' && r[1] !== '') out.push(lancDeLinha_(r)); });
  return out;
}
// Lê só as últimas linhas (os lançamentos novos entram no fim da aba). Bem mais rápido no celular.
function lerLancsFinal_(n) {
  var sh = aba_(ABA.lanc), ult = sh.getLastRow();
  if (ult < 2) return [];
  var ini = Math.max(2, ult - n + 1);
  var out = [];
  sh.getRange(ini, 1, ult - ini + 1, CABECALHO.Lancamentos.length).getValues().forEach(function (r) {
    if (r[0] !== '' && r[1] !== '') out.push(lancDeLinha_(r));
  });
  return out;
}
function linhaDoId_(id) {
  var sh = aba_(ABA.lanc), n = sh.getLastRow();
  if (n < 2) return -1;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function gravarLanc_(l, nat) {
  var sh = aba_(ABA.lanc);
  var linha = lancParaLinha_(l, nat);
  var i = l.id ? linhaDoId_(l.id) : -1;
  if (i > 0) { sh.getRange(i, 1, 1, linha.length).setValues([linha]); return l; }
  if (!l.id) { l.id = novoId_(); linha[0] = l.id; }
  sh.getRange(sh.getLastRow() + 1, 1, 1, linha.length).setValues([linha]);
  return l;
}

// ---------- saldos ----------
function lerSaldos_() {
  var out = [];
  lerAba_(ABA.saldos).forEach(function (r) {
    if (!r[0] || !r[1]) return;
    out.push({ data: dataIso_(r[0]), conta: String(r[1]).trim(), calc: numOuNulo_(r[2]), inst: numOuNulo_(r[3]), diff: numOuNulo_(r[4]), chave: String(r[5] || '') });
  });
  return out;
}

// ---------- recorrências ----------
function lerRecs_() {
  var out = [];
  lerAba_(ABA.recs).forEach(function (r) {
    if (!r[0]) return;
    out.push({
      id: String(r[0]), desc: String(r[1]), meio: String(r[2]), categoria: String(r[3]), destino: String(r[4] || ''),
      previsto: numOuNulo_(r[5]) || 0, dia: Number(r[6]) || 1, inicio: mesIso_(r[7]), n: Number(r[8]) || 0, ativo: ativo_(r[9])
    });
  });
  return out;
}
function recParaLinha_(r) {
  return [r.id, r.desc, r.meio, r.categoria, r.destino || '', r.previsto || 0, r.dia, isoMes_(r.inicio), r.n, r.ativo ? 'S' : 'N'];
}
function gravarLinhaPorChave_(nome, chaveCol, valor, linha) {
  var sh = aba_(nome), n = sh.getLastRow();
  if (n >= 2) {
    var col = sh.getRange(2, chaveCol, n - 1, 1).getValues();
    for (var i = 0; i < col.length; i++) {
      if (String(col[i][0]) === String(valor)) { sh.getRange(i + 2, 1, 1, linha.length).setValues([linha]); return; }
    }
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, linha.length).setValues([linha]);
}
