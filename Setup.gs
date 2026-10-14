/**
 * Setup.gs
 * Menu da planilha, criação das abas e importação do histórico.
 */
// Edição manual nas abas de cadastro: limpa o cache para o app enxergar na hora.
function onEdit(e) {
  try {
    var n = e && e.range && e.range.getSheet().getName();
    if (n === 'Contas' || n === 'Categorias' || n === 'Naturezas' || n === 'Tipos' || n === 'DePara') limparCache_();
  } catch (er) {}
}

function atualizarCadastros() {
  limparCache_();
  try { SpreadsheetApp.getUi().alert('Pronto. O aplicativo vai ler os cadastros de novo na próxima abertura.'); } catch (e) {}
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Finanças')
    .addItem('1. Configurar planilha', 'configurarPlanilha')
    .addSeparator()
    .addItem('Importar movimentação (aba Importar)', 'importarMovimentacao')
    .addItem('Importar saldos (aba ImportarSaldos)', 'importarSaldos')
    .addItem('Importar recorrentes (aba ImportarRecorrentes)', 'importarRecorrentes')
    .addSeparator()
    .addItem('Atualizar para a versão 4 (grupos, orçamento, painel)', 'migrarV4')
    .addSeparator()
    .addItem('Fazer backup agora', 'backupAgora')
    .addItem('Ligar backup semanal', 'ligarBackupSemanal')
    .addItem('Desligar backup semanal', 'desligarBackupSemanal')
    .addSeparator()
    .addItem('Atualizar cadastros no aplicativo', 'atualizarCadastros')
    .addItem('Mostrar endereço do aplicativo', 'mostrarEndereco')
    .addToUi();
}

function mostrarEndereco() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? 'Endereço do aplicativo:\n\n' + url : 'Ainda não há implantação. Em Implantar > Nova implantação > App da Web.');
}

// ---------- dados iniciais ----------
var SEED_NATUREZAS = [
  ['Receita', '+', 'receita'], ['Despesa', '-', 'despesa'], ['Aplicação', '-', 'neutro'], ['Resgate', '+', 'neutro'],
  ['Depósito', '+', 'neutro'], ['Retirada', '-', 'neutro'], ['Transitória', '+', 'neutro'], ['Transferência', '-', 'neutro'],
  ['Compromisso', '-', 'neutro'], ['Rendimento', '+', 'rendimento']
];
var SEED_TIPOS = [
  ['CC', 'corrente', 'Conta corrente'], ['CP', 'aplicacao', 'Poupança'], ['SIC', 'corrente', 'Sicoob'], ['CRE', 'corrente', 'Cresol'],
  ['CEF', 'corrente', 'Caixa Econômica'], ['CV', 'cartao', 'Cartão Visa'], ['CM', 'cartao', 'Cartão Mastercard'],
  ['CX', 'dinheiro', 'Dinheiro'], ['AP', 'aplicacao', 'Aplicação financeira'], ['VA', 'beneficio', 'Vale alimentação'],
  ['EU', 'moeda', 'Euro'], ['US', 'moeda', 'Dólar'], ['EM', 'emprestimo', 'Valor emprestado a receber'], ['AR', 'emprestimo', 'A receber (pessoa ou empresa)'], ['Coopera', 'outro', 'Cooperativa']
];
// [nome, tipo, faturaDe]
var SEED_CONTAS = [
  ['2278-0', 'CC'], ['Norberto', 'CV'], ['Gerti', 'CV', 'CV - Norberto'], ['Anna', 'CV', 'CV - Norberto'], ['Maria', 'CV', 'CV - Norberto'], ['Caixa', 'CX'],
  ['3891', 'CP'], ['RDC', 'AP'], ['LCA', 'AP'], ['LCI', 'AP'], ['Dilnei', 'EM'], ['Matias', 'EM'], ['Empresa', 'AR'], ['Euro', 'EU'], ['Dolar', 'US'],
  ['Norberto', 'VA'], ['29906-3', 'SIC'], ['41931-1', 'CRE'], ['575930433-6', 'CEF'], ['123912-0', 'CC'],
  ['Norberto', 'CM'], ['cashback', 'Coopera'], ['Gerti', 'CM', 'CM - Norberto']
];
var SEED_CATEGORIAS = [
  ['Água Prefeitura', 'Despesa'], ['Aluguel Terra', 'Despesa'], ['Antecipação', 'Retirada'], ['Aplicações', 'Aplicação'],
  ['Caixa', 'Transitória'], ['Pagamento de fatura', 'Retirada'], ['Combustível', 'Despesa'], ['Consórcio', 'Compromisso'], ['Dentista', 'Despesa'],
  ['Devolução antecipação', 'Depósito'], ['Mercado, Padaria e Açougue', 'Despesa'], ['Telefone', 'Despesa'], ['Lazer', 'Despesa'],
  ['Energia Elétrica', 'Despesa'], ['Plano de Saúde', 'Despesa'], ['Salário UCEFF', 'Receita'],
  ['Resgate de aplicações', 'Resgate'], ['Rendimento de aplicações', 'Rendimento'], ['Transferência entre contas', 'Transferência']
];
var SEED_DEPARA = [
  ['meio', 'CC', 'CC - 2278-0'], ['meio', 'CV', 'CV - Norberto'], ['meio', 'CM', 'CM - Norberto'],
  ['fatura', 'visa', 'CV - Norberto'], ['fatura', 'mastercard', 'CM - Norberto'], ['fatura', 'mastecard', 'CM - Norberto'], ['fatura', 'master', 'CM - Norberto'],
  ['categoria', 'Cartão', 'Pagamento de fatura'], ['natureza', 'Transitoria', 'Transitória']
];
var CAT_RENDIMENTO = 'Rendimento de aplicações';
var CAT_CARTAO = 'Pagamento de fatura';
var CAT_ANTECIPACAO = 'Antecipação';
var CAT_DEVOLUCAO = 'Devolução antecipação';
// Só importa o que for desta data em diante (os saldos informados começam em jan/2024)
var IMPORT_DESDE = '2024-01-01';

function configurarPlanilha() {
  var ss = ss_();
  Object.keys(CABECALHO).forEach(function (nome) {
    var sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
    var cab = CABECALHO[nome];
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#e6f2ef');
    sh.setFrozenRows(1);
  });
  ['Importar', 'ImportarSaldos', 'ImportarRecorrentes'].forEach(function (n) { if (!ss.getSheetByName(n)) ss.insertSheet(n); });

  semear_(ABA.nat, SEED_NATUREZAS);
  semear_(ABA.tipos, SEED_TIPOS);
  semear_(ABA.cats, SEED_CATEGORIAS.map(function (c, i) { return [i + 1, c[0], c[1], '', grupoSugerido_(c[0], c[1])]; }));
  semear_(ABA.depara, SEED_DEPARA);
  semear_(ABA.contas, SEED_CONTAS.map(function (c, i) {
    return [i + 1, c[0], c[1], c[1] + ' - ' + c[0], '', '', '', '', c[2] || '', '', '', ''];
  }));

  var l = aba_(ABA.lanc);
  l.getRange('B2:B').setNumberFormat('dd/mm/yyyy');
  l.getRange('I2:J').setNumberFormat('#,##0.00');
  aba_(ABA.saldos).getRange('A2:A').setNumberFormat('dd/mm/yyyy');
  aba_(ABA.saldos).getRange('C2:E').setNumberFormat('#,##0.00');
  aba_(ABA.orc).getRange('B2:B').setNumberFormat('#,##0.00');
  aba_(ABA.recs).getRange('H2:H').setNumberFormat('mm/yyyy');
  aba_(ABA.recs).getRange('F2:F').setNumberFormat('#,##0.00');
  aba_(ABA.contas).getRange('H2:H').setNumberFormat('dd/mm/yyyy');

  limparCache_();
  var msg = 'Planilha configurada.\n\nPróximos passos:\n1. Aba Contas: preencha dia de fechamento, vencimento e conta de pagamento dos cartões, e o saldo inicial das contas.\n2. Cole o histórico na aba Importar e use o menu Finanças.\n3. Implante como App da Web.';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

// só preenche a aba se ela ainda está vazia (não sobrescreve o que você já ajustou)
function semear_(nome, linhas) {
  var sh = aba_(nome);
  if (sh.getLastRow() > 1) return;
  sh.getRange(2, 1, linhas.length, linhas[0].length).setValues(linhas);
}

// ---------- importação ----------
function importarMovimentacao() { mostrarRelatorio_('Movimentação', importarMovimentacao_()); }
function importarSaldos() { mostrarRelatorio_('Saldos', importarSaldos_()); }
function importarRecorrentes() { mostrarRelatorio_('Recorrentes', importarRecorrentes_()); }

function mostrarRelatorio_(titulo, rel) {
  var t = 'Importação de ' + titulo + '\n\n' + rel.linhas.join('\n');
  try { SpreadsheetApp.getUi().alert(t); } catch (e) { Logger.log(t); }
}

function cabecalhoIdx_(valores) {
  var idx = {};
  valores[0].forEach(function (h, i) { var k = Fin.norm(h); if (k && idx[k] === undefined) idx[k] = i; });
  return idx;
}
function celulaData_(v) { return v instanceof Date ? dataIso_(v) : v; }
function contar_(mapa, k) { mapa[k] = (mapa[k] || 0) + 1; }

function importarMovimentacao_() {
  return comLock_(function () {
    var src = aba_('Importar');
    var vals = src.getDataRange().getValues();
    if (vals.length < 2) throw new Error('A aba Importar está vazia. Cole as linhas da planilha antiga com o cabeçalho na primeira linha.');
    var ix = cabecalhoIdx_(vals);
    ['data', 'descricao', 'conta', 'meio pagto', 'r$ efetivo', 'r$ previsto', 'categoria'].forEach(function (h) {
      if (ix[h] === undefined) throw new Error('Não achei a coluna "' + h + '" no cabeçalho da aba Importar.');
    });
    var cad = lerCadastros_();
    var ctx = { categorias: {}, contas: cad.contas, nat: cad.nat, deParaMeio: cad.deParaMeio, deParaFatura: cad.deParaFatura,
      aliasCat: cad.aliasCat, aliasNat: cad.aliasNat, catFatura: CAT_CARTAO };
    Object.keys(cad.categorias).forEach(function (k) { ctx.categorias[k] = cad.categorias[k].natureza; });

    var existentes = {};
    lerLancs_().forEach(function (l) { contar_(existentes, Fin.chaveDup(l)); });

    var novas = [], motivos = {}, catsNovas = {}, pend = 0, dups = 0, lidas = 0;
    for (var i = 1; i < vals.length; i++) {
      var r = vals[i];
      if (r.every(function (c) { return c === '' || c == null; })) continue;
      lidas++;
      var cv = ix['mes'] !== undefined ? r[ix['mes']] : '';
      var res = Fin.converterMovimentacao({
        data: celulaData_(r[ix['data']]), mes: cv, desc: r[ix['descricao']], conta: r[ix['conta']],
        meio: r[ix['meio pagto']], efetivo: r[ix['r$ efetivo']], previsto: r[ix['r$ previsto']], categoria: r[ix['categoria']]
      }, ctx);
      if (res.skip) { contar_(motivos, res.skip); continue; }
      if (res.l.data < IMPORT_DESDE) { contar_(motivos, 'anterior a ' + IMPORT_DESDE); continue; }
      if (res.novaCategoria && !catsNovas[res.novaCategoria.descricao]) {
        catsNovas[res.novaCategoria.descricao] = res.novaCategoria.natureza;
        ctx.categorias[res.novaCategoria.descricao] = res.novaCategoria.natureza;
      }
      var k = Fin.chaveDup(res.l);
      if (existentes[k] > 0) { existentes[k]--; dups++; continue; }
      if (res.pend) pend++;
      res.l.id = novoId_();
      novas.push(lancParaLinha_(res.l, cad.nat));
    }
    var nomesCats = Object.keys(catsNovas);
    if (nomesCats.length) {
      var shC = aba_(ABA.cats), base = shC.getLastRow();
      shC.getRange(base + 1, 1, nomesCats.length, 5).setValues(nomesCats.map(function (n, j) { return [base + j, n, catsNovas[n], '', grupoSugerido_(n, catsNovas[n])]; }));
    }
    if (novas.length) {
      var sh = aba_(ABA.lanc);
      sh.getRange(sh.getLastRow() + 1, 1, novas.length, novas[0].length).setValues(novas);
    }
    limparCache_();
    var linhas = ['Linhas lidas: ' + lidas, 'Importadas: ' + novas.length, 'Já existiam (ignoradas): ' + dups];
    Object.keys(motivos).forEach(function (m) { linhas.push('Puladas (' + m + '): ' + motivos[m]); });
    if (nomesCats.length) linhas.push('Categorias novas criadas: ' + nomesCats.join(', '));
    if (pend) linhas.push('Atenção: ' + pend + ' pagamento(s) de cartão sem cartão identificado. Preencha o destino nesses lançamentos ou ajuste a aba DePara.');
    return { importadas: novas.length, duplicadas: dups, puladas: motivos, categoriasNovas: nomesCats, pendencias: pend, linhas: linhas };
  });
}

function importarSaldos_() {
  return comLock_(function () {
    var vals = aba_('ImportarSaldos').getDataRange().getValues();
    if (vals.length < 2) throw new Error('A aba ImportarSaldos está vazia.');
    var ix = cabecalhoIdx_(vals);
    ['data', 'conta', 'saldo instituicao'].forEach(function (h) {
      if (ix[h] === undefined) throw new Error('Não achei a coluna "' + h + '" no cabeçalho de ImportarSaldos.');
    });
    var cad = lerCadastros_();
    var ja = {};
    lerSaldos_().forEach(function (s) { ja[s.data + '|' + s.conta] = 1; });
    var novas = [], rej = {}, dups = 0;
    for (var i = 1; i < vals.length; i++) {
      var r = vals[i];
      if (r.every(function (c) { return c === '' || c == null; })) continue;
      var data = Fin.parseDataLegado(celulaData_(r[ix['data']]), null);
      var conta = String(r[ix['conta']]).trim();
      if (!data) { contar_(rej, 'data inválida'); continue; }
      if (!cad.contasMap[conta]) { contar_(rej, 'conta desconhecida: ' + conta); continue; }
      if (ja[data + '|' + conta]) { dups++; continue; }
      ja[data + '|' + conta] = 1;
      var calc = ix['saldo lancamentos'] !== undefined ? numOuNulo_(r[ix['saldo lancamentos']]) : null;
      var inst = numOuNulo_(r[ix['saldo instituicao']]);
      var dif = ix['diferenca'] !== undefined ? numOuNulo_(r[ix['diferenca']]) : null;
      novas.push([isoData_(data), conta, calc == null ? '' : calc, inst == null ? '' : inst, dif == null ? '' : dif, conta + '|' + data]);
    }
    if (novas.length) {
      var sh = aba_(ABA.saldos);
      sh.getRange(sh.getLastRow() + 1, 1, novas.length, 6).setValues(novas);
    }
    var linhas = ['Importados: ' + novas.length, 'Já existiam: ' + dups];
    Object.keys(rej).forEach(function (m) { linhas.push('Rejeitados (' + m + '): ' + rej[m]); });
    return { importados: novas.length, duplicados: dups, rejeitados: rej, linhas: linhas };
  });
}

function importarRecorrentes_() {
  return comLock_(function () {
    var vals = aba_('ImportarRecorrentes').getDataRange().getValues();
    if (vals.length < 2) throw new Error('A aba ImportarRecorrentes está vazia.');
    var ix = cabecalhoIdx_(vals);
    ['descricao', 'meio pagto', 'r$ previsto', 'conta', 'dia', 'mes inicio', 'no meses'].forEach(function (h) {
      if (ix[h] === undefined) throw new Error('Não achei a coluna "' + h + '" no cabeçalho de ImportarRecorrentes.');
    });
    var cad = lerCadastros_();
    var ano = +hoje_().slice(0, 4);
    var ja = {};
    lerRecs_().forEach(function (r) { ja[Fin.norm(r.desc) + '|' + r.meio] = 1; });
    var novas = [], rej = {}, dups = 0, catsNovas = {};
    for (var i = 1; i < vals.length; i++) {
      var r = vals[i];
      if (r.every(function (c) { return c === '' || c == null; })) continue;
      var desc = String(r[ix['descricao']]).trim(), meio = String(r[ix['meio pagto']]).trim();
      var cat = Fin.alias(cad.aliasCat, String(r[ix['conta']]).trim());
      if (!cad.contasMap[meio]) { contar_(rej, 'meio desconhecido: ' + meio); continue; }
      if (ja[Fin.norm(desc) + '|' + meio]) { dups++; continue; }
      ja[Fin.norm(desc) + '|' + meio] = 1;
      if (!cad.categorias[cat] && !catsNovas[cat]) {
        var nat = Fin.alias(cad.aliasNat, ix['categoria'] !== undefined ? String(r[ix['categoria']]).trim() : 'Despesa');
        if (!cad.nat[nat]) { contar_(rej, 'natureza desconhecida: ' + nat); continue; }
        catsNovas[cat] = nat;
      }
      var mes = Number(r[ix['mes inicio']]) || 1;
      novas.push([novoId_(), desc, meio, cat, '', Fin.parseNum(r[ix['r$ previsto']]) || 0, Number(r[ix['dia']]) || 1,
        isoMes_(ano + '-' + Fin.pad(mes)), Number(r[ix['no meses']]) || 12, 'S']);
    }
    var nomes = Object.keys(catsNovas);
    if (nomes.length) {
      var shC = aba_(ABA.cats), base = shC.getLastRow();
      shC.getRange(base + 1, 1, nomes.length, 5).setValues(nomes.map(function (n, j) { return [base + j, n, catsNovas[n], '', grupoSugerido_(n, catsNovas[n])]; }));
    }
    if (novas.length) {
      var sh = aba_(ABA.recs);
      sh.getRange(sh.getLastRow() + 1, 1, novas.length, 10).setValues(novas);
    }
    limparCache_();
    var linhas = ['Importadas: ' + novas.length, 'Já existiam: ' + dups, 'O mês de início foi lido como mês de ' + ano + '.'];
    Object.keys(rej).forEach(function (m) { linhas.push('Rejeitadas (' + m + '): ' + rej[m]); });
    return { importadas: novas.length, duplicadas: dups, rejeitadas: rej, linhas: linhas };
  });
}

// ---------- grupos de categorias ----------
var GRUPOS_PALAVRAS = [
  ['Moradia', ['agua', 'energia', 'luz', 'aluguel', 'condominio', 'iptu', 'internet', 'telefone', 'celular', 'gas', 'reforma', 'casa', 'manutencao']],
  ['Alimentação', ['mercado', 'padaria', 'acougue', 'restaurante', 'lanche', 'feira', 'ifood', 'supermercado']],
  ['Transporte', ['combustivel', 'gasolina', 'ipva', 'seguro auto', 'pedagio', 'estacionamento', 'uber', 'oficina', 'veiculo', 'carro']],
  ['Saúde', ['saude', 'plano de', 'dentista', 'farmacia', 'medico', 'unimed', 'consulta', 'exame']],
  ['Lazer', ['lazer', 'viagem', 'cinema', 'streaming', 'assinatura', 'passeio']],
  ['Educação', ['escola', 'curso', 'faculdade', 'livro', 'mensalidade']]
];
function grupoSugerido_(nome, natureza) {
  if (natureza === 'Receita') return 'Receitas';
  if (natureza === 'Rendimento') return 'Rendimentos';
  if (natureza !== 'Despesa') return '';
  var n = Fin.norm(nome);
  for (var i = 0; i < GRUPOS_PALAVRAS.length; i++) {
    for (var j = 0; j < GRUPOS_PALAVRAS[i][1].length; j++) if (n.indexOf(GRUPOS_PALAVRAS[i][1][j]) >= 0) return GRUPOS_PALAVRAS[i][0];
  }
  return 'Outros';
}

// ---------- migração para a versão 4 (pode rodar mais de uma vez) ----------
function migrarV4() { mostrarRelatorio_('versão 4', migrarV4_()); }
function migrarV4_() {
  var aviso = '';
  try { fazerBackup_('antes da versão 4'); aviso = 'Backup feito antes de mexer nos dados (pasta "' + BACKUP_PASTA + '").'; }
  catch (e) { aviso = 'ATENÇÃO: não consegui fazer o backup antes (' + e.message + '). A migração seguiu mesmo assim.'; }
  return comLock_(function () {
    var ss = ss_(), linhas = [aviso];
    // 1. abas e cabeçalhos
    ['Orcamento', 'Categorias'].forEach(function (nome) {
      var sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
      var cab = CABECALHO[nome];
      sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#e6f2ef');
      sh.setFrozenRows(1);
      _memo.abas[nome] = sh;
    });
    // 2. naturezas
    var nats = lerAba_(ABA.nat).map(function (r) { return String(r[0]); });
    var novasNat = SEED_NATUREZAS.filter(function (n) { return nats.indexOf(n[0]) < 0; });
    if (novasNat.length) {
      var shN = aba_(ABA.nat);
      shN.getRange(shN.getLastRow() + 1, 1, novasNat.length, 3).setValues(novasNat);
      linhas.push('Naturezas criadas: ' + novasNat.map(function (n) { return n[0]; }).join(', '));
    }
    // 3. tipo AR e conta Empresa
    var tipos = lerAba_(ABA.tipos).map(function (r) { return String(r[0]); });
    if (tipos.indexOf('AR') < 0) {
      var shT = aba_(ABA.tipos);
      shT.getRange(shT.getLastRow() + 1, 1, 1, 3).setValues([['AR', 'emprestimo', 'A receber (pessoa ou empresa)']]);
      linhas.push('Tipo AR (a receber) criado.');
    }
    var contas = lerAba_(ABA.contas);
    if (!contas.some(function (r) { return String(r[3]) === 'AR - Empresa'; })) {
      var maxOrdem = contas.reduce(function (m, r) { return Math.max(m, Number(r[0]) || 0); }, 0);
      var shC = aba_(ABA.contas);
      shC.getRange(shC.getLastRow() + 1, 1, 1, 12).setValues([[maxOrdem + 1, 'Empresa', 'AR', 'AR - Empresa', 'emprestimo', 'S', '', '', '', '', '', '']]);
      linhas.push('Conta "AR - Empresa" criada (adiantamentos que a empresa vai restituir).');
    }
    // 4. categorias: natureza e grupo
    var shCat = aba_(ABA.cats), nCat = shCat.getLastRow(), mudouCat = [], semGrupo = 0;
    if (nCat >= 2) {
      var vals = shCat.getRange(2, 1, nCat - 1, 5).getValues();
      vals.forEach(function (r, i) {
        var nome = String(r[1]).trim(), nat = String(r[2]).trim();
        if (nome === 'Consórcio' && nat === 'Aplicação') { r[2] = 'Compromisso'; mudouCat.push(nome); }
        if (nome === CAT_RENDIMENTO && nat === 'Receita') { r[2] = 'Rendimento'; mudouCat.push(nome); }
        if (nome && !String(r[4] || '').trim()) { r[4] = grupoSugerido_(nome, String(r[2]).trim()); if (r[4]) semGrupo++; }
      });
      shCat.getRange(2, 1, vals.length, 5).setValues(vals);
    }
    if (mudouCat.length) linhas.push('Categorias reclassificadas: ' + mudouCat.join(', '));
    if (semGrupo) linhas.push('Grupos sugeridos para ' + semGrupo + ' categoria(s). Confira na aba Categorias, coluna Grupo.');
    // 5. lançamentos já gravados com a natureza antiga
    var shL = aba_(ABA.lanc), nL = shL.getLastRow(), trocados = 0;
    if (nL >= 2) {
      var cats = shL.getRange(2, 4, nL - 1, 1).getValues(), nt = shL.getRange(2, 6, nL - 1, 1).getValues();
      for (var i = 0; i < nt.length; i++) {
        if (cats[i][0] === 'Consórcio' && nt[i][0] === 'Aplicação') { nt[i][0] = 'Compromisso'; trocados++; }
        else if (cats[i][0] === CAT_RENDIMENTO && nt[i][0] === 'Receita') { nt[i][0] = 'Rendimento'; trocados++; }
      }
      if (trocados) shL.getRange(2, 6, nL - 1, 1).setValues(nt);
    }
    linhas.push('Lançamentos ajustados: ' + trocados);
    limparCache_();
    return { trocados: trocados, linhas: linhas };
  });
}

// ---------- backup ----------
var BACKUP_PASTA = 'Backups Finanças';
var BACKUP_MANTER = 12;

function props_() { return PropertiesService.getScriptProperties(); }
function agora_() {
  var d = new Date();
  return dataIso_(d) + ' ' + Fin.pad(d.getHours()) + ':' + Fin.pad(d.getMinutes());
}

// Copia a planilha inteira para a pasta "Backups Finanças" (ao lado da planilha) e guarda só as últimas cópias.
function fazerBackup_(motivo) {
  var arq = DriveApp.getFileById(ss_().getId());
  var pais = arq.getParents();
  var pai = pais.hasNext() ? pais.next() : DriveApp.getRootFolder();
  var achadas = pai.getFoldersByName(BACKUP_PASTA);
  var pasta = achadas.hasNext() ? achadas.next() : pai.createFolder(BACKUP_PASTA);
  var carimbo = agora_();
  var copia = arq.makeCopy('Backup ' + carimbo + (motivo ? ' (' + motivo + ')' : '') + ' - ' + ss_().getName(), pasta);
  var lista = [], it = pasta.getFiles();
  while (it.hasNext()) { var f = it.next(); if (String(f.getName()).indexOf('Backup ') === 0) lista.push(f); }
  lista.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  lista.slice(BACKUP_MANTER).forEach(function (f) { f.setTrashed(true); });
  props_().setProperty('BACKUP_ULTIMO', carimbo);
  props_().deleteProperty('BACKUP_ERRO');
  return { nome: copia.getName(), url: copia.getUrl(), guardados: Math.min(lista.length, BACKUP_MANTER) };
}

function backupAgora() {
  var msg;
  try { var r = fazerBackup_('manual'); msg = 'Backup feito:\n' + r.nome + '\n\nGuardo as últimas ' + BACKUP_MANTER + ' cópias na pasta "' + BACKUP_PASTA + '".'; }
  catch (e) { msg = 'Não consegui fazer o backup: ' + e.message; }
  try { SpreadsheetApp.getUi().alert(msg); } catch (e2) { Logger.log(msg); }
}

// Chamada pelo gatilho semanal. Se falhar, o app avisa na aba Mais.
function backupAutomatico() {
  try { fazerBackup_('semanal'); }
  catch (e) { props_().setProperty('BACKUP_ERRO', agora_() + ': ' + e.message); }
}

function backupAtivo_() {
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'backupAutomatico'; });
}
function ligarBackupSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'backupAutomatico') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backupAutomatico').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  var msg = 'Backup semanal ligado (domingo, de madrugada).';
  try { fazerBackup_('ao ligar'); msg += '\nJá fiz o primeiro agora.'; } catch (e) { msg += '\nMas o primeiro falhou: ' + e.message; }
  try { SpreadsheetApp.getUi().alert(msg); } catch (e2) { Logger.log(msg); }
}
function desligarBackupSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'backupAutomatico') ScriptApp.deleteTrigger(t); });
  try { SpreadsheetApp.getUi().alert('Backup semanal desligado.'); } catch (e) {}
}
