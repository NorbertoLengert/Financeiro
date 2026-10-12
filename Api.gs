/**
 * Api.gs
 * Funções chamadas pela tela (google.script.run). Só entram e saem textos, números,
 * listas e objetos simples (datas viajam como 'yyyy-mm-dd').
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Finanças')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function cadastrosParaTela_(cad) {
  return {
    contas: cad.contas, categorias: cad.listaCats,
    naturezas: Object.keys(cad.nat).map(function (k) { return cad.nat[k]; })
  };
}

function api_bootstrap() {
  var cad = lerCadastros_();
  var hoje = hoje_();
  var corte = Fin.addDias(hoje, -60);
  var recentes = lerLancsFinal_(500).filter(function (l) { return l.data >= corte; });
  recentes.sort(function (a, b) { return a.data < b.data ? 1 : a.data > b.data ? -1 : 0; });
  var out = cadastrosParaTela_(cad);
  out.hoje = hoje;
  out.recorrencias = lerRecs_();
  out.recentes = recentes.slice(0, 60);
  out.planilha = ss_().getUrl();
  return out;
}

// ---------- lançamentos ----------
function normalizarLanc_(l, cad) {
  var o = {
    id: l.id || '', data: String(l.data || ''), desc: String(l.desc || '').trim(),
    categoria: String(l.categoria || ''), meio: String(l.meio || ''), destino: String(l.destino || ''),
    previsto: Fin.parseNum(l.previsto), efetivo: Fin.parseNum(l.efetivo),
    conferido: !!l.conferido, recId: l.recId || '', ref: l.ref || ''
  };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(o.data)) throw new Error('Data inválida.');
  var cat = cad.categorias[o.categoria];
  if (!cat) throw new Error('Categoria desconhecida: ' + o.categoria);
  o.natureza = cat.natureza;
  if (!cad.contasMap[o.meio]) throw new Error('Conta desconhecida: ' + o.meio);
  if (o.destino && !cad.contasMap[o.destino]) throw new Error('Conta de destino desconhecida: ' + o.destino);
  if (o.destino && o.destino === o.meio) throw new Error('Origem e destino não podem ser a mesma conta.');
  if (o.natureza === 'Transferência' && !o.destino) throw new Error('Transferência precisa da conta de destino.');
  if (o.previsto == null && o.efetivo == null) throw new Error('Informe o valor.');
  if (!o.desc) o.desc = o.categoria;
  return o;
}

function api_salvarLancamento(l, opts) {
  opts = opts || {};
  return comLock_(function () {
    var cad = lerCadastros_();
    var o = normalizarLanc_(l, cad);
    var futuras = 0;
    if (o.id && opts.aplicarFuturas && o.recId && o.previsto != null) {
      futuras = aplicarPrevistoFuturo_(o);
    }
    gravarLanc_(o, cad.nat);
    return { lancamento: o, futuras: futuras };
  });
}

// muda o valor da recorrência e dos meses seguintes que ainda não foram pagos
function aplicarPrevistoFuturo_(o) {
  var recs = lerRecs_();
  for (var i = 0; i < recs.length; i++) {
    if (recs[i].id === o.recId) {
      recs[i].previsto = o.previsto;
      gravarLinhaPorChave_(ABA.recs, 1, recs[i].id, recParaLinha_(recs[i]));
    }
  }
  var cad = lerCadastros_(), n = 0;
  lerLancs_().forEach(function (x) {
    if (x.recId === o.recId && x.id !== o.id && x.data > o.data && x.efetivo == null) {
      x.previsto = o.previsto;
      gravarLanc_(x, cad.nat); n++;
    }
  });
  return n;
}

// Parcelamento: cria n lançamentos, um por mês, no mesmo dia (ou no último dia, se o mês for menor).
function api_salvarParcelado(l, n, opts) {
  opts = opts || {};
  return comLock_(function () {
    n = Math.floor(Number(n));
    if (!(n >= 2 && n <= 60)) throw new Error('Número de parcelas entre 2 e 60.');
    var cad = lerCadastros_();
    var valor = Fin.parseNum(opts.previsto ? l.previsto : l.efetivo);
    if (valor == null || valor <= 0) throw new Error('Informe o valor.');
    var base = normalizarLanc_({ data: l.data, desc: l.desc, categoria: l.categoria, meio: l.meio, destino: l.destino, efetivo: valor }, cad);
    var conta = cad.contasMap[base.meio];
    var efetivas = opts.previsto ? 'nenhuma' : (conta.classe === 'cartao' ? 'todas' : 'primeira');
    var grupo = novoId_();
    var parcelas = Fin.gerarParcelas({ data: base.data, desc: base.desc, categoria: base.categoria, meio: base.meio, destino: base.destino, valor: valor },
      n, !!opts.valorEhTotal, efetivas, grupo);
    var criados = parcelas.map(function (p) {
      var o = normalizarLanc_(p, cad); o.id = novoId_(); return o;
    });
    var linhas = criados.map(function (o) { return lancParaLinha_(o, cad.nat); });
    var sh = aba_(ABA.lanc);
    sh.getRange(sh.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
    return { lancamentos: criados };
  });
}

function api_buscar(texto) {
  var palavras = Fin.norm(texto).split(' ').filter(function (x) { return x; });
  if (!palavras.length) return [];
  var achados = lerLancs_().filter(function (l) {
    var alvo = Fin.norm(l.desc + ' ' + l.categoria + ' ' + l.meio + ' ' + l.destino);
    return palavras.every(function (p) { return alvo.indexOf(p) >= 0; });
  });
  achados.sort(function (a, b) { return a.data < b.data ? 1 : a.data > b.data ? -1 : 0; });
  return achados.slice(0, 60);
}

// Antecipações: dinheiro adiantado que alguém vai devolver (cartão, viagem pela empresa...)
function api_pendencias() {
  var adiantado = 0, devolvido = 0, nA = 0;
  lerLancs_().forEach(function (l) {
    if (l.categoria === CAT_ANTECIPACAO) { adiantado += Fin.valorUsado(l); nA++; }
    else if (l.categoria === CAT_DEVOLUCAO) devolvido += Fin.valorUsado(l);
  });
  return { antecipacao: { adiantado: Fin.round2(adiantado), devolvido: Fin.round2(devolvido), saldo: Fin.round2(adiantado - devolvido), lancamentos: nA } };
}

function api_confirmarPagamento(id, valor, data) {
  return comLock_(function () {
    var cad = lerCadastros_();
    var l = lerLancs_().filter(function (x) { return x.id === id; })[0];
    if (!l) throw new Error('Pagamento não encontrado.');
    var v = Fin.parseNum(valor);
    l.efetivo = v != null ? v : l.previsto;
    if (data) l.data = data;
    gravarLanc_(l, cad.nat);
    return l;
  });
}

function api_excluirLancamento(id) {
  return comLock_(function () {
    var i = linhaDoId_(id);
    if (i < 0) throw new Error('Lançamento não encontrado.');
    aba_(ABA.lanc).deleteRow(i);
    return true;
  });
}

function api_conferir(ids, valor) {
  return comLock_(function () {
    var sh = aba_(ABA.lanc), n = sh.getLastRow();
    if (n < 2) return 0;
    var col = sh.getRange(2, 1, n - 1, 1).getValues();
    var set = {}; ids.forEach(function (x) { set[x] = 1; });
    var c = 0;
    for (var i = 0; i < col.length; i++) {
      if (set[String(col[i][0])]) { sh.getRange(i + 2, 11).setValue(valor ? 'S' : ''); c++; }
    }
    return c;
  });
}

// ---------- relatórios ----------
function api_extrato(chave, ini, fim) {
  var cad = lerCadastros_();
  var conta = cad.contasMap[chave];
  if (!conta) throw new Error('Conta desconhecida.');
  var r = Fin.extrato(conta, lerLancs_(), lerSaldos_(), cad.nat, ini, fim);
  r.conta = conta;
  return r;
}

function api_resumo(ini, fim) {
  var cad = lerCadastros_();
  return Fin.resumo(lerLancs_(), cad.nat, ini, fim);
}

function api_saldos(data) {
  var cad = lerCadastros_(), lancs = lerLancs_(), saldos = lerSaldos_();
  var d = data || hoje_();
  return cad.contas.filter(function (c) { return c.ativo; }).map(function (c) {
    var a = Fin.ancoraAte(saldos, c.chave, d);
    return {
      chave: c.chave, nome: c.nome, tipo: c.tipo, classe: c.classe,
      real: Fin.saldoEm(c, lancs, saldos, cad.nat, d, 'real'),
      comPrevistos: Fin.saldoEm(c, lancs, saldos, cad.nat, d, 'proj'),
      ancora: a ? { data: a.data, inst: a.inst, diff: a.diff } : null
    };
  });
}

// Registra o saldo informado pela instituição e devolve a diferença contra os lançamentos.
// opts.rendimento: lança a diferença como rendimento (para aplicações).
function api_registrarSaldo(chave, data, inst, opts) {
  opts = opts || {};
  return comLock_(function () {
    var cad = lerCadastros_();
    var conta = cad.contasMap[chave];
    if (!conta) throw new Error('Conta desconhecida.');
    inst = Fin.parseNum(inst);
    if (inst == null) throw new Error('Informe o saldo.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data)) throw new Error('Data inválida.');
    var lancs = lerLancs_();
    var saldosSemEste = lerSaldos_().filter(function (s) { return !(s.conta === chave && s.data === data); });
    var interno = Fin.instParaInterno(conta, inst);
    var calc = Fin.saldoEm(conta, lancs, saldosSemEste, cad.nat, data, 'real');
    var dif = Fin.round2(interno - calc);
    var rendimentoLancado = false;
    if (opts.rendimento && dif !== 0) {
      if (!cad.categorias[CAT_RENDIMENTO]) throw new Error('Falta a categoria "' + CAT_RENDIMENTO + '".');
      var o = normalizarLanc_({
        data: data, desc: 'Rendimento ' + conta.nome, categoria: CAT_RENDIMENTO, meio: chave, efetivo: dif
      }, cad);
      gravarLanc_(o, cad.nat);
      lancs.push(o);
      calc = Fin.saldoEm(conta, lancs, saldosSemEste, cad.nat, data, 'real');
      dif = Fin.round2(interno - calc);
      rendimentoLancado = true;
    }
    var exibeCalc = conta.classe === 'cartao' ? -calc : calc;
    var exibeDif = conta.classe === 'cartao' ? -dif : dif;
    gravarLinhaPorChave_(ABA.saldos, 6, chave + '|' + data, [isoData_(data), chave, exibeCalc, inst, exibeDif, chave + '|' + data]);
    return { calculado: exibeCalc, informado: inst, diferenca: exibeDif, rendimentoLancado: rendimentoLancado };
  });
}

// ---------- cartão ----------
function gruposCartao_(cad) {
  return cad.contas.filter(function (c) { return c.ativo && c.classe === 'cartao' && !c.faturaDe; });
}

function api_faturas() {
  var cad = lerCadastros_(), lancs = lerLancs_(), hoje = hoje_();
  return gruposCartao_(cad).map(function (g) {
    var adicionais = cad.contas.filter(function (c) { return c.faturaDe === g.chave; }).map(function (c) { return c.chave; });
    return {
      chave: g.chave, nome: g.nome, diaFech: g.diaFech, diaVenc: g.diaVenc, contaPgto: g.contaPgto, adicionais: adicionais,
      faturas: Fin.faturas(g, cad.contas, lancs, cad.nat, hoje).slice(0, 8)
    };
  });
}

function api_itensFatura(chaveGrupo, id) {
  var cad = lerCadastros_();
  var g = cad.contasMap[chaveGrupo];
  var membros = cad.contas.filter(function (c) { return c.chave === g.chave || c.faturaDe === g.chave; }).map(function (c) { return c.chave; });
  var itens = lerLancs_().filter(function (l) {
    return membros.indexOf(l.meio) >= 0 && !l.destino && Fin.cicloFatura(l.data, g.diaFech || 1, g.diaVenc || 10).id === id;
  }).map(function (l) {
    return { id: l.id, data: l.data, desc: l.desc, categoria: l.categoria, meio: l.meio, valor: -Fin.sinalDe(cad.nat, l.natureza) * Fin.valorUsado(l), pendente: l.efetivo == null };
  });
  itens.sort(function (a, b) { return a.data < b.data ? -1 : a.data > b.data ? 1 : 0; });
  return itens;
}

function api_gerarPagamentoFatura(chaveGrupo, id) {
  return comLock_(function () {
    var cad = lerCadastros_(), lancs = lerLancs_();
    var g = cad.contasMap[chaveGrupo];
    if (!g) throw new Error('Cartão desconhecido.');
    if (!g.diaFech || !g.diaVenc) throw new Error('Preencha o dia de fechamento e de vencimento do cartão na aba Contas.');
    if (!g.contaPgto || !cad.contasMap[g.contaPgto]) throw new Error('Preencha a conta de pagamento do cartão na aba Contas.');
    var fat = Fin.faturas(g, cad.contas, lancs, cad.nat, hoje_()).filter(function (f) { return f.id === id; })[0];
    if (!fat) throw new Error('Fatura não encontrada.');
    if (fat.status === 'paga') throw new Error('Esta fatura já está paga.');
    if (fat.pagamentos.length && !fat.pagamento) throw new Error('Já existe pagamento desta fatura na conta (' + fat.pagamentos.map(function (p) { return p.data; }).join(', ') + '). Não gerei outro para não duplicar.');
    var base = Fin.pagamentoFatura(g, fat, CAT_CARTAO);
    if (fat.pagamento) base.id = fat.pagamento.id;
    var o = normalizarLanc_(base, cad);
    gravarLanc_(o, cad.nat);
    return o;
  });
}

// ---------- recorrências ----------
function api_gerarMes(ym) {
  return comLock_(function () {
    if (!/^\d{4}-\d{2}$/.test(ym)) throw new Error('Mês inválido.');
    var cad = lerCadastros_();
    var r = Fin.gerarRecorrencias(lerRecs_(), lerLancs_(), ym);
    var sh = aba_(ABA.lanc);
    var linhas = r.novos.map(function (x) {
      var o = normalizarLanc_(x, cad);
      o.id = novoId_();
      return lancParaLinha_(o, cad.nat);
    });
    if (linhas.length) sh.getRange(sh.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
    return { criados: linhas.length, jaExistiam: r.jaExistiam };
  });
}

function api_salvarRecorrencia(r) {
  return comLock_(function () {
    var cad = lerCadastros_();
    var o = {
      id: r.id || novoId_(), desc: String(r.desc || '').trim(), meio: r.meio, categoria: r.categoria, destino: r.destino || '',
      previsto: Fin.parseNum(r.previsto) || 0, dia: Math.min(31, Math.max(1, Number(r.dia) || 1)),
      inicio: r.inicio, n: Math.max(0, Math.floor(Number(r.n) || 0)), ativo: r.ativo !== false
    };
    if (!o.desc) throw new Error('Informe a descrição.');
    if (!cad.contasMap[o.meio]) throw new Error('Conta desconhecida.');
    if (!cad.categorias[o.categoria]) throw new Error('Categoria desconhecida.');
    if (!/^\d{4}-\d{2}$/.test(o.inicio)) throw new Error('Mês de início inválido.');
    gravarLinhaPorChave_(ABA.recs, 1, o.id, recParaLinha_(o));
    return o;
  });
}

function api_excluirRecorrencia(id) {
  return comLock_(function () {
    var sh = aba_(ABA.recs), n = sh.getLastRow();
    if (n < 2) return false;
    var col = sh.getRange(2, 1, n - 1, 1).getValues();
    for (var i = 0; i < col.length; i++) if (String(col[i][0]) === String(id)) { sh.deleteRow(i + 2); return true; }
    return false;
  });
}

// ---------- cadastros ----------
function api_salvarConta(c) {
  return comLock_(function () {
    var cad = lerCadastros_();
    var nome = String(c.nome || '').trim(), tipo = String(c.tipo || '').trim();
    if (!nome || !tipo) throw new Error('Informe tipo e nome.');
    var chave = c.chave || (tipo + ' - ' + nome);
    var atual = cad.contasMap[chave];
    var ordem = atual ? atual.ordem : (cad.contas.reduce(function (m, x) { return Math.max(m, x.ordem); }, 0) + 1);
    var classe = (cad.tipos[tipo] ? cad.tipos[tipo].classe : '') || c.classe || 'outro';
    gravarLinhaPorChave_(ABA.contas, 4, chave, [ordem, nome, tipo, chave, classe, c.ativo === false ? 'N' : 'S',
      c.saldoInicial == null || c.saldoInicial === '' ? '' : Fin.parseNum(c.saldoInicial), isoData_(c.dataSaldoInicial || ''),
      c.faturaDe || '', c.diaFech || '', c.diaVenc || '', c.contaPgto || '']);
    return chave;
  });
}

function api_salvarCategoria(c) {
  return comLock_(function () {
    var cad = lerCadastros_();
    var desc = String(c.descricao || '').trim();
    if (!desc) throw new Error('Informe o nome.');
    if (!cad.nat[c.natureza]) throw new Error('Natureza inválida.');
    var atual = cad.categorias[desc];
    var ordem = atual ? atual.ordem : (cad.listaCats.reduce(function (m, x) { return Math.max(m, x.ordem); }, 0) + 1);
    gravarLinhaPorChave_(ABA.cats, 2, desc, [ordem, desc, c.natureza, c.ativo === false ? 'N' : 'S']);
    return desc;
  });
}

function api_cadastros() { return cadastrosParaTela_(lerCadastros_()); }
