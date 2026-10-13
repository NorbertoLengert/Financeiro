/**
 * Logic.gs
 * Funções puras: não tocam na planilha. Datas sempre como texto 'yyyy-mm-dd'
 * e meses como 'yyyy-mm'. Assim dá para testar fora do Google.
 */
var Fin = (function () {

  // ---------- datas ----------
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function iso(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  function parseIso(s) { var p = String(s).split('-'); return { y: +p[0], m: +p[1], d: +p[2] }; }
  function diasNoMes(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
  function addDias(s, n) {
    var p = parseIso(s);
    var dt = new Date(Date.UTC(p.y, p.m - 1, p.d + n));
    return iso(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
  }
  function addMeses(ym, n) {
    var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + n;
    y += Math.floor(m / 12); m = ((m % 12) + 12) % 12;
    return y + '-' + pad(m + 1);
  }
  function difMeses(a, b) { return (+b.slice(0, 4) - +a.slice(0, 4)) * 12 + (+b.slice(5, 7) - +a.slice(5, 7)); }
  function diaNoMes(ym, d) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7); return iso(y, m, Math.min(d, diasNoMes(y, m))); }
  function ultimoDia(ym) { return diaNoMes(ym, 31); }
  function round2(n) { return Math.round((n + (n < 0 ? -1e-9 : 1e-9)) * 100) / 100; }

  // ---------- números e textos do legado ----------
  function norm(s) {
    return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/º/g, 'o').replace(/\s+/g, ' ').trim();
  }

  // Aceita "1,010.00", "1.010,00", "24.00", "R$ 1.000,00", " -   ", números
  function parseNum(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isNaN(v) ? null : v;
    var s = String(v).replace(/[R$\s]/g, '');
    if (s === '' || s === '-') return null;
    var neg = /^\(.*\)$/.test(s) || s.charAt(0) === '-';
    s = s.replace(/[()\-]/g, '');
    var lc = s.lastIndexOf(','), lp = s.lastIndexOf('.');
    if (lc > -1 && lp > -1) {
      s = lc > lp ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    } else if (lc > -1) {
      s = s.replace(',', '.');
    } else if (lp > -1) {
      if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
    }
    var n = parseFloat(s);
    if (isNaN(n)) return null;
    return neg ? -n : n;
  }

  // "01-janeiro" -> 1
  function parseMesCol(v) {
    if (v == null || v === '') return null;
    var n = parseInt(String(v), 10);
    return (n >= 1 && n <= 12) ? n : null;
  }

  // Data do legado. A planilha antiga exporta m/d/aaaa; outras abas usam d/m/aaaa.
  // A coluna "Mês" desempata: se o primeiro número bate com o mês, é m/d.
  function parseDataLegado(v, mesCol) {
    if (v == null || v === '') return null;
    var s = String(v).trim(), m;
    if ((m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s))) {
      return corrigirData(iso(+m[1], +m[2], +m[3]), mesCol);
    }
    m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
    if (!m) return null;
    var a = +m[1], b = +m[2], y = +m[3]; if (y < 100) y += 2000;
    var mes, dia;
    if (mesCol && a === mesCol && b !== mesCol) { mes = a; dia = b; }
    else if (mesCol && b === mesCol) { dia = a; mes = b; }
    else { dia = a; mes = b; }
    if (mes < 1 || mes > 12 || dia < 1 || dia > diasNoMes(y, mes)) return null;
    return iso(y, mes, dia);
  }
  // Quando o Sheets já converteu a data com o dia e o mês trocados
  function corrigirData(isoStr, mesCol) {
    if (!mesCol) return isoStr;
    var p = parseIso(isoStr);
    if (p.m !== mesCol && p.d === mesCol && p.m <= diasNoMes(p.y, p.d)) return iso(p.y, p.d, p.m);
    return isoStr;
  }

  // ---------- lançamentos ----------
  // l: {id,data,desc,categoria,meio,natureza,destino,previsto,efetivo,conferido,recId,ref}
  // nat: { 'Despesa': {sinal:-1, efeito:'despesa'}, ... }
  function sinalDe(nat, natureza) { return nat[natureza] ? nat[natureza].sinal : -1; }
  function valorUsado(l) { return l.efetivo != null ? l.efetivo : (l.previsto || 0); }

  // Efeito do lançamento no saldo de uma conta. modo 'real' ignora o que ainda é só previsto.
  function efeitoConta(l, chave, nat, modo) {
    var v = l.efetivo != null ? l.efetivo : (modo === 'real' ? null : (l.previsto || 0));
    if (v == null) return 0;
    var s = sinalDe(nat, l.natureza), t = 0;
    if (l.meio === chave) t += s * v;
    if (l.destino === chave) t -= s * v;
    return t;
  }

  // Cartão: o banco informa a fatura como número positivo, internamente é dívida (negativo)
  function instParaInterno(conta, inst) { return conta.classe === 'cartao' ? -inst : inst; }

  function ancoraAte(saldos, chave, dataIso) {
    var best = null;
    saldos.forEach(function (s) {
      if (s.conta === chave && s.inst != null && s.data <= dataIso && (!best || s.data > best.data)) best = s;
    });
    return best;
  }

  // Saldo no fim do dia. Parte do último saldo informado pela instituição (ou do saldo inicial).
  function saldoEm(conta, lancs, saldos, nat, dataIso, modo) {
    var a = ancoraAte(saldos, conta.chave, dataIso);
    var base, baseData;
    if (a) { base = instParaInterno(conta, a.inst); baseData = a.data; }
    else { base = conta.saldoInicial || 0; baseData = conta.dataSaldoInicial || '0000-00-00'; }
    var tot = base;
    lancs.forEach(function (l) {
      if (l.data > baseData && l.data <= dataIso) tot += efeitoConta(l, conta.chave, nat, modo);
    });
    return round2(tot);
  }

  function extrato(conta, lancs, saldos, nat, ini, fim) {
    var abertura = saldoEm(conta, lancs, saldos, nat, addDias(ini, -1), 'real');
    var linhas = [];
    var ordem = 0;
    lancs.forEach(function (l) {
      ordem++;
      if (l.data < ini || l.data > fim) return;
      if (l.meio !== conta.chave && l.destino !== conta.chave) return;
      linhas.push({ l: l, ordem: ordem });
    });
    linhas.sort(function (a, b) { return a.l.data < b.l.data ? -1 : a.l.data > b.l.data ? 1 : a.ordem - b.ordem; });
    var proj = abertura, real = abertura, ent = 0, sai = 0, naoConf = 0;
    var out = linhas.map(function (x) {
      var l = x.l;
      var vp = efeitoConta(l, conta.chave, nat, 'proj');
      var vr = efeitoConta(l, conta.chave, nat, 'real');
      proj += vp; real += vr;
      if (vp >= 0) ent += vp; else sai += vp;
      if (!l.conferido) naoConf++;
      return {
        id: l.id, data: l.data, desc: l.desc, categoria: l.categoria, natureza: l.natureza,
        meio: l.meio, destino: l.destino, valor: round2(vp), previsto: l.previsto, efetivo: l.efetivo,
        pendente: l.efetivo == null, conferido: !!l.conferido, saldo: round2(proj), recId: l.recId, ref: l.ref
      };
    });
    var atrasados = 0;
    lancs.forEach(function (l) {
      if (l.data < ini && l.efetivo == null && (l.meio === conta.chave || l.destino === conta.chave) && (l.previsto || 0) !== 0) atrasados++;
    });
    return {
      abertura: abertura, entradas: round2(ent), saidas: round2(sai),
      fechamentoReal: round2(real), fechamentoComPrevistos: round2(proj),
      naoConferidos: naoConf, previstosEmAtraso: atrasados, linhas: out
    };
  }

  // extra (opcional): { grupoDe: {categoria: grupo}, orcamento: {categoria: valorMensal}, meses: n }
  var NAT_FORA = ['Aplicação', 'Resgate', 'Compromisso'];
  function resumo(lancs, nat, ini, fim, extra) {
    extra = extra || {};
    var grupoDe = extra.grupoDe || {}, orc = extra.orcamento || {}, meses = extra.meses || 1;
    var grupos = { receita: {}, despesa: {}, rendimento: {} }, fora = {};
    function item(g, cat) {
      return g[cat] || (g[cat] = { categoria: cat, grupo: grupoDe[cat] || 'Sem grupo', previsto: 0, realizado: 0, projetado: 0 });
    }
    lancs.forEach(function (l) {
      if (l.data < ini || l.data > fim) return;
      var n = nat[l.natureza];
      if (!n) return;
      if (grupos[n.efeito]) {
        var c = item(grupos[n.efeito], l.categoria);
        c.previsto += l.previsto || 0;
        c.realizado += l.efetivo || 0;
        c.projetado += valorUsado(l);
      } else if (NAT_FORA.indexOf(l.natureza) >= 0) {
        var f = fora[l.natureza] || (fora[l.natureza] = { natureza: l.natureza, previsto: 0, realizado: 0, projetado: 0 });
        f.previsto += l.previsto || 0; f.realizado += l.efetivo || 0; f.projetado += valorUsado(l);
      }
    });
    Object.keys(orc).forEach(function (cat) { if (orc[cat] > 0) item(grupos.despesa, cat); });
    function fecha(g, comOrc) {
      var itens = Object.keys(g).map(function (k) {
        var c = g[k]; c.previsto = round2(c.previsto); c.realizado = round2(c.realizado); c.projetado = round2(c.projetado);
        if (comOrc && orc[c.categoria] > 0) c.orcado = round2(orc[c.categoria] * meses);
        return c;
      }).sort(function (a, b) { return b.projetado - a.projetado; });
      var t = { previsto: 0, realizado: 0, projetado: 0, orcado: 0 };
      itens.forEach(function (c) { t.previsto += c.previsto; t.realizado += c.realizado; t.projetado += c.projetado; t.orcado += c.orcado || 0; });
      var out = { itens: itens, previsto: round2(t.previsto), realizado: round2(t.realizado), projetado: round2(t.projetado) };
      if (comOrc) out.orcado = round2(t.orcado);
      return out;
    }
    var r = fecha(grupos.receita), d = fecha(grupos.despesa, true), rd = fecha(grupos.rendimento);
    var foraLista = NAT_FORA.filter(function (k) { return fora[k]; }).map(function (k) {
      var f = fora[k]; return { natureza: k, previsto: round2(f.previsto), realizado: round2(f.realizado), projetado: round2(f.projetado) };
    });
    return {
      receitas: r, despesas: d, rendimentos: rd, fora: foraLista,
      resultadoRealizado: round2(r.realizado - d.realizado), resultadoProjetado: round2(r.projetado - d.projetado),
      resultadoTotalRealizado: round2(r.realizado + rd.realizado - d.realizado), resultadoTotalProjetado: round2(r.projetado + rd.projetado - d.projetado)
    };
  }

  // ---------- painel ----------
  var CLASSES_PATRIMONIO = ['corrente', 'dinheiro', 'beneficio', 'aplicacao', 'emprestimo', 'outro', 'cartao'];
  var CLASSES_RESERVA = ['corrente', 'dinheiro', 'aplicacao', 'cartao'];

  function somaSaldos(contas, classes, lancs, saldos, nat, data, modo) {
    var t = 0;
    contas.forEach(function (c) { if (classes.indexOf(c.classe) >= 0) t += saldoEm(c, lancs, saldos, nat, data, modo); });
    return round2(t);
  }

  // p: { contas, lancs, saldos, nat, hoje }
  function painel(p) {
    var contas = p.contas, lancs = p.lancs, saldos = p.saldos, nat = p.nat, hoje = p.hoje;
    var cur = hoje.slice(0, 7), classeDe = {};
    contas.forEach(function (c) { classeDe[c.chave] = c.classe; });

    // receitas, despesas e rendimentos mês a mês (só o que foi pago)
    var mapa = {}, cats = {}, compromisso = 0;
    var meses = []; for (var k = 11; k >= 0; k--) meses.push(addMeses(cur, -k));
    meses.forEach(function (ym) { mapa[ym] = { ym: ym, receitas: 0, despesas: 0, rendimentos: 0 }; });
    lancs.forEach(function (l) {
      var n = nat[l.natureza];
      if (!n) return;
      if (l.natureza === 'Compromisso' && l.efetivo != null && l.data <= hoje) compromisso += l.efetivo;
      if (l.efetivo == null) return;
      var m = mapa[l.data.slice(0, 7)];
      if (!m) return;
      if (n.efeito === 'receita') m.receitas += l.efetivo;
      else if (n.efeito === 'rendimento') m.rendimentos += l.efetivo;
      else if (n.efeito === 'despesa') {
        m.despesas += l.efetivo;
        var c = cats[l.categoria] || (cats[l.categoria] = { categoria: l.categoria, mes: 0, soma12: 0 });
        if (l.data.slice(0, 7) === cur) c.mes += l.efetivo; else c.soma12 += l.efetivo;
      }
    });
    var serie = meses.map(function (ym) {
      var m = mapa[ym]; return { ym: ym, receitas: round2(m.receitas), despesas: round2(m.despesas), rendimentos: round2(m.rendimentos) };
    });
    var fechados = serie.filter(function (m) { return m.ym < cur; });
    var ult3 = fechados.slice(-3);
    var rec3 = 0, desp3 = 0;
    ult3.forEach(function (m) { rec3 += m.receitas; desp3 += m.despesas; });
    var taxa3m = ult3.length === 3 && rec3 > 0 ? round2((rec3 - desp3) / rec3 * 100) : null;
    var comDespesa = fechados.filter(function (m) { return m.despesas > 0; }).slice(-6);
    var despMedia = comDespesa.length >= 3 ? round2(comDespesa.reduce(function (t, m) { return t + m.despesas; }, 0) / comDespesa.length) : null;
    var liquido = somaSaldos(contas, CLASSES_RESERVA, lancs, saldos, nat, hoje, 'real');
    var reservaMeses = despMedia ? round2(liquido / despMedia) : null;

    // patrimônio líquido ao fim de cada mês (moeda estrangeira fica de fora: está em outra unidade)
    var primeira = null;
    saldos.forEach(function (s) { if (s.inst != null && (!primeira || s.data < primeira)) primeira = s.data; });
    var ini = addMeses(cur, -11);
    if (primeira && primeira.slice(0, 7) > ini) ini = primeira.slice(0, 7);
    var patrimonio = [];
    for (var ym = ini; ym <= cur; ym = addMeses(ym, 1)) {
      var fimMes = ym === cur ? hoje : diaNoMes(ym, 31);
      patrimonio.push({ ym: ym, valor: somaSaldos(contas, CLASSES_PATRIMONIO, lancs, saldos, nat, fimMes, 'real') });
    }

    var topDespesas = Object.keys(cats).map(function (k2) {
      var c = cats[k2], mesesBase = Math.max(1, Math.min(11, fechados.length));
      return { categoria: c.categoria, mes: round2(c.mes), media: round2(c.soma12 / mesesBase) };
    }).sort(function (a, b) { return Math.max(b.mes, b.media) - Math.max(a.mes, a.media); }).slice(0, 5);

    // próximos 30 dias
    var limite = addDias(hoje, 30), itens = [], atrasados = 0, aPagarFaturas = 0;
    lancs.forEach(function (l) {
      if (l.efetivo != null || !(l.previsto || 0)) return;
      if (l.data < hoje) { atrasados++; return; }
      if (l.data > limite) return;
      if (classeDe[l.meio] === 'cartao' || classeDe[l.destino] === 'cartao') return; // entra pela fatura
      itens.push({ id: l.id, data: l.data, desc: l.desc, categoria: l.categoria, meio: l.meio, valor: round2(sinalDe(nat, l.natureza) * l.previsto), fatura: false });
    });
    contas.forEach(function (g) {
      if (g.classe !== 'cartao' || g.faturaDe || !g.ativo) return;
      faturas(g, contas, lancs, nat, hoje).forEach(function (f) {
        if (f.status === 'paga' || f.vencimento < hoje || f.vencimento > limite || !(f.restante > 0)) return;
        itens.push({ id: 'fat|' + g.chave + '|' + f.id, data: f.vencimento, desc: 'Fatura ' + g.chave, categoria: 'Fatura do cartão', meio: g.contaPgto || '', valor: -f.restante, fatura: true });
        aPagarFaturas += Math.max(0, f.restante - f.agendado);
      });
    });
    itens.sort(function (a, b) { return a.data < b.data ? -1 : a.data > b.data ? 1 : 0; });
    var ent = 0, sai = 0;
    itens.forEach(function (i) { if (i.valor >= 0) ent += i.valor; else sai += i.valor; });
    var disp = somaSaldos(contas, ['corrente', 'dinheiro'], lancs, saldos, nat, hoje, 'real');
    var dispProj = round2(somaSaldos(contas, ['corrente', 'dinheiro'], lancs, saldos, nat, limite, 'proj') - aPagarFaturas);

    return {
      hoje: hoje, serie: serie,
      kpis: {
        resultadoMes: round2(mapa[cur].receitas - mapa[cur].despesas), receitasMes: round2(mapa[cur].receitas), despesasMes: round2(mapa[cur].despesas),
        rendimentosMes: round2(mapa[cur].rendimentos), taxaPoupanca3m: taxa3m, despesaMedia: despMedia, liquido: liquido, reservaMeses: reservaMeses,
        mesesComDados: comDespesa.length, patrimonio: patrimonio.length ? patrimonio[patrimonio.length - 1].valor : 0,
        variacaoPatrimonio: patrimonio.length > 1 ? round2(patrimonio[patrimonio.length - 1].valor - patrimonio[patrimonio.length - 2].valor) : null,
        compromissoPago: round2(compromisso)
      },
      patrimonio: patrimonio, topDespesas: topDespesas,
      proximos: { ini: hoje, fim: limite, itens: itens, entradas: round2(ent), saidas: round2(sai), atrasados: atrasados, disponivel: disp, disponivelProjetado: dispProj }
    };
  }

  // ---------- consulta por categoria e período ----------
  // f: {ini, fim, categoria, natureza, conta, texto, status: 'todos'|'pagos'|'previstos'}
  function consultar(lancs, nat, f) {
    var palavras = norm(f.texto || '').split(' ').filter(function (x) { return x; });
    var linhas = [], ent = 0, sai = 0, entR = 0, saiR = 0, meses = {};
    lancs.forEach(function (l) {
      if (f.ini && l.data < f.ini) return;
      if (f.fim && l.data > f.fim) return;
      if (f.categoria && l.categoria !== f.categoria) return;
      if (f.catSet && !f.catSet[l.categoria]) return;
      if (f.natureza && l.natureza !== f.natureza) return;
      if (f.conta && l.meio !== f.conta && l.destino !== f.conta) return;
      if (f.status === 'pagos' && l.efetivo == null) return;
      if (f.status === 'previstos' && l.efetivo != null) return;
      if (palavras.length) {
        var alvo = norm(l.desc + ' ' + l.categoria + ' ' + l.meio + ' ' + l.destino);
        if (!palavras.every(function (p) { return alvo.indexOf(p) >= 0; })) return;
      }
      var s = sinalDe(nat, l.natureza), v = valorUsado(l), vr = l.efetivo != null ? l.efetivo : 0;
      var ym = l.data.slice(0, 7), m = meses[ym] || (meses[ym] = { ym: ym, entradas: 0, saidas: 0, qtd: 0 });
      if (s > 0) { ent += v; entR += vr; m.entradas += v; } else { sai += v; saiR += vr; m.saidas += v; }
      m.qtd++;
      linhas.push(l);
    });
    linhas.sort(function (a, b) { return a.data < b.data ? 1 : a.data > b.data ? -1 : 0; });
    var lista = Object.keys(meses).sort().map(function (k) {
      var m = meses[k]; m.entradas = round2(m.entradas); m.saidas = round2(m.saidas); return m;
    });
    return {
      qtd: linhas.length, truncado: linhas.length > 300, linhas: linhas.slice(0, 300),
      entradas: round2(ent), saidas: round2(sai), entradasPagas: round2(entR), saidasPagas: round2(saiR), porMes: lista
    };
  }

  // ---------- cartão ----------
  function cicloFatura(dataIso, fech, venc) {
    var p = parseIso(dataIso), ym = p.y + '-' + pad(p.m);
    var closeYm = p.d <= Math.min(fech, diasNoMes(p.y, p.m)) ? ym : addMeses(ym, 1);
    var dueYm = venc > fech ? closeYm : addMeses(closeYm, 1);
    return { id: closeYm, fechamento: diaNoMes(closeYm, fech), vencimento: diaNoMes(dueYm, venc) };
  }

  function difDias(a, b) {
    var pa = parseIso(a), pb = parseIso(b);
    return Math.round((Date.UTC(pb.y, pb.m - 1, pb.d) - Date.UTC(pa.y, pa.m - 1, pa.d)) / 86400000);
  }

  // grupo = cartão principal. Adicionais apontam para ele em "faturaDe".
  // Pagamentos: lançamentos com destino no cartão. Os gerados pelo app têm ref FAT|...;
  // os do histórico entram na fatura cujo vencimento fica mais perto (até 15 dias).
  function faturas(grupo, contas, lancs, nat, hoje) {
    var membros = contas.filter(function (c) { return c.chave === grupo.chave || c.faturaDe === grupo.chave; })
      .map(function (c) { return c.chave; });
    var fech = grupo.diaFech || 1, venc = grupo.diaVenc || 10;
    var aReceber = {};
    contas.forEach(function (c) { if (c.classe === 'emprestimo') aReceber[c.chave] = 1; });
    var mapa = {};
    function nova(id, venc_) {
      return { id: id, fechamento: diaNoMes(id, fech), vencimento: venc_, total: 0, itens: 0, porCartao: {}, pagamentos: [], pagamento: null };
    }
    lancs.forEach(function (l) {
      if (membros.indexOf(l.meio) < 0 || (l.destino && !aReceber[l.destino])) return;
      var cic = cicloFatura(l.data, fech, venc);
      var f = mapa[cic.id] || (mapa[cic.id] = nova(cic.id, cic.vencimento));
      var v = -sinalDe(nat, l.natureza) * valorUsado(l);
      f.total += v; f.itens++;
      f.porCartao[l.meio] = round2((f.porCartao[l.meio] || 0) + v);
    });
    var lista = Object.keys(mapa).map(function (k) { return mapa[k]; });
    lancs.forEach(function (l) {
      if (l.destino !== grupo.chave) return;
      var alvo = null, ligado = false;
      if (l.ref && l.ref.indexOf('FAT|' + grupo.chave + '|') === 0) {
        var id = l.ref.split('|')[2];
        alvo = mapa[id] || (mapa[id] = nova(id, l.data)); ligado = true;
        if (lista.indexOf(alvo) < 0) lista.push(alvo);
      } else {
        var melhor = 16;
        lista.forEach(function (f) {
          var d = Math.abs(difDias(f.vencimento, l.data));
          if (d < melhor) { melhor = d; alvo = f; }
        });
      }
      if (!alvo) return;
      var p = { id: l.id, data: l.data, desc: l.desc, previsto: l.previsto, efetivo: l.efetivo, gerado: ligado };
      alvo.pagamentos.push(p);
      if (ligado) alvo.pagamento = p;
    });
    return lista.sort(function (a, b) { return a.id < b.id ? 1 : -1; }).map(function (f) {
      f.total = round2(f.total);
      var pago = 0, agendado = 0;
      f.pagamentos.forEach(function (p) {
        if (p.efetivo != null) pago += p.efetivo; else agendado += (p.previsto || 0);
      });
      f.pago = round2(pago); f.agendado = round2(agendado);
      f.restante = round2(f.total - pago);
      f.status = pago > 0 && Math.abs(f.restante) < 0.005 ? 'paga'
        : pago > 0 ? 'paga em parte'
        : agendado > 0 ? 'agendada'
        : (f.fechamento >= hoje ? 'aberta' : 'fechada');
      f.divergente = agendado > 0 && pago === 0 && Math.abs(agendado - f.total) > 0.005;
      return f;
    });
  }

  function pagamentoFatura(grupo, fat, categoriaCartao) {
    return {
      data: fat.vencimento, desc: 'Fatura ' + grupo.nome + ' ' + fat.id.slice(5, 7) + '/' + fat.id.slice(0, 4),
      categoria: categoriaCartao, meio: grupo.contaPgto, destino: grupo.chave,
      previsto: Math.max(0, fat.restante != null ? fat.restante : fat.total), efetivo: null, ref: 'FAT|' + grupo.chave + '|' + fat.id
    };
  }

  // ---------- recorrências ----------
  // rec: {id,desc,meio,categoria,destino,previsto,dia,inicio:'yyyy-mm',n,ativo}
  function gerarRecorrencias(recs, lancs, ym) {
    var existentes = {};
    lancs.forEach(function (l) { if (l.ref) existentes[l.ref] = 1; });
    var novos = [], jaExistiam = 0;
    recs.forEach(function (r) {
      if (!r.ativo) return;
      var k = difMeses(r.inicio, ym);
      if (k < 0 || (r.n > 0 && k >= r.n)) return;
      var ref = 'REC|' + r.id + '|' + ym;
      if (existentes[ref]) { jaExistiam++; return; }
      novos.push({
        data: diaNoMes(ym, r.dia), desc: r.desc, categoria: r.categoria, meio: r.meio,
        destino: r.destino || '', previsto: r.previsto || 0, efetivo: null, recId: r.id, ref: ref
      });
    });
    return { novos: novos, jaExistiam: jaExistiam };
  }

  // ---------- parcelamento ----------
  // l: {data, desc, categoria, meio, destino, valor}. efetivas: 'todas' | 'primeira' | 'nenhuma'
  // valorEhTotal: o valor informado é o total da compra (divide em n, a diferença vai na última)
  function gerarParcelas(l, n, valorEhTotal, efetivas, grupoId) {
    var total = valorEhTotal ? l.valor : round2(l.valor * n);
    var base = Math.floor(total / n * 100) / 100;
    var resto = round2(total - base * n);
    var ym = l.data.slice(0, 7), dia = +l.data.slice(8);
    var out = [];
    for (var k = 0; k < n; k++) {
      var v = round2(k === n - 1 ? base + resto : base);
      var efet = efetivas === 'todas' || (efetivas === 'primeira' && k === 0);
      out.push({
        data: diaNoMes(addMeses(ym, k), dia), desc: l.desc + ' (' + (k + 1) + '/' + n + ')',
        categoria: l.categoria, meio: l.meio, destino: l.destino || '',
        previsto: efet ? null : v, efetivo: efet ? v : null, ref: 'PAR|' + grupoId + '|' + (k + 1) + '/' + n
      });
    }
    return out;
  }

  // ---------- importação do legado ----------
  function chaveDup(l) { return [l.data, norm(l.desc), l.meio, l.categoria, round2(l.efetivo || 0), round2(l.previsto || 0)].join('|'); }

  // ctx: {categorias:{nome:natureza}, contas:[...], nat, deParaMeio:{valor:chave}, deParaFatura:[[palavra,chave]]}
  function resolverMeio(valor, ctx) {
    var v = String(valor || '').trim();
    var i;
    for (i = 0; i < ctx.contas.length; i++) if (ctx.contas[i].chave === v) return ctx.contas[i];
    if (ctx.deParaMeio[v]) {
      for (i = 0; i < ctx.contas.length; i++) if (ctx.contas[i].chave === ctx.deParaMeio[v]) return ctx.contas[i];
    }
    var doTipo = ctx.contas.filter(function (c) { return c.tipo === v; });
    return doTipo.length === 1 ? doTipo[0] : null;
  }

  function alias(mapa, v) { var k = norm(v); return (mapa && mapa[k]) ? mapa[k] : String(v == null ? '' : v).trim(); }

  // row: {data, desc, conta, meio, efetivo, previsto, categoria, mes}
  function converterMovimentacao(row, ctx) {
    var data = parseDataLegado(row.data, parseMesCol(row.mes));
    if (!data) return { skip: 'data inválida' };
    var desc = String(row.desc || '').trim();
    var ef = parseNum(row.efetivo), pr = parseNum(row.previsto);
    if (ef == null && pr == null) return { skip: 'sem valor' };
    var meio = resolverMeio(row.meio, ctx);
    if (!meio) return { skip: 'meio de pagamento desconhecido: ' + row.meio };
    var categoria = alias(ctx.aliasCat, row.conta);
    var natureza = ctx.categorias[categoria] || alias(ctx.aliasNat, row.categoria);
    if (!ctx.nat[natureza]) return { skip: 'natureza desconhecida: ' + natureza };
    var novaCategoria = ctx.categorias[categoria] ? null : { descricao: categoria, natureza: natureza };
    if (meio.classe === 'cartao' && natureza === 'Retirada' && categoria === ctx.catFatura) return { skip: 'saldo de cartão (o pagamento da fatura já cobre)' };
    var destino = '', pend = null;
    if (natureza === 'Retirada' && categoria === ctx.catFatura) {
      var d = norm(desc);
      for (var i = 0; i < ctx.deParaFatura.length; i++) {
        if (d.indexOf(norm(ctx.deParaFatura[i][0])) >= 0) { destino = ctx.deParaFatura[i][1]; break; }
      }
      if (!destino) pend = 'pagamento de cartão sem cartão identificado';
    }
    return {
      l: { data: data, desc: desc, categoria: categoria, meio: meio.chave, natureza: natureza, destino: destino, previsto: pr, efetivo: ef, conferido: false, recId: '', ref: '' },
      novaCategoria: novaCategoria, pend: pend
    };
  }

  return {
    pad: pad, iso: iso, parseIso: parseIso, addDias: addDias, addMeses: addMeses, difMeses: difMeses,
    diaNoMes: diaNoMes, ultimoDia: ultimoDia, diasNoMes: diasNoMes, round2: round2, norm: norm,
    parseNum: parseNum, parseMesCol: parseMesCol, parseDataLegado: parseDataLegado, corrigirData: corrigirData,
    sinalDe: sinalDe, valorUsado: valorUsado, efeitoConta: efeitoConta, instParaInterno: instParaInterno,
    ancoraAte: ancoraAte, saldoEm: saldoEm, extrato: extrato, resumo: resumo, painel: painel,
    cicloFatura: cicloFatura, faturas: faturas, pagamentoFatura: pagamentoFatura,
    gerarRecorrencias: gerarRecorrencias, chaveDup: chaveDup, resolverMeio: resolverMeio,
    converterMovimentacao: converterMovimentacao, consultar: consultar, alias: alias, gerarParcelas: gerarParcelas, difDias: difDias
  };
})();
