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

  function resumo(lancs, nat, ini, fim) {
    var grupos = { receita: {}, despesa: {} };
    lancs.forEach(function (l) {
      if (l.data < ini || l.data > fim) return;
      var n = nat[l.natureza];
      if (!n || !grupos[n.efeito]) return;
      var g = grupos[n.efeito];
      var c = g[l.categoria] || (g[l.categoria] = { categoria: l.categoria, previsto: 0, realizado: 0, projetado: 0 });
      c.previsto += l.previsto || 0;
      c.realizado += l.efetivo || 0;
      c.projetado += valorUsado(l);
    });
    function fecha(g) {
      var itens = Object.keys(g).map(function (k) {
        var c = g[k]; c.previsto = round2(c.previsto); c.realizado = round2(c.realizado); c.projetado = round2(c.projetado); return c;
      }).sort(function (a, b) { return b.projetado - a.projetado; });
      var t = { previsto: 0, realizado: 0, projetado: 0 };
      itens.forEach(function (c) { t.previsto += c.previsto; t.realizado += c.realizado; t.projetado += c.projetado; });
      return { itens: itens, previsto: round2(t.previsto), realizado: round2(t.realizado), projetado: round2(t.projetado) };
    }
    var r = fecha(grupos.receita), d = fecha(grupos.despesa);
    return { receitas: r, despesas: d, resultadoRealizado: round2(r.realizado - d.realizado), resultadoProjetado: round2(r.projetado - d.projetado) };
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
    var mapa = {};
    function nova(id, venc_) {
      return { id: id, fechamento: diaNoMes(id, fech), vencimento: venc_, total: 0, itens: 0, porCartao: {}, pagamentos: [], pagamento: null };
    }
    lancs.forEach(function (l) {
      if (membros.indexOf(l.meio) < 0 || l.destino) return;
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
    ancoraAte: ancoraAte, saldoEm: saldoEm, extrato: extrato, resumo: resumo,
    cicloFatura: cicloFatura, faturas: faturas, pagamentoFatura: pagamentoFatura,
    gerarRecorrencias: gerarRecorrencias, chaveDup: chaveDup, resolverMeio: resolverMeio,
    converterMovimentacao: converterMovimentacao, alias: alias, gerarParcelas: gerarParcelas, difDias: difDias
  };
})();
