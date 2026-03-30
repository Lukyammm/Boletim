const ID_PLANILHA = '1XUtI9TSMJmTpbtfLjZbJ-uarRN94lu_Aqpsc46Lxmt4';
const ABA_CAMINHADAS = 'BASE DE DADOS CAMINHADAS';
const ABA_NOTIFICA = 'NOTIFICA - BASE';
const META_INSTITUCIONAL = 80;
const FUSO_HORARIO = 'America/Fortaleza';
const ORDEM_MESES = {
  'JANEIRO': 1,
  'FEVEREIRO': 2,
  'MARÇO': 3,
  'MARCO': 3,
  'ABRIL': 4,
  'MAIO': 5,
  'JUNHO': 6,
  'JULHO': 7,
  'AGOSTO': 8,
  'SETEMBRO': 9,
  'OUTUBRO': 10,
  'NOVEMBRO': 11,
  'DEZEMBRO': 12
};
const MESES_CANONICOS = {
  1: 'Janeiro',
  2: 'Fevereiro',
  3: 'Março',
  4: 'Abril',
  5: 'Maio',
  6: 'Junho',
  7: 'Julho',
  8: 'Agosto',
  9: 'Setembro',
  10: 'Outubro',
  11: 'Novembro',
  12: 'Dezembro'
};

const CLASSIFICACOES_META_5_DIAS = [
  'EVENTO ADVERSO GRAVE',
  'EVENTO ADVERSO ÓBITO',
  'EVENTO ADVERSO OBITO',
  'NEVER EVENT'
];

const METAS_CAMINHADAS = [
  {
    codigo: '1',
    nome: 'Identificação Segura',
    observacaoIdx: 19,
    itens: [
      { codigo: '1.1', nome: 'Paciente com pulseira padrão legível', idx: 14 },
      { codigo: '1.2', nome: 'Placa de leito', idx: 15 },
      { codigo: '1.3', nome: 'Dieta', idx: 16 },
      { codigo: '1.4', nome: 'Medicamentos', idx: 17 },
      { codigo: '1.5', nome: 'Hemocomponentes', idx: 18 }
    ]
  },
  {
    codigo: '2',
    nome: 'Comunicação Efetiva',
    observacaoIdx: 24,
    itens: [
      { codigo: '2.1', nome: 'Reuniões rápidas de segurança', idx: 20 },
      { codigo: '2.2', nome: 'Comunicação de resultados críticos laboratoriais', idx: 21 },
      { codigo: '2.3', nome: 'Registros da transferência de cuidados', idx: 22 },
      { codigo: '2.4', nome: 'Registros da passagem de plantão', idx: 23 }
    ]
  },
  {
    codigo: '3',
    nome: 'Segurança da Cadeia Medicamentosa',
    observacaoIdx: 32,
    itens: [
      { codigo: '3.1', nome: 'Geladeira (temperatura e uso exclusivo)', idx: 25 },
      { codigo: '3.2', nome: 'Validade dos medicamentos da geladeira', idx: 26 },
      { codigo: '3.3', nome: 'Medicamentos multidoses', idx: 27 },
      { codigo: '3.4', nome: 'Eletrólitos', idx: 28 },
      { codigo: '3.5', nome: 'Farmacovigilância', idx: 29 },
      { codigo: '3.6', nome: 'Medicamentos via sonda', idx: 30 },
      { codigo: '3.7', nome: 'MPPs (dupla checagem)', idx: 31 }
    ]
  },
  {
    codigo: '4',
    nome: 'Cirurgia Segura',
    observacaoIdx: 39,
    itens: [
      { codigo: '4.1', nome: 'Demarcação do sítio', idx: 33 },
      { codigo: '4.2', nome: 'Tricotomia', idx: 34 },
      { codigo: '4.3', nome: 'Reserva cirúrgica', idx: 35 },
      { codigo: '4.4', nome: 'Antibioticoprofilaxia', idx: 36 },
      { codigo: '4.5', nome: 'Amostras identificadas', idx: 37 },
      { codigo: '4.6', nome: 'Avaliação anestésica', idx: 38 }
    ]
  },
  {
    codigo: '5',
    nome: 'Higienização das Mãos',
    observacaoIdx: 43,
    itens: [
      { codigo: '5.1', nome: 'Higienização correta', idx: 40 },
      { codigo: '5.2', nome: 'Dispensadores abastecidos', idx: 41 },
      { codigo: '5.3', nome: 'Sem adornos', idx: 42 }
    ]
  },
  {
    codigo: '6',
    nome: 'Prevenção de Lesão por Pressão',
    observacaoIdx: 50,
    itens: [
      { codigo: '6.1', nome: 'Avaliação na admissão', idx: 44 },
      { codigo: '6.2', nome: 'Avaliação diária', idx: 45 },
      { codigo: '6.3', nome: 'Mudança de decúbito', idx: 46 },
      { codigo: '6.4', nome: 'Cuidados com a pele', idx: 47 },
      { codigo: '6.5', nome: 'Hidratação', idx: 48 },
      { codigo: '6.6', nome: 'Monitoramento da pele', idx: 49 }
    ]
  },
  {
    codigo: '7',
    nome: 'Prevenção de Quedas',
    observacaoIdx: 57,
    itens: [
      { codigo: '7.1', nome: 'Avaliação na admissão', idx: 51 },
      { codigo: '7.2', nome: 'Avaliação diária', idx: 52 },
      { codigo: '7.3', nome: 'Grades e rodas', idx: 53 },
      { codigo: '7.4', nome: 'Calçado antiderrapante', idx: 54 },
      { codigo: '7.5', nome: 'Ambiente seguro', idx: 55 },
      { codigo: '7.6', nome: 'Sinalização de risco', idx: 56 }
    ]
  }
];

function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.api === '1') {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    return ContentService
      .createTextOutput(JSON.stringify(montarPayload(ss, extrairFiltros(params))))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Boletim COSEP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function extrairFiltros(params) {
  const bruto = params || {};
  return {
    caminhadas: {
      anos: extrairListaFiltros(bruto.anosCosep || bruto.anoCosep || bruto.ano || '', normalizarAno),
      meses: extrairListaFiltros(bruto.mesesCosep || bruto.mesCosep || bruto.mes || '', normalizarMes),
      unidades: extrairListaFiltros(bruto.unidadesCosep || bruto.unidadeCosep || bruto.unidade || '', value => String(value || '').trim())
    },
    notificacoes: {
      anos: extrairListaFiltros(bruto.anosNotifica || bruto.anoNotifica || bruto.ano || '', normalizarAno),
      meses: extrairListaFiltros(bruto.mesesNotifica || bruto.mesNotifica || bruto.mes || '', normalizarMes),
      unidades: extrairListaFiltros(bruto.unidadesNotifica || bruto.unidadeNotifica || bruto.unidade || '', value => String(value || '').trim())
    }
  };
}

function extrairListaFiltros(rawValue, normalizerFn) {
  const brutoLista = Array.isArray(rawValue) ? rawValue : String(rawValue || '').split(/[|,]/);

  const valores = brutoLista
    .map(item => String(item || '').trim())
    .filter(Boolean)
    .map(item => normalizerFn ? normalizerFn(item) : item)
    .filter(Boolean);

  return [...new Set(valores)];
}

function montarPayload(ss, filtros) {
  const filtrosAplicados = filtros || { caminhadas: {}, notificacoes: {} };
  return {
    success: true,
    geradoEm: Utilities.formatDate(new Date(), FUSO_HORARIO, "dd/MM/yyyy 'às' HH:mm"),
    filtros: getFiltros(ss),
    aplicado: filtrosAplicados,
    caminhadas: processarCaminhadas(ss, filtrosAplicados.caminhadas || {}),
    notificacoes: processarNotificacoes(ss, filtrosAplicados.notificacoes || {})
  };
}

function obterPayload(filtros) {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  return montarPayload(ss, extrairFiltros(filtros || {}));
}

function normalizarTexto(valor) {
  return String(valor == null ? '' : valor)
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function normalizarMes(valor) {
  const texto = normalizarTexto(valor);
  if (!texto) return '';

  if (ORDEM_MESES[texto]) {
    return MESES_CANONICOS[ORDEM_MESES[texto]];
  }

  const numero = Number(texto);
  if (!Number.isNaN(numero) && numero >= 1 && numero <= 12) {
    return MESES_CANONICOS[numero];
  }

  const base = String(valor == null ? '' : valor).trim().toLowerCase();
  return base ? base.charAt(0).toUpperCase() + base.slice(1) : '';
}

function normalizarAno(valor) {
  return String(valor == null ? '' : valor).trim();
}

function mesParaOrdem(valor) {
  const texto = normalizarTexto(valor);
  if (!texto) return 99;
  if (ORDEM_MESES[texto]) return ORDEM_MESES[texto];

  const numero = Number(texto);
  if (!Number.isNaN(numero) && numero >= 1 && numero <= 12) return numero;
  return 99;
}

function ordenarMeses(a, b) {
  const ordemA = mesParaOrdem(a);
  const ordemB = mesParaOrdem(b);
  if (ordemA !== ordemB) return ordemA - ordemB;
  return String(a).localeCompare(String(b), 'pt-BR');
}

function ehSim(valor) {
  const v = normalizarTexto(valor);
  return ['SIM', 'S', 'CONFORME', 'OK', 'ADEQUADO'].includes(v);
}

function ehNao(valor) {
  const v = normalizarTexto(valor);
  return ['NÃO', 'NAO', 'N', 'INCONFORME', 'INADEQUADO', 'NÃO CONFORME', 'NAO CONFORME'].includes(v);
}

function incrementarMapa(mapa, chave) {
  const valor = String(chave == null ? '' : chave).trim() || 'Não informado';
  mapa[valor] = (mapa[valor] || 0) + 1;
}

function getUnidade(row) {
  const indices = [4, 68, 69, 70, 71, 72, 73, 74, 75, 76];
  for (let i = 0; i < indices.length; i++) {
    const valor = String(row[indices[i]] == null ? '' : row[indices[i]]).trim();
    if (valor) return valor;
  }
  return 'Não informado';
}

function coletarFiltros(dados, config) {
  const anos = {};
  const meses = {};
  const unidades = {};

  dados.forEach(row => {
    const ano = normalizarAno(row[config.anoIdx]);
    const mes = normalizarMes(row[config.mesIdx]);
    const unidade = config.unidadeFn(row);

    if (ano) anos[ano] = true;
    if (mes) meses[mes] = true;
    if (unidade) unidades[unidade] = true;
  });

  return {
    anos: Object.keys(anos).sort((a, b) => Number(a) - Number(b) || a.localeCompare(b, 'pt-BR')),
    meses: Object.keys(meses).sort(ordenarMeses),
    unidades: Object.keys(unidades).sort((a, b) => a.localeCompare(b, 'pt-BR'))
  };
}

function getFiltros(ss) {
  const shCaminhadas = ss.getSheetByName(ABA_CAMINHADAS);
  const shNotifica = ss.getSheetByName(ABA_NOTIFICA);

  const dadosCaminhadas = shCaminhadas.getDataRange().getValues().slice(1);
  const dadosNotifica = shNotifica.getDataRange().getValues().slice(2);

  return {
    caminhadas: coletarFiltros(dadosCaminhadas, {
      anoIdx: 3,
      mesIdx: 2,
      unidadeFn: getUnidade
    }),
    notificacoes: coletarFiltros(dadosNotifica, {
      anoIdx: 3,
      mesIdx: 2,
      unidadeFn: row => String(row[6] || '').trim() || 'Não informado'
    })
  };
}

function processarCaminhadas(ss, filtros) {
  const sh = ss.getSheetByName(ABA_CAMINHADAS);
  const linhas = sh.getDataRange().getValues().slice(1);

  const metas = METAS_CAMINHADAS.map(meta => ({
    codigo: meta.codigo,
    nome: meta.nome,
    avaliados: 0,
    conformes: 0,
    naoConformes: 0,
    percentual: 0,
    totalObservacoes: 0,
    observacoes: [],
    itens: meta.itens.map(item => ({
      codigo: item.codigo,
      nome: item.nome,
      avaliados: 0,
      conformes: 0,
      naoConformes: 0,
      percentual: 0
    }))
  }));

  const porUnidade = {};
  const observacoesGerais = [];
  const avaliadores = {};
  const metasAcionadas = {};

  let totalAvaliacoes = 0;
  let geralConformes = 0;
  let geralNaoConformes = 0;
  let totalComObservacao = 0;
  let totalComFoto = 0;

  const anosFiltro = new Set((filtros.anos || []).map(normalizarAno).filter(Boolean));
  const mesesFiltro = new Set((filtros.meses || []).map(normalizarMes).filter(Boolean));
  const unidadesFiltro = new Set((filtros.unidades || []).map(item => String(item || '').trim()).filter(Boolean));

  const linhasFiltradas = linhas.filter(row => {
    const ano = normalizarAno(row[3]);
    const mes = normalizarMes(row[2]);
    const unidade = getUnidade(row);

    if (anosFiltro.size && !anosFiltro.has(ano)) return false;
    if (mesesFiltro.size && !mesesFiltro.has(mes)) return false;
    if (unidadesFiltro.size && !unidadesFiltro.has(unidade)) return false;
    return true;
  });

  totalAvaliacoes = linhasFiltradas.length;

  linhasFiltradas.forEach(row => {
    const unidade = getUnidade(row);
    const observacaoGeral = String(row[59] || '').trim();
    const temFoto = String(row[58] || '').trim();
    const avaliador = String(row[9] || '').trim();
    const metaPrincipal = String(row[13] || '').trim();

    incrementarMapa(porUnidade, unidade);
    if (avaliador) avaliadores[avaliador] = true;
    if (metaPrincipal) metasAcionadas[metaPrincipal] = true;
    if (temFoto) totalComFoto++;

    if (observacaoGeral) {
      totalComObservacao++;
      observacoesGerais.push({ unidade: unidade, texto: observacaoGeral });
    }

    METAS_CAMINHADAS.forEach((metaDef, metaIndex) => {
      metaDef.itens.forEach((itemDef, itemIndex) => {
        const valor = row[itemDef.idx];
        const item = metas[metaIndex].itens[itemIndex];

        if (ehSim(valor)) {
          item.conformes++;
          item.avaliados++;
          metas[metaIndex].conformes++;
          metas[metaIndex].avaliados++;
          geralConformes++;
        } else if (ehNao(valor)) {
          item.naoConformes++;
          item.avaliados++;
          metas[metaIndex].naoConformes++;
          metas[metaIndex].avaliados++;
          geralNaoConformes++;
        }
      });

      const observacaoMeta = String(row[metaDef.observacaoIdx] || '').trim();
      if (observacaoMeta) {
        metas[metaIndex].totalObservacoes++;
        metas[metaIndex].observacoes.push({ unidade: unidade, texto: observacaoMeta });
      }
    });
  });

  metas.forEach(meta => {
    meta.percentual = meta.avaliados ? Number(((meta.conformes / meta.avaliados) * 100).toFixed(1)) : 0;
    meta.itens.forEach(item => {
      item.percentual = item.avaliados ? Number(((item.conformes / item.avaliados) * 100).toFixed(1)) : 0;
    });
    meta.observacoes = meta.observacoes.slice(0, 3);
  });

  const totalAvaliado = geralConformes + geralNaoConformes;
  const conformidadeGeral = totalAvaliado ? Number(((geralConformes / totalAvaliado) * 100).toFixed(1)) : 0;

  const evolucaoMensalMap = {};
  linhas
    .filter(row => !anosFiltro.size || anosFiltro.has(normalizarAno(row[3])))
    .filter(row => !unidadesFiltro.size || unidadesFiltro.has(getUnidade(row)))
    .forEach(row => {
      const mes = normalizarMes(row[2]) || 'Sem mês';
      if (!evolucaoMensalMap[mes]) {
        evolucaoMensalMap[mes] = { conformes: 0, naoConformes: 0 };
      }

      METAS_CAMINHADAS.forEach(meta => {
        meta.itens.forEach(item => {
          const valor = row[item.idx];
          if (ehSim(valor)) evolucaoMensalMap[mes].conformes++;
          else if (ehNao(valor)) evolucaoMensalMap[mes].naoConformes++;
        });
      });
    });

  const evolucaoMensal = Object.keys(evolucaoMensalMap)
    .sort(ordenarMeses)
    .map(mes => {
      const item = evolucaoMensalMap[mes];
      const total = item.conformes + item.naoConformes;
      return {
        mes: mes,
        percentual: total ? Number(((item.conformes / total) * 100).toFixed(1)) : 0
      };
    });

  const unidadesOrdenadas = Object.entries(porUnidade)
    .sort((a, b) => b[1] - a[1])
    .map(([nome, quantidade]) => ({ nome, quantidade }));

  const metasCriticas = metas
    .slice()
    .sort((a, b) => a.percentual - b.percentual)
    .slice(0, 3)
    .map(meta => ({ codigo: meta.codigo, nome: meta.nome, percentual: meta.percentual }));

  return {
    totalAvaliacoes: totalAvaliacoes,
    totalUnidades: unidadesOrdenadas.length,
    totalAvaliadores: Object.keys(avaliadores).length,
    totalComObservacao: totalComObservacao,
    totalComFoto: totalComFoto,
    metasSelecionadas: Object.keys(metasAcionadas).length,
    geralConformes: geralConformes,
    geralNaoConformes: geralNaoConformes,
    conformidadeGeral: conformidadeGeral,
    metaInstitucional: META_INSTITUCIONAL,
    diferencaMeta: Number((conformidadeGeral - META_INSTITUCIONAL).toFixed(1)),
    metas: metas,
    porUnidade: unidadesOrdenadas,
    evolucaoMensal: evolucaoMensal,
    metasCriticas: metasCriticas,
    observacoesGerais: observacoesGerais.slice(0, 5)
  };
}

function processarNotificacoes(ss, filtros) {
  const sh = ss.getSheetByName(ABA_NOTIFICA);
  const linhas = sh.getDataRange().getValues().slice(2);

  const desempenhoResposta = {
    meta5: criarResumoPrazoResposta(5),
    meta10: criarResumoPrazoResposta(10)
  };

  const anosFiltro = new Set((filtros.anos || []).map(normalizarAno).filter(Boolean));
  const mesesFiltro = new Set((filtros.meses || []).map(normalizarMes).filter(Boolean));
  const unidadesFiltro = new Set((filtros.unidades || []).map(item => String(item || '').trim()).filter(Boolean));

  const filtradas = linhas.filter(row => {
    const mes = normalizarMes(row[2]);
    const ano = normalizarAno(row[3]);
    const setor = String(row[6] || '').trim();

    if (anosFiltro.size && !anosFiltro.has(ano)) return false;
    if (mesesFiltro.size && !mesesFiltro.has(mes)) return false;
    if (unidadesFiltro.size && !unidadesFiltro.has(setor)) return false;
    return true;
  });

  const porTipo = {};
  const porSetor = {};
  const porNatureza = {};
  const porStatus = {};
  const porStatusResposta = {};

  let afetouSim = 0;
  let afetouNao = 0;
  let concluidas = 0;
  let pendentes = 0;

  filtradas.forEach(row => {
    const tipo = String(row[8] || '').trim() || 'Não informado';
    const setor = String(row[6] || '').trim() || 'Não informado';
    const natureza = String(row[10] || '').trim() || 'Não informado';
    const status = String(row[13] || '').trim() || 'Não informado';
    const statusResposta = String(row[15] || '').trim() || 'Não informado';
    const afetou = normalizarTexto(row[11]);

    incrementarMapa(porTipo, tipo);
    incrementarMapa(porSetor, setor);
    incrementarMapa(porNatureza, natureza);
    incrementarMapa(porStatus, status);
    incrementarMapa(porStatusResposta, statusResposta);

    const classificacao = normalizarTexto(row[8]);
    const dataClassificacao = converterEmData(row[14]);
    const dataResposta = converterEmData(row[16]);
    const diasResposta = calcularDiferencaDias(dataClassificacao, dataResposta);
    const grupoPrazo = CLASSIFICACOES_META_5_DIAS.includes(classificacao) ? desempenhoResposta.meta5 : desempenhoResposta.meta10;

    atualizarResumoPrazoResposta(grupoPrazo, dataClassificacao, diasResposta);

    if (afetou === 'SIM') afetouSim++;
    else afetouNao++;

    if (normalizarTexto(status).indexOf('CONCLU') === 0) concluidas++;
    else pendentes++;
  });

  const tabela = filtradas.slice(0, 12).map(row => ({
    numeroNotivisa: row[0],
    link: row[1],
    mes: normalizarMes(row[2]),
    ano: row[3],
    codigo: row[4],
    dataOcorrencia: formatarData(row[5]),
    setorNotificado: row[6],
    localOcorrencia: row[7],
    tipoClassificacao: row[8],
    codInteracao: row[9],
    natureza: row[10],
    afetouPaciente: row[11],
    prontuario: row[12],
    status: row[13],
    dataClassificacao: formatarData(row[14]),
    statusResposta: row[15],
    dataResposta: formatarData(row[16]),
    dataConclusao: formatarData(row[17])
  }));

  return {
    total: filtradas.length,
    afetouSim: afetouSim,
    afetouNao: afetouNao,
    percentualAfetouPaciente: filtradas.length ? Number(((afetouSim / filtradas.length) * 100).toFixed(1)) : 0,
    concluidas: concluidas,
    pendentes: pendentes,
    totalSetores: Object.keys(porSetor).length,
    porTipo: ordenarMapaPorValor(porTipo),
    porSetor: ordenarMapaPorValor(porSetor),
    porNatureza: ordenarMapaPorValor(porNatureza),
    porStatus: ordenarMapaPorValor(porStatus),
    porStatusResposta: ordenarMapaPorValor(porStatusResposta),
    desempenhoResposta: {
      meta5: finalizarResumoPrazoResposta(desempenhoResposta.meta5),
      meta10: finalizarResumoPrazoResposta(desempenhoResposta.meta10)
    },
    tabela: tabela
  };
}

function criarResumoPrazoResposta(metaDias) {
  return {
    metaDias: metaDias,
    dentroPrazo: 0,
    foraPrazo: 0,
    totalComResposta: 0,
    somaDias: 0,
    mediaDias: 0
  };
}

function atualizarResumoPrazoResposta(resumo, dataClassificacao, diasResposta) {
  if (!resumo || !dataClassificacao) return;

  if (diasResposta == null) {
    resumo.foraPrazo++;
    return;
  }

  resumo.totalComResposta++;
  resumo.somaDias += diasResposta;

  if (diasResposta <= resumo.metaDias) resumo.dentroPrazo++;
  else resumo.foraPrazo++;
}

function finalizarResumoPrazoResposta(resumo) {
  const base = resumo || criarResumoPrazoResposta(0);
  const total = Number(base.totalComResposta || 0);
  return {
    metaDias: base.metaDias,
    dentroPrazo: Number(base.dentroPrazo || 0),
    foraPrazo: Number(base.foraPrazo || 0),
    totalComResposta: total,
    mediaDias: total ? Number((base.somaDias / total).toFixed(1)) : 0
  };
}

function ordenarMapaPorValor(mapa) {
  const ordenado = {};
  Object.keys(mapa)
    .sort((a, b) => mapa[b] - mapa[a] || a.localeCompare(b, 'pt-BR'))
    .forEach(chave => {
      ordenado[chave] = mapa[chave];
    });
  return ordenado;
}

function formatarData(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]' && !Number.isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, FUSO_HORARIO, 'dd/MM/yyyy');
  }
  return String(valor).trim();
}

function converterEmData(valor) {
  if (!valor) return null;

  if (Object.prototype.toString.call(valor) === '[object Date]' && !Number.isNaN(valor.getTime())) {
    return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  }

  const texto = String(valor).trim();
  if (!texto) return null;

  const partes = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (partes) {
    return new Date(Number(partes[3]), Number(partes[2]) - 1, Number(partes[1]));
  }

  const data = new Date(texto);
  if (Number.isNaN(data.getTime())) return null;

  return new Date(data.getFullYear(), data.getMonth(), data.getDate());
}

function calcularDiferencaDias(dataInicial, dataFinal) {
  if (!dataInicial || !dataFinal) return null;

  const inicio = new Date(dataInicial.getFullYear(), dataInicial.getMonth(), dataInicial.getDate());
  const fim = new Date(dataFinal.getFullYear(), dataFinal.getMonth(), dataFinal.getDate());
  const diferenca = Math.round((fim.getTime() - inicio.getTime()) / 86400000);

  return diferenca < 0 ? 0 : diferenca;
}
