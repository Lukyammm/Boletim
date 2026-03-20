const ID_PLANILHA = '1XUtI9TSMJmTpbtfLjZbJ-uarRN94lu_Aqpsc46Lxmt4';
const ABA_CAMINHADAS = 'BASE DE DADOS CAMINHADAS';
const ABA_NOTIFICA = 'NOTIFICA - BASE';
const META_INSTITUCIONAL = 85;
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
    const ano = String(params.ano || '').trim();
    const mes = String(params.mes || '').trim();
    const unidade = String(params.unidade || '').trim();
    const ss = SpreadsheetApp.openById(ID_PLANILHA);

    return ContentService
      .createTextOutput(JSON.stringify(montarPayload(ss, { ano, mes, unidade })))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Boletim COSEP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function montarPayload(ss, filtros) {
  return {
    success: true,
    geradoEm: Utilities.formatDate(new Date(), FUSO_HORARIO, "dd/MM/yyyy 'às' HH:mm"),
    filtros: getFiltros(ss),
    caminhadas: processarCaminhadas(ss, filtros),
    notificacoes: processarNotificacoes(ss, filtros)
  };
}

function obterPayload(filtros) {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const parametros = filtros || {};

  return montarPayload(ss, {
    ano: String(parametros.ano || '').trim(),
    mes: String(parametros.mes || '').trim(),
    unidade: String(parametros.unidade || '').trim()
  });
}

function normalizarTexto(valor) {
  return String(valor == null ? '' : valor)
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function normalizarMes(valor) {
  return String(valor == null ? '' : valor).trim();
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
  return ['NÃO', 'NAO', 'N', 'INCONFORME', 'INADEQUADO'].includes(v);
}

function incrementarMapa(mapa, chave) {
  const valor = String(chave == null ? '' : chave).trim() || 'Não informado';
  mapa[valor] = (mapa[valor] || 0) + 1;
}

function getUnidade(row) {
  const indices = [4, 68, 69, 70, 71, 72, 73, 74, 75, 76]; // E, BQ..BY
  for (let i = 0; i < indices.length; i++) {
    const valor = String(row[indices[i]] == null ? '' : row[indices[i]]).trim();
    if (valor) return valor;
  }
  return 'Não informado';
}

function getFiltros(ss) {
  const shCaminhadas = ss.getSheetByName(ABA_CAMINHADAS);
  const shNotifica = ss.getSheetByName(ABA_NOTIFICA);

  const dadosCaminhadas = shCaminhadas.getDataRange().getValues().slice(1);
  const dadosNotifica = shNotifica.getDataRange().getValues().slice(2);

  const anos = {};
  const meses = {};
  const unidades = {};

  dadosCaminhadas.forEach(row => {
    const ano = normalizarAno(row[3]);
    const mes = normalizarMes(row[2]);
    const unidade = getUnidade(row);
    if (ano) anos[ano] = true;
    if (mes) meses[mes] = true;
    if (unidade) unidades[unidade] = true;
  });

  dadosNotifica.forEach(row => {
    const ano = normalizarAno(row[3]);
    const mes = normalizarMes(row[2]);
    const setor = String(row[6] || '').trim();
    if (ano) anos[ano] = true;
    if (mes) meses[mes] = true;
    if (setor) unidades[setor] = true;
  });

  return {
    anos: Object.keys(anos).sort((a, b) => Number(a) - Number(b) || a.localeCompare(b, 'pt-BR')),
    meses: Object.keys(meses).sort(ordenarMeses),
    unidades: Object.keys(unidades).sort((a, b) => a.localeCompare(b, 'pt-BR'))
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

  const linhasFiltradas = linhas.filter(row => {
    const ano = normalizarAno(row[3]);
    const mes = normalizarMes(row[2]);
    const unidade = getUnidade(row);

    if (filtros.ano && ano !== filtros.ano) return false;
    if (filtros.mes && mes !== filtros.mes) return false;
    if (filtros.unidade && unidade !== filtros.unidade) return false;
    return true;
  });

  totalAvaliacoes = linhasFiltradas.length;

  linhasFiltradas.forEach(row => {
    const unidade = getUnidade(row);
    const observacaoGeral = String(row[59] || '').trim(); // BH
    const temFoto = String(row[58] || '').trim(); // BG
    const avaliador = String(row[9] || '').trim(); // J
    const metaPrincipal = String(row[13] || '').trim(); // N

    incrementarMapa(porUnidade, unidade);
    if (avaliador) avaliadores[avaliador] = true;
    if (metaPrincipal) metasAcionadas[metaPrincipal] = true;
    if (temFoto) totalComFoto++;

    if (observacaoGeral) {
      totalComObservacao++;
      observacoesGerais.push({
        unidade: unidade,
        texto: observacaoGeral
      });
    }

    METAS_CAMINHADAS.forEach((metaDef, metaIndex) => {
      let metaTemAchado = false;

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
          metaTemAchado = true;
        }
      });

      const observacaoMeta = String(row[metaDef.observacaoIdx] || '').trim();
      if (observacaoMeta) {
        metas[metaIndex].observacoes.push({
          unidade: unidade,
          texto: observacaoMeta
        });
        metaTemAchado = true;
      }

      metas[metaIndex].percentual = metas[metaIndex].avaliados
        ? Number(((metas[metaIndex].conformes / metas[metaIndex].avaliados) * 100).toFixed(1))
        : 0;
    });
  });

  metas.forEach(meta => {
    meta.percentual = meta.avaliados ? Number(((meta.conformes / meta.avaliados) * 100).toFixed(1)) : 0;
    meta.itens.forEach(item => {
      item.percentual = item.avaliados ? Number(((item.conformes / item.avaliados) * 100).toFixed(1)) : 0;
    });
    meta.observacoes = meta.observacoes.slice(0, 4);
  });

  const totalAvaliado = geralConformes + geralNaoConformes;
  const conformidadeGeral = totalAvaliado ? Number(((geralConformes / totalAvaliado) * 100).toFixed(1)) : 0;

  const evolucaoMensalMap = {};
  linhas
    .filter(row => !filtros.ano || normalizarAno(row[3]) === filtros.ano)
    .filter(row => !filtros.unidade || getUnidade(row) === filtros.unidade)
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
    observacoesGerais: observacoesGerais.slice(0, 6)
  };
}

function processarNotificacoes(ss, filtros) {
  const sh = ss.getSheetByName(ABA_NOTIFICA);
  const linhas = sh.getDataRange().getValues().slice(2);

  const filtradas = linhas.filter(row => {
    const mes = normalizarMes(row[2]);
    const ano = normalizarAno(row[3]);
    const setor = String(row[6] || '').trim();

    if (filtros.ano && ano !== filtros.ano) return false;
    if (filtros.mes && mes !== filtros.mes) return false;
    if (filtros.unidade && setor !== filtros.unidade) return false;
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

    if (afetou === 'SIM') afetouSim++;
    else afetouNao++;

    if (normalizarTexto(status) === 'CONCLUÍDO' || normalizarTexto(status) === 'CONCLUIDO') concluidas++;
    else pendentes++;
  });

  const tabela = filtradas.slice(0, 30).map(row => ({
    numeroNotivisa: row[0],
    link: row[1],
    mes: row[2],
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
    tabela: tabela
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
