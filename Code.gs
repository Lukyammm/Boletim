const ID_PLANILHA = '1XUtI9TSMJmTpbtfLjZbJ-uarRN94lu_Aqpsc46Lxmt4';
const ABA_CAMINHADAS = 'BASE DE DADOS CAMINHADAS';
const ABA_NOTIFICA = 'NOTIFICA - BASE';
const META_INSTITUCIONAL = 85;

function doGet(e) {
  e = e || {};
  const params = e.parameter || {};

  if (params.api === '1') {
    const ano = params.ano || '';
    const mes = params.mes || '';
    const unidade = params.unidade || '';

    const ss = SpreadsheetApp.openById(ID_PLANILHA);

    const payload = {
      success: true,
      filtros: getFiltros(ss),
      caminhadas: processarCaminhadas(ss, ano, mes, unidade),
      notificacoes: processarNotificacoes(ss, ano, mes, unidade)
    };

    return ContentService
      .createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Boletim COSEP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

function ehSim(valor) {
  const v = normalizarTexto(valor);
  return ['SIM', 'S', 'CONFORME', 'OK'].includes(v);
}

function ehNao(valor) {
  const v = normalizarTexto(valor);
  return ['NÃO', 'NAO', 'N', 'INCONFORME'].includes(v);
}

function getUnidade(row) {
  const indices = [4, 68, 69, 70, 71, 72, 73, 74, 75, 76]; // E, BQ..BY

  for (const i of indices) {
    const valor = String(row[i] == null ? '' : row[i]).trim();
    if (valor) return valor;
  }

  return 'Não informado';
}

function getFiltros(ss) {
  const sh = ss.getSheetByName(ABA_CAMINHADAS);
  const dados = sh.getDataRange().getValues();
  const linhas = dados.slice(1);

  const anos = [...new Set(linhas.map(r => normalizarAno(r[3])).filter(Boolean))].sort();
  const meses = [...new Set(linhas.map(r => normalizarMes(r[2])).filter(Boolean))].sort();
  const unidades = [...new Set(linhas.map(r => getUnidade(r)).filter(Boolean))].sort();

  return { anos, meses, unidades };
}

function processarCaminhadas(ss, anoFiltro, mesFiltro, unidadeFiltro) {
  const sh = ss.getSheetByName(ABA_CAMINHADAS);
  const dados = sh.getDataRange().getValues();
  const linhas = dados.slice(1);

  const metasDef = [
    {
      codigo: '1',
      nome: 'Identificação Segura',
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
      itens: [
        { codigo: '5.1', nome: 'Higienização correta', idx: 40 },
        { codigo: '5.2', nome: 'Dispensadores abastecidos', idx: 41 },
        { codigo: '5.3', nome: 'Sem adornos', idx: 42 }
      ]
    },
    {
      codigo: '6',
      nome: 'Prevenção de Lesão por Pressão',
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

  const metas = metasDef.map(meta => ({
    codigo: meta.codigo,
    nome: meta.nome,
    conformes: 0,
    naoConformes: 0,
    avaliados: 0,
    percentual: 0,
    itens: meta.itens.map(item => ({
      codigo: item.codigo,
      nome: item.nome,
      idx: item.idx,
      conformes: 0,
      naoConformes: 0,
      avaliados: 0,
      percentual: 0
    }))
  }));

  let totalAvaliacoes = 0;
  let geralConformes = 0;
  let geralNaoConformes = 0;

  const linhasFiltradas = linhas.filter(row => {
    const ano = normalizarAno(row[3]);
    const mes = normalizarMes(row[2]);
    const unidade = getUnidade(row);

    if (anoFiltro && ano !== String(anoFiltro)) return false;
    if (mesFiltro && mes !== String(mesFiltro)) return false;
    if (unidadeFiltro && unidade !== String(unidadeFiltro)) return false;

    return true;
  });

  totalAvaliacoes = linhasFiltradas.length;

  linhasFiltradas.forEach(row => {
    metas.forEach(meta => {
      meta.itens.forEach(item => {
        const valor = row[item.idx];

        if (ehSim(valor)) {
          item.conformes++;
          item.avaliados++;
          meta.conformes++;
          meta.avaliados++;
          geralConformes++;
        } else if (ehNao(valor)) {
          item.naoConformes++;
          item.avaliados++;
          meta.naoConformes++;
          meta.avaliados++;
          geralNaoConformes++;
        }
      });
    });
  });

  metas.forEach(meta => {
    meta.percentual = meta.avaliados ? +(meta.conformes / meta.avaliados * 100).toFixed(1) : 0;
    meta.itens.forEach(item => {
      item.percentual = item.avaliados ? +(item.conformes / item.avaliados * 100).toFixed(1) : 0;
    });
  });

  const geralAvaliado = geralConformes + geralNaoConformes;
  const conformidadeGeral = geralAvaliado ? +(geralConformes / geralAvaliado * 100).toFixed(1) : 0;
  const diferencaMeta = +(conformidadeGeral - META_INSTITUCIONAL).toFixed(1);

  const evolucaoMensalMap = {};
  linhas
    .filter(row => !anoFiltro || normalizarAno(row[3]) === String(anoFiltro))
    .forEach(row => {
      const mes = normalizarMes(row[2]) || 'Sem mês';
      if (!evolucaoMensalMap[mes]) {
        evolucaoMensalMap[mes] = { conformes: 0, nao: 0 };
      }

      metasDef.forEach(meta => {
        meta.itens.forEach(item => {
          const valor = row[item.idx];
          if (ehSim(valor)) evolucaoMensalMap[mes].conformes++;
          else if (ehNao(valor)) evolucaoMensalMap[mes].nao++;
        });
      });
    });

  const evolucaoMensal = Object.keys(evolucaoMensalMap).map(mes => {
    const d = evolucaoMensalMap[mes];
    const total = d.conformes + d.nao;
    return {
      mes,
      percentual: total ? +(d.conformes / total * 100).toFixed(1) : 0
    };
  });

  const observacoesGerais = linhasFiltradas
    .map(row => String(row[59] || '').trim()) // BH
    .filter(Boolean)
    .slice(0, 20);

  return {
    totalAvaliacoes,
    conformidadeGeral,
    metaInstitucional: META_INSTITUCIONAL,
    diferencaMeta,
    geralConformes,
    geralNaoConformes,
    metas,
    evolucaoMensal,
    observacoesGerais
  };
}

function processarNotificacoes(ss, anoFiltro, mesFiltro, unidadeFiltro) {
  const sh = ss.getSheetByName(ABA_NOTIFICA);
  const dados = sh.getDataRange().getValues();
  const linhas = dados.slice(2);

  const filtradas = linhas.filter(row => {
    const mes = normalizarMes(row[2]);
    const ano = normalizarAno(row[3]);
    const setor = String(row[6] || '').trim();

    if (anoFiltro && ano !== String(anoFiltro)) return false;
    if (mesFiltro && mes !== String(mesFiltro)) return false;
    if (unidadeFiltro && setor && setor !== String(unidadeFiltro)) {
      // não bloquear duro se quiser coexistir com unidade da COSEP
    }

    return true;
  });

  const total = filtradas.length;

  let afetouSim = 0;
  let afetouNao = 0;

  const porTipo = {};
  const porSetor = {};
  const porNatureza = {};
  const porStatus = {};
  const porStatusResposta = {};

  filtradas.forEach(row => {
    const tipo = String(row[8] || 'Não informado').trim();
    const setor = String(row[6] || 'Não informado').trim();
    const natureza = String(row[10] || 'Não informado').trim();
    const status = String(row[13] || 'Não informado').trim();
    const statusResposta = String(row[15] || 'Não informado').trim();
    const afetou = normalizarTexto(row[11]);

    if (afetou === 'SIM') afetouSim++;
    else afetouNao++;

    porTipo[tipo] = (porTipo[tipo] || 0) + 1;
    porSetor[setor] = (porSetor[setor] || 0) + 1;
    porNatureza[natureza] = (porNatureza[natureza] || 0) + 1;
    porStatus[status] = (porStatus[status] || 0) + 1;
    porStatusResposta[statusResposta] = (porStatusResposta[statusResposta] || 0) + 1;
  });

  const percentualAfetouPaciente = total ? +(afetouSim / total * 100).toFixed(1) : 0;

  const tabela = filtradas.slice(0, 100).map(row => ({
    numeroNotivisa: row[0],
    link: row[1],
    mes: row[2],
    ano: row[3],
    codigo: row[4],
    dataOcorrencia: row[5],
    setorNotificado: row[6],
    localOcorrencia: row[7],
    tipoClassificacao: row[8],
    codInteracao: row[9],
    natureza: row[10],
    afetouPaciente: row[11],
    prontuario: row[12],
    status: row[13],
    dataClassificacao: row[14],
    statusResposta: row[15],
    dataResposta: row[16],
    dataConclusao: row[17]
  }));

  return {
    total,
    afetouSim,
    afetouNao,
    percentualAfetouPaciente,
    porTipo,
    porSetor,
    porNatureza,
    porStatus,
    porStatusResposta,
    tabela
  };
}
