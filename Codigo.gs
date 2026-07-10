// =============================================
// CONTROLE DE PONTO — PRODUTO GENÉRICO
// Versão 2.0
// =============================================
// ESTRUTURA DE ABAS NECESSÁRIAS:
//   Registros    — batidas de ponto
//   Configurações — escala e regras de horas extras
//   Feriados      — datas especiais editáveis pela empresa
//   Painel        — gerado automaticamente
// =============================================


// =============================================
// CONFIGURAÇÕES PADRÃO (usadas se a aba não existir)
// =============================================
const CONFIG_PADRAO = {
  // Escala
  entrada: "08:00",
  saidaPadrao: "18:00",       // seg-qui ou todos os dias
  saidaSexta: "17:30",        // se houver sexta diferente
  sextaDiferente: true,       // true = sexta tem saída diferente
  intervalo: 90,              // minutos de intervalo (descontados)

  // Horas extras — faixas diurnas (em ordem de aplicação)
  // Cada faixa: { limiteMinutos: X, percentual: Y }
  // limiteMinutos: quantos minutos extras nessa faixa (null = sem limite)
  // percentual: 50 = 50% adicional sobre hora normal
  faixasExtras: [
    { limiteMinutos: 120, percentual: 50 },  // primeiras 2h: 50%
    { limiteMinutos: null, percentual: 100 } // demais: 100%
  ],

  // Extras em feriado
  percentualFeriado: 100,

  // Adicional noturno (após este horário)
  horarioNoturnoInicio: "22:00",
  horarioNoturnoFim: "05:00",
  percentualNoturno: 20,
  usaAdicionalNoturno: false
};


// =============================================
// ENTRADA — GET e POST
// =============================================
function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;

  let resultado;
  if (action === "ultimoRegistro") {
    resultado = ultimoRegistroObj(e.parameter.funcionario);
  } else if (action === "feriados") {
    resultado = { feriados: lerFeriados() };
  } else if (action === "configuracoes") {
    resultado = lerConfiguracoes();
  } else if (action === "ultimosRegistros") {
    resultado = { registros: listarUltimosRegistros(e.parameter.funcionario, 5) };
  } else if (action === "ajustesPendentes") {
    resultado = { ajustes: listarAjustesPendentes() };
  } else {
    resultado = { status: "ok", versao: "2.1" };
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(resultado) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return responder(resultado);
}

function doPost(e) {
  const dados = JSON.parse(e.postData.contents);

  if (dados.action === "registrar") return salvarRegistro(dados);
  if (dados.action === "salvarFeriado") return salvarFeriado(dados);
  if (dados.action === "removerFeriado") return removerFeriado(dados);
  if (dados.action === "salvarConfiguracoes") return salvarConfiguracoes(dados);
  if (dados.action === "pedirAjuste") return pedirAjuste(dados);
  if (dados.action === "responderAjuste") return responderAjuste(dados);

  return responder({ status: "erro", msg: "Ação desconhecida" });
}


// =============================================
// SALVAR REGISTRO
// =============================================
function salvarRegistro(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName("Registros");
  if (!aba) {
    aba = ss.insertSheet("Registros");
    aba.appendRow(["Data", "Hora", "Funcionário", "Tipo", "Justificativa", "Retroativo", "ID", "DataEfetiva", "HoraEfetiva"]);
    aba.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#1a365d").setFontColor("#ffffff");
  }

  const id = Utilities.getUuid();

  // Data/Hora = horário real do dispositivo no momento do clique — NUNCA é editado depois.
  // DataEfetiva/HoraEfetiva = o que entra no cálculo. Começa igual ao real e só muda
  // se um pedido de ajuste for aprovado pelo admin (ver responderAjuste).
  aba.appendRow([
    dados.data,
    dados.hora,
    dados.funcionario,
    dados.tipo,
    dados.justificativa || "",
    dados.retroativo ? "Sim" : "Não",
    id,
    dados.data,
    dados.hora
  ]);

  atualizarPainel();
  return responder({ status: "ok", id });
}


// =============================================
// ÚLTIMO REGISTRO DO FUNCIONÁRIO
// =============================================
function ultimoRegistroObj(funcionario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Registros");

  if (!aba || aba.getLastRow() < 2) {
    return { proximoTipo: "Entrada", ultimaData: null };
  }

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues();

  const registros = dados
    .filter(r => r[2] === funcionario && r[0] !== "")
    .sort((a, b) => {
      const dtA = converterData(formatarData(a[0])) + " " + a[1];
      const dtB = converterData(formatarData(b[0])) + " " + b[1];
      return dtB.localeCompare(dtA);
    });

  if (registros.length === 0) {
    return { proximoTipo: "Entrada", ultimaData: null };
  }

  const ultimo = registros[0];
  const ultimoTipo = ultimo[3];
  const ultimaData = formatarData(ultimo[0]);
  const proximoTipo = ultimoTipo === "Entrada" ? "Saída" : "Entrada";

  return { proximoTipo, ultimaData, ultimoTipo };
}


// =============================================
// FERIADOS
// =============================================
function lerFeriados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Feriados");
  if (!aba || aba.getLastRow() < 2) return [];

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 3).getValues();
  return dados
    .filter(r => r[0] !== "")
    .map(r => ({
      data: formatarData(r[0]),
      descricao: r[1] || "",
      tipo: r[2] || "feriado" // feriado | folga_compensatoria | ponto_facultativo
    }));
}

function salvarFeriado(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName("Feriados");
  if (!aba) {
    aba = ss.insertSheet("Feriados");
    aba.appendRow(["Data", "Descrição", "Tipo"]);
    aba.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1a365d").setFontColor("#ffffff");
  }
  aba.appendRow([dados.data, dados.descricao || "", dados.tipo || "feriado"]);
  return responder({ status: "ok" });
}

function removerFeriado(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Feriados");
  if (!aba || aba.getLastRow() < 2) return responder({ status: "ok" });

  const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
  for (let i = linhas.length - 1; i >= 0; i--) {
    if (formatarData(linhas[i][0]) === dados.data) {
      aba.deleteRow(i + 2);
    }
  }
  return responder({ status: "ok" });
}

function ehFeriado(dataStr) {
  const feriados = lerFeriados();
  return feriados.find(f => f.data === dataStr) || null;
}


// =============================================
// AJUSTES DE PONTO
// O horário real batido (Data/Hora) nunca é alterado.
// Um pedido de ajuste só muda o cálculo (DataEfetiva/HoraEfetiva)
// depois de aprovado pelo admin.
// =============================================
function listarUltimosRegistros(funcionario, limite) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Registros");
  if (!aba || aba.getLastRow() < 2) return [];

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 9).getValues();
  const doFuncionario = dados
    .filter(r => r[2] === funcionario && r[0] !== "" && r[6]) // só registros com ID (versão nova)
    .map(r => ({
      id: r[6],
      data: formatarData(r[0]),
      hora: r[1].toString(),
      tipo: r[3],
      dataEfetiva: formatarData(r[7] || r[0]),
      horaEfetiva: (r[8] || r[1]).toString(),
      ajustado: formatarData(r[7] || r[0]) !== formatarData(r[0]) || (r[8] || r[1]).toString() !== r[1].toString()
    }))
    .reverse();

  return doFuncionario.slice(0, limite || 5);
}

function pedirAjuste(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName("Ajustes");
  if (!aba) {
    aba = ss.insertSheet("Ajustes");
    aba.appendRow(["ID", "RegistroID", "Funcionário", "Data Original", "Hora Original",
      "Data Solicitada", "Hora Solicitada", "Motivo", "Status", "Data Solicitação", "Aprovado Por", "Data Resposta"]);
    aba.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#1a365d").setFontColor("#ffffff");
  }

  const id = Utilities.getUuid();
  const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  aba.appendRow([
    id,
    dados.registroId || "",
    dados.funcionario,
    dados.dataOriginal,
    dados.horaOriginal,
    dados.dataSolicitada,
    dados.horaSolicitada,
    dados.motivo || "",
    "Pendente",
    agora,
    "",
    ""
  ]);

  return responder({ status: "ok", id });
}

function listarAjustesPendentes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Ajustes");
  if (!aba || aba.getLastRow() < 2) return [];

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 12).getValues();
  return dados
    .filter(r => r[0] && r[8] === "Pendente")
    .map(r => ({
      id: r[0],
      registroId: r[1],
      funcionario: r[2],
      dataOriginal: formatarData(r[3]),
      horaOriginal: r[4].toString(),
      dataSolicitada: formatarData(r[5]),
      horaSolicitada: r[6].toString(),
      motivo: r[7],
      dataSolicitacao: r[9]
    }));
}

function responderAjuste(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAjustes = ss.getSheetByName("Ajustes");
  if (!abaAjustes || abaAjustes.getLastRow() < 2) {
    return responder({ status: "erro", msg: "Nenhum ajuste encontrado" });
  }

  const linhas = abaAjustes.getRange(2, 1, abaAjustes.getLastRow() - 1, 12).getValues();
  let linhaEncontrada = -1;
  let registroId, dataSolicitada, horaSolicitada;

  for (let i = 0; i < linhas.length; i++) {
    if (linhas[i][0] === dados.id) {
      linhaEncontrada = i + 2;
      registroId = linhas[i][1];
      dataSolicitada = linhas[i][5];
      horaSolicitada = linhas[i][6];
      break;
    }
  }

  if (linhaEncontrada === -1) return responder({ status: "erro", msg: "Ajuste não encontrado" });

  const status = dados.aprovado ? "Aprovado" : "Rejeitado";
  const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  abaAjustes.getRange(linhaEncontrada, 9).setValue(status);
  abaAjustes.getRange(linhaEncontrada, 11).setValue(dados.aprovadoPor || "Admin");
  abaAjustes.getRange(linhaEncontrada, 12).setValue(agora);

  // Só quando aprovado: atualiza DataEfetiva/HoraEfetiva no registro original.
  // O horário real batido continua intacto nas colunas Data/Hora.
  if (dados.aprovado && registroId) {
    const abaRegistros = ss.getSheetByName("Registros");
    if (abaRegistros && abaRegistros.getLastRow() >= 2) {
      const regs = abaRegistros.getRange(2, 1, abaRegistros.getLastRow() - 1, 9).getValues();
      for (let i = 0; i < regs.length; i++) {
        if (regs[i][6] === registroId) {
          abaRegistros.getRange(i + 2, 8).setValue(dataSolicitada);
          abaRegistros.getRange(i + 2, 9).setValue(horaSolicitada);
          break;
        }
      }
    }
    atualizarPainel();
  }

  return responder({ status: "ok" });
}


// =============================================
// CONFIGURAÇÕES (aba editável)
// =============================================
function lerConfiguracoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Configurações");
  if (!aba || aba.getLastRow() < 2) return CONFIG_PADRAO;

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues();
  const mapa = {};
  dados.forEach(r => {
    if (r[0]) mapa[r[0].toString().trim()] = r[1];
  });

  // Lê faixas de extras (formato JSON na célula)
  let faixas = CONFIG_PADRAO.faixasExtras;
  if (mapa["faixasExtras"]) {
    try { faixas = JSON.parse(mapa["faixasExtras"]); } catch(e) {}
  }

  return {
    entrada: mapa["entrada"] || CONFIG_PADRAO.entrada,
    saidaPadrao: mapa["saidaPadrao"] || CONFIG_PADRAO.saidaPadrao,
    saidaSexta: mapa["saidaSexta"] || CONFIG_PADRAO.saidaSexta,
    sextaDiferente: mapa["sextaDiferente"] === "true" || mapa["sextaDiferente"] === true || CONFIG_PADRAO.sextaDiferente,
    intervalo: Number(mapa["intervalo"]) || CONFIG_PADRAO.intervalo,
    faixasExtras: faixas,
    percentualFeriado: Number(mapa["percentualFeriado"]) || CONFIG_PADRAO.percentualFeriado,
    horarioNoturnoInicio: mapa["horarioNoturnoInicio"] || CONFIG_PADRAO.horarioNoturnoInicio,
    horarioNoturnoFim: mapa["horarioNoturnoFim"] || CONFIG_PADRAO.horarioNoturnoFim,
    percentualNoturno: Number(mapa["percentualNoturno"]) || CONFIG_PADRAO.percentualNoturno,
    usaAdicionalNoturno: mapa["usaAdicionalNoturno"] === "true" || mapa["usaAdicionalNoturno"] === true
  };
}

function salvarConfiguracoes(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName("Configurações");
  if (!aba) {
    aba = ss.insertSheet("Configurações");
    aba.appendRow(["Chave", "Valor", "Descrição"]);
    aba.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1a365d").setFontColor("#ffffff");
  }

  // Limpa e reescreve tudo
  if (aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, 3).clearContent();
  }

  const linhas = [
    ["entrada", dados.entrada, "Horário de entrada padrão"],
    ["saidaPadrao", dados.saidaPadrao, "Saída de seg a qui (ou todos os dias)"],
    ["saidaSexta", dados.saidaSexta, "Saída de sexta-feira"],
    ["sextaDiferente", dados.sextaDiferente ? "true" : "false", "Sexta tem horário diferente?"],
    ["intervalo", dados.intervalo, "Intervalo de almoço em minutos"],
    ["faixasExtras", JSON.stringify(dados.faixasExtras), "Faixas de horas extras (JSON)"],
    ["percentualFeriado", dados.percentualFeriado, "% adicional em feriado (ex: 100)"],
    ["usaAdicionalNoturno", dados.usaAdicionalNoturno ? "true" : "false", "Usa adicional noturno?"],
    ["horarioNoturnoInicio", dados.horarioNoturnoInicio, "Início do adicional noturno"],
    ["horarioNoturnoFim", dados.horarioNoturnoFim, "Fim do adicional noturno"],
    ["percentualNoturno", dados.percentualNoturno, "% adicional noturno (ex: 20)"]
  ];

  linhas.forEach((linha, i) => aba.getRange(i + 2, 1, 1, 3).setValues([linha]));
  aba.autoResizeColumns(1, 3);

  return responder({ status: "ok" });
}


// =============================================
// ATUALIZAR PAINEL
// =============================================
function atualizarPainel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRegistros = ss.getSheetByName("Registros");
  let abaPainel = ss.getSheetByName("Painel");

  if (!abaPainel) abaPainel = ss.insertSheet("Painel");
  if (!abaRegistros || abaRegistros.getLastRow() < 2) return;

  const cfg = lerConfiguracoes();
  const feriados = lerFeriados();
  const feriadoSet = new Set(feriados.map(f => f.data));

  const dados = abaRegistros.getRange(2, 1, abaRegistros.getLastRow() - 1, 9).getValues();
  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();
  const hojeStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Agrupa registros por funcionário + data — usa DataEfetiva/HoraEfetiva (colunas 8 e 9),
  // que refletem ajustes aprovados. Linhas antigas sem essas colunas caem no horário real (fallback).
  const mapa = {};
  dados.forEach(r => {
    if (!r[0] || !r[2]) return;
    const dataStr = formatarData(r[7] || r[0]);
    const horaStr = (r[8] || r[1]).toString();
    const chave = r[2] + "|" + dataStr;
    if (!mapa[chave]) mapa[chave] = { func: r[2], data: dataStr, registros: [] };
    mapa[chave].registros.push({ hora: horaStr, tipo: r[3] });
  });

  // Processa cada funcionário
  const porFuncionario = {};

  Object.values(mapa).forEach(dia => {
    const { func, data, registros } = dia;
    if (!porFuncionario[func]) {
      porFuncionario[func] = {
        hoje: null,
        saldoMes: 0,
        extrasMes: { minutos: 0, detalhes: [] },
        faltas: 0,
        atrasosMin: 0,
        diasMes: 0
      };
    }

    const pf = porFuncionario[func];
    registros.sort((a, b) => a.hora.localeCompare(b.hora));

    const isFeriado = feriadoSet.has(data);
    const dataObj = parsearData(data);
    const diaSemana = dataObj ? dataObj.getDay() : 1;

    // Escala esperada do dia
    const minutosEsperados = calcularMinutosEsperados(diaSemana, cfg, isFeriado);

    // Verifica se é falta (sem nenhum registro válido no dia)
    const temEntrada = registros.some(r => r.tipo === "Entrada");
    const temSaida = registros.some(r => r.tipo === "Saída");

    if (!temEntrada && !temSaida) {
      // Falta total — registra como falta, não como horas negativas
      if (!isFeriado && dataObj &&
          dataObj.getMonth() === mesAtual &&
          dataObj.getFullYear() === anoAtual) {
        pf.faltas += 1;
        // NÃO soma ao saldo de horas — falta é tratada separadamente
      }
      return;
    }

    // Calcula horas trabalhadas (pares entrada/saída)
    let minutosTrabalhados = 0;
    for (let i = 0; i < registros.length - 1; i += 2) {
      const entrada = registros[i];
      const saida = registros[i + 1];
      if (entrada && saida && entrada.tipo === "Entrada" && saida.tipo === "Saída") {
        const minEntrada = horaParaMinutos(entrada.hora);
        const minSaida = horaParaMinutos(saida.hora);
        const bruto = minSaida - minEntrada;
        minutosTrabalhados += Math.max(0, bruto - cfg.intervalo);
      }
    }

    // Calcula atraso (só se entrou, mas ficou menos que o esperado sem ser falta total)
    const saldoDia = minutosTrabalhados - minutosEsperados;

    if (saldoDia < 0 && !isFeriado) {
      pf.atrasosMin += Math.abs(saldoDia);
    }

    // Calcula horas extras e adicional noturno
    const extras = calcularExtras(registros, minutosTrabalhados, minutosEsperados, isFeriado, cfg);

    // Acumula no mês
    if (dataObj && dataObj.getMonth() === mesAtual && dataObj.getFullYear() === anoAtual) {
      pf.saldoMes += saldoDia;
      pf.diasMes += 1;
      if (extras.totalExtraMin > 0) {
        pf.extrasMes.minutos += extras.totalExtraMin;
        pf.extrasMes.detalhes.push({ data, ...extras });
      }
    }

    // Saldo de hoje
    if (data === hojeStr) {
      pf.hoje = {
        trabalhados: minutosTrabalhados,
        esperados: minutosEsperados,
        saldo: saldoDia,
        extras,
        isFeriado
      };
    }
  });

  // Monta o painel
  abaPainel.clearContents();
  const cabecalho = [
    "Funcionário",
    "Hoje — Trabalhado",
    "Hoje — Esperado",
    "Saldo Hoje",
    "Saldo do Mês",
    "H. Extras Mês",
    "Faltas no Mês",
    "Atrasos Mês",
    "Dias Registrados"
  ];
  abaPainel.appendRow(cabecalho);

  Object.entries(porFuncionario).forEach(([func, info]) => {
    const hj = info.hoje;
    abaPainel.appendRow([
      func,
      hj ? minutosParaHora(hj.trabalhados) : "-",
      hj ? minutosParaHora(hj.esperados) : "-",
      hj ? formatarSaldo(hj.saldo) : "-",
      formatarSaldo(info.saldoMes),
      minutosParaHora(info.extrasMes.minutos),
      info.faltas,
      minutosParaHora(info.atrasosMin),
      info.diasMes
    ]);
  });

  abaPainel.getRange(1, 1, 1, cabecalho.length)
    .setFontWeight("bold")
    .setBackground("#1a365d")
    .setFontColor("#ffffff");
  abaPainel.setFrozenRows(1);
  abaPainel.autoResizeColumns(1, cabecalho.length);
}


// =============================================
// CALCULAR MINUTOS ESPERADOS NO DIA
// =============================================
function calcularMinutosEsperados(diaSemana, cfg, isFeriado) {
  if (isFeriado) return 0; // Feriado = nada esperado

  const minEntrada = horaParaMinutos(cfg.entrada);
  let minSaida;

  if (cfg.sextaDiferente && diaSemana === 5) {
    minSaida = horaParaMinutos(cfg.saidaSexta);
  } else {
    minSaida = horaParaMinutos(cfg.saidaPadrao);
  }

  return Math.max(0, minSaida - minEntrada - cfg.intervalo);
}


// =============================================
// CALCULAR HORAS EXTRAS POR FAIXAS
// =============================================
function calcularExtras(registros, minutosTrabalhados, minutosEsperados, isFeriado, cfg) {
  const resultado = {
    totalExtraMin: 0,
    faixas: [],
    adicionalNoturnoMin: 0
  };

  const extraBruto = minutosTrabalhados - minutosEsperados;
  if (extraBruto <= 0 && !isFeriado) return resultado;

  if (isFeriado) {
    // Todo o tempo trabalhado no feriado é extra com percentual específico
    resultado.totalExtraMin = minutosTrabalhados;
    resultado.faixas.push({
      minutos: minutosTrabalhados,
      percentual: cfg.percentualFeriado,
      tipo: "feriado"
    });
  } else {
    // Aplica faixas em ordem
    let minutosRestantes = extraBruto;
    cfg.faixasExtras.forEach(faixa => {
      if (minutosRestantes <= 0) return;
      const nessaFaixa = faixa.limiteMinutos
        ? Math.min(minutosRestantes, faixa.limiteMinutos)
        : minutosRestantes;
      resultado.faixas.push({
        minutos: nessaFaixa,
        percentual: faixa.percentual,
        tipo: "diurno"
      });
      resultado.totalExtraMin += nessaFaixa;
      minutosRestantes -= nessaFaixa;
    });
  }

  // Adicional noturno
  if (cfg.usaAdicionalNoturno) {
    const inicioNoturno = horaParaMinutos(cfg.horarioNoturnoInicio);
    let minutosNoturnos = 0;
    for (let i = 0; i < registros.length - 1; i += 2) {
      const ent = registros[i];
      const sai = registros[i + 1];
      if (ent && sai && ent.tipo === "Entrada" && sai.tipo === "Saída") {
        const minSai = horaParaMinutos(sai.hora);
        if (minSai > inicioNoturno) {
          minutosNoturnos += minSai - Math.max(horaParaMinutos(ent.hora), inicioNoturno);
        }
      }
    }
    resultado.adicionalNoturnoMin = minutosNoturnos;
  }

  return resultado;
}


// =============================================
// FUNÇÃO PARA POPULAR PAINEL MANUALMENTE
// Execute esta função para recalcular tudo
// =============================================
function recalcularTudo() {
  atualizarPainel();
  Logger.log("Painel atualizado com sucesso.");
}


// =============================================
// MIGRAÇÃO — execute UMA VEZ após subir esta versão
// Preenche ID/DataEfetiva/HoraEfetiva nas linhas antigas
// que foram criadas antes da função de ajuste existir.
// =============================================
function migrarRegistrosAntigos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Registros");
  if (!aba || aba.getLastRow() < 2) {
    Logger.log("Nenhum registro para migrar.");
    return;
  }

  const numLinhas = aba.getLastRow() - 1;
  const dados = aba.getRange(2, 1, numLinhas, 9).getValues();
  let migradas = 0;

  dados.forEach((r, i) => {
    const linha = i + 2;
    if (!r[6]) { aba.getRange(linha, 7).setValue(Utilities.getUuid()); migradas++; }
    if (!r[7]) aba.getRange(linha, 8).setValue(r[0]);
    if (!r[8]) aba.getRange(linha, 9).setValue(r[1]);
  });

  Logger.log("Migração concluída: " + migradas + " linha(s) receberam ID novo, de " + numLinhas + " verificadas.");
}


// =============================================
// CRIAR ESTRUTURA INICIAL DA PLANILHA
// Execute uma vez ao configurar
// =============================================
function inicializarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Aba Registros
  let abaReg = ss.getSheetByName("Registros");
  if (!abaReg) {
    abaReg = ss.insertSheet("Registros");
    abaReg.appendRow(["Data", "Hora", "Funcionário", "Tipo", "Justificativa", "Retroativo", "ID", "DataEfetiva", "HoraEfetiva"]);
    abaReg.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#1a365d").setFontColor("#ffffff");
  }

  // Aba Ajustes
  let abaAju = ss.getSheetByName("Ajustes");
  if (!abaAju) {
    abaAju = ss.insertSheet("Ajustes");
    abaAju.appendRow(["ID", "RegistroID", "Funcionário", "Data Original", "Hora Original",
      "Data Solicitada", "Hora Solicitada", "Motivo", "Status", "Data Solicitação", "Aprovado Por", "Data Resposta"]);
    abaAju.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#1a365d").setFontColor("#ffffff");
  }

  // Aba Feriados
  let abaFer = ss.getSheetByName("Feriados");
  if (!abaFer) {
    abaFer = ss.insertSheet("Feriados");
    abaFer.appendRow(["Data", "Descrição", "Tipo"]);
    abaFer.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1a365d").setFontColor("#ffffff");
    // Feriados nacionais 2026
    var feriadosNacionais2026 = [
      // FERIADOS NACIONAIS
      ["01/01/2026", "Confraternização Universal", "feriado"],
      ["03/04/2026", "Paixão de Cristo", "feriado"],
      ["21/04/2026", "Tiradentes", "feriado"],
      ["01/05/2026", "Dia Mundial do Trabalho", "feriado"],
      ["07/09/2026", "Independência do Brasil", "feriado"],
      ["12/10/2026", "Nossa Senhora Aparecida", "feriado"],
      ["02/11/2026", "Finados", "feriado"],
      ["15/11/2026", "Proclamação da República", "feriado"],
      ["20/11/2026", "Consciência Negra", "feriado"],
      ["25/12/2026", "Natal", "feriado"],
      // PONTOS FACULTATIVOS (empresa decide se adota — pode remover os que não usar)
      ["16/02/2026", "Carnaval", "ponto_facultativo"],
      ["17/02/2026", "Carnaval", "ponto_facultativo"],
      ["18/02/2026", "Quarta-Feira de Cinzas (até 14h)", "ponto_facultativo"],
      ["04/06/2026", "Corpus Christi", "ponto_facultativo"],
      ["05/06/2026", "Emenda Corpus Christi", "ponto_facultativo"],
      ["28/10/2026", "Dia do Servidor Público", "ponto_facultativo"],
      ["24/12/2026", "Véspera de Natal", "ponto_facultativo"],
      ["31/12/2026", "Véspera do Ano Novo", "ponto_facultativo"]
    ];
    feriadosNacionais2026.forEach(function(f) { abaFer.appendRow(f); });
  }

  // Aba Configurações
  let abaCfg = ss.getSheetByName("Configurações");
  if (!abaCfg) {
    salvarConfiguracoes(CONFIG_PADRAO);
  }

  Logger.log("Planilha inicializada com sucesso!");
}


// =============================================
// FUNÇÕES AUXILIARES
// =============================================
function horaParaMinutos(hora) {
  if (!hora) return 0;
  const partes = hora.toString().split(":");
  return parseInt(partes[0]) * 60 + parseInt(partes[1] || 0);
}

function minutosParaHora(min) {
  const sinal = min < 0 ? "-" : "";
  const abs = Math.abs(min);
  const h = Math.floor(abs / 60);
  const m = abs % 60;
  return sinal + String(h).padStart(2, "0") + "h" + String(m).padStart(2, "0") + "m";
}

function formatarSaldo(min) {
  const sinal = min >= 0 ? "+" : "";
  return sinal + minutosParaHora(min);
}

function parsearData(dataStr) {
  if (!dataStr) return null;
  const partes = dataStr.toString().split("/");
  if (partes.length !== 3) return null;
  return new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
}

function converterData(dataStr) {
  const partes = dataStr.split("/");
  if (partes.length !== 3) return dataStr;
  return partes[2] + "/" + partes[1] + "/" + partes[0];
}

function formatarData(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const d = String(val.getDate()).padStart(2, "0");
    const m = String(val.getMonth() + 1).padStart(2, "0");
    return d + "/" + m + "/" + val.getFullYear();
  }
  const str = val.toString().trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return str;
  if (str.includes("GMT") || str.includes("UTC")) {
    const dt = new Date(str);
    if (!isNaN(dt.getTime())) {
      const d = String(dt.getUTCDate()).padStart(2, "0");
      const m = String(dt.getUTCMonth() + 1).padStart(2, "0");
      return d + "/" + m + "/" + dt.getUTCFullYear();
    }
  }
  return str;
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
