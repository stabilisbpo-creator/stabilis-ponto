// =============================================
// CONFIGURAÇÕES DA ESCALA
// =============================================
const ESCALA = {
  entrada: "08:00",
  saidaSegQuinta: "18:00",
  saidaSexta: "17:30"
  // Sem intervalo — o almoço já está embutido no horário
};

// =============================================
// FUNÇÃO PRINCIPAL — recebe GET
// =============================================
function doGet(e) {
  const callback = e.parameter.callback;
  const action = e.parameter.action;
  let resultado;

  if (action === "ultimoRegistro") {
    resultado = ultimoRegistroObj(e.parameter.funcionario);
  } else if (action === "dadosPainel") {
    resultado = getDadosPainel(e.parameter.dataInicio, e.parameter.dataFim);
  } else {
    resultado = { status: "ok" };
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
  if (dados.action === "registrar") {
    return salvarRegistro(dados);
  }
  return responder({ status: "erro", msg: "Ação desconhecida" });
}

// =============================================
// SALVAR REGISTRO NA ABA "Registros"
// =============================================
function salvarRegistro(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName("Registros");
  if (!aba) {
    aba = ss.insertSheet("Registros");
    aba.appendRow(["Data", "Hora", "Funcionário", "Tipo", "Justificativa", "Retroativo"]);
  }

  aba.appendRow([
    dados.data,
    dados.hora,
    dados.funcionario,
    dados.tipo,
    dados.justificativa || "",
    dados.retroativo ? "Sim" : "Não"
  ]);

  atualizarPainel();
  return responder({ status: "ok" });
}

// =============================================
// CONSULTAR ÚLTIMO REGISTRO DO FUNCIONÁRIO
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
      const dtA = normalizeDataStr(a[0]) + " " + normalizeHoraStr(a[1]);
      const dtB = normalizeDataStr(b[0]) + " " + normalizeHoraStr(b[1]);
      return dtB.localeCompare(dtA);
    });

  if (registros.length === 0) {
    return { proximoTipo: "Entrada", ultimaData: null };
  }

  const ultimo = registros[0];
  const ultimoTipo = ultimo[3];
  const ultimaData = normalizeDataFormatado(ultimo[0]);
  const proximoTipo = ultimoTipo === "Entrada" ? "Saída" : "Entrada";

  return { proximoTipo, ultimaData, ultimoTipo };
}

function ultimoRegistro(funcionario) {
  return responder(ultimoRegistroObj(funcionario));
}

// =============================================
// DADOS PARA O PAINEL WEB (com filtro de período)
// =============================================
function getDadosPainel(dataInicioStr, dataFimStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRegistros = ss.getSheetByName("Registros");

  if (!abaRegistros || abaRegistros.getLastRow() < 2) {
    return { funcionarios: [] };
  }

  const dados = abaRegistros.getRange(2, 1, abaRegistros.getLastRow() - 1, 6).getValues();

  const dtInicio = dataInicioStr ? parsearDataISO(dataInicioStr) : null;
  const dtFim    = dataFimStr   ? parsearDataISO(dataFimStr)    : null;

  // Agrupar por funcionário + data
  const mapa = {};
  dados.forEach(r => {
    const dataRaw = r[0];
    const horaRaw = r[1];
    const func    = r[2];
    const tipo    = r[3];
    if (!dataRaw || !func) return;

    const dataFormatada = normalizeDataFormatado(dataRaw); // dd/MM/yyyy
    const horaFormatada = normalizeHoraFormatado(horaRaw); // HH:mm:ss
    const dataUTC       = parsearDataDDMMYYYY(dataFormatada);

    if (dtInicio && dataUTC < dtInicio) return;
    if (dtFim    && dataUTC > dtFim)    return;

    const chave = func + "|" + dataFormatada;
    if (!mapa[chave]) mapa[chave] = { func, dataFormatada, dataUTC, registros: [] };
    mapa[chave].registros.push({ hora: horaFormatada, tipo });
  });

  // Calcular por funcionário
  const porFuncionario = {};
  Object.values(mapa).forEach(dia => {
    const { func, dataFormatada, dataUTC, registros } = dia;
    if (!porFuncionario[func]) porFuncionario[func] = { dias: [], totalTrabalhado: 0, totalEsperado: 0 };

    registros.sort((a, b) => a.hora.localeCompare(b.hora));

    // Dia da semana via UTC para evitar deslocamento de fuso
    const diaSemana = dataUTC.getUTCDay(); // 0=Dom, 1=Seg ... 6=Sáb
    let minutosEsperados = 0;
    if (diaSemana !== 0 && diaSemana !== 6) {
      minutosEsperados = diaSemana === 5
        ? horaParaMinutos(ESCALA.saidaSexta)     - horaParaMinutos(ESCALA.entrada)
        : horaParaMinutos(ESCALA.saidaSegQuinta) - horaParaMinutos(ESCALA.entrada);
    }

    // Monta pares Entrada+Saída — cada par vira uma linha no painel
    const pares = [];
    let minutosTrabalhados = 0;
    let i = 0;
    while (i < registros.length) {
      if (registros[i].tipo !== "Entrada") { i++; continue; }
      let j = i + 1;
      while (j < registros.length && registros[j].tipo !== "Saída") j++;
      if (j < registros.length) {
        const diff = horaParaMinutos(registros[j].hora) - horaParaMinutos(registros[i].hora);
        const minPar = (!isNaN(diff) && diff > 0) ? diff : 0;
        minutosTrabalhados += minPar;
        pares.push({ entrada: registros[i].hora, saida: registros[j].hora, trabalhado: minPar });
        i = j + 1;
      } else {
        // Entrada sem saída
        pares.push({ entrada: registros[i].hora, saida: "-", trabalhado: 0, incompleto: true });
        i++;
      }
    }

    const saldo = minutosTrabalhados - minutosEsperados;

    porFuncionario[func].dias.push({
      data:          dataFormatada,
      diaSemana:     nomeDiaSemana(diaSemana),
      pares,
      trabalhado:    minutosTrabalhados,
      esperado:      minutosEsperados,
      saldo,
      temIncompleto: registros.length % 2 !== 0
    });

    porFuncionario[func].totalTrabalhado += minutosTrabalhados;
    porFuncionario[func].totalEsperado   += minutosEsperados;
  });

  // Montar resultado ordenado por data
  const resultado = Object.entries(porFuncionario).map(([nome, info]) => {
    info.dias.sort((a, b) => {
      return parsearDataDDMMYYYY(a.data) - parsearDataDDMMYYYY(b.data);
    });
    return {
      nome,
      totalTrabalhado: info.totalTrabalhado,
      totalEsperado:   info.totalEsperado,
      saldoTotal:      info.totalTrabalhado - info.totalEsperado,
      dias:            info.dias
    };
  });

  resultado.sort((a, b) => a.nome.localeCompare(b.nome));
  return { funcionarios: resultado };
}

// =============================================
// ATUALIZAR ABA "Painel" (planilha interna)
// =============================================
function atualizarPainel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRegistros = ss.getSheetByName("Registros");
  let abaPainel = ss.getSheetByName("Painel");

  if (!abaPainel) abaPainel = ss.insertSheet("Painel");
  if (abaRegistros.getLastRow() < 2) return;

  const dados = abaRegistros.getRange(2, 1, abaRegistros.getLastRow() - 1, 6).getValues();
  const hoje    = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();
  const hojeStr  = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");

  const mapa = {};
  dados.forEach(r => {
    const dataRaw = r[0]; const horaRaw = r[1]; const func = r[2]; const tipo = r[3];
    if (!dataRaw || !func) return;
    const dataFormatada = normalizeDataFormatado(dataRaw);
    const horaFormatada = normalizeHoraFormatado(horaRaw);
    const chave = func + "|" + dataFormatada;
    if (!mapa[chave]) mapa[chave] = { func, dataFormatada, registros: [] };
    mapa[chave].registros.push({ hora: horaFormatada, tipo });
  });

  const porFuncionario = {};
  Object.values(mapa).forEach(dia => {
    const { func, dataFormatada, registros } = dia;
    if (!porFuncionario[func]) porFuncionario[func] = { hoje: null, saldoMes: 0, diasMes: 0, saldoTotal: 0, diasTotal: 0 };

    registros.sort((a, b) => a.hora.localeCompare(b.hora));

    let minutosTrabalhados = 0;
    let i = 0;
    while (i < registros.length) {
      if (registros[i].tipo !== "Entrada") { i++; continue; }
      let j = i + 1;
      while (j < registros.length && registros[j].tipo !== "Saída") j++;
      if (j < registros.length) {
        const diff = horaParaMinutos(registros[j].hora) - horaParaMinutos(registros[i].hora);
        if (!isNaN(diff) && diff > 0) minutosTrabalhados += diff;
        i = j + 1;
      } else { i++; }
    }

    const dataUTC   = parsearDataDDMMYYYY(dataFormatada);
    const diaSemana = dataUTC.getUTCDay();
    let minutosEsperados = 0;
    if (diaSemana !== 0 && diaSemana !== 6) {
      minutosEsperados = diaSemana === 5
        ? horaParaMinutos(ESCALA.saidaSexta)     - horaParaMinutos(ESCALA.entrada)
        : horaParaMinutos(ESCALA.saidaSegQuinta) - horaParaMinutos(ESCALA.entrada);
    }

    const saldoDia = minutosTrabalhados - minutosEsperados;
    porFuncionario[func].saldoTotal += saldoDia;
    porFuncionario[func].diasTotal  += 1;

    if (dataUTC.getUTCMonth() === mesAtual && dataUTC.getUTCFullYear() === anoAtual) {
      porFuncionario[func].saldoMes += saldoDia;
      porFuncionario[func].diasMes  += 1;
    }

    if (dataFormatada === hojeStr) {
      porFuncionario[func].hoje = { trabalhados: minutosTrabalhados, esperados: minutosEsperados, saldo: saldoDia };
    }
  });

  abaPainel.clearContents();
  abaPainel.appendRow(["Funcionário","Hoje - Trabalhado","Hoje - Esperado","Saldo Hoje","Saldo do Mês","Dias no Mês","Saldo Total (Banco)","Total de Dias"]);

  Object.entries(porFuncionario).forEach(([func, info]) => {
    const h = info.hoje;
    abaPainel.appendRow([
      func,
      h ? minutosParaHora(h.trabalhados) : "-",
      h ? minutosParaHora(h.esperados)   : "-",
      h ? formatarSaldo(h.saldo)         : "-",
      formatarSaldo(info.saldoMes),
      info.diasMes,
      formatarSaldo(info.saldoTotal),
      info.diasTotal
    ]);
  });

  abaPainel.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#1a365d").setFontColor("#ffffff");
  abaPainel.setFrozenRows(1);
  abaPainel.autoResizeColumns(1, 8);
}

// =============================================
// FUNÇÕES AUXILIARES DE NORMALIZAÇÃO
// =============================================

// Retorna data no formato dd/MM/yyyy — aceita Date, ISO string ou dd/MM/yyyy
function normalizeDataFormatado(dataRaw) {
  if (dataRaw instanceof Date) {
    return Utilities.formatDate(dataRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  const str = String(dataRaw).trim();
  // Formato ISO: "2026-05-19T03:00:00.000Z" ou "2026-05-19"
  if (str.indexOf("T") !== -1 || /^\d{4}-\d{2}-\d{2}$/.test(str)) {
    const datePart = str.split("T")[0];
    const p = datePart.split("-");
    if (p.length === 3) return p[2] + "/" + p[1] + "/" + p[0];
  }
  return str; // já está em dd/MM/yyyy
}

// Retorna hora no formato HH:mm:ss — aceita Date, ISO string ou HH:mm(:ss)
function normalizeHoraFormatado(horaRaw) {
  if (horaRaw instanceof Date) {
    return Utilities.formatDate(horaRaw, Session.getScriptTimeZone(), "HH:mm:ss");
  }
  const str = String(horaRaw).trim();
  // Formato ISO: "1899-12-31T21:47:21.000Z"
  if (str.indexOf("T") !== -1) {
    const timePart = str.split("T")[1].replace("Z","").split(".")[0]; // HH:mm:ss em UTC
    // Hora de 1899 no Sheets = hora pura armazenada como UTC, sem fuso
    return timePart;
  }
  // HH:mm → HH:mm:ss
  if (/^\d{2}:\d{2}$/.test(str)) return str + ":00";
  return str;
}

// Versão para ordenação (yyyy/MM/dd HH:mm:ss)
function normalizeDataStr(dataRaw) {
  const fmt = normalizeDataFormatado(dataRaw);
  const p = fmt.split("/");
  return p.length === 3 ? p[2] + "/" + p[1] + "/" + p[0] : fmt;
}

function normalizeHoraStr(horaRaw) {
  return normalizeHoraFormatado(horaRaw);
}

// =============================================
// FUNÇÕES AUXILIARES DE CÁLCULO
// =============================================

function horaParaMinutos(hora) {
  if (!hora || hora === "-") return 0;
  const partes = hora.split(":");
  return Number(partes[0]) * 60 + Number(partes[1]);
}

function minutosParaHora(min) {
  if (isNaN(min)) return "00h00m";
  const sinal = min < 0 ? "-" : "";
  const abs = Math.abs(min);
  return sinal + String(Math.floor(abs / 60)).padStart(2, "0") + "h" + String(abs % 60).padStart(2, "0") + "m";
}

function formatarSaldo(min) {
  if (isNaN(min)) return "+00h00m";
  return (min >= 0 ? "+" : "") + minutosParaHora(min);
}

// Parseia dd/MM/yyyy → Date UTC (sem deslocamento de fuso)
function parsearDataDDMMYYYY(dataStr) {
  const p = dataStr.split("/");
  if (p.length !== 3) return new Date(0);
  return new Date(Date.UTC(Number(p[2]), Number(p[1]) - 1, Number(p[0])));
}

// Parseia YYYY-MM-DD → Date UTC
function parsearDataISO(iso) {
  const p = iso.split("-");
  if (p.length !== 3) return null;
  return new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2])));
}

function nomeDiaSemana(n) {
  return ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"][n] || "";
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
