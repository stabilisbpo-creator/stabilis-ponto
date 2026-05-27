// =============================================
// CONFIGURAÇÕES DA ESCALA
// =============================================
const ESCALA = {
  entrada: "08:00",
  saidaSegQuinta: "18:00",
  saidaSexta: "17:30",
  intervalo: 90 // minutos
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
      const dtA = converterData(normalizeData(a[0])) + " " + normalizeHora(a[1]);
      const dtB = converterData(normalizeData(b[0])) + " " + normalizeHora(b[1]);
      return dtB.localeCompare(dtA);
    });

  if (registros.length === 0) {
    return { proximoTipo: "Entrada", ultimaData: null };
  }

  const ultimo = registros[0];
  const ultimoTipo = ultimo[3];
  const ultimaData = normalizeData(ultimo[0]);
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

  // Parsear filtro de datas
  const dtInicio = dataInicioStr ? parsearDataISO(dataInicioStr) : null;
  const dtFim = dataFimStr ? parsearDataISO(dataFimStr) : null;

  // Agrupar por funcionário > data
  const mapa = {};
  dados.forEach(r => {
    const dataRaw = r[0];
    const horaRaw = r[1];
    const func = r[2];
    const tipo = r[3];
    if (!dataRaw || !func) return;

    const data = normalizeData(dataRaw);
    const hora = normalizeHora(horaRaw);
    const dataObj = parsearData(data);

    // Aplicar filtro de período
    if (dtInicio && dataObj < dtInicio) return;
    if (dtFim && dataObj > dtFim) return;

    const chave = func + "|" + data;
    if (!mapa[chave]) mapa[chave] = { func, data, dataObj, registros: [] };
    mapa[chave].registros.push({ hora, tipo });
  });

  // Calcular por funcionário
  const porFuncionario = {};
  Object.values(mapa).forEach(dia => {
    const { func, data, dataObj, registros } = dia;
    if (!porFuncionario[func]) porFuncionario[func] = { dias: [], totalTrabalhado: 0, totalEsperado: 0 };

    registros.sort((a, b) => a.hora.localeCompare(b.hora));

    let minutosTrabalhados = 0;
    let i = 0;
    while (i < registros.length) {
      const entrada = registros[i];
      if (entrada.tipo !== "Entrada") { i++; continue; }
      let j = i + 1;
      while (j < registros.length && registros[j].tipo !== "Saída") j++;
      if (j < registros.length) {
        const saida = registros[j];
        const diff = horaParaMinutos(saida.hora) - horaParaMinutos(entrada.hora);
        if (!isNaN(diff)) minutosTrabalhados += diff;
        i = j + 1;
      } else {
        i++;
      }
    }

    const diaSemana = dataObj ? dataObj.getDay() : 1;
    // Sábado (6) e domingo (0) = dia não útil, esperado = 0
    let minutosEsperados = 0;
    if (diaSemana !== 0 && diaSemana !== 6) {
      minutosEsperados = diaSemana === 5
        ? horaParaMinutos(ESCALA.saidaSexta) - horaParaMinutos(ESCALA.entrada)        : horaParaMinutos(ESCALA.saidaSegQuinta) - horaParaMinutos(ESCALA.entrada);
    }

    const saldo = minutosTrabalhados - minutosEsperados;

    // Primeira e última batida do dia
    const primeiroRegistro = registros[0] ? registros[0].hora : "-";
    const ultimoRegistroHora = registros[registros.length - 1] ? registros[registros.length - 1].hora : "-";

    porFuncionario[func].dias.push({
      data: normalizeData(data),
      diaSemana: nomeDiaSemana(diaSemana),
      entrada: normalizeHora(primeiroRegistro),
      saida: normalizeHora(ultimoRegistroHora),
      trabalhado: minutosTrabalhados,
      esperado: minutosEsperados,
      saldo,
      temIncompleto: registros.length % 2 !== 0
    });

    porFuncionario[func].totalTrabalhado += minutosTrabalhados;
    porFuncionario[func].totalEsperado += minutosEsperados;
  });

  // Ordenar dias por data para cada funcionário
  const resultado = Object.entries(porFuncionario).map(([nome, info]) => {
    info.dias.sort((a, b) => {
      const pa = a.data.split("/"); const pb = b.data.split("/");
      const da = new Date(pa[2], pa[1]-1, pa[0]);
      const db = new Date(pb[2], pb[1]-1, pb[0]);
      return da - db;
    });
    return {
      nome,
      totalTrabalhado: info.totalTrabalhado,
      totalEsperado: info.totalEsperado,
      saldoTotal: info.totalTrabalhado - info.totalEsperado,
      dias: info.dias
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
  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();
  const hojeStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");

  const mapa = {};
  dados.forEach(r => {
    const dataRaw = r[0]; const horaRaw = r[1]; const func = r[2]; const tipo = r[3];
    if (!dataRaw || !func) return;
    const data = normalizeData(dataRaw);
    const hora = normalizeHora(horaRaw);
    const chave = func + "|" + data;
    if (!mapa[chave]) mapa[chave] = { func, data, registros: [] };
    mapa[chave].registros.push({ hora, tipo });
  });

  const porFuncionario = {};
  Object.values(mapa).forEach(dia => {
    const { func, data, registros } = dia;
    if (!porFuncionario[func]) porFuncionario[func] = { hoje: null, saldoMes: 0, diasMes: 0, saldoTotal: 0, diasTotal: 0 };

    registros.sort((a, b) => a.hora.localeCompare(b.hora));

    let minutosTrabalhados = 0;
    let i = 0;
    while (i < registros.length) {
      const entrada = registros[i];
      if (entrada.tipo !== "Entrada") { i++; continue; }
      let j = i + 1;
      while (j < registros.length && registros[j].tipo !== "Saída") j++;
      if (j < registros.length) {
        const saida = registros[j];
        const diff = horaParaMinutos(saida.hora) - horaParaMinutos(entrada.hora);
        if (!isNaN(diff)) minutosTrabalhados += diff;
        i = j + 1;
      } else { i++; }
    }

    const dataObj = parsearData(data);
    const diaSemana = dataObj ? dataObj.getDay() : 1;
    let minutosEsperados = 0;
    if (diaSemana !== 0 && diaSemana !== 6) {
      minutosEsperados = diaSemana === 5
        ? horaParaMinutos(ESCALA.saidaSexta) - horaParaMinutos(ESCALA.entrada)        : horaParaMinutos(ESCALA.saidaSegQuinta) - horaParaMinutos(ESCALA.entrada);
    }

    const saldoDia = minutosTrabalhados - minutosEsperados;

    porFuncionario[func].saldoTotal += saldoDia;
    porFuncionario[func].diasTotal += 1;

    if (dataObj && dataObj.getMonth() === mesAtual && dataObj.getFullYear() === anoAtual) {
      porFuncionario[func].saldoMes += saldoDia;
      porFuncionario[func].diasMes += 1;
    }

    if (data === hojeStr) {
      porFuncionario[func].hoje = { trabalhados: minutosTrabalhados, esperados: minutosEsperados, saldo: saldoDia };
    }
  });

  abaPainel.clearContents();
  abaPainel.appendRow(["Funcionário", "Hoje - Trabalhado", "Hoje - Esperado", "Saldo Hoje", "Saldo do Mês", "Dias no Mês", "Saldo Total (Banco)", "Total de Dias"]);

  Object.entries(porFuncionario).forEach(([func, info]) => {
    const h = info.hoje;
    abaPainel.appendRow([
      func,
      h ? minutosParaHora(h.trabalhados) : "-",
      h ? minutosParaHora(h.esperados) : "-",
      h ? formatarSaldo(h.saldo) : "-",
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
// FUNÇÕES AUXILIARES
// =============================================
function normalizeData(dataRaw) {
  if (dataRaw instanceof Date) {
    return Utilities.formatDate(dataRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return String(dataRaw);
}

function normalizeHora(horaRaw) {
  if (horaRaw instanceof Date) {
    return Utilities.formatDate(horaRaw, Session.getScriptTimeZone(), "HH:mm:ss");
  }
  const str = String(horaRaw);
  return str.length === 5 ? str + ":00" : str;
}

function horaParaMinutos(hora) {
  const [h, m] = hora.split(":").map(Number);
  return h * 60 + m;
}

function minutosParaHora(min) {
  const sinal = min < 0 ? "-" : "";
  const abs = Math.abs(min);
  const h = Math.floor(abs / 60);
  const m = abs % 60;
  return sinal + String(h).padStart(2, "0") + "h" + String(m).padStart(2, "0") + "m";
}

function formatarSaldo(min) {
  if (isNaN(min)) return "00h00m";
  const sinal = min >= 0 ? "+" : "";
  return sinal + minutosParaHora(min);
}

function parsearData(dataStr) {
  const partes = dataStr.split("/");
  if (partes.length !== 3) return null;
  return new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
}

function parsearDataISO(iso) {
  // Espera formato YYYY-MM-DD
  const partes = iso.split("-");
  if (partes.length !== 3) return null;
  return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
}

function converterData(dataStr) {
  const partes = dataStr.split("/");
  if (partes.length !== 3) return dataStr;
  return partes[2] + "/" + partes[1] + "/" + partes[0];
}

function nomeDiaSemana(n) {
  return ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"][n] || "";
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
