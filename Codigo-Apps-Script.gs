/* ==================================================================
   FORMULÁRIO CRIO HARMONIZAÇÃO — Amanda Sobottka
   Google Apps Script — grava leads na planilha "Leads", mantém um
   painel automático na aba "Painel" e organiza um controle de status
   comercial (Novo / Em contato / Fechou / etc.) direto na planilha.

   COMO USAR (resumo — o passo a passo completo está no guia de instalação):
   1. Crie uma planilha nova no Google Sheets.
   2. Menu Extensões > Apps Script.
   3. Apague o código de exemplo e cole TODO este arquivo no lugar.
   4. Salve (ícone de disquete ou Ctrl+S).
   5. No menu de funções (barra de cima, ao lado do ícone de play),
      selecione "criarPlanilha" e clique em Executar. Autorize as
      permissões quando pedido. Isso cria/organiza as abas "Leads" e
      "Painel" já com cabeçalhos, cores e fórmulas prontas.
   6. Implantar > Nova implantação > tipo "Aplicativo da Web".
      - Executar como: Eu (seu e-mail)
      - Quem tem acesso: Qualquer pessoa
   7. Copie a URL gerada e cole na variável APPS_SCRIPT_URL no topo
      do arquivo HTML do formulário.

   OBS: sempre que colar uma versão nova deste script, rode
   "criarPlanilha" de novo — ela atualiza cabeçalhos, cores, dropdown
   de status e fórmulas do Painel, sem apagar os leads já gravados.
   ================================================================== */

const SHEET_LEADS = "Leads";
const SHEET_PAINEL = "Painel";

// Ordem das colunas na aba Leads
// A=Data/Hora B=Nome C=Sobrenome D=Telefone E=Cidade F=Objetivo
// G=Região(ões) H=Rotina I=Quando pretende iniciar J=Consentimento K=Status
// L=Observações M=Próximo Contato
const CABECALHO_LEADS = [
  "Data/Hora", "Nome", "Sobrenome", "Telefone", "Cidade",
  "Objetivo", "Região(ões)", "Rotina", "Quando pretende iniciar", "Consentimento", "Status",
  "Observações", "Próximo Contato"
];
const COL_STATUS = 11; // coluna K
const COL_OBSERVACOES = 12; // coluna L
const COL_PROXIMO_CONTATO = 13; // coluna M

// Rótulos usados nas telas do formulário — precisam bater com o HTML
const OBJETIVOS = [
  "Tenho uma gordura localizada que não sai nem com esforço",
  "Quero emagrecer / estou começando a cuidar do corpo",
  "Ainda estou pesquisando",
];

const REGIOES = [
  "Abdômen", "Flanco", "Culote", "Coxa (interna/externa)",
  "Braço", "Papada", "Costas", "Quero uma avaliação geral",
];

const OPCOES_INICIO = [
  "Assim que possível / já quero começar",
  "Nas próximas semanas",
  "Ainda estou decidindo / só pesquisando",
];

// Etapas do processo comercial — aparecem como lista suspensa na coluna Status
const STATUS_OPCOES = [
  "Novo",
  "Em contato",
  "Avaliação agendada",
  "Fechou",
  "Não fechou",
  "Sem resposta",
];

// Cor de fundo de cada status, pra dar uma leitura visual rápida na planilha
const STATUS_CORES = {
  "Novo": "#F1F3F4",
  "Em contato": "#FFF2CC",
  "Avaliação agendada": "#C9DAF8",
  "Fechou": "#D9EAD3",
  "Não fechou": "#F4CCCC",
  "Sem resposta": "#FCE5CD",
};

// Senha simples do painel de leads (HTML separado, tipo quadro/Kanban).
// TROQUE esse valor antes de divulgar o link do painel pra atendente.
const SENHA_PAINEL = "trocar-senha-aqui";

/* ------------------------------------------------------------------
   doPost — recebe cada envio do formulário e grava uma linha em "Leads"
   ------------------------------------------------------------------ */
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_LEADS);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_LEADS);
      sheet.appendRow(CABECALHO_LEADS);
    }

    sheet.appendRow([
      new Date(),
      dados.nome || "",
      dados.sobrenome || "",
      dados.telefone || "",
      dados.cidade || "",
      dados.objetivo || "",
      dados.regioes || "",
      dados.rotina || "",
      dados.inicio || "",
      dados.consentimento ? "Sim" : "Não",
      "Novo", // status inicial — a atendente vai atualizando conforme fala com o lead
      "", // Observações — preenchida depois, pelo painel
      "", // Próximo Contato — preenchida depois, pelo painel
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ------------------------------------------------------------------
   doGet — endpoint de leitura. Sem parâmetros: só confirma que o
   Web App está ativo (usado pra testar a URL no navegador). Com
   ?acao=listar ou ?acao=atualizarStatus: alimenta o painel de leads
   (HTML separado, tipo quadro/Kanban), sempre validando a senha e
   respondendo em formato JSONP (callback) pra funcionar de qualquer
   origem sem problema de CORS.
   ------------------------------------------------------------------ */
function doGet(e) {
  const acao = e.parameter && e.parameter.acao;
  const callback = e.parameter && e.parameter.callback;

  if (acao === "listar") {
    return respostaJson(listarLeadsParaPainel(e), callback);
  }
  if (acao === "atualizarStatus") {
    return respostaJson(atualizarStatusLead(e), callback);
  }
  if (acao === "excluir") {
    return respostaJson(excluirLead(e), callback);
  }
  if (acao === "atualizarObservacoes") {
    return respostaJson(atualizarObservacoesLead(e), callback);
  }

  return ContentService.createTextOutput("Formulário Crio Harmonização — endpoint ativo.");
}

/* Empacota a resposta em JSON puro ou em JSONP (quando vem "callback"
   na URL — necessário pro painel HTML ler os dados sem bloqueio de CORS). */
function respostaJson(obj, callback) {
  const texto = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + texto + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

function senhaValida(e) {
  return !!(e.parameter && e.parameter.senha === SENHA_PAINEL);
}

/* Formata uma data como YYYY-MM-DD (mesmo formato do <input type="date">
   do painel), usando o fuso horário da própria planilha. */
function formatarDataISO(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  } catch (err) {
    return "";
  }
}

/* Devolve todos os leads da aba "Leads" em JSON, cada um com o número
   da linha na planilha (precisa pra depois atualizar o status certo). */
function listarLeadsParaPainel(e) {
  if (!senhaValida(e)) {
    return { status: "error", message: "Senha incorreta." };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LEADS);
    if (!sheet) return { status: "ok", leads: [], statusOpcoes: STATUS_OPCOES };

    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) return { status: "ok", leads: [], statusOpcoes: STATUS_OPCOES };

    const valores = sheet.getRange(2, 1, ultimaLinha - 1, CABECALHO_LEADS.length).getValues();
    const leads = [];
    valores.forEach(function (linha, i) {
      if (!linha[1]) return; // pula linhas em branco
      leads.push({
        linha: i + 2,
        dataHora: linha[0] instanceof Date ? linha[0].toISOString() : String(linha[0]),
        nome: linha[1],
        sobrenome: linha[2],
        telefone: linha[3],
        cidade: linha[4],
        objetivo: linha[5],
        regioes: linha[6],
        rotina: linha[7],
        inicio: linha[8],
        consentimento: linha[9],
        status: linha[10] || "Novo",
        observacoes: linha[11] || "",
        proximoContato: linha[12] instanceof Date ? formatarDataISO(linha[12]) : String(linha[12] || ""),
      });
    });

    return { status: "ok", leads: leads, statusOpcoes: STATUS_OPCOES };
  } catch (err) {
    return { status: "error", message: String(err) };
  }
}

/* Atualiza o status de um lead específico (identificado pelo número da
   linha) — usado quando a atendente arrasta/seleciona um novo status
   no painel HTML. */
function atualizarStatusLead(e) {
  if (!senhaValida(e)) {
    return { status: "error", message: "Senha incorreta." };
  }
  try {
    const linha = parseInt(e.parameter.linha, 10);
    const novoStatus = e.parameter.novoStatus;
    if (!linha || STATUS_OPCOES.indexOf(novoStatus) === -1) {
      return { status: "error", message: "Dados inválidos." };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LEADS);
    sheet.getRange(linha, COL_STATUS).setValue(novoStatus);
    return { status: "ok" };
  } catch (err) {
    return { status: "error", message: String(err) };
  }
}

/* Salva a observação livre e a data do próximo contato de um lead —
   preenchidas pela atendente ao clicar no card, no painel HTML. A data
   chega como texto "YYYY-MM-DD" (do <input type="date">) e é gravada
   como texto puro, sem deixar o Sheets tentar reformatar sozinho. */
function atualizarObservacoesLead(e) {
  if (!senhaValida(e)) {
    return { status: "error", message: "Senha incorreta." };
  }
  try {
    const linha = parseInt(e.parameter.linha, 10);
    if (!linha || linha < 2) {
      return { status: "error", message: "Linha inválida." };
    }
    const observacoes = e.parameter.observacoes || "";
    const proximoContato = e.parameter.proximoContato || "";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LEADS);
    sheet.getRange(linha, COL_OBSERVACOES).setValue(observacoes);
    sheet.getRange(linha, COL_PROXIMO_CONTATO).setNumberFormat("@").setValue(proximoContato);
    return { status: "ok" };
  } catch (err) {
    return { status: "error", message: String(err) };
  }
}

/* Exclui de vez a linha de um lead (ex.: cadastro de teste, duplicado,
   spam). Ação irreversível — a confirmação acontece no painel HTML. */
function excluirLead(e) {
  if (!senhaValida(e)) {
    return { status: "error", message: "Senha incorreta." };
  }
  try {
    const linha = parseInt(e.parameter.linha, 10);
    if (!linha || linha < 2) {
      return { status: "error", message: "Linha inválida." };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LEADS);
    sheet.deleteRow(linha);
    return { status: "ok" };
  } catch (err) {
    return { status: "error", message: String(err) };
  }
}

/* ------------------------------------------------------------------
   getArgSep — detecta o idioma da planilha e devolve o separador de
   argumentos certo pra fórmula ("," no inglês, ";" em português,
   espanhol, alemão, francês etc. — idiomas que usam vírgula como
   separador decimal). Sem isso, as fórmulas do Painel dão #ERROR!
   em planilhas configuradas em português.
   ------------------------------------------------------------------ */
function getArgSep() {
  var locale = "";
  try {
    locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale() || "";
  } catch (e) {
    locale = "";
  }
  var idioma = locale.split(/[_-]/)[0].toLowerCase();
  var idiomasComVirgulaDecimal = ["pt", "es", "de", "fr", "it", "nl", "pl", "ru", "tr", "sv", "da", "fi", "nb", "no", "cs", "sk", "ro", "hu", "el", "uk"];
  return idiomasComVirgulaDecimal.indexOf(idioma) !== -1 ? ";" : ",";
}

/* ------------------------------------------------------------------
   criarPlanilha — RODE ESTA FUNÇÃO sempre que colar uma versão nova
   do script. Cria/organiza as abas "Leads" e "Painel": cabeçalho,
   cores, dropdown de status e fórmulas — sem apagar leads já gravados.
   ------------------------------------------------------------------ */
function criarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SEP = getArgSep(); // "," ou ";" dependendo do idioma da planilha
  const LIMITE = 5000; // linhas cobertas pelas fórmulas e formatações (aumente se precisar)

  organizarAbaLeads(ss, LIMITE);
  organizarAbaPainel(ss, SEP, LIMITE);

  try {
    SpreadsheetApp.getUi().alert(
      "Pronto! As abas \"Leads\" e \"Painel\" foram criadas/atualizadas."
    );
  } catch (e) {
    Logger.log("Concluído. (Não foi possível mostrar o alerta de UI, mas as abas foram criadas normalmente.)");
  }
}

/* ------------------------------------------------------------------
   organizarAbaLeads — cabeçalho colorido, coluna de status com lista
   suspensa e cores, zebra nas linhas, e preenche "Novo" em quem ainda
   não tem status (leads gravados antes dessa atualização).
   ------------------------------------------------------------------ */
function organizarAbaLeads(ss, LIMITE) {
  let leads = ss.getSheetByName(SHEET_LEADS);
  if (!leads) {
    leads = ss.insertSheet(SHEET_LEADS);
  }

  // Cabeçalho
  const numCols = CABECALHO_LEADS.length;
  const headerRange = leads.getRange(1, 1, 1, numCols);
  headerRange.setValues([CABECALHO_LEADS]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#A8663F");
  headerRange.setFontColor("#FFFFFF");
  leads.setFrozenRows(1);

  // Zebra (linhas alternadas) pra facilitar a leitura
  leads.getBandings().forEach(function (b) { b.remove(); });
  leads.getRange(1, 1, LIMITE, numCols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  // Lista suspensa (dropdown) na coluna Status
  const statusRange = leads.getRange(2, COL_STATUS, LIMITE - 1, 1);
  const regra = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPCOES, true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(regra);

  // Cores por status (formatação condicional) — dá pra ver o andamento de cada lead de relance
  const regrasAntigas = leads.getConditionalFormatRules().filter(function (r) {
    return r.getRanges().every(function (rg) { return rg.getColumn() !== COL_STATUS; });
  });
  const novasRegras = STATUS_OPCOES.map(function (status) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(STATUS_CORES[status] || "#FFFFFF")
      .setRanges([statusRange])
      .build();
  });
  leads.setConditionalFormatRules(regrasAntigas.concat(novasRegras));

  // Preenche "Novo" em qualquer lead já gravado que ainda não tenha status
  const ultimaLinha = leads.getLastRow();
  if (ultimaLinha >= 2) {
    const dados = leads.getRange(2, 2, ultimaLinha - 1, 1).getValues(); // coluna Nome, só pra saber quais linhas têm lead
    const statusAtual = leads.getRange(2, COL_STATUS, ultimaLinha - 1, 1).getValues();
    const statusNovo = statusAtual.map(function (linha, i) {
      const temNome = dados[i][0] && String(dados[i][0]).trim() !== "";
      const jaTemStatus = linha[0] && String(linha[0]).trim() !== "";
      return [temNome && !jaTemStatus ? "Novo" : linha[0]];
    });
    leads.getRange(2, COL_STATUS, ultimaLinha - 1, 1).setValues(statusNovo);
  }

  leads.autoResizeColumns(1, numCols);
}

/* ------------------------------------------------------------------
   organizarAbaPainel — recria o painel com contagens, ranking de
   regiões, leads por cidade, status do funil e taxa de fechamento.
   ------------------------------------------------------------------ */
function organizarAbaPainel(ss, SEP, LIMITE) {
  let painel = ss.getSheetByName(SHEET_PAINEL);
  if (!painel) {
    painel = ss.insertSheet(SHEET_PAINEL);
  }
  painel.clear();
  painel.clearConditionalFormatRules();

  function tituloSecao(texto, linha) {
    const r = painel.getRange(linha, 1, 1, 2);
    r.merge();
    r.setValue(texto);
    r.setFontWeight("bold");
    r.setFontColor("#FFFFFF");
    r.setBackground("#1E1E1E");
    r.setHorizontalAlignment("left");
  }

  let linha = 1;

  const tituloRange = painel.getRange(linha, 1, 1, 2);
  tituloRange.merge();
  tituloRange.setValue("PAINEL DE LEADS — Crio Harmonização");
  tituloRange.setFontWeight("bold").setFontSize(14).setFontColor("#A8663F");
  linha += 2;

  painel.getRange(linha, 1).setValue("Total de leads").setFontWeight("bold");
  painel.getRange(linha, 2).setFormula('=COUNTA(' + SHEET_LEADS + '!B2:B' + LIMITE + ')');
  linha += 1;

  painel.getRange(linha, 1).setValue("Taxa de fechamento (Fechou / Total)").setFontWeight("bold");
  painel.getRange(linha, 2).setFormula(
    '=IFERROR(COUNTIF(' + SHEET_LEADS + '!K2:K' + LIMITE + SEP + '"Fechou")/COUNTA(' + SHEET_LEADS + '!B2:B' + LIMITE + ')' + SEP + '0)'
  );
  painel.getRange(linha, 2).setNumberFormat("0.0%");
  linha += 2;

  tituloSecao("STATUS DO FUNIL COMERCIAL", linha);
  linha += 1;
  STATUS_OPCOES.forEach(function (status) {
    painel.getRange(linha, 1).setValue(status);
    painel.getRange(linha, 1).setBackground(STATUS_CORES[status] || "#FFFFFF");
    painel.getRange(linha, 2).setFormula(
      '=COUNTIF(' + SHEET_LEADS + '!K2:K' + LIMITE + SEP + '"' + status + '")'
    );
    linha += 1;
  });
  linha += 1;

  tituloSecao("OBJETIVO (Tela 1)", linha);
  linha += 1;
  OBJETIVOS.forEach(function (label) {
    painel.getRange(linha, 1).setValue(label);
    painel.getRange(linha, 2).setFormula(
      '=COUNTIF(' + SHEET_LEADS + '!F2:F' + LIMITE + SEP + '"' + label + '")'
    );
    linha += 1;
  });
  linha += 1;

  tituloSecao("QUANDO PRETENDE INICIAR", linha);
  linha += 1;
  OPCOES_INICIO.forEach(function (label) {
    painel.getRange(linha, 1).setValue(label);
    painel.getRange(linha, 2).setFormula(
      '=COUNTIF(' + SHEET_LEADS + '!I2:I' + LIMITE + SEP + '"' + label + '")'
    );
    linha += 1;
  });
  linha += 1;

  tituloSecao("RANKING DE REGIÕES (Tela 3)", linha);
  linha += 1;
  const inicioRegioes = linha;
  REGIOES.forEach(function (label) {
    painel.getRange(linha, 1).setValue(label);
    painel.getRange(linha, 2).setFormula(
      '=SUMPRODUCT(--ISNUMBER(SEARCH("' + label + '"' + SEP + SHEET_LEADS + '!G2:G' + LIMITE + ')))'
    );
    linha += 1;
  });
  const fimRegioes = linha - 1;
  linha += 1;

  painel.getRange(linha, 1).setValue("Ranking ordenado (maior → menor)").setFontStyle("italic");
  linha += 1;
  painel.getRange(linha, 1).setValue("Região").setFontWeight("bold");
  painel.getRange(linha, 2).setValue("Menções").setFontWeight("bold");
  linha += 1;
  // IMPORTANTE: a fórmula de array (SORT) só pode ser escrita na célula de
  // cima-esquerda de UMA célula só — se for escrita numa faixa de 2 colunas,
  // o Sheets tenta espalhar a mesma fórmula a partir de cada célula da faixa
  // e elas colidem, dando #REF!. Por isso usamos getRange(linha,1) (1 célula).
  painel.getRange(linha, 1).setFormula(
    '=SORT(A' + inicioRegioes + ':B' + fimRegioes + SEP + '2' + SEP + 'FALSE)'
  );
  linha += (REGIOES.length + 2);

  tituloSecao("LEADS POR CIDADE", linha);
  linha += 1;
  painel.getRange(linha, 1).setFormula(
    '=IFERROR(QUERY(' + SHEET_LEADS + '!E2:E' + LIMITE + SEP +
    '"select E, count(E) where E is not null and E <> \'\' group by E order by count(E) desc label E \'Cidade\', count(E) \'Leads\'"' + SEP + '1)' + SEP + '"Sem dados ainda")'
  );
  linha += 12; // espaço reservado pra lista de cidades (ajusta sozinho se crescer)

  tituloSecao("LEADS POR MÊS", linha);
  linha += 1;
  painel.getRange(linha, 1).setFormula(
    '=IFERROR(QUERY(' + SHEET_LEADS + '!A2:A' + LIMITE + SEP +
    '"select year(A), month(A)+1, count(A) where A is not null group by year(A), month(A) order by year(A), month(A) label year(A) \'Ano\', month(A)+1 \'Mês\', count(A) \'Leads\'"' + SEP + '1)' + SEP + '"Sem dados ainda")'
  );

  painel.setColumnWidth(1, 320);
  painel.setColumnWidth(2, 110);
}
