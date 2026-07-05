/* ==================================================================
   FORMULÁRIO CRIO HARMONIZAÇÃO — Amanda Sobottka
   Google Apps Script — grava leads na planilha "Leads" e mantém um
   painel automático na aba "Painel".

   COMO USAR (resumo — o passo a passo completo está no guia de instalação):
   1. Crie uma planilha nova no Google Sheets.
   2. Menu Extensões > Apps Script.
   3. Apague o código de exemplo e cole TODO este arquivo no lugar.
   4. Salve (ícone de disquete ou Ctrl+S).
   5. No menu de funções (barra de cima, ao lado do ícone de play),
      selecione "criarPlanilha" e clique em Executar. Autorize as
      permissões quando pedido. Isso cria as abas "Leads" e "Painel"
      já com cabeçalhos e fórmulas prontas.
   6. Implantar > Nova implantação > tipo "Aplicativo da Web".
      - Executar como: Eu (seu e-mail)
      - Quem tem acesso: Qualquer pessoa
   7. Copie a URL gerada e cole na variável APPS_SCRIPT_URL no topo
      do arquivo HTML do formulário.

   OBS: se você já tinha uma versão anterior deste script rodando,
   rode "criarPlanilha" de novo depois de colar esta versão — ela
   atualiza o cabeçalho e as fórmulas do Painel para os novos campos
   (Sobrenome, Telefone, Quando pretende iniciar).
   ================================================================== */

const SHEET_LEADS = "Leads";
const SHEET_PAINEL = "Painel";

// Ordem das colunas na aba Leads
const CABECALHO_LEADS = [
  "Data/Hora", "Nome", "Sobrenome", "Telefone", "Cidade",
  "Objetivo", "Região(ões)", "Rotina", "Quando pretende iniciar", "Consentimento"
];

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

/* Permite testar a URL do Web App abrindo ela direto no navegador */
function doGet(e) {
  return ContentService.createTextOutput("Formulário Crio Harmonização — endpoint ativo.");
}

/* ------------------------------------------------------------------
   criarPlanilha — RODE ESTA FUNÇÃO UMA VEZ, MANUALMENTE, antes de
   publicar o Web App (ou de novo, se atualizar o código). Cria/atualiza
   as abas "Leads" e "Painel" já formatadas.
   ------------------------------------------------------------------ */
function criarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Aba Leads
  let leads = ss.getSheetByName(SHEET_LEADS);
  if (!leads) {
    leads = ss.insertSheet(SHEET_LEADS);
  }
  leads.getRange(1, 1, 1, CABECALHO_LEADS.length).setValues([CABECALHO_LEADS]);
  leads.setFrozenRows(1);
  leads.getRange(1, 1, 1, CABECALHO_LEADS.length).setFontWeight("bold");

  // Aba Painel
  let painel = ss.getSheetByName(SHEET_PAINEL);
  if (!painel) {
    painel = ss.insertSheet(SHEET_PAINEL);
  }
  painel.clear();

  const LIMITE = 5000; // linhas cobertas pelas fórmulas (aumente se precisar)

  // Colunas na aba Leads: A=Data/Hora B=Nome C=Sobrenome D=Telefone E=Cidade
  //                       F=Objetivo G=Região(ões) H=Rotina I=Quando pretende iniciar J=Consentimento

  let linha = 1;

  painel.getRange(linha, 1).setValue("PAINEL DE LEADS — Crio Harmonização").setFontWeight("bold").setFontSize(14);
  linha += 2;

  painel.getRange(linha, 1).setValue("Total de leads");
  painel.getRange(linha, 2).setFormula('=COUNTA(' + SHEET_LEADS + '!B2:B' + LIMITE + ')');
  linha += 2;

  painel.getRange(linha, 1).setValue("OBJETIVO (Tela 1)").setFontWeight("bold");
  linha += 1;
  OBJETIVOS.forEach(function (label) {
    painel.getRange(linha, 1).setValue(label);
    painel.getRange(linha, 2).setFormula(
      '=COUNTIF(' + SHEET_LEADS + '!F2:F' + LIMITE + ',"' + label + '")'
    );
    linha += 1;
  });
  linha += 1;

  painel.getRange(linha, 1).setValue("QUANDO PRETENDE INICIAR").setFontWeight("bold");
  linha += 1;
  OPCOES_INICIO.forEach(function (label) {
    painel.getRange(linha, 1).setValue(label);
    painel.getRange(linha, 2).setFormula(
      '=COUNTIF(' + SHEET_LEADS + '!I2:I' + LIMITE + ',"' + label + '")'
    );
    linha += 1;
  });
  linha += 1;

  painel.getRange(linha, 1).setValue("RANKING DE REGIÕES (Tela 3)").setFontWeight("bold");
  linha += 1;
  const inicioRegioes = linha;
  REGIOES.forEach(function (label) {
    painel.getRange(linha, 1).setValue(label);
    painel.getRange(linha, 2).setFormula(
      '=SUMPRODUCT(--ISNUMBER(SEARCH("' + label + '",' + SHEET_LEADS + '!G2:G' + LIMITE + ')))'
    );
    linha += 1;
  });
  const fimRegioes = linha - 1;
  linha += 1;

  painel.getRange(linha, 1).setValue("Ranking ordenado (maior → menor)").setFontStyle("italic");
  linha += 1;
  painel.getRange(linha, 1, 1, 2).setFormula(
    '={"Região","Menções"; SORT(A' + inicioRegioes + ':B' + fimRegioes + ',2,FALSE)}'
  );
  linha += (REGIOES.length + 2);

  painel.getRange(linha, 1).setValue("LEADS POR CIDADE").setFontWeight("bold");
  linha += 1;
  painel.getRange(linha, 1).setFormula(
    '=IFERROR(QUERY(' + SHEET_LEADS + '!E2:E' + LIMITE +
    ',"select E, count(E) where E is not null and E <> \'\' group by E order by count(E) desc label E \'Cidade\', count(E) \'Leads\'",1),"Sem dados ainda")'
  );

  painel.autoResizeColumns(1, 2);

  try {
    SpreadsheetApp.getUi().alert(
      "Pronto! As abas \"Leads\" e \"Painel\" foram criadas/atualizadas."
    );
  } catch (e) {
    Logger.log("Concluído. (Não foi possível mostrar o alerta de UI, mas as abas foram criadas normalmente.)");
  }
}
