/**
 * @OnlyCurrentDoc
 *
 * PROFESSOR ONLINE AUTOMÁTICO (POA)
 * v. 1.2
 * 
 * Automatiza o preenchimento de notas no portal Professor Online (SEDUC-CE)
 * a partir de uma lista copiada do Google Planilhas.
 * 
 * Esse script, em específico, seleciona os números de matrícula e notas epecíficas
 * para serem copiadas para a área de transferência.
 * 
 * Autor: Anderson Freitas (anderson.silva7@prof.ce.gov.br)
 */

// =================================================================================
// 1. CRIAÇÃO DO MENU
// =================================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 POA')
    .addItem('QUALITATIVA', 'copiarQualitativa')
    .addSeparator()
    .addItem('L. PORTUGUESA', 'copiarPortuguesa')
    .addItem('ED. FISICA', 'copiarEdFisica')
    .addItem('L. ESPANHOLA', 'copiarEspanhola')
    .addItem('L. INGLESA', 'copiarInglesa')
    .addItem('ARTES', 'copiarArtes')
    .addSeparator()
    .addItem('MATEMATICA', 'copiarMatematica')
    .addItem('BIOLOGIA', 'copiarBiologia')
    .addItem('FISICA', 'copiarFisica')
    .addItem('QUIMICA', 'copiarQuimica')
    .addSeparator()
    .addItem('FILOSOFIA', 'copiarFilosofia')
    .addItem('GEOGRAFIA', 'copiarGeografia')
    .addItem('HISTORIA', 'copiarHistoria')
    .addItem('SOCIOLOGIA', 'copiarSociologia')
    .addSeparator()
    .addItem('INF. BASICA', 'copiarInfBasica')
    .addToUi();
}


// =================================================================================
// 2. FUNÇÕES DE ATALHO (Sem alterações)
// =================================================================================

function copiarQualitativa() { copiarIntervalosDefinidos('QUALITATIVA', 'B4:B52', 'D4:D52'); }
function copiarPortuguesa()  { copiarIntervalosDefinidos('L. PORTUGUESA', 'B4:B52', 'E4:E52'); }
function copiarEdFisica()    { copiarIntervalosDefinidos('ED. FISICA', 'B4:B52', 'F4:F52'); }
function copiarEspanhola()   { copiarIntervalosDefinidos('L. ESPANHOLA', 'B4:B52', 'G4:G52'); }
function copiarInglesa()     { copiarIntervalosDefinidos('L. INGLESA', 'B4:B52', 'H4:H52'); }
function copiarArtes()       { copiarIntervalosDefinidos('ARTES', 'B4:B52', 'I4:I52'); }
function copiarMatematica()  { copiarIntervalosDefinidos('MATEMATICA', 'B4:B52', 'J4:J52'); }
function copiarBiologia()    { copiarIntervalosDefinidos('BIOLOGIA', 'B4:B52', 'K4:K52'); }
function copiarFisica()      { copiarIntervalosDefinidos('FISICA', 'B4:B52', 'L4:L52'); }
function copiarQuimica()     { copiarIntervalosDefinidos('QUIMICA', 'B4:B52', 'M4:M52'); }
function copiarFilosofia()   { copiarIntervalosDefinidos('FILOSOFIA', 'B4:B52', 'N4:N52'); }
function copiarGeografia()   { copiarIntervalosDefinidos('GEOGRAFIA', 'B4:B52', 'O4:O52'); }
function copiarHistoria()    { copiarIntervalosDefinidos('HISTORIA', 'B4:B52', 'P4:P52'); }
function copiarSociologia()  { copiarIntervalosDefinidos('SOCIOLOGIA', 'B4:B52', 'Q4:Q52'); }
function copiarInfBasica()   { copiarIntervalosDefinidos('INF. BASICA', 'B4:B52', 'R4:R52'); }


// =================================================================================
// 3. FUNÇÃO PRINCIPAL (Lógica de formatação ajustada)
// =================================================================================

function copiarIntervalosDefinidos(nomeDisciplina, intervaloMatriculas, intervaloNotas) {
  const ui = SpreadsheetApp.getUi();
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const valoresMatriculas = planilha.getRange(intervaloMatriculas).getValues();
  const valoresNotas = planilha.getRange(intervaloNotas).getValues();

  let dadosFormatados = "";
  let contadorLinhas = 0;

  for (let i = 0; i < valoresMatriculas.length; i++) {
    const matricula = valoresMatriculas[i][0];
    const notaCrua = valoresNotas[i][0];

    if (matricula) {
      // --- MUDANÇA PRINCIPAL AQUI ---
      // 1. Garante que a nota seja tratada como número, mesmo que a célula esteja vazia (será 0).
      // 2. Converte vírgulas para pontos, caso haja.
      // 3. Formata para ter exatamente uma casa decimal usando toFixed(1).
      let notaNumerica = 0;
      if (notaCrua || notaCrua === 0) {
        notaNumerica = parseFloat(String(notaCrua).replace(',', '.'));
      }
      
      // Se a conversão falhar (ex: texto na célula), a nota será 0.
      if (isNaN(notaNumerica)) {
        notaNumerica = 0;
      }

      const notaFormatada = notaNumerica.toFixed(1);
      // --- FIM DA MUDANÇA ---

      dadosFormatados += matricula + '\t' + notaFormatada + '\n';
      contadorLinhas++;
    }
  }

  dadosFormatados = dadosFormatados.trim();

  const htmlOutput = HtmlService.createHtmlOutput(
    `
    <style>
      body { font-family: Arial, sans-serif; }
      textarea { width: 98%; height: 250px; margin-top: 10px; font-family: monospace; }
      .instructions { font-size: 14px; }
      .info { font-weight: bold; }
    </style>
    <body>
      <p class="instructions">Dados da disciplina "<span class="info">${nomeDisciplina}</span>" prontos para copiar.</p>
      <p class="instructions"><b>Passo 1:</b> Clique na caixa abaixo e pressione <b>Ctrl+A</b> para selecionar tudo.</p>
      <p class="instructions"><b>Passo 2:</b> Pressione <b>Ctrl+C</b> para copiar.</p>
      <textarea readonly>${dadosFormatados}</textarea>
    </body>
    `
  )
  .setWidth(450)
  .setHeight(400);

  ui.showModalDialog(htmlOutput, `Copiar Notas de ${nomeDisciplina}`);
}
