/**
 * Cria um Google Form a partir de dados na Planilha Google ativa.
 * Linha 1: Títulos das perguntas.
 * Linha 2: Tipos das perguntas (Múltipla Escolha, Texto Curto, Data, etc.).
 * Linha 3+: Opções (para tipos que as usam).
 */
function createFormFromSheet() {
  // 1. Get the active Spreadsheet and the active Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  // 2. Get all data from the sheet
  const data = sheet.getDataRange().getValues();

  // 3. Basic validation: Check if there are at least 2 rows (Title, Type)
  if (!data || data.length < 2) {
    Logger.log("Erro: A planilha deve conter pelo menos 2 linhas: uma para títulos e uma para tipos de pergunta.");
    SpreadsheetApp.getUi().alert("Erro: A planilha deve conter pelo menos 2 linhas: uma para títulos e uma para tipos de pergunta.");
    return; // Stop the script
  }

  // 4. Create a new Google Form
  const formTitle = `Formulário criado de: ${sheetName} (${new Date().toLocaleString()})`; // Add timestamp to avoid name conflicts
  const form = FormApp.create(formTitle);
  form.setDescription(`Este formulário foi gerado automaticamente da Planilha Google '${sheetName}'.`);
  Logger.log(`Criado novo formulário: "${formTitle}"`);
  Logger.log(`URL de Edição: ${form.getEditUrl()}`);

  // 5. Extract questions (Row 1) and types (Row 2)
  const questions = data[0]; // [Question1, Question2, ...]
  const types = data[1];     // [Type1, Type2, ...]
  const numColumns = questions.length;
  const numRows = data.length;

  // 6. Process each column to create a form item based on type
  for (let j = 0; j < numColumns; j++) {
    const questionTitle = questions[j] ? questions[j].toString().trim() : "";
    const questionTypeRaw = types[j] ? types[j].toString().trim() : "";
    const questionType = questionTypeRaw.toLowerCase(); // Normalize type for comparison

    // Skip if the question title in the header row is empty
    if (!questionTitle) {
      Logger.log(`Pulando coluna ${j + 1} porque o título está vazio.`);
      continue;
    }
    // Skip if the question type is empty
    if (!questionType) {
      Logger.log(`Pulando coluna ${j + 1} ("${questionTitle}") porque o tipo está vazio na linha 2.`);
      continue;
    }

    Logger.log(`Processando coluna ${j+1}: Título="${questionTitle}", Tipo="${questionTypeRaw}"`);

    // 7. Extract choices (from Row 3 onwards) - only relevant for specific types
    const choices = [];
    if (numRows > 2) { // Only look for choices if there are rows beyond row 2
      for (let i = 2; i < numRows; i++) { // Loop through ROWS starting from the third row (index 2)
        const choiceValue = data[i][j];
        // Add the choice only if it's not empty
        if (choiceValue !== "" && choiceValue != null) {
          choices.push(choiceValue.toString()); // Ensure choices are strings
        }
      }
    }

    // 8. Add the item to the form based on its type
    try {
      switch (questionType) {
        case 'múltipla escolha':
          if (choices.length > 0) {
            form.addMultipleChoiceItem()
              .setTitle(questionTitle)
              .setChoiceValues(choices);
            Logger.log(`   Adicionado: Múltipla Escolha com ${choices.length} opções.`);
          } else {
            Logger.log(`   Aviso: Múltipla Escolha "${questionTitle}" não tem opções na planilha (Linha 3+). Item não adicionado.`);
          }
          break;

        case 'caixa de seleção':
           if (choices.length > 0) {
             form.addCheckboxItem()
              .setTitle(questionTitle)
              .setChoiceValues(choices);
             Logger.log(`   Adicionado: Caixa de Seleção com ${choices.length} opções.`);
           } else {
             Logger.log(`   Aviso: Caixa de Seleção "${questionTitle}" não tem opções na planilha (Linha 3+). Item não adicionado.`);
           }
           break;

        case 'lista suspensa':
           if (choices.length > 0) {
             form.addListItem()
               .setTitle(questionTitle)
               .setChoiceValues(choices);
             Logger.log(`   Adicionado: Lista Suspensa com ${choices.length} opções.`);
           } else {
              Logger.log(`   Aviso: Lista Suspensa "${questionTitle}" não tem opções na planilha (Linha 3+). Item não adicionado.`);
           }
           break;

        case 'texto curto':
          form.addTextItem().setTitle(questionTitle);
          Logger.log(`   Adicionado: Texto Curto.`);
          break;

        case 'parágrafo':
          form.addParagraphTextItem().setTitle(questionTitle);
          Logger.log(`   Adicionado: Parágrafo.`);
          break;

        case 'escala linear':
          // Default scale 1 to 5. Could be customized further if needed.
          form.addScaleItem()
            .setTitle(questionTitle)
            .setBounds(1, 5);
          Logger.log(`   Adicionado: Escala Linear (1-5).`);
          break;

        case 'data':
          form.addDateItem().setTitle(questionTitle);
          Logger.log(`   Adicionado: Data.`);
          break;

        case 'hora':
          form.addTimeItem().setTitle(questionTitle);
           Logger.log(`   Adicionado: Hora.`);
          break;

        case 'data e hora':
          form.addDateTimeItem().setTitle(questionTitle);
           Logger.log(`   Adicionado: Data e Hora.`);
          break;

        case 'duração':
           form.addDurationItem().setTitle(questionTitle);
           Logger.log(`   Adicionado: Duração.`);
           break;

        // Add cases for GridItem or CheckboxGridItem here if needed (more complex)

        default:
          Logger.log(`   Aviso: Tipo de pergunta não reconhecido ou inválido na coluna ${j + 1}: "${questionTypeRaw}". Item não adicionado.`);
          // Optionally, add a default text item as a fallback:
          // form.addTextItem().setTitle(questionTitle + ' (Tipo Inválido)');
      }
    } catch (error) {
        Logger.log(`   ERRO ao adicionar item para "${questionTitle}": ${error}`);
    }
  }

  // 9. Final message
  Logger.log("Processo de criação do formulário concluído.");
  SpreadsheetApp.getUi().alert(`Formulário criado! Você pode editá-lo aqui: ${form.getEditUrl()}`);
}

// A função onOpen continua a mesma - não precisa ser alterada
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('✨ Ferramentas de Formulário')
      .addItem('Criar Formulário desta Planilha', 'createFormFromSheet')
      .addToUi();
}