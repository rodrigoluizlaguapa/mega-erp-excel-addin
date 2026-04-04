const GOOGLE_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbwcNC5muN0LUsOeF2Ff1XXDFR22ZDnt10e5d7PazoElXvgs5q6rKuZSODop3IrHrALK/exec";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(() => {
      $("#btn-sync").on("click", syncData);
      $("#btn-template").on("click", gerarTemplateAgentes);
      $("#btn-enviar").on("click", processarInclusoesAgentes);
    });
  }
});

// --- FUNÇÕES DA ABA RELATÓRIOS (BAIXAR DADOS) ---
async function syncData() {
  try {
    $("#status").text("Consultando Mega ERP... Aguarde.");
    
    const modulo = $("#modulo").val();
    const dataInicio = $("#data-inicio").val();
    const dataFim = $("#data-fim").val();

    // 1. Busca os dados no Google
    const res = await fetch(`${GOOGLE_WEBAPP_URL}?action=${modulo}&inicio=${dataInicio}&fim=${dataFim}`);
    const rawText = await res.text();
    
    let result;
    try {
        result = JSON.parse(rawText);
    } catch(e) {
        $("#status").text("Erro: O servidor não retornou um formato válido.");
        return;
    }

    // 2. Normaliza os dados (não importa se vem direto num array ou dentro de "result.data")
    let arrayDeDados = [];
    if (Array.isArray(result)) {
        arrayDeDados = result;
    } else if (result && result.data && Array.isArray(result.data)) {
        arrayDeDados = result.data;
    } else if (result && typeof result === 'object') {
        arrayDeDados = [result]; 
    }

    // Se estiver vazio de fato
    if (arrayDeDados.length === 0) {
      $("#status").text("Consulta concluída, mas não há dados no período.");
      return;
    }

    // 3. Escreve os dados no Excel (em uma NOVA ABA)
    await Excel.run(async (context) => {
      let sheetName = "Relatorio_" + modulo.toUpperCase();
      let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();
      
      // Cria a aba se não existir, limpa se existir
      if (sheet.isNullObject) {
          sheet = context.workbook.worksheets.add(sheetName);
      } else {
          sheet.getUsedRange().clear();
      }
      
      // PULA PARA A ABA NOVA
      sheet.activate();

      let headers = Object.keys(arrayDeDados[0]);
      let dadosParaExcel = [headers];

      arrayDeDados.forEach(item => {
        let linha = [];
        headers.forEach(h => {
            let valor = item[h];
            if(valor === null || valor === undefined) valor = "";
            linha.push(valor);
        });
        dadosParaExcel.push(linha);
      });

      const range = sheet.getRangeByIndexes(0, 0, dadosParaExcel.length, dadosParaExcel[0].length);
      range.values = dadosParaExcel;
      
      const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
      headerRange.format.fill.color = "#0078d4";
      headerRange.format.font.color = "white";
      headerRange.format.font.bold = true;
      
      range.format.autofitColumns();
      await context.sync();
    });

    $("#status").text(`Sucesso! ${arrayDeDados.length} linhas baixadas.`);

  } catch (error) {
    console.error(error);
    $("#status").text("Erro de conexão ao baixar os dados.");
  }
}

// --- FUNÇÕES DA ABA AGENTES (ENVIAR PARA MEGA) ---
async function gerarTemplateAgentes() {
  await Excel.run(async (context) => {
    let sheetName = "Carga_Agentes";
    let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    await context.sync();
    
    // Cria a aba se não existir, limpa se já existir
    if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add(sheetName);
    } else {
        sheet.getUsedRange().clear();
    }
    
    // PULA PARA A ABA NOVA
    sheet.activate();

    const headers = [["NOME", "APELIDO", "TIPO (F/J)", "CPF_CNPJ", "LOGRADOURO", "IBGE_MUNICIPIO", "STATUS/ERRO"]];
    const range = sheet.getRange("A1:G1");
    range.values = headers;
    range.format.fill.color = "#217346";
    range.format.font.color = "white";
    range.format.font.bold = true;
    range.format.autofitColumns();
    
    $("#status").text("Aba de Carga criada! Preencha e clique em Enviar.");
    await context.sync();
  });
}

async function processarInclusoesAgentes() {
  await Excel.run(async (context) => {
    $("#status").text("Analisando e enviando dados...");
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values, rowCount");
    await context.sync();

    const dados = usedRange.values;
    if (dados.
