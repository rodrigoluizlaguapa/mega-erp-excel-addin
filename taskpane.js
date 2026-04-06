const API_URL = "http://localhost:3000";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(() => {
      $("#btn-sync").on("click", syncData);
      $("#btn-template").on("click", gerarTemplateAgentes);
      $("#btn-enviar").on("click", processarInclusoesAgentes);
    });
  }
});

/**
 * --- FUNÇÕES DA ABA RELATÓRIOS (BAIXAR DADOS) ---
 * Busca dados do servidor e cria uma nova aba formatada.
 */
async function syncData() {
  try {
    $("#status").text("Consultando Mega ERP... Aguarde.");
    
    const modulo = $("#modulo").val() || "Dados";
    const dataInicio = $("#data-inicio").val();
    const dataFim = $("#data-fim").val();

    // 1. Busca os dados no servidor
    const res = await fetch(`${API_URL}/api/lancamentos?inicio=${dataInicio}&fim=${dataFim}`);
    
    if (!res.ok) {
      $("#status").text(`Erro: ${res.status} - ${res.statusText}`);
      return;
    }

    const result = await res.json();
    
    // Normalização robusta do array de dados
    let arrayDeDados = Array.isArray(result) ? result : (result.data || []);
    if (!Array.isArray(arrayDeDados) && typeof result === 'object') {
        arrayDeDados = [result];
    }

    if (arrayDeDados.length === 0) {
      $("#status").text("Consulta concluída, mas não há dados no período.");
      return;
    }

    // 2. Escrita no Excel
    await Excel.run(async (context) => {
      const sheetName = "Relatorio_" + modulo.toUpperCase();
      let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();
      
      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add(sheetName);
      } else {
        // .clear() direto na sheet evita erro 404 se a aba estiver vazia
        sheet.getRange().clear(); 
      }
      
      sheet.activate();
      
      const headers = Object.keys(arrayDeDados[0]);
      const dadosParaExcel = [headers];
      
      arrayDeDados.forEach(item => {
        const linha = headers.map(h => (item[h] ?? ""));
        dadosParaExcel.push(linha);
      });

      const range = sheet.getRangeByIndexes(0, 0, dadosParaExcel.length, dadosParaExcel[0].length);
      range.values = dadosParaExcel;
      
      // Formatação
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
    $("#status").text("Erro de conexão ou processamento de dados.");
  }
}

/**
 * --- FUNÇÕES DA ABA AGENTES (ENVIAR PARA MEGA) ---
 * Gera o cabeçalho padrão para preenchimento.
 */
async function gerarTemplateAgentes() {
  try {
    await Excel.run(async (context) => {
      const sheetName = "Carga_Agentes";
      let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();
      
      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add(sheetName);
      } else {
        sheet.getRange().clear();
      }
      
      sheet.activate();
      
      const headers = [["NOME", "APELIDO", "TIPO (F/J)", "CPF_CNPJ", "LOGRADOURO", "IBGE_MUNICIPIO", "STATUS/ERRO"]];
      const range = sheet.getRange("A1:G1");
      range.values = headers;
      range.format.fill.color = "#217346";
      range.format.font.color = "white";
      range.format.font.bold = true;
      range.format.autofitColumns();
      
      await context.sync();
      $("#status").text("Aba de Carga criada! Preencha e clique em Enviar.");
    });
  } catch (error) {
    console.error(error);
    $("#status").text("Erro ao criar aba de template.");
  }
}

/**
 * Lê os dados da aba ativa e envia via POST para a API.
 */
async function processarInclusoesAgentes() {
  try {
    await Excel.run(async (context) => {
      $("#status").text("Analisando e enviando dados...");
      
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("values");
      await context.sync();

      if (usedRange.isNullObject || usedRange.values.length < 2) {
        $("#status").text("Nenhum dado encontrado na planilha!");
        return;
      }

      const dados = usedRange.values;
      const headers = dados[0];
      
      // Filtra linhas que podem estar vazias no final da planilha
      const payload = dados.slice(1)
        .filter(linha => linha.some(celula => celula !== "" && celula !== null))
        .map(linha => {
          let obj = {};
          headers.forEach((header, idx) => {
            obj[header] = linha[idx] ?? null;
          });
          return obj;
        });

      if (payload.length === 0) {
        $("#status").text("Não há dados válidos para enviar.");
        return;
      }

      // Envio para o servidor
      const response = await fetch(`${API_URL}/api/agentes`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        const erroTxt = await response.text();
        $("#status").text(`Erro ${response.status}: ${erroTxt || response.statusText}`);
        return;
      }

      const result = await response.json();
      
      if (result.success) {
        $("#status").text(`✓ ${payload.length} agentes enviados com sucesso!`);
      } else {
        $("#status").text(`Erro no processamento: ${result.error || 'Verifique o servidor.'}`);
      }
    });
  } catch (error) {
    console.error(error);
    $("#status").text("Erro ao ler Excel ou falha de rede.");
  }
}
