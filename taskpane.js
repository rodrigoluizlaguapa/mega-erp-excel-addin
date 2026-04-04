const GOOGLE_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbx6QhNdtZo9U1p-rsRhXvBPWkv58NHItnajCU3OQKFcRKfiiYxGdjPj5P7dZsa9cww9/exec";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(() => {
      // Conecta os cliques aos botões da interface
      $("#btn-sync").on("click", syncData);
      $("#btn-template").on("click", gerarTemplateAgentes);
      $("#btn-enviar").on("click", processarInclusoesAgentes);
    });
  }
});

// --- FUNÇÕES DA ABA RELATÓRIOS ---
// --- FUNÇÕES DA ABA RELATÓRIOS ---
async function syncData() {
  try {
    $("#status").text("Consultando Mega ERP... Aguarde.");
    
    const modulo = $("#modulo").val();
    const dataInicio = $("#data-inicio").val();
    const dataFim = $("#data-fim").val();

    // Faz a chamada ao Google Apps Script passando as datas do painel
    const res = await fetch(`${GOOGLE_WEBAPP_URL}?action=${modulo}&inicio=${dataInicio}&fim=${dataFim}`);
    const result = await res.json();

    // Se o servidor retornar vazio ou erro
    if (!result || result.length === 0) {
      $("#status").text("Nenhum dado encontrado para este período.");
      return;
    }

    // Escreve os dados no Excel
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getUsedRange().clear(); // Limpa a planilha atual

      // Se result.data for a matriz (depende de como o seu Google Script devolve)
      // Assumindo que o Apps Script devolve um array de objetos (JSON normal)
      let dadosParaExcel = [];
      let headers = [];

      if (Array.isArray(result) && result.length > 0) {
        headers = Object.keys(result[0]); // Pega o nome das colunas
        dadosParaExcel.push(headers); // Linha 1 = Cabeçalho

        // Preenche as linhas seguintes
        result.forEach(item => {
          let linha = [];
          headers.forEach(h => linha.push(item[h]));
          dadosParaExcel.push(linha);
        });
      } else if (result.data && Array.isArray(result.data)) {
        // Caso o seu script devolva dentro de um objeto { success: true, data: [...] }
        headers = Object.keys(result.data[0]);
        dadosParaExcel.push(headers);
        result.data.forEach(item => {
          let linha = [];
          headers.forEach(h => linha.push(item[h]));
          dadosParaExcel.push(linha);
        });
      }

      if (dadosParaExcel.length > 0) {
        const range = sheet.getRangeByIndexes(0, 0, dadosParaExcel.length, dadosParaExcel[0].length);
        range.values = dadosParaExcel;
        
        // Formata o cabeçalho de azul
        const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
        headerRange.format.fill.color = "#0078d4";
        headerRange.format.font.color = "white";
        headerRange.format.font.bold = true;
        
        range.format.autofitColumns(); // Ajusta a largura das colunas
      }

      await context.sync();
    });

    $("#status").text("Razão Contábil gerado com sucesso!");

  } catch (error) {
    console.error(error);
    $("#status").text("Erro ao baixar os dados. Verifique a conexão.");
  }
}

// --- FUNÇÕES DA ABA AGENTES ---
async function gerarTemplateAgentes() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getUsedRange().clear();

    const headers = [["NOME", "APELIDO", "TIPO (F/J)", "CPF_CNPJ", "LOGRADOURO", "IBGE_MUNICIPIO", "STATUS/ERRO"]];
    const range = sheet.getRange("A1:G1");
    range.values = headers;
    range.format.fill.color = "#0078d4";
    range.format.font.color = "white";
    range.format.font.bold = true;
    range.format.autofitColumns();
    
    $("#status").text("Planilha gerada! Preencha os dados e clique em Enviar.");
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
    if (dados.length <= 1) {
      $("#status").text("Nenhum dado encontrado para enviar.");
      return; 
    }

    let enviados = 0;
    let erros = 0;

    for (let i = 1; i < dados.length; i++) {
      const nome = dados[i][0];
      const tipo = dados[i][2]?.toUpperCase();
      const cpfCnpj = dados[i][3]?.toString().replace(/\D/g, '');
      const ibge = dados[i][5]?.toString();

      let erroLocal = "";
      if (!nome) erroLocal = "Nome obrigatório";
      else if (!tipo || (tipo !== 'F' && tipo !== 'J')) erroLocal = "Tipo deve ser F ou J";
      else if (!validarCpfCnpj(cpfCnpj)) erroLocal = "CPF/CNPJ Inválido";
      else if (!ibge || ibge.length !== 7) erroLocal = "IBGE deve ter 7 dígitos";

      const rangeLinha = sheet.getRange(`A${i + 1}:F${i + 1}`);
      const rangeStatus = sheet.getRange(`G${i + 1}`);

      if (erroLocal) {
        rangeLinha.format.fill.color = "#FFC7CE";
        rangeStatus.values = [[erroLocal]];
        erros++;
        continue;
      }

      try {
        const payload = {
          action: "cadastrarAgente",
          data: {
            AgenteNome: nome,
            AgenteApelido: dados[i][1] || nome,
            AgenteTipo: tipo,
            CpfCnpj: cpfCnpj,
            EnderecoLogradouro: dados[i][4] || "",
            MunicipioCodigoIBGE: ibge
          }
        };

        const response = await fetch(GOOGLE_WEBAPP_URL, {
          method: "POST",
          body: JSON.stringify(payload)
        });
        const result = await response.json();

        if (result.success) {
          rangeLinha.format.fill.color = "#C6EFCE";
          rangeStatus.values = [["Sucesso"]];
          enviados++;
        } else {
          rangeLinha.format.fill.color = "#FFC7CE";
          rangeStatus.values = [[result.response]];
          erros++;
        }
      } catch (err) {
        rangeStatus.values = [["Erro de Conexão"]];
        erros++;
      }
    }
    await context.sync();
    $("#status").text(`Operação concluída: ${enviados} Cadastros | ${erros} Erros.`);
  });
}

function validarCpfCnpj(val) {
  if (!val) return false;
  return (val.length === 11 || val.length === 14);
}
