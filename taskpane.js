const API_URL = "http://localhost:3000";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(() => {
      $("#btn-sync").on("click", syncData);
      $("#btn-template").on("click", gerarTemplateAgentes);
      $("#btn-enviar").on("click", processarInclusoesAgentes);
      console.log("Suplemento carregado e pronto.");
    });
  }
});

async function syncData() {
  try {
    $("#status").text("Verificando parâmetros...");

    // Captura valores e limpa espaços
    const modulo = ($("#modulo").val() || "Relatorio").trim();
    const dataInicio = $("#data-inicio").val();
    const dataFim = $("#data-fim").val();

    // Validação de segurança para evitar URL malformada (evita o 404)
    if (!dataInicio || !dataFim) {
      $("#status").text("Erro: Preencha as datas de início e fim.");
      return;
    }

    const urlFinal = `${API_URL}/api/lancamentos?inicio=${dataInicio}&fim=${dataFim}`;
    console.log("Chamando API:", urlFinal);
    $("#status").text("Consultando API Mega...");

    const res = await fetch(urlFinal);
    
    if (!res.ok) {
      const errorDetail = await res.text();
      throw new Error(`Erro ${res.status}: ${res.statusText}`);
    }

    const responseData = await res.json();
    const lista = responseData.data || [];

    if (lista.length === 0) {
      $("#status").text("Nenhum dado encontrado para este período.");
      return;
    }

    await Excel.run(async (context) => {
      const sheetName = "Relatorio_" + modulo.replace(/\s+/g, '_').toUpperCase();
      let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();

      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add(sheetName);
      } else {
        sheet.getRange().clear();
      }

      sheet.activate();

      // Transforma objetos JSON em matriz para o Excel
      const headers = Object.keys(lista[0]);
      const rows = lista.map(item => headers.map(h => item[h] ?? ""));
      const finalData = [headers, ...rows];

      const range = sheet.getRangeByIndexes(0, 0, finalData.length, headers.length);
      range.values = finalData;

      // Estilização básica
      const headerRange = range.getRow(0);
      headerRange.format.fill.color = "#0078d4";
      headerRange.format.font.color = "white";
      headerRange.format.font.bold = true;
      range.format.autofitColumns();

      await context.sync();
      $("#status").text(`Sucesso! ${lista.length} linhas baixadas.`);
    });

  } catch (error) {
    console.error("Erro detalhado:", error);
    $("#status").text("Falha: " + error.message);
  }
}

async function gerarTemplateAgentes() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItemOrNullObject("Carga_Agentes");
      await context.sync();

      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add("Carga_Agentes");
      }
      sheet.activate();
      sheet.getRange().clear();

      const headers = [["NOME", "APELIDO", "TIPO (F/J)", "CPF_CNPJ", "LOGRADOURO", "IBGE_MUNICIPIO", "STATUS_RETORNO"]];
      sheet.getRange("A1:G1").values = headers;
      sheet.getRange("A1:G1").format.font.bold = true;
      
      await context.sync();
      $("#status").text("Template de agentes criado.");
    });
  } catch (error) {
    $("#status").text("Erro ao criar template.");
  }
}

async function processarInclusoesAgentes() {
    try {
        await Excel.run(async (context) => {
            $("#status").text("Lendo dados da planilha...");
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();

            const [headers, ...rows] = range.values;
            const payload = rows.map(row => {
                let obj = {};
                headers.forEach((h, i) => obj[h] = row[i]);
                return obj;
            }).filter(item => item.NOME); // Ignora linhas sem nome

            $("#status").text("Enviando para o servidor...");
            const response = await fetch(`${API_URL}/api/agentes`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const resData = await response.json();
            if (resData.success) {
                $("#status").text("✓ Enviado com sucesso!");
            } else {
                $("#status").text("Erro no servidor: " + resData.error);
            }
        });
    } catch (error) {
        $("#status").text("Erro no processamento.");
    }
}
