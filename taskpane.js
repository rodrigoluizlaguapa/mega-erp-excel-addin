const GOOGLE_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbx6QhNdtZo9U1p-rsRhXvBPWkv58NHItnajCU3OQKFcRKfiiYxGdjPj5P7dZsa9cww9/exec"; 

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(() => {
      // 1. Reativa o botão do Painel Lateral
      $("#btn-sync").on("click", syncData);
      
      // 2. Associa as funções da Ribbon (sem duplicar o onReady)
      Office.actions.associate("gerarTemplateAgentes", gerarTemplateAgentes);
      Office.actions.associate("processarInclusoesAgentes", processarInclusoesAgentes);
      
      console.log("LG CFO: Motor e Ribbon unificados e prontos.");
    });
  }
});
// --- FUNÇÕES CHAMADAS PELA RIBBON ---

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
    
    await context.sync();
  });
}

async function processarInclusoesAgentes() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values, rowCount");
    await context.sync();

    const dados = usedRange.values;
    if (dados.length <= 1) return; // Só cabeçalho

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
        } else {
          rangeLinha.format.fill.color = "#FFC7CE";
          rangeStatus.values = [[result.response]];
        }
      } catch (err) {
        rangeStatus.values = [["Erro de Conexão"]];
      }
    }
    await context.sync();
  });
}

// --- UTILITÁRIOS ---

function validarCpfCnpj(val) {
  if (!val) return false;
  // Validação básica de tamanho (pode ser expandida para algoritmo de dígitos)
  return (val.length === 11 || val.length === 14);
}

async function syncData() {
  // Função para o painel lateral carregar o Razão Contábil
  const modulo = "list"; // Valor padrão
  const res = await fetch(`${GOOGLE_WEBAPP_URL}?action=${modulo}`);
  const responseData = await res.json();
  // ... lógica de inserção na planilha conforme necessário
}

// Adicione no final do taskpane.js
window.gerarTemplateAgentes = gerarTemplateAgentes;
window.processarInclusoesAgentes = processarInclusoesAgentes;
