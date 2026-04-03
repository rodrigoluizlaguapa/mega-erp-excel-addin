const GOOGLE_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbwkTGfXu8KJ_HVfFDcrswTshMST20E8BIfmOBBsgeVt_rhzUZ3HIQhWpl8EYAJHsLaU/exec"; 

let detailCache = {}; 
let selectionTimeout = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(() => {
      $("#btn-sync").on("click", syncData);
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged, 
        onSelectionChange
      );
    });
  }
});

async function syncData() {
  $("#status").text("Buscando dados via Google Workspace...");
  detailCache = {}; 

  try {
    const timestamp = new Date().getTime(); 
    const res = await fetch(`${GOOGLE_WEBAPP_URL}?action=list&t=${timestamp}`, {
      method: "GET",
      redirect: "follow"
    });
    
    const responseData = await res.json();

    if (!responseData.success) {
      throw new Error(responseData.error);
    }

    const data = responseData.data;

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getUsedRange().clear();
      
      const reportValues = [["ID (Drill Down)", "Data", "Código", "Lote"]];
      
      if (data && Array.isArray(data)) {
        data.forEach((item) => {
          reportValues.push([item.Id || "", item.Data || "", item.Codigo || "", item.CodigoLote || ""]);
        });
      }

      const range = sheet.getRangeByIndexes(0, 0, reportValues.length, 4);
      range.values = reportValues;
      range.format.autofitColumns();
      await context.sync();
    });

    $("#status").html("<span style='color:green'>Sucesso! Clique no ID para ver o Razão.</span>");
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    $("#status").html(`<span style='color:red'>Erro (Lista): ${errorMessage}</span>`);
  }
}

function renderDetails(detail) {
  let htmlLines = `<b>Lote:</b> ${detail.CodigoLote || "N/A"}<br><hr>`;
  
  if (detail.Itens && detail.Itens.length > 0) {
      detail.Itens.forEach(line => {
        htmlLines += `<div class="item-line">
          <b>D:</b> ${line.ReduzidoContaDebito || "-"} | <b>C:</b> ${line.ReduzidoContaCredito || "-"} <br>
          <b>Valor:</b> R$ ${line.Valor.toLocaleString('pt-BR')} <br>
          <small>${line.Complemento || ""}</small>
        </div>`;
      });
  } else {
      htmlLines += "Nenhum item encontrado.";
  }
  
  $("#detail-content").html(htmlLines);
}

async function onSelectionChange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("values, columnIndex");
    await context.sync();
    
    if (!range.values || !range.values[0]) return;
    const selectedId = range.values[0][0];

    if (range.columnIndex === 0 && typeof selectedId === "string" && selectedId.length > 10) {
      
      if (selectionTimeout) clearTimeout(selectionTimeout);

      selectionTimeout = setTimeout(async () => {
        $("#detail-pane").show();

        if (detailCache[selectedId]) {
          renderDetails(detailCache[selectedId]);
          return; 
        }

        $("#detail-content").html("<i>Buscando detalhes via Google... ⏳</i>");

        try {
          const timestamp = new Date().getTime();
          const res = await fetch(`${GOOGLE_WEBAPP_URL}?action=detail&id=${selectedId}&t=${timestamp}`, {
            method: "GET",
            redirect: "follow"
          });
          
          const responseData = await res.json();

          if (!responseData.success) throw new Error(responseData.error);
          
          detailCache[selectedId] = responseData.data;
          renderDetails(responseData.data);

        } catch (err) {
          const errMsg = err instanceof Error ? err.message : String(err);
          $("#detail-content").html(`<span style='color:red'>Erro (Detalhe): ${errMsg}</span>`);
        }
      }, 400); 
    }
  });
}