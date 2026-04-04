const GOOGLE_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbx6QhNdtZo9U1p-rsRhXvBPWkv58NHItnajCU3OQKFcRKfiiYxGdjPj5P7dZsa9cww9/exec"; 

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(() => {
      // Controla apenas o painel lateral
      $("#btn-sync").on("click", syncData);
    });
  }
});

async function syncData() {
  // Lógica do painel lateral (Razão Contábil, etc)
  console.log("Botão de baixar clicado!");
}
