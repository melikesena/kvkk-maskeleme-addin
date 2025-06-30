/* global Office, Word, Excel, document */

Office.onReady(() => {
  document.getElementById("maskSelectionBtn").onclick = maskSelectedText;
});

async function maskSelectedText() {
  const status = document.getElementById("status");
  status.style.color = "black";
  status.textContent = "İşlem yapılıyor...";

  try {
    if (Office.context.host === Office.HostType.Word) {
      // Word için maskeleme işlemi
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          status.style.color = "red";
          status.textContent = "Lütfen önce Word'de bir metin seçin.";
          return;
        }

        // Sunucuya seçili metni gönder
        const maskedText = await callMaskingAPI(selection.text);

        // Seçili metni maskeleme sonucu ile değiştir
        selection.insertText(maskedText, Word.InsertLocation.replace);
        await context.sync();

        status.style.color = "green";
        status.textContent = "Metin başarıyla maskelendi.";
      });

    } else if (Office.context.host === Office.HostType.Excel) {
      // Excel için maskeleme işlemi
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("values");
        await context.sync();

        const text = range.values[0][0];
        if (!text || text.trim().length === 0) {
          status.style.color = "red";
          status.textContent = "Lütfen önce Excel'de bir hücre seçin.";
          return;
        }

        // Sunucuya seçili hücre metnini gönder
        const maskedText = await callMaskingAPI(text);

        // Hücreyi maskeleme sonucu ile değiştir
        range.values = [[maskedText]];
        await context.sync();

        status.style.color = "green";
        status.textContent = "Hücre metni başarıyla maskelendi.";
      });

    } else {
      status.style.color = "red";
      status.textContent = "Bu eklenti sadece Word ve Excel üzerinde çalışır.";
    }
  } catch (error) {
    status.style.color = "red";
    status.textContent = "Hata: " + error.message;
  }
}

async function callMaskingAPI(text) {
  const response = await fetch("http://localhost:5000/mask", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text }),
  });

  if (!response.ok) {
    throw new Error("Sunucu hatası: " + response.status);
  }

  const data = await response.json();
  return data.masked_text;
}
