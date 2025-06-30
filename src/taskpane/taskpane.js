/* global document, Office, Word, Excel */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").classList.remove("hidden");

    document.getElementById("maskButton").onclick = () => {
      if (info.host === Office.HostType.Word) {
        runMaskingWord();
      } else if (info.host === Office.HostType.Excel) {
        runMaskingExcel();
      }
    };

    document.getElementById("replaceButton").onclick = () => {
      if (info.host === Office.HostType.Word) {
        replaceSelectionWithMaskedWord();
      } else if (info.host === Office.HostType.Excel) {
        // Excel'de zaten hücreyi değiştirdik, ekstra işleme gerek yok
        alert("Excel hücresi zaten güncellendi.");
      }
    };
  }
});

let lastMaskedText = "";

// API çağrısı süresini ölçen fonksiyon
async function callMaskingAPIWithTiming(text) {
  const startTime = performance.now();
  const res = await fetch("http://localhost:5000/mask", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text })
  });

  if (!res.ok) throw new Error("Sunucu hatası: " + res.status);
  const data = await res.json();
  const endTime = performance.now();
  console.log(`Maskeleme API çağrı süresi: ${(endTime - startTime).toFixed(2)} ms`);
  return data.masked_text;
}

// === WORD ===
async function runMaskingWord() {
  const resultArea = document.getElementById("resultArea");

  await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.load("text");
    await context.sync();

    if (!sel.text.trim()) {
      resultArea.innerText = "Lütfen metin seçin.";
      return;
    }

    try {
      // Ölçümlü API çağrısı kullanılıyor
      const masked = await callMaskingAPIWithTiming(sel.text);
      lastMaskedText = masked;
      resultArea.innerText = masked;
    } catch (err) {
      resultArea.innerText = "Hata: " + err.message;
    }
  });
}

async function replaceSelectionWithMaskedWord() {
  if (!lastMaskedText) {
    alert("Önce metni maskelemelisin.");
    return;
  }

  await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.insertText(lastMaskedText, Word.InsertLocation.replace);
    await context.sync();
    alert("Seçilen metin güncellendi.");
    lastMaskedText = "";
  });
}

// === EXCEL ===
async function runMaskingExcel() {
  const resultArea = document.getElementById("resultArea");

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("text");
    await context.sync();

    const text = range.text[0][0];

    if (!text.trim()) {
      resultArea.innerText = "Lütfen bir hücre seçin.";
      return;
    }

    try {
      // Ölçümlü API çağrısı kullanılıyor
      const masked = await callMaskingAPIWithTiming(text);
      range.values = [[masked]];
      await context.sync();
      resultArea.innerText = masked;
      alert("Excel hücresi güncellendi.");
    } catch (err) {
      resultArea.innerText = "Hata: " + err.message;
    }
  });
}
