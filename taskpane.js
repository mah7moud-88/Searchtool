Office.onReady(() => {

  const searchBtn = document.getElementById("searchBtn");
  const exportBtn = document.getElementById("exportBtn");
  const clearBtn = document.getElementById("clearBtn");
  const excelFile = document.getElementById("excelFile");


  // رفع الملف
  if (excelFile) {
    excelFile.addEventListener("change", importAccountsFromFile);
  }

  // البحث
  if (searchBtn) {
    searchBtn.onclick = searchAccount;
  }

  // التصدير
  if (exportBtn) {
    exportBtn.onclick = exportExcel;
  }

  // المسح
  if (clearBtn) {
    clearBtn.onclick = () => {

      document.getElementById("accountNumber").value = "";
      document.getElementById("result").innerText = "تم المسح";
      resultsData = [];

    };
  }

});


// =========================
// تنظيف النص
// =========================
function cleanValue(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "");
}


// =========================
// تخزين النتائج
// =========================
let resultsData = [];


// =========================
// رفع ملف Excel
// =========================
function importAccountsFromFile(e) {

  const file = e.target.files[0];
  if (!file) return;

  const resultDiv = document.getElementById("result");
  resultDiv.innerText = "جاري قراءة الملف...";

  const reader = new FileReader();

  reader.onerror = function () {
    resultDiv.innerText = "❌ المتصفح فشل في قراءة الملف";
  };

  reader.onload = function (evt) {

    try {

      if (typeof XLSX === "undefined") {
        resultDiv.innerText = "❌ مكتبة XLSX غير محملة";
        return;
      }

      const data = new Uint8Array(evt.target.result);

      const workbook = XLSX.read(data, {
        type: "array",
        raw: false,
        cellText: true,
        cellDates: true
      });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1
      });

      let accounts = [];

      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < rows[r].length; c++) {

          const value = String(rows[r][c] ?? "");
          if (value.trim()) {
            accounts.push(value.trim());
          }

        }
      }

      document.getElementById("accountNumber").value =
        accounts.join("\n");

      resultDiv.innerText =
        `✅ تم تحميل ${accounts.length} قيمة`;

      e.target.value = "";

    } catch (error) {
      resultDiv.innerText = "❌ " + error.message;
    }

  };

  reader.readAsArrayBuffer(file);
}


// =========================
// البحث
// =========================
async function searchAccount() {

  const resultDiv = document.getElementById("result");

  let input = document.getElementById("accountNumber").value;

  if (!input.trim()) {
    resultDiv.innerText = "اكتب أرقام الحسابات";
    return;
  }

  const accounts = input
    .split("\n")
    .map(cleanValue)
    .filter(x => x);

  resultDiv.innerText = "جاري البحث...";

  resultsData = [];

  let foundCount = 0;

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("C1:C50000");

    range.load("text");
    await context.sync();

    const values = range.text;

    let output = "";

    for (let acc of accounts) {

      let found = false;
      const isLast6 = acc.length <= 6;

      for (let r = 0; r < values.length; r++) {

        let cellValue = cleanValue(values[r][0]);
        let match = false;

        if (isLast6) {
          if (cellValue.slice(-acc.length) === acc) {
            match = true;
          }
        } else {
          if (cellValue === acc) {
            match = true;
          }
        }

        if (match) {

          const rowIndex = r + 1;

          const nameCell = sheet.getRange("B" + rowIndex);
          const accCell = sheet.getRange("C" + rowIndex);

          nameCell.load("text");
          accCell.load("text");

          await context.sync();

          const name = nameCell.text[0][0];
          const account = accCell.text[0][0];

          resultsData.push({
            name,
            account,
            status: "موجود"
          });

          output += `👤 ${name}\n📌 ${account}\n\n`;

          found = true;
          foundCount++;

          break;
        }
      }

      if (!found) {

        resultsData.push({
          name: "",
          account: acc,
          status: "غير موجود"
        });

        output += `❌ ${acc} → غير موجود\n\n`;
      }
    }

    resultDiv.innerText =
      `✅ تم العثور على ${foundCount} من أصل ${accounts.length}\n\n`
      + output;

  });

}


// =========================
// التصدير
// =========================
async function exportExcel() {

  const resultDiv = document.getElementById("result");

  if (!resultsData.length) {
    resultDiv.innerText = "لا يوجد بيانات للتصدير";
    return;
  }

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.add("Export");

    const data = [
      ["رقم الحساب", "الاسم", "الحالة"]
    ];

    resultsData.forEach(item => {
      data.push([
        item.account,
        item.name,
        item.status
      ]);
    });

    const range = sheet.getRange(`A1:C${data.length}`);
    range.values = data;
    range.format.autofitColumns();

    // تلوين غير الموجود بالأحمر
    for (let i = 1; i < data.length; i++) {

      if (data[i][2] === "غير موجود") {

        const row = sheet.getRange(`A${i + 1}:C${i + 1}`);
        row.format.fill.color = "#FFCCCC";

      }

    }

    sheet.activate();
    await context.sync();

  });

}