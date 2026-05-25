Office.onReady(() => {

  const searchBtn = document.getElementById("searchBtn");
  const exportBtn = document.getElementById("exportBtn");
  const clearBtn = document.getElementById("clearBtn");
  const excelFile = document.getElementById("excelFile");

  if (excelFile) {
    excelFile.addEventListener("change", importAccountsFromFile);
  }

  if (searchBtn) {
    searchBtn.onclick = searchAccount;
  }

  if (exportBtn) {
    exportBtn.onclick = exportExcel;
  }

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
// البيانات
// =========================
let resultsData = [];
let accountIndex = {};


// =========================
// بناء Index سريع
// =========================
function buildIndex(values) {

  accountIndex = {};

  for (let r = 0; r < values.length; r++) {

    const cell = cleanValue(values[r][0]);

    if (!cell) continue;

    // تخزين الرقم كما هو
    if (accountIndex[cell] === undefined) {
      accountIndex[cell] = r;
    }

    // دعم آخر 6 أرقام
    if (cell.length >= 6) {
      const last6 = cell.slice(-6);
      if (accountIndex[last6] === undefined) {
        accountIndex[last6] = r;
      }
    }
  }
}


// =========================
// رفع Excel
// =========================
function importAccountsFromFile(e) {

  const file = e.target.files[0];
  if (!file) return;

  const resultDiv = document.getElementById("result");
  resultDiv.innerText = "جاري قراءة الملف...";

  const reader = new FileReader();

  reader.onload = function (evt) {

    try {

      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      let accounts = [];

      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < rows[r].length; c++) {

          const value = String(rows[r][c] ?? "");
          if (value.trim()) accounts.push(value.trim());

        }
      }

      document.getElementById("accountNumber").value =
        accounts.join("\n");

      resultDiv.innerText =
        `✅ تم تحميل ${accounts.length} قيمة`;

      e.target.value = "";

    } catch (err) {
      resultDiv.innerText = "❌ " + err.message;
    }

  };

  reader.readAsArrayBuffer(file);
}


// =========================
// البحث السريع (INDEX)
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

    // ⚡ بناء Index مرة واحدة
    buildIndex(values);

    let output = "";

    for (let acc of accounts) {

      let rowIndex = accountIndex[acc];

      // لو مش موجود نجرب آخر 6 أرقام
      if (rowIndex === undefined && acc.length <= 6) {
        const last6 = acc.slice(-acc.length);
        rowIndex = accountIndex[last6];
      }

      if (rowIndex !== undefined) {

        const realRow = rowIndex + 1;

        const nameCell = sheet.getRange("B" + realRow);
        const accCell = sheet.getRange("C" + realRow);

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

        foundCount++;

      } else {

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

    // تلوين غير الموجود
    for (let i = 1; i < data.length; i++) {

      if (data[i][2] === "غير موجود") {
        sheet.getRange(`A${i + 1}:C${i + 1}`)
          .format.fill.color = "#F2F2F2";
      }

    }

    sheet.activate();
    await context.sync();

  });

}