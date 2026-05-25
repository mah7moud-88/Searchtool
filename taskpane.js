Office.onReady(() => {

  const searchBtn =
    document.getElementById("searchBtn");

  const exportBtn =
    document.getElementById("exportBtn");

  const clearBtn =
    document.getElementById("clearBtn");

  const excelFile =
    document.getElementById("excelFile");


  // =========================
  // رفع الملف
  // =========================
  if (excelFile) {

    excelFile.addEventListener(
      "change",
      importAccountsFromFile
    );

  }


  // =========================
  // البحث
  // =========================
  if (searchBtn) {

    searchBtn.onclick =
      searchAccount;

  }


  // =========================
  // التصدير
  // =========================
  if (exportBtn) {

    exportBtn.onclick =
      exportExcel;

  }


  // =========================
  // المسح
  // =========================
  if (clearBtn) {

    clearBtn.onclick = () => {

      document.getElementById(
        "accountNumber"
      ).value = "";

      document.getElementById(
        "result"
      ).innerText =
        "تم المسح";

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


let resultsData = [];


// =========================
// استيراد القيم من Excel
// =========================
function importAccountsFromFile(e) {

  const file =
    e.target.files[0];

  if (!file) return;

  const resultDiv =
    document.getElementById("result");

  resultDiv.innerText =
    "جاري قراءة الملف...";

  const reader =
    new FileReader();

  reader.onerror = function() {

    resultDiv.innerText =
      "❌ المتصفح فشل في قراءة الملف";

  };

  reader.onload = function(evt) {

    try {

      if (typeof XLSX === "undefined") {

        resultDiv.innerText =
          "❌ مكتبة XLSX غير محملة";

        return;

      }

      const data =
        new Uint8Array(
          evt.target.result
        );

      const workbook =
        XLSX.read(data, {

          type: "array",

          raw: false,

          cellText: true,

          cellDates: true

        });

      const firstSheet =
        workbook.SheetNames[0];

      const worksheet =
        workbook.Sheets[firstSheet];

      const rows =
        XLSX.utils.sheet_to_json(
          worksheet,
          {
            header: 1,
            raw: false
          }
        );

      let accounts = [];

      for (let r = 0; r < rows.length; r++) {

        const row = rows[r];

        for (let c = 0; c < row.length; c++) {

          const value =
            String(row[c] ?? "");

          if (value.trim() !== "") {

            accounts.push(
              value.trim()
            );

          }

        }

      }

      document.getElementById(
        "accountNumber"
      ).value =
        accounts.join("\n");

      resultDiv.innerText =
        `✅ تم تحميل ${accounts.length} قيمة`;

    } catch (error) {

      console.error(error);

      resultDiv.innerText =
        "❌ " + error.message;

    }

  };

  reader.readAsArrayBuffer(file);

}


// =========================
// البحث
// =========================
async function searchAccount() {

  const resultDiv =
    document.getElementById("result");

  let input =
    document.getElementById(
      "accountNumber"
    ).value;

  if (!input.trim()) {

    resultDiv.innerText =
      "اكتب أرقام الحسابات";

    return;

  }

  const accounts = input
    .split("\n")
    .map(x => cleanValue(x))
    .filter(x => x);

  resultDiv.innerText =
    "جاري البحث...";

  resultsData = [];

  let foundCount = 0;

  await Excel.run(async (context) => {

    const sheet =
      context.workbook
        .worksheets
        .getActiveWorksheet();

    const range =
      sheet.getRange("C1:C50000");

    range.load("text");

    await context.sync();

    const values =
      range.text;

    let output = "";

    for (let acc of accounts) {

      let found = false;

      const isLast6 =
        acc.length <= 6;

      for (let r = 0; r < values.length; r++) {

        let cellValue =
          cleanValue(values[r][0]);

        let match = false;

        if (isLast6) {

          if (
            cellValue.slice(-acc.length)
            === acc
          ) {

            match = true;

          }

        } else {

          if (cellValue === acc) {

            match = true;

          }

        }

        if (match) {

          const realRow =
            r + 1;

          const nameCell =
            sheet.getRange(
              "B" + realRow
            );

          const accCell =
            sheet.getRange(
              "C" + realRow
            );

          nameCell.load("text");
          accCell.load("text");

          await context.sync();

          const name =
            nameCell.text[0][0];

          const account =
            accCell.text[0][0];

          output +=
            `👤 ${name}\n📌 ${account}\n\n`;

          resultsData.push({

            name,
            account

          });

          found = true;

          foundCount++;

          break;

        }

      }

      if (!found) {

        output +=
          `❌ ${acc} → غير موجود\n\n`;

      }

    }

    output =
      `✅ تم العثور على ${foundCount} من أصل ${accounts.length}\n\n`
      + output;

    resultDiv.innerText =
      output;

  });

}


// =========================
// تصدير النتائج
// =========================
async function exportExcel() {

  const resultDiv =
    document.getElementById("result");

  if (!resultsData.length) {

    resultDiv.innerText =
      "لا يوجد بيانات للتصدير";

    return;

  }

  await Excel.run(async (context) => {

    const sheet =
      context.workbook
        .worksheets
        .add("Export");

    const data = [

      ["رقم الحساب", "الاسم"]

    ];

    resultsData.forEach(item => {

      data.push([

        item.account,
        item.name

      ]);

    });

    const range =
      sheet.getRange(
        `A1:B${data.length}`
      );

    range.values = data;

    range.format
      .autofitColumns();

    sheet.activate();

    await context.sync();

  });

}