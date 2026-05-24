Office.onReady(() => {

  const searchBtn = document.getElementById("searchBtn");
  const exportBtn = document.getElementById("exportBtn");

  if (searchBtn) searchBtn.onclick = searchAccount;
  if (exportBtn) exportBtn.onclick = exportExcel;

});

function cleanValue(value) {

  return String(value || "")
    .trim()
    .replace(/\s+/g, "");

}

let resultsData = [];

async function searchAccount() {

  const resultDiv =
    document.getElementById("result");

  let input =
    document.getElementById("accountNumber").value;

  if (!input.trim()) {

    resultDiv.innerText =
      "اكتب أرقام الحسابات";

    return;
  }

  const accounts = input
    .split("\n")
    .map(x => cleanValue(x))
    .filter(x => x);

  resultDiv.innerText = "جاري البحث...";

  resultsData = [];

  await Excel.run(async (context) => {

    const sheet =
      context.workbook.worksheets.getActiveWorksheet();

    const range =
      sheet.getRange("C1:C50000");

    range.load("text");

    await context.sync();

    const values = range.text;

    let output = "";

    for (let acc of accounts) {

      let found = false;

      const isLast6 = acc.length <= 6;

      for (let r = 0; r < values.length; r++) {

        let cellValue =
          cleanValue(values[r][0]);

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

          const realRow = r + 1;

          const nameCell =
            sheet.getRange("B" + realRow);

          const accCell =
            sheet.getRange("C" + realRow);

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
            name: name,
            account: account
          });

          found = true;
          break;
        }
      }

      if (!found) {

        output +=
          `❌ ${acc} → غير موجود\n\n`;

      }
    }

    resultDiv.innerText = output;

  });

}

async function exportExcel() {

  const resultDiv =
    document.getElementById("result");

  if (!resultsData || resultsData.length === 0) {

    resultDiv.innerText =
      "لا يوجد بيانات للتصدير";

    return;
  }

  await Excel.run(async (context) => {

    const sheet =
      context.workbook.worksheets.add("Export");

    const data = [];

    data.push(["رقم الحساب", "الاسم"]);

    resultsData.forEach(item => {

      data.push([item.account, item.name]);

    });

    const range =
      sheet.getRange(
        `A1:B${data.length}`
      );

    range.values = data;

    range.format.autofitColumns();

    sheet.activate();

    await context.sync();

  });

}