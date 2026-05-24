Office.onReady(() => {

  document.getElementById("searchBtn").onclick =
    searchByLast6;

});

function cleanValue(value) {

  return String(value || "")
    .trim()
    .replace(/\s+/g, "");

}

async function searchByLast6() {

  const resultDiv =
    document.getElementById("result");

  let input =
    document.getElementById("accountNumber").value;

  input = cleanValue(input);

  if (!input) {

    resultDiv.innerText =
      "اكتب آخر 6 أرقام";

    return;
  }

  if (input.length > 6) {

    resultDiv.innerText =
      "اكتب آخر 6 أرقام فقط";

    return;
  }

  resultDiv.innerText = "جاري البحث...";

  await Excel.run(async (context) => {

    const sheet =
      context.workbook.worksheets.getActiveWorksheet();

    const range =
      sheet.getRange("C1:C50000");

    range.load("text");

    await context.sync();

    const values = range.text;

    let output = "";
    let found = false;

    for (let r = 0; r < values.length; r++) {

      let acc = cleanValue(values[r][0]);

      // نقارن آخر 6 أرقام
      let last6 = acc.slice(-6);

      if (last6 === input) {

        const realRow = r + 1;

        const nameCell =
          sheet.getRange("B" + realRow);

        const fullAccCell =
          sheet.getRange("C" + realRow);

        nameCell.load("text");
        fullAccCell.load("text");

        await context.sync();

        output +=
          `✔ ${fullAccCell.text[0][0]} → ${nameCell.text[0][0]}\n`;

        found = true;
      }
    }

    if (!found) {

      output = "لا توجد نتائج";

    }

    resultDiv.innerText = output;

  });

}