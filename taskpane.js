Office.onReady(() => {

  document.getElementById("searchBtn").onclick =
    searchAccount;

});

function cleanValue(value) {

  return String(value || "")
    .replace(/\s+/g, "")
    .trim();

}

function removeLeadingZeros(value) {

  return cleanValue(value)
    .replace(/^0+/, "");

}

async function searchAccount() {

  const resultDiv =
    document.getElementById("result");

  const accountNumber =
    cleanValue(
      document.getElementById("accountNumber").value
    );

  const normalizedInput =
    removeLeadingZeros(accountNumber);

  if (!accountNumber) {

    resultDiv.innerText =
      "اكتب رقم الحساب";

    return;
  }

  resultDiv.innerText = "جاري البحث...";

  await Excel.run(async (context) => {

    const sheet =
      context.workbook.worksheets.getActiveWorksheet();

    // آخر صف مستخدم في العمود C
    const usedRange =
      sheet.getUsedRange();

    usedRange.load("rowCount");

    await context.sync();

    const lastRow =
      usedRange.rowCount;

    // تحميل عمود C فقط
    const accountRange =
      sheet.getRange(
        `C1:C${lastRow}`
      );

    accountRange.load("text");

    await context.sync();

    const values =
      accountRange.text;

    let foundRow = -1;

    for (let r = 0; r < values.length; r++) {

      const cellValue =
        cleanValue(values[r][0]);

      if (!cellValue) continue;

      const normalizedCell =
        removeLeadingZeros(cellValue);

      if (
        cellValue === accountNumber ||
        normalizedCell === normalizedInput
      ) {

        foundRow = r + 1;
        break;

      }
    }

    if (foundRow === -1) {

      resultDiv.innerText =
        "رقم الحساب غير موجود";

      return;
    }

    // تحميل الاسم فقط
    const nameCell =
      sheet.getRange(
        `B${foundRow}`
      );

    nameCell.load("text");

    // تحديد الخلية
    sheet
      .getRange(`C${foundRow}`)
      .select();

    await context.sync();

    resultDiv.innerText =
      "الاسم: " +
      nameCell.text[0][0];

  });

}