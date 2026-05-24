Office.onReady(() => {

  document.getElementById("searchBtn").onclick =
    searchAccount;

});

function cleanValue(value) {

  return String(value || "")
    .trim()
    .replace(/\s+/g, "");

}

// حذف الأصفار من بداية الرقم
function removeLeadingZeros(value) {

  return cleanValue(value).replace(/^0+/, "");

}

async function searchAccount() {

  const accountNumber =
    cleanValue(
      document.getElementById("accountNumber").value
    );

  const normalizedInput =
    removeLeadingZeros(accountNumber);

  if (!accountNumber) {

    document.getElementById("result").innerText =
      "اكتب رقم الحساب";

    return;
  }

  await Excel.run(async (context) => {

    const sheet =
      context.workbook.worksheets.getActiveWorksheet();

    const usedRange = sheet.getUsedRange();

    usedRange.load([
      "values",
      "rowIndex"
    ]);

    await context.sync();

    const values = usedRange.values;

    let found = false;

    for (let r = 0; r < values.length; r++) {

      for (let c = 0; c < values[r].length; c++) {

        const cellValue =
          cleanValue(values[r][c]);

        const normalizedCell =
          removeLeadingZeros(cellValue);

        // مقارنة بالقيمتين
        if (
          cellValue === accountNumber ||
          normalizedCell === normalizedInput
        ) {

          found = true;

          // الصف الحقيقي
          const realRow =
            usedRange.rowIndex + r + 1;

          // الاسم من العمود B
          const nameCell =
            sheet.getRange("B" + realRow);

          nameCell.load("text");

          // تحديد الخلية
          const foundCell =
            usedRange.getCell(r, c);

          foundCell.select();

          await context.sync();

          const finalName =
            nameCell.text[0][0];

          document.getElementById("result").innerText =
            "الاسم: " + finalName;

          break;
        }
      }

      if (found) break;
    }

    if (!found) {

      document.getElementById("result").innerText =
        "رقم الحساب غير موجود";

    }

    await context.sync();

  });

}