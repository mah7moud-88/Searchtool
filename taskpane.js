Office.onReady(() => {

  document.getElementById("searchBtn").onclick =
    searchAccount;

});

function cleanValue(value) {

  return String(value || "")
    .trim()
    .replace(/\s+/g, "");

}

function removeLeadingZeros(value) {

  return cleanValue(value)
    .replace(/^0+/, "");

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

    // عمود رقم الحساب = C
    const accountRange =
      sheet.getRange("C1:C50000");

    accountRange.load("text");

    await context.sync();

    const values = accountRange.text;

    let found = false;

    for (let r = 0; r < values.length; r++) {

      const cellValue =
        cleanValue(values[r][0]);

      const normalizedCell =
        removeLeadingZeros(cellValue);

      if (
        cellValue === accountNumber ||
        normalizedCell === normalizedInput
      ) {

        found = true;

        const realRow = r + 1;

        // الاسم من العمود B
        const nameCell =
          sheet.getRange("B" + realRow);

        nameCell.load("text");

        // تحديد خلية رقم الحساب
        const foundCell =
          sheet.getRange("C" + realRow);

        foundCell.select();

        await context.sync();

        const finalName =
          nameCell.text[0][0];

        document.getElementById("result").innerText =
          "الاسم: " + finalName;

        break;
      }
    }

    if (!found) {

      document.getElementById("result").innerText =
        "رقم الحساب غير موجود";

    }

    await context.sync();

  });

}