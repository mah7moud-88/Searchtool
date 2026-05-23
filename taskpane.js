await Excel.run(async (context) => {

  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const used = sheet.getUsedRange();

  used.load(["rowIndex", "rowCount"]);
  await context.sync();

  const nameSet = new Set(
    document.getElementById("nameBox").value
      .toLowerCase()
      .split("\n")
      .map(v => v.trim())
      .filter(Boolean)
  );

  const rawNumbers = document.getElementById("numberBox")
    .value
    .split("\n")
    .map(v => v.trim())
    .filter(Boolean);

  const numberSet = new Set(rawNumbers.map(v =>
    String(v).replace(/\D/g, "").slice(-6)
  ));

  // ❗ نجيب العمودين فقط (B و C) بدون تحميل الشيت كله
  const range = sheet.getRangeByIndexes(
    used.rowIndex,
    1,
    used.rowCount,
    2
  );

  range.load("values");
  await context.sync();

  const values = range.values;

  let showRows = [];

  for (let i = 0; i < values.length; i++) {

    const name = (values[i][0] || "").toLowerCase().trim();
    const number = String(values[i][1] || "").replace(/\D/g, "").slice(-6);

    if (
      (nameSet.size === 0 || nameSet.has(name)) &&
      (numberSet.size === 0 || numberSet.has(number))
    ) {
      showRows.push(i);
    }
  }

  // 🔥 الحل النهائي: لا rowHidden ولا loops على الشيت
  // نستخدم AutoFilter بدل الإخفاء اليدوي

  const tableRange = sheet.getUsedRange();
  tableRange.load();
  await context.sync();

  tableRange.autoFilter.apply(
    [{
      columnIndex: 1,
      criterion: nameSet.size ? Array.from(nameSet) : undefined
    }]
  );

});