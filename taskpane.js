// ====================================================
// 👁️ إظهار كل الصفوف أولاً
// ====================================================

const totalRows = values.length;

// إلغاء أي إخفاء قديم
sheet
  .getRangeByIndexes(
    range.rowIndex,
    0,
    totalRows,
    range.columnCount
  )
  .rowHidden = false;


// ====================================================
// 👁️ إخفاء الصفوف غير المطابقة فقط
// ====================================================

for (let i = 0; i < totalRows; i++) {

  if (!showRows.includes(i)) {

    sheet
      .getRangeByIndexes(
        range.rowIndex + i,
        0,
        1,
        range.columnCount
      )
      .rowHidden = true;
  }
}