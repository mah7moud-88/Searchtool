// =========================
// 📦 DATA STORAGE
// =========================
let DATA = [];

// ⚡ CACHE
let SEARCH_CACHE = [];

// =========================
// 📥 LOAD DATA
// =========================
async function loadData() {

  const status =
    document.getElementById("liveStatus");

  status.innerText =
    "جاري تحميل البيانات...";

  DATA = [];
  SEARCH_CACHE = [];

  try {

    await Excel.run(async (context) => {

      const sheet =
        context.workbook
          .worksheets
          .getActiveWorksheet();

      // ✅ فقط العمودين A و B
      const usedRange =
        sheet.getRange("A:B");

      usedRange.load([
        "rowCount",
        "columnCount"
      ]);

      await context.sync();

      const totalRows =
        usedRange.rowCount;

      const totalCols = 2;

      // ✅ حجم الدفعة
      const chunkSize = 500;

      // =========================
      // 🔄 LOAD IN CHUNKS
      // =========================
      for (
        let start = 0;
        start < totalRows;
        start += chunkSize
      ) {

        const rows =
          Math.min(
            chunkSize,
            totalRows - start
          );

        const range =
          sheet.getRangeByIndexes(
            start,
            0,
            rows,
            totalCols
          );

        range.load("text");

        await context.sync();

        const values =
          range.text || [];

        // =========================
        // 📦 SAVE DATA
        // =========================
        DATA.push(...values);

        // =========================
        // ⚡ BUILD CACHE
        // =========================
        for (const row of values) {

          SEARCH_CACHE.push({

            raw: row,

            name:
              normalizeText(row[0]),

            number:
              normalizeNumber(row[1])
          });
        }

        // =========================
        // 📊 STATUS
        // =========================
        status.innerText =
          `تحميل ${Math.min(start + rows, totalRows)} / ${totalRows}`;
      }
    });

    status.innerText =
      `تم تحميل ${DATA.length} صف ✅`;

  } catch (err) {

    console.error(err);

    status.innerText =
      "فشل تحميل البيانات ❌";
  }
}