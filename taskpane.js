Office.onReady(() => {

  const searchBtn = document.getElementById("searchBtn");
  const exportBtn = document.getElementById("exportBtn");
  const clearBtn = document.getElementById("clearBtn");
  const excelFile = document.getElementById("excelFile");
  const modeBtn = document.getElementById("modeBtn");

  // =========================
  // الأحداث
  // =========================

  if (excelFile) {
    excelFile.addEventListener(
      "change",
      importAccountsFromFile
    );
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
      document.getElementById("customerName").value = "";

      document.getElementById("result").innerText = "تم المسح";

      resultsData = [];
    };
  }


  // =========================
  // تغيير وضع البحث
  // =========================

  if (modeBtn) {

    modeBtn.onclick = () => {

      searchMode =
        searchMode === "independent"
          ? "paired"
          : "independent";

      modeBtn.innerText =
        searchMode === "independent"
          ? "🔁 وضع البحث: مستقل"
          : "🔗 وضع البحث: مطابق";
    };
  }

});


// ======================================================
// المتغيرات
// ======================================================

let searchMode = "independent";

let resultsData = [];

let accountIndex = {};
let nameIndex = {};
let pairIndex = {};

let last6Index = {};
let last5Index = {};


// ======================================================
// تنظيف البيانات
// ======================================================

function cleanValue(value) {

  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}


// ======================================================
// بناء الفهارس
// ======================================================

function buildIndex(values) {

  accountIndex = {};
  nameIndex = {};
  pairIndex = {};

  last6Index = {};
  last5Index = {};


  for (let r = 0; r < values.length; r++) {

    // B = الاسم
    // C = رقم الحساب
    // N = الملاحظة

    const name = values[r][0] ?? "";
    const account = values[r][1] ?? "";
    const note = values[r][12] ?? "";

    const cleanName = cleanValue(name);
    const cleanAcc = cleanValue(account);


    // ==================================================
    // الرقم الكامل
    // ==================================================

    if (
      cleanAcc &&
      accountIndex[cleanAcc] === undefined
    ) {

      accountIndex[cleanAcc] = [];
    }

    if (cleanAcc) {

      accountIndex[cleanAcc].push(r);
    }


    // ==================================================
    // الاسم
    // نحفظ رقم الصف وليس البيانات فقط
    // ==================================================

    if (cleanName) {

      if (!nameIndex[cleanName]) {
        nameIndex[cleanName] = [];
      }

      nameIndex[cleanName].push(r);
    }


    // ==================================================
    // الاسم + الرقم
    // ==================================================

    if (cleanName && cleanAcc) {

      const key =
        cleanName + "|" + cleanAcc;

      if (pairIndex[key] === undefined) {

        pairIndex[key] = [];
      }

      pairIndex[key].push(r);
    }


    // ==================================================
    // آخر 6 أرقام
    // ==================================================

    if (
      cleanAcc &&
      cleanAcc.length >= 6
    ) {

      const last6 =
        cleanAcc.slice(-6);


      if (!last6Index[last6]) {

        last6Index[last6] = [];
      }


      last6Index[last6].push(r);
    }


    // ==================================================
    // آخر 5 أرقام
    // ==================================================

    if (
      cleanAcc &&
      cleanAcc.length >= 5
    ) {

      const last5 =
        cleanAcc.slice(-5);


      if (!last5Index[last5]) {

        last5Index[last5] = [];
      }


      last5Index[last5].push(r);
    }

  }
}


// ======================================================
// البحث بالرقم
//
// كامل
// آخر 5
// آخر 6
// ======================================================

function findAccountRows(acc) {

  const cleanAcc =
    cleanValue(acc);


  if (!cleanAcc) {
    return [];
  }


  // ==================================================
  // 1️⃣ الرقم الكامل
  // ==================================================

  if (
    accountIndex[cleanAcc] !== undefined
  ) {

    return accountIndex[cleanAcc];
  }


  // ==================================================
  // 2️⃣ آخر 5 أرقام
  // ==================================================

  if (
    cleanAcc.length === 5
  ) {

    return (
      last5Index[cleanAcc] || []
    );
  }


  // ==================================================
  // 3️⃣ آخر 6 أرقام
  // ==================================================

  if (
    cleanAcc.length === 6
  ) {

    return (
      last6Index[cleanAcc] || []
    );
  }


  // ==================================================
  // 4️⃣ رقم أكبر من 6
  // نبحث بآخر 6
  // ==================================================

  if (
    cleanAcc.length > 6
  ) {

    const last6 =
      cleanAcc.slice(-6);


    return (
      last6Index[last6] || []
    );
  }


  // أقل من 5 أرقام
  return [];
}


// ======================================================
// رفع ملف Excel
// ======================================================

function importAccountsFromFile(e) {

  const file =
    e.target.files[0];


  if (!file) {
    return;
  }


  const reader =
    new FileReader();


  reader.onload = function (evt) {

    try {

      const data =
        new Uint8Array(
          evt.target.result
        );


      const workbook =
        XLSX.read(
          data,
          {
            type: "array"
          }
        );


      const sheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ];


      const rows =
        XLSX.utils.sheet_to_json(
          sheet,
          {
            header: 1
          }
        );


      let names = [];
      let accounts = [];


      rows.forEach(row => {

        if (row[0]) {

          names.push(
            String(row[0]).trim()
          );
        }


        if (row[1]) {

          accounts.push(
            String(row[1]).trim()
          );
        }

      });


      document.getElementById(
        "customerName"
      ).value =
        names.join("\n");


      document.getElementById(
        "accountNumber"
      ).value =
        accounts.join("\n");


      document.getElementById(
        "result"
      ).innerText =
        "تم تحميل الملف";


    } catch (error) {

      document.getElementById(
        "result"
      ).innerText =
        "❌ حدث خطأ أثناء قراءة الملف";

      console.error(error);
    }

  };


  reader.readAsArrayBuffer(file);
}


// ======================================================
// إضافة نتيجة
// ======================================================

function addResult(
  row,
  status
) {

  const name =
    row[0] ?? "";

  const account =
    row[1] ?? "";

  const note =
    row[12] ?? "";


  resultsData.push({

    name: name,

    account: account,

    note: note,

    status: status
  });


  return {
    name,
    account,
    note,
    status
  };
}


// ======================================================
// تنسيق نتيجة العرض
// ======================================================

function formatResult(
  item
) {

  return (
    `👤 ${item.name}\n` +
    `📌 ${item.account}\n` +
    `📝 ${item.note}\n\n`
  );
}


// ======================================================
// البحث
// ======================================================

async function searchAccount() {

  const resultDiv =
    document.getElementById("result");


  const accountInput =
    document
      .getElementById("accountNumber")
      .value
      .trim();


  const nameInput =
    document
      .getElementById("customerName")
      .value
      .trim();


  const accounts =
    accountInput
      .split(/\r?\n/)
      .map(x => x.trim())
      .filter(Boolean);


  const names =
    nameInput
      .split(/\r?\n/)
      .map(x => x.trim())
      .filter(Boolean);


  const totalCount =
    searchMode === "independent"
      ? accounts.length + names.length
      : Math.max(
          accounts.length,
          names.length
        );


  if (
    accounts.length === 0 &&
    names.length === 0
  ) {

    resultDiv.innerText =
      "⚠️ اكتب رقم الحساب أو الاسم أولاً";

    return;
  }


  resultDiv.innerText =
    "جاري البحث...";


  resultsData = [];

  let foundCount = 0;

  let output = "";


  try {

    await Excel.run(async (context) => {


      const sheet =
        context
          .workbook
          .worksheets
          .getActiveWorksheet();


      // ==================================================
      // B إلى N
      //
      // B = row[0] الاسم
      // C = row[1] الحساب
      // N = row[12] الملاحظة
      // ==================================================

      const range =
        sheet.getRange(
          "B1:N50000"
        );


      range.load("text");

      await context.sync();


      const values =
        range.text;


      // ==================================================
      // بناء الفهارس
      // ==================================================

      buildIndex(values);


      // ==================================================
      // الوضع المستقل
      // ==================================================

      if (
        searchMode === "independent"
      ) {


        // ==================================================
        // البحث بالأرقام
        // ==================================================

        for (
          const acc of accounts
        ) {

          const rows =
            findAccountRows(acc);


          if (rows.length > 0) {


            for (
              const rowIndex of rows
            ) {

              const row =
                values[rowIndex];


              const item =
                addResult(
                  row,
                  "موجود"
                );


              output +=
                formatResult(item);


              foundCount++;
            }


          } else {


            const item = {

              name: "",

              account: acc,

              note: "",

              status: "غير موجود"
            };


            resultsData.push(item);


            output +=
              `❌ ${acc}\n\n`;
          }

        }


        // ==================================================
        // البحث بالأسماء
        // ==================================================

        for (
          const name of names
        ) {

          const cleanName =
            cleanValue(name);


          const rows =
            nameIndex[cleanName] || [];


          if (rows.length > 0) {


            for (
              const rowIndex of rows
            ) {

              // نقرأ نفس الصف
              // وبالتالي الملاحظة صحيحة 100%

              const row =
                values[rowIndex];


              const item =
                addResult(
                  row,
                  "موجود"
                );


              output +=
                formatResult(item);


              foundCount++;
            }


          } else {


            const item = {

              name: name,

              account: "",

              note: "",

              status: "غير موجود"
            };


            resultsData.push(item);


            output +=
              `❌ ${name}\n\n`;
          }

        }

      }


      // ==================================================
      // الوضع المطابق
      // ==================================================

      else {


        for (
          let i = 0;
          i < Math.max(
            accounts.length,
            names.length
          );
          i++
        ) {


          const acc =
            accounts[i] || "";


          const name =
            names[i] || "";


          // لا يمكن عمل مطابقة بدون الاثنين

          if (
            !acc ||
            !name
          ) {

            resultsData.push({

              name: name,

              account: acc,

              note: "",

              status: "بيانات ناقصة"
            });


            output +=
              `⚠️ بيانات ناقصة\n` +
              `👤 ${name}\n` +
              `📌 ${acc}\n\n`;


            continue;
          }


          // ==================================================
          // البحث بالرقم
          // ==================================================

          const rows =
            findAccountRows(acc);


          let matched =
            false;


          // ==================================================
          // التأكد من الاسم داخل نفس الصف
          // ==================================================

          for (
            const rowIndex of rows
          ) {

            const row =
              values[rowIndex];


            const rowName =
              cleanValue(
                row[0]
              );


            if (
              rowName ===
              cleanValue(name)
            ) {


              const item =
                addResult(
                  row,
                  "موجود"
                );


              output +=
                formatResult(item);


              foundCount++;

              matched = true;

              break;
            }

          }


          // ==================================================
          // غير مطابق
          // ==================================================

          if (!matched) {

            const item = {

              name: name,

              account: acc,

              note: "",

              status: "غير مطابق"
            };


            resultsData.push(item);


            output +=
              `❌ غير مطابق\n` +
              `👤 ${name}\n` +
              `📌 ${acc}\n\n`;
          }

        }

      }


      // ==================================================
      // عرض النتائج
      // ==================================================

      resultDiv.innerText =
        `✅ تم العثور على ${foundCount} من أصل ${totalCount}\n\n` +
        output;

    });


  } catch (error) {

    console.error(error);


    resultDiv.innerText =
      "❌ حدث خطأ أثناء البحث: " +
      (
        error.message ||
        "خطأ غير معروف"
      );

  }

}


// ======================================================
// تصدير النتائج إلى Excel
// ======================================================

async function exportExcel() {

  const resultDiv =
    document.getElementById("result");


  if (
    !resultsData.length
  ) {

    resultDiv.innerText =
      "لا يوجد بيانات للتصدير";

    return;
  }


  try {

    await Excel.run(async (context) => {


      // ==================================================
      // إنشاء ورقة جديدة
      // ==================================================

      const sheet =
        context
          .workbook
          .worksheets
          .add("Export");


      // ==================================================
      // العناوين
      // ==================================================

      const data = [

        [
          "رقم الحساب",
          "الاسم",
          "الملاحظة",
          "الحالة"
        ]

      ];


      // ==================================================
      // البيانات
      // ==================================================

      resultsData.forEach(item => {

        data.push([

          item.account,

          item.name,

          item.note,

          item.status

        ]);

      });


      // ==================================================
      // كتابة البيانات
      // ==================================================

      const range =
        sheet.getRange(
          `A1:D${data.length}`
        );


      range.values =
        data;


      // ==================================================
      // تنسيق العناوين
      // ==================================================

      const header =
        sheet.getRange("A1:D1");


      header.format.font.bold =
        true;


      header.format.fill.color =
        "#D9EAF7";


      // ==================================================
      // ضبط عرض الأعمدة
      // ==================================================

      range.format.autofitColumns();


      // ==================================================
      // تفعيل ورقة التصدير
      // ==================================================

      sheet.activate();


      await context.sync();

    });


    resultDiv.innerText =
      "✅ تم تصدير النتائج بنجاح";


  } catch (error) {

    console.error(error);


    resultDiv.innerText =
      "❌ حدث خطأ أثناء التصدير: " +
      (
        error.message ||
        "خطأ غير معروف"
      );

  }

}
