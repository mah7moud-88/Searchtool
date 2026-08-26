Office.onReady(() => {

  const searchBtn = document.getElementById("searchBtn");
  const exportBtn = document.getElementById("exportBtn");
  const clearBtn = document.getElementById("clearBtn");
  const excelFile = document.getElementById("excelFile");
  const modeBtn = document.getElementById("modeBtn");

  excelFile.addEventListener("change", importAccountsFromFile);

  searchBtn.onclick = searchAccount;
  exportBtn.onclick = exportExcel;


  // =========================
  // زر المسح
  // =========================

  clearBtn.onclick = () => {

    document.getElementById("accountNumber").value = "";
    document.getElementById("customerName").value = "";

    document.getElementById("result").innerText = "تم المسح";

    resultsData = [];
  };


  // =========================
  // تغيير وضع البحث
  // =========================

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

  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}


// ======================================================
// بناء Index
// ======================================================

function buildIndex(values) {

  accountIndex = {};
  nameIndex = {};
  pairIndex = {};

  last6Index = {};
  last5Index = {};


  for (let r = 0; r < values.length; r++) {

    const name = values[r][0];
    const account = values[r][1];

    const cleanName = cleanValue(name);
    const cleanAcc = cleanValue(account);


    // ==================================================
    // الرقم الكامل
    // ==================================================

    if (
      cleanAcc &&
      accountIndex[cleanAcc] === undefined
    ) {

      accountIndex[cleanAcc] = r;
    }


    // ==================================================
    // الاسم
    // ==================================================

    if (cleanName) {

      if (!nameIndex[cleanName]) {
        nameIndex[cleanName] = [];
      }

      nameIndex[cleanName].push({

        name: name,

        account: account

      });
    }


    // ==================================================
    // الاسم + الرقم
    // ==================================================

    const key =
      cleanName + "|" + cleanAcc;


    if (!pairIndex[key]) {

      pairIndex[key] = r;

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
// آخر 6
// آخر 5
// ======================================================

function findAccountRows(acc) {

  const cleanAcc =
    cleanValue(acc);


  // ==================================================
  // 1️⃣ البحث بالرقم الكامل
  // ==================================================

  if (
    accountIndex[cleanAcc] !== undefined
  ) {

    return [
      accountIndex[cleanAcc]
    ];

  }


  // ==================================================
  // 2️⃣ البحث بآخر 5 أرقام
  // ==================================================

  if (
    cleanAcc.length === 5
  ) {

    return (
      last5Index[cleanAcc] || []
    );

  }


  // ==================================================
  // 3️⃣ البحث بآخر 6 أرقام
  // ==================================================

  if (
    cleanAcc.length === 6
  ) {

    return (
      last6Index[cleanAcc] || []
    );

  }


  // ==================================================
  // 4️⃣ لو الرقم أكثر من 6 أرقام
  // يبحث بآخر 6
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


  // ==================================================
  // أقل من 5 أرقام
  // ==================================================

  return [];

}


// ======================================================
// رفع ملف Excel
// ======================================================

function importAccountsFromFile(e) {

  const file =
    e.target.files[0];


  if (!file) return;


  const reader =
    new FileReader();


  reader.onload = function (evt) {

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

  };


  reader.readAsArrayBuffer(file);

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
      .split("\n")
      .map(x => x.trim())
      .filter(Boolean);


  const names =
    nameInput
      .split("\n")
      .map(x => x.trim())
      .filter(Boolean);


  const totalCount =
    accounts.length +
    names.length;


  resultDiv.innerText =
    "جاري البحث...";


  resultsData = [];

  let foundCount = 0;


  await Excel.run(async (context) => {


    const sheet =
      context
        .workbook
        .worksheets
        .getActiveWorksheet();


    // ==================================================
    // مهم:
    // B = الاسم
    // C = رقم الحساب
    // N = الملاحظة
    //
    // لذلك نقرأ من B إلى N
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


    let output = "";


    // ==================================================
    // وضع البحث المستقل
    // ==================================================

    if (
      searchMode === "independent"
    ) {


      // ==================================================
      // البحث بالأرقام
      // ==================================================

      for (
        let acc of accounts
      ) {


        const rows =
          findAccountRows(acc);


        if (rows.length) {


          rows.forEach(
            rowIndex => {


              const row =
                values[rowIndex];


              // B = row[0]
              // C = row[1]
              // N = row[12]

              const name =
                row[0] || "";


              const account =
                row[1] || "";


              const note =
                row[12] || "";


              // ==================================================
              // حفظ النتيجة
              // ==================================================

              resultsData.push({

                name: name,

                account: account,

                note: note,

                status: "موجود"

              });


              // ==================================================
              // عرض النتيجة
              // ==================================================

              output +=
                `👤 ${name}\n` +
                `📌 ${account}\n` +
                `📝 ${note}\n\n`;


              foundCount++;

            }
          );


        } else {


          // ==================================================
          // غير موجود
          // ==================================================

          resultsData.push({

            name: "",

            account: acc,

            note: "",

            status: "غير موجود"

          });


          output +=
            `❌ ${acc}\n\n`;

        }

      }


      // ==================================================
      // البحث بالأسماء
      // ==================================================

      for (
        let name of names
      ) {


        const records =
          nameIndex[
            cleanValue(name)
          ] || [];


        if (records.length) {


          records.forEach(
            r => {


              // الاسم والرقم موجودان في الـ index
              // لكن الملاحظة تحتاج الصف الأصلي.
              // لذلك نبحث عن الصف باستخدام الرقم.

              const rows =
                findAccountRows(
                  r.account
                );


              let note = "";


              if (rows.length) {

                const originalRow =
                  values[rows[0]];

                note =
                  originalRow[12] || "";
              }


              resultsData.push({

                name: r.name,

                account: r.account,

                note: note,

                status: "موجود"

              });


              output +=
                `👤 ${r.name}\n` +
                `📌 ${r.account}\n` +
                `📝 ${note}\n\n`;


              foundCount++;

            }
          );


        } else {


          resultsData.push({

            name: name,

            account: "",

            note: "",

            status: "غير موجود"

          });


          output +=
            `❌ ${name}\n\n`;

        }

      }

    }


    // ==================================================
    // وضع البحث المطابق
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
          accounts[i];


        const name =
          names[i];


        if (!acc || !name) {

          continue;

        }


        const rows =
          findAccountRows(acc);


        let matched = false;


        // ==================================================
        // البحث داخل النتائج المطابقة
        // ==================================================

        for (
          let rowIndex of rows
        ) {


          const row =
            values[rowIndex];


          const rowName =
            row[0] || "";


          const rowAccount =
            row[1] || "";


          const note =
            row[12] || "";


          // ==================================================
          // مطابقة الاسم
          // ==================================================

          if (
            cleanValue(rowName) ===
            cleanValue(name)
          ) {


            resultsData.push({

              name: rowName,

              account: rowAccount,

              note: note,

              status: "موجود"

            });


            output +=
              `👤 ${rowName}\n` +
              `📌 ${rowAccount}\n` +
              `📝 ${note}\n\n`;


            foundCount++;

            matched = true;

            break;

          }

        }


        // ==================================================
        // غير مطابق
        // ==================================================

        if (!matched) {


          resultsData.push({

            name: name,

            account: acc,

            note: "",

            status: "غير مطابق"

          });


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

}


// ======================================================
// Export
// ======================================================

async function exportExcel() {

  const resultDiv =
    document.getElementById("result");


  // ==================================================
  // لا توجد نتائج
  // ==================================================

  if (
    !resultsData.length
  ) {

    resultDiv.innerText =
      "لا يوجد بيانات للتصدير";

    return;

  }


  await Excel.run(async (context) => {


    // ==================================================
    // إنشاء ورقة التصدير
    // ==================================================

    const sheet =
      context
        .workbook
        .worksheets
        .add("Export");


    // ==================================================
    // عناوين الأعمدة
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
    // إضافة النتائج
    // ==================================================

    resultsData.forEach(i => {

      data.push([

        i.account,

        i.name,

        i.note,

        i.status

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
    // ضبط عرض الأعمدة
    // ==================================================

    range.format.autofitColumns();


    // ==================================================
    // تفعيل ورقة التصدير
    // ==================================================

    sheet.activate();


    await context.sync();

  });

}
