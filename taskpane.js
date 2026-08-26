Office.onReady(() => {

  const searchBtn = document.getElementById("searchBtn");
  const exportBtn = document.getElementById("exportBtn");
  const clearBtn = document.getElementById("clearBtn");
  const excelFile = document.getElementById("excelFile");
  const modeBtn = document.getElementById("modeBtn");

  excelFile.addEventListener("change", importAccountsFromFile);

  searchBtn.onclick = searchAccount;
  exportBtn.onclick = exportExcel;

  clearBtn.onclick = () => {

    document.getElementById("accountNumber").value = "";
    document.getElementById("customerName").value = "";

    document.getElementById("result").innerText = "تم المسح";

    resultsData = [];
  };

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


// =========================
// المتغيرات
// =========================

let searchMode = "independent";

let resultsData = [];

let accountIndex = {};
let nameIndex = {};
let pairIndex = {};

let last6Index = {};
let last4Index = {};


// =========================
// تنظيف البيانات
// =========================

function cleanValue(value) {

  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}


// =========================
// بناء Index
// =========================

function buildIndex(values) {

  accountIndex = {};
  nameIndex = {};
  pairIndex = {};

  last6Index = {};
  last4Index = {};

  for (let r = 0; r < values.length; r++) {

    const name = values[r][0];
    const account = values[r][1];

    const cleanName = cleanValue(name);
    const cleanAcc = cleanValue(account);


    // =========================
    // الرقم الكامل
    // =========================

    if (
      cleanAcc &&
      accountIndex[cleanAcc] === undefined
    ) {

      accountIndex[cleanAcc] = r;
    }


    // =========================
    // الاسم
    // =========================

    if (cleanName) {

      if (!nameIndex[cleanName]) {
        nameIndex[cleanName] = [];
      }

      nameIndex[cleanName].push({
        name: name,
        account: account
      });
    }


    // =========================
    // الاسم + الرقم
    // =========================

    const key =
      cleanName + "|" + cleanAcc;

    if (!pairIndex[key]) {
      pairIndex[key] = r;
    }


    // =========================
    // آخر 6 أرقام
    // =========================

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


    // =========================
    // آخر 4 أرقام
    // =========================

    if (
      cleanAcc &&
      cleanAcc.length >= 4
    ) {

      const last4 =
        cleanAcc.slice(-4);

      if (!last4Index[last4]) {
        last4Index[last4] = [];
      }

      last4Index[last4].push(r);
    }

  }
}


// =========================
// البحث بالرقم
//
// كامل
// آخر 6
// آخر 4
// =========================

function findAccountRows(acc) {

  const cleanAcc =
    cleanValue(acc);


  // =========================
  // 1️⃣ البحث بالرقم الكامل
  // =========================

  if (
    accountIndex[cleanAcc] !== undefined
  ) {

    return [
      accountIndex[cleanAcc]
    ];
  }


  // =========================
  // 2️⃣ البحث بآخر 4 أرقام
  // =========================

  if (cleanAcc.length === 4) {

    return (
      last4Index[cleanAcc] || []
    );
  }


  // =========================
  // 3️⃣ البحث بآخر 6 أرقام
  // =========================

  if (cleanAcc.length === 6) {

    return (
      last6Index[cleanAcc] || []
    );
  }


  // =========================
  // 4️⃣ لو المستخدم كتب
  // أكثر من 6 أرقام
  // =========================

  if (cleanAcc.length > 6) {

    const last6 =
      cleanAcc.slice(-6);

    return (
      last6Index[last6] || []
    );
  }


  // =========================
  // أي رقم أقل من 4
  // لا يوجد بحث
  // =========================

  return [];
}


// =========================
// رفع ملف Excel
// =========================

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


// =========================
// البحث
// =========================

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


    const range =
      sheet.getRange(
        "B1:C50000"
      );


    range.load("text");

    await context.sync();


    const values =
      range.text;


    // =========================
    // بناء الفهارس
    // =========================

    buildIndex(values);


    let output = "";


    // =========================
    // وضع البحث المستقل
    // =========================

    if (
      searchMode === "independent"
    ) {


      // =========================
      // البحث بالأرقام
      // =========================

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


              resultsData.push({

                name: row[0],

                account: row[1],

                status: "موجود"

              });


              output +=
                `👤 ${row[0]}\n` +
                `📌 ${row[1]}\n\n`;


              foundCount++;

            }
          );


        } else {


          resultsData.push({

            name: "",

            account: acc,

            status: "غير موجود"

          });


          output +=
            `❌ ${acc}\n\n`;
        }

      }


      // =========================
      // البحث بالأسماء
      // =========================

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

              resultsData.push({

                name: r.name,

                account: r.account,

                status: "موجود"

              });


              output +=
                `👤 ${r.name}\n` +
                `📌 ${r.account}\n\n`;


              foundCount++;

            }
          );


        } else {


          resultsData.push({

            name: name,

            account: "",

            status: "غير موجود"

          });


          output +=
            `❌ ${name}\n\n`;
        }

      }

    }


    // =========================
    // وضع البحث المطابق
    // =========================

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


        for (
          let rowIndex of rows
        ) {

          const row =
            values[rowIndex];


          if (
            cleanValue(row[0]) ===
            cleanValue(name)
          ) {


            resultsData.push({

              name: row[0],

              account: row[1],

              status: "موجود"

            });


            output +=
              `👤 ${row[0]}\n` +
              `📌 ${row[1]}\n\n`;


            foundCount++;

            matched = true;

            break;
          }

        }


        if (!matched) {


          resultsData.push({

            name: name,

            account: acc,

            status: "غير مطابق"

          });


          output +=
            `❌ غير مطابق\n` +
            `👤 ${name}\n` +
            `📌 ${acc}\n\n`;
        }

      }

    }


    // =========================
    // عرض النتيجة
    // =========================

    resultDiv.innerText =
      `✅ تم العثور على ${foundCount} من أصل ${totalCount}\n\n` +
      output;

  });

}


// =========================
// Export
// =========================

async function exportExcel() {

  const resultDiv =
    document.getElementById("result");


  if (!resultsData.length) {

    resultDiv.innerText =
      "لا يوجد بيانات للتصدير";

    return;
  }


  await Excel.run(async (context) => {


    const sheet =
      context
        .workbook
        .worksheets
        .add("Export");


    const data = [
      [
        "رقم الحساب",
        "الاسم",
        "الحالة"
      ]
    ];


    resultsData.forEach(i => {

      data.push([
        i.account,
        i.name,
        i.status
      ]);

    });


    const range =
      sheet.getRange(
        `A1:C${data.length}`
      );


    range.values =
      data;


    range.format.autofitColumns();


    sheet.activate();


    await context.sync();

  });

}
