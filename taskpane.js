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
    searchMode = searchMode === "independent" ? "paired" : "independent";
    modeBtn.innerText =
      searchMode === "independent"
        ? "🔁 وضع البحث: مستقل"
        : "🔗 وضع البحث: مطابق";
  };

});

let searchMode = "independent";

let resultsData = [];
let accountIndex = {};
let nameIndex = {};
let pairIndex = {};
let last6Index = {};

// =========================
// إعدادات الأعمدة
// =========================
// النطاق المقروء من الشيت: B1:N50000
// الأعمدة داخل كل صف (values[r][...]) بالترتيب:
// 0=B (الاسم) | 1=C (الحساب) | ... | 12=N (ملاحظة)
const NOTE_COL_INDEX = 12; // عمود N

// =========================
// تنظيف
// =========================
function cleanValue(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

// =========================
// index
// =========================
function buildIndex(values) {

  accountIndex = {};
  nameIndex = {};
  pairIndex = {};
  last6Index = {};

  for (let r = 0; r < values.length; r++) {

    const name = values[r][0];
    const account = values[r][1];
    const note = values[r][NOTE_COL_INDEX];

    const cleanName = cleanValue(name);
    const cleanAcc = cleanValue(account);

    if (cleanAcc && accountIndex[cleanAcc] === undefined) {
      accountIndex[cleanAcc] = r;
    }

    if (cleanName) {
      if (!nameIndex[cleanName]) nameIndex[cleanName] = [];
      nameIndex[cleanName].push({ name, account, note });
    }

    const key = cleanName + "|" + cleanAcc;
    if (!pairIndex[key]) pairIndex[key] = r;

    if (cleanAcc && cleanAcc.length >= 6) {
      const last6 = cleanAcc.slice(-6);

      if (!last6Index[last6]) {
        last6Index[last6] = [];
      }

      last6Index[last6].push(r);
    }
  }
}

// =========================
// 🔥 البحث بالرقم (كامل أو آخر 6)
// =========================
function findAccountRows(acc) {

  const cleanAcc = cleanValue(acc);

  if (accountIndex[cleanAcc] !== undefined) {
    return [accountIndex[cleanAcc]];
  }

  if (cleanAcc.length <= 6) {
    return last6Index[cleanAcc] || [];
  }

  const last6 = cleanAcc.slice(-6);
  return last6Index[last6] || [];
}

// =========================
// رفع ملف
// =========================
function importAccountsFromFile(e) {

  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (evt) {

    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    let names = [];
    let accounts = [];

    rows.forEach(row => {
      if (row[0]) names.push(String(row[0]).trim());
      if (row[1]) accounts.push(String(row[1]).trim());
    });

    document.getElementById("customerName").value = names.join("\n");
    document.getElementById("accountNumber").value = accounts.join("\n");
  };

  reader.readAsArrayBuffer(file);
}

// =========================
// البحث
// =========================
async function searchAccount() {

  const resultDiv = document.getElementById("result");

  const accountInput = document.getElementById("accountNumber").value.trim();
  const nameInput = document.getElementById("customerName").value.trim();

  const accounts = accountInput.split("\n").filter(Boolean);
  const names = nameInput.split("\n").filter(Boolean);

  const totalCount = accounts.length + names.length;

  resultDiv.innerText = "جاري البحث...";

  resultsData = [];
  let foundCount = 0;

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("B1:N50000");

    range.load("text");
    await context.sync();

    const values = range.text;

    buildIndex(values);

    let output = "";

    // =========================
    // مستقل
    // =========================
    if (searchMode === "independent") {

      for (let acc of accounts) {

        const rows = findAccountRows(acc);

        if (rows.length) {

          rows.forEach(rowIndex => {

            const row = values[rowIndex];
            const note = row[NOTE_COL_INDEX];

            resultsData.push({
              name: row[0],
              account: row[1],
              note: note,
              status: "موجود"
            });

            output += `👤 ${row[0]}\n📌 ${row[1]}\n📝 ${note || "-"}\n\n`;
            foundCount++;
          });

        } else {

          resultsData.push({
            name: "",
            account: acc,
            note: "",
            status: "غير موجود"
          });

          output += `❌ ${acc}\n\n`;
        }
      }

      for (let name of names) {

        const records = nameIndex[cleanValue(name)] || [];

        if (records.length) {

          records.forEach(r => {

            resultsData.push({
              name: r.name,
              account: r.account,
              note: r.note,
              status: "موجود"
            });

            output += `👤 ${r.name}\n📌 ${r.account}\n📝 ${r.note || "-"}\n\n`;
            foundCount++;
          });

        } else {

          resultsData.push({
            name,
            account: "",
            note: "",
            status: "غير موجود"
          });

          output += `❌ ${name}\n\n`;
        }
      }

    }

    // =========================
    // مطابق
    // =========================
    else {

      for (let i = 0; i < Math.max(accounts.length, names.length); i++) {

        const acc = accounts[i];
        const name = names[i];

        if (!acc || !name) continue;

        const rows = findAccountRows(acc);
        let matched = false;

        for (let rowIndex of rows) {

          const row = values[rowIndex];

          if (cleanValue(row[0]) === cleanValue(name)) {

            const note = row[NOTE_COL_INDEX];

            resultsData.push({
              name: row[0],
              account: row[1],
              note: note,
              status: "موجود"
            });

            output += `👤 ${row[0]}\n📌 ${row[1]}\n📝 ${note || "-"}\n\n`;

            foundCount++;
            matched = true;
            break;
          }
        }

        if (!matched) {

          resultsData.push({
            name,
            account: acc,
            note: "",
            status: "غير مطابق"
          });

          output += `❌ غير مطابق\n👤 ${name}\n📌 ${acc}\n\n`;
        }
      }
    }

    resultDiv.innerText =
      `✅ تم العثور على ${foundCount} من أصل ${totalCount}\n\n` + output;

  });
}

// =========================
// Export
// =========================
async function exportExcel() {

  const resultDiv = document.getElementById("result");

  if (!resultsData.length) {
    resultDiv.innerText = "لا يوجد بيانات للتصدير";
    return;
  }

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.add("Export");

    const data = [["رقم الحساب", "الاسم", "ملاحظة", "الحالة"]];

    resultsData.forEach(i => {
      data.push([i.account, i.name, i.note || "", i.status]);
    });

    const range = sheet.getRange(`A1:D${data.length}`);
    range.values = data;
    range.format.autofitColumns();

    sheet.activate();
    await context.sync();

  });
}