Office.onReady(async () => {

  // =========================
  // 🔢 GET LAST 6 DIGITS
  // =========================
  function getLast6(value) {

    return String(value || "")
      .replace(/\D/g, "")
      .slice(-6);
  }

  // =========================
  // ⚡ DATA STORAGE
  // =========================
  let DATA = [];

  // =========================
  // 📥 LOAD DATA FROM EXCEL
  // =========================
  async function loadData() {

    await Excel.run(async (context) => {

      const sheet =
        context.workbook
          .worksheets
          .getActiveWorksheet();

      const range =
        sheet.getUsedRange();

      range.load("values");

      await context.sync();

      DATA = range.values || [];
    });

    document.getElementById("liveStatus").innerText =
      "جاهز للبحث 🚀";
  }

  await loadData();

  // =========================
  // 📊 COUNTER
  // =========================
  function updateCounter() {

    const names =
      document.getElementById("nameBox")
        .value
        .split("\n")
        .map(v => v.trim())
        .filter(Boolean);

    const numbers =
      document.getElementById("numberBox")
        .value
        .split("\n")
        .map(v => v.trim())
        .filter(Boolean);

    document.getElementById("inputCount").innerText =
      names.length + numbers.length;
  }

  window.updateCounter = updateCounter;

  document.getElementById("nameBox")
    .addEventListener("input", updateCounter);

  document.getElementById("numberBox")
    .addEventListener("input", updateCounter);

  // =========================
  // 📁 IMPORT FILE
  // =========================
  document.getElementById("fileInput")
    .addEventListener("change", function (e) {

      const file = e.target.files[0];

      if (!file) return;

      const reader = new FileReader();

      reader.onload = function (event) {

        try {

          const workbook =
            XLSX.read(event.target.result, {
              type: "binary"
            });

          const sheet =
            workbook.Sheets[
              workbook.SheetNames[0]
            ];

          // 🔥 STORE FILE DATA
          DATA =
            XLSX.utils.sheet_to_json(sheet, {
              header: 1
            });

          // 🔥 FILL NAMES FROM COLUMN B
          document.getElementById("nameBox").value =
            DATA.slice(1)
              .map(r => String(r[1] || ""))
              .join("\n");

          // 🔥 FILL NUMBERS FROM COLUMN C
          document.getElementById("numberBox").value =
            DATA.slice(1)
              .map(r => String(r[2] || ""))
              .join("\n");

          updateCounter();

          document.getElementById("liveStatus").innerText =
            "تم تحميل الملف ✅";

          e.target.value = "";

        } catch (err) {

          console.error(err);

          document.getElementById("liveStatus").innerText =
            "خطأ في قراءة الملف ❌";

          e.target.value = "";
        }
      };

      reader.readAsBinaryString(file);
    });

  // =========================
  // 🔍 SEARCH
  // =========================
  window.searchNumber = function () {

    // 🔥 INPUT NAMES
    const names =
      document.getElementById("nameBox")
        .value
        .toLowerCase()
        .split("\n")
        .map(v => v.trim())
        .filter(Boolean);

    // 🔥 INPUT NUMBERS
    const numbers =
      document.getElementById("numberBox")
        .value
        .split("\n")
        .map(v => getLast6(v))
        .filter(Boolean);

    let results = [];

    // 🔥 LOOP DATA
    for (let i = 1; i < DATA.length; i++) {

      const row = DATA[i];

      // 🔥 COLUMN B = NAME
      const excelName =
        String(row[1] || "")
          .toLowerCase()
          .trim();

      // 🔥 COLUMN C = ACCOUNT NUMBER
      const excelNumber =
        getLast6(row[2]);

      // 🔥 PARTIAL NAME SEARCH
      const nameMatch =
        names.length === 0 ||
        names.some(n => excelName.includes(n));

      // 🔥 NUMBER MATCH
      const numberMatch =
        numbers.length === 0 ||
        numbers.includes(excelNumber);

      // 🔥 PUSH RESULT
      if (nameMatch && numberMatch) {
        results.push(row);
      }
    }

    document.getElementById("liveStatus").innerText =
      `${results.length} نتيجة`;

    renderResults(results);
  };

  // =========================
  // 📊 RENDER RESULTS
  // =========================
  function renderResults(data) {

    const box =
      document.getElementById("resultsTable");

    // 🔥 NO RESULTS
    if (!data.length) {

      box.innerHTML = `
        <div style="
          padding:12px;
          color:#b00020;
          font-weight:bold;
        ">
          ❌ لا توجد نتائج
        </div>
      `;

      return;
    }

    let html = `
      <table style="
        width:100%;
        border-collapse:collapse;
        font-size:12px;
      ">
        <thead>
          <tr style="
            background:#f3f6fb;
          ">
            <th>#</th>
            <th>الاسم</th>
            <th>رقم الحساب الكامل</th>
          </tr>
        </thead>
        <tbody>
    `;

    // 🔥 TABLE ROWS
    for (let i = 0; i < data.length; i++) {

      html += `
        <tr>
          <td>${i + 1}</td>
          <td>${String(data[i][1] || "")}</td>
          <td>${String(data[i][2] || "")}</td>
        </tr>
      `;
    }

    html += `
        </tbody>
      </table>
    `;

    box.innerHTML = html;
  }

  // =========================
  // 🧹 CLEAR
  // =========================
  window.clearBox = function () {

    document.getElementById("nameBox").value = "";

    document.getElementById("numberBox").value = "";

    document.getElementById("resultsTable").innerHTML =
      `
      <div style="
        padding:10px;
        color:#888;
      ">
        لا توجد نتائج بعد
      </div>
      `;

    document.getElementById("inputCount").innerText =
      "0";

    document.getElementById("liveStatus").innerText =
      "جاهز للبحث 🚀";
  };

  // =========================
  // 👁️ SHOW ALL
  // =========================
  window.showAllRows = async function () {

    await Excel.run(async (context) => {

      const sheet =
        context.workbook
          .worksheets
          .getActiveWorksheet();

      const used =
        sheet.getUsedRange();

      used.load();

      await context.sync();

      used.rowHidden = false;

      document.getElementById("liveStatus").innerText =
        "تم إظهار كل الصفوف 👁️";
    });
  };

});