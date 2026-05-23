Office.onReady(async () => {

  function getLast6(value) {
    return (value || "")
      .toString()
      .replace(/\D/g, "")
      .slice(-6);
  }

  // =========================
  // ⚡ CACHE DATA
  // =========================
  let DATA = [];

  async function loadData() {

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();

      range.load("values");
      await context.sync();

      DATA = range.values || [];
    });

    document.getElementById("liveStatus").innerText = "جاهز للبحث 🚀";
  }

  await loadData();

  // =========================
  // 📊 COUNTER (FIXED - NO STATUS CHANGE)
  // =========================
  function updateCounter() {

    const nameBox = document.getElementById("nameBox");
    const numberBox = document.getElementById("numberBox");

    const names =
      (nameBox.value || "")
        .split("\n")
        .map(v => v.trim())
        .filter(Boolean);

    const numbers =
      (numberBox.value || "")
        .split("\n")
        .map(v => v.trim())
        .filter(Boolean);

    document.getElementById("inputCount").innerText =
      names.length + numbers.length;
  }

  window.updateCounter = updateCounter;

  document.getElementById("nameBox").addEventListener("input", updateCounter);
  document.getElementById("numberBox").addEventListener("input", updateCounter);

  // =========================
  // 📁 FILE IMPORT (FIXED)
  // =========================
  document.getElementById("fileInput")
    .addEventListener("change", function (e) {

      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();

      reader.onload = function (event) {

        try {
          const workbook = XLSX.read(event.target.result, { type: "binary" });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(1);

          document.getElementById("nameBox").value =
            rows.map(r => r[0] || "").join("\n");

          document.getElementById("numberBox").value =
            rows.map(r => r[1] || "").join("\n");

          updateCounter();

          // 📊 رجع الحالة زي ما كانت
          document.getElementById("liveStatus").innerText =
            "جاهز للبحث 🚀";

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

    const nameSet = new Set(
      document.getElementById("nameBox")
        .value.toLowerCase()
        .split("\n")
        .map(v => v.trim())
        .filter(Boolean)
    );

    const numberSet = new Set(
      document.getElementById("numberBox")
        .value
        .split("\n")
        .map(v => getLast6(v))
        .filter(v => v.length === 6)
    );

    let results = [];

    for (let i = 0; i < DATA.length; i++) {

      const row = DATA[i];

      const name = (row[1] || "").toLowerCase().trim();
      const fullNumber = (row[2] || "").toString();
      const last6 = getLast6(fullNumber);

      const nameMatch =
        nameSet.size === 0 || nameSet.has(name);

      const numberMatch =
        numberSet.size === 0 || numberSet.has(last6);

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

    const box = document.getElementById("resultsTable");

    if (!data.length) {

      box.innerHTML = `
        <div style="padding:12px;font-weight:bold;color:#b00020;">
          ❌ لا توجد نتائج
        </div>
      `;
      return;
    }

    let html = `
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead>
          <tr style="background:#f3f6fb;">
            <th>#</th>
            <th>الاسم</th>
            <th>رقم الحساب الكامل</th>
          </tr>
        </thead>
        <tbody>
    `;

    for (let i = 0; i < data.length; i++) {

      html += `
        <tr>
          <td>${i + 1}</td>
          <td>${data[i][1] || ""}</td>
          <td>${data[i][2] || ""}</td>
        </tr>
      `;
    }

    html += `</tbody></table>`;

    box.innerHTML = html;
  }

  // =========================
  // 🧹 CLEAR
  // =========================
  window.clearBox = function () {

    document.getElementById("nameBox").value = "";
    document.getElementById("numberBox").value = "";
    document.getElementById("resultsTable").innerHTML =
      `<div style="padding:10px;color:#888;">لا توجد نتائج بعد</div>`;

    document.getElementById("inputCount").innerText = "0";
    document.getElementById("liveStatus").innerText = "جاهز للبحث 🚀";
  };

  // =========================
  // 👁️ SHOW ALL
  // =========================
  window.showAllRows = async function () {

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();

      used.load();
      await context.sync();

      used.rowHidden = false;

      document.getElementById("liveStatus").innerText = "جاهز للبحث 🚀";
    });
  };

});