Office.onReady(async () => {

  function getLast6(value) {
    return (value || "")
      .toString()
      .replace(/\D/g, "")
      .slice(-6);
  }

  let DATA = [];

  async function loadData() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();

      range.load("values");
      await context.sync();

      DATA = range.values || [];
    });

    setStatus("جاهز للبحث 🚀");
  }

  await loadData();

  // =========================
  // STATUS
  // =========================
  function setStatus(text) {
    document.getElementById("liveStatus").innerText = text;
  }

  // =========================
  // REPORT BUTTON
  // =========================
  function showReportBtn() {
    document.getElementById("reportBtn").style.display = "inline-block";
  }

  function hideReportBtn() {
    document.getElementById("reportBtn").style.display = "none";
  }

  document.getElementById("reportBtn").addEventListener("click", function () {

    const rows = [];

    const names = document.getElementById("nameBox").value
      .split("\n").map(v => v.trim()).filter(Boolean);

    const numbers = document.getElementById("numberBox").value
      .split("\n").map(v => v.trim()).filter(Boolean);

    names.forEach(n => rows.push({ type: "اسم", value: n }));
    numbers.forEach(n => rows.push({ type: "رقم", value: n }));

    if (!rows.length) return;

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Report");

    XLSX.writeFile(wb, "report.xlsx");

    setStatus("تم إنشاء التقرير 📄");
  });

  // =========================
  // SEARCH
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
      const last6 = getLast6(row[2]);

      const nameMatch =
        nameSet.size === 0 || nameSet.has(name);

      const numberMatch =
        numberSet.size === 0 || numberSet.has(last6);

      if (nameMatch && numberMatch) {
        results.push(row);
      }
    }

    const hasInput = nameSet.size > 0 || numberSet.size > 0;

    if (!hasInput) {
      setStatus("جاهز للبحث 🚀");
      hideReportBtn();
      renderResults(results);
      return;
    }

    // ❌ NO RESULTS
    if (results.length === 0) {
      setStatus("غير موجود في الشيت ❌");
      showReportBtn();
      renderResults([]);
      return;
    }

    // ✔ RESULTS
    setStatus(`${results.length} نتيجة`);
    hideReportBtn();

    const allNamesMatched = [...nameSet].every(n =>
      results.some(r => (r[1] || "").toLowerCase().trim() === n)
    );

    const allNumbersMatched = [...numberSet].every(n =>
      results.some(r => getLast6(r[2]) === n)
    );

    if (allNamesMatched && allNumbersMatched) {
      setStatus(`${results.length} نتيجة - كل البيانات صحيحة ✅`);
    }

    renderResults(results);
  };

  // =========================
  // TABLE
  // =========================
  function renderResults(data) {

    const box = document.getElementById("resultsTable");

    if (!data.length) {
      box.innerHTML = `<div style="padding:10px;color:#b00020;">❌ لا توجد نتائج</div>`;
      return;
    }

    let html = `
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead>
          <tr style="background:#f3f6fb;">
            <th>#</th>
            <th>الاسم</th>
            <th>الرقم</th>
          </tr>
        </thead>
        <tbody>
    `;

    data.forEach((r, i) => {
      html += `
        <tr>
          <td>${i + 1}</td>
          <td>${r[1] || ""}</td>
          <td>${r[2] || ""}</td>
        </tr>
      `;
    });

    html += `</tbody></table>`;
    box.innerHTML = html;
  }

  // =========================
  // CLEAR
  // =========================
  window.clearBox = function () {

    document.getElementById("nameBox").value = "";
    document.getElementById("numberBox").value = "";

    document.getElementById("inputCount").innerText = "0";

    setStatus("جاهز للبحث 🚀");
    hideReportBtn();

    document.getElementById("resultsTable").innerHTML =
      `<div style="padding:10px;color:#888;">لا توجد نتائج بعد</div>`;
  };

  // =========================
  // SHOW ALL
  // =========================
  window.showAllRows = async function () {

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();

      used.load();
      await context.sync();

      used.rowHidden = false;
    });

    setStatus("جاهز للبحث 🚀");
  };

});