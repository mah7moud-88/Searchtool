Office.onReady(async () => {

  // =========================
  // 👤 CLEAN TEXT
  // =========================
  function normalizeText(value) {

    return String(value || "")
      .toLowerCase()

      // حذف المسافات
      .replace(/\s+/g, "")

      // حذف التشكيل
      .replace(/[\u064B-\u065F]/g, "")

      // توحيد الحروف
      .replace(/[أإآ]/g, "ا")
      .replace(/ة/g, "ه")
      .replace(/ى/g, "ي")

      // حذف الرموز
      .replace(/[^\w\u0600-\u06FF]/g, "");
  }

  // =========================
  // 🔢 CLEAN NUMBER
  // =========================
  function normalizeNumber(value) {

    return String(value || "")

      // عربي → إنجليزي
      .replace(/[٠-٩]/g, d =>
        "٠١٢٣٤٥٦٧٨٩".indexOf(d)
      )

      // حذف أي شيء غير رقم
      .replace(/[^\d]/g, "");
  }

  // =========================
  // 📦 DATA STORAGE
  // =========================
  let DATA = [];

  // =========================
  // 📥 LOAD DATA
  // =========================
  async function loadData() {

    document.getElementById("liveStatus").innerText =
      "جاري تحميل البيانات...";

    await Excel.run(async (context) => {

      const sheet =
        context.workbook
          .worksheets
          .getActiveWorksheet();

      const range =
        sheet.getUsedRange();

      range.load("text");

      await context.sync();

      DATA = range.text || [];
    });

    document.getElementById("liveStatus").innerText =
      "جاهز للبحث 🚀";
  }

  await loadData();

  // =========================
  // 📊 UPDATE COUNTER
  // =========================
  function updateCounter() {

    const names =
      document.getElementById("nameBox")
        .value
        .split("\n")
        .filter(v => v.trim());

    const numbers =
      document.getElementById("numberBox")
        .value
        .split("\n")
        .filter(v => v.trim());

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

          DATA =
            XLSX.utils.sheet_to_json(sheet, {
              header: 1,
              raw: false,
              defval: ""
            });

          // 👤 الاسم = العمود A
          document.getElementById("nameBox").value =
            DATA.slice(1)
              .map(r => String(r[0] || ""))
              .join("\n");

          // 🔢 الرقم = العمود B
          document.getElementById("numberBox").value =
            DATA.slice(1)
              .map(r => String(r[1] || ""))
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

    // 👤 INPUT NAMES
    const names =
      document.getElementById("nameBox")
        .value
        .split("\n")
        .map(v => normalizeText(v))
        .filter(Boolean);

    // 🔢 INPUT NUMBERS
    const numbers =
      document.getElementById("numberBox")
        .value
        .split("\n")
        .map(v => normalizeNumber(v))
        .filter(Boolean);

    let matchedResults = [];
    let unmatchedResults = [];

    // =========================
    // 🔍 LOOP INPUTS
    // =========================
    const maxLength =
      Math.max(names.length, numbers.length);

    for (let x = 0; x < maxLength; x++) {

      const inputName =
        names[x] || "";

      const inputNumber =
        numbers[x] || "";

      let found = false;

      // =========================
      // 🔍 SEARCH IN DATA
      // =========================
      for (let i = 1; i < DATA.length; i++) {

        const row = DATA[i];

        // 👤 العمود A
        const excelName =
          normalizeText(row[0]);

        // 🔢 العمود B
        const excelNumber =
          normalizeNumber(row[1]);

        const nameMatch =
          !inputName ||
          excelName.includes(inputName);

        const numberMatch =
          !inputNumber ||
          excelNumber.includes(inputNumber);

        // ✅ MATCHED
        if (nameMatch && numberMatch) {

          matchedResults.push(row);

          found = true;

          break;
        }
      }

      // ❌ NOT MATCHED OR NOT FOUND
      if (!found) {

        let reason = "";

        // 🔍 فحص الاسم
        const nameExists = DATA.some(r => {

          const excelName =
            normalizeText(r[0]);

          return excelName.includes(inputName);
        });

        // 🔍 فحص الرقم
        const numberExists = DATA.some(r => {

          const excelNumber =
            normalizeNumber(r[1]);

          return excelNumber.includes(inputNumber);
        });

        // 🚫 غير موجود
        if (
          (inputName && !nameExists) ||
          (inputNumber && !numberExists)
        ) {

          reason = "غير موجود";
        }

        // ❌ غير مطابق
        else {

          reason = "غير مطابق";
        }

        unmatchedResults.push({

          name:
            inputName || "-",

          number:
            inputNumber || "-",

          reason
        });
      }
    }

    // =========================
    // 📊 STATUS
    // =========================
    const notFoundCount =
      unmatchedResults.filter(
        x => x.reason === "غير موجود"
      ).length;

    const mismatchCount =
      unmatchedResults.filter(
        x => x.reason === "غير مطابق"
      ).length;

    document.getElementById("liveStatus").innerHTML = `
      ✅ صحيح: ${matchedResults.length}
      &nbsp;&nbsp; | &nbsp;&nbsp;
      ❌ غير مطابق: ${mismatchCount}
      &nbsp;&nbsp; | &nbsp;&nbsp;
      🚫 غير موجود: ${notFoundCount}
    `;

    // =========================
    // ✅ SUCCESS MESSAGE
    // =========================
    const successMsg =
      document.getElementById("successMsg");

    successMsg.style.display = "block";

    successMsg.style.padding = "8px";
    successMsg.style.borderRadius = "8px";

    if (
      mismatchCount === 0 &&
      notFoundCount === 0
    ) {

      successMsg.style.color =
        "#1a7f37";

      successMsg.style.background =
        "#eefbf1";

      successMsg.style.border =
        "1px solid #c9ebd1";

      successMsg.innerText =
        "كل القيم صحيحة ✅";
    }

    else {

      successMsg.style.color =
        "#b00020";

      successMsg.style.background =
        "#fff1f1";

      successMsg.style.border =
        "1px solid #ffd0d0";

      successMsg.innerText =
        `يوجد ${unmatchedResults.length} قيم غير صحيحة`;
    }

    // =========================
    // 📊 RENDER RESULTS
    // =========================
    renderAllResults(
      matchedResults,
      unmatchedResults
    );
  };

  // =========================
  // 📊 RENDER RESULTS
  // =========================
  function renderAllResults(
    matchedData,
    unmatchedData
  ) {

    const box =
      document.getElementById("resultsTable");

    let html = "";

    // =========================
    // ✅ MATCHED TABLE
    // =========================
    if (matchedData.length) {

      html += `

        <div style="
          margin-bottom:20px;
          border:1px solid #d6f0dc;
          border-radius:12px;
          overflow:hidden;
        ">

          <div style="
            background:#1a7f37;
            color:white;
            padding:10px;
            font-weight:bold;
            text-align:center;
          ">
            ✅ القيم الصحيحة
          </div>

          <table style="
            width:100%;
            border-collapse:collapse;
            font-size:12px;
          ">

            <thead>
              <tr style="background:#f4fff6;">
                <th style="padding:10px;">#</th>
                <th style="padding:10px;">الاسم</th>
                <th style="padding:10px;">الرقم</th>
              </tr>
            </thead>

            <tbody>
      `;

      for (let i = 0; i < matchedData.length; i++) {

        html += `
          <tr style="border-top:1px solid #eee;">

            <td style="padding:10px;">
              ${i + 1}
            </td>

            <td style="
              padding:10px;
              text-align:right;
              direction:rtl;
              unicode-bidi:plaintext;
              font-family:'Segoe UI',sans-serif;
            ">
              ${matchedData[i][0]}
            </td>

            <td style="
              padding:10px;
              direction:ltr;
            ">
              ${matchedData[i][1]}
            </td>

          </tr>
        `;
      }

      html += `
            </tbody>
          </table>
        </div>
      `;
    }

    // =========================
    // ❌ UNMATCHED TABLE
    // =========================
    if (unmatchedData.length) {

      html += `

        <div style="
          border:1px solid #ffd0d0;
          border-radius:12px;
          overflow:hidden;
        ">

          <div style="
            background:#b00020;
            color:white;
            padding:10px;
            font-weight:bold;
            text-align:center;
          ">
            ❌ القيم غير المطابقة / غير الموجودة
          </div>

          <table style="
            width:100%;
            border-collapse:collapse;
            font-size:12px;
          ">

            <thead>
              <tr style="background:#fff5f5;">
                <th style="padding:10px;">#</th>
                <th style="padding:10px;">الاسم</th>
                <th style="padding:10px;">الرقم</th>
                <th style="padding:10px;">الحالة</th>
              </tr>
            </thead>

            <tbody>
      `;

      for (let i = 0; i < unmatchedData.length; i++) {

        html += `
          <tr style="border-top:1px solid #eee;">

            <td style="padding:10px;">
              ${i + 1}
            </td>

            <td style="
              padding:10px;
              text-align:right;
              direction:rtl;
              unicode-bidi:plaintext;
              white-space:normal;
              word-break:break-word;
              font-family:'Segoe UI',sans-serif;
            ">
              ${unmatchedData[i].name || "-"}
            </td>

            <td style="
              padding:10px;
              direction:ltr;
            ">
              ${unmatchedData[i].number || "-"}
            </td>

            <td style="
              padding:10px;
              font-weight:bold;
              color:
                ${unmatchedData[i].reason === "غير موجود"
                  ? "#d97706"
                  : "#b00020"};
            ">
              ${unmatchedData[i].reason}
            </td>

          </tr>
        `;
      }

      html += `
            </tbody>
          </table>
        </div>
      `;
    }

    // =========================
    // 🚫 NOTHING
    // =========================
    if (
      !matchedData.length &&
      !unmatchedData.length
    ) {

      html = `
        <div style="
          padding:20px;
          text-align:center;
          color:#888;
        ">
          لا توجد نتائج
        </div>
      `;
    }

    box.innerHTML = html;
  }

  // =========================
  // 🧹 CLEAR
  // =========================
  window.clearBox = function () {

    document.getElementById("nameBox").value = "";
    document.getElementById("numberBox").value = "";

    document.getElementById("resultsTable").innerHTML = `
      <div style="
        padding:20px;
        text-align:center;
        color:#888;
      ">
        لا توجد نتائج بعد
      </div>
    `;

    document.getElementById("inputCount").innerText =
      "0";

    document.getElementById("liveStatus").innerText =
      "جاهز للبحث 🚀";

    document.getElementById("successMsg")
      .style.display = "none";
  };

  // =========================
  // 👁️ SHOW ALL
  // =========================
  window.showAllRows = function () {

    document.getElementById("liveStatus").innerText =
      "تم إظهار البيانات 👁️";
  };

});