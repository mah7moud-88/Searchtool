Office.onReady(() => {

  function getLast6(value) {
    const v = String(value || "").replace(/\D/g, "");
    return v.slice(-6);
  }

  window.updateCounter = function () {

    const names =
      document.getElementById("nameBox").value
        .split("\n")
        .filter(v => v.trim());

    const numbers =
      document.getElementById("numberBox").value
        .split("\n")
        .filter(v => v.trim());

    document.getElementById("inputCount").innerText =
      names.length + numbers.length;
  };

  document.getElementById("nameBox")
    .addEventListener("input", updateCounter);

  document.getElementById("numberBox")
    .addEventListener("input", updateCounter);


  // 📁 رفع ملف Excel
  document.getElementById("fileInput")
    .addEventListener("change", function (e) {

      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();

      reader.onload = function (event) {

        const workbook = XLSX.read(event.target.result, { type: "binary" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(1);

        document.getElementById("nameBox").value =
          rows.map(r => r[0] || "").join("\n");

        document.getElementById("numberBox").value =
          rows.map(r => r[1] || "").join("\n");

        setTimeout(updateCounter, 0);
      };

      reader.readAsBinaryString(file);
    });


  // 🔍 البحث
  window.searchNumber = async function () {

    await Excel.run(async (context) => {

      const nameSet = new Set(
        document.getElementById("nameBox")
          .value
          .toLowerCase()
          .split("\n")
          .map(v => v.trim())
          .filter(Boolean)
      );

      const rawNumbers =
        document.getElementById("numberBox")
          .value
          .split("\n")
          .map(v => v.trim())
          .filter(Boolean);

      const numberSet =
        new Set(rawNumbers.map(v => getLast6(v)));

      const sheet =
        context.workbook.worksheets.getActiveWorksheet();


      // ====================================================
      // ✅ الحل القوي: قراءة B و C مباشرة
      // ====================================================

      const used = sheet.getUsedRange();
      used.load("rowCount");
      await context.sync();

      const rowCount = used.rowCount;

      const range = sheet.getRangeByIndexes(0, 1, rowCount, 2);
      range.load("values");

      await context.sync();

      const values = range.values;


      let showRows = [];
      let foundNames = new Set();
      let foundNumbers = new Set();

      let notFound = [];
      let mismatch = [];

      const sheetNameSet = new Set();
      const sheetNumberSet = new Set();
      const nameToNumbers = new Map();


      // 📥 قراءة البيانات (B = اسم / C = رقم)
      for (let i = 0; i < values.length; i++) {

        const row = values[i];

        const name = String(row[0] || "").toLowerCase().trim();
        const number = getLast6(row[1]);

        if (!name && !number) continue;

        sheetNameSet.add(name);
        sheetNumberSet.add(number);

        if (!nameToNumbers.has(name)) {
          nameToNumbers.set(name, new Set());
        }

        nameToNumbers.get(name).add(number);
      }


      // 🔍 البحث
      for (let i = 0; i < values.length; i++) {

        const row = values[i];

        const name = String(row[0] || "").toLowerCase().trim();
        const number = getLast6(row[1]);

        let found = false;

        if (nameSet.size && numberSet.size) {
          found = nameSet.has(name) && numberSet.has(number);
        }
        else if (nameSet.size) {
          found = nameSet.has(name);
        }
        else {
          found = numberSet.has(number);
        }

        if (found) {
          showRows.push(i);

          if (nameSet.has(name)) foundNames.add(name);
          if (numberSet.has(number)) foundNumbers.add(number);
        }
      }


      // ❌ غير موجود
      nameSet.forEach(name => {
        if (!sheetNameSet.has(name)) {
          notFound.push(`الاسم غير موجود: ${name}`);
        }
      });

      rawNumbers.forEach(fullNum => {
        const last6 = getLast6(fullNum);
        if (!sheetNumberSet.has(last6)) {
          notFound.push(`الرقم غير موجود: ${fullNum}`);
        }
      });


      // ⚠️ غير مطابق
      const mismatchSet = new Set();

      rawNumbers.forEach(fullNum => {

        const last6 = getLast6(fullNum);

        if (!sheetNumberSet.has(last6)) return;

        let matched = false;

        for (let inputName of nameSet) {

          if (!nameToNumbers.has(inputName)) continue;

          if (nameToNumbers.get(inputName).has(last6)) {
            matched = true;
            break;
          }
        }

        if (!matched) {
          mismatchSet.add(`❌ الرقم غير مطابق للاسم\n🔢 الرقم: ${fullNum}`);
        }
      });

      mismatch = Array.from(mismatchSet);


      // 📊 الحالة
      const totalInputs = nameSet.size + numberSet.size;
      const totalFound = foundNames.size + foundNumbers.size;

      document.getElementById("liveStatus").innerText =
        `${totalFound} من ${totalInputs} مطابق`;


      // ====================================================
      // 👁️ إظهار/إخفاء الصفوف (حل ثابت)
      // ====================================================

      const totalRows = values.length;

      // إظهار الكل أولاً
      sheet.getRangeByIndexes(0, 0, totalRows, 3).rowHidden = false;

      // إخفاء غير المطابق
      for (let i = 0; i < totalRows; i++) {

        if (!showRows.includes(i)) {

          sheet.getRangeByIndexes(i, 0, 1, 3).rowHidden = true;
        }
      }


      // 📦 التقرير
      const missingBox = document.getElementById("missingBox");

      missingBox.style.display = "block";

      missingBox.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;font-weight:bold;margin-bottom:8px;">
          <span>📌 تقرير التحقق</span>
          <span id="copyAllBtn" style="cursor:pointer;font-size:18px;">📋</span>
        </div>

        ${
          mismatch.length
            ? `<div style="color:#b97700;font-weight:bold;">⚠️ غير مطابق:</div>
               ${mismatch.map(v => `<div>• ${v}</div>`).join("")}`
            : ""
        }

        ${
          notFound.length
            ? `<div style="color:#b00000;font-weight:bold;margin-top:10px;">❌ غير موجود:</div>
               ${notFound.map(v => `<div>• ${v}</div>`).join("")}`
            : ""
        }

        ${
          !notFound.length && !mismatch.length
            ? `<div style="color:#1a7f37;font-weight:bold;">✅ كل البيانات صحيحة</div>`
            : ""
        }
      `;


      document.getElementById("copyAllBtn").onclick = function () {

        const text = [
          ...mismatch.map(v => "⚠️ " + v),
          ...notFound.map(v => "❌ " + v)
        ].join("\n");

        if (!text) return;

        navigator.clipboard.writeText(text);
        this.innerText = "✔️";

        setTimeout(() => {
          this.innerText = "📋";
        }, 1200);
      };

      await context.sync();
    });
  };


  // 👁️ إظهار كل الصفوف
  window.showAllRows = async function () {

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const used = sheet.getUsedRange();
      used.load();
      await context.sync();

      used.rowHidden = false;

      document.getElementById("liveStatus").innerText = "جاهز للبحث";
    });
  };


  // 🧹 مسح
  window.clearBox = function () {

    document.getElementById("nameBox").value = "";
    document.getElementById("numberBox").value = "";
    document.getElementById("missingBox").style.display = "none";
    document.getElementById("liveStatus").innerText = "جاهز للبحث";

    updateCounter();
  };

});