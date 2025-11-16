document.addEventListener("DOMContentLoaded", () => {
  const fileInputs = [
    document.getElementById("file1"),
    document.getElementById("file2"),
    document.getElementById("file3"),
  ];

  const submitBtn = document.getElementById("submit-btn");
  const statusText = document.getElementById("status-text");
  const form = document.getElementById("upload-form");
  const uploadCard = document.getElementById("upload-card");
  const workspace = document.getElementById("workspace");

  const filterCityEl = document.getElementById("filter-city");
  const filterRoleEl = document.getElementById("filter-role");
  const filterDateEl = document.getElementById("filter-date");
  const filterTrainingShopEl = document.getElementById("filter-training-shop");
  const report1Body = document.getElementById("report1-body");

  const report2Body = document.getElementById("report2-body");

  const filter3CityEl = document.getElementById("filter3-city");
  const filter3RoleEl = document.getElementById("filter3-role");
  const report3Body = document.getElementById("report3-body");

  const refreshButtons = document.querySelectorAll(".js-refresh");
  const refreshFile1 = document.getElementById("refresh-file1");
  const refreshFile2 = document.getElementById("refresh-file2");
  const refreshFile3 = document.getElementById("refresh-file3");

  // 8 слотов используются в 1-й выгрузке
  const HOUR_LABELS = [
    "7–8",
    "9–10",
    "11–12",
    "13–14",
    "15–16",
    "17–18",
    "19–20",
    "21–22"
  ];

  const MAIN_CITIES = [
    "Волгоград",
    "Воронеж",
    "Екатеринбург",
    "Казань",
    "Краснодар",
    "Красноярск",
    "Нижний Новгород",
    "Новосибирск",
    "Омск",
    "Пермь",
    "Самара",
    "Санкт-Петербург",
    "Тюмень",
    "Уфа",
    "Челябинск"
  ];

  const RAW_ROLE_BY_DISPLAY = {
    "ЭПЗ": "shopper",
    "ВК": "driver",
    "БД": "universal"
  };

  const RAW_ROLE3_BY_DISPLAY = {
    "ЭПЗ": "эпз",
    "ВК": "вк",
    "БД": "бд"
  };

  let export1Data = [];
  let export1ColumnMap = null;

  let export2Data = [];
  let export3Data = [];
  let export3ColumnMap = null;

  function normalizeDate(value) {
    if (!value) return null;

    if (value instanceof Date) {
      const y = value.getFullYear();
      const m = String(value.getMonth() + 1).padStart(2, "0");
      const d = String(value.getDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }

    const str = String(value).trim();

    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
      return str;
    }

    const m = str.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);
    if (m) {
      const d = m[1].padStart(2, "0");
      const mo = m[2].padStart(2, "0");
      const y = m[3];
      return `${y}-${mo}-${d}`;
    }

    return null;
  }

  function detectHourColumnMap(headers) {
    const map = {};

    headers.forEach((h) => {
      if (!h || typeof h !== "string") return;
      const trimmed = h.trim();
      if (!/^sh/i.test(trimmed)) return;

      const m = trimmed.match(/(\d{1,2})\D+(\d{1,2})/);
      if (!m) return;

      const start = parseInt(m[1], 10);
      const end = parseInt(m[2], 10);
      if (isNaN(start) || isNaN(end)) return;

      const key = `${start}–${end}`;
      if (HOUR_LABELS.includes(key)) {
        map[key] = h;
      }
    });

    return map;
  }

  function buildColumnMap(headers) {
    const map = {
      shop: null,
      trainingShop: null,
      day: null,
      trainingDateDay: null,
      approved: null,
      city: null,
      role: null,
      hourColumns: {}
    };

    headers.forEach((h) => {
      const key = String(h).trim().toLowerCase();

      if (key === "магазин") {
        map.shop = h;
      } else if (key === "магазин обучения") {
        map.trainingShop = h;
      } else if (key === "день") {
        map.day = h;
      } else if (key === "дата обучения: день") {
        map.trainingDateDay = h;
      } else if (key === "оформлено") {
        map.approved = h;
      } else if (key === "город") {
        map.city = h;
      } else if (key === "роль") {
        map.role = h;
      }
    });

    map.hourColumns = detectHourColumnMap(headers);
    return map;
  }

  // 3-я выгрузка: все колонки, кроме Город / Роль / Grand Total, идут как слоты
  function buildExport3ColumnMap(headers) {
    const map = {
      city: null,
      role: null,
      grandTotal: null,
      slotColumns: []
    };

    headers.forEach((h) => {
      if (!h) return;
      const key = String(h).trim().toLowerCase();

      if (key === "город") {
        map.city = h;
      } else if (key === "роль") {
        map.role = h;
      } else if (key === "grand total") {
        map.grandTotal = h;
      } else {
        map.slotColumns.push(h);
      }
    });

    return map;
  }

  // Обновление заголовков 3-й таблицы по slotColumns
  function updateExport3Headers() {
    if (!export3ColumnMap) return;
    const slotColumns = export3ColumnMap.slotColumns || [];
    const headerCells = document.querySelectorAll(
        '#report3-table thead th[data-slot-index]'
    );

    headerCells.forEach((th) => {
      const idx = parseInt(th.getAttribute("data-slot-index"), 10);
      const colName = slotColumns[idx];
      th.textContent = colName || "";
    });
  }

  function rebuildTrainingShopFilterOptions() {
    if (!filterTrainingShopEl || !export1Data.length || !export1ColumnMap) return;

    const trainingShopCol = export1ColumnMap.trainingShop;
    if (!trainingShopCol) return;

    const valuesSet = new Set();

    export1Data.forEach((row) => {
      const val = row[trainingShopCol];
      if (val === null || val === undefined) return;
      const str = String(val).trim();
      if (!str) return;
      valuesSet.add(str);
    });

    const values = Array.from(valuesSet);
    values.sort((a, b) => a.localeCompare(b, "ru"));

    filterTrainingShopEl.innerHTML = "";

    const defaultOption = document.createElement("option");
    defaultOption.textContent = "Все магазины обучения";
    defaultOption.value = "Все магазины обучения";
    filterTrainingShopEl.appendChild(defaultOption);

    values.forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      filterTrainingShopEl.appendChild(opt);
    });
  }

  function renderExport1Table() {
    if (!report1Body) return;
    report1Body.innerHTML = "";

    if (!export1Data.length || !export1ColumnMap) {
      return;
    }

    const cityFilter = filterCityEl ? filterCityEl.value : "";
    const roleFilter = filterRoleEl ? filterRoleEl.value : "Все роли";
    const dateFilter = filterDateEl ? filterDateEl.value : "";
    const trainingShopFilter = filterTrainingShopEl
        ? filterTrainingShopEl.value
        : "Все магазины обучения";

    const cityCol = export1ColumnMap.city;
    const roleCol = export1ColumnMap.role;
    const dateCol = export1ColumnMap.trainingDateDay;

    const shopCol = export1ColumnMap.shop;
    const trainingShopCol = export1ColumnMap.trainingShop;
    const dayCol = export1ColumnMap.day || export1ColumnMap.trainingDateDay;
    const approvedCol = export1ColumnMap.approved;
    const hourColumns = export1ColumnMap.hourColumns || {};

    const filtered = export1Data.filter((row) => {
      if (cityCol && cityFilter) {
        const cityVal = String(row[cityCol] ?? "").trim();

        if (cityFilter === "Остальные города") {
          if (!cityVal) return false;
          if (MAIN_CITIES.includes(cityVal)) return false;
        } else {
          if (cityVal !== cityFilter) return false;
        }
      }

      if (roleFilter && roleFilter !== "Все роли" && roleCol) {
        const rawNeeded = RAW_ROLE_BY_DISPLAY[roleFilter];
        const rawVal = String(row[roleCol] || "").trim().toLowerCase();
        if (rawNeeded && rawVal !== rawNeeded) return false;
      }

      if (dateFilter && dateCol) {
        const rawDateVal = row[dateCol];
        let match = false;

        const normalizedRowDate = normalizeDate(rawDateVal);
        if (normalizedRowDate) {
          match = normalizedRowDate === dateFilter;
        } else {
          const selectedDay = parseInt(dateFilter.slice(8, 10), 10);
          const rawStr = String(rawDateVal ?? "").trim();
          const asInt = parseInt(rawStr, 10);

          if (!isNaN(asInt) && !isNaN(selectedDay)) {
            match = asInt === selectedDay;
          }
        }

        if (!match) return false;
      }

      if (
          trainingShopFilter &&
          trainingShopFilter !== "Все магазины обучения" &&
          trainingShopCol
      ) {
        const val = String(row[trainingShopCol] ?? "").trim();
        if (val !== trainingShopFilter) return false;
      }

      return true;
    });

    const rowsToRender = filtered.slice(0, 1000);

    rowsToRender.forEach((row) => {
      const tr = document.createElement("tr");

      const shopCell = document.createElement("th");
      shopCell.className = "report-table__row-header";
      shopCell.textContent = shopCol ? (row[shopCol] ?? "") : "";
      tr.appendChild(shopCell);

      const trainCell = document.createElement("td");
      trainCell.textContent = trainingShopCol ? (row[trainingShopCol] ?? "") : "";
      tr.appendChild(trainCell);

      const dayCell = document.createElement("td");
      let dayValue = dayCol ? row[dayCol] : "";

      const isEmptyDay =
          dayValue === null ||
          dayValue === undefined ||
          String(dayValue).trim() === "";

      if (isEmptyDay && dateCol) {
        const rawDate = row[dateCol];
        if (rawDate instanceof Date) {
          dayValue = rawDate.toLocaleDateString("ru-RU");
        } else if (rawDate) {
          dayValue = String(rawDate);
        }
      }

      dayCell.textContent =
          dayValue === null || dayValue === undefined ? "" : String(dayValue);
      tr.appendChild(dayCell);

      const approvedCell = document.createElement("td");
      approvedCell.textContent = approvedCol ? (row[approvedCol] ?? "") : "";
      tr.appendChild(approvedCell);

      HOUR_LABELS.forEach((label) => {
        const td = document.createElement("td");
        const colName = hourColumns[label];
        const value = colName ? row[colName] : "";
        td.textContent =
            value === null || value === undefined ? "" : String(value);
        tr.appendChild(td);
      });

      report1Body.appendChild(tr);
    });
  }

  function normalizeExport2Row(row) {
    const keys = Object.keys(row);
    if (keys.length < 7) {
      return null;
    }

    const [
      kCity,
      kEpz,
      kEpzOut,
      kVk,
      kVkOut,
      kBd,
      kBdOut
    ] = keys;

    return {
      city: row[kCity] ?? "",
      epz: row[kEpz] ?? "",
      epzOut: row[kEpzOut] ?? "",
      vk: row[kVk] ?? "",
      vkOut: row[kVkOut] ?? "",
      bd: row[kBd] ?? "",
      bdOut: row[kBdOut] ?? ""
    };
  }

  function renderExport2Table() {
    if (!report2Body) return;
    report2Body.innerHTML = "";

    if (!export2Data.length) {
      return;
    }

    export2Data.forEach((row) => {
      const parsed = normalizeExport2Row(row);
      if (!parsed) return;

      const tr = document.createElement("tr");

      const tdCity = document.createElement("td");
      tdCity.textContent = parsed.city || "";
      tr.appendChild(tdCity);

      const tdEpz = document.createElement("td");
      tdEpz.textContent = parsed.epz || "";
      tr.appendChild(tdEpz);

      const tdEpzOut = document.createElement("td");
      tdEpzOut.textContent = parsed.epzOut || "";
      tr.appendChild(tdEpzOut);

      const tdVk = document.createElement("td");
      tdVk.textContent = parsed.vk || "";
      tr.appendChild(tdVk);

      const tdVkOut = document.createElement("td");
      tdVkOut.textContent = parsed.vkOut || "";
      tr.appendChild(tdVkOut);

      const tdBd = document.createElement("td");
      tdBd.textContent = parsed.bd || "";
      tr.appendChild(tdBd);

      const tdBdOut = document.createElement("td");
      tdBdOut.textContent = parsed.bdOut || "";
      tr.appendChild(tdBdOut);

      report2Body.appendChild(tr);
    });
  }

  // 3-я выгрузка
  function renderExport3Table() {
    if (!report3Body) return;
    report3Body.innerHTML = "";

    if (!export3Data.length || !export3ColumnMap) {
      return;
    }

    const cityFilter = filter3CityEl ? filter3CityEl.value : "";
    const roleFilter = filter3RoleEl ? filter3RoleEl.value : "Все роли";

    const cityCol = export3ColumnMap.city;
    const roleCol = export3ColumnMap.role;
    const grandTotalCol = export3ColumnMap.grandTotal;
    const slotColumns = export3ColumnMap.slotColumns || [];

    const filtered = export3Data.filter((row) => {
      if (cityCol && cityFilter) {
        const cityVal = String(row[cityCol] ?? "").trim();

        if (cityFilter === "Остальные города") {
          if (!cityVal) return false;
          if (MAIN_CITIES.includes(cityVal)) return false;
        } else {
          if (cityVal !== cityFilter) return false;
        }
      }

      if (roleFilter && roleFilter !== "Все роли" && roleCol) {
        const needed = RAW_ROLE3_BY_DISPLAY[roleFilter];
        const rawVal = String(row[roleCol] || "").trim().toLowerCase();
        if (needed && rawVal !== needed) return false;
      }

      return true;
    });

    const rowsToRender = filtered.slice(0, 1000);

    const headerCells = document.querySelectorAll(
        '#report3-table thead th[data-slot-index]'
    );

    rowsToRender.forEach((row) => {
      const tr = document.createElement("tr");

      const tdCity = document.createElement("td");
      tdCity.textContent = cityCol ? (row[cityCol] ?? "") : "";
      tr.appendChild(tdCity);

      const tdRole = document.createElement("td");
      tdRole.textContent = roleCol ? (row[roleCol] ?? "") : "";
      tr.appendChild(tdRole);

      const tdGT = document.createElement("td");
      tdGT.className = "report-table__td--grand-total";
      tdGT.textContent = grandTotalCol ? (row[grandTotalCol] ?? "") : "";
      tr.appendChild(tdGT);

      // Проходим по фактическим заголовкам-слотам (7 штук),
      // чтобы количество колонок совпадало с шапкой.
      headerCells.forEach((th) => {
        const idx = parseInt(th.getAttribute("data-slot-index"), 10);
        const colName = slotColumns[idx];
        const value = colName ? row[colName] : "";
        const td = document.createElement("td");
        td.textContent =
            value === null || value === undefined ? "" : String(value);
        tr.appendChild(td);
      });

      report3Body.appendChild(tr);
    });
  }

  function checkFiles() {
    const allSelected = fileInputs.every(input => input.files.length > 0);
    submitBtn.disabled = !allSelected;

    statusText.textContent = allSelected
        ? "Все файлы выбраны. Нажмите «Продолжить», чтобы перейти к визуальному отчёту."
        : "Загрузите все три файла, чтобы перейти к просмотру данных.";
  }

  fileInputs.forEach(input => {
    input.addEventListener("change", checkFiles);
  });

  function loadExport1(file) {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const json = XLSX.utils.sheet_to_json(worksheet, {
          defval: "",
          raw: false
        });

        export1Data = json;

        const headers = json.length ? Object.keys(json[0]) : [];
        export1ColumnMap = buildColumnMap(headers);

        rebuildTrainingShopFilterOptions();

        uploadCard.classList.add("card--hidden");
        workspace.classList.remove("workspace--hidden");

        renderExport1Table();
        renderExport2Table();
        renderExport3Table();

        workspace.scrollIntoView({ behavior: "smooth", block: "start" });
      } catch (err) {
        console.error("Ошибка при чтении файла 1 выгрузки:", err);
        alert("Не удалось прочитать файл «Выгрузка 1». Проверьте формат и названия столбцов.");
      } finally {
        if (refreshFile1) refreshFile1.value = "";
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function loadExport2(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        export2Data = XLSX.utils.sheet_to_json(worksheet, {
          defval: "",
          raw: false
        });
        console.log("Выгрузка 2 обновлена, строк:", export2Data.length);

        renderExport2Table();
      } catch (err) {
        console.error("Ошибка при чтении файла 2 выгрузки:", err);
        alert("Не удалось прочитать файл «Выгрузка 2».");
      } finally {
        if (refreshFile2) refreshFile2.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadExport3(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        export3Data = XLSX.utils.sheet_to_json(worksheet, {
          defval: "",
          raw: false
        });

        const headers = export3Data.length ? Object.keys(export3Data[0]) : [];
        export3ColumnMap = buildExport3ColumnMap(headers);

        console.log("Выгрузка 3 обновлена, строк:", export3Data.length);

        updateExport3Headers();
        renderExport3Table();
      } catch (err) {
        console.error("Ошибка при чтении файла 3 выгрузки:", err);
        alert("Не удалось прочитать файл «Выгрузка 3».");
      } finally {
        if (refreshFile3) refreshFile3.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  }

  if (filterCityEl) filterCityEl.addEventListener("change", renderExport1Table);
  if (filterRoleEl) filterRoleEl.addEventListener("change", renderExport1Table);
  if (filterDateEl) filterDateEl.addEventListener("change", renderExport1Table);
  if (filterTrainingShopEl) filterTrainingShopEl.addEventListener("change", renderExport1Table);

  if (filter3CityEl) filter3CityEl.addEventListener("change", renderExport3Table);
  if (filter3RoleEl) filter3RoleEl.addEventListener("change", renderExport3Table);

  refreshButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const exp = btn.dataset.export;
      if (exp === "1" && refreshFile1) {
        refreshFile1.click();
      } else if (exp === "2" && refreshFile2) {
        refreshFile2.click();
      } else if (exp === "3" && refreshFile3) {
        refreshFile3.click();
      }
    });
  });

  if (refreshFile1) {
    refreshFile1.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (file) loadExport1(file);
    });
  }

  if (refreshFile2) {
    refreshFile2.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (file) loadExport2(file);
    });
  }

  if (refreshFile3) {
    refreshFile3.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (file) loadExport3(file);
    });
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();

    const file1 = document.getElementById("file1").files[0];
    const file2 = document.getElementById("file2").files[0];
    const file3 = document.getElementById("file3").files[0];

    if (!file1 || !file2 || !file3) {
      alert("Пожалуйста, загрузите все три файла (Выгрузка 1, 2 и 3).");
      return;
    }

    loadExport1(file1);
    loadExport2(file2);
    loadExport3(file3);
  });
});
