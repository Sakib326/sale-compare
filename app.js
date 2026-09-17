// State
let excelWorkbook = null;
let excelFileName = "";
let docFile = null;
let docFileName = "";
let docXmlContent = null;
let comparisonRows = [];

// DOM Elements
const excelInput = document.getElementById("excelInput");
const excelFileNameSpan = document.getElementById("excelFileName");
const excelStatus = document.getElementById("excelStatus");
const excelDropZone = document.getElementById("excelDropZone");

const docInput = document.getElementById("docInput");
const docFileNameSpan = document.getElementById("docFileName");
const docStatus = document.getElementById("docStatus");
const docDropZone = document.getElementById("docDropZone");

const compareBtn = document.getElementById("compareBtn");
const downloadBtn = document.getElementById("downloadBtn");

const statsContainer = document.getElementById("statsContainer");
const matchedCountEl = document.getElementById("matchedCount");
const excelUnmatchedCountEl = document.getElementById("excelUnmatchedCount");
const docUnmatchedCountEl = document.getElementById("docUnmatchedCount");
const totalRowCountEl = document.getElementById("totalRowCount");

const filterBar = document.getElementById("filterBar");
const statusFilter = document.getElementById("statusFilter");

const previewSection = document.getElementById("previewSection");
const previewTableBody = document.getElementById("previewTableBody");
const previewTableFoot = document.getElementById("previewTableFoot");
const previewCountBadge = document.getElementById("previewCountBadge");

// Event Listeners
excelInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) handleExcelFile(file);
});

docInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) handleDocFile(file);
});

setupDragAndDrop(excelDropZone, handleExcelFile);
setupDragAndDrop(docDropZone, handleDocFile);

function setupDragAndDrop(zone, handler) {
  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    zone.classList.add("drag-over");
  });
  zone.addEventListener("dragleave", () => {
    zone.classList.remove("drag-over");
  });
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handler(e.dataTransfer.files[0]);
    }
  });
}

function handleExcelFile(file) {
  excelFileName = file.name;
  excelFileNameSpan.textContent = "Change Excel File";
  excelStatus.textContent = `✓ Selected: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  excelStatus.className = "file-status active";
  excelDropZone.classList.add("loaded");

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      excelWorkbook = XLSX.read(data, { type: "array" });
      checkReadyToCompare();
    } catch (err) {
      alert("Error reading Excel file: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleDocFile(file) {
  docFile = file;
  docFileName = file.name;
  docFileNameSpan.textContent = "Change Docx File";
  docStatus.textContent = `✓ Selected: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  docStatus.className = "file-status active";
  docDropZone.classList.add("loaded");

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const zip = await JSZip.loadAsync(e.target.result);
      const docXml = await zip.file("word/document.xml").async("text");
      docXmlContent = docXml;
      checkReadyToCompare();
    } catch (err) {
      alert("Error reading Docx XML: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function checkReadyToCompare() {
  if (excelWorkbook && docXmlContent) {
    compareBtn.removeAttribute("disabled");
  }
}

// Compare & Generate Logic
compareBtn.addEventListener("click", () => {
  if (!excelWorkbook || !docXmlContent) {
    alert("Please upload both the Excel file and Word document (.docx).");
    return;
  }
  executeComparison();
});

function executeComparison() {
  // 1. Extract and expand Excel Items
  const sheetName = excelWorkbook.SheetNames[0];
  const worksheet = excelWorkbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  let headerRowIndex = -1;
  let colIndices = {
    productName: -1,
    model: -1,
    barcode: -1,
    quantity: -1,
    unitPrice: -1,
    discount: -1,
  };

  for (let r = 0; r < Math.min(rawRows.length, 15); r++) {
    const row = rawRows[r] || [];
    const pIndex = row.findIndex(
      (c) => c && String(c).trim().toLowerCase() === "product name",
    );
    if (pIndex !== -1) {
      headerRowIndex = r;
      row.forEach((cellVal, idx) => {
        if (!cellVal) return;
        const col = String(cellVal).trim().toLowerCase();
        if (col === "product name") colIndices.productName = idx;
        else if (col === "model") colIndices.model = idx;
        else if (col === "barcode") colIndices.barcode = idx;
        else if (col === "quantity" || col === "qty") colIndices.quantity = idx;
        else if (col === "unit price" || col === "unitprice")
          colIndices.unitPrice = idx;
        else if (col === "discount" || col === "dis") colIndices.discount = idx;
      });
      break;
    }
  }

  if (headerRowIndex === -1 || colIndices.productName === -1) {
    alert("Could not locate 'Product Name' column in the Excel file.");
    return;
  }

  const excelItems = [];
  for (let r = headerRowIndex + 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || row.length === 0) continue;

    const productName =
      row[colIndices.productName] !== undefined
        ? String(row[colIndices.productName]).trim()
        : "";
    const model =
      colIndices.model !== -1 && row[colIndices.model] !== undefined
        ? String(row[colIndices.model]).trim()
        : "";
    const barcode =
      colIndices.barcode !== -1 && row[colIndices.barcode] !== undefined
        ? String(row[colIndices.barcode]).trim()
        : "";
    if (!productName && !barcode) continue;

    let qty = 1;
    if (colIndices.quantity !== -1 && row[colIndices.quantity] !== undefined) {
      const parsedQty = parseFloat(row[colIndices.quantity]);
      if (!isNaN(parsedQty) && parsedQty > 0) qty = Math.round(parsedQty);
    }

    const unitPrice =
      colIndices.unitPrice !== -1 && row[colIndices.unitPrice] !== undefined
        ? row[colIndices.unitPrice]
        : "";
    const discount =
      colIndices.discount !== -1 && row[colIndices.discount] !== undefined
        ? row[colIndices.discount]
        : 0;

    // Expand rows for Quantity > 1
    for (let i = 0; i < qty; i++) {
      excelItems.push({
        productName: productName,
        model: model,
        barcode: barcode,
        unitPrice: unitPrice,
        discount: discount,
      });
    }
  }

  // 2. Parse Word Document (.docx) XML
  const docItems = parseDocxXml(docXmlContent);

  // 3. Comparison & Matching by Model
  // Match key: Model (normalized uppercase)
  const docUsed = new Array(docItems.length).fill(false);
  let matchedCount = 0;
  let excelUnmatchedCount = 0;
  comparisonRows = [];

  for (const e of excelItems) {
    const eModelKey = (e.model || "").trim().toUpperCase();
    const eModelNorm = normalizeCode(e.model);
    const eBarcodeKey = (e.barcode || "").trim().toUpperCase();
    let matchedDocIdx = -1;
    let isPrefixMismatch = false;

    // PRIMARY KEY 1: Model Match
    // 1A. Exact Model match
    if (eModelKey) {
      for (let i = 0; i < docItems.length; i++) {
        if (
          !docUsed[i] &&
          (docItems[i].model || "").trim().toUpperCase() === eModelKey
        ) {
          matchedDocIdx = i;
          break;
        }
      }
    }

    // 1B. Normalized Model match (handles spacing/prefix differences e.g. "EF-2634-CPH J" vs "EF-2634-CPHJ" and "EF-26428-FTZF" vs "26428-FTZF")
    if (matchedDocIdx === -1 && eModelNorm) {
      for (let i = 0; i < docItems.length; i++) {
        if (!docUsed[i]) {
          const dNorm = normalizeCode(docItems[i].model);
          if (
            dNorm &&
            (dNorm === eModelNorm ||
              dNorm.includes(eModelNorm) ||
              eModelNorm.includes(dNorm))
          ) {
            matchedDocIdx = i;
            const dRaw = (docItems[i].model || "").trim().toUpperCase();
            if (eModelKey !== dRaw) {
              isPrefixMismatch = true;
            }
            break;
          }
        }
      }
    }

    // PRIMARY KEY 2: Barcode (Excel) <=> Style No. (Doc)
    if (matchedDocIdx === -1 && eBarcodeKey) {
      for (let i = 0; i < docItems.length; i++) {
        if (
          !docUsed[i] &&
          (docItems[i].barcode || "").trim().toUpperCase() === eBarcodeKey
        ) {
          matchedDocIdx = i;
          const dRaw = (docItems[i].model || "").trim().toUpperCase();
          if (eModelKey !== dRaw) {
            isPrefixMismatch = true;
          }
          break;
        }
      }
    }

    if (matchedDocIdx !== -1) {
      docUsed[matchedDocIdx] = true;
      matchedCount++;
      const d = docItems[matchedDocIdx];

      // Calculate differences
      const uPriceNum =
        typeof e.unitPrice === "number"
          ? e.unitPrice
          : parseFloat(String(e.unitPrice).replace(/,/g, ""));
      const docPriceNum =
        typeof d.unitPrice === "number"
          ? d.unitPrice
          : parseFloat(String(d.unitPrice).replace(/,/g, ""));
      const priceDiff =
        !isNaN(uPriceNum) && !isNaN(docPriceNum)
          ? Math.round((uPriceNum - docPriceNum) * 100) / 100
          : 0;

      const discNum =
        typeof e.discount === "number"
          ? e.discount
          : parseFloat(String(e.discount).replace(/,/g, ""));
      const docDiscNum =
        typeof d.discount === "number"
          ? d.discount
          : parseFloat(String(d.discount).replace(/,/g, ""));
      const discDiff =
        !isNaN(discNum) && !isNaN(docDiscNum)
          ? Math.round((discNum - docDiscNum) * 100) / 100
          : 0;

      comparisonRows.push({
        status: isPrefixMismatch ? "PREFIX_MISMATCH" : "MATCHED",
        statusLabel: isPrefixMismatch ? "Prefix Mismatch" : "Matched",
        docSl: d.docSl || "-",
        productName: e.productName,
        model: e.model,
        docModel: d.model,
        barcode: e.barcode,
        unitPrice: e.unitPrice,
        docUnitPrice: d.unitPrice,
        priceDiff: priceDiff,
        discount: e.discount,
        docDiscount: d.discount,
        discDiff: discDiff,
      });
    } else {
      excelUnmatchedCount++;
      comparisonRows.push({
        status: "EXCEL_ONLY",
        statusLabel: "Excel Only",
        productName: e.productName,
        model: e.model,
        barcode: e.barcode,
        unitPrice: e.unitPrice,
        docUnitPrice: "NOT MATCHED",
        priceDiff: "NOT MATCHED",
        discount: e.discount,
        docDiscount: "NOT MATCHED",
        discDiff: "NOT MATCHED",
      });
    }
  }

  // Remaining doc items not matched
  let docUnmatchedCount = 0;
  for (let i = 0; i < docItems.length; i++) {
    if (!docUsed[i]) {
      docUnmatchedCount++;
      const d = docItems[i];
      comparisonRows.push({
        status: "DOC_ONLY",
        statusLabel: "Doc Only",
        productName: d.goodsName,
        model: d.model,
        barcode: d.barcode,
        unitPrice: "NOT MATCHED",
        docUnitPrice: d.unitPrice,
        priceDiff: "NOT MATCHED",
        discount: "NOT MATCHED",
        docDiscount: d.discount,
        discDiff: "NOT MATCHED",
      });
    }
  }

  // 4. Update Stats UI
  matchedCountEl.textContent = matchedCount;
  excelUnmatchedCountEl.textContent = excelUnmatchedCount;
  docUnmatchedCountEl.textContent = docUnmatchedCount;
  totalRowCountEl.textContent = comparisonRows.length;
  statsContainer.style.display = "grid";
  filterBar.style.display = "flex";

  // 5. Render Preview Table
  renderFilteredRows();

  // 6. Show Download Button
  downloadBtn.style.display = "inline-flex";
}

// Helper to normalize model string (removes hyphens, spaces, punctuation for robust comparison)
function normalizeCode(str) {
  return String(str || "")
    .replace(/[^A-Za-z0-9]/g, "")
    .toUpperCase();
}

// XML parser for docx paragraphs & tabs
function parseDocxXml(xmlStr) {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlStr, "application/xml");
  const pNodes = xmlDoc.getElementsByTagName("w:p");

  const lines = [];
  for (let i = 0; i < pNodes.length; i++) {
    const p = pNodes[i];
    const rNodes = p.getElementsByTagName("w:r");
    let lineStr = "";

    for (let j = 0; j < rNodes.length; j++) {
      const r = rNodes[j];
      const tabs = r.getElementsByTagName("w:tab");
      for (let t = 0; t < tabs.length; t++) lineStr += "\t";

      const tNodes = r.getElementsByTagName("w:t");
      for (let t = 0; t < tNodes.length; t++) {
        lineStr += tNodes[t].textContent || "";
      }
    }
    lineStr = lineStr.trim();
    if (lineStr) lines.push(lineStr);
  }

  const items = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const parts = line
      .split("\t")
      .map((s) => s.trim())
      .filter(Boolean);

    // Identify sales data rows (having at least 5 tab fields, skipping table headers and summary totals)
    if (
      parts.length >= 5 &&
      !line.toLowerCase().includes("bill number") &&
      !line.toLowerCase().startsWith("sub total") &&
      !line.toLowerCase().startsWith("total ")
    ) {
      let overflow = "";
      // Check if next line is an overflow model code without tabs and not a summary line
      if (
        i + 1 < lines.length &&
        !lines[i + 1].includes("\t") &&
        !lines[i + 1].toLowerCase().startsWith("sub total") &&
        !lines[i + 1].toLowerCase().startsWith("total ")
      ) {
        overflow = lines[i + 1];
        i += 2;
      } else {
        i += 1;
      }

      // If first part is numeric (the SL), offset the fields by 1
      let offset = 0;
      let docSlNumber = items.length + 1;
      if (/^\d+$/.test(parts[0])) {
        docSlNumber = parseInt(parts[0], 10);
        offset = 1;
      }

      const billNo = parts[offset] || "";
      const goodsField = parts[offset + 1] || "";
      const styleNo = parts[offset + 2] || "";
      const priceStr = parts[offset + 3] || "";
      const disStr = parts[offset + 4] || "";

      let model = "";
      let goodsName = goodsField;

      if (overflow) {
        model = overflow.trim();
      } else {
        // The last token of goodsField is the model primary key (e.g. EF-2529-CPH)
        const tokens = goodsField.split(/\s+/);
        if (tokens.length > 1) {
          model = tokens[tokens.length - 1].trim();
          goodsName = tokens.slice(0, -1).join(" ");
        } else {
          model = tokens[0] || "";
        }
      }

      const priceNum = parseFloat(priceStr.replace(/,/g, ""));
      const disNum = parseFloat(disStr.replace(/,/g, ""));

      items.push({
        docSl: docSlNumber,
        billNo: billNo,
        goodsName: goodsName,
        model: model,
        barcode: styleNo,
        unitPrice: isNaN(priceNum) ? priceStr : priceNum,
        discount: isNaN(disNum) ? disStr : disNum,
      });
      continue;
    }
    i++;
  }

  return items;
}

// Filter handling
statusFilter.addEventListener("change", () => {
  renderFilteredRows();
});

function renderFilteredRows() {
  const filterVal = statusFilter.value;
  let filtered = comparisonRows;

  if (filterVal === "MATCHED") {
    filtered = comparisonRows.filter((r) => r.status === "MATCHED");
  } else if (filterVal === "PREFIX_MISMATCH") {
    filtered = comparisonRows.filter((r) => r.status === "PREFIX_MISMATCH");
  } else if (filterVal === "UNMATCHED") {
    filtered = comparisonRows.filter(
      (r) => r.status === "EXCEL_ONLY" || r.status === "DOC_ONLY"
    );
  } else if (filterVal === "EXCEL_ONLY") {
    filtered = comparisonRows.filter((r) => r.status === "EXCEL_ONLY");
  } else if (filterVal === "DOC_ONLY") {
    filtered = comparisonRows.filter((r) => r.status === "DOC_ONLY");
  }

  renderPreview(filtered);
}

// Render Preview Table
function renderPreview(rows) {
  previewSection.style.display = "block";
  previewCountBadge.textContent = `${rows.length} rows displayed`;
  previewTableBody.innerHTML = "";

  // Render all rows directly in the browser
  for (let i = 0; i < rows.length; i++) {
    const item = rows[i];
    const tr = document.createElement("tr");

    let rowClass = "row-matched";
    let tagClass = "matched";
    if (item.status === "PREFIX_MISMATCH") {
      rowClass = "row-prefix-mismatch";
      tagClass = "prefix-mismatch";
    } else if (item.status === "EXCEL_ONLY") {
      rowClass = "row-excel-only";
      tagClass = "excel-only";
    } else if (item.status === "DOC_ONLY") {
      rowClass = "row-doc-only";
      tagClass = "doc-only";
    }
    tr.className = rowClass;

    const unitPriceDisplay =
      item.unitPrice === "NOT MATCHED"
        ? `<span class="unmatched-cell">NOT MATCHED</span>`
        : typeof item.unitPrice === "number"
          ? item.unitPrice.toLocaleString()
          : escapeHtml(item.unitPrice);

    const docPriceDisplay =
      item.docUnitPrice === "NOT MATCHED"
        ? `<span class="unmatched-cell">NOT MATCHED</span>`
        : typeof item.docUnitPrice === "number"
          ? item.docUnitPrice.toLocaleString()
          : escapeHtml(item.docUnitPrice);

    const discountDisplay =
      item.discount === "NOT MATCHED"
        ? `<span class="unmatched-cell">NOT MATCHED</span>`
        : escapeHtml(item.discount);

    const docDiscountDisplay =
      item.docDiscount === "NOT MATCHED"
        ? `<span class="unmatched-cell">NOT MATCHED</span>`
        : escapeHtml(item.docDiscount);

    // Format difference values
    let priceDiffDisplay = `<span class="unmatched-cell">NOT MATCHED</span>`;
    let priceDiffClass = "col-diff-cell";
    if (item.priceDiff !== "NOT MATCHED") {
      const diffVal = Number(item.priceDiff);
      const sign = diffVal > 0 ? "+" : "";
      priceDiffDisplay = `${sign}${diffVal.toFixed(2)}`;
      priceDiffClass +=
        diffVal === 0
          ? " diff-zero"
          : diffVal > 0
            ? " diff-positive"
            : " diff-negative";
    }

    let discDiffDisplay = `<span class="unmatched-cell">NOT MATCHED</span>`;
    let discDiffClass = "col-diff-cell";
    if (item.discDiff !== "NOT MATCHED") {
      const diffVal = Number(item.discDiff);
      const sign = diffVal > 0 ? "+" : "";
      discDiffDisplay = `${sign}${diffVal.toFixed(2)}`;
      discDiffClass +=
        diffVal === 0
          ? " diff-zero"
          : diffVal > 0
            ? " diff-positive"
            : " diff-negative";
    }

    let modelDisplay = `<code>${escapeHtml(item.model)}</code>`;
    if (item.status === "PREFIX_MISMATCH" && item.docModel && item.model !== item.docModel) {
      modelDisplay = `<code>${escapeHtml(item.model)}</code><div style="font-size:0.75rem; color:#f59e0b; margin-top:2px;">(Doc: ${escapeHtml(item.docModel)})</div>`;
    }

    tr.innerHTML = `
      <td style="color: var(--text-secondary);">${i + 1}</td>
      <td style="font-weight: 700; color: var(--accent-blue); text-align: center;">${item.docSl !== "-" ? item.docSl : `<span class="unmatched-cell">-</span>`}</td>
      <td><span class="status-tag ${tagClass}">${escapeHtml(item.statusLabel)}</span></td>
      <td style="font-weight: 600;">${escapeHtml(item.productName)}</td>
      <td>${modelDisplay}</td>
      <td><code>${escapeHtml(item.barcode)}</code></td>
      <td>${unitPriceDisplay}</td>
      <td style="background: rgba(56, 189, 248, 0.05); font-weight: 600;">${docPriceDisplay}</td>
      <td class="${priceDiffClass}">${priceDiffDisplay}</td>
      <td>${discountDisplay}</td>
      <td style="background: rgba(56, 189, 248, 0.05); font-weight: 600;">${docDiscountDisplay}</td>
      <td class="${discDiffClass}">${discDiffDisplay}</td>
    `;
    previewTableBody.appendChild(tr);
  }

  // Calculate and render Sums in Table Footer
  let sumUnitPrice = 0;
  let sumDocUnitPrice = 0;
  let sumPriceDiff = 0;
  let sumDiscount = 0;
  let sumDocDiscount = 0;
  let sumDiscDiff = 0;

  for (const r of rows) {
    if (typeof r.unitPrice === "number") sumUnitPrice += r.unitPrice;
    if (typeof r.docUnitPrice === "number") sumDocUnitPrice += r.docUnitPrice;
    if (typeof r.priceDiff === "number") sumPriceDiff += r.priceDiff;
    if (typeof r.discount === "number") sumDiscount += r.discount;
    if (typeof r.docDiscount === "number") sumDocDiscount += r.docDiscount;
    if (typeof r.discDiff === "number") sumDiscDiff += r.discDiff;
  }

  if (previewTableFoot) {
    previewTableFoot.innerHTML = `
      <tr>
        <td colspan="6" class="total-label">TOTAL / SUM (${rows.length} rows)</td>
        <td class="sum-value">${sumUnitPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td class="sum-value">${sumDocUnitPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td class="sum-diff">${sumPriceDiff > 0 ? "+" : ""}${sumPriceDiff.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td class="sum-value">${sumDiscount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td class="sum-value">${sumDocDiscount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td class="sum-diff">${sumDiscDiff > 0 ? "+" : ""}${sumDiscDiff.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
      </tr>
    `;
  }
}

// Download Excel with Colors, Differences, Doc SL, and Sums row
downloadBtn.addEventListener("click", () => {
  if (comparisonRows.length === 0) return;

  // Prepare worksheet data with requested columns
  const exportData = comparisonRows.map((r) => ({
    Status: r.statusLabel,
    "Doc SL": r.docSl,
    "Product Name": r.productName,
    Model: r.model,
    Barcode: r.barcode,
    "Unit price": r.unitPrice,
    "doc u. price": r.docUnitPrice,
    "Price Diff": r.priceDiff,
    Discount: r.discount,
    "doc discount": r.docDiscount,
    "Disc Diff": r.discDiff,
  }));

  // Calculate totals for sums row
  let sumUnitPrice = 0;
  let sumDocUnitPrice = 0;
  let sumPriceDiff = 0;
  let sumDiscount = 0;
  let sumDocDiscount = 0;
  let sumDiscDiff = 0;

  for (const r of comparisonRows) {
    if (typeof r.unitPrice === "number") sumUnitPrice += r.unitPrice;
    if (typeof r.docUnitPrice === "number") sumDocUnitPrice += r.docUnitPrice;
    if (typeof r.priceDiff === "number") sumPriceDiff += r.priceDiff;
    if (typeof r.discount === "number") sumDiscount += r.discount;
    if (typeof r.docDiscount === "number") sumDocDiscount += r.docDiscount;
    if (typeof r.discDiff === "number") sumDiscDiff += r.discDiff;
  }

  // Append Total / Sum Row
  exportData.push({
    Status: "TOTAL SUM",
    "Doc SL": "",
    "Product Name": "",
    Model: "",
    Barcode: "",
    "Unit price": Math.round(sumUnitPrice * 100) / 100,
    "doc u. price": Math.round(sumDocUnitPrice * 100) / 100,
    "Price Diff": Math.round(sumPriceDiff * 100) / 100,
    Discount: Math.round(sumDiscount * 100) / 100,
    "doc discount": Math.round(sumDocDiscount * 100) / 100,
    "Disc Diff": Math.round(sumDiscDiff * 100) / 100,
  });

  const worksheet = XLSX.utils.json_to_sheet(exportData);

  // Set column widths
  worksheet["!cols"] = [
    { wch: 14 }, // Status
    { wch: 10 }, // Doc SL
    { wch: 28 }, // Product Name
    { wch: 18 }, // Model
    { wch: 18 }, // Barcode
    { wch: 14 }, // Unit price
    { wch: 15 }, // doc u. price
    { wch: 14 }, // Price Diff
    { wch: 12 }, // Discount
    { wch: 15 }, // doc discount
    { wch: 14 }, // Disc Diff
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Reconciliation Result");

  XLSX.writeFile(workbook, "Sales_Reconciliation_Comparison.xlsx");
});

function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
