let globalVendorDb = [];

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").addEventListener("click", fetchInvoices);
    document.getElementById("approveBtn").addEventListener("click", writeToExcel);
    document.getElementById("retryBtn").addEventListener("click", fetchInvoices);
    document.getElementById("cancelBtn").addEventListener("click", () => resetUI(true));
    
    // 🌟 ดักจับเหตุการณ์ปุ่ม Export
    document.getElementById("exportExpressBtn").addEventListener("click", exportToExpress);

    document.getElementById("cardsArea").addEventListener("click", function (e) {
      if (e.target && e.target.classList.contains("preview-img-btn")) {
        let url = e.target.getAttribute("data-url");
        showPreviewImage(url);
      }
      if (e.target && e.target.classList.contains("restore-btn")) {
        let idx = e.target.getAttribute("data-index");
        restoreInvoice(parseInt(idx));
      }
      if (e.target && (e.target.classList.contains("autocalc-btn") || e.target.closest(".autocalc-btn"))) {
        let btn = e.target.classList.contains("autocalc-btn") ? e.target : e.target.closest(".autocalc-btn");
        let idx = btn.getAttribute("data-index");
        let base = btn.getAttribute("data-base");
        autoCalculate(parseInt(idx), base);
      }
    });
    
    document.getElementById("cardsArea").addEventListener("change", function (e) {
        if (e.target && e.target.classList.contains("vendor-search-input")) {
            let idx = e.target.getAttribute("data-index");
            let selectedName = e.target.value;
            let matchedVendor = globalVendorDb.find(v => v.name === selectedName);
            
            if (matchedVendor) {
                pendingInvoices[idx].vendor_name = matchedVendor.name;
                pendingInvoices[idx].tax_id = matchedVendor.tax_id;
                pendingInvoices[idx].v_code = matchedVendor.v_code;
                pendingInvoices[idx].is_vendor_matched = true;
                
                document.getElementById(`tax_id_${idx}`).value = matchedVendor.tax_id;
                document.getElementById(`vcode_${idx}`).value = matchedVendor.v_code;
                updateState(idx, 'tax_id', matchedVendor.tax_id); 
            } else {
                pendingInvoices[idx].vendor_name = selectedName;
                pendingInvoices[idx].v_code = "";
                pendingInvoices[idx].is_vendor_matched = false;
                document.getElementById(`vcode_${idx}`).value = "";
                updateState(idx, 'tax_id', pendingInvoices[idx].tax_id);
            }
        }
    });

    document.addEventListener('keydown', function (event) {
      if (event.ctrlKey && event.key === 'Enter') {
        if (document.getElementById("reviewContainer").style.display === "block") {
          writeToExcel();
        }
      }
    });

    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange);
    
    await loadVendorDataForUI();
  }
});

function showToast(message, type = "success") {
    const container = document.getElementById("toastContainer");
    const toastId = "toast_" + Date.now();
    
    let icon = type === "success" ? "✅ " : (type === "error" ? "❌ " : "⚠️ ");
    
    const toastHTML = `
        <div id="${toastId}" class="message-bar ${type}">
            <span>${icon}${message}</span>
            <span class="close-msg" onclick="document.getElementById('${toastId}').remove()">✕</span>
        </div>
    `;
    container.innerHTML = toastHTML; 
    
    if(type === "success") {
        setTimeout(() => {
            let el = document.getElementById(toastId);
            if(el) el.remove();
        }, 5000);
    }
}

async function loadVendorDataForUI() {
    try {
        let res = await fetch("http://127.0.0.1:8080/get-vendors");
        let json = await res.json();
        if(json.status === "success") {
            globalVendorDb = json.data;
            let listHTML = "";
            globalVendorDb.forEach(v => {
                listHTML += `<option value="${v.name}">[${v.v_code}]${v.tax_id}</option>`;
            });
            document.getElementById("vendorList").innerHTML = listHTML;
        }
    } catch (e) {
        showToast("ไม่สามารถโหลดฐานข้อมูล V-Code ได้ (เซิร์ฟเวอร์ออฟไลน์)", "error");
    }
}

let pendingInvoices = [];

async function fetchInvoices() {
  const fileInput = document.getElementById("fileInput");
  if (fileInput.files.length === 0) return showToast("กรุณาเลือกไฟล์เอกสารก่อน", "warning");

  const formData = new FormData();
  for (let i = 0; i < fileInput.files.length; i++) {
    formData.append("file", fileInput.files[i]);
  }

  try {
    const runBtn = document.getElementById("run");
    const progressBar = document.getElementById("loadingProgress");
    const progressText = document.getElementById("loadingText");
    
    runBtn.disabled = true;
    runBtn.querySelector(".ms-Button-label").innerText = "กำลังประมวลผล...";
    progressBar.style.display = "block";
    progressText.style.display = "block";

    let response = await fetch("http://127.0.0.1:8080/extract-invoice", { method: "POST", body: formData });
    let result = await response.json();

    if (result.status === "success") {
      let invoices = Array.isArray(result.data) ? result.data : [result.data];

      let validDocsCount = invoices.filter(inv => inv.is_valid_invoice).length;
      let invalidDocsCount = invoices.length - validDocsCount;

      pendingInvoices = await filterInvoicesWithExcel(invoices);

      if (pendingInvoices.length > 0) {
        pendingInvoices.forEach(inv => inv._backup = JSON.parse(JSON.stringify(inv)));
        renderReviewCards();
        
        let msg = `ดึงข้อมูลสำเร็จ ${validDocsCount} รายการ`;
        if(invalidDocsCount > 0) msg += ` (ข้ามเอกสารที่ไม่ใช่บิล ${invalidDocsCount} รายการ)`;
        showToast(msg, "success");
        
      } else {
        showToast("ไม่มีข้อมูลที่สามารถนำเข้าได้ (อาจเป็นบิลซ้ำ หรือไม่ใช่ใบกำกับภาษี)", "warning");
        resetUI(false);
      }
    } else {
      throw new Error(result.message || "เซิร์ฟเวอร์ตอบกลับผิดพลาด");
    }
  } catch (error) {
    showToast(`เกิดข้อผิดพลาด: ${error.message}`, "error");
    resetUI(false);
  } finally {
      document.getElementById("loadingProgress").style.display = "none";
      document.getElementById("loadingText").style.display = "none";
  }
}

function renderReviewCards() {
  document.getElementById("uploadSection").style.display = "none";
  document.getElementById("reviewContainer").style.display = "block";
  document.getElementById("docCount").innerText = `${pendingInvoices.filter(i => i.is_valid_invoice).length} รายการ`;

  const cardsArea = document.getElementById("cardsArea");
  cardsArea.innerHTML = "";

  pendingInvoices.forEach((inv, index) => {
    if (inv.is_valid_invoice === false) {
      cardsArea.innerHTML += `
            <div class="review-card" style="border-left: 4px solid #a19f9d; text-align: center; background: #faf9f8;">
                <div style="font-weight: 600; font-size: 13px; color: #605e5c;">หน้า ${inv.page_number}</div>
                <div style="color: #a19f9d; font-size: 11px; margin-top: 5px;">⚠️ เอกสารนี้ไม่ใช่ใบกำกับภาษี (ระบบข้ามการบันทึก)</div>
            </div>
        `;
      return;
    }

    if (inv.is_wrong_company === true) {
      cardsArea.innerHTML += `
            <div class="review-card" style="border-left: 4px solid #a4262c; text-align: center; background: #fde7e9;">
                <div style="font-weight: 600; font-size: 13px; color: #a4262c;">หน้า ${inv.page_number}</div>
                <div style="color: #a4262c; font-size: 12px; margin-top: 5px; font-weight: bold;">⚠️ ห้ามบันทึก! บิลนี้ไม่ใช่ของบริษัทคุณ</div>
                <div style="color: #605e5c; font-size: 11px;">(Tax ID ผู้ซื้อไม่ตรงกับที่อยู่เซลล์ A4)</div>
            </div>
        `;
      return;
    }

    let mathError = Math.abs((parseFloat(inv.subtotal) + parseFloat(inv.vat)) - parseFloat(inv.total)) > 0.1;
    let cleanTaxId = inv.tax_id.replace(/\D/g, '');
    let taxIdError = cleanTaxId.length !== 13;
    let missingVCodeError = !inv.is_vendor_matched;

    let borderColor = "#107c41";
    let alertMsg = "";

    if (mathError) {
      borderColor = "#d13438";
      alertMsg += `<div style="color: #d13438; font-size: 11px; margin-bottom: 4px;">⚠️ ยอดรวมไม่ตรงกัน โปรดแก้ไข!</div>`;
    }
    if (taxIdError) {
      borderColor = mathError ? "#d13438" : "#ffaa44";
      alertMsg += `<div style="color: #c97104; font-size: 11px; margin-bottom: 4px;">⚠️ เลขผู้เสียภาษีไม่ครบ 13 หลัก!</div>`;
    }
    if (missingVCodeError) {
      borderColor = (mathError || taxIdError) ? borderColor : "#ffaa44";
      alertMsg += `<div style="color: #c97104; font-size: 11px; margin-bottom: 6px;">⚠️ โปรดค้นหารหัสผู้ขายจากฐานข้อมูล V-Code</div>`;
    }

    let vendorBadge = inv.is_vendor_matched ? `<span class="matched-badge">✔️ ในระบบ</span>` : ``;

    let cardHTML = `
            <div id="card_${index}" class="review-card" style="border-left: 4px solid ${borderColor};">
                <div class="review-header">
                    <span style="font-weight: 600; font-size: 13px; color: #323130;">บิลที่ ${index + 1} (หน้า ${inv.page_number})</span>
                    <div>
                        <span class="preview-img-btn" data-url="${inv.file_path}" style="cursor: pointer; color: #0078d4; font-size: 12px; margin-right: 10px;">📸 รูปบิล</span>
                        <span class="restore-btn" data-index="${index}" style="cursor: pointer; color: #605e5c; font-size: 12px;">🔄 คืนค่า</span>
                    </div>
                </div>
                
                <div id="alert_${index}">${alertMsg}</div>
                
                <div style="display: flex; flex-direction: column; gap: 8px; font-size: 12px;">
                    <label>เลขบิล: <input tabindex="1" type="text" style="width: 100%; box-sizing: border-box;" id="inv_no_${index}" value="${inv.invoice_no}" onchange="updateState(${index}, 'invoice_no', this.value)"></label>
                    
                    <div style="display: flex; gap: 8px; align-items: flex-end;">
                        <label style="flex: 3;">ผู้ขาย: ${vendorBadge}
                            <input tabindex="1" list="vendorList" type="text" class="vendor-search-input" data-index="${index}" style="width: 100%; box-sizing: border-box; ${!inv.is_vendor_matched ? 'border-color: #ffaa44; background: #fff4ce;' : ''}" id="vendor_${index}" value="${inv.vendor_name}">
                        </label>
                        <label style="flex: 1;">V-Code:
                            <input type="text" style="width: 100%; box-sizing: border-box;" class="vcode-input" id="vcode_${index}" value="${inv.v_code}" readonly tabindex="-1">
                        </label>
                    </div>

                    <label>Tax ID (13 หลัก): <input tabindex="1" type="text" style="width: 100%; box-sizing: border-box; ${taxIdError ? 'border-color: #ffaa44; background: #fff4ce;' : ''}" id="tax_id_${index}" value="${inv.tax_id}" onchange="updateState(${index}, 'tax_id', this.value)"></label>
                    
                    <div style="display: flex; gap: 8px; align-items: center; background: #f3f2f1; padding: 8px; border-radius: 4px; margin-top: 5px;">
                        <label style="flex: 1; margin: 0;">มูลค่า: 
                            <div style="display: flex; gap: 2px;">
                                <input tabindex="1" type="number" style="width: 100%; border: none;" id="subtotal_${index}" value="${inv.subtotal}" onchange="updateState(${index}, 'subtotal', this.value)">
                                <button tabindex="-1" class="autocalc-btn" data-index="${index}" data-base="subtotal" title="คำนวณ VAT" style="background: #e1dfdd; border: none; cursor: pointer; border-radius: 2px;">🪄</button>
                            </div>
                        </label>
                        <label style="flex: 1; margin: 0;">VAT: <input tabindex="1" type="number" style="width: 100%; border: none;" id="vat_${index}" value="${inv.vat}" onchange="updateState(${index}, 'vat', this.value)"></label>
                        <label style="flex: 1; margin: 0;">สุทธิ: 
                            <div style="display: flex; gap: 2px;">
                                <input tabindex="1" type="number" style="width: 100%; font-weight: bold; border: none; color: #107c41;" id="total_${index}" value="${inv.total}" onchange="updateState(${index}, 'total', this.value)">
                                <button tabindex="-1" class="autocalc-btn" data-index="${index}" data-base="total" title="ถอด VAT" style="background: #e1dfdd; border: none; cursor: pointer; border-radius: 2px;">🪄</button>
                            </div>
                        </label>
                    </div>
                </div>
            </div>
        `;
    cardsArea.innerHTML += cardHTML;
  });
}

window.autoCalculate = function (index, base) {
  let inv = pendingInvoices[index];
  if (base === 'total') {
    let t = parseFloat(inv.total) || 0;
    let sub = t / 1.07;
    let v = t - sub;
    inv.subtotal = parseFloat(sub.toFixed(2));
    inv.vat = parseFloat(v.toFixed(2));
  } else if (base === 'subtotal') {
    let sub = parseFloat(inv.subtotal) || 0;
    let v = sub * 0.07;
    let t = sub + v;
    inv.vat = parseFloat(v.toFixed(2));
    inv.total = parseFloat(t.toFixed(2));
  }

  document.getElementById(`subtotal_${index}`).value = inv.subtotal;
  document.getElementById(`vat_${index}`).value = inv.vat;
  document.getElementById(`total_${index}`).value = inv.total;

  updateState(index, 'total', inv.total.toString());
  showToast("คำนวณ VAT อัตโนมัติสำเร็จ", "success");
};

function updateState(index, field, value) {
  pendingInvoices[index][field] = field.includes('total') || field === 'vat' || field === 'subtotal' ? parseFloat(value) || 0 : value;
  let inv = pendingInvoices[index];

  if (field === 'tax_id') {
    inv.tax_id = value.replace(/\D/g, '');
    document.getElementById(`tax_id_${index}`).value = inv.tax_id;
  }

  inv.is_math_valid = Math.abs((inv.subtotal + inv.vat) - inv.total) <= 0.1;

  let cardDiv = document.getElementById(`card_${index}`);
  let alertDiv = document.getElementById(`alert_${index}`);
  let taxInput = document.getElementById(`tax_id_${index}`);
  let vendorInput = document.getElementById(`vendor_${index}`);

  let mathError = !inv.is_math_valid;
  let taxIdError = inv.tax_id.length !== 13;
  let missingVCodeError = !inv.is_vendor_matched;

  let borderColor = "#107c41";
  let alertMsg = "";

  if (mathError) {
    borderColor = "#d13438";
    alertMsg += `<div style="color: #d13438; font-size: 11px; margin-bottom: 2px;">⚠️ ยอดรวมไม่ตรงกัน โปรดแก้ไข!</div>`;
  }
  if (taxIdError) {
    borderColor = mathError ? "#d13438" : "#ffaa44";
    alertMsg += `<div style="color: #c97104; font-size: 11px; margin-bottom: 2px;">⚠️ เลขผู้เสียภาษีไม่ครบ 13 หลัก!</div>`;
    if (taxInput) { taxInput.style.borderColor = "#ffaa44"; taxInput.style.background = "#fff4ce"; }
  } else {
    if (taxInput) { taxInput.style.borderColor = "#605e5c"; taxInput.style.background = ""; }
  }
  
  if (missingVCodeError) {
      borderColor = (mathError || taxIdError) ? borderColor : "#ffaa44";
      alertMsg += `<div style="color: #c97104; font-size: 11px; margin-bottom: 5px;">⚠️ โปรดค้นหารหัสผู้ขายจากฐานข้อมูล V-Code</div>`;
      if (vendorInput) { vendorInput.style.borderColor = "#ffaa44"; vendorInput.style.background = "#fff4ce"; }
  } else {
      if (vendorInput) { vendorInput.style.borderColor = "#605e5c"; vendorInput.style.background = ""; }
  }

  if (cardDiv) cardDiv.style.borderLeft = `4px solid ${borderColor}`;
  if (alertDiv) alertDiv.innerHTML = alertMsg;
}

window.restoreInvoice = function (index) {
  let backupData = pendingInvoices[index]._backup;
  if (backupData) {
    pendingInvoices[index] = JSON.parse(JSON.stringify(backupData));
    pendingInvoices[index]._backup = backupData;

    let inv = pendingInvoices[index];
    document.getElementById(`inv_no_${index}`).value = inv.invoice_no;
    document.getElementById(`vendor_${index}`).value = inv.vendor_name;
    document.getElementById(`tax_id_${index}`).value = inv.tax_id;
    document.getElementById(`subtotal_${index}`).value = inv.subtotal;
    document.getElementById(`vat_${index}`).value = inv.vat;
    document.getElementById(`total_${index}`).value = inv.total;
    document.getElementById(`vcode_${index}`).value = inv.v_code;

    updateState(index, 'total', inv.total.toString());
    showToast("คืนค่าการสกัดดั้งเดิมสำเร็จ", "success");
  }
};

function showPreviewImage(url) {
  let imgEl = document.getElementById("billPreviewImg");
  let noImgText = document.getElementById("noImageText");
  if (url) {
    imgEl.src = url;
    imgEl.style.display = "block";
    noImgText.style.display = "none";
    document.getElementById("imageViewerSection").scrollIntoView({ behavior: 'smooth' });
  }
}

function resetUI(clearFile = true) {
  pendingInvoices = [];
  document.getElementById("uploadSection").style.display = "block";
  document.getElementById("reviewContainer").style.display = "none";
  document.getElementById("run").querySelector(".ms-Button-label").innerText = "สกัดข้อมูลอัตโนมัติ";
  document.getElementById("run").disabled = false;
  document.getElementById("loadingProgress").style.display = "none";
  document.getElementById("loadingText").style.display = "none";

  if (clearFile) {
    document.getElementById("fileInput").value = "";
  }
}

async function filterInvoicesWithExcel(invoices) {
  let validList = [];
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    let companyInfoRange = sheet.getRange("A4");
    companyInfoRange.load("values");

    let headerRange = sheet.getRange("A2:K2");
    headerRange.load("values");
    let searchRange = sheet.getRange("A1:A500");
    let totalCell = searchRange.findOrNullObject("รวม", { completeMatch: true });
    totalCell.load("rowIndex");
    await context.sync();

    let myCompanyTaxId = "";
    if (companyInfoRange.values && companyInfoRange.values[0] && companyInfoRange.values[0][0]) {
      let textA4 = companyInfoRange.values[0][0].toString();
      let match = textA4.match(/\d{13}/);
      if (match) {
        myCompanyTaxId = match[0];
      }
    }

    let headerText = headerRange.values[0].join(" ");
    let targetMonth = 0;
    const months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
    months.forEach((m, i) => { if (headerText.includes(m)) targetMonth = i + 1; });

    if (totalCell.isNullObject) {
      showToast("ไม่พบบรรทัด 'รวม' ในฟอร์ม Excel กรุณาตรวจสอบ", "error");
      return;
    }

    let invoiceNoRange = sheet.getRangeByIndexes(5, 2, totalCell.rowIndex - 5, 1);
    invoiceNoRange.load("values");
    await context.sync();

    let existingInvoices = new Set();
    invoiceNoRange.values.forEach(row => { if (row[0]) existingInvoices.add(row[0].toString().trim()); });

    for (let inv of invoices) {
      inv.is_wrong_company = false;
      if (myCompanyTaxId && inv.buyer_tax_id) {
        if (inv.buyer_tax_id !== myCompanyTaxId) {
          inv.is_wrong_company = true;
        }
      }

      if (inv.is_valid_invoice === false || inv.is_wrong_company === true) {
        validList.push(inv);
        continue;
      }

      if (targetMonth !== 0 && inv.doc_month !== 0 && inv.doc_month !== targetMonth) continue;
      if (existingInvoices.has(inv.invoice_no.toString().trim())) continue;
      
      validList.push(inv);
    }
  });
  return validList;
}

async function writeToExcel() {
  let finalInvoicesToWrite = pendingInvoices.filter(inv => inv.is_valid_invoice !== false && inv.is_wrong_company !== true);

  const hasMathError = finalInvoicesToWrite.some(inv => !inv.is_math_valid);
  const hasTaxError = finalInvoicesToWrite.some(inv => inv.tax_id.replace(/\D/g, '').length !== 13);
  const hasVCodeError = finalInvoicesToWrite.some(inv => !inv.is_vendor_matched);

  if (hasMathError || hasTaxError || hasVCodeError) {
    showToast("กรุณาแก้ไขข้อมูลที่ผิดพลาด หรือกรอกรหัส V-Code ให้ครบก่อนบันทึก", "error");
    return;
  }

  if (finalInvoicesToWrite.length === 0) {
    showToast("ไม่มีเอกสารที่สมบูรณ์ให้บันทึกลงตาราง", "warning");
    resetUI(true);
    return;
  }

  try {
    document.getElementById("approveBtn").disabled = true;
    document.getElementById("approveBtn").querySelector(".ms-Button-label").innerText = "กำลังเขียนลงตาราง...";

    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();

      let searchRange = sheet.getRange("A1:A500");
      let totalCell = searchRange.findOrNullObject("รวม", { completeMatch: true });
      totalCell.load("rowIndex");
      await context.sync();
      let totalRowIdx = totalCell.rowIndex;

      let dataRange = sheet.getRangeByIndexes(5, 2, totalRowIdx - 5, 1);
      dataRange.load("values");
      await context.sync();

      let firstEmptyRowOffset = 0;
      dataRange.values.forEach((v, i) => { if (v[0] !== "" && v[0] !== null) firstEmptyRowOffset = i + 1; });

      let targetRowIdx = 5 + firstEmptyRowOffset;
      let availableBlankRows = totalRowIdx - targetRowIdx;
      let rowCount = finalInvoicesToWrite.length;

      if (rowCount > availableBlankRows) {
        let insertRange = sheet.getRangeByIndexes(totalRowIdx, 0, rowCount - availableBlankRows, 12);
        insertRange.insert(Excel.InsertShiftDirection.down);
      }

      let dataList = [];
      for (let i = 0; i < rowCount; i++) {
        let inv = finalInvoicesToWrite[i];
        dataList.push([
          firstEmptyRowOffset + i + 1,
          "'" + inv.date,
          inv.invoice_no,
          inv.vendor_name,
          inv.tax_id,
          inv.is_hq ? "/" : "",
          inv.is_hq ? "" : inv.branch_no,
          inv.total,
          inv.subtotal,
          inv.vat,
          inv.file_path,
          inv.v_code
        ]);
      }

      let writeRange = sheet.getRangeByIndexes(targetRowIdx, 0, rowCount, 12);
      writeRange.values = dataList;
      writeRange.format.font.name = "Tahoma";
      writeRange.format.font.size = 10;
      sheet.getRangeByIndexes(targetRowIdx, 0, rowCount, 2).format.horizontalAlignment = "Center";

      let secretUrlColumn = sheet.getRange("K:K");
      secretUrlColumn.columnHidden = true;

      await context.sync();
    });

    showToast(`นำเข้าสำเร็จ ${finalInvoicesToWrite.length} รายการ เรียบร้อยแล้ว!`, "success");
    resetUI(true);

  } catch (error) {
    showToast(`เกิดข้อผิดพลาดในการเขียนตาราง: ${error.message}`, "error");
    document.getElementById("approveBtn").disabled = false;
    document.getElementById("approveBtn").querySelector(".ms-Button-label").innerText = "บันทึกข้อมูลลงตาราง";
  }
}

async function onSelectionChange(event) {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = context.workbook.getSelectedRange();
      range.load("rowIndex");
      await context.sync();

      let rowIndex = range.rowIndex;

      if (rowIndex >= 4) {
        let linkCell = sheet.getRangeByIndexes(rowIndex, 10, 1, 1);
        linkCell.load("values");
        await context.sync();

        let url = linkCell.values[0][0];
        let imgEl = document.getElementById("billPreviewImg");
        let noImgText = document.getElementById("noImageText");

        if (url && typeof url === 'string' && url.includes("http://127.0.0.1:8080/saved_bills")) {
          imgEl.src = url;
          imgEl.style.display = "block";
          noImgText.style.display = "none";
        } else {
          imgEl.style.display = "none";
          noImgText.style.display = "block";
        }
      }
    });
  } catch (error) {
    console.error("Selection Change Error:", error);
  }
}

// 🌟 ฟังก์ชันใหม่: ดึงข้อมูลจากตารางมาสร้างไฟล์ CSV สำหรับโปรแกรมบัญชี
async function exportToExpress() {
    try {
        let exportData = [];
        let exportBtn = document.getElementById("exportExpressBtn");
        exportBtn.disabled = true;
        exportBtn.querySelector(".ms-Button-label").innerText = "กำลังสร้างไฟล์...";

        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // หาบรรทัดคำว่า "รวม"
            let searchRange = sheet.getRange("A1:A500");
            let totalCell = searchRange.findOrNullObject("รวม", { completeMatch: true });
            totalCell.load("rowIndex");
            await context.sync();
            
            if (totalCell.isNullObject) {
                throw new Error("ไม่พบบรรทัด 'รวม' ในฟอร์ม");
            }
            
            // อ่านตารางข้อมูลตั้งแต่บรรทัดที่ 5 ถึงก่อนบรรทัดรวม (เอาถึงคอลัมน์ L ที่เก็บ V-Code ไว้)
            let dataRange = sheet.getRangeByIndexes(5, 1, totalCell.rowIndex - 5, 11);
            dataRange.load("values");
            await context.sync();

            // คัดกรองและจัด Format ข้อมูล
            dataRange.values.forEach(row => {
                let dateStr = row[0]; // วันที่
                let invoiceNo = row[1]; // เลขบิล
                let taxId = row[3]; // เลขภาษี
                let total = row[6]; // ยอดรวม
                let vCode = row[10]; // 🌟 V-Code อยู่ในคอลัมน์ L (index 10)

                // ถ้าบรรทัดไหนมีเลขบิล และมี V-Code แสดงว่าพร้อมส่งออก
                if (invoiceNo && vCode) {
                    // ลบเครื่องหมาย ' ออกจากวันที่
                    let cleanDate = dateStr ? dateStr.toString().replace(/'/g, "") : "";
                    
                    // สร้าง Array 1 บรรทัด (จัดรูปแบบตามใจชอบได้เลย)
                    exportData.push([vCode, cleanDate, invoiceNo, taxId, total]);
                }
            });
        });

        if (exportData.length === 0) {
            showToast("ไม่พบข้อมูลที่มี V-Code พร้อมสำหรับการส่งออก", "warning");
            exportBtn.disabled = false;
            exportBtn.querySelector(".ms-Button-label").innerText = "⬇️ ส่งออกไฟล์สำหรับ Express";
            return;
        }

        // 🌟 สร้างเนื้อหาไฟล์ CSV
        let csvContent = "V-Code,Date,InvoiceNo,TaxID,TotalAmount\n";
        exportData.forEach(row => {
            csvContent += row.join(",") + "\n";
        });

        // 🌟 สร้างไฟล์และสั่งดาวน์โหลดลงเครื่อง
        let blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' }); // ใส่ BOM ให้รองรับภาษาไทย
        let link = document.createElement("a");
        let url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.setAttribute("download", `Express_Import_${Date.now()}.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        showToast(`ส่งออกไฟล์ ${exportData.length} รายการสำเร็จ! (เช็คที่แฟ้มดาวน์โหลด)`, "success");
        
    } catch (error) {
        showToast(`เกิดข้อผิดพลาดในการส่งออก: ${error.message}`, "error");
    } finally {
        let exportBtn = document.getElementById("exportExpressBtn");
        exportBtn.disabled = false;
        exportBtn.querySelector(".ms-Button-label").innerText = "⬇️ ส่งออกไฟล์สำหรับ Express";
    }
}