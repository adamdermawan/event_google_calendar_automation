// ============================================================
// SALES PERFORMANCE DASHBOARD - Unified Backend
// Deals + Revenue + Calendar
// ============================================================

const SALES_SPREADSHEET_ID = "1m0tiOB6MWexUC4Bx-Slfn7u3PMldvnu4B4WFBnT2i0A";
const DOC_LIBRARY_SPREADSHEET_ID = "1QThThONBgXncK2gV7viV1-H5sjKzrKWE1dsh1r9Oj6c";

const SHEET_NAME = "Event Data";
const SHEET_MASTER_AM = "Event Data AM";
const SHEET_MASTER_PROD = "Event Data";
const SHEET_REVENUE = "Event Data";
const SHEET_CALENDARS = "Event Data";
const SHEET_PIPELINE_CUTOFF = "Event Data";
const SHEET_ORDER_DETAIL = "Event Data";
const SHEET_REVENUE_XLS = "Event Data";

const DOC_SHEET_NAME = "DATA";
const DOC_HEADERS = [
  "FOLDER ID",
  "ISP",
  "CITIES",
  "FILE NAME",
  "FILE URL",
  "FILE ID",
  "FILE DATE",
  "DATE UPLOADED",
  "DESKRIPSI",
  "JENIS FILE",
  "STATUS",
  "OWNER"
];

const DOC_STATUS_OPTIONS = ["Draft", "Active", "Archived"];
const DOC_FILE_TYPE_OPTIONS = ["PDF", "DOC", "TXT", "XLS", "PPT", "Image", "Other"];

const DROPDOWN_OPTIONS = {
  namaAM: ["Bunga", "Teuku", "Hans", "Filbert", "Lita", "Rodhi", "Navalin", "Khairina", "Amri"],
  serviceName: ["Leased Line", "IPT - IX", "IPT - Mix", "CDN", "IPLC", "Colocation", "Dark Fiber", "Cross Connect", "Homepass (B2s/OA)"],
  stage: ["New", "Requirement", "Presentation Solution", "Negotiation", "Closed Won", "Closed Lost"],
  jenisProduct: ["Wholesale", "B2S", "Open Access"],
  orderStatus: ["Cold", "Hot", "Won", "Lost"],
};

function getSheet() {
  var ss = getSalesSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + SHEET_NAME + "' tidak ditemukan.");
  return sheet;
}

function getCalendarSheet() {
  var ss = getSalesSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_CALENDARS);
  if (!sheet) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim().toLowerCase() === SHEET_CALENDARS.trim().toLowerCase()) {
        sheet = sheets[i];
        break;
      }
    }
  }
  if (!sheet) throw new Error("Sheet '" + SHEET_CALENDARS + "' tidak ditemukan.");
  return sheet;
}

function getRevenueSheet() {
  var ss = getSalesSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_REVENUE);
  if (!sheet) throw new Error("Sheet '" + SHEET_REVENUE + "' tidak ditemukan.");
  return sheet;
}

function getRevenueXLSSheet() {
  var ss = getSalesSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_REVENUE_XLS);
  if (!sheet) throw new Error("Sheet '" + SHEET_REVENUE_XLS + "' tidak ditemukan.");
  return sheet;
}

function getPipelineCutoffSheet() {
  var ss = getSalesSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_PIPELINE_CUTOFF);
  if (!sheet) throw new Error("Sheet '" + SHEET_PIPELINE_CUTOFF + "' tidak ditemukan.");
  return sheet;
}

function getOrderDetailSheet() {
  var ss = getSalesSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_ORDER_DETAIL);
  if (!sheet) throw new Error("Sheet '" + SHEET_ORDER_DETAIL + "' tidak ditemukan.");
  return sheet;
}

function getSalesSpreadsheet_() {
  return SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
}

function getDocumentSpreadsheet_() {
  return SpreadsheetApp.openById(DOC_LIBRARY_SPREADSHEET_ID);
}

function normalizeHeaderName(header) {
  return String(header || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function buildHeaderMap(headers) {
  var map = {};
  headers.forEach(function(header, index) {
    var key = normalizeHeaderName(header);
    if (key && map[key] === undefined) map[key] = index;
  });
  return map;
}

function getValueByHeader(row, headerMap, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headerMap[normalizeHeaderName(names[i])];
    if (idx !== undefined) return row[idx];
  }
  return "";
}

// ============================================================
// LEADS ID GENERATOR
// ============================================================

function getISPCode(isp) {
  if (!isp) return "UNK";
  var words = isp.trim().split(/\s+/).filter(function(w) { return w; });
  if (words.length === 1) return words[0].replace(/[^a-zA-Z0-9]/g, "").slice(0, 4).toUpperCase();
  if (words.length === 2) return words[0].replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
  return words.map(function(w) { return w[0].toUpperCase(); }).join("");
}

function getAMCode(namaAM) {
  if (!namaAM) return "UNK";
  try {
    var sheet = getSalesSpreadsheet_().getSheetByName(SHEET_MASTER_AM);
    if (!sheet) return namaAM.replace(/[^a-zA-Z0-9]/g, "").slice(0, 3).toUpperCase();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === namaAM.trim().toLowerCase()) {
        return String(data[i][1]).trim().toUpperCase();
      }
    }
  } catch (e) {}
  return namaAM.replace(/[^a-zA-Z0-9]/g, "").slice(0, 3).toUpperCase();
}

function getProductCode(serviceName, jenisProduct) {
  if (!serviceName && !jenisProduct) return "UNK";
  try {
    var sheet = getSalesSpreadsheet_().getSheetByName(SHEET_MASTER_PROD);
    if (!sheet) return "UNK";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var svc = String(data[i][0]).trim().toLowerCase();
      var prod = String(data[i][1]).trim().toLowerCase();
      if (svc === (serviceName || "").trim().toLowerCase() &&
          prod === (jenisProduct || "").trim().toLowerCase()) {
        return String(data[i][2]).trim().toUpperCase();
      }
    }
  } catch (e) {}
  var s = (serviceName || "").replace(/[^a-zA-Z0-9]/g, "").slice(0, 2).toUpperCase();
  var p = (jenisProduct || "").replace(/[^a-zA-Z0-9]/g, "").slice(0, 2).toUpperCase();
  return s + p;
}

function formatTanggalLeads(tanggal) {
  var tgl;
  if (!tanggal) {
    tgl = new Date();
  } else if (tanggal instanceof Date) {
    tgl = tanggal;
  } else {
    var parts = String(tanggal).split("-");
    if (parts.length === 3) {
      tgl = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    } else {
      tgl = new Date(tanggal);
    }
  }
  var dd = String(tgl.getDate()).padStart(2, "0");
  var mm = String(tgl.getMonth() + 1).padStart(2, "0");
  var yy = String(tgl.getFullYear()).slice(-2);
  return dd + mm + yy;
}

function makeSignature(namaAM, isp, serviceName, jenisProduct, tanggal) {
  var tgl = formatTanggalLeads(tanggal);
  return [namaAM, isp, serviceName, jenisProduct, tgl]
    .map(function(s) { return String(s || "").trim().toLowerCase(); })
    .join("|");
}

function buildLeadsID(namaAM, isp, serviceName, jenisProduct, tanggal, urut) {
  var amCode = getAMCode(namaAM);
  var ispCode = getISPCode(isp);
  var prodCode = getProductCode(serviceName, jenisProduct);
  var tglCode = formatTanggalLeads(tanggal);
  var urutStr = "LD" + String(urut).padStart(3, "0");
  return urutStr + "-" + amCode + "-" + ispCode + "-" + prodCode + "-" + tglCode;
}

function generateNewLeadsID(namaAM, isp, serviceName, jenisProduct, tanggal) {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  var sig = makeSignature(namaAM, isp, serviceName, jenisProduct, tanggal);
  var count = 0;

  if (lastRow >= 2) {
    var vals = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    vals.forEach(function(row) {
      var isEmpty = row.slice(0, 6).every(function(c) { return c === "" || c === null; });
      if (isEmpty) return;
      var rowSig = makeSignature(
        String(row[1] || ""),
        String(row[2] || ""),
        String(row[3] || ""),
        String(row[5] || ""),
        row[0]
      );
      if (rowSig === sig) count++;
    });
  }

  return buildLeadsID(namaAM, isp, serviceName, jenisProduct, tanggal, count + 1);
}

function backfillLeadsID() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Sheet DATA kosong.");
    return;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  var updates = [];
  var sigCount = {};
  var filled = 0;

  values.forEach(function(row) {
    var isEmpty = row.slice(0, 6).every(function(cell) { return cell === "" || cell === null; });
    if (isEmpty) {
      updates.push([""]);
      return;
    }

    var tanggal = row[0];
    var namaAM = String(row[1] || "").trim();
    var isp = String(row[2] || "").trim();
    var serviceName = String(row[3] || "").trim();
    var jenisProduct = String(row[5] || "").trim();

    var sig = makeSignature(namaAM, isp, serviceName, jenisProduct, tanggal);
    sigCount[sig] = (sigCount[sig] || 0) + 1;
    updates.push([buildLeadsID(namaAM, isp, serviceName, jenisProduct, tanggal, sigCount[sig])]);
    filled++;
  });

  sheet.getRange(2, 12, updates.length, 1).setValues(updates);
  SpreadsheetApp.getUi().alert("✅ Backfill selesai!\n" + filled + " Leads ID berhasil di-generate di kolom L.");
}

function debugBackfill() {
  var sheet = getSheet();
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var sigCount = {};

  values.forEach(function(row, idx) {
    var namaAM = String(row[1] || "").trim();
    var isp = String(row[2] || "").trim();
    var serviceName = String(row[3] || "").trim();
    var jenisProduct = String(row[5] || "").trim();
    var tanggal = row[0];

    var sig = makeSignature(namaAM, isp, serviceName, jenisProduct, tanggal);
    sigCount[sig] = (sigCount[sig] || 0) + 1;
    Logger.log("Baris " + (idx + 2) + " | sig: " + sig + " | urut: " + sigCount[sig]);
  });
}

// ============================================================
// WEB APP & MENU
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile("dashboard")
    .setTitle("Sales Performance Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Sales Dashboard")
    .addItem("Buka Dashboard (Sidebar)", "openSidebar")
    .addItem("Buka Dashboard (Web App)", "openWebApp")
    .addSeparator()
    .addItem("🔄 Generate / Backfill Leads ID", "backfillLeadsID")
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("dashboard")
    .setTitle("Sales Performance Dashboard")
    .setWidth(1400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openWebApp() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput(
    "<script>window.open('" + url + "','_blank');google.script.host.close();</script>"
  );
  SpreadsheetApp.getUi().showModalDialog(html, "Membuka Web App...");
}

// ============================================================
// DEALS API
// ============================================================

function getDropdownOptions() {
  return { success: true, options: DROPDOWN_OPTIONS };
}

function getAllDeals() {
  try {
    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    var values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    var data = values
      .filter(function(row) {
        return row.some(function(cell) { return cell !== "" && cell !== null; });
      })
      .map(function(row, i) {
        return {
          rowIndex: i + 2,
          tanggal: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
          namaAM: String(row[1] || ""),
          isp: String(row[2] || ""),
          serviceName: String(row[3] || ""),
          stage: String(row[4] || ""),
          jenisProduct: String(row[5] || ""),
          estimasiHP: Number(row[6]) || 0,
          amountMRC: Number(row[7]) || 0,
          amountOTC: Number(row[8]) || 0,
          orderStatus: String(row[9] || ""),
          keterangan: String(row[10] || ""),
          leadsID: String(row[11] || ""),
          activityID: String(row[12] || "")
        };
      });

    return { success: true, data: data };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addDeal(d) {
  try {
    var sheet = getSheet();
    var leadsID = generateNewLeadsID(d.namaAM, d.isp, d.serviceName, d.jenisProduct, d.tanggal);
    var newRow = [
      d.tanggal ? new Date(d.tanggal) : new Date(),
      d.namaAM,
      d.isp,
      d.serviceName,
      d.stage,
      d.jenisProduct,
      d.estimasiHP || 0,
      d.amountMRC || 0,
      d.amountOTC || 0,
      d.orderStatus,
      d.keterangan,
      leadsID,
      ""
    ];
    var lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, 13).setValues([newRow]);
    return { success: true, rowIndex: lastRow, leadsID: leadsID };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateDeal(d) {
  try {
    var sheet = getSheet();
    var existingLeadsID = sheet.getRange(d.rowIndex, 12).getValue();
    var updatedRow = [
      d.tanggal ? new Date(d.tanggal) : new Date(),
      d.namaAM,
      d.isp,
      d.serviceName,
      d.stage,
      d.jenisProduct,
      d.estimasiHP || 0,
      d.amountMRC || 0,
      d.amountOTC || 0,
      d.orderStatus,
      d.keterangan,
      existingLeadsID,
      d.activityID || ""
    ];
    sheet.getRange(d.rowIndex, 1, 1, 13).setValues([updatedRow]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteDeal(rowIndex) {
  try {
    getSheet().deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateStage(rowIndex, newStage) {
  try {
    getSheet().getRange(rowIndex, 5).setValue(newStage);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// REVENUE API
// ============================================================

function normalizeJenisProduk(raw) {
  if (!raw) return "Lainnya";
  var s = String(raw).trim();
  var lo = s.toLowerCase();
  if (lo === "b2s/oa" || lo === "b2s/open access") return "B2S/Open Access";
  if (lo === "wholesale" || lo === "wholesale product") return "Wholesale";
  if (lo === "media product") return "Media Product";
  if (lo === "not found" || s === "") return "Lainnya";
  return s;
}

function getRevenueData() {
  try {
    var allRevenue = [];
    
    // Ambil data dari REVENUE sheet biasa
    var sheet = getRevenueSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
      var MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      
      values.forEach(function(row, i) {
        if (!row.some(function(cell) { return cell !== "" && cell !== null; })) return;
        
        var monthRaw = row[8];
        var monthLabel = "";
        if (monthRaw instanceof Date) {
          monthLabel = MONTH_NAMES[monthRaw.getMonth()];
        } else if (monthRaw) {
          var s = String(monthRaw).trim();
          var parts = s.split(/[\s\/\-]+/);
          for (var p = 0; p < parts.length; p++) {
            var part = parts[p].replace(/[^a-zA-Z]/g, "");
            if (part.length === 3) {
              var idx = MONTH_NAMES.findIndex(function(m) {
                return m.toLowerCase() === part.toLowerCase();
              });
              if (idx >= 0) {
                monthLabel = MONTH_NAMES[idx];
                break;
              }
            }
          }
          if (!monthLabel) monthLabel = s;
        }
        
        allRevenue.push({
          rowIndex: i + 2,
          year: row[0] ? String(row[0]).trim() : "",
          sid: String(row[1] || "").trim(),
          isp: String(row[2] || "").trim(),
          serviceName: String(row[3] || "").trim(),
          jenisProduk: normalizeJenisProduk(row[4]),
          namaAM: String(row[5] || "").trim(),
          subscType: String(row[6] || "").trim(),
          amount: Number(row[7]) || 0,
          month: monthLabel,
          source: "Non XL" // Marker untuk source
        });
      });
    }
    
    // Ambil data dari REVENUE XLS sheet
    try {
      var sheetXLS = getRevenueXLSSheet();
      var lastRowXLS = sheetXLS.getLastRow();
      if (lastRowXLS >= 2) {
        var valuesXLS = sheetXLS.getRange(2, 1, lastRowXLS - 1, 9).getValues();
        
        valuesXLS.forEach(function(row, i) {
          if (!row.some(function(cell) { return cell !== "" && cell !== null; })) return;
          
          var monthRaw = row[8];
          var monthLabel = "";
          if (monthRaw instanceof Date) {
            monthLabel = MONTH_NAMES[monthRaw.getMonth()];
          } else if (monthRaw) {
            var s = String(monthRaw).trim();
            var parts = s.split(/[\s\/\-]+/);
            for (var p = 0; p < parts.length; p++) {
              var part = parts[p].replace(/[^a-zA-Z]/g, "");
              if (part.length === 3) {
                var idx = MONTH_NAMES.findIndex(function(m) {
                  return m.toLowerCase() === part.toLowerCase();
                });
                if (idx >= 0) {
                  monthLabel = MONTH_NAMES[idx];
                  break;
                }
              }
            }
            if (!monthLabel) monthLabel = s;
          }
          
          allRevenue.push({
            rowIndex: i + 2 + 100000, // Offset untuk membedakan dari sheet biasa
            year: row[0] ? String(row[0]).trim() : "",
            sid: String(row[1] || "").trim(),
            isp: String(row[2] || "").trim(),
            serviceName: String(row[3] || "").trim(),
            jenisProduk: normalizeJenisProduk(row[4]),
            namaAM: String(row[5] || "").trim(),
            subscType: String(row[6] || "").trim(),
            amount: Number(row[7]) || 0,
            month: monthLabel,
            source: "XL" // Marker untuk source XL
          });
        });
      }
    } catch (e) {
      Logger.log("Error loading REVENUE XLS: " + e.message);
      // Continue dengan data Non XL saja jika XLS error
    }
    
    return { success: true, data: allRevenue };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// PIPELINE CUTOFF API
// ============================================================

function getPipelineCutoffData() {
  try {
    var sheet = getPipelineCutoffSheet();
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow < 2 || lastColumn < 1) return { success: true, data: [] };

    var values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    var headers = values[0];
    var rows = values.slice(1);
    var headerMap = buildHeaderMap(headers);

    var data = rows
      .filter(function(row) {
        return row.some(function(cell) { return cell !== "" && cell !== null; });
      })
      .map(function(row, i) {
        var createdDate = getValueByHeader(row, headerMap, ["Opportunity Created Date"]);
        return {
          rowIndex: i + 2,
          account: String(getValueByHeader(row, headerMap, ["Account"]) || "").trim(),
          stage: String(getValueByHeader(row, headerMap, ["Stage"]) || "").trim(),
          opportunityType: String(getValueByHeader(row, headerMap, ["Opportunity Type"]) || "").trim(),
          type: String(getValueByHeader(row, headerMap, ["Type"]) || "").trim(),
          opportunityCreatedDate: createdDate instanceof Date && !isNaN(createdDate)
            ? Utilities.formatDate(createdDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
            : String(createdDate || "").trim(),
          totalFYRevenue: Number(getValueByHeader(row, headerMap, ["Total FY Revenue (converted)"])) || 0,
          contractMonth: Number(getValueByHeader(row, headerMap, ["Contract (Month)"])) || 0,
          ownerName: String(getValueByHeader(row, headerMap, ["Owner Name", "Owner: Full Name"]) || "").trim(),
          productName: String(getValueByHeader(row, headerMap, ["Product Name", "Quote Line Product Name"]) || "").trim(),
          opportunityName: String(getValueByHeader(row, headerMap, ["Opportunity Name", "Opportunity: Opportunity Name"]) || "").trim()
        };
      });

    return { success: true, data: data };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// CALENDAR API
// ============================================================

function getCalendarData() {
  try {
    var sheet = getCalendarSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var data = values
      .map(function(row, index) {
        return {
          row: row,
          rowIndex: index + 2
        };
      })
      .filter(function(entry) {
        return entry.row.some(function(cell) { return cell !== "" && cell !== null; });
      })
      .map(function(entry) {
        var row = entry.row;
        var dateRaw = row[0];
        var dateStr = "";
        if (dateRaw instanceof Date && !isNaN(dateRaw)) {
          dateStr = Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (dateRaw) {
          dateStr = String(dateRaw).trim();
        }
        var docs = parseCalendarDocs_(row[7]);
        return {
          rowIndex: entry.rowIndex,
          date: dateStr,
          year: row[1] ? String(row[1]).trim() : "",
          company: String(row[2] || "").trim(),
          isp: String(row[3] || "").trim(),
          activity: String(row[4] || "").trim(),
          location: String(row[5] || "").trim(),
          pic: String(row[6] || "").trim(),
          docs: docs,
          docsText: docs.map(function(doc) {
            return [doc.fileName, doc.fileUrl, doc.deskripsi].filter(Boolean).join(" ");
          }).join(" "),
          docsCount: docs.length
        };
      });

    return { success: true, data: data };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addCalendarEvent(d) {
  try {
    var sheet = getCalendarSheet();
    var dateObj = d.date ? new Date(d.date) : new Date();
    var year = dateObj.getFullYear();
    var newRow = [
      dateObj,
      year,
      d.company || "",
      d.isp || "",
      d.activity || "",
      d.location || "",
      d.pic || "",
      serializeCalendarDocs_(d.docs)
    ];
    var lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, 8).setValues([newRow]);
    return { success: true, rowIndex: lastRow };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateCalendarEvent(d) {
  try {
    var sheet = getCalendarSheet();
    var dateObj = d.date ? new Date(d.date) : new Date();
    var year = dateObj.getFullYear();
    var updatedRow = [
      dateObj,
      year,
      d.company || "",
      d.isp || "",
      d.activity || "",
      d.location || "",
      d.pic || "",
      serializeCalendarDocs_(d.docs)
    ];
    sheet.getRange(d.rowIndex, 1, 1, 8).setValues([updatedRow]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteCalendarEvent(rowIndex) {
  try {
    getCalendarSheet().deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function parseCalendarDocs_(rawValue) {
  if (rawValue === "" || rawValue === null || rawValue === undefined) return [];
  var value = String(rawValue).trim();
  if (!value) return [];

  try {
    var parsed = JSON.parse(value);
    var list = Array.isArray(parsed) ? parsed : [parsed];
    return list
      .map(function(item) { return normalizeCalendarDoc_(item); })
      .filter(function(item) { return item.fileName || item.fileUrl; });
  } catch (err) {
    return [normalizeCalendarDoc_(value)];
  }
}

function normalizeCalendarDoc_(item) {
  if (typeof item === "string") {
    var label = item;
    if (/^https?:\/\//i.test(item)) {
      var parts = item.split("/");
      label = decodeURIComponent(parts[parts.length - 1] || "Dokumen");
    }
    return {
      fileName: label || "Dokumen",
      fileUrl: item,
      fileId: "",
      fileDate: "",
      dateUploaded: "",
      jenisFile: /^https?:\/\//i.test(item) ? "Link" : "Catatan",
      deskripsi: "",
      previewUrl: /^https?:\/\//i.test(item) ? item : "",
      downloadUrl: /^https?:\/\//i.test(item) ? item : ""
    };
  }

  item = item || {};
  var fileId = String(item.fileId || "").trim();
  var fileUrl = String(item.fileUrl || item.url || "").trim();
  return {
    fileName: String(item.fileName || item.name || "Dokumen").trim(),
    fileUrl: fileUrl,
    fileId: fileId,
    fileDate: String(item.fileDate || "").trim(),
    dateUploaded: String(item.dateUploaded || "").trim(),
    jenisFile: String(item.jenisFile || "Other").trim(),
    deskripsi: String(item.deskripsi || "").trim(),
    previewUrl: fileUrl || (fileId ? "https://drive.google.com/file/d/" + encodeURIComponent(fileId) + "/view" : ""),
    downloadUrl: fileId ? "https://drive.google.com/uc?export=download&id=" + encodeURIComponent(fileId) : fileUrl
  };
}

function serializeCalendarDocs_(docs) {
  if (docs === "" || docs === null || docs === undefined) return "";
  if (typeof docs === "string") {
    var parsed = parseCalendarDocs_(docs);
    return parsed.length ? JSON.stringify(parsed.map(stripCalendarDocForStorage_)) : String(docs).trim();
  }

  var list = Array.isArray(docs) ? docs : [docs];
  list = list
    .map(function(item) { return normalizeCalendarDoc_(item); })
    .filter(function(item) { return item.fileName || item.fileUrl; })
    .map(stripCalendarDocForStorage_);
  return list.length ? JSON.stringify(list) : "";
}

function stripCalendarDocForStorage_(doc) {
  return {
    fileName: doc.fileName || "",
    fileUrl: doc.fileUrl || "",
    fileId: doc.fileId || "",
    fileDate: doc.fileDate || "",
    dateUploaded: doc.dateUploaded || "",
    jenisFile: doc.jenisFile || "",
    deskripsi: doc.deskripsi || ""
  };
}

// ============================================================
// ORDER DETAIL API
// ============================================================

function getOrderDetailData() {
  try {
    var sheet = getOrderDetailSheet();
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow < 2 || lastColumn < 1) return { success: true, headers: [], data: [] };

    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    var selectedColumns = [
      "PO Number",
      "Order Created Date",
      "Customer ID",
      "Account Name",
      "Order Owner",
      "Sales Head",
      "Location 1",
      "Location 2",
      "Order Status",
      "Service ID#",
      "Product Name",
      "International Capacity",
      "Local Capacity",
      "Product Unit",
      "Product Sub Family",
      "Linknet Departement",
      "Subscript Type#",
      "Payment Term",
      "Order Currency",
      "Amount Calculate",
      "WO Status",
      "Order Type",
      "Reason For Termination",
      "Order Number",
      "Order Product",
    ];
    var headerMap = buildHeaderMap(headers);

    var data = rows
      .map(function(row, index) {
        return {
          row: row,
          rowIndex: index + 2
        };
      })
      .filter(function(entry) {
        return entry.row.some(function(cell) { return cell !== "" && cell !== null; });
      })
      .map(function(entry) {
        var row = entry.row;
        var detail = {};
        headers.forEach(function(header, colIndex) {
          detail[header] = formatOrderDetailCell_(row[colIndex]);
        });

        var orderCurrency = getValueByHeader(row, headerMap, ["Order Currency"]);
        var amountCalculate = getValueByHeader(row, headerMap, ["Amount Calculate"]);

        var summary = {
          rowIndex: entry.rowIndex,
          actionLabel: "Aksi"
        };

        selectedColumns.forEach(function(name) {
          summary[name] = formatOrderDetailCell_(getValueByHeader(row, headerMap, [name]));
        });

        summary["Amount Display"] = formatOrderDetailAmount_(orderCurrency, amountCalculate);

        return {
          rowIndex: entry.rowIndex,
          summary: summary,
          detail: detail
        };
      });

    return {
      success: true,
      headers: headers,
      selectedColumns: selectedColumns,
      data: data
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getOrderDetailRow(rowIndex) {
  try {
    var targetRow = Number(rowIndex || 0);
    if (targetRow < 2) throw new Error("Baris order detail tidak valid.");

    var sheet = getOrderDetailSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var values = sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0];
    var detail = {};

    headers.forEach(function(header, index) {
      detail[header] = formatOrderDetailCell_(values[index]);
    });

    return { success: true, rowIndex: targetRow, headers: headers, detail: detail };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateOrderDetailRow(payload) {
  try {
    payload = payload || {};
    var rowIndex = Number(payload.rowIndex || 0);
    if (rowIndex < 2) throw new Error("Baris order detail tidak valid.");

    var sheet = getOrderDetailSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var current = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    var updated = current.slice();
    var detail = payload.detail || {};

    headers.forEach(function(header, index) {
      if (Object.prototype.hasOwnProperty.call(detail, header)) {
        updated[index] = parseOrderDetailInput_(detail[header], current[index]);
      }
    });

    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([updated]);
    return { success: true, message: "Order Detail berhasil diperbarui." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteOrderDetailRow(rowIndex) {
  try {
    var targetRow = Number(rowIndex || 0);
    if (targetRow < 2) throw new Error("Baris order detail tidak valid.");
    getOrderDetailSheet().deleteRow(targetRow);
    return { success: true, message: "Order Detail berhasil dihapus." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function formatOrderDetailAmount_(currency, amount) {
  var curr = String(currency || "").trim().toUpperCase();
  var raw = amount;
  var num = Number(raw);
  if (raw === "" || raw === null || isNaN(num)) return "";
  if (curr === "IDR") return "Rp. " + num.toLocaleString("id-ID");
  if (!curr) return num.toLocaleString("id-ID");
  return curr + " " + num.toLocaleString("id-ID");
}

function formatOrderDetailCell_(value) {
  if (value === null || value === undefined) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value);
}

function parseOrderDetailInput_(value, existingValue) {
  if (Object.prototype.toString.call(existingValue) === "[object Date]" && !isNaN(existingValue.getTime())) {
    if (!value) return "";
    return new Date(value);
  }

  if (typeof existingValue === "number") {
    if (value === "" || value === null || value === undefined) return "";
    var numeric = Number(String(value).replace(/,/g, ""));
    return isNaN(numeric) ? value : numeric;
  }

  return value;
}

// ============================================================
// DOCUMENT LIBRARY API
// ============================================================

function getDocumentLibraryInitialData() {
  getOrCreateDocumentSheet_();
  return {
    folders: getDocumentFolderSummaries(),
    options: buildDocumentOptions_()
  };
}

function setupDocumentLibrarySheet() {
  getOrCreateDocumentSheet_();
  return {
    success: true,
    message: "Sheet document library siap digunakan.",
    folders: getDocumentFolderSummaries(),
    options: buildDocumentOptions_()
  };
}

function getDocumentFolderSummaries() {
  var rows = getDocumentLibraryRows_();
  var grouped = {};

  rows.forEach(function(row) {
    var folderId = String(row.folderId || "").trim();
    if (!folderId) return;

    if (!grouped[folderId]) {
      grouped[folderId] = {
        folderId: folderId,
        isp: row.isp || "",
        cities: row.cities || "",
        dateUploaded: row.dateUploaded || "",
        deskripsi: row.deskripsi || "",
        status: row.status || "",
        owner: row.owner || "",
        fileCount: 0,
        jenisFileSet: {}
      };
    }

    var folder = grouped[folderId];
    folder.fileCount += row.fileName || row.fileId || row.fileUrl ? 1 : 0;
    if (!folder.isp && row.isp) folder.isp = row.isp;
    if (!folder.cities && row.cities) folder.cities = row.cities;
    if (!folder.dateUploaded && row.dateUploaded) folder.dateUploaded = row.dateUploaded;
    if (!folder.deskripsi && row.deskripsi) folder.deskripsi = row.deskripsi;
    if (!folder.status && row.status) folder.status = row.status;
    if (!folder.owner && row.owner) folder.owner = row.owner;
    if (row.jenisFile) folder.jenisFileSet[String(row.jenisFile).trim()] = true;
  });

  return Object.keys(grouped).sort().map(function(folderId) {
    var folder = grouped[folderId];
    folder.jenisFiles = Object.keys(folder.jenisFileSet);
    delete folder.jenisFileSet;
    return folder;
  });
}

function getDocumentFolderDetail(folderId) {
  var targetId = String(folderId || "").trim();
  if (!targetId) throw new Error("Folder ID tidak ditemukan.");

  var rows = getDocumentLibraryRows_().filter(function(row) {
    return String(row.folderId || "").trim() === targetId;
  });

  if (!rows.length) throw new Error("Folder tidak ditemukan.");

  var first = rows[0];
  return {
    folder: {
      folderId: targetId,
      isp: first.isp || "",
      cities: first.cities || "",
      dateUploaded: first.dateUploaded || "",
      deskripsi: first.deskripsi || "",
      status: first.status || "",
      owner: first.owner || "",
      fileCount: rows.filter(function(row) {
        return row.fileName || row.fileId || row.fileUrl;
      }).length,
      jenisFiles: uniqueDocumentValues_(rows.map(function(row) { return row.jenisFile; }))
    },
    files: rows
  };
}

function saveDocumentFolder(payload) {
  payload = payload || {};

  var sheet = getOrCreateDocumentSheet_();
  var folderId = String(payload.folderId || "").trim() || generateNextDocumentFolderId_();
  var shared = normalizeDocumentFolderPayload_(payload);
  var items = payload.files && payload.files.length ? payload.files : [createBlankDocumentFilePayload_()];
  var existingRows = findDocumentFolderRows_(folderId);

  if (existingRows.length) {
    existingRows.forEach(function(item) {
      var existing = mapDocumentRowToObject_(item.values, item.rowNumber);
      var merged = buildDocumentRowFromFolderAndFile_(folderId, shared, existing);
      sheet.getRange(item.rowNumber, 1, 1, DOC_HEADERS.length).setValues([merged]);
    });
  }

  items.forEach(function(fileItem) {
    if (Number(fileItem.rowNumber || 0) > 1) return;
    if (existingRows.length && items.length === 1 && isBlankDocumentFilePayload_(fileItem)) return;
    sheet.appendRow(buildDocumentRowFromFolderAndFile_(folderId, shared, normalizeDocumentFilePayload_(fileItem)));
  });

  return {
    success: true,
    folderId: folderId,
    message: existingRows.length ? "Folder berhasil diupdate." : "Folder berhasil dibuat.",
    folders: getDocumentFolderSummaries(),
    options: buildDocumentOptions_(),
    detail: getDocumentFolderDetail(folderId)
  };
}

function addFilesToDocumentFolder(payload) {
  payload = payload || {};
  var folderId = String(payload.folderId || "").trim();
  if (!folderId) throw new Error("Folder ID wajib ada.");

  var existingRows = findDocumentFolderRows_(folderId);
  if (!existingRows.length) throw new Error("Folder tidak ditemukan.");

  var baseRow = mapDocumentRowToObject_(existingRows[0].values, existingRows[0].rowNumber);
  var shared = normalizeDocumentFolderPayload_(baseRow);
  var files = payload.files && payload.files.length ? payload.files : [];
  if (!files.length) throw new Error("File baru belum ada.");

  var sheet = getOrCreateDocumentSheet_();
  files.forEach(function(fileItem) {
    sheet.appendRow(buildDocumentRowFromFolderAndFile_(folderId, shared, normalizeDocumentFilePayload_(fileItem)));
  });

  return {
    success: true,
    message: files.length + " file berhasil ditambahkan ke folder.",
    folders: getDocumentFolderSummaries(),
    options: buildDocumentOptions_(),
    detail: getDocumentFolderDetail(folderId)
  };
}

function updateDocumentFile(payload) {
  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  if (rowNumber < 2) throw new Error("Baris file tidak valid.");

  var sheet = getOrCreateDocumentSheet_();
  var existingValues = sheet.getRange(rowNumber, 1, 1, DOC_HEADERS.length).getValues()[0];
  var existing = mapDocumentRowToObject_(existingValues, rowNumber);
  var shared = normalizeDocumentFolderPayload_(existing);
  var mergedValues = buildDocumentRowFromFolderAndFile_(existing.folderId, shared, normalizeDocumentFilePayload_(payload));
  sheet.getRange(rowNumber, 1, 1, DOC_HEADERS.length).setValues([mergedValues]);

  return {
    success: true,
    message: "File berhasil diupdate.",
    folders: getDocumentFolderSummaries(),
    options: buildDocumentOptions_(),
    detail: getDocumentFolderDetail(existing.folderId)
  };
}

function deleteDocumentFile(rowNumber) {
  var targetRow = Number(rowNumber || 0);
  if (targetRow < 2) throw new Error("Baris file tidak valid.");

  var sheet = getOrCreateDocumentSheet_();
  var currentValues = sheet.getRange(targetRow, 1, 1, DOC_HEADERS.length).getValues()[0];
  var current = mapDocumentRowToObject_(currentValues, targetRow);
  var folderId = current.folderId;
  sheet.deleteRow(targetRow);

  var remaining = findDocumentFolderRows_(folderId);
  return {
    success: true,
    folderRemoved: remaining.length === 0,
    message: "File berhasil dihapus.",
    folders: getDocumentFolderSummaries(),
    options: buildDocumentOptions_(),
    detail: remaining.length ? getDocumentFolderDetail(folderId) : null
  };
}

function uploadDocumentFile(payload) {
  if (!payload) throw new Error("Payload file tidak ditemukan.");

  var fileName = String(payload.fileName || "").trim();
  var mimeType = String(payload.mimeType || "application/octet-stream").trim();
  var base64Data = String(payload.base64Data || "").trim();

  if (!fileName) throw new Error("Nama file wajib ada.");
  if (!base64Data) throw new Error("Isi file tidak ditemukan.");

  var file = getDocumentUploadFolder_().createFile(
    Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName)
  );

  return {
    success: true,
    fileName: file.getName(),
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + encodeURIComponent(file.getId()),
    previewUrl: file.getUrl(),
    mimeType: mimeType
  };
}

function getDocumentLibraryRows_() {
  var sheet = getOrCreateDocumentSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, DOC_HEADERS.length).getValues()
    .filter(function(row) {
      return row.some(function(cell) {
        return String(cell || "").trim() !== "";
      });
    })
    .map(function(row, index) {
      return mapDocumentRowToObject_(row, index + 2);
    });
}

function getOrCreateDocumentSheet_() {
  var ss = getDocumentSpreadsheet_();
  var sheet = ss.getSheetByName(DOC_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(DOC_SHEET_NAME);

  var headers = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, DOC_HEADERS.length).getValues()[0]
    : [];

  var needsHeader = DOC_HEADERS.some(function(header, index) {
    return String(headers[index] || "").trim() !== header;
  });

  if (needsHeader) {
    sheet.getRange(1, 1, 1, DOC_HEADERS.length).setValues([DOC_HEADERS]);
    sheet.getRange(1, 1, 1, DOC_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, DOC_HEADERS.length);
  }

  return sheet;
}

function normalizeDocumentFolderPayload_(payload) {
  return {
    isp: String(payload.isp || "").trim(),
    cities: String(payload.cities || "").trim(),
    dateUploaded: normalizeDocumentDateInput_(payload.dateUploaded),
    deskripsi: String(payload.deskripsi || "").trim(),
    status: String(payload.status || "").trim(),
    owner: String(payload.owner || "").trim()
  };
}

function normalizeDocumentFilePayload_(payload) {
  return {
    rowNumber: Number(payload.rowNumber || 0),
    fileName: String(payload.fileName || "").trim(),
    fileUrl: String(payload.fileUrl || "").trim(),
    fileId: String(payload.fileId || "").trim(),
    fileDate: normalizeDocumentDateInput_(payload.fileDate),
    jenisFile: String(payload.jenisFile || "").trim()
  };
}

function buildDocumentRowFromFolderAndFile_(folderId, shared, file) {
  return [
    folderId,
    shared.isp || "",
    shared.cities || "",
    file.fileName || "",
    file.fileUrl || "",
    file.fileId || "",
    file.fileDate || "",
    shared.dateUploaded || "",
    shared.deskripsi || "",
    file.jenisFile || "",
    shared.status || "",
    shared.owner || ""
  ];
}

function createBlankDocumentFilePayload_() {
  return { rowNumber: 0, fileName: "", fileUrl: "", fileId: "", fileDate: "", jenisFile: "" };
}

function isBlankDocumentFilePayload_(payload) {
  return !String(payload.fileName || "").trim() &&
    !String(payload.fileUrl || "").trim() &&
    !String(payload.fileId || "").trim() &&
    !String(payload.fileDate || "").trim() &&
    !String(payload.jenisFile || "").trim();
}

function findDocumentFolderRows_(folderId) {
  var target = String(folderId || "").trim();
  if (!target) return [];

  var sheet = getOrCreateDocumentSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, DOC_HEADERS.length).getValues();
  var matches = [];

  values.forEach(function(row, index) {
    if (String(row[0] || "").trim() === target) {
      matches.push({ rowNumber: index + 2, values: row });
    }
  });

  return matches;
}

function generateNextDocumentFolderId_() {
  var rows = getDocumentLibraryRows_();
  var max = 0;

  rows.forEach(function(row) {
    var match = String(row.folderId || "").match(/FLD-(\d+)/i);
    if (!match) return;
    var num = Number(match[1]);
    if (num > max) max = num;
  });

  return "FLD-" + Utilities.formatString("%04d", max + 1);
}

function getDocumentUploadFolder_() {
  var folderName = "Document Archive Uploads";
  var ss = getDocumentSpreadsheet_();

  try {
    var file = DriveApp.getFileById(ss.getId());
    var parents = file.getParents();
    if (parents.hasNext()) {
      var parentFolder = parents.next();
      var folders = parentFolder.getFoldersByName(folderName);
      if (folders.hasNext()) return folders.next();
      try {
        return parentFolder.createFolder(folderName);
      } catch (error) {
        return getOrCreateRootDocumentUploadFolder_(folderName);
      }
    }
  } catch (error) {
    return getOrCreateRootDocumentUploadFolder_(folderName);
  }

  return getOrCreateRootDocumentUploadFolder_(folderName);
}

function getOrCreateRootDocumentUploadFolder_(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function buildDocumentOptions_() {
  var rows = getDocumentLibraryRows_();
  return {
    status: uniqueDocumentValues_(DOC_STATUS_OPTIONS.concat(rows.map(function(row) { return row.status; }))),
    jenisFile: uniqueDocumentValues_(DOC_FILE_TYPE_OPTIONS.concat(rows.map(function(row) { return row.jenisFile; }))),
    isp: uniqueDocumentValues_(rows.map(function(row) { return row.isp; })),
    cities: uniqueDocumentValues_(rows.map(function(row) { return row.cities; })),
    owner: uniqueDocumentValues_(rows.map(function(row) { return row.owner; }))
  };
}

function mapDocumentRowToObject_(row, rowNumber) {
  var fileUrl = String(row[4] || "").trim();
  var fileId = String(row[5] || "").trim();
  return {
    rowNumber: rowNumber,
    folderId: row[0],
    isp: row[1],
    cities: row[2],
    fileName: row[3],
    fileUrl: fileUrl,
    fileId: fileId,
    fileDate: formatDocumentOutputDate_(row[6]),
    dateUploaded: formatDocumentOutputDate_(row[7]),
    deskripsi: row[8],
    jenisFile: row[9],
    status: row[10],
    owner: row[11],
    previewUrl: fileUrl || (fileId ? "https://drive.google.com/file/d/" + encodeURIComponent(fileId) + "/view" : ""),
    downloadUrl: fileId ? "https://drive.google.com/uc?export=download&id=" + encodeURIComponent(fileId) : fileUrl
  };
}

function uniqueDocumentValues_(values) {
  var map = {};
  return (values || []).filter(function(value) {
    var text = String(value || "").trim();
    var key = text.toLowerCase();
    if (!text || map[key]) return false;
    map[key] = true;
    return true;
  });
}

function normalizeDocumentDateInput_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value).trim();
}

function formatDocumentOutputDate_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value);
}

// ============================================================
// ADD THIS TO YOUR EXISTING backend.gs
// Sheet name: ganti sesuai nama sheet spreadsheet event kamu
// ============================================================

const SHEET_EVENT_MARKETING = "Event Data"; // <-- ganti dengan nama sheet kamu
const EVENT_SPREADSHEET_ID  = "1m0tiOB6MWexUC4Bx-Slfn7u3PMldvnu4B4WFBnT2i0A"; // <-- ganti

function getEventSpreadsheet_() {
  // Jika event data ada di spreadsheet YANG SAMA dengan CBD App:
  // return getSalesSpreadsheet_();
  // Jika beda spreadsheet, pakai ini:
  return SpreadsheetApp.openById(EVENT_SPREADSHEET_ID);
}

function getEventData() {
  try {
    var ss    = getEventSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_EVENT_MARKETING);
    if (!sheet) throw new Error("Sheet '" + SHEET_EVENT_MARKETING + "' tidak ditemukan.");

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [], totals: { cost: 0, revenue: 0 } };

    // Baca semua data mulai row 2 (skip header)
    var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    var totalCost    = 0;
    var totalRevenue = 0;
    var data = [];

    values.forEach(function(row, i) {
      // Skip baris kosong atau baris TOTAL
      var year = row[0];
      if (!year || String(year).toUpperCase() === "TOTAL") return;

      var cost    = Number(String(row[7]).replace(/[^0-9.-]/g, "")) || 0;
      var revenue = Number(String(row[8]).replace(/[^0-9.-]/g, "")) || 0;

      totalCost    += cost;
      totalRevenue += revenue;

      data.push({
        rowIndex : i + 2,
        year     : Number(year) || 0,
        month    : String(row[1] || "").trim(),
        category : String(row[2] || "").trim(),
        isp      : String(row[3] || "").trim(),
        brand    : String(row[4] || "").trim(),
        product  : String(row[5] || "").trim(),   // B2S / Wholesale / Open Access
        status   : String(row[6] || "").trim(),
        cost     : cost,
        revenue  : revenue
      });
    });

    return {
      success : true,
      data    : data,
      totals  : { cost: totalCost, revenue: totalRevenue }
    };

  } catch (e) {
    return { success: false, error: e.message };
  }
}