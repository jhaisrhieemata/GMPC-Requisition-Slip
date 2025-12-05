// === CONFIG ===
const TEMPLATE_ID = "1bss96EkFlNl_GQ35E8PdxeZrgO-02yrJILtevKzE4SE"; // Google Doc template ID
const FOLDER_ID = "1p-nYwiyl6XsXXV93_nJf5OoeXXmaWZWL"; // Folder for PDFs
const SHEET_FILE_ID = "1MVY1ucbqCTRQkoEEMaQc6tEI6u62psbup6iL023xGsI"; // Spreadsheet ID
const REDIRECT_URL = "https://sites.google.com/view/giantmotoprocorp";

// === MASTER ITEM LIST ===
const ITEM_LIST = {
  ITM001: "BOND PAPER (LONG)",
  ITM002: "BOND PAPER (SHORT)",
  ITM003: "BOND PAPER A4",
  ITM004: "BOND PAPER NEWSPRINT (SHORT)",
  ITM005: "BOND PAPER NEWSPRINT (LONG)",
  ITM006: "VELLUM PAPER",
  ITM007: "LOGBOOK VALIANT RECORD (150 & 300P PAGES)",
  ITM008: "NOTEBOOK",
  ITM009: "CARBON PAPER (LONG BLUE)",
  ITM010: "BROWN ENVELOPE (LONG)",
  ITM011: "BROWN ENVELOPE (SHORT)",
  ITM012: "BROWN FOLDER (LONG)",
  ITM013: "BROWN FOLDER (SHORT)",
  ITM014: "WHITE FOLDER (LONG)",
  ITM015: "WHITE FOLDER (SHORT)",
  ITM016: "LETTER SHORT WHITE ENVELOPE",
  ITM017: "LETTER LONG WHITE ENVELOPE",
  ITM018: "CARD HOLDER",
  ITM019: "FLEXSTICK BLACK BALLPEN (0.5)",
  ITM020: "FLEXSTICK RED BALLPEN (0.5)",
  ITM021: "WYTEBORD MARKER PILOT (BLACK)",
  ITM022: "WYTEBORD MARKER PILOT (RED)",
  ITM023: "HIGHLIGHTER",
  ITM024: "PENCIL",
  ITM025: "SHARPENER",
  ITM026: "ERASER",
  ITM027: "CORRECTION TAPE",
  ITM028: "CANDIES",
  ITM029: "PUNCHER",
  ITM030: "STAPLER",
  ITM031: "SCISSORS",
  ITM032: "GLUE",
  ITM033: "RULER",
  ITM034: "CALCULATOR MX-12B",
  ITM035: "TAPE DISPENSER BIG",
  ITM036: "SCOTCH TAPE (1 INCH)",
  ITM037: "SCOTCH TAPE (2 INCH)",
  ITM038: "MASKING TAPE",
  ITM039: "PACKAGING TAPE BROWN",
  ITM040: "DOUBLE-SIDED TAPE (1 INCH)",
  ITM041: "DUCT TAPE",
  ITM043: "TYPEWRITER BLACK NYLON",
  ITM044: "INK EPSON 003 (YELLOW)",
  ITM045: "INK EPSON 003 (MAGENTA)",
  ITM046: "INK EPSON 003 (CYAN)",
  ITM047: "INK EPSON 003 (BLACK)",
  ITM048: "INK BROTHER BT5000Y",
  ITM049: "INK BROTHER BT5000M",
  ITM050: "INK BROTHER BT5000C",
  ITM051: "INK BROTHER BTD60BK",
  ITM052: "INK HP (YELLOW)",
  ITM053: "INK HP (MAGENTA)",
  ITM054: "INK HP (CYAN)",
  ITM055: "INK HP (BLACK)",
  ITM056: "PRIDE SOAP POWDER",
  ITM057: "STAMP PAD INK BLACK 30ML",
  ITM058: "STAMP PAD INK BLUE 30ML",
  ITM059: "STAMP PAD INK RED 30ML",
  ITM060: "SUPER COLOR MARKER PILOT (BLACK)",
  ITM061: "JOY",
  ITM062: "DOWNY",
  ITM063: "SAFEGUARD SOAP",
  ITM064: "STICKY NOTES 2X2 A02",
  ITM065: "STICKY NOTES A03",
  ITM066: "STICKY NOTES A05",
  ITM067: "FASTENER",
  ITM068: "PAPER CLIPS 50MM",
  ITM069: "PAPER CLIPS JUMBO",
  ITM070: "BINDER CLIPS 2 INCHES",
  ITM071: "BINDER CLIPS 1 5/8 41MM",
  ITM072: "BINDER CLIPS 1 1/4 32MM",
  ITM073: "BINDER CLIPS 1 INCHES",
  ITM074: "STAPLE WIRE #35",
  ITM075: "STAPLE WIRE #10-M",
  ITM076: "RUBBER BAND ROUND RB-350",
  ITM077: "CABLE TIE",
  ITM078: "COFFEE CUPS",
  ITM079: "DRINKING CUPS",
  ITM080: "COFFEE STICK",
  ITM081: "CREAMER",
  ITM082: "BROWN SUGAR",
  ITM083: "STIRRER",
  ITM084: "SUPER COLOR MARKER REFILL INK (BLACK)",
  ITM085: "WYTEBORD MARKER REFILL INK PILOT (BLACK)",
  ITM086: "WYTEBORD MARKER REFILL INK PILOT (RED)",
  ITM087: "ALBATROSS",
  ITM088: "ZONROX (ORIGINAL)",
  ITM089: "ZONROX (COLORSAFE)",
  ITM090: "GLASS CLEANER",
  ITM091: "HANDWASH",
  ITM092: "RAGS",
  ITM093: "TRASH BAG (XXL)",
  ITM094: "TRASH BAG (XL)",
  ITM095: "TRASH BAG (LARGE)",
  ITM096: "TRASH BAG (MEDIUM)",
  ITM097: "TRASH BAG (SMALL)",
  ITM098: "TISSUE PAPER",
  ITM099: "INK EPSON SET",
  ITM100: "INK BROTHER SET",
  ITM101: "INK HP SET",
  ITM102: "STAMP PAD BLUE NO.2",
  ITM103: "STAMP PAD BLACK NO.2",
  ITM104: "DTR CARD",
  ITM105: "FASTENER LONG",
};

// === FRONTEND LOADER ===
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Requisition Form v15")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// === MAIN SUBMISSION HANDLER (Optimized for Speed) ===
function saveAndCreatePdf(data) {
  if (!data || typeof data !== "object") throw new Error("Invalid data.");
  if (!Array.isArray(data.items)) data.items = [];

  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sheetName = getSheetNameByPurpose(data.purpose, data.branch);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // Create header only if new sheet
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Timestamp",
      "Branch",
      "Date",
      "To",
      "Purpose",
      "ITEM_ID",
      "Qty",
      "Unit",
      "Description",
      "UPrice",
      "Amount",
      "Total",
      "Requested By",
      "Status",
      "Release Date",
      "Received By",
      "PDF URL",
    ]);
  }

  const ts = new Date();
  const pdfUrl = createPdfFromTemplate(data, ts);

  // Prepare fast bulk append
  const newRows = data.items.map((it) => {
    const itemId = findItemId(it.description);
    return [
      ts,
      data.branch,
      data.date,
      data.to,
      data.purpose,
      itemId,
      it.qty,
      it.unit,
      it.description,
      it.uprice,
      it.amount,
      data.total,
      data.requested_by,
      "Pending",
      "",
      "",
      pdfUrl,
    ];
  });

  // Write all items in one call (faster)
  if (newRows.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }

  // Try to schedule SUMMARY generation (non-blocking)
  try {
    ScriptApp.newTrigger("generateMasterSummary")
      .timeBased()
      .after(1000)
      .create();
  } catch (e) {
    // ignore trigger creation errors
  }

  return pdfUrl;
}

// === GET SHEET NAME BY PURPOSE/BRANCH ===
function getSheetNameByPurpose(purpose, branch) {
  const p = purpose ? purpose.toUpperCase() : "";
  if (p === "SPECIAL REQUEST") return "SPECIAL REQUEST";
  return branch ? branch.trim().toUpperCase() : "GENERAL";
}

// === FIND ITEM ID ===
function findItemId(description) {
  if (!description) return "";
  const d = description.toString().trim().toUpperCase();
  for (let id in ITEM_LIST) {
    if (ITEM_LIST[id].toUpperCase() === d) return id;
  }
  return "";
}

// === FAST PDF CREATION (no external API, lightweight) ===
function createPdfFromTemplate(data, timestamp) {
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const copy = templateFile.makeCopy(
    `Requisition ${Utilities.formatDate(
      timestamp,
      "Asia/Manila",
      "yyyy-MM-dd HH:mm:ss"
    )}`
  );

  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  try {
    body.replaceText("{{BRANCH}}", data.branch || "");
    body.replaceText("{{DATE}}", data.date || "");
    body.replaceText("{{TO}}", data.to || "");
    body.replaceText("{{PURPOSE}}", data.purpose || "");
    body.replaceText("{{TOTAL}}", data.total || "");
    body.replaceText("{{NOTES}}", data.note || "");
    body.replaceText("{{REQUESTED_BY}}", data.requested_by || "");

    // Insert items table
    const search = body.findText("{{ITEMS}}");
    if (search) {
      const element = search.getElement();
      const parent = element.getParent();

      const tableData = [["Qty", "Unit", "Description", "UPrice", "Amount"]];
      (data.items || []).forEach((it) => {
        tableData.push([
          it.qty || "",
          it.unit || "",
          it.description || "",
          it.uprice || "",
          it.amount || "",
        ]);
      });
      tableData.push(["", "", "", "Total", data.total || ""]);

      const table = body.insertTable(body.getChildIndex(parent) + 1, tableData);
      table.setBorderWidth(1);
      table.setFontSize(9);

      for (let i = 0; i < table.getNumRows(); i++) {
        const row = table.getRow(i);
        for (let j = 0; j < row.getNumCells(); j++) {
          const cell = row.getCell(j);
          cell
            .setPaddingTop(3)
            .setPaddingBottom(3)
            .setPaddingLeft(5)
            .setPaddingRight(5);
          try {
            if (j === 0) cell.setWidth(40);
            if (j === 1) cell.setWidth(60);
            if (j === 2) cell.setWidth(240);
            if (j === 3) cell.setWidth(80);
            if (j === 4) cell.setWidth(80);
          } catch (e) {}
          const text = cell.getChild(0);
          if (
            text &&
            text.getType &&
            text.getType() === DocumentApp.ElementType.PARAGRAPH
          ) {
            const para = text.asParagraph();
            if (i === 0) {
              para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
              para.setBold(true);
              try {
                cell.setBackgroundColor("#f0f0f0");
              } catch (e) {}
            } else if ([0, 1, 3, 4].includes(j)) {
              para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            } else {
              para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
            }
          }
        }
      }
      element.asText().setText("");
    }

    // Signature insertion
    if (data.requested_by_signature) {
      const parts = data.requested_by_signature.split(",");
      const base64Signature = parts.length > 1 ? parts[1] : parts[0];
      const base64Data = base64Signature.trim();
      if (base64Data) {
        const imgBytes = Utilities.base64Decode(
          base64Data.replace(/^data:image\/png;base64,/, "")
        );
        const sigBlob = Utilities.newBlob(
          imgBytes,
          "image/png",
          "signature.png"
        );
        const searchSig = body.findText("{{REQUESTED_BY_SIGNATURE}}");
        if (searchSig) {
          const el = searchSig.getElement();
          const parent = el.getParent();
          const insertIndex = parent.getChildIndex(el);
          el.asText().setText("");
          const spacerBlob = Utilities.newBlob(
            Utilities.base64Decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABAQMAAAAl21bKAAAAA1BMVEUAAACnej3aAAAAAXRSTlMAQObYZgAAAApJREFUCNdjYAAAAAIAAeIhvDMAAAAASUVORK5CYII="
            ),
            "image/png",
            "spacer.png"
          );
          parent
            .insertInlineImage(insertIndex, spacerBlob)
            .setWidth(95)
            .setHeight(8);
          parent
            .insertInlineImage(insertIndex + 1, sigBlob)
            .setWidth(150)
            .setHeight(55);
        }
      }
    }

    doc.saveAndClose();
  } catch (e) {
    Logger.log("PDF generation error: " + e);
  }

  const pdfBlob = copy.getBlob().getAs("application/pdf");
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const pdfFile = folder.createFile(pdfBlob).setName(copy.getName() + ".pdf");
  pdfFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  // Trash temp doc
  DriveApp.getFileById(copy.getId()).setTrashed(true);

  return pdfFile.getUrl();
}
// === REAL-TIME STOCK FETCHER (based on REAL-TIME STOCKS SHEET) ===
function getRunningStocks() {
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sh = ss.getSheetByName("REAL-TIME STOCKS");
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  // Columns:
  // A = Item_ID
  // B = Description
  // C = Real-Time Stocks
  const data = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  const stockMap = {};

  data.forEach((r) => {
    const description = (r[1] || "").toString().trim();
    const stock = Number(r[2]) || 0;
    if (description) {
      stockMap[description] = stock;
    }
  });

  return stockMap;
}

// === BRANCH LIST FETCHER ===
function getBranchList() {
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sh = ss.getSheetByName("Branch List");
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // Branch header is in column B (index 2). Read from row 2, column 2.
  const values = sh.getRange(2, 2, lastRow - 1, 1).getValues();
  const branches = [];
  const seen = {};
  values.forEach((r) => {
    const v = (r[0] || "").toString().trim();
    if (v && !seen[v]) {
      branches.push(v);
      seen[v] = true;
    }
  });
  return branches;
}
