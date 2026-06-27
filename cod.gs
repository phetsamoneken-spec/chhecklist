const SPREADSHEET_ID = "1nYX0Wq8fUUkQmtZuN3wEMA_8W-0Vo0aFdVezo9oaRnI";
const SHEET_NAME = "Outlets";

/**
 * ຟັງຊັນຫຼັກໃນການເປີດໜ້າເວັບແອັບ (Web App)
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Outlet Checklist')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ຟັງຊັນຊ່ວຍໃນການດຶງກຳນົດ Spreadsheet Object ຈາກ ID
 */
function getTargetSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // ຫາກບໍ່ມີ Sheet ຊື່ "Outlets" ໃຫ້ສ້າງຂຶ້ນມາໃໝ່ອັດຕະໂນມັດ
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["ID", "Name", "Location", "Checked", "Phone", "Details"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
  }
  return sheet;
}

/**
 * ດຶງຂໍ້ມູນ Outlets ທັງໝົດຈາກ Sheet
 */
function getOutletsFromSheet() {
  try {
    const sheet = getTargetSheet();
    const data = sheet.getDataRange().getValues();
    const outlets = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "") continue; 
      outlets.push({
        id: Number(data[i][0]),
        name: data[i][1] ? data[i][1].toString() : "",
        location: data[i][2] ? data[i][2].toString() : "",
        checked: data[i][3] === true || data[i][3] === "TRUE",
        phone: data[i][4] ? data[i][4].toString() : "",
        details: data[i][5] ? data[i][5].toString() : ""
      });
    }
    return outlets;
  } catch (error) {
    Logger.log("Error fetching outlets: " + error.toString());
    return [];
  }
}

/**
 * ບັນທຶກ ຫຼື ຂຽນທັບຂໍ້ມູນທັງໝົດລົງໃນ Google Sheet (ສຳລັບການ Import ຂໍ້ມູນຈຳນວນຫຼາຍ)
 */
function saveAllOutletsToSheet(outletsList) {
  try {
    const sheet = getTargetSheet();
    
    // ລຶບຂໍ້ມູນເກົ່າທັງໝົດ (ຍົກເວັ້ນ Header ແຖວທີ 1)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    
    // ກຽມຂໍ້ມູນແຖວໃໝ່
    const rows = outletsList.map(item => [
      item.id,
      item.name,
      item.location,
      item.checked,
      item.phone || "ບໍ່ມີຂໍ້ມູນ",
      item.details || ""
    ]);
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }
    return true;
  } catch (error) {
    Logger.log("Error saving all outlets: " + error.toString());
    return false;
  }
}

/**
 * ອັບເດດສະຖານະ (Checked: TRUE/FALSE)
 */
function updateStatusInSheet(id, isChecked) {
  try {
    const sheet = getTargetSheet();
    const data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === id) {
        sheet.getRange(i + 1, 4).setValue(isChecked); 
        return true;
      }
    }
    return false;
  } catch (error) {
    Logger.log("Error updating status: " + error.toString());
    return false;
  }
}

/**
 * ເພີ່ມຮ້ານຄ້າໃໝ່ (New Outlet)
 */
function addNewOutletToSheet(newShop) {
  try {
    const sheet = getTargetSheet();
    sheet.appendRow([
      newShop.id,
      newShop.name,
      newShop.location,
      newShop.checked,
      newShop.phone,
      newShop.details
    ]);
    return true;
  } catch (error) {
    Logger.log("Error adding new outlet: " + error.toString());
    return false;
  }
}

/**
 * ລຶບຮ້ານຄ້າລາຍຮ້ານອອກຈາກ Google Sheet ຕາມ ID
 */
function deleteOutletFromSheet(id) {
  try {
    const sheet = getTargetSheet();
    const data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === id) {
        sheet.deleteRow(i + 1); 
        return true;
      }
    }
    return false;
  } catch (error) {
    Logger.log("Error deleting outlet from sheet: " + error.toString());
    return false;
  }
}

/**
 * ຣີເຊັດສະຖານະຂອງທຸກຮ້ານຄ້າໃຫ້ເປັນ FALSE
 */
function resetAllStatusInSheet() {
  try {
    const sheet = getTargetSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 4, lastRow - 1, 1).setValue(false);
    }
    return true;
  } catch (error) {
    Logger.log("Error resetting statuses: " + error.toString());
    return false;
  }
}

/**
 * ລຶບຂໍ້ມູນຮ້ານຄ້າທັງໝົດໃນ Sheet
 */
function clearAllOutletsInSheet() {
  try {
    const sheet = getTargetSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return true;
  } catch (error) {
    Logger.log("Error clearing outlets: " + error.toString());
    return false;
  }
}
