var ss = SpreadsheetApp.getActiveSpreadsheet();

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FitNavy - Calorie Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// 1. ตรวจสอบการเข้าสู่ระบบ
function checkUserLogin(username) {
  try {
    var sheet = ss.getSheetByName("Profile");
    if (!sheet) return { exists: false, message: "ไม่พบชีท Profile" };
    
    var rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { exists: false };
    
    var searchUser = username.toString().trim().toLowerCase();
    
    for (var i = 1; i < rows.length; i++) {
      var dbUser = rows[i][0] ? rows[i][0].toString().trim().toLowerCase() : "";
      if (dbUser === searchUser) {
        var createdDateVal = rows[i][9] ? new Date(rows[i][9]) : new Date();
        return {
          exists: true,
          profile: {
            username: rows[i][0].toString().trim(),
            age: rows[i][1] || 25,
            gender: rows[i][2] || "male",
            weight: rows[i][3] || 60,
            height: rows[i][4] || 170,
            activity: rows[i][5] || "1.55",
            goal: rows[i][6] || "maintain",
            targetWeightLoss: rows[i][7] || 0,
            targetCalories: rows[i][8] || 1500,
            createdDate: Utilities.formatDate(createdDateVal, Session.getScriptTimeZone(), "yyyy-MM-dd")
          }
        };
      }
    }
    return { exists: false };
  } catch (err) {
    return { exists: false, error: err.toString() };
  }
}

// 2. สร้างโปรไฟล์ใหม่
function createNewProfile(data) {
  try {
    var sheet = ss.getSheetByName("Profile");
    if (!sheet) return { success: false, message: "ไม่พบชีท Profile" };
    
    var rows = sheet.getDataRange().getValues();
    var searchUser = data.username.toString().trim().toLowerCase();
    
    for (var i = 1; i < rows.length; i++) {
      var dbUser = rows[i][0] ? rows[i][0].toString().trim().toLowerCase() : "";
      if (dbUser === searchUser) {
        return { success: false, message: "ชื่อผู้ใช้นี้มีในระบบแล้ว กรุณากดเข้าสู่ระบบ" };
      }
    }
    
    var now = new Date();
    sheet.appendRow([
      data.username.trim(),
      data.age,
      data.gender,
      data.weight,
      data.height,
      data.activity,
      data.goal,
      data.targetWeightLoss || 0,
      data.targetCalories,
      now
    ]);
    
    data.createdDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    return { success: true, profile: data };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// 3. ดึงข้อมูลประจำวัน
function getDailyData(username, dateStr) {
  try {
    var sheet = ss.getSheetByName("Data");
    if (!sheet) return [];
    
    var rows = sheet.getDataRange().getValues();
    var items = [];
    var searchUser = username.toString().trim().toLowerCase();
    
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      
      var rowDate = Utilities.formatDate(new Date(rows[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var dbUser = rows[i][1] ? rows[i][1].toString().trim().toLowerCase() : "";
      
      if (rowDate === dateStr && dbUser === searchUser) {
        items.push({
          type: rows[i][2],
          name: rows[i][3],
          calories: Number(rows[i][4]),
          detail: rows[i][5]
        });
      }
    }
    return items;
  } catch (err) {
    return [];
  }
}

// 4. บันทึกข้อมูลหลายรายการ
function saveBatchData(dataList) {
  try {
    var sheet = ss.getSheetByName("Data");
    var pSheet = ss.getSheetByName("Profile");
    
    dataList.forEach(function(data) {
      sheet.appendRow([
        data.date,
        data.username,
        data.type,
        data.name,
        data.calories,
        data.detail,
        new Date()
      ]);

      if (data.type === 'weight') {
        var pRows = pSheet.getDataRange().getValues();
        for (var i = 1; i < pRows.length; i++) {
          if (pRows[i][0] && pRows[i][0].toString().trim().toLowerCase() === data.username.toString().trim().toLowerCase()) {
            pSheet.getRange(i + 1, 4).setValue(data.calories);
            break;
          }
        }
      }
    });
    
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}