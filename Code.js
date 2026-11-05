function doPost(e) {
  try {
    var request = JSON.parse(e.postData.contents);
    var action = request.action;

    if (action === "login") {
      var result = checkLogin(request.username, request.password);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    } else if (action === "getData") {
      var data = fetchAllData();
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: data })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("API Endpoint is running.").setMimeType(ContentService.MimeType.TEXT);
}

function checkLogin(username, passwordMd5) {
  var sheet = SpreadsheetApp.openById("1RBXdJs6HZB0rPCRU0YyOcEUt_usBGhYQq0ePGpLLUCg").getSheetByName("TaiKhoanQuanTri");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var dbUser = data[i][2] ? data[i][2].toString().trim() : ""; 
    var dbPass = data[i][3] ? data[i][3].toString().trim().toLowerCase() : ""; 
    var dbStatus = data[i][4] ? data[i][4].toString().trim() : ""; 
    
    var inputUser = username ? username.toString().trim() : "";
    var inputPassMd5 = passwordMd5 ? passwordMd5.toString().trim().toLowerCase() : "";

    if (dbUser === inputUser && dbPass === inputPassMd5) {
      if (dbStatus.toLowerCase() !== "hoạt động" && dbStatus.toLowerCase() !== "active" && dbStatus !== "true" && dbStatus !== "1") {
         return { success: false, message: "Tài khoản đang bị khóa." };
      }
      return { success: true, name: data[i][1] };
    }
  }
  return { success: false, message: "Tên đăng nhập hoặc mật khẩu không đúng." };
}

function formatDateString(dateStr) {
  if (!dateStr) return "";
  var str = dateStr.toString().trim();
  var match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(.*)$/);
  if (match) {
    var month = match[1].padStart(2, '0');
    var day = match[2].padStart(2, '0');
    var year = match[3];
    var time = match[4];
    return day + '/' + month + '/' + year + time;
  }
  return str;
}

function fetchAllData() {
  return {
    masoi: getGroupedData("1-H_mbJpwwLQvzLsjuF2ZG1lQaMlV0eXm0nIkICHphg8", "DangKyLangMaSoi", "LichLangMaSoi", "MaSuKienMaSoi", ["DiaDiem"]),
    cafe: getGroupedData("1F6Xo5Xq0WRefvGZMtxk-sqp5yqXTwJ6Gy-bGRVzImms", "DangKyCafe", "LichCafe", "MaSuKienCafe", ["DiaDiem"]),
    event: getGroupedData("1xlGdKp9WZY2J6WcUvkfYrx1GTJHvrdrLvzeFmQPY86o", "DSDKSuKien", "ThongTinSuKien", ["MaSuKienWorkshop", "MaDangKySuKien"], ["TenSuKien"]),
    talk: getGroupedData("1WQKsrBCiY6DrEnLSMq1jxH8fn29t0npvwKq1DdvYb3o", "DangKyPrideTalk", "LichPrideTalk", "MaSuKienTalk", ["TenChuDe"])
  };
}

function getGroupedData(spreadsheetId, sheetRegName, sheetInfoName, eventIdKeys, extraInfoKeys) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheetReg = ss.getSheetByName(sheetRegName);
  var sheetInfo = ss.getSheetByName(sheetInfoName);
  
  var regData = sheetReg.getDataRange().getDisplayValues();
  var infoData = sheetInfo.getDataRange().getDisplayValues();
  
  var infoHeaders = infoData[0];
  var regHeaders = regData[0];
  
  var regEventIdCol = typeof eventIdKeys === 'string' ? regHeaders.indexOf(eventIdKeys) : regHeaders.indexOf(eventIdKeys[0]);
  var infoEventIdCol = typeof eventIdKeys === 'string' ? infoHeaders.indexOf(eventIdKeys) : infoHeaders.indexOf(eventIdKeys[1]);
  
  var infoDateCol = infoHeaders.indexOf("NgayToChuc");
  var extraInfoCol = infoHeaders.indexOf(extraInfoKeys[0]);
  var infoStatusCol = infoHeaders.indexOf("TrangThai"); 
  var infoStatusCheckInCol = infoHeaders.indexOf("TrangThaiCheckIn"); // Cột trạng thái checkin
  
  var dateCol = regHeaders.indexOf("ThoiGianDangKy") > -1 ? regHeaders.indexOf("ThoiGianDangKy") : regHeaders.indexOf("ThoiGian");
  var nameCol = regHeaders.indexOf("HoVaTen") > -1 ? regHeaders.indexOf("HoVaTen") : regHeaders.indexOf("HoTen");
  var phoneCol = regHeaders.indexOf("SoDienThoai");
  var statusCol = regHeaders.indexOf("TrangThaiPheDuyet") > -1 ? regHeaders.indexOf("TrangThaiPheDuyet") : regHeaders.indexOf("TrangThai");
  
  var regMaBiMatCol = regHeaders.indexOf("MaBiMat"); // Mã bí mật
  var regCheckInCol = regHeaders.indexOf("CheckIn"); // Thời gian checkin

  var eventsMap = {};
  for (var i = 1; i < infoData.length; i++) {
    var id = infoData[i][infoEventIdCol];
    if (id) {
      eventsMap[id] = {
        date: formatDateString(infoData[i][infoDateCol]),
        extra: extraInfoCol > -1 ? infoData[i][extraInfoCol] : "",
        status: infoStatusCol > -1 ? infoData[i][infoStatusCol] : "",
        statusCheckIn: infoStatusCheckInCol > -1 ? infoData[i][infoStatusCheckInCol] : ""
      };
    }
  }

  var groups = {};
  for (var j = 1; j < regData.length; j++) {
    var eId = regData[j][regEventIdCol];
    if (!eId) continue;
    
    if (!groups[eId]) {
      groups[eId] = {
        info: eventsMap[eId] || { date: "Không xác định", extra: "", status: "", statusCheckIn: "" },
        participants: []
      };
    }
    
    groups[eId].participants.push({
      time: formatDateString(regData[j][dateCol]),
      name: regData[j][nameCol],
      phone: regData[j][phoneCol],
      status: regData[j][statusCol],
      maBiMat: regMaBiMatCol > -1 ? regData[j][regMaBiMatCol] : "",
      checkIn: regCheckInCol > -1 ? regData[j][regCheckInCol] : ""
    });
  }
  
  return groups;
}