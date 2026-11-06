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
    var p1 = parseInt(match[1], 10);
    var p2 = parseInt(match[2], 10);
    var year = match[3];
    var time = match[4];
    
    if (p1 > 12) {
       return match[1].padStart(2, '0') + '/' + match[2].padStart(2, '0') + '/' + year + time;
    } else {
       return match[2].padStart(2, '0') + '/' + match[1].padStart(2, '0') + '/' + year + time;
    }
  }
  return str;
}

function fetchAllData() {
  return {
    masoi: getGroupedData("1-H_mbJpwwLQvzLsjuF2ZG1lQaMlV0eXm0nIkICHphg8", "DangKyLangMaSoi", "LichLangMaSoi", "MaSuKienMaSoi", ["DiaDiem"]),
    cafe: getGroupedData("1F6Xo5Xq0WRefvGZMtxk-sqp5yqXTwJ6Gy-bGRVzImms", "DangKyCafe", "LichCafe", "MaSuKienCafe", ["DiaDiem"]),
    event: getGroupedData("1xlGdKp9WZY2J6WcUvkfYrx1GTJHvrdrLvzeFmQPY86o", "DSDKSuKien", "ThongTinSuKien", ["MaSuKienWorkshop", "MaDangKySuKien"], ["TenSuKien"]),
    talk: getGroupedData("1WQKsrBCiY6DrEnLSMq1jxH8fn29t0npvwKq1DdvYb3o", "DangKyPrideTalk", "LichPrideTalk", "MaSuKienTalk", ["TenChuDe"]),
    sport: getGroupedData("1Mu25JGgxcaKwMZ2iklVcb9wvEVibR4QW-WUIxkQl-ac", "DangKySport", "LichSport", "MaSport", ["TenSport"]),
    feedback: getFeedbackData("18VmspeT4fjVB1DG2ye8FzhkqdHOxgqyL1MYyBjSTEnI")
  };
}

function getFeedbackData(spreadsheetId) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheetInfo = ss.getSheetByName("SuKienDanhGia");
    var sheetReview = ss.getSheetByName("NoiDungDanhGia");
    
    var infoData = sheetInfo.getDataRange().getDisplayValues();
    var reviewData = sheetReview.getDataRange().getDisplayValues();
    
    var infoHeaders = infoData[0];
    var reviewHeaders = reviewData[0];
    
    var infoEventIdCol = infoHeaders.indexOf("MaSuKienDanhGia");
    var infoNameCol = infoHeaders.indexOf("TenSuKien");
    var infoDateCol = infoHeaders.indexOf("NgayToChuc");
    
    var revEventIdCol = reviewHeaders.indexOf("MaSuKienDanhGia");
    var revFormatCol = reviewHeaders.indexOf("DanhGiaHinhThuc");
    var revContentCol = reviewHeaders.indexOf("DanhGiaNoiDung");
    var revUsefulCol = reviewHeaders.indexOf("DanhGiaBoIch");
    var revCommentCol = reviewHeaders.indexOf("GopY");
    var revTimeCol = reviewHeaders.indexOf("ThoiGianDanhGia");

    var eventsMap = {};
    for (var i = 1; i < infoData.length; i++) {
      var id = infoData[i][infoEventIdCol];
      if (id) {
        eventsMap[id] = {
          name: infoData[i][infoNameCol] ? infoData[i][infoNameCol].toString().trim() : "",
          date: infoData[i][infoDateCol] ? infoData[i][infoDateCol].toString().trim() : "" 
        };
      }
    }

    var groups = {};
    for (var j = 1; j < reviewData.length; j++) {
      var eId = reviewData[j][revEventIdCol];
      if (!eId) continue;
      
      if (!groups[eId]) {
        groups[eId] = {
          info: eventsMap[eId] || { name: "Không xác định", date: "Không xác định" },
          reviews: []
        };
      }
      
      groups[eId].reviews.push({
        time: reviewData[j][revTimeCol] ? reviewData[j][revTimeCol].toString().trim() : "",
        format: parseFloat(reviewData[j][revFormatCol]) || 0,
        content: parseFloat(reviewData[j][revContentCol]) || 0,
        useful: parseFloat(reviewData[j][revUsefulCol]) || 0,
        comment: reviewData[j][revCommentCol] || ""
      });
    }
    
    for (var key in eventsMap) {
      if (!groups[key]) {
         groups[key] = {
           info: eventsMap[key],
           reviews: []
         }
      }
    }
    return groups;
  } catch (e) {
    return {};
  }
}

function getGroupedData(spreadsheetId, sheetRegName, sheetInfoName, eventIdKeys, extraInfoKeys) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheetReg = ss.getSheetByName(sheetRegName);
    var sheetInfo = ss.getSheetByName(sheetInfoName);
    
    var regData = sheetReg.getDataRange().getDisplayValues();
    var infoData = sheetInfo.getDataRange().getDisplayValues();
    
    var infoHeaders = infoData[0];
    var regHeaders = regData[0];
    
    var regEventIdCol = typeof eventIdKeys === 'string' ? regHeaders.indexOf(eventIdKeys) : regHeaders.indexOf(eventIdKeys[0]);
    var infoEventIdCol = typeof eventIdKeys === 'string' ? infoHeaders.indexOf(eventIdKeys) : infoHeaders.indexOf(eventIdKeys[1]);
    
    // Fallback nếu người dùng đặt tên cột khoá chính hơi khác (ví dụ MaSuKienSport thay vì MaSport)
    if(regEventIdCol === -1) {
      for(var k=0; k<regHeaders.length; k++) {
        if(regHeaders[k].toLowerCase().indexOf("ma") > -1) { regEventIdCol = k; break; }
      }
    }
    
    var infoDateCol = infoHeaders.indexOf("NgayToChuc");
    var extraInfoCol = infoHeaders.indexOf(extraInfoKeys[0]);
    var infoStatusCol = infoHeaders.indexOf("TrangThai"); 
    var infoStatusCheckInCol = infoHeaders.indexOf("TrangThaiCheckIn");
    
    var dateCol = regHeaders.indexOf("ThoiGianDangKy") > -1 ? regHeaders.indexOf("ThoiGianDangKy") : regHeaders.indexOf("ThoiGian");
    var nameCol = regHeaders.indexOf("HoVaTen") > -1 ? regHeaders.indexOf("HoVaTen") : (regHeaders.indexOf("HoTen") > -1 ? regHeaders.indexOf("HoTen") : regHeaders.indexOf("Tên"));
    var phoneCol = regHeaders.indexOf("SoDienThoai");
    var statusCol = regHeaders.indexOf("TrangThaiPheDuyet") > -1 ? regHeaders.indexOf("TrangThaiPheDuyet") : regHeaders.indexOf("TrangThai");
    
    var regMaBiMatCol = regHeaders.indexOf("MaBiMat");
    var regCheckInCol = regHeaders.indexOf("CheckIn");

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
      
      // Fix: Nếu Sheet (như Sport) không có cột Trạng thái, tự động gán là "Đã duyệt"
      var statusVal = "Đã duyệt";
      if (statusCol > -1 && regData[j][statusCol] !== undefined && regData[j][statusCol].toString().trim() !== "") {
          statusVal = regData[j][statusCol];
      }

      groups[eId].participants.push({
        time: formatDateString(regData[j][dateCol]),
        name: regData[j][nameCol],
        phone: regData[j][phoneCol],
        status: statusVal, 
        maBiMat: regMaBiMatCol > -1 ? regData[j][regMaBiMatCol] : "",
        checkIn: regCheckInCol > -1 ? regData[j][regCheckInCol] : ""
      });
    }
    
    return groups;
  } catch (e) {
    return {};
  }
}