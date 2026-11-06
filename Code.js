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
    feedback: getFeedbackData("18VmspeT4fjVB1DG2ye8FzhkqdHOxgqyL1MYyBjSTEnI"),
    tickets: getTicketData("1xlGdKp9WZY2J6WcUvkfYrx1GTJHvrdrLvzeFmQPY86o"),
    members: getMembersData("1LLS2iXbgQzg9exlMlKLVZX9wGLqAcOcC2gBViVlqHRk") // API lấy Thành viên
  };
}

// HÀM MỚI: LẤY DỮ LIỆU THÀNH VIÊN
function getMembersData(spreadsheetId) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName("ThanhVien");
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var cMaSo = headers.indexOf("MaSoThanhVien");
    var cTen = headers.indexOf("TenThanhVien");
    var cBietDanh = headers.indexOf("BietDanh");
    var cHinhAnh = headers.indexOf("HinhAnhDaiDien");
    var cNguon = headers.indexOf("NguonDangKy");
    var cMongMuon = headers.indexOf("MongMuon");
    var cTinhThanh = headers.indexOf("TinhThanh");
    var cNamSinh = headers.indexOf("NamSinh");
    var cSDT = headers.indexOf("SoDienThoai");
    var cThoiGian = headers.indexOf("ThoiGianDangKy");
    var cTrangThai = headers.indexOf("TrangThai");

    var membersList = [];
    for (var i = 1; i < data.length; i++) {
      membersList.push({
        maSo: cMaSo > -1 ? data[i][cMaSo].toString().trim() : "",
        ten: cTen > -1 ? data[i][cTen].toString().trim() : "",
        bietDanh: cBietDanh > -1 ? data[i][cBietDanh].toString().trim() : "",
        hinhAnh: cHinhAnh > -1 ? data[i][cHinhAnh].toString().trim() : "",
        nguon: cNguon > -1 ? data[i][cNguon].toString().trim() : "",
        mongMuon: cMongMuon > -1 ? data[i][cMongMuon].toString().trim() : "",
        tinhThanh: cTinhThanh > -1 ? data[i][cTinhThanh].toString().trim() : "",
        namSinh: cNamSinh > -1 ? data[i][cNamSinh].toString().trim() : "",
        sdt: cSDT > -1 ? data[i][cSDT].toString().trim() : "",
        thoiGian: cThoiGian > -1 ? formatDateString(data[i][cThoiGian]) : "",
        trangThai: cTrangThai > -1 ? data[i][cTrangThai].toString().trim() : ""
      });
    }
    return membersList;
  } catch (e) {
    return [];
  }
}

function getTicketData(spreadsheetId) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheetInfo = ss.getSheetByName("ThongTinSuKien");
    var sheetReg = ss.getSheetByName("DSDKSuKien");
    var infoData = sheetInfo.getDataRange().getDisplayValues();
    var regData = sheetReg.getDataRange().getDisplayValues();
    var infoHeaders = infoData[0];
    var regHeaders = regData[0];

    var iMaSuKien = infoHeaders.indexOf("MaDangKySuKien");
    var iTenSuKien = infoHeaders.indexOf("TenSuKien");
    var iNgayToChuc = infoHeaders.indexOf("NgayToChuc");
    var iGioToChuc = infoHeaders.indexOf("GioToChuc");
    var iDiaDiem = infoHeaders.indexOf("DiaDiemToChuc");
    var iGhiChuTrenVe = infoHeaders.indexOf("GhiChuTrenVe");
    var rThoiGian = regHeaders.indexOf("ThoiGian");
    var rHoTen = regHeaders.indexOf("HoTen");
    var rSoDienThoai = regHeaders.indexOf("SoDienThoai");
    var rTrangThai = regHeaders.indexOf("TrangThai");
    var rMaBiMat = regHeaders.indexOf("MaBiMat");
    var rMaSuKien = regHeaders.indexOf("MaSuKienWorkshop");

    var result = { events: [], registrations: {} };
    var eventMap = {};
    for (var i = 1; i < infoData.length; i++) {
      var eId = infoData[i][iMaSuKien];
      if (eId) {
        result.events.push({
          MaDangKySuKien: eId, TenSuKien: infoData[i][iTenSuKien] || "",
          NgayToChuc: formatDateString(infoData[i][iNgayToChuc]), 
          GioToChuc: infoData[i][iGioToChuc] || "",
          DiaDiemToChuc: infoData[i][iDiaDiem] || "", GhiChuTrenVe: infoData[i][iGhiChuTrenVe] || ""
        });
        eventMap[eId] = true;
        result.registrations[eId] = [];
      }
    }
    for (var j = 1; j < regData.length; j++) {
      var eId = regData[j][rMaSuKien];
      if (eId && eventMap[eId]) {
        result.registrations[eId].push({
          ThoiGian: formatDateString(regData[j][rThoiGian]), HoTen: regData[j][rHoTen] || "",
          SoDienThoai: regData[j][rSoDienThoai] || "", TrangThai: regData[j][rTrangThai] || "", MaBiMat: regData[j][rMaBiMat] || ""
        });
      }
    }
    return result;
  } catch (e) { return { events: [], registrations: {} }; }
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
        eventsMap[id] = { name: infoData[i][infoNameCol] ? infoData[i][infoNameCol].toString().trim() : "", date: infoData[i][infoDateCol] ? infoData[i][infoDateCol].toString().trim() : "" };
      }
    }
    var groups = {};
    for (var j = 1; j < reviewData.length; j++) {
      var eId = reviewData[j][revEventIdCol];
      if (!eId) continue;
      if (!groups[eId]) { groups[eId] = { info: eventsMap[eId] || { name: "Không xác định", date: "Không xác định" }, reviews: [] }; }
      groups[eId].reviews.push({
        time: reviewData[j][revTimeCol] ? reviewData[j][revTimeCol].toString().trim() : "",
        format: parseFloat(reviewData[j][revFormatCol]) || 0, content: parseFloat(reviewData[j][revContentCol]) || 0,
        useful: parseFloat(reviewData[j][revUsefulCol]) || 0, comment: reviewData[j][revCommentCol] || ""
      });
    }
    for (var key in eventsMap) { if (!groups[key]) { groups[key] = { info: eventsMap[key], reviews: [] } } }
    return groups;
  } catch (e) { return {}; }
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
        groups[eId] = { info: eventsMap[eId] || { date: "Không xác định", extra: "", status: "", statusCheckIn: "" }, participants: [] };
      }
      var statusVal = "Đã duyệt";
      if (statusCol > -1 && regData[j][statusCol] !== undefined && regData[j][statusCol].toString().trim() !== "") {
          statusVal = regData[j][statusCol];
      }
      groups[eId].participants.push({
        time: formatDateString(regData[j][dateCol]),
        name: regData[j][nameCol], phone: regData[j][phoneCol], status: statusVal, 
        maBiMat: regMaBiMatCol > -1 ? regData[j][regMaBiMatCol] : "",
        checkIn: regCheckInCol > -1 ? regData[j][regCheckInCol] : ""
      });
    }
    return groups;
  } catch (e) { return {}; }
}