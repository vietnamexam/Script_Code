// --- FILE TỔNG TRÊN GITHUB ---

function mainDoGet(e) { 

  const params = e?.parameter || {};
  const type = params.type; 
  const action = params.action || e.parameter.action;
  var mon = e.parameter.monId || e.parameter.mon || params.mon;

if (!mon || mon === "undefined" || mon === "unknown") {     
    mon = "chung";
}

if (action === "checkAdminOTP") {
    const userOTP = (params.otp || "").trim();
    const inputIDGV = (params.idgv || "").trim().toLowerCase();

    // Kiểm tra IDGV (Giữ nguyên logic cũ của bạn)
    const idgvSheet = ssAdmin.getSheetByName("idgv");
    const dataIDGV = idgvSheet.getRange("A:A").getValues().flat();
    const isIdValid = dataIDGV.some(id => String(id).trim().toLowerCase() === inputIDGV);

    if (!isIdValid) return createResponse("error", "ID Giáo viên không tồn tại!");

    // Kiểm tra OTP và lấy tên môn
    const assignedMon = MAP_MON[userOTP];

    if (assignedMon) {
      return createResponse("success", "Xác minh thành công", {
        verified: true,
        mon: assignedMon // Gửi về cho React: "toan", "ly",...
      });
    } else {
      return createResponse("error", "Mật khẩu Admin không đúng!");
    }
  }
  // load ngân hàng đề
  if (action === "loadQuestions") {     
    var sheetNH = getsheetname(mon); 
    var values = sheetNH.getDataRange().getValues();
    if (values.length <= 1) {
      return createResponse("success", "Không có dữ liệu", []);
    }
    var rows = values.slice(1);

    var result = rows.map(function (r) {

      var obj = {
        id: r[0],
        classTag: r[1],
        type: r[2],
        part: r[3],
        question: r[4]
      };

      if (r[2] === "mcq") {
        obj.o = r[5] ? JSON.parse(r[5]) : [];
        obj.a = r[6];
      }

      if (r[2] === "true-false") {
        obj.s = r[5] ? JSON.parse(r[5]) : [];
      }

      if (r[2] === "short-answer") {
        obj.a = r[6];
      }

      return obj;
    });

    return createResponse("success", "Load thành công", result);
  }


  // Reset QuiZ
  if (action === "resetQuiz") {
  return resetQuizData(e.parameter.password);
}



  if (action === 'checkTeacher') {
    try {
      const idInput = (params.idgv || "").toString().trim();
      if (!idInput) return createResponse("error", "Chưa nhập ID giáo viên");

      const sheet = ss.getSheetByName("idgv");
      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        // Ép cả 2 về String để so sánh cho chuẩn
        let idInSheet = data[i][0].toString().trim();
        
        if (idInSheet === idInput) {
          return createResponse("success", "OK", { 
            name: data[i][1], 
            link: data[i][2] 
          });
        }
      }
      return createResponse("error", "Không tìm thấy ID: " + idInput);
    } catch (err) {
      return createResponse("error", "Lỗi Script: " + err.toString());
    }
  }
  
  if (action === 'getLG') {
    var idTraCuu = params.id;    
    var targetSheet = getsheetname(mon);
    if (!idTraCuu) return ContentService.createTextOutput("Thiếu ID rồi!").setMimeType(ContentService.MimeType.TEXT);

    var data = targetSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === idTraCuu.toString().trim()) {
        var loigiai = data[i][7] || "";

        // Ép kiểu về String để đảm bảo không bị lỗi tệp
        return ContentService.createTextOutput(String(loigiai))
          .setMimeType(ContentService.MimeType.TEXT);
      }
    }
    return ContentService.createTextOutput("Không tìm thấy ID này!").setMimeType(ContentService.MimeType.TEXT);
  }
   if (action === 'updateLG') {
  var sheetNH = getsheetname(mon);
  var data = JSON.parse(e.postData.contents);
  var id = data.id;
  var lg = data.loigiai;

  var values = sheetNH.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    if (values[i][0] == id) {
      sheetNH.getRange(i + 1, 8).setValue(lg); // cột loigiai
      break;
    }
  }

  return createResponse("success", "Đã cập nhật lời giải!");
}
 // lấy dạng câu hỏi
  if (action === 'getAppConfig') {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: getAppConfig()
    })).setMimeType(ContentService.MimeType.JSON);
  }
// THÊM NHÁNH NÀY CHO MA TRẬN
if (action === 'getAppConfigmt') {
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: getAppConfigmt()
  })).setMimeType(ContentService.MimeType.JSON);
}


   
// 4. KIỂM TRA GIÁO VIÊN (Dành cho Module Giáo viên tạo đề word)
    
   
   // Trong hàm doGet(e) của Google Apps Script
if (action === "getRouting") {
  const sheet = ss.getSheetByName("idgv");
  const rows = sheet.getDataRange().getValues();
  const data = [];
  for (var i = 1; i < rows.length; i++) {
    data.push({
      idNumber: rows[i][0], // Cột A
      link: rows[i][2]      // Cột C
    });
  }
  return createResponse("success", "OK", data);
}

  // 1. ĐĂNG KÝ / ĐĂNG NHẬP
  var sheetAcc = ss.getSheetByName("account");
  if (action === "register") {
    var phone = params.phone;
    var pass = params.pass;
    var rows = sheetAcc.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1].toString() === phone) return ContentService.createTextOutput("exists");
    }
    sheetAcc.appendRow([new Date(), "'" + phone, pass, "VIP0"]);
    return ContentService.createTextOutput("success");
  }

  if (action === "login") {
    var phone = params.phone;
    var pass = params.pass;
    var rows = sheetAcc.getDataRange().getValues();
    
    for (var i = 1; i < rows.length; i++) {
      // Kiểm tra số điện thoại (cột B) và mật khẩu (cột C)
      if (rows[i][1].toString() === phone && rows[i][2].toString() === pass) {
        
        return createResponse("success", "OK", { 
          phoneNumber: rows[i][1].toString(), 
          vip: rows[i][3] ? rows[i][3].toString() : "VIP0",
          name: rows[i][4] ? rows[i][4].toString() : "" // Lấy thêm cột E (tên người dùng)
        });
      }
    }
    return ContentService.createTextOutput("fail");
  }

  // 2. LẤY DANH SÁCH ỨNG DỤNG
  if (params.sheet === "ungdung") {
    var sheet = ss.getSheetByName("ungdung");
    var rows = sheet.getDataRange().getValues();
    var data = [];
    for (var i = 1; i < rows.length; i++) {
      data.push({ name: rows[i][0], icon: rows[i][1], link: rows[i][2] });
    }
    return resJSON(data);
  }

  // 3. TOP 10
  if (type === 'top10') {
    const sheet = ss.getSheetByName("Top10Display");
    if (!sheet) return createResponse("error", "Không tìm thấy sheet Top10Display");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return createResponse("success", "Chưa có dữ liệu Top 10", []);
    const values = sheet.getRange(2, 1, Math.min(10, lastRow - 1), 10).getValues();
    const top10 = values.map((row, index) => ({
      rank: index + 1, name: row[0], phoneNumber: row[1], score: row[2],
      time: row[3], sotk: row[4], bank: row[5], idPhone: row[9]
    }));
    return createResponse("success", "OK", top10);
  }

  // 4. THỐNG KÊ ĐÁNH GIÁ
  if (type === 'getStats') {
    const stats = { ratings: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 } };
    const sheetRate = ss.getSheetByName("danhgia");
    if (sheetRate) {
      const rateData = sheetRate.getDataRange().getValues();
      for (let i = 1; i < rateData.length; i++) {
        const star = parseInt(rateData[i][1]);
        if (star >= 1 && star <= 5) stats.ratings[star]++;
      }
    }
    return createResponse("success", "OK", stats);
  }

  // 5. LẤY MẬT KHẨU QUIZ
  if (type === 'getPass') {    
    const password = ADMIN_RESET_PASSWORD;
    return resJSON({ password: password.toString() });
  }

  // 6. XÁC MINH THÍ SINH
  if (type === 'verifyStudent') {
    const idNumber = params.idnumber;
    const sbd = params.sbd;
    const mon = params.mon;

    const sheet = ss.getSheetByName("danhsach");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[1][5].toString().trim() === idNumber.trim() && data[i][0].toString().trim() === sbd.trim()) {
        return createResponse("success", "OK", {
          name: data[i][1], class: data[i][2], limit: data[i][3],
          limittab: data[i][4], taikhoanapp: data[i][6], idnumber: idNumber, sbd: sbd
        });
      }
    }
    return createResponse("error", "Thí sinh không tồn tại!");
  }

  // 7. LẤY CÂU HỎI THEO ID
  if (action === 'getQuestionById') {
    var id = params.id;
    
      Logger.log("MON NHẬN: " + mon);

      var sheetNH = getsheetname(mon);

      Logger.log("SHEET ĐỌC: " + sheet.getName());
    var dataNH = sheetNH.getDataRange().getValues();
    for (var i = 1; i < dataNH.length; i++) {
      if (dataNH[i][0].toString() === id.toString()) {
        return createResponse("success", "OK", {
          idquestion: dataNH[i][0],
          classTag: dataNH[i][1],
          question: dataNH[i][4],
          options: dataNH[i][5],
          answer: dataNH[i][6],
          loigiai: dataNH[i][7],
          datetime: dataNH[i][8]
        });
      }
    }
    return resJSON({ status: 'error' });
  }

  // 8. LẤY MA TRẬN ĐỀ
  if (type === 'getExamCodes') {
    const teacherId = params.idnumber;
    const sheet = ss.getSheetByName("matran");
    const data = sheet.getDataRange().getValues();
    const results = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0].toString().trim() === teacherId.trim() || row[0].toString() === "SYSTEM") {
        try {
          results.push({
            code: row[1].toString(), name: row[2].toString(), topics: JSON.parse(row[3]),
            fixedConfig: {
              duration: parseInt(row[4]), numMC: JSON.parse(row[5]), scoreMC: parseFloat(row[6]),
              mcL3: JSON.parse(row[7]), mcL4: JSON.parse(row[8]), numTF: JSON.parse(row[9]),
              scoreTF: parseFloat(row[10]), tfL3: JSON.parse(row[11]), tfL4: JSON.parse(row[12]),
              numSA: JSON.parse(row[13]), scoreSA: parseFloat(row[14]), saL3: JSON.parse(row[15]), saL4: JSON.parse(row[16])
            }
          });
        } catch (err) {}
      }
    }
    return createResponse("success", "OK", results);
  }

  // 9. LẤY TẤT CẢ CÂU HỎI (Hàm này thầy bị trùng, em gom lại bản chuẩn nhất)
  if (action === "getQuestions") {
    var sheet = getsheetname(mon);
    if (!sheet) {
  return createResponse("error", "Không tìm thấy sheet môn: " + mon, []);
}
    var lastRow = sheet.getLastRow();
    var rows = sheet.getRange(2,1,lastRow-1,9).getValues();
    var questions = [];
    for (var i = 0; i < rows.length; i++) {
      var raw = rows[i][2];
      if (!raw) continue;
      try {
        var jsonText = raw.replace(/(\w+)\s*:/g, '"$1":').replace(/'/g, '"');
        var obj = JSON.parse(jsonText);
        if (!obj.classTag) obj.classTag = rows[i][1];
        obj.loigiai = rows[i][7] || "";
        questions.push(obj);
      } catch (e) { }
    }
    return createResponse("success", "OK", questions);
  }

  return createResponse("error", "Yêu cầu không hợp lệ");
}
// =====================================================================================================================Hết Doget =======================================
function mainDoPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {    
    const idgv = (e.parameter.idgv || JSON.parse(e.postData.contents).idgv || "").toString().trim();
    const data = JSON.parse(e.postData.contents || "{}");
    const params = e.parameter || {};    
    const action = params.action || e.parameter.action;
    var mon = e.parameter.monId || e.parameter.mon;

  if (!mon || mon === "undefined" || mon === "unknown") {
  mon = "chung";
  }
   
    
        // 1. NHÁNH LƯU CẤU HÌNH (Ổn định theo kiểu saveMatrix)
    if (action === 'saveExamConfig') {
      // BƯỚC 1: Xác định file đích (Master hay Hàng xóm)
      const targetSS = getSpreadsheetByTarget(idgv);
      const sheet = targetSS.getSheetByName("exams") || targetSS.insertSheet("exams");
      
      // Tạo tiêu đề nếu sheet mới
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(["exams", "IdNumber", "MCQ", "scoremcq", "TF", "scoretf", "SA", "scoresa", "fulltime", "mintime", "tab", "dateclose"] );
      }

      // Chuẩn bị dữ liệu hàng (Row Data)
      const rowData = [
        data.exams, idgv, data.MCQ, data.scoremcq, data.TF, data.scoretf, data.SA, data.scoresa, data.fulltime, data.mintime, data.tab, data.dateclose 
      ];

      // BƯỚC 2: Kiểm tra xem mã đề đã tồn tại chưa để ghi đè (Giống logic Ma trận)
      const vals = sheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < vals.length; i++) {
        // Nếu trùng mã đề (cột A) và trùng ID GV (cột B)
        if (vals[i][0].toString() === data.exams.toString() && vals[i][1].toString() === idgv.toString()) {
          rowIndex = i + 1; 
          break;
        }
      }

      // BƯỚC 3: Ghi dữ liệu
      if (rowIndex > 0) {
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }

      return createResponse("success", "✅ Đã lưu cấu hình đề [" + data.exams + "] vào file: " + targetSS.getName());
    }
    // 5. UPLOAD DỮ LIỆU ĐỀ THI TỪ WORD (Teacher)
    if (action === 'uploadExamData') {
      const gvSS = getSpreadsheetByTarget(data.idgv);
      const sheet = gvSS.getSheetByName("exam_data") || gvSS.insertSheet("exam_data");
      const now = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yy");
      data.questions.forEach(q => {
        sheet.appendRow([
          data.examCode, q.classTag || "", q.type, 
          JSON.stringify(q), now, q.loigiai || ""
        ]);
      });
      return createResponse("success", "Đã tải lên " + data.questions.length + " câu!");
    }


    // 1. NHÁNH LỜI GIẢI (saveLG)
   if (action === 'saveLG') {
  
  var sheetNH = getsheetname(mon);
  var lastRow = sheetNH.getLastRow();
  if (lastRow < 1) return createResponse("error", "Database thiếu dòng 1 (header)!");

  // Lấy ID ở cột A (cột 1) để tra cứu
  var idValues = sheetNH.getRange(1, 1, lastRow, 1).getValues().flat().map(String);
  var count = 0;

  // data là mảng lời giải từ Web gửi về
  data.forEach(function (item) {
    var targetId = String(item.id || "").trim();
    if (!targetId) return;

    var rowIndex = idValues.indexOf(targetId);
    if (rowIndex !== -1) {
      var rowNum = rowIndex + 1;
      var contentLG = item.loigiai || item.lg || "";
      
      // Ghi vào cột K (cột 8)
      sheetNH.getRange(rowNum, 8).setValue(contentLG);
      count++;
    }
  });

  return createResponse("success", "Đã cập nhật " + count + " lời giải qua Web thành công!");
}
    // 2. NHÁNH MA TRẬN (saveMatrix)
    if (action === "saveMatrix") {
      const sheetMatran = ss.getSheetByName("matran") || ss.insertSheet("matran");
      const toStr = (v) => (v != null) ? String(v).trim() : "";
      const toNum = (v) => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
      const toJson = (v) => {
        if (!v || v === "" || (Array.isArray(v) && v.length === 0)) return "[]";
        if (typeof v === 'object') return JSON.stringify(v);
        let s = String(v).trim();
        return s.startsWith("[") ? s : "[" + s + "]";
      };
      const rowData = [
        toStr(data.gvId), toStr(data.makiemtra), toStr(data.name), toJson(data.topics),
        toNum(data.duration), toJson(data.numMC), toNum(data.scoreMC), toJson(data.mcL3),
        toJson(data.mcL4), toJson(data.numTF), toNum(data.scoreTF), toJson(data.tfL3),
        toJson(data.tfL4), toJson(data.numSA), toNum(data.scoreSA), toJson(data.saL3), toJson(data.saL4)
      ];
      const vals = sheetMatran.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0].toString() === toStr(data.gvId) && vals[i][1].toString() === toStr(data.makiemtra)) {
          rowIndex = i + 1; break;
        }
      }
      if (rowIndex > 0) { sheetMatran.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]); } 
      else { sheetMatran.appendRow(rowData); }
      return createResponse("success", "✅ Đã tạo ma trận " + data.makiemtra + " thành công!");
    }

    // 3. NHÁNH LƯU CÂU HỎI MỚI (saveQuestions)
    if (action === 'saveQuestions') {

      var now = new Date();
      
      var sheetNH = getsheetname(mon);
      if (!sheetNH) return createResponse("error", "Không tìm thấy sheet môn: " + params.mon);      
      var lastRow = sheetNH.getLastRow();
      if (lastRow < 1) return createResponse("error", "Database thiếu dòng 1 (header)!");

      var startRow = sheetNH.getLastRow() + 1;

      var rows = data.map(function (item) {
        return [
          item.id,
          item.classTag,
          item.type,
          item.part,
          item.question,
          item.options || "",
          item.answer || "",
          item.loigiai || "",
          now
        ];
      });

      if (rows.length > 0) {
    // THAY ĐỔI Ở ĐÂY: getRange(dòng bắt đầu, CỘT BẮT ĐẦU = 4, số dòng, số cột)
        sheetNH.getRange(startRow, 1, rows.length, rows[0].length)
      .setValues(rows);
    
    // Định dạng WrapText cho các cột nội dung từ cột H (8) đến cột K (11)
    // Thay vì "D:H" (toàn bộ sheet), chỉ làm cho những dòng vừa nạp để tránh lag
    sheetNH.getRange(startRow, 4, rows.length, 5).setWrap(true);
        }
      var lastRow = sheetNH.getLastRow();
      // Tự chỉnh chiều cao từ dòng 2 trở xuống
      if (lastRow > 1) {
        sheetNH.autoResizeRows(startRow, rows.length);
      }

      return createResponse("success", "Đã lưu " + rows.length + " câu hỏi thành công!");
    }


    // 4. XÁC MINH GIÁO VIÊN (verifyGV)
    if (action === "verifyGV") {
      var sheetGV = ss.getSheetByName("idgv");
      var rows = sheetGV.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString().trim() === data.idnumber.toString().trim() && rows[i][1].toString().trim() === data.password.toString().trim()) {
          return resJSON({ status: "success" });
        }
      }
      return resJSON({ status: "error", message: "ID hoặc Mật khẩu GV không đúng!" });
    }

    // 5. CẬP NHẬT CÂU HỎI (updateQuestion)
    if (action === 'updateQuestion') {

  const payload = JSON.parse(e.postData.contents);
  const item = payload.data;

  var targetId = item.id || item.idquestion;
  if (!targetId) return resJSON({ status:'error', message:'ID gửi lên bị trống!' });

  var sheetNH = getsheetname(mon);
  sheetNH.getRange("E:H").setNumberFormat("@");

  var lastRow = sheetNH.getLastRow();

  var idValues = sheetNH.getRange(1,1,lastRow,1).getValues().flat().map(String);

  var rowIndex = idValues.indexOf(String(targetId));

  if (rowIndex !== -1) {

    var rowNum = rowIndex + 1;

    sheetNH.getRange(rowNum,2).setValue(item.classTag || "");
    sheetNH.getRange(rowNum,5).setValue(item.question || "");
    sheetNH.getRange(rowNum,6).setValue(item.options || "");
    sheetNH.getRange(rowNum,7).setValue(String(item.answer || ""));
    sheetNH.getRange(rowNum,8).setValue(item.loigiai || "");
    sheetNH.getRange(rowNum,9).setValue(new Date().toLocaleString('vi-VN'));

    return resJSON({status:'success'});
  }

  return resJSON({status:'error',message:'Không tìm thấy ID: '+targetId});
}

    // 6. XÁC MINH ADMIN (verifyAdmin)
    if (action === "verifyAdmin") {
           if (data.password.toString().trim() === ADMIN_RESET_PASSWORD) return resJSON({ status: "success", message: "Chào Admin!" });
      return resJSON({ status: "error", message: "Sai mật khẩu!" });
    }

    // 7. LƯU TỪ WORD (uploadWord)
    if (action === "uploadWord") {
      const sheetExams = ss.getSheetByName("Exams") || ss.insertSheet("Exams");
      const sheetBank = ss.getSheetByName("QuestionBank") || ss.insertSheet("QuestionBank");
      sheetExams.appendRow([data.config.title, data.idNumber, data.config.duration, data.config.minTime, data.config.tabLimit, JSON.stringify(data.config.points)]);
      data.questions.forEach(function (q) { sheetBank.appendRow([data.config.title, q.part, q.type, q.classTag, q.question, q.answer, q.image]); });
      return createResponse("success", "UPLOAD_DONE");
    }

    // 8. NHÁNH THEO TYPE (quiz, rating, ketqua)
    if (data.type === 'rating') {
      let sheetRate = ss.getSheetByName("danhgia") || ss.insertSheet("danhgia");
      sheetRate.appendRow([new Date(), data.stars, data.name, data.class, data.idNumber, data.comment || "", data.taikhoanapp]);
      return createResponse("success", "Đã nhận đánh giá");
    }
    if (data.type === 'quiz') {
      let sheetQuiz = ss.getSheetByName("ketquaQuiZ") || ss.insertSheet("ketquaQuiZ");
      sheetQuiz.appendRow([new Date(), data.examCode || "QUIZ", data.name || "N/A", data.className || "", data.school || "", data.phoneNumber || "", data.score || 0, data.totalTime || "00:00", data.stk || "", data.bank || ""]);
      return createResponse("success", "Đã lưu kết quả Quiz");
    }

    // 9. LƯU KẾT QUẢ THI TỔNG HỢP (Mặc định nếu có data.examCode)
    if (data.examCode) {
      let sheetResult = ss.getSheetByName("ketqua") || ss.insertSheet("ketqua");
      sheetResult.appendRow([new Date(), data.examCode, data.sbd, data.name, data.className, data.score, data.totalTime, JSON.stringify(data.details)]);
      return createResponse("success", "Đã lưu kết quả thi");
    }
    return createResponse("error", "Không khớp lệnh nào!");

  }
  catch (err) {
    return createResponse("error", err.toString());
  } finally {
    lock.releaseLock();
  }
}
// ====================================================================CÁC HÀM PHỤ TRỢ (Để hết vào đây)
function getLinkFromRouting(idNumber) {
  const sheet = ss.getSheetByName("idgv");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // Cột A: idNumber, Cột C: linkscript
    if (data[i][0].toString().trim() === idNumber.toString().trim()) {
      return data[i][2].toString().trim();
    }
  }
  return null;
}

function getSpreadsheetByTarget(targetId) {
  if (!targetId) return ss;
  
  const sheet = ss.getSheetByName("idgv");
  const rows = sheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    // Cột A: idNumber, Cột C: linkscript (URL Spreadsheet)
    if (rows[i][0].toString().trim() === targetId.toString().trim()) {
      let url = rows[i][2].toString().trim();
      if (url && url.startsWith("http")) {
        try {
          // Nếu link chính là file Master thì không cần mở lại
          if (url.indexOf(SPREADSHEET_ID) !== -1) return ss;
          return SpreadsheetApp.openByUrl(url);
        } catch (e) {
          console.log("Không thể mở link riêng của GV, dùng file Master làm mặc định.");
        }
      }
      break;
    }
  }
  return ss; 
}

function replaceIdInBlock(block, newId) {
  if (block.match(/id\s*:\s*\d+/)) return block.replace(/id\s*:\s*\d+/, "id: " + newId);
  return block.replace("{", "{\nid: " + newId + ",");
}


function getAppConfig() {
  var sheetCD = ssAdmin.getSheetByName("dangcd");
  var dataCD = sheetCD.getDataRange().getValues();

  var topics = [];
  var classesMap = {}; // Dùng để lọc danh sách lớp không trùng lặp

  // Chạy từ dòng 2 (bỏ tiêu đề)
  for (var i = 1; i < dataCD.length; i++) {
    var lop = dataCD[i][0];   // Cột A: lop
    var idcd = dataCD[i][1];  // Cột B: idcd
    var namecd = dataCD[i][2]; // Cột C: namecd

    if (lop) {
      // 1. Đẩy vào danh sách chuyên đề
      topics.push({
        grade: lop,
        id: idcd,
        name: namecd
      });

      // 2. Thu thập danh sách lớp (để nạp vào CLASS_ID bên React)
      // Ví dụ: Trong sheet có lớp 10, 11, 12 thì CLASS_ID sẽ có các lớp tương ứng
      classesMap[lop] = true;
    }
  }

  return {
    topics: topics,
    classes: Object.keys(classesMap).sort(function (a, b) { return a - b; }) // Trả về [9, 10, 11, 12] chẳng hạn
  };
}

function getAppConfigmt() {
  try {
    // Lưu ý: Đảm bảo ssAdmin đã được khai báo ở đầu script của bạn
    var sheetCD = ssAdmin.getSheetByName("dangcd");
    if (!sheetCD) return { topics: [] };

    var dataCD = sheetCD.getDataRange().getValues();
    var topics = [];

    // Chạy từ dòng 2 (bỏ tiêu đề)
    for (var i = 1; i < dataCD.length; i++) {
      var lop = dataCD[i][0];    // Cột A: lop
      var idcd = dataCD[i][1];   // Cột B: idcd
      var namecd = dataCD[i][2]; // Cột C: namecd

      if (idcd) {
        topics.push({
          grade: lop,            // Khối lớp (10, 11, 12)
          id: String(idcd),      // ID chuyên đề (để lưu vào matrix)
          name: String(namecd)   // Tên để hiển thị cho GV chọn
        });
      }
    }

    return { topics: topics };
  } catch (e) {
    return { topics: [], error: e.toString() };
  }
}
// Reset QuiZ
 function resetQuizData(password) {
  if (password !== ADMIN_RESET_PASSWORD) {
    return createResponse("error", "Sai mật khẩu!");
  }

  const sheet = ss.getSheetByName("ketquaQuiZ");

  if (!sheet) {
    return createResponse("error", "Không tìm thấy sheet ketquaQuiZ");
  }

  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  return createResponse("success", "Đã reset Quiz admin2");
}
function createResponseW(status, message, data) {
  var output = JSON.stringify({
    status: status,
    message: message,
    data: data || null
  });
  
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
function createResponse(status, message, data) {
  var output = JSON.stringify({
    status: status,
    message: message,
    data: data || null
  });
  
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

// Giữ lại resJSON để phục vụ các đoạn code cũ đang gọi tên này
function resJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==== Hàm lấy tên sheet
function getsheetname(mon) {
  var monId = String(mon || "").toLowerCase().trim();
  const cleanMon = clean(monId);
  const sheetName = "NH" + cleanMon;

  var sheet = ssAdmin.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Không tìm thấy sheet: " + sheetName);
    return null;
  }
  return sheet;
}

// chuẩn hóa text
function clean(text, fallback = "") {
  return String(text ?? fallback)
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " "); // gom nhiều space thành 1
}

function parseQuestionFromCell(text, id) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const qLine = lines.find(l => l.startsWith('?'));
  const question = qLine ? qLine.slice(1).trim() : '';
  const options = lines.filter(l => /^[A-D]\./.test(l)).map(l => l.slice(2).trim());
  const ansLine = lines.find(l => l.startsWith('='));
  const ansIndex = ansLine ? ansLine.replace('=', '').trim().charCodeAt(0) - 65 : -1;
  return { id, type: 'mcq', question, o: options, a: options[ansIndex] || '' };
}
