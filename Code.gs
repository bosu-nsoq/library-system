function onOpen() { 
  SpreadsheetApp.getUi().createMenu('📚 WNTK Library').addItem('เปิดโปรแกรม', 'showModal').addToUi(); 
}

function showModal() { 
  var html = HtmlService.createHtmlOutputFromFile('Page').setWidth(450).setHeight(800); 
  SpreadsheetApp.getUi().showModalDialog(html, 'WNTK Library System'); 
}

function doGet(e) { 
  return HtmlService.createHtmlOutputFromFile('Page')
    .setTitle("WNTK Library System")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

// --- 1. Login System & Role Management ---
// ตรวจสอบและเพิ่มคอลัมน์ Role ใน Sheet 'Staff' หลังคอลัมน์ Name (คอลัมน์ D)
function ensureStaffSheetRoleColumn(sheet) {
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  if (lastCol < 4) {
    sheet.getRange(1, 4).setValue("Role");
    sheet.getRange(1, 4).setFontWeight("bold").setBackground("#D8F3DC");
  }
  
  if (lastRow > 1) {
    const roles = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    let updated = false;
    for (let r = 0; r < roles.length; r++) {
      if (!roles[r][0] || String(roles[r][0]).trim() === "") {
        roles[r][0] = "admin"; // กำหนดให้ทุกคนมีบทบาทเริ่มต้นเป็น admin ตามคำสั่งของผู้ใช้
        updated = true;
      }
    }
    if (updated) {
      sheet.getRange(2, 4, lastRow - 1, 1).setValues(roles);
    }
  }
}

function loginStaff(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Staff");
  if (!sheet) return { status: "error", msg: "ไม่พบ Sheet 'Staff'" };
  
  ensureStaffSheetRoleColumn(sheet);
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(pin).trim()) {
      checkAndAutoBackupOnLogin();
      const rawRole = String(data[i][3] || 'admin').trim().toLowerCase();
      const isStaff = (rawRole === 'staff' || rawRole === 'youth' || rawRole === 'ยุวบรรณารักษ์');
      const role = isStaff ? 'staff' : 'admin';
      return { 
        status: "success", 
        name: data[i][2],
        role: role,
        libraryData: getFullLibraryData() // โหลดข้อมูลทั้งหมดลงเครื่องทันทีที่ล็อกอิน
      };
    }
  }
  return { status: "fail" };
}

// ฟังก์ชันเพิ่มยุวบรรณารักษ์ (สำหรับครูบรรณารักษ์ / Admin)
function addYouthStaff(pin, name) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Staff");
    if (!sheet) {
      sheet = ss.insertSheet("Staff");
      sheet.appendRow(["ID", "PIN", "Name", "Role"]);
      sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#D8F3DC");
    }
    
    ensureStaffSheetRoleColumn(sheet);
    
    pin = String(pin || '').trim();
    name = String(name || '').trim();
    
    if (!pin) return { status: "error", msg: "กรุณาระบุรหัสใช้งาน (PIN)" };
    if (!name) return { status: "error", msg: "กรุณาระบุชื่อยุวบรรณารักษ์" };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === pin) {
        return { status: "error", msg: `รหัสใช้งาน (PIN) นี้มีผู้ใช้งานแล้ว (${data[i][2]})` };
      }
    }
    
    const nextNum = Math.max(1, data.length);
    const nextId = "ST-" + String(nextNum).padStart(3, '0');
    
    sheet.appendRow([nextId, pin, name, "staff"]);
    return { status: "success", msg: `เพิ่มยุวบรรณารักษ์ "${name}" สำเร็จ!` };
  } catch(e) {
    return { status: "error", msg: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

// --- 2. History & Print Data ---
function getRecentLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Logs");
  const bSheet = ss.getSheetByName("Books");
  const lastRow = logSheet.getLastRow();
  const numRows = 25;
  if (lastRow < 2) return []; 
  const startRow = Math.max(2, lastRow - numRows + 1);
  const data = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();

  // สร้าง map ชื่อหนังสือเพื่อดึงชื่อเรื่องมาแสดงในประวัติ
  let bookNameMap = {};
  if (bSheet) {
    const bData = bSheet.getDataRange().getValues();
    for (let b = 1; b < bData.length; b++) {
      const bid = String(bData[b][0] || '').trim();
      if (bid) {
        bookNameMap[bid] = String(bData[b][1] || '').trim();
      }
    }
  }

  let logs = [];
  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let d = new Date(row[0]);
    let dateStr = !isNaN(d.getTime()) ? (Utilities.formatDate(d, "GMT+7", "dd/MM/") + (d.getFullYear() + 543)) : "";
    let timeStr = !isNaN(d.getTime()) ? (d.getHours().toString().padStart(2, '0') + ":" + d.getMinutes().toString().padStart(2, '0')) : "";
    
    // กำหนดส่ง (dueDate)
    let dueDateStr = "";
    if (row[5]) {
      let dd = new Date(row[5]);
      if (!isNaN(dd.getTime())) {
        dueDateStr = Utilities.formatDate(dd, "GMT+7", "dd/MM/") + (dd.getFullYear() + 543);
      } else {
        dueDateStr = String(row[5]).trim();
      }
    }

    const bookId = String(row[4] || '').trim();
    const bookTitle = bookNameMap[bookId] || bookId;

    logs.push({
      rowIndex: startRow + i,
      date: dateStr,
      time: timeStr,
      action: String(row[1] || '').trim(),
      studentId: String(row[2] || '').trim(),
      studentName: String(row[3] || '').trim(),
      bookId: bookId,
      bookName: bookTitle,
      dueDate: dueDateStr,
      note: String(row[6] || '').trim(),
      staffName: String(row[7] || '').trim()
    });
  }
  return logs.reverse();
}

function deleteSpecificLog(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Logs");
  try {
    const lastRow = logSheet.getLastRow();
    if (rowIndex > lastRow || rowIndex < 2) return "error:❌ ไม่พบรายการนี้";
    
    const rowData = logSheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
    const action = rowData[1];
    const studentId = String(rowData[2]).trim();
    const note = String(rowData[6]);
    
    if (action === "คืน" && studentId !== "") {
       const studentSheet = ss.getSheetByName("Students");
       const sRow = findRow(studentSheet, studentId);
       if (sRow !== -1) {
          let currentScore = studentSheet.getRange(sRow, 4).getValue() || 0;
          
          if (note.includes("+10")) {
             currentScore -= 10;
          } else if (note.includes("หักคะแนน")) { 
             currentScore += 5;
          }
          
          if (currentScore < 0) currentScore = 0;
          studentSheet.getRange(sRow, 4).setValue(currentScore);
       }
    }

    logSheet.deleteRow(rowIndex);
    return "success:🗑️ ลบรายการและปรับปรุงคะแนนเรียบร้อย";
  } catch(e) { return "error:❌ เกิดข้อผิดพลาด"; }
}

// --- 3. Transaction (รองรับหลายเล่ม + ระบบจำนวนหนังสือ + แสดงชื่อผู้ยืมในชีต) ---
function processMultipleTransactions(studentId, bookIds, action, staffName) {
  if (!Array.isArray(bookIds) || bookIds.length === 0) return "error:❌ กรุณาสแกนรหัสให้ครบ";
  if (bookIds.length === 1) return processTransaction(studentId, bookIds[0], action, staffName);

  let successCount = 0;
  for (let i = 0; i < bookIds.length; i++) {
    let bid = String(bookIds[i]).trim();
    if (!bid) continue;
    let res = processTransaction(studentId, bid, action, staffName);
    if (res.startsWith("success:")) successCount++;
  }
  const actionText = (action === 'borrow') ? 'ยืม' : 'คืน';
  if (successCount === bookIds.length) {
    return `success:✅ ${actionText}สำเร็จ ${successCount} เล่ม`;
  } else if (successCount > 0) {
    return `warning:⚠️ ${actionText}สำเร็จ ${successCount}/${bookIds.length} เล่ม`;
  } else {
    return `error:❌ ทำรายการไม่สำเร็จ`;
  }
}

function processTransaction(studentId, bookId, action, staffName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookSheet = ss.getSheetByName("Books");
  const studentSheet = ss.getSheetByName("Students");
  const logSheet = ss.getSheetByName("Logs");
  studentId = String(studentId).trim(); bookId = String(bookId).trim();
  if (studentId == "" || bookId == "") return "error:❌ กรุณาสแกนรหัสให้ครบ";
  const sRow = findRow(studentSheet, studentId);
  if (sRow == -1) return "error:❌ ไม่พบรหัสนักเรียน (" + studentId + ")";
  const bRow = findRow(bookSheet, bookId);
  if (bRow == -1) return "error:❌ ไม่พบรหัสหนังสือ (" + bookId + ")";
  
  const studentName = String(studentSheet.getRange(sRow, 2).getValue()).trim();
  const totalQty = Number(bookSheet.getRange(bRow, 6).getValue()) || 1;
  const borrowersStr = String(bookSheet.getRange(bRow, 5).getValue()).trim();
  let borrowers = borrowersStr ? borrowersStr.split(",").map(s => s.trim()).filter(s => s !== "") : [];
  const available = totalQty - borrowers.length;
  const today = new Date();
  
  if (action == "borrow") {
    let autoReturnNote = "";
    if (available <= 0) {
      // ⚡ หากมีการยืมหนังสือที่ยังไม่ได้ทำการคืน ให้ระบบทำการคืน (คิด/หักคะแนนคนยืมคนแรกก่อน) แล้วค่อยทำรายการยืมสำหรับคนที่สองทีหลัง
      if (borrowers.indexOf(studentName) !== -1 || borrowers.indexOf(studentId) !== -1) {
        return "warning:⚠️ นักเรียนคนนี้ (" + studentName + ") กำลังยืมหนังสือเล่มนี้อยู่แล้ว";
      }

      let prevBorrowerIdent = borrowers.length > 0 ? borrowers[0] : "";
      let prevStudentId = "";
      let prevStudentName = prevBorrowerIdent;
      let prevDueDate = null;

      const lastBorrowLog = findLastBorrowLogOptimized(logSheet, bookId, "");
      if (lastBorrowLog) {
        prevDueDate = lastBorrowLog.dueDate ? new Date(lastBorrowLog.dueDate) : null;
        if (lastBorrowLog.studentId) prevStudentId = lastBorrowLog.studentId;
        if (lastBorrowLog.studentName) prevStudentName = lastBorrowLog.studentName;
      }

      let prevSRow = -1;
      if (prevStudentId) prevSRow = findRow(studentSheet, prevStudentId);
      if (prevSRow === -1 && prevStudentName) {
        const sData = studentSheet.getDataRange().getValues();
        for (let j = 1; j < sData.length; j++) {
          if (String(sData[j][1]).trim() === prevStudentName) {
            prevSRow = j + 1;
            prevStudentId = String(sData[j][0]).trim();
            break;
          }
        }
      }

      let prevScoreMsg = "";
      if (prevSRow !== -1) {
        let prevScore = Number(studentSheet.getRange(prevSRow, 4).getValue()) || 0;
        if (prevDueDate && !isNaN(prevDueDate.getTime())) {
          const checkDate = new Date(); checkDate.setHours(0,0,0,0);
          const dDate = new Date(prevDueDate); dDate.setHours(0,0,0,0);
          if (checkDate <= dDate) {
            prevScore += 10;
            prevScoreMsg = "(+10 คะแนน)";
          } else {
            prevScore -= 5;
            prevScoreMsg = "(คืนช้า หักคะแนน -5)";
          }
        } else {
          prevScore += 10;
          prevScoreMsg = "(+10 คะแนน)";
        }
        studentSheet.getRange(prevSRow, 4).setValue(Math.max(0, prevScore));
      }

      logSheet.appendRow([today, "คืน", prevStudentId, prevStudentName, bookId, "", "คืนอัตโนมัติ(ส่งต่อ) " + prevScoreMsg, staffName]);

      if (borrowers.length > 0) {
        borrowers.shift();
      }
      autoReturnNote = `[คืนจาก ${prevStudentName || 'ผู้ยืมเดิม'} ${prevScoreMsg}] `;
    }

    if (countHolding(bookSheet, studentId, studentName) >= 3) return "warning:⚠️ ยืมเกินโควตา 3 เล่ม";
    
    // ✅ บันทึกเป็นชื่อผู้ยืมในชีต Books
    borrowers.push(studentName);
    const newAvailable = totalQty - borrowers.length;
    bookSheet.getRange(bRow, 4).setValue(newAvailable > 0 ? "ว่าง" : "ถูกยืม");
    bookSheet.getRange(bRow, 5).setValue(borrowers.join(", "));
    
    const dueDate = new Date(); dueDate.setDate(today.getDate() + 3); 
    logSheet.appendRow([today, "ยืม", studentId, studentName, bookId, dueDate, "", staffName]);
    
    return "success:✅ " + autoReturnNote + "ยืมสำเร็จ!";
    
  } else if (action == "return") {
    const lastLog = findLastBorrowLogOptimized(logSheet, bookId, studentId); 
    let newScore = studentSheet.getRange(sRow, 4).getValue() || 0;
    let msg = "";
    if (lastLog) {
       const dueDate = new Date(lastLog.dueDate);
       const checkDate = new Date(); checkDate.setHours(0,0,0,0); dueDate.setHours(0,0,0,0);
       if (checkDate <= dueDate) { newScore += 10; msg = "(+10 คะแนน)"; } 
       else { newScore -= 5; msg = "(คืนช้า หักคะแนน)"; }
       studentSheet.getRange(sRow, 4).setValue(newScore);
    }
    
    // ✅ ลบชื่อผู้ยืมออก (และรองรับรายการเก่าที่เคยเก็บเป็นรหัสด้วย)
    let idx = borrowers.indexOf(studentName);
    if (idx === -1) idx = borrowers.indexOf(studentId);
    if (idx !== -1) {
       borrowers.splice(idx, 1);
    } else if (borrowers.length > 0) {
       borrowers.pop();
    }
    
    const newAvailable = totalQty - borrowers.length;
    bookSheet.getRange(bRow, 4).setValue(newAvailable > 0 ? "ว่าง" : "ถูกยืม");
    bookSheet.getRange(bRow, 5).setValue(borrowers.join(", "));
    logSheet.appendRow([new Date(), "คืน", studentId, studentName, bookId, "", msg, staffName]);
    return "success:✅ คืนสำเร็จ! " + msg;
  }
}

// --- 4. Rewards & Stats ---
function processRedemption(studentId, rewardId, cost, rewardName, staffName) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const sSheet = ss.getSheetByName("Students");
   const logSheet = ss.getSheetByName("Logs");
   const sRow = findRow(sSheet, studentId);
   if (sRow == -1) return "error:❌ ไม่พบข้อมูลนักเรียน";
   let currentScore = sSheet.getRange(sRow, 4).getValue() || 0;
   if (currentScore < cost) return "error:❌ คะแนนไม่พอแลก!";
   sSheet.getRange(sRow, 4).setValue(currentScore - cost);
   logSheet.appendRow([new Date(), "แลกของรางวัล", studentId, sSheet.getRange(sRow, 2).getValue(), rewardId, "", rewardName + " (-"+cost+")", staffName]);
   return "success:✅ แลกสำเร็จ!";
}

function checkDateMatch(rowDate, today, range) {
    if (!rowDate) return false;
    if (typeof range === 'object') {
       if (range.type === 'custom_day') {
           let val = String(range.value).trim();
           let y, m, d;
           if (val.includes('/')) {
               let parts = val.split('/');
               d = parseInt(parts[0], 10);
               m = parseInt(parts[1], 10) - 1;
               y = parseInt(parts[2], 10);
           } else {
               let parts = val.split('-');
               y = parseInt(parts[0], 10);
               m = parseInt(parts[1], 10) - 1;
               d = parseInt(parts[2], 10);
           }
           if (y > 2400) y -= 543;
           return (rowDate.getDate() == d && rowDate.getMonth() == m && rowDate.getFullYear() == y);
       } else if (range.type === 'custom_month') {
           let parts = range.value.split('-');
           let y = parseInt(parts[0], 10);
           if (y > 2400) y -= 543;
           return (rowDate.getMonth() == parseInt(parts[1], 10) - 1 && rowDate.getFullYear() == y);
       } else if (range.type === 'custom_year') {
           let y = parseInt(range.value, 10);
           if (y > 2400) y -= 543;
           return (rowDate.getFullYear() == y);
       }
    } else {
       if (range === 'day') return (rowDate.getDate()===today.getDate() && rowDate.getMonth()===today.getMonth() && rowDate.getFullYear()===today.getFullYear());
       else if (range === 'month') return (rowDate.getMonth()===today.getMonth() && rowDate.getFullYear()===today.getFullYear());
       else if (range === 'year') return (rowDate.getFullYear()===today.getFullYear());
       else return true;
    }
    return false;
}

function getAdvancedStats(range) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logData = ss.getSheetByName("Logs").getDataRange().getValues();
  const bookData = ss.getSheetByName("Books").getDataRange().getValues();
  let bookMap = {};
  for(let i=1; i<bookData.length; i++) { bookMap[String(bookData[i][0]).trim()] = bookData[i][1]; }
  const today = new Date();
  
  let borrows = 0, returns = 0;
  let studentCount = {};
  let bookCount = {};

  for(let i=1; i<logData.length; i++) {
    let rowDate = new Date(logData[i][0]);
    if (!logData[i][0] || isNaN(rowDate.getTime())) continue;
    let action = logData[i][1];
    let sName = logData[i][3];
    let bId = String(logData[i][4]).trim();
    
    if (checkDateMatch(rowDate, today, range)) {
       if (action === 'ยืม') {
           borrows++;
           if (sName) studentCount[sName] = (studentCount[sName]||0) + 1;
           if (bId) bookCount[bId] = (bookCount[bId]||0) + 1;
       } else if (action === 'คืน' || action === 'คืน(ระบบตัด)') {
           returns++;
       }
    }
  }
  
  let topStudents = Object.keys(studentCount).map(k => ({name: k, count: studentCount[k]})).sort((a,b)=>b.count-a.count).slice(0,3);
  let topBooks = Object.keys(bookCount).map(k => ({name: bookMap[k] || k, count: bookCount[k]})).sort((a,b)=>b.count-a.count).slice(0,3);
  
  return { borrows, returns, topStudents, topBooks };
}

function getLogsByRange(range) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const logData = ss.getSheetByName("Logs").getDataRange().getValues();
   const bookData = ss.getSheetByName("Books").getDataRange().getValues();
   
   let bookMap = {};
   if(bookData.length > 1) {
      for(let j=1; j<bookData.length; j++) { bookMap[String(bookData[j][0]).trim()] = bookData[j][1]; }
   }

   if(logData.length < 2) return [];
   const today = new Date();
   let result = [];
   
   for(let i=1; i<logData.length; i++) {
      let d = new Date(logData[i][0]);
      if (!logData[i][0] || isNaN(d.getTime())) continue;

      if(checkDateMatch(d, today, range)) {
         let bId = String(logData[i][4]).trim();
         let dBE = d.getFullYear() + 543;
         let dateBE = Utilities.formatDate(d, "GMT+7", "dd/MM/") + dBE;
         result.push({
           date: dateBE,
           time: Utilities.formatDate(d, "GMT+7", "HH:mm"),
           action: logData[i][1],
           studentId: logData[i][2],
           studentName: logData[i][3],
           bookId: bId,
           bookName: bookMap[bId] || bId, 
           staff: logData[i][7] || ''
         });
      }
   }
   return result.reverse(); 
}

// --- Helpers ---
function findRow(sheet, id) { const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(id).trim()) return i + 1; } return -1; }
function getAllStudentNames() { const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students").getDataRange().getValues(); let map = {}; for (let i = 1; i < data.length; i++) { map[String(data[i][0]).trim()] = data[i][1]; } return map; }

function getNames(studentId, bookId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  let sInfo = null, bInfo = null;

  // 1. ดึงข้อมูลนักเรียน (ประมวลผลบนหน่วยความจำ รวดเร็ว)
  if (studentId) { 
    const sSheet = ss.getSheetByName("Students");
    if (sSheet) {
      const sData = sSheet.getDataRange().getValues();
      const targetSid = String(studentId).trim();
      for (let i = 1; i < sData.length; i++) {
        if (String(sData[i][0]).trim() === targetSid) {
          sInfo = { 
            name: String(sData[i][1] || '').trim(), 
            grade: String(sData[i][2] || '').trim(), 
            score: Number(sData[i][3]) || 0 
          };
          break;
        }
      }
    }
  }

  // 2. ดึงข้อมูลหนังสือ พร้อมคำนวณจำนวนวันการยืม และผู้ยืมคนล่าสุด
  if (bookId) { 
    const bookSheet = ss.getSheetByName("Books");
    if (bookSheet) {
      const bData = bookSheet.getDataRange().getValues();
      const targetBid = String(bookId).trim();
      for (let i = 1; i < bData.length; i++) {
        if (String(bData[i][0]).trim() === targetBid) {
          const bName = String(bData[i][1] || '').trim();
          const sheetStatus = String(bData[i][3] || '').trim();
          const totalQty = Math.max(1, Number(bData[i][5]) || 1);
          const borrowersStr = String(bData[i][4] || '').trim();
          const sheetBorrowers = borrowersStr ? borrowersStr.split(",").map(s => s.trim()).filter(Boolean) : [];

          // ตรวจสอบประวัติจาก Logs ตามลำดับเวลา เพื่อหาสถานะยืมปัจจุบันและผู้ยืมล่าสุด
          let activeBorrows = [];
          let latestBorrower = "";
          let latestBorrowDate = null;
          const logSheet = ss.getSheetByName("Logs");
          if (logSheet) {
            const logData = logSheet.getDataRange().getValues();
            for (let j = 1; j < logData.length; j++) {
              if (String(logData[j][4]).trim() === targetBid) {
                const action = String(logData[j][1] || '').trim();
                const sId = String(logData[j][2] || '').trim();
                const sName = String(logData[j][3] || '').trim();
                const rawDate = logData[j][0];
                if (sName) latestBorrower = sName;
                if (rawDate) latestBorrowDate = new Date(rawDate);

                if (action === "ยืม") {
                  activeBorrows.push({ studentId: sId, studentName: sName, date: new Date(rawDate) });
                } else if (action === "คืน") {
                  let idx = activeBorrows.findIndex(b => b.studentId === sId || b.studentName === sName);
                  if (idx !== -1) activeBorrows.splice(idx, 1);
                  else if (activeBorrows.length > 0) activeBorrows.shift();
                }
              }
            }
          }

          if (sheetBorrowers.length > 0) {
            latestBorrower = sheetBorrowers[sheetBorrowers.length - 1];
          }

          const isCurrentlyBorrowed = (sheetStatus === "ถูกยืม") || (activeBorrows.length > 0) || (sheetBorrowers.length >= totalQty);
          let borrowDays = 0;
          let borrowDaysText = "";

          const today = new Date();
          today.setHours(0,0,0,0);

          if (activeBorrows.length > 0) {
            const lastActive = activeBorrows[activeBorrows.length - 1];
            if (lastActive.studentName) latestBorrower = lastActive.studentName;
            if (lastActive.date && !isNaN(lastActive.date.getTime())) {
              const bDate = new Date(lastActive.date);
              bDate.setHours(0,0,0,0);
              borrowDays = Math.max(0, Math.round((today.getTime() - bDate.getTime()) / (1000 * 60 * 60 * 24)));
              borrowDaysText = borrowDays === 0 ? "ยืมวันนี้" : `ยืมไปแล้ว ${borrowDays} วัน`;
            }
          } else if (latestBorrowDate && !isNaN(latestBorrowDate.getTime()) && isCurrentlyBorrowed) {
            const bDate = new Date(latestBorrowDate);
            bDate.setHours(0,0,0,0);
            borrowDays = Math.max(0, Math.round((today.getTime() - bDate.getTime()) / (1000 * 60 * 60 * 24)));
            borrowDaysText = borrowDays === 0 ? "ยืมวันนี้" : `ยืมไปแล้ว ${borrowDays} วัน`;
          }

          let available = 0;
          let statusText = "";

          if (isCurrentlyBorrowed) {
            const effectiveCount = Math.max(sheetBorrowers.length, activeBorrows.length, 1);
            available = Math.max(0, totalQty - effectiveCount);
            if (available > 0) {
              statusText = `ว่าง จำนวน ${available} เล่ม`;
            } else {
              statusText = `ถูกยืม (ว่าง 0 เล่ม)`;
            }
          } else {
            available = totalQty;
            statusText = `ว่าง จำนวน ${available} เล่ม`;
            latestBorrower = "";
          }

          let currentBorrowers = [];
          if (activeBorrows.length > 0) {
            currentBorrowers = activeBorrows.map(b => b.studentName || b.studentId).filter(Boolean);
            for (let s of sheetBorrowers) {
              if (!currentBorrowers.includes(s) && currentBorrowers.length < totalQty) {
                currentBorrowers.push(s);
              }
            }
          } else if (sheetBorrowers.length > 0) {
            currentBorrowers = sheetBorrowers.slice();
          }

          let copies = [];
          if (totalQty > 1) {
            for (let k = 1; k <= totalQty; k++) {
              if (k <= currentBorrowers.length) {
                const bName = currentBorrowers[k - 1];
                copies.push({
                  copyNum: k,
                  borrower: bName,
                  isAvailable: false,
                  text: `เล่มที่ ${k}: อยู่ที่ ${bName}`
                });
              } else {
                copies.push({
                  copyNum: k,
                  borrower: "",
                  isAvailable: true,
                  text: `เล่มที่ ${k}: ว่าง`
                });
              }
            }
          }

          bInfo = { 
            name: bName, 
            status: statusText, 
            available: available, 
            totalQty: totalQty,
            copies: copies,
            borrowerName: latestBorrower, 
            borrowDays: borrowDays, 
            borrowDaysText: borrowDaysText
          };
          break;
        }
      }
    }
  }

  return {student: sInfo, book: bInfo};
}

function countHolding(sheet, studentId, studentName) { 
  const data = sheet.getRange("E2:E").getValues(); 
  let count = 0; 
  const targetId = String(studentId).trim();
  const targetName = studentName ? String(studentName).trim() : "";
  for (let i = 0; i < data.length; i++) { 
    let cell = String(data[i][0]).trim();
    if (!cell) continue;
    let list = cell.split(",").map(s => s.trim());
    list.forEach(item => { 
      if (item === targetId || (targetName && item === targetName)) count++; 
    });
  } 
  return count; 
}

function findLastBorrowLogOptimized(sheet, bookId, studentId) { 
  const data = sheet.getDataRange().getValues(); 
  for (let i = data.length - 1; i >= 0; i--) { 
    if (String(data[i][4]).trim() == String(bookId).trim() && data[i][1] == "ยืม") {
      if (!studentId || String(data[i][2]).trim() == String(studentId).trim()) {
        return { 
          dueDate: data[i][5],
          studentId: String(data[i][2]).trim(),
          studentName: String(data[i][3]).trim(),
          rowIndex: i + 1
        }; 
      }
    }
  } 
  return null; 
}

function formatDate(date) { 
  let d = new Date(date); 
  if (isNaN(d.getTime())) return "";
  let day = String(d.getDate()).padStart(2, '0');
  let month = String(d.getMonth() + 1).padStart(2, '0');
  let year = d.getFullYear() + 543;
  return `${day}/${month}/${year}`; 
}

function addNewBook(id, name, author, qty) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Books");
  const addQty = Math.max(1, parseInt(qty, 10) || 1);
  const existingRow = findRow(sheet, id);
  
  if (existingRow != -1) {
    const currentTotal = Number(sheet.getRange(existingRow, 6).getValue()) || 1;
    const newTotal = currentTotal + addQty;
    sheet.getRange(existingRow, 6).setValue(newTotal);
    
    const borrowersStr = String(sheet.getRange(existingRow, 5).getValue()).trim();
    const borrowers = borrowersStr ? borrowersStr.split(",").map(s => s.trim()).filter(s => s !== "") : [];
    if (newTotal - borrowers.length > 0) {
      sheet.getRange(existingRow, 4).setValue("ว่าง");
    }
    return `success:✅ เพิ่มจำนวนหนังสือรหัสเดิมเรียบร้อย (+${addQty} เล่ม รวม ${newTotal} เล่ม)`;
  }
  
  sheet.appendRow([id, name, author, "ว่าง", "", addQty]);
  return `success:✅ เพิ่มหนังสือเรียบร้อย (${addQty} เล่ม)`;
}

function getStudentScoreAndRewards(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sSheet = ss.getSheetByName("Students");
  const rSheet = ss.getSheetByName("Rewards");
  if (!sSheet || !rSheet) return { error: "❌ ไม่พบ Sheet 'Students' หรือ 'Rewards'" };
  
  const targetId = String(studentId).trim();
  const sData = sSheet.getDataRange().getValues();
  let sRowData = null;
  for (let i = 1; i < sData.length; i++) {
    if (String(sData[i][0]).trim() === targetId) {
      sRowData = sData[i];
      break;
    }
  }
  if (!sRowData) return { error: "❌ ไม่พบข้อมูลนักเรียน (" + targetId + ")" };
  
  const rData = rSheet.getDataRange().getValues();
  const rewards = [];
  for (let j = 1; j < rData.length; j++) {
    const rId = String(rData[j][0] || '').trim();
    const rName = String(rData[j][1] || '').trim();
    const rCost = Number(rData[j][2]) || 0;
    if (rId && rName) {
      rewards.push({ id: rId, name: rName, cost: rCost });
    }
  }
  
  return { 
    name: String(sRowData[1] || '').trim(), 
    grade: String(sRowData[2] || '').trim(), 
    score: Number(sRowData[3]) || 0, 
    rewards: rewards 
  };
}

// ฟังก์ชันโหลดข้อมูลทั้งระบบเพื่อทำ In-Memory Cache (ความเร็วระดับ 0.01 วินาที)
function getFullLibraryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sSheet = ss.getSheetByName("Students");
  const bSheet = ss.getSheetByName("Books");
  const lSheet = ss.getSheetByName("Logs");

  let students = {};
  if (sSheet) {
    const sData = sSheet.getDataRange().getValues();
    for (let i = 1; i < sData.length; i++) {
      const id = String(sData[i][0] || '').trim();
      if (id) {
        students[id] = { 
          name: String(sData[i][1] || '').trim(), 
          grade: String(sData[i][2] || '').trim(), 
          score: Number(sData[i][3]) || 0 
        };
      }
    }
  }

  // ติดตามสถานะยืมปัจจุบันและผู้ยืมล่าสุดจาก Logs (ประมวลผลตามลำดับเวลา)
  let activeBorrows = {}; // bid -> [{ studentId, studentName, date }]
  let latestBorrowerMap = {}; // bid -> studentName
  let latestBorrowDateMap = {}; // bid -> Date
  if (lSheet) {
    const lData = lSheet.getDataRange().getValues();
    for (let j = 1; j < lData.length; j++) {
      const action = String(lData[j][1] || '').trim();
      const sId = String(lData[j][2] || '').trim();
      const sName = String(lData[j][3] || '').trim();
      const bid = String(lData[j][4] || '').trim();
      const rawDate = lData[j][0];

      if (!bid) continue;
      if (!activeBorrows[bid]) activeBorrows[bid] = [];
      if (sName) latestBorrowerMap[bid] = sName;
      if (rawDate) latestBorrowDateMap[bid] = new Date(rawDate);

      if (action === "ยืม") {
        activeBorrows[bid].push({
          studentId: sId,
          studentName: sName,
          date: new Date(rawDate)
        });
      } else if (action === "คืน") {
        let idx = activeBorrows[bid].findIndex(b => b.studentId === sId || b.studentName === sName);
        if (idx !== -1) {
          activeBorrows[bid].splice(idx, 1);
        } else if (activeBorrows[bid].length > 0) {
          activeBorrows[bid].shift();
        }
      }
    }
  }

  const today = new Date();
  today.setHours(0,0,0,0);

  let books = {};
  if (bSheet) {
    const bData = bSheet.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      const bid = String(bData[i][0] || '').trim();
      if (bid) {
        const bName = String(bData[i][1] || '').trim();
        const sheetStatus = String(bData[i][3] || '').trim();
        const totalQty = Math.max(1, Number(bData[i][5]) || 1);
        const borrowersStr = String(bData[i][4] || '').trim();
        const sheetBorrowers = borrowersStr ? borrowersStr.split(",").map(s => s.trim()).filter(Boolean) : [];
        
        const activeList = activeBorrows[bid] || [];
        const isCurrentlyBorrowed = (sheetStatus === "ถูกยืม") || (activeList.length > 0) || (sheetBorrowers.length >= totalQty);

        let latestBorrower = "";
        let borrowDays = 0;
        let borrowDaysText = "";

        if (activeList.length > 0) {
          const lastActive = activeList[activeList.length - 1];
          latestBorrower = lastActive.studentName || sheetBorrowers[sheetBorrowers.length - 1] || "";
          if (lastActive.date && !isNaN(lastActive.date.getTime())) {
            const bDate = new Date(lastActive.date);
            bDate.setHours(0,0,0,0);
            borrowDays = Math.max(0, Math.round((today.getTime() - bDate.getTime()) / (1000 * 60 * 60 * 24)));
            borrowDaysText = borrowDays === 0 ? "ยืมวันนี้" : `ยืมไปแล้ว ${borrowDays} วัน`;
          }
        } else if (sheetBorrowers.length > 0) {
          latestBorrower = sheetBorrowers[sheetBorrowers.length - 1];
        } else if (latestBorrowerMap[bid]) {
          latestBorrower = latestBorrowerMap[bid];
        }

        let available = 0;
        let statusText = "";

        if (isCurrentlyBorrowed) {
          const effectiveCount = Math.max(sheetBorrowers.length, activeList.length, 1);
          available = Math.max(0, totalQty - effectiveCount);
          if (available > 0) {
            statusText = `ว่าง จำนวน ${available} เล่ม`;
          } else {
            statusText = `ถูกยืม (ว่าง 0 เล่ม)`;
          }
        } else {
          available = totalQty;
          statusText = `ว่าง จำนวน ${available} เล่ม`;
          latestBorrower = "";
        }

        let currentBorrowers = [];
        if (activeList.length > 0) {
          currentBorrowers = activeList.map(b => b.studentName || b.studentId).filter(Boolean);
          for (let s of sheetBorrowers) {
            if (!currentBorrowers.includes(s) && currentBorrowers.length < totalQty) {
              currentBorrowers.push(s);
            }
          }
        } else if (sheetBorrowers.length > 0) {
          currentBorrowers = sheetBorrowers.slice();
        }

        let copies = [];
        if (totalQty > 1) {
          for (let k = 1; k <= totalQty; k++) {
            if (k <= currentBorrowers.length) {
              const bName = currentBorrowers[k - 1];
              copies.push({
                copyNum: k,
                borrower: bName,
                isAvailable: false,
                text: `เล่มที่ ${k}: อยู่ที่ ${bName}`
              });
            } else {
              copies.push({
                copyNum: k,
                borrower: "",
                isAvailable: true,
                text: `เล่มที่ ${k}: ว่าง`
              });
            }
          }
        }

        books[bid] = {
          name: bName,
          category: String(bData[i][2] || 'ทั่วไป').trim(),
          author: String(bData[i][2] || 'ทั่วไป').trim(),
          status: statusText,
          available: available,
          totalQty: totalQty,
          copies: copies,
          borrowerName: latestBorrower,
          borrowDays: borrowDays,
          borrowDaysText: borrowDaysText
        };
      }
    }
  }

  const rSheet = ss.getSheetByName("Rewards");
  let rewards = [];
  if (rSheet) {
    const rData = rSheet.getDataRange().getValues();
    for (let j = 1; j < rData.length; j++) {
      const rId = String(rData[j][0] || '').trim();
      const rName = String(rData[j][1] || '').trim();
      const rCost = Number(rData[j][2]) || 0;
      if (rId && rName) {
        rewards.push({ id: rId, name: rName, cost: rCost });
      }
    }
  }

  return { students: students, books: books, rewards: rewards };
}

// ฟังก์ชันดึงรายชื่อหนังสือที่ถูกยืมอยู่ขณะนี้ (ยังไม่คืน)
function getCurrentlyBorrowedBooks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookSheet = ss.getSheetByName("Books");
  const logSheet = ss.getSheetByName("Logs");
  
  let bookMap = {};
  if (bookSheet) {
    const bData = bookSheet.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      const bid = String(bData[i][0] || '').trim();
      if (bid) {
        bookMap[bid] = {
          name: String(bData[i][1] || '').trim(),
          status: String(bData[i][3] || '').trim(),
          borrowersStr: String(bData[i][4] || '').trim(),
          totalQty: Math.max(1, Number(bData[i][5]) || 1)
        };
      }
    }
  }

  let activeBorrows = {}; // bid -> [{ studentId, studentName, date }]
  if (logSheet) {
    const lData = logSheet.getDataRange().getValues();
    for (let j = 1; j < lData.length; j++) {
      const action = String(lData[j][1] || '').trim();
      const sId = String(lData[j][2] || '').trim();
      const sName = String(lData[j][3] || '').trim();
      const bid = String(lData[j][4] || '').trim();
      const rawDate = lData[j][0];

      if (!bid) continue;
      if (!activeBorrows[bid]) activeBorrows[bid] = [];

      if (action === "ยืม") {
        activeBorrows[bid].push({
          studentId: sId,
          studentName: sName,
          date: new Date(rawDate)
        });
      } else if (action === "คืน") {
        let idx = activeBorrows[bid].findIndex(b => b.studentId === sId || b.studentName === sName);
        if (idx !== -1) {
          activeBorrows[bid].splice(idx, 1);
        } else if (activeBorrows[bid].length > 0) {
          activeBorrows[bid].shift();
        }
      }
    }
  }

  const today = new Date();
  today.setHours(0,0,0,0);
  let result = [];

  for (let bid in activeBorrows) {
    const list = activeBorrows[bid];
    for (let k = 0; k < list.length; k++) {
      const b = list[k];
      let diffDays = 0;
      let borrowDateThai = "-";
      if (b.date && !isNaN(b.date.getTime())) {
        const d = new Date(b.date);
        d.setHours(0,0,0,0);
        diffDays = Math.max(0, Math.round((today.getTime() - d.getTime()) / (1000 * 60 * 60 * 24)));
        const dBE = b.date.getFullYear() + 543;
        borrowDateThai = Utilities.formatDate(b.date, "GMT+7", "dd/MM/") + dBE;
      }
      result.push({
        bookId: bid,
        bookName: (bookMap[bid] ? bookMap[bid].name : bid) || bid,
        studentId: b.studentId || "-",
        borrowerName: b.studentName || "ไม่ระบุชื่อ",
        days: diffDays,
        borrowDaysText: diffDays === 0 ? "ยืมวันนี้" : `ยืมไปแล้ว ${diffDays} วัน`,
        borrowDateThai: borrowDateThai
      });
    }
  }

  // เผื่อมีรายการที่ถูกยืมในชีต Books แต่ยังไม่มี log ให้แสดงด้วย
  for (let bid in bookMap) {
    const bInfo = bookMap[bid];
    const isRecorded = result.some(r => r.bookId === bid);
    if (!isRecorded && (bInfo.status === "ถูกยืม" || bInfo.borrowersStr !== "")) {
      const borrowers = bInfo.borrowersStr ? bInfo.borrowersStr.split(",").map(s => s.trim()).filter(Boolean) : ["ไม่ระบุชื่อ"];
      borrowers.forEach(borrowerName => {
        result.push({
          bookId: bid,
          bookName: bInfo.name || bid,
          studentId: "-",
          borrowerName: borrowerName,
          days: 0,
          borrowDaysText: "กำลังถูกยืม",
          borrowDateThai: "-"
        });
      });
    }
  }

  result.sort((a, b) => b.days - a.days);
  return result;
}

// ฟังก์ชันลบรายการยืม (บันทึกการคืนหนังสือจากหน้าต่างหนังสือที่ถูกยืม)
function markBookAsReturned(bookId, studentId, borrowerName, staffName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookSheet = ss.getSheetByName("Books");
  const studentSheet = ss.getSheetByName("Students");
  const logSheet = ss.getSheetByName("Logs");
  
  bookId = String(bookId || '').trim();
  studentId = String(studentId || '').trim();
  borrowerName = String(borrowerName || '').trim();
  staffName = String(staffName || 'เจ้าหน้าที่').trim();

  if (!bookId) return "error:❌ ไม่พบรหัสหนังสือ";

  // 1. ค้นหาแถวหนังสือในชีต Books
  let bRow = findRow(bookSheet, bookId);
  let bookName = bookId;
  if (bRow !== -1) {
    bookName = String(bookSheet.getRange(bRow, 2).getValue() || bookId).trim();
    const totalQty = Math.max(1, Number(bookSheet.getRange(bRow, 6).getValue()) || 1);
    const borrowersStr = String(bookSheet.getRange(bRow, 5).getValue()).trim();
    let borrowers = borrowersStr ? borrowersStr.split(",").map(s => s.trim()).filter(s => s !== "") : [];
    
    // ลบชื่อหรือรหัสผู้ยืมออกจากรายการผู้ยืม
    let idx = -1;
    if (borrowerName && borrowerName !== "-" && borrowerName !== "ไม่ระบุชื่อ") {
      idx = borrowers.indexOf(borrowerName);
    }
    if (idx === -1 && studentId && studentId !== "-") {
      idx = borrowers.indexOf(studentId);
    }
    if (idx !== -1) {
      borrowers.splice(idx, 1);
    } else if (borrowers.length > 0) {
      borrowers.pop();
    }

    const newAvailable = Math.max(0, totalQty - borrowers.length);
    bookSheet.getRange(bRow, 4).setValue(newAvailable > 0 ? "ว่าง" : "ถูกยืม");
    bookSheet.getRange(bRow, 5).setValue(borrowers.join(", "));
  }

  // 2. ถ้ามีข้อมูลนักเรียน ให้คำนวณคะแนน
  let sRow = -1;
  let finalStudentName = borrowerName;
  if (studentId && studentId !== "-") {
    sRow = findRow(studentSheet, studentId);
  }
  if (sRow === -1 && borrowerName && borrowerName !== "-" && borrowerName !== "ไม่ระบุชื่อ" && studentSheet) {
    const sData = studentSheet.getDataRange().getValues();
    for (let i = 1; i < sData.length; i++) {
      if (String(sData[i][1]).trim() === borrowerName) {
        sRow = i + 1;
        studentId = String(sData[i][0]).trim();
        break;
      }
    }
  }

  let scoreMsg = "";
  if (sRow !== -1 && studentSheet) {
    finalStudentName = String(studentSheet.getRange(sRow, 2).getValue()).trim();
    const lastLog = findLastBorrowLogOptimized(logSheet, bookId, studentId);
    let newScore = Number(studentSheet.getRange(sRow, 4).getValue()) || 0;
    if (lastLog && lastLog.dueDate) {
      const dueDate = new Date(lastLog.dueDate);
      const checkDate = new Date();
      checkDate.setHours(0,0,0,0);
      dueDate.setHours(0,0,0,0);
      if (checkDate <= dueDate) {
        newScore += 10;
        scoreMsg = "(+10 คะแนน)";
      } else {
        newScore -= 5;
        scoreMsg = "(คืนช้า หักคะแนน)";
      }
      studentSheet.getRange(sRow, 4).setValue(Math.max(0, newScore));
    }
  }

  // 3. บันทึกประวัติการคืนลงในชีต Logs
  if (logSheet) {
    const sIdForLog = (studentId && studentId !== "-") ? studentId : "";
    const sNameForLog = (finalStudentName && finalStudentName !== "-" && finalStudentName !== "ไม่ระบุชื่อ") ? finalStudentName : "ผู้ยืมทั่วไป";
    logSheet.appendRow([new Date(), "คืน", sIdForLog, sNameForLog, bookId, "", scoreMsg, staffName]);
  }

  const freshBorrowed = getCurrentlyBorrowedBooks();
  const freshLibData = getFullLibraryData();

  return {
    status: "success",
    msg: "success:✅ คืนหนังสือสำเร็จเรียบร้อย " + scoreMsg,
    borrowedBooks: freshBorrowed,
    libraryData: freshLibData
  };
}

// ฟังก์ชันดึงรายชื่อนักเรียนทั้งหมดพร้อมชั้นเรียนสำหรับพิมพ์บัตรสมาชิก
function getAllStudentsForCards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sSheet = ss.getSheetByName("Students");
  if (!sSheet) return [];
  const sData = sSheet.getDataRange().getValues();
  let list = [];
  for (let i = 1; i < sData.length; i++) {
    let id = String(sData[i][0] || '').trim();
    let name = String(sData[i][1] || '').trim();
    let grade = String(sData[i][2] || '').trim();
    if (id && name) {
      list.push({ id: id, name: name, grade: grade || '-' });
    }
  }
  return list;
}

// ฟังก์ชันดึงรายการหนังสือทั้งหมดสำหรับพิมพ์สติกเกอร์บาร์โค้ดติดสันหนังสือ
function getAllBooksForLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bSheet = ss.getSheetByName("Books");
  if (!bSheet) return [];
  const bData = bSheet.getDataRange().getValues();
  let list = [];
  for (let i = 1; i < bData.length; i++) {
    let id = String(bData[i][0] || '').trim();
    let name = String(bData[i][1] || '').trim();
    let category = String(bData[i][2] || 'ทั่วไป').trim();
    let totalQty = Math.max(1, Number(bData[i][5]) || 1);
    if (id && name) {
      list.push({ id: id, name: name, category: category, author: category, totalQty: totalQty });
    }
  }
  return list;
}

// ==========================================
// 💾 ระบบสำรองข้อมูลอัตโนมัติ (Automated Backup System)
// ==========================================

const BACKUP_SHEET_NAME = "Backups";
const MAX_BACKUP_RECORDS = 30; // เก็บประวัติสำรองข้อมูล 30 รายการล่าสุด

// ฟังก์ชันแปลง Date เป็น วัน เดือน ปี พ.ศ. พร้อมเวลา (dd/mm/yyyy พ.ศ. HH:mm:ss)
function formatThaiDateTimeBE(date) {
  if (!date) return '-';
  const d = new Date(date);
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const yearBE = d.getFullYear() + 543;
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  const seconds = String(d.getSeconds()).padStart(2, '0');
  return `${day}/${month}/${yearBE} ${hours}:${minutes}:${seconds}`;
}

// ตรวจสอบและดึงชีตสำหรับบันทึกรายการสำรองข้อมูล (Backups)
function getOrCreateBackupLogSheet(ss) {
  let sheet = ss.getSheetByName(BACKUP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BACKUP_SHEET_NAME);
    sheet.appendRow(["วันที่/เวลา", "ชื่อไฟล์สำรอง", "ลิงก์ไฟล์", "รหัสไฟล์", "ประเภท", "สถานะ"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#D8F3DC");
  }
  return sheet;
}

// สร้างไฟล์สำรองข้อมูล (Full Spreadsheet Snapshot) คัดลอกทุกชีต (Books, Students, Logs, Staff)
// โดยจะลบไฟล์สำรองเดิมใน Google Drive และสำรองทับไฟล์เดิมเสมอ
function createLibraryBackup(triggerType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const thaiDateStr = formatThaiDateTimeBE(now);
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const yearBE = now.getFullYear() + 543;
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    const timestamp = `${yearBE}${month}${day}_${hours}${minutes}${seconds}`;
    const backupName = `Backup_WNTK_Library_${timestamp}`;

    // 1. ลบไฟล์สำรองเดิมใน Google Drive และล้างประวัติเดิมเพื่อสำรองทับไฟล์เดิม
    const logSheet = getOrCreateBackupLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      const oldData = logSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (let i = 0; i < oldData.length; i++) {
        const oldFileId = String(oldData[i][3] || '').trim();
        if (oldFileId) {
          try {
            if (typeof DriveApp !== 'undefined') {
              const oldFile = DriveApp.getFileById(oldFileId);
              if (oldFile) oldFile.setTrashed(true);
            }
          } catch(e) {
            console.warn("Could not trash old backup file:", e);
          }
        }
      }
      logSheet.deleteRows(2, lastRow - 1);
    }

    // 2. สร้าง Google Spreadsheet ไฟล์ใหม่สำหรับเก็บสำรองข้อมูล
    const newSs = SpreadsheetApp.create(backupName);
    const newSsId = newSs.getId();
    const newSsUrl = newSs.getUrl();
    const defaultSheet = newSs.getSheets()[0];

    // 3. คัดลอกทุกชีตในระบบ (ยกเว้นชีต Backups เอง)
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const s = sheets[i];
      if (s.getName() !== BACKUP_SHEET_NAME) {
        const copied = s.copyTo(newSs);
        copied.setName(s.getName());
      }
    }

    // ลบ Sheet1 เริ่มต้นของ Spreadsheet ใหม่
    if (newSs.getSheets().length > 1) {
      newSs.deleteSheet(defaultSheet);
    }

    // 4. ถ้าผู้ใช้เปิดสิทธิ์ DriveApp ไว้ ให้พยายามย้ายเข้าโฟลเดอร์ WNTK_Library_Backups (แบบ Graceful)
    try {
      if (typeof DriveApp !== 'undefined') {
        const folders = DriveApp.getFoldersByName("WNTK_Library_Backups");
        let folder;
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = DriveApp.createFolder("WNTK_Library_Backups");
        }
        const file = DriveApp.getFileById(newSsId);
        folder.addFile(file);
      }
    } catch (driveErr) {
      // ทำงานต่อได้ปกติ แม้ DriveApp ยังไม่ได้รับสิทธิ์
    }

    // 5. บันทึกประวัติไฟล์สำรองใหม่ลงชีต Backups (มีไฟล์ปัจจุบัน 1 ไฟล์เสมอ)
    logSheet.appendRow([
      thaiDateStr,
      backupName,
      newSsUrl,
      newSsId,
      triggerType || "กำหนดเอง (Manual)",
      "สำเร็จ"
    ]);

    return {
      status: "success",
      msg: `สำรองข้อมูลทับไฟล์เดิมสำเร็จเรียบร้อย (${thaiDateStr})`,
      fileName: backupName,
      fileUrl: newSsUrl,
      time: thaiDateStr
    };
  } catch (err) {
    return {
      status: "error",
      msg: "เกิดข้อผิดพลาดในการสำรองข้อมูล: " + err.toString()
    };
  }
}

// ฟังก์ชันลบไฟล์สำรองข้อมูลที่เลือก
function deleteLibraryBackup(fileId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = getOrCreateBackupLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    fileId = String(fileId || '').trim();
    if (!fileId) return { status: "error", msg: "ไม่พบรหัสไฟล์สำรอง" };
    
    let deletedName = "";
    if (lastRow > 1) {
      const data = logSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        const rowFileId = String(data[i][3] || '').trim();
        if (rowFileId === fileId) {
          deletedName = String(data[i][1] || '').trim();
          logSheet.deleteRow(i + 2);
        }
      }
    }
    
    try {
      if (typeof DriveApp !== 'undefined' && fileId) {
        const file = DriveApp.getFileById(fileId);
        if (file) file.setTrashed(true);
      }
    } catch (driveErr) {
      console.warn("DriveApp trash warning:", driveErr);
    }
    
    return { 
      status: "success", 
      msg: `ลบไฟล์สำรอง ${deletedName ? '"' + deletedName + '"' : ''} เรียบร้อยแล้ว` 
    };
  } catch (err) {
    return { 
      status: "error", 
      msg: "เกิดข้อผิดพลาดในการลบไฟล์สำรอง: " + err.toString() 
    };
  }
}

// ดึงข้อมูลประวัติการสำรองข้อมูลและสถานะระบบ
function getLibraryBackupInfo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = getOrCreateBackupLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    
    let backups = [];
    if (lastRow > 1) {
      const data = logSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        const dateStr = String(data[i][0] || '').trim();
        const name = String(data[i][1] || '').trim();
        const url = String(data[i][2] || '').trim();
        const fileId = String(data[i][3] || '').trim();
        const type = String(data[i][4] || 'อัตโนมัติ').trim();
        if (name && url) {
          backups.push({
            name: name,
            createdAt: dateStr,
            url: url,
            fileId: fileId,
            type: type,
            size: "Google Sheets"
          });
        }
        if (backups.length >= 15) break;
      }
    }

    // ตรวจสอบว่าวันนี้มีการสำรองข้อมูลแล้วหรือยัง (สำหรับ Auto Backup Daily Check)
    const today = new Date();
    const todayDay = String(today.getDate()).padStart(2, '0');
    const todayMonth = String(today.getMonth() + 1).padStart(2, '0');
    const todayYearBE = String(today.getFullYear() + 543);
    const todayPrefix = `${todayDay}/${todayMonth}/${todayYearBE}`;
    
    let backedUpToday = false;
    if (backups.length > 0 && backups[0].createdAt.startsWith(todayPrefix)) {
      backedUpToday = true;
    }

    return {
      status: "success",
      folderUrl: "https://drive.google.com/drive/my-drive",
      hasAutoTrigger: true,
      triggerSchedule: "ทำงานเมื่อล็อกอินเข้าใช้งานครั้งแรกของวัน",
      maxFiles: 1,
      totalBackups: Math.max(0, lastRow - 1),
      backedUpToday: backedUpToday,
      backups: backups
    };
  } catch (err) {
    return {
      status: "error",
      msg: "ไม่สามารถดึงข้อมูลสำรอง: " + err.toString(),
      folderUrl: "https://drive.google.com/drive/my-drive",
      hasAutoTrigger: true,
      backups: []
    };
  }
}

// ตรวจสอบและสำรองข้อมูลอัตโนมัติประจำวันเมื่อเจ้าหน้าที่ล็อกอิน
function checkAndAutoBackupOnLogin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = getOrCreateBackupLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    
    const today = new Date();
    const todayDay = String(today.getDate()).padStart(2, '0');
    const todayMonth = String(today.getMonth() + 1).padStart(2, '0');
    const todayYearBE = String(today.getFullYear() + 543);
    const todayPrefix = `${todayDay}/${todayMonth}/${todayYearBE}`;

    if (lastRow > 1) {
      const lastDate = String(logSheet.getRange(lastRow, 1).getValue() || '').trim();
      if (lastDate.startsWith(todayPrefix)) {
        return; // วันนี้สำรองข้อมูลไปแล้ว
      }
    }
    // วันนี้ยังไม่มีการสำรองข้อมูล ให้สร้างสแนปช็อตอัตโนมัติ
    createLibraryBackup("อัตโนมัติประจำวัน (Daily Auto)");
  } catch (e) {
    console.error("Auto backup on login check failed:", e);
  }
}

// ฟังก์ชันสำหรับเรียกใช้งานโดยอัตโนมัติจาก Time-driven Trigger หรือ Event
function autoLibraryBackupDailyTrigger() {
  createLibraryBackup("อัตโนมัติ (Daily Trigger)");
}

// ฟังก์ชัน API รับ request จากหน้าเว็บภายนอก
function doPost(e) {
  try {
    var contents = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var data = JSON.parse(contents);
    var action = data.action;
    var args = data.args || [];
    
    var fn = (typeof globalThis !== 'undefined' && typeof globalThis[action] === 'function') 
             ? globalThis[action] 
             : (typeof this[action] === 'function' ? this[action] : null);

    if (!fn) {
      return ContentService.createTextOutput(JSON.stringify({ 
        error: "ไม่พบฟังก์ชัน: " + action 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var result = fn.apply(null, args);
    return ContentService.createTextOutput(JSON.stringify(result !== undefined ? result : { status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      error: "Error: " + err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
