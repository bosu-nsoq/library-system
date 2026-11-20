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

// --- 1. Login System ---
function loginStaff(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Staff");
  if (!sheet) return { status: "error", msg: "ไม่พบ Sheet 'Staff'" };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(pin)) {
      return { 
        status: "success", 
        name: data[i][2],
        libraryData: getFullLibraryData() // โหลดข้อมูลทั้งหมดลงเครื่องทันทีที่ล็อกอิน
      };
    }
  }
  return { status: "fail" };
}

// --- 2. History & Print Data ---
function getRecentLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Logs");
  const lastRow = logSheet.getLastRow();
  const numRows = 10;
  if (lastRow < 2) return []; 
  const startRow = Math.max(2, lastRow - numRows + 1);
  const data = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();
  let logs = [];
  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let d = new Date(row[0]);
    let timeStr = d.getHours().toString().padStart(2, '0') + ":" + d.getMinutes().toString().padStart(2, '0');
    logs.push({ rowIndex: startRow + i, time: timeStr, action: row[1], studentName: row[3], itemName: row[4], note: row[6] });
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

  let results = [];
  let successCount = 0;
  for (let i = 0; i < bookIds.length; i++) {
    let bid = String(bookIds[i]).trim();
    if (!bid) continue;
    let res = processTransaction(studentId, bid, action, staffName);
    let cleanMsg = res.split(":")[1] || res;
    if (res.startsWith("success:")) successCount++;
    results.push(`[${bid}] ${cleanMsg}`);
  }
  let prefix = successCount > 0 ? "success:" : "warning:";
  return prefix + `ทำรายการสำเร็จ ${successCount}/${bookIds.length} เล่ม\n` + results.join("\n");
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
    if (available <= 0) return "warning:⚠️ หนังสือไม่ว่าง (ถูกยืมครบแล้ว)";
    if (countHolding(bookSheet, studentId, studentName) >= 3) return "warning:⚠️ ยืมเกินโควตา 3 เล่ม";
    
    // ✅ บันทึกเป็นชื่อผู้ยืมในชีต Books
    borrowers.push(studentName);
    const newAvailable = totalQty - borrowers.length;
    bookSheet.getRange(bRow, 4).setValue(newAvailable > 0 ? "ว่าง" : "ถูกยืม");
    bookSheet.getRange(bRow, 5).setValue(borrowers.join(", "));
    
    const dueDate = new Date(); dueDate.setDate(today.getDate() + 3); 
    logSheet.appendRow([today, "ยืม", studentId, studentName, bookId, dueDate, "", staffName]);
    
    return "success:✅ ยืมสำเร็จ!";
    
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
           let parts = range.value.split('-');
           return (rowDate.getDate() == parseInt(parts[2], 10) && rowDate.getMonth() == parseInt(parts[1], 10) - 1 && rowDate.getFullYear() == parseInt(parts[0], 10));
       } else if (range.type === 'custom_month') {
           let parts = range.value.split('-');
           return (rowDate.getMonth() == parseInt(parts[1], 10) - 1 && rowDate.getFullYear() == parseInt(parts[0], 10));
       } else if (range.type === 'custom_year') {
           return (rowDate.getFullYear() == parseInt(range.value, 10));
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
         result.push({
           date: Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy"),
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
            score: Number(sData[i][3]) || 0 
          };
          break;
        }
      }
    }
  }

  // 2. ดึงข้อมูลหนังสือ พร้อมคำนวณจำนวนวันการยืม (เช่น ด.ช.เอ ยืมไปแล้ว 7 วัน)
  if (bookId) { 
    const bookSheet = ss.getSheetByName("Books");
    if (bookSheet) {
      const bData = bookSheet.getDataRange().getValues();
      const targetBid = String(bookId).trim();
      for (let i = 1; i < bData.length; i++) {
        if (String(bData[i][0]).trim() === targetBid) {
          const bName = String(bData[i][1] || '').trim();
          const totalQty = Number(bData[i][5]) || 1;
          const borrowersStr = String(bData[i][4] || '').trim();
          const borrowers = borrowersStr ? borrowersStr.split(",").map(s => s.trim()).filter(s => s !== "") : [];
          const available = Math.max(0, totalQty - borrowers.length);

          let borrowDaysText = "";
          let borrowerName = "";
          let borrowDays = 0;

          // คำนวณจำนวนวันที่ยืมจาก Logs
          if (borrowers.length > 0) {
            borrowerName = borrowers[borrowers.length - 1];
            const logSheet = ss.getSheetByName("Logs");
            if (logSheet) {
              const logData = logSheet.getDataRange().getValues();
              for (let j = logData.length - 1; j >= 1; j--) {
                if (String(logData[j][4]).trim() === targetBid && String(logData[j][1]).trim() === "ยืม") {
                  const borrowDate = new Date(logData[j][0]);
                  if (!isNaN(borrowDate.getTime())) {
                    const today = new Date();
                    today.setHours(0,0,0,0);
                    const bDate = new Date(borrowDate);
                    bDate.setHours(0,0,0,0);
                    const diffDays = Math.round((today.getTime() - bDate.getTime()) / (1000 * 60 * 60 * 24));
                    borrowDays = Math.max(0, diffDays);
                    borrowDaysText = borrowDays === 0 ? "ยืมวันนี้" : `ยืมไปแล้ว ${borrowDays} วัน`;
                    if (logData[j][3]) {
                      borrowerName = String(logData[j][3]).trim();
                    }
                  }
                  break;
                }
              }
            }
          }

          let statusText = "";
          if (available > 0) {
            statusText = totalQty > 1 ? `ว่าง (${available}/${totalQty} เล่ม)` : "ว่าง";
          } else {
            const detailStr = [borrowerName, borrowDaysText].filter(Boolean).join(" ");
            statusText = totalQty > 1 
              ? (detailStr ? `ถูกยืมครบแล้ว (${detailStr})` : `ถูกยืมครบแล้ว (0/${totalQty} เล่ม)`)
              : (detailStr ? `ถูกยืม (${detailStr})` : "ถูกยืม");
          }

          bInfo = { 
            name: bName, 
            status: statusText, 
            available: available, 
            totalQty: totalQty,
            borrowerName: borrowerName,
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
        return { dueDate: data[i][5] }; 
      }
    }
  } 
  return null; 
}

function formatDate(date) { let d = new Date(date); return `${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()+543}`; }

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
          score: Number(sData[i][3]) || 0 
        };
      }
    }
  }

  // คำนวณข้อมูลการยืมล่าสุดจาก Logs
  let lastBorrowMap = {};
  if (lSheet) {
    const lData = lSheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0,0,0,0);
    for (let j = lData.length - 1; j >= 1; j--) {
      const bid = String(lData[j][4] || '').trim();
      const action = String(lData[j][1] || '').trim();
      if (bid && action === "ยืม" && !lastBorrowMap[bid]) {
        const bDate = new Date(lData[j][0]);
        if (!isNaN(bDate.getTime())) {
          bDate.setHours(0,0,0,0);
          const days = Math.max(0, Math.round((today.getTime() - bDate.getTime()) / (1000 * 60 * 60 * 24)));
          lastBorrowMap[bid] = {
            days: days,
            borrowDaysText: days === 0 ? "ยืมวันนี้" : `ยืมไปแล้ว ${days} วัน`,
            borrowerName: String(lData[j][3] || '').trim()
          };
        }
      }
    }
  }

  let books = {};
  if (bSheet) {
    const bData = bSheet.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      const bid = String(bData[i][0] || '').trim();
      if (bid) {
        const bName = String(bData[i][1] || '').trim();
        const totalQty = Number(bData[i][5]) || 1;
        const borrowersStr = String(bData[i][4] || '').trim();
        const borrowers = borrowersStr ? borrowersStr.split(",").map(s => s.trim()).filter(Boolean) : [];
        const available = Math.max(0, totalQty - borrowers.length);

        let borrowDaysText = "";
        let borrowerName = "";
        let borrowDays = 0;
        if (borrowers.length > 0 && lastBorrowMap[bid]) {
          borrowDaysText = lastBorrowMap[bid].borrowDaysText;
          borrowerName = lastBorrowMap[bid].borrowerName || borrowers[borrowers.length - 1];
          borrowDays = lastBorrowMap[bid].days;
        }

        let statusText = "";
        if (available > 0) {
          statusText = totalQty > 1 ? `ว่าง (${available}/${totalQty} เล่ม)` : "ว่าง";
        } else {
          const detailStr = [borrowerName, borrowDaysText].filter(Boolean).join(" ");
          statusText = totalQty > 1 
            ? (detailStr ? `ถูกยืมครบแล้ว (${detailStr})` : `ถูกยืมครบแล้ว (0/${totalQty} เล่ม)`)
            : (detailStr ? `ถูกยืม (${detailStr})` : "ถูกยืม");
        }

        books[bid] = {
          name: bName,
          status: statusText,
          available: available,
          totalQty: totalQty,
          borrowerName: borrowerName,
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
