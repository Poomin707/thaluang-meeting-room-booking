/**
 * ระบบเว็บแอปพลิเคชันจองห้องประชุมออนไลน์
 * วิทยาลัยเทคนิคท่าหลวงซิเมนต์ไทยอนุสรณ์
 * 
 * Google Apps Script Backend (Code.gs)
 * Spreadsheet ID: 1Ap5fdiPYEs8LhEEfAUmQRVaDK9k0D_m_ZyovxPCgk3k
 */

// รหัส Google Spreadsheet ของระบบ
const SPREADSHEET_ID = "1Ap5fdiPYEs8LhEEfAUmQRVaDK9k0D_m_ZyovxPCgk3k";
const SHEET_NAME_BOOKINGS = "Bookings";
const SHEET_NAME_ROOMS = "Rooms";
const SHEET_NAME_LOGS = "AuditLogs";

/**
 * ฟังก์ชันหลักสำหรับแสดงผล Web App (GET Request)
 */
function doGet(e) {
  // หากมีการร้องขอ API ผ่านพารามิเตอร์ action
  if (e && e.parameter && e.parameter.action) {
    return handleApiGet(e);
  }
  
  // ให้บริการหน้าเว็บแอปพลิเคชัน
  const htmlOutput = HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("ระบบจองห้องประชุมออนไลน์ - วท.ท่าหลวงซิเมนต์ไทยอนุสรณ์")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
  return htmlOutput;
}

/**
 * ฟังก์ชันรับข้อมูล POST Request (สำหรับ API ภายนอกหรือ AJAX)
 */
function doPost(e) {
  try {
    let requestData = {};
    if (e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else {
      requestData = e.parameter;
    }

    const action = requestData.action;
    let result = { success: false, message: "Unknown action" };

    switch (action) {
      case "createBooking":
        result = createBooking(requestData.data);
        break;
      case "updateApproval":
        result = updateApprovalStatus(
          requestData.bookingId,
          requestData.step,
          requestData.status,
          requestData.actorName,
          requestData.comment
        );
        break;
      case "cancelBooking":
        result = cancelBooking(requestData.bookingId, requestData.reason);
        break;
      case "checkRoomAvailability":
        result = checkRoomAvailability(
          requestData.roomName,
          requestData.startDate,
          requestData.startTime,
          requestData.endDate,
          requestData.endTime,
          requestData.excludeBookingId
        );
        break;
      default:
        result = { success: false, message: "Action not recognized" };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * จัดการ API GET Request
 */
function handleApiGet(e) {
  const action = e.parameter.action;
  let result = {};

  try {
    switch (action) {
      case "initSheets":
        result = initSheets();
        break;
      case "getBookings":
        result = { success: true, data: getAllBookings() };
        break;
      case "getDashboardStats":
        result = { success: true, data: getDashboardStats() };
        break;
      case "getRooms":
        result = { success: true, data: getRoomList() };
        break;
      default:
        result = { success: false, message: "Invalid action" };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * เชื่อมต่อ Spreadsheet
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * ฟังก์ชันสร้างและตั้งค่าชีตอัตโนมัติ (รันครั้งแรกเพียงคลิกเดียว)
 */
function initSheets() {
  const ss = getSpreadsheet();
  
  // 1. ตรวจสอบ/สร้างชีต Bookings
  let bookingSheet = ss.getSheetByName(SHEET_NAME_BOOKINGS);
  if (!bookingSheet) {
    bookingSheet = ss.insertSheet(SHEET_NAME_BOOKINGS);
  }
  
  const bookingHeaders = [
    "BookingID",               // รหัสการจอง เช่น TL-20261028-001
    "Timestamp",               // เวลาที่ทำรายการ
    "FullName",                // ชื่อ-นามสกุล ผู้ขอใช้
    "Position",                // ตำแหน่ง
    "Department",              // หน่วยงาน / สาขาวิชา
    "Phone",                   // หมายเลขโทรศัพท์
    "Objective",               // วัตถุประสงค์การใช้งาน
    "RoomName",                // ห้องประชุมที่ขอใช้
    "StartDate",               // วันที่เริ่มต้น (YYYY-MM-DD)
    "EndDate",                 // วันที่สิ้นสุด (YYYY-MM-DD)
    "StartTime",               // เวลาเริ่มต้น (HH:mm)
    "EndTime",                 // เวลาสิ้นสุด (HH:mm)
    "AttendeeCount",           // จำนวนผู้เข้าใช้โดยประมาณ (คน)
    "Equipments",              // ครุภัณฑ์ที่ขอใช้ (คั่นด้วยจุลภาค)
    "OtherEquipment",          // ครุภัณฑ์หรืออุปกรณ์อื่นๆ เพิ่มเติม
    "Remarks",                 // หมายเหตุเพิ่มเติม
    // สายการพิจารณาและตรวจสอบ (Workflow)
    "Step1_Certifier_Name",    // หัวหน้างาน/ผู้รับรอง (นายเอกราช เขียวทอง)
    "Step1_Status",            // สถานะ: รอรับรอง / รับรองแล้ว / ปฏิเสธ
    "Step1_Timestamp",         // วันเวลาที่รับรอง
    "Step2_Inspector_Name",    // จนท. ผู้ตรวจสอบห้อง (นายวัชรพงศ์ โสภา)
    "Step2_RoomStatus",        // สถานะ: พร้อมใช้งาน / ไม่พร้อมใช้งาน
    "Step2_Note",              // ความเห็นการตรวจสอบ
    "Step2_Timestamp",         // วันเวลาที่ตรวจ
    "Step3_Approver_Name",     // ผู้พิจารณาอนุมัติ (นายชาตรี เริงชัยภูมิ)
    "Step3_Status",            // สถานะ: รออนุมัติ / อนุมัติ / ไม่อนุมัติ
    "Step3_Note",              // ความเห็นผู้พิจารณาอนุมัติ
    "Step3_Timestamp",         // วันเวลาที่อนุมัติ
    "Overall_Status"           // สถานะรวม: รอตรวจสอบ / พร้อมใช้งาน-รออนุมัติ / อนุมัติเรียบร้อย / ไม่อนุมัติ / ยกเลิก
  ];
  
  bookingSheet.getRange(1, 1, 1, bookingHeaders.length).setValues([bookingHeaders]);
  bookingSheet.getRange(1, 1, 1, bookingHeaders.length)
    .setFontWeight("bold")
    .setBackground("#1E3A8A") // Dark Navy
    .setFontColor("#FFFFFF");
  bookingSheet.setFrozenRows(1);
  
  // 2. ตรวจสอบ/สร้างชีต Rooms
  let roomSheet = ss.getSheetByName(SHEET_NAME_ROOMS);
  if (!roomSheet) {
    roomSheet = ss.insertSheet(SHEET_NAME_ROOMS);
  }
  
  const roomHeaders = ["RoomID", "RoomName", "Capacity", "Location", "Facilities", "Status", "ImageURL"];
  roomSheet.getRange(1, 1, 1, roomHeaders.length).setValues([roomHeaders]);
  roomSheet.getRange(1, 1, 1, roomHeaders.length)
    .setFontWeight("bold")
    .setBackground("#0F766E")
    .setFontColor("#FFFFFF");
  roomSheet.setFrozenRows(1);
  
  // เติมข้อมูลห้องเริ่มต้นหากว่างอยู่
  if (roomSheet.getLastRow() <= 1) {
    const defaultRooms = [
      ["R01", "หอประชุมอดิเรกสาร", 500, "อาคารอำนวยการ ชั้น 3", "เครื่องฉายภาพ, ระบบเสียงเวที, แอร์, โต๊ะ-เก้าอี้", "พร้อมใช้งาน", ""],
      ["R02", "ห้องประชุมสักทอง", 80, "อาคารอำนวยการ ชั้น 2", "Smart TV/โปรเจกเตอร์, ระบบเสียงประชุม, แอร์, Wifi", "พร้อมใช้งาน", ""],
      ["R03", "อาคารโดมเอนกประสงค์", 1000, "ลานกิจกรรมกลาง", "ระบบเสียงสนาม, เวทียก, พัดลมขนาดใหญ่", "พร้อมใช้งาน", ""],
      ["R04", "ห้องประชุม (กาแฟ)", 25, "อาคารวิทยบริการ ชั้น 1", "จอแสดงผล, เครื่องปรับอากาศ, บาร์กาแฟ, Wifi", "พร้อมใช้งาน", ""],
      ["R05", "ห้องประชุมหน้าห้องผู้อำนวยการ", 20, "อาคารอำนวยการ ชั้น 2", "จอโปรเจกเตอร์, เครื่องปรับอากาศ, ชุดไมโครโฟนประชุม", "พร้อมใช้งาน", ""],
      ["R06", "ห้องคอมพิวเตอร์กลาง (ห้องโสต)", 50, "ศูนย์วิทยบริการและดิจิทัล", "คอมพิวเตอร์ 50 เครื่อง, เครื่องฉายภาพ, แอร์, อินเทอร์เน็ตความเร็วสูง", "พร้อมใช้งาน", ""],
      ["R07", "อื่นๆ", 50, "ตามระบุในคำขอ", "ตามที่ระบุในรายการครุภัณฑ์", "พร้อมใช้งาน", ""]
    ];
    roomSheet.getRange(2, 1, defaultRooms.length, roomHeaders.length).setValues(defaultRooms);
  }
  
  return { success: true, message: "เตรียมโครงสร้างชีตเรียบร้อยแล้ว!" };
}

/**
 * ดึงรายการห้องประชุมทั้งหมด
 */
function getRoomList() {
  const ss = getSpreadsheet();
  let roomSheet = ss.getSheetByName(SHEET_NAME_ROOMS);
  if (!roomSheet) {
    initSheets();
    roomSheet = ss.getSheetByName(SHEET_NAME_ROOMS);
  }
  
  const data = roomSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rooms = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const room = {};
    headers.forEach((header, index) => {
      room[header] = row[index];
    });
    rooms.push(room);
  }
  return rooms;
}

/**
 * ดึงข้อมูลการจองทั้งหมด
 */
function getAllBookings() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_BOOKINGS);
  if (!sheet) {
    initSheets();
    sheet = ss.getSheetByName(SHEET_NAME_BOOKINGS);
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const bookings = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const booking = {};
    headers.forEach((header, index) => {
      let val = row[index];
      // Format วันที่ให้เป็น ISO string สม่ำเสมอ
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
      }
      booking[header] = val;
    });
    bookings.push(booking);
  }
  
  // เรียงลำดับตามวันที่จองล่าสุดไปเก่า
  bookings.sort((a, b) => new Date(b.StartDate + "T" + (b.StartTime || "00:00")) - new Date(a.StartDate + "T" + (a.StartTime || "00:00")));
  return bookings;
}

/**
 * ตรวจสอบห้องว่างและเช็คความทับซ้อน (Conflict Checking)
 */
function checkRoomAvailability(roomName, startDate, startTime, endDate, endTime, excludeBookingId) {
  if (roomName === "อื่นๆ") {
    return { available: true, conflicts: [] };
  }
  
  const allBookings = getAllBookings();
  const conflicts = [];
  
  const reqStart = new Date(startDate + "T" + startTime);
  const reqEnd = new Date(endDate + "T" + endTime);
  
  if (isNaN(reqStart.getTime()) || isNaN(reqEnd.getTime())) {
    return { available: false, error: "รูปแบบวันเวลาไม่ถูกต้อง" };
  }
  
  if (reqStart >= reqEnd) {
    return { available: false, error: "เวลาเริ่มต้นต้องมาก่อนเวลาสิ้นสุด" };
  }
  
  for (let b of allBookings) {
    // ข้ามรายการที่ถูกยกเลิก หรือไม่อนุมัติ
    if (b.Overall_Status === "ไม่อนุมัติ" || b.Overall_Status === "ยกเลิก") {
      continue;
    }
    
    // ข้าม id ตัวเองหากเป็นการแก้ไข
    if (excludeBookingId && b.BookingID === excludeBookingId) {
      continue;
    }
    
    if (b.RoomName === roomName) {
      const bStart = new Date(b.StartDate + "T" + b.StartTime);
      const bEnd = new Date(b.EndDate + "T" + b.EndTime);
      
      // ช่วงเวลาเหลื่อมกันถ้า (reqStart < bEnd) และ (reqEnd > bStart)
      if (reqStart < bEnd && reqEnd > bStart) {
        conflicts.push({
          bookingId: b.BookingID,
          objective: b.Objective,
          fullName: b.FullName,
          department: b.Department,
          startDate: b.StartDate,
          startTime: b.StartTime,
          endDate: b.EndDate,
          endTime: b.EndTime,
          status: b.Overall_Status
        });
      }
    }
  }
  
  return {
    available: conflicts.length === 0,
    conflicts: conflicts
  };
}

/**
 * บันทึกการจองใหม่ (Create Booking)
 */
function createBooking(data) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME_BOOKINGS);
    if (!sheet) {
      initSheets();
      sheet = ss.getSheetByName(SHEET_NAME_BOOKINGS);
    }
    
    // 1. ตรวจสอบ Conflict ซ้ำก่อนบันทึก
    const availability = checkRoomAvailability(
      data.roomName,
      data.startDate,
      data.startTime,
      data.endDate,
      data.endTime,
      null
    );
    
    if (!availability.available) {
      return {
        success: false,
        message: "ไม่สามารถจองได้ เนื่องจากช่วงเวลานี้มีรายการจองอื่นซ้อนทับอยู่แล้ว",
        conflicts: availability.conflicts
      };
    }
    
    // 2. สร้าง Booking ID ไม่ซ้ำ
    const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd");
    const count = sheet.getLastRow(); // แถวปัจจุบัน
    const randomSuffix = Math.floor(100 + Math.random() * 900);
    const bookingId = `TL-${todayStr}-${count.toString().padStart(3, '0')}-${randomSuffix}`;
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    // จัดรูปแบบ Equipment
    const equipmentsStr = Array.isArray(data.equipments) ? data.equipments.join(", ") : (data.equipments || "");
    
    // กำหนดค่าเริ่มต้นตามเอกสารราชการ
    const defaultCertifier = "นายเอกราช เขียวทอง (หัวหน้างานศูนย์ดิจิทัลและสื่อสารองค์กร)";
    const defaultInspector = "นายวัชรพงศ์ โสภา (เจ้าหน้าที่งานศูนย์ดิจิทัลและสื่อสารองค์กร)";
    const defaultApprover = "นายชาตรี เริงชัยภูมิ (รองผู้อำนวยการฝ่ายพัฒนายุทธศาสตร์และแผนงาน)";
    
    const newRow = [
      bookingId,
      timestamp,
      data.fullName || "",
      data.position || "",
      data.department || "",
      data.phone || "",
      data.objective || "",
      data.roomName || "",
      data.startDate || "",
      data.endDate || "",
      data.startTime || "",
      data.endTime || "",
      Number(data.attendeeCount) || 0,
      equipmentsStr,
      data.otherEquipment || "",
      data.remarks || "",
      // Step 1: หัวหน้างาน/ผู้รับรอง
      defaultCertifier,
      "รอการรับรอง",
      "",
      // Step 2: เจ้าหน้าที่ตรวจสอบห้อง
      defaultInspector,
      "รอการตรวจสอบ",
      "",
      "",
      // Step 3: ผู้พิจารณาอนุมัติ
      defaultApprover,
      "รอการอนุมัติ",
      "",
      "",
      // สถานะรวม
      "รอตรวจสอบ"
    ];
    
    sheet.appendRow(newRow);
    
    // ส่งอีเมลแจ้งเตือน (ถ้ากำหนด email ไว้)
    notifyBookingCreated(bookingId, data);
    
    return {
      success: true,
      message: "บันทึกการจองสำเร็จ รหัสการจอง: " + bookingId,
      bookingId: bookingId
    };
    
  } catch (err) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาดในการบันทึก: " + err.toString()
    };
  }
}

/**
 * อัปเดตสถานะ Workflow การพิจารณาและตรวจสอบ (3 ขั้นตอน)
 * step: 1 (รับรอง), 2 (ตรวจสอบห้อง), 3 (อนุมัติ)
 */
function updateApprovalStatus(bookingId, step, status, actorName, note) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_BOOKINGS);
    if (!sheet) return { success: false, message: "ไม่พบชีตข้อมูล" };
    
    const data = sheet.getDataRange().getValues();
    let targetRowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === bookingId) {
        targetRowIndex = i + 1; // 1-based index ใน Apps Script
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      return { success: false, message: "ไม่พบรหัสการจอง: " + bookingId };
    }
    
    const nowTime = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    // คอลัมน์อิงตาม Header ที่กำหนดไว้ใน initSheets:
    // 17: Step1_Certifier_Name, 18: Step1_Status, 19: Step1_Timestamp
    // 20: Step2_Inspector_Name, 21: Step2_RoomStatus, 22: Step2_Note, 23: Step2_Timestamp
    // 24: Step3_Approver_Name, 25: Step3_Status, 26: Step3_Note, 27: Step3_Timestamp
    // 28: Overall_Status
    
    if (step == 1) {
      // ขั้นตอนที่ 1: หัวหน้างาน/ผู้รับรอง
      sheet.getRange(targetRowIndex, 18).setValue(status); // รับรองแล้ว / ไม่รับรอง
      sheet.getRange(targetRowIndex, 19).setValue(nowTime);
      if (actorName) sheet.getRange(targetRowIndex, 17).setValue(actorName);
      
      if (status === "ไม่รับรอง") {
        sheet.getRange(targetRowIndex, 28).setValue("ไม่ผ่านการรับรอง");
      } else {
        sheet.getRange(targetRowIndex, 28).setValue("รับรองแล้ว-รอตรวจห้อง");
      }
    } 
    else if (step == 2) {
      // ขั้นตอนที่ 2: เจ้าหน้าที่ตรวจสอบห้อง (พร้อมใช้งาน / ไม่พร้อมใช้งาน)
      sheet.getRange(targetRowIndex, 21).setValue(status);
      sheet.getRange(targetRowIndex, 22).setValue(note || "");
      sheet.getRange(targetRowIndex, 23).setValue(nowTime);
      if (actorName) sheet.getRange(targetRowIndex, 20).setValue(actorName);
      
      if (status === "พร้อมใช้งาน") {
        sheet.getRange(targetRowIndex, 28).setValue("ห้องพร้อม-รออนุมัติ");
      } else {
        sheet.getRange(targetRowIndex, 28).setValue("ห้องไม่พร้อมใช้งาน");
      }
    } 
    else if (step == 3) {
      // ขั้นตอนที่ 3: ผู้พิจารณาอนุมัติ (อนุมัติ / ไม่อนุมัติ)
      sheet.getRange(targetRowIndex, 25).setValue(status);
      sheet.getRange(targetRowIndex, 26).setValue(note || "");
      sheet.getRange(targetRowIndex, 27).setValue(nowTime);
      if (actorName) sheet.getRange(targetRowIndex, 24).setValue(actorName);
      
      if (status === "อนุมัติ") {
        sheet.getRange(targetRowIndex, 28).setValue("อนุมัติเรียบร้อย");
      } else {
        sheet.getRange(targetRowIndex, 28).setValue("ไม่อนุมัติ");
      }
    }
    
    return {
      success: true,
      message: `อัปเดตขั้นตอนที่ ${step} สำเร็จ สถานะ: ${status}`
    };
    
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.toString() };
  }
}

/**
 * ยกเลิกการจอง
 */
function cancelBooking(bookingId, reason) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_BOOKINGS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === bookingId) {
        sheet.getRange(i + 1, 28).setValue("ยกเลิก");
        sheet.getRange(i + 1, 16).setValue("ขอยกเลิก: " + (reason || "ไม่ระบุ"));
        return { success: true, message: "ยกเลิกการจองสำเร็จ" };
      }
    }
    return { success: false, message: "ไม่พบรหัสการจอง" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * สรุปสถิติสำหรับ Dashboard
 */
function getDashboardStats() {
  const bookings = getAllBookings();
  const rooms = getRoomList();
  const today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  
  let todayCount = 0;
  let approvedCount = 0;
  let pendingCount = 0;
  let rejectedCount = 0;
  
  const roomUsageCount = {};
  const departmentCount = {};
  
  bookings.forEach(b => {
    // เช็คการจองวันนี้
    if (b.StartDate <= today && b.EndDate >= today && b.Overall_Status !== "ยกเลิก" && b.Overall_Status !== "ไม่อนุมัติ") {
      todayCount++;
    }
    
    if (b.Overall_Status === "อนุมัติเรียบร้อย") {
      approvedCount++;
    } else if (b.Overall_Status === "ไม่อนุมัติ" || b.Overall_Status === "ห้องไม่พร้อมใช้งาน") {
      rejectedCount++;
    } else if (b.Overall_Status !== "ยกเลิก") {
      pendingCount++;
    }
    
    // นับสถิติห้อง
    if (b.RoomName) {
      roomUsageCount[b.RoomName] = (roomUsageCount[b.RoomName] || 0) + 1;
    }
    // นับสถิติแผนก
    if (b.Department) {
      departmentCount[b.Department] = (departmentCount[b.Department] || 0) + 1;
    }
  });
  
  return {
    totalBookings: bookings.length,
    todayBookings: todayCount,
    approvedBookings: approvedCount,
    pendingBookings: pendingCount,
    rejectedBookings: rejectedCount,
    totalRooms: rooms.length,
    roomUsageCount: roomUsageCount,
    departmentCount: departmentCount
  };
}

/**
 * ส่งอีเมลแจ้งเตือนการจอง (ถ้าต้องการเปิดใช้งาน)
 */
function notifyBookingCreated(bookingId, data) {
  try {
    const notifyEmail = Session.getActiveUser().getEmail();
    if (!notifyEmail) return;
    
    const subject = `[ระบบจองห้องประชุม] คำขอจองใหม่: ${data.roomName} (${bookingId})`;
    const body = `
      เรียน ผู้เกี่ยวข้อง
      
      มีการส่งคำขอใช้ห้องประชุมใหม่เข้ามาในระบบ ดังนี้:
      - รหัสการจอง: ${bookingId}
      - ผู้ขอใช้: ${data.fullName} (${data.position})
      - แผนก/สาขาวิชา: ${data.department} โทร. ${data.phone}
      - วัตถุประสงค์: ${data.objective}
      - ห้องประชุม: ${data.roomName}
      - วันที่: ${data.startDate} ถึง ${data.endDate} เวลา ${data.startTime} - ${data.endTime}
      - จำนวนผู้เข้าร่วม: ${data.attendeeCount} คน
      
      กรุณาเข้าสู่ระบบเพื่อดำเนินการตรวจสอบและอนุมัติตามขั้นตอน
      
      วิทยาลัยเทคนิคท่าหลวงซิเมนต์ไทยอนุสรณ์
    `;
    
    MailApp.sendEmail(notifyEmail, subject, body);
  } catch (e) {
    Logger.log("Mail notification skipped: " + e.toString());
  }
}
