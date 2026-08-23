/**
 * ระบบเช็คชื่อนักเรียนออนไลน์ (Student Attendance System)
 * Google Apps Script Backend (Code.gs)
 * Features: Anti-Cheat System (TOTP Dynamic OTP, GPS Geofencing, Device Binding, Live Selfie),
 * Admin Auth (admin888), Class Management, and Excel Import.
 * Pure Clean Database (No Mock / No Demo Data)
 */

// Global Sheet Configuration
var SHEET_STUDENTS = 'Students';
var SHEET_ATTENDANCE = 'Attendance';
var SHEET_SESSIONS = 'Sessions';
var SHEET_CLASSES = 'Classes';

var ADMIN_PASSWORD = 'admin888';

/**
 * Serves the HTML Web Application
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('ระบบเช็คชื่อนักเรียนออนไลน์ - Anti-Cheat Edition')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Gets or creates the active spreadsheet
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Initializes Database with default sheets (Completely empty clean database)
 */
function initDatabase() {
  var ss = getSpreadsheet();
  
  // 1. Students Sheet
  var studentSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!studentSheet) {
    studentSheet = ss.insertSheet(SHEET_STUDENTS);
    studentSheet.appendRow(['StudentID', 'เลขที่', 'ชื่อ', 'ห้อง', 'DeviceID']);
  } else {
    studentSheet.clear();
    studentSheet.appendRow(['StudentID', 'เลขที่', 'ชื่อ', 'ห้อง', 'DeviceID']);
  }
  
  // 2. Attendance Sheet
  var attendanceSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!attendanceSheet) {
    attendanceSheet = ss.insertSheet(SHEET_ATTENDANCE);
    attendanceSheet.appendRow([
      'Date', 'StudentID', 'Status', 'CheckinTime', 'GPSLocation', 'DistanceMeters', 'DeviceID', 'SelfiePhoto', 'CheckinType'
    ]);
  } else {
    attendanceSheet.clear();
    attendanceSheet.appendRow([
      'Date', 'StudentID', 'Status', 'CheckinTime', 'GPSLocation', 'DistanceMeters', 'DeviceID', 'SelfiePhoto', 'CheckinType'
    ]);
  }
  
  // 3. Sessions Sheet (For Projector and Live Check-in)
  var sessionSheet = ss.getSheetByName(SHEET_SESSIONS);
  if (!sessionSheet) {
    sessionSheet = ss.insertSheet(SHEET_SESSIONS);
    sessionSheet.appendRow([
      'SessionID', 'ClassName', 'Date', 'CreatedAt', 'ExpiresAt', 'Lat', 'Lng', 'RadiusMeters', 'SecretSeed', 'RequireSelfie', 'RequireGPS', 'RequireDeviceBinding', 'Status'
    ]);
  } else {
    sessionSheet.clear();
    sessionSheet.appendRow([
      'SessionID', 'ClassName', 'Date', 'CreatedAt', 'ExpiresAt', 'Lat', 'Lng', 'RadiusMeters', 'SecretSeed', 'RequireSelfie', 'RequireGPS', 'RequireDeviceBinding', 'Status'
    ]);
  }

  // 4. Classes Sheet
  var classSheet = ss.getSheetByName(SHEET_CLASSES);
  if (!classSheet) {
    classSheet = ss.insertSheet(SHEET_CLASSES);
    classSheet.appendRow(['ClassName', 'CreatedAt']);
  } else {
    classSheet.clear();
    classSheet.appendRow(['ClassName', 'CreatedAt']);
  }
  
  return { status: 'success', message: 'ล้างข้อมูลและเริ่มต้นฐานข้อมูลใหม่เรียบร้อยแล้ว' };
}

/**
 * Verifies Admin Password
 */
function verifyAdminPassword(password) {
  if (password === ADMIN_PASSWORD) {
    return { status: 'success', message: 'เข้าสู่ระบบ Admin สำเร็จ' };
  }
  return { status: 'error', message: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
}

/**
 * Gets list of all classes (Empty array if no classes added yet)
 */
function getClassList() {
  var ss = getSpreadsheet();
  var classSheet = ss.getSheetByName(SHEET_CLASSES);
  var classMap = {};

  if (classSheet && classSheet.getLastRow() > 1) {
    var data = classSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var c = String(data[i][0]).trim();
      if (c) classMap[c] = true;
    }
  }

  // Also check distinct classes from Students sheet
  var studentSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (studentSheet && studentSheet.getLastRow() > 1) {
    var sData = studentSheet.getDataRange().getValues();
    for (var j = 1; j < sData.length; j++) {
      var sc = String(sData[j][3]).trim();
      if (sc) classMap[sc] = true;
    }
  }

  var classes = Object.keys(classMap);
  classes.sort(function(a, b) { return a.localeCompare(b, 'th'); });
  return classes;
}

/**
 * Adds a new classroom
 */
function addClass(className) {
  if (!className || !className.trim()) {
    return { status: 'error', message: 'กรุณาระบุชื่อห้องเรียน' };
  }
  className = className.trim();

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLASSES);
  if (!sheet) {
    initDatabase();
    sheet = ss.getSheetByName(SHEET_CLASSES);
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === className.toLowerCase()) {
      return { status: 'error', message: 'มีห้องเรียนชื่อนี้อยู่แล้วในระบบ' };
    }
  }

  sheet.appendRow([className, new Date().toISOString()]);
  return { status: 'success', message: 'เพิ่มห้องเรียน "' + className + '" เรียบร้อยแล้ว', classList: getClassList() };
}

/**
 * Deletes a classroom and removes its student associations
 */
function deleteClass(className) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLASSES);
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางห้องเรียน' };

  className = String(className).trim();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === className.toLowerCase()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  // Also remove students belonging to this class in Students sheet
  var studentSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (studentSheet && studentSheet.getLastRow() > 1) {
    var sData = studentSheet.getDataRange().getValues();
    for (var j = sData.length - 1; j >= 1; j--) {
      if (String(sData[j][3]).trim().toLowerCase() === className.toLowerCase()) {
        studentSheet.deleteRow(j + 1);
      }
    }
  }

  return { status: 'success', message: 'ลบห้องเรียน "' + className + '" เรียบร้อยแล้ว', classList: getClassList() };
}

/**
 * Gets list of all students (optionally filtered by class)
 */
function getStudents(classFilter) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) {
    initDatabase();
    sheet = ss.getSheetByName(SHEET_STUDENTS);
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var students = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var studentId = String(row[0]).trim();
    var studentNo = String(row[1]).trim();
    var name = String(row[2]).trim();
    var className = String(row[3]).trim();
    var deviceId = row[4] ? String(row[4]).trim() : '';
    
    if (!name) continue;

    if (!classFilter || classFilter === 'all' || className.toLowerCase() === classFilter.toLowerCase()) {
      students.push({
        studentId: studentId,
        studentNo: studentNo,
        name: name,
        className: className,
        deviceId: deviceId
      });
    }
  }
  
  students.sort(function(a, b) {
    if (a.className !== b.className) {
      return a.className.localeCompare(b.className, 'th');
    }
    return (parseInt(a.studentNo, 10) || 0) - (parseInt(b.studentNo, 10) || 0);
  });
  
  return students;
}

/**
 * Batch import students (from Excel or bulk form)
 */
function importStudentsBatch(studentsList) {
  if (!studentsList || studentsList.length === 0) {
    return { status: 'error', message: 'ไม่มีข้อมูลนักเรียนสำหรับนำเข้า' };
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  var classSheet = ss.getSheetByName(SHEET_CLASSES);
  if (!sheet || !classSheet) {
    initDatabase();
    sheet = ss.getSheetByName(SHEET_STUDENTS);
    classSheet = ss.getSheetByName(SHEET_CLASSES);
  }

  var classesMap = {};
  var cData = classSheet.getDataRange().getValues();
  for (var c = 1; c < cData.length; c++) {
    classesMap[String(cData[c][0]).trim()] = true;
  }

  var rowsToAdd = [];
  var newClasses = {};

  for (var k = 0; k < studentsList.length; k++) {
    var item = studentsList[k];
    var sNo = String(item.studentNo || (k + 1)).trim();
    var sName = String(item.name || '').trim();
    var sClass = String(item.className || '').trim();
    var sId = String(item.studentId || '').trim();

    if (!sName) continue;

    if (!sId) {
      sId = 'STD' + String(new Date().getTime() + k).slice(-6);
    }

    // Auto add new class to Classes sheet if not existing
    if (sClass && !classesMap[sClass] && !newClasses[sClass]) {
      newClasses[sClass] = true;
      classSheet.appendRow([sClass, new Date().toISOString()]);
      classesMap[sClass] = true;
    }

    rowsToAdd.push([sId, sNo, sName, sClass, '']);
  }

  if (rowsToAdd.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
  }

  return {
    status: 'success',
    message: 'นำเข้ารายชื่อนักเรียนสำเร็จ ' + rowsToAdd.length + ' คน',
    count: rowsToAdd.length,
    classList: getClassList()
  };
}

/**
 * =========================================================================
 * ANTI-CHEAT SESSION MANAGEMENT & LIVE PROJECTOR BACKEND
 * =========================================================================
 */
function createCheckinSession(config) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SESSIONS);
  if (!sheet) {
    initDatabase();
    sheet = ss.getSheetByName(SHEET_SESSIONS);
  }
  
  var now = new Date();
  var nowMs = now.getTime();
  var durationMinutes = config.durationMinutes || 15;
  var expiresAtMs = nowMs + (durationMinutes * 60 * 1000);
  var sessionId = 'SES' + String(nowMs).slice(-6);
  var secretSeed = 'SEC_' + Math.random().toString(36).substring(2, 10).toUpperCase();
  var dateStr = formatDateString(now);
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(config.className) && String(data[i][12]) === 'Active') {
      sheet.getRange(i + 1, 13).setValue('Closed');
    }
  }
  
  var newRow = [
    sessionId,
    config.className || 'all',
    dateStr,
    now.toISOString(),
    new Date(expiresAtMs).toISOString(),
    config.lat || 0,
    config.lng || 0,
    config.radiusMeters || 50,
    secretSeed,
    config.requireSelfie !== false ? 'YES' : 'NO',
    config.requireGPS !== false ? 'YES' : 'NO',
    config.requireDeviceBinding !== false ? 'YES' : 'NO',
    'Active'
  ];
  
  sheet.appendRow(newRow);
  
  return {
    status: 'success',
    message: 'เปิดคาบเช็คชื่อเรียบร้อยแล้ว',
    session: {
      sessionId: sessionId,
      className: config.className || 'all',
      date: dateStr,
      createdAt: now.toISOString(),
      expiresAt: new Date(expiresAtMs).toISOString(),
      lat: config.lat || 0,
      lng: config.lng || 0,
      radiusMeters: config.radiusMeters || 50,
      secretSeed: secretSeed,
      requireSelfie: config.requireSelfie !== false,
      requireGPS: config.requireGPS !== false,
      requireDeviceBinding: config.requireDeviceBinding !== false,
      status: 'Active'
    }
  };
}

function getActiveSession(className) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SESSIONS);
  if (!sheet) return null;
  
  var data = sheet.getDataRange().getValues();
  var nowMs = new Date().getTime();
  
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var sId = String(row[0]);
    var cName = String(row[1]);
    var expStr = String(row[4]);
    var expMs = new Date(expStr).getTime();
    var status = String(row[12]);
    
    if (status === 'Active') {
      if (nowMs > expMs) {
        sheet.getRange(i + 1, 13).setValue('Expired');
        continue;
      }
      
      if (!className || className === 'all' || cName === 'all' || cName === className) {
        return {
          sessionId: sId,
          className: cName,
          date: formatDateString(row[2]),
          createdAt: String(row[3]),
          expiresAt: expStr,
          lat: parseFloat(row[5]) || 0,
          lng: parseFloat(row[6]) || 0,
          radiusMeters: parseInt(row[7], 10) || 50,
          secretSeed: String(row[8]),
          requireSelfie: String(row[9]) === 'YES',
          requireGPS: String(row[10]) === 'YES',
          requireDeviceBinding: String(row[11]) === 'YES',
          status: 'Active',
          serverTimeMs: nowMs
        };
      }
    }
  }
  
  return null;
}

function closeCheckinSession(sessionId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SESSIONS);
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางคาบเรียน' };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(sessionId)) {
      sheet.getRange(i + 1, 13).setValue('Closed');
      return { status: 'success', message: 'ปิดคาบเช็คชื่อเรียบร้อยแล้ว' };
    }
  }
  return { status: 'error', message: 'ไม่พบคราบเรียนนี้' };
}

function generateTOTP(secretSeed, intervalSeconds, timestampMs) {
  intervalSeconds = intervalSeconds || 15;
  var counter = Math.floor((timestampMs || Date.now()) / (intervalSeconds * 1000));
  var str = secretSeed + '_' + counter;
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    var char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  var code = Math.abs(hash) % 1000000;
  var codeStr = String(code);
  while (codeStr.length < 6) {
    codeStr = '0' + codeStr;
  }
  return codeStr;
}

function calculateHaversineDistance(lat1, lon1, lat2, lon2) {
  var R = 6371e3;
  var phi1 = lat1 * Math.PI / 180;
  var phi2 = lat2 * Math.PI / 180;
  var deltaPhi = (lat2 - lat1) * Math.PI / 180;
  var deltaLambda = (lon2 - lon1) * Math.PI / 180;

  var a = Math.sin(deltaPhi / 2) * Math.sin(deltaPhi / 2) +
          Math.cos(phi1) * Math.cos(phi2) *
          Math.sin(deltaLambda / 2) * Math.sin(deltaLambda / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return Math.round(R * c);
}

/**
 * =========================================================================
 * STUDENT SELF CHECK-IN & ANTI-CHEAT VALIDATION ENGINE
 * =========================================================================
 */
function validateAndCheckinStudent(payload) {
  var ss = getSpreadsheet();
  var sessionSheet = ss.getSheetByName(SHEET_SESSIONS);
  var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  
  if (!sessionSheet || !attSheet) {
    initDatabase();
    sessionSheet = ss.getSheetByName(SHEET_SESSIONS);
    attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  }
  
  var sessionId = payload.sessionId;
  var studentId = String(payload.studentId || '').trim();
  var enteredOtp = String(payload.otp || '').trim();
  var studentLat = parseFloat(payload.lat);
  var studentLng = parseFloat(payload.lng);
  var deviceId = String(payload.deviceId || '').trim();
  var selfieBase64 = payload.selfieBase64 || '';
  
  if (!studentId) {
    return { status: 'error', message: 'กรุณาระบุรหัสนักศึกษา' };
  }
  
  var students = getStudents();
  var studentObj = null;
  for (var s = 0; s < students.length; s++) {
    if (students[s].studentId === studentId) {
      studentObj = students[s];
      break;
    }
  }
  if (!studentObj) {
    return { status: 'error', message: 'ไม่พบรหัสนักศึกษานี้ในระบบ' };
  }
  
  var activeSession = getActiveSession(studentObj.className);
  if (!activeSession) {
    return { status: 'error', message: 'ไม่มีคาบเรียนที่เปิดรับเช็คชื่อสำหรับห้องนี้ในขณะนี้' };
  }
  if (sessionId && activeSession.sessionId !== sessionId) {
    return { status: 'error', message: 'รหัสคาบเรียนไม่ตรงกับคาบที่เปิดอยู่' };
  }
  
  var now = new Date();
  var nowMs = now.getTime();
  
  var validOtpCurrent = generateTOTP(activeSession.secretSeed, 15, nowMs);
  var validOtpPrev = generateTOTP(activeSession.secretSeed, 15, nowMs - 15000);
  
  if (enteredOtp !== validOtpCurrent && enteredOtp !== validOtpPrev) {
    return {
      status: 'error',
      message: 'รหัส OTP 6 หลักไม่ถูกต้องหรือหมดอายุแล้ว กรุณาดูรหัสล่าสุดบนหน้าจอโปรเจกเตอร์'
    };
  }
  
  var distanceMeters = 0;
  if (activeSession.requireGPS && activeSession.lat && activeSession.lng) {
    if (isNaN(studentLat) || isNaN(studentLng) || (studentLat === 0 && studentLng === 0)) {
      return {
        status: 'error',
        message: 'กรุณาเปิด GPS และอนุญาตการเข้าถึงตำแหน่งพิกัดเพื่อเช็คชื่อ'
      };
    }
    
    distanceMeters = calculateHaversineDistance(
      activeSession.lat,
      activeSession.lng,
      studentLat,
      studentLng
    );
    
    if (distanceMeters > activeSession.radiusMeters) {
      return {
        status: 'error',
        message: 'คุณอยู่นอกพื้นที่ห้องเรียน (ระยะห่าง ' + distanceMeters + ' เมตร เกินกว่าที่กำหนด ' + activeSession.radiusMeters + ' เมตร)'
      };
    }
  }
  
  var attData = attSheet.getDataRange().getValues();
  var dateStr = activeSession.date;
  
  if (activeSession.requireDeviceBinding && deviceId) {
    for (var i = 1; i < attData.length; i++) {
      var rowDate = formatDateString(attData[i][0]);
      var rowStd = String(attData[i][1]);
      var rowDev = String(attData[i][6] || '');
      
      if (rowDate === dateStr) {
        if (rowDev === deviceId && rowStd !== studentId) {
          return {
            status: 'error',
            message: 'อุปกรณ์นี้ (Device ID) ได้ใช้เช็คชื่อให้นักศึกษาคนอื่นไปแล้วในคาบนี้ ห้ามเช็คชื่อแทนกัน'
          };
        }
        if (rowStd === studentId && (attData[i][2] === 'Present' || attData[i][2] === 'Late')) {
          return {
            status: 'error',
            message: 'นักศึกษาท่านนี้ได้ทำการเช็คชื่อในคาบนี้เรียบร้อยแล้ว'
          };
        }
      }
    }
  }
  
  if (activeSession.requireSelfie && (!selfieBase64 || selfieBase64.length < 50)) {
    return {
      status: 'error',
      message: 'กรุณาถ่ายรูปเซลฟี่ยืนยันตัวตนสดผ่านกล้องหน้า'
    };
  }
  
  var checkinTimeStr = now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  var gpsLocStr = (studentLat && studentLng) ? (studentLat.toFixed(6) + ',' + studentLng.toFixed(6)) : '-';
  
  var existingRowIndex = -1;
  for (var r = 1; r < attData.length; r++) {
    if (formatDateString(attData[r][0]) === dateStr && String(attData[r][1]) === studentId) {
      existingRowIndex = r + 1;
      break;
    }
  }
  
  if (existingRowIndex > 0) {
    attSheet.getRange(existingRowIndex, 3).setValue('Present');
    attSheet.getRange(existingRowIndex, 4).setValue(checkinTimeStr);
    attSheet.getRange(existingRowIndex, 5).setValue(gpsLocStr);
    attSheet.getRange(existingRowIndex, 6).setValue(distanceMeters);
    attSheet.getRange(existingRowIndex, 7).setValue(deviceId);
    attSheet.getRange(existingRowIndex, 8).setValue(selfieBase64);
    attSheet.getRange(existingRowIndex, 9).setValue('Self Check-in (Verified)');
  } else {
    attSheet.appendRow([
      dateStr,
      studentId,
      'Present',
      checkinTimeStr,
      gpsLocStr,
      distanceMeters,
      deviceId,
      selfieBase64,
      'Self Check-in (Verified)'
    ]);
  }
  
  return {
    status: 'success',
    message: 'เช็คชื่อสำเร็จเรียบร้อยแล้ว!',
    student: {
      studentId: studentObj.studentId,
      studentNo: studentObj.studentNo,
      name: studentObj.name,
      className: studentObj.className,
      checkinTime: checkinTimeStr,
      distanceMeters: distanceMeters,
      deviceId: deviceId.substring(0, 8) + '...',
      selfiePhoto: selfieBase64
    }
  };
}

function getSessionLiveAttendance(sessionId) {
  var ss = getSpreadsheet();
  var sessionSheet = ss.getSheetByName(SHEET_SESSIONS);
  var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  
  if (!sessionSheet || !attSheet) return { list: [], count: 0, total: 0 };
  
  var session = null;
  var sData = sessionSheet.getDataRange().getValues();
  for (var i = 1; i < sData.length; i++) {
    if (String(sData[i][0]) === String(sessionId)) {
      session = {
        sessionId: sData[i][0],
        className: sData[i][1],
        date: formatDateString(sData[i][2]),
        status: sData[i][12]
      };
      break;
    }
  }
  
  if (!session) return { list: [], count: 0, total: 0 };
  
  var students = getStudents(session.className);
  var studentMap = {};
  for (var s = 0; s < students.length; s++) {
    studentMap[students[s].studentId] = students[s];
  }
  
  var attData = attSheet.getDataRange().getValues();
  var checkedList = [];
  
  for (var j = 1; j < attData.length; j++) {
    var dStr = formatDateString(attData[j][0]);
    var sId = String(attData[j][1]);
    var status = String(attData[j][2]);
    
    if (dStr === session.date && (status === 'Present' || status === 'มาเรียน') && studentMap[sId]) {
      checkedList.push({
        studentId: sId,
        studentNo: studentMap[sId].studentNo,
        name: studentMap[sId].name,
        className: studentMap[sId].className,
        checkinTime: String(attData[j][3] || '-'),
        distanceMeters: attData[j][5] || 0,
        selfiePhoto: String(attData[j][7] || ''),
        checkinType: String(attData[j][8] || 'Self Check-in')
      });
    }
  }
  
  checkedList.reverse();
  
  return {
    list: checkedList,
    count: checkedList.length,
    total: students.length
  };
}

function getAttendanceRecords(dateStr, className) {
  var ss = getSpreadsheet();
  var attendanceSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!attendanceSheet) return {};
  
  var students = getStudents(className);
  var studentIds = {};
  for (var s = 0; s < students.length; s++) {
    studentIds[students[s].studentId] = students[s];
  }
  
  var attData = attendanceSheet.getDataRange().getValues();
  var records = {};
  
  for (var i = 1; i < attData.length; i++) {
    var rowDate = formatDateString(attData[i][0]);
    var sId = String(attData[i][1]);
    var status = String(attData[i][2]);
    
    if (rowDate === dateStr && studentIds[sId]) {
      records[sId] = {
        status: status,
        checkinTime: String(attData[i][3] || '-'),
        gpsLocation: String(attData[i][4] || '-'),
        distanceMeters: attData[i][5] || 0,
        deviceId: String(attData[i][6] || ''),
        selfiePhoto: String(attData[i][7] || ''),
        checkinType: String(attData[i][8] || 'Manual')
      };
    }
  }
  
  return records;
}

function saveAttendanceBatch(attendanceList) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) {
    initDatabase();
    sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  }
  
  if (!attendanceList || attendanceList.length === 0) {
    return { status: 'error', message: 'ไม่มีข้อมูลบันทึก' };
  }
  
  var existingData = sheet.getDataRange().getValues();
  var rowMap = {};
  
  for (var i = 1; i < existingData.length; i++) {
    var dStr = formatDateString(existingData[i][0]);
    var sId = String(existingData[i][1]);
    rowMap[dStr + '_' + sId] = i + 1;
  }
  
  var nowTimeStr = new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
  
  for (var k = 0; k < attendanceList.length; k++) {
    var item = attendanceList[k];
    var formattedDate = formatDateString(item.date);
    var key = formattedDate + '_' + item.studentId;
    
    if (rowMap[key]) {
      var rowIndex = rowMap[key];
      sheet.getRange(rowIndex, 3).setValue(item.status);
    } else {
      sheet.appendRow([
        formattedDate,
        item.studentId,
        item.status,
        nowTimeStr,
        '-',
        0,
        '-',
        '',
        'Manual (Teacher)'
      ]);
    }
  }
  
  return { status: 'success', message: 'บันทึกการเช็คชื่อเรียบร้อยแล้ว (' + attendanceList.length + ' คน)' };
}

function getDashboardData(targetDateStr) {
  var ss = getSpreadsheet();
  var students = getStudents();
  var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  
  var formattedTargetDate = targetDateStr ? formatDateString(targetDateStr) : formatDateString(new Date());
  
  var totalStudents = students.length;
  var presentCount = 0;
  var absentCount = 0;
  var sickCount = 0;
  var leaveCount = 0;
  var lateCount = 0;
  
  var todayRecords = {};
  
  if (attSheet && attSheet.getLastRow() > 1) {
    var attData = attSheet.getDataRange().getValues();
    for (var i = 1; i < attData.length; i++) {
      var d = formatDateString(attData[i][0]);
      var sId = String(attData[i][1]);
      var st = String(attData[i][2]);
      
      if (d === formattedTargetDate) {
        todayRecords[sId] = st;
      }
    }
  }
  
  for (var s = 0; s < students.length; s++) {
    var status = todayRecords[students[s].studentId] || 'Unchecked';
    if (status === 'Present' || status === 'มาเรียน') presentCount++;
    else if (status === 'Absent' || status === 'ขาดเรียน') absentCount++;
    else if (status === 'Sick' || status === 'ลาป่วย') sickCount++;
    else if (status === 'Leave' || status === 'ลากิจ') leaveCount++;
    else if (status === 'Late' || status === 'มาสาย') lateCount++;
  }
  
  var classMap = {};
  for (var j = 0; j < students.length; j++) {
    var cName = students[j].className;
    if (!classMap[cName]) {
      classMap[cName] = { total: 0, present: 0, absent: 0, sick: 0, leave: 0, late: 0, unchecked: 0 };
    }
    classMap[cName].total++;
    var cStatus = todayRecords[students[j].studentId] || 'Unchecked';
    if (cStatus === 'Present' || cStatus === 'มาเรียน') classMap[cName].present++;
    else if (cStatus === 'Absent' || cStatus === 'ขาดเรียน') classMap[cName].absent++;
    else if (cStatus === 'Sick' || cStatus === 'ลาป่วย') classMap[cName].sick++;
    else if (cStatus === 'Leave' || cStatus === 'ลากิจ') classMap[cName].leave++;
    else if (cStatus === 'Late' || cStatus === 'มาสาย') classMap[cName].late++;
    else classMap[cName].unchecked++;
  }
  
  var classSummary = [];
  for (var c in classMap) {
    var item = classMap[c];
    var rate = item.total > 0 ? ((item.present / item.total) * 100).toFixed(1) : 0;
    classSummary.push({
      className: c,
      total: item.total,
      present: item.present,
      absent: item.absent,
      sick: item.sick,
      leave: item.leave,
      late: item.late,
      unchecked: item.unchecked,
      rate: rate
    });
  }
  classSummary.sort(function(a, b) { return a.className.localeCompare(b.className, 'th'); });
  
  return {
    date: formattedTargetDate,
    totalStudents: totalStudents,
    present: presentCount,
    absent: absentCount,
    sick: sickCount,
    leave: leaveCount,
    late: lateCount,
    unchecked: totalStudents - (presentCount + absentCount + sickCount + leaveCount + lateCount),
    classSummary: classSummary,
    classes: getClassList()
  };
}

function getStudentStats(studentId) {
  var ss = getSpreadsheet();
  var students = getStudents();
  var student = null;
  studentId = String(studentId).trim();

  for (var i = 0; i < students.length; i++) {
    if (String(students[i].studentId).trim() === studentId) {
      student = students[i];
      break;
    }
  }
  
  if (!student) {
    return { status: 'error', message: 'ไม่พบข้อมูลนักเรียน' };
  }
  
  var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  var present = 0, absent = 0, sick = 0, leave = 0, late = 0;
  
  if (attSheet && attSheet.getLastRow() > 1) {
    var attData = attSheet.getDataRange().getValues();
    for (var j = 1; j < attData.length; j++) {
      var sId = String(attData[j][1]).trim();
      if (sId === studentId) {
        var status = String(attData[j][2]);
        if (status === 'Present' || status === 'มาเรียน') present++;
        else if (status === 'Absent' || status === 'ขาดเรียน') absent++;
        else if (status === 'Sick' || status === 'ลาป่วย') sick++;
        else if (status === 'Leave' || status === 'ลากิจ') leave++;
        else if (status === 'Late' || status === 'มาสาย') late++;
      }
    }
  }
  
  var totalDays = present + absent + sick + leave + late;
  var attendancePercentage = totalDays > 0 ? (((present + late) / totalDays) * 100).toFixed(2) : '100.00';
  
  return {
    student: student,
    totalDays: totalDays,
    present: present,
    absent: absent,
    sick: sick,
    leave: leave,
    late: late,
    attendancePercentage: attendancePercentage
  };
}

function addStudent(studentData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) {
    initDatabase();
    sheet = ss.getSheetByName(SHEET_STUDENTS);
  }
  
  var newId = studentData.studentId ? String(studentData.studentId).trim() : ('STD' + String(new Date().getTime()).slice(-6));
  sheet.appendRow([newId, studentData.studentNo, studentData.name, studentData.className, '']);
  return { status: 'success', message: 'เพิ่มนักเรียนเรียบร้อยแล้ว', studentId: newId };
}

function updateStudent(studentData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางนักเรียน' };
  
  var studentId = String(studentData.studentId).trim();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === studentId) {
      var row = i + 1;
      sheet.getRange(row, 2).setValue(studentData.studentNo);
      sheet.getRange(row, 3).setValue(studentData.name);
      sheet.getRange(row, 4).setValue(studentData.className);
      return { status: 'success', message: 'แก้ไขข้อมูลเรียบร้อยแล้ว' };
    }
  }
  return { status: 'error', message: 'ไม่พบนักเรียนในระบบ' };
}

function deleteStudent(studentId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางนักเรียน' };
  
  studentId = String(studentId).trim();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === studentId) {
      sheet.deleteRow(i + 1);
      return { status: 'success', message: 'ลบข้อมูลนักเรียนเรียบร้อยแล้ว' };
    }
  }
  return { status: 'error', message: 'ไม่พบนักเรียนในระบบ' };
}

/**
 * Deletes all students from the Students sheet
 */
function deleteAllStudents() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางนักเรียน' };
  
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  
  return { status: 'success', message: 'ลบรายชื่อนักเรียนทั้งหมดเรียบร้อยแล้ว' };
}

function formatDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = String(val.getMonth() + 1);
    if (m.length < 2) m = '0' + m;
    var d = String(val.getDate());
    if (d.length < 2) d = '0' + d;
    return y + '-' + m + '-' + d;
  }
  
  var str = String(val).trim();
  if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
    return str.substring(0, 10);
  }
  
  if (str.indexOf('/') > -1) {
    var parts = str.split('/');
    if (parts.length === 3) {
      var day = parts[0].length === 1 ? '0' + parts[0] : parts[0];
      var month = parts[1].length === 1 ? '0' + parts[1] : parts[1];
      var year = parts[2];
      if (parseInt(year, 10) > 2500) {
        year = String(parseInt(year, 10) - 543);
      } else if (year.length === 2) {
        year = '20' + year;
      }
      return year + '-' + month + '-' + day;
    }
  }
  return str;
}
