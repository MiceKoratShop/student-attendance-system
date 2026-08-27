
/**
 * ฟังก์ชันสำหรับกด "▷ เรียกใช้" (Run) ในหน้า Google Apps Script Editor 
 * เพื่อขออนุญาตสิทธิ์การส่งข้อมูลภายนอก (UrlFetchApp / LINE API) และเข้าถึง Google Sheets ครั้งเดียว
 */
/**
 * ฟังก์ชันสำหรับบังคับให้ Google Apps Script เด้งหน้าต่าง "ตรวจสอบสิทธิ์" (Authorization Required)
 * ห้ามครอบ try-catch เพื่อให้ Google Engine ตรวจพบ Scope ที่ขาดและเด้ง Pop-up ทันที
 */
function authorizePermissions() {
  UrlFetchApp.fetch('https://www.google.com');
  SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties();
  ScriptApp.getProjectTriggers();
  Logger.log('สิทธิ์ทั้งหมดได้รับการอนุมัติเรียบร้อยแล้ว (UrlFetchApp, Sheets, Triggers)');
}

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
var SHEET_ACTIVITIES = 'Activities';

var ADMIN_PASSWORD = 'admin888';

/**
 * Serves the HTML Web Application or handles API GET requests
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var data = null;
    var args = [];
    if (e.parameter.args) {
      try { args = JSON.parse(e.parameter.args); } catch (ex) {}
    }
    if (e.parameter.data) {
      try {
        data = JSON.parse(e.parameter.data);
      } catch (ex) {
        data = e.parameter.data;
      }
    }
    var result = handleApiRouter(action, data, args);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('ระบบเช็คชื่อนักเรียนออนไลน์ - Anti-Cheat Edition')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles API POST requests (For GitHub Pages and external HTTP clients)
 */
function doPost(e) {
  var action = '';
  var data = null;
  var args = [];

  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      // Automatic LINE Bot Webhook Handler (Auto-detects Group ID when bot joins group or receives message)
      if (body && body.events && Array.isArray(body.events)) {
        return handleLineBotWebhook(body.events);
      }
      action = body.action || '';
      data = body.data !== undefined ? body.data : (body.payload !== undefined ? body.payload : body);
      args = body.args || [];
    } catch (ex) {
      action = e.parameter ? e.parameter.action : '';
      data = e.parameter;
    }
  } else if (e && e.parameter) {
    action = e.parameter.action || '';
    data = e.parameter.data ? JSON.parse(e.parameter.data) : e.parameter;
    if (e.parameter.args) {
      try { args = JSON.parse(e.parameter.args); } catch (ex) {}
    }
  }

  var result = handleApiRouter(action, data, args);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Central API Router for all actions
 */
function handleApiRouter(action, data, args) {
  args = args || [];
  function arg(index, key, fallback) {
    if (args && args.length > index && args[index] !== undefined && args[index] !== null) {
      return args[index];
    }
    if (data && typeof data === 'object' && key && data[key] !== undefined && data[key] !== null) {
      return data[key];
    }
    if (index === 0 && data !== undefined && data !== null && typeof data !== 'object') {
      return data;
    }
    return fallback;
  }

  try {
    switch (action) {
      case 'initDatabase':
        return initDatabase();
      case 'verifyAdminPassword':
        return verifyAdminPassword(arg(0, 'password', ''));
      case 'getClassList':
        return getClassList();
      case 'addClass':
        return addClass(arg(0, 'className', ''));
      case 'deleteClass':
        return deleteClass(arg(0, 'className', ''));
      case 'getActivityList':
        return getActivityList();
      case 'addActivity':
        return addActivity(arg(0, 'name', ''), arg(1, 'type', 'general'));
      case 'deleteActivity':
        return deleteActivity(arg(0, 'name', ''));
      case 'getStudents':
        return getStudents(arg(0, 'classFilter', 'all'));
      case 'addStudent':
        return addStudent(arg(0, 'studentData', data));
      case 'updateStudent':
        return updateStudent(arg(0, 'studentData', data));
      case 'deleteStudent':
        var sIdParam = arg(0, 'studentId', data);
        if (typeof sIdParam === 'object' && sIdParam !== null) {
          sIdParam = sIdParam.studentId || sIdParam.id || (Array.isArray(sIdParam) ? sIdParam[0] : '');
        }
        return deleteStudent(sIdParam);
      case 'deleteStudentsBatch':
        var bIdsParam = arg(0, 'studentIds', data);
        return deleteStudentsBatch(bIdsParam);
      case 'deleteAllStudents':
        return deleteAllStudents();
      case 'importStudentsBatch':
        return importStudentsBatch(arg(0, 'studentsList', data));
      case 'createCheckinSession':
        return createCheckinSession(arg(0, 'config', data));
      case 'getActiveSession':
        return getActiveSession(arg(0, 'className', ''), arg(1, 'sessionId', ''));
      case 'getActiveSessions':
        return getActiveSessions();
      case 'closeCheckinSession':
        return closeCheckinSession(arg(0, 'sessionId', ''));
      case 'validateAndCheckinStudent':
        return validateAndCheckinStudent(arg(0, 'payload', data));
      case 'getSessionLiveAttendance':
        return getSessionLiveAttendance(arg(0, 'sessionId', ''));
      case 'deleteLiveCheckinRecord':
        return deleteLiveCheckinRecord(arg(0, 'studentId', ''), arg(1, 'dateStr', ''));
      case 'getAttendanceRecords':
        return getAttendanceRecords(arg(0, 'dateStr', ''), arg(1, 'className', ''), arg(2, 'activityName', ''));
      case 'saveAttendanceBatch':
        return saveAttendanceBatch(arg(0, 'attendanceList', data), arg(1, 'activityName', ''));
      case 'getDashboardData':
        return getDashboardData(arg(0, 'targetDateStr', ''), arg(1, 'className', 'all'), arg(2, 'activityName', 'all'));
      case 'getStudentStats':
        return getStudentStats(arg(0, 'studentId', ''));
      case 'getLineConfig':
        return getLineConfig();
      case 'saveLineConfig':
        return saveLineConfig(arg(0, 'config', data));
      case 'getDetectedLineGroups':
        return getDetectedLineGroups();
      case 'testLineConnection':
        return testLineConnection(arg(0, 'token', ''), arg(1, 'webhookUrl', ''), arg(2, 'lineGroupId', ''), arg(3, 'customMsg', ''));
      case 'sendLineReport':
        return sendLineReport(arg(0, 'options', data));
      case 'setupDailyLineTrigger':
        return setupDailyLineTrigger(arg(0, 'enable', false), arg(1, 'targetHour', 16));
      default:
        return { status: 'error', message: 'Unknown action: ' + action };
    }
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
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
      'Date', 'StudentID', 'Status', 'CheckinTime', 'GPSLocation', 'DistanceMeters', 'DeviceID', 'SelfiePhoto', 'CheckinType', 'SessionID', 'Note'
    ]);
  } else {
    attendanceSheet.clear();
    attendanceSheet.appendRow([
      'Date', 'StudentID', 'Status', 'CheckinTime', 'GPSLocation', 'DistanceMeters', 'DeviceID', 'SelfiePhoto', 'CheckinType', 'SessionID', 'Note'
    ]);
  }
  
  // 3. Sessions Sheet (For Projector and Live Check-in)
  var sessionSheet = ss.getSheetByName(SHEET_SESSIONS);
  var sessionHeaders = [
    'SessionID', 'ClassName', 'Date', 'CreatedAt', 'ExpiresAt', 'Lat', 'Lng', 'RadiusMeters', 'SecretSeed', 'RequireSelfie', 'RequireGPS', 'RequireDeviceBinding', 'Status', 'OTPIntervalSeconds', 'ActivityName'
  ];
  if (!sessionSheet) {
    sessionSheet = ss.insertSheet(SHEET_SESSIONS);
    sessionSheet.appendRow(sessionHeaders);
  } else {
    sessionSheet.clear();
    sessionSheet.appendRow(sessionHeaders);
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
  
  // 5. Activities & Subjects Sheet
  var actSheet = ss.getSheetByName(SHEET_ACTIVITIES);
  if (!actSheet) {
    actSheet = ss.insertSheet(SHEET_ACTIVITIES);
  } else {
    actSheet.clear();
  }
  actSheet.appendRow(['Name', 'Type', 'CreatedAt']);
  var defaultActivities = [
    ['เช็คชื่อมาเรียน', 'general', new Date()],
    ['เช็คชื่อเข้าแถว', 'general', new Date()],
    ['เช็คชื่อคาบโฮมรูม', 'general', new Date()],
    ['เช็คชื่อประจำวัน', 'general', new Date()],
    ['คาบที่ 1 (08:30 - 09:30)', 'period', new Date()],
    ['คาบที่ 2 (09:30 - 10:30)', 'period', new Date()],
    ['คาบที่ 3 (10:30 - 11:30)', 'period', new Date()],
    ['คาบที่ 4 (11:30 - 12:30)', 'period', new Date()],
    ['คาบที่ 5 (13:00 - 14:00)', 'period', new Date()],
    ['คาบที่ 6 (14:00 - 15:00)', 'period', new Date()],
    ['คาบที่ 7 (15:00 - 16:00)', 'period', new Date()],
    ['คาบเช้า', 'period', new Date()],
    ['คาบบ่าย', 'period', new Date()],
    ['คณิตศาสตร์', 'subject', new Date()],
    ['ภาษาไทย', 'subject', new Date()],
    ['ภาษาอังกฤษ', 'subject', new Date()],
    ['วิทยาศาสตร์', 'subject', new Date()],
    ['สังคมศึกษา', 'subject', new Date()],
    ['การเขียนโปรแกรม', 'subject', new Date()],
    ['สุขศึกษาและพลศึกษา', 'subject', new Date()],
    ['ศิลปะ', 'subject', new Date()],
    ['การบัญชีเบื้องต้น', 'subject', new Date()],
    ['คอมพิวเตอร์กราฟิก', 'subject', new Date()],
    ['กิจกรรมวันไหว้ครู', 'activity', new Date()],
    ['กิจกรรมจิตอาสา', 'activity', new Date()],
    ['กิจกรรมลูกเสือ/เนตรนารี', 'activity', new Date()],
    ['กิจกรรมอบรมคุณธรรม', 'activity', new Date()],
    ['กิจกรรมชุมนุม / ชมรม', 'activity', new Date()],
    ['กิจกรรมวันภาษาไทย', 'activity', new Date()],
    ['กิจกรรมวันวิทยาศาสตร์', 'activity', new Date()]
  ];
  actSheet.getRange(2, 1, defaultActivities.length, 3).setValues(defaultActivities);

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
 * Helper to clear CacheService
 */
function clearScriptCache() {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('cache_class_list');
    cache.remove('cache_students_all');
    cache.remove('cache_activity_list');
  } catch (e) {}
}

/**
 * Gets list of all activities and subjects
 */
function getActivityList() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('cache_activity_list');
    if (cached) {
      return { status: 'success', activities: JSON.parse(cached) };
    }
  } catch (e) {}

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ACTIVITIES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ACTIVITIES);
    sheet.appendRow(['Name', 'Type', 'CreatedAt']);
    var defaults = [
      ['เช็คชื่อมาเรียน', 'general', new Date()],
      ['เช็คชื่อเข้าแถว', 'general', new Date()],
      ['เช็คชื่อคาบโฮมรูม', 'general', new Date()],
      ['เช็คชื่อประจำวัน', 'general', new Date()],
      ['คาบที่ 1 (08:30 - 09:30)', 'period', new Date()],
      ['คาบที่ 2 (09:30 - 10:30)', 'period', new Date()],
      ['คาบที่ 3 (10:30 - 11:30)', 'period', new Date()],
      ['คาบที่ 4 (11:30 - 12:30)', 'period', new Date()],
      ['คาบที่ 5 (13:00 - 14:00)', 'period', new Date()],
      ['คาบที่ 6 (14:00 - 15:00)', 'period', new Date()],
      ['คาบที่ 7 (15:00 - 16:00)', 'period', new Date()],
      ['คาบเช้า', 'period', new Date()],
      ['คาบบ่าย', 'period', new Date()],
      ['คณิตศาสตร์', 'subject', new Date()],
      ['ภาษาไทย', 'subject', new Date()],
      ['ภาษาอังกฤษ', 'subject', new Date()],
      ['วิทยาศาสตร์', 'subject', new Date()],
      ['สังคมศึกษา', 'subject', new Date()],
      ['การเขียนโปรแกรม', 'subject', new Date()],
      ['สุขศึกษาและพลศึกษา', 'subject', new Date()],
      ['ศิลปะ', 'subject', new Date()],
      ['กิจกรรมวันไหว้ครู', 'activity', new Date()],
      ['กิจกรรมจิตอาสา', 'activity', new Date()],
      ['กิจกรรมลูกเสือ/เนตรนารี', 'activity', new Date()],
      ['กิจกรรมอบรมคุณธรรม', 'activity', new Date()],
      ['กิจกรรมชุมนุม / ชมรม', 'activity', new Date()]
    ];
    sheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }

  var list = [];
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][0] || '').trim();
      var type = String(data[i][1] || 'general').trim();
      if (name) {
        list.push({ name: name, type: type });
      }
    }
  }

  if (list.length === 0) {
    list = [
      { name: 'เช็คชื่อมาเรียน', type: 'general' },
      { name: 'เช็คชื่อเข้าแถว', type: 'general' },
      { name: 'เช็คชื่อคาบโฮมรูม', type: 'general' },
      { name: 'เช็คชื่อประจำวัน', type: 'general' },
      { name: 'คาบที่ 1 (08:30 - 09:30)', type: 'period' },
      { name: 'คาบที่ 2 (09:30 - 10:30)', type: 'period' },
      { name: 'คาบที่ 3 (10:30 - 11:30)', type: 'period' },
      { name: 'คาบที่ 4 (11:30 - 12:30)', type: 'period' },
      { name: 'คาบที่ 5 (13:00 - 14:00)', type: 'period' },
      { name: 'คาบเช้า', type: 'period' },
      { name: 'คาบบ่าย', type: 'period' },
      { name: 'คณิตศาสตร์', type: 'subject' },
      { name: 'ภาษาไทย', type: 'subject' },
      { name: 'ภาษาอังกฤษ', type: 'subject' },
      { name: 'วิทยาศาสตร์', type: 'subject' },
      { name: 'สังคมศึกษา', type: 'subject' },
      { name: 'การเขียนโปรแกรม', type: 'subject' },
      { name: 'กิจกรรมวันไหว้ครู', type: 'activity' },
      { name: 'กิจกรรมจิตอาสา', type: 'activity' },
      { name: 'กิจกรรมลูกเสือ/เนตรนารี', type: 'activity' }
    ];
  }

  try {
    var cache = CacheService.getScriptCache();
    cache.put('cache_activity_list', JSON.stringify(list), 300);
  } catch (e) {}

  return { status: 'success', activities: list };
}

/**
 * Adds a new activity or subject
 */
function addActivity(name, type) {
  name = String(name || '').trim();
  type = String(type || 'general').trim().toLowerCase();
  if (type === 'assembly' || type === 'homeroom' || type === 'daily') type = 'general';
  if (!name) return { status: 'error', message: 'กรุณาระบุชื่อรายวิชาหรือกิจกรรม' };

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ACTIVITIES);
  if (!sheet) {
    getActivityList();
    sheet = ss.getSheetByName(SHEET_ACTIVITIES);
  }

  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === name.toLowerCase()) {
        return { status: 'error', message: 'รายวิชาหรือกิจกรรมนี้มีอยู่ในระบบแล้ว' };
      }
    }
  }

  sheet.appendRow([name, type, new Date()]);
  SpreadsheetApp.flush();
  clearScriptCache();
  return { status: 'success', message: 'เพิ่ม ' + name + ' เรียบร้อยแล้ว', activities: getActivityList().activities };
}

/**
 * Deletes an activity or subject
 */
function deleteActivity(name) {
  name = String(name || '').trim();
  if (!name) return { status: 'error', message: 'กรุณาระบุชื่อที่ต้องการลบ' };

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ACTIVITIES);
  if (!sheet) return { status: 'error', message: 'ไม่พบตาราง Activities' };

  if (sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === name.toLowerCase()) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        clearScriptCache();
        return { status: 'success', message: 'ลบ ' + name + ' เรียบร้อยแล้ว', activities: getActivityList().activities };
      }
    }
  }
  return { status: 'error', message: 'ไม่พบรายการที่ต้องการลบ' };
}

/**
 * Gets list of all classes (Empty array if no classes added yet)
 */
function getClassList() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('cache_class_list');
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {}

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

  try {
    CacheService.getScriptCache().put('cache_class_list', JSON.stringify(classes), 600);
  } catch (e) {}

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
  SpreadsheetApp.flush();
  clearScriptCache();
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
  SpreadsheetApp.flush();
  clearScriptCache();

  return { status: 'success', message: 'ลบห้องเรียน "' + className + '" เรียบร้อยแล้ว', classList: getClassList() };
}

/**
 * Helper to check if a student classroom matches a single or comma-separated class filter
 */

/**
 * Helper to check if an activity name matches a single or comma-separated activity filter
 */
function isActivityMatch(filter, activityName) {
  if (!filter || filter === 'all' || filter === 'ทุกกิจกรรม' || filter === 'undefined') return true;
  var sAct = String(activityName || 'เช็คชื่อมาเรียน').trim().toLowerCase();
  var parts = String(filter).split(',');
  for (var p = 0; p < parts.length; p++) {
    var item = parts[p].trim().toLowerCase();
    if (item === 'all' || item === 'ทุกกิจกรรม' || item === sAct || sAct.indexOf(item) > -1 || item.indexOf(sAct) > -1) {
      return true;
    }
  }
  return false;
}

function isClassMatch(filter, studentClass) {
  if (!filter || filter === 'all' || filter === 'ทุกห้องเรียน' || filter === 'undefined') return true;
  var sClass = String(studentClass || '').trim().toLowerCase();
  if (!sClass) return false;
  var parts = String(filter).split(',');
  for (var p = 0; p < parts.length; p++) {
    var item = parts[p].trim().toLowerCase();
    if (item === 'all' || item === 'ทุกห้องเรียน' || item === sClass) {
      return true;
    }
  }
  return false;
}

/**
 * Gets list of all students (optionally filtered by class or multi-classes)
 */
function getStudents(classFilter) {
  var allStudents = null;
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('cache_students_all');
    if (cached) {
      allStudents = JSON.parse(cached);
    }
  } catch (e) {}

  if (!allStudents) {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_STUDENTS);
    if (!sheet) {
      initDatabase();
      sheet = ss.getSheetByName(SHEET_STUDENTS);
    }
    
    var data = sheet.getDataRange().getValues();
    allStudents = [];
    if (data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var studentId = String(row[0]).trim();
        var studentNo = String(row[1]).trim();
        var name = String(row[2]).trim();
        var className = String(row[3]).trim();
        var deviceId = row[4] ? String(row[4]).trim() : '';
        
        if (!name) continue;
        allStudents.push({
          studentId: studentId,
          studentNo: studentNo,
          name: name,
          className: className,
          deviceId: deviceId
        });
      }
    }

    try {
      CacheService.getScriptCache().put('cache_students_all', JSON.stringify(allStudents), 600);
    } catch (e) {}
  }

  var students = [];
  for (var k = 0; k < allStudents.length; k++) {
    var s = allStudents[k];
    if (isClassMatch(classFilter, s.className)) {
      students.push(s);
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
  if (classSheet.getLastRow() > 1) {
    var cData = classSheet.getDataRange().getValues();
    for (var c = 1; c < cData.length; c++) {
      classesMap[String(cData[c][0]).trim().toLowerCase()] = true;
    }
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
    var classKey = sClass.toLowerCase();
    if (sClass && !classesMap[classKey] && !newClasses[classKey]) {
      newClasses[classKey] = true;
      classSheet.appendRow([sClass, new Date().toISOString()]);
      classesMap[classKey] = true;
    }

    rowsToAdd.push([sId, sNo, sName, sClass, '']);
  }

  if (rowsToAdd.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
    SpreadsheetApp.flush();
  }

  clearScriptCache();

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
  config = config || {};
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SESSIONS);
  if (!sheet) {
    initDatabase();
    sheet = ss.getSheetByName(SHEET_SESSIONS);
  }
  
  if (sheet.getMaxColumns() < 17) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 17 - sheet.getMaxColumns());
  }

  var now = new Date();
  var nowMs = now.getTime();
  var durationMinutes = config.durationMinutes || 15;
  var otpIntervalSeconds = parseInt(config.otpIntervalSeconds, 10) || 15;
  if (otpIntervalSeconds < 5) otpIntervalSeconds = 5;
  var activityName = String(config.activityName || 'เช็คชื่อมาเรียน').trim();
  if (!activityName) activityName = 'เช็คชื่อมาเรียน';

  // Check currently active sessions
  var activeSessions = getActiveSessions();
  if (activeSessions && activeSessions.length >= 3) {
    return {
      status: 'error',
      message: 'เปิดคาบเช็คชื่อพร้อมกันได้สูงสุด 3 คาบ กรุณาปิดคาบเก่าก่อนหรือรอให้หมดเวลา',
      activeSessions: activeSessions
    };
  }

  var expiresAtMs = nowMs + (durationMinutes * 60 * 1000);
  var sessionId = 'SES' + String(nowMs).slice(-6);
  var secretSeed = 'SEC_' + Math.random().toString(36).substring(2, 10).toUpperCase();
  var dateStr = formatDateString(now);
  
  // Calculate Late Cut-off Timestamp (lateAtMs)
  var lateAtMs = 0;
  var lateAfterMinutes = 0;
  if (config.lateAfterMinutes && parseInt(config.lateAfterMinutes, 10) > 0) {
    lateAfterMinutes = parseInt(config.lateAfterMinutes, 10);
    lateAtMs = nowMs + (lateAfterMinutes * 60 * 1000);
  } else if (config.lateCutoffTime) {
    var timeParts = String(config.lateCutoffTime).split(':');
    if (timeParts.length >= 2) {
      var lateDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), parseInt(timeParts[0], 10), parseInt(timeParts[1], 10), 0);
      lateAtMs = lateDate.getTime();
      if (lateAtMs > nowMs) {
        lateAfterMinutes = Math.round((lateAtMs - nowMs) / (60 * 1000));
      }
    }
  } else if (config.lateAt) {
    lateAtMs = new Date(config.lateAt).getTime();
    if (lateAtMs > nowMs) {
      lateAfterMinutes = Math.round((lateAtMs - nowMs) / (60 * 1000));
    }
  }

  var lateAtIso = lateAtMs > 0 ? new Date(lateAtMs).toISOString() : '';

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
    'Active',
    otpIntervalSeconds,
    activityName,
    lateAtIso,
    lateAfterMinutes
  ];
  
  sheet.appendRow(newRow);
  SpreadsheetApp.flush();
  
  var newSession = {
    sessionId: sessionId,
    className: config.className || 'all',
    activityName: activityName,
    date: dateStr,
    createdAt: now.toISOString(),
    expiresAt: new Date(expiresAtMs).toISOString(),
    lat: config.lat || 0,
    lng: config.lng || 0,
    radiusMeters: config.radiusMeters || 50,
    secretSeed: secretSeed,
    otpIntervalSeconds: otpIntervalSeconds,
    requireSelfie: config.requireSelfie !== false,
    requireGPS: config.requireGPS !== false,
    requireDeviceBinding: config.requireDeviceBinding !== false,
    status: 'Active',
    lateAt: lateAtIso,
    lateAtMs: lateAtMs,
    lateAfterMinutes: lateAfterMinutes,
    serverTimeMs: nowMs
  };

  var allActive = getActiveSessions();

  return {
    status: 'success',
    message: 'เปิดคาบเช็คชื่อเรียบร้อยแล้ว',
    session: newSession,
    activeSessions: allActive
  };
}

function getActiveSessions() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SESSIONS);
  if (!sheet) return [];
  
  if (sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var nowMs = new Date().getTime();
  var activeList = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var sId = String(row[0] || '').trim();
    var cName = String(row[1] || 'all').trim();
    var expStr = String(row[4] || '').trim();
    var expMs = new Date(expStr).getTime();
    var status = String(row[12] || '').trim();
    var otpInterval = parseInt(row[13], 10) || 15;
    if (otpInterval < 5) otpInterval = 15;
    var actName = String(row[14] || 'เช็คชื่อมาเรียน').trim();
    var lateAtStr = String(row[15] || '').trim();
    var lateAtMs = lateAtStr ? new Date(lateAtStr).getTime() : 0;
    var lateAfterMins = parseInt(row[16], 10) || 0;
    
    if (status === 'Active') {
      if (nowMs > expMs) {
        try {
          if (sheet.getMaxColumns() >= 13) {
            sheet.getRange(i + 1, 13).setValue('Expired');
          }
        } catch (e) {}
        continue;
      }
      
      activeList.push({
        sessionId: sId,
        className: cName || 'all',
        activityName: actName || 'เช็คชื่อมาเรียน',
        date: formatDateString(row[2]),
        createdAt: String(row[3]),
        expiresAt: expStr,
        lat: parseFloat(row[5]) || 0,
        lng: parseFloat(row[6]) || 0,
        radiusMeters: parseInt(row[7], 10) || 50,
        secretSeed: String(row[8]),
        otpIntervalSeconds: otpInterval,
        requireSelfie: String(row[9]) === 'YES',
        requireGPS: String(row[10]) === 'YES',
        requireDeviceBinding: String(row[11]) === 'YES',
        status: 'Active',
        lateAt: lateAtStr,
        lateAtMs: lateAtMs,
        lateAfterMinutes: lateAfterMins,
        serverTimeMs: nowMs
      });
    }
  }
  
  try {
    SpreadsheetApp.flush();
  } catch (e) {}
  
  return activeList;
}

function getActiveSession(className, sessionId) {
  var activeList = getActiveSessions();
  if (!activeList || activeList.length === 0) return null;

  if (sessionId) {
    for (var i = 0; i < activeList.length; i++) {
      if (activeList[i].sessionId === String(sessionId).trim()) {
        return activeList[i];
      }
    }
  }

  if (className && className !== 'all') {
    for (var j = 0; j < activeList.length; j++) {
      var cName = activeList[j].className;
      if (cName === 'all' || isClassMatch(cName, className) || isClassMatch(className, cName)) {
        return activeList[j];
      }
    }
  }

  // Return the most recent active session
  return activeList[activeList.length - 1] || null;
}

function closeCheckinSession(sessionId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SESSIONS);
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางคาบเรียน' };
  
  var data = sheet.getDataRange().getValues();
  var cleanId = String(sessionId || '').trim();
  var closedCount = 0;

  for (var i = 1; i < data.length; i++) {
    var rowSessionId = String(data[i][0] || '').trim();
    var rowStatus = String(data[i][12] || '').trim();
    
    if (cleanId && cleanId !== 'all' && cleanId !== 'undefined') {
      if (rowSessionId === cleanId) {
        sheet.getRange(i + 1, 13).setValue('Closed');
        closedCount++;
      }
    } else {
      if (rowStatus === 'Active') {
        sheet.getRange(i + 1, 13).setValue('Closed');
        closedCount++;
      }
    }
  }
  SpreadsheetApp.flush();
  var remainingActive = getActiveSessions();
  return {
    status: 'success',
    message: cleanId && cleanId !== 'all' ? 'ปิดคาบเช็คชื่อ (' + cleanId + ') เรียบร้อยแล้ว' : 'ปิดคาบเช็คชื่อทั้งหมดเรียบร้อยแล้ว',
    closedCount: closedCount,
    activeSessions: remainingActive
  };
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
  
  var activeSession = getActiveSession(studentObj.className, sessionId);
  if (!activeSession) {
    return { status: 'error', message: 'ไม่มีคาบเรียนที่เปิดรับเช็คชื่อสำหรับห้องนี้ในขณะนี้' };
  }
  
  var now = new Date();
  var nowMs = now.getTime();
  var intervalSec = activeSession.otpIntervalSeconds || 15;
  var intervalMs = intervalSec * 1000;
  
  var validOtpCurrent = generateTOTP(activeSession.secretSeed, intervalSec, nowMs);
  var validOtpPrev = generateTOTP(activeSession.secretSeed, intervalSec, nowMs - intervalMs);
  
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
  var currentActivity = activeSession.activityName || 'เช็คชื่อมาเรียน';
  
  if (activeSession.requireDeviceBinding && deviceId) {
    for (var i = 1; i < attData.length; i++) {
      var rowDate = formatDateString(attData[i][0]);
      var rowStd = String(attData[i][1]);
      var rowDev = String(attData[i][6] || '');
      var rowAct = String(attData[i][8] || 'เช็คชื่อมาเรียน');
      var rowSes = String(attData[i][9] || '');
      
      if (rowDate === dateStr && (rowSes === activeSession.sessionId || rowAct === currentActivity)) {
        if (rowDev === deviceId && rowStd !== studentId) {
          return {
            status: 'error',
            message: 'อุปกรณ์นี้ (Device ID) ได้ใช้เช็คชื่อ (' + currentActivity + ') ให้นักศึกษาคนอื่นไปแล้ว ห้ามเช็คชื่อแทนกัน'
          };
        }
        if (rowStd === studentId && (attData[i][2] === 'Present' || attData[i][2] === 'Late')) {
          return {
            status: 'error',
            message: 'นักศึกษาท่านนี้ได้ทำการเช็คชื่อ (' + currentActivity + ') ในคาบนี้เรียบร้อยแล้ว'
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
  
  // Determine Present vs Late status
  var checkinStatus = 'Present';
  var isLate = false;
  if (activeSession.lateAtMs && activeSession.lateAtMs > 0 && nowMs > activeSession.lateAtMs) {
    checkinStatus = 'Late';
    isLate = true;
  }

  var existingRowIndex = -1;
  for (var r = 1; r < attData.length; r++) {
    var rDate = formatDateString(attData[r][0]);
    var rStd = String(attData[r][1]);
    var rAct = String(attData[r][8] || 'เช็คชื่อมาเรียน');
    var rSes = String(attData[r][9] || '');
    if (rDate === dateStr && rStd === studentId && (rSes === activeSession.sessionId || rAct === currentActivity)) {
      existingRowIndex = r + 1;
      break;
    }
  }
  
  if (existingRowIndex > 0) {
    attSheet.getRange(existingRowIndex, 3).setValue(checkinStatus);
    attSheet.getRange(existingRowIndex, 4).setValue(checkinTimeStr);
    attSheet.getRange(existingRowIndex, 5).setValue(gpsLocStr);
    attSheet.getRange(existingRowIndex, 6).setValue(distanceMeters);
    attSheet.getRange(existingRowIndex, 7).setValue(deviceId);
    attSheet.getRange(existingRowIndex, 8).setValue(selfieBase64);
    attSheet.getRange(existingRowIndex, 9).setValue(currentActivity);
    attSheet.getRange(existingRowIndex, 10).setValue(activeSession.sessionId || '');
  } else {
    attSheet.appendRow([
      dateStr,
      studentId,
      checkinStatus,
      checkinTimeStr,
      gpsLocStr,
      distanceMeters,
      deviceId,
      selfieBase64,
      currentActivity,
      activeSession.sessionId || ''
    ]);
  }
  SpreadsheetApp.flush();
  
  var statusMsg = isLate ? 'เช็คชื่อ (' + currentActivity + ') สำเร็จ! (สถานะ: มาสาย เนื่องจากเกินเวลาที่กำหนด)' : 'เช็คชื่อ (' + currentActivity + ') สำเร็จเรียบร้อยแล้ว!';

  return {
    status: 'success',
    message: statusMsg,
    checkinStatus: checkinStatus,
    isLate: isLate,
    student: {
      studentId: studentObj.studentId,
      studentNo: studentObj.studentNo,
      name: studentObj.name,
      className: studentObj.className,
      activityName: currentActivity,
      checkinTime: checkinTimeStr,
      checkinStatus: checkinStatus,
      isLate: isLate,
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
  
  if (!sessionSheet || !attSheet) return { list: [], count: 0, countPresent: 0, countLate: 0, total: 0 };
  
  var session = null;
  var sData = sessionSheet.getDataRange().getValues();
  for (var i = 1; i < sData.length; i++) {
    if (String(sData[i][0]) === String(sessionId)) {
      var lateAtStr = String(sData[i][15] || '').trim();
      session = {
        sessionId: sData[i][0],
        className: sData[i][1],
        date: formatDateString(sData[i][2]),
        status: sData[i][12],
        activityName: String(sData[i][14] || 'เช็คชื่อมาเรียน'),
        lateAt: lateAtStr,
        lateAtMs: lateAtStr ? new Date(lateAtStr).getTime() : 0,
        lateAfterMinutes: parseInt(sData[i][16], 10) || 0
      };
      break;
    }
  }
  
  if (!session) return { list: [], count: 0, countPresent: 0, countLate: 0, total: 0 };
  
  var students = getStudents(session.className);
  var studentMap = {};
  for (var s = 0; s < students.length; s++) {
    studentMap[students[s].studentId] = students[s];
  }
  
  var attData = attSheet.getDataRange().getValues();
  var checkedList = [];
  var countPresent = 0;
  var countLate = 0;
  
  for (var j = 1; j < attData.length; j++) {
    var dStr = formatDateString(attData[j][0]);
    var sId = String(attData[j][1]);
    var status = String(attData[j][2]);
    var rowAct = String(attData[j][8] || 'เช็คชื่อมาเรียน');
    var rowSes = String(attData[j][9] || '');
    
    if (dStr === session.date && (status === 'Present' || status === 'มาเรียน' || status === 'Late' || status === 'มาสาย') && studentMap[sId]) {
      if (rowSes === session.sessionId || rowAct === session.activityName || (!rowSes && session.activityName === 'เช็คชื่อมาเรียน')) {
        var isRowLate = (status === 'Late' || status === 'มาสาย');
        if (isRowLate) countLate++;
        else countPresent++;

        checkedList.push({
          studentId: sId,
          studentNo: studentMap[sId].studentNo,
          name: studentMap[sId].name,
          className: studentMap[sId].className,
          activityName: rowAct,
          status: isRowLate ? 'Late' : 'Present',
          isLate: isRowLate,
          checkinTime: String(attData[j][3] || '-'),
          distanceMeters: attData[j][5] || 0,
          selfiePhoto: String(attData[j][7] || ''),
          checkinType: rowAct
        });
      }
    }
  }
  
  checkedList.reverse();
  
  return {
    list: checkedList,
    count: checkedList.length,
    countPresent: countPresent,
    countLate: countLate,
    total: students.length,
    activityName: session.activityName,
    lateAt: session.lateAt,
    lateAtMs: session.lateAtMs,
    lateAfterMinutes: session.lateAfterMinutes
  };
}

/**
 * Deletes a live check-in attendance record for a student on a specific date (e.g. invalid photo / cheat attempt)
 */
function deleteLiveCheckinRecord(studentId, dateStr) {
  var ss = getSpreadsheet();
  var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!attSheet) return { status: 'error', message: 'ไม่พบชีต Attendance' };

  var data = attSheet.getDataRange().getValues();
  var found = false;
  var targetDate = dateStr ? formatDateString(dateStr) : formatDateString(new Date());
  var cleanStudentId = String(studentId || '').trim().toLowerCase();

  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = formatDateString(data[i][0]);
    var rowStudentId = String(data[i][1] || '').trim().toLowerCase();
    if (rowStudentId === cleanStudentId && (!targetDate || rowDate === targetDate)) {
      attSheet.deleteRow(i + 1);
      found = true;
    }
  }
  SpreadsheetApp.flush();

  return {
    status: 'success',
    message: found ? 'ลบรายการเช็คชื่อเรียบร้อยแล้ว' : 'ไม่พบรายการเช็คชื่อ'
  };
}

function getAttendanceRecords(dateStr, className, activityName) {
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
  var cleanAct = String(activityName || '').trim();
  
  for (var i = 1; i < attData.length; i++) {
    var rowDate = formatDateString(attData[i][0]);
    var sId = String(attData[i][1]);
    var status = String(attData[i][2]);
    var rowAct = String(attData[i][8] || 'เช็คชื่อมาเรียน');
    var rowSes = String(attData[i][9] || '');
    var rowNote = String(attData[i][10] || '');
    
    if (rowDate === dateStr && studentIds[sId]) {
      if (isActivityMatch(cleanAct, rowAct)) {
        records[sId] = {
          status: status,
          checkinTime: String(attData[i][3] || '-'),
          gpsLocation: String(attData[i][4] || '-'),
          distanceMeters: attData[i][5] || 0,
          deviceId: String(attData[i][6] || ''),
          selfiePhoto: String(attData[i][7] || ''),
          checkinType: rowAct,
          activityName: rowAct,
          sessionId: rowSes,
          note: rowNote
        };
      }
    }
  }
  
  return records;
}

function saveAttendanceBatch(attendanceList, activityName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) {
    initDatabase();
    sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  }
  
  if (!attendanceList || attendanceList.length === 0) {
    return { status: 'error', message: 'ไม่มีข้อมูลบันทึก' };
  }
  
  var defaultAct = String(activityName || 'เช็คชื่อมาเรียน').trim();
  if (!defaultAct || defaultAct === 'all') defaultAct = 'เช็คชื่อมาเรียน';

  var existingData = sheet.getDataRange().getValues();
  var rowMap = {};
  
  for (var i = 1; i < existingData.length; i++) {
    var dStr = formatDateString(existingData[i][0]);
    var sId = String(existingData[i][1]);
    var rAct = String(existingData[i][8] || 'เช็คชื่อมาเรียน');
    rowMap[dStr + '_' + sId + '_' + rAct] = i + 1;
    if (!rowMap[dStr + '_' + sId]) {
      rowMap[dStr + '_' + sId] = i + 1;
    }
  }
  
  var nowTimeStr = new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
  
  for (var k = 0; k < attendanceList.length; k++) {
    var item = attendanceList[k];
    var formattedDate = formatDateString(item.date);
    var itemAct = String(item.activityName || defaultAct).trim();
    var key = formattedDate + '_' + item.studentId + '_' + itemAct;
    var noteVal = String(item.note || '').trim();
    
    if (rowMap[key]) {
      var rowIndex = rowMap[key];
      sheet.getRange(rowIndex, 3).setValue(item.status);
      sheet.getRange(rowIndex, 11).setValue(noteVal);
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
        itemAct,
        item.sessionId || '',
        noteVal
      ]);
    }
  }
  SpreadsheetApp.flush();
  
  return { status: 'success', message: 'บันทึกการเช็คชื่อ (' + defaultAct + ') เรียบร้อยแล้ว (' + attendanceList.length + ' คน)' };
}

function getDashboardData(targetDateStr, className, activityName) {
  var ss = getSpreadsheet();
  var students = getStudents();
  var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  
  var formattedTargetDate = targetDateStr ? formatDateString(targetDateStr) : formatDateString(new Date());
  var filterClass = String(className || 'all').trim();
  var filterAct = String(activityName || 'all').trim();
  
  var totalStudents = 0;
  var presentCount = 0;
  var absentCount = 0;
  var sickCount = 0;
  var leaveCount = 0;
  var lateCount = 0;
  
  var filteredStudents = [];
  for (var s = 0; s < students.length; s++) {
    if (isClassMatch(filterClass, students[s].className)) {
      filteredStudents.push(students[s]);
    }
  }
  totalStudents = filteredStudents.length;

  var todayRecords = {};
  var activityStatsMap = {};
  
  if (attSheet && attSheet.getLastRow() > 1) {
    var attData = attSheet.getDataRange().getValues();
    for (var i = 1; i < attData.length; i++) {
      var d = formatDateString(attData[i][0]);
      var sId = String(attData[i][1]);
      var st = String(attData[i][2]);
      var rowAct = String(attData[i][8] || 'เช็คชื่อมาเรียน').trim();
      
      if (d === formattedTargetDate) {
        if (!activityStatsMap[rowAct]) {
          activityStatsMap[rowAct] = { name: rowAct, total: 0, present: 0, absent: 0, sick: 0, leave: 0, late: 0 };
        }
        activityStatsMap[rowAct].total++;
        if (st === 'Present' || st === 'มาเรียน') activityStatsMap[rowAct].present++;
        else if (st === 'Absent' || st === 'ขาดเรียน') activityStatsMap[rowAct].absent++;
        else if (st === 'Sick' || st === 'ลาป่วย') activityStatsMap[rowAct].sick++;
        else if (st === 'Leave' || st === 'ลากิจ') activityStatsMap[rowAct].leave++;
        else if (st === 'Late' || st === 'มาสาย') activityStatsMap[rowAct].late++;

        if (filterAct === 'all' || rowAct === filterAct) {
          if (!todayRecords[sId] || st === 'Present' || st === 'มาเรียน') {
            todayRecords[sId] = st;
          }
        }
      }
    }
  }
  
  for (var k = 0; k < filteredStudents.length; k++) {
    var status = todayRecords[filteredStudents[k].studentId] || 'Unchecked';
    if (status === 'Present' || status === 'มาเรียน') presentCount++;
    else if (status === 'Absent' || status === 'ขาดเรียน') absentCount++;
    else if (status === 'Sick' || status === 'ลาป่วย') sickCount++;
    else if (status === 'Leave' || status === 'ลากิจ') leaveCount++;
    else if (status === 'Late' || status === 'มาสาย') lateCount++;
  }
  
  var classMap = {};
  for (var j = 0; j < filteredStudents.length; j++) {
    var cName = filteredStudents[j].className;
    if (!classMap[cName]) {
      classMap[cName] = { total: 0, present: 0, absent: 0, sick: 0, leave: 0, late: 0, unchecked: 0 };
    }
    classMap[cName].total++;
    var cStatus = todayRecords[filteredStudents[j].studentId] || 'Unchecked';
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
  classSummary.sort(function(a, b) {
    var diff = parseFloat(b.rate) - parseFloat(a.rate);
    if (diff !== 0) return diff;
    return a.className.localeCompare(b.className, 'th');
  });

  var activitySummary = [];
  for (var aKey in activityStatsMap) {
    var aItem = activityStatsMap[aKey];
    var aRate = aItem.total > 0 ? ((aItem.present / aItem.total) * 100).toFixed(1) : 0;
    activitySummary.push({
      name: aItem.name,
      total: aItem.total,
      present: aItem.present,
      absent: aItem.absent,
      sick: aItem.sick,
      leave: aItem.leave,
      late: aItem.late,
      rate: aRate
    });
  }
  activitySummary.sort(function(a, b) { return parseFloat(b.rate) - parseFloat(a.rate); });
  
  return {
    date: formattedTargetDate,
    filterClass: filterClass,
    filterActivity: filterAct,
    totalStudents: totalStudents,
    present: presentCount,
    absent: absentCount,
    sick: sickCount,
    leave: leaveCount,
    late: lateCount,
    unchecked: totalStudents - (presentCount + absentCount + sickCount + leaveCount + lateCount),
    classSummary: classSummary,
    activitySummary: activitySummary,
    activities: getActivityList().activities,
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
  
  var newId = (studentData.studentId && String(studentData.studentId).trim()) ? String(studentData.studentId).trim() : ('STD' + String(new Date().getTime()).slice(-6));
  
  // Check if duplicate studentId
  if (sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === newId.toLowerCase()) {
        return { status: 'error', message: 'รหัสนักเรียน "' + newId + '" มีอยู่ในระบบแล้ว' };
      }
    }
  }

  sheet.appendRow([newId, studentData.studentNo, studentData.name, studentData.className, '']);
  SpreadsheetApp.flush();
  clearScriptCache();
  return { status: 'success', message: 'เพิ่มนักเรียนเรียบร้อยแล้ว', studentId: newId };
}

function updateStudent(studentData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางนักเรียน' };
  
  var targetId = String(studentData.oldStudentId || studentData.studentId).trim();
  var newStudentId = String(studentData.studentId || targetId).trim();
  var data = sheet.getDataRange().getValues();
  
  // If studentId is changed, check if new studentId already exists on another row
  if (targetId.toLowerCase() !== newStudentId.toLowerCase()) {
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][0]).trim().toLowerCase() === newStudentId.toLowerCase()) {
        return { status: 'error', message: 'รหัสนักเรียน "' + newStudentId + '" มีอยู่ในระบบแล้ว' };
      }
    }
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === targetId.toLowerCase()) {
      var row = i + 1;
      sheet.getRange(row, 1).setValue(newStudentId);
      sheet.getRange(row, 2).setValue(studentData.studentNo);
      sheet.getRange(row, 3).setValue(studentData.name);
      sheet.getRange(row, 4).setValue(studentData.className);
      
      // If student ID changed, also update studentId in Attendance records
      if (targetId.toLowerCase() !== newStudentId.toLowerCase()) {
        var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
        if (attSheet && attSheet.getLastRow() > 1) {
          var attData = attSheet.getDataRange().getValues();
          for (var a = 1; a < attData.length; a++) {
            if (String(attData[a][1]).trim().toLowerCase() === targetId.toLowerCase()) {
              attSheet.getRange(a + 1, 2).setValue(newStudentId);
            }
          }
        }
      }

      SpreadsheetApp.flush();
      clearScriptCache();
      return { status: 'success', message: 'แก้ไขข้อมูลนักเรียนเรียบร้อยแล้ว' };
    }
  }
  return { status: 'error', message: 'ไม่พบนักเรียนในระบบ' };
}

function deleteStudent(studentId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางนักเรียน' };
  
  if (typeof studentId === 'object' && studentId !== null) {
    studentId = studentId.studentId || studentId.id || (Array.isArray(studentId) ? studentId[0] : '');
  }
  studentId = String(studentId || '').trim().toLowerCase();
  if (!studentId) return { status: 'error', message: 'กรุณาระบุรหัสนักเรียน' };

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { status: 'error', message: 'ไม่พบนักเรียนในระบบ' };

  var data = sheet.getRange(1, 1, lastRow, 5).getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === studentId) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      clearScriptCache();
      return { status: 'success', message: 'ลบข้อมูลนักเรียนเรียบร้อยแล้ว' };
    }
  }
  return { status: 'error', message: 'ไม่พบนักเรียนในระบบ' };
}

/**
 * Deletes multiple students in a single high-performance batch operation
 */
function deleteStudentsBatch(studentIds) {
  if (!studentIds) {
    return { status: 'error', message: 'กรุณาเลือกนักเรียนที่ต้องการลบ' };
  }

  if (typeof studentIds === 'string') {
    try {
      studentIds = JSON.parse(studentIds);
    } catch (e) {
      studentIds = [studentIds];
    }
  }

  if (Array.isArray(studentIds) && studentIds.length === 1 && Array.isArray(studentIds[0])) {
    studentIds = studentIds[0];
  }

  if (!Array.isArray(studentIds) || studentIds.length === 0) {
    return { status: 'error', message: 'กรุณาเลือกนักเรียนที่ต้องการลบ' };
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางนักเรียน' };

  var targetMap = {};
  for (var k = 0; k < studentIds.length; k++) {
    var rawItem = studentIds[k];
    var sid = (typeof rawItem === 'object' && rawItem) ? (rawItem.studentId || rawItem.id || '') : String(rawItem || '');
    sid = String(sid).trim().toLowerCase();
    if (sid) targetMap[sid] = true;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { status: 'success', message: 'ลบข้อมูลนักเรียนเรียบร้อยแล้ว', deletedCount: 0 };
  }

  var data = sheet.getRange(1, 1, lastRow, 5).getValues();
  var remainingRows = [data[0]]; // Preserve header row
  var deletedCount = 0;

  for (var i = 1; i < data.length; i++) {
    var studentId = String(data[i][0]).trim().toLowerCase();
    if (targetMap[studentId]) {
      deletedCount++;
    } else {
      remainingRows.push(data[i]);
    }
  }

  sheet.clear();
  if (remainingRows.length > 0) {
    sheet.getRange(1, 1, remainingRows.length, 5).setValues(remainingRows);
  }
  SpreadsheetApp.flush();
  clearScriptCache();

  return {
    status: 'success',
    message: 'ลบข้อมูลนักเรียนที่เลือกเรียบร้อยแล้ว (' + deletedCount + ' คน)',
    deletedCount: deletedCount
  };
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
  SpreadsheetApp.flush();
  clearScriptCache();
  
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

/**
 * =========================================================================
 * LINE GROUP NOTIFICATION & AUTOMATED SUMMARY SYSTEM
 * =========================================================================
 */


/**
 * Handles incoming Webhook events from LINE Messaging API
 * Automatically captures Group ID (e.g. Parents Group) when the bot is invited or messaged
 */
function handleLineBotWebhook(events) {
  try {
    var props = PropertiesService.getScriptProperties();
    var config = getLineConfig();
    var token = config.lineToken || '';
    var rawGroups = props.getProperty('KNOWN_LINE_GROUPS');
    var knownGroups = rawGroups ? JSON.parse(rawGroups) : [];

    for (var i = 0; i < events.length; i++) {
      var ev = events[i];
      var gId = (ev.source && (ev.source.groupId || ev.source.roomId || ev.source.userId)) || '';
      var sType = (ev.source && ev.source.type) || 'group';

      if (gId) {
        props.setProperty('LAST_DETECTED_GROUP_ID', gId);

        var existingIdx = -1;
        for (var k = 0; k < knownGroups.length; k++) {
          if (knownGroups[k].id === gId) { existingIdx = k; break; }
        }
        var groupObj = {
          id: gId,
          type: sType,
          joinedAt: new Date().toISOString(),
          label: (sType === 'group' ? 'กลุ่ม LINE ' : 'แชท ') + gId.substring(0, 8) + '...'
        };
        if (existingIdx > -1) {
          knownGroups[existingIdx] = groupObj;
        } else {
          knownGroups.unshift(groupObj);
        }

        props.setProperty('KNOWN_LINE_GROUPS', JSON.stringify(knownGroups.slice(0, 15)));

        // Send friendly greeting & group ID confirmation if join or message
        if (token && ev.replyToken && (ev.type === 'join' || (ev.type === 'message' && ev.message && ev.message.text))) {
          var userText = (ev.message && ev.message.text ? String(ev.message.text).trim() : '');
          if (ev.type === 'join' || userText === '!id' || userText === 'กลุ่ม' || userText === 'เช็คชื่อ' || userText.indexOf('รหัสกลุ่ม') > -1) {
            var replyText = '🎉 สวัสดีครับผู้ปกครองและอาจารย์ทุกท่าน!\n' +
                            '🤖 บอทระบบเช็คชื่อนักศึกษาได้เชื่อมต่อกับกลุ่มนี้เรียบร้อยแล้ว\n' +
                            '📌 รหัสกลุ่มของคุณคือ: ' + gId + '\n' +
                            '✨ ระบบจะส่งรายงานสรุปการมาเรียนเข้ากลุ่มนี้ตามเวลาที่กำหนดครับ';
            
            var replyPayload = {
              replyToken: ev.replyToken,
              messages: [{ type: 'text', text: replyText }]
            };

            UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
              method: 'post',
              contentType: 'application/json',
              headers: { 'Authorization': 'Bearer ' + token },
              payload: JSON.stringify(replyPayload),
              muteHttpExceptions: true
            });
          }
        }
      }
    }
  } catch (err) {
    Logger.log('Error in handleLineBotWebhook: ' + err);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getDetectedLineGroups() {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('KNOWN_LINE_GROUPS');
    var lastId = props.getProperty('LAST_DETECTED_GROUP_ID') || '';
    var groups = raw ? JSON.parse(raw) : [];
    return {
      status: 'success',
      lastGroupId: lastId,
      groups: groups
    };
  } catch (e) {
    return { status: 'error', message: e.message, groups: [], lastGroupId: '' };
  }
}

function getLineConfig() {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('LINE_CONFIG');
    if (raw) {
      return JSON.parse(raw);
    }
  } catch (e) {
    Logger.log('Error in getLineConfig: ' + e);
  }
  return {
    lineToken: '',
    lineGroupId: '',
    webhookUrl: '',
    targetClass: 'all',
    targetActivity: 'all',
    statusFilter: 'all',
    messageStyle: 'formal',
    autoSendEnabled: false,
    autoSendHour: 16
  };
}

function saveLineConfig(config) {
  config = config || {};
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('LINE_CONFIG', JSON.stringify(config));
    
    // Manage automated daily trigger
    setupDailyLineTrigger(config.autoSendEnabled === true, parseInt(config.autoSendHour, 10) || 16);
    
    return {
      status: 'success',
      message: 'บันทึกการตั้งค่า LINE และกำหนดเวลาส่งอัตโนมัติเรียบร้อยแล้ว'
    };
  } catch (e) {
    return {
      status: 'error',
      message: 'เกิดข้อผิดพลาดในการบันทึกการตั้งค่า: ' + e.message
    };
  }
}

function setupDailyLineTrigger(enable, targetHour) {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'sendDailyAttendanceLineReport') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    if (enable) {
      if (targetHour < 0 || targetHour > 23) targetHour = 16;
      ScriptApp.newTrigger('sendDailyAttendanceLineReport')
        .timeBased()
        .everyDays(1)
        .atHour(targetHour)
        .create();
    }
  } catch (e) {
    Logger.log('Trigger Setup Error: ' + e);
  }
}

function testLineConnection(token, webhookUrl, lineGroupId, customMsg) {
  token = String(token || '').trim();
  webhookUrl = String(webhookUrl || '').trim();
  lineGroupId = String(lineGroupId || '').trim();
  customMsg = String(customMsg || '').trim();
  
  if (!token && !webhookUrl) {
    return { status: 'error', message: 'กรุณาระบุ LINE Channel Access Token หรือ Webhook URL' };
  }
  
  var now = new Date();
  var timeStr = (now.getHours() < 10 ? '0' + now.getHours() : now.getHours()) + ':' + 
                (now.getMinutes() < 10 ? '0' + now.getMinutes() : now.getMinutes()) + ' น.';
  var testMsg = customMsg || ('🔔 [ทดสอบระบบ] การเชื่อมต่อ LINE สำเร็จ!\n' +
                '📅 วันที่: ' + formatDateString(now) + ' (' + timeStr + ')\n' +
                '✨ ระบบเช็คชื่อนักศึกษาพร้อมส่งรายงานสรุปเข้า LINE แล้วครับ');

  return dispatchLineNotification(token, testMsg, lineGroupId, webhookUrl);
}

function dispatchLineNotification(token, msg, lineGroupId, webhookUrl) {
  token = String(token || '').trim();
  lineGroupId = String(lineGroupId || '').trim();
  webhookUrl = String(webhookUrl || '').trim();

  if (!token && !webhookUrl) {
    return { status: 'error', message: 'กรุณาระบุ LINE Token หรือ Webhook URL' };
  }

  // 1. LINE Messaging API (Channel Access Token - Long token > 60 chars)
  if (token && token.length > 60) {
    var endpoint = 'https://api.line.me/v2/bot/message/broadcast';
    var payloadObj = {
      messages: [{ type: 'text', text: msg }]
    };

    if (lineGroupId && (lineGroupId.startsWith('C') || lineGroupId.startsWith('U') || lineGroupId.startsWith('R'))) {
      endpoint = 'https://api.line.me/v2/bot/message/push';
      payloadObj.to = lineGroupId;
    }

    var msgApiOptions = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      payload: JSON.stringify(payloadObj),
      muteHttpExceptions: true
    };

    try {
      var msgApiRes = UrlFetchApp.fetch(endpoint, msgApiOptions);
      var msgApiCode = msgApiRes.getResponseCode();
      if (msgApiCode === 200) {
        return { status: 'success', message: 'ส่งข้อความผ่าน LINE Messaging API (LINE OA) สำเร็จ!' };
      } else {
        var errText = msgApiRes.getContentText();
        // Fallback to broadcast if push to specific ID failed
        if (endpoint.indexOf('/push') > -1) {
          var bcOptions = {
            method: 'post',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + token },
            payload: JSON.stringify({ messages: [{ type: 'text', text: msg }] }),
            muteHttpExceptions: true
          };
          var bcRes = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', bcOptions);
          if (bcRes.getResponseCode() === 200) {
            return { status: 'success', message: 'ส่งข้อความแบบ Broadcast ผ่าน LINE OA สำเร็จ!' };
          }
        }
        return { status: 'error', message: 'LINE Messaging API ตอบกลับข้อผิดพลาด (HTTP ' + msgApiCode + '): ' + errText };
      }
    } catch (e) {
      return { status: 'error', message: 'เกิดข้อผิดพลาดในการส่ง LINE Messaging API: ' + e.message };
    }
  }

  // 2. LINE Notify API (Legacy Token <= 60 chars)
  if (token) {
    try {
      var notifyOptions = {
        method: 'post',
        headers: {
          'Authorization': 'Bearer ' + token
        },
        payload: {
          'message': '\n' + msg
        },
        muteHttpExceptions: true
      };
      var notifyRes = UrlFetchApp.fetch('https://notify-api.line.me/api/notify', notifyOptions);
      var notifyCode = notifyRes.getResponseCode();
      if (notifyCode === 200) {
        return { status: 'success', message: 'ส่งข้อความผ่าน LINE Notify สำเร็จ!' };
      } else {
        return { status: 'error', message: 'LINE Notify ตอบกลับข้อผิดพลาด (HTTP ' + notifyCode + '): ' + notifyRes.getContentText() };
      }
    } catch (e) {
      return { status: 'error', message: 'เกิดข้อผิดพลาดในการส่ง LINE Notify: ' + e.message };
    }
  }

  // 3. Webhook URL
  if (webhookUrl) {
    try {
      var whOptions = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          content: msg,
          message: msg,
          text: msg
        }),
        muteHttpExceptions: true
      };
      var whRes = UrlFetchApp.fetch(webhookUrl, whOptions);
      var whCode = whRes.getResponseCode();
      if (whCode >= 200 && whCode < 300) {
        return { status: 'success', message: 'ส่งข้อความผ่าน Webhook สำเร็จ (HTTP ' + whCode + ')' };
      } else {
        return { status: 'error', message: 'Webhook ตอบกลับข้อผิดพลาด (HTTP ' + whCode + '): ' + whRes.getContentText() };
      }
    } catch (e) {
      return { status: 'error', message: 'เกิดข้อผิดพลาดในการส่ง Webhook: ' + e.message };
    }
  }

  return { status: 'error', message: 'ไม่สามารถส่งข้อความได้' };
}

function sendLineReport(options) {
  options = options || {};
  var config = getLineConfig();
  
  var token = String(options.token || config.lineToken || '').trim();
  var lineGroupId = String(options.lineGroupId || config.lineGroupId || '').trim();
  var webhookUrl = String(options.webhookUrl || config.webhookUrl || '').trim();
  
  if (!token && !webhookUrl) {
    return { status: 'error', message: 'ยังไม่ได้ตั้งค่า LINE Token หรือ Webhook URL' };
  }

  // If customMessage is provided directly, dispatch it immediately
  if (options.customMessage && String(options.customMessage).trim().length > 0) {
    var customText = String(options.customMessage).trim();
    return dispatchLineNotification(token, customText, lineGroupId, webhookUrl);
  }
  
  var targetDate = options.date || formatDateString(new Date());
  var targetClass = options.className || config.targetClass || 'all';
  var targetActivity = options.activityName || config.targetActivity || 'all';
  var statusFilter = options.statusFilter || config.statusFilter || 'all';
  var messageStyle = options.messageStyle || config.messageStyle || 'formal';
  
  var recordsMap = getAttendanceRecords(targetDate, targetClass, targetActivity);
  var students = getStudents(targetClass);
  
  var countPresent = 0;
  var countAbsent = 0;
  var countSick = 0;
  var countLeave = 0;
  var countLate = 0;
  var absentList = [];
  var sickList = [];
  var leaveList = [];
  var lateList = [];
  var presentList = [];
  
  for (var i = 0; i < students.length; i++) {
    var s = students[i];
    var sId = String(s.studentId).trim();
    var cleanId = sId.toLowerCase();
    var rec = recordsMap[sId] || recordsMap[cleanId];
    var st = 'Absent';
    var note = '';
    var cTime = '-';
    
    if (rec) {
      if (typeof rec === 'object') {
        st = rec.status || 'Present';
        note = rec.note || '';
        cTime = rec.checkinTime || '-';
      } else {
        st = String(rec);
      }
    }
    
    var sItem = {
      no: s.studentNo || (i + 1),
      name: s.name,
      studentId: s.studentId,
      className: s.className,
      status: st,
      note: note,
      time: cTime
    };
    
    if (st === 'Present' || st === 'มาเรียน') {
      countPresent++;
      presentList.push(sItem);
    } else if (st === 'Absent' || st === 'ขาด' || st === 'ขาดเรียน') {
      countAbsent++;
      absentList.push(sItem);
    } else if (st === 'Sick' || st === 'ลาป่วย') {
      countSick++;
      sickList.push(sItem);
    } else if (st === 'Leave' || st === 'ลากิจ') {
      countLeave++;
      leaveList.push(sItem);
    } else if (st === 'Late' || st === 'มาสาย') {
      countLate++;
      lateList.push(sItem);
    }
  }
  
  var totalStudents = students.length;
  var attRate = totalStudents > 0 ? (((countPresent + countLate) / totalStudents) * 100).toFixed(1) : '100.0';
  
  // Format message text
  var msg = '';
  var isMultiClass = (targetClass === 'all' || targetClass === 'ทุกห้องเรียน' || String(targetClass).indexOf(',') > -1);
  var classLabel = (targetClass === 'all' || targetClass === 'ทุกห้องเรียน') ? 'ทุกห้องเรียน' : ('ห้อง ' + targetClass);
  var actLabel = (targetActivity === 'all' || !targetActivity) ? 'เช็คชื่อประจำวัน' : targetActivity;
  
  if (messageStyle === 'parent' || messageStyle === 'parents') {
    // 👨‍👩‍👧‍👦 Parent Alert Template (เน้นความชัดเจนและสุภาพสำหรับกลุ่มผู้ปกครอง)
    msg += '📢 [รายงานการมาเรียนถึงผู้ปกครอง]\n';
    msg += '🏫 ' + classLabel + '\n';
    msg += '📅 วันที่: ' + targetDate + ' (' + actLabel + ')\n';
    msg += '------------------------\n';
    msg += '🟢 นักเรียนที่มาเรียน: ' + countPresent + ' คน (' + attRate + '%)\n';
    msg += '🔴 นักเรียนที่ขาดเรียน: ' + countAbsent + ' คน\n';
    msg += '🟡 ลาป่วย/ลากิจ: ' + (countSick + countLeave) + ' คน\n';
    msg += '🌸 มาสาย: ' + countLate + ' คน\n';
    msg += '👥 จำนวนนักเรียนทั้งหมด: ' + totalStudents + ' คน\n';
    msg += '------------------------\n';

    if (absentList.length > 0) {
      msg += '\n⚠️ รายชื่อนักเรียนที่ไม่พบการเช็คชื่อวันนี้ (' + absentList.length + ' คน):\n';
      for (var pAb = 0; pAb < absentList.length; pAb++) {
        var rTagAb = isMultiClass ? ' (' + absentList[pAb].className + ')' : '';
        msg += '• เลขที่ ' + absentList[pAb].no + ' ' + absentList[pAb].name + rTagAb + (absentList[pAb].note ? ' [' + absentList[pAb].note + ']' : '') + '\n';
      }
    }

    if (lateList.length > 0) {
      msg += '\n⏰ นักเรียนที่มาสาย (' + lateList.length + ' คน):\n';
      for (var pLt = 0; pLt < lateList.length; pLt++) {
        var rTagLt = isMultiClass ? ' (' + lateList[pLt].className + ')' : '';
        msg += '• เลขที่ ' + lateList[pLt].no + ' ' + lateList[pLt].name + rTagLt + ' [เข้าเรียนเวลา ' + lateList[pLt].time + ']\n';
      }
    }

    if (sickList.length > 0 || leaveList.length > 0) {
      msg += '\n✉️ นักเรียนที่แจ้งลา (' + (sickList.length + leaveList.length) + ' คน):\n';
      for (var pSk = 0; pSk < sickList.length; pSk++) {
        var rTagSk = isMultiClass ? ' (' + sickList[pSk].className + ')' : '';
        msg += '• [ป่วย] เลขที่ ' + sickList[pSk].no + ' ' + sickList[pSk].name + rTagSk + (sickList[pSk].note ? ' (' + sickList[pSk].note + ')' : '') + '\n';
      }
      for (var pLv = 0; pLv < leaveList.length; pLv++) {
        var rTagLv = isMultiClass ? ' (' + leaveList[pLv].className + ')' : '';
        msg += '• [กิจ] เลขที่ ' + leaveList[pLv].no + ' ' + leaveList[pLv].name + rTagLv + (leaveList[pLv].note ? ' (' + leaveList[pLv].note + ')' : '') + '\n';
      }
    }

    msg += '\n🙏 หากผู้ปกครองท่านใดมีข้อสงสัย หรือประสงค์แจ้งการลาของบุตรหลาน สามารถติดต่อครูประจำชั้นได้เลยครับ ขอบคุณครับ';
  } else if (messageStyle === 'compact') {
    // 📢 Compact Alert
    msg += '📢 [สรุปการเช็คชื่อ] ' + classLabel + '\n';
    msg += '📅 วันที่: ' + targetDate + '\n';
    msg += '🏷️ รายการ: ' + actLabel + '\n';
    msg += '------------------------\n';
    msg += '🟢 มาเรียน: ' + countPresent + ' คน (' + attRate + '%)\n';
    msg += '🔴 ขาดเรียน: ' + countAbsent + ' คน\n';
    msg += '🟡 ลาป่วย: ' + countSick + ' คน | 🟣 ลากิจ: ' + countLeave + ' คน\n';
    msg += '🌸 มาสาย: ' + countLate + ' คน\n';
    msg += '👥 รวมทั้งหมด: ' + totalStudents + ' คน\n';
    if (absentList.length > 0) {
      msg += '\n❌ รายชื่อนักเรียนขาดเรียน (' + absentList.length + ' คน):\n';
      for (var a = 0; a < absentList.length; a++) {
        var rTag = isMultiClass ? ' (' + absentList[a].className + ')' : '';
        msg += '• เลขที่ ' + absentList[a].no + ' ' + absentList[a].name + rTag + (absentList[a].note ? ' [' + absentList[a].note + ']' : '') + '\n';
      }
    }
  } else if (messageStyle === 'detailed') {
    // 📋 Full Detailed List
    msg += '📋 รายงานการเช็คชื่อนักเรียนฉบับละเอียด\n';
    msg += '🏫 ' + classLabel + ' • ' + actLabel + '\n';
    msg += '📅 วันที่: ' + targetDate + '\n';
    msg += '------------------------\n';
    msg += '📊 สถิติ: มา ' + countPresent + ' | ขาด ' + countAbsent + ' | ป่วย ' + countSick + ' | กิจ ' + countLeave + ' | สาย ' + countLate + ' (รวม ' + totalStudents + ' คน, ' + attRate + '%)\n';
    msg += '------------------------\n';
    for (var d = 0; d < students.length; d++) {
      var curS = students[d];
      var curRec = recordsMap[curS.studentId] || recordsMap[String(curS.studentId).toLowerCase()];
      var curSt = 'ขาดเรียน';
      var curStIcon = '🔴';
      var curNote = '';
      var curT = '';
      if (curRec && typeof curRec === 'object') {
        if (curRec.status === 'Present' || curRec.status === 'มาเรียน') { curSt = 'มา'; curStIcon = '🟢'; }
        else if (curRec.status === 'Sick' || curRec.status === 'ลาป่วย') { curSt = 'ลาป่วย'; curStIcon = '🟡'; }
        else if (curRec.status === 'Leave' || curRec.status === 'ลากิจ') { curSt = 'ลากิจ'; curStIcon = '🟣'; }
        else if (curRec.status === 'Late' || curRec.status === 'มาสาย') { curSt = 'มาสาย'; curStIcon = '🌸'; }
        if (curRec.checkinTime && curRec.checkinTime !== '-') curT = ' [' + curRec.checkinTime + ']';
        if (curRec.note) curNote = ' (' + curRec.note + ')';
      }
      var rTagD = isMultiClass ? ' (' + curS.className + ')' : '';
      msg += (d + 1) + '. ' + curStIcon + ' ' + curS.name + rTagD + ' - ' + curSt + curT + curNote + '\n';
    }
  } else {
    // 📊 Formal Summary (Default)
    msg += '📊 รายงานสรุปการเช็คชื่อนักเรียน\n';
    msg += '🏫 ห้องเรียน: ' + classLabel + '\n';
    msg += '🏷️ กิจกรรม/วิชา: ' + actLabel + '\n';
    msg += '📅 วันที่: ' + targetDate + '\n';
    msg += '========================\n';
    msg += '🟢 มาเรียน: ' + countPresent + ' คน\n';
    msg += '🔴 ขาดเรียน: ' + countAbsent + ' คน\n';
    msg += '🟡 ลาป่วย: ' + countSick + ' คน\n';
    msg += '🟣 ลากิจ: ' + countLeave + ' คน\n';
    msg += '🌸 มาสาย: ' + countLate + ' คน\n';
    msg += '👥 รวมทั้งหมด: ' + totalStudents + ' คน\n';
    msg += '📈 อัตราการมาเรียน: ' + attRate + '%\n';
    msg += '========================\n';
    
    if (absentList.length > 0) {
      msg += '\n❌ ขาดเรียน (' + absentList.length + ' คน):\n';
      for (var ab = 0; ab < absentList.length; ab++) {
        var rTagAb = isMultiClass ? ' (' + absentList[ab].className + ')' : '';
        msg += '• ' + absentList[ab].no + '. ' + absentList[ab].name + rTagAb + (absentList[ab].note ? ' [' + absentList[ab].note + ']' : '') + '\n';
      }
    }
    
    if (sickList.length > 0 || leaveList.length > 0) {
      msg += '\n✉️ ลาป่วย/ลากิจ (' + (sickList.length + leaveList.length) + ' คน):\n';
      for (var sk = 0; sk < sickList.length; sk++) {
        var rTagSk = isMultiClass ? ' (' + sickList[sk].className + ')' : '';
        msg += '• [ป่วย] ' + sickList[sk].no + '. ' + sickList[sk].name + rTagSk + (sickList[sk].note ? ' [' + sickList[sk].note + ']' : '') + '\n';
      }
      for (var lv = 0; lv < leaveList.length; lv++) {
        var rTagLv = isMultiClass ? ' (' + leaveList[lv].className + ')' : '';
        msg += '• [กิจ] ' + leaveList[lv].no + '. ' + leaveList[lv].name + rTagLv + (leaveList[lv].note ? ' [' + leaveList[lv].note + ']' : '') + '\n';
      }
    }

    if (lateList.length > 0) {
      msg += '\n⏰ มาสาย (' + lateList.length + ' คน):\n';
      for (var lt = 0; lt < lateList.length; lt++) {
        var rTagLt = isMultiClass ? ' (' + lateList[lt].className + ')' : '';
        msg += '• ' + lateList[lt].no + '. ' + lateList[lt].name + rTagLt + ' [' + lateList[lt].time + ']\n';
      }
    }
  }

  return dispatchLineNotification(token, msg, lineGroupId, webhookUrl);
}

function sendDailyAttendanceLineReport() {
  var config = getLineConfig();
  if (!config || !config.autoSendEnabled) return;
  var today = formatDateString(new Date());
  sendLineReport({
    date: today,
    className: config.targetClass || 'all',
    activityName: config.targetActivity || 'all',
    statusFilter: config.statusFilter || 'all',
    messageStyle: config.messageStyle || 'formal'
  });
}
