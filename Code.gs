function formatTimeString(val) {
  if (!val || val === '-' || val === 'null' || val === 'undefined') return '-';
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, 'Asia/Bangkok', 'HH:mm');
    } catch (e) {
      var h = String(val.getHours());
      if (h.length < 2) h = '0' + h;
      var m = String(val.getMinutes());
      if (m.length < 2) m = '0' + m;
      return h + ':' + m;
    }
  }
  var str = String(val).trim();
  if (!str || str === '-' || str === 'undefined' || str === 'null') return '-';

  // If long date string like "Sat Dec 30 1899 09:01:00..." or ISO string
  var timeMatch = str.match(/(?:(?:[A-Za-z]{3}\s+[A-Za-z]{3}\s+\d+\s+\d+\s+)|T|^)(\d{1,2}):(\d{2})(?::\d{2})?/);
  if (timeMatch) {
    var hh = timeMatch[1].length === 1 ? '0' + timeMatch[1] : timeMatch[1];
    var mm = timeMatch[2];
    return hh + ':' + mm;
  }
  var generalMatch = str.match(/(\d{1,2}):(\d{2})/);
  if (generalMatch) {
    var hh2 = generalMatch[1].length === 1 ? '0' + generalMatch[1] : generalMatch[1];
    var mm2 = generalMatch[2];
    return hh2 + ':' + mm2;
  }
  return str;
}

function formatThaiDate(dateVal) {
  if (!dateVal) return '';
  var d = new Date();
  if (typeof dateVal === 'object' && dateVal instanceof Date) {
    d = dateVal;
  } else {
    var str = String(dateVal).trim();
    if (str.indexOf('/') > -1) {
      var parts = str.split('/');
      if (parts.length === 3) {
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1;
        var year = parseInt(parts[2], 10);
        if (year > 2500) year -= 543;
        d = new Date(year, month, day);
      }
    } else if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
      var yParts = str.substring(0, 10).split('-');
      d = new Date(parseInt(yParts[0], 10), parseInt(yParts[1], 10) - 1, parseInt(yParts[2], 10));
    }
  }
  var months = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];
  return d.getDate() + ' ' + months[d.getMonth()] + ' ' + (d.getFullYear() + 543);
}

function getThaiDayOfWeek(dateStr) {
  var d = new Date();
  if (dateStr) {
    var str = String(dateStr).trim();
    if (str.indexOf('/') > -1) {
      var parts = str.split('/');
      if (parts.length === 3) {
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1;
        var year = parseInt(parts[2], 10);
        if (year > 2500) year -= 543;
        d = new Date(year, month, day);
      }
    } else if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
      var yParts = str.substring(0, 10).split('-');
      d = new Date(parseInt(yParts[0], 10), parseInt(yParts[1], 10) - 1, parseInt(yParts[2], 10));
    }
  }
  var days = ['วันอาทิตย์', 'วันจันทร์', 'วันอังคาร', 'วันพุธ', 'วันพฤหัสบดี', 'วันศุกร์', 'วันเสาร์'];
  return days[d.getDay()] || 'สรุปการเช็คชื่อ';
}

/**
 * Convert any student selfie (Base64, Google Drive link, or direct URL)
 * into a direct public Google CDN HTTPS URL (https://lh3.googleusercontent.com/d/FILE_ID)
 * that LINE Messaging API can load and render 100% reliably in Flex Messages.
 */
function convertSelfieToPublicUrl(selfieData, studentId, dateStr) {
  if (!selfieData) return '';
  var str = String(selfieData).trim();
  if (!str || str === '-' || str === 'undefined' || str === 'null') return '';
  
  // 1. If already a direct public Google CDN or image URL
  if (str.indexOf('https://lh3.googleusercontent.com/d/') === 0 || 
      str.indexOf('https://images.unsplash.com/') === 0 || 
      str.indexOf('https://i.imgur.com/') === 0) {
    return str;
  }
  
  // 2. If Google Drive URL, extract File ID and convert to direct public Google CDN
  var driveMatch = str.match(/(?:id=|\/d\/)([a-zA-Z0-9_-]{20,})/);
  if (driveMatch && driveMatch[1]) {
    var fileId = driveMatch[1];
    try {
      if (typeof DriveApp !== 'undefined') {
        var driveFile = DriveApp.getFileById(fileId);
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    } catch (e) {}
    return 'https://lh3.googleusercontent.com/d/' + fileId;
  }
  
  // 3. If Base64 image (data:image/...;base64,... or raw base64 string)
  if (str.indexOf('data:image/') === 0 || str.indexOf('base64,') > -1 || str.length > 500) {
    try {
      if (typeof Utilities !== 'undefined' && typeof DriveApp !== 'undefined') {
        var cleanBase64 = str;
        var mimeType = 'image/jpeg';
        var ext = 'jpg';
        
        if (str.indexOf('data:') === 0) {
          var parts = str.split(';base64,');
          mimeType = parts[0].replace('data:', '');
          cleanBase64 = parts[1] || '';
          if (mimeType.indexOf('png') > -1) ext = 'png';
        }
        
        var decodedBytes = Utilities.base64Decode(cleanBase64);
        var fileName = 'selfie_' + String(studentId || 'std').replace(/[^a-zA-Z0-9_-]/g, '_') + '_' + String(dateStr || 'today').replace(/[^0-9-]/g, '') + '_' + new Date().getTime() + '.' + ext;
        var blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
        
        // Find or create "Student_Attendance_Selfies" folder
        var folderName = 'Student_Attendance_Selfies';
        var folders = DriveApp.getFoldersByName(folderName);
        var folder;
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = DriveApp.createFolder(folderName);
        }
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        var savedFile = folder.createFile(blob);
        savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var savedFileId = savedFile.getId();
        
        return 'https://lh3.googleusercontent.com/d/' + savedFileId;
      }
    } catch (e) {
      Logger.log('Error converting base64 to Drive public URL: ' + e.message);
    }
  }
  
  // 4. If direct HTTPS image URL
  if (str.indexOf('https://') === 0 && (str.indexOf('.jpg') > -1 || str.indexOf('.png') > -1 || str.indexOf('.jpeg') > -1 || str.indexOf('.webp') > -1)) {
    return str;
  }
  
  return '';
}

function buildLineFlexAttendanceMessage(data) {
  var targetDate = data.targetDate || 'วันนี้';
  var rawDate = data.rawDate || '';
  var dayName = getThaiDayOfWeek(rawDate || targetDate);
  var classLabel = data.classLabel || 'ทุกห้องเรียน';
  var actLabel = data.actLabel || 'เช็คชื่อมาเรียน';
  
  var totalStudents = data.totalStudents || 0;
  var countPresent = data.countPresent || 0;
  var countAbsent = data.countAbsent || 0;
  var countSick = data.countSick || 0;
  var countLeave = data.countLeave || 0;
  var countLate = data.countLate || 0;
  var attRate = data.attRate || '100.0';
  
  var presentList = data.presentList || [];
  var absentList = data.absentList || [];
  var lateList = data.lateList || [];
  var sickList = data.sickList || [];
  var leaveList = data.leaveList || [];
  var isMultiClass = data.isMultiClass || false;

  var bodyContents = [];

  // 1. Overall Status Card / Row
  bodyContents.push({
    "type": "box",
    "layout": "horizontal",
    "paddingBottom": "10px",
    "contents": [
      {
        "type": "box",
        "layout": "vertical",
        "width": "85px",
        "contents": [
          {
            "type": "text",
            "text": attRate + "%",
            "weight": "bold",
            "size": "lg",
            "color": "#10B981"
          },
          {
            "type": "text",
            "text": "เข้าเรียน",
            "size": "xxs",
            "color": "#64748B",
            "margin": "xs"
          }
        ]
      },
      {
        "type": "box",
        "layout": "vertical",
        "flex": 1,
        "contents": [
          {
            "type": "text",
            "text": "มาเรียน " + (countPresent + countLate) + " จาก " + totalStudents + " คน",
            "weight": "bold",
            "size": "sm",
            "color": "#1E293B",
            "wrap": true
          },
          {
            "type": "text",
            "text": "ตรงเวลา " + countPresent + " • สาย " + countLate + " • ลา " + (countSick + countLeave) + " • ขาด " + countAbsent,
            "size": "xs",
            "color": "#64748B",
            "margin": "xs",
            "wrap": true
          }
        ]
      }
    ]
  });

  // 1.5 Top 3 Early Birds Champions 3D Podium Box (👑 Gold / Silver / Bronze)
  var earlyBirds = [];
  for (var eb = 0; eb < presentList.length; eb++) {
    var pTime = formatTimeString(presentList[eb].time);
    if (pTime && pTime !== '-') {
      earlyBirds.push({
        no: presentList[eb].no,
        name: presentList[eb].name,
        className: presentList[eb].className,
        time: pTime,
        selfiePhoto: presentList[eb].selfiePhoto || ''
      });
    }
  }
  earlyBirds.sort(function(a, b) {
    return String(a.time).localeCompare(String(b.time));
  });

  if (earlyBirds.length > 0) {
    var top1 = earlyBirds[0];
    var top2 = earlyBirds[1] || null;
    var top3 = earlyBirds[2] || null;

    var championColumns = [];

    // Helper to build a column box in LINE Flex (Strictly compliant with LINE Flex JSON)
    function buildFlexChampCol(item, rankNum, medalEmoji, rankLabel, crownColor, badgeBgHex, isChamp) {
      var colContents = [];
      
      // Crown & Rank Header
      colContents.push({
        "type": "text",
        "text": "👑 " + rankNum,
        "size": isChamp ? "xs" : "xxs",
        "color": crownColor,
        "weight": "bold",
        "align": "center"
      });

      // Photo (Convert Base64/Drive to direct public HTTPS CDN URL) or Stylized Avatar Box
      var directPhotoUrl = convertSelfieToPublicUrl(item.selfiePhoto, item.no || item.studentId, targetDate);
      if (directPhotoUrl && directPhotoUrl.indexOf('https://') === 0 && directPhotoUrl.indexOf('data:') === -1 && directPhotoUrl.length < 1000) {
        colContents.push({
          "type": "box",
          "layout": "vertical",
          "width": isChamp ? "56px" : "44px",
          "height": isChamp ? "56px" : "44px",
          "cornerRadius": "100px",
          "borderColor": crownColor,
          "borderWidth": isChamp ? "bold" : "normal",
          "margin": "xs",
          "contents": [
            {
              "type": "image",
              "url": directPhotoUrl,
              "size": "full",
              "aspectRatio": "1:1",
              "aspectMode": "cover"
            }
          ]
        });
      } else {
        colContents.push({
          "type": "box",
          "layout": "vertical",
          "width": isChamp ? "44px" : "36px",
          "height": isChamp ? "44px" : "36px",
          "cornerRadius": "100px",
          "backgroundColor": isChamp ? "#3B2800" : "#1E293B",
          "borderColor": crownColor,
          "borderWidth": isChamp ? "bold" : "normal",
          "alignItems": "center",
          "justifyContent": "center",
          "margin": "xs",
          "contents": [
            {
              "type": "text",
              "text": medalEmoji,
              "size": isChamp ? "md" : "sm",
              "align": "center"
            }
          ]
        });
      }

      // Rank Pill Badge
      colContents.push({
        "type": "box",
        "layout": "vertical",
        "backgroundColor": badgeBgHex,
        "cornerRadius": "4px",
        "paddingStart": "4px",
        "paddingEnd": "4px",
        "paddingTop": "1px",
        "paddingBottom": "1px",
        "margin": "xs",
        "contents": [
          {
            "type": "text",
            "text": medalEmoji + " " + rankLabel,
            "size": "xxs",
            "color": crownColor,
            "weight": "bold",
            "align": "center"
          }
        ]
      });

      // Student Name
      var rClassTag = (isMultiClass && item.className) ? ' (' + item.className + ')' : '';
      colContents.push({
        "type": "text",
        "text": item.name + rClassTag,
        "size": "xxs",
        "color": "#FFFFFF",
        "weight": "bold",
        "wrap": true,
        "maxLines": 2,
        "align": "center",
        "margin": "xs"
      });

      // Check-in Time
      colContents.push({
        "type": "text",
        "text": "⏰ " + item.time + " น.",
        "size": "xxs",
        "color": isChamp ? "#FDE047" : "#CBD5E1",
        "weight": isChamp ? "bold" : "regular",
        "align": "center",
        "margin": "none"
      });

      return {
        "type": "box",
        "layout": "vertical",
        "flex": isChamp ? 2 : 1,
        "alignItems": "center",
        "paddingAll": "4px",
        "contents": colContents
      };
    }

    // 2nd Place (Silver) - Left
    if (top2) {
      championColumns.push(buildFlexChampCol(top2, '#2', '🥈', 'อันดับ 2', '#C0C0C0', '#334155', false));
    } else {
      championColumns.push({ "type": "box", "layout": "vertical", "flex": 1, "contents": [{ "type": "text", "text": " ", "size": "xxs" }] });
    }

    // 1st Place (Gold Champion) - Center
    championColumns.push(buildFlexChampCol(top1, '#1', '🥇', 'อันดับ 1', '#FFD700', '#78350F', true));

    // 3rd Place (Bronze) - Right
    if (top3) {
      championColumns.push(buildFlexChampCol(top3, '#3', '🥉', 'อันดับ 3', '#F97316', '#431407', false));
    } else {
      championColumns.push({ "type": "box", "layout": "vertical", "flex": 1, "contents": [{ "type": "text", "text": " ", "size": "xxs" }] });
    }

    bodyContents.push({
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#0F172A",
      "borderColor": "#FBBF24",
      "borderWidth": "semi-bold",
      "cornerRadius": "10px",
      "paddingAll": "10px",
      "margin": "md",
      "contents": [
        {
          "type": "text",
          "text": "👑 🏆 3 อันดับแรกที่เช็คชื่อไวที่สุดประจำวัน",
          "weight": "bold",
          "size": "xs",
          "color": "#FFD700",
          "align": "center"
        },
        {
          "type": "box",
          "layout": "horizontal",
          "alignItems": "flex-end",
          "margin": "sm",
          "contents": championColumns
        }
      ]
    });
  }

  bodyContents.push({
    "type": "separator",
    "color": "#F1F5F9"
  });

  // 2. Statistics Grid Rows
  var statRows = [
    {
      label: "มาเรียนตรงเวลา",
      count: countPresent + " คน",
      color: "#10B981",
      iconText: "มาเรียน",
      subText: "ตรงเวลา",
      desc: "เข้าเรียนตามเวลาปกติ"
    }
  ];

  if (countLate > 0) {
    statRows.push({
      label: "มาสาย",
      count: countLate + " คน",
      color: "#FF6B00",
      iconText: "มาสาย",
      subText: countLate + " คน",
      desc: "เกินเวลาที่กำหนด"
    });
  }

  if (countSick + countLeave > 0) {
    statRows.push({
      label: "ลากิจ / ลาป่วย",
      count: (countSick + countLeave) + " คน",
      color: "#EAB308",
      iconText: "แจ้งลา",
      subText: (countSick + countLeave) + " คน",
      desc: "ป่วย " + countSick + " • กิจ " + countLeave
    });
  }

  if (countAbsent > 0) {
    statRows.push({
      label: "ขาดเรียน",
      count: countAbsent + " คน",
      color: "#EF4444",
      iconText: "ขาดเรียน",
      subText: countAbsent + " คน",
      desc: "ไม่พบการเช็คชื่อในระบบ"
    });
  }

  for (var r = 0; r < statRows.length; r++) {
    var row = statRows[r];
    bodyContents.push({
      "type": "box",
      "layout": "horizontal",
      "paddingTop": "8px",
      "paddingBottom": "8px",
      "contents": [
        {
          "type": "box",
          "layout": "vertical",
          "width": "85px",
          "contents": [
            {
              "type": "text",
              "text": row.iconText,
              "weight": "bold",
              "size": "sm",
              "color": row.color
            },
            {
              "type": "text",
              "text": row.subText,
              "size": "xxs",
              "color": "#94A3B8"
            }
          ]
        },
        {
          "type": "box",
          "layout": "vertical",
          "flex": 1,
          "contents": [
            {
              "type": "text",
              "text": row.label + " (" + row.count + ")",
              "weight": "bold",
              "size": "sm",
              "color": "#1E293B"
            },
            {
              "type": "text",
              "text": row.desc,
              "size": "xs",
              "color": "#64748B",
              "margin": "xs"
            }
          ]
        }
      ]
    });

    if (r < statRows.length - 1) {
      bodyContents.push({
        "type": "separator",
        "color": "#F1F5F9"
      });
    }
  }

  // 3. Exception & Attendance Details List
  bodyContents.push({
    "type": "separator",
    "margin": "md",
    "color": "#E2E8F0"
  });

  bodyContents.push({
    "type": "text",
    "text": "📌 รายชื่อและการลงเวลาเช็คชื่อ",
    "weight": "bold",
    "size": "xs",
    "color": "#475569",
    "margin": "md"
  });

  // A. PRESENT STUDENTS (มาเรียนตรงเวลา เริ่มจากลำดับที่ 4, 5, 6... เพราะ 1-3 อยู่ในโพเดียมแล้ว)
  var sortedAllPresent = [];
  for (var sap = 0; sap < presentList.length; sap++) {
    var sTime = formatTimeString(presentList[sap].time);
    sortedAllPresent.push({
      no: presentList[sap].no,
      name: presentList[sap].name,
      className: presentList[sap].className,
      time: sTime,
      selfiePhoto: presentList[sap].selfiePhoto || ''
    });
  }
  sortedAllPresent.sort(function(a, b) {
    if (a.time !== '-' && b.time !== '-' && a.time !== b.time) {
      return String(a.time).localeCompare(String(b.time));
    }
    return (parseInt(a.no, 10) || 0) - (parseInt(b.no, 10) || 0);
  });

  if (sortedAllPresent.length > 3) {
    var prTexts = [];
    var remCount = sortedAllPresent.length - 3;
    for (var p = 3; p < Math.min(sortedAllPresent.length, 15); p++) {
      var pItem = sortedAllPresent[p];
      var pClass = (isMultiClass || pItem.className) ? " (" + (pItem.className || "") + ")" : "";
      var pTime = pItem.time;
      var timeTag = (pTime && pTime !== '-') ? " ⏰ " + pTime + " น." : "";
      prTexts.push((p + 1) + ". เลขที่ " + pItem.no + " " + pItem.name + pClass + timeTag);
    }
    if (sortedAllPresent.length > 15) prTexts.push("...และอีก " + (sortedAllPresent.length - 15) + " คน");

    bodyContents.push({
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#F0FDF4",
      "cornerRadius": "6px",
      "paddingAll": "8px",
      "margin": "sm",
      "contents": [
        {
          "type": "text",
          "text": "🟢 ผู้ที่มาเรียนตรงเวลา (ลำดับที่ 4 เป็นต้นไป - " + remCount + " คน):",
          "weight": "bold",
          "size": "xxs",
          "color": "#16A34A"
        },
        {
          "type": "text",
          "text": prTexts.join("\n"),
          "size": "xxs",
          "color": "#166534",
          "wrap": true,
          "margin": "xs"
        }
      ]
    });
  } else if (sortedAllPresent.length > 0) {
    bodyContents.push({
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#F0FDF4",
      "cornerRadius": "6px",
      "paddingAll": "8px",
      "margin": "sm",
      "contents": [
        {
          "type": "text",
          "text": "🟢 ผู้ที่มาเรียนตรงเวลา (" + sortedAllPresent.length + " คน):",
          "weight": "bold",
          "size": "xxs",
          "color": "#16A34A"
        },
        {
          "type": "text",
          "text": "✨ นักเรียนที่มาตรงเวลาทุกคนได้รับเกียรติอยู่ในโพเดียม 3 อันดับแรกแล้ว",
          "size": "xxs",
          "color": "#15803D",
          "wrap": true,
          "margin": "xs"
        }
      ]
    });
  }

  // B. LATE STUDENTS (มาสาย)
  if (lateList.length > 0) {
    var ltTexts = [];
    for (var lt = 0; lt < Math.min(lateList.length, 8); lt++) {
      var ltItem = lateList[lt];
      var ltClass = (isMultiClass || ltItem.className) ? " (" + (ltItem.className || "") + ")" : "";
      var ltTime = formatTimeString(ltItem.time);
      var ltTimeTag = (ltTime && ltTime !== '-') ? " ⏰ " + ltTime + " น." : "";
      ltTexts.push("• " + ltItem.no + ". " + ltItem.name + ltClass + ltTimeTag);
    }
    if (lateList.length > 8) ltTexts.push("...และอีก " + (lateList.length - 8) + " คน");

    bodyContents.push({
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#FFFBEB",
      "cornerRadius": "6px",
      "paddingAll": "8px",
      "margin": "sm",
      "contents": [
        {
          "type": "text",
          "text": "🌸 มาสาย (" + lateList.length + " คน):",
          "weight": "bold",
          "size": "xxs",
          "color": "#D97706"
        },
        {
          "type": "text",
          "text": ltTexts.join("\n"),
          "size": "xxs",
          "color": "#92400E",
          "wrap": true,
          "margin": "xs"
        }
      ]
    });
  }

  // C. LEAVE STUDENTS (แจ้งลา)
  if (sickList.length > 0 || leaveList.length > 0) {
    var lvTexts = [];
    for (var sk = 0; sk < Math.min(sickList.length, 6); sk++) {
      var skItem = sickList[sk];
      var skClass = (isMultiClass || skItem.className) ? " (" + (skItem.className || "") + ")" : "";
      lvTexts.push("• [ป่วย] " + skItem.no + ". " + skItem.name + skClass + (skItem.note ? " [" + skItem.note + "]" : ""));
    }
    for (var lv = 0; lv < Math.min(leaveList.length, 6); lv++) {
      var lvItem = leaveList[lv];
      var lvClass = (isMultiClass || lvItem.className) ? " (" + (lvItem.className || "") + ")" : "";
      lvTexts.push("• [กิจ] " + lvItem.no + ". " + lvItem.name + lvClass + (lvItem.note ? " [" + lvItem.note + "]" : ""));
    }

    bodyContents.push({
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#FEFCE8",
      "cornerRadius": "6px",
      "paddingAll": "8px",
      "margin": "sm",
      "contents": [
        {
          "type": "text",
          "text": "🟡 แจ้งลา (" + (sickList.length + leaveList.length) + " คน):",
          "weight": "bold",
          "size": "xxs",
          "color": "#CA8A04"
        },
        {
          "type": "text",
          "text": lvTexts.join("\n"),
          "size": "xxs",
          "color": "#854D0E",
          "wrap": true,
          "margin": "xs"
        }
      ]
    });
  }

  // D. ABSENT STUDENTS (ขาดเรียน)
  if (absentList.length > 0) {
    var abTexts = [];
    for (var a = 0; a < Math.min(absentList.length, 8); a++) {
      var abItem = absentList[a];
      var abClass = (isMultiClass || abItem.className) ? " (" + (abItem.className || "") + ")" : "";
      abTexts.push("• " + abItem.no + ". " + abItem.name + abClass);
    }
    if (absentList.length > 8) abTexts.push("...และอีก " + (absentList.length - 8) + " คน");

    bodyContents.push({
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#FEF2F2",
      "cornerRadius": "6px",
      "paddingAll": "8px",
      "margin": "sm",
      "contents": [
        {
          "type": "text",
          "text": "🔴 ไม่พบการเช็คชื่อ (" + absentList.length + " คน):",
          "weight": "bold",
          "size": "xxs",
          "color": "#DC2626"
        },
        {
          "type": "text",
          "text": abTexts.join("\n"),
          "size": "xxs",
          "color": "#991B1B",
          "wrap": true,
          "margin": "xs"
        }
      ]
    });
  }

  var flexBubble = {
    "type": "bubble",
    "size": "mega",
    "header": {
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#1A1A1A",
      "paddingTop": "18px",
      "paddingBottom": "18px",
      "paddingStart": "20px",
      "paddingEnd": "20px",
      "contents": [
        {
          "type": "text",
          "text": "🟠 สรุปรายงานการมาเรียนประจำวัน",
          "weight": "bold",
          "color": "#FF6B00",
          "size": "xs"
        },
        {
          "type": "text",
          "text": dayName,
          "weight": "bold",
          "color": "#FFFFFF",
          "size": "xxl",
          "margin": "sm"
        },
        {
          "type": "text",
          "text": formatThaiDate(targetDate) + " • " + classLabel + (actLabel !== 'เช็คชื่อมาเรียน' ? " • " + actLabel : "") + " (" + totalStudents + " คน)",
          "color": "#9CA3AF",
          "size": "xs",
          "margin": "xs"
        }
      ]
    },
    "body": {
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#FFFFFF",
      "paddingAll": "16px",
      "contents": bodyContents
    },
    "footer": {
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#1A1A1A",
      "paddingTop": "12px",
      "paddingBottom": "12px",
      "paddingStart": "16px",
      "paddingEnd": "16px",
      "contents": (function() {
        var fArr = [
          {
            "type": "text",
            "text": "ขอให้เป็นวันที่ยอดเยี่ยมในการเรียนการสอนครับ ✨",
            "color": "#F59E0B",
            "size": "xs",
            "weight": "bold",
            "align": "center",
            "wrap": true
          }
        ];
        try {
          var wUrl = ScriptApp.getService().getUrl();
          if (wUrl && wUrl.indexOf('http') === 0) {
            var rLink = wUrl + '?page=report&date=' + encodeURIComponent(targetDate) + '&class=' + encodeURIComponent(data.targetClass || data.className || 'all') + '&activity=' + encodeURIComponent(data.targetActivity || data.activityName || 'all');
            fArr.push({
              "type": "button",
              "style": "primary",
              "color": "#06C755",
              "height": "sm",
              "margin": "md",
              "action": {
                "type": "uri",
                "label": "📱 ดูรายงานออนไลน์",
                "uri": rLink
              }
            });
          }
        } catch (ex) {}
        return fArr;
      })()
    }
  };

  return {
    type: "flex",
    altText: "📊 สรุปรายงานการมาเรียน - " + dayName + " " + targetDate + " (" + classLabel + ")",
    contents: flexBubble
  };
}


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
var SHEET_SCHEDULE = 'Schedule';

var ADMIN_PASSWORD = 'admin888';

/**
 * Serves the HTML Web Application or handles API GET requests
 */
function doGet(e) {
  if (e && e.parameter) {
    if (e.parameter.page === 'report1.2.5' || e.parameter.page === 'Report1.2.5' || e.parameter.view === 'report1.2.5' || e.parameter.p === 'report1.2.5' || e.parameter.page === 'report1-2.5' || e.parameter.page === 'Report 1-2.5' || e.parameter.page === 'report125') {
      try {
        var repTemplate125 = HtmlService.createTemplateFromFile('Report1.2.5');
        return repTemplate125.evaluate()
          .setTitle('รายงานเช็คชื่อ ปวส.1/2 และ ปวส.1/5 - ภาพเซลฟี่ยืนยันตัวตน')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } catch (err) {
        var repTemplateAlt125 = HtmlService.createTemplateFromFile('report1.2.5');
        return repTemplateAlt125.evaluate()
          .setTitle('รายงานเช็คชื่อ ปวส.1/2 และ ปวส.1/5 - ภาพเซลฟี่ยืนยันตัวตน')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    }
    if (e.parameter.page === 'report' || e.parameter.view === 'report' || e.parameter.p === 'report') {
      try {
        var repTemplate = HtmlService.createTemplateFromFile('report');
        return repTemplate.evaluate()
          .setTitle('รายงานสรุปการมาเรียนประจำวัน - สำหรับผู้ปกครอง')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } catch (err) {
        var repTemplateAlt = HtmlService.createTemplateFromFile('Report');
        return repTemplateAlt.evaluate()
          .setTitle('รายงานสรุปการมาเรียนประจำวัน - สำหรับผู้ปกครอง')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    }
    if (e.parameter.action) {
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

/**
 * Public API for Parent Daily Report Page (Read-only, high performance)
 */
function getPublicDailyReportData(dateStr, className, activityName) {
  var targetDate = dateStr ? formatDateString(dateStr) : Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  var cleanClass = String(className || 'all').trim();
  var cleanAct = String(activityName || 'all').trim();
  var actLabel = (cleanAct === 'all' || !cleanAct) ? 'เช็คชื่อมาเรียน' : cleanAct;
  
  // Filter students strictly for the specified class(es) only
  var students = getStudents(cleanClass);
  var recordsMap = getAttendanceRecords(targetDate, cleanClass, cleanAct, 'daily');
  
  var totalStudents = students.length;
  var countPresent = 0;
  var countOnTime = 0;
  var countLate = 0;
  var countSick = 0;
  var countLeave = 0;
  var countAbsent = 0;
  
  var studentReports = [];
  var onTimePresentList = [];
  
  for (var i = 0; i < students.length; i++) {
    var s = students[i];
    var sId = String(s.studentId || '').trim();
    var cleanId = sId.toLowerCase();
    var rec = recordsMap[sId] || recordsMap[cleanId];
    
    var st = 'Absent';
    var note = '';
    var cTime = '-';
    var photo = s.photo || s.selfiePhoto || s.avatar || s.imageUrl || '';
    var dist = 0;
    
    if (rec) {
      if (typeof rec === 'object') {
        st = rec.status || 'Present';
        note = rec.note || '';
        cTime = rec.checkinTime || '-';
        if (rec.selfiePhoto) photo = rec.selfiePhoto;
        dist = rec.distanceMeters || 0;
      } else {
        st = String(rec);
      }
    }
    
    var statusLabel = 'ขาดเรียน';
    var statusClass = 'absent';
    
    if (st === 'Present' || st === 'มาเรียน') {
      countPresent++;
      countOnTime++;
      statusLabel = 'มาเรียนตรงเวลา';
      statusClass = 'present';
      onTimePresentList.push({
        no: s.studentNo || (i + 1),
        studentId: s.studentId,
        name: s.name,
        className: s.className,
        time: cTime,
        selfiePhoto: photo
      });
    } else if (st === 'Late' || st === 'มาสาย') {
      countPresent++;
      countLate++;
      statusLabel = 'มาสาย';
      statusClass = 'late';
    } else if (st === 'Sick' || st === 'ลาป่วย') {
      countSick++;
      statusLabel = 'ลาป่วย';
      statusClass = 'sick';
    } else if (st === 'Leave' || st === 'ลากิจ') {
      countLeave++;
      statusLabel = 'ลากิจ';
      statusClass = 'leave';
    } else {
      countAbsent++;
      statusLabel = 'ขาดเรียน';
      statusClass = 'absent';
    }
    
    var convertedPhoto = photo ? convertSelfieToPublicUrl(photo, s.studentId, targetDate) : '';
    var rowActName = (rec && rec.activityName) ? rec.activityName : actLabel;
    studentReports.push({
      no: s.studentNo || (i + 1),
      studentId: s.studentId,
      name: s.name,
      className: s.className,
      activityName: rowActName,
      status: st,
      statusLabel: statusLabel,
      statusClass: statusClass,
      checkinTime: cTime,
      selfiePhoto: convertedPhoto || photo,
      distanceMeters: dist,
      note: note
    });
  }
  
  // Sort Top 3 Early Birds
  onTimePresentList.sort(function(a, b) {
    return String(a.time).localeCompare(String(b.time));
  });
  
  var topChampions = [];
  for (var c = 0; c < Math.min(3, onTimePresentList.length); c++) {
    var champItem = onTimePresentList[c];
    champItem.rank = c + 1;
    champItem.selfiePhoto = convertSelfieToPublicUrl(champItem.selfiePhoto, champItem.no || champItem.studentId, targetDate);
    topChampions.push(champItem);
  }
  
  var attRate = totalStudents > 0 ? ((countPresent / totalStudents) * 100).toFixed(1) : '0.0';
  
  var availClasses = [];
  try { availClasses = getClassList(); } catch (e) { console.warn('getClassList err:', e); }
  var availActivities = [];
  try {
    var actObj = getActivityList();
    if (actObj && actObj.activities) availActivities = actObj.activities;
  } catch (e) { console.warn('getActivityList err:', e); }

  return {
    status: 'success',
    date: targetDate,
    availableClasses: availClasses,
    availableActivities: availActivities,
    dateThai: formatThaiDate(targetDate),
    dayName: getThaiDayOfWeek(targetDate),
    className: (cleanClass === 'all' || cleanClass === 'ทุกห้องเรียน' || cleanClass === 'ทุกห้อง') ? 'ทุกห้องเรียน' : ('ห้อง ' + cleanClass.replace(/^ห้อง\s*/, '')),
    activityName: actLabel,
    stats: {
      totalStudents: totalStudents,
      countPresent: countPresent,
      countOnTime: countOnTime,
      countLate: countLate,
      countSick: countSick,
      countLeave: countLeave,
      countAbsent: countAbsent,
      attRate: attRate
    },
    topChampions: topChampions,
    students: studentReports
  };
}


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
        return getAttendanceRecords(arg(0, 'dateStr', ''), arg(1, 'className', ''), arg(2, 'activityName', ''), arg(3, 'reportType', 'daily'), arg(4, 'endDateStr', ''));
      case 'getPublicDailyReportData':
        return getPublicDailyReportData(arg(0, 'dateStr', ''), arg(1, 'className', ''), arg(2, 'activityName', ''));
      case 'getStudentPortalLiveAttendance':
        return getStudentPortalLiveAttendance(arg(0, 'className', 'all'), arg(1, 'sessionId', ''));
      case 'getRangeAttendanceSummary':
        return getRangeAttendanceSummary(arg(0, 'startDateStr', ''), arg(1, 'endDateStr', ''), arg(2, 'className', ''), arg(3, 'activityName', ''));
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
      case 'getAutoCheckinSchedules':
        return getAutoCheckinSchedules();
      case 'saveAutoCheckinSchedules':
        return saveAutoCheckinSchedules(arg(0, 'schedules', data));
      case 'resolveActiveSchedule':
        return resolveActiveSchedule(arg(0, 'className', 'all'), arg(1, 'clientDateStr', ''), arg(2, 'clientTimeStr', ''));
      case 'saveStudentFacePhoto':
        return saveStudentFacePhoto(arg(0, 'studentId', ''), arg(1, 'photoBase64', ''));
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
  if (!filter || filter === 'all' || filter === 'ทุกกิจกรรม' || filter === 'ทุกรายวิชา' || filter === 'ทุกกิจกรรม / ทุกรายวิชา' || filter === 'undefined') return true;
  var sAct = String(activityName || 'เช็คชื่อมาเรียน').trim().toLowerCase();
  var parts = String(filter).split(',');
  for (var p = 0; p < parts.length; p++) {
    var item = parts[p].trim().toLowerCase();
    if (!item || item === 'all' || item === 'ทุกกิจกรรม' || item === 'ทุกรายวิชา' || item === 'ทุกกิจกรรม / ทุกรายวิชา' || item === sAct || sAct.indexOf(item) > -1 || item.indexOf(sAct) > -1) {
      return true;
    }
  }
  return false;
}

function isClassMatch(filter, studentClass) {
  if (!filter || filter === 'all' || filter === 'ทุกห้องเรียน' || filter === 'ทุกห้อง' || filter === 'undefined') return true;
  var sClass = String(studentClass || '').trim().toLowerCase().replace(/^ห้อง\s*/, '');
  if (!sClass) return false;
  var parts = String(filter).split(',');
  for (var p = 0; p < parts.length; p++) {
    var item = parts[p].trim().toLowerCase().replace(/^ห้อง\s*/, '');
    if (!item || item === 'all' || item === 'ทุกห้องเรียน' || item === 'ทุกห้อง' || item === sClass || sClass === item || sClass.indexOf(item) > -1 || item.indexOf(sClass) > -1) {
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
        var photo = row[5] ? String(row[5]).trim() : '';
        
        if (!name) continue;
        allStudents.push({
          studentId: studentId,
          studentNo: studentNo,
          name: name,
          className: className,
          deviceId: deviceId,
          photo: photo
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

function getActiveSession(className, sessionId, activityName) {
  var activeList = getActiveSessions();
  var sIdClean = String(sessionId || '').trim();
  var cNameClean = String(className || 'all').trim();
  var actNameClean = String(activityName || '').trim();

  // Priority 1: Match EXACT sessionId from active list
  if (sIdClean && sIdClean !== 'all' && sIdClean !== 'undefined') {
    for (var i = 0; i < activeList.length; i++) {
      if (activeList[i].sessionId === sIdClean) {
        return activeList[i];
      }
    }
  }

  // Priority 2: If sessionId is an automated schedule sessionId
  if (sIdClean && sIdClean.indexOf('AUTOSES_') === 0) {
    var autoResExact = resolveActiveSchedule(cNameClean);
    if (autoResExact && autoResExact.status === 'open') {
      return {
        sessionId: autoResExact.sessionId,
        className: autoResExact.className || 'all',
        activityName: autoResExact.activityName || 'เช็คชื่อมาเรียน',
        date: autoResExact.dateStr,
        createdAt: new Date().toISOString(),
        expiresAt: new Date(Date.now() + 3600000).toISOString(),
        lat: autoResExact.gpsLat || 0,
        lng: autoResExact.gpsLng || 0,
        radiusMeters: autoResExact.radiusMeters || 100,
        secretSeed: autoResExact.secretSeed,
        otpIntervalSeconds: autoResExact.otpIntervalSeconds || 15,
        requireOTP: autoResExact.requireOTP !== false,
        requireSelfie: autoResExact.requireSelfie !== false,
        requireGPS: autoResExact.requireGPS === true,
        requireDeviceBinding: true,
        status: 'Active',
        isAutomated: true,
        sessionPhase: autoResExact.sessionPhase,
        statusResult: autoResExact.statusResult,
        serverTimeMs: Date.now()
      };
    }
  }

  // Priority 3: Match both Activity Name & Class Name
  if (actNameClean && actNameClean !== 'all') {
    for (var a = 0; a < activeList.length; a++) {
      var itm = activeList[a];
      var isActOk = isActivityMatch(actNameClean, itm.activityName);
      var isClsOk = (cNameClean === 'all' || isClassMatch(itm.className, cNameClean) || isClassMatch(cNameClean, itm.className));
      if (isActOk && isClsOk) {
        return itm;
      }
    }
  }

  // Priority 4: Match Class Name
  if (cNameClean && cNameClean !== 'all') {
    for (var j = 0; j < activeList.length; j++) {
      var cName = activeList[j].className;
      if (cName === 'all' || isClassMatch(cName, cNameClean) || isClassMatch(cNameClean, cName)) {
        return activeList[j];
      }
    }
  }

  // Priority 5: Latest active session
  if (activeList.length > 0) {
    return activeList[activeList.length - 1];
  }

  // Priority 6: Fallback to dynamic automated daily schedule session if active!
  var autoRes = resolveActiveSchedule(cNameClean || 'all');
  if (autoRes && autoRes.status === 'open') {
    return {
      sessionId: autoRes.sessionId,
      className: autoRes.className || 'all',
      activityName: autoRes.activityName || 'เช็คชื่อมาเรียน',
      date: autoRes.dateStr,
      createdAt: new Date().toISOString(),
      expiresAt: new Date(Date.now() + 3600000).toISOString(),
      lat: autoRes.gpsLat || 0,
      lng: autoRes.gpsLng || 0,
      radiusMeters: autoRes.radiusMeters || 100,
      secretSeed: autoRes.secretSeed,
      otpIntervalSeconds: autoRes.otpIntervalSeconds || 15,
      requireOTP: autoRes.requireOTP !== false,
      requireSelfie: autoRes.requireSelfie !== false,
      requireGPS: autoRes.requireGPS === true,
      requireDeviceBinding: true,
      status: 'Active',
      isAutomated: true,
      sessionPhase: autoRes.sessionPhase,
      statusResult: autoRes.statusResult,
      serverTimeMs: Date.now()
    };
  }

  return null;
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
  
  var sessionId = String(payload.sessionId || '').trim();
  var studentId = String(payload.studentId || '').trim();
  var selectedClass = String(payload.className || payload.selectedClass || '').trim();
  var enteredOtp = String(payload.otp || '').trim();
  var studentLat = parseFloat(payload.lat);
  var studentLng = parseFloat(payload.lng);
  var deviceId = String(payload.deviceId || '').trim();
  var selfieBase64 = payload.selfieBase64 || '';
  
  if (!studentId) {
    return { status: 'error', message: 'กรุณาระบุรหัสนักศึกษา' };
  }
  
  // 1. Resolve the active session
  var activeSession = getActiveSession(selectedClass, sessionId, payload.activityName);
  if (!activeSession) {
    return { status: 'error', message: 'ไม่มีคาบเรียนหรือกิจกรรมที่เปิดรับเช็คชื่อสำหรับห้องนี้ในขณะนี้' };
  }

  // 2. Multi-Roster Security: Find student record matching both studentId AND the active session / selected class
  var students = getStudents();
  var studentObj = null;

  // Search Priority 1: Match studentId AND activeSession.className
  for (var s = 0; s < students.length; s++) {
    if (String(students[s].studentId).trim() === studentId) {
      if (activeSession.className === 'all' || isClassMatch(activeSession.className, students[s].className)) {
        studentObj = students[s];
        break;
      }
    }
  }

  // Search Priority 2: Match studentId AND selectedClass
  if (!studentObj && selectedClass) {
    for (var s2 = 0; s2 < students.length; s2++) {
      if (String(students[s2].studentId).trim() === studentId) {
        if (isClassMatch(selectedClass, students[s2].className)) {
          studentObj = students[s2];
          break;
        }
      }
    }
  }

  // Search Priority 3: Fallback to any matching studentId
  if (!studentObj) {
    for (var s3 = 0; s3 < students.length; s3++) {
      if (String(students[s3].studentId).trim() === studentId) {
        studentObj = students[s3];
        break;
      }
    }
  }

  if (!studentObj) {
    return { status: 'error', message: 'ไม่พบรหัสนักศึกษานี้ในระบบ' };
  }

  // 3. Strict Classroom Guard: Verify that the student's class matches the active session
  if (activeSession.className && activeSession.className !== 'all') {
    var isClassOk = isClassMatch(activeSession.className, studentObj.className) || 
                    (selectedClass && isClassMatch(activeSession.className, selectedClass));
    if (!isClassOk) {
      return {
        status: 'error',
        message: 'คุณเลือกห้อง/กลุ่ม ("' + (selectedClass || studentObj.className) + '") ซึ่งไม่ได้เปิดให้เช็คชื่อในขณะนี้ (ขณะนี้เปิดเฉพาะ: ห้อง ' + activeSession.className + ' - ' + activeSession.activityName + ')'
      };
    }
  }
  
  var now = new Date();
  var nowMs = now.getTime();
  var intervalSec = activeSession.otpIntervalSeconds || 15;
  var intervalMs = intervalSec * 1000;
  
  if (activeSession.requireOTP !== false && activeSession.requireOtp !== false && activeSession.useOtp !== false) {
    if (!enteredOtp || enteredOtp.length < 6) {
      return {
        status: 'error',
        message: 'กรุณากรอกรหัส OTP 6 หลัก'
      };
    }
    var validOtpCurrent = generateTOTP(activeSession.secretSeed, intervalSec, nowMs);
    var validOtpPrev = generateTOTP(activeSession.secretSeed, intervalSec, nowMs - intervalMs);
    
    if (enteredOtp !== validOtpCurrent && enteredOtp !== validOtpPrev) {
      return {
        status: 'error',
        message: 'รหัส OTP 6 หลักไม่ถูกต้องหรือหมดอายุแล้ว กรุณาดูรหัสล่าสุดบนหน้าจอโปรเจกเตอร์'
      };
    }
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
  var serverThaiDate = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');
  var dateStr = (activeSession && activeSession.date) ? activeSession.date : serverThaiDate;
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
  
  var checkinTimeStr = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');
  var gpsLocStr = (!isNaN(studentLat) && !isNaN(studentLng) && studentLat !== 0) ? (studentLat.toFixed(6) + ',' + studentLng.toFixed(6)) : '-';
  
  // Determine Present vs Late status
  var checkinStatus = 'Present';
  var isLate = false;
  if (activeSession.isAutomated && activeSession.statusResult) {
    checkinStatus = activeSession.statusResult;
    isLate = (checkinStatus === 'Late');
  } else if (activeSession.lateAtMs && activeSession.lateAtMs > 0 && nowMs > activeSession.lateAtMs) {
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
  
  // Convert Base64 selfie to direct public Google CDN URL for LINE Flex support & fast sheets
  var selfieToStore = selfieBase64;
  if (selfieBase64 && (selfieBase64.indexOf('data:image/') === 0 || selfieBase64.indexOf('base64,') > -1 || selfieBase64.length > 500)) {
    try {
      var pubUrl = convertSelfieToPublicUrl(selfieBase64, studentId, dateStr);
      if (pubUrl && pubUrl.indexOf('https://') === 0) {
        selfieToStore = pubUrl;
      }
    } catch (e) {}
  }

  if (existingRowIndex > 0) {
    attSheet.getRange(existingRowIndex, 3).setValue(checkinStatus);
    attSheet.getRange(existingRowIndex, 4).setValue(checkinTimeStr);
    attSheet.getRange(existingRowIndex, 5).setValue(gpsLocStr);
    attSheet.getRange(existingRowIndex, 6).setValue(distanceMeters);
    attSheet.getRange(existingRowIndex, 7).setValue(deviceId);
    attSheet.getRange(existingRowIndex, 8).setValue(selfieToStore);
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
      selfieToStore,
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

/**
 * Realtime Live Attendance for Student Portal
 * Returns list of students who have already checked in if a session or schedule is active/open.
 * If the session is closed/expired, returns status: 'closed' and isOpen: false so the UI stops displaying.
 */
function getStudentPortalLiveAttendance(className, sessionId) {
  className = String(className || 'all').trim();
  sessionId = String(sessionId || '').trim();

  var ss = getSpreadsheet();
  var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  var sessionSheet = ss.getSheetByName(SHEET_SESSIONS);
  if (!attSheet) return { status: 'closed', isOpen: false, list: [], count: 0, countPresent: 0, countLate: 0, total: 0 };

  var todayStr = formatDateString(new Date());

  // 1. Check for Active Manual Projector Sessions
  var activeManualSessions = getActiveSessions() || [];
  var matchedManual = null;
  if (sessionId) {
    for (var m = 0; m < activeManualSessions.length; m++) {
      if (String(activeManualSessions[m].sessionId) === String(sessionId)) {
        matchedManual = activeManualSessions[m];
        break;
      }
    }
  }
  if (!matchedManual && className !== 'all' && className !== '') {
    for (var m2 = 0; m2 < activeManualSessions.length; m2++) {
      if (isClassMatch(activeManualSessions[m2].className, className)) {
        matchedManual = activeManualSessions[m2];
        break;
      }
    }
  }
  if (!matchedManual && activeManualSessions.length > 0) {
    matchedManual = activeManualSessions[0];
  }

  var targetSession = null;
  if (matchedManual && matchedManual.status === 'Active') {
    targetSession = {
      sessionId: matchedManual.sessionId,
      className: matchedManual.className,
      activityName: matchedManual.activityName || 'เช็คชื่อมาเรียน',
      date: matchedManual.date || todayStr,
      isAutomated: false,
      status: 'Active'
    };
  } else {
    // 2. Check Automated Daily Schedule
    var autoRes = resolveActiveSchedule(className || 'all');
    if (autoRes && autoRes.status === 'open') {
      targetSession = {
        sessionId: autoRes.sessionId,
        className: autoRes.className,
        activityName: autoRes.activityName || 'เช็คชื่อมาเรียน',
        date: autoRes.dateStr || todayStr,
        isAutomated: true,
        phaseLabel: autoRes.phaseLabel,
        status: 'Active'
      };
    }
  }

  // If NO session is open right now -> Return closed state immediately
  if (!targetSession) {
    return {
      status: 'closed',
      isOpen: false,
      list: [],
      count: 0,
      countPresent: 0,
      countLate: 0,
      total: 0
    };
  }

  // Fetch student roster to match names and student numbers
  var students = getStudents(targetSession.className || className || 'all');
  var studentMap = {};
  for (var s = 0; s < students.length; s++) {
    var stdIdKey = String(students[s].studentId).trim().toLowerCase();
    studentMap[stdIdKey] = students[s];
  }

  var attData = attSheet.getDataRange().getValues();
  var checkedList = [];
  var countPresent = 0;
  var countLate = 0;

  for (var j = 1; j < attData.length; j++) {
    var rowD = formatDateString(attData[j][0]);
    var sIdRaw = String(attData[j][1]).trim();
    var sIdKey = sIdRaw.toLowerCase();
    var status = String(attData[j][2]).trim();
    var checkinTime = formatTimeString(attData[j][3]);
    var gpsLoc = String(attData[j][4] || '-').trim();
    var dist = attData[j][5] || 0;
    var selfie = String(attData[j][7] || '').trim();
    var rowAct = String(attData[j][8] || 'เช็คชื่อมาเรียน').trim();
    var rowSes = String(attData[j][9] || '').trim();

    if (rowD === targetSession.date && (status === 'Present' || status === 'มาเรียน' || status === 'Late' || status === 'มาสาย')) {
      var isActMatch = isActivityMatch(targetSession.activityName, rowAct);
      var isSesMatch = (rowSes === targetSession.sessionId || !rowSes || !targetSession.sessionId);
      var stObj = studentMap[sIdKey];

      // If class is specified, filter by student's class
      var isClassOk = true;
      if (className && className !== 'all' && stObj) {
        isClassOk = isClassMatch(targetSession.className, stObj.className) || isClassMatch(className, stObj.className);
      }

      if (isActMatch && isSesMatch && isClassOk) {
        var isLate = (status === 'Late' || status === 'มาสาย');
        if (isLate) countLate++;
        else countPresent++;

        checkedList.push({
          studentId: sIdRaw,
          studentNo: stObj ? stObj.studentNo : '',
          name: stObj ? stObj.name : (attData[j][2] || 'นักศึกษา'),
          className: stObj ? stObj.className : (className !== 'all' ? className : 'ห้องเรียน'),
          status: isLate ? 'Late' : 'Present',
          statusText: isLate ? 'มาสาย' : 'มาเรียน',
          checkinTime: checkinTime,
          gpsLocation: gpsLoc,
          distanceMeters: dist,
          selfiePhoto: selfie,
          activityName: rowAct
        });
      }
    }
  }

  // Sort newest first
  checkedList.reverse();

  return {
    status: 'open',
    isOpen: true,
    session: {
      sessionId: targetSession.sessionId,
      className: targetSession.className,
      activityName: targetSession.activityName,
      date: targetSession.date,
      isAutomated: !!targetSession.isAutomated,
      phaseLabel: targetSession.phaseLabel || ''
    },
    list: checkedList,
    count: checkedList.length,
    countPresent: countPresent,
    countLate: countLate,
    total: students.length
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
function deleteLiveCheckinRecord(studentId, dateStr, activityName, sessionId) {
  var ss = getSpreadsheet();
  var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!attSheet) return { status: 'error', message: 'ไม่พบชีต Attendance' };

  var data = attSheet.getDataRange().getValues();
  var found = false;
  var targetDate = dateStr ? formatDateString(dateStr) : formatDateString(new Date());
  var cleanStudentId = String(studentId || '').trim().toLowerCase();
  var cleanAct = activityName ? String(activityName).trim().toLowerCase() : '';
  var cleanSes = sessionId ? String(sessionId).trim() : '';

  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = formatDateString(data[i][0]);
    var rowStudentId = String(data[i][1] || '').trim().toLowerCase();
    var rowAct = String(data[i][8] || 'เช็คชื่อมาเรียน').trim().toLowerCase();
    var rowSes = String(data[i][9] || '').trim();

    var isMatch = (rowStudentId === cleanStudentId && (!targetDate || rowDate === targetDate));
    if (isMatch) {
      if (cleanSes && rowSes) {
        if (cleanSes === rowSes) {
          attSheet.deleteRow(i + 1);
          found = true;
        }
      } else if (cleanAct) {
        if (isActivityMatch(cleanAct, rowAct)) {
          attSheet.deleteRow(i + 1);
          found = true;
        }
      } else {
        attSheet.deleteRow(i + 1);
        found = true;
      }
    }
  }
  SpreadsheetApp.flush();

  return {
    status: 'success',
    message: found ? 'ลบรายการเช็คชื่อเรียบร้อยแล้ว' : 'ไม่พบรายการเช็คชื่อ'
  };
}


function getRangeAttendanceSummary(startDateStr, endDateStr, className, activityName) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
    if (!sheet) return { status: 'success', recordedDates: [], students: [], totalRecordedDays: 0, summaryTotals: {} };

    var startD = startDateStr ? formatDateString(startDateStr) : formatDateString(new Date());
    var endD = endDateStr ? formatDateString(endDateStr) : startD;
    if (endD < startD) {
      var tmp = startD; startD = endD; endD = tmp;
    }

    var cleanClass = String(className || 'all').trim();
    var cleanAct = String(activityName || 'all').trim();

    var allStudents = getStudents(cleanClass) || [];
    var studentMap = {};
    var studentOrder = [];

    for (var s = 0; s < allStudents.length; s++) {
      var st = allStudents[s];
      var sId = String(st.studentId || '').trim();
      if (!sId) continue;
      studentMap[sId.toLowerCase()] = {
        studentId: sId,
        studentNo: st.studentNo || (s + 1),
        name: st.name || '',
        className: st.className || '',
        present: 0,
        late: 0,
        sick: 0,
        leave: 0,
        absent: 0,
        totalRecords: 0
      };
      studentOrder.push(sId.toLowerCase());
    }

    var datesWithDataMap = {};
    var lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var rowD = formatDateString(row[0]);
        if (rowD && rowD >= startD && rowD <= endD) {
          var rowAct = String(row[8] || 'เช็คชื่อมาเรียน').trim();
          if (!isActivityMatch(cleanAct, rowAct)) continue;

          var sId = String(row[1] || '').trim().toLowerCase();
          var stObj = studentMap[sId];
          if (!stObj) continue;

          datesWithDataMap[rowD] = true;

          var status = String(row[2] || 'Present').trim();
          if (status === 'Present' || status === 'มาเรียน') {
            stObj.present++;
          } else if (status === 'Late' || status === 'มาสาย') {
            stObj.late++;
          } else if (status === 'Sick' || status === 'ลาป่วย') {
            stObj.sick++;
          } else if (status === 'Leave' || status === 'ลากิจ') {
            stObj.leave++;
          } else {
            stObj.absent++;
          }
          stObj.totalRecords++;
        }
      }
    }

    var recordedDates = Object.keys(datesWithDataMap).sort();
    var totalDays = recordedDates.length;

    var resultList = [];
    var totPresent = 0, totLate = 0, totSick = 0, totLeave = 0, totAbsent = 0;

    for (var k = 0; k < studentOrder.length; k++) {
      var item = studentMap[studentOrder[k]];
      if (!item) continue;
      if (totalDays > item.totalRecords) {
        var unrecorded = totalDays - item.totalRecords;
        item.absent += unrecorded;
      }
      var attended = item.present + item.late;
      var effTotal = Math.max(totalDays, item.present + item.late + item.sick + item.leave + item.absent);
      var attRate = effTotal > 0 ? ((attended / effTotal) * 100).toFixed(1) : '100.0';

      var evalLabel = 'ดีเยี่ยม';
      var evalClass = 'excellent';
      var rNum = parseFloat(attRate);
      if (rNum >= 85) {
        evalLabel = 'ดีเยี่ยม';
        evalClass = 'excellent';
      } else if (rNum >= 80) {
        evalLabel = 'ปกติ';
        evalClass = 'good';
      } else if (rNum >= 60) {
        evalLabel = 'เสี่ยงติด มส.';
        evalClass = 'warning';
      } else {
        evalLabel = 'หมดสิทธิ์สอบ';
        evalClass = 'critical';
      }

      totPresent += item.present;
      totLate += item.late;
      totSick += item.sick;
      totLeave += item.leave;
      totAbsent += item.absent;

      resultList.push({
        studentId: item.studentId,
        studentNo: item.studentNo,
        name: item.name,
        className: item.className,
        present: item.present,
        late: item.late,
        sick: item.sick,
        leave: item.leave,
        absent: item.absent,
        attended: attended,
        totalDays: effTotal,
        attendanceRate: attRate,
        evaluation: evalLabel,
        evalClass: evalClass
      });
    }

    return {
      status: 'success',
      startDate: startD,
      endDate: endD,
      startDateThai: formatThaiDate(startD),
      endDateThai: formatThaiDate(endD),
      className: (cleanClass === 'all' || cleanClass === 'ทุกห้องเรียน') ? 'ทุกห้องเรียน' : ('ห้อง ' + cleanClass.replace(/^ห้อง\s*/, '')),
      activityName: (cleanAct === 'all' || !cleanAct) ? 'ทุกกิจกรรม/วิชา' : cleanAct,
      recordedDates: recordedDates,
      totalRecordedDays: totalDays,
      summaryTotals: {
        totalStudents: resultList.length,
        present: totPresent,
        late: totLate,
        sick: totSick,
        leave: totLeave,
        absent: totAbsent,
        avgRate: resultList.length > 0 && totalDays > 0 ? (((totPresent + totLate) / (resultList.length * totalDays)) * 100).toFixed(1) : '100.0'
      },
      students: resultList
    };
  } catch (err) {
    return {
      status: 'error',
      message: 'เกิดข้อผิดพลาดในการดึงข้อมูลสรุป: ' + (err.message || String(err)),
      recordedDates: [],
      students: [],
      totalRecordedDays: 0,
      summaryTotals: {}
    };
  }
}


function getAttendanceRecords(dateStr, className, activityName, reportType, endDateStr) {
  var ss = getSpreadsheet();
  var attendanceSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!attendanceSheet) return {};
  
  var startDate = dateStr ? formatDateString(dateStr) : formatDateString(new Date());
  var endDate = endDateStr ? formatDateString(endDateStr) : startDate;
  if (endDate < startDate) {
    var tempDate = startDate;
    startDate = endDate;
    endDate = tempDate;
  }
  
  var cleanAct = String(activityName || 'all').trim();
  var cleanClass = String(className || 'all').trim();
  var repType = String(reportType || 'daily').toLowerCase();
  
  var students = getStudents(cleanClass);
  var studentIds = {};
  for (var s = 0; s < students.length; s++) {
    var rawId = String(students[s].studentId || '').trim();
    if (rawId) {
      studentIds[rawId] = students[s];
      studentIds[rawId.toLowerCase()] = students[s];
    }
  }
  
  var attData = attendanceSheet.getDataRange().getValues();
  var records = {};
  
  for (var i = 1; i < attData.length; i++) {
    var rawDateVal = attData[i][0];
    var rowDate = formatDateString(rawDateVal);
    var sId = String(attData[i][1] || '').trim();
    var status = String(attData[i][2] || 'Present').trim();
    var checkinTime = formatTimeString(attData[i][3]);
    var gpsLocation = String(attData[i][4] || '-').trim();
    var distanceMeters = attData[i][5] || 0;
    var deviceId = String(attData[i][6] || '').trim();
    var selfiePhoto = String(attData[i][7] || '').trim();
    var rowAct = String(attData[i][8] || 'เช็คชื่อมาเรียน').trim();
    var rowSes = String(attData[i][9] || '').trim();
    var rowNote = String(attData[i][10] || '').trim();
    
    if (!sId) continue;

    // Date matching logic (Supports single date, date range, monthly, semester)
    var isDateMatch = false;
    if (repType === 'monthly') {
      isDateMatch = (rowDate.substring(0, 7) === startDate.substring(0, 7));
    } else if (repType === 'semester') {
      isDateMatch = true;
    } else if (startDate !== endDate || repType === 'range') {
      isDateMatch = (rowDate >= startDate && rowDate <= endDate);
    } else {
      isDateMatch = (rowDate === startDate);
    }
    
    // Match date and activity
    if (isDateMatch) {
      if (isActivityMatch(cleanAct, rowAct)) {
        var recObj = {
          status: status,
          checkinTime: checkinTime,
          gpsLocation: gpsLocation,
          distanceMeters: distanceMeters,
          deviceId: deviceId,
          selfiePhoto: selfiePhoto,
          checkinType: rowAct,
          activityName: rowAct,
          sessionId: rowSes,
          note: rowNote,
          date: rowDate
        };
        // For range queries, keep the latest or best attendance status
        if (!records[sId] || status === 'Present' || status === 'มาเรียน') {
          records[sId] = recObj;
          records[sId.toLowerCase()] = recObj;
        }
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
    var sId = String(existingData[i][1] || '').trim().toLowerCase();
    var rAct = String(existingData[i][8] || 'เช็คชื่อมาเรียน').trim().toLowerCase();
    rowMap[dStr + '_' + sId + '_' + rAct] = i + 1;
  }
  
  var nowTimeStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm');
  
  for (var k = 0; k < attendanceList.length; k++) {
    var item = attendanceList[k];
    var formattedDate = formatDateString(item.date);
    var itemAct = String(item.activityName || defaultAct).trim();
    var cleanSId = String(item.studentId || '').trim().toLowerCase();
    var key = formattedDate + '_' + cleanSId + '_' + itemAct.toLowerCase();
    var rowIndex = rowMap[key];
    var noteVal = String(item.note || '').trim();
    
    if (rowIndex) {
      sheet.getRange(rowIndex, 1).setValue(formattedDate);
      sheet.getRange(rowIndex, 3).setValue(item.status);
      sheet.getRange(rowIndex, 4).setValue(nowTimeStr);
      sheet.getRange(rowIndex, 9).setValue(itemAct);
      sheet.getRange(rowIndex, 11).setValue(noteVal);
      if (item.selfiePhoto) sheet.getRange(rowIndex, 8).setValue(item.selfiePhoto);
    } else {
      sheet.appendRow([
        formattedDate,
        String(item.studentId).trim(),
        item.status,
        nowTimeStr,
        '-',
        0,
        '-',
        item.selfiePhoto || '',
        itemAct,
        item.sessionId || '',
        noteVal
      ]);
      rowMap[key] = sheet.getLastRow();
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
    try {
      return Utilities.formatDate(val, 'Asia/Bangkok', 'yyyy-MM-dd');
    } catch (e) {
      var y = val.getFullYear();
      if (y > 2500) y = y - 543;
      var m = String(val.getMonth() + 1);
      if (m.length < 2) m = '0' + m;
      var d = String(val.getDate());
      if (d.length < 2) d = '0' + d;
      return y + '-' + m + '-' + d;
    }
  }
  
  var str = String(val).trim();
  if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
    var yPart = parseInt(str.substring(0, 4), 10);
    if (yPart > 2500) {
      return String(yPart - 543) + str.substring(4, 10);
    }
    return str.substring(0, 10);
  }
  
  if (str.indexOf('/') > -1) {
    var parts = str.split('/');
    if (parts.length === 3) {
      var day = parts[0].trim().length === 1 ? '0' + parts[0].trim() : parts[0].trim();
      var month = parts[1].trim().length === 1 ? '0' + parts[1].trim() : parts[1].trim();
      var year = parts[2].trim().substring(0, 4);
      if (parseInt(year, 10) > 2500) {
        year = String(parseInt(year, 10) - 543);
      } else if (year.length === 2) {
        year = '20' + year;
      }
      return year + '-' + month + '-' + day;
    }
  }

  try {
    var parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, 'Asia/Bangkok', 'yyyy-MM-dd');
    }
  } catch (ex) {}

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

        // ONLY reply when joining or when user explicitly asks for Group ID command (!id, /id, รหัสกลุ่ม)
        if (token && ev.replyToken) {
          var isJoin = (ev.type === 'join');
          var isMsg = (ev.type === 'message' && ev.message && ev.message.type === 'text');
          var userText = (isMsg && ev.message.text) ? String(ev.message.text).trim().toLowerCase() : '';
          var isCommand = (userText === '!id' || userText === '/id' || userText === '!groupid' || userText === 'รหัสกลุ่ม' || userText === 'groupid' || userText === '!เช็คชื่อ');

          if (isJoin || isCommand) {
            var replyText = '🎉 สวัสดีครับผู้ปกครองและอาจารย์ทุกท่าน!\n' +
                            '🤖 บอทระบบเช็คชื่อนักศึกษาได้เชื่อมต่อกับกลุ่มนี้เรียบร้อยแล้ว\n' +
                            '📌 รหัสกลุ่มของคุณคือ: ' + gId + '\n' +
                            '✨ ระบบจะส่งรายงานสรุปการมาเรียนเข้ากลุ่มนี้ตามเวลาที่คุณครูกำหนดครับ';
            
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

function testLineConnection(token, webhookUrl, lineGroupId, customMsg, sendType) {
  token = String(token || '').trim();
  webhookUrl = String(webhookUrl || '').trim();
  lineGroupId = String(lineGroupId || '').trim();
  customMsg = String(customMsg || '').trim();
  sendType = String(sendType || 'flex').trim(); // 'flex' or 'text'
  
  if (!token && !webhookUrl) {
    return { status: 'error', message: 'กรุณาระบุ LINE Channel Access Token หรือ Webhook URL' };
  }
  
  var now = new Date();
  var timeStr = (now.getHours() < 10 ? '0' + now.getHours() : now.getHours()) + ':' + 
                (now.getMinutes() < 10 ? '0' + now.getMinutes() : now.getMinutes()) + ' น.';
  var testMsg = customMsg || ('🔔 [ทดสอบระบบ] การเชื่อมต่อ LINE สำเร็จ!\n' +
                '📅 วันที่: ' + formatDateString(now) + ' (' + timeStr + ')\n' +
                '✨ ระบบเช็คชื่อนักศึกษาพร้อมส่งรายงานสรุปเข้า LINE แล้วครับ');

  var flexPayload = null;
  if (sendType === 'flex' || sendType !== 'text') {
    flexPayload = buildLineFlexAttendanceMessage({
      targetDate: formatDateString(now),
      rawDate: formatDateString(now),
      classLabel: 'ทดสอบทุกห้องเรียน (Demo)',
      actLabel: 'เช็คชื่อมาเรียน (ทดสอบระบบ)',
      totalStudents: 3,
      countPresent: 3,
      countAbsent: 0,
      countSick: 0,
      countLeave: 0,
      countLate: 0,
      attRate: '100.0',
      presentList: [
        { no: 1, name: 'นักเรียนตัวอย่าง 1 (แชมป์)', className: 'ม.4/1', time: '08:01' },
        { no: 2, name: 'นักเรียนตัวอย่าง 2', className: 'ม.4/1', time: '08:03' },
        { no: 3, name: 'นักเรียนตัวอย่าง 3', className: 'ม.4/1', time: '08:05' }
      ],
      absentList: [],
      lateList: [],
      sickList: [],
      leaveList: [],
      isMultiClass: false
    });
  }

  return dispatchLineNotification(token, testMsg, lineGroupId, webhookUrl, flexPayload);
}

function dispatchLineNotification(token, msg, lineGroupId, webhookUrl, flexPayload) {
  token = String(token || '').trim();
  lineGroupId = String(lineGroupId || '').trim();
  webhookUrl = String(webhookUrl || '').trim();

  if (!token && !webhookUrl) {
    return { status: 'error', message: 'กรุณาระบุ LINE Token หรือ Webhook URL' };
  }

  // 1. LINE Messaging API (Channel Access Token - Long token > 60 chars)
  if (token && token.length > 60) {
    var endpoint = 'https://api.line.me/v2/bot/message/broadcast';
    var isFlex = (flexPayload && flexPayload.type === 'flex');
    var messageItem = isFlex ? flexPayload : { type: 'text', text: msg };
    var payloadObj = {
      messages: [ messageItem ]
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
        var succText = isFlex ? 'ส่งการ์ด Flex Message เข้า LINE สำเร็จแล้ว! 🌟' : 'ส่งข้อความ Text เข้า LINE สำเร็จแล้ว! 📝';
        return { status: 'success', message: succText };
      } else {
        var errText = msgApiRes.getContentText();
        // If Flex Message fails with 400, attempt automatic resilient fallback to text message
        if (isFlex && msgApiCode === 400 && msg) {
          try {
            var fallbackPayload = { messages: [{ type: 'text', text: msg }] };
            if (payloadObj.to) fallbackPayload.to = payloadObj.to;
            var fbOptions = {
              method: 'post',
              contentType: 'application/json',
              headers: { 'Authorization': 'Bearer ' + token },
              payload: JSON.stringify(fallbackPayload),
              muteHttpExceptions: true
            };
            var fbRes = UrlFetchApp.fetch(endpoint, fbOptions);
            if (fbRes.getResponseCode() === 200) {
              return {
                status: 'success',
                message: 'ส่งข้อความสรุปรายงานเข้า LINE สำเร็จแล้ว! 📝'
              };
            }
          } catch (fbEx) {}
        }
        if (msgApiCode === 429 || errText.indexOf('monthly limit') > -1 || errText.indexOf('rate limit') > -1) {
          return {
            status: 'error',
            code: 429,
            message: 'โควตาส่งข้อความฟรีรายเดือนของบัญชี LINE OA นี้เต็มแล้ว (500 ข้อความ/เดือน)\n\n💡 แนะนำ: สามารถใช้ Token จาก LINE OA ใหม่, เปลี่ยนใช้ LINE Notify Token (ฟรีไม่จำกัด), หรือกดปุ่มคัดลอกข้อความไปวางในกลุ่ม LINE ได้เลยครับ'
          };
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

  // 3. Custom Webhook URL
  if (webhookUrl) {
    try {
      var whOptions = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          message: msg,
          flexMessage: flexPayload,
          timestamp: new Date().toISOString()
        }),
        muteHttpExceptions: true
      };
      var whRes = UrlFetchApp.fetch(webhookUrl, whOptions);
      var whCode = whRes.getResponseCode();
      if (whCode >= 200 && whCode < 300) {
        return { status: 'success', message: 'ส่งข้อความผ่าน Custom Webhook สำเร็จ!' };
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
    return { status: 'error', message: 'ยังไม่ได้ตั้งค่า LINE Channel Access Token หรือ Webhook URL' };
  }

  var targetDate = options.date || formatDateString(new Date());
  var targetClass = options.className || config.targetClass || 'all';
  var targetActivity = options.activityName || config.targetActivity || 'all';
  var statusFilter = options.statusFilter || config.statusFilter || 'all';
  var messageStyle = options.messageStyle || config.messageStyle || 'flex';
  
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
    var sId = String(s.studentId || '').trim();
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
    
    var sPhoto = '';
    if (rec && typeof rec === 'object' && rec.selfiePhoto) {
      sPhoto = rec.selfiePhoto;
    } else if (s.photo || s.selfiePhoto || s.avatar || s.imageUrl || s.img) {
      sPhoto = s.photo || s.selfiePhoto || s.avatar || s.imageUrl || s.img;
    }

    var sItem = {
      no: s.studentNo || (i + 1),
      name: s.name,
      studentId: s.studentId,
      className: s.className,
      status: st,
      note: note,
      time: cTime,
      selfiePhoto: sPhoto
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
  var totalAll = countPresent + countAbsent + countSick + countLeave + countLate;
  var totalBase = totalStudents > 0 ? totalStudents : (totalAll > 0 ? totalAll : 1);
  var attRate = totalBase > 0 ? (((countPresent + countLate) / totalBase) * 100).toFixed(1) : '100.0';
  var pctPresent = ((countPresent / totalBase) * 100).toFixed(1);
  var pctLate = ((countLate / totalBase) * 100).toFixed(1);
  var pctSickLeave = (((countSick + countLeave) / totalBase) * 100).toFixed(1);
  var pctAbsent = ((countAbsent / totalBase) * 100).toFixed(1);
  
  function makeProgressBar(rate) {
    var num = parseFloat(rate) || 0;
    var blocks = Math.round(num / 10);
    if (blocks > 10) blocks = 10;
    if (blocks < 0) blocks = 0;
    var bar = '';
    for (var b = 0; b < blocks; b++) bar += '🟩';
    for (var w = blocks; w < 10; w++) bar += '⬜';
    return bar;
  }

  var isMultiClass = (targetClass === 'all' || targetClass === 'ทุกห้องเรียน' || String(targetClass).indexOf(',') > -1);
  var classLabel = (targetClass === 'all' || targetClass === 'ทุกห้องเรียน') ? 'ทุกห้องเรียน' : ('ห้อง ' + targetClass);
  var actLabel = (targetActivity === 'all' || targetActivity === 'ทุกกิจกรรม' || !targetActivity) ? 'เช็คชื่อมาเรียน' : targetActivity;
  
  // Determine if sending as Flex Card or Text-only
  var sendType = options.sendType || options.format || (options.useFlex === false ? 'text' : (options.messageStyle || config.messageStyle || 'flex'));
  var isTextOnly = (sendType === 'text' || sendType === 'text_only' || options.useFlex === false || (sendType !== 'flex' && !options.useFlex));

  // Build beautiful LINE Flex Message if not text-only
  var flexPayload = null;
  if (!isTextOnly) {
    flexPayload = buildLineFlexAttendanceMessage({
      targetDate: targetDate,
      rawDate: targetDate,
      classLabel: classLabel,
      actLabel: actLabel,
      totalStudents: totalStudents,
      countPresent: countPresent,
      countAbsent: countAbsent,
      countSick: countSick,
      countLeave: countLeave,
      countLate: countLate,
      attRate: attRate,
      presentList: presentList,
      absentList: absentList,
      lateList: lateList,
      sickList: sickList,
      leaveList: leaveList,
      isMultiClass: isMultiClass
    });
  }

  // Format standard text message as fallback
  var divLine = '━━━━━━━━━━━━━━━━━━━━\n';
  var msg = divLine +
            '👨‍👩‍👧‍👦 แจ้งเตือนสถิติการมาเรียน (กลุ่มผู้ปกครอง)\n' +
            '🏫 ห้องเรียน: ' + classLabel + '\n' +
            '📅 ประจำวันที่: ' + targetDate + '\n' +
            '🏷️ รายการ: ' + actLabel + '\n' +
            divLine +
            '📊 อัตราการเข้าเรียน: ' + attRate + '%\n' +
            '[' + makeProgressBar(attRate) + ']\n\n' +
            '📋 สรุปยอดการมาเรียนวันนี้:\n' +
            '👥 นักเรียนทั้งหมด: ' + totalStudents + ' คน (100%)\n' +
            '🟢 มาเรียนตรงเวลา: ' + countPresent + ' คน (' + pctPresent + '%)\n' +
            '🌸 มาสาย: ' + countLate + ' คน (' + pctLate + '%)\n' +
            '🟡 ลาป่วย/ลากิจ: ' + (countSick + countLeave) + ' คน (' + pctSickLeave + '%)\n' +
            '🔴 ขาดเรียน (ไม่พบเช็คชื่อ): ' + countAbsent + ' คน (' + pctAbsent + '%)\n' +
            divLine;

  var textEarlyBirds = [];
  for (var teb = 0; teb < presentList.length; teb++) {
    if (presentList[teb].time && presentList[teb].time !== '-') {
      textEarlyBirds.push(presentList[teb]);
    }
  }
  textEarlyBirds.sort(function(a, b) {
    return String(a.time).localeCompare(String(b.time));
  });

  if (textEarlyBirds.length > 0) {
    msg += '\n🏆 อันดับเช็คชื่อไวที่สุดประจำวัน:\n';
    if (textEarlyBirds[0]) msg += '🥇 อันดับ 1: 👑 ' + textEarlyBirds[0].name + ' ⏰ ' + textEarlyBirds[0].time + ' น.\n';
    if (textEarlyBirds[1]) msg += '🥈 อันดับ 2: 👑 ' + textEarlyBirds[1].name + ' ⏰ ' + textEarlyBirds[1].time + ' น.\n';
    if (textEarlyBirds[2]) msg += '🥉 อันดับ 3: 👑 ' + textEarlyBirds[2].name + ' ⏰ ' + textEarlyBirds[2].time + ' น.\n';
  }

  if (textEarlyBirds.length > 3) {
    msg += '\n🟢 รายชื่อผู้ที่มาเรียนตรงเวลา (ลำดับที่ 4 เป็นต้นไป - ' + (textEarlyBirds.length - 3) + ' คน):\n';
    for (var pPr = 3; pPr < textEarlyBirds.length; pPr++) {
      var rTagPr = isMultiClass ? ' (' + textEarlyBirds[pPr].className + ')' : '';
      var tPr = textEarlyBirds[pPr].time;
      var tStrPr = (tPr && tPr !== '-') ? ' ⏰ เข้า ' + tPr + ' น.' : '';
      msg += '  ' + (pPr + 1) + '. เลขที่ ' + textEarlyBirds[pPr].no + ' ' + textEarlyBirds[pPr].name + rTagPr + tStrPr + '\n';
    }
  } else if (textEarlyBirds.length > 0) {
    msg += '\n🟢 ผู้ที่มาเรียนตรงเวลา: ทั้งหมดอยู่ในโพเดียม 3 อันดับแรกข้างต้น ✨\n';
  }

  if (absentList.length > 0) {
    msg += '\n🔴 รายชื่อนักเรียนที่ไม่พบการเช็คชื่อ (' + absentList.length + ' คน):\n';
    for (var pAb = 0; pAb < absentList.length; pAb++) {
      var rTagAb = isMultiClass ? ' (' + absentList[pAb].className + ')' : '';
      msg += '  ' + (pAb + 1) + '. เลขที่ ' + absentList[pAb].no + ' ' + absentList[pAb].name + rTagAb + (absentList[pAb].note ? ' 📝 ' + absentList[pAb].note : '') + '\n';
    }
  }

  if (lateList.length > 0) {
    msg += '\n🌸 รายชื่อนักเรียนที่มาสาย (' + lateList.length + ' คน):\n';
    for (var pLt = 0; pLt < lateList.length; pLt++) {
      var rTagLt = isMultiClass ? ' (' + lateList[pLt].className + ')' : '';
      msg += '  ' + (pLt + 1) + '. เลขที่ ' + lateList[pLt].no + ' ' + lateList[pLt].name + rTagLt + ' ⏰ เข้าเรียน ' + formatTimeString(lateList[pLt].time) + ' น.\n';
    }
  }

  if (sickList.length > 0 || leaveList.length > 0) {
    msg += '\n🟡 รายชื่อนักเรียนที่แจ้งลา (' + (sickList.length + leaveList.length) + ' คน):\n';
    var lCount = 1;
    for (var pSk = 0; pSk < sickList.length; pSk++) {
      var rTagSk = isMultiClass ? ' (' + sickList[pSk].className + ')' : '';
      msg += '  ' + lCount + '. [ลาป่วย] เลขที่ ' + sickList[pSk].no + ' ' + sickList[pSk].name + rTagSk + (sickList[pSk].note ? ' 📝 ' + sickList[pSk].note : '') + '\n';
      lCount++;
    }
    for (var pLv = 0; pLv < leaveList.length; pLv++) {
      var rTagLv = isMultiClass ? ' (' + leaveList[pLv].className + ')' : '';
      msg += '  ' + lCount + '. [ลากิจ] เลขที่ ' + leaveList[pLv].no + ' ' + leaveList[pLv].name + rTagLv + (leaveList[pLv].note ? ' 📝 ' + leaveList[pLv].note : '') + '\n';
      lCount++;
    }
  }

  if (countAbsent === 0 && countLate === 0 && countSick === 0 && countLeave === 0) {
    msg += '\n✨ ยอดเยี่ยมมากครับ! วันนี้นักเรียนทุกคนมาเรียนครบตรงเวลา 100% 🎉\n';
  }

  var parentReportUrl = '';
  try {
    var scriptUrl = ScriptApp.getService().getUrl();
    if (scriptUrl && scriptUrl.indexOf('http') === 0) {
      parentReportUrl = scriptUrl + '?page=report&date=' + encodeURIComponent(targetDate) + '&class=' + encodeURIComponent(targetClass) + '&activity=' + encodeURIComponent(targetActivity);
    }
  } catch (ex) {}

  if (parentReportUrl) {
    msg += '\n📱 ลิงก์ดูรายงานฉบับเต็มสำหรับผู้ปกครอง:\n' + parentReportUrl + '\n';
  }

  msg += '\n' + divLine +
         '💬 หากผู้ปกครองท่านใดประสงค์แจ้งการลา หรือมีข้อสอบถาม สามารถติดต่ออาจารย์ที่ปรึกษาได้โดยตรงครับ\n' +
         '🙏 ขอบพระคุณผู้ปกครองทุกท่านที่ให้ความร่วมมือครับ';

  var finalTextMessage = (options.customMessage && String(options.customMessage).trim().length > 0) ? String(options.customMessage).trim() : msg;
  return dispatchLineNotification(token, finalTextMessage, lineGroupId, webhookUrl, isTextOnly ? null : flexPayload);
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


/**
 * =========================================================================
 * AUTOMATED DAILY CHECK-IN SCHEDULE & CALENDAR ENGINE
 * =========================================================================
 */
function getAutoCheckinSchedules() {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('AUTO_CHECKIN_SCHEDULES');
    if (raw !== null && raw !== undefined) {
      var list = JSON.parse(raw);
      if (Array.isArray(list)) return list;
    }
  } catch (e) {
    Logger.log('Error in getAutoCheckinSchedules: ' + e);
  }

  // Default Preset Schedule (Mon-Fri 07:00 - 08:30 On-time, 08:30-09:00 Late)
  var defaultSchedule = [
    {
      id: 'SCH_DEFAULT_ALL',
      className: 'all',
      title: 'ตารางเช็คชื่อประจำวันมาตรฐาน (ทุกห้องเรียน)',
      enabled: true,
      days: [1, 2, 3, 4, 5],
      activityName: 'เช็คชื่อมาเรียน',
      morningStart: '07:00',
      morningOnTimeEnd: '08:30',
      morningLateEnd: '09:00',
      morningCutoff: '11:59',
      afternoonStart: '12:00',
      afternoonOnTimeEnd: '13:30',
      afternoonLateEnd: '14:00',
      afternoonCutoff: '17:00',
      afternoonEnabled: true,
      requireOTP: true,
      otpIntervalSeconds: 15,
      requireDeviceBinding: true,
      requireGPS: false,
      requireSelfie: true,
      gpsLat: 0,
      gpsLng: 0,
      radiusMeters: 100
    }
  ];

  try {
    var props2 = PropertiesService.getScriptProperties();
    props2.setProperty('AUTO_CHECKIN_SCHEDULES', JSON.stringify(defaultSchedule));
  } catch (e2) {}

  return defaultSchedule;
}

function saveAutoCheckinSchedules(schedules) {
  try {
    schedules = schedules || [];
    if (typeof schedules === 'string') {
      try { schedules = JSON.parse(schedules); } catch (e) {}
    }
    var props = PropertiesService.getScriptProperties();
    props.setProperty('AUTO_CHECKIN_SCHEDULES', JSON.stringify(schedules));

    // Also mirror to Google Sheet if sheet exists
    try {
      var ss = getSpreadsheet();
      var schSheet = ss.getSheetByName(SHEET_SCHEDULE);
      if (schSheet) {
        schSheet.clear();
        schSheet.appendRow(['ID', 'ClassName', 'Days', 'ActivityName', 'MorningStart', 'MorningOnTimeEnd', 'MorningLateEnd', 'MorningCutoff', 'AfternoonStart', 'AfternoonOnTimeEnd', 'AfternoonLateEnd', 'AfternoonCutoff', 'Enabled', 'RequireGPS', 'RequireSelfie', 'UpdatedAt']);
        for (var i = 0; i < schedules.length; i++) {
          var s = schedules[i];
          schSheet.appendRow([
            s.id || ('SCH_' + (i + 1)),
            s.className || 'all',
            Array.isArray(s.days) ? s.days.join(',') : String(s.days || '1,2,3,4,5'),
            s.activityName || 'เช็คชื่อมาเรียน',
            s.morningStart || '07:00',
            s.morningOnTimeEnd || '08:30',
            s.morningLateEnd || '09:00',
            s.morningCutoff || '11:59',
            s.afternoonStart || '12:00',
            s.afternoonOnTimeEnd || '13:30',
            s.afternoonLateEnd || '14:00',
            s.afternoonCutoff || '17:00',
            s.enabled !== false ? 'YES' : 'NO',
            s.requireGPS === true ? 'YES' : 'NO',
            s.requireSelfie !== false ? 'YES' : 'NO',
            new Date().toISOString()
          ]);
        }
      }
    } catch (sheetErr) {
      Logger.log('Could not mirror to Schedule sheet: ' + sheetErr);
    }

    return {
      status: 'success',
      message: 'บันทึกการตั้งค่าตารางเช็คชื่อรายวันอัตโนมัติเรียบร้อยแล้ว',
      schedules: schedules
    };
  } catch (e) {
    return {
      status: 'error',
      message: 'เกิดข้อผิดพลาดในการบันทึกตารางเวลา: ' + e.message
    };
  }
}

function resolveActiveSchedule(className, clientDateStr, clientTimeStr) {
  var schedules = getAutoCheckinSchedules();
  var now = new Date();
  var nowMinutes = now.getHours() * 60 + now.getMinutes();
  var dayOfWeek = now.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat

  className = String(className || 'all').trim();

  // If client provided a specific time (e.g. "08:15")
  if (clientTimeStr && String(clientTimeStr).indexOf(':') > -1) {
    var parts = String(clientTimeStr).split(':');
    var h = parseInt(parts[0], 10);
    var m = parseInt(parts[1], 10);
    if (!isNaN(h) && !isNaN(m)) {
      nowMinutes = h * 60 + m;
    }
  }

  function parseMin(tStr, defaultMin) {
    if (!tStr || String(tStr).indexOf(':') === -1) return defaultMin;
    var p = String(tStr).split(':');
    return (parseInt(p[0], 10) * 60) + parseInt(p[1], 10);
  }

  // Find matching active schedule for this class
  var matchedSchedule = null;
  for (var i = 0; i < schedules.length; i++) {
    var sch = schedules[i];
    if (sch.enabled === false) continue;

    // Check Class Match
    var isMatch = (sch.className === 'all' || sch.className === 'ทุกห้องเรียน' || sch.className === className);
    if (!isMatch && sch.classes && Array.isArray(sch.classes)) {
      isMatch = (sch.classes.indexOf(className) > -1 || sch.classes.indexOf('all') > -1);
    }
    if (!isMatch && typeof sch.className === 'string' && sch.className.indexOf(',') > -1) {
      var cSplits = sch.className.split(',');
      for (var cs = 0; cs < cSplits.length; cs++) {
        if (cSplits[cs].trim() === className || cSplits[cs].trim() === 'all') {
          isMatch = true;
          break;
        }
      }
    }
    if (!isMatch) continue;

    // Check Day Match
    var isDayMatch = false;
    if (sch.days === 'all' || !sch.days) {
      isDayMatch = true;
    } else if (Array.isArray(sch.days)) {
      isDayMatch = (sch.days.indexOf(dayOfWeek) > -1 || sch.days.indexOf(String(dayOfWeek)) > -1);
    } else if (typeof sch.days === 'string') {
      var dArr = sch.days.split(',');
      isDayMatch = (dArr.indexOf(String(dayOfWeek)) > -1 || dArr.indexOf('all') > -1);
    }
    if (!isDayMatch) continue;

    matchedSchedule = sch;
    break;
  }

  if (!matchedSchedule) {
    return {
      status: 'locked',
      code: 'NOT_SCHEDULED',
      message: 'ห้องนี้ไม่ได้เปิดการตั้งค่าเช็คชื่ออัตโนมัติ หรือไม่มีตารางสำหรับวันนี้',
      className: className,
      schedule: null
    };
  }

  var mStart = parseMin(matchedSchedule.morningStart, 7 * 60);         // 07:00 (420)
  var mOnTime = parseMin(matchedSchedule.morningOnTimeEnd, 8 * 60 + 30); // 08:30 (510)
  var mLate = parseMin(matchedSchedule.morningLateEnd, 9 * 60);          // 09:00 (540)
  var mCutoff = parseMin(matchedSchedule.morningCutoff, 11 * 60 + 59);  // 11:59 (719)

  var aStart = parseMin(matchedSchedule.afternoonStart, 12 * 60);        // 12:00 (720)
  var aOnTime = parseMin(matchedSchedule.afternoonOnTimeEnd, 13 * 60 + 30); // 13:30 (810)
  var aLate = parseMin(matchedSchedule.afternoonLateEnd, 14 * 60);       // 14:00 (840)
  var aCutoff = parseMin(matchedSchedule.afternoonCutoff, 17 * 60);     // 17:00 (1020)

  var sessionPhase = null;
  var statusResult = 'Present';
  var phaseLabel = '';
  var activeActivity = matchedSchedule.activityName || 'เช็คชื่อมาเรียน';

  if (nowMinutes >= mStart && nowMinutes <= mCutoff) {
    // MORNING SESSION
    if (nowMinutes <= mOnTime) {
      sessionPhase = 'MORNING_ON_TIME';
      statusResult = 'Present';
      phaseLabel = 'มาเรียนตรงเวลา (07:00 - 08:30 น.)';
    } else if (nowMinutes <= mLate) {
      sessionPhase = 'MORNING_LATE';
      statusResult = 'Late';
      phaseLabel = 'มาสาย (08:30 - 09:00 น.)';
    } else {
      sessionPhase = 'MORNING_LATE_EXTENDED';
      statusResult = 'Late';
      phaseLabel = 'มาสาย (หลัง 09:00 น.)';
    }
  } else if (matchedSchedule.afternoonEnabled !== false && nowMinutes >= aStart && nowMinutes <= aCutoff) {
    // AFTERNOON SESSION
    activeActivity = 'เช็คชื่อช่วงบ่าย';
    if (nowMinutes <= aOnTime) {
      sessionPhase = 'AFTERNOON_ON_TIME';
      statusResult = 'Present';
      phaseLabel = 'เช็คชื่อช่วงบ่าย (ตรงเวลา 12:00 - 13:30 น.)';
    } else {
      sessionPhase = 'AFTERNOON_LATE';
      statusResult = 'Late';
      phaseLabel = 'ช่วงบ่าย (มาสาย)';
    }
  }

  if (!sessionPhase) {
    return {
      status: 'locked',
      code: 'OUT_OF_HOURS',
      message: 'ขณะนี้อยู่นอกช่วงเวลาเช็คชื่อที่กำหนด',
      currentMinutes: nowMinutes,
      schedule: matchedSchedule,
      allowedTimes: 'รอบเช้า ' + (matchedSchedule.morningStart || '07:00') + ' - ' + (matchedSchedule.morningLateEnd || '09:00') + ' น. / รอบบ่าย ' + (matchedSchedule.afternoonStart || '12:00') + ' - ' + (matchedSchedule.afternoonCutoff || '17:00') + ' น.'
    };
  }

  var todayDateStr = clientDateStr || formatDateString(now);
  var cleanClassTag = className.replace(/[^a-zA-Z0-9]/g, '_');
  var dailySeed = 'SEC_' + todayDateStr.replace(/-/g, '') + '_' + cleanClassTag;
  var sessionId = 'AUTOSES_' + todayDateStr.replace(/-/g, '') + '_' + (nowMinutes < aStart ? 'MORN' : 'AFT');

  return {
    status: 'open',
    sessionPhase: sessionPhase,
    statusResult: statusResult,
    phaseLabel: phaseLabel,
    activityName: activeActivity,
    className: className,
    date: todayDateStr,
    dateStr: todayDateStr,
    sessionId: sessionId,
    secretSeed: dailySeed,
    otpIntervalSeconds: matchedSchedule.otpIntervalSeconds || 15,
    requireOTP: matchedSchedule.requireOTP !== false,
    requireSelfie: matchedSchedule.requireSelfie !== false,
    requireGPS: matchedSchedule.requireGPS === true,
    requireDeviceBinding: matchedSchedule.requireDeviceBinding !== false,
    gpsLat: matchedSchedule.gpsLat || 0,
    gpsLng: matchedSchedule.gpsLng || 0,
    radiusMeters: matchedSchedule.radiusMeters || 100,
    schedule: matchedSchedule,
    serverTimeMs: now.getTime()
  };
}


/**
 * Saves or updates student reference face photo (Base64 data URI or image URL)
 */
function saveStudentFacePhoto(studentId, photoBase64) {
  studentId = String(studentId || '').trim();
  if (!studentId) {
    return { status: 'error', message: 'ไม่พบรหัสนักศึกษา' };
  }
  if (!photoBase64 || photoBase64.length < 15) {
    return { status: 'error', message: 'รูปภาพไม่ถูกต้อง' };
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) {
    initDatabase();
    sheet = ss.getSheetByName(SHEET_STUDENTS);
  }

  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === studentId.toLowerCase()) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow === -1) {
    return { status: 'error', message: 'ไม่พบข้อมูลนักศึกษารหัส: ' + studentId };
  }

  // Ensure header in column 6 (Photo) exists
  if (!sheet.getRange(1, 6).getValue()) {
    sheet.getRange(1, 6).setValue('Photo');
  }

  // Column 6 is Photo
  sheet.getRange(foundRow, 6).setValue(photoBase64);
  SpreadsheetApp.flush();
  clearScriptCache();

  return {
    status: 'success',
    message: 'บันทึกใบหน้าหลักของนักศึกษาเรียบร้อยแล้ว',
    studentId: studentId,
    photo: photoBase64
  };
}
