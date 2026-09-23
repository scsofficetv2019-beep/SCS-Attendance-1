/**
 * ==============================================================================
 * Google Apps Script Backend for SCS Attendance System
 * ==============================================================================
 * 
 * 📌 সহজে ডেপ্লয় করার নিয়মাবলী (Easy Deployment Steps):
 * 1. আপনার Google Drive-এ যান অথবা ব্রাউজারে sheets.new লিখে একটি নতুন Google Spreadsheet তৈরি করুন।
 * 2. স্প্রেডশিটের ওপরের মেনু থেকে Extensions > Apps Script (এক্সটেনশন > অ্যাপস স্ক্রিপ্ট) এ ক্লিক করুন।
 * 3. Apps Script এডিটরে বিদ্যমান কোড মুছে ফেলে এই "Code.gs" ফাইলের সম্পূর্ণ কোড পেস্ট করুন।
 * 4. বামপাশের Files তালিকার পাশে "+" (Plus) আইকনে ক্লিক করে "HTML" নির্বাচন করুন।
 *    - ফাইলের নাম দিন: Index (শুধু Index লিখবেন, .html লিখতে হবে না)।
 *    - এবার app/Index.html (বা apps-script/Index.html) ফাইলের সম্পূর্ণ কোড কপি করে Index ফাইলে পেস্ট করে দিন।
 * 5. ওপরের ডানে "Deploy" (ডিপ্লয়) বাটনে ক্লিক করে "New deployment" নির্বাচন করুন।
 *    - বামপাশে গিয়ার (Gear) আইকন থেকে "Web app" সিলেক্ট করুন।
 *    - Description: SCS Attendance System
 *    - Execute as: Me (আমার অ্যাকাউন্টে)
 *    - Who has access: Anyone (যে কেউ)
 * 6. "Deploy" বাটনে ক্লিক করুন, পারমিশন চাইলে "Authorize access" দিয়ে অনুমোদন দিন।
 * 7. প্রাপ্ত Web App URL-এ প্রবেশ করলেই আপনার পুরো SCS Attendance অ্যাপটি লাইভ হয়ে যাবে!
 * 
 * ==============================================================================
 */

function doGet(e) {
  try {
    var fileNames = ['Index', 'index', 'Index.html', 'index.html', 'App', 'app'];
    var htmlOutput = null;
    
    for (var i = 0; i < fileNames.length; i++) {
      try {
        htmlOutput = HtmlService.createHtmlOutputFromFile(fileNames[i]);
        if (htmlOutput) break;
      } catch (err) {}
    }
    
    if (!htmlOutput) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:\'Hind Siliguri\',sans-serif,Arial;padding:30px;color:#991b1b;background:#fee2e2;border:2px solid #ef4444;border-radius:16px;max-width:650px;margin:40px auto;line-height:1.7;">' +
        '<h2 style="color:#b91c1c;margin-top:0;font-size:22px;">⚠️ Index.html ফাইলটি পাওয়া যায়নি!</h2>' +
        '<p style="font-size:15px;">Google Apps Script এডিটরে বামপাশে <b>+</b> আইকনে ক্লিক করুন, <b>HTML</b> নির্বাচন করুন এবং ফাইলের নাম দিন <b>Index</b> (হুবহু Index লিখুন)।</p>' +
        '<p style="font-size:15px;">এরপর প্রজেক্টের <b>Index.html</b> ফাইল থেকে সম্পূর্ণ কোডটি কপি করে পেস্ট করে Save (সংরক্ষণ) করুন।</p>' +
        '<hr style="border:none;border-top:1px solid #fca5a5;margin:20px 0;">' +
        '<p style="font-size:13px;color:#7f1d1d;"><b>সহজ নির্দেশিকা:</b> Code.gs ফাইলে দেওয়া কোডটি কোড এডিটরে পেস্ট করুন এবং HTML ফাইলে Index.html এর কোড পেস্ট করে Deploy > New deployment > Web app হিসেবে পাবলিশ করুন।</p>' +
        '</div>'
      ).setTitle('SCS Attendance - সেটআপ নির্দেশিকা');
    }
    
    return htmlOutput
      .setTitle('SCS ATTENDANCE')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (fatalErr) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:30px;color:#991b1b;background:#fee2e2;border:2px solid #ef4444;border-radius:12px;max-width:600px;margin:40px auto;">' +
      '<h2>doGet Execution Error</h2>' +
      '<p>' + fatalErr.toString() + '</p>' +
      '</div>'
    );
  }
}

/**
 * Handle optional HTTP POST JSON requests
 */
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    var action = data.action || (e && e.parameter ? e.parameter.action : '');
    var result = { success: false, error: 'Unknown action: ' + action };

    if (action === 'getInitialData') {
      result = getInitialData();
    } else if (action === 'saveAttendanceRecord') {
      result = saveAttendanceRecord(data.record);
    } else if (action === 'updateAttendanceList') {
      result = updateAttendanceList(data.records);
    } else if (action === 'updateEmployeeList') {
      result = updateEmployeeList(data.employees);
    } else if (action === 'updateOfficeSettings') {
      result = updateOfficeSettings(data.officeSetup);
    } else if (action === 'updateDriveSettings') {
      result = updateDriveSettings(data.driveSetup);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getActiveSpreadsheetSafe() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    return null;
  }
}

/**
 * Helper to safely format dates and times
 */
function formatCellDateSafe(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || 'Asia/Dhaka', 'yyyy-MM-dd');
  }
  return String(val).trim();
}

function formatCellTimeSafe(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || 'Asia/Dhaka', 'hh:mm a');
  }
  return String(val).trim();
}

/**
 * Get Initial Data from Google Sheets
 * If sheets do not exist yet, creates them with proper structure and default data.
 */
function getInitialData() {
  try {
    var ss = getActiveSpreadsheetSafe();
    if (!ss) {
      return {
        success: true,
        standalone: true,
        employees: [],
        attendance: [],
        officeSetup: {
          officeInTime: '09:00',
          officeOutTime: '17:00',
          maxLateMinutes: '15',
          holidays: []
        },
        driveSetup: { folderId: '', folderUrl: '' }
      };
    }

    // 1. Employees Sheet
    var empSheet = ss.getSheetByName('Employees');
    if (!empSheet) {
      empSheet = ss.insertSheet('Employees');
      empSheet.appendRow(['ID', 'EmpID', 'Name', 'Designation', 'Department', 'Mobile', 'PhotoUrl', 'SortOrder']);
      
      var defaultEmps = [
        ['emp-001', 'SCS-101', 'মোঃ রফিকুল ইসলাম', 'অফিস ইনচার্জ (Supervisor)', 'Operations', '01711000001', 'https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=256&h=256&fit=crop&crop=face', 1],
        ['emp-002', 'SCS-102', 'তানজিলা আক্তার', 'সিনিয়র অ্যাকাউন্ট্যান্ট', 'Accounts & Finance', '01711000002', 'https://images.unsplash.com/photo-1494790108377-be9c29b29330?w=256&h=256&fit=crop&crop=face', 2],
        ['emp-003', 'SCS-103', 'কামরুল হাসান', 'ফিল্ড এক্সিকিউটিভ', 'Field Operations', '01711000003', 'https://images.unsplash.com/photo-1500648767791-00dcc994a43e?w=256&h=256&fit=crop&crop=face', 3],
        ['emp-004', 'SCS-104', 'সাদিয়া রহমান', 'অ্যাসিস্ট্যান্ট কো-অর্ডিনেটর', 'Coordination', '01711000004', 'https://images.unsplash.com/photo-1534528741775-53994a69daeb?w=256&h=256&fit=crop&crop=face', 4],
        ['emp-005', 'SCS-105', 'মাহমুদুল হক', 'সাপোর্ট এক্সিকিউটিভ', 'Support', '01711000005', 'https://images.unsplash.com/photo-1472099645785-5658abf4ff4e?w=256&h=256&fit=crop&crop=face', 5]
      ];
      defaultEmps.forEach(function(row) {
        empSheet.appendRow(row);
      });
    }

    var empData = empSheet.getDataRange().getValues();
    var employees = [];
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][0]) {
        employees.push({
          id: String(empData[i][0]),
          empId: String(empData[i][1]),
          name: String(empData[i][2]),
          designation: String(empData[i][3]),
          department: String(empData[i][4]),
          mobile: String(empData[i][5]),
          photoUrl: String(empData[i][6]),
          sortOrder: Number(empData[i][7]) || i
        });
      }
    }

    // 2. Attendance Sheet
    var attSheet = ss.getSheetByName('Attendance');
    if (!attSheet) {
      attSheet = ss.insertSheet('Attendance');
      attSheet.appendRow(['ID', 'EmpID', 'Date', 'Type', 'Time', 'Lat', 'Lng', 'LocationText', 'Note', 'CreatedBy']);
    }
    var attData = attSheet.getDataRange().getValues();
    var attendance = [];
    for (var j = 1; j < attData.length; j++) {
      if (attData[j][0]) {
        attendance.push({
          id: String(attData[j][0]),
          empId: String(attData[j][1]),
          date: formatCellDateSafe(attData[j][2]),
          type: String(attData[j][3]),
          time: formatCellTimeSafe(attData[j][4]),
          lat: Number(attData[j][5]) || 0,
          lng: Number(attData[j][6]) || 0,
          locationText: String(attData[j][7] || ''),
          note: String(attData[j][8] || ''),
          createdBy: String(attData[j][9] || 'user')
        });
      }
    }

    // 3. Settings Sheet
    var settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('Settings');
      settingsSheet.appendRow(['Key', 'Value']);
      settingsSheet.appendRow(['officeInTime', '09:00']);
      settingsSheet.appendRow(['officeOutTime', '17:00']);
      settingsSheet.appendRow(['maxLateMinutes', '15']);
      settingsSheet.appendRow(['holidays', '2026-09-04 (Friday),2026-09-11 (Friday),2026-09-18 (Friday),2026-09-25 (Friday),2026-09-16 (Eid-e-Miladunnabi)']);
      settingsSheet.appendRow(['driveFolderId', '1AbCdEfGhIjKlMnOpQrStUvWxYz']);
      settingsSheet.appendRow(['driveFolderUrl', 'https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz']);
    }
    var setRows = settingsSheet.getDataRange().getValues();
    var settings = {};
    for (var k = 1; k < setRows.length; k++) {
      if (setRows[k][0]) {
        settings[String(setRows[k][0])] = String(setRows[k][1]);
      }
    }

    return {
      success: true,
      employees: employees,
      attendance: attendance,
      officeSetup: {
        officeInTime: settings.officeInTime || '09:00',
        officeOutTime: settings.officeOutTime || '17:00',
        maxLateMinutes: settings.maxLateMinutes || '15',
        holidays: settings.holidays ? settings.holidays.split(',') : []
      },
      driveSetup: {
        folderId: settings.driveFolderId || '',
        folderUrl: settings.driveFolderUrl || ''
      }
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * Save single Attendance Record to Google Sheet
 */
function saveAttendanceRecord(record) {
  try {
    var ss = getActiveSpreadsheetSafe();
    if (!ss) return { status: 'error', message: 'No active Google Sheet linked' };
    var sheet = ss.getSheetByName('Attendance');
    if (!sheet) {
      sheet = ss.insertSheet('Attendance');
      sheet.appendRow(['ID', 'EmpID', 'Date', 'Type', 'Time', 'Lat', 'Lng', 'LocationText', 'Note', 'CreatedBy']);
    }
    sheet.appendRow([
      record.id,
      record.empId,
      record.date,
      record.type,
      record.time,
      record.lat || 0,
      record.lng || 0,
      record.locationText || '',
      record.note || '',
      record.createdBy || 'user'
    ]);
    return { status: 'success' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Update Full Attendance List in Google Sheet (on admin edit, delete, or bulk save)
 */
function updateAttendanceList(attendanceRecords) {
  try {
    var ss = getActiveSpreadsheetSafe();
    if (!ss) return { status: 'error', message: 'No active Google Sheet linked' };
    var sheet = ss.getSheetByName('Attendance');
    if (!sheet) {
      sheet = ss.insertSheet('Attendance');
    }
    sheet.clear();
    sheet.appendRow(['ID', 'EmpID', 'Date', 'Type', 'Time', 'Lat', 'Lng', 'LocationText', 'Note', 'CreatedBy']);
    if (attendanceRecords && attendanceRecords.length > 0) {
      var rows = [];
      for (var i = 0; i < attendanceRecords.length; i++) {
        var r = attendanceRecords[i];
        rows.push([
          r.id,
          r.empId,
          r.date,
          r.type,
          r.time,
          r.lat || 0,
          r.lng || 0,
          r.locationText || '',
          r.note || '',
          r.createdBy || 'user'
        ]);
      }
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { status: 'success' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Update Employee List in Google Sheet
 */
function updateEmployeeList(employees) {
  try {
    var ss = getActiveSpreadsheetSafe();
    if (!ss) return { status: 'error', message: 'No active Google Sheet linked' };
    var sheet = ss.getSheetByName('Employees');
    if (!sheet) {
      sheet = ss.insertSheet('Employees');
    }
    sheet.clear();
    sheet.appendRow(['ID', 'EmpID', 'Name', 'Designation', 'Department', 'Mobile', 'PhotoUrl', 'SortOrder']);
    if (employees && employees.length > 0) {
      var rows = [];
      for (var i = 0; i < employees.length; i++) {
        var e = employees[i];
        rows.push([
          e.id,
          e.empId,
          e.name,
          e.designation || '',
          e.department || '',
          e.mobile || '',
          e.photoUrl || '',
          i + 1
        ]);
      }
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { status: 'success' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Update Office Settings in Google Sheet
 */
function updateOfficeSettings(officeSetup) {
  try {
    var ss = getActiveSpreadsheetSafe();
    if (!ss) return { status: 'error', message: 'No active Google Sheet linked' };
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) {
      sheet = ss.insertSheet('Settings');
    }
    
    // Read current settings to keep other keys intact
    var currentRows = sheet.getDataRange().getValues();
    var config = {};
    for (var i = 1; i < currentRows.length; i++) {
      if (currentRows[i][0]) config[String(currentRows[i][0])] = String(currentRows[i][1]);
    }
    
    config['officeInTime'] = officeSetup.officeInTime || '09:00';
    config['officeOutTime'] = officeSetup.officeOutTime || '17:00';
    config['maxLateMinutes'] = officeSetup.maxLateMinutes || '15';
    config['holidays'] = Array.isArray(officeSetup.holidays) ? officeSetup.holidays.join(',') : (officeSetup.holidays || '');
    
    sheet.clear();
    sheet.appendRow(['Key', 'Value']);
    var keys = Object.keys(config);
    if (keys.length > 0) {
      var rows = [];
      for (var k = 0; k < keys.length; k++) {
        rows.push([keys[k], config[keys[k]]]);
      }
      sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }
    return { status: 'success' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Update Drive Settings in Google Sheet
 */
function updateDriveSettings(driveSetup) {
  try {
    var ss = getActiveSpreadsheetSafe();
    if (!ss) return { status: 'error', message: 'No active Google Sheet linked' };
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) {
      sheet = ss.insertSheet('Settings');
    }
    var currentRows = sheet.getDataRange().getValues();
    var config = {};
    for (var i = 1; i < currentRows.length; i++) {
      if (currentRows[i][0]) config[String(currentRows[i][0])] = String(currentRows[i][1]);
    }
    config['driveFolderId'] = driveSetup.folderId || '';
    config['driveFolderUrl'] = driveSetup.folderUrl || '';

    sheet.clear();
    sheet.appendRow(['Key', 'Value']);
    var keys = Object.keys(config);
    if (keys.length > 0) {
      var rows = [];
      for (var k = 0; k < keys.length; k++) {
        rows.push([keys[k], config[keys[k]]]);
      }
      sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }
    return { status: 'success' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
