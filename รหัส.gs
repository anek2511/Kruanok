/**
 * @OnlyCurrentDoc
 */

// =================================================================
// GLOBAL CONSTANTS & CONFIGURATION
// =================================================================
const ss = SpreadsheetApp.getActiveSpreadsheet();
const STUDENTS_SHEET = "นักเรียน";
const SUBJECTS_SHEET = "รายวิชา";
const CLASSES_SHEET = "ชั้นเรียน";
const ASSIGNMENTS_SHEET = "งานที่มอบหมาย";
const SUBMISSIONS_SHEET = "ข้อมูลการส่งงาน";
const GRADEBOOK_SHEET = "สรุปคะแนน";

const HEADERS = {
  [STUDENTS_SHEET]: ["รหัสนักเรียน", "ชื่อ-สกุล", "ชั้นเรียน"],
  [SUBJECTS_SHEET]: ["รายวิชา"],
  [CLASSES_SHEET]: ["ชั้นเรียน"],
  [ASSIGNMENTS_SHEET]: ["AssignmentID", "ชื่องาน", "รายวิชา", "ชั้นเรียน", "วันที่สั่งงาน"],
  [SUBMISSIONS_SHEET]: ["SubmissionID", "AssignmentID", "วันที่ส่ง", "รหัสนักเรียน", "ชื่อ-สกุล", "ชั้นเรียน", "รายวิชา", "ชื่องาน", "Link", "ไฟล์แนบ (URL)", "หมายเหตุ", "คะแนน", "Comment"],
  [GRADEBOOK_SHEET]: ["รหัสนักเรียน", "ชื่อ-สกุล", "ชั้นเรียน", "รายวิชา", "คะแนนเก็บ", "คะแนนงาน", "กลางภาค", "ปลายภาค", "จิตพิสัย", "รวม", "เกรด"]
};

// =================================================================
// WEB APP ENTRY POINT
// =================================================================
function doGet(e) {
  setupInitialSheets();
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  html.setTitle("ระบบส่งงานนักเรียน").addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  return html;
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// =================================================================
// INITIAL SETUP - Ensures sheets and headers exist
// =================================================================
function setupInitialSheets() {
  Object.keys(HEADERS).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, HEADERS[sheetName].length).setValues([HEADERS[sheetName]]).setFontWeight("bold");
    } else {
      if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, HEADERS[sheetName].length).setValues([HEADERS[sheetName]]).setFontWeight("bold");
      }
    }
  });
}

// =================================================================
// AUTHENTICATION
// =================================================================
function checkLogin(credentials) {
  if (credentials.username === "admin" && credentials.password === "1234") {
    PropertiesService.getUserProperties().setProperty('isLoggedIn', 'true');
    return { success: true };
  }
  return { success: false };
}

function checkAuthStatus() {
  return PropertiesService.getUserProperties().getProperty('isLoggedIn') === 'true';
}

function logout() {
  PropertiesService.getUserProperties().deleteProperty('isLoggedIn');
  return { success: true };
}

function isAdmin() {
  if (!checkAuthStatus()) {
    throw new Error("Permission Denied. Administrator access required.");
  }
}

// =================================================================
// DATA FETCHING FUNCTIONS (Public & Admin)
// =================================================================
function getInitialData() {
  return {
    classes: getSheetData(CLASSES_SHEET),
    subjects: getSheetData(SUBJECTS_SHEET)
  };
}

function getSheetData(sheetName) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().filter(String);
  } catch (e) {
    return [];
  }
}

function getStudentsByClass(className) {
  try {
    const sheet = ss.getSheetByName(STUDENTS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const allStudents = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    const filtered = allStudents.filter(row => row[2] === className);
    return filtered.map(row => ({ studentId: row[0], studentName: row[1] }));
  } catch (e) {
    return [];
  }
}

function getAssignmentsForStudent(studentId) {
  try {
    if (!studentId) return {};

    const studentsSheet = ss.getSheetByName(STUDENTS_SHEET);
    const assignmentsSheet = ss.getSheetByName(ASSIGNMENTS_SHEET);
    const submissionsSheet = ss.getSheetByName(SUBMISSIONS_SHEET);

    // 1. Find student's class
    const students = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 3).getValues();
    const studentInfo = students.find(s => s[0].toString() === studentId.toString());
    if (!studentInfo) return {}; // Student not found
    const className = studentInfo[2];

    // 2. Get all assignments for that class
    const allAssignments = assignmentsSheet.getLastRow() > 1 ? assignmentsSheet.getRange(2, 1, assignmentsSheet.getLastRow() - 1, 4).getValues() : [];
    const classAssignments = allAssignments.filter(a => a[3] === className);

    // 3. Get all submissions by this student
    const allSubmissions = submissionsSheet.getLastRow() > 1 ? submissionsSheet.getRange(2, 1, submissionsSheet.getLastRow() - 1, 4).getValues() : [];
    const submittedAssignmentIds = new Set(
        allSubmissions
            .filter(sub => sub[3].toString() === studentId.toString())
            .map(sub => sub[1]) // Get AssignmentID
    );

    // 4. Filter for unsubmitted assignments and group by subject
    const unsubmittedBySubject = {};
    classAssignments.forEach(assignment => {
      const [assignmentId, assignmentName, subject] = assignment;
      if (!submittedAssignmentIds.has(assignmentId)) {
        if (!unsubmittedBySubject[subject]) {
          unsubmittedBySubject[subject] = [];
        }
        unsubmittedBySubject[subject].push({
          assignmentId: assignmentId,
          assignmentName: assignmentName
        });
      }
    });

    return unsubmittedBySubject;
  } catch (e) {
    Logger.log(e);
    return { error: e.message };
  }
}


// =================================================================
// DATA MANIPULATION FUNCTIONS (Admin Only)
// =================================================================
function addData(sheetName, value) {
  isAdmin();
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
    sheet.appendRow([value]);
    return { success: true, message: `เพิ่มข้อมูล '${value}' เรียบร้อยแล้ว` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addStudent(studentData) {
  isAdmin();
  try {
    const sheet = ss.getSheetByName(STUDENTS_SHEET);
    if (!sheet) throw new Error(`Sheet "${STUDENTS_SHEET}" not found.`);
    sheet.appendRow([studentData.studentId, studentData.studentName, studentData.className]);
    return { success: true, message: `เพิ่มนักเรียน '${studentData.studentName}' เรียบร้อยแล้ว` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function importStudents(data) {
  isAdmin();
  try {
    const sheet = ss.getSheetByName(STUDENTS_SHEET);
    if (!sheet) throw new Error(`Sheet "${STUDENTS_SHEET}" not found.`);
    sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    return { success: true, message: `นำเข้าข้อมูลนักเรียน ${data.length} คนสำเร็จ` };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการนำเข้า: ' + e.message };
  }
}

function createAssignment(assignmentData) {
  isAdmin();
  try {
    const sheet = ss.getSheetByName(ASSIGNMENTS_SHEET);
    const assignmentId = "ASN" + new Date().getTime();
    sheet.appendRow([assignmentId, assignmentData.assignmentName, assignmentData.subject, assignmentData.className, new Date()]);
    return { success: true, message: `มอบหมายงาน '${assignmentData.assignmentName}' เรียบร้อย` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// =================================================================
// SUBMISSION HANDLING
// =================================================================
function saveSubmission(data) {
  try {
    const submissionSheet = ss.getSheetByName(SUBMISSIONS_SHEET);
    const studentSheet = ss.getSheetByName(STUDENTS_SHEET);
    const assignmentSheet = ss.getSheetByName(ASSIGNMENTS_SHEET);

    const students = studentSheet.getRange(2, 1, studentSheet.getLastRow() - 1, 2).getValues();
    const student = students.find(row => row[0].toString() === data.studentId.toString());
    const studentName = student ? student[1] : "ไม่พบชื่อ";

    const assignments = assignmentSheet.getRange(2, 1, assignmentSheet.getLastRow() - 1, 2).getValues();
    const assignment = assignments.find(row => row[0].toString() === data.assignmentId.toString());
    const assignmentName = assignment ? assignment[1] : "ไม่พบชื่องาน";

    let fileUrl = "";
    if (data.fileData) {
      const folderName = "StudentSubmissions";
      let folders = DriveApp.getFoldersByName(folderName);
      let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      const blob = Utilities.newBlob(Utilities.base64Decode(data.fileData.split(',')[1]), data.fileType, data.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    const submissionId = "SUB" + new Date().getTime();
    const submissionRow = [
      submissionId, data.assignmentId, new Date(data.submissionDate), data.studentId, studentName,
      data.className, data.subject, assignmentName, data.link, fileUrl, data.notes, "รอตรวจ", ""
    ];
    submissionSheet.appendRow(submissionRow);
    return { success: true, message: "บันทึกการส่งงานเรียบร้อยแล้ว" };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.message };
  }
}


// =================================================================
// REPORTS AND DASHBOARD
// =================================================================
function getDashboardReport(filters) {
  try {
    const studentsSheet = ss.getSheetByName(STUDENTS_SHEET);
    const assignmentsSheet = ss.getSheetByName(ASSIGNMENTS_SHEET);
    const submissionsSheet = ss.getSheetByName(SUBMISSIONS_SHEET);

    if (studentsSheet.getLastRow() < 2 || assignmentsSheet.getLastRow() < 2) return { labels: [], series: [] };

    let students = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 3).getValues();
    let assignments = assignmentsSheet.getRange(2, 1, assignmentsSheet.getLastRow() - 1, 4).getValues();
    let submissions = (submissionsSheet.getLastRow() > 1) ?
      submissionsSheet.getRange(2, 1, submissionsSheet.getLastRow() - 1, 8).getValues().map(s => s[1] + "_" + s[3]) : []; // AssignmentID_StudentID
    const submissionSet = new Set(submissions);

    if (filters.className && filters.className !== 'all') {
      students = students.filter(s => s[2] === filters.className);
      assignments = assignments.filter(a => a[3] === filters.className);
    }
    if (filters.subject && filters.subject !== 'all') {
      assignments = assignments.filter(a => a[2] === filters.subject);
    }
    if (filters.studentId && filters.studentId !== 'all') {
      students = students.filter(s => s[0].toString() === filters.studentId.toString());
    }

    const reportData = {}; // Key: "Class - Subject", Value: { total: 0, submitted: 0 }

    assignments.forEach(assignment => {
      const assignmentId = assignment[0];
      const subject = assignment[2];
      const className = assignment[3];
      const key = `${className} - ${subject}`;

      if (!reportData[key]) {
        reportData[key] = { total: 0, submitted: 0 };
      }

      const studentsInClass = students.filter(s => s[2] === className);
      studentsInClass.forEach(student => {
        reportData[key].total++;
        if (submissionSet.has(assignmentId + "_" + student[0])) {
          reportData[key].submitted++;
        }
      });
    });

    const labels = Object.keys(reportData);
    const series = labels.map(key => {
      const { total, submitted } = reportData[key];
      return total > 0 ? parseFloat(((submitted / total) * 100).toFixed(2)) : 0;
    });

    return { labels, series };
  } catch (e) {
    return { error: e.message };
  }
}

function getUnsubmittedReport(filters) {
  try {
    const studentsSheet = ss.getSheetByName(STUDENTS_SHEET);
    const assignmentsSheet = ss.getSheetByName(ASSIGNMENTS_SHEET);
    const submissionsSheet = ss.getSheetByName(SUBMISSIONS_SHEET);

    let students = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 3).getValues();
    let assignments = assignmentsSheet.getRange(2, 1, assignmentsSheet.getLastRow() - 1, 4).getValues();
    let submissions = (submissionsSheet.getLastRow() > 1) ?
      submissionsSheet.getRange(2, 1, submissionsSheet.getLastRow() - 1, 4).getValues() : []; // [subId, asnId, date, studentId]

    if (filters.className && filters.className !== 'all') {
      students = students.filter(s => s[2] === filters.className);
      assignments = assignments.filter(a => a[3] === filters.className);
    }
    if (filters.subject && filters.subject !== 'all') {
      assignments = assignments.filter(a => a[2] === filters.subject);
    }

    const report = [];
    assignments.forEach(assignment => {
      const [assignmentId, assignmentName, subject, className] = assignment;
      const studentsInClass = students.filter(s => s[2] === className);
      const submittedStudents = new Set(
        submissions.filter(s => s[1] === assignmentId).map(s => s[3].toString())
      );

      const unsubmittedStudents = studentsInClass
        .filter(student => !submittedStudents.has(student[0].toString()))
        .map(student => student[1]); 

      if (unsubmittedStudents.length > 0) {
        report.push({
          assignmentName: `${assignmentName} (${subject})`,
          unsubmitted: unsubmittedStudents
        });
      }
    });

    return report;
  } catch (e) {
    return { error: e.message };
  }
}

function getStudentScores(studentId) {
  if (!studentId) return { error: "กรุณาระบุรหัสนักเรียน" };
  try {
    const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
    if (sheet.getLastRow() < 2) return { no_data: true };

    const allSubmissions = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const studentSubmissions = allSubmissions.filter(row => row[3].toString() === studentId.toString());

    if (studentSubmissions.length === 0) return { no_data: true };

    const scoresBySubject = {};
    studentSubmissions.forEach(sub => {
      const subject = sub[6];
      if (!scoresBySubject[subject]) {
        scoresBySubject[subject] = { submissions: [], totalScore: 0 };
      }
      const scoreValue = parseFloat(sub[11]);
      if (!isNaN(scoreValue)) {
        scoresBySubject[subject].totalScore += scoreValue;
      }
      scoresBySubject[subject].submissions.push({
        submissionId: sub[0], assignmentName: sub[7], date: sub[2] ? new Date(sub[2]).toLocaleDateString('th-TH') : '',
        score: sub[11], comment: sub[12] || "-", link: sub[8], fileUrl: sub[9]
      });
    });
    return scoresBySubject;
  } catch (e) {
    return { error: e.message };
  }
}

// NEW FUNCTION to get final grades for a student
function getStudentGrades(studentId) {
  if (!studentId) return { error: "กรุณาระบุรหัสนักเรียน" };
  try {
    const sheet = ss.getSheetByName(GRADEBOOK_SHEET);
    if (sheet.getLastRow() < 2) return { no_data: true };

    const allGrades = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const studentGradesData = allGrades.filter(row => row[0].toString() === studentId.toString());

    if (studentGradesData.length === 0) return { no_data: true };
    
    // HEADERS index: ["รหัสนักเรียน"[0], "ชื่อ-สกุล"[1], "ชั้นเรียน"[2], "รายวิชา"[3], ..., "รวม"[9], "เกรด"[10]]
    const formattedGrades = studentGradesData.map(row => {
      return {
        subject: row[3], // รายวิชา
        total: row[9],   // รวม
        grade: row[10]   // เกรด
      };
    });

    return formattedGrades;

  } catch (e) {
    Logger.log(e);
    return { error: e.message };
  }
}


function getAdminReportData(filters) {
  isAdmin();
  try {
    const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
    if (sheet.getLastRow() < 2) return [];

    let submissions = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    if (filters.className && filters.className !== 'all') {
      submissions = submissions.filter(s => s[5] === filters.className);
    }
    if (filters.subject && filters.subject !== 'all') {
      submissions = submissions.filter(s => s[6] === filters.subject);
    }
    if (filters.studentId && filters.studentId !== 'all') {
      submissions = submissions.filter(s => s[3].toString() === filters.studentId.toString());
    }

    return submissions.map(s => ({
      submissionId: s[0], studentName: s[4], subject: s[6], assignmentName: s[7],
      date: s[2] ? new Date(s[2]).toLocaleDateString('th-TH') : '', score: s[11],
      comment: s[12] || "", link: s[8], fileUrl: s[9]
    }));
  } catch (e) {
    return { error: e.message };
  }
}

function updateScore(data) {
  isAdmin();
  try {
    const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const submissionIdCol = 0;
    const scoreCol = 11;
    const commentCol = 12;

    for (let i = 1; i < values.length; i++) {
      if (values[i][submissionIdCol] === data.submissionId) {
        sheet.getRange(i + 1, scoreCol + 1).setValue(data.score);
        sheet.getRange(i + 1, commentCol + 1).setValue(data.comment);
        return { success: true, message: "อัปเดตคะแนนเรียบร้อย" };
      }
    }
    return { success: false, message: "ไม่พบข้อมูลการส่งงาน" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}


function getGradebookData(className, subject) {
  isAdmin();
  try {
    const students = getStudentsByClass(className);
    const gradebookSheet = ss.getSheetByName(GRADEBOOK_SHEET);
    const scoreData = gradebookSheet.getLastRow() > 1 ?
      gradebookSheet.getRange(2, 1, gradebookSheet.getLastRow() - 1, gradebookSheet.getLastColumn()).getValues() : [];

    return students.map(student => {
      const row = scoreData.find(r =>
        r[2] === className &&
        r[3] === subject &&
        r[0].toString() === student.studentId.toString()
      );
      return {
        studentId: student.studentId,
        studentName: student.studentName,
        collected: row ? row[4] : 0,
        work: row ? row[5] : 0,
        midterm: row ? row[6] : 0,
        final: row ? row[7] : 0,
        attitude: row ? row[8] : 0,
        total: row ? row[9] : 0,
        grade: row ? row[10] : "-"
      };
    });
  } catch (e) {
    return { error: e.message };
  }
}

function saveGradebookData(className, subject, dataToSave) {
  isAdmin();
  try {
    const sheet = ss.getSheetByName(GRADEBOOK_SHEET);
    let allData = sheet.getDataRange().getValues();
    const dataMap = new Map();
    for (let i = 1; i < allData.length; i++) {
        const key = `${allData[i][2]}_${allData[i][3]}_${allData[i][0]}`; // Class_Subject_StudentID
        dataMap.set(key, i + 1); 
    }

    const rowsToAdd = [];
    dataToSave.forEach(student => {
        const key = `${className}_${subject}_${student.studentId}`;
        const newRowData = [
            student.studentId, student.studentName, className, subject,
            student.collected, student.work, student.midterm, student.final, student.attitude,
            student.total, student.grade
        ];

        if (dataMap.has(key)) {
            const rowNum = dataMap.get(key);
            sheet.getRange(rowNum, 1, 1, newRowData.length).setValues([newRowData]);
        } else {
            rowsToAdd.push(newRowData);
        }
    });

    if (rowsToAdd.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }
    return { success: true, message: "บันทึกข้อมูลคะแนนและเกรดเรียบร้อย" };
  } catch (e) {
      return { success: false, message: e.message };
  }
}

