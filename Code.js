// Ganti nilai di bawah ini dengan ID Spreadsheet Anda.
const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE";
const RESPONSE_SHEET_NAME = "Responses"; // Nama sheet untuk menyimpan jawaban siswa
const SCHEDULE_SHEET_NAME = "Jadwal"; // Nama sheet untuk menyimpan jadwal
const ADMIN_EMAIL = "admin@example.com"; // [PENTING] Ganti dengan email Anda
const DEFAULT_TIME = 180; // Waktu default per soal dalam detik
const UPLOAD_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID_HERE"; // [PENTING] Ganti dengan ID folder Google Drive untuk menyimpan file soal

// [DIUBAH] Daftar nama kelas diperbarui sesuai permintaan
const CLASS_NAMES_X = ["E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9"];
const CLASS_NAMES_XII = ["XII 1", "XII 2"];
const CLASS_NAMES = [...CLASS_NAMES_X, ...CLASS_NAMES_XII]; // Gabungkan semua kelas
const TRIGGER_FUNCTION_NAME = "checkScheduledQuizStatus";

/**
 * [BARU] Fungsi diagnostik untuk memeriksa izin akses ke Spreadsheet dan Folder.
 * Jalankan fungsi ini secara manual dari editor skrip untuk memverifikasi izin.
 */
function checkPermissions() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log(`BERHASIL mengakses Spreadsheet: "${spreadsheet.getName()}"`);
  } catch (e) {
    Logger.log(
      `GAGAL mengakses Spreadsheet ID "${SPREADSHEET_ID}". Pesan Error: ${e.message}`
    );
    Logger.log(
      "Pastikan akun Anda memiliki setidaknya akses 'Viewer' ke Google Sheet ini."
    );
  }

  try {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    Logger.log(`BERHASIL mengakses Folder: "${folder.getName()}"`);
  } catch (e) {
    Logger.log(
      `GAGAL mengakses Folder ID "${UPLOAD_FOLDER_ID}". Pesan Error: ${e.message}`
    );
    Logger.log(
      "Pastikan akun Anda memiliki setidaknya akses 'Viewer' ke Folder Google Drive ini."
    );
  }
}

/**
 * Fungsi utama untuk menyajikan aplikasi web.
 */
function doGet(e) {
  let page = e.parameter.page || "quiz";

  if (page === "dashboard") {
    return HtmlService.createTemplateFromFile("dashboard")
      .evaluate()
      .setTitle("E-Quiz - Dashboard")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  }

  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("E-Quiz")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

/**
 * Mengambil waktu pengerjaan soal pertama dari sheet aktif.
 * Jika tidak ditemukan, mengembalikan DEFAULT_TIME.
 */
function getQuizTimeFromSheet() {
  try {
    const activeSheetName = getActiveQuestionSheetName();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(activeSheetName);
    if (!sheet) return DEFAULT_TIME;
    
    // Baris 1 adalah header, baris 2 adalah data soal pertama
    const timeValue = sheet.getRange(2, 8).getValue(); // Kolom H (indeks 8) adalah "Waktu (detik)"
    return parseInt(timeValue) || DEFAULT_TIME;
  } catch(e) {
    return DEFAULT_TIME;
  }
}

/**
 * [DIPERBARUI] Fungsi untuk menangani upload file soal dengan logika yang lebih stabil.
 */
function uploadSoalFile(fileData, sheetName) {
  try {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.data),
      fileData.mimeType,
      fileData.fileName
    );

    // Proses impor otomatis khusus untuk file Excel (.xlsx)
    if (
      fileData.mimeType ===
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      fileData.fileName.toLowerCase().endsWith(".xlsx")
    ) {
      const targetSs = SpreadsheetApp.openById(SPREADSHEET_ID);
      if (targetSs.getSheetByName(sheetName)) {
        return {
          success: false,
          error: `Gagal: Sheet dengan nama "${sheetName}" sudah ada. Silakan gunakan nama lain.`,
        };
      }

      let tempSheetFileId = null;
      try {
        // Konversi file .xlsx menjadi Google Sheet sementara
        const tempSheetFile = Drive.Files.insert(
          {
            title: `TEMP_CONVERT_${sheetName}_${new Date().getTime()}`,
            mimeType: "application/vnd.google-apps.spreadsheet",
          },
          blob
        );

        tempSheetFileId = tempSheetFile.id;

        // Jeda singkat untuk memastikan file siap diakses
        Utilities.sleep(5000);

        // Buka spreadsheet sementara dan baca datanya
        const tempSs = SpreadsheetApp.openById(tempSheetFileId);
        const sourceSheet = tempSs.getSheets()[0];
        if (!sourceSheet) {
          throw new Error(
            "File Excel yang diupload kosong atau tidak dapat dibaca."
          );
        }
        const data = sourceSheet.getDataRange().getValues();

        // Ekstrak gambar dari xlsx dan ganti di data
        const cellImageMap = extractImagesFromXlsx(blob);
        const headers = data[0];
        if (headers && Object.keys(cellImageMap).length > 0) {
            const imageIdColIndex = headers.indexOf("imageId");
            if (imageIdColIndex !== -1) {
                for (let r = 1; r < data.length; r++) {
                    const mapKey = r + "_" + imageIdColIndex;
                    if (cellImageMap[mapKey]) {
                        data[r][imageIdColIndex] = cellImageMap[mapKey];
                    }
                }
            }
        }

        // Buat sheet baru di spreadsheet tujuan dan tulis datanya
        const newSheet = targetSs.insertSheet(sheetName);
        newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      } finally {
        // Pastikan file sementara selalu dihapus
        if (tempSheetFileId) {
          Drive.Files.remove(tempSheetFileId);
        }
      }

      // Simpan juga file .xlsx asli di folder Drive sebagai cadangan
      folder.createFile(blob);

      return {
        success: true,
        message: `File "${fileData.fileName}" berhasil diimpor sebagai sheet baru bernama "${sheetName}".`,
      };
    }

    // Untuk tipe file lain, hanya simpan di folder Drive
    const newFile = folder.createFile(blob);
    return {
      success: true,
      message: `File "${
        fileData.fileName
      }" berhasil diupload ke Drive. ID File: ${newFile.getId()}`,
    };
  } catch (e) {
    Logger.log("Upload Error Stack: " + e.stack);
    let errorMessage = e.message || String(e);

    if (errorMessage.includes("Drive.Files.insert is not a function")) {
      errorMessage =
        "Layanan 'Drive API' belum diaktifkan. Silakan aktifkan di editor skrip (Layanan > Drive API > Tambah).";
    } else if (
      errorMessage.includes("file is inaccessible") ||
      errorMessage.includes("Bad value")
    ) {
      errorMessage =
        "File sementara tidak dapat diakses, mungkin ada masalah propagasi di Google Drive. Coba lagi dalam beberapa saat.";
    }

    return { success: false, error: "Gagal mengupload file: " + errorMessage };
  }
}

/**
 * Mendapatkan URL aplikasi web.
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Fungsi untuk mengatur jenjang kelas yang aktif.
 */
function setActiveGrade(grade) {
  try {
    PropertiesService.getScriptProperties().setProperty("ACTIVE_GRADE", grade);
    return { success: true, activeGrade: grade };
  } catch (e) {
    Logger.log(e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * Fungsi untuk mendapatkan jenjang kelas yang aktif.
 */
function getActiveGrade() {
  return (
    PropertiesService.getScriptProperties().getProperty("ACTIVE_GRADE") || "X"
  ); // Default ke Kelas X
}

/**
 * Fungsi yang dipanggil oleh halaman kuis untuk mendapatkan pengaturan awal.
 */
function getQuizPageData() {
  try {
    const activeGrade = getActiveGrade();
    let classList;
    if (activeGrade === "X") {
      classList = CLASS_NAMES_X;
    } else {
      // 'XII'
      classList = CLASS_NAMES_XII;
    }
    const time = getQuizTimeFromSheet();
    return { classes: classList, quizTime: time };
  } catch (e) {
    Logger.log(e.stack);
    return { error: e.message };
  }
}

/**
 * Helper function untuk memfilter jadwal yang akan ditampilkan di dashboard.
 */
function filterSchedulesForDisplay(allSchedules) {
  const now = new Date().getTime();
  const upcomingSchedules = [];
  const activeSchedules = [];
  const finishedSchedules = [];

  allSchedules.forEach((s) => {
    const startTime = new Date(s.startTime).getTime();
    const endTime = new Date(s.endTime).getTime();
    if (endTime < now) {
      finishedSchedules.push(s);
    } else if (startTime > now) {
      upcomingSchedules.push(s);
    } else {
      activeSchedules.push(s);
    }
  });

  // Urutkan jadwal yang sudah selesai berdasarkan waktu berakhirnya, dari yang terbaru ke terlama
  finishedSchedules.sort(
    (a, b) => new Date(b.endTime).getTime() - new Date(a.endTime).getTime()
  );

  // Ambil hanya jadwal terakhir yang selesai, jika ada
  const lastFinishedSchedule =
    finishedSchedules.length > 0 ? [finishedSchedules[0]] : [];

  // Gabungkan jadwal yang aktif, yang akan datang, dan yang terakhir selesai
  return [...activeSchedules, ...upcomingSchedules, ...lastFinishedSchedule];
}

// --- FUNGSI MANAJEMEN JADWAL & TRIGGER ---

/**
 * Menghapus semua trigger proyek yang ada untuk penjadwalan.
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Mengatur satu trigger berulang yang memeriksa status setiap menit.
 */
function setupRecurringTrigger() {
  deleteTriggers(); // Hapus trigger lama untuk menghindari duplikasi
  ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log("Trigger pengecekan jadwal berulang telah dibuat.");
}

/**
 * Mendapatkan semua jadwal dari sheet "Jadwal".
 */
function getQuizSchedules() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SCHEDULE_SHEET_NAME);
      sheet.appendRow([
        "ID",
        "Waktu Mulai",
        "Waktu Selesai",
        "Nama Soal",
        "Nama Kelas",
      ]);
      sheet.setFrozenRows(1);
      return [];
    }

    const range = sheet.getDataRange();
    if (range.getNumRows() <= 1) return [];

    const values = range.getValues();
    const headers = values.shift();

    const idIndex = headers.indexOf("ID");
    const startIndex = headers.indexOf("Waktu Mulai");
    const endIndex = headers.indexOf("Waktu Selesai");
    const sheetIndex = headers.indexOf("Nama Soal");
    const classIndex = headers.indexOf("Nama Kelas");

    if ([idIndex, startIndex, endIndex, sheetIndex, classIndex].includes(-1)) {
      Logger.log("Header di sheet Jadwal tidak lengkap atau salah nama.");
      return [];
    }

    const schedules = values
      .map((row) => {
        if (row.join("").trim() === "") return null;
        return {
          id: row[idIndex].toString(),
          startTime: new Date(row[startIndex]).toISOString(),
          endTime: new Date(row[endIndex]).toISOString(),
          sheetName: row[sheetIndex],
          classNames: row[classIndex]
            .toString()
            .split(",")
            .map((name) => name.trim()),
        };
      })
      .filter((s) => s !== null);

    return schedules;
  } catch (e) {
    Logger.log("Gagal membaca jadwal dari sheet: " + e.stack);
    return [];
  }
}

/**
 * Menambah atau memperbarui jadwal kuis di sheet.
 */
function addOrUpdateQuizSchedule(scheduleData) {
  try {
    const schedules = getQuizSchedules();

    // Tambahkan offset timezone WIB (+07:00) ke input pengguna
    const newStartTime = new Date(scheduleData.startTime + "+07:00").getTime();
    const newEndTime = new Date(scheduleData.endTime + "+07:00").getTime();

    // Cek apakah ada jadwal yang bentrok
    const isOverlapping = schedules.some((s) => {
      if (scheduleData.id && s.id === scheduleData.id) {
        return false;
      }
      const existingStartTime = new Date(s.startTime).getTime();
      const existingEndTime = new Date(s.endTime).getTime();
      return newStartTime < existingEndTime && newEndTime > existingStartTime;
    });

    if (isOverlapping) {
      return {
        success: false,
        error:
          "Jadwal yang Anda masukkan bentrok dengan jadwal yang sudah ada.",
      };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SCHEDULE_SHEET_NAME);
      sheet.appendRow([
        "ID",
        "Waktu Mulai",
        "Waktu Selesai",
        "Nama Soal",
        "Nama Kelas",
      ]);
      sheet.setFrozenRows(1);
    }
    if (sheet.getLastRow() < 1) {
      sheet.appendRow([
        "ID",
        "Waktu Mulai",
        "Waktu Selesai",
        "Nama Soal",
        "Nama Kelas",
      ]);
      sheet.setFrozenRows(1);
    }

    if (sheet.getLastColumn() < 5) {
      throw new Error(
        "Sheet 'Jadwal' tidak memiliki format yang benar. Harap hapus dan biarkan skrip membuatnya kembali, atau pastikan memiliki kolom: ID, Waktu Mulai, Waktu Selesai, Nama Soal, Nama Kelas."
      );
    }
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const idIndex = headers.indexOf("ID");

    if (idIndex === -1) {
      throw new Error(
        "Kolom 'ID' tidak ditemukan di sheet 'Jadwal'. Mohon periksa header."
      );
    }

    const startTimeForSheet = new Date(newStartTime);
    const endTimeForSheet = new Date(newEndTime);
    const classNames = scheduleData.classNames.join(",");

    if (scheduleData.id) {
      const data = sheet.getDataRange().getValues();
      const rowIndex =
        data.findIndex((row) => row[idIndex] == scheduleData.id) + 1;
      if (rowIndex > 1) {
        sheet
          .getRange(rowIndex, 1, 1, 5)
          .setValues([
            [
              scheduleData.id,
              startTimeForSheet,
              endTimeForSheet,
              scheduleData.sheetName,
              classNames,
            ],
          ]);
      } else {
        throw new Error(
          "Jadwal dengan ID tersebut tidak ditemukan untuk diupdate."
        );
      }
    } else {
      const newId = new Date().getTime().toString();
      sheet.appendRow([
        newId,
        startTimeForSheet,
        endTimeForSheet,
        scheduleData.sheetName,
        classNames,
      ]);
    }

    SpreadsheetApp.flush(); // Paksa penyimpanan perubahan sebelum melanjutkan.

    if (
      ScriptApp.getProjectTriggers().filter(
        (t) => t.getHandlerFunction() === TRIGGER_FUNCTION_NAME
      ).length === 0
    ) {
      setupRecurringTrigger();
    }

    checkScheduledQuizStatus();

    return {
      success: true,
      schedules: filterSchedulesForDisplay(getQuizSchedules()),
    };
  } catch (e) {
    Logger.log(`addOrUpdateQuizSchedule Error: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

/**
 * Fungsi yang dijalankan oleh trigger untuk mengaktifkan/menonaktifkan kuis.
 */
function checkScheduledQuizStatus() {
  try {
    const props = PropertiesService.getScriptProperties();
    const schedules = getQuizSchedules();

    if (schedules && schedules.length > 0) {
      const now = new Date();
      const activeSchedule = schedules.find((s) => {
        const startTime = new Date(s.startTime);
        const endTime = new Date(s.endTime);
        return now >= startTime && now < endTime;
      });

      if (activeSchedule) {
        const updates = {
          "QUIZ_GLOBAL_STATUS": "active",
          "ACTIVE_QUESTION_SHEET": activeSchedule.sheetName
        };
        CLASS_NAMES.forEach((className) => {
          const isClassActive = activeSchedule.classNames.includes(className);
          updates["QUIZ_STATUS_" + className] = isClassActive.toString();
        });
        props.setProperties(updates);
        Logger.log(
          `Jadwal aktif ditemukan: ID ${activeSchedule.id}. Kuis diaktifkan.`
        );
      } else {
        const updates = { "QUIZ_GLOBAL_STATUS": "paused" };
        CLASS_NAMES.forEach((className) => {
          updates["QUIZ_STATUS_" + className] = "false";
        });
        props.setProperties(updates);
        Logger.log(
          "Ada jadwal, tetapi tidak ada yang aktif saat ini. Kuis dijeda."
        );
      }
    } else {
      Logger.log("Tidak ada jadwal ditemukan. Kontrol manual tetap aktif.");
    }
  } catch (e) {
    Logger.log(`checkScheduledQuizStatus Error: ${e.stack}`);
  }
}

/**
 * Menghapus satu jadwal kuis dari sheet.
 */
function deleteQuizSchedule(scheduleId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
    if (!sheet) return { success: true, schedules: [] };

    const data = sheet.getDataRange().getValues();
    const idIndex = data[0].indexOf("ID");
    const rowIndexToDelete =
      data.findIndex((row) => row[idIndex] == scheduleId) + 1;

    if (rowIndexToDelete > 1) {
      sheet.deleteRow(rowIndexToDelete);
      SpreadsheetApp.flush();
    }

    const remainingSchedules = getQuizSchedules();
    if (remainingSchedules.length === 0) {
      deleteTriggers();
      const updates = { "QUIZ_GLOBAL_STATUS": "paused" };
      CLASS_NAMES.forEach((className) => {
        updates["QUIZ_STATUS_" + className] = "false";
      });
      PropertiesService.getScriptProperties().setProperties(updates);
      Logger.log(
        "Jadwal terakhir dihapus. Beralih ke kontrol manual dengan semua kelas nonaktif."
      );
    }

    checkScheduledQuizStatus();

    return {
      success: true,
      schedules: filterSchedulesForDisplay(remainingSchedules),
    };
  } catch (e) {
    Logger.log(`deleteQuizSchedule Error: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

/**
 * Menghapus semua jadwal dari sheet.
 */
function deleteAllSchedules() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
    if (sheet) {
      if (sheet.getLastRow() > 1) {
        sheet
          .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
          .clearContent();
        SpreadsheetApp.flush();
      }
    }

    deleteTriggers();

    const updates = { "QUIZ_GLOBAL_STATUS": "paused" };
    CLASS_NAMES.forEach((className) => {
      updates["QUIZ_STATUS_" + className] = "false";
    });
    PropertiesService.getScriptProperties().setProperties(updates);
    Logger.log(
      "Semua jadwal dihapus. Beralih ke kontrol manual dengan semua kelas nonaktif."
    );

    return { success: true };
  } catch (e) {
    Logger.log(`deleteAllSchedules Error: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

// --- FUNGSI INTI APLIKASI KUIS ---

/**
 * Mengambil nama sheet soal yang aktif.
 */
function getActiveQuestionSheetName() {
  return (
    PropertiesService.getScriptProperties().getProperty(
      "ACTIVE_QUESTION_SHEET"
    ) || "Sheet1"
  );
}

/**
 * Mengambil semua data kuis: soal, ID gambar, URL gambar, dan status kuis.
 */
function getQuizData() {
  try {
    const activeSheetName = getActiveQuestionSheetName();
    const sheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(activeSheetName);
    if (!sheet)
      throw new Error(`Sheet soal aktif "${activeSheetName}" tidak ditemukan.`);

    const quizStatus =
      PropertiesService.getScriptProperties().getProperty(
        "QUIZ_GLOBAL_STATUS"
      ) || "active";
    if (quizStatus === "paused") {
      const schedules = getQuizSchedules();
      const now = new Date();
      const upcomingSchedule = schedules.find(
        (s) => new Date(s.startTime) > now
      );
      if (upcomingSchedule) {
        const startTime = new Date(upcomingSchedule.startTime);
        return {
          error: `Kuis belum dimulai. Kuis akan dimulai pada ${startTime.toLocaleString(
            "id-ID",
            { timeZone: "Asia/Jakarta" }
          )}.`,
        };
      }
      return {
        error:
          "Kuis saat ini sedang dijeda oleh admin atau waktu pengerjaan telah berakhir.",
      };
    }

    const values = sheet.getDataRange().getDisplayValues();
    values.shift();

    const questions = [];
    const imageIds = new Set();

    values.forEach((row) => {
      if (row[0] && row[1] && row[2]) {
        const imageId = row[6] ? row[6].trim() : null;
        if (imageId) imageIds.add(imageId);
        questions.push({
          question: row[0],
          options: {
            A: row[1],
            B: row[2],
            C: row[3] || null,
            D: row[4] || null,
            E: row[5] || null,
          },
          imageId: imageId,
          time: parseInt(row[7]) || DEFAULT_TIME,
        });
      }
    });

    const images = {};
    imageIds.forEach((id) => {
      try {
        const file = DriveApp.getFileById(id);
        const blob = file.getBlob();
        images[
          id
        ] = `data:${blob.getContentType()};base64,${Utilities.base64Encode(
          blob.getBytes()
        )}`;
      } catch (imgErr) {
        Logger.log(`Gagal memproses gambar ID ${id}: ${imgErr.message}`);
      }
    });

    return { questions, images, quizStatus };
  } catch (e) {
    Logger.log(e.stack);
    return { error: `Gagal mengambil data kuis: ${e.message}` };
  }
}

/**
 * Memeriksa jawaban, menghitung skor, dan mencatatnya.
 */
function checkAnswers(submissionData) {
  try {
    const {
      answers: userAnswers,
      name: userName,
      class: userClass,
      ipAddress,
    } = submissionData;
    if (!userAnswers || !userName || !userClass)
      throw new Error("Data nama, kelas, atau jawaban tidak lengkap.");

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const activeSheetName = getActiveQuestionSheetName();
    const questionSheet = ss.getSheetByName(activeSheetName);
    if (!questionSheet)
      throw new Error(`Sheet soal aktif "${activeSheetName}" tidak ditemukan.`);

    const questionValues = questionSheet.getDataRange().getValues();
    questionValues.shift();

    let userScore = 0;
    let maxScore = 0;
    const answerDetails = {};

    questionValues.forEach((row, index) => {
      if (row[0]) {
        const correctAnswer = String(row[8] || "")
          .trim()
          .toUpperCase();
        const questionScore = parseInt(row[9]) || 0;
        maxScore += questionScore;

        const userAnswerObj = userAnswers[index];
        const userAnswer = userAnswerObj ? userAnswerObj.answer : "";
        answerDetails[`Soal ${index + 1}`] = userAnswer;

        if (userAnswer && userAnswer.toUpperCase() === correctAnswer) {
          userScore += questionScore;
        }
      }
    });

    const responseSheet = setupResponseSheet(ss, questionValues.length);
    const timestamp = new Date();
    const scoreString = `${userScore} / ${maxScore}`;

    const newRow = ["", userName, userClass, scoreString, ipAddress];
    const questionHeaders = Object.keys(answerDetails).sort(
      (a, b) => parseInt(a.split(" ")[1]) - parseInt(b.split(" ")[1])
    );
    questionHeaders.forEach((header) => newRow.push(answerDetails[header]));

    const newRowData = [timestamp, ...newRow.slice(1)];
    responseSheet.appendRow(newRowData);
    SpreadsheetApp.flush();

    return { score: userScore, total: maxScore };
  } catch (e) {
    Logger.log(e.stack);
    return { error: `Gagal memeriksa jawaban: ${e.message}` };
  }
}

/**
 * Mengatur header pada sheet "Responses".
 */
function setupResponseSheet(spreadsheet, questionCount) {
  let sheet = spreadsheet.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(RESPONSE_SHEET_NAME);
  }

  const baseHeaders = [
    "Waktu Pengerjaan",
    "Nama",
    "Kelas",
    "Skor",
    "Alamat IP",
  ];
  const questionHeaders = Array.from(
    { length: questionCount },
    (_, i) => `Soal ${i + 1}`
  );
  const fullHeaders = baseHeaders.concat(questionHeaders);

  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent();
  sheet.getRange(1, 1, 1, fullHeaders.length).setValues([fullHeaders]);

  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * Memeriksa apakah siswa sudah menyelesaikan kuis atau kuis untuk kelasnya aktif.
 */
function checkStudentCompletion(studentDetails) {
  try {
    const { name, studentClass } = studentDetails;
    if (!name || !studentClass) throw new Error("Nama dan kelas diperlukan.");

    const schedules = getQuizSchedules();
    const now = new Date();

    const upcomingScheduleForClass = schedules
      .filter(
        (s) =>
          s.classNames.includes(studentClass) && new Date(s.startTime) > now
      )
      .sort((a, b) => new Date(a.startTime) - new Date(b.startTime))[0];

    const classStatus = PropertiesService.getScriptProperties().getProperty(
      "QUIZ_STATUS_" + studentClass
    );

    if (classStatus !== "true") {
      if (upcomingScheduleForClass) {
        const startTime = new Date(upcomingScheduleForClass.startTime);
        return {
          error: `Kuis untuk kelas Anda belum dimulai. Kuis akan dimulai pada ${startTime.toLocaleString(
            "id-ID",
            { timeZone: "Asia/Jakarta" }
          )}.`,
        };
      }
      return { completed: false, active: false };
    }

    const sheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(
        RESPONSE_SHEET_NAME
      );
    if (!sheet) return { completed: false, active: true };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const nameIndex = headers.indexOf("Nama");
    const classIndex = headers.indexOf("Kelas");
    const scoreIndex = headers.indexOf("Skor");

    if (nameIndex === -1 || classIndex === -1 || scoreIndex === -1) {
      throw new Error(
        "Header 'Nama', 'Kelas', atau 'Skor' tidak ditemukan di sheet Responses."
      );
    }

    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      if (
        String(row[nameIndex]).trim() === name &&
        String(row[classIndex]).trim() === studentClass
      ) {
        const scoreParts = String(row[scoreIndex]).split("/");
        return {
          completed: true,
          active: true,
          score: scoreParts[0] ? scoreParts[0].trim() : 0,
          total: scoreParts[1] ? scoreParts[1].trim() : 0,
        };
      }
    }
    return { completed: false, active: true };
  } catch (e) {
    Logger.log(e.stack);
    return { error: e.message };
  }
}

/**
 * Mengambil hasil kuis dari sheet "Responses".
 */
function getQuizResults() {
  try {
    const sheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(
        RESPONSE_SHEET_NAME
      );
    if (!sheet) return { results: [] };

    const range = sheet.getDataRange();
    if (range.getNumRows() <= 1) return { results: [] };

    const values = range.getDisplayValues();
    const headers = values.shift();

    const coreHeaders = [
      "Waktu Pengerjaan",
      "Alamat IP",
      "Nama",
      "Kelas",
      "Skor",
    ];
    const headerIndices = coreHeaders.map((h) => headers.indexOf(h));

    const results = values
      .map((row, index) => {
        if (row.join("").trim() === "") return null;

        let resultObj = { rowIndex: index + 2 };
        headerIndices.forEach((hIndex, i) => {
          if (hIndex !== -1) {
            resultObj[coreHeaders[i]] = row[hIndex];
          }
        });
        return resultObj;
      })
      .filter((r) => r !== null)
      .reverse();

    return { results: results };
  } catch (e) {
    Logger.log(e.stack);
    return { error: `Gagal mengambil hasil: ${e.message}` };
  }
}

/**
 * Menghapus data satu siswa dari sheet "Responses".
 */
function resetSingleStudentData(rowIndex) {
  try {
    const sheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(
        RESPONSE_SHEET_NAME
      );
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    Logger.log(e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * Mengirim email permintaan izin mengerjakan ulang.
 */
function sendPermissionRequestEmail(studentDetails) {
  if (!ADMIN_EMAIL || ADMIN_EMAIL === "ganti.dengan.email.admin@gmail.com") {
    return {
      success: false,
      error: "Alamat email admin belum diatur di file Code.js.",
    };
  }
  try {
    const { name, studentClass } = studentDetails;
    const subject = `Permintaan Mengerjakan Ulang Kuis dari ${name}`;
    const body = `Siswa berikut meminta izin untuk mengerjakan ulang kuis:\n\nNama: ${name}\nKelas: ${studentClass}\n\nAnda dapat mereset sesi siswa ini melalui dashboard aplikasi kuis.`;
    MailApp.sendEmail(ADMIN_EMAIL, subject, body);
    return { success: true };
  } catch (e) {
    Logger.log(e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * Memeriksa password dashboard.
 */
function checkPassword(password) {
  const CORRECT_PASSWORD = "demo";
  return password === CORRECT_PASSWORD;
}

/**
 * Mengatur status jeda global untuk kuis.
 */
function setQuizStatus(status) {
  try {
    if (getQuizSchedules().length > 0) {
      return {
        success: false,
        error: "Tidak dapat mengubah status manual saat jadwal otomatis aktif.",
      };
    }
    PropertiesService.getScriptProperties().setProperty(
      "QUIZ_GLOBAL_STATUS",
      status
    );
    return { success: true, status: status };
  } catch (e) {
    Logger.log(e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * Mendapatkan status jeda global kuis.
 */
function getQuizStatus() {
  return (
    PropertiesService.getScriptProperties().getProperty("QUIZ_GLOBAL_STATUS") ||
    "active"
  );
}

/**
 * Mengatur status aktif kuis per kelas
 */
function setQuizClassStatus(className, isActive) {
  try {
    if (getQuizSchedules().length > 0) {
      return {
        success: false,
        error:
          "Tidak dapat mengubah status kelas secara manual saat jadwal otomatis aktif.",
      };
    }
    PropertiesService.getScriptProperties().setProperty(
      "QUIZ_STATUS_" + className,
      isActive.toString()
    );
    return { success: true };
  } catch (e) {
    Logger.log(e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * Menjalankan pengecekan jadwal lalu mengembalikan status terbaru.
 */
function runCheckAndGetStatuses() {
  checkScheduledQuizStatus(); // Jalankan pengecekan status terlebih dahulu
  return getQuizClassStatuses(); // Lalu kembalikan status yang sudah diperbarui
}

/**
 * Mendapatkan semua status kuis per kelas
 */
function getQuizClassStatuses() {
  try {
    const statuses = {};
    const allProps = PropertiesService.getScriptProperties().getProperties();
    CLASS_NAMES.forEach((className) => {
      const status = allProps["QUIZ_STATUS_" + className];
      statuses[className] = status === "true";
    });
    return { statuses: statuses };
  } catch (e) {
    Logger.log(e.stack);
    return { error: e.message };
  }
}

/**
 * Mendapatkan pengaturan aplikasi untuk dashboard.
 */
function getAppSettings() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheets = ss.getSheets();
    const sheetNames = allSheets
      .map((sheet) => sheet.getName())
      .filter(
        (name) => name !== RESPONSE_SHEET_NAME && name !== SCHEDULE_SHEET_NAME
      );

    const activeSheet = getActiveQuestionSheetName();

    return {
      sheets: sheetNames,
      activeSheet: activeSheet,
    };
  } catch (e) {
    Logger.log(e.stack);
    return { error: e.message };
  }
}

/**
 * Mengatur sheet soal mana yang aktif.
 */
function setActiveQuestionSheet(sheetName) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      "ACTIVE_QUESTION_SHEET",
      sheetName
    );
    return { success: true };
  } catch (e) {
    Logger.log(e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * Menggabungkan semua panggilan data untuk dashboard menjadi satu, dengan filter jadwal.
 */
function getDashboardInitialData() {
  try {
    checkScheduledQuizStatus();

    const resultsData = getQuizResults();
    const quizStatus = getQuizStatus();
    const classStatusesData = getQuizClassStatuses();
    const appSettingsData = getAppSettings();
    const allSchedules = getQuizSchedules();
    const activeGrade = getActiveGrade();

    const schedulesToDisplay = filterSchedulesForDisplay(allSchedules);

    if (resultsData.error || classStatusesData.error || appSettingsData.error) {
      throw new Error("Gagal mengambil salah satu set data dashboard.");
    }

    return {
      results: resultsData.results,
      quizStatus: quizStatus,
      classStatuses: classStatusesData.statuses,
      appSettings: appSettingsData,
      schedules: schedulesToDisplay,
      classNames: CLASS_NAMES,
      activeGrade: activeGrade,
    };
  } catch (e) {
    Logger.log(e.stack);
    return { error: e.message };
  }
}

/**
 * Ekstrak gambar dari file .xlsx dan upload ke Google Drive
 */
function extractImagesFromXlsx(blob) {
  try {
    const zipBlob = blob.copyBlob().setContentType("application/zip");
    const files = Utilities.unzip(zipBlob);
    
    // 1. Upload media files
    const mediaFiles = files.filter(f => f.getName().indexOf('xl/media/') === 0);
    if (mediaFiles.length === 0) return {}; 
    
    const targetFolderId = UPLOAD_FOLDER_ID; // Menggunakan folder ID yang sudah ada di config
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const mediaPathToDriveId = {};
    
    mediaFiles.forEach(file => {
      // Create file in Drive
      const driveFile = targetFolder.createFile(file);
      const filename = file.getName().split('/').pop(); 
      mediaPathToDriveId["../media/" + filename] = driveFile.getId();
    });
    
    // 2. Parse drawing rels
    const cellImageMap = {};
    const relsFiles = files.filter(f => f.getName().indexOf('xl/drawings/_rels/') === 0);
    const drawingFiles = files.filter(f => f.getName().indexOf('xl/drawings/') === 0 && f.getName().endsWith('.xml') && f.getName().indexOf('_rels') === -1);
    
    for (let i = 0; i < drawingFiles.length; i++) {
      const drawingFile = drawingFiles[i];
      const drawingName = drawingFile.getName().split('/').pop();
      const relsName = drawingName + ".rels";
      const relsFile = relsFiles.find(f => f.getName().endsWith(relsName));
      
      let rIdToDriveId = {};
      if (relsFile) {
        try {
          const relsXml = XmlService.parse(relsFile.getDataAsString());
          const root = relsXml.getRootElement();
          const ns = root.getNamespace();
          const rels = root.getChildren('Relationship', ns);
          rels.forEach(rel => {
            const id = rel.getAttribute('Id').getValue();
            const target = rel.getAttribute('Target').getValue(); // e.g. "../media/image1.png"
            
            const targetFilename = target.split('/').pop();
            const key = "../media/" + targetFilename;
            if (mediaPathToDriveId[key]) {
              rIdToDriveId[id] = mediaPathToDriveId[key];
            }
          });
        } catch(e) { Logger.log("Rels parse error: " + e); }
      }
      
      // 3. Parse drawing file
      try {
        const xmlString = drawingFile.getDataAsString();
        const drawingXml = XmlService.parse(xmlString);
        const root = drawingXml.getRootElement();
        const xdrNs = XmlService.getNamespace("http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        const aNs = XmlService.getNamespace("http://schemas.openxmlformats.org/drawingml/2006/main");
        const rNs = XmlService.getNamespace("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        
        const anchors = root.getChildren(); 
        anchors.forEach(anchor => {
          let from = anchor.getChild('from', xdrNs);
          if (!from) return;
          
          let col = from.getChild('col', xdrNs).getText();
          let row = from.getChild('row', xdrNs).getText();
          
          let pic = anchor.getChild('pic', xdrNs);
          if (!pic) return;
          
          let blipFill = pic.getChild('blipFill', xdrNs);
          if (!blipFill) return;
          
          let blip = blipFill.getChild('blip', aNs);
          if (!blip) return;
          
          let embedAttr = blip.getAttribute('embed', rNs);
          if (embedAttr) {
            let rId = embedAttr.getValue();
            if (rIdToDriveId[rId]) {
               cellImageMap[row + "_" + col] = rIdToDriveId[rId];
            }
          }
        });
      } catch(e) { Logger.log("Drawing parse error: " + e); }
    }
    
    return cellImageMap;
  } catch (e) {
    Logger.log("extractImagesFromXlsx error: " + e);
    return {};
  }
}
