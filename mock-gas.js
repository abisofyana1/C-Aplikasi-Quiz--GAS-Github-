/**
 * mock-gas.js
 * Mock adapter untuk menjalankan aplikasi Google Apps Script di GitHub Pages / Static Server
 * Menyediakan objek dummy google.script.run dan penanganan navigasi halaman.
 */

(function () {
  const MOCK_STORAGE_KEY_PREFIX = "quiz_app_demo_";

  function getStored(key, defaultVal) {
    try {
      const val = localStorage.getItem(MOCK_STORAGE_KEY_PREFIX + key);
      return val ? JSON.parse(val) : defaultVal;
    } catch (e) {
      return defaultVal;
    }
  }

  function setStored(key, val) {
    try {
      localStorage.setItem(MOCK_STORAGE_KEY_PREFIX + key, JSON.stringify(val));
    } catch (e) {}
  }

  // Inisialisasi Data Demo
  const defaultClassesX = ["E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9"];
  const defaultClassesXII = ["XII 1", "XII 2"];
  const allDefaultClasses = [...defaultClassesX, ...defaultClassesXII];

  const defaultQuestions = [
    {
      question: "Perhatikan cuplikan kode JavaScript berikut. Manakah tipe data dari nilai `[1, 2, 3]`?",
      options: {
        A: "Object (Array)",
        B: "String",
        C: "Boolean",
        D: "Number",
        E: "Undefined"
      },
      imageId: null,
      time: 90
    },
    {
      question: "Dalam ekosistem Google Apps Script, layanan apakah yang digunakan untuk memanipulasi Google Spreadsheet?",
      options: {
        A: "DriveApp",
        B: "SpreadsheetApp",
        C: "DocumentApp",
        D: "GmailApp",
        E: "FormApp"
      },
      imageId: null,
      time: 90
    },
    {
      question: "Komponen utama yang digunakan untuk menghosting website statis langsung dari repositori GitHub secara gratis adalah...",
      options: {
        A: "GitHub Copilot",
        B: "GitHub Gist",
        C: "GitHub Pages",
        D: "GitHub Codespaces",
        E: "GitHub Actions"
      },
      imageId: null,
      time: 90
    },
    {
      question: "Manakah method JavaScript yang digunakan untuk menyimpan data dalam bentuk string ke penyimpanan lokal browser?",
      options: {
        A: "localStorage.getItem()",
        B: "localStorage.clear()",
        C: "localStorage.setItem()",
        D: "sessionStorage.remove()",
        E: "window.saveData()"
      },
      imageId: null,
      time: 60
    },
    {
      question: "Apa fungsi dari tag `<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">` pada HTML?",
      options: {
        A: "Mengatur judul halaman web",
        B: "Menghubungkan file CSS eksternal",
        C: "Mengoptimalkan tampilan web agar responsif di perangkat mobile",
        D: "Menonaktifkan javascript di browser",
        E: "Mengatur kata sandi aplikasi"
      },
      imageId: null,
      time: 60
    }
  ];

  const questionAnswerKeys = ["A", "B", "C", "C", "C"]; // Kunci jawaban untuk dummy questions (skor 20/soal, total 100)

  const defaultResults = [
    {
      rowIndex: 2,
      "Waktu Pengerjaan": "19/10/2026, 08.15.22",
      "Alamat IP": "182.253.112.45",
      "Nama": "AHMAD BUDI SANTOSO",
      "Kelas": "E1",
      "Skor": "100/100"
    },
    {
      rowIndex: 3,
      "Waktu Pengerjaan": "19/10/2026, 08.18.05",
      "Alamat IP": "182.253.112.46",
      "Nama": "CITRA LESTARI DEWI",
      "Kelas": "E1",
      "Skor": "80/100"
    },
    {
      rowIndex: 4,
      "Waktu Pengerjaan": "19/10/2026, 08.20.40",
      "Alamat IP": "182.253.114.12",
      "Nama": "DIMAS PRASETYA",
      "Kelas": "E2",
      "Skor": "100/100"
    },
    {
      rowIndex: 5,
      "Waktu Pengerjaan": "19/10/2026, 08.22.15",
      "Alamat IP": "182.253.114.15",
      "Nama": "FANI ANGGRAINI",
      "Kelas": "E3",
      "Skor": "60/100"
    },
    {
      rowIndex: 6,
      "Waktu Pengerjaan": "19/10/2026, 08.25.00",
      "Alamat IP": "182.253.115.89",
      "Nama": "GILANG RAMADHAN",
      "Kelas": "XII 1",
      "Skor": "100/100"
    },
    {
      rowIndex: 7,
      "Waktu Pengerjaan": "19/10/2026, 08.27.33",
      "Alamat IP": "182.253.115.90",
      "Nama": "HANA SALSABILA",
      "Kelas": "XII 2",
      "Skor": "80/100"
    }
  ];

  // Inisialisasi storage jika belum ada
  if (!localStorage.getItem(MOCK_STORAGE_KEY_PREFIX + "initialized")) {
    setStored("results", defaultResults);
    setStored("activeGrade", "X");
    setStored("quizStatus", "active");
    const classStatuses = {};
    allDefaultClasses.forEach(c => {
      classStatuses[c] = true;
    });
    setStored("classStatuses", classStatuses);
    localStorage.setItem(MOCK_STORAGE_KEY_PREFIX + "initialized", "true");
  }

  // Gantikan scriptlet URL links agar link Kuis & Dashboard bekerja di static hosting
  document.addEventListener("DOMContentLoaded", () => {
    const isDashboardPage = window.location.pathname.toLowerCase().includes("dashboard") || window.location.search.includes("page=dashboard");
    
    // Ganti tautan navigasi
    document.querySelectorAll("a").forEach(a => {
      const href = a.getAttribute("href") || "";
      if (href.includes("page=quiz") || href.includes("<?= url ?>?page=quiz")) {
        a.setAttribute("href", "index.html");
      } else if (href.includes("page=dashboard") || href.includes("<?= url ?>?page=dashboard")) {
        a.setAttribute("href", "dashboard.html");
      }
    });

    // Tampilkan banner notifikasi Demo
    const demoBanner = document.createElement("div");
    demoBanner.style.cssText = `
      background: linear-gradient(90deg, #1e3a8a, #3b82f6);
      color: #ffffff;
      padding: 8px 16px;
      font-size: 13px;
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
      position: sticky;
      top: 0;
      z-index: 99999;
      box-shadow: 0 2px 4px rgba(0,0,0,0.15);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    `;
    demoBanner.innerHTML = `
      <span><strong>Live Demo Mode:</strong> Data simulasi kuis offline (Password Dashboard: <code style="background:rgba(255,255,255,0.25);padding:2px 6px;border-radius:4px;font-weight:bold;">demo</code>).</span>
      <button id="demo-reset-session-btn" style="background:#ef4444;color:#fff;border:none;padding:3px 10px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;">Kembali ke Halaman Awal / Reset</button>
    `;
    document.body.insertBefore(demoBanner, document.body.firstChild);

    document.getElementById("demo-reset-session-btn")?.addEventListener("click", () => {
      localStorage.removeItem("studentInfo");
      localStorage.removeItem("quizProgress");
      localStorage.removeItem("retakeRequestSent");
      window.location.href = "index.html";
    });
  });

  // Handler Factory untuk google.script.run
  class MockRunner {
    constructor(successHandler = null, failureHandler = null) {
      this._successHandler = successHandler || (() => {});
      this._failureHandler = failureHandler || ((err) => console.error("Mock GAS Error:", err));
    }

    withSuccessHandler(handler) {
      return new MockRunner(handler, this._failureHandler);
    }

    withFailureHandler(handler) {
      return new MockRunner(this._successHandler, handler);
    }

    _asyncResolve(data) {
      setTimeout(() => {
        try {
          this._successHandler(data);
        } catch (e) {
          console.error("Error inside successHandler:", e);
        }
      }, 150);
    }

    checkPassword(pass) {
      // Password Live Demo yang dibuat
      const isValid = (pass === "demo");
      this._asyncResolve(isValid);
    }

    getQuizPageData() {
      const activeGrade = getStored("activeGrade", "X");
      const classes = activeGrade === "X" ? defaultClassesX : defaultClassesXII;
      this._asyncResolve({
        classes: classes,
        quizTime: 90
      });
    }

    getQuizData() {
      const quizStatus = getStored("quizStatus", "active");
      this._asyncResolve({
        questions: defaultQuestions,
        images: {},
        quizStatus: quizStatus
      });
    }

    getQuizGlobalStatus() {
      const status = getStored("quizStatus", "active");
      this._asyncResolve(status);
    }

    checkStudentCompletion(student) {
      const classStatuses = getStored("classStatuses", {});
      const isClassActive = classStatuses[student.studentClass] !== false;
      const results = getStored("results", []);

      const existing = results.find(r => 
        String(r.Nama).trim().toUpperCase() === String(student.name).trim().toUpperCase() &&
        String(r.Kelas).trim() === String(student.studentClass).trim()
      );

      if (existing) {
        const parts = String(existing.Skor).split("/");
        this._asyncResolve({
          completed: true,
          active: isClassActive,
          score: parts[0] ? parts[0].trim() : 0,
          total: parts[1] ? parts[1].trim() : 100
        });
      } else {
        this._asyncResolve({
          completed: false,
          active: isClassActive
        });
      }
    }

    checkAnswers(submissionData) {
      const { answers, name, class: studentClass, ipAddress } = submissionData;
      let userScore = 0;
      let maxScore = 100;
      const scorePerQuestion = 20;

      defaultQuestions.forEach((q, idx) => {
        const userAns = answers[idx] ? answers[idx].answer : null;
        if (userAns && userAns === questionAnswerKeys[idx]) {
          userScore += scorePerQuestion;
        }
      });

      const now = new Date();
      const timeStr = now.toLocaleDateString("id-ID") + ", " + now.toLocaleTimeString("id-ID").replace(/:/g, ".");
      const currentResults = getStored("results", []);

      const newRecord = {
        rowIndex: currentResults.length + 2,
        "Waktu Pengerjaan": timeStr,
        "Alamat IP": ipAddress || "127.0.0.1 (Demo)",
        "Nama": name.toUpperCase(),
        "Kelas": studentClass,
        "Skor": `${userScore}/${maxScore}`
      };

      currentResults.unshift(newRecord);
      setStored("results", currentResults);

      this._asyncResolve({
        score: userScore,
        total: maxScore
      });
    }

    sendPermissionRequestEmail(studentDetails) {
      this._asyncResolve({
        success: true,
        message: "Permintaan demo berhasil disimulasikan."
      });
    }

    getDashboardInitialData() {
      const results = getStored("results", defaultResults);
      const quizStatus = getStored("quizStatus", "active");
      const classStatuses = getStored("classStatuses", {});
      allDefaultClasses.forEach(c => {
        if (classStatuses[c] === undefined) classStatuses[c] = true;
      });
      const activeGrade = getStored("activeGrade", "X");

      const now = new Date();
      const dummySchedules = [
        {
          id: "SCHED_001",
          sheetName: "Soal Pemrograman Web",
          startTime: new Date(now.getTime() - 3600000).toISOString(),
          endTime: new Date(now.getTime() + 7200000).toISOString(),
          classNames: ["E1", "E2", "E3"]
        }
      ];

      this._asyncResolve({
        results: results,
        quizStatus: quizStatus,
        classStatuses: classStatuses,
        appSettings: {
          sheets: ["Soal Pemrograman Web", "Soal Desain Grafis"]
        },
        schedules: dummySchedules,
        classNames: allDefaultClasses,
        activeGrade: activeGrade
      });
    }

    runCheckAndGetStatuses() {
      const classStatuses = getStored("classStatuses", {});
      allDefaultClasses.forEach(c => {
        if (classStatuses[c] === undefined) classStatuses[c] = true;
      });
      this._asyncResolve({ statuses: classStatuses });
    }

    setQuizStatus(status) {
      setStored("quizStatus", status);
      this._asyncResolve({ success: true, status: status });
    }

    setQuizClassStatus(className, isActive) {
      const classStatuses = getStored("classStatuses", {});
      classStatuses[className] = isActive;
      setStored("classStatuses", classStatuses);
      this._asyncResolve({ success: true });
    }

    setActiveGrade(grade) {
      setStored("activeGrade", grade);
      this._asyncResolve({ success: true, activeGrade: grade });
    }

    resetSingleStudentData(rowIndex) {
      let results = getStored("results", []);
      results = results.filter(r => r.rowIndex !== rowIndex);
      setStored("results", results);
      this._asyncResolve({ success: true });
    }

    deleteQuizSchedule(scheduleId) {
      this._asyncResolve({ success: true, schedules: [] });
    }

    deleteAllSchedules() {
      this._asyncResolve({ success: true });
    }

    saveQuizSchedule(scheduleData) {
      this._asyncResolve({ success: true });
    }

    uploadSoalFile(fileData, sheetName) {
      this._asyncResolve({ success: true });
    }
  }

  // Pasang mock ke window jika belum ada google.script.run (misal saat running di GitHub Pages / Static Server)
  if (typeof window.google === "undefined" || !window.google.script || !window.google.script.run) {
    window.google = window.google || {};
    window.google.script = window.google.script || {};
    window.google.script.run = new MockRunner();
  }
})();
