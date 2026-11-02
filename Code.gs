// Jalankan setupLdpSystem dari dropdown editor. Proyek standalone memakai
// Script Property LDP_SPREADSHEET_ID (ID atau URL spreadsheet database).
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("LDP Admin")
    .addItem("Setup / Perbarui sistem", "setupLdpSystem")
    .addItem("Provision akun dari User Import", "provisionUsersFromImportSheet")
    .addToUi();
}
function setupSpreadsheetId_() {
  var value = String(props_().getProperty("LDP_SPREADSHEET_ID") || "").trim();
  if (!value) return "";
  var url = value.match(/^https:\/\/docs\.google\.com\/spreadsheets\/(?:u\/\d+\/)?d\/([A-Za-z0-9_-]+)(?:[\/?#]|$)/i);
  if (url) return url[1];
  if (!/^[A-Za-z0-9_-]{20,}$/.test(value))
    throw new Error("LDP_SPREADSHEET_ID harus berisi ID atau URL Google Sheets database LDP, bukan URL proyek Apps Script.");
  return value;
}
function setupSpreadsheet_() {
  var id = setupSpreadsheetId_(), active = SpreadsheetApp.getActiveSpreadsheet();
  if (id && active && active.getId() !== id)
    throw new Error("Spreadsheet aktif berbeda dari LDP_SPREADSHEET_ID. Periksa ID database pada Project Settings > Script properties sebelum menjalankan setup.");
  if (id) {
    try { return SpreadsheetApp.openById(id); }
    catch (error) {
      throw new Error("Database LDP tidak dapat dibuka. Periksa LDP_SPREADSHEET_ID dan akses edit akun Google ke spreadsheet tersebut. Detail: " + error.message);
    }
  }
  if (active) return active;
  throw new Error("Isi LDP_SPREADSHEET_ID di Project Settings > Script properties dengan ID atau URL spreadsheet database LDP, lalu jalankan setupLdpSystem kembali.");
}
function requireSpreadsheetEditor_() {
  // getUi is available for maintenance launched from a bound spreadsheet.
  var boundEditor = false;
  try { SpreadsheetApp.getUi(); boundEditor = true; } catch (error) {}
  if (boundEditor) return;
  // Never use effectiveUser as caller identity: web apps may run as their owner.
  var email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!email)
    throw new Error("Akun Google pemanggil tidak dapat diverifikasi. Jalankan fungsi dari editor Apps Script dengan akun pemilik spreadsheet database, bukan dari halaman aplikasi.");
  var ss = setupSpreadsheet_();
  var allowed = String(props_().getProperty("LDP_MAINTENANCE_EMAILS") || "")
    .split(/[,;\n]/).map(function (value) { return value.trim().toLowerCase(); }).filter(Boolean);
  if (allowed.indexOf(email) >= 0) return;
  var owner = DriveApp.getFileById(ss.getId()).getOwner();
  if (owner && String(owner.getEmail() || "").trim().toLowerCase() === email) return;
  throw new Error("Setup hanya untuk pemilik spreadsheet database atau email pada Script Property LDP_MAINTENANCE_EMAILS. Gunakan akun Google yang benar; properti tambahan hanya diisi oleh pengelola proyek.");
}
function setupLdpSystem() {
  requireSpreadsheetEditor_();
  var result = setupLdpSystem_();
  console.log(result.message);
  if (result.status === "FAILED") throw new Error(result.message);
  return result;
}
function provisionUsersFromImportSheet() {
  requireSpreadsheetEditor_();
  return provisionUsersFromImportSheet_();
}
function initializeAdminPassword() {
  requireSpreadsheetEditor_();
  return initializeAdminPassword_();
}

/**
 * LDP Monitor 2.0 — optimized Google Apps Script V8 backend
 * Entry points are at the end; internal/editor functions end in _.
 * Install using PANDUAN.md. A 200-user live load test remains required.
 */

var SHEET_SETTINGS = "Settings";
var SHEET_USERS = "Users";
var SHEET_USER_IMPORT = "User Import";
var SHEET_PARTICIPANTS = "Participants";
var SHEET_ENROLLMENTS = "Program Enrollments";
var SHEET_QLEAP = "QLEAP Mapping";
var SHEET_TIMELINE = "Timeline"; // legacy only
var SHEET_SUBMISSIONS = "Project Submissions";
var SHEET_REVIEWS = "Slide Reviews"; // legacy only
var SHEET_MENTORING = "Mentoring Evidence";
var SHEET_CERTIFICATES = "Certificates";
var SHEET_CERT_IMPORT = "Certificate Import";
var SHEET_AUDIT = "Audit Log";
var SHEET_PANEL = "Panel Assessments";
var SHEET_QLEAP_POSTTEST = "QLEAP Post Test"; // NEW

var CERT_PROGRAM_ORDER = ["ASHDP", "TSHDP", "SLDP"];

var ROLE_PARTICIPANT = "PARTICIPANT";
var ROLE_SUPERIOR = "SUPERIOR";
var ROLE_CLP = "CLP";
var ROLE_ADMIN = "ADMIN";
var ALL_ROLES = [ROLE_PARTICIPANT, ROLE_SUPERIOR, ROLE_CLP, ROLE_ADMIN];

var TZ = "Asia/Jakarta";
var PASSING_SCORE = 80;

var PARTICIPANT_IMPORT_HEADERS = ["NIK", "Nama", "Job Title", "Region", "Program", "Batch", "Segmen"];
var PARTICIPANT_STATUS_OPTIONS = [
  "ELIGIBLE",
  "NOT ELIGIBLE",
  "PROMOSI",
  "RESIGN",
  "BYPASS",
];
var CURRENT_PHASE_OPTIONS = [
  "PHASE 1",
  "PHASE 2",
  "PHASE 3",
  "PANELISASI",
  "COMPLETED",
];
var PROGRAM_TRACK_OPTIONS = ["STANDARDISASI", "TALENT", "BOTH"]; // NEW

var ALLOWED_UPLOAD_MIME = {
  "application/vnd.ms-powerpoint": true,
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": true,
  "application/pdf": true,
  "image/png": true,
  "image/jpeg": true,
};

var SHEET_HEADERS = {};
SHEET_HEADERS[SHEET_SETTINGS] = ["Key", "Value", "Description"];
SHEET_HEADERS[SHEET_USERS] = [
  "NIK",
  "Name",
  "Role",
  "Region",
  "Superior_NIK",
  "Email",
  "PasswordHash",
  "Salt",
  "Active",
  "CreatedAt",
  "UpdatedAt",
];
SHEET_HEADERS[SHEET_USER_IMPORT] = [
  "NIK",
  "Name",
  "Role",
  "Region",
  "Superior_NIK",
  "Email",
  "Active",
  "Password",
];
SHEET_HEADERS[SHEET_PARTICIPANTS] = [
  "NIK",
  "Name",
  "JoinDate",
  "JobTitle",
  "Region",
  "City",
  "Email",
  "Program",
  "Batch",
  "Phase1Score",
  "Phase1Status",
  "Phase2Score",
  "Phase2Status",
  "Phase3Score",
  "Phase3Status",
  "ImprovementObservation",
  "Theme",
  "ProjectTitle",
  "ActionPlan",
  "Implementation",
  "PanelDate",
  "ParticipantStatus",
  "OverallProgress",
  "BehaviourObservation",
  "PanelScore",
  "FinalScore",
  "FinalStatus",
  "CertificateUrl",
  "SlideUrl",
  "Superior_NIK",
  "CLP_NIK",
  "UpdatedAt",
  "Superior_Name",
  "Location",
  "CityName",
  "CurrentPhase",
  "CurrentEnrollmentId",
  "ProgramTrack",
];
SHEET_HEADERS[SHEET_ENROLLMENTS] = [
  "EnrollmentID",
  "NIK",
  "Program",
  "Batch",
  "Region",
  "CurrentPhase",
  "EnrollmentStatus",
  "ProgramStatus",
  "IsActive",
  "Phase1Score",
  "Phase2Score",
  "Phase3Score",
  "FinalScore",
  "CertificateUrl",
  "StartedAt",
  "CompletedAt",
  "CreatedAt",
  "UpdatedAt",
];
SHEET_HEADERS[SHEET_QLEAP] = [
  "Program",
  "Phase",
  "Module",
  "QLEAPUrl",
  "UpdatedBy",
  "UpdatedAt",
];
SHEET_HEADERS[SHEET_SUBMISSIONS] = [
  "SubmissionID",
  "NIK",
  "Phase",
  "Version",
  "FileName",
  "FileUrl",
  "DriveFileId",
  "SubmittedAt",
  "Status",
  "Deadline",
];
SHEET_HEADERS[SHEET_REVIEWS] = [
  "ReviewID",
  "SubmissionID",
  "NIK",
  "Phase",
  "SlideNo",
  "SlideTopic",
  "ReviewerNIK",
  "ReviewerRole",
  "ReviewStatus",
  "Feedback",
  "UpdatedAt",
];
SHEET_HEADERS[SHEET_MENTORING] = [
  "EvidenceID",
  "NIK",
  "MentorRole",
  "MentorNIK",
  "MentoringDate",
  "Notes",
  "FileName",
  "FileUrl",
  "DriveFileId",
  "VerificationStatus",
  "VerifiedBy",
  "VerifiedAt",
  "Feedback",
];
SHEET_HEADERS[SHEET_CERTIFICATES] = [
  "CertificateID",
  "NIK",
  "FinalScore",
  "FinalStatus",
  "FileOrLink",
  "Published",
  "UploadedBy",
  "UploadedAt",
  "CertificateNumber",
  "DisplayName",
  "Program",
  "GeneratedAt",
  "IssueDate",
];
SHEET_HEADERS[SHEET_CERT_IMPORT] = [
  "ImportID",
  "BatchID",
  "RowNo",
  "NIK",
  "DisplayName",
  "Batch",
  "Program",
  "FinalStatus",
  "FinalScore",
  "CertificateNumber",
  "IssueDate",
  "RowStatus",
  "Message",
  "FileOrLink",
  "UploadedBy",
  "UploadedAt",
  "GeneratedAt",
];
SHEET_HEADERS[SHEET_AUDIT] = [
  "Timestamp",
  "NIK",
  "Role",
  "Action",
  "Entity",
  "EntityID",
  "Detail",
];
SHEET_HEADERS[SHEET_PANEL] = [
  "AssessmentID",
  "NIK",
  "AssessorNIK",
  "AssessorName",
  "AssessorRole",
  "PanelDate",
  "Scores",
  "TotalScore",
  "Recommendation",
  "Notes",
  "CreatedAt",
  "UpdatedAt",
];
SHEET_HEADERS[SHEET_QLEAP_POSTTEST] = [
  "UploadBatchID",
  "NIK",
  "EmployeeName",
  "JobTitle",
  "WorkLocation",
  "BusinessUnit",
  "ElearningStatus",
  "Modules",
  "ExamName",
  "Score",
  "Duration",
  "ExamDate",
  "Status",
  "UploadedBy",
  "UploadedAt",
];

var USER_IMPORT_ROLE_ALIASES = {
  PARTICIPANT: ["PARTICIPANT", "PESERTA"],
  SUPERIOR: ["SUPERIOR", "ATASAN"],
  CLP: ["CLP", "TRAINER", "CORPORATE LEARNING PARTNER"],
  ADMIN: ["ADMIN", "ADMINISTRATOR"],
};
var PROGRAM_PHASE_ASSESSMENT_CONFIG = {
  "TSHDP|PHASE 1": {
    program: "TSHDP",
    phase: "PHASE 1",
    settingKey: "PHASE1_TSHDP_SHEET",
    defaultSheet: "Phase 1 TSHDP",
    headerRow: 2,
    dataStartRow: 3,
    nikHeaders: ["NIK", "Employee ID"],
    batchHeaders: ["Batch"],
    headers: [
      "Region",
      "NIK",
      "Nama Karyawan",
      "Batch",
      "Goal Setting & Strategy",
      "Analisis Data Operasional Depo",
      "Market Research",
      "Service Excellence",
      "Leadership & Pengembangan Tim",
      "Skala Prioritas",
    ],
    modules: [
      {
        label: "Goal Setting & Strategy",
        headers: ["Goal Setting & Strategy"],
      },
      {
        label: "Analisis Data Operasional Depo",
        headers: ["Analisis Data Operasional Depo"],
      },
      {
        label: "Market Research",
        headers: ["Market Research"],
      },
      {
        label: "Service Excellence",
        headers: ["Service Excellence"],
      },
      {
        label: "Leadership & Pengembangan Tim",
        headers: ["Leadership & Pengembangan Tim"],
      },
      {
        label: "Skala Prioritas",
        headers: ["Skala Prioritas"],
      },
    ],
  },
  "TSHDP|PHASE 2": {
    program: "TSHDP",
    phase: "PHASE 2",
    settingKey: "PHASE2_TSHDP_SHEET",
    defaultSheet: "Phase 2 TSHDP",
    headerRow: 1,
    dataStartRow: 2,
    nikHeaders: ["NIK", "Employee ID"],
    batchHeaders: ["Batch"],
    headers: [
      "NO",
      "Region",
      "NIK",
      "Nama Karyawan",
      "Batch",
      "In Class Training Score Business Process Analisyst",
      "In Class Training Score Sales Forecasting",
      "In Class Training Score Profiling Dealer",
      "In Class Training Score Monitor Dealer",
      "In Class Training Score Solution and Recovery",
      "In Class Training Score Market Research",
      "In Class Training Score Negotiation Skill",
      "In Class Training Score TAM POS",
      "Count Phase 2 - Self Learning",
      "% QTY Pengerjaan Phase 2",
      "Avarage Score Phase 2 - Self Learning",
    ],
    modules: [
      {
        label: "Business Process Analyst",
        headers: [
          "In Class Training Score Business Process Analisyst",
          "In Class Training Score Business Process Analyst",
        ],
      },
      {
        label: "Sales Forecasting",
        headers: ["In Class Training Score Sales Forecasting"],
      },
      {
        label: "Profiling Dealer",
        headers: ["In Class Training Score Profiling Dealer"],
      },
      {
        label: "Monitor Dealer",
        headers: ["In Class Training Score Monitor Dealer"],
      },
      {
        label: "Solution & Recovery",
        headers: ["In Class Training Score Solution and Recovery"],
      },
      {
        label: "Market Research",
        headers: ["In Class Training Score Market Research"],
      },
      {
        label: "Negotiation Skill",
        headers: ["In Class Training Score Negotiation Skill"],
      },
      {
        label: "TAM POS",
        headers: ["In Class Training Score TAM POS"],
      },
    ],
  },
  "ASHDP|PHASE 1": {
    program: "ASHDP",
    phase: "PHASE 1",
    settingKey: "PHASE1_ASHDP_SHEET",
    defaultSheet: "Phase 1 ASHDP",
    headerRow: 1,
    dataStartRow: 2,
    nikHeaders: ["Employee ID", "NIK"],
    batchHeaders: ["Batch"],
    headers: [
      "Employee ID",
      "Employee Name",
      "Job Title Name",
      "Location Group Name",
      "Location City Name",
      "Batch",
      "Strategic Thinking",
      "Strategic Leadership",
      "Developing Others",
      "Leading Others",
      "Reliable Partner",
      "Quality Oriented",
      "Achievement Oriented",
      "Creative Agility",
      "Technology Savvy",
      "Average Score",
      "QTY Pengerjaan",
      "% QTY Pengerjaan",
    ],
    modules: [
      {
        label: "Strategic Thinking",
        headers: ["Strategic Thinking"],
      },
      {
        label: "Strategic Leadership",
        headers: ["Strategic Leadership"],
      },
      {
        label: "Developing Others",
        headers: ["Developing Others"],
      },
      {
        label: "Leading Others",
        headers: ["Leading Others"],
      },
      {
        label: "Reliable Partner",
        headers: ["Reliable Partner"],
      },
      {
        label: "Quality Oriented",
        headers: ["Quality Oriented"],
      },
      {
        label: "Achievement Oriented",
        headers: ["Achievement Oriented"],
      },
      {
        label: "Creative Agility",
        headers: ["Creative Agility"],
      },
      {
        label: "Technology Savvy",
        headers: ["Technology Savvy"],
      },
    ],
  },
  "ASHDP|PHASE 2": {
    program: "ASHDP",
    phase: "PHASE 2",
    settingKey: "PHASE2_ASHDP_SHEET",
    defaultSheet: "Phase 2 ASHDP",
    headerRow: 1,
    dataStartRow: 2,
    nikHeaders: ["Employee ID", "NIK"],
    batchHeaders: ["Batch"],
    headers: [
      "Employee ID",
      "Employee Name",
      "Job Title Name",
      "Location Group Name",
      "Location City Name",
      "Batch",
      "Analisa Potensi Pasar",
      "Profiling Dealer",
      "Monitor Your Dealer",
      "Solution & Recovery",
      "Prediksi Penjualan dalam Bisnis Distribusi",
      "Meningkatkan Efektifitas Proses Kerja",
      "TAM POS",
      "Membuat Laporan yang Efektif",
      "Dealer Retention",
      "Average Score",
      "QTY Pengerjaan",
      "% QTY Pengerjaan",
      "Status Kelulusan",
    ],
    modules: [
      {
        label: "Analisa Potensi Pasar",
        headers: ["Analisa Potensi Pasar"],
      },
      {
        label: "Profiling Dealer",
        headers: ["Profiling Dealer"],
      },
      {
        label: "Monitor Your Dealer",
        headers: ["Monitor Your Dealer"],
      },
      {
        label: "Solution & Recovery",
        headers: ["Solution & Recovery"],
      },
      {
        label: "Prediksi Penjualan dalam Bisnis Distribusi",
        headers: ["Prediksi Penjualan dalam Bisnis Distribusi"],
      },
      {
        label: "Meningkatkan Efektifitas Proses Kerja",
        headers: ["Meningkatkan Efektifitas Proses Kerja"],
      },
      {
        label: "TAM POS",
        headers: ["TAM POS"],
      },
      {
        label: "Membuat Laporan yang Efektif",
        headers: ["Membuat Laporan yang Efektif"],
      },
      {
        label: "Dealer Retention",
        headers: ["Dealer Retention"],
      },
    ],
  },
};

// Pure normalization, compatible import parsers, and assessment templates.
function resolveImportRole_(raw) {
  var s = String(raw || "")
    .trim()
    .toUpperCase();
  if (!s) return "";
  if (ALL_ROLES.indexOf(s) !== -1) return s;
  var resolved = "";
  Object.keys(USER_IMPORT_ROLE_ALIASES).forEach(function (canonical) {
    if (resolved) return;
    if (USER_IMPORT_ROLE_ALIASES[canonical].indexOf(s) !== -1)
      resolved = canonical;
  });
  return resolved;
}

function parseRoles_(raw) {
  return parseRolesStrict_(raw).roles;
}

function parseRolesStrict_(raw) {
  var roles = [],
    invalid = [];
  String(raw || "")
    .split(/[,;\/|+&]+/)
    .forEach(function (token) {
      var t = String(token).trim();
      if (!t) return;
      var r = resolveImportRole_(t);
      if (!r) {
        invalid.push(t);
        return;
      }
      if (roles.indexOf(r) === -1) roles.push(r);
    });
  return { roles: roles, invalid: invalid };
}

function joinRoles_(roles) {
  return (roles || []).join(",");
}

function userHasRole_(userData, role) {
  return parseRoles_(userData && userData.Role).indexOf(role) !== -1;
}

function normalizeNik_(value) {
  if (value === null || value === undefined) return "";
  var s = String(value).trim();
  if (s.indexOf(".") !== -1) s = s.replace(/\.0+$/, "");
  return s;
}

function normalizeBatchKey_(value) {
  var s = String(value === null || value === undefined ? "" : value)
    .trim()
    .toUpperCase();
  if (!s) return "";
  s = s.replace(/^BATCH\s*/i, "");
  s = s.replace(/#/g, " ");
  s = s
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!s) return "";
  return s
    .split(" ")
    .map(function (token) {
      return /^\d+$/.test(token) ? String(parseInt(token, 10)) : token;
    })
    .join(" ");
}

function sameBatch_(a, b) {
  var ka = normalizeBatchKey_(a);
  var kb = normalizeBatchKey_(b);
  if (!ka || !kb) return String(a || "").trim() === String(b || "").trim();
  return ka === kb;
}

function normalizeRegion_(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim().toUpperCase();
}

function parseBoolean_(value) {
  if (typeof value === "boolean") return value;
  var s = String(value).trim().toUpperCase();
  return (
    s === "TRUE" ||
    s === "1" ||
    s === "YES" ||
    s === "ACTIVE" ||
    s === "AKTIF" ||
    s === "Y"
  );
}

function formatDate_(d) {
  if (!(d instanceof Date)) return d;
  return Utilities.formatDate(d, TZ, "yyyy-MM-dd");
}

function formatDateTime_(d) {
  if (!(d instanceof Date)) return d;
  return Utilities.formatDate(d, TZ, "yyyy-MM-dd'T'HH:mm:ss");
}

function toIsoDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  var m = String(v || "").match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[1] + "-" + m[2] + "-" + m[3] : "";
}

function hashPassword_(password, salt) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + "::" + password,
  );
  return bytes
    .map(function (b) {
      return ("0" + (b & 0xff).toString(16)).slice(-2);
    })
    .join("");
}

function generateSalt_() {
  return Utilities.getUuid() + "-" + new Date().getTime();
}

function newId_(prefix) {
  return (
    prefix +
    "-" +
    Utilities.getUuid().replace(/-/g, "").substring(0, 12).toUpperCase()
  );
}

function sanitizeFileName_(name) {
  var base = String(name || "file")
    .replace(/[^a-zA-Z0-9._\- ]/g, "_")
    .trim();
  return base.length > 0 ? base : "file";
}

function requireFields_(obj, fields) {
  fields.forEach(function (f) {
    if (obj[f] === undefined || obj[f] === null || obj[f] === "") {
      throw new Error('Field "' + f + '" wajib diisi.');
    }
  });
}

function safeCell_(text) {
  var t = String(text || "");
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function normalizeParticipantStatus_(value) {
  var s = String(value || "")
    .trim()
    .toUpperCase();
  if (!s || s === "ACTIVE" || s === "AKTIF") return "ELIGIBLE";
  if (s === "PROMOTED" || s === "PROMOSI") return "PROMOSI";
  if (s === "RESIGNED" || s === "RESIGN") return "RESIGN";
  if (s.replace(/[\s_-]+/g, "") === "BYPASS") return "BYPASS";
  if (s === "NOT ELIGIBLE" || s === "NOT_ELIGIBLE" || s === "TIDAK ELIGIBLE")
    return "NOT ELIGIBLE";
  if (s === "ELIGIBLE") return "ELIGIBLE";
  return s;
}

function normalizeCurrentPhase_(value) {
  var s = String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
  if (!s) return "";
  if (s === "PANEL" || s === "PANELIZATION" || s === "PANELISATION")
    return "PANELISASI";
  if (s === "COMPLETE" || s === "DONE" || s === "FINISHED") return "COMPLETED";
  if (/^PHASE\s*1$/.test(s)) return "PHASE 1";
  if (/^PHASE\s*2$/.test(s)) return "PHASE 2";
  if (/^PHASE\s*3$/.test(s)) return "PHASE 3";
  return s;
}

function phaseRank_(phase) {
  var order = {
    "": 0,
    "PHASE 1": 1,
    "PHASE 2": 2,
    "PHASE 3": 3,
    PANELISASI: 4,
    COMPLETED: 5,
  };
  return order[normalizeCurrentPhase_(phase)] || 0;
}

function progressPercentFromPhase_(phase) {
  var p = normalizeCurrentPhase_(phase);
  return (
    {
      "": 0,
      "PHASE 1": 25,
      "PHASE 2": 50,
      "PHASE 3": 75,
      PANELISASI: 90,
      COMPLETED: 100,
    }[p] || 0
  );
}

function buildProgressStages_(currentPhase) {
  var phases = ["PHASE 1", "PHASE 2", "PHASE 3", "PANELISASI"];
  var current = normalizeCurrentPhase_(currentPhase);
  if (current === "COMPLETED") {
    return phases.map(function (p) {
      return { stage: p, status: "COMPLETED" };
    });
  }
  var idx = phases.indexOf(current);
  return phases.map(function (p, i) {
    if (idx < 0) return { stage: p, status: "NOT STARTED" };
    if (i < idx) return { stage: p, status: "COMPLETED" };
    if (i === idx) return { stage: p, status: "ON GOING" };
    return { stage: p, status: "NOT STARTED" };
  });
}

var TSHDP_PHASE1_WEIGHTS_ = [0.20, 0.15, 0.15, 0.15, 0.20, 0.15];
PROGRAM_PHASE_ASSESSMENT_CONFIG["TSHDP|PHASE 1"].modules.forEach(function(m, i) { m.weight = TSHDP_PHASE1_WEIGHTS_[i]; });
// SLDP initially follows ASHDP. B2C TSHDP keeps its own sources and history.
["PHASE 1","PHASE 2"].forEach(function(phase) {
  var sldp=JSON.parse(JSON.stringify(PROGRAM_PHASE_ASSESSMENT_CONFIG["ASHDP|"+phase])),
    tshdp=JSON.parse(JSON.stringify(PROGRAM_PHASE_ASSESSMENT_CONFIG["TSHDP|"+phase])),num=phase.slice(-1);
  sldp.program="SLDP";sldp.segment="B2C";sldp.settingKey="PHASE"+num+"_SLDP_SHEET";sldp.defaultSheet="Phase "+num+" SLDP B2C";
  tshdp.segment="B2C";tshdp.settingKey="PHASE"+num+"_B2C_TSHDP_SHEET";tshdp.defaultSheet="Phase "+num+" TSHDP B2C";
  PROGRAM_PHASE_ASSESSMENT_CONFIG["SLDP|"+phase]=sldp;
  PROGRAM_PHASE_ASSESSMENT_CONFIG["B2C|TSHDP|"+phase]=tshdp;
});
function assessmentAverage_(cfg, modules) {
  var valid = modules.filter(function(m) { return m.score !== ""; });
  if (!valid.length) return "";
  if (cfg.program === "TSHDP" && cfg.phase === "PHASE 1") {
    // An incomplete study case has no final score. A real zero remains a score.
    if (valid.length !== cfg.modules.length) return "";
    return Math.round(modules.reduce(function(sum, m, i) { return sum + m.score * cfg.modules[i].weight; }, 0) * 100) / 100;
  }
  return Math.round(valid.reduce(function(sum, m) { return sum + m.score; }, 0) / valid.length * 100) / 100;
}
function getAssessmentConfig_(program, phase) {
  var prog = String(program || "")
    .trim()
    .toUpperCase();
  var ph = normalizeCurrentPhase_(phase);
  if(activePrograms_().indexOf(prog)<0) return null;
  return PROGRAM_PHASE_ASSESSMENT_CONFIG[(activeSegment_()==="B2C" && prog==="TSHDP" ? "B2C|" : "") + prog + "|" + ph] || null;
}

function headerKey_(v) {
  return String(v || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function findHeaderIndex_(headers, aliases) {
  var keys = headers.map(headerKey_);
  for (var i = 0; i < aliases.length; i++) {
    var idx = keys.indexOf(headerKey_(aliases[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

function parseTabularBase64_(fileName, mimeType, base64) {
  var raw = String(base64 || "");
  var marker = raw.indexOf("base64,");
  if (marker !== -1) raw = raw.substring(marker + 7);
  var bytes = Utilities.base64Decode(raw);
  var lower = String(fileName || "").toLowerCase();
  var mime = String(mimeType || "");

  var isCsv =
    lower.slice(-4) === ".csv" ||
    mime === "text/csv" ||
    mime === "application/csv";
  if (isCsv) {
    var text = Utilities.newBlob(bytes)
      .getDataAsString("UTF-8")
      .replace(/^\uFEFF/, "");
    var delimiter = text.split(";").length > text.split(",").length ? ";" : ",";
    return Utilities.parseCsv(text, delimiter);
  }

  var isExcel =
    lower.slice(-5) === ".xlsx" ||
    lower.slice(-4) === ".xls" ||
    mime.indexOf("spreadsheetml") !== -1 ||
    mime === "application/vnd.ms-excel";
  if (!isExcel)
    throw new Error("Format file tidak didukung. Gunakan .xlsx atau .csv.");

  var blob = Utilities.newBlob(
    bytes,
    mime || MimeType.MICROSOFT_EXCEL,
    sanitizeFileName_(fileName || "import.xlsx"),
  );
  var tempId = null;
  try {
    tempId = convertBlobToSheet_(blob);
    var sheet = SpreadsheetApp.openById(tempId).getSheets()[0];
    var lastRow = sheet.getLastRow(),
      lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) return [];
    return sheet.getRange(1, 1, lastRow, lastCol).getValues();
  } finally {
    if (tempId) {
      try {
        DriveApp.getFileById(tempId).setTrashed(true);
      } catch (e) {}
    }
  }
}

function convertBlobToSheet_(blob) {
  var name =
    "TMP_CERT_IMPORT_" +
    Utilities.formatDate(new Date(), TZ, "yyyyMMdd_HHmmss");
  if (typeof Drive === "undefined" || !Drive.Files) {
    throw new Error(
      "Untuk file Excel (.xlsx), aktifkan Advanced Drive Service. Alternatif: simpan sebagai .csv lalu upload .csv.",
    );
  }
  if (typeof Drive.Files.create === "function") {
    return Drive.Files.create(
      { name: name, mimeType: MimeType.GOOGLE_SHEETS },
      blob,
    ).id;
  }
  return Drive.Files.insert(
    { title: name, mimeType: MimeType.GOOGLE_SHEETS },
    blob,
    { convert: true },
  ).id;
}

function parseImportDate_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  if (typeof v === "number" && v > 20000 && v < 80000) {
    var base = new Date(Date.UTC(1899, 11, 30));
    base.setUTCDate(base.getUTCDate() + Math.floor(v));
    return Utilities.formatDate(base, TZ, "yyyy-MM-dd");
  }
  var s = String(v).trim();
  var iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso)
    return (
      iso[1] + "-" + ("0" + iso[2]).slice(-2) + "-" + ("0" + iso[3]).slice(-2)
    );
  var dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
  if (dmy)
    return (
      dmy[3] + "-" + ("0" + dmy[2]).slice(-2) + "-" + ("0" + dmy[1]).slice(-2)
    );
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime()))
    return Utilities.formatDate(parsed, TZ, "yyyy-MM-dd");
  return "";
}

function pickCell_(rowArr, idx) {
  if (idx === undefined || idx === null || idx < 0) return "";
  var v = rowArr[idx];
  return v === undefined || v === null ? "" : v;
}

function parseQleapExport_(fileName, mimeType, base64) {
  var rows = parseTabularBase64_(fileName, mimeType, base64);
  if (!rows || rows.length < 4)
    throw new Error("File kosong atau format tidak sesuai template QLEAP.");

  var headerRowIdx = -1;
  for (var i = 0; i < Math.min(rows.length, 6); i++) {
    var r = rows[i] || [];
    if (
      r.some(function (c) {
        return (
          String(c || "")
            .trim()
            .toUpperCase() === "NIK"
        );
      })
    ) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error('Header "NIK" tidak ditemukan.');

  var headerRow = rows[headerRowIdx];
  var getIdx = function (names) {
    for (var i = 0; i < headerRow.length; i++) {
      var k = String(headerRow[i] || "")
        .trim()
        .toUpperCase();
      if (names.indexOf(k) !== -1) return i;
    }
    return -1;
  };

  var nikCol = getIdx(["NIK"]);
  var nameCol = getIdx(["NAME", "NAMA"]);
  var jobCol = getIdx(["JOB TITLE", "JOBTITLE", "JABATAN"]);
  var locCol = getIdx(["WORK LOCATION", "WORKLOCATION", "LOKASI"]);
  var buCol = getIdx(["BUSINESS UNIT", "BUSINESSUNIT"]);
  var elStatCol = getIdx([
    "E-LEARNING STATUS",
    "ELEARNING STATUS",
    "STATUS E-LEARNING",
  ]);
  var modCol = getIdx(["MODULES", "MODULE"]);

  if (nikCol < 0) throw new Error("Kolom NIK tidak ditemukan.");

  var groupRow = rows[headerRowIdx + 1] || [];
  var subRow = rows[headerRowIdx + 2] || [];

  var examBlocks = [];
  var currentGroup = "";
  for (var c = nikCol + 1; c < subRow.length; c++) {
    var g = String(groupRow[c] || "").trim();
    if (g) currentGroup = g;
    if (
      String(subRow[c] || "")
        .trim()
        .toUpperCase() === "SCORE"
    ) {
      examBlocks.push({
        examName: (currentGroup || "Exam " + (examBlocks.length + 1)).trim(),
        scoreCol: c,
        durCol: c + 1,
        dateCol: c + 2,
        statusCol: c + 3,
      });
    }
  }
  if (!examBlocks.length)
    throw new Error('Tidak ada blok "Score" di sub-header.');

  var dataStart = headerRowIdx + 3;
  var records = [];
  for (var r2 = dataStart; r2 < rows.length; r2++) {
    var row = rows[r2] || [];
    var nik = normalizeNik_(row[nikCol]);
    if (!nik) continue;
    var base = {
      nik: nik,
      name: String(row[nameCol] || "").trim(),
      jobTitle: String(row[jobCol] || "").trim(),
      workLocation: String(row[locCol] || "").trim(),
      businessUnit: String(row[buCol] || "").trim(),
      elStat: String(row[elStatCol] || "").trim(),
      modules: row[modCol] !== undefined ? row[modCol] : "",
    };
    examBlocks.forEach(function (b) {
      var rawScore = row[b.scoreCol];
      if (rawScore === "" || rawScore === null || rawScore === undefined)
        return;
      var score = Number(rawScore);
      if (isNaN(score)) return;
      records.push({
        nik: base.nik,
        name: base.name,
        jobTitle: base.jobTitle,
        workLocation: base.workLocation,
        businessUnit: base.businessUnit,
        elStat: base.elStat,
        modules: base.modules,
        examName: b.examName,
        score: score,
        duration: String(row[b.durCol] || "").trim(),
        examDate: formatQleapDate_(row[b.dateCol]),
        status: String(row[b.statusCol] || "").trim(),
      });
    });
  }
  return records;
}

function formatQleapDate_(v) {
  if (!v) return "";
  if (v instanceof Date)
    return Utilities.formatDate(v, TZ, "yyyy-MM-dd HH:mm:ss");
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/);
  if (m)
    return (
      m[1] + "-" + m[2] + "-" + m[3] + " " + m[4] + ":" + m[5] + ":" + m[6]
    );
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  return s;
}

// ---------------------------------------------------------------------------
// Repository and request lifecycle. Login bootstrapping syncs trusted User Import only.
// ---------------------------------------------------------------------------
var LDP_VERSION_ = "2.9.0";
var LDP_SCHEMA_ = "3";
var SHEET_JOBS_ = "Certificate Jobs";
var CTX_ = null;
var MUTATING_FIELDS_ = [
  "Role",
  "Region",
  "Superior_NIK",
  "Active",
  "PasswordHash",
  "Salt",
];
SHEET_HEADERS[SHEET_USERS].push("AuthVersion");
SHEET_HEADERS[SHEET_ENROLLMENTS].push("AssessmentSummary", "RecordVersion");
SHEET_HEADERS[SHEET_CERTIFICATES].push("EnrollmentID", "Batch", "JobID");
SHEET_HEADERS[SHEET_SUBMISSIONS].push("EnrollmentID", "OperationID");
SHEET_HEADERS[SHEET_MENTORING].push("EnrollmentID", "OperationID");
SHEET_HEADERS[SHEET_CERT_IMPORT].push("EnrollmentID", "JobID");
SHEET_HEADERS[SHEET_JOBS_] = [
  "JobID",
  "EnrollmentID",
  "NIK",
  "Program",
  "Batch",
  "DisplayName",
  "CertificateNumber",
  "IssueDate",
  "TemplateID",
  "Status",
  "Attempts",
  "LeaseUntil",
  "LeaseToken",
  "Error",
  "FileUrl",
  "ImportID",
  "CreatedBy",
  "CreatedAt",
  "UpdatedAt",
];

function context_() {
  if (!CTX_)
    CTX_ = {
      tables: {},
      headers: {},
      sheets: {},
      memo: {},
      dirty: {},
      depth: 0,
      lockDepth: 0,
      bulk: null,
      reads: 0,
      writes: 0,
      cells: 0,
    };
  return CTX_;
}
function nowIso_() {
  return new Date().toISOString();
}
function props_() {
  return PropertiesService.getScriptProperties();
}
function revision_() {
  var c = context_();
  if (c.revision === undefined)
    c.revision = props_().getProperty("LDP_DATA_REV") || "0";
  return c.revision;
}
// Global revision still identifies writes to clients. Read caches use only the
// tables they depend on, so a queue lease or audit entry cannot evict a roster.
function tableRevisions_() {
  var c = context_();
  if (!c.tableRevisions) c.tableRevisions = parseJson_(props_().getProperty("LDP_TABLE_REVISIONS"), {});
  return c.tableRevisions;
}
function sourceRevision_(names) {
  var revisions = tableRevisions_();
  return cacheKey_("", JSON.stringify([LDP_VERSION_, props_().getProperty("LDP_SPREADSHEET_ID") || "bound",
    revisions["*"] || "0", names.map(function (name) { return [name, revisions[name] || "0"]; })]));
}
function resetSourceRevisions_() {
  // Manual/formula/structural Sheet changes can have dependencies outside the
  // application write paths. Conservatively invalidate every source in that case.
  return withLock_(function () {
    var value = Utilities.getUuid(), revisions = {"*": value};
    props_().setProperty("LDP_TABLE_REVISIONS", JSON.stringify(revisions));
    props_().setProperty("LDP_DATA_REV", value);
    context_().tableRevisions = revisions; context_().revision = value;
  });
}
function withLock_(fn) {
  var c = context_();
  if (c.lockDepth) return fn();
  var lock = LockService.getScriptLock();
  var start = Date.now();
  lock.waitLock(8000);
  c.lockWait = (c.lockWait || 0) + Date.now() - start;
  c.lockDepth++;
  try {
    return fn();
  } finally {
    c.lockDepth--;
    try {
      flushRevision_();
    } finally {
      lock.releaseLock();
    }
  }
}
function flushRevision_() {
  var c = context_();
  if (!c.lockDepth && Object.keys(c.dirty).length && !c.revisionFlushed) {
    c.revision = Utilities.getUuid();
    // Reload under the write lock; never overwrite another execution's revisions.
    var revisions = parseJson_(props_().getProperty("LDP_TABLE_REVISIONS"), {});
    Object.keys(c.revisionDirty || c.dirty).forEach(function (name) { revisions[name] = c.revision; });
    props_().setProperty("LDP_TABLE_REVISIONS", JSON.stringify(revisions));
    c.tableRevisions = revisions; c.revisionDirty = {};
    props_().setProperty("LDP_DATA_REV", c.revision);
    c.revisionFlushed = true;
  }
}
function normalizeSegment_(value) {
  var s = String(value === undefined || value === null ? "" : value).trim().toUpperCase().replace(/\s+/g, "");
  if (!s) return "B2B";
  if (s !== "B2B" && s !== "B2C") throw new Error("Segmen harus B2B atau B2C.");
  return s;
}
function activeSegment_() { return context_().segment || "B2B"; }
function activePrograms_(segment) {
  var s = segment || activeSegment_();
  return s === "ALL" ? CERT_PROGRAM_ORDER.slice() : normalizeSegment_(s) === "B2C" ? ["SLDP","TSHDP"] : ["ASHDP","TSHDP"];
}
function requestSegment_(name, args) {
  if (name === "certificateWorker") return "ALL";
  var explicit, marker;
  Array.prototype.forEach.call(args || [], function (arg) {
    if (!arg || typeof arg !== "object" || Array.isArray(arg)) return;
    if (arg.__ldpContext === 1) marker = arg.segment;
    else if (arg.segment !== undefined) explicit = arg.segment;
  });
  var value = marker === undefined ? explicit : marker;
  if (value === "ALL") {
    if (["login","logout","getDashboardBootstrap","switchRole","getLoginConfig"].indexOf(name)<0)
      throw new Error("Pilih satu segmen B2B atau B2C untuk membuka data.");
    return "ALL";
  }
  var segment = normalizeSegment_(value);
  if (explicit !== undefined && normalizeSegment_(explicit)!==segment)
    throw new Error("Segmen tujuan berbeda dari halaman aktif. Buka segmen yang sesuai.");
  return segment;
}
function isSegmentedTable_(name) {
  if ([SHEET_USERS,SHEET_USER_IMPORT,SHEET_SETTINGS].indexOf(name)>=0) return false;
  if (Object.prototype.hasOwnProperty.call(SHEET_HEADERS,name)) return true;
  return name === (getSetting_("PHASE3_SHEET") || "Phase 3") || name === (getSetting_("QLEAP_MAPPING_SHEET") || SHEET_QLEAP) ||
    name === (getSetting_("SOURCE_PARTICIPANT_SHEET") || "Legacy Participants");
}
function segmentRows_(name, rows) {
  if (!isSegmentedTable_(name)) return rows;
  // Segment alone does not turn a formerly empty spreadsheet row into a record.
  rows=rows.filter(function(row){return Object.keys(row.data).some(function(k){
    return k!=="Segment" && k!=="Segmen" && row.data[k]!=="" && row.data[k]!==null && row.data[k]!==undefined;
  });});
  rows.forEach(function (row) { row.data.Segment = normalizeSegment_(row.data.Segment || row.data.Segmen); });
  var segment = activeSegment_();
  return segment === "ALL" ? rows : rows.filter(function (row) { return row.data.Segment === segment; });
}
function allSegmentRows_(name) {
  var c=context_(),previous=c.segment;c.segment="ALL";
  try {return readAllAsObjects_(name);}finally {c.segment=previous;}
}
function segmentWrite_(name, headers, value) {
  if (!isSegmentedTable_(name)) return value;
  var segment = value.Segment || value.Segmen || (activeSegment_() === "ALL" ? "B2B" : activeSegment_());
  segment = normalizeSegment_(segment);
  if (activeSegment_() !== "ALL" && segment !== activeSegment_()) throw new Error("Data berada pada segmen lain.");
  if (segment === "B2C" && headers.indexOf("Segment")<0)
    throw new Error("Struktur segmen belum tersedia. Jalankan setupLdpSystem satu kali untuk menambahkan kolom Segment.");
  return Object.assign({},value,{Segment:segment});
}
function storedSegmentForRow_(sheet, headers, rowIndex) {
  var c=context_(),name=sheet.getName(),row=(c.tables[name] || []).filter(function(r){return r.rowIndex===rowIndex;})[0];
  if(row) return normalizeSegment_(row.data.Segment || row.data.Segmen);
  var pending=c.bulk && c.bulk.appends[name],col=headers.indexOf("Segment");
  if(pending && rowIndex>=pending.start && rowIndex<pending.start+pending.rows.length)
    return normalizeSegment_(col<0 ? "" : pending.rows[rowIndex-pending.start][col]);
  return normalizeSegment_(col<0 ? "" : sheet.getRange(rowIndex,col+1,1,1).getValues()[0][0]);
}
function api_(name, fn, args) {
  var previous = CTX_;
  CTX_ = null;
  var c = context_(),
    start = Date.now(),
    ok = false;
  try {
    c.segment = requestSegment_(name,args);
    if (
      name !== "getLoginConfig" &&
      props_().getProperty("LDP_SCHEMA_VERSION") !== LDP_SCHEMA_
    )
      throw new Error(
        "Struktur aplikasi belum diperbarui. Admin perlu menjalankan setupLdpSystem dari editor Apps Script.",
      );
    var result = fn();
    ok = true;
    return result;
  } finally {
    if (Object.keys(c.dirty).length && !c.revisionFlushed) {
      withLock_(function () {
        flushRevision_();
      });
    }
    console.log(
      JSON.stringify({
        app: "LDP",
        version: LDP_VERSION_,
        endpoint: name,
        ok: ok,
        ms: Date.now() - start,
        reads: c.reads,
        writes: c.writes,
        cells: c.cells,
        cacheHits: c.cacheHits || 0,
        cacheMisses: c.cacheMisses || 0,
        lockWaitMs: c.lockWait || 0,
      }),
    );
    CTX_ = previous;
  }
}
function clone_(v) {
  return JSON.parse(JSON.stringify(v));
}
function getSheet_(name) {
  var c = context_();
  if (!c.ss) {
    var spreadsheetId = props_().getProperty("LDP_SPREADSHEET_ID");
    c.ss = spreadsheetId
      ? SpreadsheetApp.openById(spreadsheetId)
      : SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!c.ss)
    throw new Error(
      "Spreadsheet database belum terhubung. Isi Script Property LDP_SPREADSHEET_ID atau buka editor dari spreadsheet database, lalu jalankan setupLdpSystem.",
    );
  if (!c.sheets[name]) c.sheets[name] = c.ss.getSheetByName(name);
  if (!c.sheets[name])
    throw new Error(
      "Sheet " + name + " tidak tersedia. Jalankan setupLdpSystem dari editor.",
    );
  return c.sheets[name];
}
function cacheKey_(prefix, value) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
  );
  return (
    prefix +
    bytes
      .map(function (b) {
        return ("0" + (b & 255).toString(16)).slice(-2);
      })
      .join("")
  );
}
function getHeaders_(sheet) {
  var c = context_(),
    name = sheet.getName();
  if (c.headers[name]) return c.headers[name];
  var key = cacheKey_("ldph3:", sourceRevision_([name]));
  var raw = CacheService.getScriptCache().get(key);
  if (raw && !c.dirty[name]) {
    try {
      return (c.headers[name] = JSON.parse(raw));
    } catch (e) {}
  }
  var n = sheet.getLastColumn();
  var h = n
    ? sheet
        .getRange(1, 1, 1, n)
        .getValues()[0]
        .map(function (x) {
          return String(x).trim();
        })
    : [];
  if (h.length)
    try {
      CacheService.getScriptCache().put(key, JSON.stringify(h), 60);
    } catch (e) {}
  return (c.headers[name] = h);
}
function rowToObject_(headers, values) {
  var o = {};
  headers.forEach(function (h, i) {
    var v = values[i];
    o[h] = v instanceof Date ? v.toISOString() : v === undefined ? "" : v;
  });
  return o;
}
function readSnapshot_(key) {
  var cache = CacheService.getScriptCache(),
    raw = cache.get(key);
  if (!raw) {context_().cacheMisses=(context_().cacheMisses || 0)+1;return null;}
  try {
    var manifest = JSON.parse(raw),
      keys = [];
    for (var i = 0; i < manifest.parts; i++) keys.push(key + ":" + i);
    var parts = cache.getAll(keys),
      s = "";
    for (var j = 0; j < keys.length; j++) {
      if (parts[keys[j]] === undefined) return null;
      s += parts[keys[j]];
    }
    var snapshot=JSON.parse(s);
    context_().cacheHits=(context_().cacheHits || 0)+1;
    return snapshot;
  } catch (e) {
    return null;
  }
}
function putSnapshot_(key, rows, ttl) {
  var raw = JSON.stringify(rows),
    size = 20000,
    count = Math.ceil(raw.length / size);
  if (count > 64) return; // Fall back to Sheets, never truncate data.
  var values = {};
  for (var i = 0; i < count; i++)
    values[key + ":" + i] = raw.slice(i * size, (i + 1) * size);
  try {
    var cache = CacheService.getScriptCache();
    ttl = ttl || 40;
    cache.putAll(values, ttl + 5);
    cache.put(key, JSON.stringify({ parts: count }), ttl);
  } catch (e) {
    console.warn("Snapshot cache unavailable");
  }
}
function readAllAsObjects_(name) {
  var c = context_();
  if (c.tables[name]) return segmentRows_(name,c.tables[name]);
  var key = cacheKey_("ldpt3:", sourceRevision_([name]));
  // Users contains credential material; auth always reads the current row.
  var allowCache =
    name !== SHEET_USERS && name !== SHEET_AUDIT && !c.dirty[name] && !c.bulk;
  if (allowCache && !c.memo.bypassSnapshots) {
    var cached = readSnapshot_(key);
    if (cached && cached.format === "LDP_TABLE_1" && Array.isArray(cached.rows))
      return segmentRows_(name,(c.tables[name] = cached.rows.map(function(row) {
        return {rowIndex:row[0],data:rowToObject_(cached.headers,row.slice(1))};
      })));
  }
  var s = getSheet_(name),
    h = getHeaders_(s),
    last = s.getLastRow(),
    out = [];
  if (last >= 2 && h.length) {
    var values = s.getRange(2, 1, last - 1, h.length).getValues();
    c.reads++;
    c.cells += values.length * h.length;
    values.forEach(function (v, i) {
      if (
        v.some(function (x) {
          return x !== "" && x !== null && x !== undefined;
        })
      )
        out.push({ rowIndex: i + 2, data: rowToObject_(h, v) });
    });
  }
  c.tables[name] = out;
  if (allowCache) putSnapshot_(key,{format:"LDP_TABLE_1",headers:h,rows:out.map(function(row) {
    return [row.rowIndex].concat(h.map(function(header) {return row.data[header];}));
  })});
  return segmentRows_(name,out);
}
function invalidateSheet_(name) {
  var c = context_();
  c.dirty[name] = true;
  (c.revisionDirty || (c.revisionDirty = {}))[name] = true;
  c.revisionFlushed = false;
  delete c.memo["projected:" + name];
  if (name === SHEET_SETTINGS) {
    c.memo = {};
    return;
  }
  if (name === SHEET_ENROLLMENTS) { delete c.memo.enrollments; delete c.memo.enrollmentsLight; }
  if (name === SHEET_USERS) delete c.memo.userRows;
  if (name === SHEET_SUBMISSIONS || name === SHEET_MENTORING)
    delete c.memo.workspaceIndex;
  if (name === SHEET_CERTIFICATES) delete c.memo.certIndex;
  if (name === SHEET_QPT_LATEST_ || name === SHEET_QLEAP_POSTTEST)
    delete c.memo.latestQpt;
  if (name === (getSetting_("QLEAP_MAPPING_SHEET") || SHEET_QLEAP))
    delete c.memo.qleap;
  if (name === (getSetting_("PHASE3_SHEET") || "Phase 3")) delete c.memo.phase3;
  Object.keys(c.memo).forEach(function(k) {if(k.indexOf("assessment:"+name+"|")===0)delete c.memo[k];});
  Object.keys(c.memo).forEach(function (k) {
    if (k.indexOf("detail:") === 0) delete c.memo[k];
  });
}
function patchColumns_(sheet, headers, rowIndex, obj) {
  var c = context_(),
    cols = [];
  var pending = c.bulk && c.bulk.appends[sheet.getName()];
  if (
    pending &&
    rowIndex >= pending.start &&
    rowIndex < pending.start + pending.rows.length
  ) {
    headers.forEach(function (h, i) {
      if (Object.prototype.hasOwnProperty.call(obj, h))
        pending.rows[rowIndex - pending.start][i] =
          obj[h] === undefined || obj[h] === null ? "" : obj[h];
    });
    return;
  }
  headers.forEach(function (h, i) {
    if (Object.prototype.hasOwnProperty.call(obj, h))
      cols.push({
        col: i + 1,
        value: obj[h] === null || obj[h] === undefined ? "" : obj[h],
      });
  });
  var groups = [];
  cols.forEach(function (x) {
    var g = groups[groups.length - 1];
    if (g && g.col + g.values.length === x.col) g.values.push(x.value);
    else groups.push({ row: rowIndex, col: x.col, values: [x.value] });
  });
  if (c.bulk) {
    var name = sheet.getName();
    c.bulk.patches[name] = (c.bulk.patches[name] || []).concat(groups);
  } else
    groups.forEach(function (g) {
      sheet.getRange(g.row, g.col, 1, g.values.length).setValues([g.values]);
      c.writes++;
    });
}
function updateRowByIndex_(sheet, headers, rowIndex, obj) {
  var c = context_(),
    name = sheet.getName(),
    patch = Object.assign({}, obj);
  if (isSegmentedTable_(name) && activeSegment_()!=="ALL") {
    var stored=storedSegmentForRow_(sheet,headers,rowIndex);
    if (stored!==activeSegment_()) throw new Error("Data berada pada segmen lain.");
    if (patch.Segment!==undefined && normalizeSegment_(patch.Segment)!==stored)
      throw new Error("Segmen enrollment tidak dapat dipindahkan melalui perubahan profil.");
  }
  return withLock_(function () {
    // Only fields being changed are written. Formula and other editors' columns survive.
    if (
      name === SHEET_USERS &&
      MUTATING_FIELDS_.some(function (h) {
        return Object.prototype.hasOwnProperty.call(patch, h);
      })
    ) {
      var col = headers.indexOf("AuthVersion");
      if (col < 0) throw new Error("Kolom AuthVersion belum tersedia.");
      var pendingUser = (c.tables[name] || []).filter(function (r) {
        return r.rowIndex === rowIndex;
      })[0];
      var current = pendingUser
        ? pendingUser.data.AuthVersion
        : sheet.getRange(rowIndex, col + 1, 1, 1).getValues()[0][0];
      patch.AuthVersion = Number(current || 0) + 1;
    }
    if (name === SHEET_ENROLLMENTS) {
      patch.RecordVersion = Utilities.getUuid();
    }
    patchColumns_(sheet, headers, rowIndex, patch);
    if (c.tables[name]) {
      var r = c.tables[name].filter(function (x) {
        return x.rowIndex === rowIndex;
      })[0];
      if (r) Object.assign(r.data, patch);
    }
    invalidateSheet_(name);
    return patch;
  });
}
function appendObject_(sheet, headers, obj) {
  return withLock_(function () {
    var c = context_(),
      name = sheet.getName(),
      copy = segmentWrite_(name,headers,Object.assign({}, obj));
    if (name === SHEET_USERS) copy.AuthVersion = Number(copy.AuthVersion || 1);
    if (name === SHEET_ENROLLMENTS) copy.RecordVersion = Utilities.getUuid();
    var row = headers.map(function (h) {
        return copy[h] === undefined || copy[h] === null ? "" : copy[h];
      }),
      index;
    if (c.bulk) {
      var app =
        c.bulk.appends[name] ||
        (c.bulk.appends[name] = { start: sheet.getLastRow() + 1, rows: [] });
      index = app.start + app.rows.length;
      app.rows.push(row);
    } else {
      sheet.appendRow(row);
      index = sheet.getLastRow();
      c.writes++;
    }
    if (c.tables[name])
      c.tables[name].push({
        rowIndex: index,
        data: rowToObject_(headers, row),
      });
    invalidateSheet_(name);
    return index;
  });
}
function bulkWrites_(fn) {
  var c = context_();
  if (c.bulk) return fn();
  return withLock_(function () {
    // Reload inside the write lock; cached snapshots must not decide new row positions.
    c.tables = {};
    c.memo = {};
    c.bulk = { patches: {}, appends: {} };
    try {
      var result = fn(),
        bulk = c.bulk;
      c.bulk = null;
      Object.keys(bulk.patches).forEach(function (name) {
        var byShape = {};
        bulk.patches[name].forEach(function (p) {
          var k = p.col + ":" + p.values.length;
          (byShape[k] || (byShape[k] = [])).push(p);
        });
        Object.keys(byShape).forEach(function (k) {
          var rows = byShape[k].sort(function (a, b) {
              return a.row - b.row;
            }),
            groups = [];
          rows.forEach(function (p) {
            var g = groups[groups.length - 1];
            if (g && g.row + g.values.length === p.row) g.values.push(p.values);
            else groups.push({ row: p.row, col: p.col, values: [p.values] });
          });
          groups.forEach(function (g) {
            getSheet_(name)
              .getRange(g.row, g.col, g.values.length, g.values[0].length)
              .setValues(g.values);
            c.writes++;
          });
        });
      });
      Object.keys(bulk.appends).forEach(function (name) {
        var a = bulk.appends[name];
        getSheet_(name)
          .getRange(a.start, 1, a.rows.length, a.rows[0].length)
          .setValues(a.rows);
        c.writes++;
      });
      return result;
    } catch (e) {
      c.bulk = null;
      c.tables = {};
      c.memo = {};
      throw e;
    }
  });
}
function uniqueMatch_(rows, label) {
  if (rows.length > 1)
    throw new Error(
      "Data duplikat pada " +
        label +
        ". Admin perlu menyelesaikan konflik data.",
    );
  return rows[0] || null;
}
function entityKey_(nik, program, batch, segment) {
  return [
    normalizeSegment_(segment || (activeSegment_() === "ALL" ? "B2B" : activeSegment_())),
    normalizeNik_(nik),
    String(program || "")
      .trim()
      .toUpperCase(),
    normalizeBatchKey_(batch),
  ].join("|");
}
function csvEscape_(value) {
  var s = String(value === undefined || value === null ? "" : value);
  if (/^[=+@\-]/.test(s)) s = "'" + s;
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
function isPassingStatus_(v) {
  return (
    ["LULUS", "PASS", "PASSED", "COMPLETE"].indexOf(
      String(v || "")
        .trim()
        .toUpperCase(),
    ) >= 0
  );
}
function computeResult_(p) {
  var s = String(p.FinalStatus || p.ProgramStatus || "").toUpperCase();
  if(s==="RESIGN"||s==="RESIGNED")return "RESIGN";
  if (isPassingStatus_(s)) return "PASSED";
  if (s === "REMEDIAL") return "REMEDIAL";
  if (["TIDAK LULUS", "NOT PASSED", "FAIL", "FAILED"].indexOf(s) >= 0)
    return "NOT PASSED";
  return normalizeParticipantStatus_(p.ParticipantStatus) === "RESIGN"
    ? "NOT PASSED"
    : "ON GOING";
}
function inferCurrentPhase_(p) {
  if (
    isPassingStatus_(p.FinalStatus) ||
    ["TIDAK LULUS", "NOT PASSED"].indexOf(p.FinalStatus) >= 0
  )
    return "COMPLETED";
  if (p.PanelScore !== "" && p.PanelScore !== undefined) return "PANELISASI";
  if (p.Phase3Score !== "" && p.Phase3Score !== undefined) return "PHASE 3";
  if (p.Phase2Score !== "" && p.Phase2Score !== undefined) return "PHASE 2";
  return "PHASE 1";
}
function numberOrBlank_(v) {
  if (v === "" || v === null || v === undefined) return "";
  var n = Number(v);
  return isFinite(n) ? n : "";
}
function validScore_(v) {
  var n = numberOrBlank_(v);
  if (n !== "" && (n < 0 || n > 100))
    throw new Error("Nilai harus 0 sampai 100.");
  return n;
}
function parseJson_(v, fallback) {
  try {
    return JSON.parse(v || "") || fallback;
  } catch (e) {
    return fallback;
  }
}
function timeValue_(v) {
  var s = String(v || "");
  if (/^\d{4}-\d\d-\d\d[T ]\d\d:\d\d:\d\d$/.test(s))
    s = s.replace(" ", "T") + "+07:00";
  return new Date(s).getTime() || 0;
}
function paginate_(rows, filter) {
  var size = Math.max(10, Math.min(100, Number(filter.pageSize) || 30)),
    page = Math.max(
      1,
      Math.min(
        Math.max(1, Math.ceil(rows.length / size)),
        Math.floor(Number(filter.page) || 1),
      ),
    );
  return {
    rows: rows.slice((page - 1) * size, page * size),
    total: rows.length,
    page: page,
    pageSize: size,
    pages: Math.max(1, Math.ceil(rows.length / size)),
  };
}

// ---------------------------------------------------------------------------
// Settings and authentication. Revocation is checked against a current user row.
// ---------------------------------------------------------------------------
function getAllSettings_() {
  var c = context_();
  if (!c.settings) {
    c.settings = {};
    readAllAsObjects_(SHEET_SETTINGS).forEach(function (r) {
      c.settings[r.data.Key] = r.data.Value;
    });
  }
  return c.settings;
}
function getSetting_(key) {
  var s = getAllSettings_();
  return Object.prototype.hasOwnProperty.call(s, key) ? s[key] : "";
}
function setSetting_(key, value) {
  var rows = readAllAsObjects_(SHEET_SETTINGS),
    s = getSheet_(SHEET_SETTINGS),
    h = getHeaders_(s),
    r = uniqueMatch_(
      rows.filter(function (x) {
        return x.data.Key === key;
      }),
      "Settings " + key,
    );
  if (r) updateRowByIndex_(s, h, r.rowIndex, { Value: value });
  else appendObject_(s, h, { Key: key, Value: value });
  context_().settings = null;
}
function getRoleSettings_() {
  return {
    clpEnabled: parseBoolean_(getSetting_("CLP_ENABLED")),
    superiorEnabled: parseBoolean_(getSetting_("SUPERIOR_ENABLED")),
  };
}
function isRoleEnabled_(role) {
  if (role === "ADMIN" || role === "PARTICIPANT") return true;
  var s = getRoleSettings_();
  return role === "CLP"
    ? s.clpEnabled
    : role === "SUPERIOR" && s.superiorEnabled;
}
function findUserRowByNik_(nik, hint) {
  var s = getSheet_(SHEET_USERS),
    h = getHeaders_(s),
    nc = h.indexOf("NIK"),
    target = normalizeNik_(nik);
  if (nc < 0) throw new Error("Kolom NIK pada Users tidak tersedia.");
  if (hint && hint >= 2 && hint <= s.getLastRow()) {
    var v = s.getRange(hint, 1, 1, h.length).getValues()[0];
    context_().reads++;
    if (normalizeNik_(v[nc]) === target)
      return { rowIndex: hint, data: rowToObject_(h, v) };
  }
  var key = cacheKey_("ldpu:", revision_()),
    idx = context_().memo.userRows;
  if (!idx) {
    idx = readSnapshot_(key);
    if (!idx) {
      idx = {};
      var last = s.getLastRow();
      if (last >= 2) {
        s.getRange(2, nc + 1, last - 1, 1)
          .getDisplayValues()
          .forEach(function (r, i) {
            var n = normalizeNik_(r[0]);
            if (n) {
              if (idx[n]) idx[n] = -1;
              else idx[n] = i + 2;
            }
          });
        context_().reads++;
      }
      putSnapshot_(key, idx);
    }
    context_().memo.userRows = idx;
  }
  var row = idx[target];
  if (row === -1) throw new Error("NIK duplikat pada Users. Hubungi Admin.");
  if (!row) return null;
  var values = s.getRange(row, 1, 1, h.length).getValues()[0];
  context_().reads++;
  if (normalizeNik_(values[nc]) !== target) {
    CacheService.getScriptCache().remove(key);
    context_().memo.userRows = null;
    throw new Error("Posisi data berubah. Silakan ulangi login.");
  }
  return { rowIndex: row, data: rowToObject_(h, values) };
}
function requireSession_(token) {
  if (!token) throw new Error("AUTH_REQUIRED: Silakan login terlebih dahulu.");
  var session = parseJson_(
    CacheService.getScriptCache().get("sess_" + token),
    null,
  );
  if (!session || session.expiresAt <= Date.now())
    throw new Error("AUTH_EXPIRED: Sesi berakhir. Silakan login kembali.");
  var match = findUserRowByNik_(session.nik, session.userRowIndex),
    u = match && match.data;
  if (
    !u ||
    !parseBoolean_(u.Active) ||
    !userHasRole_(u, session.role) ||
    !isRoleEnabled_(session.role) ||
    String(u.AuthVersion || 1) !== String(session.authVersion)
  ) {
    CacheService.getScriptCache().remove("sess_" + token);
    throw new Error("AUTH_REVOKED: Akses akun berubah. Silakan login kembali.");
  }
  session.region = normalizeRegion_(u.Region);
  session.name = u.Name;
  session.superiorNik = normalizeNik_(u.Superior_NIK);
  return session;
}
function requireRole_(session, roles) {
  if (roles.indexOf(session.role) < 0)
    throw new Error("Akses tidak diizinkan untuk peran Anda.");
}
function authenticateAdminPassword_(match, password) {
  var nik = normalizeNik_(match.data.NIK), cache = CacheService.getScriptCache(),
    failKey = cacheKey_("ldpfail:", nik), failed = Number(cache.get(failKey) || 0), created = false;
  if (failed >= 8) throw new Error("Terlalu banyak percobaan login. Coba kembali dalam 15 menit.");
  if (!password) throw new Error("Masukkan password Admin. Saat login pertama, buat password minimal 8 karakter.");
  function verify(current) {
    if (!current || !parseBoolean_(current.data.Active) || !userHasRole_(current.data, ROLE_ADMIN))
      throw new Error("NIK tidak terdaftar atau akses Admin belum aktif.");
    if (current.data.PasswordHash) {
      if (hashPassword_(password, current.data.Salt) !== current.data.PasswordHash) {
        cache.put(failKey, String(failed + 1), 900);
        throw new Error("Password tidak sesuai.");
      }
      return current;
    }
    if (password.length < 8 || !password.trim()) throw new Error("Buat password Admin minimal 8 karakter.");
    var salt = generateSalt_(), sh = getSheet_(SHEET_USERS), patch = {
      Salt:salt, PasswordHash:hashPassword_(password, salt), UpdatedAt:nowIso_()
    };
    var saved = updateRowByIndex_(sh, getHeaders_(sh), current.rowIndex, patch);
    SpreadsheetApp.flush();
    Object.assign(current.data, saved); created = true;
    return current;
  }
  if (match.data.PasswordHash) match = verify(match);
  else match = withLock_(function() {
    // A competing first login must verify the password committed by the winner.
    delete context_().tables[SHEET_USERS];
    delete context_().memo.userRows;
    return verify(findUserRowByNik_(nik, match.rowIndex));
  });
  cache.remove(failKey);
  return {match:match, passwordCreated:created};
}
function login_(payload) {
  payload = payload || {};
  var nik = normalizeNik_(payload.nik),
    role = String(payload.role || "")
      .trim()
      .toUpperCase(),
    password = String(payload.password || "");
  if (!nik || ALL_ROLES.indexOf(role) < 0 || !isRoleEnabled_(role))
    throw new Error("NIK atau peran tidak valid.");
  var match = payload.fast === true ? findUserRowByNik_(nik) : null,
    sync = payload.fast === true && match ? opportunisticUserSync_() : autoSyncUserImport_();
  if (!match || sync.changed) match = findUserRowByNik_(nik);
  var u = match && match.data, passwordCreated = false;
  if (!u && !sync.ok) throw new Error("User Import belum dapat disinkronkan: " + sync.message);
  if (!u || !parseBoolean_(u.Active) || !userHasRole_(u, role))
    throw new Error("NIK tidak terdaftar atau akses peran belum aktif.");
  if (role === ROLE_ADMIN) {
    var authenticated = authenticateAdminPassword_(match, password);
    match = authenticated.match; u = match.data; passwordCreated = authenticated.passwordCreated;
  }
  if (
    role === ROLE_PARTICIPANT &&
    !readAllAsObjects_(SHEET_PARTICIPANTS).some(function (r) {
      return normalizeNik_(r.data.NIK) === nik;
    })
  )
    throw new Error("NIK belum terdaftar sebagai peserta LDP.");
  var ttl = Math.min(
      21600,
      Math.max(3600, (Number(getSetting_("SESSION_HOURS")) || 6) * 3600),
    ),
    token = Utilities.getUuid();
  var session = {
    token: token,
    nik: nik,
    name: u.Name,
    role: role,
    roles: parseRoles_(u.Role).filter(isRoleEnabled_),
    region: normalizeRegion_(u.Region),
    userRowIndex: match.rowIndex,
    authVersion: String(u.AuthVersion || 1),
    issuedAt: Date.now(),
    expiresAt: Date.now() + ttl * 1000,
  };
  CacheService.getScriptCache().put(
    "sess_" + token,
    JSON.stringify(session),
    ttl,
  );
  // Operational login telemetry avoids serial Spreadsheet audit writes during bursts.
  console.log(JSON.stringify({ event: "LOGIN_SUCCESS", nik: nik, role: role }));
  var response = {
    token: token,
    nik: nik,
    name: u.Name,
    role: role,
    roles: session.roles,
    region: session.region,
    expiresAt: session.expiresAt,
    passwordCreated: passwordCreated,
  };
  if (payload.bootstrap === true) response.dashboard = dashboardBootstrap_(session);
  return response;
}
function switchRole_(token, role, password) {
  var s = requireSession_(token), match = findUserRowByNik_(s.nik, s.userRowIndex), u = match.data;
  role = String(role || "").toUpperCase();
  if (!userHasRole_(u, role) || !isRoleEnabled_(role)) throw new Error("Peran tidak tersedia.");
  if (role === ROLE_ADMIN) {
    var authenticated = authenticateAdminPassword_(match, String(password || ""));
    u = authenticated.match.data;
  }
  s.role = role; s.roles = parseRoles_(u.Role).filter(isRoleEnabled_);
  s.authVersion = String(u.AuthVersion || 1);
  CacheService.getScriptCache().put("sess_" + token, JSON.stringify(s),
    Math.max(1, Math.ceil((s.expiresAt - Date.now()) / 1000)));
  return {ok:true, role:role, roles:s.roles};
}
function logout_(token) {
  if (token) CacheService.getScriptCache().remove("sess_" + token);
  return { ok: true };
}
function getLoginConfig_(options) {
  var ready = props_().getProperty("LDP_SCHEMA_VERSION") === LDP_SCHEMA_;
  if (!ready) return { appName: "LDP Monitor", ready: false, roles: [] };
  if (!(options && options.light === true)) autoSyncUserImport_();
  var roles = [
    { value: "PARTICIPANT", label: "Peserta" },
    { value: "ADMIN", label: "Admin" },
  ];
  if (isRoleEnabled_("CLP")) roles.splice(1, 0, { value: "CLP", label: "CLP" });
  if (isRoleEnabled_("SUPERIOR"))
    roles.splice(1, 0, { value: "SUPERIOR", label: "Superior" });
  return {
    appName: getSetting_("APP_NAME") || "LDP Monitor",
    logoUrl: getSetting_("LOGO_URL") || "",
    roles: roles,
    ready: true,
    version: LDP_VERSION_,
  };
}

// ---------------------------------------------------------------------------
// Scope, enrollment identities and read models.
// ---------------------------------------------------------------------------
var PARTICIPANT_LIST_FIELDS_ = ["Segment","NIK","Name","JobTitle","Region","City","Program","Batch",
  "Phase2Score","Phase3Score","PanelScore","FinalStatus","ParticipantStatus","Superior_NIK",
  "CLP_NIK","UpdatedAt","Location","CityName","CurrentPhase","CurrentEnrollmentId","ProgramTrack"];
var ENROLLMENT_LIST_FIELDS_ = ["Segment","EnrollmentID","NIK","Program","Batch","Region","CurrentPhase",
  "EnrollmentStatus","ProgramStatus","Phase1Score","Phase2Score","Phase3Score","FinalScore","UpdatedAt","RecordVersion"];
function readSummaryRows_(name, fields) {
  if (name === SHEET_USERS) throw new Error("Data akun tidak boleh disimpan dalam cache ringkasan.");
  var c = context_(), memo = "projected:" + name;
  if (c.memo[memo]) return segmentRows_(name,c.memo[memo]);
  function project(rows) { return rows.map(function (row) {
    var data = {}; fields.forEach(function (field) { data[field] = row.data[field] === undefined ? "" : row.data[field]; });
    return {rowIndex:row.rowIndex,data:data};
  }); }
  if (c.tables[name]) return segmentRows_(name,(c.memo[memo] = project(c.tables[name])));
  var key = cacheKey_("ldpsummary:", sourceRevision_([name]) + "|" + JSON.stringify(fields)),
    allowCache = !c.bulk && !c.dirty[name];
  if (allowCache && !c.memo.bypassSnapshots) {
    var saved = readSnapshot_(key);
    if (saved && saved.format === "LDP_SUMMARY_1")
      return segmentRows_(name,(c.memo[memo] = saved.rows.map(function (row) { return {rowIndex:row[0],data:rowToObject_(fields,row.slice(1))}; })));
  }
  var sheet = getSheet_(name), headers = getHeaders_(sheet), last = sheet.getLastRow(), rows = [],
    wanted = fields.map(function (field) { return headers.indexOf(field); }),
    columns = wanted.filter(function (col) { return col >= 0; }).sort(function (a,b) { return a-b; }), groups = [];
  // Small rosters use one batched block per table to avoid extra service calls.
  // Larger scans split around long columns; only selected fields enter cache.
  // Header order may vary, so no data column positions are hard-coded.
  columns.forEach(function (col) {
    var previous = groups[groups.length-1], split = !previous;
    if (previous && last > 500) for (var i=previous.end+1;i<col;i++)
      if (/summary|snapshot|notes|feedback|answersjson|detailsjson|actionplan|implementation|projecttitle|theme/i.test(headers[i])) { split=true;break; }
    if (split) groups.push({start:col,end:col}); else previous.end=col;
  });
  if (last >= 2 && columns.length) {
    var values = Array.from({length:last-1}, function () { return fields.map(function () { return ""; }); });
    groups.forEach(function (group) {
      var count = group.end-group.start+1,
        block = sheet.getRange(2,group.start+1,last-1,count).getValues();
      c.reads++;c.cells+=(last-1)*count;
      wanted.forEach(function (col,index) {
        if (col>=group.start && col<=group.end) block.forEach(function (row,r) { values[r][index]=row[col-group.start]; });
      });
    });
    values.forEach(function (row,index) {
      if (row.some(function (v) { return v!=="" && v!==null && v!==undefined; }))
        rows.push({rowIndex:index+2,data:rowToObject_(fields,row)});
    });
  }
  if (allowCache) putSnapshot_(key,{format:"LDP_SUMMARY_1",rows:rows.map(function (row) {
    return [row.rowIndex].concat(fields.map(function (field) { return row.data[field]; }));
  })},40);
  return segmentRows_(name,(c.memo[memo] = rows));
}
function scopedParticipants_(s, light) {
  var rows = light ? readSummaryRows_(SHEET_PARTICIPANTS,PARTICIPANT_LIST_FIELDS_) : readAllAsObjects_(SHEET_PARTICIPANTS);
  return rows.filter(function (r) {
    var p = r.data;
    return (
      s.role === "ADMIN" ||
      (s.role === "PARTICIPANT" && normalizeNik_(p.NIK) === s.nik) ||
      (s.role === "CLP" && normalizeRegion_(p.Region) === s.region) ||
      (s.role === "SUPERIOR" && normalizeNik_(p.Superior_NIK) === s.nik)
    );
  });
}
function participantTarget_(s, target) {
  target = typeof target === "object" ? target : { nik: target };
  var nik = normalizeNik_(target.nik),
    list = scopedParticipants_(s).filter(function (r) {
      return normalizeNik_(r.data.NIK) === nik;
    });
  if (target.program)
    list = list.filter(function (r) {
      return (
        String(r.data.Program).toUpperCase() ===
        String(target.program).toUpperCase()
      );
    });
  if (target.batch !== undefined)
    list = list.filter(function (r) {
      return sameBatch_(r.data.Batch, target.batch);
    });
  if (target.enrollmentId) {
    var er = enrollmentById_(target.enrollmentId);
    if (!er || normalizeNik_(er.data.NIK) !== nik)
      throw new Error("Enrollment tidak ditemukan.");
    list = list.filter(function (r) {
      return (
        entityKey_(r.data.NIK, r.data.Program, r.data.Batch, r.data.Segment) ===
        entityKey_(er.data.NIK, er.data.Program, er.data.Batch, er.data.Segment)
      );
    });
  }
  var r = uniqueMatch_(list, "peserta target");
  if (!r) throw new Error("Peserta tidak ditemukan atau di luar akses Anda.");
  return r;
}
function enrollmentIndex_(light) {
  var c = context_(), memo = light ? "enrollmentsLight" : "enrollments";
  if (light && c.memo.enrollments) return c.memo.enrollments;
  if (!c.memo[memo]) {
    var idx = { byId: {}, byKey: {}, byNik: {} };
    (light ? readSummaryRows_(SHEET_ENROLLMENTS,ENROLLMENT_LIST_FIELDS_) : readAllAsObjects_(SHEET_ENROLLMENTS)).forEach(function (r) {
      var d = r.data,
        n = normalizeNik_(d.NIK),
        k = entityKey_(n, d.Program, d.Batch, d.Segment);
      if (idx.byKey[k]) throw new Error("Enrollment duplikat untuk " + k);
      if (!d.EnrollmentID)
        throw new Error(
          "EnrollmentID kosong pada baris " +
            r.rowIndex +
            ". Jalankan setupLdpSystem.",
        );
      if (idx.byId[d.EnrollmentID])
        throw new Error("EnrollmentID duplikat: " + d.EnrollmentID);
      idx.byKey[k] = r;
      idx.byId[d.EnrollmentID] = r;
      (idx.byNik[n] || (idx.byNik[n] = [])).push(r);
    });
    c.memo[memo] = idx;
  }
  return c.memo[memo];
}
function enrollmentById_(id) {
  return enrollmentIndex_().byId[String(id)] || null;
}
function enrollmentFor_(p, light) {
  return (
    enrollmentIndex_(light).byKey[entityKey_(p.NIK, p.Program, p.Batch, p.Segment)] || null
  );
}
// Repair only roster identities that the authenticated caller may already access.
// Established records take the read-only path; no setup or schema formatting here.
function ensureParticipantEnrollments_(session, nik) {
  var targetNik = normalizeNik_(nik), keys = {}, repairs = [], existing = {};
  var rows = scopedParticipants_(session).filter(function(r) {
    return normalizeNik_(r.data.NIK) === targetNik;
  });
  if (!rows.length) throw new Error("Peserta tidak ditemukan atau di luar akses Anda.");
  rows.forEach(function(r) {
    var d = r.data, key = entityKey_(d.NIK, d.Program, d.Batch, d.Segment);
    if (keys[key]) throw new Error("Participants duplikat: " + key);
    if (activePrograms_().indexOf(String(d.Program || "").toUpperCase()) < 0 || !normalizeBatchKey_(d.Batch))
      throw new Error("Program/batch peserta belum lengkap. Perbaiki melalui Participant Control.");
    keys[key] = r;
  });
  readAllAsObjects_(SHEET_ENROLLMENTS).forEach(function(r) {
    var d = r.data, key = entityKey_(d.NIK, d.Program, d.Batch, d.Segment);
    if (!keys[key]) return;
    if (existing[key]) throw new Error("Enrollment duplikat: " + key);
    existing[key] = r;
  });
  Object.keys(keys).forEach(function(key) {
    var p = keys[key].data, er = existing[key];
    if (!er || !er.data.EnrollmentID || p.CurrentEnrollmentId !== er.data.EnrollmentID) repairs.push(p);
  });
  repairs.forEach(function(p) {
    ensureImportEnrollments_(String(p.Program).toUpperCase(), p.Batch, targetNik, session);
  });
  return repairs.length;
}
function authorizeEnrollment_(s, target) {
  // Authorize the supplied identity before any repair; never derive it from untrusted data.
  var p = participantTarget_(s, target);
  ensureParticipantEnrollments_(s, p.data.NIK);
  p = participantTarget_(s, target);
  var er = enrollmentFor_(p.data);
  if (!er) throw new Error("Program peserta belum terhubung. Muat ulang detail peserta.");
  return { participant: p, enrollment: er };
}
function hasMentoringEvidence_(items) {
  return (items || []).some(function (m) {
    return !!String(m.FileUrl || m.DriveFileId || "").trim();
  });
}
function legacyRecordScope_(record, p, light) {
  if (normalizeSegment_(record.Segment)!==normalizeSegment_(p.Segment)) return false;
  if (record.EnrollmentID)
    return (
      record.EnrollmentID ===
      (enrollmentFor_(p, light) || { data: {} }).data.EnrollmentID
    );
  var list = (enrollmentIndex_(light).byNik[normalizeNik_(p.NIK)] || []).filter(function(r){return normalizeSegment_(r.data.Segment)===normalizeSegment_(p.Segment);});
  return list.length === 1;
}
function workspaceIndex_() {
  var c = context_();
  if (!c.memo.workspaceIndex) {
    var index = { subs: {}, evidence: {} };
    readAllAsObjects_(SHEET_SUBMISSIONS).forEach(function (r) {
      var n = normalizeNik_(r.data.NIK);
      (index.subs[n] || (index.subs[n] = [])).push(r.data);
    });
    readAllAsObjects_(SHEET_MENTORING).forEach(function (r) {
      var n = normalizeNik_(r.data.NIK);
      (index.evidence[n] || (index.evidence[n] = [])).push(r.data);
    });
    c.memo.workspaceIndex = index;
  }
  return c.memo.workspaceIndex;
}
function overviewPanelEligible_(p) {
  var index=workspaceIndex_(),n=normalizeNik_(p.NIK),phases={},mentors={};
  (index.subs[n] || []).some(function(d) {
    if(legacyRecordScope_(d,p,true)) phases[String(d.Phase).toUpperCase()]=true;
    return phases.PROPOSAL && phases.PROGRESS && phases.FINAL;
  });
  if(!phases.PROPOSAL || !phases.PROGRESS || !phases.FINAL) return false;
  (index.evidence[n] || []).some(function(d) {
    if(!legacyRecordScope_(d,p,true) || !hasMentoringEvidence_([d])) return false;
    var role=String(d.MentorRole || '').toUpperCase();
    if(role==='CLP') mentors.CLP=true;
    if(['ASH','TSH','LOB','SUPERIOR'].indexOf(role)>=0) mentors.SUPERIOR=true;
    return mentors.CLP && mentors.SUPERIOR;
  });
  return !!(mentors.CLP && mentors.SUPERIOR);
}
function workspace_(p) {
  var c=context_();workspaceIndex_();
  var n = normalizeNik_(p.NIK),
    out = { submissions: {}, mentoring: [], unassignedLegacy: 0 };
  (c.memo.workspaceIndex.subs[n] || []).forEach(function (d) {
    if (!legacyRecordScope_(d, p)) {
      out.unassignedLegacy++;
      return;
    }
    var ph = String(d.Phase).toUpperCase();
    (out.submissions[ph] || (out.submissions[ph] = [])).push(d);
  });
  Object.keys(out.submissions).forEach(function (ph) {
    out.submissions[ph].sort(function (a, b) {
      return Number(b.Version) - Number(a.Version);
    });
  });
  (c.memo.workspaceIndex.evidence[n] || []).forEach(function (d) {
    if (legacyRecordScope_(d, p)) out.mentoring.push(d);
    else out.unassignedLegacy++;
  });
  out.mentoring.sort(function (a, b) {
    return timeValue_(b.MentoringDate) - timeValue_(a.MentoringDate);
  });
  var missing = ["PROPOSAL", "PROGRESS", "FINAL"]
    .filter(function (ph) {
      return !out.submissions[ph] || !out.submissions[ph].length;
    })
    .map(function (ph) {
      return ph + " belum diunggah";
    });
  ["SUPERIOR","CLP"].forEach(function(role) {
    var records=out.mentoring.filter(function(m) {var r=String(m.MentorRole||"").toUpperCase();return role==="CLP"?r==="CLP":["ASH","TSH","LOB","SUPERIOR"].indexOf(r)>=0;});
    if(!hasMentoringEvidence_(records))missing.push("Bukti mentoring "+role+" belum diunggah");
  });
  out.eligibility = { eligible: !missing.length, missing: missing };
  return out;
}
function certificateFor_(p) {
  var er = enrollmentFor_(p, true);
  if (!er) return null;
  var c = context_();
  if (!c.memo.certIndex) {
    var index = {};
    readAllAsObjects_(SHEET_CERTIFICATES).forEach(function (r) {
      if (r.data.EnrollmentID)
        (index[r.data.EnrollmentID] || (index[r.data.EnrollmentID] = [])).push(
          r,
        );
    });
    c.memo.certIndex = index;
  }
  return uniqueMatch_(
    c.memo.certIndex[er.data.EnrollmentID] || [],
    "sertifikat enrollment",
  );
}
function participantRow_(p, options) {
  var er = enrollmentFor_(p, !!(options && options.light)),
    e = er ? er.data : {},
    w = options && options.light ? {eligibility:{eligible:null,missing:[]}} : workspace_(p);
  return {
    segment: normalizeSegment_(p.Segment),
    nik: normalizeNik_(p.NIK),
    name: p.Name,
    jobTitle: p.JobTitle,
    location: p.Location || "",
    cityName: p.CityName || p.City || "",
    region: p.Region,
    program: p.Program,
    batch: p.Batch,
    enrollmentId: e.EnrollmentID || "",
    recordVersion: e.RecordVersion || "",
    currentPhase: normalizeCurrentPhase_(
      e.CurrentPhase || p.CurrentPhase || inferCurrentPhase_(p),
    ),
    participantStatus: normalizeParticipantStatus_(p.ParticipantStatus),
    programTrack: normalizeProgramTrack_(p.ProgramTrack),
    phase1Score: numberOrBlank_(e.Phase1Score),
    phase2Score: numberOrBlank_(e.Phase2Score),
    phase3Score: numberOrBlank_(e.Phase3Score),
    finalScore: numberOrBlank_(e.FinalScore),
    finalStatus: e.ProgramStatus || "",
    result: computeResult_({
      ProgramStatus: e.ProgramStatus,
      ParticipantStatus: p.ParticipantStatus,
    }),
    eligible: w.eligibility.eligible,
    eligibilityKnown: !(options && options.light),
    missingRequirements: w.eligibility.missing,
    superiorNik: p.Superior_NIK || "",
    clpNik: p.CLP_NIK || "",
    updatedAt: e.UpdatedAt || p.UpdatedAt,
  };
}
function programDashboard_(p, s, includeDetails) {
  var er = enrollmentFor_(p),
    e = er ? er.data : {},
    sum = parseJson_(e.AssessmentSummary, {}),
    cert = certificateFor_(p),
    c = cert && cert.data;
  var out = {
    segment: normalizeSegment_(p.Segment),
    enrollmentId: e.EnrollmentID || "",
    program: p.Program,
    batch: p.Batch,
    region: e.Region || p.Region,
    currentPhase: normalizeCurrentPhase_(
      e.CurrentPhase || p.CurrentPhase || inferCurrentPhase_(p),
    ),
    programStatus: e.ProgramStatus || "",
    phase1: {
      score: numberOrBlank_(e.Phase1Score),
      available: e.Phase1Score !== "" && e.Phase1Score !== undefined,
      remedials: (sum["PHASE 1"] || {}).remedials || [],
      remedialKnown: !!sum["PHASE 1"],
    },
    phase2: {
      score: numberOrBlank_(e.Phase2Score),
      available: e.Phase2Score !== "" && e.Phase2Score !== undefined,
      remedials: (sum["PHASE 2"] || {}).remedials || [],
      remedialKnown: !!sum["PHASE 2"],
    },
    phase3: {
      score: numberOrBlank_(e.Phase3Score),
      available: e.Phase3Score !== "" && e.Phase3Score !== undefined,
    },
    behaviourObservation:
      e.BehaviourScore !== "" && e.BehaviourScore !== undefined
        ? e.BehaviourScore
        : numberOrBlank_(p.BehaviourObservation),
    behaviourSummary: parseJson_(e.BehaviourSummary, {}),
    panelScore: numberOrBlank_(e.PanelScore) !== "" ? numberOrBlank_(e.PanelScore) : numberOrBlank_(p.PanelScore),
    panelSummary: parseJson_(e.PanelSummary, {}),
    finalSummary: parseJson_(e.FinalSummary, {}),
    panelDate: p.PanelDate || "",
    finalScore: numberOrBlank_(e.FinalScore),
    finalStatus: e.ProgramStatus || "",
    result: computeResult_({
      ProgramStatus: e.ProgramStatus,
      ParticipantStatus: p.ParticipantStatus,
    }),
    recordVersion: e.RecordVersion || "",
    certificateUrl:
      c && (s.role === "ADMIN" || parseBoolean_(c.Published))
        ? c.FileOrLink || ""
        : "",
    certificatePending: !c || !c.FileOrLink,
  };
  // Older sheets may store the existing observation on a 0–100 scale.
  // Preserve its number and identify the scale; do not silently convert old data.
  out.behaviourScale = numberOrBlank_(out.behaviourObservation) > 5 ? 100 : 5;
  if (context_().memo.liveAssessmentReads || context_().memo.refreshAssessmentSummary) {
    var anyLiveAssessment=false;
    ["PHASE 1","PHASE 2","PHASE 3"].forEach(function(phase,i) {
      var target=out["phase"+(i+1)],a;
      try {a=readAssessmentDetail_(p.NIK,phase,p.Program,p.Batch);}
      catch(error) {
        if(!context_().memo.refreshAssessmentSummary) throw error;
        target.error=true;target.score="";target.available=false;out.assessmentError=true;
        console.warn(JSON.stringify({event:"ASSESSMENT_READ_FAILED",program:p.Program,phase:phase,message:String(error.message || error)}));
        return;
      }
      participantFallbackAssessment_(a,e);
      if (!a.available) return;
      anyLiveAssessment=true;
      if(phase==="PHASE 3" && !a.completedModules && numberOrBlank_(target.score)!=="")return;
      sum[phase]=Object.assign({},sum[phase]||{},{completedModules:a.completedModules,totalModules:a.totalModules,aggregateFallback:!!a.participantFallback});
      target.rawCompleted=a.completedModules;target.participantFallback=!!a.participantFallback;target.score=a.averageScore;target.available=a.completedModules>0||!!a.participantFallback;
      target.remedials=a.modules.filter(function(m){return m.result==="REMEDIAL REQUIRED";});target.remedialKnown=true;
    });
    if(anyLiveAssessment)out.finalScore=finalGrade_(Object.assign({},e,{Phase1Score:out.phase1.score,Phase2Score:out.phase2.score,Phase3Score:out.phase3.score}),p,sum).score;
    if(out.assessmentError) out.finalScore="";
  }
  legacyDashboard_(out,e,p);
  participantFallbackDashboard_(out,e,p);
  out.progressStages = buildProgressStages_(out.currentPhase);
  ["phase1", "phase2"].forEach(function (k) {
    out[k].remedialCount = out[k].remedials.length;
  });
  if (includeDetails) {
    out.phase1.detail = readAssessmentDetail_(
      p.NIK,
      "PHASE 1",
      p.Program,
      p.Batch,
    );
    out.phase2.detail = readAssessmentDetail_(
      p.NIK,
      "PHASE 2",
      p.Program,
      p.Batch,
    );
  }
  return out;
}
// Program Track is assigned per enrollment in Participants, independently of scores.
function normalizeProgramTrack_(value) {
  var track=String(value || "").trim().toUpperCase();
  return track === "STANDARISASI" ? "STANDARDISASI" : track;
}
function matchesProgramTrack_(value, filter) {
  var track=normalizeProgramTrack_(value), wanted=normalizeProgramTrack_(filter);
  if(!wanted) return true;
  if(wanted === "UNASSIGNED") return PROGRAM_TRACK_OPTIONS.indexOf(track)<0;
  if(PROGRAM_TRACK_OPTIONS.indexOf(wanted)<0) throw new Error("Jalur program tidak valid.");
  return track===wanted || (track==="BOTH" && wanted!=="BOTH");
}
function programStats_(participants) {
  return activePrograms_().map(function (program) {
    var unique = {},
      g = {
        program: program,
        total: 0,
        passed: 0,
        onGoing: 0,
        notPassed: 0,
        remedial: 0,
        promoted: 0,
        resigned: 0,
        bypassed: 0,
        talentPool: 0,
        standarisasi: 0,
        unassignedTrack: 0,
        bothTracks: 0,
        certified: 0,
        eligiblePanel: 0,
        regions: {},
      };
    participants.forEach(function (r) {
      var p = r.data;
      if (String(p.Program).toUpperCase() !== program) return;
      var key = entityKey_(p.NIK, p.Program, p.Batch, p.Segment);
      if (unique[key]) throw new Error("Data peserta duplikat " + key);
      unique[key] = true;
      var row = participantRow_(p,{light:true}),
        status = row.participantStatus,
        result = row.result;
      g.total++;
      var track=normalizeProgramTrack_(p.ProgramTrack);
      if(track === "TALENT" || track === "BOTH") g.talentPool++;
      if(track === "STANDARDISASI" || track === "BOTH") g.standarisasi++;
      if(track === "BOTH") g.bothTracks++;
      if(PROGRAM_TRACK_OPTIONS.indexOf(track)<0) g.unassignedTrack++;
      g.regions[p.Region] = (g.regions[p.Region] || 0) + 1;
      if (result === "PASSED") {
        g.passed++;
      } else if (result === "REMEDIAL") {
        g.remedial++;
        g.notPassed++;
      } else if (result === "NOT PASSED") g.notPassed++;
      else g.onGoing++;
      if (status === "PROMOSI") g.promoted++;
      if (status === "RESIGN") g.resigned++;
      if (status === "BYPASS") g.bypassed++;
      if (overviewPanelEligible_(p)) g.eligiblePanel++;
      var c = certificateFor_(p);
      if (c && c.data.FileOrLink && parseBoolean_(c.data.Published))
        g.certified++;
    });
    g.passRate = g.total ? Math.round((g.passed / g.total) * 100) : 0;
    g.regionList = Object.keys(g.regions)
      .sort()
      .map(function (region) {
        return { region: region, count: g.regions[region] };
      });
    delete g.regions;
    return g;
  });
}
function getDashboard_(token) {
  var s = requireSession_(token),
    rows = scopedParticipants_(s),
    base = {
      session: {
        nik: s.nik,
        name: s.name,
        role: s.role,
        roles: s.roles,
        region: s.region,
      },
      appName: getSetting_("APP_NAME") || "LDP Monitor",
      version: LDP_VERSION_,
      dataVersion: revision_(),
      roleSettings: getRoleSettings_(),
    };
  if (s.role === "PARTICIPANT") {
    if(rows.length)ensureParticipantEnrollments_(s,s.nik);
    rows=scopedParticipants_(s);
    context_().memo.liveAssessmentReads = true;
    base.programDashboards = rows.map(function (r) {
      return programDashboard_(r.data, s, false);
    });
    base.participant = rows.length ? participantRow_(rows[0].data) : null;
    if (!base.participant) throw new Error("Data peserta tidak ditemukan.");
  } else base.programStats = programStats_(rows);
  return base;
}
// Progress is separate from employment status and academic pass/fail.
// Reading this panel never repairs enrollment, calculates grades or writes sheets.
function programProgressRow_(p) {
  var er = enrollmentFor_(p, true), e = er ? er.data : {},
    phase = normalizeCurrentPhase_(e.CurrentPhase || p.CurrentPhase || inferCurrentPhase_(p)),
    result = String(e.ProgramStatus || p.FinalStatus || "").trim().toUpperCase(),
    participantStatus = normalizeParticipantStatus_(p.ParticipantStatus),
    complete = phase === "COMPLETED" ||
      normalizeCurrentPhase_(e.EnrollmentStatus) === "COMPLETED" ||
      isPassingStatus_(result) || result === "COMPLETED" ||
      ["TIDAK LULUS", "NOT PASSED", "FAIL", "FAILED"].indexOf(result) >= 0,
    progressStatus = complete ? "COMPLETE" : participantStatus === "RESIGN" ? "STOPPED" : "ON GOING";
  return {
    segment:normalizeSegment_(p.Segment),
    nik: normalizeNik_(p.NIK), name: String(p.Name || ""),
    jobTitle: p.JobTitle || "", region: p.Region || "",
    program: String(p.Program || "").trim().toUpperCase(), batch: String(p.Batch || ""),
    enrollmentId: e.EnrollmentID || "", currentPhase: phase,
    participantStatus: participantStatus, progressStatus: progressStatus,
    result: computeResult_({ProgramStatus: result, ParticipantStatus: participantStatus})
  };
}
function getProgramProgress_(token, filters) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN", "CLP", "SUPERIOR"]);
  filters = filters || {};
  if(filters.fresh) context_().memo.bypassSnapshots=true;
  var pageKey = cacheKey_("ldpprogress:", JSON.stringify([activeSegment_(),sourceRevision_([SHEET_PARTICIPANTS,SHEET_ENROLLMENTS,SHEET_SETTINGS]),
    s.role,s.nik,s.region || "",filters.programs || {}]));
  if (!filters.fresh) { var cached = readSnapshot_(pageKey); if (cached) return cached; }
  var grouped = {ASHDP: [], TSHDP: [], SLDP: []}, seen = Object.create(null);
  scopedParticipants_(s,true).forEach(function(r) {
    var program = String(r.data.Program || "").trim().toUpperCase();
    if (!grouped[program]) return;
    var key = entityKey_(r.data.NIK, program, r.data.Batch, r.data.Segment);
    if (seen[key]) throw new Error("Data peserta duplikat " + key);
    seen[key] = true;
    grouped[program].push(programProgressRow_(r.data));
  });
  var result = {segment:activeSegment_(),programs: activePrograms_().map(function(program) {
    var f = (filters.programs || {})[program] || {}, all = grouped[program],
      regions = Object.keys(all.reduce(function(a,r) { if(r.region) a[r.region]=true; return a; }, Object.create(null))).sort(),
      regional = all.filter(function(r) { return !f.region || r.region === f.region; }),
      batches = Object.create(null);
    regional.forEach(function(r) {
      var key = normalizeBatchKey_(r.batch), b = batches[key];
      if (!b) b = batches[key] = {batch:r.batch, total:0, onGoing:0, complete:0, stopped:0};
      b.total++;
      b[r.progressStatus === "COMPLETE" ? "complete" : r.progressStatus === "STOPPED" ? "stopped" : "onGoing"]++;
    });
    var batchList = Object.keys(batches).map(function(k) {
      var b=batches[k];
      b.status=b.onGoing ? "ON GOING" : b.complete === b.total ? "COMPLETE" : "STOPPED";
      return b;
    }).sort(function(a,b) { return a.batch.localeCompare(b.batch, undefined, {numeric:true}); });
    var allowed = Object.create(null);
    batchList.forEach(function(b) { if(!f.batchStatus || b.status === f.batchStatus) allowed[normalizeBatchKey_(b.batch)]=true; });
    var scope = regional.filter(function(r) {
      return allowed[normalizeBatchKey_(r.batch)] && (!f.batch || sameBatch_(r.batch,f.batch));
    });
    var summary={total:scope.length,onGoing:0,complete:0,stopped:0};
    scope.forEach(function(r) {
      summary[r.progressStatus === "COMPLETE" ? "complete" : r.progressStatus === "STOPPED" ? "stopped" : "onGoing"]++;
    });
    summary.completionRate=summary.total ? Math.round(summary.complete/summary.total*100) : 0;
    var q=String(f.search || "").trim().toLowerCase(), rows=scope.filter(function(r) {
      return (!f.progressStatus || r.progressStatus === f.progressStatus) &&
        (!q || (r.nik+" "+r.name).toLowerCase().indexOf(q)>=0);
    });
    rows.sort(function(a,b) { return a.name.localeCompare(b.name) || a.nik.localeCompare(b.nik) || a.batch.localeCompare(b.batch); });
    var page=paginate_(rows,{page:f.page,pageSize:25});
    page.program=program; page.summary=summary; page.regions=regions;
    page.batches=batchList.filter(function(b) {return !f.batchStatus || b.status === f.batchStatus;});
    return page;
  })};
  putSnapshot_(pageKey,result,30);
  return result;
}
function getProgramProgress(t, f) {
  return api_("getProgramProgress", function() { return getProgramProgress_(t,f); }, arguments);
}


// Bootstrap reads roster/enrollment metadata only. It never reads grades,
// certificates or workspace, and never repairs/writes enrollment during login.
function dashboardBase_(s) {
  return {
    segment:activeSegment_(),availableSegments:s.role===ROLE_PARTICIPANT ? participantSegments_(s) : ["B2B","B2C"],
    session:{nik:s.nik,name:s.name,role:s.role,roles:s.roles,region:s.region},
    appName:getSetting_("APP_NAME") || "Leadership Development Program",
    version:LDP_VERSION_, dataVersion:revision_(), roleSettings:getRoleSettings_()
  };
}
function participantSegments_(s) {
  var c=context_(),previous=c.segment;c.segment="ALL";
  try {return Object.keys(scopedParticipants_(s,true).reduce(function(a,r){a[normalizeSegment_(r.data.Segment)]=true;return a;},{}));}
  finally {c.segment=previous;}
}
function bootstrapProgram_(p) {
  var er=enrollmentFor_(p,true),e=er ? er.data : {},
    phase=normalizeCurrentPhase_(e.CurrentPhase || p.CurrentPhase || inferCurrentPhase_(p));
  return {
    enrollmentId:e.EnrollmentID || "", segment:normalizeSegment_(p.Segment), program:p.Program, batch:p.Batch,
    region:e.Region || p.Region, currentPhase:phase,
    programStatus:e.ProgramStatus || p.FinalStatus || "",
    result:computeResult_({ProgramStatus:e.ProgramStatus || p.FinalStatus,ParticipantStatus:p.ParticipantStatus}),
    recordVersion:e.RecordVersion || "", detailsPending:true,
    phase1:{score:"",available:false,remedials:[],remedialCount:0},
    phase2:{score:"",available:false,remedials:[],remedialCount:0},
    phase3:{score:"",available:false}, finalScore:"", panelScore:"",
    behaviourObservation:"", behaviourScale:5, certificateUrl:"",
    certificatePending:true, progressStages:buildProgressStages_(phase)
  };
}
function dashboardBootstrap_(s) {
  var base=dashboardBase_(s);
  base.detailsPending=true;
  if(activeSegment_()==="ALL"){base.segmentChoiceRequired=true;base.detailsPending=false;base.programStats=[];base.programStatsPending=false;base.programDashboards=[];return base;}
  if(s.role === ROLE_PARTICIPANT) {
    var rows=scopedParticipants_(s,true);
    if(!rows.length) throw new Error("Data peserta tidak ditemukan.");
    base.participant=participantRow_(rows[0].data,{light:true});
    base.programDashboards=rows.map(function(r) {return bootstrapProgram_(r.data);});
    var personal=readSnapshot_(participantDashboardCacheKey_(s));
    if(personal) {Object.assign(base,personal);base.summaryCached=true;}
  } else {
    var stats=readSnapshot_(overviewCacheKey_(s));
    base.programStats=stats || [];base.programStatsPending=!stats;
    if(stats) {base.detailsPending=false;base.summaryCached=true;}
  }
  return base;
}
function participantDashboardCacheKey_(s) {
  return cacheKey_('ldppdash:',[LDP_VERSION_,activeSegment_(),revision_(),s.nik].join('|'));
}
function overviewCacheKey_(s) {
  var scope=s.role===ROLE_ADMIN ? 'ALL' : s.role===ROLE_CLP ? s.region : s.nik;
  return cacheKey_('ldpoverview:',[activeSegment_(),sourceRevision_([SHEET_PARTICIPANTS,SHEET_ENROLLMENTS,
    SHEET_SUBMISSIONS,SHEET_MENTORING,SHEET_CERTIFICATES,SHEET_SETTINGS]),s.role,scope].join('|'));
}
function getDashboardBootstrap(t) {
  return api_("getDashboardBootstrap",function() {return dashboardBootstrap_(requireSession_(t));}, arguments);
}
function getParticipantDashboardData(t, options) {
  return api_("getParticipantDashboardData",function() {
    var s=requireSession_(t);
    requireRole_(s,[ROLE_PARTICIPANT]);
    context_().memo.bypassSnapshots=!!(options && options.fresh);
    var key=participantDashboardCacheKey_(s),cached=options && options.fresh ? null : readSnapshot_(key);
    if(cached) return cached;
    var rows=scopedParticipants_(s);
    if(!rows.length) throw new Error("Data peserta tidak ditemukan.");
    // Source snapshots last at most 30s, are shared between participants,
    // and are invalidated by the existing edit/change/import revisions.
    // Explicit refresh bypasses the source snapshots.
    context_().memo.refreshAssessmentSummary=true;
    context_().memo.liveAssessmentReads=!!(options && options.fresh);
    var result={programDashboards:rows.map(function(r) {return programDashboard_(r.data,s,false);}),
      dataVersion:revision_(),loadedAt:nowIso_(),detailsPending:false};
    if(!result.programDashboards.some(function(p){return p.assessmentError;})) putSnapshot_(key,result,30);
    else CacheService.getScriptCache().remove(key);
    return result;
  }, arguments);
}
function getOverviewStats(t, options) {
  return api_("getOverviewStats",function() {
    var s=requireSession_(t);
    requireRole_(s,[ROLE_ADMIN,ROLE_CLP,ROLE_SUPERIOR]);
    context_().memo.bypassSnapshots=!!(options && options.fresh);
    // Recheck authorization before looking in cache. Different access scopes
    // never share aggregates, even if the supplied client filters are forged.
    var key=overviewCacheKey_(s),
      stats=options && options.fresh ? null : readSnapshot_(key);
    if(!stats) {stats=programStats_(scopedParticipants_(s,true));putSnapshot_(key,stats,30);}
    return {programStats:stats,programStatsPending:false,detailsPending:false,dataVersion:revision_()};
  }, arguments);
}
// Established accounts can continue using their freshly validated Users row
// when an import/write holds the lock. Unknown accounts still require sync.
function opportunisticUserSync_() {
  if(CacheService.getScriptCache().get(userImportCheckKey_())) return {ok:true,skipped:true};
  var c=context_(),lock=LockService.getScriptLock();
  if(typeof lock.tryLock !== "function") return autoSyncUserImport_();
  if(!lock.tryLock(0)) return {ok:true,deferred:true};
  c.lockDepth++;
  try {return autoSyncUserImport_();}
  finally {c.lockDepth--;try {flushRevision_();} finally {lock.releaseLock();}}
}


function getParticipantPage_(token, f) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN", "CLP", "SUPERIOR"]);
  f = f || {};
  var pageCacheKey = f.light === true ? cacheKey_("ldppage:", JSON.stringify([
    activeSegment_(),sourceRevision_([SHEET_PARTICIPANTS,SHEET_ENROLLMENTS,SHEET_SETTINGS]),
    s.role, s.nik, s.region || "", String(f.search || "").toLowerCase(),
    f.program || "", f.region || "", f.batch || "", f.status || "", f.phase || "", normalizeProgramTrack_(f.track),
    f.page || 1, f.pageSize || 30
  ])) : "";
  if (f.fresh === true) context_().memo.bypassSnapshots = true;
  else if (pageCacheKey) {
    var cachedPage = readSnapshot_(pageCacheKey);
    if (cachedPage && cachedPage.format === "LDP_PAGE_1") return cachedPage.data;
  }
  var scoped = scopedParticipants_(s, f.light === true);
  var q = String(f.search || "").toLowerCase(),
    rows = scoped
      .filter(function (r) {
        var p=r.data;
        return (!q || (normalizeNik_(p.NIK) + " " + p.Name).toLowerCase().indexOf(q) >= 0) &&
          (!f.program || p.Program === f.program) && (!f.region || p.Region === f.region) &&
          (!f.batch || sameBatch_(p.Batch,f.batch)) &&
          matchesProgramTrack_(p.ProgramTrack,f.track) &&
          (!f.status || normalizeParticipantStatus_(p.ParticipantStatus) === f.status) &&
          (!f.phase || normalizeCurrentPhase_((enrollmentFor_(p) || {data:{}}).data.CurrentPhase || p.CurrentPhase || inferCurrentPhase_(p)) === f.phase);
      });
  rows.sort(function (a, b) {
    return String(a.data.Name).localeCompare(String(b.data.Name));
  });
  var out = paginate_(rows, f), batches = {};
  out.rows = out.rows.map(function(r) {return participantRow_(r.data,{light:f.light === true});});
  scoped.forEach(function(r) {
    var d=r.data;
    if ((!f.program || d.Program===f.program) && (!f.region || d.Region===f.region) && matchesProgramTrack_(d.ProgramTrack,f.track) && normalizeBatchKey_(d.Batch))
      batches[normalizeBatchKey_(d.Batch)]=String(d.Batch);
  });
  out.batches=Object.keys(batches).map(function(k){return batches[k];}).sort(function(a,b){return a.localeCompare(b,undefined,{numeric:true});});
  out.regions = Object.keys(
    scoped.reduce(function (a, r) {
      a[r.data.Region] = true;
      return a;
    }, {}),
  ).sort();
  if (pageCacheKey) putSnapshot_(pageCacheKey, {format:"LDP_PAGE_1",data:out}, 30);
  return out;
}
function getParticipantFullDetail_(token, nik) {
  var s = requireSession_(token),
    rows = scopedParticipants_(s).filter(function (r) {
      return normalizeNik_(r.data.NIK) === normalizeNik_(nik);
    });
  if (!rows.length) throw new Error("Peserta di luar akses Anda.");
  ensureParticipantEnrollments_(s, nik);
  // Repair can reload request tables inside its lock; use the refreshed roster.
  rows = scopedParticipants_(s).filter(function(r) {
    return normalizeNik_(r.data.NIK) === normalizeNik_(nik);
  });
  context_().memo.liveAssessmentReads = true;
  var p = rows[0].data;
  return {
    participant: {
      NIK: p.NIK,
      Name: p.Name,
      JobTitle: p.JobTitle,
      Location: p.Location || "",
      CityName: p.CityName || p.City || "",
      Region: p.Region,
      ParticipantStatus: normalizeParticipantStatus_(p.ParticipantStatus),
    },
    enrollments: rows.map(function (r) {
      return programDashboard_(r.data, s, false);
    }),
  };
}
function getParticipantWorkspace_(token, target) {
  var s = requireSession_(token);
  if(target && target.fresh) context_().memo.bypassSnapshots=true;
  var t = authorizeEnrollment_(s, target);
  var w = workspace_(t.participant.data);
  w.participant = participantRow_(t.participant.data,{light:true});
  w.participant.eligible=w.eligibility.eligible;
  w.participant.eligibilityKnown=true;
  w.participant.missingRequirements=w.eligibility.missing;
  w.enrollment = programDashboard_(t.participant.data, s, false);
  return w;
}

// ---------------------------------------------------------------------------
// Assessment reads use exact NIK + program + batch. No shifted-cell fallback.
// ---------------------------------------------------------------------------
function assessmentBatchCompatible_(source, target) {
  if (sameBatch_(source, target)) return true;
  var a = normalizeBatchKey_(source), b = normalizeBatchKey_(target),
    ay = /^(.*) (20\d{2})$/.exec(a), by = /^(.*) (20\d{2})$/.exec(b);
  // Never collapse two explicit, different years into the same batch.
  if (ay && by) return false;
  return !!((ay && ay[1] === b) || (by && by[1] === a));
}
// For large Phase 1/2 sheets, a fresh popup reads a narrow NIK column and only
// the matching rows. Small sheets and many-attempt histories keep the bulk path.
function assessmentHeaderAt_(grid, index, cfg) {
  var h=(grid[index] || []).map(String),nc=findHeaderIndex_(h,cfg.nikHeaders);
  if(nc<0) return null;
  var modules=cfg.modules.filter(function(m) {return findHeaderIndex_(h,m.headers)>=0;});
  return modules.length ? {headers:h,nik:nc,batch:findHeaderIndex_(h,cfg.batchHeaders),
    program:findHeaderIndex_(h,["Program"]),index:index} : null;
}
function targetedAssessmentGrid_(sh,cfg,nik,last,cols) {
  if(last*cols<=40000) return null;
  var c=context_(),probeRows=Math.min(last,Math.max(10,cfg.headerRow || 1)),
    probe=sh.getRange(1,1,probeRows,cols).getValues(),
    header=assessmentHeaderAt_(probe,(cfg.headerRow || 1)-1,cfg);
  c.reads++;c.cells+=probeRows*cols;
  if(!header) for(var i=0;i<Math.min(probe.length,10);i++) {
    header=assessmentHeaderAt_(probe,i,cfg);if(header) break;
  }
  if(!header) return null;
  var start=Math.max(header.index+2,header.index===(cfg.headerRow || 1)-1 ? cfg.dataStartRow : header.index+2),
    grid=[],matches=[],target=normalizeNik_(nik);
  grid[header.index]=probe[header.index];
  if(last<start) return grid;
  var niks=sh.getRange(start,header.nik+1,last-start+1,1).getDisplayValues();
  c.reads++;c.cells+=last-start+1;
  niks.forEach(function(row,i) {if(normalizeNik_(row[0])===target)matches.push({row:start+i,nik:row[0]});});
  if(matches.length>12) return null;
  matches.forEach(function(m) {
    var row=sh.getRange(m.row,1,1,cols).getValues()[0];
    c.reads++;c.cells+=cols;
    // Recheck the identity: a concurrent row move must not disclose another NIK.
    var current=normalizeNik_(row[header.nik]),expected=normalizeNik_(m.nik);
    var formattedNumber=typeof row[header.nik]==="number" && /^\d+$/.test(expected) &&
      current===expected.replace(/^0+(?=\d)/,"");
    if(current!==expected && !formattedNumber)
      throw new Error("Posisi nilai berubah. Silakan buka kembali detail phase.");
    row[header.nik]=m.nik;
    grid[m.row-1]=row;
  });
  return grid;
}

function assessmentTable_(program, phase) {
  var sourceCfg = getAssessmentConfig_(program, phase);
  if (!sourceCfg) return null;
  var cfg = Object.assign({}, sourceCfg), name = getSetting_(cfg.settingKey) || cfg.defaultSheet,
    c = context_(), mk = "assessment:" + name + "|" + program + "|" + phase;
  if (c.memo[mk]) return c.memo[mk];
  Object.keys(PROGRAM_PHASE_ASSESSMENT_CONFIG).forEach(function(k) {
    var other=PROGRAM_PHASE_ASSESSMENT_CONFIG[k];
    if((other.segment || "B2B")!==(cfg.segment || "B2B") && String(getSetting_(other.settingKey)||other.defaultSheet).toLowerCase()===String(name).toLowerCase())
      throw new Error("Sheet nilai B2B dan B2C harus terpisah. Periksa pengaturan sumber " + name + ".");
  });
  var sh = getSheet_(name), key = cacheKey_("ldpa:", LDP_VERSION_ + "|" + name + "|" + revision_()),
    grid = !c.memo.liveAssessmentReads && !c.dirty[name] && !c.bulk ? readSnapshot_(key) : null,
    last,cols;
  if(c.memo.liveAssessmentReads && c.memo.assessmentTargetNik) {
    last=sh.getLastRow();cols=sh.getLastColumn();
    grid=targetedAssessmentGrid_(sh,cfg,c.memo.assessmentTargetNik,last,cols);
  }
  if (!grid) {
    if(last===undefined) last=sh.getLastRow();
    if(cols===undefined) cols=sh.getLastColumn();
    grid = last && cols ? sh.getRange(1,1,last,cols).getValues() : [];
    c.reads++; c.cells += last * cols;
    if (!c.dirty[name] && !c.bulk) putSnapshot_(key, grid, 30);
  }
  function headerAt(index) {return assessmentHeaderAt_(grid,index,cfg);}
  var foundHeader = headerAt((cfg.headerRow || 1) - 1);
  if (!foundHeader) for (var hi=0;hi<Math.min(grid.length,10);hi++) {foundHeader=headerAt(hi);if(foundHeader)break;}
  if (!foundHeader) throw new Error("Header NIK dan kompetensi pada " + name + " belum dikenali. Periksa nama sheet dan baris header.");
  var headers=foundHeader.headers,nc=foundHeader.nik,bc=foundHeader.batch,pc=foundHeader.program,byKey={},byNik={};
  cfg.headerRow=foundHeader.index+1;
  cfg.dataStartRow=Math.max(cfg.headerRow+1,foundHeader.index===(sourceCfg.headerRow||1)-1 ? sourceCfg.dataStartRow : cfg.headerRow+1);
  for (var i=cfg.dataStartRow-1;i<grid.length;i++) {
    if(!grid[i]) continue;
    var n=normalizeNik_(grid[i][nc]);
    if (!n || (pc>=0 && String(grid[i][pc]||"").trim() && String(grid[i][pc]).trim().toUpperCase()!==program)) continue;
    var sourceBatch=bc>=0?grid[i][bc]:"",k=entityKey_(n,program,sourceBatch);
    (byKey[k]||(byKey[k]=[])).push(i+1);
    (byNik[n]||(byNik[n]=[])).push({row:i+1,batch:sourceBatch});
  }
  return (c.memo[mk]={cfg:cfg,sheet:sh,name:name,headers:headers,grid:grid,byKey:byKey,byNik:byNik,batchColumn:bc});
}
function findAssessmentRow_(sheet, cfg, nik, batch) {
  var t=assessmentTable_(cfg.program,cfg.phase),n=normalizeNik_(nik),k=entityKey_(n,cfg.program,batch),rows=t.byKey[k]||[],mode="EXACT",reason="";
  if (rows.length>1) throw new Error("Assessment duplikat untuk " + k + " pada " + t.name);
  if (!rows.length) {
    var candidates=t.byNik[n]||[], identities={};
    [SHEET_PARTICIPANTS,SHEET_ENROLLMENTS].forEach(function(name) {
      readAllAsObjects_(name).forEach(function(r) {
        var d=r.data;
        if(normalizeNik_(d.NIK)===n && String(d.Program).toUpperCase()===cfg.program)
          identities[entityKey_(n,cfg.program,d.Batch, d.Segment)]=d.Batch;
      });
    });
    var eligible=candidates.filter(function(x) {
      if (normalizeBatchKey_(x.batch) && !assessmentBatchCompatible_(x.batch,batch)) return false;
      var matching=Object.keys(identities).filter(function(ik) {return !normalizeBatchKey_(x.batch)||assessmentBatchCompatible_(x.batch,identities[ik]);});
      return matching.length===1 && matching[0]===k;
    });
    if (eligible.length>1) throw new Error("Assessment duplikat untuk NIK " + n + " pada " + t.name + ". Lengkapi Batch agar nilai tidak tertukar.");
    if (eligible.length===1) {rows=[eligible[0].row];mode=normalizeBatchKey_(eligible[0].batch)?"BATCH_ALIAS":"LEGACY_NIK";}
    else reason=candidates.length
      ? 'Data nilai NIK ini ditemukan pada sheet ' + t.name + ', tetapi Batch belum cocok atau belum dapat dibedakan. Batch peserta: ' + batch + '. Batch di sheet: ' + candidates.map(function(x){return String(x.batch||"(kosong)");}).filter(function(v,i,a){return a.indexOf(v)===i;}).slice(0,5).join(', ') + '. Periksa kolom Batch sebelum menghubungkan nilai.'
      : 'NIK ' + n + ' belum ditemukan pada sheet ' + t.name + '. Periksa NIK dan nama sheet sumber nilai.';
  }
  return {rowNumber:rows[0]||0,headers:t.headers,headerRow:t.cfg.headerRow,matchMode:rows.length?mode:"UNMATCHED",message:reason};
}
function getPassingScore_() {
  var n = Number(getSetting_("PASSING_SCORE"));
  return isFinite(n) &&
    n >= 0 &&
    n <= 100 &&
    String(getSetting_("PASSING_SCORE")) !== ""
    ? n
    : 80;
}
function getQleapIndex_() {
  var c = context_();
  if (!c.memo.qleap) {
    var map = {},
      name = getSetting_("QLEAP_MAPPING_SHEET") || SHEET_QLEAP;
    readAllAsObjects_(name).forEach(function (r) {
      var d = r.data;
      map[
        String(d.Program).toUpperCase() +
          "|" +
          normalizeCurrentPhase_(d.Phase) +
          "|" +
          normalizeModuleKey_(d.Module)
      ] = d.QLEAPUrl || d["QLEAP URL"] || "";
    });
    c.memo.qleap = map;
  }
  return c.memo.qleap;
}
function normalizeModuleKey_(v) {
  return String(v || "")
    .trim()
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function qleapUrl_(program, phase, module) {
  return (
    getQleapIndex_()[
      String(program).toUpperCase() +
        "|" +
        normalizeCurrentPhase_(phase) +
        "|" +
        normalizeModuleKey_(module)
    ] || ""
  );
}
function readAssessmentDetail_(nik, phase, program, batch) {
  phase = normalizeCurrentPhase_(phase);
  program = String(program || "").toUpperCase();
  var key = "detail:" + entityKey_(nik, program, batch) + "|" + phase,
    c = context_();
  if (c.memo[key]) return c.memo[key];
  if (phase === "PHASE 3")
    return (c.memo[key] = readPhase3Assessment_(nik, program, batch));
  var t = assessmentTable_(program, phase);
  if (!t) throw new Error("Program atau phase assessment tidak valid.");
  var found = findAssessmentRow_(t.sheet, t.cfg, nik, batch),
    row = found.rowNumber ? t.grid[found.rowNumber - 1] : [],
    passing = getPassingScore_();
  var modules = t.cfg.modules.map(function (m) {
    var col = findHeaderIndex_(t.headers, m.headers);
    if (col < 0) throw new Error("Kolom modul " + m.label + " tidak tersedia.");
    var score = numberOrBlank_(row[col]);
    return {
      module: m.label,
      score: score,
      weight: m.weight === undefined ? null : m.weight,
      result:
        score === ""
          ? "NOT AVAILABLE"
          : score < passing
            ? "REMEDIAL REQUIRED"
            : "PASS",
      qleapUrl: qleapUrl_(program, phase, m.label),
      editable: true,
    };
  });
  var valid = modules.filter(function (m) {
      return m.score !== "";
    }),
    avg = assessmentAverage_(t.cfg, modules);
  return (c.memo[key] = {
    nik: normalizeNik_(nik),
    program: program,
    batch: batch,
    phase: phase,
    available: !!found.rowNumber,
    matchMode: found.matchMode,
    sheetName: t.name,
    rowNumber: found.rowNumber,
    modules: modules,
    finalScore: avg,
    averageScore: avg,
    passingScore: passing,
    completedModules: valid.length,
    totalModules: modules.length,
    completionStatus:
      valid.length === modules.length
        ? "COMPLETE"
        : valid.length
          ? "PARTIAL"
          : "NOT STARTED",
    message: found.rowNumber
      ? ""
      : found.message || "Data nilai belum cocok dengan identitas peserta pada sheet sumber.",
  });
}
function readPhase3Assessment_(nik, program, batch) {
  var c = context_(),
    name = getSetting_("PHASE3_SHEET") || "Phase 3";
  if (!c.memo.phase3) {
    var sh = getSheet_(name),
      n = sh.getLastRow(),
      m = sh.getLastColumn(),
      key = cacheKey_("ldpp3:", revision_());
    var grid = c.bulk && c.tables[name] ? [getHeaders_(sh)].concat(c.tables[name].map(function(r){return getHeaders_(sh).map(function(h){return r.data[h]===undefined?"":r.data[h];});})) : (!c.memo.liveAssessmentReads && !c.bulk && !c.dirty[name] ? readSnapshot_(key) : null);
    if (!grid) {
      grid = n && m ? sh.getRange(1, 1, n, m).getDisplayValues() : [];
      c.reads++; c.cells += n * m;
      if (!c.bulk && !c.dirty[name]) putSnapshot_(key, grid, 30);
    }
    c.memo.phase3 = grid;
  }
  var values = c.memo.phase3,
    h = (values[0] || []).map(headerKey_),
    target = normalizeNik_(nik),
    pi = h.indexOf("program"),
    bi = h.indexOf("batch"),
    si = h.indexOf("segment"),
    ni = h.indexOf("nik"),
    found = {},
    latest = {}, answers = {};
  var post = h.indexOf("posttest"),
    pre = h.indexOf("pretest"),
    ts = h.indexOf("updatedat");
  function take(type, raw, time) {
    var n = numberOrBlank_(raw);
    if (n === "") return;
    if (latest[type] === undefined || time >= latest[type]) {
      found[type] = n;
      latest[type] = time;
    }
  }
  if (ni >= 0 && pi >= 0 && bi >= 0) {
    values.slice(1).forEach(function (r) {
      if (
        normalizeNik_(r[ni]) !== target ||
        entityKey_(target, r[pi], r[bi], si>=0?r[si]:"B2B") !== entityKey_(target, program, batch)
      )
        return;
      var time = ts >= 0 ? timeValue_(r[ts]) : 0;
      if (pre >= 0) take("PRE-TEST", r[pre], time);
      if (post >= 0) { take("POST-TEST", r[post], time); if(time===latest["POST-TEST"]){var ac=h.indexOf("answersjson");answers=ac>=0?parseJson_(r[ac],{}):{};} }
    });
  }
  if (!Object.keys(found).length && activeSegment_()!=="B2C") {
    // Legacy form exports have no enrollment context. Never reuse them for two enrollments.
    var ers = enrollmentIndex_().byNik[target] || [];
    if (ers.length === 1) {
      var ncs = [],
        tcs = [],
        scs = [];
      h.forEach(function (k, i) {
        if (k === "nomorindukkaryawan") ncs.push(i);
        if (k === "pretestatauposttest") tcs.push(i);
        if (k === "score") scs.push(i);
      });
      values.slice(1).forEach(function (r, index) {
        ncs.forEach(function (nc, j) {
          if (normalizeNik_(r[nc]) === target)
            take(
              String(
                r[tcs[j] === undefined ? nc + 3 : tcs[j]] || "",
              ).toUpperCase(),
              r[scs[j] === undefined ? nc + 4 : scs[j]],
              index,
            );
        });
      });
    }
  }
  if(found["POST-TEST"]===undefined)found["POST-TEST"]="";
  var modules = Object.keys(found).map(function (k) {
      return {
        module: k,
        displayLabel: k === "POST-TEST" ? "Soft Competency" : "Pre-test (informasi)",
        score: found[k],
        assessed: k === "POST-TEST",
        result:
          k === "POST-TEST"
            ? found[k] === "" ? "NOT AVAILABLE" : found[k] < getPassingScore_()
              ? "REMEDIAL REQUIRED"
              : "PASS"
            : "INFO",
        editable: k === "POST-TEST",
        qleapUrl: k === "POST-TEST" ? qleapUrl_(program, "PHASE 3", "POST-TEST") : "",
      };
    }),
    score = found["POST-TEST"] !== undefined ? found["POST-TEST"] : "";
  return {
    nik: target,
    program: program,
    batch: batch,
    phase: "PHASE 3",
    available: Object.keys(latest).length > 0,
    answers: answers,
    breakdownNote: "Soft Competency: nilai total post-test Phase 3 dari file sumber.",
    modules: modules,
    finalScore: score,
    averageScore: score,
    completedModules: score === "" ? 0 : 1,
    totalModules: 1,
    completionStatus: score === "" ? "NOT STARTED" : "COMPLETE",
    passingScore: getPassingScore_(),
    editable: true,
    message: Object.keys(latest).length
      ? ""
      : "Nilai Phase 3 belum tersedia untuk program dan batch ini. Gunakan kolom NIK, Program, Batch, PreTest, PostTest pada sheet Phase 3.",
  };
}
function getParticipantAssessmentDetail_(token, nik, phase, program, batch) {
  var s = requireSession_(token),
    t = authorizeEnrollment_(s, { nik: nik, program: program, batch: batch });
  context_().memo.liveAssessmentReads = true;
  context_().memo.assessmentTargetNik = normalizeNik_(nik);
  var a = clone_(
    readAssessmentDetail_(
      nik,
      phase,
      t.participant.data.Program,
      t.participant.data.Batch,
    ),
  );
  legacyAssessment_(a,t.enrollment.data);
  participantFallbackAssessment_(a,t.enrollment.data);
  a.recordVersion = t.enrollment.data.RecordVersion || "";
  var sources=(parseJson_(t.enrollment.data.AssessmentSummary,{})[normalizeCurrentPhase_(phase)]||{}).sources||{};
  a.hasManualOverrides=Object.keys(sources).some(function(k){return /^MANUAL:/.test(sources[k].hash||"");});
  return a;
}
function validateUrl_(url) {
  if (url && !/^https?:\/\/[^\s]+$/i.test(url))
    throw new Error("Link harus berupa URL http atau https yang valid.");
  return String(url || "").trim();
}
function saveQleapMappings_(program, phase, items, user) {
  var name = getSetting_("QLEAP_MAPPING_SHEET") || SHEET_QLEAP,
    sh = getSheet_(name),
    h = getHeaders_(sh),
    rows = readAllAsObjects_(name);
  items.forEach(function (x) {
    var module = String(x.module || "").trim();
    if (!module) throw new Error("Nama modul wajib diisi.");
    var patch = {
        Program: program,
        Phase: phase,
        Module: module,
        QLEAPUrl: validateUrl_(x.url),
        UpdatedBy: user,
        UpdatedAt: nowIso_(),
      },
      r = uniqueMatch_(
        rows.filter(function (r) {
          return (
            String(r.data.Program).toUpperCase() === program &&
            normalizeCurrentPhase_(r.data.Phase) === phase &&
            normalizeModuleKey_(r.data.Module) === normalizeModuleKey_(module)
          );
        }),
        "QLEAP Mapping",
      );
    if (r) updateRowByIndex_(sh, h, r.rowIndex, patch);
    else appendObject_(sh, h, patch);
  });
}
function saveAssessmentScores_(token, data) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  data = data || {};
  if(normalizeCurrentPhase_(data.phase)==="PHASE 3")return savePhase3Scores_(token,data);
  var program = String(data.program || "").toUpperCase(),
    phase = normalizeCurrentPhase_(data.phase),
    cfg = getAssessmentConfig_(program, phase);
  if (!cfg) throw new Error("Phase ini tidak mendukung edit nilai.");
  var scores = {},
    allowed = cfg.modules.map(function (m) {
      return m.label;
    });
  (data.scores || []).forEach(function (x) {
    if (allowed.indexOf(x.module) < 0) throw new Error("Modul tidak dikenali.");
    scores[x.module] = validScore_(x.score);
  });
  if (!Object.keys(scores).length)
    throw new Error("Pilih nilai modul yang akan disimpan.");
  (data.qleapLinks || []).forEach(function (x) {
    if (allowed.indexOf(x.module) < 0)
      throw new Error("Modul QLEAP tidak dikenal.");
    validateUrl_(x.url);
  });
  return bulkWrites_(function () {
    ensureAutoRecapSchema_();
    var target = authorizeEnrollment_(s, data),
      p = target.participant.data,
      e = target.enrollment.data;
    if (data.recordVersion && data.recordVersion !== e.RecordVersion)
      throw new Error(
        "CONFLICT: Nilai telah diperbarui pengguna lain. Muat ulang detail sebelum menyimpan.",
      );
    var table = assessmentTable_(program, phase),
      found = findAssessmentRow_(table.sheet, cfg, p.NIK, p.Batch),
      r = found.rowNumber;
    if (!r) {
      r = Math.max(table.sheet.getLastRow() + 1, table.cfg.dataStartRow);
      var identity = {};
      table.headers.forEach(function (h) {
        var k = headerKey_(h);
        if (["nik", "employeeid"].indexOf(k) >= 0) identity[h] = p.NIK;
        else if (k === "batch") identity[h] = p.Batch;
        else if (["namakaryawan", "employeename", "name"].indexOf(k) >= 0)
          identity[h] = p.Name;
        else if (k === "region" || k === "locationgroupname")
          identity[h] = p.Region;
        else if (k === "jobtitlename") identity[h] = p.JobTitle;
        else if (k === "locationcityname") identity[h] = p.CityName || p.City;
      });
      patchColumns_(table.sheet, table.headers, r, identity);
      table.grid[r - 1] = table.headers.map(function (h) {
        return identity[h] || "";
      });
      table.byKey[entityKey_(p.NIK, program, p.Batch, p.Segment)] = [r];
    }
    var patch = {},
      changedModules = [];
    cfg.modules.forEach(function (m) {
      if (!Object.prototype.hasOwnProperty.call(scores, m.label)) return;
      var col = findHeaderIndex_(table.headers, m.headers);
      if (numberOrBlank_(table.grid[r - 1][col]) !== scores[m.label])
        changedModules.push(m.label);
      patch[table.headers[col]] = scores[m.label];
    });
    patchColumns_(table.sheet, table.headers, r, patch);
    invalidateSheet_(table.name);
    // Update only in-memory score inputs; no getValues/setValues rewrite of formula cells.
    var row = table.grid[r - 1];
    Object.keys(patch).forEach(function (h) {
      row[table.headers.indexOf(h)] = patch[h];
    });
    var passing = getPassingScore_(),
      modules = cfg.modules.map(function (m) {
        var value = numberOrBlank_(
          row[findHeaderIndex_(table.headers, m.headers)],
        );
        return {
          module: m.label,
          score: value,
          result:
            value === ""
              ? "NOT AVAILABLE"
              : value < passing
                ? "REMEDIAL REQUIRED"
                : "PASS",
          qleapUrl: qleapUrl_(program, phase, m.label),
          editable: true,
        };
      }),
      valid = modules.filter(function (m) {
        return m.score !== "";
      }),
      avg = assessmentAverage_(cfg, modules);
    var summary = parseJson_(e.AssessmentSummary, {}),
      remedials = modules.filter(function (m) {
        return m.result === "REMEDIAL REQUIRED";
      });
    var manualSources = Object.assign({}, (summary[phase] || {}).sources || {});
    changedModules.forEach(function (module) {
      manualSources[module] = {
        examAt: nowIso_(),
        hash: "MANUAL:" + Utilities.getUuid(),
        sourceLabel: "INPUT ADMIN",
        sourceRow: r,
      };
    });
    summary[phase] = {
      sources: manualSources,
      remedials: remedials,
      completedModules: valid.length,
      totalModules: modules.length,
      updatedAt: nowIso_(),
      updatedBy: s.nik,
    };
    var erPatch = {
        AssessmentSummary: JSON.stringify(summary),
        UpdatedAt: nowIso_(),
      },
      field = phase === "PHASE 1" ? "Phase1Score" : "Phase2Score";
    erPatch[field] = avg;
    erPatch.FinalScore = finalGrade_(Object.assign({},e,erPatch),p,summary).score;
    updateRowByIndex_(
      getSheet_(SHEET_ENROLLMENTS),
      getHeaders_(getSheet_(SHEET_ENROLLMENTS)),
      target.enrollment.rowIndex,
      erPatch,
    );
    var pp = { UpdatedAt: nowIso_(), FinalScore: erPatch.FinalScore };
    pp[field] = avg;
    pp[phase === "PHASE 1" ? "Phase1Status" : "Phase2Status"] = !valid.length
      ? "NOT STARTED"
      : remedials.length
        ? "REMEDIAL"
        : valid.length < modules.length
          ? "PARTIAL"
          : "PASS";
    updateRowByIndex_(
      getSheet_(SHEET_PARTICIPANTS),
      getHeaders_(getSheet_(SHEET_PARTICIPANTS)),
      target.participant.rowIndex,
      pp,
    );
    if (data.qleapLinks && data.qleapLinks.length)
      saveQleapMappings_(program, phase, data.qleapLinks, s.nik);
    var cert = certificateFor_(p);
    if (cert && !(legacySnapshot_(target.enrollment.data).version && legacySnapshot_(target.enrollment.data).locked !== false))
      updateRowByIndex_(
        getSheet_(SHEET_CERTIFICATES),
        getHeaders_(getSheet_(SHEET_CERTIFICATES)),
        cert.rowIndex,
        { FinalScore: erPatch.FinalScore },
      );
    context_().memo["assessment:"+table.name+"|"+program+"|"+phase]=table;
    target.updatedBy=s.nik; target.manualPhase=phase; refreshEnrollmentSummary_(target);
    auditLog_(
      s.nik,
      s.role,
      "SAVE_ASSESSMENT",
      "Enrollment",
      e.EnrollmentID,
      JSON.stringify({ phase: phase, modules: changedModules }),
    );
    return {
      ok: true,
      message: "Nilai tersimpan.",
      assessment: {
        nik: p.NIK,
        program: program,
        batch: p.Batch,
        phase: phase,
        available: true,
        modules: modules,
        averageScore: avg,
        finalScore: avg,
        passingScore: passing,
        completedModules: valid.length,
        totalModules: modules.length,
        recordVersion: target.enrollment.data.RecordVersion,
      },
    };
  });
}

// ---------------------------------------------------------------------------
// Participant and account mutations. Identity is explicit for every program.
// ---------------------------------------------------------------------------
function validateParticipant_(d, basicImport) {
  requireFields_(d, [
    "nik",
    "name",
    "region",
    "batch",
    "program",
    "status",
    "currentPhase",
  ]);
  if (!basicImport) requireFields_(d, ["jobTitle", "location", "cityName"]);
  d = Object.assign({}, d);
  d.segment=normalizeSegment_(d.segment || (activeSegment_()==="ALL" ? "B2B" : activeSegment_()));
  if(activeSegment_()!=="ALL" && d.segment!==activeSegment_()) throw new Error("Segmen peserta berbeda dari halaman aktif.");
  d.nik = normalizeNik_(d.nik);
  d.program = String(d.program).trim().toUpperCase();
  d.region = normalizeRegion_(d.region);
  d.status = normalizeParticipantStatus_(d.status);
  d.currentPhase = normalizeCurrentPhase_(d.currentPhase);
  if (activePrograms_(d.segment).indexOf(d.program) < 0)
    throw new Error('Program "' + d.program + '" tidak valid. Pilih program sesuai segmen.' + (/^REGION\b/i.test(d.program) ? ' Isi Region terbaca di kolom Program. Periksa kolom Region dan Program.' : ''));
  if (PARTICIPANT_STATUS_OPTIONS.indexOf(d.status) < 0)
    throw new Error('Status peserta "' + d.status + '" tidak valid.');
  if (CURRENT_PHASE_OPTIONS.indexOf(d.currentPhase) < 0)
    throw new Error('Phase "' + d.currentPhase + '" tidak valid.');
  if(d.programTrack !== undefined) d.programTrack=normalizeProgramTrack_(d.programTrack);
  if (d.programTrack && PROGRAM_TRACK_OPTIONS.indexOf(d.programTrack) < 0)
    throw new Error("Program Track tidak valid.");
  return d;
}
function ensureParticipantLogin_(p) {
  var sh = getSheet_(SHEET_USERS),
    h = getHeaders_(sh),
    rows = readAllAsObjects_(SHEET_USERS),
    r = uniqueMatch_(
      rows.filter(function (r) {
        return normalizeNik_(r.data.NIK) === normalizeNik_(p.NIK);
      }),
      "Users " + p.NIK,
    );
  if (r) {
    var roles = parseRoles_(r.data.Role);
    if (roles.indexOf("PARTICIPANT") < 0) roles.push("PARTICIPANT");
    updateRowByIndex_(sh, h, r.rowIndex, {
      Name: p.Name,
      Role: joinRoles_(roles),
      Region: roles.some(function(role){return [ROLE_ADMIN,ROLE_CLP,ROLE_SUPERIOR].indexOf(role)>=0;}) ? r.data.Region : p.Region,
      UpdatedAt: nowIso_(),
    });
  } else
    appendObject_(sh, h, {
      NIK: p.NIK,
      Name: p.Name,
      Role: "PARTICIPANT",
      Region: p.Region,
      Active: true,
      CreatedAt: nowIso_(),
      UpdatedAt: nowIso_(),
      AuthVersion: 1,
    });
}
function upsertParticipant_(d, s) {
  var rows = readAllAsObjects_(SHEET_PARTICIPANTS),
    key = entityKey_(d.nik, d.program, d.batch, d.segment),
    r = uniqueMatch_(
      rows.filter(function (r) {
        return entityKey_(r.data.NIK, r.data.Program, r.data.Batch, r.data.Segment) === key;
      }),
      "Participants " + key,
    ),
    er = enrollmentIndex_().byKey[key],
    now = nowIso_();
  if (d.rosterImport && r) {
    // A roster refresh must not reset the participant's learning history.
    d = Object.assign({}, d, {
      status: normalizeParticipantStatus_(r.data.ParticipantStatus),
      currentPhase: normalizeCurrentPhase_((er && er.data.CurrentPhase) || r.data.CurrentPhase || inferCurrentPhase_(r.data)),
      makeCurrent: er ? parseBoolean_(er.data.IsActive) : false,
    });
    ["jobTitle", "location", "cityName"].forEach(function(k) {
      if (d[k] === undefined || d[k] === "") {
        var field = {jobTitle:"JobTitle", location:"Location", cityName:"CityName"}[k];
        d[k] = r.data[field] || (k === "cityName" ? r.data.City : "") || "";
      }
    });
  }
  if (er && d.recordVersion && er.data.RecordVersion !== d.recordVersion)
    throw new Error("CONFLICT: Data peserta sudah berubah. Muat ulang daftar.");
  if (
    d.status === "PROMOSI" &&
    !(d.rosterImport && r) &&
    (!er || !isPassingStatus_(er.data.ProgramStatus))
  )
    throw new Error("Peserta harus lulus sebelum ditandai promosi.");
  if(!d.rosterImport && d.currentPhase === "PANELISASI" && (!er || (er.data.CurrentPhase !== "PANELISASI" && !workspace_(r.data).eligibility.eligible)))throw new Error("Lengkapi project dan mentoring Superior/CLP sebelum masuk panel.");
  var enrollmentId = er ? er.data.EnrollmentID : newId_("ENR"),
    patch = {
      Segment: d.segment,
      NIK: d.nik,
      Name: safeCell_(d.name),
      JobTitle: safeCell_(d.jobTitle),
      Location: safeCell_(d.location),
      CityName: safeCell_(d.cityName),
      City: safeCell_(d.cityName),
      Region: d.region,
      Program: d.program,
      Batch: String(d.batch).trim(),
      ParticipantStatus: d.status,
      CurrentPhase: d.currentPhase,
      CurrentEnrollmentId: enrollmentId,
      UpdatedAt: now,
    };
  if (d.programTrack !== undefined) patch.ProgramTrack = d.programTrack;
  if (d.superiorNik !== undefined)
    patch.Superior_NIK = normalizeNik_(d.superiorNik);
  if (d.clpNik !== undefined) patch.CLP_NIK = normalizeNik_(d.clpNik);
  var ep = {
    Segment: d.segment,
    EnrollmentID: enrollmentId,
    NIK: d.nik,
    Program: d.program,
    Batch: String(d.batch).trim(),
    Region: d.region,
    CurrentPhase: d.currentPhase,
    EnrollmentStatus: d.currentPhase === "COMPLETED" ? "COMPLETED" : "ON GOING",
    IsActive: d.makeCurrent !== false,
    UpdatedAt: now,
  };
  if (d.currentPhase === "COMPLETED")
    ep.CompletedAt = (er && er.data.CompletedAt) || now;
  if (d.makeCurrent !== false)
    readAllAsObjects_(SHEET_ENROLLMENTS)
      .filter(function (x) {
        return (
          normalizeNik_(x.data.NIK) === d.nik &&
          x.data.EnrollmentID !== enrollmentId &&
          parseBoolean_(x.data.IsActive)
        );
      })
      .forEach(function (x) {
        updateRowByIndex_(
          getSheet_(SHEET_ENROLLMENTS),
          getHeaders_(getSheet_(SHEET_ENROLLMENTS)),
          x.rowIndex,
          { IsActive: false, UpdatedAt: now },
        );
      });
  if (er)
    updateRowByIndex_(
      getSheet_(SHEET_ENROLLMENTS),
      getHeaders_(getSheet_(SHEET_ENROLLMENTS)),
      er.rowIndex,
      ep,
    );
  else {
    ep.CreatedAt = now;
    ep.StartedAt = now;
    ep.ProgramStatus = "ON GOING";
    appendObject_(
      getSheet_(SHEET_ENROLLMENTS),
      getHeaders_(getSheet_(SHEET_ENROLLMENTS)),
      ep,
    );
  }
  if (r)
    updateRowByIndex_(
      getSheet_(SHEET_PARTICIPANTS),
      getHeaders_(getSheet_(SHEET_PARTICIPANTS)),
      r.rowIndex,
      patch,
    );
  else {
    patch.JoinDate = formatDate_(new Date());
    appendObject_(
      getSheet_(SHEET_PARTICIPANTS),
      getHeaders_(getSheet_(SHEET_PARTICIPANTS)),
      patch,
    );
  }
  ensureParticipantLogin_(patch);
  return { created: !r, enrollmentId: enrollmentId };
}
function saveParticipantControl_(token, data) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var d = validateParticipant_(data);
  return bulkWrites_(function () {
    var out = upsertParticipant_(d, s);
    auditLog_(
      s.nik,
      s.role,
      "SAVE_PARTICIPANT",
      "Enrollment",
      out.enrollmentId,
      d.program + " " + d.batch,
    );
    return { ok: true, created: out.created, enrollmentId: out.enrollmentId };
  });
}
function uploadParticipants_(token, data) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  data = data || {};
  requireFields_(data, ["fileName", "base64"]);
  validateImportSize_(data);
  var table = parseTabularBase64_(data.fileName, data.mimeType, data.base64),
    head = (table[0] || []).map(headerKey_),
    aliases = {
      segment: ["segmen","segment"],
      nik: ["nik", "employeeid"],
      name: ["nama", "namalengkap", "name", "employeename"],
      region: ["region", "orggroupname"],
      program: ["program"],
      batch: ["batch"],
      jobTitle: ["jobtitle", "jabatan"],
      location: ["location", "lokasi"],
      cityName: ["cityname", "city", "locationcityname", "kota"],
    }, map = {};
  Object.keys(aliases).forEach(function(k) {
    var cols = [];
    head.forEach(function(h,i) { if (aliases[k].indexOf(h) >= 0) cols.push(i); });
    if (cols.length > 1) throw new Error("Kolom " + k + " ganda pada file import.");
    map[k] = cols.length ? cols[0] : -1;
    var required = ["nik", "name", "region", "program", "batch"].indexOf(k) >= 0;
    // Keep older clients working when they explicitly send a program for the old template.
    if (required && map[k] < 0 && !(k === "program" && data.program))
      throw new Error("Kolom " + k + " tidak tersedia. Gunakan template NIK, Nama, Job Title, Region, Program, Batch, Segmen.");
  });
  var input = [], seen = {};
  table.slice(1).forEach(function(row, i) {
    if (!row.some(function(v) { return String(v === undefined || v === null ? "" : v).trim(); })) return;
    var d = {segment:map.segment<0?"B2B":activeSegment_(),program:data.program, status:data.status || "ELIGIBLE", currentPhase:data.currentPhase || "PHASE 1", makeCurrent:data.makeCurrent !== false};
    Object.keys(map).forEach(function(k) {
      if (map[k] >= 0) d[k] = String(row[map[k]] === undefined || row[map[k]] === null ? "" : row[map[k]]).trim();
    });
    try { d = validateParticipant_(d, true); }
    catch(error) { throw new Error("Baris " + (i + 2) + ": " + error.message); }
    d.rosterImport = true;
    var key = entityKey_(d.nik, d.program, d.batch, d.segment);
    if (seen[key]) throw new Error("Input duplikat pada baris " + (i + 2) + " untuk " + key);
    seen[key] = true;
    input.push(d);
  });
  if (!input.length) throw new Error("Tidak ada data peserta.");
  if (input.length > 250) throw new Error("Maksimal 250 peserta per import. Pisahkan file untuk menjaga waktu proses.");
  return bulkWrites_(function() {
    var added = 0, updated = 0, roster = {};
    readAllAsObjects_(SHEET_PARTICIPANTS).forEach(function(r) {
      var k = entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment);
      if (seen[k] && roster[k]) throw new Error("Participants duplikat: " + k);
      roster[k] = r;
    });
    // Preflight the target identity sets before staging any participant or account writes.
    enrollmentIndex_();
    var userNiks = {};
    readAllAsObjects_(SHEET_USERS).forEach(function(r) {
      var n = normalizeNik_(r.data.NIK);
      if (userNiks[n]) throw new Error("Users duplikat " + n);
      userNiks[n] = true;
    });
    input.forEach(function(d) {
      var key = entityKey_(d.nik,d.program,d.batch, d.segment);
      if (roster[key] && !enrollmentIndex_().byKey[key])
        ensureImportEnrollments_(d.program,d.batch,d.nik,s);
      var r = upsertParticipant_(d,s);
      if (r.created) added++; else updated++;
    });
    auditLog_(s.nik,s.role,"IMPORT_PARTICIPANTS","Participants","",input.length + " record");
    return {ok:true, added:added, updated:updated, failed:0, errors:[], message:added + " peserta ditambahkan, " + updated + " diperbarui. Akun peserta baru otomatis dibuat."};
  });
}
function bulkUpdateParticipantPhase_(token, data) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var phase = normalizeCurrentPhase_(data.currentPhase);
  if (CURRENT_PHASE_OPTIONS.indexOf(phase) < 0)
    throw new Error("Phase tidak valid.");
  if (!data.enrollmentIds || !data.enrollmentIds.length)
    throw new Error("Pilih enrollment yang akan diubah.");
  return bulkWrites_(function () {
    var targets = data.enrollmentIds.map(function (id) {
      var er = enrollmentById_(id);
      if (!er) throw new Error("Enrollment tidak ditemukan.");
      return authorizeEnrollment_(s, { nik: er.data.NIK, enrollmentId: id });
    });
    if(phase === "PANELISASI")targets.forEach(function(t){var w=workspace_(t.participant.data);if(!w.eligibility.eligible)throw new Error("Syarat panel belum lengkap: "+w.eligibility.missing.join(", "));});
    targets.forEach(function (t) {
      var patch = {
        CurrentPhase: phase,
        EnrollmentStatus: phase === "COMPLETED" ? "COMPLETED" : "ON GOING",
        CompletedAt:
          phase === "COMPLETED"
            ? t.enrollment.data.CompletedAt || nowIso_()
            : "",
        UpdatedAt: nowIso_(),
      };
      updateRowByIndex_(
        getSheet_(SHEET_ENROLLMENTS),
        getHeaders_(getSheet_(SHEET_ENROLLMENTS)),
        t.enrollment.rowIndex,
        patch,
      );
      updateRowByIndex_(
        getSheet_(SHEET_PARTICIPANTS),
        getHeaders_(getSheet_(SHEET_PARTICIPANTS)),
        t.participant.rowIndex,
        { CurrentPhase: phase, UpdatedAt: patch.UpdatedAt },
      );
    });
    return {
      ok: true,
      updated: targets.length,
      message: targets.length + " enrollment diperbarui.",
    };
  });
}
function updateParticipantResult_(token,target,result) {
  var s=requireSession_(token);requireRole_(s,["ADMIN"]);
  return bulkWrites_(function(){var t=authorizeEnrollment_(s,target);refreshEnrollmentSummary_(t);auditLog_(s.nik,s.role,"RECALCULATE_RESULT","Enrollment",t.enrollment.data.EnrollmentID,"Automatic 10/15/15/30/30");return {ok:true,message:"Hasil dihitung otomatis dari seluruh komponen nilai. Koreksi nilai melalui detail phase."};});
}

function updateParticipantStatus_(token, target, status) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  status = normalizeParticipantStatus_(status);
  if (PARTICIPANT_STATUS_OPTIONS.indexOf(status) < 0)
    throw new Error("Status tidak valid.");
  return bulkWrites_(function () {
    var t = authorizeEnrollment_(s, target);
    if (
      status === "PROMOSI" &&
      !isPassingStatus_(t.enrollment.data.ProgramStatus)
    )
      throw new Error("Peserta harus lulus sebelum ditandai promosi.");
    updateRowByIndex_(
      getSheet_(SHEET_PARTICIPANTS),
      getHeaders_(getSheet_(SHEET_PARTICIPANTS)),
      t.participant.rowIndex,
      { ParticipantStatus: status, UpdatedAt: nowIso_() },
    );
    return { ok: true };
  });
}
function auditLog_(nik, role, action, entity, id, detail) {
  var s = getSheet_(SHEET_AUDIT);
  appendObject_(s, getHeaders_(s), {
    Timestamp: nowIso_(),
    NIK: nik,
    Role: role,
    Action: action,
    Entity: entity,
    EntityID: id,
    Detail: detail,
  });
}
function getAuditPage_(token, f) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  f = f || {};
  var sh = getSheet_(SHEET_AUDIT),
    h = getHeaders_(sh),
    last = sh.getLastRow(),
    before = Math.min(last, Number(f.before) || last),
    size = Math.min(100, Math.max(10, Number(f.pageSize) || 50)),
    start = Math.max(2, before - size + 1),
    rows = [];
  if (before >= 2)
    rows = sh
      .getRange(start, 1, before - start + 1, h.length)
      .getValues()
      .map(function (r) {
        return rowToObject_(h, r);
      })
      .filter(function(d){return normalizeSegment_(d.Segment)===activeSegment_();})
      .reverse();
  return {
    rows: rows,
    before: before,
    nextBefore: start > 2 ? start - 1 : null,
  };
}
function publicUser_(u) {
  return {
    nik: u.NIK,
    name: u.Name,
    role: u.Role,
    roles: parseRoles_(u.Role),
    region: u.Region,
    superiorNik: u.Superior_NIK,
    email: u.Email,
    active: parseBoolean_(u.Active),
    passwordReady: !!u.PasswordHash,
  };
}
function getUserAccountsPage_(token, f) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  f = f || {};
  var q = String(f.search || "").toLowerCase();
  var page=paginate_(
    readAllAsObjects_(SHEET_USERS)
      .filter(function (r) {
        var u=r.data;
        return !q || (u.NIK + " " + u.Name).toLowerCase().indexOf(q) >= 0;
      }),
    f,
  );
  page.rows=page.rows.map(function(r){return publicUser_(r.data);});
  return page;
}
function saveUserAccount_(token, data) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  requireFields_(data, ["nik", "name", "role"]);
  var roles = parseRolesStrict_(data.role);
  if (!roles.roles.length || roles.invalid.length)
    throw new Error("Role tidak valid.");
  var active = parseBoolean_(data.active),
    nik = normalizeNik_(data.nik);
  if (nik === s.nik && (!active || roles.roles.indexOf("ADMIN") < 0))
    throw new Error("Akun Admin sendiri harus tetap aktif sebagai Admin.");
  return bulkWrites_(function () {
    var sh = getSheet_(SHEET_USERS),
      h = getHeaders_(sh),
      r = uniqueMatch_(
        readAllAsObjects_(SHEET_USERS).filter(function (r) {
          return normalizeNik_(r.data.NIK) === nik;
        }),
        "Users",
      ),
      patch = {
        NIK: nik,
        Name: safeCell_(data.name),
        Role: joinRoles_(roles.roles),
        Region: normalizeRegion_(data.region),
        Superior_NIK: normalizeNik_(data.superiorNik),
        Email: String(data.email || "").trim(),
        Active: active,
        UpdatedAt: nowIso_(),
      };
    if (r) updateRowByIndex_(sh, h, r.rowIndex, patch);
    else {
      patch.CreatedAt = nowIso_();
      appendObject_(sh, h, patch);
    }
    auditLog_(
      s.nik,
      s.role,
      "SAVE_USER",
      "Users",
      nik,
      "Akun dan akses diperbarui",
    );
    return { ok: true, created: !r };
  });
}
function softDeleteUser_(token, nik) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  if (normalizeNik_(nik) === s.nik)
    throw new Error("Akun sendiri tidak dapat dinonaktifkan.");
  return bulkWrites_(function () {
    var r = findUserRowByNik_(nik);
    if (!r) throw new Error("User tidak ditemukan.");
    updateRowByIndex_(
      getSheet_(SHEET_USERS),
      getHeaders_(getSheet_(SHEET_USERS)),
      r.rowIndex,
      { Active: false, Role: "", UpdatedAt: nowIso_() },
    );
    return { ok: true };
  });
}
function resetAdminPassword_(token, nik, password) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  if (String(password || "").length < 8)
    throw new Error("Password minimal 8 karakter.");
  return withLock_(function () {
    var r = findUserRowByNik_(nik);
    if (!r || !userHasRole_(r.data, "ADMIN"))
      throw new Error("Akun Admin tidak ditemukan.");
    var salt = generateSalt_();
    updateRowByIndex_(
      getSheet_(SHEET_USERS),
      getHeaders_(getSheet_(SHEET_USERS)),
      r.rowIndex,
      {
        Salt: salt,
        PasswordHash: hashPassword_(password, salt),
        UpdatedAt: nowIso_(),
      },
    );
    auditLog_(
      s.nik,
      s.role,
      "RESET_ADMIN_PASSWORD",
      "Users",
      nik,
      "Password direset; sesi lama dicabut",
    );
    return { ok: true };
  });
}
function saveAdminRoleSettings_(token, data) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  return bulkWrites_(function () {
    var roles = [
        {
          role: "CLP",
          key: "CLP_ENABLED",
          value: parseBoolean_(data.clpEnabled),
        },
        {
          role: "SUPERIOR",
          key: "SUPERIOR_ENABLED",
          value: parseBoolean_(data.superiorEnabled),
        },
      ],
      changed = [];
    roles.forEach(function (r) {
      if (parseBoolean_(getSetting_(r.key)) !== r.value) changed.push(r.role);
      setSetting_(r.key, r.value ? "TRUE" : "FALSE");
    });
    if (changed.length)
      readAllAsObjects_(SHEET_USERS).forEach(function (r) {
        if (
          changed.some(function (role) {
            return userHasRole_(r.data, role);
          })
        )
          updateRowByIndex_(
            getSheet_(SHEET_USERS),
            getHeaders_(getSheet_(SHEET_USERS)),
            r.rowIndex,
            { AuthVersion: Number(r.data.AuthVersion || 1) + 1 },
          );
      });
    return { ok: true, settings: getRoleSettings_() };
  });
}

// ---------------------------------------------------------------------------
// File uploads: reserve operation, upload outside the lock, then commit metadata.
// ---------------------------------------------------------------------------
var SHEET_UPLOADS_ = "Upload Operations";
var SHEET_QPT_BATCHES_ = "QLEAP Upload Batches";
var SHEET_QPT_LATEST_ = "QLEAP Latest";
SHEET_HEADERS[SHEET_UPLOADS_] = [
  "OperationID",
  "NIK",
  "EnrollmentID",
  "Kind",
  "Phase",
  "Status",
  "LeaseUntil",
  "FileId",
  "FileUrl",
  "ResultJSON",
  "Error",
  "UpdatedAt",
];
SHEET_HEADERS[SHEET_QPT_BATCHES_] = [
  "UploadBatchID",
  "Checksum",
  "FileName",
  "Rows",
  "Participants",
  "Exams",
  "UploadedBy",
  "UploadedAt",
  "Status",
];
SHEET_HEADERS[SHEET_QPT_LATEST_] = SHEET_HEADERS[SHEET_QLEAP_POSTTEST].slice();
function validateImportSize_(d) {
  var raw = String(d.base64 || "")
      .split("base64,")
      .pop(),
    max = Number(getSetting_("MAX_FILE_MB")) || 20;
  if (!raw || !/^[A-Za-z0-9+/=\r\n]+$/.test(raw))
    throw new Error("Isi file tidak valid.");
  if ((raw.length * 3) / 4 > max * 1024 * 1024)
    throw new Error("Ukuran file melebihi " + max + " MB.");
}
function validateUpload_(d) {
  validateImportSize_(d);
  if (!ALLOWED_UPLOAD_MIME[d.mimeType])
    throw new Error("Gunakan PPT, PPTX, PDF, PNG, atau JPG.");
}
function getRootFolder_() {
  var id = getSetting_("ROOT_FOLDER_ID");
  if (!id)
    throw new Error("Folder upload belum disiapkan. Jalankan setupLdpSystem.");
  return DriveApp.getFolderById(id);
}
function getOrCreateChildFolder_(parent, name) {
  var found = parent.getFoldersByName(sanitizeFileName_(name));
  return found.hasNext()
    ? found.next()
    : parent.createFolder(sanitizeFileName_(name));
}
function uploadFile_(token, d, kind) {
  var s = requireSession_(token);
  requireRole_(s, ["PARTICIPANT"]);
  requireFields_(d, [
    "nik",
    "enrollmentId",
    "fileName",
    "mimeType",
    "base64",
    "operationId",
  ]);
  validateUpload_(d);
  if (normalizeNik_(d.nik) !== s.nik)
    throw new Error("Hanya dapat mengunggah dokumen sendiri.");
  if (!/^[A-Za-z0-9_-]{12,100}$/.test(d.operationId))
    throw new Error("Identitas upload tidak valid. Muat ulang form.");
  var t = authorizeEnrollment_(s, d),
    phase = String(d.phase || "").toUpperCase();
  if (kind === "PROJECT") {
    if (["PROPOSAL", "PROGRESS", "FINAL"].indexOf(phase) < 0)
      throw new Error("Tahap project tidak valid.");
    if (
      phaseRank_(t.enrollment.data.CurrentPhase) < 3 ||
      !hasMentoringEvidence_(workspace_(t.participant.data).mentoring)
    )
      throw new Error(
        "Project aktif setelah Phase 3 dan minimal satu bukti mentoring tersedia.",
      );
  } else if (
    ["ASH", "TSH", "LOB", "CLP"].indexOf(
      String(d.mentorRole || "").toUpperCase(),
    ) < 0
  )
    throw new Error("Pilih mentor ASH, TSH, LOB, atau CLP.");
  var reserved = bulkWrites_(function () {
    var sh = getSheet_(SHEET_UPLOADS_),
      h = getHeaders_(sh),
      r = uniqueMatch_(
        readAllAsObjects_(SHEET_UPLOADS_).filter(function (r) {
          return r.data.OperationID === d.operationId;
        }),
        "operation",
      );
    if (r) {
      if (
        r.data.NIK !== s.nik ||
        r.data.EnrollmentID !== d.enrollmentId ||
        r.data.Kind !== kind ||
        r.data.Phase !== phase
      )
        throw new Error("Identitas upload sudah digunakan.");
      if (r.data.Status === "DONE")
        return {
          done: parseJson_(r.data.ResultJSON, {
            ok: true,
            fileUrl: r.data.FileUrl,
          }),
        };
      if (
        r.data.Status === "PROCESSING" &&
        Number(r.data.LeaseUntil) > Date.now()
      )
        throw new Error("Upload ini masih diproses. Tunggu sebelum mengulang.");
      updateRowByIndex_(sh, h, r.rowIndex, {
        Status: "PROCESSING",
        LeaseUntil: Date.now() + 330000,
        Error: "",
        UpdatedAt: nowIso_(),
      });
      return { row: r.rowIndex };
    }
    return {
      row: appendObject_(sh, h, {
        OperationID: d.operationId,
        NIK: s.nik,
        EnrollmentID: d.enrollmentId,
        Kind: kind,
        Phase: phase,
        Status: "PROCESSING",
        LeaseUntil: Date.now() + 330000,
        UpdatedAt: nowIso_(),
      }),
    };
  });
  if (reserved.done) return reserved.done;
  try {
    var folder = getOrCreateChildFolder_(
        getOrCreateChildFolder_(
          getRootFolder_(),
          kind === "PROJECT" ? "Project" : "Evidence",
        ),
        d.enrollmentId,
      ),
      fileName = d.operationId + "_" + sanitizeFileName_(d.fileName),
      found = folder.getFilesByName(fileName),
      file;
    if (found.hasNext()) file = found.next();
    else
      file = folder.createFile(
        Utilities.newBlob(
          Utilities.base64Decode(String(d.base64).split("base64,").pop()),
          d.mimeType,
          fileName,
        ),
      );
    return bulkWrites_(function () {
      // Reloaded under the commit lock; a retry never inserts the same metadata twice.
      var op = uniqueMatch_(
        readAllAsObjects_(SHEET_UPLOADS_).filter(function (r) {
          return r.data.OperationID === d.operationId;
        }),
        "operation",
      );
      if (op && op.data.Status === "DONE")
        return parseJson_(op.data.ResultJSON, {
          ok: true,
          fileUrl: op.data.FileUrl,
        });
      var obj = {
          NIK: s.nik,
          EnrollmentID: d.enrollmentId,
          OperationID: d.operationId,
          FileName: sanitizeFileName_(d.fileName),
          FileUrl: file.getUrl(),
          DriveFileId: file.getId(),
        },
        result = { ok: true, fileUrl: file.getUrl() };
      if (kind === "PROJECT") {
        var existing = readAllAsObjects_(SHEET_SUBMISSIONS).filter(
          function (r) {
            return (
              r.data.EnrollmentID === d.enrollmentId && r.data.Phase === phase
            );
          },
        );
        obj.SubmissionID = newId_("SUB");
        obj.Phase = phase;
        obj.Version =
          existing.reduce(function (a, r) {
            return Math.max(a, Number(r.data.Version) || 0);
          }, 0) + 1;
        obj.SubmittedAt = nowIso_();
        obj.Status = "SUBMITTED";
        appendObject_(
          getSheet_(SHEET_SUBMISSIONS),
          getHeaders_(getSheet_(SHEET_SUBMISSIONS)),
          obj,
        );
        result.version = obj.Version;
      } else {
        obj.EvidenceID = newId_("EVD");
        obj.MentorRole = String(d.mentorRole).toUpperCase();
        obj.MentorNIK = normalizeNik_(d.mentorNik);
        obj.MentoringDate =
          toIsoDate_(d.mentoringDate) || formatDate_(new Date());
        obj.Notes = safeCell_(d.notes || "");
        obj.Feedback = obj.Notes;
        appendObject_(
          getSheet_(SHEET_MENTORING),
          getHeaders_(getSheet_(SHEET_MENTORING)),
          obj,
        );
      }
      updateRowByIndex_(
        getSheet_(SHEET_UPLOADS_),
        getHeaders_(getSheet_(SHEET_UPLOADS_)),
        op.rowIndex,
        {
          Status: "DONE",
          LeaseUntil: "",
          FileId: file.getId(),
          FileUrl: file.getUrl(),
          ResultJSON: JSON.stringify(result),
          UpdatedAt: nowIso_(),
        },
      );
      var updatedTarget=authorizeEnrollment_(s,{nik:s.nik,enrollmentId:d.enrollmentId});
      updatedTarget.updatedBy=s.nik;refreshEnrollmentSummary_(updatedTarget);
      auditLog_(
        s.nik,
        s.role,
        "UPLOAD_" + kind,
        kind,
        d.enrollmentId,
        sanitizeFileName_(d.fileName),
      );
      return result;
    });
  } catch (e) {
    bulkWrites_(function () {
      var op = readAllAsObjects_(SHEET_UPLOADS_).filter(function (r) {
        return r.data.OperationID === d.operationId;
      })[0];
      if (op && op.data.Status !== "DONE")
        updateRowByIndex_(
          getSheet_(SHEET_UPLOADS_),
          getHeaders_(getSheet_(SHEET_UPLOADS_)),
          op.rowIndex,
          {
            Status: "FAILED",
            Error: String(e.message || e),
            LeaseUntil: "",
            UpdatedAt: nowIso_(),
          },
        );
    });
    throw e;
  }
}

// ---------------------------------------------------------------------------
// QLEAP raw history and latest read model. Latest means latest examination date.
// ---------------------------------------------------------------------------
function qptKey_(d) {
  return (
    normalizeNik_(d.NIK) +
    "|" +
    String(d.ExamName || "")
      .trim()
      .toUpperCase()
  );
}
function qptIsNewer_(a, b) {
  if (!b) return true;
  var at = timeValue_(a.ExamDate),
    bt = timeValue_(b.ExamDate);
  if (at !== bt) return at > bt;
  return timeValue_(a.UploadedAt) >= timeValue_(b.UploadedAt);
}
function latestQpt_() {
  var c = context_();
  if (c.memo.latestQpt) return c.memo.latestQpt;
  var rows = readAllAsObjects_(SHEET_QPT_LATEST_);
  if (rows.length)
    return (c.memo.latestQpt = rows.map(function (r) {
      return r.data;
    }));
  var map = {};
  readAllAsObjects_(SHEET_QLEAP_POSTTEST).forEach(function (r) {
    var k = qptKey_(r.data);
    if (qptIsNewer_(r.data, map[k])) map[k] = r.data;
  });
  return (c.memo.latestQpt = Object.keys(map).map(function (k) {
    return map[k];
  }));
}
function uploadQleapPostTest_(token, d) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  validateImportSize_(d);
  var checksum = cacheKey_("", String(d.base64).split("base64,").pop()),
    records = parseQleapExport_(d.fileName, d.mimeType, d.base64);
  if (!records.length) throw new Error("Tidak ada nilai exam yang valid.");
  if (records.length > 25000)
    throw new Error("Maksimal 25.000 hasil exam per import. Pisahkan file.");
  records.forEach(function (r) {
    validScore_(r.score);
  });
  return bulkWrites_(function () {
    var duplicate = readAllAsObjects_(SHEET_QPT_BATCHES_).filter(function (r) {
      return r.data.Checksum === checksum && r.data.Status === "COMMITTED";
    })[0];
    if (duplicate)
      return {
        ok: true,
        duplicate: true,
        batchId: duplicate.data.UploadBatchID,
        inserted: 0,
        message: "File ini sudah diimport. Tidak ada duplikasi data.",
      };
    var batchId = newId_("QPT"),
      now = nowIso_(),
      niks = {},
      exams = {},
      old = readAllAsObjects_(SHEET_QPT_LATEST_),
      byKey = {};
    old.forEach(function (r) {
      byKey[qptKey_(r.data)] = r;
    });
    // Upgrade any legacy raw rows before combining a first new import.
    if (!old.length)
      latestQpt_().forEach(function (o) {
        var ri = appendObject_(
          getSheet_(SHEET_QPT_LATEST_),
          getHeaders_(getSheet_(SHEET_QPT_LATEST_)),
          o,
        );
        byKey[qptKey_(o)] = { rowIndex: ri, data: o };
      });
    records.forEach(function (r) {
      var obj = {
        UploadBatchID: batchId,
        NIK: r.nik,
        EmployeeName: safeCell_(r.name),
        JobTitle: safeCell_(r.jobTitle),
        WorkLocation: safeCell_(r.workLocation),
        BusinessUnit: safeCell_(r.businessUnit),
        ElearningStatus: r.elStat,
        Modules: r.modules,
        ExamName: safeCell_(r.examName),
        Score: r.score,
        Duration: r.duration,
        ExamDate: r.examDate,
        Status: r.status,
        UploadedBy: s.nik,
        UploadedAt: now,
      };
      appendObject_(
        getSheet_(SHEET_QLEAP_POSTTEST),
        getHeaders_(getSheet_(SHEET_QLEAP_POSTTEST)),
        obj,
      );
      var k = qptKey_(obj),
        prev = byKey[k];
      if (qptIsNewer_(obj, prev && prev.data)) {
        if (prev)
          updateRowByIndex_(
            getSheet_(SHEET_QPT_LATEST_),
            getHeaders_(getSheet_(SHEET_QPT_LATEST_)),
            prev.rowIndex,
            obj,
          );
        else {
          var row = appendObject_(
            getSheet_(SHEET_QPT_LATEST_),
            getHeaders_(getSheet_(SHEET_QPT_LATEST_)),
            obj,
          );
          byKey[k] = { rowIndex: row, data: obj };
        }
      }
      niks[r.nik] = true;
      exams[r.examName] = true;
    });
    appendObject_(
      getSheet_(SHEET_QPT_BATCHES_),
      getHeaders_(getSheet_(SHEET_QPT_BATCHES_)),
      {
        UploadBatchID: batchId,
        Checksum: checksum,
        FileName: sanitizeFileName_(d.fileName),
        Rows: records.length,
        Participants: Object.keys(niks).length,
        Exams: Object.keys(exams).length,
        UploadedBy: s.nik,
        UploadedAt: now,
        Status: "COMMITTED",
      },
    );
    return {
      ok: true,
      batchId: batchId,
      inserted: records.length,
      message: records.length + " hasil exam berhasil diimport.",
    };
  });
}
function qptRecap_(s, f) {
  f = f || {};
  var cacheKey=cacheKey_('ldpqptpage:',JSON.stringify([LDP_VERSION_,activeSegment_(),revision_(),s.role,s.nik,s.region || '',f.page || 1,f.pageSize || 30,f.search || '',!!f.onlyRemedial]));
  if(f.fresh) context_().memo.bypassSnapshots=true;
  else {var cached=readSnapshot_(cacheKey);if(cached) return cached;}
  var scope = {};
  scopedParticipants_(s).forEach(function (r) {
    scope[normalizeNik_(r.data.NIK)] = true;
  });
  var byNik = {},
    passing = getPassingScore_();
  latestQpt_().forEach(function (d) {
    var nik = normalizeNik_(d.NIK);
    if (!scope[nik] || (s.role === "PARTICIPANT" && nik !== s.nik)) return;
    var p =
        byNik[nik] ||
        (byNik[nik] = { nik: nik, name: d.EmployeeName, exams: [] }),
      score = numberOrBlank_(d.Score),
      result =
        score === ""
          ? "NOT AVAILABLE"
          : score < passing
            ? "REMEDIAL REQUIRED"
            : "PASS";
    p.exams.push({
      examName: d.ExamName,
      score: score,
      result: result,
      examDate: d.ExamDate,
      duration: d.Duration,
      qleapUrl:
        result === "REMEDIAL REQUIRED"
          ? qleapUrl_("POST TEST", "POST TEST", d.ExamName)
          : "",
    });
  });
  var q = String(f.search || "").toLowerCase(),
    rows = Object.keys(byNik)
      .map(function (n) {
        var p = byNik[n],
          valid = p.exams.filter(function (e) {
            return e.score !== "";
          });
        p.remedialCount = p.exams.filter(function (e) {
          return e.result === "REMEDIAL REQUIRED";
        }).length;
        p.averageScore = valid.length
          ? Math.round(
              (valid.reduce(function (a, e) {
                return a + e.score;
              }, 0) /
                valid.length) *
                100,
            ) / 100
          : "";
        return p;
      })
      .filter(function (p) {
        return (
          (!q || (p.nik + " " + p.name).toLowerCase().indexOf(q) >= 0) &&
          (!f.onlyRemedial || p.remedialCount > 0)
        );
      })
      .sort(function (a, b) {
        return String(a.name).localeCompare(String(b.name));
      });
  var out = paginate_(rows, f);
  out.passingScore = passing;
  out.summary = {
    totalParticipants: rows.length,
    totalExams: rows.reduce(function (a, p) {
      return a + p.exams.length;
    }, 0),
    totalRemedialExams: rows.reduce(function (a, p) {
      return a + p.remedialCount;
    }, 0),
  };
  putSnapshot_(cacheKey,out,30);
  return out;
}
function getQleapPostTestRecap_(token, f) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN", "CLP"]);
  return qptRecap_(s, f);
}
function getMyQleapPostTest_(token, options) {
  var s = requireSession_(token);
  requireRole_(s, ["PARTICIPANT"]);
  return qptRecap_(s, options || {});
}
function getQleapMappings_(token) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  return readAllAsObjects_(
    getSetting_("QLEAP_MAPPING_SHEET") || SHEET_QLEAP,
  ).map(function (r) {
    return {
      program: r.data.Program,
      phase: r.data.Phase,
      module: r.data.Module,
      url: r.data.QLEAPUrl || r.data["QLEAP URL"] || "",
    };
  });
}
function saveQleapMapping_(token, d) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var program = String(d.program || "").toUpperCase(),
    phase = normalizeCurrentPhase_(d.phase);
  if (
    activePrograms_().concat(["POST TEST"]).indexOf(program) < 0 ||
    ["PHASE 1", "PHASE 2", "PHASE 3", "POST TEST"].indexOf(phase) < 0
  )
    throw new Error("Program/phase mapping tidak valid.");
  validateUrl_(d.url);
  return bulkWrites_(function () {
    saveQleapMappings_(
      program,
      phase,
      [{ module: d.module, url: d.url }],
      s.nik,
    );
    return { ok: true };
  });
}

// ---------------------------------------------------------------------------
// Certificates: one queue and one worker for every source of generation.
// ---------------------------------------------------------------------------
function certificateIssueDate_(value) {
  // Sheet Date cells become ISO timestamps in rowToObject_; restore the Jakarta
  // calendar date before building the template date. Never coerce null/0 to epoch.
  if (value instanceof Date)
    return isFinite(value.getTime()) ? Utilities.formatDate(value, TZ, "yyyy-MM-dd") : "";
  if (typeof value === "number") {
    // Explicit Excel/Sheets serial dates, not Unix milliseconds or zero.
    if (!isFinite(value) || value <= 20000 || value >= 80000) return "";
    return Utilities.formatDate(new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000), TZ, "yyyy-MM-dd");
  }
  if (typeof value !== "string") return "";
  var text = value.trim(), dmy = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (dmy) text = dmy[3] + "-" + ("0" + dmy[2]).slice(-2) + "-" + ("0" + dmy[1]).slice(-2);
  var match = text.match(/^(\d{4}-\d{2}-\d{2})(T(?:[01]\d|2[0-3]):[0-5]\d:[0-5]\d(?:\.\d{1,3})?(?:Z|[+-]\d{2}:\d{2}))?$/);
  if (!match) return "";
  var calendar = new Date(match[1] + "T12:00:00Z");
  if (!isFinite(calendar.getTime()) || calendar.toISOString().slice(0, 10) !== match[1]) return "";
  if (!match[2]) return match[1];
  var instant = new Date(text);
  return isFinite(instant.getTime()) ? Utilities.formatDate(instant, TZ, "yyyy-MM-dd") : "";
}
function resolveCertificateIssueDate_(requested, existing) {
  function supplied(v) { return v !== undefined && v !== null && String(v).trim() !== ""; }
  var value = supplied(requested) ? requested : supplied(existing) ? existing : new Date(),
    iso = certificateIssueDate_(value);
  if (!iso) throw new Error("Tanggal terbit sertifikat tidak valid. Pilih kembali tanggal terbit sebelum membuat sertifikat.");
  return iso;
}
function nextCertificateNumber_(program) {
  var year = Utilities.formatDate(new Date(), TZ, "yyyy"),
    prefix =
      (getSetting_("CERTIFICATE_PREFIX") || "LDP") +
      "/" +
      (activeSegment_()==="B2C" ? "B2C/" : "") + program +
      "/" +
      year +
      "/",
    key = "CERTSEQ_" + (activeSegment_()==="B2C" ? "B2C_" : "") + program + "_" + year,
    max = Number(props_().getProperty(key) || 0);
  readAllAsObjects_(SHEET_CERTIFICATES).forEach(function (r) {
    var n = String(r.data.CertificateNumber || "");
    if (n.indexOf(prefix) === 0)
      max = Math.max(max, parseInt(n.slice(prefix.length), 10) || 0);
  });
  max++;
  props_().setProperty(key, String(max));
  return prefix + ("000000" + max).slice(-6);
}
// Historical graduates need their recorded pass, not reconstructed old uploads.
// This policy is confined to certificates; workspace/panel gates stay unchanged.
function historicalCertificatePass_(p, e) {
  if(!p || !e || !e.EnrollmentID ||
    entityKey_(p.NIK,p.Program,p.Batch, p.Segment)!==entityKey_(e.NIK,e.Program,e.Batch, e.Segment)) return false;
  function recordedPass(value) {
    return ["LULUS","PASS","PASSED"].indexOf(String(value || "").trim().toUpperCase())>=0;
  }
  // Completion percentages and COMPLETED/COMPLETE alone are not a historical pass.
  if(!recordedPass(e.ProgramStatus)) return false;
  var h=legacySnapshot_(e), frozen=h.values || {}, fallback=h.fallbackValues || {};
  if(h.version && (h.migratedAt || h.source) &&
    (recordedPass(frozen.ProgramStatus || frozen.FinalStatus) || recordedPass(fallback.FinalStatus))) return true;
  // Older databases initialized before migration snapshots still carry cohort years.
  var years=normalizeBatchKey_(e.Batch).match(/\b(?:19|20)\d{2}\b/g) || [];
  return years.length===1 && Number(years[0])<Number(String(formatDate_(new Date())).slice(0,4));
}
function certificateEligibility_(t, workspace) {
  var p=t.participant.data,e=t.enrollment.data;
  if(!isPassingStatus_(e.ProgramStatus)) return {eligible:false,historical:false};
  if(historicalCertificatePass_(p,e)) return {eligible:true,historical:true};
  return {eligible:!!(workspace || workspace_(p)).eligibility.eligible,historical:false};
}
function enqueueCertificate_(t, s, options) {
  options = options || {};
  var p = t.participant.data,
    e = t.enrollment.data,
    id = e.EnrollmentID;
  if (!isPassingStatus_(e.ProgramStatus))
    throw new Error(
      "Sertifikat hanya dapat dibuat untuk enrollment yang lulus.",
    );
  if(!certificateEligibility_(t).eligible)throw new Error("Lengkapi Proposal, Progress, Final, mentoring Superior dan CLP sebelum menerbitkan sertifikat.");
  var existing = certificateFor_(p),
    cert = existing && existing.data,
    issue = resolveCertificateIssueDate_(options.issueDate, cert && cert.IssueDate),
    name = String(options.displayName || (cert && cert.DisplayName) || p.Name).trim();
  if (cert && cert.FileOrLink && !options.regenerate)
    return { jobId: cert.JobID || "", ready: true };
  var pending = readAllAsObjects_(SHEET_JOBS_).filter(function (r) {
    return (
      r.data.EnrollmentID === id &&
      ["QUEUED", "PROCESSING"].indexOf(r.data.Status) >= 0
    );
  })[0];
  if (pending) {
    var old = pending.data,
      same = certificateIssueDate_(old.IssueDate) === issue &&
        String(old.DisplayName || "").trim() === name &&
        (!options.certificateNumber || String(options.certificateNumber).trim() === String(old.CertificateNumber || "").trim());
    if (same) return { jobId: old.JobID, queued: true };
  }
  var template = String(getSetting_("CERTIFICATE_TEMPLATE_ID") || "").trim();
  if (!template)
    throw new Error(
      "Template Google Slides sertifikat belum diatur di Settings.",
    );
  var number =
    String(
      options.certificateNumber || (cert && cert.CertificateNumber) || "",
    ).trim() || nextCertificateNumber_(p.Program);
  var duplicate = allSegmentRows_(SHEET_CERTIFICATES).some(function (r) {
    return (
      r.data.EnrollmentID !== id &&
      String(r.data.CertificateNumber).toUpperCase() === number.toUpperCase()
    );
  });
  if (duplicate)
    throw new Error("Nomor sertifikat sudah dipakai enrollment lain.");
  // Changed queued/in-flight metadata gets a new immutable job. A stale worker
  // cannot publish its old PDF because its old job is cancelled before commit.
  if (pending) cancelCertificateJobs_(id);
  var jobId = newId_("CJOB"), now = nowIso_(),
    importId = options.importId || (pending && pending.data.ImportID) || "";
  appendObject_(getSheet_(SHEET_JOBS_), getHeaders_(getSheet_(SHEET_JOBS_)), {
    JobID: jobId,
    EnrollmentID: id,
    NIK: p.NIK,
    Program: p.Program,
    Batch: p.Batch,
    DisplayName: name,
    CertificateNumber: number,
    IssueDate: issue,
    TemplateID: template,
    Status: "QUEUED",
    Attempts: 0,
    ImportID: importId,
    CreatedBy: s.nik,
    CreatedAt: now,
    UpdatedAt: now,
  });
  if (pending && importId) {
    var importRow = readAllAsObjects_(SHEET_CERT_IMPORT).filter(function (r) { return r.data.ImportID === importId; })[0];
    if (importRow) updateRowByIndex_(getSheet_(SHEET_CERT_IMPORT), getHeaders_(getSheet_(SHEET_CERT_IMPORT)), importRow.rowIndex,
      { JobID: jobId, RowStatus: "QUEUED", DisplayName: name, CertificateNumber: number, IssueDate: issue, Message: "" });
  }
  var patch = {
    EnrollmentID: id,
    NIK: p.NIK,
    Program: p.Program,
    Batch: p.Batch,
    FinalScore: e.FinalScore,
    FinalStatus: e.ProgramStatus,
    CertificateNumber: number,
    DisplayName: name,
    IssueDate: issue,
    JobID: jobId,
    FileOrLink: "",
    GeneratedAt: "",
    Published: false,
    UploadedAt: now,
    UploadedBy: s.nik,
  };
  if (existing)
    updateRowByIndex_(
      getSheet_(SHEET_CERTIFICATES),
      getHeaders_(getSheet_(SHEET_CERTIFICATES)),
      existing.rowIndex,
      patch,
    );
  else {
    patch.CertificateID = newId_("CERT");
    appendObject_(
      getSheet_(SHEET_CERTIFICATES),
      getHeaders_(getSheet_(SHEET_CERTIFICATES)),
      patch,
    );
  }
  props_().setProperty("LDP_CERT_PENDING", "1");
  return { jobId: jobId, queued: true };
}
function generateCertificateFile_(job) {
  // This function deliberately holds no ScriptLock during Drive/Slides work.
  var issue = certificateIssueDate_(job.IssueDate);
  if (!issue) throw new Error("Tanggal terbit sertifikat tidak valid. Pilih kembali tanggal terbit dan buat ulang sertifikat.");
  var issuedAt = new Date(issue + "T12:00:00+07:00");
  if (!isFinite(issuedAt.getTime())) throw new Error("Tanggal terbit sertifikat tidak valid.");
  var folder = getOrCreateChildFolder_(
      getOrCreateChildFolder_(getRootFolder_(), "Certificates"),
      job.EnrollmentID,
    ),
    fileName = "Certificate_" + job.JobID + ".pdf",
    existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) {
    var done = existing.next();
    return { url: done.getUrl(), fileId: done.getId() };
  }
  var template = DriveApp.getFileById(job.TemplateID);
  if (template.getMimeType() !== "application/vnd.google-apps.presentation")
    throw new Error("Template harus berupa Google Slides asli.");
  var slideName = "Certificate_" + job.JobID,
    slides = folder.getFilesByName(slideName),
    copy = slides.hasNext()
      ? slides.next()
      : template.makeCopy(slideName, folder),
    deck = SlidesApp.openById(copy.getId()),
    label = getSetting_("PROGRAM_LABEL_" + job.Program) || job.Program;
  var replacements = {
    "{{SEGMENT}}": normalizeSegment_(job.Segment),
    "{{NAME}}": job.DisplayName,
    "{{PROGRAM_NAME}}": String(label).toUpperCase(),
    "{{PROGRAM_CODE}}": job.Program,
    "{{CERT_NO}}": job.CertificateNumber,
    "{{DATE_ISSUED}}": Utilities.formatDate(
      issuedAt,
      TZ,
      "MMMM d, yyyy",
    ).toUpperCase(),
    "{{MONTH_YEAR}}": Utilities.formatDate(
      issuedAt,
      TZ,
      "MMMM yyyy",
    ).toUpperCase(),
    "{{SIGNER_NAME}}": getSetting_("CERTIFICATE_SIGNER_NAME") || "",
    "{{SIGNER_TITLE}}": getSetting_("CERTIFICATE_SIGNER_TITLE") || "",
  };
  deck.getSlides().forEach(function (slide) {
    Object.keys(replacements).forEach(function (tag) {
      slide.replaceAllText(tag, String(replacements[tag] || ""));
    });
  });
  deck.saveAndClose();
  var blob = DriveApp.getFileById(copy.getId())
      .getAs("application/pdf")
      .setName(fileName),
    pdf = folder.createFile(blob);
  return { url: pdf.getUrl(), fileId: pdf.getId() };
}
function processCertificateJobs_() {
  return api_("certificateWorker", function () {
    if (props_().getProperty("LDP_CERT_PENDING") !== "1")
      return { ok: true, processed: 0, failed: 0, idle: true };
    var started = Date.now(),
      processed = 0,
      failed = 0;
    while (processed < 3 && Date.now() - started < 150000) {
      var claim = bulkWrites_(function () {
        var allJobs = readAllAsObjects_(SHEET_JOBS_);
        allJobs.forEach(function (r) {
          if (
            r.data.Status === "PROCESSING" &&
            Number(r.data.LeaseUntil) < Date.now() &&
            Number(r.data.Attempts) >= 3
          )
            updateRowByIndex_(
              getSheet_(SHEET_JOBS_),
              getHeaders_(getSheet_(SHEET_JOBS_)),
              r.rowIndex,
              {
                Status: "FAILED",
                LeaseToken: "",
                LeaseUntil: "",
                Error: "Batas percobaan tercapai setelah proses terputus.",
                UpdatedAt: nowIso_(),
              },
            );
        });
        var r = allJobs.filter(function (r) {
          var d = r.data;
          return (
            (d.Status === "QUEUED" ||
              (d.Status === "PROCESSING" &&
                Number(d.LeaseUntil) < Date.now())) &&
            Number(d.Attempts || 0) < 3
          );
        })[0];
        if (!r) return null;
        var token = Utilities.getUuid(),
          patch = {
            Status: "PROCESSING",
            LeaseUntil: Date.now() + 330000,
            LeaseToken: token,
            Attempts: Number(r.data.Attempts || 0) + 1,
            UpdatedAt: nowIso_(),
          };
        updateRowByIndex_(
          getSheet_(SHEET_JOBS_),
          getHeaders_(getSheet_(SHEET_JOBS_)),
          r.rowIndex,
          patch,
        );
        return clone_(Object.assign({}, r.data, patch));
      });
      if (!claim) break;
      processed++;
      try {
        var file = generateCertificateFile_(claim);
        bulkWrites_(function () {
          var jr = readAllAsObjects_(SHEET_JOBS_).filter(function (r) {
            return r.data.JobID === claim.JobID;
          })[0];
          if (
            !jr ||
            jr.data.Status !== "PROCESSING" ||
            jr.data.LeaseToken !== claim.LeaseToken
          )
            return;
          var er = enrollmentById_(claim.EnrollmentID);
          if (!er || !isPassingStatus_(er.data.ProgramStatus)) {
            updateRowByIndex_(
              getSheet_(SHEET_JOBS_),
              getHeaders_(getSheet_(SHEET_JOBS_)),
              jr.rowIndex,
              {
                Status: "CANCELLED",
                LeaseUntil: "",
                Error: "Status enrollment bukan lulus",
                UpdatedAt: nowIso_(),
              },
            );
            return;
          }
          var pr = uniqueMatch_(
            readAllAsObjects_(SHEET_PARTICIPANTS).filter(function (r) {
              return entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment)===entityKey_(er.data.NIK,er.data.Program,er.data.Batch, er.data.Segment);
            }), "participant");
          if(!pr || !certificateEligibility_({participant:pr,enrollment:er}).eligible) {
            updateRowByIndex_(getSheet_(SHEET_JOBS_),getHeaders_(getSheet_(SHEET_JOBS_)),jr.rowIndex,{
              Status:"CANCELLED",LeaseUntil:"",Error:"Syarat penerbitan sertifikat belum terpenuhi.",UpdatedAt:nowIso_()
            });
            return;
          }
          var cert = uniqueMatch_(
            readAllAsObjects_(SHEET_CERTIFICATES).filter(function (r) {
              return r.data.EnrollmentID === claim.EnrollmentID;
            }),
            "certificate",
          );
          if (!cert) throw new Error("Record sertifikat tidak ditemukan.");
          updateRowByIndex_(
            getSheet_(SHEET_CERTIFICATES),
            getHeaders_(getSheet_(SHEET_CERTIFICATES)),
            cert.rowIndex,
            { FileOrLink: file.url, GeneratedAt: nowIso_(), Published: true },
          );
          updateRowByIndex_(
            getSheet_(SHEET_ENROLLMENTS),
            getHeaders_(getSheet_(SHEET_ENROLLMENTS)),
            er.rowIndex,
            { CertificateUrl: file.url, UpdatedAt: nowIso_() },
          );
          if (pr)
            updateRowByIndex_(
              getSheet_(SHEET_PARTICIPANTS),
              getHeaders_(getSheet_(SHEET_PARTICIPANTS)),
              pr.rowIndex,
              { CertificateUrl: file.url },
            );
          updateRowByIndex_(
            getSheet_(SHEET_JOBS_),
            getHeaders_(getSheet_(SHEET_JOBS_)),
            jr.rowIndex,
            {
              Status: "DONE",
              LeaseUntil: "",
              FileUrl: file.url,
              Error: "",
              UpdatedAt: nowIso_(),
            },
          );
          if (claim.ImportID) {
            var imp = readAllAsObjects_(SHEET_CERT_IMPORT).filter(function (r) {
              return r.data.ImportID === claim.ImportID;
            })[0];
            if (imp)
              updateRowByIndex_(
                getSheet_(SHEET_CERT_IMPORT),
                getHeaders_(getSheet_(SHEET_CERT_IMPORT)),
                imp.rowIndex,
                {
                  RowStatus: "GENERATED",
                  FileOrLink: file.url,
                  GeneratedAt: nowIso_(),
                  Message: "",
                },
              );
          }
        });
      } catch (err) {
        failed++;
        bulkWrites_(function () {
          var r = readAllAsObjects_(SHEET_JOBS_).filter(function (r) {
            return r.data.JobID === claim.JobID;
          })[0];
          if (r && r.data.LeaseToken === claim.LeaseToken)
            updateRowByIndex_(
              getSheet_(SHEET_JOBS_),
              getHeaders_(getSheet_(SHEET_JOBS_)),
              r.rowIndex,
              {
                Status: Number(r.data.Attempts) >= 3 ? "FAILED" : "QUEUED",
                LeaseUntil: "",
                Error: String(err.message || err),
                UpdatedAt: nowIso_(),
              },
            );
        });
        console.error("Certificate job " + claim.JobID + ": " + err.message);
      }
    }
    bulkWrites_(function () {
      var pending = readAllAsObjects_(SHEET_JOBS_).some(function (r) {
        return ["QUEUED", "PROCESSING"].indexOf(r.data.Status) >= 0;
      });
      if (pending) props_().setProperty("LDP_CERT_PENDING", "1");
      else props_().deleteProperty("LDP_CERT_PENDING");
    });
    return { ok: true, processed: processed, failed: failed };
  }, arguments);
}
function getCertificatesOverview_(token, f) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  f = f || {};
  var cacheKey=cacheKey_('ldpcertpage:',JSON.stringify([LDP_VERSION_,activeSegment_(),revision_(),s.role,s.nik,f.page || 1,f.pageSize || 30,f.search || '']));
  if(f.fresh) context_().memo.bypassSnapshots=true;
  else {var cached=readSnapshot_(cacheKey);if(cached) return cached;}
  var q = String(f.search || "").toLowerCase(),
    jobs = {};
  readAllAsObjects_(SHEET_JOBS_).forEach(function (r) {
    jobs[r.data.JobID] = r.data;
  });
  var page=paginate_(scopedParticipants_(s).filter(function(r) {
    return !q || (r.data.NIK + ' ' + r.data.Name).toLowerCase().indexOf(q)>=0;
  }),f);
  page.rows=page.rows.map(function (r) {
      var p = r.data,
        er = enrollmentFor_(p),
        cert = certificateFor_(p),
        c = (cert && cert.data) || {},
        job = jobs[c.JobID] || {};
      return {
        nik: p.NIK,
        name: p.Name,
        program: p.Program,
        batch: p.Batch,
        enrollmentId: er && er.data.EnrollmentID,
        finalScore: er && er.data.FinalScore,
        finalStatus: er && er.data.ProgramStatus,
        passing: er && isPassingStatus_(er.data.ProgramStatus),
        certificateNumber: c.CertificateNumber || "",
        displayName: c.DisplayName || p.Name,
        fileUrl: c.FileOrLink || "",
        active: parseBoolean_(c.Published),
        issueDate: certificateIssueDate_(c.IssueDate),
        issueDateInvalid: c.IssueDate !== undefined && c.IssueDate !== null && String(c.IssueDate).trim() !== "" && !certificateIssueDate_(c.IssueDate),
        jobStatus: job.Status || "",
        jobError: job.Error || "",
      };
    });
  putSnapshot_(cacheKey,page,30);
  return page;
}
function cancelCertificateJobs_(enrollmentId) {
  readAllAsObjects_(SHEET_JOBS_)
    .filter(function (r) {
      return (
        r.data.EnrollmentID === enrollmentId &&
        ["QUEUED", "PROCESSING"].indexOf(r.data.Status) >= 0
      );
    })
    .forEach(function (r) {
      updateRowByIndex_(
        getSheet_(SHEET_JOBS_),
        getHeaders_(getSheet_(SHEET_JOBS_)),
        r.rowIndex,
        { Status: "CANCELLED", LeaseToken: "", UpdatedAt: nowIso_() },
      );
    });
}
function manageCertificate_(token, d) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  validateUrl_(d.fileOrLink);
  return bulkWrites_(function () {
    var t = authorizeEnrollment_(s, d);
    if (!isPassingStatus_(t.enrollment.data.ProgramStatus))
      throw new Error("Enrollment belum lulus.");
    var current = certificateFor_(t.participant.data), old = current && current.data,
      issue = resolveCertificateIssueDate_(d.issueDate, old && old.IssueDate),
      metadataChanged = old && (issue !== certificateIssueDate_(old.IssueDate) ||
        (d.displayName && String(d.displayName).trim() !== String(old.DisplayName || t.participant.data.Name).trim()) ||
        (d.certificateNumber && String(d.certificateNumber).trim() !== String(old.CertificateNumber || "").trim())),
      generatedFile = old && old.JobID && old.FileOrLink && d.fileOrLink === old.FileOrLink && metadataChanged &&
        readAllAsObjects_(SHEET_JOBS_).some(function (r) {
          return r.data.JobID === old.JobID && r.data.Status === "DONE" && r.data.FileUrl === old.FileOrLink;
        });
    if (generatedFile) d = Object.assign({}, d, { fileOrLink: "", regenerate: true });
    else if (d.regenerate) d = Object.assign({}, d, { fileOrLink: "" });
    if (d.fileOrLink) {
      var c = certificateFor_(t.participant.data),
        patch = {
          EnrollmentID: t.enrollment.data.EnrollmentID,
          NIK: t.participant.data.NIK,
          Program: t.participant.data.Program,
          Batch: t.participant.data.Batch,
          FinalScore: t.enrollment.data.FinalScore,
          FinalStatus: t.enrollment.data.ProgramStatus,
          FileOrLink: d.fileOrLink,
          Published: !!(c && c.data.FileOrLink === d.fileOrLink && parseBoolean_(c.data.Published)),
          UploadedAt: nowIso_(),
          UploadedBy: s.nik,
        };
      var number = String(
        d.certificateNumber || (c && c.data.CertificateNumber) || "",
      ).trim();
      if (
        number &&
        allSegmentRows_(SHEET_CERTIFICATES).some(function (r) {
          return (
            r.data.EnrollmentID !== t.enrollment.data.EnrollmentID &&
            String(r.data.CertificateNumber || "").toUpperCase() ===
              number.toUpperCase()
          );
        })
      )
        throw new Error("Nomor sertifikat sudah dipakai enrollment lain.");
      patch.CertificateNumber =
        number || nextCertificateNumber_(t.participant.data.Program);
      patch.DisplayName = safeCell_(
        String(
          d.displayName || (c && c.data.DisplayName) || t.participant.data.Name,
        ).trim(),
      );
      patch.IssueDate = issue;
      cancelCertificateJobs_(t.enrollment.data.EnrollmentID);
      if (c)
        updateRowByIndex_(
          getSheet_(SHEET_CERTIFICATES),
          getHeaders_(getSheet_(SHEET_CERTIFICATES)),
          c.rowIndex,
          patch,
        );
      else {
        patch.CertificateID = newId_("CERT");
        appendObject_(
          getSheet_(SHEET_CERTIFICATES),
          getHeaders_(getSheet_(SHEET_CERTIFICATES)),
          patch,
        );
      }
      return {
        ok: true,
        message: "Link sertifikat tersimpan. Periksa preview lalu aktifkan.",
      };
    }
    var job = enqueueCertificate_(t, s, d);
    return {
      ok: true,
      queued: job.queued,
      jobId: job.jobId,
      message: job.ready
        ? "Sertifikat sudah tersedia."
        : "Sertifikat sedang dibuat dari template dan akan otomatis aktif setelah selesai.",
    };
  });
}
function toggleCertificateActive_(token, target, active) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  return bulkWrites_(function () {
    var t = authorizeEnrollment_(s, target),
      c = certificateFor_(t.participant.data);
    if (!c || !c.data.FileOrLink)
      throw new Error("File sertifikat belum tersedia.");
    if(active && !certificateEligibility_(t).eligible)throw new Error("Sertifikat memerlukan status lulus dan kelengkapan project/mentoring, kecuali peserta historis yang sudah tercatat lulus.");
    if (active && !isPassingStatus_(t.enrollment.data.ProgramStatus))
      throw new Error("Enrollment belum lulus.");
    updateRowByIndex_(
      getSheet_(SHEET_CERTIFICATES),
      getHeaders_(getSheet_(SHEET_CERTIFICATES)),
      c.rowIndex,
      { Published: parseBoolean_(active) },
    );
    return { ok: true };
  });
}
function retryCertificateJob_(token, jobId) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  return bulkWrites_(function () {
    var r = readAllAsObjects_(SHEET_JOBS_).filter(function (r) {
      return r.data.JobID === jobId;
    })[0];
    if (!r || r.data.Status !== "FAILED")
      throw new Error("Hanya job gagal yang dapat diulang.");
    updateRowByIndex_(
      getSheet_(SHEET_JOBS_),
      getHeaders_(getSheet_(SHEET_JOBS_)),
      r.rowIndex,
      {
        Status: "QUEUED",
        Attempts: 0,
        LeaseUntil: "",
        Error: "",
        UpdatedAt: nowIso_(),
      },
    );
    props_().setProperty("LDP_CERT_PENDING", "1");
    return { ok: true };
  });
}
function getCertificateImportTemplate_(token) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  return {
    filename: "Template_Sertifikat_LDP.csv",
    content:
      "\uFEFFNIK,Program,Batch,Nama Sertifikat,Final Status,Nilai Final,No Sertifikat,Tanggal Terbit\r\n",
  };
}
function uploadCertificateImport_(token, d) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  validateImportSize_(d);
  var rows = parseTabularBase64_(d.fileName, d.mimeType, d.base64),
    h = (rows[0] || []).map(headerKey_),
    ix = {
      nik: findHeaderIndex_(h, ["NIK"]),
      program: findHeaderIndex_(h, ["Program"]),
      batch: findHeaderIndex_(h, ["Batch"]),
      name: findHeaderIndex_(h, ["Nama Sertifikat", "Nama", "Name"]),
      status: findHeaderIndex_(h, ["Final Status", "FinalStatus"]),
      score: findHeaderIndex_(h, ["Nilai Final", "FinalScore"]),
      number: findHeaderIndex_(h, ["No Sertifikat", "CertificateNumber"]),
      issue: findHeaderIndex_(h, ["Tanggal Terbit", "IssueDate"]),
    };
  if (ix.nik < 0 || ix.program < 0 || ix.batch < 0)
    throw new Error("Template wajib memuat NIK, Program, dan Batch.");
  if (rows.length > 251)
    throw new Error("Maksimal 250 record per import sertifikat.");
  var input = [],
    seen = {};
  rows.slice(1).forEach(function (r, i) {
    if (
      !r.some(function (v) {
        return String(v || "").trim();
      })
    )
      return;
    var target = {
        nik: normalizeNik_(r[ix.nik]),
        program: String(r[ix.program] || "").toUpperCase(),
        batch: String(r[ix.batch] || "").trim(),
      },
      key = entityKey_(target.nik, target.program, target.batch, target.segment);
    if (seen[key])
      throw new Error("Enrollment duplikat di file pada baris " + (i + 2));
    seen[key] = true;
    var t = authorizeEnrollment_(s, target);
    if (!isPassingStatus_(t.enrollment.data.ProgramStatus))
      throw new Error(
        "Enrollment " +
          key +
          " belum lulus. Tetapkan hasilnya pada Participant Control.",
      );
    var rawIssue = ix.issue >= 0 ? r[ix.issue] : "", issue = certificateIssueDate_(rawIssue);
    if (rawIssue !== undefined && rawIssue !== null && String(rawIssue).trim() !== "" && !issue)
      throw new Error("Tanggal terbit tidak valid pada baris " + (i + 2) + ". Gunakan YYYY-MM-DD atau DD/MM/YYYY.");
    input.push({
      target: target,
      t: t,
      name: ix.name >= 0 ? r[ix.name] : "",
      number: ix.number >= 0 ? r[ix.number] : "",
      issue: issue,
    });
  });
  return bulkWrites_(function () {
    var batchId = newId_("CIMP");
    input.forEach(function (x, i) {
      appendObject_(
        getSheet_(SHEET_CERT_IMPORT),
        getHeaders_(getSheet_(SHEET_CERT_IMPORT)),
        {
          ImportID: newId_("CIMPR"),
          BatchID: batchId,
          RowNo: i + 2,
          NIK: x.target.nik,
          Program: x.target.program,
          Batch: x.target.batch,
          EnrollmentID: x.t.enrollment.data.EnrollmentID,
          DisplayName: x.name || x.t.participant.data.Name,
          FinalStatus: x.t.enrollment.data.ProgramStatus,
          FinalScore: x.t.enrollment.data.FinalScore,
          CertificateNumber: x.number || "",
          IssueDate: x.issue || formatDate_(new Date()),
          RowStatus: "UPLOADED",
          UploadedBy: s.nik,
          UploadedAt: nowIso_(),
        },
      );
    });
    return {
      ok: true,
      batchId: batchId,
      message: input.length + " sertifikat siap masuk antrean.",
    };
  });
}
function generateCertificatesFromImport_(token, options) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  options = options || {};
  return bulkWrites_(function () {
    var pending = readAllAsObjects_(SHEET_CERT_IMPORT)
      .filter(function (r) {
        return (
          (!options.batchId || r.data.BatchID === options.batchId) &&
          r.data.RowStatus === "UPLOADED"
        );
      })
      .slice(0, 50);
    pending.forEach(function (r) {
      var t = authorizeEnrollment_(s, {
          nik: r.data.NIK,
          enrollmentId: r.data.EnrollmentID,
        }),
        j = enqueueCertificate_(t, s, {
          displayName: r.data.DisplayName,
          certificateNumber: r.data.CertificateNumber,
          issueDate: r.data.IssueDate,
          importId: r.data.ImportID,
        });
      updateRowByIndex_(
        getSheet_(SHEET_CERT_IMPORT),
        getHeaders_(getSheet_(SHEET_CERT_IMPORT)),
        r.rowIndex,
        {
          RowStatus: j.ready ? "GENERATED" : "QUEUED",
          JobID: j.jobId || "",
          Message: "",
        },
      );
    });
    return {
      ok: true,
      queued: pending.length,
      message: pending.length + " sertifikat ditambahkan ke antrean.",
    };
  });
}
function getCertificateImportStatus_(token, batchId) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var all = readAllAsObjects_(SHEET_CERT_IMPORT);
  if (!batchId && all.length) batchId = all[all.length - 1].data.BatchID;
  return {
    batchId: batchId || "",
    rows: all
      .filter(function (r) {
        return r.data.BatchID === batchId;
      })
      .map(function (r) {
        return r.data;
      }),
  };
}

// ---------------------------------------------------------------------------
// Protected setup/migration. Private continuation workers end in _.
// ---------------------------------------------------------------------------
// Setup is resumable. A private continuation trigger processes short chunks.
// Only missing schema fields and missing migration links are written.
var LDP_SETUP_KEY_ = "LDP_SETUP_STATE";
var LDP_SETUP_STEPS_ = ["SCHEMA", "SEGMENTS", "DEFAULTS", "ASSESSMENTS", "ROOT", "USERS", "ENROLLMENTS", "PARTICIPANTS", "CERTIFICATES", "LEGACY_CERTIFICATES", "TRIGGERS"];
function setupState_() { return parseJson_(props_().getProperty(LDP_SETUP_KEY_), null); }
function saveSetupState_(job) { job.updatedAt = nowIso_(); props_().setProperty(LDP_SETUP_KEY_, JSON.stringify(job)); }
function setupStatus_(job) {
  if (!job) return {status:"NOT_STARTED", message:"Jalankan setupLdpSystem."};
  return {status:job.status,stage:job.stage,cursor:job.cursor,current:job.current||"",error:job.error||"",attempts:(job.current==="Membuka database"?job.openAttempts:job.attempts)||0,updatedAt:job.updatedAt,report:job.report,
    message:job.status==="DONE"?"SETUP LDP BERHASIL. Perbarui deployment lalu muat ulang aplikasi.":job.status==="FAILED"?"Setup berhenti pada "+job.current+". "+job.error+" Periksa kendala lalu jalankan setupLdpSystem untuk melanjutkan.":"Setup bertahap: "+job.stage+" · posisi "+job.cursor+". Proses dilanjutkan otomatis; jalankan getLdpSetupStatus untuk melihat progres."};
}
function removeSetupTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==="continueLdpSetup_")ScriptApp.deleteTrigger(t);});
}
function ensureSetupTrigger_() {
  var list=ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()==="continueLdpSetup_";});
  list.slice(1).forEach(function(t){ScriptApp.deleteTrigger(t);});
  if(!list.length)ScriptApp.newTrigger("continueLdpSetup_").timeBased().everyMinutes(1).create();
}
function setupLdpSystem_() {
  CTX_=null;context_().segment="ALL";
  var ss=setupSpreadsheet_(),id=ss.getId();
  var executor=String(Session.getEffectiveUser().getEmail()||"").trim().toLowerCase();
  var job=withLock_(function(){var old=setupState_();
    if(old&&old.status==="RUNNING"&&Number(old.leaseUntil)>Date.now())return old;
    if(!old||old.status==="DONE"||old.databaseId!==id||old.version!==LDP_VERSION_)
      old={version:LDP_VERSION_,databaseId:id,executor:executor,status:"PENDING",stage:"SCHEMA",cursor:0,attempts:0,report:{createdEnrollments:0,linkedCertificates:0,stagedLegacyCertificates:0,conflicts:[]},startedAt:nowIso_()};
    else {old.status="PENDING";old.attempts=0;old.error="";old.executor=executor;}
    props_().setProperty("LDP_SPREADSHEET_ID",id);saveSetupState_(old);return old;
  });
  if(job.status==="RUNNING")return setupStatus_(job);
  try{ensureSetupTrigger_();}
  catch(error){job.status="FAILED";job.current="Pembuatan trigger kelanjutan";job.error=String(error.message||error);saveSetupState_(job);throw error;}
  return continueLdpSetup_();
}
function getLdpSetupStatus() {
  requireSpreadsheetEditor_();
  var result=setupStatus_(setupState_());console.log(JSON.stringify(result));return result;
}
function setupNext_(job) { var index=LDP_SETUP_STEPS_.indexOf(job.stage);job.stage=LDP_SETUP_STEPS_[index+1]||"DONE";job.cursor=0; }
function mergeSetupReport_(job,delta) {
  ["createdEnrollments","linkedCertificates","stagedLegacyCertificates"].forEach(function(k){job.report[k]+=Number(delta[k]||0);});
  (delta.conflicts||[]).forEach(function(message){if(job.report.conflicts.indexOf(message)<0&&job.report.conflicts.length<30)job.report.conflicts.push(message);});
}
function continueLdpSetup_() {
  CTX_=null;context_().segment="ALL";
  var token=Utilities.getUuid(),job=withLock_(function(){var current=setupState_();
    if(!current||["DONE","FAILED"].indexOf(current.status)>=0)return null;
    if(current.status==="RUNNING"&&Number(current.leaseUntil)>Date.now())return null;
    var executor=String(Session.getEffectiveUser().getEmail()||"").trim().toLowerCase();
    if(current.executor&&current.executor!==executor)return null;
    if(current.version!==LDP_VERSION_||current.databaseId!==setupSpreadsheetId_()){
      current.status="FAILED";current.error="Versi kode atau database berubah selama setup. Jalankan setupLdpSystem dari editor.";saveSetupState_(current);return null;
    }
    current.status="RUNNING";current.leaseToken=token;current.leaseUntil=Date.now()+420000;saveSetupState_(current);return current;
  });
  if(!job){var inactive=setupState_();if(inactive&&["DONE","FAILED"].indexOf(inactive.status)>=0)try{removeSetupTrigger_();}catch(cleanupError){}return setupStatus_(inactive);}
  var started=Date.now(),steps=0;
  try{
    job.current="Membuka database";
    if(Number(job.openAttempts||0)>=3)throw new Error("Batas tiga percobaan membuka database tercapai. Periksa akses atau kondisi Google Sheets.");
    job.openAttempts=Number(job.openAttempts||0)+1;saveSetupState_(job);
    var ss=SpreadsheetApp.openById(job.databaseId);context_().ss=ss;job.openAttempts=0;
    while(job.stage!=="DONE"&&steps<12&&Date.now()-started<25000){
      job.current=job.stage==="SCHEMA"?("Header: "+Object.keys(SHEET_HEADERS)[job.cursor]):job.stage+" · "+job.cursor;
      if(Number(job.attempts||0)>=3)throw new Error("Batas tiga percobaan tahap tercapai setelah proses sebelumnya terputus. Periksa sheet yang ditampilkan pada status.");
      job.attempts=Number(job.attempts||0)+1;saveSetupState_(job);
      runSetupStep_(job,ss);steps++;job.attempts=0;job.error="";job.leaseUntil=Date.now()+420000;
      // Progress advances only after the current step's data writes have succeeded.
      saveSetupState_(job);
    }
    if(job.stage==="DONE"){
      resetSourceRevisions_();
      props_().setProperty("LDP_SCHEMA_VERSION",LDP_SCHEMA_);
      job.status="DONE";job.current="Selesai";job.leaseToken="";job.leaseUntil=0;saveSetupState_(job);
      try{removeSetupTrigger_();}catch(cleanupError){console.warn("Trigger kelanjutan belum terhapus: "+cleanupError.message);}
    }else{job.status="PENDING";job.leaseToken="";job.leaseUntil=0;saveSetupState_(job);}
  }catch(error){
    // Schema and migration steps are idempotent. Retry only transient service failures.
    var transient=/timed out|timeout|temporarily unavailable|internal error|try again later/i.test(String(error.message||error));
    job.error=String(error.message||error).slice(0,1000);job.status=transient&&Number(job.current==="Membuka database"?job.openAttempts:job.attempts)<3?"PENDING":"FAILED";job.leaseToken="";job.leaseUntil=0;saveSetupState_(job);
    if(job.status==="FAILED")try{removeSetupTrigger_();}catch(cleanupError){}
    console.warn(job.current+": "+job.error);
  }
  var result=setupStatus_(job);console.log(JSON.stringify(result));return result;
}
function setupDefaults_() {
  var defaults={PROGRAM_LABEL_SLDP:"SLDP",PHASE1_SLDP_SHEET:"Phase 1 SLDP B2C",PHASE2_SLDP_SHEET:"Phase 2 SLDP B2C",PHASE1_B2C_TSHDP_SHEET:"Phase 1 TSHDP B2C",PHASE2_B2C_TSHDP_SHEET:"Phase 2 TSHDP B2C",APP_NAME:"LDP Monitor",TIMEZONE:TZ,MAX_FILE_MB:"20",SESSION_HOURS:"6",CLP_ENABLED:"FALSE",SUPERIOR_ENABLED:"FALSE",PASSING_SCORE:"80",ROOT_FOLDER_ID:"",SOURCE_PARTICIPANT_SHEET:"Legacy Participants",PHASE1_TSHDP_SHEET:"Phase 1 TSHDP",PHASE2_TSHDP_SHEET:"Phase 2 TSHDP",PHASE1_ASHDP_SHEET:"Phase 1 ASHDP",PHASE2_ASHDP_SHEET:"Phase 2 ASHDP",PHASE3_SHEET:"Phase 3",QLEAP_MAPPING_SHEET:"QLEAP Mapping",CERTIFICATE_TEMPLATE_ID:"",CERTIFICATE_PREFIX:"LDP",CERTIFICATE_SIGNER_NAME:"",CERTIFICATE_SIGNER_TITLE:"",PROGRAM_LABEL_ASHDP:"Area Sales Head Development Program",PROGRAM_LABEL_TSHDP:"Territory Sales Head Development Program"};
  bulkWrites_(function(){var existing=getAllSettings_(),sh=getSheet_(SHEET_SETTINGS),headers=getHeaders_(sh);Object.keys(defaults).forEach(function(k){if(!Object.prototype.hasOwnProperty.call(existing,k))appendObject_(sh,headers,{Key:k,Value:defaults[k]});});});
  context_().settings=null;
}
function ensureSetupColumns_(sheet,count) {var current=sheet.getMaxColumns();if(current<count)sheet.insertColumnsAfter(current,count-current);}
function ensureSheetWithHeaders_(ss,name,headers) {
  var sheet=ss.getSheetByName(name)||ss.insertSheet(name),last=sheet.getLastColumn();
  var old=last?sheet.getRange(1,1,1,last).getValues()[0].map(String):[];
  var missing=headers.filter(function(h){return old.indexOf(h)<0;});
  if(missing.length){ensureSetupColumns_(sheet,last+missing.length);sheet.getRange(1,last+1,1,missing.length).setValues([missing]);}
  // Do not freeze or restyle sheets during setup. Avoid cosmetic calls on the live timeout path.
  return sheet;
}
function ensureSetupAssessment_(ss,cfg) {
  var name=getSetting_(cfg.settingKey)||cfg.defaultSheet;if(ss.getSheetByName(name))return;
  var sheet=ss.insertSheet(name);ensureSetupColumns_(sheet,cfg.headers.length);
  sheet.getRange(cfg.headerRow,1,1,cfg.headers.length).setValues([cfg.headers]);
}
function ensureAssessmentTemplateSheets_(ss) {
  Object.keys(PROGRAM_PHASE_ASSESSMENT_CONFIG).forEach(function(k){ensureSetupAssessment_(ss,PROGRAM_PHASE_ASSESSMENT_CONFIG[k]);});
}
function setupDataBatch_(job,stage) {
  return bulkWrites_(function(){
    var table=stage==="USERS"?SHEET_USERS:stage==="ENROLLMENTS"?SHEET_ENROLLMENTS:stage==="CERTIFICATES"?SHEET_CERTIFICATES:SHEET_PARTICIPANTS;
    var all=readAllAsObjects_(table),selected=all.filter(function(r){return r.rowIndex>=Number(job.cursor||2);}).slice(0,50),delta={conflicts:[]};
    var erIndex=["PARTICIPANTS","CERTIFICATES","LEGACY_CERTIFICATES"].indexOf(stage)>=0?enrollmentIndex_():null;
    var participantCounts={},certificateIds={};
    if(stage==="PARTICIPANTS")all.forEach(function(r){var p=r.data,k=entityKey_(p.NIK,p.Program,p.Batch, p.Segment);participantCounts[k]=(participantCounts[k]||0)+1;});
    if(stage==="LEGACY_CERTIFICATES")readAllAsObjects_(SHEET_CERTIFICATES).forEach(function(r){if(r.data.EnrollmentID)certificateIds[r.data.EnrollmentID]=true;});
    selected.forEach(function(r){
      var d=r.data,sh=getSheet_(table),headers=getHeaders_(sh);
      if(stage==="USERS"){if(!d.AuthVersion)updateRowByIndex_(sh,headers,r.rowIndex,{AuthVersion:1});return;}
      if(stage==="ENROLLMENTS"){if(!d.RecordVersion||!d.EnrollmentID)updateRowByIndex_(sh,headers,r.rowIndex,{EnrollmentID:d.EnrollmentID||newId_("ENR"),RecordVersion:Utilities.getUuid()});return;}
      if(stage==="PARTICIPANTS"){
        var program=String(d.Program||"").toUpperCase(),key=entityKey_(d.NIK,program,d.Batch, d.Segment);
        if(!d.NIK||activePrograms_().indexOf(program)<0||!normalizeBatchKey_(d.Batch))return;
        if(participantCounts[key]>1){delta.conflicts.push("Participants duplikat "+key);return;}
        var er=erIndex.byKey[key];
        if(!er){var newEnrollment={Segment:d.Segment,EnrollmentID:newId_("ENR"),NIK:normalizeNik_(d.NIK),Program:program,Batch:d.Batch,Region:d.Region,CurrentPhase:normalizeCurrentPhase_(d.CurrentPhase||inferCurrentPhase_(d)),EnrollmentStatus:normalizeCurrentPhase_(d.CurrentPhase||inferCurrentPhase_(d))==="COMPLETED"?"COMPLETED":"ON GOING",ProgramStatus:d.FinalStatus||"ON GOING",IsActive:false,Phase1Score:d.Phase1Score,Phase2Score:d.Phase2Score,Phase3Score:d.Phase3Score,FinalScore:d.FinalScore,CertificateUrl:d.CertificateUrl||"",CreatedAt:nowIso_(),UpdatedAt:nowIso_()};
          var es=getSheet_(SHEET_ENROLLMENTS),row=appendObject_(es,getHeaders_(es),newEnrollment);er={rowIndex:row,data:newEnrollment};erIndex.byKey[key]=er;erIndex.byId[newEnrollment.EnrollmentID]=er;(erIndex.byNik[newEnrollment.NIK]||(erIndex.byNik[newEnrollment.NIK]=[])).push(er);delta.createdEnrollments=(delta.createdEnrollments||0)+1;
        }
        if(d.CurrentEnrollmentId!==er.data.EnrollmentID)updateRowByIndex_(sh,headers,r.rowIndex,{CurrentEnrollmentId:er.data.EnrollmentID});return;
      }
      if(stage==="CERTIFICATES"){
        if(d.EnrollmentID)return;
        var list=(erIndex.byNik[normalizeNik_(d.NIK)]||[]).filter(function(e){return normalizeSegment_(e.data.Segment)===normalizeSegment_(d.Segment)&&(!d.Program||String(e.data.Program).toUpperCase()===String(d.Program).toUpperCase())&&(!d.Batch||sameBatch_(e.data.Batch,d.Batch));});
        if(list.length===1){updateRowByIndex_(sh,headers,r.rowIndex,{EnrollmentID:list[0].data.EnrollmentID,Program:list[0].data.Program,Batch:list[0].data.Batch});delta.linkedCertificates=(delta.linkedCertificates||0)+1;}
        else delta.conflicts.push("Sertifikat belum dipetakan NIK "+d.NIK+" CertificateID "+d.CertificateID);return;
      }
      var existingEr=erIndex.byKey[entityKey_(d.NIK,d.Program,d.Batch, d.Segment)],url=d.CertificateUrl||(existingEr&&existingEr.data.CertificateUrl);
      if(!existingEr||!url||certificateIds[existingEr.data.EnrollmentID])return;
      validateUrl_(url);var cs=getSheet_(SHEET_CERTIFICATES);
      appendObject_(cs,getHeaders_(cs),{Segment:d.Segment,CertificateID:newId_("CERT"),EnrollmentID:existingEr.data.EnrollmentID,NIK:d.NIK,Program:d.Program,Batch:d.Batch,DisplayName:d.Name,FinalScore:existingEr.data.FinalScore,FinalStatus:existingEr.data.ProgramStatus,FileOrLink:url,Published:false,UploadedBy:"MIGRATION",UploadedAt:nowIso_()});certificateIds[existingEr.data.EnrollmentID]=true;delta.stagedLegacyCertificates=(delta.stagedLegacyCertificates||0)+1;
    });
    return {next:selected.length?selected[selected.length-1].rowIndex+1:Number(job.cursor||2),done:selected.length<50,delta:delta};
  });
}
function setupAppTriggers_(ss) {
  if(readAllAsObjects_(SHEET_JOBS_).some(function(r){return["QUEUED","PROCESSING"].indexOf(r.data.Status)>=0;}))props_().setProperty("LDP_CERT_PENDING","1");
  ScriptApp.getProjectTriggers().forEach(function(t){if(["autoGenerateCertificates","sendDeadlineReminders","processCertificateJobs_","handleSpreadsheetEdit_","handleSpreadsheetChange_"].indexOf(t.getHandlerFunction())>=0)ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger("processCertificateJobs_").timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger("handleSpreadsheetEdit_").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("handleSpreadsheetChange_").forSpreadsheet(ss).onChange().create();
}
function runSetupStep_(job,ss) {
  if(job.stage==="SCHEMA"){var names=Object.keys(SHEET_HEADERS);if(job.cursor<names.length){ensureSheetWithHeaders_(ss,names[job.cursor],SHEET_HEADERS[names[job.cursor]]);job.cursor++;}if(job.cursor>=names.length){props_().setProperty("LDP_DATA_REV",Utilities.getUuid());context_().revision=undefined;context_().headers={};context_().tables={};context_().memo={};setupNext_(job);}return;}
  if(job.stage==="SEGMENTS"){
    var segmentNames=Object.keys(SHEET_HEADERS).filter(function(name){return [SHEET_USERS,SHEET_USER_IMPORT,SHEET_SETTINGS].indexOf(name)<0;});
    if(job.cursor>=segmentNames.length){delete job.segmentRow;setupNext_(job);return;}
    var segmentName=segmentNames[job.cursor],sheet=ss.getSheetByName(segmentName),h=getHeaders_(sheet),col=h.indexOf("Segment"),start=Number(job.segmentRow || 2),last=sheet.getLastRow();
    if(col<0 || start>last){job.cursor++;delete job.segmentRow;return;}
    var count=Math.min(200,last-start+1);
    withLock_(function(){
      var identityCol=h.indexOf("NIK");if(identityCol<0)identityCol=h.findIndex(function(k){return k!=="Segment";});
      var values=sheet.getRange(start,col+1,count,1).getValues(),
        identities=identityCol>=0?sheet.getRange(start,identityCol+1,count,1).getValues():[],runs=[],run;
      values.forEach(function(row,i){
        if(!String((identities[i]||[])[0] || "").trim() || String(row[0] || "").trim()) {run=null;return;}
        if(!run){run={row:start+i,values:[]};runs.push(run);}run.values.push(["B2B"]);
      });
      runs.forEach(function(r){sheet.getRange(r.row,col+1,r.values.length,1).setValues(r.values);context_().writes++;});
      if(runs.length){delete context_().tables[segmentName];invalidateSheet_(segmentName);}
    });
    job.segmentRow=start+count;if(job.segmentRow>last){job.cursor++;delete job.segmentRow;}return;
  }
  if(job.stage==="DEFAULTS"){setupDefaults_();setupNext_(job);return;}
  if(job.stage==="ASSESSMENTS"){
    var keys=Object.keys(PROGRAM_PHASE_ASSESSMENT_CONFIG);
    if(job.cursor<keys.length)ensureSetupAssessment_(ss,PROGRAM_PHASE_ASSESSMENT_CONFIG[keys[job.cursor]]);
    else if(job.cursor===keys.length){var name=getSetting_("QLEAP_MAPPING_SHEET")||SHEET_QLEAP;if(name!==SHEET_QLEAP)ensureSheetWithHeaders_(ss,name,SHEET_HEADERS[SHEET_QLEAP]);}
    else{var phase3=getSetting_("PHASE3_SHEET")||"Phase 3";ensureSheetWithHeaders_(ss,phase3,["NIK","Program","Batch","PreTest","PostTest","UpdatedAt","Segment"]);}
    job.cursor++;if(job.cursor>=keys.length+2)setupNext_(job);return;
  }
  if(job.stage==="ROOT"){if(!getSetting_("ROOT_FOLDER_ID"))setSetting_("ROOT_FOLDER_ID",DriveApp.createFolder("LDP Monitor Files").getId());setupNext_(job);return;}
  if(["USERS","ENROLLMENTS","PARTICIPANTS","CERTIFICATES","LEGACY_CERTIFICATES"].indexOf(job.stage)>=0){var result=setupDataBatch_(job,job.stage);mergeSetupReport_(job,result.delta);job.cursor=result.next;if(result.done)setupNext_(job);return;}
  if(job.stage==="TRIGGERS"){setupAppTriggers_(ss);setupNext_(job);return;}
  throw new Error("Tahap setup tidak dikenal: "+job.stage);
}

function handleSpreadsheetEdit_(event) {
  CTX_ = null;
  try {
    if (event && event.range && event.range.getSheet().getName() === SHEET_USER_IMPORT &&
        props_().getProperty("LDP_SCHEMA_VERSION") === LDP_SCHEMA_) {
      CacheService.getScriptCache().remove(userImportCheckKey_());
      syncUserImport_(true, null);
      return;
    }
    if (
      event &&
      event.range &&
      event.range.getSheet().getName() === SHEET_USERS
    ) {
      var sh = event.range.getSheet(),
        h = getHeaders_(sh),
        col = h.indexOf("AuthVersion");
      if (col >= 0)
        for (
          var row = Math.max(2, event.range.getRow());
          row < event.range.getRow() + event.range.getNumRows();
          row++
        ) {
          var value = Number(
            sh.getRange(row, col + 1, 1, 1).getValues()[0][0] || 1,
          );
          sh.getRange(row, col + 1, 1, 1).setValues([[value + 1]]);
        }
    }
  } finally {
    resetSourceRevisions_();
  }
}
function handleSpreadsheetChange_(event) {
  CTX_ = null;
  resetSourceRevisions_();
}
function initializeAdminPassword_() {
  CTX_ = null;
  var nik = props_().getProperty("LDP_BOOTSTRAP_ADMIN_NIK"),
    password = props_().getProperty("LDP_BOOTSTRAP_ADMIN_PASSWORD");
  if (!nik || !password || password.length < 8)
    throw new Error(
      "Isi Script Properties LDP_BOOTSTRAP_ADMIN_NIK dan LDP_BOOTSTRAP_ADMIN_PASSWORD, minimal 8 karakter.",
    );
  try {
    return api_("initializeAdminPassword", function () {
      return withLock_(function () {
        var user = findUserRowByNik_(nik);
        if (!user || !userHasRole_(user.data, "ADMIN"))
          throw new Error("Provision akun ADMIN terlebih dahulu.");
        var salt = generateSalt_();
        updateRowByIndex_(
          getSheet_(SHEET_USERS),
          getHeaders_(getSheet_(SHEET_USERS)),
          user.rowIndex,
          {
            PasswordHash: hashPassword_(password, salt),
            Salt: salt,
            Active: true,
            UpdatedAt: nowIso_(),
          },
        );
        return "Password Admin telah diinisialisasi.";
      });
    }, arguments);
  } finally {
    props_().deleteProperty("LDP_BOOTSTRAP_ADMIN_PASSWORD");
  }
}
function provisionUsersFromImportSheet_() {
  return api_("provisionUsers", function () {
    return provisionUsers_({ nik: "EDITOR", role: "ADMIN" });
  }, arguments);
}
// Read the trusted import sheet live. Public login payloads cannot assign roles.
function readUserImportInput_() {
  var sh = getSheet_(SHEET_USER_IMPORT), last = sh.getLastRow(), columns = sh.getLastColumn();
  if (!columns) throw new Error("Header User Import belum tersedia.");
  var headers = sh.getRange(1, 1, 1, columns).getDisplayValues()[0].map(function(h) { return String(h).trim(); });
  if (headers.indexOf("Password") < 0) {
    if (columns >= sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
    sh.getRange(1, columns + 1, 1, 1).setValues([["Password"]]);
    columns++; headers.push("Password");
    delete context_().headers[SHEET_USER_IMPORT];
    context_().writes++; invalidateSheet_(SHEET_USER_IMPORT);
    SpreadsheetApp.flush();
  }
  if (headers.indexOf("Password") !== headers.lastIndexOf("Password")) throw new Error("Header User Import tidak valid: Password");
  var values = sh.getRange(1, 1, Math.max(1, last), columns).getDisplayValues();
  context_().reads += 2; context_().cells += values.length * columns;
  var seen = {};
  ["NIK", "Name", "Role", "Active"].forEach(function(h) {
    if (headers.indexOf(h) < 0 || headers.lastIndexOf(h) !== headers.indexOf(h))
      throw new Error("Header User Import tidak valid: " + h);
  });
  return values.slice(1).map(function(row, i) {
    if (!row.some(function(v) { return String(v).trim() !== ""; })) return null;
    var d = rowToObject_(headers, row), nik = normalizeNik_(d.NIK), roles = parseRolesStrict_(d.Role);
    if (!nik || !roles.roles.length || roles.invalid.length)
      throw new Error("NIK/Role tidak valid pada User Import baris " + (i + 2));
    if (seen[nik]) throw new Error("NIK ganda pada User Import: " + nik);
    seen[nik] = true;
    var password = roles.roles.indexOf(ROLE_ADMIN) >= 0 ? String(d.Password || "") : "";
    if (password && (password.length < 8 || !password.trim()))
      throw new Error("Password Admin pada User Import baris " + (i + 2) + " minimal 8 karakter.");
    return {NIK:nik, Name:String(d.Name || "").trim(), Role:joinRoles_(roles.roles),
      Region:normalizeRegion_(d.Region), Superior_NIK:normalizeNik_(d.Superior_NIK),
      Email:String(d.Email || "").trim(), Active:parseBoolean_(d.Active), Password:password};
  }).filter(function(d) { return !!d; });
}
function userImportCheckKey_() {
  return cacheKey_("ldpui:checked:", LDP_VERSION_ + "|" + (props_().getProperty("LDP_SPREADSHEET_ID") || "bound"));
}
function syncUserImport_(force, actor) {
  var cache = CacheService.getScriptCache(), key = userImportCheckKey_();
  if (!force && cache.get(key)) return {ok:true, skipped:true};
  return withLock_(function() {
    // Collapse simultaneous login checks; only the first reads the import sheet.
    if (!force && cache.get(key)) return {ok:true, skipped:true};
    var input = readUserImportInput_(), signature = cacheKey_("", JSON.stringify(input)),
      signatureKey = "LDP_USER_IMPORT_SIGNATURE";
    if (!force && props_().getProperty(signatureKey) === signature) {
      cache.put(key, "1", 30); return {ok:true, skipped:true};
    }
    var changed = 0;
    bulkWrites_(function() {
      var users = readAllAsObjects_(SHEET_USERS), idx = {}, sh = getSheet_(SHEET_USERS), h = getHeaders_(sh), participantNiks=null;
      users.forEach(function(r) {
        var nik = normalizeNik_(r.data.NIK);
        if (idx[nik]) throw new Error("Users duplikat " + nik);
        idx[nik] = r;
      });
      input.forEach(function(d) {
        var r = idx[d.NIK], patch = {};
        if(r && userHasRole_(r.data,ROLE_PARTICIPANT) && parseRoles_(d.Role).indexOf(ROLE_PARTICIPANT)<0) {
          if(!participantNiks) {
            participantNiks={};allSegmentRows_(SHEET_PARTICIPANTS).forEach(function(row) {participantNiks[normalizeNik_(row.data.NIK)]=true;});
          }
          if(participantNiks[d.NIK]) d=Object.assign({},d,{Role:joinRoles_(parseRoles_(d.Role).concat([ROLE_PARTICIPANT]))});
        }
        Object.keys(d).forEach(function(k) {
          if (k === "Password") return;
          var previous = r && r.data[k];
          if (k === "Active") previous = parseBoolean_(previous);
          if (!r || String(previous === undefined || previous === null ? "" : previous) !== String(d[k])) patch[k] = d[k];
        });
        // The import password is never copied to Users, sessions, logs, or a snapshot cache.
        if (d.Password && (!r || !r.data.PasswordHash || hashPassword_(d.Password, r.data.Salt) !== r.data.PasswordHash)) {
          patch.Salt = generateSalt_(); patch.PasswordHash = hashPassword_(d.Password, patch.Salt);
        }
        if (!Object.keys(patch).length) return;
        patch.UpdatedAt = nowIso_();
        if (r) updateRowByIndex_(sh, h, r.rowIndex, patch);
        else { patch.CreatedAt = nowIso_(); var row = appendObject_(sh, h, patch); idx[d.NIK] = {rowIndex:row, data:patch}; }
        changed++;
      });
      if (actor && changed) auditLog_(actor.nik, actor.role, "PROVISION_USERS", "Users", "", changed + " akun diperbarui");
    });
    // Blank source passwords retain credentials; changed access/passwords revoke old sessions.
    // Accounts absent from the source are retained.
    if (changed) SpreadsheetApp.flush();
    props_().setProperty(signatureKey, signature);
    cache.put(key, "1", 30);
    console.log(JSON.stringify({event:"USER_IMPORT_SYNC", scanned:input.length, changed:changed}));
    return {ok:true, changed:changed, message:input.length + " akun diperiksa; " + changed + " akun diperbarui."};
  });
}
function provisionUsers_(s) {
  return syncUserImport_(true, s);
}
function autoSyncUserImport_() {
  // Source errors are visible to the operator without blocking established users.
  try { return syncUserImport_(false, null); }
  catch (error) {
    console.warn("USER_IMPORT_SYNC_FAILED: " + String(error.message || error));
    return {ok:false, message:String(error.message || error)};
  }
}
function syncParticipantsFromSourceSheet_(s) {
  var name = getSetting_("SOURCE_PARTICIPANT_SHEET");
  if (!name) throw new Error("Sheet sumber belum dikonfigurasi.");
  var source = readAllAsObjects_(name);
  return bulkWrites_(function () {
    var count = 0,
      seen = {};
    source.forEach(function (r) {
      var d = r.data,
        nik = normalizeNik_(d.NIK || d["Employee ID"]);
      if (!nik) return;
      var matches = readAllAsObjects_(SHEET_PARTICIPANTS).filter(function (x) {
        return (
          normalizeNik_(x.data.NIK) === nik &&
          (!d.Program ||
            String(x.data.Program).toUpperCase() ===
              String(d.Program).toUpperCase()) &&
          (!d.Batch || sameBatch_(x.data.Batch, d.Batch))
        );
      });
      if (matches.length !== 1)
        throw new Error(
          "Sumber NIK " +
            nik +
            " harus memuat Program/Batch yang cocok dan unik. Gunakan import peserta untuk enrollment baru.",
        );
      var p = matches[0],
        key = entityKey_(p.data.NIK, p.data.Program, p.data.Batch, p.data.Segment);
      if (seen[key]) throw new Error("Profil sumber duplikat: " + key);
      seen[key] = true;
      var patch = { UpdatedAt: nowIso_() },
        fields = {
          Name: d.Name || d["Employee Name"],
          JobTitle: d.JobTitle || d["Job Title Name"],
          Email: d.Email,
          Location: d.Location,
          CityName: d.CityName || d.City,
        };
      Object.keys(fields).forEach(function (k) {
        if (fields[k] !== undefined && fields[k] !== "")
          patch[k] = safeCell_(String(fields[k]));
      });
      updateRowByIndex_(
        getSheet_(SHEET_PARTICIPANTS),
        getHeaders_(getSheet_(SHEET_PARTICIPANTS)),
        p.rowIndex,
        patch,
      );
      count++;
    });
    return {
      ok: true,
      message:
        count +
        " profil disinkronkan. Phase, status, program, batch, dan nilai tidak ditimpa.",
    };
  });
}
function exportParticipantsCsv_(token, f) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN", "CLP"]);
  f = f || {};
  var rows = scopedParticipants_(s)
      .map(function (r) {
        return participantRow_(r.data);
      })
      .filter(function (r) {
        return (
          (!f.program || r.program === f.program) &&
          (!f.region || r.region === f.region) &&
          (!f.batch || sameBatch_(r.batch, f.batch)) &&
          matchesProgramTrack_(r.programTrack,f.track)
        );
      }),
    fields = [
      "segment",
      "nik",
      "name",
      "jobTitle",
      "region",
      "program",
      "batch",
      "programTrack",
      "participantStatus",
      "currentPhase",
      "phase1Score",
      "phase2Score",
      "phase3Score",
      "finalScore",
      "finalStatus",
      "enrollmentId",
    ];
  return {
    filename: "LDP_Peserta.csv",
    content:
      "\uFEFF" +
      [fields.map(csvEscape_).join(",")]
        .concat(
          rows.map(function (r) {
            return fields
              .map(function (h) {
                return csvEscape_(r[h]);
              })
              .join(",");
          }),
        )
        .join("\r\n"),
  };
}
function getParticipantImportTemplate_(token) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  return {
    filename: "Template_Peserta_LDP.csv",
    content:
      "\uFEFF" + PARTICIPANT_IMPORT_HEADERS.map(csvEscape_).join(",") + "\r\n",
  };
}
function getSystemSettings_(token) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var allowed = [
    "APP_NAME",
    "MAX_FILE_MB",
    "PASSING_SCORE",
    "CERTIFICATE_TEMPLATE_ID",
    "CERTIFICATE_PREFIX",
    "CERTIFICATE_SIGNER_NAME",
    "CERTIFICATE_SIGNER_TITLE",
  ];
  var out = {};
  allowed.forEach(function (k) {
    out[k] = getSetting_(k);
  });
  return { values: out, roleSettings: getRoleSettings_() };
}
function saveSystemSettings_(token, d) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var allowed = [
    "APP_NAME",
    "MAX_FILE_MB",
    "PASSING_SCORE",
    "CERTIFICATE_TEMPLATE_ID",
    "CERTIFICATE_PREFIX",
    "CERTIFICATE_SIGNER_NAME",
    "CERTIFICATE_SIGNER_TITLE",
  ];
  Object.keys(d || {}).forEach(function (k) {
    if (allowed.indexOf(k) < 0) throw new Error("Setting tidak diizinkan.");
    if (
      k === "MAX_FILE_MB" &&
      (!isFinite(Number(d[k])) || Number(d[k]) < 1 || Number(d[k]) > 20)
    )
      throw new Error("Batas file harus 1 sampai 20 MB.");
    if (k === "PASSING_SCORE") validScore_(d[k]);
  });
  return bulkWrites_(function () {
    Object.keys(d || {}).forEach(function (k) {
      setSetting_(k, String(d[k]));
    });
    return { ok: true };
  });
}

// Public endpoints form the complete RPC surface. No editor/worker entry is public.
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("LDP Monitor")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}
function getLoginConfig(options) {
  return api_("getLoginConfig", function () {
    return getLoginConfig_(options);
  }, arguments);
}
function login(d) {
  return api_("login", function () {
    return login_(d);
  }, arguments);
}
function logout(t) {
  return api_("logout", function () {
    return logout_(t);
  }, arguments);
}
function switchRole(t, r, p) {
  return api_("switchRole", function () {
    return switchRole_(t, r, p);
  }, arguments);
}
function getDashboard(t) {
  return api_("getDashboard", function () {
    return getDashboard_(t);
  }, arguments);
}
function getParticipantPage(t, f) {
  return api_("getParticipantPage", function () {
    return getParticipantPage_(t, f);
  }, arguments);
}
function getParticipantFullDetail(t, n) {
  return api_("getParticipantFullDetail", function () {
    return getParticipantFullDetail_(t, n);
  }, arguments);
}
function getParticipantWorkspace(t, d) {
  return api_("getParticipantWorkspace", function () {
    return getParticipantWorkspace_(t, d);
  }, arguments);
}
function getParticipantAssessmentDetail(t, n, p, g, b) {
  return api_("getParticipantAssessmentDetail", function () {
    return getParticipantAssessmentDetail_(t, n, p, g, b);
  }, arguments);
}
function saveAssessmentScores(t, d) {
  return api_("saveAssessmentScores", function () {
    return saveAssessmentScores_(t, d);
  }, arguments);
}
function saveParticipantControl(t, d) {
  return api_("saveParticipantControl", function () {
    return saveParticipantControl_(t, d);
  }, arguments);
}
function uploadParticipants(t, d) {
  return api_("uploadParticipants", function () {
    return uploadParticipants_(t, d);
  }, arguments);
}
function bulkUpdateParticipantPhase(t, d) {
  return api_("bulkUpdateParticipantPhase", function () {
    return bulkUpdateParticipantPhase_(t, d);
  }, arguments);
}
function updateParticipantResult(t, d, r) {
  return api_("updateParticipantResult", function () {
    return updateParticipantResult_(t, d, r);
  }, arguments);
}
function updateParticipantStatus(t, d, s) {
  return api_("updateParticipantStatus", function () {
    return updateParticipantStatus_(t, d, s);
  }, arguments);
}
function uploadProject(t, d) {
  return api_("uploadProject", function () {
    return uploadFile_(t, d, "PROJECT");
  }, arguments);
}
function uploadMentoringEvidence(t, d) {
  return api_("uploadMentoringEvidence", function () {
    return uploadFile_(t, d, "EVIDENCE");
  }, arguments);
}
function getUserAccountsPage(t, f) {
  return api_("getUserAccountsPage", function () {
    return getUserAccountsPage_(t, f);
  }, arguments);
}
function saveUserAccount(t, d) {
  return api_("saveUserAccount", function () {
    return saveUserAccount_(t, d);
  }, arguments);
}
function softDeleteUser(t, n) {
  return api_("softDeleteUser", function () {
    return softDeleteUser_(t, n);
  }, arguments);
}
function resetAdminPassword(t, n, p) {
  return api_("resetAdminPassword", function () {
    return resetAdminPassword_(t, n, p);
  }, arguments);
}
function provisionUsersFromImportSheetUI(t) {
  return api_("provisionUsersFromImportSheetUI", function () {
    var s = requireSession_(t);
    requireRole_(s, ["ADMIN"]);
    return provisionUsers_(s);
  }, arguments);
}
function syncParticipantsFromSourceSheetUI(t) {
  return api_("syncParticipantsFromSourceSheetUI", function () {
    var s = requireSession_(t);
    requireRole_(s, ["ADMIN"]);
    return syncParticipantsFromSourceSheet_(s);
  }, arguments);
}
function saveAdminRoleSettings(t, d) {
  return api_("saveAdminRoleSettings", function () {
    return saveAdminRoleSettings_(t, d);
  }, arguments);
}
function getAuditPage(t, f) {
  return api_("getAuditPage", function () {
    return getAuditPage_(t, f);
  }, arguments);
}
function getCertificatesOverview(t, f) {
  return api_("getCertificatesOverview", function () {
    return getCertificatesOverview_(t, f);
  }, arguments);
}
function manageCertificate(t, d) {
  return api_("manageCertificate", function () {
    return manageCertificate_(t, d);
  }, arguments);
}
function toggleCertificateActive(t, d, a) {
  return api_("toggleCertificateActive", function () {
    return toggleCertificateActive_(t, d, a);
  }, arguments);
}
function retryCertificateJob(t, id) {
  return api_("retryCertificateJob", function () {
    return retryCertificateJob_(t, id);
  }, arguments);
}
function getCertificateImportTemplate(t) {
  return api_("getCertificateImportTemplate", function () {
    return getCertificateImportTemplate_(t);
  }, arguments);
}
function uploadCertificateImport(t, d) {
  return api_("uploadCertificateImport", function () {
    return uploadCertificateImport_(t, d);
  }, arguments);
}
function generateCertificatesFromImport(t, o) {
  return api_("generateCertificatesFromImport", function () {
    return generateCertificatesFromImport_(t, o);
  }, arguments);
}
function getCertificateImportStatus(t, id) {
  return api_("getCertificateImportStatus", function () {
    return getCertificateImportStatus_(t, id);
  }, arguments);
}
function getQleapPostTestRecap(t, f) {
  return api_("getQleapPostTestRecap", function () {
    return getQleapPostTestRecap_(t, f);
  }, arguments);
}
function getMyQleapPostTest(t, options) {
  return api_("getMyQleapPostTest", function () {
    return getMyQleapPostTest_(t, options);
  }, arguments);
}
function uploadQleapPostTest(t, d) {
  return api_("uploadQleapPostTest", function () {
    return uploadQleapPostTest_(t, d);
  }, arguments);
}
function getQleapMappings(t) {
  return api_("getQleapMappings", function () {
    return getQleapMappings_(t);
  }, arguments);
}
function saveQleapMapping(t, d) {
  return api_("saveQleapMapping", function () {
    return saveQleapMapping_(t, d);
  }, arguments);
}
function getSystemSettings(t) {
  return api_("getSystemSettings", function () {
    return getSystemSettings_(t);
  }, arguments);
}
function saveSystemSettings(t, d) {
  return api_("saveSystemSettings", function () {
    return saveSystemSettings_(t, d);
  }, arguments);
}
function getParticipantImportTemplate(t) {
  return api_("getParticipantImportTemplate", function () {
    return getParticipantImportTemplate_(t);
  }, arguments);
}
function exportParticipantsCsv(t, f) {
  return api_("exportParticipantsCsv", function () {
    return exportParticipantsCsv_(t, f);
  }, arguments);
}

function getMigrationIssues_(token) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var out = [];
  [
    [SHEET_CERTIFICATES, "CertificateID", "Sertifikat"],
    [SHEET_SUBMISSIONS, "SubmissionID", "Project"],
    [SHEET_MENTORING, "EvidenceID", "Mentoring"],
  ].forEach(function (cfg) {
    readAllAsObjects_(cfg[0]).forEach(function (r) {
      if (!r.data.EnrollmentID)
        out.push({
          kind: cfg[2],
          id: r.data[cfg[1]],
          nik: normalizeNik_(r.data.NIK),
          name:
            r.data.FileName ||
            r.data.DisplayName ||
            r.data.CertificateNumber ||
            "",
          fileUrl: r.data.FileUrl || r.data.FileOrLink || "",
        });
    });
  });
  return out.slice(0, 500);
}
function assignLegacyRecord_(token, d) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var map = {
      Sertifikat: [SHEET_CERTIFICATES, "CertificateID"],
      Project: [SHEET_SUBMISSIONS, "SubmissionID"],
      Mentoring: [SHEET_MENTORING, "EvidenceID"],
    },
    cfg = map[d.kind];
  if (!cfg) throw new Error("Jenis data tidak valid.");
  return bulkWrites_(function () {
    var t = authorizeEnrollment_(s, d),
      r = uniqueMatch_(
        readAllAsObjects_(cfg[0]).filter(function (r) {
          return (
            String(r.data[cfg[1]]) === String(d.id) &&
            normalizeNik_(r.data.NIK) === normalizeNik_(d.nik)
          );
        }),
        "record lama",
      );
    if (!r) throw new Error("Dokumen tidak ditemukan.");
    if (r.data.EnrollmentID && r.data.EnrollmentID !== d.enrollmentId)
      throw new Error("Dokumen sudah dipetakan.");
    var patch = { EnrollmentID: d.enrollmentId };
    if (d.kind === "Sertifikat") {
      if (certificateFor_(t.participant.data))
        throw new Error(
          "Enrollment sudah memiliki sertifikat. Periksa konflik sebelum menggabungkan.",
        );
      patch.Program = t.participant.data.Program;
      patch.Batch = t.participant.data.Batch;
    }
    updateRowByIndex_(
      getSheet_(cfg[0]),
      getHeaders_(getSheet_(cfg[0])),
      r.rowIndex,
      patch,
    );
    auditLog_(s.nik, s.role, "ASSIGN_LEGACY", cfg[0], d.id, d.enrollmentId);
    return { ok: true };
  });
}
function getMigrationIssues(t) {
  return api_("getMigrationIssues", function () {
    return getMigrationIssues_(t);
  }, arguments);
}
function assignLegacyRecord(t, d) {
  return api_("assignLegacyRecord", function () {
    return assignLegacyRecord_(t, d);
  }, arguments);
}

// Explicit source-to-summary refresh, separate from dashboard reads.
function refreshEnrollmentSummary_(t) {
  var sourceTransition=activateSourceGrading_(t)||t.sourceTransition, sourceDriven=legacySnapshot_(t.enrollment.data).mode==="SOURCE";
  var p=t.participant,er=t.enrollment,summary=parseJson_(er.data.AssessmentSummary,{}),patch={UpdatedAt:nowIso_()};
  ["PHASE 1","PHASE 2","PHASE 3"].forEach(function(ph,i){
    var a=readAssessmentDetail_(p.data.NIK,ph,p.data.Program,p.data.Batch),manualPhase3=ph==="PHASE 3"&&t.importSources&&t.importSources[ph]&&t.updatedBy!=="RAW_IMPORT";if(!a.available&&!manualPhase3&&!sourceDriven)return;
    if(!sourceDriven&&ph==="PHASE 3"&&!a.completedModules&&numberOrBlank_(er.data.Phase3Score)!==""&&!(t.importSources&&t.importSources[ph]&&t.updatedBy!=="RAW_IMPORT"))return;
    patch["Phase"+(i+1)+"Score"]=a.averageScore;
    var existing=summary[ph]||{};
    summary[ph]=Object.assign({},existing,{sources:Object.assign({},existing.sources||{},(t.importSources||{})[ph]||{}),remedials:a.modules.filter(function(m){return m.assessed!==false&&m.result==="REMEDIAL REQUIRED";}),completedModules:a.completedModules,totalModules:a.totalModules,updatedAt:nowIso_(),updatedBy:t.updatedBy||"SHEET_SYNC"});
  });
  if(sourceTransition){["BehaviourScore","BehaviourCoreScore","BehaviourTechnicalScore","BehaviourSummary","PanelScore","PanelSummary","LegacySnapshot"].forEach(function(k){patch[k]=er.data[k];});}
  applyParticipantFallbacks_(t,patch,summary);
  patch.AssessmentSummary=JSON.stringify(summary);
  var grade=finalGrade_(Object.assign({},er.data,patch),p.data,summary),workspace=workspace_(p.data);
  grade.missingRequirements=workspace.eligibility.missing;
  grade=resolveParticipantFinal_(t,patch,summary,grade);
  // Numeric result follows the workbook. Project/mentoring remains the panel and certificate gate.
  var status=grade.complete?(grade.score>=80?"LULUS":"TIDAK LULUS"):"ON GOING";
  patch.FinalScore=grade.score;patch.ProgramStatus=status;patch.FinalSummary=JSON.stringify(grade);
  if(grade.complete){patch.CurrentPhase="COMPLETED";patch.EnrollmentStatus="COMPLETED";patch.CompletedAt=er.data.CompletedAt||nowIso_();}
  else if(er.data.CurrentPhase==="COMPLETED"){patch.CurrentPhase="PANELISASI";patch.EnrollmentStatus="ON GOING";patch.CompletedAt="";}
  if(sourceDriven&&!grade.complete){var incomplete=["PHASE 1","PHASE 2","PHASE 3"].filter(function(ph){var x=summary[ph];return !x||!x.aggregateFallback&&(!x.totalModules||x.completedModules<x.totalModules);});patch.CurrentPhase=incomplete[0]||"PANELISASI";patch.EnrollmentStatus="ON GOING";patch.CompletedAt="";}
  var historicalLocked=protectLegacySummary_(t,patch,summary);
  var oldStatus=er.data.ProgramStatus,oldScore=numberOrBlank_(er.data.FinalScore);
  updateRowByIndex_(getSheet_(SHEET_ENROLLMENTS),getHeaders_(getSheet_(SHEET_ENROLLMENTS)),er.rowIndex,patch);
  Object.assign(er.data,patch);
  var pp={UpdatedAt:patch.UpdatedAt,FinalScore:patch.FinalScore,FinalStatus:patch.ProgramStatus};
  if(patch.BehaviourScore!==undefined)pp.BehaviourObservation=patch.BehaviourScore;
  ["Phase1Score","Phase2Score","Phase3Score","PanelScore","CurrentPhase"].forEach(function(k){if(patch[k]!==undefined)pp[k]=patch[k];});
  ["PHASE 1","PHASE 2","PHASE 3"].forEach(function(ph,i){var a=summary[ph];if(a)pp["Phase"+(i+1)+"Status"]=!a.completedModules?"NOT STARTED":a.remedials.length?"REMEDIAL":a.completedModules<a.totalModules?"PARTIAL":"PASS";});
  updateRowByIndex_(getSheet_(SHEET_PARTICIPANTS),getHeaders_(getSheet_(SHEET_PARTICIPANTS)),p.rowIndex,pp);Object.assign(p.data,pp);
  if(historicalLocked){return;}
  var cert=certificateFor_(p.data),certificateReady=certificateEligibility_(t,workspace).eligible;
  if(cert)updateRowByIndex_(getSheet_(SHEET_CERTIFICATES),getHeaders_(getSheet_(SHEET_CERTIFICATES)),cert.rowIndex,{FinalScore:patch.FinalScore,FinalStatus:status,Published:status==="LULUS"&&certificateReady&&oldScore===grade.score?cert.data.Published:false});
  if(status!=="LULUS"||!certificateReady||oldScore!==grade.score)cancelCertificateJobs_(er.data.EnrollmentID);
  if(status==="LULUS"&&certificateReady&&getSetting_("CERTIFICATE_TEMPLATE_ID")&&(oldStatus!==status||oldScore!==grade.score||!cert||!cert.data.FileOrLink))enqueueCertificate_(t,{nik:t.updatedBy||"AUTO_RECAP",role:"ADMIN"},{regenerate:!!(cert&&cert.data.FileOrLink&&oldScore!==grade.score)});
}

function refreshAssessmentSummaries_(token, ids) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  if (!Array.isArray(ids) || !ids.length || ids.length > 30)
    throw new Error("Pilih 1–30 enrollment per sinkronisasi.");
  return bulkWrites_(function () {
    var seen = {},
      targets = [];
    ids.forEach(function (id) {
      if (seen[id]) return;
      seen[id] = true;
      var er = enrollmentById_(id);
      if (!er) throw new Error("Enrollment tidak ditemukan.");
      targets.push(
        authorizeEnrollment_(s, { nik: er.data.NIK, enrollmentId: id }),
      );
    });
    targets.forEach(refreshEnrollmentSummary_);
    auditLog_(
      s.nik,
      s.role,
      "SYNC_ASSESSMENT_SUMMARIES",
      "Enrollment",
      "",
      targets.length + " enrollment",
    );
    return {
      ok: true,
      message:
        targets.length +
        " ringkasan nilai diperbarui dari sheet. Kelulusan otomatis dari nilai akhir lengkap, minimum 80.",
    };
  });
}
function refreshAssessmentSummaries(t, ids) {
  return api_("refreshAssessmentSummaries", function () {
    return refreshAssessmentSummaries_(t, ids);
  }, arguments);
}
// Optional full refresh from the editor; dashboard never calls this action.
function rebuildAssessmentSummaries_() {
  return api_("rebuildAssessmentSummaries", function () {
    var ids = readAllAsObjects_(SHEET_ENROLLMENTS).map(function (r) {
        return r.data.EnrollmentID;
      }),
      report = { updated: 0, conflicts: [] };
    ids.forEach(function (id) {
      try {
        bulkWrites_(function () {
          var er = enrollmentById_(id);
          if (!er) throw new Error("Enrollment tidak ditemukan");
          var p = participantTarget_(
            { role: "ADMIN" },
            { nik: er.data.NIK, enrollmentId: id },
          );
          refreshEnrollmentSummary_({ participant: p, enrollment: er });
        });
        report.updated++;
      } catch (e) {
        report.conflicts.push(id + ": " + e.message);
      }
    });
    console.log(JSON.stringify(report));
    return report;
  }, arguments);
}

// Raw syllabus / behavior imports. Source files remain separate from participant summaries.
var SHEET_IMPORTS_ = "Import Sessions";
var SHEET_IMPORT_MAP_ = "Import Module Mapping";
var SHEET_SCORE_EVENTS_ = "Assessment Import Records";
var SHEET_BEHAVIOR_ = "Behavior Observations";
SHEET_HEADERS[SHEET_IMPORTS_] = [
  "UploadID",
  "RequestID",
  "OwnerNIK",
  "Kind",
  "Program",
  "Batch",
  "Phase",
  "Aggregation",
  "FileName",
  "FileHash",
  "SourceFileID",
  "PayloadFileID",
  "Status",
  "CreatedAt",
  "ExpiresAt",
  "PreviewVersion",
  "NextIndex",
  "Written",
  "Skipped",
  "Failed",
  "ResultJSON",
  "LeaseToken",
  "LeaseUntil",
  "UpdatedAt",
];
SHEET_HEADERS[SHEET_IMPORT_MAP_] = [
  "Program",
  "Phase",
  "SourceKey",
  "SourceLabel",
  "Module",
  "UpdatedBy",
  "UpdatedAt",
];
SHEET_HEADERS[SHEET_SCORE_EVENTS_] = [
  "RecordID",
  "EnrollmentID",
  "NIK",
  "Program",
  "Batch",
  "Phase",
  "Module",
  "Score",
  "AnswersJSON",
  "ExamAt",
  "SourceLabel",
  "ContentHash",
  "UploadID",
  "SourceRow",
  "ImportedBy",
  "ImportedAt",
];
SHEET_HEADERS[SHEET_BEHAVIOR_] = [
  "ObservationID",
  "EnrollmentID",
  "NIK",
  "Program",
  "Batch",
  "ObservedAt",
  "ObserverNIK",
  "ObserverName",
  "Region",
  "CoreScore",
  "TechnicalScore",
  "AverageScore",
  "DetailsJSON",
  "Strengths",
  "ImprovementAreas",
  "Recommendations",
  "ContentHash",
  "UploadID",
  "SourceRow",
  "ImportedBy",
  "ImportedAt",
];
SHEET_HEADERS[SHEET_ENROLLMENTS] = SHEET_HEADERS[SHEET_ENROLLMENTS].concat([
  "BehaviourScore",
  "BehaviourCoreScore",
  "BehaviourTechnicalScore",
  "BehaviourSummary",
]);

function rawText_(v) {
  return String(v === null || v === undefined ? "" : v).trim();
}
function rawNumber_(v, min, max) {
  var s = rawText_(v);
  if (!s) return "";
  if (typeof v !== "number" && !/^[+-]?\d+(?:[.,]\d+)?$/.test(s))
    throw new Error("Angka tidak valid: " + s.slice(0, 80));
  var n = typeof v === "number" ? v : Number(s.replace(",", "."));
  if (!isFinite(n) || n < min || n > max)
    throw new Error("Nilai harus " + min + "–" + max + ": " + s.slice(0, 80));
  return n;
}
function rawNik_(v) {
  var s = rawText_(v);
  if (/^[0-9]+(?:\.[0-9]+)?e[+-]?[0-9]+$/i.test(s)) {
    var n = Number(s);
    if (!Number.isSafeInteger(n))
      throw new Error("NIK dalam notasi ilmiah tidak valid.");
    s = String(n);
  }
  return normalizeNik_(s);
}
function rawDate_(v) {
  if (v instanceof Date) return isFinite(v.getTime()) ? v.toISOString() : "";
  var s = rawText_(v);
  if (!s) return "";
  if (typeof v === "number" || /^\d{5}(?:\.\d+)?$/.test(s)) {
    var num = Number(v);
    if (num > 20000 && num < 80000)
      return new Date(
        Date.UTC(1899, 11, 30) + Math.round(num * 86400000) - 7 * 3600000,
      ).toISOString();
  }
  var m = s.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d\d)(?::(\d\d))?)?$/,
  );
  if (m)
    s =
      m[3] +
      "-" +
      ("0" + m[2]).slice(-2) +
      "-" +
      ("0" + m[1]).slice(-2) +
      "T" +
      ("0" + (m[4] || 0)).slice(-2) +
      ":" +
      (m[5] || "00") +
      ":" +
      (m[6] || "00") +
      "+07:00";
  else if (/^\d{4}-\d\d-\d\d(?:[ T]\d\d:\d\d(?::\d\d)?)?$/.test(s)) {
    s = s.replace(" ", "T");
    if (s.length === 10) s += "T00:00:00";
    if (s.length === 16) s += ":00";
    s += "+07:00";
  } else if (!/^\d{4}-\d\d-\d\dT.*(?:Z|[+-]\d\d:\d\d)$/.test(s)) return "";
  var date = new Date(s);
  return isFinite(date.getTime()) ? date.toISOString() : "";
}
function rawModuleKey_(v) {
  return normalizeModuleKey_(v)
    .replace(/\bphase\s*[123]\b/g," ")
    .replace(/\b(self learning|post test|posttest|pre test|pretest|exam)\b/g, " ")
    .replace(/\bai technology savvy\b/g,"technology savvy")
    .replace(/\bdelaer\b/g, "dealer")
    .replace(/\bothers\b/g, "other")
    .replace(/\s+/g, " ")
    .trim();
}

function suggestedModule_(program, phase, titles) {
  var cfg = rawAssessmentConfig_(program, phase),
    keys = titles.map(rawModuleKey_),
    extra = {
      "Business Process Analyst": [
        "Meningkatkan Efektifitas Proses Kerja",
        "Business Process & Analysis",
      ],
      "Sales Forecasting": ["Prediksi Penjualan dalam Bisnis Distribusi"],
      "Solution & Recovery": [
        "Recovery Dealer",
        "Solution and Recovery Dealer",
        "Solution and Recovery",
      ],
      "Market Research": ["Analisa Potensi Pasar", "Market Analysis"],
      "Negotiation Skill": ["Negotiation for Effective Closing"],
      "Monitor Your Dealer": ["Monitor Dealer"],
      "Monitor Dealer": ["Monitor Your Dealer"],
      "POST-TEST": ["Supervisory Development Program"],
      "Analisa Potensi Pasar": ["Market Research Analysis Planning"],
      "Developing Others": ["Developing Other"],
      "Leading Others": ["Leading Other"],
    },
    hits = [];
  cfg.modules.forEach(function (m) {
    var aliases = [m.label]
      .concat(m.headers, extra[m.label] || [])
      .map(rawModuleKey_);
    if (
      keys.some(function (k) {
        return k && aliases.indexOf(k) >= 0;
      })
    )
      hits.push(m.label);
  });
  return hits.length === 1 ? hits[0] : "";
}
function rawHeader_(rows, aliases) {
  for (var i = 0; i < Math.min(rows.length, 8); i++) {
    if (findHeaderIndex_(rows[i] || [], aliases) >= 0) return i;
  }
  throw new Error(
    "Kolom identitas peserta tidak ditemukan pada 8 baris pertama.",
  );
}
function parseRawSyllabus_(rows, program, phase) {
  var hr = rawHeader_(rows, ["NIK", "Employee ID"]),
    h = rows[hr],
    nikCol = findHeaderIndex_(h, ["NIK", "Employee ID"]),
    nameCol = findHeaderIndex_(h, [
      "Name",
      "Nama",
      "Nama Peserta",
      "Nama Karyawan",
      "Employee Name",
    ]),
    batchCol = findHeaderIndex_(h, ["Batch"]),
    columns = [],
    sub = -1;
  for (var r = hr + 1; r < Math.min(rows.length, hr + 4); r++)
    if (
      (rows[r] || []).some(function (v) {
        return headerKey_(v) === "score";
      })
    ) {
      sub = r;
      break;
    }
  if (sub >= 0) {
    var group = "",
      exam = "";
    for (var c = 0; c < rows[sub].length; c++) {
      if (rawText_(h[c])) group = rawText_(h[c]);
      if (sub > hr + 1 && rawText_((rows[sub - 1] || [])[c]))
        exam = rawText_(rows[sub - 1][c]);
      if (headerKey_(rows[sub][c]) !== "score") continue;
      var g = rawText_(h[c]) || group,
        x = rawText_((rows[sub - 1] || [])[c]) || exam || g,
        label = g + (x && x !== g ? " — " + x : "");
      columns.push({
        id: "col_" + c,
        col: c,
        dateCol: c + 2,
        sourceLabel: label,
        sourceKey: rawModuleKey_(g) + "|" + rawModuleKey_(x),
        suggested: suggestedModule_(program, phase, [x]) || suggestedModule_(program, phase, [g]),
      });
    }
  } else {
    var excluded = [
      "nik",
      "employeeid",
      "name",
      "nama",
      "namapeserta",
      "employeename",
      "batch",
      "program",
      "phase",
      "region",
      "jobtitle",
      "worklocation",
      "businessunit",
      "uniqueid",
      "averagescore", "finalscore", "finalstatus", "totalscore", "nilaitotal", "nilaiakhir",
      "modules",
      "elearningstatus", "namakaryawan", "namapenilai", "penilai", "orggroupname", "employeemail", "employeeemail",
      "locationcityname", "locationgroupname", "skorakhir", "status", "statuskelulusan", "jumlahterisi", "keterangan",
      "countphase2selflearning", "qtypengerjaan", "qtypengerjaanphase2", "avaragescorephase2selflearning",
    ];
    h.forEach(function (v, c) {
      if (rawText_(v) && excluded.indexOf(headerKey_(v)) < 0)
        columns.push({
          id: "col_" + c,
          col: c,
          dateCol: -1,
          sourceLabel: rawText_(v),
          sourceKey: rawModuleKey_(v),
          suggested: suggestedModule_(program, phase, [v]),
        });
    });
  }
  if (!columns.length)
    throw new Error("Kolom nilai / blok Score tidak ditemukan.");
  var data = [];
  for (var i = sub >= 0 ? sub + 1 : hr + 1; i < rows.length; i++) {
    var row = rows[i] || [];
    if (
      !row.some(function (v) {
        return rawText_(v) !== "";
      })
    )
      continue;
    var nik = rawNik_(row[nikCol]),
      name = rawText_(row[nameCol]);
    if (!nik && !name) {
      var populated = columns.filter(function(col) { return rawText_(row[col.col]) !== ""; });
      if (!populated.length) continue;
      if (program === "TSHDP" && phase === "PHASE 1" && i === hr + 1 && populated.length === 6 &&
          Math.abs(populated.reduce(function(sum,col) { return sum + Number(row[col.col]); },0) - 1) < 0.00001 &&
          populated.every(function(col) { return Number(row[col.col]) >= 0 && Number(row[col.col]) <= 1; })) continue;
    }
    data.push({
      rowNo: i + 1,
      nik: nik,
      name: name,
      batch: batchCol >= 0 ? rawText_(row[batchCol]) : "",
      cells: columns.map(function (col) {
        return {
          columnId: col.id,
          value: row[col.col] === undefined ? "" : row[col.col],
          examAt: col.dateCol >= 0 ? rawDate_(row[col.dateCol]) : "",
          invalidDate:
            col.dateCol >= 0 &&
            rawText_(row[col.dateCol]) !== "" &&
            !rawDate_(row[col.dateCol]),
        };
      }),
    });
  }
  return {
    format: sub >= 0 ? "SYLLABUS_EXPORT" : "SCORE_TABLE",
    columns: columns,
    rows: data,
  };
}
function parseRawBehavior_(rows) {
  var hr = rawHeader_(rows, ["NIK Peserta", "NIK", "Employee ID"]),
    h = rows[hr],
    at = function (a) {
      return findHeaderIndex_(h, a);
    },
    ni = at(["NIK Peserta", "NIK", "Employee ID"]),
    nm = at(["Nama Peserta", "Name", "Employee Name"]),
    ts = at(["Timestamp", "ObservedAt", "Tanggal Observasi"]),
    oi = at(["NIK Observer"]),
    on = at(["Nama Observer (Penilai)", "Observer Name"]),
    re = at(["Region"]),
    co = at(["Score Core Competency", "Score Core Comp"]),
    te = at(["Score Technical Competency", "Score Technical Competcny", "Score Tech Comp"]),
    av = at(["Average Score", "Total Score"]),
    st = at(["Kekuatan Utama"]),
    im = at(["Area Perbaikan"]),
    rc = at(["Rekomendasi Pengembangan"]),
    choices = [],
    ratings = [];
  if (ts < 0 || co < 0 || te < 0)
    throw new Error(
      "Gunakan sheet Form Responses 1 atau Recap dengan Timestamp, Score Core Competency, dan Score Technical Competency.",
    );
  h.forEach(function (v, i) {
    if (/^(pilihnamapeserta|namapeserta?region|namapesetaregion)/.test(headerKey_(v))) choices.push(i);
    if (/^ratingkompetensi/.test(headerKey_(v))) ratings.push(i);
  });
  var data = [];
  for (var i = hr + 1; i < rows.length; i++) {
    var row = rows[i] || [],
      nik = rawNik_(row[ni]),
      name = rawText_(row[nm]),
      date = rawDate_(row[ts]),
      observer = rawText_(row[on]),
      selected = choices
        .map(function (c) {
          return rawText_(row[c]);
        })
        .filter(Boolean);
    if (!nik && !name && !rawText_(row[ts]) && !observer && !selected.length)
      continue;
    var issue = "",
      candidate = [];
    selected.forEach(function (v) {
      var m = v.match(/^(\d{4,18})(?:\.0)?\s+(.+)$/);
      if (m) candidate.push(m);
    });
    if (!nik && candidate.length === 1) {
      nik = rawNik_(candidate[0][1]);
      if (!name) name = candidate[0][2];
    }
    if (
      candidate.some(function (m) {
        return nik && rawNik_(m[1]) !== nik;
      })
    )
      issue = "NIK Peserta berbeda dari NIK pada pilihan peserta.";
    var details = { ratings: {}, checks: {} };
    h.forEach(function (label, c) {
      var value = row[c];
      if (ratings.indexOf(c) >= 0 && rawText_(value))
        details.ratings[rawText_(label)] = value;
      else if (["ya", "tidak"].indexOf(rawText_(value).toLowerCase()) >= 0)
        details.checks[rawText_(label)] = rawText_(value);
    });
    if (te >= 0 && co >= 0)
      for (var col = co + 1; col < te; col++)
        if (rawText_(row[col]) && /^\d+(?:[.,]\d+)?$/.test(rawText_(row[col])))
          details.ratings[rawText_(h[col])] = row[col];
    var coreValues=ratings.map(function(c){return rawNumber_(row[c],1,5);}).filter(function(v){return v!=="";}),technicalValues=[];
    if(co>=0&&te>co)for(var c=co+1;c<te;c++)technicalValues.push(rawNumber_(row[c],1,5));
    var mean=function(v){return v.reduce(function(a,b){return a+b;},0)/v.length;},core=coreValues.length===9?mean(coreValues):row[co],tech=technicalValues.length===9&&technicalValues.every(function(v){return v!=="";})?mean(technicalValues):row[te];
    data.push({
      rowNo: i + 1,
      nik: nik,
      name: name,
      observedAt: date,
      rawTimestamp: rawText_(row[ts]),
      observerNik: rawNik_(row[oi]),
      observerName: observer,
      region: rawText_(row[re]),
      core: core === undefined ? "" : core,
      technical: tech === undefined ? "" : tech,
      average: numberOrBlank_(core)!==""&&numberOrBlank_(tech)!==""?(Number(core)+Number(tech))/2:(av>=0?row[av]:""),
      details: details,
      strengths: rawText_(row[st]),
      improvementAreas: rawText_(row[im]),
      recommendations: rawText_(row[rc]),
      issue: issue,
    });
  }
  return { format: "BEHAVIOR_1_TO_5", columns: [], rows: data };
}
function parseRawWorkbook_(d) {
  var name = rawText_(d.fileName).toLowerCase();
  if (/\.csv$/.test(name))
    return {
      sheetName: "CSV",
      rows: parseTabularBase64_(d.fileName, "text/csv", d.base64),
    };
  if (!/\.(xlsx|xls)$/.test(name))
    throw new Error("Gunakan CSV, XLSX, atau XLS.");
  var bytes = Utilities.base64Decode(rawText_(d.base64).split("base64,").pop()),
    id = null;
  try {
    id = convertBlobToSheet_(
      Utilities.newBlob(
        bytes,
        d.mimeType || MimeType.MICROSOFT_EXCEL,
        d.fileName,
      ),
    );
    var sheets = SpreadsheetApp.openById(id).getSheets(),
      sheet = sheets[0];
    if (d.kind === "SYLLABUS") {
      var phaseKey = headerKey_(normalizeCurrentPhase_(d.phase || "PHASE 2"));
      var programKey = headerKey_(d.program);
      sheet = sheets.filter(function(x) { return headerKey_(x.getName()) === phaseKey + programKey; })[0] ||
        sheets.filter(function(x) { return headerKey_(x.getName()) === phaseKey; })[0] || sheet;
    }
    if (d.kind === "SYLLABUS" && d.program === "TSHDP" && normalizeCurrentPhase_(d.phase) === "PHASE 1")
      sheet=sheets.filter(function(s){return headerKey_(s.getName())==="rekapnilai";})[0]||sheet;
    if (d.kind === "BEHAVIOR" || d.kind === "PANEL")
      sheet =
        sheets.filter(function (s) {
          return /^form responses/i.test(s.getName());
        })[0] ||
        sheets.filter(function (s) {
          return /^recap$/i.test(s.getName());
        })[0] ||
        sheets.filter(function (s) { return /^bm$/i.test(s.getName()); })[0] ||
        sheet;
    var nr = sheet.getLastRow(),
      nc = sheet.getLastColumn();
    if (nr > 5003 || nc > 250)
      throw new Error("Maksimal 5.000 baris data dan 250 kolom per file.");
    return {
      sheetName: sheet.getName(),
      rows: nr && nc ? sheet.getRange(1, 1, nr, nc).getValues() : [],
    };
  } finally {
    if (id)
      try {
        DriveApp.getFileById(id).setTrashed(true);
      } catch (error) {}
  }
}
function getRawImportOptions_(token) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var optionsKey = cacheKey_("ldprio:", LDP_VERSION_ + "|" + activeSegment_() + "|" + revision_()), cachedOptions = readSnapshot_(optionsKey);
  if (cachedOptions) return cachedOptions;
  getSheet_(SHEET_IMPORTS_);
  var programs = {};
  activePrograms_().forEach(function (p) {
    programs[p] = { batches: ["AUTO"], phases: {} };
    ["PHASE 1", "PHASE 2"].forEach(function (ph) {
      programs[p].phases[ph] = getAssessmentConfig_(p, ph).modules.map(
        function (m) {
          return { module: m.label, weight: m.weight === undefined ? null : m.weight, url: qleapUrl_(p, ph, m.label) };
        },
      );
    });
    programs[p].phases["PHASE 3"] = [{module:"POST-TEST", url:qleapUrl_(p,"PHASE 3","POST-TEST"), weight:null}];
  });
  readAllAsObjects_(SHEET_ENROLLMENTS).forEach(function (r) {
    var e = r.data,
      p = programs[e.Program];
    if (
      p &&
      !p.batches.some(function (b) {
        return sameBatch_(b, e.Batch);
      })
    )
      p.batches.push(String(e.Batch));
  });
  // A pasted participant roster can supply batches before its enrollment backfill.
  readAllAsObjects_(SHEET_PARTICIPANTS).forEach(function(r) {
    var d = r.data, program = programs[String(d.Program || "").toUpperCase()];
    if (program && normalizeNik_(d.NIK) && normalizeBatchKey_(d.Batch) && !program.batches.some(function(b){return sameBatch_(b,d.Batch);})) program.batches.push(String(d.Batch));
  });
  Object.keys(programs).forEach(function (p) {
    programs[p].batches.sort(function (a, b) {
      if(a==="AUTO")return -1;if(b==="AUTO")return 1;return a.localeCompare(b, undefined, { numeric: true });
    });
  });
  var result = {programs:programs, maxFileMB:Math.min(5, Number(getSetting_("MAX_FILE_MB")) || 20), maxParticipants:250, behaviorScale:5,sourceDefaults:activeSegment_()==="B2C"?{}:RAW_SOURCE_DEFAULTS_};
  putSnapshot_(optionsKey, result);
  return result;
}
function getAssessmentImportTemplate_(token, d) {
  var session = requireSession_(token); requireRole_(session, ["ADMIN"]); d = d || {};
  var program = String(d.program || "").toUpperCase(), phase = normalizeCurrentPhase_(d.phase), cfg = getAssessmentConfig_(program, phase);
  if (!cfg) throw new Error("Pilih program dan Phase 1 atau Phase 2.");
  var batch = String(d.batch || "").trim();
  if (!batch) throw new Error("Pilih batch terlebih dahulu.");
  var header = ["NIK", "Name", "Batch", "Nama Penilai"].concat(cfg.modules.map(function(m) { return m.label; }));
  var source = scopedParticipants_(session).filter(function(r) { return String(r.data.Program).toUpperCase() === program && (batch==="AUTO"||sameBatch_(r.data.Batch,batch)); });
  var rows = [header], seen = {};
  source.forEach(function(r) {
    var p = r.data, nik = normalizeNik_(p.NIK); var identity=entityKey_(nik,program,p.Batch, p.Segment);if(seen[identity])throw new Error("Peserta duplikat: "+identity);seen[identity]=true;
    rows.push([nik,p.Name,p.Batch,""].concat(cfg.modules.map(function(){return "";})));
  });
  return {filename:"Template_" + program + "_" + phase.replace(/ /g,"_") + "_" + sanitizeFileName_(batch) + ".csv",
    content:"\ufeff" + rows.map(function(row){return row.map(csvEscape_).join(",");}).join("\r\n")};
}
function getAssessmentImportTemplate(token, d) { return api_("getAssessmentImportTemplate", function(){return getAssessmentImportTemplate_(token,d);}, arguments); }

function importTargetIndex_() {
  var map = {};
  readAllAsObjects_(SHEET_PARTICIPANTS).forEach(function (r) {
    var k = entityKey_(r.data.NIK, r.data.Program, r.data.Batch, r.data.Segment);
    if (map[k]) throw new Error("Participants duplikat: " + k);
    map[k] = r;
  });
  return map;
}
function behaviorAggregate_(records, policy) {
  var unique = {},
    list = [];
  records.forEach(function (d) {
    if (d.ContentHash) unique[d.ContentHash] = d;
  });
  Object.keys(unique).forEach(function (k) {
    list.push(unique[k]);
  });
  list.sort(function (a, b) {
    return (
      timeValue_(b.ObservedAt) - timeValue_(a.ObservedAt) ||
      String(b.ContentHash || b.ObservationID).localeCompare(
        String(a.ContentHash || a.ObservationID),
      )
    );
  });
  var chosen = list;
  if (policy === "LATEST") chosen = list.slice(0, 1);
  if (policy === "RATER_LATEST") {
    var raters = {};
    chosen = [];
    list.forEach(function (d) {
      var k =
        rawText_(d.ObserverNIK) ||
        rawText_(d.ObserverName).toLowerCase() ||
        "UNKNOWN";
      if (!raters[k]) {
        raters[k] = true;
        chosen.push(d);
      }
    });
  }
  function avg(k) {
    var v = chosen
      .map(function (d) {
        return numberOrBlank_(d[k]);
      })
      .filter(function (x) {
        return x !== "";
      });
    return v.length
      ? Math.round(
          (v.reduce(function (a, b) {
            return a + b;
          }, 0) /
            v.length) *
            10000,
        ) / 10000
      : "";
  }
  return {
    score: avg("AverageScore"),
    core: avg("CoreScore"),
    technical: avg("TechnicalScore"),
    observations: list.length,
    usedObservations: chosen.length,
    policy: policy,
    scale: 5,
    lastObservedAt: list.length ? list[0].ObservedAt : "",
    updatedAt: nowIso_(),
  };
}
function behaviorData_(r, context, enrollmentId) {
  var core = rawNumber_(r.core, 1, 5),
    technical = rawNumber_(r.technical, 1, 5);
  if (core === "" || technical === "")
    throw new Error("Score Core dan Technical wajib terisi (skala 1–5).");
  var average = rawNumber_(r.average, 1, 5);
  if (average === "") average = (core + technical) / 2;
  if (!r.observedAt)
    throw new Error("Timestamp observasi kosong atau tidak dikenali.");
  if (r.issue) throw new Error(r.issue);
  var obj = {
    EnrollmentID: enrollmentId,
    NIK: r.nik,
    Program: context.program,
    Batch: context.batch,
    ObservedAt: r.observedAt,
    ObserverNIK: r.observerNik,
    ObserverName: r.observerName,
    Region: r.region,
    CoreScore: core,
    TechnicalScore: technical,
    AverageScore: average,
    DetailsJSON: JSON.stringify(r.details),
    Strengths: r.strengths,
    ImprovementAreas: r.improvementAreas,
    Recommendations: r.recommendations,
    SourceRow: r.rowNo,
  };
  obj.ContentHash = cacheKey_(
    "",
    JSON.stringify([
      obj.EnrollmentID,
      obj.ObservedAt,
      obj.ObserverNIK,
      obj.ObserverName,
      core,
      technical,
      average,
      obj.DetailsJSON,
      obj.Strengths,
      obj.ImprovementAreas,
      obj.Recommendations,
    ]),
  );
  return obj;
}
function buildRawPreview_(raw, scope, requestedMappings, s) {
  var cfg =
      scope.kind === "SYLLABUS"
        ? rawAssessmentConfig_(scope.program, scope.phase)
        : null,
    allowed = cfg
      ? cfg.modules.map(function (m) {
          return m.label;
        })
      : [],
    saved = {};
  if (cfg)
    readAllAsObjects_(SHEET_IMPORT_MAP_).forEach(function (r) {
      var d = r.data;
      if (
        d.Program === scope.program &&
        normalizeCurrentPhase_(d.Phase) === scope.phase
      )
        saved[d.SourceKey] = d.Module;
    });
  var mappings = {},
    columns = (raw.columns || []).map(function (c) {
      var choice =
        requestedMappings &&
        Object.prototype.hasOwnProperty.call(requestedMappings, c.id)
          ? requestedMappings[c.id]
          : saved[c.sourceKey] || c.suggested || "";
      if (choice && choice !== "IGNORE" && allowed.indexOf(choice) < 0)
        throw new Error("Kompetensi tujuan tidak dikenal: " + choice);
      mappings[c.id] = choice;
      return Object.assign({}, c, { module: choice });
    });
  var colMap = {};
  columns.forEach(function (c) {
    colMap[c.id] = c;
  });
  var participants = importTargetIndex_(),
    ers = enrollmentIndex_(),
    items = {},
    errors = [],
    warnings = [],
    blankRows = 0,
    totalCells = 0;
  function err(row, msg) {
    errors.push({ rowNo: row.rowNo, nik: row.nik, message: msg });
  }
  var workingRows=(raw.rows||[]).map(function(r){return Object.assign({},r);});
  if(scope.kind==="PANEL")resolveRawPanelNames_({rows:workingRows},scope,participants);
  workingRows.forEach(function (r) {
    if(r.issue){err(r,r.issue);return;}
    if (!r.nik) {
      err(
        r,
        "NIK Peserta kosong. Gunakan NIK atau pilihan peserta yang diawali NIK.",
      );
      return;
    }
    var rowBatch;try{rowBatch=rawTargetBatch_(r,scope,participants);}catch(error){err(r,error.message);return;}
    var key = entityKey_(r.nik, scope.program, rowBatch, r.segment),
      p = participants[key],
      er = ers.byKey[key];
    if (!p || !er) {
      err(
        r,
        "NIK belum terdaftar pada " +
          scope.program +
          " batch " +
          scope.batch +
          ". Peserta tidak dibuat otomatis.",
      );
      return;
    }
    if (r.batch && !assessmentBatchCompatible_(r.batch, rowBatch)) {
      err(r, "Batch pada file berbeda dari batch yang dipilih.");
      return;
    }
    if (
      r.name &&
      normalizeModuleKey_(r.name) !== normalizeModuleKey_(p.data.Name)
    )
      warnings.push({
        rowNo: r.rowNo,
        nik: r.nik,
        message:
          "Nama file berbeda dari nama database; pencocokan tetap memakai NIK.",
      });
    if (!items[key])
      items[key] = {
        enrollmentId: er.data.EnrollmentID,
        nik: r.nik,
        name: p.data.Name,
        program: scope.program,
        batch: er.data.Batch,
        phase: scope.phase,
        expectedVersion: er.data.RecordVersion || "",
        sourceRows: [],
        records: [],
        preTests: [],
        scores: {},
        sources: {},
        currentScore:
          scope.kind === "BEHAVIOR" || scope.kind === "PANEL"
            ? numberOrBlank_(er.data[scope.kind === "PANEL" ? "PanelScore" : "BehaviourScore"])
            : numberOrBlank_(
                er.data[
                  scope.phase === "PHASE 1" ? "Phase1Score" : scope.phase === "PHASE 2" ? "Phase2Score" : "Phase3Score"
                ],
              ),
      };
    var item = items[key];
    item.sourceRows.push(r.rowNo);
    if (scope.kind === "BEHAVIOR" || scope.kind === "PANEL") {
      try {
        item.records.push(scope.kind === "PANEL" ? panelData_(r,Object.assign({},scope,{batch:er.data.Batch}),er.data.EnrollmentID) : behaviorData_(r,Object.assign({},scope,{batch:er.data.Batch}),er.data.EnrollmentID));
      } catch (error) {
        err(r, error.message);
      }
      return;
    }
    var records = [], invalid = false;
    if(scope.phase==="PHASE 3"&&rawText_(r.preTest)!=="") {
      try {var pre=rawNumber_(r.preTest,0,100);if(!r.testAt)throw new Error("Tanggal pre-test tidak dikenali.");
        item.preTests.push({score:pre,at:r.testAt});
        var pr={EnrollmentID:er.data.EnrollmentID,NIK:r.nik,Program:scope.program,Batch:er.data.Batch,Phase:scope.phase,Module:"PRE-TEST",Score:pre,ExamAt:r.testAt,SourceLabel:"PRE-TEST",SourceRow:r.rowNo,AnswersJSON:JSON.stringify(r.answers||{})};pr.ContentHash=cacheKey_("",JSON.stringify([pr.EnrollmentID,pr.Phase,pr.Module,pr.Score,pr.ExamAt,pr.AnswersJSON]));item.records.push(pr);
      }catch(error){err(r,error.message);return;}
    }
    r.cells.forEach(function (cell) {
      var col = colMap[cell.columnId],
        module = mappings[cell.columnId];
      if (!module || module === "IGNORE" || rawText_(cell.value) === "") return;
      try {
        if (cell.invalidDate) throw new Error("Tanggal ujian tidak dikenali.");
        var score = rawNumber_(cell.value, 0, 100);
        if (score === "") return;
        var data = {
          EnrollmentID: er.data.EnrollmentID,
          NIK: r.nik,
          Program: scope.program,
          Batch: er.data.Batch,
          Phase: scope.phase,
          Module: module,
          Score: score,
          AnswersJSON: JSON.stringify(r.answers||{}),
          ExamAt: cell.examAt || "",
          SourceLabel: col.sourceLabel,
          SourceRow: r.rowNo,
        };
        data.ContentHash = cacheKey_(
          "",
          JSON.stringify([
            data.EnrollmentID,
            data.Phase,
            module,
            data.ExamAt,
            score,
            col.sourceKey,
          ]),
        );
        records.push(data);
      } catch (error) {
        invalid = true;
        err(r, col.sourceLabel + ": " + error.message);
      }
    });
    if (invalid) return;
    if (!records.length) {
      blankRows++;
      return;
    }
    item.records = item.records.concat(records);
  });
  var behaviorIndex = {};
  if (scope.kind === "BEHAVIOR" || scope.kind === "PANEL")
    readAllAsObjects_(scope.kind === "PANEL" ? SHEET_PANEL_OBSERVATIONS_ : SHEET_BEHAVIOR_).forEach(function (r) {
      var d = r.data;
      (
        behaviorIndex[d.EnrollmentID] || (behaviorIndex[d.EnrollmentID] = [])
      ).push(d);
    });
  var result = [];
  Object.keys(items).forEach(function (key) {
    var item = items[key];
    if (!item.records.length) return;
    if (scope.kind === "BEHAVIOR" || scope.kind === "PANEL") {
      item.preview = behaviorAggregate_(
        (behaviorIndex[item.enrollmentId] || []).concat(item.records),
        scope.aggregation,
      );
      result.push(item);
      return;
    }
    var expectedScores = {};
    readAssessmentDetail_(
      item.nik,
      scope.phase,
      scope.program,
      item.batch,
    ).modules.forEach(function (m) {
      expectedScores[m.module] = m.score;
    });
    var candidates = {};
    item.records.forEach(function(d){if(allowed.indexOf(d.Module)<0)return;var old=candidates[d.Module];if(!old||d.Score>old.Score||(d.Score===old.Score&&timeValue_(d.ExamAt)>timeValue_(old.ExamAt)))candidates[d.Module]=d;});
    var e=ers.byId[item.enrollmentId].data,previousSources=(parseJson_(e.AssessmentSummary,{})[scope.phase]||{}).sources||{};
    Object.keys(candidates).forEach(function(module){
      var d=candidates[module],previous=previousSources[module],current=expectedScores[module];
      if(previous&&/^MANUAL:/.test(previous.hash||"")){warnings.push({rowNo:d.SourceRow,nik:item.nik,message:module+": koreksi Admin dipertahankan; percobaan masuk riwayat."});return;}
      if(current!==undefined&&current!==""&&current>=d.Score)return;
      item.scores[module]=d.Score;item.sources[module]={examAt:d.ExamAt,hash:d.ContentHash,sourceLabel:d.SourceLabel,sourceRow:d.SourceRow,policy:"HIGHEST"};totalCells++;
    });
    item.expectedScores = expectedScores;
    item.preview = {
      modules: Object.keys(item.scores).length,
      sourceRecords: item.records.length,
      finalScore: assessmentAverage_(cfg, cfg.modules.map(function(m) { return {score:Object.prototype.hasOwnProperty.call(item.scores,m.label) ? item.scores[m.label] : (expectedScores[m.label] === undefined ? "" : expectedScores[m.label])}; })),
      weighted: scope.program === "TSHDP" && scope.phase === "PHASE 1",
    };
    result.push(item);
  });
  if (result.length > 250)
    throw new Error("Maksimal 250 peserta per upload. Pisahkan file.");
  return {
    items: result,
    mappings: mappings,
    columns: columns,
    summary: {
      segment:activeSegment_(),
      kind: scope.kind,
      format: raw.format,
      program: scope.program,
      batch: scope.batch,
      phase: scope.phase,
      aggregation: scope.aggregation,
      sourceRows: raw.rows.length,
      matchedParticipants: result.length,
      scoreCells: totalCells,
      blankRows: blankRows,
      errorCount: errors.length,
      warningCount: warnings.length,
      unmappedColumns: columns.filter(function (c) {
        return !c.module;
      }).length,
    },
    errors: errors,
    warnings: warnings,
  };
}
function importSession_(id) {
  return uniqueMatch_(
    readAllAsObjects_(SHEET_IMPORTS_).filter(function (r) {
      return r.data.UploadID === id;
    }),
    "import session",
  );
}
function requireImport_(s, id) {
  var r = importSession_(id);
  if (!r || r.data.OwnerNIK !== s.nik)
    throw new Error("Upload tidak ditemukan atau milik Admin lain.");
  return r;
}
function importPayload_(id) {
  var key = cacheKey_("ldp-import-payload:", id),
    cached = readSnapshot_(key);
  if (cached) return cached;
  var parsed = JSON.parse(
    DriveApp.getFileById(id).getBlob().getDataAsString("UTF-8"),
  );
  putSnapshot_(key, parsed);
  return parsed;
}
function importFolder_() {
  var key = "LDP_IMPORT_FOLDER_ID",
    id = props_().getProperty(key);
  if (id) return DriveApp.getFolderById(id);
  var folder = DriveApp.createFolder("LDP Import Sources");
  props_().setProperty(key, folder.getId());
  return folder;
}
function saveImportPayload_(id, payload) {
  var folder = importFolder_();
  return folder
    .createFile(
      Utilities.newBlob(
        JSON.stringify(payload),
        "application/json",
        id + "_preview_" + Utilities.getUuid() + ".json",
      ),
    )
    .getId();
}
function previewResponse_(job, payload) {
  var p = payload.preview;
  return {
    uploadId: job.UploadID,
    status: job.Status,
    previewVersion: Number(job.PreviewVersion || 1),
    expiresAt: job.ExpiresAt,
    sheetName: payload.sheetName,
    summary: p.summary,
    columns: p.columns,
    errors: p.errors.slice(0, 100),
    warnings: p.warnings.slice(0, 100),
    rows: p.items.slice(0, 50).map(function (i) {
      return {
        nik: i.nik,
        name: i.name,
        batch: i.batch,
        enrollmentId: i.enrollmentId,
        currentScore: i.currentScore,
        scores: i.scores,
        preview: i.preview,
        sourceRows: i.sourceRows,
      };
    }),
    totalItems: p.items.length,
  };
}
function ensureImportEnrollments_(program, batch, targetNik, session) {
  if (activePrograms_().indexOf(program) < 0 || !normalizeBatchKey_(batch)) throw new Error("Program/batch tujuan tidak valid.");
  return withLock_(function() { var result = bulkWrites_(function() {
    // bulkWrites_ reloads live rows under the script lock, preventing duplicate creation.
    if (session) participantTarget_(session, {nik:targetNik, program:program, batch:batch});
    var participants = readAllAsObjects_(SHEET_PARTICIPANTS).filter(function(r){return String(r.data.Program).toUpperCase() === program && sameBatch_(r.data.Batch,batch) && normalizeNik_(r.data.NIK) && (!targetNik || normalizeNik_(r.data.NIK) === normalizeNik_(targetNik));});
    var seen = {}, existing = {};
    participants.forEach(function(r){var k=entityKey_(r.data.NIK,program,batch, r.data.Segment);if(seen[k])throw new Error("Participants duplikat: " + k);seen[k]=true;});
    readAllAsObjects_(SHEET_ENROLLMENTS).forEach(function(r){var d=r.data,k=entityKey_(d.NIK,d.Program,d.Batch, d.Segment);if(seen[k]){if(existing[k])throw new Error("Enrollment duplikat: " + k);existing[k]=r;}});
    var es=getSheet_(SHEET_ENROLLMENTS),eh=getHeaders_(es),ps=getSheet_(SHEET_PARTICIPANTS),ph=getHeaders_(ps),created=0,linked=0;
    participants.forEach(function(r){
      var d=r.data,k=entityKey_(d.NIK,program,batch, d.Segment),er=existing[k],id=er && er.data.EnrollmentID;
      if(!er){id=newId_("ENR");var value={Segment:d.Segment,EnrollmentID:id,NIK:normalizeNik_(d.NIK),Program:program,Batch:d.Batch,Region:d.Region,CurrentPhase:normalizeCurrentPhase_(d.CurrentPhase||inferCurrentPhase_(d)),EnrollmentStatus:normalizeCurrentPhase_(d.CurrentPhase||inferCurrentPhase_(d))==="COMPLETED"?"COMPLETED":"ON GOING",ProgramStatus:d.FinalStatus||"ON GOING",IsActive:false,Phase1Score:d.Phase1Score,Phase2Score:d.Phase2Score,Phase3Score:d.Phase3Score,FinalScore:d.FinalScore,BehaviourScore:d.BehaviourObservation,CertificateUrl:d.CertificateUrl||"",CreatedAt:nowIso_(),UpdatedAt:nowIso_()};
        appendObject_(es,eh,value);created++;
      }else if(!id){id=newId_("ENR");updateRowByIndex_(es,eh,er.rowIndex,{EnrollmentID:id,UpdatedAt:nowIso_()});}
      if(d.CurrentEnrollmentId!==id){updateRowByIndex_(ps,ph,r.rowIndex,{CurrentEnrollmentId:id});linked++;}
    });
    return {created:created,linked:linked};
  }); if(result.created || result.linked) SpreadsheetApp.flush(); return result; });
}
function prepareRawImport_(token, d) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  d = d || {};
  requireFields_(d,["kind","program","batch","requestId"]);
  if(d.sourceUrl){sourceSpreadsheetId_(d.sourceUrl);d=Object.assign({},d,{fileName:"Sumber_"+d.program+"_"+(d.phase||d.kind)+".json"});}
  else {requireFields_(d,["fileName","base64"]);validateImportSize_(d);}
  var sourceFingerprint=d.sourceUrl?"SHEET:"+rawText_(d.sourceUrl):d.base64;
  if ((rawText_(d.base64).length * 3) / 4 > 5 * 1024 * 1024)
    throw new Error("File import maksimal 5 MB.");
  var scope = {
    segment:activeSegment_(),
    kind: String(d.kind).toUpperCase(),
    program: String(d.program).toUpperCase(),
    batch: rawText_(d.batch),
    phase: normalizeCurrentPhase_(d.phase || "PHASE 2"),
    aggregation: d.aggregation || (String(d.kind).toUpperCase()==="PANEL"?"ALL_MEAN":"LATEST"),
  };
  if (
    ["SYLLABUS", "BEHAVIOR", "PANEL"].indexOf(scope.kind) < 0 ||
    activePrograms_().indexOf(scope.program) < 0
  )
    throw new Error("Jenis / program tidak valid.");
  if (
    scope.kind === "SYLLABUS" &&
    (["PHASE 1", "PHASE 2", "PHASE 3"].indexOf(scope.phase) < 0 ||
      !rawAssessmentConfig_(scope.program, scope.phase))
  )
    throw new Error("Pilih Phase 1–3.");
  if (["LATEST", "RATER_LATEST", "ALL_MEAN"].indexOf(scope.aggregation) < 0)
    throw new Error("Cara rekap behavior tidak valid.");
  if (!/^[A-Za-z0-9_-]{12,100}$/.test(d.requestId))
    throw new Error("Identitas upload tidak valid.");
  // This write action links existing roster rows only; it never creates a new participant.
  ensureAutoRecapSchema_();
  ensureRawImportEnrollments_(scope.program, scope.batch);
  var reserved = bulkWrites_(function () {
    requireSession_(token);
    var existing = uniqueMatch_(
      readAllAsObjects_(SHEET_IMPORTS_).filter(function (r) {
        return r.data.OwnerNIK === s.nik && r.data.RequestID === d.requestId;
      }),
      "request upload",
    );
    if (existing) {
      if (
        existing.data.FileHash !== cacheKey_("", sourceFingerprint) ||
        existing.data.Kind !== scope.kind ||
        existing.data.Program !== scope.program ||
        !sameBatch_(existing.data.Batch, scope.batch) ||
        existing.data.Phase !== (scope.kind === "SYLLABUS" ? scope.phase : "")
      )
        throw new Error("RequestID sudah dipakai file / tujuan lain.");
      if (existing.data.Status === "PREVIEW")
        return { existing: existing.data };
      throw new Error(
        "Upload ini sudah diproses. Mulai upload baru atau buka riwayatnya.",
      );
    }
    var id = newId_("IMP"),
      now = nowIso_(),
      row = appendObject_(
        getSheet_(SHEET_IMPORTS_),
        getHeaders_(getSheet_(SHEET_IMPORTS_)),
        {
          UploadID: id,
          RequestID: d.requestId,
          OwnerNIK: s.nik,
          Kind: scope.kind,
          Program: scope.program,
          Batch: scope.batch,
          Phase: scope.kind === "SYLLABUS" ? scope.phase : "",
          Aggregation: scope.aggregation,
          FileName: sanitizeFileName_(d.fileName),
          FileHash: cacheKey_("", sourceFingerprint),
          Status: "PREPARING",
          CreatedAt: now,
          ExpiresAt: new Date(Date.now() + 1800000).toISOString(),
          PreviewVersion: 1,
          NextIndex: 0,
          Written: 0,
          Skipped: 0,
          Failed: 0,
          LeaseUntil: Date.now() + 330000,
          UpdatedAt: now,
        },
      );
    return { id: id, row: row };
  });
  if (reserved.existing) {
    if (timeValue_(reserved.existing.ExpiresAt) < Date.now())
      throw new Error("Preview sudah kedaluwarsa. Analisis file kembali.");
    return previewResponse_(
      reserved.existing,
      importPayload_(reserved.existing.PayloadFileID),
    );
  }
  try {
    var parsed = (d.sourceUrl?readRawSourceWorkbook_(Object.assign({},d,scope)):parseRawWorkbook_(Object.assign({},d,scope)));
    if (!parsed.rows.length) throw new Error("File kosong.");
    var raw =
      scope.kind === "SYLLABUS"
        ? (scope.phase === "PHASE 3" ? parseRawPhase3_(parsed.rows,scope.program) : parseRawSyllabus_(parsed.rows, scope.program, scope.phase))
        : scope.kind === "PANEL" ? parseRawPanel_(parsed.rows,scope.program) : parseRawBehavior_(parsed.rows);
    if (raw.rows.length > (scope.kind === "SYLLABUS" ? 500 : 2000))
      throw new Error(
        "Maksimal 500 baris syllabus atau 2.000 observasi per file.",
      );
    var preview = buildRawPreview_(raw, scope, null, s),
      payload = {
        scope: scope,
        raw: raw,
        preview: preview,
        sheetName: parsed.sheetName,
        sourceUrl: d.sourceUrl||"",
      },
      payloadID = saveImportPayload_(reserved.id, payload),
      folder = importFolder_(),
      source = folder.createFile(
        Utilities.newBlob(
          d.sourceUrl?JSON.stringify({sourceUrl:d.sourceUrl,sheetName:parsed.sheetName,rows:parsed.rows}):Utilities.base64Decode(rawText_(d.base64).split("base64,").pop()),
          d.sourceUrl?"application/json":d.mimeType || "application/octet-stream",
          reserved.id + "_" + sanitizeFileName_(d.fileName),
        ),
      );
    var job = bulkWrites_(function () {
      requireSession_(token);
      var r = requireImport_(s, reserved.id);
      updateRowByIndex_(
        getSheet_(SHEET_IMPORTS_),
        getHeaders_(getSheet_(SHEET_IMPORTS_)),
        r.rowIndex,
        {
          PayloadFileID: payloadID,
          SourceFileID: source.getId(),
          Status: "PREVIEW",
          LeaseUntil: "",
          ResultJSON: JSON.stringify({
            preview: preview.summary,
            totalItems: preview.items.length,
            outcomes: [],
          }),
          UpdatedAt: nowIso_(),
        },
      );
      return r.data;
    });
    return previewResponse_(job, payload);
  } catch (error) {
    bulkWrites_(function () {
      var r = requireImport_(s, reserved.id);
      updateRowByIndex_(
        getSheet_(SHEET_IMPORTS_),
        getHeaders_(getSheet_(SHEET_IMPORTS_)),
        r.rowIndex,
        {
          Status: "FAILED",
          LeaseUntil: "",
          ResultJSON: JSON.stringify({ error: String(error.message || error) }),
          UpdatedAt: nowIso_(),
        },
      );
    });
    throw error;
  }
}
function previewRawImport_(token, id, mappings, version) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var job = requireImport_(s, id).data;
  if (job.Status !== "PREVIEW" || timeValue_(job.ExpiresAt) < Date.now())
    throw new Error(
      "Preview tidak aktif / kedaluwarsa. Analisis file kembali.",
    );
  var payload = importPayload_(job.PayloadFileID);
  payload.preview = buildRawPreview_(payload.raw, payload.scope, mappings, s);
  var file = saveImportPayload_(id, payload);
  var updated = bulkWrites_(function () {
    requireSession_(token);
    var r = requireImport_(s, id);
    if (
      r.data.Status !== "PREVIEW" ||
      Number(r.data.PreviewVersion) !== Number(version)
    )
      throw new Error(
        "CONFLICT: Preview sudah berubah. Muat ulang panel upload.",
      );
    updateRowByIndex_(
      getSheet_(SHEET_IMPORTS_),
      getHeaders_(getSheet_(SHEET_IMPORTS_)),
      r.rowIndex,
      {
        PayloadFileID: file,
        PreviewVersion: Number(version) + 1,
        ResultJSON: JSON.stringify({
          preview: payload.preview.summary,
          totalItems: payload.preview.items.length,
          outcomes: [],
        }),
        UpdatedAt: nowIso_(),
      },
    );
    return r.data;
  });
  return previewResponse_(updated, payload);
}

// A short lease prevents two tabs from committing the same participant chunk.
// Sheets writes are not transactional. A failed chunk always needs review.
function importJobResponse_(d) {
  var result = parseJson_(d.ResultJSON, {});
  return {
    uploadId: d.UploadID,
    status: d.Status,
    kind: d.Kind,
    program: d.Program,
    batch: d.Batch,
    phase: d.Phase,
    fileName: d.FileName,
    previewVersion: Number(d.PreviewVersion || 1),
    nextIndex: Number(d.NextIndex || 0),
    written: Number(d.Written || 0),
    skipped: Number(d.Skipped || 0),
    failed: Number(d.Failed || 0),
    totalItems: Number(result.totalItems || 0),
    summary: result.preview || {},
    outcomes: result.outcomes || [],
    error: result.error || "",
    createdAt: d.CreatedAt,
    updatedAt: d.UpdatedAt,
  };
}
function importNeedsReview_(s, id, lease, error) {
  return bulkWrites_(function () {
    var r = requireImport_(s, id);
    if (r.data.Status !== "PROCESSING" || r.data.LeaseToken !== lease)
      return importJobResponse_(r.data);
    var result = parseJson_(r.data.ResultJSON, {});
    result.error = String(error.message || error).slice(0, 1200);
    updateRowByIndex_(
      getSheet_(SHEET_IMPORTS_),
      getHeaders_(getSheet_(SHEET_IMPORTS_)),
      r.rowIndex,
      {
        Status: "NEEDS_REVIEW",
        LeaseUntil: "",
        ResultJSON: JSON.stringify(result),
        UpdatedAt: nowIso_(),
      },
    );
    return importJobResponse_(r.data);
  });
}
function importAssessmentPlan_(item, scope, target, table) {
  var found = findAssessmentRow_(table.sheet, table.cfg, item.nik, item.batch),
    row = found.rowNumber ? table.grid[found.rowNumber - 1] : [],
    summary = parseJson_(target.enrollment.data.AssessmentSummary, {}),
    lineage = (summary[scope.phase] || {}).sources || {},
    scores = {},
    sources = {};
  table.cfg.modules.forEach(function (m) {
    var col = findHeaderIndex_(table.headers, m.headers);
    if (col < 0) throw new Error("Kolom modul " + m.label + " belum tersedia.");
    if (!Object.prototype.hasOwnProperty.call(item.scores, m.label)) return;
    var actual = numberOrBlank_(row[col]);
    if (item.expectedScores && actual !== item.expectedScores[m.label])
      throw new Error(
        "Nilai sumber berubah setelah preview. Analisis file kembali.",
      );
    var previous = lineage[m.label],
      incoming = item.sources[m.label];
    if (previous && /^MANUAL:/.test(previous.hash || "")) return;
    if (actual !== "" && actual >= item.scores[m.label]) return;
    if (
      previous &&
      previous.hash === incoming.hash &&
      actual === item.scores[m.label]
    )
      return;
    scores[m.label] = item.scores[m.label];
    sources[m.label] = incoming;
  });
  return {
    target: target,
    table: table,
    rowNumber: found.rowNumber,
    scores: scores,
    sources: sources,
  };
}
function writeImportedScores_(plan) {
  var t = plan.target,
    p = t.participant.data,
    table = plan.table,
    r = plan.rowNumber;
  if (!Object.keys(plan.scores).length) return false;
  if (!r) {
    r = Math.max(
      table.nextRow || table.grid.length + 1,
      table.cfg.dataStartRow,
    );
    table.nextRow = r + 1;
    var identity = {};
    table.headers.forEach(function (h) {
      var k = headerKey_(h);
      if (["nik", "employeeid"].indexOf(k) >= 0) identity[h] = p.NIK;
      else if (k === "batch") identity[h] = p.Batch;
      else if (["namakaryawan", "employeename", "name"].indexOf(k) >= 0)
        identity[h] = p.Name;
      else if (k === "region" || k === "locationgroupname")
        identity[h] = p.Region;
      else if (k === "jobtitlename") identity[h] = p.JobTitle;
      else if (k === "locationcityname")
        identity[h] = p.CityName || p.City || "";
    });
    patchColumns_(table.sheet, table.headers, r, identity);
    table.grid[r - 1] = table.headers.map(function (h) {
      return identity[h] === undefined ? "" : identity[h];
    });
    table.byKey[entityKey_(p.NIK, p.Program, p.Batch, p.Segment)] = [r];
  }
  var patch = {};
  table.cfg.modules.forEach(function (m) {
    if (Object.prototype.hasOwnProperty.call(plan.scores, m.label))
      patch[table.headers[findHeaderIndex_(table.headers, m.headers)]] =
        plan.scores[m.label];
  });
  patchColumns_(table.sheet, table.headers, r, patch);
  Object.keys(patch).forEach(function (h) {
    table.grid[r - 1][table.headers.indexOf(h)] = patch[h];
  });
  invalidateSheet_(table.name);
  context_().memo["assessment:" + table.name + "|" + table.cfg.program + "|" + table.cfg.phase] = table;
  t.importSources = {};
  t.importSources[table.cfg.phase] = plan.sources;
  t.updatedBy = "RAW_IMPORT";
  refreshEnrollmentSummary_(t);
  return true;
}
function safeImportRecord_(record) {
  var out = {};
  Object.keys(record).forEach(function (k) {
    var v = record[k];
    out[k] = typeof v === "string" && /^[=+@]/.test(v) ? "'" + v : v;
  });
  return out;
}
function applyRawImportChunk_(token, s, payload, job, start, lease) {
  if(payload.scope.kind === "LEGACY")return applyLegacyMigrationChunk_(token,s,payload,job,start,lease);
  return bulkWrites_(function () {
    requireRole_(requireSession_(token), ["ADMIN"]);
    var current = requireImport_(s, job.UploadID).data;
    if (
      current.Status !== "PROCESSING" ||
      current.LeaseToken !== lease ||
      Number(current.NextIndex || 0) !== start
    )
      throw new Error("CONFLICT: Proses upload berubah.");
    var scope = payload.scope,
      items = payload.preview.items.slice(start, start + 20),
      plans = [],
      outcomes = [],
      historyName =
        scope.kind === "SYLLABUS" ? SHEET_SCORE_EVENTS_ : scope.kind === "PANEL" ? SHEET_PANEL_OBSERVATIONS_ : SHEET_BEHAVIOR_,
      history = readAllAsObjects_(historyName),
      hashes = {},
      behaviorByEnrollment = {},
      table =
        scope.kind === "SYLLABUS" && scope.phase !== "PHASE 3"
          ? assessmentTable_(scope.program, scope.phase)
          : null;
    history.forEach(function (r) {
      var d = r.data;
      hashes[d.ContentHash] = true;
      if (scope.kind === "BEHAVIOR" || scope.kind === "PANEL")
        (
          behaviorByEnrollment[d.EnrollmentID] ||
          (behaviorByEnrollment[d.EnrollmentID] = [])
        ).push(d);
    });
    var panelParticipants=scope.kind === "PANEL" ? importTargetIndex_() : null;
    // Validate every target and every required source sheet before planning any write.
    items.forEach(function (item) {
      try {
        var t = authorizeEnrollment_(s, {
          nik: item.nik,
          enrollmentId: item.enrollmentId,
          program: scope.program,
          batch: item.batch,
        });
        if (
          entityKey_(
            t.participant.data.NIK,
            t.participant.data.Program,
            t.participant.data.Batch, t.participant.data.Segment,
          ) !== entityKey_(item.nik, scope.program, item.batch, item.segment)
        )
          throw new Error("Program / batch tujuan berubah.");
        if (
          String(t.enrollment.data.RecordVersion || "") !==
          String(item.expectedVersion)
        )
          throw new Error(
            "Data peserta berubah setelah preview. Analisis file kembali.",
          );
        if(scope.kind === "PANEL")item.records.forEach(function(r){if(r.SourceNameMatch){var check={name:r.SourceParticipantName,region:r.Region,nik:""};resolveRawPanelNames_({rows:[check]},scope,panelParticipants);if(check.issue||check.nik!==item.nik)throw new Error("CONFLICT: Identitas nama peserta panel berubah / ambigu setelah preview.");}});
        var plan = { item: item, target: t };
        if (table) {
          ["PHASE 1", "PHASE 2", "PHASE 3"].forEach(function (ph) {
            readAssessmentDetail_(item.nik, ph, scope.program, item.batch);
          });
          plan.assessment = importAssessmentPlan_(item, scope, t, table);
        }
        if(scope.kind === "SYLLABUS" && scope.phase === "PHASE 3")plan.phase3=phase3ImportPlan_(item,scope,t);
        plans.push(plan);
      } catch (error) {
        outcomes.push({
          nik: item.nik,
          status: "FAILED",
          message: error.message,
        });
      }
    });
    var written = 0,
      skipped = 0,
      appended = 0,
      hs = getSheet_(historyName),
      hh = getHeaders_(hs),
      now = nowIso_();
    plans.forEach(function (plan) {
      var item = plan.item,
        added = [];
      item.records.forEach(function (data) {
        if (hashes[data.ContentHash]) return;
        var record = Object.assign({}, data, {
          UploadID: job.UploadID,
          ImportedBy: s.nik,
          ImportedAt: now,
        });
        record[scope.kind === "SYLLABUS" ? "RecordID" : "ObservationID"] =
          newId_(scope.kind === "SYLLABUS" ? "SCR" : "OBS");
        appendObject_(hs, hh, safeImportRecord_(record));
        hashes[data.ContentHash] = true;
        added.push(record);
        appended++;
      });
      var transitioned=activateSourceGrading_(plan.target);
      var changed = false;
      if (scope.kind === "SYLLABUS")
        changed = scope.phase === "PHASE 3" ? writeImportedPhase3_(plan.phase3) : writeImportedScores_(plan.assessment);
      else {
        var all = (behaviorByEnrollment[item.enrollmentId] || []).concat(added),
          sum = behaviorAggregate_(all, scope.aggregation),
          old = parseJson_(plan.target.enrollment.data[scope.kind === "PANEL" ? "PanelSummary" : "BehaviourSummary"], {});
        if (
          added.length ||
          old.policy !== sum.policy ||
          old.score !== sum.score ||
          old.observations !== sum.observations
        ) {
          var ep=scope.kind === "PANEL" ? {PanelScore:Math.round(sum.score*2000)/100,PanelSummary:JSON.stringify(sum),UpdatedAt:now} : {BehaviourScore:sum.score,BehaviourCoreScore:sum.core,BehaviourTechnicalScore:sum.technical,BehaviourSummary:JSON.stringify(sum),UpdatedAt:now};
          updateRowByIndex_(getSheet_(SHEET_ENROLLMENTS),getHeaders_(getSheet_(SHEET_ENROLLMENTS)),plan.target.enrollment.rowIndex,ep);
          Object.assign(plan.target.enrollment.data,ep);
          var pp=scope.kind === "PANEL"?{PanelScore:ep.PanelScore,UpdatedAt:now}:{BehaviourObservation:sum.score,UpdatedAt:now};
          updateRowByIndex_(getSheet_(SHEET_PARTICIPANTS),getHeaders_(getSheet_(SHEET_PARTICIPANTS)),plan.target.participant.rowIndex,pp);
          Object.assign(plan.target.participant.data,pp);
          plan.target.updatedBy=s.nik;refreshEnrollmentSummary_(plan.target);
          changed = true;
        }
      }
      if(transitioned&&!changed){plan.target.updatedBy=s.nik;refreshEnrollmentSummary_(plan.target);changed=true;}
      if (changed || added.length) {
        written++;
        outcomes.push({
          nik: item.nik,
          status: "WRITTEN",
          message: added.length + " catatan baru",
        });
      } else {
        skipped++;
        outcomes.push({
          nik: item.nik,
          status: "SKIPPED",
          message: "Catatan sama / nilai tertinggi atau koreksi Admin dipertahankan.",
        });
      }
    });
    if (scope.kind === "SYLLABUS") {
      var mapSheet = getSheet_(SHEET_IMPORT_MAP_),
        mh = getHeaders_(mapSheet),
        existing = readAllAsObjects_(SHEET_IMPORT_MAP_),
        used = {},
        choices = {};
      payload.preview.columns.forEach(function (c) {
        (choices[c.sourceKey] || (choices[c.sourceKey] = {}))[
          c.module || "UNMAPPED"
        ] = true;
      });
      payload.preview.columns.forEach(function (c) {
        if (
          !c.module ||
          used[c.sourceKey] ||
          Object.keys(choices[c.sourceKey]).length !== 1
        )
          return;
        used[c.sourceKey] = true;
        var patch = {
            Program: scope.program,
            Phase: scope.phase,
            SourceKey: c.sourceKey,
            SourceLabel: c.sourceLabel,
            Module: c.module,
            UpdatedBy: s.nik,
            UpdatedAt: now,
          },
          r = uniqueMatch_(
            existing.filter(function (r) {
              return (
                r.data.Program === scope.program &&
                r.data.Phase === scope.phase &&
                r.data.SourceKey === c.sourceKey
              );
            }),
            "mapping import",
          );
        if (r) updateRowByIndex_(mapSheet, mh, r.rowIndex, patch);
        else appendObject_(mapSheet, mh, patch);
      });
    }
    auditLog_(
      s.nik,
      s.role,
      "RAW_IMPORT_CHUNK",
      scope.kind,
      job.UploadID,
      JSON.stringify({
        from: start,
        count: items.length,
        written: written,
        skipped: skipped,
        failed: outcomes.filter(function (x) {
          return x.status === "FAILED";
        }).length,
        records: appended,
      }),
    );
    return {
      next: start + items.length,
      written: written,
      skipped: skipped,
      failed: outcomes.filter(function (x) {
        return x.status === "FAILED";
      }).length,
      outcomes: outcomes,
    };
  });
}
function commitRawImport_(token, id, version, cursor) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var initial = requireImport_(s, id).data;
  if (initial.Status === "PREVIEW") {
    var first = importPayload_(initial.PayloadFileID);
    if (first.preview.summary.unmappedColumns)
      throw new Error("Selesaikan pemetaan kolom sebelum menyimpan.");
    if (!first.preview.items.length)
      throw new Error("Tidak ada peserta valid untuk disimpan.");
  }
  var reserved = bulkWrites_(function () {
    requireSession_(token);
    var r = requireImport_(s, id),
      d = r.data;
    if (
      d.Status === "COMPLETE" ||
      d.Status === "COMPLETE_WITH_ERRORS" ||
      d.Status === "NEEDS_REVIEW"
    )
      return { response: importJobResponse_(d) };
    if (d.Status === "PROCESSING") {
      if (Number(d.LeaseUntil) > Date.now())
        return { response: importJobResponse_(d) };
      var stale = parseJson_(d.ResultJSON, {});
      stale.error =
        "Proses sebelumnya terputus. Periksa sheet dan hasil sebelum membuat preview baru.";
      updateRowByIndex_(
        getSheet_(SHEET_IMPORTS_),
        getHeaders_(getSheet_(SHEET_IMPORTS_)),
        r.rowIndex,
        {
          Status: "NEEDS_REVIEW",
          LeaseUntil: "",
          ResultJSON: JSON.stringify(stale),
          UpdatedAt: nowIso_(),
        },
      );
      return { response: importJobResponse_(d) };
    }
    if (["PREVIEW", "READY"].indexOf(d.Status) < 0)
      throw new Error("Upload tidak siap disimpan.");
    if (
      Number(d.PreviewVersion) !== Number(version) ||
      Number(d.NextIndex || 0) !== Number(cursor)
    )
      throw new Error(
        "CONFLICT: Posisi / versi preview berubah. Muat status upload.",
      );
    if (d.Status === "PREVIEW" && timeValue_(d.ExpiresAt) < Date.now())
      throw new Error("Preview kedaluwarsa. Analisis file kembali.");
    var lease = Utilities.getUuid();
    updateRowByIndex_(
      getSheet_(SHEET_IMPORTS_),
      getHeaders_(getSheet_(SHEET_IMPORTS_)),
      r.rowIndex,
      {
        Status: "PROCESSING",
        LeaseToken: lease,
        LeaseUntil: Date.now() + 330000,
        UpdatedAt: nowIso_(),
      },
    );
    return { job: clone_(d), lease: lease };
  });
  if (reserved.response) return reserved.response;
  try {
    var payload = importPayload_(reserved.job.PayloadFileID);
    if (payload.preview.summary.unmappedColumns)
      throw new Error("Selesaikan pemetaan kolom sebelum menyimpan.");
    if (!payload.preview.items.length)
      throw new Error("Tidak ada peserta valid untuk disimpan.");
    var chunk = applyRawImportChunk_(
      token,
      s,
      payload,
      reserved.job,
      Number(cursor),
      reserved.lease,
    );
    // Acknowledge only after all score, summary and history writes succeeded.
    return bulkWrites_(function () {
      requireSession_(token);
      var r = requireImport_(s, id),
        d = r.data;
      if (d.Status !== "PROCESSING" || d.LeaseToken !== reserved.lease)
        throw new Error("CONFLICT: Lease upload berubah.");
      var result = parseJson_(d.ResultJSON, {}),
        failed = Number(d.Failed || 0) + chunk.failed;
      result.totalItems = payload.preview.items.length;
      result.outcomes = (result.outcomes || []).concat(chunk.outcomes);
      updateRowByIndex_(
        getSheet_(SHEET_IMPORTS_),
        getHeaders_(getSheet_(SHEET_IMPORTS_)),
        r.rowIndex,
        {
          Status:
            chunk.next >= payload.preview.items.length
              ? failed || payload.preview.summary.errorCount
                ? "COMPLETE_WITH_ERRORS"
                : "COMPLETE"
              : "READY",
          NextIndex: chunk.next,
          Written: Number(d.Written || 0) + chunk.written,
          Skipped: Number(d.Skipped || 0) + chunk.skipped,
          Failed: failed,
          ResultJSON: JSON.stringify(result),
          LeaseToken: "",
          LeaseUntil: "",
          UpdatedAt: nowIso_(),
        },
      );
      return importJobResponse_(d);
    });
  } catch (error) {
    return importNeedsReview_(s, id, reserved.lease, error);
  }
}
function getRawImportStatus_(token, id) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var d = requireImport_(s, id).data;
  if (d.Status === "PROCESSING" && Number(d.LeaseUntil) <= Date.now())
    return importNeedsReview_(
      s,
      id,
      d.LeaseToken,
      new Error("Proses terputus. Rekonsiliasi data sebelum preview baru."),
    );
  return importJobResponse_(d);
}
function getRawImportHistory_(token) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  var sh = getSheet_(SHEET_IMPORTS_),
    h = getHeaders_(sh),
    last = sh.getLastRow(),
    start = Math.max(2, last - 99);
  if (last < 2) return [];
  return sh
    .getRange(start, 1, last - start + 1, h.length)
    .getValues()
    .map(function (row) {
      return rowToObject_(h, row);
    })
    .filter(function (d) {
      return d.OwnerNIK === s.nik && normalizeSegment_(d.Segment)===activeSegment_();
    })
    .reverse()
    .map(importJobResponse_);
}
function saveAssessmentQleapLinks_(token, d) {
  var s = requireSession_(token);
  requireRole_(s, ["ADMIN"]);
  d = d || {};
  var program = String(d.program || "").toUpperCase(),
    phase = normalizeCurrentPhase_(d.phase),
    cfg = getAssessmentConfig_(program, phase);
  if (activePrograms_().indexOf(program) < 0 || (!cfg && phase !== "PHASE 3"))
    throw new Error("Pilih program dan Phase 1–3.");
  var allowed = phase === "PHASE 3" ? ["POST-TEST"] : cfg.modules.map(function (m) {
      return m.label;
    }),
    seen = {};
  if (!Array.isArray(d.links) || d.links.length > allowed.length)
    throw new Error("Daftar tautan tidak valid.");
  d.links.forEach(function (x) {
    if (allowed.indexOf(x.module) < 0 || seen[x.module])
      throw new Error("Modul QLEAP tidak valid / duplikat.");
    seen[x.module] = true;
    validateUrl_(x.url);
  });
  return bulkWrites_(function () {
    requireSession_(token);
    saveQleapMappings_(program, phase, d.links, s.nik);
    auditLog_(
      s.nik,
      s.role,
      "UPDATE_PHASE_QLEAP",
      "QLEAP Mapping",
      program + "|" + phase,
      d.links.length + " modul",
    );
    return { ok: true };
  });
}
function getBehaviorDetail_(token, d) {
  var s = requireSession_(token),
    t = authorizeEnrollment_(s, d),
    id = t.enrollment.data.EnrollmentID,
    list = readAllAsObjects_(SHEET_BEHAVIOR_)
      .filter(function (r) {
        return r.data.EnrollmentID === id;
      })
      .map(function (r) {
        var x = r.data;
        return {
          id: x.ObservationID,
          observedAt: x.ObservedAt,
          observerName: x.ObserverName,
          region: x.Region,
          core: numberOrBlank_(x.CoreScore),
          technical: numberOrBlank_(x.TechnicalScore),
          score: numberOrBlank_(x.AverageScore),
          details: parseJson_(x.DetailsJSON, {}),
          strengths: x.Strengths,
          improvementAreas: x.ImprovementAreas,
          recommendations: x.Recommendations,
        };
      })
      .sort(function (a, b) {
        return timeValue_(b.observedAt) - timeValue_(a.observedAt);
      });
  return {
    enrollmentId: id,
    historicalScore: legacyObservationScore_(t.enrollment.data,"BehaviourScore"),
    participantFallbackScore: participantObservationFallback_(t.enrollment.data,"Behaviour"),
    summary: parseJson_(t.enrollment.data.BehaviourSummary, {}),
    observations: list.slice(0, 20),
    total: list.length,
  };
}
function getRawImportOptions(t) {
  return api_("getRawImportOptions", function () {
    return getRawImportOptions_(t);
  }, arguments);
}
function prepareRawImport(t, d) {
  return api_("prepareRawImport", function () {
    return prepareRawImport_(t, d);
  }, arguments);
}
function previewRawImport(t, id, m, v) {
  return api_("previewRawImport", function () {
    return previewRawImport_(t, id, m, v);
  }, arguments);
}
function commitRawImport(t, id, v, c) {
  return api_("commitRawImport", function () {
    return commitRawImport_(t, id, v, c);
  }, arguments);
}
function getRawImportStatus(t, id) {
  return api_("getRawImportStatus", function () {
    return getRawImportStatus_(t, id);
  }, arguments);
}
function getRawImportHistory(t) {
  return api_("getRawImportHistory", function () {
    return getRawImportHistory_(t);
  }, arguments);
}
function saveAssessmentQleapLinks(t, d) {
  return api_("saveAssessmentQleapLinks", function () {
    return saveAssessmentQleapLinks_(t, d);
  }, arguments);
}
function getBehaviorDetail(t, d) {
  return api_("getBehaviorDetail", function () {
    return getBehaviorDetail_(t, d);
  }, arguments);
}

function getRawImportPreview(t, id) {
  return api_("getRawImportPreview", function () {
    var s = requireSession_(t);
    requireRole_(s, ["ADMIN"]);
    var d = requireImport_(s, id).data;
    if (d.Status !== "PREVIEW" || timeValue_(d.ExpiresAt) < Date.now())
      throw new Error(
        "Preview tidak aktif / kedaluwarsa. Analisis file kembali.",
      );
    return previewResponse_(d, importPayload_(d.PayloadFileID));
  }, arguments);
}

// Version 2.3: one import pipeline, with source-specific parsers and enrollment summaries.
var SHEET_PANEL_OBSERVATIONS_ = "Panel Observations";
SHEET_HEADERS[SHEET_PANEL_OBSERVATIONS_] = ["ObservationID","EnrollmentID","NIK","Program","Batch","ObservedAt","ObserverNIK","ObserverName","Region","AverageScore","DetailsJSON","Recommendations","ContentHash","UploadID","SourceRow","ImportedBy","ImportedAt"];
SHEET_HEADERS[SHEET_ENROLLMENTS] = SHEET_HEADERS[SHEET_ENROLLMENTS].concat(["PanelScore","PanelSummary","FinalSummary"]);
function ensureAutoRecapSchema_() {
  var c=context_(); if(c.memo.autoRecapSchema) return;
  withLock_(function() {
    getSheet_(SHEET_ENROLLMENTS);
    var definitions={};
    definitions[SHEET_SCORE_EVENTS_]=["AnswersJSON"];
    definitions[SHEET_ENROLLMENTS]=["PanelScore","PanelSummary","FinalSummary"];
    definitions[SHEET_PANEL_OBSERVATIONS_]=SHEET_HEADERS[SHEET_PANEL_OBSERVATIONS_];
    definitions[getSetting_("PHASE3_SHEET")||"Phase 3"]=["NIK","Program","Batch","PreTest","PostTest","UpdatedAt","AnswersJSON","Segment"];
    Object.keys(definitions).forEach(function(name) {
      var sh=c.ss.getSheetByName(name),old=sh?getHeaders_(sh):[],missing=definitions[name].filter(function(h){return old.indexOf(h)<0;});
      if(!sh||missing.length) {
        sh=ensureSheetWithHeaders_(c.ss,name,definitions[name]); c.sheets[name]=sh;
        delete c.headers[name]; delete c.tables[name]; invalidateSheet_(name);
      }
    });
    c.memo.autoRecapSchema=true;
  });
}
function rawAssessmentConfig_(program,phase) {
  return getAssessmentConfig_(program,phase) || (phase==="PHASE 3" && activePrograms_().indexOf(program)>=0 ? {program:program,phase:phase,modules:[{label:"POST-TEST",headers:["POST-TEST","PostTest","Supervisory Development Program"]}]} : null);
}
function parseRawPhase3_(rows,program) {
  var h=rows[0]||[],ni=findHeaderIndex_(h,["Nomor Induk Karyawan"]);
  if(ni<0) return parseRawSyllabus_(rows,program,"PHASE 3");
  var ts=findHeaderIndex_(h,["Timestamp"]),ty=findHeaderIndex_(h,["Pre-Test atau Post-Test"]),sc=findHeaderIndex_(h,["Score"]),nm=findHeaderIndex_(h,["Nama","Name"]),re=findHeaderIndex_(h,["Region"]);
  if(ts<0||ty<0||sc<0) throw new Error("Form Phase 3 membutuhkan Timestamp, jenis pre/post-test, dan Score.");
  return {format:"PHASE3_FORM_TOTAL",columns:[{id:"post",col:sc,dateCol:ts,sourceLabel:"Supervisory Development Program · POST-TEST",sourceKey:"phase3posttest",suggested:"POST-TEST"}],rows:rows.slice(1).map(function(row,i) {
    if(!row.some(function(v){return rawText_(v)!=="";})) return null;
    var type=headerKey_(row[ty]);
    if(type!=="pretest"&&type!=="posttest") throw new Error("Baris "+(i+2)+": jenis tes tidak dikenal.");
    var answers={};h.forEach(function(label,c){if(c>sc&&rawText_(row[c]))answers[rawText_(label)]=rawText_(row[c]);});
    return {rowNo:i+2,nik:rawNik_(row[ni]),name:rawText_(row[nm]),region:rawText_(row[re]),preTest:type==="pretest"?row[sc]:"",testAt:rawDate_(row[ts]),answers:answers,cells:type==="posttest"?[{columnId:"post",value:row[sc],examAt:rawDate_(row[ts]),invalidDate:!rawDate_(row[ts])}]:[]};
  }).filter(Boolean)};
}
function parseRawPanel_(rows,program) {
  var hr=rawHeader_(rows,["Nama Panelis"]),h=rows[hr],at=function(a){return findHeaderIndex_(h,a);},ts=at(["Timestamp"]),oi=at(["NIK Panelis", "NIK"]),on=at(["Nama Panelis"]),re=at(["Region","Silahkan pilih Region"]),blocks=[];
  h.forEach(function(label,i) {
    var k=headerKey_(label);
    if(program==="TSHDP" && /^niknamapeserta\d*$/.test(k)) blocks.push({identity:i,ratings:[]});
    if((program==="ASHDP"||program==="SLDP") && /^namapeserta\d*$/.test(k)) blocks.push({identity:i,ratings:[i+1,i+2,i+3,i+4,i+5,i+6,i+7,i+8],feedback:i+9});
  });
  if(program==="TSHDP") {
    var first=blocks.length?blocks[blocks.length-1].identity+1:-1;
    blocks.forEach(function(b){b.ratings=Array.from({length:8},function(_,i){return first+i;});b.feedback=first+8;});
  }
  if(ts<0||on<0||!blocks.length) throw new Error("Gunakan Form Responses 1 asli panel, termasuk kolom pilihan peserta dan 8 rubrik penilaian.");
  var out=[];
  rows.slice(hr+1).forEach(function(row,index) {
    if(!row.some(function(v){return rawText_(v)!=="";})) return;
    var chosen=blocks.filter(function(b){return rawText_(row[b.identity]);}),issue="",nik="",name="",ratings={},values=[];
    if(chosen.length!==1) issue="Harus tepat satu pilihan peserta panel per respons.";
    var b=chosen[0];
    if(b) {
      name=rawText_(row[b.identity]);
      if(program==="TSHDP") {var match=name.match(/^(\d{4,18})(?:\.0)?\s+(.+)$/);if(match){nik=rawNik_(match[1]);name=match[2];}}
      b.ratings.forEach(function(c) {try{var v=rawNumber_(row[c],1,5);if(v==="")throw new Error("kosong");values.push(v);var label=rawText_(h[c]).replace(/_x000a_/gi," ");if((program==="ASHDP"||program==="SLDP")&&b.identity!==blocks[0].identity)label=label.replace(/\s+[2-6]$/,"");ratings[label]=v;}catch(e){if(!issue)issue="Delapan rubrik panel wajib berupa angka 1–5.";}});
    }
    out.push({rowNo:hr+index+2,nik:nik,name:name,region:rawText_(row[re]),observedAt:rawDate_(row[ts]),observerNik:rawNik_(row[oi]),observerName:rawText_(row[on]),average:values.length===8?values.reduce(function(a,b){return a+b;},0)/8:"",details:{ratings:ratings,checks:{}},recommendations:b?rawText_(row[b.feedback]):"",issue:issue});
  });
  return {format:"PANEL_8_RUBRICS",columns:[],rows:out};
}
function panelData_(r,context,enrollmentId) {
  if(r.issue)throw new Error(r.issue);
  if(!r.observedAt)throw new Error("Timestamp panel tidak dikenali.");
  if(!r.observerNik&&!r.observerName)throw new Error("Identitas panelis wajib terisi.");
  var score=rawNumber_(r.average,1,5);if(score==="")throw new Error("Delapan nilai rubrik panel belum lengkap.");
  var obj={EnrollmentID:enrollmentId,NIK:r.nik,Program:context.program,Batch:context.batch,ObservedAt:r.observedAt,ObserverNIK:r.observerNik,ObserverName:r.observerName,Region:r.region,AverageScore:score,DetailsJSON:JSON.stringify(r.details),Recommendations:r.recommendations,SourceRow:r.rowNo,SourceNameMatch:!!r.matchedByName,SourceParticipantName:r.name};
  obj.ContentHash=cacheKey_("",JSON.stringify([enrollmentId,obj.ObservedAt,obj.ObserverNIK,obj.ObserverName,score,obj.DetailsJSON,obj.Recommendations]));return obj;
}
function resolveRawPanelNames_(raw,scope,participants) {
  // ASHDP form's NIK identifies the panelist; never use it as the participant NIK.
  raw.rows.forEach(function(r) {
    if(r.nik||r.issue)return;
    if(!r.region){r.issue="Region wajib terisi untuk pencocokan nama peserta panel.";return;}
    var matches=Object.keys(participants).map(function(k){return participants[k].data;}).filter(function(p) {
      return String(p.Program).toUpperCase()===scope.program && (scope.batch==="AUTO"||sameBatch_(p.Batch,scope.batch)) && normalizeModuleKey_(p.Name)===normalizeModuleKey_(r.name) && (!r.region||normalizeRegion_(p.Region)===normalizeRegion_(r.region));
    });
    if(matches.length===1){r.nik=normalizeNik_(matches[0].NIK);r.matchedByName=true;}
    else r.issue=matches.length?"Nama peserta ambigu pada program/batch/region tujuan; lengkapi NIK peserta di sumber.":"Nama peserta panel belum cocok dengan roster program/batch/region tujuan.";
  });
}
function phase3Table_() {
  var name=getSetting_("PHASE3_SHEET")||"Phase 3",sh=getSheet_(name);return {name:name,sheet:sh,headers:getHeaders_(sh),rows:readAllAsObjects_(name)};
}
function phase3ImportPlan_(item,scope,target) {
  scope=Object.assign({},scope,{batch:item.batch});
  var a=readAssessmentDetail_(item.nik,"PHASE 3",scope.program,scope.batch),actual={};a.modules.forEach(function(m){actual[m.module]=m.score;});
  Object.keys(item.expectedScores||{}).forEach(function(k){if(actual[k]!==item.expectedScores[k])throw new Error("Nilai sumber berubah setelah preview. Analisis file kembali.");});
  return {item:item,target:target,scope:scope};
}
function writeImportedPhase3_(plan) {
  var item=plan.item,t=plan.target,scope=plan.scope,table=phase3Table_(),key=entityKey_(item.nik,scope.program,scope.batch, scope.segment),matches=table.rows.filter(function(r){return entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment)===key;});
  if(matches.length>1)throw new Error("Phase 3 duplikat untuk program dan batch tujuan.");
  if(!Object.keys(item.scores).length&&!item.preTests.length)return false;
  var row=matches[0],patch={NIK:item.nik,Program:scope.program,Batch:scope.batch,UpdatedAt:nowIso_()},source=item.sources["POST-TEST"],prev=(parseJson_(t.enrollment.data.AssessmentSummary,{})["PHASE 3"]||{}).sources||{};
  if(Object.prototype.hasOwnProperty.call(item.scores,"POST-TEST")) {
    patch.PostTest=item.scores["POST-TEST"];
    var rec=item.records.filter(function(r){return r.ContentHash===(source||{}).hash;})[0];
    if(rec&&rec.AnswersJSON)patch.AnswersJSON=rec.AnswersJSON;
  }
  if(item.preTests.length) {
    var pre=item.preTests.reduce(function(a,b){return a.score>=b.score?a:b;});
    patch.PreTest=row&&numberOrBlank_(row.data.PreTest)!==""?Math.max(Number(row.data.PreTest),pre.score):pre.score;
  }
  if(row)updateRowByIndex_(table.sheet,table.headers,row.rowIndex,patch);else appendObject_(table.sheet,table.headers,patch);
  t.importSources={"PHASE 3":Object.assign({},prev,item.sources)};t.updatedBy="RAW_IMPORT";refreshEnrollmentSummary_(t);return true;
}
function savePhase3Scores_(token,data) {
  var s=requireSession_(token);requireRole_(s,["ADMIN"]);
  return bulkWrites_(function(){
    ensureAutoRecapSchema_();var t=authorizeEnrollment_(s,data),e=t.enrollment.data;
    if(!data.recordVersion||data.recordVersion!==e.RecordVersion)throw new Error("CONFLICT: Nilai berubah. Muat ulang detail.");
    var input=(data.scores||[]).filter(function(m){return m.module==="POST-TEST";});
    if(input.length!==1||(data.scores||[]).some(function(m){return m.module!=="POST-TEST";}))throw new Error("Edit hanya nilai post-test Phase 3.");
    var score=validScore_(input[0].score),table=phase3Table_(),matches=table.rows.filter(function(r){return entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment)===entityKey_(data.nik,data.program,data.batch, data.segment);});
    if(matches.length>1)throw new Error("Phase 3 duplikat.");
    var patch={NIK:normalizeNik_(data.nik),Program:data.program,Batch:data.batch,PostTest:score,UpdatedAt:nowIso_()};
    if(matches[0])updateRowByIndex_(table.sheet,table.headers,matches[0].rowIndex,patch);else appendObject_(table.sheet,table.headers,patch);
    t.importSources={"PHASE 3":{"POST-TEST":{examAt:nowIso_(),hash:"MANUAL:"+Utilities.getUuid(),sourceLabel:"INPUT ADMIN"}}};t.updatedBy=s.nik;t.manualPhase="PHASE 3";refreshEnrollmentSummary_(t);
    auditLog_(s.nik,s.role,"SAVE_ASSESSMENT","Enrollment",e.EnrollmentID,"PHASE 3 manual override");return {ok:true,message:"Nilai Phase 3 tersimpan; rekap akhir diperbarui otomatis."};
  });
}
function finalGrade_(e,p,summary) {
  var missing=[],components=[{key:"Phase1Score",label:"Phase 1",weight:.10},{key:"Phase2Score",label:"Phase 2",weight:.15},{key:"Phase3Score",label:"Soft Competency",weight:.15},{key:"BehaviourScore",label:"Behavior",weight:.30},{key:"PanelScore",label:"Panel",weight:.30}],total=0;
  components.forEach(function(c,i){
    var raw=e[c.key];
    if(i===3&&numberOrBlank_(raw)==="")raw=p.BehaviourObservation;
    if(i===4&&numberOrBlank_(raw)==="")raw=p.PanelScore;
    var n=numberOrBlank_(raw);
    if(i===3&&n!==""&&n>=1&&n<=5&&!(parseJson_(e.BehaviourSummary,{}).aggregateFallback&&parseJson_(e.BehaviourSummary,{}).scale===100))n=n*20;
    var ph=summary["PHASE "+(i+1)];
    if(n===""||n<0||n>100||(i<3&&ph&&!ph.aggregateFallback&&ph.completedModules<ph.totalModules)){missing.push(c.label);c.score="";return;}
    c.score=n;total+=n*c.weight;
  });
  return {score:missing.length?"":Math.round((total+1e-9)*100)/100,complete:!missing.length,missing:missing,components:components,passingScore:80};
}
function getPanelDetail_(token,d) {
  var s=requireSession_(token),t=authorizeEnrollment_(s,d),id=t.enrollment.data.EnrollmentID,c=context_();getSheet_(SHEET_ENROLLMENTS);
  var sh=c.ss.getSheetByName(SHEET_PANEL_OBSERVATIONS_),list=sh?readAllAsObjects_(SHEET_PANEL_OBSERVATIONS_).filter(function(r){return r.data.EnrollmentID===id;}).map(function(r){var x=r.data;return {id:x.ObservationID,observedAt:x.ObservedAt,observerName:x.ObserverName,region:x.Region,score:numberOrBlank_(x.AverageScore),details:parseJson_(x.DetailsJSON,{}),recommendations:x.Recommendations};}).sort(function(a,b){return timeValue_(b.observedAt)-timeValue_(a.observedAt);}):[];
  return {enrollmentId:id,historicalScore:legacyObservationScore_(t.enrollment.data,"PanelScore"),participantFallbackScore:participantObservationFallback_(t.enrollment.data,"Panel"),summary:parseJson_(t.enrollment.data.PanelSummary,{}),observations:list.slice(0,20),total:list.length};
}
function getPanelDetail(token,d){return api_("getPanelDetail",function(){return getPanelDetail_(token,d);}, arguments);}
function releaseAssessmentOverrides(token,d) {
  return api_("releaseAssessmentOverrides",function(){var s=requireSession_(token);requireRole_(s,["ADMIN"]);return bulkWrites_(function(){
    var t=authorizeEnrollment_(s,d),e=t.enrollment.data;if(!d.recordVersion||d.recordVersion!==e.RecordVersion)throw new Error("CONFLICT: Muat ulang detail nilai.");
    var phase=normalizeCurrentPhase_(d.phase);if(["PHASE 1","PHASE 2","PHASE 3"].indexOf(phase)<0)throw new Error("Phase tidak valid.");
    var summary=parseJson_(e.AssessmentSummary,{}),value=summary[phase]||{},count=0;
    Object.keys(value.sources||{}).forEach(function(k){if(/^MANUAL:/.test(value.sources[k].hash||"")){delete value.sources[k];count++;}});
    summary[phase]=value;updateRowByIndex_(getSheet_(SHEET_ENROLLMENTS),getHeaders_(getSheet_(SHEET_ENROLLMENTS)),t.enrollment.rowIndex,{AssessmentSummary:JSON.stringify(summary),UpdatedAt:nowIso_()});
    auditLog_(s.nik,s.role,"RELEASE_MANUAL_SCORES","Enrollment",e.EnrollmentID,phase+": "+count+" modules");return {ok:true,message:"Proteksi dilepas. Impor berikutnya boleh menaikkan nilai; nilai yang sudah ada tetap menjadi batas minimum."};
  });}, arguments);
}
function ensureRawImportEnrollments_(program,batch) {
  if(batch!=="AUTO")return ensureImportEnrollments_(program,batch);
  var batches={};readAllAsObjects_(SHEET_PARTICIPANTS).forEach(function(r){if(String(r.data.Program).toUpperCase()===program&&normalizeBatchKey_(r.data.Batch))batches[String(r.data.Batch)]=true;});
  Object.keys(batches).forEach(function(b){ensureImportEnrollments_(program,b);});
}
function rawTargetBatch_(row,scope,participants) {
  if(scope.batch!=="AUTO")return scope.batch;
  var list=Object.keys(participants).map(function(k){return participants[k].data;}).filter(function(p){return normalizeNik_(p.NIK)===row.nik&&String(p.Program).toUpperCase()===scope.program;});
  if(row.batch){var exact=list.filter(function(p){return sameBatch_(p.Batch,row.batch);});list=exact.length?exact:list.filter(function(p){return assessmentBatchCompatible_(row.batch,p.Batch);});}
  if(list.length!==1)throw new Error(list.length?"NIK mempunyai beberapa batch pada program ini. Pilih batch tujuan secara khusus.":"Peserta/batch belum terdaftar pada program tujuan.");
  return list[0].Batch;
}

// Historical participant migration uses the existing Import Sessions lease/chunk pipeline.
var LEGACY_SOURCE_DEFAULTS_ = [
  {program:"ASHDP",url:"https://docs.google.com/spreadsheets/d/1CNa3Fev2UwW3ql16gBSzCyZjtQflwuhm39UCRkaXDew/edit"},
  {program:"TSHDP",url:"https://docs.google.com/spreadsheets/d/1g6H10pn4796FEtIO3SMhEc9W-wHxKtIfV3LmV30_1FM/edit"}
];
SHEET_HEADERS[SHEET_ENROLLMENTS].push("LegacySnapshot");
function legacySnapshot_(e){return parseJson_(e.LegacySnapshot,{});}
function ensureLegacySchema_(){
  ensureAutoRecapSchema_();var c=context_(),sh=getSheet_(SHEET_ENROLLMENTS),h=getHeaders_(sh);
  if(h.indexOf("LegacySnapshot")<0)withLock_(function(){ensureSheetWithHeaders_(c.ss,SHEET_ENROLLMENTS,["LegacySnapshot"]);delete c.headers[SHEET_ENROLLMENTS];invalidateSheet_(SHEET_ENROLLMENTS);});
}
function legacyBlank_(v){return v===null||v===undefined||String(v).trim()==="";}
function legacyLink_(v){var s=rawText_(v);return /^(?:-|n\/?a|#n\/a|none)$/i.test(s)?"":s;}
function legacyNumber_(v){
  if(legacyBlank_(v)||/^(?:-|n\/?a|#n\/a|#ref!|#value!|#div\/0!)$/i.test(rawText_(v)))return "";
  var n=Number(rawText_(v).replace(/\s/g,"").replace(",","."));
  if(!isFinite(n)||n<0||n>100)throw new Error("Nilai harus 0–100: "+rawText_(v));return n;
}
function legacyRegion_(v){var s=rawText_(v).toUpperCase(),m=s.match(/\bREGION\s*[-_#]?\s*(\d{1,2})\b/)||s.match(/^(\d{1,2})$/);return m?"REGION "+Number(m[1]):s;}
function legacyBatch_(v){var k=normalizeBatchKey_(v);if(!/^\d+(?: 20\d\d)?$/.test(k))throw new Error("Batch belum dapat dikenali: "+rawText_(v));return "Batch "+k;}
function legacyStatus_(v){var k=headerKey_(v);if(!k||k==="na")return "";if(k==="resign"||k==="resigned")return "RESIGN";if(k==="move"||k==="moved")return "MOVE";if(["pass","passed","lulus"].indexOf(k)>=0)return "LULUS";if(["notpass","notpassed","fail","failed","tidaklulus","gagal"].indexOf(k)>=0)return "TIDAK LULUS";if(["ongoing","onprogress","inprogress","belumselesai"].indexOf(k)>=0)return "ON GOING";if(k.indexOf("remedial")>=0)return "REMEDIAL";throw new Error("Status kelulusan belum dikenal: "+rawText_(v));}
function parseLegacyRosterBase_(rows,program,reference){
  var out=[],errors=[],h=null,headerRow=0,skipped=0;
  rows.forEach(function(row,index){
    if(findHeaderIndex_(row,["NIK","Employee ID"])>=0&&findHeaderIndex_(row,["Batch"])>=0){h=row;headerRow=index+1;return;}
    if(!h){skipped++;return;}
    function val(aliases,last){var found=[];h.forEach(function(x,i){if(aliases.map(headerKey_).indexOf(headerKey_(x))>=0)found.push(i);});var i=last?found[found.length-1]:found.filter(function(c){return !legacyBlank_(row[c]);})[0];if(i===undefined)i=found[0];return i===undefined?"":row[i];}
    var rawProgram=rawText_(val(["Program"])).toUpperCase();
    var rawNik=val(["NIK","Employee ID"]),rawName=val(["Name","Nama","Nama Lengkap","Employee Name"]),rawBatch=val(["Batch"]);
    if(legacyBlank_(rawNik)&&legacyBlank_(rawName)&&legacyBlank_(rawBatch)){skipped++;return;}
    if(legacyBlank_(rawNik)&&legacyBlank_(rawBatch)){skipped++;return;}
    try{
      if(rawProgram&&rawProgram!==program)throw new Error("Program pada baris sumber berbeda dari program tujuan.");
      var nik=rawNik_(rawNik);if(!/^\d{4,18}$/.test(nik))throw new Error("NIK kosong/tidak valid; tidak dibuat dari nama.");
      var name=rawText_(rawName),region=legacyRegion_(val(["Region","Location Group Name","Org Group Name"]));
      if(!name||!region)throw new Error("Nama dan region wajib tersedia.");
      var batch=legacyBatch_(rawBatch),modern=program==="TSHDP"||findHeaderIndex_(h,["Avarage Score Phase 1 - Self Learning","Average Score Phase 1 - Self Learning","Phase1Score"])>=0;
      var data={NIK:nik,Name:name,Program:program,Batch:batch,Region:region,
        JobTitle:legacyLink_(val(["Job Title","Job Title Name","JobTitle"])),CityName:legacyLink_(val(["Location City Name","City","CityName"])),Email:legacyLink_(val(["Employee Email Primary","Employee Email","Email"])),
        FinalScore:legacyNumber_(val(["Final Score","FinalScore"])),FinalStatus:legacyStatus_(val(["Final Status","FinalStatus"])),
        PanelScore:legacyNumber_(val(["Panel Score","PanelScore","Panel"],true)),BehaviourObservation:legacyNumber_(val(["Behaviours Monitoring","Behaviours Observation","BehaviourObservation","Observasi"])),
        CertificateUrl:legacyLink_(val(["E-Certificate","CertificateUrl"])),Theme:legacyLink_(val(["Theme"])),ProjectTitle:legacyLink_(val(["Judul"])),ActionPlan:legacyLink_(val(["Action Plan"])),Implementation:legacyLink_(val(["Implementation"])),SlideUrl:legacyLink_(val(["Slide","SlideUrl"])),
        PanelDate:rawDate_(val(["Panel Date","PanelDate","Panel"]))};
      var rawPS=val(["Participant Status","Status"]),ps=normalizeParticipantStatus_(rawPS);if(!legacyBlank_(rawPS)&&PARTICIPANT_STATUS_OPTIONS.indexOf(ps)>=0)data.ParticipantStatus=ps;if(data.FinalStatus==="RESIGN")data.ParticipantStatus="RESIGN";
      var track=headerKey_(val(["Status LDP","Program Track"]));if(track.indexOf("talent")>=0&&track.indexOf("stand")>=0)data.ProgramTrack="BOTH";else if(track.indexOf("talent")>=0)data.ProgramTrack="TALENT";else if(track.indexOf("stand")>=0)data.ProgramTrack="STANDARDISASI";
      if(data.CertificateUrl)validateUrl_(data.CertificateUrl);if(data.SlideUrl)validateUrl_(data.SlideUrl);
      var metrics=[];
      if(modern){
        data.Phase1Score=legacyNumber_(val(["Score Phase 1","Phase1Score","Avarage Score Phase 1 - Self Learning","Average Score Phase 1 - Self Learning"]));
        data.Phase2Score=legacyNumber_(val(["Score Phase 2","Phase2Score","Avarage Score Phase 2 - Technical Competency","Average Score Phase 2 - Technical Competency"]));
        data.Phase3Score=legacyNumber_(val(["Score Phase 3","Phase3Score","Avarage Score Phase 3 - Supervisory Development Program","Average Score Phase 3 - Supervisory Development Program"]));
      }else ["Post Test Average Score","In Class Training Score","In Class Training"].forEach(function(label){var value=legacyNumber_(val([label]));if(value!=="")metrics.push({label:label,score:value});});
      data.CurrentPhase=["LULUS","TIDAK LULUS"].indexOf(data.FinalStatus)>=0?"COMPLETED":data.PanelScore!==""?"PANELISASI":data.Phase3Score!==undefined&&data.Phase3Score!==""?"PHASE 3":data.Phase2Score!==undefined&&data.Phase2Score!==""?"PHASE 2":"PHASE 1";
      var ref={program:program,spreadsheetId:reference.spreadsheetId||"",sheetName:reference.sheetName||"",row:index+1,headerRow:headerRow,batchLabel:rawText_(rawBatch)};
      data.Segment=activeSegment_()==="ALL"?"B2B":activeSegment_();out.push({key:entityKey_(nik,program,batch,data.Segment),data:data,metrics:metrics,source:ref,legacyLayout:!modern});
    }catch(error){errors.push({id:"source:"+program+":"+(index+1),program:program,rowNo:index+1,nik:rawText_(rawNik),message:error.message,resolvable:false});}
  });
  if(!h)throw new Error("Header NIK/Employee ID dan Batch tidak ditemukan pada master peserta "+program+".");
  return {rows:out,errors:errors,skipped:skipped};
}
function readLegacySource_(source){
  var p=String(source.program||"").toUpperCase();if(activePrograms_().indexOf(p)<0)throw new Error("Program sumber tidak valid.");
  var m=rawText_(source.url).match(/^https:\/\/docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]{20,100})(?:[/?#]|$)/);if(!m)throw new Error("Gunakan link Google Sheets database peserta.");
  if(m[1]===props_().getProperty("LDP_SPREADSHEET_ID"))throw new Error("Sumber migrasi harus berbeda dari database aplikasi.");
  var ss=SpreadsheetApp.openById(m[1]),title=typeof ss.getName==="function"?String(ss.getName()).toUpperCase():"",detected=activePrograms_().filter(function(x){return title.indexOf(x)>=0;});if(detected.length>1||detected.length===1&&detected[0]!==p)throw new Error("Nama database sumber tidak cocok dengan program tujuan.");var sheets=ss.getSheets(),preferred=p==="ASHDP"?["Participant","All Participants","Participants"]:["All Participants","Participant","Participants"],sh=null;
  preferred.some(function(name){sh=sheets.filter(function(x){return headerKey_(x.getName())===headerKey_(name);})[0];return !!sh;});
  if(!sh)throw new Error("Tab master peserta tidak ditemukan pada sumber "+p+".");
  var n=sh.getLastRow(),c=sh.getLastColumn();if(n>5000||c>100)throw new Error("Master peserta maksimal 5.000 baris dan 100 kolom.");
  if(n*c>250000)throw new Error("Master terlalu besar; gunakan salinan tab peserta saja.");
  return parseLegacyMaster_(n&&c?sh.getRange(1,1,n,c).getValues():[],p,{spreadsheetId:m[1],sheetName:sh.getName()});
}
function legacyTargets_(){
  var p={},e={},c={};readAllAsObjects_(SHEET_PARTICIPANTS).forEach(function(r){var k=entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment);(p[k]||(p[k]=[])).push(r);});readAllAsObjects_(SHEET_ENROLLMENTS).forEach(function(r){var k=entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment);(e[k]||(e[k]=[])).push(r);});readAllAsObjects_(SHEET_CERTIFICATES).forEach(function(r){var id=r.data.EnrollmentID;if(id)(c[id]||(c[id]=[])).push(r);});return {p:p,e:e,c:c};
}
function legacyTargetHash_(p,e,certs){return cacheKey_("",JSON.stringify([p?p.data:null,e?e.data:null,(certs||[]).map(function(r){return r.data;})]));}
function legacySame_(field,a,b){
  if(field==="BehaviourScore"||field==="BehaviourObservation"){var av=numberOrBlank_(a),bv=numberOrBlank_(b);if(av!==""&&bv!==""&&av>5&&bv>=1&&bv<=5)b=bv*20;else if(av!==""&&bv!==""&&bv>5&&av>=1&&av<=5)a=av*20;}
  if(/Score$/.test(field)||field==="BehaviourObservation")return numberOrBlank_(a)!==""&&numberOrBlank_(b)!==""&&Math.abs(Number(a)-Number(b))<.011;
  if(field==="Batch")return sameBatch_(a,b);if(field==="Region")return legacyRegion_(a)===legacyRegion_(b);
  if(field==="FinalStatus"){try{return legacyStatus_(a)===legacyStatus_(b);}catch(error){return String(a)===String(b);}}
  if(field==="Name")return normalizeModuleKey_(a)===normalizeModuleKey_(b);
  return rawText_(a).toLowerCase()===rawText_(b).toLowerCase();
}
function legacyCurrent_(p,e,field){var mapping={FinalStatus:"ProgramStatus",BehaviourObservation:"BehaviourScore"},ek=mapping[field]||field;return e&&!legacyBlank_(e.data[ek])?e.data[ek]:p?p.data[field]:"";}
function buildLegacyPreview_(raw,resolutions){
  resolutions=resolutions||{};var targets=legacyTargets_(),groups={},errors=(raw.errors||[]).slice(),items=[],display=[],duplicateRows=0,reviewCount=0;
  raw.rows.forEach(function(r){r=Object.assign({},r,{data:sourceRosterData_(r.data),metrics:r.metrics||[]});(groups[r.key]||(groups[r.key]=[])).push(r);});
  Object.keys(groups).forEach(function(k){
    var rows=groups[k],r=rows[0],signature=JSON.stringify([r.data,r.metrics,r.fallback]),conflicting=rows.some(function(x){return JSON.stringify([x.data,x.metrics,x.fallback])!==signature;}),ps=targets.p[k]||[],es=targets.e[k]||[];
    var issue="";if(conflicting)issue="Satu NIK/program/batch memiliki data sumber yang bertentangan.";else if(ps.length>1||es.length>1)issue="Identitas duplikat pada database aplikasi.";
    if(issue){errors.push({id:k,program:r.data.Program,rowNo:r.source.row,nik:r.data.NIK,message:issue,resolvable:false});display.push({key:k,data:r.data,status:"PERLU DICEK",conflicts:[issue],resolvable:false});return;}
    duplicateRows+=rows.length-1;
    var p=ps[0],e=es[0],cs=targets.c[e&&e.data.EnrollmentID]||[],conflicts=[];if(cs.length>1){errors.push({id:k,program:r.data.Program,rowNo:r.source.row,nik:r.data.NIK,message:"Sertifikat enrollment duplikat pada aplikasi.",resolvable:false});return;}
    Object.keys(r.data).forEach(function(field){var v=r.data[field],current=field==="CertificateUrl"&&cs[0]&&!legacyBlank_(cs[0].data.FileOrLink)?cs[0].data.FileOrLink:legacyCurrent_(p,e,field);if(!legacyBlank_(v)&&!legacyBlank_(current)&&!legacySame_(field,current,v))conflicts.push({field:field,current:current,source:v});});
    var mode=["KEEP_CURRENT","USE_SOURCE"].indexOf(resolutions[k])>=0?resolutions[k]:"";
    if(conflicts.length&&!mode){reviewCount++;errors.push({id:k,program:r.data.Program,rowNo:r.source.row,nik:r.data.NIK,message:"Data aplikasi berbeda: "+conflicts.map(function(c){return c.field;}).join(", ")+".",resolvable:true});}
    var item={key:k,nik:r.data.NIK,name:r.data.Name,program:r.data.Program,batch:r.data.Batch,data:r.data,metrics:r.metrics,fallback:r.fallback||{},source:r.source,sourceRows:rows.map(function(x){return x.source.row;}),conflicts:conflicts,resolution:mode,expectedHash:legacyTargetHash_(p,e,cs),action:p||e?"LENGKAPI":"TAMBAH",legacyLayout:r.legacyLayout};
    display.push({key:k,data:r.data,status:conflicts.length&&!mode?"PERLU DICEK":mode==="KEEP_CURRENT"?"PERTAHANKAN APLIKASI":mode==="USE_SOURCE"?"PAKAI SUMBER DIPILIH":item.action,conflicts:conflicts,resolvable:conflicts.length>0});
    if(!conflicts.length||mode)items.push(item);
  });
  return {items:items,display:display,errors:errors,columns:[],warnings:[],summary:{kind:"LEGACY",program:"ALL",batch:"AUTO",phase:"",unmappedColumns:0,sourceRows:raw.rows.length,matchedParticipants:items.length,reviewCount:reviewCount,errorCount:errors.length,warningCount:0,duplicateRows:duplicateRows,selectedSource:items.filter(function(i){return i.resolution==="USE_SOURCE";}).length,keptCurrent:items.filter(function(i){return i.resolution==="KEEP_CURRENT";}).length,skippedRows:raw.skipped||0}};
}
function legacyPreviewResponse_(job,payload){return {uploadId:job.UploadID,status:job.Status,previewVersion:Number(job.PreviewVersion||1),expiresAt:job.ExpiresAt,summary:payload.preview.summary,rows:payload.preview.display,errors:payload.preview.errors,totalItems:payload.preview.items.length};}
function getLegacyMigrationOptions(token){return api_("getLegacyMigrationOptions",function(){var s=requireSession_(token);requireRole_(s,["ADMIN"]);return {sources:(activeSegment_()==="B2C"?[]:LEGACY_SOURCE_DEFAULTS_),maxRows:5000,gradingMode:"SOURCE"};}, arguments);}
function prepareLegacyMigration(token,d){return api_("prepareLegacyMigration",function(){
  var s=requireSession_(token);requireRole_(s,["ADMIN"]);d=d||{};
  if(!/^[A-Za-z0-9_-]{12,100}$/.test(d.requestId||""))throw new Error("RequestID migrasi tidak valid.");
  var sources=d.sources||(activeSegment_()==="B2C"?[]:LEGACY_SOURCE_DEFAULTS_);if(!Array.isArray(sources)||!sources.length||sources.length>2)throw new Error("Pilih satu atau dua database sumber.");
  var programs={};sources.forEach(function(x){if(programs[x.program])throw new Error("Program sumber ganda.");programs[x.program]=true;});
  ensureLegacySchema_();var signature=cacheKey_("",JSON.stringify(sources));
  var reserved=bulkWrites_(function(){var existing=uniqueMatch_(readAllAsObjects_(SHEET_IMPORTS_).filter(function(r){return r.data.OwnerNIK===s.nik&&r.data.RequestID===d.requestId;}),"request migrasi");if(existing){if(existing.data.Kind!=="LEGACY"||existing.data.FileHash!==signature)throw new Error("RequestID sudah dipakai tujuan lain.");if(existing.data.Status!=="PREVIEW"||timeValue_(existing.data.ExpiresAt)<Date.now())throw new Error("Migrasi sudah diproses/preview tidak aktif. Buka riwayat atau analisis baru.");return {existing:existing.data};}
    var id=newId_("MIG"),value={UploadID:id,OwnerNIK:s.nik,RequestID:d.requestId,Kind:"LEGACY",Program:"ALL",Batch:"AUTO",FileName:"Database peserta lama",FileHash:signature,Status:"PREPARING",CreatedAt:nowIso_(),UpdatedAt:nowIso_(),ExpiresAt:new Date(Date.now()+3600000).toISOString(),PreviewVersion:1,NextIndex:0,Written:0,Skipped:0,Failed:0,LeaseUntil:Date.now()+330000};appendObject_(getSheet_(SHEET_IMPORTS_),getHeaders_(getSheet_(SHEET_IMPORTS_)),value);return {id:id};});
  if(reserved.existing)return legacyPreviewResponse_(reserved.existing,importPayload_(reserved.existing.PayloadFileID));
  try{var raw={rows:[],errors:[],skipped:0};sources.forEach(function(x){var r=readLegacySource_(x);raw.rows=raw.rows.concat(r.rows);raw.errors=raw.errors.concat(r.errors);raw.skipped+=r.skipped;});if(raw.rows.length>2000)throw new Error("Maksimal 2.000 peserta per migrasi; pilih satu program/salinan sumber yang lebih kecil.");
    var payload={scope:{kind:"LEGACY",program:"ALL",batch:"AUTO"},raw:raw,resolutions:{},preview:buildLegacyPreview_(raw,{}),sources:sources},file=saveImportPayload_(reserved.id,payload),job=bulkWrites_(function(){requireRole_(requireSession_(token),["ADMIN"]);var r=requireImport_(s,reserved.id);updateRowByIndex_(getSheet_(SHEET_IMPORTS_),getHeaders_(getSheet_(SHEET_IMPORTS_)),r.rowIndex,{PayloadFileID:file,Status:"PREVIEW",LeaseUntil:"",ResultJSON:JSON.stringify({preview:payload.preview.summary,totalItems:payload.preview.items.length,outcomes:[]}),UpdatedAt:nowIso_()});return r.data;});return legacyPreviewResponse_(job,payload);
  }catch(error){bulkWrites_(function(){var r=requireImport_(s,reserved.id);updateRowByIndex_(getSheet_(SHEET_IMPORTS_),getHeaders_(getSheet_(SHEET_IMPORTS_)),r.rowIndex,{Status:"FAILED",LeaseUntil:"",ResultJSON:JSON.stringify({error:String(error.message||error)}),UpdatedAt:nowIso_()});});throw error;}
}, arguments);}
function getLegacyMigrationPreview(token,id){return api_("getLegacyMigrationPreview",function(){var s=requireSession_(token);requireRole_(s,["ADMIN"]);var job=requireImport_(s,id).data;if(job.Kind!=="LEGACY")throw new Error("Bukan migrasi peserta lama.");return legacyPreviewResponse_(job,importPayload_(job.PayloadFileID));}, arguments);}
function reviewLegacyMigration(token,id,version,resolutions){return api_("reviewLegacyMigration",function(){
  var s=requireSession_(token);requireRole_(s,["ADMIN"]);var job=requireImport_(s,id).data;if(job.Kind!=="LEGACY"||job.Status!=="PREVIEW"||timeValue_(job.ExpiresAt)<Date.now())throw new Error("Preview migrasi tidak aktif.");
  var payload=importPayload_(job.PayloadFileID);payload.resolutions=Object.assign({},payload.resolutions||{},resolutions||{});Object.keys(payload.resolutions).forEach(function(k){if(["KEEP_CURRENT","USE_SOURCE"].indexOf(payload.resolutions[k])<0)throw new Error("Pilihan pemeriksaan tidak valid.");});payload.preview=buildLegacyPreview_(payload.raw,payload.resolutions);var file=saveImportPayload_(id,payload);
  var updated=bulkWrites_(function(){requireRole_(requireSession_(token),["ADMIN"]);var r=requireImport_(s,id);if(r.data.Status!=="PREVIEW"||Number(r.data.PreviewVersion)!==Number(version))throw new Error("CONFLICT: Preview migrasi berubah.");updateRowByIndex_(getSheet_(SHEET_IMPORTS_),getHeaders_(getSheet_(SHEET_IMPORTS_)),r.rowIndex,{PayloadFileID:file,PreviewVersion:Number(version)+1,ResultJSON:JSON.stringify({preview:payload.preview.summary,totalItems:payload.preview.items.length,outcomes:[]}),UpdatedAt:nowIso_()});return r.data;});return legacyPreviewResponse_(updated,payload);
}, arguments);}
function applyLegacyMigrationChunk_(token,s,payload,job,start,lease){return bulkWrites_(function(){
  requireRole_(requireSession_(token),["ADMIN"]);var current=requireImport_(s,job.UploadID).data;if(current.Kind!=="LEGACY"||current.Status!=="PROCESSING"||current.LeaseToken!==lease||Number(current.NextIndex||0)!==start)throw new Error("CONFLICT: Posisi migrasi berubah.");
  var targets=legacyTargets_(),plans=[],outcomes=[],items=payload.preview.items.slice(start,start+20);
  items.forEach(function(item){try{var ps=targets.p[item.key]||[],es=targets.e[item.key]||[];if(ps.length>1||es.length>1||legacyTargetHash_(ps[0],es[0],targets.c[es[0]&&es[0].data.EnrollmentID]||[])!==item.expectedHash)throw new Error("Data berubah setelah preview. Analisis ulang baris ini.");plans.push({item:item,p:ps[0],e:es[0]});}catch(error){outcomes.push({nik:item.nik,status:"FAILED",message:error.message});}});
  var users={},certs={};readAllAsObjects_(SHEET_USERS).forEach(function(r){var n=normalizeNik_(r.data.NIK);if(users[n])throw new Error("Users duplikat: "+n);users[n]=r;});readAllAsObjects_(SHEET_CERTIFICATES).forEach(function(r){var id=r.data.EnrollmentID;if(id)(certs[id]||(certs[id]=[])).push(r);});
  var written=0,skipped=0;plans.forEach(function(plan){
    var item=plan.item,p=plan.p,e=plan.e,data=sourceRosterData_(item.data),useSource=item.resolution==="USE_SOURCE",now=nowIso_(),id=e&&e.data.EnrollmentID?e.data.EnrollmentID:newId_("ENR"),pp={},ep={},activeCert=(certs[id]||[])[0],fieldMap={FinalStatus:"ProgramStatus",BehaviourObservation:"BehaviourScore"};
    Object.keys(data).forEach(function(k){var incoming=data[k],actual=k==="CertificateUrl"&&activeCert&&!legacyBlank_(activeCert.data.FileOrLink)?activeCert.data.FileOrLink:legacyCurrent_(p,e,k);if(!legacyBlank_(incoming)&&(legacyBlank_(actual)||useSource&&!legacySame_(k,actual,incoming)))pp[k]=typeof incoming==="string"?safeCell_(incoming):incoming;});
    // Enrollment mirrors use the existing canonical record first, then fill missing values.
    ["NIK","Program","Batch","Region","CertificateUrl"].forEach(function(k){var ek=fieldMap[k]||k,v=k==="CertificateUrl"&&activeCert&&!legacyBlank_(activeCert.data.FileOrLink)?activeCert.data.FileOrLink:legacyCurrent_(p,e,k);if(legacyBlank_(v)||useSource&&!legacyBlank_(data[k]))v=data[k];if(!legacyBlank_(v)&&(!e||legacyBlank_(e.data[ek])||useSource&&!legacySame_(k,e.data[ek],v)))ep[ek]=v;});

    if(!p){pp=Object.assign({ParticipantStatus:"ELIGIBLE"},data,pp);Object.keys(pp).forEach(function(k){if(typeof pp[k]==="string")pp[k]=safeCell_(pp[k]);});pp.CurrentEnrollmentId=id;pp.UpdatedAt=now;}
    if(!e){ep=Object.assign({EnrollmentID:id,NIK:item.nik,Program:item.program,Batch:item.batch,Region:data.Region,CurrentPhase:"PHASE 1",EnrollmentStatus:"ON GOING",ProgramStatus:"ON GOING",IsActive:false,CreatedAt:now,UpdatedAt:now},ep);}
    var merged=Object.assign({},p?p.data:{},pp),me=Object.assign({},e?e.data:{},ep),snapshot=legacySnapshot_(me);
    if(!snapshot.version){snapshot={version:2,locked:false,mode:"SOURCE",values:{},overrides:{},metrics:item.metrics||[],fallbackValues:item.fallback||{},source:item.source,migratedAt:now,migratedBy:s.nik,needsSourceReset:true};ep.LegacySnapshot=JSON.stringify(snapshot);}
    var fallbackChanged=JSON.stringify(snapshot.fallbackValues||{})!==JSON.stringify(item.fallback||{});
    if(fallbackChanged){snapshot.fallbackValues=item.fallback||{};snapshot.metrics=item.metrics||[];ep.LegacySnapshot=JSON.stringify(snapshot);}
    var needsRegrade=fallbackChanged||snapshot.locked!==false||snapshot.mode!=="SOURCE"||snapshot.needsSourceReset;
    if(e&&!e.data.EnrollmentID)ep.EnrollmentID=id;
    if(merged.CurrentEnrollmentId!==id)pp.CurrentEnrollmentId=id;
    var changed=!p||!e||Object.keys(pp).length>0||Object.keys(ep).length>0;
    if(Object.keys(pp).length){pp.UpdatedAt=now;if(p)updateRowByIndex_(getSheet_(SHEET_PARTICIPANTS),getHeaders_(getSheet_(SHEET_PARTICIPANTS)),p.rowIndex,pp);else{var row=appendObject_(getSheet_(SHEET_PARTICIPANTS),getHeaders_(getSheet_(SHEET_PARTICIPANTS)),pp);p={rowIndex:row,data:Object.assign({},pp)};}}
    if(Object.keys(ep).length){ep.UpdatedAt=now;if(e)updateRowByIndex_(getSheet_(SHEET_ENROLLMENTS),getHeaders_(getSheet_(SHEET_ENROLLMENTS)),e.rowIndex,ep);else{var erow=appendObject_(getSheet_(SHEET_ENROLLMENTS),getHeaders_(getSheet_(SHEET_ENROLLMENTS)),ep);e={rowIndex:erow,data:Object.assign({},ep)};}}
    var u=users[item.nik];if(!u||!userHasRole_(u.data,"PARTICIPANT")){if(u){var roles=parseRoles_(u.data.Role);roles.push("PARTICIPANT");updateRowByIndex_(getSheet_(SHEET_USERS),getHeaders_(getSheet_(SHEET_USERS)),u.rowIndex,{Role:joinRoles_(roles),UpdatedAt:now});if(item.nik===s.nik){var live=readAllAsObjects_(SHEET_USERS).filter(function(r){return normalizeNik_(r.data.NIK)===s.nik;})[0];s.authVersion=String(live.data.AuthVersion);s.roles=parseRoles_(live.data.Role).filter(isRoleEnabled_);CacheService.getScriptCache().put("sess_"+token,JSON.stringify(s),Math.max(1,Math.ceil((s.expiresAt-Date.now())/1000)));}}else ensureParticipantLogin_(merged);users[item.nik]={data:Object.assign({},u?u.data:{},{Role:joinRoles_(parseRoles_(u&&u.data.Role).concat(["PARTICIPANT"]))})};changed=true;}
    var url=rawText_(me.CertificateUrl||merged.CertificateUrl);if(url&&!(certs[id]||[]).length){var c={CertificateID:newId_("CERT"),EnrollmentID:id,NIK:item.nik,Program:item.program,Batch:item.batch,DisplayName:merged.Name,FinalScore:me.FinalScore,FinalStatus:me.ProgramStatus,FileOrLink:url,Published:false,UploadedBy:s.nik,UploadedAt:now};appendObject_(getSheet_(SHEET_CERTIFICATES),getHeaders_(getSheet_(SHEET_CERTIFICATES)),c);certs[id]=[{data:c}];changed=true;}
    if(useSource&&item.conflicts.length)auditLog_(s.nik,s.role,"LEGACY_SOURCE_SELECTED","Enrollment",id,JSON.stringify({source:item.source,changes:item.conflicts}));
    if(useSource&&data.CertificateUrl&&(certs[id]||[]).length===1){var cr=certs[id][0];if(cr.rowIndex&&(cr.data.FileOrLink!==data.CertificateUrl||!legacySame_("FinalScore",cr.data.FinalScore,me.FinalScore)||cr.data.FinalStatus!==me.ProgramStatus)){updateRowByIndex_(getSheet_(SHEET_CERTIFICATES),getHeaders_(getSheet_(SHEET_CERTIFICATES)),cr.rowIndex,{FileOrLink:data.CertificateUrl,FinalScore:me.FinalScore,FinalStatus:me.ProgramStatus,Published:false,UploadedBy:s.nik,UploadedAt:now});changed=true;}}
    if(needsRegrade){Object.assign(p.data,pp);Object.assign(e.data,ep);var target={participant:p,enrollment:e,updatedBy:s.nik};refreshEnrollmentSummary_(target);changed=true;}
    if(changed){written++;outcomes.push({nik:item.nik,status:"WRITTEN",message:"Identitas disimpan. Nilai dan hasil mengikuti sumber tiap tahap."});}else{skipped++;outcomes.push({nik:item.nik,status:"SKIPPED",message:"Data sudah tersimpan; tidak ada perubahan."});}
  });
  auditLog_(s.nik,s.role,"MIGRATE_LEGACY_PARTICIPANTS","Import",job.UploadID,JSON.stringify({start:start,written:written,skipped:skipped,failed:outcomes.filter(function(x){return x.status==="FAILED";}).length}));
  return {next:start+items.length,written:written,skipped:skipped,failed:outcomes.filter(function(x){return x.status==="FAILED";}).length,outcomes:outcomes};
});}
function protectLegacySummary_(t,patch,summary){
  var hist=legacySnapshot_(t.enrollment.data);if(!hist.version||hist.locked===false)return false;
  var fields=["Phase1Score","Phase2Score","Phase3Score","BehaviourScore","PanelScore"],diffs=[];hist.overrides=hist.overrides||{};
  fields.forEach(function(k,i){
    var ph="PHASE "+(i+1),completed=summary[ph]&&summary[ph].completedModules===summary[ph].totalModules&&summary[ph].totalModules>0;
    if(i<3&&completed&&t.manualPhase===ph&&Object.prototype.hasOwnProperty.call(patch,k))hist.overrides[k]=patch[k];
    var frozen=Object.prototype.hasOwnProperty.call(hist.overrides,k)?hist.overrides[k]:hist.values[k];
    if(frozen!==undefined){if((i>=3||completed)&&patch[k]!==undefined&&patch[k]!==""&&!legacySame_(k,patch[k],frozen))diffs.push({phase:i<3?"PHASE "+(i+1):i===3?"Behavior":"Panel",historical:frozen,source:patch[k]});patch[k]=frozen;}
  });
  hist.differences=diffs;patch.LegacySnapshot=JSON.stringify(hist);patch.FinalScore=hist.values.FinalScore;patch.ProgramStatus=hist.values.ProgramStatus||"ON GOING";
  patch.CurrentPhase=t.enrollment.data.CurrentPhase;patch.EnrollmentStatus=t.enrollment.data.EnrollmentStatus;patch.CompletedAt=t.enrollment.data.CompletedAt||"";
  patch.FinalSummary=JSON.stringify({historical:true,locked:true,score:patch.FinalScore,source:hist.source,updatedAt:nowIso_()});return true;
}
function legacyDashboard_(out,e,p){var h=legacySnapshot_(e);if(!h.version)return;out.historicalInfo={locked:h.locked!==false,metrics:h.metrics||[],differences:h.differences||[],migratedAt:h.migratedAt};if(h.locked===false)return;["Phase1Score","Phase2Score","Phase3Score"].forEach(function(k,i){var v=Object.prototype.hasOwnProperty.call(h.overrides||{},k)?h.overrides[k]:h.values[k];if(v!==undefined){out["phase"+(i+1)].score=v;out["phase"+(i+1)].available=v!=="";}});if(h.values.BehaviourScore!==undefined){out.behaviourObservation=h.values.BehaviourScore;out.behaviourScale=Number(h.values.BehaviourScore)>5?100:5;}if(h.values.PanelScore!==undefined)out.panelScore=h.values.PanelScore;out.finalScore=h.values.FinalScore;out.programStatus=out.finalStatus=h.values.ProgramStatus;out.result=computeResult_({ProgramStatus:h.values.ProgramStatus,ParticipantStatus:p&&p.ParticipantStatus});}
function legacyAssessment_(a,e){var h=legacySnapshot_(e);if(!h.version||h.locked===false)return;var k="Phase"+a.phase.replace(/\D/g,"")+"Score",v=Object.prototype.hasOwnProperty.call(h.overrides||{},k)?h.overrides[k]:h.values[k];if(v===undefined)return;a.sourceAverageScore=a.averageScore;a.historicalScore=v;a.averageScore=a.finalScore=v;a.historicalOnly=!a.available;}
function getLegacyMigrationHistory(token){return api_("getLegacyMigrationHistory",function(){var s=requireSession_(token);requireRole_(s,["ADMIN"]);return getRawImportHistory_(token).filter(function(j){return j.kind==="LEGACY";});}, arguments);}
function releaseHistoricalGrading(token,d){return api_("releaseHistoricalGrading",function(){var s=requireSession_(token);requireRole_(s,["ADMIN"]);return bulkWrites_(function(){
  var t=authorizeEnrollment_(s,d),e=t.enrollment.data,h=legacySnapshot_(e);if(!h.version)throw new Error("Enrollment ini tidak memiliki data historis.");if(!d.recordVersion||String(d.recordVersion)!==String(e.RecordVersion))throw new Error("CONFLICT: Data berubah; buka ulang detail peserta.");
  h.locked=false;h.releasedAt=nowIso_();h.releasedBy=s.nik;var value=JSON.stringify(h);updateRowByIndex_(getSheet_(SHEET_ENROLLMENTS),getHeaders_(getSheet_(SHEET_ENROLLMENTS)),t.enrollment.rowIndex,{LegacySnapshot:value,UpdatedAt:nowIso_()});t.enrollment.data.LegacySnapshot=value;t.updatedBy=s.nik;refreshEnrollmentSummary_(t);auditLog_(s.nik,s.role,"RELEASE_HISTORICAL_GRADING","Enrollment",e.EnrollmentID,"Explicit Admin switch to current grading rules");return {ok:true,message:"Penilaian sekarang memakai aturan berjalan. Snapshot historis tetap disimpan."};
});}, arguments);}
function legacyObservationScore_(e,key){var h=legacySnapshot_(e);return h.version&&h.locked!==false&&h.values[key]!==undefined?h.values[key]:null;}

// Raw workbooks, not master totals, are the grading authority.
var RAW_SOURCE_DEFAULTS_ = {
  TSHDP: [
    {id:"p1",label:"Phase 1 · per kompetensi",kind:"SYLLABUS",phase:"PHASE 1",url:"https://docs.google.com/spreadsheets/d/1jsa5H2EE2x9ivfhrDhPoFFHmiXwdRG33/edit"},
    {id:"p2",label:"Phase 2 · per kompetensi",kind:"SYLLABUS",phase:"PHASE 2",url:"https://docs.google.com/spreadsheets/d/1B6OI9QfTfcjRJKJSN-6OywA4c7s8fQthW-e6gJPskFg/edit"},
    {id:"p3",label:"Phase 3 · Soft Competency",kind:"SYLLABUS",phase:"PHASE 3",url:"https://docs.google.com/spreadsheets/d/1riAVblI5r7XpHAZJSABYXslLcNc1WpcVQWDFpcMbmUY/edit"},
    {id:"behavior",label:"Behavior Monitoring",kind:"BEHAVIOR",phase:"",url:"https://docs.google.com/spreadsheets/d/11qMLOSyRnXF4-Wyj4YcbZgUq5gIFOXG1aVDHzPhgTxQ/edit"},
    {id:"panel",label:"Panel Score",kind:"PANEL",phase:"",url:"https://docs.google.com/spreadsheets/d/1X43AUHWg22_9tsPPmeMFLzmePSRJTW5_fJXkGNiiKK0/edit"}
  ],
  ASHDP: [
    {id:"p1",label:"Phase 1 · per kompetensi",kind:"SYLLABUS",phase:"PHASE 1",url:"https://docs.google.com/spreadsheets/d/1ReCs7MP3YAGyf31G_asMhyumUhUKufP2/edit"},
    {id:"p2",label:"Phase 2 · per kompetensi",kind:"SYLLABUS",phase:"PHASE 2",url:"https://docs.google.com/spreadsheets/d/16QvHc8YSiA6QiawV5qUZB2dI0QMVPabI/edit"},
    {id:"p3",label:"Phase 3 · Soft Competency",kind:"SYLLABUS",phase:"PHASE 3",url:"https://docs.google.com/spreadsheets/d/1vz0h5cz4YJ39TYHxEWem8MlgN-xwHYWJ/edit"},
    {id:"behavior",label:"Behavior Monitoring",kind:"BEHAVIOR",phase:"",url:"https://docs.google.com/spreadsheets/d/1u-RFl8vzb7LrdMxIB8yp-UhM6EVUe7UYSdcL_zCSaRo/edit"},
    {id:"panel",label:"Panel Score",kind:"PANEL",phase:"",url:"https://docs.google.com/spreadsheets/d/1vfB1h-o7KO0V8m2TzRgjYtCI7pSGCvQyFcT2n8KVC2I/edit"}
  ]
};
function sourceRosterData_(data){
  var result={},allowed=["Segment","NIK","Name","Program","Batch","Region","JobTitle","CityName","Email","ParticipantStatus","ProgramTrack","CertificateUrl","Theme","ProjectTitle","ActionPlan","Implementation","SlideUrl","PanelDate"];
  allowed.forEach(function(k){if(Object.prototype.hasOwnProperty.call(data,k))result[k]=data[k];});return result;
}
function parseLegacyIdentity_(rows,program,reference){
  // Mask grade columns before parsing so broken old formulas or stale status cannot
  // block roster import or silently become current grades.
  var ignored=["Final Score","FinalScore","Final Status","FinalStatus","Panel Score","PanelScore","Panel","Behaviours Monitoring","Behaviours Observation","BehaviourObservation","Observasi","Score Phase 1","Score Phase 2","Score Phase 3","Phase1Score","Phase2Score","Phase3Score","Post Test Average Score","In Class Training Score","In Class Training","Avarage Score Phase 1 - Self Learning","Average Score Phase 1 - Self Learning","Avarage Score Phase 2 - Technical Competency","Average Score Phase 2 - Technical Competency","Avarage Score Phase 3 - Supervisory Development Program","Average Score Phase 3 - Supervisory Development Program"].map(headerKey_);
  var clean=rows.map(function(row){return findHeaderIndex_(row,["NIK","Employee ID"])>=0&&findHeaderIndex_(row,["Batch"])>=0?row.map(function(v){return ignored.indexOf(headerKey_(v))>=0?"":v;}):row;});
  var parsed=parseLegacyRosterBase_(clean,program,reference);
  parsed.rows.forEach(function(r){
    r.data=sourceRosterData_(r.data);r.metrics=[];
    var header=rows[r.source.headerRow-1]||[],at=findHeaderIndex_(header,["Final Status","FinalStatus"]),status=headerKey_((rows[r.source.row-1]||[])[at]);
    if(status==="resign"||status==="resigned")r.data.ParticipantStatus="RESIGN";
  });return parsed;
}
function activateSourceGrading_(t){
  var e=t.enrollment.data,p=t.participant.data,h=legacySnapshot_(e);
  if(!h.version||(h.mode==="SOURCE"&&h.locked===false&&!h.needsSourceReset))return false;
  // Preserve the old snapshot only as an archive; no source workbook is modified.
  h.originalValues=h.originalValues||Object.assign({},h.values||{});
  if(!h.fallbackValues&&h.values)h.fallbackValues=Object.assign({},h.values,{FinalStatus:h.values.ProgramStatus||""});
  h.version=2;h.locked=false;h.mode="SOURCE";h.activatedAt=nowIso_();delete h.needsSourceReset;
  ["Phase1Score","Phase2Score","Phase3Score","FinalScore"].forEach(function(k){e[k]="";p[k]="";});
  ["BEHAVIOR","PANEL"].forEach(function(kind){
    var name=kind==="PANEL"?SHEET_PANEL_OBSERVATIONS_:SHEET_BEHAVIOR_,key=kind==="PANEL"?"Panel":"Behaviour",old=parseJson_(e[key+"Summary"],{}),policy=["LATEST","RATER_LATEST","ALL_MEAN"].indexOf(old.policy)>=0?old.policy:kind==="PANEL"?"ALL_MEAN":"LATEST";
    var records=readAllAsObjects_(name).filter(function(r){return r.data.EnrollmentID===e.EnrollmentID;}).map(function(r){return r.data;});
    var sum=records.length?behaviorAggregate_(records,policy):null;
    e[key+"Score"]=sum?(kind==="PANEL"?Math.round(sum.score*2000)/100:sum.score):"";e[key+"Summary"]=sum?JSON.stringify(sum):"";
    if(kind==="BEHAVIOR"){e.BehaviourCoreScore=sum?sum.core:"";e.BehaviourTechnicalScore=sum?sum.technical:"";p.BehaviourObservation=e.BehaviourScore;}else p.PanelScore=e.PanelScore;
  });
  e.ProgramStatus="ON GOING";p.FinalStatus="ON GOING";e.LegacySnapshot=JSON.stringify(h);t.sourceTransition=true;return true;
}
function sourceSpreadsheetId_(url){
  var match=rawText_(url).match(/^https:\/\/docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]{20,100})(?:[/?#]|$)/);
  if(!match)throw new Error("Gunakan link Google Sheets atau workbook Excel di Google Drive.");
  if(match[1]===props_().getProperty("LDP_SPREADSHEET_ID"))throw new Error("Pilih file mentah tahap, bukan database aplikasi.");return match[1];
}
function readRawSourceWorkbook_(d){
  var id=sourceSpreadsheetId_(d.sourceUrl),ss;
  try {ss=SpreadsheetApp.openById(id);}catch(error){
    var file=DriveApp.getFileById(id),blob=file.getBlob(),mime=blob.getContentType();
    if(!/spreadsheetml|ms-excel/.test(mime))throw error;
    if(blob.getBytes().length>5*1024*1024)throw new Error("File sumber maksimal 5 MB.");
    return parseRawWorkbook_(Object.assign({},d,{fileName:file.getName(),mimeType:mime,base64:Utilities.base64Encode(blob.getBytes())}));
  }
  var title=typeof ss.getName==="function"?String(ss.getName()).toUpperCase():"",detected=activePrograms_().filter(function(p){return title.indexOf(p)>=0;});
  if(detected.length>1||detected.length===1&&detected[0]!==d.program)throw new Error("Program file sumber berbeda dari tujuan.");
  var sheets=ss.getSheets(),sh=null,gid=rawText_(d.sourceUrl).match(/[?#&]gid=(\d+)/);
  if(gid){sh=sheets.filter(function(x){return String(x.getSheetId())===gid[1];})[0];if(!sh)throw new Error("Tab gid sumber tidak ditemukan.");}
  if(!sh){
    var preferred=d.kind==="SYLLABUS"?(d.program==="TSHDP"&&d.phase==="PHASE 1"?["rekapnilai"]:[headerKey_(d.phase)+headerKey_(d.program),headerKey_(d.phase)]):["formresponses1","recap","bm"];
    preferred.some(function(k){sh=sheets.filter(function(x){return headerKey_(x.getName())===k;})[0];return !!sh;});
    if(!sh&&sheets.length===1)sh=sheets[0];
    if(!sh)throw new Error("Tab raw tahap belum dikenali. Gunakan link tab dengan gid atau upload file mentah.");
  }
  var nr=sh.getLastRow(),nc=sh.getLastColumn();if(nr>5003||nc>250||nr*nc>500000)throw new Error("Tab sumber terlalu besar; maksimal 5.000 baris data, 250 kolom dan 500.000 sel.");
  return {sheetName:sh.getName(),rows:nr&&nc?sh.getRange(1,1,nr,nc).getValues():[]};
}

// Read-only plans keep all source selection and matching rules in the existing import pipeline.
function getLegacySyncPlan(token,program){return api_("getLegacySyncPlan",function(){
  if(activeSegment_()==="B2C")throw new Error("Link arsip lama B2C belum ditambahkan. Gunakan Tarik Data Lama untuk memasukkan link B2C atau Upload & Rekap untuk file mentah.");
  var s=requireSession_(token);requireRole_(s,["ADMIN"]);program=rawText_(program).toUpperCase();
  if(program&&activePrograms_().indexOf(program)<0)throw new Error("Program sinkron tidak valid.");
  var programs=program?[program]:activePrograms_().slice(),sources=(activeSegment_()==="B2C"?[]:LEGACY_SOURCE_DEFAULTS_).filter(function(x){return programs.indexOf(x.program)>=0;}),steps=[{id:"roster",kind:"ROSTER",label:"Identitas peserta dan batch",sources:sources}];
  programs.forEach(function(p){RAW_SOURCE_DEFAULTS_[p].forEach(function(x){steps.push({id:p+"_"+x.id,kind:x.kind,phase:x.phase,program:p,label:p+" · "+(x.id==="p3"?"Soft Competency":x.label),sourceUrl:x.url});});});
  return {owner:s.nik,program:program||"ALL",key:cacheKey_("",JSON.stringify(steps)),steps:steps};
}, arguments);}
function getLegacySyncRequest(token,requestId){return api_("getLegacySyncRequest",function(){
  var s=requireSession_(token);requireRole_(s,["ADMIN"]);
  if(!/^[A-Za-z0-9_-]{12,100}$/.test(requestId||""))throw new Error("Request sinkron tidak valid.");
  var job=uniqueMatch_(readAllAsObjects_(SHEET_IMPORTS_).filter(function(r){return r.data.OwnerNIK===s.nik&&r.data.RequestID===requestId;}),"request sinkron");
  if(!job)return null;
  var response=importJobResponse_(job.data);response.leaseUntil=Number(job.data.LeaseUntil||0);return response;
}, arguments);}

// Historic totals fill only stages with no competency/observation data of their own.
function parseLegacyMaster_(rows,program,reference){
  var parsed=parseLegacyIdentity_(rows,program,reference);
  parsed.rows.forEach(function(r){
    var h=rows[r.source.headerRow-1]||[],row=rows[r.source.row-1]||[];
    function value(aliases,last){var at=[];h.forEach(function(x,i){if(aliases.map(headerKey_).indexOf(headerKey_(x))>=0)at.push(i);});var col=last?at[at.length-1]:at.filter(function(i){return !legacyBlank_(row[i]);})[0];return col===undefined?"":row[col];}
    function number(aliases,last){try{return legacyNumber_(value(aliases,last));}catch(error){return "";}}
    var f={FinalScore:number(["Final Score","FinalScore"]),PanelScore:number(["Panel Score","PanelScore","Panel"],true),BehaviourScore:number(["Behaviours Monitoring","Behaviours Observation","BehaviourObservation","Observasi"])};
    try{f.FinalStatus=legacyStatus_(value(["Final Status","FinalStatus"]));}catch(error){f.FinalStatus="";}
    f.Phase1Score=number(["Score Phase 1","Phase1Score","Avarage Score Phase 1 - Self Learning","Average Score Phase 1 - Self Learning"]);
    f.Phase2Score=number(["Score Phase 2","Phase2Score","Avarage Score Phase 2 - Technical Competency","Average Score Phase 2 - Technical Competency"]);
    f.Phase3Score=number(["Score Phase 3","Phase3Score","Avarage Score Phase 3 - Supervisory Development Program","Average Score Phase 3 - Supervisory Development Program"]);
    r.metrics=[];["Post Test Average Score","In Class Training Score","In Class Training"].forEach(function(label){var score=number([label]);if(score!=="")r.metrics.push({label:label,score:score});});
    r.fallback=f;
  });return parsed;
}
function applyParticipantFallbacks_(t,patch,summary){
  var e=t.enrollment.data,h=legacySnapshot_(e),f=h.fallbackValues;if(!f)return;
  ["PHASE 1","PHASE 2","PHASE 3"].forEach(function(ph,i){
    var a=summary[ph]||{},key="Phase"+(i+1)+"Score";a.aggregateFallback=false;
    if(!a.completedModules&&numberOrBlank_(f[key])!==""){patch[key]=numberOrBlank_(f[key]);a.aggregateFallback=true;}
    summary[ph]=a;
  });
  ["Behaviour","Panel"].forEach(function(prefix){
    var key=prefix+"Score",sk=prefix+"Summary",current=Object.assign({},e,patch),sum=parseJson_(current[sk],{});
    if((numberOrBlank_(current[key])===""||sum.aggregateFallback)&&numberOrBlank_(f[key])!==""){
      patch[key]=numberOrBlank_(f[key]);patch[sk]=JSON.stringify({score:patch[key],scale:100,aggregateFallback:true,source:"PARTICIPANT",observations:0,usedObservations:0});
    }else if(sum.aggregateFallback&&numberOrBlank_(f[key])===""){patch[key]="";patch[sk]="";}
  });
}
function resolveParticipantFinal_(t,patch,summary,grade){
  var e=Object.assign({},t.enrollment.data,patch),h=legacySnapshot_(e),f=h.fallbackValues;if(!f)return grade;
  var allRawAbsent=["PHASE 1","PHASE 2","PHASE 3"].every(function(ph){return !Number((summary[ph]||{}).completedModules||0);})&&["Behaviour","Panel"].every(function(k){var x=parseJson_(e[k+"Summary"],{});return x.aggregateFallback||!Number(x.observations||0);});
  grade.components.forEach(function(c,i){c.source=i<3?(summary["PHASE "+(i+1)]||{}).aggregateFallback?"PARTICIPANT":"RAW":parseJson_(e[(i===3?"Behaviour":"Panel")+"Summary"],{}).aggregateFallback?"PARTICIPANT":"RAW";});
  if(allRawAbsent&&numberOrBlank_(f.FinalScore)!==""&&f.FinalStatus!=="ON GOING"&&(Number(f.FinalScore)>0||["LULUS","TIDAK LULUS"].indexOf(f.FinalStatus)>=0)){grade.detailsMissing=grade.missing;grade.missing=[];grade.score=numberOrBlank_(f.FinalScore);grade.complete=true;grade.source="PARTICIPANT";grade.aggregateFallback=true;}
  else if(allRawAbsent&&f.FinalStatus==="ON GOING"){grade.score="";grade.complete=false;grade.source="PARTICIPANT";}
  return grade;
}
function participantFallbackAssessment_(a,e){
  var h=legacySnapshot_(e),f=h.fallbackValues||{},key="Phase"+String(a.phase).replace(/\D/g,"")+"Score";
  if(!a.completedModules&&numberOrBlank_(f[key])!==""){a.averageScore=a.finalScore=numberOrBlank_(f[key]);a.available=true;a.participantFallback=true;a.message="Nilai total dari Participant; rincian kompetensi belum tersedia untuk batch ini.";}
}
function participantObservationFallback_(e,prefix){var s=parseJson_(e[prefix+"Summary"],{});return s.aggregateFallback?numberOrBlank_(e[prefix+"Score"]):null;}
function participantFallbackDashboard_(out,e,p){
  if(!legacySnapshot_(e).fallbackValues)return;
  var sum=parseJson_(e.AssessmentSummary,{}),b=parseJson_(e.BehaviourSummary,{}),fs=parseJson_(e.FinalSummary,{});
  if(b.aggregateFallback)out.behaviourScale=100;
  if(fs.aggregateFallback&&fs.source==="PARTICIPANT"&&!out.assessmentError&&![1,2,3].some(function(i){return Number(out["phase"+i].rawCompleted||0)>0;}))out.finalScore=e.FinalScore;
  [1,2,3].forEach(function(i){var a=sum["PHASE "+i]||{},target=out["phase"+i];if(target.rawCompleted===undefined)target.participantFallback=!!a.aggregateFallback;});
}

function dashboardSyncRows_(d){
  d=d||{};var program=rawText_(d.program).toUpperCase(),ids=d.ids||[];
  if(program&&activePrograms_().indexOf(program)<0)throw new Error("Program tidak valid.");
  if(!Array.isArray(ids)||ids.length>5000)throw new Error("Daftar peserta tidak valid.");
  var wanted={},known={};ids.forEach(function(id){wanted[String(id)]=true;});
  readAllAsObjects_(SHEET_ENROLLMENTS).forEach(function(r){if(wanted[r.data.EnrollmentID])known[entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment)]=true;});
  var rows=readAllAsObjects_(SHEET_PARTICIPANTS).filter(function(r){var p=r.data;return (!program||String(p.Program).toUpperCase()===program)&&(!d.batch||sameBatch_(p.Batch,d.batch))&&(!d.region||normalizeRegion_(p.Region)===normalizeRegion_(d.region))&&matchesProgramTrack_(p.ProgramTrack,d.track)&&(!ids.length||known[entityKey_(p.NIK,p.Program,p.Batch, p.Segment)]);});
  rows.sort(function(a,b){return a.rowIndex-b.rowIndex;});
  return {rows:rows,key:cacheKey_("",JSON.stringify(rows.map(function(r){return [r.rowIndex,entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment)];})))};
}
function getDashboardScoreSyncPlan(token,d){return api_("getDashboardScoreSyncPlan",function(){
  var s=requireSession_(token);requireRole_(s,["ADMIN"]);context_().memo.bypassSnapshots=true;var plan=dashboardSyncRows_(d);return {key:plan.key,total:plan.rows.length,chunkSize:20};
}, arguments);}
function dashboardParticipantValues_(p){
  var out={};["Phase1Score","Phase2Score","Phase3Score","FinalScore","PanelScore"].forEach(function(k){out[k]=legacyNumber_(p[k]);});
  out.BehaviourScore=legacyNumber_(p.BehaviourObservation);out.FinalStatus="";
  try{out.FinalStatus=legacyStatus_(p.FinalStatus);}catch(error){}
  return out;
}
function dashboardObservation_(records,prefix,old){
  if(!records.length)return null;
  records.forEach(function(r){var n=numberOrBlank_(r.AverageScore);if(n===""||n<1||n>5)throw new Error("Nilai "+prefix+" pada sheet observasi tidak valid.");});
  var policy=["LATEST","RATER_LATEST","ALL_MEAN"].indexOf(old.policy)>=0?old.policy:prefix==="Panel"?"ALL_MEAN":"LATEST",sum=behaviorAggregate_(records,policy);
  return {score:prefix==="Panel"?Math.round(sum.score*2000)/100:sum.score,summary:sum};
}
function syncDashboardScores(token,d){return api_("syncDashboardScores",function(){
  var s=requireSession_(token);requireRole_(s,["ADMIN"]);d=d||{};
  if(!/^[A-Za-z0-9_-]{12,100}$/.test(d.requestId||""))throw new Error("Request sinkron tidak valid.");
  if(!Number.isInteger(Number(d.cursor))||Number(d.cursor)<0)throw new Error("Posisi sinkron tidak valid.");
  ensureLegacySchema_();var receipt=null,result=bulkWrites_(function(){
    var c=context_();c.memo.bypassSnapshots=true;c.memo.liveAssessmentReads=true;
    var signature=cacheKey_("",JSON.stringify([d.key,d.cursor,d.program||"",d.region||"",d.batch||"",normalizeProgramTrack_(d.track),d.ids||[]])),cache=CacheService.getScriptCache(),cacheKey=cacheKey_("ldplocalsync:",s.nik+"|"+d.requestId),cached=cache.get(cacheKey);
    if(cached){cached=JSON.parse(cached);if(cached.signature!==signature)throw new Error("RequestID sudah dipakai tujuan lain.");return cached.result;}
    var plan=dashboardSyncRows_(d);if(plan.key!==d.key)throw new Error("CONFLICT: Daftar peserta berubah. Klik Sinkron Nilai kembali.");
    var start=Number(d.cursor),items=plan.rows.slice(start,start+20),all=readAllAsObjects_(SHEET_PARTICIPANTS),pKeys={},nikCounts={},ers={},outcomes=[],written=0,skipped=0;
    all.forEach(function(r){var k=entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment);pKeys[k]=(pKeys[k]||0)+1;nikCounts[normalizeNik_(r.data.NIK)]=(nikCounts[normalizeNik_(r.data.NIK)]||0)+1;});
    readAllAsObjects_(SHEET_ENROLLMENTS).forEach(function(r){var k=entityKey_(r.data.NIK,r.data.Program,r.data.Batch, r.data.Segment);(ers[k]||(ers[k]=[])).push(r);});
    var behavior=readAllAsObjects_(SHEET_BEHAVIOR_),panels=readAllAsObjects_(SHEET_PANEL_OBSERVATIONS_),legacyPanels=readAllAsObjects_(SHEET_PANEL);
    items.forEach(function(pRow){try{
      var p=pRow.data,k=entityKey_(p.NIK,p.Program,p.Batch, p.Segment),matches=ers[k]||[],er=matches[0],old=er?er.data:{},now=nowIso_();
      if(pKeys[k]>1||matches.length>1)throw new Error("Identitas NIK/program/batch duplikat; tidak ditimpa.");
      if(!normalizeNik_(p.NIK)||activePrograms_().indexOf(String(p.Program).toUpperCase())<0||!normalizeBatchKey_(p.Batch))throw new Error("NIK, program atau batch belum lengkap.");
      var local=dashboardParticipantValues_(p),h=legacySnapshot_(old),previousFallback=h.fallbackValues||{},f=Object.assign({},previousFallback),patch={},sum={},id=old.EnrollmentID||newId_("ENR");
      ["Phase1Score","Phase2Score","Phase3Score","FinalScore","PanelScore","BehaviourScore"].forEach(function(key){if(numberOrBlank_(f[key])===""&&numberOrBlank_(old[key])!=="")f[key]=numberOrBlank_(old[key]);});
      if(!f.FinalStatus)try{f.FinalStatus=legacyStatus_(old.ProgramStatus);}catch(ignore){}
      Object.keys(local).forEach(function(key){if(!legacyBlank_(local[key]))f[key]=local[key];});
      var componentKeys=["Phase1Score","Phase2Score","Phase3Score","BehaviourScore","PanelScore"],hasLocalComponents=componentKeys.some(function(key){return numberOrBlank_(local[key])!=="";});
      if(numberOrBlank_(local.FinalScore)===""&&hasLocalComponents){f.FinalScore="";if(componentKeys.every(function(key){return numberOrBlank_(local[key])!=="";})&&componentKeys.some(function(key){return Number(local[key])>0;}))f.FinalStatus="";}
      if(!h.localParticipantOriginal){h.localParticipantOriginal={capturedAt:now,row:pRow.rowIndex,values:local,participantStatus:p.ParticipantStatus||"",originalEnrollment:{Phase1Score:old.Phase1Score,Phase2Score:old.Phase2Score,Phase3Score:old.Phase3Score,BehaviourScore:old.BehaviourScore,PanelScore:old.PanelScore,FinalScore:old.FinalScore,ProgramStatus:old.ProgramStatus}};}
      h.version=3;h.locked=false;h.mode="SOURCE";h.fallbackValues=f;h.localSource="PARTICIPANTS";
      ["PHASE 1","PHASE 2","PHASE 3"].forEach(function(ph,i){
        var a=readAssessmentDetail_(p.NIK,ph,p.Program,p.Batch),key="Phase"+(i+1)+"Score",fallback=!a.completedModules&&numberOrBlank_(f[key])!=="";
        patch[key]=fallback?numberOrBlank_(f[key]):a.averageScore;
        sum[ph]=Object.assign({},(parseJson_(old.AssessmentSummary,{})[ph]||{}),{completedModules:a.completedModules,totalModules:a.totalModules,aggregateFallback:fallback,remedials:a.modules.filter(function(m){return m.assessed!==false&&m.result==="REMEDIAL REQUIRED";})});
      });
      ["Behaviour","Panel"].forEach(function(prefix){
        var rows=(prefix==="Panel"?panels:behavior).filter(function(r){return r.data.EnrollmentID===id||!r.data.EnrollmentID&&normalizeNik_(r.data.NIK)===normalizeNik_(p.NIK)&&nikCounts[normalizeNik_(p.NIK)]===1;}).map(function(r){return r.data;}),result=dashboardObservation_(rows,prefix,parseJson_(old[prefix+"Summary"],{}));
        if(!result&&prefix==="Panel"){
          var legacy=legacyPanels.filter(function(r){return normalizeNik_(r.data.NIK)===normalizeNik_(p.NIK)&&nikCounts[normalizeNik_(p.NIK)]===1&&numberOrBlank_(r.data.TotalScore)!=="";});
          if(legacy.length){var byRater={};legacy.forEach(function(r){var key=r.data.AssessorNIK||r.data.AssessmentID||r.rowIndex;if(!byRater[key]||timeValue_(r.data.UpdatedAt||r.data.CreatedAt)>=timeValue_(byRater[key].UpdatedAt||byRater[key].CreatedAt))byRater[key]=r.data;});var values=Object.keys(byRater).map(function(key){return legacyNumber_(byRater[key].TotalScore);}),score=values.reduce(function(a,b){return a+b;},0)/values.length;result={score:Math.round(score*100)/100,summary:{score:score/20,observations:values.length,usedObservations:values.length,policy:"RATER_LATEST",source:"PANEL_ASSESSMENTS"}};}
        }
        var key=prefix+"Score";patch[key]=result?result.score:numberOrBlank_(f[key]);patch[prefix+"Summary"]=JSON.stringify(result?result.summary:patch[key]!==""?{score:patch[key],scale:100,aggregateFallback:true,source:"PARTICIPANT",observations:0,usedObservations:0}:{});
      });
      patch.AssessmentSummary=JSON.stringify(sum);patch.LegacySnapshot=JSON.stringify(h);
      var t={participant:pRow,enrollment:{data:Object.assign({},old,patch)}},grade=resolveParticipantFinal_(t,patch,sum,finalGrade_(Object.assign({},old,patch),p,sum)),stopped=normalizeParticipantStatus_(p.ParticipantStatus)==="RESIGN"||headerKey_(p.FinalStatus)==="resign";
      if(stopped&&numberOrBlank_(local.FinalScore)!==""){grade.score=local.FinalScore;grade.source="PARTICIPANT";}
      patch.FinalScore=grade.score;patch.ProgramStatus=stopped?"RESIGN":grade.complete?(grade.score>=80?"LULUS":"TIDAK LULUS"):"ON GOING";patch.FinalSummary=JSON.stringify(grade);
      var incomplete=["PHASE 1","PHASE 2","PHASE 3"].filter(function(ph){var a=sum[ph];return !a.aggregateFallback&&a.completedModules<a.totalModules;});
      patch.CurrentPhase=stopped?(old.CurrentPhase||p.CurrentPhase||"PHASE 1"):grade.complete?"COMPLETED":incomplete[0]||"PANELISASI";
      patch.EnrollmentStatus=stopped?(old.EnrollmentStatus||"ON GOING"):grade.complete?"COMPLETED":"ON GOING";patch.CompletedAt=grade.complete&&!stopped?(old.CompletedAt||now):old.CompletedAt||"";
      var changed=!er||Object.keys(patch).some(function(key){return String(old[key]===undefined?"":old[key])!==String(patch[key]);});
      if(changed){patch.UpdatedAt=now;if(er)updateRowByIndex_(getSheet_(SHEET_ENROLLMENTS),getHeaders_(getSheet_(SHEET_ENROLLMENTS)),er.rowIndex,patch);else{Object.assign(patch,{EnrollmentID:id,NIK:normalizeNik_(p.NIK),Program:p.Program,Batch:p.Batch,Region:p.Region,IsActive:false,CreatedAt:now});var ri=appendObject_(getSheet_(SHEET_ENROLLMENTS),getHeaders_(getSheet_(SHEET_ENROLLMENTS)),patch);ers[k]=[{rowIndex:ri,data:patch}];}written++;}else skipped++;
      outcomes.push({nik:normalizeNik_(p.NIK),program:p.Program,batch:p.Batch,status:changed?"UPDATED":"UNCHANGED"});
    }catch(error){outcomes.push({nik:normalizeNik_(pRow.data.NIK),status:"FAILED",message:error.message});}});
    var result={next:Math.min(start+items.length,plan.rows.length),total:plan.rows.length,written:written,skipped:skipped,failed:outcomes.filter(function(x){return x.status==="FAILED";}).length,outcomes:outcomes,done:start+items.length>=plan.rows.length};
    if(written)auditLog_(s.nik,s.role,"SYNC_DASHBOARD_FROM_SHEETS","Enrollment","",JSON.stringify({start:start,written:written,source:"LOCAL_SHEETS",participantWrites:0}));
    receipt={key:cacheKey,value:JSON.stringify({signature:signature,result:result})};return result;
  });
  if(receipt)try{CacheService.getScriptCache().put(receipt.key,receipt.value,1200);}catch(ignore){}
  return result;
}, arguments);}

// Enrollment-scoped records carry Segment; accounts and global settings remain shared.
Object.keys(SHEET_HEADERS).forEach(function(name){
  if([SHEET_USERS,SHEET_USER_IMPORT,SHEET_SETTINGS].indexOf(name)<0 && SHEET_HEADERS[name].indexOf("Segment")<0)SHEET_HEADERS[name].push("Segment");
});
