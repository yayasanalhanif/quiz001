/**
 * UNIVERSAL QUIZ ENGINE V3.2 — BACKEND FINAL
 * Pasangan untuk Admin V3.2.
 *
 * Fitur:
 * - Admin login
 * - Quiz CRUD
 * - Questions
 * - Learning Materials
 * - Participants
 * - Penilaian essay
 * - Ranking & winners
 * - WhatsApp winner contact
 * - HadiahStatus / HadiahContactedAt / HadiahDiberikanAt / HadiahCatatan
 * - Publication brief
 *
 * CATATAN:
 * Backend ini sengaja memakai deteksi header sehingga tetap kompatibel
 * dengan beberapa nama sheet/header lama.
 */

// ============================================================
// KONFIGURASI
// ============================================================

const UQE_CONFIG = {
  SHEET_ID: '1CPPivGfjLp-Xg0qEKIP4CJ2A6wonQ18FANmtyczDtW8',

  ADMIN_USER: 'admin',
  ADMIN_PASS: '123456',

  // Nama file HTML di Apps Script
  ADMIN_HTML: 'Admin',
  PARTICIPANT_HTML: 'Index'
};

// Nama sheet yang didukung.
// Sistem akan memakai nama pertama yang ditemukan.
const SHEET_ALIASES = {
  QUIZZES: ['Quizzes', 'Quiz', 'Kuis'],
  QUESTIONS: ['Questions', 'Soal'],
  MATERIALS: ['Materials', 'Materi'],
  PARTICIPANTS: ['Participants', 'Peserta'],
  SCORES: ['Scores', 'Penilaian'],
  SETTINGS: ['Settings', 'Pengaturan']
};

// ============================================================
// WEB APP
// ============================================================

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};

  if (p.page === 'admin') {
    return HtmlService
      .createHtmlOutputFromFile(UQE_CONFIG.ADMIN_HTML)
      .setTitle('Universal Quiz Engine — Admin V3.2')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService
    .createHtmlOutputFromFile(UQE_CONFIG.PARTICIPANT_HTML)
    .setTitle('Sayembara Mutiara Ad-Dhuha')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// CORE HELPERS
// ============================================================

function getSS_() {
  return SpreadsheetApp.openById(UQE_CONFIG.SHEET_ID);
}

function normalize_(v) {
  return String(v == null ? '' : v).trim();
}

function auth_(username, password) {
  return normalize_(username) === UQE_CONFIG.ADMIN_USER &&
         normalize_(password) === UQE_CONFIG.ADMIN_PASS;
}

function requireAdmin_(username, password) {
  const u = normalize_(username);
  const p = String(password || '');
  
  // Cek Admin Utama
  if (u === UQE_CONFIG.ADMIN_USER && p === UQE_CONFIG.ADMIN_PASS) return true;
  
  // Cek Lulu
  if (u === 'Lulu' && p === 'abumaryam') return true;
  
  throw new Error('Username atau password admin salah.');
}

function findSheet_(names) {
  const ss = getSS_();
  for (const name of names) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
}

function sheet_(key, createIfMissing) {
  const sh = findSheet_(SHEET_ALIASES[key]);
  if (sh) return sh;
  if (!createIfMissing) return null;

  const ss = getSS_();
  return ss.insertSheet(SHEET_ALIASES[key][0]);
}

function headers_(sh) {
  if (!sh || sh.getLastColumn() < 1) return [];
  return sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(normalize_);
}

function headerMap_(sh) {
  const h = headers_(sh);
  const m = {};
  h.forEach((x, i) => {
    if (x) m[x] = i + 1;
  });
  return m;
}

function ensureColumns_(sh, cols) {
  if (!sh) throw new Error('Sheet tidak ditemukan.');

  let h = headers_(sh);

  if (!h.length || h.every(x => !x)) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    return headerMap_(sh);
  }

  cols.forEach(col => {
    if (!h.includes(col)) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(col);
      h.push(col);
    }
  });

  return headerMap_(sh);
}

function ensureDatabase_() {
  const ss = getSS_();

  const defs = {
    QUIZZES: [
      'QuizID','Slug','Judul','Tagline','Deskripsi','DurasiMenit',
      'TanggalMulai','TanggalSelesai','Hadiah','JumlahPemenang','Status',
      'AcakSoal','AcakPilihan','Penyelenggara','SumberMateri','Pemateri',
      'LinkMateri','Peserta','Biaya','Kontak','Instagram','Facebook',
      'Website','Ketentuan','CTA','LogoURL','QRURL','PosterNotes',
      'PesanPemenang','CreatedAt','UpdatedAt'
    ],
    QUESTIONS: [
      'QuestionID','QuizID','NoUrut','Pertanyaan','Tipe','Bobot',
      'OptionsJSON','AnswerKey','Aktif','CreatedAt','UpdatedAt','Status'
    ],
    MATERIALS: [
      'MaterialID','QuizID','NoUrut','Tipe','Judul','Deskripsi',
      'URL','Konten','Aktif','CreatedAt','UpdatedAt'
    ],
    PARTICIPANTS: [
      'ID','QuizID','Timestamp','Nama','WhatsApp','Email',
      'AnswersJSON','Total','Status','FinishTime','Duration',
      'WinnerRank','HadiahStatus','HadiahContactedAt',
      'HadiahDiberikanAt','HadiahCatatan','Catatan','LastUpdated'
    ],
    SCORES: [
      'ID','QuizID','ScoresJSON','Total','NotesJSON','Status','UpdatedAt'
    ]
  };

  Object.keys(defs).forEach(k => {
    const sh = sheet_(k, true);
    ensureColumns_(sh, defs[k]);
  });

  return true;
}

// ============================================================
// LOGIN
// ============================================================

function adminLogin(username, password) {
  const u = normalize_(username);
  const p = String(password || '');

  // Cek Admin Utama
  if (u === UQE_CONFIG.ADMIN_USER && p === UQE_CONFIG.ADMIN_PASS) {
    return { ok: true, role: 'admin' };
  }

  // Cek Lulu (Viewer)
  if (u === 'Lulu' && p === 'abumaryam') {
    return { ok: true, role: 'viewer' };
  }

  return { ok: false };
}

// ============================================================
// QUIZ
// ============================================================

function listQuizzes(username, password) {
  requireAdmin_(username, password);

  try {
    ensureDatabase_();
    const sh = sheet_('QUIZZES');
    const hm = headerMap_(sh);
    const rows = dataRows_(sh);

    return {
      ok: true,
      quizzes: rows.map(r => quizFromRow_(r, hm))
    };
  } catch (e) {
    return { ok: false, message: e.message || String(e) };
  }
}

function quizFromRow_(r, hm) {
  return {
    QuizID: val_(r, hm, 'QuizID'),
    Slug: val_(r, hm, 'Slug'),
    Judul: val_(r, hm, 'Judul'),
    Tagline: val_(r, hm, 'Tagline'),
    Deskripsi: val_(r, hm, 'Deskripsi'),
    DurasiMenit: Number(val_(r, hm, 'DurasiMenit') || 30),
    TanggalMulai: dateOut_(val_(r, hm, 'TanggalMulai')),
    TanggalSelesai: dateOut_(val_(r, hm, 'TanggalSelesai')),
    Hadiah: Number(val_(r, hm, 'Hadiah') || 0),
    JumlahPemenang: Number(val_(r, hm, 'JumlahPemenang') || 0),
    Status: val_(r, hm, 'Status') || 'DRAFT',
    AcakSoal: bool_(val_(r, hm, 'AcakSoal')),
    AcakPilihan: bool_(val_(r, hm, 'AcakPilihan')),
    Penyelenggara: val_(r, hm, 'Penyelenggara'),
    SumberMateri: val_(r, hm, 'SumberMateri'),
    Pemateri: val_(r, hm, 'Pemateri'),
    LinkMateri: val_(r, hm, 'LinkMateri'),
    Peserta: val_(r, hm, 'Peserta'),
    Biaya: val_(r, hm, 'Biaya'),
    Kontak: val_(r, hm, 'Kontak'),
    Instagram: val_(r, hm, 'Instagram'),
    Facebook: val_(r, hm, 'Facebook'),
    Website: val_(r, hm, 'Website'),
    Ketentuan: val_(r, hm, 'Ketentuan'),
    CTA: val_(r, hm, 'CTA'),
    LogoURL: val_(r, hm, 'LogoURL'),
    QRURL: val_(r, hm, 'QRURL'),
    PosterNotes: val_(r, hm, 'PosterNotes'),
    PesanPemenang: val_(r, hm, 'PesanPemenang')
  };
}

function updateQuiz(username, password, d) {
  if (normalize_(username) === 'Lulu') throw new Error('Akses ditolak.'); 
  requireAdmin_(username, password);

  try {
    ensureDatabase_();
    if (!d || !normalize_(d.QuizID)) {
      return { ok: false, message: 'QuizID wajib diisi.' };
    }

    const sh = sheet_('QUIZZES');
    const hm = headerMap_(sh);
    const rowIndex = findRowBy_(sh, hm, 'QuizID', normalize_(d.QuizID));

    const fields = [
      'QuizID','Slug','Judul','Tagline','Deskripsi','DurasiMenit',
      'TanggalMulai','TanggalSelesai','Hadiah','JumlahPemenang','Status',
      'AcakSoal','AcakPilihan','Penyelenggara','SumberMateri','Pemateri',
      'LinkMateri','Peserta','Biaya','Kontak','Instagram','Facebook',
      'Website','Ketentuan','CTA','LogoURL','QRURL','PosterNotes',
      'PesanPemenang'
    ];

    let row;
    if (rowIndex < 0) {
      row = new Array(sh.getLastColumn()).fill('');
      fields.forEach(f => row[hm[f] - 1] = d[f] == null ? '' : d[f]);
      if (hm.CreatedAt) row[hm.CreatedAt - 1] = new Date();
      if (hm.UpdatedAt) row[hm.UpdatedAt - 1] = new Date();
      sh.appendRow(row);
    } else {
      row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
      fields.forEach(f => {
        if (hm[f]) row[hm[f] - 1] = d[f] == null ? '' : d[f];
      });
      if (hm.UpdatedAt) row[hm.UpdatedAt - 1] = new Date();
      sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    }

    return { ok: true, quiz: d };
  } catch (e) {
    return { ok: false, message: e.message || String(e) };
  }
}

function duplicateQuiz(username, password, oldId, newId, newTitle) {
    if (normalize_(username) === 'Lulu') throw new Error('Akses ditolak.'); 
  requireAdmin_(username, password);

  try {
    ensureDatabase_();

    if (findRowBy_(sheet_('QUIZZES'), headerMap_(sheet_('QUIZZES')), 'QuizID', newId) > 0) {
      return { ok: false, message: 'QuizID baru sudah digunakan.' };
    }

    const old = getQuizObject_(oldId);
    if (!old) return { ok: false, message: 'Kuis sumber tidak ditemukan.' };

    old.QuizID = newId;
    old.Slug = String(newId).toLowerCase();
    old.Judul = newTitle || (old.Judul + ' — Salinan');
    old.Status = 'DRAFT';

    const saved = updateQuiz(username, password, old);
    if (!saved.ok) return saved;

    // Duplikasi soal
    const qs = getQuestions_(oldId);
    saveQuestions(username, password, newId, qs.map(q => {
      const x = Object.assign({}, q);
      x.QuestionID = '';
      return x;
    }));

    // Duplikasi materi
    const ms = getMaterials_(oldId);
    saveMaterials(username, password, newId, ms.map(m => {
      const x = Object.assign({}, m);
      x.MaterialID = '';
      return x;
    }));

    return { ok: true, quiz: old };
  } catch (e) {
    return { ok: false, message: e.message || String(e) };
  }
}

// ============================================================
// PUBLIC QUIZ
// ============================================================

function getPublicQuiz(quizId) {
  try {
    ensureDatabase_();
    const qid = String(quizId || 'Q001').trim().toLowerCase();
    
    const sh = sheet_('QUIZZES');
    if (!sh) return { ok: false, message: 'Sheet QUIZZES tidak ditemukan.' };

    const hm = headerMap_(sh);
    const rows = dataRows_(sh);
    
    const row = rows.find(r => String(val_(r, hm, 'QuizID') || '').trim().toLowerCase() === qid);
    if (!row) return { ok: false, message: 'Kuis tidak ditemukan.' };

    const qObj = quizFromRow_(row, hm);
    const questions = getQuestions_(qObj.QuizID, false);
    const materials = getMaterials_(qObj.QuizID);

    const quiz = {
      id: qObj.QuizID,
      title: qObj.Judul,
      tagline: qObj.Tagline,
      description: qObj.Deskripsi,
      durationMinutes: qObj.DurasiMenit,
      prize: qObj.Hadiah,
      winners: qObj.JumlahPemenang,
      status: qObj.Status,
      hasLearningMaterial: materials.length > 0 || Boolean(qObj.LinkMateri),
      materialLink: qObj.LinkMateri,
      materialLinkEmbedUrl: youtubeEmbedUrl_(qObj.LinkMateri, 'youtube'),
      source: qObj.SumberMateri,
      speaker: qObj.Pemateri,
      audience: qObj.Peserta,
      fee: qObj.Biaya,
      terms: qObj.Ketentuan,
      cta: qObj.CTA,
      start: qObj.TanggalMulai,
      end: qObj.TanggalSelesai,
      questions: questions.map(q => ({
        QuestionID: q.QuestionID,
        Pertanyaan: q.Pertanyaan,
        Tipe: q.Tipe,
        Bobot: q.Bobot,
        Options: q.Options
      })),
      materials: materials
    };

    return { ok: true, quiz: quiz };
  } catch (e) {
    return { ok: false, message: 'Error: ' + e.message };
  }
}

function getQuizObject_(quizId) {
  const sh = sheet_('QUIZZES');
  if (!sh) return null;
  const hm = headerMap_(sh);
  const row = findRowBy_(sh, hm, 'QuizID', normalize_(quizId));
  if (row < 2) return null;
  return quizFromRow_(sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0], hm);
}

// ============================================================
// QUESTIONS
// ============================================================

function getQuestions_(quizId, includeAll) {
  const sh = sheet_('QUESTIONS');
  if (!sh) return [];

  const hm = headerMap_(sh);
  const qid = normalize_(quizId);

  return dataRows_(sh)
    .filter(r => normalize_(val_(r, hm, 'QuizID')) === qid)
    .filter(r => boolDefault_(val_(r, hm, 'Aktif'), true))
    .filter(r => {
      if (includeAll === true) return true;
      const status = normalize_(val_(r, hm, 'Status')).toUpperCase();
      return !status || status === 'OPEN';
    })
    .sort((a,b) => Number(val_(a,hm,'NoUrut')||0) - Number(val_(b,hm,'NoUrut')||0))
    .map(r => ({
      QuestionID: val_(r,hm,'QuestionID'),
      QuizID: val_(r,hm,'QuizID'),
      NoUrut: Number(val_(r,hm,'NoUrut')||0),
      Pertanyaan: val_(r,hm,'Pertanyaan'),
      Tipe: val_(r,hm,'Tipe') || 'essay',
      Bobot: Number(val_(r,hm,'Bobot')||0),
      Options: parseJson_(val_(r,hm,'OptionsJSON'), []),
      AnswerKey: val_(r,hm,'AnswerKey'),
      Aktif: boolDefault_(val_(r,hm,'Aktif'), true),
      Status: normalize_(val_(r,hm,'Status')).toUpperCase() || 'OPEN'
    }));
}

function getAdminQuestions(username, password, quizId) {
  requireAdmin_(username, password);
  try {
    ensureDatabase_();
    return {ok:true, questions:getQuestions_(quizId, true)};
  } catch(e) {
    return {ok:false, message:e.message || String(e)};
  }
}

function saveQuestions(username, password, quizId, questions) {
  requireAdmin_(username, password);

  try {
    ensureDatabase_();
    const sh = sheet_('QUESTIONS');
    const hm = headerMap_(sh);

    // Hapus record soal lama untuk quiz ini, lalu tulis ulang.
    const last = sh.getLastRow();
    if (last >= 2) {
      const rows = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
      const keep = rows.filter(r => val_(r,hm,'QuizID') !== normalize_(quizId));
      sh.getRange(2,1,last-1,sh.getLastColumn()).clearContent();
      if (keep.length) sh.getRange(2,1,keep.length,keep[0].length).setValues(keep);
    }

    (questions || []).forEach((q, i) => {
      const row = new Array(sh.getLastColumn()).fill('');
      row[hm.QuestionID-1] = normalize_(q.QuestionID) || ('QST-' + Utilities.getUuid().slice(0,8));
      row[hm.QuizID-1] = quizId;
      row[hm.NoUrut-1] = i + 1;
      row[hm.Pertanyaan-1] = q.Pertanyaan || '';
      row[hm.Tipe-1] = q.Tipe || 'essay';
      row[hm.Bobot-1] = Number(q.Bobot || 0);
      row[hm.OptionsJSON-1] = JSON.stringify(q.Options || []);
      if (hm.AnswerKey) row[hm.AnswerKey-1] = q.AnswerKey || '';
      row[hm.Aktif-1] = q.Aktif !== false;

      let status = normalize_(q.Status).toUpperCase();
      if (!['DRAFT','OPEN','CLOSED'].includes(status)) status = 'DRAFT';
      if (hm.Status) row[hm.Status-1] = status;

      if (hm.CreatedAt) row[hm.CreatedAt-1] = new Date();
      if (hm.UpdatedAt) row[hm.UpdatedAt-1] = new Date();
      sh.appendRow(row);
    });

    return { ok:true, count:(questions || []).length };
  } catch(e) {
    return { ok:false, message:e.message || String(e) };
  }
}

// ============================================================
// MATERIALS
// ============================================================

function getMaterials_(quizId) {
  const sh = sheet_('MATERIALS');
  if (!sh) return [];

  const hm = headerMap_(sh);
  return dataRows_(sh)
    .filter(r => val_(r,hm,'QuizID') === normalize_(quizId))
    .filter(r => boolDefault_(val_(r,hm,'Aktif'), true))
    .sort((a,b) => Number(val_(a,hm,'NoUrut')||0) - Number(val_(b,hm,'NoUrut')||0))
    .map(r => ({
      MaterialID: val_(r,hm,'MaterialID'),
      QuizID: val_(r,hm,'QuizID'),
      NoUrut: Number(val_(r,hm,'NoUrut')||0),
      Tipe: val_(r,hm,'Tipe') || 'link',
      Judul: val_(r,hm,'Judul'),
      Deskripsi: val_(r,hm,'Deskripsi'),
      URL: val_(r,hm,'URL'),
      embedUrl: youtubeEmbedUrl_(val_(r,hm,'URL'), val_(r,hm,'Tipe')),
      Konten: val_(r,hm,'Konten'),
      Aktif: boolDefault_(val_(r,hm,'Aktif'), true)
    }));
}

function getMaterials(username, password, quizId) {
  requireAdmin_(username, password);
  try {
    ensureDatabase_();
    return {ok:true, materials:getMaterials_(quizId)};
  } catch(e) {
    return {ok:false,message:e.message || String(e)};
  }
}

function saveMaterials(username, password, quizId, materials) {
  requireAdmin_(username, password);

  try {
    ensureDatabase_();
    const sh = sheet_('MATERIALS');
    const hm = headerMap_(sh);
    const last = sh.getLastRow();

    if (last >= 2) {
      const rows = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
      const keep = rows.filter(r => val_(r,hm,'QuizID') !== normalize_(quizId));
      sh.getRange(2,1,last-1,sh.getLastColumn()).clearContent();
      if (keep.length) sh.getRange(2,1,keep.length,keep[0].length).setValues(keep);
    }

    (materials || []).forEach((m,i) => {
      const row = new Array(sh.getLastColumn()).fill('');
      row[hm.MaterialID-1] = normalize_(m.MaterialID) || ('MAT-' + Utilities.getUuid().slice(0,8));
      row[hm.QuizID-1] = quizId;
      row[hm.NoUrut-1] = i + 1;
      row[hm.Tipe-1] = m.Tipe || 'link';
      row[hm.Judul-1] = m.Judul || '';
      row[hm.Deskripsi-1] = m.Deskripsi || '';
      row[hm.URL-1] = m.URL || '';
      row[hm.Konten-1] = m.Konten || '';
      row[hm.Aktif-1] = m.Aktif !== false;
      if (hm.CreatedAt) row[hm.CreatedAt-1] = new Date();
      if (hm.UpdatedAt) row[hm.UpdatedAt-1] = new Date();
      sh.appendRow(row);
    });

    return {ok:true,count:(materials || []).length};
  } catch(e) {
    return {ok:false,message:e.message || String(e)};
  }
}

// ============================================================
// PARTICIPANTS / DASHBOARD
// ============================================================

function getAdminDashboard(username, password, quizId) {
  try {
    // Dipindahkan ke dalam try-catch agar jika sesi terputus, error-nya jelas tertangkap
    requireAdmin_(username, password);
    ensureDatabase_();

    const qid = normalize_(quizId);
    if (!qid) {
      return { ok: false, message: 'QuizID kosong.' };
    }

    const quiz = getQuizObject_(qid);
    if (!quiz) {
      return { ok: false, message: 'Kuis tidak ditemukan: ' + qid };
    }

    let data = [];
    try {
      const sh = sheet_('PARTICIPANTS');
      if (sh) {
        const hm = headerMap_(sh);
        data = dataRows_(sh)
          .filter(r => normalize_(val_(r, hm, 'QuizID')) === qid)
          .map(r => {
            const id = val_(r, hm, 'ID') || val_(r, hm, 'ParticipantID');
            
            // Pengamanan ketat agar answers selalu terbaca sebagai Array
            let ans = parseJson_(val_(r, hm, 'AnswersJSON'), []);
            if (!Array.isArray(ans)) ans = [];

            return {
              id: id,
              timestamp: dateOut_(val_(r, hm, 'Timestamp')),
              nama: val_(r, hm, 'Nama'),
              wa: val_(r, hm, 'WhatsApp'),
              email: val_(r, hm, 'Email'),
              answers: ans,
              total: Number(val_(r, hm, 'Total') || 0),
              status: val_(r, hm, 'Status') || 'Masuk',
              finishTime: dateOut_(val_(r, hm, 'FinishTime') || val_(r, hm, 'WaktuSelesai')),
              duration: val_(r, hm, 'Duration') || val_(r, hm, 'Durasi') || '-',
              winnerRank: Number(val_(r, hm, 'WinnerRank') || 0),
              HadiahStatus: val_(r, hm, 'HadiahStatus') || 'Belum Diproses',
              HadiahContactedAt: dateOut_(val_(r, hm, 'HadiahContactedAt')),
              HadiahDiberikanAt: dateOut_(val_(r, hm, 'HadiahDiberikanAt')),
              HadiahCatatan: val_(r, hm, 'HadiahCatatan'),
              note: val_(r, hm, 'Catatan')
            };
          });
      }
    } catch (participantError) {
      data = [];
    }

    const norm = x => String(x || '').trim().toUpperCase();
    const selesai = data.filter(x => ['SELESAI', 'FINISHED'].includes(norm(x.status))).length;
    const mengerjakan = data.filter(x => ['MENGERJAKAN', 'STARTED'].includes(norm(x.status))).length;
    const habis = data.filter(x => ['WAKTU HABIS', 'TIMEOUT'].includes(norm(x.status))).length;
    const belumDinilai = data.filter(x => !Number(x.total) && ['BELUM DINILAI', 'SELESAI', 'WAKTU HABIS'].includes(norm(x.status))).length;

    let questions = [];
    try {
      questions = getQuestions_(qid, true);
    } catch (questionError) {
      questions = [];
    }

    const result = {
      ok: true,
      quiz: { id: quiz.QuizID, title: quiz.Judul, prize: Number(quiz.Hadiah || 0), winners: Number(quiz.JumlahPemenang || 0) },
      stats: { total: data.length, selesai: selesai, mengerjakan: mengerjakan, habis: habis, belumDinilai: belumDinilai },
      data: data,
      questions: questions
    };

    // Serialisasi ulang untuk membuang objek yang tidak kompatibel (mis. Date) yang bisa membuat gagal muat
    return JSON.parse(JSON.stringify(result));

  } catch (e) {
    return { ok: false, message: 'getAdminDashboard Error: ' + (e.message || String(e)) };
  }
}

function participantFromRow_(r,hm,scores) {
  const id = val_(r,hm,'ID');
  const score = scores[id] || {};
  
  // Ambil dan pastikan answers selalu berformat Array
  let answers = parseJson_(val_(r,hm,'AnswersJSON'), []);
  if (!Array.isArray(answers)) {
    answers = [];
  }

  return {
    id:id,
    timestamp:dateOut_(val_(r,hm,'Timestamp')),
    nama:val_(r,hm,'Nama'),
    wa:val_(r,hm,'WhatsApp'),
    email:val_(r,hm,'Email'),
    answers:answers.map(a => typeof a === 'object' ? a : {answer:String(a||'')}),
    scores:score.scores || [],
    total:Number(score.total != null ? score.total : val_(r,hm,'Total') || 0),
    status:score.status || val_(r,hm,'Status') || 'Masuk',
    finishTime:dateOut_(score.finishTime || val_(r,hm,'FinishTime')),
    duration:score.duration || val_(r,hm,'Duration') || '-',
    winnerRank:Number(val_(r,hm,'WinnerRank') || 0),
    HadiahStatus:val_(r,hm,'HadiahStatus') || 'Belum Diproses',
    HadiahContactedAt:dateOut_(val_(r,hm,'HadiahContactedAt')),
    HadiahDiberikanAt:dateOut_(val_(r,hm,'HadiahDiberikanAt')),
    HadiahCatatan:val_(r,hm,'HadiahCatatan'),
    note:val_(r,hm,'Catatan')
  };
}

// ============================================================
// SCORES
// ============================================================

function getScoreMap_(quizId) {
  const sh = sheet_('SCORES');
  if (!sh) return {};
  const hm = headerMap_(sh);
  if (!hm.ID || !hm.QuizID) return {};
  const map = {};

  dataRows_(sh)
    .filter(r => normalize_(val_(r,hm,'QuizID')) === normalize_(quizId))
    .forEach(r => {
      const id = normalize_(val_(r,hm,'ID'));
      if (!id) return;
      map[id] = {
        scores:parseJson_(val_(r,hm,'ScoresJSON'), []),
        total:Number(val_(r,hm,'Total') || 0),
        notes:parseJson_(val_(r,hm,'NotesJSON'), {}),
        status:val_(r,hm,'Status'),
        finishTime:val_(r,hm,'FinishTime'),
        duration:val_(r,hm,'Duration')
      };
    });
  return map;
}

function saveScores(username,password,participantId,scores,notes) {
  requireAdmin_(username,password);

  try {
    ensureDatabase_();

    const p = findParticipant_(participantId);
    if (!p) return {ok:false,message:'Peserta tidak ditemukan.'};

    const q = getQuestions_(p.quizId);
    const arr = q.map(x => Number(scores && scores[x.QuestionID] || 0));
    const total = arr.reduce((a,b)=>a+b,0);

    const sh = sheet_('SCORES');
    const hm = headerMap_(sh);
    const rowIndex = findRowBy_(sh,hm,'ID',participantId);

    const row = new Array(sh.getLastColumn()).fill('');
    row[hm.ID-1] = participantId;
    row[hm.QuizID-1] = p.quizId;
    row[hm.ScoresJSON-1] = JSON.stringify(arr);
    row[hm.Total-1] = total;
    row[hm.NotesJSON-1] = JSON.stringify(notes || {});
    row[hm.Status-1] = 'Dinilai';
    if (hm.UpdatedAt) row[hm.UpdatedAt-1] = new Date();

    if (rowIndex < 0) sh.appendRow(row);
    else sh.getRange(rowIndex,1,1,row.length).setValues([row]);

    // Sinkron ke Participants
    const ps = sheet_('PARTICIPANTS');
    const ph = headerMap_(ps);
    const pr = findRowBy_(ps,ph,'ID',participantId);
    if (pr >= 2) {
      if (ph.Total) ps.getRange(pr,ph.Total).setValue(total);
      if (ph.Status) ps.getRange(pr,ph.Status).setValue('Dinilai');
      if (ph.Catatan) ps.getRange(pr,ph.Catatan).setValue(JSON.stringify(notes || {}));
      if (ph.LastUpdated) ps.getRange(pr,ph.LastUpdated).setValue(new Date());
    }

    return {ok:true,total:total};
  } catch(e) {
    return {ok:false,message:e.message || String(e)};
  }
}

// Alias kompatibilitas
function saveGrade(username,password,participantId,scores,notes) {
  return saveScores(username,password,participantId,scores,notes);
}

// ============================================================
// RANKING & WINNERS
// ============================================================

function publishWinners(username,password,quizId) {
  requireAdmin_(username,password);

  try {
    const dashboard = getAdminDashboard(username,password,quizId);
    if (!dashboard.ok) return dashboard;

    const winnersCount = Number(dashboard.quiz.winners || 0);
    if (winnersCount <= 0) return {ok:false,message:'Jumlah pemenang belum diatur.'};

    const eligible = dashboard.data
      .filter(p => {
        const st = String(p.status || '').toUpperCase();
        return ['SELESAI','FINISHED','DINILAI'].includes(st) || Number(p.total) > 0;
      })
      .filter(p => p.status !== 'Batal')
      .sort((a,b) => {
        const n = Number(b.total||0) - Number(a.total||0);
        if (n) return n;
        const da = a.finishTime ? new Date(a.finishTime).getTime() : Infinity;
        const db = b.finishTime ? new Date(b.finishTime).getTime() : Infinity;
        return da - db;
      });

    const winners = eligible.slice(0,winnersCount);
    const ps = sheet_('PARTICIPANTS');
    const hm = headerMap_(ps);

    winners.forEach((w,i) => {
      const row = findRowBy_(ps,hm,'ID',w.id);
      if (row >= 2) {
        if (hm.WinnerRank) ps.getRange(row,hm.WinnerRank).setValue(i+1);
        if (hm.Status) ps.getRange(row,hm.Status).setValue('Pemenang');
        if (hm.HadiahStatus) {
          const current = ps.getRange(row,hm.HadiahStatus).getValue();
          if (!current) ps.getRange(row,hm.HadiahStatus).setValue('Belum Diproses');
        }
        if (hm.LastUpdated) ps.getRange(row,hm.LastUpdated).setValue(new Date());
      }
    });

    return {
      ok:true,
      winners:winners.map((w,i)=>({
        rank:i+1,id:w.id,nama:w.nama,wa:w.wa,total:w.total,status:'Pemenang',
        HadiahStatus:w.HadiahStatus || 'Belum Diproses'
      }))
    };
  } catch(e) {
    return {ok:false,message:e.message || String(e)};
  }
}

// ============================================================
// HADIAH
// ============================================================

function updateGiftStatus(username,password,participantId,status,note) {
  requireAdmin_(username,password);

  try {
    const allowed = ['Belum Diproses','Sudah Dihubungi','Hadiah Diberikan'];
    status = normalize_(status);
    if (!allowed.includes(status)) {
      return {ok:false,message:'Status hadiah tidak valid.'};
    }

    const sh = sheet_('PARTICIPANTS');
    const hm = headerMap_(sh);
    const row = findRowBy_(sh,hm,'ID',participantId);
    if (row < 2) return {ok:false,message:'Peserta tidak ditemukan.'};

    if (hm.HadiahStatus) sh.getRange(row,hm.HadiahStatus).setValue(status);
    if (hm.HadiahCatatan && note !== undefined) {
      sh.getRange(row,hm.HadiahCatatan).setValue(note || '');
    }

    if (status === 'Sudah Dihubungi' && hm.HadiahContactedAt) {
      sh.getRange(row,hm.HadiahContactedAt).setValue(new Date());
    }

    if (status === 'Hadiah Diberikan' && hm.HadiahDiberikanAt) {
      sh.getRange(row,hm.HadiahDiberikanAt).setValue(new Date());
    }

    if (hm.LastUpdated) sh.getRange(row,hm.LastUpdated).setValue(new Date());

    return {
      ok:true,
      participantId:participantId,
      status:status
    };
  } catch(e) {
    return {ok:false,message:e.message || String(e)};
  }
}

// ============================================================
// PUBLICATION
// ============================================================

function getPublicationBrief(username,password,quizId) {
  requireAdmin_(username,password);

  try {
    const q = getQuizObject_(quizId);
    if (!q) return {ok:false,message:'Kuis tidak ditemukan.'};

    const qs = getQuestions_(quizId);

    const brief = [
      'DATA PUBLIKASI QUIZ',
      '',
      'NAMA KEGIATAN: ' + q.Judul,
      'TEMA/TAGLINE: ' + (q.Tagline || '-'),
      'PENYELENGGARA: ' + (q.Penyelenggara || '-'),
      'DESKRIPSI: ' + (q.Deskripsi || '-'),
      'SUMBER MATERI: ' + (q.SumberMateri || '-'),
      'PEMATERI: ' + (q.Pemateri || '-'),
      'LINK MATERI: ' + (q.LinkMateri || '-'),
      'PESERTA: ' + (q.Peserta || '-'),
      'BIAYA: ' + (q.Biaya || '-'),
      'DURASI: ' + q.DurasiMenit + ' menit',
      'JUMLAH PEMENANG: ' + q.JumlahPemenang,
      'HADIAH / PEMENANG: Rp' + Number(q.Hadiah||0).toLocaleString('id-ID'),
      'TOTAL HADIAH: Rp' + (Number(q.Hadiah||0)*Number(q.JumlahPemenang||0)).toLocaleString('id-ID'),
      'KONTAK: ' + (q.Kontak || '-'),
      'INSTAGRAM: ' + (q.Instagram || '-'),
      'FACEBOOK: ' + (q.Facebook || '-'),
      'WEBSITE: ' + (q.Website || '-'),
      '',
      'KETENTUAN:',
      q.Ketentuan || '-',
      '',
      'CTA: ' + (q.CTA || '-'),
      'CATATAN POSTER: ' + (q.PosterNotes || '-'),
      '',
      'JUMLAH SOAL: ' + qs.length,
      'QuizID: ' + q.QuizID
    ].join('\n');

    return {
      ok:true,
      quiz:{
        id:q.QuizID,
        title:q.Judul,
        tagline:q.Tagline,
        durationMinutes:q.DurasiMenit,
        winners:q.JumlahPemenang,
        prize:q.Hadiah
      },
      questionCount:qs.length,
      brief:brief
    };
  } catch(e) {
    return {ok:false,message:e.message || String(e)};
  }
}

// ============================================================
// PARTICIPANT API — kompatibilitas dasar
// ============================================================

function submitQuizParticipant(data) {
  try {
    ensureDatabase_();

    const quizId = normalize_(data.quizId || data.QuizID);
    const quiz = getQuizObject_(quizId);
    if (!quiz) return {ok:false,message:'Kuis tidak ditemukan.'};

    const id = normalize_(data.id) || ('P-' + Utilities.getUuid().slice(0,8).toUpperCase());
    const sh = sheet_('PARTICIPANTS');
    const hm = headerMap_(sh);

    const row = new Array(sh.getLastColumn()).fill('');
    row[hm.ID-1] = id;
    row[hm.QuizID-1] = quizId;
    row[hm.Timestamp-1] = new Date();
    row[hm.Nama-1] = data.nama || data.Nama || '';
    row[hm.WhatsApp-1] = data.wa || data.WhatsApp || '';
    row[hm.Email-1] = data.email || data.Email || '';
    row[hm.AnswersJSON-1] = JSON.stringify(data.answers || []);
    row[hm.Total-1] = 0;
    row[hm.Status-1] = data.status || 'Masuk';
    if (hm.LastUpdated) row[hm.LastUpdated-1] = new Date();
    if (hm.HadiahStatus) row[hm.HadiahStatus-1] = 'Belum Diproses';

    sh.appendRow(row);
    return {ok:true,id:id};
  } catch(e) {
    return {ok:false,message:e.message || String(e)};
  }
}

function getQuizState(id,wa,quizId) {
  try {
    const p = findParticipant_(id);
    if (!p) return {ok:false};
    if (normalize_(quizId) && normalize_(p.quizId) !== normalize_(quizId)) return {ok:false};
    if (normalize_(wa) && normalizeWa_(p.wa) !== normalizeWa_(wa)) return {ok:false};

    const sh = sheet_('PARTICIPANTS');
    const hm = headerMap_(sh);
    const rowIndex = findRowBy_(sh,hm,'ID',id);
    let row = rowIndex >= 2 ? sh.getRange(rowIndex,1,1,sh.getLastColumn()).getValues()[0] : [];
    const quiz = getQuizObject_(p.quizId);
    const duration = Math.max(1, Number(
      val_(row,hm,'Durasi') || val_(row,hm,'Duration') || (quiz && quiz.DurasiMenit) || 30
    ));
    const started = parseDateSafe_(val_(row,hm,'WaktuMulai')) || parseDateSafe_(val_(row,hm,'Timestamp')) || parseDateSafe_(p.timestamp);
    let deadline = parseDateSafe_(val_(row,hm,'BatasWaktu'));
    if (!deadline && started) deadline = new Date(started.getTime() + duration * 60000);

    const participant = Object.assign({}, p, {
      startTime: started ? started.toISOString() : '',
      deadline: deadline ? deadline.toISOString() : '',
      duration: duration
    });

    // Kembalikan data sesi di level atas agar kompatibel dengan Index.html
    // sekaligus tetap menyediakan object participant untuk kompatibilitas lama.
    return {ok:true, participant:participant, startTime:participant.startTime, deadline:participant.deadline, durationMinutes:duration, id:participant.id, quizId:participant.quizId, wa:participant.wa};
  } catch(e) {
    return {ok:false,message:e.message || String(e)};
  }
}

// ============================================================
// PARTICIPANT QUIZ SESSION — V3.2 COMPATIBILITY
// ============================================================

/**
 * Dipanggil langsung oleh Index.html saat tombol MULAI KUIS ditekan.
 * Versi backend lama tidak memiliki fungsi ini sehingga Google Apps Script
 * mengembalikan: [fn] is not a function.
 */
function startQuiz(data) {
  try {
    data = data || {};
    ensureDatabase_();

    const quizId = normalize_(data.quizId || data.QuizID);
    if (!quizId) return {ok:false, message:'QuizID tidak ditemukan.'};

    const quiz = getQuizObject_(quizId);
    if (!quiz) return {ok:false, message:'Kuis tidak ditemukan.'};

    const status = normalize_(quiz.Status).toUpperCase();
    const isOpen = ['OPEN','ACTIVE','PUBLISHED'].includes(status);
    if (!isOpen) {
      return {ok:false, message:'Kuis saat ini ' + (quiz.Status || 'belum dibuka') + '.'};
    }

    const now = new Date();
    const startDate = parseDateSafe_(quiz.TanggalMulai);
    const endDate = parseQuizEndDate_(quiz.TanggalSelesai);
    if (startDate && now < startDate) {
      return {ok:false, message:'Kuis belum dimulai. Periode mulai: ' + startDate.toLocaleString('id-ID')};
    }
    if (endDate && now > endDate) {
      return {ok:false, message:'Periode kuis sudah berakhir.'};
    }

    if (!normalize_(data.nama)) return {ok:false, message:'Nama wajib diisi.'};
    if (!normalize_(data.wa)) return {ok:false, message:'Nomor WhatsApp wajib diisi.'};
    if (data.agree !== true) return {ok:false, message:'Persetujuan wajib dicentang.'};

    // Satu nomor WhatsApp hanya boleh mengikuti satu kali untuk satu kuis.
    const waNorm = normalizeWa_(data.wa);
    const participantSheet = sheet_('PARTICIPANTS');
    const participantHM = headerMap_(participantSheet);
    const existing = dataRows_(participantSheet).find(r =>
      normalize_(val_(r, participantHM, 'QuizID')) === quizId &&
      normalizeWa_(val_(r, participantHM, 'WhatsApp')) === waNorm
    );
    if (existing) {
      const existingId = normalize_(val_(existing, participantHM, 'ID'));
      const existingStatus = normalize_(val_(existing, participantHM, 'Status'));
      return {ok:false, message:'Nomor WhatsApp ini sudah terdaftar pada sayembara ini. ID peserta: ' + existingId + ' (' + existingStatus + ').'};
    }

    // Jika materi pembelajaran tersedia, peserta harus menekan tombol siap.
    const materials = getMaterials_(quizId);
    if (materials.length && data.ready !== true) {
      return {ok:false, message:'Silakan pelajari materi terlebih dahulu.'};
    }

    const duration = Math.max(1, Number(quiz.DurasiMenit || 30));
    const deadline = new Date(now.getTime() + duration * 60000);

    const sh = sheet_('PARTICIPANTS');
    const hm = headerMap_(sh);
    const id = 'P-' + Utilities.getUuid().slice(0,8).toUpperCase();

    const row = new Array(sh.getLastColumn()).fill('');
    row[hm.ID-1] = id;
    row[hm.QuizID-1] = quizId;
    row[hm.Timestamp-1] = now;
    if (hm.WaktuMulai) row[hm.WaktuMulai-1] = now;
    if (hm.BatasWaktu) row[hm.BatasWaktu-1] = deadline;
    if (hm.Durasi) row[hm.Durasi-1] = duration;
    row[hm.Nama-1] = normalize_(data.nama);
    row[hm.WhatsApp-1] = normalize_(data.wa);
    row[hm.Email-1] = normalize_(data.email);
    row[hm.AnswersJSON-1] = JSON.stringify([]);
    row[hm.Total-1] = 0;
    row[hm.Status-1] = 'MENGERJAKAN';
    row[hm.FinishTime-1] = '';
    row[hm.Duration-1] = '';
    row[hm.WinnerRank-1] = '';
    if (hm.HadiahStatus) row[hm.HadiahStatus-1] = 'Belum Diproses';
    if (hm.LastUpdated) row[hm.LastUpdated-1] = now;

    sh.appendRow(row);

    return {
      ok:true,
      id:id,
      quizId:quizId,
      nama:normalize_(data.nama),
      wa:normalize_(data.wa),
      deadline:deadline.toISOString(),
      startTime:now.toISOString(),
      durationMinutes:duration
    };
  } catch (e) {
    return {ok:false, message:e.message || String(e)};
  }
}

/**
 * Dipanggil Index.html saat peserta menekan KIRIM JAWABAN.
 * Mendukung MCQ/truefalse/checkbox auto-score dan essay untuk dinilai admin.
 */
function submitEntry(payload) {
  try {
    payload = payload || {};
    ensureDatabase_();

    const quizId = normalize_(payload.quizId || payload.QuizID);
    const participantId = normalize_(payload.participantId || payload.id);
    const answers = payload.answers || {};

    if (!quizId || !participantId) {
      return {ok:false, message:'Data sesi kuis tidak lengkap.'};
    }

    const p = findParticipant_(participantId);
    if (!p || p.quizId !== quizId) {
      return {ok:false, message:'Sesi peserta tidak ditemukan atau tidak sesuai kuis.'};
    }

    if (['SELESAI','DINILAI','Pemenang'.toUpperCase()].includes(String(p.status || '').toUpperCase())) {
      return {ok:false, message:'Jawaban peserta sudah dikirim sebelumnya.'};
    }

    const sh = sheet_('PARTICIPANTS');
    const hm = headerMap_(sh);
    const rowIndex = findRowBy_(sh, hm, 'ID', participantId);
    if (rowIndex < 2) return {ok:false, message:'Data peserta tidak ditemukan.'};

    const row = sh.getRange(rowIndex,1,1,sh.getLastColumn()).getValues()[0];
    const startValue = val_(row, hm, 'Timestamp');
    const start = parseDateSafe_(startValue) || new Date();
    const quiz = getQuizObject_(quizId);
    const duration = Math.max(1, Number(quiz && quiz.DurasiMenit || 30));
    const deadline = new Date(start.getTime() + duration * 60000);
    const now = new Date();
    const timedOut = payload.timedOut === true || now.getTime() > deadline.getTime();

    const qs = getQuestions_(quizId);
    const answerRows = [];
    let total = 0;
    let hasEssay = false;

    qs.forEach(q => {
      const raw = answers[q.QuestionID];
      const answer = raw == null ? (Array.isArray(raw) ? [] : '') : raw;
      let score = '';
      let note = '';

      if (q.Tipe === 'essay') {
        hasEssay = true;
      } else {
        score = autoScoreQuestion_(q, answer);
        total += Number(score || 0);
      }

      answerRows.push({
        QuestionID:q.QuestionID,
        answer:answer,
        score:score,
        note:note
      });
    });

    row[hm.AnswersJSON-1] = JSON.stringify(answerRows);
    row[hm.Total-1] = total;
    row[hm.Status-1] = timedOut ? 'WAKTU HABIS' : (hasEssay ? 'Belum Dinilai' : 'SELESAI');
    if (hm.FinishTime) row[hm.FinishTime-1] = now;
    if (hm.Duration) row[hm.Duration-1] = formatDuration_(Math.max(0, now.getTime() - start.getTime()));
    if (hm.LastUpdated) row[hm.LastUpdated-1] = now;
    sh.getRange(rowIndex,1,1,row.length).setValues([row]);

    return {
      ok:true,
      id:participantId,
      quizId:quizId,
      total:total,
      status:row[hm.Status-1],
      timedOut:timedOut,
      hasEssay:hasEssay
    };
  } catch (e) {
    return {ok:false, message:e.message || String(e)};
  }
}

function autoScoreQuestion_(q, answer) {
  const opts = Array.isArray(q.Options) ? q.Options : [];
  if (!opts.length) return 0;

  if (q.Tipe === 'checkbox') {
    const selected = Array.isArray(answer) ? answer.map(normalize_) : [];
    const correct = opts.filter(o => o && o.correct).map(o => normalize_(o.value));
    if (!correct.length) return 0;
    if (selected.length !== correct.length) return 0;
    const a = selected.slice().sort().join('|');
    const b = correct.slice().sort().join('|');
    if (a !== b) return 0;
    const scores = opts.filter(o => o && o.correct).map(o => Number(o.score || q.Bobot || 0));
    return scores.reduce((x,y)=>x+y,0) || Number(q.Bobot || 0);
  }

  const selected = normalize_(answer);
  const opt = opts.find(o => normalize_(o && o.value) === selected);
  if (!opt) return 0;
  if (opt.score !== '' && opt.score != null && !isNaN(Number(opt.score))) return Number(opt.score);
  return opt.correct ? Number(q.Bobot || 0) : 0;
}

function formatDuration_(ms) {
  const sec = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(sec / 3600);
  const m = Math.floor((sec % 3600) / 60);
  const s = sec % 60;
  return (h ? String(h).padStart(2,'0') + ':' : '') +
         String(m).padStart(2,'0') + ':' + String(s).padStart(2,'0');
}

function youtubeEmbedUrl_(url, type) {
  const u = normalize_(url);
  if (String(type || '').toLowerCase() !== 'youtube' || !u) return '';

  let id = '';
  let m = u.match(/[?&]v=([A-Za-z0-9_-]{6,})/);
  if (m) id = m[1];
  if (!id) {
    m = u.match(/youtu\.be\/([A-Za-z0-9_-]{6,})/i);
    if (m) id = m[1];
  }
  if (!id) {
    m = u.match(/youtube\.com\/(?:embed|shorts)\/([A-Za-z0-9_-]{6,})/i);
    if (m) id = m[1];
  }
  return id ? 'https://www.youtube.com/embed/' + id : '';
}

// ============================================================
// GENERIC DATA HELPERS
// ============================================================

function dataRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
}

function val_(row,hm,key) {
  return hm[key] ? row[hm[key]-1] : '';
}

function bool_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE' || String(v) === '1';
}

function boolDefault_(v,def) {
  if (v === '' || v == null) return def;
  return bool_(v);
}

function parseJson_(v,def) {
  if (v == null || v === '') return def;
  if (Array.isArray(v) || typeof v === 'object') return v;
  try { return JSON.parse(String(v)); } catch(e) { return def; }
}

function parseQuizEndDate_(v) {
  const d = parseDateSafe_(v);
  if (!d) return null;
  // Bila admin mengisi tanggal tanpa jam (mis. 2026-09-05),
  // periode berakhir pada 23:59:59.999 hari tersebut.
  const s = normalize_(v);
  if (/^\d{4}[\/-]\d{1,2}[\/-]\d{1,2}$/.test(s) || /^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}$/.test(s)) {
    d.setHours(23,59,59,999);
  }
  return d;
}

function parseDateSafe_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = normalize_(v);
  if (!s) return null;

  // ISO / normal JS date strings
  let d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  // dd/MM/yyyy [HH:mm[:ss]]
  let m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    d = new Date(
      Number(m[3]), Number(m[2]) - 1, Number(m[1]),
      Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)
    );
    if (!isNaN(d.getTime())) return d;
  }

  // yyyy/MM/dd or yyyy-MM-dd [HH:mm[:ss]]
  m = s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    d = new Date(
      Number(m[1]), Number(m[2]) - 1, Number(m[3]),
      Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)
    );
    if (!isNaN(d.getTime())) return d;
  }

  return null;
}

function dateOut_(v) {
  const d = parseDateSafe_(v);
  return d ? d.toISOString() : normalize_(v);
}

function findRowBy_(sh,hm,key,value) {
  if (!sh || !hm[key] || sh.getLastRow() < 2) return -1;

  const col = hm[key];
  const vals = sh.getRange(2,col,sh.getLastRow()-1,1).getValues();

  for (let i=0;i<vals.length;i++) {
    if (normalize_(vals[i][0]) === normalize_(value)) return i+2;
  }
  return -1;
}

function normalizeWa_(value) {
  let s = String(value || '').replace(/[^\d]/g,'');
  if (!s) return '';
  if (s.indexOf('62') === 0) return s;
  if (s.indexOf('0') === 0) return '62' + s.slice(1);
  if (s.indexOf('8') === 0) return '62' + s;
  return s;
}

function findParticipant_(id) {
  const sh = sheet_('PARTICIPANTS');
  if (!sh) return null;

  const hm = headerMap_(sh);
  const row = findRowBy_(sh,hm,'ID',id);
  if (row < 2) return null;

  const r = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];

  return {
    id:val_(r,hm,'ID'),
    quizId:val_(r,hm,'QuizID'),
    timestamp:dateOut_(val_(r,hm,'Timestamp')),
    nama:val_(r,hm,'Nama'),
    wa:val_(r,hm,'WhatsApp'),
    email:val_(r,hm,'Email'),
    answers:parseJson_(val_(r,hm,'AnswersJSON'),[]),
    total:Number(val_(r,hm,'Total')||0),
    status:val_(r,hm,'Status') || 'Masuk'
  };
}

// ============================================================
// SETUP / DIAGNOSTIC
// ============================================================

function setupQuizEngine() {
  ensureDatabase_();
  return 'Universal Quiz Engine V3.2 database siap.';
}

function upgradeQuizEngineV32() {
  ensureDatabase_();

  // Tambahkan kolom hadiah tanpa menghapus data lama.
  const p = sheet_('PARTICIPANTS');
  ensureColumns_(p, [
    'HadiahStatus',
    'HadiahContactedAt',
    'HadiahDiberikanAt',
    'HadiahCatatan'
  ]);

  // Default status hanya untuk sel kosong.
  const hm = headerMap_(p);
  if (hm.HadiahStatus && p.getLastRow() >= 2) {
    const range = p.getRange(2,hm.HadiahStatus,p.getLastRow()-1,1);
    const vals = range.getValues();
    let changed = false;

    vals.forEach(r => {
      if (!normalize_(r[0])) {
        r[0] = 'Belum Diproses';
        changed = true;
      }
    });

    if (changed) range.setValues(vals);
  }

  return 'Universal Quiz Engine V3.2 siap. Struktur hadiah ditambahkan tanpa menghapus data lama.';
}

function DEBUG_UQE_V32() {
  const ss = getSS_();

  const result = {
    spreadsheet:ss.getName(),
    spreadsheetId:ss.getId(),
    sheets:{
      quizzes:!!sheet_('QUIZZES'),
      questions:!!sheet_('QUESTIONS'),
      materials:!!sheet_('MATERIALS'),
      participants:!!sheet_('PARTICIPANTS'),
      scores:!!sheet_('SCORES')
    },
    adminLogin:adminLogin(UQE_CONFIG.ADMIN_USER,UQE_CONFIG.ADMIN_PASS)
  };

  Logger.log(JSON.stringify(result,null,2));
  return result;
}

function getPublicParticipants(quizId) {
  try {
    ensureDatabase_();
    const qid = normalize_(quizId);
    if (!qid) return { ok: false, message: 'QuizID kosong.' };

    const sh = sheet_('PARTICIPANTS');
    if (!sh) return { ok: true, participants: [] };

    const hm = headerMap_(sh);
    const rows = dataRows_(sh)
      .filter(r => normalize_(val_(r, hm, 'QuizID')) === qid)
      .map(r => ({
        nama: val_(r, hm, 'Nama'),
        timestamp: dateOut_(val_(r, hm, 'Timestamp')),
        duration: val_(r, hm, 'Duration') || val_(r, hm, 'Durasi') || '-',
        total: Number(val_(r, hm, 'Total') || 0),
        status: val_(r, hm, 'Status') || 'Masuk'
      }))
      .sort((a, b) => {
        const diff = Number(b.total) - Number(a.total);
        if (diff !== 0) return diff;
        return 0;
      });

    return { ok: true, participants: rows };
  } catch (e) {
    return { ok: false, message: e.message || String(e) };
  }
}
