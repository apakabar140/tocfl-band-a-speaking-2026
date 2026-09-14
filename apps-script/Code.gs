const CONFIG = Object.freeze({
  MEMBER_SHEET_ID: '1quA4P4-b4cUU-9Mh_L949N4QQ9qDoHs9rLPWsB5K-ig',
  QUESTION_SHEET_ID: '1z9UtZTHph_V3tvtxq_Eh4aHjs4UNgQfAZVmhg_tuDAY',
  MEMBER_TAB: '會員名單',
  PROFILE_TAB: 'Supabase會員備份',
  PROGRESS_TAB: '逐題進度',
  SESSION_SECONDS: 21600,
});

function doGet() {
  return json_({ ok: true, service: 'tocfl-speaking-band-a' });
}

function setupCheck() {
  const memberBook = SpreadsheetApp.openById(CONFIG.MEMBER_SHEET_ID);
  const questionBook = SpreadsheetApp.openById(CONFIG.QUESTION_SHEET_ID);
  const requiredMemberTabs = [CONFIG.MEMBER_TAB, CONFIG.PROFILE_TAB, CONFIG.PROGRESS_TAB];
  const requiredQuestionTabs = ['回答問題', '經驗描述', '影片描述', '平台介面翻譯表'];
  const missingMemberTabs = requiredMemberTabs.filter(name => !memberBook.getSheetByName(name));
  const missingQuestionTabs = requiredQuestionTabs.filter(name => !questionBook.getSheetByName(name));
  if (missingMemberTabs.length || missingQuestionTabs.length) {
    throw new Error('缺少必要分頁：' + missingMemberTabs.concat(missingQuestionTabs).join('、'));
  }
  console.log('設定檢查完成：會員與題庫試算表均可存取，必要分頁完整。');
  return true;
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const routes = {
      content: () => content_(),
      login: () => login_(body),
      restore: () => restore_(body.token),
      recordPractice: () => recordPractice_(body.token, body.questionId),
      recordMock: () => recordMock_(body.token),
      logout: () => logout_(body.token),
    };
    if (!routes[body.action]) throw new Error('unknown_action');
    return json_(routes[body.action]());
  } catch (error) {
    console.error(error);
    return json_({ ok: false, error: 'request_failed' });
  }
}

function content_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('published_content_v1');
  if (cached) return JSON.parse(cached);
  const book = SpreadsheetApp.openById(CONFIG.QUESTION_SHEET_ID);
  const result = {
    ok: true,
    questions: {
      question: questionRows_(book.getSheetByName('回答問題'), false),
      experience: questionRows_(book.getSheetByName('經驗描述'), false),
      sequence: questionRows_(book.getSheetByName('影片描述'), true),
    },
    translations: translationRows_(book.getSheetByName('平台介面翻譯表')),
  };
  cache.put('published_content_v1', JSON.stringify(result), 300);
  return result;
}

function questionRows_(sheet, isSequence) {
  const values = sheet.getDataRange().getDisplayValues();
  const header = headerMap_(values[2]);
  return values.slice(3)
    .filter(r => r[header['題目 ID']] && r[header['中文題目']] && r[header['題目狀態']] === '網站使用中')
    .map(r => {
      const item = {
        id: String(r[header['題目 ID']]).toLowerCase(),
        prompt: r[header['中文題目']],
        hints: splitLines_(r[header['提示詞']]),
        frames: splitLines_(r[header['句型鷹架']]),
        checks: splitLines_(r[header['自我檢查']]).map(x => x.replace(/^\d+[.、]\s*/, '')),
        prep: Number(r[header['準備秒數']] || 0),
        answer: Number(r[header['回答秒數']] || 0),
        sort: Number(r[header['題目排序']] || 9999),
        version: r[header['版本']] || '',
      };
      if (isSequence) {
        item.scenes = ['第 1 格', '第 2 格', '第 3 格', '第 4 格']
          .map(name => scene_(r[header[name]]))
          .filter(Boolean);
        item.frameSeconds = Number(r[header['每格秒數']] || 3);
        item.mediaFolder = r[header['素材資料夾']] || '';
      }
      return item;
    })
    .sort((a, b) => a.sort - b.sort || a.id.localeCompare(b.id));
}

function translationRows_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const header = headerMap_(values[1]);
  const languages = { zh: '中文', id: '印尼文', en: '英文', vi: '越南文', th: '泰文' };
  const out = { zh: {}, id: {}, en: {}, vi: {}, th: {} };
  values.slice(2)
    .filter(r => r[header['文字 ID']] && r[header['網站發布狀態']] === '網站使用中')
    .forEach(r => Object.keys(languages).forEach(code => {
      const value = r[header[languages[code]]];
      if (value) out[code][r[header['文字 ID']]] = value;
    }));
  return out;
}

function splitLines_(value) {
  return String(value || '').split(/\r?\n/).map(x => x.trim()).filter(Boolean);
}

function scene_(value) {
  const text = String(value || '').trim();
  if (!text) return null;
  const parts = text.split(/\s+/);
  return [parts.shift(), parts.join(' ')];
}

function login_(body) {
  const passport = normalizePassport_(body.passport);
  const birthday = normalizeBirthday_(body.birthday);
  if (!passport || birthday.length !== 8) return { ok: false, error: 'invalid_login' };

  const sheet = SpreadsheetApp.openById(CONFIG.MEMBER_SHEET_ID).getSheetByName(CONFIG.MEMBER_TAB);
  const values = sheet.getDataRange().getDisplayValues();
  const header = headerMap_(values[0]);
  const row = values.slice(1).find(r =>
    normalizePassport_(r[header['護照號碼']]) === passport &&
    normalizeBirthday_(r[header['出生年月日']]) === birthday
  );
  if (!row || String(row[header['開放使用']]).trim() !== '是') {
    return { ok: false, error: 'invalid_login' };
  }

  const passportHash = sha256_(passport);
  const profile = findProfileByHash_(passportHash);
  if (!profile) return { ok: false, error: 'profile_missing' };

  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put(token, JSON.stringify({ userId: profile.userId }), CONFIG.SESSION_SECONDS);
  touchMember_(passport, 'login');
  return { ok: true, token, profile: publicProfile_(profile), completedIds: completedIds_(profile.userId) };
}

function restore_(token) {
  const session = session_(token);
  const profile = findProfileByUserId_(session.userId);
  return { ok: true, profile: publicProfile_(profile), completedIds: completedIds_(session.userId) };
}

function recordPractice_(token, questionId) {
  const session = session_(token);
  const id = String(questionId || '').toLowerCase().trim();
  if (!/^[qes]\d{1,3}$/.test(id)) throw new Error('invalid_question');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.MEMBER_SHEET_ID).getSheetByName(CONFIG.PROGRESS_TAB);
    const values = sheet.getDataRange().getDisplayValues();
    const index = values.slice(2).findIndex(r => r[0] === session.userId && r[1].toLowerCase() === id);
    const now = new Date();
    if (index >= 0) {
      const rowNumber = index + 3;
      sheet.getRange(rowNumber, 3, 1, 2).setValues([[Number(values[rowNumber - 1][2] || 0) + 1, now]]);
    } else {
      sheet.appendRow([session.userId, id, 1, now]);
    }
    updateMemberStats_(session.userId, false);
  } finally {
    lock.releaseLock();
  }
  return restore_(token);
}

function recordMock_(token) {
  const session = session_(token);
  updateMemberStats_(session.userId, true);
  return restore_(token);
}

function logout_(token) {
  if (token) CacheService.getScriptCache().remove(String(token));
  return { ok: true };
}

function session_(token) {
  const raw = token && CacheService.getScriptCache().get(String(token));
  if (!raw) throw new Error('invalid_session');
  CacheService.getScriptCache().put(String(token), raw, CONFIG.SESSION_SECONDS);
  return JSON.parse(raw);
}

function profiles_() {
  const values = SpreadsheetApp.openById(CONFIG.MEMBER_SHEET_ID).getSheetByName(CONFIG.PROFILE_TAB).getDataRange().getDisplayValues();
  return values.slice(2).map(r => ({
    userId: r[0], passportHash: r[1], displayName: r[2], interfaceLanguage: r[3],
    enabled: String(r[4]).toLowerCase() === 'true', completedQuestions: Number(r[5] || 0),
    practiceCount: Number(r[6] || 0), mockCount: Number(r[7] || 0),
  }));
}

function findProfileByHash_(hash) { return profiles_().find(p => p.passportHash === hash); }
function findProfileByUserId_(id) { return profiles_().find(p => p.userId === id); }
function publicProfile_(p) {
  if (!p || !p.enabled) throw new Error('profile_disabled');
  return { display_name:p.displayName, interface_language:p.interfaceLanguage, enabled:p.enabled,
    completed_questions:p.completedQuestions, practice_count:p.practiceCount, mock_count:p.mockCount };
}

function completedIds_(userId) {
  const values = SpreadsheetApp.openById(CONFIG.MEMBER_SHEET_ID).getSheetByName(CONFIG.PROGRESS_TAB).getDataRange().getDisplayValues();
  return values.slice(2).filter(r => r[0] === userId).map(r => r[1]);
}

function updateMemberStats_(userId, isMock) {
  const profileSheet = SpreadsheetApp.openById(CONFIG.MEMBER_SHEET_ID).getSheetByName(CONFIG.PROFILE_TAB);
  const profiles = profileSheet.getDataRange().getDisplayValues();
  const rowIndex = profiles.findIndex((r, i) => i > 1 && r[0] === userId);
  if (rowIndex < 0) throw new Error('profile_missing');
  const progress = SpreadsheetApp.openById(CONFIG.MEMBER_SHEET_ID)
    .getSheetByName(CONFIG.PROGRESS_TAB).getDataRange().getDisplayValues()
    .slice(2).filter(r => r[0] === userId);
  const completed = new Set(progress.map(r => r[1])).size;
  const practice = progress.reduce((sum, r) => sum + Number(r[2] || 0), 0);
  const mock = Number(profiles[rowIndex][7] || 0) + (isMock ? 1 : 0);
  profileSheet.getRange(rowIndex + 1, 6, 1, 3).setValues([[
    completed, practice, mock
  ]]);
  syncMemberStats_(profiles[rowIndex][1], completed, practice, mock);
}

function syncMemberStats_(passportHash, completed, practice, mock) {
  const sheet = SpreadsheetApp.openById(CONFIG.MEMBER_SHEET_ID).getSheetByName(CONFIG.MEMBER_TAB);
  const values = sheet.getDataRange().getDisplayValues();
  const h = headerMap_(values[0]);
  const index = values.slice(1).findIndex(r => sha256_(normalizePassport_(r[h['護照號碼']])) === passportHash);
  if (index < 0) throw new Error('member_missing');
  const rowNumber = index + 2;
  if (h['完成題數'] !== undefined) sheet.getRange(rowNumber, h['完成題數'] + 1).setValue(completed);
  if (h['練習次數'] !== undefined) sheet.getRange(rowNumber, h['練習次數'] + 1).setValue(practice);
  if (h['完整模考次數'] !== undefined) sheet.getRange(rowNumber, h['完整模考次數'] + 1).setValue(mock);
  if (h['最近練習日期'] !== undefined) sheet.getRange(rowNumber, h['最近練習日期'] + 1).setValue(new Date());
}

function touchMember_(passport, type) {
  const sheet = SpreadsheetApp.openById(CONFIG.MEMBER_SHEET_ID).getSheetByName(CONFIG.MEMBER_TAB);
  const values = sheet.getDataRange().getDisplayValues();
  const h = headerMap_(values[0]);
  const index = values.slice(1).findIndex(r => normalizePassport_(r[h['護照號碼']]) === passport);
  if (index < 0) return;
  const name = type === 'login' ? '最近登入日期' : '最近練習日期';
  if (h[name] !== undefined) sheet.getRange(index + 2, h[name] + 1).setValue(new Date());
}

function headerMap_(row) { const out={}; row.forEach((v,i)=>{ if(String(v).trim()) out[String(v).trim()]=i; }); return out; }
function normalizePassport_(v) { return String(v || '').toUpperCase().replace(/\s+/g, '').trim(); }
function normalizeBirthday_(v) { return String(v || '').replace(/\D/g, '').trim(); }
function sha256_(v) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, v, Utilities.Charset.UTF_8).map(b => ('0'+((b+256)%256).toString(16)).slice(-2)).join(''); }
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
