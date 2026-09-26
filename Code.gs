
/**
 * ระบบรับสมัครนักเรียนออนไลน์
 * Google Apps Script + Google Sheets + Google Drive
 *
 * สิ่งสำคัญ:
 * 1) เปิดโปรเจกต์นี้ใน Google Apps Script
 * 2) Run setupSystem() ครั้งแรกด้วยบัญชีเจ้าของระบบ
 * 3) Deploy > New deployment > Web app
 * 4) Execute as: Me
 * 5) Who has access: ตามนโยบายของโรงเรียน
 *
 * ระบบสร้างฐานข้อมูลและโฟลเดอร์ให้อัตโนมัติ ไม่มีข้อมูลผู้สมัครจำลอง
 */

const APP = {
  TZ: 'Asia/Bangkok',
  SESSION_HOURS: 8,
  SHEETS: {
    SETTINGS: 'Settings',
    ROUNDS: 'AdmissionRounds',
    PROGRAMS: 'Programs',
    STUDENTS: 'Students',
    APPLICATIONS: 'Applications',
    DOCUMENTS: 'Documents',
    PAYMENTS: 'Payments',
    EXAMS: 'Exams',
    SCORES: 'Scores',
    USERS: 'Users',
    AUDIT: 'AuditLog'
  },
  DOC_TYPES: {
    PHOTO: 'รูปถ่าย',
    ID_CARD: 'สำเนาบัตรประชาชน',
    HOUSE: 'สำเนาทะเบียนบ้าน',
    TRANSCRIPT: 'ปพ.1',
    PAYMENT: 'หลักฐานการชำระเงิน'
  },
  STATUS: {
    APPLICATION_DRAFT: 'ร่าง',
    APPLICATION_SUBMITTED: 'รอตรวจสอบ',
    DOCUMENT_FIX: 'เอกสารไม่สมบูรณ์',
    ELIGIBLE: 'มีสิทธิ์สอบ',
    EXAM_PASSED: 'สอบผ่าน',
    SELECTED: 'ผ่านการคัดเลือก',
    WAITLIST: 'ตัวสำรอง',
    NOT_SELECTED: 'ไม่ผ่านการคัดเลือก',
    CONFIRMED: 'ยืนยันสิทธิ์',
    CANCELLED: 'ยกเลิก'
  }
};

const HEADERS = {
  Settings: ['key','value','updatedAt'],
  AdmissionRounds: [
    'roundId','academicYear','roundName','roundType','description',
    'openDate','closeDate','examDate','resultDate','confirmStart','confirmEnd',
    'fee','status','maxApplicants','createdAt','updatedAt','createdBy'
  ],
  Programs: [
    'programId','roundId','programName','programCode','quotaName',
    'quotaType','capacity','minimumGpa','description','status','createdAt','updatedAt'
  ],
  Students: [
    'studentId','citizenId','prefix','firstName','lastName','birthDate',
    'phone','email','passwordHash','passwordSet','registeredAt','updatedAt',
    'pdpaConsent','pdpaConsentAt','status'
  ],
  Applications: [
    'applicationId','studentId','roundId','programId','applicationNo',
    'status','submittedAt','updatedAt','prefix','firstName','lastName',
    'nickname','birthDate','nationality','religion','phone','email',
    'addressNo','village','soi','road','subdistrict','district','province','postalCode',
    'parent1Prefix','parent1Name','parent1Relation','parent1Occupation','parent1Phone',
    'parent2Prefix','parent2Name','parent2Relation','parent2Occupation','parent2Phone',
    'schoolName','schoolProvince','graduationYear','gpa','studentCode',
    'specialNeeds','disability','emergencyName','emergencyPhone','notes',
    'createdAt','createdBy','updatedBy'
  ],
  Documents: [
    'documentId','applicationId','studentId','documentType','fileName','mimeType',
    'driveFileId','driveUrl','status','reviewNote','uploadedAt','reviewedAt','reviewedBy'
  ],
  Payments: [
    'paymentId','applicationId','studentId','roundId','amount','method',
    'referenceNo','slipDocumentId','bankDate','bankTime','status','reviewNote',
    'paidAt','verifiedAt','verifiedBy','createdAt'
  ],
  Exams: [
    'examId','applicationId','studentId','roundId','examDate','building',
    'room','seatNo','examTime','status','createdAt','updatedAt'
  ],
  Scores: [
    'scoreId','examId','applicationId','studentId','roundId','subject',
    'score','maxScore','importBatch','createdAt','createdBy'
  ],
  Users: [
    'userId','username','passwordHash','displayName','role','permissions',
    'status','createdAt','updatedAt','lastLoginAt'
  ],
  AuditLog: [
    'logId','timestamp','actorType','actorId','action','targetType','targetId','details'
  ]
};

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appName = getSetting('schoolName') || 'ระบบรับสมัครนักเรียน';
  return t.evaluate()
    .setTitle(t.appName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================
 * INITIAL SETUP / DATABASE
 * ========================= */

function setupSystem() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  let ss;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (err) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('ระบบรับสมัครนักเรียน - ฐานข้อมูล');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  Object.keys(HEADERS).forEach(function(sheetName) {
    const headers = HEADERS[sheetName];
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);
    ensureHeaders_(sh, headers);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length)
      .setFontWeight('bold')
      .setBackground('#0F2C59')
      .setFontColor('#FFFFFF');
  });

  let defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && Object.keys(HEADERS).indexOf(defaultSheet.getName()) === -1 &&
      ss.getSheets().length > Object.keys(HEADERS).length) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }

  setupDriveFolders_();

  // บัญชีผู้ดูแลเริ่มต้นตามที่ผู้ใช้กำหนด
  const users = getAllRows_(APP.SHEETS.USERS);
  const existing = users.find(function(u){ return String(u.username) === 'sxaiq54'; });
  if (!existing) {
    appendRow_(APP.SHEETS.USERS, {
      userId: uid_('USR'),
      username: 'sxaiq54',
      passwordHash: hashPassword_('Sxxnga2011.54'),
      displayName: 'ผู้ดูแลระบบ',
      role: 'ADMIN',
      permissions: JSON.stringify(['*']),
      status: 'ACTIVE',
      createdAt: now_(),
      updatedAt: now_(),
      lastLoginAt: ''
    });
  }

  setDefaultSettings_();

  logAudit_('SYSTEM','setupSystem','SYSTEM','SYSTEM','สร้าง/ตรวจสอบโครงสร้างระบบ','');
  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    driveRootId: PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID')
  };
}

function setDefaultSettings_() {
  const defaults = {
    schoolName: 'ระบบรับสมัครนักเรียน',
    schoolLogoFileId: '',
    schoolLogoUrl: '',
    bannerFileId: '',
    bannerUrl: '',
    mapFileId: '',
    mapUrl: '',
    homeTitle: 'ระบบรับสมัครนักเรียน',
    homeSubtitle: 'ระบบรับสมัครนักเรียนออนไลน์',
    homeNotice: 'กรุณาตรวจสอบประกาศและรอบการรับสมัครก่อนเริ่มสมัคร',
    pdpaText: 'ข้าพเจ้ายินยอมให้สถานศึกษาจัดเก็บ ใช้ และประมวลผลข้อมูลส่วนบุคคลเพื่อดำเนินการรับสมัคร คัดเลือก สอบ ประกาศผล และงานที่เกี่ยวข้องตามกฎหมายคุ้มครองข้อมูลส่วนบุคคล',
    applicationFeeDefault: '0',
    contactPhone: '',
    contactEmail: '',
    address: '',
    updatedAt: now_()
  };
  Object.keys(defaults).forEach(function(k) {
    if (getSetting(k) === null) setSetting(k, defaults[k]);
  });
}

function setupDriveFolders_() {
  const props = PropertiesService.getScriptProperties();
  let rootId = props.getProperty('ROOT_FOLDER_ID');
  let root;
  if (rootId) {
    try { root = DriveApp.getFolderById(rootId); } catch (e) { root = null; }
  }
  if (!root) {
    root = DriveApp.createFolder('ระบบรับสมัครนักเรียน - เอกสาร');
    props.setProperty('ROOT_FOLDER_ID', root.getId());
  }
  const names = ['ผู้สมัคร','รูปถ่าย','บัตรประชาชน','ทะเบียนบ้าน','ปพ.1','หลักฐานการชำระเงิน','รายงาน PDF','Import'];
  names.forEach(function(name) {
    let found = false;
    const it = root.getFoldersByName(name);
    if (it.hasNext()) found = true;
    if (!found) root.createFolder(name);
  });
}

function ensureReady_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) setupSystem();
}

function getSS_() {
  ensureReady_();
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
}

function getSheet_(name) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    ensureHeaders_(sh, HEADERS[name] || []);
  }
  return sh;
}

function ensureHeaders_(sh, headers) {
  if (!headers || !headers.length) return;
  const current = sh.getLastColumn() ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] : [];
  const needs = current.length !== headers.length || headers.some(function(h,i){ return current[i] !== h; });
  if (needs) {
    sh.clear();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
}

function getAllRows_(sheetName) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).filter(function(row){
    return row.some(function(v){ return v !== '' && v !== null; });
  }).map(function(row){
    const o = {};
    headers.forEach(function(h,i){ o[h] = row[i]; });
    return o;
  });
}

function findOne_(sheetName, field, value) {
  const rows = getAllRows_(sheetName);
  const needle = String(value == null ? '' : value);
  return rows.find(function(r){ return String(r[field] == null ? '' : r[field]) === needle; }) || null;
}

function filterRows_(sheetName, predicate) {
  return getAllRows_(sheetName).filter(predicate);
}

function appendRow_(sheetName, obj) {
  const sh = getSheet_(sheetName);
  const headers = HEADERS[sheetName];
  const row = headers.map(function(h){ return obj[h] !== undefined && obj[h] !== null ? obj[h] : ''; });
  sh.appendRow(row);
  return obj;
}

function updateRowById_(sheetName, idField, idValue, patch) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) throw new Error('ไม่พบข้อมูล');
  const headers = values[0];
  const idx = headers.indexOf(idField);
  if (idx < 0) throw new Error('ไม่พบฟิลด์ ' + idField);
  for (let r=1; r<values.length; r++) {
    if (String(values[r][idx]) === String(idValue)) {
      Object.keys(patch).forEach(function(k) {
        const c = headers.indexOf(k);
        if (c >= 0) sh.getRange(r+1,c+1).setValue(patch[k]);
      });
      return true;
    }
  }
  throw new Error('ไม่พบข้อมูล');
}

function deleteRowById_(sheetName, idField, idValue) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = headers.indexOf(idField);
  for (let r=1;r<values.length;r++) {
    if (String(values[r][idx]) === String(idValue)) {
      sh.deleteRow(r+1);
      return true;
    }
  }
  return false;
}

function uid_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g,'').substring(0,16).toUpperCase();
}

function now_() {
  return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd HH:mm:ss');
}

function today_() {
  return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd');
}

function parseDate_(s) {
  if (!s) return null;
  if (Object.prototype.toString.call(s) === '[object Date]') return s;
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
}

function hashPassword_(password) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b){ return ('0' + ((b<0?b+256:b).toString(16))).slice(-2); }).join('');
}

function getSetting(key) {
  ensureReady_();
  const row = findOne_(APP.SHEETS.SETTINGS,'key',key);
  return row ? String(row.value == null ? '' : row.value) : null;
}

function setSetting(key,value) {
  const row = findOne_(APP.SHEETS.SETTINGS,'key',key);
  if (row) {
    updateRowById_(APP.SHEETS.SETTINGS,'key',key,{value:value,updatedAt:now_()});
  } else {
    appendRow_(APP.SHEETS.SETTINGS,{key:key,value:value,updatedAt:now_()});
  }
}

/* =========================
 * DRIVE / FILE UPLOAD
 * ========================= */

function getRootFolder_() {
  setupDriveFolders_();
  return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
}

function getChildFolder_(name) {
  const root = getRootFolder_();
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function saveBase64File_(fileObj, folderName, prefix) {
  if (!fileObj || !fileObj.data) throw new Error('ไม่พบไฟล์');
  const cleanName = String(fileObj.name || 'upload').replace(/[\\/:*?"<>|]/g,'_');
  const bytes = Utilities.base64Decode(String(fileObj.data).split(',').pop());
  const blob = Utilities.newBlob(bytes, fileObj.mimeType || 'application/octet-stream', cleanName);
  const folder = getChildFolder_(folderName);
  const file = folder.createFile(blob);
  file.setName((prefix || '') + cleanName);
  return {
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    url: 'https://drive.google.com/file/d/' + file.getId() + '/view'
  };
}

function validateUpload_(fileObj, maxMb) {
  if (!fileObj || !fileObj.data) throw new Error('กรุณาเลือกไฟล์');
  const size = Math.ceil((String(fileObj.data).length * 3) / 4 / 1024 / 1024);
  if (size > maxMb) throw new Error('ไฟล์มีขนาดเกิน ' + maxMb + ' MB');
  const allowed = [
    'image/jpeg','image/png','application/pdf',
    'image/webp'
  ];
  if (allowed.indexOf(String(fileObj.mimeType || '').toLowerCase()) < 0) {
    throw new Error('ชนิดไฟล์ไม่รองรับ กรุณาใช้ JPG, PNG, WEBP หรือ PDF');
  }
}

/* =========================
 * PUBLIC / BOOTSTRAP
 * ========================= */

function getPublicBootstrap() {
  ensureReady_();
  const rounds = getOpenRounds_().map(function(r){
    return {
      roundId:r.roundId, academicYear:r.academicYear, roundName:r.roundName,
      roundType:r.roundType, description:r.description,
      openDate:r.openDate, closeDate:r.closeDate, examDate:r.examDate,
      fee:Number(r.fee || 0), status:r.status, maxApplicants:Number(r.maxApplicants || 0),
      programs: getProgramsByRound_(r.roundId)
    };
  });
  const settings = {};
  ['schoolName','schoolLogoUrl','bannerUrl','homeTitle','homeSubtitle','homeNotice',
   'pdpaText','contactPhone','contactEmail','address','mapUrl'].forEach(function(k){ settings[k]=getSetting(k)||''; });
  return {ok:true,settings:settings,rounds:rounds};
}

function getOpenRounds_() {
  const rows = getAllRows_(APP.SHEETS.ROUNDS);
  const now = new Date();
  return rows.filter(function(r){
    const open = parseDate_(r.openDate);
    const close = parseDate_(r.closeDate);
    const active = String(r.status) === 'OPEN';
    return active && open && close && now >= open && now <= new Date(close.getTime()+86399999);
  });
}

function getProgramsByRound_(roundId) {
  return getAllRows_(APP.SHEETS.PROGRAMS)
    .filter(function(p){ return String(p.roundId)===String(roundId) && String(p.status)!=='INACTIVE'; })
    .map(function(p){
      const used = getAllRows_(APP.SHEETS.APPLICATIONS)
        .filter(function(a){ return String(a.roundId)===String(roundId) && String(a.programId)===String(p.programId) &&
          ['CANCELLED','NOT_SELECTED'].indexOf(String(a.status))<0; }).length;
      return Object.assign({},p,{
        capacity:Number(p.capacity||0),
        minimumGpa:Number(p.minimumGpa||0),
        remaining:Math.max(0,Number(p.capacity||0)-used)
      });
    });
}

/* =========================
 * STUDENT REGISTRATION / LOGIN
 * ========================= */

function registerStudent(data) {
  ensureReady_();
  const citizenId = normalizeCitizenId_(data.citizenId);
  validateCitizenId_(citizenId);
  if (!data.firstName || !data.lastName || !data.phone) throw new Error('กรุณากรอกข้อมูลที่จำเป็นให้ครบ');
  if (String(data.pdpaConsent) !== 'true') throw new Error('กรุณายอมรับความยินยอม PDPA ก่อนสมัคร');
  if (findOne_(APP.SHEETS.STUDENTS,'citizenId',citizenId)) throw new Error('เลขบัตรประชาชนนี้ลงทะเบียนแล้ว กรุณาเข้าสู่ระบบ');

  const student = {
    studentId:uid_('STD'),
    citizenId:citizenId,
    prefix:String(data.prefix||''),
    firstName:String(data.firstName).trim(),
    lastName:String(data.lastName).trim(),
    birthDate:String(data.birthDate||''),
    phone:String(data.phone).trim(),
    email:String(data.email||'').trim(),
    passwordHash:'',
    passwordSet:'NO',
    registeredAt:now_(),
    updatedAt:now_(),
    pdpaConsent:'YES',
    pdpaConsentAt:now_(),
    status:'ACTIVE'
  };
  appendRow_(APP.SHEETS.STUDENTS,student);
  logAudit_('STUDENT',student.studentId,'REGISTER','STUDENT',student.studentId,'สมัครบัญชีผู้สมัคร');
  return {ok:true,studentId:student.studentId,message:'สมัครสมาชิกสำเร็จ กรุณาเข้าสู่ระบบด้วยเลขบัตรประชาชน 13 หลัก'};
}

function normalizeCitizenId_(id) {
  return String(id||'').replace(/\D/g,'');
}

function validateCitizenId_(id) {
  if (!/^\d{13}$/.test(id)) throw new Error('เลขบัตรประชาชนต้องมี 13 หลัก');
  // ตรวจ checksum บัตรประชาชนไทย
  let sum=0;
  for(let i=0;i<12;i++) sum += Number(id.charAt(i))*(13-i);
  const check=(11-(sum%11))%10;
  if(check !== Number(id.charAt(12))) throw new Error('เลขบัตรประชาชนไม่ถูกต้อง');
  return true;
}

function loginStudent(citizenId) {
  const id=normalizeCitizenId_(citizenId);
  validateCitizenId_(id);
  const student=findOne_(APP.SHEETS.STUDENTS,'citizenId',id);
  if(!student || String(student.status)!=='ACTIVE') throw new Error('ไม่พบบัญชีผู้สมัคร กรุณาสมัครสมาชิกก่อน');
  const token=createSession_('STUDENT',student.studentId,{citizenId:id});
  return {ok:true,token:token,student:safeStudent_(student),dashboard:getStudentDashboard_(student.studentId)};
}

function safeStudent_(s) {
  const o=Object.assign({},s);
  delete o.passwordHash;
  return o;
}

/* =========================
 * OFFICER / ADMIN AUTH
 * ========================= */

function loginStaff(username,password) {
  const u=findOne_(APP.SHEETS.USERS,'username',String(username||'').trim());
  if(!u || String(u.status)!=='ACTIVE') throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  if(String(u.passwordHash)!==hashPassword_(password||'')) throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  updateRowById_(APP.SHEETS.USERS,'userId',u.userId,{lastLoginAt:now_(),updatedAt:now_()});
  const token=createSession_('STAFF',u.userId,{role:u.role,username:u.username});
  logAudit_('STAFF',u.userId,'LOGIN','USER',u.userId,'เข้าสู่ระบบ');
  return {ok:true,token:token,user:safeUser_(u)};
}

function safeUser_(u) {
  const o=Object.assign({},u); delete o.passwordHash; return o;
}

function createSession_(type,id,data) {
  const token=Utilities.getUuid()+'-'+Utilities.getUuid();
  CacheService.getScriptCache().put('SESSION_'+token,JSON.stringify({
    type:type,id:id,data:data,createdAt:Date.now()
  }),APP.SESSION_HOURS*3600);
  return token;
}

function getSession_(token) {
  if(!token) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const raw=CacheService.getScriptCache().get('SESSION_'+token);
  if(!raw) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  return JSON.parse(raw);
}

function requireStaff_(token, roles) {
  const s=getSession_(token);
  if(s.type!=='STAFF') throw new Error('ไม่มีสิทธิ์');
  if(roles && roles.length && roles.indexOf(s.data.role)<0 && s.data.role!=='ADMIN') throw new Error('ไม่มีสิทธิ์สำหรับรายการนี้');
  return s;
}

function requireStudent_(token) {
  const s=getSession_(token);
  if(s.type!=='STUDENT') throw new Error('ไม่มีสิทธิ์');
  return s;
}

/* =========================
 * STUDENT DASHBOARD / APPLICATION
 * ========================= */

function getStudentDashboard(token) {
  const s=requireStudent_(token);
  const student=findOne_(APP.SHEETS.STUDENTS,'studentId',s.id);
  return {ok:true,student:safeStudent_(student),dashboard:getStudentDashboard_(s.id)};
}

function getStudentDashboard_(studentId) {
  const apps=getAllRows_(APP.SHEETS.APPLICATIONS).filter(function(a){return String(a.studentId)===String(studentId);});
  return apps.map(function(a){
    const round=findOne_(APP.SHEETS.ROUNDS,'roundId',a.roundId)||{};
    const program=findOne_(APP.SHEETS.PROGRAMS,'programId',a.programId)||{};
    const docs=getAllRows_(APP.SHEETS.DOCUMENTS).filter(function(d){return String(d.applicationId)===String(a.applicationId);});
    const exam=findOne_(APP.SHEETS.EXAMS,'applicationId',a.applicationId);
    const payment=findOne_(APP.SHEETS.PAYMENTS,'applicationId',a.applicationId);
    return {
      application:a,round:round,program:program,documents:docs,exam:exam,payment:payment
    };
  });
}

function createApplication(token,data) {
  const s=requireStudent_(token);
  const student=findOne_(APP.SHEETS.STUDENTS,'studentId',s.id);
  if(!student) throw new Error('ไม่พบบัญชีผู้สมัคร');
  if(String(student.pdpaConsent)!=='YES') throw new Error('ไม่พบการยินยอม PDPA');

  const round=findOne_(APP.SHEETS.ROUNDS,'roundId',data.roundId);
  if(!round || String(round.status)!=='OPEN') throw new Error('รอบรับสมัครนี้ไม่เปิดรับสมัคร');
  const open=parseDate_(round.openDate), close=parseDate_(round.closeDate), now=new Date();
  if(!open || !close || now<open || now>new Date(close.getTime()+86399999)) throw new Error('หมดระยะเวลารับสมัครแล้ว');

  const existing=getAllRows_(APP.SHEETS.APPLICATIONS).find(function(a){
    return String(a.studentId)===String(s.id) && String(a.roundId)===String(data.roundId) &&
      ['CANCELLED'].indexOf(String(a.status))<0;
  });
  if(existing) throw new Error('คุณมีใบสมัครในรอบนี้แล้ว');

  const program=findOne_(APP.SHEETS.PROGRAMS,'programId',data.programId);
  if(!program) throw new Error('ไม่พบแผนการเรียน');
  const usedRound=getAllRows_(APP.SHEETS.APPLICATIONS).filter(function(a){
    return String(a.roundId)===String(data.roundId)&&['CANCELLED'].indexOf(String(a.status))<0;
  }).length;
  if(Number(round.maxApplicants||0)>0 && usedRound>=Number(round.maxApplicants)) throw new Error('รอบรับสมัครเต็มตามจำนวนที่กำหนด');
  const used=getAllRows_(APP.SHEETS.APPLICATIONS).filter(function(a){
    return String(a.roundId)===String(data.roundId)&&String(a.programId)===String(data.programId)&&
      ['CANCELLED','NOT_SELECTED'].indexOf(String(a.status))<0;
  }).length;
  if(used>=Number(program.capacity||0)) throw new Error('แผนการเรียนนี้เต็มแล้ว');

  if(!data.firstName || !data.lastName || !data.birthDate || !data.phone) throw new Error('กรุณากรอกข้อมูลส่วนตัวให้ครบ');
  if(!data.parent1Name || !data.parent1Phone) throw new Error('กรุณากรอกข้อมูลผู้ปกครองให้ครบ');
  if(!data.schoolName || !data.gpa) throw new Error('กรุณากรอกข้อมูลการศึกษาให้ครบ');
  if(Number(data.gpa)<0 || Number(data.gpa)>4) throw new Error('เกรดเฉลี่ยต้องอยู่ระหว่าง 0.00 - 4.00');
  if(Number(data.gpa)<Number(program.minimumGpa||0)) throw new Error('GPA ต่ำกว่าเกณฑ์ขั้นต่ำของแผนการเรียนนี้');

  const appId=uid_('APP');
  const applicationNo=generateApplicationNo_(round,program);
  const app={
    applicationId:appId,studentId:s.id,roundId:data.roundId,programId:data.programId,
    applicationNo:applicationNo,status:APP.STATUS.APPLICATION_SUBMITTED,submittedAt:now_(),updatedAt:now_(),
    prefix:String(data.prefix||''),firstName:String(data.firstName).trim(),lastName:String(data.lastName).trim(),
    nickname:String(data.nickname||''),birthDate:String(data.birthDate||''),nationality:String(data.nationality||'ไทย'),
    religion:String(data.religion||''),phone:String(data.phone||''),email:String(data.email||''),
    addressNo:String(data.addressNo||''),village:String(data.village||''),soi:String(data.soi||''),
    road:String(data.road||''),subdistrict:String(data.subdistrict||''),district:String(data.district||''),
    province:String(data.province||''),postalCode:String(data.postalCode||''),
    parent1Prefix:String(data.parent1Prefix||''),parent1Name:String(data.parent1Name||''),
    parent1Relation:String(data.parent1Relation||''),parent1Occupation:String(data.parent1Occupation||''),
    parent1Phone:String(data.parent1Phone||''),
    parent2Prefix:String(data.parent2Prefix||''),parent2Name:String(data.parent2Name||''),
    parent2Relation:String(data.parent2Relation||''),parent2Occupation:String(data.parent2Occupation||''),
    parent2Phone:String(data.parent2Phone||''),schoolName:String(data.schoolName||''),
    schoolProvince:String(data.schoolProvince||''),graduationYear:String(data.graduationYear||''),
    gpa:String(data.gpa||''),studentCode:String(data.studentCode||''),
    specialNeeds:String(data.specialNeeds||''),disability:String(data.disability||''),
    emergencyName:String(data.emergencyName||''),emergencyPhone:String(data.emergencyPhone||''),
    notes:String(data.notes||''),createdAt:now_(),createdBy:'STUDENT',updatedBy:'STUDENT'
  };
  appendRow_(APP.SHEETS.APPLICATIONS,app);
  logAudit_('STUDENT',s.id,'CREATE','APPLICATION',appId,'สร้างใบสมัคร '+applicationNo);
  return {ok:true,application:app};
}

function generateApplicationNo_(round,program) {
  const prefix=String(round.academicYear||'0000').replace(/\D/g,'').slice(-4);
  const rows=getAllRows_(APP.SHEETS.APPLICATIONS);
  let max=0;
  rows.forEach(function(a){
    const m=String(a.applicationNo||'').match(/(\d{4})$/);
    if(m) max=Math.max(max,Number(m[1]));
  });
  return 'APP-'+prefix+'-'+String(max+1).padStart(4,'0');
}

function uploadStudentDocument(token,data) {
  const s=requireStudent_(token);
  const app=findOne_(APP.SHEETS.APPLICATIONS,'applicationId',data.applicationId);
  if(!app || String(app.studentId)!==String(s.id)) throw new Error('ไม่พบใบสมัคร');
  validateUpload_(data.file,10);
  const folderMap={};
  folderMap[APP.DOC_TYPES.PHOTO]='รูปถ่าย';
  folderMap[APP.DOC_TYPES.ID_CARD]='บัตรประชาชน';
  folderMap[APP.DOC_TYPES.HOUSE]='ทะเบียนบ้าน';
  folderMap[APP.DOC_TYPES.TRANSCRIPT]='ปพ.1';
  folderMap[APP.DOC_TYPES.PAYMENT]='หลักฐานการชำระเงิน';
  const saved=saveBase64File_(data.file,folderMap[data.documentType]||'ผู้สมัคร',app.applicationNo+'_');
  const doc={
    documentId:uid_('DOC'),applicationId:app.applicationId,studentId:s.id,
    documentType:data.documentType,fileName:saved.name,mimeType:saved.mimeType,
    driveFileId:saved.id,driveUrl:saved.url,status:'PENDING',reviewNote:'',
    uploadedAt:now_(),reviewedAt:'',reviewedBy:''
  };
  appendRow_(APP.SHEETS.DOCUMENTS,doc);
  updateRowById_(APP.SHEETS.APPLICATIONS,'applicationId',app.applicationId,{updatedAt:now_()});
  logAudit_('STUDENT',s.id,'UPLOAD','DOCUMENT',doc.documentId,doc.documentType);
  return {ok:true,document:doc};
}

function submitPaymentSlip(token,data) {
  const s=requireStudent_(token);
  const app=findOne_(APP.SHEETS.APPLICATIONS,'applicationId',data.applicationId);
  if(!app || String(app.studentId)!==String(s.id)) throw new Error('ไม่พบใบสมัคร');
  const round=findOne_(APP.SHEETS.ROUNDS,'roundId',app.roundId);
  const fee=Number(round ? round.fee : 0);
  if(!data.file) throw new Error('กรุณาแนบสลิป');
  validateUpload_(data.file,10);
  const saved=saveBase64File_(data.file,'หลักฐานการชำระเงิน',app.applicationNo+'_');
  const doc={
    documentId:uid_('DOC'),applicationId:app.applicationId,studentId:s.id,
    documentType:APP.DOC_TYPES.PAYMENT,fileName:saved.name,mimeType:saved.mimeType,
    driveFileId:saved.id,driveUrl:saved.url,status:'PENDING',reviewNote:'',
    uploadedAt:now_(),reviewedAt:'',reviewedBy:''
  };
  appendRow_(APP.SHEETS.DOCUMENTS,doc);
  const pay={
    paymentId:uid_('PAY'),applicationId:app.applicationId,studentId:s.id,roundId:app.roundId,
    amount:fee,method:String(data.method||'โอนเงิน'),referenceNo:String(data.referenceNo||''),
    slipDocumentId:doc.documentId,bankDate:String(data.bankDate||''),bankTime:String(data.bankTime||''),
    status:'PENDING',reviewNote:'',paidAt:now_(),verifiedAt:'',verifiedBy:'',createdAt:now_()
  };
  appendRow_(APP.SHEETS.PAYMENTS,pay);
  logAudit_('STUDENT',s.id,'PAYMENT_SUBMIT','PAYMENT',pay.paymentId,'แจ้งชำระเงิน');
  return {ok:true,payment:pay};
}

function confirmRight(token,applicationId) {
  const s=requireStudent_(token);
  const app=findOne_(APP.SHEETS.APPLICATIONS,'applicationId',applicationId);
  if(!app || String(app.studentId)!==String(s.id)) throw new Error('ไม่พบใบสมัคร');
  if(String(app.status)!==APP.STATUS.SELECTED) throw new Error('ยังไม่อยู่ในสถานะที่สามารถยืนยันสิทธิ์ได้');
  const round=findOne_(APP.SHEETS.ROUNDS,'roundId',app.roundId);
  const now=new Date(), start=parseDate_(round.confirmStart), end=parseDate_(round.confirmEnd);
  if(start && now<start) throw new Error('ยังไม่ถึงช่วงเวลายืนยันสิทธิ์');
  if(end && now>new Date(end.getTime()+86399999)) throw new Error('หมดเวลายืนยันสิทธิ์');
  updateRowById_(APP.SHEETS.APPLICATIONS,'applicationId',applicationId,{status:APP.STATUS.CONFIRMED,updatedAt:now_(),updatedBy:'STUDENT'});
  logAudit_('STUDENT',s.id,'CONFIRM_RIGHT','APPLICATION',applicationId,'ยืนยันสิทธิ์');
  return {ok:true};
}

/* =========================
 * STAFF DASHBOARD
 * ========================= */

function getStaffDashboard(token) {
  requireStaff_(token,['ADMIN','OFFICER','FINANCE','ACADEMIC']);
  const apps=getAllRows_(APP.SHEETS.APPLICATIONS);
  const payments=getAllRows_(APP.SHEETS.PAYMENTS);
  const rounds=getAllRows_(APP.SHEETS.ROUNDS);
  const programs=getAllRows_(APP.SHEETS.PROGRAMS);
  const today=today_();
  const byProgram={};
  apps.forEach(function(a){
    const p=findOne_(APP.SHEETS.PROGRAMS,'programId',a.programId);
    const key=p?p.programName:'ไม่ระบุ';
    byProgram[key]=(byProgram[key]||0)+1;
  });
  return {
    totalApplicants:apps.length,
    todayApplicants:apps.filter(function(a){return String(a.submittedAt).indexOf(today)===0;}).length,
    paid:payments.filter(function(p){return String(p.status)==='VERIFIED';}).length,
    selected:apps.filter(function(a){return String(a.status)===APP.STATUS.SELECTED;}).length,
    waitlist:apps.filter(function(a){return String(a.status)===APP.STATUS.WAITLIST;}).length,
    eligible:apps.filter(function(a){return String(a.status)===APP.STATUS.ELIGIBLE;}).length,
    rounds:rounds.length,programs:programs.length,byProgram:byProgram
  };
}

/* =========================
 * ADMIN: ROUNDS / PROGRAMS
 * ========================= */

function listRounds(token) {
  requireStaff_(token,['ADMIN','OFFICER']);
  return getAllRows_(APP.SHEETS.ROUNDS).sort(function(a,b){return String(b.createdAt).localeCompare(String(a.createdAt));});
}

function saveRound(token,data) {
  const s=requireStaff_(token,['ADMIN']);
  if(!data.roundName || !data.academicYear || !data.openDate || !data.closeDate) throw new Error('กรุณากรอกข้อมูลรอบรับสมัครให้ครบ');
  const open=parseDate_(data.openDate), close=parseDate_(data.closeDate);
  if(!open || !close || open>close) throw new Error('ช่วงวันเปิดรับสมัครไม่ถูกต้อง');
  if(data.examDate && !parseDate_(data.examDate)) throw new Error('วันสอบไม่ถูกต้อง');
  if(data.resultDate && !parseDate_(data.resultDate)) throw new Error('วันประกาศผลไม่ถูกต้อง');
  const obj={
    roundId:data.roundId||uid_('ROUND'),academicYear:String(data.academicYear),
    roundName:String(data.roundName),roundType:String(data.roundType||'ทั่วไป'),
    description:String(data.description||''),openDate:String(data.openDate),closeDate:String(data.closeDate),
    examDate:String(data.examDate||''),resultDate:String(data.resultDate||''),
    confirmStart:String(data.confirmStart||''),confirmEnd:String(data.confirmEnd||''),
    fee:Number(data.fee||0),status:String(data.status||'DRAFT'),
    maxApplicants:Number(data.maxApplicants||0),createdAt:data.roundId?(findOne_(APP.SHEETS.ROUNDS,'roundId',data.roundId)||{}).createdAt||now_():now_(),
    updatedAt:now_(),createdBy:s.id
  };
  if(data.roundId) updateRowById_(APP.SHEETS.ROUNDS,'roundId',data.roundId,obj);
  else appendRow_(APP.SHEETS.ROUNDS,obj);
  logAudit_('STAFF',s.id,data.roundId?'UPDATE':'CREATE','ROUND',obj.roundId,obj.roundName);
  return {ok:true,round:obj};
}

function toggleRound(token,roundId,status) {
  const s=requireStaff_(token,['ADMIN']);
  const r=findOne_(APP.SHEETS.ROUNDS,'roundId',roundId);
  if(!r) throw new Error('ไม่พบรอบ');
  if(status==='OPEN') {
    // ไม่บังคับให้มีเพียงหนึ่งรอบ สามารถเปิดหลายรอบได้
    const open=parseDate_(r.openDate),close=parseDate_(r.closeDate);
    if(!open||!close) throw new Error('กรุณาตั้งวันเปิดและปิดก่อน');
  }
  updateRowById_(APP.SHEETS.ROUNDS,'roundId',roundId,{status:status,updatedAt:now_()});
  logAudit_('STAFF',s.id,'ROUND_STATUS','ROUND',roundId,status);
  return {ok:true};
}

function listPrograms(token,roundId) {
  requireStaff_(token,['ADMIN','OFFICER','ACADEMIC']);
  return getAllRows_(APP.SHEETS.PROGRAMS).filter(function(p){return !roundId || String(p.roundId)===String(roundId);});
}

function saveProgram(token,data) {
  const s=requireStaff_(token,['ADMIN']);
  const round=findOne_(APP.SHEETS.ROUNDS,'roundId',data.roundId);
  if(!round) throw new Error('ต้องเลือกรอบรับสมัคร');
  if(!data.programName || Number(data.capacity)<1) throw new Error('กรุณากรอกชื่อแผนและจำนวนรับ');
  const obj={
    programId:data.programId||uid_('PROG'),roundId:data.roundId,programName:String(data.programName),
    programCode:String(data.programCode||''),quotaName:String(data.quotaName||''),
    quotaType:String(data.quotaType||'ทั่วไป'),capacity:Number(data.capacity),
    minimumGpa:Number(data.minimumGpa||0),description:String(data.description||''),
    status:String(data.status||'ACTIVE'),
    createdAt:data.programId?(findOne_(APP.SHEETS.PROGRAMS,'programId',data.programId)||{}).createdAt||now_():now_(),
    updatedAt:now_()
  };
  if(data.programId) updateRowById_(APP.SHEETS.PROGRAMS,'programId',data.programId,obj);
  else appendRow_(APP.SHEETS.PROGRAMS,obj);
  logAudit_('STAFF',s.id,data.programId?'UPDATE':'CREATE','PROGRAM',obj.programId,obj.programName);
  return {ok:true,program:obj};
}

/* =========================
 * APPLICATION MANAGEMENT
 * ========================= */

function listApplications(token,filters) {
  requireStaff_(token,['ADMIN','OFFICER','ACADEMIC','FINANCE']);
  filters=filters||{};
  let rows=getAllRows_(APP.SHEETS.APPLICATIONS);
  if(filters.roundId) rows=rows.filter(function(a){return String(a.roundId)===String(filters.roundId);});
  if(filters.status) rows=rows.filter(function(a){return String(a.status)===String(filters.status);});
  if(filters.search) {
    const q=String(filters.search).toLowerCase();
    rows=rows.filter(function(a){
      return [a.applicationNo,a.firstName,a.lastName,a.citizenId,a.phone].join(' ').toLowerCase().indexOf(q)>=0;
    });
  }
  return rows.map(function(a){
    const s=findOne_(APP.SHEETS.STUDENTS,'studentId',a.studentId)||{};
    const p=findOne_(APP.SHEETS.PROGRAMS,'programId',a.programId)||{};
    const r=findOne_(APP.SHEETS.ROUNDS,'roundId',a.roundId)||{};
    return {
      application:a,
      student:{studentId:s.studentId,citizenId:s.citizenId,phone:s.phone,email:s.email},
      program:p,round:r
    };
  });
}

function getApplicationDetail(token,applicationId) {
  requireStaff_(token,['ADMIN','OFFICER','ACADEMIC','FINANCE']);
  const a=findOne_(APP.SHEETS.APPLICATIONS,'applicationId',applicationId);
  if(!a) throw new Error('ไม่พบใบสมัคร');
  return {
    application:a,
    student:findOne_(APP.SHEETS.STUDENTS,'studentId',a.studentId),
    round:findOne_(APP.SHEETS.ROUNDS,'roundId',a.roundId),
    program:findOne_(APP.SHEETS.PROGRAMS,'programId',a.programId),
    documents:getAllRows_(APP.SHEETS.DOCUMENTS).filter(function(d){return String(d.applicationId)===String(applicationId);}),
    payment:findOne_(APP.SHEETS.PAYMENTS,'applicationId',applicationId),
    exam:findOne_(APP.SHEETS.EXAMS,'applicationId',applicationId),
    scores:getAllRows_(APP.SHEETS.SCORES).filter(function(x){return String(x.applicationId)===String(applicationId);})
  };
}

function updateApplicationByStaff(token,applicationId,patch) {
  const s=requireStaff_(token,['ADMIN','OFFICER']);
  const allowed=HEADERS.Applications.filter(function(k){
    return ['applicationId','studentId','applicationNo','createdAt','createdBy','submittedAt'].indexOf(k)<0;
  });
  const clean={};
  Object.keys(patch||{}).forEach(function(k){if(allowed.indexOf(k)>=0) clean[k]=patch[k];});
  clean.updatedAt=now_();clean.updatedBy=s.id;
  updateRowById_(APP.SHEETS.APPLICATIONS,'applicationId',applicationId,clean);
  logAudit_('STAFF',s.id,'UPDATE','APPLICATION',applicationId,JSON.stringify(clean));
  return {ok:true};
}

function reviewDocument(token,documentId,status,note) {
  const s=requireStaff_(token,['ADMIN','OFFICER','ACADEMIC']);
  if(['APPROVED','REJECTED','PENDING'].indexOf(String(status))<0) throw new Error('สถานะเอกสารไม่ถูกต้อง');
  const d=findOne_(APP.SHEETS.DOCUMENTS,'documentId',documentId);
  if(!d) throw new Error('ไม่พบเอกสาร');
  updateRowById_(APP.SHEETS.DOCUMENTS,'documentId',documentId,{
    status:status,reviewNote:String(note||''),reviewedAt:now_(),reviewedBy:s.id
  });
  const docs=getAllRows_(APP.SHEETS.DOCUMENTS).filter(function(x){return String(x.applicationId)===String(d.applicationId);});
  if(docs.some(function(x){return String(x.status)==='REJECTED';})) {
    updateRowById_(APP.SHEETS.APPLICATIONS,'applicationId',d.applicationId,{status:APP.STATUS.DOCUMENT_FIX,updatedAt:now_()});
  }
  logAudit_('STAFF',s.id,'REVIEW_DOCUMENT','DOCUMENT',documentId,status+' '+String(note||''));
  return {ok:true};
}

/* =========================
 * PAYMENT VERIFICATION
 * ========================= */

function listPayments(token,filters) {
  requireStaff_(token,['ADMIN','FINANCE']);
  filters=filters||{};
  let rows=getAllRows_(APP.SHEETS.PAYMENTS);
  if(filters.status) rows=rows.filter(function(p){return String(p.status)===String(filters.status);});
  if(filters.roundId) rows=rows.filter(function(p){return String(p.roundId)===String(filters.roundId);});
  return rows.map(function(p){
    return Object.assign({},p,{application:findOne_(APP.SHEETS.APPLICATIONS,'applicationId',p.applicationId)||{}});
  });
}

function verifyPayment(token,paymentId,status,note) {
  const s=requireStaff_(token,['ADMIN','FINANCE']);
  if(['VERIFIED','REJECTED','PENDING'].indexOf(String(status))<0) throw new Error('สถานะการชำระเงินไม่ถูกต้อง');
  const p=findOne_(APP.SHEETS.PAYMENTS,'paymentId',paymentId);
  if(!p) throw new Error('ไม่พบรายการชำระเงิน');
  updateRowById_(APP.SHEETS.PAYMENTS,'paymentId',paymentId,{
    status:status,reviewNote:String(note||''),verifiedAt:now_(),verifiedBy:s.id
  });
  logAudit_('STAFF',s.id,'VERIFY_PAYMENT','PAYMENT',paymentId,status);
  return {ok:true};
}


function importBankPayments(token,data) {
  const s=requireStaff_(token,['ADMIN','FINANCE']);
  if(!data || !data.csv) throw new Error('ไม่พบไฟล์ CSV/TXT');
  const rows=Utilities.parseCsv(String(data.csv));
  if(rows.length<2) throw new Error('ไฟล์ไม่มีข้อมูล');
  const header=rows[0].map(function(x){return String(x).trim().toLowerCase();});
  const findIdx=function(names){for(const n of names){const i=header.indexOf(n);if(i>=0)return i;}return -1;};
  const iApp=findIdx(['applicationno','เลขที่สมัคร','application_no']);
  const iCitizen=findIdx(['citizenid','เลขบัตรประชาชน','เลขประจำตัวประชาชน']);
  const iRef=findIdx(['referenceno','reference','เลขอ้างอิง','เลขรายการ']);
  const iAmount=findIdx(['amount','ยอดเงิน','จำนวนเงิน']);
  const iDate=findIdx(['bankdate','date','วันที่']);
  const iTime=findIdx(['banktime','time','เวลา']);
  if(iApp<0 && iCitizen<0) throw new Error('ไฟล์ธนาคารต้องมีเลขที่สมัครหรือเลขบัตรประชาชน');
  let matched=0,updated=0,skipped=0;
  for(let r=1;r<rows.length;r++){
    if(!rows[r].some(Boolean)) continue;
    let app=null;
    if(iApp>=0 && rows[r][iApp]) app=findOne_(APP.SHEETS.APPLICATIONS,'applicationNo',rows[r][iApp]);
    if(!app && iCitizen>=0 && rows[r][iCitizen]){
      const st=findOne_(APP.SHEETS.STUDENTS,'citizenId',normalizeCitizenId_(rows[r][iCitizen]));
      if(st) app=getAllRows_(APP.SHEETS.APPLICATIONS).find(function(a){return String(a.studentId)===String(st.studentId)&&String(a.status)!=='CANCELLED';})||null;
    }
    if(!app){skipped++;continue;}
    matched++;
    let pay=findOne_(APP.SHEETS.PAYMENTS,'applicationId',app.applicationId);
    if(!pay){
      pay={
        paymentId:uid_('PAY'),applicationId:app.applicationId,studentId:app.studentId,roundId:app.roundId,
        amount:iAmount>=0?Number(rows[r][iAmount]||0):Number((findOne_(APP.SHEETS.ROUNDS,'roundId',app.roundId)||{}).fee||0),
        method:'นำเข้าจากธนาคาร',referenceNo:iRef>=0?String(rows[r][iRef]||''):'',
        slipDocumentId:'',bankDate:iDate>=0?String(rows[r][iDate]||''):'',
        bankTime:iTime>=0?String(rows[r][iTime]||''):'',
        status:'VERIFIED',reviewNote:'นำเข้าจากไฟล์ธนาคาร',paidAt:now_(),verifiedAt:now_(),verifiedBy:s.id,createdAt:now_()
      };
      appendRow_(APP.SHEETS.PAYMENTS,pay);updated++;
    } else {
      updateRowById_(APP.SHEETS.PAYMENTS,'paymentId',pay.paymentId,{
        amount:iAmount>=0?Number(rows[r][iAmount]||pay.amount):pay.amount,
        referenceNo:iRef>=0?String(rows[r][iRef]||pay.referenceNo):pay.referenceNo,
        bankDate:iDate>=0?String(rows[r][iDate]||pay.bankDate):pay.bankDate,
        bankTime:iTime>=0?String(rows[r][iTime]||pay.bankTime):pay.bankTime,
        status:'VERIFIED',verifiedAt:now_(),verifiedBy:s.id,reviewNote:'นำเข้าจากไฟล์ธนาคาร'
      });updated++;
    }
  }
  logAudit_('STAFF',s.id,'IMPORT_BANK','PAYMENT','IMPORT',`matched=${matched},updated=${updated},skipped=${skipped}`);
  return {ok:true,matched:matched,updated:updated,skipped:skipped};
}

/* =========================
 * EXAM MANAGEMENT
 * ========================= */

function assignSeats(token,options) {
  const s=requireStaff_(token,['ADMIN','ACADEMIC']);
  options=options||{};
  if(!options.roundId) throw new Error('ต้องเลือกรอบ');
  const apps=getAllRows_(APP.SHEETS.APPLICATIONS).filter(function(a){
    return String(a.roundId)===String(options.roundId) && String(a.status)===APP.STATUS.ELIGIBLE;
  });
  if(!apps.length) throw new Error('ยังไม่มีผู้มีสิทธิ์สอบในรอบนี้');

  const existing=getAllRows_(APP.SHEETS.EXAMS);
  const rooms=String(options.rooms||'').split(',').map(function(x){return x.trim();}).filter(Boolean);
  if(!rooms.length) throw new Error('กรุณาระบุห้องสอบอย่างน้อย 1 ห้อง');

  const seatsPerRoom=Number(options.seatsPerRoom||30);
  const examTime=String(options.examTime||'');
  let created=0;
  apps.forEach(function(a,index){
    const roomIndex=Math.floor(index/seatsPerRoom);
    const room=rooms[roomIndex % rooms.length];
    const seat=(index%seatsPerRoom)+1;
    const old=existing.find(function(e){return String(e.applicationId)===String(a.applicationId);});
    const obj={
      examId:old?old.examId:uid_('EXAM'),applicationId:a.applicationId,studentId:a.studentId,roundId:a.roundId,
      examDate:String(options.examDate||''),building:String(options.building||''),
      room:room,seatNo:seat,examTime:examTime,status:'ASSIGNED',
      createdAt:old?old.createdAt:now_(),updatedAt:now_()
    };
    if(old) updateRowById_(APP.SHEETS.EXAMS,'examId',old.examId,obj);
    else appendRow_(APP.SHEETS.EXAMS,obj);
    created++;
  });
  logAudit_('STAFF',s.id,'ASSIGN_SEATS','ROUND',options.roundId,'จัดเลขที่นั่ง '+created+' รายการ');
  return {ok:true,count:created};
}

function listExams(token,roundId) {
  requireStaff_(token,['ADMIN','ACADEMIC','OFFICER']);
  return getAllRows_(APP.SHEETS.EXAMS).filter(function(e){return !roundId||String(e.roundId)===String(roundId);})
    .map(function(e){
      const a=findOne_(APP.SHEETS.APPLICATIONS,'applicationId',e.applicationId)||{};
      return {exam:e,application:a};
    });
}

function importScores(token,data) {
  const s=requireStaff_(token,['ADMIN','ACADEMIC']);
  if(!data || !data.csv) throw new Error('ไม่พบไฟล์ CSV');
  const rows=Utilities.parseCsv(String(data.csv));
  if(rows.length<2) throw new Error('ไฟล์ไม่มีข้อมูล');
  const header=rows[0].map(function(x){return String(x).trim().toLowerCase();});
  const idx=function(names){
    for(const n of names){const i=header.indexOf(n);if(i>=0)return i;}
    return -1;
  };
  const iCitizen=idx(['citizenid','เลขบัตรประชาชน','เลขประจำตัวประชาชน']);
  const iSubject=idx(['subject','วิชา','รายวิชา']);
  const iScore=idx(['score','คะแนน']);
  const iMax=idx(['maxscore','คะแนนเต็ม']);
  if(iCitizen<0||iSubject<0||iScore<0) throw new Error('หัวตารางต้องมี เลขบัตรประชาชน, วิชา, คะแนน');
  const batch=uid_('BATCH');
  let count=0;
  for(let r=1;r<rows.length;r++){
    if(!rows[r].some(Boolean)) continue;
    const citizen=normalizeCitizenId_(rows[r][iCitizen]);
    const st=findOne_(APP.SHEETS.STUDENTS,'citizenId',citizen);
    if(!st) continue;
    const app=getAllRows_(APP.SHEETS.APPLICATIONS).find(function(a){return String(a.studentId)===String(st.studentId)&&String(a.status)===APP.STATUS.ELIGIBLE;});
    if(!app) continue;
    const exam=findOne_(APP.SHEETS.EXAMS,'applicationId',app.applicationId);
    if(!exam) continue;
    appendRow_(APP.SHEETS.SCORES,{
      scoreId:uid_('SCORE'),examId:exam.examId,applicationId:app.applicationId,studentId:st.studentId,
      roundId:app.roundId,subject:String(rows[r][iSubject]),score:Number(rows[r][iScore]||0),
      maxScore:iMax>=0?Number(rows[r][iMax]||0):0,importBatch:batch,createdAt:now_(),createdBy:s.id
    });
    count++;
  }
  logAudit_('STAFF',s.id,'IMPORT_SCORES','SCORE',batch,'นำเข้าคะแนน '+count);
  return {ok:true,count:count,batch:batch};
}

function publishResults(token,roundId,items) {
  const s=requireStaff_(token,['ADMIN','ACADEMIC']);
  if(!roundId) throw new Error('ต้องเลือกรอบ');
  if(!Array.isArray(items)||!items.length) throw new Error('ไม่มีรายการผลสอบ');
  items.forEach(function(x){
    const app=findOne_(APP.SHEETS.APPLICATIONS,'applicationId',x.applicationId);
    if(!app || String(app.roundId)!==String(roundId)) return;
    const status=['SELECTED','WAITLIST','NOT_SELECTED'].indexOf(String(x.status))>=0 ? String(x.status) : APP.STATUS.NOT_SELECTED;
    updateRowById_(APP.SHEETS.APPLICATIONS,'applicationId',app.applicationId,{status:status,updatedAt:now_(),updatedBy:s.id});
  });
  logAudit_('STAFF',s.id,'PUBLISH_RESULTS','ROUND',roundId,'เผยแพร่ผลสอบ');
  return {ok:true};
}

/* =========================
 * USER MANAGEMENT
 * ========================= */

function listUsers(token) {
  requireStaff_(token,['ADMIN']);
  return getAllRows_(APP.SHEETS.USERS).map(safeUser_);
}

function saveUser(token,data) {
  const s=requireStaff_(token,['ADMIN']);
  if(!data.username || !data.role) throw new Error('กรุณากรอกชื่อผู้ใช้และสิทธิ์');
  if(data.userId) {
    const old=findOne_(APP.SHEETS.USERS,'userId',data.userId);
    if(!old) throw new Error('ไม่พบผู้ใช้');
    const patch={
      username:String(data.username),displayName:String(data.displayName||''),
      role:String(data.role),permissions:JSON.stringify(data.permissions||[]),
      status:String(data.status||'ACTIVE'),updatedAt:now_()
    };
    if(data.password) patch.passwordHash=hashPassword_(data.password);
    updateRowById_(APP.SHEETS.USERS,'userId',data.userId,patch);
  } else {
    if(findOne_(APP.SHEETS.USERS,'username',data.username)) throw new Error('ชื่อผู้ใช้นี้มีแล้ว');
    if(!data.password) throw new Error('ผู้ใช้ใหม่ต้องกำหนดรหัสผ่าน');
    appendRow_(APP.SHEETS.USERS,{
      userId:uid_('USR'),username:String(data.username),passwordHash:hashPassword_(data.password),
      displayName:String(data.displayName||''),role:String(data.role||'OFFICER'),
      permissions:JSON.stringify(data.permissions||[]),status:String(data.status||'ACTIVE'),
      createdAt:now_(),updatedAt:now_(),lastLoginAt:''
    });
  }
  logAudit_('STAFF',s.id,'USER_SAVE','USER',data.userId||'NEW',data.username);
  return {ok:true};
}

/* =========================
 * SETTINGS
 * ========================= */

function getAdminSettings(token) {
  requireStaff_(token,['ADMIN']);
  const keys=HEADERS.Settings;
  const out={};
  getAllRows_(APP.SHEETS.SETTINGS).forEach(function(r){out[r.key]=r.value;});
  return out;
}

function saveSetting(token,key,value) {
  const s=requireStaff_(token,['ADMIN']);
  const allowed=['schoolName','homeTitle','homeSubtitle','homeNotice','pdpaText','contactPhone','contactEmail','address',
    'schoolLogoFileId','schoolLogoUrl','bannerFileId','bannerUrl','mapFileId','mapUrl','applicationFeeDefault'];
  if(allowed.indexOf(String(key))<0) throw new Error('ไม่อนุญาตให้แก้ไขรายการนี้');
  setSetting(String(key),String(value||''));
  logAudit_('STAFF',s.id,'SETTING','SETTING',key,String(value||''));
  return {ok:true};
}

function uploadSettingImage(token,data) {
  const s=requireStaff_(token,['ADMIN']);
  validateUpload_(data.file,5);
  if(['logo','banner','map'].indexOf(String(data.kind))<0) throw new Error('ประเภทไฟล์ไม่ถูกต้อง');
  const saved=saveBase64File_(data.file,'ผู้สมัคร','SYSTEM_'+String(data.kind)+'_');
  const idKey=String(data.kind)==='logo'?'schoolLogoFileId':(String(data.kind)==='banner'?'bannerFileId':'mapFileId');
  const urlKey=String(data.kind)==='logo'?'schoolLogoUrl':(String(data.kind)==='banner'?'bannerUrl':'mapUrl');
  setSetting(idKey,saved.id);
  setSetting(urlKey,saved.url);
  logAudit_('STAFF',s.id,'UPLOAD_SETTING_IMAGE','SETTING',data.kind,saved.id);
  return {ok:true,url:saved.url,fileId:saved.id};
}

/* =========================
 * REPORTS / EXPORT
 * ========================= */

function exportApplicationsCsv(token,filters) {
  requireStaff_(token,['ADMIN','OFFICER','ACADEMIC','FINANCE']);
  const rows=listApplications(token,filters||{});
  const headers=['เลขที่สมัคร','เลขบัตรประชาชน','ชื่อ-สกุล','รอบ','แผนการเรียน','สถานะ','โทรศัพท์','GPA'];
  const data=[headers].concat(rows.map(function(x){
    const a=x.application,s=x.student,p=x.program,r=x.round;
    return [a.applicationNo,s.citizenId,[a.prefix,a.firstName,a.lastName].join(' ').trim(),
      r.roundName,p.programName,a.status,a.phone,a.gpa];
  }));
  const csv=data.map(function(row){return row.map(csvEscape_).join(',');}).join('\r\n');
  return {ok:true,fileName:'applications_'+today_()+'.csv',data:'data:text/csv;charset=utf-8,'+encodeURIComponent('\uFEFF'+csv)};
}

function csvEscape_(v) {
  const s=String(v==null?'':v).replace(/"/g,'""');
  return '"' + s + '"';
}

function exportApplicationsXlsx(token,filters) {
  requireStaff_(token,['ADMIN','OFFICER','ACADEMIC','FINANCE']);
  const rows=listApplications(token,filters||{});
  const tmp=SpreadsheetApp.create('Export ผู้สมัคร '+now_());
  const sh=tmp.getSheets()[0];
  sh.setName('Applicants');
  const headers=['เลขที่สมัคร','เลขบัตรประชาชน','ชื่อ-สกุล','รอบ','แผนการเรียน','สถานะ','โทรศัพท์','GPA'];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  const data=rows.map(function(x){
    const a=x.application,s=x.student,p=x.program,r=x.round;
    return [a.applicationNo,s.citizenId,[a.prefix,a.firstName,a.lastName].join(' ').trim(),
      r.roundName,p.programName,a.status,a.phone,a.gpa];
  });
  if(data.length) sh.getRange(2,1,data.length,headers.length).setValues(data);
  sh.autoResizeColumns(1,headers.length);
  SpreadsheetApp.flush();
  const blob=DriveApp.getFileById(tmp.getId()).getBlob().getAs(MimeType.MICROSOFT_EXCEL);
  const dataUrl='data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,'+
    Utilities.base64Encode(blob.getBytes());
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return {ok:true,fileName:'applications_'+today_()+'.xlsx',data:dataUrl};
}

function exportPdf(token,type,filters) {
  requireStaff_(token,['ADMIN','OFFICER','ACADEMIC','FINANCE']);
  filters=filters||{};
  const rows=listApplications(token,filters);
  let title='รายงานข้อมูลผู้สมัคร';
  if(type==='eligible') title='รายชื่อผู้มีสิทธิ์สอบ';
  if(type==='selected') title='รายชื่อผู้ผ่านการคัดเลือก';
  if(type==='waitlist') title='รายชื่อตัวสำรอง';
  if(type==='confirmed') title='รายชื่อผู้ยืนยันสิทธิ์';
  let filtered=rows;
  if(type==='eligible') filtered=rows.filter(function(x){return x.application.status===APP.STATUS.ELIGIBLE;});
  if(type==='selected') filtered=rows.filter(function(x){return x.application.status===APP.STATUS.SELECTED;});
  if(type==='waitlist') filtered=rows.filter(function(x){return x.application.status===APP.STATUS.WAITLIST;});
  if(type==='confirmed') filtered=rows.filter(function(x){return x.application.status===APP.STATUS.CONFIRMED;});

  const doc=DocumentApp.create(title+' '+today_());
  const body=doc.getBody();
  body.clear();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(42).setMarginRight(42);
  const school=getSetting('schoolName')||'สถานศึกษา';
  let p=body.appendParagraph(school);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontFamily('TH Sarabun PSK').setFontSize(18);
  p=body.appendParagraph(title);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontFamily('TH Sarabun PSK').setFontSize(18);
  p=body.appendParagraph('วันที่จัดทำ '+formatThaiDate_(new Date()));
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontFamily('TH Sarabun PSK').setFontSize(14);
  body.appendParagraph('');
  const table=body.appendTable();
  const hdr=['ลำดับ','เลขที่สมัคร','ชื่อ-สกุล','แผนการเรียน','สถานะ'];
  const hr=table.appendTableRow();
  hdr.forEach(function(h){hr.appendTableCell(h);});
  filtered.forEach(function(x,i){
    const row=table.appendTableRow();
    const a=x.application,p=x.program;
    [String(i+1),a.applicationNo,[a.prefix,a.firstName,a.lastName].join(' ').trim(),
      p.programName,a.status].forEach(function(v){row.appendTableCell(String(v||''));});
  });
  for(let r=0;r<table.getNumRows();r++){
    for(let c=0;c<table.getRow(r).getNumCells();c++){
      table.getRow(r).getCell(c).editAsText().setFontFamily('TH Sarabun PSK').setFontSize(12);
    }
  }
  doc.saveAndClose();
  const file=DriveApp.getFileById(doc.getId());
  const pdf=file.getBlob().getAs(MimeType.PDF);
  const dataUrl='data:application/pdf;base64,'+Utilities.base64Encode(pdf.getBytes());
  file.setTrashed(true);
  return {ok:true,fileName:title+'_'+today_()+'.pdf',data:dataUrl,count:filtered.length};
}

function generateExamTicket(token,applicationId) {
  const s=requireStudent_(token);
  const a=findOne_(APP.SHEETS.APPLICATIONS,'applicationId',applicationId);
  if(!a||String(a.studentId)!==String(s.id)) throw new Error('ไม่พบใบสมัคร');
  if([APP.STATUS.ELIGIBLE,APP.STATUS.SELECTED,APP.STATUS.CONFIRMED].indexOf(String(a.status))<0) throw new Error('ยังไม่มีสิทธิ์พิมพ์บัตรสอบ');
  const e=findOne_(APP.SHEETS.EXAMS,'applicationId',applicationId);
  if(!e) throw new Error('ยังไม่ได้จัดเลขที่นั่งสอบ');
  const p=findOne_(APP.SHEETS.PROGRAMS,'programId',a.programId)||{};
  const r=findOne_(APP.SHEETS.ROUNDS,'roundId',a.roundId)||{};
  const doc=DocumentApp.create('บัตรประจำตัวสอบ '+a.applicationNo);
  const body=doc.getBody();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);
  let x=body.appendParagraph(getSetting('schoolName')||'สถานศึกษา');
  x.setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontFamily('TH Sarabun PSK').setFontSize(22);
  x=body.appendParagraph('บัตรประจำตัวสอบ');
  x.setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontFamily('TH Sarabun PSK').setFontSize(20);
  const data=[
    ['เลขที่สมัคร',a.applicationNo],['ชื่อ-สกุล',[a.prefix,a.firstName,a.lastName].join(' ').trim()],
    ['รอบรับสมัคร',r.roundName],['แผนการเรียน',p.programName],['วันสอบ',e.examDate],
    ['เวลา',e.examTime],['อาคาร',e.building],['ห้องสอบ',e.room],['เลขที่นั่ง',e.seatNo]
  ];
  const table=body.appendTable(data);
  for(let rr=0;rr<table.getNumRows();rr++){
    table.getRow(rr).getCell(0).editAsText().setBold(true);
    for(let cc=0;cc<table.getRow(rr).getNumCells();cc++){
      table.getRow(rr).getCell(cc).editAsText().setFontFamily('TH Sarabun PSK').setFontSize(15);
    }
  }
  body.appendParagraph('');
  const photoDoc=getAllRows_(APP.SHEETS.DOCUMENTS).filter(function(d){
    return String(d.applicationId)===String(applicationId) && String(d.documentType)===APP.DOC_TYPES.PHOTO && String(d.status)!=='REJECTED';
  }).sort(function(a,b){return String(b.uploadedAt).localeCompare(String(a.uploadedAt));})[0];
  if(photoDoc){
    try{
      const image=body.appendImage(DriveApp.getFileById(photoDoc.driveFileId).getBlob());
      image.setWidth(120).setHeight(150);
    }catch(e){}
  }
  const schoolAddress=getSetting('address')||'';
  body.appendParagraph('สถานที่สอบ: '+schoolAddress).editAsText().setFontFamily('TH Sarabun PSK').setFontSize(14);
  const mapId=getSetting('mapFileId');
  if(mapId){
    try{
      body.appendParagraph('แผนที่สถานที่สอบ').editAsText().setFontFamily('TH Sarabun PSK').setFontSize(14).setBold(true);
      const mapImage=body.appendImage(DriveApp.getFileById(mapId).getBlob());
      mapImage.setWidth(420);
    }catch(e){}
  }
  doc.saveAndClose();
  const file=DriveApp.getFileById(doc.getId());
  const pdf=file.getBlob().getAs(MimeType.PDF);
  const dataUrl='data:application/pdf;base64,'+Utilities.base64Encode(pdf.getBytes());
  file.setTrashed(true);
  return {ok:true,fileName:'ExamTicket_'+a.applicationNo+'.pdf',data:dataUrl};
}

function formatThaiDate_(d) {
  const months=['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return d.getDate()+' '+months[d.getMonth()]+' '+(d.getFullYear()+543);
}

/* =========================
 * AUDIT
 * ========================= */

function logAudit_(actorType,actorId,action,targetType,targetId,details) {
  try {
    appendRow_(APP.SHEETS.AUDIT,{
      logId:uid_('LOG'),timestamp:now_(),actorType:actorType,actorId:actorId,
      action:action,targetType:targetType,targetId:targetId,details:String(details||'')
    });
  } catch(e) {}
}

/* =========================
 * GENERAL ADMIN DATA
 * ========================= */

function getRoundsAndProgramsForAdmin(token) {
  requireStaff_(token,['ADMIN','OFFICER','ACADEMIC','FINANCE']);
  return {
    rounds:getAllRows_(APP.SHEETS.ROUNDS),
    programs:getAllRows_(APP.SHEETS.PROGRAMS)
  };
}

function getPublicApplicantForm(token) {
  const s=requireStudent_(token);
  return {
    student:safeStudent_(findOne_(APP.SHEETS.STUDENTS,'studentId',s.id)),
    rounds:getOpenRounds_().map(function(r){
      return {roundId:r.roundId,roundName:r.roundName,academicYear:r.academicYear,fee:r.fee,
        programs:getProgramsByRound_(r.roundId)};
    })
  };
}
