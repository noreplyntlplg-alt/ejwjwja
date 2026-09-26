/**
 * ระบบรับสมัครนักเรียน - Google Apps Script + Google Sheets + Google Drive
 * เวอร์ชัน: 1.0.0
 * หมายเหตุ: โค้ดนี้ตั้งใจแยกไฟล์เพื่อให้นำเข้า Apps Script ได้ง่าย
 */

const CONFIG = {
  APP_NAME: 'ระบบรับสมัครนักเรียน',
  ADMIN_USERNAME: 'sxaiq54',
  ADMIN_PASSWORD: 'Sxxnga2011.54',
  MAX_FILE_BYTES: 8 * 1024 * 1024,
  ALLOWED_IMAGE_MIME: ['image/jpeg', 'image/png', 'image/webp'],
  ALLOWED_DOC_MIME: ['application/pdf', 'image/jpeg', 'image/png', 'image/webp'],
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
  }
};

const HEADERS = {
  Settings: ['key','value','updatedAt'],
  AdmissionRounds: ['roundId','academicYear','roundName','description','startDate','endDate','examDate','resultDate','confirmStartDate','confirmEndDate','applicationFee','status','createdAt','updatedAt'],
  Programs: ['programId','roundId','programName','programType','quotaName','capacity','usedCapacity','minimumGpa','description','status','createdAt','updatedAt'],
  Students: ['studentId','citizenId','firstName','lastName','birthDate','gender','phone','email','address','subdistrict','district','province','postalCode','parentName','parentRelation','parentOccupation','parentPhone','previousSchool','previousProvince','gpa','createdAt','updatedAt','status'],
  Applications: ['applicationId','studentId','roundId','programId','applicationNo','photoFileId','pdpaConsent','pdpaConsentAt','applicationStatus','documentStatus','paymentStatus','examStatus','resultStatus','seatNo','examRoom','examBuilding','submittedAt','updatedAt','staffNote'],
  Documents: ['documentId','applicationId','documentType','fileName','fileId','mimeType','fileSize','verificationStatus','verificationNote','uploadedAt','verifiedAt','verifiedBy'],
  Payments: ['paymentId','applicationId','paymentMethod','amount','referenceNo','slipFileId','bankDate','bankTime','verificationStatus','verificationNote','paidAt','verifiedAt','verifiedBy'],
  Exams: ['examId','roundId','applicationId','seatNo','room','building','examDate','examTime','status','createdAt','updatedAt'],
  Scores: ['scoreId','roundId','applicationId','subject','score','importedAt','importedBy'],
  Users: ['userId','username','passwordHash','fullName','role','active','createdAt','updatedAt'],
  AuditLog: ['logId','actorType','actorId','action','targetType','targetId','detail','createdAt']
};

function doGet() {
  setupSystem_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSystem_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('MASTER_SPREADSHEET_ID');
  let ss;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.APP_NAME + ' - ฐานข้อมูล');
    props.setProperty('MASTER_SPREADSHEET_ID', ss.getId());
  }
  Object.keys(HEADERS).forEach(name => ensureSheet_(ss, name, HEADERS[name]));
  ensureDriveFolders_();
  ensureDefaultSettings_();
  ensureDefaultAdmin_();
  return ss.getId();
}

function getDb_() {
  setupSystem_();
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID'));
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const current = sh.getRange(1,1,1,headers.length).getValues()[0];
  const different = headers.some((h,i) => current[i] !== h);
  if (different) sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  return sh;
}

function ensureDriveFolders_() {
  const props = PropertiesService.getScriptProperties();
  let rootId = props.getProperty('ROOT_FOLDER_ID');
  let root;
  if (rootId) { try { root = DriveApp.getFolderById(rootId); } catch(e) {} }
  if (!root) {
    root = DriveApp.createFolder(CONFIG.APP_NAME + ' - เอกสาร');
    props.setProperty('ROOT_FOLDER_ID', root.getId());
  }
  ['รูปถ่าย','บัตรประชาชน','ทะเบียนบ้าน','ปพ1','สลิปชำระเงิน','รายงาน'].forEach(n => {
    const p = 'FOLDER_' + n;
    let id = props.getProperty(p);
    let folder;
    if (id) { try { folder = DriveApp.getFolderById(id); } catch(e) {} }
    if (!folder) {
      folder = root.createFolder(n);
      props.setProperty(p, folder.getId());
    }
  });
}

function ensureDefaultSettings_() {
  const sh = getDbRaw_().getSheetByName(CONFIG.SHEETS.SETTINGS);
  const existing = getAllObjects_(sh);
  const map = {};
  existing.forEach(r => map[r.key] = r.value);
  const defaults = {
    schoolName: 'กรุณาตั้งชื่อโรงเรียน',
    schoolLogoFileId: '',
    bannerFileId: '',
    homeNotice: 'กรุณาตั้งค่าประกาศจากเมนูผู้ดูแลระบบ',
    contactPhone: '',
    contactEmail: '',
    address: '',
    systemOpenMessage: 'ระบบรับสมัครออนไลน์',
    pdpaText: 'ข้าพเจ้ายินยอมให้สถานศึกษาจัดเก็บ ใช้ และประมวลผลข้อมูลส่วนบุคคลเพื่อวัตถุประสงค์ในการรับสมัคร คัดเลือก สอบ ประกาศผล และการดำเนินงานทางการศึกษาที่เกี่ยวข้อง ตามกฎหมายคุ้มครองข้อมูลส่วนบุคคล'
  };
  Object.keys(defaults).forEach(k => {
    if (map[k] === undefined) appendRow_(sh, [k, defaults[k], now_()]);
  });
}

function ensureDefaultAdmin_() {
  const sh = getDbRaw_().getSheetByName(CONFIG.SHEETS.USERS);
  const users = getAllObjects_(sh);
  const found = users.some(u => String(u.username).toLowerCase() === CONFIG.ADMIN_USERNAME.toLowerCase());
  if (!found) {
    appendRow_(sh, [Utilities.getUuid(), CONFIG.ADMIN_USERNAME, sha256_(CONFIG.ADMIN_PASSWORD), 'ผู้ดูแลระบบ', 'admin', true, now_(), now_()]);
  }
}

function getDbRaw_() {
  const id = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

function now_() { return new Date(); }
function iso_(v) { return v ? new Date(v).toISOString() : ''; }
function sha256_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}
function clean_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function safeDate_(v) { if (!v) return ''; const d = new Date(v); return isNaN(d.getTime()) ? '' : d; }
function json_(obj) { return JSON.stringify(obj); }

function getAllObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(r => r.some(x => x !== '')).map(r => {
    const o = {}; headers.forEach((h,i) => o[h] = r[i]); return o;
  });
}

function findBy_(sheetName, key, value) {
  return getAllObjects_(getDb_().getSheetByName(sheetName)).find(r => String(r[key]) === String(value)) || null;
}
function filterBy_(sheetName, key, value) {
  return getAllObjects_(getDb_().getSheetByName(sheetName)).filter(r => String(r[key]) === String(value));
}
function appendRow_(sheet, row) { sheet.appendRow(row); }

function updateObject_(sheetName, key, keyValue, patch) {
  const sh = getDb_().getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idx = headers.indexOf(key);
  if (idx < 0) throw new Error('ไม่พบคีย์ ' + key);
  for (let r=1; r<data.length; r++) {
    if (String(data[r][idx]) === String(keyValue)) {
      Object.keys(patch).forEach(k => {
        const c = headers.indexOf(k);
        if (c >= 0) data[r][c] = patch[k];
      });
      sh.getRange(r+1,1,1,headers.length).setValues([data[r]]);
      return true;
    }
  }
  return false;
}

function audit_(actorType, actorId, action, targetType, targetId, detail) {
  appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.AUDIT), [Utilities.getUuid(), actorType, actorId || '', action, targetType || '', targetId || '', detail || '', now_()]);
}

function getPublicData() {
  setupSystem_();
  const settings = getSettings_();
  const rounds = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.ROUNDS)).filter(r => String(r.status) === 'OPEN').map(normalizeDates_);
  const programs = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.PROGRAMS)).filter(p => String(p.status) === 'ACTIVE').map(normalizeDates_);
  programs.forEach(p => { p.remaining = Math.max(0, Number(p.capacity || 0) - Number(p.usedCapacity || 0)); });
  return {ok:true, settings, rounds, programs};
}

function getSettings_() {
  const rows = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.SETTINGS));
  const o = {}; rows.forEach(r => o[r.key] = r.value);
  if (o.schoolLogoFileId) o.schoolLogoData = fileDataUrl_(o.schoolLogoFileId);
  if (o.bannerFileId) o.bannerData = fileDataUrl_(o.bannerFileId);
  return o;
}
function saveSettings(data, session) {
  requireRole_(session, ['admin']);
  const sh = getDb_().getSheetByName(CONFIG.SHEETS.SETTINGS);
  const allowed = ['schoolName','schoolLogoFileId','bannerFileId','homeNotice','contactPhone','contactEmail','address','systemOpenMessage','pdpaText'];
  allowed.forEach(k => {
    if (data[k] === undefined) return;
    const row = findBy_('Settings','key',k);
    if (row) updateObject_('Settings','key',k,{value:data[k],updatedAt:now_()});
    else appendRow_(sh,[k,data[k],now_()]);
  });
  audit_('admin',session.userId,'UPDATE_SETTINGS','Settings','',JSON.stringify(data));
  return {ok:true,settings:getSettings_()};
}

function registerStudent(data) {
  setupSystem_();
  const round = findBy_('AdmissionRounds','roundId',data.roundId);
  if (!round || String(round.status) !== 'OPEN' || !withinDates_(round.startDate, round.endDate)) throw new Error('รอบรับสมัครนี้ยังไม่เปิดหรือปิดรับสมัครแล้ว');
  validateCitizenId_(data.citizenId);
  if (findBy_('Students','citizenId',data.citizenId)) throw new Error('หมายเลขบัตรประชาชนนี้มีบัญชีผู้สมัครแล้ว');
  const required = ['firstName','lastName','birthDate','phone','parentName','parentPhone','previousSchool'];
  required.forEach(k => { if (!clean_(data[k])) throw new Error('กรุณากรอกข้อมูล ' + k); });
  const studentId = 'ST' + Utilities.getUuid().replace(/-/g,'').substring(0,12).toUpperCase();
  appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.STUDENTS), [studentId,clean_(data.citizenId),clean_(data.firstName),clean_(data.lastName),safeDate_(data.birthDate),clean_(data.gender),clean_(data.phone),clean_(data.email),clean_(data.address),clean_(data.subdistrict),clean_(data.district),clean_(data.province),clean_(data.postalCode),clean_(data.parentName),clean_(data.parentRelation),clean_(data.parentOccupation),clean_(data.parentPhone),clean_(data.previousSchool),clean_(data.previousProvince),clean_(data.gpa),now_(),now_(),'ACTIVE']);
  audit_('student',studentId,'REGISTER','Student',studentId,'สร้างบัญชีผู้สมัคร');
  return {ok:true,studentId,message:'สมัครสมาชิกสำเร็จ กรุณาเข้าสู่ระบบด้วยหมายเลขบัตรประชาชน 13 หลัก'};
}

function loginStudent(citizenId) {
  validateCitizenId_(citizenId);
  const s = findBy_('Students','citizenId',citizenId);
  if (!s) throw new Error('ไม่พบบัญชีผู้สมัคร กรุณาสมัครสมาชิกก่อน');
  if (String(s.status) !== 'ACTIVE') throw new Error('บัญชีผู้สมัครนี้ถูกระงับ');
  const token = createSession_({type:'student',userId:s.studentId,citizenId:s.citizenId,expires:Date.now()+6*60*60*1000});
  audit_('student',s.studentId,'LOGIN','Student',s.studentId,'เข้าสู่ระบบ');
  return {ok:true,token,user:{type:'student',id:s.studentId,name:s.firstName+' '+s.lastName,citizenId:s.citizenId}};
}

function loginStaff(username,password) {
  const u = findBy_('Users','username',username);
  if (!u || String(u.active) !== 'true' && String(u.active) !== 'TRUE' && u.active !== true) throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  if (sha256_(password) !== String(u.passwordHash)) throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  const token = createSession_({type:u.role,userId:u.userId,username:u.username,expires:Date.now()+8*60*60*1000});
  audit_(u.role,u.userId,'LOGIN','User',u.userId,'เข้าสู่ระบบเจ้าหน้าที่');
  return {ok:true,token,user:{type:u.role,id:u.userId,name:u.fullName,username:u.username}};
}

function createSession_(obj) {
  const token = Utilities.getUuid() + '.' + Utilities.getUuid();
  CacheService.getScriptCache().put('SESSION_'+token, JSON.stringify(obj), 21600);
  return token;
}
function getSession_(token) {
  if (!token) throw new Error('กรุณาเข้าสู่ระบบ');
  const raw = CacheService.getScriptCache().get('SESSION_'+token);
  if (!raw) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const s = JSON.parse(raw);
  if (s.expires < Date.now()) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  return s;
}
function requireRole_(token, roles) {
  const s = getSession_(token);
  if (roles.indexOf(s.type) < 0) throw new Error('ไม่มีสิทธิ์ดำเนินการ');
  return s;
}

function getStudentPortal(token) {
  const s = requireRole_(token,['student']);
  const student = findBy_('Students','studentId',s.userId);
  const apps = filterBy_('Applications','studentId',s.userId).map(normalizeDates_);
  apps.forEach(a => {
    a.round = findBy_('AdmissionRounds','roundId',a.roundId);
    a.program = findBy_('Programs','programId',a.programId);
    a.documents = filterBy_('Documents','applicationId',a.applicationId).map(normalizeDates_);
    a.payment = filterBy_('Payments','applicationId',a.applicationId).map(normalizeDates_)[0] || null;
    a.exam = filterBy_('Exams','applicationId',a.applicationId).map(normalizeDates_)[0] || null;
  });
  return {ok:true,student:normalizeDates_(student),applications:apps,settings:getSettings_(),rounds:getPublicData().rounds,programs:getPublicData().programs};
}

function createApplication(token,data) {
  const s = requireRole_(token,['student']);
  const round = findBy_('AdmissionRounds','roundId',data.roundId);
  if (!round || String(round.status)!=='OPEN' || !withinDates_(round.startDate,round.endDate)) throw new Error('รอบรับสมัครไม่พร้อมใช้งาน');
  if (!data.pdpaConsent) throw new Error('กรุณายอมรับความยินยอม PDPA ก่อนส่งใบสมัคร');
  const program = findBy_('Programs','programId',data.programId);
  if (!program || String(program.status)!=='ACTIVE' || String(program.roundId)!==String(data.roundId)) throw new Error('ไม่พบแผนการเรียนที่เลือก');
  if (Number(program.usedCapacity||0) >= Number(program.capacity||0)) throw new Error('แผนการเรียนนี้เต็มแล้ว');
  const existing = filterBy_('Applications','studentId',s.userId).find(a => String(a.roundId)===String(data.roundId));
  if (existing) throw new Error('คุณมีใบสมัครในรอบนี้แล้ว');
  const applicationId = 'APP' + Utilities.getUuid().replace(/-/g,'').substring(0,12).toUpperCase();
  const appNo = makeApplicationNo_(round,applicationId);
  const status = 'รอตรวจสอบ';
  appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.APPLICATIONS), [applicationId,s.userId,data.roundId,data.programId,appNo,'',true,now_(),status,'รอตรวจสอบ','ยังไม่ชำระ','รอดำเนินการ','รอดำเนินการ','', '', '', '',now_(),now_(),'']);
  updateObject_('Programs','programId',data.programId,{usedCapacity:Number(program.usedCapacity||0)+1,updatedAt:now_()});
  updateObject_('Students','studentId',s.userId,{firstName:clean_(data.firstName),lastName:clean_(data.lastName),birthDate:safeDate_(data.birthDate),gender:clean_(data.gender),phone:clean_(data.phone),email:clean_(data.email),address:clean_(data.address),subdistrict:clean_(data.subdistrict),district:clean_(data.district),province:clean_(data.province),postalCode:clean_(data.postalCode),parentName:clean_(data.parentName),parentRelation:clean_(data.parentRelation),parentOccupation:clean_(data.parentOccupation),parentPhone:clean_(data.parentPhone),previousSchool:clean_(data.previousSchool),previousProvince:clean_(data.previousProvince),gpa:clean_(data.gpa),updatedAt:now_()});
  audit_('student',s.userId,'CREATE_APPLICATION','Application',applicationId,appNo);
  return {ok:true,applicationId,applicationNo:appNo};
}

function uploadDocument(token, payload) {
  const s = requireRole_(token,['student']);
  const app = findBy_('Applications','applicationId',payload.applicationId);
  if (!app || String(app.studentId)!==String(s.userId)) throw new Error('ไม่พบใบสมัครหรือไม่มีสิทธิ์');
  const type = clean_(payload.documentType);
  const folder = getDocumentFolder_(type);
  const file = saveBase64File_(payload.base64,payload.fileName,payload.mimeType,folder);
  const existing = filterBy_('Documents','applicationId',app.applicationId).find(d => String(d.documentType)===type);
  if (existing) {
    updateObject_('Documents','documentId',existing.documentId,{fileName:file.getName(),fileId:file.getId(),mimeType:file.getMimeType(),fileSize:file.getSize(),verificationStatus:'รอตรวจสอบ',verificationNote:'',uploadedAt:now_(),verifiedAt:'',verifiedBy:''});
  } else {
    appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.DOCUMENTS),[Utilities.getUuid(),app.applicationId,type,file.getName(),file.getId(),file.getMimeType(),file.getSize(),'รอตรวจสอบ','',now_(),'','']);
  }
  updateObject_('Applications','applicationId',app.applicationId,{documentStatus:'รอตรวจสอบ',applicationStatus:'รอตรวจสอบเอกสาร',updatedAt:now_()});
  audit_('student',s.userId,'UPLOAD_DOCUMENT','Document',app.applicationId,type);
  return {ok:true,message:'อัปโหลดเอกสารสำเร็จ'};
}

function uploadPaymentSlip(token,payload) {
  const s = requireRole_(token,['student']);
  const app = findBy_('Applications','applicationId',payload.applicationId);
  if (!app || String(app.studentId)!==String(s.userId)) throw new Error('ไม่พบใบสมัครหรือไม่มีสิทธิ์');
  const file = saveBase64File_(payload.base64,payload.fileName,payload.mimeType,getDocumentFolder_('สลิปชำระเงิน'));
  const existing = filterBy_('Payments','applicationId',app.applicationId)[0];
  const row = {paymentMethod:'โอนเงิน',amount:Number(payload.amount||0),referenceNo:clean_(payload.referenceNo),slipFileId:file.getId(),bankDate:safeDate_(payload.bankDate),bankTime:clean_(payload.bankTime),verificationStatus:'รอตรวจสอบ',verificationNote:'',paidAt:now_(),verifiedAt:'',verifiedBy:''};
  if (existing) updateObject_('Payments','paymentId',existing.paymentId,Object.assign(row,{}));
  else appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.PAYMENTS),[Utilities.getUuid(),app.applicationId,row.paymentMethod,row.amount,row.referenceNo,row.slipFileId,row.bankDate,row.bankTime,row.verificationStatus,row.verificationNote,row.paidAt,row.verifiedAt,row.verifiedBy]);
  updateObject_('Applications','applicationId',app.applicationId,{paymentStatus:'รอตรวจสอบ',updatedAt:now_()});
  return {ok:true,message:'แจ้งชำระเงินแล้ว กรุณารอเจ้าหน้าที่ตรวจสอบ'};
}

function getFileData(token,fileId) {
  const s = getSession_(token);
  if (s.type!=='student' && s.type!=='admin' && s.type!=='staff') throw new Error('ไม่มีสิทธิ์');
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  return {ok:true,name:file.getName(),mimeType:blob.getContentType(),data:Utilities.base64Encode(blob.getBytes())};
}

function getStaffDashboard(token) {
  requireRole_(token,['staff','admin']);
  const apps = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.APPLICATIONS));
  const rounds = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.ROUNDS));
  const programs = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.PROGRAMS));
  const payments = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.PAYMENTS));
  const verifiedPaid = payments.filter(p=>String(p.verificationStatus)==='อนุมัติ').reduce((a,p)=>a+Number(p.amount||0),0);
  return {ok:true,stats:{applications:apps.length,paid:verifiedPaid,passed:apps.filter(a=>String(a.resultStatus)==='ผ่านการคัดเลือก').length,reserve:apps.filter(a=>String(a.resultStatus)==='ตัวสำรอง').length},rounds:rounds.map(normalizeDates_),programs:programs.map(normalizeDates_)};
}

function listApplications(token, filters) {
  requireRole_(token,['staff','admin']);
  filters = filters || {};
  let apps = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.APPLICATIONS));
  if (filters.roundId) apps = apps.filter(a=>String(a.roundId)===String(filters.roundId));
  if (filters.status) apps = apps.filter(a=>String(a.applicationStatus)===String(filters.status));
  if (filters.resultStatus) apps = apps.filter(a=>String(a.resultStatus)===String(filters.resultStatus));
  const students = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.STUDENTS));
  const programs = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.PROGRAMS));
  const rounds = getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.ROUNDS));
  return apps.map(a=>{
    const s=students.find(x=>String(x.studentId)===String(a.studentId))||{};
    const p=programs.find(x=>String(x.programId)===String(a.programId))||{};
    const r=rounds.find(x=>String(x.roundId)===String(a.roundId))||{};
    return Object.assign(normalizeDates_(a),{student:normalizeDates_(s),program:p,round:r,documents:filterBy_('Documents','applicationId',a.applicationId).map(normalizeDates_),payment:filterBy_('Payments','applicationId',a.applicationId).map(normalizeDates_)[0]||null,exam:filterBy_('Exams','applicationId',a.applicationId).map(normalizeDates_)[0]||null});
  });
}

function verifyDocument(token, documentId, status, note) {
  const s = requireRole_(token,['staff','admin']);
  if (['อนุมัติ','ปฏิเสธ','รอตรวจสอบ'].indexOf(status)<0) throw new Error('สถานะเอกสารไม่ถูกต้อง');
  updateObject_('Documents','documentId',documentId,{verificationStatus:status,verificationNote:clean_(note),verifiedAt:now_(),verifiedBy:s.userId});
  const d=findBy_('Documents','documentId',documentId);
  if(d){
    const docs=filterBy_('Documents','applicationId',d.applicationId);
    const allApproved=docs.length>0 && docs.every(x=>String(x.verificationStatus)==='อนุมัติ');
    const anyRejected=docs.some(x=>String(x.verificationStatus)==='ปฏิเสธ');
    updateObject_('Applications','applicationId',d.applicationId,{documentStatus:anyRejected?'เอกสารไม่สมบูรณ์':allApproved?'เอกสารครบถ้วน':'รอตรวจสอบ',applicationStatus:anyRejected?'เอกสารไม่สมบูรณ์':allApproved?'มีสิทธิ์เข้าสู่ขั้นตอนต่อไป':'รอตรวจสอบเอกสาร',updatedAt:now_()});
  }
  audit_(s.type,s.userId,'VERIFY_DOCUMENT','Document',documentId,status+' '+note);
  return {ok:true};
}

function verifyPayment(token,paymentId,status,note) {
  const s=requireRole_(token,['staff','admin']);
  updateObject_('Payments','paymentId',paymentId,{verificationStatus:status,verificationNote:clean_(note),verifiedAt:now_(),verifiedBy:s.userId});
  const p=findBy_('Payments','paymentId',paymentId);
  if(p) updateObject_('Applications','applicationId',p.applicationId,{paymentStatus:status,updatedAt:now_()});
  return {ok:true};
}

function createRound(token,data) {
  requireRole_(token,['admin']);
  if (!clean_(data.academicYear) || !clean_(data.roundName) || !data.startDate || !data.endDate) throw new Error('กรุณากรอกข้อมูลรอบรับสมัครให้ครบ');
  const start=safeDate_(data.startDate),end=safeDate_(data.endDate);
  if (!start || !end || start>end) throw new Error('ช่วงวันรับสมัครไม่ถูกต้อง');
  const roundId='RND'+Utilities.getUuid().replace(/-/g,'').substring(0,10).toUpperCase();
  appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.ROUNDS),[roundId,clean_(data.academicYear),clean_(data.roundName),clean_(data.description),start,end,safeDate_(data.examDate),safeDate_(data.resultDate),safeDate_(data.confirmStartDate),safeDate_(data.confirmEndDate),Number(data.applicationFee||0),'OPEN',now_(),now_()]);
  audit_('admin',getSession_(token).userId,'CREATE_ROUND','Round',roundId,JSON.stringify(data));
  return {ok:true,roundId};
}

function updateRound(token,roundId,data) {
  requireRole_(token,['admin']);
  const patch={};
  ['academicYear','roundName','description','applicationFee','status'].forEach(k=>{if(data[k]!==undefined)patch[k]=data[k];});
  ['startDate','endDate','examDate','resultDate','confirmStartDate','confirmEndDate'].forEach(k=>{if(data[k]!==undefined)patch[k]=safeDate_(data[k]);});
  patch.updatedAt=now_();
  updateObject_('AdmissionRounds','roundId',roundId,patch);
  return {ok:true};
}

function createProgram(token,data) {
  requireRole_(token,['admin']);
  const round=findBy_('AdmissionRounds','roundId',data.roundId);
  if(!round)throw new Error('ไม่พบรอบรับสมัคร');
  if(!clean_(data.programName) || Number(data.capacity||0)<=0)throw new Error('กรุณาระบุชื่อแผนและจำนวนรับ');
  const id='PRG'+Utilities.getUuid().replace(/-/g,'').substring(0,10).toUpperCase();
  appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.PROGRAMS),[id,data.roundId,clean_(data.programName),clean_(data.programType),clean_(data.quotaName),Number(data.capacity),0,Number(data.minimumGpa||0),clean_(data.description),'ACTIVE',now_(),now_()]);
  return {ok:true,programId:id};
}

function updateProgram(token,programId,data) {
  requireRole_(token,['admin']);
  const patch={};
  ['programName','programType','quotaName','capacity','minimumGpa','description','status'].forEach(k=>{if(data[k]!==undefined)patch[k]=data[k];});
  patch.updatedAt=now_();
  updateObject_('Programs','programId',programId,patch);
  return {ok:true};
}

function createUser(token,data) {
  requireRole_(token,['admin']);
  if(!data.username||!data.password||!data.fullName||['staff','admin'].indexOf(data.role)<0)throw new Error('ข้อมูลผู้ใช้ไม่ครบ');
  if(findBy_('Users','username',data.username))throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้ว');
  const id='USR'+Utilities.getUuid().replace(/-/g,'').substring(0,10).toUpperCase();
  appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.USERS),[id,clean_(data.username),sha256_(data.password),clean_(data.fullName),data.role,true,now_(),now_()]);
  return {ok:true,userId:id};
}

function updateUser(token,userId,data) {
  requireRole_(token,['admin']);
  const patch={};
  if(data.fullName!==undefined)patch.fullName=clean_(data.fullName);
  if(data.role!==undefined)patch.role=data.role;
  if(data.active!==undefined)patch.active=Boolean(data.active);
  if(data.password)patch.passwordHash=sha256_(data.password);
  patch.updatedAt=now_();
  updateObject_('Users','userId',userId,patch);
  return {ok:true};
}

function listUsers(token) { requireRole_(token,['admin']); return getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.USERS)).map(u=>{delete u.passwordHash;return normalizeDates_(u);}); }
function listRounds(token) { requireRole_(token,['staff','admin']); return getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.ROUNDS)).map(normalizeDates_); }
function listPrograms(token) { requireRole_(token,['staff','admin']); return getAllObjects_(getDb_().getSheetByName(CONFIG.SHEETS.PROGRAMS)).map(normalizeDates_); }

function assignExamSeats(token,roundId) {
  const s=requireRole_(token,['staff','admin']);
  const apps=listApplications(token,{roundId}).filter(a=>String(a.documentStatus)==='เอกสารครบถ้วน' || String(a.applicationStatus).indexOf('มีสิทธิ์')>=0);
  if(!apps.length)throw new Error('ไม่พบผู้สมัครที่พร้อมจัดเลขที่นั่งสอบ');
  const rooms = buildRooms_(apps.length);
  const existing=filterBy_('Exams','roundId',roundId);
  apps.forEach((a,i)=>{
    const room=rooms[Math.floor(i/30)]; const seat=String((i%30)+1).padStart(3,'0');
    const ex=existing.find(x=>String(x.applicationId)===String(a.applicationId));
    const vals={seatNo:seat,room:room.room,building:room.building,examDate:findBy_('AdmissionRounds','roundId',roundId).examDate,examTime:'09:00',status:'กำหนดแล้ว',updatedAt:now_()};
    if(ex)updateObject_('Exams','examId',ex.examId,vals); else appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.EXAMS),[Utilities.getUuid(),roundId,a.applicationId,seat,room.room,room.building,findBy_('AdmissionRounds','roundId',roundId).examDate,'09:00','กำหนดแล้ว',now_(),now_()]);
    updateObject_('Applications','applicationId',a.applicationId,{seatNo:seat,examRoom:room.room,examBuilding:room.building,examStatus:'มีสิทธิ์สอบ',updatedAt:now_()});
  });
  audit_(s.type,s.userId,'ASSIGN_EXAM_SEATS','Round',roundId,'จำนวน '+apps.length);
  return {ok:true,count:apps.length};
}
function buildRooms_(count){const rooms=[];for(let i=0;i<Math.ceil(count/30);i++)rooms.push({room:'ห้อง '+(i+1),building:'อาคารสอบ'});return rooms;}

function importScores(token,roundId,csvText) {
  const s=requireRole_(token,['staff','admin']);
  if(!csvText)throw new Error('ไม่พบข้อมูล CSV');
  const rows=Utilities.parseCsv(csvText);
  if(rows.length<2)throw new Error('ไฟล์ไม่มีข้อมูล');
  const headers=rows[0].map(x=>clean_(x).toLowerCase());
  const appIdx=headers.indexOf('applicationid')>=0?headers.indexOf('applicationid'):headers.indexOf('เลขที่ใบสมัคร');
  const subjectIdx=headers.indexOf('subject')>=0?headers.indexOf('subject'):headers.indexOf('วิชา');
  const scoreIdx=headers.indexOf('score')>=0?headers.indexOf('score'):headers.indexOf('คะแนน');
  if(appIdx<0||subjectIdx<0||scoreIdx<0)throw new Error('CSV ต้องมีคอลัมน์ applicationId, subject, score หรือชื่อภาษาไทยที่กำหนด');
  rows.slice(1).forEach(r=>{const app=findBy_('Applications','applicationId',r[appIdx]);if(!app||String(app.roundId)!==String(roundId))return;appendRow_(getDb_().getSheetByName(CONFIG.SHEETS.SCORES),[Utilities.getUuid(),roundId,r[appIdx],r[subjectIdx],Number(r[scoreIdx]),now_(),s.userId]);});
  return {ok:true};
}

function publishResults(token,roundId,items) {
  const s=requireRole_(token,['staff','admin']);
  if(!Array.isArray(items)||!items.length)throw new Error('ไม่พบรายการผลการคัดเลือก');
  items.forEach(x=>{
    const app=findBy_('Applications','applicationId',x.applicationId);if(!app||String(app.roundId)!==String(roundId))return;
    updateObject_('Applications','applicationId',x.applicationId,{resultStatus:x.resultStatus,applicationStatus:x.resultStatus==='ผ่านการคัดเลือก'?'สอบผ่าน':x.resultStatus,updatedAt:now_()});
  });
  audit_(s.type,s.userId,'PUBLISH_RESULTS','Round',roundId,'เผยแพร่ผล '+items.length+' รายการ');
  return {ok:true};
}

function exportApplicationsCsv(token,filters) {
  requireRole_(token,['staff','admin']);
  const rows=listApplications(token,filters||{});
  const out=[['เลขที่ใบสมัคร','รอบ','เลขบัตรประชาชน','ชื่อ-สกุล','แผนการเรียน','สถานะ','เอกสาร','ชำระเงิน','ผลคัดเลือก','เลขที่นั่งสอบ','ห้องสอบ']];
  rows.forEach(a=>out.push([a.applicationNo,a.round.roundName,a.student.citizenId,a.student.firstName+' '+a.student.lastName,a.program.programName,a.applicationStatus,a.documentStatus,a.paymentStatus,a.resultStatus,a.seatNo,a.examRoom]));
  return makeCsvResponse_(out,'รายชื่อผู้สมัคร.csv');
}

function exportPdf(token,type,filters) {
  requireRole_(token,['staff','admin']);
  const rows=listApplications(token,filters||{});
  let title='รายงานผู้สมัคร';
  if(type==='exam')title='รายชื่อผู้มีสิทธิ์สอบ';
  if(type==='passed')title='รายชื่อผู้ผ่านการคัดเลือก';
  if(type==='reserve')title='รายชื่อผู้สมัครตัวสำรอง';
  const filtered=type==='exam'?rows.filter(a=>String(a.examStatus)==='มีสิทธิ์สอบ'):type==='passed'?rows.filter(a=>String(a.resultStatus)==='ผ่านการคัดเลือก'):type==='reserve'?rows.filter(a=>String(a.resultStatus)==='ตัวสำรอง'):rows;
  const settings=getSettings_();
  const html='<html><head><meta charset="UTF-8"><style>@page{size:A4;margin:18mm}body{font-family:TH Sarabun PSK,Arial,sans-serif;font-size:16pt}h1{text-align:center;font-size:24pt;margin:0 0 4mm}p{text-align:center}.meta{margin-bottom:4mm}table{width:100%;border-collapse:collapse;font-size:13pt}th,td{border:1px solid #333;padding:4px 6px}th{background:#eee;text-align:center}.footer{margin-top:8mm;text-align:right}</style></head><body><h1>'+esc_(settings.schoolName||CONFIG.APP_NAME)+'</h1><h1>'+esc_(title)+'</h1><p>จัดทำเมื่อ '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm')+'</p><table><thead><tr><th>ลำดับ</th><th>เลขที่ใบสมัคร</th><th>ชื่อ-สกุล</th><th>แผนการเรียน</th><th>สถานะ/ผล</th><th>เลขที่นั่ง</th><th>ห้อง</th></tr></thead><tbody>'+filtered.map((a,i)=>'<tr><td style="text-align:center">'+(i+1)+'</td><td>'+esc_(a.applicationNo)+'</td><td>'+esc_((a.student.firstName||'')+' '+(a.student.lastName||''))+'</td><td>'+esc_(a.program.programName||'')+'</td><td>'+esc_(a.resultStatus||a.applicationStatus||'')+'</td><td style="text-align:center">'+esc_(a.seatNo||'')+'</td><td>'+esc_(a.examRoom||'')+'</td></tr>').join('')+'</tbody></table><div class="footer">ระบบรับสมัครนักเรียน</div></body></html>';
  const blob=Utilities.newBlob(html,'text/html','report.html').getAs(MimeType.PDF).setName(title+'_'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss')+'.pdf');
  return {ok:true,name:blob.getName(),mimeType:MimeType.PDF,data:Utilities.base64Encode(blob.getBytes())};
}


function uploadSystemFile(token,payload,fileKind) {
  requireRole_(token,['admin']);
  if (['logo','banner'].indexOf(fileKind)<0) throw new Error('ประเภทไฟล์ระบบไม่ถูกต้อง');
  if (!CONFIG.ALLOWED_IMAGE_MIME.includes(payload.mimeType)) throw new Error('โลโก้/แบนเนอร์ต้องเป็น JPG, PNG หรือ WEBP');
  const props=PropertiesService.getScriptProperties();
  const root=DriveApp.getFolderById(props.getProperty('ROOT_FOLDER_ID'));
  const folder=root.createFolder(fileKind==='logo'?'ตั้งค่าโลโก้':'ตั้งค่าแบนเนอร์');
  const file=saveBase64File_(payload.base64,payload.fileName,payload.mimeType,folder);
  const key=fileKind==='logo'?'schoolLogoFileId':'bannerFileId';
  updateObject_('Settings','key',key,{value:file.getId(),updatedAt:now_()});
  return {ok:true,fileId:file.getId()};
}

function confirmAdmission(token,applicationId) {
  const s=requireRole_(token,['student']);
  const app=findBy_('Applications','applicationId',applicationId);
  if(!app || String(app.studentId)!==String(s.userId)) throw new Error('ไม่พบใบสมัครหรือไม่มีสิทธิ์');
  if(String(app.resultStatus)!=='ผ่านการคัดเลือก') throw new Error('ใบสมัครนี้ยังไม่อยู่ในสถานะที่ยืนยันสิทธิ์ได้');
  const round=findBy_('AdmissionRounds','roundId',app.roundId);
  if(round && (round.confirmStartDate || round.confirmEndDate) && !withinDates_(round.confirmStartDate,round.confirmEndDate)) throw new Error('ขณะนี้ยังไม่อยู่ในช่วงเวลายืนยันสิทธิ์');
  updateObject_('Applications','applicationId',applicationId,{resultStatus:'ยืนยันสิทธิ์',applicationStatus:'ยืนยันสิทธิ์',updatedAt:now_()});
  audit_('student',s.userId,'CONFIRM_ADMISSION','Application',applicationId,'ยืนยันสิทธิ์');
  return {ok:true};
}

function exportApplicationsXlsx(token,filters) {
  requireRole_(token,['staff','admin']);
  const rows=listApplications(token,filters||{});
  const temp=SpreadsheetApp.create('รายงานผู้สมัคร_'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss'));
  const sh=temp.getSheets()[0];
  sh.setName('Applicants');
  const out=[['เลขที่ใบสมัคร','รอบ','เลขบัตรประชาชน','ชื่อ-สกุล','แผนการเรียน','สถานะ','เอกสาร','ชำระเงิน','ผลคัดเลือก','เลขที่นั่งสอบ','ห้องสอบ']];
  rows.forEach(a=>out.push([a.applicationNo,a.round?.roundName||'',a.student?.citizenId||'',(a.student?.firstName||'')+' '+(a.student?.lastName||''),a.program?.programName||'',a.applicationStatus,a.documentStatus,a.paymentStatus,a.resultStatus,a.seatNo,a.examRoom]));
  sh.getRange(1,1,out.length,out[0].length).setValues(out);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,out[0].length).setFontWeight('bold');
  SpreadsheetApp.flush();
  const url='https://docs.google.com/spreadsheets/d/'+temp.getId()+'/export?format=xlsx';
  const blob=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:false}).getBlob().setName('รายงานผู้สมัคร.xlsx');
  DriveApp.getFileById(temp.getId()).setTrashed(true);
  return {ok:true,name:blob.getName(),mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',data:Utilities.base64Encode(blob.getBytes())};
}

function makeCsvResponse_(rows,name){
  const csv=rows.map(r=>r.map(v=>'"'+String(v===undefined?'':v).replace(/"/g,'""')+'"').join(',')).join('\n');
  return {ok:true,name,mimeType:'text/csv;charset=utf-8',data:Utilities.base64Encode(Utilities.newBlob('\ufeff'+csv,'text/csv').getBytes())};
}

function saveBase64File_(base64,name,mimeType,folder){
  if(!base64||!name||!mimeType)throw new Error('ข้อมูลไฟล์ไม่ครบ');
  if(CONFIG.ALLOWED_DOC_MIME.indexOf(mimeType)<0)throw new Error('ชนิดไฟล์นี้ไม่รองรับ');
  const bytes=Utilities.base64Decode(String(base64).replace(/^data:[^;]+;base64,/,'') );
  if(bytes.length>CONFIG.MAX_FILE_BYTES)throw new Error('ไฟล์มีขนาดเกิน 8 MB');
  return folder.createFile(Utilities.newBlob(bytes,mimeType,name));
}
function getDocumentFolder_(type){
  const props=PropertiesService.getScriptProperties();
  const map={'รูปถ่าย':'FOLDER_รูปถ่าย','บัตรประชาชน':'FOLDER_บัตรประชาชน','ทะเบียนบ้าน':'FOLDER_ทะเบียนบ้าน','ปพ.1':'FOLDER_ปพ1','สลิปชำระเงิน':'FOLDER_สลิปชำระเงิน'};
  const id=props.getProperty(map[type]||'FOLDER_ปพ1');
  return DriveApp.getFolderById(id);
}
function fileDataUrl_(id){try{const b=DriveApp.getFileById(id).getBlob();return 'data:'+b.getContentType()+';base64,'+Utilities.base64Encode(b.getBytes());}catch(e){return '';}}
function validateCitizenId_(id){id=clean_(id).replace(/\D/g,'');if(!/^\d{13}$/.test(id))throw new Error('หมายเลขบัตรประชาชนต้องมี 13 หลัก');let sum=0;for(let i=0;i<12;i++)sum+=Number(id.charAt(i))*(13-i);const check=(11-(sum%11))%10;if(check!==Number(id.charAt(12)))throw new Error('หมายเลขบัตรประชาชนไม่ถูกต้อง');return true;}
function withinDates_(start,end){const n=new Date();return (!start||n>=new Date(start))&&(!end||n<=new Date(end));}
function makeApplicationNo_(round,id){const y=String(round.academicYear||new Date().getFullYear()).replace(/\D/g,'').slice(-4);return 'A'+y+'-'+id.slice(-6);}
function normalizeDates_(o){const x=Object.assign({},o);Object.keys(x).forEach(k=>{if(Object.prototype.toString.call(x[k])==='[object Date]')x[k]=iso_(x[k]);});return x;}
function esc_(s){return String(s===undefined?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
