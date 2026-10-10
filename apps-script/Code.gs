/** Bound to the owner's Google Sheet. Deploy: execute as accessing user, access only myself. */
const TABLES_ = {
  projects: ['App_Projekte', ['id','title','status','area','weeklyHours','goal']],
  tasks: ['App_Aufgaben', ['id','projectId','title','status','priority','due','scope','milestoneId','area','estimate','note']],
  milestones: ['App_Meilensteine', ['id','title','year','projectId','area','due','proof','why']],
  times: ['App_Zeiten', ['id','date','projectId','hours','confidence','note']],
  reviews: ['App_Rueckblicke', ['id','date','text']],
  sources: ['App_Quellen', ['id','title','status','note']],
  history: ['App_Verlauf', ['id','date','text']]
};
/** Run from the Apps Script editor once; not callable from the website. */
function setup_() {
  const email = Session.getEffectiveUser().getEmail();
  if (!email) throw new Error('In deinem Google-Konto anmelden.');
  const p=PropertiesService.getScriptProperties();
  if(p.getProperty('OWNER_EMAIL') && p.getProperty('OWNER_EMAIL')!==email) throw new Error('Eigentümer stimmt nicht überein.');
  p.setProperty('OWNER_EMAIL',email);
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  p.setProperty('SPREADSHEET_ID',ss.getId());
  for(const key in TABLES_){const [name,headers]=TABLES_[key];if(!ss.getSheetByName(name)){const sheet=ss.insertSheet(name);sheet.getRange(1,1,1,headers.length).setValues([headers]);sheet.setFrozenRows(1);sheet.getRange(1,1,1,headers.length).setBackground('#1b7469').setFontColor('#ffffff').setFontWeight('bold');sheet.autoResizeColumns(1,headers.length);}}
  if(!ss.getSheetByName('App_Ausrichtung')){const sh=ss.insertSheet('App_Ausrichtung');sh.getRange('A1:B1').setValues([['Feld','Wert (JSON)']]);const a={vision:'',why:'',notWant:'',want:'',skills:'',people:'',words:'Zeit, Freiheit, Familie, Gesundheit, Freude',steps:[],book:false,sleep:56,buffer:7,actual:[0,0,0,0,0,0],target:[0,0,0,0,0,0],feel:[3,3,3,3,3,3]};sh.getRange(2,1,Object.keys(a).length,2).setValues(Object.entries(a).map(([k,v])=>[k,JSON.stringify(v)]));sh.setFrozenRows(1);}
}
function assertOwner_(){const owner=PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL'),active=Session.getActiveUser().getEmail();if(!owner||!active||active!==owner)throw new Error('Nur der eingerichtete Eigentümer hat Zugriff.');}
function book_(){assertOwner_();const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');if(!id)throw new Error('setup_ zuerst im Editor ausführen.');return SpreadsheetApp.openById(id);}
function doGet(){assertOwner_();return HtmlService.createHtmlOutputFromFile('Index').setTitle('Projektzentrale').addMetaTag('viewport','width=device-width, initial-scale=1');}
function getState(){return readState_();}
function readState_(){const ss=book_(),s={schema:1,alignment:{}};for(const key in TABLES_){const [name,headers]=TABLES_[key],sh=ss.getSheetByName(name);if(!sh)throw new Error('Einrichtung unvollständig: '+name);const values=sh.getDataRange().getValues();s[key]=values.slice(1).filter(r=>r[0]).map(r=>Object.fromEntries(headers.map((h,i)=>[h,decodeCell_(r[i])])));}
  const values=ss.getSheetByName('App_Ausrichtung').getDataRange().getValues();values.slice(1).filter(r=>r[0]).forEach(r=>s.alignment[r[0]]=JSON.parse(r[1]));s.revision=revision_(s);return s;}
function revision_(state){const s=JSON.parse(JSON.stringify(state));delete s.revision;return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(s)));}
// Prefix spreadsheet formulas as literal text. Undo only our escaping on read.
function encodeCell_(v){if(v===null||v===undefined)return '';if(typeof v==='string'&&/^[=+\-@']/.test(v))return "'"+v;return v;}
function decodeCell_(v){if(v instanceof Date)return Utilities.formatDate(v,'America/Santo_Domingo','yyyy-MM-dd');if(typeof v==='string'&&/^'[=+\-@']/.test(v))return v.slice(1);return v;}
function validate_(s){if(!s||s.schema!==1||!s.alignment)throw new Error('Ungültiges Schema.');if(JSON.stringify(s).length>1500000)throw new Error('Daten zu groß.');for(const k in TABLES_){if(!Array.isArray(s[k])||s[k].length>5000)throw new Error('Ungültige Tabelle '+k);const seen=new Set();s[k].forEach(r=>{if(!r||typeof r.id!=='string'||!r.id||seen.has(r.id))throw new Error('Datensatz-ID fehlt oder ist doppelt.');seen.add(r.id);for(const key of TABLES_[k][1])if(r[key]!==undefined&&typeof r[key]!=='string'&&typeof r[key]!=='number')throw new Error('Ungültiger Feldtyp.');});}
  for(const k of ['actual','target','feel'])if(!Array.isArray(s.alignment[k])||s.alignment[k].length!==6||s.alignment[k].some(x=>!Number.isFinite(x)))throw new Error('Sechs Lebensbereiche erforderlich.');for(const k of ['actual','target'])if(s.alignment[k].some(x=>x<0)||s.alignment[k].reduce((a,b)=>a+b,0)+Number(s.alignment.sleep)+Number(s.alignment.buffer)>168)throw new Error('Zeitbudget überschritten.');if(!Number.isFinite(s.alignment.sleep)||!Number.isFinite(s.alignment.buffer)||s.alignment.sleep<0||s.alignment.buffer<0)throw new Error('Ungültiges Zeitbudget.');if(!Array.isArray(s.alignment.steps)||typeof s.alignment.words!=='string')throw new Error('Ausrichtung fehlt.');s.times.forEach(t=>{if(!Number.isFinite(t.hours)||t.hours<=0||t.hours>24)throw new Error('Ungültige Arbeitszeit.');});}
function saveState(next,expectedRevision){assertOwner_();validate_(next);const lock=LockService.getScriptLock();lock.waitLock(15000);try{const ss=book_(),before=readState_();if(before.revision!==expectedRevision)throw new Error('Das Sheet wurde inzwischen geändert. Neu laden, dann Änderung erneut vornehmen.');
  // Write a recoverable backup before replacing rows. Originals outside App_* are untouched.
  const backup=ss.getSheetByName('App_Backups')||ss.insertSheet('App_Backups');if(backup.getLastRow()===0)backup.appendRow(['Zeitpunkt','Tabellenname','Datensatz (JSON)']);const backupRows=[];for(const k in TABLES_)before[k].forEach(r=>backupRows.push([new Date().toISOString(),k,JSON.stringify(r)]));backupRows.push([new Date().toISOString(),'alignment',JSON.stringify(before.alignment)]);if(backupRows.length)backup.getRange(backup.getLastRow()+1,1,backupRows.length,3).setValues(backupRows);
  try{writeState_(ss,next);SpreadsheetApp.flush();return {revision:readState_().revision};}catch(err){writeState_(ss,before);SpreadsheetApp.flush();throw err;}
}finally{lock.releaseLock();}}
function writeState_(ss,s){for(const k in TABLES_){const [name,headers]=TABLES_[k],sh=ss.getSheetByName(name),oldCount=Math.max(0,sh.getLastRow()-1);if(s[k].length)sh.getRange(2,1,s[k].length,headers.length).setValues(s[k].map(r=>headers.map(h=>encodeCell_(r[h]))));if(oldCount>s[k].length)sh.getRange(s[k].length+2,1,oldCount-s[k].length,headers.length).clearContent();}
 const sh=ss.getSheetByName('App_Ausrichtung'),rows=Object.entries(s.alignment).map(([k,v])=>[k,JSON.stringify(v)]),oldCount=Math.max(0,sh.getLastRow()-1);sh.getRange(2,1,rows.length,2).setValues(rows);if(oldCount>rows.length)sh.getRange(rows.length+2,1,oldCount-rows.length,2).clearContent();}
