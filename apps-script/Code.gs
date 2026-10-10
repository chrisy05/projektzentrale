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
  if(!ss.getSheetByName('App_Ausrichtung')){const sh=ss.insertSheet('App_Ausrichtung');sh.getRange('A1:B1').setValues([['Feld','Wert (JSON)']]);const a={startYear:2027,vision:'',why:'',notWant:'',want:'',skills:'',people:'',words:'Zeit, Freiheit, Familie, Gesundheit, Freude',steps:[],book:false,sleep:56,buffer:7,actual:[0,0,0,0,0,0],target:[0,0,0,0,0,0],feel:[3,3,3,3,3,3]};sh.getRange(2,1,Object.keys(a).length,2).setValues(Object.entries(a).map(([k,v])=>[k,JSON.stringify(v)]));sh.setFrozenRows(1);}
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
function validate_(s){
 const keys=['projects','tasks','milestones','times','reviews','sources','history'];
 if(!s||s.schema!==1||!s.alignment||JSON.stringify(s).length>1500000||!keys.every(k=>Array.isArray(s[k])&&s[k].length<=5000))throw Error('Ungültiges Datenformat.');
 const a=s.alignment;
 if(!['actual','target','feel'].every(k=>Array.isArray(a[k])&&a[k].length===6&&a[k].every(Number.isFinite)))throw Error('Die sechs Lebensbereiche fehlen.');
 if(!Number.isFinite(a.sleep)||a.sleep<0||!Number.isFinite(a.buffer)||a.buffer<0||['actual','target'].some(k=>a[k].some(x=>x<0)||a[k].reduce((x,y)=>x+y,0)+a.sleep+a.buffer>168))throw Error('Ungültiges Zeitbudget: maximal 168 Stunden.');
 if(a.feel.some(x=>!Number.isInteger(x)||x<1||x>5)||!Array.isArray(a.steps)||a.steps.some(x=>!Number.isInteger(x)||x<0||x>9)||typeof a.book!=='boolean')throw Error('Ungültige Ausrichtung.');
 for(const k of ['words','vision','why','notWant','want','skills','people'])if(typeof a[k]!=='string')throw Error('Ausrichtung unvollständig.');
 if(a.startYear!==undefined&&(!Number.isInteger(a.startYear)||a.startYear<1900||a.startYear>2200))throw Error('Bitte ein gültiges Startjahr angeben.');
 for(const k of keys){const seen=new Set();for(const r of s[k]){if(!r||Array.isArray(r)||typeof r.id!=='string'||!r.id||seen.has(r.id)||/[^a-zA-Z0-9_-]/.test(r.id))throw Error('Ungültige oder doppelte Datensatz-ID.');seen.add(r.id);for(const v of Object.values(r))if(typeof v!=='string'&&typeof v!=='number')throw Error('Ungültiger Feldtyp.');}}
 for(const k of ['projects','tasks','milestones','sources'])for(const r of s[k])if(typeof r.title!=='string'||!r.title.trim())throw Error('Titel fehlt.');
 for(const p of s.projects)if(!['Aktiv','Wartend','Passiv','Abgeschlossen'].includes(p.status)||!Number.isFinite(p.weeklyHours)||p.weeklyHours<0||p.weeklyHours>168)throw Error('Ungültiges Projektbudget oder Status.');
 for(const t of s.tasks)if(!['Offen','Wartend','Erledigt'].includes(t.status)||!['Heute','Woche','Später'].includes(t.scope)||!['Hoch','Normal','Niedrig'].includes(t.priority)||typeof t.note!=='string')throw Error('Ungültiger Aufgabenstatus.');
 for(const m of s.milestones)if(!Number.isInteger(m.year)||m.year<1||m.year>5)throw Error('Ungültiges Planungsjahr.');
 for(const t of s.times)if(!Number.isFinite(t.hours)||t.hours<=0||t.hours>24||!/^\d{4}-\d{2}-\d{2}$/.test(t.date)||!['Bestätigt','Geschätzt'].includes(t.confidence))throw Error('Ungültiger Zeiteintrag.');
 return s;
}
function saveState(next,expectedRevision){assertOwner_();validate_(next);const lock=LockService.getScriptLock();lock.waitLock(15000);try{const ss=book_(),before=readState_();if(before.revision!==expectedRevision)throw new Error('Das Sheet wurde inzwischen geändert. Neu laden, dann Änderung erneut vornehmen.');
  // Write a recoverable backup before replacing rows. Originals outside App_* are untouched.
  const backup=ss.getSheetByName('App_Backups')||ss.insertSheet('App_Backups');if(backup.getLastRow()===0)backup.appendRow(['Zeitpunkt','Tabellenname','Datensatz (JSON)']);const backupRows=[];for(const k in TABLES_)before[k].forEach(r=>backupRows.push([new Date().toISOString(),k,JSON.stringify(r)]));backupRows.push([new Date().toISOString(),'alignment',JSON.stringify(before.alignment)]);if(backupRows.length)ensureGrid_(backup,backup.getLastRow()+backupRows.length,3);if(backupRows.length)backup.getRange(backup.getLastRow()+1,1,backupRows.length,3).setValues(backupRows);
  try{writeState_(ss,next);SpreadsheetApp.flush();return {revision:readState_().revision};}catch(err){writeState_(ss,before);SpreadsheetApp.flush();throw err;}
}finally{lock.releaseLock();}}
function writeState_(ss,s){for(const k in TABLES_){const [name,headers]=TABLES_[k],sh=ss.getSheetByName(name),oldCount=Math.max(0,sh.getLastRow()-1);ensureGrid_(sh,s[k].length+1,headers.length);if(s[k].length)sh.getRange(2,1,s[k].length,headers.length).setValues(s[k].map(r=>headers.map(h=>encodeCell_(r[h]))));if(oldCount>s[k].length)sh.getRange(s[k].length+2,1,oldCount-s[k].length,headers.length).clearContent();}
 const sh=ss.getSheetByName('App_Ausrichtung'),rows=Object.entries(s.alignment).map(([k,v])=>[k,JSON.stringify(v)]),oldCount=Math.max(0,sh.getLastRow()-1);sh.getRange(2,1,rows.length,2).setValues(rows);if(oldCount>rows.length)sh.getRange(rows.length+2,1,oldCount-rows.length,2).clearContent();}

function ensureGrid_(sheet,rows,cols){if(rows>sheet.getMaxRows())sheet.insertRowsAfter(sheet.getMaxRows(),rows-sheet.getMaxRows());if(cols>sheet.getMaxColumns())sheet.insertColumnsAfter(sheet.getMaxColumns(),cols-sheet.getMaxColumns());}
