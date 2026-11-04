/**
 * TDT DATA — APPS SCRIPT BACKEND
 * Separate spreadsheet. No PHO/BMT code or data is referenced.
 * Source headings are mapped by EXACT COLUMN POSITION (including duplicate Images).
 * Writes require an authorized editor, the correct 39-column sheet, and lock protection.
 */
const TDT_CONFIG = Object.freeze({
  spreadsheetId: PropertiesService.getScriptProperties().getProperty('TDT_SPREADSHEET_ID') || '',
  sheetGid: 1026053180,
  headerRow: 1,       // Change only if actual headings are on a different row.
  allowWrites: true,  // Saving works for approved editors ONLY after all 39 sheet headers pass validation.
  viewerEmails: [],   // Add approved signed-in Google account emails (or keep empty if only editors).
  editorEmails: [],   // REQUIRED: Add your own signed-in email before saving will work.
  release: '2.1',
  maxFerritinResults: 100, // Each extra value/date pair is appended AFTER the original 39 columns.
  preventDuplicateCr: true, // Prevent accidental duplicate patient records by exact CR number.
  timeZone: 'Asia/Kolkata',
  pageSize: 35,
  imageMaxBytes: 2 * 1024 * 1024,
  imageFolderId: '' // OPTIONAL private Drive folder ID for image uploads. Keep blank to disable uploads.
});

// Heading text and order are preserved, including original Ferriitn -4 spelling.
const TDT_HEADERS = Object.freeze([
  'Date','Name','CR . No','AGE','SEX','Number','HEIGHT','WEIGHT',
  'Comment','Pre T Hb','Relevant','Hemotrack Update',
  'Ferritin 1','Date.','Ferritin 2','Date 2','Ferritin-3','Date -3',
  'Ferriitn -4','Date -4','Ferritin -5','Date -5','Ferritin -6','DATE-6',
  'Ferritin-7','Date-7','Ferritin-8','Date-8','Ferritin-9','Date-9',
  'Tanner Staging','IDEAL HT','IDEAL WT','Delay','HT %ile','WT %ile',
  'Image','Images','Images'
]);


/** Source: IAP 2015 Tables II–V, published in Indian Pediatrics 52:47–55.
 * 50th centile only; age 5–18 years in 0.5-year increments.
 * Values between ages are linearly interpolated for display/reference, not clinical diagnosis.
 * https://www.indianpediatrics.net/jan2015/jan-47-55.htm
 */
const TDT_IAP50 = Object.freeze({
  M: {
    height:[108.9,111.9,114.8,117.8,120.7,123.5,126.4,129.1,131.8,134.5,137.2,139.9,142.7,145.5,148.4,151.4,154.3,157.2,159.9,162.3,164.5,166.5,168.1,169.6,171.0,172.3,173.6],
    weight:[17.1,18.2,19.3,20.7,21.9,23.3,24.8,26.4,27.9,29.4,31.1,32.8,34.7,36.9,39.0,41.2,43.3,45.7,48.2,50.8,53.1,55.2,56.8,58.2,59.5,60.6,61.6]
  },
  F: {
    height:[107.5,110.5,113.5,116.5,119.4,122.4,125.4,128.4,131.4,134.4,137.4,140.4,143.3,145.9,148.4,150.5,152.2,153.6,154.7,155.5,156.1,156.6,156.9,157.2,157.4,157.6,157.8],
    weight:[16.4,17.6,18.7,19.9,21.2,22.5,24.0,25.5,27.2,29.0,31.0,33.2,35.4,37.6,39.8,41.8,43.6,45.1,46.4,47.5,48.4,49.1,49.7,50.3,50.9,51.5,52.0]
  }
});
function tdtIap50_(ageValue, sexValue) {
  const age=Number(ageValue);
  const raw=String(sexValue||'').toUpperCase().trim();
  const sex=raw==='MALE'?'M':raw==='FEMALE'?'F':raw;
  if (!Number.isFinite(age)||String(ageValue??'').trim()===''||age<5||age>18||!TDT_IAP50[sex])return null;
  const lower=Math.floor((age-5)*2+1e-9);
  const upper=Math.min(26,lower+1);
  const fraction=Math.max(0,Math.min(1,(age-5)*2-lower));
  const series=TDT_IAP50[sex];
  const round=v=>Math.round(v*10)/10;
  return {height:round(series.height[lower]*(1-fraction)+series.height[upper]*fraction),
    weight:round(series.weight[lower]*(1-fraction)+series.weight[upper]*fraction)};
}
function tdtExtraCount_(sh) {
  const end=sh.getLastColumn();
  if (end<=39) return 0;
  const extra=end-39;
  if(extra%2!==0 || extra/2> TDT_CONFIG.maxFerritinResults-9)
    throw new Error('Unexpected columns after the original 39. Extra ferritin header mapping cannot be verified.');
  const headers=sh.getRange(TDT_CONFIG.headerRow,40,1,extra).getDisplayValues()[0];
  for(let i=0;i<extra/2;i++){
    const number=10+i;
    if(tdtNorm_(headers[2*i])!==tdtNorm_('Ferritin '+number)||
       tdtNorm_(headers[2*i+1])!==tdtNorm_('Date-'+number))
      throw new Error('Column '+(40+2*i)+' onwards contains unexpected data. Do not overwrite it; review extra ferritin columns.');
  }
  return extra/2;
}
function tdtExtendFerritin_(sh, desiredCount) {
  const existing=tdtExtraCount_(sh);
  if(desiredCount<=existing)return existing;
  if(!Number.isInteger(desiredCount)||desiredCount>TDT_CONFIG.maxFerritinResults-9)
    throw new Error('Too many ferritin results. The configured safety limit is '+TDT_CONFIG.maxFerritinResults+'.');
  const last=39+existing*2;
  const add=(desiredCount-existing)*2;
  if (sh.getMaxColumns()<last+add)sh.insertColumnsAfter(sh.getMaxColumns(),last+add-sh.getMaxColumns());
  const names=[];
  for(let n=existing+10;n<desiredCount+10;n++)names.push('Ferritin '+n,'Date-'+n);
  sh.getRange(TDT_CONFIG.headerRow,last+1,1,names.length).setValues([names]);
  sh.getRange(TDT_CONFIG.headerRow,last+1,1,names.length).setFontWeight('bold').setBackground('#e1f2e8');
  // Reset value-column formats; keep CR/mobiles as text, ferritin as numeric.
  for(let i=0;i<add;i+=2){
    sh.getRange(TDT_CONFIG.headerRow+1,last+i+1,Math.max(1,sh.getMaxRows()-TDT_CONFIG.headerRow),1).setNumberFormat('0.##');
    sh.getRange(TDT_CONFIG.headerRow+1,last+i+2,Math.max(1,sh.getMaxRows()-TDT_CONFIG.headerRow),1).setNumberFormat('yyyy-mm-dd');
  }
  return desiredCount;
}
function tdtValidISO_(value,allowFuture) {
  const s=String(value||'');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(s))return false;
  const m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const y=Number(m[1]),month=Number(m[2]),d=Number(m[3]);
  if(y<1900||y>2100||month<1||month>12||d<1||d>new Date(Date.UTC(y,month,0)).getUTCDate())return false;
  return allowFuture || s<=Utilities.formatDate(new Date(),TDT_CONFIG.timeZone,'yyyy-MM-dd');
}
function tdtExtraInput_(item) {
  const value=String(item?.value??'').trim(),date=String(item?.date??'').trim();
  if(value||date){
    if(!value||!date)throw new Error('Every new ferritin result must have both a value and a date.');
    const number=Number(value);
    if(!Number.isFinite(number)||number<0||number>100000)
      throw new Error('Ferritin value must be a valid nonnegative number.');
    if(!tdtValidISO_(date,false))throw new Error('Select a valid ferritin date (not in the future).');
  }
  return {value:value===''?'':Number(value),date};
}

function doGet() {
  const t = HtmlService.createTemplateFromFile('TDT_Data_UI');
  t.release = TDT_CONFIG.release;
  return t.evaluate().setTitle('TDT DATA | Patient Data Entry')
    .addMetaTag('viewport','width=device-width,initial-scale=1,viewport-fit=cover');
}

function tdtAuthorized_(editing) {
  const allow = editing ? TDT_CONFIG.editorEmails : TDT_CONFIG.viewerEmails.concat(TDT_CONFIG.editorEmails);
  if (!allow || !allow.length) throw new Error('TDT DATA access is not configured. Add your signed-in Google account to editorEmails (and optionally viewerEmails) in TDT_Data_Server.gs.');
  if (editing && !TDT_CONFIG.allowWrites) throw new Error('Saving is disabled. Verify the schema, then enable allowWrites in TDT_Data_Server.gs.');
  const identity = String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  if (!identity) throw new Error('Google did not provide a signed-in user email. Deploy for approved signed-in users; anonymous web-app access is not allowed.');
  if (!allow.some(email=>String(email).toLowerCase().trim()===identity)) throw new Error('This Google account is not authorized to '+(editing?'edit':'view')+' patient records.');
  return identity;
}

function tdtSheet_() {
  if (!TDT_CONFIG.spreadsheetId) throw new Error('Missing TDT_SPREADSHEET_ID Script Property. Set it in Apps Script > Project Settings before using the app.');
  const ss = SpreadsheetApp.openById(TDT_CONFIG.spreadsheetId);
  const sh = ss.getSheets().find(s=>s.getSheetId()===TDT_CONFIG.sheetGid);
  if (!sh) throw new Error('The configured tab gid '+TDT_CONFIG.sheetGid+' was not found. No records changed.');
  if (sh.getMaxColumns()<TDT_HEADERS.length) throw new Error('The source sheet has fewer than 39 columns. No records changed.');
  return sh;
}

function tdtNorm_(s) { return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,''); }
function tdtCheckSchema_(sh) {
  const actual = sh.getRange(TDT_CONFIG.headerRow,1,1,TDT_HEADERS.length).getDisplayValues()[0];
  const errors=[];
  TDT_HEADERS.forEach((expected,i)=>{
    const norm=tdtNorm_(actual[i]);
    const aliases=i===18 ? ['ferriitn4','ferritin4'] : [tdtNorm_(expected)];
    if (!aliases.includes(norm)) errors.push('Column '+(i+1)+': expected "'+expected+'", got "'+actual[i]+'"');
  });
  if(errors.length) throw new Error('Column mapping mismatch (saving blocked). '+errors.slice(0,5).join('; ')+(errors.length>5?'; plus '+(errors.length-5)+' more':'')+'. Verify headerRow and source headings.');
  return actual;
}

function tdtIso_(v) {
  if(v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v,TDT_CONFIG.timeZone,'yyyy-MM-dd');
  const s=String(v===null||v===undefined?'':v).trim();
  if(!s) return '';
  const iso=s.match(/^(\d{4})-(\d{2})-(\d{2})/); if(iso) return iso[0];
  const dmy=s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
  if(dmy) return dmy[3]+'-'+dmy[2].padStart(2,'0')+'-'+dmy[1].padStart(2,'0');
  const parsed=new Date(s);
  return isNaN(parsed.getTime()) ? '' : Utilities.formatDate(parsed,TDT_CONFIG.timeZone,'yyyy-MM-dd');
}
function tdtDisplay_(v,i) {
  if (v===null||v===undefined||v==='') return '';
  if ((i===0 || (i>=13&&i<=29&&i%2===1) || (i>=40&&i%2===0)) && !(v instanceof Date)) {
    // Normalize existing date strings such as 04/11/2026 for HTML date inputs.
    return tdtIso_(v) || String(v);
  }
  if (v instanceof Date) return Utilities.formatDate(v,TDT_CONFIG.timeZone,(i===0 || (i>=13&&i<=29&&i%2===1) || (i>=40&&i%2===0))?'yyyy-MM-dd':'yyyy-MM-dd HH:mm');
  if (typeof v==='boolean') return v?'YES':'NO';
  return String(v);
}
function tdtFingerprint_(row) {
  const normalized=row.map(x=>x instanceof Date?x.getTime():x);
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(normalized));
  return Utilities.base64EncodeWebSafe(bytes).slice(0,40);
}
function tdtToItem_(rowNumber,row) {
  const extras=[];
  for(let i=39;i<row.length;i+=2)extras.push({number:10+(i-39)/2,value:tdtDisplay_(row[i],i),date:tdtDisplay_(row[i+1],i+1)});
  return {rowNumber,fingerprint:tdtFingerprint_(row),values:row.slice(0,39).map((x,i)=>tdtDisplay_(x,i)),extraFerritin:extras};
}
function tdtRowsForDate_(sh,date) {
  const total=Math.max(0,sh.getLastRow()-TDT_CONFIG.headerRow);
  if (!total) return [];
  const raw=sh.getRange(TDT_CONFIG.headerRow+1,1,total,1).getValues();
  const matches=[];
  raw.forEach((r,i)=>{if (tdtIso_(r[0])===date)matches.push(TDT_CONFIG.headerRow+1+i);});
  return matches.reverse();
}
function tdtReadSparse_(sh,rowNumbers) {
  const sorted=rowNumbers.slice().sort((a,b)=>a-b), groups=[],width=39+2*tdtExtraCount_(sh);
  sorted.forEach(r=>{const last=groups[groups.length-1];if(last&&last.end+1===r&&last.end-last.start<59)last.end=r;else groups.push({start:r,end:r});});
  const found=new Map();
  groups.forEach(g=>sh.getRange(g.start,1,g.end-g.start+1,width).getValues().forEach((row,j)=>found.set(g.start+j,row)));
  return rowNumbers.map(n=>tdtToItem_(n,found.get(n)));
}

function tdtBootstrap() {
  const signedIn=tdtAuthorized_(false);
  const sh=tdtSheet_(),headers=tdtCheckSchema_(sh),extraFerritinCount=tdtExtraCount_(sh);
  const n=Math.max(0,sh.getLastRow()-TDT_CONFIG.headerRow);
  const dateValues=n?sh.getRange(TDT_CONFIG.headerRow+1,1,n,1).getValues():[];
  const counts={};dateValues.forEach(r=>{const d=tdtIso_(r[0]);if(d)counts[d]=(counts[d]||0)+1;});
  const dates=Object.keys(counts).sort().reverse();
  return {title:sh.getParent().getName(),tabName:sh.getName(),release:TDT_CONFIG.release,
    headers,dates:dates.slice(0,2000),counts,extraFerritinCount,
    iap50:TDT_IAP50,iapSource:'IAP 2015, Indian Pediatrics 52:47-55, Tables II–V; ages 5–18',
    writeEnabled:TDT_CONFIG.allowWrites && TDT_CONFIG.editorEmails.some(e=>String(e).toLowerCase().trim()===signedIn),
    imageUploadEnabled:TDT_CONFIG.allowWrites && !!TDT_CONFIG.imageFolderId && TDT_CONFIG.editorEmails.some(e=>String(e).toLowerCase().trim()===signedIn),
    pageSize:TDT_CONFIG.pageSize,today:Utilities.formatDate(new Date(),TDT_CONFIG.timeZone,'yyyy-MM-dd')};
}
function tdtListDate(date,page) {
  tdtAuthorized_(false);
  const sh=tdtSheet_(); tdtCheckSchema_(sh);
  const d=String(date||'');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('Select a valid date.');
  const extraCount=tdtExtraCount_(sh);
  const indices=tdtRowsForDate_(sh,d),total=indices.length, pages=Math.ceil(total/TDT_CONFIG.pageSize);
  const p=Math.max(0,Math.min(Math.max(0,Number(page)||0),Math.max(0,pages-1)));
  const subset=indices.slice(p*TDT_CONFIG.pageSize,(p+1)*TDT_CONFIG.pageSize);
  const items=tdtReadSparse_(sh,subset).map(item=>({
    rowNumber:item.rowNumber,fingerprint:item.fingerprint,
    name:item.values[1],cr:item.values[2],sex:item.values[4],age:item.values[3],
    ferritinCount:[12,14,16,18,20,22,24,26,28].filter(i=>item.values[i]!=='').length+item.extraFerritin.filter(x=>x.value!=='').length,
    ferritinLatest:(()=>{for(let i=28;i>=12;i-=2)if(item.values[i]!=='')return item.values[i];return '';})()
  }));
  return {date:d,total,page:p,pages,items};
}
function tdtFindCR(query) {
  tdtAuthorized_(false);
  const sh=tdtSheet_(); tdtCheckSchema_(sh);
  const q=String(query||'').trim();
  if(q.length<2||q.length>100)throw new Error('Enter at least 2 characters of CR number.');
  const n=Math.max(0,sh.getLastRow()-TDT_CONFIG.headerRow);
  const matches=[];
  if(n){
    sh.getRange(TDT_CONFIG.headerRow+1,3,n,1).getDisplayValues().forEach((r,i)=>{
      if(String(r[0]||'').toLowerCase().includes(q.toLowerCase())) matches.push(TDT_CONFIG.headerRow+1+i);
    });
  }
  const limited=matches.slice(-40).reverse();
  return {total:matches.length,items:tdtReadSparse_(sh,limited).map(x=>({rowNumber:x.rowNumber,name:x.values[1],cr:x.values[2],date:x.values[0]}))};
}
function tdtGetRecord(rowNumber) {
  tdtAuthorized_(false);
  const sh=tdtSheet_(); tdtCheckSchema_(sh);
  const n=Number(rowNumber);
  if(!Number.isInteger(n)||n<=TDT_CONFIG.headerRow||n>sh.getLastRow())throw new Error('Invalid record reference.');
  return tdtToItem_(n,sh.getRange(n,1,1,39+2*tdtExtraCount_(sh)).getValues()[0]);
}

function tdtCleanValue_(value,old,i) {
  const s=String(value===null||value===undefined?'':value).trim();
  if(s.length>5000)throw new Error('Value exceeds 5000 characters in '+(TDT_HEADERS[i]||'extra ferritin'));
  if(!s)return '';
  if(i===0 || (i>=13&&i<=29&&i%2===1)){
    if(!tdtValidISO_(s,i===0))throw new Error('Invalid date in '+TDT_HEADERS[i]);
    return old instanceof Date?new Date(s+'T12:00:00+05:30'):s;
  }
  if(i===1){
    const name=s.toUpperCase().replace(/ {2,}/g,' ');
    if(!/^[A-Z]+(?: [A-Z]+)*$/.test(name))throw new Error('Name must use A–Z capital letters and spaces only.');
    return name;
  }
  if(i===3){
    if(!/^\d+(?:\.\d{1,2})?$/.test(s)||Number(s)>120)throw new Error('Age must be a number of years (0–120).');
    return Number(s);
  }
  if(i===4){
    if(!['M','F','MALE','FEMALE'].includes(s.toUpperCase()))throw new Error('Sex must be M or F.');
    return s.toUpperCase().startsWith('M')?'M':'F';
  }
  if(i===5){
    if(!/^\d{10}$/.test(s))throw new Error('Mobile number must contain exactly 10 digits.');
    return s; // Preserve leading zero, if present.
  }
  if(i===10 || i===11){
    if(!['YES','NO'].includes(s.toUpperCase()))throw new Error(TDT_HEADERS[i]+' must be YES or NO.');
    return s.toUpperCase();
  }
  if(i===30){
    if(!/^[1-5]$/.test(s))throw new Error('Tanner stage must be 1–5.');
    return Number(s);
  }
  if([6,7,9,12,14,16,18,20,22,24,26,28,31,32,34,35].includes(i)){
    if(!/^\d+(?:\.\d{1,2})?$/.test(s)||Number(s)>100000)
      throw new Error(TDT_HEADERS[i]+' requires a valid nonnegative number.');
    return Number(s);
  }
  // Preserve CR, name, free-text comments, original image links and other source data.
  if(s[0]==='=')return "'"+s;
  return s;
}
function tdtSaveRecord(payload) {
  tdtAuthorized_(true);
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw new Error('Another save is currently running. Please retry.');
  try{
    const sh=tdtSheet_();tdtCheckSchema_(sh);
    let extraCount=tdtExtraCount_(sh);
    if(!payload||!Array.isArray(payload.values)||payload.values.length!==39)
      throw new Error('Expected all 39 original fields. Nothing was saved.');
    const incoming=payload.values.slice();
    let rowNumber=Number(payload.rowNumber)||0;
    const editing=Boolean(rowNumber);
    if(editing&&(!Number.isInteger(rowNumber)||rowNumber<=TDT_CONFIG.headerRow||rowNumber>sh.getLastRow()))
      throw new Error('Invalid record row.');
    if(!editing)rowNumber=sh.getLastRow()+1;
    const width=39+extraCount*2;
    const previous=editing?sh.getRange(rowNumber,1,1,width).getValues()[0]:Array(width).fill('');
    const userIndices=editing?(Array.isArray(payload.changedIndices)?payload.changedIndices:[]):TDT_HEADERS.map((_,i)=>i);
    if(userIndices.some(i=>!Number.isInteger(i)||i<0||i>=39))throw new Error('Invalid changed-field reference.');
    const changed=new Set(userIndices);
    // The first column always records the actual saving date in IST for new entries.
    // It remains immutable for existing patient records.
    const today=Utilities.formatDate(new Date(),TDT_CONFIG.timeZone,'yyyy-MM-dd');
    if(!editing){incoming[0]=today;changed.add(0);}
    else changed.delete(0);
    for(const i of [0,1,2]){
      const val=changed.has(i)?String(incoming[i]??'').trim():tdtDisplay_(previous[i],i).trim();
      if(!val)throw new Error('Date, Name and CR . No are required.');
    }
    if(editing && tdtFingerprint_(previous)!==String(payload.fingerprint||''))
      throw new Error('Record changed in the Sheet. Reload it before saving to prevent overwriting newer data.');
    // Require the newly selected rapid-entry clinical data to have validated formats.
    // Legacy untouched values remain preserved.
    for(const i of [3,4,5,6,7,10,11,30])if(changed.has(i))tdtCleanValue_(incoming[i],previous[i],i);
    if(TDT_CONFIG.preventDuplicateCr && (!editing||changed.has(2))){
      const proposed=(changed.has(2)?String(incoming[2]??''):tdtDisplay_(previous[2],2)).trim().toLowerCase();
      const total=Math.max(0,sh.getLastRow()-TDT_CONFIG.headerRow);
      if(proposed&&total){
        const rows=sh.getRange(TDT_CONFIG.headerRow+1,3,total,1).getDisplayValues();
        const duplicate=rows.some((rr,j)=>TDT_CONFIG.headerRow+1+j!==rowNumber&&String(rr[0]||'').trim().toLowerCase()===proposed);
        if(duplicate)throw new Error('This CR number already exists. Open the patient using Find by CR before editing.');
      }
    }
    // IAP 2015 50th percentile height/weight are calculated by the server. No false
    // reference values outside 5–18 years or for unrecognized sex.
    if(!editing || changed.has(3)||changed.has(4)){
      const age=changed.has(3)?incoming[3]:tdtDisplay_(previous[3],3);
      const sex=changed.has(4)?incoming[4]:tdtDisplay_(previous[4],4);
      const median=tdtIap50_(age,sex);
      if(median){incoming[31]=String(median.height);incoming[32]=String(median.weight);changed.add(31);changed.add(32);}
      else if(changed.has(3)||changed.has(4)||!editing){
        // Clear a now-invalid reference rather than retaining an incorrect old P50.
        incoming[31]='';incoming[32]='';changed.add(31);changed.add(32);
      }
    }
    const payloadExtra=Array.isArray(payload.extraFerritin)?payload.extraFerritin:[];
    if(payloadExtra.length>TDT_CONFIG.maxFerritinResults-9)throw new Error('Too many ferritin results.');
    const changedExtras=new Set(editing?(Array.isArray(payload.changedExtraIndices)?payload.changedExtraIndices:[]):payloadExtra.map((_,i)=>i));
    if([...changedExtras].some(i=>!Number.isInteger(i)||i<0||i>=payloadExtra.length))
      throw new Error('Invalid additional ferritin field index.');
    const extrasClean=[];
    for(let i=0;i<payloadExtra.length;i++){
      if(!changedExtras.has(i)){extrasClean.push(null);continue;}
      const raw=payloadExtra[i];
      if(!raw||typeof raw!=='object')throw new Error('Invalid additional ferritin entry.');
      // A pair may be cleared when explicitly edited; it is never silently erased.
      extrasClean.push(tdtExtraInput_(raw));
    }
    // Base ferritin pairs must have a valid value and date when the operator changes
    // either half, but old incomplete legacy entries may remain untouched.
    for(let i=12;i<30;i+=2){
      if(!changed.has(i)&&!changed.has(i+1)&&editing)continue;
      const value=String(changed.has(i)?incoming[i]??'':tdtDisplay_(previous[i],i)).trim();
      const dt=String(changed.has(i+1)?incoming[i+1]??'':tdtDisplay_(previous[i+1],i+1)).trim();
      if(value||dt){
        if(!value||!dt)throw new Error('Ferritin '+(1+(i-12)/2)+' requires both value and date.');
        if(!tdtValidISO_(dt,false))throw new Error('Invalid ferritin date.');
        if(!Number.isFinite(Number(value))||Number(value)<0||Number(value)>100000)throw new Error('Invalid ferritin value.');
      }
    }
    const formulas=editing?sh.getRange(rowNumber,1,1,39).getFormulas()[0]:Array(39).fill('');
    const changes=[];
    for(let i=0;i<39;i++){
      if(!changed.has(i))continue;
      if(formulas[i]){
        if(String(incoming[i])!==tdtDisplay_(previous[i],i))throw new Error('Column '+(i+1)+' contains a formula and cannot be edited.');
        continue;
      }
      const text=String(incoming[i]??'').trim();
      if(editing&&text===tdtDisplay_(previous[i],i).trim())continue;
      changes.push({col:i+1,val:tdtCleanValue_(text,previous[i],i)});
    }
    // Only expand columns when a new extra ferritin pair has an actual value/date.
    let desired=extraCount;
    for(let i=0;i<extrasClean.length;i++)if(extrasClean[i]&&(extrasClean[i].value!==''||extrasClean[i].date!==''))desired=Math.max(desired,i+1);
    if(desired>extraCount)extraCount=tdtExtendFerritin_(sh,desired);
    const extraFormulas=(editing&&extraCount)?sh.getRange(rowNumber,40,1,extraCount*2).getFormulas()[0]:[];
    for(let i=0;i<extrasClean.length;i++){
      const extra=extrasClean[i];if(!extra)continue;
      if(extraFormulas[2*i]||extraFormulas[2*i+1])throw new Error('Ferritin '+(i+10)+' contains a spreadsheet formula; edit is blocked.');
      const oldValue=i<(previous.length-39)/2?previous[39+i*2]:'';
      const oldDate=i<(previous.length-39)/2?previous[39+i*2+1]:'';
      if(i>=extraCount)continue;
      if(!editing||String(extra.value)!==tdtDisplay_(oldValue,39+i*2))changes.push({col:40+i*2,val:extra.value});
      if(!editing||String(extra.date)!==tdtDisplay_(oldDate,40+i*2))changes.push({col:41+i*2,val:extra.date});
    }
    if(!changes.length&&editing)return {saved:true,created:false,item:tdtToItem_(rowNumber,previous)};
    // Changes are written only to their targeted columns, never rewriting legacy images
    // or user-maintained formulas in untouched cells.
    changes.sort((a,b)=>a.col-b.col);
    // Explicit text number-format protects leading zeroes in CR and mobile fields.
    if(changes.some(x=>x.col===3))sh.getRange(rowNumber,3).setNumberFormat('@');
    if(changes.some(x=>x.col===6))sh.getRange(rowNumber,6).setNumberFormat('@');
    let start=changes[0].col,part=[changes[0].val];
    function write(){sh.getRange(rowNumber,start,1,part.length).setValues([part]);}
    for(let i=1;i<changes.length;i++){
      if(changes[i].col===start+part.length)part.push(changes[i].val);
      else{write();start=changes[i].col;part=[changes[i].val];}
    }
    write();SpreadsheetApp.flush();
    const full=sh.getRange(rowNumber,1,1,39+2*extraCount).getValues()[0];
    return {saved:true,created:!editing,item:tdtToItem_(rowNumber,full),extraFerritinCount:extraCount};
  }finally{lock.releaseLock();}
}
function tdtUploadImage(slot,filename,mimeType,dataUrl) {
  tdtAuthorized_(true);
  const i=Number(slot);
  if(!Number.isInteger(i)||i<1||i>3)throw new Error('Invalid image slot.');
  const mime=String(mimeType||'').toLowerCase();
  if(!['image/jpeg','image/png','image/webp'].includes(mime))throw new Error('Only JPG, PNG or WebP images are supported.');
  const str=String(dataUrl||'');
  const match=str.match(/^data:(image\/jpeg|image\/png|image\/webp);base64,([a-z0-9+/=]+)$/i);
  if(!match||match[1].toLowerCase()!==mime)throw new Error('Invalid image file.');
  if(match[2].length>TDT_CONFIG.imageMaxBytes*4/3+8)throw new Error('Image exceeds the 2 MB limit. Compress before uploading.');
  const bytes=Utilities.base64Decode(match[2]);
  if(bytes.length>TDT_CONFIG.imageMaxBytes)throw new Error('Image exceeds the 2 MB limit.');
  if(!TDT_CONFIG.imageFolderId)throw new Error('Private image folder ID not configured. Image upload is disabled.');
  const folder=DriveApp.getFolderById(TDT_CONFIG.imageFolderId);
  if(folder.getSharingAccess()!==DriveApp.Access.PRIVATE)
    throw new Error('Image folder has broad link/domain access. Restrict it to specific authorized people before uploading.');
  // UUID filename intentionally excludes patient name and CR number.
  const ext=mime==='image/jpeg'?'jpg':mime==='image/png'?'png':'webp';
  const blob=Utilities.newBlob(bytes,mime,Utilities.getUuid()+'-slot'+i+'.'+ext);
  const file=folder.createFile(blob);
  // No anyone-with-link sharing and no public visibility are enabled.
  return {url:file.getUrl(),name:file.getName()};
}


/**
 * Run this manually in Apps Script to verify the target tab and header mapping.
 * Returns only non-patient metadata; never logs personal or clinical values.
 */
function tdtValidateSetup() {
  tdtAuthorized_(false);
  const sheet=tdtSheet_();
  const headings=tdtCheckSchema_(sheet);
  return {
    project:'TDT DATA',
    tabName:sheet.getName(),
    tabGid:sheet.getSheetId(),
    columns:headings.length,
    records:Math.max(0,sheet.getLastRow()-TDT_CONFIG.headerRow),
    writeEnabled:TDT_CONFIG.allowWrites,
    mapping:'PASS'
  };
}