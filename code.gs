const APP_VERSION = '1.1.0';
const SHEET_SAHIN = "سادس";
const SHEET_AWAL = "أول متوسط";
const SHEET_SETTINGS = "إعدادات_الواجبات";
const SHEET_QUESTIONS = "بنك_الأسئلة_سادس";
const SHEET_QUESTIONS_FIRST = "بنك_الأسئلة_أول_متوسط";
function questionsSheetName_(className) {
  return normalizeClassName(className) === 'أول متوسط' ? SHEET_QUESTIONS_FIRST : SHEET_QUESTIONS;
}
const SHEET_NOTIFICATIONS = "تنبيهات";
const SHEET_CHECK_SIXTH = "تحقق سادس";
const SHEET_CHECK_FIRST = "تحقق أول متوسط";


function normalizeSectionValue_(value){var s=String(value==null?'':value).trim().replace(/[،,]/g,'.');if(/^\d+(?:\.0+)?$/.test(s))return String(Number(s));return s;}
function normalizeClassName(name) {
  var value = String(name || '').trim();
  if (value === 'أولى متوسط' || value === 'اولى متوسط' || value === 'أول متوسط') return 'أول متوسط';
  return value;
}
function findClassSheet(ss, className) {
  var wanted = normalizeClassName(className);
  var names = wanted === 'أول متوسط' ? ['أول متوسط','أولى متوسط','اولى متوسط'] : [String(className).trim()];
  for (var i = 0; i < names.length; i++) { var sh = ss.getSheetByName(names[i]); if (sh) return sh; }
  return null;
}
function displayAssignmentId(value) { return String(value == null ? '' : value).trim().replace(/^'+/, ''); }
function canonicalAssignmentId(value) {
  var s = displayAssignmentId(value).replace(/\s/g, '');
  var m = s.match(/^(\d{1,2})[-\/]\s*(\d{1,2})(?:[-\/]\d{2,4})?$/);
  if (m) return Number(m[1]) + '-' + Number(m[2]);
  var d = new Date(s);
  if (!isNaN(d.getTime()) && /^\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}$/.test(s)) return d.getMonth() + 1 + '-' + d.getDate();
  return s;
}

const SHEET_SYSTEM = "إعدادات_النظام";
const DEFAULT_TEACHER_HASH = "43f64dc77762f69f9f52d5f70b53170679cb9abfc688f4cf77bdfc8077f022bc";
function hashSecret_(value) {
  var bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value == null ? '' : value), Utilities.Charset.UTF_8);
  return bytes.map(function(b){var v=b<0?b+256:b;return ('0'+v.toString(16)).slice(-2);}).join('');
}
function getTeacherHash_() {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(SHEET_SYSTEM);
  if (!sh) { sh=ss.insertSheet(SHEET_SYSTEM); sh.getRange('A1:B2').setValues([['الإعداد','القيمة'],['كلمة سر المعلم',DEFAULT_TEACHER_HASH]]); sh.hideSheet(); return DEFAULT_TEACHER_HASH; }
  var value=sh.getRange('B2').getDisplayValue().trim();
  if (!value) { sh.getRange('B2').setValue(DEFAULT_TEACHER_HASH); sh.hideSheet(); return DEFAULT_TEACHER_HASH; }
  sh.hideSheet(); return value;
}
function verifyTeacherPassword(password) { return hashSecret_(password) === getTeacherHash_(); }
function loginTeacher(password){var props=PropertiesService.getScriptProperties(),now=Date.now(),locked=Number(props.getProperty('teacher_lock_until')||0),fails=Number(props.getProperty('teacher_failed_count')||0);if(locked>now)return {success:false,message:'تم إيقاف المحاولات مؤقتًا. حاول بعد '+Math.ceil((locked-now)/60000)+' دقيقة'};if(!verifyTeacherPassword(password)){fails++;props.setProperty('teacher_failed_count',String(fails));if(fails>=5){props.setProperty('teacher_lock_until',String(now+15*60000));props.setProperty('teacher_failed_count','0');return {success:false,message:'تم إيقاف الدخول 15 دقيقة بعد محاولات فاشلة كثيرة'};}return {success:false,message:'كلمة سر المعلم غير صحيحة، المحاولة '+fails+' من 5'};}props.deleteProperty('teacher_failed_count');props.deleteProperty('teacher_lock_until');var token=Utilities.getUuid()+'-'+Utilities.getUuid();props.setProperty('teacher_session_'+token,'1');return {success:true,token:token,expiresIn:0};}
function requireTeacherSession_(token){var props=PropertiesService.getScriptProperties();if(!token||props.getProperty('teacher_session_'+String(token))!=='1')throw new Error('جلسة المعلم غير صالحة. سجّل الدخول مرة أخرى.');return true;}
function logoutTeacher(token){if(token)PropertiesService.getScriptProperties().deleteProperty('teacher_session_'+String(token));return {success:true};}

function changeTeacherPassword(oldPassword,newPassword,token) { requireTeacherSession_(token);
  if (!verifyTeacherPassword(oldPassword)) return {success:false,message:'كلمة السر الحالية غير صحيحة'};
  if (String(newPassword || '').trim().length < 3) return {success:false,message:'كلمة السر الجديدة يجب أن تكون 3 رموز على الأقل'};
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(SHEET_SYSTEM) || ss.insertSheet(SHEET_SYSTEM);
  sh.getRange('A1:B2').setValues([['الإعداد','القيمة'],['كلمة سر المعلم',hashSecret_(newPassword)]]); sh.hideSheet();
  return {success:true,message:'تم تغيير كلمة سر المعلم بنجاح'};
}

function ensureNotificationsSheet_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(SHEET_NOTIFICATIONS);
  if(!sh){sh=ss.insertSheet(SHEET_NOTIFICATIONS);sh.getRange(1,1,1,5).setValues([['الصف','نص التنبيه','من','إلى','تاريخ الإنشاء']]);}
  else if(sh.getLastRow()<1)sh.getRange(1,1,1,5).setValues([['الصف','نص التنبيه','من','إلى','تاريخ الإنشاء']]);
  return sh;
}
function getTeacherNotifications(token){requireTeacherSession_(token);var sh=ensureNotificationsSheet_();if(sh.getLastRow()<2)return[];return sh.getRange(2,1,sh.getLastRow()-1,5).getDisplayValues().map(function(r,i){return {row:i+2,className:r[0],message:r[1],startDate:r[2],endDate:r[3],createdAt:r[4]};});}
function saveTeacherNotification(item,token){requireTeacherSession_(token);var cls=normalizeClassName(item&&item.className),msg=String(item&&item.message||'').trim(),start=String(item&&item.startDate||'').trim(),end=String(item&&item.endDate||'').trim();if(!cls||!msg||!start||!end) return {success:false,message:'أكمل الصف والرسالة وفترة النشر'};if(new Date(start).getTime()>=new Date(end).getTime())return {success:false,message:'يجب أن يكون وقت الانتهاء بعد وقت البدء'};var sh=ensureNotificationsSheet_(),row=Number(item&&item.row||0),vals=[[cls,msg,start,end,new Date()]];if(row>=2&&row<=sh.getLastRow())sh.getRange(row,1,1,5).setValues(vals);else{sh.appendRow(vals[0]);row=sh.getLastRow();}return {success:true,message:'تم حفظ التنبيه بنجاح',row:row};}
function deleteTeacherNotification(row,token){requireTeacherSession_(token);var sh=ensureNotificationsSheet_(),n=Number(row);if(n>=2&&n<=sh.getLastRow())sh.deleteRow(n);return {success:true,message:'تم حذف التنبيه'};}
function getActiveStudentNotifications(className){var sh=ensureNotificationsSheet_(),now=Date.now(),out=[];if(sh.getLastRow()<2)return out;sh.getRange(2,1,sh.getLastRow()-1,5).getDisplayValues().forEach(function(r){if(normalizeClassName(r[0])!==normalizeClassName(className))return;var a=new Date(r[2]).getTime(),b=new Date(r[3]).getTime();if((!a||now>=a)&&(!b||now<=b))out.push({message:r[1],startDate:r[2],endDate:r[3]});});return out;}

const SHEET_ADVERTISEMENTS = "إعلانات";
const ADVERTISEMENT_HEADERS = ['الصف','النوع','الرمز','اللون','المحتوى','كرابط','من','إلى','تاريخ الإنشاء','حجم الفقاعة','نبض الفقاعة','محاذاة النص'];
function ensureAdvertisementsSheet_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(SHEET_ADVERTISEMENTS);
  if(!sh) sh=ss.insertSheet(SHEET_ADVERTISEMENTS);
  if(sh.getLastColumn()<ADVERTISEMENT_HEADERS.length) sh.insertColumnsAfter(Math.max(1,sh.getLastColumn()),ADVERTISEMENT_HEADERS.length-Math.max(1,sh.getLastColumn()));
  sh.getRange(1,1,1,ADVERTISEMENT_HEADERS.length).setValues([ADVERTISEMENT_HEADERS]);
  return sh;
}
function advertisementColor_(value){
  var allowed=['red','green','blue','orange','purple','white','teal','pink'];
  var v=String(value||'red').trim().toLowerCase();
  return allowed.indexOf(v)>=0?v:'red';
}
function advertisementSize_(value){
  var allowed=['medium','small','xsmall','large'];
  var v=String(value||'medium').trim().toLowerCase();
  if(v==='large') return 'medium';
  return allowed.indexOf(v)>=0?v:'medium';
}
function advertisementAnimation_(value){
  var allowed=['pulse','bounce','float','shake','none'];
  var v=String(value||'pulse').trim().toLowerCase();
  return allowed.indexOf(v)>=0?v:'pulse';
}
function advertisementAlign_(value){
  var allowed=['center','right','left'];
  var v=String(value||'center').trim().toLowerCase();
  return allowed.indexOf(v)>=0?v:'center';
}
function getTeacherAdvertisements(token){
  requireTeacherSession_(token);var sh=ensureAdvertisementsSheet_();if(sh.getLastRow()<2)return[];
  return sh.getRange(2,1,sh.getLastRow()-1,ADVERTISEMENT_HEADERS.length).getDisplayValues().map(function(r,i){return {row:i+2,className:r[0],type:r[1],icon:r[2],color:advertisementColor_(r[3]),content:r[4],isLink:String(r[5]).toLowerCase()==='true',startDate:r[6],endDate:r[7],createdAt:r[8],size:advertisementSize_(r[9]),animation:advertisementAnimation_(r[10]),align:advertisementAlign_(r[11])};});
}
function saveTeacherAdvertisement(item,token){
  requireTeacherSession_(token);item=item||{};var cls=normalizeClassName(item.className),type=String(item.type||'إعلان').trim()||'إعلان',icon=String(item.icon||'📢').trim()||'📢',content=String(item.content||''),start=String(item.startDate||'').trim(),end=String(item.endDate||'').trim(),isLink=!!item.isLink||/^https?:\/\//i.test(content);
  if(!cls||!content.trim()||!start||!end)return {success:false,message:'أكمل الصف والمحتوى وفترة النشر'};
  if(!/^https?:\/\//i.test(content)&&isLink)return {success:false,message:'الرابط يجب أن يبدأ بـ http:// أو https://'};
  if(new Date(start).getTime()>=new Date(end).getTime())return {success:false,message:'يجب أن يكون وقت الانتهاء بعد وقت البدء'};
  var sh=ensureAdvertisementsSheet_(),row=Number(item.row||0),vals=[[cls,type,icon,advertisementColor_(item.color),content,isLink,start,end,new Date(),advertisementSize_(item.size),advertisementAnimation_(item.animation),advertisementAlign_(item.align)]];
  if(row>=2&&row<=sh.getLastRow())sh.getRange(row,1,1,ADVERTISEMENT_HEADERS.length).setValues(vals);else{sh.appendRow(vals[0]);row=sh.getLastRow();}
  return {success:true,message:'تم حفظ الإعلان بنجاح',row:row};
}
function deleteTeacherAdvertisement(row,token){
  requireTeacherSession_(token);var sh=ensureAdvertisementsSheet_(),n=Number(row);if(n>=2&&n<=sh.getLastRow())sh.deleteRow(n);return {success:true,message:'تم حذف الإعلان'};
}
function getActiveStudentAdvertisements(className){
  var sh=ensureAdvertisementsSheet_(),now=Date.now(),out=[];if(sh.getLastRow()<2)return out;
  sh.getRange(2,1,sh.getLastRow()-1,ADVERTISEMENT_HEADERS.length).getDisplayValues().forEach(function(r){if(normalizeClassName(r[0])!==normalizeClassName(className))return;var a=new Date(r[6]).getTime(),b=new Date(r[7]).getTime();if((!a||now>=a)&&(!b||now<=b)){var link=String(r[5]).toLowerCase()==='true'||/^https?:\/\//i.test(String(r[4]||''));out.push({row:r[0]+'-'+r[8],type:r[1]||'إعلان',icon:r[2]||'📢',color:advertisementColor_(r[3]),content:r[4],isLink:link,startDate:r[6],endDate:r[7],size:advertisementSize_(r[9]),animation:advertisementAnimation_(r[10]),align:advertisementAlign_(r[11])});}});return out;
}

function checkSheetName_(className){return normalizeClassName(className)==='أول متوسط'?SHEET_CHECK_FIRST:SHEET_CHECK_SIXTH;}
function ensureCheckSheet_(className){var ss=SpreadsheetApp.getActiveSpreadsheet(),name=checkSheetName_(className),sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()<1)sh.getRange(1,1,1,2).setValues([['الفصل','اسم الطالب']]);return sh;}
function checkLesson_(v){var raw=String(v==null?'':v).trim(),id=canonicalAssignmentId(raw);return /^\d{1,2}-\d{1,2}$/.test(String(id||''))?id:'';}
function getCheckSheetData_(className){var sh=ensureCheckSheet_(className),lastCol=Math.max(2,sh.getLastColumn()),rawHeaders=sh.getRange(1,1,1,lastCol).getDisplayValues()[0],valid=[];rawHeaders.slice(2).forEach(function(h,j){var id=checkLesson_(h);if(id)valid.push({id:id,col:j+2});});var rawRows=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,lastCol).getDisplayValues():[],rows=rawRows.map(function(r){return [r[0]||'',r[1]||''].concat(valid.map(function(x){return r[x.col]||'';}));});return {sheet:sh,headers:valid.map(function(x){return x.id;}),rows:rows};}
function getCheckGradesForStudent_(className,section,name){var d=getCheckSheetData_(className),out={},idx=-1;for(var i=0;i<d.rows.length;i++)if(normalizeSectionValue_(d.rows[i][0])===normalizeSectionValue_(section)&&String(d.rows[i][1]).trim()===String(name||'').trim()){idx=i;break;}if(idx<0)return out;d.headers.forEach(function(h,j){if(h&&String(d.rows[idx][j+2]||'').trim()!=='')out[h]=d.rows[idx][j+2];});return out;}
function getUnderstandingChecks(className,section,token){requireTeacherSession_(token);var ss=SpreadsheetApp.getActiveSpreadsheet(),d=getCheckSheetData_(className),settings=ss.getSheetByName(SHEET_SETTINGS),headers=d.headers.slice();if(settings&&settings.getLastRow()>=2){settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues().forEach(function(r){if(normalizeClassName(r[0])===normalizeClassName(className)&&String(r[3]).trim()==='منشور'){var id=canonicalAssignmentId(r[1]);if(id&&headers.indexOf(id)<0)headers.push(id);}});}var sh=findClassSheet(ss,className),students=[];if(sh&&sh.getLastRow()>=2)sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues().forEach(function(r){if(String(r[1]).trim()&&(section==='__ALL__'||normalizeSectionValue_(r[2])===normalizeSectionValue_(section)))students.push({name:String(r[1]).trim(),section:String(r[2]).trim()});});var rows=d.rows.filter(function(r){return (section==='__ALL__'||normalizeSectionValue_(r[0])===normalizeSectionValue_(section));});return {headers:headers,rows:rows,students:sortStudentRows_(students,section==='__ALL__')};}
function saveUnderstandingChecks(className,section,headers,rows,token){requireTeacherSession_(token);var sh=ensureCheckSheet_(className),valid=(headers||[]).map(checkLesson_).filter(Boolean),lastCol=Math.max(2,sh.getLastColumn()),existing=sh.getRange(1,1,1,lastCol).getDisplayValues()[0].slice(2).map(checkLesson_);valid.forEach(function(h){if(existing.indexOf(h)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(h);existing.push(h);}});var allHeaders=existing,values=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,Math.max(2,sh.getLastColumn())).getDisplayValues():[];(rows||[]).forEach(function(item){var sectionValue=String(item.section==null?section:item.section).trim(),name=String(item.name||'').trim();if(!name)return;var ri=-1;for(var i=0;i<values.length;i++)if(String(values[i][0]).trim()===sectionValue&&String(values[i][1]).trim()===name){ri=i;break;}if(ri<0){sh.appendRow([sectionValue,name]);values.push([sectionValue,name]);ri=values.length-1;}var rowNum=ri+2;allHeaders.forEach(function(h,j){if(Object.prototype.hasOwnProperty.call(item.scores||{},h)){var raw=item.scores[h],txt=String(raw==null?'':raw).trim();if(txt===''){sh.getRange(rowNum,j+3).clearContent();}else{var n=Number(String(txt).replace(',','.'));if(!isNaN(n)&&n>=0&&n<=10)sh.getRange(rowNum,j+3).setNumberFormat('0.##').setValue(n);}}});});return {success:true,message:'تم حفظ درجات تحقق الفهم بنجاح'};}
function parseCsvChecks_(text){var lines=String(text||'').replace(/^\uFEFF/,'').split(/\r?\n/).filter(function(x){return x.trim()!=='';});return lines.map(function(line){var delim=line.indexOf('\t')>=0?'\t':',';var out=[],cur='',quote=false;for(var i=0;i<line.length;i++){var c=line[i];if(c==='"'){if(quote&&line[i+1]==='"'){cur+='"';i++;}else quote=!quote;}else if(c===delim&&!quote){out.push(cur.trim());cur='';}else cur+=c;}out.push(cur.trim());return out;});}
function xmlEscape_(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}
function xmlUnescape_(v){return String(v||'').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');}
function colLetter_(n){var out='';while(n>0){var r=(n-1)%26;out=String.fromCharCode(65+r)+out;n=Math.floor((n-1)/26);}return out;}
function makeXlsxBlob_(rows,name){var sheetRows=[];rows.forEach(function(row,ri){var cells=[];row.forEach(function(v,ci){var text=String(v==null?'':v);var ref=colLetter_(ci+1)+(ri+1);if(ri>0&&ci>1&&text!==''&&!isNaN(Number(text))){cells.push('<c r="'+ref+'"><v>'+xmlEscape_(text)+'</v></c>');}else{cells.push('<c r="'+ref+'" t="inlineStr"><is><t>'+xmlEscape_(text)+'</t></is></c>');}});sheetRows.push('<row r="'+(ri+1)+'">'+cells.join('')+'</row>');});var sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:'+colLetter_(Math.max(1,rows[0].length))+rows.length+'"/><sheetData>'+sheetRows.join('')+'</sheetData></worksheet>';var workbook='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="تحقق" sheetId="1" r:id="rId1"/></sheets></workbook>';var rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';var wbRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';var types='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>';var blobs=[Utilities.newBlob(types,'application/xml','[Content_Types].xml'),Utilities.newBlob(rels,'application/xml','_rels/.rels'),Utilities.newBlob(workbook,'application/xml','xl/workbook.xml'),Utilities.newBlob(wbRels,'application/xml','xl/_rels/workbook.xml.rels'),Utilities.newBlob(sheet,'application/xml','xl/worksheets/sheet1.xml')];return Utilities.zip(blobs,name);}
function parseXlsxRows_(bytes){var files=Utilities.unzip(Utilities.newBlob(bytes,'application/zip','import.xlsx')),rows=[],sheet=null,shared=[];files.forEach(function(f){if(f.getName()==='xl/sharedStrings.xml'){var sx=f.getDataAsString(),sr=/<si[\s\S]*?<\/si>/g,sm;while((sm=sr.exec(sx))){var ts=[],tr=/<t[^>]*>([\s\S]*?)<\/t>/g,tm;while((tm=tr.exec(sm[0])))ts.push(tm[1]);shared.push(xmlUnescape_(ts.join('')));}}if(/xl\/worksheets\/sheet1\.xml$/i.test(f.getName()))sheet=f;});if(!sheet)return rows;var xml=sheet.getDataAsString(),rowRe=/<row[\s\S]*?<\/row>/g,rm;while((rm=rowRe.exec(xml))){var cells=[],cellRe=/<c([^>]*)>([\s\S]*?)<\/c>/g,cm;while((cm=cellRe.exec(rm[0]))){var attrs=cm[1],body=cm[2],v='',type=(attrs.match(/\bt="([^"]+)"/)||[])[1]||'';if(type==='inlineStr'){v=(body.match(/<t[^>]*>([\s\S]*?)<\/t>/)||[])[1]||'';}else{v=(body.match(/<v[^>]*>([\s\S]*?)<\/v>/)||[])[1]||'';if(type==='s'&&shared[Number(v)]!==undefined)v=shared[Number(v)];}cells.push(xmlUnescape_(v.replace(/<[^>]+>/g,'')));}rows.push(cells);}return rows;}
function importUnderstandingChecksFile(className,section,dataUrl,token){requireTeacherSession_(token);var parts=String(dataUrl||'').split(','),mime=(parts[0]||''),bytes=Utilities.base64Decode(parts[1]||''),rows=[];if(/spreadsheetml|xlsx/i.test(mime)||/\.xlsx/i.test(mime)){rows=parseXlsxRows_(bytes);}else{rows=parseCsvChecks_(Utilities.newBlob(bytes,'text/plain','import.csv').getDataAsString('UTF-8'));}if(rows.length<1)return {success:false,message:'الملف فارغ أو غير صالح'};var first=rows[0].map(function(x){return String(x||'').trim();}),nameCol=-1,sectionCol=-1;first.forEach(function(h,i){var x=h.toLowerCase();if(x==='اسم الطالب'||x==='الاسم'||x==='name'||x==='student name')nameCol=i;if(x==='الفصل'||x==='section'||x==='الشعبة')sectionCol=i;});if(nameCol<0)nameCol=0;if(sectionCol<0&&nameCol===0&&first.length>1&&(/فصل|section|الشعبة/i.test(first[1])))sectionCol=1;if(sectionCol<0)sectionCol=1;var lessonCols=[];first.forEach(function(h,i){var id=checkLesson_(h);if(i!==nameCol&&i!==sectionCol&&id)lessonCols.push({index:i,id:id});});var data=rows.slice(1).map(function(r){var nm=String(r[nameCol]||'').trim(),sec=sectionCol<r.length?String(r[sectionCol]||'').trim():String(section||'').trim(),scores={};lessonCols.forEach(function(x){var v=String(r[x.index]||'').trim();if(v!=='')scores[x.id]=v;});return {section:sec||section,name:nm,scores:scores};}).filter(function(x){return x.name;});return saveUnderstandingChecks(className,section,lessonCols.map(function(x){return x.id;}),data,token);}
function exportUnderstandingChecks(className,section,token){requireTeacherSession_(token);var ss=SpreadsheetApp.getActiveSpreadsheet(),d=getCheckSheetData_(className),headers=d.headers.slice(),settings=ss.getSheetByName(SHEET_SETTINGS);if(settings&&settings.getLastRow()>=2)settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues().forEach(function(r){if(normalizeClassName(r[0])===normalizeClassName(className)&&String(r[3]).trim()==='منشور'){var id=canonicalAssignmentId(r[1]);if(id&&headers.indexOf(id)<0)headers.push(id);}});var sh=findClassSheet(ss,className),saved={};d.rows.forEach(function(r){var key=String(r[0]).trim()+'\u0000'+String(r[1]).trim();if(section==='__ALL__'||String(r[0]).trim()===String(section).trim())saved[key]=r;});var rows=[['اسم الطالب','الفصل'].concat(headers)];if(sh&&sh.getLastRow()>=2)sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues().forEach(function(r){if(section!=='__ALL__'&&String(r[2]).trim()!==String(section).trim())return;if(!String(r[1]).trim())return;var old=saved[String(r[2]).trim()+'\u0000'+String(r[1]).trim()]||[];rows.push([r[1],r[2]].concat(headers.map(function(_,j){return old[j+2]||'';})));});var name='تحقق من فهمك - '+normalizeClassName(className)+' - فصل '+section+'.xlsx',blob=makeXlsxBlob_(rows,name);return {success:true,name:name,dataUrl:'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,'+Utilities.base64Encode(blob.getBytes())};}

const SHEET_VIOLATIONS = 'المخالفات';
const SHEET_STUDENT_TRASH = 'سلة محذوفات الطلاب';
const SHEET_STUDENT_LOG = 'سجل تغييرات الطلاب';
function ensureViolationsSheet_(){var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(SHEET_VIOLATIONS);if(!sh){sh=ss.insertSheet(SHEET_VIOLATIONS);sh.getRange(1,1,1,5).setValues([['المرحلة','اسم الطالب','الفصل','التاريخ','نوع المخالفة']]);}else if(sh.getLastRow()<1)sh.getRange(1,1,1,5).setValues([['المرحلة','اسم الطالب','الفصل','التاريخ','نوع المخالفة']]);return sh;}
function importTrackingWorkbookSheets(sheets,token){requireTeacherSession_(token);var names=Object.keys(sheets||{}),total=0;names.forEach(function(sheetName){var rows=sheets[sheetName]||[],lower=String(sheetName).toLowerCase(),isV=/مخالف|violation/.test(lower),cls=/أول|اول|first|متوسط/.test(lower)?SHEET_AWAL:SHEET_SAHIN;if(isV){var sh=ensureViolationsSheet_(),head=rows.length?rows[0].map(function(x){return String(x||'').trim();}):[],idx=function(a,d){var i=head.findIndex(function(x){return a.indexOf(x)>=0;});return i<0?d:i;},ni=idx(['اسم الطالب','الاسم'],0),si=idx(['الفصل','الشعبة'],1),di=idx(['التاريخ','Date'],2),ti=idx(['نوع المخالفة','المخالفة','Violation'],3),existing=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,5).getDisplayValues():[],out=[];(rows.slice(1)||[]).forEach(function(r){var row=[cls,String(r[ni]||'').trim(),String(r[si]||'').trim(),String(r[di]||'').trim(),String(r[ti]||'').trim()];if(!row[1]||!row[4])return;var dup=existing.some(function(e){return e[0]===row[0]&&e[1]===row[1]&&e[2]===row[2]&&e[3]===row[3]&&e[4]===row[4];});if(!dup){out.push(row);existing.push(row);}});if(out.length)sh.getRange(sh.getLastRow()+1,1,out.length,5).setValues(out);total+=out.length;}else{var head=rows.length?rows[0].map(function(x){return String(x||'').trim();}):[],ni=head.findIndex(function(x){return x==='اسم الطالب'||x==='الاسم'||x.toLowerCase()==='name';}),si=head.findIndex(function(x){return x==='الفصل'||x==='الشعبة'||x.toLowerCase()==='section';});if(ni<0)ni=0;if(si<0)si=1;var lessons=head.filter(function(h,i){return i!==ni&&i!==si&&checkLesson_(h);}),items=[];(rows.slice(1)||[]).forEach(function(r){var name=String(r[ni]||'').trim(),section=String(r[si]||'').trim(),scores={};if(!name)return;head.forEach(function(h,i){var id=checkLesson_(h),v=String(r[i]||'').trim();if(id&&i!==ni&&i!==si&&v!=='')scores[id]=v;});items.push({name:name,section:section,scores:scores});});if(items.length)saveUnderstandingChecks(cls,'__ALL__',lessons,items,token);total+=items.length;}});return {success:true,message:'تمت مطابقة واستيراد تحقق من فهمك والمخالفات دون تكرار. السجلات المعالجة: '+total};}
function importTrackingWorkbook(checkRows,violationRows,className,token){var sheets={};sheets['تحقق '+normalizeClassName(className)]=checkRows||[];sheets['مخالفات '+normalizeClassName(className)]=violationRows||[];return importTrackingWorkbookSheets(sheets,token);}
function getStudentViolations(className,section,name){var sh=ensureViolationsSheet_(),out=[];if(sh.getLastRow()<2)return out;sh.getRange(2,1,sh.getLastRow()-1,5).getDisplayValues().forEach(function(r){if(normalizeClassName(r[0])===normalizeClassName(className)&&String(r[1]).trim()===String(name||'').trim()&&normalizeSectionValue_(r[2])===normalizeSectionValue_(section))out.push({date:r[3],type:r[4]});});return out;}
function normalizeStudentName_(value){return String(value==null?'':value).replace(/[\u0640]/g,'').replace(/\s+/g,' ').trim();}
function sortStudentRows_(rows,allSections){return (rows||[]).sort(function(a,b){var sa=String(a.section||'').trim(),sb=String(b.section||'').trim();if(allSections){var sc=sa.localeCompare(sb,'ar',{numeric:true,sensitivity:'base'});if(sc)return sc;}return normalizeStudentName_(a.name).localeCompare(normalizeStudentName_(b.name),'ar',{numeric:true,sensitivity:'base'});});}

function ensureStudentTrashSheet_(){var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(SHEET_STUDENT_TRASH);if(!sh)sh=ss.insertSheet(SHEET_STUDENT_TRASH);if(sh.getLastRow()<1)sh.getRange(1,1,1,7).setValues([['تاريخ الحذف','الصف','رقم الصف القديم','كلمة المرور','اسم الطالب','الفصل','بيانات الصف الكاملة']]);return sh;}
function ensureStudentLogSheet_(){var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(SHEET_STUDENT_LOG);if(!sh)sh=ss.insertSheet(SHEET_STUDENT_LOG);if(sh.getLastRow()<1)sh.getRange(1,1,1,7).setValues([['التاريخ','العملية','الصف','كلمة المرور','الاسم القديم','الاسم الجديد','التفاصيل']]);return sh;}
function logStudentChange_(action,className,oldName,newName,id,details){var sh=ensureStudentLogSheet_();sh.appendRow([new Date(),action,normalizeClassName(className),String(id==null?'':id),String(oldName||''),String(newName||''),String(details||'')]);}
function getStudentManagementData(className,token){requireTeacherSession_(token);var sh=findClassSheet(SpreadsheetApp.getActiveSpreadsheet(),className),out=[];if(!sh||sh.getLastRow()<2)return out;sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues().forEach(function(r,i){if(normalizeStudentName_(r[1]))out.push({row:i+2,id:String(r[0]||'').trim(),name:String(r[1]||'').trim(),section:String(r[2]||'').trim()});});return sortStudentRows_(out,true);}
function addStudent(className,name,section,token){requireTeacherSession_(token);var cls=normalizeClassName(className),nm=normalizeStudentName_(name),sec=String(section||'').trim();if(!cls||!nm||!sec)return {success:false,message:'أكمل الصف واسم الطالب والفصل'};var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=findClassSheet(ss,cls);if(!sh)return {success:false,message:'ورقة الصف غير موجودة'};var rows=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues():[];if(rows.some(function(r){return normalizeStudentName_(r[1])===nm;}))return {success:false,message:'اسم الطالب موجود مسبقًا في هذا الصف. استخدم تعديل الطالب أو نقله.'};var id=0;sh.appendRow([id,nm,sec]);logStudentChange_('إضافة',cls,'',nm,id,'إضافة طالب جديد');return {success:true,message:'تمت إضافة الطالب بنجاح',student:{id:String(id),name:nm,section:sec}};}
function updateStudent(className,row,name,section,token){requireTeacherSession_(token);var cls=normalizeClassName(className),sh=findClassSheet(SpreadsheetApp.getActiveSpreadsheet(),cls),n=Number(row),nm=normalizeStudentName_(name),sec=String(section||'').trim();if(!sh||n<2||n>sh.getLastRow()||!nm||!sec)return {success:false,message:'بيانات الطالب غير مكتملة أو غير موجودة'};var old=sh.getRange(n,1,1,3).getDisplayValues()[0],rows=sh.getLastRow()>=2?sh.getRange(2,2,sh.getLastRow()-1,2).getDisplayValues():[];for(var i=0;i<rows.length;i++){var rr=i+2;if(rr!==n&&normalizeStudentName_(rows[i][0])===nm)return {success:false,message:'لا يمكن التعديل: اسم الطالب موجود مسبقًا في هذا الصف'};}sh.getRange(n,2,1,2).setValues([[nm,sec]]);logStudentChange_('تعديل',cls,old[1],nm,old[0],'الفصل: '+old[2]+' ← '+sec);return {success:true,message:'تم تعديل بيانات الطالب بنجاح'};}
function moveStudent(className,row,section,token){requireTeacherSession_(token);var sh=findClassSheet(SpreadsheetApp.getActiveSpreadsheet(),className),n=Number(row);if(!sh||n<2||n>sh.getLastRow())return {success:false,message:'الطالب غير موجود'};var old=sh.getRange(n,1,1,3).getDisplayValues()[0];return updateStudent(className,row,old[1],section,token);}
function deleteStudent(className,row,token){requireTeacherSession_(token);var cls=normalizeClassName(className),sh=findClassSheet(SpreadsheetApp.getActiveSpreadsheet(),cls),n=Number(row);if(!sh||n<2||n>sh.getLastRow())return {success:false,message:'الطالب غير موجود'};var vals=sh.getRange(n,1,1,sh.getLastColumn()).getValues()[0],display=sh.getRange(n,1,1,3).getDisplayValues()[0];if(!normalizeStudentName_(display[1]))return {success:false,message:'الطالب غير موجود'};var trash=ensureStudentTrashSheet_();trash.appendRow([new Date(),cls,n,String(display[0]||''),String(display[1]||''),String(display[2]||''),JSON.stringify(vals)]);sh.deleteRow(n);logStudentChange_('حذف إلى السلة',cls,display[1],'',display[0],'تم حفظ الصف الكامل في سلة المحذوفات');return {success:true,message:'تم نقل الطالب إلى سلة المحذوفات ويمكن استعادته لاحقًا'};}
function getDeletedStudents(token){requireTeacherSession_(token);var sh=ensureStudentTrashSheet_(),out=[];if(sh.getLastRow()<2)return out;sh.getRange(2,1,sh.getLastRow()-1,7).getDisplayValues().forEach(function(r,i){if(r[4])out.push({row:i+2,deletedAt:r[0],className:r[1],oldRow:r[2],id:r[3],name:r[4],section:r[5]});});return out;}
function restoreStudent(trashRow,token){requireTeacherSession_(token);var trash=ensureStudentTrashSheet_(),n=Number(trashRow);if(n<2||n>trash.getLastRow())return {success:false,message:'السجل المحذوف غير موجود'};var r=trash.getRange(n,1,1,7).getDisplayValues()[0],cls=normalizeClassName(r[1]),sh=findClassSheet(SpreadsheetApp.getActiveSpreadsheet(),cls);if(!sh)return {success:false,message:'ورقة الصف غير موجودة'};var existing=sh.getLastRow()>=2?sh.getRange(2,2,sh.getLastRow()-1,1).getDisplayValues():[];if(existing.some(function(x){return normalizeStudentName_(x[0])===normalizeStudentName_(r[4]);}))return {success:false,message:'لا يمكن الاستعادة: يوجد طالب بالاسم نفسه في الصف'};var vals;try{vals=JSON.parse(r[6]||'[]');}catch(e){return {success:false,message:'بيانات السجل المحذوف غير صالحة'};}if(!Array.isArray(vals)||!vals.length)return {success:false,message:'بيانات السجل المحذوف ناقصة'};sh.getRange(sh.getLastRow()+1,1,1,vals.length).setValues([vals]);trash.deleteRow(n);logStudentChange_('استعادة',cls,'',r[4],r[3],'استعادة من سلة المحذوفات');return {success:true,message:'تمت استعادة الطالب بجميع بيانات صفه'};}

function createTeacherBackup(token){requireTeacherSession_(token);var ss=SpreadsheetApp.getActiveSpreadsheet(),stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'GMT','yyyy-MM-dd_HH-mm-ss'),baseName='بوابة الواجبات v.'+APP_VERSION+' - نسخة احتياطية - '+stamp,blobs=[];var spreadsheetId=ss.getId(),sheetUrl='https://docs.google.com/spreadsheets/d/'+spreadsheetId+'/export?format=xlsx';var xlsx=UrlFetchApp.fetch(sheetUrl,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});if(xlsx.getResponseCode()!==200)throw new Error('تعذر تصدير جدول البيانات. تحقق من صلاحيات Drive ثم أعد المحاولة.');blobs.push(Utilities.newBlob(xlsx.getBlob().getBytes(),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','الجدول.xlsx'));var sourceUrl='https://script.googleapis.com/v1/projects/'+encodeURIComponent(ScriptApp.getScriptId())+'/content';var src=UrlFetchApp.fetch(sourceUrl,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});if(src.getResponseCode()!==200)throw new Error('تعذر قراءة ملفات المشروع تلقائيًا. فعّل Google Apps Script API للمشروع ثم أعد المحاولة.');var files=JSON.parse(src.getContentText()).files||[],foundCode=false,foundIndex=false;files.forEach(function(f){if(!f.source)return;var n=String(f.name||''),lower=n.toLowerCase();if(lower==='code.gs'||lower.endsWith('/code.gs')){blobs.push(Utilities.newBlob(f.source,'text/plain','code.gs'));foundCode=true;}else if(lower==='index'||lower==='index.html'||lower.endsWith('/index.html')){blobs.push(Utilities.newBlob(f.source,'text/html','index.txt'));foundIndex=true;}});if(!foundCode||!foundIndex)throw new Error('لم يتم العثور على code.gs أو index في محتوى المشروع.');var manifest={name:'بوابة الواجبات v.'+APP_VERSION,createdAt:new Date().toISOString(),spreadsheetId:spreadsheetId,files:['الجدول.xlsx','code.gs','index.txt']};blobs.push(Utilities.newBlob(JSON.stringify(manifest,null,2),'application/json','معلومات النسخة.json'));var zip=Utilities.zip(blobs,baseName+'.zip'),file=DriveApp.createFile(zip);return {success:true,name:file.getName(),url:file.getDownloadUrl(),fileId:file.getId(),message:'تم إنشاء النسخة الاحتياطية بنجاح: '+file.getName()};}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover');
}

// جلب بيانات الطلاب
function getStudentsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = {};
  
  [SHEET_SAHIN, SHEET_AWAL].forEach(function(sheetName) {
    var sheet = findClassSheet(ss, sheetName);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var values = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
        var students = [];
        values.forEach(function(row) {
          if (row[1]) {
            students.push({ id: String(row[0]).trim(), name: String(row[1]).trim(), section: String(row[2] || '').trim() });
          }
        });
        data[sheetName] = sortStudentRows_(students, true);
      } else {
        data[sheetName] = [];
      }
    }
  });
  return data;
}

// حفظ أو تحديث الواجب مع أسئلته كاملة من التطبيق
function saveCompleteAssignment(data,token) { requireTeacherSession_(token);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var className = data.className;
  var assignmentId = canonicalAssignmentId(data.assignmentId);
  var originalAssignmentId = canonicalAssignmentId(data.originalAssignmentId || assignmentId);
  var title = data.title;
  var questions = data.questions; // مصفوفة الأسئلة
  
  // 1. إدارة جدول الإعدادات
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
    settingsSheet.appendRow(["الصف", "رقم الواجب", "عنوان الدرس", "الحالة", "تاريخ البدء", "تاريخ الانتهاء"]);
  }
  
  var lastRow = settingsSheet.getLastRow();
  var foundRow = -1;
  if (lastRow >= 2) {
    var rows = settingsSheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    for (var i = 0; i < rows.length; i++) {
      if (normalizeClassName(rows[i][0]) === normalizeClassName(className) && (canonicalAssignmentId(rows[i][1]) === originalAssignmentId || (data.isEdit && canonicalAssignmentId(rows[i][1]) === assignmentId))) {
        foundRow = i + 2;
        break;
      }
    }
  }
  
  // في وضع التعديل: ابحث بالرقم الأصلي أو الحالي حتى لو كان تنسيق الرقم في الجدول مختلفاً.
  if (data.isEdit && foundRow === -1 && lastRow >= 2) {
    var editRows = settingsSheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    for (var editIndex = 0; editIndex < editRows.length; editIndex++) {
      if (normalizeClassName(editRows[editIndex][0]) === normalizeClassName(className) && (canonicalAssignmentId(editRows[editIndex][1]) === originalAssignmentId || canonicalAssignmentId(editRows[editIndex][1]) === assignmentId)) {
        foundRow = editIndex + 2;
        break;
      }
    }
  }
  var duplicateSettings = settingsSheet.getRange(2, 1, Math.max(settingsSheet.getLastRow() - 1, 1), 2).getDisplayValues();
  for (var duplicateIndex = 0; duplicateIndex < duplicateSettings.length; duplicateIndex++) {
    var duplicateRowNumber = duplicateIndex + 2;
    var duplicateClass = normalizeClassName(duplicateSettings[duplicateIndex][0]);
    var duplicateId = canonicalAssignmentId(duplicateSettings[duplicateIndex][1]);
    var sameAssignment = duplicateClass === normalizeClassName(className) && duplicateId === assignmentId;
    var isOriginalEditRow = data.isEdit && duplicateClass === normalizeClassName(className) && duplicateId === originalAssignmentId;
    if (sameAssignment && duplicateRowNumber !== foundRow && !isOriginalEditRow) return { success: false, message: 'رقم الواجب موجود مسبقاً لهذا الصف، اختر رقماً آخر' };
  }
  if (foundRow === -1) {
    if (data.isEdit) return { success: false, message: 'لم يتم العثور على الواجب الأصلي للتعديل' };
    var allSettings = settingsSheet.getRange(2, 1, Math.max(settingsSheet.getLastRow() - 1, 1), 2).getDisplayValues();
    for (var du = 0; du < allSettings.length; du++) if (normalizeClassName(allSettings[du][0]) === normalizeClassName(className) && canonicalAssignmentId(allSettings[du][1]) === assignmentId) return { success: false, message: 'رقم الواجب موجود مسبقاً لهذا الصف' };
    settingsSheet.appendRow([className, "'" + assignmentId, title, "غير منشور", "", ""]);
    foundRow = settingsSheet.getLastRow();
  } else {
    settingsSheet.getRange(foundRow, 1, 1, 3).setValues([[className, "'" + assignmentId, title]]);
  }
  
  // 2. تحديث بنك الأسئلة (حذف القديم لهذا الواجب وإضافة الجديد)
  var qSheet = ss.getSheetByName(questionsSheetName_(className));
  if (!qSheet) {
    qSheet = ss.insertSheet(questionsSheetName_(className));
    qSheet.appendRow(["رقم الواجب", "رقم السؤال", "نص السؤال", "خيار 1", "خيار 2", "خيار 3", "خيار 4", "الإجابة الصحيحة", "الدرجة", "تخطيط الخيارات"]);
  }
  
  var qLastRow = qSheet.getLastRow();
  if (qLastRow >= 2) {
    var qValues = qSheet.getRange(2, 1, qLastRow - 1, 1).getDisplayValues();
    for (var r = qValues.length; r >= 1; r--) {
      if (canonicalAssignmentId(qValues[r - 1][0]) === originalAssignmentId || canonicalAssignmentId(qValues[r - 1][0]) === assignmentId) {
        qSheet.deleteRow(r + 1);
      }
    }
  }
  
  // إدخال الأسئلة الجديدة
  questions.forEach(function(q, index) {
    qSheet.appendRow([
      "'" + assignmentId,
      index + 1,
      q.text,
      q.choices[0] || "",
      q.choices[1] || "",
      q.choices[2] || "",
      q.choices[3] || "",
      q.correct,
      q.score,
      q.choiceLayout || "column",
      q.imageUrl || "", q.choiceImages && q.choiceImages[0] || "", q.choiceImages && q.choiceImages[1] || "", q.choiceImages && q.choiceImages[2] || "", q.choiceImages && q.choiceImages[3] || ""
    ]);
  });
  
  // 3. إضافة عمود الدرجات في جدول درجات الصف (يبدأ من العمود D فصاعداً)
  var gradeSheet = findClassSheet(ss, className);
  if (gradeSheet) {
    var lastCol = gradeSheet.getLastColumn();
    var colFound = false;
    if (lastCol >= 4) {
      var headers = gradeSheet.getRange(1, 4, 1, lastCol - 3).getDisplayValues()[0];
      for (var j = 0; j < headers.length; j++) {
        if (canonicalAssignmentId(headers[j]) === originalAssignmentId || canonicalAssignmentId(headers[j]) === assignmentId) {
          if (canonicalAssignmentId(headers[j]) === originalAssignmentId && originalAssignmentId !== assignmentId) gradeSheet.getRange(1, j + 4).setValue(assignmentId);
          colFound = true;
          break;
        }
      }
    }
    if (!colFound) {
      var targetCol = lastCol < 4 ? 4 : lastCol + 1;
      gradeSheet.getRange(1, targetCol).setValue(assignmentId);
    }
  }
  
  clearPublishedAssignmentsCache_(className);
  return { success: true, message: "✅ تم حفظ الواجب وأسئلته بنجاح!" };
}

function checkAssignmentDuplicate(className, assignmentId,token) { requireTeacherSession_(token);
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow()<2) return {exists:false};
  var rows=sheet.getRange(2,1,sheet.getLastRow()-1,2).getDisplayValues(), cls=normalizeClassName(className), id=canonicalAssignmentId(assignmentId);
  for (var i=0;i<rows.length;i++) if (normalizeClassName(rows[i][0])===cls && canonicalAssignmentId(rows[i][1])===id) return {exists:true,message:'هذا الواجب موجود مسبقاً لهذا الصف'};
  return {exists:false};
}

// بيانات التقارير الشاملة للمعلم
function getAllAssignmentsForTeacher(token) { requireTeacherSession_(token);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  var assignments = rows.filter(function(r) { return String(r[0]).trim() && String(r[1]).trim(); }).map(function(r) {
    var cls=String(r[0]).trim(), id=displayAssignmentId(r[1]), solved=0, total=0, activityCount=0, gradeSheet=findClassSheet(ss,cls);
    if (gradeSheet && gradeSheet.getLastRow() >= 2) {
      total=gradeSheet.getRange(2,2,gradeSheet.getLastRow()-1,1).getDisplayValues().filter(function(x){return String(x[0]).trim();}).length;
      var headers=gradeSheet.getRange(1,4,1,Math.max(gradeSheet.getLastColumn()-3,1)).getDisplayValues()[0], col=-1;
      for(var hi=0;hi<headers.length;hi++) if(canonicalAssignmentId(headers[hi])===canonicalAssignmentId(id)){col=hi+4;break;}
      if(col>0) solved=gradeSheet.getRange(2,col,Math.max(gradeSheet.getLastRow()-1,1),1).getDisplayValues().filter(function(x){return String(x[0]).trim();}).length;
    }
    var activitySheet=ss.getSheetByName('تسليمات_الأنشطة'); if(activitySheet&&activitySheet.getLastRow()>=2){var acRows=activitySheet.getRange(2,1,activitySheet.getLastRow()-1,Math.max(activitySheet.getLastColumn(),9)).getValues();acRows.forEach(function(ar){var score=activityScoreValue_(ar[7]);if(normalizeClassName(ar[1])===normalizeClassName(cls)&&canonicalAssignmentId(ar[3])===canonicalAssignmentId(id)&&!(score!==''&&Number(score)===0))activityCount++;});}
    return { className: cls, id: id, title: String(r[2] || ''), status: String(r[3] || '').trim() || 'غير منشور', startDate: r[4] || '', endDate: r[5] || '', solvedCount: solved, totalStudents: total, activityCount: activityCount };
  });
  assignments.sort(function(a,b){var pa=String(a.id||'').match(/^(\d+)-(\d+)/),pb=String(b.id||'').match(/^(\d+)-(\d+)/);if(pa&&pb)return Number(pb[1])-Number(pa[1])||Number(pb[2])-Number(pa[2]);if(pa&&!pb)return -1;if(!pa&&pb)return 1;return String(b.id||'').localeCompare(String(a.id||''));});
  return assignments;
}
// نشر الواجب مع أوقات البدء والانتهاء
function publishAssignmentData(className, assignmentId, startDate, endDate,token) { requireTeacherSession_(token);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'لا يوجد جدول إعدادات للواجبات' };
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  var targetClass = normalizeClassName(className);
  var targetId = displayAssignmentId(assignmentId);
  for (var i = 0; i < rows.length; i++) {
    if (normalizeClassName(rows[i][0]) === targetClass && canonicalAssignmentId(rows[i][1]) === canonicalAssignmentId(targetId)) {
      // إذا تُرك وقت البداية فارغًا أثناء التمديد، يبدأ النشر من هذه اللحظة.
      var effectiveStart = String(startDate || '').trim() ? startDate : new Date();
      sheet.getRange(i + 2, 4, 1, 3).setValues([['منشور', effectiveStart, endDate]]);
      clearPublishedAssignmentsCache_(targetClass);
      return { success: true, message: '🚀 تم نشر الواجب بنجاح وإتاحته للطلاب!' };
    }
  }
  return { success: false, message: 'لم يتم العثور على الواجب: ' + targetId };
}
// جلب الواجبات المتاحة والمنشورة للطالب مع حالة الحل
function changeStudentPin(className, studentName, oldPin, newPin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = findClassSheet(ss, className);
  if (!sheet || sheet.getLastRow() < 2) return {success:false, message:'لم يتم العثور على ورقة الصف'};
  var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getDisplayValues();
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][1]).trim() === String(studentName).trim()) {
      if (String(rows[i][0]).trim() !== String(oldPin).trim()) return {success:false, message:'الرقم الحالي غير صحيح'};
      if (!String(newPin || '').trim() || String(newPin).trim().length < 1) return {success:false, message:'اكتب كلمة سر جديدة'};
      sheet.getRange(i+2, 1).setNumberFormat('@').setValue(String(newPin));
      return {success:true, message:'تم تغيير الرقم السري بنجاح'};
    }
  }
  return {success:false, message:'لم يتم العثور على الطالب'};
}

function studentActivityNameKey_(value){return String(value||'').replace(/[\s\u200f\u200eـ]+/g,'').trim();}

// ==================== Activity matrix storage v2 ====================
const ACTIVITY_SHEET_SIXTH = 'تسليم أنشطة سادس';
const ACTIVITY_SHEET_FIRST = 'تسليم أنشطة أول متوسط';
function activitySheetName_(className){ return activityClassKey_(className)==='أول متوسط' ? ACTIVITY_SHEET_FIRST : ACTIVITY_SHEET_SIXTH; }
function activityMatrixHeaders_(className){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), settings=ss.getSheetByName(SHEET_SETTINGS), ids=[];
  if(settings&&settings.getLastRow()>=2) settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues().forEach(function(r){
    if(activityClassKey_(r[0])===activityClassKey_(className)){var id=canonicalAssignmentId(r[1]);if(/^\d+-\d+$/.test(id)&&ids.indexOf(id)<0)ids.push(id);}
  });
  ids.sort(function(a,b){var x=a.split('-'),y=b.split('-');return Number(x[0])-Number(y[0])||Number(x[1])-Number(y[1]);});
  var h=['اسم الطالب','الصف','الفصل']; ids.forEach(function(id){h.push(id+' رابط',id+' درجة',id+' حالة',id+' ملاحظات');}); return h;
}
function ensureActivityMatrixSheet_(className){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),name=activitySheetName_(className),sh=ss.getSheetByName(name),wanted=activityMatrixHeaders_(className);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()<1){sh.getRange(1,1,1,wanted.length).setValues([wanted]);}
  else {var current=sh.getRange(1,1,1,Math.max(4,sh.getLastColumn())).getDisplayValues()[0],missing=wanted.filter(function(x){return current.indexOf(x)<0;});if(missing.length)sh.getRange(1,sh.getLastColumn()+1,1,missing.length).setValues([missing]);}
  sh.setFrozenRows(1);sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold');return sh;
}
function activityMatrixCols_(sh){var h=sh.getRange(1,1,1,Math.max(4,sh.getLastColumn())).getDisplayValues()[0],o={};h.forEach(function(x,i){o[String(x).trim()]=i+1;});return o;}
function normalizeSectionValue_(section){var s=String(section==null?'':section).trim().replace(',','.');if(/^\d+(?:\.0+)?$/.test(s))return String(Number(s));return s;}
function activityMatrixRowKey_(name,cls,section){return String(name||'').trim()+'|'+activityClassKey_(cls)+'|'+normalizeSectionValue_(section);}
function activityMatrixRows_(sh){var last=Math.max(4,sh.getLastColumn()),out=[];if(sh.getLastRow()>=2)sh.getRange(2,1,sh.getLastRow()-1,last).getValues().forEach(function(r,i){if(String(r[0]||'').trim())out.push({row:i+2,values:r});});return out;}
function activityMatrixMap_(className){
  var sh=ensureActivityMatrixSheet_(className),cols=activityMatrixCols_(sh),out={};activityMatrixRows_(sh).forEach(function(x){var r=x.values,n=String(r[0]||'').trim(),c=activityClassKey_(r[1]),sec=String(r[2]||'').trim();if(!n)return;var key=activityMatrixRowKey_(n,c,sec),v={row:x.row,name:n,className:c,section:sec};Object.keys(cols).forEach(function(h){var m=String(h).match(/^(\d+-\d+) (رابط|درجة|حالة|ملاحظات)$/);if(m){v[m[1]+'_'+m[2]]=r[cols[h]-1]||'';}});out[key]=v;});return {sheet:sh,cols:cols,map:out};}
function activityMatrixFindValue_(d,className,name,section){var exact=d.map[activityMatrixRowKey_(name,className,section)]||d.map[activityMatrixRowKey_(name,className,'')];if(exact)return exact;var target=String(name||'').trim(),ck=activityClassKey_(className),found=null;Object.keys(d.map).some(function(k){var v=d.map[k];if(String(v.name||'').trim()===target&&activityClassKey_(v.className)===ck){found=v;return true;}return false;});return found||{};}
function activityMatrixGet_(className,name,section,id){var d=activityMatrixMap_(className),v=activityMatrixFindValue_(d,className,name,section);return {data:d,value:v};}
function activityMatrixFindOrCreateRow_(d,className,name,section){var key=activityMatrixRowKey_(name,className,section),v=d.map[key];if(v)return v;var sh=d.sheet,row=sh.getLastRow()+1,arr=[];for(var i=1;i<=sh.getLastColumn();i++)arr.push('');arr[0]=String(name).trim();arr[1]=activityClassKey_(className);arr[2]=String(section||'').trim();sh.getRange(row,1,1,arr.length).setValues([arr]);v={row:row,name:arr[0],className:arr[1],section:arr[2]};d.map[key]=v;return v;}
function activityWriteStatus_(){}
function activityMatrixStatus_(hasUrl,sc){return sc!==''&&Number(sc)>0?'تم التقييم':hasUrl?'تم الإرسال':'لم يرسل';}

function getStudentAssignmentsLegacy_(className,studentName,requestedSection){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),cls=normalizeClassName(className),name=String(studentName||'').trim(),wanted=String(requestedSection||'').trim(),settings=ss.getSheetByName(SHEET_SETTINGS),solved={},section='';
  var grade=findClassSheet(ss,cls);if(grade&&grade.getLastRow()>=2){var nr=grade.getRange(2,2,grade.getLastRow()-1,2).getDisplayValues();for(var i=0;i<nr.length;i++)if(String(nr[i][0]).trim()===name&&(!wanted||wanted==='__ALL__'||String(nr[i][1]).trim()===wanted)){section=String(nr[i][1]).trim();if(grade.getLastColumn()>=4){var hh=grade.getRange(1,4,1,grade.getLastColumn()-3).getDisplayValues()[0],vv=grade.getRange(i+2,4,1,grade.getLastColumn()-3).getDisplayValues()[0];hh.forEach(function(h,j){if(String(vv[j]).trim()!=='')solved[canonicalAssignmentId(h)]=vv[j];});}break;}}
  var checks=getCheckGradesForStudent_(cls,section,name),d=activityMatrixMap_(cls),now=Date.now();
  return publishedAssignmentRowsCached_(cls).filter(function(r){return activityClassKey_(r[0])===activityClassKey_(cls);}).map(function(r){var id=canonicalAssignmentId(r[1]),st=r[4]?new Date(r[4]).getTime():0,en=r[5]?new Date(r[5]).getTime():0,x=activityMatrixFindValue_(d,cls,name,section),url=normalizeDriveImageUrl_(x[id+'_رابط']||''),sc=activityScoreValue_(x[id+'_درجة']);return {id:id,title:r[2],startDate:r[4]||'',endDate:r[5]||'',isSolved:Object.prototype.hasOwnProperty.call(solved,id),isExpired:!!(en&&en<now),isNotStarted:!!(st&&st>now),score:solved[id]||null,isActivitySubmitted:(!!url||sc!==''),isActivityGraded:sc!==''&&Number(sc)>0,activityUrl:url,activityNote:x[id+'_ملاحظات']||'',activityStatus:x[id+'_حالة']||activityMatrixStatus_(!!url,sc),checkScore:checks[id]||null};});
}
function secureFolderTree_(folder){var files=folder.getFiles();while(files.hasNext()){try{files.next().setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.VIEW);}catch(e){}}var folders=folder.getFolders();while(folders.hasNext())secureFolderTree_(folders.next());}
function secureExistingPlatformFiles(){['منصة الواجبات','صور منصة الواجبات','صور أسئلة المنصة','صور خيارات المنصة','صور أنشطة الطلاب'].forEach(function(name){var it=DriveApp.getFoldersByName(name);while(it.hasNext())secureFolderTree_(it.next());});return 'تم جعل ملفات منصة الواجبات خاصة بالكامل';}
function authorizePlatformDrive() {
  // إنشاء مجلد تجريبي يطلب صلاحية Drive الكاملة (قراءة وكتابة)، ثم حذفه.
  var probe=DriveApp.createFolder('منصة الواجبات - اختبار الصلاحية');
  var id=probe.getId();
  probe.setTrashed(true);
  var imageRoot=getOrCreateFolder_('منصة الواجبات'); ['سادس','أول متوسط'].forEach(function(c){var cf=getOrCreateChildFolder_(imageRoot,c);getOrCreateChildFolder_(cf,'صور الأسئلة');getOrCreateChildFolder_(cf,'صور الخيارات');getOrCreateChildFolder_(cf,'صور الأنشطة');}); secureExistingPlatformFiles();
  return 'تم تفويض Google Drive بصلاحية القراءة والكتابة بنجاح: '+id;
}
function getOrCreateFolder_(name) {
  var it=DriveApp.getFoldersByName(name); return it.hasNext()?it.next():DriveApp.createFolder(name);
}
function normalizeDriveImageUrl_(url) {
  var value=String(url||'').trim(); if(!value)return '';
  var m=value.match(/[-\w]{20,}/); return m ? 'https://drive.google.com/uc?export=view&id='+m[0] : value;
}
function getOrCreateChildFolder_(parent,name) {
  var it=parent.getFoldersByName(String(name)); return it.hasNext()?it.next():parent.createFolder(String(name));
}
function getClassImageFolder_(className,kind) { var root=getOrCreateFolder_('منصة الواجبات'), cls=getOrCreateChildFolder_(root,normalizeClassName(className)); return getOrCreateChildFolder_(cls,kind); }
function saveImageToDrive_(dataUrl,fileName,folderName) {
  var parts=String(dataUrl||'').split(','), meta=parts[0]||'', raw=parts[1]||'';
  if(!raw) throw new Error('بيانات الصورة غير موجودة');
  var mime=(meta.match(/data:([^;]+);/)||[])[1]||'image/jpeg';
  var bytes=Utilities.base64Decode(raw), blob=Utilities.newBlob(bytes,mime,fileName), parent=folderName==='صور الأسئلة'||folderName==='صور الخيارات'?getOrCreateChildFolder_(getOrCreateFolder_('صور منصة الواجبات'),folderName):getOrCreateFolder_(folderName), file=parent.createFile(blob);
  
  return normalizeDriveImageUrl_(file.getUrl());
}
function getDriveImageDataUrl(fileUrl) {
  var m=String(fileUrl||'').match(/[-\w]{20,}/); if(!m) throw new Error('رابط الصورة غير صالح');
  var blob=DriveApp.getFileById(m[0]).getBlob(), bytes=blob.getBytes();
  var encoded=Utilities.base64Encode(bytes); return 'data:'+blob.getContentType()+';base64,'+encoded;
}
function uploadQuestionImage(dataUrl,className,assignmentId,questionNumber,choiceIndex,token) { requireTeacherSession_(token);
  var folder=getClassImageFolder_(className,choiceIndex===null||choiceIndex===undefined?'صور الأسئلة':'صور الخيارات');
  var name='واجب '+canonicalAssignmentId(assignmentId)+' - سؤال '+questionNumber+(choiceIndex===null||choiceIndex===undefined?'':' - خيار '+(Number(choiceIndex)+1))+'.jpg';
  var parts=String(dataUrl||'').split(','),mime=((parts[0]||'').match(/data:([^;]+);/)||[])[1]||'image/jpeg',file=folder.createFile(Utilities.newBlob(Utilities.base64Decode(parts[1]||''),mime,name)); return {success:true,url:normalizeDriveImageUrl_(file.getUrl())};
}
function activityScoreValue_(value){
  if(Object.prototype.toString.call(value)==='[object Date]'&&!isNaN(value.getTime())){
    var base=new Date(1899,11,30); return Math.round((value.getTime()-base.getTime())/86400000*100)/100;
  }
  var text=String(value==null?'':value).trim().replace(/[،,]/g,'.');
  if(text==='')return '';
  var n=Number(text); return isNaN(n)?'':n;
}
function ensureActivityColumns_(sh){
  if(!sh)return;
  if(sh.getLastColumn()<9)sh.insertColumnsAfter(Math.max(sh.getLastColumn(),1),9-Math.max(sh.getLastColumn(),1));
  var headers=sh.getRange(1,1,1,9).getDisplayValues()[0];
  if(!String(headers[7]||'').trim())sh.getRange(1,8).setValue('الدرجة');
  if(!String(headers[8]||'').trim())sh.getRange(1,9).setValue('ملاحظات');
  if(sh.getLastRow()>=2){
    sh.getRange(2,8,sh.getLastRow()-1,1).setNumberFormat('0.##');
    var rows=sh.getRange(2,1,sh.getLastRow()-1,Math.max(sh.getLastColumn(),9)).getValues();
    for(var i=0;i<rows.length;i++){
      var hasSubmission=String(rows[i][0]||'').trim()&&String(rows[i][4]||'').trim();
      var score=rows[i][7];
      // الدرجة تُدخل يدويًا من تقرير الأنشطة؛ لا تمنح 10 تلقائيًا عند رفع الملف.
    }
  }
}
function uploadStudentActivityLegacy_(dataUrl,className,section,studentName,assignmentId){
  var id=canonicalAssignmentId(assignmentId),d=activityMatrixMap_(className),existing=activityMatrixGet_(className,studentName,section,id).value,oldScore=activityScoreValue_(existing[id+'_درجة']);
  if(oldScore!==''&&Number(oldScore)>0)return {success:false,message:'تم تثبيت درجة النشاط من المعلم ولا يمكن إعادة الإرسال'};
  var root=getClassImageFolder_(className,'صور الأنشطة'),folder=getOrCreateChildFolder_(root,id),safe=String(studentName).replace(/[\\/:*?"<>|]/g,'_'),parts=String(dataUrl||'').split(','),mime=((parts[0]||'').match(/data:([^;]+);/)||[])[1]||'image/jpeg',file=folder.createFile(Utilities.newBlob(Utilities.base64Decode(parts[1]||''),mime,safe+' - واجب '+id+'.jpg')),row=activityMatrixFindOrCreateRow_(d,className,studentName,section),cols=d.cols,sh=d.sheet;
  if(cols[id+' رابط'])sh.getRange(row.row,cols[id+' رابط']).setValue(file.getUrl());if(cols[id+' درجة'])sh.getRange(row.row,cols[id+' درجة']).clearContent();if(cols[id+' حالة'])sh.getRange(row.row,cols[id+' حالة']).setValue('تم الإرسال');if(cols[id+' ملاحظات'])sh.getRange(row.row,cols[id+' ملاحظات']).clearContent();sh.getRange(row.row,4).setValue(new Date());
  return {success:true,url:file.getUrl(),message:'تم إرسال النشاط بنجاح'};
}
function activityClassKey_(name){
  var value=normalizeClassName(name).replace(/[\s\u200f\u200e]+/g,'');
  if(value==='سادس'||value==='السادس'||value.indexOf('سادس')>=0||value.indexOf('السادس')>=0)return 'سادس';
  if(value==='أولمتوسط'||value==='اولمتوسط'||value==='أولىمتوسط'||value==='اولىمتوسط'||value.indexOf('متوسط')>=0)return 'أول متوسط';
  return normalizeClassName(name);
}

function findOldActivityFileUrl_(className,assignmentId,studentName){
  try{
    var rootIt=DriveApp.getFoldersByName('منصة الواجبات'),id=canonicalAssignmentId(assignmentId),student=String(studentName||'').trim();
    while(rootIt.hasNext()){
      var root=rootIt.next(),classNames=[String(className||'').trim(),'سادس','الصف السادس','أول متوسط','أولى متوسط','اولى متوسط'],seen={};
      for(var ci=0;ci<classNames.length;ci++){
        var classIt=root.getFoldersByName(classNames[ci]);
        while(classIt.hasNext()){
          var classFolder=classIt.next(),activityIt=classFolder.getFoldersByName('صور الأنشطة');
          while(activityIt.hasNext()){
            var activityFolder=activityIt.next(),lessonIt=activityFolder.getFoldersByName(id);
            while(lessonIt.hasNext()){
              var files=lessonIt.next().getFiles();
              while(files.hasNext()){
                var file=files.next(),name=String(file.getName()||'');
                if(name.indexOf(student)>=0&&(name.indexOf(id)>=0||name.indexOf('واجب')>=0))return normalizeDriveImageUrl_(file.getUrl());
              }
            }
          }
        }
      }
    }
  }catch(e){}
  return '';
}

function findOldActivityFilesMap_(className,assignmentId){
  var out={};
  try{
    var rootIt=DriveApp.getFoldersByName('منصة الواجبات'),id=canonicalAssignmentId(assignmentId);
    while(rootIt.hasNext()){
      var root=rootIt.next(),classNames=[String(className||'').trim(),'سادس','الصف السادس','أول متوسط','أولى متوسط','اولى متوسط'],seen={};
      for(var ci=0;ci<classNames.length;ci++){
        var classIt=root.getFoldersByName(classNames[ci]);
        while(classIt.hasNext()){
          var classFolder=classIt.next(),activityIt=classFolder.getFoldersByName('صور الأنشطة');
          while(activityIt.hasNext()){
            var activityFolder=activityIt.next(),lessonIt=activityFolder.getFoldersByName(id);
            while(lessonIt.hasNext()){
              var files=lessonIt.next().getFiles();
              while(files.hasNext()){
                var file=files.next(),fileName=String(file.getName()||''),student=fileName.split(' - واجب')[0].trim();
                if(student)out[student]=normalizeDriveImageUrl_(file.getUrl());
              }
            }
          }
        }
      }
    }
  }catch(e){}
  return out;
}
function getActivityManagementLegacy_(className,section,unit,token){requireTeacherSession_(token);var ss=SpreadsheetApp.getActiveSpreadsheet(),settings=ss.getSheetByName(SHEET_SETTINGS),assignments=[],u=String(unit||'').trim();if(settings&&settings.getLastRow()>=2)settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues().forEach(function(r){var id=canonicalAssignmentId(r[1]);if(activityClassKey_(r[0])===activityClassKey_(className)&&String(r[3]).trim()==='منشور'&&String(id).indexOf(u+'-')===0)assignments.push({id:id,title:r[2]||''});});var cls=findClassSheet(ss,className),students=[];if(cls&&cls.getLastRow()>=2)cls.getRange(2,1,cls.getLastRow()-1,3).getDisplayValues().forEach(function(r){if(String(r[1]).trim()&&(section==='__ALL__'||normalizeSectionValue_(r[2])===normalizeSectionValue_(section)))students.push({name:String(r[1]).trim(),section:String(r[2]||'').trim()});});var d=activityMatrixMap_(className);students=sortStudentRows_(students,section==='__ALL__');return {assignments:assignments,students:students,rows:students.map(function(st){var x=activityMatrixFindValue_(d,className,st.name,st.section);return {name:st.name,section:st.section,activities:assignments.map(function(a){var sc=activityScoreValue_(x[a.id+'_درجة']),url=normalizeDriveImageUrl_(x[a.id+'_رابط']||'');return {id:a.id,url:url,score:sc,note:x[a.id+'_ملاحظات']||'',status:x[a.id+'_حالة']||activityMatrixStatus_(!!url,sc)};})};})};}
function saveActivityManagementLegacy_(className,section,rows,token){requireTeacherSession_(token);var d=activityMatrixMap_(className),sh=d.sheet,cols=d.cols;(rows||[]).forEach(function(st){var row=activityMatrixFindOrCreateRow_(d,className,st.name,st.section||section);(st.activities||[]).forEach(function(a){var id=canonicalAssignmentId(a.id),raw=String(a.score==null?'':a.score).trim();if(raw!==''&&!/^\d+(?:[.,]\d+)?$/.test(raw))return;var sc=raw===''?'':Number(raw.replace(',','.')),url=String(a.url||'').trim(),base=row.row;if(cols[id+' درجة']){if(sc==='')sh.getRange(base,cols[id+' درجة']).clearContent();else sh.getRange(base,cols[id+' درجة']).setValue(sc);}if(cols[id+' رابط']&&url)sh.getRange(base,cols[id+' رابط']).setValue(url);if(cols[id+' حالة'])sh.getRange(base,cols[id+' حالة']).setValue(activityMatrixStatus_(!!(url||sh.getRange(base,cols[id+' رابط']).getDisplayValue()),sc));if(cols[id+' ملاحظات'])sh.getRange(base,cols[id+' ملاحظات']).setValue(String(a.note||''));});});return {success:true,message:'تم حفظ درجات الأنشطة وتحديث الحالة في نفس صف الطالب'};}
function getAssignmentQuestions(className, assignmentId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(questionsSheetName_(className));
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), 15)).getDisplayValues();
  var target = canonicalAssignmentId(assignmentId);
  return rows.filter(function(r) { return canonicalAssignmentId(r[0]) === target; }).map(function(r) {
    return { questionNumber: r[1], text: r[2], choices: [r[3], r[4], r[5], r[6]].filter(function(c) { return String(c).trim() !== ''; }), correct: String(r[7] || '').trim(), score: Number(r[8]) || 1, choiceLayout: r[9] === 'grid' ? 'grid' : 'column', imageUrl:normalizeDriveImageUrl_(r[10]), choiceImages:[normalizeDriveImageUrl_(r[11]),normalizeDriveImageUrl_(r[12]),normalizeDriveImageUrl_(r[13]),normalizeDriveImageUrl_(r[14])] };
  });
}
function exportActivitiesExcel(className,token){return exportHomeworkReportExcel(className,'activities',token);}
function exportComprehensiveExcel(className,token){return exportHomeworkReportExcel(className,'homework',token);}

// حفظ نتيجة الطالب في الشيت
function submitStudentScore(studentName, className, assignmentId, totalScore, answersJson) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findClassSheet(ss, className);
  if (!sheet) return { success: false, message: 'لم يتم العثور على ورقة الصف' };
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  var target = canonicalAssignmentId(assignmentId), colIndex = -1;
  if (lastCol >= 4) {
    var headers = sheet.getRange(1, 4, 1, lastCol - 3).getDisplayValues()[0];
    for (var i = 0; i < headers.length; i++) if (canonicalAssignmentId(headers[i]) === target) { colIndex = i + 4; break; }
  }
  if (colIndex === -1) return { success: false, message: 'لم يتم العثور على عمود درجة الواجب ' + target };
  var studentRow = -1;
  if (lastRow >= 2) {
    var names = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
    for (var j = 0; j < names.length; j++) if (String(names[j][0]).trim() === String(studentName).trim()) { studentRow = j + 2; break; }
  }
  if (studentRow === -1) return { success: false, message: 'لم يتم العثور على الطالب في ورقة الصف' };
  sheet.getRange(studentRow, colIndex).setValue(totalScore);
  if (answersJson) PropertiesService.getScriptProperties().setProperty('answers_' + normalizeClassName(className) + '_' + studentName + '_' + target, String(answersJson));
  return { success: true, message: 'تم حفظ الدرجة' };
}



// ==================== الإنجازات ====================
const SHEET_ACHIEVEMENTS_SIXTH = "إنجازات سادس";
const SHEET_ACHIEVEMENTS_FIRST = "إنجازات أول متوسط";
const SHEET_ACHIEVEMENT_META = "إعدادات الإنجازات";

function ensureAchievementMetaSheet_(){var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(SHEET_ACHIEVEMENT_META);if(!sh)sh=ss.insertSheet(SHEET_ACHIEVEMENT_META);if(sh.getLastRow()<1)sh.getRange(1,1,1,6).setValues([['الصف','اسم الإنجاز','الوحدة','الدرجة الكاملة','النوع','تضاف إلى']]);else{if(sh.getLastColumn()<5)sh.getRange(1,5).setValue('النوع');if(sh.getLastColumn()<6)sh.getRange(1,6).setValue('تضاف إلى');}return sh;}
function getAchievementMeta_(className,name){var sh=ensureAchievementMetaSheet_();if(sh.getLastRow()<2)return {unit:'',max:'10',type:'video',target:'achievement'};var r=sh.getRange(2,1,sh.getLastRow()-1,Math.max(5,sh.getLastColumn())).getDisplayValues();for(var i=0;i<r.length;i++)if(normalizeClassName(r[i][0])===normalizeClassName(className)&&String(r[i][1]).trim()===String(name).trim())return {unit:String(r[i][2]||''),max:String(r[i][3]||'10'),type:String(r[i][4]||'video'),target:String(r[i][5]||'achievement').trim()==='test'?'test':'achievement'};return {unit:'',max:'10',type:'video',target:'achievement'};}
function achievementsSheetName_(className){return normalizeClassName(className)==='أول متوسط'?SHEET_ACHIEVEMENTS_FIRST:SHEET_ACHIEVEMENTS_SIXTH;}
function ensureAchievementsSheet_(className){var ss=SpreadsheetApp.getActiveSpreadsheet(),name=achievementsSheetName_(className),sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()<1||sh.getLastColumn()<3)sh.getRange(1,1,1,3).setValues([['اسم الطالب','الصف','الفصل']]);else{var h=sh.getRange(1,1,1,Math.max(3,sh.getLastColumn())).getDisplayValues()[0];if(h[0]!=='اسم الطالب')sh.getRange(1,1,1,3).setValues([['اسم الطالب','الصف','الفصل']]);}return sh;}
function achievementCell_(url,score){return JSON.stringify({url:String(url||''),score:score==null?'':String(score),updatedAt:new Date().toISOString()});}
function parseAchievementCell_(value){var raw=String(value==null?'':value).trim();if(!raw)return {url:'',score:''};try{var o=JSON.parse(raw);return {url:String(o.url||''),score:o.score==null?'':String(o.score)};}catch(e){return {url:raw,score:''};}}
function achievementColumn_(sh,name,create){var target=String(name||'').trim();if(!target)return -1;var last=Math.max(3,sh.getLastColumn()),h=sh.getRange(1,1,1,last).getDisplayValues()[0];for(var i=3;i<h.length;i++)if(String(h[i]).trim()===target)return i+1;if(create){var c=Math.max(4,sh.getLastColumn()+1);sh.getRange(1,c).setValue(target).setWrap(true);return c;}return -1;}
function getAchievementDefinitions(className,token){requireTeacherSession_(token);var sh=ensureAchievementsSheet_(className),h=sh.getRange(1,1,1,Math.max(3,sh.getLastColumn())).getDisplayValues()[0];return h.slice(3).filter(function(x){return String(x).trim();}).map(function(x){var m=getAchievementMeta_(className,String(x).trim());return {name:String(x).trim(),unit:m.unit,max:m.max,type:m.type,target:m.target};});}
function saveAchievementDefinition(className,name,unit,maxScore,target,token){requireTeacherSession_(token);var n=String(name||'').trim(),cls=normalizeClassName(className),u=String(unit||'').trim(),mx=String(maxScore||'10').trim(),tg=String(target||'achievement').trim()==='test'?'test':'achievement';if(!cls||!n||!/^[0-9]+$/.test(u)||!/^[0-9]+(?:[.,]\d+)?$/.test(mx))return {success:false,message:'اختر الصف والوحدة واكتب اسم الإنجاز والدرجة الكاملة'};var sh=ensureAchievementsSheet_(cls);if(achievementColumn_(sh,n,false)>0)return {success:false,message:'هذا الإنجاز موجود مسبقاً في هذا الصف'};achievementColumn_(sh,n,true);var ms=ensureAchievementMetaSheet_();ms.appendRow([cls,n,u,Number(mx.replace(',','.')),'video',tg]);return {success:true,message:'تم حفظ الإنجاز للصف '+cls+' في الوحدة '+u};}
function saveContestDefinition(className,name,unit,maxScore,rows,token){requireTeacherSession_(token);var n=String(name||'').trim(),cls=normalizeClassName(className),u=String(unit||'').trim(),mx=String(maxScore||'10').trim();if(!cls||!n||!/^[0-9]+$/.test(u)||!/^[0-9]+(?:[.,][0-9]+)?$/.test(mx))return {success:false,message:'بيانات المسابقة غير مكتملة'};var sh=ensureAchievementsSheet_(cls);if(achievementColumn_(sh,n,false)>0)return {success:false,message:'اسم المسابقة موجود مسبقاً لهذا الصف'};var c=achievementColumn_(sh,n,true),ms=ensureAchievementMetaSheet_();ms.appendRow([cls,n,u,Number(mx.replace(',','.')),'contest']);var students=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues():[];(rows||[]).forEach(function(item){var name0=String(item.name||'').trim(),score=String(item.score==null?'':item.score).trim();if(!name0)return;var row=-1;for(var i=0;i<students.length;i++)if(String(students[i][0]).trim()===name0&&normalizeClassName(students[i][1])===cls){row=i+2;break;}if(row<0){row=sh.getLastRow()+1;sh.getRange(row,1,1,3).setValues([[name0,cls,'']]);students.push([name0,cls,'']);}sh.getRange(row,c).setValue(achievementCell_('',score)).setWrap(true);});return {success:true,message:'تمت إضافة المسابقة ودرجات الطلاب'};}
function deleteAchievementDefinition(className,name,token){requireTeacherSession_(token);var sh=ensureAchievementsSheet_(className),c=achievementColumn_(sh,name,false);if(c<4)return {success:false,message:'الإنجاز غير موجود'};sh.deleteColumn(c);return {success:true,message:'تم حذف الإنجاز وبياناته من ورقة الصف'};}
function getStudentAchievements(className,section,studentName){var sh=ensureAchievementsSheet_(className),last=Math.max(3,sh.getLastColumn()),h=sh.getRange(1,1,1,last).getDisplayValues()[0],rows=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues():[],wanted=String(studentName||'').trim(),cls=normalizeClassName(className),sec=String(section||'').trim(),candidates=[];for(var i=0;i<rows.length;i++)if(String(rows[i][0]).trim()===wanted&&normalizeClassName(rows[i][1])===cls){if(String(rows[i][2]).trim()===sec)candidates.unshift(i+2);else candidates.push(i+2);}var valsByRow={};candidates.forEach(function(r){valsByRow[r]=sh.getRange(r,1,1,last).getDisplayValues()[0];});return h.slice(3).map(function(n,j){var chosen={url:'',score:''};for(var k=0;k<candidates.length;k++){var cell=parseAchievementCell_(valsByRow[candidates[k]][j+3]);if(cell.url||cell.score!==''){chosen=cell;break;}}var meta=getAchievementMeta_(className,String(n||'').trim());return {name:String(n||'').trim(),url:chosen.url,score:chosen.score,unit:meta.unit,max:meta.max,type:meta.type,target:meta.target,addedToTest:meta.target==='test'&&Number(chosen.score||0)>0,submitted:meta.type==='contest'?String(chosen.score||'').trim()!=='' : !!chosen.url};}).filter(function(x){return x.name;});}
function getAchievementReport(className,achievementName,token){
  requireTeacherSession_(token);
  var cls=normalizeClassName(className),name=String(achievementName||'').trim();
  var rosterSheet=findClassSheet(SpreadsheetApp.getActiveSpreadsheet(),cls),students=[];
  if(rosterSheet&&rosterSheet.getLastRow()>=2){
    rosterSheet.getRange(2,1,rosterSheet.getLastRow()-1,3).getDisplayValues().forEach(function(row){
      var studentName=String(row[1]||'').trim();
      if(studentName)students.push({name:studentName,section:String(row[2]||'').trim()});
    });
  }
  var achievementSheet=ensureAchievementsSheet_(cls),column=achievementColumn_(achievementSheet,name,false),byStudent={},byName={};
  if(column>=4&&achievementSheet.getLastRow()>=2){
    achievementSheet.getRange(2,1,achievementSheet.getLastRow()-1,Math.max(column,3)).getDisplayValues().forEach(function(row){
      if(normalizeClassName(row[1])!==cls)return;
      var studentName=String(row[0]||'').trim(),section=String(row[2]||'').trim();
      if(!studentName)return;
      var value=parseAchievementCell_(row[column-1]);
      byStudent[studentName+'|'+section]=value;
      if(!byName[studentName])byName[studentName]=value;
    });
  }
  var meta=getAchievementMeta_(cls,name);
  return students.map(function(student){
    var exact=byStudent[student.name+'|'+student.section],value=(exact&&(exact.url||exact.score!==''))?exact:(byName[student.name]||exact||{url:'',score:''});
    return {name:student.name,section:student.section,url:value.url,score:value.score,unit:meta.unit,max:meta.max,type:meta.type};
  });
}

function saveStudentAchievement(className,section,studentName,achievementName,url){var n=String(achievementName||'').trim(),u=String(url||'').trim();if(!/^https?:\/\//i.test(u))return {success:false,message:'استخدم رابطًا يبدأ بـ https:// أو http://'};if(!n||!u)return {success:false,message:'أدخل رابط الإنجاز'};var sh=ensureAchievementsSheet_(className),c=achievementColumn_(sh,n,true),last=sh.getLastRow(),row=-1;if(last>=2){var rows=sh.getRange(2,1,last-1,3).getDisplayValues();for(var i=0;i<rows.length;i++)if(String(rows[i][0]).trim()===String(studentName||'').trim()&&normalizeClassName(rows[i][1])===normalizeClassName(className)&&String(rows[i][2]).trim()===String(section||'').trim()){row=i+2;break;}}if(row<0){row=sh.getLastRow()+1;sh.getRange(row,1,1,3).setValues([[studentName,normalizeClassName(className),section||'']]);}sh.getRange(row,c).setValue(achievementCell_(u,'')).setWrap(true);return {success:true,message:'تم حفظ رابط الإنجاز وإرساله للمعلم'};}
function saveAchievementGrades(className,achievementName,rows,token){requireTeacherSession_(token);var sh=ensureAchievementsSheet_(className),c=achievementColumn_(sh,achievementName,true),last=sh.getLastRow(),data=last>=2?sh.getRange(2,1,last-1,Math.max(3,c)).getDisplayValues():[];(rows||[]).forEach(function(item){var name=String(item.name||'').trim(),section=String(item.section==null?'':item.section).trim(),score=String(item.score==null?'':item.score).trim();if(!name)return;if(score!==''&&!/^\d+$/.test(score))return;if(score!==''&&(Number(score)<0||Number(score)>20))return;var row=-1,fallback=-1;for(var i=0;i<data.length;i++){if(String(data[i][0]).trim()!==name||normalizeClassName(data[i][1])!==normalizeClassName(className))continue;var rowSection=String(data[i][2]||'').trim();if(rowSection===section){row=i+2;break;}if(fallback<0)fallback=i+2;}if(row<0)row=fallback;if(row<0){row=sh.getLastRow()+1;sh.getRange(row,1,1,3).setValues([[name,normalizeClassName(className),section]]);data.push([name,normalizeClassName(className),section]);}var old=parseAchievementCell_(sh.getRange(row,c).getDisplayValue());sh.getRange(row,c).setValue(achievementCell_(old.url,score)).setWrap(true);});return {success:true,message:'تم حفظ درجات الإنجاز بنجاح'};}

function deleteAssignment(className, assignmentId,token) { requireTeacherSession_(token);
  var ss=SpreadsheetApp.getActiveSpreadsheet(), targetClass=normalizeClassName(className), targetId=canonicalAssignmentId(assignmentId), settings=ss.getSheetByName(SHEET_SETTINGS);
  if(!settings||settings.getLastRow()<2) return {success:false,message:'لا يوجد جدول إعدادات للواجبات'};
  var rows=settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues(), removed=false;
  for(var i=rows.length-1;i>=0;i--) if(normalizeClassName(rows[i][0])===targetClass&&canonicalAssignmentId(rows[i][1])===targetId){settings.deleteRow(i+2);removed=true;}
  if(!removed) return {success:false,message:'لم يتم العثور على الواجب المراد حذفه'};
  var qs=ss.getSheetByName(questionsSheetName_(className));
  if(qs&&qs.getLastRow()>=2){var q=qs.getRange(2,1,qs.getLastRow()-1,1).getDisplayValues();for(var j=q.length-1;j>=0;j--)if(canonicalAssignmentId(q[j][0])===targetId)qs.deleteRow(j+2);}
  var grade=findClassSheet(ss,className);
  if(grade&&grade.getLastColumn()>=4){var hs=grade.getRange(1,4,1,grade.getLastColumn()-3).getDisplayValues()[0];for(var c=hs.length-1;c>=0;c--)if(canonicalAssignmentId(hs[c])===targetId)grade.deleteColumn(c+4);}
  return {success:true,message:'تم حذف الواجب وأسئلته ودرجاته بالكامل'};
}


function getStudentReviewAnswers(studentName, className, assignmentId) {
  var key = 'answers_' + normalizeClassName(className) + '_' + studentName + '_' + canonicalAssignmentId(assignmentId);
  return PropertiesService.getScriptProperties().getProperty(key) || '{}';
}


function testSheetName_(className){return normalizeClassName(className)==='أول متوسط'?'اختبارات أول متوسط':'اختبارات سادس';}
function ensureTestSheet_(className){var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(testSheetName_(className));if(!sh)sh=ss.insertSheet(testSheetName_(className));if(sh.getLastRow()<1)sh.getRange(1,1,1,2).setValues([['اسم الطالب','الفصل']]);if(sh.getLastRow()<2){var roster=rosterForExport_(className);if(roster.length)sh.getRange(2,1,roster.length,2).setValues(roster.map(function(st){return [st.name,st.section];}));}return sh;}
function ensureTestColumn_(sh,unit){var name='اختبار'+String(Number(unit));var last=Math.max(2,sh.getLastColumn()),h=sh.getRange(1,1,1,last).getDisplayValues()[0];for(var i=2;i<h.length;i++)if(String(h[i]).trim()===name)return i+1;var c=Math.max(3,sh.getLastColumn()+1);sh.getRange(1,c).setValue(name);return c;}
function getTestManagement(className,section,unit,token){requireTeacherSession_(token);var ss=SpreadsheetApp.getActiveSpreadsheet(),cls=findClassSheet(ss,className),students=[];if(cls&&cls.getLastRow()>=2)cls.getRange(2,1,cls.getLastRow()-1,3).getDisplayValues().forEach(function(r){if(String(r[1]).trim()&&(section==='__ALL__'||normalizeSectionValue_(r[2])===normalizeSectionValue_(section)))students.push({name:String(r[1]).trim(),section:String(r[2]||'').trim()});});var sh=ensureTestSheet_(className),c=ensureTestColumn_(sh,unit),map={};if(sh.getLastRow()>=2){var rows=sh.getRange(2,1,sh.getLastRow()-1,Math.max(3,c)).getDisplayValues();rows.forEach(function(r){map[String(r[0]).trim()+'|'+String(r[1]).trim()]=r[c-1]||'';});}students=sortStudentRows_(students,section==='__ALL__');return {students:students,unit:Number(unit),rows:students.map(function(st){return {name:st.name,section:st.section,score:map[st.name+'|'+st.section]||''};})};}
function saveTestGrades(className,unit,rows,token){requireTeacherSession_(token);var sh=ensureTestSheet_(className),c=ensureTestColumn_(sh,unit),last=sh.getLastRow(),existing=last>=2?sh.getRange(2,1,last-1,2).getDisplayValues():[],ignored=0;var index={};existing.forEach(function(r,i){index[String(r[0]).trim()+'|'+String(r[1]).trim()]=i+2;});(rows||[]).forEach(function(x){var name=String(x.name||'').trim(),sec=String(x.section||'').trim();if(!name)return;var key=name+'|'+sec,row=index[key];if(!row){ignored++;return;}var raw=String(x.score==null?'':x.score).trim();if(raw===''){sh.getRange(row,c).clearContent();return;}var n=Number(raw.replace(',','.'));if(!isNaN(n))sh.getRange(row,c).setNumberFormat('0.##').setValue(Math.max(0,Math.min(20,n)));});return {success:true,ignored:ignored,message:'تم حفظ درجات الاختبار'+(ignored?'، وتم تجاهل '+ignored+' طالب غير موجود':'')};}
function importTestGrades(className,unit,rows,token){requireTeacherSession_(token);var sh=ensureTestSheet_(className),c=ensureTestColumn_(sh,unit),last=sh.getLastRow(),existing=last>=2?sh.getRange(2,1,last-1,2).getDisplayValues():[],index={};existing.forEach(function(r,i){index[String(r[0]).trim()+'|'+String(r[1]).trim()]=i+2;});var ignored=0;(rows||[]).forEach(function(x){var name=String(x.name||'').trim(),sec=String(x.section||'').trim(),row=index[name+'|'+sec];if(!row){ignored++;return;}var n=Number(String(x.score==null?'':x.score).replace(',','.'));if(!isNaN(n))sh.getRange(row,c).setNumberFormat('0.##').setValue(Math.max(0,Math.min(20,n)));});return {success:true,ignored:ignored,message:'تم استيراد الدرجات'+(ignored?'، وتم تجاهل '+ignored+' طالب غير موجود':'')};}
function exportTestTemplate(className,unit,token){requireTeacherSession_(token);var rows=[['اسم الطالب','الفصل','الدرجة']],roster=rosterForExport_(className);roster.forEach(function(st){rows.push([st.name,st.section,'']);});var name='درجات الاختبار '+normalizeClassName(className)+' - اختبار'+Number(unit)+'.xlsx',blob=makeXlsxBlob_(rows,name);return {success:true,name:name,dataUrl:'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,'+Utilities.base64Encode(blob.getBytes())};}
function getTestScores_(className){var sh=ensureTestSheet_(className),out={};if(sh.getLastRow()<2)return out;var last=sh.getLastColumn(),h=sh.getRange(1,1,1,last).getDisplayValues()[0],rows=sh.getRange(2,1,sh.getLastRow()-1,last).getDisplayValues();rows.forEach(function(r){var key=String(r[0]).trim()+'|'+String(r[1]).trim();out[key]={};for(var i=2;i<h.length;i++){var m=String(h[i]).match(/^اختبار(\d+)$/);if(m)out[key][String(Number(m[1]))]=r[i]||'';}});return out;}
function numberOrZero_(v){var n=Number(String(v==null?'':v).replace(',','.'));return isNaN(n)?0:n;}
function rosterForExport_(className){var sh=findClassSheet(SpreadsheetApp.getActiveSpreadsheet(),className),out=[];if(sh&&sh.getLastRow()>=2)sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues().forEach(function(r){if(String(r[1]).trim())out.push({name:String(r[1]).trim(),section:String(r[2]||'').trim()});});return sortStudentRows_(out,true);}
function exportTrackingWorkbook(token){requireTeacherSession_(token);var ss=SpreadsheetApp.getActiveSpreadsheet(),allSheets=[];[SHEET_SAHIN,SHEET_AWAL].forEach(function(cls){var roster=rosterForExport_(cls),settings=ss.getSheetByName(SHEET_SETTINGS),ids=[],settingsRows=settings&&settings.getLastRow()>=2?settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues():[];settingsRows.forEach(function(r){if(normalizeClassName(r[0])===normalizeClassName(cls)&&String(r[3]).trim()==='منشور'){var id=canonicalAssignmentId(r[1]);if(/^\d+-\d+$/.test(String(id))&&ids.indexOf(id)<0)ids.push(id);}});ids.sort(function(a,b){return String(a).localeCompare(String(b),undefined,{numeric:true});});var grade=findClassSheet(ss,cls),gmap={};if(grade&&grade.getLastRow()>=2&&grade.getLastColumn()>=4){var gradeLast=grade.getLastColumn(),gradeHeaders=grade.getRange(1,1,1,gradeLast).getDisplayValues()[0],gradeRows=grade.getRange(2,1,grade.getLastRow()-1,gradeLast).getDisplayValues();gradeRows.forEach(function(r){var rowKey=String(r[1]).trim()+'|'+String(r[2]||'').trim();gmap[rowKey]={};for(var gi=3;gi<gradeHeaders.length;gi++){var gid=canonicalAssignmentId(gradeHeaders[gi]);if(gid)gmap[rowKey][gid]=r[gi]||'0';}});}var matrix=activityMatrixMap_(cls),amap={};Object.keys(matrix.map).forEach(function(k){var x=matrix.map[k];ids.forEach(function(id){var u=normalizeDriveImageUrl_(x[id+'_رابط']||''),sc=activityScoreValue_(x[id+'_درجة']);if(u||sc!=='')amap[k+'|'+id]={score:sc===''?0:sc,url:u};});});var cd=getCheckSheetData_(cls),cmap={};cd.rows.forEach(function(r){var key=String(r[1]).trim()+'|'+String(r[0]).trim(),map={};cd.headers.forEach(function(h,j){map[h]=r[j+2]||'';});cmap[key]=map;});var checkHeaders=cd.headers.filter(function(id){return ids.indexOf(id)>=0;});var ach=ensureAchievementsSheet_(cls),ah=ach.getRange(1,1,1,Math.max(3,ach.getLastColumn())).getDisplayValues()[0].slice(3).filter(function(x){return String(x).trim();}),meta=ah.map(function(x){return getAchievementMeta_(cls,x);}),avals=ach.getLastRow()>=2?ach.getRange(2,1,ach.getLastRow()-1,Math.max(3,ach.getLastColumn())).getDisplayValues():[],amap2={};avals.forEach(function(r){var k=String(r[0]).trim()+'|'+String(r[2]||'').trim();amap2[k]=ah.map(function(_,j){return parseAchievementCell_(r[j+3]);});});var tests=getTestScores_(cls);
function baseRows(headers,fn){var out=[['اسم الطالب','الفصل'].concat(headers)];roster.forEach(function(st){out.push([st.name,st.section].concat(headers.map(function(x){return fn(st,x);})));});return out;}
allSheets.push({name:'واجبات '+cls,rows:baseRows(ids,function(st,id){return gmap[st.name+'|'+st.section]&&gmap[st.name+'|'+st.section][id]?numberOrZero_(gmap[st.name+'|'+st.section][id]):0;})});
allSheets.push({name:'أنشطة '+cls,rows:baseRows(ids,function(st,id){var x=amap[st.name+'|'+st.section+'|'+id];return x?numberOrZero_(x.score):0;})});
allSheets.push({name:'تحقق '+cls,rows:baseRows(checkHeaders,function(st,id){var r=cmap[st.name+'|'+st.section]||{};return numberOrZero_(r[id]);})});
var achRows=[['اسم الطالب','الفصل']],achSub=['',''];
ah.forEach(function(x,j){achRows[0].push(x);achSub.push('الدرجة');if(meta[j].type==='video'){achRows[0].push('');achSub.push('الرابط');}});
achRows.push(achSub);
roster.forEach(function(st){var cells=[st.name,st.section],vals=amap2[st.name+'|'+st.section]||[];ah.forEach(function(x,j){var q=vals[j]||{};cells.push(numberOrZero_(q.score));if(meta[j].type==='video')cells.push(q.url||'');});achRows.push(cells);});
allSheets.push({name:'إنجازات '+cls,rows:achRows});
var testHeaders=[],th=ensureTestSheet_(cls).getRange(1,1,1,Math.max(2,ensureTestSheet_(cls).getLastColumn())).getDisplayValues()[0];th.forEach(function(x){if(/^اختبار\d+$/.test(x))testHeaders.push(x);});
allSheets.push({name:'اختبارات '+cls,rows:baseRows(testHeaders,function(st,id){var t=tests[st.name+'|'+st.section]||{};return numberOrZero_(t[id.replace('اختبار','')]);})});});var name='كشوف المتابعة - جميع البيانات.xlsx',blob=makeMultiSheetXlsxBlob_(allSheets,name);return {success:true,name:name,dataUrl:'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,'+Utilities.base64Encode(blob.getBytes())};}
function makeMultiSheetXlsxBlob_(sheets,name){var wb=[],rels=[],files=[],over=[],idx=1;sheets.forEach(function(x){var rows=x.rows||[],sr=[];rows.forEach(function(row,ri){var cs=[];row.forEach(function(v,ci){var text=String(v==null?'':v),ref=colLetter_(ci+1)+(ri+1);if(ri>1&&ci>1&&text!==''&&!isNaN(Number(text)))cs.push('<c r="'+ref+'"><v>'+xmlEscape_(text)+'</v></c>');else cs.push('<c r="'+ref+'" t="inlineStr"><is><t>'+xmlEscape_(text)+'</t></is></c>');});sr.push('<row r="'+(ri+1)+'">'+cs.join('')+'</row>');});var fn='sheet'+idx+'.xml';files.push(Utilities.newBlob('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'+sr.join('')+'</sheetData></worksheet>','application/xml','xl/worksheets/'+fn));wb.push('<sheet name="'+xmlEscape_(x.name)+'" sheetId="'+idx+'" r:id="rId'+idx+'"/>');rels.push('<Relationship Id="rId'+idx+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/'+fn+'"/>');over.push('<Override PartName="/xl/worksheets/'+fn+'" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');idx++;});var types='<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+over.join('')+'</Types>',root='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',workbook='<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+wb.join('')+'</sheets></workbook>',wbrels='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+rels.join('')+'</Relationships>';files.unshift(Utilities.newBlob(types,'application/xml','[Content_Types].xml'),Utilities.newBlob(root,'application/xml','_rels/.rels'),Utilities.newBlob(workbook,'application/xml','xl/workbook.xml'),Utilities.newBlob(wbrels,'application/xml','xl/_rels/workbook.xml.rels'));return Utilities.zip(files,name);}
function getHomeworkManagement(className,section,unit,token){requireTeacherSession_(token);var ss=SpreadsheetApp.getActiveSpreadsheet(),settings=ss.getSheetByName(SHEET_SETTINGS),ids=[],titles={};if(settings&&settings.getLastRow()>=2)settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues().forEach(function(r){var id=canonicalAssignmentId(r[1]),m=String(id).match(/^(\d+)-/);if(normalizeClassName(r[0])===normalizeClassName(className)&&m&&String(Number(m[1]))===String(Number(unit))){ids.push(id);titles[id]=r[2]||'';}});var cls=findClassSheet(ss,className),students=[];if(cls&&cls.getLastRow()>=2)cls.getRange(2,1,cls.getLastRow()-1,3).getDisplayValues().forEach(function(r){if(String(r[1]).trim()&&(section==='__ALL__'||normalizeSectionValue_(r[2])===normalizeSectionValue_(section)))students.push({name:r[1],section:r[2]});});var gradeHeaders=[],values={};if(cls&&cls.getLastRow()>=2&&cls.getLastColumn()>=4){gradeHeaders=cls.getRange(1,4,1,cls.getLastColumn()-3).getDisplayValues()[0];cls.getRange(2,1,cls.getLastRow()-1,Math.max(3,cls.getLastColumn())).getDisplayValues().forEach(function(r){var n=String(r[1]).trim();if(n)values[n]={};gradeHeaders.forEach(function(h,i){values[n][canonicalAssignmentId(h)]=r[i+3]||'';});});}return {assignments:ids.map(function(id){return {id:id,title:titles[id]||''};}),rows:students.map(function(st){var x=values[st.name]||{};return {name:st.name,section:st.section,scores:ids.reduce(function(o,id){o[id]=x[id]||'';return o;},{})};})};}
function saveHomeworkManagement(className,section,unit,rows,token){requireTeacherSession_(token);var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=findClassSheet(ss,className);if(!sh)return {success:false,message:'لم يتم العثور على ورقة الصف'};var ids=[],settings=ss.getSheetByName(SHEET_SETTINGS);if(settings&&settings.getLastRow()>=2)settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues().forEach(function(r){var id=canonicalAssignmentId(r[1]),m=String(id).match(/^(\d+)-/);if(normalizeClassName(r[0])===normalizeClassName(className)&&m&&String(Number(m[1]))===String(Number(unit)))ids.push(id);});var headers=sh.getRange(1,4,1,Math.max(1,sh.getLastColumn()-3)).getDisplayValues()[0],cols={};headers.forEach(function(h,i){cols[canonicalAssignmentId(h)]=i+4;});var names=sh.getRange(2,2,Math.max(1,sh.getLastRow()-1),1).getDisplayValues();(rows||[]).forEach(function(st){var ri=-1;for(var i=0;i<names.length;i++)if(String(names[i][0]).trim()===String(st.name).trim()){ri=i+2;break;}if(ri<0)return;ids.forEach(function(id){var c=cols[id];if(!c)return;var raw=String((st.scores||{})[id]||'').trim();if(raw===''||/^\d+(?:[.,]\d+)?$/.test(raw))sh.getRange(ri,c).setValue(raw===''?'':Number(raw.replace(',','.')));});});return {success:true,message:'تم حفظ درجات الواجبات'};}
function publishedAssignmentRowsCached_(className){
  var cls=normalizeClassName(className), key='published_assignments_v2_'+encodeURIComponent(cls), cache=CacheService.getScriptCache(), hit=cache.get(key);
  if(hit){try{return JSON.parse(hit)||[];}catch(e){}}
  var ss=SpreadsheetApp.getActiveSpreadsheet(),settings=ss.getSheetByName(SHEET_SETTINGS),out=[];
  if(settings&&settings.getLastRow()>=2) settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues().forEach(function(r){
    if(normalizeClassName(r[0])===cls && String(r[3]).trim()==='منشور') out.push(r);
  });
  try{cache.put(key,JSON.stringify(out),30);}catch(e){}
  return out;
}
function clearPublishedAssignmentsCache_(className){
  try{CacheService.getScriptCache().remove('published_assignments_v2_'+encodeURIComponent(normalizeClassName(className)));}catch(e){}
}
function getAvailableUnits(className,token){
  requireTeacherSession_(token);var seen={};
  publishedAssignmentRowsCached_(className).forEach(function(r){var m=String(r[1]).trim().match(/^(\d+)-/);if(m)seen[String(Number(m[1]))]=true;});
  return Object.keys(seen).sort(function(a,b){return Number(a)-Number(b);});
}

function evaluationPhraseSheet_(){var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName('عبارات_التقييم');if(sh&&sh.getLastColumn()>=2&&String(sh.getRange(1,2).getDisplayValue()).trim()==='ممتاز'){sh.insertColumnBefore(2);sh.getRange(1,2).setValue('متفوق');}if(!sh){sh=ss.insertSheet('عبارات_التقييم');sh.getRange(1,1,1,6).setValues([['القسم','100٪','99–75٪','74–50٪','49–1٪','0٪']]);var rows=[['homework','✓ إنجاز كامل','أداء ممتاز','أداء جيد','يحتاج إلى متابعة','لم ينفذ'],['activity','★ مشاركة متميزة','مشاركة متميزة','مشاركة جيدة','يحتاج إلى تشجيع','لم يرسل'],['check','✓ إتقان تام','أتقن','فهم جزئي','يحتاج دعم','لم ينفذ'],['achievement','◆ إنجاز رائع','إنجاز متميز','إنجاز جيد','يحتاج إلى تطوير','لم ينفذ'],['test','🏆 درجة كاملة','إتقان ممتاز','مستوى جيد جدًا','يحتاج إلى مراجعة','لم يختبر'],['total','🌟 أداء استثنائي','ممتاز','جيد','ضعيف','صفر']];sh.getRange(2,1,rows.length,6).setValues(rows);}return sh;}
function getEvaluationPhrases(token){requireTeacherSession_(token);var sh=evaluationPhraseSheet_(),out={};if(sh.getLastRow()>=2)sh.getRange(2,1,sh.getLastRow()-1,6).getDisplayValues().forEach(function(r){out[r[0]]={outstanding:r[1],excellent:r[2],good:r[3],weak:r[4],zero:r[5]};});return out;}
function getEvaluationPhrasesPublic(){var sh=evaluationPhraseSheet_(),out={};if(sh.getLastRow()>=2)sh.getRange(2,1,sh.getLastRow()-1,6).getDisplayValues().forEach(function(r){out[r[0]]={outstanding:r[1],excellent:r[2],good:r[3],weak:r[4],zero:r[5]};});return out;}
function saveEvaluationPhrases(data,token){requireTeacherSession_(token);var sh=evaluationPhraseSheet_(),keys=['homework','activity','check','achievement','test','total'];keys.forEach(function(k,i){var r=data&&data[k]||{};sh.getRange(i+2,2,1,5).setValues([[String(r.outstanding||'').trim(),String(r.excellent||'').trim(),String(r.good||'').trim(),String(r.weak||'').trim(),String(r.zero||'').trim()]]);});return {success:true,message:'تم حفظ عبارات التقييم'};}
function getStudentTestScore(className,section,name){var key=String(name||'').trim()+'|'+String(section||'').trim(),t=getTestScores_(className),r=t[key]||{},units={},ach=ensureAchievementsSheet_(className),last=Math.max(3,ach.getLastColumn()),h=ach.getRange(1,1,1,last).getDisplayValues()[0],row=-1,rs=ach.getLastRow()>=2?ach.getRange(2,1,ach.getLastRow()-1,3).getDisplayValues():[];Object.keys(r).forEach(function(k){units[String(Number(k))]=true;});for(var i=0;i<rs.length;i++)if(String(rs[i][0]).trim()===String(name||'').trim()&&normalizeClassName(rs[i][1])===normalizeClassName(className)&&(String(rs[i][2]).trim()===String(section||'').trim()||!String(section||'').trim())){row=i+2;break;}var extras={};if(row>0){var cells=ach.getRange(row,1,1,last).getDisplayValues()[0];h.slice(3).forEach(function(n,j){var m=getAchievementMeta_(className,String(n||'').trim());var sc=numberOrZero_(parseAchievementCell_(cells[j+3]).score);if(m.target==='test'&&sc>0){var u=String(Number(m.unit));units[u]=true;extras[u]=(extras[u]||0)+sc;}});}var vals=Object.keys(units).map(function(k){return Math.min(20,numberOrZero_(r[k])+numberOrZero_(extras[k]));});return vals.length?Math.max.apply(null,vals):0;}

// ==================== Activity long storage v3 ====================
function longActivitySheetName_(className){return activityClassKey_(className)==='أول متوسط'?'تسليم أنشطة أول متوسط':'تسليم أنشطة سادس';}
function longActivityEnsure_(className){var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(longActivitySheetName_(className));if(!sh)sh=ss.insertSheet(longActivitySheetName_(className));if(sh.getLastRow()<1)sh.getRange(1,1,1,8).setValues([['اسم الطالب','الفصل','الوحدة','رقم الدرس','الرابط','الدرجة','الحالة','الملاحظة']]);return sh;}
function longActivitySec_(v){return normalizeSectionValue_(v);}
function longActivityKey_(name,section,id){var p=String(id||'').split('-');return String(name||'').trim()+'|'+longActivitySec_(section)+'|'+String(Number(p[0]))+'|'+String(Number(p[1]));}
function longActivityData_(className){var sh=longActivityEnsure_(className),rows=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,8).getValues():[],map={};rows.forEach(function(r,i){var id=Number(r[2])+'-'+Number(r[3]),k=longActivityKey_(r[0],r[1],id);if(String(r[0]||'').trim())map[k]={row:i+2,name:String(r[0]).trim(),section:longActivitySec_(r[1]),id:id,url:String(r[4]||'').trim(),score:activityScoreValue_(r[5]),status:String(r[6]||'').trim(),note:String(r[7]||'').trim()};});return {sheet:sh,map:map};}
function longActivityFind_(d,name,section,id){var x=d.map[longActivityKey_(name,section,id)];if(x)return x;var target=String(name||'').trim(),found=null;Object.keys(d.map).some(function(k){var v=d.map[k];if(v.name===target&&v.id===canonicalAssignmentId(id)){found=v;return true;}return false;});return found;}
function longActivityIds_(className,unit){var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(SHEET_SETTINGS),out=[];if(sh&&sh.getLastRow()>=2)sh.getRange(2,1,sh.getLastRow()-1,6).getDisplayValues().forEach(function(r){var id=canonicalAssignmentId(r[1]);if(activityClassKey_(r[0])===activityClassKey_(className)&&String(r[3]).trim()==='منشور'&&(!unit||String(id).indexOf(String(Number(unit))+'-')===0)&&out.indexOf(id)<0)out.push(id);});out.sort(function(a,b){var x=a.split('-'),y=b.split('-');return Number(x[0])-Number(y[0])||Number(x[1])-Number(y[1]);});return out;}
function getStudentAssignments(className,studentName,requestedSection){var ss=SpreadsheetApp.getActiveSpreadsheet(),cls=normalizeClassName(className),name=String(studentName||'').trim(),wanted=longActivitySec_(requestedSection),settings=ss.getSheetByName(SHEET_SETTINGS),solved={},section=wanted,grade=findClassSheet(ss,cls);if(grade&&grade.getLastRow()>=2){var nr=grade.getRange(2,2,grade.getLastRow()-1,2).getDisplayValues();for(var i=0;i<nr.length;i++)if(String(nr[i][0]).trim()===name&&(!wanted||wanted==='__ALL__'||longActivitySec_(nr[i][1])===wanted)){section=longActivitySec_(nr[i][1]);if(grade.getLastColumn()>=4){var hh=grade.getRange(1,4,1,grade.getLastColumn()-3).getDisplayValues()[0],vv=grade.getRange(i+2,4,1,grade.getLastColumn()-3).getDisplayValues()[0];hh.forEach(function(h,j){if(String(vv[j]).trim()!=='')solved[canonicalAssignmentId(h)]=vv[j];});}break;}}var checks=getCheckGradesForStudent_(cls,section,name),d=longActivityData_(cls),now=Date.now();return (settings&&settings.getLastRow()>=2?settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues():[]).filter(function(r){return activityClassKey_(r[0])===activityClassKey_(cls)&&String(r[3]).trim()==='منشور';}).map(function(r){var id=canonicalAssignmentId(r[1]),st=r[4]?new Date(r[4]).getTime():0,en=r[5]?new Date(r[5]).getTime():0,x=longActivityFind_(d,name,section,id)||{},sc=activityScoreValue_(x.score),url=normalizeDriveImageUrl_(x.url||'');return {id:id,title:r[2],startDate:r[4]||'',endDate:r[5]||'',isSolved:Object.prototype.hasOwnProperty.call(solved,id),isExpired:!!(en&&en<Date.now()),isNotStarted:!!(st&&st>Date.now()),score:solved[id]||null,isActivitySubmitted:!!url||sc!=='',isActivityGraded:sc!==''&&Number(sc)>0,activityUrl:url,activityScore:sc,activityNote:x.note||'',activityStatus:x.status||activityMatrixStatus_(!!url,sc),checkScore:checks[id]||null};});}
function uploadStudentActivity(dataUrl,className,section,studentName,assignmentId){var settings=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS),aid=canonicalAssignmentId(assignmentId),activeRows=settings&&settings.getLastRow()>=2?settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues():[],published=activeRows.filter(function(r){return normalizeClassName(r[0])===normalizeClassName(className)&&canonicalAssignmentId(r[1])===aid&&String(r[3]).trim()==='منشور';})[0];if(published&&published[5]&&new Date(published[5]).getTime()<Date.now())return {success:false,message:'انتهى وقت تسليم الواجب، لا يمكن إرسال النشاط بعد انتهاء النشر'};var id=aid,d=longActivityData_(className),old=longActivityFind_(d,studentName,section,id),oldScore=old?activityScoreValue_(old.score):'';if(oldScore!==''&&Number(oldScore)>0)return {success:false,message:'تم تثبيت درجة النشاط من المعلم ولا يمكن إعادة الإرسال'};var root=getClassImageFolder_(className,'صور الأنشطة'),folder=getOrCreateChildFolder_(root,id),safe=String(studentName).replace(/[\\/:*?"<>|]/g,'_'),parts=String(dataUrl||'').split(','),mime=((parts[0]||'').match(/data:([^;]+);/)||[])[1]||'image/jpeg',file=folder.createFile(Utilities.newBlob(Utilities.base64Decode(parts[1]||''),mime,safe+' - واجب '+id+'.jpg')),sh=d.sheet,row=old?old.row:sh.getLastRow()+1,p=id.split('-');sh.getRange(row,1,1,8).setValues([[String(studentName).trim(),longActivitySec_(section),Number(p[0]),Number(p[1]),file.getUrl(),'','تم الإرسال','']]);return {success:true,url:file.getUrl(),message:'تم إرسال النشاط بنجاح'};}
function getActivityManagement(className,section,unit,token){requireTeacherSession_(token);var ids=longActivityIds_(className,unit),cls=findClassSheet(SpreadsheetApp.getActiveSpreadsheet(),className),students=[],d=longActivityData_(className);if(cls&&cls.getLastRow()>=2)cls.getRange(2,1,cls.getLastRow()-1,3).getDisplayValues().forEach(function(r){if(String(r[1]).trim()&&(section==='__ALL__'||longActivitySec_(r[2])===longActivitySec_(section)))students.push({name:String(r[1]).trim(),section:longActivitySec_(r[2])});});students=sortStudentRows_(students,section==='__ALL__');return {assignments:ids.map(function(id){return {id:id,title:''};}),students:students,rows:students.map(function(st){return {name:st.name,section:st.section,activities:ids.map(function(id){var x=longActivityFind_(d,st.name,st.section,id)||{},sc=activityScoreValue_(x.score),url=normalizeDriveImageUrl_(x.url||'');return {id:id,url:url,score:sc,note:x.note||'',status:x.status||activityMatrixStatus_(!!url,sc)};})};})};}
function saveActivityManagement(className,section,rows,token){requireTeacherSession_(token);var d=longActivityData_(className),sh=d.sheet;(rows||[]).forEach(function(st){(st.activities||[]).forEach(function(a){var id=canonicalAssignmentId(a.id),p=id.split('-'),raw=String(a.score==null?'':a.score).trim();if(raw!==''&&!/^\d+(?:[.,،]\d+)?$/.test(raw))return;var old=longActivityFind_(d,st.name,st.section,id),row=old?old.row:sh.getLastRow()+1,url=String(a.url||'').trim(),sc=raw===''?'':Number(raw.replace(/[،,]/g,'.'));if(!old){sh.getRange(row,1,1,8).setValues([[String(st.name).trim(),longActivitySec_(st.section||section),Number(p[0]),Number(p[1]),url,sc,activityMatrixStatus_(!!url,sc),String(a.note||'')]]);d.map[longActivityKey_(st.name,st.section,id)]={row:row,name:String(st.name).trim(),section:longActivitySec_(st.section||section),id:id,url:url,score:sc,status:activityMatrixStatus_(!!url,sc),note:String(a.note||'')};}else{if(url)sh.getRange(row,5).setValue(url);if(sc==='')sh.getRange(row,6).clearContent();else {sh.getRange(row,6).setValue(sc);sh.getRange(row,8).clearContent();}sh.getRange(row,7).setValue(activityMatrixStatus_(!!(url||old.url),sc));if(sc==='')sh.getRange(row,8).setValue(String(a.note||''));}});});return {success:true,message:'تم حفظ الدرجة والرابط والحالة في نفس صف الطالب والدرس'};}


var DEFAULT_ACTIVITY_NOTE_='المطلوب ارسال صوره من ورقه تدرب وحل المسائل بعد كتابتها صحيحه من نموذج الاجابه بخط يد الطالب';
function getActivityDefaultNote(token){requireTeacherSession_(token);return PropertiesService.getScriptProperties().getProperty('activity_default_note')||DEFAULT_ACTIVITY_NOTE_;}
function saveActivityDefaultNote(note,token){requireTeacherSession_(token);var v=String(note==null?'':note).trim();PropertiesService.getScriptProperties().setProperty('activity_default_note',v);return {success:true,message:'تم حفظ ملاحظة النشاط الافتراضية'};}
function setActivityNote(className,section,studentName,assignmentId,action,token){requireTeacherSession_(token);var d=longActivityData_(className),id=canonicalAssignmentId(assignmentId),old=longActivityFind_(d,studentName,section,id),row=old?old.row:d.sheet.getLastRow()+1,p=id.split('-'),note=String(action||'')==='delete'?'':(PropertiesService.getScriptProperties().getProperty('activity_default_note')||DEFAULT_ACTIVITY_NOTE_);if(!old)d.sheet.getRange(row,1,1,8).setValues([[String(studentName||'').trim(),longActivitySec_(section),Number(p[0]),Number(p[1]),'','','',note]]);else d.sheet.getRange(row,8).setValue(note);return {success:true,note:note,message:note?'تم إرسال الملاحظة للطالب':'تم حذف الملاحظة'};}
// تقرير الطلاب الضعاف: متوسط الوحدات المحددة، مع نفس مقياس ملخص الطالب (60 درجة).
function getWeakStudentsReportCached(className, section, units, token){
  requireTeacherSession_(token);
  var key='weak_report_v2_'+encodeURIComponent(normalizeClassName(className)+'|'+String(section||'')+'|'+(units||[]).join(',')),cache=CacheService.getScriptCache(),hit=cache.get(key);
  if(hit){try{return JSON.parse(hit)||[];}catch(e){}}
  var rows=getWeakStudentsReport(className,section,units,token);
  try{cache.put(key,JSON.stringify(rows),30);}catch(e){}
  return rows;
}
function getWeakStudentsReport(className, section, units, token){
  requireTeacherSession_(token);
  var ss=SpreadsheetApp.getActiveSpreadsheet(),cls=normalizeClassName(className);
  units=(units||[]).map(function(x){return String(Number(x));}).filter(function(x){return x!=='NaN';});
  if(!units.length)return [];
  var sh=findClassSheet(ss,cls),roster=[];
  if(sh&&sh.getLastRow()>=2){var last=Math.max(3,sh.getLastColumn()),headers=sh.getRange(1,1,1,last).getDisplayValues()[0],rows=sh.getRange(2,1,sh.getLastRow()-1,last).getDisplayValues();rows.forEach(function(r){var name=String(r[1]||'').trim(),sec=String(r[2]||'').trim();if(!name||(section!=='__ALL__'&&normalizeSectionValue_(sec)!==normalizeSectionValue_(section)))return;var hw={};headers.slice(3).forEach(function(h,i){var id=canonicalAssignmentId(h);if(/^\d+-/.test(id))hw[id]=r[i+3]||'';});roster.push({name:name,section:sec,hw:hw});});}
  var settings=ss.getSheetByName(SHEET_SETTINGS),idsByUnit={};
  if(settings&&settings.getLastRow()>=2)settings.getRange(2,1,settings.getLastRow()-1,6).getDisplayValues().forEach(function(r){if(normalizeClassName(r[0])!==cls||String(r[3]).trim()!=='منشور')return;var id=canonicalAssignmentId(r[1]),m=String(id).match(/^(\d+)-/);if(m&&units.indexOf(String(Number(m[1])))>=0)(idsByUnit[m[1]]||(idsByUnit[m[1]]=[])).push(id);});
  var activity=longActivityData_(cls),tests=getTestScores_(cls),cd=getCheckSheetData_(cls),checkMap={},ach=ensureAchievementsSheet_(cls),alast=Math.max(3,ach.getLastColumn()),ahead=ach.getRange(1,1,1,alast).getDisplayValues()[0].slice(3),ameta={};ahead.forEach(function(n){ameta[String(n)]=getAchievementMeta_(cls,String(n));});
  var avals=ach.getLastRow()>=2?ach.getRange(2,1,ach.getLastRow()-1,alast).getDisplayValues():[],achMap={};avals.forEach(function(r){var k=String(r[0]).trim()+'|'+String(r[2]||'').trim(),arr=achMap[k]||[];ahead.forEach(function(n,j){var cell=parseAchievementCell_(r[j+3]);if(cell.url||cell.score!=='')arr.push({unit:String(Number((ameta[String(n)]||{}).unit)),score:cell.score,target:(ameta[String(n)]||{}).target});});achMap[k]=arr;});
  cd.rows.forEach(function(r){var k=String(r[1]).trim()+'|'+String(r[0]).trim(),m={};cd.headers.forEach(function(h,j){m[canonicalAssignmentId(h)]=r[j+2]||'';});checkMap[k]=m;});
  function avg(vals,max){var a=vals.filter(function(v){return v!==null&&v!==''&&!isNaN(Number(v));}).map(Number);return a.length?Math.min(max,a.reduce(function(x,y){return x+y;},0)/a.length):0;}
  var out=[];roster.forEach(function(st){var h=[],a=[],c=[],e=[],tst=[],key=st.name+'|'+st.section,check=checkMap[key]||{},achRows=achMap[key]||[],testRow=tests[key]||{};units.forEach(function(u){var ids=idsByUnit[u]||[];ids.forEach(function(id){h.push(Number(st.hw[id])||0);var x=longActivityFind_(activity,st.name,st.section,id)||{};a.push(activityScoreValue_(x.score));c.push(Number(check[id])||0);});var es=achRows.filter(function(x){return x.unit===u&&x.target!=='test';}).map(function(x){return Number(x.score)||0;});e.push(es.length?Math.min(10,es.reduce(function(x,y){return x+y;},0)):0);tst.push(Math.min(20,Number(testRow[u])||0));});var r={name:st.name,className:className,section:st.section,homework:avg(h,10),activity:avg(a,10),check:avg(c,10),achievement:avg(e,10),test:avg(tst,20)};r.total=r.homework+r.activity+r.check+r.achievement+r.test;if(r.total<30)out.push(r);});return out.sort(function(x,y){return String(x.section||'').localeCompare(String(y.section||''),'ar',{numeric:true,sensitivity:'base'})||normalizeStudentName_(x.name).localeCompare(normalizeStudentName_(y.name),'ar',{numeric:true,sensitivity:'base'});});
}
