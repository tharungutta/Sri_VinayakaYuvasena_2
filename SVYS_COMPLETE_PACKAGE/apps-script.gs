const SHEET_MEMBERS = 'Members';
const SHEET_PAYMENTS = 'Payments';
const SHEET_EXPENSES = 'Expenses';
const DRIVE_ROOT = 'SRI VINAYAKA YUVA SENA';

function setupSheets(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let m=ss.getSheetByName(SHEET_MEMBERS)||ss.insertSheet(SHEET_MEMBERS);
  let p=ss.getSheetByName(SHEET_PAYMENTS)||ss.insertSheet(SHEET_PAYMENTS);
  let e=ss.getSheetByName(SHEET_EXPENSES)||ss.insertSheet(SHEET_EXPENSES);
  if(m.getLastRow()===0)m.appendRow(['Member ID','Name','Mobile No']);
  if(p.getLastRow()===0)p.appendRow(['Payment ID','Member ID','Member Name','Mobile No','Month','Amount','Payment Date','Screenshot URL','Screenshot File','Created At']);
  if(e.getLastRow()===0)e.appendRow(['Expense ID','Category','Title','Amount','Date','Paid To','Receipt URL','Receipt File','Created At']);
  let a=ss.getSheetByName('Login Activity')||ss.insertSheet('Login Activity');
  if(a.getLastRow()===0)a.appendRow(['Login ID','Firebase UID','Display Name','Email','Login At','Method']);
}
function doGet(){
  setupSheets();
  return jsonOut(getData_());
}
function doPost(e){
  try{
    setupSheets();
    const body=JSON.parse(e.postData.contents||'{}');
    if(body.action==='logLogin') return jsonOut(logLogin_(body));
    if(body.action==='getLoginActivity'){auth_(body);return jsonOut(getLoginActivity_());}
    if(body.action==='addContribution') return jsonOut(addContribution_(body));
    if(body.action==='addMember') return jsonOut(addMember_(body));
    if(body.action==='addExpense') return jsonOut(addExpense_(body));
    if(body.action==='deleteExpense') return jsonOut(deleteExpense_(body));
    return jsonOut({ok:false,error:'Unknown action'});
  }catch(err){return jsonOut({ok:false,error:String(err.message||err)})}
}
function auth_(body){
  const expectedUser=PropertiesService.getScriptProperties().getProperty('ADMIN_USERNAME')||'Tharun';
  const expectedPass=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD')||'Tharun@18';
  if(body.username!==expectedUser||body.password!==expectedPass)throw new Error('Invalid administrator credentials.');
}
function addMember_(b){
  auth_(b);
  const ss=SpreadsheetApp.getActive();const sh=ss.getSheetByName(SHEET_MEMBERS);
  const name=String(b.name||'').trim(),phone=String(b.phone||'').trim();
  if(!name)throw new Error('Member name is required.');
  if(!/^\d{10}$/.test(phone))throw new Error('Mobile number must contain 10 digits.');
  const values=sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++)if(String(values[i][1]).trim().toLowerCase()===name.toLowerCase())throw new Error('This member already exists.');
  const id='M'+Date.now();sh.appendRow([id,name,phone]);
  return getData_();
}
function addContribution_(b){
  auth_(b);
  if(!b.memberId||!b.month||!b.amount||!b.date)throw new Error('Member, month, amount and date are required.');
  if(!b.screenshotData)throw new Error('Payment screenshot is mandatory.');
  const ss=SpreadsheetApp.getActive();const ms=ss.getSheetByName(SHEET_MEMBERS),ps=ss.getSheetByName(SHEET_PAYMENTS);
  const members=ms.getDataRange().getValues();let member=null;
  for(let i=1;i<members.length;i++)if(String(members[i][0])===String(b.memberId))member={id:members[i][0],name:members[i][1],phone:members[i][2]};
  if(!member)throw new Error('Member not found.');
  const bytes=Utilities.base64Decode(String(b.screenshotData).split(',').pop());
  const mime=b.mimeType||'image/jpeg';
  const blob=Utilities.newBlob(bytes,mime,b.screenshotName||('payment-'+Date.now()+'.jpg'));
  const root=getOrCreateFolder_(DRIVE_ROOT);
  const yearFolder=getOrCreateFolder_(root,new Date(b.date).getFullYear()+'-'+(new Date(b.date).getFullYear()+1));
  const monthFolder=getOrCreateFolder_(yearFolder,String(b.month));
  const memberFolder=getOrCreateFolder_(monthFolder,member.name);
  const file=memberFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  const id='P'+Date.now();
  ps.appendRow([id,member.id,member.name,member.phone,b.month,Number(b.amount),b.date,file.getUrl(),file.getName(),new Date()]);
  return getData_();
}
function addExpense_(b){
  auth_(b);
  if(!b.category||!b.title||!b.amount||!b.date)throw new Error('Category, title, amount and date are required.');
  const ss=SpreadsheetApp.getActive();
  let es=ss.getSheetByName(SHEET_EXPENSES);
  if(!es){setupSheets();es=ss.getSheetByName(SHEET_EXPENSES);}
  let receiptUrl='',receiptName='';
  if(b.receiptData){
    const bytes=Utilities.base64Decode(String(b.receiptData).split(',').pop());
    const mime=b.mimeType||'image/jpeg';
    receiptName=b.receiptName||('receipt-'+Date.now()+'.jpg');
    const blob=Utilities.newBlob(bytes,mime,receiptName);
    const root=getOrCreateFolder_(DRIVE_ROOT);
    const expFolder=getOrCreateFolder_(root,'Expenses');
    const file=expFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
    receiptUrl=file.getUrl();
  }
  const id='E'+Date.now();
  es.appendRow([id,b.category,b.title,Number(b.amount),b.date,b.paidTo||'',receiptUrl,receiptName,new Date()]);
  return getData_();
}
function deleteExpense_(b){
  auth_(b);
  if(!b.id)throw new Error('Expense ID is required.');
  const ss=SpreadsheetApp.getActive();
  const es=ss.getSheetByName(SHEET_EXPENSES);
  if(!es)throw new Error('Expenses sheet not found.');
  const rows=es.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(b.id)){
      es.deleteRow(i+1);
      return getData_();
    }
  }
  throw new Error('Expense record not found.');
}
function getOrCreateFolder_(parent,name){
  if(name===undefined){name=parent;parent=DriveApp.getRootFolder();} // root call: getOrCreateFolder_('Folder name')
  const it=parent.getFoldersByName(name);return it.hasNext()?it.next():parent.createFolder(name);
}

function firebaseApiKey_(){
  return PropertiesService.getScriptProperties().getProperty('FIREBASE_WEB_API_KEY')||'AIzaSyBd6kCvSAtdJBJajM2MX01Ux6fiLJGivII';
}
function diagnoseLoginLogging(){
  setupSheets();
  const sh=SpreadsheetApp.getActive().getSheetByName('Login Activity');
  Logger.log('1) Login Activity sheet: '+(sh?'found':'MISSING'));
  try{sh.appendRow(['TEST','','','','','diagnose']);sh.deleteRow(sh.getLastRow());Logger.log('2) Writing to the sheet: OK');}
  catch(e){Logger.log('2) Writing to the sheet FAILED: '+e.message);}
  try{
    const r=UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key='+encodeURIComponent(firebaseApiKey_()),{method:'post',contentType:'application/json',payload:JSON.stringify({idToken:'diagnostic'}),muteHttpExceptions:true});
    const code=r.getResponseCode(),txt=r.getContentText().slice(0,300);
    Logger.log('3) Firebase check -> HTTP '+code+' '+txt);
    if(code===400)Logger.log('   GOOD: key and API work (400 INVALID_ID_TOKEN is expected for a fake token).');
    else if(code===403)Logger.log('   PROBLEM: key is blocked. Either the Identity Toolkit API is disabled, or the API key has website/HTTP-referrer restrictions that block Apps Script. In Google Cloud Console > APIs & Services > Credentials, set this key to no application restriction (keep API restriction to Identity Toolkit API).');
  }catch(e){Logger.log('3) UrlFetchApp FAILED (usually missing authorization — run this function and click Review permissions > Allow): '+e.message);}
}
function logLogin_(b){
  const apiKey=firebaseApiKey_();
  if(!b.idToken)throw new Error('Missing Firebase login token.');
  const response=UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key='+encodeURIComponent(apiKey),{method:'post',contentType:'application/json',payload:JSON.stringify({idToken:b.idToken}),muteHttpExceptions:true});
  const result=JSON.parse(response.getContentText()||'{}');
  if(response.getResponseCode()!==200||!result.users||!result.users.length)throw new Error('Firebase token verification failed.');
  const u=result.users[0];
  if(!u.emailVerified)throw new Error('Email address is not verified.');
  const sh=SpreadsheetApp.getActive().getSheetByName('Login Activity');
  const at=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
  sh.appendRow(['L'+Date.now(),String(u.localId||''),String(u.displayName||''),String(u.email||''),at,String(b.method||'password')]);
  return {ok:true};
}
function getLoginActivity_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('Login Activity');
  const values=sh.getDataRange().getValues();
  const rows=values.slice(1).filter(r=>r[0]).slice(-200).reverse().map(r=>({displayName:String(r[2]||''),email:String(r[3]||''),loginAt:r[4] instanceof Date?Utilities.formatDate(r[4],Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'):String(r[4]||''),method:String(r[5]||'password')}));
  return {ok:true,activity:rows};
}
function getData_(){
  const ss=SpreadsheetApp.getActive();
  const ms=ss.getSheetByName(SHEET_MEMBERS),ps=ss.getSheetByName(SHEET_PAYMENTS),es=ss.getSheetByName(SHEET_EXPENSES);
  const mv=ms?ms.getDataRange().getValues():[],pv=ps?ps.getDataRange().getValues():[],ev=es?es.getDataRange().getValues():[];
  const members=mv.slice(1).filter(r=>r[0]).map(r=>({id:String(r[0]),name:String(r[1]||''),phone:String(r[2]||'')}));
  const payments=pv.slice(1).filter(r=>r[0]).map(r=>({id:String(r[0]),memberId:String(r[1]),memberName:String(r[2]||''),phone:String(r[3]||''),month:String(r[4]||''),amount:Number(r[5]||0),date:r[6] instanceof Date?Utilities.formatDate(r[6],Session.getScriptTimeZone(),'yyyy-MM-dd'):String(r[6]||''),screenshotUrl:String(r[7]||''),screenshotName:String(r[8]||'')}));
  const expenses=ev.slice(1).filter(r=>r[0]).map(r=>({id:String(r[0]),category:String(r[1]||'Miscellaneous'),title:String(r[2]||''),amount:Number(r[3]||0),date:r[4] instanceof Date?Utilities.formatDate(r[4],Session.getScriptTimeZone(),'yyyy-MM-dd'):String(r[4]||''),paidTo:String(r[5]||''),receiptUrl:String(r[6]||''),receiptName:String(r[7]||'')}));
  return {ok:true,members,payments,expenses};
}
function jsonOut(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}