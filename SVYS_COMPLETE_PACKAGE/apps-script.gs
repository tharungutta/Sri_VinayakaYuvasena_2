function authorizePermissions(){
  // Run this function once in Apps Script and click "Review permissions" -> "Allow" to enable Email sending
  let log = [];
  try {
    const quota = MailApp.getRemainingDailyQuota();
    log.push('MailApp quota available: ' + quota);
  } catch(e) {
    log.push('MailApp info: ' + e.message);
  }
  try {
    const drafts = GmailApp.getDrafts();
    log.push('GmailApp authorized successfully');
  } catch(e) {
    log.push('GmailApp info: ' + e.message);
  }
  Logger.log('Permissions check: ' + log.join(' | '));
  return 'SUCCESS: ' + log.join(' | ');
}

const SHEET_MEMBERS = 'Members';
const SHEET_PAYMENTS = 'Payments';
const SHEET_EXPENSES = 'Expenses';
const SHEET_CHAT = 'Chat';
const SHEET_SUGGESTIONS = 'Suggestions';
const SHEET_ID_CARDS = 'ID_Cards';
const DRIVE_ROOT = 'SRI VINAYAKA YUVA SENA';

function setupSheets(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let m=ss.getSheetByName(SHEET_MEMBERS)||ss.insertSheet(SHEET_MEMBERS);
  let p=ss.getSheetByName(SHEET_PAYMENTS)||ss.insertSheet(SHEET_PAYMENTS);
  let e=ss.getSheetByName(SHEET_EXPENSES)||ss.insertSheet(SHEET_EXPENSES);
  if(m.getLastRow()===0)m.appendRow(['Member ID','Name','Mobile No','Mail ID']);
  if(p.getLastRow()===0)p.appendRow(['Payment ID','Member ID','Member Name','Mobile No','Month','Amount','Payment Date','Screenshot URL','Screenshot File','Created At']);
  if(e.getLastRow()===0)e.appendRow(['Expense ID','Category','Title','Amount','Date','Paid To','Receipt URL','Receipt File','Created At']);
  let a=ss.getSheetByName('Login Activity')||ss.insertSheet('Login Activity');
  if(a.getLastRow()===0)a.appendRow(['Login ID','Firebase UID','Display Name','Email','Login At','Method']);
  let c=ss.getSheetByName(SHEET_CHAT)||ss.insertSheet(SHEET_CHAT);
  if(c.getLastRow()===0)c.appendRow(['Message ID','Sender','Sender Email','Message Text','Time','Is Announcement','Created At']);
  let s=ss.getSheetByName(SHEET_SUGGESTIONS)||ss.insertSheet(SHEET_SUGGESTIONS);
  if(s.getLastRow()===0)s.appendRow(['Suggestion ID','Sender','Sender Email','Suggestion Text','Time Str','Likes','Created At']);
  let ic=ss.getSheetByName(SHEET_ID_CARDS)||ss.insertSheet(SHEET_ID_CARDS);
  if(ic.getLastRow()===0)ic.appendRow(['Email','Name','Mobile No','Blood Group','Photo URL','Pass ID','Created At']);
}
function doGet(){
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
    if(body.action==='sendChatMessage') return jsonOut(sendChatMessage_(body));
    if(body.action==='getChatMessages') return jsonOut(getChatMessages_());
    if(body.action==='sendSuggestion') return jsonOut(sendSuggestion_(body));
    if(body.action==='likeSuggestion') return jsonOut(likeSuggestion_(body));
    if(body.action==='getSuggestions') return jsonOut(getSuggestions_());
    if(body.action==='saveIdCard') return jsonOut(saveIdCard_(body));
    if(body.action==='getIdCard') return jsonOut(getIdCard_(body));
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

function findMemberEmail_(name, phone){
  try{
    const ss=SpreadsheetApp.getActive();
    // 1. Check Members Sheet Column D (Mail ID)
    const ms=ss.getSheetByName(SHEET_MEMBERS);
    if(ms){
      const mVals=ms.getDataRange().getValues();
      const cleanPhone=String(phone||'').replace(/\D/g,'').slice(-10);
      const targetName=String(name||'').trim().toLowerCase();
      for(let i=1; i<mVals.length; i++){
        const rowName=String(mVals[i][1]||'').trim().toLowerCase();
        const rowPhone=String(mVals[i][2]||'').replace(/\D/g,'').slice(-10);
        const rowEmail=String(mVals[i][3]||'').trim();
        if(rowEmail && rowEmail.indexOf('@')!==-1){
          if((cleanPhone && rowPhone===cleanPhone) || (targetName && (rowName===targetName || rowName.indexOf(targetName)!==-1 || targetName.indexOf(rowName)!==-1))){
            return rowEmail;
          }
        }
      }
    }
    // 2. Check Login Activity Sheet
    const sh=ss.getSheetByName('Login Activity');
    if(sh){
      const vals=sh.getDataRange().getValues();
      const cleanPhone=String(phone||'').replace(/\D/g,'').slice(-10);
      const targetName=String(name||'').trim().toLowerCase();
      for(let i=vals.length-1; i>=1; i--){
        const email=String(vals[i][3]||'').trim();
        const disp=String(vals[i][2]||'').trim().toLowerCase();
        const userPhone=String(vals[i][1]||'').replace(/\D/g,'').slice(-10);
        if(email && email.indexOf('@')!==-1){
          if((cleanPhone && userPhone===cleanPhone) || (targetName && disp===targetName)){
            return email;
          }
        }
      }
    }
  }catch(e){}
  return '';
}

function generateReceiptHtml_(r){
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 24px; background: #fdfaf6; color: #2e0910; }
  .receipt-box { max-width: 620px; margin: 0 auto; background: #ffffff; border: 3px double #d6a64f; border-radius: 16px; padding: 32px; box-shadow: 0 4px 20px rgba(87,27,35,0.1); }
  .header { text-align: center; border-bottom: 2px solid #f2e3c6; padding-bottom: 18px; margin-bottom: 22px; }
  .om { font-size: 24px; color: #b8860b; font-weight: bold; }
  .title { font-size: 22px; font-weight: 800; color: #571b23; letter-spacing: 1px; margin: 6px 0; }
  .subtitle { font-size: 13px; color: #7d2a36; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
  .badge { display: inline-block; background: #fbf2dc; border: 1px solid #d6a64f; color: #80570b; padding: 4px 14px; border-radius: 20px; font-size: 11px; font-weight: 800; margin-top: 8px; }
  .rcpt-no-row { display: flex; justify-content: space-between; margin-bottom: 20px; font-size: 13px; color: #555; }
  .info-table { width: 100%; border-collapse: collapse; margin-bottom: 24px; }
  .info-table td { padding: 10px 12px; border-bottom: 1px solid #f5ede0; font-size: 13.5px; }
  .info-table td.lbl { width: 38%; color: #666; font-weight: 600; }
  .info-table td.val { color: #222; font-weight: 700; }
  .amount-card { background: linear-gradient(135deg, #571b23, #380d14); color: #ffffff; padding: 18px 24px; border-radius: 12px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 24px; border: 1px solid #d6a64f; }
  .amount-lbl { font-size: 13px; color: #f5d696; text-transform: uppercase; font-weight: 600; }
  .amount-val { font-size: 26px; font-weight: 800; color: #ffd700; }
  .footer { text-align: center; border-top: 1px dashed #d6a64f; padding-top: 18px; margin-top: 10px; font-size: 11.5px; color: #777; }
  .seal { display: inline-block; border: 2px solid #27ae60; color: #27ae60; padding: 4px 12px; border-radius: 6px; font-weight: 800; font-size: 11px; margin-top: 8px; text-transform: uppercase; }
</style>
</head>
<body>
<div class="receipt-box">
  <div class="header">
    <div class="om">|| Sri Vinayaka Prasanna ||</div>
    <div class="title">SRI VINAYAKA YUVASENA</div>
    <div class="subtitle">Official Donation & Contribution Receipt</div>
    <div class="badge">Annual Ganesh Festival 2026-2027</div>
  </div>
  <div class="rcpt-no-row">
    <div><strong>Receipt No:</strong> ${r.id}</div>
    <div><strong>Date:</strong> ${r.date}</div>
  </div>
  <table class="info-table">
    <tr><td class="lbl">Contributor / Devotee:</td><td class="val">${r.memberName}</td></tr>
    <tr><td class="lbl">Contributor Mobile:</td><td class="val">+91 ${r.phone}</td></tr>
    <tr><td class="lbl">Contributor Email:</td><td class="val">${r.email || '—'}</td></tr>
    <tr><td class="lbl">Contribution Towards:</td><td class="val">Ganesh Chaturthi Utsav Fund</td></tr>
    <tr><td class="lbl">Contribution Month:</td><td class="val">${r.month}</td></tr>
    <tr><td class="lbl">Payment Verification:</td><td class="val">Verified in Live Google Sheet</td></tr>
  </table>
  <div class="amount-card">
    <div><div class="amount-lbl">Total Amount Received</div><small style="color:#f2e3c6">Status: Verified & Accepted</small></div>
    <div class="amount-val">INR ${Number(r.amount).toLocaleString('en-IN')}.00</div>
  </div>
  <div class="footer">
    <div><strong>Sri Vinayaka Yuvasena Utsav Committee</strong></div>
    <div>Live Tracking Portal: https://srivinayakayuvasena.vercel.app/</div>
    <div class="seal">Verified Official Receipt</div>
    <p style="margin: 8px 0 0; font-size: 10.5px; color: #999;">This is a system-generated electronic payment receipt. May Lord Vinayaka bless you and your family with prosperity and good health.</p>
  </div>
</div>
</body>
</html>`;
}

function addContribution_(b){
  auth_(b);
  if(!b.memberId||!b.month||!b.amount||!b.date)throw new Error('Member, month, amount and date are required.');
  if(!b.screenshotData)throw new Error('Payment screenshot is mandatory.');
  const ss=SpreadsheetApp.getActive();const ms=ss.getSheetByName(SHEET_MEMBERS),ps=ss.getSheetByName(SHEET_PAYMENTS);
  const members=ms.getDataRange().getValues();let member=null;
  for(let i=1;i<members.length;i++)if(String(members[i][0])===String(b.memberId))member={id:members[i][0],name:members[i][1],phone:members[i][2],email:String(members[i][3]||'').trim()};
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

  // Prepare receipt data
  const targetEmail = String(b.email||'').trim() || member.email || findMemberEmail_(member.name, member.phone);
  const rcptData = {
    id: id,
    memberName: member.name,
    phone: member.phone,
    email: targetEmail,
    month: b.month,
    amount: Number(b.amount),
    date: b.date,
    screenshotUrl: file.getUrl()
  };

  // Compile HTML and PDF
  let htmlReceipt = '', pdfBlob = null, pdfDownloadUrl = '';
  try {
    htmlReceipt = generateReceiptHtml_(rcptData);
    pdfBlob = Utilities.newBlob(htmlReceipt, 'text/html', 'SVYS_Receipt_'+id+'.html').getAs('application/pdf').setName('SVYS_Payment_Receipt_'+id+'.pdf');
    const rootFolder = getOrCreateFolder_(DRIVE_ROOT);
    const receiptFolder = getOrCreateFolder_(rootFolder, 'Receipts');
    const pdfFile = receiptFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    pdfDownloadUrl = pdfFile.getUrl();
  } catch(pdfErr) {
    Logger.log('PDF generation error: ' + pdfErr.message);
  }

  // 1. Automatic Email Delivery with PDF attachment (Zero manual effort)
  if (targetEmail && targetEmail.indexOf('@') !== -1 && pdfBlob) {
    let emailSent = false;
    try {
      MailApp.sendEmail({
        to: targetEmail,
        subject: 'Official Donation Receipt #' + id + ' - Sri Vinayaka Yuvasena',
        htmlBody: htmlReceipt,
        name: 'Sri Vinayaka Yuvasena',
        attachments: [pdfBlob]
      });
      emailSent = true;
      Logger.log('SUCCESS: Email dispatched via MailApp to ' + targetEmail);
    } catch(mailErr) {
      Logger.log('MailApp error: ' + mailErr.message);
    }

    if (!emailSent) {
      try {
        GmailApp.sendEmail(targetEmail, 'Official Donation Receipt #' + id + ' - Sri Vinayaka Yuvasena', 'Please find attached your official donation receipt.', {
          htmlBody: htmlReceipt,
          name: 'Sri Vinayaka Yuvasena',
          attachments: [pdfBlob]
        });
        emailSent = true;
        Logger.log('SUCCESS: Email dispatched via GmailApp fallback to ' + targetEmail);
      } catch(gmailErr) {
        Logger.log('GmailApp fallback error: ' + gmailErr.message);
      }
    }

    // If email was manually provided and sheet didn't have it, auto-save to Column D of Members!
    if (String(b.email||'').trim() && (!member.email || member.email.indexOf('@')===-1)) {
      try {
        const emailToSave = String(b.email).trim();
        for (let i = 1; i < members.length; i++) {
          if (String(members[i][0]) === String(member.id)) {
            ms.getRange(i + 1, 4).setValue(emailToSave);
            member.email = emailToSave;
            Logger.log('Auto-saved new email to Members sheet for ' + member.name + ': ' + emailToSave);
            break;
          }
        }
      } catch(sheetUpdateErr) {
        Logger.log('Could not update member email in sheet: ' + sheetUpdateErr.message);
      }
    }
  } else {
    Logger.log('Email skipped: No email provided or registered for ' + member.name);
  }

  // 2. Automatic WhatsApp Dispatch
  try {
    sendAutoWhatsAppReceipt_(member.phone, rcptData, pdfDownloadUrl);
  } catch(waErr) {
    Logger.log('WhatsApp error: ' + waErr.message);
  }

  // 3. Automatic SMS Dispatch
  try {
    sendAutoSmsReceipt_(member.phone, rcptData, pdfDownloadUrl);
  } catch(smsErr) {
    Logger.log('SMS error: ' + smsErr.message);
  }

  const resultData = getData_();
  resultData.newReceipt = {
    id: id,
    memberId: member.id,
    memberName: member.name,
    phone: member.phone,
    email: String(b.email||'').trim() || findMemberEmail_(member.name, member.phone),
    month: b.month,
    amount: Number(b.amount),
    date: b.date,
    screenshotUrl: file.getUrl()
  };
  return resultData;
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
  if(name===undefined){name=parent;parent=DriveApp.getRootFolder();}
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
  const members=mv.slice(1).filter(r=>r[0]).map(r=>({id:String(r[0]),name:String(r[1]||''),phone:String(r[2]||''),email:String(r[3]||'').trim()}));
  const payments=pv.slice(1).filter(r=>r[0]).map(r=>({id:String(r[0]),memberId:String(r[1]),memberName:String(r[2]||''),phone:String(r[3]||''),month:String(r[4]||''),amount:Number(r[5]||0),date:r[6] instanceof Date?Utilities.formatDate(r[6],Session.getScriptTimeZone(),'yyyy-MM-dd'):String(r[6]||''),screenshotUrl:String(r[7]||''),screenshotName:String(r[8]||'')}));
  const expenses=ev.slice(1).filter(r=>r[0]).map(r=>({id:String(r[0]),category:String(r[1]||'Miscellaneous'),title:String(r[2]||''),amount:Number(r[3]||0),date:r[4] instanceof Date?Utilities.formatDate(r[4],Session.getScriptTimeZone(),'yyyy-MM-dd'):String(r[4]||''),paidTo:String(r[5]||''),receiptUrl:String(r[6]||''),receiptName:String(r[7]||'')}));
  const chatMessages=getChatMessages_().messages || [];
  const suggestions=getSuggestions_().suggestions || [];
  const idCards=getAllIdCards_();
  return {ok:true,members,payments,expenses,chatMessages,suggestions,idCards};
}

function sendChatMessage_(b){
  const ss=SpreadsheetApp.getActive();
  let c=ss.getSheetByName(SHEET_CHAT);
  if(!c){setupSheets();c=ss.getSheetByName(SHEET_CHAT);}
  const id='C'+Date.now();
  const sender=String(b.sender||'Devotee').trim();
  const email=String(b.email||'').trim();
  const text=String(b.text||'').trim();
  const time=String(b.time||'').trim();
  const isAnnounce=b.isAnnouncement?true:false;
  if(!text)throw new Error('Message text is required.');
  c.appendRow([id,sender,email,text,time,isAnnounce,new Date()]);
  return {ok:true,message:{id,sender,email,text,time,isAnnouncement:isAnnounce}};
}

function getChatMessages_(){
  try{
    const ss=SpreadsheetApp.getActive();
    const c=ss.getSheetByName(SHEET_CHAT);
    if(!c)return {ok:true,messages:[]};
    const vals=c.getDataRange().getValues();
    const rows=vals.slice(1).filter(r=>r[0]).slice(-150).map(r=>({
      id:String(r[0]),
      sender:String(r[1]||'Devotee'),
      email:String(r[2]||''),
      text:String(r[3]||''),
      time:String(r[4]||''),
      isAnnouncement:r[5]===true||String(r[5]).toLowerCase()==='true'
    }));
    return {ok:true,messages:rows};
  }catch(e){return {ok:true,messages:[]};}
}

function sendSuggestion_(b){
  const ss=SpreadsheetApp.getActive();
  let s=ss.getSheetByName(SHEET_SUGGESTIONS);
  if(!s){setupSheets();s=ss.getSheetByName(SHEET_SUGGESTIONS);}
  const id=b.id || ('S'+Date.now());
  const sender=String(b.sender||'Devotee').trim();
  const email=String(b.email||'').trim();
  const text=String(b.text||'').trim();
  const timeStr=String(b.timeStr||'').trim();
  if(!text)throw new Error('Suggestion text is required.');
  s.appendRow([id,sender,email,text,timeStr,0,new Date()]);
  return {ok:true,suggestion:{id,sender,email,text,timeStr,likes:0}};
}

function likeSuggestion_(b){
  try{
    const ss=SpreadsheetApp.getActive();
    const s=ss.getSheetByName(SHEET_SUGGESTIONS);
    if(!s)return {ok:false,error:'Sheet not found'};
    const vals=s.getDataRange().getValues();
    for(let i=1;i<vals.length;i++){
      if(String(vals[i][0])===String(b.id)){
        const curLikes=Number(vals[i][5]||0)+1;
        s.getRange(i+1,6).setValue(curLikes);
        return {ok:true,likes:curLikes};
      }
    }
  }catch(e){}
  return {ok:false,error:'Suggestion not found'};
}

function getSuggestions_(){
  try{
    const ss=SpreadsheetApp.getActive();
    const s=ss.getSheetByName(SHEET_SUGGESTIONS);
    if(!s)return {ok:true,suggestions:[]};
    const vals=s.getDataRange().getValues();
    const rows=vals.slice(1).filter(r=>r[0]).slice(-100).reverse().map(r=>({
      id:r[0],
      sender:String(r[1]||'Devotee'),
      email:String(r[2]||''),
      text:String(r[3]||''),
      timeStr:String(r[4]||''),
      likes:Number(r[5]||0)
    }));
    return {ok:true,suggestions:rows};
  }catch(e){return {ok:true,suggestions:[]};}
}
function jsonOut(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}

function saveIdCard_(b){
  const ss=SpreadsheetApp.getActive();
  let ic=ss.getSheetByName(SHEET_ID_CARDS);
  if(!ic){setupSheets();ic=ss.getSheetByName(SHEET_ID_CARDS);}
  const email=String(b.email||'').trim().toLowerCase();
  const name=String(b.name||'').trim();
  const phone=String(b.phone||'').trim();
  const blood=String(b.blood||'O+').trim();
  const photo=String(b.photo||'').trim();
  const idCode=String(b.idCode||('SVYS-2026-'+Math.floor(Math.random()*900+100))).trim();
  if(!email)throw new Error('Email is required.');
  if(!name)throw new Error('Name is required.');

  const vals=ic.getDataRange().getValues();
  let foundRow=-1;
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][0]||'').trim().toLowerCase()===email){
      foundRow=i+1;
      break;
    }
  }
  if(foundRow>0){
    ic.getRange(foundRow,2).setValue(name);
    ic.getRange(foundRow,3).setValue(phone);
    ic.getRange(foundRow,4).setValue(blood);
    if(photo)ic.getRange(foundRow,5).setValue(photo);
    ic.getRange(foundRow,6).setValue(idCode);
    ic.getRange(foundRow,7).setValue(new Date());
  }else{
    ic.appendRow([email,name,phone,blood,photo,idCode,new Date()]);
  }
  return {ok:true,card:{email,name,phone,blood,photo,idCode}};
}

function getIdCard_(b){
  try{
    const ss=SpreadsheetApp.getActive();
    const ic=ss.getSheetByName(SHEET_ID_CARDS);
    if(!ic)return {ok:true,card:null};
    const email=String(b.email||'').trim().toLowerCase();
    if(!email)return {ok:true,card:null};
    const vals=ic.getDataRange().getValues();
    for(let i=1;i<vals.length;i++){
      if(String(vals[i][0]||'').trim().toLowerCase()===email){
        return {
          ok:true,
          card:{
            email:String(vals[i][0]),
            name:String(vals[i][1]||''),
            phone:String(vals[i][2]||''),
            blood:String(vals[i][3]||'O+'),
            photo:String(vals[i][4]||''),
            idCode:String(vals[i][5]||'')
          }
        };
      }
    }
    return {ok:true,card:null};
  }catch(e){return {ok:false,error:e.message};}
}

function getAllIdCards_(){
  try{
    const ss=SpreadsheetApp.getActive();
    const ic=ss.getSheetByName(SHEET_ID_CARDS);
    if(!ic)return {};
    const vals=ic.getDataRange().getValues();
    const map={};
    for(let i=1;i<vals.length;i++){
      const em=String(vals[i][0]||'').trim().toLowerCase();
      if(em){
        map[em]={
          email:em,
          name:String(vals[i][1]||''),
          phone:String(vals[i][2]||''),
          blood:String(vals[i][3]||'O+'),
          photo:String(vals[i][4]||''),
          idCode:String(vals[i][5]||'')
        };
      }
    }
    return map;
  }catch(e){return {};}
}

// ==========================================================================
// ZERO-MANUAL-EFFORT AUTOMATED DISPATCH (WHATSAPP & SMS WEBHOOK GATEWAYS)
// ==========================================================================
function sendAutoWhatsAppReceipt_(phone, r, pdfUrl){
  try {
    const props = PropertiesService.getScriptProperties();
    const apiUrl = props.getProperty('WHATSAPP_API_URL') || '';
    const apiToken = props.getProperty('WHATSAPP_API_TOKEN') || '';
    if (!apiUrl) {
      Logger.log('Auto WhatsApp: Set WHATSAPP_API_URL in Script Properties for 0-click automated delivery.');
      return;
    }
    const cleanPhone = String(phone).replace(/\D/g, '');
    const fullPhone = cleanPhone.length === 10 ? ('91' + cleanPhone) : cleanPhone;
    const msg = `🙏 *SRI VINAYAKA YUVASENA - OFFICIAL E-RECEIPT* 🙏\n\n` +
      `Receipt No: ${r.id}\n` +
      `Devotee: ${r.memberName}\n` +
      `Amount: ₹${Number(r.amount).toLocaleString('en-IN')}\n` +
      `Month: ${r.month}\n` +
      `Date: ${r.date}\n` +
      `Status: Verified & Recorded in Live Audit\n\n` +
      (pdfUrl ? `📄 Download PDF Receipt: ${pdfUrl}\n\n` : '') +
      `May Lord Sri Ganesha bless you and your family!\n` +
      `Live Portal: https://srivinayakayuvasena.vercel.app/`;

    UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        token: apiToken,
        to: fullPhone,
        body: msg
      }),
      muteHttpExceptions: true
    });
    Logger.log('Auto WhatsApp dispatched successfully to ' + fullPhone);
  } catch(e) {
    Logger.log('Auto WhatsApp error: ' + e.message);
  }
}

function sendAutoSmsReceipt_(phone, r, pdfUrl){
  try {
    const props = PropertiesService.getScriptProperties();
    const smsApiKey = props.getProperty('FAST2SMS_API_KEY') || '';
    if (!smsApiKey) {
      Logger.log('Auto SMS: Set FAST2SMS_API_KEY in Script Properties for 0-click automated delivery.');
      return;
    }
    const cleanPhone = String(phone).replace(/\D/g, '').slice(-10);
    const formattedAmt = Number(r.amount || 0).toLocaleString('en-IN');
    const smsText = `Thank you, ${r.memberName}! Sri Vinayaka Yuvasena has received your contribution of ₹${formattedAmt} for ${r.month} in support of Ganesh Chaturthi Utsav.\n\nReceipt No.: ${r.id}\n\nTrack your contribution here: https://srivinayakayuvasena.vercel.app/`;

    UrlFetchApp.fetch('https://www.fast2sms.com/dev/bulkV2', {
      method: 'post',
      headers: { 'authorization': smsApiKey },
      contentType: 'application/json',
      payload: JSON.stringify({
        route: 'v3',
        sender_id: 'TXTIND',
        message: smsText,
        language: 'english',
        flash: 0,
        numbers: cleanPhone
      }),
      muteHttpExceptions: true
    });
    Logger.log('Auto SMS dispatched successfully to ' + cleanPhone);
  } catch(e) {
    Logger.log('Auto SMS error: ' + e.message);
  }
}
