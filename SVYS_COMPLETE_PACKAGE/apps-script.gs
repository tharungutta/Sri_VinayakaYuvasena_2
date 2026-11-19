function authorizePermissions(){
  // Run this function once in Apps Script and click "Allow" to enable Email sending
  const quota = MailApp.getRemainingDailyQuota();
  Logger.log('SUCCESS: MailApp is authorized! Remaining daily emails: ' + quota);
  return 'SUCCESS: MailApp is authorized! Daily quota: ' + quota;
}

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
    try {
      MailApp.sendEmail({
        to: targetEmail,
        subject: 'Official Donation Receipt #' + id + ' - Sri Vinayaka Yuvasena',
        htmlBody: htmlReceipt,
        name: 'Sri Vinayaka Yuvasena',
        attachments: [pdfBlob]
      });
      Logger.log('SUCCESS: Email dispatched to ' + targetEmail);
    } catch(mailErr) {
      Logger.log('MailApp error: ' + mailErr.message);
    }
  } else {
    Logger.log('Email skipped: No email registered for ' + member.name);
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
  return {ok:true,members,payments,expenses};
}
function jsonOut(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}

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
