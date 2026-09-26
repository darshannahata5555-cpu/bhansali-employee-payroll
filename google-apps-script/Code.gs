const SHEETS = {
  employees: ['Employee ID','Name','Phone','PAN','Designation','Monthly Salary','Status','Joining Date','Email'],
  attendance: ['Employee ID','Date','Status'],
  expenses: ['Expense ID','Employee ID','Date','Category','Amount','Description'],
  employeeStoreExpenses: ['Expense ID','Employee ID','Date','Category','Amount','Description'],
  ownerStoreExpenses: ['Expense ID','Date','Category','Amount','Description'],
  payroll: ['Month','Employee ID','Days in Month','Days Worked','Daily Salary','Earned Salary','Deductions','Net Salary','Generated At']
};

function doGet() { return json_({ok:true, service:'Bhansali Medisales Payroll API'}); }

function data_() {
  return {
    employees: records_(title_('employees')).map(r=>({id:text_(r['Employee ID']),name:text_(r['Name']),phone:text_(r['Phone']),email:text_(r['Email']),pan:text_(r['PAN']),designation:text_(r['Designation']),salary:Number(r['Monthly Salary'])||0,status:text_(r['Status'])||'Active',joined:date_(r['Joining Date'])})),
    attendance: records_(title_('attendance')).map(r=>({employeeId:text_(r['Employee ID']),date:date_(r['Date']),status:text_(r['Status'])})),
    expenses: records_(title_('expenses')).map(r=>({id:text_(r['Expense ID']),employeeId:text_(r['Employee ID']),date:date_(r['Date']),category:text_(r['Category']),amount:Number(r['Amount'])||0,description:text_(r['Description'])})),
    employeeStoreExpenses: records_(title_('employeeStoreExpenses')).map(r=>({id:text_(r['Expense ID']),employeeId:text_(r['Employee ID']),date:date_(r['Date']),category:text_(r['Category']),amount:Number(r['Amount'])||0,description:text_(r['Description'])})),
    ownerStoreExpenses: records_(title_('ownerStoreExpenses')).map(r=>({id:text_(r['Expense ID']),date:date_(r['Date']),category:text_(r['Category']),amount:Number(r['Amount'])||0,description:text_(r['Description'])})),
    payroll: records_(title_('payroll')).map(r=>({month:text_(r['Month']),employeeId:text_(r['Employee ID']),days:Number(r['Days in Month'])||0,worked:Number(r['Days Worked'])||0,daily:Number(r['Daily Salary'])||0,gross:Number(r['Earned Salary'])||0,deduct:Number(r['Deductions'])||0,net:Number(r['Net Salary'])||0}))
  };
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (!validToken_(body.token)) return json_({ok:false, error:'Invalid access token'});
    if (body.action === 'load') return json_({ok:true, data:data_()});
    if (body.action === 'syncAll') {
      writeEmployees_(body.data.employees || []);
      writeAttendance_(body.data.attendance || []);
      writeExpenses_(body.data.expenses || []);
      writeEmployeeStoreExpenses_(body.data.employeeStoreExpenses || []);
      writeOwnerStoreExpenses_(body.data.ownerStoreExpenses || []);
      writePayroll_(body.data);
      return json_({ok:true, syncedAt:new Date().toISOString()});
    }
    return json_({ok:false, error:'Unknown action'});
  } catch (err) { return json_({ok:false, error:String(err)}); }
}

function setupWorkbook() {
  const currentEmployees = data_().employees;
  Object.keys(SHEETS).forEach(key => sheet_(title_(key), SHEETS[key]));
  writeEmployees_(currentEmployees);
  Object.keys(SHEETS).forEach(key => formatSheet_(sheet_(title_(key), SHEETS[key]), SHEETS[key].length));
  const props=PropertiesService.getScriptProperties();
  let token=props.getProperty('BMS_API_TOKEN');
  if(!token){token=Utilities.getUuid().replace(/-/g,'');props.setProperty('BMS_API_TOKEN',token);}
  Logger.log('BMS private access token: '+token);
}

function writeEmployees_(rows) { replace_(title_('employees'), SHEETS.employees, rows.map(x=>[x.id,x.name,x.phone,x.pan,x.designation,Number(x.salary),x.status,x.joined,x.email||''])); }
function writeAttendance_(rows) { replace_(title_('attendance'), SHEETS.attendance, rows.map(x=>[x.employeeId,x.date,x.status])); }
function writeExpenses_(rows) { replace_(title_('expenses'), SHEETS.expenses, rows.map(x=>[x.id,x.employeeId,x.date,x.category,Number(x.amount),x.description])); }
function writeEmployeeStoreExpenses_(rows) { replace_(title_('employeeStoreExpenses'), SHEETS.employeeStoreExpenses, rows.map(x=>[x.id,x.employeeId,x.date,x.category,Number(x.amount),x.description])); }
function writeOwnerStoreExpenses_(rows) { replace_(title_('ownerStoreExpenses'), SHEETS.ownerStoreExpenses, rows.map(x=>[x.id,x.date,x.category,Number(x.amount),x.description])); }
function writePayroll_(data) {
  const month = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const days = new Date(Number(month.slice(0,4)), Number(month.slice(5,7)), 0).getDate();
  const rows = (data.employees||[]).filter(e=>e.status==='Active').map(e=>{
    const worked=(data.attendance||[]).filter(a=>a.employeeId===e.id&&a.date.indexOf(month)===0&&a.status==='Present').length;
    const deductions=(data.expenses||[]).filter(x=>x.employeeId===e.id&&x.date.indexOf(month)===0).reduce((s,x)=>s+Number(x.amount),0);
    const daily=Number(e.salary)/days, earned=daily*worked;
    return [month,e.id,days,worked,daily,earned,deductions,Math.max(0,earned-deductions),new Date()];
  });
  replace_(title_('payroll'), SHEETS.payroll, rows);
}
function replace_(name, headers, rows) { const s=sheet_(name,headers); s.clearContents(); s.getRange(1,1,1,headers.length).setValues([headers]); if(rows.length)s.getRange(2,1,rows.length,headers.length).setValues(rows); formatSheet_(s,headers.length); }
function sheet_(name,headers) { const ss=SpreadsheetApp.getActive(); let s=ss.getSheetByName(name); if(!s)s=ss.insertSheet(name); if(s.getLastRow()===0)s.getRange(1,1,1,headers.length).setValues([headers]); return s; }
function records_(name) { const s=SpreadsheetApp.getActive().getSheetByName(name); if(!s||s.getLastRow()<2)return []; const width=s.getLastColumn(); const headers=s.getRange(1,1,1,width).getDisplayValues()[0]; return s.getRange(2,1,s.getLastRow()-1,width).getValues().filter(r=>String(r[0]).trim()!=='').map(row=>headers.reduce((obj,h,i)=>{obj[h]=row[i];return obj;},{})); }
function formatSheet_(s,width) { s.setFrozenRows(1); s.getRange(1,1,1,width).setFontWeight('bold').setBackground('#00796b').setFontColor('#ffffff'); s.autoResizeColumns(1,width); }
function date_(value) { if(!value)return ''; if(Object.prototype.toString.call(value)==='[object Date]')return Utilities.formatDate(value,Session.getScriptTimeZone(),'yyyy-MM-dd'); return String(value).slice(0,10); }
function text_(value) { return value==null?'':String(value); }
function validToken_(token) { const expected=PropertiesService.getScriptProperties().getProperty('BMS_API_TOKEN'); return Boolean(expected)&&token===expected; }
function title_(key) { return key.replace(/([A-Z])/g,' $1').replace(/^./,c=>c.toUpperCase()); }
function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
