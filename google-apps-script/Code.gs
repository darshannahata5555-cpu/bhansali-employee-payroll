const SHEETS = {
  employees: ['Employee ID','Name','Phone','PAN','Designation','Monthly Salary','Status','Joining Date'],
  attendance: ['Employee ID','Date','Status'],
  expenses: ['Expense ID','Employee ID','Date','Category','Amount','Description'],
  payroll: ['Month','Employee ID','Days in Month','Days Worked','Daily Salary','Earned Salary','Deductions','Net Salary','Generated At']
};

function doGet() { return json_({ok:true, service:'Bhansali Medisales Payroll API'}); }

function data_() {
  return {
    employees: rows_(title_('employees')).map(r=>({id:String(r[0]),name:String(r[1]),phone:String(r[2]),pan:String(r[3]),designation:String(r[4]),salary:Number(r[5])||0,status:String(r[6]||'Active'),joined:date_(r[7])})),
    attendance: rows_(title_('attendance')).map(r=>({employeeId:String(r[0]),date:date_(r[1]),status:String(r[2])})),
    expenses: rows_(title_('expenses')).map(r=>({id:String(r[0]),employeeId:String(r[1]),date:date_(r[2]),category:String(r[3]),amount:Number(r[4])||0,description:String(r[5])})),
    payroll: rows_(title_('payroll')).map(r=>({month:String(r[0]),employeeId:String(r[1]),days:Number(r[2]),worked:Number(r[3]),daily:Number(r[4]),gross:Number(r[5]),deduct:Number(r[6]),net:Number(r[7])}))
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
      writePayroll_(body.data);
      return json_({ok:true, syncedAt:new Date().toISOString()});
    }
    return json_({ok:false, error:'Unknown action'});
  } catch (err) { return json_({ok:false, error:String(err)}); }
}

function setupWorkbook() {
  Object.keys(SHEETS).forEach(key => sheet_(title_(key), SHEETS[key]));
  const props=PropertiesService.getScriptProperties();
  let token=props.getProperty('BMS_API_TOKEN');
  if(!token){token=Utilities.getUuid().replace(/-/g,'');props.setProperty('BMS_API_TOKEN',token);}
  Logger.log('BMS private access token: '+token);
}

function writeEmployees_(rows) { replace_(title_('employees'), SHEETS.employees, rows.map(x=>[x.id,x.name,x.phone,x.pan,x.designation,Number(x.salary),x.status,x.joined])); }
function writeAttendance_(rows) { replace_(title_('attendance'), SHEETS.attendance, rows.map(x=>[x.employeeId,x.date,x.status])); }
function writeExpenses_(rows) { replace_(title_('expenses'), SHEETS.expenses, rows.map(x=>[x.id,x.employeeId,x.date,x.category,Number(x.amount),x.description])); }
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
function replace_(name, headers, rows) { const s=sheet_(name,headers); s.clearContents(); s.getRange(1,1,1,headers.length).setValues([headers]); if(rows.length)s.getRange(2,1,rows.length,headers.length).setValues(rows); s.setFrozenRows(1); s.autoResizeColumns(1,headers.length); }
function sheet_(name,headers) { const ss=SpreadsheetApp.getActive(); let s=ss.getSheetByName(name); if(!s)s=ss.insertSheet(name); if(s.getLastRow()===0){s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#eeeeee');s.setFrozenRows(1);} return s; }
function rows_(name) { const s=SpreadsheetApp.getActive().getSheetByName(name); if(!s||s.getLastRow()<2)return []; return s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).getValues().filter(r=>String(r[0]).trim()!==''); }
function date_(value) { if(!value)return ''; if(Object.prototype.toString.call(value)==='[object Date]')return Utilities.formatDate(value,Session.getScriptTimeZone(),'yyyy-MM-dd'); return String(value).slice(0,10); }
function validToken_(token) { const expected=PropertiesService.getScriptProperties().getProperty('BMS_API_TOKEN'); return Boolean(expected)&&token===expected; }
function title_(key) { return key.charAt(0).toUpperCase()+key.slice(1); }
function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
