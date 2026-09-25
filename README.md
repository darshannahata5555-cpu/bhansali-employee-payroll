# Bhansali Medisales Employee Payroll

A simple employee salary dashboard backed by Google Sheets.

Live dashboard: https://darshannahata5555-cpu.github.io/bhansali-employee-payroll/

## Run the dashboard

1. Run `npm install`.
2. Run `npm run dev`.
3. Open the local URL shown in the terminal.

The app opens in demo mode and stores changes in the browser until Google Sheets is connected. Once connected, Google Sheets becomes the source of truth: records load from the sheet at startup and changes are backed up automatically. Browser storage remains as an offline cache.

## Connect Google Sheets

1. Open the supplied **Bhansali Medisales — Employee Payroll** Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Replace the editor contents with `google-apps-script/Code.gs` from this project.
4. Run `setupWorkbook` once and approve Google access.
5. Open the execution log and copy the generated **BMS private access token**.
6. Select **Deploy → New deployment → Web app**.
7. Set **Execute as** to yourself. Deploy the web app.
8. Open **Settings** in the dashboard and enter the private token once. The deployed backend URL is already configured in the app.

The dashboard automatically backs up employees, attendance, expenses, and payroll after each change. The circular-arrow button can be used for an immediate manual sync.

## Salary rule

`Daily salary = Monthly salary ÷ Calendar days in the month`

`Net payable = Daily salary × Present days − Employee expenses`
