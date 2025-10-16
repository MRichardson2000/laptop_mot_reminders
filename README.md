💻 Laptop MOT Reminder Tool
This script automates the process of identifying laptops due for an MOT within 30 days and sends reminder emails to the helpdesk team. It helps ensure MOT's are booked in
and not forgotten about. 

-- THIS IS ON EBSFS005 (PRTG Server) --

📦 Features
- Scans a XLSX file containing laptop MOT details with date column at the end.
- Identifies laptops due for a MOT within the next 30 days.
- Sends a summary email via Outlook to the helpdesk.

🧰 Requirements
- Python 3.8+
- pandas
- openpyxl
- pywin32 (for Outlook integration)


Install dependencies:
run uv sync to pull dependencies from the toml file



📁 File Structure
project/
│
├── main.py                # Main script
├── config.py              # Contains XLSX file path
└── README.md              # Documentation



⚙️ Configuration
Edit config.py to set your file path if you need to change them but it's a constant so should only need changing for testing:
XLSX_FILE = "K:/IT/Restricted/Ecology Network/devops_automation_mr/eol_laptop_reminders/eol_laptops.xlsx"




🚀 Usage
Set up a batch file with the below, put your user path in and then specify on your c drive where you cloned the repo. This is just where most of my stuff went:
C:\Users\YOURUSERPATHHERE\AppData\Local\Programs\Python\Python313\python.exe C:\Utilities\Python\eol_laptop_reminders\main.py
Then set up a schedule task to run once a week on your chosen day and time that targets this batch file and runs it. 


This will:
- Read the laptop MOT file on the K drive.
- Check for laptops with MOT dates within 30 days.
- Send an email to the helpdesk with the list of upcoming MOT's due.

✉️ Email Setup
The script uses win32com.client to send emails via Outlook. The recipient is currently set to:
mail.To = "helpdesk@ecology.co.uk"


You can change this to any valid email address or distribution list.

🧪 Testing
To test with alternate files you just pass in your test file path. It uses the constant if nothings passed in:
mot_reminder(xlsx_path="test_inventory.xlsx")



📌 Notes
- The script assumes the Last MOT date is in the 15th column of the spreadsheet.
- The first row is skipped (assumed to be headers).
- run a uv sync in the terminal to pull the dependencies from the toml
- if you need uv, run pip install uv in the terminal to add to your global scope, then run uv sync
