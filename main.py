import win32com.client
import pandas as pd
from datetime import datetime as dt, timedelta as td
from config import XLSX_FILE


def send_reminder(reminder_list: list) -> None:
    """
    This is the function that sends the reminder email. It uses the win32com module and as you can see below you specify the email, subject and
    body and then it sends the email, we're currently sending this to helpdesk for centralisation but this can be changed if required.
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "MOT Reminders"
    mail.To = ""
    mail.Body = "The following MOTs are due within the next 30 days:\n\n" + "\n".join(
        reminder_list
    )
    mail.Send()
    print("Reminder email sent.")


def mot_reminder(file_path: str = XLSX_FILE) -> None:
    """
    This function has a file path parameter passed in but it uses a constant. You don't need to pass anything in when you call it unless you want to change
    which path is being used, maybe for testing for example. We iterate through the excel document using pandas. We're targetting the last MOT Completion
    date column so if index is equal to 0 we just skip and continue as the first row is the headers. We then read the dates in the column ([15]) and
    if it is coming up in the next 30 days, we append the computer name to the reminders list and then the date is formatted nicely. If there is reminders
    it calls the send reminders function and passes the reminders list in. This sends the email. Otherwise it prints that no mot's are due within the
    next 30 days.
    """
    df = pd.read_excel(file_path, engine="openpyxl")
    today = dt.today()
    reminder_threshold = today + td(days=30)
    reminders = []
    for index, row in df.iterrows():
        if index == 0:
            continue
        next_due_date = row[15]
        if pd.notna(next_due_date) and today <= next_due_date <= reminder_threshold:
            computer_name = row[0]
            reminders.append(
                f"{computer_name}: MOT due on {next_due_date.strftime('%Y-%m-%d')}"
            )
    if reminders:
        send_reminder(reminders)
    else:
        print("No MOTs due within the next 30 days.")


if __name__ == "__main__":
    mot_reminder()
