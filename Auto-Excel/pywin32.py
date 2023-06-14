import win32com.client as win32
from win32com.client import constants
import os
from datetime import datetime
import time

def RGB(r, g, b):
    return r + (g * 256) + (b * 256 * 256)

def generate_email_summaries():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False

    wb = excel.Workbooks.Open('D:\\Data Analysis Projects\\VBA\\Income Table.xlsb')
    ws = wb.Sheets('Income Table')
    lastRow = ws.Cells(ws.Rows.Count, 1).End(constants.xlUp).Row

    outApp = win32.Dispatch('Outlook.Application')
    signature = outApp.CreateItem(0).HTMLBody

    for i in range(2, lastRow + 1):
        country = ws.Cells(i, 2).Value
        code = ws.Cells(i, 5).Value
        file_name = country + '.jpg'
        location_save = 'D:\\Data Analysis Projects\\VBA\\' + file_name

        wb_new = excel.Workbooks.Add()
        ws_new = wb_new.Sheets(1)

        ws_new.Range("B1:C1").Merge()
        ws_new.Range("B1").Value = country
        ws_new.Range("B1").Font.Bold = True
        ws_new.Range("B1").HorizontalAlignment = constants.xlCenter # -4108 is the constant value for xlCenter
        ws_new.Range("B1").Interior.Color = RGB(173, 216, 230)
        ws_new.Range("B3").Value = "ID"
        ws_new.Range("B4").Value = "Income"

        if code == 2:
            ws_new.Range("B6").Value = "Occupation"
            ws_new.Range("C3").Value = ws.Cells(i, 1).Value
            ws_new.Range("C4").Value = ws.Cells(i, 3).Value
            ws_new.Range("C6").Value = ws.Cells(i, 6).Value
            ws_new.Range("C3:C6").HorizontalAlignment = constants.xlRight # -4152 is the constant value for xlRight
            ws_new.Range("B6:C6").Interior.Color = RGB(224, 224, 224)
            ws_new.Range("B6:C6").Borders(constants.xlEdgeBottom).LineStyle = constants.xlContinuous
            ws_new.Range("B6:C6").Borders(constants.xlEdgeBottom).Weight = constants.xlThick
            ws_new.Columns.AutoFit()
        else:
            ws_new.Range("B6").Value = "Country"
            ws_new.Range("B7").Value = "Occupation"
            ws_new.Range("C3").Value = ws.Cells(i, 1).Value
            ws_new.Range("C4").Value = ws.Cells(i, 3).Value
            ws_new.Range("C6").Value = ws.Cells(i, 4).Value
            ws_new.Range("C7").Value = ws.Cells(i, 6).Value
            ws_new.Range("C3:C7").HorizontalAlignment = constants.xlRight  # -4152 is the constant value for xlRight
            ws_new.Range("B6:C7").Interior.Color = RGB(224, 224, 224)
            ws_new.Range("B6:C7").Borders(constants.xlEdgeBottom).LineStyle = constants.xlContinuous
            ws_new.Range("B6:C7").Borders(constants.xlEdgeBottom).Weight = constants.xlThick
            ws_new.Columns.AutoFit()

        ws_new.Range("B1:C" + str(6 if code == 2 else 7)).CopyPicture(Format=constants.xlPicture)

        objChart = ws_new.ChartObjects().Add(0, 0, ws_new.Range("B1:C" + str(6 if code == 2 else 7)).Width,
                                             ws_new.Range("B1:C" + str(6 if code == 2 else 7)).Height) # Chart is initially set to 0,0 so that it will be positioned at the top left corner of the ws.new workbooks
        objChart.Activate()
        objChart.Chart.Paste()
        objChart.Chart.Export(location_save)
        objChart.Delete()

        wb_new.Close(False)

        outApp = win32.Dispatch('Outlook.Application')
        # Get the updated HTML body after the email has been displayed
        outMail = outApp.CreateItem(0)



        # Display the email, allowing the default signature to be retained
        outMail.Display()
        signature = outMail.HTMLBody
        outMail.Close(0)


        # Year: %Y for a full four-digit year, %y for a two-digit year. For example, "2023" or "23".
        # Day of the month: %d for a zero-padded day of the month (01, 02, ..., 31).
        # Time: %H for hour and %M for minute and you can use them together like %H:%M:%S to get time in the format "13:45:30", for instance.

        current_month = datetime.now().strftime('%B')
        email_address = (ws.Cells(i, 2).Value.replace(" ", ".").lower() + "@vmail.com")

        str_body = "<BODY style=font-size:16pt;><p>Hiya " + ws.Cells(i, 2).Value + ";" + \
                   "<p>I hope you have been well and that this message finds you." + \
                   " I'd like to be selected for an interview with your organization as I believe I would be able to provide a unique insight and skills that the team and company would benefit from." + \
                   "<p> I am typically free to interview on Thursdays and on occasion Wednesdays." + \
                   " (username instead of link)." + \
                   "<p>I hope that we will be able to speak soon!</BODY>"

        outMail.To = email_address
        outMail.CC = ""
        outMail.Subject = ws.Cells(i, 2).Value + " - " + current_month + " VBA ChatGPT"




        # Append the signature to the email body
        outMail.HTMLBody = str_body + signature

        # Attach file and send the email
        outMail.Attachments.Add(location_save)

        outMail.Save()


    wb.Close(False)
    excel.Quit()

    ws = None
    wb = None
    excel = None
    outMail = None
    outApp = None

generate_email_summaries()