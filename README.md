# Excel VBA Property Management System

This project is an Excel-based property management system designed to automate data entry, reporting, and communication tasks. It leverages VBA macros to provide an efficient and user-friendly workflow for managing real estate data. Below is a detailed breakdown of the system's features, functionalities, and implementation.

---

## Features

1. **Form Sheet (`FORM`)**  
   - A user-friendly interface for entering property details:
     - Property type
     - Property size
     - Location
     - Price
     - Completion date  
   - **Buttons on the form:**
     - **Confirm Entry**: Transfers data to the `DATA` sheet.  
       *(Code required:
        `[Sub ConfirmDataEntry()
    
    Dim wsForm As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim missingData As Boolean
    Dim namedRanges As Range
    
    
    Set wsForm = ThisWorkbook.Sheets("Form")
    Set wsData = ThisWorkbook.Sheets("Data")

    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1
    
   
    Set namedRanges = Union(wsForm.Range("tybe"), wsForm.Range("space"), wsForm.Range("location"), wsForm.Range("price"), wsForm.Range("date"))

   
    missingData = False
    For Each cell In namedRanges
        If cell.Value = "" Then
            missingData = True
            Exit For
        End If
    Next cell

    If missingData Then
        MsgBox "Please fill in all the required fields before submitting.", vbExclamation
        Exit Sub
    End If

   
    wsData.Cells(lastRow, 1).Value = wsForm.Range("tybe").Value ' نوع العقار
    wsData.Cells(lastRow, 2).Value = wsForm.Range("space").Value ' مساحة العقار
    wsData.Cells(lastRow, 3).Value = wsForm.Range("location").Value ' الموقع
    wsData.Cells(lastRow, 4).Value = wsForm.Range("price").Value ' السعر
    wsData.Cells(lastRow, 5).Value = wsForm.Range("date").Value ' تاريخ الانتهاء

    
    wsForm.Range("tybe,space,location,price,date").ClearContents

    ThisWorkbook.RefreshAll
   
    MsgBox "Data has been successfully transferred to the Data sheet!", vbInformation
End Sub
]`)*
     - **Print Report**: Generates and prints a report from the `DATA` sheet.  
       *(Code required: `[Sub PrintDataSheet()
    Dim dataSheet As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim printRange As Range

   
    On Error Resume Next
    Set dataSheet = ThisWorkbook.Sheets("Data")
    On Error GoTo 0


    If dataSheet Is Nothing Then
        MsgBox "The sheet 'Data' does not exist.", vbExclamation
        Exit Sub
    End If

    
    With dataSheet
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With

    Set printRange = dataSheet.Range(dataSheet.Cells(1, 1), dataSheet.Cells(lastRow, lastCol))

   
    With dataSheet.PageSetup
        .PrintArea = printRange.Address
    End With


    printRange.PrintOut

    
    MsgBox "The 'Data' sheet has been printed successfully!", vbInformation
End Sub
]`)*
     - **Print Dashboard**: Prints the dashboard from the `DASHBOARD` sheet.  
       *(Code required: `[Sub PrintDashboard()
    Dim dashboardSheet As Worksheet

    
    On Error Resume Next
    Set dashboardSheet = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0
    
    If dashboardSheet Is Nothing Then
        MsgBox "The sheet 'Dashboard' does not exist.", vbExclamation
        Exit Sub
    End If

  
    With dashboardSheet
        .PrintOut
    End With
    
    
    MsgBox "The 'Dashboard' sheet has been printed successfully!", vbInformation
End Sub
]`)*
     - **Send Screenshot**:  
       - Takes screenshots of the `DATA` and `DASHBOARD` sheets.  
       - Opens a web browser, searches for the phone number in the `SEND` sheet, and sends the screenshots via WhatsApp.  
       - Handles invalid phone numbers with an error message.  
       *(Code required: `[Sub SenddateWhatsAppMessage()
    Dim contact As String
    Dim sendSheet As Worksheet
    Dim dataSheet As Worksheet
    
  
    Set sendSheet = ThisWorkbook.Sheets("Send")
    Set dataSheet = ThisWorkbook.Sheets("Data")
    
    
    contact = Trim(sendSheet.Cells(2, 2).Value)
    

    If contact = "" Then
        MsgBox "No phone number found in cell B2.", vbExclamation
        Exit Sub
    End If
    
   
    ThisWorkbook.FollowHyperlink "https://web.whatsapp.com/"
    Application.Wait (Now + TimeValue("00:00:30")) ' Wait longer to ensure the page is fully loaded
    
    
    Dim i As Integer
    For i = 1 To 10
        Call SendKeys("{Tab}", True)
        Application.Wait (Now + TimeValue("00:00:01"))
    Next i
    
    
    Application.Wait (Now + TimeValue("00:00:03"))
    Call SendKeys(contact, True)
    Application.Wait (Now + TimeValue("00:00:02"))
    Call SendKeys("~", True) ' Press Enter key to search
    Application.Wait (Now + TimeValue("00:00:03"))
    
    
    dataSheet.Range("A2:E12").Copy
    Application.Wait (Now + TimeValue("00:00:02"))
    
    Application.SendKeys "^v"
    Application.Wait (Now + TimeValue("00:00:02"))
    
    
    Call SendKeys("~", True) ' Press Enter key to send
    Application.Wait (Now + TimeValue("00:00:02"))
    
    
    MsgBox "Message has been sent successfully!", vbInformation
End Sub


Sub SenddashboardWhatsAppMessage()
    Dim contact As String
    Dim sendSheet As Worksheet
    Dim dataSheet As Worksheet
    
   
    Set sendSheet = ThisWorkbook.Sheets("Send")
    Set dataSheet = ThisWorkbook.Sheets("Dashboard") ' Change the sheet name to "Dashboard"
    
    
    contact = Trim(sendSheet.Cells(2, 2).Value)
    
    
    If contact = "" Then
        MsgBox "No phone number found in cell B2.", vbExclamation
        Exit Sub
    End If
    
 
    ThisWorkbook.FollowHyperlink "https://web.whatsapp.com/"
    Application.Wait (Now + TimeValue("00:00:30")) ' Wait longer to ensure the page is fully loaded
    
  
    Dim i As Integer
    For i = 1 To 10
        Call SendKeys("{Tab}", True)
        Application.Wait (Now + TimeValue("00:00:01"))
    Next i
    
   
    Application.Wait (Now + TimeValue("00:00:03"))
    Call SendKeys(contact, True)
    Application.Wait (Now + TimeValue("00:00:02"))
    Call SendKeys("~", True) ' Press Enter key to search
    Application.Wait (Now + TimeValue("00:00:03"))
    
   
    dataSheet.Range("A4:T33").Copy ' Update the range to A4:T33
    Application.Wait (Now + TimeValue("00:00:02"))
    
    Application.SendKeys "^v"
    Application.Wait (Now + TimeValue("00:00:02"))
    
    
    Call SendKeys("~", True) ' Press Enter key to send
    Application.Wait (Now + TimeValue("00:00:02"))
    
   
    MsgBox "Message has been sent successfully!", vbInformation
End Sub

]`)*  

2. **Data Sheet (`DATA`)**  
   - A repository for all entered property data.  
   - Populated automatically through the form.  

3. **Pivot Table Sheet (`PIVOT TABLE`)**  
   - Contains pivot tables for advanced data analysis.
   - Automatically generated from the `DATA` sheet.  
     *(Code required: `[Insert code for pivot table creation here]`)*

4. **Dashboard Sheet (`DASHBOARD`)**  
   - Displays visual insights and summaries of the data.  
   - Designed for easy printing and sharing.  

5. **Send Sheet (`SEND`)**  
   - Stores a phone number for sending WhatsApp messages.
   - Updates dynamically to send to a new number when changed.

---

## How It Works

1. Open the Excel file and enable macros to activate the system.
2. Enter property details in the `FORM` sheet and click "Confirm Entry" to save the data.
3. Use the provided buttons to print reports, dashboards, or send screenshots via WhatsApp:
   - Screenshots are taken automatically, and the system opens the browser to send them.
   - Handles errors gracefully, ensuring seamless operation.

---

## Prerequisites

- Microsoft Excel (supports VBA macros)
- Internet connection for WhatsApp functionality
- Enabled macros in Excel settings

---

## Installation

1. Download the Excel macro-enabled file (`.xlsm`) and ensure macros are enabled.
2. Watch the following tutorial videos:
   - **Activation Video**: Explains how to enable and start the system.
   - **System Usage Video**: Demonstrates system functionalities.

---

## Code Details

All VBA macros are embedded within the file. For customization or troubleshooting, the VBA editor (Alt + F11) can be accessed to view or modify the code.


---

## Project Files

- **Excel File**: The macro-enabled workbook containing the system.
- **Videos**:
  - `activation.mp4`: Activation tutorial.
  - `usage.mp4`: System demonstration.

---

