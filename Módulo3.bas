Attribute VB_Name = "Módulo3"
Sub ExportAppointmentsToExcel_RENATO()
    Const SCRIPT_NAME = "Export Appointments to Excel"
    Dim olkFld As Object, _
        olkLst As Object, _
        olkApt As Object, _
        excApp As Object, _
        excWkb As Object, _
        excWks As Object, _
        lngRow As Long, _
        intCnt As Integer
    Set olkFld = Application.ActiveExplorer.CurrentFolder
    
    Session.GetDefaultFolder (olFolderCalendar)
    
    If olkFld.DefaultItemType = olAppointmentItem Then
        'strFilename = InputBox("Enter a filename (including path) to save the exported appointments to.", SCRIPT_NAME)
        strFilename = "C:\Users\rmour\Desktop\calendario_export.xlsx"
        If strFilename <> "" Then
            Set excApp = CreateObject("Excel.Application")
            Set excWkb = excApp.Workbooks.Add()
            Set excWks = excWkb.Worksheets(1)
            'Write Excel Column Headers
            With excWks
                .Cells(1, 1) = "Organizer"
                .Cells(1, 2) = "Created"
                .Cells(1, 3) = "Subject"
                .Cells(1, 4) = "Start"
                .Cells(1, 5) = "Required"
                .Cells(1, 6) = "Optional"
            End With
            lngRow = 2
            Set olkLst = olkFld.Items
            olkLst.Sort "[Start]"
            olkLst.IncludeRecurrences = True
            'Write appointments to spreadsheet
            For Each olkApt In Application.ActiveExplorer.CurrentFolder.Items
                'Only export appointments
                If olkApt.Class = olAppointment Then
                    'Add a row for each field in the message you want to export
                    excWks.Cells(lngRow, 1) = olkApt.Organizer
                    excWks.Cells(lngRow, 2) = olkApt.CreationTime
                    excWks.Cells(lngRow, 3) = olkApt.Subject
                    excWks.Cells(lngRow, 4) = olkApt.Start
                    excWks.Cells(lngRow, 5) = olkApt.RequiredAttendees
                    excWks.Cells(lngRow, 6) = olkApt.OptionalAttendees
                    lngRow = lngRow + 1
                    intCnt = intCnt + 1
                End If
            Next
            excWks.Columns("A:F").AutoFit
            excWkb.SaveAs strFilename
            excWkb.Close
            MsgBox "Process complete.  A total of " & intCnt & " appointments were exported.", vbInformation + vbOKOnly, SCRIPT_NAME
        End If
    Else
        MsgBox "Operation cancelled.  The selected folder is not a calendar.  You must select a calendar for this macro to work.", vbCritical + vbOKOnly, SCRIPT_NAME
    End If
    Set excWks = Nothing
    Set excWkb = Nothing
    Set excApp = Nothing
    Set olkApt = Nothing
    Set olkLst = Nothing
    Set olkFld = Nothing
End Sub




Sub RunAllInboxRules2()
    Dim st As Outlook.Store
    Dim myRules As Outlook.Rules
    Dim rl As Outlook.Rule
    Dim count As Integer
    Dim ruleList As String
    'On Error Resume Next
    
    Call MarkAsUnread
     
    ' get default store (where rules live)
    Set st = Application.Session.DefaultStore
    ' get rules
    Set myRules = st.GetRules
     
    ' iterate all the rules
    For Each rl In myRules
        ' determine if it's an Inbox rule
        If rl.RuleType = olRuleReceive Then
            ' if so, run it
            rl.Execute ShowProgress:=True
            count = count + 1
            ruleList = ruleList & vbCrLf & rl.name
        End If
    Next
     
    ' tell the user what you did
    ruleList = "These rules were executed against the Inbox: " & vbCrLf & ruleList
    MsgBox ruleList, vbInformation, "Macro: RunAllInboxRules"
     
    Set rl = Nothing
    Set st = Nothing
    Set myRules = Nothing
End Sub


Sub MarkAsUnread()
'Application.ScreenUpdating = False

Dim objInbox As Outlook.MAPIFolder
Dim objOutlook As Object, objnSpace As Object, objMessage As Object

Set objOutlook = CreateObject("Outlook.Application")
Set objnSpace = objOutlook.GetNamespace("MAPI")
Set objInbox = objnSpace.GetDefaultFolder(olFolderInbox)

For Each objMessage In objInbox.Items
objMessage.UnRead = True
Next

Set objOutlook = Nothing
Set objnSpace = Nothing
Set objInbox = Nothing

'Application.ScreenUpdating = True
End Sub

