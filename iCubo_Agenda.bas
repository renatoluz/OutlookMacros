Attribute VB_Name = "iCubo_Agenda"
Public Sub SendPrettyAgenda()
Dim oNamespace As NameSpace
Dim oFolder As Folder
Dim oCalendarSharing As CalendarSharing
Dim objMail As MailItem
Dim wd As Integer

Set oNamespace = Application.GetNamespace("MAPI")
Set oFolder = oNamespace.GetDefaultFolder(olFolderCalendar)
Set oCalendarSharing = oFolder.GetCalendarExporter

' get the day - send sat/sun/monday out Fri night
' Sun = 1, Mon = 2, Tue = 3, Wed = 4, Thu = 5, Fri = 6, Sat = 7
' none set Sat/Sun
wd = Weekday(Date)
If wd >= 2 And wd <= 7 Then
    lDays = Date + 1
ElseIf wd = 1 Then
    lDays = Date + 7
End If

With oCalendarSharing
' options are olFreeBusyAndSubject, olFullDetails, olFreeBusyOnly
    .CalendarDetail = olFreeBusyAndSubject
    .IncludeWholeCalendar = False
    .IncludeAttachments = False
    .IncludePrivateDetails = True
    .RestrictToWorkingHours = False
    .StartDate = Date + 1
    .EndDate = lDays
End With

' prepare as email
' options: olCalendarMailFormatEventList, olCalendarMailFormatDailySchedule
Set objMail = oCalendarSharing.ForwardAsICal(olCalendarMailFormatDailySchedule)
 
 ' Send the mail item to the specified recipient.
 With objMail
 .Recipients.Add "renato@icubo.digital"
 .Subject = "*****Agenda de " & Format(Now + 1, "DD/MM/YYYY")

' Remove the attached ics
 .Attachments.Remove (1)
 .Display 'for testing, change to .send
 End With

Set oCalendarSharing = Nothing
Set oFolder = Nothing
Set oNamespace = Nothing
End Sub



Sub CreateListofAppt()
   
   Dim CalFolder As Outlook.MAPIFolder
   Dim CalItems As Outlook.Items
   Dim ResItems As Outlook.Items
   Dim sFilter, strSubject, strAppt As String
   Dim iNumRestricted As Integer
   Dim itm, apptSnapshot As Object
   Dim tStart As Date, tEnd As Date, tFullWeek As Date
   Dim wd As Integer
  
   ' Use the default calendar folder
   Set CalFolder = Session.GetDefaultFolder(olFolderCalendar)
   Set CalItems = CalFolder.Items

   ' Sort all of the appointments based on the start time
   CalItems.Sort "[Start]"
   CalItems.IncludeRecurrences = True

   ' Set an end date
    tStart = Format(Date + 1, "Short Date")
    tEnd = Format(Date + 2, "Short Date")
    tFullWeek = Format(Date + 6, "Short Date")
 
    wd = Weekday(Date)
   ' Sun = 1, Mon = 2, Tues = 3, Wed = 4, Thu = 5, Fri = 6, Sat = 7
' get next day appt, do whole week on sunday
If wd >= 2 And wd <= 6 Then
   sFilter = "[Start] >= '" & tStart & "' AND [Start] <= '" & tEnd & "'"
ElseIf wd = 1 Then
   sFilter = "[Start] >= '" & tStart & "' AND [Start] <= '" & tFullWeek & "'"
End If

Debug.Print sFilter
   Set ResItems = CalItems.Restrict(sFilter)

   iNumRestricted = 0

   'Loop through the items in the collection.
   For Each itm In ResItems
   Debug.Print ResItems.count
      iNumRestricted = iNumRestricted + 1
      
 ' Create list of appointments
  strAppt = strAppt & vbCrLf & itm.Subject & vbTab & " >> " & vbTab & itm.Start & vbTab & " to: " & vbTab & Format(itm.End, "h:mm AM/PM")

   Next
   
' After the last occurrence is checked
' Open a new email message form and insert the list of dates
  Set apptSnapshot = Application.CreateItem(olMailItem)
  With apptSnapshot
    .Body = strAppt & vbCrLf & "Total appointments; " & iNumRestricted
    .To = "luz.renato@gmail.com, renato@icubo.digital"
    .Subject = "Appointments for " & tStart
    .Display 'or .send
  End With

   Set itm = Nothing
   Set apptSnapshot = Nothing
   Set ResItems = Nothing
   Set CalItems = Nothing
   Set CalFolder = Nothing
   
End Sub


Private Sub Application_Reminder(ByVal Item As Object)
'IPM.TaskItem to watch for Task Reminders
If Item.MessageClass <> "IPM.Appointment" Then
  Exit Sub
End If

If Item.Categories <> "Send Message" Then
  Exit Sub
End If

' call the macro:
SendPrettyAgenda

' or
' CreateListofAppt

End Sub



