Attribute VB_Name = "iCubo"
'Option Explicit

'https://wellsr.com/vba/2016/outlook/export-outlook-address-book-to-excel/

Sub ExportOutlookAddressBook()
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Exports your Outlook Address Book to Excel.
'NOTES: This macro runs on Excel.
' Add the Microsoft Outlook Reference library to the project to get this to run
'Application.ScreenUpdating = False
Dim olApp As Outlook.Application
Dim olNs As Outlook.NameSpace
Dim olAL As Outlook.AddressList
Dim olEntry As Outlook.AddressEntry
Set olApp = Outlook.Application
Set olNs = olApp.GetNamespace("MAPI")
Set olAL = olNs.AddressLists("Contatos") 'Change name if different contacts list name
ActiveWorkbook.ActiveSheet.Range("a1").Select
For Each olEntry In olAL.AddressEntries
     ' your looping code here
     ActiveCell.Value = olEntry.GetContact.FullName 'display name
     ActiveCell.Offset(0, 1).Value = olEntry.Address 'email address
     ActiveCell.Offset(0, 2).Value = olEntry.GetContact.MobileTelephoneNumber 'cell phone number
     ActiveCell.Offset(1, 0).Select
Next olEntry
Set olApp = Nothing
Set olNs = Nothing
Set olAL = Nothing
Application.ScreenUpdating = True
End Sub




'https://www.ozgrid.com/forum/forum/help-forums/excel-general/113789-macro-to-export-outlook-calendar-to-excel


Sub ListAppointments()

    Dim olApp As Object
    Dim olNs As Object
    Dim olFolder As Object
    Dim olApt As Object
    Dim NextRow As Long
    
    Set olApp = CreateObject("Outlook.Application")
    
    Set olNs = olApp.GetNamespace("MAPI")
    
    Set olFolder = olNs.GetDefaultFolder(9) 'olFolderCalendar
    
    Range("A1:D1").Value = Array("Subject", "Start", "End", "Location")
    
    NextRow = 2
    
    For Each olApt In olFolder.Items
        Cells(NextRow, "A").Value = olApt.Subject
        Cells(NextRow, "B").Value = olApt.Start
        Cells(NextRow, "C").Value = olApt.End
        Cells(NextRow, "D").Value = olApt.Location
        NextRow = NextRow + 1
    Next olApt
    
    Set olApt = Nothing
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
    
    Columns.AutoFit
    
End Sub




