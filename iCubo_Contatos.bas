Attribute VB_Name = "iCubo_Contatos"
Sub ExportOutlookAddressBook_NEW()
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
'Windows("Book1.xlsx").Activate
'ActiveWorkbook.ActiveSheet.Range("a1").Select
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


Sub contatos()

    Dim xlApp As Object
    Dim sourceWB
    Dim sourceWS
    
    Set xlApp = CreateObject("Excel.Application")
    
    With xlApp
        .Visible = True
        .EnableEvents = True
    End With
    
    strFile = "C:\Users\rmour\Desktop\PRODUTOS\CONTATOS_CLIENTES\CONTATOS.xlsm"  'Put your file path.
    
    Set sourceWB = xlApp.Workbooks.Open(strFile, , False, , , , , , , True)
    'Set sourceWH = sourceWB.Worksheets("Sheet1")
    sourceWB.Activate
    
End Sub




