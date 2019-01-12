Attribute VB_Name = "iCubo_Inventario"
'Function openExcel()
'
'Dim xlApp As Excel.Application
'Dim sourceWB As Workbook
'Dim sourceWS As Worksheet
'
'Set xlApp = New Excel.Application
'
'With xlApp
'.Visible = True
'.EnableEvents = False
'End With
'
'
'strFile = "C:\Users\rmour\Desktop\PRODUTOS\INVENTARIO\MERCADO.xlsx"  'Put your file path.
'
'Set sourceWB = Workbooks.Open(strFile, , False, , , , , , , True)
'Set sourceWH = sourceWB.Worksheets("SalesForm")
'sourceWB.Activate
'End Function

Sub Mercado()
Call Inventario
End Sub


Public Function Inventario()

    Dim xlApp As Object
    Dim sourceWB
    Dim sourceWS
    
    Set xlApp = CreateObject("Excel.Application")
    
    With xlApp
        .Visible = True
        .EnableEvents = True
    End With
    
    strFile = "C:\Users\rmour\Desktop\PRODUTOS\INVENTARIO\MERCADO.xlsm"  'Put your file path.
    
    Set sourceWB = xlApp.Workbooks.Open(strFile, , False, , , , , , , True)
    'Set sourceWH = sourceWB.Worksheets("Sheet1")
    sourceWB.Activate
    
End Function




'Sub Mercado()
'Call Inventario
'End Sub


Sub Prontuario_Pacientes()

    Dim xlApp As Object
    Dim sourceWB
    Dim sourceWS
    
    Set xlApp = CreateObject("Excel.Application")
    
    With xlApp
        .Visible = True
        .EnableEvents = True
    End With
    
    strFile = "C:\Users\rmour\Desktop\PRODUTOS\MEDICO_FINAL\prontuario_paciente.xlsm"  'Put your file path.
    
    Set sourceWB = xlApp.Workbooks.Open(strFile, , False, , , , , , , True)
    'Set sourceWH = sourceWB.Worksheets("Sheet1")
    sourceWB.Activate
    
End Sub


Sub Marketing_Pacientes()

    Dim xlApp As Object
    Dim sourceWB
    Dim sourceWS
    
    Set xlApp = CreateObject("Excel.Application")
    
    With xlApp
        .Visible = True
        .EnableEvents = True
    End With
    
    strFile = "C:\Users\rmour\Desktop\PRODUTOS\MEDICO_FINAL\iCubo_Client_Routine.xlsm"  'Put your file path.
    
    Set sourceWB = xlApp.Workbooks.Open(strFile, , False, , , , , , , True)
    'Set sourceWH = sourceWB.Worksheets("Sheet1")
    sourceWB.Activate
    
End Sub

Sub backup_files()
'MsgBox ("Deseja executar o backup?")
Answer = MsgBox("Deseja executar o backup?", vbYesNo + vbQuestion, "BACKUP")
If Answer = 6 Then
    Shell "cmd /c C:/Users/rmour/AppData/Local/Programs/Python/Python36-32/python.exe C:/Users/rmour/Desktop/PRODUTOS/BACKUP/backup.py"
End If
End Sub

Sub stop_backup_files()
'MsgBox ("Deseja executar o backup?")
Answer = MsgBox("Deseja PARAR o backup?", vbYesNo + vbQuestion, "PARAR O BACKUP")
If Answer = 6 Then
    Shell "cmd /c C:/Users/rmour/AppData/Local/Programs/Python/Python36-32/python.exe C:/Users/rmour/Desktop/PRODUTOS/BACKUP/CLOSE_BACKUP.py"
End If
End Sub




Sub financeiro()

    Dim xlApp As Object
    Dim sourceWB
    Dim sourceWS
    
    Set xlApp = CreateObject("Excel.Application")
    
    With xlApp
        .Visible = True
        .EnableEvents = True
    End With
    
    strFile = "C:\Users\rmour\Desktop\DINO\financas.xlsm"  'Put your file path.
    
    Set sourceWB = xlApp.Workbooks.Open(strFile, , False, , , , , , , True)
    'Set sourceWH = sourceWB.Worksheets("Sheet1")
    sourceWB.Activate
    
End Sub



Sub Missao()

    Dim xlApp As Object
    Dim sourceWB
    Dim sourceWS
    
    Set xlApp = CreateObject("Excel.Application")
    
    With xlApp
        .Visible = True
        .EnableEvents = True
    End With
    
    strFile = "C:\Users\rmour\Desktop\iCubo\MISSAO\MISSAO.xlsm"  'Put your file path.
    
    Set sourceWB = xlApp.Workbooks.Open(strFile, , False, , , , , , , True)
    'Set sourceWH = sourceWB.Worksheets("Sheet1")
    sourceWB.Activate
    
End Sub




Sub OpenFileTXT()

Shell "C:\Windows\Notepad.exe C:\Users\rmour\Desktop\DINO\MINIMALISMO.txt"

'Dim sFilePath As String
'Dim fileNumber As Integer
'' Full path of the textFile which you want
'' to open.
'sFilePath = "C:\Users\rmour\Desktop\DINO\MINIMALISMO.txt"
'
'' Assign a unique file numner
'fileNumber = FreeFile
'' Below statement will open the
'' above text file in output mode
'Open sFilePath For Output As #fileNumber
''Close

End Sub



Sub users()

    Dim xlApp As Object
    Dim sourceWB
    Dim sourceWS
    
    Set xlApp = CreateObject("Excel.Application")
    
    With xlApp
        .Visible = True
        .EnableEvents = True
    End With
    
    strFile = "C:\Users\rmour\Desktop\DINO\USERS.xlsm"  'Put your file path.
    
    Set sourceWB = xlApp.Workbooks.Open(strFile, , False, , , , , , , True)
    'Set sourceWH = sourceWB.Worksheets("Sheet1")
    sourceWB.Activate
    
End Sub


Sub People()

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



