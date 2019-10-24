Attribute VB_Name = "UpdateConciExcel"
Option Explicit

 Private Const xlUp As Long = -4162
 Private Const xlValues As Long = -4163
 Private Const xlWhole As Long = 1
 Private Const xlPrevious As Long = 2
 Private Const xlByRows As Long = 1

 Sub AtualizaConciliadoraExcel(olItem As Outlook.MailItem)
 Dim enviro, strPath As String
 Dim xlApp As Object
 Dim bXStarted As Boolean
 Dim xlWB, xlSheet As Object
 
 Dim vText, vText2, vText3, vText4 As Variant
 Dim sText As String
 Dim rCount As Long
 
 
 Dim Reg1 As Object
 Dim M1 As Object
 Dim M As Object
              
    enviro = CStr(Environ("USERPROFILE"))
    
    strPath = enviro & "\Documents\Conciliadoras Opt-in.xlsx"
    
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err <> 0 Then
        Application.StatusBar = "Aguarde enquanto o Excel é aberto ..."
        Set xlApp = CreateObject("Excel.Application")
        bXStarted = True
    End If
    On Error GoTo 0
     
    'Apre a planilha para realizar a atualização
    Set xlWB = xlApp.Workbooks.Open(strPath)
    Set xlSheet = xlWB.Sheets("Opt-in")

    'Find the next empty line of the worksheet
     'rCount = xlSheet.Range("B" & xlSheet.Rows.Count).End(xlUp).Row
     'rCount = rCount + 1
     
     sText = olItem.Body

     Set Reg1 = CreateObject("VBScript.RegExp")
    ' \s* = invisible spaces
    ' \d* = match digits
    ' \w* = match alphanumeric
     
    With Reg1
        .Pattern = "(Razao:\s*([^)]+)\s*CNPJ:\s*(\d*)/(\d*)-(\d*))"
    End With
    If Reg1.test(sText) Then
     
    ' each "(\w*)" and the "(\d)" are assigned a vText variable
        Set M1 = Reg1.Execute(sText)
        For Each M In M1
           vText = Trim(M.SubMatches(1))
           vText2 = Trim(M.SubMatches(2))
           vText3 = Trim(M.SubMatches(3))
           vText4 = Trim(M.SubMatches(4))
        Next
    End If
    
    With xlSheet.Range("d1")
        rCount = .Find(What:=vText2 & "/" & vText3 & "-" & vText4, _
                        After:=.Cells(1), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
    End With

    'If Not rCount Is Nothing Then
        xlSheet.Range("e" & rCount) = Now
        xlSheet.Range("f" & rCount) = "Aprovado"
    'Else
    '    MsgBox "Não Localizado"
    'End If


     xlWB.Close 1
     If bXStarted Then
         xlApp.Quit
     End If
     Set M = Nothing
     Set M1 = Nothing
     Set Reg1 = Nothing
     Set xlApp = Nothing
     Set xlWB = Nothing
     Set xlSheet = Nothing
 End Sub

