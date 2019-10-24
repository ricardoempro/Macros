Attribute VB_Name = "InsereConciCliExcel"
Option Explicit
 Private Const xlUp As Long = -4162
 Private Const xlValues As Long = -4163
 Private Const xlWhole As Long = 1
 Private Const xlPrevious As Long = 2
 Private Const xlByRows As Long = 1

 Sub InsereConcClienteExcel(item As Outlook.MailItem)
 Dim enviro, strPath As String
 Dim xlApp, xlWB, xlSheet, xlSheetConci, rangeConci, rangeCli As Object
 Dim rCount, rCountConci As Long
 Dim bXStarted As Boolean
 Dim ReglAprovacao, ReglInfoCliente, ReglInfoConcRazao, ReglInfoConcCnpj  As Object
 Dim razaoSocial, cnpjCliente, razaoConci, cnpjConci, aprovador, rgAprov, cpfAprov, conciAprov, cnpjConciAprov As Variant
 Dim M1 As Object
 Dim M As Object
 Dim splitRazConci As Variant
              
    enviro = CStr(Environ("USERPROFILE"))

    strPath = enviro & "\Documents\Controle Solução SE.xlsx"
     
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
    Set xlSheet = xlWB.Sheets("Controle Cliente")
    Set xlSheetConci = xlWB.Sheets("Controle Conciliadora")
    
    Set ReglInfoCliente = CreateObject("VBScript.RegExp")
    With ReglInfoCliente
        .Pattern = "(razão\s*social:([^)]+)CNPJ:([^)]+)grupo\s*econômico:)"
    End With
    If ReglInfoCliente.test(item.Body) Then
        Set M1 = ReglInfoCliente.Execute(item.Body)
        For Each M In M1
           razaoSocial = Trim(M.SubMatches(1))
           cnpjCliente = Trim(M.SubMatches(2))
        Next
    End If
    
    Set ReglInfoConcRazao = CreateObject("VBScript.RegExp")
    With ReglInfoConcRazao
        .Pattern = "(nome\s*da\s*conciliadora:([^)]+)CNPJ\s*da\s*conciliadora)"
    End With
    If ReglInfoConcRazao.test(item.Body) Then
        Set M1 = ReglInfoConcRazao.Execute(item.Body)
        For Each M In M1
           razaoConci = Trim(Replace(M.SubMatches(1), vbCrLf, ""))
        Next
    End If
    
    Set ReglInfoConcCnpj = CreateObject("VBScript.RegExp")
    With ReglInfoConcCnpj
        .Pattern = "(CNPJ\s*da\s*conciliadora:([^)]+)eu,)"
    End With
    If ReglInfoConcCnpj.test(item.Body) Then
        Set M1 = ReglInfoConcCnpj.Execute(item.Body)
        For Each M In M1
           cnpjConci = Trim(Replace(M.SubMatches(1), vbCrLf, ""))
        Next
    End If

    Set ReglAprovacao = CreateObject("VBScript.RegExp")
    ' \s* = invisible spaces
    ' \d* = match digits
    ' \w* = match alphanumeric
    ' ([^)]+) = Considera tudo até o prox. match
    With ReglAprovacao
        .Pattern = "(eu,([^)]+),\s*portador\s*do\s*Documento\s*de\s*Identidade\s*nº([^)]+)e\s*" & _
                   "do\s*CPF\s*nº([^)]+),\s*declaro\s*que\s*estou\s*de\s*acordo\s*com\s*o\s*" & _
                   "compartilhamento\s*de\s*informações\s*com\s*a\s*conciliadora([^)]+),\s*CNPJ([^)]+)com\s*essa)"
    End With
    If ReglAprovacao.test(item.Body) Then
    ' each "(\w*)" and the "(\d)" are assigned a vText variable
        Set M1 = ReglAprovacao.Execute(item.Body)
        For Each M In M1
           aprovador = Trim(M.SubMatches(1))
           rgAprov = Trim(M.SubMatches(2))
           cpfAprov = Trim(M.SubMatches(3))
           conciAprov = Trim(M.SubMatches(4))
           cnpjConciAprov = Trim(Replace(M.SubMatches(5), vbCrLf, "")) 'remove quebra de linha com replace
        Next
    End If


    If Len(aprovador) > 0 And _
        Len(rgAprov) > 0 And _
        Len(cpfAprov) > 0 And _
        Len(conciAprov) > 0 And _
        Len(cnpjConciAprov) > 0 _
    Then
    
                            
        splitRazConci = Split(conciAprov)
        conciAprov = Trim(splitRazConci(0))
        
        With xlSheet.Range("A1")
        
           Set rangeCli = .Find(What:=StripIllegalChar(cnpjCliente) & "_" & StripIllegalChar(conciAprov), _
                            After:=.Cells(1), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False)
            
        End With
        
        If rangeCli Is Nothing Then
            'Acha a proxima linha na planilha
            rCount = xlSheet.Range("A" & xlSheet.Rows.Count).End(xlUp).Row
            rCount = rCount + 1
    
            xlSheet.Range("A" & rCount) = StripIllegalChar(cnpjCliente) & "_" & StripIllegalChar(conciAprov)
            xlSheet.Range("B" & rCount) = razaoSocial
            xlSheet.Range("C" & rCount) = cnpjCliente
            xlSheet.Range("D" & rCount) = razaoConci
            xlSheet.Range("E" & rCount) = cnpjConci
            xlSheet.Range("G" & rCount) = Now
            xlSheet.Range("H" & rCount) = aprovador & " / " & cpfAprov & " / " & rgAprov
            xlSheet.Range("I" & rCount) = item.Sender.Address
        Else
            MsgBox "A confirmação do cliente " & razaoSocial & " CNPJ: " & _
                cnpjCliente & " para a conciliadora " & razaoConci & _
                " CNPJ: " & cnpjConci & " já foi inserida na planilha de controle. Verifique", vbOKOnly, "Atenção!"
        End If
        
        With xlSheetConci.Range("A1")
        
           Set rangeConci = .Find(What:=cnpjConci, _
                            After:=.Cells(1), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False)
            
        End With
        
        If rangeConci Is Nothing Then
            rCountConci = xlSheetConci.Range("A" & xlSheetConci.Rows.Count).End(xlUp).Row
            rCountConci = rCountConci + 1
            xlSheetConci.Range("A" & rCountConci) = cnpjConci
            xlSheetConci.Range("B" & rCountConci) = razaoConci
        End If

    End If

    xlWB.Close 1
    If bXStarted Then
       xlApp.Quit
    End If
    
    Set M = Nothing
    Set M1 = Nothing
    Set ReglAprovacao = Nothing
    Set ReglInfoCliente = Nothing
    Set ReglInfoConcRazao = Nothing
    Set ReglInfoConcCnpj = Nothing
    Set xlApp = Nothing
    Set xlWB = Nothing
    Set xlSheet = Nothing
    Set xlSheetConci = Nothing
    
 End Sub
 
 Function StripIllegalChar(StrInput)
    
    Dim RegX            As Object
    
    Set RegX = CreateObject("vbscript.regexp")
    
    RegX.Pattern = "[\" & Chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,\.\-]"
    RegX.IgnoreCase = True
    RegX.Global = True
    
    StripIllegalChar = RegX.Replace(StrInput, "")
    
ExitFunction:
    
    Set RegX = Nothing
    
End Function
