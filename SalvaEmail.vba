Option Explicit
 
Sub SalvaEmail(Item As Outlook.MailItem)
 Dim FileName As String
 Dim ReglAprovacao, ReglInfoCliente, ReglInfoConcRazao, ReglInfoConcCnpj  As Object
 Dim M1 As Object
 Dim M As Object
 Dim razaoSocial, cnpjCliente, razaoConci, cnpjConci, aprovador, rgAprov, cpfAprov, conciAprov, cnpjConciAprov As Variant
 Dim splitRazConci As Variant
 

    Set ReglInfoCliente = CreateObject("VBScript.RegExp")
    With ReglInfoCliente
        .Pattern = "(razão\s*social:([^)]+)CNPJ:([^)]+)grupo\s*econômico:)"
    End With
    If ReglInfoCliente.test(Item.Body) Then
    ' each "(\w*)" and the "(\d)" are assigned a vText variable
        Set M1 = ReglInfoCliente.Execute(Item.Body)
        For Each M In M1
           razaoSocial = Trim(M.SubMatches(1))
           cnpjCliente = Trim(M.SubMatches(2))
        Next
    End If
    
    Set ReglInfoConcRazao = CreateObject("VBScript.RegExp")
    With ReglInfoConcRazao
        .Pattern = "(nome\s*da\s*conciliadora:([^)]+)CNPJ\s*da\s*conciliadora)"
    End With
    If ReglInfoConcRazao.test(Item.Body) Then
        Set M1 = ReglInfoConcRazao.Execute(Item.Body)
        For Each M In M1
           razaoConci = Trim(Replace(M.SubMatches(1), vbCrLf, ""))
        Next
    End If
    
    Set ReglInfoConcCnpj = CreateObject("VBScript.RegExp")
    With ReglInfoConcCnpj
        .Pattern = "(CNPJ\s*da\s*conciliadora:([^)]+)eu,)"
    End With
    If ReglInfoConcCnpj.test(Item.Body) Then
        Set M1 = ReglInfoConcCnpj.Execute(Item.Body)
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
    If ReglAprovacao.test(Item.Body) Then
    ' each "(\w*)" and the "(\d)" are assigned a vText variable
        Set M1 = ReglAprovacao.Execute(Item.Body)
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

        FileName = Environ("USERPROFILE") & "\Documents\Emails\" & StripIllegalChar(cnpjCliente) & "_" _
                    & StripIllegalChar(conciAprov) & ".msg"
 
        Item.SaveAs FileName, 3
    Else
        MsgBox "A confirmação do cliente " & razaoSocial & " CNPJ: " & _
                cnpjCliente & " para a conciliadora " & razaoConci & _
                " CNPJ: " & cnpjConci & " não retornou com todos os dados.", vbOKOnly, "Atenção!"
    End If
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
