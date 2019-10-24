Attribute VB_Name = "ModuloSalvaEmail"
Option Explicit
 
Sub SalvaEmail(item As Outlook.MailItem)
 Dim FileName As String
 Dim ReglAprovacao, ReglInfoCliente, ReglInfoConcRazao, ReglInfoConcCnpj  As Object
 Dim M1 As Object
 Dim M As Object
 Dim razaoSocial, cnpjCliente, razaoConci, cnpjConci, aprovador, rgAprov, cpfAprov, conciAprov, cnpjConciAprov As Variant
 Dim splitRazConci As Variant
 

    Set ReglInfoCliente = CreateObject("VBScript.RegExp")
    With ReglInfoCliente
        .Pattern = "(raz�o\s*social:([^)]+)CNPJ:([^)]+)grupo\s*econ�mico:)"
    End With
    If ReglInfoCliente.test(item.Body) Then
    ' each "(\w*)" and the "(\d)" are assigned a vText variable
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
    ' ([^)]+) = Considera tudo at� o prox. match
    With ReglAprovacao
        .Pattern = "(eu,([^)]+),\s*portador\s*do\s*Documento\s*de\s*Identidade\s*n�([^)]+)e\s*" & _
                   "do\s*CPF\s*n�([^)]+),\s*declaro\s*que\s*estou\s*de\s*acordo\s*com\s*o\s*" & _
                   "compartilhamento\s*de\s*informa��es\s*com\s*a\s*conciliadora([^)]+),\s*CNPJ([^)]+)com\s*essa)"
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

        FileName = Environ("USERPROFILE") & "\Documents\Emails\" & StripIllegalChar(cnpjCliente) & "_" _
                    & StripIllegalChar(conciAprov) & ".msg"
 
        item.SaveAs FileName, 3
    Else
        MsgBox "A confirma��o do cliente " & razaoSocial & " CNPJ: " & _
                cnpjCliente & " para a conciliadora " & razaoConci & _
                " CNPJ: " & cnpjConci & " n�o retornou com todos os dados.", vbOKOnly, "Aten��o!"
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
