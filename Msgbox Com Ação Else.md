Sub MsgBox_Com_Ação_Else

If Sheets("SOLICITACAO").Range("I3").Value = "SALDO INSUFICIENTE" Then

MsgBox "TRANSPORTE SEM SALDO OU INSUFICIENTE!", vbExclamation, "Atenção"

ElseIf Sheets("SOLICITACAO").Range("B3").Value = "INSIRA TODAS AS INFORMAÇÕES" Then

MsgBox "PREENCHA TODAS AS INFORMAÇÕES!", vbExclamation, "Atenção"

ElseIf Sheets("SOLICITACAO").Range("A3").Value = Sheets("SOLICITACAO").Range("E3").Value And Sheets("SOLICITACAO").Range("E3").Value <> "" Then

MsgBox "TRANSPORTE UTILIZADO, INSIRA OUTRO NÚMERO DE TRANSPORTE!", vbExclamation, "Atenção"

Else

End If

End Sub