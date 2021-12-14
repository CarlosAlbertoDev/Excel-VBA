Sub Enviar_Email()

Application.ScreenUpdating = False

Dim DEST As String
Dim DESTcc As String
Dim MOT As String
Dim Placa As String

DEST = Sheets("SOLICITACAO").Range("C5").Value
DESTcc = Sheets("SOLICITACAO").Range("C6").Value
MOT = Sheets("SOLICITACAO").Range("J3").Value
Placa = Sheets("SOLICITACAO").Range("K3").Value

Sheets("SOLICITACAO").Select
        
Range("B2:K3").Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
     
ActiveWorkbook.EnvelopeVisible = True

With ActiveSheet.MailEnvelope

 .Item.Subject = "Abastecimento" & " " & MOT & "-" & Placa
 
 .Introduction = "Bom Dia," _
 & " por favor, liberar abastecimento." _
 & " Empresa: Chess Log"
      
      .Item.To = DEST
      
      .Item.Cc = DESTcc
            
      .Item.Send

   End With
   
   Sheets("SOLICITACAO").Select

Range("C3").Select
Range("C3").ClearContents
Range("E3").Select
Range("E3").ClearContents
Range("H3").Select
Range("H3").ClearContents

Sheets("SOLICITACAO").Select

Application.ScreenUpdating = True
   
End If

End Sub