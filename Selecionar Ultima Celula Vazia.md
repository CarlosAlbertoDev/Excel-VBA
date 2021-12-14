Sub Selecionar_Ultima_Celula_Vazia()

Sheets("ABASTECIMENTOS").Select
Range("B1").Select

Do While ActiveCell.Value <> ""

ActiveCell.Offset(1, 0).Select

Loop

End Sub