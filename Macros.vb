Sub Macro16()
'
' Macro16 Macro
'

'
linha = Sheets("PEDIDOS").Range("B1048576").End(xlUp).Row + 1


Sheets("PEDIDOS").Unprotect Password:="0309"
    Sheets("PEDIDOS").Cells(linha, 2).Value = Sheets("CAIXA").Range("C2").Value
    Sheets("PEDIDOS").Cells(linha, 4).Value = Sheets("CAIXA").Range("D5").Value
    Sheets("PEDIDOS").Cells(linha, 5).Value = Sheets("CAIXA").Range("D6").Value
    Sheets("PEDIDOS").Cells(linha, 6).Value = Sheets("CAIXA").Range("D7").Value
    Sheets("PEDIDOS").Cells(linha, 7).Value = Sheets("CAIXA").Range("D8").Value
    Sheets("PEDIDOS").Cells(linha, 8).Value = Sheets("CAIXA").Range("D9").Value
    Sheets("PEDIDOS").Cells(linha, 9).Value = Sheets("CAIXA").Range("D10").Value
    Sheets("PEDIDOS").Cells(linha, 10).Value = Sheets("CAIXA").Range("D11").Value
    Sheets("PEDIDOS").Cells(linha, 11).Value = Sheets("CAIXA").Range("D12").Value
    Sheets("PEDIDOS").Cells(linha, 12).Value = Sheets("CAIXA").Range("D13").Value
    Sheets("PEDIDOS").Cells(linha, 13).Value = Sheets("CAIXA").Range("D14").Value
Sheets("PEDIDOS").Protect Password:="0309"
    Range("A1").Select
    Sheets("CAIXA").Select
    Range("D5:D14").Select
    Selection.ClearContents
    Range("C2").Select
    Selection.ClearContents
End Sub

Sub Altera_status()
valor = Sheets("CHECKOUT").Range("B1").Value
linha = Sheets("PEDIDOS").Range("A1:A1048576").Find(What:=valor).Row

Sheets("PEDIDOS").Unprotect Password:="0309"
Sheets("PEDIDOS").Cells(linha, 3).Value = "SIM"
Sheets("PEDIDOS").Protect Password:="0309"
Sheets("CHECKOUT").Range("B1").Select
Selection.ClearContents
End Sub
