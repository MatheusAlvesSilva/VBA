Sub compilação()

linha = 1
linha_fim = Range("A1").End(xlDown).Row

Range("N2:N" & linha_fim).Copy
Range("AL1").PasteSpecial
Application.CutCopyMode = False
ActiveSheet.Range("$AL$1:$AL$" & linha_fim).RemoveDuplicates Columns:=1, Header:= _
    xlNo
    
linha_fim = Range("AL1").End(xlDown).Row
    
While linha <= linha_fim
    Sheets.Add After:=ActiveSheet
    
    ActiveSheet.Name = Sheets("BaseGeral").Cells(linha, 38)
    
    Sheets("BaseGeral").Range("A1:T1").Copy
    ActiveSheet.Range("A1").PasteSpecial
    
    linha = linha + 1
Wend

Sheets("BaseGeral").Range("AL:AL").Clear

linha = 2

While Sheets("BaseGeral").Cells(linha, 1) <> ""
    Sheets("BaseGeral").Range("A" & linha & ":T" & linha).Copy
    
    REGIÃO = Sheets("BaseGeral").Cells(linha, 14)
    Sheets("BaseGeral").Select
    Range("A100000").End(xlUp).Offset(1, 0).PasteSpecial
    Application.CutCopyMode = False
    
    Sheets("BaseGeral").Select

    linha = linha + 1
Wend

For Each aba In ThisWorkbook.Sheets
    aba.Columns("A:T").AutoFit
Next

End Sub
