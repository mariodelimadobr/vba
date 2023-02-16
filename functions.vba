Sub RemoverCaracteres()

    Dim cel As Range
    
    For Each cel In Sheets("DADOS").UsedRange
        If cel.Value <> "" Then
            cel.Value = Replace(cel.Value, ",", "")
            cel.Value = Replace(cel.Value, ";", "")
            cel.Value = Replace(cel.Value, "'", "")
            cel.Value = Replace(cel.Value, """", "")
        End If
    Next cel

End Sub

Sub TratarCPF()
    Dim ultimaLinha As Long
    ultimaLinha = Sheets("DADOS").Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim cpf As Range
    For Each cpf In Sheets("DADOS").Range("B2:B" & ultimaLinha)
        'Remove pontos, hífens e espaços do CPF
        cpf.Value = Replace(Replace(Replace(cpf.Value, ".", ""), "-", ""), " ", "")
        
        'Formata o CPF com 11 dígitos
        cpf.NumberFormat = "00000000000"
    Next cpf
End Sub

Sub RemoverQuebrasDeLinha()
    Dim celula As Range
    
    For Each celula In ActiveSheet.UsedRange.Cells
        celula.Value = Replace(celula.Value, vbLf, "")
        celula.Value = Replace(celula.Value, vbCr, "")
    Next celula
End Sub

Sub PreencherCelulasVazias()

    Dim ultimaLinha As Long
    ultimaLinha = Sheets("DADOS").Cells(Rows.Count, "A").End(xlUp).Row ' Encontra a última linha da coluna A
    
    Dim celula As Range
    For Each celula In Sheets("DADOS").Range("A1:A" & ultimaLinha) ' Percorre a coluna A
        If celula.Value = "" Then ' Se a célula estiver vazia
            celula.Value = celula.End(xlUp).Value ' Copia o valor da célula acima
        End If
    Next celula

End Sub


