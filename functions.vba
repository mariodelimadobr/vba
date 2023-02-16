Function Function_InserirColunaID()
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Inserir uma coluna antes da coluna A
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'Definir o cabeçalho da coluna como "id"
    Range("A1").Value = "id"
    
    'Preencher a coluna "id" com uma sequência numérica de 1 até o último registro
    Range("A2:A" & lastRow).Formula = "=ROW()-1"
    Range("A2:A" & lastRow).Value = Range("A2:A" & lastRow).Value
    
    'Selecionar a célula A1
    Range("A1").Select
End Function

Function Function_RemoveEspacos()

    Dim cel As Range
    
    For Each cel In ActiveSheet.UsedRange
        If VarType(cel.Value) = vbString Then
            cel.Value = Trim(cel.Value)
        End If
    Next cel

End Function

Function Function_TransformaDatas()
    Dim dataCelula As Range
    
    For Each dataCelula In coluna.Cells
        If IsDate(dataCelula.Value) Then
            dataCelula.Value = Format(dataCelula.Value, "yyyy-mm-dd")
        End If
    Next dataCelula
    
    Set TransformaDatas = coluna
End Function

Function Function_RemoverCaracteres()

    Dim cel As Range
    
    For Each cel In Sheets("DADOS").UsedRange
        If cel.Value <> "" Then
            cel.Value = Replace(cel.Value, ",", "")
            cel.Value = Replace(cel.Value, ";", "")
            cel.Value = Replace(cel.Value, "'", "")
            cel.Value = Replace(cel.Value, """", "")
        End If
    Next cel

End Function

Function Function_TratarCPF(col As String) As Range
    Dim ultimaLinha As Long
    ultimaLinha = Sheets("DADOS").Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim cpf As Range
    
    For Each cpf In Sheets("DADOS").Range(col & "2:" & col & ultimaLinha)
        'Remove pontos, hífens e espaços do CPF
        cpf.Value = Replace(Replace(Replace(cpf.Value, ".", ""), "-", ""), " ", "")
        
        'Formata o CPF com 11 dígitos
        cpf.NumberFormat = "00000000000"
    Next cpf
    
    'retorna o range tratado
    Set TratarCPF = Sheets("DADOS").Range(col & "2:" & col & ultimaLinha)
End Function

Function Function_RemoverQuebrasDeLinha()
    Dim celula As Range
    
    For Each celula In ActiveSheet.UsedRange.Cells
        celula.Value = Replace(celula.Value, vbLf, "")
        celula.Value = Replace(celula.Value, vbCr, "")
    Next celula
End Function

Function Function_PreencherCelulasVazias()

    Dim ultimaLinha As Long
    ultimaLinha = Sheets("DADOS").Cells(Rows.Count, "A").End(xlUp).Row ' Encontra a última linha da coluna A
    
    Dim celula As Range
    For Each celula In Sheets("DADOS").Range("A1:A" & ultimaLinha) ' Percorre a coluna A
        If celula.Value = "" Then ' Se a célula estiver vazia
            celula.Value = celula.End(xlUp).Value ' Copia o valor da célula acima
        End If
    Next celula

End Function

