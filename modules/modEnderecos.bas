Attribute VB_Name = "modEnderecos"
Private Sub cadastro()
Dim ws As Worksheet
Dim obj As cEnderecos
Dim lRow As Long, x As Long
        
Set ws = Worksheets("ENDERECOS")
Set obj = New cEnderecos
       
''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
For x = 2 To lRow - 1
        
    With obj
        .id = CStr(ws.Range("A" & x).Value)
        .FK = CStr(ws.Range("B" & x).Value)
        
        .Cep = CStr(ws.Range("C" & x).Value)
        .Numero = CStr(ws.Range("D" & x).Value)
        .Complemento = CStr(ws.Range("E" & x).Value)
        .Logradouro = CStr(ws.Range("F" & x).Value)
        .Bairro = CStr(ws.Range("G" & x).Value)
        .Cidade = CStr(ws.Range("H" & x).Value)
        .Estado = CStr(ws.Range("I" & x).Value)
    
        .CadastroCategoria = "CLIENTE_OBRA"
        .Procedure = "spEnderecos"

        .add obj
    End With
    
    If obj.id = "0" Then
        obj.Insert carregarBanco, obj
    ElseIf obj.id <> "" And obj.Cep <> "" Then
        obj.Update carregarBanco, obj
    Else
        '' ATENÇÃO: CUIDADO AO EXCLUIR O ENDERECO DO CLIENTE!!!!!!
        obj.Delete carregarBanco, obj
    End If
    
Next x
                      
Set obj = Nothing


End Sub

Private Sub ListarEnderecos()
Dim obj As cEnderecos
Set obj = New cEnderecos

Dim col As cEnderecos
Set col = obj.getEnderecos(carregarBanco, "vw_cep")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("ENDERECOS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

For Each obj In col.Itens

    With obj
    
        ws.Range("A" & lRow).Value = .id
        ws.Range("B" & lRow).Value = .FK

        ws.Range("C" & lRow).Value = "'" & .Cep
        ws.Range("D" & lRow).Value = .Numero
        ws.Range("E" & lRow).Value = .Complemento
        ws.Range("F" & lRow).Value = .Logradouro
        ws.Range("G" & lRow).Value = .Bairro
        ws.Range("H" & lRow).Value = .Cidade
        ws.Range("I" & lRow).Value = .Estado

        lRow = lRow + 1
    
    End With
    
Next obj

End Sub


Private Sub ListarEnderecosID()
Dim obj As cEnderecos
Set obj = New cEnderecos

Dim col As cEnderecos
Set col = obj.getEnderecosID(carregarBanco, "vw_cep", "04531070")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("ENDERECOS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj

            ws.Range("C" & lRow).Value = .Cep
            ws.Range("D" & lRow).Value = .Numero
            ws.Range("E" & lRow).Value = .Complemento
            ws.Range("F" & lRow).Value = .Logradouro
            ws.Range("G" & lRow).Value = .Bairro
            ws.Range("H" & lRow).Value = .Cidade
            ws.Range("I" & lRow).Value = .Estado
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub

Private Sub carregarEnderecoPorCep()
Dim obj As cEnderecos
Set obj = New cEnderecos

Dim col As cEnderecos
Set col = obj.getEnderecosCEP(carregarBanco, "vw_cep", "04531070")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("ENDERECOS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj

            ws.Range("C" & lRow).Value = .Cep
            ws.Range("D" & lRow).Value = .Numero
            ws.Range("E" & lRow).Value = .Complemento
            ws.Range("F" & lRow).Value = .Logradouro
            ws.Range("G" & lRow).Value = .Bairro
            ws.Range("H" & lRow).Value = .Cidade
            ws.Range("I" & lRow).Value = .Estado
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub


Private Sub carregarEnderecoCep()
Dim lRow As Long, x As Long
Dim obj As cEnderecos
Dim col As cEnderecos

Dim ws As Worksheet
Set ws = Worksheets("ENDERECOS")

lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row - 1

For x = 2 To lRow - 1
    If CStr(ws.Range("C" & x).Value) <> "" Then
        Set obj = New cEnderecos
        Set col = obj.getEnderecosCEP(carregarBanco, "vw_cep", CStr(ws.Range("C" & x).Value))
        For Each obj In col.Itens
            With obj
                ws.Range("F" & x).Value = .Logradouro
                ws.Range("G" & x).Value = .Bairro
                ws.Range("H" & x).Value = .Cidade
                ws.Range("I" & x).Value = .Estado
            End With
        Next obj
    End If
Next x

End Sub

