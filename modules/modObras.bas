Attribute VB_Name = "modObras"
Sub ListarObras()
Dim obj As cEnderecos
Set obj = New cEnderecos

Dim col As cEnderecos
Set col = obj.getEnderecos(carregarBanco, "vw_clientes_obras")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("OBRAS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj
    
            ws.Range("A" & lRow).Value = .id
            ws.Range("B" & lRow).Value = .FK
                        
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

