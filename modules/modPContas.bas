Attribute VB_Name = "modPContas"
Public Sub cadastro()

Dim ws As Worksheet
Dim obj As cNotas
Dim lRow As Long, x As Long
        
Set ws = Worksheets("PCONTAS")
Set obj = New cNotas
       
''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1
            
        With obj
            .id = CStr(ws.Range("A" & x).Value)
            .FK = CStr(ws.Range("B" & x).Value)
            
            .Titulo = CStr(ws.Range("C" & x).Value)
            .Descricao = CStr(ws.Range("D" & x).Value)
            .CadastroCategoria = "PCONTA"
            .Procedure = "spFinancas"
            
            .add obj
        End With
        
        If obj.id = "0" Then
            obj.Insert carregarBanco, obj
        ElseIf obj.id <> "" And obj.Titulo <> "" Then
            obj.Update carregarBanco, obj
        Else
            obj.Delete carregarBanco, obj
        End If
        
    Next x
                      
Set obj = Nothing

End Sub

Sub Listar()
Dim obj As cNotas
Set obj = New cNotas

Dim col As cNotas
Set col = obj.getNotasID(carregarBanco, "vw_pcontas", "53")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("PCONTAS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

For Each obj In col.Itens

    With obj

        ws.Range("A" & lRow).Value = .id
        ws.Range("B" & lRow).Value = .FK
                    
        ws.Range("C" & lRow).Value = .Titulo
        ws.Range("D" & lRow).Value = .Descricao

        lRow = lRow + 1
    
    End With
    
Next obj

End Sub


