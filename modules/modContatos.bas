Attribute VB_Name = "modContatos"
Private Sub cadastro()
Dim ws As Worksheet
Dim obj As cContatos
Dim lRow As Long, x As Long
        
Set ws = Worksheets("CONTATOS")
Set obj = New cContatos

        
''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1
            
        With obj
            .id = CStr(ws.Range("A" & x).Value)
            .FK = CStr(ws.Range("B" & x).Value)

            .ContatoNome = CStr(ws.Range("C" & x).Value)
            .ContatoTelefone = CStr(ws.Range("D" & x).Value)
            .ContatoEmail = CStr(ws.Range("E" & x).Value)
            
            .CadastroCategoria = "CONTATO_CLIENTE"
            .Procedure = "spContatos"

            .add obj
        End With
        
        If obj.id = "0" Then
            obj.Insert carregarBanco, obj
        ElseIf obj.id <> "" And obj.ContatoNome <> "" Then
            obj.Update carregarBanco, obj
        Else
            obj.Delete carregarBanco, obj
        End If
        
    Next x
                      
Set obj = Nothing


End Sub

Private Sub ListarContatos()
Dim obj As cContatos
Set obj = New cContatos

Dim col As cContatos
Set col = obj.getContatos(carregarBanco, "vw_clientes_obras")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("CONTATOS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 3).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj
    
            ws.Range("A" & lRow).Value = .id
            ws.Range("B" & lRow).Value = .FK
                        
            ws.Range("C" & lRow).Value = .ContatoNome
            ws.Range("D" & lRow).Value = .ContatoTelefone
            ws.Range("E" & lRow).Value = .ContatoEmail
'            ws.Range("F" & lRow).Value = .ContatoObservacao
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub

