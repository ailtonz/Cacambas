Attribute VB_Name = "modAterros"
Private Sub cadastro()
Dim ws As Worksheet
Dim obj As cEntidades
Dim lRow As Long, x As Long
        
Set ws = Worksheets("ATERROS")
Set obj = New cEntidades
        
''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1
            
        With obj
            .id = CStr(ws.Range("A" & x).Value)
            .FK = CStr(ws.Range("B" & x).Value)
            
            .CadastroTipo = CStr(ws.Range("C" & x).Value)
            .CnpjCpf = CStr(ws.Range("D" & x).Value)
            .IeRg = CStr(ws.Range("E" & x).Value)
            .Nome = CStr(ws.Range("F" & x).Value)
            .NomeFantasia = CStr(ws.Range("G" & x).Value)
            .CadastroPropaganda = CStr(ws.Range("H" & x).Value)
            .CadastroObservacao = CStr(ws.Range("I" & x).Value)
            .CadastroStatus = CStr(ws.Range("J" & x).Value)
            
            .CadastroCategoria = "ATERRO"
            .Procedure = "spEntidades"

            .add obj
        End With
        
        If obj.id = "0" Then
            obj.Insert carregarBanco, obj
        ElseIf obj.id <> "" And obj.Nome <> "" Then
            obj.Update carregarBanco, obj
        Else
            obj.Delete carregarBanco, obj
        End If
        
    Next x
                      
Set obj = Nothing


End Sub

Private Sub Listar()
Dim obj As cEntidades
Set obj = New cEntidades

Dim col As cEntidades
Set col = obj.getEntidades(carregarBanco, "vw_aterros")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("ATERROS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj
    
            ws.Range("A" & lRow).Value = .id
            ws.Range("B" & lRow).Value = .FK
            
            ws.Range("C" & lRow).Value = .CadastroTipo
            ws.Range("D" & lRow).Value = .CnpjCpf
            ws.Range("E" & lRow).Value = .IeRg
            ws.Range("F" & lRow).Value = .Nome
            ws.Range("G" & lRow).Value = .NomeFantasia
            ws.Range("H" & lRow).Value = .CadastroPropaganda
            ws.Range("I" & lRow).Value = .CadastroObservacao
            ws.Range("J" & lRow).Value = .CadastroStatus
            
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub

