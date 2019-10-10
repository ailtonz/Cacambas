Attribute VB_Name = "modClientes"
Private Sub cadastro()

Dim ws As Worksheet
Dim obj As cEntidades
Dim lRow As Long, x As Long
        
Set ws = Worksheets("CLIENTES")
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
            
            
            
            .CadastroCategoria = "CLIENTE"
            .Procedure = "spEntidades"
            
'            .EnderecoCep = CStr(ws.Range("G" & x).Value)
'            .EnderecoNumero = CStr(ws.Range("H" & x).Value)
'            .EnderecoComplemento = CStr(ws.Range("I" & x).Value)
'            .EnderecoLogradouro = CStr(ws.Range("J" & x).Value)
'            .EnderecoBairro = CStr(ws.Range("K" & x).Value)
'            .EnderecoCidade = CStr(ws.Range("L" & x).Value)
'            .EnderecoEstado = CStr(ws.Range("M" & x).Value)
        
'            .ContatoNome = CStr(ws.Range("R" & x).Value)
'            .ContatoTelefone = CStr(ws.Range("S" & x).Value)
'            .ContatoEmail = CStr(ws.Range("T" & x).Value)

            .add obj
        End With
        
        If obj.id = "0" Then
            obj.Insert carregarBanco, obj
        ElseIf obj.id <> "" And obj.Nome <> "" Then
            obj.Update carregarBanco, obj
        Else
            obj.Delete carregarBanco, obj
        End If



'    If (obj.Update(carregarBanco, obj)) Then
'        MsgBox "Alteração realizada com sucesso!", vbInformation + vbOKOnly, "Alteração"
'    Else
'        MsgBox "Não foi possivel realizar alteração!", vbCritical + vbOKOnly, "Alteração - ERRO!"
'    End If
        
    Next x
                      
Set obj = Nothing

End Sub

Private Sub ListarClientes()
Dim obj As cEntidades
Set obj = New cEntidades

Dim col As cEntidades
Set col = obj.getEntidadesID(carregarBanco, "vw_clientes", "1710482")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("CLIENTES")

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

Private Sub ListarClientesContatos()
Dim obj As cContatos
Set obj = New cContatos

Dim col As cContatos
Set col = obj.getContatos(carregarBanco, "vw_clientes")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("CLIENTES")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 18).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj
    
'            ws.Range("A" & lRow).Value = .ID
'            ws.Range("B" & lRow).Value = .FK
            
            ws.Range("R" & lRow).Value = .ContatoNome
            ws.Range("S" & lRow).Value = .ContatoTelefone
            ws.Range("T" & lRow).Value = .ContatoEmail
            
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub




