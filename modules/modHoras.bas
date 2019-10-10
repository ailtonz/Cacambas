Attribute VB_Name = "modHoras"
Private Sub cadastro()
Dim ws As Worksheet
Dim obj As cHoras
Dim lRow As Long, x As Long
            
Set ws = Worksheets("RESERVATORIO")
Set obj = New cHoras

        
''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 5 To lRow - 1
            
        With obj
            .id = CStr(ws.Range("A" & x).Text)
            .FK = CStr(ws.Range("B" & x).Text)

            .Cliente = CStr(ws.Range("C" & x).Text)
            .Tarefa = CStr(ws.Range("D" & x).Text)
            .Controle = CStr(ws.Range("E" & x).Text)
            .DataTarefa = CStr(ws.Range("F" & x).Text)
            .SalarioHora = Trim(CStr(ws.Range("G" & x).Text))
            .CargaHoraria = CStr(ws.Range("H" & x).Text)
            .Turno01_Entrada = CStr(ws.Range("I" & x).Text)
            .Turno01_Saida = CStr(ws.Range("J" & x).Text)
            .Turno02_Entrada = CStr(ws.Range("L" & x).Text)
            .Turno02_Saida = CStr(ws.Range("M" & x).Text)

            .Procedure = "spHoras"

            .add obj
        End With
        
        If obj.id = "0" Then
            obj.Insert carregarBanco, obj
        ElseIf obj.id <> "" And obj.FK <> "" Then
            obj.Update carregarBanco, obj
        Else
            obj.Delete carregarBanco, obj
        End If
        
    Next x
                      
Set obj = Nothing


End Sub

Sub Listar()
Dim obj As cHoras
Set obj = New cHoras

Dim col As cHoras
'Set col = obj.getHoras(carregarBanco, "vw_horas")
'Set col = obj.getHorasPorFuncionarioPeriodo(carregarBanco, "vw_horas", "AILTON", "fev_15")
Set col = obj.getHorasPorFuncionario(carregarBanco, "vw_horas", "AILTON")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("HORAS")

''find  first empty row in database
'lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
lRow = 5
    For Each obj In col.Itens
    
        With obj
    
            ws.Range("A" & lRow).Value = .id
            ws.Range("B" & lRow).Value = .FK
                        
            ws.Range("C" & lRow).Value = .Cliente
            ws.Range("D" & lRow).Value = .Tarefa
            ws.Range("E" & lRow).Value = .Controle
            ws.Range("F" & lRow).Value = .DataTarefa
            ws.Range("G" & lRow).Value = IIf(Trim(.SalarioHora) = Null, "0", CDbl(Trim(.SalarioHora)))
            ws.Range("H" & lRow).Value = .CargaHoraria
            ws.Range("I" & lRow).Value = .Turno01_Entrada
            ws.Range("J" & lRow).Value = .Turno01_Saida
            ws.Range("L" & lRow).Value = .Turno02_Entrada
            ws.Range("M" & lRow).Value = .Turno02_Saida
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub
