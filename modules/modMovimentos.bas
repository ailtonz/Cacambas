Attribute VB_Name = "modMovimentos"
Public Sub cadastro()
Dim ws As Worksheet
Dim obj As cMovimentos
Dim lRow As Long, x As Long
        
Set ws = Worksheets("MOVIMENTOS")
Set obj = New cMovimentos
       
''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1
            
        With obj
            .id = CStr(ws.Range("A" & x).Value)
            .FK = CStr(ws.Range("B" & x).Value)
            
            .DataDeEmissao = ws.Range("C" & x).Value
            .Documento = ws.Range("D" & x).Value
            .Observacao = ws.Range("E" & x).Value
            .DataDeVencimento = ws.Range("F" & x).Value
            .ValorOriginal = Replace(ws.Range("G" & x).Value, ",", ".")
            .DataDePagamento = ws.Range("H" & x).Value
            .ValorFinal = Replace(ws.Range("I" & x).Value, ",", ".")
            .Movimento = ws.Range("J" & x).Value
            .Grupo = ws.Range("K" & x).Value
            .Conta = ws.Range("L" & x).Value
            .Transacao = ws.Range("M" & x).Value
            .Frequencia = ws.Range("N" & x).Value
            
            .Procedure = "spMovimento"
            
            .add obj
        End With
        
        If obj.id = "0" Then
            obj.Insert carregarBanco, obj
        ElseIf obj.id <> "0" And obj.FK <> "" Then
            obj.Update carregarBanco, obj
        Else
            obj.Delete carregarBanco, obj
        End If
        
    Next x
                      
Set obj = Nothing


End Sub

Sub Listar()
Dim obj As cMovimentos
Set obj = New cMovimentos

Dim col As cMovimentos
'Set col = obj.getMovimentosID(carregarBanco, "vw_movimentos", "2407")
Set col = obj.getMovimentos(carregarBanco, "vw_movimentos")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("MOVIMENTOS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj
    
            ws.Range("A" & lRow).Value = .id
            ws.Range("B" & lRow).Value = .FK
                        
            ws.Range("C" & lRow).Value = .DataDeEmissao
            ws.Range("D" & lRow).Value = .Documento
            ws.Range("E" & lRow).Value = .Observacao
            ws.Range("F" & lRow).Value = .DataDeVencimento
            ws.Range("G" & lRow).Value = .ValorOriginal
            ws.Range("H" & lRow).Value = .DataDePagamento
            ws.Range("I" & lRow).Value = .ValorFinal
            ws.Range("J" & lRow).Value = .Movimento
            ws.Range("K" & lRow).Value = .Grupo
            ws.Range("L" & lRow).Value = .Conta
            ws.Range("M" & lRow).Value = .Transacao
            ws.Range("N" & lRow).Value = .Frequencia
            
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub

Sub ListarDados()
Dim obj As cMovimentos
Set obj = New cMovimentos

Dim col As cMovimentos
Set col = obj.getMovimentosDados(carregarBanco, "vw_movimentos")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("dados")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj
    
            ws.Range("A" & lRow).Value = .id
            ws.Range("B" & lRow).Value = .FK
                        
            ws.Range("C" & lRow).Value = .DataDeEmissao
            ws.Range("D" & lRow).Value = .Documento
            ws.Range("E" & lRow).Value = .Observacao
            ws.Range("F" & lRow).Value = .DataDeVencimento
            ws.Range("G" & lRow).Value = .ValorOriginal
            ws.Range("H" & lRow).Value = .DataDePagamento
            ws.Range("I" & lRow).Value = .ValorFinal
            ws.Range("J" & lRow).Value = .Movimento
            ws.Range("K" & lRow).Value = .Grupo
            ws.Range("L" & lRow).Value = .Conta
            ws.Range("M" & lRow).Value = .Transacao
            ws.Range("N" & lRow).Value = .Frequencia
            
            ws.Range("O" & lRow).Value = .Ano
            ws.Range("P" & lRow).Value = .Mes
            ws.Range("Q" & lRow).Value = .Ref
            ws.Range("R" & lRow).Value = .Plano
            
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub

