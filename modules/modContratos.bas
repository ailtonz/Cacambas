Attribute VB_Name = "modContratos"
Private Sub cadastroCobranca()
Dim lRow As Long, x As Long

Dim obj As cCobranca
Set obj = New cCobranca

Dim ws As Worksheet
Set ws = Worksheets("CONTRATOS")

carregarBanco
        
''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
For x = 2 To lRow - 1
                    
    With obj
        .id = CStr(ws.Range("A" & x).Value)
        .FK = CStr(ws.Range("B" & x).Value)

        .CobrancaTipo = CStr(ws.Range("C" & x).Value)
        .CobrancaInscricao = CStr(ws.Range("D" & x).Value)
        .CobrancaSacado = CStr(ws.Range("E" & x).Value)
        
        .CobrancaCep = CStr(ws.Range("F" & x).Value)
        .CobrancaLogradouro = CStr(ws.Range("G" & x).Value)
        .CobrancaBairro = CStr(ws.Range("H" & x).Value)
        .CobrancaCidade = CStr(ws.Range("I" & x).Value)
        .CobrancaEstado = CStr(ws.Range("J" & x).Value)
        
        .CobrancaContato = CStr(ws.Range("K" & x).Value)
        .CobrancaTelefone = CStr(ws.Range("L" & x).Value)
        .CobrancaEmail = CStr(ws.Range("M" & x).Value)
                
        .CadastroCategoria = "CLIENTE_OBRA"
        .Procedure = "spCobranca"

        .add obj
    End With
    
    If obj.id = "0" Then
        obj.Insert carregarBanco, obj
    ElseIf obj.id <> "" And obj.CobrancaTipo <> "" Then
        obj.Update carregarBanco, obj
    End If
    
Next x
                      
Set obj = Nothing


End Sub

Private Sub cadastro()
Dim ws As Worksheet
Dim obj As cContratos
Dim lRow As Long, x As Long
        
Set ws = Worksheets("CONTRATOS")
Set obj = New cContratos
        
''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1
            
        With obj
            .id = CStr(ws.Range("A" & x).Value)
            .FK = CStr(ws.Range("B" & x).Value)

            .ContratoInicio = CStr(ws.Range("N" & x).Value)
            .ContratoTerminio = CStr(ws.Range("O" & x).Value)
            .ContratoValor = CStr(ws.Range("P" & x).Value)
            .ContratoNF = CStr(ws.Range("Q" & x).Value)
            .ContratoISS = CStr(ws.Range("R" & x).Value)
            .ContratoCTR = CStr(ws.Range("S" & x).Value)
            .ContratoPeriodoLocacao = CStr(ws.Range("T" & x).Value)
            .ContratoTransacao = CStr(ws.Range("U" & x).Value)
            .ContratoCondicoes = CStr(ws.Range("V" & x).Value)
            .ContratoRetiradaAutomatica = CStr(ws.Range("W" & x).Value)
            .ContratoVctoAposEntrega = CStr(ws.Range("X" & x).Value)
            .ContratoMultaMora = CStr(ws.Range("Y" & x).Value)
            .ContratoMultaDia = CStr(ws.Range("Z" & x).Value)
            .ContratoObservacao = CStr(ws.Range("AA" & x).Value)
            .ContratoObsColoca = CStr(ws.Range("AB" & x).Value)
            .ContratoObsTroca = CStr(ws.Range("AC" & x).Value)
            .ContratoObsRetira = CStr(ws.Range("AD" & x).Value)
            .ContratoObsLigacao = CStr(ws.Range("AE" & x).Value)
                        
            .CadastroCategoria = "CLIENTE_OBRA"
            .Procedure = "spContratos"

            .add obj
        End With
        
        If obj.id = "0" Then
            obj.Insert carregarBanco, obj
        ElseIf obj.id <> "" And obj.ContratoInicio <> "" Then
            obj.Update carregarBanco, obj
        End If
        
    Next x
                      
Set obj = Nothing


End Sub


Private Sub ListarContratos()
Dim obj As cContratos
Set obj = New cContratos

Dim col As cContratos
Set col = obj.getContratos(carregarBanco, "vw_clientes_obras")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("CONTRATOS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 3).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj
    
            ws.Range("A" & lRow).Value = .id
            ws.Range("B" & lRow).Value = .FK
                        
            ws.Range("N" & lRow).Value = .ContratoInicio
            ws.Range("O" & lRow).Value = .ContratoTerminio
            ws.Range("P" & lRow).Value = .ContratoValor
            ws.Range("Q" & lRow).Value = .ContratoNF
            ws.Range("R" & lRow).Value = .ContratoISS
            ws.Range("S" & lRow).Value = .ContratoCTR
            ws.Range("T" & lRow).Value = .ContratoPeriodoLocacao
            ws.Range("U" & lRow).Value = .ContratoTransacao
            ws.Range("V" & lRow).Value = .ContratoCondicoes
            ws.Range("W" & lRow).Value = .ContratoRetiradaAutomatica
            ws.Range("X" & lRow).Value = .ContratoVctoAposEntrega
            ws.Range("Y" & lRow).Value = .ContratoMultaMora
            ws.Range("Z" & lRow).Value = .ContratoMultaDia
            ws.Range("AA" & lRow).Value = .ContratoObservacao
            ws.Range("AB" & lRow).Value = .ContratoObsColoca
            ws.Range("AC" & lRow).Value = .ContratoObsTroca
            ws.Range("AD" & lRow).Value = .ContratoObsRetira
            ws.Range("AE" & lRow).Value = .ContratoObsLigacao
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub


Private Sub ListarContratosID()
Dim obj As cContratos
Set obj = New cContratos

Dim col As cContratos
Set col = obj.getContratosID(carregarBanco, "vw_clientes_obras", "1708280")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("CONTRATOS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj
    
            ws.Range("A" & lRow).Value = .id
            ws.Range("B" & lRow).Value = .FK
                        
            ws.Range("N" & lRow).Value = .ContratoInicio
            ws.Range("O" & lRow).Value = .ContratoTerminio
            ws.Range("P" & lRow).Value = .ContratoValor
            ws.Range("Q" & lRow).Value = .ContratoNF
            ws.Range("R" & lRow).Value = .ContratoISS
            ws.Range("S" & lRow).Value = .ContratoCTR
            ws.Range("T" & lRow).Value = .ContratoPeriodoLocacao
            ws.Range("U" & lRow).Value = .ContratoTransacao
            ws.Range("V" & lRow).Value = .ContratoCondicoes
            ws.Range("W" & lRow).Value = .ContratoRetiradaAutomatica
            ws.Range("X" & lRow).Value = .ContratoVctoAposEntrega
            ws.Range("Y" & lRow).Value = .ContratoMultaMora
            ws.Range("Z" & lRow).Value = .ContratoMultaDia
            ws.Range("AA" & lRow).Value = .ContratoObservacao
            ws.Range("AB" & lRow).Value = .ContratoObsColoca
            ws.Range("AC" & lRow).Value = .ContratoObsTroca
            ws.Range("AD" & lRow).Value = .ContratoObsRetira
            ws.Range("AE" & lRow).Value = .ContratoObsLigacao
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub

Private Sub ListarCobrancaID()
Dim obj As cCobranca
Set obj = New cCobranca

Dim col As cCobranca
Set col = obj.getCobrancaID(carregarBanco, "vw_clientes_obras", "1708280")

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("CONTRATOS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each obj In col.Itens
    
        With obj
    
'            ws.Range("A" & lRow).Value = .ID
'            ws.Range("B" & lRow).Value = .FK
                        
            ws.Range("C" & lRow).Value = .CobrancaTipo
            ws.Range("D" & lRow).Value = .CobrancaInscricao
            ws.Range("E" & lRow).Value = .CobrancaSacado
                        
            ws.Range("F" & lRow).Value = .CobrancaCep
            ws.Range("G" & lRow).Value = .CobrancaLogradouro
            ws.Range("H" & lRow).Value = .CobrancaBairro
            ws.Range("I" & lRow).Value = .CobrancaCidade
            ws.Range("J" & lRow).Value = .CobrancaEstado
            
            ws.Range("K" & lRow).Value = .CobrancaContato
            ws.Range("L" & lRow).Value = .CobrancaTelefone
            ws.Range("M" & lRow).Value = .CobrancaEmail
    
            lRow = lRow + 1
        
        End With
        
    Next obj

End Sub



