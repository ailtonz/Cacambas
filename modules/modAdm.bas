Attribute VB_Name = "modAdm"
Private Sub carregarCliente()
Dim obj As New cMatriz

AbrirCadastros obj.getMatrizFrm(carregarBanco, "vw_matriz", "CLIENTES"), "1710495"

Set obj = Nothing
    
End Sub

Private Sub carregarTransportador()
Dim obj As New cMatriz

AbrirCadastros obj.getMatrizFrm(carregarBanco, "vw_matriz", "TRANSPORTADOR"), "1710494"

Set obj = Nothing
    
End Sub

Private Sub carregarFuncionario()
Dim obj As New cMatriz

AbrirCadastros obj.getMatrizFrm(carregarBanco, "vw_matriz", "FUNCIONARIO"), "1710433"

Set obj = Nothing
    
End Sub

Private Sub carregarAterro()
Dim obj As New cMatriz

AbrirCadastros obj.getMatrizFrm(carregarBanco, "vw_matriz", "ATERROS"), "1710487"

Set obj = Nothing

End Sub

Private Sub carregarPropagandas()
Dim obj As New cMatriz

AbrirNotas obj.getMatrizFrm(carregarBanco, "vw_matriz", "PROPAGANDAS")

Set obj = Nothing

End Sub

Private Sub carregarMateriais()
Dim obj As New cMatriz

AbrirNotas obj.getMatrizFrm(carregarBanco, "vw_matriz", "MATERIAIS")

Set obj = Nothing

End Sub

Private Sub carregarLinks()
Dim obj As New cMatriz

AbrirNotas obj.getMatrizFrm(carregarBanco, "vw_matriz", "LINKS")

Set obj = Nothing

End Sub

Private Sub carregarVeiculos()
Dim obj As New cMatriz

AbrirNotas obj.getMatrizFrm(carregarBanco, "vw_matriz", "VEICULOS")

Set obj = Nothing

End Sub

Private Sub carregarCargos()
Dim obj As New cMatriz

AbrirNotas obj.getMatrizFrm(carregarBanco, "vw_matriz", "CARGOS")

Set obj = Nothing

End Sub

Private Sub carregarCondicoes()
Dim obj As New cMatriz

AbrirNotas obj.getMatrizFrm(carregarBanco, "vw_matriz", "CONDICOES")

Set obj = Nothing

End Sub

Private Sub carregarTransacoes()
Dim obj As New cMatriz

AbrirNotas obj.getMatrizFrm(carregarBanco, "vw_matriz", "TRANSACOES")

Set obj = Nothing

End Sub

Private Sub carregarPContas()
Dim obj As New cMatriz

AbrirNotas obj.getMatrizFrm(carregarBanco, "vw_matriz", "PCONTAS")

Set obj = Nothing

End Sub


