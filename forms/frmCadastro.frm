VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadastro 
   Caption         =   "CADASTRO"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10920
   OleObjectBlob   =   "frmCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sFrm As String
Public sProcedure As String
Public sCategoria As String
Public sConsulta As String
Public sTitulo As String
Public sID As String
Public sFK As String

Private Sub UserForm_Activate()
    Me.Caption = sTitulo
    
    If sID <> "0" Then
        carregarCampos
    Else
        limparCampos
    End If
    
End Sub

Private Sub cmdSalvar_Click()
    salvar objEntidade
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub

Private Sub cmdContratos_Click()
    carregarContratos
End Sub

Private Sub cmdObras_Click()
    carregarObras
End Sub

Private Sub cmdContatos_Click()
    carregarContatos
End Sub

Private Sub cmdObservacoes_Click()
    carregarObservacoes
End Sub


Private Sub salvar(ByVal objEntidade As cEntidades)

    If Me.cmdSalvar.Caption = "NOVO" Then
        If (objEntidade.Insert(carregarBanco, objEntidade) = True) Then
            MsgBox "Cadastro realizado com sucesso!", vbInformation + vbOKOnly, "Cadastro"
        Else
            MsgBox "Não foi possivel realizar o cadastro!", vbCritical + vbOKOnly, "Cadastro - ERRO!"
        End If
    ElseIf Me.cmdSalvar.Caption = "SALVAR" Then
        If (objEntidade.Update(carregarBanco, objEntidade) = True) Then
            MsgBox "Alteração realizada com sucesso!", vbInformation + vbOKOnly, "Alteração"
        Else
            MsgBox "Não foi possivel realizar alteração!", vbCritical + vbOKOnly, "Alteração - ERRO!"
        End If
    ElseIf Me.cmdSalvar.Caption = "EXCLUIR" Then
        If mostrarRegistro = vbYes Then
            If (objEntidade.Delete(carregarBanco, objEntidade) = True) Then
                MsgBox "Exclusão realizada com sucesso!", vbInformation + vbOKOnly, "Exclusão"
            Else
                MsgBox "Não foi possivel realizar Exclusão!", vbCritical + vbOKOnly, "Exclusão - ERRO!"
            End If
        End If
    End If

End Sub


Private Sub limparCampos()
    
    Me.txtID.Value = "0"
    Me.txtFK.Value = "0"
    
    Me.cboTipo.Text = ""
    Me.cboStatus.Text = ""
    
    Me.txtCNPJ_CPF.Text = ""
    Me.txtIE_RG.Text = ""
    
    Me.txtNome.Text = ""
    Me.txtApelido.Text = ""
    Me.txtObservacao.Text = ""
    Me.cboPropaganda.Text = ""
    
    Me.txtCep.Text = ""
    Me.txtNumero.Text = ""
    Me.txtComplemento.Text = ""
    Me.txtBairro.Text = ""
    Me.txtLogradouro.Text = ""
    Me.cboCidade.Text = ""
    Me.cboEstado.Text = ""
    
    Me.cmdSalvar.Caption = "NOVO"
    Me.txtNome.SetFocus
    
End Sub

Private Sub carregarContratos()
Dim obj As New cMatriz

    AbrirContratos obj.getMatrizSubFrm(carregarBanco, "vw_matriz", sFrm, "CONTRATOS"), sID

Set obj = Nothing

End Sub

Private Sub carregarObras()
Dim obj As New cMatriz

    AbrirObras obj.getMatrizSubFrm(carregarBanco, "vw_matriz", sFrm, "OBRAS"), sID

Set obj = Nothing
    
End Sub

Private Sub carregarContatos()
Dim obj As New cMatriz

    AbrirContatos obj.getMatrizSubFrm(carregarBanco, "vw_matriz", sFrm, "CONTATOS"), sID

Set obj = Nothing

End Sub

Private Sub carregarObservacoes()
Dim obj As New cMatriz

    AbrirObservacoes obj.getMatrizSubFrm(carregarBanco, "vw_matriz", sFrm, "OBSERVAÇÕES"), sID

Set obj = Nothing

End Sub

Private Sub carregarCampos()
    carregarEntidade
End Sub

Private Sub carregarEntidade()
Dim obj As New cEntidades

    camposEntidade obj.getEntidadesID(carregarBanco, sConsulta, sID)
        
Set obj = Nothing
        
End Sub

Private Sub camposEntidade(ByVal obj As cEntidades)
'Dim oEndereco As New cEnderecos

With obj

    Me.txtID.Value = .id
    Me.txtFK.Value = .FK
    
    Me.cboTipo.Value = .CadastroTipo
    Me.txtCNPJ_CPF.Value = .CnpjCpf
    Me.txtIE_RG.Value = .IeRg
    Me.txtNome.Value = .Nome
    Me.txtApelido.Value = .NomeFantasia
    Me.cboStatus.Value = .CadastroStatus
    Me.cboPropaganda.Value = .CadastroPropaganda
    Me.txtObservacao.Value = .CadastroObservacao
    
    Me.txtCep.Text = .ENDERECO.Cep
    Me.txtNumero.Text = .ENDERECO.Numero
    Me.txtComplemento.Text = .ENDERECO.Complemento
    Me.txtLogradouro.Text = .ENDERECO.Logradouro
    Me.txtBairro.Text = .ENDERECO.Bairro
    Me.cboCidade.Text = .ENDERECO.Cidade
    Me.cboEstado.Text = .ENDERECO.Estado

End With

End Sub

Private Function objEntidade() As cEntidades
Dim obj As cEntidades
Set obj = New cEntidades

With obj

    .id = Me.txtID.Value
    .FK = Me.txtFK.Value
    
    .CadastroCategoria = sCategoria
    .Procedure = sProcedure
    
    .CadastroTipo = Me.cboTipo.Value
    .CnpjCpf = Me.txtCNPJ_CPF.Value
    .IeRg = Me.txtIE_RG.Value
    .Nome = Me.txtNome.Value
    .NomeFantasia = Me.txtApelido.Value
    .CadastroStatus = Me.cboStatus.Value
    .CadastroPropaganda = Me.cboPropaganda.Value
    .CadastroObservacao = Me.txtObservacao.Value
    
    .ENDERECO.Cep = Me.txtCep.Text
    .ENDERECO.Numero = Me.txtNumero.Text
    .ENDERECO.Complemento = Me.txtComplemento.Text
    .ENDERECO.Logradouro = Me.txtLogradouro.Text
    .ENDERECO.Bairro = Me.txtBairro.Text
    .ENDERECO.Cidade = Me.cboCidade.Text
    .ENDERECO.Estado = Me.cboEstado.Text
    
    .add obj

End With

Set objEntidade = obj
Set obj = Nothing

End Function
