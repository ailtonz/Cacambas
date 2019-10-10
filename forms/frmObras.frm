VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmObras 
   Caption         =   "OBRAS"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10995
   OleObjectBlob   =   "frmObras.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmObras"
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
'    Me.fmDados.SetFocus
    limparCampos
    carregarCampos
End Sub

Private Sub cmdCancelar_Click()
'    limparCampos
End Sub

Private Sub cmdSalvar_Click()
    salvarRegistro
End Sub

Private Sub salvarRegistro()
'Dim ws As Worksheet
'Set ws = Worksheets(ActiveSheet.Name)
'
'Dim Obj As cEnderecos
'Set Obj = New cEnderecos
'
'    With Obj
'        .id = CStr(Me.txtID.Value)
'        .FK = CStr("0")
'
'        If Me.cmdSalvar.Caption = "EXCLUIR" Then
'            .Nome = CStr("")
'        Else
'            .Nome = CStr(Me.txtNome.Value)
'        End If
'
'        .NomeFantasia = CStr(Me.txtApelido.Value)
'        .CadastroCategoria = sCatPrimario
'        .Procedure = sPrdPrimario
'
'        .add Obj
'    End With
'
'
'    If Me.cmdSalvar.Caption = "NOVO" Then
'        If (Obj.Insert(carregarBanco, Obj) = True) Then
'            MsgBox "Cadastro realizado com sucesso!", vbInformation + vbOKOnly, "Cadastro"
'        Else
'            MsgBox "Não foi possivel realizar o cadastro!", vbCritical + vbOKOnly, "Cadastro - ERRO!"
'        End If
'    ElseIf Me.cmdSalvar.Caption = "SALVAR" Then
'        If (Obj.Update(carregarBanco, Obj) = True) Then
'            MsgBox "Alteração realizada com sucesso!", vbInformation + vbOKOnly, "Alteração"
'        Else
'            MsgBox "Não foi possivel realizar alteração!", vbCritical + vbOKOnly, "Alteração - ERRO!"
'        End If
'    ElseIf Me.cmdSalvar.Caption = "EXCLUIR" Then
'        If mostrarRegistro = vbYes Then
'            If (Obj.Delete(carregarBanco, Obj) = True) Then
'                MsgBox "Exclusão realizada com sucesso!", vbInformation + vbOKOnly, "Exclusão"
'            Else
'                MsgBox "Não foi possivel realizar Exclusão!", vbCritical + vbOKOnly, "Exclusão - ERRO!"
'            End If
'        End If
'    End If
'
'    limparCampos
'
'Set Obj = Nothing

End Sub

Private Sub limparCampos()
    
    Me.txtID.Value = "0"
    Me.txtFK.Value = "0"
    
'    Me.cboTipo.Text = ""
'    Me.cboStatus.Text = ""
'
'    Me.txtCNPJ_CPF.Text = ""
'    Me.txtIE_RG.Text = ""
'
'    Me.txtNome.Text = ""
'    Me.txtApelido.Text = ""
'    Me.txtObservacao.Text = ""
'
'    Me.txtCep.Text = ""
'    Me.txtNumero.Text = ""
'    Me.txtComplemento.Text = ""
'    Me.txtLogradouro.Text = ""
'    Me.cboCidade.Text = ""
'    Me.cboEstado.Text = ""
        
    
    Me.cmdSalvar.Caption = "NOVO"
'    Me.txtNome.SetFocus
    
End Sub

Private Sub cmdContatos_Click()
    carregarContatos
End Sub

Private Sub cmdContratos_Click()
    carregarContratos
End Sub

Private Sub cmdObservacoes_Click()
    carregarObservacoes
End Sub

Private Sub carregarContratos()
Dim obj As New cMatriz

    AbrirContratos obj.getMatrizSubFrm(carregarBanco, "vw_matriz", sFrm, "CONTRATOS"), sID

Set obj = Nothing

End Sub

Private Sub carregarObservacoes()
Dim obj As New cMatriz

    AbrirObras obj.getMatrizSubFrm(carregarBanco, "vw_matriz", sFrm, "OBSERVAÇÕES"), sID

Set obj = Nothing

End Sub

Private Sub carregarContatos()
Dim obj As New cMatriz

    AbrirObras obj.getMatrizSubFrm(carregarBanco, "vw_matriz", sFrm, "CONTATOS"), sID

Set obj = Nothing

End Sub

Private Sub carregarCampos()
    carregarEndereco
    carregarCobranca
End Sub

Private Sub carregarEndereco()
Dim obj As New cEnderecos

    camposEndereco obj.getEnderecosObj(carregarBanco, sConsulta, sID)

Set obj = Nothing

End Sub

Private Sub camposEndereco(ByVal obj As cEnderecos)

With obj

    Me.txtCep.Text = .Cep
    Me.txtNumero.Text = .Numero
    Me.txtComplemento.Text = .Complemento
    Me.txtLogradouro.Text = .Logradouro
    Me.cboCidade.Text = .Cidade
    Me.cboEstado.Text = .Estado

End With

End Sub

Private Sub carregarCobranca()
Dim obj As New cCobranca

    camposCobranca obj.getCobrancaObj(carregarBanco, sConsulta, sID)

Set obj = Nothing

End Sub

Private Sub camposCobranca(ByVal obj As cCobranca)

With obj

    Me.txtCobrancaCep.Text = .CobrancaCep
    Me.txtCobrancaLogradouro.Text = .CobrancaLogradouro
    Me.cboCobrancaCidade.Text = .CobrancaCidade
    Me.cboCobrancaEstado.Text = .CobrancaEstado

End With

End Sub
