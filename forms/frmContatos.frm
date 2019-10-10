VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContatos 
   Caption         =   "CONTATOS"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8220.001
   OleObjectBlob   =   "frmContatos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmContatos"
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
    listarRegistros
    limparCampos
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub
Private Sub cmdSalvar_Click()
    salvarRegistro
End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '' CARREGAR REGISTRO
    If Not IsNull(Me.lstRegistros.Value) Then
        carregarCampos
        Me.cmdSalvar.Caption = "SALVAR"
    End If
    
End Sub

Private Sub lstRegistros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '' EXCLUIR REGISTRO
    If KeyCode = vbKeyDelete Then
        If Not IsNull(Me.lstRegistros.Value) Then
            carregarCampos
            Me.cmdSalvar.Caption = "EXCLUIR"
            Me.cmdSalvar.SetFocus
        End If
    End If
End Sub

Private Sub salvarRegistro()
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

Dim obj As cContatos
Set obj = New cContatos

    With obj
        .id = CStr(Me.txtID.Value)
        .FK = CStr(Me.txtFK.Value)
        
        If Me.cmdSalvar.Caption = "EXCLUIR" Then
            .ContatoNome = CStr("")
        Else
            .ContatoNome = CStr(Me.txtNome.Value)
        End If
        
        .ContatoTelefone = CStr(Me.txtTelefone.Value)
        .ContatoEmail = CStr(Me.txtEmail.Value)
        .ContatoObservacao = CStr(Me.txtObservacoes.Value)
        
        .CadastroCategoria = sCategoria
        .Procedure = sProcedure
        
        .add obj
    End With
    

    If Me.cmdSalvar.Caption = "NOVO" Then
        If (obj.Insert(carregarBanco, obj) = True) Then
            MsgBox "Cadastro realizado com sucesso!", vbInformation + vbOKOnly, "Cadastro"
        Else
            MsgBox "Não foi possivel realizar o cadastro!", vbCritical + vbOKOnly, "Cadastro - ERRO!"
        End If
    ElseIf Me.cmdSalvar.Caption = "SALVAR" Then
        If (obj.Update(carregarBanco, obj) = True) Then
            MsgBox "Alteração realizada com sucesso!", vbInformation + vbOKOnly, "Alteração"
        Else
            MsgBox "Não foi possivel realizar alteração!", vbCritical + vbOKOnly, "Alteração - ERRO!"
        End If
    ElseIf Me.cmdSalvar.Caption = "EXCLUIR" Then
        If mostrarRegistro = vbYes Then
            If (obj.Delete(carregarBanco, obj) = True) Then
                MsgBox "Exclusão realizada com sucesso!", vbInformation + vbOKOnly, "Exclusão"
            Else
                MsgBox "Não foi possivel realizar Exclusão!", vbCritical + vbOKOnly, "Exclusão - ERRO!"
            End If
        End If
    End If

    listarRegistros
    limparCampos

Set obj = Nothing


End Sub

Private Sub listarRegistros()
Dim Prf As cContatos
Dim col As cContatos

Set Prf = New cContatos

Set col = Prf.getContatosID(carregarBanco, sConsulta, sFK)

With Me.lstRegistros
    .Clear
    .ColumnCount = 6
    .ColumnWidths = "0;0;100;100;100;90"
    
    For Each Prf In col.Itens
        .AddItem Prf.id
        .List(.ListCount - 1, 1) = Prf.FK
                
        .List(.ListCount - 1, 2) = Prf.ContatoNome
        .List(.ListCount - 1, 3) = Prf.ContatoTelefone
        .List(.ListCount - 1, 4) = Prf.ContatoEmail
        .List(.ListCount - 1, 5) = Prf.ContatoObservacao
        
    Next Prf

End With

End Sub

Private Sub limparCampos()
    
    Me.txtID.Value = "0"
    Me.txtFK.Value = sFK
        
    Me.txtNome.Value = ""
    Me.txtTelefone.Value = ""
    Me.txtEmail.Value = ""
    Me.txtObservacoes.Value = ""
        
    Me.cmdSalvar.Caption = "NOVO"
    Me.txtNome.SetFocus
    
End Sub

Private Sub carregarCampos()

    Me.txtID.Value = Me.lstRegistros.Value
    Me.txtFK.Value = Me.lstRegistros.Column(1)
        
    Me.txtNome.Value = Me.lstRegistros.Column(2)
    Me.txtTelefone.Value = Me.lstRegistros.Column(3)
    Me.txtEmail.Value = Me.lstRegistros.Column(4)
    Me.txtObservacoes.Value = Me.lstRegistros.Column(5)
        
End Sub

Private Function mostrarRegistro() As Variant
Dim retVal As Variant
Dim strTitulo As String: strTitulo = IIf(Not IsNull(Me.lstRegistros.Column(1)), Me.lstRegistros.Column(1), 0)
Dim strDescricao As String: strDescricao = IIf(Not IsNull(Me.lstRegistros.Column(2)), Me.lstRegistros.Column(2), "")


    retVal = MsgBox("Você deseja realmente EXCLUIR o registro abaixo:" & vbNewLine & _
            vbNewLine & _
            "TITULO: " & strTitulo & vbNewLine & _
            "DESCRIÇÃO : " & strDescricao & vbNewLine & _
            vbNewLine, vbCritical + vbYesNo, "EXCLUSÃO DE REGISTRO!")
            
    mostrarRegistro = retVal
            
Set retVal = Nothing

End Function



