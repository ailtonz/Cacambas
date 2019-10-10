VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNotas 
   Caption         =   "NOTAS"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610.001
   OleObjectBlob   =   "frmNotas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNotas"
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

Dim obj As cNotas
Set obj = New cNotas

    With obj
        .id = CStr(Me.txtID.Value)
        .FK = CStr("0")
        
        If Me.cmdSalvar.Caption = "EXCLUIR" Then
            .Titulo = CStr("")
        Else
            .Titulo = CStr(Me.txtTitulo.Value)
        End If
        
        .Descricao = CStr(Me.txtDescricao.Value)
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

    limparCampos
    listarRegistros
    

Set obj = Nothing


End Sub

Private Sub listarRegistros()
Dim Prf As cNotas
Dim col As cNotas

Set Prf = New cNotas

Set col = Prf.getNotas(carregarBanco, sConsulta)

With Me.lstRegistros
    .Clear
    .ColumnCount = 3
    .ColumnWidths = "0;90;0"
    
    For Each Prf In col.Itens
        .AddItem Prf.id
        .List(.ListCount - 1, 1) = Prf.Titulo
        .List(.ListCount - 1, 2) = Prf.Descricao
        
    Next Prf

End With

End Sub

Private Sub limparCampos()
    
    Me.txtID.Value = "0"
    Me.txtTitulo.Value = ""
    Me.txtDescricao.Value = ""
    
    Me.cmdSalvar.Caption = "NOVO"
    Me.txtTitulo.SetFocus
    
End Sub

Private Sub carregarCampos()

    Me.txtID.Value = Me.lstRegistros.Value
    Me.txtTitulo.Value = Me.lstRegistros.Column(1)
    Me.txtDescricao.Value = Me.lstRegistros.Column(2)
        
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


