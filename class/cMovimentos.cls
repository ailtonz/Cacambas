VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMovimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pFK As String
Private pProcedure As String
Private pCadastroCategoria As String

Private pDataDeEmissao As String
Private pDocumento As String
Private pObservacao As String
Private pDataDeVencimento As String
Private pValorOriginal As String
Private pDataDePagamento As String
Private pValorFinal As String
Private pMovimento As String
Private pGrupo As String
Private pConta As String
Private pTransacao As String
Private pFrequencia As String

Private pAno As String
Private pMes As String
Private pRef As String
Private pPlano As String


''---------------
'' propriedades
''---------------

Public Property Get id() As String
    id = pID
End Property

Public Property Let id(ByVal vValor As String)
    pID = vValor
End Property

Public Property Get FK() As String
    FK = pFK
End Property

Public Property Let FK(ByVal vValor As String)
    pFK = vValor
End Property

Public Property Get Procedure() As String
    Procedure = pProcedure
End Property

Public Property Let Procedure(ByVal vValor As String)
    pProcedure = vValor
End Property

Public Property Get CadastroCategoria() As String
    CadastroCategoria = pCadastroCategoria
End Property

Public Property Let CadastroCategoria(ByVal vValor As String)
    pCadastroCategoria = vValor
End Property

Public Property Get DataDeEmissao() As String
    DataDeEmissao = pDataDeEmissao
End Property

Public Property Let DataDeEmissao(ByVal vValor As String)
    pDataDeEmissao = vValor
End Property

Public Property Get Documento() As String
    Documento = pDocumento
End Property

Public Property Let Documento(ByVal vValor As String)
    pDocumento = vValor
End Property

Public Property Get Observacao() As String
    Observacao = pObservacao
End Property

Public Property Let Observacao(ByVal vValor As String)
    pObservacao = vValor
End Property

Public Property Get DataDeVencimento() As String
    DataDeVencimento = pDataDeVencimento
End Property

Public Property Let DataDeVencimento(ByVal vValor As String)
    pDataDeVencimento = vValor
End Property

Public Property Get ValorOriginal() As String
    ValorOriginal = pValorOriginal
End Property

Public Property Let ValorOriginal(ByVal vValor As String)
    pValorOriginal = vValor
End Property

Public Property Get DataDePagamento() As String
    DataDePagamento = pDataDePagamento
End Property

Public Property Let DataDePagamento(ByVal vValor As String)
    pDataDePagamento = vValor
End Property

Public Property Get ValorFinal() As String
    ValorFinal = pValorFinal
End Property

Public Property Let ValorFinal(ByVal vValor As String)
    pValorFinal = vValor
End Property

Public Property Get Movimento() As String
    Movimento = pMovimento
End Property

Public Property Let Movimento(ByVal vValor As String)
    pMovimento = vValor
End Property

Public Property Get Grupo() As String
    Grupo = pGrupo
End Property

Public Property Let Grupo(ByVal vValor As String)
    pGrupo = vValor
End Property

Public Property Get Conta() As String
    Conta = pConta
End Property

Public Property Let Conta(ByVal vValor As String)
    pConta = vValor
End Property

Public Property Get Transacao() As String
    Transacao = pTransacao
End Property

Public Property Let Transacao(ByVal vValor As String)
    pTransacao = vValor
End Property

Public Property Get Frequencia() As String
    Frequencia = pFrequencia
End Property

Public Property Let Frequencia(ByVal vValor As String)
    pFrequencia = vValor
End Property

Public Property Get Ano() As String
    Ano = pAno
End Property

Public Property Let Ano(ByVal vValor As String)
    pAno = vValor
End Property

Public Property Get Mes() As String
    Mes = pMes
End Property

Public Property Let Mes(ByVal vValor As String)
    pMes = vValor
End Property

Public Property Get Ref() As String
    Ref = pRef
End Property

Public Property Let Ref(ByVal vValor As String)
    pRef = vValor
End Property

Public Property Get Plano() As String
    Plano = pPlano
End Property

Public Property Let Plano(ByVal vValor As String)
    pPlano = vValor
End Property

''---------------
'' administração
''---------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCol = Nothing
End Sub

Public Sub add(ByVal rec As cMovimentos, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As cMovimentos)
    pCol.add rec, rec.id
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function cMovimentos(ByVal i As Variant) As cMovimentos
    Set cMovimentos = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As cMovimentos
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As cDB, obj As cMovimentos) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = obj.Procedure
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@pID", adVarChar, adParamInput, 50, obj.id)
    .Parameters.Append .CreateParameter("@pFK", adVarChar, adParamInput, 50, obj.FK)
    
'    .Parameters.Append .CreateParameter("@pDataDeEmissao", adVarChar, adParamInput, 10, obj.DataDeEmissao)
    .Parameters.Append .CreateParameter("@pDocumento", adVarChar, adParamInput, 50, obj.Documento)
    .Parameters.Append .CreateParameter("@pObservacao", adVarChar, adParamInput, 100, obj.Observacao)
    .Parameters.Append .CreateParameter("@pDataDeVencimento", adVarChar, adParamInput, 10, obj.DataDeVencimento)
    .Parameters.Append .CreateParameter("@pValorOriginal", adVarChar, adParamInput, 50, obj.ValorOriginal)
    .Parameters.Append .CreateParameter("@pDataDePagamento", adVarChar, adParamInput, 10, obj.DataDePagamento)
    .Parameters.Append .CreateParameter("@pValorFinal", adVarChar, adParamInput, 50, obj.ValorFinal)
    .Parameters.Append .CreateParameter("@pMovimento", adVarChar, adParamInput, 50, obj.Movimento)
    .Parameters.Append .CreateParameter("@pGrupo", adVarChar, adParamInput, 50, obj.Grupo)
    .Parameters.Append .CreateParameter("@pConta", adVarChar, adParamInput, 50, obj.Conta)
    .Parameters.Append .CreateParameter("@pTransacao", adVarChar, adParamInput, 50, obj.Transacao)
    .Parameters.Append .CreateParameter("@pFrequencia", adVarChar, adParamInput, 50, obj.Frequencia)

    Set rst = .Execute(lRecordsAffected)
End With
Connection.Close

  If lRecordsAffected = 0 Then
    Insert = False
  Else
    'Return a good return code
    Insert = True
  End If

Exit_Insert:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Insert:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Insert()" & vbNewLine & err.Description)
  Insert = False
  Resume Exit_Insert


End Function

Public Function Update(banco As cDB, obj As cMovimentos) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = obj.Procedure
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@pID", adVarChar, adParamInput, 50, obj.id)
    .Parameters.Append .CreateParameter("@pFK", adVarChar, adParamInput, 50, obj.FK)
    
    .Parameters.Append .CreateParameter("@pDataDeEmissao", adVarChar, adParamInput, 10, obj.DataDeEmissao)
    .Parameters.Append .CreateParameter("@pDocumento", adVarChar, adParamInput, 50, obj.Documento)
    .Parameters.Append .CreateParameter("@pObservacao", adVarChar, adParamInput, 100, obj.Observacao)
    .Parameters.Append .CreateParameter("@pDataDeVencimento", adVarChar, adParamInput, 10, obj.DataDeVencimento)
    .Parameters.Append .CreateParameter("@pValorOriginal", adVarChar, adParamInput, 50, obj.ValorOriginal)
    .Parameters.Append .CreateParameter("@pDataDePagamento", adVarChar, adParamInput, 10, obj.DataDePagamento)
    .Parameters.Append .CreateParameter("@pValorFinal", adVarChar, adParamInput, 50, obj.ValorFinal)
    .Parameters.Append .CreateParameter("@pMovimento", adVarChar, adParamInput, 50, obj.Movimento)
    .Parameters.Append .CreateParameter("@pGrupo", adVarChar, adParamInput, 50, obj.Grupo)
    .Parameters.Append .CreateParameter("@pConta", adVarChar, adParamInput, 50, obj.Conta)
    .Parameters.Append .CreateParameter("@pTransacao", adVarChar, adParamInput, 50, obj.Transacao)
    .Parameters.Append .CreateParameter("@pFrequencia", adVarChar, adParamInput, 50, obj.Frequencia)


    Set rst = .Execute(lRecordsAffected)
End With
Connection.Close

  If lRecordsAffected = 0 Then
    Update = False
  Else
    'Return a good return code
    Update = True
  End If

Exit_Update:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Update:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Update()" & vbNewLine & err.Description)
  Update = False
  Resume Exit_Update

End Function

Public Function Delete(banco As cDB, obj As cMovimentos) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = obj.Procedure
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@pID", adVarChar, adParamInput, 50, obj.id)
    .Parameters.Append .CreateParameter("@pFK", adVarChar, adParamInput, 50, obj.FK)
    
    .Parameters.Append .CreateParameter("@pDataDeEmissao", adVarChar, adParamInput, 10, obj.DataDeEmissao)
    .Parameters.Append .CreateParameter("@pDocumento", adVarChar, adParamInput, 50, obj.Documento)
    .Parameters.Append .CreateParameter("@pObservacao", adVarChar, adParamInput, 100, obj.Observacao)
    .Parameters.Append .CreateParameter("@pDataDeVencimento", adVarChar, adParamInput, 10, obj.DataDeVencimento)
    .Parameters.Append .CreateParameter("@pValorOriginal", adVarChar, adParamInput, 50, obj.ValorOriginal)
    .Parameters.Append .CreateParameter("@pDataDePagamento", adVarChar, adParamInput, 10, obj.DataDePagamento)
    .Parameters.Append .CreateParameter("@pValorFinal", adVarChar, adParamInput, 50, obj.ValorFinal)
    .Parameters.Append .CreateParameter("@pMovimento", adVarChar, adParamInput, 50, obj.Movimento)
    .Parameters.Append .CreateParameter("@pGrupo", adVarChar, adParamInput, 50, obj.Grupo)
    .Parameters.Append .CreateParameter("@pConta", adVarChar, adParamInput, 50, obj.Conta)
    .Parameters.Append .CreateParameter("@pTransacao", adVarChar, adParamInput, 50, obj.Transacao)
    .Parameters.Append .CreateParameter("@pFrequencia", adVarChar, adParamInput, 50, obj.Frequencia)


    Set rst = .Execute(lRecordsAffected)
End With
Connection.Close

  If lRecordsAffected = 0 Then
    Delete = False
  Else
    'Return a good return code
    Delete = True
  End If

Exit_Delete:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Delete:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Delete()" & vbNewLine & err.Description)
  Delete = False
  Resume Exit_Delete

End Function

Public Function getMovimentos(banco As cDB, qryEntidade As String) As cMovimentos
On Error GoTo Err_getMovimentos
Dim colObj As New cMovimentos
Dim recObj As New cMovimentos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cMovimentos
    With recObj

        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))

        .DataDeEmissao = CStr(rst.Fields("DataDeEmissao"))
        .Documento = CStr(IIf(IsNull(rst.Fields("Documento")), "", rst.Fields("Documento")))
        .Observacao = CStr(IIf(IsNull(rst.Fields("Observacao")), "", rst.Fields("Observacao")))
        .DataDeVencimento = CStr(rst.Fields("DataDeVencimento"))
        .ValorOriginal = CStr(rst.Fields("ValorOriginal"))
        .DataDePagamento = CStr(rst.Fields("DataDePagamento"))
        .ValorFinal = CStr(rst.Fields("ValorFinal"))
        .Movimento = CStr(IIf(IsNull(rst.Fields("Movimento")), "", rst.Fields("Movimento")))
        .Grupo = CStr(IIf(IsNull(rst.Fields("Plano")), "", rst.Fields("Plano")))
        .Conta = CStr(IIf(IsNull(rst.Fields("Conta")), "", rst.Fields("Conta")))
        .Transacao = CStr(IIf(IsNull(rst.Fields("Transacao")), "", rst.Fields("Transacao")))
        .Frequencia = CStr(IIf(IsNull(rst.Fields("Frequencia")), "", rst.Fields("Frequencia")))

        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getMovimentos:
    Set getMovimentos = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getMovimentos:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getMovimentos()" & vbNewLine & err.Description)
    Resume Exit_getMovimentos

End Function


Public Function getMovimentosID(banco As cDB, qryEntidade As String, i As String) As cMovimentos
On Error GoTo Err_getMovimentosID
Dim colObj As New cMovimentos
Dim recObj As New cMovimentos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where id = " & i & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cMovimentos
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))

        .DataDeEmissao = CStr(rst.Fields("DataDeEmissao"))
        .Documento = CStr(IIf(IsNull(rst.Fields("Documento")), "", rst.Fields("Documento")))
        .Observacao = CStr(IIf(IsNull(rst.Fields("Observacao")), "", rst.Fields("Observacao")))
        .DataDeVencimento = CStr(rst.Fields("DataDeVencimento"))
        .ValorOriginal = CStr(rst.Fields("ValorOriginal"))
        .DataDePagamento = CStr(rst.Fields("DataDePagamento"))
        .ValorFinal = CStr(rst.Fields("ValorFinal"))
        .Movimento = CStr(IIf(IsNull(rst.Fields("Movimento")), "", rst.Fields("Movimento")))
        .Grupo = CStr(IIf(IsNull(rst.Fields("Plano")), "", rst.Fields("Plano")))
        .Conta = CStr(IIf(IsNull(rst.Fields("Conta")), "", rst.Fields("Conta")))
        .Transacao = CStr(IIf(IsNull(rst.Fields("Transacao")), "", rst.Fields("Transacao")))
        .Frequencia = CStr(IIf(IsNull(rst.Fields("Frequencia")), "", rst.Fields("Frequencia")))
                
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getMovimentosID:
    Set getMovimentosID = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getMovimentosID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getMovimentosID()" & vbNewLine & err.Description)
    Resume Exit_getMovimentosID

End Function

Public Function getMovimentosDados(banco As cDB, qryEntidade As String) As cMovimentos
On Error GoTo Err_getMovimentosDados
Dim colObj As New cMovimentos
Dim recObj As New cMovimentos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cMovimentos
    With recObj

        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))

        .DataDeEmissao = CStr(rst.Fields("DataDeEmissao"))
        .Documento = CStr(IIf(IsNull(rst.Fields("Documento")), "", rst.Fields("Documento")))
        .Observacao = CStr(IIf(IsNull(rst.Fields("Observacao")), "", rst.Fields("Observacao")))
        .DataDeVencimento = CStr(rst.Fields("DataDeVencimento"))
        .ValorOriginal = CStr(rst.Fields("ValorOriginal"))
        .DataDePagamento = CStr(rst.Fields("DataDePagamento"))
        .ValorFinal = CStr(rst.Fields("ValorFinal"))
        .Movimento = CStr(IIf(IsNull(rst.Fields("Movimento")), "", rst.Fields("Movimento")))
        .Grupo = CStr(IIf(IsNull(rst.Fields("Plano")), "", rst.Fields("Plano")))
        .Conta = CStr(IIf(IsNull(rst.Fields("Conta")), "", rst.Fields("Conta")))
        .Transacao = CStr(IIf(IsNull(rst.Fields("Transacao")), "", rst.Fields("Transacao")))
        .Frequencia = CStr(IIf(IsNull(rst.Fields("Frequencia")), "", rst.Fields("Frequencia")))
        
        .Ano = CStr(IIf(IsNull(rst.Fields("Ano")), "", rst.Fields("Ano")))
        .Mes = CStr(IIf(IsNull(rst.Fields("Mes")), "", rst.Fields("Mes")))
        .Ref = CStr(IIf(IsNull(rst.Fields("Ref")), "", rst.Fields("Ref")))
        .Plano = CStr(IIf(IsNull(rst.Fields("Plano")), "", rst.Fields("Plano")))
        

        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getMovimentosDados:
    Set getMovimentosDados = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getMovimentosDados:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getMovimentosDados()" & vbNewLine & err.Description)
    Resume Exit_getMovimentosDados

End Function
