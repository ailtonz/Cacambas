VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEnderecos"
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

Private pCep As String
Private pNumero As String
Private pComplemento As String
Private pLogradouro As String
Private pBairro As String
Private pCidade As String
Private pEstado As String


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

Public Property Get Cep() As String
    Cep = pCep
End Property

Public Property Let Cep(ByVal vValor As String)
    pCep = vValor
End Property

Public Property Get Numero() As String
    Numero = pNumero
End Property

Public Property Let Numero(ByVal vValor As String)
    pNumero = vValor
End Property

Public Property Get Complemento() As String
    Complemento = pComplemento
End Property

Public Property Let Complemento(ByVal vValor As String)
    pComplemento = vValor
End Property

Public Property Get Logradouro() As String
    Logradouro = pLogradouro
End Property

Public Property Let Logradouro(ByVal vValor As String)
    pLogradouro = vValor
End Property

Public Property Get Bairro() As String
    Bairro = pBairro
End Property

Public Property Let Bairro(ByVal vValor As String)
    pBairro = vValor
End Property

Public Property Get Cidade() As String
    Cidade = pCidade
End Property

Public Property Let Cidade(ByVal vValor As String)
    pCidade = vValor
End Property

Public Property Get Estado() As String
    Estado = pEstado
End Property

Public Property Let Estado(ByVal vValor As String)
    pEstado = vValor
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

Public Sub add(ByVal rec As cEnderecos, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As cEnderecos)
    pCol.add rec, rec.id
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function cEnderecos(ByVal i As Variant) As cEnderecos
    Set cEnderecos = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As cEnderecos
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As cDB, obj As cEnderecos) As Boolean
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
    .Parameters.Append .CreateParameter("@pCadastroCategoria", adVarChar, adParamInput, 50, obj.CadastroCategoria)
        
    .Parameters.Append .CreateParameter("@pEnderecoCep", adVarChar, adParamInput, 8, obj.Cep)
    .Parameters.Append .CreateParameter("@pEnderecoNumero", adVarChar, adParamInput, 10, obj.Numero)
    .Parameters.Append .CreateParameter("@pEnderecoComplemento", adVarChar, adParamInput, 50, obj.Complemento)
    .Parameters.Append .CreateParameter("@pEnderecoLogradouro", adVarChar, adParamInput, 200, obj.Logradouro)
    .Parameters.Append .CreateParameter("@pEnderecoBairro", adVarChar, adParamInput, 50, obj.Bairro)
    .Parameters.Append .CreateParameter("@pEnderecoCidade", adVarChar, adParamInput, 100, obj.Cidade)
    .Parameters.Append .CreateParameter("@pEnderecoEstado", adVarChar, adParamInput, 2, obj.Estado)


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

Public Function Update(banco As cDB, obj As cEnderecos) As Boolean
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
    .Parameters.Append .CreateParameter("@pCadastroCategoria", adVarChar, adParamInput, 50, obj.CadastroCategoria)
        
    .Parameters.Append .CreateParameter("@pEnderecoCep", adVarChar, adParamInput, 8, obj.Cep)
    .Parameters.Append .CreateParameter("@pEnderecoNumero", adVarChar, adParamInput, 10, obj.Numero)
    .Parameters.Append .CreateParameter("@pEnderecoComplemento", adVarChar, adParamInput, 50, obj.Complemento)
    .Parameters.Append .CreateParameter("@pEnderecoLogradouro", adVarChar, adParamInput, 200, obj.Logradouro)
    .Parameters.Append .CreateParameter("@pEnderecoBairro", adVarChar, adParamInput, 50, obj.Bairro)
    .Parameters.Append .CreateParameter("@pEnderecoCidade", adVarChar, adParamInput, 100, obj.Cidade)
    .Parameters.Append .CreateParameter("@pEnderecoEstado", adVarChar, adParamInput, 2, obj.Estado)

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

Public Function Delete(banco As cDB, obj As cEnderecos) As Boolean
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
    .Parameters.Append .CreateParameter("@pCadastroCategoria", adVarChar, adParamInput, 50, obj.CadastroCategoria)
        
    .Parameters.Append .CreateParameter("@pEnderecoCep", adVarChar, adParamInput, 8, obj.Cep)
    .Parameters.Append .CreateParameter("@pEnderecoNumero", adVarChar, adParamInput, 10, obj.Numero)
    .Parameters.Append .CreateParameter("@pEnderecoComplemento", adVarChar, adParamInput, 50, obj.Complemento)
    .Parameters.Append .CreateParameter("@pEnderecoLogradouro", adVarChar, adParamInput, 200, obj.Logradouro)
    .Parameters.Append .CreateParameter("@pEnderecoBairro", adVarChar, adParamInput, 50, obj.Bairro)
    .Parameters.Append .CreateParameter("@pEnderecoCidade", adVarChar, adParamInput, 100, obj.Cidade)
    .Parameters.Append .CreateParameter("@pEnderecoEstado", adVarChar, adParamInput, 2, obj.Estado)


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

Public Function getEnderecos(banco As cDB, qry As String) As cEnderecos
On Error GoTo Err_getEnderecos
Dim colObj As New cEnderecos
Dim recObj As New cEnderecos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qry
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cEnderecos
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .Cep = CStr(IIf(IsNull(rst.Fields("EnderecoCep")), "", rst.Fields("EnderecoCep")))
        .Numero = CStr(IIf(IsNull(rst.Fields("EnderecoNumero")), "", rst.Fields("EnderecoNumero")))
        .Complemento = CStr(IIf(IsNull(rst.Fields("EnderecoComplemento")), "", rst.Fields("EnderecoComplemento")))
        .Logradouro = CStr(IIf(IsNull(rst.Fields("EnderecoLogradouro")), "", rst.Fields("EnderecoLogradouro")))
        .Bairro = CStr(IIf(IsNull(rst.Fields("EnderecoBairro")), "", rst.Fields("EnderecoBairro")))
        .Cidade = CStr(IIf(IsNull(rst.Fields("EnderecoCidade")), "", rst.Fields("EnderecoCidade")))
        .Estado = CStr(IIf(IsNull(rst.Fields("EnderecoEstado")), "", rst.Fields("EnderecoEstado")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getEnderecos:
    Set getEnderecos = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getEnderecos:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getEnderecos()" & vbNewLine & err.Description)
    Resume Exit_getEnderecos

End Function



Public Function getEnderecosCEP(banco As cDB, qry As String, Cep As String) As cEnderecos
On Error GoTo Err_getEnderecosCEP
Dim colObj As New cEnderecos
Dim recObj As New cEnderecos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qry & " where EnderecoCep = " & Cep & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cEnderecos
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .Cep = CStr(IIf(IsNull(rst.Fields("EnderecoCep")), "", rst.Fields("EnderecoCep")))
        .Numero = CStr(IIf(IsNull(rst.Fields("EnderecoNumero")), "", rst.Fields("EnderecoNumero")))
        .Complemento = CStr(IIf(IsNull(rst.Fields("EnderecoComplemento")), "", rst.Fields("EnderecoComplemento")))
        .Logradouro = CStr(IIf(IsNull(rst.Fields("EnderecoLogradouro")), "", rst.Fields("EnderecoLogradouro")))
        .Bairro = CStr(IIf(IsNull(rst.Fields("EnderecoBairro")), "", rst.Fields("EnderecoBairro")))
        .Cidade = CStr(IIf(IsNull(rst.Fields("EnderecoCidade")), "", rst.Fields("EnderecoCidade")))
        .Estado = CStr(IIf(IsNull(rst.Fields("EnderecoEstado")), "", rst.Fields("EnderecoEstado")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getEnderecosCEP:
    Set getEnderecosCEP = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getEnderecosCEP:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getEnderecosCEP()" & vbNewLine & err.Description)
    Resume Exit_getEnderecosCEP

End Function

Public Function getEnderecosObj(banco As cDB, qry As String, id As String) As cEnderecos
On Error GoTo Err_getEnderecosObj
Dim obj As New cEnderecos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qry & " where ID = " & id & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    With obj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .Cep = CStr(IIf(IsNull(rst.Fields("EnderecoCep")), "", rst.Fields("EnderecoCep")))
        .Numero = CStr(IIf(IsNull(rst.Fields("EnderecoNumero")), "", rst.Fields("EnderecoNumero")))
        .Complemento = CStr(IIf(IsNull(rst.Fields("EnderecoComplemento")), "", rst.Fields("EnderecoComplemento")))
        .Logradouro = CStr(IIf(IsNull(rst.Fields("EnderecoLogradouro")), "", rst.Fields("EnderecoLogradouro")))
        .Bairro = CStr(IIf(IsNull(rst.Fields("EnderecoBairro")), "", rst.Fields("EnderecoBairro")))
        .Cidade = CStr(IIf(IsNull(rst.Fields("EnderecoCidade")), "", rst.Fields("EnderecoCidade")))
        .Estado = CStr(IIf(IsNull(rst.Fields("EnderecoEstado")), "", rst.Fields("EnderecoEstado")))
        
        .add obj
    End With
    rst.MoveNext
Loop


Exit_getEnderecosObj:
    Set getEnderecosObj = obj
    Set obj = Nothing
    
    Exit Function

Err_getEnderecosObj:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getEnderecosObj()" & vbNewLine & err.Description)
    Resume Exit_getEnderecosObj

End Function


Public Function getEnderecosID(banco As cDB, qry As String, id As String) As cEnderecos
On Error GoTo Err_getEnderecosID
Dim colObj As New cEnderecos
Dim recObj As New cEnderecos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qry & " where ID = " & id & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cEnderecos
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .Cep = CStr(IIf(IsNull(rst.Fields("EnderecoCep")), "", rst.Fields("EnderecoCep")))
        .Numero = CStr(IIf(IsNull(rst.Fields("EnderecoNumero")), "", rst.Fields("EnderecoNumero")))
        .Complemento = CStr(IIf(IsNull(rst.Fields("EnderecoComplemento")), "", rst.Fields("EnderecoComplemento")))
        .Logradouro = CStr(IIf(IsNull(rst.Fields("EnderecoLogradouro")), "", rst.Fields("EnderecoLogradouro")))
        .Bairro = CStr(IIf(IsNull(rst.Fields("EnderecoBairro")), "", rst.Fields("EnderecoBairro")))
        .Cidade = CStr(IIf(IsNull(rst.Fields("EnderecoCidade")), "", rst.Fields("EnderecoCidade")))
        .Estado = CStr(IIf(IsNull(rst.Fields("EnderecoEstado")), "", rst.Fields("EnderecoEstado")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getEnderecosID:
    Set getEnderecosID = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getEnderecosID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getEnderecosID()" & vbNewLine & err.Description)
    Resume Exit_getEnderecosID

End Function

Public Function getEnderecosFK(banco As cDB, qry As String, id As String) As cEnderecos
On Error GoTo Err_getEnderecosFK
Dim colObj As New cEnderecos
Dim recObj As New cEnderecos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qry & " where FK = " & id & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cEnderecos
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .Cep = CStr(IIf(IsNull(rst.Fields("EnderecoCep")), "", rst.Fields("EnderecoCep")))
        .Numero = CStr(IIf(IsNull(rst.Fields("EnderecoNumero")), "", rst.Fields("EnderecoNumero")))
        .Complemento = CStr(IIf(IsNull(rst.Fields("EnderecoComplemento")), "", rst.Fields("EnderecoComplemento")))
        .Logradouro = CStr(IIf(IsNull(rst.Fields("EnderecoLogradouro")), "", rst.Fields("EnderecoLogradouro")))
        .Bairro = CStr(IIf(IsNull(rst.Fields("EnderecoBairro")), "", rst.Fields("EnderecoBairro")))
        .Cidade = CStr(IIf(IsNull(rst.Fields("EnderecoCidade")), "", rst.Fields("EnderecoCidade")))
        .Estado = CStr(IIf(IsNull(rst.Fields("EnderecoEstado")), "", rst.Fields("EnderecoEstado")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getEnderecosFK:
    Set getEnderecosFK = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getEnderecosFK:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getEnderecosFK()" & vbNewLine & err.Description)
    Resume Exit_getEnderecosFK

End Function

