VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEntidades"
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

Private pCnpjCpf As String
Private pIeRg As String
Private pNome As String
Private pNomeFantasia As String
Private pCadastroTipo As String
Private pCadastroStatus As String
Private pCadastroPropaganda As String
Private pCadastroObservacao As String

Private pENDERECO As cEnderecos

''---------------
'' propriedades
''---------------
Public Property Get ENDERECO() As cEnderecos
    ENDERECO = pENDERECO
End Property

Public Property Let ENDERECO(ByVal vValor As cEnderecos)
    pENDERECO = vValor
End Property


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

Public Property Get CnpjCpf() As String
    CnpjCpf = pCnpjCpf
End Property

Public Property Let CnpjCpf(ByVal vValor As String)
    pCnpjCpf = vValor
End Property

Public Property Get IeRg() As String
    IeRg = pIeRg
End Property

Public Property Let IeRg(ByVal vValor As String)
    pIeRg = vValor
End Property

Public Property Get Nome() As String
    Nome = pNome
End Property

Public Property Let Nome(ByVal vValor As String)
    pNome = vValor
End Property

Public Property Get NomeFantasia() As String
    NomeFantasia = pNomeFantasia
End Property

Public Property Let NomeFantasia(ByVal vValor As String)
    pNomeFantasia = vValor
End Property

Public Property Get CadastroTipo() As String
    CadastroTipo = pCadastroTipo
End Property

Public Property Let CadastroTipo(ByVal vValor As String)
    pCadastroTipo = vValor
End Property

Public Property Get CadastroStatus() As String
    CadastroStatus = pCadastroStatus
End Property

Public Property Let CadastroStatus(ByVal vValor As String)
    pCadastroStatus = vValor
End Property

Public Property Get CadastroPropaganda() As String
    CadastroPropaganda = pCadastroPropaganda
End Property

Public Property Let CadastroPropaganda(ByVal vValor As String)
    pCadastroPropaganda = vValor
End Property

Public Property Get CadastroObservacao() As String
    CadastroObservacao = pCadastroObservacao
End Property

Public Property Let CadastroObservacao(ByVal vValor As String)
    pCadastroObservacao = vValor
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

Public Sub add(ByVal rec As cEntidades, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As cEntidades)
    pCol.add rec, rec.id
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function cEntidades(ByVal i As Variant) As cEntidades
    Set cEntidades = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As cEntidades
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As cDB, obj As cEntidades) As Boolean
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

    .Parameters.Append .CreateParameter("@pCadastroTipo", adVarChar, adParamInput, 2, obj.CadastroTipo)
    .Parameters.Append .CreateParameter("@pCnpjCpf", adVarChar, adParamInput, 14, obj.CnpjCpf)
    .Parameters.Append .CreateParameter("@pIeRg", adVarChar, adParamInput, 16, obj.IeRg)
    .Parameters.Append .CreateParameter("@pNome", adVarChar, adParamInput, 100, obj.Nome)
    .Parameters.Append .CreateParameter("@pNomeFantasia", adVarChar, adParamInput, 50, obj.NomeFantasia)
    .Parameters.Append .CreateParameter("@pCadastroStatus", adVarChar, adParamInput, 10, obj.CadastroStatus)
    .Parameters.Append .CreateParameter("@pCadastroPropaganda", adVarChar, adParamInput, 50, obj.CadastroPropaganda)
    .Parameters.Append .CreateParameter("@pCadastroObservacao", adVarChar, adParamInput, 1000, obj.CadastroObservacao)
    
    .Parameters.Append .CreateParameter("@pEnderecoCep", adVarChar, adParamInput, 8, obj.ENDERECO.Cep)
    .Parameters.Append .CreateParameter("@pEnderecoNumero", adVarChar, adParamInput, 10, obj.ENDERECO.Numero)
    .Parameters.Append .CreateParameter("@pEnderecoComplemento", adVarChar, adParamInput, 50, obj.ENDERECO.Complemento)
    .Parameters.Append .CreateParameter("@pEnderecoLogradouro", adVarChar, adParamInput, 200, obj.ENDERECO.Logradouro)
    .Parameters.Append .CreateParameter("@pEnderecoBairro", adVarChar, adParamInput, 50, obj.ENDERECO.Bairro)
    .Parameters.Append .CreateParameter("@pEnderecoCidade", adVarChar, adParamInput, 100, obj.ENDERECO.Cidade)
    .Parameters.Append .CreateParameter("@pEnderecoEstado", adVarChar, adParamInput, 2, obj.ENDERECO.Estado)

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

Public Function Update(banco As cDB, obj As cEntidades) As Boolean
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

    .Parameters.Append .CreateParameter("@pCadastroTipo", adVarChar, adParamInput, 2, obj.CadastroTipo)
    .Parameters.Append .CreateParameter("@pCnpjCpf", adVarChar, adParamInput, 14, obj.CnpjCpf)
    .Parameters.Append .CreateParameter("@pIeRg", adVarChar, adParamInput, 16, obj.IeRg)
    .Parameters.Append .CreateParameter("@pNome", adVarChar, adParamInput, 100, obj.Nome)
    .Parameters.Append .CreateParameter("@pNomeFantasia", adVarChar, adParamInput, 50, obj.NomeFantasia)
    .Parameters.Append .CreateParameter("@pCadastroStatus", adVarChar, adParamInput, 10, obj.CadastroStatus)
    .Parameters.Append .CreateParameter("@pCadastroPropaganda", adVarChar, adParamInput, 50, obj.CadastroPropaganda)
    .Parameters.Append .CreateParameter("@pCadastroObservacao", adVarChar, adParamInput, 1000, obj.CadastroObservacao)
    
    .Parameters.Append .CreateParameter("@pEnderecoCep", adVarChar, adParamInput, 8, obj.ENDERECO.Cep)
    .Parameters.Append .CreateParameter("@pEnderecoNumero", adVarChar, adParamInput, 10, obj.ENDERECO.Numero)
    .Parameters.Append .CreateParameter("@pEnderecoComplemento", adVarChar, adParamInput, 50, obj.ENDERECO.Complemento)
    .Parameters.Append .CreateParameter("@pEnderecoLogradouro", adVarChar, adParamInput, 200, obj.ENDERECO.Logradouro)
    .Parameters.Append .CreateParameter("@pEnderecoBairro", adVarChar, adParamInput, 50, obj.ENDERECO.Bairro)
    .Parameters.Append .CreateParameter("@pEnderecoCidade", adVarChar, adParamInput, 100, obj.ENDERECO.Cidade)
    .Parameters.Append .CreateParameter("@pEnderecoEstado", adVarChar, adParamInput, 2, obj.ENDERECO.Estado)
    
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

Public Function Delete(banco As cDB, obj As cEntidades) As Boolean
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

    .Parameters.Append .CreateParameter("@pCadastroTipo", adVarChar, adParamInput, 2, obj.CadastroTipo)
    .Parameters.Append .CreateParameter("@pCnpjCpf", adVarChar, adParamInput, 14, obj.CnpjCpf)
    .Parameters.Append .CreateParameter("@pIeRg", adVarChar, adParamInput, 16, obj.IeRg)
    .Parameters.Append .CreateParameter("@pNome", adVarChar, adParamInput, 100, obj.Nome)
    .Parameters.Append .CreateParameter("@pNomeFantasia", adVarChar, adParamInput, 50, obj.NomeFantasia)
    
    .Parameters.Append .CreateParameter("@pCadastroStatus", adVarChar, adParamInput, 10, obj.CadastroStatus)
    .Parameters.Append .CreateParameter("@pCadastroPropaganda", adVarChar, adParamInput, 50, obj.CadastroPropaganda)
    .Parameters.Append .CreateParameter("@pCadastroObservacao", adVarChar, adParamInput, 1000, obj.CadastroObservacao)
    
    .Parameters.Append .CreateParameter("@pEnderecoCep", adVarChar, adParamInput, 8, obj.ENDERECO.Cep)
    .Parameters.Append .CreateParameter("@pEnderecoNumero", adVarChar, adParamInput, 10, obj.ENDERECO.Numero)
    .Parameters.Append .CreateParameter("@pEnderecoComplemento", adVarChar, adParamInput, 50, obj.ENDERECO.Complemento)
    .Parameters.Append .CreateParameter("@pEnderecoLogradouro", adVarChar, adParamInput, 200, obj.ENDERECO.Logradouro)
    .Parameters.Append .CreateParameter("@pEnderecoBairro", adVarChar, adParamInput, 50, obj.ENDERECO.Bairro)
    .Parameters.Append .CreateParameter("@pEnderecoCidade", adVarChar, adParamInput, 100, obj.ENDERECO.Cidade)
    .Parameters.Append .CreateParameter("@pEnderecoEstado", adVarChar, adParamInput, 2, obj.ENDERECO.Estado)
    
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

Public Function getEntidades(banco As cDB, qryEntidade As String) As cEntidades
On Error GoTo Err_getEntidades
Dim colObj As New cEntidades
Dim recObj As New cEntidades

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
    Set recObj = New cEntidades
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .CadastroTipo = CStr(IIf(IsNull(rst.Fields("Cadastro_Tipo")), "", rst.Fields("Cadastro_Tipo")))
        .CnpjCpf = CStr(IIf(IsNull(rst.Fields("Cnpj_Cpf")), "", rst.Fields("Cnpj_Cpf")))
        .IeRg = CStr(IIf(IsNull(rst.Fields("Ie_Rg")), "", rst.Fields("Ie_Rg")))
        .Nome = CStr(IIf(IsNull(rst.Fields("Nome")), "", rst.Fields("Nome")))
        .NomeFantasia = CStr(IIf(IsNull(rst.Fields("NomeFantasia")), "", rst.Fields("NomeFantasia")))
        
        .CadastroPropaganda = CStr(IIf(IsNull(rst.Fields("CadastroPropaganda")), "", rst.Fields("CadastroPropaganda")))
        .CadastroObservacao = CStr(IIf(IsNull(rst.Fields("CadastroObservacao")), "", rst.Fields("CadastroObservacao")))
        .CadastroStatus = CStr(IIf(IsNull(rst.Fields("CadastroStatus")), "", rst.Fields("CadastroStatus")))
        
        .ENDERECO.Cep = CStr(IIf(IsNull(rst.Fields("Cep")), "", rst.Fields("Cep")))
        .ENDERECO.Numero = CStr(IIf(IsNull(rst.Fields("Numero")), "", rst.Fields("Numero")))
        .ENDERECO.Complemento = CStr(IIf(IsNull(rst.Fields("Complemento")), "", rst.Fields("Complemento")))
        .ENDERECO.Logradouro = CStr(IIf(IsNull(rst.Fields("Logradouro")), "", rst.Fields("Logradouro")))
        .ENDERECO.Bairro = CStr(IIf(IsNull(rst.Fields("Bairro")), "", rst.Fields("Bairro")))
        .ENDERECO.Cidade = CStr(IIf(IsNull(rst.Fields("Cidade")), "", rst.Fields("Cidade")))
        .ENDERECO.Estado = CStr(IIf(IsNull(rst.Fields("Estado")), "", rst.Fields("Estado")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getEntidades:
    Set getEntidades = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getEntidades:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getEntidades()" & vbNewLine & err.Description)
    Resume Exit_getEntidades

End Function

Public Function getEntidadesObj(banco As cDB, qryEntidade As String, id As String) As cEntidades
On Error GoTo Err_getEntidadesObj
Dim obj As New cEntidades

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

Dim strSQL As String: strSQL = "select * from " & qryEntidade & " where id = " & id & ""

'Saida strSQL, "sql.txt"

With cd
    .ActiveConnection = cnn
    .CommandText = strSQL
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    With obj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .CadastroTipo = CStr(IIf(IsNull(rst.Fields("CadastroTipo")), "", rst.Fields("CadastroTipo")))
        .CnpjCpf = CStr(IIf(IsNull(rst.Fields("Cnpj_Cpf")), "", rst.Fields("Cnpj_Cpf")))
        .IeRg = CStr(IIf(IsNull(rst.Fields("Ie_Rg")), "", rst.Fields("Ie_Rg")))
        .Nome = CStr(IIf(IsNull(rst.Fields("Nome")), "", rst.Fields("Nome")))
        .NomeFantasia = CStr(IIf(IsNull(rst.Fields("Fantasia")), "", rst.Fields("Fantasia")))
        
        .CadastroPropaganda = CStr(IIf(IsNull(rst.Fields("CadastroPropaganda")), "", rst.Fields("CadastroPropaganda")))
        .CadastroObservacao = CStr(IIf(IsNull(rst.Fields("CadastroObservacao")), "", rst.Fields("CadastroObservacao")))
        .CadastroStatus = CStr(IIf(IsNull(rst.Fields("CadastroStatus")), "", rst.Fields("CadastroStatus")))
               
        .ENDERECO.Cep = CStr(IIf(IsNull(rst.Fields("Cep")), "", rst.Fields("Cep")))
        .ENDERECO.Numero = CStr(IIf(IsNull(rst.Fields("Numero")), "", rst.Fields("Numero")))
        .ENDERECO.Complemento = CStr(IIf(IsNull(rst.Fields("Complemento")), "", rst.Fields("Complemento")))
        .ENDERECO.Logradouro = CStr(IIf(IsNull(rst.Fields("Logradouro")), "", rst.Fields("Logradouro")))
        .ENDERECO.Bairro = CStr(IIf(IsNull(rst.Fields("Bairro")), "", rst.Fields("Bairro")))
        .ENDERECO.Cidade = CStr(IIf(IsNull(rst.Fields("Cidade")), "", rst.Fields("Cidade")))
        .ENDERECO.Estado = CStr(IIf(IsNull(rst.Fields("Estado")), "", rst.Fields("Estado")))
        
        .add obj
    End With
    rst.MoveNext
Loop


Exit_getEntidadesObj:
    Set getEntidadesObj = obj
    Set obj = Nothing
    
    Exit Function

Err_getEntidadesObj:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getEntidadesObj()" & vbNewLine & err.Description)
    Resume Exit_getEntidadesObj

End Function

Public Function getEntidadesID(banco As cDB, qryEntidade As String, id As String) As cEntidades
On Error GoTo Err_getEntidadesID
Dim colObj As New cEntidades
Dim recObj As New cEntidades

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

Dim oEndereco As New cEnderecos
'Set oEndereco = New cEnderecos

'Dim rEndereco As cEnderecos
'Set rEndereco = New cEnderecos

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where id = " & id & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cEntidades
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .CadastroTipo = CStr(IIf(IsNull(rst.Fields("CadastroTipo")), "", rst.Fields("CadastroTipo")))
        .CnpjCpf = CStr(IIf(IsNull(rst.Fields("Cnpj_Cpf")), "", rst.Fields("Cnpj_Cpf")))
        .IeRg = CStr(IIf(IsNull(rst.Fields("Ie_Rg")), "", rst.Fields("Ie_Rg")))
        .Nome = CStr(IIf(IsNull(rst.Fields("Nome")), "", rst.Fields("Nome")))
        .NomeFantasia = CStr(IIf(IsNull(rst.Fields("Fantasia")), "", rst.Fields("Fantasia")))
        
        .CadastroPropaganda = CStr(IIf(IsNull(rst.Fields("CadastroPropaganda")), "", rst.Fields("CadastroPropaganda")))
        .CadastroObservacao = CStr(IIf(IsNull(rst.Fields("CadastroObservacao")), "", rst.Fields("CadastroObservacao")))
        .CadastroStatus = CStr(IIf(IsNull(rst.Fields("CadastroStatus")), "", rst.Fields("CadastroStatus")))
        
        Set oEndereco = New cEnderecos
        With oEndereco
            .Cep = CStr(IIf(IsNull(rst.Fields("Cep")), "", rst.Fields("Cep")))
            .Numero = CStr(IIf(IsNull(rst.Fields("Numero")), "", rst.Fields("Numero")))
            .Complemento = CStr(IIf(IsNull(rst.Fields("Complemento")), "", rst.Fields("Complemento")))
            .Logradouro = CStr(IIf(IsNull(rst.Fields("Logradouro")), "", rst.Fields("Logradouro")))
            .Bairro = CStr(IIf(IsNull(rst.Fields("Bairro")), "", rst.Fields("Bairro")))
            .Cidade = CStr(IIf(IsNull(rst.Fields("Cidade")), "", rst.Fields("Cidade")))
            .Estado = CStr(IIf(IsNull(rst.Fields("Estado")), "", rst.Fields("Estado")))
            .add oEndereco
        End With
        
        .ENDERECO.add oEndereco
        
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getEntidadesID:
    Set getEntidadesID = colObj
    Set recObj = Nothing
    Set colObj = Nothing
    Set oEndereco = Nothing

    Exit Function

Err_getEntidadesID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getEntidadesID()" & vbNewLine & err.Description)
    Resume Exit_getEntidadesID

End Function
