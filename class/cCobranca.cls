VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCobranca"
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

Private pCobrancaTipo As String
Private pCobrancaInscricao As String
Private pCobrancaSacado As String
Private pCobrancaCep As String
Private pCobrancaLogradouro As String
Private pCobrancaBairro As String
Private pCobrancaCidade As String
Private pCobrancaEstado As String

Private pCobrancaContato As String
Private pCobrancaTelefone As String
Private pCobrancaEmail As String


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

Public Property Get CobrancaTipo() As String
    CobrancaTipo = pCobrancaTipo
End Property

Public Property Let CobrancaTipo(ByVal vValor As String)
    pCobrancaTipo = vValor
End Property

Public Property Get CobrancaInscricao() As String
    CobrancaInscricao = pCobrancaInscricao
End Property

Public Property Let CobrancaInscricao(ByVal vValor As String)
    pCobrancaInscricao = vValor
End Property

Public Property Get CobrancaSacado() As String
    CobrancaSacado = pCobrancaSacado
End Property

Public Property Let CobrancaSacado(ByVal vValor As String)
    pCobrancaSacado = vValor
End Property

Public Property Get CobrancaCep() As String
    CobrancaCep = pCobrancaCep
End Property

Public Property Let CobrancaCep(ByVal vValor As String)
    pCobrancaCep = vValor
End Property

Public Property Get CobrancaLogradouro() As String
    CobrancaLogradouro = pCobrancaLogradouro
End Property

Public Property Let CobrancaLogradouro(ByVal vValor As String)
    pCobrancaLogradouro = vValor
End Property

Public Property Get CobrancaBairro() As String
    CobrancaBairro = pCobrancaBairro
End Property

Public Property Let CobrancaBairro(ByVal vValor As String)
    pCobrancaBairro = vValor
End Property

Public Property Get CobrancaCidade() As String
    CobrancaCidade = pCobrancaCidade
End Property

Public Property Let CobrancaCidade(ByVal vValor As String)
    pCobrancaCidade = vValor
End Property

Public Property Get CobrancaEstado() As String
    CobrancaEstado = pCobrancaEstado
End Property

Public Property Let CobrancaEstado(ByVal vValor As String)
    pCobrancaEstado = vValor
End Property

Public Property Get CobrancaContato() As String
    CobrancaContato = pCobrancaContato
End Property

Public Property Let CobrancaContato(ByVal vValor As String)
    pCobrancaContato = vValor
End Property

Public Property Get CobrancaTelefone() As String
    CobrancaTelefone = pCobrancaTelefone
End Property

Public Property Let CobrancaTelefone(ByVal vValor As String)
    pCobrancaTelefone = vValor
End Property

Public Property Get CobrancaEmail() As String
    CobrancaEmail = pCobrancaEmail
End Property

Public Property Let CobrancaEmail(ByVal vValor As String)
    pCobrancaEmail = vValor
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

Public Sub add(ByVal rec As cCobranca, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As cCobranca)
    pCol.add rec, rec.id
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function cCobranca(ByVal i As Variant) As cCobranca
    Set cCobranca = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As cCobranca
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As cDB, obj As cCobranca) As Boolean
''' ATENÇÃO: USAR APENAS QUANDO HOUVER + DE 1 REGISTRO (FK)

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
    
    .Parameters.Append .CreateParameter("@pCobrancaTipo", adVarChar, adParamInput, 14, obj.CobrancaTipo)
    .Parameters.Append .CreateParameter("@pCobrancaInscricao", adVarChar, adParamInput, 14, obj.CobrancaInscricao)
    .Parameters.Append .CreateParameter("@pCobrancaSacado", adVarChar, adParamInput, 30, obj.CobrancaSacado)

    .Parameters.Append .CreateParameter("@pCobrancaCep", adVarChar, adParamInput, 8, obj.CobrancaCep)
    .Parameters.Append .CreateParameter("@pCobrancaLogradouro", adVarChar, adParamInput, 40, obj.CobrancaLogradouro)
    .Parameters.Append .CreateParameter("@pCobrancaBairro", adVarChar, adParamInput, 12, obj.CobrancaBairro)
    .Parameters.Append .CreateParameter("@pCobrancaCidade", adVarChar, adParamInput, 15, obj.CobrancaCidade)
    .Parameters.Append .CreateParameter("@pCobrancaEstado", adVarChar, adParamInput, 2, obj.CobrancaEstado)

    .Parameters.Append .CreateParameter("@pCobrancaContato", adVarChar, adParamInput, 15, obj.CobrancaContato)
    .Parameters.Append .CreateParameter("@pCobrancaTelefone", adVarChar, adParamInput, 50, obj.CobrancaTelefone)
    .Parameters.Append .CreateParameter("@pCobrancaEmail", adVarChar, adParamInput, 120, obj.CobrancaEmail)


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

Public Function Update(banco As cDB, obj As cCobranca) As Boolean
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
    
    .Parameters.Append .CreateParameter("@pCobrancaTipo", adVarChar, adParamInput, 14, obj.CobrancaTipo)
    .Parameters.Append .CreateParameter("@pCobrancaInscricao", adVarChar, adParamInput, 14, obj.CobrancaInscricao)
    .Parameters.Append .CreateParameter("@pCobrancaSacado", adVarChar, adParamInput, 30, obj.CobrancaSacado)

    .Parameters.Append .CreateParameter("@pCobrancaCep", adVarChar, adParamInput, 8, obj.CobrancaCep)
    .Parameters.Append .CreateParameter("@pCobrancaLogradouro", adVarChar, adParamInput, 40, obj.CobrancaLogradouro)
    .Parameters.Append .CreateParameter("@pCobrancaBairro", adVarChar, adParamInput, 12, obj.CobrancaBairro)
    .Parameters.Append .CreateParameter("@pCobrancaCidade", adVarChar, adParamInput, 15, obj.CobrancaCidade)
    .Parameters.Append .CreateParameter("@pCobrancaEstado", adVarChar, adParamInput, 2, obj.CobrancaEstado)

    .Parameters.Append .CreateParameter("@pCobrancaContato", adVarChar, adParamInput, 15, obj.CobrancaContato)
    .Parameters.Append .CreateParameter("@pCobrancaTelefone", adVarChar, adParamInput, 50, obj.CobrancaTelefone)
    .Parameters.Append .CreateParameter("@pCobrancaEmail", adVarChar, adParamInput, 120, obj.CobrancaEmail)


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

Public Function Delete(banco As cDB, obj As cCobranca) As Boolean
''' ATENÇÃO: USAR APENAS QUANDO HOUVER + DE 1 REGISTRO (FK)

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

    .Parameters.Append .CreateParameter("@pCobrancaTipo", adVarChar, adParamInput, 14, obj.CobrancaTipo)
    .Parameters.Append .CreateParameter("@pCobrancaInscricao", adVarChar, adParamInput, 14, obj.CobrancaInscricao)
    .Parameters.Append .CreateParameter("@pCobrancaSacado", adVarChar, adParamInput, 30, obj.CobrancaSacado)

    .Parameters.Append .CreateParameter("@pCobrancaCep", adVarChar, adParamInput, 8, obj.CobrancaCep)
    .Parameters.Append .CreateParameter("@pCobrancaLogradouro", adVarChar, adParamInput, 40, obj.CobrancaLogradouro)
    .Parameters.Append .CreateParameter("@pCobrancaBairro", adVarChar, adParamInput, 12, obj.CobrancaBairro)
    .Parameters.Append .CreateParameter("@pCobrancaCidade", adVarChar, adParamInput, 15, obj.CobrancaCidade)
    .Parameters.Append .CreateParameter("@pCobrancaEstado", adVarChar, adParamInput, 2, obj.CobrancaEstado)

    .Parameters.Append .CreateParameter("@pCobrancaContato", adVarChar, adParamInput, 15, obj.CobrancaContato)
    .Parameters.Append .CreateParameter("@pCobrancaTelefone", adVarChar, adParamInput, 50, obj.CobrancaTelefone)
    .Parameters.Append .CreateParameter("@pCobrancaEmail", adVarChar, adParamInput, 120, obj.CobrancaEmail)


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

Public Function getCobranca(banco As cDB, qryEntidade As String) As cCobranca
On Error GoTo Err_getCobranca
Dim colObj As New cCobranca
Dim recObj As New cCobranca

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
    Set recObj = New cCobranca
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
                              
        .CobrancaTipo = CStr(IIf(IsNull(rst.Fields("CobrancaTipo")), "", rst.Fields("CobrancaTipo")))
        .CobrancaInscricao = CStr(IIf(IsNull(rst.Fields("CobrancaInscricao")), "", rst.Fields("CobrancaInscricao")))
        .CobrancaSacado = CStr(IIf(IsNull(rst.Fields("CobrancaSacado")), "", rst.Fields("CobrancaSacado")))
               
        .CobrancaCep = CStr(IIf(IsNull(rst.Fields("CobrancaCep")), "", rst.Fields("CobrancaCep")))
        .CobrancaLogradouro = CStr(IIf(IsNull(rst.Fields("CobrancaLogradouro")), "", rst.Fields("CobrancaLogradouro")))
        .CobrancaBairro = CStr(IIf(IsNull(rst.Fields("CobrancaBairro")), "", rst.Fields("CobrancaBairro")))
        .CobrancaCidade = CStr(IIf(IsNull(rst.Fields("CobrancaCidade")), "", rst.Fields("CobrancaCidade")))
        .CobrancaEstado = CStr(IIf(IsNull(rst.Fields("CobrancaEstado")), "", rst.Fields("CobrancaEstado")))
        
        .CobrancaContato = CStr(IIf(IsNull(rst.Fields("CobrancaContato")), "", rst.Fields("CobrancaContato")))
        .CobrancaTelefone = CStr(IIf(IsNull(rst.Fields("CobrancaTelefone")), "", rst.Fields("CobrancaTelefone")))
        .CobrancaEmail = CStr(IIf(IsNull(rst.Fields("CobrancaEmail")), "", rst.Fields("CobrancaEmail")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getCobranca:
    Set getCobranca = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getCobranca:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getCobranca()" & vbNewLine & err.Description)
    Resume Exit_getCobranca

End Function

Public Function getCobrancaObj(banco As cDB, qry As String, id As String) As cCobranca
On Error GoTo Err_getCobrancaObj
Dim obj As New cCobranca

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
               
        .CobrancaCep = CStr(IIf(IsNull(rst.Fields("CobrancaCep")), "", rst.Fields("CobrancaCep")))
        .CobrancaLogradouro = CStr(IIf(IsNull(rst.Fields("CobrancaLogradouro")), "", rst.Fields("CobrancaLogradouro")))
        .CobrancaBairro = CStr(IIf(IsNull(rst.Fields("CobrancaBairro")), "", rst.Fields("CobrancaBairro")))
        .CobrancaCidade = CStr(IIf(IsNull(rst.Fields("CobrancaCidade")), "", rst.Fields("CobrancaCidade")))
        .CobrancaEstado = CStr(IIf(IsNull(rst.Fields("CobrancaEstado")), "", rst.Fields("CobrancaEstado")))
        
        .add obj
    End With
    rst.MoveNext
Loop


Exit_getCobrancaObj:
    Set getCobrancaObj = obj
    Set obj = Nothing
    
    Exit Function

Err_getCobrancaObj:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getCobrancaObj()" & vbNewLine & err.Description)
    Resume Exit_getCobrancaObj

End Function

Public Function getCobrancaID(banco As cDB, qryEntidade As String, id As String) As cCobranca
On Error GoTo Err_getCobrancaID
Dim colObj As New cCobranca
Dim recObj As New cCobranca

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where ID = " & id & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cCobranca
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .CobrancaTipo = CStr(IIf(IsNull(rst.Fields("CobrancaTipo")), "", rst.Fields("CobrancaTipo")))
        .CobrancaInscricao = CStr(IIf(IsNull(rst.Fields("CobrancaInscricao")), "", rst.Fields("CobrancaInscricao")))
        .CobrancaSacado = CStr(IIf(IsNull(rst.Fields("CobrancaSacado")), "", rst.Fields("CobrancaSacado")))
               
        .CobrancaCep = CStr(IIf(IsNull(rst.Fields("CobrancaCep")), "", rst.Fields("CobrancaCep")))
        .CobrancaLogradouro = CStr(IIf(IsNull(rst.Fields("CobrancaLogradouro")), "", rst.Fields("CobrancaLogradouro")))
        .CobrancaBairro = CStr(IIf(IsNull(rst.Fields("CobrancaBairro")), "", rst.Fields("CobrancaBairro")))
        .CobrancaCidade = CStr(IIf(IsNull(rst.Fields("CobrancaCidade")), "", rst.Fields("CobrancaCidade")))
        .CobrancaEstado = CStr(IIf(IsNull(rst.Fields("CobrancaEstado")), "", rst.Fields("CobrancaEstado")))
        
        .CobrancaContato = CStr(IIf(IsNull(rst.Fields("CobrancaContato")), "", rst.Fields("CobrancaContato")))
        .CobrancaTelefone = CStr(IIf(IsNull(rst.Fields("CobrancaTelefone")), "", rst.Fields("CobrancaTelefone")))
        .CobrancaEmail = CStr(IIf(IsNull(rst.Fields("CobrancaEmail")), "", rst.Fields("CobrancaEmail")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getCobrancaID:
    Set getCobrancaID = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getCobrancaID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getCobrancaID()" & vbNewLine & err.Description)
    Resume Exit_getCobrancaID

End Function

Public Function getCobrancaFK(banco As cDB, qryEntidade As String, id As String) As cCobranca
On Error GoTo Err_getCobrancaFK
Dim colObj As New cCobranca
Dim recObj As New cCobranca

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where FK = " & id & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cCobranca
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .CobrancaTipo = CStr(IIf(IsNull(rst.Fields("CobrancaTipo")), "", rst.Fields("CobrancaTipo")))
        .CobrancaInscricao = CStr(IIf(IsNull(rst.Fields("CobrancaInscricao")), "", rst.Fields("CobrancaInscricao")))
        .CobrancaSacado = CStr(IIf(IsNull(rst.Fields("CobrancaSacado")), "", rst.Fields("CobrancaSacado")))
               
        .CobrancaCep = CStr(IIf(IsNull(rst.Fields("CobrancaCep")), "", rst.Fields("CobrancaCep")))
        .CobrancaLogradouro = CStr(IIf(IsNull(rst.Fields("CobrancaLogradouro")), "", rst.Fields("CobrancaLogradouro")))
        .CobrancaBairro = CStr(IIf(IsNull(rst.Fields("CobrancaBairro")), "", rst.Fields("CobrancaBairro")))
        .CobrancaCidade = CStr(IIf(IsNull(rst.Fields("CobrancaCidade")), "", rst.Fields("CobrancaCidade")))
        .CobrancaEstado = CStr(IIf(IsNull(rst.Fields("CobrancaEstado")), "", rst.Fields("CobrancaEstado")))
        
        .CobrancaContato = CStr(IIf(IsNull(rst.Fields("CobrancaContato")), "", rst.Fields("CobrancaContato")))
        .CobrancaTelefone = CStr(IIf(IsNull(rst.Fields("CobrancaTelefone")), "", rst.Fields("CobrancaTelefone")))
        .CobrancaEmail = CStr(IIf(IsNull(rst.Fields("CobrancaEmail")), "", rst.Fields("CobrancaEmail")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getCobrancaFK:
    Set getCobrancaFK = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getCobrancaFK:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getCobrancaFK()" & vbNewLine & err.Description)
    Resume Exit_getCobrancaFK

End Function

