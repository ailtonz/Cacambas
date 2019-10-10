VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cContatos"
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

Private pContatoNome As String
Private pContatoTelefone As String
Private pContatoEmail As String
Private pContatoObservacao As String


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

Public Property Get ContatoNome() As String
    ContatoNome = pContatoNome
End Property

Public Property Let ContatoNome(ByVal vValor As String)
    pContatoNome = vValor
End Property

Public Property Get ContatoTelefone() As String
    ContatoTelefone = pContatoTelefone
End Property

Public Property Let ContatoTelefone(ByVal vValor As String)
    pContatoTelefone = vValor
End Property

Public Property Get ContatoEmail() As String
    ContatoEmail = pContatoEmail
End Property

Public Property Let ContatoEmail(ByVal vValor As String)
    pContatoEmail = vValor
End Property

Public Property Get ContatoObservacao() As String
    ContatoObservacao = pContatoObservacao
End Property

Public Property Let ContatoObservacao(ByVal vValor As String)
    pContatoObservacao = vValor
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

Public Sub add(ByVal rec As cContatos, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As cContatos)
    pCol.add rec, rec.id
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function cContatos(ByVal i As Variant) As cContatos
    Set cContatos = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As cContatos
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As cDB, obj As cContatos) As Boolean
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
    
    .Parameters.Append .CreateParameter("@pContatoNome", adVarChar, adParamInput, 50, obj.ContatoNome)
    .Parameters.Append .CreateParameter("@pContatoTelefone", adVarChar, adParamInput, 50, obj.ContatoTelefone)
    .Parameters.Append .CreateParameter("@pContatoEmail", adVarChar, adParamInput, 120, obj.ContatoEmail)
    

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

Public Function Update(banco As cDB, obj As cContatos) As Boolean
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
    .Parameters.Append .CreateParameter("@pCadastroCategoria", adVarChar, adParamInput, 10, obj.CadastroCategoria)
    
    .Parameters.Append .CreateParameter("@pContatoNome", adVarChar, adParamInput, 50, obj.ContatoNome)
    .Parameters.Append .CreateParameter("@pContatoTelefone", adVarChar, adParamInput, 50, obj.ContatoTelefone)
    .Parameters.Append .CreateParameter("@pContatoEmail", adVarChar, adParamInput, 120, obj.ContatoEmail)


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

Public Function Delete(banco As cDB, obj As cContatos) As Boolean
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
    .Parameters.Append .CreateParameter("@pCadastroCategoria", adVarChar, adParamInput, 255, obj.CadastroCategoria)
    
    .Parameters.Append .CreateParameter("@pContatoNome", adVarChar, adParamInput, 50, obj.ContatoNome)
    .Parameters.Append .CreateParameter("@pContatoTelefone", adVarChar, adParamInput, 50, obj.ContatoTelefone)
    .Parameters.Append .CreateParameter("@pContatoEmail", adVarChar, adParamInput, 120, obj.ContatoEmail)


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



Public Function getContatos(banco As cDB, qryEntidade As String) As cContatos
On Error GoTo Err_getContatos
Dim colObj As New cContatos
Dim recObj As New cContatos

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
    Set recObj = New cContatos
        With recObj
           
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .ContatoNome = CStr(IIf(IsNull(rst.Fields("ContatoNome")), "", rst.Fields("ContatoNome")))
        .ContatoTelefone = CStr(IIf(IsNull(rst.Fields("ContatoTelefone")), "", rst.Fields("ContatoTelefone")))
        .ContatoEmail = CStr(IIf(IsNull(rst.Fields("ContatoEmail")), "", rst.Fields("ContatoEmail")))
'        .ContatoObservacao = CStr(IIf(IsNull(rst.Fields("ContatoObservacao")), "", rst.Fields("ContatoObservacao")))
            
        colObj.all recObj
        End With
    rst.MoveNext
Loop


Exit_getContatos:
    Set getContatos = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getContatos:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getContatos()" & vbNewLine & err.Description)
    Resume Exit_getContatos

End Function

Public Function getContatosID(banco As cDB, qryEntidade As String, i As String) As cContatos
On Error GoTo Err_getContatosID
Dim colObj As New cContatos
Dim recObj As New cContatos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where ID = " & i & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cContatos
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .ContatoNome = CStr(IIf(IsNull(rst.Fields("ContatoNome")), "", rst.Fields("ContatoNome")))
        .ContatoTelefone = CStr(IIf(IsNull(rst.Fields("ContatoTelefone")), "", rst.Fields("ContatoTelefone")))
        .ContatoEmail = CStr(IIf(IsNull(rst.Fields("ContatoEmail")), "", rst.Fields("ContatoEmail")))
        .ContatoObservacao = CStr(IIf(IsNull(rst.Fields("ContatoObservacao")), "", rst.Fields("ContatoObservacao")))
                
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getContatosID:
    Set getContatosID = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getContatosID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getContatosID()" & vbNewLine & err.Description)
    Resume Exit_getContatosID

End Function

Public Function getContatosFK(banco As cDB, qryEntidade As String, i As String) As cContatos
On Error GoTo Err_getContatosFK
Dim colObj As New cContatos
Dim recObj As New cContatos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where FK = " & i & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cContatos
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
               
        .ContatoNome = CStr(IIf(IsNull(rst.Fields("ContatoNome")), "", rst.Fields("ContatoNome")))
        .ContatoTelefone = CStr(IIf(IsNull(rst.Fields("ContatoTelefone")), "", rst.Fields("ContatoTelefone")))
        .ContatoEmail = CStr(IIf(IsNull(rst.Fields("ContatoEmail")), "", rst.Fields("ContatoEmail")))
        .ContatoObservacao = CStr(IIf(IsNull(rst.Fields("ContatoObservacao")), "", rst.Fields("ContatoObservacao")))
                
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getContatosFK:
    Set getContatosFK = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getContatosFK:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getContatosFK()" & vbNewLine & err.Description)
    Resume Exit_getContatosFK

End Function
