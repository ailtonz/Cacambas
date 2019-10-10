VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNotas"
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

Private pTitulo As String
Private pDescricao As String

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

Public Property Get Titulo() As String
    Titulo = pTitulo
End Property

Public Property Let Titulo(ByVal vValor As String)
    pTitulo = vValor
End Property

Public Property Get Descricao() As String
    Descricao = pDescricao
End Property

Public Property Let Descricao(ByVal vValor As String)
    pDescricao = vValor
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

Public Sub add(ByVal rec As cNotas, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As cNotas)
    pCol.add rec, rec.id
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function cNotas(ByVal i As Variant) As cNotas
    Set cNotas = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As cNotas
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As cDB, obj As cNotas) As Boolean
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
    .Parameters.Append .CreateParameter("@pCadastroCategoria", adVarChar, adParamInput, 255, obj.CadastroCategoria)
        
    .Parameters.Append .CreateParameter("@pTitulo", adVarChar, adParamInput, 100, obj.Titulo)
    .Parameters.Append .CreateParameter("@pDescricao", adVarChar, adParamInput, 1000, obj.Descricao)
    

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

Public Function Update(banco As cDB, obj As cNotas) As Boolean
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
    .Parameters.Append .CreateParameter("@pCadastroCategoria", adVarChar, adParamInput, 255, obj.CadastroCategoria)
        
    .Parameters.Append .CreateParameter("@pTitulo", adVarChar, adParamInput, 100, obj.Titulo)
    .Parameters.Append .CreateParameter("@pDescricao", adVarChar, adParamInput, 1000, obj.Descricao)


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

Public Function Delete(banco As cDB, obj As cNotas) As Boolean
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
        
    .Parameters.Append .CreateParameter("@pTitulo", adVarChar, adParamInput, 100, obj.Titulo)
    .Parameters.Append .CreateParameter("@pDescricao", adVarChar, adParamInput, 1000, obj.Descricao)


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

Public Function getNotas(banco As cDB, qryEntidade As String) As cNotas
On Error GoTo Err_getNotas
Dim colObj As New cNotas
Dim recObj As New cNotas

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
    Set recObj = New cNotas
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .Titulo = CStr(IIf(IsNull(rst.Fields("Titulo")), "", rst.Fields("Titulo")))
        .Descricao = CStr(IIf(IsNull(rst.Fields("Descricao")), "", rst.Fields("Descricao")))
                
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getNotas:
    Set getNotas = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getNotas:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getNotas()" & vbNewLine & err.Description)
    Resume Exit_getNotas

End Function

Public Function getNotasID(banco As cDB, qryEntidade As String, i As String) As cNotas
On Error GoTo Err_getNotasID
Dim colObj As New cNotas
Dim recObj As New cNotas

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
    Set recObj = New cNotas
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .Titulo = CStr(IIf(IsNull(rst.Fields("Titulo")), "", rst.Fields("Titulo")))
        .Descricao = CStr(IIf(IsNull(rst.Fields("Descricao")), "", rst.Fields("Descricao")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getNotasID:
    Set getNotasID = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getNotasID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getNotasID()" & vbNewLine & err.Description)
    Resume Exit_getNotasID

End Function

Public Function getNotasFK(banco As cDB, qryEntidade As String, i As String) As cNotas
On Error GoTo Err_getNotasFK
Dim colObj As New cNotas
Dim recObj As New cNotas

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
    Set recObj = New cNotas
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .Titulo = CStr(IIf(IsNull(rst.Fields("Titulo")), "", rst.Fields("Titulo")))
        .Descricao = CStr(IIf(IsNull(rst.Fields("Descricao")), "", rst.Fields("Descricao")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getNotasFK:
    Set getNotasFK = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getNotasFK:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getNotasFK()" & vbNewLine & err.Description)
    Resume Exit_getNotasFK

End Function
