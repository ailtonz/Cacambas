VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMatriz"
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

Private pFrm As String
Private pPrc As String
Private pQry As String
Private pCat As String
Private pTitle As String

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

Public Property Get Frm() As String
    Frm = pFrm
End Property

Public Property Let Frm(ByVal vValor As String)
    pFrm = vValor
End Property

Public Property Get Prc() As String
    Prc = pPrc
End Property

Public Property Let Prc(ByVal vValor As String)
    pPrc = vValor
End Property

Public Property Get qry() As String
    qry = pQry
End Property

Public Property Let qry(ByVal vValor As String)
    pQry = vValor
End Property

Public Property Get Cat() As String
    Cat = pCat
End Property

Public Property Let Cat(ByVal vValor As String)
    pCat = vValor
End Property

Public Property Get Title() As String
    Title = pTitle
End Property

Public Property Let Title(ByVal vValor As String)
    pTitle = vValor
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

Public Sub add(ByVal rec As cMatriz, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As cMatriz)
    pCol.add rec, rec.id
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function cMatriz(ByVal i As Variant) As cMatriz
    Set cMatriz = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As cMatriz
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As cDB, obj As cMatriz) As Boolean
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
        
    .Parameters.Append .CreateParameter("@pFrm", adVarChar, adParamInput, 50, obj.Frm)
    .Parameters.Append .CreateParameter("@pPrc", adVarChar, adParamInput, 50, obj.Prc)
    .Parameters.Append .CreateParameter("@pQry", adVarChar, adParamInput, 50, obj.qry)
    .Parameters.Append .CreateParameter("@pCat", adVarChar, adParamInput, 50, obj.Cat)
    .Parameters.Append .CreateParameter("@pTitle", adVarChar, adParamInput, 50, obj.Title)
    

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

Public Function Update(banco As cDB, obj As cMatriz) As Boolean
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

    .Parameters.Append .CreateParameter("@pFrm", adVarChar, adParamInput, 50, obj.Frm)
    .Parameters.Append .CreateParameter("@pPrc", adVarChar, adParamInput, 50, obj.Prc)
    .Parameters.Append .CreateParameter("@pQry", adVarChar, adParamInput, 50, obj.qry)
    .Parameters.Append .CreateParameter("@pCat", adVarChar, adParamInput, 50, obj.Cat)
    .Parameters.Append .CreateParameter("@pTitle", adVarChar, adParamInput, 50, obj.Title)


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

Public Function Delete(banco As cDB, obj As cMatriz) As Boolean
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

    .Parameters.Append .CreateParameter("@pFrm", adVarChar, adParamInput, 50, obj.Frm)
    .Parameters.Append .CreateParameter("@pPrc", adVarChar, adParamInput, 50, obj.Prc)
    .Parameters.Append .CreateParameter("@pQry", adVarChar, adParamInput, 50, obj.qry)
    .Parameters.Append .CreateParameter("@pCat", adVarChar, adParamInput, 50, obj.Cat)
    .Parameters.Append .CreateParameter("@pTitle", adVarChar, adParamInput, 50, obj.Title)


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

Public Function getMatriz(banco As cDB, qryEntidade As String) As cMatriz
On Error GoTo Err_getMatriz
Dim colObj As New cMatriz
Dim recObj As New cMatriz

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
    Set recObj = New cMatriz
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .Frm = CStr(IIf(IsNull(rst.Fields("Frm")), "", rst.Fields("Frm")))
        .Prc = CStr(IIf(IsNull(rst.Fields("Prc")), "", rst.Fields("Prc")))
        .qry = CStr(IIf(IsNull(rst.Fields("Qry")), "", rst.Fields("Qry")))
        .Cat = CStr(IIf(IsNull(rst.Fields("Cat")), "", rst.Fields("Cat")))
        .Title = CStr(IIf(IsNull(rst.Fields("Title")), "", rst.Fields("Title")))
                
        
                
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getMatriz:
    Set getMatriz = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getMatriz:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getMatriz()" & vbNewLine & err.Description)
    Resume Exit_getMatriz

End Function

Public Function getMatrizFrm(banco As cDB, qryEntidade As String, sFrm As String) As cMatriz
On Error GoTo Err_getMatrizFrm
Dim recObj As New cMatriz

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where frm = '" & sFrm & "'"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .Frm = CStr(IIf(IsNull(rst.Fields("Frm")), "", rst.Fields("Frm")))
        .Prc = CStr(IIf(IsNull(rst.Fields("Prc")), "", rst.Fields("Prc")))
        .qry = CStr(IIf(IsNull(rst.Fields("Qry")), "", rst.Fields("Qry")))
        .Cat = CStr(IIf(IsNull(rst.Fields("Cat")), "", rst.Fields("Cat")))
        .Title = CStr(IIf(IsNull(rst.Fields("Title")), "", rst.Fields("Title")))
        
        .add recObj
    End With
    rst.MoveNext
Loop


Exit_getMatrizFrm:
    Set getMatrizFrm = recObj
    Set recObj = Nothing

    Exit Function

Err_getMatrizFrm:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getMatrizFrm()" & vbNewLine & err.Description)
    Resume Exit_getMatrizFrm

End Function

Public Function getMatrizSubFrm(banco As cDB, qryEntidade As String, sFrm As String, qrySubFrm As String) As cMatriz
On Error GoTo Err_getMatrizSubFrm
Dim recObj As New cMatriz

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

Dim strSQL As String

strSQL = "SELECT * FROM vw_matriz WHERE vw_matriz.frm IN ('" & qrySubFrm & "') " & _
        "AND vw_matriz.FK IN ( SELECT vw_matriz.id FROM vw_matriz   WHERE vw_matriz.FK = '0' and vw_matriz.frm = '" & sFrm & "')"

With cd
    .ActiveConnection = cnn
    .CommandText = strSQL
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .Frm = CStr(IIf(IsNull(rst.Fields("Frm")), "", rst.Fields("Frm")))
        .Prc = CStr(IIf(IsNull(rst.Fields("Prc")), "", rst.Fields("Prc")))
        .qry = CStr(IIf(IsNull(rst.Fields("Qry")), "", rst.Fields("Qry")))
        .Cat = CStr(IIf(IsNull(rst.Fields("Cat")), "", rst.Fields("Cat")))
        .Title = CStr(IIf(IsNull(rst.Fields("Title")), "", rst.Fields("Title")))
        
        .add recObj
    End With
    rst.MoveNext
Loop


Exit_getMatrizSubFrm:
    Set getMatrizSubFrm = recObj
    Set recObj = Nothing

    Exit Function

Err_getMatrizSubFrm:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getMatrizSubFrm()" & vbNewLine & err.Description)
    Resume Exit_getMatrizSubFrm

End Function

'Public Function getMatrizFrm(banco As cDB, qryEntidade As String, sFrm As String) As cMatriz
'On Error GoTo Err_getMatrizFrm
'Dim colObj As New cMatriz
'Dim recObj As New cMatriz
'
'Dim cnn As New ADODB.Connection
'Set cnn = OpenConnectionNEW(banco)
'
'Dim cd As ADODB.Command
'Set cd = New ADODB.Command
'
'Dim rst As ADODB.Recordset
'
'With cd
'    .ActiveConnection = cnn
'    .CommandText = "select * from " & qryEntidade & " where frm = '" & sFrm & "'"
'    .CommandType = adCmdText
'    Set rst = .Execute
'End With
'
'Do While Not rst.EOF
'    Set recObj = New cMatriz
'    With recObj
'
'        .ID = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
'        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
'
'        .Frm = CStr(IIf(IsNull(rst.Fields("Frm")), "", rst.Fields("Frm")))
'        .Prc = CStr(IIf(IsNull(rst.Fields("Prc")), "", rst.Fields("Prc")))
'        .Qry = CStr(IIf(IsNull(rst.Fields("Qry")), "", rst.Fields("Qry")))
'        .Cat = CStr(IIf(IsNull(rst.Fields("Cat")), "", rst.Fields("Cat")))
'        .Title = CStr(IIf(IsNull(rst.Fields("Title")), "", rst.Fields("Title")))
'
'        colObj.all recObj
'    End With
'    rst.MoveNext
'Loop
'
'
'Exit_getMatrizFrm:
'    Set getMatrizFrm = colObj
'    Set recObj = Nothing
'    Set colObj = Nothing
'
'    Exit Function
'
'Err_getMatrizFrm:
'    Call MsgBox("Class: " & TypeName(Me) & " , Function: getMatrizFrm()" & vbNewLine & Err.Description)
'    Resume Exit_getMatrizFrm
'
'End Function


Public Function getMatrizID(banco As cDB, qryEntidade As String, i As String) As cMatriz
On Error GoTo Err_getMatrizID
Dim colObj As New cMatriz
Dim recObj As New cMatriz

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
    Set recObj = New cMatriz
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .Frm = CStr(IIf(IsNull(rst.Fields("Frm")), "", rst.Fields("Frm")))
        .Prc = CStr(IIf(IsNull(rst.Fields("Prc")), "", rst.Fields("Prc")))
        .qry = CStr(IIf(IsNull(rst.Fields("Qry")), "", rst.Fields("Qry")))
        .Cat = CStr(IIf(IsNull(rst.Fields("Cat")), "", rst.Fields("Cat")))
        .Title = CStr(IIf(IsNull(rst.Fields("Title")), "", rst.Fields("Title")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getMatrizID:
    Set getMatrizID = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getMatrizID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getMatrizID()" & vbNewLine & err.Description)
    Resume Exit_getMatrizID

End Function

Public Function getMatrizFK(banco As cDB, qryEntidade As String, i As String) As cMatriz
On Error GoTo Err_getMatrizFK
Dim colObj As New cMatriz
Dim recObj As New cMatriz

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
    Set recObj = New cMatriz
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .Frm = CStr(IIf(IsNull(rst.Fields("Frm")), "", rst.Fields("Frm")))
        .Prc = CStr(IIf(IsNull(rst.Fields("Prc")), "", rst.Fields("Prc")))
        .qry = CStr(IIf(IsNull(rst.Fields("Qry")), "", rst.Fields("Qry")))
        .Cat = CStr(IIf(IsNull(rst.Fields("Cat")), "", rst.Fields("Cat")))
        .Title = CStr(IIf(IsNull(rst.Fields("Title")), "", rst.Fields("Title")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getMatrizFK:
    Set getMatrizFK = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getMatrizFK:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getMatrizFK()" & vbNewLine & err.Description)
    Resume Exit_getMatrizFK

End Function
