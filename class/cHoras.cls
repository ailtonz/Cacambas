VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHoras"
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

Private pCliente As String
Private pTarefa As String
Private pControle As String
Private pDataTarefa As String
Private pSalarioHora As String
Private pCargaHoraria As String
Private pTurno01_Entrada As String
Private pTurno01_Saida As String
Private pTurno02_Entrada As String
Private pTurno02_Saida As String

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

Public Property Get Cliente() As String
    Cliente = pCliente
End Property

Public Property Let Cliente(ByVal vValor As String)
    pCliente = vValor
End Property

Public Property Get Tarefa() As String
    Tarefa = pTarefa
End Property

Public Property Let Tarefa(ByVal vValor As String)
    pTarefa = vValor
End Property

Public Property Get Controle() As String
    Controle = pControle
End Property

Public Property Let Controle(ByVal vValor As String)
    pControle = vValor
End Property

Public Property Get DataTarefa() As String
    DataTarefa = pDataTarefa
End Property

Public Property Let DataTarefa(ByVal vValor As String)
    pDataTarefa = vValor
End Property

Public Property Get SalarioHora() As String
    SalarioHora = pSalarioHora
End Property

Public Property Let SalarioHora(ByVal vValor As String)
    pSalarioHora = vValor
End Property

Public Property Get CargaHoraria() As String
    CargaHoraria = pCargaHoraria
End Property

Public Property Let CargaHoraria(ByVal vValor As String)
    pCargaHoraria = vValor
End Property

Public Property Get Turno01_Entrada() As String
    Turno01_Entrada = pTurno01_Entrada
End Property

Public Property Let Turno01_Entrada(ByVal vValor As String)
    pTurno01_Entrada = vValor
End Property

Public Property Get Turno01_Saida() As String
    Turno01_Saida = pTurno01_Saida
End Property

Public Property Let Turno01_Saida(ByVal vValor As String)
    pTurno01_Saida = vValor
End Property

Public Property Get Turno02_Entrada() As String
    Turno02_Entrada = pTurno02_Entrada
End Property

Public Property Let Turno02_Entrada(ByVal vValor As String)
    pTurno02_Entrada = vValor
End Property

Public Property Get Turno02_Saida() As String
    Turno02_Saida = pTurno02_Saida
End Property

Public Property Let Turno02_Saida(ByVal vValor As String)
    pTurno02_Saida = vValor
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

Public Sub add(ByVal rec As cHoras, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As cHoras)
    pCol.add rec, rec.id
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function cHoras(ByVal i As Variant) As cHoras
    Set cHoras = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As cHoras
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As cDB, obj As cHoras) As Boolean
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
    
    .Parameters.Append .CreateParameter("@pCliente", adVarChar, adParamInput, 50, obj.Cliente)
    .Parameters.Append .CreateParameter("@pTarefa", adVarChar, adParamInput, 50, obj.Tarefa)
    .Parameters.Append .CreateParameter("@pControle", adVarChar, adParamInput, 100, obj.Controle)
    .Parameters.Append .CreateParameter("@pDataTarefa", adVarChar, adParamInput, 10, obj.DataTarefa)
    .Parameters.Append .CreateParameter("@pSalarioHora", adVarChar, adParamInput, 50, obj.SalarioHora)
    .Parameters.Append .CreateParameter("@pCargaHoraria", adVarChar, adParamInput, 10, obj.CargaHoraria)
    .Parameters.Append .CreateParameter("@pTurno01_Entrada", adVarChar, adParamInput, 50, obj.Turno01_Entrada)
    .Parameters.Append .CreateParameter("@pTurno01_Saida", adVarChar, adParamInput, 50, obj.Turno01_Saida)
    .Parameters.Append .CreateParameter("@pTurno02_Entrada", adVarChar, adParamInput, 50, obj.Turno02_Entrada)
    .Parameters.Append .CreateParameter("@pTurno02_Saida", adVarChar, adParamInput, 50, obj.Turno02_Saida)

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

Public Function Update(banco As cDB, obj As cHoras) As Boolean
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
    
    .Parameters.Append .CreateParameter("@pCliente", adVarChar, adParamInput, 50, obj.Cliente)
    .Parameters.Append .CreateParameter("@pTarefa", adVarChar, adParamInput, 50, IIf(obj.Tarefa <> "", obj.Tarefa, "OK"))
    .Parameters.Append .CreateParameter("@pControle", adVarChar, adParamInput, 100, obj.Controle)
    .Parameters.Append .CreateParameter("@pDataTarefa", adVarChar, adParamInput, 10, obj.DataTarefa)
    .Parameters.Append .CreateParameter("@pSalarioHora", adVarChar, adParamInput, 50, obj.SalarioHora)
    .Parameters.Append .CreateParameter("@pCargaHoraria", adVarChar, adParamInput, 10, obj.CargaHoraria)
    .Parameters.Append .CreateParameter("@pTurno01_Entrada", adVarChar, adParamInput, 50, obj.Turno01_Entrada)
    .Parameters.Append .CreateParameter("@pTurno01_Saida", adVarChar, adParamInput, 50, obj.Turno01_Saida)
    .Parameters.Append .CreateParameter("@pTurno02_Entrada", adVarChar, adParamInput, 50, obj.Turno02_Entrada)
    .Parameters.Append .CreateParameter("@pTurno02_Saida", adVarChar, adParamInput, 50, obj.Turno02_Saida)



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

Public Function Delete(banco As cDB, obj As cHoras) As Boolean
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
    
    .Parameters.Append .CreateParameter("@pCliente", adVarChar, adParamInput, 10, obj.Cliente)
    .Parameters.Append .CreateParameter("@pTarefa", adVarChar, adParamInput, 50, obj.Tarefa)
    .Parameters.Append .CreateParameter("@pControle", adVarChar, adParamInput, 100, obj.Controle)
    .Parameters.Append .CreateParameter("@pDataTarefa", adVarChar, adParamInput, 10, obj.DataTarefa)
    .Parameters.Append .CreateParameter("@pSalarioHora", adVarChar, adParamInput, 50, obj.SalarioHora)
    .Parameters.Append .CreateParameter("@pCargaHoraria", adVarChar, adParamInput, 10, obj.CargaHoraria)
    .Parameters.Append .CreateParameter("@pTurno01_Entrada", adVarChar, adParamInput, 50, obj.Turno01_Entrada)
    .Parameters.Append .CreateParameter("@pTurno01_Saida", adVarChar, adParamInput, 50, obj.Turno01_Saida)
    .Parameters.Append .CreateParameter("@pTurno02_Entrada", adVarChar, adParamInput, 50, obj.Turno02_Entrada)
    .Parameters.Append .CreateParameter("@pTurno02_Saida", adVarChar, adParamInput, 50, obj.Turno02_Saida)



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

Public Function getHoras(banco As cDB, qryEntidade As String) As cHoras
On Error GoTo Err_getHoras
Dim colObj As New cHoras
Dim recObj As New cHoras

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
    Set recObj = New cHoras
    With recObj

        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))

        .Cliente = CStr(rst.Fields("Cliente"))
        .Tarefa = CStr(IIf(IsNull(rst.Fields("Tarefa")), "", rst.Fields("Tarefa")))
        .Controle = CStr(IIf(IsNull(rst.Fields("Controle")), "", rst.Fields("Controle")))
        .DataTarefa = CStr(rst.Fields("DataTarefa"))
        .SalarioHora = CStr(rst.Fields("SalarioHora"))
        .CargaHoraria = CStr(rst.Fields("CargaHoraria"))
        .Turno01_Entrada = CStr(rst.Fields("Turno01_Entrada"))
        .Turno01_Saida = CStr(IIf(IsNull(rst.Fields("Turno01_Saida")), "", rst.Fields("Turno01_Saida")))
        .Turno02_Entrada = CStr(IIf(IsNull(rst.Fields("Turno02_Entrada")), "", rst.Fields("Turno02_Entrada")))
        .Turno02_Saida = CStr(IIf(IsNull(rst.Fields("Turno02_Saida")), "", rst.Fields("Turno02_Saida")))

        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getHoras:
    Set getHoras = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getHoras:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getHoras()" & vbNewLine & err.Description)
    Resume Exit_getHoras

End Function


Public Function getHoraID(banco As cDB, qryEntidade As String, i As String) As cHoras
On Error GoTo Err_getHoraID
Dim colObj As New cHoras
Dim recObj As New cHoras

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
    Set recObj = New cHoras
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))

        .Cliente = CStr(rst.Fields("Cliente"))
        .Tarefa = CStr(IIf(IsNull(rst.Fields("Tarefa")), "", rst.Fields("Tarefa")))
        .Controle = CStr(IIf(IsNull(rst.Fields("Controle")), "", rst.Fields("Controle")))
        .DataTarefa = CStr(rst.Fields("DataTarefa"))
        .SalarioHora = CStr(rst.Fields("SalarioHora"))
        .CargaHoraria = CStr(rst.Fields("CargaHoraria"))
        .Turno01_Entrada = CStr(rst.Fields("Turno01_Entrada"))
        .Turno01_Saida = CStr(IIf(IsNull(rst.Fields("Turno01_Saida")), "", rst.Fields("Turno01_Saida")))
        .Turno02_Entrada = CStr(IIf(IsNull(rst.Fields("Turno02_Entrada")), "", rst.Fields("Turno02_Entrada")))
        .Turno02_Saida = CStr(IIf(IsNull(rst.Fields("Turno02_Saida")), "", rst.Fields("Turno02_Saida")))
                
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getHoraID:
    Set getHoraID = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getHoraID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getHoraID()" & vbNewLine & err.Description)
    Resume Exit_getHoraID

End Function

Public Function getHorasPorFuncionarioPeriodo(banco As cDB, qryEntidade As String, Funcionario As String, Periodo As String) As cHoras
On Error GoTo Err_getHorasPorFuncionarioPeriodo
Dim colObj As New cHoras
Dim recObj As New cHoras

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

'' select * from vw_horas WHERE vw_horas.controle = 'fev_15' and vw_horas.Tarefa = 'ailton';

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where controle = '" & Periodo & "' and Cliente = '" & Funcionario & "' order by DataTarefa"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cHoras
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))

        .Cliente = CStr(rst.Fields("Cliente"))
        .Tarefa = CStr(IIf(IsNull(rst.Fields("Tarefa")), "", rst.Fields("Tarefa")))
        .Controle = CStr(IIf(IsNull(rst.Fields("Controle")), "", rst.Fields("Controle")))
        .DataTarefa = CStr(rst.Fields("DataTarefa"))
        .SalarioHora = CStr(rst.Fields("SalarioHora"))
        .CargaHoraria = CStr(rst.Fields("CargaHoraria"))
        .Turno01_Entrada = CStr(rst.Fields("Turno01_Entrada"))
        .Turno01_Saida = CStr(IIf(IsNull(rst.Fields("Turno01_Saida")), "", rst.Fields("Turno01_Saida")))
        .Turno02_Entrada = CStr(IIf(IsNull(rst.Fields("Turno02_Entrada")), "", rst.Fields("Turno02_Entrada")))
        .Turno02_Saida = CStr(IIf(IsNull(rst.Fields("Turno02_Saida")), "", rst.Fields("Turno02_Saida")))
                
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getHorasPorFuncionarioPeriodo:
    Set getHorasPorFuncionarioPeriodo = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getHorasPorFuncionarioPeriodo:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getHorasPorFuncionarioPeriodo()" & vbNewLine & err.Description)
    Resume Exit_getHorasPorFuncionarioPeriodo

End Function

Public Function getHorasPorFuncionarioInicioTerminio(banco As cDB, qryEntidade As String, Funcionario As String, Inicio As String, Terminio As String) As cHoras
On Error GoTo Err_getHorasPorFuncionarioInicioTerminio
Dim colObj As New cHoras
Dim recObj As New cHoras

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

'' select * from vw_horas WHERE vw_horas.controle = 'fev_15' and vw_horas.Tarefa = 'ailton';

With cd
    .ActiveConnection = cnn
    .CommandText = "SELECT * FROM " & qryEntidade & " WHERE tarefa = '" & Funcionario & "' AND DataTarefa BETWEEN ('" & Inicio & "') AND ('" & Terminio & "') order by DataTarefa"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cHoras
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))

        .Cliente = CStr(rst.Fields("Cliente"))
        .Tarefa = CStr(IIf(IsNull(rst.Fields("Tarefa")), "", rst.Fields("Tarefa")))
        .Controle = CStr(IIf(IsNull(rst.Fields("Controle")), "", rst.Fields("Controle")))
        .DataTarefa = CStr(rst.Fields("DataTarefa"))
        .SalarioHora = CStr(rst.Fields("SalarioHora"))
        .CargaHoraria = CStr(rst.Fields("CargaHoraria"))
        .Turno01_Entrada = CStr(rst.Fields("Turno01_Entrada"))
        .Turno01_Saida = CStr(IIf(IsNull(rst.Fields("Turno01_Saida")), "", rst.Fields("Turno01_Saida")))
        .Turno02_Entrada = CStr(IIf(IsNull(rst.Fields("Turno02_Entrada")), "", rst.Fields("Turno02_Entrada")))
        .Turno02_Saida = CStr(IIf(IsNull(rst.Fields("Turno02_Saida")), "", rst.Fields("Turno02_Saida")))
                
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getHorasPorFuncionarioInicioTerminio:
    Set getHorasPorFuncionarioInicioTerminio = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getHorasPorFuncionarioInicioTerminio:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getHorasPorFuncionarioInicioTerminio()" & vbNewLine & err.Description)
    Resume Exit_getHorasPorFuncionarioInicioTerminio

End Function

Public Function getHorasPorFuncionario(banco As cDB, qryEntidade As String, Funcionario As String) As cHoras
On Error GoTo Err_getHorasPorFuncionario
Dim colObj As New cHoras
Dim recObj As New cHoras

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

'' select * from vw_horas WHERE vw_horas.controle = 'fev_15' and vw_horas.Tarefa = 'ailton';

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where Cliente = '" & Funcionario & "' order by ID"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cHoras
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))

        .Cliente = CStr(rst.Fields("Cliente"))
        .Tarefa = CStr(IIf(IsNull(rst.Fields("Tarefa")), "", rst.Fields("Tarefa")))
        .Controle = CStr(IIf(IsNull(rst.Fields("Controle")), "", rst.Fields("Controle")))
        .DataTarefa = CStr(rst.Fields("DataTarefa"))
        .SalarioHora = CStr(rst.Fields("SalarioHora"))
        .CargaHoraria = CStr(rst.Fields("CargaHoraria"))
        .Turno01_Entrada = CStr(rst.Fields("Turno01_Entrada"))
        .Turno01_Saida = CStr(IIf(IsNull(rst.Fields("Turno01_Saida")), "", rst.Fields("Turno01_Saida")))
        .Turno02_Entrada = CStr(IIf(IsNull(rst.Fields("Turno02_Entrada")), "", rst.Fields("Turno02_Entrada")))
        .Turno02_Saida = CStr(IIf(IsNull(rst.Fields("Turno02_Saida")), "", rst.Fields("Turno02_Saida")))
                
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getHorasPorFuncionario:
    Set getHorasPorFuncionario = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getHorasPorFuncionario:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getHorasPorFuncionario()" & vbNewLine & err.Description)
    Resume Exit_getHorasPorFuncionario

End Function

