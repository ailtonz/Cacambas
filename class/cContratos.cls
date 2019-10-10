VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cContratos"
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


Private pContratoInicio As String
Private pContratoTerminio As String
Private pContratoNF As String
Private pContratoISS As String
Private pContratoCTR As String

Private pContratoPeriodoLocacao As String
Private pContratoTransacao As String
Private pContratoCondicoes As String
Private pContratoRetiradaAutomatica As String
Private pContratoVctoAposEntrega As String

Private pContratoValor As String
Private pContratoMultaMora As String
Private pContratoMultaDia As String

Private pContratoObservacao As String
Private pContratoObsColoca As String
Private pContratoObsTroca As String
Private pContratoObsRetira As String
Private pContratoObsLigacao As String


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

Public Property Get ContratoInicio() As String
    ContratoInicio = pContratoInicio
End Property

Public Property Let ContratoInicio(ByVal vValor As String)
    pContratoInicio = vValor
End Property

Public Property Get ContratoTerminio() As String
    ContratoTerminio = pContratoTerminio
End Property

Public Property Let ContratoTerminio(ByVal vValor As String)
    pContratoTerminio = vValor
End Property

Public Property Get ContratoNF() As String
    ContratoNF = pContratoNF
End Property

Public Property Let ContratoNF(ByVal vValor As String)
    pContratoNF = vValor
End Property

Public Property Get ContratoISS() As String
    ContratoISS = pContratoISS
End Property

Public Property Let ContratoISS(ByVal vValor As String)
    pContratoISS = vValor
End Property

Public Property Get ContratoCTR() As String
    ContratoCTR = pContratoCTR
End Property

Public Property Let ContratoCTR(ByVal vValor As String)
    pContratoCTR = vValor
End Property

Public Property Get ContratoPeriodoLocacao() As String
    ContratoPeriodoLocacao = pContratoPeriodoLocacao
End Property

Public Property Let ContratoPeriodoLocacao(ByVal vValor As String)
    pContratoPeriodoLocacao = vValor
End Property

Public Property Get ContratoTransacao() As String
    ContratoTransacao = pContratoTransacao
End Property

Public Property Let ContratoTransacao(ByVal vValor As String)
    pContratoTransacao = vValor
End Property

Public Property Get ContratoCondicoes() As String
    ContratoCondicoes = pContratoCondicoes
End Property

Public Property Let ContratoCondicoes(ByVal vValor As String)
    pContratoCondicoes = vValor
End Property

Public Property Get ContratoRetiradaAutomatica() As String
    ContratoRetiradaAutomatica = pContratoRetiradaAutomatica
End Property

Public Property Let ContratoRetiradaAutomatica(ByVal vValor As String)
    pContratoRetiradaAutomatica = vValor
End Property

Public Property Get ContratoVctoAposEntrega() As String
    ContratoVctoAposEntrega = pContratoVctoAposEntrega
End Property

Public Property Let ContratoVctoAposEntrega(ByVal vValor As String)
    pContratoVctoAposEntrega = vValor
End Property

Public Property Get ContratoValor() As String
    ContratoValor = pContratoValor
End Property

Public Property Let ContratoValor(ByVal vValor As String)
    pContratoValor = vValor
End Property

Public Property Get ContratoMultaMora() As String
    ContratoMultaMora = pContratoMultaMora
End Property

Public Property Let ContratoMultaMora(ByVal vValor As String)
    pContratoMultaMora = vValor
End Property

Public Property Get ContratoMultaDia() As String
    ContratoMultaDia = pContratoMultaDia
End Property

Public Property Let ContratoMultaDia(ByVal vValor As String)
    pContratoMultaDia = vValor
End Property

Public Property Get ContratoObservacao() As String
    ContratoObservacao = pContratoObservacao
End Property

Public Property Let ContratoObservacao(ByVal vValor As String)
    pContratoObservacao = vValor
End Property

Public Property Get ContratoObsColoca() As String
    ContratoObsColoca = pContratoObsColoca
End Property

Public Property Let ContratoObsColoca(ByVal vValor As String)
    pContratoObsColoca = vValor
End Property

Public Property Get ContratoObsTroca() As String
    ContratoObsTroca = pContratoObsTroca
End Property

Public Property Let ContratoObsTroca(ByVal vValor As String)
    pContratoObsTroca = vValor
End Property

Public Property Get ContratoObsRetira() As String
    ContratoObsRetira = pContratoObsRetira
End Property

Public Property Let ContratoObsRetira(ByVal vValor As String)
    pContratoObsRetira = vValor
End Property

Public Property Get ContratoObsLigacao() As String
    ContratoObsLigacao = pContratoObsLigacao
End Property

Public Property Let ContratoObsLigacao(ByVal vValor As String)
    pContratoObsLigacao = vValor
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

Public Sub add(ByVal rec As cContratos, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As cContratos)
    pCol.add rec, rec.id
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function cContratos(ByVal i As Variant) As cContratos
    Set cContratos = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As cContratos
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As cDB, obj As cContratos) As Boolean
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

    .Parameters.Append .CreateParameter("@pContratoInicio", adVarChar, adParamInput, 10, obj.ContratoInicio)
    .Parameters.Append .CreateParameter("@pContratoTerminio", adVarChar, adParamInput, 10, obj.ContratoTerminio)
    .Parameters.Append .CreateParameter("@pContratoNF", adVarChar, adParamInput, 1, obj.ContratoNF)
    .Parameters.Append .CreateParameter("@pContratoISS", adVarChar, adParamInput, 1, obj.ContratoISS)
    .Parameters.Append .CreateParameter("@pContratoCTR", adVarChar, adParamInput, 1, obj.ContratoCTR)
    .Parameters.Append .CreateParameter("@pContratoPeriodoLocacao", adVarChar, adParamInput, 3, obj.ContratoPeriodoLocacao)
    .Parameters.Append .CreateParameter("@pContratoTransacao", adVarChar, adParamInput, 50, obj.ContratoTransacao)
    .Parameters.Append .CreateParameter("@pContratoCondicoes", adVarChar, adParamInput, 50, obj.ContratoCondicoes)
    .Parameters.Append .CreateParameter("@pContratoRetiradaAutomatica", adVarChar, adParamInput, 3, obj.ContratoRetiradaAutomatica)
    .Parameters.Append .CreateParameter("@pContratoVctoAposEntrega", adVarChar, adParamInput, 3, obj.ContratoVctoAposEntrega)
    .Parameters.Append .CreateParameter("@pContratoValor", adVarChar, adParamInput, 50, obj.ContratoValor)
    .Parameters.Append .CreateParameter("@pContratoMultaMora", adVarChar, adParamInput, 50, obj.ContratoMultaMora)
    .Parameters.Append .CreateParameter("@pContratoMultaDia", adVarChar, adParamInput, 50, obj.ContratoMultaDia)
    .Parameters.Append .CreateParameter("@pContratoObservacao", adVarChar, adParamInput, 120, obj.ContratoObservacao)
    .Parameters.Append .CreateParameter("@pContratoObsColoca", adVarChar, adParamInput, 120, obj.ContratoObsColoca)
    .Parameters.Append .CreateParameter("@pContratoObsTroca", adVarChar, adParamInput, 120, obj.ContratoObsTroca)
    .Parameters.Append .CreateParameter("@pContratoObsRetira", adVarChar, adParamInput, 120, obj.ContratoObsRetira)
    .Parameters.Append .CreateParameter("@pContratoObsLigacao", adVarChar, adParamInput, 120, obj.ContratoObsLigacao)


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

Public Function Update(banco As cDB, obj As cContratos) As Boolean
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

    .Parameters.Append .CreateParameter("@pContratoInicio", adVarChar, adParamInput, 10, obj.ContratoInicio)
    .Parameters.Append .CreateParameter("@pContratoTerminio", adVarChar, adParamInput, 10, obj.ContratoTerminio)
    .Parameters.Append .CreateParameter("@pContratoNF", adVarChar, adParamInput, 1, obj.ContratoNF)
    .Parameters.Append .CreateParameter("@pContratoISS", adVarChar, adParamInput, 1, obj.ContratoISS)
    .Parameters.Append .CreateParameter("@pContratoCTR", adVarChar, adParamInput, 1, obj.ContratoCTR)
    .Parameters.Append .CreateParameter("@pContratoPeriodoLocacao", adVarChar, adParamInput, 3, obj.ContratoPeriodoLocacao)
    .Parameters.Append .CreateParameter("@pContratoTransacao", adVarChar, adParamInput, 50, obj.ContratoTransacao)
    .Parameters.Append .CreateParameter("@pContratoCondicoes", adVarChar, adParamInput, 50, obj.ContratoCondicoes)
    .Parameters.Append .CreateParameter("@pContratoRetiradaAutomatica", adVarChar, adParamInput, 3, obj.ContratoRetiradaAutomatica)
    .Parameters.Append .CreateParameter("@pContratoVctoAposEntrega", adVarChar, adParamInput, 3, obj.ContratoVctoAposEntrega)
    .Parameters.Append .CreateParameter("@pContratoValor", adVarChar, adParamInput, 50, obj.ContratoValor)
    .Parameters.Append .CreateParameter("@pContratoMultaMora", adVarChar, adParamInput, 50, obj.ContratoMultaMora)
    .Parameters.Append .CreateParameter("@pContratoMultaDia", adVarChar, adParamInput, 50, obj.ContratoMultaDia)
    .Parameters.Append .CreateParameter("@pContratoObservacao", adVarChar, adParamInput, 120, obj.ContratoObservacao)
    .Parameters.Append .CreateParameter("@pContratoObsColoca", adVarChar, adParamInput, 120, obj.ContratoObsColoca)
    .Parameters.Append .CreateParameter("@pContratoObsTroca", adVarChar, adParamInput, 120, obj.ContratoObsTroca)
    .Parameters.Append .CreateParameter("@pContratoObsRetira", adVarChar, adParamInput, 120, obj.ContratoObsRetira)
    .Parameters.Append .CreateParameter("@pContratoObsLigacao", adVarChar, adParamInput, 120, obj.ContratoObsLigacao)


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

Public Function getContratos(banco As cDB, qryEntidade As String) As cContratos
On Error GoTo Err_getContratos
Dim colObj As New cContratos
Dim recObj As New cContratos

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
    Set recObj = New cContratos
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .ContratoInicio = CStr(IIf(IsNull(rst.Fields("ContratoInicio")), "", rst.Fields("ContratoInicio")))
        .ContratoTerminio = CStr(IIf(IsNull(rst.Fields("ContratoTerminio")), "", rst.Fields("ContratoTerminio")))
        .ContratoNF = CStr(IIf(IsNull(rst.Fields("ContratoNF")), "", rst.Fields("ContratoNF")))
        .ContratoISS = CStr(IIf(IsNull(rst.Fields("ContratoISS")), "", rst.Fields("ContratoISS")))
        .ContratoCTR = CStr(IIf(IsNull(rst.Fields("ContratoCTR")), "", rst.Fields("ContratoCTR")))
        
        .ContratoPeriodoLocacao = CStr(IIf(IsNull(rst.Fields("ContratoPeriodoLocacao")), "", rst.Fields("ContratoPeriodoLocacao")))
        .ContratoTransacao = CStr(IIf(IsNull(rst.Fields("ContratoTransacao")), "", rst.Fields("ContratoTransacao")))
        .ContratoCondicoes = CStr(IIf(IsNull(rst.Fields("ContratoCondicoes")), "", rst.Fields("ContratoCondicoes")))
        .ContratoRetiradaAutomatica = CStr(IIf(IsNull(rst.Fields("ContratoRetiradaAutomatica")), "", rst.Fields("ContratoRetiradaAutomatica")))
        .ContratoVctoAposEntrega = CStr(IIf(IsNull(rst.Fields("ContratoVctoAposEntrega")), "", rst.Fields("ContratoVctoAposEntrega")))
        .ContratoValor = CStr(IIf(IsNull(rst.Fields("ContratoValor")), "", rst.Fields("ContratoValor")))
        .ContratoMultaMora = CStr(IIf(IsNull(rst.Fields("ContratoMultaMora")), "", rst.Fields("ContratoMultaMora")))
        .ContratoMultaDia = CStr(IIf(IsNull(rst.Fields("ContratoMultaDia")), "", rst.Fields("ContratoMultaDia")))
        
        .ContratoObservacao = CStr(IIf(IsNull(rst.Fields("ContratoObservacao")), "", rst.Fields("ContratoObservacao")))
        .ContratoObsColoca = CStr(IIf(IsNull(rst.Fields("ContratoObsColoca")), "", rst.Fields("ContratoObsColoca")))
        .ContratoObsTroca = CStr(IIf(IsNull(rst.Fields("ContratoObsTroca")), "", rst.Fields("ContratoObsTroca")))
        .ContratoObsRetira = CStr(IIf(IsNull(rst.Fields("ContratoObsRetira")), "", rst.Fields("ContratoObsRetira")))
        .ContratoObsLigacao = CStr(IIf(IsNull(rst.Fields("ContratoObsLigacao")), "", rst.Fields("ContratoObsLigacao")))

        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getContratos:
    Set getContratos = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getContratos:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getContratos()" & vbNewLine & err.Description)
    Resume Exit_getContratos

End Function

Public Function getContratosID(banco As cDB, qryEntidade As String, id As String) As cContratos
On Error GoTo Err_getContratosID
Dim colObj As New cContratos
Dim recObj As New cContratos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from " & qryEntidade & " where id = " & id & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New cContratos
    With recObj
       
        .id = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .ContratoInicio = CStr(IIf(IsNull(rst.Fields("ContratoInicio")), "", rst.Fields("ContratoInicio")))
        .ContratoTerminio = CStr(IIf(IsNull(rst.Fields("ContratoTerminio")), "", rst.Fields("ContratoTerminio")))
        .ContratoNF = CStr(IIf(IsNull(rst.Fields("ContratoNF")), "", rst.Fields("ContratoNF")))
        .ContratoISS = CStr(IIf(IsNull(rst.Fields("ContratoISS")), "", rst.Fields("ContratoISS")))
        .ContratoCTR = CStr(IIf(IsNull(rst.Fields("ContratoCTR")), "", rst.Fields("ContratoCTR")))
        
        .ContratoPeriodoLocacao = CStr(IIf(IsNull(rst.Fields("ContratoPeriodoLocacao")), "", rst.Fields("ContratoPeriodoLocacao")))
        .ContratoTransacao = CStr(IIf(IsNull(rst.Fields("ContratoTransacao")), "", rst.Fields("ContratoTransacao")))
        .ContratoCondicoes = CStr(IIf(IsNull(rst.Fields("ContratoCondicoes")), "", rst.Fields("ContratoCondicoes")))
        .ContratoRetiradaAutomatica = CStr(IIf(IsNull(rst.Fields("ContratoRetiradaAutomatica")), "", rst.Fields("ContratoRetiradaAutomatica")))
        .ContratoVctoAposEntrega = CStr(IIf(IsNull(rst.Fields("ContratoVctoAposEntrega")), "", rst.Fields("ContratoVctoAposEntrega")))
        .ContratoValor = CStr(IIf(IsNull(rst.Fields("ContratoValor")), "", rst.Fields("ContratoValor")))
        .ContratoMultaMora = CStr(IIf(IsNull(rst.Fields("ContratoMultaMora")), "", rst.Fields("ContratoMultaMora")))
        .ContratoMultaDia = CStr(IIf(IsNull(rst.Fields("ContratoMultaDia")), "", rst.Fields("ContratoMultaDia")))
        
        .ContratoObservacao = CStr(IIf(IsNull(rst.Fields("ContratoObservacao")), "", rst.Fields("ContratoObservacao")))
        .ContratoObsColoca = CStr(IIf(IsNull(rst.Fields("ContratoObsColoca")), "", rst.Fields("ContratoObsColoca")))
        .ContratoObsTroca = CStr(IIf(IsNull(rst.Fields("ContratoObsTroca")), "", rst.Fields("ContratoObsTroca")))
        .ContratoObsRetira = CStr(IIf(IsNull(rst.Fields("ContratoObsRetira")), "", rst.Fields("ContratoObsRetira")))
        .ContratoObsLigacao = CStr(IIf(IsNull(rst.Fields("ContratoObsLigacao")), "", rst.Fields("ContratoObsLigacao")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getContratosID:
    Set getContratosID = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getContratosID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getContratosID()" & vbNewLine & err.Description)
    Resume Exit_getContratosID

End Function
