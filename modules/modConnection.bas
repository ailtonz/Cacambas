Attribute VB_Name = "modConnection"


'Public Function OpenConnection(strBanco As infBanco) As ADODB.Connection
''' Build the connection string depending on the source
'Dim connectionString As String
'
'Select Case strBanco.strSource
'    Case "Access"
'        connectionString = "Provider=" & strBanco.strDriver & ";Data Source=" & strBanco.strDatabase
'    Case "Access2003"
'        connectionString = "Driver={" & strBanco.strDriver & "};Dbq=" & strBanco.strLocation & strBanco.strDatabase & ";Uid=" & strBanco.strUser & ";PWD=" & strBanco.strPassword & ""
'    Case "SQLite"
'        connectionString = "Driver={" & strBanco.strDriver & "};Database=" & strBanco.strDatabase
'    Case "MySQL"
'        connectionString = "Driver={" & strBanco.strDriver & "};Server=" & strBanco.strLocation & ";Database=" & strBanco.strDatabase & ";PORT=" & strBanco.strPort & ";UID=" & strBanco.strUser & ";PWD=" & strBanco.strPassword
'    Case "PostgreSQL"
'        connectionString = "Driver={" & strBanco.strDriver & "};Server=" & strBanco.strLocation & ";Database=" & strBanco.strDatabase & ";UID=" & strBanco.strUser & ";PWD=" & strBanco.strPassword
'    Case "SQL Server"
'        connectionString = "Driver={" & strBanco.strDriver & "};Server=" & strBanco.strLocation & ";Database=" & strBanco.strDatabase & ";Uid=" & strBanco.strUser & ";Pwd=" & strBanco.strPassword & ";"
'
'End Select
'
''' Create and open a new connection to the selected source
'Set OpenConnection = New ADODB.Connection
'Call OpenConnection.Open(connectionString)
'
'End Function

Public Function OpenConnectionNEW(banco As cDB) As ADODB.Connection
'' Build the connection string depending on the source
Dim connectionString As String
    
Select Case banco.Source
    Case "Access"
        connectionString = "Provider=" & banco.Driver & ";Data Source=" & banco.Database
    Case "Access2003"
        connectionString = "Driver={" & banco.Driver & "};Dbq=" & banco.Location & banco.Database & ";Uid=" & banco.User & ";PWD=" & banco.Password & ""
    Case "SQLite"
        connectionString = "Driver={" & banco.Driver & "};Database=" & banco.Database
    Case "MySQL"
        connectionString = "Driver={" & banco.Driver & "};Server=" & banco.Location & ";Database=" & banco.Database & ";PORT=" & banco.Port & ";UID=" & banco.User & ";PWD=" & banco.Password
    Case "PostgreSQL"
        connectionString = "Driver={" & banco.Driver & "};Server=" & banco.Location & ";Database=" & banco.Database & ";UID=" & banco.User & ";PWD=" & banco.Password
End Select

'' Create and open a new connection to the selected source
Set OpenConnectionNEW = New ADODB.Connection
Call OpenConnectionNEW.Open(connectionString)
   
End Function

Public Function carregarBanco() As cDB
Dim Bnc As New cDB

Dim wsBnc As Worksheet
Set wsBnc = Worksheets("cfg")

    With Bnc
        .Source = wsBnc.Range("C2").Value
        .Driver = wsBnc.Range("C3").Value
        .Location = wsBnc.Range("C4").Value
        .Database = wsBnc.Range("C5").Value
        .User = wsBnc.Range("C6").Value
        .Password = wsBnc.Range("C7").Value
        .Port = wsBnc.Range("C8").Value
        .add Bnc
    End With

Set carregarBanco = Bnc
Set wsBnc = Nothing

End Function

Public Sub AbrirNotas(ByVal obj As cMatriz)

With frmNotas
    .sCategoria = obj.Cat
    .sConsulta = obj.qry
    .sProcedure = obj.Prc
    .sTitulo = obj.Title
    .Show
End With

End Sub

Public Sub AbrirCadastros(ByVal obj As cMatriz, ByVal id As String)

With frmCadastro
    .sFrm = obj.Frm
    .sCategoria = obj.Cat
    .sConsulta = obj.qry
    .sProcedure = obj.Prc
    .sTitulo = obj.Title
    .sID = id
    .Show
End With

End Sub

Public Sub AbrirObras(ByVal obj As cMatriz, ByVal id As String)

With frmObras
    .sFrm = obj.Frm
    .sCategoria = obj.Cat
    .sConsulta = obj.qry
    .sProcedure = obj.Prc
    .sTitulo = obj.Title
    .sID = id
    .Show
End With

End Sub

Public Sub AbrirContatos(ByVal obj As cMatriz, ByVal id As String)

With frmContatos
    .sFrm = obj.Frm
    .sCategoria = obj.Cat
    .sConsulta = obj.qry
    .sProcedure = obj.Prc
    .sTitulo = obj.Title
    .sID = id
    .Show
End With

End Sub

Public Sub AbrirContratos(ByVal obj As cMatriz, ByVal id As String)

With frmContratos
    .sFrm = obj.Frm
    .sCategoria = obj.Cat
    .sConsulta = obj.qry
    .sProcedure = obj.Prc
    .sTitulo = obj.Title
    .sID = id
    .Show
End With

End Sub

Public Sub AbrirObservacoes(ByVal obj As cMatriz, ByVal id As String)

With frmObservacoes
    .sFrm = obj.Frm
    .sCategoria = obj.Cat
    .sConsulta = obj.qry
    .sProcedure = obj.Prc
    .sTitulo = obj.Title
    .sID = id
    .Show
End With

End Sub
