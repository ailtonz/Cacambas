Attribute VB_Name = "modFunctions"
Public Function Saida(strConteudo As String, strArquivo As String)
    Open CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & strArquivo For Append As #1
    Print #1, strConteudo
    Close #1
End Function
