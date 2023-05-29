Attribute VB_Name = "MDlUtil"
Option Explicit
Option Private Module
Public Const senha As String = ""

Function getCaminho(nome As String) As String

    Dim caminho As String
    
    caminho = Replace(Replace(ThisWorkbook.Names(nome), "=", ""), """", "")
    
    getCaminho = caminho

End Function

