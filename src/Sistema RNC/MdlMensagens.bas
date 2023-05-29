Attribute VB_Name = "MdlMensagens"
Option Explicit

Function mensagemInformacao(ByVal mensagem As String) As String

    MsgBox mensagem, VBA.VbMsgBoxStyle.vbInformation, "Gerenciamento de Dados"
    
End Function

Function mensagemErro(ByVal mensagem As String) As String

    MsgBox mensagem, VBA.VbMsgBoxStyle.vbCritical, "Gerenciamento de Dados"
    
End Function
