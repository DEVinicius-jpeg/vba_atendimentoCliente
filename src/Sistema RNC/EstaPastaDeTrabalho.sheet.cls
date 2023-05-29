Public IdUsuario As Integer
Public nomeUsuario As String

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    fechaBancoDeDados
End Sub

Sub Workbook_Open()
    Application.Visible = False
    Frm_Login.Show
    Application.Visible = True
End Sub
