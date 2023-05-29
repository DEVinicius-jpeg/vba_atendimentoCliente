VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Login 
   Caption         =   "Login ao Sistema de RNC"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5760
   OleObjectBlob   =   "Frm_Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btEntrar_Click()

    Dim rs As ADODB.Recordset
    Dim querySelect As String
    
    
    conexaoAccess
    
    querySelect = "SELECT * FROM Usuarios"
    
    Set rs = getRecordset(querySelect)
    
    ' Verifica se o Login está correto
    While UCase(Me.tbLogin) <> rs.Fields("login_usuario")
        rs.MoveNext
        If rs.EOF = True Then
            mensagemInformacao "Usuário não cadastrado"
            fechaBancoDeDados
            Exit Sub
        End If
    Wend
    ' Verifica se a senha está correta
    While UCase(Me.tbSenha) <> rs.Fields("senha_usuario")
        rs.MoveNext
        If rs.EOF = True Then
            mensagemInformacao "Senha incorreta"
            fechaBancoDeDados
            Exit Sub
        End If
    Wend
    
    nomeUsuario = rs.Fields("login_usuario")
    
    IdUsuario = rs.Fields("ID_usuario")
    
    fechaBancoDeDados
    
    Unload Frm_Login
    
    Frm_RNC.Show
    
End Sub

Private Sub chExibir_Click()
    '
    If Me.chExibir.Value = True Then
        Me.tbSenha.PasswordChar = ""
    Else
        Me.tbSenha.PasswordChar = "*"
    End If
    
End Sub

Private Sub UserForm_Click()

End Sub
