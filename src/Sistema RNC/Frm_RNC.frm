VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_RNC 
   Caption         =   "Relatório de Não Conformidade"
   ClientHeight    =   9135.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14910
   OleObjectBlob   =   "Frm_RNC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_RNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim querySelect As String
Dim rs As Recordset
Private xformat() As New ClassFormat

Private Sub btVisualizar_Click()
    Frm_VisualizaRnc.Show
End Sub

Private Sub userform_initialize()
    
    querySelect = "SELECT TOP 50 a.ID_rnc, a.data_abertura, a.nome_cliente, b.ds_naoconformidade, a.status ,a.situacao FROM Rnc a LEFT JOIN NaoConformidade b ON a.ID_naoconformidade = b.ID_naoconformidade ORDER BY ID_rnc DESC"
    
    Call carregaListViewRnc(querySelect)
    Call comboBoxNaoConformidade
    Call comboBoxStatusSituacao
    Call chamarFormat
 
End Sub

Function carregaListViewRnc(ByVal sql As String)

    Dim rs As ADODB.Recordset
    Dim linha As Integer
    Dim Inprocess As Boolean
    
    On Error GoTo error
    
    lvRnc.ColumnHeaders.Clear
    lvRnc.ListItems.Clear
    
    Inprocess = False
    
    conexaoAccess
    
    Set rs = getRecordset(sql)
    
    Inprocess = True
    
    With lvRnc
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .CheckBoxes = True
        .MultiSelect = False
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Add Text:="RNC", Width:=45, Alignment:=0
        .ColumnHeaders.Add Text:="Data", Width:=60, Alignment:=0
        .ColumnHeaders.Add Text:="Cliente", Width:=210, Alignment:=0
        .ColumnHeaders.Add Text:="Não Conformidade", Width:=100, Alignment:=0
        .ColumnHeaders.Add Text:="Status", Width:=100, Alignment:=0
        .ColumnHeaders.Add Text:="Situação", Width:=111, Alignment:=0
        
        linha = lvRnc.ListItems.Count + 1
        
        Do Until rs.EOF
            .ListItems.Add = rs.Fields("ID_rnc")
            .ListItems(linha).SubItems(1) = rs.Fields("data_abertura")
            .ListItems(linha).SubItems(2) = rs.Fields("nome_cliente")
            .ListItems(linha).SubItems(3) = rs.Fields("ds_naoconformidade")
            .ListItems(linha).SubItems(4) = statusRnc(, rs.Fields("status"))
            .ListItems(linha).SubItems(5) = situacaoRnc(, rs.Fields("situacao"))
            linha = lvRnc.ListItems.Count + 1
            rs.MoveNext
        Loop
    End With
    
    lbregistros.Caption = lvRnc.ListItems.Count
    
Function_Exit:
    fechaBancoDeDados
    Exit Function
    
error:
    If Inprocess = True Then
        Call carregaListViewRnc(querySelect)
    End If
    
    mensagemErro Err.Description
    Resume Function_Exit

End Function

Function buscaAvancada() As String

    Dim sql As String
    Dim controle As Control
    Dim I As Integer
        
    sql = "SELECT a.ID_rnc, a.data_abertura, a.nome_cliente, b.ds_naoconformidade, a.status ,a.situacao FROM Rnc a, NaoConformidade b, Produtos c WHERE a.ID_naoconformidade = b.ID_naoconformidade AND a.ID_rnc = c.ID_rnc"
    
    For I = 0 To Me.MultiPage1.Pages.Count - 1
        For Each controle In Me.MultiPage1.Pages(I).Controls
            If TypeName(controle) = "TextBox" Or TypeName(controle) = "ComboBox" Then
                If controle.Text <> Empty Then
                    Select Case controle.Tag
                        Case Is = "a.ID_rnc"
                            sql = sql & " AND " & controle.Tag & " = " & controle.Text
                        Case Is = "c.cd_produto"
                            sql = sql & " AND " & controle.Tag & " = " & controle.Text
                        Case Is = "a.nome_cliente", "c.ds_produto"
                            sql = sql & " AND " & controle.Tag & " LIKE " & "'%" & controle.Text & "%'"
                        Case Is = "b.ds_naoconformidade"
                            sql = sql & " AND " & controle.Tag & " = " & "'" & controle.Text & "'"
                        Case Is = "a.status"
                            sql = sql & " AND " & controle.Tag & " = " & statusRnc(controle.Text)
                        Case Is = "a.situacao"
                            sql = sql & " AND " & controle.Tag & " = " & situacaoRnc(controle.Text)
                        Case Is = "c.lote"
                            If Left(CStr(controle.Text), 1) = "2" Then
                                sql = sql & " AND " & controle.Tag & " = " & controle.Text
                            Else
                                sql = sql & " AND " & "c.op" & " = " & controle.Text
                            End If
                        Case Is = "a.data_abertura"
                            If Me.tbDataFinal.Text = "" Then
                                sql = sql & " AND " & controle.Tag & " = " & "#" & Format(controle.Text, "yyyy/MM/dd") & "#"
                            Else
                                sql = sql & " AND " & controle.Tag & " BETWEEN " & "#" & Format(controle.Text, "yyyy/MM/dd") & "#" & " AND " & "#" & Format(Me.tbDataFinal.Text, "yyyy/MM/dd") & "#"
                            End If
                    End Select
                End If
            End If
        Next controle
    Next I
    sql = sql & " GROUP BY a.ID_rnc, a.data_abertura, a.nome_cliente, b.ds_naoconformidade, a.status ,a.situacao ORDER BY a.ID_rnc DESC;"
    buscaAvancada = sql
End Function

Private Sub chamarFormat() 'Rotina que confere todos os objetos do fomulário para colocar a mascara de entrada de acordo com o objeto

    Dim I As Integer
    Dim cont As Integer
    
    cont = Me.Controls.Count - 1
    
    ReDim xformat(0 To cont)
        For I = 0 To cont
            Select Case Me.Controls(I).Tag
                Case Is = "a.ID_rnc", "c.cd_produto", "c.lote" ' Mascara de apenas numeros
                    Set xformat(I).toNumero = Me.Controls(I)
                Case Is = "a.data_abertura", "data" ' mascara de data
                    Set xformat(I).toData = Me.Controls(I)
            End Select
        Next
End Sub

Private Sub opAbertura_Click()
    Me.tbDataInicial.Tag = "a.data_abertura"
End Sub

Private Sub opFechamento_Click()
    Me.tbDataInicial.Tag = "a.data_fechamento"
End Sub

Private Sub comboBoxStatusSituacao()
    
    Dim I As Integer
    
    For I = 0 To 3
        Me.cbStatus.AddItem statusRnc(, I)
        Me.cbSituacao.AddItem situacaoRnc(, I)
    Next I

End Sub

Private Sub comboBoxNaoConformidade() 'Gera a lista suspensa das não conformidades

    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "SELECT * FROM NaoConformidade ORDER BY ds_naoconformidade ASC"

    conexaoAccess
    
    Set rs = getRecordset(sql)
    
    While Not rs.EOF
    
        Me.cbNaoConformidade.AddItem rs.Fields(1)
        rs.MoveNext
        
    Wend
    
   fechaBancoDeDados

End Sub

Private Sub btNovo_Click()
    
    Frm_CadastroDeRNC.Show
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub btPesquisar_Click()

    On Error GoTo erro
  
    Call carregaListViewRnc(buscaAvancada)
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)
End Sub

Private Sub tbRnc_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub tbCliente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub cbNaoConformidade_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub cbStatus_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub cbSituacao_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub tbDataInicial_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub tbDataFinal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub tbOpLote_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub tbCodigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub

Private Sub tbNomeProduto_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error GoTo erro
    If KeyCode = VBA.vbKeyReturn Then
        Call carregaListViewRnc(buscaAvancada)
    End If
    
    Exit Sub
erro:
    Call carregaListViewRnc(querySelect)

End Sub
