VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_VisualizaRnc 
   Caption         =   "UserForm1"
   ClientHeight    =   11430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12600
   OleObjectBlob   =   "Frm_VisualizaRnc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_VisualizaRnc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub userform_initialize()

    Dim contado, linha As Integer
    Dim controle As Control
    Dim rs As Recordset
    Dim querySelectRnc, querySelectProd, querySelectDesc, querySelectAnexos As String
    
    querySelectRnc = "SELECT a.ID_rnc, a.data_abertura, a.nome_cliente, a.cd_cliente, a.CPF_CNPJ, a.contato_cliente, a.numero_contato, d.ds_departamento, b.ds_area, a.gestao_risco, a.gestao_mudanca, a.status, a.reincidente , a.obs_rnc , c.ds_naoconformidade, a.situacao, a.data_fechamento FROM Rnc a, AreaDeteccao b, NaoConformidade c, Departamento d WHERE a.ID_rnc = " & Frm_RNC.lvRnc.SelectedItem & ";"
    
    querySelectProd = "SELECT p.cd_produto, p.ds_produto, p.op, p.lote, p.qtd_reclamado, p.qtd_produzido, p.data_vencimento, p.data_fabricacao FROM Produtos p WHERE p.ID_rnc = " & Frm_RNC.lvRnc.SelectedItem & ";"
    
    querySelectDesc = "SELECT a.ds_fatoOcorrido, a.Nome1, a.Data1, a.ds_analiseCausa, a.Nome2, a.Data2, a.ds_acaoCorrecao, a.Nome3, a.Data3, a.ds_implementacaoAcaoCorrecao, a.Nome4, a.Data4, a.ds_acompanhamentoAcaoCorrecao, a.Nome5, a.Data5, a.ds_vericaoEficacia, a.Nome6, a.Data6, a.ds_propostaPreventiva, a.Nome7, a.Data7, a.ds_acompanhamentoAcaoPreventiva, a.Nome8, a.Data8, a.ds_conclusao, a.Nome9, a.Data9 FROM Acompanhamento a WHERE a.ID_rnc = " & Frm_RNC.lvRnc.SelectedItem & ";"
    
    querySelectAnexos = "SELECT * FROM Imagens i WHERE i.ID_rnc = " & Frm_RNC.lvRnc.SelectedItem & ";"
    
    
    conexaoAccess
    
    Set rs = getRecordset(querySelectRnc)
    
    Me.tbRnc.Text = "RNC Nº " & rs.Fields(0)
    Me.tbDataRNC.Text = rs.Fields(1)
    Me.tbCliente.Text = rs.Fields(2)
    Me.tbCodigo.Text = rs.Fields(3)
    Me.tbCPF_CNPJ.Text = rs.Fields(4)
    Me.tbContato.Text = rs.Fields(5)
    Me.tbTelefone.Text = rs.Fields(6)
    Me.tbDepartamento.Text = rs.Fields(7)
    Me.tbAreaDeteccao.Text = rs.Fields(8)
    Me.chGestaoDeRisco.Value = rs.Fields(9)
    Me.chGestaoDeMudanca.Value = rs.Fields(10)
    Me.chReincidente.Value = rs.Fields(11)
    Me.tbObservacao.Text = rs.Fields(12)
    Me.tbNaoConformidade.Text = rs.Fields(13)
    
    Set rs = getRecordset(querySelectDesc)
    
    contador = 0
    
    For Each controle In Me.MultiPage2.Pages(0).Controls
        If TypeName(controle) = "TextBox" Then
            If IsNull(rs.Fields(contador)) = False Then
                controle.Text = rs.Fields(contador)
            End If
            contador = contador + 1
        End If
    Next controle
    
    Set rs = getRecordset(querySelectProd)
    
        With Me.lvProdutos
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Add Text:="Codigo", Width:=50, Alignment:=0
        .ColumnHeaders.Add Text:="Nome Do Produto ", Width:=180, Alignment:=0
        .ColumnHeaders.Add Text:="Nº OP", Width:=50, Alignment:=0
        .ColumnHeaders.Add Text:="Lote", Width:=50, Alignment:=0
        .ColumnHeaders.Add Text:="Reclamado", Width:=55, Alignment:=0
        .ColumnHeaders.Add Text:="Qtd Total", Width:=55, Alignment:=0
        .ColumnHeaders.Add Text:="Vencimento", Width:=60, Alignment:=0
        .ColumnHeaders.Add Text:="Fabricação", Width:=60, Alignment:=0
        
        linha = lvProdutos.ListItems.Count + 1
        
        Do Until rs.EOF
            .ListItems.Add = rs.Fields(0)
            .ListItems(linha).SubItems(1) = rs.Fields(1)
            .ListItems(linha).SubItems(2) = rs.Fields(2)
            .ListItems(linha).SubItems(3) = rs.Fields(2)
            .ListItems(linha).SubItems(4) = rs.Fields(4)
            .ListItems(linha).SubItems(5) = rs.Fields(5)
            .ListItems(linha).SubItems(6) = rs.Fields(6)
            .ListItems(linha).SubItems(7) = rs.Fields(7)
            linha = lvProdutos.ListItems.Count + 1
            rs.MoveNext
        Loop
    End With
    fechaBancoDeDados
    
    
End Sub

Private Sub UserForm_Activate()
    With Me
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = .InsideHeight * 1.9
    End With
End Sub


