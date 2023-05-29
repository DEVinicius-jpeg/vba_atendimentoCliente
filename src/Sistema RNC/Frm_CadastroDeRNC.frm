VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_CadastroDeRNC 
   Caption         =   "Cadastro de RNC"
   ClientHeight    =   9585.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10095
   OleObjectBlob   =   "Frm_CadastroDeRNC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_CadastroDeRNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variaveis privadas do fomulário
Private caminhoImagem As String
Private lotePai As Long
Private dataFabricacao As Date
Private codigoProduto As Long
Private xformat() As New ClassFormat

Private Sub userform_initialize() 'Chama rotinas quando o formulario é iniciado

    Call comboBoxDepartamento
    Call comboBoxNaoConformidade
    Call inicializaDados
    Call chamarFormat
    Call carregaListViewProdutos
    Call carregaListViewAnexos
    
End Sub

Private Sub carregaListViewProdutos() 'ListView dos produtos

    With lvProdutos
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Add Text:="Lote", Width:=65, Alignment:=0
        .ColumnHeaders.Add Text:="Nome Do Produto ", Width:=200, Alignment:=0
        .ColumnHeaders.Add Text:="Vencimento", Width:=60, Alignment:=0
        .ColumnHeaders.Add Text:="Qtd Total", Width:=60, Alignment:=0
        .ColumnHeaders.Add Text:="Reclamado", Width:=60, Alignment:=0
        .ColumnHeaders.Add Text:="OP", Width:=0, Alignment:=0
        .ColumnHeaders.Add Text:="Codigo", Width:=0, Alignment:=0
        .ColumnHeaders.Add Text:="Fabricação", Width:=0, Alignment:=0
        
    End With

End Sub

Private Sub lvProdutos_KeyDown(KeyCode As Integer, ByVal Shift As Integer) ' Evento que exclui um item selecionado da lista de produtos ao teclar o backspace
    If KeyCode = VBA.vbKeyBack Then
        lvProdutos.ListItems.Remove (lvProdutos.SelectedItem.Index)
    End If
End Sub

Private Sub carregaListViewAnexos() 'ListView de anexos

    With lvAnexos
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Add Text:="Imagem", Width:=249, Alignment:=0
    End With

End Sub

Private Sub lvAnexos_KeyDown(KeyCode As Integer, ByVal Shift As Integer) 'Evento que exclui um item selecionado da lista de anexos ao teclar o backspace
    If KeyCode = VBA.vbKeyBack Then
        lvAnexos.ListItems.Remove (lvAnexos.SelectedItem.Index)
    End If
End Sub

Private Sub inicializaDados() ' Gera a data da RNC e da analise das causas
        
    Me.tbDataRNC.Text = Date
    Me.tbData.Text = Date
    Me.tbNome.Text = nomeUsuario
    
End Sub

Private Sub chamarFormat() 'Rotina que confere todos os objetos do fomulário para colocar a mascara de entrada de acordo com o objeto

    Dim I As Integer
    Dim cont As Integer
    
    cont = Me.Controls.Count - 1
    
    ReDim xformat(0 To cont)
        For I = 0 To cont
            Select Case Me.Controls(I).Tag
                Case Is = "numero", "CPF CNPJ", "Código" ' Mascara de apenas numeros
                    Set xformat(I).toNumero = Me.Controls(I)
                Case Is = "Data" ' mascara de data
                    Set xformat(I).toData = Me.Controls(I)
                Case Is = "Telefone" ' mascara de telefone
                    Set xformat(I).toTelefone = Me.Controls(I)
            End Select
        Next
End Sub

Private Sub comboBoxDepartamento() 'Gera a lista suspensa dos departamentos

    Dim querySelect As String
    Dim rs As ADODB.Recordset
    
    querySelect = "SELECT * FROM Departamento ORDER BY ds_departamento ASC"

    conexaoAccess
    
    Set rs = getRecordset(querySelect)
    
    While Not rs.EOF
    
        Me.cbDepartamento.AddItem rs.Fields(1)
        rs.MoveNext
        
    Wend
    
   fechaBancoDeDados
   
End Sub

Private Sub comboBoxNaoConformidade() 'Gera a lista suspensa das não conformidades
                                        '1# Gerar a lista de não conformidades de acordo com o tipo de RNC
    Dim querySelect As String
    Dim rs As ADODB.Recordset
    
    querySelect = "SELECT * FROM NaoConformidade ORDER BY ds_naoconformidade ASC"

    conexaoAccess
    
    Set rs = getRecordset(querySelect)
    
    While Not rs.EOF
    
        Me.cbNaoConformidade.AddItem rs.Fields(1)
        rs.MoveNext
        
    Wend
    
   fechaBancoDeDados

End Sub

Private Sub cbDepartamento_Change() 'Gera a lista suspensa da area de detecção conforme o departamento

    Dim querySelect As String
    Dim rs As ADODB.Recordset
    
    Me.cbAreaDeteccao.Clear
    
    querySelect = "SELECT A.* FROM AreaDeteccao A LEFT JOIN Departamento D ON D.ID_departamento = A.ID_departamento WHERE ds_departamento LIKE '" & Me.cbDepartamento.Value & "' ORDER BY ds_area ASC"

    conexaoAccess
    
    Set rs = getRecordset(querySelect)
    
    While Not rs.EOF
    
        Me.cbAreaDeteccao.AddItem rs.Fields(2)
        rs.MoveNext
        
    Wend
    
   fechaBancoDeDados
   
End Sub

Private Sub tbCodigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'Consulta o Código digitado na textBox no banco de dados do SQLServer e rotorna inf do cliente

    Dim comp As Integer
    Dim rs As ADODB.Recordset
    Dim querySelect As String
    
    querySelect = "SELECT Nome, CNPJ_CPF FROM CLIENTE WHERE cd_cliente = '" & Me.tbCodigo.Text & "'"
    
    comp = Len(Me.tbCodigo.Text)

    If KeyCode = VBA.vbKeyReturn Then
    
        If comp = 6 Then
               
            conexaoSQLServer
            
            Set rs = getRecordset(querySelect)
            
            If rs.BOF = False Then
            
                Me.tbCliente.Text = RTrim(rs.Fields("Nome"))
                Me.tbCPF_CNPJ.Text = RTrim(rs.Fields("CNPJ_CPF"))
                fechaBancoDeDados
                Me.tbContato.SetFocus
                
            Else
                fechaBancoDeDados
                mensagemInformacao "Código Invalido!"
                Me.tbCliente.Text = Empty
                Me.tbCPF_CNPJ.Text = Empty
                Me.tbCPF_CNPJ.SetFocus
            End If
            
        Else
            mensagemInformacao "Código Invalido!"
            
        End If
        
    End If
End Sub

Private Sub tbCPF_CNPJ_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'Consulta o CNPJ digitado na textBox no banco de dados do SQLServer e rotorna inf do cliente
                                                                                                
    Dim comp As Integer
    Dim rs As ADODB.Recordset
    Dim querySelect As String
    
    querySelect = "SELECT Nome, cd_cliente FROM CLIENTE WHERE CNPJ_CPF = '" & Me.tbCPF_CNPJ.Text & "'"
    
    comp = Len(Me.tbCPF_CNPJ.Text)

    If KeyCode = VBA.vbKeyReturn Then
    
        If comp = 14 Or comp = 11 Then
               
            conexaoSQLServer
            
            Set rs = getRecordset(querySelect)
            
            If rs.BOF = False Then
            
                Me.tbCliente.Text = RTrim(rs.Fields("Nome"))
                Me.tbCodigo.Text = rs.Fields("cd_cliente")
                fechaBancoDeDados
                Me.tbContato.SetFocus
                
            Else
                fechaBancoDeDados
                mensagemInformacao "CNPJ / CPF Invalido!"
                Me.tbCliente.Text = Empty
                Me.tbCodigo.Text = Empty
                Me.tbCPF_CNPJ.SetFocus
            End If
            
        Else
            mensagemInformacao "CNPJ / CPF Invalido!"
            
        End If
        
    End If
    
End Sub

Private Sub tbLote_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) ''Consulta o lote digitado na textBox no banco de dados do Fb e rotorna inf do produto
    
    If KeyCode = VBA.vbKeyReturn Then
        
        Dim rs As ADODB.Recordset
        Dim querySelect As String
        Dim validade As Long
        
        On Error GoTo erro
        conexaoFireBird
        
        querySelect = "SELECT opi.OP_ID, ft.FT_VALIDADE, opi.OI_DATA, p.PD_NOME ,p.PD_ID, opi.OI_QUANTIDADE FROM ORDEM_PRODUCAO_ITEM opi LEFT JOIN FICHA_TECNICA ft ON opi.FT_ID = ft.FT_ID LEFT JOIN PRODUTO p ON ft.PD_ID = p.PD_ID WHERE opi.OI_ID = '" & Me.tbLote.Value & "'"
                
        Set rs = getRecordset(querySelect)
        
        If rs.BOF = False Then
            
            lotePai = rs.Fields("OP_ID")
            validade = rs.Fields("FT_VALIDADE")
            dataFabricacao = rs.Fields("OI_DATA")
            codigoProduto = rs.Fields("PD_ID")
        
            Me.tbQtdProducao.Value = rs.Fields("OI_QUANTIDADE")
            Me.tbNomeProduto.Value = RTrim(rs.Fields("PD_NOME"))
            Me.tbVencimento.Value = CalculoValidade(rs.Fields("OI_DATA"), validade)
            fechaBancoDeDados
            Me.tbQtdReclamacao.SetFocus
            
        Else
erro:
            fechaBancoDeDados
            mensagemInformacao "Lote invalido"
            Me.tbQtdProducao.Value = Empty
            Me.tbNomeProduto.Value = Empty
            Me.tbVencimento.Value = Empty
            
        End If
              
    End If

End Sub

Private Sub tbQtdReclamacao_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) ' Evento para quando o usuario teclar enter enviar as inf do produto para a listView de produtos

    Dim linha As Integer
    Dim existeLote As Boolean
    Dim I As Integer
    
    existeLote = False
    
    If KeyCode = VBA.vbKeyReturn Then
        If Me.tbQtdReclamacao.Value > 0 And Me.tbQtdReclamacao.Value <> "" And Me.tbQtdReclamacao.Value <= Me.tbQtdProducao.Value Then
            linha = lvProdutos.ListItems.Count + 1
            With lvProdutos
                For I = 1 To .ListItems.Count
                    If .ListItems(I) = Me.tbLote.Value Then existeLote = True
                Next I
                If existeLote = False Then
                    lvProdutos.ListItems.Add = Me.tbLote.Value
                    lvProdutos.ListItems(linha).SubItems(1) = Me.tbNomeProduto.Text
                    lvProdutos.ListItems(linha).SubItems(2) = Me.tbVencimento.Text
                    lvProdutos.ListItems(linha).SubItems(3) = Me.tbQtdProducao.Value
                    lvProdutos.ListItems(linha).SubItems(4) = Me.tbQtdReclamacao.Value
                    lvProdutos.ListItems(linha).SubItems(5) = lotePai
                    lvProdutos.ListItems(linha).SubItems(6) = codigoProduto
                    lvProdutos.ListItems(linha).SubItems(7) = dataFabricacao
                        
                    Me.tbLote.Text = Empty
                    Me.tbNomeProduto.Text = Empty
                    Me.tbQtdProducao.Text = Empty
                    Me.tbVencimento.Text = Empty
                    Me.tbQtdReclamacao.Text = Empty
            
                    Me.tbLote.SetFocus
                Else
                    mensagemInformacao "Lote já inserido!"
                End If
            End With
        Else
            Me.tbQtdReclamacao.SetFocus
            mensagemInformacao "A Quantidade Reclamada é Invalida!"
        End If
    End If
    
End Sub

Private Sub btCarregarImagem_Click() 'Abre o diretorio para o usuario selecionar uma imagem

    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = "C:\Users\Cameras\Downloads"
        .Title = "Selecione uma Imagem"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Pictures Files", "*.jpg, *.jpeg", 1
        .Show
        On Error GoTo erro
        caminhoImagem = .SelectedItems.Item(1)
        Me.Image1.Picture = LoadPicture(caminhoImagem)
    End With
    
    If VBA.InStr(FullPath, ".xlsx") = 0 Then
    
erro:
        Exit Sub
        
    End If
    
End Sub

Private Sub btEnviarImagem_Click() ' Envia a imagem selecionada pelo usuario para a listview de anexos

    Dim I As Integer
    Dim existeImagem As Boolean
    
    existeImagem = False
    
    With lvAnexos
        For I = 1 To .ListItems.Count
            If caminhoImagem = .ListItems(I) Then existeImagem = True
        Next I
            
        If existeImagem = False Then
            .ListItems.Add = caminhoImagem
            Me.Image1.Picture = Nothing
        Else
            mensagemInformacao "Imagem já inserida!"
            Me.Image1.Picture = Nothing
        End If
 
    End With
End Sub

Private Sub btEnviar_Click() 'Cadastrar RNC

    'Controle
    Dim status As Integer, situacao As Integer
    Dim dataNovo As Date
    'Cabeçalho
    Dim Data As Date
    'ComboBox
    Dim areaDeDeteccao As Integer, departamento As Integer, naoConformidade As Integer
    'CheckBox
    Dim reincidente As Integer
    'pgCliente
    Dim Codigo As String, cliente As String, Contato As String, cpf_cnpfj As String, Telefone As String, obs As String
    'pgDescrições
    Dim fatoOcorrido As String, nome As String
    Dim dataRegistro As Date
    'Conexão
    Dim rs As ADODB.Recordset
    
    conexaoAccess
    
    querySelect = "SELECT ID_departamento FROM Departamento WHERE ds_departamento = '" & Me.cbDepartamento.Value & "'"
    Set rs = getRecordset(querySelect)
    If rs.EOF = False Then departamento = rs.Fields("ID_departamento")
    
    querySelect = "SELECT ID_area FROM AreaDeteccao WHERE ds_area = '" & Me.cbAreaDeteccao.Value & "'"
    Set rs = getRecordset(querySelect)
    If rs.EOF = False Then areaDeDeteccao = rs.Fields("ID_AREA")
    
    querySelect = "SELECT ID_naoconformidade FROM NaoConformidade WHERE ds_naoconformidade = '" & Me.cbNaoConformidade.Value & "'"
    Set rs = getRecordset(querySelect)
    If rs.EOF = False Then naoConformidade = rs.Fields("ID_naoconformidade")
    
    fechaBancoDeDados
    
    'Cabeçalho
    Data = Format(Me.tbDataRNC.Value, "yyyy/MM/dd")
    'CheckBox
    reincidente = CBool(Me.chReincidente.Value)
    'pgCliente
    cliente = UCase(Me.tbCliente.Value)
    Contato = UCase(Me.tbContato.Value)
    Codigo = Me.tbCodigo.Value
    CPF_CNPJ = Me.tbCPF_CNPJ.Value
    Telefone = Me.tbTelefone.Value
    obs = UCase(Me.tbObservacao.Value)
    'pgDescrições
    fatoOcorrido = UCase(Me.tbDescricao.Value)
    nome = UCase(Me.tbNome.Value)
    dataRegistro = Format(Me.tbData.Value, "yyyy/MM/dd")
    'Controle
    situacao = 0
    status = 0
    dataNovo = Format(Now, "yyyy/MM/dd hh:mm:ss")
    

    If verificaVazio(Frm_CadastroDeRNC) = False Then
        Call inserirDados(dataNovo, Data, cliente, Codigo, CPF_CNPJ, Contato, Telefone, departamento, areaDeDeteccao, naoConformidade, reincidente, status, situacao, _
                        obs, fatoOcorrido, nome, dataRegistro, Me.lvProdutos, Me.lvAnexos)
                        
        Unload Frm_CadastroDeRNC
   End If
End Sub

Private Sub btSair_Click() ' Cancelar cadastro

    Unload Frm_CadastroDeRNC
    
End Sub
