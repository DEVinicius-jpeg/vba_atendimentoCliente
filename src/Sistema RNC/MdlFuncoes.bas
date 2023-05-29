Attribute VB_Name = "MdlFuncoes"
Public Function inserirDados(ByVal dataNovo As Date, ByVal dataRnc As Date, ByVal cliente As String, ByVal Codigo As String, ByVal cpf_cnpfj As String, ByVal Contato As String, ByVal Telefone As String, ByVal departamento As Integer, ByVal areaDeDeteccao As Integer, ByVal naoConformidade As Integer, _
                        ByVal reincidente As Integer, ByVal status As Integer, ByVal situacao As Integer, ByVal obs As String, ByVal fatoOcorrido As String, ByVal nome As String, ByVal dataRegistro As Date, ByVal listaProdutos As Object, ByVal listaAnexo)
    '
    'Realiza o Insert no banco de dados do Access
    '
    
    'Variaveis da ListView de Produtos
    Dim op As Long, codigoProduto As Long, lote As Long
    Dim vencimento As Date, dataFabricacao As Date
    Dim qtdProducao As Integer, qtdReclamacao As Integer
    Dim descricaoProduto As String
    
    'Variável da ListView de Imagens
    Dim imagem As String
    
    'Variáveis de Controle
    Dim lastID As Long
    Dim I As Integer
    Dim queryInsert As String, querySelect As String
    
    'Variáveis de conexão(DAO) com o Banco de dados do ACCESS
    'É utilizado a conexão DAO ao invés da ADODB pois a DAO possibilita acesso aos objetos do ACCESS.
    Dim objWs As DAO.Workspace
    Dim blnInprocess As Boolean
    Dim db As DAO.Database

    On Error GoTo error
    
    blnInprocess = False
    Set objWs = DBEngine.Workspaces(0)
    Set db = objWs.OpenDatabase(getCaminho("caminhoBancoAccess"), False, False, "MS Access;PWD=")
    blnInprocess = True
    objWs.BeginTrans 'Início da Transação
    
    'Insert 1: Insert na tabela de RNC
    queryInsert = "Insert Into Rnc( data_novo, data_abertura, nome_cliente, cd_cliente, CPF_CNPJ, contato_cliente, numero_contato, ID_departamento, ID_area, ID_naoconformidade, reincidente, status, Situacao, obs_rnc)"
    queryInsert = queryInsert & "Values('" & dataNovo & "','" & dataRnc & "','" & cliente & "','" & Codigo & "','" & cpf_cnpfj & "','" & Contato & "', '" & Telefone & "', " & departamento & ", " & areaDeDeteccao & ", " & naoConformidade & ", " & reincidente & ", " & status & ", " & situacao & ",'" & obs & "');"
    db.Execute queryInsert, dbFailOnError

    Set rds = db.OpenRecordset("SELECT @@IDENTITY AS newID") 'Pegando o ID do último insert realizado

    lastID = rds.Fields("newID")
    
    'Insert 2: Insert na tabela de Descrições
    queryInsert = "Insert Into Acompanhamento(ID_rnc, ds_fatoOcorrido, Nome1, Data1)"
    queryInsert = queryInsert & "Values(" & lastID & ",'" & fatoOcorrido & "','" & nome & "','" & dataRegistro & "')"
    db.Execute queryInsert, dbFailOnError
    
    'Insert 3: Insert na Tabela de Produtos
    With listaProdutos
        For I = 1 To .ListItems.Count
            lote = .ListItems(I)
            descricaoProduto = RTrim(.ListItems(I).SubItems(1))
            vencimento = .ListItems(I).SubItems(2)
            qtdProducao = .ListItems(I).SubItems(3)
            qtdReclamacao = .ListItems(I).SubItems(4)
            op = .ListItems(I).SubItems(5)
            codigoProduto = .ListItems(I).SubItems(6)
            dataFabricacao = .ListItems(I).SubItems(7)
            queryInsert = "Insert Into Produtos(ID_rnc, cd_produto, ds_produto, lote, op, qtd_produzido, qtd_reclamado, data_fabricacao, data_vencimento)"
            queryInsert = queryInsert & "Values(" & lastID & ", " & codigoProduto & ",'" & descricaoProduto & "', " & lote & ", " & op & "," & qtdProducao & ", " & qtdReclamacao & ",'" & dataFabricacao & "','" & vencimento & "')"
            db.Execute queryInsert, dbFailOnError
        Next I
    End With
    
    'Insert 4: Insert na Tabela de imagens
    With listaAnexo
        For I = 1 To .ListItems.Count
            imagem = moverImagem(.ListItems(I), lastID, cliente, I) 'chama a função mover imagem para pergar o DirDestino da Imagem no servidor
            queryInsert = "Insert Into Imagens(ID_rnc, caminho_imagem)"
            queryInsert = queryInsert & "Values(" & lastID & ", '" & imagem & "')"
            db.Execute queryInsert, dbFailOnError
        Next I
    End With
    
    objWs.CommitTrans 'Commit caso não ocorra erro
    blnInprocess = False
    
Function_Exit:
    'Fechando banco de dados e limpando o RecordSet
    Set objWs = Nothing
    Set rds = Nothing
    db.Close
    Exit Function
    
error:
    If blnInprocess Then 'Condição que só realizado o rollback caso o erro tenha ocorrido na operação.
        objWs.Rollback 'RollBack caso ocorra erro
        mensagemErro "O envio não foi concluido, por favor procure o Administrador" & Err.Description
    End If
    
    Resume Function_Exit
    
End Function

Public Function verificaVazio(ByVal form As UserForm) As Boolean
    '
    'Verica todos os objts do formulario identificando se algum o objt obrigatório está vazio
    '

    Dim controle As Control
    
    verificaVazio = False
    
    For Each controle In form.Controls
        'Objts que não podem estar vazio
        If TypeName(controle) = "TextBox" Or TypeName(controle) = "ComboBox" Then
            'Verificando as exceções
            If controle.name = "tbObservacao" Or controle.name = "tbLote" Or controle.name = "tbNomeProduto" Or controle.name = "tbVencimento" Or controle.name = "tbQtdProducao" Or controle.name = "tbQtdReclamacao" Or controle.name = "cbTipoRnc" Then
                GoTo proximo
            Else
                If controle.Text = Empty Then
                    verificaVazio = True
                    mensagemInformacao "O campo " & controle.Tag & " deve ser preenchido !"
                End If
            End If
        ElseIf TypeName(controle) = "ListView4" Then
            If controle.ListItems.Count = 0 Then
                verificaVazio = True
                mensagemInformacao "A lista de " & controle.Tag & " não pode estar vazio !"
            End If
        End If
        If verificaVazio = True Then Exit Function
proximo:
    Next controle

End Function

Public Function moverImagem(ByVal strCaminho As String, ByVal ID As Integer, ByVal nomeCliente As String, ByVal contador As Integer) As String
    '
    'Cria uma cópia da imagem fornecida como paramêtro e cola a imagem e um novo diretório, retorno o seu diretório da cópia
    '

    Dim strCaminhoDestino As String
    
    strCaminhoDestino = getCaminho("caminhoPastaFotos") & ID & "_" & nomeCliente & "_" & "(" & contador & ")" & Mid(strCaminho, InStr(1, strCaminho, ".", 1), Len(strCaminho))
    
    FileCopy strCaminho, strCaminhoDestino
    
    moverImagem = strCaminhoDestino

End Function

Public Function CalculoValidade(Data As String, validade As Long)
    '
    'Retorna a data em "mês/ano" de acordo com o número de anos informados(String)
    '

    Dim Mes As Long, Ano As Long, Calculo As Long
    
    Ano = Year(CDate(Data))
    
    Mes = Month(CDate(Data))
    
    If validade > 12 Then
    
        Calculo = validade \ 12

        Ano = Ano + Calculo
        
    ElseIf Mes + validade > 12 Then
    
        Ano = Ano + 1
    
    End If
    
    Mes = ((Mes + validade) Mod 12)
    
    If Mes = 0 Then
        Mes = 12
    End If
    
    CalculoValidade = CStr(MonthName(Mes, True)) + "/" + CStr(Ano)

End Function

Public Function statusRnc(Optional str As String, Optional num As Integer) As Variant
    '
    'Controle de status tanto para o front quanto para o back.
    '

    str = UCase(str)

    If str = "" Then
        'Verificação do numero para retornar a String.
        Select Case num
            Case Is = 0
                statusRnc = "ABERTO"
            Case Is = 1
                statusRnc = "EM ANÁLISE"
            Case Is = 2
                statusRnc = "FECHADO"
            Case Is = 3
                statusRnc = "CANCELADO"
        End Select
    Else
        'Verificação da String para retornar o numero.
        Select Case str
            Case Is = "ABERTO"
                statusRnc = 0
            Case Is = "EM ANÁLISE"
                statusRnc = 1
            Case Is = "FECHADO"
                statusRnc = 2
            Case Is = "CANCELADO"
                statusRnc = 3
        End Select
    End If
End Function

Public Function situacaoRnc(Optional str As String, Optional num As Integer) As Variant
    '
    'Controle de situação tanto para o front quanto para o back.
    '

    str = UCase(str)
    
    If str = "" Then
        'Verificação do numero para retornar a String.
        Select Case num
            Case Is = 0
                situacaoRnc = "AGUARDANDO CONCLUSÃO"
            Case Is = 1
                situacaoRnc = "PROCEDENTE"
            Case Is = 2
                situacaoRnc = "IMPROCEDENTE"
            Case Is = 3
                situacaoRnc = "CANCELADO"
        End Select
    Else
        'Verificação da String para retornar o numero.
        Select Case str
            Case Is = "AGUARDANDO CONCLUSÃO"
                situacaoRnc = 0
            Case Is = "PROCEDENTE"
                situacaoRnc = 1
            Case Is = "IMPROCEDENTE"
                situacaoRnc = 2
            Case Is = "CANCELADO"
                situacaoRnc = 3
        End Select
    End If
End Function
