Attribute VB_Name = "MdlBancoDados"
Option Explicit
Option Private Module

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim CMD As ADODB.Command
Dim caminhoBancoDeDados As String

Public Sub conexaoAccess()

    On Error GoTo erro
    caminhoBancoDeDados = getCaminho("caminhoBancoAccess")
    
    If Dir(caminhoBancoDeDados) = vbNullString Then
        
        mensagemErro "O Banco de Dados não foi encontrado"
        
    End If
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set CMD = New ADODB.Command
    
    With cn
        .connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;"
        '.Properties("Jet OLEDB:DataBase Password=") = senha
        .Mode = adModeReadWrite
        .Open caminhoBancoDeDados
        
    End With
    
    rs.ActiveConnection = cn
    
    Exit Sub

erro:
    mensagemErro Err.Description
    
    fechaBancoDeDados
        

End Sub

Public Sub conexaoSQLServer()

    On Error GoTo erro
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set CMD = New ADODB.Command
    
    With cn
        .connectionString = "PROVIDER=SQLOLEDB;DATA SOURCE=SERVER;INITIAL CATALOG=;User Id=;Password=;"
        .Open
        
    End With
    
    rs.ActiveConnection = cn
    
    Exit Sub

erro:
    mensagemErro Err.Description
    
    fechaBancoDeDados
        

End Sub

Public Sub conexaoFireBird()

    On Error GoTo erro
    
    If Dir(caminhoBancoDeDados) = vbNullString Then
        
        mensagemErro "O Banco de Dados não foi encontrado"
    
    End If
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set CMD = New ADODB.Command
    
    With cn
        .connectionString = "DRIVER=Firebird/Interbase(r) driver;UID=;PWD=;DBNAME="
        .Mode = adModeReadWrite
        .Open
    
    End With
    
    rs.ActiveConnection = cn
    
    Exit Sub

erro:
    mensagemErro Err.Description
    
    fechaBancoDeDados


End Sub
Public Sub fechaBancoDeDados()

    On Error Resume Next
    
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
    Set CMD = Nothing
    
End Sub

Function getRecordset(ByVal sql As String, Optional ByVal cursorType As CursorTypeEnum = adOpenForwardOnly, _
    Optional ByVal lockType As LockTypeEnum = adLockReadOnly) As Recordset

    If rs.State = adStateOpen Then
    
        rs.Close
        
    End If
        
    rs.Open sql, cn, cursorType, lockType
    
    Set getRecordset = rs
    
    
End Function

Function getComando(ByVal sql As String) As Command

    Set CMD = New ADODB.Command
    
    With CMD
        .ActiveConnection = cn
        .CommandTimeout = 0
        .CommandType = adCmdText
        .CommandText = sql
    
        
    End With
    
    Set getComando = CMD
    
End Function

