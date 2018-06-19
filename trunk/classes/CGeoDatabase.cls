VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGeoDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public geoDatabase As New TeDatabase           'aqui ele deixa como herança todos os métodos do TeDatabase, quando o usuário instanciar um novo objeto do tipo CGeoDatabase
' Função para ativar a conexão com o banco de dados utilizando a biblioteca TeDatabase
'
' retorna se a conexão realizou ou não
' conexão - string de conexão com o banco de dados geográfico
' provedor - 0 - Access, 1 - SQLServer, 2 - Oracle, 3 - Firebird, 4 - Postgres (isto vem do VB6 em object browser)
' usuário - nome do usuário Terralib que se conecta ao banco geográfico
'
Public Function configura(conexao As ADODB.Connection, provedor As cAppType, usuario As String) As Boolean
    On Error GoTo Trata_Erro:
    
    If typeconnection <> POSTGRESQL Then
        'se não for Postgresss
        geoDatabase.UserName = usuario
        geoDatabase.Provider = provedor
        geoDatabase.Connection = conexao
    Else
        'implementar
    End If
    Exit Function

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      ErroUsuario.Registra "CGeoDatabase", "configura", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
   End If
End Function
