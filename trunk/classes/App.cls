VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''Option Explicit
'''
'''Public Enum cAppType
'''   Access = 0
'''   SqlServer = 1
'''   Oracle = 2
'''   FireBird = 3
'''End Enum
'''
'''Private mvarTypeConnection As String 'local copy
'''Private mvarServerName As String 'local copy
'''Private mvarCatalog As String 'local copy
'''Private mvarPathFileName As String 'local copy
'''Private mvarServiceName As String 'local copy
'''Private mvarUserName As String 'local copy
'''Private mvarPassWord As String 'local copy
'''
''''Dim mTypeConnection As String, mServerName As String, mCatalog As String, mPathFileName As String, mServiceName As String, mPassWord As String, mUser As String
'''
'''Public Property Let ServiceName(ByVal vData As String)
'''
'''    mvarServiceName = vData
'''
'''End Property
'''
'''Public Property Get ServiceName() As String
'''
'''    ServiceName = mvarServiceName
'''
'''End Property
'''
'''Public Property Let PathFileName(ByVal vData As String)
'''
'''    mvarPathFileName = vData
'''
'''End Property
'''
'''Public Property Get PathFileName() As String
'''
'''    PathFileName = mvarPathFileName
'''
'''End Property
'''
'''Public Property Let Catalog(ByVal vData As String)
'''
'''    mvarCatalog = vData
'''
'''End Property
'''
'''Public Property Get Catalog() As String
'''
'''    Catalog = mvarCatalog
'''
'''End Property
'''
'''Public Property Let ServerName(ByVal vData As String)
'''
'''    mvarServerName = vData
'''
'''End Property
'''
'''Public Property Get ServerName() As String
'''
'''    ServerName = mvarServerName
'''
'''End Property
'''
'''Public Property Let TypeConnection(ByVal vData As String)
'''
'''    mvarTypeConnection = vData
'''
'''End Property
'''
'''Public Property Get TypeConnection() As String
'''
'''    TypeConnection = mvarTypeConnection
'''
'''End Property
'''
'''Public Property Let UserName(ByVal vData As String)
'''
'''    mvarUserName = vData
'''
'''End Property
'''
'''Public Property Get UserName() As String
'''
'''    UserName = mvarUserName
'''
'''End Property
'''
'''Public Property Let PassWord(ByVal vData As String)
'''
'''    mvarPassWord = vData
'''
'''End Property
'''
'''Public Property Get PassWord() As String
'''
'''    PassWord = mvarPassWord
'''
'''End Property
'''
'''Public Function appNewRegistry(appName As String, Conn As ADODB.Connection, Optional TypeConn As cAppType) As Boolean
'''
''''FUN��O PARA CRIAR NOVO ARQUIVO DE INFORMA��ES SOBRE CONEX�O DE BANCO DE DADOS
'''
'''    appNewRegistry = FrmConnection.Init(appName, Conn, TypeConn) ' CRIA NOVA CONEX�O CASO N�O EXISTA
'''
'''    Open App.path & "\" & appName & ".cfg" For Input As #1
'''        Input #1, mvarTypeConnection
'''        Input #1, mvarServerName
'''        Input #1, mvarCatalog
'''        Input #1, mvarPathFileName
'''        Input #1, mvarServiceName
'''        Input #1, mvarUserName
'''        Input #1, mvarPassWord
'''    Close #1
'''
'''    TypeConnection = mvarTypeConnection
'''    ServerName = mvarServerName
'''    Catalog = mvarCatalog
'''    PathFileName = mvarPathFileName
'''    ServiceName = mvarServiceName
'''    UserName = mvarUserName
'''    PassWord = mvarPassWord
'''
'''End Function
'''
'''Public Function appGetRegistry(appName As String, MyConn As ADODB.Connection, Optional TypeConn As cAppType) As Boolean
'''On Error GoTo Trata_Erro
'''
'''   Set MyConn = New ADODB.Connection
'''
'''   'Open App.Path & "\" & appName & ".cfg" For Input As #1
'''   Open App.path & "\" & appName & ".cfg" For Input As #1
'''      Input #1, mvarTypeConnection
'''      Input #1, mvarServerName
'''      Input #1, mvarCatalog
'''      Input #1, mvarPathFileName
'''      Input #1, mvarServiceName
'''      Input #1, mvarUserName
'''      Input #1, mvarPassWord
'''   Close #1
'''
'''
'''
'''   TypeConnection = mvarTypeConnection
'''   ServerName = mvarServerName
'''   Catalog = mvarCatalog
'''   PathFileName = mvarPathFileName
'''   ServiceName = mvarServiceName
'''   UserName = mvarUserName
'''   PassWord = mvarPassWord
'''
'''   SELECT Case TypeConnection
'''      Case Access
'''         MyConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False"
'''         TypeConn = TypeConnection
'''         appGetRegistry = True
'''      Case SqlServer
'''         MyConn.Open "Provider=SQLOLEDB.1;Persist Security Info=True;Data Source=" & ServerName & ";User ID=" & UserName & ";Password=" & PassWord & ";Initial Catalog=" & Catalog
'''         TypeConn = TypeConnection
'''         appGetRegistry = True
'''      Case Oracle
'''         MyConn.Open "Provider=OraOLEDB.Oracle.1;Password=" & PassWord & ";Persist Security Info=True;User ID=" & UserName & ";Data Source=" & ServiceName
'''         TypeConn = TypeConnection
'''         appGetRegistry = True
'''   End SELECT
'''
'''   Exit Function
'''
'''
'''Trata_Erro:
''''Error_appGetRegistry:
'''    Close #1
'''    MsgBox "N�o foi poss�vel estabelecer a conex�o", vbExclamation
'''
'''
''''    If Err.Number = 0 Or Err.Number = 20 Then
''''        Resume Next
''''    ElseIf Err.Number = 53 Then 'ARQUIVO DE REGISTRO .CFG N�O FOI ENCONTRADO
''''        'MsgBox "Bem vindo ao GeoSan!",
''''        MsgBox "O arquivo para conex�o autom�tica com banco de dados ser� criado.", vbInformation, "Primeiro acesso"
''''    Else
''''        Close #1
''''        Open App.Path & "\GeoSanLog.txt" For Output As #1
''''        Print #1, Now & " " & strUser & " " & Versao_Geo & " - NexusConnection.DLL - Public Function appGetRegistry - " & Err.Number & " - " & Err.Description
''''        Close #1
''''        MsgBox "Um posss�vel erro foi identificado:" & Chr(13) & Chr(13) & "N�o foi poss�vel estabelecer a conex�o" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informa��es desta ocorrencia.", vbInformation
''''    End If
'''
'''End Function
