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
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum cAppType
   ACCESS = 0
   SQLSERVER = 1
   ORACLE = 2
   FireBird = 3
   POSTGRESQL = 4
End Enum

Private SenhaDecripto As String

Private mvarTypeConnection As String 'local copy
Private mvarServerName As String 'local copy
Private mvarCatalog As String 'local copy
Private mvarPathFileName As String 'local copy
Private mvarServiceName As String 'local copy
Private mvarUserName As String 'local copy
Private mvarPassWord As String 'local copy

'Dim mTypeConnection As String, mServerName As String, mCatalog As String, mPathFileName As String, mServiceName As String, mPassWord As String, mUser As String

Public Property Let ServiceName(ByVal vData As String)
    mvarServiceName = vData
End Property

Public Property Get ServiceName() As String
    ServiceName = mvarServiceName
End Property

Public Property Let PathFileName(ByVal vData As String)
    mvarPathFileName = vData
End Property

Public Property Get PathFileName() As String
    PathFileName = mvarPathFileName
End Property

Public Property Let Catalog(ByVal vData As String)
    mvarCatalog = vData
End Property

Public Property Get Catalog() As String
    Catalog = mvarCatalog
End Property

Public Property Let ServerName(ByVal vData As String)
    mvarServerName = vData
End Property

Public Property Get ServerName() As String
    ServerName = mvarServerName
End Property

Public Property Let TypeConnection(ByVal vData As String)
    mvarTypeConnection = vData
End Property

Public Property Get TypeConnection() As String
    TypeConnection = mvarTypeConnection
End Property

Public Property Let UserName(ByVal vData As String)
    mvarUserName = vData
End Property

Public Property Get UserName() As String
    UserName = mvarUserName
End Property

Public Property Let PassWord(ByVal vData As String)
    mvarPassWord = vData
End Property

Public Property Get PassWord() As String
    PassWord = mvarPassWord
End Property

Public Function appNewRegistry(appName As String, Conn As ADODB.Connection, Optional TypeConn As cAppType) As Boolean

'FUNÇÃO PARA CRIAR NOVO ARQUIVO DE INFORMAÇÕES SOBRE CONEXÃO DE BANCO DE DADOS

    appNewRegistry = FrmConnection.Init(appName, Conn, TypeConn) ' CRIA NOVA CONEXÃO CASO NÃO EXISTA
    
   ' BUSCA NO ARQUIVO GEOSAN.INI AS INFORMAÇÕES DE CONEXÃO
   mPROVEDOR = Mid(ReadINI("CONEXAO", "PROVEDOR", App.Path & "\GEOSAN.ini"), 1, 1) ' PEGA SÓ O NÚMERO DA CONEXÃO
   mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.Path & "\GEOSAN.ini")
   mSERVICO = ReadINI("CONEXAO", "SERVIÇO", App.Path & "\GEOSAN.ini")
   mPORTA = ReadINI("CONEXAO", "PORTA", App.Path & "\GEOSAN.ini")
   mBANCO = ReadINI("CONEXAO", "BANCO", App.Path & "\GEOSAN.ini")
   mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.Path & "\GEOSAN.ini")
   funDecripta (ReadINI("CONEXAO", "SENHA", App.Path & "\GEOSAN.ini"))
   mSENHA = SenhaDecripto
   
   If IsNumeric(mPROVEDOR) = True Then
      
      TypeConnection = mPROVEDOR
   
   Else
      appNewRegistry = False
      MsgBox "Não foi possível estabelecer a conexão." & Chr(13) & Chr(13) & "Não há provedor definido.", vbExclamation, ""
      Exit Function
   End If
    
    
End Function

Public Function appGetRegistry(appName As String, MyConn As ADODB.Connection, Optional TypeConn As cAppType) As Boolean
On Error GoTo Trata_Erro

   Set MyConn = New ADODB.Connection
   Dim prov As String
   ' BUSCA NO ARQUIVO GEOSAN.INI AS INFORMAÇÕES DE CONEXÃO
   mPROVEDOR = Mid(ReadINI("CONEXAO", "PROVEDOR", App.Path & "\GEOSAN.ini"), 1, 1) ' PEGA SÓ O NÚMERO DA CONEXÃO
   prov = Mid(ReadINI("CONEXAO", "PROVEDOR", App.Path & "\GEOSAN.ini"), 3, 20) ' PEGA SÓ O NÚMERO DA CONEXÃO
   mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.Path & "\GEOSAN.ini")
   mSERVICO = ReadINI("CONEXAO", "SERVIÇO", App.Path & "\GEOSAN.ini")
   mPORTA = ReadINI("CONEXAO", "PORTA", App.Path & "\GEOSAN.ini")
   mBANCO = ReadINI("CONEXAO", "BANCO", App.Path & "\GEOSAN.ini")
   mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.Path & "\GEOSAN.ini")
   funDecripta (ReadINI("CONEXAO", "SENHA", App.Path & "\GEOSAN.ini"))
   
   mSENHA = SenhaDecripto
   
   If IsNumeric(mPROVEDOR) = True Then
      
      TypeConn = mPROVEDOR
   
   Else
      appGetRegistry = False
      'MsgBox "Não foi possível estabelecer a conexão." & Chr(13) & Chr(13) & "Não há provedor definido.", vbExclamation, ""
      Exit Function
   End If
   If prov = "SQL Server 2008" Then
   prov = "SQLNCLI10.1"
   End If
     If prov = "SQL Server 2005" Then
    prov = "SQLOLEDB.1"
   End If
   
   
   Dim st As String
   
   Select Case TypeConn
   Case 0
         stc = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mBANCO & ";Persist Security Info=False"
      Case 1
         stc = "Provider=" + prov + ";Persist Security Info=True;Data Source=" & mSERVIDOR & ";User ID=" & mUSUARIO & ";Password=" & mSENHA & ";Initial Catalog=" & mBANCO
      Case 2
         stc = "Provider=OraOLEDB.Oracle.1;Password=" & mSENHA & ";Persist Security Info=True;User ID=" & mUSUARIO & ";Data Source=" & mSERVICO
      Case 4
         stc = "Provider=PostgreSQL.1;Data Source=" & mSERVIDOR & ";User ID=" & mUSUARIO & ";Password=" & mSENHA & ";location=" & mBANCO
   End Select
   
   
   
   
   MyConn.ConnectionString = stc
   MyConn.Open
   appGetRegistry = True
        
   Exit Function

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
    
      Resume Next
      
   Else
      
      MsgBox "Não foi possível estabelecer a conexão" & Err.Number & " - " & Err.Description, vbExclamation
      appGetRegistry = False
      
   End If
   
End Function





Public Function funDecripta(ByVal strDecripta As String)


    Dim IntTam As Integer
    Dim nStr As String
    Dim i As Integer
    Dim letra As String
    IntTam = Len(strDecripta)
    nStr = ""
    
    'desconsidera os os numeros de HH-MM-SS
    strDecripta = Mid(strDecripta, 6, 5) & Mid(strDecripta, 16, 5) & Mid(strDecripta, 26, 5) & _
                  Mid(strDecripta, 36, 5) & Mid(strDecripta, 46, 5) & Mid(strDecripta, 56, 200)
    
    i = 1
    Do While Not i = IntTam - 29
        letra = Mid(strDecripta, i, 5)
        Select Case letra
        Case "14334"
            nStr = nStr & "a"
        Case "14212"
            nStr = nStr & "A"
        Case "24334"
            nStr = nStr & "á"
        Case "24134"
            nStr = nStr & "â"
        Case "24234"
            nStr = nStr & "ã"
        Case "24314"
            nStr = nStr & "à"
        Case "24324"
            nStr = nStr & "b"
        Case "14223"
            nStr = nStr & "B"
        Case "11211"
            nStr = nStr & "ç"
        Case "11311"
            nStr = nStr & "Ç"
        Case "13334"
            nStr = nStr & "c"
        Case "14324"
            nStr = nStr & "C"
        Case "24344"
            nStr = nStr & "d"
        Case "14444"
            nStr = nStr & "D"
        Case "12314"
            nStr = nStr & "e"
        Case "21111"
            nStr = nStr & "E"
        Case "24321"
            nStr = nStr & "é"
        Case "32314"
            nStr = nStr & "ê"
        Case "31314"
            nStr = nStr & "f"
        Case "21311"
            nStr = nStr & "F"
        Case "32134"
            nStr = nStr & "g"
        Case "21341"
            nStr = nStr & "G"
        Case "31324"
            nStr = nStr & "h"
        Case "22111"
            nStr = nStr & "H"
        Case "32124"
            nStr = nStr & "i"
        Case "21112"
            nStr = nStr & "I"
        Case "31334"
            nStr = nStr & "í"
        Case "32333"
            nStr = nStr & "ì"
        Case "11314"
            nStr = nStr & "j"
        Case "23122"
            nStr = nStr & "J"
        Case "33134"
            nStr = nStr & "k"
        Case "23411"
            nStr = nStr & "K"
        Case "33314"
            nStr = nStr & "l"
        Case "32222"
            nStr = nStr & "L"
        Case "43423"
            nStr = nStr & "m"
        Case "32111"
            nStr = nStr & "M"
        Case "42423"
            nStr = nStr & "n"
        Case "33221"
            nStr = nStr & "N"
        Case "43234"
            nStr = nStr & "o"
        Case "33233"
            nStr = nStr & "O"
        Case "42444"
            nStr = nStr & "ô"
        Case "43223"
            nStr = nStr & "õ"
        Case "42433"
            nStr = nStr & "ò"
        Case "43231"
            nStr = nStr & "ó"
        Case "22223"
            nStr = nStr & "p"
        Case "33444"
            nStr = nStr & "P"
        Case "43233"
            nStr = nStr & "q"
        Case "34442"
            nStr = nStr & "Q"
        Case "43421"
            nStr = nStr & "r"
        Case "34332"
            nStr = nStr & "R"
        Case "13443"
            nStr = nStr & "s"
        Case "34222"
            nStr = nStr & "S"
        Case "44444"
            nStr = nStr & "t"
        Case "34112"
            nStr = nStr & "T"
        Case "13444"
            nStr = nStr & "u"
        Case "41311"
            nStr = nStr & "U"
        Case "11111"
            nStr = nStr & "ú"
        Case "13243"
            nStr = nStr & "ù"
        Case "11115"
            nStr = nStr & "û"
        Case "13241"
            nStr = nStr & "v"
        Case "41222"
            nStr = nStr & "V"
        Case "12443"
            nStr = nStr & "x"
        Case "41133"
            nStr = nStr & "X"
        Case "13244"
            nStr = nStr & "y"
        Case "42231"
            nStr = nStr & "Y"
        Case "13441"
            nStr = nStr & "w"
        Case "42222"
            nStr = nStr & "W"
        Case "11313"
            nStr = nStr & "z"
        Case "42213"
            nStr = nStr & "Z"
        Case "11312"
            nStr = nStr & "@"
        Case "11114"
            nStr = nStr & "%"
        Case "12341"
            nStr = nStr & "&"
        Case "13343"
            nStr = nStr & "*"
        Case "12342"
            nStr = nStr & "("
        Case "13344"
            nStr = nStr & ")"
        Case "12333"
            nStr = nStr & "$"
        Case "23334"
            nStr = nStr & "!"
        Case "13331"
            nStr = nStr & "#"
        Case "21242"
            nStr = nStr & "?"
        Case "22313"
            nStr = nStr & "1"
        Case "23424"
            nStr = nStr & "2"
        Case "24131"
            nStr = nStr & "3"
        Case "41414"
            nStr = nStr & "4"
        Case "22314"
            nStr = nStr & "5"
        Case "23423"
            nStr = nStr & "6"
        Case "44134"
            nStr = nStr & "7"
        Case "21241"
            nStr = nStr & "8"
        Case "22312"
            nStr = nStr & "9"
        Case "23231"
            nStr = nStr & "0"
        Case "34123"
            nStr = nStr & " "
        Case "14121"
            nStr = nStr & "_"
        Case "14144"
            nStr = nStr & "/"
        Case "12131"
            nStr = nStr & "\"
        Case "12124"
            nStr = nStr & "-"
        Case "21421"
            nStr = nStr & ";"
        Case "21321"
            nStr = nStr & ":"
        Case "14431"
            nStr = nStr & ","
        Case "13421"
            nStr = nStr & "."
        Case "11213"
            nStr = nStr & "+"
        Case "11212"
            nStr = nStr & "="
        
        Case Else
            'MsgBox "Código de criptografia inválido!"
            'mStrDeCriptografa = ""
            Exit Function
        End Select
        i = i + 5
    Loop
    SenhaDecripto = nStr
    'mStrDeCriptografa = nStr
Exit Function
End Function




