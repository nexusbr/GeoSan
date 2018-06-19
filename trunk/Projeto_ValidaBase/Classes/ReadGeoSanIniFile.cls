VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGeoSanIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private CGeoSanIniFile As Object
Private strConn As String                   'String da conexão completa com o banco de dados geográfico
Private m_arquivo As String                 'Caminho e nome completo do arquivo de configurações c:\...\CONTROLES\GEOSAN.ini, nesta variável é rescebido internamente onde ele irá procurar o arquivo .INI
Dim mProvedor As String     'Tipo de banco de dados, Oracle, SQLServer ou Postgres


'Define funções privadas desta classe para ler e gravar no arquivo de inicialização .INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Retorna a string de conexão com o banco de dados a partir das informações lidas no arquivo de inicialização .INI
Property Get StrConexao() As String
    leStringConexao
    StrConexao = strConn
End Property
'Retorna a string de conexão com o banco de dados a partir das informações lidas no arquivo de inicialização .INI
Property Get TipoBDados() As String
    leStringConexao
    TipoBDados = mProvedor
End Property
'Recebe o nome do arquivo a ser procurado
Property Let arquivo(ByVal newValue As String)
If newValue = "" Then
    Err.Raise 5
Else
    m_arquivo = newValue
End If
End Property
'Tira e encriptação da senha presente no arquivo .INI
Public Function FunDecripta(ByVal strDecripta As String) As String
    Dim IntTam As Integer
    Dim i As Integer
    Dim letra As String
    Dim nStr As String
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
            MsgBox "Código de criptografia inválido!"
            'mStrDeCriptografa = ""
            Exit Function
        End Select
        i = i + 5
    Loop
  FunDecripta = nStr
    'mStrDeCriptografa = nStr
Exit Function
End Function
'Lê os dados desejados do arquivo de inicizalização .INI
'secao = Nome da seção a qual o dado pertençe, que encontra-se entre chaves [xxxx]
'entrada = Nome do dado a ser lido, este nome encontra-se antes do sinal de =
'aquivo = Caminho e nome completo do arquivo de configurações c:\...\CONTROLES\GEOSAN.ini
Private Function ReadINI(secao As String, entrada As String, arquivo As String)
    Dim retlen As String
    Dim Ret As String
    Ret = String$(255, 0)
    retlen = GetPrivateProfileString(secao, entrada, "", Ret, Len(Ret), arquivo)
    Ret = Left$(Ret, retlen)
    ReadINI = Ret
End Function
'Obtem a string de conexão completa de acordo com o arquivo de inicialização
Private Sub leStringConexao()
    Dim mSERVIDOR As String     'Endereço IP ou nome do servidor de banco de dados
    Dim mPORTA As String        'Número da porta em que se acessa o banco de dados junto ao servidor de banco de dados
    Dim mBANCO As String        'Nome do banco de dados
    Dim mUSUARIO As String      'Nome do usuário que acessa todo o banco de dados
    Dim senha As String         'Senha encriptada do usuário que acessa o banco de dados
    Dim decriptada As String    'Senha deencriptada do usuário que acessa o banco de dados
    Dim nStr As String
    
    mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", m_arquivo)
    mPORTA = ReadINI("CONEXAO", "PORTA", m_arquivo)
    mBANCO = ReadINI("CONEXAO", "BANCO", m_arquivo)
    mUSUARIO = ReadINI("CONEXAO", "USUARIO", m_arquivo)
    senha = ReadINI("CONEXAO", "SENHA", m_arquivo)
    nStr = FunDecripta(senha)
    mProvedor = ReadINI("CONEXAO", "PROVEDOR", m_arquivo)
    Select Case mProvedor
        Case "Oracle"
            'Exemplo de conexão: Provider=OraOLEDB.Oracle.1;Password=geosan;Persist Security Info=True;User ID=geosan;Data Source=servidor07;Extended Properties
        Case "1-SQL Server 2005"
            'Exemplo de conexão: Provider=SQLOLEDB.1;Password=senha;Persist Security Info=True;User ID=usuario;Initial Catalog=nome do banco;Data Source=local
            strConn = "Provider=SQLOLEDB.1;Password=" + nStr + ";Persist Security Info=True;User ID=" + mUSUARIO + ";Initial Catalog=" + mBANCO + ";Data Source=" + mSERVIDOR
        Case "Postgres"
            'Exemplo de conexão: DRIVER={PostgreSQL Unicode}; DATABASE=banco_geosan; SERVER=localhost; PORT=5432; UID=Administrador; PWD=secreta; ByteaAsLongVarBinary=1"
            strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"
        Case Else
            MsgBox "Banco de dados incorreto, somente são aceitos SQLServer, Oracle e Postgres"
        End Select
End Sub

