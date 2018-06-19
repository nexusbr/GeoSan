VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmConnection 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Configuração de Conexão"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4050
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameOracle 
      BackColor       =   &H8000000E&
      Height          =   1395
      Left            =   150
      TabIndex        =   25
      Top             =   1185
      Visible         =   0   'False
      Width           =   3705
      Begin VB.TextBox txtOracleServico 
         Height          =   285
         Left            =   810
         TabIndex        =   26
         Top             =   585
         Width           =   2715
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Serviço:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   90
         TabIndex        =   27
         Top             =   585
         Width           =   855
      End
   End
   Begin VB.Frame frameAcs 
      BackColor       =   &H8000000E&
      Height          =   1395
      Left            =   150
      TabIndex        =   21
      Top             =   1185
      Visible         =   0   'False
      Width           =   3705
      Begin VB.CommandButton cmdFile 
         Caption         =   "..."
         Height          =   255
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtAcsBanco 
         Height          =   285
         Left            =   810
         TabIndex        =   22
         Top             =   600
         Width           =   2355
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Banco:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   90
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame frameSql 
      BackColor       =   &H8000000E&
      Height          =   1395
      Left            =   150
      TabIndex        =   16
      Top             =   1185
      Visible         =   0   'False
      Width           =   3705
      Begin VB.TextBox txtSqlBanco 
         Height          =   285
         Left            =   840
         TabIndex        =   18
         Top             =   375
         Width           =   2715
      End
      Begin VB.TextBox txtSqlServidor 
         Height          =   285
         Left            =   840
         TabIndex        =   17
         Top             =   825
         Width           =   2715
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Banco:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Servidor:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   825
         Width           =   855
      End
   End
   Begin VB.Frame framePostgres 
      BackColor       =   &H8000000E&
      Height          =   1395
      Left            =   150
      TabIndex        =   9
      Top             =   1185
      Visible         =   0   'False
      Width           =   3705
      Begin VB.TextBox txtPortaPostgres 
         Height          =   285
         Left            =   2580
         TabIndex        =   12
         Text            =   "5432"
         Top             =   615
         Width           =   960
      End
      Begin VB.TextBox txtServidorPostgres 
         Height          =   285
         Left            =   825
         TabIndex        =   11
         Text            =   "localhost"
         Top             =   285
         Width           =   2715
      End
      Begin VB.TextBox txtBancoPostgres 
         Height          =   285
         Left            =   825
         TabIndex        =   10
         Text            =   "postgres"
         Top             =   960
         Width           =   2715
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Porta:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1890
         TabIndex        =   15
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Servidor:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   105
         TabIndex        =   14
         Top             =   285
         Width           =   690
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Banco:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   645
      End
   End
   Begin VB.Frame fraSenha 
      BackColor       =   &H8000000E&
      Height          =   1215
      Left            =   150
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   3705
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   720
         Width           =   2715
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   270
         Width           =   2715
      End
      Begin VB.Label lblSenha 
         BackColor       =   &H8000000E&
         Caption         =   "Senha:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblUser 
         BackColor       =   &H8000000E&
         Caption         =   "Usuário:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Sair"
      Height          =   345
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4020
      Width           =   885
   End
   Begin VB.Frame FrameOption 
      BackColor       =   &H8000000E&
      Caption         =   "Tipo de Conexão"
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   3705
      Begin VB.ComboBox cboTPConexao 
         Height          =   315
         ItemData        =   "FrmConnection.frx":0000
         Left            =   345
         List            =   "FrmConnection.frx":0013
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   390
         Width           =   2940
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   345
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4020
      Width           =   885
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strEncripta As String
Public str As String

Private NovaStr As String
Private senhaCripto As String

Private Registrou As Boolean, apName As String, MyConn As New ADODB.Connection, Tconn As String
Dim providerSql As String
Dim strin As String
   

Public Function Init(appName As String, Conn As ADODB.Connection, Optional TypeConn As cAppType) As Boolean
    
    Registrou = False
    apName = appName 'NOME DO APLICATIVO QUE CHAMOU A FUNÇÃO
    Me.Show vbModal

    If Tconn <> "" Then
        Init = Registrou
        TypeConn = Tconn
    End If
    
    Set Conn = MyConn
    Set MyConn = Nothing

End Function

Private Sub cboTPConexao_Click()
   
   strin = Me.cboTPConexao.Text
   Me.fraSenha.Visible = False
   frameSql.Visible = False
   frameOracle.Visible = False
   frameAcs.Visible = False
   framePostgres.Visible = False
   
   If strin = "Access" Then
      
      frameAcs.Visible = True
      Me.txtAcsBanco.SetFocus
      Me.txtUser.TabIndex = Me.txtAcsBanco.TabIndex + 1
      Me.txtSenha.TabIndex = Me.txtUser.TabIndex + 1
   
   ElseIf strin = "SQL Server 2005" Then

      frameSql.Visible = True
      Me.fraSenha.Visible = True
      Me.txtSqlBanco.SetFocus
      Me.txtSqlServidor.TabIndex = Me.txtSqlBanco.TabIndex + 1
      Me.txtUser.TabIndex = Me.txtSqlServidor.TabIndex + 1
      Me.txtSenha.TabIndex = Me.txtUser.TabIndex + 1
      providerSql = "SQLOLEDB.1"
      
      ElseIf strin = "SQL Server 2008" Then
      frameSql.Visible = True
      Me.fraSenha.Visible = True
      Me.txtSqlBanco.SetFocus
      Me.txtSqlServidor.TabIndex = Me.txtSqlBanco.TabIndex + 1
      Me.txtUser.TabIndex = Me.txtSqlServidor.TabIndex + 1
      Me.txtSenha.TabIndex = Me.txtUser.TabIndex + 1
      providerSql = "SQLNCLI10.1"
      
   ElseIf strin = "Oracle" Then
   
      frameOracle.Visible = True
      Me.fraSenha.Visible = True
      Me.txtOracleServico.SetFocus
      Me.txtUser.TabIndex = Me.txtOracleServico.TabIndex + 1
      Me.txtSenha.TabIndex = Me.txtUser.TabIndex + 1
   
   ElseIf strin = "PostgreSQL" Then
   
      Me.framePostgres.Visible = True
      Me.fraSenha.Visible = True
      Me.txtServidorPostgres.SetFocus
      Me.txtPortaPostgres.TabIndex = Me.txtServidorPostgres.TabIndex + 1
      Me.txtBancoPostgres.TabIndex = Me.txtPortaPostgres.TabIndex + 1
      Me.txtUser.TabIndex = Me.txtBancoPostgres.TabIndex + 1
      Me.txtSenha.TabIndex = Me.txtUser.TabIndex + 1
      
   End If
   
End Sub

Private Sub cmdClose_Click()
   
   Unload Me

End Sub

Private Sub cmdFile_Click()

   CommonDialog1.FileName = ""
   txtAcsBanco = CommonDialog1.FileName
   CommonDialog1.Filter = "MDB(*.mdb)|*.mdb|"
   CommonDialog1.DialogTitle = "Abrir Arquivo"
   CommonDialog1.ShowOpen
   If CommonDialog1.FileName <> "" Then
      txtAcsBanco = CommonDialog1.FileName
   End If
   
End Sub

Private Sub cmdOk_Click()
On Error GoTo Trata_Erro
   
   mPROVEDOR = ""
   mSERVIDOR = ""
   mSERVICO = ""
   mPORTA = ""
   mBANCO = ""
   mUSUARIO = ""
   mSENHA = ""
   
   ' CADA TIPO DE BANCO DE DADOS REQUER UMA QUANTIDADE DE PARÂMETROS DE CONEXAO
   ' ORACLE = SERVIÇO,USUARIO,SENHA
   ' ACCESS = BANCO
   ' SQLSERVER = BANCO,SERVIDOR,USUARIO,SENHA
   ' POSTGRES = SERVIDOR,PORTA,BANCO,USUARIO,SENHA
   
   Dim strin As String
   strin = Me.cboTPConexao.Text
   
   If strin = "Access" Then
      If Trim(txtAcsBanco.Text) <> "" Then
         mBANCO = Me.txtAcsBanco.Text
      
      Else
         exibeMensagem
         Exit Sub
      End If
   
   ElseIf strin = "SQL Server 2005" Then
   
      If Trim(txtSqlServidor.Text) <> "" And Trim(txtSqlBanco.Text) <> "" And txtUser.Text <> "" And txtSenha.Text <> "" Then
      
          mPROVEDOR = "1"
          mSERVIDOR = Me.txtSqlServidor.Text
          mBANCO = Me.txtSqlBanco.Text
          mUSUARIO = Me.txtUser.Text
          mSENHA = Me.txtSenha.Text
      
      Else
         exibeMensagem
         Exit Sub
      End If
      ElseIf strin = "SQL Server 2008" Then
   
      If Trim(txtSqlServidor.Text) <> "" And Trim(txtSqlBanco.Text) <> "" And txtUser.Text <> "" And txtSenha.Text <> "" Then
      
          mPROVEDOR = "1"
          mSERVIDOR = Me.txtSqlServidor.Text
          mBANCO = Me.txtSqlBanco.Text
          mUSUARIO = Me.txtUser.Text
          mSENHA = Me.txtSenha.Text
      
      Else
         exibeMensagem
         Exit Sub
      End If
   
   ElseIf strin = "Oracle" Then
   
      If txtOracleServico <> "" And txtUser.Text <> "" And txtSenha.Text <> "" Then
         
          mPROVEDOR = 2
          mSERVICO = Me.txtOracleServico.Text
          mUSUARIO = Me.txtUser.Text
          mSENHA = Me.txtSenha.Text

      Else
         exibeMensagem
         Exit Sub
      End If
   
   ElseIf strin = "PostgreSQL" Then
   
      If Me.txtServidorPostgres.Text <> "" And Me.txtBancoPostgres.Text <> "" And Me.txtPortaPostgres.Text <> "" And txtUser.Text <> "" And txtSenha.Text <> "" Then
                                
         mPROVEDOR = "4"
         mSERVIDOR = Me.txtServidorPostgres.Text
         mPORTA = Me.txtPortaPostgres.Text
         mBANCO = Me.txtBancoPostgres.Text
         mUSUARIO = Me.txtUser.Text
         mSENHA = Me.txtSenha.Text
         
      Else
         exibeMensagem
         Exit Sub
      End If
   
   End If
   
   Me.MousePointer = vbHourglass
   
   ' FAZ A TENTATIVA DE CONEXÃO USANDO AS VARIÁVEIS CARREGADAS
   If strin = "Access" Then
      stc = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mBANCO & ";Persist Security Info=False"
   ElseIf strin = "SQL Server 2005" Then
      stc = "Provider= " + providerSql + ";Persist Security Info=True;Data Source=" & mSERVIDOR & ";User ID=" & mUSUARIO & ";Password=" & mSENHA & ";Initial Catalog=" & mBANCO
   ElseIf strin = "SQL Server 2008" Then
      stc = "Provider= " + providerSql + ";Persist Security Info=True;Data Source=" & mSERVIDOR & ";User ID=" & mUSUARIO & ";Password=" & mSENHA & ";Initial Catalog=" & mBANCO
   ElseIf strin = "Oracle" Then
      stc = "Provider=OraOLEDB.Oracle.1;Password=" & mSENHA & ";Persist Security Info=True;User ID=" & mUSUARIO & ";Data Source=" & mSERVICO
   ElseIf strin = "PostgreSQL" Then
      stc = "Provider=PostgreSQL.1;Data Source=" & mSERVIDOR & ";User ID=" & mUSUARIO & ";Password=" & mSENHA & ";location=" & mBANCO
   End If
               
   
   MyConn.Open stc
   Tconn = mPROVEDOR
               
   ' SE A CONEXAO FOI BEM SUCEDIDA, GRAVA AS INFORMAÇÕES NO GEOSAN.INI
   
   Call WriteINI("CONEXAO", "PROVEDOR", mPROVEDOR & "-" & strin, App.Path & "\GEOSAN.INI")
   Call WriteINI("CONEXAO", "SERVIDOR", mSERVIDOR, App.Path & "\GEOSAN.INI")
   Call WriteINI("CONEXAO", "SERVIÇO", mSERVICO, App.Path & "\GEOSAN.INI")
   Call WriteINI("CONEXAO", "PORTA", mPORTA, App.Path & "\GEOSAN.INI")
   Call WriteINI("CONEXAO", "BANCO", mBANCO, App.Path & "\GEOSAN.INI")
   Call WriteINI("CONEXAO", "USUARIO", mUSUARIO, App.Path & "\GEOSAN.INI")
   funEncripta (txtSenha.Text)
   funMontaStr (NovaStr)
   Call WriteINI("CONEXAO", "SENHA", senhaCripto, App.Path & "\GEOSAN.INI")
      
   Registrou = True
   Unload Me
    
Trata_Erro:
    Me.MousePointer = vbDefault
    Close #1
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    
    ElseIf Err.Number = 3705 Then
        MyConn.Close
        Resume
    
    ElseIf Trim(Tconn) = "" Then ' SIGNIFICA QUE NÃO CONSEGUIU CONECTAR NA BASE SELECIONADA E O ERRO FOI DE CONEXÃO
            
        MsgBox "Não foi possível estabelecer a conexão " & strin & " por motivo de: " & Chr(13) & Chr(13) & Err.Number & " - " & Err.Description, vbInformation
        'MsgBox Err.Number & " - " & Err.Description
    
    ElseIf Err.Number = 55 Then
        Close #1
        Resume
    
    ElseIf Err.Number = 75 Then 'ERRO DE ACESSO AO ARQUIVO
        MsgBox "Não foi possível gravar o arquivo " & apName & ".cfg na pasta '" & App.Path & "'." & Chr(13) & Chr(13) & "É necessário que o usuário possua permissão para gravar arquivos nesta pasta.", vbExclamation, "Erro de acesso"
    Else
        Close #1
        Open App.Path & "\GeoSanLog.txt" For Output As #1
        Print #1, Now & " - NexusConnection.DLL - frmConnection - cmdOK_Click - " & Err.Number & " - " & Err.Description
        Close #1
        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & "Não foi possível estabelecer a conexão" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrencia.", vbInformation
    End If

End Sub

Private Sub exibeMensagem()
   MsgBox "Todos os campos são obrigatórios", vbExclamation, ""
End Sub

Private Sub cmdSelecionar_Click()

   CommonDialog1.FileName = ""
   txtAcsBanco = CommonDialog1.FileName
   CommonDialog1.Filter = "MDB(*.mdb)|*.mdb|"
   CommonDialog1.DialogTitle = "Abrir Arquivo"
   CommonDialog1.ShowOpen
   If CommonDialog1.FileName <> "" Then
      txtAcsBanco = CommonDialog1.FileName
   End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
   End If
   
End Sub

Private Sub Form_Load()
   'Click
End Sub


Public Function funEncripta(ByVal strEncripta As String)
    Dim IntTam As Integer
    Dim nStr As String
    Dim i As Integer
    Dim letra As String
    IntTam = Len(strEncripta)
    nStr = ""
    'strEncripta = LCase(strEncripta)
    For i = 1 To IntTam
        letra = Mid(strEncripta, i, 1)
        Select Case letra
        Case "a"
            nStr = nStr & "14334"
        Case "A"
            nStr = nStr & "14212"
        Case "á"
            nStr = nStr & "24334"
        Case "â"
            nStr = nStr & "24134"
        Case "ã"
            nStr = nStr & "24234"
        Case "à"
            nStr = nStr & "24314"
        Case "b"
            nStr = nStr & "24324"
        Case "B"
            nStr = nStr & "14223"
        Case "ç"
            nStr = nStr & "11211"
        Case "Ç"
            nStr = nStr & "11311"
        Case "c"
            nStr = nStr & "13334"
        Case "C"
            nStr = nStr & "14324"
        Case "d"
            nStr = nStr & "24344"
        Case "D"
            nStr = nStr & "14444"
        Case "e"
            nStr = nStr & "12314"
        Case "E"
            nStr = nStr & "21111"
        Case "é"
            nStr = nStr & "24321"
        Case "ê"
            nStr = nStr & "32314"
        Case "f"
            nStr = nStr & "31314"
        Case "F"
            nStr = nStr & "21311"
        Case "g"
            nStr = nStr & "32134"
        Case "G"
            nStr = nStr & "21341"
        Case "h"
            nStr = nStr & "31324"
        Case "H"
            nStr = nStr & "22111"
        Case "i"
            nStr = nStr & "32124"
        Case "I"
            nStr = nStr & "21112"
        Case "í"
            nStr = nStr & "31334"
        Case "ì"
            nStr = nStr & "32333"
        Case "j"
            nStr = nStr & "11314"
        Case "J"
            nStr = nStr & "23122"
        Case "k"
            nStr = nStr & "33134"
        Case "K"
            nStr = nStr & "23411"
        Case "l"
            nStr = nStr & "33314"
        Case "L"
            nStr = nStr & "32222"
        Case "m"
            nStr = nStr & "43423"
        Case "M"
            nStr = nStr & "32111"
        Case "n"
            nStr = nStr & "42423"
        Case "N"
            nStr = nStr & "33221"
        Case "o"
            nStr = nStr & "43234"
        Case "O"
            nStr = nStr & "33233"
        Case "ô"
            nStr = nStr & "42444"
        Case "õ"
            nStr = nStr & "43223"
        Case "ò"
            nStr = nStr & "42433"
        Case "ó"
            nStr = nStr & "43231"
        Case "p"
            nStr = nStr & "22223"
        Case "P"
            nStr = nStr & "33444"
        Case "q"
            nStr = nStr & "43233"
        Case "Q"
            nStr = nStr & "34442"
        Case "r"
            nStr = nStr & "43421"
        Case "R"
            nStr = nStr & "34332"
        Case "s"
            nStr = nStr & "13443"
        Case "S"
            nStr = nStr & "34222"
        Case "t"
            nStr = nStr & "44444"
        Case "T"
            nStr = nStr & "34112"
        Case "u"
            nStr = nStr & "13444"
        Case "U"
            nStr = nStr & "41311"
        Case "ú"
            nStr = nStr & "11111"
        Case "ù"
            nStr = nStr & "13243"
        Case "û"
            nStr = nStr & "11115"
        Case "v"
            nStr = nStr & "13241"
        Case "V"
            nStr = nStr & "41222"
        Case "x"
            nStr = nStr & "12443"
        Case "X"
            nStr = nStr & "41133"
        Case "y"
            nStr = nStr & "13244"
        Case "Y"
            nStr = nStr & "42231"
        Case "w"
            nStr = nStr & "13441"
        Case "W"
            nStr = nStr & "42222"
        Case "z"
            nStr = nStr & "11313"
        Case "Z"
            nStr = nStr & "42213"
        Case "@"
            nStr = nStr & "11312"
        Case "%"
            nStr = nStr & "11114"
        Case "&"
            nStr = nStr & "12341"
        Case "*"
            nStr = nStr & "13343"
        Case "("
            nStr = nStr & "12342"
        Case ")"
            nStr = nStr & "13344"
        Case "$"
            nStr = nStr & "12333"
        Case "!"
            nStr = nStr & "23334"
        Case "#"
            nStr = nStr & "13331"
        Case "?"
            nStr = nStr & "21242"
        Case "1"
            nStr = nStr & "22313"
        Case "2"
            nStr = nStr & "23424"
        Case "3"
            nStr = nStr & "24131"
        Case "4"
            nStr = nStr & "41414"
        Case "5"
            nStr = nStr & "22314"
        Case "6"
            nStr = nStr & "23423"
        Case "7"
            nStr = nStr & "44134"
        Case "8"
            nStr = nStr & "21241"
        Case "9"
            nStr = nStr & "22312"
        Case "0"
            nStr = nStr & "23231"
        Case " "
            nStr = nStr & "34123"
        Case "_"
            nStr = nStr & "14121"
        Case "/"
            nStr = nStr & "14144"
        Case "\"
            nStr = nStr & "12131"
        Case "-"
            nStr = nStr & "12124"
        Case ";"
            nStr = nStr & "21421"
        Case ":"
            nStr = nStr & "21321"
        Case ","
            nStr = nStr & "14431"
        Case "."
            nStr = nStr & "13421"
        Case "+"
            nStr = nStr & "11213"
        Case "="
            nStr = nStr & "11212"
        Case Else
            MsgBox "Caractere não encontrado: " & letra
        End Select
    Next
    NovaStr = nStr
    strEncripta = nStr
    Exit Function
    
End Function

Public Function funMontaStr(ByVal strEncripta As String)
    Dim strData As String
    
    strData = Format(Now, "HHMMSS")
    'strData = ""
    funEncripta (strData)
    strData = NovaStr
    'strData = mStrCriptografa
    
    strEncripta = Mid(strData, 26, 5) & Mid(strEncripta, 1, 5) & Mid(strData, 21, 5) & Mid(strEncripta, 6, 5) _
                & Mid(strData, 16, 5) & Mid(strEncripta, 11, 5) & Mid(strData, 11, 5) & Mid(strEncripta, 16, 5) _
                & Mid(strData, 6, 5) & Mid(strEncripta, 21, 5) & Mid(strData, 1, 5) & Mid(strEncripta, 26, 200)
    
    senhaCripto = strEncripta
    'mStrCriptografa = strEncripta
    Exit Function
    
    
End Function

'
'Public Function funDecripta(ByVal strDecripta As String)
'
'
'    Dim IntTam As Integer
'    Dim nStr As String
'    Dim i As Integer
'    Dim letra As String
'    IntTam = Len(strDecripta)
'    nStr = ""
'
'    'desconsidera os os numeros de HH-MM-SS
'    strDecripta = Mid(strDecripta, 6, 5) & Mid(strDecripta, 16, 5) & Mid(strDecripta, 26, 5) & _
'                  Mid(strDecripta, 36, 5) & Mid(strDecripta, 46, 5) & Mid(strDecripta, 56, 200)
'
'    i = 1
'    Do While Not i = IntTam - 29
'        letra = Mid(strDecripta, i, 5)
'        Select Case letra
'        Case "14334"
'            nStr = nStr & "a"
'        Case "14212"
'            nStr = nStr & "A"
'        Case "24334"
'            nStr = nStr & "á"
'        Case "24134"
'            nStr = nStr & "â"
'        Case "24234"
'            nStr = nStr & "ã"
'        Case "24314"
'            nStr = nStr & "à"
'        Case "24324"
'            nStr = nStr & "b"
'        Case "14223"
'            nStr = nStr & "B"
'        Case "11211"
'            nStr = nStr & "ç"
'        Case "11311"
'            nStr = nStr & "Ç"
'        Case "13334"
'            nStr = nStr & "c"
'        Case "14324"
'            nStr = nStr & "C"
'        Case "24344"
'            nStr = nStr & "d"
'        Case "14444"
'            nStr = nStr & "D"
'        Case "12314"
'            nStr = nStr & "e"
'        Case "21111"
'            nStr = nStr & "E"
'        Case "24321"
'            nStr = nStr & "é"
'        Case "32314"
'            nStr = nStr & "ê"
'        Case "31314"
'            nStr = nStr & "f"
'        Case "21311"
'            nStr = nStr & "F"
'        Case "32134"
'            nStr = nStr & "g"
'        Case "21341"
'            nStr = nStr & "G"
'        Case "31324"
'            nStr = nStr & "h"
'        Case "22111"
'            nStr = nStr & "H"
'        Case "32124"
'            nStr = nStr & "i"
'        Case "21112"
'            nStr = nStr & "I"
'        Case "31334"
'            nStr = nStr & "í"
'        Case "32333"
'            nStr = nStr & "ì"
'        Case "11314"
'            nStr = nStr & "j"
'        Case "23122"
'            nStr = nStr & "J"
'        Case "33134"
'            nStr = nStr & "k"
'        Case "23411"
'            nStr = nStr & "K"
'        Case "33314"
'            nStr = nStr & "l"
'        Case "32222"
'            nStr = nStr & "L"
'        Case "43423"
'            nStr = nStr & "m"
'        Case "32111"
'            nStr = nStr & "M"
'        Case "42423"
'            nStr = nStr & "n"
'        Case "33221"
'            nStr = nStr & "N"
'        Case "43234"
'            nStr = nStr & "o"
'        Case "33233"
'            nStr = nStr & "O"
'        Case "42444"
'            nStr = nStr & "ô"
'        Case "43223"
'            nStr = nStr & "õ"
'        Case "42433"
'            nStr = nStr & "ò"
'        Case "43231"
'            nStr = nStr & "ó"
'        Case "22223"
'            nStr = nStr & "p"
'        Case "33444"
'            nStr = nStr & "P"
'        Case "43233"
'            nStr = nStr & "q"
'        Case "34442"
'            nStr = nStr & "Q"
'        Case "43421"
'            nStr = nStr & "r"
'        Case "34332"
'            nStr = nStr & "R"
'        Case "13443"
'            nStr = nStr & "s"
'        Case "34222"
'            nStr = nStr & "S"
'        Case "44444"
'            nStr = nStr & "t"
'        Case "34112"
'            nStr = nStr & "T"
'        Case "13444"
'            nStr = nStr & "u"
'        Case "41311"
'            nStr = nStr & "U"
'        Case "11111"
'            nStr = nStr & "ú"
'        Case "13243"
'            nStr = nStr & "ù"
'        Case "11115"
'            nStr = nStr & "û"
'        Case "13241"
'            nStr = nStr & "v"
'        Case "41222"
'            nStr = nStr & "V"
'        Case "12443"
'            nStr = nStr & "x"
'        Case "41133"
'            nStr = nStr & "X"
'        Case "13244"
'            nStr = nStr & "y"
'        Case "42231"
'            nStr = nStr & "Y"
'        Case "13441"
'            nStr = nStr & "w"
'        Case "42222"
'            nStr = nStr & "W"
'        Case "11313"
'            nStr = nStr & "z"
'        Case "42213"
'            nStr = nStr & "Z"
'        Case "11312"
'            nStr = nStr & "@"
'        Case "11114"
'            nStr = nStr & "%"
'        Case "12341"
'            nStr = nStr & "&"
'        Case "13343"
'            nStr = nStr & "*"
'        Case "12342"
'            nStr = nStr & "("
'        Case "13344"
'            nStr = nStr & ")"
'        Case "12333"
'            nStr = nStr & "$"
'        Case "23334"
'            nStr = nStr & "!"
'        Case "13331"
'            nStr = nStr & "#"
'        Case "21242"
'            nStr = nStr & "?"
'        Case "22313"
'            nStr = nStr & "1"
'        Case "23424"
'            nStr = nStr & "2"
'        Case "24131"
'            nStr = nStr & "3"
'        Case "41414"
'            nStr = nStr & "4"
'        Case "22314"
'            nStr = nStr & "5"
'        Case "23423"
'            nStr = nStr & "6"
'        Case "44134"
'            nStr = nStr & "7"
'        Case "21241"
'            nStr = nStr & "8"
'        Case "22312"
'            nStr = nStr & "9"
'        Case "23231"
'            nStr = nStr & "0"
'        Case "34123"
'            nStr = nStr & " "
'        Case "14121"
'            nStr = nStr & "_"
'        Case "14144"
'            nStr = nStr & "/"
'        Case "12131"
'            nStr = nStr & "\"
'        Case "12124"
'            nStr = nStr & "-"
'        Case "21421"
'            nStr = nStr & ";"
'        Case "21321"
'            nStr = nStr & ":"
'        Case "14431"
'            nStr = nStr & ","
'        Case "13421"
'            nStr = nStr & "."
'        Case "11213"
'            nStr = nStr & "+"
'        Case "11212"
'            nStr = nStr & "="
'
'        Case Else
'            MsgBox "Código de criptografia inválido!"
'            'mStrDeCriptografa = ""
'            Exit Function
'        End Select
'        i = i + 5
'    Loop
'    'strEncripta = nStr
'    'mStrDeCriptografa = nStr
'Exit Function
'End Function

Private Sub txtBancoPostgres_Change()
   Me.txtUser.TabIndex = Me.txtBancoPostgres.TabIndex + 1
   Me.txtSenha.TabIndex = Me.txtUser.TabIndex + 1
End Sub

Private Sub txtOracleServico_Change()
   Me.txtUser.TabIndex = Me.txtOracleServico.TabIndex + 1
   Me.txtSenha.TabIndex = Me.txtUser.TabIndex + 1
End Sub

Private Sub txtSqlServidor_Change()
   Me.txtUser.TabIndex = Me.txtSqlServidor.TabIndex + 1

End Sub






