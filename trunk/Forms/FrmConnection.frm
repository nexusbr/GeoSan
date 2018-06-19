VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmConnection 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conexão"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4035
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSenha 
      BackColor       =   &H8000000E&
      Height          =   1215
      Left            =   150
      TabIndex        =   19
      Top             =   2325
      Width           =   3705
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   720
         Width           =   2715
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   270
         Width           =   2715
      End
      Begin VB.Label lblSenha 
         BackColor       =   &H8000000E&
         Caption         =   "Senha:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblUser 
         BackColor       =   &H8000000E&
         Caption         =   "Usuário:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Sair"
      Height          =   345
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3690
      Width           =   885
   End
   Begin VB.Frame FrameOption 
      BackColor       =   &H8000000E&
      Caption         =   "Tipo de Conexão"
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   150
      TabIndex        =   6
      Top             =   150
      Width           =   3705
      Begin VB.OptionButton optFireBird 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fire Bird"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2550
         TabIndex        =   11
         Top             =   570
         Width           =   1095
      End
      Begin VB.OptionButton optOracle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Oracle"
         Height          =   225
         Left            =   2550
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optAccess 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Access"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSql 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SQL Server / MSDE"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   570
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   345
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3690
      Width           =   885
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameAcs 
      BackColor       =   &H8000000E&
      Height          =   1215
      Left            =   150
      TabIndex        =   12
      Top             =   1125
      Width           =   3705
      Begin VB.CommandButton cmdFile 
         Caption         =   "..."
         Height          =   255
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   510
         Width           =   375
      End
      Begin VB.TextBox txtAcsBanco 
         Height          =   285
         Left            =   810
         TabIndex        =   13
         Top             =   510
         Width           =   2355
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Banco:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   510
         Width           =   855
      End
   End
   Begin VB.Frame frameOracle 
      BackColor       =   &H8000000E&
      Height          =   1215
      Left            =   150
      TabIndex        =   16
      Top             =   1125
      Width           =   3705
      Begin VB.TextBox txtOracleServico 
         Height          =   285
         Left            =   810
         TabIndex        =   17
         Top             =   510
         Width           =   2715
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Serviço:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   510
         Width           =   855
      End
   End
   Begin VB.Frame frameSql 
      BackColor       =   &H8000000E&
      Height          =   1215
      Left            =   150
      TabIndex        =   5
      Top             =   1125
      Width           =   3705
      Begin VB.TextBox txtSqlBanco 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   270
         Width           =   2715
      End
      Begin VB.TextBox txtSqlServidor 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   2715
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Banco:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Servidor:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Registrou As Boolean, apName As String, MyConn As New ADODB.Connection, Tconn As String



Public Function init(appName As String, Conn As ADODB.Connection, Optional TypeConn As cAppType) As Boolean
    
    Registrou = False
    apName = appName 'NOME DO APLICATIVO QUE CHAMOU A FUNÇÃO
    Me.Show vbModal
    
    If Tconn <> "" Then
        init = Registrou
        TypeConn = Tconn
    End If
    
    Set Conn = MyConn
    Set MyConn = Nothing

End Function

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

Private Sub cmdOK_Click()
On Error GoTo Trata_Erro
    Tconn = ""
    If optAccess.value Then 'OPÇÃO DE CONEXÃO DE BANCO DE DADOS ACCESS
        If Trim(txtAcsBanco.Text) <> "" Then
            Me.MousePointer = vbHourglass
            MyConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtAcsBanco & ";Persist Security Info=False"
            Tconn = Access
        Else
            MsgBox "Escolha o banco de dados para conexão do GeoSan.", vbInformation, "Atenção!"
            cmdFile.SetFocus
            Exit Sub
        End If
    ElseIf optSql.value Then ''OPÇÃO DE CONEXÃO DE BANCO DE DADOS SQL SERVER
    
        If Trim(txtSqlServidor.Text) <> "" Then
            If Trim(txtSqlBanco.Text) <> "" Then
                If txtUser.Text <> "" Then
                    If txtSenha.Text <> "" Then
                        Me.MousePointer = vbHourglass
                        MyConn.Open "Provider=SQLOLEDB.1;Persist Security Info=True;Data Source=" & txtSqlServidor & ";User ID=" & txtUser.Text & ";Password=" & txtSenha.Text & ";Initial Catalog=" & txtSqlBanco
                        Tconn = SqlServer
                    Else
                        MsgBox "Digite a senha do banco SQL para conexão com o GeoSan.", vbInformation, "Atenção!"
                        txtSenha.SetFocus
                        Exit Sub
                    End If
                Else
                    MsgBox "Digite o nome de usuário do banco SQL para conexão com o GeoSan.", vbInformation, "Atenção!"
                    txtUser.SetFocus
                    Exit Sub
                End If
            Else
                MsgBox "Digite o nome do banco SQL para conexão com o GeoSan.", vbInformation, "Atenção!"
                txtSqlBanco.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Digite o nome do servidor sql para conexão com o GeoSan.", vbInformation, "Atenção!"
            txtSqlServidor.SetFocus
            Exit Sub
        End If
    
    ElseIf optOracle.value Then 'OPÇÃO DE CONEXÃO DE BANCO DE DADOS ORACLE
        
        If txtOracleServico <> "" Then
            If txtUser.Text <> "" Then
                If txtSenha.Text <> "" Then
                    Me.MousePointer = vbHourglass
                    MyConn.Open "Provider=OraOLEDB.Oracle.1;Password=" & txtSenha.Text & ";Persist Security Info=True;User ID=" & txtUser.Text & ";Data Source=" & txtOracleServico
                    Tconn = Oracle
                Else
                    MsgBox "Digite a senha do banco Oracle para conexão com o GeoSan.", vbInformation, "Atenção!"
                    txtSenha.SetFocus
                    Exit Sub
                End If
            Else
                MsgBox "Digite o nome de usuário do banco Oracle para conexão com o GeoSan.", vbInformation, "Atenção!"
                txtUser.SetFocus
                Exit Sub
            End If
        End If
        
    End If
    
    If Trim(Tconn) <> "" Then 'SE A CONEXÃO FOI BEM SUCEDIDA, GRAVA AS INFORMAÇÕES NO ARQUIVO
        Open App.path & "\Controles\GeoSan.cfg" For Output As #1 'ABRE O ARQUIVO EXCLUINDO DADOS ANTERIORES
            Print #1, Tconn
            Print #1, txtSqlServidor.Text
            Print #1, txtSqlBanco.Text
            Print #1, txtAcsBanco.Text
            Print #1, txtOracleServico.Text
            Print #1, txtUser.Text
            Print #1, txtSenha.Text
        Close #1
            
    End If
    
    MsgBox "Banco de dados redirecionado com sucesso." & Chr(13) & Chr(13) & "Reinicie o sistema para ativar.", vbInformation
    
    
    'Shell App.path & "\" & App.EXEName & ".exe"
    'Shell App.path & "\" & "GeoSan.exe"
    End 'finaliza o aplicativo
    
Trata_Erro:
    Me.MousePointer = vbDefault
    Close #1
    If Err.Number = 0 Or Err.Number = 20 Then ' ERROS 0 E 20 SÃO DESCONSIDERADOS
        Resume Next
    ElseIf Err.Number = 3705 Then
        MyConn.Close
        Resume
    ElseIf Trim(Tconn) = "" Then ' SIGNIFICA QUE NÃO CONSEGUIU CONECTAR NA BASE SELECIONADA E O ERRO FOI DE CONEXÃO
        If optOracle.value Then
            MsgBox "Não foi possível estabelecer a conexão com o banco Oracle : " & Chr(13) & Chr(13) & Err.Description, vbInformation
        
        ElseIf optSql.value Then
            MsgBox "Não foi possível estabelecer a conexão com o banco SQL : " & Chr(13) & Chr(13) & Err.Description, vbInformation
        
        ElseIf optAccess.value Then
            MsgBox "Não foi possível estabelecer a conexão com o banco Access : " & Chr(13) & Chr(13) & Err.Description, vbInformation
        
        End If
    ElseIf Err.Number = 55 Then
        Close #1
        Resume
    ElseIf Err.Number = 75 Then 'ERRO DE ACESSO AO ARQUIVO
        MsgBox "Não foi possível gravar o arquivo " & apName & ".cfg na pasta '" & App.path & "'." & Chr(13) & Chr(13) & "É necessário que o usuário possua permissão para gravar arquivos nesta pasta.", vbExclamation, "Erro de acesso"
    Else
    
       PrintErro CStr(Me.Name), "cmdOK_Click", CStr(Err.Number), CStr(Err.Description), True
    
    End If

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
   Click
End Sub

Private Sub optAccess_Click()
   Click
End Sub

Sub Click()
   
'ARRUMA A TELA PARA QUE O USUÁRIO INSIRA AS INFORMAÇÕES DE CONEXÃO DE BANCO DE DADOS
   
   txtAcsBanco = ""
   txtOracleServico = ""
   txtSqlBanco = ""
   txtSqlServidor = ""
   If optAccess.value Then
      frameAcs.Visible = True
      frameOracle.Visible = False
      frameSql.Visible = False
      LiberaFrameLogin False
'      cmdClose.Top = FrameOption.Height + frameAcs.Height + 120
'      cmdOk.Top = FrameOption.Height + frameAcs.Height + 120
'      Me.Height = FrameOption.Height + frameAcs.Height + cmdOk.Height + 500
   ElseIf optSql.value Then
      frameAcs.Visible = False
      frameOracle.Visible = False
      frameSql.Visible = True
      LiberaFrameLogin True
'      cmdClose.Top = FrameOption.Height + frameSql.Height + 50
'      cmdOk.Top = FrameOption.Height + frameSql.Height + 50
'      Me.Height = FrameOption.Height + frameSql.Height + cmdOk.Height + 500
   ElseIf optOracle.value Then
      frameAcs.Visible = False
      frameOracle.Visible = True
      frameSql.Visible = False
      LiberaFrameLogin True
'      cmdClose.Top = FrameOption.Height + frameOracle.Height + 50
'      cmdOk.Top = FrameOption.Height + frameOracle.Height + 50
'      Me.Height = FrameOption.Height + frameOracle.Height + cmdOk.Height + 500
   End If

End Sub

Sub LiberaFrameLogin(pLibera As Boolean)

    fraSenha.Enabled = pLibera
    lblUser.Enabled = pLibera
    lblSenha.Enabled = pLibera
    txtUser.Enabled = pLibera
    txtSenha.Enabled = pLibera
    
    If pLibera = False Then
        txtUser.Text = ""
        txtSenha.Text = ""
    End If

End Sub

Private Sub optFireBird_Click()
   Click
End Sub

Private Sub optOracle_Click()
   Click
End Sub

Private Sub optSql_Click()
   Click
End Sub



