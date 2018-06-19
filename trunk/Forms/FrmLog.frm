VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H80000005&
   Caption         =   "Acesso ao Sistema"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsrLog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1350
      TabIndex        =   0
      Top             =   1185
      Width           =   1950
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1335
      TabIndex        =   2
      Top             =   2175
      Width           =   870
   End
   Begin VB.TextBox txtUsrPwd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1350
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1695
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2415
      TabIndex        =   3
      Top             =   2175
      Width           =   870
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   465
      Stretch         =   -1  'True
      Top             =   165
      Width           =   2370
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   1215
      Width           =   900
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   1725
      Width           =   810
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Logou As Boolean
Private UName As Long
Private MyConn As ADODB.Connection
Private Usr As New NexusUsers.clsUsers


Private Sub cmdOK_Click()
'   If Not Usr.FindUser(MyConn, txtUsrLog) Then
'      MsgBox "Usuario não encontrado", vbExclamation
'   Else
'      If Usr.UsrBrk = True Then
'         MsgBox "Usuario não Bloqueado, contate o administrador", vbExclamation
'      ElseIf Usr.UsrExp = True Then
'         MsgBox "Usuario não expirado, contate o administrador", vbExclamation
'      ElseIf Not Usr.UsrPwd = txtUsrPwd Then
'         MsgBox "Senha inválida, tente novamente", vbExclamation
'      Else
'         UName = Usr.UsrId
'         Logou = True
'         Unload Me
'      End If
'   End If

    Dim rs As ADODB.Recordset
    Set rs = Conn.execute("SELECT USRLOG, USRFUN FROM SYSTEMUSERS")
    
    
    If rs.EOF = False Then
        rs.MoveFirst
        Do While Not rs.EOF = True
            If rs!UsrLog = strUser Then
                Exit Do
            End If
            rs.MoveNext
        Loop
        If rs.EOF = False Then
            If rs!UsrFun = 3 Then 'ADMINISTRADOR
                'NÃO É DESABILITADA NENHUMA FUNÇÃO
            ElseIf rs!UsrFun = 2 Then 'OPERADOR
                
            ElseIf rs!UsrFun = 1 Then 'VISITANTE
                FrmMain.mnuDrawLineWater.Enabled = False
                FrmMain.mnuDrawPointInLineWater.Enabled = False
                FrmMain.mnuMovePointWithLines.Enabled = False
                FrmMain.mnuInsertDocs.Enabled = False
                FrmMain.mnuDeleteLineWater.Enabled = False
                FrmMain.mnuDrawRamal.Enabled = False
                FrmMain.mnuInsertLabel.Enabled = False
                FrmMain.mnuCadastros.Visible = False
                FrmMain.mnuAdmin.Visible = False
                FrmMain.tbToolBar.Buttons("kdrawnetworkline").Enabled = False
                FrmMain.tbToolBar.Buttons("kmovenetworknode").Enabled = False
                FrmMain.tbToolBar.Buttons("kinsertnetworknode").Enabled = False
                FrmMain.tbToolBar.Buttons("kinsertdoc").Enabled = False
                FrmMain.tbToolBar.Buttons("kdelete").Enabled = False
                FrmMain.tbToolBar.Buttons("kdrawramal").Enabled = False
            Else
                MsgBox "Não foi encontrado permissão para este usuário.", vbExclamation
                End
            End If
        End If
    Else
        MsgBox "Não foram encontrados usuários no banco de dados.", vbExclamation
        End
    End If

End Sub




