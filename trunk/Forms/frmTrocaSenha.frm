VERSION 5.00
Begin VB.Form frmTrocaSenha 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Troca de Senha"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2475
      TabIndex        =   5
      Top             =   2625
      Width           =   1035
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1335
      TabIndex        =   4
      Top             =   2625
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   45
      Width           =   3390
      Begin VB.TextBox txtConfirmaSenha 
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
         Left            =   1230
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1830
         Width           =   1875
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1230
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   405
         Width           =   1875
      End
      Begin VB.TextBox txtNovaSenha 
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
         Left            =   1230
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1350
         Width           =   1875
      End
      Begin VB.TextBox txtSenhaAtual 
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
         Left            =   1230
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   870
         Width           =   1875
      End
      Begin VB.Label Label4 
         Caption         =   "Confirmação"
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   1845
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Nova senha"
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   1365
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Usuário"
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   450
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Senha atual"
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   915
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmTrocaSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strsql As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String
Dim k As String
Dim l As String


Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
a = "USRLOG"
b = "USRPWD"
c = "SYSTEMUSERS"

    If txtNovaSenha.Text = txtConfirmaSenha.Text And txtNovaSenha.Text <> "" Then
    
        Dim rs As ADODB.Recordset
     If frmCanvas.TipoConexao <> 4 Then
        Set rs = Conn.execute("SELECT USRLOG,USRPWD FROM SYSTEMUSERS WHERE USRLOG = '" & txtUsuario.Text & "' and USRPWD = '" & txtSenhaAtual.Text & "'")
        Else
        Set rs = Conn.execute("SELECT " + """" + a + """" + "," + """" + b + """" + " FROM " + """" + c + """" + " WHERE " + """" + a + """" + " = '" & txtUsuario.Text & "' and " + """" + b + """" + " = '" & txtSenhaAtual.Text & "'")
        End If
        If rs.EOF = True Then 'A senha não foi encontrada
            MsgBox "A senha atual não confere.", vbInformation + vbOKOnly, ""
            Exit Sub
        End If
a = "SYSTEMUSERS"
b = "USRPWD"
c = "USRLOG"
     If frmCanvas.TipoConexao <> 4 Then
        strsql = "UPDATE SYSTEMUSERS SET USRPWD= '" & txtNovaSenha.Text & "' WHERE USRLOG = '" & txtUsuario.Text & "'"
        Else
        strsql = "UPDATE " + """" + a + """" + " SET " + """" + b + """" + "= '" & txtNovaSenha.Text & "' WHERE " + """" + c + """" + " = '" & txtUsuario.Text & "'"
        End If
        Conn.execute strsql
        MsgBox "Sua senha foi alterada com sucesso!", vbInformation, "Nova senha"
        Unload Me
    Else
    
        MsgBox "A Nova senha não foi confirmada corretamente.", vbExclamation + vbOKOnly, ""
        txtNovaSenha.Text = ""
        txtConfirmaSenha.Text = ""
        txtNovaSenha.SetFocus
        
    End If

End Sub

