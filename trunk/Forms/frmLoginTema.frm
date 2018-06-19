VERSION 5.00
Begin VB.Form frmLoginTema 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login de Administrador:"
   ClientHeight    =   1500
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   886.25
   ScaleMode       =   0  'User
   ScaleWidth      =   3563.3
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1080
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1665
      TabIndex        =   4
      Top             =   1020
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2565
      TabIndex        =   5
      Top             =   1020
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "Login"
      Height          =   270
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   150
      Width           =   570
   End
   Begin VB.Label lblLabels 
      Caption         =   "Senha"
      Height          =   270
      Index           =   1
      Left            =   315
      TabIndex        =   2
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "frmLoginTema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub cmdCancel_Click()
   
   Me.Hide

End Sub

Private Sub cmdOK_Click()
    'check for correct password
    Dim rs As ADODB.Recordset
    
   If txtPassword.Text <> "" And Me.txtUserName.Text <> "" Then
      Set rs = New ADODB.Recordset
      
a = "USRLOG"
b = "USRFUN"
c = "SYSTEMUSERS"
d = "USRPwd"
e = "HIDROMETRADO"
f = "ECONOMIAS"
g = "CONSUMO_LPS"
h = "TB_LIGACOES"
i = "HIDROMETRADO"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"

      If frmCanvas.TipoConexao <> 4 Then
      rs.Open "SELECT USRLOG, USRFUN FROM SYSTEMUSERS WHERE USRLOG = '" & Me.txtUserName.Text & "' AND USRPwd = '" & Me.txtPassword.Text & "'", Conn, adOpenDynamic, adLockReadOnly
      Else
      rs.Open "SELECT " + """" + a + """" + "," + """" + b + """" + " FROM " + """" + c + """" + " WHERE " + """" + a + """" + " = '" & Me.txtUserName.Text & "' AND " + """" + d + """" + " = '" & Me.txtPassword.Text & "'", Conn, adOpenDynamic, adLockOptimistic
      End If
      
      If rs.EOF = False Then
         If rs!UsrFun = 1 Then  '   ADMINISTRADOR
            
            FrmMain.pctSfondo.Visible = True
            
            Me.txtUserName.Text = ""
            Me.txtPassword.Text = ""
            
            Me.Hide
         End If
      Else
         MsgBox "Login ou senha incorreto.", vbInformation, ""
      End If
      rs.Close
   End If
   
    
End Sub

