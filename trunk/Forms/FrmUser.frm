VERSION 5.00
Begin VB.Form FrmUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Usuário"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3600
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5640
      Left            =   90
      TabIndex        =   11
      Top             =   -15
      Width           =   3375
      Begin VB.TextBox txtConfSenha 
         Height          =   405
         Left            =   315
         TabIndex        =   19
         Top             =   2430
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEmail 
         Height          =   330
         Left            =   825
         TabIndex        =   4
         Top             =   2025
         Width           =   2340
      End
      Begin VB.TextBox txtDepto 
         Height          =   330
         Left            =   810
         TabIndex        =   3
         Top             =   1599
         Width           =   2340
      End
      Begin VB.TextBox txtSenha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   825
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1176
         Width           =   2340
      End
      Begin VB.TextBox txtLoginOld 
         Height          =   390
         Left            =   120
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2430
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.CheckBox chkBloqueado 
         Caption         =   "Bloqueado"
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   2460
         Width           =   1170
      End
      Begin VB.Frame Frame2 
         Caption         =   "Permissão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   195
         TabIndex        =   15
         Top             =   2895
         Width           =   2970
         Begin VB.OptionButton optViewer 
            Caption         =   "Visualizador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   630
            TabIndex        =   20
            Top             =   1500
            Width           =   1650
         End
         Begin VB.OptionButton optVisitante 
            Caption         =   "Visitante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   630
            TabIndex        =   8
            Top             =   1110
            Width           =   1290
         End
         Begin VB.OptionButton optUsuario 
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
            Height          =   300
            Left            =   630
            TabIndex        =   7
            Top             =   735
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optAdministrador 
            Caption         =   "Administrador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   630
            TabIndex        =   6
            Top             =   375
            Width           =   1590
         End
      End
      Begin VB.CommandButton cmdFechar 
         Caption         =   "Sair"
         Height          =   375
         Left            =   2070
         TabIndex        =   10
         Top             =   5040
         Width           =   1080
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   900
         TabIndex        =   9
         Top             =   5040
         Width           =   1080
      End
      Begin VB.TextBox txtLogin 
         Height          =   330
         Left            =   825
         TabIndex        =   1
         Top             =   753
         Width           =   2340
      End
      Begin VB.TextBox txtNome 
         Height          =   330
         Left            =   825
         TabIndex        =   0
         Top             =   330
         Width           =   2340
      End
      Begin VB.Label Label4 
         Caption         =   "E-mail"
         Height          =   315
         Left            =   225
         TabIndex        =   18
         Top             =   2070
         Width           =   495
      End
      Begin VB.Label label5 
         Caption         =   "Depto."
         Height          =   315
         Left            =   225
         TabIndex        =   17
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Senha"
         Height          =   315
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Login"
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private MyConn As ADODB.Connection
'Private MyRs As ADODB.Recordset
'Private MyUsers As New NexusUsers.clsUsers
'Private itmx As ListItem
'Private ChangePwd As Boolean
Private strsql As String
Private intBloq As Byte
Private intNivel As Byte
Private blnSenha As Boolean
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
   Unload Me
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Trata_Erro
    If Trim(txtLogin.Text) = "" Or Trim(txtNome.Text) = "" Or Trim(txtSenha.Text) = "" Then
        Exit Sub
    End If
   
    If txtLogin.Text = strUser Then
        ConfereSenha
        If blnSenha = False Then Exit Sub
    End If
    
    If chkBloqueado.value = 1 Then
        intBloq = 1
    Else
        intBloq = 0
    End If
    
    Dim intAcesso As Byte
    If optAdministrador.value = True Then
        intAcesso = 1
    ElseIf optUsuario.value = True Then
        intAcesso = 2
    ElseIf optVisitante.value = True Then
        intAcesso = 3
    ElseIf Me.optViewer.value = True Then
        intAcesso = 4
    End If

RECOMANDO:
a = "SYSTEMUSERS"
b = "USRFUN"
c = "USRBRK"
d = "USRLOG"
e = "USRMAIL"
f = "USRDEPTO"
g = "USRDATA"
h = "USRPWD"
i = "USRDEP"
j = "USRNOM"
k = "USRLOG"


    If Me.Caption = " Usuário - Editar" Then
        If frmCanvas.TipoConexao <> 4 Then
        strsql = "UPDATE SYSTEMUSERS SET USRFUN='" & intAcesso & "',USRBRK ='" & intBloq & "',"
        strsql = strsql & "USRLOG='" & txtLogin.Text & "',USRMAIL='" & txtEmail.Text & "',USRDEPTO='" & txtDepto.Text & "',"
        strsql = strsql & "USRDATA='" & Format(Now, "DDMMYYYY") & "', USRPWD= '" & txtSenha.Text & "',USRDEP = 1 "
        strsql = strsql & "WHERE USRNOM= '" & txtNome.Text & "' AND USRLOG = '" & txtLoginOld.Text & "'"
        Else
        strsql = "UPDATE " + """" + a + """" + " SET " + """" + b + """" + "='" & intAcesso & "'," + """" + c + """" + " ='" & intBloq & "',"
        strsql = strsql + """" + d + """" + "='" & txtLogin.Text & "'," + """" + e + """" + "='" & txtEmail.Text & "'," + """" + f + """" + "='" & txtDepto.Text & "',"
        strsql = strsql + """" + g + """" + "='" & Format(Now, "DDMMYYYY") & "', " + """" + h + """" + "= '" & txtSenha.Text & "'," + """" + i + """" + " = '1' "
        strsql = strsql + "WHERE " + """" + j + """" + "= '" & txtNome.Text & "' AND " + """" + k + """" + " = '" & txtLoginOld.Text & "'"
        End If
        
        Conn.execute strsql
        
        'MsgBox "Usuário '" & txtNome.Text & "' atualizado com êxito!", vbInformation, "Confirmação"
        
        Unload Me
    Else
       
        Dim rs As ADODB.Recordset
a = "USRLOG"
b = "SYSTEMUSERS"


     If frmCanvas.TipoConexao <> 4 Then
        Set rs = Conn.execute("SELECT USRLOG FROM SYSTEMUSERS WHERE USRLOG = '" & txtLogin.Text & "'")
        Else
        Set rs = Conn.execute("SELECT " + """" + a + """" + " FROM " + """" + b + """" + " WHERE " + """" + a + """" + " = '" & txtLogin.Text & "'")
        End If
        If rs.EOF = True Then 'O NOVO NOME DE USUÁRIO NÃO EXISTE
            
      a = "SYSTEMUSERS"
      b = "USRLOG"
      c = "USRNOM"
      d = "USRFUN"
      e = "USRPWD"
      f = "USRDEP"
      g = "USRBRK"
      h = "USRDATA"
      i = "USRMAIL"
      j = "USRDEPTO"
  
     If frmCanvas.TipoConexao = 1 Then
         
     strsql = "INSERT INTO SYSTEMUSERS(USRLOG,USRNOM,USRFUN,USRPWD,USRDEP,USRBRK,USRDATA,USRMAIL,USRDEPTO) values "
            strsql = strsql & "('" & txtLogin.Text & "','" & txtNome.Text & "','" & intAcesso & "','" & txtSenha.Text & "','1','" & intBloq
            strsql = strsql & "','" & Format(Now, "DDMMYYYY") & "','" & txtEmail.Text & "','" & txtDepto.Text & "')"
     
     
     ElseIf frmCanvas.TipoConexao = 2 Then
      strsql = "INSERT INTO SYSTEMUSERS(USRLOG,USRNOM,USRFUN,USRPWD,USRDEP,USRBRK,USRDATA,USRMAIL,USRDEPTO) values "
            strsql = strsql & "('" & txtLogin.Text & "','" & txtNome.Text & "','" & intAcesso & "','" & txtSenha.Text & "','1','" & intBloq
            strsql = strsql & "','" & "0" & "','" & txtEmail.Text & "','" & txtDepto.Text & "')"
     
     
     Else
     
  strsql = "INSERT INTO " + """" + a + """" + "(" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + "," + """" + e + """" + "," + """" + f + """" + "," + """" + g + """" + "," + """" + h + """" + "," + """" + i + """" + "," + """" + j + """" + "," + """" + "USREXP" + """" + ") values "
            strsql = strsql & "('" & txtLogin.Text & "','" & txtNome.Text & "','" & intAcesso & "','" & txtSenha.Text & "','1','" & intBloq
            strsql = strsql & "','" & Format(Now, "DDMMYYYY") & "','" & txtEmail.Text & "','" & txtDepto.Text & "','False')"
            
            
          '  MsgBox strsql
     End If
            
            
          
             
            Conn.execute strsql
            MsgBox "Usuário '" & txtNome.Text & "' cadastrado com êxito!", vbInformation, "Confirmação"
            txtNome.Text = ""
            txtLogin.Text = ""
            txtSenha.Text = ""
            txtDepto.Text = ""
            txtEmail.Text = ""
            chkBloqueado.value = 0
            optUsuario.value = True
            txtNome.SetFocus
        Else
            MsgBox "O login '" & txtLogin.Text & "' pertence a outro usuário." & Chr(13) & "Escolha outro login para este novo usuário.", vbInformation
            txtLogin.SetFocus
            txtLogin.SelStart = 0
            txtLogin.SelLength = Len(txtLogin.Text)
        End If
    End If
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = 3265 Then
        Err.Clear
        Resume Next
    ElseIf Err.Number = -2147217900 Then
        Err.Clear
        MsgBox "Atualizando o banco de dados para esta nova funcionalidade.", vbInformation, ""
        'NOME DE COLUNA INVÁLIDO, CRIANDO COLUNAS
        strsql = "ALTER TABLE SYSTEMUSERS ADD (USRMAIL CHAR(50),USRDEPTO CHAR(50),USRDATA CHAR(8))"
        'strsql = "ALTER TABLE SYSTEMUSERS DROP (USRDEPTO)"
        Conn.execute strsql
        GoTo RECOMANDO 'RETORNA AO PONTO DE SALVAMENTO/EDIÇÃO
    Else
       
      PrintErro CStr(Me.Name), "Public Function Init", CStr(Err.Number), CStr(Err.Description), True
          
    End If
End Sub

Private Function ConfereSenha()
   
        frmConfSenha.Show 1
        If txtConfSenha.Text <> txtSenha.Text Then
            MsgBox "A senha não confere.", vbExclamation, ""
            txtSenha.SelLength = (Len(txtSenha.Text))
            txtSenha.SetFocus
            blnSenha = False
        Else
            blnSenha = True
        End If

End Function

Private Sub txtDepto_GotFocus()
    txtDepto.SelLength = (Len(txtDepto.Text))
End Sub

Private Sub txtEmail_GotFocus()
    txtEmail.SelLength = (Len(txtEmail.Text))
End Sub

Private Sub txtNome_GotFocus()
    txtNome.SelLength = (Len(txtNome.Text))
End Sub

Private Sub txtLogin_GotFocus()
    txtLogin.SelLength = (Len(txtLogin.Text))
End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.SelLength = (Len(txtSenha.Text))
End Sub

