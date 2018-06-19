VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUserControle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Controle de Usuários GeoSan - Acesso de Administrador"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   400
      Left            =   9660
      TabIndex        =   4
      Top             =   1395
      Width           =   1120
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   400
      Left            =   9660
      TabIndex        =   3
      Top             =   930
      Width           =   1120
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   400
      Left            =   9660
      TabIndex        =   2
      Top             =   225
      Width           =   1120
   End
   Begin MSFlexGridLib.MSFlexGrid Flex1 
      Height          =   2490
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4392
      _Version        =   393216
      Cols            =   6
      BackColorBkg    =   14737632
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   400
      Left            =   9660
      TabIndex        =   0
      Top             =   2235
      Width           =   1120
   End
End
Attribute VB_Name = "frmUserControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Co As Integer
Dim Li As Integer
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

Private Sub cmdExcluir_Click()
    If Flex1.TextMatrix(Flex1.Row, 1) = "Admin. do Sistema" Then
        MsgBox "O usuário 'Admin. do Sistema' não pode ser excluído.", vbOKOnly + vbExclamation, "Excluir"
        Exit Sub
    End If
    If MsgBox("Confirma exclusão do usuário '" & Flex1.TextMatrix(Flex1.Row, 2) & "' ?", vbQuestion + vbYesNo + vbDefaultButton2, "Excluir") = vbYes Then
    a = "SYSTEMUSERS"
b = "USRLOG"

     If frmCanvas.TipoConexao <> 4 Then
        Conn.execute ("DELETE FROM SYSTEMUSERS WHERE USRLOG = '" & Flex1.TextMatrix(Flex1.Row, 2) & "' ")
        Else
         Conn.execute ("DELETE FROM " + """" + a + """" + " WHERE " + """" + b + """" + " = '" & Flex1.TextMatrix(Flex1.Row, 2) & "' ")
        End If
        CarregaForm
    End If
    
    
End Sub

Private Sub Flex1_DblClick()
    Editar
    CarregaForm
End Sub

Private Sub cmdEditar_Click()
    Editar
    CarregaForm
End Sub
Private Function Editar()
    FrmUser.txtLogin.Text = Flex1.TextMatrix(Flex1.Row, 2)
    FrmUser.txtLoginOld.Text = Flex1.TextMatrix(Flex1.Row, 2)
    FrmUser.txtNome.Text = Flex1.TextMatrix(Flex1.Row, 1)
    FrmUser.txtDepto.Text = Flex1.TextMatrix(Flex1.Row, 4)
    FrmUser.txtEmail.Text = Flex1.TextMatrix(Flex1.Row, 5)
    FrmUser.txtSenha.Text = Flex1.TextMatrix(Flex1.Row, 8)
    
    If FrmUser.txtNome.Text = "Admin. do Sistema" Then
        FrmUser.optAdministrador.value = True
        FrmUser.txtLogin.Enabled = False
        FrmUser.chkBloqueado.value = False
        FrmUser.chkBloqueado.Enabled = False
        FrmUser.Frame2.Enabled = False
    End If
    
    If Flex1.TextMatrix(Flex1.Row, 7) = "Sim" Then
        FrmUser.chkBloqueado.value = 1
    End If
    
    If Flex1.TextMatrix(Flex1.Row, 3) = "Administrador" Then
        FrmUser.optAdministrador = True
    ElseIf Flex1.TextMatrix(Flex1.Row, 3) = "Usuário" Then
        FrmUser.optUsuario = True
    ElseIf Flex1.TextMatrix(Flex1.Row, 3) = "Visitante" Then
        FrmUser.optVisitante = True
    ElseIf Flex1.TextMatrix(Flex1.Row, 3) = "Visualizador" Then
        FrmUser.optViewer = True
    
    End If
    
    FrmUser.Caption = " Usuário - Editar"
    FrmUser.txtNome.Enabled = False
    
    FrmUser.Show 1
End Function


Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdNovo_Click()
    FrmUser.Caption = " Usuário - Novo"
    FrmUser.Show 1
    CarregaForm
End Sub

Private Function CarregaForm()
On Error GoTo Trata_Erro
    Flex1.Clear
    Flex1.Rows = 1
    Flex1.Cols = 9
    Flex1.ColWidth(0) = 350
    Flex1.ColWidth(1) = 1700 '"Nome do usuário"
    Flex1.ColWidth(2) = 1300 '"Login"
    Flex1.ColWidth(3) = 1100 '"Permissão"
    Flex1.ColWidth(4) = 2000 '"Departamento"
    Flex1.ColWidth(5) = 1950 '"E-mail"
    Flex1.ColWidth(6) = 900 '"Data"
    Flex1.ColWidth(7) = 900 '"Bloqueado"
    Flex1.ColWidth(8) = 0 '"Senha
    
    Flex1.TextMatrix(0, 1) = "Nome do usuário"
    Flex1.TextMatrix(0, 2) = "Login"
    Flex1.TextMatrix(0, 3) = "Permissão"
    Flex1.TextMatrix(0, 4) = "Departamento"
    Flex1.TextMatrix(0, 5) = "E-mail"
    Flex1.TextMatrix(0, 6) = "Data"
    Flex1.TextMatrix(0, 7) = "Bloqueado"
    
    Dim rs As ADODB.Recordset
    
    a = "SYSTEMUSERS"
    b = "USRNOM"
    
    If frmCanvas.TipoConexao <> 4 Then
    Set rs = Conn.execute("SELECT * FROM SYSTEMUSERS ORDER BY USRNOM")
    Else
    Set rs = Conn.execute("SELECT * FROM " + """" + a + """" + " ORDER BY " + """" + b + """" + "")
    End If
    
    If rs.EOF = False Then
        rs.MoveFirst
        Li = 1
        Do While Not rs.EOF = True
            Flex1.Rows = Flex1.Rows + 1
            Flex1.TextMatrix(Li, 1) = rs!UsrNom
            Flex1.TextMatrix(Li, 2) = rs!UsrLog
            If rs!UsrFun = 1 Then
                Flex1.TextMatrix(Li, 3) = "Administrador"
            ElseIf rs!UsrFun = 2 Then
                Flex1.TextMatrix(Li, 3) = "Usuário"
            ElseIf rs!UsrFun = 3 Then
                Flex1.TextMatrix(Li, 3) = "Visitante"
            Else
               Flex1.TextMatrix(Li, 3) = "Visualizador"
            End If
            Flex1.TextMatrix(Li, 4) = rs!USRDEPTO
            Flex1.TextMatrix(Li, 5) = rs!USRMAIL
            Flex1.TextMatrix(Li, 6) = Format(rs!USRDATA, "00-00-0000")
            If rs!UsrBrk <> 0 Then
                 Flex1.TextMatrix(Li, 7) = "Sim"
            Else
                 Flex1.TextMatrix(Li, 7) = "Não"
            End If
            Flex1.TextMatrix(Li, 8) = rs!UsrPwd
            Li = Li + 1
            rs.MoveNext
        Loop
    End If
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = 3265 Then
        Err.Clear
        Resume Next
    ElseIf Err.Number = 94 Then
        Err.Clear
        Resume Next
    Else
       
      PrintErro CStr(Me.Name), "Private Function CarregaForm()", CStr(Err.Number), CStr(Err.Description), True
          
    End If
End Function

Private Sub Form_Load()
    CarregaForm
End Sub

Private Sub Flex1_Click()
    'MsgBox Flex1.TextMatrix(Flex1.Row, 2)
    Li = Flex1.Row
    Co = 2
End Sub


