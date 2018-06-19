VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndicProdutRedesDeAgua 
   Caption         =   "Indicador de Produtividade - Redes de Agua"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgBar1 
      Height          =   315
      Left            =   165
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Max             =   5
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   360
      Left            =   5085
      TabIndex        =   1
      Top             =   840
      Width           =   1140
   End
   Begin VB.TextBox txtCaminho 
      Height          =   330
      Left            =   165
      TabIndex        =   0
      Top             =   390
      Width           =   6060
   End
   Begin VB.Label lblCaminho 
      Caption         =   "Caminho do Arquivo"
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   135
      Width           =   2985
   End
End
Attribute VB_Name = "frmIndicProdutRedesDeAgua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoRede As String ' Rede de Agua ou Esgoto


Private Sub Form_Load()

   'A CHAMADA DO FORM CARREGA A VARIÁVEL TipoRede

   If UCase(TipoRede) = "ESGOTO" Then
      
      Me.Caption = "Indicador de Produtividade - Redes de Esgoto"
      
      txtCaminho.Text = App.path & "\Indicador_RedesDeEsgoto_" & Format(Now, "YYYYMMDD") & ".txt"
      
   Else
      
      Me.Caption = "Indicador de Produtividade - Redes de Água"
      
      txtCaminho.Text = App.path & "\Indicador_RedesDeAgua_" & Format(Now, "YYYYMMDD") & ".txt"
   
   End If
   
    
End Sub

Private Sub cmdGerar_Click()

   If TipoRede = "ESGOTO" Then

      If RelProdutividade("ESGOTO") = True Then
         MsgBox "Relatório gerado com sucesso!", vbInformation, ""
      Else
         MsgBox "Falha ao gerar o relatório!", vbInformation, ""
      End If
      
   Else
   
      If RelProdutividade("AGUA") = True Then
         MsgBox "Relatório gerado com sucesso!", vbInformation, ""
      Else
         MsgBox "Falha ao gerar o relatório!", vbInformation, ""
      End If
   
   
   End If
      
   Unload Me
   

End Sub

