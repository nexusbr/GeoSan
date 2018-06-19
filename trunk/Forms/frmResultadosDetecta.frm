VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmResultadosDetecta 
   Caption         =   "Resultados do Calculo de Rede"
   ClientHeight    =   11580
   ClientLeft      =   3510
   ClientTop       =   2430
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   ScaleHeight     =   11580
   ScaleWidth      =   13290
   Begin TabDlg.SSTab SST 
      Height          =   10305
      Left            =   60
      TabIndex        =   3
      Top             =   690
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   18177
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Uma Coisa"
      TabPicture(0)   =   "frmResultadosDetecta.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fg1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Outra Coisa"
      TabPicture(1)   =   "frmResultadosDetecta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fg2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Outra Outra Coisa"
      TabPicture(2)   =   "frmResultadosDetecta.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fg3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid fg1 
         Height          =   9885
         Left            =   -74940
         TabIndex        =   4
         Top             =   360
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   17436
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid fg2 
         Height          =   9885
         Left            =   -74940
         TabIndex        =   5
         Top             =   360
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   17436
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid fg3 
         Height          =   9885
         Left            =   60
         TabIndex        =   6
         Top             =   360
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   17436
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin VB.TextBox txtIteracoes 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   11070
      Width           =   4275
   End
   Begin VB.CommandButton cmFechar 
      Caption         =   "Fechar"
      Height          =   315
      Left            =   12000
      TabIndex        =   0
      Top             =   11130
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Resultados do Cálculo Hidráulico da Rede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2070
      TabIndex        =   2
      Top             =   90
      Width           =   9075
   End
End
Attribute VB_Name = "frmResultadosDetecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmFechar_Click()
Unload Me
End Sub

Private Sub Form_Load()

fg1.ColWidth(0) = 615
fg1.ColWidth(1) = 765
fg1.ColWidth(2) = 675
fg1.ColWidth(3) = 2625
fg1.ColWidth(4) = 2565
fg1.ColWidth(5) = 1065
fg1.ColWidth(6) = 1455
fg1.ColWidth(7) = 720
fg1.ColWidth(8) = 1080

fg2.ColWidth(0) = 630
fg2.ColWidth(1) = 765
fg2.ColWidth(2) = 690
fg2.ColWidth(3) = 1620
fg2.ColWidth(4) = 1545
fg2.ColWidth(5) = 1380
fg2.ColWidth(6) = 1320


fg3.ColWidth(0) = 330
fg3.ColWidth(1) = 2490
fg3.ColWidth(2) = 915

End Sub


