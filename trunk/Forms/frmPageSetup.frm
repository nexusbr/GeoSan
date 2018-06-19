VERSION 5.00
Begin VB.Form frmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar Página"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5685
      TabIndex        =   22
      Top             =   4125
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4575
      TabIndex        =   21
      Top             =   4125
      Width           =   1050
   End
   Begin VB.Frame Frame4 
      Caption         =   "Margem"
      Height          =   1515
      Left            =   3030
      TabIndex        =   12
      Top             =   2415
      Width           =   3720
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2715
         TabIndex        =   19
         Text            =   "0.75"
         Top             =   885
         Width           =   795
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   810
         TabIndex        =   17
         Text            =   "0.75"
         Top             =   900
         Width           =   795
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2715
         TabIndex        =   15
         Text            =   "0.75"
         Top             =   315
         Width           =   795
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   810
         TabIndex        =   13
         Text            =   "0.75"
         Top             =   330
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abaixo"
         Height          =   195
         Left            =   2130
         TabIndex        =   20
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topo"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   975
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direita"
         Height          =   195
         Left            =   2130
         TabIndex        =   16
         Top             =   390
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Esqerda"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   405
         Width           =   585
      End
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "Orientação"
      Height          =   1515
      Left            =   150
      TabIndex        =   7
      Top             =   2415
      Width           =   2715
      Begin VB.OptionButton opPortrait 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   120
         Picture         =   "frmPageSetup.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   255
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton opLandscape 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   1440
         Picture         =   "frmPageSetup.frx":274E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paisagem"
         Height          =   195
         Left            =   1590
         TabIndex        =   11
         Top             =   1140
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retrato"
         Height          =   195
         Left            =   315
         TabIndex        =   10
         Top             =   1110
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Visualização"
      Height          =   2055
      Left            =   4470
      TabIndex        =   3
      Top             =   180
      Width           =   2295
      Begin VB.Image imgPortrait 
         Height          =   1470
         Left            =   600
         Picture         =   "frmPageSetup.frx":5034
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image imgLandScape 
         Height          =   1080
         Left            =   360
         Picture         =   "frmPageSetup.frx":A19E
         Top             =   600
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Customizado"
      Height          =   735
      Left            =   150
      TabIndex        =   2
      Top             =   1515
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Papel"
      Height          =   1170
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   3975
      Begin VB.ComboBox cmbSource 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPageSetup.frx":F880
         Left            =   945
         List            =   "frmPageSetup.frx":F882
         TabIndex        =   5
         Top             =   720
         Width           =   2790
      End
      Begin VB.ComboBox cmbPageSize 
         Height          =   315
         ItemData        =   "frmPageSetup.frx":F884
         Left            =   945
         List            =   "frmPageSetup.frx":F89D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2790
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleção"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   750
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamanho"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private orientation As Boolean
Private pageSize As Integer
Private isOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
orientation = opPortrait.value
pageSize = cmbPageSize.ListIndex
isOK = True
Unload Me
End Sub



Private Sub Form_Load()
    cmbPageSize.ListIndex = 0
    isOK = False
End Sub

Private Sub opLandscape_Click()
imgLandScape.Visible = True
imgPortrait.Visible = False
End Sub

Private Sub opPortrait_Click()
imgLandScape.Visible = False
imgPortrait.Visible = True

End Sub

Public Function getOrientation() As Boolean
    getOrientation = orientation
End Function

Public Sub setOrientation(value As Boolean)
    opPortrait.value = value
    opLandscape.value = Not value
End Sub

Public Function getPageSize() As Integer
    getPageSize = pageSize
End Function


Public Sub setPageSize(value As Integer)
    cmbPageSize.ListIndex = value
End Sub

Function getOK() As Boolean
    getOK = isOK
End Function

