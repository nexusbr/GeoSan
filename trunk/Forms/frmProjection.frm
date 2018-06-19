VERSION 5.00
Begin VB.Form frmProjection 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informe projeção"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ab_frame 
      Height          =   5715
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Height          =   345
         Left            =   1440
         TabIndex        =   28
         Top             =   5250
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   2700
         TabIndex        =   27
         Top             =   5250
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5205
         Left            =   30
         TabIndex        =   1
         Top             =   0
         Width           =   3765
         Begin VB.Frame Frame3 
            Caption         =   "Parâmetros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3555
            Left            =   180
            TabIndex        =   4
            Top             =   1530
            Width           =   3465
            Begin VB.ComboBox cboZone 
               Height          =   315
               Left            =   1140
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   240
               Width           =   2205
            End
            Begin VB.TextBox txtParPadrao2 
               BackColor       =   &H8000000E&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   16
               Text            =   "0"
               Top             =   1470
               Width           =   2175
            End
            Begin VB.TextBox txtParPadrao1 
               BackColor       =   &H8000000E&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   15
               Text            =   "0"
               Top             =   1170
               Width           =   2175
            End
            Begin VB.TextBox txtLatOrigem 
               BackColor       =   &H8000000E&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   14
               Text            =   "0"
               Top             =   870
               Width           =   2175
            End
            Begin VB.TextBox txtOffSetX 
               BackColor       =   &H8000000E&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   13
               Text            =   "500000"
               Top             =   1770
               Width           =   2175
            End
            Begin VB.TextBox txtOffSetY 
               BackColor       =   &H8000000E&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   12
               Text            =   "10000000"
               Top             =   2070
               Width           =   2175
            End
            Begin VB.TextBox txtLongOrigem 
               BackColor       =   &H8000000E&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   11
               Text            =   "-45"
               Top             =   570
               Width           =   2175
            End
            Begin VB.TextBox txtEscala 
               BackColor       =   &H80000014&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   10
               Text            =   "0.9996"
               Top             =   2370
               Width           =   2175
            End
            Begin VB.Frame Frame4 
               Caption         =   "Hemisférios"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               TabIndex        =   5
               Top             =   2850
               Width           =   3195
               Begin VB.OptionButton rdNorteSul 
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   1
                  Left            =   1950
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   255
               End
               Begin VB.OptionButton rdNorteSul 
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   0
                  Left            =   540
                  TabIndex        =   6
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   195
               End
               Begin VB.Label Label17 
                  Caption         =   "Norte"
                  Height          =   195
                  Left            =   840
                  TabIndex        =   9
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label Label18 
                  Caption         =   "Sul"
                  Height          =   195
                  Left            =   2250
                  TabIndex        =   8
                  Top             =   240
                  Width           =   495
               End
            End
            Begin VB.Label Label14 
               Caption         =   "Par.Padrão 2:"
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   1500
               Width           =   1935
            End
            Begin VB.Label Label13 
               Caption         =   "Par.Padrão 1:"
               Height          =   195
               Left            =   120
               TabIndex        =   23
               Top             =   1200
               Width           =   1935
            End
            Begin VB.Label Label12 
               Caption         =   "Latit.Origem:"
               Height          =   225
               Left            =   120
               TabIndex        =   22
               Top             =   900
               Width           =   2055
            End
            Begin VB.Label Label7 
               Caption         =   "Zone:"
               Height          =   285
               Left            =   120
               TabIndex        =   21
               Top             =   300
               Width           =   1935
            End
            Begin VB.Label Label8 
               Caption         =   "OffSet X:"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   1800
               Width           =   1935
            End
            Begin VB.Label Label9 
               Caption         =   "OffSet Y:"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   2100
               Width           =   1935
            End
            Begin VB.Label Label10 
               Caption         =   "Long.Origem:"
               Height          =   225
               Left            =   120
               TabIndex        =   18
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label11 
               Caption         =   "Escala:"
               Height          =   315
               Left            =   120
               TabIndex        =   17
               Top             =   2370
               Width           =   1935
            End
         End
         Begin VB.ComboBox cboProjecao 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmProjection.frx":0000
            Left            =   180
            List            =   "frmProjection.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   540
            Width           =   3495
         End
         Begin VB.ComboBox cboDatum 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmProjection.frx":0004
            Left            =   180
            List            =   "frmProjection.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1170
            Width           =   3495
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            Caption         =   "Datum"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   210
            TabIndex        =   26
            Top             =   930
            Width           =   1095
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            Caption         =   "Projeção"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   210
            TabIndex        =   25
            Top             =   300
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmProjection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Confirm As Boolean

Public Function init() As Boolean
   'LoozeXP1.InitIDESubClassing
   Confirm = False
   cboProjecao.Clear
   cboProjecao.AddItem "Albers"
   cboProjecao.AddItem "Conforme de Lambert"
   cboProjecao.AddItem "LatLong"
   cboProjecao.AddItem "Mercator"
   cboProjecao.AddItem "Miller"
   cboProjecao.AddItem "Policonica"
   cboProjecao.AddItem "Sinusoidal"
   cboProjecao.AddItem "UTM"
   
   cboDatum.Clear
   cboDatum.AddItem "SAD69"
   cboDatum.AddItem "WGS84"
   cboDatum.AddItem "Córrego Alegre"
   cboDatum.AddItem "Indiano"
   cboDatum.AddItem "Astro-Chuá"
   cboDatum.AddItem "NAD27"
   cboDatum.AddItem "NAD83"
   cboDatum.AddItem "Esférico"
   cboDatum.AddItem "SIRGAS2000"
   Dim a As Integer
   For a = 1 To 60
      cboZone.AddItem a
   Next
   cboZone.Text = 23
   Me.Show vbModal
   
   init = Confirm
   'LoozeXP1.EndWinXPCSubClassing
End Function


Private Sub cboZone_Click()

   If cboZone.Text <= 30 Then
      txtLongOrigem.Text = -(((31 - cboZone.Text) * 6) - 3)
   Else
      txtLongOrigem.Text = ((cboZone.Text - 31) * 6) + 3
   End If
   
End Sub

Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   Confirm = True
   Me.Hide
End Sub

