VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOdometro2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "                      "
      ForeColor       =   &H80000008&
      Height          =   1890
      Left            =   4035
      TabIndex        =   3
      Top             =   585
      Width           =   1275
      Begin MSComctlLib.ProgressBar caixa2 
         Height          =   1755
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   3096
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Max             =   100
         Orientation     =   1
         Scrolling       =   1
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3135
      Top             =   1845
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2550
      Top             =   1830
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "                      "
      ForeColor       =   &H80000008&
      Height          =   1890
      Left            =   675
      TabIndex        =   0
      Top             =   585
      Width           =   1275
      Begin MSComctlLib.ProgressBar caixa1 
         Height          =   1755
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   3096
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Max             =   100
         Orientation     =   1
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.ProgressBar Tubo 
      Height          =   150
      Left            =   1905
      TabIndex        =   2
      Top             =   2325
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1
      Max             =   10
   End
End
Attribute VB_Name = "frmOdometro2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Timer1.Enabled = True
    Timer2.Enabled = True

End Sub

Private Sub Form_Load()

caixa1.Max = 1000
caixa1.Value = 1000
caixa2.Max = 1000
caixa2.Value = 1
Tubo.Max = 10
Tubo.Value = 1
End Sub

Private Sub Timer1_Timer()
    
    caixa1.Value = caixa1.Value - 1
    caixa2.Value = caixa2.Value + 1

End Sub

Private Sub Timer2_Timer()
    
    If Tubo.Value >= 10 Then
        Tubo.Value = 1
    Else
        Tubo.Value = Tubo.Value + 1
    End If
    
End Sub
