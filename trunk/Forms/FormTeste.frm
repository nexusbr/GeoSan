VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   2430
      Width           =   1830
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   3735
      ReadOnly        =   0   'False
      TabIndex        =   0
      Top             =   1950
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
    Me.File1.path = Drive1.Drive
    
End Sub
