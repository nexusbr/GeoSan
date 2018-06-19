VERSION 5.00
Begin VB.Form frmAutoLoginCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login Automático"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Definir login automático para o usuário"
      Height          =   975
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   5145
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   390
         Left            =   3915
         TabIndex        =   2
         Top             =   390
         Width           =   1005
      End
      Begin VB.TextBox txtNomeUsuario 
         Height          =   360
         Left            =   225
         TabIndex        =   1
         Top             =   390
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmAutoLoginCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


