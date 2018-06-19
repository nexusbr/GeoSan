VERSION 5.00
Begin VB.Form frmAlteraPorSelecao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alterações por Seleção"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   2490
      TabIndex        =   2
      Top             =   3465
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   405
      Left            =   3690
      TabIndex        =   1
      Top             =   3465
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modificar Informações dos Locais Selecionados"
      Height          =   3150
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   4620
   End
End
Attribute VB_Name = "frmAlteraPorSelecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
