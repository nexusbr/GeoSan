VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TeComDatabase.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3300
      TabIndex        =   0
      Top             =   2295
      Width           =   900
   End
   Begin TECOMDATABASELibCtl.TeDatabase TeDatabase1 
      Left            =   615
      OleObjectBlob   =   "Form1.frx":0000
      Top             =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Private Sub Command1_Click()



    TeDatabase1.Provider = 1

    'configura o componente para a conexao com o banco principal

    TeDatabase1.Connection = conn

    If TeDatabase1.setCurrentLayer("WATERLINES") Then

        'Retorna a coordenada do ponto 2 da linha 0xx999

        If TeDatabase1.getPointOfLine(0, 12253, 0, x, y) Then
            MsgBox "Coordenada do Terceiro Ponto da linha : " & CStr(x) & "_ " & CStr(y)
        End If

    End If
 'x = TeDatabase1.getPointOfLine(0, 1345, 0, x, y)
 

End Sub
