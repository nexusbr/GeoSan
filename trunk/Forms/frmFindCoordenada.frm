VERSION 5.00
Begin VB.Form frmFindCoordenada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizar Coordenadas"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCoordY 
      Height          =   315
      Left            =   225
      TabIndex        =   2
      Top             =   1140
      Width           =   3390
   End
   Begin VB.TextBox txtCoordX 
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   465
      Width           =   3390
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Localizar"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   1665
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Coordenada Y"
      Height          =   225
      Left            =   225
      TabIndex        =   4
      Top             =   900
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "Coordenada X"
      Height          =   225
      Left            =   255
      TabIndex        =   3
      Top             =   210
      Width           =   1770
   End
End
Attribute VB_Name = "frmFindCoordenada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLocalizar_Click()
On Error GoTo Trata_Erro
    
    Dim xMin, xMax, yMin, yMax As Double
    
    'x = InputBox("Informe a Coordena X ")
    'y = InputBox("Informe a Coordena y ")
    
    If IsNumeric(Me.txtCoordX) = True And IsNumeric(Me.txtCoordY.Text) Then
        xMin = Me.txtCoordX.Text - 150
        xMax = Me.txtCoordY.Text + 150
        yMin = Me.txtCoordY.Text - 130
        yMax = Me.txtCoordY.Text + 130
        
        
        
        frmCanvas.TCanvas.setWorld xMin, yMin, xMax, yMax
        
        frmCanvas.TCanvas.plotView
    
        Unload Me
    Else
        MsgBox "Insira valores numéricos.", vbInformation, ""
    End If
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    ElseIf Err.Number = 91 Or Err.Number = 13 Then
        MsgBox "Não há mapa ativo.", vbInformation, "Geosan"
    Else
       Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
       Print #1, Now & " - Private Sub cmdLocalizar_Click - " & Err.Number & " - " & Err.Description
       Close #1
       MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
    End If
End Sub
