VERSION 5.00
Begin VB.Form frmCalibrarZoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calibrar Zoom"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   390
      Left            =   1695
      TabIndex        =   5
      Top             =   1845
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fator  1.1 a 10 "
      Height          =   1575
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   2595
      Begin VB.TextBox txtZoomMais 
         Height          =   345
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   2
         Top             =   900
         Width           =   765
      End
      Begin VB.TextBox txtZoomMenos 
         Height          =   345
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   1
         Top             =   435
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "Zoom Menos"
         Height          =   285
         Left            =   210
         TabIndex        =   4
         Top             =   510
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom Mais"
         Height          =   285
         Left            =   210
         TabIndex        =   3
         Top             =   960
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmCalibrarZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
   
   If Me.txtZoomMenos.Text <> "" And Me.txtZoomMais.Text <> "" Then
      
      Dim menos As Double
      Dim mais As Double
      
      menos = Me.txtZoomMenos.Text
      
      
      
'      If CDbl(Me.txtZoomMenos.Text) >= 1.1 And CDbl(Me.txtZoomMenos.Text) <= 10 Then
'         If CDbl(Me.txtZoomMais.Text) >= 1.1 And CDbl(Me.txtZoomMais.Text) <= 10 Then
            
            Call WriteINI("MAPA", "ZOOM_MENOS", Replace(Me.txtZoomMenos.Text, ",", "."), App.path & "\CONTROLES\GEOSAN.INI")
            Call WriteINI("MAPA", "ZOOM_MAIS", Replace(Me.txtZoomMais.Text, ",", "."), App.path & "\CONTROLES\GEOSAN.INI")
            
            dblFatorZoomMenos = CDbl(Me.txtZoomMenos.Text)
            dblFatorZoomMais = CDbl(Me.txtZoomMais.Text)
            
            Unload Me
         
'         Else
'            MsgBox "Valores inválidos.", vbInformation, ""
'         End If
'      Else
'         MsgBox "Valores inválidos.", vbInformation, ""
'
'      End If
   End If

End Sub

Private Sub Form_Load()
On Error GoTo Trata_Erro
   
   Me.txtZoomMenos.Text = Replace(ReadINI("MAPA", "ZOOM_MENOS", App.path & "\CONTROLES\GEOSAN.ini"), ",", ".")
   Me.txtZoomMais.Text = Replace(ReadINI("MAPA", "ZOOM_MAIS", App.path & "\CONTROLES\GEOSAN.ini"), ",", ".")
   
Trata_Erro:

End Sub


Private Sub txtZoomMais_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
   
   Else
      KeyAscii = 0
   End If

End Sub

Private Sub txtZoomMenos_KeyPress(KeyAscii As Integer)

   If KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
   
   Else
      KeyAscii = 0
   End If

End Sub

