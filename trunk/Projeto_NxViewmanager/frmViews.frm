VERSION 5.00
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TECOMC~1.DLL"
Object = "{F03ABD98-7B60-43E4-9934-DA5F0D19FDAC}#1.0#0"; "TeComViewManager.dll"
Object = "{EE78E37B-39BE-42FA-80B7-E525529739F7}#1.0#0"; "TECOMV~2.DLL"
Begin VB.Form frmViews 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleção de vista"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   795
   End
   Begin VB.CommandButton cmdDeleteView 
      Caption         =   "Excluir"
      Height          =   315
      Left            =   1973
      TabIndex        =   3
      Top             =   600
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   795
   End
   Begin VB.ComboBox cboViews 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Text            =   "Nova Vista"
      Top             =   210
      Width           =   3675
   End
   Begin TeComViewDatabaseLibCtl.TeViewDatabase TeViewDatabase1 
      Left            =   1680
      OleObjectBlob   =   "frmViews.frx":0000
      Top             =   0
   End
   Begin TECOMVIEWMANAGERLibCtl.TeViewManager TeViewManager1 
      Left            =   2160
      OleObjectBlob   =   "frmViews.frx":0024
      Top             =   120
   End
   Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
      Left            =   3120
      OleObjectBlob   =   "frmViews.frx":0048
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   "Vista:"
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Confirm As Boolean, tvw As Object, tcs As Object, man As Object, data As Object
Private mViews_name() As String
Private Icount As Integer

Public Function Init(mtvw As Object, tcs2 As Object, man2 As Object, data2 As Object) As Boolean
   'LoozeXP1.InitSubClassing
   Init = False
   Set tvw = mtvw
   Set tcs = tcs2
   Set man = man2
   Set data = data2
   LoadView
   Me.Show vbModal
   Unload Me
   Init = Confirm
End Function
Private Sub LoadView()
   Dim a As Integer, rs As ADODB.Recordset
   ReDim mViews_name(0) As String
   Icount = -1
   Dim aa As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
aa = "te_view"
b = "user_name"
c = "name"

    If TypeConn <> 4 Then
   Set rs = conn.Execute("Select * From Te_View where user_name='" & tvw.UserName & "' order by name")
   Else
   Set rs = conn.Execute("Select * From " + """" + aa + """" + " where " + """" + b + """" + "='" & tvw.UserName & "' order by " + """" + c + """" + "")
   End If
   
   While Not rs.EOF
      cboViews.AddItem rs.Fields("name").Value
      ReDim Preserve mViews_name(a) As String
      mViews_name(a) = rs.Fields("name").Value
      Icount = a
      a = a + 1
      rs.MoveNext
   Wend
 
   rs.Close
   Set rs = Nothing
End Sub

Private Sub cmdCancelar_Click()

    Unload Me

End Sub

Private Sub cmdDeleteView_Click()
   Dim a As Integer
   
   Dim vistas As Integer
   vistas = Icount
  ' mViews_name2 = mViews_name
   
   For a = 0 To vistas  'mtvw.getViewCount() - 1
      'If mtvw.getViewName(A) = cboViews.Text Then
      If mViews_name(a) = cboViews.Text Then
         If MsgBox("Tem certeza que deseja excluir a vista: " & cboViews, 36) = vbYes Then
            tvw.removeView cboViews.Text
            'cboViews.RemoveItem cboViews.ListIndex
            cboViews.Text = ""
            
         End If
      End If
   Next
  cboViews.Clear
  
   LoadView
   
End Sub

Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
  
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 
 Dim retlen As String
 Dim Ret As String
 
 Ret = String$(255, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
 Ret = Left$(Ret, retlen)
 ReadINI = Ret

End Function


Private Sub cmdOK_Click()
   On Error GoTo cmdOk_sair
   Dim LayerName As String, ThemeName As String
   
   

   
   If cboViews.Text <> "" Then
      If Not GetView(cboViews.Text) Then
         If MsgBox("Deseja criar uma vista com o nome de: " & cboViews.Text, 36) = vbYes Then
    
             If FrmLayerTheme.Init(tvw, LayerName, ThemeName) Then
               tvw.addView cboViews.Text
               tvw.setActiveView cboViews.Text
               'FrmLayerTheme.Temas (ThemeName)
              
             
              
              
              
              
              
              
               tcs.ResetView
               If tvw.addTheme(LayerName, cboViews.Text, ThemeName) Then
               tcs.ResetView
               frmTheme.Init tvw, ThemeName, LayerName
               
               End If
               Me.Hide
               Confirm = True
            End If
         End If
      Else
         Screen.MousePointer = vbHourglass
         tvw.setActiveView cboViews.Text
         Confirm = True
         Me.Hide
         Screen.MousePointer = vbNormal
      End If
   Else
      MsgBox "É necessário setar uma vista, selecione ou crie uma nova", vbExclamation
   End If
  
   Exit Sub
cmdOk_sair:
   MsgBox Err.Description, vbExclamation
End Sub

Private Function GetView(mView As String) As Boolean
   Dim a As Integer
   For a = 0 To Icount
      If mViews_name(a) = cboViews.Text Then GetView = True
   Next
End Function

Private Sub Form_Unload(Cancel As Integer)
   'LoozeXP1.EndWinXPCSubClassing
End Sub

