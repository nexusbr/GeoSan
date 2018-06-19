VERSION 5.00
Begin VB.Form FrmTypes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Tipos"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   2850
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirmation 
      Caption         =   "Confirmar"
      Height          =   315
      Left            =   1770
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtSpecification_ 
      Height          =   315
      Left            =   1590
      TabIndex        =   3
      Top             =   720
      Width           =   2235
   End
   Begin VB.TextBox txtDescription_ 
      Height          =   315
      Left            =   1590
      TabIndex        =   0
      Top             =   210
      Width           =   2235
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Especificação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   270
      Width           =   930
   End
End
Attribute VB_Name = "FrmTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ID As Long, LayerName As String
Dim rs As New ADODB.Recordset, Confirmou As Boolean
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String
Dim k As String
Dim l As String




Public Function init(mid As Long, mlayerName As String) As Boolean
   'LoozeXP1.InitSubClassing
   ID = mid
   LayerName = mlayerName

   
    If frmCanvas.TipoConexao <> 4 Then
         
    If ID > 0 Then
a = LayerName
b = "a"
c = "TYPES"
d = "ID_TYPES"

      rs.Open "SELECT * From " & LayerName & " Types where id_type=" & ID, Conn
     
      txtDescription_ = rs!DESCRIPTION_
      txtSpecification_ = IIf(IsNull(rs!Specification_), "", rs!Specification_)
      rs.Close
   Else
      txtDescription_ = ""
      txtSpecification_ = ""
   End If
   
     
     Else
     
   If ID > 0 Then
   
         rs.Open "SELECT * From " + """" + LayerName & c + """" + " where " + """" + d + """" + "= '" & ID & "'", Conn, adOpenDynamic, adLockOptimistic
      txtDescription_ = rs!DESCRIPTION_
      txtSpecification_ = IIf(IsNull(rs!Specification_), "", rs!Specification_)
      rs.Close
   Else
      txtDescription_ = ""
      txtSpecification_ = ""
   End If
   
   End If
  
   Me.Show vbModal
   'LoozeXP1.EndWinXPCSubClassing
End Function

Private Sub cmdCancel_Click()
   Set rs = Nothing
   Confirmou = False
   Unload Me
End Sub

Private Sub cmdConfirmation_Click()
a = LayerName
b = "a"
c = "TYPES"
d = "ID_TYPE"
   On Error GoTo cmdConfirmation_Click_err
     If frmCanvas.TipoConexao <> 4 Then
   If ID > 0 Then
      rs.Open "SELECT * From " & LayerName & "Types where id_type=" & ID, Conn, adOpenKeyset, adLockOptimistic
      rs!DESCRIPTION_ = txtDescription_
      rs!Specification_ = txtSpecification_
   Else
      rs.Open LayerName & "Types", Conn, adOpenKeyset, adLockOptimistic
      rs.AddNew
      rs!DESCRIPTION_ = txtDescription_
      rs!Specification_ = txtSpecification_
   End If
   Else
   If ID > 0 Then
    rs.Open "SELECT * From " + """" + LayerName & c + """" + "  where " + """" + d + """" + "='" & ID & "'", Conn, adOpenDynamic, adLockOptimistic
      rs!DESCRIPTION_ = txtDescription_
      rs!Specification_ = txtSpecification_
   Else
      rs.Open LayerName & "Types", Conn, adOpenDynamic, adLockOptimistic
      rs.AddNew
      rs!DESCRIPTION_ = txtDescription_
      rs!Specification_ = txtSpecification_
   End If
   
   End If
   rs.Update
   rs.Close
   Set rs = Nothing
   Confirmou = True
   Unload Me
   Exit Sub
cmdConfirmation_Click_err:
   MsgBox Err.Description, vbExclamation
End Sub

