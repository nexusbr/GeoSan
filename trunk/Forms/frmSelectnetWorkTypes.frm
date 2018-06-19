VERSION 5.00
Begin VB.Form frmSelectnetWorkTypes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecione o tipo de rede a exportar"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
   Begin VB.ComboBox cboTypes 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2685
   End
End
Attribute VB_Name = "frmSelectnetWorkTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Confirm As Boolean, mtype As Long
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


Public Function init(type_id As Long) As Boolean

   Confirm = False
   Dim rs As ADODB.Recordset
a = "ID_TYPE"
b = "DESCRIPTION_"
c = "WATERLINESTYPES"


     If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute("SELECT id_type as " + """" + "ID_type" + """" + ", Description_ as " + """" + "Description_" + """" + " From waterlinesTypes")
   Else
      Set rs = Conn.execute("SELECT " + """" + a + """" + " as ID_type, " + """" + b + """" + "  as " + """" + "Description_" + """" + " From " + """" + c + """" + "")
   End If
   While Not rs.EOF
      cboTypes.AddItem rs(1).value
      cboTypes.ItemData(cboTypes.NewIndex) = rs(0).value
      rs.MoveNext
   Wend
   rs.Close
   Set rs = Nothing
   Me.Show vbModal
   init = Confirm
   type_id = mtype
End Function

Private Sub cmdOK_Click()
   If cboTypes.ListIndex >= 0 Then
      Confirm = True
      mtype = cboTypes.ItemData(cboTypes.ListIndex)
      Unload Me
   Else
      MsgBox "selecione um na caixa de seleção", vbExclamation
   End If
End Sub

