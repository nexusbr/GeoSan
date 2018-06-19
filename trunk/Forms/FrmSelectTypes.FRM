VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSelectTypes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GeoSan"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6390
   Icon            =   "FrmSelectTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   4170
      TabIndex        =   7
      Top             =   3600
      Width           =   2205
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Height          =   345
         Left            =   570
         TabIndex        =   8
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecione um tipo de item de uma rede "
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
      Height          =   705
      Left            =   30
      TabIndex        =   5
      Top             =   3600
      Width           =   4035
      Begin VB.ComboBox cboLayer 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   3675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criar Tipos e Propriedades Adicionais  P/ Rede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3555
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6345
      Begin VB.Frame fraTipo 
         Caption         =   "Tipo"
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
         Height          =   3075
         Left            =   180
         TabIndex        =   2
         Top             =   270
         Width           =   2955
         Begin MSComctlLib.ListView lvType 
            Height          =   2625
            Left            =   150
            TabIndex        =   3
            ToolTipText     =   "Click com o botão direito do mouse p/ menu"
            Top             =   270
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   4630
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame fraSubTipo 
         Caption         =   "Propriedades adicionais"
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
         Height          =   3075
         Left            =   3300
         TabIndex        =   1
         Top             =   270
         Width           =   2895
         Begin MSComctlLib.ListView lvSubType 
            Height          =   2595
            Left            =   150
            TabIndex        =   4
            ToolTipText     =   "Click com o botão direito do mouse p/ menu"
            Top             =   240
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   4577
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
   Begin VB.Menu mnuTypes 
      Caption         =   "type"
      Visible         =   0   'False
      Begin VB.Menu mnuAddType 
         Caption         =   "Adicionar"
      End
      Begin VB.Menu mnuEditType 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuRemoveType 
         Caption         =   "Excluir"
      End
   End
   Begin VB.Menu mnuSubType 
      Caption         =   "SubType"
      Visible         =   0   'False
      Begin VB.Menu mnuAddSubType 
         Caption         =   "Adicionar"
      End
      Begin VB.Menu mnuEditSubType 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuRemoveSubType 
         Caption         =   "Excluir"
      End
   End
End
Attribute VB_Name = "FrmSelectTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

Private rs As New ADODB.Recordset, itmx As ListItem


Public Function init() As Boolean
   'LoozeXP1.InitSubClassing
   LoadCboLayer
   Me.Show vbModal
End Function

Private Sub LoadCboLayer()


a = "te_layer"
b = "name"




 If frmCanvas.TipoConexao <> 4 Then
   rs.Open "SELECT * from te_layer where name in('waterlines','watercomponents','sewerlines','sewercomponents','drainlines','draincomponents')", Conn
   Else
    rs.Open "SELECT * from " + """" + a + """" + " where " + """" + b + """" + " in('waterlines','watercomponents','sewerlines','sewercomponents','drainlines','draincomponents'')", Conn, adOpenDynamic, adLockOptimistic
   End If
   While Not rs.EOF
      cboLayer.AddItem UCase(rs.Fields("name").value)
      cboLayer.ItemData(cboLayer.NewIndex) = rs.Fields("layer_id").value
      rs.MoveNext
   Wend
   rs.Close
   
   If cboLayer.ListCount > 0 Then
      cboLayer.ListIndex = 0
      LoadLvType
   End If
End Sub


Private Sub LoadLvType()
   
   
   lvSubType.ListItems.Clear
   lvType.ListItems.Clear
   lvType.ColumnHeaders.Clear
   lvType.ColumnHeaders.Add , , "", lvType.Width
   a = cboLayer.Text
b = "a"
c = "TYPES"

 If frmCanvas.TipoConexao <> 4 Then
      rs.Open "SELECT * from " & cboLayer.Text & "Types", Conn
   Else
      rs.Open "SELECT * from " + """" + cboLayer.Text + c + """" + ", Conn, adOpenDynamic, adLockOptimistic"
   End If
   While Not rs.EOF
      Set itmx = lvType.ListItems.Add(, , rs.Fields("Description_").value)
      itmx.Tag = rs.Fields("id_type").value
      rs.MoveNext
   Wend
   rs.Close
   If lvType.ListItems.count > 0 Then
      lvType.ListItems.Item(1).Selected = True
      LoadLvSubType lvType.SelectedItem.Tag
   End If
   
End Sub
Private Sub LoadLvSubType(mtype As Long)
   
   lvSubType.ListItems.Clear
   lvSubType.ColumnHeaders.Clear
   lvSubType.ColumnHeaders.Add , , "", lvSubType.Width
   a = cboLayer.Text
b = "a"
c = "SUBTYPES"
d = mtype
e = "'d'"
f = "ID_TYPE"

 If frmCanvas.TipoConexao <> 4 Then
   rs.Open "SELECT * from " & cboLayer.Text & "SubTypes where id_type = " & mtype, Conn
   Else
   rs.Open "SELECT * from " + """" + cboLayer.Text + c + """" + " where " + """" + f + """" + "='" & mtype & "'", Conn, adOpenDynamic, adLockOptimistic
   End If
   While Not rs.EOF
      Set itmx = lvSubType.ListItems.Add(, , rs.Fields("Description_").value)
      itmx.Tag = rs.Fields("id_SubType").value
      rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub cboLayer_Click()
   LoadLvType
End Sub

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rs = Nothing
   Set itmx = Nothing
   'LoozeXP1.EndWinXPCSubClassing
End Sub

Private Sub lvSubType_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not lvType.HitTest(x, y) Is Nothing Then
      If Button = 2 And Not lvSubType.HitTest(x, y) Is Nothing Then
         mnuEditSubType.Enabled = True
         mnuRemoveSubType.Enabled = True
         lvSubType.HitTest(x, y).Selected = True
         PopupMenu mnuSubType
      ElseIf Button = 2 Then
         mnuEditSubType.Enabled = False
         mnuRemoveSubType.Enabled = False
         PopupMenu mnuSubType
      End If
   End If
End Sub

Private Sub lvType_ItemClick(ByVal Item As MSComctlLib.ListItem)
   LoadLvSubType Item.Tag
End Sub

Private Sub lvType_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 And Not lvType.HitTest(x, y) Is Nothing Then
      mnuEditType.Enabled = True
      mnuRemoveType.Enabled = True
      lvType.HitTest(x, y).Selected = True
      PopupMenu mnuTypes
   ElseIf Button = 2 Then
      mnuEditType.Enabled = False
      mnuRemoveType.Enabled = False
      
      PopupMenu mnuTypes
   End If
End Sub

Private Sub mnuAddSubType_Click()
   Dim frm As New FrmSubTypes
   If Not lvType Is Nothing Then
      frm.init cboLayer.Text, lvType.SelectedItem.Tag, 0, True
      LoadLvSubType lvType.SelectedItem.Tag
   End If
   Set frm = Nothing
End Sub

Private Sub mnuAddType_Click()
   Dim frm As New FrmTypes
   frm.init 0, cboLayer.Text
   LoadLvType
   Set frm = Nothing
End Sub

Private Sub mnuEditSubType_Click()
   Dim frm As New FrmSubTypes
   If Not lvSubType Is Nothing And Not lvType Is Nothing Then
      frm.init cboLayer.Text, lvType.SelectedItem.Tag, lvSubType.SelectedItem.Tag, False
      LoadLvSubType lvType.SelectedItem.Tag
   End If
   Set frm = Nothing
End Sub

Private Sub mnuEditType_Click()
   Dim frm As New FrmTypes
   If Not lvType.SelectedItem Is Nothing Then
      frm.init lvType.SelectedItem.Tag, cboLayer.Text
      LoadLvType
   End If
   Set frm = Nothing
End Sub

Private Sub mnuRemoveSubType_Click()
   'implementar
   If Not lvSubType.SelectedItem Is Nothing And Not lvType.SelectedItem Is Nothing Then
      If MsgBox("Deseja realmente excluir a propriedade adicional deste tipo", 36) = vbYes Then
         Set rs = Conn.execute("SELECT count(*) from " & cboLayer.Text & "Data where id_type=" & lvType.SelectedItem.Tag & " and id_SubType=" & lvSubType.SelectedItem.Tag)
         If rs.Fields(0).value = 0 Then
         
a = cboLayer.Text
b = "a"
c = "SELECTIONS"
d = "ID_SUBTYPE"
e = "SUBTYPES"
f = cboLayer.Text

h = "ID_TYPE"
i = lvType.SelectedItem.Tag

k = "SUBTYPES"


            If frmCanvas.TipoConexao <> 4 Then
         
            Conn.execute "delete from " & cboLayer.Text & "SELECTions where id_type = " & lvType.SelectedItem.Tag & " and  id_subtype = " & lvSubType.SelectedItem.Tag
            Conn.execute "delete from " & cboLayer.Text & "subTypes where  id_type = " & lvType.SelectedItem.Tag & " and  id_subtype = " & lvSubType.SelectedItem.Tag
            
            Else
              
            Conn.execute "delete from " + """" + cboLayer.Text + c + """" + " where " + """" + h + """" + " =  '" & lvType.SelectedItem.Tag & "'  and  " + """" + h + """" + " = '" & lvSubType.SelectedItem.Tag & "'"
            Conn.execute "delete from " + """" + cboLayer.Text + k + """" + " where  " + """" + h + """" + " = '" & lvType.SelectedItem.Tag & "' and  " + """" + d + """" + " = '" & lvSubType.SelectedItem.Tag & "'"
            
            End If
            
            lvSubType.ListItems.Remove lvSubType.SelectedItem.index
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub mnuRemoveType_Click()
a = cboLayer.Text
b = "a"
c = "SELECTIONS"
d = "ID_TYPE"
e = "SUBTYPES"
f = cboLayer.Text

h = "ID_TYPE"
i = lvType.SelectedItem.Tag

k = "TYPES"
   If Not lvType.SelectedItem Is Nothing Then
      Set rs = Conn.execute("SELECT count(*) from " & cboLayer.Text & " where id_type =" & lvType.SelectedItem.Tag)
      If Not rs.EOF Then
         If rs(0).value = 0 Then
          If frmCanvas.TipoConexao <> 4 Then
            Conn.execute "delete from " & cboLayer.Text & "SELECTions where id_type=" & lvType.SelectedItem.Tag
            Conn.execute "delete from " & cboLayer.Text & "subtypes where id_type=" & lvType.SelectedItem.Tag
            Conn.execute "delete from " & cboLayer.Text & "types where id_type=" & lvType.SelectedItem.Tag
            Else
  
            Conn.execute "delete from " + """" + cboLayer.Text + c + """" + " where " + """" + d + """" + "='" & lvType.SelectedItem.Tag & "'"
            Conn.execute "delete from " + """" + cboLayer.Text + e + """" + " where " + """" + d + """" + "='" & lvType.SelectedItem.Tag & "'"
            Conn.execute "delete from " + """" + cboLayer.Text + k + """" + " where " + """" + h + """" + "='" & lvType.SelectedItem.Tag & "'"
            End If
         Else
            MsgBox "Não é possível excluir este tipo pois ele está relacionado um ou mais trechos da rede", vbExclamation
         End If
      End If
      rs.Close
      LoadLvType
   End If
End Sub


Public Function FunDecripta(ByVal strDecripta As String) As String


    Dim IntTam As Integer
    Dim i As Integer
    Dim letra As String
    Dim nStr As String
    IntTam = Len(strDecripta)
    nStr = ""

    'desconsidera os os numeros de HH-MM-SS
    strDecripta = mid(strDecripta, 6, 5) & mid(strDecripta, 16, 5) & mid(strDecripta, 26, 5) & _
                  mid(strDecripta, 36, 5) & mid(strDecripta, 46, 5) & mid(strDecripta, 56, 200)

    i = 1
    Do While Not i = IntTam - 29
        letra = mid(strDecripta, i, 5)
        Select Case letra
        Case "14334"
            nStr = nStr & "a"
        Case "14212"
            nStr = nStr & "A"
        Case "24334"
            nStr = nStr & "á"
        Case "24134"
            nStr = nStr & "â"
        Case "24234"
            nStr = nStr & "ã"
        Case "24314"
            nStr = nStr & "à"
        Case "24324"
            nStr = nStr & "b"
        Case "14223"
            nStr = nStr & "B"
        Case "11211"
            nStr = nStr & "ç"
        Case "11311"
            nStr = nStr & "Ç"
        Case "13334"
            nStr = nStr & "c"
        Case "14324"
            nStr = nStr & "C"
        Case "24344"
            nStr = nStr & "d"
        Case "14444"
            nStr = nStr & "D"
        Case "12314"
            nStr = nStr & "e"
        Case "21111"
            nStr = nStr & "E"
        Case "24321"
            nStr = nStr & "é"
        Case "32314"
            nStr = nStr & "ê"
        Case "31314"
            nStr = nStr & "f"
        Case "21311"
            nStr = nStr & "F"
        Case "32134"
            nStr = nStr & "g"
        Case "21341"
            nStr = nStr & "G"
        Case "31324"
            nStr = nStr & "h"
        Case "22111"
            nStr = nStr & "H"
        Case "32124"
            nStr = nStr & "i"
        Case "21112"
            nStr = nStr & "I"
        Case "31334"
            nStr = nStr & "í"
        Case "32333"
            nStr = nStr & "ì"
        Case "11314"
            nStr = nStr & "j"
        Case "23122"
            nStr = nStr & "J"
        Case "33134"
            nStr = nStr & "k"
        Case "23411"
            nStr = nStr & "K"
        Case "33314"
            nStr = nStr & "l"
       Case "32222"
            nStr = nStr & "L"
        Case "43423"
            nStr = nStr & "m"
        Case "32111"
            nStr = nStr & "M"
        Case "42423"
            nStr = nStr & "n"
        Case "33221"
            nStr = nStr & "N"
        Case "43234"
            nStr = nStr & "o"
        Case "33233"
            nStr = nStr & "O"
        Case "42444"
            nStr = nStr & "ô"
        Case "43223"
            nStr = nStr & "õ"
        Case "42433"
            nStr = nStr & "ò"
        Case "43231"
            nStr = nStr & "ó"
        Case "22223"
            nStr = nStr & "p"
        Case "33444"
            nStr = nStr & "P"
        Case "43233"
            nStr = nStr & "q"
        Case "34442"
            nStr = nStr & "Q"
        Case "43421"
            nStr = nStr & "r"
        Case "34332"
            nStr = nStr & "R"
        Case "13443"
            nStr = nStr & "s"
        Case "34222"
            nStr = nStr & "S"
        Case "44444"
            nStr = nStr & "t"
        Case "34112"
            nStr = nStr & "T"
        Case "13444"
            nStr = nStr & "u"
        Case "41311"
            nStr = nStr & "U"
        Case "11111"
            nStr = nStr & "ú"
        Case "13243"
            nStr = nStr & "ù"
        Case "11115"
            nStr = nStr & "û"
        Case "13241"
           nStr = nStr & "v"
        Case "41222"
            nStr = nStr & "V"
        Case "12443"
            nStr = nStr & "x"
        Case "41133"
            nStr = nStr & "X"
        Case "13244"
            nStr = nStr & "y"
        Case "42231"
            nStr = nStr & "Y"
        Case "13441"
            nStr = nStr & "w"
        Case "42222"
            nStr = nStr & "W"
        Case "11313"
            nStr = nStr & "z"
        Case "42213"
            nStr = nStr & "Z"
        Case "11312"
            nStr = nStr & "@"
        Case "11114"
            nStr = nStr & "%"
        Case "12341"
            nStr = nStr & "&"
        Case "13343"
            nStr = nStr & "*"
        Case "12342"
            nStr = nStr & "("
        Case "13344"
            nStr = nStr & ")"
        Case "12333"
            nStr = nStr & "$"
        Case "23334"
            nStr = nStr & "!"
        Case "13331"
            nStr = nStr & "#"
        Case "21242"
            nStr = nStr & "?"
        Case "22313"
            nStr = nStr & "1"
        Case "23424"
            nStr = nStr & "2"
        Case "24131"
            nStr = nStr & "3"
        Case "41414"
            nStr = nStr & "4"
        Case "22314"
           nStr = nStr & "5"
        Case "23423"
            nStr = nStr & "6"
        Case "44134"
            nStr = nStr & "7"
        Case "21241"
            nStr = nStr & "8"
       Case "22312"
           nStr = nStr & "9"
       Case "23231"
            nStr = nStr & "0"
        Case "34123"
            nStr = nStr & " "
        Case "14121"
            nStr = nStr & "_"
        Case "14144"
            nStr = nStr & "/"
        Case "12131"
            nStr = nStr & "\"
        Case "12124"
            nStr = nStr & "-"
        Case "21421"
            nStr = nStr & ";"
        Case "21321"
            nStr = nStr & ":"
        Case "14431"
            nStr = nStr & ","
        Case "13421"
            nStr = nStr & "."
        Case "11213"
            nStr = nStr & "+"
        Case "11212"
            nStr = nStr & "="

        Case Else
            MsgBox "Código de criptografia inválido!"
            'mStrDeCriptografa = ""
            Exit Function
        End Select
        i = i + 5
    Loop
  FunDecripta = nStr
    'mStrDeCriptografa = nStr

Exit Function
End Function



