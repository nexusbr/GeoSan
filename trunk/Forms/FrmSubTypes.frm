VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSubTypes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Propriedade Adcional do Tipo"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5115
   Icon            =   "FrmSubTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Ok"
      Height          =   285
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Salvar SubTipo"
      Top             =   4710
      Width           =   885
   End
   Begin VB.Frame frmSelecoes 
      Caption         =   "Ítens de seleção"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2895
      Left            =   30
      TabIndex        =   10
      Top             =   1770
      Width           =   5055
      Begin VB.CommandButton cmdDel 
         Caption         =   "Excluir"
         Height          =   285
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Salvar SubTipo"
         Top             =   2400
         Width           =   885
      End
      Begin VB.CommandButton cmdRename 
         Caption         =   "Renomear"
         Height          =   285
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Salvar SubTipo"
         Top             =   2400
         Width           =   885
      End
      Begin VB.TextBox txtOption 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Adicionar"
         Height          =   345
         Left            =   3990
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Salvar SubTipo"
         Top             =   600
         Width           =   885
      End
      Begin MSComctlLib.ListView LvSeletions 
         Height          =   1245
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   2196
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   8387
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtDescription 
         Height          =   315
         Left            =   1620
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lblOpcaoSelect 
         AutoSize        =   -1  'True
         Caption         =   "Opção"
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
         TabIndex        =   12
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblDescrSelect 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
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
         Left            =   1620
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   870
      End
   End
   Begin VB.Frame FraSubTypes 
      Caption         =   "Propriedades do Tipo: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   5055
      Begin VB.TextBox txtValue 
         Height          =   315
         Left            =   3870
         TabIndex        =   17
         Top             =   510
         Width           =   1005
      End
      Begin VB.TextBox txtMax 
         Height          =   315
         Left            =   3030
         TabIndex        =   16
         Top             =   510
         Width           =   765
      End
      Begin VB.TextBox txtMin 
         Height          =   315
         Left            =   2190
         TabIndex        =   15
         Top             =   510
         Width           =   765
      End
      Begin VB.TextBox txtDescription_ 
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   510
         Width           =   1965
      End
      Begin VB.ComboBox cboTipoDado 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1110
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Seleção"
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
         Height          =   600
         Left            =   2160
         TabIndex        =   5
         Top             =   900
         Width           =   1635
         Begin VB.OptionButton optNao 
            Caption         =   "Não"
            Height          =   255
            Left            =   720
            TabIndex        =   1
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optSim 
            Caption         =   "Sim"
            Height          =   255
            Left            =   90
            TabIndex        =   0
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de dado"
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
         Left            =   150
         TabIndex        =   9
         Top             =   870
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor Padrão"
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
         Left            =   3900
         TabIndex        =   8
         Top             =   270
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mínimo"
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
         Left            =   2220
         TabIndex        =   7
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Máximo"
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
         Left            =   3060
         TabIndex        =   6
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
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
         TabIndex        =   4
         Top             =   270
         Width           =   870
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "Selecao"
      Visible         =   0   'False
      Begin VB.Menu mnuRename 
         Caption         =   "Renomear"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Remover"
      End
   End
End
Attribute VB_Name = "FrmSubTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlayer As String
Private mtype As Integer
Private mSubType As Integer
Private mConfirm As Boolean
Private mNew As Boolean
Private i As ListItem


Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim ii As String
Dim j As String
Dim k As String
Dim l As String
Dim m As String
Dim n As String
Dim o As String
Dim z As String


Public Function init(LayerName As String, id_Type As Integer, Id_SubType As Integer, mNew_ As Boolean)
   'LoozeXP1.InitSubClassing
   mlayer = LayerName
   mtype = id_Type
   mSubType = Id_SubType
   mNew = mNew_
      'preenche combo de tipos de dado
   With cboTipoDado
      .AddItem "Float"
      .ItemData(cboTipoDado.NewIndex) = 5
      .AddItem "Char"
      .ItemData(cboTipoDado.NewIndex) = 129
      .AddItem "VarChar"
      .ItemData(cboTipoDado.NewIndex) = 200
      .AddItem "nChar"
      .ItemData(cboTipoDado.NewIndex) = 130
      .AddItem "nVarChar"
      .ItemData(cboTipoDado.NewIndex) = 202
      .AddItem "Bit"
      .ItemData(cboTipoDado.NewIndex) = 11
      .AddItem "TinyInt"
      .ItemData(cboTipoDado.NewIndex) = 17
      .AddItem "SmallInt"
      .ItemData(cboTipoDado.NewIndex) = 2
      .AddItem "Int"
      .ItemData(cboTipoDado.NewIndex) = 3
      .AddItem "Decimal"
      .ItemData(cboTipoDado.NewIndex) = 13
      .AddItem "DateTime"
      .ItemData(cboTipoDado.NewIndex) = 135
   End With

   LoadSubType
   Me.Show vbModal
   'LoozeXP1.EndWinXPCSubClassing
   init = mConfirm
End Function


Public Sub LoadSubType()
   Dim rs As New ADODB.Recordset
   Dim SQL As String

a = mlayer
b = "a"
c = "SUB_TYPES"
d = "ID_TYPE"
e = "ID_SUBTYPE"
f = mtype

i = mSubType


   If Not mNew Then
      
      'preenche campos de subtipos
      
     If frmCanvas.TipoConexao <> 4 Then
      SQL = "SELECT * from " & mlayer & "SubTypes where id_Type = " & mtype & " and id_subType = " & mSubType
      Else
      SQL = "SELECT * from " + """" + mlayer & c + """" + " where " + d + "= '" & mtype & "' And " + e + " = '" & mSubType & "'"
      End If
      
      Set rs = Conn.execute(SQL)
      If Not rs.EOF Then
         txtDescription_.Text = rs.Fields("description_").value
         'cboTipoDado.Text = cboTipoDado.li .ItemData(cboTipoDado.Text)
         If rs.Fields("min_").value <> "" Then txtMin.Text = rs.Fields("min_").value
         If rs.Fields("max_").value <> "" Then txtMax.Text = rs.Fields("max_").value
         If rs!Selection_ = True Then
            optSim.value = True
         Else
            optNao.value = True
         End If
         txtValue.Text = rs.Fields("defaultValue").value
         cboTipoDado.ListIndex = GetCboListIndex(rs.Fields("DataType").value, cboTipoDado)
         
      End If
      rs.Close
      
a = mlayer
b = "a"
c = "SELECTIONS"
d = "ID_TYPE"
e = "ID_SUBTYPE"
f = mtype

i = mSubType

If frmCanvas.TipoConexao <> 4 Then
      SQL = "SELECT * from " & mlayer & "SELECTions where id_Type = " & mtype & " and id_subType = " & mSubType
      Else
      SQL = "SELECT * from " + """" + mlayer & c + """" + " where " + """" + d + """" + " = '" & mtype & "' and " + """" + e + """" + " = '" & mSubType & "'"
      End If
      
      Set rs = Conn.execute(SQL)
         While Not rs.EOF
            Set i = LvSeletions.ListItems.Add(, , rs.Fields("option_").value)
            If rs.Fields("description_").value <> Null Then i.SubItems(1) = rs.Fields("description_").value
            i.Tag = rs.Fields("value_").value
            rs.MoveNext
         Wend
      rs.Close
   
   End If
   Set rs = Nothing
End Sub

Private Sub cmdDel_Click()
   mnuDel_Click
End Sub

Private Sub cmdRename_Click()
   mnuRename_Click
End Sub

'salvar subtipo
Private Sub cmdSalvar_Click()
   Dim rs As New ADODB.Recordset
   Dim SQL As String
   Dim a As Integer
   Dim PValue As Integer

   
   'insere dados de subtipo e seleção
   If mNew Then

       c = "ID_TYPE"
      d = "ID_SUBTYPE"
        e = "SUBTYPES"
    
    

     If frmCanvas.TipoConexao <> 4 Then
         
      rs.Open "SELECT max(id_subtype) from " & mlayer & "subtypes where id_type = " & mtype, Conn
      Else
       rs.Open "SELECT max(" + """" + d + """" + ") from " + """" + mlayer & e + """" + " where " + """" + c + """" + " = '" & mtype & "'", Conn, adOpenDynamic, adLockOptimistic
      End If
      
      If rs.EOF Then
        mSubType = 0
      Else
        mSubType = IIf(IsNull(rs(0).value), 0, rs(0).value) + 1
      End If
      rs.Close
      
             
      b = "SUBTYPES"
      c = "ID_TYPE"
      d = "ID_SUBTYPE"
      e = "DESCRIPTION_"
      f = "SELECTION_"
      g = "MIN_"
      h = "MAX_"
      ii = "DEFAULTVALUE"
      j = "DATATYPE"

     If frmCanvas.TipoConexao <> 4 Then
         
     SQL = "Insert into " & mlayer & "SubTypes " & "(id_type, id_subtype, description_, SELECTion_, min_, max_, defaultvalue, datatype) " & _
         "Values (" & mtype & "," & mSubType & ",'" & txtDescription_.Text & "'," & IIf(optSim, 1, 0) & "," & IIf(txtMin.Text = "", 0, txtMin.Text) & "," & IIf(txtMax.Text = "", 0, txtMax.Text) & ",'" & IIf(txtValue.Text = "", 0, txtValue.Text) & "'," & cboTipoDado.ItemData(cboTipoDado.ListIndex) & ")"
     Else
     
      SQL = "Insert into " + """" + mlayer & b + """" + " (" + """" + c + """" + ", " + """" + d + """" + ", " + """" + e + """" + ", " + """" + f + """" + ", " + """" + g + """" + ", " + """" + h + """" + ", " + """" + ii + """" + ", " + """" + j + """" + ") " & _
         "Values ('" & mtype & "','" & mSubType & "','" & txtDescription_.Text & "','" & IIf(optSim, 1, 0) & "','" & IIf(txtMin.Text = "", 0, txtMin.Text) & "','" & IIf(txtMax.Text = "", 0, txtMax.Text) & "','" & IIf(txtValue.Text = "", 0, txtValue.Text) & "','" & cboTipoDado.ItemData(cboTipoDado.ListIndex) & "')"
     End If
      
    
      Conn.execute SQL
      
      'insere seleção
      If optSim.value = True Then
         With LvSeletions
            For a = 1 To .ListItems.count
              If Not .ListItems.Item(a).Tag = "" Then
          c = "ID_TYPE"
      d = "ID_SUBTYPE"
      a = mlayer
      b = "a"
      e = "SELECTIONS"
      g = mSubType
     
      i = "VALUE_"
      
     
          If frmCanvas.TipoConexao <> 4 Then
            Set rs = Conn.execute("SELECT max(value_) as " + """" + "Proximo" + """" + " From " & mlayer & "SELECTions Where Id_Type = " & mtype & " and Id_SubType = " & mSubType)
             
                 Else
                  Set rs = Conn.execute("SELECT max(" + """" + i + """" + ") as " + """" + "Proximo" + """" + " From " + """" + mlayer & e + """" + " Where " + """" + c + """" + "= '" & mtype & "' and " + """" + d + """" + " = '" & mSubType & "'")
                 End If
                    If rs.EOF Then
                       PValue = rs!proximo + 1
                    Else
                       PValue = 1
                    End If
                 rs.Close
                 Set rs = Nothing
               
                b = "SELECTions"
      c = "ID_TYPE"
      d = "ID_SUBTYPE"
      e = "DESCRIPTION_"
      f = "SELECTIONS"
      g = "MIN_"
      h = "MAX_"
       ii = "Value_"
       j = "OPTION"



      If frmCanvas.TipoConexao <> 4 Then
          SQL = "Insert into " & mlayer & "SELECTions (id_type, id_subtype, option_, description_,Value_) " & _
                     " Values (" & mtype & " , " & mSubType & ", '" & .ListItems(a).Text & "','" & .ListItems(a).SubItems(1) & "'," & PValue & ")"
    
     Else
     
      SQL = "Insert into " + """" + mlayer & f + """" + " (" + """" + c + """" + ", " + """" + d + """" + ", " + """" + j + """" + ", " + """" + e + """" + ", " + """" + ii + """" + ") " & _
        " Values ('" & mtype & "' , '" & mSubType & "', '" & .ListItems(a).Text & "','" & .ListItems(a).SubItems(1) & "','" & PValue & "')"
     End If
                
                     
                 Conn.execute (SQL)
   
              End If
            Next
         End With
      End If
   Else
      b = "SUBTYPES"
      c = "DESCRIPTION_"
      d = cboTipoDado.ItemData(cboTipoDado.ListIndex)
      e = "DESCRIPTION_"
      f = "SELECTION_"
      g = IIf(optSim, 1, 0)
      h = "MAX_"
      ii = "MIN_"
      j = "OPTION"
      k = "DEFAULTVALUE"
            l = "DATATYPE"
                  m = "ID_TYPE"
                  a = "ID_SUBTYPE"
                  z = "'g'"
                  Dim zz As String
                  zz = mlayer + b + " set " + c
      If frmCanvas.TipoConexao <> 4 Then
      SQL = "update '" & mlayer & "' SubTypes " & "set description_ = '" & txtDescription_ & _
            "', SELECTion_= " & IIf(optSim, 1, 0) & _
            " , min_ = " & txtMin & _
            " , max_ = " & txtMax & _
            " , defaultvalue = " & txtValue & _
            " , datatype= " & cboTipoDado.ItemData(cboTipoDado.ListIndex) & _
            "   where id_type = " & mtype & " and id_subtype =" & mSubType
         Else
         SQL = "update  " + """" + mlayer + "SUBTYPES" + """" + " set " + """" + c + """" + " = '" & txtDescription_ & "', " + """" + f + """" + "= '" + """" + g + """" + "' ," + """" + ii + """" + " = '" & txtMin & "' , " + """" + h + """" + " = '" & txtMax & "' , " + """" + k + """" + " = '" & txtValue & "' , " + """" + l + """" + "= '" + """" + d + """" + "' where " + """" + m + """" + "= '" & mtype & "' and " + """" + a + """" + " = '" & mSubType & "'"
         
         End If
     
      Conn.execute SQL
       b = "VALUE_"
      c = "SELECTIONS"
      d = "ID_SUBTYPE"
       a = "ID_TYPE"
       e = mlayer
       f = "e"
      If optSim.value = True Then
         With LvSeletions
            For a = 1 To .ListItems.count
              If .ListItems.Item(a).Tag = "" Then
              
               If frmCanvas.TipoConexao <> 4 Then
                 Set rs = Conn.execute("SELECT max(value_) as " + """" + "Proximo" + """" + " From " & mlayer & "SELECTions Where Id_Type = " & mtype & " and Id_SubType = " & mSubType)
                 
                 Else
                 Set rs = Conn.execute("SELECT max(" + """" + b + """" + ") as " + """" + "Proximo" + """" + " From " + """" + mlayer + c + """" + " Where " + """" + a + """" + " = '" & mtype & "' and " + """" + d + """" + " = '" & mSubType & "'")
                 End If
                    If Not rs.EOF Then
                       If IsNull(rs!proximo) Then
                          PValue = 0
                       Else
                          PValue = Val(rs!proximo) + 1
                       End If
                    Else
                       PValue = 1
                    End If
                 rs.Close
                 Set rs = Nothing
  


              a = "SELECTIONS"
      b = "ID_TYPE"
      c = "ID_SUBTYPE"
      d = "OPTION_"
      e = "DESCRIPTION_"
      f = "VALUE"
      



     If frmCanvas.TipoConexao <> 4 Then
         
     SQL = "Insert into " & mlayer & "SELECTions (id_type, id_subtype, option_, description_,Value_) " & _
                     " Values (" & mtype & " , " & mSubType & ", '" & .ListItems(a).Text & "','" & .ListItems(a).SubItems(1) & "'," & PValue & ")"
                     
     Else
     
      SQL = "INSERT INTO " + """" + mlayer + a + """" + " (" + """" + b + """" + "," + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + "," + """" + e + """" + "," + """" + f + """" + ") "
     SQL = SQL & " Values ('" & mtype & "' , '" & mSubType & "', '" & .ListItems(a).Text & "','" & .ListItems(a).SubItems(1) & "','" & PValue & "')"
     End If
               
                 
                     
                 Conn.execute (SQL)
               Else
      b = "SUBTYPES"
      c = "OPTION_"
      d = "DESCRIPTION_"
      e = "ID_TYPE"
      f = "ID_SUBTYPE"
      g = "VALUE_"
      a = mlayer
      h = "a"
      i = "SELECTIONS"
               If frmCanvas.TipoConexao <> 4 Then
                 SQL = " update " & mlayer & "SELECTions set " & _
                       " option_='" & .ListItems.Item(a).Text & "'," & _
                       " description_='" & .ListItems.Item(a).SubItems(1) & "'" & _
                       " " & _
                       " where id_type = " & mtype & " and id_SubType =" & mSubType & " and value_='" & .ListItems.Item(a).Tag & "'"
                       Else
                       SQL = " update " + """" + mlayer + i + """" + " set " + """" + _
                       c + """" + "= '" & .ListItems.Item(a).Text & "'," & _
                        "+ """" +d+ """" +='" & .ListItems.Item(a).SubItems(1) & "'" & _
                       " " & _
                       " where + """" + e+ """" + = '" & mtype & "' and " + """" + f + """" + " ='" & mSubType & "' and " + """" + g + """" + " ='" & .ListItems.Item(a).Tag & "'"
                       End If
                 Conn.execute SQL
              
              End If
            Next
         End With
      Else
         'Sql = "delete from " & mlayer & "SELECTions " & _
               "where id_type = " & mtype & " and id_subtype = " & mSubType & " "
         Conn.execute (SQL)
      
      End If
      
   
   End If
   
   mConfirm = True
   Unload Me
   
   
End Sub



Private Sub cmdAdd_Click()
    
   Set i = LvSeletions.ListItems.Add(, , txtOption)
   i.SubItems(1) = txtDescription
   txtOption = ""
   txtDescription = ""
               
End Sub

Private Sub LvSeletions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not LvSeletions.HitTest(x, y) Is Nothing Then
      LvSeletions.HitTest(x, y).Selected = True
      If Button = 2 Then PopupMenu mnuSelect
   End If
End Sub


Private Sub mnuDel_Click()
   On Error GoTo mnuDel_Click_err
   Dim SQL As String
   If Not LvSeletions.SelectedItem Is Nothing Then
      If LvSeletions.SelectedItem.Tag <> "" Then
 b = mlayer
      
      d = "SELECTIONS"
      e = "ID_TYPE"
      f = "ID_SUBTYPE"
      g = "VALUE_"
      a = mlayer
      
      i = "SELECTIONS"
      j = mSubType
     
               If frmCanvas.TipoConexao <> 4 Then
         SQL = "delete from " & mlayer & "SELECTions " & _
                  "where id_type = " & mtype & " and id_subtype = " & mSubType & " and value_='" & LvSeletions.SelectedItem.Tag & "'"
                  Else
                  SQL = "delete from " + """" + mlayer + d + """" + "where " + """" + e + """" + " = '" & mtype & "' and " + """" + f + """" + " = '" & mSubType & "' and " + """" + g + """" + "='" & LvSeletions.SelectedItem.Tag & "'"
                  End If
                  
                  
         Conn.execute (SQL)
      End If
      LvSeletions.ListItems.Remove LvSeletions.SelectedItem.index
   End If
   Exit Sub
mnuDel_Click_err:
   
End Sub

Private Sub mnuRename_Click()
   LvSeletions.StartLabelEdit
End Sub

Private Sub optNao_Click()
   If optNao.value Then
      frmSelecoes.Enabled = False
      If LvSeletions.ListItems.count > 0 Then
         MsgBox "Os itens de seleçao serão excluidos", vbInformation
      End If
   End If
End Sub

Private Sub optSim_Click()
   If optSim.value Then
      frmSelecoes.Enabled = True
   End If
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii = 44 Or KeyAscii = 46) Or (KeyAscii >= 48 And KeyAscii <= 57)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii = 44 Or KeyAscii = 46) Or (KeyAscii >= 48 And KeyAscii <= 57)) Then
      KeyAscii = 0
   End If
End Sub

'verifica se os campos estão preenchidos
Private Function verificaCampoSubTipo() As Boolean
Dim mensage As String

   If txtDescription.Text = "" Then mensage = mensage & "Descrição" & vbCrLf
   If txtValue.Text = "" Then mensage = mensage & "Valor Default" & vbCrLf
   If optSim.value = False And optNao.value = False Then mensage = mensage & "Opção de Seleção" & vbCrLf
   If txtMin.Text = "" Then mensage = mensage & "Valor Mínimo" & vbCrLf
   If txtMax.Text = "" Then mensage = mensage & "Valor Máximo" & vbCrLf
   If cboTipoDado.Text = "" Then mensage = mensage & "Tipo de Dado " & vbCrLf
   If optSim = True Then
      If LvSeletions.ListItems.count <= 0 Then mensage = mensage & "Insira um Tipo de Seleção" & vbCrLf
   End If
   
   If mensage <> "" Then
      verificaCampoSubTipo = False
      MsgBox "Os campos abaixo devem ser preenchidos: " & vbCrLf & vbCrLf & mensage, vbInformation, "GeoSan"
   Else
      verificaCampoSubTipo = True
   End If

End Function



