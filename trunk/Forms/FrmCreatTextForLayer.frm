VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TECOMD~1.DLL"
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TECOMC~1.DLL"
Begin VB.Form FrmCreatTextForLayer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cria texto para o Plano"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLayer 
      Height          =   315
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   150
      Width           =   3285
   End
   Begin VB.Frame Frame4 
      Caption         =   "Atributos"
      Height          =   2145
      Left            =   120
      TabIndex        =   13
      Top             =   540
      Width           =   3855
      Begin VB.CommandButton cmdRemover 
         Caption         =   "<"
         Height          =   345
         Left            =   1770
         TabIndex        =   19
         Top             =   1290
         Width           =   315
      End
      Begin VB.CommandButton CmdInserir 
         Caption         =   ">"
         Height          =   345
         Left            =   1770
         TabIndex        =   18
         Top             =   660
         Width           =   315
      End
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   2130
         TabIndex        =   17
         Top             =   240
         Width           =   1635
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Posição do texto em relação ao objeto"
      Height          =   585
      Left            =   150
      TabIndex        =   9
      Top             =   3390
      Width           =   3825
      Begin VB.OptionButton optEnd 
         Caption         =   "Fim"
         Height          =   255
         Left            =   2700
         TabIndex        =   12
         Top             =   300
         Width           =   795
      End
      Begin VB.OptionButton optCenter 
         Caption         =   "Centro"
         Height          =   285
         Left            =   1530
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optInit 
         Caption         =   "Início"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   300
         Width           =   975
      End
      Begin TECOMDATABASELibCtl.TeDatabase DB 
         Left            =   3600
         OleObjectBlob   =   "FrmCreatTextForLayer.frx":0000
         Top             =   360
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecione a Geometria"
      Height          =   615
      Left            =   150
      TabIndex        =   5
      Top             =   2730
      Width           =   3825
      Begin VB.OptionButton optPoints 
         Caption         =   "Pontos"
         Height          =   255
         Left            =   2700
         TabIndex        =   8
         Top             =   270
         Width           =   795
      End
      Begin VB.OptionButton optLines 
         Caption         =   "Linhas"
         Height          =   285
         Left            =   1530
         TabIndex        =   7
         Top             =   270
         Width           =   855
      End
      Begin VB.OptionButton optPolygons 
         Caption         =   "Poligon."
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Separador de campos"
      Height          =   585
      Left            =   150
      TabIndex        =   2
      Top             =   4020
      Width           =   3825
      Begin VB.OptionButton optHifem 
         Caption         =   "Hífem"
         Height          =   195
         Left            =   690
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton OptSpace 
         Caption         =   "Espaço"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   270
         Width           =   885
      End
      Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
         Left            =   3360
         OleObjectBlob   =   "FrmCreatTextForLayer.frx":0024
         Top             =   240
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   3030
      TabIndex        =   1
      Top             =   4710
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Confimar"
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   4710
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Plano:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   180
      Width           =   2235
   End
End
Attribute VB_Name = "FrmCreatTextForLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private cgeo As New clsGeoReference
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
Function init() As Boolean

   Dim a As Integer
   'carregar Plano
   db.Provider = typeconnection
   'DB.Connection = Conn
   
   
   
   
   Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String


mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
frmCanvas.FunDecripta (Senha)
decriptada = frmCanvas.Senha


 TeAcXConnection1.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA
 
   db.connection = TeAcXConnection1.objectConnection_
   
   

   
   cboLayer.Clear
   For a = 0 To db.getLayerCount - 1
      Select Case UCase(db.getLayerName(a))
         Case "WATERLINES", "WATERCOMPONENTS", "SEWERLINES", "SEWERCOMPONENTS", "DRAINLINES", "DRAINCOMPONENTS" _
               , "RAMAIS", "DOCUMENTOS", "AMARRACAO", "IMAGEM"
         Case Else
            cboLayer.AddItem db.getLayerName(a)
      End Select
   Next
   Me.Show , FrmMain
End Function

Private Sub cboLayer_Click()
   Dim rs As ADODB.Recordset, layer_id As Integer, attrib_link As String, a As Integer
   Dim bb As String
   Dim cc As String
   If cboLayer.ListIndex >= 0 Then
      db.setCurrentLayer cboLayer.Text
      If db.existsRepresentation(1) Then
         optPolygons.Enabled = True
      Else
         optPolygons.Enabled = False
      End If
      If db.existsRepresentation(2) Then
         optLines.Enabled = True
      Else
         optLines.Enabled = False
      End If
      If db.existsRepresentation(4) Then
         optPoints.Enabled = True
      Else
         optPoints.Enabled = False
      End If
      List1.Clear
      List2.Clear
      If Not optLines.Enabled And Not optPoints.Enabled And Not optPolygons.Enabled Then
         optLines.value = False
         optPoints.value = False
         optPolygons.value = False
         MsgBox "Este plano não contém nenhuma geometria de pontos, linhas ou poligonos e não pde ser gerado texto para o mesmo", vbExclamation
         Exit Sub
      End If
      cgeo.GetLayerAttrib cboLayer.Text, layer_id, attrib_link
      If attrib_link = "" Then Exit Sub
      
      'alterado em 19/10/2010
       If frmCanvas.TipoConexao <> 4 Then
      
      Set rs = Conn.execute("SELECT * from " & cboLayer & " where " & attrib_link & "='0'")
      For a = 0 To rs.Fields.count - 1
         List1.AddItem rs.Fields(a).Name
      Next
   End If
   Else
   bb = "& cboLayer &"
   cc = "& attrib_link &"
   Set rs = Conn.execute("SELECT  from "" & cboLayer & "" where "" & attrib_link & ""='0'")
      For a = 0 To rs.Fields.count - 1
         List1.AddItem rs.Fields(a).Name
      Next
   End If
   
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   
End Sub

Private Sub CmdInserir_Click()
   If List1.Text <> "" Then
      List2.AddItem List1.Text
      List1.RemoveItem List1.ListIndex
   End If
End Sub

Private Sub cmdOK_Click()
   Dim a As Integer, Geometria As TeAllGeometries, _
   Position As TECOMDATABASELibCtl.TePositions, Fields As String, sep As String
   If Not optLines.value And Not optPoints.value And Not optPolygons.value Then
      MsgBox "Este plano não contém nenhuma geometria selecionada", vbExclamation
      Exit Sub
   End If
   
   If List2.ListCount = 0 Then
      MsgBox "Selecione ao menos um atributo", vbExclamation
      Exit Sub
   End If
   
   For a = 0 To List2.ListCount - 1
      If Fields = "" Then
         Fields = List2.list(a)
      Else
         Fields = Fields & "," & List2.list(a)
      End If
   Next

   If optPoints.value Then
      Geometria = taPOINTS
   End If
   
   If optLines.value Then
      Geometria = taLINES
   End If

   If optPolygons.value Then
      Geometria = taPOLYGONS
   End If


   If optHifem.value Then
      sep = "-"
   Else
      sep = " "
   End If

   If optCenter And optPolygons Then
      Position = TeCenterPolygon
   ElseIf optCenter And optLines Then
      Position = TeMiddleLine
   ElseIf optEnd Then
      Position = TeEnd
   ElseIf optInit Then
      Position = TeInit
   End If

   InsertText Geometria, Fields, sep, Position
End Sub


Sub InsertText(Geometria As TeAllGeometries, mFields As String, sep As String, Position As TECOMDATABASELibCtl.TePositions)
   Dim rs As ADODB.Recordset, layer_id As Integer, attrib_link As String, SQL As String, a As Integer
   
   cgeo.GetLayerAttrib cboLayer.Text, layer_id, attrib_link
   db.setCurrentLayer cboLayer.Text

   If db.existsRepresentation(128) = 1 Then
   

a = "Texts"
b = layer_id
c = "b"

     If frmCanvas.TipoConexao <> 4 Then
      Conn.execute "delete from texts" & layer_id
      Else
      Conn.execute "delete from " + """" + a + Trim(str(b)) + """"
      End If
   Else
      db.addGeometryRepresentation cboLayer, 128
      db.setCurrentLayer cboLayer.Text
   End If


   Select Case Geometria
      Case TeAllGeometries.taLINES
         SQL = "lines"
      Case TeAllGeometries.taPOINTS
         SQL = "points"
      Case TeAllGeometries.taPOLYGONS
         SQL = "polygons"
   End Select


 
            
         If frmCanvas.TipoConexao <> 4 Then
          Set rs = Conn.execute("SELECT count(*) from " & SQL & layer_id & _
            " inner join " & cboLayer.Text & " on object_id=" & attrib_link)
                  Else
                  Dim va2 As String
                  Dim ve2 As String
                  Dim vi2 As String
                  va2 = "geom_id"
                   ve2 = "object_id"
                   
                  Set rs = Conn.execute("SELECT count(*) from " & """" + SQL & Trim(str(layer_id)) & """" + " inner join " & """" + cboLayer.Text & """" + " on " + """" + ve2 + """" + "='" & attrib_link & "'")
                  End If
            

            
   If rs(0).value > 0 Then
      With frmProgressBar

         .ProgressBar1.value = 0
         .ProgressBar1.Max = rs.Fields(0).value
         .Show , FrmMain
         DoEvents
         If frmCanvas.TipoConexao <> 4 Then
         Set rs = Conn.execute("SELECT geom_id, object_id," & mFields & " from " & SQL & layer_id & _
                  " inner join " & cboLayer.Text & " on object_id=" & attrib_link)
                  Else
                  Dim va As String
                  Dim ve As String
                  Dim vi As String
                  va = "geom_id"
                   ve = "object_id"
                   
                  Set rs = Conn.execute("SELECT " + """" + va + """" + ", " + """" + ve + """" + "," + """" + mFields + """" + " from " + """" + SQL & layer_id + """" + " inner join " + """" + cboLayer.Text + """" + " on " + """" + ve + """" + "='" & attrib_link & "'")
                  End If
                  
                  
         While Not rs.EOF

            For a = 2 To rs.Fields.count - 1
               If a = 2 Then
                  SQL = IIf(IsNull(rs.Fields(a).value), "", rs.Fields(a).value)
               Else
                  SQL = SQL & sep & IIf(IsNull(rs.Fields(a).value), "", rs.Fields(a).value)
               End If
            Next
            db.insertTextFromGeometryReference , rs!object_id, rs!geom_id, Position, _
                           , , , Geometria, _
                           SQL, 2, 2, True
            .ProgressBar1.value = .ProgressBar1.value + 1
            .Caption = "Processado " & .ProgressBar1.value & " de " & .ProgressBar1.Max
            DoEvents
            rs.MoveNext
         Wend

      End With
      Unload frmProgressBar
      MsgBox "Processamento Concluído", vbInformation
   Else
      MsgBox "Nenhuma Geometria Encontrada", vbExclamation
      End If
   
   
   
   
 
   rs.Close
   Set rs = Nothing
End Sub

Private Sub cmdRemover_Click()
   If List2.Text <> "" Then
      List1.AddItem List2.Text
      List2.RemoveItem List2.ListIndex
   End If
End Sub


