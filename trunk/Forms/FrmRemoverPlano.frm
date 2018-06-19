VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TeComDatabase.dll"
Begin VB.Form FrmRemoverPlano 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Remover Plano de Informações"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   285
      Left            =   2010
      TabIndex        =   2
      Top             =   2790
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Confirmar"
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   2790
      Width           =   945
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2805
   End
   Begin TECOMDATABASELibCtl.TeDatabase db 
      Left            =   1830
      OleObjectBlob   =   "FrmRemoverPlano.frx":0000
      Top             =   3090
   End
End
Attribute VB_Name = "FrmRemoverPlano"
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


Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   On Error GoTo comeco_err
   Dim cgeo As New clsGeoReference, layer_id As Integer, rs As ADODB.Recordset
   
   If MsgBox("Deseja realmente excluir este plano", 36) = vbYes Then
      Usuario.UsrId = Sec.OpenLogin(Conn)
      If nxUser.GetPermission(Conn, Usuario.UsrId, "ADMINISTRAR", nxDelete) Then
         On Error GoTo Continuar
         layer_id = cgeo.GetLayerID(List1.Text)
       
         If db.deleteLayer(List1.Text) = 0 Then
Continuar:
            If layer_id > 0 Then
            
            
a = "te_visual"
b = "legend_id"
c = "te_grouping"
d = "theme_id"
e = "te_theme"
f = "layer_id"
g = "te_legend"
h = "te_theme_application"

     If frmCanvas.TipoConexao <> 4 Then
            Conn.execute "delete from te_visual where legend_id in(" & _
                        "SELECT legend_id from te_legend where theme_id in(" & _
                        "SELECT theme_id from te_theme where layer_id = " & layer_id & "))"
   
            Conn.execute "delete from te_grouping where theme_id in(" & _
                        "SELECT theme_id from te_theme where layer_id =" & layer_id & ")"
   
            Conn.execute "delete from te_legend where theme_id in(" & _
                        "SELECT theme_id from te_theme where layer_id =" & layer_id & ")"
   
            Conn.execute "delete from te_theme_application where theme_id in " & _
                        "(SELECT theme_id from te_theme where layer_id = " & layer_id & ")"
            
            Conn.execute "delete from te_theme where layer_id =" & layer_id
            Else
            
            Conn.execute "delete from " + """" + a + """" + " where " + """" + b + """" + " in(" & _
                        "SELECT " + """" + b + """" + " from " + """" + g + """" + " where " + """" + d + """" + " in(" & _
                        "SELECT " + """" + d + """" + " from " + """" + e + """" + " where " + """" + f + """" + " = '" & layer_id & "'))"
   
            Conn.execute "delete from " + """" + c + """" + " where " + """" + d + """" + " in(" & _
                        "SELECT " + """" + d + """" + " from " + """" + e + """" + " where " + """" + f + """" + " ='" & layer_id & "')"
   
            Conn.execute "delete from " + """" + g + """" + " where " + """" + d + """" + " in(" & _
                        "SELECT " + """" + d + """" + " from " + """" + e + """" + " where " + """" + f + """" + " ='" & layer_id & "')"
   
            Conn.execute "delete from " + """" + h + """" + " where " + """" + d + """" + " in (" & _
                        "(SELECT " + """" + d + """" + " from " + """" + e + """" + " where " + """" + e + """" + " = '" & layer_id & "')"
            
            Conn.execute "delete from " + e + " where " + f + " ='" & layer_id & "'"
            End If
            
 
a = "te_representation"

c = layer_id

            
e = "geom_table"
f = "geom_type"




            If frmCanvas.TipoConexao <> 4 Then
            Set rs = Conn.execute("SELECT geom_table, geom_type from te_representation where layer_id=" & layer_id)
            Else
            Set rs = Conn.execute("SELECT " + """" + e + """" + "," + """" + f + """" + " from " + """" + a + """" + " where " + """" + c + """" + " = '" & layer_id & "'")
            End If
            While Not rs.EOF
               If rs.Fields("geom_type").value = 128 Then
                  DropTable rs.Fields("geom_table").value & "_txvisual"
               End If
               DropTable rs.Fields("geom_table").value
               rs.MoveNext
            Wend
            rs.Close
            If frmCanvas.TipoConexao <> 4 Then
            Conn.execute "delete from te_representation where layer_id =" & layer_id
            Else
            Conn.execute "delete from " + """" + a + """" + " where " + """" + c + """" + " ='" & layer_id & "'"
            End If
a = "attr_table"
b = "te_layer_table"
c = "layer_id"
            If frmCanvas.TipoConexao <> 4 Then
            Set rs = Conn.execute("SELECT attr_table from te_layer_table where layer_id=" & layer_id)
            Else
            Set rs = Conn.execute("SELECT " + """" + a + """" + " from " + """" + b + """" + " where " + """" + c + """" + " = '" & layer_id & "'")
            End If
            While Not rs.EOF
               DropTable rs.Fields("attr_table").value
               rs.MoveNext
            Wend
            rs.Close
            
            
            
a = "te_layer_table"
b = "layer_id"

e = "te_layer"

            If frmCanvas.TipoConexao <> 4 Then
            Conn.execute "delete from te_layer_table where layer_id=" & layer_id
            
            Conn.execute "delete from te_layer where layer_id=" & layer_id
            Else
             Conn.execute "delete from " + """" + a + """" + " where " + """" + b + """" + "='" & layer_id & "'"
            
            Conn.execute "delete from " + """" + e + """" + " where " + """" + b + """" + "= '" & layer_id & "'"
            End If
            End If
         End If
         
a = "X_MANAGERPROPERTIESB"
b = "TABLENAME"
c = "X_LAYERSCOMPONENTS"
d = "LAYERLINE"
e = "LAYERCOMPONENT"

         If frmCanvas.TipoConexao <> 4 Then
         Conn.execute "DELETE FROM X_ManagerPropertiesB WHERE TABLENAME='" & List1.Text & "'"
         Conn.execute "DELETE FROM X_LayersComponents WHERE LAYERLINE='" & List1.Text & "' AND LAYERCOMPONENT='" & List1.Text & "'"
         Else
         Conn.execute "DELETE FROM " + """" + a + """" + " WHERE " + """" + b + """" + "='" & List1.Text & "'"
         Conn.execute "DELETE FROM " + """" + c + """" + " WHERE " + """" + d + """" + "='" & List1.Text & "' AND " + """" + e + """" + "='" & List1.Text & "'"
         End If
  
         List1.RemoveItem List1.ListIndex
         MsgBox "Plano removido com sucesso", vbExclamation
      End If
   End If
   Exit Sub
comeco_err:
   MsgBox Err.Description
End Sub

Private Sub DropTable(TableName As String)
   On Error GoTo DropTable_sair
   Conn.execute "drop table " & TableName
DropTable_sair:
End Sub


Private Sub Form_Load()
   Dim a As Integer
   db.Provider = typeconnection
   db.Connection = Conn
   Dim rs As ADODB.Recordset
a = "name"
b = "te_layer"


     If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute("SELECT NAME from TE_LAYER ORDER BY NAME")
   Else
  Set rs = Conn.execute("SELECT " + """" + a + """" + " from " + """" + b + """" + "  ORDER BY " + """" + a + """" + "")
   End If
   While Not rs.EOF
   
      Select Case UCase(rs.Fields("NAME").value)
         Case "WATERLINES", "WATERCOMPONENTS", "SEWERLINES", "SEWERCOMPONENTS", "DRAINLINES", "DRAINCOMPONENTS" _
               , "RAMAIS", "DOCUMENTOS", "AMARRACAO", "IMAGEM"
         Case Else
            List1.AddItem rs.Fields("NAME").value
      End Select
      rs.MoveNext
   Wend
   rs.Close
   Set rs = Nothing
End Sub


