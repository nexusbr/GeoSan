VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TECOMD~1.DLL"
Object = "{2CCABA93-B681-4E7F-8047-BD4D623301BA}#1.0#0"; "TECOMI~1.DLL"
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TECOMC~1.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar "
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   2175
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   4305
      Begin TECOMIMPORTLibCtl.TeImport TeImport1 
         Left            =   3120
         OleObjectBlob   =   "frmImportFile.frx":0000
         Top             =   240
      End
      Begin MSComDlg.CommonDialog cdl 
         Left            =   1770
         Top             =   1590
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Importar"
         Height          =   375
         Left            =   420
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2730
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox cboLayer 
         Height          =   315
         ItemData        =   "frmImportFile.frx":0024
         Left            =   360
         List            =   "frmImportFile.frx":0026
         TabIndex        =   5
         Top             =   450
         Width           =   3525
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1020
         Width           =   3165
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "...."
         Height          =   285
         Left            =   3570
         TabIndex        =   1
         Top             =   1020
         Width           =   285
      End
      Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
         Left            =   3840
         OleObjectBlob   =   "frmImportFile.frx":0028
         Top             =   840
      End
      Begin TECOMDATABASELibCtl.TeDatabase TeDatabase1 
         Left            =   3840
         OleObjectBlob   =   "frmImportFile.frx":004C
         Top             =   360
      End
      Begin VB.Label lblLayer 
         AutoSize        =   -1  'True
         Caption         =   "Plano de Informação"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   210
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Arquivo"
         Height          =   225
         Left            =   360
         TabIndex        =   3
         Top             =   810
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmImportFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Mylayer As String
Private Conn  As ADODB.connection
Dim tdb As TeDatabase
Dim TeImp As TeImport
Dim aa As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
     Dim Abertura As Integer
'Public Function init(mConn As ADODB.Connection, tImp As TeImport, mtdb As TeDatabase) As String
  ' 'LoozeXP1.InitIDESubClassing
  ' Set Conn = mConn
  ' Set TeImp = tImp
  ' Set tdb = mtdb
   
  ' Me.Show vbModal
   
  ' init = Mylayer
  ' 'LoozeXP1.EndWinXPCSubClassing
'End Function

Public Function init(mConn As ADODB.connection, ti As TeImport, td As TeDatabase) As String
   'LoozeXP1.InitIDESubClassing
   Set Conn = mConn
   Set tdb = td
   Set TeImp = ti
      
   Me.Show vbModal
   
   init = Mylayer
   'LoozeXP1.EndWinXPCSubClassing
End Function




Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdImport_Click()
   'On Error GoTo Trata_Erro
   Dim result As Integer, a As Integer, IniciouTransacao As Boolean
   Dim rs As New ADODB.Recordset
   Dim SQL As String
   Dim dy As String
   Dim dh As String
   dy = "te_layer"
   dh = "name"
   IniciouTransacao = True
   Dim a1, a2, a3, a4, a5, a6, a7, a8, a9, a10 As String



   If cboLayer.Text = "" Then
      MsgBox "O campo nome do plano esta vazio !", vbInformation
      Exit Sub
   End If

   If txtFileName.Text = "" Then
      MsgBox "O campo nome do arquivo está vazio!", vbInformation
      Exit Sub
   End If
   If frmCanvas.TipoConexao <> 4 Then

   SQL = "SELECT * FROM TE_LAYER WHERE Name = '" & UCase(cboLayer.Text) & "'"
   Else
    SQL = "SELECT  * FROM " + """" + dy + """" + " WHERE " + """" + dh + """" + "  = '" & cboLayer.Text & "'"
    End If
   rs.Open SQL, Conn, adOpenDynamic, adLockOptimistic

   If Not rs.EOF Then
      rs.Close
      MsgBox "Plano já existe", vbExclamation
      Exit Sub
   Else
   rs.Close
      
      
   If frmProjection.init() Then

      'CRIA O NOVO LAYER  ''ATENÇÃO! TUDO EM CASE SENSITIVE!!
      
      'cboLayer.Text = "teste_teste"
      
   ' Set TeImp = TeImport1
'Set tdb = TeDatabase1

  
a1 = UCase(cboLayer.Text)
a2 = UCase(frmProjection.cboProjecao)
a3 = UCase(frmProjection.cboDatum)
a4 = UCase(frmProjection.txtLongOrigem)
a5 = UCase(frmProjection.txtLatOrigem)
a6 = UCase(frmProjection.txtOffSetX)
a7 = UCase(frmProjection.txtOffSetY)
a8 = UCase(frmProjection.txtParPadrao1)
a9 = UCase(frmProjection.txtParPadrao2)


   
      
If TeImp.createLayer(a1, a2, a3, a4, a5, a6, a7, a8, a9, 2) = True Then
         
End If





      
     ' If TeImp.createLayer(a1, a2, a3, a4, a5, a6, a7, a8, a9, 2) = True Then
         

     ' End If
      'If TeImp.createLayer("hhhyuzzt55hh", "UTM", "SAD69", "-45", "0", "50000", "1000000", "0", "0", 2) = True Then
        

     ' End If
      
      'ADICIONA AS GEOMETRIAS
      tdb.addGeometryRepresentation UCase(cboLayer.Text), 1 'poligono
      tdb.addGeometryRepresentation UCase(cboLayer.Text), 4 'pontos
      tdb.addGeometryRepresentation UCase(cboLayer.Text), 128 'textos
   
   Else

      Exit Sub
   End If


      Screen.MousePointer = vbHourglass
      Select Case UCase(Right(txtFileName.Text, 3))

        Case "GEO":

           result = TeImp.importGeoTab(txtFileName.Text, cboLayer.Text, "")

        Case "SHP":

           result = TeImp.importSHP(txtFileName.Text, cboLayer.Text, "")

        Case "MIF":

          result = TeImp.importMIF(txtFileName.Text, cboLayer.Text, "", "")

      End Select


      Screen.MousePointer = vbNormal
      
      If result = 1 Then

        Mylayer = cboLayer.Text
        
    
              If frmCanvas.TipoConexao <> 4 Then
         
       Conn.execute "INSERT INTO X_LayersComponents (LAYERLINE,LAYERCOMPONENT,TYPEREFERENCE,TYPETEXT) VALUES('" & Mylayer & "','" & Mylayer & "',4,1)"
     Else
     aa = "LAYERLINE"
      b = "LAYERCOMPONENT"
      c = "TYPEREFERENCE"
      d = "TYPETEXT"
      h = "X_LAYERSCOMPONENTS"

     
 Conn.execute "INSERT INTO " + """" + h + """" + " ( " + """" + aa + """" + " ," + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + " ) VALUES('" & Mylayer & "','" & Mylayer & "','4','1')"
     End If
     
        
     
        
        
        
        
        Set rs = Conn.execute("SELECT * from " & cboLayer.Text)
        For a = 0 To rs.Fields.count - 1
        
        
            If frmCanvas.TipoConexao <> 4 Then
         
     Conn.execute "Insert into X_ManagerPropertiesB (TableName,FieldSequence) values('" & cboLayer.Text & "'," & a & ")"
     Else
     
      aa = "TABLENAME"
      b = "FIELDSEQUENCE"
      h = "X_MANAGERPROPERTIESB"
     
 Conn.execute "INSERT INTO " + """" + h + """" + " ( " + """" + aa + """" + " ," + """" + b + """" + ") values('" & cboLayer.Text & "','" & a & "')"
     End If
           
            
            
        Next
                'CORREÇÃO DO BUG DO TE_IMPORT
        rs.Close
        Dim dw As String ' alterado em 20/10/2010
        Dim dr As String
        Dim dp As String
         dw = "table_id"
        dr = "te_layer_table"
        dp = "attr_table"
        
        If frmCanvas.TipoConexao <> 4 Then

        Set rs = Conn.execute("SELECT Max(table_id) , count(*) from te_layer_table where attr_table='" & cboLayer.Text & "'")
        Else
        Set rs = Conn.execute("SELECT Max(" + """" + dw + """" + ") , count(*) from " + """" + dr + """" + " where " + """" + dp + """" + "='" & cboLayer.Text & "'")
        End If
        If Not rs.EOF Then
            If rs.Fields(1).value > 1 Then
            a = "te_layer_table"
            b = "table_id"
          
           
            If frmCanvas.TipoConexao <> 4 Then

            Conn.execute ("delete from TE_LAYER_TABLE where TABLE_ID = rs.Fields(0).value")
            Else
               Conn.execute ("delete from " + """" + aa + """" + " where " + """" + b + """" + " ='" & rs.Fields(0).value & "'")
            End If
        End If
        rs.Close
        
        
        
        Set rs = Nothing
        MsgBox "Importação Concluida!", vbInformation
         Unload Me
         While Not FrmMain.ActiveForm Is Nothing
             
            Unload FrmMain.ActiveForm
            FrmMain.SetFocus

         Wend
         FrmMain.tbToolBar_ButtonClick FrmMain.tbToolBar.Buttons("knew")
      Else

        MsgBox "Ocorreu um erro na importação!", vbInformation
      End If

   End If
End If



'Trata_Erro:
  ' If Err.Number = 0 Or Err.Number = 20 Then
   '   Resume Next
  ' Else
   '   PrintErro "frmImportFile", "cmdImport_Click()", CStr(Err.Number), CStr(Err.Description), True
   
      'Resume
   
  ' End If
End Sub

Private Sub cmdOpenFile_Click()
   CDL.Filter = "Arquivo Shape (*.shp)|*.shp|Arquivo SPRING (*.geo)|*.geo|Arquivo MapInfo (*.mif)|*.mif|"
   CDL.FilterIndex = 1
   CDL.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
   CDL.ShowOpen
   txtFileName.Text = CDL.FileName
End Sub

