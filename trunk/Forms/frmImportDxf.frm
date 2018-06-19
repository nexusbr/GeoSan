VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportDxf 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar Dxf"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   3705
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   4245
      Begin VB.CommandButton cmdImport 
         Caption         =   "Importar"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   3120
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2550
         TabIndex        =   13
         Top             =   3120
         Width           =   1305
      End
      Begin VB.ComboBox cboLayer 
         Height          =   315
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   3525
      End
      Begin VB.Frame Frame 
         Caption         =   "Geometrias"
         Height          =   1005
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   3525
         Begin MSComDlg.CommonDialog cdl 
            Left            =   2940
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CheckBox chk 
            Caption         =   "Ponto"
            Height          =   255
            Index           =   2
            Left            =   2550
            TabIndex        =   12
            Top             =   300
            Width           =   735
         End
         Begin VB.CheckBox chk 
            Caption         =   "Linha"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   11
            Top             =   300
            Width           =   735
         End
         Begin VB.CheckBox chk 
            Caption         =   "Poligono"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   300
            Width           =   945
         End
         Begin VB.CheckBox chk 
            Caption         =   "Texto"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   9
            Top             =   630
            Width           =   795
         End
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1050
         Width           =   3165
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "...."
         Height          =   285
         Left            =   3600
         TabIndex        =   1
         Top             =   1050
         Width           =   285
      End
      Begin VB.ComboBox cboDXFlayer 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1620
         Width           =   3525
      End
      Begin VB.Label lblLayer 
         AutoSize        =   -1  'True
         Caption         =   "Plano de Informação"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   210
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Arquivo"
         Height          =   225
         Left            =   360
         TabIndex        =   6
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Planos do arquivo"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1350
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmImportDxf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TeImp As TeImport

Private tdb As TeDatabase

Private Conn As ADODB.connection

Public Function init(mConn As ADODB.connection, tImp As TeImport, mtdb As TeDatabase) As Boolean
   'LoozeXP1.InitIDESubClassing
   
   Dim contplano As Integer
   Dim Cont As Integer
   
   Set Conn = mConn
   Set TeImp = tImp
   Set tdb = mtdb
   chk(0).Enabled = False
   chk(1).Enabled = False
   chk(2).Enabled = False
   chk(3).Enabled = False
   
   contplano = mtdb.getLayerCount()
   For Cont = 0 To contplano - 1
      cboLayer.AddItem mtdb.getLayerName(Cont)
   Next
   
   Me.Show vbModal
   Unload Me
   'LoozeXP1.InitIDESubClassing
End Function

Private Sub cboDXFlayer_Click()
   On Error GoTo cboDXFlayer_Click_err
   chk(0).Enabled = False
   chk(1).Enabled = False
   chk(2).Enabled = False
   chk(3).Enabled = False
   If cboDXFlayer.Text <> "" Then
      If (TeImp.getDxfGeometryLayer(txtFileName.Text, cboDXFlayer.Text, 1)) = 1 Then
         chk(0).Enabled = True
      End If
      If (TeImp.getDxfGeometryLayer(txtFileName.Text, cboDXFlayer.Text, 2)) = 1 Then
         chk(1).Enabled = True
      End If
      If (TeImp.getDxfGeometryLayer(txtFileName.Text, cboDXFlayer.Text, 4)) = 1 Then
         chk(2).Enabled = True
      End If
      If (TeImp.getDxfGeometryLayer(txtFileName.Text, cboDXFlayer.Text, 128)) = 1 Then
         chk(3).Enabled = True
      End If
   End If
   Exit Sub
cboDXFlayer_Click_err:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub chk_Click(index As Integer)
   Dim a As Integer
   For a = 0 To 3
      If a <> index Then
         chk(a).value = 0
      End If
   Next
End Sub

Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdImport_Click()
   
   'On Error GoTo AoReCarregar
      Dim af As String
      Dim ag As String
   af = "te_layer"
    ag = "name"
   
   On Error GoTo cmdimport_err
   
   Dim SQL As String, rs As ADODB.Recordset, georepresentacao As Integer
   
   If cboDXFlayer.Text = "" Then
     MsgBox "Plano do Arquivo Esta Vazio!", vbInformation
     Exit Sub
   End If

   If txtFileName.Text = "" Then
     MsgBox "O Nome do Arquivo Esta Vazio !", vbExclamation
     Exit Sub
   End If
   If cboLayer.Text = "" Then
     MsgBox "O Nome do Plano Esta Vazio !", vbExclamation
     Exit Sub
   End If
   Conn.BeginTrans
   If frmCanvas.TipoConexao <> 4 Then

   SQL = "SELECT * FROM TE_LAYER WHERE Name = '" & UCase(cboLayer.Text) & "'"
   Else
   SQL = "SELECT  * FROM " + """" + af + """" + " WHERE " + """" + ag + """" + " = '" & UCase(cboLayer.Text) & "'"
   End If
   Set rs = New ADODB.Recordset
   rs.Open SQL, Conn, adOpenDynamic, adLockOptimistic

   If rs.EOF Then
      cboLayer.Text = UCase(cboLayer.Text)

      If frmProjection.init() Then
         With frmProjection
            TeImp.createLayer cboLayer.Text, .cboProjecao, .cboDatum, .txtLongOrigem, _
               .txtLatOrigem, .txtOffSetX, .txtOffSetY, .txtParPadrao1, .txtParPadrao2, 1
            tdb.addGeometryRepresentation cboLayer.Text, 128
            tdb.addGeometryRepresentation cboLayer.Text, 2
            tdb.addGeometryRepresentation cboLayer.Text, 4
            Unload frmProjection
         End With
      Else
         rs.Close
         Set rs = Nothing
         Conn.RollbackTrans
         Exit Sub
      End If

   End If

   rs.Close
   Set rs = Nothing

   If chk(3).value = 1 Then
      georepresentacao = 128
   End If

   If chk(2).value = 1 Then
      georepresentacao = 4
   End If

   If chk(1).value = 1 Then
      georepresentacao = 2
   End If

   If chk(0).value = 1 Then
      georepresentacao = 1
   End If

   If georepresentacao = 0 Then
      Conn.RollbackTrans
      MsgBox "É necessário selecionar uma geometria!", vbInformation
      
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   
   
   
   If TeImp.importDXF(cboLayer.Text, txtFileName.Text, cboDXFlayer.Text, georepresentacao) = 1 Then
     'frmAvi.Visible = False
     Screen.MousePointer = vbDefault

     Conn.CommitTrans
     
     MsgBox "Importação Concluida !", vbInformation
   Else
     'frmAvi.Visible = False
     Screen.MousePointer = vbDefault
     MsgBox "Ocorreu um erro na importação !", vbInformation
   End If
   
   Unload Me
   While Not FrmMain.ActiveForm Is Nothing
       
      Unload FrmMain.ActiveForm
      FrmMain.SetFocus
      'FrmMain.ActiveForm.SetFocus
   Wend
   FrmMain.tbToolBar_ButtonClick FrmMain.tbToolBar.Buttons("knew")
   Exit Sub
cmdimport_err:
   Conn.RollbackTrans
End Sub

Private Sub cmdOpenFile_Click()
   On Error GoTo cmdOpenFile_Click_err
   Dim contplano As Integer
   Dim Cont As Integer
   CDL.Filter = "Arquivo AutoCad (*.dxf)|*.dxf|"
   CDL.FilterIndex = 1
   CDL.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
   CDL.ShowOpen
   txtFileName.Text = CDL.FileName
   cboDXFlayer.Clear
   If txtFileName.Text <> "" Then
      contplano = TeImp.getDxfLayersCount(txtFileName.Text)
      For Cont = 0 To contplano - 1
         cboDXFlayer.AddItem TeImp.getDxfLayersFromIndex(txtFileName.Text, Cont)
      Next Cont
   Else
      cboDXFlayer.Clear: cboDXFlayer.ListIndex = -1
   End If
   Exit Sub
cmdOpenFile_Click_err:
   MsgBox Err.Description, vbExclamation
End Sub


