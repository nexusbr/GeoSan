VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSuppliersSub 
   Caption         =   "Selecione "
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "FrmSuppliersSub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtCompanyName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtManufacturersID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   4455
         Begin MSComctlLib.ListView Lv 
            Height          =   2055
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Companhia"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmSuppliersSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Man_ID As Long


Public Function init(TableName As String) As Long
   Dim rs As ADODB.Recordset, i As ListItem
   If frmCanvas.TipoConexao <> 4 Then
   
   Set rs = Conn.execute("SELECT * from " & TableName & " order by CompanyName")
   Else
   
   Set rs = Conn.execute("SELECT * from " + """" & UCase(TableName) & """" + " order by " + """" + "COMPANYNAME" + """" + "")
   End If
   
   
   Lv.ColumnHeaders(1).Width = 1000
   Lv.ColumnHeaders(2).Width = Lv.Width - 1100
   
   While Not rs.EOF
      Set i = Lv.ListItems.Add(, "a" & CStr(rs.Fields(0).value), rs.Fields(0).value)
         i.SubItems(1) = rs.Fields(1).value
         i.Tag = rs.Fields(0).value
      rs.MoveNext
   Wend
   
   rs.Close
   Set rs = Nothing
   Me.Show vbModal
   init = Man_ID
End Function


Private Sub Lv_DblClick()
   If Not (Lv.SelectedItem Is Nothing) Then
      Man_ID = Lv.SelectedItem.Tag
   End If
   Unload Me
End Sub


Private Sub txtCompanyName_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim a As Integer
   If txtCompanyName = "" Then Exit Sub
   For a = 1 To Lv.ListItems.Count
      If UCase(Left(Lv.ListItems(a).SubItems(1), Len(txtCompanyName.Text))) = _
         UCase(txtCompanyName.Text) Then
         Lv.ListItems.Item(Lv.ListItems(a).key).Selected = True
         If KeyCode = vbKeyReturn Then Lv_DblClick
         Exit For
      End If
   Next
End Sub


Private Sub txtManufacturerID_KeyUp(KeyCode As Integer, Shift As Integer)
'   Dim A As Integer
'   'If txtManufacturerId = "" Then Exit Sub
'   For A = 1 To Lv.ListItems.Count
'      If UCase(Left(Lv.ListItems(A).Text, Len(txtManufacturerId.Text))) = _
'         UCase(txtManufacturerId.Text) Then
'         Lv.ListItems.Item(Lv.ListItems(A).Key).SELECTed = True
'         If KeyCode = vbKeyReturn Then Lv_DblClick
'         Exit For
'      End If
'   Next
   
End Sub




