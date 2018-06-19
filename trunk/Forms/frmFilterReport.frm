VERSION 5.00
Begin VB.Form frmFilterReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filtro"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Data de Instalação"
      Height          =   885
      Left            =   120
      TabIndex        =   4
      Top             =   570
      Width           =   3495
      Begin VB.TextBox mskDataFim 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox mskDataInicio 
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Até"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "De"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   390
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cboManufacture 
      Height          =   315
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2235
   End
   Begin VB.Label Label3 
      Caption         =   "Marca"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmFilterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private str As String
Dim a1 As String
Dim b1 As String
Dim a2 As String
Dim a3 As String
Dim a4 As String
Dim a5 As String

Function init() As String
   'LoozeXP1.InitIDESubClassing
   Dim rs As ADODB.Recordset
   Dim dd As String
   Dim df As String
   dd = "X_MANUFACTURERS" ' alterado em 20/10/2010
   df = "COMPANYNAME"
   
   If frmCanvas.TipoConexao <> 4 Then

   Set rs = Conn.execute("SELECT * from x_manufacturers order by companyName")
   Else
   ' WritePrivateProfileString "A", "A", "SELECT  from " + """" + dd + """" + " order by " + """" + df + """" + "", App.path & "\DEBUG.INI"
   Set rs = Conn.execute("SELECT * from " + """" + dd + """" + " order by " + """" + df + """" + "")
   End If
   While Not rs.EOF
      cboManufacture.AddItem rs!CompanyName
      cboManufacture.ItemData(cboManufacture.NewIndex) = rs!ManufacturerID
      rs.MoveNext
   Wend
   Set rs = Nothing
   Me.Show vbModal
   init = str
'   'LoozeXP1.InitIDESubClassing
End Function

Private Sub cmdCancel_Click()
   blnGeraRel = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
a1 = "DATEINSTALLATION"
b1 = "MANUFACTURER"

  If frmCanvas.TipoConexao <> 4 Then
   If IsDate(mskDataInicio.Text) And IsDate(mskDataFim.Text) Then
      str = " where DateInstallation >='" & Format(mskDataInicio.Text, "yyyymmdd") & "'" & _
            "AND DateInstallation <='" & Format(mskDataFim.Text, "yyyymmdd") & "'"
   End If
   If cboManufacture.ListIndex > -1 Then
      If str <> "" Then
         str = str & " AND Manufacturer=" & cboManufacture.ItemData(cboManufacture.ListIndex)
      Else
         str = " where Manufacturer=" & cboManufacture.ItemData(cboManufacture.ListIndex)
      End If
   End If
   
   Else
    If IsDate(mskDataInicio.Text) And IsDate(mskDataFim.Text) Then
      str = " where " + """" + a1 + """" + " >='" & Format(mskDataInicio.Text, "yyyymmdd") & "'" & _
            "AND " + """" + a1 + """" + " <='" & Format(mskDataFim.Text, "yyyymmdd") & "'"
   End If
   If cboManufacture.ListIndex > -1 Then
      If str <> "" Then
         str = str & " AND " + """" + b1 + """" + "='" & cboManufacture.ItemData(cboManufacture.ListIndex) & "'"
      Else
         str = " where " + """" + b1 + """" + "='" & cboManufacture.ItemData(cboManufacture.ListIndex) & "'"
      End If
   End If
   
   End If
   
   
   
   blnGeraRel = True
   Unload Me
End Sub




