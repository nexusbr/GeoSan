VERSION 5.00
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmEPANET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Exporta��o EPANET"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   ControlBox      =   0   'False
   Icon            =   "FrmEPANET.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Caminho de Exporta��o"
      Height          =   990
      Left            =   120
      TabIndex        =   4
      Top             =   210
      Width           =   6165
      Begin VB.TextBox txtArquivo 
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   375
         Width           =   5325
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   330
         Left            =   5550
         TabIndex        =   5
         Top             =   375
         Width           =   435
      End
   End
   Begin VB.TextBox txtTimer 
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Text            =   "20:00:00"
      Top             =   1335
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3450
      Top             =   1305
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3420
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   5190
      TabIndex        =   1
      Top             =   1335
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4035
      TabIndex        =   0
      Top             =   1335
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   165
      TabIndex        =   7
      Top             =   1335
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
      Left            =   4680
      OleObjectBlob   =   "FrmEPANET.frx":1CFA
      Top             =   120
   End
   Begin VB.Label Label4 
      Caption         =   "Hor�rio"
      Height          =   225
      Left            =   645
      TabIndex        =   3
      Top             =   1395
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "FrmEPANET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'EpanetExport Vers�o 06.10.09

Option Explicit
Public conn_recebida As ADODB.Connection
Public Provider As Integer
Public PLANO As String

Private rsTP As ADODB.Recordset
Private rsST As ADODB.Recordset

Dim i As Integer

'Declara��es necess�rias para a fun��o GetMyDocumentsDirectory()
Const REG_SZ = 1
Const REG_BINARY = 3
Const HKEY_CURRENT_USER = &H80000001
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_READ = &H20000
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_QUERY_VALUE = &H1
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
    ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
    ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Fim das declara��es necess�rias para a fun��o GetMyDocumentsDirectory()
'
'
'Rotina inicial da aplica��o
'
'
'
Public Sub init()
   cmdConfirmar.Default = True
   txtArquivo.Text = GetMyDocumentsDirectory() & "\GeoSan_Exp_Epanet_" & Format(Now, "YYYY-MM-DD-HHMMSS") & ".INP"
   Me.Show
End Sub
Private Sub cmdCancelar_Click()
   Cancelar = True
   Unload Me
End Sub
'Subrotina que inicia o timer e inicia a exporta��o para o Epanet
'
'
Private Sub Timer1_Timer()
    MousePointer = vbHourglass              'ativa a ampulheta
    'iniciaExportacaoParaEpanet              'inicia a convers�o para o EPANET
    Dim trechos As New TrechosRedeEpanet
    Dim totalTrechosExportar As Integer
    totalTrechosExportar = banco.ObtemNumeroTrechosQueSeraoExportados
    Me.ProgressBar1.Value = 1
    If totalTrechosExportar > 0 Then
        'existe pelo menos um trecho a ser exportado para o Epanet
        Me.ProgressBar1.Max = totalTrechosExportar
    Else
        'n�o existem trechos a serem exportados para o Epanet
        MsgBox "N�o h� dados selecionados para exportar.", vbInformation, ""
        End
    End If
    'In�cio da fun��o de exporta��o para o EPANET. Ao final dela ser� chamado o ModExport pela rotina ExportaEPANet que gera em mem�ria toda a exporta��o
    'para depois gerar em arquivo atrav�s de outra rotina. Esta rotina incicia quando o timer � iniciado
    'Revisar este coment�rio
    trechos.Exporta
    MousePointer = vbDefault                'desativa a ampulheta
    Timer1.Enabled = False                  'desativa o timer
    End
End Sub
'Subrotina que ir� iniciar a exporta��o para o Epanet - usu�rio selecionou o bot�o de exportar
'
'
Private Sub cmdConfirmar_Click()
    Timer1.Enabled = True               'ativa o timer
    Me.ProgressBar1.Visible = True      'ativa a visualiza��o da barra de progresso
    Me.cmdConfirmar.Enabled = False
End Sub

'Seleciona o nome do arquivo a exportar para o EPANET
'
'
'
Private Sub cmdPath_Click()
   cdl.Filter = "Epanet (.inp)|*.INP|Todos tipos (*.*)|*.*|"
   cdl.FileName = txtArquivo.Text
   cdl.InitDir = Environ$("USERPROFILE") & "\my documents"
   cdl.ShowSave
   txtArquivo.Text = cdl.FileName
End Sub

'Obtem o nome do diret�rio dos Meus Documentos do usu�rio que est� logado
'
'GetMyDocumentsDirectory() - retorna o caminho do diret�rio
'
Function GetMyDocumentsDirectory() As String
    Dim lRes As Long
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    Dim strData As Integer
    RegOpenKeyEx HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", 0, KEY_READ, lRes
    lResult = RegQueryValueEx(lRes, "Personal", 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(lRes, "Personal", 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                GetMyDocumentsDirectory = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        End If
    End If
    RegCloseKey lRes
End Function
