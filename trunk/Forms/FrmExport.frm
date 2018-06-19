VERSION 5.00
Object = "{C51C74EC-6107-4A01-8400-40B53BB20D42}#1.0#0"; "TeComExport.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Exportação DXF"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   7890
      Top             =   6495
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExport.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TECOMEXPORTLibCtl.TeExport TeExport1 
      Left            =   2505
      OleObjectBlob   =   "FrmExport.frx":079D
      Top             =   3270
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Fechar"
      Height          =   315
      Left            =   4125
      TabIndex        =   4
      Top             =   9255
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   315
      Left            =   5220
      TabIndex        =   3
      Top             =   9255
      Width           =   975
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "..."
      Height          =   315
      Left            =   5790
      TabIndex        =   1
      Top             =   8775
      Width           =   405
   End
   Begin VB.TextBox txtNomeArquivo 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   8775
      Width           =   5550
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5445
      Left            =   150
      TabIndex        =   6
      Top             =   225
      Width           =   6030
      Begin VB.OptionButton optSelecaoTemas 
         Caption         =   "Seleção de Temas"
         Height          =   315
         Left            =   300
         TabIndex        =   10
         Top             =   1140
         Width           =   1740
      End
      Begin VB.OptionButton optPlanoAtual 
         Caption         =   "Plano Atual Selecionado"
         Height          =   360
         Left            =   300
         TabIndex        =   9
         Top             =   720
         Width           =   3240
      End
      Begin VB.OptionButton optSelecionados 
         Caption         =   "Itens Selecionados no Mapa"
         Height          =   420
         Left            =   300
         TabIndex        =   8
         Top             =   315
         Width           =   2445
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         Left            =   195
         TabIndex        =   7
         Top             =   1170
         Width           =   5610
         Begin MSComctlLib.ListView Lv 
            Height          =   3495
            Left            =   105
            TabIndex        =   2
            Top             =   345
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   6165
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Opção de Arquivos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   135
      TabIndex        =   11
      Top             =   5850
      Width           =   6045
      Begin VB.OptionButton optArquivoCompleto 
         Caption         =   "1 Completo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   630
         TabIndex        =   23
         Top             =   465
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   540
         TabIndex        =   18
         Top             =   480
         Width           =   2175
         Begin VB.CheckBox Check4 
            Caption         =   "Linhas"
            Height          =   255
            Left            =   375
            TabIndex        =   22
            Top             =   360
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Polígonos"
            Height          =   255
            Left            =   375
            TabIndex        =   21
            Top             =   1350
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Textos"
            Height          =   255
            Left            =   375
            TabIndex        =   20
            Top             =   1020
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Pontos"
            Height          =   255
            Left            =   375
            TabIndex        =   19
            Top             =   690
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.OptionButton optArquivosSeparados 
         Caption         =   "Separados por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   480
         Width           =   1650
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   3165
         TabIndex        =   12
         Top             =   480
         Width           =   2175
         Begin VB.CheckBox chkPontos 
            Caption         =   "Pontos"
            Height          =   255
            Left            =   375
            TabIndex        =   16
            Top             =   690
            Width           =   855
         End
         Begin VB.CheckBox chkTextos 
            Caption         =   "Textos"
            Height          =   255
            Left            =   375
            TabIndex        =   15
            Top             =   1020
            Width           =   855
         End
         Begin VB.CheckBox chkPoligonos 
            Caption         =   "Polígonos"
            Height          =   255
            Left            =   375
            TabIndex        =   14
            Top             =   1335
            Width           =   1095
         End
         Begin VB.CheckBox chkLinhas 
            Caption         =   "Linhas"
            Height          =   255
            Left            =   375
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Diretório \ Pasta \ Arquivo"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   165
      TabIndex        =   5
      Top             =   8535
      Width           =   3150
   End
End
Attribute VB_Name = "FrmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TCanvas As TeCanvas
Private Conn As ADODB.connection
Private rs As Recordset
Private strLayerAtual As String
Private Representacao As Integer
Dim blnExportaTemasSelecionados As Boolean
Dim strDir As String
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
Dim xx As String ' alerado em 20/10/2010
   Dim xu As String
   Dim xt As String
   Dim xq As String
   Dim xp As String
   Dim xw As String
   Dim xf As String
   Dim xj As String
   Dim xn As String
   Dim xm As String
   Dim xv As String
   Dim RsB As Recordset
       Dim conexao As New ADODB.connection
     ' Dim conexao2 As New ADODB.connection
  

Private Sub Form_Load()
   'LoozeXP1.InitSubClassing
End Sub

Public Function init(mConn As ADODB.connection, mTeCanvas As TeCanvas, OnnerForm As Object)
On Error GoTo Trata_Erro

   Dim itmx As ListItem, strsql As String
   Set Conn = mConn
   
   xx = "te_layer"
   xu = "te_representation"
   xt = "name"
   xq = "theme_id"
   xp = "te_theme"
   xw = "view_id"
   xf = "layer_id"
   xj = "geom_type"
   xn = "user_name"
   xm = "te_view"
   xv = "generate_attribute_where"
   
   
   TeExport1.connection = mConn
   Set TCanvas = mTeCanvas
  ' Set Conn = mConn
   If frmCanvas.TipoConexao <> 4 Then
  
    strsql = "SELECT " & vbCr
    strsql = strsql & "l.name as " + """" + "layername" + """" + ", " & vbCr
    strsql = strsql & "user_name as " + """" + "usrnom" + """" + ", " & vbCr
    strsql = strsql & "t.name as " + """" + "thenom" + """" + ", " & vbCr
    strsql = strsql & "generate_attribute_where, " & vbCr
    strsql = strsql & "theme_id " & vbCr
    strsql = strsql & "From " & vbCr
    strsql = strsql & "((te_view v Inner Join te_theme t on t.view_id=v.view_id) " & vbCr
    strsql = strsql & "inner join te_representation r on t.layer_id=r.layer_id) " & vbCr
    strsql = strsql & "inner join te_layer l on t.layer_id=l.layer_id " & vbCr
    strsql = strsql & "Where " & vbCr
    strsql = strsql & "geom_type <= 128 " & vbCr
    strsql = strsql & "and v.user_name = '" & usuario.UseName & "' " & vbCr
    strsql = strsql & "and v.name = '" & FrmMain.ViewManager1.tvw.getActiveView & "' " & vbCr
    strsql = strsql & "Order By " & vbCr
    strsql = strsql & "t.Name " & vbCr
   Else
   strsql = "SELECT " & vbCr
    strsql = strsql & "" + """" + xx + """" + "." + """" + xt + """" + "as " + """" + "layername" + """" + ", " & vbCr
    strsql = strsql & "" + """" + xn + """" + " as " + """" + "usrnom" + """" + ", " & vbCr
    strsql = strsql & "" + """" + xp + """" + "." + """" + xt + """" + " as " + """" + "thenom" + """" + ", " & vbCr
    strsql = strsql & "" + """" + xv + """" + ", " & vbCr
    strsql = strsql & "" + """" + xp + """" + "" & vbCr
    strsql = strsql & "From " & vbCr
    strsql = strsql & "((" + """" + xm + """" + " Inner Join " + """" + xp + """" + " on " + """" + xp + """" + "." + """" + xw + """" + "=" + """" + xm + """" + "." + """" + xw + """" + ") " & vbCr
    strsql = strsql & "inner join " + """" + xu + """" + " on " + """" + xp + """" + "." + """" + xf + """" + "=" + """" + xu + """" + "." + """" + xf + """" + ") " & vbCr
    strsql = strsql & "inner join " + """" + xx + """" + " on " + """" + xp + """" + "." + """" + xf + """" + "=" + """" + xx + """" + "." + """" + xf + """" + " " & vbCr
    strsql = strsql & "Where " & vbCr
    strsql = strsql & "" + """" + xj + """" + " <= '128' " & vbCr
    strsql = strsql & "and" + """" + xm + """" + "." + """" + xn + """" + " = '" & usuario.UseName & "' " & vbCr
    strsql = strsql & "and" + """" + xm + """" + "." + """" + xt + """" + " = '" & FrmMain.ViewManager1.tvw.getActiveView & "' " & vbCr
    strsql = strsql & "Order By " & vbCr
    strsql = strsql & "" + """" + xp + """" + "." + """" + xt + """" + " " & vbCr
    
    'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", strsql, App.path & "\DEBUG.INI"
    
   End If
   
  
   
   ' Conn, adOpenDynamic, adLockOptimistic
     Set RsB = New ADODB.Recordset
     Dim strConn As String
   
   
       If frmCanvas.TipoConexao <> 4 Then
       

       
        Set RsB = mConn.execute(strsql)
       Else
       Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
Dim nStr As String

mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
nStr = frmCanvas.FunDecripta(Senha)
       
      strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"

    conexao.Open strConn
       Set RsB = conexao.execute(strsql)
    End If
    
    
     'WritePrivateProfileString "A", "A", strsql, App.path & "\DEBUG.INI"
     
   
 ' rs.Open strsql, mConn, adOpenDynamic, adLockOptimistic
    Lv.ColumnHeaders.Clear
    Lv.ListItems.Clear
    Lv.ColumnHeaders.Add , , "Usuário", 2000
    Lv.ColumnHeaders.Add , , "Tema", Lv.Width - 2000
    
    Dim TemaOld As String
    
    While Not RsB.EOF
      If TemaOld <> RsB!thenom Then
         Set itmx = Lv.ListItems.Add(, , UCase(RsB!UsrNom), , 1)
         itmx.SubItems(1) = UCase(RsB!thenom)
         itmx.Tag = RsB!LayerName
      End If
      TemaOld = RsB!thenom
      RsB.MoveNext
    Wend

    RsB.Close
    Set RsB = Nothing
    
    
    'VERIFICA SE HÁ LAYER ATUAL SELECIONADO, CASO SIM, POSSIBILITA EXPORTAR 'PLANO ATUAL'
    strLayerAtual = TCanvas.getCurrentLayer
    If strLayerAtual <> "" Then
        Me.optPlanoAtual.Caption = "Plano Atual: (" & strLayerAtual & ")"
    Else
        Me.optPlanoAtual.Enabled = False
    End If
    
    'VERIFICA SE HÁ ITENS SELECIONADOS NO MAPA, CASO SIM, POSSIBILITA EXPORTAR 'ITENS SELECIONADOS'
    If (TCanvas.getSelectCount(128) + TCanvas.getSelectCount(1) + TCanvas.getSelectCount(2) + TCanvas.getSelectCount(4)) = 0 Then
        Me.optSelecionados.Enabled = False
    End If
    
    Me.Show vbModal, OnnerForm
    
Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   
   PrintErro CStr(Me.Name), "Public Function Init()", CStr(Err.Number), CStr(Err.Description), True
   
End If
    
End Function

Private Sub cmdCancel_Click()
   'Set TCanvas = Nothing
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'LoozeXP1.EndWinXPCSubClassing
End Sub

Private Sub cmdOK_Click()
On Error GoTo Trata_Erro

Dim retval As String
Dim x As Integer
Dim strarquivo As String
Dim strFile As String
    '1 = Polígonos
    '2 = Linhas
    '4 = Pontos
    '128 = Textos
    
    If Trim(txtNomeArquivo.Text) = "" Then
        If Me.optArquivosSeparados.value = True Then
            MsgBox "Informe: Diretório\Pasta\ para salvar os arquivos.", vbExclamation
        Else
            MsgBox "Informe: Diretório\Pasta\Nome do Arquivo.", vbExclamation
        End If
        
        Exit Sub
    End If

    strDir = Me.txtNomeArquivo.Text
    x = Len(strDir)
    Do While Not x = 1                      'JONATHAS
        If mid(strDir, x, 1) = "\" Then     '13/02/09
            strDir = mid(strDir, 1, x)
            Exit Do
        End If
        x = x - 1
    Loop
    strFile = mid(txtNomeArquivo.Text, (Len(strDir) + 1), (Len(txtNomeArquivo.Text) - Len(strDir)))
    
    If x = 1 Then
        MsgBox "Caminho \ Nome do Arquivo incorreto.", vbExclamation
        Exit Sub
    End If
                
    If UCase(strFile) = ".DXF" Then
        MsgBox "Caminho \ Nome do Arquivo incorreto.", vbExclamation
        Exit Sub
    End If
                
    Screen.MousePointer = vbHourglass

    If Me.optPlanoAtual.value = True Then 'Layer selecionado
        If Me.optArquivoCompleto.value = True Then 'exporta um arquivo completo com todas as geometrias
            
            retval = Dir(txtNomeArquivo.Text)
            If retval <> "" Then 'verifica se o arquivo existe na pasta
                If MsgBox("O arquivo " & txtNomeArquivo.Text & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                    Screen.MousePointer = vbHourglass
                    TeExport1.exportDXF txtNomeArquivo.Text, strLayerAtual, 0, ""
                End If
            Else 'O arquivo ainda não existe na pasta
                Screen.MousePointer = vbHourglass
                TeExport1.exportDXF txtNomeArquivo.Text, strLayerAtual, 0, ""
            End If
            Screen.MousePointer = vbDefault
            MsgBox "Arquivo " & txtNomeArquivo.Text & " exportado com sucesso.   ", vbInformation
            Exit Sub
            
        Else 'exporta 1 arquivo para cada tipo de geometria do layer selecionado
                   
            If (Me.chkLinhas.value + Me.chkPoligonos.value + Me.chkTextos.value + Me.chkPontos.value) = 0 Then
                MsgBox "selecione os tipos de informação que deseja exportar.", vbInformation
                Exit Sub
            End If
                   
           
            If Me.chkLinhas.value = 1 Then 'EXPORTA LINHAS
                strarquivo = strDir & strLayerAtual & "_LINHAS.dxf"
                retval = Dir(strarquivo)
                If retval <> "" Then 'verifica se o arquivo existe na pasta
                    If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                        Screen.MousePointer = vbHourglass
                        TeExport1.exportDXF strarquivo, strLayerAtual, 2, ""
                    End If
                Else 'O arquivo ainda não existe na pasta
                    Screen.MousePointer = vbHourglass
                    TeExport1.exportDXF strarquivo, strLayerAtual, 2, ""
                End If
                Screen.MousePointer = vbDefault
            End If
            
            If Me.chkPontos.value = 1 Then 'EXPORTA PONTOS
                strarquivo = strDir & strLayerAtual & "_PONTOS.dxf"
                retval = Dir(strarquivo)
                If retval <> "" Then 'verifica se o arquivo existe na pasta
                    If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                        Screen.MousePointer = vbHourglass
                        TeExport1.exportDXF strarquivo, strLayerAtual, 4, ""
                    End If
                Else 'O arquivo ainda não existe na pasta
                    Screen.MousePointer = vbHourglass
                    TeExport1.exportDXF strarquivo, strLayerAtual, 4, ""
                End If
                Screen.MousePointer = vbDefault
            End If
            
            If Me.chkTextos.value = 1 Then 'EXPORTA TEXTOS
                strarquivo = strDir & strLayerAtual & "_TEXTO.dxf"
                retval = Dir(strarquivo)
                If retval <> "" Then 'verifica se o arquivo existe na pasta
                    If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                        Screen.MousePointer = vbHourglass
                        TeExport1.exportDXF strarquivo, strLayerAtual, 128, ""
                    End If
                Else 'O arquivo ainda não existe na pasta
                    Screen.MousePointer = vbHourglass
                    TeExport1.exportDXF strarquivo, strLayerAtual, 128, ""
                End If
                Screen.MousePointer = vbDefault
            End If
            
            If Me.chkPoligonos.value = 1 Then 'EXPORTA POLIGONOS
                strarquivo = strDir & strLayerAtual & "_POLIGONOS.dxf"
                retval = Dir(strarquivo)
                If retval <> "" Then 'verifica se o arquivo existe na pasta
                    If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                        Screen.MousePointer = vbHourglass
                        TeExport1.exportDXF strarquivo, strLayerAtual, 1, ""
                    End If
                Else 'O arquivo ainda não existe na pasta
                    Screen.MousePointer = vbHourglass
                    TeExport1.exportDXF strarquivo, strLayerAtual, 1, ""
                End If
                Screen.MousePointer = vbDefault
            End If
            
            MsgBox "Arquivo(s) exportado(s) com sucesso para o diretório " & strDir & ".   ", vbInformation
            
        End If
        
        
    ElseIf Me.optSelecionados.value = True Then ' Exporta itens selecionados no mapa
        
        ExportaComponentesSelecionados (Representacao)

        
    ElseIf Me.optSelecaoTemas.value = True Then
        Dim i, j As Integer
        For i = 1 To Lv.ListItems.count
            If Lv.ListItems.Item(i).Checked Then
                j = j + 1
                Exit For
            End If
        Next
        If j = 0 Then
            MsgBox "Não há itens selecionados.", vbExclamation
            Exit Sub
        End If
        blnExportaTemasSelecionados = False
        ExportaTemasSelecionados
            
        If blnExportaTemasSelecionados = True Then
            MsgBox "Temas selecionados exportados com sucesso.", vbInformation
        End If
    End If
    
    Screen.MousePointer = vbDefault
        
Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   
   PrintErro CStr(Me.Name), "Private Sub cmdOK_Click()", CStr(Err.Number), CStr(Err.Description), True
   Screen.MousePointer = vbNormal

End If

End Sub

Private Sub cmdOpenFile_Click()
On Error GoTo Trata_Erro
   With CDL
        .FileName = ".dxf"
        .Filter = "DXF (*.dxf)|*.dxf" '| MIF (*.mif;*.mif)|*.mif| shp(*.shp;*.shp)|*.shp"
        
        If Me.optSelecaoTemas.value = True Then
            CDL.DialogTitle = "Selecione o diretório que deseja salvar"
        Else
            CDL.DialogTitle = "Salvar como"
        End If
        .ShowSave
        txtNomeArquivo.Text = .FileName

        If Me.optArquivosSeparados.value = True And txtNomeArquivo.Text <> "" Or Me.optSelecaoTemas.value = True Then
                                                    'PROCEDIMENTO PARA CRIAR
            Dim x As Integer                        'NOME DE ARQUIVO
            Dim strarquivo As String
            strDir = Me.txtNomeArquivo.Text
            x = Len(strDir)
            Do While Not x = 1                      'JONATHAS
                If mid(strDir, x, 1) = "\" Then     '13/02/09
                    strDir = mid(strDir, 1, x)
                    Exit Do
                End If
                x = x - 1
            Loop
            MsgBox "Os nomes de arquivos serão criados automaticamente no diretório selecionado:" & Chr(13) & Chr(13) & strDir, vbInformation
            txtNomeArquivo.Text = strDir
            txtNomeArquivo.Locked = True
            
        End If
   
   End With
Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   
   Screen.MousePointer = vbNormal
   PrintErro CStr(Me.Name), "Private Sub cmdOpenFile_Click()", CStr(Err.Number), CStr(Err.Description), True
   
End If

End Sub

Private Sub ExportaComponentesSelecionados(pRep As Integer)
On Error GoTo Trata_Erro
   
    'CARREGA UM RECORDSET TEMPORÁRIO COM TODOS AS ID's DAS GEOMETRIAS SELECIONADAS NO MAPA
    
    Dim ExisteSelecao   As Boolean
    
   
    Dim a As Integer
    Dim ca As String
    Dim cb As String
    Dim cc As String
    ca = "GEOM_ID" ' alterado em 20/10/2010
    cb = "OBJECT_ID"
    cc = "X_TEMPOBJECTSSELECIONADOS"
    
    Dim strWhere As String
    If frmCanvas.TipoConexao <> 4 Then

    strWhere = " WHERE GEOM_ID IN(SELECT OBJECT_ID FROM X_TempObjectsSelecionados)"
    Else
    
strWhere = " WHERE " + """" + ca + """" + " IN(SELECT " + """" + cb + """" + " FROM " + """" + cc + ")"
       End If
        If optArquivoCompleto.value = True Then
            
            'GERA UM ARQUIVO COMPLETO COM TODO TIPO DE GEOMETRIA
            
            Dim retval As String
            retval = Dir(txtNomeArquivo.Text)
            If retval <> "" Then 'verifica se o arquivo existe na pasta
                If MsgBox("O arquivo " & txtNomeArquivo.Text & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                    Screen.MousePointer = vbHourglass
                    

                     TeExport1.exportDXF txtNomeArquivo.Text, strLayerAtual, 0, strWhere
                
                
                End If
            Else 'O arquivo ainda não existe na pasta
                Screen.MousePointer = vbHourglass
                TeExport1.exportDXF txtNomeArquivo.Text, strLayerAtual, 0, strWhere
            End If
            Screen.MousePointer = vbDefault
            MsgBox "Arquivo " & txtNomeArquivo.Text & " exportado com sucesso.   ", vbInformation
            Exit Sub
            
        Else ' EXPORTA CADA TIPO DE GEOMETRIA EM UM ARQUIVO
            
            Dim strarquivo As String
            Dim x As Integer
            strDir = Me.txtNomeArquivo.Text
            x = Len(strDir)
            Do While Not x = 1                      'JONATHAS
                If mid(strDir, x, 1) = "\" Then     '13/02/09
                    strDir = mid(strDir, 1, x)
                    Exit Do
                End If
                x = x - 1
            Loop
            If x = 1 Then
                MsgBox "Caminho \ Nome do Arquivo incorreto.", vbExclamation
                Exit Sub
            End If
                
            
            If Me.chkTextos.value = 1 Then 'EXPORTA TEXTOS
               
               
a = "X_TEMPOBJECTSSELECIONADOS"

     If frmCanvas.TipoConexao <> 4 Then
               Conn.execute "DELETE FROM X_TempObjectsSelecionados"
               Else
                Conn.execute "DELETE FROM " + """" + a + """"
               End If
               For a = 0 To TCanvas.getSelectCount(128) - 1 'TEXTOS SELECIONADOS
               
                       a = "OBJECT_ID_"
      b = "NRO_LIGACAO"
      c = "INSCRICAO_LOTE"
      d = "TIPO"
      e = "HIDROMETRADO"
      f = "ECONOMIAS"
      g = "CONSUMO_LPS"
      h = "X_TEMPOBJECTSSELECIONADOS"



     If frmCanvas.TipoConexao <> 4 Then
         
    Conn.execute ("INSERT INTO X_TempObjectsSelecionados (OBJECT_ID) VALUES ('" & TCanvas.getSelectGeoId(a, 128) & "') ")
     
     Else
     
     Conn.execute ("INSERT INTO " + """" + h + """" + "(" + """" + a + """" + ") VALUES ('" & TCanvas.getSelectGeoId(a, 128) & "') ")
     End If
                  
                  
                  
               Next
            
                strarquivo = strDir & strLayerAtual & "_TEXTOS_SELECIONADOS.dxf"
                retval = Dir(strarquivo)
                If retval <> "" Then 'verifica se o arquivo existe na pasta
                    If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                        Screen.MousePointer = vbHourglass
                        TeExport1.exportDXF strarquivo, strLayerAtual, 128, strWhere
                    End If
                Else 'O arquivo ainda não existe na pasta
                    Screen.MousePointer = vbHourglass
                    TeExport1.exportDXF strarquivo, strLayerAtual, 128, strWhere
                End If
                Screen.MousePointer = vbDefault
                MsgBox "Arquivo " & strarquivo & " exportado com sucesso.   ", vbInformation
                Screen.MousePointer = vbHourglass
            End If
                
            If Me.chkLinhas.value = 1 Then 'EXPORTA LINHAS
            a = "OBJECT_ID_"
                 If frmCanvas.TipoConexao <> 4 Then
               Conn.execute "DELETE FROM X_TempObjectsSelecionados"
               Else
                Conn.execute "DELETE FROM " + a + ""
               End If
               
               For a = 0 To TCanvas.getSelectCount(2) - 1 'LINHAS SELECIONADAS
               
                                 a = "OBJECT_ID_"
      b = "NRO_LIGACAO"
      c = "INSCRICAO_LOTE"
      d = "TIPO"
      e = "HIDROMETRADO"
      f = "ECONOMIAS"
      g = "CONSUMO_LPS"
      h = "X_TEMPOBJECTSSELECIONADOS"



     If frmCanvas.TipoConexao <> 4 Then
         
     Conn.execute ("INSERT INTO X_TempObjectsSelecionados (OBJECT_ID) VALUES ('" & TCanvas.getSelectGeoId(a, 2) & "') ")
     Else
     
       Conn.execute ("INSERT INTO " + """" + h + """" + "(" + """" + a + """" + ") VALUES ('" & TCanvas.getSelectGeoId(a, 2) & "') ")
     End If
               
               
                 
               Next
                
                strarquivo = strDir & strLayerAtual & "_LINHAS_SELECIONADAS.dxf"
                retval = Dir(strarquivo)
                If retval <> "" Then 'verifica se o arquivo existe na pasta
                    If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                        Screen.MousePointer = vbHourglass
                        TeExport1.exportDXF strarquivo, strLayerAtual, 2, strWhere
                    End If
                Else 'O arquivo ainda não existe na pasta
                    Screen.MousePointer = vbHourglass
                    TeExport1.exportDXF strarquivo, strLayerAtual, 2, strWhere
                End If
                Screen.MousePointer = vbDefault
                MsgBox "Arquivo " & strarquivo & " exportado com sucesso.   ", vbInformation
                Screen.MousePointer = vbHourglass
            End If

            If Me.chkPontos.value = 1 Then 'EXPORTA PONTOS
            a = "X_TEMPOBJECTSSELECIONADOS"
            
     If frmCanvas.TipoConexao <> 4 Then
               Conn.execute "DELETE FROM X_TempObjectsSelecionados"
               Else
               Conn.execute "DELETE FROM " + """" + a + """"
               End If
               For a = 0 To TCanvas.getSelectCount(4) - 1 'LINHAS SELECIONADAS
      a = "OBJECT_ID_"
      b = "NRO_LIGACAO"
      c = "INSCRICAO_LOTE"
      d = "TIPO"
      e = "HIDROMETRADO"
      f = "ECONOMIAS"
      g = "CONSUMO_LPS"
      h = "X_TEMPOBJECTSSELECIONADOS"



     If frmCanvas.TipoConexao <> 4 Then
         
          Conn.execute ("INSERT INTO X_TempObjectsSelecionados (OBJECT_ID) VALUES ('" & TCanvas.getSelectGeoId(a, 4) & "') ")
     Else
     
       Conn.execute ("INSERT INTO " + """" + h + """" + "(" + """" + a + """" + ") VALUES ('" & TCanvas.getSelectGeoId(a, 4) & "') ")
   
     End If
               

               Next
            
                strarquivo = strDir & strLayerAtual & "_PONTOS_SELECIONADOS.dxf"
                retval = Dir(strarquivo)
                If retval <> "" Then 'verifica se o arquivo existe na pasta
                    If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                        Screen.MousePointer = vbHourglass
                        TeExport1.exportDXF strarquivo, strLayerAtual, 4, strWhere
                    End If
                Else 'O arquivo ainda não existe na pasta
                    Screen.MousePointer = vbHourglass
                    TeExport1.exportDXF strarquivo, strLayerAtual, 4, strWhere
                End If
                Screen.MousePointer = vbDefault
                MsgBox "Arquivo " & strarquivo & " exportado com sucesso.   ", vbInformation
                Screen.MousePointer = vbHourglass
            End If
            
            If Me.chkPoligonos.value = 1 Then 'EXPORTA POLIGONOS
a = "X_TEMPOBJECTSSELECIONADOS"

 If frmCanvas.TipoConexao <> 4 Then
                Conn.execute "DELETE FROM X_TempObjectsSelecionados"
               Else
              Conn.execute "DELETE FROM " + """" + a + """"
               End If
               
               For a = 0 To TCanvas.getSelectCount(1) - 1 'LINHAS SELECIONADAS
               
               a = "OBJECT_ID_"
      b = "NRO_LIGACAO"
      c = "INSCRICAO_LOTE"
      d = "TIPO"
      e = "HIDROMETRADO"
      f = "ECONOMIAS"
      g = "CONSUMO_LPS"
      h = "X_TEMPOBJECTSSELECIONADOS"



     If frmCanvas.TipoConexao <> 4 Then
         
         Conn.execute ("INSERT INTO X_TempObjectsSelecionados (OBJECT_ID) VALUES ('" & TCanvas.getSelectGeoId(a, 1) & "') ")
     Else
     
       Conn.execute ("INSERT INTO " + """" + h + """" + "(" + """" + a + """" + ") VALUES ('" & TCanvas.getSelectGeoId(a, 1) & "') ")
   
     End If
               
                  
               Next

            
                strarquivo = strDir & strLayerAtual & "_POLIGONOS_SELECIONADOS.dxf"
                retval = Dir(strarquivo)
                If retval <> "" Then 'verifica se o arquivo existe na pasta
                    If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                        Screen.MousePointer = vbHourglass
                        TeExport1.exportDXF strarquivo, strLayerAtual, 1, strWhere
                    End If
                Else 'O arquivo ainda não existe na pasta
                    Screen.MousePointer = vbHourglass
                    TeExport1.exportDXF strarquivo, strLayerAtual, 1, strWhere
                End If
                Screen.MousePointer = vbDefault
                MsgBox "Arquivo " & strarquivo & " exportado com sucesso.   ", vbInformation

            End If
        
        End If
    
    
Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   
   Screen.MousePointer = vbNormal
   PrintErro CStr(Me.Name), "Private Sub ExportaComponentesSelecionados", CStr(Err.Number), CStr(Err.Description), True

End If

End Sub

Private Function ExportaTemasSelecionados()
    On Error GoTo Trata_Erro
    
    Dim a As Integer
    Dim mPath As String
    Dim mWhere As String
    Dim b As Integer
    Dim strTema, strFiltro, strarquivo, retval As String
    
    xx = "te_layer"
   xu = "te_representation"
   xt = "name"
   xq = "theme_id"
   xp = "te_theme"
   xw = "view_id"
   xf = "layer_id"
   xj = "geom_type"
   xn = "user_name"
   xm = "te_view"
   xv = "generate_attribute_where"
   
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If frmCanvas.TipoConexao <> 4 Then
    rs.Open "SELECT l.name, t.name as tema,t.theme_id, t.generate_attribute_where from Te_Theme t inner join te_layer l on t.layer_id=l.layer_id", Conn, adOpenDynamic

Else
    rs.Open "SELECT " + """" + xx + """" + "." + """" + xt + """" + ", " + """" + xp + """" + "." + """" + xt + """" + " as " + """" + "tema" + """" + "," + """" + xp + """" + "." + """" + xq + """" + ", " + """" + xp + """" + "." + """" + xv + """" + " from " + """" + xp + """" + " inner join" + """" + xx + """" + " on " + """" + xp + """" + "." + """" + xf + """" + "=" + """" + xx + """" + "." + """" + xf + """" + "", conexao, adOpenDynamic, adLockOptimistic

End If

    For a = 1 To Lv.ListItems.count
        If Lv.ListItems.Item(a).Checked Then
            rs.Filter = "tema='" & Lv.ListItems.Item(a).ListSubItems(1) & "'"
            
            'MsgBox rs!Name & " " & rs!TEMA & " " & rs!Theme_id
            
            If Not rs.EOF Then
                With txtNomeArquivo
                   'VERIFICA SE O NOME DO TEMA PODE SER UM NOME VÁLIDO DE ARQUIVO
                    
                    mPath = Lv.ListItems.Item(a).SubItems(1)
                    For b = 1 To Len(mPath)
                        If Not ((Asc(mid(mPath, b, 1)) >= 48 And Asc(mid(mPath, b, 1)) <= 57) Or _
                            (Asc(mid(mPath, b, 1)) >= 65 And Asc(mid(mPath, b, 1)) <= 90) Or _
                            (Asc(mid(mPath, b, 1)) >= 97 And Asc(mid(mPath, b, 1)) <= 122)) Then
                            mPath = Replace(mPath, mid(mPath, b, 1), "_", 1, 1)
                        End If
                    Next
                    strTema = mPath
            
                    strFiltro = " WHERE " & Trim(rs!generate_attribute_where)
                    If Trim(strFiltro) = "WHERE" Then
                        strFiltro = ""
                    End If
                    'trocar o nome do layer no loop
        
                    If Me.optArquivoCompleto.value = True Then
                        strarquivo = strDir & Lv.ListItems.Item(a).Tag & "_" & strTema & "_COMPLETO.dxf"
                        'strarquivo = Me.txtNomeArquivo.Text & Lv.ListItems.Item(a).Tag & "_" & strTema & "_COMPLETO.dxf"
                        retval = Dir(strarquivo)
                        If retval <> "" Then 'verifica se o arquivo existe na pasta
                            If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                                Screen.MousePointer = vbHourglass
                                    
                                TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 0, strFiltro

                            End If
                        Else 'O arquivo ainda não existe na pasta
                            Screen.MousePointer = vbHourglass
                            
                            TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 0, strFiltro

                        End If
                        Screen.MousePointer = vbDefault
                        
                    Else ' ARQUIVO SEPARADO
                        
                        If Me.chkPoligonos.value = 1 Then 'EXPORTA POLIGONOS
                            strarquivo = Me.txtNomeArquivo.Text & Lv.ListItems.Item(a).Tag & "_TEMA_" & strTema & "_POLIGONOS.dxf"
                            retval = Dir(strarquivo)
                            If retval <> "" Then 'verifica se o arquivo existe na pasta
                                If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                                    Screen.MousePointer = vbHourglass
                                        
                                    TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 1, strFiltro
                                
                                End If
                            Else 'O arquivo ainda não existe na pasta
                                Screen.MousePointer = vbHourglass
                                If strFiltro = "" Then
                                    TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 1, ""
                                Else
                                    TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 1, strFiltro
                                End If
                            End If
                            Screen.MousePointer = vbDefault
                        End If
        
                        If Me.chkLinhas.value = 1 Then 'EXPORTA LINHAS
                            strarquivo = Me.txtNomeArquivo.Text & Lv.ListItems.Item(a).Tag & "_TEMA_" & strTema & "_LINHAS.dxf"
                            retval = Dir(strarquivo)
                            If retval <> "" Then 'verifica se o arquivo existe na pasta
                                If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                                    Screen.MousePointer = vbHourglass
                                    
                                    TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 2, strFiltro
                                
                                End If
                            Else 'O arquivo ainda não existe na pasta
                                Screen.MousePointer = vbHourglass
                                    
                                TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 2, strFiltro

                            End If
                            Screen.MousePointer = vbDefault
                        End If
                        
                        If Me.chkPontos.value = 1 Then 'EXPORTA PONTOS
                            strarquivo = Me.txtNomeArquivo.Text & Lv.ListItems.Item(a).Tag & "_TEMA_" & strTema & "_PONTOS.dxf"
                            retval = Dir(strarquivo)
                            If retval <> "" Then 'verifica se o arquivo existe na pasta
                                If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                                    Screen.MousePointer = vbHourglass

                                    TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 4, strFiltro

                                End If
                            Else 'O arquivo ainda não existe na pasta
                                Screen.MousePointer = vbHourglass
                                    
                                TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 4, strFiltro

                            End If
                            Screen.MousePointer = vbDefault
                        End If
                        
                        If Me.chkTextos.value = 1 Then 'EXPORTA TEXTOS
                            strarquivo = Me.txtNomeArquivo.Text & Lv.ListItems.Item(a).Tag & "_TEMA_" & strTema & "_TEXTO.dxf"
                            retval = Dir(strarquivo)
                            If retval <> "" Then 'verifica se o arquivo existe na pasta
                                If MsgBox("O arquivo " & strarquivo & " já existe na pasta selecionada." & Chr(13) & Chr(13) & "Deseja sobrescrever?", vbExclamation + vbYesNo + vbDefaultButton2, "") = vbYes Then
                                    Screen.MousePointer = vbHourglass
                                    
                                    TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 128, strFiltro
                                
                                End If
                            Else 'O arquivo ainda não existe na pasta
                                Screen.MousePointer = vbHourglass
                                
                                TeExport1.exportDXF strarquivo, Lv.ListItems.Item(a).Tag, 128, strFiltro
                                
                            End If
                            Screen.MousePointer = vbDefault
                        End If
                        
                    End If
                    
                End With
            End If
        End If
      
    Next
    Screen.MousePointer = vbDefault
    rs.Close
    Set rs = Nothing
    blnExportaTemasSelecionados = True

Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   
   Screen.MousePointer = vbNormal
   PrintErro CStr(Me.Name), "Private Function ExportaTemasSelecionados()", CStr(Err.Number), CStr(Err.Description), True

End If

End Function

Private Sub optSelecionados_Click()
    habilitaLista
    
    
    Me.optArquivosSeparados.value = True
    Me.optArquivoCompleto.Visible = False
    Me.Frame5.Visible = False
    Me.Check1.Visible = False
    Me.Check2.Visible = False
    Me.Check3.Visible = False
    Me.Check4.Visible = False
    
End Sub

Private Sub optPlanoAtual_Click()
    
    habilitaLista

    Me.optArquivoCompleto.Enabled = True
    Me.optArquivoCompleto.value = True
    Me.Frame5.Enabled = True
    
    Me.optArquivoCompleto.Visible = True
    Me.Frame5.Visible = True
    Me.Check1.Visible = True
    Me.Check2.Visible = True
    Me.Check3.Visible = True
    Me.Check4.Visible = True
    
End Sub

Private Sub optSelecaoTemas_Click()
    habilitaLista
    
    Me.optArquivoCompleto.Enabled = True
    Me.optArquivoCompleto.value = True
    Me.Frame5.Enabled = True
    
    Me.optArquivoCompleto.Visible = True
    Me.Frame5.Visible = True
    Me.Check1.Visible = True
    Me.Check2.Visible = True
    Me.Check3.Visible = True
    Me.Check4.Visible = True
    
    
End Sub

Private Function habilitaLista()
    If Me.optSelecaoTemas.value = False Then
        Me.Frame1.Enabled = False
        Me.Lv.Enabled = False
        Dim i As Integer
        Dim j As Integer
        j = 1
        i = Lv.ListItems.count
        Do While Not j = i + 1
            Lv.ListItems.Item(j).Checked = False
            j = j + 1
        Loop
    Else
        Me.Frame1.Enabled = True
        Me.Lv.Enabled = True
    End If
End Function

Private Sub optArquivoCompleto_Click()
    Me.chkLinhas.value = 0
    Me.chkPoligonos.value = 0
    Me.chkPontos.value = 0
    Me.chkTextos.value = 0
    Me.optArquivoCompleto.value = True
    Frame3.Enabled = False
    Frame5.Enabled = True
    Me.txtNomeArquivo.Text = ""
    Me.txtNomeArquivo.Locked = False
End Sub

Private Sub Check1_Click()
    Me.Check1.value = 1
    Me.txtNomeArquivo.Locked = False
End Sub
Private Sub Check2_Click()
    Me.Check2.value = 1
    Me.txtNomeArquivo.Locked = False
End Sub
Private Sub Check3_Click()
    Me.Check3.value = 1
    Me.txtNomeArquivo.Locked = False
End Sub
Private Sub Check4_Click()
    Me.Check4.value = 1
    Me.txtNomeArquivo.Locked = False
End Sub

Private Sub optArquivosSeparados_Click()
    Frame3.Enabled = True
    Me.txtNomeArquivo.Text = ""
    Me.txtNomeArquivo.Locked = True
End Sub

Private Sub chkLinhas_Click()
    Me.optArquivosSeparados.value = True
    Me.txtNomeArquivo.Locked = True
End Sub

Private Sub chkPoligonos_Click()
    Me.optArquivosSeparados.value = True
    Me.txtNomeArquivo.Locked = True
End Sub

Private Sub chkPontos_Click()
    Me.optArquivosSeparados.value = True
    Me.txtNomeArquivo.Locked = True
End Sub

Private Sub chkTextos_Click()
    Me.optArquivosSeparados.value = True
    Me.txtNomeArquivo.Locked = True
End Sub




