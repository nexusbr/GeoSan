VERSION 5.00
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F03ABD98-7B60-43E4-9934-DA5F0D19FDAC}#1.0#0"; "TeComViewManager.dll"
Object = "{EE78E37B-39BE-42FA-80B7-E525529739F7}#1.0#0"; "TeComViewDatabase.dll"
Begin VB.UserControl ViewManager 
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ScaleHeight     =   4605
   ScaleWidth      =   3795
   Begin MSComctlLib.ImageList imgLista 
      Left            =   3150
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   26
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewManager.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewManager.ctx":0386
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewManager.ctx":0738
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewManager.ctx":0BB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewManager.ctx":0FCB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tv 
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   8017
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgLista"
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin TeComViewDatabaseLibCtl.TeViewDatabase TeViewDatabase1 
      Left            =   2160
      OleObjectBlob   =   "ViewManager.ctx":13D8
      Top             =   2400
   End
   Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection2 
      Left            =   2160
      OleObjectBlob   =   "ViewManager.ctx":13FC
      Top             =   1440
   End
   Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
      Left            =   2760
      OleObjectBlob   =   "ViewManager.ctx":1420
      Top             =   2280
   End
   Begin TECOMVIEWMANAGERLibCtl.TeViewManager TeViewManager1 
      Left            =   2160
      OleObjectBlob   =   "ViewManager.ctx":1444
      Top             =   2040
   End
   Begin VB.Menu mnuTheme 
      Caption         =   "Themes"
      Visible         =   0   'False
      Begin VB.Menu mnuThemeProperties 
         Caption         =   "Propriedades"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRenameTheme 
         Caption         =   "Renomear"
      End
      Begin VB.Menu mnuNewTheme 
         Caption         =   "Novo"
      End
      Begin VB.Menu mnuDeleteTheme 
         Caption         =   "Excluir"
      End
      Begin VB.Menu mnuThemeOn 
         Caption         =   "Ligar/Desligar"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "Temas"
         Begin VB.Menu mnuAllTheme_on 
            Caption         =   "Ligar todos os temas"
         End
         Begin VB.Menu mnuAllText_on 
            Caption         =   "Ligar todos os textos"
         End
         Begin VB.Menu mnuAllTheme_off 
            Caption         =   "Desligar todos os temas"
         End
         Begin VB.Menu mnuAllText_off 
            Caption         =   "Desligar todos os textos"
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSnap 
         Caption         =   "Snap (lig/desl)"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepresentation 
         Caption         =   "Representações"
         Begin VB.Menu mnuRepVisible 
            Caption         =   "Ligar/Desligar"
            Begin VB.Menu mnuPolygons 
               Caption         =   "Polígonos "
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuLines 
               Caption         =   "Linhas "
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuPoints 
               Caption         =   "Pontos"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuTexts 
               Caption         =   "Textos"
               Checked         =   -1  'True
            End
         End
         Begin VB.Menu mnuRepEnableSub 
            Caption         =   "Ativar/Desativar (não salva)"
            Begin VB.Menu mnuRepEnable 
               Caption         =   "Polígonos"
               Checked         =   -1  'True
               Index           =   1
            End
            Begin VB.Menu mnuRepEnable 
               Caption         =   "Linhas"
               Checked         =   -1  'True
               Index           =   2
            End
            Begin VB.Menu mnuRepEnable 
               Caption         =   "Pontos"
               Checked         =   -1  'True
               Index           =   4
            End
            Begin VB.Menu mnuRepEnable 
               Caption         =   "Textos"
               Checked         =   -1  'True
               Index           =   128
            End
            Begin VB.Menu mnuRepEnable 
               Caption         =   "Ativar Todas"
               Index           =   512
            End
         End
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "Vista"
      End
   End
End
Attribute VB_Name = "ViewManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum mImg
  moff = 1
  mOn = 2
  mOnSet = 3
  mOnSnap = 4
  mOnSnapLock = 5
End Enum
Public Event onReset(ViewName As String)
Public tvw As Object, tvm As Object, tcs As Object, mConn As ADODB.Connection, Provider As Integer
Private ThemeName As String, indexTheme As Integer
Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
Dim tax As TeAcXConnection
Dim manager As TeViewManager
Dim carrega As Integer
Dim database As TeViewDatabase
Dim usuario As String
Dim strConn As String
Dim cont As Integer
Dim conexao As New ADODB.Connection





Public Sub Start()






   Set conn = mConn
   TypeConn = Provider
  
  If carrega <> 10 Then
   If TypeConn = 4 Then

   Set tax = TeAcXConnection1
Set manager = TeViewManager1
   Set database = TeViewDatabase1
   mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.Path & "\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.Path & "\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.Path & "\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.Path & "\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.Path & "\GEOSAN.ini")
usuario = ReadINI("CONEXAO", "USER", App.Path & "\GEOSAN.ini")
decriptada = FunDecripta(Senha)
 strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + decriptada + "; ByteaAsLongVarBinary=1;"

    conexao.Open strConn

 tax.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA
 
  
 manager.UserName = usuario
 manager.Provider = 4
 manager.Connection = tax.objectConnection_

  manager.Start

 
   
 database.UserName = usuario
 database.Provider = 4
 database.Connection = tax.objectConnection_

 
 carrega = 10
 End If
  End If
   
   LoadTheme
End Sub

Public Function TvSetCurrentLayer(LayerName As String) As Boolean
   Dim a As Integer
   For a = 1 To Tv.Nodes.Count
      If Tv.Nodes.Item(a).Tag = LayerName Then
         Tv_NodeClick Tv.Nodes.Item(a)
         TvSetCurrentLayer = True
      End If
   Next
End Function

Public Sub UserControl_Resize()
   Tv.Top = 0
   Tv.Left = 0
   Tv.Height = Height
   Tv.Width = Width
End Sub
'Esta Subrotina roda logo no início quando vai carregar a árvore de temas a direita do GeoSan
'Nela são carregados os temas ativos da vista do usuário e salvos no arquivo FTema.txt
Private Sub LoadTheme()
    On Error GoTo Trata_Erro
    Dim aa As String
    Dim b As String
    Dim c As String
    Dim d As String
    Dim e As String
    Dim f As String
    Dim g As String
    Dim h As String
    Dim i As String
    Dim j As String

    Call SaveLoadGlobalData("C:\Arquivos de programas\GeoSan" + "\controles\variaveisGlobais.txt", False) 'Carrega as variáveis globais do GeoSan.exe
    aa = "te_theme"
    b = "te_view"
    c = "theme_id"
    d = "generate_attribute_where"
    e = "name"
    f = "te_layer"
    g = "view_id"
    h = "layer_id"
    i = "user_name"
    j = "priority"
    'Me.Caption = "Vista: " & tvw.getActiveView
    Tv.Nodes.Clear
    Dim rs As ADODB.Recordset
    If TypeConn <> 4 Then
        ' Dim ff As String
        ' ff = "Select t.THEME_ID as TID_T, T.GENERATE_ATTRIBUTE_WHERE AS GAW_T, t.Name as name_T, v.Name as name_V, l.name as name_l From ((Te_Theme t inner join Te_View v on v.view_id=t.view_id) inner join te_layer l on t.layer_id=l.layer_id) where v.name='" & tvw.getactiveview & "' and v.user_name='" & tvw.UserName & "' order by t.priority"
        ' WritePrivateProfileString "A", "A", ff, App.Path & "\DEBUG.INI"
        Set rs = mConn.Execute("Select t.THEME_ID as " + """" + "TID_T" + """" + ", T.GENERATE_ATTRIBUTE_WHERE AS " + """" + "GAW_T" + """" + ", t.Name as " + """" + "name_T" + """" + ", v.Name as " + """" + "name_V" + """" + ", l.name as " + """" + "name_l" + """" + " From ((Te_Theme t inner join Te_View v on v.view_id=t.view_id) inner join te_layer l on t.layer_id=l.layer_id) where v.name='" & tvw.getActiveView & "' and v.user_name='" & tvw.UserName & "' order by t.priority")
    Else
        ' Select "te_theme"."theme_id" as TID_T, "te_theme"."generate_attribute_where" AS GAW_T, "te_theme"."name" as name_T, "te_view"."name" as name_V,
        ' "te_layer"."name" as name_l From (("te_theme" inner join "te_view" on "te_view"."view_id"="te_theme"."view_id") inner join "te_layer"  on "te_theme"."layer_id"="te_layer"."layer_id")
        ' where "te_view"."name"='Nova Vista'
        'and "te_view"."user_name"='Administrador' order by "te_theme"."priority";
        'Dim aaa As String
        'aaa = "Select " + """" + aa + """" + "." + """" + c + """" + " as " + """" + "TID_T" + """" + ", " + """" + aa + """" + "." + """" + d + """" + " AS " + """" + "GAW_T" + """" + ", " + """" + aa + """" + "." + """" + e + """" + " as " + """" + "name_T" + """" + ", " + """" + b + """" + "." + """" + e + """" + " as " + """" + "name_V" + """" + ", " + """" + f + """" + "." + """" + e + """" + " as " + """" + "name_l" + """" + " From ((" + """" + aa + """" + "  inner join " + """" + b + """" + "  on " + """" + b + """" + "." + """" + g + """" + "=" + """" + aa + """" + "." + """" + g + """" + ") inner join " + """" + f + """" + "  on " + """" + aa + """" + "." + """" + h + """" + "=" + """" + f + """" + " ." + """" + h + """" + " ) where " + """" + b + """" + " ." + """" + e + """" + " ='" & tvw.getactiveview & "' and " + """" + b + """" + " ." + """" + i + """" + "='" & tvw.UserName & "' order by " + """" + aa + """" + "." + """" + j + """" + ""
        'WritePrivateProfileString "A", "A", aaa, App.Path & "\DEBUG.INI"
        'MsgBox "ARQUIVO DEBUG SALVO"
        'Dim aaaa2 As String
        'aaaa2 = "Select " + """" + aa + """" + "." + """" + c + """" + " as " + """" + "TID_T" + """" + ", " + """" + aa + """" + "." + """" + d + """" + " AS " + """" + "GAW_T" + """" + ", " + """" + aa + """" + "." + """" + e + """" + " as " + """" + "name_t" + """" + ", " + """" + b + """" + "." + """" + e + """" + " as " + """" + "name_V" + """" + ", " + """" + f + """" + "." + """" + e + """" + " as " + """" + "name_l" + """" + " From ((" + """" + aa + """" + "  inner join " + """" + b + """" + "  on " + """" + b + """" + "." + """" + g + """" + "=" + """" + aa + """" + "." + """" + g + """" + ") inner join " + """" + f + """" + "  on " + """" + aa + """" + "." + """" + h + """" + "=" + """" + f + """" + " ." + """" + h + """" + " ) where " + """" + b + """" + " ." + """" + e + """" + " ='" & tvw.getactiveview & "' and " + """" + b + """" + " ." + """" + i + """" + "='" & tvw.UserName & "' order by " + """" + aa + """" + "." + """" + j + """" + ""
        'WritePrivateProfileString "A", "A", aaaa2, App.Path & "\DEBUG.INI"
        Set rs = mConn.Execute("Select " + """" + aa + """" + "." + """" + c + """" + " as " + """" + "TID_T" + """" + ", " + """" + aa + """" + "." + """" + d + """" + " AS " + """" + "GAW_T" + """" + ", " + """" + aa + """" + "." + """" + e + """" + " as " + """" + "name_t" + """" + ", " + """" + b + """" + "." + """" + e + """" + " as " + """" + "name_V" + """" + ", " + """" + f + """" + "." + """" + e + """" + " as " + """" + "name_l" + """" + " From ((" + """" + aa + """" + "  inner join " + """" + b + """" + "  on " + """" + b + """" + "." + """" + g + """" + "=" + """" + aa + """" + "." + """" + g + """" + ") inner join " + """" + f + """" + "  on " + """" + aa + """" + "." + """" + h + """" + "=" + """" + f + """" + " ." + """" + h + """" + " ) where " + """" + b + """" + " ." + """" + e + """" + " ='" & tvw.getActiveView & "' and " + """" + b + """" + " ." + """" + i + """" + "='" & tvw.UserName & "' order by " + """" + aa + """" + "." + """" + j + """" + "")
        'Dim ay As String
        'ay = "Select " + """" + aa + """" + "." + """" + c + """" + " as " + """" + "TID_T" + """" + ", " + """" + aa + """" + "." + """" + d + """" + " AS " + """" + "GAW_T" + """" + ", " + """" + aa + """" + "." + """" + e + """" + " as " + """" + "name_t" + """" + ", " + """" + b + """" + "." + """" + e + """" + " as " + """" + "name_V" + """" + ", " + """" + f + """" + "." + """" + e + """" + " as " + """" + "name_l" + """" + " From ((" + """" + aa + """" + "  inner join " + """" + b + """" + "  on " + """" + b + """" + "." + """" + g + """" + "=" + """" + aa + """" + "." + """" + g + """" + ") inner join " + """" + f + """" + "  on " + """" + aa + """" + "." + """" + h + """" + "=" + """" + f + """" + " ." + """" + h + """" + " ) where " + """" + b + """" + " ." + """" + e + """" + " ='" & tvw.getactiveview & "' and " + """" + b + """" + " ." + """" + i + """" + "='" & tvw.UserName & "' order by " + """" + aa + """" + "." + """" + j + """" + ""
        'MsgBox "ARQUIVO DEBUG SALVO"
        'WritePrivateProfileString "A", "A", ay, App.Path & "\DEBUG.INI"
        ' While Not rs.EOF
        'MsgBox ("Entrei")
        ' rs.MoveNext
        '    Wend
    End If
    'Set Rs = mConn.Execute("Select t.Name as name_T, v.Name as name_V, l.name as name_l From ((Te_Theme t inner join Te_View v on v.view_id=t.view_id) inner join te_layer l on t.layer_id=l.layer_id) where v.name='" & tvw.getactiveview & "' and v.user_name='" & tvw.userName & "' order by t.priority desc")
    Dim a As Integer, T As Node
    ' With tvw
    If tvw.getActiveView = "" Then
        Exit Sub
    End If
    Close #3
    Open glo.diretorioGeoSan & "\Controles\FTema.txt" For Output As #3 ' GRAVA UM ARQUIVO EXTERNO QUE SERÁ USADO COMO LOG
    While Not rs.EOF
        'THEME_ID; STRING DO FILTRO
        '                 ID DO TEMA          ;            NOME DA LAYER        ;       COMANDO DO FILTRO
        Print #3, rs.Fields("TID_T").Value & ";" & rs.Fields("name_t").Value & ";" & rs.Fields("GAW_T").Value
        Set T = Tv.Nodes.Add(, , rs.Fields("name_t").Value, rs.Fields("name_t").Value, GetVisibledTheme(rs.Fields("name_t").Value))
        If tvw.existPoint(tvw.getActiveView(), rs.Fields("name_t").Value) Then
            T.ForeColor = tvw.getPointColor(tvw.getActiveView(), rs.Fields("name_t").Value)
        ElseIf tvw.existLine(tvw.getActiveView(), rs.Fields("name_t").Value) Then
            T.ForeColor = tvw.getLineColor(tvw.getActiveView(), rs.Fields("name_t").Value)
        ElseIf tvw.existPolygon(tvw.getActiveView(), rs.Fields("name_t").Value) Then
            If tvw.getPolygonStyle(tvw.getActiveView(), rs.Fields("name_t").Value) > 0 Then
                T.ForeColor = tvw.getPolygonColor(tvw.getActiveView(), rs.Fields("name_t").Value)
            Else
                T.ForeColor = tvw.getPolygonContourColor(tvw.getActiveView(), rs.Fields("name_t").Value)
            End If
        End If
        T.Tag = tvw.getLayerNameFromTheme(tvw.getActiveView(), rs.Fields("name_t").Value)
        rs.MoveNext
    Wend
    Close #3
    'For A = 0 To .getThemeCount(.getActiveView()) - 1
    '         Set T = Tv.Nodes.Add(, , .getThemeName(A), .getThemeName(A), GetVisibledTheme(.getThemeName(A)))
    '         If .existPoint(.getActiveView(), .getThemeName(A)) Then
    '            T.ForeColor = .getPointColor(.getActiveView(), .getThemeName(A))
    '         ElseIf .existLine(.getActiveView(), .getThemeName(A)) Then
    '            T.ForeColor = .getLineColor(.getActiveView(), .getThemeName(A))
    '         ElseIf .existPolygon(.getActiveView(), .getThemeName(A)) Then
    '            If .getPolygonStyle(.getActiveView(), .getThemeName(A)) > 0 Then
    '               T.ForeColor = .getPolygonStyle(.getActiveView(), .getThemeName(A))
    '            Else
    '               T.ForeColor = .getPolygonContourColor(.getActiveView(), .getThemeName(A))
    '            End If
    '         End If
    '         T.Tag = .getLayerNameFromTheme(.getActiveView(), .getThemeName(A))
    'Next
    ' End With
    rs.Close
    aa = "te_theme"
    g = "view_id"
    j = "priority"
    'MsgBox "ressequenciando temas"
    '******** RESSEQUENCIAMENTO DE TEMAS
    Dim conexao As New ADODB.Connection
    Dim strConn As String
    If TypeConn = 4 Then
        strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + decriptada + "; ByteaAsLongVarBinary=1;"
        conexao.Open strConn
    End If
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim ViewOld As String, ViewAtual As String
    If TypeConn <> 4 Then
        rs2.Open "SELECT * FROM TE_THEME ORDER BY VIEW_ID, PRIORITY", conn, adOpenDynamic, adLockOptimistic
    Else
        rs2.Open "SELECT * FROM " + """" + "te_theme" + """" + " ORDER BY " + """" + "view_id" + """" + "," + """" + "priority" + """", conexao, adOpenDynamic, adLockOptimistic
    End If
    Do While Not rs2.EOF = True
        If ViewOld = rs2!view_id Then
            cont = cont + 1
        Else
            cont = 0
        End If
        rs2!priority = cont
        rs2.Update
        ViewOld = rs2!view_id
        rs2.MoveNext
    Loop
    rs2.Close
    '******** RESSEQUENCIAMENTO DE TEMAS FIM
     'conexao.Close
    Set rs = Nothing
    Set rs2 = Nothing
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #1
        Open App.Path & "\GeoSanLog.txt" For Append As #1
        'Print #1, Now & " - NxViewManager - frmTheme - Private Sub cmdOK_Click() - " & Err.Number & " - " & Err.Description
        Close #1
        'MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & "Não foi possível estabelecer a conexão" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrencia.", vbInformation
    End If
End Sub

Private Function LoadThemeProperties() As Boolean

   If frmTheme.Init(tvw, ThemeName, Tv.Nodes.Item(indexTheme).Tag) Then
      tcs.ResetView
      LoadTheme
      tcs.plotView
   End If

End Function

Private Sub mnuAllText_off_Click()
   ThemeVisibled 128, False
End Sub

Private Sub mnuAllText_on_Click()
   ThemeVisibled 128, True
End Sub

Private Sub ThemeVisibled(Rep As Integer, mOn As Boolean)
   Dim a As Integer
   For a = 1 To Tv.Nodes.Count
      If Rep = 0 Then
         tcs.setVisibleThemeByThemeName Tv.Nodes(a).Text, mOn
         tvw.setVisibleThemeStatus tvw.getActiveView, Tv.Nodes(a).Text, mOn
         Tv.Nodes(a).Image = IIf(mOn, 2, 1)
      Else
         If tvw.existText(tvw.getActiveView, Tv.Nodes(a).Text) Then
            If tcs.getRepresentationTheme(Tv.Nodes(a).Text) >= 128 And Not mOn Then
               tcs.setRepresentationTheme Tv.Nodes(a).Text, tcs.getRepresentationTheme(Tv.Nodes(a).Text) - 128
               tvw.setVisibleTextStatus tvw.getActiveView, Tv.Nodes(a).Text, mOn
               
            ElseIf tcs.getRepresentationTheme(Tv.Nodes(a).Text) < 128 And mOn Then
               tcs.setRepresentationTheme Tv.Nodes(a).Text, tcs.getRepresentationTheme(Tv.Nodes(a).Text) + 128
               tvw.setVisibleTextStatus tvw.getActiveView, Tv.Nodes(a).Text, mOn
            End If
         End If
      End If
   Next
   tcs.plotView
End Sub

Private Sub mnuAllTheme_off_Click()
   ThemeVisibled 0, False
End Sub

Private Sub mnuAllTheme_on_Click()
   ThemeVisibled 0, True
End Sub

Private Sub mnuDeleteTheme_Click()
   On Error GoTo mnuDeleteTheme_Click_err
   Dim rs As New ADODB.Recordset, a As Integer
   
Dim vetor As Variant
Dim str As String
Dim ArrayTema(50) As String
Dim i, j As Integer
Dim aa As String
Dim b As String
   Dim d As String
Dim c As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim jj As String
Dim ii As String
Dim k As String
      
If TypeConn <> 4 Then
   
      If MsgBox("Tem certeza que deseja excluir o tema: " & Tv.Nodes.Item(indexTheme).Text, 36, "Atenção") = vbYes Then
                     
        tvw.removeTheme tvw.getActiveView, Tv.Nodes.Item(indexTheme).Text
        rs.Open "Select priority from te_theme where view_id in(select view_id from te_view where name='" & tvw.getActiveView & "') order by priority", conn, adOpenKeyset, adLockOptimistic, adCmdText
        While Not rs.EOF
           rs.Fields("priority") = a
           rs.Update
           
           a = a + 1
           rs.MoveNext
        Wend
  
     
        Dim tName As String
        tName = Tv.Nodes.Item(indexTheme).Text
      
        Close #3
        intTema = 0
        strCmdFiltro = ""
        Open App.Path & "\FTema.txt" For Input As #3
        Do While Not EOF(3)
             Line Input #3, str
             vetor = Split(str, ";")
             If CStr(vetor(1)) = tName Then
                 intTema = vetor(0)
                 'Exit Do
             Else
                ArrayTema(i) = str
                i = i + 1
             End If
        Loop
        Close #3
        
        Open App.Path & "\FTema.txt" For Output As #3
        Do While Not j = i
            Print #3, ArrayTema(j)
            j = j + 1
        Loop
        Close #3
       ' alterado em 21/10/2010
                
     '   If intTema <> 0 Then
      
            conn.Execute ("DELETE FROM NXGS_FILT_TEMA WHERE THEME_ID = " & intTema)
            
            
        'MsgBox "FILTRO TEMA EXCLUIDO"
       
     '  End If

        rs.Close
        Tv.Nodes.Remove indexTheme
        tcs.ResetView
        RaiseEvent onReset(tvw.getActiveView)
        tcs.plotView
        
    End If
        
  Else
  
  
  



aa = "te_theme"
b = "te_view"
c = "theme_id"
d = "generate_attribute_where"
e = "name"
f = "te_layer"
g = "view_id"
h = "layer_id"
ii = "user_name"
jj = "priority"
   


     
        If MsgBox("Tem certeza que deseja excluir o tema: " & Tv.Nodes.Item(indexTheme).Text, 36, "Atenção") = vbYes Then
                       'tvw.Start
         tvw.removeTheme tvw.getActiveView, Tv.Nodes.Item(indexTheme).Text
    
         rs.Open "Select " + """" + jj + """" + " from " + """" + aa + """" + " where " + """" + g + """" + " in(select " + """" + g + """" + " from " + """" + b + """" + " where " + """" + e + """" + "='" & tvw.getActiveView & "') order by " + """" + jj + """" + "", conexao, adOpenDynamic, adLockOptimistic
        End If
        
      
      
        While Not rs.EOF
           rs.Fields("priority") = a
           rs.Update
           
           a = a + 1
           rs.MoveNext
        Wend
 
     
      
        tName = Tv.Nodes.Item(indexTheme).Text
      
        Close #3
        intTema = 0
        strCmdFiltro = ""
        Open App.Path & "\FTema.txt" For Input As #3
        Do While Not EOF(3)
             Line Input #3, str
             vetor = Split(str, ";")
             If CStr(vetor(1)) = tName Then
                 intTema = vetor(0)
                 'Exit Do
             Else
                ArrayTema(i) = str
                i = i + 1
             End If
        Loop
        Close #3
        
        Open App.Path & "\FTema.txt" For Output As #3
        Do While Not j = i
            Print #3, ArrayTema(j)
            j = j + 1
        Loop
        Close #3
       ' alterado em 21/10/2010
                
        If intTema <> 0 Then
        
             aa = "NXGS_FILT_TEMA"
      b = "THEME_ID"
      
        
           ' conn.Execute ("DELETE FROM " + """" + aa + """" + " WHERE " + """" + c + """" + " = '" & intTema & "' ")
             conn.Execute ("DELETE FROM " + """" + aa + """" + " WHERE " + """" + b + """" + " = '" & intTema & "' ")
        'MsgBox "FILTRO TEMA EXCLUIDO"
        
       End If

        rs.Close
        Tv.Nodes.Remove indexTheme
        tcs.ResetView
        RaiseEvent onReset(tvw.getActiveView)
        tcs.plotView
        
    End If
  
  
  
  
  
  
  
         
        
 
 
   Set rs = Nothing
   Exit Sub
   
mnuDeleteTheme_Click_err:
 
  ' MsgBox Err.Description & vbCrLf & "Vista: " & tvw.getactiveview & " Tema: " & Tv.Nodes.Item(indexTheme).Text
 
End Sub

Private Sub mnuLines_Click()

   SetRepTheme tcs.getRepresentationTheme(Tv.Nodes.Item(indexTheme).Text), 2

End Sub


Private Sub mnuNewTheme_Click()
   On Error GoTo mnunewtheme_error
   
   Dim mLayerName As String, mThemeName As String
   
   If FrmLayerTheme.Init(tvw, mLayerName, mThemeName) Then
      tvw.addTheme mLayerName, tvw.getActiveView, mThemeName
      frmTheme.Init tvw, mThemeName, mLayerName
      tcs.ResetView
      LoadTheme
      tcs.plotView
   End If
   
   Exit Sub

mnunewtheme_error:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuPoints_Click()
   SetRepTheme tcs.getRepresentationTheme(Tv.Nodes.Item(indexTheme).Text), 4
End Sub

Private Sub mnuPolygons_Click()
   SetRepTheme tcs.getRepresentationTheme(Tv.Nodes.Item(indexTheme).Text), 1
End Sub

Private Sub mnuRenameTheme_Click()
   Tv.Nodes.Item(indexTheme).Selected = True
   Tv.StartLabelEdit
End Sub

Private Sub mnuTexts_Click()
   SetRepTheme tcs.getRepresentationTheme(Tv.Nodes.Item(indexTheme).Text), 128
End Sub

Private Sub mnuThemeOn_Click()
    
    If Tv.Nodes.Count > 0 Then
        Select Case Tv.Nodes.Item(indexTheme).Image
            Case 1 'ligado sem snap
                Tv.Nodes.Item(indexTheme).Image = mOn
            Case Else
                Tv.Nodes.Item(indexTheme).Image = moff
        End Select
        tcs.setVisibleThemeByThemeName Tv.Nodes.Item(indexTheme).Text, IIf(Tv.Nodes.Item(indexTheme).Image = 2, True, False)
        tvw.setVisibleThemeStatus tvw.getActiveView, Tv.Nodes.Item(indexTheme).Text, IIf(Tv.Nodes.Item(indexTheme).Image = 2, True, False)
        tcs.plotView
    End If
    
End Sub

Private Sub mnuThemeProperties_Click()
   On Error GoTo mnuThemeProperties_Click
   LoadThemeProperties
   Exit Sub
mnuThemeProperties_Click:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuView_Click()
   Dim Frm As New frmViews
   If TypeConn <> 4 Then
   If Frm.Init(tvw, tcs, manager, database) Then
      tcs.ResetView
      RaiseEvent onReset(tvw.getActiveView)
      LoadTheme
      tcs.plotView
   End If
   Else
   'manager.Start
   
   If Frm.Init(TeViewDatabase1, tcs, manager, database) Then
      tcs.ResetView
      RaiseEvent onReset(tvw.getActiveView)
      LoadTheme
      tcs.plotView
   End If
   
   End If
   
   
   Set Frm = Nothing
End Sub


Private Sub Tv_AfterLabelEdit(Cancel As Integer, NewString As String)
   On Error GoTo Tv_AfterLabelEdit_err
   
   
   
   If Not NewString = "" Then
   
      If tvw.renameTheme(tvw.getActiveView, Tv.SelectedItem.Text, NewString) = False Then
         
         Cancel = 1
         
      End If
   
   Else
      MsgBox "O nome do tema não pode ser vazio", vbExclamation
      Cancel = 1
      Exit Sub
   End If
   
      tcs.ResetView
      LoadTheme
      tcs.plotView
      
   
   Exit Sub

Tv_AfterLabelEdit_err:
   Cancel = 1
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub Tv_DblClick()
   mnuThemeOn_Click
   'If Not Tv.SelectedItem Is Nothing Then
    'If frmTheme.Init(tvw, Tv.SelectedItem.Text, Tv.SelectedItem.Tag) Then
     'tcs.ResetView
      'LoadTheme
       'tcs.plotView
     'End If
  'End If
End Sub
Private Sub GetRepresentation(ActiveView As String, mtheme As String)
   
   With tvw
      If .existPolygon(ActiveView, mtheme) Then
         mnuPolygons.Enabled = True
      Else
         mnuPolygons.Enabled = False
      End If
      If .existLine(ActiveView, mtheme) Then
         mnuLines.Enabled = True
      Else
         mnuLines.Enabled = False
      End If
      
      If .existPoint(ActiveView, mtheme) Then
         mnuPoints.Enabled = True
      Else
         mnuPoints.Enabled = False
      End If
      If .existText(ActiveView, mtheme) Then
         mnuTexts.Enabled = True
      Else
         mnuTexts.Enabled = False
      End If
   End With

End Sub


Private Sub Tv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo Tv_MouseDown_err
   If Not tvw Is Nothing And Not mConn Is Nothing Then
      If Not Tv.HitTest(x, y) Is Nothing Then
         ThemeName = Tv.HitTest(x, y).Text
         indexTheme = Tv.HitTest(x, y).Index
         If Button = 2 Then
            GetRepresentation tvw.getActiveView, ThemeName
            SetMnuRepTheme tvw.getThemeRepresentation(tvw.getActiveView, ThemeName)
            mnuThemeProperties.Caption = "Propriedades: " & Tv.HitTest(x, y).Text
            mnuMenuEnabled True
            PopupMenu mnuTheme
         End If
      ElseIf Tv.HitTest(x, y) Is Nothing And Button = 2 Then
         mnuMenuEnabled False
         PopupMenu mnuTheme
      End If
   End If
   Exit Sub
Tv_MouseDown_err:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuMenuEnabled(mEnabled As Boolean)
   mnuTheme.Enabled = mEnabled
   mnuThemeProperties.Enabled = mEnabled
   mnuNewTheme.Enabled = IIf(tvw.getActiveView = "", mEnabled, True)
   mnuDeleteTheme.Enabled = mEnabled
   mnuRenameTheme.Enabled = mEnabled
   mnuThemeOn.Enabled = mEnabled
   mnuAll.Enabled = mEnabled
   mnuSnap.Enabled = mEnabled
   mnuRepresentation.Enabled = mEnabled
End Sub

Private Sub Tv_NodeClick(ByVal Node As MSComctlLib.Node)
   
   
   
   
   
   Dim a As Integer
   For a = 1 To Tv.Nodes.Count 'Limpa os Seleciocionados e os snepados
      If Tv.Nodes.Item(a).Image = mOnSet Or Tv.Nodes.Item(a).Image = mOnSnapLock Then
         Tv.Nodes.Item(a).Image = mOn
      End If
   Next
   For a = 1 To Tv.Nodes.Count 'Alterar a imagem de todos os temas do mesmo plano p/ Setado(se ativo)
      If Tv.Nodes.Item(a).Image <> moff Then
         If Node.Image <> moff And Node.Tag = Tv.Nodes.Item(a).Tag Then
            tcs.setActiveGeometry 0 ' ativa todas as representações
            mnuRepEnable(1).Checked = True
            mnuRepEnable(2).Checked = True
            mnuRepEnable(4).Checked = True
            mnuRepEnable(128).Checked = True
            
            tcs.setCurrentLayer Node.Tag
            Dim layerAtual As String
            layerAtual = tcs.getCurrentLayer
           ' frmTheme.Layer (layerAtual)
            tcs.Normal
            tcs.Select
            RaiseEvent onReset(tvw.getActiveView)
            Tv.Nodes.Item(a).Image = mOnSet
         End If
      End If
   Next
  
End Sub

Private Sub mnuSnap_Click()
   If Tv.Nodes.Item(indexTheme).Tag <> tcs.getCurrentLayer() Then
      Select Case Tv.Nodes.Item(indexTheme).Image
         Case mOn 'ligado sem snap
            LoadImageSnap Tv.Nodes.Item(indexTheme).Tag, mOnSnap
            tcs.addLayerToSnap Tv.Nodes.Item(indexTheme).Tag
         Case mOnSnap
            LoadImageSnap Tv.Nodes.Item(indexTheme).Tag, mOn
            tcs.removeLayerToSnap Tv.Nodes.Item(indexTheme).Tag
      End Select
   End If
End Sub
Public Sub LoadImageSnap(mLayerName As String, Img As mImg)
   Dim a As Integer
   For a = 1 To Tv.Nodes.Count
      If Tv.Nodes.Item(a).Tag = mLayerName Then Tv.Nodes.Item(a).Image = Img
   Next
End Sub

Private Function GetVisibledTheme(mtheme As String) As Integer


 If TypeConn <> 4 Then
   If tvw.visibleThemeStatus(tvw.getActiveView, mtheme) Then
      GetVisibledTheme = 2
   Else
      GetVisibledTheme = 1
   End If
   Else

'manager.Start
'manager.setActiveView TeViewDatabase1.getActiveView
'manager.saveAsLastView TeViewDatabase1.getActiveView

'MsgBox manager.visibleThemeStatus(mtheme)
'FrmLayerTheme.Temas (mtheme)
'manager.Start
'If FrmLayerTheme.Temas2 = "0" Then


 If database.visibleThemeStatus(database.getActiveView, mtheme) Then
      GetVisibledTheme = 2
   Else
      GetVisibledTheme = 1
   End If

'Else

  ' If FrmLayerTheme.Temas2 = "1" Then
  '    GetVisibledTheme = 2
  ' Else
  '    GetVisibledTheme = 1
  ' End If
   
 ' End If
   End If
End Function

Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
  
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 
 Dim retlen As String
 Dim Ret As String
 
 Ret = String$(255, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
 Ret = Left$(Ret, retlen)
 ReadINI = Ret

End Function








Private Function GetSnapLayer(mLayer As String) As Integer
   Dim a As Integer, rtn_layer As String
   For a = 0 To tcs.getLayersToSnapCount() - 1
      tcs.getLayerToSnap a, rtn_layer
      If mLayer = rtn_layer Then GetSnapLayer = 4
   Next
End Function

Private Sub SetRepTheme(Repvisibled As Integer, Rep As Integer)
      Select Case Rep
        Case 1
          Select Case Repvisibled
            Case 1, 5, 7, 129, 131, 135 'Poligonus
               tcs.setRepresentationTheme ThemeName, Repvisibled - 1
               tvw.setVisiblePolygonStatus tvw.getActiveView, ThemeName, False
            Case Else
               tcs.setRepresentationTheme ThemeName, Repvisibled + 1
               tvw.setVisiblePolygonStatus tvw.getActiveView, ThemeName, True
          End Select
        Case 2
          Select Case Repvisibled
            Case 2, 3, 6, 7, 130, 134
               tcs.setRepresentationTheme ThemeName, Repvisibled - 2
               tvw.setVisibleLineStatus tvw.getActiveView, ThemeName, False
            Case Else
               tcs.setRepresentationTheme ThemeName, Repvisibled + 2
               tvw.setVisibleLineStatus tvw.getActiveView, ThemeName, True
          End Select
        Case 4
          Select Case Repvisibled
            Case 4, 5, 6, 7, 132, 134, 135
               tcs.setRepresentationTheme ThemeName, Repvisibled - 4
               tvw.setVisiblePointStatus tvw.getActiveView, ThemeName, False
            Case Else
               tcs.setRepresentationTheme ThemeName, Repvisibled + 4
               tvw.setVisiblePointStatus tvw.getActiveView, ThemeName, False
          End Select
        Case 128
          Select Case Repvisibled
            Case Is >= 128
               tcs.setRepresentationTheme ThemeName, Repvisibled - 128
               tvw.setVisibleTextStatus tvw.getActiveView, ThemeName, False
            Case Else
               tcs.setRepresentationTheme ThemeName, Repvisibled + 128
               tvw.setVisibleTextStatus tvw.getActiveView, ThemeName, True
          End Select
      End Select
      tcs.plotView
End Sub
Private Sub SetMnuRepTheme(Repvisibled As Integer)
      Select Case Repvisibled
        Case 1, 5, 7, 129, 131, 135 'Poligonus
           mnuPolygons.Checked = True
        Case Else
           mnuPolygons.Checked = False
      End Select
      Select Case Repvisibled
        Case 2, 3, 6, 7, 130, 134
           mnuLines.Checked = True
        Case Else
           mnuLines.Checked = False
      End Select
      Select Case Repvisibled
        Case 4, 5, 6, 7, 132, 134, 135
           mnuPoints.Checked = True
        Case Else
           mnuPoints.Checked = False
      End Select
      Select Case Repvisibled
        Case Is >= 128
           mnuTexts.Checked = True
        Case Else
           mnuTexts.Checked = False
      End Select
End Sub

Private Sub mnuRepEnable_Click(Index As Integer)
   If Index <> 512 Then
      mnuRepEnable(1).Checked = False: mnuRepEnable(2).Checked = False: mnuRepEnable(4).Checked = False:  mnuRepEnable(128).Checked = False
      mnuRepEnable(Index).Checked = True
      tcs.setActiveGeometry Index
   Else
      mnuRepEnable(1).Checked = True: mnuRepEnable(2).Checked = True: mnuRepEnable(4).Checked = True:  mnuRepEnable(128).Checked = True
      tcs.setActiveGeometry 0
   End If
End Sub

Private Sub Tv_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo Tv_OLEDragDrop_Err
   Dim a, mStep As Integer
   If Tv.HitTest(x, y) Is Nothing Then Exit Sub
   mStep = IIf(Tv.Nodes(indexTheme).Index > Tv.HitTest(x, y).Index, -1, 1)
   For a = Tv.Nodes(indexTheme).Index To (Tv.HitTest(x, y).Index - mStep) Step mStep
      If mStep = -1 Then
         tvw.moveUp tvw.getActiveView, Tv.Nodes(indexTheme).Text
      Else
         tvw.moveDown tvw.getActiveView, Tv.Nodes(indexTheme).Text
      End If
   Next
   

   LoadTheme
   

   
   tcs.ResetView
   tcs.plotView
   Exit Sub
Tv_OLEDragDrop_Err:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Public Sub ResetView()
   'Set tvw = Nothing
  ' Set tcs = Nothing
      
  ' tcs.ResetView
  ' tcs.plotView
   Tv.Nodes.Clear
   
End Sub


Public Function FunDecripta(ByVal strDecripta As String) As String


    Dim IntTam As Integer
    Dim i As Integer
    Dim letra As String
    IntTam = Len(strDecripta)
    Dim nStr As String
    nStr = ""

    'desconsidera os os numeros de HH-MM-SS
    strDecripta = Mid(strDecripta, 6, 5) & Mid(strDecripta, 16, 5) & Mid(strDecripta, 26, 5) & _
                  Mid(strDecripta, 36, 5) & Mid(strDecripta, 46, 5) & Mid(strDecripta, 56, 200)

    i = 1
    Do While Not i = IntTam - 29
        letra = Mid(strDecripta, i, 5)
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


