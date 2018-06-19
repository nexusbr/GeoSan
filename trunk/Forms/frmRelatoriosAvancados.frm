VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRelatoriosAvancados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório por Seleção de Polígono"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   180
      TabIndex        =   0
      Top             =   5790
      Width           =   6225
   End
   Begin VB.Frame Frame1 
      Height          =   5250
      Left            =   195
      TabIndex        =   2
      Top             =   165
      Width           =   7335
      Begin VB.OptionButton optRelRamalPadrao 
         Caption         =   "Relatorio Padrão"
         Height          =   2415
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   7005
      End
      Begin VB.OptionButton Option4 
         Caption         =   "RAMAL; CONSUMO MÉDIO; ECONOMIAS; TIPO; HIDROMETRADO"
         Height          =   375
         Left            =   255
         TabIndex        =   16
         Top             =   1995
         Width           =   6930
      End
      Begin VB.OptionButton Option3 
         Caption         =   "SELEÇÃO > >"
         Height          =   540
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   5415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "MATERIAL; COMPRIMENTO; DIAMETRO INTERNO"
         Height          =   390
         Left            =   240
         TabIndex        =   4
         Top             =   855
         Width           =   6915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ID; MATERIAL; COMPRIMENTO; DIAMETRO INTERNO"
         Height          =   420
         Left            =   240
         TabIndex        =   3
         Top             =   510
         Value           =   -1  'True
         Width           =   6930
      End
      Begin VB.Label Label3 
         Caption         =   "Ramais de Água"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   15
         Top             =   1740
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Redes de Água"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   14
         Top             =   270
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdGerarRelatorio 
      Caption         =   "Procurar "
      Height          =   390
      Left            =   6495
      TabIndex        =   1
      Top             =   5775
      Width           =   1020
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecione os Campos"
      Height          =   3600
      Left            =   165
      TabIndex        =   6
      Top             =   6615
      Width           =   7350
      Begin VB.CommandButton cmdUP 
         Height          =   435
         Left            =   6525
         Picture         =   "frmRelatoriosAvancados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   450
         Width           =   645
      End
      Begin VB.CommandButton cmdDown 
         Height          =   435
         Left            =   6525
         Picture         =   "frmRelatoriosAvancados.frx":0552
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   945
         Width           =   645
      End
      Begin VB.ListBox List1 
         Height          =   2985
         ItemData        =   "frmRelatoriosAvancados.frx":0AA4
         Left            =   165
         List            =   "frmRelatoriosAvancados.frx":0AA6
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   375
         Width           =   2730
      End
      Begin VB.CommandButton cmdAddCamppo 
         Height          =   435
         Left            =   3000
         Picture         =   "frmRelatoriosAvancados.frx":0AA8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   450
         Width           =   645
      End
      Begin VB.CommandButton cmdRemCampo 
         Height          =   435
         Left            =   3000
         Picture         =   "frmRelatoriosAvancados.frx":111A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   945
         Width           =   645
      End
      Begin VB.ListBox List2 
         Height          =   2985
         ItemData        =   "frmRelatoriosAvancados.frx":178C
         Left            =   3705
         List            =   "frmRelatoriosAvancados.frx":178E
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   390
         Width           =   2730
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Caminho"
      Height          =   255
      Left            =   195
      TabIndex        =   5
      Top             =   5535
      Width           =   2835
   End
End
Attribute VB_Name = "frmRelatoriosAvancados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String
Dim COMP As String
Dim rs As ADODB.Recordset
Dim qtdCampos As Integer

Dim rsBig As New ADODB.Recordset
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim ii As String
Dim j As String
Dim k As String

Dim selecionaArquivo As New CArquivo                                                    'para obter o nome e diretório onde o arquivo será salvo
Dim nomeArquivo As String                                                               'nome completo do arquivo com o drive e diretório no qual será salvo
Dim diretorioMyDocuments As String                                                      'diretório meus documentos inicial
Dim filelocation As String                                                              'nome completo do arquivo onde será salvo o relatório
' Gera o relatório do tipo 1
'
'
'
Private Function OPT1() As Boolean
   
On Error GoTo Trata_Erro

   'a. Object_ID do trecho de rede.
   'b. Número do setor de abastecimento.
   'c. Comprimento da rede.
   'd. Material.
   'e. Diâmetro.


a = "OBJECT_ID_"
b = "WATERLINES"
c = "LENGTHCALCULATED"
d = "X_MATERIAL"
e = "MATERIALNAME"
f = "INTERNALDIAMETER"
g = "POLIGONO_SELECAO"
h = "USUARIO"
i = "TIPO"
ii = "MATERIAL"
j = "MATERIALID"
   If frmCanvas.TipoConexao = 1 Then 'SQL
      strsql = "SELECT LEN(W.OBJECT_ID_) AS TAM, W.OBJECT_ID_, LENGTHCALCULATED, M.MATERIALNAME, W.INTERNALDIAMETER FROM WATERLINES W INNER JOIN X_MATERIAL M ON W.MATERIAL = M.MATERIALID WHERE w.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 1) ORDER BY TAM, OBJECT_ID_"
   
   ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
      strsql = "SELECT OBJECT_ID_, LENGTHCALCULATED, MATERIALNAME,INTERNALDIAMETER FROM WATERLINES A INNER JOIN X_MATERIAL M ON A.MATERIAL = M.MATERIALID AND EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE A.LINE_ID = P.OBJECT_ID_ AND P.USUARIO = '" & strUser & "' AND P.TIPO = 1) ORDER BY OBJECT_ID_"
   ElseIf frmCanvas.TipoConexao = 4 Then
   'SELECT length("WATERLINES"."OBJECT_ID_") AS TAM, "WATERLINES"."OBJECT_ID_", "LENGTHCALCULATED", "X_MATERIAL"."MATERIALNAME", "WATERLINES"."INTERNALDIAMETER" FROM "WATERLINES"
   'INNER JOIN "X_MATERIAL" ON "WATERLINES"."MATERIAL" = "X_MATERIAL"."MATERIALID" WHERE "WATERLINES"."OBJECT_ID_" IN (SELECT "OBJECT_ID_" FROM "POLIGONO_SELECAO" WHERE "USUARIO" = 'Administrador' AND "TIPO" = '1') ORDER BY TAM, "OBJECT_ID_"
   strsql = " SELECT length(" + """" + b + """" + "." + """" + a + """" + ") AS " + """" + "TAM" + """" + ", " + """" + b + """" + "." + """" + a + """" + "," + """" + c + """" + "," + """" + d + """" + "." + """" + e + """" + "," + """" + b + """" + "." + """" + f + """" + " FROM " + """" + b + """" + "INNER JOIN " + """" + d + """" + "ON " + """" + b + """" + "." + """" + ii + """" + " = " + """" + d + """" + "." + """" + j + """" + " WHERE " + """" + b + """" + "." + """" + a + """" + " IN(SELECT " + """" + a + """" + " FROM " + """" + g + """" + "WHERE " + """" + h + """" + " ='" & strUser & "' AND " + """" + i + """" + " = '1') ORDER BY " + """" + "TAM" + """" + ", " + """" + a + """" + ""
   End If
   
   Set rs = New ADODB.Recordset
      rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
   
   Open Me.Text1.Text For Output As #1
   Print #1, "OBJECT_ID;MATERIAL;COMPRIMENTO;DIAMETRO INTERNO"
   
   Do While Not rs.EOF = True
   
      COMP = Replace(rs!LENGTHCALCULATED, ".", ",")
   
      Print #1, rs!Object_id_ & ";" & rs!MATERIALNAME & ";" & COMP & ";" & rs!INTERNALDIAMETER
   
      rs.MoveNext
   Loop
   Close #1
     MsgBox "Relatório gerado com sucesso!", vbInformation, ""
   rs.Close
   Set rs = Nothing
   
   OPT1 = True

Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
ElseIf Err.Number = 52 Then
   MsgBox "Caminho de arquivo incorreto!", vbExclamation, ""
   Err.Clear
Else
   MsgBox Err.Number & " - " & Err.Description
End If

End Function
' Gera o relatório do tipo 2
'
'
'
Private Function OPT2() As Boolean

a = "OBJECT_ID_"
b = "WATERLINES"
c = "LENGTHCALCULATED"
d = "X_MATERIAL"
e = "MATERIALNAME"
f = "INTERNALDIAMETER"
g = "POLIGONO_SELECAO"
h = "USUARIO"
i = "TIPO"
ii = "MATERIAL"
j = "LINE_ID"
k = "MATERIALID"
On Error GoTo Trata_Erro
   
   If frmCanvas.TipoConexao = 1 Then ' SQL
      strsql = "SELECT m.materialname, w.internaldiameter, sum(w.lengthcalculated) as metros from waterlines w inner join x_material m on w.material = m.materialid WHERE w.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 1) group by m.materialname,w.internaldiameter order by m.materialname,w.internaldiameter"
   
   ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
   
      strsql = "SELECT M.MATERIALNAME, W.INTERNALDIAMETER, SUM(W.LENGTHCALCULATED) AS " + """" + "METROS" + """" + " FROM WATERLINES W INNER JOIN X_MATERIAL M ON W.MATERIAL = M.MATERIALID AND EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE W.LINE_ID = P.OBJECT_ID_ AND P.USUARIO = '" & strUser & "' AND P.TIPO = 1) GROUP BY M.MATERIALNAME,W.INTERNALDIAMETER ORDER BY M.MATERIALNAME,W.INTERNALDIAMETER"
   
   ElseIf frmCanvas.TipoConexao = 4 Then
  ' strsql = " SELECT "X_MATERIAL"."MATERIALNAME", "WATERLINES"."INTERNALDIAMETER", SUM("WATERLINES"."LENGTHCALCULATED") AS METROS FROM "WATERLINES" INNER JOIN "X_MATERIAL"  ON
  ' "WATERLINES"."MATERIAL" = "X_MATERIAL"."MATERIALID" AND EXISTS (SELECT '1' FROM "POLIGONO_SELECAO"  WHERE ("WATERLINES"."LINE_ID" = (CAST ("POLIGONO_SELECAO"."OBJECT_ID_" AS Integer))) AND "POLIGONO_SELECAO"."USUARIO" = 'Administrador'
' AND "POLIGONO_SELECAO"."TIPO" = '1') GROUP BY "X_MATERIAL"."MATERIALNAME","WATERLINES"."INTERNALDIAMETER" ORDER BY "X_MATERIAL"."MATERIALNAME","WATERLINES"."INTERNALDIAMETER""
    
  '  strsql = " SELECT " + d + "." + e + "," + b + "." + f + ", SUM(" + b + "." + j + "=(CAST(" + g + "." + a + " AS INTEGER))) AND " + g + "." + h + "where= '" & strUser & "' AND " + g + "." + i + "='1') GROUP BY " + d + "." + e + "," + b + "." + f + " ORDER BY " + d + "." + e + "," + b + "." + f + ""
    strsql = " SELECT " + """" + d + """" + "." + """" + e + """" + "," + """" + b + """" + "." + """" + f + """" + ", SUM(" + """" + b + """" + "." + """" + c + """" + ") AS " + """" + "METROS" + """" + " from " + """" + b + """" + "inner join " + """" + d + """" + " on " + """" + b + """" + "." + """" + ii + """" + "=" + """" + d + """" + "." + """" + k + """" + " where " + """" + b + """" + "." + """" + a + """" + " IN(SELECT " + """" + a + """" + " from " + """" + g + """" + " where " + """" + h + """" + " = '" & strUser & "' AND " + """" + i + """" + "='1') GROUP BY " + """" + d + """" + "." + """" + e + """" + "," + """" + b + """" + "." + """" + f + """" + " ORDER BY " + """" + d + """" + "." + """" + e + """" + "," + """" + b + """" + "." + """" + f + """" + ""

   
   
   End If
   
   Set rs = New ADODB.Recordset
      rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
   
   Open Me.Text1.Text For Output As #1
   Print #1, "MATERIAL;DIAMETRO INTERNO;COMPRIMENTO"
   
   Do While Not rs.EOF = True
   
      COMP = Replace(rs!METROS, ".", ",")
   
      Print #1, rs!MATERIALNAME & ";" & rs!INTERNALDIAMETER & ";" & COMP
      
      rs.MoveNext
   Loop
        MsgBox "Relatório gerado com sucesso!", vbInformation, ""
   Close #1
   rs.Close
   Set rs = Nothing
   
   OPT2 = True

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
ElseIf Err.Number = 52 Then
   MsgBox "Caminho de arquivo incorreto!", vbExclamation, ""
   Err.Clear
Else
   
   PrintErro CStr(Me.Name), "Private Function OPT2()", CStr(Err.Number), CStr(Err.Description), True

End If

End Function
' Gera o relatório do tipo 4
'
'
'
Private Function OPT4() As Boolean

On Error GoTo Trata_Erro
   Dim calcConsumo As New CConsumo                              'para calcular as conversões de l/s e m3/mês
   Dim consumoM3 As String                                      'consumo em m3/mês
   Dim strNro As String
   Dim strRamal As String
   Dim strTipo As String
   Dim strConsumo As String
   Dim strHidro As String


   Dim ha As String
   Dim he As String
   Dim hi As String
   Dim ho As String
   Dim hu As String
   Dim hb As String
   Dim hc As String
   Dim hd As String
   
ha = "OBJECT_ID_"
he = "CONSUMO_LPS"
hi = "ECONOMIAS"
ho = "TIPO"
hu = "HIDROMETRADO"
hb = "RAMAIS_AGUA_LIGACAO"
hc = "POLIGONO_SELECAO"
hd = "USUARIO"
          

   'strsql = "SELECT RAL.OBJECT_ID_ AS RAMAL,RAL.NRO_LIGACAO AS NROLIGA ,RAL.TIPO AS TIPO, RAL.CONSUMO_LPS AS CONSUMO, RA.OBJECT_ID_TRECHO AS REDE FROM RAMAIS_AGUA_LIGACAO AS RAL INNER JOIN RAMAIS_AGUA AS RA ON RAL.OBJECT_ID_ = RA.OBJECT_ID_ WHERE RA.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2)"
   
   If frmCanvas.TipoConexao = 1 Then
      strsql = "SELECT OBJECT_ID_ AS RAMAL, SUM(CONSUMO_LPS) AS CONSUMO, SUM(ECONOMIAS) AS ECONOMIAS, TIPO, HIDROMETRADO FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2) GROUP BY OBJECT_ID_,TIPO,HIDROMETRADO"
   
   ElseIf frmCanvas.TipoConexao = 2 Then
      strsql = "SELECT OBJECT_ID_ AS " + """" + "RAMAL" + """" + ", SUM(CONSUMO_LPS) AS " + """" + "CONSUMO" + """" + ", SUM(ECONOMIAS) AS " + """" + "ECONOMIAS" + """" + ", TIPO, HIDROMETRADO FROM RAMAIS_AGUA_LIGACAO R WHERE EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE R.OBJECT_ID_ = P.OBJECT_ID_ AND P.USUARIO = '" & strUser & "' AND P.TIPO = 2) GROUP BY OBJECT_ID_, TIPO, HIDROMETRADO"
ElseIf frmCanvas.TipoConexao = 4 Then
 strsql = "SELECT " + """" + ha + """" + " AS " + """" + "RAMAL" + """" + ", SUM(" + """" + he + """" + ") AS " + """" + "CONSUMO" + """" + ", SUM(" + """" + hi + """" + ") AS " + """" + "ECONOMIAS" + """" + ", " + """" + ho + """" + ", " + """" + hu + """" + " FROM " + """" + hb + """" + " WHERE " + """" + ha + """" + " IN (SELECT " + """" + ha + """" + " FROM " + """" + hc + """" + " WHERE " + """" + hd + """" + " = '" & strUser & "' AND " + """" + ho + """" + " = '2') GROUP BY " + """" + ha + """" + "," + """" + ho + """" + "," + """" + hu + """" + ""
   
     

   
   End If
   
   Open Me.Text1.Text For Output As #1
   Print #1, "NÚMERO DO RAMAL;CONSUMO MÉDIO (l/s);CONSUMO MÉDIO (m3/mês);NÚMERO DE ECONOMIAS;TIPO;HIDROMETRADO"
   
   Set rs = New ADODB.Recordset
    rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
      
   If rs.EOF = False Then
      
      Do While Not rs.EOF = True
         
         If rs!ramal <> "" Then
            strRamal = Trim(rs!ramal)
         Else
            strRamal = ""
         End If
         
         If rs!consumo <> "" Then
            strConsumo = Replace(rs!consumo, ".", ",")
         Else
            strConsumo = 0
         End If
         
         If rs!ECONOMIAS <> "" Then
            strNro = Trim(rs!ECONOMIAS)
         Else
            strNro = ""
         End If
         
         If rs!tipo <> "" Then
            strTipo = Trim(rs!tipo)
         Else
            strTipo = ""
         End If
         
         If rs!HIDROMETRADO <> "" Then
            strHidro = Trim(rs!HIDROMETRADO)
         Else
            strHidro = ""
         End If
         consumoM3 = CStr(calcConsumo.lps2m3mes(CDbl(strConsumo)))
         consumoM3 = Round(consumoM3, 2)
         consumoM3 = Replace(consumoM3, ".", ",")
         Print #1, strRamal & ";" & strConsumo & ";" & consumoM3 & ";" & strNro & ";" & strTipo & ";" & strHidro
     
         rs.MoveNext
      Loop
   End If
   
   Close #1
   rs.Close
   Set rs = Nothing
   
  MsgBox "Relatório gerado com sucesso!", vbInformation, ""
   OPT4 = True

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Or Err.Number = 55 Then
   Resume Next
ElseIf Err.Number = 52 Then
   MsgBox "Caminho de arquivo incorreto!", vbExclamation, ""
   Err.Clear
Else
      PrintErro CStr(Me.Name), "OPT4", CStr(Err.Number), CStr(Err.Description), True
End If

End Function
' Gera o relatório do tipo 3 - Seleção
'
'
'
Private Function OPT3() As Boolean
    On Error GoTo Trata_Erro:
    Dim strCabecalho As String
    Dim blnInnerJoinMaterial As Boolean
    Dim blnInnerJoinTipo As Boolean
    Dim strPrint As String
    Dim strCampo As String
    Dim i As Integer
    
    strsql = ""
    Screen.MousePointer = vbHourglass                                                   'mostra a ampulheta para o usuário
    For i = 0 To List2.ListCount - 1
        strCampo = ""
        If List2.list(i) = "DATA DE DESENHO" Then
            strCampo = "DATA_LOG"
        ElseIf List2.list(i) = "DATALOG" Then
            'List1.AddItem
        ElseIf List2.list(i) = "DATA DE INSTALAÇÃO" Then
            strCampo = "DATEINSTALLATION"
        ElseIf List2.list(i) = "DISTÂNCIA DA DIVISA" Then
            strCampo = "DIVIDEDDISTANCE"
        ElseIf List2.list(i) = "DIÂMETRO EXTERNO" Then
            strCampo = "EXTERNALDIAMETER"
        ElseIf List2.list(i) = "CÓD COMPONENTE FINAL" Then
            strCampo = "FINALCOMPONENT"
        ElseIf List2.list(i) = "COTA TERRENO FINAL" Then
            strCampo = "FINALGROUNDHEIGHT"
        ElseIf List2.list(i) = "PROFUNDIDADE FINAL" Then
            strCampo = "FINALTUBEDEEPNESS"
        ElseIf List2.list(i) = "TIPO DE REDE" Then
            strCampo = "ID_TYPE"
        ElseIf List2.list(i) = "VALIDADE" Then
            strCampo = "INFORMATIONVALIDITY"
        ElseIf List2.list(i) = "CÓD COMPONENTE INICAL" Then
            strCampo = "INITIALCOMPONENT"
        ElseIf List2.list(i) = "COTA TERRENO INICIAL" Then
            strCampo = "INITIALGROUNDHEIGHT"
        ElseIf List2.list(i) = "PROFUNDIDADE INICIAL" Then
            strCampo = "INITIALTUBEDEEPNESS"
        ElseIf List2.list(i) = "DIÂMETRO INTERNO" Then
            strCampo = "INTERNALDIAMETER"
        ElseIf List2.list(i) = "COMPRIM. DIGITADO" Then
            strCampo = "LENGTH"
        ElseIf List2.list(i) = "COMPRIM. CALCULADO" Then
            strCampo = "LENGTHCALCULATED"
        ElseIf List2.list(i) = "LINE_ID" Then
            'List1.AddItem ""
        ElseIf List2.list(i) = "LOCALIZAÇÃO" Then
            strCampo = "LOCATION"
        ElseIf List2.list(i) = "FABRICANTE" Then
            strCampo = "MANUFACTURER"
        ElseIf List2.list(i) = "MATERIAL NOME" Then
            strCampo = "MATERIAL"
        ElseIf List2.list(i) = "ID REDE" Then
            strCampo = "OBJECT_ID_"
        ElseIf List2.list(i) = "RUGOSIDADE" Then
            strCampo = "ROUGHNESS"
        ElseIf List2.list(i) = "CÓD SETOR" Then
            strCampo = "SECTOR"
        ElseIf List2.list(i) = "LADO DA RUA" Then
            strCampo = "SIDESTREET"
        ElseIf List2.list(i) = "ESTADO" Then
            strCampo = "STATE"
        ElseIf List2.list(i) = "FORNECEDOR" Then
            strCampo = "SUPPLIER"
        ElseIf List2.list(i) = "DENSIDADE" Then
            strCampo = "THICKNESS"
        ElseIf List2.list(i) = "TROUBLE" Then
            'List1.AddItem ""
        ElseIf List2.list(i) = "USUARIO CADASTRO" Then
            strCampo = "USUARIO_LOG"
        End If
        If strsql = "" Then
            If frmCanvas.TipoConexao <> 4 Then
                If strCampo = "MATERIAL" Then
                    strsql = "X_MATERIAL.MATERIALNAME"
                    blnInnerJoinMaterial = True
                ElseIf strCampo = "ID_TYPE" Then
                    strsql = "WATERLINESTYPES.DESCRIPTION_"
                    blnInnerJoinTipo = True
                Else
                    strsql = "WATERLINES." & strCampo
                End If
            Else
                If strCampo = "MATERIAL" Then
                    strsql = """" + "X_MATERIAL" + """" + "." + """" + "MATERIALNAME" + """"
                    blnInnerJoinMaterial = True
                ElseIf strCampo = "ID_TYPE" Then
                    strsql = """" + "WATERLINESTYPES" + """" + "." + """" + "DESCRIPTION_" + """"
                    blnInnerJoinTipo = True
                Else
                    strsql = """" + "WATERLINES" + """" + "." + """" + strCampo + """"
                End If
            End If
        qtdCampos = 1
        Else
            If frmCanvas.TipoConexao <> 4 Then
                If strCampo = "MATERIAL" Then
                    strsql = strsql & "," & "X_MATERIAL.MATERIALNAME"
                    blnInnerJoinMaterial = True
                ElseIf strCampo = "ID_TYPE" Then
                    strsql = strsql & "," & "WATERLINESTYPES.DESCRIPTION_"
                    blnInnerJoinTipo = True
                Else
                    strsql = strsql & "," & "WATERLINES." & strCampo
                End If
            Else
                If strCampo = "MATERIAL" Then
                    strsql = strsql & "," & """" + "X_MATERIAL" + """" + "." + """" + "MATERIALNAME" + """"
                    blnInnerJoinMaterial = True
                ElseIf strCampo = "ID_TYPE" Then
                    strsql = strsql & "," & """" + "WATERLINESTYPES" + """" + "." + """" + "DESCRIPTION_" + """"
                    blnInnerJoinTipo = True
                Else
                    strsql = strsql & "," & """" + "WATERLINES" + """" + "." & """" + strCampo + """"
                End If
            End If
            qtdCampos = qtdCampos + 1
        End If
    Next
    strCabecalho = ""
    For i = 0 To List2.ListCount - 1
        If strCabecalho = "" Then
            strCabecalho = List2.list(i)
        Else
            strCabecalho = strCabecalho & ";" & List2.list(i)
        End If
    Next
    'strsql = Replace(strsql, ",", ";")
    Open Me.Text1.Text For Output As #1
    Print #1, strCabecalho 'IMPRIME CABEÇALHO
    If blnInnerJoinMaterial = True And blnInnerJoinTipo = True Then 'INNER JOIN DE TIPO E MATERIAL
        If frmCanvas.TipoConexao = 1 Then
            strsql = "SELECT " & strsql & " FROM WATERLINES WATERLINES INNER JOIN X_MATERIAL X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID INNER JOIN WATERLINESTYPES WATERLINESTYPES ON WATERLINESTYPES.ID_TYPE = WATERLINES.ID_TYPE WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 1)"
        ElseIf frmCanvas.TipoConexao = 2 Then
            strsql = "SELECT " & strsql & " FROM WATERLINES WATERLINES INNER JOIN X_MATERIAL X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID INNER JOIN WATERLINESTYPES WATERLINESTYPES ON WATERLINESTYPES.ID_TYPE = WATERLINES.ID_TYPE AND EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE WATERLINES.LINE_ID = P.OBJECT_ID_ AND P.USUARIO = '" & strUser & "' AND P.TIPO = 1)"
        ElseIf frmCanvas.TipoConexao = 4 Then
            a = "WATERLINES"
            b = "X_MATERIAL"
            c = "MATERIALID"
            d = "WATERLINESTYPES"
            e = "OBJECT_ID_"
            f = "POLIGONO_SELECAO"
            g = "USUARIO"
            h = "TIPO"
            ii = "MATERIAL"
            j = "ID_TYPE"
            strsql = "SELECT " & strsql & " FROM " + """" + a + """" + " INNER JOIN " + """" + b + """" + "  ON " + """" + a + """" + "." + """" + ii + """" + " = " + """" + b + """" + "." + """" + c + """" + " INNER JOIN " + """" + d + """" + " ON " + """" + d + """" + "." + """" + j + """" + " = " + """" + a + """" + "." + """" + j + """" + " WHERE " + """" + a + """" + "." + """" + e + """" + " IN (SELECT " + """" + e + """" + " FROM " + """" + f + """" + " WHERE " + """" + g + """" + " = '" & strUser & "' AND " + """" + h + """" + " = '1')"
        End If
    ElseIf blnInnerJoinMaterial = False And blnInnerJoinTipo = True Then 'INNER JOIN SOMENTE DE TIPO
        If frmCanvas.TipoConexao = 1 Then
            strsql = "SELECT " & strsql & " FROM WATERLINES WATERLINES INNER JOIN WATERLINESTYPES WATERLINESTYPES ON WATERLINESTYPES.ID_TYPE = WATERLINES.ID_TYPE WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 1)"
        ElseIf frmCanvas.TipoConexao = 2 Then
            strsql = "SELECT " & strsql & " FROM WATERLINES WATERLINES INNER JOIN WATERLINESTYPES WATERLINESTYPES ON WATERLINESTYPES.ID_TYPE = WATERLINES.ID_TYPE AND EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE WATERLINES.LINE_ID = P.OBJECT_ID_ AND P.USUARIO = '" & strUser & "' AND P.TIPO = 1)"
        ElseIf frmCanvas.TipoConexao = 4 Then
            a = "WATERLINES"
            b = "X_MATERIAL"
            c = "MATERIALID"
            d = "WATERLINESTYPES"
            e = "OBJECT_ID_"
            f = "POLIGONO_SELECAO"
            g = "USUARIO"
            h = "TIPO"
            ii = "MATERIAL"
            j = "ID_TYPE"
            strsql = "SELECT " + """" + strsql + """" + " FROM " + """" + a + """" + " INNER JOIN " + """" + d + """" + """" + d + """" + "  ON " + """" + d + """" + "." + """" + j + """" + " = " + """" + a + """" + "." + """" + j + """" + "  WHERE " + """" + a + """" + "." + """" + e + """" + " IN '(SELECT " + """" + e + """" + " FROM " + """" + f + """" + " WHERE " + """" + g + """" + " = '" & strUser & "' AND " + """" + h + """" + " = '1')'"
        End If
    ElseIf blnInnerJoinMaterial = True And blnInnerJoinTipo = False Then 'INNER JOIN SOMENTE DE MATERIAL
        If frmCanvas.TipoConexao = 1 Then
            strsql = "SELECT " & strsql & " FROM WATERLINES WATERLINES INNER JOIN X_MATERIAL X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 1)"
        ElseIf frmCanvas.TipoConexao = 2 Then
            strsql = "SELECT " & strsql & " FROM WATERLINES WATERLINES INNER JOIN X_MATERIAL X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID AND EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE WATERLINES.LINE_ID = P.OBJECT_ID_ AND P.USUARIO = '" & strUser & "' AND P.TIPO = 1)"
        ElseIf frmCanvas.TipoConexao = 4 Then
            a = "WATERLINES"
            b = "X_MATERIAL"
            c = "MATERIALID"
            d = "WATERLINESTYPES"
            e = "OBJECT_ID_"
            f = "POLIGONO_SELECAO"
            g = "USUARIO"
            h = "TIPO"
            ii = "MATERIAL"
            j = "ID_TYPE"
            strsql = "SELECT " + """" + strsql + """" + " FROM " + """" + a + """" + """" + a + """" + " INNER JOIN " + """" + b + """" + """" + b + """" + "  ON " + """" + a + """" + "." + """" + i + """" + " = " + """" + b + """" + "." + """" + c + """" + "  WHERE " + """" + a + """" + "." + """" + e + """" + " IN '(SELECT " + """" + e + """" + " FROM " + """" + f + """" + " WHERE " + """" + g + """" + " = '" & strUser & "' AND " + """" + h + """" + " = '1')'"
        End If
    Else
        If frmCanvas.TipoConexao = 1 Then   'NENHUM INNER JOIN
            strsql = "SELECT " & strsql & " FROM WATERLINES WATERLINES WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 1)"
        ElseIf frmCanvas.TipoConexao = 2 Then
            strsql = "SELECT " & strsql & " FROM WATERLINES WATERLINES WHERE EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE WATERLINES.LINE_ID = P.OBJECT_ID_ AND P.USUARIO = '" & strUser & "' AND P.TIPO = 1)"
        ElseIf frmCanvas.TipoConexao = 4 Then
            a = "WATERLINES"
            b = "X_MATERIAL"
            c = "MATERIALID"
            d = "WATERLINESTYPES"
            e = "OBJECT_ID_"
            f = "POLIGONO_SELECAO"
            g = "USUARIO"
            h = "TIPO"
            ii = "MATERIAL"
            j = "ID_TYPE"
            strsql = "SELECT " + strsql + " FROM " + """" + a + """" + " WHERE " + """" + a + """" + "." + """" + e + """" + " IN (SELECT " + """" + e + """" + " FROM " + """" + f + """" + " WHERE " + """" + g + """" + " = '" & strUser & "' AND " + """" + h + """" + " = '1')"
        End If
    End If
    Set rs = New ADODB.Recordset
    rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
    'Print #1, strsql
    Do While Not rs.EOF = True
        strPrint = ""
        For i = 0 To qtdCampos - 1
            If rs.Fields(i).value <> "" Then
                If strPrint <> "" Then
                    strPrint = strPrint & ";" & rs.Fields(i).value
                Else
                    strPrint = rs.Fields(i).value
                End If
            Else
                If strPrint <> "" Then
                    strPrint = strPrint & ";"
                Else
                    strPrint = ";"
                End If
            End If
        Next
        strPrint = Replace(strPrint, ".", ",")
        Print #1, strPrint
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Close #1
    MsgBox "Relatório gerado com sucesso!", vbInformation, ""
    OPT3 = True
    Screen.MousePointer = vbNormal
    Exit Function

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Or Err.Number = 55 Then
        Resume Next
    ElseIf Err.Number = 52 Then
        Screen.MousePointer = vbNormal
        MsgBox "Caminho de arquivo incorreto!", vbExclamation, ""
        Err.Clear
    Else
        Screen.MousePointer = vbNormal
        ErroUsuario.Registra "frmRelatoriosAvancados", "OPT3", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function


Private Sub cmdAddCamppo_Click()
'Adiciona um campo na lista de campos selecionados retirando-o da lista de disponíveis

Dim i As Integer
reinicia:
   For i = 0 To List1.ListCount - 1
  
      If List1.Selected(i) = True Then
         List2.AddItem List1.list(i)
         List1.RemoveItem (i)
         List1.Refresh
         GoTo reinicia
      
      End If
   
   Next
End Sub
' Gera o relatório com os componentes que estão dentro do polígono selecionado
'
'
'
Private Sub cmdGerarRelatorio_Click()
    Dim string2 As String
    Dim user As String
    Dim stringFinal As String
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
    Dim retorno As Integer                                                                                      'Se é para processar ou não o relatório que demora para executar

    CommonDialog1.Filter = "Texto (.txt)|*.TXT|Todos tipos (*.*)|*.*|"                                          'configura o filtro do arquivo
    CommonDialog1.FileName = nomeArquivo                                                                        'informa a caixa de diálogo que será aberta o nome do arquivo inicial sugerido
    CommonDialog1.InitDir = diretorioMyDocuments                                                                'sugero o diretório inicial
    CommonDialog1.ShowSave                                                                                      'abre a caixa de diálogo par ao usuário digitar o nome do arquivo e selecionar o diretório, se desejar
    filelocation = CommonDialog1.FileName
    Me.Text1.Text = filelocation
    If Me.Text1.Text <> "" Then ' se há um nome de arquivo...
        If Me.Option1.value = True Then
            If OPT1 = False Then
                Exit Sub
            End If
        ElseIf Me.Option2.value = True Then
            If OPT2 = False Then
                Exit Sub
            End If
        ElseIf Me.Option3.value = True Then
            If OPT3 = False Then
                Exit Sub
            End If
        ElseIf Me.Option4.value = True Then
            If OPT4 = False Then
                Exit Sub
            End If
        ElseIf Me.optRelRamalPadrao.value = True Then                                                           'este é um relatório padrão o qual é executado segundo a querie de número 22 e 23 cadastrada em GS_QUERYS_CLIENT
            If frmCanvas.TipoConexao = 1 Then                                                                   'se for SQLServer
                Set rs = New ADODB.Recordset
                strsql = "SELECT QUERYSTRING FROM GS_QUERYS_CLIENT WHERE QUERY_ID = 22"                         'obtem a querie padrão personalizada pela empresa de saneamento
                string2 = "SELECT QUERYSTRING FROM GS_QUERYS_CLIENT WHERE QUERY_ID = 23"
                rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
                If rs.EOF = False Then
                    strsql = rs(0).value                                                                        'obtem a querie 22
                End If
                rs.Close
                rs.Open string2, Conn, adOpenDynamic, adLockOptimistic
                If rs.EOF = False Then
                    string2 = rs(0).value                                                                       'obtem a querie 23
                End If
                rs.Close
                user = strUser
                stringFinal = strsql + " " + "'" + user + "'" + " " + string2                                   'junta as duas queries colocando o filtro por usuário. Elas estão por usuário pois vão procurar o que foi selecionado pelo usuário na tabela POLIGONO_SELECAO
                retorno = MsgBox("Este relatório irá demorar vários minutos e não poderá ser cancelado. Deseja realmente continuar?", vbYesNo)
                If retorno = vbYes Then
                    If PrintSelect(Me.Text1.Text, stringFinal) = False Then Exit Sub 'É CHAMADO O MÉTODO PRINTSELECT
                    MsgBox "Relatório gerado com sucesso!", vbInformation, ""
                End If
            End If
            If frmCanvas.TipoConexao = 2 Then
                'precisa revisar a implementação
                Set rs = New ADODB.Recordset
                strsql = "SELECT QUERYSTRING FROM GS_QUERYS_CLIENT WHERE QUERY_ID = 22"
                string2 = "SELECT QUERYSTRING FROM GS_QUERYS_CLIENT WHERE QUERY_ID = 23"
                rs.Open strsql, Conn, adOpenForwardOnly, adLockReadOnly
                If rs.EOF = False Then
                    strsql = rs(0).value
                End If
                rs.Close
                rs.Open string2, Conn, adOpenForwardOnly, adLockReadOnly
                If rs.EOF = False Then
                    string2 = rs(0).value
                End If
                rs.Close
                user = strUser
                stringFinal = strsql + " " + "'" + user + "'" + " " + string2
                'É CHAMADO O MÉTODO PRINTSELECT
                If PrintSelect(Me.Text1.Text, stringFinal) = False Then Exit Sub
                MsgBox "Relatório gerado com sucesso!", vbInformation, ""
            End If
            If frmCanvas.TipoConexao = 4 Then
                'precisa revisar a implementação
                Set rs = New ADODB.Recordset
                a = "QUERYSTRING"
                b = "GS_QUERYS_CLIENT"
                c = "QUERY_ID"
                strsql = "SELECT " + """" + a + """" + " FROM " + """" + b + """" + " WHERE " + """" + c + """" + " = '22'"
                string2 = "SELECT " + """" + a + """" + " FROM " + """" + b + """" + " WHERE " + """" + c + """" + " = '23'"
                rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
                If rs.EOF = False Then
                    strsql = rs(0).value
                End If
                rs.Close
                rs.Open string2, Conn, adOpenDynamic, adLockOptimistic
                If rs.EOF = False Then
                    string2 = rs(0).value
                End If
                rs.Close
                user = strUser
                stringFinal = strsql + " " + "'" + user + "'" + " " + string2
                'É CHAMADO O MÉTODO PRINTSELECT
                If PrintSelect(Me.Text1.Text, stringFinal) = False Then Exit Sub
                MsgBox "Relatório gerado com sucesso!", vbInformation, ""
            End If
        End If
    End If
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmRelatoriosAvancados", "cmdGerarRelatorio_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub

Private Sub cmdRemCampo_Click()
'remove um campo selecionado da lista de campos selecionados devolvendo o para a lista de disponíveis

Dim i As Integer
reinicia:
   For i = 0 To List2.ListCount - 1
      If List2.Selected(i) = True Then
         List1.AddItem List2.list(i)
         List2.RemoveItem (i)
         List2.Refresh
         GoTo reinicia
      End If
   Next
End Sub
'evento que modifica a posição do campo dentro do list de campos selecionados, mudando o campo para cima
'
'
'
Private Sub cmdUP_Click()
    On Error GoTo Trata_Erro
    Dim intSELECT As Integer
    Dim i, jjj As Integer
    Dim campo(50) As String

    intSELECT = 100
    For i = 0 To List2.ListCount - 1
        campo(i) = List2.list(i)
        If List2.Selected(i) And i <> 0 Then                                       'verifica se este foi o que o usuário selecionou
            If intSELECT = 100 Then
                intSELECT = (i - 1)
                campo(i) = List2.list(i - 1)
                campo(i - 1) = List2.list(i)
            End If
        End If
    Next
    i = List2.ListCount - 1
    List2.Clear
    For jjj = 0 To i
        If campo(jjj) <> "" Then
            List2.AddItem campo(jjj)
        End If
    Next
    If intSELECT <> 100 Then Me.List2.Selected(intSELECT) = True
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = 381 Or Err.Number = 9 Then
        ErroUsuario.Registra "frmRelatoriosAvancados", "cmdUP_Click - Erro 381 ou 9", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    Else
        ErroUsuario.Registra "frmRelatoriosAvancados", "cmdUP_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'evento que modifica a posição do campo dentro do list de campos selecionados, mudando o campo para baixo
'
'
'
Private Sub cmdDown_Click()
    On Error GoTo Trata_Erro
    Dim i, jjj As Integer
    Dim campo(50) As String
    Dim intSELECT As Integer
    
    For i = 0 To List2.ListCount - 1
        campo(i) = List2.list(i)
        If List2.Selected(i) And List2.ListCount - 1 <> i Then
            intSELECT = (i + 1)
            campo(i) = List2.list(i + 1)
            campo(i + 1) = List2.list(i)
            i = i + 1
        End If
    Next
    List2.Clear
    For jjj = 0 To i - 1
        If campo(jjj) <> "" Then
            List2.AddItem campo(jjj)
        End If
    Next
    Me.List2.Selected(intSELECT) = True
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = 381 Or Err.Number = 9 Then
        ErroUsuario.Registra "frmRelatoriosAvancados", "cmdDown_Click - Erro 381 ou 9", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    Else
        ErroUsuario.Registra "frmRelatoriosAvancados", "cmdDown_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
' Carrega o formulário inicial de geração de relatórios
'
'
'
Private Sub Form_Load()
    On errror GoTo Trata_Erro:
    Dim Texto As String
    Dim rs As ADODB.Recordset
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
    
    Set rs = New ADODB.Recordset
    nomeArquivo = selecionaArquivo.ConfiguraNomeArquivo("exportação_dados_redes", "txt", diretorioMyDocuments)  'obtem o nome do arquivo sugerido e o diretório meus documentos do usuário
    Me.Text1 = diretorioMyDocuments & "\" & nomeArquivo                                                         'coloca na caixa de diálogo o nome sugerido
    Me.Height = 7000
    a = "QUERYSTRING"
    b = "GS_QUERYS_CLIENT"
    c = "QUERY_ID"
    If frmCanvas.TipoConexao <> 4 Then                                                                          'se não for Postgres
        Texto = Replace(RetornaCabecalho("SELECT QUERYSTRING FROM GS_QUERYS_CLIENT WHERE QUERY_ID = 22", "SELECT QUERYSTRING FROM GS_QUERYS_CLIENT WHERE QUERY_ID = 23"), ";", "; ")
    Else                                                                                                        'se for Postgres
        Texto = Replace(RetornaCabecalho("SELECT " + """" + a + """" + " FROM " + """" + b + """" + " WHERE " + """" + c + """" + " = '22'", "SELECT " + """" + a + """" + " FROM " + """" + b + """" + " WHERE " + """" + c + """" + " = '23'"), ";", "; ")
    End If
    If Trim(Texto) = "" Then
        Texto = "Não configurado no banco de dados (QUERY_ID=22, QUERY_ID=23)"
        ErroUsuario.Registra "frmRelatoriosAvancados", "Form_Load", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails, Texto
        Exit Sub
    End If
    optRelRamalPadrao.Caption = "Relatório personalizado com as seguintes informações: " & Texto                              'mostra a querie na caixa de diálogo
    'no Form Load, a estrutura de colunas da tabela waterlines é copiada para a parte de relatórios avançados
    'ocorrendo uma 'tradução' para que, montar o relatório, fique mais simplificado ao usuário do sistema
    If frmCanvas.TipoConexao = 1 Then
        rs.Open "SELECT upper(NAME) as CAMPOS FROM SYSCOLUMNS WHERE ID IN (SELECT ID FROM SYSOBJECTS WHERE NAME = 'WATERLINES') ORDER BY CAMPOS", Conn
        If rs.EOF = False Then
            Do While Not rs.EOF
                If rs!campos = "DATA_LOG" Then
                    List1.AddItem "DATA DE DESENHO"
                ElseIf rs!campos = "DATALOG" Then ' não é copiado
                    'List1.AddItem
                ElseIf rs!campos = "DATEINSTALLATION" Then
                    List1.AddItem "DATA DE INSTALAÇÃO"
                ElseIf rs!campos = "DIVIDEDDISTANCE" Then
                    List1.AddItem "DISTÂNCIA DA DIVISA"
                ElseIf rs!campos = "EXTERNALDIAMETER" Then
                    List1.AddItem "DIÂMETRO EXTERNO"
                ElseIf rs!campos = "FINALCOMPONENT" Then
                    List1.AddItem "CÓD COMPONENTE FINAL"
                ElseIf rs!campos = "FINALGROUNDHEIGHT" Then
                    List1.AddItem "COTA TERRENO FINAL"
                ElseIf rs!campos = "FINALTUBEDEEPNESS" Then
                    List1.AddItem "PROFUNDIDADE FINAL"
                ElseIf rs!campos = "ID_TYPE" Then
                    List1.AddItem "TIPO DE REDE"
                ElseIf rs!campos = "INFORMATIONVALIDITY" Then
                    List1.AddItem "VALIDADE"
                ElseIf rs!campos = "INITIALCOMPONENT" Then
                    List1.AddItem "CÓD COMPONENTE INICAL"
                ElseIf rs!campos = "INITIALGROUNDHEIGHT" Then
                    List1.AddItem "COTA TERRENO INICIAL"
                ElseIf rs!campos = "INITIALTUBEDEEPNESS" Then
                    List1.AddItem "PROFUNDIDADE INICIAL"
                ElseIf rs!campos = "INTERNALDIAMETER" Then
                    List1.AddItem "DIÂMETRO INTERNO"
                ElseIf rs!campos = "LENGTH" Then
                    List1.AddItem "COMPRIM. DIGITADO"
                ElseIf rs!campos = "LENGTHCALCULATED" Then
                    List1.AddItem "COMPRIM. CALCULADO"
                ElseIf rs!campos = "LINE_ID" Then ' LINE_ID é o mesmo que OBJECT_ID_
                    'List1.AddItem ""
                ElseIf rs!campos = "LOCATION" Then
                    List1.AddItem "LOCALIZAÇÃO"
                ElseIf rs!campos = "MANUFACTURER" Then
                    List1.AddItem "FABRICANTE"
                ElseIf rs!campos = "MATERIAL" Then
                    List1.AddItem "MATERIAL NOME"
                ElseIf rs!campos = "OBJECT_ID_" Then
                    List1.AddItem "ID REDE"
                ElseIf rs!campos = "ROUGHNESS" Then
                    List1.AddItem "RUGOSIDADE"
                ElseIf rs!campos = "SECTOR" Then
                    List1.AddItem "CÓD SETOR"
                ElseIf rs!campos = "SIDESTREET" Then
                    List1.AddItem "LADO DA RUA"
                ElseIf rs!campos = "STATE" Then
                    List1.AddItem "ESTADO"
                ElseIf rs!campos = "SUPPLIER" Then
                    List1.AddItem "FORNECEDOR"
                ElseIf rs!campos = "THICKNESS" Then
                    List1.AddItem "DENSIDADE"
                ElseIf rs!campos = "TROUBLE" Then
                    'List1.AddItem ""
                ElseIf rs!campos = "USUARIO_LOG" Then
                    List1.AddItem "USUARIO CADASTRO"
                End If
                rs.MoveNext
            Loop
        End If
    End If
    If frmCanvas.TipoConexao = 2 Then
        List1.AddItem "DATA DE DESENHO"
        List1.AddItem "DATA DE INSTALAÇÃO"
        List1.AddItem "DISTÂNCIA DA DIVISA"
        List1.AddItem "DIÂMETRO EXTERNO"
        List1.AddItem "CÓD COMPONENTE FINAL"
        List1.AddItem "COTA TERRENO FINAL"
        List1.AddItem "PROFUNDIDADE FINAL"
        List1.AddItem "TIPO DE REDE"
        List1.AddItem "VALIDADE"
        List1.AddItem "CÓD COMPONENTE INICAL"
        List1.AddItem "COTA TERRENO INICIAL"
        List1.AddItem "PROFUNDIDADE INICIAL"
        List1.AddItem "DIÂMETRO INTERNO"
        List1.AddItem "COMPRIM. DIGITADO"
        List1.AddItem "COMPRIM. CALCULADO"
        List1.AddItem "LOCALIZAÇÃO"
        List1.AddItem "FABRICANTE"
        List1.AddItem "MATERIAL NOME"
        List1.AddItem "ID REDE"
        List1.AddItem "RUGOSIDADE"
        List1.AddItem "CÓD SETOR"
        List1.AddItem "LADO DA RUA"
        List1.AddItem "ESTADO"
        List1.AddItem "FORNECEDOR"
        List1.AddItem "DENSIDADE"
        List1.AddItem "USUARIO CADASTRO"
    End If
    If frmCanvas.TipoConexao = 4 Then
        List1.AddItem "DATA DE DESENHO"
        List1.AddItem "DATA DE INSTALAÇÃO"
        List1.AddItem "DISTÂNCIA DA DIVISA"
        List1.AddItem "DIÂMETRO EXTERNO"
        List1.AddItem "CÓD COMPONENTE FINAL"
        List1.AddItem "COTA TERRENO FINAL"
        List1.AddItem "PROFUNDIDADE FINAL"
        List1.AddItem "TIPO DE REDE"
        List1.AddItem "VALIDADE"
        List1.AddItem "CÓD COMPONENTE INICAL"
        List1.AddItem "COTA TERRENO INICIAL"
        List1.AddItem "PROFUNDIDADE INICIAL"
        List1.AddItem "DIÂMETRO INTERNO"
        List1.AddItem "COMPRIM. DIGITADO"
        List1.AddItem "COMPRIM. CALCULADO"
        List1.AddItem "LOCALIZAÇÃO"
        List1.AddItem "FABRICANTE"
        List1.AddItem "MATERIAL NOME"
        List1.AddItem "ID REDE"
        List1.AddItem "RUGOSIDADE"
        List1.AddItem "CÓD SETOR"
        List1.AddItem "LADO DA RUA"
        List1.AddItem "ESTADO"
        List1.AddItem "FORNECEDOR"
        List1.AddItem "DENSIDADE"
        List1.AddItem "USUARIO CADASTRO"
    End If
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmRelatoriosAvancados", "Form_Load", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub

Private Sub Option1_Click()
   MIN_FORM
End Sub

Private Sub Option2_Click()
   MIN_FORM
End Sub

Private Sub Option3_Click()
   Me.Height = 12235

   
End Sub

Private Function MIN_FORM()
   Me.Height = 7000

End Function


Private Sub Option4_Click()
   MIN_FORM
End Sub

Private Sub optRelRamalPadrao_Click()
   
   'frmRelAvanComando.Text1.Text = Me.optRelRamalPadrao.Caption
   'frmRelAvanComando.Show 1

End Sub

