VERSION 5.00
Begin VB.Form frmAlteraConsumoPorPoligono 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alteração de Consumo por Polígono"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAtribuir 
      Caption         =   "Atribuir"
      Height          =   405
      Left            =   6615
      TabIndex        =   7
      Top             =   4860
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Fechar"
      Height          =   405
      Left            =   5385
      TabIndex        =   6
      Top             =   4860
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ligações de água selecionadas pelo polígono"
      Height          =   4485
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   7650
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "  Tipo                 Quantidade  Economias  Consumo LPS     Medido"
         Top             =   345
         Width           =   7275
      End
      Begin VB.Frame Frame7 
         Caption         =   "Consumo (médio/ligação)"
         Height          =   1035
         Left            =   180
         TabIndex        =   2
         Top             =   3240
         Width           =   2745
         Begin VB.OptionButton optLitrosSegundo 
            Caption         =   "LPS"
            Height          =   285
            Left            =   150
            TabIndex        =   5
            Top             =   615
            Width           =   870
         End
         Begin VB.OptionButton optMetroCubico 
            Caption         =   "M³/Mês"
            Height          =   285
            Left            =   150
            TabIndex        =   4
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.TextBox txtConsumo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1290
            TabIndex        =   3
            Text            =   "0.00"
            ToolTipText     =   "Informe o consumo médio de uma ligação"
            Top             =   435
            Width           =   1215
         End
      End
      Begin VB.ListBox lstTipos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         ItemData        =   "frmAlteraConsumoPorPoligono.frx":0000
         Left            =   180
         List            =   "frmAlteraConsumoPorPoligono.frx":0007
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   630
         Width           =   7275
      End
   End
End
Attribute VB_Name = "frmAlteraConsumoPorPoligono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtribuir_Click()
On Error GoTo Trata_Erro

Dim i, j As Integer
Dim strCMD As String
Dim strTipo As String
Dim strHidro As String
Dim dblConsumo As Double
Dim strConsumo As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim ii As String
Dim jj As String
Dim k As String
Dim l As String

   'CAPTA QUAIS TIPOS SERÃO ALTERADOS
   MousePointer = vbHourglass

   For i = 0 To lstTipos.ListCount - 1
      If lstTipos.Selected(i) = True Then
         strTipo = ""
         For j = 1 To Len(Me.lstTipos.list(i)) 'PROCURA O TIPO DA LIGAÇÃO
            If IsNumeric(mid(Me.lstTipos.list(i), j, 1)) = False Then
               strTipo = strTipo & mid(Me.lstTipos.list(i), j, 1)
            Else
               strTipo = Trim(strTipo)
               Exit For
            End If
         Next
         
         strHidro = ""
         j = 0
         Do While Not j = 3    'PROCURA O sim ou nao PARA HIDROMETRADO
            j = j + 1
            f = Len(Me.lstTipos.list(i)) - j + 1
            If IsNumeric(mid(Me.lstTipos.list(i), f, 1)) = False Then
               strHidro = mid(Me.lstTipos.list(i), f, 1) & strHidro
            End If
         Loop
     
         If Trim(strTipo) = "" Then
            MousePointer = vbDefault
            MsgBox "Não foi possível identificar o tipo da ligação.", vbExclamation, ""
            Exit Sub
         End If
      
         'CAPTURA O CONSUMO DIGITADO E CONVERTE SE NECESSÁRIO
         If CDbl(Me.txtConsumo.Text) > 0 Then
            If Me.optMetroCubico.value = True Then
                'SE FOR METRO CUBICO, CONVERTE PARA LITROS POR SEGUNDO
                dblConsumo = Replace(Me.txtConsumo.Text, ".", ",") * 0.00038580246
            Else
                dblConsumo = Replace(Me.txtConsumo.Text, ".", ",")
            End If
         Else
            dblConsumo = 0
         End If
         
         strConsumo = Replace(dblConsumo, ",", ".") ' converte a virgula para ponto para comando SQL
         
         If frmCanvas.TipoConexao = 1 Then
         
            strCMD = "UPDATE RAMAIS_AGUA_LIGACAO SET CONSUMO_LPS = '" & strConsumo & "' WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2) and TIPO = '" & strTipo & "' AND HIDROMETRADO = '" & strHidro & "'"
         
         ElseIf frmCanvas.TipoConexao = 2 Then
         
            strCMD = "UPDATE RAMAIS_AGUA_LIGACAO R SET CONSUMO_LPS = '" & strConsumo & "' WHERE EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE R.OBJECT_ID_ = P.OBJECT_ID_ AND P.USUARIO = '" & strUser & "' AND P.TIPO = 2) and TIPO = '" & strTipo & "' AND HIDROMETRADO = '" & strHidro & "'"
           
           ElseIf frmCanvas.TipoConexao = 4 Then
a = "RAMAIS_AGUA_LIGACAO"
b = "CONSUMO_LPS"
c = "OBJECT_ID_"
d = "POLIGONO_SELECAO"
e = "TIPO"
f = "HIDROMETRADO"

 
           strCMD = "UPDATE " + """" + a + """" + " SET " + """" + b + """" + " = '" & strConsumo & "' WHERE " + """" + c + """" + " IN (SELECT " + """" + c + """" + " FROM " + """" + d + """" + " WHERE " + """" + "USUARIO" + """" + " = '" & strUser & "' AND " + """" + e + """" + " = '2') and " + """" + e + """" + " = '" & strTipo & "' AND " + """" + f + """" + " = '" & strHidro & "'"
               
         End If
         
         Conn.execute (strCMD)

          
      End If
   
   Next
   
   CarregaList
   
   MousePointer = vbDefault

   
'   'FAZ UM SELECT CONTANDO QUANTAS LIGAÇÕES SERÃO AFETADAS PELO COMANDO
'   Dim rs As New ADODB.Recordset
'
'   strCMD = "SELECT COUNT(NRO_LIGACAO) AS QTD FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN ("
'   strCMD = strCMD & "SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2)"
'   strCMD = strCMD & ") AND TIPO IN (" & strTP & ")"
'
'
'
'   If rs.EOF = False Then
'      If CLng(rs!qtd) > 0 Then
'
'         If MsgBox("De acordo com a seleção, serão alteradas " & rs!qtd & " ligações." & Chr(13) & Chr(13) & "Deseja continuar?", vbDefaultButton2 + vbQuestion + vbYesNo, "") = vbYes Then
'
'            'PREPARA O COMANDO SQL
'            strCMD = "UPDATE RAMAIS_AGUA_LIGACAO SET CONSUMO_LPS = " & strConsumo & " WHERE OBJECT_ID_ IN ("
'            strCMD = strCMD & "SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2)"
'            strCMD = strCMD & ") AND TIPO IN (" & strTP & ")"
'
'            MousePointer = vbHourglass
'
'            'EXECUTA A ATUALIZAÇÃO
'            Conn.execute (strCMD)
'
'            MousePointer = vbDefault
'            MsgBox "Atualização concluída!", vbInformation, ""
'         End If
'
'      Else
'         MsgBox "Nenhuma ligação do Tipo selecionado foi encontrada na selecão.", vbInformation, ""
'      End If
'      Me.cmdCancelar.Caption = "Fechar"
'   End If
'   rs.Close


Trata_Erro:

   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   ElseIf Err.Number = 13 Then
      MsgBox "Insira somente números para valores de consumo.", vbExclamation, ""
      Err.Clear
   Else
      MsgBox Err.Number & " " & Err.Description
      Err.Clear
   End If
   MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   CarregaList

End Sub
Private Function CarregaList() As Boolean
On Error GoTo Trata_Erro
   
   Dim rs As New ADODB.Recordset
   Dim str As String
   
   Dim strTipo As String
   Dim strQTD As String
   Dim strEcon As String
   Dim strCons As String
   Dim strHidro As String
   
   Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim ii As String
Dim jj As String
Dim k As String
Dim l As String

a = "TIPO"
b = "ECONOMIAS"
c = "CONSUMO_LPS"
d = "HIDROMETRADO"
e = "RAMAIS_AGUA_LIGACAO"
f = "OBJECT_ID_"
g = "POLIGONO_SELECAO"
h = "USUARIO"
   
   If frmCanvas.TipoConexao = 1 Then
   
      str = "SELECT DISTINCT TIPO, COUNT(TIPO) AS QTD, SUM(ECONOMIAS) AS ECON, SUM(CONSUMO_LPS) AS CONSUMO,HIDROMETRADO AS HIDRO FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2) GROUP BY TIPO, HIDROMETRADO ORDER BY TIPO, HIDROMETRADO"
   
   ElseIf frmCanvas.TipoConexao = 2 Then
      
      str = "SELECT DISTINCT TIPO, COUNT(TIPO) AS " + """" + "QTD" + """" + ", SUM(ECONOMIAS) AS " + """" + "ECON" + """" + ", SUM(CONSUMO_LPS) AS " + """" + "CONSUMO" + """" + ",HIDROMETRADO AS " + """" + "HIDRO" + """" + " FROM RAMAIS_AGUA_LIGACAO RAL WHERE EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE RAL.OBJECT_ID_ = P.OBJECT_ID_ AND P.USUARIO = '" & strUser & "' AND P.TIPO = 2) GROUP BY TIPO, HIDROMETRADO ORDER BY TIPO, HIDROMETRADO"
   
  Else
   'SELECT DISTINCT "TIPO", COUNT("TIPO") AS QTD, SUM("ECONOMIAS") AS ECON, SUM("CONSUMO_LPS") AS CONSUMO,"HIDROMETRADO" AS HIDRO FROM "RAMAIS_AGUA_LIGACAO" WHERE EXISTS (SELECT 1 FROM "POLIGONO_SELECAO" WHERE "RAMAIS_AGUA_LIGACAO"."OBJECT_ID_" = "POLIGONO_SELECAO"."OBJECT_ID_" AND "POLIGONO_SELECAO"."USUARIO" = 'Administrador' AND "POLIGONO_SELECAO"."TIPO" = '2') GROUP BY "TIPO", "HIDROMETRADO" ORDER BY "TIPO", "HIDROMETRADO"
    

   str = "SELECT DISTINCT " + """" + a + """" + ", COUNT(" + """" + a + """" + ") AS " + """" + "QTD" + """" + ", SUM(" + """" + b + """" + ") AS " + """" + "ECON" + """" + ", SUM(" + """" + c + """" + ") AS " + """" + "CONSUMO" + """" + "," + """" + d + """" + " AS " + """" + "HIDRO" + """" + " FROM " + """" + e + """" + " WHERE " + """" + f + """" + "IN (SELECT " + """" + f + """" + " FROM " + """" + g + """" + " WHERE " + """" + h + """" + " = '" & strUser & "' AND " + """" + a + """" + " = 2) GROUP BY " + """" + a + """" + ", " + """" + d + """" + " ORDER BY " + """" + a + """" + ", " + """" + d + """" + ""
      'MsgBox str
      
      
      
      


'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
   End If
   
   
   'O RESULTADO DO SELECT É ALGO COMO TIPO
   
   'CORTADO             61    65    .11612643   nao
   'CORTADO             132   143   .50921627   sim
   'FICTÍCIA            3     3     .00000000   nao
   'LIGADO              109   119   .24729915   nao
   'LIGADO              1748  2171  8.50811021  sim
   'LIGADO EM ANALISE   152   184   .49035469   nao
   'LIGADO EM ANALISE   229   251   .54128078   sim
   'SUPRIMIDO           6     6     .00810185   nao
   
   Me.lstTipos.Clear

   Set rs = New ADODB.Recordset
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic
   
   If rs.EOF = False Then
      Do While Not rs.EOF

         If rs!tipo <> "" Then strTipo = rs!tipo Else strTipo = ""
         If rs!qtd <> "" Then strQTD = rs!qtd Else strQTD = ""
         If rs!econ <> "" Then strEcon = rs!econ Else strEcon = ""
         If rs!CONSUMO <> "" Then strCons = rs!CONSUMO Else strCons = ""
         If rs!hidro <> "" Then strHidro = rs!hidro Else strHidro = ""
         
         'os campos são redimensionados para formar uma grade virtual no ListView
         str = strTipo & Space(21 - Len(strTipo)) & strQTD & Space(12 - Len(strQTD)) & strEcon & Space(11 - Len(strEcon)) & strCons & Space(16 - Len(strCons)) & strHidro
         
         Me.lstTipos.AddItem str
         
         rs.MoveNext
      DoEvents
      Loop
   End If
   
   rs.Close
   Set rs = Nothing
   
   
Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   MsgBox Err.Number & " - " & Err.Description
End If

End Function

