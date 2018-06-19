VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   4995
      TabIndex        =   1
      Top             =   2040
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Import"
      Height          =   375
      Left            =   4980
      TabIndex        =   0
      Top             =   1545
      Width           =   1140
   End
   Begin VB.PictureBox OdImport 
      Height          =   480
      Left            =   930
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   690
      Width           =   1200
   End
   Begin VB.PictureBox OdExport 
      Height          =   480
      Left            =   1290
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   1245
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Dim lineSet As OdLineSettings
'
''tpPOLYGONS = 1
''tpLINES = 2
''tpPOINTS = 4
''tpTEXTS = 128
'
'
'Private Sub Command1_Click()
'
'OdImport.Provider = 2
'OdImport.Connection = Conn
'OdImport.Open ("C:\Projeto_GeoSan\RUI\ruas.dwg")
'
'Dim i As Integer
'Dim Layer As String
'
'i = OdImport.getLayerCount
'For i = 0 To i
'
'   Layer = OdImport.getLayerName(i)
'Next
'
'If OdImport.importToLayer("TESTE_IMP_SEMXY34", "DESENHO") = True Then
'   MsgBox "IMPORTADO"
'Else
'   MsgBox "ERRO NA IMPORTAÇÃO"
'End If
'
'
'
'End Sub
''
''
'Private Sub cmdExport_Click()
'
'
'
'
'
'' SELECT Object_id_ FROM POLIGONO_SELECAO WHERE USUARIO = 'a' and Tipo = 1 = OK
'
'
'
'   If OdExport.buildLayer("WATERLINES", 2, " where object_id in (SELECT Object_id_ FROM POLIGONO_SELECAO WHERE USUARIO = 'a' and Tipo = 1  )") = True Then
'   End If
'
'      ' where object_id in (SELECT Object_id_ FROM POLIGONO_SELECAO WHERE USUARIO = 'a' and Tipo = 0)
'   If OdExport.buildLayer("WATERCOMPONENTS", 4, "") = True Then
'
'   End If
'
'   If OdExport.buildLayer("WATERLINES", 128, " where object_id in (SELECT Object_id_ FROM POLIGONO_SELECAO WHERE USUARIO = 'a' and Tipo = 1  )") = True Then
'   End If
'
'
'   If OdExport.Save("C:\Projeto_GeoSan\export_16501.dxf", Dxf, ACAD2000) = True Then
'
'
'   End If
'
'
'End Sub
'''
''Private Sub Command3_Click()
''
''
''      strsql = "SELECT * from waterlines"
''
''      'strsql = "SELECT OBJECT_ID_ AS RAMAL, SUM(CONSUMO_LPS) AS CONSUMO, SUM(ECONOMIAS) AS ECONOMIAS, TIPO, HIDROMETRADO FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2) GROUP BY OBJECT_ID_,TIPO,HIDROMETRADO"
''      Dim rs As New ADODB.Recordset
''
''         If strsql <> "" Then
''            Set rs = New ADODB.Recordset
''
''            rs.Open strsql, Conn, adOpenForwardOnly, adLockReadOnly
''
''            'monta a string de colunas
''            NomeCol = rs.Fields(0).Name
''            For i = 1 To rs.Fields.Count - 1
''               NomeCol = NomeCol & ";" & rs.Fields(i).Name
''            Next
''
''            'obtenho o número de colunas e o número de linhas
''            dbvetor = rs.GetRows
''
''            colunas = UBound(dbvetor, 1)
''            registros = UBound(dbvetor, 2)
''
''            Open "C:\GeoTeste.txt" For Append As #1
''            Print #1, NomeCol
''            For i = 0 To registros
''               For j = 0 To colunas
''                  LINHA = LINHA & dbvetor(j, i) & ";"
''               Next j
''                  Print #1, LINHA
''                  LINHA = ""
''            Next i
''            Close #1
''
''        End If
''
''End Sub
'
'
'
'
'Private Sub Form_Load()
'
'   OdExport.Provider = frmCanvas.TipoConexao
'   OdExport.Connection = Conn
'
'End Sub
'
'
'Private Sub OdExport_buildingGeometry(ByVal representation As Long, ByVal geom_id As Long, ByVal object_id As String)
'
'
'If representation = 2 Then ' LINHAS
'
'   Dim lineSet As New OdLineSettings
'
'   If Len(object_id) = 3 Then
'      lineSet.lineRColor = 0
'      lineSet.lineGColor = 0
'      lineSet.lineBColor = 255
'      lineSet.lineWidth = 2
'   ElseIf Len(object_id) = 4 Then
'      lineSet.lineRColor = 255
'      lineSet.lineGColor = 0
'      lineSet.lineBColor = 0
'      lineSet.lineWidth = 7
'   Else
'      lineSet.lineRColor = 0
'      lineSet.lineGColor = 255
'      lineSet.lineBColor = 0
'      lineSet.lineWidth = 5
'   End If
'
'   OdExport.setLineStyle lineSet
'
'ElseIf representation = 128 Then ' TEXTOS
'
'   Dim textSet As New OdTextSettings
'
'   textSet.FontName = "Arial"
'   textSet.FontSize = 2
'   textSet.textAngle = 0
'   textSet.textRColor = 255
'   textSet.textGColor = 255
'   textSet.textBColor = 255
'
'   OdExport.setTextStyle textSet
'
'ElseIf representation = 4 Then ' PONTOS
'
'   Dim pointSet As New OdPointSettings
'
'   pointSet.FileName = ""
'   pointSet.pointRColor = 0
'   pointSet.pointGColor = 0
'   pointSet.pointBColor = 255
'   pointSet.pointSize = 4
'
'   OdExport.setPointStyle pointSet
'
'End If
'
'
'End Sub
'
