VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmConsumoLoteGraf 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Grafico"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6330
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4935
      Left            =   0
      OleObjectBlob   =   "frmConsumoLoteGraf.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmConsumoLoteGraf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub init(Periodo1 As Integer, Periodo2 As Integer, Ano1 As Integer, Ano2 As Integer, _
         CLASSIFICACAO_FISCAL As String, NRO_LIGACAO As String, Agrupado As Boolean)
On Error GoTo Trata_Erro
   
  
   Screen.MousePointer = vbHourglass
   
   Dim str As String
   Dim i As Integer, j As Integer
   Dim rs As New ADODB.Recordset
   Dim gu1 As String
   Dim gu2 As String
   Dim gu3 As String
   Dim gu4 As String
   Dim gu5 As String
   Dim gu6 As String
   Dim gu7 As String
   Dim gu8 As String
   Dim gu9 As String
   Dim gu10 As String
   Dim gu11 As String
  gu1 = "QUERYSTRING"
   gu2 = "GS_QUERYS_CLIENT"
   gu3 = "CLIENT_ID"
   gu4 = "QUERY_ID"
   gu5 = "ligacao"
   gu6 = "histconsumo"
   gu7 = "PERIODO"
   gu8 = "ano"
   gu9 = "NRO_LIGACAO"
   gu10 = "CLASSIFICACAO_FISCAL"
   gu11 = "CONSUMO_MEDIDO"
   If frmCanvas.TipoConexao <> 4 Then

   
   If Agrupado Then
      
      Set rs = Conn.execute("SELECT querystring from gs_querys_client where client_id=3 and query_id=5")
      str = rs.Fields("querystring").value
      
      
      str = Replace(str, "LI.CLASSIFICACAO_FISCAL = '' and", "")
      
      str = Replace(str, "@CLASSIFICACAO_FISCAL", CLASSIFICACAO_FISCAL)
      str = Replace(str, "@PERIODO1", Periodo1)
      str = Replace(str, "@PERIODO2", Periodo2)
      str = Replace(str, "@ANO1", Ano1)
      str = Replace(str, "@ANO2", Ano2)
      str = Replace(str, "@NRO_LIGACAO", NRO_LIGACAO)
'Imprima str
      rs.Close
   Else
      MsgBox "Você só pode apresentar gráfico de apenas uma linha", vbExclamation
      Exit Sub
      str = "SELECT h.PERIODO || '/' || h.ano as Mes_Ano,"
      str = str & "     l.NRO_LIGACAO,"
      str = str & "     h.CONSUMO_MEDIDO as M3"
      str = str & " from ligacao l, histconsumo h"
      str = str & " Where l.NRO_LIGACAO = h.NRO_LIGACAO And h.ANO >= " & Ano1
      str = str & "             and h.PERIODO >= " & Periodo1 & " and h.ANO <= " & Ano2 & " and h.PERIODO <= " & Periodo2
      str = str & "             and l.CLASSIFICACAO_FISCAL = '" & CLASSIFICACAO_FISCAL & "'"
      str = str & "             and l.NRO_LIGACAO IN(" & NRO_LIGACAO & ")"
      
     
   End If
   
   Else
      If Agrupado Then
      
      Set rs = Conn.execute("SELECT " + """" + gu1 + """" + " from " + """" + gu2 + """" + " where " + """" + gu3 + """" + "='3' and " + """" + gu4 + """" + "='5'")
      str = rs.Fields("querystring").value
      
      
      str = Replace(str, "LI.CLASSIFICACAO_FISCAL = '' and", "")
      
      str = Replace(str, "@CLASSIFICACAO_FISCAL", CLASSIFICACAO_FISCAL)
      str = Replace(str, "@PERIODO1", Periodo1)
      str = Replace(str, "@PERIODO2", Periodo2)
      str = Replace(str, "@ANO1", Ano1)
      str = Replace(str, "@ANO2", Ano2)
      str = Replace(str, "@NRO_LIGACAO", NRO_LIGACAO)
'Imprima str
      rs.Close
   Else
      MsgBox "Você só pode apresentar gráfico de apenas uma linha", vbExclamation
      Exit Sub
       str = "SELECT " + """" + gu6 + """" + "." + """" + gu7 + """" + " || '/' || " + """" + gu6 + """" + ".ano as Mes_Ano,"
      str = str & "     " + """" + gu5 + """" + "." + """" + gu9 + """" + ","
      str = str & "     " + """" + gu6 + """" + "." + """" + gu11 + """" + " as " + """" + " M3" + """" + ""
      str = str & " from " + """" + gu5 + """" + " , " + """" + gu6 + """" + " "
      str = str & " Where " + """" + gu5 + """" + "." + """" + gu9 + """" + " = " + """" + gu6 + """" + "." + """" + gu9 + """" + " And " + """" + gu6 + """" + "." + """" + gu8 + """" + " >= " & Ano1
      str = str & "             and " + """" + gu6 + """" + "." + """" + gu7 + """" + " >= " & Periodo1 & " and " + """" + gu6 + """" + "." + """" + gu8 + """" + " <= " & Ano2 & " and " + """" + gu6 + """" + "." + """" + gu7 + """" + " <= " & Periodo2
      str = str & "             and " + """" + gu5 + """" + "." + """" + gu10 + """" + " = '" & CLASSIFICACAO_FISCAL & "'"
      str = str & "             and " + """" + gu5 + """" + "." + """" + gu9 + """" + " IN(" & NRO_LIGACAO & ")"
   End If
   End If
   
   rs.CursorType = 3
   rs.Open str, ConnSec, adOpenDynamic, adLockOptimistic
   With MSChart1
      .chartType = VtChChartType2dLine
      .ColumnCount = rs.Fields.Count - 1
      .RowCount = rs.RecordCount
      .ShowLegend = False
      i = 1
      While Not rs.EOF
         .Row = i
         For j = 1 To .ColumnCount
            .Column = j
            .ColumnLabel = rs.Fields(j).Name
            .Data = IIf(IsNull(rs.Fields(j).value), 0, rs.Fields(j).value)
         Next
         .RowLabel = rs.Fields(0).value
         rs.MoveNext
         i = i + 1
      Wend
   End With
   rs.Close
   Set rs = Nothing
    Screen.MousePointer = vbNormal
    Me.Show vbModal
    
   
Trata_Erro:
    
    Screen.MousePointer = vbNormal
    
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        
        PrintErro CStr(Me.Name), "Private Sub Init", CStr(Err.Number), CStr(Err.Description), True
        
    End If
End Sub






