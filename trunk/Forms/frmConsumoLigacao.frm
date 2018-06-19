VERSION 5.00
Begin VB.Form frmConsumoLote 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consultar Consumo"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAgrupar 
      Caption         =   "AGRUPAR LIGAÇÕES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4110
      TabIndex        =   10
      Top             =   3180
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin VB.PictureBox dtData 
      Height          =   375
      Index           =   0
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox lvLigacoes 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   60
      ScaleHeight     =   1875
      ScaleWidth      =   6315
      TabIndex        =   3
      Top             =   360
      Width           =   6375
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   5100
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "Grafico"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox lvConsumo 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   60
      ScaleHeight     =   1755
      ScaleWidth      =   6315
      TabIndex        =   5
      Top             =   3480
      Width           =   6375
   End
   Begin VB.PictureBox dtData 
      Height          =   375
      Index           =   1
      Left            =   1860
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Inicial:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "CONSUMO:"
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
      Left            =   60
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "CONSUMIDORES:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   90
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Final:"
      Height          =   255
      Left            =   1860
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "frmConsumoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Inscricao As String
       Dim ve As String
         Dim vi As String
         Dim vo As String
         Dim vu As String
         Dim vc As String
          Dim vd As String
          Dim ve As String
          Dim vf As String

Public Sub Init(object_id As String, layer_name As String)
On Error GoTo Trata_Erro

   If ConnSec.State = 0 Then
        MsgBox "Não há conexão disponível com o banco de dados comercial para executar esta função.", vbInformation, "Conexão não disponível"
        Exit Sub
   End If
      
   If layer_name <> GetLayerLot() Then ' forçar como LOTES
      MsgBox "O plano de lotes não está cadastrado corretamente ou não está selecionado  " & GetLayerLot, vbExclamation
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   Dim i As ListItem, rs As ADODB.Recordset, str As String
   va = """RAMAIS_AGUA_LIGACAO"""
         ve = """querystring"""
         vi = """gs_querys_client"""
         vo = """query_id"""
         vu = """client_id"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute("Select querystring from gs_querys_client where query_id=3 and client_id=1")
   Else
    Set rs = Conn.execute("Select " + ve + " from " + vi + " where " + vo + "='3' and " + vu + "='1'")
   
   
   End If
   str = rs.Fields("querystring").Value
   str = Replace(str, "@OBJECT_ID_", object_id)
   rs.Close
   Set rs = Conn.execute(str) 'select inscrição from lotes where object_id_in()
   Inscricao = rs(0).Value
   va = """RAMAIS_AGUA_LIGACAO"""
         ve = """querystring"""
         vi = """gs_querys_client"""
         vo = """query_id"""
         vu = """client_id"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute("select querystring from gs_querys_client where client_id=1 and query_id=2")
   Else
   Set rs = Conn.execute("select " + ve + " from " + vi + " where " + vu + "='1' and " + vo + "='2'")
  
   
   End If
   
   str = rs.Fields("querystring").Value
   str = Replace(str, "@CLASSIFICACAO_FISCAL", "'" & Inscricao & "'")
   str = Replace(str, "@NRO_LIGACAO", "''")
   rs.Close
   Set rs = ConnSec.execute(str)
   lvLigacoes.ListItems.Clear
   While Not rs.EOF
      Set i = lvLigacoes.ListItems.Add(, , rs(0).Value)
      i.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      i.SubItems(2) = ""
      i.SubItems(3) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
      rs.MoveNext
   Wend
   rs.Close
   dtData(1).Month = Month(Date)
   dtData(1).Year = Year(Date)
   Set rs = Nothing
   Select Case Month(Date)
      Case 1
         dtData(0).Month = 11
         dtData(0).Year = Year(Date) - 1
      Case 2
         dtData(0).Month = 12
         dtData(0).Year = Year(Date) - 1
      Case Else
         dtData(0).Month = Month(Date) - 2
   End Select
   Screen.MousePointer = vbNormal
   Me.Show vbModal
   Exit Sub
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
       Print #1, Now & " - frmConsumoLote - Private Sub Init - " & Err.Number & " - " & Err.Description
       Close #1
       MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
    End If
End Sub

Private Sub chkAgrupar_Click()
   dtData_Click 0
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGrafico_Click()
On Error GoTo Trata_Erro
   
   
    Screen.MousePointer = vbHourglass
    Dim Ligacoes As String, a As Integer
    For a = 1 To lvLigacoes.ListItems.Count
       If lvLigacoes.ListItems(a).Checked Then
          If Ligacoes = "" Then
             Ligacoes = lvLigacoes.ListItems(a).Text
          Else
             Ligacoes = Ligacoes & "," & lvLigacoes.ListItems(a).Text
          End If
       End If
    Next
    Screen.MousePointer = vbNormal
    If Ligacoes <> "" Then
       frmConsumoLoteGraf.Init dtData(0).Month, IIf(dtData(1).Month = 12, 1, dtData(1).Month + 1), dtData(0).Year, IIf(dtData(1).Month = 12, dtData(1).Year + 1, dtData(1).Year), _
                               Inscricao, Ligacoes, IIf(chkAgrupar = 1, True, False)
    End If

Trata_Erro:
    Screen.MousePointer = vbNormal
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
       Print #1, Now & " - frmConsumoLote - Private Sub cmdGrafico_Click - " & Err.Number & " - " & Err.Description
       Close #1
       MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
    End If
End Sub

Private Sub dtData_Change(index As Integer)
   If chkAgrupar.Value = 1 Then
      ResultadoAgrupado
   Else
      ResultadoDesagrupado
   End If

End Sub

Private Sub dtData_Click(index As Integer)
   If chkAgrupar.Value = 1 Then
      ResultadoAgrupado
   Else
      ResultadoDesagrupado
   End If
   
End Sub

Private Sub lvLigacoes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   dtData_Click 0
End Sub

Private Sub ResultadoAgrupado()
On Error GoTo Trata_Erro
   Dim i As ListItem, rs As ADODB.Recordset, str As String, a As Integer, Ligacoes As String

   Dim Total As Double
   Dim Media As Double
   Dim TotalConsumo As Double
   
   Dim mChecked As Boolean

   For a = 1 To lvLigacoes.ListItems.Count
      If lvLigacoes.ListItems(a).Checked Then mChecked = True
   Next
   
   If Not mChecked Then
      'MsgBox "Selecione ao menos um hidrometro", vbExclamation
      lvConsumo.ListItems.Clear
      Exit Sub
   End If
      
    va = """RAMAIS_AGUA_LIGACAO"""
         ve = """querystring"""
         vi = """gs_querys_client"""
         vo = """query_id"""
         vu = """client_id"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
Set rs = Conn.execute("select querystring from gs_querys_client where client_id=1 and query_id=1")
    
   Else
      Set rs = Conn.execute("select " + ve + " from " + vi + " where " + vu + "='1' and " + vo + "='1'")
   
   End If
   
   
   str = rs.Fields("querystring").Value
   
   str = Replace(str, "@CLASSIFICACAO_FISCAL", Inscricao)
   str = Replace(str, "@PERIODO1", dtData(0).Month)
   str = Replace(str, "@PERIODO2", IIf(dtData(1).Month = 12, 1, dtData(1).Month + 1))
   str = Replace(str, "@ANO1", dtData(0).Year)
   str = Replace(str, "@ANO2", IIf(dtData(1).Month = 12, dtData(1).Year + 1, dtData(1).Year))
   
   
   For a = 1 To lvLigacoes.ListItems.Count
      If lvLigacoes.ListItems(a).Checked Then
         If Ligacoes = "" Then
            Ligacoes = lvLigacoes.ListItems(a).Text
         Else
            Ligacoes = Ligacoes & "," & lvLigacoes.ListItems(a).Text
         End If
      End If
   Next
   
   str = Replace(str, "@NRO_LIGACAO", Ligacoes)
   
   rs.Close
   
   
   Set rs = ConnSec.execute(str)
   
   'h.ano,h.PERIODO,sum(h.consumo) as Consumo_Faturado, sum(h.CONSUMO_MEDIDO) as Consumo_Medido,count(*) as qtde_ligacoes
   
   lvConsumo.ColumnHeaders.Clear
   lvConsumo.ColumnHeaders.Add , , "Mes_Ano"
   lvConsumo.ColumnHeaders.Add , , "Cons. Medido"
   lvConsumo.ColumnHeaders.Add , , "Cons. Faturado"
   lvConsumo.ListItems.Clear
   While Not rs.EOF
      Set i = lvConsumo.ListItems.Add(, , rs(0).Value)
      i.SubItems(1) = IIf(IsNull(rs(1).Value), 0, rs(1).Value)
      i.SubItems(2) = CStr(IIf(IsNull(rs(2).Value), 0, rs(2).Value))
      TotalConsumo = TotalConsumo + IIf(IsNull(rs(1).Value), 0, rs(1).Value)
      Total = Total + IIf(IsNull(rs(2).Value), 0, rs(2).Value)
      Media = Media + 1
      rs.MoveNext
   Wend
   
   rs.Close
   
   Set rs = Nothing
   
   Set i = lvConsumo.ListItems.Add(, , "--")
   i.SubItems(1) = "----"
   i.SubItems(2) = "----"
   
   
   Set i = lvConsumo.ListItems.Add(, , "TOTAL")
   i.SubItems(1) = CStr(TotalConsumo)
   i.SubItems(2) = CStr(Total)
   
   Set i = lvConsumo.ListItems.Add(, , "--")
   i.SubItems(1) = "----"
   i.SubItems(2) = "----"
   
   Set i = lvConsumo.ListItems.Add(, , "MEDIA")
   If Media > 0 Then
      i.SubItems(1) = CStr(TotalConsumo / Media)
      i.SubItems(2) = CStr(Total / Media)
   Else
      i.SubItems(1) = CStr(0)
      i.SubItems(2) = CStr(0)
   End If
   Set rs = Nothing
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
       Print #1, Now & " - frmConsumoLote - Private Sub ResultadoAgrupado - " & Err.Number & " - " & Err.Description
       Close #1
       MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
    End If
End Sub

Private Sub ResultadoDesagrupado()
On Error GoTo Trata_Erro
   Dim i As ListItem, rs As ADODB.Recordset, str As String, a As Integer, Ligacoes As String

   Dim Total As Double
   Dim Media As Double
   Dim TotalConsumo As Double
   
   Dim mChecked As Boolean, UltimaLigacao As String

   For a = 1 To lvLigacoes.ListItems.Count
      If lvLigacoes.ListItems(a).Checked Then mChecked = True
   Next
   
   If Not mChecked Then
      lvConsumo.ListItems.Clear
      Exit Sub
   End If
      
   va = """RAMAIS_AGUA_LIGACAO"""
         ve = """querystring"""
         vi = """gs_querys_client"""
         vo = """query_id"""
         vu = """client_id"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
             
   Set rs = Conn.execute("select querystring from gs_querys_client where client_id=1 and query_id=4")
   Else
   
    Set rs = Conn.execute("select " + ve + " from " + vi + " where " + vu + "='1' and " + vo + "='4'")
  
   End If
   
   str = rs.Fields("querystring").Value
   
   str = Replace(str, "@CLASSIFICACAO_FISCAL", Inscricao)
   str = Replace(str, "@PERIODO1", dtData(0).Month)
   str = Replace(str, "@PERIODO2", IIf(dtData(1).Month = 12, 1, dtData(1).Month + 1))
   str = Replace(str, "@ANO1", dtData(0).Year)
   str = Replace(str, "@ANO2", IIf(dtData(1).Month = 12, dtData(1).Year + 1, dtData(1).Year))
   
   
   For a = 1 To lvLigacoes.ListItems.Count
      If lvLigacoes.ListItems(a).Checked Then
         If Ligacoes = "" Then
            Ligacoes = lvLigacoes.ListItems(a).Text
         Else
            Ligacoes = Ligacoes & "," & lvLigacoes.ListItems(a).Text
         End If
      End If
   Next
   
   str = Replace(str, "@NRO_LIGACAO", Ligacoes)
   
   rs.Close
   
   
   Set rs = ConnSec.execute(str)
   
   'l.NRO_LIGACAO, h.ano, h.PERIODO, h.consumo as Consumo_Faturado, h.CONSUMO_MEDIDO as Consumo_Medido
   
   lvConsumo.ColumnHeaders.Clear
   lvConsumo.ColumnHeaders.Add , , "Ligação"
   lvConsumo.ColumnHeaders.Add , , "Mes_Ano"
   lvConsumo.ColumnHeaders.Add , , "Cons. Medido"
   lvConsumo.ColumnHeaders.Add , , "Cons. Faturado"
   lvConsumo.ListItems.Clear
   If rs.EOF Then
      Exit Sub
   Else
      UltimaLigacao = rs(0).Value
   End If
   While Not rs.EOF
      
      If rs(0).Value <> UltimaLigacao Then
         Set i = lvConsumo.ListItems.Add(, , "--")
         i.SubItems(1) = "----"
         i.SubItems(2) = "----"
         i.SubItems(3) = "----"
         
         
         Set i = lvConsumo.ListItems.Add(, , "TOTAL")
         i.SubItems(1) = "-->"
         i.SubItems(2) = CStr(TotalConsumo)
         i.SubItems(3) = CStr(Total)
         
         Set i = lvConsumo.ListItems.Add(, , "--")
         i.SubItems(1) = "----"
         i.SubItems(2) = "----"
         i.SubItems(3) = "----"
         
         Set i = lvConsumo.ListItems.Add(, , "MEDIA")
         i.SubItems(1) = "-->"
         i.SubItems(2) = CStr(TotalConsumo / Media)
         i.SubItems(3) = CStr(Total / Media)
         
         Total = 0
         TotalConsumo = 0
         Media = 0
         
         Set i = lvConsumo.ListItems.Add(, , "")
         
      End If
   
      Set i = lvConsumo.ListItems.Add(, , rs(0).Value)
      i.SubItems(1) = rs(1).Value
      i.SubItems(2) = CStr(IIf(IsNull(rs(2).Value), 0, rs(2).Value))
      i.SubItems(3) = CStr(IIf(IsNull(rs(3).Value), 0, rs(3).Value))
      TotalConsumo = TotalConsumo + IIf(IsNull(rs(2).Value), 0, rs(2).Value)
      Total = Total + IIf(IsNull(rs(3).Value), 0, rs(3).Value)
      Media = Media + 1
      UltimaLigacao = rs(0).Value
      rs.MoveNext
   
   
   Wend
   
   rs.Close
   
   Set i = lvConsumo.ListItems.Add(, , "--")
   i.SubItems(1) = "----"
   i.SubItems(2) = "----"
   i.SubItems(3) = "----"
   
   
   Set i = lvConsumo.ListItems.Add(, , "TOTAL")
   i.SubItems(1) = "-->"
   i.SubItems(2) = CStr(TotalConsumo)
   i.SubItems(3) = CStr(Total)
   
   Set i = lvConsumo.ListItems.Add(, , "--")
   i.SubItems(1) = "----"
   i.SubItems(2) = "----"
   i.SubItems(3) = "----"
   
   Set i = lvConsumo.ListItems.Add(, , "MEDIA")
   i.SubItems(1) = "-->"
   If Media > 0 Then
      i.SubItems(2) = CStr(TotalConsumo / Media)
      i.SubItems(3) = CStr(Total / Media)
   Else
      i.SubItems(2) = CStr(0)
      i.SubItems(3) = CStr(0)
   End If
   Set rs = Nothing
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
       Print #1, Now & " - frmConsumoLote - Private Sub ResultadoDesagrupado - " & Err.Number & " - " & Err.Description
       Close #1
       MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
    End If
End Sub

Private Function GetLayerLot() As String
On Error GoTo Trata_Erro
   Dim rs As ADODB.Recordset
   va = """RAMAIS_AGUA_LIGACAO"""
         ve = """gs_querys_system"""
         vi = """name"""
         vo = """query_id"""
         vu = """client_id"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
   
   Set rs = Conn.execute("Select * from gs_querys_system where name='INSCRICAO_LOTE'")
   Else
    Set rs = Conn.execute("Select * from " + ve + " where " + vi + "='INSCRICAO_LOTE'")
 
   
   End If
   GetLayerLot = rs.Fields("reference_layer").Value
   rs.Close
   Set rs = Nothing
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
       Print #1, Now & " - frmConsumoLote - Private Function GetLayerLot - " & Err.Number & " - " & Err.Description
       Close #1
       MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
    End If
End Function
