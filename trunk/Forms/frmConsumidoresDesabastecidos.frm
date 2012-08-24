VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConsumidoresDesabastecidos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ligações cadastras nos trechos selecionados"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   5520
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGerartxt 
      Caption         =   "Gerar txt"
      Height          =   315
      Left            =   6390
      TabIndex        =   4
      Top             =   4860
      Width           =   1215
   End
   Begin VB.TextBox txtQde 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   4830
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   7680
      TabIndex        =   1
      Top             =   4860
      Width           =   1155
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4665
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   8229
      SortKey         =   2
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nro Ligação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Endereço"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Usuário"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tel. Res."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tel. Com."
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Total de Ligações:"
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
      Left            =   150
      TabIndex        =   3
      Top             =   4890
      Width           =   1635
   End
End
Attribute VB_Name = "frmConsumidoresDesabastecidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim conexao As New ADODB.connection
        Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
Dim nStr As String
Dim strConn As String
Dim connection As Integer


Public Function init(Object_id_trecho As String) As Boolean
   On Error GoTo Trata_Erro
   'LoozeXP1.InitIDESubClassing
   Dim TABELACOMERCIAL As String
   Dim count2 As Integer
  Dim lig As String
   count2 = 0
   Dim str As String, rs As ADODB.Recordset, itmx As ListItem, QtdeLig As Integer, rs2 As ADODB.Recordset
   Dim fg As String
   Dim fh As String
   Dim fi As String
   Dim fk As String
   Dim fl As String
   Dim fm As String
   fg = "RAMAIS_AGUA"
   fh = "RAMAIS_AGUA_LIGACAO"
   fi = "NRO_LIGACAO"
   fk = "ramais_agua"
   fl = "OBJECT_ID_"
   fm = "OBJECT_ID_TRECHO"
   
   
  ' str = "SELECT l.nro_ligacao from ramais_agua r inner join ramais_agua_ligacao l " & _
   '     "on r.object_id_=l.object_id_ " & _
      '  "where object_id_trecho in(" & Object_id_trecho & ")"
      'alerado em 20/10/2010
       If frmCanvas.TipoConexao = 4 Then
      If connection <> 10 Then
      
       

mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
nStr = frmCanvas.FunDecripta(Senha)
       
strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"

conexao.Open strConn

connection = 10

End If
 End If
 If frmCanvas.TipoConexao <> 4 Then
 str = "SELECT ramais_agua_ligacao.nro_ligacao from ramais_agua_ligacao  inner join ramais_agua on ramais_agua.object_id_=ramais_agua_ligacao.object_id_ where object_id_trecho in(" & Object_id_trecho & ")"
 Set rs = Conn.execute(str)
 Else
    

     str = "SELECT " + """" + fh + """" + "." + """" + fi + """" + " from " + """" + fh + """" + "  inner join " + """" + fg + """" + "  on " + """" + fg + """" + "." + """" + fl + """" + "=" + """" + fh + """" + "." + """" + fl + """" + " where " + """" + fm + """" + " in(" & Object_id_trecho & ")"
     Set rs = conexao.execute(str)
     End If
     
   
   
   'While Not rs.EOF
    '  lig = rs.Fields(0).value
   
     ' rs.MoveNext
   'Wend
   'rs.Close
   
   TABELACOMERCIAL = GetQueryProcess(19)
   If frmCanvas.TipoConexao <> 4 Then
Dim dw As String
dw = "GS_TEMP"
    Set rs2 = Conn.execute("SELECT * FROM GS_TEMP")
   Else
   Set rs2 = conexao.execute("SELECT  * FROM " + """" + "GS_TEMP" + """" + "")
   
   End If
   While Not rs2.EOF
      count2 = 1
      rs2.MoveNext
   Wend
     rs2.Close
     
   If count2 = 1 Then
   
   If frmCanvas.TipoConexao <> 4 Then
   
   
   
   ConnSec.execute "Delete  From " & TABELACOMERCIAL
   Else
      conexao.execute "Delete  From " + """" + TABELACOMERCIAL + """"
   
   End If
   
End If


'Conn.execute ("INSERT INTO GS_TEMP(NRO_LIGACAO) VALUES (" & Object_id_trecho & ")")


   
   
   
   Dim ddd As String
   Dim rsNro_Ligacao As ADODB.Recordset
   Set rsNro_Ligacao = New ADODB.Recordset
   

 If frmCanvas.TipoConexao = 1 Then

  rsNro_Ligacao.Open TABELACOMERCIAL, ConnSec, adOpenKeyset, adLockOptimistic, adCmdTable
  
  
 ElseIf frmCanvas.TipoConexao = 2 Then

   ddd = "SELECT  * FROM GS_TEMP"
    rsNro_Ligacao.Open ddd, ConnSec, adOpenDynamic, adLockOptimistic
  
  
  
  Else
  
   ddd = "SELECT  * FROM " + """" + "GS_TEMP" + """" + ""
    rsNro_Ligacao.Open ddd, conexao, adOpenDynamic, adLockOptimistic
 End If
   
   While Not rs.EOF
      rsNro_Ligacao.AddNew
      rsNro_Ligacao.Fields(0).value = rs.Fields(0).value
      rsNro_Ligacao.Update
      rs.MoveNext
   Wend
 '  rs.Close
 

    'Lv.ListItems.Clear
    
     ' Set itmx = Lv.ListItems.Add(, , 0)
    '  itmx.SubItems(1) = 0
    '  itmx.SubItems(2) = 0
    '  itmx.SubItems(3) = 0
    '  itmx.SubItems(4) = 0
      
   '   txtQde = QtdeLig
  ' Me.Show vbModal
  ' 'LoozeXP1.EndWinXPCSubClassing
    
    
    
    
    
    
      
   str = GetQueryProcess(18)
   If frmCanvas.TipoConexao <> 4 Then
   Set rs = ConnSec.execute(str)
   Else
   Set rs = conexao.execute(str)
   
   End If
   'Set rs = ConnSec.execute("SELECT LI.NRO_LIGACAO, (LI.ENDERECO + '-' + ISNULL(LI.NUM_CASA,'') + '-' +  ISNULL(LI.COMPL_LOGRADOURO,'') + '-' + ISNULL(LI.BAIRRO,'')) as Endereco, LI.CONSUMIDOR, LI.TEL_RES AS TELRES, LI.TEL_COM AS TELCOM FROM NXGS_V_LIG_COMERCIAL LI INNER JOIN gs_temp G ON G.NRO_LIGACAO = LI.NRO_LIGACAO")
   
 If frmCanvas.TipoConexao <> 4 Then
   Lv.ListItems.Clear
   While Not rs.EOF
      Set itmx = Lv.ListItems.Add(, , rs.Fields(0).value)
      itmx.SubItems(1) = rs.Fields(1).value
      itmx.SubItems(2) = rs.Fields(2).value
      itmx.SubItems(3) = rs.Fields(3).value
      itmx.SubItems(4) = rs.Fields(4).value

      QtdeLig = QtdeLig + 1
     rs.MoveNext
   Wend
   rs.Close
   Else
   
    Lv.ListItems.Clear
   While Not rs.EOF
      Set itmx = Lv.ListItems.Add(, , rs.Fields(0).value)
      itmx.SubItems(1) = rs.Fields(1).value
      itmx.SubItems(2) = rs.Fields(5).value
      itmx.SubItems(3) = rs.Fields(6).value
      itmx.SubItems(4) = rs.Fields(7).value

      QtdeLig = QtdeLig + 1
     rs.MoveNext
   Wend
   rs.Close
   
   End If
   
   
   
   
   
   

txtQde = QtdeLig
   Me.Show vbModal
   'LoozeXP1.EndWinXPCSubClassing
   
   

   Exit Function
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Or Err.Number = 94 Then
       Resume Next
   Else
   
   
         'LoozeXP1.EndWinXPCSubClassing
    

   
      
      'PrintErro CStr(Me.Name), "Public Function Init", CStr(Err.Number), CStr(Err.Description), True
      
   End If
    
End Function

Private Sub cmdGerartxt_Click()
On Error GoTo Trata_Erro
   Dim a As Integer
   cdl1.FileName = ""
   cdl1.ShowSave
   If cdl1.FileName <> "" Then
      Open cdl1.FileName For Output As #1
      For a = 1 To Lv.ListItems.count
         Print #1, Lv.ListItems.Item(a).Text & ";" & _
                     Lv.ListItems.Item(a).SubItems(1) & ";" & _
                     Lv.ListItems.Item(a).SubItems(2) & ";" & _
                     Lv.ListItems.Item(a).SubItems(3) & ";" & _
                     Lv.ListItems.Item(a).SubItems(4)
      Next
      Close #1
      
      MsgBox "Gerado com sucesso", vbInformation
      Shell "notepad.exe " & cdl1.FileName, vbNormalFocus
   End If
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
       PrintErro CStr(Me.Name), "Private Sub cmdGerartxt_Click", CStr(Err.Number), CStr(Err.Description), True
   End If
End Sub

Private Sub cmdOK_Click()
   Unload Me
End Sub



