VERSION 5.00
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmEPANET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Exportação EPANET"
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
      Caption         =   "Caminho de Exportação"
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
      Caption         =   "Horário"
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
'EpanetExport Versão 06.10.09

Option Explicit
Public conn As ADODB.Connection
Public Provider As Integer
Public PLANO As String

Private rsTP As ADODB.Recordset
Private rsST As ADODB.Recordset

Dim i As Integer

'Declarações necessárias para a função GetMyDocumentsDirectory()
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
'Fim das declarações necessárias para a função GetMyDocumentsDirectory()
'
'
'Rotina inicial da aplicação
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
'Início da função de exportação para o EPANET. Ao final dela será chamado o ModExport pela rotina ExportaEPANet que gera em memória toda a exportação
'para depois gerar em arquivo através de outra rotina. Esta rotina incicia quando o timer é iniciado
Private Function INICIAR()
    On Error GoTo Trata_Erro
    Dim retval As String
    Dim usuario As String
    Dim arquivoLog As String                        'nome do arquivo de log com todas as operações ao exportar para o Epanet
    Dim tipoErro As String                          'indica para o arquivo de log o tipo de erro que pode estar acontecendo
   
    tipoErro = "Erro não localizado"
    arquivoLog = "\Controles\ExportaEpanet" & DateValue(Now) & "  " & TimeValue(Now) & ".log"    'define o nome completo do arquivo de log do sistema, incluíndo a data e hora em que o mesmo será gerado pela primeira vez
    arquivoLog = Replace(arquivoLog, "/", "-")                      'troca caractere / especial que não é aceito como parte do nome do arquivo
    arquivoLog = Replace(arquivoLog, ":", "-")                      'troca caractere : especial que não é aceito como parte do nome do arquivo
    arquivoLog = App.Path & arquivoLog                              'adiciona a localização do caminho onde o aplicativo está instalado
    Open arquivoLog For Append As #5                                'Inicia o log do sistema, abrindo o arquivo sem apagar o log anterior, mantendo sempre o histórico
    Print #5, vbCrLf & "ExportEpanet;*************************************************************************************************"  'Pula uma linha antes de iniciar a escrita
    Print #5, "ExportEpanet;Início do processamento da exportação para o Epanet: " & DateValue(Now) & " - " & TimeValue(Now)
    'Neste arquivo existe gravado o nome do usuário ativo, que indica que usuário exportou para o Epanet o polígono de seleção
    'Através dele será feita uma pesquisa no banco de dados POLIGONO_SELECAO para ver que redes OBJECT_ID_s serão exportados para o EPANET
    'Lembrando que
    '                TIPO = 0 - Nós
    '                TIPO = 1 - Redes
    '                TIPO = 2 - Ramais
    retval = Dir(App.Path & "\Controles\UserLog.txt")
   
    'verifica se o arquivo existe na pasta
    If retval <> "" Then
        'Abre e lê o arquivo para ver que usuário será consultado no polígono selecionado, pois podem existir vários usuários realizando esta operação ao mesmo tempo
        Open App.Path & "\Controles\UserLog.txt" For Input As #3
        Line Input #3, usuario
        Close #3
    Else
        'Avisa e cai fora, pois não dá para executar a operação
        MsgBox "É necessário criar a seleção por polígono.", vbOKOnly + vbInformation, "Mensagem"
        End
    End If
    
    'Liga a ampulheta no ponteiro do mouse
    MousePointer = vbHourglass
    
    'Atualiza todas as rugosidades de todas as tubulações, conforme o tipo de material. Foi considerada uma tubulação de 20 anos de idade
    If conn.Provider <> "PostgreSQL.1" Then
        'Caso o banco de dados seja Oracle ou SQLServer
        'Sempre que alguma linha for alterada na tabela X_Material, estas rugosidades deverão ser revistas
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 111 WHERE MATERIAL = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 111 WHERE MATERIAL = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 1 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 1 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 120 WHERE MATERIAL = 2 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 120 WHERE MATERIAL = 2 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 110 WHERE MATERIAL = 3 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 110 WHERE MATERIAL = 3 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 105 WHERE MATERIAL = 4 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 105 WHERE MATERIAL = 4 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 90 WHERE MATERIAL = 5 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 90 WHERE MATERIAL = 5 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 6 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 6 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 7 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 7 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 8 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 8 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 9 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 9 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 10 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 10 AND ROUGHNESS = 0")
        Print #5, "ExportEpanet;UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 11 AND ROUGHNESS = 0"
        conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 11 AND ROUGHNESS = 0")
    Else
        'Caso seja Postgres
        'Sempre que alguma linha for alterada na tabela X_Material, estas rugosidades deverão ser revistas
        Print #5, "ExportEpanet;UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '0'"""
        conn.Execute ("UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '111' WHERE " + """" + "MATERIAL" + """" + " = '0'"
        conn.Execute ("UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '111' WHERE " + """" + "MATERIAL" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '1' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '1' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '120' WHERE " + """" + "MATERIAL" + """" + " = '2' AND " + """" + "ROUGHNESS" + """" + " = '0' "
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '120' WHERE " + """" + "MATERIAL" + """" + " = '2' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '110' WHERE " + """" + "MATERIAL" + """" + " = '3' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '110' WHERE " + """" + "MATERIAL" + """" + " = '3' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '105' WHERE " + """" + "MATERIAL" + """" + " = '4' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '105' WHERE " + """" + "MATERIAL" + """" + " = '4' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '90' WHERE " + """" + "MATERIAL" + """" + " = '5' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '90' WHERE " + """" + "MATERIAL" + """" + " = '5' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '6' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '6' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '7' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '7' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '8' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '8' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '9' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '9' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '10' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '10' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '11' AND " + """" + "ROUGHNESS" + """" + " = '0'"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '11' AND " + """" + "ROUGHNESS" + """" + " = '0'")
    End If
    'Volta o mouse para o normal
    FrmEPANET.MousePointer = vbDefault
    
    Dim Rs As ADODB.Recordset
    Dim str As String
    Dim Tipo As String
    Dim setor As String
    Dim strtot As String                        'armazena a querie para obter o número total de segmentos de rede que serão exportados para o Epanet (TIPO=1 na tabela POLIGONO_SELECAO)
    Dim totalTrechosExportar As Integer         'número total de trechos de rede de água que serão exportados para o Epanet, disponíveis em (TIPO=1 na tabela POLIGONO_SELECAO)
    
    'Zera todos os materiais de tubulações quando o mesmo não estiver cadastrado
    If conn.Provider <> "PostgreSQL.1" Then
        'Se for Oracle ou SQLServer
        Print #5, "ExportEpanet;UPDATE WATERLINES SET MATERIAL = 0 WHERE MATERIAL IS NULL"
        conn.Execute ("UPDATE WATERLINES SET MATERIAL = 0 WHERE MATERIAL IS NULL")
    Else
        'Se for Postgres
        Print #5, "ExportEpanet;UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "MATERIAL" + """" + " = '0' WHERE " + """" + "MATERIAL" + """" + " IS NULL"
        conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "MATERIAL" + """" + " = '0' WHERE " + """" + "MATERIAL" + """" + " IS NULL")
    End If
    
    'Agora temos que descobrir todos os trechos de rede (TIPO=1 na tabela POLIGONO_SELECAO) que serão exportados para o Epanet, do usuário ativo
    If Provider = 1 Then
        'Se for SQLServer
        str = "SELECT * FROM WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID"
        str = str & " WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & usuario & "' AND TIPO = 1)"
        Print #5, "ExportEpanet;" & str
    ElseIf Provider = 2 Then
        'Se for Oracle
        str = "SELECT * FROM WATERLINES WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID "
        str = str & " WHERE EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE WATERLINES.LINE_ID = P.OBJECT_ID_ AND P.USUARIO = '" & usuario & "' AND P.TIPO = 1)"
        Print #5, "ExportEpanet;" & str
    End If
    If conn.Provider = "PostgreSQL.1" Then
        'Se for Postgres
        str = "SELECT * FROM " + """" + "WATERLINES" + """" + " INNER JOIN " + """" + "X_MATERIAL" + """" + " ON " + """" + "WATERLINES" + """" + "." + """" + "MATERIAL" + """" + " = " + """" + "X_MATERIAL" + """" + "." + """" + "MATERIALID" + """" + " "
        str = str & " WHERE " + """" + "WATERLINES" + """" + "." + """" + "OBJECT_ID_" + """" + " IN (SELECT " + """" + "OBJECT_ID_" + """" + " FROM " + """" + "POLIGONO_SELECAO" + """" + " WHERE " + """" + "USUARIO" + """" + " = '" & usuario & "' AND " + """" + "TIPO" + """" + " = '1')"
        Print #5, "ExportEpanet;" & str
    End If
    
    'Prepara querie para verificar quantos trechos de rede iremos ler, substinuindo o início da querie acima para contar (SELECT COUNT(*))
    If conn.Provider <> "PostgreSQL.1" Then
        'Se for Oracle ou SQLServer
        strtot = Replace(str, "SELECT *", "SELECT COUNT(*)")
        Print #5, "ExportEpanet;" & strtot
    Else
        'Se for Postgres
        strtot = Replace(str, "SELECT *", "SELECT COUNT(*)")
        Print #5, "ExportEpanet;" & strtot
    End If
    Set Rs = New ADODB.Recordset
    'Obtem o número total de trechos de rede a serem exportados para o Epanet os quais são do TIPO=1 na tabela POLIGONO_SELECAO
    If conn.Provider <> "PostgreSQL.1" Then
        'Se SQLServer ou Oracle
        tipoErro = "sql abertura cursor: " & strtot & " - string de conexão: " & conn
        Rs.Open strtot, conn, adOpenDynamic, adLockReadOnly
        Print #5, "ExportEpanet;Executou a abertura do cursor com a querie: " & strtot
        tipoErro = "Erro não localizado"
    Else
        'Se Postgres
        Rs.Open strtot, conn, adOpenDynamic, adLockOptimistic
        Print #5, "ExportEpanet;Executou a abertura do cursor com a querie: " & strtot
    End If
    Me.ProgressBar1.Value = 1
    totalTrechosExportar = Rs(0).Value      'obtem o número total de trechos de rede que serão exportados
    If totalTrechosExportar > 0 Then
        'existe pelo menos um trecho a ser exportado para o Epanet
        Me.ProgressBar1.Max = totalTrechosExportar
    Else
        'não existem trechos a serem exportados para o Epanet
        MsgBox "Não há dados selecionados para exportar.", vbInformation, ""
        Print #5, "ExportEpanet;Total de trechos que serão exportados: " & totalTrechosExportar & " Devido a isso a exportação está sendo abortada."
        Exit Function
    End If
    Print #5, "ExportEpanet;Total de trechos que serão exportados: " & totalTrechosExportar
    Rs.Close
    Set Rs = Nothing
   
    'Agora que temos trechos a serem exportados, vamos exportar para o Epanet
    Set Rs = New ADODB.Recordset
    tipoErro = "sql abertura cursor: " & str & " - string de conexão: " & conn
    Print #5, tipoErro
    Rs.Open str, conn, adOpenDynamic, adLockReadOnly
    conn.CommandTimeout = 300
    tipoErro = "Erro não localizado"
    If Rs.EOF = False Then
        'Fecha temporariamente a conexão com o arquivo de log
        Print #5, "ExportEpanet;Querie a ser enviada para o ExportaEPANet: " & Rs.Source
        Print #5, "ExportEpanet;Fim do processamento inicial da exportação para o Epanet. Iniciará a exportação de " & totalTrechosExportar & " trechos de rede para o Epanet."
        Close #5                                           'Fecha o arquivo de log do sistema
        'Chama rotina de exportação, passando o cursor com a querie com todos os segmentos de rede a serem exportados
        ExportaEPANet Rs, conn, arquivoLog
    Else
        MsgBox "Não há informações selecionadas para exportar.", vbInformation, ""
        Print #5, "ExportEpanet;Fim do processamento inicial da exportação para o Epanet defido a falta de informações para exportar."
        Close #5                                           'Fecha o arquivo de log do sistema
    End If

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #2
        Open App.Path & "\LogErroExportEPANET.txt" For Append As #2
        Print #2, Now & "  - Private Sub cmdConfirmar_Click() - Tipo de erro: " & tipoErro & " - Linha: " & intLinhaCod & " - " & Err.Number & " - " & Err.Description
        Close #2
        Print #5, "ExportEpanet;Fim do processamento da exportação para o Epanet: " & DateValue(Now) & " - " & TimeValue(Now)
        Print #5, "ExportEpanet;*************************************************************************************************"
        Close #5                                           'Fecha o arquivo de log do sistema
        MsgBox "Exportação para o Epanet concluída com não conformidades. Verifique o log no arquivo " & arquivoLog
        MsgBox "Um posssível erro foi identificado na rotina 'INICIAR()':" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo " & App.Path & "\LogErroExportEPANET.txt" & " com informações desta ocorrencia.", vbInformation
    End If
End Function

'Subrotina que inicia o timer e inicia a exportação para o Epanet
'
'
Private Sub Timer1_Timer()
    MousePointer = vbHourglass              'ativa a ampulheta
    INICIAR                                 'inicia a conversão para o EPANET
    MousePointer = vbDefault                'desativa a ampulheta
    Timer1.Enabled = False                  'desativa o timer
    End
End Sub
'Subrotina que irá iniciar a exportação para o Epanet
'
'
Private Sub cmdConfirmar_Click()
    Timer1.Enabled = True               'ativa o timer
    Me.ProgressBar1.Visible = True      'ativa a visualização da barra de progresso
    Me.cmdConfirmar.Enabled = False
End Sub


Private Sub cmdPath_Click()
   cdl.Filter = "Epanet (.inp)|*.INP|Todos tipos (*.*)|*.*|"
   cdl.FileName = txtArquivo.Text
   cdl.InitDir = Environ$("USERPROFILE") & "\my documents"
   cdl.ShowSave
   
   
   txtArquivo.Text = cdl.FileName
End Sub
'Rotina para atualizar rugozidades nas tubulações. Sempre que alguma linha for alterada na tabela X_Material, esta rotina deverá ser revista
'
'
'
Private Sub Command1_Click()
    If MsgBox("Deseja aplicar fórmula Material x Rugosidade?", vbYesNo + vbQuestion, "Confirmar Ação") = vbYes Then
        FrmEPANET.MousePointer = vbHourglass
        If conn.Provider <> "PostgreSQL.1" Then
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 111 WHERE MATERIAL = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 1 AND ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 120 WHERE MATERIAL = 2 AND ROUGHNESS = 0 ")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 110 WHERE MATERIAL = 3 AND ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 105 WHERE MATERIAL = 4 AND ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 90 WHERE MATERIAL = 5 AND ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 6 AND ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 7 AND ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 8 AND ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 9 AND ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 10 AND ROUGHNESS = 0")
            conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 11 AND ROUGHNESS = 0")
        Else
            conn.Execute ("UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '111' WHERE " + """" + "MATERIAL" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '1' AND " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '120' WHERE " + """" + "MATERIAL" + """" + " = '2' AND " + """" + "ROUGHNESS" + """" + " = '0' ")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '110' WHERE " + """" + "MATERIAL" + """" + " = '3' AND " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '105' WHERE " + """" + "MATERIAL" + """" + " = '4' AND " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '90' WHERE " + """" + "MATERIAL" + """" + " = '5' AND " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '6' AND " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '7' AND " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '8' AND " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '9' AND " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '10' AND " + """" + "ROUGHNESS" + """" + " = '0'")
            conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '140' WHERE " + """" + "MATERIAL" + """" + " = '11' AND " + """" + "ROUGHNESS" + """" + " = '0'")
        End If
        FrmEPANET.MousePointer = vbDefault
        MsgBox "Fórmula aplicada com sucesso!", vbInformation, ""
    End If
End Sub
'Obtem o nome do diretório dos Meus Documentos do usuário que está logado
'
'GetMyDocumentsDirectory() - retorna o caminho do diretório
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

