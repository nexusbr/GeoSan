Attribute VB_Name = "Global"

Option Explicit

Public Const VK_ESCAPE = &H1B                                                              'definie a tecla ESC para eventos de interrupção do programa
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer      'para habilitar o timer e poder interromper tarefas que demoram muito

Public varGlobais As New CVariaveis             'variáveis globais que todas as rotinas podem acessar
Public cGeoDatabase As New cGeoDatabase         'conexão do TeDatabase única para toda a aplicação
Public cGeoViewDatabase As New CViewDatabase    'conexão com o TeViewManager para toda a aplicação
Public ErroUsuario As New CPrintErro            'classe responsável por apresentar caixa de diálogo de erro e registrar o erro no arquivo de log
Public Email As New CEmail                      'Classe responsável pelo envio de emails
Public arquivo As New CArquivo                  'Classe de operação de arquivos e diretórios
Public Type Ramais                              'utilizado para mover os ramais quando um nó de um trecho de rede é movido
    objIdTrecho As String
    objIdRamal As String
    geomIdRamal As String
    Distancia As Double
    comprTrecho As Double
    xHidrom As Double
    yHidrom As Double
End Type
Public ramalMovendo() As Ramais
Private AbrirArquivo As New clsAbreArquivo      'Classe que abre um arquivo conforme a extensão do mesmo
Public Versao_Geo As String                     'Número da versão do software no formato XX.YY.ZZ.WWDim exp As New GeosanExport
Public exp As New GeosanExport                  'médotos de exportação para o formato shape
'FUNÇÕES PARA LER E GRAVAR NO ARQUIVO .INI-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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
Private Type Usuarios
   UsrId As Long
   UseName As String
End Type
' typeconnection - represents the database type
' access = 0
' sqlserver = 1
' oracle = 2
' firebird = 3
' postgres = 4

Public usuario As Usuarios, typeconnection As cAppType
Public strUser As String 'VARIAVEL GLOBAL DE USUÁRIO LOGADO
Public xWorld As Double
Public yWorld As Double
Public canvasScale As Double
Public Object_id_Show As String
Public blnLocalizandoConsumidor As Boolean
Public strLayerAtivo As String
Public idPoligonSel As String
Public ramal_Object_id_trecho As Long
Public idAutoLote As String 'CODIGO DO LOTE QUE é UTILIZADO NO CADASTRO DO RAMAL
'MEDIR DISTANCIA ENTRE DOIS PONTOS
Public X1i As Double
Public Y1i As Double
Public X1 As Double
Public Y1 As Double
Public XYInicio As Boolean
Public CanvasXmin_ As Double
Public CanvasYmin_ As Double
Public CanvasXmax_ As Double
Public CanvasYmax_ As Double
Public strViewAtiva_ As String
Public blnGeraRel As Boolean
Type TListaNo
    indice As Integer
    object_id As String
    X As Double
    Y As Double
End Type
Public Enum nxSqlOperations
   nxSELECT = 1
   nxUpdate = 2
   nxInsert = 3
   nxDelete = 4
End Enum
Public blnMonitorar As Boolean
'Public frmCanvas.TipoConexao As Integer
Public ConnPostgresPorta As String ' informa ao a conexão TeConnectio a porta de conexão postgres
Public blnAutoLogin As Boolean
Public dblFatorZoomMais As Double
Public dblFatorZoomMenos As Double
Public Conn As New ADODB.connection

Public stopProcess As Boolean
Public Sec As New NSecurity.AppMode
Public nxUser As New NexusUsers.clsUsers
Public ConnSec As New ADODB.connection
Public msEmpresa As String
Public msconexao As String
Public msServidor As String
Public msBanco As String
Public msPathFileName As String
Public msServiceName As String
Public MyConn As ADODB.connection
Public idLinhaEmDesenho As String
'Variáveis para operações com polígonos
'IDENTIFICAÇÃO PARA SABER SE ESTÁ SENDO ANALISADO UM POLÍGONO VIRTUAL OU POLÍGONO DE LAYER
Public blnPoligonoVirtual As Boolean
Public lngTotalRedesDentro As Long
Public lngTotalRedesDivisa As Long
Public lngTotalRamaisDentro As Long
Public lngTotalRamaisDivisa As Long
Public lngTotalPontosDentro As Long
Public lngTotalPontosDivisa As Long
Public ArrRedesDentro() As Long
Public ArrRedesDivisa() As Long
Public ArrRamaisDentro() As Long
Public ArrRamaisDivisa() As Long
Public ArrPontosDentro() As Long
Public ArrPontosDivisa() As Long
'**************************************************************************************
Private aListaNo() As TListaNo
Public Enum mSaveType
   Single_Point = 0
   Single_Line = 1
   Multiples_Point = 2
   Multiples_Line = 3
End Enum
Public Enum TypeGeometry
   Polyguns = 1         'TePOLYGONS
   points = 4           'TePOINTS
   lines = 2            'TeLINES
   texts = 128          'TeTEXT
End Enum
Enum TipoRelatorio
   RedeMaterialDiametro = 0
   RegistrosEstadoEstado = 1
   ComponentsRede = 2
End Enum

'The GetDC function retrieves a handle of a display device context (DC) for the client area of the specified window.
'The display device context can be used in subsequent GDI functions to draw in the client area of the window.
'Utilizada na função para converter Twits para Pixels
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
  
'Returns pixels per inch
'Utilizada na função para converter Twits para Pixels
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
' Subrotina principal de entrada do GeoSan
'
' Sempre ao gerar um novo executável atualizar o número da versão do GeoSan a seguir
'
'
Public Sub Main()
    On Error GoTo Trata_Erro
    Dim momento As String
    Dim nC As New NexusConnection.App               'ConectaBanco
    Dim contador As String * 100
    Dim tipo As String
    Dim connn As String
    Dim rs As ADODB.Recordset
    Dim retval As String
    Dim frmauto As New frmAutoLogin
    Dim stringconexao As String
    Dim strBanco As String
    Dim s As String
    Dim strMais As String
    Dim strMenos As String
    Dim gsParameters As New clsGS_Parameters
    Dim atualizaAplicacao As CAtualiza                                                              'para poder atualizar GeoSanIni.exe se necessário, assim da próxima vez que entrar no GeoSan ele está com a versão atualizada
    Dim retornoAtualizaAplicacao As Boolean                                                         'para saber se atualizou ou não com sucesso a aplicação GeoSanIni.exe
    
    If UCase(ReadINI("EMAIL", "ENVIARMENSAGENS", App.path & "\CONTROLES\GEOSAN.INI")) = "SIM" Then  'informa todo o sistema de é ou não para enviar por email as mensagens de erro que ocorrem
        glo.enviaEmails = True                                                                      'sempre enviar emails de erros. Salva globalmente
    Else
        glo.enviaEmails = False                                                                     'nunca enviar emails de erros
    End If
    'Configura a versão atual do GeoSan
    Versao_Geo = App.Major & "." & App.Minor & "." & App.Revision
    Versao_Geo = "08.01.00"
    glo.diretorioGeoSan = App.path                                                                  'salva globalmente o caminho onde encontra-se o GeoSan.exe
    SaveLoadGlobalData glo.diretorioGeoSan + "/controles/variaveisGlobais.txt", True                'salva em um arquivo todas as variáveis globais para poderem ser acessadas por outras aplicações
    connn = ""
    If Not nC.appGetRegistry(App.EXEName, Conn, typeconnection) Then
        If Not nC.appNewRegistry(App.EXEName, Conn, typeconnection) Then
            End
        End If
        typeconnection = nC.typeconnection
    End If
    Set nC = Nothing
    FrmMain.Show
    Set rs = New ADODB.Recordset
    '%%%% AUTO LOGIN %%%%
    retval = Dir(App.path & "\Controles\AutoLogin.txt")
    If retval <> "" Then 'verifica se o arquivo existe na pasta
        blnAutoLogin = True
        Open App.path & "\Controles\AutoLogin.txt" For Input As #3
        Input #3, strUser
        Close #3
        If Trim(strUser) = "" Then
            MsgBox "Arquivo de login automático inválido.", vbExclamation, ""
            Kill App.path & "\Controles\AutoLogin.txt"
            End
        End If
        a = "USRLOG"
        c = "SYSTEMUSERS"
        b = "USRFUN"
        'manoel alterou em 18/10/2010
        If frmCanvas.TipoConexao <> 4 Then
            rs.Open ("SELECT * FROM SYSTEMUSERS WHERE USRLOG = '" & strUser & "'"), Conn, adOpenDynamic, adLockReadOnly
            If rs.EOF = False Then
                If Sec.MyUsers.SelectData(Conn, rs!UsrId) Then
                    usuario.UseName = Sec.MyUsers.UsrLog
                End If
            End If
        Else
            rs.Open ("SELECT * FROM " + """" + c + """" + " WHERE " + """" + a + """" + " = '" & strUser & "'"), Conn, adOpenDynamic, adLockOptimistic
            If rs.EOF = False Then
                If Sec.MyUsers.SelectData(Conn, rs!UsrId) Then
                    usuario.UseName = Sec.MyUsers.UsrLog
                End If
            End If
        End If
        frmauto.Show 1
    Else 'O arquivo não existe na pasta
        blnAutoLogin = False
        nxUser.TipoConexao (frmCanvas.TipoConexao)
        usuario.UsrId = Sec.OpenLogin(Conn)                                 'Abre a tela de Usuário e Senha para preenchimento
        If Sec.MyUsers.SelectData(Conn, usuario.UsrId) Then
            usuario.UseName = Sec.MyUsers.UsrLog
            strUser = Sec.MyUsers.UsrLog
        Else
            Set Sec = Nothing
            End
        End If
    End If
    'Valida o perfil do usuário -
    Set rs = New ADODB.Recordset
    a = "USRLOG"
    c = "SYSTEMUSERS"
    b = "USRFUN"
    If frmCanvas.TipoConexao <> 4 Then
        stringconexao = "SELECT USRLOG, USRFUN FROM SYSTEMUSERS WHERE USRLOG = '" & strUser & "' ORDER BY USRLOG"
    Else
        stringconexao = "Select " + """" + a + """" + "," + """" + b + """" + " from  " + """" + c + """" + "Where " + """" + a + """" + "=" + " '" & strUser & "' ORDER BY " + """" + a + """" + ""
    End If
    rs.Open stringconexao, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        If blnAutoLogin = True Then
            If rs!UsrFun < 3 Then 'para logina automático somente pode ser usuário tipo visitante ou visualizador
                MsgBox "Este usuário não pode iniciar com login automático." & Chr(13) & Chr(13) & "Senha requerida.", vbExclamation, ""
                Kill App.path & "\Controles\AutoLogin.txt"
                End
            End If
        End If
        If rs!UsrFun = 1 Then 'ADMINISTRADOR
            FrmMain.mnuChangePassword.Visible = False 'desabilita a troca de senha pelo menu arquivo, pois o admin. pode fazer isso por outro menu
            FrmMain.mnuAutoLogin.Visible = False
            FrmMain.mnuCalculaZNo = True                                        'Exibe a opção de o usuário selecionar se deseja ou não que as cotas sejam calculadas enquanto ele desenha uma rede
            FrmMain.mnuAtualizaCotas.Visible = True                             'Permite atualizar todas as cotas de todos os nós das redes da cidade toda
        ElseIf rs!UsrFun = 2 Then                                               'USUÁRIO
            FrmMain.mnuUsers.Enabled = False                                    'NÃO PERMITE QUE SEJAM EDITADOS USUÁRIOS
            FrmMain.mnuProdutividade.Enabled = False                            'NÃO PERMITE GERAR RELATORIO DE PRODUTIVIDADE
            FrmMain.mnuUpdate_Demand.Visible = False                            'NÃO PERMITE ATALIZAÇÃO DE DEMANDA
            FrmMain.mnuAutoLogin.Visible = False                                'Não permite o login automático
            FrmMain.mnuExporta_GeoSan.Visible = False                           'Não permite exportar para o formato shape
            FrmMain.mnuAtualizaCotas.Visible = False                                    'não permite atualizar todas as cotas de todos os nós das redes da cidade toda
            FrmMain.mnuCalculaZNo = True                                        'Exibe a opção de o usuário selecionar se deseja ou não que as cotas sejam calculadas enquanto ele desenha uma rede
        ElseIf rs!UsrFun = 3 Then 'VISITANTE - BLOQUEIA A MAIORIA DAS FUNÇÕES
            blnAutoLogin = True
            '          If blnAutoLogin = True Then              'CASO LOGIN AUTOMÁTICO, NÃO PERMITE ALTERAR VISTAS
            '               FrmMain.pctSfondo.Visible = False   'PARA ALTERAR VISTAS DEVE SE ENTRAR COM O USUÁRIO
            '               FrmMain.mnuLayers.Visible = False   'E SENHA NO MODO CONVENCIONAL
            '          End If
            FrmMain.mnuCalculaZNo = False                                        'Não exibe a opção de o usuário selecionar se deseja ou não que as cotas sejam calculadas enquanto ele desenha uma rede
            FrmMain.mnuDrawLineWater.Visible = False
            FrmMain.mnuDrawPointInLineWater.Visible = False
            FrmMain.mnuMovePointWithLines.Visible = False
            'FrmMain.mnuInsertDocs.Visible = False
            FrmMain.mnuDeleteLineWater.Visible = False
            FrmMain.mnuDrawRamal.Visible = False
            'FrmMain.mnuInsertLabel.Visible = False
            FrmMain.mnuCadastros.Visible = False
            FrmMain.mnuAdmin.Visible = False
            FrmMain.mnuProdutividade.Visible = False
            FrmMain.mnuCarregaPoligono.Visible = False
            FrmMain.mnuUpdate_Demand.Visible = False
            FrmMain.mnuImport.Visible = False
            FrmMain.mnuEditBar30.Visible = False
            FrmMain.mnuEditBar80.Visible = False
            FrmMain.mnusep1234.Visible = False
            FrmMain.mnusep9999.Visible = False
            FrmMain.tbToolBar.Buttons("ksave").Visible = False
            FrmMain.tbToolBar.Buttons("kdrawnetworkline").Visible = False
            FrmMain.tbToolBar.Buttons("kmovenetworknode").Visible = False
            FrmMain.tbToolBar.Buttons("kinsertnetworknode").Visible = False
            'FrmMain.tbToolBar.Buttons("kinsertdoc").Visible = False
            FrmMain.tbToolBar.Buttons("kdelete").Visible = False
            FrmMain.tbToolBar.Buttons("kdrawramal").Visible = False
            FrmMain.tbToolBar.Buttons("kdrawramalAuto").Visible = False
            FrmMain.tbToolBar.Buttons("kdrawramalAddConsumer").Visible = False
            FrmMain.tbToolBar.Buttons("mnuPoligono").Visible = False
            FrmMain.tbToolBar.Buttons("kdelete").Visible = False
            FrmMain.tbToolBar.Buttons("ksearchinnetwork").Visible = False
            FrmMain.tbToolBar.Buttons("kMoveConsumidorGPS").Visible = False
            'FrmMain.tbToolBar.Buttons("kdeclivity").Visible = False
            FrmMain.mnuExporta_GeoSan.Visible = False                           'Não permite exportar para o formato shape
            FrmMain.mnuAtualizaCotas.Visible = False                                    'não permite atualizar todas as cotas de todos os nós das redes da cidade toda
        ElseIf rs!UsrFun = 4 Then                                               'VISUALIZADOR - BLOQUEIA A MAIORIA DAS FUNÇÕES
            blnAutoLogin = True
            FrmMain.mnuCalculaZNo = False                                       'Não exibe a opção de o usuário selecionar se deseja ou não que as cotas sejam calculadas enquanto ele desenha uma rede
            FrmMain.pctSfondo.Visible = False
            FrmMain.mnuLayers.Checked = False
            FrmMain.mnuExpAutoCad.Visible = False
            FrmMain.mnuExportLocalNos.Visible = False
            FrmMain.mnusep01001.Visible = False
            FrmMain.mnusep011101.Visible = False
            FrmMain.mnuChangePassword.Visible = False
            FrmMain.mnuFixaIcone.Visible = False
            FrmMain.mnuRel.Visible = False
            FrmMain.mnuFileBar2.Visible = False
            'FrmMain.mnuFilePrint.Visible = False
            FrmMain.mnu_Find_Object.Visible = False
            FrmMain.mnuDrawLineWater.Visible = False
            FrmMain.mnuDrawPointInLineWater.Visible = False
            FrmMain.mnuMovePointWithLines.Visible = False
            'FrmMain.mnuInsertDocs.Visible = False
            FrmMain.mnuDeleteLineWater.Visible = False
            FrmMain.mnuDrawRamal.Visible = False
            ' FrmMain.mnuInsertLabel.Visible = False
            FrmMain.mnuCadastros.Visible = False
            FrmMain.mnuAdmin.Visible = False
            FrmMain.mnuProdutividade.Visible = False
            FrmMain.mnuCarregaPoligono.Visible = False
            FrmMain.mnuUpdate_Demand.Visible = False
            FrmMain.mnuImport.Visible = False
            FrmMain.mnuEditBar30.Visible = False
            FrmMain.mnuEditBar80.Visible = False
            FrmMain.mnusep1234.Visible = False
            FrmMain.mnusep9999.Visible = False
            FrmMain.tbToolBar.Buttons("ksave").Visible = False
            FrmMain.tbToolBar.Buttons("kdrawnetworkline").Visible = False
            FrmMain.tbToolBar.Buttons("kmovenetworknode").Visible = False
            FrmMain.tbToolBar.Buttons("kinsertnetworknode").Visible = False
            'FrmMain.tbToolBar.Buttons("kinsertdoc").Visible = False
            FrmMain.tbToolBar.Buttons("kdelete").Visible = False
            FrmMain.tbToolBar.Buttons("kdrawramal").Visible = False
            FrmMain.tbToolBar.Buttons("kdrawramalAuto").Visible = False
            FrmMain.tbToolBar.Buttons("kdrawramalAddConsumer").Visible = False
            FrmMain.tbToolBar.Buttons("kMoveConsumidorGPS").Visible = False
            FrmMain.tbToolBar.Buttons("mnuPoligono").Visible = False
            FrmMain.tbToolBar.Buttons("kdelete").Visible = False
            FrmMain.tbToolBar.Buttons("ksearchinnetwork").Visible = False
            'FrmMain.tbToolBar.Buttons("kdeclivity").Visible = False
            FrmMain.mnuExporta_GeoSan.Visible = False                           'Não permite exportar para o formato shape
            FrmMain.mnuAtualizaCotas.Visible = False                                    'não permite atualizar todas as cotas de todos os nós das redes da cidade toda
        Else
            MsgBox "Não foi encontrada a permissão para este usuário.", vbExclamation, ""
            rs.Close
            End
        End If
        rs.Close
    Else
        MsgBox "Usuário não cadastrado.", vbExclamation, ""
        rs.Close
        End
    End If
    If frmCanvas.TipoConexao <> 4 Then
        If UCase(ReadINI("MAPA", "CORRIGIR_QUADRANTE", App.path & "\CONTROLES\GEOSAN.INI")) = "SIM" Then
            CorrigeQuadrante
        End If
    Else
        If UCase(ReadINI("MAPA", "CORRIGIR_QUADRANTE", App.path & "\CONTROLES\GEOSAN.INI")) = "SIM" Then
            CorrigeQuadrante
        End If
    End If
    s = mid(ReadINI("CONEXAO", "PROVEDOR", App.path & "\CONTROLES\GEOSAN.ini"), 1, 1)
    If Trim(s) = "" Or IsNumeric(s) = False Then
        MsgBox "Informação de tipo de conexão inválida. (Geosan.ini)", vbCritical, ""
        End
    Else
        'frmCanvas.TipoConexao = s
        ' VERIFICA QUAL É O TIPO DO BANCO E FAZ A VERIFICAÇÃO DE VERSÃO DO BANCO DE DADOS
        ' CASO POSTGRES, IDENTIFICA A PORTA DE CONEXÃO
        ' Select Case frmCanvas.TipoConexao
        ' Case 1 ' SqlServer
        'VerificaBaseSQL
        ' Case 2 ' Oracle
        'VerificaBaseORACLE
        ' Case 4 ' PostgreSQL
        ' VerificaBasePOSTGRES
        'ConnPostgresPorta = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
        'If ConnPostgresPorta = "" Or IsNumeric(ConnPostgresPorta) = False Then
        ' MsgBox "Informação de porta de conexão inválida. (Geosan.ini)", vbCritical, ""
        '  End
        'End If
        'End Select
    End If
    strBanco = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
    FrmMain.Caption = "NEXUS - GeoSan " & Versao_Geo & " [Banco: " & strBanco & "]"
    'CARREGA AS CONFIGURAÇÕES DO GEOSAN.INI NO ZOOM DO USUÁRIO NA MÁQUINA
    strMais = Replace(ReadINI("MAPA", "ZOOM_MAIS", App.path & "\CONTROLES\GEOSAN.ini"), ",", ".")
    strMenos = Replace(ReadINI("MAPA", "ZOOM_MENOS", App.path & "\CONTROLES\GEOSAN.ini"), ",", ".")
    If IsNumeric(strMais) = True And IsNumeric(strMenos) = True Then
        dblFatorZoomMais = strMais
        dblFatorZoomMenos = strMenos
    Else
        dblFatorZoomMais = 2
        dblFatorZoomMenos = 2
    End If
    momento = "ConnComercial"
    'Estabelece objeto de conexão para o Comercial
    If gsParameters.getData(Conn, tipo) Then
        If UCase(gsParameters.String_Connection_Secundary) = UCase(Conn.ConnectionString) Then
            'Seta conexão comercial sendo a mesma do GeoSan
            Set ConnSec = Conn
        Else
            'Seta conexão comercial outro banco
            Set rs = New ADODB.Recordset
            a = """STRING_CONNECTION_SECUNDARY"""
            c = """GS_PARAMETERS"""
            If frmCanvas.TipoConexao <> 4 Then
                stringconexao = "SELECT STRING_CONNECTION_SECUNDARY FROM GS_PARAMETERS"
            Else
                stringconexao = "Select " + a + "  from  " + c + ""
            End If
            rs.Open (stringconexao), Conn, adOpenDynamic, adLockReadOnly
            If rs.EOF = False Then
                strBanco = rs!String_Connection_Secundary
                ConnSec.Open gsParameters.String_Connection_Secundary
            Else
                MsgBox "Não há string de conexão para o banco de dados comercial", vbInformation, ""
            End If
            rs.Close
            'MsgBox "Conexão com banco comercial apontando para outro banco", vbInformation, "Conexão comercial"
        End If
    End If
    exp.AtivaRamaisGeoSan                                                   'precisa ativar o te_representation, uma vez que na exportação que pode ter ocorrido ou ter sido cancelada, pode ter sido apagado
    cGeoDatabase.configura Conn, typeconnection, usuario.UseName            'aqui ele inicializa a conexão com o banco de dados, com TeDatabase, para fazer todas as operações necessárias ao longo de toda a aplicação
    Set atualizaAplicacao = New CAtualiza                                   'para atualizar o GeoSanIni.exe se necessário
    'retornoAtualizaAplicacao = atualizaAplicacao.AtualizaAplicacaoLocal     'retorna verdadeiro se atualizou com sucesso, falso se houve uma falha em localização de algum arquivo
    
pulaConexaoComercial:
    momento = ""
    Set gsParameters = Nothing
    Set Sec = Nothing
    Set nxUser = Nothing
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Or Err.Number = 374 Then
        Resume Next
    ElseIf momento = "ConnComercial" Then
        MsgBox "A conexão com o banco de dados comercial não pode ser estabelecida pelo seguinte motivo: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Verifique a tabela GS_PARAMETERS, pois ela não está apontando corretamente para a vista do banco comercial.", vbInformation, "Conexão comercial"
        GoTo pulaConexaoComercial
    Else
        ErroUsuario.Registra "Global", "Main", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
' Lê o arquivo de inicialização do GeoSan e retorna o parâmetro solicitado do mesmo
'
' Secao - o nome da seção presente no arquivo .ini entre colchetes []
' Entrada - nome do parâmetro ao qual se deseja obter a informação de entrada, o qual está dentro da seção apontada. Fica antes do sinal de igual
' arquivo - nome do arquivo .ini que será lido
'
Public Function ReadINI(Secao As String, Entrada As String, arquivo As String)
    Dim retlen As String
    Dim Ret As String
    
    Ret = String$(255, 0)                                                           'string que conterá o parâmetro de retorno. Preenche com o caractere ASCII 0 255 vezes
    retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), arquivo)
    Ret = Left$(Ret, retlen)
    ReadINI = Ret
End Function

Public Sub WriteINI(Secao As String, Entrada As String, Texto As String, arquivo As String)
  
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
  'texto= valor que vem depois do igual
  
  WritePrivateProfileString Secao, Entrada, Texto, arquivo

End Sub

Public Function Imprima(str As String) As Boolean

   Open "c:\GeoPrint.txt" For Append As #3
   Print #3, str
   Close #3

End Function
' Retorna em uma String o nome das colunas que foram retornadas num select
'
' strSelect - querie sql
' strSelect2 - querie sql
'
Public Function RetornaCabecalho(ByVal strSelect As String, ByVal strSelect2 As String) As String
    On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim nomeCols As String
    Dim strSel As String
    Dim strSel3 As String
    Dim strSel4 As String
    Dim i As Integer
    
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    strSel4 = strSelect2
    rs2.Open strSel4, Conn, adOpenDynamic, adLockOptimistic
    If rs2.EOF = False Then
        strSel4 = rs2(0).value
    End If
    rs.Open strSelect, Conn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = False Then
        strSel = rs!querystring
        Set rs = New ADODB.Recordset
        strSel3 = strSel + " " + "'" + strUser + "'" + " " + strSel4            'adiciona a querie existente em GS_QUERYS_CLIENT de 22 + usuário logado + 23 formando uma única querie
        rs.Open strSel3, Conn, adOpenDynamic, adLockOptimistic
        'monta a string de colunas, obtendo o nome de todas as colunas existentes na querie cocactenada 22 + usuário logado + 23
        nomeCols = rs.Fields(0).Name
        For i = 1 To rs.Fields.count - 1
            nomeCols = nomeCols & ";" & rs.Fields(i).Name
        Next
        RetornaCabecalho = nomeCols
    End If
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "Global", "RetornaCabecalho", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function
' Gera um relatório e salva em arquivo texto a partir de um nome do arquivo a salvar e uma querie SQL
'
' strArqDestino - nome do arquivo destino onde será salvo o relatório
' strSelect - querie do relatório
'
Public Function PrintSelect(ByVal strArqDestino As String, ByVal strSelect As String) As Boolean
    On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset
    Dim nomeCols As String
    Dim i As Long
    Dim j As Long
    Dim dbvetor As Variant                                                              'vetor com todos os nomes das colunas que são retornadas pela querie
    Dim colunas As Integer
    Dim registros As Long
    Dim linha As String
    Dim numeroDeColunas As Integer                                                      'número total de colunas retornadas na querie
    
    Screen.MousePointer = vbHourglass                                                   'mostra a ampulheta para o usuário
    If Trim(strArqDestino) = "" Then
        MsgBox "Não há caminho de arquivo para gerar o relatório.", vbInformation, ""
        PrintSelect = False
        Exit Function
    End If
    If Trim(strSelect) = "" Then
        MsgBox "Não há um script SQL definido para gerar o relatório.", vbInformation, ""
        PrintSelect = False
        Exit Function
    End If
    Set rs = New ADODB.Recordset
    rs.Open strSelect, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        Open strArqDestino For Output As #1
        numeroDeColunas = rs.Fields.count                                               'obtem o número total de colunas
        'monta a string de colunas para imprimir no cabeçalho
        nomeCols = rs.Fields(0).Name
        For i = 1 To numeroDeColunas - 1
            nomeCols = nomeCols & ";" & rs.Fields(i).Name
        Next
        'obtenho o número de colunas e o número de linhas
        dbvetor = rs.GetRows
        colunas = UBound(dbvetor, 1)
        registros = UBound(dbvetor, 2)
        'imprime o cabeçalho
        Print #1, nomeCols
        For i = 0 To registros
            For j = 0 To colunas
                linha = linha & dbvetor(j, i) & ";"
            Next j
            Print #1, linha
            linha = ""
        Next i
        Close #1
        PrintSelect = True
    Else
        MsgBox "Não existe informação para gerar o relatório.", vbInformation, ""
    End If
    rs.Close
    Screen.MousePointer = vbNormal
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Screen.MousePointer = vbNormal
        ErroUsuario.Registra "Global", "PrintSelect", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function
'Acrescenta ao arquivo de log o erro ocorrido
'
'Modulo - string que contém em arquivo VB o erro ocorreu
'EVENTO - string que contém em que rotina o erro ocorreu
'ErrDescr - string com a descrição do erro ocorrido
'ExibeMensagem - se é para exibir ou não uma mensagem para o usuário
'linha - número da linha em que o erro ocorreu
'
Public Function PrintErro(ByVal Modulo As String, ByVal EVENTO As String, ByVal ErrNum As String, ByVal ErrDescr As String, ByVal ExibeMensagem As Boolean, Optional ByVal linha As Integer = 0)
      Close #1 'FECHA O ARQUIVO DE LOG
      Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
      Print #1, "DATA"; Tab(16); Now
      Print #1, "USUÁRIO"; Tab(16); strUser
      Print #1, "VERSÃO"; Tab(16); Versao_Geo
      Print #1, "MÓDULO"; Tab(16); Modulo
      Print #1, "EVENTO"; Tab(16); EVENTO
      Print #1, "LINHA"; Tab(16); CStr(linha)
      Print #1, "MOTIVO"; Tab(16); ErrNum
      Print #1, "DESCRIÇÃO"; Tab(16); ErrDescr
      Print #1, ""
      Print #1, "-----------------------------------------------------------------------------------------------------"
      Print #1, ""
      Close #1 'FECHA O ARQUIVO
      'SE O PARÂMETRO ExibeMensagem = True , EXIBE MENSAGEM PARA O USUÁRIO
      If ExibeMensagem = True Then
         MsgBox "A operação não pode ser completada, consulte o arquivo: " & App.path & "\Controles\GeoSanLog.txt" & " para maiores detalhes.", vbInformation
      End If
End Function


Public Function CorrigeQuadrante()


On Error GoTo Trata_Erro
   
   Dim rs As New ADODB.Recordset
   Dim S_LOWER_X As String, S_LOWER_Y As String, S_UCASE_X As String, S_UCASE_Y As String
    If frmCanvas.TipoConexao <> 4 Then

   Set rs = Conn.execute("SELECT LOWER_X,LOWER_Y,UCASE_X,UCASE_Y FROM TE_LAYER where NAME = 'QUADRANTE_REF'")

   'Set rs = CONN.Execute("SELECT LOWER_X,LOWER_Y,UCASE_X,UCASE_Y FROM TE_VIEW where lower_y <> null ORDER BY LOWER_X DESC")
   If rs.EOF = False Then
      
   'update te_layer set lower_x = 288212;
   'update te_layer set lower_y = 7424974;
   'update te_layer set UCASE_x = 320593;
   'update te_layer set UCASE_y = 7445655;
      
      S_LOWER_X = rs!lower_x
      S_LOWER_Y = rs!lower_y
      S_UCASE_X = rs!UCASE_X
      S_UCASE_Y = rs!UCASE_Y
      
   Else
      MsgBox "Não foi possível atualizar o quadrante." & Chr(13) & Chr(13) & "Verifique se possui o layer QUADRANTE_REF e se ele possui coordenadas.", vbInformation, ""
      Exit Function
   End If
   
   
   rs.Close
   
   
   If MsgBox("Os quadrantes dos Layers serão atualizados com os valores:" & _
      Chr(13) & Chr(13) & " LOWER_X  " & S_LOWER_X & _
      Chr(13) & " LOWER_Y  " & S_LOWER_Y & _
      Chr(13) & " UCASE_X   " & S_UCASE_X & _
      Chr(13) & " UCASE_Y   " & S_UCASE_Y & _
      Chr(13) & Chr(13) & "Se estes forem valores validos clique em SIM, caso desconheça clique em NÃO.   " & _
      Chr(13) & Chr(13) & "Atenção: Esta operação não poderá ser desfeita.", vbQuestion + vbYesNo + vbDefaultButton2, "") = vbYes Then
         
         

      Conn.execute ("UPDATE TE_LAYER SET LOWER_X = '" & S_LOWER_X & "' , LOWER_Y = '" & S_LOWER_Y & "', UCASE_X = '" & S_UCASE_X & "', UCASE_Y = '" & S_UCASE_Y & "'")

      Conn.execute ("UPDATE TE_VIEW SET LOWER_X = '" & S_LOWER_X & "' , LOWER_Y = '" & S_LOWER_Y & "', UCASE_X = '" & S_UCASE_X & "', UCASE_Y = '" & S_UCASE_Y & "'")

      Conn.execute ("UPDATE TE_REPRESENTATION SET LOWER_X = '" & S_LOWER_X & "' , LOWER_Y = '" & S_LOWER_Y & "', UCASE_X = '" & S_UCASE_X & "', UCASE_Y = '" & S_UCASE_Y & "'")
      
      Conn.execute ("UPDATE TE_THEME SET LOWER_X = '" & S_LOWER_X & "' , LOWER_Y = '" & S_LOWER_Y & "', UCASE_X = '" & S_UCASE_X & "', UCASE_Y = '" & S_UCASE_Y & "'")
      
      MsgBox "Concluído!", vbInformation, ""
      
   End If
   
  End If
   If frmCanvas.TipoConexao = 4 Then
   
   a = "LOWER_X"
   b = "LOWER_Y"
   c = "UCASE_X"
   d = "UCASE_Y"
   e = "te_layer"
    f = "UCASE_X"
   g = "UCASE_Y"
   h = "te_representation"
   i = "te_view"
   j = "te_theme"
   k = "name"
   l = "te_layer"
   
   Set rs = Conn.execute("SELECT " + """" + a + """" + "," + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + " from " + """" + l + """" + "where = " + """" + k + """" + " = 'QUADRANTE_REF'")

   'Set rs = CONN.Execute("SELECT LOWER_X,LOWER_Y,UCASE_X,UCASE_Y FROM TE_VIEW where lower_y <> null ORDER BY LOWER_X DESC")
   If rs.EOF = False Then
      
   'update te_layer set lower_x = 288212;
   'update te_layer set lower_y = 7424974;
   'update te_layer set UCASE_x = 320593;
   'update te_layer set UCASE_y = 7445655;
      
      S_LOWER_X = rs!lower_x
      S_LOWER_Y = rs!lower_y
      S_UCASE_X = rs!UCASE_X
      S_UCASE_Y = rs!UCASE_Y
      
   Else
      MsgBox "Não foi possível atualizar o quadrante." & Chr(13) & Chr(13) & "Verifique se possui o layer QUADRANTE_REF e se ele possui coordenadas.", vbInformation, ""
      Exit Function
   End If
   
   
   rs.Close
   
   
   If MsgBox("Os quadrantes dos Layers serão atualizados com os valores:" & _
      Chr(13) & Chr(13) & " LOWER_X  " & S_LOWER_X & _
      Chr(13) & " LOWER_Y  " & S_LOWER_Y & _
      Chr(13) & " UCASE_X   " & S_UCASE_X & _
      Chr(13) & " UCASE_Y   " & S_UCASE_Y & _
      Chr(13) & Chr(13) & "Se estes forem valores validos clique em SIM, caso desconheça clique em NÃO.   " & _
      Chr(13) & Chr(13) & "Atenção: Esta operação não poderá ser desfeita.", vbQuestion + vbYesNo + vbDefaultButton2, "") = vbYes Then
         
         
         
      'Conn.execute ("UPDATE TE_LAYER SET LOWER_X = '" & S_LOWER_X & "' , LOWER_Y = '" & S_LOWER_Y & "', UCASE_X = '" & S_UCASE_X & "', UCASE_Y = '" & S_UCASE_Y & "'")
  Conn.execute ("UPDATE " + """" + l + """" + " SET " + """" + a + """" + " = '" & S_LOWER_X & "' , " + """" + b + """" + " = '" & S_LOWER_Y & "', " + """" + c + """" + " = '" & S_UCASE_X & "', " + """" + d + """" + " = '" & S_UCASE_Y & "'")
  
  Conn.execute ("UPDATE " + """" + i + """" + " SET " + """" + a + """" + " = '" & S_LOWER_X & "' , " + """" + b + """" + " = '" & S_LOWER_Y & "', " + """" + c + """" + " = '" & S_UCASE_X & "', " + """" + d + """" + " = '" & S_UCASE_Y & "'")

      Conn.execute ("UPDATE " + """" + h + """" + " SET " + """" + a + """" + " = '" & S_LOWER_X & "' , " + """" + b + """" + " = '" & S_LOWER_Y & "', " + """" + c + """" + " = '" & S_UCASE_X & "', " + """" + d + """" + " = '" & S_UCASE_Y & "'")

      
      Conn.execute ("UPDATE " + """" + j + """" + " SET " + """" + a + """" + " = '" & S_LOWER_X & "' , " + """" + b + """" + " = '" & S_LOWER_Y & "', " + """" + c + """" + " = '" & S_UCASE_X & "', " + """" + d + """" + " = '" & S_UCASE_Y & "'")

      
      MsgBox "Concluído!", vbInformation, ""
   
   End If
   
End If
Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   MsgBox "Não foi possível atualizar o quadrante." & Chr(13) & Chr(13) & "Verifique se possui o layer QUADRANTE_REF e se ele possui coordenadas.", vbInformation, ""

End If

End Function
Public Function VerificaBasePOSTGRES()




End Function


Public Function VerificaBaseSQL() 'ATUALIZAÇÃO DA BASE DE DADOS SE SQL

On Error GoTo Trata_Erro
  
    Dim SQL As String
    Dim arrSQL(100) As String
    Dim i As Integer
    Dim ctErro As Integer
    Dim Erro As Integer
Dim g13 As String
g13 = "NX_BASE"
Inicio:
    Dim rsBase As ADODB.Recordset
        If frmCanvas.TipoConexao <> 4 Then
    SQL = "SELECT VERSAO FROM NX_BASE"
    Else
     SQL = "SELECT VERSAO FROM " + g13 + ""
    End If
    Set rsBase = Conn.execute(SQL)
    If rsBase.EOF = False Then
       If Trim(rsBase!versao) = "6.0.0" Then
        

            Exit Function
        Else
                  If MsgBox("A versão da base de dados não é compatível com a versão do aplicativo." & Chr(13) & Chr(13) & "Deseja atualizar a base de dados?", vbQuestion + vbYesNo, "") = vbYes Then
               GoTo ATUALIZACAO600
            Else
               GoTo fim
            End If
        End If
    Else
    
       
        If MsgBox("A versão da base de dados não é compatível com a versão do aplicativo." & Chr(13) & Chr(13) & "Deseja atualizar a base de dados?", vbQuestion + vbYesNo, "") = vbYes Then
            ' arrSQL(0) = "Insert into NX_BASE(versao) values('5.8.0')"
           '  Conn.execute (arrSQL(0))
             ' If frmCanvas.TipoConexao <> 4 Then
              
            ' SQL = "SELECT VERSAO FROM NX_BASE"
            ' Else
             'g13 = "NX_BASE"
            ' 'SQL = "SELECT VERSAO FROM " + g13 + ""
             GoTo ATUALIZACAO600
   ' End If
            ' Set rsBase = Conn.execute(SQL)
             'If rsBase.EOF = False Then
               ' GoTo ATUALIZACAO600
            ' End If
         Else
            End
        End If
       

   End If


CRIA_NXBASE:
    SQL = "CREATE TABLE NX_BASE (VERSAO VARCHAR(10))" 'cria a tabela NX_BASE
    Conn.execute (SQL)
    SQL = "INSERT INTO NX_BASE (VERSAO) VALUES ('0.0.0')" 'insere valor para que o próximo update funcione
    Conn.execute (SQL)
    GoTo Inicio

ATUALIZACAO600:
    
    i = 0
 If Trim(rsBase!versao) < "5.9.8" Then
    arrSQL(0) = "ALTER TABLE SYSTEMUSERS ADD USRMAIL VARCHAR(50)"
    arrSQL(1) = "ALTER TABLE SYSTEMUSERS ADD USRDEPTO VARCHAR(50)"
    arrSQL(2) = "ALTER TABLE SYSTEMUSERS ADD USRDATA VARCHAR(8)"
   
    'ATUALIZADO PARA APARECER O FILTRO ATUAL/EXISTENTE EM FILTROS DE LAYERS
    arrSQL(3) = "CREATE TABLE NXGS_FILT_TEMA (THEME_ID INTEGER, FILT_1 VARCHAR(100), FILT_2 VARCHAR(100),FILT_3 VARCHAR(100))"
        
    arrSQL(4) = "ALTER TABLE WATERLINES ADD ROUGHNESS FLOAT(14) DEFAULT 0"
    arrSQL(5) = "ALTER TABLE WATERLINES ADD DATEINSTALLATION DATETIME"
    arrSQL(6) = "ALTER TABLE WATERLINES ADD SIDESTREET NUMERIC"
    arrSQL(7) = "ALTER TABLE WATERLINES ADD DIVIDEDDISTANCE FLOAT(10)"
    arrSQL(8) = "ALTER TABLE WATERLINES ADD TROUBLE FLOAT(2)"
    arrSQL(9) = "ALTER TABLE WATERLINES ADD MANUFACTURER NUMERIC"
    arrSQL(10) = "ALTER TABLE WATERLINES ALTER COLUMN INITIALGROUNDHEIGHT FLOAT"
    
    'ATUALIZADO PARA APARECER USUÁRIO E DATA NO GRID DOS ATRIBUTOS 26/11/08
    arrSQL(11) = "ALTER TABLE WATERLINES ADD USUARIO_LOG VARCHAR(50)"
    arrSQL(12) = "ALTER TABLE WATERLINES ADD DATA_LOG VARCHAR(50)"
    arrSQL(13) = "ALTER TABLE WATERLINES ADD DATALOG DATETIME NOT NULL DEFAULT(GETDATE())"
    
    arrSQL(14) = "ALTER TABLE SEWERLINES ADD ROUGHNESS FLOAT(14) DEFAULT 0"
    arrSQL(15) = "ALTER TABLE SEWERLINES ADD DATEINSTALLATION DATETIME"
    arrSQL(16) = "ALTER TABLE SEWERLINES ADD SIDESTREET NUMERIC"
    arrSQL(17) = "ALTER TABLE SEWERLINES ADD DIVIDEDDISTANCE FLOAT(10)"
    arrSQL(18) = "ALTER TABLE SEWERLINES ADD TROUBLE FLOAT(2) DEFAULT 0"
    arrSQL(19) = "ALTER TABLE SEWERLINES ADD MANUFACTURER NUMERIC"
    arrSQL(20) = "ALTER TABLE SEWERLINES ALTER COLUMN INITIALGROUNDHEIGHT FLOAT"
    
    arrSQL(21) = "ALTER TABLE SEWERLINES ADD USUARIO_LOG VARCHAR(50)"
    arrSQL(22) = "ALTER TABLE SEWERLINES ADD DATA_LOG VARCHAR(50)"
    arrSQL(23) = "ALTER TABLE SEWERLINES ADD DATALOG DATETIME NOT NULL DEFAULT(GETDATE())"
    
    'Versão 5.6.8
    arrSQL(24) = "ALTER TABLE SEWERCOMPONENTS ADD DATEINSTALLATION DATETIME"
    arrSQL(25) = "ALTER TABLE SEWERCOMPONENTS ADD TROUBLE FLOAT(2) DEFAULT 0"
    arrSQL(26) = "ALTER TABLE SEWERCOMPONENTS ADD PATTERN FLOAT(10)"
    arrSQL(27) = "ALTER TABLE SEWERCOMPONENTS ADD SECTOR NUMERIC"
    arrSQL(28) = "ALTER TABLE SEWERCOMPONENTS ADD INFORMATIONVALIDITY FLOAT(10)"
    arrSQL(29) = "ALTER TABLE SEWERCOMPONENTS ADD GROUNDHEIGHTFINAL FLOAT(10) DEFAULT 0"
    
    arrSQL(30) = "DELETE FROM X_YES_NO"
    arrSQL(31) = "SET IDENTITY_INSERT X_YES_NO ON"
    arrSQL(32) = "INSERT INTO X_YES_NO (ID,DESCRIPTION) VALUES (0,'Não')"

    arrSQL(33) = "INSERT INTO X_YES_NO (ID,DESCRIPTION) VALUES (1,'Sim')"
    arrSQL(34) = "SET IDENTITY_INSERT X_YES_NO OFF"
    arrSQL(35) = "UPDATE WATERCOMPONENTS SET TROUBLE = 0 WHERE TROUBLE IS NULL"
    arrSQL(36) = "UPDATE WATERCOMPONENTS SET TROUBLE = 0 WHERE TROUBLE = 2"
    
    arrSQL(37) = "ALTER TABLE WATERCOMPONENTS ADD CONSTRAINT DF_WATERCOMPONENTS_TROUBLE DEFAULT (0) FOR TROUBLE"
    
    'RAMAIS_AGUA
    arrSQL(38) = "ALTER TABLE RAMAIS_AGUA ADD DATA_LOG VARCHAR(30)"
    arrSQL(39) = "ALTER TABLE RAMAIS_AGUA ADD USUARIO_LOG VARCHAR(30)"
    
    'RAMAIS_AGUA_LIGACAO
    arrSQL(40) = "ALTER TABLE RAMAIS_AGUA_LIGACAO ADD CONSUMO_LPS [numeric](24, 8) NULL DEFAULT ('0')"
    arrSQL(41) = "ALTER TABLE RAMAIS_AGUA_LIGACAO ADD TIPO VARCHAR(20)"
    arrSQL(42) = "ALTER TABLE RAMAIS_AGUA_LIGACAO ADD ECONOMIAS [numeric](10, 0)"
    arrSQL(43) = "ALTER TABLE RAMAIS_AGUA_LIGACAO ADD HIDROMETRADO [varchar] (3)"

    arrSQL(44) = "ALTER TABLE DRAINLINES ADD USUARIO_LOG VARCHAR(50)"
    arrSQL(45) = "ALTER TABLE DRAINLINES ADD DATA_LOG VARCHAR(50)"
    arrSQL(46) = "ALTER TABLE DRAINLINES ADD DATALOG DATETIME NOT NULL DEFAULT(GETDATE())"
    arrSQL(47) = "ALTER TABLE DRAINLINES ALTER COLUMN INITIALGROUNDHEIGHT FLOAT"
    arrSQL(48) = "ALTER TABLE DRAINLINES ADD ROUGHNESS FLOAT(14) default 0"
    arrSQL(49) = "ALTER TABLE DRAINLINES ADD DATEINSTALLATION DATETIME"
    arrSQL(50) = "ALTER TABLE DRAINLINES ADD SIDESTREET NUMERIC"
    arrSQL(51) = "ALTER TABLE DRAINLINES ADD DIVIDEDDISTANCE FLOAT(10)"
    arrSQL(52) = "ALTER TABLE DRAINLINES ADD TROUBLE NUMERIC"
    arrSQL(53) = "ALTER TABLE DRAINLINES ADD MANUFACTURER NUMERIC"
    
    arrSQL(54) = "CREATE TABLE POLIGONO_SELECAO (OBJECT_ID_ VARCHAR(10),USUARIO VARCHAR(20), TIPO INT)" ' INSERIDA FUNCIONALIDADE DE EXPORTAR POR SELECAO
    
    
   For i = 0 To 100
      If arrSQL(i) <> "" Then
         Conn.execute (arrSQL(i))
      End If
   Next

End If

    arrSQL(0) = "CREATE TABLE DRAINCOMPONENTSSELECTIONS (ID_TYPE INT ,ID_SUBTYPE INT ,OPTION_ VARCHAR(25),VALUE_ INT,DESCRIPTION_ VarChar(30))"
    arrSQL(1) = "CREATE TABLE DRAINLINES (LINE_ID INT NOT NULL , OBJECT_ID_ INT NOT NULL , ID_TYPE INT DEFAULT 0,INITIALGROUNDHEIGHT INT DEFAULT 0,FINALGROUNDHEIGHT INT DEFAULT 0,INITIALTUBEDEEPNESS INT DEFAULT 0,FINALTUBEDEEPNESS INT DEFAULT 0,INTERNALDIAMETER INT DEFAULT 0,EXTERNALDIAMETER INT DEFAULT 0,INITIALCOMPONENT INT  DEFAULT 0,FINALCOMPONENT INT  DEFAULT 0,THICKNESS INT  DEFAULT 0,MATERIAL INT  DEFAULT 0,LENGTH INT  DEFAULT 0,LENGTHCALCULATED INT  DEFAULT 0,SUPPLIER INT  DEFAULT 0,LOCATION INT  DEFAULT 0,STATE INT  DEFAULT 0,INFORMATIONVALIDITY INT  DEFAULT 0, SECTOR INT  DEFAULT 0,MANUFACTURER INT  DEFAULT 0,ROUGHNESS INT  DEFAULT 0,DATEINSTALLATION DATETIME,SIDESTREET NCHAR(50),DIVIDEDDISTANCE INT  DEFAULT 0,USUARIO_LOG NCHAR(200),DATA_LOG varchar(50))"
    arrSQL(2) = "CREATE TABLE DRAINLINESSUBTYPES (ID_TYPE INT DEFAULT 0, ID_SUBTYPE INT, DESCRIPTION_ VARCHAR(50), SELECTION_ INT DEFAULT 0, MAX_ INT DEFAULT 0, MIN_ INT DEFAULT 0, DEFAULTVALUE VARCHAR(200), DATATYPE INT DEFAULT 0, EPAREF VARCHAR(10))"
    arrSQL(3) = "CREATE TABLE DRAINLINESDATA (ID_TYPE INT DEFAULT 0, ID_SUBTYPE INT DEFAULT 0, OBJECT_ID INT DEFAULT 0, VALUE_ INT DEFAULT 0)"
    arrSQL(4) = "CREATE TABLE DRAINLINESTYPES(ID_TYPE INT, DESCRIPTION VARCHAR(25), SPECIFICATION VARCHAR(100))"
    arrSQL(5) = "CREATE TABLE DRAINLINESSELECTIONS(ID_TYPE INT DEFAULT 0, ID_SUBTYPE INT DEFAULT 0,OPTION_ VARCHAR(200),VALUE_ INT DEFAULT 0,DESCRIPTION_ VARCHAR(200))"
    arrSQL(6) = "CREATE TABLE DRAINCOMPONENTSDATA(OBJECT_ID INT DEFAULT 0,ID_TYPE INT DEFAULT 0,ID_SUBTYPE INT DEFAULT 0,VALUE VARCHAR(50))"
    arrSQL(7) = "INSERT INTO GS_LAYER_CONFIG_LAYERS(LAYER_ID, LAYER_REFERENCE,TYPE_OPERATION,DESCRIPTION_OPERATION) VALUES(5,6,0,'INSERIR REDE DE DRENAGEM')"
    arrSQL(8) = "INSERT INTO GS_LAYER_CONFIG_LAYERS(LAYER_ID, LAYER_REFERENCE,TYPE_OPERATION,DESCRIPTION_OPERATION) VALUES(6,5,0,'INSERIR COMPONTE DE DRENAGEM')"
    arrSQL(12) = "Delete  from GS_LAYER_CONFIG_LAYERS where layer_id = 5"
    arrSQL(13) = "Delete from GS_LAYER_CONFIG_LAYERS where layer_id = 6"
    arrSQL(9) = "CREATE TABLE DRAINCOMPONENTS(COMPONENT_ID INT , OBJECT_ID_ INT DEFAULT 0, ID_TYPE INT,YEAROFCONSTRUCTION INT DEFAULT 0, STATE INT DEFAULT 0, LOCATION INT DEFAULT 0, SUPPLIER INT DEFAULT 0, MANUFACTURER INT DEFAULT 0, GROUNDHEIGHT INT DEFAULT 0, INFORMATIONVALIDITY INT DEFAULT 0, NOTES INT DEFAULT 0, DEMAND INT DEFAULT 0, ESPECIAL INT DEFAULT 0, GROUNDHEIGHTFINAL INT DEFAULT 0, ROUGHNESS INT DEFAULT 0, DATEINSTALLATION DATETIME,  TROUBLE NCHAR(50), PATTERN NCHAR(50), SECTOR INT DEFAULT 0)"
    arrSQL(10) = "CREATE TABLE DRAINCOMPONENTSTYPES(ID_TYPE INT,DESCRIPTION VARCHAR(25),SPECIFICATION_ VARCHAR(100))"
    arrSQL(11) = "CREATE TABLE DRAINCOMPONENTSSUBTYPES(ID_TYPE INT, ID_SUBTYPE INT)"
      
      arrSQL(14) = "Update NX_BASE set versao = '6.0.0'"

          
                     
    Conn.execute (arrSQL(14))
    Conn.execute (arrSQL(0))
    Conn.execute (arrSQL(1))
    Conn.execute (arrSQL(2))
    Conn.execute (arrSQL(3))
    Conn.execute (arrSQL(4))
    Conn.execute (arrSQL(5))
    Conn.execute (arrSQL(6))
    Conn.execute (arrSQL(9))
    Conn.execute (arrSQL(10))
    Conn.execute (arrSQL(11))
    Conn.execute (arrSQL(12))
    Conn.execute (arrSQL(13))
    Conn.execute (arrSQL(7))
    Conn.execute (arrSQL(8))



 
  If Erro = 0 Then
      'SQL = "UPDATE NX_BASE SET VERSAO = '6.0.0'"
     ' Conn.execute (SQL)
      MsgBox "Base de dados atualizada com sucesso!", vbInformation, ""
   Else
      MsgBox "A atualização não foi concluída." & Chr(13) & Chr(13) & "Consulte o arquivo de Log para maiores informações", vbInformation, ""
   End If
     

fim:

Trata_Erro:
   Close #1
    If Err.Number = 0 Or Err.Number = 20 Then
    
        Resume Next
    ElseIf Err.Number = -2147217900 Then  'A coluna ja existe na tabela
        
        PrintErro "Global", arrSQL(i), CStr(Err.Number), CStr(Err.Description), False
        Err.Clear
        Resume Next
    
    ElseIf Err.Number = -2147217865 Then  'A tabela não existe
        If i = 0 Then
            Err.Clear
            GoTo CRIA_NXBASE
        Else
            
            PrintErro "Global", arrSQL(i), CStr(Err.Number), CStr(Err.Description), False
            Err.Clear
            Resume Next
            
        End If
    Else
         
         PrintErro "Global", arrSQL(i), CStr(Err.Number), CStr(Err.Description), False

         Resume Next
        
    End If

End Function

Public Function VerificaBaseORACLE() 'ATUALIZAÇÃO DA BASE DE DADOS SE ORACLE
On Error GoTo Trata_Erro_oracle

    Dim SQL As String
    Dim arrSQL(100) As String
    Dim i As Integer
    Dim Erro As Byte
    Dim ctErro As Integer
   
Inicio:
   Dim rsBase As ADODB.Recordset
    SQL = "SELECT VERSAO FROM NX_BASE"
    Set rsBase = Conn.execute(SQL)
    If rsBase.EOF = False Then
    
     If Trim(rsBase!versao) = "6.0.0" Then
     
   
           Exit Function
        Else
            If MsgBox("A versão da base de dados não é compatível com a versão do aplicativo." & Chr(13) & Chr(13) & "Deseja atualizar a base de dados?", vbQuestion + vbYesNo, "") = vbYes Then
               GoTo ATUALIZA600
            Else
               
       GoTo fim
            End If
        End If
    Else
    
       
If MsgBox("A versão da base de dados não é compatível com a versão do aplicativo." & Chr(13) & Chr(13) & "Deseja atualizar a base de dados?", vbQuestion + vbYesNo, "") = vbYes Then

     
            
             'arrSQL(0) = "Insert into NX_BASE(versao) values('5.8.0')"
        
       ' Conn.execute (arrSQL(0))
       
         '   SQL = "SELECT VERSAO FROM NX_BASE"
   ' Set rsBase = Conn.execute(SQL)
    'If rsBase.EOF = False Then
            GoTo ATUALIZA600
          '  End If
            Else
               End
           End If
       


End If
CRIA_NXBASE:
    SQL = "CREATE TABLE NX_BASE (VERSAO VARCHAR(10))" 'cria a tabela NX_BASE
    Conn.execute (SQL)
    SQL = "INSERT INTO NX_BASE (VERSAO) VALUES ('0.0.0')" 'insere valor para que o próximo update funcione
    Conn.execute (SQL)
    GoTo Inicio


ATUALIZA600:
  
    
  
 If Trim(rsBase!versao) < "5.9.8" Then


   'CARREGA EM UM ARRAY TODOS OS SCRIPTS DE ATUALIZAÇÃO DA BASE DE DADOS

    arrSQL(0) = "ALTER TABLE SYSTEMUSERS ADD (USRMAIL VARCHAR2(50))"
    arrSQL(1) = "ALTER TABLE SYSTEMUSERS ADD (USRDEPTO VARCHAR2(50))"
    arrSQL(2) = "ALTER TABLE SYSTEMUSERS ADD (USRDATA VARCHAR2(8))"
    
    'ATUALIZADO PARA APARECER O FILTRO ATUAL/EXISTENTE EM FILTROS DE LAYERS
    arrSQL(3) = "CREATE TABLE NXGS_FILT_TEMA (THEME_ID INTEGER, FILT_1 VARCHAR2(100), FILT_2 VARCHAR2(100),FILT_3 VARCHAR2(100))"
    
    arrSQL(4) = "ALTER TABLE WATERLINES ADD (ROUGHNESS NUMBER(14,0) default 0)"
    arrSQL(5) = "ALTER TABLE WATERLINES ADD (DATEINSTALLATION date)"
    arrSQL(6) = "ALTER TABLE WATERLINES ADD (SIDESTREET NUMBER(2))"
    arrSQL(7) = "ALTER TABLE WATERLINES ADD (DIVIDEDDISTANCE NUMBER(10))"
    arrSQL(8) = "ALTER TABLE WATERLINES ADD (TROUBLE NUMBER(2))"
    arrSQL(9) = "ALTER TABLE WATERLINES ADD (MANUFACTURER NUMBER(10))"
    arrSQL(10) = "ALTER TABLE WATERLINES MODIFY (INITIALGROUNDHEIGHT NUMBER(38,2))"
    
    
    'ATUALIZADO PARA APARECER USUÁRIO E DATA NO GRID DOS ATRIBUTOS 26/11/08
    arrSQL(11) = "ALTER TABLE WATERLINES ADD (USUARIO_LOG VARCHAR2(50))"
    arrSQL(12) = "ALTER TABLE WATERLINES ADD (DATA_LOG VARCHAR2(50))"
    arrSQL(13) = "ALTER TABLE WATERLINES ADD (DATALOG DATE DEFAULT SYSDATE)"
   
    arrSQL(14) = "ALTER TABLE SEWERLINES ADD (ROUGHNESS NUMBER(14,0) default 0)"
    arrSQL(15) = "ALTER TABLE SEWERLINES ADD (DATEINSTALLATION date)"
    arrSQL(16) = "ALTER TABLE SEWERLINES ADD (SIDESTREET NUMBER(2))"
    arrSQL(17) = "ALTER TABLE SEWERLINES ADD (DIVIDEDDISTANCE NUMBER(10))"
    arrSQL(18) = "ALTER TABLE SEWERLINES ADD (TROUBLE NUMBER(2))"
    arrSQL(19) = "ALTER TABLE SEWERLINES ADD (MANUFACTURER NUMBER(10))"
    arrSQL(20) = "ALTER TABLE SEWERLINES MODIFY (INITIALGROUNDHEIGHT NUMBER(38,2))" 'manoel alterou de 38,16 para 38,2
    
    arrSQL(21) = "ALTER TABLE SEWERLINES ADD (USUARIO_LOG VARCHAR2(50))"
    arrSQL(22) = "ALTER TABLE SEWERLINES ADD (DATA_LOG VARCHAR2(50))"
    arrSQL(23) = "ALTER TABLE SEWERLINES ADD (DATALOG DATE DEFAULT SYSDATE)"
    
    'Versão 5.6.8
    arrSQL(24) = "ALTER TABLE SEWERCOMPONENTS ADD (DATEINSTALLATION DATE)"
    arrSQL(25) = "ALTER TABLE SEWERCOMPONENTS ADD (TROUBLE NUMBER(2,0) DEFAULT (0))"
    arrSQL(26) = "ALTER TABLE SEWERCOMPONENTS ADD (PATTERN NUMBER(10,0))"
    arrSQL(27) = "ALTER TABLE SEWERCOMPONENTS ADD (SECTOR NUMBER(10,0))"
    arrSQL(28) = "ALTER TABLE SEWERCOMPONENTS ADD (INFORMATIONVALIDITY NUMBER(10,0))"
    arrSQL(29) = "ALTER TABLE SEWERCOMPONENTS ADD (GROUNDHEIGHTFINAL NUMBER(14,2) DEFAULT (0))"  'manoel alterou de 14,16 para 14,2


    'MODIFICAÇÕES REFERENTES A INFORMAÇÃO NÃO CONFORMIDADE
    arrSQL(30) = "DELETE FROM X_YES_NO"
    
    arrSQL(31) = "ALTER TABLE X_YES_NO MODIFY CONSTRAINT PK_X_YES_NO DISABLE"
    
    arrSQL(32) = "ALTER TRIGGER X_YES_NO_TRI DISABLE"
    
    arrSQL(33) = "INSERT INTO X_YES_NO (ID,DESCRIPTION) VALUES (0,'Não')"
    arrSQL(34) = "INSERT INTO X_YES_NO (ID,DESCRIPTION) VALUES (1,'Sim')"
    arrSQL(35) = "UPDATE WATERCOMPONENTS SET TROUBLE = 0 WHERE TROUBLE IS NULL"
    arrSQL(36) = "UPDATE WATERCOMPONENTS SET TROUBLE = 0 WHERE TROUBLE = 2"
    arrSQL(37) = "ALTER TABLE WATERCOMPONENTS MODIFY (TROUBLE DEFAULT (0))"
    
    'MODIFICAÇÕES REFERENTES A INFORMAÇÃO USUARIO E DATA DE CADASTRO DOS RAMAIS_AGUA
    arrSQL(38) = "ALTER TABLE RAMAIS_AGUA ADD DATA_LOG VARCHAR2(30)"
    arrSQL(39) = "ALTER TABLE RAMAIS_AGUA ADD USUARIO_LOG VARCHAR2(30)"
    
    arrSQL(40) = "ALTER TABLE NXGS_V_LIG_COMERCIAL ADD TIPO VARCHAR2(20)"

   'Alterar o valor do campo QUERYSTRING da tabela GS_QUERYS_CLIENT onde a QUERY_ID = 2 salvando o novo valor conforme abaixo:
   'SELECT LI.NRO_LIGACAO, LI.CLASSIFICACAO_FISCAL, LI.ENDERECO + '-' + ISNULL(LI.NUM_CASA,'') + '-' +  ISNULL(LI.COMPL_LOGRADOURO,'') + '-' + ISNULL(LI.BAIRRO,'') as Endereco, LI.CONSUMIDOR, LI.COD_LOGRADOURO as CODLOGRAD,TIPO FROM NXGS_V_LIG_COMERCIAL LI WHERE (LI.CLASSIFICACAO_FISCAL in (@CLASSIFICACAO_FISCAL) OR LI.NRO_LIGACAO in (@NRO_LIGACAO))
   
   'Alterar e atualizar a tabela RAMAIS_AGUA_LIGACAO incluindo os campos 'TIPO' e 'CONSUMO_LPS'
    arrSQL(41) = "ALTER TABLE RAMAIS_AGUA_LIGACAO ADD TIPO VARCHAR2(20) DEFAULT 'HIDROMETRADA'"
    arrSQL(42) = "ALTER TABLE RAMAIS_AGUA_LIGACAO ADD CONSUMO_LPS NUMBER(24, 8) DEFAULT (0)"
    
    arrSQL(43) = "ALTER TABLE RAMAIS_AGUA_LIGACAO ADD HIDROMETRADO VARCHAR2(3)"
    arrSQL(44) = "ALTER TABLE RAMAIS_AGUA_LIGACAO ADD ECONOMIAS NUMBER(10,0) DEFAULT (0)"
    
    
    'ESTA ATUALIZAÇÃO DEVE SER FEITA MANUALMENTE
    'arrSQL() = "UPDATE RAMAIS_AGUA_LIGACAO SET TIPO = 'HIDROMETRADA'"
    'arrSQL() = "UPDATE RAMAIS_AGUA_LIGACAO SET CONSUMO_LPS = 0"
    
    arrSQL(45) = "CREATE TABLE POLIGONO_SELECAO (OBJECT_ID_ VARCHAR2(10),USUARIO VARCHAR2(20), TIPO NUMBER(1,0))" ' INSERIDA FUNCIONALIDADE DE EXPORTAR POR SELECAO
    
    'ErrX = 1
    'ALTERAÇÕES DA DRAINLINES SEMPRE POR ULTIMO
    arrSQL(46) = "ALTER TABLE DRAINLINES ADD (USUARIO_LOG VARCHAR2(50))"
    arrSQL(47) = "ALTER TABLE DRAINLINES ADD (DATA_LOG VARCHAR2(50))"
    arrSQL(48) = "ALTER TABLE DRAINLINES ADD (DATALOG DATE DEFAULT SYSDATE)"
    arrSQL(49) = "ALTER TABLE DRAINLINES MODIFY (INITIALGROUNDHEIGHT NUMBER(38,2))"
    arrSQL(50) = "ALTER TABLE DRAINLINES ADD (ROUGHNESS NUMBER(14,0) default 0)"
    arrSQL(51) = "ALTER TABLE DRAINLINES ADD (DATEINSTALLATION date)"
    arrSQL(52) = "ALTER TABLE DRAINLINES ADD (SIDESTREET NUMBER(2))"
    arrSQL(53) = "ALTER TABLE DRAINLINES ADD (DIVIDEDDISTANCE NUMBER(10))"
    arrSQL(54) = "ALTER TABLE DRAINLINES ADD (TROUBLE NUMBER(2))"
    arrSQL(55) = "ALTER TABLE DRAINLINES ADD (MANUFACTURER NUMBER(10))"
    
    'EXECUTA TODOS OS SCRIPS ARMAZENADOS NO ARRAY
    
   For i = 0 To 100
      If arrSQL(i) <> "" Then
         Conn.execute (arrSQL(i))
      End If
   Next

 'arrSQL(56) = "Update NX_BASE set versao = '5.8.0'"
          
        
          
                     ' Conn.execute (arrSQL(56))


End If


  
   


    arrSQL(15) = "Delete  from GS_LAYER_CONFIG_LAYERS where layer_id = 5"
    arrSQL(16) = "Delete from GS_LAYER_CONFIG_LAYERS where layer_id = 6"
    
    arrSQL(0) = "INSERT INTO GS_LAYER_CONFIG_LAYERS(LAYER_ID, LAYER_REFERENCE,TYPE_OPERATION,DESCRIPTION_OPERATION) VALUES(5,6,0,'INSERIR REDE DE DRENAGEM')"
    
    arrSQL(1) = "INSERT INTO GS_LAYER_CONFIG_LAYERS(LAYER_ID, LAYER_REFERENCE,TYPE_OPERATION,DESCRIPTION_OPERATION) VALUES(6,5,0,'INSERIR COMPONTE DE DRENAGEM')"
    
    arrSQL(2) = "INSERT INTO GS_LAYER_TYPE_REFERENCE(TYPE_REFERENCE,DESCRIPTION) VALUES(9,'DOCUMENTOS')"
    
    arrSQL(3) = "INSERT INTO GS_LAYER_TYPE_REFERENCE(TYPE_REFERENCE,DESCRIPTION) VALUES(10,'CEL_PONTO_ATRIB')"
    
    arrSQL(4) = "INSERT INTO GS_LAYER_TYPE_REFERENCE(TYPE_REFERENCE,DESCRIPTION) VALUES(11,'LOTES')"
    
    arrSQL(5) = "CREATE TABLE DRAINLINES(LINE_ID NUMBER(*,0), OBJECT_ID_ NUMBER(*,0) DEFAULT 0, ID_TYPE NUMBER(*,0) DEFAULT 0, INITIALGROUNDHEIGHT NUMBER(*,0) DEFAULT 0, FINALGROUNDHEIGHT NUMBER(*,0) DEFAULT 0,INITIALTUBEDEEPNESS NUMBER(*,0) DEFAULT 0, FINALTUBEDEEPNESS NUMBER(*,0) DEFAULT 0, INTERNALDIAMETER NUMBER(*,0) DEFAULT 0,EXTERNALDIAMETER NUMBER(*,0) DEFAULT 0, INITIALCOMPONENT NUMBER(*,0) DEFAULT 0,FINALCOMPONENT NUMBER(*,0) DEFAULT 0,THICKNESS NUMBER(*,0) DEFAULT 0,MATERIAL NUMBER(*,0) DEFAULT 0,LENGTH NUMBER(*,0) DEFAULT 0,LENGTHCALCULATED NUMBER(*,0) DEFAULT 0,SUPPLIER NUMBER(*,0) DEFAULT 0,LOCATION NUMBER(*,0) DEFAULT 0,STATE NUMBER(*,0) DEFAULT 0,INFORMATIONVALIDITY NUMBER(*,0) DEFAULT 0,SECTOR NUMBER(*,0) DEFAULT 0,MANUFACTURER NUMBER(*,0) DEFAULT 0,ROUGHNESS NUMBER(*,0) DEFAULT 0,DATEINSTALLATION DATE DEFAULT sysdate,SIDESTREET NCHAR(50) DEFAULT 0,DIVIDEDDISTANCE NUMBER(*,0) DEFAULT 0,USUARIO_LOG NCHAR(200) DEFAULT 0,DATA_LOG varchar(50) DEFAULT '0')"
    arrSQL(6) = " CREATE TABLE DRAINLINESTYPES(ID_TYPE NUMBER(*,0), DESCRIPTION_ VARCHAR2(25), SPECIFICATION_ VARCHAR2(100))"
    
    arrSQL(7) = " CREATE TABLE DRAINLINESDATA(ID_TYPE NUMBER(*,0) DEFAULT 0, ID_SUBTYPE NUMBER(*,0) DEFAULT 0, OBJECT_ID_ NUMBER(*,0) DEFAULT 0, VALUE_ NUMBER(*,0) DEFAULT 0)"
    
    arrSQL(8) = "CREATE TABLE DRAINLINESSELECTIONS(ID_TYPE NUMBER(30,0) DEFAULT 0, ID_SUBTYPE NUMBER(30,0) DEFAULT 0, OPTION_ VARCHAR2(200), VALUE_ NUMBER DEFAULT 0, DESCRIPTION_ VARCHAR2(200))"
    
    arrSQL(9) = "CREATE TABLE DRAINLINESSUBTYPES(ID_TYPE INT DEFAULT 0, ID_SUBTYPE INT,DESCRIPTION_ VARCHAR(50),SELECTION_ INT DEFAULT 0,MAX_ INT DEFAULT 0,MIN_ INT DEFAULT 0,DEFAULTVALUE VARCHAR(200),DATATYPE INT DEFAULT 0,EPAREF VARCHAR(10))"
    
    arrSQL(10) = "CREATE TABLE DRAINCOMPONENTS(COMPONENT_ID NUMBER(*,0), OBJECT_ID_ NUMBER(*,0) DEFAULT 0, ID_TYPE NUMBER(*,0), YEAROFCONSTRUCTION NUMBER(*,0) DEFAULT 0, STATE NUMBER(*,0) DEFAULT 0, LOCATION NUMBER(*,0) DEFAULT 0, SUPPLIER NUMBER(*,0) DEFAULT 0, MANUFACTURER NUMBER(*,0) DEFAULT 0, GROUNDHEIGHT NUMBER(*,0) DEFAULT 0, INFORMATIONVALIDITY NUMBER(*,0) DEFAULT 0,NOTES NUMBER(*,0) DEFAULT 0,DEMAND NUMBER(*,0) DEFAULT 0,ESPECIAL NUMBER(*,0) DEFAULT 0,GROUNDHEIGHTFINAL NUMBER(*,0) DEFAULT 0,ROUGHNESS NUMBER(*,0) DEFAULT 0,DATEINSTALLATION DATE,TROUBLE NCHAR(50),PATTERN NCHAR(50),SECTOR NUMBER(*,0) DEFAULT 0)"
    
    arrSQL(11) = "CREATE TABLE DRAINCOMPONENTSTYPES(ID_TYPE NUMBER(*,0), DESCRIPTION VARCHAR2(25 BYTE), SPECIFICATION_ VARCHAR2(100 BYTE))"
    
    arrSQL(12) = "CREATE TABLE DRAINCOMPONENTSSUBTYPES(ID_TYPE NUMBER, ID_SUBTYPE NUMBER)"
    
    
    arrSQL(13) = " CREATE TABLE DRAINCOMPONENTSDATA(OBJECT_ID_ NUMBER(10,0) NOT NULL ,ID_TYPE NUMBER(10,0) NOT NULL , ID_SUBTYPE NUMBER(10,0) NOT NULL , VALUE_ VARCHAR2(50 BYTE) NOT NULL)"
    
    arrSQL(14) = "CREATE TABLE DRAINCOMPONENTSSELECTIONS(ID_TYPE NUMBER(10,0) NOT NULL ENABLE, ID_SUBTYPE NUMBER(10,0) NOT NULL ENABLE, OPTION_ VARCHAR2(25 BYTE) NOT NULL ENABLE, VALUE_ NUMBER(3,0) NOT NULL ENABLE, DESCRIPTION_ VARCHAR2(30 BYTE))"
    

         
        
        
          
    Conn.execute (arrSQL(5))
    Conn.execute (arrSQL(6))
    Conn.execute (arrSQL(7))
    Conn.execute (arrSQL(8))
    Conn.execute (arrSQL(9))
    Conn.execute (arrSQL(10))
    Conn.execute (arrSQL(11))
    Conn.execute (arrSQL(12))
    Conn.execute (arrSQL(13))
    Conn.execute (arrSQL(14))
    Conn.execute (arrSQL(15))
    Conn.execute (arrSQL(16))
    Conn.execute (arrSQL(0))
    Conn.execute (arrSQL(1))
    Conn.execute (arrSQL(2))
    Conn.execute (arrSQL(3))
    Conn.execute (arrSQL(4))
    
        

 If Erro = 0 Then
    '  SQL = "UPDATE NX_BASE SET VERSAO = '5.8.0'"
     ' Conn.execute (SQL)
      MsgBox "Base de dados atualizada com sucesso!", vbInformation, ""
   Else
      MsgBox "A atualização não foi concluída." & Chr(13) & Chr(13) & "Consulte o arquivo de Log para maiores informações", vbInformation, ""
   End If







fim:



Trata_Erro_oracle:
   Dim strDescription As String
   strDescription = mid(Err.Description, 1, 9)
    If Err.Number = 0 Or Err.Number = 20 Then
    
        Resume Next
    ElseIf strDescription = "ORA-00955" Or strDescription = "ORA-01440" Or strDescription = "ORA-01430" Then 'A tabela não existe
        
       PrintErro "Global", arrSQL(i), CStr(Err.Number), CStr(Err.Description), False
       Err.Clear
       Resume Next
    
    ElseIf mid(Err.Description, 1, 9) = "ORA-00942" Then 'A tabela não existe
      
      If i = 0 Then
         Err.Clear
         GoTo CRIA_NXBASE
      Else
       
         PrintErro "Global", arrSQL(i), CStr(Err.Number), CStr(Err.Description), False
         Err.Clear
         Resume Next
       
      End If
    
    Else
        
      PrintErro "Global", arrSQL(i), CStr(Err.Number), CStr(Err.Description), False
      Err.Clear
        
      Erro = 1 'MARCA QUE HOUVE 1 ERRO DESCONHECIDO, ENTÃO NÃO ATUALIZA O VALOR DE VERSÃO DO BANCO
    
    End If

End Function

'Public Function ConectaBanco() As Boolean
'On Error GoTo ErrorConnection
'
'    Dim lnFile As Integer
'    Dim lsParam As String
'    Dim lsConnect As String
'
'
'    Dim SqlServidor As String
'    Dim SqlBanco As String
'    Dim AcsBanco As String
'    Dim OracleServico As String
'
'    Dim strSenha As String
'
'    If Dir(App.path & "\Controles\GeoSan.cfg") = "" Then
'        MsgBox "Arquivo de inicialização do sistema não encontrado.", vbInformation, "Conexão indefinida"
'        FrmConnection.Show (1)
'        'ConectaBanco = False
'    End If
'
'    lnFile = FreeFile
'    Open App.path & "\Controles\GeoSan.cfg" For Input As #lnFile
'        Input #lnFile, lsParam  'Tipo de Conexão
'            frmCanvas.TipoConexao = lsParam
'        Input #lnFile, lsParam  'Servidor SQL
'            SqlServidor = lsParam
'        Input #lnFile, lsParam  'Banco SQL
'            SqlBanco = lsParam
'        Input #lnFile, lsParam  'Banco Access
'            AcsBanco = lsParam
'        Input #lnFile, lsParam  'Serviço Oracle
'            OracleServico = lsParam
'        Input #lnFile, lsParam  'Usuário
'            strUser = lsParam
'        Input #lnFile, lsParam  'Senha do banco de dados
'            strSenha = lsParam
'    Close #lnFile
'
'    Set MyConn = New ADODB.Connection
'
'    Select Case frmCanvas.TipoConexao
'        Case 0 'ACCESS
'            MyConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AcsBanco & ";Persist Security Info=False"
'            ConectaBanco = True
'        Case 1 'SqlServer
'            MyConn.Open "Provider=SQLOLEDB.1;Persist Security Info=True;Data Source=" & SqlServidor & ";User ID=" & strUser & ";Password=" & strSenha & ";Initial Catalog=" & SqlBanco
'            ConectaBanco = True
'        Case 2 'Oracle
'            MyConn.Open "Provider=OraOLEDB.Oracle.1;Password=" & strSenha & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & OracleServico
'            ConectaBanco = True
'    End Select
'    typeconnection = frmCanvas.TipoConexao
'Exit Function
'
'ErrorConnection:
'If Err.Number = 0 Or Err.Number = 20 Then
'    Resume Next
'ElseIf Err.Number = 5 Then
'    MsgBox "Caminho de arquivo não encontrado", vbInformation, "GeoSan.cfg"
'ElseIf Err.Number = 55 Then
'
'    Close #lnFile
'    Resume
'End If
'
'   MsgBox "Não foi possível estabelecer conexão com o Servidor de banco de dados." & Chr(13) & Chr(10) & "Erro -> " & Err.Description, vbInformation, "Erro !"
'   PrintErro "Global", "Public Function ConectaBanco() As Boolean", CStr(Err.Number), CStr(Err.Description), False
'   Err.Clear
'
'   ConectaBanco = False
'End Function

Public Function GetCboListIndex(ID As Long, mCbo As ComboBox) As Integer
   Dim a As Integer
   For a = 0 To mCbo.ListCount - 1
       If mCbo.ItemData(a) = ID Then
         GetCboListIndex = a
         Exit Function
       End If
   Next
   GetCboListIndex = -1
End Function



Function capturaRede(ByVal LayName As String, ByVal tcs As TeCanvas, ByRef allSELECTComponents As String) As Boolean
   '##################################################################################
   ' Autor     : Luis Claudio Rodrigues Domingues / Rodrigo Viviani    Data: 11/12/06
   ' Nome      : capturaRede
   ' Descrição : Captura sequenciamente todos elementos da redes criando uma tabela temporária
   '             e retorna todos o componentes selecionados
   '##################################################################################
   On Error GoTo captureRede_err
   a = "X_TEMPCALCULENODE"
   
     If frmCanvas.TipoConexao <> 4 Then

         Conn.execute "delete from X_TempCalculeNode  "
   Else
   Conn.execute "delete from " + """" + a + """"
   End If
   
   Dim componentesSelecionados As String ', Frm As frmSELECTnetWorkTypes
   Dim Object_id_Line_1 As String, contador As Integer, tipo As Long
   Dim rsValidaTrecho As ADODB.Recordset, rsTemp As ADODB.Recordset
   
   'Retorna tipo de rede a exportar
   Dim frm  As frmSelectnetWorkTypes
   Set frm = New frmSelectnetWorkTypes
   Dim k As String
   Dim l As String
   Dim m As String
   Dim n As String
   Dim o As String
   
   If Not frm.Init(tipo) Then Exit Function
   Set frm = Nothing
   '########################
   'Captura todos os nós selecionados
   For contador = 0 To tcs.getSelectCount(points) - 1
     If contador = 0 Then
        componentesSelecionados = "'" & tcs.getSelectObjectId(contador, points) & "'"
     Else
        componentesSelecionados = componentesSelecionados & ",'" & tcs.getSelectObjectId(contador, points) & "'"
     End If
   Next
   'Retorna o primeiro trecho encontrado a partir do primenrio nó selecionado
   If frmCanvas.TipoConexao <> 4 Then
   Set rsValidaTrecho = Conn.execute("SELECT object_id_ from WaterLines where (initialComponent=" & tcs.getSelectObjectId(0, points) & " or finalcomponent=" & tcs.getSelectObjectId(0, points) & ") and id_Type=" & tipo)
   Else
   k = "OBJECT_ID"
   l = "WATERLINES"
   m = "INITIALCOMPONENT"
   n = "FINALCOMPONENT"
   o = "ID_TYPE"
    Set rsValidaTrecho = Conn.execute("SELECT " + """" + k + """" + " from " + """" + l + """" + " where (" + """" + m + """" + "='" & tcs.getSelectObjectId(0, points) & "' or " + """" + n + """" + "='" & tcs.getSelectObjectId(0, points) & "') and " + """" + o + """" + "='" & tipo & "'")
   End If
   tcs.setCurrentLayer "WATERLINES"
   If Not rsValidaTrecho.EOF Then
      Object_id_Line_1 = rsValidaTrecho.Fields("object_id_").value
   Else
      MsgBox "Trecho não encontrado" '
      Exit Function
   End If
   If Not (rsValidaTrecho Is Nothing) Then
      If rsValidaTrecho.State = adStateOpen Then rsValidaTrecho.Close
   End If
   Set rsValidaTrecho = Nothing
         
   'Abre a Tabela  temporaria em modo de adçao
   Set rsTemp = New ADODB.Recordset
   rsTemp.Open "X_TempCalculeNode", Conn, adOpenKeyset, adLockOptimistic, adCmdTable
   'rsTemp.AddNew
   
   ' Chama a função recursiva que busca os trecho/componentes da rede sequenciamente
   Dim cgeo As New clsGeoReference
   Set cgeo.tcs = tcs
   cgeo.GetSubPrimaria Object_id_Line_1, rsTemp, tipo
   
   'Cancela o ultimo registro pois não foi inserido  nada nele vide codigo recursivo
   rsTemp.Update
   'fecha o recordset
   rsTemp.Close
   'mata o recordset
   Set rsTemp = Nothing
   allSELECTComponents = componentesSelecionados
   capturaRede = True
   Exit Function
captureRede_err:
   


   
End Function

Public Sub ExporteCrede(MyObject As String, nomeArq As String, LayName As String)
   Dim objRecordset As Recordset
   Dim sQry As String
   Dim arqrede As Integer
   Dim iNo As Long
   Dim minX, minY, maxX, maxY As Long
   Dim SubType As Long
   Dim cgeo As New clsGeoReference
    'On Error GoTo gtErro
    Dim xi As Double, yi As Double, xf As Double, yf As Double
    arqrede = FreeFile
    Open nomeArq For Output As arqrede
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -
    '                       Versão de formato de arquivo
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -
    Print #arqrede, "[VERSAO]"
    Print #arqrede, "1.0.0"
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -
    '                               Dados Gerais
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -
    Print #arqrede, "[DADGE]"
    'Vazão inicial do regime permanente
    Print #arqrede, "0"
    'Diâmetro inicial do regime permanente
    Print #arqrede, "1000"
    'Nó de referência piezométrica
    Print #arqrede, "1"
    'Cota de referencia piezométrica
    Print #arqrede, "0"
    'Diâmetro mínimo para o cálculo de diâmetros
    Print #arqrede, "50"
    'Diâmetro máximo para o cálculo de diâmetros
    Print #arqrede, "400"
    'Velocidade mínima para o dimensionamento
    Print #arqrede, "0,6"
    'Número máximo de iterações para o regime permanete
    Print #arqrede, "100"
    'Tolerância máxima para o regime permanente (m3/s)
    Print #arqrede, "0,00001"
    'Instante inicial para a simulação extendida ou para o transitório
    Print #arqrede, "0"
    'Instante final para a simulação extendida ou para o transitório
    Print #arqrede, "0"
    'Intervalo de tempo de cálculo
    Print #arqrede, "0,05"
    'Passo de saída para gravação de resultados
    Print #arqrede, "10"
    'Módulo de elasticidade do fluido
    Print #arqrede, "2200000000"
    'Massa específica do fluido
    Print #arqrede, "998,2"
    'Viscosidade cinemática do fluido
    Print #arqrede, "0,000001"
    'Recobrimento mínimo
    Print #arqrede, "1,5"
    'Método de cálculo: F. Universal = 0; Hazen-Willians=1
    Print #arqrede, "0"
    'Coordenada UTM Este da origem da área de trabalho
    Set objRecordset = New Recordset
    Dim icont As Integer ', MyObject
    
    If frmCanvas.TipoConexao <> 4 Then
    
    
    sQry = convertQuery("SELECT MIN(X) FROM POINTS" & cgeo.GetLayerID(LayName) & " where [object_id] in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    minX = Int(objRecordset.Fields(0) / 100) * 100
    Print #arqrede, CStr(minX) * 100
    'Coordenada UTM Norte da origem da área de trabalho
    sQry = convertQuery("SELECT MIN(Y) FROM POINTS" & cgeo.GetLayerID(LayName) & " where [object_id] in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    minY = Int(objRecordset.Fields(0) / 100) * 100
    Print #arqrede, CStr(minY) * 100
    'Comprimento na direção Este (cm)
    sQry = convertQuery("SELECT max(x) from points" & cgeo.GetLayerID(LayName) & " where [object_id] in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    maxX = Int(CSng(objRecordset.Fields(0) / 100) * 100 + 100)
    Print #arqrede, CStr(maxX - minX)
    'Comprimento na direção Norte (cm)
    sQry = convertQuery("SELECT MAX(Y) FROM POINTS" & cgeo.GetLayerID(LayName) & " where [object_id] in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    maxY = Int(CSng(objRecordset.Fields(0) / 100) * 100 + 100)
    Print #arqrede, CStr(maxY - minY)
    'Espaçamento do grid em metros
    Print #arqrede, CalculaAreaTrab(maxX - minX, maxY - minY)
    'Espaçamento do grid em twips
    'talvez não seja necessário, mas aí vai um teste
    Print #arqrede, 13785
    'Quantidade de nós
    sQry = convertQuery("SELECT COUNT(*) FROM POINTS" & cgeo.GetLayerID(LayName) & " where [object_id] in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    Print #arqrede, CStr(objRecordset.Fields(0))
    ReDim aListaNo(1 To objRecordset.Fields(0))
    'Quatidade de trechos

    sQry = "SELECT count(DISTINCT(LIN)) from X_TempCalculeNode"
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    'Print #arqrede, CStr(objRecordset.Fields(0))
    Print #arqrede, CStr(objRecordset.Fields(0))
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -
    '                                       Nós
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -
    sQry = convertQuery("SELECT " & _
           "P.Y, " & _
           "P.X, " & _
           "C.ID_TYPE, " & _
           "P.[OBJECT_ID], " & _
           "C.GROUNDHEIGHT, " & _
           "C.Demand " & _
           "FROM POINTS" & cgeo.GetLayerID(LayName) & " P " & _
           "LEFT JOIN WATERCOMPONENTS C ON P.[OBJECT_ID] = C.OBJECT_ID_ " & _
           "Where [object_id] in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    iNo = 1
    Dim RsCa As ADODB.Recordset
    While (Not objRecordset.EOF)
        'Nó
        Print #arqrede, "[NO]"
        'N em cm
        Print #arqrede, CStr(objRecordset.Fields("Y") * 100)
        'E em cm
        Print #arqrede, CStr(objRecordset.Fields("X") * 100)
        'Tipo
        Print #arqrede, CStr(CalculaSubType(objRecordset.Fields("id_TYPE")))
        'Rótulo[6]
        Print #arqrede, Left(objRecordset.Fields("OBJECT_ID"), 6)
        'Rótulo[20]
        Print #arqrede, Left(objRecordset.Fields("OBJECT_ID"), 20)
        'Valor característico
        'Print #arqrede, "0"
        
         Select Case objRecordset.Fields("id_TYPE")
           Case 0, 2, 18, 22 ' demanda ou vazao pontual
                'Desconhecido=0; Hidrante=6; Curva=7;
                'Conexão T=8; Cruzeta=9; Tap=10; Redução=11;
                'Hidrômetro = 12; Consumidor=2
                Print #arqrede, objRecordset.Fields("Demand")
           Case 20 'Bomba=1 ou buster 'sentido VALOR
                Set RsCa = Conn.execute("SELECT Value_ from watercomponentsdata where  id_type = 11  and id_subtype = 1 and object_id_ = " & objRecordset.Fields("OBJECT_ID"))
                If Not RsCa.EOF Then
                   Print #arqrede, RsCa(0) 'objRecordset.Fields("GROUNDHEIGHTFINAL")
                Else
                   Print #arqrede, 0
                End If
                RsCa.Close
           Case 19 'Reservatorio=3 Cota do nival dagua
                Set RsCa = Conn.execute("SELECT Value_ from watercomponentsdata where  id_type = 19  and id_subtype = 7 and object_id_ = " & objRecordset.Fields("OBJECT_ID"))
                If Not RsCa.EOF Then
                   Print #arqrede, RsCa(0) 'objRecordset.Fields("GROUNDHEIGHTFINAL")
                Else
                   Print #arqrede, 0
                End If
                RsCa.Close
           Case 21 'VRP=4  Perda de Carga - UNIDADE(m)
                Print #arqrede, "0"
                'Print #arqrede, objRecordset.Fields("GROUNDHEIGHTFINAL")
           Case 1 'V Controle    vc = porcentagem de abertura
                Set RsCa = Conn.execute("SELECT value_ from watercomponentsdata where  id_type = 1  and id_subtype = 3 and Object_id_=" & objRecordset.Fields("OBJECT_ID"))
                If Not RsCa.EOF Then
                   Print #arqrede, RsCa(0) 'objRecordset.Fields("GROUNDHEIGHTFINAL")
                Else
                   Print #arqrede, 0
                End If
                RsCa.Close
           Case Else
                Print #arqrede, "0"
         End Select
        'Print #arqrede, ValorCaracterisco(objRecordset.Fields("id_TYPE"))
        'k
        Print #arqrede, "0"
        'Cota do terreno
        Print #arqrede, CStr(objRecordset.Fields("GROUNDHEIGHT"))
        'Cota do nó
        Print #arqrede, CStr(objRecordset.Fields("GROUNDHEIGHT") - 1)
        'Pressão mín admitida (mca)
        Print #arqrede, "15"
        'Pressão máx admitida (mca)
        Print #arqrede, "50"
        'Nome da lei de operação: pra compatibilidade com os outros formatos de cabecalho
        Print #arqrede, ""
        'Lei de especificação da condição de contorno
        Print #arqrede, "Default"
        'Nó de jusante
        Print #arqrede, "0"
        'Confirmação de cota
        Print #arqrede, "-1"
        'booster: Indicação de sentido, ou
        'vrp: cirtério de cálculo
        Print #arqrede, "0"
        'situação de rede
        Print #arqrede, "Rede existente"
        aListaNo(iNo).indice = iNo
        aListaNo(iNo).object_id = objRecordset.Fields("OBJECT_ID")
        aListaNo(iNo).X = CStr(objRecordset.Fields("X") * 100)
        aListaNo(iNo).Y = CStr(objRecordset.Fields("Y") * 100)
        iNo = iNo + 1
        objRecordset.MoveNext
    Wend
    Set RsCa = Nothing
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -
    '                                     Trechos
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -
   sQry = "SELECT DISTINCT " & _
              "CPI AS  " + """" + "INITIALCOMPONENT" + """" + ", " & _
              "CPF AS " + """" + "FINALCOMPONENT" + """" + ", " & _
              "OBJECT_ID_, " & _
              "LENGTH, " & _
              "LengthCalculated," & _
              "INTERNALDiameter, " & _
              "Thickness, " & _
              "MATERIAL " & _
              "FROM WATERLINES INNER JOIN X_TempCalculeNode on lin = object_id_"


    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
 
    While (Not objRecordset.EOF)
        'Trecho
        Print #arqrede, "[TRECHO]"
        'Nó inicial
        Print #arqrede, CStr(QualIndiceNo(objRecordset.Fields("INITIALCOMPONENT")))
        'Nó final
        Print #arqrede, CStr(QualIndiceNo(objRecordset.Fields("FINALCOMPONENT")))
        'Rótulo[6]
        Print #arqrede, Left(objRecordset.Fields("OBJECT_ID_"), 6)
        'Rótulo[40]
        Print #arqrede, Left(objRecordset.Fields("OBJECT_ID_"), 40)
        'Compr. de cálculo (m)
        'Print #arqrede, CStr(IIf(objRecordset.Fields("LENGTH") = 0, objRecordset.Fields("LengthCalculated"), objRecordset.Fields("LENGTH")))
        QualCoordernada objRecordset.Fields("INITIALCOMPONENT"), xi, yi
        QualCoordernada objRecordset.Fields("FINALCOMPONENT"), xf, yf
        
        Print #arqrede, Round(Sqr(((xf - xi) * (xf - xi)) + ((xf - xi) * (xf - xi))), 2)
        'Diâmetro nominal (mm)
        Print #arqrede, CStr(objRecordset.Fields("INTERNALDIAMETER"))
        'Diâmetro externo real (m)
        Print #arqrede, CStr(objRecordset.Fields("INTERNALDIAMETER"))
        'Espessura da parede do tubo (mm)
        Print #arqrede, CStr(objRecordset.Fields("Thickness"))
        'Material
        Print #arqrede, Left(objRecordset.Fields("MATERIAL"), 30)
        'Superfície
        Print #arqrede, Left("Sem revestimeto", 30)
        'Escoramento
        Print #arqrede, Left("Sem escoramento", 30)
        'Compactação de aterro: compactado = -1; não compactado = 0
        Print #arqrede, "0"
        'Velocidade mínima (m/s)
        Print #arqrede, "0,5"
        'Velocidade máxima (m/s)
        Print #arqrede, "5"
        'Critério de interpolação de dados: global=0; cota dos nós=1; curvas de nível=2; pontos cotados=3
        Print #arqrede, "0"
        'Situação de rede
        Print #arqrede, Left("Rede existente", 30)
        objRecordset.MoveNext
    Wend
    Close #arqrede
    MsgBox "Exportação concluída"
    objRecordset.Close
    Set cgeo = Nothing
    Set objRecordset = Nothing
    ReDim aListaNo(0) As TListaNo
    Exit Sub
    

Else
Dim h As String
Dim i As String
Dim j As String
Dim l As String
Dim m As String
Dim n As String
Dim o As String
h = "MIN(X)"
i = "POINTS"
j = "object_id"
l = "MIN(Y)"
m = "MAX(X)"
n = "MAX(Y)"
o = "X_TempCalculeNode"


'Alterado dia 19/10/2010

sQry = convertQuery("SELECT " + """" + h + """" + " FROM " + """" + i & cgeo.GetLayerID(LayName) & " where " + """" + j + """" + " in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    minX = Int(objRecordset.Fields(0) / 100) * 100
    Print #arqrede, CStr(minX) * 100
    'Coordenada UTM Norte da origem da área de trabalho
    sQry = convertQuery("SELECT " + """" + l + """" + " FROM " + """" + i & cgeo.GetLayerID(LayName) & " where " + """" + j + """" + " in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    minY = Int(objRecordset.Fields(0) / 100) * 100
    Print #arqrede, CStr(minY) * 100
    'Comprimento na direção Este (cm)
    sQry = convertQuery("SELECT " + """" + m + """" + " from " + """" + l + """" + cgeo.GetLayerID(LayName) & " where " + """" + j + """" + " in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    maxX = Int(CSng(objRecordset.Fields(0) / 100) * 100 + 100)
    Print #arqrede, CStr(maxX - minX)
    'Comprimento na direção Norte (cm)
    sQry = convertQuery("SELECT " + """" + n + """" + " FROM " + """" + l & cgeo.GetLayerID(LayName) & " where " + """" + j + """" + " in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    maxY = Int(CSng(objRecordset.Fields(0) / 100) * 100 + 100)
    Print #arqrede, CStr(maxY - minY)
    'Espaçamento do grid em metros
    Print #arqrede, CalculaAreaTrab(maxX - minX, maxY - minY)
    'Espaçamento do grid em twips
    'talvez não seja necessário, mas aí vai um teste
    Print #arqrede, 13785
    'Quantidade de nós
    sQry = convertQuery("SELECT COUNT(*) FROM " + """" + l + """" + cgeo.GetLayerID(LayName) & " where " + """" + j + """" + " in(" & MyObject & ")", CInt(typeconnection))
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    Print #arqrede, CStr(objRecordset.Fields(0))
    ReDim aListaNo(1 To objRecordset.Fields(0))
    'Quatidade de trechos
  a = "LIN"
    sQry = "SELECT count(DISTINCT(" + """" + a + """" + ")) from " + """" + o + """"
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    'Print #arqrede, CStr(objRecordset.Fields(0))
    Print #arqrede, CStr(objRecordset.Fields(0))
    Dim p As String
     Dim q As String
      Dim s As String
       Dim t As String
        Dim u As String
         Dim v As String
         Dim sq As String
         Dim sa As String
          Dim r, X, z, xz, Y As String
         p = "y"
         q = "ID_TYPE"
        r = "OBJECT_ID_"
         t = "GROUNDHEIGHT"
         u = "Demand"
         v = "x"
         X = "points2"
         z = "WATERCOMPONENTS"
         xz = "object_id"
        sa = cgeo.GetLayerID(LayName)
        sq = "sa"
        
          sQry = convertQuery("SELECT " + """" + X + """" + "." + """" + Y + """" + "," + """" + X + """" + "." + """" + v + """" + "," + """" + z + """" + "." + """" + q + """" + "," + """" + X + """" + "." + """" + r + """" + "," + """" + z + """" + "." + """" + t + """" + "," + """" + X + """" + "." + """" + u + """" + " From " + """" + X + sq + """" + " LEFT JOIN " + """" + z + """" + " ON " + """" + xz + """" + "." + """" + p + """" + " = " + """" + z + """" + "." + """" + r + """" + " Where " + """" + xz + """" + " in(" & MyObject & ")", CInt(typeconnection))
    'pode está errado ********
        
        
 
    
    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
    iNo = 1
  
    While (Not objRecordset.EOF)
        'Nó
        Print #arqrede, "[NO]"
        'N em cm
        Print #arqrede, CStr(objRecordset.Fields("Y") * 100)
        'E em cm
        Print #arqrede, CStr(objRecordset.Fields("X") * 100)
        'Tipo
        Print #arqrede, CStr(CalculaSubType(objRecordset.Fields("id_TYPE")))
        'Rótulo[6]
        Print #arqrede, Left(objRecordset.Fields("OBJECT_ID"), 6)
        'Rótulo[20]
        Print #arqrede, Left(objRecordset.Fields("OBJECT_ID"), 20)
        'Valor característico
        'Print #arqrede, "0"
       
    
    
    'alterado manoel 19/10/2010
    Dim mn As String
    Dim ml As String
    Dim mm As String
    Dim mo As String
    Dim mp As String
    mn = "VALUE_"
    ml = "WATERCOMPONENTSDATA"
    mm = "ID_TYPE"
    mo = "ID_SUBTYPE"
    mp = "OBJECT_ID_"
    
    
    
 
    
    
    Select Case objRecordset.Fields("id_TYPE")
           Case 0, 2, 18, 22 ' demanda ou vazao pontual
                'Desconhecido=0; Hidrante=6; Curva=7;
                'Conexão T=8; Cruzeta=9; Tap=10; Redução=11;
                'Hidrômetro = 12; Consumidor=2
                Print #arqrede, objRecordset.Fields("Demand")
           Case 20 'Bomba=1 ou buster 'sentido VALOR
                Set RsCa = Conn.execute("SELECT " + """" + mn + """" + " from " + """" + ml + """" + " where  " + """" + mm + """" + " = '11'  and " + """" + mo + """" + " = '1' and " + """" + mp + """" + " = '" & objRecordset.Fields("OBJECT_ID") & "'")
                If Not RsCa.EOF Then
                   Print #arqrede, RsCa(0) 'objRecordset.Fields("GROUNDHEIGHTFINAL")
                Else
                   Print #arqrede, 0
                End If
                RsCa.Close
           Case 19 'Reservatorio=3 Cota do nival dagua
                Set RsCa = Conn.execute("SELECT " + """" + mn + """" + " from " + """" + ml + """" + " where  " + """" + mm + """" + " = '19'  and " + """" + mo + """" + " = '7' and " + """" + mp + """" + " = '" & objRecordset.Fields("OBJECT_ID") & "'")
                If Not RsCa.EOF Then
                   Print #arqrede, RsCa(0) 'objRecordset.Fields("GROUNDHEIGHTFINAL")
                Else
                   Print #arqrede, 0
                End If
                RsCa.Close
           Case 21 'VRP=4  Perda de Carga - UNIDADE(m)
                Print #arqrede, "0"
                'Print #arqrede, objRecordset.Fields("GROUNDHEIGHTFINAL")
           Case 1 'V Controle    vc = porcentagem de abertura
                Set RsCa = Conn.execute("SELECT " + """" + mn + """" + " from " + """" + ml + """" + " where  " + """" + mm + """" + " = '1'  and " + """" + mo + """" + " = '3' and " + """" + mp + """" + "='" & objRecordset.Fields("OBJECT_ID") & "'")
                If Not RsCa.EOF Then
                   Print #arqrede, RsCa(0) 'objRecordset.Fields("GROUNDHEIGHTFINAL")
                Else
                   Print #arqrede, 0
                End If
                RsCa.Close
           Case Else
                Print #arqrede, "0"
         End Select
        'Print #arqrede, ValorCaracterisco(objRecordset.Fields("id_TYPE"))
        'k
        Print #arqrede, "0"
        'Cota do terreno
        Print #arqrede, CStr(objRecordset.Fields("GROUNDHEIGHT"))
        'Cota do nó
        Print #arqrede, CStr(objRecordset.Fields("GROUNDHEIGHT") - 1)
        'Pressão mín admitida (mca)
        Print #arqrede, "15"
        'Pressão máx admitida (mca)
        Print #arqrede, "50"
        'Nome da lei de operação: pra compatibilidade com os outros formatos de cabecalho
        Print #arqrede, ""
        'Lei de especificação da condição de contorno
        Print #arqrede, "Default"
        'Nó de jusante
        Print #arqrede, "0"
        'Confirmação de cota
        Print #arqrede, "-1"
        'booster: Indicação de sentido, ou
        'vrp: cirtério de cálculo
        Print #arqrede, "0"
        'situação de rede
        Print #arqrede, "Rede existente"
        aListaNo(iNo).indice = iNo
        aListaNo(iNo).object_id = objRecordset.Fields("OBJECT_ID")
        aListaNo(iNo).X = CStr(objRecordset.Fields("X") * 100)
        aListaNo(iNo).Y = CStr(objRecordset.Fields("Y") * 100)
        iNo = iNo + 1
        objRecordset.MoveNext
    Wend
    Set RsCa = Nothing
  
    '19/10/2010 alterado manoel

a = "CPI"
b = "CPF"
c = "OBJECT_ID_"
d = "LENGTH"
e = "LengthCalculated"
f = "INTERNALDiameter"
g = "Thickness"
h = "MATERIAL"
i = "WATERLINES"
j = "X_TempCalculeNode"
k = "lin"
        
    
    
    
    
   
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -
    '                                     Trechos
    '- x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x - x -

    
   sQry = "SELECT DISTINCT " + """" + a + """" + " AS  INITIALCOMPONENT, " + """" + b + """" + " AS FINALCOMPONENT, " + """" + c + """" + ", " + """" + d + """" + ", " + """" + e + """" + "," + """" + f + """" + ", " + """" + g + """" + ", " + """" + h + """" + "FROM " + """" + i + """" + " INNER JOIN " + """" + j + """" + " on " + """" + k + """" + "=  " + """" + "OBJECT_ID_" + """" + ""
       
              


    objRecordset.Close
    objRecordset.Open sQry, _
                      Conn, _
                      adOpenKeyset, _
                      adLockOptimistic
                      
                      

    
    
    
    While (Not objRecordset.EOF)
        'Trecho
        Print #arqrede, "[TRECHO]"
        'Nó inicial
        Print #arqrede, CStr(QualIndiceNo(objRecordset.Fields("INITIALCOMPONENT")))
        'Nó final
        Print #arqrede, CStr(QualIndiceNo(objRecordset.Fields("FINALCOMPONENT")))
        'Rótulo[6]
        Print #arqrede, Left(objRecordset.Fields("OBJECT_ID_"), 6)
        'Rótulo[40]
        Print #arqrede, Left(objRecordset.Fields("OBJECT_ID_"), 40)
        'Compr. de cálculo (m)
        'Print #arqrede, CStr(IIf(objRecordset.Fields("LENGTH") = 0, objRecordset.Fields("LengthCalculated"), objRecordset.Fields("LENGTH")))
        QualCoordernada objRecordset.Fields("INITIALCOMPONENT"), xi, yi
        QualCoordernada objRecordset.Fields("FINALCOMPONENT"), xf, yf
        
        Print #arqrede, Round(Sqr(((xf - xi) * (xf - xi)) + ((xf - xi) * (xf - xi))), 2)
        'Diâmetro nominal (mm)
        Print #arqrede, CStr(objRecordset.Fields("INTERNALDIAMETER"))
        'Diâmetro externo real (m)
        Print #arqrede, CStr(objRecordset.Fields("INTERNALDIAMETER"))
        'Espessura da parede do tubo (mm)
        Print #arqrede, CStr(objRecordset.Fields("Thickness"))
        'Material
        Print #arqrede, Left(objRecordset.Fields("MATERIAL"), 30)
        'Superfície
        Print #arqrede, Left("Sem revestimeto", 30)
        'Escoramento
        Print #arqrede, Left("Sem escoramento", 30)
        'Compactação de aterro: compactado = -1; não compactado = 0
        Print #arqrede, "0"
        'Velocidade mínima (m/s)
        Print #arqrede, "0,5"
        'Velocidade máxima (m/s)
        Print #arqrede, "5"
        'Critério de interpolação de dados: global=0; cota dos nós=1; curvas de nível=2; pontos cotados=3
        Print #arqrede, "0"
        'Situação de rede
        Print #arqrede, Left("Rede existente", 30)
        objRecordset.MoveNext
    Wend
    Close #arqrede
    MsgBox "Exportação concluída"
    objRecordset.Close
    Set cgeo = Nothing
    Set objRecordset = Nothing
    ReDim aListaNo(0) As TListaNo
    Exit Sub
    End If
    
    
gtErro:
    Close #arqrede
    Set cgeo = Nothing
    MsgBox Err.Description, vbExclamation, "GeoSan"
End Sub
'até aqui dia 19/10/2010

Function Log10(X)
    Log10 = Log(X) / Log(10)
End Function


Private Function CalculaSubType(tipo As Long) As Long
'Tipologia obtida de TeConCanvas.idl
    Select Case tipo
        Case 0, 2, 18
            'Desconhecido=0; Hidrante=6; Curva=7;
            'Conexão T=8; Cruzeta=9; Tap=10; Redução=11;
            'Hidrômetro = 12
            CalculaSubType = 0
        Case 20 'Bomba=1
            CalculaSubType = 1
        Case 2 'Consumidor=2
            CalculaSubType = 5
        Case 19 'Reservatorio=3
            CalculaSubType = 2
        Case 21 'VRP=4
            CalculaSubType = 4
        Case 1 'Valvula=5
            CalculaSubType = 3
    End Select
End Function

Private Function ValorCaracterisco(tipo As Long) As Long
'Tipologia obtida de TeConCanvas.idl
    Select Case tipo
        Case 0, 2, 18 ' demanda ou vazao pontual
            'Desconhecido=0; Hidrante=6; Curva=7;
            'Conexão T=8; Cruzeta=9; Tap=10; Redução=11;
            'Hidrômetro = 12; Consumidor=2
            ValorCaracterisco = 11
        Case 20 'Bomba=1 ou buster 'sentido
            ValorCaracterisco = 1
        Case 19 'Reservatorio=3 Cota do nival dagua
            ValorCaracterisco = 2
        Case 21 'VRP=4  Perda de Carga
            ValorCaracterisco = 4
        Case 1 'V Controle    vc = porcentagem de abertura
            ValorCaracterisco = 3
        Case Else
            ValorCaracterisco = 0
    End Select
End Function



Function CalculaAreaTrab(DeltaE As Long, DeltaN As Long) As Long
Dim GridTeste As Long
Dim OrdemGrand As Long
Dim Resp
    If DeltaE < DeltaN Then
        OrdemGrand = 10 ^ Int(Log10(DeltaE))
        GridTeste = (DeltaE \ OrdemGrand) * (OrdemGrand / 10)
    Else
        OrdemGrand = 10 ^ Int(Log10(DeltaN))
        GridTeste = (DeltaN \ OrdemGrand) * (OrdemGrand / 10)
    End If
    If GridTeste < 100 Then GridTeste = 100
    CalculaAreaTrab = GridTeste
End Function

Private Function QualIndiceNo(object_id) As Long
Dim i As Long
    For i = LBound(aListaNo) To UBound(aListaNo)
        If aListaNo(i).object_id = object_id Then
            QualIndiceNo = aListaNo(i).indice
            Exit Function
        End If
    Next i
    MsgBox " NÃO ENCONTROU"
End Function

Private Function QualCoordernada(object_id, ByRef X As Double, ByRef Y As Double) As Long
Dim i As Long
    For i = LBound(aListaNo) To UBound(aListaNo)
        If aListaNo(i).object_id = object_id Then
            X = aListaNo(i).X
            Y = aListaNo(i).Y
            Exit Function
        End If
    Next i
End Function

Public Function OpenReport(LayerName As String)
   GeraRelatorioHtm RedeMaterialDiametro, LayerName
End Function

Public Function GeraRelatorioHtm(tipo As TipoRelatorio, Layer As String, Optional Filtro As Boolean) As Boolean
   Select Case tipo
      Case RedeMaterialDiametro
         GeraRelatorioHtm_RedeMaterialDiametro Layer
      Case RegistrosEstadoEstado
         GeraRelatorioHtm_RegistrosLocalizacaoEstado
      Case ComponentsRede
         GeraRelatorioHtm_ComponentsRede Layer, Filtro
   End Select

End Function
' Função para gerar relatório de redes
'
' LayerName - nome do layer em que será gerado o relatório
'
Private Function GeraRelatorioHtm_RedeMaterialDiametro(LayerName As String)
    On Error GoTo Trata_Erro:
    Dim rs As ADODB.Recordset, Material As String
    Dim CompCalc As Double, ComplMed As Double, Qtde As Long
    Dim CompCalcTot As Double, ComplMedTot As Double, QtdeTot As Long
    Dim aa As String
    Dim bb As String
    Dim cc As String
    Dim dd As String
    Dim ee As String
    Dim ff As String
    Dim gg As String
    Dim hh As String
    Dim ii As String
    Dim ll As String
    Dim sPathUser As String                 'caminho do diretório do usuário em My Documents
    Dim mensagem As String                  'mensagem de erro, caso ocorra
    Dim diret As New CEncontraDiretorio     'para localizar um determinado diretório do Windows
    
    'diret = New CEncontraDiretorio          'inicializa o objeto para localizar o diretório de documentos do usuário
   aa = "MATERIALNAME"
   bb = "INTERNALDIAMETER"
   cc = "LENGTH"
   dd = "LENGTHCALCULATED"
   ee = UCase(LayerName)
 
   gg = "X_MATERIAL"
   hh = "MATERIALID"
   ii = "MATERIAL"
   a = "LINES"
   
   
    If frmCanvas.TipoConexao <> 4 Then
   
   Set rs = Conn.execute("SELECT  MaterialName, internalDiameter,count(*) as " + """" + "Qtde" + """" + " ,sum(length) as " + """" + "Compl" + """" + ", sum(lengthcalculated) as " + """" + "complCalc" + """" + "" & _
              " from " & LayerName & "lines Left Join X_material on Material = Materialid Group by MaterialName, internalDiameter,MATERIAL " & _
              "Order by MaterialName, InternalDiameter")
              Else
              Set rs = Conn.execute("SELECT  " + """" + aa + """" + ", " + """" + bb + """" + ",count(*) as " + """" + "Qtde" + """" + ",sum(" + """" + cc + """" + ") as " + """" + "Compl" + """" + ", sum(" + """" + dd + """" + ") as " + """" + "complCalc" + """" + "" & _
              " from " + """" + ee + a + """" + " Left Join " + """" + gg + """" + "on " + """" + ii + """" + " = " + """" + hh + """" + " Group by " + """" + aa + """" + ", " + """" + bb + """" + "," + """" + ii + """" + " " & _
              "Order by " + """" + aa + """" + ", " + """" + bb + """" + "")
              End If
   Dim str As String
   
   
   
   str = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
   str = str & "<HTML><HEAD>"
   str = str & "<META http-equiv=Content-Type content=text/html; charset=ISO-8859-1/>"
   str = str & "<META content='MSHTML 6.00.6000.16481' name=GENERATOR></HEAD>"
   str = str & "<BODY>"
   str = str & "<TABLE style='WIDTH: 781px; HEIGHT: 103px' cellSpacing=1 cellPadding=1 width=781 border=1 id=TABLE1>"
   str = str & "  <TR>"
   str = str & "    <TD>"
   str = str & "      <P align=center><FONT color=blue size=6><EM><STRONG>NEXUS</STRONG></EM></FONT></P></TD>"
   str = str & "   <TD COLSPAN=3><P align=center><FONT size=3>Relatório Rede de " & IIf(LayerName = "water", "Água", "Esgoto") & " -  Materiais por Diametro</FONT></P></TD>"
   str = str & "    <TD>"
   str = str & "      <P align=center><FONT size=3>" & Date & "</FONT></P></TD></TR>"
   str = str & ""
   While Not rs.EOF
      If Material <> rs!MATERIALNAME Then
         str = str & "  <TR>"
         str = str & "    <TD COLSPAN=2>Material: " & UCase(rs!MATERIALNAME) & "</TD>"
         str = str & "    <TD COLSPAN=3>"
         str = str & "      <P align=center>Comprimento em Metros</P>  </TD></TR>"
         str = str & "  <TR>"
         str = str & "   <TD WIDTH='20%'></TD>"
         str = str & "    <TD WIDTH='20%'>Diâmetro</TD>"
         str = str & "    <TD WIDTH='20%'>Qdte Tubos</TD>"
         str = str & "    <TD WIDTH='20%'>Medido</TD>"
         str = str & "    <TD WIDTH='20%'>Calculado</TD></TR>"
      End If
      str = str & "  <TR>"
      str = str & "    <TD></TD>"
      str = str & "    <TD>"
      str = str & "      <P align=center>" & rs!INTERNALDIAMETER & "</P></TD>"
      str = str & "    <TD>"
      str = str & "      <P align=center>&nbsp;" & rs!Qtde & "</P></TD>"
      str = str & "    <TD>"
      str = str & "      <P align=center>" & rs!Compl & "</P></TD>"
      str = str & "    <TD>"
      str = str & "      <P align=center>" & Round(rs!complCalc, 2) & "</P></TD>"
      str = str & "     "
      str = str & "  </TR>"
      Material = IIf(IsNull(rs!MATERIALNAME), "", rs!MATERIALNAME)
      CompCalc = CompCalc + Round(rs!complCalc)
      ComplMed = ComplMed + rs!Compl
      Qtde = Qtde + rs!Qtde
      rs.MoveNext
      If rs.EOF Then
         str = str & "  <TR>"
         str = str & "   <TD WIDTH='20%'>Total</TD>"
         str = str & "    <TD WIDTH='20%'>"
         str = str & "      <P align=center></P></TD>"
         str = str & "    <TD WIDTH='20%'>"
         str = str & "      <P align=center>" & Qtde & "</P> </TD>"
         str = str & "    <TD WIDTH='20%'>"
         str = str & "      <P align=center>" & ComplMed & "</P></TD>"
         str = str & "    <TD WIDTH='20%'>"
         str = str & "      <P align=center>" & CompCalc & "</P></TD></TR>"
         str = str & "  <TR></TR>"
         CompCalcTot = CompCalcTot + CompCalc
         ComplMedTot = ComplMedTot + ComplMed
         QtdeTot = QtdeTot + Qtde
         Material = 0
         CompCalc = 0
         ComplMed = 0
         Qtde = 0
      ElseIf Material <> rs!MATERIALNAME Then
         str = str & "  <TR>"
         str = str & "   <TD WIDTH='20%'>Total</TD>"
         str = str & "    <TD WIDTH='20%'>"
         str = str & "      <P align=center></P></TD>"
         str = str & "    <TD WIDTH='20%'>"
         str = str & "      <P align=center>" & Qtde & "</P> </TD>"
         str = str & "    <TD WIDTH='20%'>"
         str = str & "      <P align=center>" & ComplMed & "</P></TD>"
         str = str & "    <TD WIDTH='20%'>"
         str = str & "      <P align=center>" & CompCalc & "</P></TD></TR>"
         str = str & "  <TR></TR>"
         CompCalcTot = CompCalcTot + CompCalc
         ComplMedTot = ComplMedTot + ComplMed
         QtdeTot = QtdeTot + Qtde
         Material = 0
         CompCalc = 0
         ComplMed = 0
         Qtde = 0
      End If
   Wend
   str = str & "</TABLE>"
   str = str & "<HR align=left style='WIDTH: 781px; HEIGHT: 2px' SIZE=2>"
   str = str & "<TABLE style='WIDTH: 781px; HEIGHT: 32px' cellSpacing=1 cellPadding=1 width=781 border=1 id=TABLE1>"
   str = str & "  <TR>"
   str = str & "   <TD WIDTH='20%'>TOTAL GERAL</TD>"
   str = str & "    <TD WIDTH='20%'></TD>"
   str = str & "    <TD WIDTH='20%'>" & QtdeTot & "</TD>"
   str = str & "    <TD WIDTH='20%'>" & ComplMedTot & "</TD>"
   str = str & "    <TD WIDTH='20%'>" & CompCalcTot & "</TD></TR>"
   str = str & "</TABLE>"
   str = str & ""
   str = str & "</BODY>"
   str = str & "</HTML>"
   sPathUser = diret.ObtemDiretorio(CSIDL_PERSONAL) + "RelatorioRede.htm"           'obtem o diretório de documentos do usuário
   mensagem = "Caminho do relatório: " & sPathUser
   Open sPathUser For Output As #1
   Print #1, str
   Close #1
   rs.Close
   Set rs = Nothing
   AbrirArquivo.Abre (sPathUser)
   Exit Function

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "Global", "GeraRelatorioHtm_RedeMaterialDiametro", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails, mensagem
    End If
End Function
' Função para gerar relatórios de registros cadastrados
'
'
Private Function GeraRelatorioHtm_RegistrosLocalizacaoEstado()
    Dim rs As ADODB.Recordset, Material As String
    Dim TotalComp As Integer
    Dim str As String
    Dim COD_VRP As Integer
    'alterei em 19/10/2010 manoel
    Dim ff As String
    Dim vv As String
    Dim xx As String
    Dim hh As String
    Dim oo As String
    Dim rr As String
    Dim ss As String
    Dim ww As String
    Dim jj As String
    Dim gg As String
    Dim tt As String
    Dim nn As String
    Dim uu As String
    Dim qq As String
    Dim pp As String
    Dim zz As String
    Dim sPathUser As String                     ' caminho do diretório do usuário em My Documents
    Dim mensagem As String                      'mensagem de erro, caso ocorra
    Dim diret As New CEncontraDiretorio         'para localizar um determinado diretório do Windows
    'Dim ss As String
   
   'PEGA O CÓDIGO IDENTIFICADOR
    If frmCanvas.TipoConexao <> 4 Then
   Set rs = New ADODB.Recordset
   rs.Open ("SELECT * FROM WATERCOMPONENTSTYPES WHERE DESCRIPTION_ = 'VRP'"), Conn, adOpenDynamic, adLockReadOnly
   If rs.EOF = False Then
      COD_VRP = rs!id_Type
   End If
   
   rs.Close
   
   
   str = "SELECT x.stateName as " + """" + "Estado" + """" + ", "
   str = str & "    l.LocationName as " + """" + "Localizacao" + """" + ","
   str = str & "    CASE WHEN d.value_ = 0 THEN 'Desconhecido'"
   str = str & "         WHEN d.value_ = 1 THEN 'Aberto'"
   str = str & "         WHEN d.value_ = 2 THEN 'Fechado'"
   str = str & "         Else 'Desconhecido' END Abertura ,"
   str = str & "    count(*) as " + """" + "Qtde" + """" + ""
   str = str & " from watercomponents w "
   str = str & "    left join x_state x on x.stateid=w.state "
   str = str & "    left join x_Location l on l.locationid= w.location "
   str = str & "    left join watercomponentsdata d on d.object_id_= w.object_id_ and d.id_type = w.id_type And d.Id_SubType = 2 "
   str = str & " Where w.id_Type = " & COD_VRP & " group by x.statename,l.LocationName,d.value_ "
   str = str & " order by x.statename,l.LocationName,d.value_ "
   Else
   Set rs = New ADODB.Recordset
   a = "WATERCOMPONENTSTYPES"
      b = "DESCRIPTION_"
   rs.Open ("SELECT * FROM " + """" + a + """" + " WHERE " + """" + b + """" + " = 'VRP'"), Conn, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      COD_VRP = rs!id_Type
   End If
   ff = "WATERCOMPONENTSTYPES"
    xx = "STATENAME"
    hh = "VALUE_"
    oo = "WATERCOMPONENTS"
    rr = "X_STATE"
    ss = "STATEID"
    ww = "STATE"
    jj = "X_LOCATION"
    gg = "LOCATIONID"
    tt = "LOCATION"
    nn = "WATERCOMPONENTSDATA"
    uu = "OBJECT_ID_"
    qq = "ID_TYPE"
    pp = "ID_SUBTYPE"
    zz = "LOCATIONNAME"
   
   
   str = "SELECT " + """" + rr + """" + "." + """" + xx + """" + " as " + """" + "Estado" + """" + ", "
   str = str & "    " + """" + jj + """" + "." + """" + zz + """" + " as " + """" + "Localizacao" + """" + ","
   str = str & "    CASE WHEN " + """" + nn + """" + "." + """" + hh + """" + " = '0' THEN 'Desconhecido'"
   str = str & "         WHEN " + """" + nn + """" + "." + """" + hh + """" + " = '1' THEN 'Aberto'"
   str = str & "         WHEN " + """" + nn + """" + "." + """" + hh + """" + " = '2' THEN 'Fechado'"
   str = str & "         Else 'Desconhecido' END Abertura ,"
   str = str & "    count(*) as " + """" + "Qtde" + """" + " "
   str = str & " from " + """" + oo + """" + " "
   str = str & "    left join " + """" + rr + """" + " on " + """" + ss + """" + "=" + """" + oo + """" + "." + """" + ww + """" + " "
   str = str & "    left join " + """" + jj + """" + " on " + """" + jj + """" + "." + """" + gg + """" + "= " + """" + oo + """" + "." + """" + tt + """" + " "
   str = str & "    left join " + """" + nn + """" + " on " + """" + nn + """" + "." + """" + uu + """" + "= " + """" + oo + """" + "." + """" + uu + """" + "and " + """" + nn + """" + "." + """" + qq + """" + " = " + """" + oo + """" + "." + """" + qq + """" + " And " + """" + nn + """" + "." + """" + pp + """" + " = '2' "
   str = str & " Where " + """" + oo + """" + "." + """" + qq + """" + " = '" & COD_VRP & "' group by " + """" + xx + """" + "," + """" + ss + """" + "," + """" + hh + """" + "," + """" + jj + """" + "." + """" + zz + """" + " "
   str = str & " order by " + """" + rr + """" + "." + """" + xx + """" + "," + """" + jj + """" + "." + """" + zz + """" + "," + """" + nn + """" + "." + """" + hh + """" + " "
   'alterado  até aqui em 19/10/2010
   End If
   
   ' WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
   Set rs = Conn.execute(str)
   str = ""
   str = str & "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
   str = str & "<HTML><HEAD>"
   str = str & "<META http-equiv=Content-Type content='text/html; charset=charset=ISO-8859-1'>"
   str = str & "<META content='MSHTML 6.00.6000.16481' name=GENERATOR></HEAD>"
   str = str & "<BODY>"
   str = str & "<TABLE style='WIDTH: 781px; HEIGHT: 103px' cellSpacing=1 cellPadding=1 width=781 border=1 id=TABLE1>"
   str = str & "<TR>"
   str = str & " <TD><P align=center><FONT color=blue size=6><EM><STRONG>NEXUS</STRONG></EM></FONT></P></TD>"
   str = str & "   <TD COLSPAN=2><P align=center><FONT size=4>Relatorio de Valvulas(VRP) da Rede</FONT></P></TD>"
   str = str & "    <TD><P align=center><FONT size=1>" & Date & "</FONT></P></TD></TR>"
   str = str & "  <TR>"
   str = str & "    <TD width='25%'><P align=center><FONT color=black><STRONG>Estado</STRONG></FONT></P></TD>"
   str = str & " <TD width='25%'><P align=center><FONT color=black><STRONG>Localizacao</STRONG></FONT></P></TD>"
   str = str & "    <TD width='25%'><P align=center><FONT color=black><STRONG>Aberto/Fechado</STRONG></FONT></P></TD>"
   str = str & " <TD width='25%'><P align=center><FONT color=black><STRONG>Quantidade</STRONG></FONT></P></TD></TR>"
   While Not rs.EOF
      str = str & "<TR>"
      str = str & " <TD><P align=center><FONT color=black>" & rs!Estado & "</FONT></P></TD>"
      str = str & " <TD><P align=center><FONT color=black>" & rs!Localizacao & "</FONT></P></TD>"
      str = str & " <TD><P align=center><FONT color=black>" & rs!Abertura & "</FONT></P></TD>"
      str = str & " <TD><P align=center><FONT color=black>" & rs!Qtde & "</FONT></P></TD></TR>"
      TotalComp = TotalComp + rs!Qtde
      rs.MoveNext
   Wend
   rs.Close
   Set rs = Nothing
   str = str & "<HR align=left style='WIDTH: 781px; HEIGHT: 2px' SIZE=2></TABLE>"
   str = str & "<TABLE style='WIDTH: 782px; HEIGHT: 12px' cellSpacing=1 cellPadding=1 width=782 border=1 id=TABLE2>"
   str = str & "  <TR>"
   str = str & "    <TD><P align=center><FONT color=black>TOTAL</FONT></P></TD>"
   str = str & "    <TD><P align=center><FONT color=black></FONT></P></TD>"
   str = str & "    <TD><P align=center><FONT color=black></FONT></P></TD>"
   str = str & "    <TD><P align=center><FONT color=black>" & TotalComp & "</FONT></P></TD></TR></TABLE>"
   str = str & "</BODY>"
   str = str & "</HTML>"
   sPathUser = diret.ObtemDiretorio(CSIDL_PERSONAL) + "RelatorioRegistros.htm"           'obtem o diretório de documentos do usuário
   mensagem = "Caminho do relatório: " & sPathUser
   Open sPathUser For Output As #1
   Print #1, str
   Close #1
   AbrirArquivo.Abre (sPathUser)
   Exit Function
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "Global", "GeraRelatorioHtm_RegistrosLocalizacaoEstado", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails, mensagem
    End If
End Function
' Gera relatório de todas componentes de uma rede
'
' LayerName - nome do layer em que será gerado o relatório
' Filtro -
'
Public Function GeraRelatorioHtm_ComponentsRede(LayerName As String, Optional Filtro As Boolean)
    On Error GoTo Trata_Erro:
    Dim rs As ADODB.Recordset, Material As String
    Dim TotalComp As Long
    Dim str As String
    Dim sPathUser As String                     ' caminho do diretório do usuário em My Documents
    Dim mensagem As String                      'mensagem de erro, caso ocorra
    Dim diret As New CEncontraDiretorio         'para localizar um determinado diretório do Windows
    
    If frmCanvas.TipoConexao <> 4 Then
   
   blnGeraRel = True
   
   str = "SELECT CASE " & _
       " WHEN Description_ IS NULL THEN 'Desconhecido' " & _
       " Else Description_ " & _
       " END TIPO, " & _
       " count(*) as " + """" + "QTDE" + """" + "" & _
       " From " & LayerName & " c " & _
       " left join " & LayerName & "Types t on t.id_type=c.id_type "
   If Filtro Then
      str = str & frmFilterReport.Init()
   End If
   str = str & " Group By Description_ "
   
   If blnGeraRel = False Then
      Exit Function
   End If
   
   'alterado em 19/10/2010
   Dim iu As String
   Dim io As String
   Dim ia As String
   Dim ie As String
   Dim ii As String
    Dim iy As String
   iu = "CASE"
   io = "DESCRIPTION_"
   ia = "TIPO"
   ie = UCase(LayerName)
   ii = "TYPES"
   iy = "ID_TYPE"
   Else
   blnGeraRel = True
   
   'SELECT CASE WHEN "DESCRIPTION_"IS NULL THEN 'Desconhecido'  Else "DESCRIPTION_" END "TIPO", count(*) as "QTDE"
 'From "WATERCOMPONENTS"   left join "WATERCOMPONENTSTYPES" on   "WATERCOMPONENTSTYPES"."ID_TYPE"="WATERCOMPONENTS"."ID_TYPE"
 'group by "DESCRIPTION_"
 
   str = "SELECT CASE WHEN " + """" + "DESCRIPTION_" + """" + "IS NULL THEN 'Desconhecido'  Else " + """" + "DESCRIPTION_" + """" + " END " + """" + "TIPO" + """" + ", count(*) as " + """" + "QTDE" + """" + "   From " + """" + UCase(LayerName) + """" + "   left join " + """" + UCase(LayerName) + "TYPES" + """" + " on   " + """" + UCase(LayerName) + "TYPES" + """" + "." + """" + "ID_TYPE" + """" + "=" + """" + UCase(LayerName) + """" + "." + """" + "ID_TYPE" + """" + "group by " + """" + "DESCRIPTION_" + """" + ""
   ' WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
  ' MsgBox str
  
  If frmCanvas.TipoConexao <> 4 Then
  
   If Filtro Then
      str = str & frmFilterReport.Init()
      
      
   End If
   End If
   
   If frmCanvas.TipoConexao = 4 Then
  
   If Filtro Then
     
      
       str = "SELECT CASE WHEN " + """" + "DESCRIPTION_" + """" + "IS NULL THEN 'Desconhecido'  Else " + """" + "DESCRIPTION_" + """" + " END " + """" + "TIPO" + """" + ", count(*) as " + """" + "QTDE" + """" + "   From " + """" + UCase(LayerName) + """" + "   left join " + """" + UCase(LayerName) + "TYPES" + """" + " on   " + """" + UCase(LayerName) + "TYPES" + """" + "." + """" + "ID_TYPE" + """" + "=" + """" + UCase(LayerName) + """" + "." + """" + "ID_TYPE" + """" + frmFilterReport.Init() + " group by " + """" + "DESCRIPTION_" + """" + ""
   'MsgBox str
   End If
   End If
   
   
   
   If blnGeraRel = False Then
      Exit Function
   End If
   End If
   
   Set rs = Conn.execute(str)
   
   str = ""
   str = str & "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>"
   str = str & "<html xmlns='http://www.w3.org/1999/xhtml' >"
   str = str & "<head>"
   str = str & "<title>Untitled Page</title>"
   str = str & "</head>"
   str = str & "<body style='font-weight: bold; text-align: left'>"
   str = str & "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;"
   str = str & "&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;"
   str = str & "Relátório de componentes de rede de " & IIf(LayerName = "watercomponents", "Água", "Esgoto") & "<br />"
   str = str & "<hr style='width: 591px' />"
   str = str & "<table style='width: 589px'>"
   str = str & "<tr>"
   str = str & "<td style='font-weight: bold; width: 283px'>"
   str = str & "Tipo</td>"
   str = str & "<td style='font-weight: bold; width: 280px'>"
   str = str & "Quantidade</td>"
   str = str & "</tr>"
   While Not rs.EOF
      str = str & "<tr>"
      str = str & "<td style='width: 283px'>" & rs!tipo
      str = str & "</td>"
      str = str & "<td style='width: 280px'>" & rs!Qtde
      str = str & "</td>"
      TotalComp = TotalComp + rs!Qtde
      rs.MoveNext
   Wend
   str = str & "</tr>"
   str = str & "</table>"
   str = str & "<br />"
   str = str & "Total Componentes: &nbsp; &nbsp;&nbsp; " & TotalComp
   str = str & "</body>"
   str = str & "</html>"
    sPathUser = diret.ObtemDiretorio(CSIDL_PERSONAL) + "RelatorioComponentes.htm"           'obtem o diretório de documentos do usuário
    mensagem = "Caminho do relatório: " & sPathUser
    Open sPathUser For Output As #1
    Print #1, str
    Close #1
    AbrirArquivo.Abre (sPathUser)
    Exit Function
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "Global", "GeraRelatorioHtm_ComponentsRede", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails, mensagem
    End If
End Function



Public Function convertQuery(SQL As String, tipo As Integer) As String
   If tipo = 2 Then
      SQL = Replace(SQL, "[", "")
      SQL = Replace(SQL, "]", "")
   End If
   convertQuery = SQL
End Function

Public Function RetornaNomeMes(Mes As Integer, Optional nomeCompleto As Boolean = False) As String
   Select Case Mes
      Case 1
         RetornaNomeMes = IIf(nomeCompleto, "janeiro", "jan")
      Case 2
         RetornaNomeMes = IIf(nomeCompleto, "fevereiro", "fev")
      Case 3
         RetornaNomeMes = IIf(nomeCompleto, "março", "mar")
      Case 4
         RetornaNomeMes = IIf(nomeCompleto, "abril", "abr")
      Case 5
         RetornaNomeMes = IIf(nomeCompleto, "maio", "mai")
      Case 6
         RetornaNomeMes = IIf(nomeCompleto, "junho", "jun")
      Case 7
         RetornaNomeMes = IIf(nomeCompleto, "julho", "jul")
      Case 8
         RetornaNomeMes = IIf(nomeCompleto, "agosto", "ago")
      Case 9
         RetornaNomeMes = IIf(nomeCompleto, "setembro", "set")
      Case 10
         RetornaNomeMes = IIf(nomeCompleto, "outubro", "out")
      Case 11
         RetornaNomeMes = IIf(nomeCompleto, "novembro", "nov")
      Case 12
         RetornaNomeMes = IIf(nomeCompleto, "dezembro", "dez")
   End Select
End Function
'Procura no banco de dados a querie correspondente ao ID fornecido
'
'query_id           - número da querie que será procurada na tabela GS_QUERYS_CLIENT
'GetQueryProcess    - retorna a string da querie
'
Public Function GetQueryProcess(query_id As Integer) As String
    Dim rs As Recordset
    Dim zx As String
    Dim zy As String
    Dim zu As String
    Dim zo As String
    
    zx = "QUERYSTRING"
    zy = "GS_QUERYS_CLIENT"
    zu = "QUERY_ID"
    If frmCanvas.TipoConexao <> 4 Then
        'Conexão com bancos SQLServer ou Oracle
        Set rs = Conn.execute("SELECT querystring from gs_querys_client where query_id=" & query_id) '& " and client_id=" & client_id)
        If rs.EOF = False Then
            GetQueryProcess = rs.Fields("querystring").value
        Else
            MsgBox "Não há a QUERY_ID n." & query_id & " na tabela GS_QUERYS_CLIENT.", vbInformation, "Falta de registro detectada"
        End If
    Else
        'conexão com Postgres
        'alterado em 19/10/2010
        Set rs = Conn.execute("SELECT " + """" + zx + """" + " from " + """" + zy + """" + " where " + """" + zu + """" + "='" & query_id & "'")
        If rs.EOF = False Then
           GetQueryProcess = rs.Fields("querystring").value
        Else
           MsgBox "Não há a QUERY_ID n." & query_id & " na tabela GS_QUERYS_CLIENT.", vbInformation, "Falta de registro detectada"
        End If
    End If
    rs.Close
    'MsgBox GetQueryProcess - simplesmente para mostrar a querie lida no banco de dados para debug
    Set rs = Nothing
End Function


Public Function VerificaSeNumerico(ByVal ColunaNome As String, ByVal TabelaNome As String) As Boolean

   'VERIFICA SE O CAMPO DE UMA DETERMINADA TABELA 'E NUMÉRICO
   'RETORNA FALSE SE ALGUM CAMPO 'E VAZIO OU TEXTO
   Dim strsql As String
   Dim aa As String
   Dim dd As String
   Dim xx As String
   

   aa = ColunaNome
  dd = TabelaNome
  
   
    'alterado em 19/10/2010
     If frmCanvas.TipoConexao <> 4 Then
   strsql = "SELECT " & ColunaNome & " FROM " & TabelaNome & " WHERE ISNUMERIC(" & ColunaNome & ") = 0"
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   rs.Open strsql, Conn, adOpenForwardOnly, adLockReadOnly
   
   If rs.EOF = False Then
      VerificaSeNumerico = False
   Else
      VerificaSeNumerico = True
   End If
   
    Else
    
    
    strsql = "SELECT " + """" + ColunaNome + """" + " FROM " + """" + dd + """" + " WHERE isnumeric(" + """" + ColunaNome + """" + ") = '0'"


   Set rs = New ADODB.Recordset
   rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
   
   If rs.EOF = False Then
      VerificaSeNumerico = False
   Else
      VerificaSeNumerico = True
   End If
    
    End If
   rs.Close
   
End Function
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

