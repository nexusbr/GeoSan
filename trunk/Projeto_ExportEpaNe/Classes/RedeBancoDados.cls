VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RedeBancoDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_oConexao As New ADODB.Connection                          'conexão com o banco de dados para poder ler as linhas das tabelas do mesmo
Private rs_trechosDeRede As New ADODB.Recordset                     'lista de trechos de rede que foram selecionados pela cerca pelo usuário do GeoSan para serem exportados para o Epanet
Private usuarioQueExportouRedes As String                           'nome do usuário que exportou as redes do GeoSan
'Recebe a conexão com o banco de dados a partir de instanciada a classe e passada a conexão para a mesma, armazenando a conexão até o final da aplicação
'Isto permitirá todas as operações com banco de dados
'
'
'
Property Set Conexao(oConexao As ADODB.Connection)
    Set m_oConexao = oConexao
End Property

'Abre a conexão com os trechos de rede que serão exportados para o EPANET
'
'
Function IniciaLeituraTrechosRede()
    Dim str As String
    
    'Agora temos que descobrir todos os trechos de rede (TIPO=1 na tabela POLIGONO_SELECAO) que serão exportados para o Epanet, do usuário ativo
    str = "SELECT * FROM WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID"
    str = str & " WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & usuarioQueExportouRedes & "' AND TIPO = 1)"
    
    'Agora que temos trechos a serem exportados, vamos exportar para o Epanet
    rs_trechosDeRede.Open str, m_oConexao, adOpenDynamic, adLockReadOnly
    conn.CommandTimeout = 300
End Function

'Obtem a quantidade total de trechos de rede que serão processados
'
'
Function ObtemNumeroTrechosQueSeraoExportados() As Integer
    Dim Rs As ADODB.Recordset
    Dim strtot As String                        'armazena a querie para obter o número total de segmentos de rede que serão exportados para o Epanet (TIPO=1 na tabela POLIGONO_SELECAO)
    Dim totalTrechosExportar As Integer         'número total de trechos de rede de água que serão exportados para o Epanet, disponíveis em (TIPO=1 na tabela POLIGONO_SELECAO)
    
    'Agora temos que descobrir todos os trechos de rede (TIPO=1 na tabela POLIGONO_SELECAO) que serão exportados para o Epanet, do usuário ativo
    strtot = "SELECT COUNT(*) FROM WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID"
    strtot = strtot & " WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & usuarioQueExportouRedes & "' AND TIPO = 1)"
    Set Rs = New ADODB.Recordset
    'Obtem o número total de trechos de rede a serem exportados para o Epanet os quais são do TIPO=1 na tabela POLIGONO_SELECAO
    Rs.Open strtot, m_oConexao, adOpenDynamic, adLockReadOnly
    totalTrechosExportar = Rs(0).Value      'obtem o número total de trechos de rede que serão exportados
    Rs.Close
    Set Rs = Nothing
    ObtemNumeroTrechosQueSeraoExportados = totalTrechosExportar
End Function

'Obtem o nome do usuário que está exportando para o EPANET
'
'
Function ObtemNomeUsuario()
    Dim retval As String
    Dim usuario As String

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
    usuarioQueExportouRedes = usuario
End Function

'Atualiza todas as rugosidades de todas as tubulações, conforme o tipo de material. Foi considerada uma tubulação de 20 anos de idade
'Sempre que alguma linha for alterada na tabela X_Material, estas rugosidades deverão ser revistas
'
'
'
Public Function AtualizaAsRugosidades()
    'Liga a ampulheta no ponteiro do mouse
    MousePointer = vbHourglass
    
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 0")                                            'zera todas as rugozidades
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 111 WHERE MATERIAL = 0")                       'Desconhecido
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 1 AND ROUGHNESS = 0")     'Cerâmico
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 120 WHERE MATERIAL = 2 AND ROUGHNESS = 0")     'Concreto
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 110 WHERE MATERIAL = 3 AND ROUGHNESS = 0")     'FF
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 105 WHERE MATERIAL = 4 AND ROUGHNESS = 0")     'Fibrocimento
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 90  WHERE MATERIAL = 5 AND ROUGHNESS = 0")     'PEAD
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 6 AND ROUGHNESS = 0")     'PRSV
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 7 AND ROUGHNESS = 0")     'PVC branco
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 8 AND ROUGHNESS = 0")     'PVC corrugado
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 9 AND ROUGHNESS = 0")     'PVC DeFoFo
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 10 AND ROUGHNESS = 0")    'PVC Ocre
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 11 AND ROUGHNESS = 0")    'PVC PBA
    
    'Volta o mouse para o normal
    FrmEPANET.MousePointer = vbDefault
End Function

'Zera todos os materiais onde os mesmos estiverem com valor nulo
'
'
'
Public Function ZeraMateriais()
    
    m_oConexao.Execute ("UPDATE WATERLINES SET MATERIAL = 0 WHERE MATERIAL IS NULL")                        'Zera todos os materiais de tubulações quando o mesmo não estiver cadastrado
End Function
