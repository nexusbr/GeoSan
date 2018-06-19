VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeosanExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 1
Attribute VB_Description = "Exporta os dados do GeoSan para outros sistemas"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Cria a tabela GS_REDES a qual contem de forma desnormalizada todos os atributos de um trecho de rede de água
'
'
Public Sub CriaTabelaRedes()
    On Error GoTo Trata_Erro
    Dim str As String                                'string SQL para criar a tabela GS_REDES
   
    'prepara a string de conexão para criar a tabela de atributos de redes
    str = "CREATE TABLE GS_REDES(object_id_ nvarchar(255) NOT NULL, idtipo int NULL, tiporede varchar(25) NULL, "
    str = str + "especifi varchar(100) NULL, altinic float NULL, altfinal float NULL, profinic float NULL, "
    str = str + "proffin float NULL, diamint varchar(25) NULL, diamext varchar(25) NULL, noinic int NULL, "
    str = str + "nofinal int NULL, espessur float NULL, material int NULL, matnome nvarchar(50) NOT NULL,comprim float NULL, "
    str = str + "compcalc float NULL, forneced int NULL, forempre nvarchar(40) NOT NULL, localiza int NULL, "
    str = str + "locanome nvarchar(25) NOT NULL, estado int NULL, estnome nvarchar(25) NOT NULL, validinf int NULL, "
    str = str + "rugosid float NOT NULL, setor float NOT NULL, datainst varchar(25) NULL, ladorua smallint NULL, "
    str = str + "ladodesc varchar(50) NOT NULL, distdiv numeric(18, 2) NULL, problema smallint NULL, usuario varchar(50) NULL, "
    str = str + "datausu varchar(50) NULL, data varchar (25) NOT NULL, fabrican numeric(18, 0) NULL, "
    str = str + "fabrnome nvarchar(40) NOT NULL, LINE_ID int NULL, txtacm varchar (50), txtab varchar (50), angulol float NULL, "   'nome fabricante, texto acima, texto abaixo, angulo da linha
    str = str + "txtabl1 varchar (50), txtabl2 varchar (50) )"                                                                      'texto abaixo linha 1, texto abaixo linha 2
    
    'conforme o tipo de banco de dados
    Select Case frmCanvas.TipoConexao
        Case 1, 2, 3
            'SQLServer ou Oracle
            On Error Resume Next                    'abre o controle de erro caso a tabela não exista no banco de dados
            Conn.execute "DROP TABLE GS_REDES"      'elimina a tabela caso a mesma exista
            On Error GoTo Trata_Erro                'muda o controle de erro para avisar caso não consiga criar a tabela
            Conn.execute str                        'cria a tabela GS_REDES
        Case 4
            'Postgres - necessita testar esta funcionalidade, ainda não foi validada
            Dim strConn As String                           'string de conexão com o banco de dados Postgres
            Dim conexao As New ADODB.connection             'conexão com o banco de dados Postgres
            Dim mPROVEDOR As String
            Dim mSERVIDOR As String
            Dim mPORTA As String
            Dim mBANCO As String
            Dim mUSUARIO As String
            Dim Senha As String
            Dim decriptada As String
            Dim nStr As String
            
            mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
            mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
            mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
            mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
            Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
            nStr = frmCanvas.FunDecripta(Senha)
            decriptada = frmCanvas.Senha
            strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"
            conexao.Open strConn                    'abre a conexão com o banco de dados Postgres
            On Error Resume Next                    'abre o controle de erro caso a tabela não exista no banco de dados
            conexao.execute "DROP TABLE GS_REDES"   'elimina a tabela caso a mesma exista
            On Error GoTo Trata_Erro                'muda o controle de erro para avisar caso não consiga criar a tabela
            conexao.execute str                     'cria a tabela GS_REDES
        Case Else
            MsgBox "Classe GeoSanExport - método CriaTabelaRedes: Conexão com o banco de dados não localizada."
    End Select
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "CriaTabelaRedes-Não consegue criar a tabela no banco de dados com o comando: " & str, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Cria a tabela GS_CONSUMIDORES a qual contem os dados das ligações de água
'
'
Public Sub CriaTabelaConsumidores()
    On Error GoTo Trata_Erro
    Dim str As String                                'string SQL para criar a tabela GS_REDES
    'prepara a string de conexão para criar a tabela de ramais de ligação de água
    'Object_id_
    'NRO_LIGACAO - NRO_LIG
    'NRO_LIGACAO_SEM_DV - NROL_SDV
    'INSCRICAO_LOTE - INSCLOTE
    'Auto
    'CONSUMO_LPS - CONS_LPS
    'tipo
    'COD_LOGRADOURO - COD_LOG
    'ENDERECO
    'NUM_CASA
    'COMPL_LOGRADOURO - COMPLLOG
    'BAIRRO
    'HIDROMETRO - HIDROM
    'COD_CONSUMIDOR - COD_CONS
    'COD_CONSUMIDOR_SEM_DV - CODCONSDV
    'CONSUMIDOR - CONSUM
    'TEL_RES
    'TEL_COM
    'ECONOMIAS - ECONOM
    'HIDROMETRADO - HIDROM
    'DATA_IMPLANTACAO_AGUA - DATAIMP
    'DATA_IMPLANTACAO_ESGOTO - DATAIMPE
    'ROTA_LEITURA - ROTALEIT
    'GRUPO_LEITURA - GRUPOL
    'ROTEIRO_LEITURA - ROTLEIT
    'SEQUENCIA_LEITURA - SEQLEI
    'CONSUMO_FATURADO - CONSFAT
    'CONSUMO_MEDIDO - CONMED
    'Mes
    'ANO
    
    str = "CREATE TABLE GS_CONSUMIDORES (OBJECT_ID_ nvarchar(250) NULL, NRO_LIG nvarchar(50) NULL,NROL_SDV nvarchar(50), INSCLOTE nvarchar(50) NULL,"
    str = str + " AUTO nvarchar(10), CONS_LPS nvarchar(24) NULL, TIPO nvarchar(20) NULL,"
    str = str + " COD_LOG nvarchar(30) NULL, ENDERECO nvarchar(250) NULL, NUM_CASA nvarchar (15) NULL,"
    str = str + " COMPLLOG nvarchar(60) NULL, BAIRRO nvarchar(150) NULL, HIDROM nvarchar(30) NULL,"
    str = str + " COD_CONS nvarchar(11) NULL, CODCONSDV nvarchar(11) NULL, CONSUM nvarchar(150) NULL,"
    str = str + " TEL_RES nvarchar(20) NULL, TEL_COM nvarchar(20) NULL, ECONOM nvarchar(10) NULL,"
    str = str + " DATAIMP nvarchar(30), DATAIMPE nvarchar(30),"
    str = str + " ROTALEIT nvarchar(50) NULL, GRUPOL nvarchar(10) NULL, ROTLEIT nvarchar(10) NULL,"
    str = str + " SEQLEI nvarchar(10) NULL,"
    str = str + " CONSFAT nvarchar(15) NULL, CONMED nvarchar(15) NULL, MES nvarchar(2), ANO nvarchar(2))"
    'conforme o tipo de banco de dados
    Select Case frmCanvas.TipoConexao
        Case 1, 2, 3
            'SQLServer ou Oracle
            On Error Resume Next                    'abre o controle de erro caso a tabela não exista no banco de dados
            Conn.execute "DROP TABLE GS_CONSUMIDORES"      'elimina a tabela caso a mesma exista
            On Error GoTo Trata_Erro                'muda o controle de erro para avisar caso não consiga criar a tabela
            Conn.execute str                        'cria a tabela GS_REDES
        Case 4
            'Postgres - necessita testar esta funcionalidade, ainda não foi validada
            Dim strConn As String                           'string de conexão com o banco de dados Postgres
            Dim conexao As New ADODB.connection             'conexão com o banco de dados Postgres
            Dim mPROVEDOR As String
            Dim mSERVIDOR As String
            Dim mPORTA As String
            Dim mBANCO As String
            Dim mUSUARIO As String
            Dim Senha As String
            Dim decriptada As String
            Dim nStr As String
            
            mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
            mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
            mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
            mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
            Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
            nStr = frmCanvas.FunDecripta(Senha)
            decriptada = frmCanvas.Senha
            strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"
            conexao.Open strConn                    'abre a conexão com o banco de dados Postgres
            On Error Resume Next                    'abre o controle de erro caso a tabela não exista no banco de dados
            conexao.execute "DROP TABLE GS_CONSUMIDORES"   'elimina a tabela caso a mesma exista
            On Error GoTo Trata_Erro                'muda o controle de erro para avisar caso não consiga criar a tabela
                Open "d:\temp.txt" For Output As #1
                Print #1, str
                Close #1
            conexao.execute str                     'cria a tabela GS_REDES
        Case Else
            MsgBox "Classe GS_CONSUMIDORES - método CriaTabelaConsumidores: Conexão com o banco de dados não localizada."
    End Select
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "CriaTabelaConsumidores-Não consegue criar a tabela no banco de dados com o comando: " & str, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Cria a tabela GS_RAMAIS a qual contem as linhas dos ramais de ligação de água
'
'
Public Sub CriaTabelaRamais()
    On Error GoTo Trata_Erro
    Dim str As String                                'string SQL para criar a tabela GS_REDES
   
    'prepara a string de conexão para criar a tabela de ramais de ligação de água
    str = "CREATE TABLE GS_RAMAIS (object_id_ nvarchar(255) NULL, objidtre nvarchar(255) NULL, codlog int NULL,"
    str = str + " distTest numeric(18, 2) NULL, distLado numeric(18, 2) NULL, compram numeric(18, 2) NULL,"
    str = str + " profram numeric(18, 2) NULL, poslote int NULL, id int not NULL, datalog varchar (30) NULL, usulog varchar (30) NULL)"
   
    'conforme o tipo de banco de dados
    Select Case frmCanvas.TipoConexao
        Case 1, 2, 3
            'SQLServer ou Oracle
            On Error Resume Next                    'abre o controle de erro caso a tabela não exista no banco de dados
            Conn.execute "DROP TABLE GS_RAMAIS"      'elimina a tabela caso a mesma exista
            On Error GoTo Trata_Erro                'muda o controle de erro para avisar caso não consiga criar a tabela
            Conn.execute str                        'cria a tabela GS_REDES
        Case 4
            'Postgres - necessita testar esta funcionalidade, ainda não foi validada
            Dim strConn As String                           'string de conexão com o banco de dados Postgres
            Dim conexao As New ADODB.connection             'conexão com o banco de dados Postgres
            Dim mPROVEDOR As String
            Dim mSERVIDOR As String
            Dim mPORTA As String
            Dim mBANCO As String
            Dim mUSUARIO As String
            Dim Senha As String
            Dim decriptada As String
            Dim nStr As String
            
            mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
            mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
            mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
            mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
            Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
            nStr = frmCanvas.FunDecripta(Senha)
            decriptada = frmCanvas.Senha
            strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"
            conexao.Open strConn                    'abre a conexão com o banco de dados Postgres
            On Error Resume Next                    'abre o controle de erro caso a tabela não exista no banco de dados
            conexao.execute "DROP TABLE GS_REDES"   'elimina a tabela caso a mesma exista
            On Error GoTo Trata_Erro                'muda o controle de erro para avisar caso não consiga criar a tabela
            conexao.execute str                     'cria a tabela GS_REDES
        Case Else
            MsgBox "Classe GeoSanExport - método CriaTabelaRedes: Conexão com o banco de dados não localizada."
    End Select
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "CriaTabelaRamais-Não consegue criar a tabela no banco de dados com o comando: " & str, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Cria a tabela GS_NOS a qual contem de forma desnormalizada todos os atributos de um nó de rede de água
'
'
Public Sub CriaTabelaNos()
    On Error GoTo Trata_Erro
    Dim str As String                               'string SQL para criar a tabela GS_NOS
   
    'prepara a string de conexão para criar a tabela de atributos de nós
    str = "CREATE TABLE GS_NOS(object_id_ nvarchar(50) NOT NULL, idtipo int NULL, idsubtipo int NULL, "
    str = str + "valor varchar(50) NULL, estado int NULL, localiza int NULL, cota float NULL, "
    str = str + "notacao nvarchar(50) NULL, demanda float NULL, datainst varchar (50) NULL, descric varchar (25) NULL, "
    str = str + "especif varchar (50) NULL, x float NULL, y float NULL)"
    
    'conforme o tipo de banco de dados
    Select Case frmCanvas.TipoConexao
        Case 1, 2, 3
            'SQLServer ou Oracle
            On Error Resume Next                    'abre o controle de erro caso a tabela não exista no banco de dados
            Conn.execute "DROP TABLE GS_NOS"        'elimina a tabela caso a mesma exista
            On Error GoTo Trata_Erro                'muda o controle de erro para avisar caso não consiga criar a tabela
            Conn.execute str                        'cria a tabela GS_NOS
        Case 4
            'Postgres - necessita testar esta funcionalidade, ainda não foi validada
            Dim strConn As String                           'string de conexão com o banco de dados Postgres
            Dim conexao As New ADODB.connection             'conexão com o banco de dados Postgres
            Dim mPROVEDOR As String
            Dim mSERVIDOR As String
            Dim mPORTA As String
            Dim mBANCO As String
            Dim mUSUARIO As String
            Dim Senha As String
            Dim decriptada As String
            Dim nStr As String
            
            mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
            mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
            mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
            mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
            Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
            nStr = frmCanvas.FunDecripta(Senha)
            decriptada = frmCanvas.Senha
            strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"
            conexao.Open strConn                    'abre a conexão com o banco de dados Postgres
            On Error Resume Next                    'abre o controle de erro caso a tabela não exista no banco de dados
            conexao.execute "DROP TABLE GS_REDES"   'elimina a tabela caso a mesma exista
            On Error GoTo Trata_Erro                'muda o controle de erro para avisar caso não consiga criar a tabela
            conexao.execute str                     'cria a tabela GS_NOS
        Case Else
            MsgBox "Classe GeoSanExport - método CriaTabelaNos: Conexão com o banco de dados não localizada."
    End Select
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "CriaTabelaNos-Não consegue criar a tabela no banco de dados com o comando: " & str, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Verifica se existe conexão com o banco de dados comercial
'
'
'
Private Function ExisteConexaoBancoComercial() As Boolean
    Dim str As String
    Dim rs As New ADODB.Recordset
    On Error GoTo Trata_Erro:

    str = "SELECT * FROM NXGS_V_LIG_COMERCIAL"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic      'abre a conexão
    If rs.EOF = True Then
        ExisteConexaoBancoComercial = False
    Else
        ExisteConexaoBancoComercial = True
    End If
    rs.Close
    Exit Function
    
Trata_Erro:
    ExisteConexaoBancoComercial = False                     'deu erro, então não existe a conexão com o banco comercial
End Function




'Insere todos dos dados de todas as redes na tabela de atributos de redes GS_REDES
'
'retorna verdadeiro se conseguiu inserir os consumidores
'
Public Function InsereConsumidores() As Boolean
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    Dim data As String
    Dim frm As New frmProgressBar
    Dim obj_id, NRO_LIG, NROL_SDV, INSCLOTE, CONS_LPS, TIPOC, COD_LOG, ENDERECO, NUM_CASA As String     'textos já preparados para inserir consumidores
    Dim COMPLLOG, BAIRRO, HIDROM, COD_CONS, CODCONSDV, CONSUM, TEL_RES As String
    Dim TEL_COM, ECONOM, DATAIMP, DATAIMPE, ROTALEIT, GRUPOL, ROTLEIT, SEQLEI As String
    Dim AUTOC, CONSUFAT, CONMED, MESC, Ano, CONSFAT As String                                           'Não utilizados no momento
    Dim logErro As String
    
    Dim testeValor As Long
    
    InsereConsumidores = True
    If ExisteConexaoBancoComercial = False Then                                                         'não localizou a tabela comercial de consumidores
        InsereConsumidores = False
        Exit Function
    End If
    frm.Show , FrmMain
    frm.ProgressBar1.Min = 0
    str = "SELECT COUNT(*) FROM RAMAIS_AGUA_LIGACAO a"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic      'abre a conexão
    frm.ProgressBar1.Max = rs.Fields(0) + 1000
    rs.Close
    
    'prepara a querie para ler todos os dados de um trecho de rede de água
    str = "select * from RAMAIS_AGUA_LIGACAO a "
    str = str + "inner join NXGS_V_LIG_COMERCIAL b on Convert(varchar, Convert(bigint,a.NRO_LIGACAO) / 10) = b.NRO_LIGACAO_SEM_DV "
    'str = str + "inner join NXGS_V_LIG_COM_CONSUMO_MEDIO on NXGS_V_LIG_COM_CONSUMO_MEDIO.NRO_LIGACAO_SEM_DV = RAMAIS_AGUA_LIGACAO.NRO_LIGACAO /10"
    logErro = "1 - " + str
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic      'abre a conexão
    Do While Not rs.EOF
        logErro = "6a - moveu para o próximo, object_id:" + rs.Fields("Object_id_").value + " SQL: " + str
        DoEvents                                                            'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
        If varGlobais.pararExecucao = True Then
            Screen.MousePointer = vbNormal
            rs.Close
            Unload frm
            InsereConsumidores = False
            Exit Function
        End If
        logErro = "6b - moveu para o próximo, object_id:" + rs.Fields("Object_id_").value + " SQL: " + str
        'prepara textos já concatenados para redes longas e curtas
        obj_id = rs.Fields("Object_id_").value
        FrmMain.sbStatusBar.Panels(2).Text = "Insere consumidor: " & obj_id                             'apresenta na barra de status o andamento da inserção
        NRO_LIG = CStr(IIf(IsNull(rs.Fields("NRO_LIGACAO").value), "", rs.Fields("NRO_LIGACAO").value))
        NROL_SDV = CStr(IIf(IsNull(rs.Fields("NRO_LIGACAO_SEM_DV").value), "", rs.Fields("NRO_LIGACAO_SEM_DV").value))
        INSCLOTE = IIf(IsNull(rs.Fields("INSCRICAO_LOTE").value), "", rs.Fields("INSCRICAO_LOTE").value)
        AUTOC = ""
        CONS_LPS = Replace(CStr(FormatNumber(IIf(IsNull(rs.Fields("CONSUMO_LPS").value), 0, rs.Fields("CONSUMO_LPS").value), 8)), ",", ".")
        TIPOC = IIf(IsNull(rs.Fields("TIPO").value), "", rs.Fields("TIPO").value)
        COD_LOG = IIf(IsNull(rs.Fields("COD_LOGRADOURO").value), "", rs.Fields("COD_LOGRADOURO").value)
        logErro = "6c - moveu para o próximo, object_id:" + rs.Fields("Object_id_").value + " SQL: " + str
        ENDERECO = Replace(IIf(IsNull(rs.Fields("ENDERECO").value), "", rs.Fields("ENDERECO").value), Chr(39), " ")
        NUM_CASA = IIf(IsNull(rs.Fields("NUM_CASA").value), "", rs.Fields("NUM_CASA").value)
        COMPLLOG = Replace(IIf(IsNull(rs.Fields("COMPL_LOGRADOURO").value), "", rs.Fields("COMPL_LOGRADOURO").value), Chr(39), " ")
        BAIRRO = Replace(IIf(IsNull(rs.Fields("BAIRRO").value), "", rs.Fields("BAIRRO").value), Chr(39), " ")
        HIDROM = IIf(IsNull(rs.Fields("HIDROMETRO").value), "", rs.Fields("HIDROMETRO").value)
        COD_CONS = CStr(IIf(IsNull(rs.Fields("COD_CONSUMIDOR").value), "", rs.Fields("COD_CONSUMIDOR").value))
        CODCONSDV = CStr(IIf(IsNull(rs.Fields("COD_CONSUMIDOR_SEM_DV").value), "", rs.Fields("COD_CONSUMIDRO_SEM_DV").value))
        logErro = "6d - moveu para o próximo, object_id:" + rs.Fields("Object_id_").value + " SQL: " + str
        CONSUM = Replace(IIf(IsNull(rs.Fields("CONSUMIDOR").value), "", rs.Fields("CONSUMIDOR").value), Chr(39), " ")
        TEL_RES = IIf(IsNull(rs.Fields("TEL_RES").value), "", rs.Fields("TEL_RES").value)
        TEL_COM = IIf(IsNull(rs.Fields("TEL_COM").value), "", rs.Fields("TEL_COM").value)
        ECONOM = CStr(FormatNumber(IIf(IsNull(rs.Fields("ECONOMIAS").value), 0, rs.Fields("ECONOMIAS").value), 0))
        DATAIMP = CStr(IIf(IsNull(rs.Fields("DATA_IMPLANTACAO_AGUA").value), "", rs.Fields("DATA_IMPLANTACAO_AGUA").value))
        DATAIMPE = CStr(IIf(IsNull(rs.Fields("DATA_IMPLANTACAO_ESGOTO").value), "", rs.Fields("DATA_IMPLANTACAO_ESGOTO").value))
        ROTALEIT = CStr(IIf(IsNull(rs.Fields("ROTA_LEITURA").value), "", rs.Fields("ROTA_LEITURA").value))
        logErro = "6e - moveu para o próximo, object_id:" + rs.Fields("Object_id_").value + " SQL: " + str
        GRUPOL = CStr(IIf(IsNull(rs.Fields("GRUPO_LEITURA").value), "", rs.Fields("GRUPO_LEITURA").value))
        ROTLEIT = CStr(IIf(IsNull(rs.Fields("ROTEIRO_LEITURA").value), "", rs.Fields("ROTEIRO_LEITURA").value))
        SEQLEI = CStr(IIf(IsNull(rs.Fields("SEQUENCIA_LEITURA").value), "", rs.Fields("SEQUENCIA_LEITURA").value))
        logErro = "6f - moveu para o próximo, object_id:" + rs.Fields("Object_id_").value + " SQL: " + str
        CONSFAT = ""
        CONMED = ""
        MESC = ""
        Ano = ""
        logErro = "2 - configuração de variáveis"
        'prepara insert com todos os dados da querie acima
        str = "INSERT INTO GS_CONSUMIDORES (OBJECT_ID_, NRO_LIG,NROL_SDV, INSCLOTE,"
        str = str + " AUTO, CONS_LPS, TIPO,"
        str = str + " COD_LOG, ENDERECO, NUM_CASA,"
        str = str + " COMPLLOG, BAIRRO, HIDROM,"
        str = str + " COD_CONS, CODCONSDV, CONSUM,"
        str = str + " TEL_RES, TEL_COM, ECONOM,"
        str = str + " DATAIMP, DATAIMPE,"
        str = str + " ROTALEIT, GRUPOL, ROTLEIT,"
        str = str + " SEQLEI,  CONSFAT, CONMED, MES, ANO) "
        str = str + "Values "
        str = str + "( '" + obj_id + "' , '" + NRO_LIG + "' , '" + NROL_SDV
        str = str + "' , '" + INSCLOTE + "' , '" + AUTOC + "' , '" + CONS_LPS + "' , '" + TIPOC
        str = str + "' , '" + COD_LOG + "' , '" + ENDERECO + "' , '" + NUM_CASA
        str = str + "' , '" + COMPLLOG + "' , '" + BAIRRO + "' , '" + HIDROM
        str = str + "' , '" + COD_CONS + "' , '" + CODCONSDV + "' , '" + CONSUM
        str = str + "' , '" + TEL_RES + "' , '" + TEL_COM + "' , '" + ECONOM
        str = str + "' , '" + DATAIMP + "' , '" + DATAIMPE + "' , '" + ROTALEIT
        str = str + "' , '" + GRUPOL + "' , '" + ROTLEIT + "' , '" + SEQLEI
        str = str + "' , '" + CONSFAT + "' , '" + CONMED + "' , '" + MESC + "' , '" + Ano + "' )"
        logErro = "3 - " + str
        Conn.execute (str)              'insere a linha
        logErro = "4 - " + str
        rs.MoveNext                     'próxima
        logErro = "5 - " + str
        
        testeValor = frm.ProgressBar1
'        If testeValor >= 101702 Then
'            MsgBox ("Chegou:" + CStr(testeValor))
'        End If
        frm.ProgressBar1.value = frm.ProgressBar1.value + 1
        logErro = "8 - " + str
    Loop
    logErro = "7 - Finalizou - " + str
    rs.Close
    Unload frm
    InsereConsumidores = True
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        InsereConsumidores = False
        ErroUsuario.Registra "GeosanExport", "InsereConsumidores - " & logErro & " - ", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function
'Insere todos dos dados de todas as redes na tabela de atributos de redes GS_REDES
'
'
Public Sub InsereRedes()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    Dim data As String
    Dim frm As New frmProgressBar
    Dim textoAcima, textoAbaixo, textoAbaixoCurtoLinha1, textoAbaixoCurtoLinha2             'textos já preparados para colocar na rede, o curto é para quando o segumento de rede for muito pequeno
    Dim diametroExterno, diametroInterno As String                                          'para facilitar a criação da querie Insert
    Dim profundidadeInicial, profundidadeFinal As String                                    'profundidades com ponto decimal
    
    frm.Show , FrmMain
    frm.ProgressBar1.Min = 0
    str = "SELECT COUNT(*) FROM WATERLINES"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic      'abre a conexão
    frm.ProgressBar1.Max = rs.Fields(0)
    rs.Close
    
    'prepara a querie para ler todos os dados de um trecho de rede de água
    str = "SELECT * FROM WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID "
    str = str + "INNER JOIN X_LOCATION ON WATERLINES.LOCATION = X_LOCATION.LOCATIONID "
    str = str + "INNER JOIN X_SIDESTREET ON WATERLINES.SIDESTREET = X_SIDESTREET.SIDESTREET_ID "
    str = str + "INNER JOIN X_SUPPLIERS ON WATERLINES.SUPPLIER = X_SUPPLIERS.SUPPLIERID "
    str = str + "INNER JOIN X_STATE ON WATERLINES.STATE = X_STATE.STATEID "
    str = str + "INNER JOIN WATERLINESTYPES ON WATERLINES.ID_TYPE = WATERLINESTYPES.ID_TYPE "
    str = str + "INNER JOIN X_MANUFACTURERS ON WATERLINES.MANUFACTURER = X_MANUFACTURERS.MANUFACTURERID "
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic      'abre a conexão
    Do While Not rs.EOF
        DoEvents                                                            'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
        If varGlobais.pararExecucao = True Then
            Screen.MousePointer = vbNormal
            rs.Close
            Unload frm
            Exit Sub
        End If
        FrmMain.sbStatusBar.Panels(2).Text = "Insere rede: " & rs.Fields("Object_id_").value                             'apresenta na barra de status o andamento da inserção
        'prepara textos já concatenados para redes longas e curtas
        textoAcima = Replace(CStr(rs.Fields("LengthCalculated").value), ",", ".") + " m"
        textoAbaixo = rs.Fields("MaterialName").value + " - " + " Ø" + CStr(Format((rs.Fields("InternalDiameter").value), "#"))
        textoAbaixoCurtoLinha1 = rs.Fields("MaterialName").value
        textoAbaixoCurtoLinha2 = "Ø" + CStr(Format((rs.Fields("InternalDiameter").value), "#"))
        diametroExterno = CStr(FormatNumber((IIf(IsNull(rs.Fields("ExternalDiameter").value), 0, rs.Fields("ExternalDiameter").value)), 0))
        diametroInterno = "Ø" + CStr(FormatNumber((rs.Fields("InternalDiameter").value), 0))
        profundidadeInicial = Replace(CStr(rs.Fields("InitialTubeDeepness").value), ",", ".")
        profundidadeFinal = Replace(CStr(rs.Fields("FinalTubeDeepness").value), ",", ".")
        
        'prepara a coluna Data de Instalação
        If IsNull(rs.Fields("DateInstallation").value) Then
            data = ""
        Else
            data = rs.Fields("DateInstallation").value
        End If
        'prepara insert com todos os dados da querie acima
        str = "INSERT INTO GS_REDES (object_id_,idtipo,tiporede,especifi,altinic,altfinal"
        str = str + ",profinic,proffin,diamint,diamext,noinic"
        str = str + ",nofinal,espessur,material,matnome,comprim,compcalc,forneced"
        str = str + ",forempre,localiza,locanome,estado,estnome,validinf,rugosid"
        str = str + ",setor,datainst,ladorua,ladodesc,distdiv,problema,usuario"
        str = str + ",datausu,data,fabrican,fabrnome,LINE_ID"
        str = str + ",txtacm, txtab, txtabl1, txtabl2, angulol)"
        str = str + "Values "
        str = str + "( '" + rs.Fields("Object_id_").value + "' ," + CStr(rs.Fields("id_Type").value) + ", '" + rs.Fields("Description_").value + "' "
        str = str + ", '" + rs.Fields("Specification_").value + "' ," + CStr(rs.Fields("InitialGroundHeight").value) + "," + CStr(rs.Fields("FinalGroundHeight").value)
        str = str + "," + profundidadeInicial + "," + profundidadeFinal + ", '" + diametroInterno
        str = str + "' , '" + diametroExterno + "' ," + CStr(rs.Fields("InitialComponent").value) + "," + CStr(rs.Fields("FinalComponent").value)
        str = str + "," + CStr(rs.Fields("Thickness").value) + "," + CStr(rs.Fields("Material").value) + ", '" + rs.Fields("MaterialName").value + "' "
        str = str + "," + Replace(CStr(rs.Fields("LENGTH").value), ",", ".") + "," + Replace(CStr(rs.Fields("LengthCalculated").value), ",", ".") + "," + CStr(rs.Fields("Supplier").value)
        str = str + ", '" + rs.Fields("CompanyName").value + "' ," + CStr(rs.Fields("Location").value) + ", '" + rs.Fields("LocationName").value + "' "
        str = str + "," + CStr(rs.Fields("State").value) + ", '" + rs.Fields("StateName").value + "' ," + CStr(rs.Fields("InformationValidity").value)
        str = str + "," + CStr(rs.Fields("RoughNess").value) + "," + CStr(rs.Fields("Sector").value) + ", '" + data + "' "
        str = str + "," + CStr(rs.Fields("SideStreet").value) + ", '" + rs.Fields("Description").value + "' ," + Replace(CStr(rs.Fields("DividedDistance").value), ",", ".")
        str = str + "," + IIf(IsNull(rs.Fields("Trouble").value), "''", rs.Fields("Trouble").value) + ", '" + rs.Fields("USUARIO_LOG").value + "' , '" + rs.Fields("DATA_LOG").value + "' "
        str = str + ", '" + CStr(rs.Fields("DATALOG").value) + "' ," + CStr(rs.Fields("MANUFACTURER").value) + ", '" + rs.Fields("CompanyName").value + "' "
        str = str + "," + CStr(rs.Fields("LINE_ID").value)  'adicionar aqui os textos das redes acima e ambaixo (montar) e o ângulo da linha
        str = str + ", '" + textoAcima + "' , '" + textoAbaixo + "' , '" + textoAbaixoCurtoLinha1 + "' , '" + textoAbaixoCurtoLinha2 + "' , " + "0.0" + ")"
        Conn.execute (str)              'insere a linha
        rs.MoveNext                     'próxima
        frm.ProgressBar1.value = frm.ProgressBar1.value + 1
    Loop
    rs.Close
    Unload frm
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "InsereRedes", CStr(Err.Number), CStr(Err.Description) & " - " & str & " - ", True, glo.enviaEmails
    End If
End Sub
'Insere todos dos dados de todas os nós na tabela de atributos de redes GS_NOS
'
'
Public Sub InsereNos()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    Dim data As String
    Dim frm As New frmProgressBar
    
    frm.Show , FrmMain
    frm.ProgressBar1.Min = 0
    str = "SELECT COUNT(*) FROM WATERCOMPONENTS left join watercomponentsTypes on watercomponentstypes.id_type = watercomponents.id_type "
    'left join watercomponentsdata on watercomponentsdata.object_id_ = watercomponents.object_id_
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic      'abre a conexão
    frm.ProgressBar1.Max = rs.Fields(0)
    rs.Close
    
    'prepara a querie para ler todos os dados dos NÓS dos trechos de redes de água
    str = "SELECT watercomponents.object_id_ ,watercomponents.id_type, watercomponents.state "
    str = str + ",watercomponents.location, watercomponents.groundheight, watercomponents.notes "
    str = str + ",watercomponents.demand, watercomponents.dateinstallation, watercomponents.component_id "
    str = str + ",watercomponentsTypes.description_, watercomponentsTypes.specification_, x, y"
    ' str = str + ", watercomponentsdata.Id_subtype ,watercomponentsdata.Value_ "
    str = str + " from watercomponents "
    str = str + "left join watercomponentsTypes on watercomponentstypes.id_type = watercomponents.id_type "
    'str = str + "left join watercomponentsdata on watercomponentsdata.object_id_ = watercomponents.object_id_ " estava acrescentando mais linhas do que o desejado
    str = str + "left join points2 on watercomponents.object_id_ = points2.object_id "
    str = str + "order by convert (numeric(50),watercomponents.object_id_)"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic      'abre a conexão
    Do While Not rs.EOF
        DoEvents                                                            'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
        If varGlobais.pararExecucao = True Then
            Screen.MousePointer = vbNormal
            rs.Close
            Unload frm
            Exit Sub
        End If
        FrmMain.sbStatusBar.Panels(2).Text = "Insere nó: " & rs.Fields("Object_id_").value                             'apresenta na barra de status o andamento da inserção
        'prepara a coluna Data de Instalação
        If IsNull(rs.Fields("dateinstallation").value) Then
            data = "sem data"
        Else
            data = rs.Fields("dateinstallation").value
        End If
        'prepara insert com todos os dados da querie acima
        str = "INSERT INTO GS_NOS "
        str = str + "(object_id_, "
        '" idtipo, idsubtipo, "
        str = str + "valor, estado, localiza, cota, "
        str = str + "notacao, demanda, datainst, descric, "
        str = str + "especif, x, y, idtipo, idsubtipo) "
        str = str + "Values "
        str = str + "( '" + rs.Fields("Object_id_").value + "' ," + CStr(IIf(IsNull(rs.Fields("id_Type").value), "''", rs.Fields("id_Type").value)) + ", "
        '+ CStr(IIf(IsNull(rs.Fields("Id_subtype").value), "''", rs.Fields("Id_subtype").value)) +  ", '" + IIf(IsNull(rs.Fields("Value_").value), "", rs.Fields("Value_").value) + "' ,"
        str = str + CStr(IIf(IsNull(rs.Fields("state").value), "''", rs.Fields("state").value)) + "," + CStr(IIf(IsNull(rs.Fields("location").value), "''", rs.Fields("location").value))
        str = str + "," + Replace(CStr(rs.Fields("groundheight").value), ",", ".")
        'str = str + ", 'XXXX"
        str = str + ", '" + IIf(IsNull(rs.Fields("notes").value), "''", rs.Fields("notes").value)
        str = str + "' ," + Replace(CStr(rs.Fields("demand").value), ",", ".")
        str = str + ", '" + data + "' , '" + IIf(IsNull(rs.Fields("description_").value), "''", IIf(IsNull(rs.Fields("description_").value), "''", rs.Fields("description_").value)) + "' , '" + IIf(IsNull(rs.Fields("specification_").value), "''", rs.Fields("specification_").value)
        str = str + "'," + Replace(CStr(rs.Fields("x").value), ",", ".") + " , " + Replace(CStr(rs.Fields("y").value), ",", ".") + " , 0, 0 )"
        Conn.execute (str)              'insere a linha
        rs.MoveNext                     'próxima
        frm.ProgressBar1.value = frm.ProgressBar1.value + 1
    Loop
    rs.Close
    Unload frm
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "InsereNos", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
Public Sub InsereRamais()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    Dim data As String
    Dim frm As New frmProgressBar
    Dim distanciaTestada, distanciaLado, comprimentoRamal, profundidadeRamal As String           'textos já preparados para colocar no ramal
    Dim posicionamentoLote, ID, codigoLogradouro As String                                       'para facilitar a criação da querie Insert
    
    frm.Show , FrmMain
    frm.ProgressBar1.Min = 0
    str = "SELECT COUNT(*) FROM RAMAIS_AGUA"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic      'abre a conexão
    frm.ProgressBar1.Max = rs.Fields(0)
    rs.Close
    'prepara a querie para ler todos os dados de um trecho de rede de água
    str = "SELECT * FROM RAMAIS_AGUA"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic      'abre a conexão
    Do While Not rs.EOF
        DoEvents                                                            'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
        If varGlobais.pararExecucao = True Then
            Screen.MousePointer = vbNormal
            rs.Close
            Unload frm
            Exit Sub
        End If
        'prepara textos já concatenados para redes longas e curtas
        distanciaTestada = Replace(CStr(IIf(IsNull(rs.Fields("DISTANCIA_TESTADA").value), "0", rs.Fields("DISTANCIA_TESTADA").value)), ",", ".")
        distanciaLado = Replace(CStr(IIf(IsNull(rs.Fields("DISTANCIA_LADO").value), "0", rs.Fields("DISTANCIA_LADO").value)), ",", ".")
        comprimentoRamal = Replace(CStr(IIf(IsNull(rs.Fields("COMPRIMENTO_RAMAL").value), "0", rs.Fields("COMPRIMENTO_RAMAL").value)), ",", ".")
        profundidadeRamal = Replace(CStr(IIf(IsNull(rs.Fields("PROFUNDIDADE_RAMAL").value), "0", rs.Fields("PROFUNDIDADE_RAMAL").value)), ",", ".")
        posicionamentoLote = CStr(IIf(IsNull(rs.Fields("POSICIONAMENTO_LOTE").value), "0", rs.Fields("POSICIONAMENTO_LOTE").value))
        ID = CStr(rs.Fields("ID").value)
        FrmMain.sbStatusBar.Panels(2).Text = "Insere ramal: " & ID                              'apresenta na barra de status o andamento da inserção
        codigoLogradouro = CStr(IIf(IsNull(rs.Fields("COD_LOGRAD").value), " ", rs.Fields("COD_LOGRAD").value))
        str = "INSERT INTO GS_RAMAIS (object_id_, objidtre, codlog,"
        str = str + " distTest, distLado, compram,"
        str = str + " profram, poslote, id, datalog, usulog)"
        str = str + " Values "
        str = str + "( '" + rs.Fields("OBJECT_ID_").value + "' , '" + rs.Fields("OBJECT_ID_TRECHO").value + "' , '" + codigoLogradouro + "' "
        str = str + ", '" + distanciaTestada + "' ," + distanciaLado + "," + comprimentoRamal
        str = str + "," + profundidadeRamal + "," + posicionamentoLote + "," + ID
        str = str + ", '" + rs.Fields("DATA_LOG").value + "' , '" + rs.Fields("USUARIO_LOG").value + "' )"
        Conn.execute (str)              'insere a linha
        rs.MoveNext                     'próxima
        frm.ProgressBar1.value = frm.ProgressBar1.value + 1
    Loop
    rs.Close
    Unload frm
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "InsereRamais", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Exporta o arquivo no formato shape de redes de água
'
'
Public Sub ExportaRedesAguaShp()
Dim conexao As New ADODB.connection



End Sub
'Insere a indicação de existência de uma nova tabela desnormalizada GS_CONSUMIDORES, de atributos das ligações de água
'
'
Public Sub InsereTabAtributoConsumidores()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    'prepara a querie para verificar se a tabela de atributos GS_CONSUMIDORES já não foi inserida anteriormente (layer_id = 7)
    str = "Select attr_table from te_layer_table where attr_table = 'GS_CONSUMIDORES'"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        'se não foi inserida, insere
        str = "INSERT INTO te_layer_table (layer_id, attr_table, unique_id, attr_link, attr_time_unit, attr_table_type)"
        str = str + "VALUES (7, 'GS_CONSUMIDORES', 'object_id_', 'object_id_', 1, 1)"
        Conn.execute (str)
        'Conn.Close
    End If
    'rs.Close
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "InsereTabAtributoConsumidores", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Insere a indicação de existência de uma nova tabela de atributos GS_RAMAIS (layer_id=7), de atributos de ramais de ligação de água
'
'
Public Sub InsereTabAtributoRamais()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    'prepara a querie para verificar se a tabela de atributos GS_RAMAIS já não foi inserida anteriormente
    str = "Select attr_table from te_layer_table where attr_table = 'GS_RAMAIS'"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        'se não foi inserida, insere
        str = "INSERT INTO te_layer_table (layer_id, attr_table, unique_id, attr_link, attr_time_unit, attr_table_type)"
        str = str + "VALUES (7, 'GS_RAMAIS', 'object_id_', 'object_id_', 1, 1)"
        Conn.execute (str)
        'Conn.Close
    End If
    'rs.Close
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "InsereTabAtributoRamais", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub

'Insere a indicação de existência de uma nova tabela desnormalizada GS_REDE, de atributos de redes
'
'
Public Sub InsereTabAtributoRedes()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    'prepara a querie para verificar se a tabela de atributos GS_REDES já não foi inserida anteriormente
    str = "Select attr_table from te_layer_table where attr_table = 'GS_REDES'"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        'se não foi inserida, insere
        str = "INSERT INTO te_layer_table (layer_id, attr_table, unique_id, attr_link, attr_time_unit, attr_table_type)"
        str = str + "VALUES (1, 'GS_REDES', 'object_id_', 'object_id_', 1, 1)"
        Conn.execute (str)
        'Conn.Close
    End If
    'rs.Close
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "InsereTabAtributoRedes", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Insere a indicação de existência de uma nova tabela desnormalizada GS_NOS, de atributos de redes
'
'
Public Sub InsereTabAtributoNos()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    'prepara a querie para verificar se a tabela de atributos GS_NOS já não foi inserida anteriormente
    str = "Select attr_table from te_layer_table where attr_table = 'GS_NOS'"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        'se não foi inserida, insere
        str = "INSERT INTO te_layer_table (layer_id, attr_table, unique_id, attr_link, attr_time_unit, attr_table_type)"
        str = str + "VALUES (2, 'GS_NOS', 'object_id_', 'object_id_', 1, 1)"
        Conn.execute (str)
        'Conn.Close
    End If
    'rs.Close
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "InsereTabAtributoNos", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Na versão 3.3.1 do Terraview não é possível exportar um layer que possua mais de uma representação (texto + ponto por exemplo)
'Nesta rotina ele reativa a representação de pontos e linhas de consumidores de ligação de água
'
'
'
Public Sub AtivaRamaisGeoSan()
    On Error GoTo Trata_Erro
    Dim str As String
    Dim conexao As New ADODB.connection
    Dim localErro As String
    
    conexao.Open Conn
    'apaga todas as representações de consumidores
    str = "DELETE FROM te_representation where layer_id = 7"
    localErro = str
    Conn.execute (str)
    'insere a representação de pontos para poder exportar o shape. Assim a Te_export sabe o que exportar
    str = "INSERT INTO te_representation (layer_id, geom_type, geom_table) "
    str = str + "values (7, 4, 'Points7') "
    localErro = str
    Conn.execute (str)
    'insere a representação de linhas para poder exportar o shape. Assim a Te_export sabe o que exportar
    str = "INSERT INTO te_representation (layer_id, geom_type, geom_table) "
    str = str + "values (7, 2, 'Lines7') "
    Conn.execute (str)
    conexao.Close
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Dim descricaoErro As String
        
        descricaoErro = CStr(Err.Description) + " - '" + localErro + "'"
        ErroUsuario.Registra "GeosanExport", "AtivaRamaisGeoSan", CStr(Err.Number), descricaoErro, True, glo.enviaEmails
    End If
End Sub
'Na versão 3.3.1 do Terraview não é possível exportar um layer que possua mais de uma representação (texto + ponto por exemplo)
'Nesta rotina ele ativa a representação de texto de consumidores de ligação de água
'
'
'
Public Sub AtivaExportacaoConsumidores()
    On Error GoTo Trata_Erro
    Dim str As String
    Dim localErro As String
    
    'apaga todas as representações de consumidores
    localErro = "DELETE FROM te_representation where layer_id = 7"
    str = "DELETE FROM te_representation where layer_id = 7"
    Conn.execute (str)
    'insere a representação de pontos para poder exportar o shape. Assim a Te_export sabe o que exportar
    str = "INSERT INTO te_representation (layer_id, geom_type, geom_table) "
    str = str + "values (7, 4, 'Points7') "
    localErro = str
    Conn.execute (str)
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Dim descricaoErro As String
        
        descricaoErro = CStr(Err.Description) + " - '" + localErro + "'"
        ErroUsuario.Registra "GeosanExport", "AtivaExportacaoConsumidores", CStr(Err.Number), descricaoErro, True, glo.enviaEmails
    End If
End Sub
'Na versão 3.3.1 do Terraview não é possível exportar um layer que possua mais de uma representação (texto + ponto por exemplo)
'Nesta rotina ele ativa a representação de texto de ramais de ligação de água
'
'
'
Public Sub AtivaExportacaoRamais()
    On Error GoTo Trata_Erro
    Dim str As String
    
    'apaga todas as representações de ramais
    str = "DELETE FROM te_representation where layer_id = 7"
    Conn.execute (str)
    'insere a representação de linhas para poder exportar o shape. Assim a Te_export sabe o que exportar
    str = "INSERT INTO te_representation (layer_id, geom_type, geom_table) "
    str = str + "values (7, 2, 'Lines7') "
    Conn.execute (str)
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "AtivaExportacaoRamais", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Na versão 3.3.1 do Terraview não é possível exportar um layer que possua mais de uma representação (texto + ponto por exemplo)
'Nesta rotina ele ativa a representação de texto do nó
'
'
'
Public Sub AtivaExportacaoNos()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    'verifica se existe atributo do tipo texto (128) para o layer de WATERCOMPONENTS (2)
    str = "Select layer_id from te_representation where repres_id = 4 and layer_id = 2 and geom_type = 128"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then                      'encontrou a linha em te_representation
        str = "UPDATE te_representation set layer_id = 3 where repres_id = 4 and layer_id = 2 and geom_type = 128"
        Conn.execute (str)
    End If
    rs.Close
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "AtivaExportacaoNos", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Na versão 3.3.1 do Terraview não é possível exportar um layer que possua mais de uma representação (texto + ponto por exemplo)
'Nesta rotina ele desativa a representação de texto do nó
'
'
'
Public Sub DesativaExportacaoNos()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    'verifica se existe atributo do tipo texto (128) para o layer de WATERCOMPONENTS (2)
    str = "Select layer_id from te_representation where repres_id = 4 and layer_id = 3 and geom_type = 128"
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then                      'encontrou a linha em te_representation
        str = "UPDATE te_representation set layer_id = 2 where repres_id = 4 and layer_id = 3 and geom_type = 128"
        Conn.execute (str)
    End If
    rs.Close
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeosanExport", "DesativaExportacaoNos", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
