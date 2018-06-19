Attribute VB_Name = "Inicio"
Option Explicit
Public banco As New RedeBancoDados

Public conn As New ADODB.Connection
Dim logExportacao As New logExportacao

Public Provider As cAppType
Private ConnectString As String
Private Const PLANO = "WATERLINES"

Public rsTrechosExportados As New ADODB.Recordset           'Criado para armazenar os trechos que já foram exportados
Public rsNosExportados As New ADODB.Recordset              'Criado para armazenar os Nós que já foram exportados
Public rsNosTmp As New ADODB.Recordset 'Criado para armazenar todos os dados de todos nos - Copia do Watercomponenstes/Points

'Criados Para Armazenar os Componentes / Trechos
Public rsTrechos As New ADODB.Recordset
Public rsCoordinates As New ADODB.Recordset
Public rsPipes As New ADODB.Recordset
Public rsJunctions As New ADODB.Recordset
Public rsPumps As New ADODB.Recordset
Public rsValves As New ADODB.Recordset
Public rsReservoirs As New ADODB.Recordset
Public rsTanks As New ADODB.Recordset
Public rsVertices As New ADODB.Recordset 'Vertices da linha com exceção do inicial e final

Enum Pipes 'LINHAS EM GERAL (TUBULAÇÃO)
   Normal = 0
   Valvula_simples = 1
   Valvula_Bomba = 2
End Enum

Enum TipoValvulas '
   Val_Retencao = 4
   Val_Gaveta = 3
   Val_Esferica = 2
   val_Desconhecida = 0
End Enum

Enum TipoNos
   No_Valvulas = 1
   No_Valvulas_99 = 99
   No_Bombas = 20
   No_Reservatorios = 40
   No_Tanques = 19
End Enum
'Inicia a aplicação por aqui
'
'
'
Sub Main()
    Dim nc As New NexusConnection.App
    If Not nc.appGetRegistry("Exporte_EPANET", conn, Provider) Then
        If Not nc.appNewRegistry("Exporte_EPANET", conn, Provider) Then End
    End If
    Set banco.Conexao = conn
    banco.ObtemNomeUsuario                                          'obtem o nome do usuário que exportou as redes
    banco.IniciaLeituraTrechosRede                                  'preparar o cursor para ler os trechos de rede que foram selecionados pelo usuário.

    
    logExportacao.AbrePrimeiraVez
    logExportacao.GravaTexto "ExportEpanet;*************************************************************************************************"
    logExportacao.GravaTexto "ExportEpanet;Início do processamento da exportação para o Epanet: " & DateValue(Now) & " - " & TimeValue(Now)

    With FrmEPANET
        Set .conn_recebida = conn
        FrmEPANET.Provider = Provider
        FrmEPANET.PLANO = PLANO
        FrmEPANET.init
    End With
End Sub

