Attribute VB_Name = "Module1"
Option Explicit
Private conn As New ADODB.Connection

Public Provider As cAppType
Private ConnectString As String
Private Const PLANO = "WATERLINES"

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

Sub Main()
   Dim nc As New NexusConnection.App
   If Not nc.appGetRegistry("Exporte_EPANET", conn, Provider) Then
      If Not nc.appNewRegistry("Exporte_EPANET", conn, Provider) Then End
   End If

   'Set nc = Nothing
   
   With FrmEPANET
      Set .conn = conn
      .Provider = Provider
      .PLANO = PLANO
      .init
   End With


End Sub

