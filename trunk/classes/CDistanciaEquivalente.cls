VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDistanciaEquivalente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para calcular a nova dist�ncia do ramal, por regra de 3, ap�s o trecho ser movido

' Calcula a nova dist�ncia do in�cio de um ramal de �gua ap�s mover a rede
'
' distanciaRamalDepoisMovido - retorna a nova dist�ncia do ramal em rela��o ao in�cio da linha
' compLinhaAntes - comprimento da linha depois de movida
' compLinhaDepois - comprimento inicial da linha antes de mover
' distanciaRamalAntes - dist�ncia do ramal para o in�cio da linha
'
Public Function distanciaRamalDepoisMovido(compLinhaAntes As Double, compLinhaDepois As Double, distanciaRamalAntes As Double) As Double
    On Error GoTo Trata_Erro:
    distanciaRamalDepoisMovido = compLinhaAntes * distanciaRamalAntes / compLinhaDepois
    Exit Function

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      ErroUsuario.Registra "CDistanciaEquivalente", "distanciaRamalDepoisMovido", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
   End If
End Function
