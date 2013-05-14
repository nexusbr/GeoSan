VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVariaveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para armazenamento de variáveis globais de uma forma um pouco mais elegante
'
Private objIDsTrechos As Variant
Private totalTrechosMovendo As Integer
'Retorna valor de do object_id do trecho
'
Property Get objIdTre() As Variant
    objIdTre = objIDsTrechos
End Property
'  Recebe valor do object_id do trecho
'
Property Let objIdTre(ByVal objId As Variant)
    objIDsTrechos = objId
End Property
'Retorna valor do número total de trechos que estão sendo movidos
'
Property Get totalTrechos() As Variant
    totalTrechos = totalTrechosMovendo
End Property
'  Recebe valor do número total de trechos que estão sendo movidos
'
Property Let totalTrechos(ByVal trechos As Variant)
    totalTrechosMovendo = trechos
End Property
Private Function variaveis()

End Function
