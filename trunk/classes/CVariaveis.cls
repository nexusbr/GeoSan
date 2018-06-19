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
Private objIDTrechoSelecionado As String
Private totalTrechosMovendo As Integer
Private retornaEmailErro As Boolean
Private paraExe As Boolean
Private objIdNoMovendo As String
Private moveVertice As Boolean                          'indica se está ou não movendo um vértice de uma rede
Private podeRealizarCommit As Boolean                   'tem objetivo de informar se é para realizar um Commit ou Rollback, pois houve falha
Private calcularOZNo As Boolean                         'para saber se a cota Z do nó deve ser calculada no momento em que estou desenhando a rede
Private numVertRamal As Integer                         'número do vértice do ramal que está sendo cadastrado pelo usuário, 1 = primeiro, 2 = segundo, ....
' Retorna qual o número do vértice do ramal que está sendo entrado pelo usuário
'
Property Get numeroVerticeDoRamal() As Integer
    numeroVerticeDoRamal = numVertRamal
End Property
' Recebe qual o número do vértice do ramal que está sendo entrado pelo usuário
'
Property Let numeroVerticeDoRamal(ByVal numVRamal As Integer)
    numVertRamal = numVRamal
End Property
' Retorna se pode ou não realizar um Commit no banco de dados
'
Property Get deveCalcularZNo() As Boolean
    deveCalcularZNo = calcularOZNo
End Property
' Recebe se pode ou não realizar um Commit no banco de dados
'
Property Let deveCalcularZNo(ByVal calculaZNo As Boolean)
    calcularOZNo = calculaZNo
End Property
' Retorna se pode ou não realizar um Commit no banco de dados
'
Property Get realizaCommit() As Boolean
    realizaCommit = podeRealizarCommit
End Property
' Recebe se pode ou não realizar um Commit no banco de dados
'
Property Let realizaCommit(ByVal podeComitar As Boolean)
    podeRealizarCommit = podeComitar
End Property
' Retorna se o vértice está sendo movido ou não
'
Property Get moverVertice() As Boolean
    moverVertice = moveVertice
End Property
' Recebe se o vértice está sendo movido ou não
'
Property Let moverVertice(ByVal vertice As Boolean)
    moveVertice = vertice
End Property

' Retorna valor de do object_id do trecho
'
Property Get objIdTreSelecionado() As String
    objIdTreSelecionado = objIDTrechoSelecionado
End Property
' Recebe valor do object_id do trecho
'
Property Let objIdTreSelecionado(ByVal objIdTreSel As String)
    objIDTrechoSelecionado = objIdTreSel
End Property
' Retorna valor de do object_id do trecho
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
'Retorna valor se é para enviar emails de erros
'
Property Get enviarEmails() As Variant
    enviarEmails = retornaEmailErro
End Property
'  Recebe valor se é para enviar emails de erros
'
Property Let seEnviaEmailErro(ByVal enviaEmails As Variant)
    retornaEmailErro = enviaEmails
End Property
'Retorna valor se é para enviar emails de erros
'
Property Get pararExecucao() As Variant
    pararExecucao = paraExe
End Property
'  Recebe valor se é para enviar emails de erros
'
Property Let pararExecucao(ByVal para As Variant)
    paraExe = para
End Property
'Retorna valor do object_id do nó selecionado para mover
'
Property Get objIdNoSelecionado() As Variant
    objIdNoSelecionado = objIdNoMovendo
End Property
'  Recebe valor do object_id do nó selecionado para mover
'
Property Let objIdNoSelecionado(ByVal objId As Variant)
    objIdNoMovendo = objId
End Property
Private Function variaveis()

End Function
