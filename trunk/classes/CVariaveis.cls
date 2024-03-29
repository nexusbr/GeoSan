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
' Classe para armazenamento de vari�veis globais de uma forma um pouco mais elegante
'
Private objIDsTrechos As Variant
Private objIDTrechoSelecionado As String
Private totalTrechosMovendo As Integer
Private retornaEmailErro As Boolean
Private paraExe As Boolean
Private objIdNoMovendo As String
Private moveVertice As Boolean                          'indica se est� ou n�o movendo um v�rtice de uma rede
Private podeRealizarCommit As Boolean                   'tem objetivo de informar se � para realizar um Commit ou Rollback, pois houve falha
Private calcularOZNo As Boolean                         'para saber se a cota Z do n� deve ser calculada no momento em que estou desenhando a rede
Private numVertRamal As Integer                         'n�mero do v�rtice do ramal que est� sendo cadastrado pelo usu�rio, 1 = primeiro, 2 = segundo, ....
' Retorna qual o n�mero do v�rtice do ramal que est� sendo entrado pelo usu�rio
'
Property Get numeroVerticeDoRamal() As Integer
    numeroVerticeDoRamal = numVertRamal
End Property
' Recebe qual o n�mero do v�rtice do ramal que est� sendo entrado pelo usu�rio
'
Property Let numeroVerticeDoRamal(ByVal numVRamal As Integer)
    numVertRamal = numVRamal
End Property
' Retorna se pode ou n�o realizar um Commit no banco de dados
'
Property Get deveCalcularZNo() As Boolean
    deveCalcularZNo = calcularOZNo
End Property
' Recebe se pode ou n�o realizar um Commit no banco de dados
'
Property Let deveCalcularZNo(ByVal calculaZNo As Boolean)
    calcularOZNo = calculaZNo
End Property
' Retorna se pode ou n�o realizar um Commit no banco de dados
'
Property Get realizaCommit() As Boolean
    realizaCommit = podeRealizarCommit
End Property
' Recebe se pode ou n�o realizar um Commit no banco de dados
'
Property Let realizaCommit(ByVal podeComitar As Boolean)
    podeRealizarCommit = podeComitar
End Property
' Retorna se o v�rtice est� sendo movido ou n�o
'
Property Get moverVertice() As Boolean
    moverVertice = moveVertice
End Property
' Recebe se o v�rtice est� sendo movido ou n�o
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
'Retorna valor do n�mero total de trechos que est�o sendo movidos
'
Property Get totalTrechos() As Variant
    totalTrechos = totalTrechosMovendo
End Property
'  Recebe valor do n�mero total de trechos que est�o sendo movidos
'
Property Let totalTrechos(ByVal trechos As Variant)
    totalTrechosMovendo = trechos
End Property
'Retorna valor se � para enviar emails de erros
'
Property Get enviarEmails() As Variant
    enviarEmails = retornaEmailErro
End Property
'  Recebe valor se � para enviar emails de erros
'
Property Let seEnviaEmailErro(ByVal enviaEmails As Variant)
    retornaEmailErro = enviaEmails
End Property
'Retorna valor se � para enviar emails de erros
'
Property Get pararExecucao() As Variant
    pararExecucao = paraExe
End Property
'  Recebe valor se � para enviar emails de erros
'
Property Let pararExecucao(ByVal para As Variant)
    paraExe = para
End Property
'Retorna valor do object_id do n� selecionado para mover
'
Property Get objIdNoSelecionado() As Variant
    objIdNoSelecionado = objIdNoMovendo
End Property
'  Recebe valor do object_id do n� selecionado para mover
'
Property Let objIdNoSelecionado(ByVal objId As Variant)
    objIdNoMovendo = objId
End Property
Private Function variaveis()

End Function
