VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCoordIniFinLinha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Esta classe tem a finalidade de a partir de uma conexão com o banco de dados e um object_id de uma linha
' retornar as coordenadas inicial e final da mesma
'
'
'
Private conexaoTdb As TeDatabase     'cria uma variável privada e interna a classe
Private xi As Double
Private yi As Double
Private xf As Double
Private yf As Double
Private lin As New CLine2D
' Recebe o objeto da conexão TeDatabase
'
Public Property Let conexao(ByVal vData As TeDatabase)
    Set conexaoTdb = vData
End Property
'Retorna coordenadas do primeiro e último vértices da linha
'
Property Get linha() As CLine2D
    Set linha = lin
End Property
' Obtem as coordenadas x e y dos nós inicial e final da linha do trecho de rede
'
Public Function GetStartEndPointsTrecho(objId As String) As Integer
    Dim retorno As Boolean
    
    frmCanvas.TeDatabase4.connection = Conn
    If frmCanvas.TeDatabase4.setCurrentLayer("WaterLines") Then
        retorno = frmCanvas.TeDatabase4.getPointOfLine(0, objId, 0, xi, yi)
        retorno = frmCanvas.TeDatabase4.getPointOfLine(0, objId, frmCanvas.TeDatabase4.getQuantityPointsLine(0, objId) - 1, xf, yf)
        lin.xi = xi
        lin.yi = yi
        lin.xf = xf
        lin.yf = yf
    End If
End Function
' Obtem as coordenadas x e y dos nós inicial e final da linha do ramal
'
Public Function GetStartEndPointsRamal(objId As String) As Integer
    Dim retorno As Boolean
    
    frmCanvas.TeDatabase4.connection = Conn
    If frmCanvas.TeDatabase4.setCurrentLayer("RAMAIS_AGUA") Then
        retorno = frmCanvas.TeDatabase4.getPointOfLine(0, objId, 0, xi, yi)
        retorno = frmCanvas.TeDatabase4.getPointOfLine(0, objId, frmCanvas.TeDatabase4.getQuantityPointsLine(0, objId) - 1, xf, yf)
        lin.xi = xi
        lin.yi = yi
        lin.xf = xf
        lin.yf = yf
    End If
End Function
