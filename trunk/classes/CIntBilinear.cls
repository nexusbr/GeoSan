VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIntBilinear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para o cálculo da interpolação bilinear para descobrir a cota z de um ponto que está dentro de quatro pontos conhecidos
'
'
'
' Função que retora a cota z de um ponto dentro do um polígono definido por 4 pontos
'
' Retorna a coordenada z do ponto de interpolação
' xInt - coordenada x do ponto que temos que descobrir a cota
' yInt - coordenada y do ponto que temos que descobrir a cota
' xInt1 - coordenada x do primeiro ponto superior a esquerda
' yInt1 - coordenada y do primeiro ponto superior a esquerda
' xInt2 - coordendad x do segundo ponto superior a direita
' yInt2 - coordenada y do segundo ponto superior a direita
' xInt3 - coordenada x do terceiro ponto inferior a esquerda
' yInt3 - coordenada y do terceiro ponto inferior a esquerda
' xInt4 - coordenada x do quarto ponto inferior a direita
' yInt4 - coordenada y do quarto ponto inferior a direita
' z1 - coordenada z do primeiro ponto
' z2 - coordenada z do segundo ponto
' z3 - coordenada z do terceiro ponto
' z4 - coordenada z do quarto ponto
'

Public Function CalcInterpolacao(xInt As Double, yInt As Double, xInt1 As Double, yInt1 As Double, xInt2 As Double, yInt2 As Double, xInt3 As Double, yInt3 As Double, xInt4 As Double, yInt4 As Double, z1 As Double, z2 As Double, z3 As Double, z4 As Double) As Double
    On Error GoTo Trata_Erro:
    Dim i1 As Double                                                                        'primeira interpolaçao em x
    Dim i2 As Double                                                                        'segunda interpolaçao em x
    Dim i3 As Double                                                                        'interpolaçao em y
    
    i1 = ((xInt2 - xInt) / (xInt2 - xInt1)) * z1 + ((xInt - xInt1) / (xInt2 - xInt1)) * z2
    i2 = ((xInt2 - xInt) / (xInt2 - xInt1)) * z3 + ((xInt - xInt1) / (xInt2 - xInt1)) * z4
    i3 = ((yInt - yInt3) / (yInt1 - yInt3)) * i1 + ((yInt1 - yInt) / (yInt1 - yInt3)) * i2
    CalcInterpolacao = i3
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CIntBilinear", "CalcInterpolacao - divisão por zero no MDT", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function

