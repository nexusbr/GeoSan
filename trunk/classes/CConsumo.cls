VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para calcular valores relativos a consumos
'
'
'
' Converte litros por segundo para metros 3 por mês
'
' Retorna o valor convertido para m3/mes
' valorLps - consumo em litros por segundo para ser convertido
'
Public Function lps2m3mes(valorLps As Double) As Double
    lps2m3mes = valorLps * 2592                     '86.4 m3/dia x 30 dias
End Function
' Converte  metros 3 por mês para litros por segundo
'
' Retorna o valor convertido para l/s
' valorM3mes - consumo em metros 3 por mês para ser convertido
'
Public Function m3mes2lps(valorM3mes As Double) As Double
    m3mes2lps = valorM3mes / 2592
End Function
