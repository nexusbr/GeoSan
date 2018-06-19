VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rugosidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Esta classe tem por finalidade realizar as operações de rugosidades nas tubulações de água.
'
'

'Atualiza todas as rugosidades de todas as tubulações, conforme o tipo de material. Foi considerada uma tubulação de 20 anos de idade
'Sempre que alguma linha for alterada na tabela X_Material, estas rugosidades deverão ser revistas
'
' conn - representa a conexão com o banco de dados geográfico do GeoSan
'
Public Function AtualizaAsRugosidades(bancoGeoSan As ADODB.Connection) As Boolean
    
    'Liga a ampulheta no ponteiro do mouse
    MousePointer = vbHourglass
    
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 0")                                            'zera todas as rugozidades
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 111 WHERE MATERIAL = 0")                       'Desconhecido
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 1 AND ROUGHNESS = 0")     'Cerâmico
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 120 WHERE MATERIAL = 2 AND ROUGHNESS = 0")     'Concreto
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 110 WHERE MATERIAL = 3 AND ROUGHNESS = 0")     'FF
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 105 WHERE MATERIAL = 4 AND ROUGHNESS = 0")     'Fibrocimento
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 90  WHERE MATERIAL = 5 AND ROUGHNESS = 0")     'PEAD
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 6 AND ROUGHNESS = 0")     'PRSV
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 7 AND ROUGHNESS = 0")     'PVC branco
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 8 AND ROUGHNESS = 0")     'PVC corrugado
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 9 AND ROUGHNESS = 0")     'PVC DeFoFo
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 10 AND ROUGHNESS = 0")    'PVC Ocre
    bancoGeoSan.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 11 AND ROUGHNESS = 0")    'PVC PBA
    
    'Volta o mouse para o normal
    FrmEPANET.MousePointer = vbDefault

End Function

Public Function ZeraMateriais(bancoGeoSan As ADODB.Connection)
    'Zera todos os materiais de tubulações quando o mesmo não estiver cadastrado
    bancoGeoSan.Execute ("UPDATE WATERLINES SET MATERIAL = 0 WHERE MATERIAL IS NULL")
End Function
