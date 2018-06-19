VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObtemZdoNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para obter a cota z de um nó da rede
Dim obtemZdoNo As New CIntBilinear                                      'para obter a cota z do nó em questão
' Localiza a cota z de um determinado ponto
'
' xNo - coordenada x do ponto em que será localizada a cota z
' yNo - coordenada y do ponto em que será localizada a cota z
' object_Id_No - obj_id somente para em caso de erros saber onde o mesmo ocorreu
'
Public Function ObtemZ(xNo As Double, yNo As Double, object_Id_No As String) As Double
    On Error GoTo Trata_Erro
    Dim retorno As Boolean                                          'retorna false caso tenha ocorrido algum erro
    Dim idObjPonto As String                                        'object_id do ponto o qual é obtido inicialmente e não sabemos qual é
    Dim idObjPonto1 As String                                       'object_id do ponto 1 que esta ao redor do pesquisado
    Dim idObjPonto2 As String                                       'object_id do ponto 2 que esta ao redor do pesquisado
    Dim idObjPonto3 As String                                       'object_id do ponto 3 que esta ao redor do pesquisado
    Dim idObjPonto4 As String                                       'object_id do ponto 4 que esta ao redor do pesquisado
    Dim xMdt As Double
    Dim yMdt As Double
    Dim xInt1 As Double                                             'coordenada x do primeiro ponto superior a esquerda
    Dim yInt1 As Double                                             'coordenada y do primeiro ponto superior a esquerda
    Dim xInt2 As Double                                             'coordendad x do segundo ponto superior a direita
    Dim yInt2 As Double                                             'coordenada y do segundo ponto superior a direita
    Dim xInt3 As Double                                             'coordenada x do terceiro ponto inferior a esquerda
    Dim yInt3 As Double                                             'coordenada y do terceiro ponto inferior a esquerda
    Dim xInt4 As Double                                             'coordenada x do quarto ponto inferior a direita
    Dim yInt4 As Double                                             'coordenada y do quarto ponto inferior a direita
    Dim z1 As Double                                                'coordenada z do primeiro ponto
    Dim z2 As Double                                                'coordenada z do segundo ponto
    Dim z3 As Double                                                'coordenada z do terceiro ponto
    Dim z4 As Double                                                'coordenada z do quarto ponto
    Dim zNo As Double                                               'cota z do nó que estamos querendo descobrir
    Dim quad As Integer                                             'quadrante em que se localiza o nó
    Dim idGeomPonto As Long
    Dim rs As New ADODB.Recordset
    Dim distanciaMinima                                             'metade da diagonal da grade. Se temos uma grade de 10 x 10 a diagonal é SQRT (10*10 + 10*10) = 14.1421. Metade é 7.0710
    Dim mensagemDeErro As String                                    'para informar o motivo de ter encontrado o erro ao atualizar a cota
    Dim nomeColunaObjIdMdt As String                                'nome da coluna no layer MDT, que possui os objcts_ids
    Dim teNexusDatabase As New CDatabaseNexus                       'classe que acessa as tabelas Terralib
    
    nomeColunaObjIdMdt = teNexusDatabase.GetNomeColunaObjectId("MDT")
    distanciaMinima = 7.071                                         'distância da diagonal (5x5metros)
    mensagemDeErro = ""
    retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo, yNo, distanciaMinima, idGeomPonto, idObjPonto)                  'obtem o ponto do mdt que está mais perto do nó da rede a uma distância de até 5 metros
    cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xMdt, yMdt                                                                'obtem as coordenadas x e y do elemento tipo ponto (4) do geometry_id fornecido
    If xNo >= xMdt And yNo <= yMdt Then     'está no Quadrante 4
        quad = 4
        xInt1 = xMdt
        yInt1 = yMdt
        idObjPonto1 = idObjPonto
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo + 10#, yNo, distanciaMinima, idGeomPonto, idObjPonto2)
        If retorno = False Then
            mensagemDeErro = "O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (2A). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt2, yInt2
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo, yNo - 10#, distanciaMinima, idGeomPonto, idObjPonto3)
        If retorno = False Then
            mensagemDeErro = mensagemDeErro + " O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (3A). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt3, yInt3
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo + 10#, yNo - 10#, distanciaMinima, idGeomPonto, idObjPonto4)
        If retorno = False Then
            mensagemDeErro = mensagemDeErro + " O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (4A). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt4, yInt4
    ElseIf xNo > xMdt And yNo > yMdt Then   'está no Quadrante 2
        quad = 2
        xInt3 = xMdt
        yInt3 = yMdt
        idObjPonto3 = idObjPonto
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo + 10#, yNo, distanciaMinima, idGeomPonto, idObjPonto4)
        If retorno = False Then
            mensagemDeErro = "O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (2B). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt4, yInt4
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo, yNo + 10#, distanciaMinima, idGeomPonto, idObjPonto1)
        If retorno = False Then
            mensagemDeErro = mensagemDeErro + " O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (3B). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt1, yInt1
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo + 10#, yNo + 10#, distanciaMinima, idGeomPonto, idObjPonto2)
        If retorno = False Then
            mensagemDeErro = mensagemDeErro + " O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (4B). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt2, yInt2
    ElseIf xNo < xMdt And yNo < yMdt Then   'está no Quadrante 3
        quad = 3
        xInt2 = xMdt
        yInt2 = yMdt
        idObjPonto2 = idObjPonto
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo - 10#, yNo, distanciaMinima, idGeomPonto, idObjPonto1)
        If retorno = False Then
            mensagemDeErro = "O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (2C). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt1, yInt1
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo, yNo - 10#, distanciaMinima, idGeomPonto, idObjPonto4)
        If retorno = False Then
            mensagemDeErro = mensagemDeErro + " O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (3C). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt4, yInt4
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo - 10#, yNo - 10#, distanciaMinima, idGeomPonto, idObjPonto3)
        If retorno = False Then
            mensagemDeErro = mensagemDeErro + " O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (4C). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt3, yInt3
    ElseIf xNo < xMdt And yNo > yMdt Then   'está no Quadrante 1
        quad = 1
        xInt4 = xMdt
        yInt4 = yMdt
        idObjPonto4 = idObjPonto
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo - 10#, yNo, distanciaMinima, idGeomPonto, idObjPonto3)
        If retorno = False Then
            mensagemDeErro = "O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (2D). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt3, yInt3
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo, yNo + 10#, distanciaMinima, idGeomPonto, idObjPonto2)
        If retorno = False Then
            mensagemDeErro = mensagemDeErro + " O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (3D). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt2, yInt2
        retorno = cGeoDatabase.geoDatabase.locateNearestGeometry(tpPOINTS, xNo - 10#, yNo + 10#, distanciaMinima, idGeomPonto, idObjPonto1)
        If retorno = False Then
            mensagemDeErro = mensagemDeErro + " O nó da malha do MDT não está sendo localizado, não consegue atríbuir a cota ao nó (4D). "
        End If
        cGeoDatabase.geoDatabase.getCenterGeometry idGeomPonto, 0, 4, xInt1, yInt1
    End If
    
    'obtem agora as cotas dos 4 pontos do MDT
    mensagemDeErro = mensagemDeErro + "Nó com obj_id = " & object_Id_No & " coord. x = " & CStr(Round(xNo, 2)) & " coord. y = " & CStr(Round(yNo, 2)) & "MDT com obj_ids: ObjId1 = " & idObjPonto1 & ", ObjId2 = " & idObjPonto2 & ", ObjId3 = " & idObjPonto3 & ", ObjId4 = " & idObjPonto4
    If idObjPonto1 = Null Or idObjPonto2 = Null Or idObjPonto3 = Null Or idObjPonto4 = Null Then                                        'verifica se ele encontrou os nós do MDT, pois se não encontra não adianta realizar o cálculo
        ErroUsuario.Registra "CObtemZdoNo", "ObtemZ", CStr(Err.Number), CStr(Err.Description), False, glo.enviaEmails, mensagemDeErro   'não exibe mensagem para o usuário
        ObtemZ = 0                                                                                                                      'retorna sem calcular a cota Z
    Else                                                                                                                                'encontrou
        rs.Open "SELECT * from MDT where " & nomeColunaObjIdMdt & " = " & idObjPonto1, Conn, adOpenKeyset, adLockOptimistic
        If Not rs.EOF Then
            z1 = rs(0).value
        End If
        rs.Close
        
        rs.Open "SELECT * from mdt where " & nomeColunaObjIdMdt & " = " & idObjPonto2, Conn, adOpenKeyset, adLockOptimistic
        If Not rs.EOF Then
            z2 = rs(0).value
        End If
        rs.Close
        
        rs.Open "SELECT * from MDT where " & nomeColunaObjIdMdt & " = " & idObjPonto3, Conn, adOpenKeyset, adLockOptimistic
        If Not rs.EOF Then
            z3 = rs(0).value
        End If
        rs.Close
        
        rs.Open "SELECT * from MDT where " & nomeColunaObjIdMdt & " = " & idObjPonto4, Conn, adOpenKeyset, adLockOptimistic
        If Not rs.EOF Then
            z4 = rs(0).value
        End If
        rs.Close
        
        z = obtemZdoNo.CalcInterpolacao(xNo, yNo, xInt1, yInt1, xInt2, yInt2, xInt3, yInt3, xInt4, yInt4, z1, z2, z3, z4)
        ObtemZ = Round(z, 2)                                    'arredonda com apenas duas casas decimais
    End If
    Exit Function

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
    Resume Next
    Else
    ErroUsuario.Registra "CObtemZdoNo", "ObtemZ", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails, mensagemDeErro
    End If
End Function
' Obtem o z de um nó a partir do Object_id do nó
'
' objId - do nó com o qual precisa-se procurar o z
' layerGeometriasNos - nome do layer de geometrias dos nós da rede de água ou esgoto
'
Public Function CObtemZFromObjIdNo(objId As String, layerGeometriasNos As String) As Double
    On Error GoTo Trata_Erro
    Dim RSNos As ADODB.Recordset                            'para localizar as coordenadas x e y do nó que está sendo inserido
    Dim xNo As Double                                       'coordenada x do nó que está sendo inserido
    Dim yNo As Double                                       'coordenada y do nó que está sendo inserido
    Dim zNo As Double                                       'cota z do nó inserido que será calculada
    Dim getZNo As New CObtemZdoNo                           'método para o cálculo da cota z do nó inserido
    Dim retorno As Boolean
    
    Set RSNos = New ADODB.Recordset
    RSNos.Open "Select x, y from " & layerGeometriasNos & " where object_id = " & objId, Conn, adOpenKeyset, adLockOptimistic         'Localiza na tabela de geometria as coordenadas do nó a partir do object_id. Points2 - Água. Points4 - Esgoto
    If RSNos.EOF = False Then
        xNo = RSNos.Fields("x").value
        yNo = RSNos.Fields("y").value
    End If
    RSNos.Close
    retorno = cGeoDatabase.geoDatabase.setCurrentLayer("mdt")                                                   'ativa o layer do MDT
    If retorno = False Then
        MsgBox ("É necessário que o layer MDT esteja adicionado na vista e ligado para processar a cota.")
        CObtemZFromObjIdNo = 0
    Else
        zNo = getZNo.ObtemZ(xNo, yNo, objId)                                                                               'localiza o z desta nó
        CObtemZFromObjIdNo = zNo
    End If
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CObtemZdoNo", "CObtemZFromObjIdNo", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function
