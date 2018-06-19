Attribute VB_Name = "mdlQuerys"
 Option Explicit
 Dim aa As String
   Dim bb As String
   Dim sql As String, lista() As String
   Dim sa As String
   Dim sb As String
   Dim sc As String
   Dim sd As String
   Dim se As String
   Dim sf As String
   Dim sg As String
   Dim sh As String
   Dim si As String
   Dim sj As String
   Dim sl As String
   Dim sm As String
   Dim sn As String
   Dim so As String
   Dim sp As String
   Dim sq As String
   Dim sr As String
   Dim ss As String
   Dim st As String
   Dim su As String
   Dim sv As String
   Dim sx As String
   Dim sz As String
   Dim sk As String
   Dim sw As String
   Dim swx As String
   Dim sss As String
   Dim ssv As String
   Dim ssz As String
    Dim ssr As String
    Dim ssa As String
    Dim sst As String
    Dim ssq As String
    Dim sse As String
    Dim ssj As String
    Dim ssd As String
    Dim aspas As String
Public tipoConex As Integer
Dim aspasduplas As String


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'FUNÇÕES PARA LER E GRAVAR NO ARQUIVO .INI-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



Public Function getPmsdp(layerName As String, TypeQuery As Integer, object_id As String, tipoProvedor As Integer) As String
aspasduplas = """"
        aspas = ""
tipoConex = tipoProvedor

'--PMSDP - Properties Manager Select Default Properties
'
'/*Existem 3 Tipos de saidas que serão determinas pela entrada de um parametro para n layers

'0 - Single Select
'1 - Multiple Select with Alias
'2 - Single Select with alias
'3 - To Insert

'*/
'
'--#####################################################################
'--# ATENÇAO: #
'--# #
'--# 1 - Ao Modificar um campos(INSERIR,MODIFICAR OU EXCLUIR) de um Layer, #
'--# deverá faze-lo igualmente para todas as saídas do respectivo layer #
'--# 2 - Ao Criar um novo Layer será necessário a inserção de todas as saídas #
'--# #
'--#####################################################################

On Error GoTo Trata_Erro

Dim sql As String, FieldName As String
sa = "INITIALGROUNDHEIGHT"
   sb = "FINALGROUNDHEIGHT"
   sc = "INITIALTUBEDEEPNESS"
   sd = "FINALTUBEDEEPNESS"
   se = "INTERNALDIAMETER"
   sf = "EXTERNALDIAMETER"
   sg = "INITIALCOMPONENT"
   sh = "FINALCOMPONENT"
   si = "THICKNESS"
   sj = "MATERIAL"
   sl = "LENGTH"
   sm = "LENGTHCALCULATED"
   sn = "SUPPLIER"
   so = "MANUFACTURER"
   sp = "LOCATION"
   sq = "STATE"
   sr = "ROUGHNESS"
   ss = "SECTOR"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   sv = "SIDESTREET"
   sx = "DIVIDEDDISTANCE"
   sz = "USUARIO_LOG"
   sk = "DATA_LOG"
   sw = "OBJECT_ID_"
   aa = "SEWERLINES"
   bb = "DRAINLINES"
   swx = "ID_TYPE"
   sss = "YEAROFCONSTRUCTION"
   ssv = "NOTES"
   ssz = "TROUBLE"
   ssr = "PATTERN"
   ssa = "COMPONENT_ID"
   sst = "WATERLINES"
   ssq = "WATERCOMPONENTS"
   sse = "DEMAND"
   ssj = "CALCULENODE"
   ssd = "LINE_ID"

Select Case UCase(layerName) 'LAYER REDES
    
   Case UCase("waterlines"), UCase("drainlines"), UCase("sewerlines")
    If tipoConex <> 4 Then

        sql = sql & " SELECT Line_Id as " + aspasduplas + "LINHA" + aspasduplas + ","
        sql = sql & " Id_Type as TIPO,"

        sql = sql & " InitialGroundHeight as " + aspasduplas + "[TERRENO - COTA INICIAL]" + aspasduplas + ","
        sql = sql & " FinalGroundHeight as " + aspasduplas + "[TERRENO - COTA FINAL]" + aspasduplas + ","
        sql = sql & " InitialTubeDeepness as " + aspasduplas + "[PEÇA - COTA INICIAL]" + aspasduplas + ","
        sql = sql & " FinalTubeDeepness as " + aspasduplas + "[PEÇA - COTA FINAL]" + aspasduplas + ","

        sql = sql & " InternalDiameter as " + aspasduplas + "[DIAMETRO INTER.(mm)]" + aspasduplas + ","
        sql = sql & " ExternalDiameter as " + aspasduplas + "[DIAMETRO EXT.(mm)]" + aspasduplas + ","
        sql = sql & " InitialComponent as " + aspasduplas + "[INICIAL COMPONENTE]" + aspasduplas + ","
        sql = sql & " FinalComponent as " + aspasduplas + "[FINAL COMPONENTE]" + aspasduplas + ","
        sql = sql & " Thickness as " + aspasduplas + "DENSIDADE" + aspasduplas + ","
        sql = sql & " MATERIAL,"
        sql = sql & " Length as " + aspasduplas + "[COMPRIMENTO(m)]" + aspasduplas + ","
        sql = sql & " LengthCalculated as " + aspasduplas + "[COMPR. CALCULADO]" + aspasduplas + ","
        sql = sql & " Supplier as " + aspasduplas + "FORNECEDOR" + aspasduplas + ","
        sql = sql & " Manufacturer as " + aspasduplas + "FABRICANTE" + aspasduplas + ","
        sql = sql & " Location as " + aspasduplas + "LOCALIZAÇÃO" + aspasduplas + ","
        sql = sql & " State as " + aspasduplas + "ESTADO" + aspasduplas + ","
        
        'colocado o campo rugosidade para todos TIPOs de encanamento
        sql = sql & " RoughNess as " + aspasduplas + "RUGOSIDADE" + aspasduplas + "," 'inserido em 19/08 - Jonathas
        
        sql = sql & " Sector as " + aspasduplas + "SETOR" + aspasduplas + ","
        sql = sql & " InformationValidity As " + aspasduplas + "VALIDADE" + aspasduplas + ","
        sql = sql & " DateInstallation As " + aspasduplas + "[DATA_DE_INSTALAÇÃO]" + aspasduplas + ", SideStreet as " + aspasduplas + "[LADO_DA_RUA]" + aspasduplas + ", DividedDistance as " + aspasduplas + "[DISTÂNCIA_DA_DIVISA]" + aspasduplas + ""
        
        sql = sql & ", USUARIO_LOG as " + aspasduplas + "USUÁRIO" + aspasduplas + ", DATA_LOG AS " + aspasduplas + "[DATA CADASTRO]" + aspasduplas + "" '*****
        
        sql = sql & " From " & layerName
         'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
        Else 'alterado em 21/10/2010
        
        sql = sql & " SELECT " + """" + ssd + """" + "as " + """" + "LINHA" + """" + ","
        sql = sql & " " + """" + swx + """" + " as" + """" + "TIPO" + """" + ","

        sql = sql & "" + """" + sa + """" + " as" + """" + " [TERRENO - COTA INICIAL]" + """" + ","
        sql = sql & " " + """" + sb + """" + " as" + """" + " [TERRENO - COTA FINAL]" + """" + ","
        sql = sql & " " + """" + sc + """" + " as" + """" + " [PEÇA - COTA INICIAL]" + """" + ","
        sql = sql & " " + """" + sd + """" + " as" + """" + " [PEÇA - COTA FINAL]" + """" + ","

        sql = sql & " " + """" + se + """" + " as" + """" + " [DIAMETRO INTER.(mm)]" + """" + ","
        sql = sql & " " + """" + sf + """" + " as" + """" + " [DIAMETRO EXT.(mm)]" + """" + ","
        sql = sql & " " + """" + sg + """" + " as" + """" + " [INICIAL COMPONENTE]" + """" + ","
        sql = sql & " " + """" + sh + """" + " as" + """" + " [FINAL COMPONENTE]" + """" + ","
        sql = sql & " " + """" + si + """" + " as" + """" + " DENSIDADE" + """" + ","
        sql = sql & " " + """" + sj + """" + ","
        sql = sql & " " + """" + sl + """" + " as" + """" + " [COMPRIMENTO(m)]" + """" + ","
        sql = sql & " " + """" + sm + """" + " as" + """" + " [COMPR. CALCULADO]" + """" + ","
        sql = sql & " " + """" + sn + """" + " as" + """" + " FORNECEDOR" + """" + ","
        sql = sql & " " + """" + so + """" + " as" + """" + " FABRICANTE" + """" + ","
        sql = sql & " " + """" + sp + """" + " as" + """" + " LOCALIZAÇÃO" + """" + ","
        sql = sql & " " + """" + sq + """" + " as" + """" + " ESTADO" + """" + ","
        
        'colocado o campo rugosidade para todos TIPOs de encanamento
        sql = sql & " " + """" + sr + """" + " as" + """" + " RUGOSIDADE" + """" + "," 'inserido em 19/08 - Jonathas
        
        sql = sql & " " + """" + ss + """" + " as" + """" + " SETOR" + """" + ","
        sql = sql & " " + """" + st + """" + " as" + """" + " VALIDADE" + """" + ","
        sql = sql & " " + """" + su + """" + " as" + """" + " [DATA_DE_INSTALAÇÃO]" + """" + ", " + """" + sv + """" + " as" + """" + " [LADO_DA_RUA]" + """" + ", " + """" + sx + """" + " as" + """" + " [DISTÂNCIA_DA_DIVISA] " + """" + ""
        
        sql = sql & ", " + """" + sz + """" + " as" + """" + " USUÁRIO" + """" + ", " + """" + sk + """" + " AS" + """" + " [DATA CADASTRO] " + """" + "" '*****
        
        sql = sql & " From " + """" + layerName + """" + ""


 
        'PARA QUE O GRID CARREGUE COM AS INFORMAÇÕES DEFAULT É NECESSÁRIO QUE A TABELA X_STATE POSSUA
        
        'Criar a estrutura mínima na tabela X_STATE com valor StateID = 1 e StateName = Normal

        'CASO CONTRÁRIO, O LOG DE USUÁRIO E DATA DE DESENHO NÃO FUNCIONARÁ CORRETAMENTE
        End If
       
        Select Case TypeQuery
            
            Case 0
                'colocado o campo rugosidade para todos TIPOs de encanamento
                'retirado INICIAL COTA TERRENO;FINAL COTA TERRENO;FINAL PROFUNDIDADE;INICIAL PROFUNDIDADE PARA SEWERLINES

 If tipoConex <> 4 Then
                     sql = "select Line_Id,id_Type,InitialGroundHeight,FinalGroundHeight,initialTubeDeepness,FinalTubeDeepness,InternalDiameter,ExternalDiameter,InitialComponent,FinalComponent,Thickness,MATERIAL,Length,LengthCalculated,Supplier,Manufacturer,Location,State,RoughNess,Sector,InformationValidity, DateInstallation, SideStreet, DividedDistance from " & layerName & " where object_id_ = '0'"
                     Else
                        sql = "select " + """" + ssd + """" + "," + """" + swx + """" + "," + """" + sa + """" + "," + """" + sb + """" + "," + """" + sc + """" + "," + """" + sd + """" + "," + """" + se + """" + "," + """" + sf + """" + "," + """" + sg + """" + "," + """" + sh + """" + "," + """" + si + """" + "," + """" + sj + """" + "," + """" + sl + """" + "," + """" + sm + """" + "," + """" + sn + """" + "," + """" + so + """" + "," + """" + sp + """" + "," + """" + sq + """" + "," + """" + sr + """" + "," + """" + ss + """" + "," + """" + st + """" + ", " + """" + su + """" + ", " + """" + sv + """" + ", " + """" + sx + """" + " from "" +""""+ layerName +""""+ "" where " + """" + sw + """" + " = '0'"
                     
                     
                     End If
                
                'sql = "select Line_Id,id_Type,InitialGroundHeight,FinalGroundHeight,initialTubeDeepness,FinalTubeDeepness,InternalDiameter,ExternalDiameter,InitialComponent,FinalComponent,Thickness,MATERIAL,Length,LengthCalculated,Supplier,Manufacturer,Location,State," & IIf(UCase(layerName) = UCase("waterlines"), "RoughNess,", "") & "Sector,InformationValidity, DateInstallation, SideStreet, DividedDistance from " & layerName & " where object_id_ = '0'"
                'comando original
                'sql = "select Line_Id,id_Type,InitialGroundHeight,FinalGroundHeight,initialTubeDeepness,FinalTubeDeepness,InternalDiameter,ExternalDiameter,InitialComponent,FinalComponent,Thickness,MATERIAL,Length,LengthCalculated,Supplier,Manufacturer,Location,State," & IIf(layerName = "waterlines", "RoughNess,", "") & "Sector,InformationValidity, DateInstallation, SideStreet, DividedDistance from " & layerName & " where object_id_ = '0'"
            Case 1 'Multiple Select with Alias
                
 If tipoConex <> 4 Then
                'sql = sql & " Where OBJECT_ID_ in (" & object_id & ")"
                
                sql = sql & " Where line_id in (" & object_id & ")"  ' ANTERIOR A 19/10/2009
                Else
                
                sql = sql & " Where " + """" + ssd + """" + " in ('" + Round(object_id) + "')"  ' ANTERIOR A 19/10/2009
                End If
            Case 2 'Single Select with alias
                 If tipoConex <> 4 Then
                
                'sql = sql & " Where OBJECT_ID_ in (" & object_id & ")"
                
                sql = sql & " Where line_id in (" & object_id & ")" ' ANTERIOR A 19/10/2009
                Else
                 sql = sql & " Where " + """" + ssd + """" + " in ('" + object_id + "')"  ' ANTERIOR A 19/10/2009
                
                End If
                
            Case 3 'Default
         
              
              
               
                
           ' Case UCase("watercomponents"), UCase("sewercomponents"), UCase("draincomponents")
                
                
            'QUANDO É SELECIONADO 'DESENHAR REDES', ESTE SELECT ABAIXO CARREGA AS TAGS NO GERENCIADOR DE ATRIBUTOS
            If tipoConex <> 4 Then

                           sql = "Select '0' as" + """" + "LINHA" + """" + ","
                     sql = sql & " '0' as" + """" + "TIPO" + """" + ","
                     sql = sql & "'0' as" + """" + "[TERRENO - COTA INICIAL]" + """" + ","
                     sql = sql & "'0' as" + """" + "[TERRENO - COTA FINAL]" + """" + ","
                     sql = sql & "'0' as" + """" + "[PEÇA - COTA INICIAL]" + """" + ","
                     sql = sql & "'0' as" + """" + "[PEÇA - COTA FINAL]" + """" + ","
                     sql = sql & "'0' as" + """" + "[DIAMETRO INTERNO]" + """" + ","
                     sql = sql & "'0' as" + """" + "[DIAMETRO EXTERNO]" + """" + ","
                     sql = sql & "'0' as" + """" + "[INICIAL COMPONENTE]" + """" + ","
                     sql = sql & "'0' as" + """" + "[FINAL COMPONENTE]" + """" + ","
                     sql = sql & "'0' as" + """" + "DENSIDADE" + """" + ","
                     sql = sql & "'0' as" + """" + "MATERIAL " + """" + ","
                     sql = sql & "'0' as" + """" + "COMPRIMENTO" + """" + ","
                     sql = sql & "'0' as" + """" + "[COMPR. CALCULADO]" + """" + ","
                     sql = sql & "'0' as" + """" + "FORNECEDOR" + """" + ","
                     sql = sql & "'0' as" + """" + "FABRICANTE" + """" + ","
                     sql = sql & "'0' as" + """" + "LOCALIZAÇÃO" + """" + ","
                     sql = sql & "'0' as" + """" + "ESTADO" + """" + ","
                     sql = sql & "'0' as" + """" + "RUGOSIDADE" + """" + ","
                     sql = sql & "'0' as" + """" + "SETOR" + """" + "," + "''" + " as" + """" + " VALIDADE" + """" + ","
                     sql = sql & "'' as" + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ","
                     sql = sql & "'1' as" + """" + "[LADO_DA_RUA]" + """" + ","
                     sql = sql & "'0' as" + """" + "[DISTÂNCIA_DA_DIVISA]" + """" + ","
                     sql = sql & "'' as" + """" + "USUÁRIO" + """" + ","
                     sql = sql & "'' as " + """" + "[DATA CADASTRO]" + """" + ""
                     sql = sql & "from x_state where stateid =1"
                     
                     Else ' alterado em 21/10/2010
                     
                     Dim cf As String
                     Dim cg As String
                     cf = "X_STATE"
                     cg = "STATEID"
                     sql = "Select '0' as" + """" + "LINHA" + """" + ","
                     sql = sql & " '0' as" + """" + "TIPO" + """" + ","
                     sql = sql & "'0' as" + """" + "[TERRENO - COTA INICIAL]" + """" + ","
                     sql = sql & "'0' as" + """" + "[TERRENO - COTA FINAL]" + """" + ","
                     sql = sql & "'0' as" + """" + "[PEÇA - COTA INICIAL]" + """" + ","
                     sql = sql & "'0' as" + """" + "[PEÇA - COTA FINAL]" + """" + ","
                     sql = sql & "'0' as" + """" + "[DIAMETRO INTERNO]" + """" + ","
                     sql = sql & "'0' as" + """" + "[DIAMETRO EXTERNO]" + """" + ","
                     sql = sql & "'0' as" + """" + "[INICIAL COMPONENTE]" + """" + ","
                     sql = sql & "'0' as" + """" + "[FINAL COMPONENTE]" + """" + ","
                     sql = sql & "'0' as" + """" + "DENSIDADE" + """" + ","
                     sql = sql & "'0' as" + """" + "MATERIAL " + """" + ","
                     sql = sql & "'0' as" + """" + "COMPRIMENTO" + """" + ","
                     sql = sql & "'0' as" + """" + "[COMPR. CALCULADO]" + """" + ","
                     sql = sql & "'0' as" + """" + "FORNECEDOR" + """" + ","
                     sql = sql & "'0' as" + """" + "FABRICANTE" + """" + ","
                     sql = sql & "'0' as" + """" + "LOCALIZAÇÃO" + """" + ","
                     sql = sql & "'0' as" + """" + "ESTADO" + """" + ","
                     sql = sql & "'0' as" + """" + "RUGOSIDADE" + """" + ","
                     sql = sql & "'0' as" + """" + "SETOR" + """" + "," + "''" + " as" + """" + " VALIDADE" + """" + ","
                     sql = sql & "'' as" + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ","
                     sql = sql & "'1' as" + """" + "[LADO_DA_RUA]" + """" + ","
                     sql = sql & "'0' as" + """" + "[DISTÂNCIA_DA_DIVISA]" + """" + ","
                     sql = sql & "'' as" + """" + "USUÁRIO" + """" + ","
                     sql = sql & "'' as " + """" + "[DATA CADASTRO]" + """" + ""
                     sql = sql & "from " + """" + cf + """" + " where " + """" + cg + """" + " ='1'"
                     End If
                 'MsgBox sql

                   End Select

                
       
        
        'MsgBox sql
        
''''        MsgBox TypeQuery
''''        Select Case TypeQuery
''''
''''            Case 0
''''                'colocado o campo rugosidade para todos TIPOs de encanamento
''''                sql = "select Line_Id,id_Type,InitialGroundHeight,FinalGroundHeight,initialTubeDeepness,FinalTubeDeepness,InternalDiameter,ExternalDiameter,InitialComponent,FinalComponent,Thickness,MATERIAL,Length,LengthCalculated,Supplier,Manufacturer,Location,State,RoughNess,Sector,InformationValidity, DateInstallation, SideStreet, DividedDistance from " & layerName & " where object_id_ = '0'"
''''                'sql = "select Line_Id,id_Type,InitialGroundHeight,FinalGroundHeight,initialTubeDeepness,FinalTubeDeepness,InternalDiameter,ExternalDiameter,InitialComponent,FinalComponent,Thickness,MATERIAL,Length,LengthCalculated,Supplier,Manufacturer,Location,State," & IIf(UCase(layerName) = UCase("waterlines"), "RoughNess,", "") & "Sector,InformationValidity, DateInstallation, SideStreet, DividedDistance from " & layerName & " where object_id_ = '0'"
''''                'comando original
''''                'sql = "select Line_Id,id_Type,InitialGroundHeight,FinalGroundHeight,initialTubeDeepness,FinalTubeDeepness,InternalDiameter,ExternalDiameter,InitialComponent,FinalComponent,Thickness,MATERIAL,Length,LengthCalculated,Supplier,Manufacturer,Location,State," & IIf(layerName = "waterlines", "RoughNess,", "") & "Sector,InformationValidity, DateInstallation, SideStreet, DividedDistance from " & layerName & " where object_id_ = '0'"
''''            Case 1 'Multiple Select with Alias
''''                sql = sql & " Where line_id in (" & object_id & ")"
''''            Case 2 'Single Select with alias
''''                sql = sql & " Where line_id in (" & object_id & ")"
''''            Case 3 'tupla default
''''                'colocado o campo rugosidade para todos TIPOs de encanamento
''''
''''                'original
''''                'sql = "Select 0 as LINHA, 0 as TIPO, 0 as [INICIAL COTA TERRENO], 0 as [FINAL COTA TERRENO], 0 as [INICIAL PROFUNDIDADE], 0 as  [FINAL PROFUNDIDADE], 0 as [DIAMETRO INTERNO], 0 as [DIAMETRO EXTERNO], 0 as [INICIAL COMPONENTE], 0 as [FINAL COMPONENTE], 0 as DENSIDADE, 0 as MATERIAL , 0 as COMPRIMENTO, 0 as [COMPR. CALCULADO],0 as FORNECEDOR, 0 as FABRICANTE, 0 as LOCALIZAÇÃO, 0 as ESTADO, 0 as RUGOSIDADE, 1 as SETOR, '' as VALIDADE, '' As [DATA_DE_INSTALAÇÃO], 1 as [LADO_DA_RUA], 0 as [DISTÂNCIA_DA_DIVISA] from x_state where stateid  =1"
''''
''''                'novo
''''                sql = "Select 0 as LINHA, 0 as TIPO, 0 as [INICIAL COTA TERRENO], 0 as [FINAL COTA TERRENO], 0 as [INICIAL PROFUNDIDADE], 0 as  [FINAL PROFUNDIDADE], 0 as [DIAMETRO INTERNO], 0 as [DIAMETRO EXTERNO], 0 as [INICIAL COMPONENTE], 0 as [FINAL COMPONENTE], 0 as DENSIDADE, 0 as MATERIAL , 0 as COMPRIMENTO, 0 as [COMPR. CALCULADO],0 as FORNECEDOR, 0 as FABRICANTE, 0 as LOCALIZAÇÃO, 0 as ESTADO, 0 as RUGOSIDADE, 1 as SETOR, '' as VALIDADE, '' As [DATA_DE_INSTALAÇÃO], 1 as [LADO_DA_RUA], 0 as [DISTÂNCIA_DA_DIVISA], '' as USUÁRIO, '' AS [DATA CADASTRO] from x_state where stateid  =1"
''''
''''
''''                'sql = "Select 0 as LINHA, 0 as TIPO, 0 as [INICIAL COTA TERRENO], 0 as [FINAL COTA TERRENO], 0 as [INICIAL PROFUNDIDADE], 0 as  [FINAL PROFUNDIDADE], 0 as [DIAMETRO INTERNO], 0 as [DIAMETRO EXTERNO], 0 as [INICIAL COMPONENTE], 0 as [FINAL COMPONENTE], 0 as DENSIDADE, 0 as MATERIAL , 0 as COMPRIMENTO, 0 as [COMPR. CALCULADO],0 as FORNECEDOR, 0 as FABRICANTE, 0 as LOCALIZAÇÃO, 0 as ESTADO, " & IIf(UCase(layerName) = UCase("waterlines"), "0 as Rugosidade,", "") & " 1 as SETOR, '' as VALIDADE, '' As [DATA_DE_INSTALAÇÃO], 1 as [LADO_DA_RUA], 0 as [DISTÂNCIA_DA_DIVISA] from x_state where stateid  =1"
''''        End Select
            
'        '************* MONITORAMENTO ***************
'        Close #2
'        Open App.Path & "\GeoSanLog.txt" For Append As #2
'        Print #2, Now & "Public Function getPmsdp - case 1 SQL = " & sql & " TIPO select " & TypeQuery
'        Close #2
'        '***************** FIM *********************
    

      
    Case UCase("watercomponents"), UCase("sewercomponents"), UCase("draincomponents")
            If tipoConex <> 4 Then
            
        sql = sql & " Select Component_id as " + """" + "COMPONENTE" + """" + ","
        sql = sql & " ID_Type as " + """" + "TIPO" + """" + ","
        sql = sql & " YearOfConstruction as " + """" + "[ANO DE FABRICAÇÃO]" + """" + ","
        sql = sql & " State as " + """" + "ESTADO" + """" + ","
        sql = sql & " Location as " + """" + "LOCALIZAÇÃO" + """" + ","
        sql = sql & " Supplier as " + """" + "FORNECEDOR" + """" + ","
        sql = sql & " Manufacturer as " + """" + "FABRICANTE" + """" + " ,"
        sql = sql & " GroundHeight as " + """" + "[COTA DO TERRENO]" + """" + ","
        
        If UCase(layerName) = "SEWERCOMPONENTS" Or UCase(layerName) = "DRAINCOMPONENTS" Then
            sql = sql & " GroundHeightFinal as " + """" + "[COTA DO FUNDO]" + """" + ","
        End If
        
        If UCase(layerName) = UCase("watercomponents") Then
            sql = sql & " Demand as " + """" + "DEMANDA" + """" + ","
            sql = sql & " CalculeNode as " + """" + "[NÓ DE CÁLCULO]" + """" + ","
        End If
        
        sql = sql & " InformationValidity as " + """" + "VALIDADE" + """" + ","
        sql = sql & " Notes As " + """" + "Observação" + """" + ","
        sql = sql & " Trouble as " + """" + "[NÃO_CONFORMIDADE]" + """" + ", DateInstallation As " + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ", Pattern as " + """" + "[PADRÃO_CONSUMO]" + """" + ", Sector as " + """" + "[SETOR]" + """"
        'sql = sql & " ANGLE,NOME_CELUL,ORIGEM_CAL,X_,Y_,COR,TAMANHO_X,TAMANHO_Y,CENT_CEL_X,CENT_CEL_Y,COR_CELULA,ESC_CEL_X,ESC_CEL_Y "
        sql = sql & " From " & layerName
        
        Select Case TypeQuery
            Case 0
                sql = "select Component_id , id_Type, YearOfConstruction, State, Location, Supplier, Manufacturer, GroundHeight" & IIf(UCase(layerName) <> UCase("watercomponents"), ", GroundHeightFinal", ", Demand, calculenode ") & ", InformationValidity, Notes,Trouble, DateInstallation, Pattern, Sector from " & layerName & " where object_id_ = 1"
            Case 1 'Multiple Select with Alias
                sql = sql & " Where Component_id in (" & object_id & ")"
            Case 2 'Single Select with alias
            
                sql = sql & " Where Component_id in (" & object_id & ")"
                
                
'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
            Case 3 'tupla default
               ' sql = "Select  0 as "+""""+"COMPONENTE"+""""+", 0 as "+""""+"TIPO"+""""+", 0 as "+""""+"[ANO DE FABRICAÇÃO]"+""""+", 0 as "+""""+"ESTADO"+""""+", 0 as "+""""+"LOCALIZAÇÃO"+""""+", 0 as "+""""+"FORNECEDOR"+""""+", 0 as "+""""+"FABRICANTE"+""""+" , 0 as "+""""+"[COTA DO TERRENO]"+"""" & IIf(UCase(layerName) <> UCase("watercomponents"), ", 0 as "+""""+"[COTA DO FUNDO]"+""""+ ", 0 as "+""""+"DEMANDA"+""""+", 0 as "+""""+"[NÓ DE CÁLCULO]"+""""+") & ", 0 as "+""""+"VALIDADE"+""""+", '' as "+""""+"Observação"+""""+", 0 as "+""""+"[NÃO_CONFORMIDADE]"+""""+", '' As "+""""+"[DATA_DE_INSTALAÇÃO]"+""""+", '' as "+""""+"[PADRÃO_CONSUMO]"+""""+", '' as "+""""+"[SETOR]"+""""+" from x_state where stateid  =1"
                
                
                sql = "Select 0 as " + aspasduplas + "LINHA" + aspasduplas + ",  0 as " + aspasduplas + "TIPO" + aspasduplas + ", 0 as " + aspasduplas + "[TERRENO - COTA INICIAL]" + aspasduplas + ", 0 as " + aspasduplas + "[TERRENO - COTA FINAL]" + aspasduplas + ", 0 as " + aspasduplas + "[PEÇA - COTA INICIAL]" _
                + aspasduplas + ", 0 as " + aspasduplas + "[PEÇA - COTA FINAL]" + aspasduplas + ", 0 as " + aspasduplas + "[DIAMETRO INTERNO]" + aspasduplas + ", 0 as " + aspasduplas + "[DIAMETRO EXTERNO]" + aspasduplas + ", 0 as " + aspasduplas + "[INICIAL COMPONENTE]" + aspasduplas + ", 0 as " + aspasduplas + "[FINAL COMPONENTE]" + aspasduplas + ", 0 as " + aspasduplas + "DENSIDADE" _
                + aspasduplas + ", 0 as " + aspasduplas + "MATERIAL" + aspasduplas + " , 0 as " + aspasduplas + "COMPRIMENTO" + aspasduplas + ", 0 as " + aspasduplas + "[COMPR. CALCULADO]" + aspasduplas + ", 0 as " + aspasduplas + "FORNECEDOR" + aspasduplas + ", 0 as " + aspasduplas + "FABRICANTE" + aspasduplas + ", 0 as " + aspasduplas + "LOCALIZAÇÃO" + aspasduplas + ", 0 as " _
                + aspasduplas + "ESTADO" + aspasduplas + ", 0 as " + aspasduplas + "RUGOSIDADE" + aspasduplas + ", 0 as " + aspasduplas + "SETOR" + aspasduplas + ", '' as " + aspasduplas + "VALIDADE" + aspasduplas + ", '' As " + aspasduplas + "[DATA_DE_INSTALAÇÃO]" + aspasduplas + ", 1 as " + aspasduplas + "[LADO_DA_RUA]" + aspasduplas + ", 0 as " + aspasduplas + "[DISTÂNCIA_DA_DIVISA]" + aspasduplas + ", '' as " + aspasduplas + "USUÁRIO" + aspasduplas + ", '' AS " + aspasduplas + "[DATA CADASTRO]" + aspasduplas + " from x_state where stateid  =1"
     
     
        
         Case Else 'qualque plano
        sql = "SELECT * FROM " & layerName & " WHERE OBJECT_ID_ in(" & object_id & ")"
   End Select


    Else 'alterado em 21/10/2010
    
    Dim ja As String
    Dim je As String
     Dim jo As String
     Dim a As String
    ja = "SEWERCOMPONENTS"
    je = "DRAINCOMPONENTS"
    ssq = "WATERCOMPONENTS"
    jo = "OBJECT_ID_"
    swx = "ID_TYPE"
    sss = "YEAROFCONSTRUCTION"
    sq = "STATE"
    sp = "LOCATION"
    sn = "SUPPLIER"
    so = "MANUFACTURER"
    sa = "INITIALGROUNDHEIGHT"
    sf = layerName
    sb = "FINALGROUNDHEIGHT"
    sse = "DEMAND"
    ssj = "CALCULENODE"
    ssv = "NOTES"
    ssz = "TROUBLE"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   ssr = "PATTERN"
    ss = "SECTOR"
    sw = "OBJECT_ID"
    a = "COMPONENT_ID"
    Dim man1 As String
        Dim man2 As String
        man1 = "STATEID"
        man2 = "X_TATE"
   ' Case UCase("+ssq+"), UCase("+ja+"), UCase("" + je + "")
       ' Case UCase("watercomponents"), UCase("sewercomponents"), UCase("draincomponents")
         
        sql = sql & " Select " + """" + jo + """" + " as" + """" + " COMPONENTE" + """" + ","
        sql = sql & " " + """" + swx + """" + " as" + """" + " TIPO" + """" + ","
        sql = sql & " " + """" + sss + """" + " as" + """" + " [ANO DE FABRICAÇÃO]" + """" + ","
        sql = sql & " " + """" + sq + """" + " as" + """" + " ESTADO" + """" + ","
        sql = sql & " " + """" + sp + """" + " as" + """" + " LOCALIZAÇÃO" + """" + ","
        sql = sql & " " + """" + sn + """" + " as" + """" + " FORNECEDOR" + """" + ","
        sql = sql & " " + """" + so + """" + " as" + """" + " FABRICANTE " + """" + ","
        sql = sql & " " + """" + sa + """" + " as" + """" + " [COTA DO TERRENO]" + """" + ","
        
        If UCase(layerName) = "SEWERCOMPONENTS" Or UCase(layerName) = "DRAINCOMPONENTS" Then
            sql = sql + """" + sb + """" + " as" + """" + "[COTA DO FUNDO]" + """" + ","
        End If
        
        If UCase(layerName) = UCase("watercomponents") Then
            sql = sql + """" + sse + """" + " as" + """" + " DEMANDA" + """" + ","
            sql = sql + """" + ssj + """" + " as" + """" + " [NÓ DE CÁLCULO]" + """" + ","
        End If
        
        sql = sql + """" + st + """" + " as" + """" + " VALIDADE" + """" + ","
        sql = sql + """" + ssv + """" + " as" + """" + " Observação" + """" + ","
        sql = sql + """" + ssz + """" + " as" + """" + " [NÃO_CONFORMIDADE]" + """" + ", " + """" + su + """" + " as " + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ", " + """" + ssr + """" + " as" + """" + " [PADRÃO_CONSUMO]" + """" + ", " + """" + ss + """" + " as" + """" + " [SETOR]" + """" + ""
        'sql = sql & " ANGLE,NOME_CELUL,ORIGEM_CAL,X_,Y_,COR,TAMANHO_X,TAMANHO_Y,CENT_CEL_X,CENT_CEL_Y,COR_CELULA,ESC_CEL_X,ESC_CEL_Y "
        sql = sql & " From " + """" + layerName + """"
        
        Select Case TypeQuery
            Case 0
                sql = "select " + """" + jo + """" + " , " + """" + swx + """" + ", " + """" + sss + """" + ", " + """" + sq + """" + ", " + """" + sp + """" + ", " + """" + sn + """" + ", " + """" + so + """" + ", " + """" + sa + """" + "" & IIf(UCase(layerName) <> UCase(" +""""+ ssq +""""+ "), ", " + """" + sb + """" + "", ", " + """" + sse + """" + ", " + """" + ssj + """" + " ") & ", " + """" + st + """" + ", " + """" + ssv + """" + "," + """" + ssz + """" + ", " + """" + su + """" + ", " + """" + ssr + """" + ", " + """" + ss + """" + " from "" & layerName & "" where " + """" + sw + """" + " = '1'"
            Case 1 'Multiple Select with Alias
                sql = sql & " Where " + """" + a + """" + " in ('" + object_id + "')"
            Case 2 'Single Select with alias
                sql = sql & " Where " + """" + a + """" + " in ('" + object_id + "')"
            Case 3 'tupla default

             sql = "Select  '0' as " + """" + " COMPONENTE " + """" + ", '0' as " + """" + "TIPO" + """" + ", '0' as " + """" + "[ANO DE FABRICAÇÃO]" + """" + ", '0' as " + """" + "ESTADO" + """" + ", '0' as " + """" + "LOCALIZAÇÃO" + """" + ", '0' as " + """" + "FORNECEDOR, '0' as FABRICANTE " + """" + ", '0' as " + """" + "[COTA DO TERRENO]" + """" + " " & IIf(layerName <> "WATERCOMPONENTS", ", '0' as " + """" + "[COTA DO FUNDO]" + """" + "", ", '0' as " + """" + "DEMANDA" + """" + ", '0' as " + """" + "[NÓ DE CÁLCULO]" + """" + "") & ", '0' as " + """" + "VALIDADE" + """" + ", '' as " + """" + "Observação" + """" + ", '0' as " + """" + "[NÃO_CONFORMIDADE]" + """" + ", '' As " + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ", '' as [PADRÃO_CONSUMO], '' as " + """" + "[SETOR]" + """" + " from " + """" + man2 + """" + " where " + """" + man1 + """" + "  ='1'"

    Case Else 'qualque plano
        sql = "SELECT * FROM " + """" + layerName + """" + " WHERE " + """" + sw + """" + " in('" + object_id + "')"
    End Select
      End If
        
End Select
      ' MsgBox "ARQUIVO DEBUG SALVO"
      ' WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
 
 'MsgBox "ARQUIVO DEBUG SALVO"
 
 
 
' WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
 
 
getPmsdp = convertQuery(sql, tipoProvedor)


''************* MONITORAMENTO ***************
'Close #2
'Open App.Path & "\GeoSanLog.txt" For Append As #2
'Print #2, Now & "Public Function getPmsdp - case 2 SQL = " & sql & " TIPO select " & TypeQuery
'Close #2
''***************** FIM *********************


Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
           ' MsgBox "ARQUIVO DEBUG SALVO"
            'WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function getPmsdp - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If


End Function


Public Function getPmssp(layerName As String, id_Type As Integer, object_id_ As String, conn As ADODB.Connection, tipoProvedor As Integer) As String
'verifica existencia de dados para sub-TIPO

On Error GoTo Trata_Erro
'On Error GoTo getPmssp_err
tipoConex = tipoProvedor
Dim rs As New ADODB.Recordset, sql As String

Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim ana As String
Dim ana2 As String
a = "OBJECT_ID_"
b = "DATA"
c = "SUBTYPES"
d = "ID_TYPE"
e = "ID_SUBTYPE"

ana = layerName + b
ana2 = layerName + c
'a = id_Type
'MsgBox a
If tipoConex <> 4 Then
sql = sql & " Select count(Object_id_) From " & layerName & "Data A "
sql = sql & " Inner join " & layerName & "SubTypes B On A.Id_Type = b.ID_Type and A.Id_SubType = b.ID_SubType"
sql = sql & " Where b.Id_Type =" & id_Type & " and Object_id_='" & object_id_ & "'"
rs.Open sql, conn
sql = ""
Else

'sql = "Select count(" + "" + "OBJECT_ID" + "" + ") From " + "WATERLINESDATA" + ""
'sql = "Select count(" + """" + "OBJECT_ID_" + """" + ") From " + """" + "WATERLINESDATA" + """" + "  Inner join " + """" + "WATERLINESSUBTYPES" + """" + " On " + """" + "WATERLINESDATA" + """" + "." + """" + "ID_TYPE" + """" + " = +""""+" + "WATERLINESSUBTYPES" + """" + "." + """" + "ID_TYPE" + """" + " and " + """" + "WATERLINESDATA" + """" + "." + """" + "ID_SUBTYPE" + """" + " = " + """" + "WATERLINESSUBTYPES" + """" + "." + """" + "ID_SUBTYPE" + """" + " Where " + """" + "WATERLINESSUBTYPES" + """" + "." + """" + "ID_TYPE" + """" + " ='0' and " + """" + "OBJECT_ID_" + """" + "='0'"

'sql = "Select count(" + a + ") From " + " + ana + " + "  Inner join " + " + ana2 + " + " On " + " + ana + " + "." + e + " = " + " + ana2 + " + "." + d + " and " + " + ana + " + "." + e + " = " + " + ana2 + " + "." + e + " Where " + " + ana2 + " + "." + d + " ='" + id_Type + "' and " + a + "='" + object_id_ + "'"
sql = sql & " Select count(" + """" + a + """" + ") From " + """" + ana + """" + " Inner join " + """" + ana2 + """" + " On " + """" + ana + """" + "." + """" + d + """" + " = " + """" + ana2 + """" + "." + """" + d + """" + " and " + """" + ana + """" + "." + """" + e + """" + " = " + """" + ana2 + """" + "." + """" + e + """" + " Where " + """" + ana2 + """" + "." + """" + d + """" + " ='" & id_Type & "'" + " and " + """" + a + """" + "='" & object_id_ & "'"

'sql = sql & " Select count(" + """" + a + """" + ") From " + """" + ana + """" + " Inner join " + """" + ana2 + """" + " On " + """" + ana + """" + "." + """" + d + """" + " = " + """" + ana2 + """" + "." + """" + d + """" + " and " + """" + ana + """" + "." + """" + e + """" + " = " + """" + ana2 + """" + "." + """" + e + """" + " Where " + """" + ana2 + """" + "." + """" + d + """" + " = '+id_Type+'" '+ "'" + " and " + """" + a + """" + "='" + object_id_ + "'"


          
          '  MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"


'MsgBox sql

rs.Open sql, conn, adOpenDynamic, adLockOptimistic
sql = ""


End If



Dim a1 As String
Dim b1 As String
Dim c1 As String
Dim d1 As String
Dim e1 As String
Dim a2 As String
Dim b2 As String
Dim c2 As String
Dim d2 As String
Dim e2 As String
Dim a3 As String
Dim b3 As String
Dim c3 As String
Dim d3 As String
Dim e3 As String
Dim e6 As String
Dim a4 As String
Dim b4 As String
Dim c4 As String
Dim d4 As String
Dim e4 As String

Dim aa As String
Dim bb As String


If Not rs.EOF Then


    If rs(0) > 0 Then
         If tipoProvedor = 2 Then
            'sql = sql & "Select '" & object_id_ & "', Selection_, Max_,Min_,DataType,B.Description_, A.Id_Type, B.Id_subType , A.Value_ as Value_Ref,"
            'sql = sql & " (Select Option_"
           ' sql = sql & " From WaterComponentsSelections C"
           ' sql = sql & " Where C.id_Type = A.id_Type"
           ' sql = sql & " and C.Id_SubType= A.Id_subType and C.Value_=A.Value_)as " + """" + "value_" + """" + ""
           ' sql = sql & " From watercomponentsData A"
           ' sql = sql & " Left Join watercomponentsSubTypes B"
           ' sql = sql & " On A.Id_Type = b.ID_Type and A.Id_SubType = b.ID_SubType"
           ' sql = sql & " Where Object_id_='" & object_id_ & "' and b.id_Type =" & id_Type
           ' sql = sql & " Union"
           ' sql = sql & " (Select '" & object_id_ & "',Selection_,Max_,Min_,DataType,A.Description_, A.Id_Type,A.Id_subType,A.DefaultValue as " + """" + "Value_Ref" + """" + ","
           ' sql = sql & " case  "
           ' sql = sql & " When Selection_= 1 then b.Option_"
           ' sql = sql & " when Selection_ <> 1 then A.DefaultValue"
           ' sql = sql & " End Value_"
           ' sql = sql & " From " & layerName & "SubTypes A Left Join " & layerName & "Selections B"
          '  sql = sql & " On A.Id_Type = B.Id_Type and A.Id_subType = B.Id_subType and B.Value_ = A.DefaultValue"
          '  sql = sql & " Where A.Id_Type = " & id_Type & " and a.id_subtype in(Select id_SubType from " & layerName & "Data where object_id_ = " & object_id_ & ")) ORDER BY Id_subType"
         
         
             sql = sql & "Select '" & object_id_ & "', Selection_, Max_,Min_,DataType,B.Description_, A.Id_Type, B.Id_subType , A.Value_ as Value_Ref,"
            sql = sql & " (Select Option_"
            sql = sql & " From WaterComponentsSelections C"
            sql = sql & " Where C.id_Type = A.id_Type"
            sql = sql & " and C.Id_SubType= A.Id_subType and C.Value_=A.Value_)as " + """" + "value_" + """" + ""
            sql = sql & " From watercomponentsData A"
            sql = sql & " Left Join watercomponentsSubTypes B"
            sql = sql & " On A.Id_Type = b.ID_Type and A.Id_SubType = b.ID_SubType"
            sql = sql & " Where Object_id_='" & object_id_ & "' and b.id_Type =" & id_Type
            
            
            'sql = sql & " Union"
           ' sql = sql & " (Select '" & object_id_ & "',Selection_,Max_,Min_,DataType,A.Description_, A.Id_Type,A.Id_subType,A.DefaultValue as " + """" + "Value_Ref" + """" + ","
           ' sql = sql & " case  "
           ' sql = sql & " When Selection_= 1 then b.Option_"
           ' sql = sql & " when Selection_ <> 1 then A.DefaultValue"
           ' sql = sql & " End Value_"
           ' sql = sql & " From " & layerName & "SubTypes A Left Join " & layerName & "Selections B"
           ' sql = sql & " On A.Id_Type = B.Id_Type and A.Id_subType = B.Id_subType inner join " & layerName & "data w on w.Value_ = B.Value_"
           ' sql = sql & " Where A.Id_Type = " & id_Type & "  and W.object_id_='" & object_id_ & "') ORDER BY Id_subType"
      
      
       
      
      
      
      
      
         
   '    Select '27553',Selection_,Max_,Min_,DataType,A.Description_, A.Id_Type,A.Id_subType,A.DefaultValue as "Value_Ref", case
     '  When Selection_= 1 then b.Option_ when Selection_ <> 1 then A.DefaultValue End Value_
'From WATERLINESSubTypes A Left Join WATERLINESSelections B On A.Id_Type = B.Id_Type and A.Id_subType =B.Id_subType inner
'join waterlinesdata w on w.Value_= B.Value_ Where A.Id_Type = 0
         
         ElseIf tipoProvedor = 1 Then
            sql = sql & " (Select '" & object_id_ & "', Selection_, Max_,Min_,DataType,B.Description_, A.Id_Type, B.Id_subType , A.Value_ as Value_Ref,"
            sql = sql & " Case Selection_ when 0 then A.Value_"
            sql = sql & " Else (Select Value_=Option_ From " & layerName & "Selections C"
            sql = sql & " where C.Id_Type= A.Id_Type and C.Id_SubType= A.Id_subType and C.Value_=A.Value_) end As Value_"
            sql = sql & " From " & layerName & "Data A"
            sql = sql & " Left Join " & layerName & "SubTypes B"
            sql = sql & " On A.Id_Type = b.ID_Type and A.Id_SubType = b.ID_SubType"
            sql = sql & " Where Object_id_='" & object_id_ & "' and b.id_Type =" & id_Type & ")"
            sql = sql & " Union"
            sql = sql & " (Select '" & object_id_ & "',Selection_,Max_,Min_,DataType,A.Description_, A.Id_Type,A.Id_subType,A.DefaultValue as Value_Ref,"
            sql = sql & " case Selection_ When 1 then b.Option_ else A.DefaultValue End as Value_"
            sql = sql & " From " & layerName & "SubTypes A Left Join " & layerName & "Selections B"
            sql = sql & " On A.Id_Type = B.Id_Type and A.Id_subType = B.Id_subType and B.Value_ = A.DefaultValue"
            sql = sql & " Where A.Id_Type = " & id_Type & " and a.id_subtype not in(Select id_SubType from " & layerName & "Data where object_id_ = " & object_id_ & "))"
            
                    


           
            
            
            Else
            

e6 = "OPTION_"
a = object_id_
b = "SELECTION_"
c = "MAX_"
d = "MIN_"
e = "DATATYPE"
a1 = "ID_TYPE"
b1 = "ID_SUBTYPE"
c2 = "VALUE_"
d2 = "SELECTION_"
e2 = "SUB_TYPES"
a3 = layerName
b3 = "SELECTIONS"
c3 = a3 + b3
d3 = "SUBTYPES"
e3 = "DESCRIPTION"
a4 = "DATA"
b4 = a3 + a4
c4 = "SUBTYPES"
d4 = a3 + c4
e4 = "DEFAULTVALUE"
Dim za66 As String
za66 = id_Type
Dim za67 As String
za67 = object_id_

'sql = sql & "Select " + a + "," + """" + "SELECTION_" + """" + "," + """" + "MAX_" + """" + "," + """" + "MIN_" + """" + "," + """" + "DATATYPE" + """" + "," + "B." + """" + "DESCRIPTION" + """" + "," + "A." + """" + "ID_TYPE" + """" + "," + "B." + """" + "ID_SUBTYPE" + """" + "," + "A." + """" + "VALUE_" + """" + " as " + """" + "Value_Ref" + """" + "," + " Case " + """" + "SELECTION_" + """" + " when '0' then CAST(A." + """" + "VALUE_" + """" + " AS BOOLEAN)"
'sql = sql & "Else (Select  CAST(A." + """" + "VALUE_" + """" + " AS VARCHAR) = " + """" + "OPTION_" + """" + " From " + """" + a3 + "SELECTIONS" + """" + " C where C." + """" + "ID_TYPE" + """" + " = A." + """" + "ID_TYPE" + """" + " and C." + """" + "ID_SUBTYPE" + """" + "= A." + """" + "ID_SUBTYPE" + """" + " and C." + """" + "VALUE_" + """" + "=CAST(A." + """" + "VALUE_" + """" + "As Integer))"
'sql = sql & "end As " + """" + "Value_" + """" + " From " + """" + a3 + "DATA" + """" + " A Left Join " + """" + a3 + "SUBTYPES" + """" + " B On A." + """" + "ID_TYPE" + """" + " = B." + """" + "ID_TYPE" + """" + " and A." + """" + "ID_SUBTYPE" + """" + " = B." + """" + "ID_SUBTYPE" + """" + " Where " + """" + "OBJECT_ID_" + """" + "='" + za67 + "' and B." + """" + "ID_TYPE" + """" + " = '" + za66 + "' Union (Select " + za67 + ","
'sql = sql & """" + "SELECTION_" + """" + "," + """" + "MAX_" + """" + "," + """" + "MIN_" + """" + "," + """" + "DATATYPE" + """" + "," + "A." + """" + "DESCRIPTION" + """" + ",A." + """" + "ID_TYPE" + """" + ",A." + """" + "ID_SUBTYPE" + """" + ",A." + """" + "DEFAULTVALUE" + """" + " as " + """" + "Value_Ref" + """" + ", Case " + """" + "SELECTION_" + """" + "  When '1' then CAST(B." + """" + "OPTION_" + """" + " AS BOOLEAN) else CAST(A." + """" + "DEFAULTVALUE" + """"
'sql = sql & " AS BOOLEAN) End as " + """" + "Value_" + """" + " From " + """" + a3 + "SUBTYPES" + """" + " A Left Join " + """" + a3 + "SELECTIONS" + """" + " B On A." + """" + "ID_TYPE" + """" + " = B." + """" + "ID_TYPE" + """" + " and A." + """" + "ID_SUBTYPE" + """" + " = B." + """" + "ID_SUBTYPE" + """" + " and B." + """" + "VALUE_" + """" + " = CAST(A." + """" + "DEFAULTVALUE" + """" + " AS INTEGER)"
'sql = sql & "Where A." + """" + "ID_TYPE" + """" + " = '0' And A." + """" + "ID_SUB"
'sql = sql & "TYPE" + """" + " not in(Select " + """" + "ID_SUBTYPE" + """" + " from " + """" + a3 + "DATA" + """" + " where " + """" + "OBJECT_ID_" + """" + " = '" + object_id_ + "'))"


If layerName = "DRAINLINES" Or layerName = "DRAINCOMPONENTS" Then
sql = sql & "Select " + a + "," + """" + "SELECTION_" + """" + "," + """" + "MAX_" + """" + "," + """" + "MIN_" + """" + "," + """" + "DATATYPE" + """" + "," + "B." + """" + "DESCRIPTION" + """" + "," + "A." + """" + "ID_TYPE" + """" + "," + "B." + """" + "ID_SUBTYPE" + """" + "," + "A." + """" + "VALUE_" + """" + " as " + """" + "Value_Ref" + """" + ","
sql = sql & "(Select  " + """" + "OPTION_" + """" + " From " + """" + a3 + "SELECTIONS" + """" + " C where C." + """" + "ID_TYPE" + """" + " = A." + """" + "ID_TYPE" + """" + " and C." + """" + "ID_SUBTYPE" + """" + "= A." + """" + "ID_SUBTYPE" + """" + " and C." + """" + "VALUE_" + """" + "=CAST(A." + """" + "VALUE_" + """" + "As Integer))"
sql = sql & "As " + """" + "Value_" + """" + " From " + """" + a3 + "DATA" + """" + " A Left Join " + """" + a3 + "SUBTYPES" + """" + " B On A." + """" + "ID_TYPE" + """" + " = B." + """" + "ID_TYPE" + """" + " and A." + """" + "ID_SUBTYPE" + """" + " = B." + """" + "ID_SUBTYPE" + """" + " Where " + """" + "OBJECT_ID_" + """" + "='" + za67 + "' and B." + """" + "ID_TYPE" + """" + " = '" + za66 + "' Union (Select " + za67 + ","
sql = sql & """" + "SELECTION_" + """" + "," + """" + "MAX_" + """" + "," + """" + "MIN_" + """" + "," + """" + "DATATYPE" + """" + "," + "A." + """" + "DESCRIPTION" + """" + ",A." + """" + "ID_TYPE" + """" + ",A." + """" + "ID_SUBTYPE" + """" + ",CAST(A." + """" + "DEFAULTVALUE" + """" + " AS Integer) as " + """" + "Value_Ref" + """" + ", B." + """" + "OPTION_" + """"
sql = sql & " as " + """" + "Value_" + """" + " From " + """" + a3 + "SUBTYPES" + """" + " A Left Join " + """" + a3 + "SELECTIONS" + """" + " B On A." + """" + "ID_TYPE" + """" + " = B." + """" + "ID_TYPE" + """" + " and A." + """" + "ID_SUBTYPE" + """" + " = B." + """" + "ID_SUBTYPE" + """" + " and B." + """" + "VALUE_" + """" + " = CAST(A." + """" + "DEFAULTVALUE" + """" + " AS INTEGER)"
sql = sql & "Where A." + """" + "ID_TYPE" + """" + " = '0' And A." + """" + "ID_SUB"
sql = sql & "TYPE" + """" + " not in(Select " + """" + "ID_SUBTYPE" + """" + " from " + """" + a3 + "DATA" + """" + " where " + """" + "OBJECT_ID_" + """" + " = '" + object_id_ + "'))"


Else


sql = sql & "Select " + a + "," + """" + "SELECTION_" + """" + "," + """" + "MAX_" + """" + "," + """" + "MIN_" + """" + "," + """" + "DATATYPE" + """" + "," + "B." + """" + "DESCRIPTION" + """" + "," + "A." + """" + "ID_TYPE" + """" + "," + "B." + """" + "ID_SUBTYPE" + """" + "," + "A." + """" + "VALUE_" + """" + " as " + """" + "Value_Ref" + """" + ","
sql = sql & "(Select  " + """" + "OPTION_" + """" + " From " + """" + a3 + "SELECTIONS" + """" + " C where C." + """" + "ID_TYPE" + """" + " = A." + """" + "ID_TYPE" + """" + " and C." + """" + "ID_SUBTYPE" + """" + "= A." + """" + "ID_SUBTYPE" + """" + " and C." + """" + "VALUE_" + """" + "=CAST(A." + """" + "VALUE_" + """" + "As Integer))"
sql = sql & "As " + """" + "Value_" + """" + " From " + """" + a3 + "DATA" + """" + " A Left Join " + """" + a3 + "SUBTYPES" + """" + " B On A." + """" + "ID_TYPE" + """" + " = B." + """" + "ID_TYPE" + """" + " and A." + """" + "ID_SUBTYPE" + """" + " = B." + """" + "ID_SUBTYPE" + """" + " Where " + """" + "OBJECT_ID_" + """" + "='" + za67 + "' and B." + """" + "ID_TYPE" + """" + " = '" + za66 + "' Union (Select " + za67 + ","
sql = sql & """" + "SELECTION_" + """" + "," + """" + "MAX_" + """" + "," + """" + "MIN_" + """" + "," + """" + "DATATYPE" + """" + "," + "A." + """" + "DESCRIPTION" + """" + ",A." + """" + "ID_TYPE" + """" + ",A." + """" + "ID_SUBTYPE" + """" + ",A." + """" + "DEFAULTVALUE" + """" + " as " + """" + "Value_Ref" + """" + ", B." + """" + "OPTION_" + """"
sql = sql & " as " + """" + "Value_" + """" + " From " + """" + a3 + "SUBTYPES" + """" + " A Left Join " + """" + a3 + "SELECTIONS" + """" + " B On A." + """" + "ID_TYPE" + """" + " = B." + """" + "ID_TYPE" + """" + " and A." + """" + "ID_SUBTYPE" + """" + " = B." + """" + "ID_SUBTYPE" + """" + " and B." + """" + "VALUE_" + """" + " = CAST(A." + """" + "DEFAULTVALUE" + """" + " AS INTEGER)"
sql = sql & "Where A." + """" + "ID_TYPE" + """" + " = '0' And A." + """" + "ID_SUB"
sql = sql & "TYPE" + """" + " not in(Select " + """" + "ID_SUBTYPE" + """" + " from " + """" + a3 + "DATA" + """" + " where " + """" + "OBJECT_ID_" + """" + " = '" + object_id_ + "'))"



End If



            
'm 'sgBox "ARQUIVO DEBUG SALVO"
' WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"




           ' sql = sql & " (Select " + a + "," + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + "," + """" + e + """" + "," + """" + a3 + d3 + """" + "." + """" + e3 + """" + ",A" + "." + """" + a1 + """" + "," + """" + a3 + d3 + """" + "." + """" + b1 + """" + " ,A" + "." + """" + c2 + """" + " as " + """" + "Value_Ref" + """" + ","
          '  sql = sql & " Case " + """" + d2 + """" + " when '0' then CAST(A" + "." + """" + c2 + """" + "AS BOOLEAN)"
           ' sql = sql & " Else (Select CAST(" + """" + c2 + """" + " AS BOOLEAN)= CAST(" + """" + e6 + """" + " AS BOOLEAN)  From " + """" + c3 + """"
          '  sql = sql & " where " + """" + c3 + """" + "." + """" + a1 + """" + " = A" + "." + """" + a1 + """" + " and " + """" + a3 + b3 + """" + "." + """" + b1 + """" + "= A" + "." + """" + b1 + """" + " and " + """" + a3 + b3 + """" + "." + """" + c2 + """" + "=A" + "." + """" + c2 + """" + ") end As" + """" + "Value_" + """" + ""
          '  sql = sql & " From " + """" + b4 + """" + ""
          '  sql = sql & " Left Join " + """" + d4 + """" + ""
          '  sql = sql & " On " + """" + b4 + """" + "." + """" + a1 + """" + " = " + """" + d4 + """" + "." + """" + a1 + """" + " and " + """" + b4 + """" + "." + """" + b1 + """" + " = " + """" + d4 + """" + "." + """" + a1 + """" + ""
          '  sql = sql & " Where " + """" + "OBJECT_ID_" + """" + "='" + a + "' and " + """" + d4 + """" + "." + """" + a1 + """" + " ='" + za66 + "')"
          '  sql = sql & " Union"
          '  sql = sql & " (Select " + """" + object_id_ + """" + "," + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + "," + """" + e + """" + "," + """" + b4 + """" + "." + """" + a3 + """" + "," + """" + b4 + """" + "." + """" + a1 + """" + "," + """" + b4 + """" + "." + _
            """" + b1 + """" + "," + """" + b4 + """" + "." + """" + e4 + """" + " as " + """" + "Value_Ref" + """" + ","
          '  sql = sql & " case " + """" + d2 + """" + " When '1' then " + """" + d4 + """" + "." + """" + e6 + """" + " else " + """" + b4 + """" + "." + e4 + " End" + " as " + """" + "Value_" + """" + ""
         ' '  sql = sql & " From " + """" + d4 + """" + "Left Join " + """" + c3 + """" + ""
        '    sql = sql & " On " + """" + b4 + """" + "." + """" + a1 + """" + " = " + """" + d4 + """" + "." + """" + a1 + """" + " and " + """" + b4 + """" + "." + """" + b1 + """" + " = " + """" + d4 + """" + "." + """" + b1 + """" + " and " + """" + d4 + """" + "." + """" + c2 + """" + " = " + """" + b4 + """" + "." + """" + e4 + """" + ""
         '   sql = sql & " Where " + """" + b4 + """" + "." + """" + a1 + """" + " = '" & id_Type & "' And " + """" + b4 + """" + "." + """" + b1 + """" + " not in(Select " + """" + b1 + """" + " from " + """" + b4 + """" + " where " + """" + "OBJECT_ID_" + """" + " = '" & object_id_ & "'))"
            
            End If
            
            
'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
            
        
    Else

    
e6 = "OPTION_"
a = object_id_
b = "SELECTION_"
c = "MAX_"
d = "MIN_"
e = "DATATYPE"
a1 = "ID_TYPE"
b1 = "ID_SUBTYPE"
c2 = "VALUE_"
d2 = "SELECTION_"
e2 = "SUB_TYPES"
a3 = layerName
b3 = "SELECTIONS"
c3 = a3 + b3
d3 = "SUBTYPES"
e3 = "DESCRIPTION"
a4 = "DATA"
b4 = a3 + a4
c4 = "SUBTYPES"
d4 = a3 + c4
e4 = "DEFAULTVALUE"
    If tipoConex <> 4 Then
        sql = sql & " Select '" & object_id_ & "',Selection_,Max_,Min_,DataType,A.Description_,A.Id_Type,A.Id_subType,A.DefaultValue as " + """" + "Value_Ref" + """" + ", case Selection_ When 1 then b.Option_ else A.DefaultValue End as " + """" + "Value_" + """" + ""
        sql = sql & " From " & layerName & "SubTypes A Left Join " & layerName & "Selections B"
        sql = sql & " On A.Id_Type = B.Id_Type and A.Id_subType = B.Id_subType and B.Value_ = A.DefaultValue "
        sql = sql & " Where A.Id_Type = " & id_Type & " ORDER BY A.Id_subType"
        
           '    MsgBox "ARQUIVO DEBUG SALVO"
' WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
        
        Else
        
        
          sql = sql + " Select " + "'" + object_id_ + "'" + "," + """" + "SELECTION_" + """" + "," + """" + "MAX_" + """" + "," + """" + "MIN_" + """" + "," + """" + "DATATYPE" + """" + "," + """" + layerName + "SUBTYPES" + """" + "." + """" + "DESCRIPTION" + """" + "," + """" + layerName + "SUBTYPES" + """" + "." + """" + "ID_TYPE" + """" + "," + """" + layerName + "SUBTYPES" + """" + "." + """" + "ID_SUBTYPE" + """" + "," + """" + layerName + "SUBTYPES" + """" + "." + """" + "DEFAULTVALUE" + """" + " as" + """" + "Value_Ref" + """" + ", case " + """" + "SELECTION_" + """" + " When '1' then " + """" + layerName + "SELECTIONS" + """" + "." + """" + "OPTION_" + """" + " else " + """" + layerName + "SUBTYPES" + """" + "." + """" + "DEFAULTVALUE" + """" + " End as " + """" + "Value_" + """" + ""
        sql = sql & " From " + """" + layerName + "SUBTYPES" + """" + " Left Join " + """" + layerName + "SELECTIONS" + """" + ""
        sql = sql & " On " + """" + layerName + "SUBTYPES" + """" + "." + """" + "ID_TYPE" + """" + " = " + """" + layerName + "SELECTIONS" + """" + "." + """" + "ID_TYPE" + """" + " and " + """" + layerName + "SUBTYPES" + """" + "." + """" + "ID_SUBTYPE" + """" + " = " + """" + layerName + "SELECTIONS" + """" + "." + """" + "ID_SUBTYPE" + """" + " and " + """" + layerName + "SELECTIONS" + """" + "." + """" + "VALUE_" + """" + " = " + "CAST(" + """" + layerName + "SUBTYPES" + """" + "." + """" + "DEFAULTVALUE" + """" + " AS INTEGER)" + ""
        sql = sql & " Where " + """" + layerName + "SUBTYPES" + """" + "." + """" + "ID_TYPE" + """" + " = '" & id_Type & "' ORDER BY " + """" + layerName + "SUBTYPES" + """" + "." + """" + "ID_SUBTYPE" + """" + ""
       
       
     '  MsgBox "ARQUIVO DEBUG SALVO"
' WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
       
       ' Else
    '   SELECT '0',"SELECTION_","MAX_","MIN_","DATATYPE","WATERLINESSUBTYPES"."DESCRIPTION_","WATERLINESSUBTYPES"."ID_SUBTYPE","WATERLINESSUBTYPES"."DEFAULTVALUE" AS
 '"VALUE_REF",CASE "SELECTION_" WHEN '1' THEN "WATERLINESSELECTIONS"."OPTION_" ELSE "WATERLINESSUBTYPES"."DEFAULTVALUE" END AS "VALUE_" FROM
 '"WATERLINESSUBTYPES" LEFT JOIN "WATERLINESSELECTIONS" ON "WATERLINESSUBTYPES"."ID_TYPE"="WATERLINESSELECTIONS"."ID_TYPE" AND "WATERLINESSUBTYPES"."ID_SUBTYPE"="WATERLINESSELECTIONS"."ID_SUBTYPE"
 ' AND "WATERLINESSELECTIONS"."VALUE_"=cast("WATERLINESSUBTYPES"."DEFAULTVALUE" AS INTEGER) WHERE "WATERLINESSUBTYPES"."ID_TYPE"='0' ORDER BY "WATERLINESSUBTYPES"."ID_SUBTYPE";
        End If
    End If
   ' MsgBox "Arquivo Gravado"
   ' WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
    'Open App.Path & "\GeoSanLog.txt" For Append As #2
    'Print #2, sql
    'Close #2
    
End If
' MsgBox "Arquivo Gravado"
   ' WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
rs.Close
getPmssp = sql
' MsgBox "Arquivo Gravado"
 '   WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"


''************* MONITORAMENTO ***************
'Close #2
'Open App.Path & "\GeoSanLog.txt" For Append As #2
'Print #2, Now & "Public Function getPmssp - SQL = " & sql
'Close #2
''***************** FIM *********************

Exit Function

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function getPmssp - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function

Public Function getPmSpo(layerName As String, id_Type As Integer, id_SubType As Integer) As String
On Error GoTo Trata_Erro

    Dim str As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String

a = "SELECTIONS"
b = "ID_TYPE"
c = "SUBTYPES"
d = "ID_TYPE"
e = "ID_SUBTYPE"

If tipoConex <> 4 Then
    str = "Select * From " & layerName & "Selections Where Id_Type="
    str = str & id_Type & " and Id_SubType=" & id_SubType
    Else
      str = "Select * From " + """" + layerName & a + """" + " Where " + """" + b + """" + "='"
    str = str & id_Type & "' and " + """" + e + """" + "='" & id_SubType & "'"
    
    End If
    
    getPmSpo = str

'    '************* MONITORAMENTO ***************
'    Close #2
'    Open App.Path & "\GeoSanLog.txt" For Append As #2
'    Print #2, Now & " Function getPmSpo str = " & str
'    Close #2
'    '***************** FIM *********************

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function getPmSpo - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function



Public Function FunDecripta(ByVal strDecripta As String) As String


    Dim IntTam As Integer
    Dim i As Integer
    Dim letra As String
    IntTam = Len(strDecripta)
    Dim nStr As String
    nStr = ""

    'desconsidera os os numeros de HH-MM-SS
    strDecripta = Mid(strDecripta, 6, 5) & Mid(strDecripta, 16, 5) & Mid(strDecripta, 26, 5) & _
                  Mid(strDecripta, 36, 5) & Mid(strDecripta, 46, 5) & Mid(strDecripta, 56, 200)

    i = 1
    Do While Not i = IntTam - 29
        letra = Mid(strDecripta, i, 5)
        Select Case letra
        Case "14334"
            nStr = nStr & "a"
        Case "14212"
            nStr = nStr & "A"
        Case "24334"
            nStr = nStr & "á"
        Case "24134"
            nStr = nStr & "â"
        Case "24234"
            nStr = nStr & "ã"
        Case "24314"
            nStr = nStr & "à"
        Case "24324"
            nStr = nStr & "b"
        Case "14223"
            nStr = nStr & "B"
        Case "11211"
            nStr = nStr & "ç"
        Case "11311"
            nStr = nStr & "Ç"
        Case "13334"
            nStr = nStr & "c"
        Case "14324"
            nStr = nStr & "C"
        Case "24344"
            nStr = nStr & "d"
        Case "14444"
            nStr = nStr & "D"
        Case "12314"
            nStr = nStr & "e"
        Case "21111"
            nStr = nStr & "E"
        Case "24321"
            nStr = nStr & "é"
        Case "32314"
            nStr = nStr & "ê"
        Case "31314"
            nStr = nStr & "f"
        Case "21311"
            nStr = nStr & "F"
        Case "32134"
            nStr = nStr & "g"
        Case "21341"
            nStr = nStr & "G"
        Case "31324"
            nStr = nStr & "h"
        Case "22111"
            nStr = nStr & "H"
        Case "32124"
            nStr = nStr & "i"
        Case "21112"
            nStr = nStr & "I"
        Case "31334"
            nStr = nStr & "í"
        Case "32333"
            nStr = nStr & "ì"
        Case "11314"
            nStr = nStr & "j"
        Case "23122"
            nStr = nStr & "J"
        Case "33134"
            nStr = nStr & "k"
        Case "23411"
            nStr = nStr & "K"
        Case "33314"
            nStr = nStr & "l"
       Case "32222"
            nStr = nStr & "L"
        Case "43423"
            nStr = nStr & "m"
        Case "32111"
            nStr = nStr & "M"
        Case "42423"
            nStr = nStr & "n"
        Case "33221"
            nStr = nStr & "N"
        Case "43234"
            nStr = nStr & "o"
        Case "33233"
            nStr = nStr & "O"
        Case "42444"
            nStr = nStr & "ô"
        Case "43223"
            nStr = nStr & "õ"
        Case "42433"
            nStr = nStr & "ò"
        Case "43231"
            nStr = nStr & "ó"
        Case "22223"
            nStr = nStr & "p"
        Case "33444"
            nStr = nStr & "P"
        Case "43233"
            nStr = nStr & "q"
        Case "34442"
            nStr = nStr & "Q"
        Case "43421"
            nStr = nStr & "r"
        Case "34332"
            nStr = nStr & "R"
        Case "13443"
            nStr = nStr & "s"
        Case "34222"
            nStr = nStr & "S"
        Case "44444"
            nStr = nStr & "t"
        Case "34112"
            nStr = nStr & "T"
        Case "13444"
            nStr = nStr & "u"
        Case "41311"
            nStr = nStr & "U"
        Case "11111"
            nStr = nStr & "ú"
        Case "13243"
            nStr = nStr & "ù"
        Case "11115"
            nStr = nStr & "û"
        Case "13241"
           nStr = nStr & "v"
        Case "41222"
            nStr = nStr & "V"
        Case "12443"
            nStr = nStr & "x"
        Case "41133"
            nStr = nStr & "X"
        Case "13244"
            nStr = nStr & "y"
        Case "42231"
            nStr = nStr & "Y"
        Case "13441"
            nStr = nStr & "w"
        Case "42222"
            nStr = nStr & "W"
        Case "11313"
            nStr = nStr & "z"
        Case "42213"
            nStr = nStr & "Z"
        Case "11312"
            nStr = nStr & "@"
        Case "11114"
            nStr = nStr & "%"
        Case "12341"
            nStr = nStr & "&"
        Case "13343"
            nStr = nStr & "*"
        Case "12342"
            nStr = nStr & "("
        Case "13344"
            nStr = nStr & ")"
        Case "12333"
            nStr = nStr & "$"
        Case "23334"
            nStr = nStr & "!"
        Case "13331"
            nStr = nStr & "#"
        Case "21242"
            nStr = nStr & "?"
        Case "22313"
            nStr = nStr & "1"
        Case "23424"
            nStr = nStr & "2"
        Case "24131"
            nStr = nStr & "3"
        Case "41414"
            nStr = nStr & "4"
        Case "22314"
           nStr = nStr & "5"
        Case "23423"
            nStr = nStr & "6"
        Case "44134"
            nStr = nStr & "7"
        Case "21241"
            nStr = nStr & "8"
       Case "22312"
           nStr = nStr & "9"
       Case "23231"
            nStr = nStr & "0"
        Case "34123"
            nStr = nStr & " "
        Case "14121"
            nStr = nStr & "_"
        Case "14144"
            nStr = nStr & "/"
        Case "12131"
            nStr = nStr & "\"
        Case "12124"
            nStr = nStr & "-"
        Case "21421"
            nStr = nStr & ";"
        Case "21321"
            nStr = nStr & ":"
        Case "14431"
            nStr = nStr & ","
        Case "13421"
            nStr = nStr & "."
        Case "11213"
            nStr = nStr & "+"
        Case "11212"
            nStr = nStr & "="

        Case Else
            MsgBox "Código de criptografia inválido!"
            'mStrDeCriptografa = ""
            Exit Function
        End Select
        i = i + 5
    Loop
  FunDecripta = nStr
    'mStrDeCriptografa = nStr

Exit Function
End Function








Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
  
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 
 Dim retlen As String
 Dim Ret As String
 
 Ret = String$(255, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
 Ret = Left$(Ret, retlen)
 ReadINI = Ret

End Function

Public Function getPMIAS(layerName As String, object_id_ As String, id_Type As Integer, id_SubType As Integer, Value_ As String, conn As ADODB.Connection) As String
    
On Error GoTo Trata_Erro

    Dim str As String, encontrou As Integer
    Dim rs As New ADODB.Recordset
    Dim gg As String
    Dim ga As String
     Dim ga1 As String
     Dim ga2 As String
     Dim ga3 As String
    Dim count3 As Integer
    If tipoConex = 4 And count3 <> 10 Then
    Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
Dim conexao As New ADODB.Connection
Dim strConn As String
Dim nStr As String
mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.Path & "\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.Path & "\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.Path & "\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.Path & "\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.Path & "\GEOSAN.ini")
nStr = FunDecripta(Senha)
decriptada = nStr
  strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"

    conexao.Open strConn
    count3 = 10
    End If
    
    
    
    
    
    
    
    
      gg = "DATA"
    ga = "OBJECT_ID_"
    ga1 = "ID_TYPE"
      ga2 = "VALUE_"
        ga3 = "ID_SUBTYPE"
If tipoConex <> 4 Then
    str = "Delete From " & layerName & "Data Where Object_id_= " & object_id_ & " and  Id_Type <> " & id_Type
  '  MsgBox str
    conn.Execute (str)
    str = "Select Count(Value_) From " & layerName & "Data"
    str = str & " Where Object_id_= " & object_id_ & " and Id_Type = " & id_Type & " and id_SubType = " & id_SubType
   ' MsgBox str
Else
Dim za2 As String
za2 = layerName + "DATA"
Dim za3, za4 As String
za3 = object_id_
za4 = id_Type
Dim za5, za6, za7 As String
za5 = object_id_
za6 = id_Type
za7 = id_SubType

str = "Delete From " + """" + za2 + """" + " Where " + """" + "OBJECT_ID_" + """" + " =  '" + za3 + "'" + " and " + """" + "ID_TYPE" + """" + " <> '" + za4 + "'"
    
    conexao.Execute (str)
    str = "Select Count(" + """" + ga2 + """" + ") From " + """" + layerName + gg + """"
    str = str & " Where " + """" + ga + """" + "='" + za5 + "' and " + """" + ga1 + """" + " ='" & za6 & "'and" + """" + ga3 + """" + " = '" + za7 + "'"


End If
 If tipoConex <> 4 Then
    rs.Open str, conn, adOpenDynamic, adLockOptimistic
Else
rs.Open str, conexao, adOpenDynamic, adLockOptimistic

End If
    If rs.Fields(0).Value > 0 Then
    If tipoConex = 1 Then
        str = "Update " & layerName & "Data Set Value_ = '" & Value_
        str = str & "' Where Object_id_=" & object_id_ & " and Id_Type=" & id_Type & " and id_SubType=" & id_SubType
        
        
'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", str, App.Path & "\DEBUG.INI"
        
        
           conn.Execute (str)
           
           
            ElseIf tipoConex = 2 Then
        str = "Update " & layerName & "Data Set Value_ = " & Value_
        str = str & " Where Object_id_=" & object_id_ & " and Id_Type=" & id_Type & " and id_SubType=" & id_SubType
         conn.Execute (str)
    Else
    Dim d As String
    Dim e As String
    Dim g As String
    Dim j As String
    Dim k As String
    Dim m As String
    Dim f As String
    Dim n As String
    d = layerName
    e = "DATA"
    f = d + e
    k = "OBJECT_ID_"
    m = "ID_TYPE"
    n = "ID_SUBTYPE"
    g = "VALUE_"
    
    str = "Update" & """" + layerName + "DATA" + """" + "Set" + """" + g + """" + "='" & Value_ & "'"
        str = str & " Where " + """" + k + """" + " ='" & object_id_ & "' and " + """" + m + """" + " ='" & id_Type & "' and" + """" + n + """" + " ='" & id_SubType & "'"
   conn.Execute (str)
    End If
    Else
    If tipoConex <> 4 Then
    
        str = "Insert Into " & layerName & "Data (Object_id_, Id_Type, id_SubType,Value_)"
        str = str & "Values (" & object_id_ & ", " & id_Type & ", " & id_SubType & "," & Value_ & ")"
        conn.Execute (str)
        Else
      Dim za9 As String
      za9 = Value_
      Dim za10 As String
      za10 = za7
        str = "Insert Into " + """" + layerName + "DATA" + """" + " (" + """" + "OBJECT_ID_" + """" + "," + """" + "ID_TYPE" + """" + "," + """" + "ID_SUBTYPE" + """" + "," + """" + "VALUE_" + """" + ")"
        str = str & "Values ('" & za5 & "', '" & za4 & "', '" & za10 & "','" & za9 & "')"
        conexao.Execute (str)
    End If
    
    End If
    
    
'    '************* MONITORAMENTO ***************
'    Close #2
'    Open App.Path & "\GeoSanLog.txt" For Append As #2
'    Print #2, Now & " Function getPMIAS str = " & str
'    Close #2
'    '***************** FIM *********************

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function getPMIAS - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
    
End Function

Public Function IAD(layerName As String, str As String, conn As ADODB.Connection) As String
On Error GoTo Trata_Erro

Dim encontrou As Integer, rs As New ADODB.Recordset
Dim sql As String, lista() As String
lista = Split(str, ",")
 Dim tas, tas2 As String
            
If tipoConex <> 4 Then
Select Case UCase(layerName)
    Case "DRAINCOMPONENTS"
        sql = "Select count(Component_id) From DrainComponents Where Component_id=" & lista(0)
        rs.Open sql, conn
        sql = getSewerDrainComponents(str, rs.Fields(0).Value, False)
    Case "DRAINLINES"
        sql = "Select count(Line_id) From Drainlines Where Line_id=" & lista(0)
        rs.Open sql, conn
        sql = getSewerDrainLines(str, rs.Fields(0).Value, False)
    Case "SEWERCOMPONENTS"
        sql = "Select count(Component_id) From SewerComponents Where Component_id=" & lista(0)
        rs.Open sql, conn
        sql = getSewerDrainComponents(str, rs.Fields(0).Value, True)
    Case "SEWERLINES"
        sql = "Select count(Line_id) From Sewerlines Where Line_id=" & lista(0)
        rs.Open sql, conn
        sql = getSewerDrainLines(str, rs.Fields(0).Value, True)
    Case "WATERCOMPONENTS"
        sql = "Select count(Component_id) From WaterComponents Where component_id =" & lista(0)
        rs.Open sql, conn
        sql = getWaterComponents(str, rs.Fields(0).Value)
    Case "WATERLINES"
        sql = "Select count(Line_id) From WaterLines Where Line_id=" & lista(0)
        rs.Open sql, conn
        sql = getWaterLines(str, rs.Fields(0).Value)
End Select
Else

Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String
Dim k As String
Dim l As String


a = "COMPONENT_ID"
b = "DRAINCOMPONENTS"
c = "DRAINLINES"
d = "SEWERCOMPONENTS"
e = "SEWERLINES"
f = "WATERCOMPONENTS"
g = "WATERLINES"
h = "LINE_ID"


             
Select Case UCase(layerName)
    Case "DRAINCOMPONENTS"
     tas = lista(0)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = "Select count(" + """" + a + """" + ") From " + """" + b + """" + " Where " + """" + a + """" + "='" & tas2 & "'"
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic
        sql = getSewerDrainComponents(str, rs.Fields(0).Value, False)
    Case "DRAINLINES"
     tas = lista(0)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = "Select count(" + """" + h + """" + ") From " + """" + c + """" + " Where " + """" + h + """" + "='" & tas2 & "'"
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic
        sql = getSewerDrainLines(str, rs.Fields(0).Value, False)
    Case "SEWERCOMPONENTS"
     tas = lista(0)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = "Select count(" + """" + a + """" + ") From " + """" + d + """" + " Where " + """" + a + """" + "='" & tas2 & "'"
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic
        sql = getSewerDrainComponents(str, rs.Fields(0).Value, True)
    Case "SEWERLINES"
     tas = lista(0)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = "Select count(" + """" + h + """" + ") From " + """" + e + """" + " Where " + """" + h + """" + "='" & tas2 & "'"
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic
        sql = getSewerDrainLines(str, rs.Fields(0).Value, True)
    Case "WATERCOMPONENTS"
     tas = lista(0)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = "Select count(" + """" + a + """" + ") From " + """" + f + """" + " Where " + """" + a + """" + " ='" & tas2 & "'"
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic
        sql = getWaterComponents(str, rs.Fields(0).Value)
    Case "WATERLINES"
     tas = lista(0)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = "Select count(" + """" + h + """" + ") From " + """" + g + """" + " Where " + """" + h + """" + "='" & tas2 & "'"
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic
        sql = getWaterLines(str, rs.Fields(0).Value)
End Select

End If
IAD = sql

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function IAD - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function

Public Function getSewerDrainLines(str As String, rs As Integer, ehSewer As Boolean) As String
On Error GoTo Trata_Erro
    
   Dim sql As String, lista() As String
   
      
   If rs > 0 Then
       If ehSewer = True Then
       If tipoConex <> 4 Then
           sql = "Update SewerLines Set"
           Else
    aa = "SEWERLINES"
   bb = "DRAINLINES"
           sql = "Update " + """" + aa + """" + " Set"
           End If
       Else
       If tipoConex <> 4 Then
           sql = "Update DrainLines Set"
           Else
    aa = "SEWERLINES"
   bb = "DRAINLINES"
            sql = "Update " + """" + bb + """" + " Set"
       End If
       End If
        If tipoConex <> 4 Then
           lista = Split(str, ",")
       sql = sql & " Id_Type=" & lista(1)
       
       sql = sql & ", InitialGroundHeight=" & lista(2)
       sql = sql & ", FinalGroundHeight=" & lista(3)
       sql = sql & ", InitialTubeDeepness=" & lista(4)
       sql = sql & ", FinalTubeDeepness=" & lista(5)
       
       sql = sql & ", InternalDiameter=" & lista(6)
       sql = sql & ", ExternalDiameter=" & lista(7)
       sql = sql & ", InitialComponent=" & lista(8)
       sql = sql & ", FinalComponent=" & lista(9)
       sql = sql & ", Thickness=" & lista(10)
       sql = sql & ", Material=" & lista(11)
       sql = sql & ", Length=" & lista(12)
       sql = sql & ", LengthCalculated=" & lista(13)
       sql = sql & ", Supplier=" & lista(14)
       sql = sql & ", Manufacturer=" & lista(15)
       sql = sql & ", Location=" & lista(16)
       
'       sql = sql & ", State=" & lista(17)
'       sql = sql & ", Sector=" & lista(18)
'       sql = sql & ", InformationValidity=" & lista(19)
'       sql = sql & ", DateInstallation=" & lista(20)
'       sql = sql & ", SideStreet=" & lista(21)
'       sql = sql & ", DividedDistance=" & lista(22)
'       sql = sql & " Where object_id_=" & lista(0)
            
            'copiado do getSewerWaterLines
            
         sql = sql & ", State=" & lista(17)
         sql = sql & ", RoughNess=" & lista(18)
         sql = sql & ", Sector=" & lista(19)
         sql = sql & ", InformationValidity=" & lista(20)
         sql = sql & ", DateInstallation=" & lista(21)
         sql = sql & ", SideStreet=" & lista(22)
         sql = sql & ", DividedDistance=" & lista(23)
         
         sql = sql & ", USUARIO_LOG=" & lista(24) '***** INCLUIDO EM 14/07/09 JONATHAS
         sql = sql & ", DATA_LOG=" & lista(25)    '***** INCLUIDO EM 14/07/09 JONATHAS
         
         sql = sql & " Where object_id_=" & lista(0)
         
         
         Else
          sa = "INITIALGROUNDHEIGHT"
   sb = "FINALGROUNDHEIGHT"
   sc = "INITIALTUBEDEEPNESS"
   sd = "FINALTUBEDEEPNESS"
   se = "INTERNALDIAMETER"
   sf = "EXTERNALDIAMETER"
   sg = "INITIALCOMPONENT"
   sh = "FINALCOMPONENT"
   si = "THICKNESS"
   sj = "MATERIAL"
   sl = "LENGTH"
   sm = "LENGTHCALCULATED"
   sn = "SUPPLIER"
   so = "MANUFACTURER"
   sp = "LOCATION"
   sq = "STATE"
   sr = "ROUGHNESS"
   ss = "SECTOR"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   sv = "SIDESTREET"
   sx = "DIVIDEDDISTANCE"
   sz = "USUARIO_LOG"
   sk = "DATA_LOG"
   sw = "OBJECT_ID_"
   aa = "SEWERLINES"
   bb = "DRAINLINES"
   swx = "ID_TYPE"
   sss = "YEAROFCONSTRUCTION"
   ssv = "NOTES"
   ssz = "TROUBLE"
   ssr = "PATTERN"
   ssa = "COMPONENT_ID"
   sst = "WATERLINES"
   ssq = "WATERCOMPONENTS"
   sse = "DEMAND"
   ssj = "CALCULENODE"
   ssd = "LINE_ID"
      lista = Split(str, ",")
         
         
         Dim tas, tas2 As String
             tas = lista(1)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
         sql = sql & " " + """" + swx + """" + "='" & tas2 & "'"
           
             tas = lista(2)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       
       sql = sql & ", " + """" + sa + """" + "='" & tas2 & "'"
           
             tas = lista(3)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sb + """" + "= '" & tas2 & "'"
           
             tas = lista(4)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sc + """" + "= '" & tas2 & "'"
           
             tas = lista(5)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sd + """" + "='" & tas2 & "'"
           
             tas = lista(6)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       
       sql = sql & ", " + """" + se + """" + "='" & tas2 & "'"
           
             tas = lista(7)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sf + """" + "='" & tas2 & "'"
           
             tas = lista(8)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sg + """" + "='" & tas2 & "'"
           
             tas = lista(9)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sh + """" + "='" & tas2 & "'"
           
             tas = lista(10)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + si + """" + "='" & tas2 & "'"
           
             tas = lista(11)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sj + """" + "='" & tas2 & "'"
           
             tas = lista(12)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sl + """" + "='" & tas2 & "'"
           
             tas = lista(13)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sm + """" + "='" & tas2 & "'"
           
             tas = lista(14)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sn + """" + "='" & tas2 & "'"
           
             tas = lista(15)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + so + """" + "='" & tas2 & "'"
           
             tas = lista(16)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", " + """" + sp + """" + "='" & tas2 & "'"
           
             tas = lista(17)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       
'       sql = sql & ", State=" & lista(17)
'       sql = sql & ", Sector=" & lista(18)
'       sql = sql & ", InformationValidity=" & lista(19)
'       sql = sql & ", DateInstallation=" & lista(20)
'       sql = sql & ", SideStreet=" & lista(21)
'       sql = sql & ", DividedDistance=" & lista(22)
'       sql = sql & " Where object_id_=" & lista(0)
            
            'copiado do getSewerWaterLines
            
         sql = sql & ", " + """" + sq + """" + "='" & tas2 & "'"
           
             tas = lista(18)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
         sql = sql & ", " + """" + sr + """" + "='" & tas2 & "'"
           
             tas = lista(19)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
         sql = sql & ", " + """" + ss + """" + "='" & tas2 & "'"
           
             tas = lista(20)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
         sql = sql & ", " + """" + st + """" + "='" & tas2 & "'"
           
             tas = lista(21)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
         sql = sql & ", " + """" + su + """" + "='" & tas2 & "'"
           
             tas = lista(22)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
         sql = sql & ", " + """" + sv + """" + "='" & tas2 & "'"
           
             tas = lista(23)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
         sql = sql & ", " + """" + sx + """" + "='" & tas2 & "'"
           
             tas = lista(24)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
         
         sql = sql & ", " + """" + sz + """" + "= '" & tas2 & "'"
           
             tas = lista(25)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If '***** INCLUIDO EM 14/07/09 JONATHAS"
         sql = sql & ", " + """" + sk + """" + "='" & tas2 & "'"
           
             tas = lista(0)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If '***** INCLUIDO EM 14/07/09 JONATHAS
         
         sql = sql & " Where " + """" + sw + """" + "='" & tas2 & "'"
      
   End If
  End If
  If rs = 0 Or rs < 0 Then
     If ehSewer = True Then
   If tipoConex <> 4 Then
      aa = "SEWERLINES"
   bb = "DRAINLINES"
           sql = "Insert Into SewerLines"
           
       Else
       sql = "Insert Into " + """" + aa + """" + ""
       End If
       
       Else
       If tipoConex <> 4 Then
        aa = "SEWERLINES"
   bb = "DRAINLINES"
           sql = "Insert Into DrainLines"
           Else
           sql = "Insert Into " + """" + bb + """" + ""
           End If
           
       End If
       If tipoConex <> 4 Then
          lista = Split(str, ",")
       sql = sql & " (object_id_, Id_Type, InitialGroundHeight, FinalGroundHeight, InitialTubeDeepness, FinalTubeDeepness, InternalDiameter, ExternalDiameter, InitialComponent, FinalComponent,"
       sql = sql & " Thickness, MATERIAL,  Length, LengthCalculated, Supplier, Manufacturer, Location, State, Sector, InformationValidity,DateInstallation,SideStreet,DividedDistance, USUARIO_LOG,DATA_LOG)"
       sql = sql & " Values (" & lista(1)
       sql = sql & ", " & lista(2)
       sql = sql & ", " & lista(3)
       sql = sql & ", " & lista(4)
       sql = sql & ", " & lista(5)
       sql = sql & ", " & lista(6)
       sql = sql & ", " & lista(7)
       sql = sql & ", " & lista(8)
       sql = sql & ", " & lista(9)
       sql = sql & ", " & lista(10)
       sql = sql & ", " & lista(11)
       sql = sql & ", " & lista(12)
       sql = sql & ", " & lista(13)
       sql = sql & ", " & lista(14)
       sql = sql & ", " & lista(15)
       sql = sql & ", " & lista(16)
       sql = sql & ", " & lista(17)
       sql = sql & ", " & lista(18)
       sql = sql & ", " & lista(19)
       sql = sql & ", " & lista(21)
       sql = sql & ", " & lista(22)
       sql = sql & ", " & lista(23)
       sql = sql & ", " & lista(24)
       sql = sql & ", " & lista(25)
       sql = sql & ")"
 
   
   
   

'    '************* MONITORAMENTO ***************
'    Close #2
'    Open App.Path & "\GeoSanLog.txt" For Append As #2
'    Print #2, Now & " Function getSewerDrainLines SQL = " & sql
'    Close #2
'    '***************** FIM *********************


   Else
   sa = "INITIALGROUNDHEIGHT"
   sb = "FINALGROUNDHEIGHT"
   sc = "INITIALTUBEDEEPNESS"
   sd = "FINALTUBEDEEPNESS"
   se = "INTERNALDIAMETER"
   sf = "EXTERNALDIAMETER"
   sg = "INITIALCOMPONENT"
   sh = "FINALCOMPONENT"
   si = "THICKNESS"
   sj = "MATERIAL"
   sl = "LENGTH"
   sm = "LENGTHCALCULATED"
   sn = "SUPPLIER"
   so = "MANUFACTURER"
   sp = "LOCATION"
   sq = "STATE"
   sr = "ROUGHNESS"
   ss = "SECTOR"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   sv = "SIDESTREET"
   sx = "DIVIDEDDISTANCE"
   sz = "USUARIO_LOG"
   sk = "DATA_LOG"
   sw = "OBJECT_ID_"
   aa = "SEWERLINES"
   bb = "DRAINLINES"
   swx = "ID_TYPE"
   sss = "YEAROFCONSTRUCTION"
   ssv = "NOTES"
   ssz = "TROUBLE"
   ssr = "PATTERN"
   ssa = "COMPONENT_ID"
   sst = "WATERLINES"
   ssq = "WATERCOMPONENTS"
   sse = "DEMAND"
   ssj = "CALCULENODE"
   ssd = "LINE_ID"
      lista = Split(str, ",")
   sql = sql & " (" + """" + sw + """" + ", " + """" + swx + """" + ", " + """" + sa + """" + ", " + """" + sb + """" + ", " + """" + sc + """" + ", " + """" + sd + """" + ", " + """" + se + """" + ", " + """" + sf + """" + ", " + """" + sg + """" + ", " + """" + sh + """" + ","
       sql = sql & " " + """" + si + """" + ", " + """" + sj + """" + ",  " + """" + sl + """" + ", " + """" + sm + """" + ", " + """" + sn + """" + ", " + """" + so + """" + ", " + """" + sp + """" + ", " + """" + sq + """" + ", " + """" + ss + """" + ", " + """" + st + """" + "," + """" + su + """" + "," + """" + sv + """" + "," + """" + sx + """" + ", " + """" + sz + """" + "," + """" + sk + """" + ")"
     MsgBox (sql)
             tas = lista(1)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
       sql = sql & " Values ('" & tas2 & "'"
           
             tas = lista(2)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(3)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(4)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(5)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(6)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(7)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(8)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(9)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(10)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(11)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(12)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(13)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(14)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(15)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(16)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(17)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(18)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(19)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(20)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(21)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(22)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(23)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(24)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
           
             tas = lista(25)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
       sql = sql & ", '" & tas2 & "'"
       sql = sql & ")"
   End If
 
   
   End If
   
  getSewerDrainLines = sql
'    '************* MONITORAMENTO ***************
'    Close #2
'    Open App.Path & "\GeoSanLog.txt" For Append As #2
'    Print #2, Now & " Function getSewerDrainLines SQL = " & sql
'    Close #2
'    '***************** FIM *********************

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
             
      'Open App.Path & "\GeoSanLog.txt" For Append As #1
      'Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function getSewerDrainLines - " & Err.Number & " - " & Err.Description
      'Close #1
      'MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function

Public Function getSewerDrainComponents(str As String, rs As Integer, ehSewer As Boolean) As String
On Error GoTo Trata_Erro
    Dim aa As String
    Dim bb As String
    aa = "SEWERCOMPONENTS"
    bb = "DRAINCOMPONENTS"
    Dim sql As String, lista() As String
   
   
      
    lista = Split(str, ",")
    If rs > 0 Then
        If ehSewer = True Then
         If tipoConex <> 4 Then
            sql = "Update SewerComponents Set"
            Else
            sql = "Update " + """" + aa + """" + " Set"
            End If
        Else
         If tipoConex <> 4 Then
            sql = "Update DrainComponents Set"
            Else
            sql = "Update " + """" + bb + """" + " Set"
            End If
        End If
         If tipoConex <> 4 Then
        sql = sql & " Id_Type=" & lista(1)
        sql = sql & ", YearOfConstruction=" & lista(2)
        sql = sql & ", State=" & lista(3)
        sql = sql & ", Location=" & lista(4)
        sql = sql & ", Supplier=" & lista(5)
        sql = sql & ", Manufacturer=" & lista(6)
        sql = sql & ", GroundHeight=" & lista(7)
        sql = sql & ", GroundHeightFinal=" & lista(8)
        sql = sql & ", InformationValidity=" & lista(9)
        sql = sql & ", Notes=" & lista(10)
        
        sql = sql & ", TROUBLE=" & lista(11) ' NÃO CONFORMIDADE
        sql = sql & ", DATEINSTALLATION=" & lista(12) 'DATA DE INSTALAÇÃO
        sql = sql & ", PATTERN=" & lista(13) 'PADRÃO CONSULMO
        sql = sql & ", SECTOR=" & lista(14) 'SETOR
        
        sql = sql & " Where Component_id=" & lista(0)
    Else
     
    sa = "INITIALGROUNDHEIGHT"
   sb = "FINALGROUNDHEIGHT"
   sc = "INITIALTUBEDEEPNESS"
   sd = "FINALTUBEDEEPNESS"
   se = "INTERNALDIAMETER"
   sf = "EXTERNALDIAMETER"
   sg = "INITIALCOMPONENT"
   sh = "FINALCOMPONENT"
   si = "THICKNESS"
   sj = "MATERIAL"
   sl = "LENGTH"
   sm = "LENGTHCALCULATED"
   sn = "SUPPLIER"
   so = "MANUFACTURER"
   sp = "LOCATION"
   sq = "STATE"
   sr = "ROUGHNESS"
   ss = "SECTOR"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   sv = "SIDESTREET"
   sx = "DIVIDEDDISTANCE"
   sz = "USUARIO_LOG"
   sk = "DATA_LOG"
   sw = "USUARIO_LOG"
   aa = "SEWERLINES"
   bb = "DRAINLINES"
   swx = "ID_TYPE"
   sss = "YEAROFCONSTRUCTION"
   ssv = "NOTES"
   ssz = "TROUBLE"
   ssr = "PATTERN"
   ssa = "COMPONENT_ID"
    
    Dim tas, tas2 As String
             tas = lista(1)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
    sql = sql & " " + """" + swx + """" + "='" & tas2 & "'"
           
             tas = lista(2)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", " + """" + sss + """" + "='" & tas2 & "'"
           
             tas = lista(3)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", " + """" + sq + """" + "='" & tas2 & "'"
           
             tas = lista(4)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", " + """" + sp + """" + "='" & tas2 & "'"
           
             tas = lista(5)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", " + """" + sn + """" + "='" & tas2 & "'"
           
             tas = lista(6)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", " + """" + so + """" + "='" & tas2 & "'"
           
             tas = lista(7)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", " + """" + sa + """" + "='" & tas2 & "'"
           
             tas = lista(8)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", " + """" + sb + """" + "='" & tas2 & "'"
           
             tas = lista(9)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", " + """" + st + """" + "='" & tas2 & "'"
           
             tas = lista(10)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", " + """" + ssv + """" + "='" & tas2 & "'"
           
             tas = lista(11)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        
        sql = sql & ", " + """" + ssz + """" + "='" & tas2 & "'"
           
             tas = lista(12)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If ' NÃO CONFORMIDADE
        sql = sql & ", " + """" + su + """" + "='" & tas2 & "'"
           
             tas = lista(13)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If 'DATA DE INSTALAÇÃO
        sql = sql & ", " + """" + ssr + """" + "='" & tas2 & "'"
           
             tas = lista(14)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If 'PADRÃO CONSULMO
        sql = sql & ", " + """" + ss + """" + "='" & tas2 & "'"
           
             tas = lista(0)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If 'SETOR
        
        sql = sql & " Where " + """" + ssa + """" + "='" & tas2 & "'"
        End If
        Else
    
        If ehSewer = True Then
         If tipoConex <> 4 Then
            sql = "Insert Into SewerComponents"
            Else
            sql = "Insert Into " + """" + aa + """" + ""
            End If
        Else
         If tipoConex <> 4 Then
            sql = "Insert Into DrainComponents"
            Else
            sql = "Insert Into " + """" + bb + """" + ""
            
            End If
        End If
        If tipoConex <> 4 Then
        sql = sql & " (object_id_, Id_Type, YearOfConstruction, State, Location, Manufacturer, GroundHeight, GroundHeightFinal, InformationValidity, Notes)"
        sql = sql & " Values (" & lista(1)
        sql = sql & ", " & lista(2)
        sql = sql & ", " & lista(3)
        sql = sql & ", " & lista(4)
        sql = sql & ", " & lista(5)
        sql = sql & ", " & lista(6)
        sql = sql & ", " & lista(7)
        sql = sql & ", " & lista(8)
        sql = sql & ", " & lista(9)
        sql = sql & ", " & lista(10)
        sql = sql & ")"
        Else
        
         
    sa = "INITIALGROUNDHEIGHT"
   sb = "FINALGROUNDHEIGHT"
   sc = "INITIALTUBEDEEPNESS"
   sd = "FINALTUBEDEEPNESS"
   se = "INTERNALDIAMETER"
   sf = "EXTERNALDIAMETER"
   sg = "INITIALCOMPONENT"
   sh = "FINALCOMPONENT"
   si = "THICKNESS"
   sj = "MATERIAL"
   sl = "LENGTH"
   sm = "LENGTHCALCULATED"
   sn = "SUPPLIER"
   so = "MANUFACTURER"
   sp = "LOCATION"
   sq = "STATE"
   sr = "ROUGHNESS"
   ss = "SECTOR"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   sv = "SIDESTREET"
   sx = "DIVIDEDDISTANCE"
   sz = "USUARIO_LOG"
   sk = "DATA_LOG"
   sw = "USUARIO_LOG"
   aa = "SEWERLINES"
   bb = "DRAINLINES"
   swx = "ID_TYPE"
   sss = "YEAROFCONSTRUCTION"
   ssv = "NOTES"
   ssz = "TROUBLE"
   ssr = "PATTERN"
   ssa = "COMPONENT_ID"
                
                sql = sql & " (" + """" + sw + """" + ", " + """" + swx + """" + ", " + """" + sss + """" + ", " + """" + sq + """" + ", " + """" + sp + """" + ", " + """" + so + """" + ", " + """" + sa + """" + ", " + """" + sb + """" + ", " + """" + st + """" + ", " + """" + ssv + """" + ")"
                
             tas = lista(1)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & " Values ('" & tas2 & "'"
           
             tas = lista(2)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", '" & tas2 & "'"
           
             tas = lista(3)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", '" & tas2 & "'"
           
             tas = lista(4)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", '" & tas2 & "'"
           
             tas = lista(5)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", '" & tas2 & "'"
           
             tas = lista(6)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", '" & tas2 & "'"
           
             tas = lista(7)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", '" & tas2 & "'"
           
             tas = lista(8)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", '" & tas2 & "'"
           
             tas = lista(9)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", '" & tas2 & "'"
           
             tas = lista(10)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
        sql = sql & ", '" & tas2 & "'"
        sql = sql & ")"
        
        End If
    End If
    getSewerDrainComponents = sql

'    '************* MONITORAMENTO ***************
    Close #2
    Open App.Path & "\GeoSanLog.txt" For Append As #2
    Print #2, sql
    Close #2
'    '***************** FIM *********************


Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then

      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function getSewerDrainComponents - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function

Public Function getWaterLines(str As String, rs As Integer) As String
On Error GoTo Trata_Erro
     
     Dim sql As String, lista() As String

           lista = Split(str, ",")
  
    
        If rs > 0 Then
        If tipoConex <> 4 Then
        
                   sql = "Update WaterLines Set"
            sql = sql & " Id_Type=" & lista(1)
            sql = sql & ", InitialGroundHeight=" & lista(2)
            sql = sql & ", FinalGroundHeight=" & lista(3)
            sql = sql & ", InitialTubeDeepness=" & lista(4)
            sql = sql & ", FinalTubeDeepness=" & lista(5)
            sql = sql & ", InternalDiameter=" & lista(6)
            sql = sql & ", ExternalDiameter=" & lista(7)
            sql = sql & ", InitialComponent=" & lista(8)
            sql = sql & ", FinalComponent=" & lista(9)
            sql = sql & ", Thickness=" & lista(10)
            sql = sql & ", Material=" & lista(11)
            sql = sql & ", Length=" & lista(12)
            sql = sql & ", LengthCalculated=" & lista(13)
            sql = sql & ", Supplier=" & lista(14)
            sql = sql & ", Manufacturer=" & lista(15)
            sql = sql & ", Location=" & lista(16)
            sql = sql & ", State=" & lista(17)
            sql = sql & ", RoughNess=" & lista(18)
            sql = sql & ", Sector=" & lista(19)
            sql = sql & ", InformationValidity=" & lista(20)
            sql = sql & ", DateInstallation=" & lista(21)
            sql = sql & ", SideStreet=" & lista(22)
            sql = sql & ", DividedDistance=" & lista(23)
            
            sql = sql & ", USUARIO_LOG=" & lista(24) '***** INCLUIDO EM 25/11/08 JONATHAS
            sql = sql & ", DATA_LOG=" & lista(25)    '***** INCLUIDO EM 25/11/08 JONATHAS
            
              sql = sql & " Where object_id_=" & lista(0)
             Else
             
              
    sa = "INITIALGROUNDHEIGHT"
   sb = "FINALGROUNDHEIGHT"
   sc = "INITIALTUBEDEEPNESS"
   sd = "FINALTUBEDEEPNESS"
   se = "INTERNALDIAMETER"
   sf = "EXTERNALDIAMETER"
   sg = "INITIALCOMPONENT"
   sh = "FINALCOMPONENT"
   si = "THICKNESS"
   sj = "MATERIAL"
   sl = "LENGTH"
   sm = "LENGTHCALCULATED"
   sn = "SUPPLIER"
   so = "MANUFACTURER"
   sp = "LOCATION"
   sq = "STATE"
   sr = "ROUGHNESS"
   ss = "SECTOR"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   sv = "SIDESTREET"
   sx = "DIVIDEDDISTANCE"
   sz = "USUARIO_LOG"
   sk = "DATA_LOG"
   sw = "USUARIO_LOG"
   aa = "SEWERLINES"
   bb = "DRAINLINES"
   swx = "ID_TYPE"
   sss = "YEAROFCONSTRUCTION"
   ssv = "NOTES"
   ssz = "TROUBLE"
   ssr = "PATTERN"
   ssa = "COMPONENT_ID"
          
                          
                   
          
             Dim tas, tas2 As String
             tas = lista(1)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
             
                sql = "Update " + """" + sst + """" + " Set"
            
             
            sql = sql & " " + """" + swx + """" + "='" & tas2 & "'"
           
             tas = lista(2)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sa + """" + "='" & tas2 & "'"
           
             tas = lista(3)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sb + """" + "='" & tas2 & "'"
           
             tas = lista(4)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sc + """" + "='" & tas2 & "'"
           
             tas = lista(5)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sd + """" + "='" & tas2 & "'"
           
             tas = lista(6)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + se + """" + "='" & tas2 & "'"
           
             tas = lista(7)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sf + """" + "='" & tas2 & "'"
           
             tas = lista(8)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sg + """" + "='" & tas2 & "'"
           
             tas = lista(9)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sh + """" + "='" & tas2 & "'"
           
             tas = lista(10)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + si + """" + "='" & tas2 & "'"
           
             tas = lista(11)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sj + """" + "='" & tas2 & "'"
           
             tas = lista(12)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sl + """" + "='" & tas2 & "'"
           
             tas = lista(13)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sm + """" + "='" & tas2 & "'"
           
             tas = lista(14)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sn + """" + "='" & tas2 & "'"
           
             tas = lista(15)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + so + """" + "='" & tas2 & "'"
           
             tas = lista(16)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sp + """" + "='" & tas2 & "'"
           
             tas = lista(17)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sq + """" + "='" & tas2 & "'"
           
             tas = lista(18)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sr + """" + "='" & tas2 & "'"
           
             tas = lista(19)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + ss + """" + "='" & tas2 & "'"
           
             tas = lista(20)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + st + """" + "='" & tas2 & "'"
           
             tas = lista(21)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + su + """" + "='" & tas2 & "'"
           
             tas = lista(22)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sv + """" + "='" & tas2 & "'"
           
             tas = lista(23)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sx + """" + "='" & tas2 & "'"
           
             tas = lista(24)
            tas2 = Replace(tas, "'", "")
              If tas2 = "" Then
             tas2 = "0"
              End If
            sql = sql & ", " + """" + sz + """" + "='" & tas2 & "'"
           
             tas = lista(25)
            tas2 = Replace(tas, "'", "") '***** INCLUIDO EM 25/11/08 JONATHAS
              If tas2 = "" Then
             tas2 = "0"
               End If
            sql = sql & ", " + """" + sk + """" + "='" & tas2 & "'"
           
             tas = lista(0)
            tas2 = Replace(tas, "'", "") '***** INCLUIDO EM 25/11/08 JONATHAS
              If tas2 = "" Then
             tas2 = "0"
               End If
              sql = sql & " Where " + """" + "OBJECT_ID_" + """" + "='" & tas2 & "'"
             
                         
                                                
             
             
                        
             
             
             
            ' sql = "Update " + """" + sst + """" + " Set"
            
             
          '  sql = sql & " " + """" + swx + """" + "='" & lista(1) & "'"
          '  sql = sql & ", " + """" + sa + """" + "='" & lista(2) & "'"
          '  sql = sql & ", " + """" + sb + """" + "='" & lista(3) & "'"
          '  sql = sql & ", " + """" + sc + """" + "='" & lista(4) & "'"
          '  sql = sql & ", " + """" + sd + """" + "='" & lista(5) & "'"
          '  sql = sql & ", " + """" + se + """" + "='" & lista(6) & "'"
           ' sql = sql & ", " + """" + sf + """" + "='" & lista(7) & "'"
          '  sql = sql & ", " + """" + sg + """" + "='" & lista(8) & "'"
          '  sql = sql & ", " + """" + sg + """" + "='" & lista(9) & "'"
           ' sql = sql & ", " + """" + si + """" + "='" & lista(10) & "'"
          '  sql = sql & ", " + """" + sj + """" + "='" & lista(11) & "'"
          '  sql = sql & ", " + """" + sl + """" + "='" & lista(12) & "'"
           ' sql = sql & ", " + """" + sm + """" + "='" & lista(13) & "'"
         '   sql = sql & ", " + """" + sn + """" + "='" & lista(14) & "'"
          '  sql = sql & ", " + """" + so + """" + "='" & lista(15) & "'"
          '  sql = sql & ", " + """" + sp + """" + "='" & lista(16) & "'"
          '  sql = sql & ", " + """" + sq + """" + "='" & lista(17) & "'"
         '   sql = sql & ", " + """" + sr + """" + "='" & lista(18) & "'"
         '   sql = sql & ", " + """" + ss + """" + "='" & lista(19) & "'"
          '  sql = sql & ", " + """" + st + """" + "='" & lista(20) & "'"
          '  sql = sql & ", " + """" + su + """" + "='" & lista(21) & "'"
         '   sql = sql & ", " + """" + sv + """" + "='" & lista(22) & "'"
        '    sql = sql & ", " + """" + sx + """" + "='" & lista(23) & "'"
          '
          '  sql = sql & ", " + """" + sz + """" + "='" & lista(24) & "'" '***** INCLUIDO EM 25/11/08 JONATHAS
           ' sql = sql & ", " + """" + sk + """" + "='" & lista(25) & "'" '***** INCLUIDO EM 25/11/08 JONATHAS
           '   sql = sql & " Where " + """" + sw + """" + "='" & lista(0) & "'"
End If
'MsgBox tipoConex
'
'            If tipoConex = 1 Then ' SQL
'                'SQL SERVER REQUER A SINTAXE no formato datalog = '2008-12-17 12:40'
'                'FORMATO RECEBIDO NA LISTA(25) = 17/12/08 10:14
'                '                          |          ANO            |            MES             |             DIA                |           HORAS
'                If Len(lista(26)) = 18 Then
'                    'MsgBox "tipo conex 1 = " & sql & ", DATALOG='20" & Mid(Trim(lista(25)), 10, 2) & "-" & Mid(Trim(lista(25)), 5, 2) & "-" & Mid(Trim(lista(25)), 2, 2) & " " & Mid(Trim(lista(25)), 13, 5) & "'"
'                    sql = sql & ", DATALOG='20" & Mid(Trim(lista(26)), 10, 2) & "-" & Mid(Trim(lista(26)), 5, 2) & "-" & Mid(Trim(lista(26)), 2, 2) & " " & Mid(Trim(lista(26)), 13, 5) & "'"
'
'                ElseIf Len(lista(26)) = 16 Then
'                    sql = sql & ", DATALOG='20" & Mid(Trim(lista(26)), 8, 2) & "-" & Mid(Trim(lista(26)), 5, 2) & "-" & Mid(Trim(lista(26)), 2, 2) & " " & Mid(Trim(lista(26)), 11, 5) & "'"
'                End If
'
'            ElseIf tipoConex = 2 Then 'ORACLE
'                'ORACLE REQUER A SINTAXE no formato datalog = '17/12/08'
'
'MsgBox lista(26) & " Tamanho " & Len(lista(26))
'
'                If Len(lista(26)) = 18 Then
'                    sql = sql & ", DATALOG='" & Mid(lista(26), 2, 6) & Mid(lista(26), 10, 2) & "'"
'
'                ElseIf Len(lista(26)) = 16 Then
'                    sql = sql & ", DATALOG='" & Mid(lista(26), 2, 6) & "20" & Mid(lista(26), 8, 2) & "'"
'                End If
'
'            End If
            
          
            
'MsgBox sql
        Else
        If tipoConex <> 4 Then
            sql = "Insert Into WaterLines"
            sql = sql & " (object_id_, Id_Type, InitialGroundHeight, FinalGroundHeight, InitialTubeDeepness, FinalTubeDeepness, InternalDiameter, ExternalDiameter, InitialComponent, FinalComponent,"
            sql = sql & " Thickness, MATERIAL,  Length, LengthCalculated, Supplier, Location, State, Sector, RoughNess, InformationValidity)"
            sql = sql & " Values (" & lista(1)
            sql = sql & ", " & lista(2)
            sql = sql & ", " & lista(3)
            sql = sql & ", " & lista(4)
            sql = sql & ", " & lista(5)
            sql = sql & ", " & lista(6)
            sql = sql & ", " & lista(7)
            sql = sql & ", " & lista(8)
            sql = sql & ", " & lista(9)
            sql = sql & ", " & lista(10)
            sql = sql & ", " & lista(11)
            sql = sql & ", " & lista(12)
            sql = sql & ", " & lista(13)
            sql = sql & ", " & lista(14)
            sql = sql & ", " & lista(15)
            sql = sql & ", " & lista(16)
            sql = sql & ", " & lista(17)
            sql = sql & ", " & lista(18)
            sql = sql & ", " & lista(19)
            sql = sql & ", " & lista(20)
            sql = sql & ")"
        
        
   Else
    
    sa = "INITIALGROUNDHEIGHT"
   sb = "FINALGROUNDHEIGHT"
   sc = "INITIALTUBEDEEPNESS"
   sd = "FINALTUBEDEEPNESS"
   se = "INTERNALDIAMETER"
   sf = "EXTERNALDIAMETER"
   sg = "INITIALCOMPONENT"
   sh = "FINALCOMPONENT"
   si = "THICKNESS"
   sj = "MATERIAL"
   sl = "LENGTH"
   sm = "LENGTHCALCULATED"
   sn = "SUPPLIER"
   so = "MANUFACTURER"
   sp = "LOCATION"
   sq = "STATE"
   sr = "ROUGHNESS"
   ss = "SECTOR"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   sv = "SIDESTREET"
   sx = "DIVIDEDDISTANCE"
   sz = "USUARIO_LOG"
   sk = "DATA_LOG"
   sw = "USUARIO_LOG"
   aa = "SEWERLINES"
   bb = "DRAINLINES"
   swx = "ID_TYPE"
   sss = "YEAROFCONSTRUCTION"
   ssv = "NOTES"
   ssz = "TROUBLE"
   ssr = "PATTERN"
   ssa = "COMPONENT_ID"
   
   
   tas = lista(1)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
 
    sql = "Insert Into " + """" + sst + """" + ""
            sql = sql & " (" + """" + sw + """" + ", " + """" + swx + """" + ", " + """" + sa + """" + ", " + """" + sb + """" + ", " + """" + sc + """" + ", " + """" + sd + """" + ", " + """" + se + """" + ", " + """" + sf + """" + ", " + """" + sg + """" + ", " + """" + sh + """" + ","
            sql = sql & " " + """" + si + """" + ", " + """" + sj + """" + ",  " + """" + sl + """" + ", " + """" + sm + """" + ", " + """" + sn + """" + ", " + """" + sp + """" + ", " + """" + sq + """" + ", " + """" + ss + """" + ", " + """" + sr + """" + ", " + """" + st + """" + ")"
            sql = sql & " Values ('" & tas2 & "'"
            
            tas = lista(2)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            
            tas = lista(3)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            
            tas = lista(4)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            tas = lista(5)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            
             tas = lista(6)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            sql = sql & ", '" & tas2 & "'"
           
            tas = lista(7)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            sql = sql & ", '" & tas2 & "'"
            
            tas = lista(8)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            sql = sql & ", '" & tas2 & "'"
            
            tas = lista(9)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            sql = sql & ", '" & tas2 & "'"
            
            tas = lista(10)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            sql = sql & ", '" & tas2 & "'"
            
            tas = lista(11)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            sql = sql & ", '" & tas2 & "'"
            
             tas = lista(12)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
                        
            sql = sql & ", '" & tas2 & "'"
            tas = lista(13)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            tas = lista(14)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            tas = lista(15)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            tas = lista(16)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            tas = lista(17)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            tas = lista(18)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            tas = lista(19)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
            tas = lista(20)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
            
            sql = sql & ", '" & tas2 & "'"
           
            sql = sql & ")"
        End If
        End If
        getWaterLines = sql

'        '************* MONITORAMENTO ***************
'        Close #2
'        Open App.Path & "\GeoSanLog.txt" For Append As #2
'        Print #2, Now & " Function getWaterLines SQL = " & sql
'        Close #2
'        '***************** FIM *********************


Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
      Else
      'Open App.Path & "\GeoSanLog.txt" For Append As #1
      'Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function getWaterLines - " & Err.Number & " - " & Err.Description
      'Close #1
      'MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function

Public Function getWaterComponents(str As String, rs As Integer) As String
On Error GoTo Trata_Erro

    Dim sql As String, lista() As String
  
      Dim tas As String
    Dim tas2 As String
     
       lista = Split(str, ",")
    
    
    If rs > 0 Then
     If tipoConex <> 4 Then
        sql = "Update WaterComponents Set"
        sql = sql & " Id_Type=" & lista(1)
        sql = sql & ", YearOfConstruction=" & lista(2)
        sql = sql & ", State=" & lista(3)
        sql = sql & ", Location=" & lista(4)
        sql = sql & ", Supplier=" & lista(5)
        sql = sql & ", Manufacturer=" & lista(6)
        sql = sql & ", GroundHeight=" & lista(7)
        sql = sql & ", Demand=" & lista(8)
        sql = sql & ", CalculeNode=" & lista(9)
        sql = sql & ", InformationValidity=" & lista(10)
        sql = sql & ", Notes=" & lista(11)
        sql = sql & ", Trouble=" & lista(12)
        sql = sql & ", DateInstallation=" & lista(13)
        sql = sql & ", Pattern=" & lista(14)
        sql = sql & ", Sector=" & lista(15)
        sql = sql & " Where Component_id=" & lista(0)
        Else
         sa = "INITIALGROUNDHEIGHT"
   sb = "FINALGROUNDHEIGHT"
   sc = "INITIALTUBEDEEPNESS"
   sd = "FINALTUBEDEEPNESS"
   se = "INTERNALDIAMETER"
   sf = "EXTERNALDIAMETER"
   sg = "INITIALCOMPONENT"
   sh = "FINALCOMPONENT"
   si = "THICKNESS"
   sj = "MATERIAL"
   sl = "LENGTH"
   sm = "LENGTHCALCULATED"
   sn = "SUPPLIER"
   so = "MANUFACTURER"
   sp = "LOCATION"
   sq = "STATE"
   sr = "ROUGHNESS"
   ss = "SECTOR"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   sv = "SIDESTREET"
   sx = "DIVIDEDDISTANCE"
   sz = "USUARIO_LOG"
   sk = "DATA_LOG"
   sw = "OBJECT_ID_"
   aa = "SEWERLINES"
   bb = "DRAINLINES"
   swx = "ID_TYPE"
   sss = "YEAROFCONSTRUCTION"
   ssv = "NOTES"
   ssz = "TROUBLE"
   ssr = "PATTERN"
   ssa = "COMPONENT_ID"
   sst = "WATERLINES"
   ssq = "WATERCOMPONENTS"
   sse = "DEMAND"
   ssj = "CALCULENODE"
   ssd = "LINE_ID"
    lista = Split(str, ",")
        
        sql = "Update " + """" + ssq + """" + " Set"
        tas = lista(1)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & " " + """" + swx + """" + "='" & tas2 & "'"
        tas = lista(2)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + sss + """" + "='" & tas2 & "'"
        tas = lista(3)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + sq + """" + "='" & tas2 & "'"
        tas = lista(4)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + sp + """" + "='" & tas2 & "'"
        tas = lista(5)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + sn + """" + "='" & tas2 & "'"
        tas = lista(6)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + so + """" + "='" & tas2 & "'"
        tas = lista(7)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + sa + """" + "='" & tas2 & "'"
        tas = lista(8)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + sse + """" + "='" & tas2 & "'"
        tas = lista(9)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + ssj + """" + "='" & tas2 & "'"
        tas = lista(10)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + st + """" + "='" & tas2 & "'"
        tas = lista(11)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + ssv + """" + "='" & tas2 & "'"
        tas = lista(12)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + ssz + """" + "='" & tas2 & "'"
        tas = lista(13)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + su + """" + "='" & tas2 & "'"
        tas = lista(14)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + ssr + """" + "='" & tas2 & "'"
        tas = lista(15)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", " + """" + ss + """" + "='" & tas2 & "'"
        tas = lista(0)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & " Where " + """" + ssa + """" + "='" & tas2 & "'"
        End If
    Else
     If tipoConex <> 4 Then
        sql = "Insert Into WaterComponents"
        sql = sql & " (object_id_, Id_Type, YearOfConstruction, State, Location, Manufacturer, GroundHeight,InformationValidity, Notes, Demand, CalculeNode)"
        sql = sql & " Values (" & lista(1)
        sql = sql & ", " & lista(2)
        sql = sql & ", " & lista(3)
        sql = sql & ", " & lista(4)
        sql = sql & ", " & lista(5)
        sql = sql & ", " & lista(6)
        sql = sql & ", " & lista(9)
        sql = sql & ", " & lista(10)
        sql = sql & ", " & lista(7)
        sql = sql & ", " & lista(8)
        sql = sql & ", " & lista(11)
        sql = sql & ")"
       Else
       
        sa = "INITIALGROUNDHEIGHT"
   sb = "FINALGROUNDHEIGHT"
   sc = "INITIALTUBEDEEPNESS"
   sd = "FINALTUBEDEEPNESS"
   se = "INTERNALDIAMETER"
   sf = "EXTERNALDIAMETER"
   sg = "INITIALCOMPONENT"
   sh = "FINALCOMPONENT"
   si = "THICKNESS"
   sj = "MATERIAL"
   sl = "LENGTH"
   sm = "LENGTHCALCULATED"
   sn = "SUPPLIER"
   so = "MANUFACTURER"
   sp = "LOCATION"
   sq = "STATE"
   sr = "ROUGHNESS"
   ss = "SECTOR"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   sv = "SIDESTREET"
   sx = "DIVIDEDDISTANCE"
   sz = "USUARIO_LOG"
   sk = "DATA_LOG"
   sw = "OBJECT_ID_"
   aa = "SEWERLINES"
   bb = "DRAINLINES"
   swx = "ID_TYPE"
   sss = "YEAROFCONSTRUCTION"
   ssv = "NOTES"
   ssz = "TROUBLE"
   ssr = "PATTERN"
   ssa = "COMPONENT_ID"
   sst = "WATERLINES"
   ssq = "WATERCOMPONENTS"
   sse = "DEMAND"
   ssj = "CALCULENODE"
   ssd = "LINE_ID"
    lista = Split(str, ",")
          
        sql = "Insert Into " + """" + ssq + """" + ""
        sql = sql & " (" + """" + sw + """" + ", " + """" + swx + """" + ", " + """" + sss + """" + ", " + """" + sq + """" + ", " + """" + sp + """" + ", " + """" + so + """" + ", " + """" + sa + """" + "," + """" + st + """" + ", " + """" + ssv + """" + ", " + """" + sse + """" + ", " + """" + ssj + """" + ")"
        tas = lista(1)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & " Values ('" & tas2 & "'"
        tas = lista(2)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        tas = lista(3)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        tas = lista(4)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        tas = lista(5)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        tas = lista(6)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        tas = lista(9)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        tas = lista(10)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        tas = lista(7)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        tas = lista(8)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        tas = lista(11)
             tas2 = Replace(tas, "'", "")
             If tas2 = "" Then
             tas2 = "0"
             End If
        sql = sql & ", '" & tas2 & "'"
        sql = sql & ")"
    End If
        
        
    End If
    
  

    getWaterComponents = sql

'    '************* MONITORAMENTO ***************
'    Close #2
'    Open App.Path & "\GeoSanLog.txt" For Append As #2
'    Print #2, Now & " Function getWaterComponents SQL = " & sql
'    Close #2
'    '***************** FIM *********************


   
    


'    '************* MONITORAMENTO ***************
'    Close #2
'    Open App.Path & "\GeoSanLog.txt" For Append As #2
'    Print #2, Now & " Function getWaterComponents SQL = " & sql
'    Close #2
'    '***************** FIM *********************

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
      Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function getWaterComponents - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
    
End Function

Public Function convertQuery(sql As String, tipo As Integer) As String
On Error GoTo Trata_Erro

   If tipo = 2 Then
      sql = Replace(sql, "[", "")
      sql = Replace(sql, "]", "")
   End If
   convertQuery = sql

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - mdlQuerys - Public Function convertQuery - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If


End Function







