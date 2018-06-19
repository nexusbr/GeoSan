VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl Manager 
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   ScaleHeight     =   3885
   ScaleWidth      =   4755
   Begin VB.TextBox txtInput 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3570
      TabIndex        =   2
      Text            =   "txtInput"
      Top             =   1350
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ListBox LstOpt 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   840
      Index           =   0
      Left            =   2970
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1170
      Picture         =   "Manager.ctx":0000
      ScaleHeight     =   240
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   2370
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3555
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6271
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      BackColor       =   12648447
      ForeColor       =   0
      BackColorFixed  =   65535
      ForeColorFixed  =   0
      BackColorSel    =   32896
      BackColorBkg    =   65535
      GridColorFixed  =   12632256
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ma As String
Dim mu As String
Dim mi As String
Dim mo As String

Public Enum ModeSelect
   Mode_Single = 0
   mode_multiple = 1
   Mode_Single_Alias = 2
   Mode_To_Insert = 3
   Mode_Multiple_Dif = 4
End Enum


Private Foco As Boolean
Public conn As ADODB.Connection
Public Itens As CItens

Public layerName As String
Public StepBy As String
Public Large As Long
Public PerCol0 As Integer
Public PerCol1 As Integer
Public tipoProvider As Integer
 Dim aspas As String

 
 Dim contador As Integer
 



Public Function InitConn(Cn As ADODB.Connection, TypeConnection As Integer)
On Error GoTo Trata_Erro
   
   
   tipoConex = TypeConnection
   
   Set conn = Cn
   tipoProvider = TypeConnection
   Set Itens = New CItens
   If Large = 0 Then Large = 100
   If PerCol0 = 0 Then PerCol0 = 50
   If PerCol1 = 0 Then PerCol1 = 50
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Public Function InitConn - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
End Function

Public Function ResetConn(Cn As ADODB.Connection)
   Set Itens = Nothing
   Set conn = Nothing
End Function

Public Function LoadDefaultProperties(object_id_ As String, LayerName_ As String, Optional ToInsert As Boolean) As Boolean
   aspas = ""
   'MsgBox aspas
On Error GoTo Trata_Erro
Dim linha As String
   
   Dim RsPM As ADODB.Recordset, RsRl As ADODB.Recordset, RsB As Recordset, RsPS As ADODB.Recordset
   Dim IcontFil As Integer, FSelec As Boolean
  
   
 
 
   contador = 0
     
 Set RsPM = New ADODB.Recordset
   ClearObjects
   linha = "1.1"
   'Código inserido em 03/04/2007 em substituição às linhas comentadas em seguida para eliminar procedures ®viviani
    
   Set RsRl = New ADODB.Recordset
   Set RsB = New ADODB.Recordset
   
     Set RsPM = New ADODB.Recordset
   
   
   
   
   
   
  
    Dim a1 As String
Dim a2 As String
Dim a3 As String
Dim a4 As String
Dim a5 As String
Dim a6 As String
Dim a7 As String
Dim a8 As String


a1 = "X_MANAGERPROPERTIESB"
a2 = "TABLENAMEIN"
a3 = "X_MANAGERPROPERTIES"
a4 = "FIELDNAMERIN"
a5 = "TABLENAME"
a6 = "FIELDNAME"

a7 = "FIELDNAMERIN"
a8 = "FIELDSEQUENCE"
     
   ' Text1.Text = mdlQuerys.getPmsdp(LayerName_, 3, object_id_, tipoProvider)
  
    RsB.CursorLocation = adUseClient
 If tipoConex <> 4 Then
   RsB.Open "Select * from X_ManagerPropertiesB Where TABLENAME ='" & UCase(LayerName_) & "'", conn, adOpenDynamic, adLockOptimistic
   Else
    RsB.Open "Select * from " + """" + a1 + """" + " Where " + """" + a5 + """" + "='" & LayerName_ & "'", conn, adOpenDynamic, adLockOptimistic
    

'MsgBox "ARQUIVO DEBUG SALVO"
' WritePrivateProfileString "A", "A", "Select * from " + """" + a1 + """" + " Where " + """" + a5 + """" + "='" & LayerName_ & "'", App.Path & "\DEBUG.INI"
     End If
   linha = "3"
   RsRl.CursorLocation = adUseClient
   
     
    If tipoConex <> 4 Then
    RsRl.Open "Select X_ManagerProperties.*,  Upper(FieldNameRIn) As fLink from X_ManagerProperties Where Upper(TableNamein)='" & UCase(LayerName_) & "'", conn, adOpenDynamic, adLockOptimistic
  
  Else
      
  
   RsRl.Open "Select * from " + """" + a3 + """" + " Where " + """" + a2 + """" + "='" & LayerName_ & "'", conn, adOpenDynamic, adLockOptimistic
  
   End If
   
   

   
    If ToInsert Then
 
        Set RsPM = conn.Execute(mdlQuerys.getPmsdp(LayerName_, 3, object_id_, tipoProvider))
       ' Dim va As String
       ' va = mdlQuerys.getPmsdp(LayerName_, 3, object_id_, tipoProvider)
       'RsPM.Open( va, conn, adOpenDynamic, adLockOptimistic
    Else
        Set RsPM = conn.Execute(mdlQuerys.getPmsdp(LayerName_, 2, object_id_, tipoProvider))
    End If
 
    If Trim(object_id_) = "" Then
        Exit Function
    End If

'   'Ocorrência de procedure
'   'PmSDP
'   If ToInsert Then
'      Set RsPM = Conn.Execute(PmSDP(LayerName_, 3, Object_id_))
'   Else
'      Set RsPM = Conn.Execute("pmsdp '" & LayerName_ & "'," & 2 & "," & Object_id_ & "," & 0)
'   End If
  ' If tipoConex <> 4 Then 'alterado em 21/10/2010
   
   linha = "1.4"
   StepBy = "ClearObjects"
   linha = "1.5"

   linha = "2"
   
   


'"Provider=PostgreSQL.1;Password=gustavo;User ID=postgres;Data Source=localhost;Location=geopost"
 
   
    ' RsB.Filter = "TABLENAME='WATERLINES'  AND FIELDSEQUENCE='0'"
   
   linha = "4"
   StepBy = "OpenObjectsRs"
   
   'Else
   
   
    'mu = UPPER(TableName)
    mi = "LayerName_"
    'mo = UPPER(FieldNameRIn)
    Dim gu As String
    'gu = UPPER(TableNameIn)

   
 
   
   'MsgBox "Select * from " + ma + " Where UPPER(""+TableName+"") ='" & UCase(LayerName_) & "'"


   
   If Not RsPM.EOF Then
      For IcontFil = 0 To RsPM.Fields.Count - 1
         'Atribui nova linha na Grid
         If Not RsPM.Fields(IcontFil).Type = adDBTimeStamp Then
            If IcontFil > 0 Then
              Grid.AddItem RsPM.Fields(IcontFil).Name, IcontFil
              Grid.TextArray(IcontFil * Grid.Cols + Grid.Col) = IIf(IsNull(RsPM.Fields(IcontFil).Value), 0, Trim(RsPM.Fields(IcontFil).Value))
              
            Else
              Grid.TextArray(IcontFil * Grid.Cols) = RsPM.Fields(IcontFil).Name
              Grid.TextArray(IcontFil * Grid.Cols + Grid.Col) = IIf(IsNull(RsPM.Fields(IcontFil).Value), 0, Trim(RsPM.Fields(IcontFil).Value))
            End If
         Else
            If IcontFil > 0 Then
              Grid.AddItem RsPM.Fields(IcontFil).Name, IcontFil
              Grid.TextArray(IcontFil * Grid.Cols + Grid.Col) = IIf(IsNull(RsPM.Fields(IcontFil).Value), "", Trim(RsPM.Fields(IcontFil).Value))
              
            Else
              Grid.TextArray(IcontFil * Grid.Cols) = RsPM.Fields(IcontFil).Name
              Grid.TextArray(IcontFil * Grid.Cols + Grid.Col) = IIf(IsNull(RsPM.Fields(IcontFil).Value), "", Trim(RsPM.Fields(IcontFil).Value))
            End If
         End If
         'Verifica se atributo é de seleção
         'RsRl.Filter = "FieldNameRIn='" & UCase(RsPM.Fields(IcontFil).Name) & "'"
          If tipoConex <> 4 Then
         RsRl.Filter = "Flink='" & Trim(UCase(RsPM.Fields(IcontFil).Name)) & "'"
         Else
         'MsgBox RsPM.Fields(IcontFil).Name
          RsRl.Filter = "FIELDNAMERIN='" & Trim(RsPM.Fields(IcontFil).Name) & "'"
         End If
         If Not RsRl.EOF Then
            FSelec = True
            linha = "5"
             If tipoConex <> 4 Then
            Set RsPS = conn.Execute("Select " & RsRl.Fields("FieldNameRel") & _
            " as  " + """" + "ID_type" + """" + "," & RsRl.Fields("FieldNameOut") & " as " + """" + "Description_" + """" + " From " & RsRl.Fields("TableNameOut"))
            
            Else
              Set RsPS = conn.Execute("Select " + """" + UCase(RsRl.Fields("FieldNameRel")) + """" + _
            " as  " + """" + "ID_TYPE" + """" + "," + """" + UCase(RsRl.Fields("FieldNameOut")) + """" + " as " + """" + "DESCRIPTION_" + """" + " From " + """" + UCase(RsRl.Fields("TableNameOut")) + """" + "")
            
            End If
            linha = "6"
            'Novo objetos de seleção(pictures e list)
            LoadObjectSelect IcontFil
            'Carrega Nova List
            While Not RsPS.EOF
               Dim type2 As String
               type2 = RsPS.Fields("ID_TYPE").Value
               If RsPM.Fields(IcontFil).Value = type2 Then
                  Grid.TextArray(IcontFil * Grid.Cols + Grid.Col) = RsPS.Fields("DESCRIPTION_").Value
               End If
               LstOpt(IcontFil).AddItem RsPS.Fields("DESCRIPTION_")
               LstOpt(IcontFil).ItemData(LstOpt(IcontFil).NewIndex) = RsPS.Fields("ID_TYPE")
               RsPS.MoveNext
            Wend
            RsPS.Close
            Set RsPS = Nothing
         Else
            FSelec = False
         End If
         'Dim va1 As String
         'va1 = RsPM.Fields(IcontFil).Name
         
        RsB.Filter = "FIELDNAME='" & RsPM.Fields(IcontFil).Name & "' OR FIELDSEQUENCE='" & IcontFil & "'"
        
           ' RsB.Filter = "FIELDNAME='WATERLINES' " 'Or FIELDSEQUENCE='0'"
        
        
         'RsB.Filter = "FIELDNAME='" & va1 & "' Or FIELDSEQUENCE='0'"
              'RsB.Filter = "FIELDNAME='" & RsPM.Fields(IcontFil).Name & "' Or FIELDSEQUENCE='23'"
         'Atribui novo campo na classe
         StepBy = "SetCls"
         
'         DESABILITADO EM 27/11/2008 JONATHAS
'         If RsPM.Fields(IcontFil).Type = 131 Or RsPM.Fields(IcontFil).Type = 139 Then 'Numerico
'            Itens.Add IIf(RsB.EOF, True, False), 0, 0, FSelec, 0, 0, IIf(RsPM.Fields(IcontFil).NumericScale > 0, 1131, 131), _
'                RsPM.Fields(IcontFil).Name, 0, IIf(IsNull(RsPM.Fields(IcontFil).Value), 0, RsPM.Fields(IcontFil).Value), , , CStr(IcontFil)
'         Else
'            Itens.Add IIf(RsB.EOF, True, False), 0, 0, FSelec, 0, 0, RsPM.Fields(IcontFil).Type, _
'                RsPM.Fields(IcontFil).Name, 0, IIf(IsNull(RsPM.Fields(IcontFil).Value), 0, RsPM.Fields(IcontFil).Value), , , CStr(IcontFil)
'         End If
         
        'INCLUIDAS FUNÇÕES PARA INSERIR CAMPOS USUÁRIO E DATA CADASTRO 27/11/2008
        If RsPM.Fields(IcontFil).Name = "USUÁRIO" Or RsPM.Fields(IcontFil).Name = "DATA CADASTRO" Then
               Itens.Add False, 0, 0, False, 0, 0, RsPM.Fields(IcontFil).Type, _
                   RsPM.Fields(IcontFil).Name, 0, IIf(IsNull(RsPM.Fields(IcontFil).Value), 0, Trim(RsPM.Fields(IcontFil).Value)), , , CStr(IcontFil)
            'ITENS.Add ENABLED, TYPE, SUB_TYPE, SELECTION, MAX, MIN, DATA TYPE, NAME, V.DISPLAY, V.STORE, ESPECIFIC?, CHANGED?, SKEY
        Else
            If RsPM.Fields(IcontFil).Type = 131 Or RsPM.Fields(IcontFil).Type = 139 Then 'Numerico
               Itens.Add IIf(RsB.EOF, True, False), 0, 0, FSelec, 0, 0, IIf(RsPM.Fields(IcontFil).NumericScale > 0, 1131, 131), _
                   RsPM.Fields(IcontFil).Name, 0, IIf(IsNull(RsPM.Fields(IcontFil).Value), 0, RsPM.Fields(IcontFil).Value), , , CStr(IcontFil)
            Else
               Itens.Add IIf(RsB.EOF, True, False), 0, 0, FSelec, 0, 0, RsPM.Fields(IcontFil).Type, _
                   RsPM.Fields(IcontFil).Name, 0, IIf(IsNull(RsPM.Fields(IcontFil).Value), 0, Trim(RsPM.Fields(IcontFil).Value)), , , CStr(IcontFil)
            End If
        End If
        
        
        
   
      
      
      Next
      StepBy = "CarregaPadrão"
   End If
   layerName = LayerName_
   If Not (RsB Is Nothing) Then
      If RsB.State = adStateOpen Then RsB.Close
   End If
   Set RsB = Nothing
   If Not (RsRl Is Nothing) Then
      If RsRl.State = adStateOpen Then RsRl.Close
   End If
   Set RsRl = Nothing
   If Not (RsPM Is Nothing) Then
      If RsPM.State = adStateOpen Then RsPM.Close
   End If
   Set RsPM = Nothing
   If StepBy = "CarregaPadrão" Then
   If tipoConex <> 4 Then
      If UCase(layerName) = "WATERLINES" Or UCase(layerName) = "WATERCOMPONENTS" Or UCase(layerName) = "SEWERLINES" Or UCase(layerName) = "SEWERCOMPONENTS" Or UCase(layerName) = "DRAINLINES" Or UCase(layerName) = "DRAINCOMPONENTS" Then
         If LoadSpecificProperties(LayerName_, object_id_, Itens.Item(2).ValueStore) Then
            LoadDefaultProperties = True
         End If
      End If
      Else
        If UCase(layerName) = "WATERLINES" Or UCase(layerName) = "WATERCOMPONENTS" Or UCase(layerName) = "SEWERLINES" Or UCase(layerName) = "SEWERCOMPONENTS" Or UCase(layerName) = "DRAINLINES" Or UCase(layerName) = "DRAINCOMPONENTS" Then
         If LoadSpecificProperties(LayerName_, object_id_, Itens.Item(2).ValueStore) Then
            LoadDefaultProperties = True
         End If
      End If
      
      End If
      
      
   End If
   txtInput.Visible = False
   Exit Function

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Public Function LoadDefaultProperties - " & "Linha = " & linha & " - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência. ", vbInformation
   End If

'LoadDefaultProperties_Error:
' MsgBox Err.Description & " - " & StepBy, vbCritical
End Function

Private Sub ClearObjectsSpecifics()
On Error GoTo Trata_Erro
   
   Dim a As Integer
   For a = Itens.Count To 1 Step -1
      If Itens.Item(a).Specific_ Then
         If Itens.Item(a).Selection_ Then
            Unload LstOpt(a - 1)
            Unload Pic(a - 1)
         End If
         Itens.Remove a
         Grid.RemoveItem a - 1
      Else
         Exit Sub
      End If
   Next

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub ClearObjectsSpecifics - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If


End Sub

Private Sub ClearObjects()
On Error GoTo Trata_Erro
   Dim a  As Integer
   If Not (Itens Is Nothing) Then
      For a = Itens.Count To 1 Step -1
         If Itens.Item(a).Selection_ Then
            Unload LstOpt(a - 1)
            Unload Pic(a - 1)
         End If
         Itens.Remove a
         If a > 1 Then Grid.RemoveItem a - 1
      Next
   End If
   txtInput = ""
   txtInput.Visible = False
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub ClearObjects - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

   
   
End Sub


'#################################################################
'
' Carregar as Prorpiedades específicas dos Objectos
'
'#################################################################

Private Function LoadSpecificProperties(LayerName_ As String, object_id_ As String, TypeObject As Integer) As Boolean
   
'MsgBox "Private Function LoadSpecificProperties(LayerName_ As String, object_id_ As String, TypeObject As Integer) As Boolean"


On Error GoTo Trata_Erro
   
   Dim RsPS As ADODB.Recordset, RsPSelect As ADODB.Recordset, IcontFil As Integer

   ClearObjectsSpecifics
   StepBy = "ClearObjectsSpecifics"
   IcontFil = Grid.Rows
   
   'Código inserido em 03/04/2007 em substituição às linhas comentadas em seguida para eliminar procedures ®viviani
   Set RsPS = conn.Execute(mdlQuerys.getPmssp(layerName, TypeObject, object_id_, conn, tipoProvider))
'   'Ocorrência de procedure
'   'PmSSP
'   Set RsPS = Conn.Execute("PmSSP '" & LayerName_ & "'," & TypeObject & ", " & Object_id_)
   While Not RsPS.EOF
 If tipoConex <> 4 Then
      Grid.AddItem RsPS!Description_, IcontFil
   
      Else
       Grid.AddItem RsPS!Description, IcontFil
      
      End If
      
      
      Grid.TextArray(IcontFil * Grid.Cols + Grid.Col) = IIf(IsNull(RsPS!Value_), RsPS!Value_Ref, RsPS!Value_)

 If tipoConex <> 4 Then
      Itens.Add True, RsPS!id_Type, RsPS!id_SubType, RsPS!Selection_, RsPS!Max_, RsPS!Min_, _
          RsPS!DataType, RsPS!Description_, IIf(IsNull(RsPS!Value_), RsPS!Value_Ref, RsPS!Value_), RsPS!Value_Ref, True
          
          Else
          Itens.Add True, RsPS!id_Type, RsPS!id_SubType, RsPS!Selection_, RsPS!Max_, RsPS!Min_, _
          RsPS!DataType, RsPS!Description, IIf(IsNull(RsPS!Value_), RsPS!Value_Ref, RsPS!Value_), RsPS!Value_Ref, True
          
          
          End If
      If RsPS.Fields("Selection_") Then
      
      'Código inserido em 03/04/2007 em substituição às linhas comentadas em seguida para eliminar procedures ®viviani
      Set RsPSelect = conn.Execute(getPmSpo(layerName, RsPS.Fields("Id_Type"), RsPS.Fields("Id_SubType")))
'       'Ocorrência de procedure
'       'PmSPO
'        Set RsPSelect = Conn.Execute("PmSPO '" & LayerName_ & "'," & RsPS.Fields("Id_Type") & "," & RsPS.Fields("Id_SubType"))
         
         If RsPSelect.EOF Then
            Exit Function
         End If
         
         LoadObjectSelect IcontFil
         
         While Not RsPSelect.EOF
            With RsPSelect
               LstOpt(IcontFil).AddItem .Fields("Option_")
               LstOpt(IcontFil).ItemData(LstOpt(IcontFil).NewIndex) = .Fields("Value_")
               .MoveNext
            End With
         Wend
         RsPSelect.Close
         Set RsPSelect = Nothing
      
      End If
      RsPS.MoveNext
      IcontFil = IcontFil + 1
   Wend
   txtInput.Visible = False
   RsPS.Close
   Set RsPS = Nothing
   
   
   Exit Function
Exit_Lsp:
   If Not (RsPS Is Nothing) Then
      If RsPS.State = adStateOpen Then RsPS.Close
   End If
   Set RsPS = Nothing

   If Not (RsPSelect Is Nothing) Then
      If RsPSelect.State = adStateOpen Then RsPSelect.Close
   End If
   Set RsPSelect = Nothing
   

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Function LoadSpecificProperties - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function

'Linhas comentadas em 04/04/2007 por ®viviani
'Private Function PmSDP(LayerName_ As String, TypeSelect As ModeSelect, Object_id_ As String, Optional V_Obj As String) As String
''Ocorrência de procedure
'   'PmSDP
'   PmSDP = "PmSDP '" & LayerName_ & "'," & TypeSelect & "," & Object_id_ & ",'" & V_Obj & "'"
'End Function

Public Function LoadComunsObjects(VarObjects As String, LayerName_ As String, Optional Adm As Boolean)
   
'MsgBox "Public Function LoadComunsObjects(VarObjects As String, LayerName_ As String, Optional Adm As Boolean)"
   
On Error GoTo Trata_Erro

'   On Error GoTo LoadComunsObjects_Err
   LoadDefaultProperties Left(VarObjects, InStr(1, VarObjects, ",") - 1), LayerName_
   
   Dim RsPS As ADODB.Recordset
   Dim a As Integer, LastValue As String
   
   'Código inserido em 03/04/2007 em substituição às linhas comentadas em seguida para eliminar procedures ®viviani
   Set RsPS = conn.Execute(mdlQuerys.getPmsdp(layerName, mode_multiple, VarObjects, tipoProvider))
   
'   'Ocorrência de procedure
'   'PmSDP
'   Set RsPS = Conn.Execute(PmSDP(layername, mode_multiple, 0, VarObjects))
   
   'ClearObjectsSpecifics
   If Not Adm Then
      For a = 0 To RsPS.Fields.Count - 1
         LastValue = IIf(IsNull(RsPS.Fields(a).Value), "", RsPS.Fields(a).Value)
         Itens.Item(a + 1).Enabled_ = True
         Do While Not RsPS.EOF
            If Not RsPS.Fields(a) = LastValue Then
               Itens.Item(a + 1).Enabled_ = False
               If Itens.Item(a + 1).Selection_ Then
                  Itens.Item(a + 1).Selection_ = False
                  Unload LstOpt(a)
                  Unload Pic(a)
               End If
               Grid.TextMatrix(a, 1) = ""
               Exit Do
            End If
            RsPS.MoveNext
         Loop
         'RsPS.Requery
         RsPS.MoveFirst
      Next
   End If
   RsPS.Close
   Set RsPS = Nothing
   Exit Function


Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   ElseIf Err.Number = -2147467259 Then
      Close #1
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Public Function LoadComunsObjects - " & Err.Number & " - " & Err.Description
      Close #1
      RsPS.Requery
      Resume
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Public Function LoadComunsObjects - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If


End Function

Public Function SaveMultProperties(object_id_ As String, LayerName_ As String) As Boolean
On Error GoTo Trata_Erro
   
   'MsgBox "Public Function SaveMultProperties(object_id_ As String, LayerName_ As String) As Boolean"
   
   'On Error GoTo sair_SaveMultProperties
   Dim a As Integer, strSQL As String, rs As ADODB.Recordset, JaInseriu As Boolean
   
   'Código inserido em 03/04/2007 em substituição às linhas comentadas em seguida para eliminar procedures ®viviani
   Set rs = conn.Execute(mdlQuerys.getPmsdp(LayerName_, Mode_Single, 0, tipoProvider))
      Dim a1 As String
Dim a2 As String
Dim a3 As String
Dim a4 As String
Dim a5 As String
Dim a6 As String
Dim a7 As String
Dim a8 As String

 If tipoConex <> 4 Then
   strSQL = "update " & LayerName_ & " set "
   
   For a = 0 To rs.Fields.Count - 1
      If JaInseriu Then
         If Itens.Item(a + 1).Changed_ Then
            Select Case Itens.Item(a + 1).DataType
               Case MyFloat, MyDecimal, MyBit, MyInt, MyNumber, MySmallInt, MyNumberScale, MyTinyInt
                  strSQL = strSQL & "," & rs.Fields(a).Name & "=" & Replace(Itens.Item(a + 1).ValueStore, ",", ".")
               Case MyDateTime
                  strSQL = strSQL & "," & rs.Fields(a).Name & "=" & IIf(Itens.Item(a + 1).ValueStore = "", "NULL", "'" & Format(Itens.Item(a + 1).ValueStore, "yyyymmdd") & "'")
               Case Else
                  strSQL = strSQL & "," & rs.Fields(a).Name & "='" & Itens.Item(a + 1).ValueStore & "'"
            End Select
         End If
      Else
         If Itens.Item(a + 1).Changed_ Then
            Select Case Itens.Item(a + 1).DataType
               Case MyFloat, MyDecimal, MyBit, MyInt, MyNumber, MySmallInt, MyNumberScale, MyTinyInt
                  strSQL = strSQL & rs.Fields(a).Name & "=" & Replace(Itens.Item(a + 1).ValueStore, ",", ".")
               Case MyDateTime
                  strSQL = strSQL & rs.Fields(a).Name & "=" & IIf(Itens.Item(a + 1).ValueStore = "", "NULL", "'" & Format(Itens.Item(a + 1).ValueStore, "yyyymmdd") & "'")
               Case Else
                  strSQL = strSQL & rs.Fields(a).Name & "='" & Itens.Item(a + 1).ValueStore & "'"
            End Select
            JaInseriu = True
         End If
      End If
   Next
   strSQL = strSQL & " Where " & rs(0).Name & " in (" & object_id_ & ")"
   
   Else
   
    strSQL = "update " + """" + LayerName_ + """" + "set "
   
   For a = 0 To rs.Fields.Count - 1
      If JaInseriu Then
         If Itens.Item(a + 1).Changed_ Then
            Select Case Itens.Item(a + 1).DataType
               Case MyFloat, MyDecimal, MyBit, MyInt, MyNumber, MySmallInt, MyNumberScale, MyTinyInt
                  strSQL = strSQL & "," & rs.Fields(a).Name & "='" & Replace(Itens.Item(a + 1).ValueStore, ",", ".")
               Case MyDateTime
                  strSQL = strSQL & "," & rs.Fields(a).Name & "='" & IIf(Itens.Item(a + 1).ValueStore = "", "NULL", "'" & Format(Itens.Item(a + 1).ValueStore, "yyyymmdd") & "'")
               Case Else
                  strSQL = strSQL & "," & rs.Fields(a).Name & "='" & Itens.Item(a + 1).ValueStore & "'"
            End Select
         End If
      Else
         If Itens.Item(a + 1).Changed_ Then
            Select Case Itens.Item(a + 1).DataType
               Case MyFloat, MyDecimal, MyBit, MyInt, MyNumber, MySmallInt, MyNumberScale, MyTinyInt
                  strSQL = strSQL & rs.Fields(a).Name & "=" & Replace(Itens.Item(a + 1).ValueStore, ",", ".")
               Case MyDateTime
                  strSQL = strSQL & rs.Fields(a).Name & "=" & IIf(Itens.Item(a + 1).ValueStore = "", "NULL", "'" & Format(Itens.Item(a + 1).ValueStore, "yyyymmdd") & "'")
               Case Else
                  strSQL = strSQL & rs.Fields(a).Name & "='" & Itens.Item(a + 1).ValueStore & "'"
            End Select
            JaInseriu = True
         End If
      End If
   Next
   strSQL = strSQL & " Where " + """" + rs(0).Name + """" + " in ('" & object_id_ & "')"
   
   
   End If
   
   
   
   If JaInseriu Then conn.Execute strSQL
   rs.Close
   Set rs = Nothing
    '##### Salva Especificos
   If Itens.Item(2).Changed_ Then
      For a = 1 To Itens.Count
         If Itens.Item(a).Specific_ And Itens.Item(a).Changed_ Then
            getPMIAS layerName, IIf(object_id_ = "", 0, object_id_), Itens.Item(a).Type_, Itens.Item(a).SubType, Replace(Itens.Item(a).ValueStore, ",", "."), conn
         End If
      Next
   End If
   Exit Function
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Public Function SaveMultProperties - " & strSQL & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
   
End Function

Public Function SaveProperties(object_id_ As String, USER_LOG As String) As Boolean
On Error GoTo Trata_Erro
   'On Error GoTo SavePropriets_Error
   Dim a As Integer, str As String
   '##### Salva Padrao

   For a = 1 To Itens.Count

      If Not Itens.Item(a).Specific_ Then
         If a = 1 Then
            str = IIf(object_id_ = "", 0, object_id_)
         Else
            If Itens.Item(a).Name_ = "USUÁRIO" Then
                str = str & ",'" & USER_LOG & "'"
            
            ElseIf Itens.Item(a).Name_ = "[DATA CADASTRO]" Or Itens.Item(a).Name_ = "DATA CADASTRO" Then
                str = str & ",'" & Format(Now, "DD/MM/YY") & " " & Format(Time, "HH:MM") & "'"
                
                'completanto para o DATALOG
                'str = str & ",'" & Format(Now, "DD/MM/YY") & " " & Format(Time, "HH:MM") & "'"
            ElseIf Itens.Item(a).Name_ <> "DATA_DE_INSTALAÇÃO" Then

               Select Case Itens.Item(a).DataType

                  Case MyChar, MynChar, MyVarChar, MynVarChar
                     str = str & IIf(Itens.Item(a).ValueStore = "", ",''", ",'" & Replace(Itens.Item(a).ValueStore, ",", ".") & "'")

                  Case MyDateTime
                     
                     str = str & IIf(Itens.Item(a).ValueStore = "", ",NULL", ",'" & Format(Itens.Item(a).ValueStore, "yyyymmdd") & "'")

                  Case Else
                     str = str & IIf(Itens.Item(a).ValueStore = "", ",0", "," & Replace(Itens.Item(a).ValueStore, ",", "."))

               End Select
            Else

               If IsNull(Itens.Item(a).ValueStore) Or IsNumeric(Itens.Item(a).ValueStore) Or Itens.Item(a).ValueStore = "" Then
                  str = str & ",NULL"
               Else
                  str = str & ",'" & Format(Itens.Item(a).ValueStore, "yyyymmdd") & "'"
               End If
            End If
         End If
         
      End If
   Next


'    '**************** MONITORAMENTO ******************
'    Close #1
'    Open App.Path & "\GeoSanLog.txt" For Append As #1
'    Print #1, Now & " - PManager4.DLL - " & layerName & " - " & str
'    Close #1
'    '********************* FIM ***********************

   If str <> "" Then
      conn.Execute (IAD(layerName, str, conn))
   End If
   '##### Salva Especificos
   
Dim aa As Integer

For a = 1 To Itens.Count
'If Itens.Item(a).Specific_ Then
' If Itens.item(a).ValueStore <> item.ItemData(aa) Then


 If Trim(Trim(Itens.Item(a).Name_)) <> "USUÁRIO" And Trim(Itens.Item(a).Name_) <> "INICIAL COMPONENTE" And Trim(Itens.Item(a).Name_) <> "FINAL COMPONENTE" _
 And Trim(Itens.Item(a).Name_) <> "COMPR. CALCULADO" And Trim(Itens.Item(a).Name_) <> "DATA CADASTRO" And Trim(Itens.Item(a).Name_) <> "[DATA_DE_INSTALAÇÃO]" _
 And Trim(Itens.Item(a).Name_) <> "[DATA CADASTRO]" Then
 If Itens.Item(a).ValueStore = "" Then
 Itens.Item(a).ValueStore = 0
 End If
getPMIAS layerName, IIf(object_id_ = "", 0, object_id_), Itens.Item(a).Type_, Itens.Item(a).SubType, Replace(Itens.Item(a).ValueStore, ",", "."), conn
 End If
Next

   SaveProperties = True
   Exit Function

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Public Function SaveProperties - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function




Private Function LoadObjectSelect(Index As Integer) As Boolean
'Cria em Tempo de Execução Objectos para seleção

On Error GoTo Trata_Erro
   
   'On Error Resume Next
   Load LstOpt(Index)
   Load Pic(Index)
   LstOpt(Index).Width = Grid.CellWidth
   LstOpt(Index).Top = Grid.RowPos(Index) + 270 + Grid.Top
   LstOpt(Index).Left = Grid.CellLeft + Grid.Left
   LstOpt(Index).ZOrder 0
   Pic(Index).Top = Grid.RowPos(Index) + 40 + Grid.Top
   Pic(Index).Left = Grid.CellLeft + Grid.CellWidth - 230 + Grid.Left
   Pic(Index).ZOrder 0
   Pic(Index).Visible = True
   'On Error GoTo 0
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Function LoadObjectSelect - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
End Function

Private Sub Grid_DblClick()
On Error GoTo Trata_Erro
   
   If Itens.Item(Grid.Row + 1).Selection_ Then
      VisibledObjectsSelect Grid.Row
   Else
      Grid_EnterCell
   End If

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub Grid_DblClick - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

Private Sub VisibledObjectsSelect(Index As Integer)
On Error GoTo Trata_Erro
   txtInput.Visible = False
   LstOpt(Index).Visible = True
   LstOpt(Index).SetFocus

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub VisibledObjectsSelect - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

'/////////////////////////////////////////////////////////////////////////
'// evento determina quando a celula recebeu o foco e posiciona os elementos
'// de entrada(seleção ou não)

Private Sub Grid_EnterCell()
On Error GoTo Trata_Erro
   'On Error GoTo Grid_EnterCell_Err
   If Not Itens.Item(Grid.Row + 1).Selection_ And Itens.Item(Grid.Row + 1).Enabled_ Then
      PicVisibled Grid.Row
      txtInput.Text = Grid.Text
      txtInput.Top = Grid.CellTop + Grid.Top
      txtInput.Left = Grid.CellLeft + Grid.Left
      txtInput.Visible = True
      txtInput.SetFocus
      txtInput.SelStart = 0
      txtInput.SelLength = Len(txtInput.Text)
   Else
      PicVisibled Grid.Row
      txtInput.Visible = False
   End If
   Exit Sub
'Grid_EnterCell_Err:

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub Grid_EnterCell - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

'/////////////////////////////////////////////////////////////////////////
'// evento Ao Precionar teclas
Private Sub Grid_KeyPress(KeyAscii As Integer)
On Error GoTo Trata_Erro
   
   If KeyAscii = vbKeyReturn Then Grid_DblClick
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub Grid_KeyPress - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
   
End Sub

''/////////////////////////////////////////////////////////////////////////
''// evento ao click na lista seleciona o item para grid
'Private Sub LstOpt_Click(Index As Integer)
'On Error GoTo Trata_Erro
'
'   If Foco Then
'      Grid.TextMatrix(Index, 1) = LstOpt(Index).Text
'      Itens.Item(Index + 1).ValueStore = LstOpt(Index).ItemData(LstOpt(Index).ListIndex)
'      Itens.Item(Index + 1).Changed_ = True
'      If Index = 1 Then
'         LoadSpecificProperties layerName, Itens.Item(1).ValueStore, Itens.Item(2).ValueStore
'      End If
'      LstOpt(Index).Visible = False
'   End If
'   Foco = True
'
'Trata_Erro:
'   If Err.Number = 0 Or Err.Number = 20 Then
'      Resume Next
'   Else
'      Open App.Path & "\GeoSanLog.txt" For Append As #1
'      Print #1, Now & " - PManager4.DLL - Manager - Private Sub LstOpt_Click - " & Err.Number & " - " & Err.Description
'      Close #1
'      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
'   End If
'
'End Sub

'/////////////////////////////////////////////////////////////////////////
'// evento ao click na lista seleciona o item para grid

Private Sub LstOpt_Click(Index As Integer)
On Error GoTo Trata_Erro
Dim rs As ADODB.Recordset
Dim strSQL As String
  Dim af As String
         Dim ag As String
   If Foco Then
      Grid.TextMatrix(Index, 1) = LstOpt(Index).Text
      Itens.Item(Index + 1).ValueStore = LstOpt(Index).ItemData(LstOpt(Index).ListIndex)
      Itens.Item(Index + 1).Changed_ = True
      If Index = 1 Then
  
         'VERIFICA SE HÁ ITENS DE SELECTIONS E SUBTYPES (GERAM NOVOS CAMPOS PARA SELEÇÃO) PARA O LAYER SELECIONADO
          If tipoConex <> 4 Then '
         strSQL = "SELECT * FROM " & layerName & "SELECTIONS"
         Set rs = conn.Execute(strSQL)
         If rs.EOF = False Then
            strSQL = "SELECT * FROM " & layerName & "SUBTYPES"
            Set rs = conn.Execute(strSQL)
            If rs.EOF = False Then
               
               LoadSpecificProperties layerName, Itens.Item(1).ValueStore, Itens.Item(2).ValueStore
               
            End If
         End If
         
         Else
       
         af = "SELECTIONS"
         ag = "SUBTYPES"
         strSQL = "SELECT * FROM " + """" + layerName + af + """"
         Set rs = conn.Execute(strSQL)
         If rs.EOF = False Then
            strSQL = "SELECT * FROM " + """" + layerName + ag + """"
            Set rs = conn.Execute(strSQL)
            If rs.EOF = False Then
               
               LoadSpecificProperties layerName, Itens.Item(1).ValueStore, Itens.Item(2).ValueStore
               
            End If
         End If
         End If
         
      End If
      LstOpt(Index).Visible = False
   End If
   Foco = True
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub LstOpt_Click - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
End Sub


'/////////////////////////////////////////////////////////////////////////
'// evento precionar a teclas entrada, e movimento
Private Sub LstOpt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Trata_Erro
   If KeyCode = vbKeyReturn Then
      Foco = True
      LstOpt_Click Index
   ElseIf KeyCode = vbKeyEscape Then
      Foco = False
      LstOpt(Index).Visible = False
      Grid.SetFocus
   Else
      Foco = False
   End If

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub LstOpt_KeyDown - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

'/////////////////////////////////////////////////////////////////////////
'// evento click da picture, abre a lista para seleção
Private Sub Pic_Click(Index As Integer)
On Error GoTo Trata_Erro
   PicVisibled Index
   If LstOpt(Index).Visible Then
      LstOpt(Index).Visible = False
   Else
      Foco = False
      LstOpt(Index).Visible = True
      LstOpt(Index).Text = Grid.TextMatrix(Index, 1)
      LstOpt(Index).ZOrder 0
      LstOpt(Index).SetFocus
      Foco = True
      If LstOpt(Index).Top + LstOpt(Index).Height > Grid.Height Then
         LstOpt(Index).Height = LstOpt(Index).Height - (LstOpt(Index).Top + LstOpt(Index).Height - Grid.Height)
      Else
         LstOpt(Index).Height = 1425
      End If
   End If

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub Pic_Click - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

'/////////////////////////////////////////////////////////////////////////
'// Procedimento que habilita a visualização da list referente
Private Sub PicVisibled(Index As Integer)
On Error GoTo Trata_Erro
   
   Dim a As Integer
   For a = 0 To Itens.Count - 1
      If a <> Index Then
         If Itens.Item(a + 1).Selection_ Then
            LstOpt(a).Visible = False
         End If
      End If
   Next
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub PicVisibled - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Trata_Erro
   Select Case KeyCode
      Case vbKeyDown
         UpdGrid txtInput.Text
         If Grid.Row < Grid.Rows - 1 Then
            Grid.Row = Grid.Row + 1
         End If
         Foco = True
      Case vbKeyUp
         UpdGrid txtInput.Text
         If Grid.Row > 0 Then
            Grid.Row = Grid.Row - 1
         End If
         Foco = True
      Case vbKeyReturn
         UpdGrid txtInput.Text
         If Grid.Row < Grid.Rows - 1 Then
            Grid.Row = Grid.Row + 1
         End If
         Foco = True
      Case vbKeyEscape
         txtInput.Text = Grid.Text
   End Select

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub txtInput_KeyDown - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

Private Sub txtInput_GotFocus()
On Error GoTo Trata_Erro
   
   txtInput.SelStart = 0
   txtInput.SelLength = Len(txtInput.Text)

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub txtInput_GotFocus - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
End Sub

Private Function ValidateTypeData(myData As Variant, ByRef tipo As Integer, ByRef linha As Integer) As Boolean
On Error GoTo Trata_Erro
   linha = Grid.Row + 1
   tipo = Itens.Item(Grid.Row + 1).DataType
   Select Case Itens.Item(Grid.Row + 1).DataType
      Case MyBit
         If IsInteger(myData) Then
            If (myData >= 0 And myData <= 1) Then ValidateTypeData = True
         End If
      Case MyTinyInt
         If IsInteger(myData) Then
            If (myData >= 0 And myData <= 255) Then ValidateTypeData = True
         End If
      Case MySmallInt
         If IsInteger(myData) Then
            If (myData >= 0 And myData <= 32768) Then ValidateTypeData = True
         End If
      Case MyInt, MyNumber
         If IsInteger(myData) Then
            If (myData >= 0 And myData <= 2147483648#) Then ValidateTypeData = True
         End If
      Case MyDecimal, MyFloat, MyNumberScale
         If IsNumeric(myData) Then ValidateTypeData = True
      Case MyChar, MynChar, MynVarChar, MyVarChar
         ValidateTypeData = True
      Case MyDateTime
         If IsDate(myData) Or myData = "" Then ValidateTypeData = True
         Case 0
          If IsInteger(myData) Then
            If (myData >= 0 And myData <= 2147483648#) Then ValidateTypeData = True
         End If
          Case 4
          If IsInteger(myData) Then
            If (myData >= 0 And myData <= 2147483648#) Then ValidateTypeData = True
         End If
   End Select

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Function ValidateTypeData - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function

Private Function IsInteger(myData As Variant) As Boolean
On Error GoTo Trata_Erro
   
   If IsNumeric(myData) Then
      If InStr(1, myData, ".") = 0 And InStr(1, myData, ",") = 0 Then IsInteger = True
   End If

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Function IsInteger - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
End Function

Private Sub UpdGrid(myData)
On Error GoTo Trata_Erro

   Dim Rtn As Integer, tipo As Integer, linha As Integer
   Rtn = ValidateTypeData(myData, tipo, linha)
   If Rtn Then
      Grid.Text = myData
      Itens.Item(Grid.Row + 1).ValueStore = myData
      Itens.Item(Grid.Row + 1).Changed_ = True
   Else
      MsgBox "Tipo de dado não permitido o campo:" & Itens.Item(linha).Name_ & vbCrLf & " Tipo de dado identificado: " & tipo, vbExclamation
   End If
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub UpdGrid - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
   
End Sub

Public Sub GridEnabled(Enabled_ As Boolean)
On Error GoTo Trata_Erro
   
   Dim a As Integer
   Grid.Enabled = Enabled_
   txtInput.Visible = Enabled_
   For a = 1 To Itens.Count
      If Itens.Item(a).Selection_ Then
         Pic(a - 1).Enabled = Enabled_
      End If
   Next
   Exit Sub

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Public Sub GridEnabled - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

Public Sub GridVisibled(Visibled_ As Boolean)
   On Error GoTo Trata_Erro
   
   Dim a As Integer
   Grid.Visible = Visibled_
   txtInput.Visible = False
   For a = 1 To Itens.Count
      If Itens.Item(a).Selection_ Then
         Pic(a - 1).Visible = Visibled_
      End If
   Next
   Exit Sub
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Public Sub GridVisibled - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

'Muda o tamanho da lateral com as propriedades das redes
'
'
'
Public Sub Resize(mwith As Double, mheight As Double)
    On Error GoTo Trata_Erro
    Dim a As Integer
    Dim intMomento As Integer
    Dim strMomento As Integer
    
    Grid.AllowUserResizing = flexResizeBoth
    If mwith > 14000 Or mwith < 0 Then Exit Sub
    If mheight > 14000 Or mheight < 0 Then Exit Sub
    Grid.ScrollBars = flexScrollBarNone
    Grid.Width = mwith
    Grid.Height = mheight
    Grid.ColWidth(0) = mwith / 100 * 50
    If Grid.Height \ Grid.RowHeight(0) = Grid.Rows And Grid.Height Mod Grid.RowHeight(0) >= 75 Then
        Grid.ColWidth(1) = (mwith) / 100 * 48.5
    ElseIf Grid.Height \ Grid.RowHeight(0) <= Grid.Rows Then
        Grid.ColWidth(1) = (mwith - 800) / 100 * 48.5
    Else
    Grid.ColWidth(1) = mwith / 100 * 48.5
    End If
    txtInput.Width = Grid.ColWidth(1)
    If Itens.Count > 0 Then
        For a = (Grid.Rows - 1) To 0 Step -1
            If Itens.Item(a + 1).Selection_ Then
                Grid.Row = a
                Pic(a).Left = Grid.CellLeft + Grid.CellWidth - Pic(a).Width + Grid.Left
                Pic(a).Top = 242 * (a - Grid.TopRow) + 20 'Grid.CellTop + Grid.Top
                LstOpt(a).Width = Grid.CellWidth
                LstOpt(a).Left = Grid.CellLeft + Grid.Left
                LstOpt(a).Top = 242 * (a - (Grid.TopRow - 1)) 'Grid.RowPos(A) + 270 + Grid.Top
            End If
        Next
    End If
    Grid.ScrollBars = flexScrollBarBoth
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Grid.ScrollBars = flexScrollBarBoth
        Open App.Path & "\GeoSanLog.txt" For Append As #1
        Print #1, Now & " - PManager4.DLL - Manager - Public Sub Resize - " & intMomento & " - " & strMomento & " - " & Err.Number & " - " & Err.Description
        Close #1
        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
    End If
End Sub

Private Sub Grid_Scroll()
On Error GoTo Trata_Erro
   With Grid
      Resize .Width, .Height
   End With

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Grid.ScrollBars = flexScrollBarBoth
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - Manager - Private Sub Grid_Scroll - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

