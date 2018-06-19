VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClsAttributes"
Attribute VB_Ext_KEY = "Member0" ,"ClsAttributes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(Enabled_ As Boolean, Type_ As Long, SubType As Long, Selection_ As Boolean, Max_ As Long, Min_ As Long, DataType As Integer, Name_ As String, ValueDisplay As String, ValueStore As String, Optional Specific_ As Boolean, Optional Changed_ As Boolean, Optional sKey As String) As ClsAttributes
On Error GoTo Trata_Erro
    
    'create a new object
    Dim objNewMember As ClsAttributes
    Set objNewMember = New ClsAttributes
    objNewMember.Enabled_ = Enabled_
    objNewMember.Type_ = Type_
    objNewMember.SubType = SubType
    objNewMember.Selection_ = Selection_
    objNewMember.Max_ = Max_
    objNewMember.Min_ = Min_
    objNewMember.DataType = DataType
    objNewMember.Name_ = Name_
    objNewMember.ValueDisplay = ValueDisplay
    objNewMember.ValueStore = ValueStore
    objNewMember.Changed_ = Changed_
    objNewMember.Specific_ = Specific_
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    Set Add = objNewMember
    Set objNewMember = Nothing

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - CItens - Public Function Add - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Function

Public Property Get Item(vntIndexKey As Variant) As ClsAttributes
On Error GoTo Trata_Erro

Set Item = mCol(vntIndexKey)

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - CItens - Public Property Get Item - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Property

Public Property Get Count() As Long
On Error GoTo Trata_Erro

Count = mCol.Count
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - CItens - Public Property Get Count - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
End Property

Public Sub Remove(vntIndexKey As Variant)
On Error GoTo Trata_Erro

   mCol.Remove vntIndexKey
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - CItens - Public Sub Remove - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
End Sub

Public Property Get NewEnum() As IUnknown
On Error GoTo Trata_Erro

   Set NewEnum = mCol.[_NewEnum]
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - CItens - Public Property Get NewEnum - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
End Property

Private Sub Class_Initialize()
On Error GoTo Trata_Erro
    
   Set mCol = New Collection

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - CItens - Private Sub Class_Initialize - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If

End Sub

Private Sub Class_Terminate()
On Error GoTo Trata_Erro

Set mCol = Nothing
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      Open App.Path & "\GeoSanLog.txt" For Append As #1
      Print #1, Now & " - PManager4.DLL - CItens - Private Sub Class_Terminate - " & Err.Number & " - " & Err.Description
      Close #1
      MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   End If
   
End Sub



