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
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClsAttributes"
Attribute VB_Ext_KEY = "Member0" ,"ClsAttributes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(Enabled_ As Boolean, Type_ As Long, SubType As Long, Selection_ As Boolean, Max_ As Long, Min_ As Long, DataType As Integer, Name_ As String, ValueDisplay As String, ValueStore As String, Optional Specific_ As Boolean, Optional Changed_ As Boolean, Optional sKey As String) As ClsAttributes
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
    objNewMember.ValueStore = Trim(ValueStore)
    
    objNewMember.Changed_ = Changed_
    objNewMember.Specific_ = Specific_
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As ClsAttributes
  On Error GoTo sair_item
  Set Item = mCol(vntIndexKey)
  Exit Sub
sair_item:
  Resume Next
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
   Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub



