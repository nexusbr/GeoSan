VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum TypesOfDatas
   
   MyFloat = 5
   MyChar = 129
   MyVarChar = 200
   MynChar = 130
   MynVarChar = 202
   MyBit = 11
   MyTinyInt = 17
   MySmallInt = 2
   MyInt = 3
   MyDecimal = 13
   MyDateTime = 135
   MyNumber = 131
   MyNumberScale = 1131
   
End Enum



'local variable(s) to hold property value(s)
Private mvarEnabled_ As Boolean 'local copy
Private mvarType_ As Long 'local copy
Private mvarSubType As Long 'local copy
Private mvarSelection_ As Boolean 'local copy
Private mvarMax_ As Long 'local copy
Private mvarMin_ As Long 'local copy
Private mvarDataType As TypesOfDatas 'local copy
Private mvarName_ As String 'local copy
Private mvarValueDisplay As String 'local copy
Private mvarValueStore As String 'local copy
Private mvarChanged_ As Boolean 'local copy
Private mvarSpecific_ As Boolean 'Local copy


Public Property Let ValueStore(ByVal vData As String)
    mvarValueStore = vData
End Property

Public Property Get ValueStore() As String
    ValueStore = mvarValueStore
End Property

Public Property Let ValueDisplay(ByVal vData As String)
    mvarValueDisplay = vData
End Property

Public Property Get ValueDisplay() As String
    ValueDisplay = mvarValueDisplay
End Property

Public Property Let Name_(ByVal vData As String)
    mvarName_ = vData
End Property

Public Property Get Name_() As String
    Name_ = mvarName_
End Property

Public Property Let DataType(ByVal vData As TypesOfDatas)
    mvarDataType = vData
End Property

Public Property Get DataType() As TypesOfDatas
    DataType = mvarDataType
End Property

Public Property Let Min_(ByVal vData As Long)
    mvarMin_ = vData
End Property

Public Property Get Min_() As Long
    Min_ = mvarMin_
End Property

Public Property Let Max_(ByVal vData As Long)
    mvarMax_ = vData
End Property

Public Property Get Max_() As Long
    Max_ = mvarMax_
End Property

Public Property Let Selection_(ByVal vData As Boolean)
    mvarSelection_ = vData
End Property

Public Property Get Selection_() As Boolean
    Selection_ = mvarSelection_
End Property

Public Property Let SubType(ByVal vData As Long)
    mvarSubType = vData
End Property

Public Property Get SubType() As Long
    SubType = mvarSubType
End Property

Public Property Let Type_(ByVal vData As Long)
    mvarType_ = vData
End Property

Public Property Get Type_() As Long
    Type_ = mvarType_
End Property

Public Property Let Enabled_(ByVal vData As Boolean)
    mvarEnabled_ = vData
End Property

Public Property Get Enabled_() As Boolean
    Enabled_ = mvarEnabled_
End Property

Public Property Let Changed_(ByVal vData As Boolean)
    mvarChanged_ = vData
End Property

Public Property Get Changed_() As Boolean
    Changed_ = mvarChanged_
End Property

Public Property Let Specific_(ByVal vData As Boolean)
    mvarSpecific_ = vData
End Property

Public Property Get Specific_() As Boolean
    Specific_ = mvarSpecific_
End Property

