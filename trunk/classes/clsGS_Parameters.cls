VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGS_Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarString_Connection_Secundary As String 'local copy
Public Property Let String_Connection_Secundary(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.String_Connection_Secundary = 5
    mvarString_Connection_Secundary = vData
End Property


Public Property Get String_Connection_Secundary() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.String_Connection_Secundary
    String_Connection_Secundary = mvarString_Connection_Secundary
End Property


Public Function getData(Conn As ADODB.connection, tipo As String) As Boolean
   Dim rs As ADODB.Recordset
   
   
   Dim stringconexao As String

    Dim a As String
    Dim b As String
    Dim c As String
     Dim d As String
     Dim e As String
    a = "STRING_CONNECTION_SECUNDARY"
      c = "GS_PARAMETERS"
      
If frmCanvas.TipoConexao <> 4 Then
      stringconexao = "SELECT String_Connection_Secundary from GS_PARAMETERS"
   Else
   stringconexao = "Select " + """" + a + """" + " from " + """" + c + """"
   
  
   End If
   
   Set rs = Conn.execute(stringconexao)
    
    If rs.EOF = False Then
    
        If rs.Fields("String_Connection_Secundary").value <> "" Then
            String_Connection_Secundary = rs.Fields("String_Connection_Secundary").value
        End If
    End If
    getData = True
    rs.Close
    Set rs = Nothing
End Function
