VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTerraConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private rs As ADODB.Recordset
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim contador As Integer

'Configura as coordenadas do box do usuário da última dela que ele deixou quando saiu
'
'
'
Public Function SetWorldByUser(username As String, xmin As Double, ymin As Double, xmax As Double, ymax As Double)
    Dim mPROVEDOR As String
    Dim mSERVIDOR As String
    Dim mPORTA As String
    Dim mBANCO As String
    Dim mUSUARIO As String
    Dim Senha As String
    Dim decriptada As String
    Dim conexao As New ADODB.connection
    Dim strConn As String
    Dim nStr As String
    Dim a As String
    Dim b As String
    Dim c As String
    
    On Error GoTo Trata_Erro

    If username <> "" Then
        a = "TERRACONFIG"
        b = "TEC_USER"
        If contador <> 10 Then
            If frmCanvas.TipoConexao = 4 Then
                mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
                mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
                mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
                mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
                Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
                nStr = frmCanvas.FunDecripta(Senha)
                decriptada = frmCanvas.Senha
                strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"
                conexao.Open strConn
                contador = 10
            End If
        End If
        Set rs = New ADODB.Recordset
        If frmCanvas.TipoConexao <> 4 Then
            rs.Open "SELECT * from TerraConfig where tec_user='" & username & "'", Conn, adOpenKeyset, adLockOptimistic, adCmdText
        Else
            rs.Open "SELECT * from " + """" + a + """" + " where " + """" + b + """" + "='" & username & "'", conexao, adOpenDynamic, adLockOptimistic
        End If
        If rs.EOF Then
            rs.Close
            If frmCanvas.TipoConexao <> 4 Then
                rs.Open "SELECT * FROM TERRACONFIG", Conn, adOpenDynamic, adLockOptimistic
            Else
                rs.Open "SELECT * FROM " + """" + a + """" + ", conexao, adOpenDynamic, adLockOptimistic"
            End If
            rs.AddNew
        End If
        rs.Fields("TEC_USER").value = username
        rs.Fields("TEC_XMIN").value = xmin
        rs.Fields("TEC_YMIN").value = ymin
        rs.Fields("TEC_XMAX").value = xmax
        rs.Fields("TEC_YMAX").value = ymax
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
    Exit Function
    
Trata_Erro:
        If Err.Number = 0 Or Err.Number = 20 Then
            Resume Next
        Else
            PrintErro "clsTerraConfig", "Public Function SetWorldByUser: " + App.path, CStr(Err.Number), CStr(Err.Description), True
        End If

End Function

'Lê as coordenadas do box do usuário, que ele deixou quando saiu pela última vez
'
'
'
Public Function GetWorldByUser(username As String, xmin, ymin, xmax, ymax, tipo As CONNECTION_TYPE) As Boolean
    Dim stringconexao As String
    Dim a As String
    Dim b As String
    Dim c As String
    Dim d As String
    Dim e As String
    On Error GoTo Trata_Erro

    Set rs = New ADODB.Recordset
    If Conn.State = 0 Then
        Conn.Open
    End If
    a = "TERRACONFIG"
    c = "TEC_USER"
    If frmCanvas.TipoConexao <> 4 Then
        stringconexao = "SELECT * from TerraConfig where tec_user='" & strUser & "'"
    Else
        stringconexao = "Select * from  " + """" + a + """" + "Where " + """" + c + """" + " = '" & strUser & "'"
    End If
    rs.Open stringconexao, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        xmin = rs.Fields("tec_xmin").value
        ymin = rs.Fields("tec_ymin").value
        xmax = rs.Fields("tec_xmax").value
        ymax = rs.Fields("tec_ymax").value
        GetWorldByUser = True
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsTerraConfig", "Public Function GetWorldByUser: " + App.path + " - '" + stringconexao + "'", CStr(Err.Number), CStr(Err.Description), True
        Err.Clear
    End If
End Function
