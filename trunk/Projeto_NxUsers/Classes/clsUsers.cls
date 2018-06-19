VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsers"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum mOperation
   mSelect = 1
   mUpdate = 2
   mInsert = 3
   mDelete = 4
End Enum
 Dim tipoDeConexao As String
Private mvarUsrId As Long 'local copy
Private mvarUsrLog As String 'local copy
Private mvarUsrNom As String 'local copy
Private mvarUsrFun As Long 'local copy
Private mvarUsrDep As Long 'local copy
Private mvarUsrPwd As String 'local copy
Private mvarUsrExp As Boolean 'local copy
Private mvarUsrBrk As Boolean 'local copy
Private mvarDepartaments As clsDepartaments
Private mvarGroups As clsGroups
Private mvarResources As clsResources
Private mvarResourcesGroups As clsResourcesGroups
Private mvarUsers As clsUsers
Private mvarUsersFunctions As clsUsers
Private mvarUsersGroups As clsUsersGroups

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------




Public Property Let UsersGroups(ByVal vData As clsUsersGroups)
    mvarUsersGroups = vData
End Property

Public Property Get UsersGroups() As clsUsersGroups
  If mvarUsersGroups Is Nothing Then
     Set mvarUsersGroups = New clsUsersGroups
  End If
  Set UsersGroups = mvarUsersGroups
End Property

Public Property Let UsersFunctions(ByVal vData As clsUsersFunctions)
    mvarUsersFunctions = vData
End Property

Public Property Get UsersFunctions() As clsUsersFunctions
  If mvarUsersFunctions Is Nothing Then
     Set mvarUsersFunctions = New clsUsersFunctions
  End If
  Set UsersFunctions = mvarUsersFunctions
End Property

Public Property Let Users(ByVal vData As clsUsers)
    mvarUsers = vData
End Property

Public Property Get Users() As clsUsers
  If mvarUsers Is Nothing Then
     Set mvarUsers = New clsUsers
  End If
  Set Users = mvarUsers
End Property

Public Property Let Departaments(ByVal vData As clsDepartaments)
    mvarDepartaments = vData
End Property

Public Property Get Departaments() As clsDepartaments
  If mvarDepartaments Is Nothing Then
     Set mvarDepartaments = New clsDepartaments
  End If
  Set Departaments = mvarDepartaments
End Property

Public Property Let Groups(ByVal vData As clsGroups)
    mvarGroups = vData
End Property

Public Property Get Groups() As clsGroups
  If mvarGroups Is Nothing Then
     Set mvarGroups = New clsGroups
  End If
  Set Groups = mvarGroups
End Property

Public Property Let Resources(ByVal vData As clsResources)
    mvarResources = vData
End Property

Public Property Get Resources() As clsResources
  If mvarResources Is Nothing Then
     Set mvarResources = New clsResources
  End If
  Set Resources = mvarResources
End Property

Public Property Let ResourcesGroups(ByVal vData As clsResourcesGroups)
    mvarResourcesGroups = vData
End Property

Public Property Get ResourcesGroups() As clsResourcesGroups
  If mvarResourcesGroups Is Nothing Then
     Set mvarResourcesGroups = New clsResourcesGroups
  End If
  Set ResourcesGroups = mvarResourcesGroups
End Property

Public Property Let UsrBrk(ByVal vData As Boolean)
    mvarUsrBrk = vData
End Property

Public Property Get UsrBrk() As Boolean
    UsrBrk = mvarUsrBrk
End Property

Public Property Let UsrExp(ByVal vData As Boolean)
    mvarUsrExp = vData
End Property

Public Property Get UsrExp() As Boolean
    UsrExp = mvarUsrExp
End Property

Public Property Let UsrPwd(ByVal vData As String)
    mvarUsrPwd = vData
End Property

Public Property Get UsrPwd() As String
    UsrPwd = mvarUsrPwd
End Property

Public Property Let UsrDep(ByVal vData As Long)
    mvarUsrDep = vData
End Property

Public Property Get UsrDep() As Long
    UsrDep = mvarUsrDep
End Property

Public Property Let UsrFun(ByVal vData As Long)
    mvarUsrFun = vData
End Property

Public Property Get UsrFun() As Long
    UsrFun = mvarUsrFun
End Property

Public Property Let UsrNom(ByVal vData As String)
    mvarUsrNom = vData
End Property

Public Property Get UsrNom() As String
    UsrNom = mvarUsrNom
End Property

Public Property Let UsrLog(ByVal vData As String)
    mvarUsrLog = vData
End Property

Public Property Get UsrLog() As String
    UsrLog = mvarUsrLog
End Property

Public Property Let Usrid(ByVal vData As Long)
    mvarUsrId = vData
End Property

Public Property Get Usrid() As Long
    Usrid = mvarUsrId
End Property

Public Function UpdateData(Conn As ADODB.Connection) As Boolean
   On Error GoTo UpdateData_error
   
Dim a As String
Dim b As String
Dim c As String
Dim d As String


   Dim MyRs As ADODB.Recordset 'local copy
   Set MyRs = New ADODB.Recordset
   If Form1.TipoConexao <> 4 Then
   
   MyRs.Open "Select * from  SystemUsers " & _
             "Where UsrId = " & Usrid, Conn, _
               adOpenKeyset, adLockOptimistic, adCmdText
               Else
a = "SYSTEMUSERS"
b = "USRID"

               MyRs.Open "Select * from  " + """" + a + """" + _
             "Where " + """" + b + """" + " = '" & Usrid & "' ", Conn, _
              adOpenDynamic, adLockOptimistic
               End If
               
   If Not MyRs.EOF Then
      MyRs.Fields("UsrBrk").Value = UsrBrk
      MyRs.Fields("UsrDep").Value = UsrDep
      MyRs.Fields("UsrExp").Value = UsrExp
      MyRs.Fields("UsrFun").Value = UsrFun
      MyRs.Fields("UsrNom").Value = UsrNom
      MyRs.Fields("UsrPwd").Value = UsrPwd
      MyRs.Fields("UsrLog").Value = UsrLog
      MyRs.Update
      UpdateData = True
   End If
   CloseObjectRs MyRs
   Set MyRs = Nothing
UpdateData_error:
   CloseObjectRs MyRs
   Set MyRs = Nothing
End Function

Public Function SelectData(Conn As ADODB.Connection, UserId As Long) As Boolean
   On Error GoTo SelectData_error
   Dim MyRs As ADODB.Recordset 'local copy
       Dim stringconexao As String

    Dim a As String
    Dim b As String
    Dim c As String
     Dim d As String
     Dim e As String
    
      
If Form1.TipoConexao <> 4 Then
      stringconexao = "Select * from  SystemUsers C Where C.UsrId = " & UserId
   Else
   a = "USRID"
      c = "SYSTEMUSERS"
   stringconexao = "Select * from  " + """" + c + """" + "Where " + """" + c + """" + "." + """" + a + """" + " = '" & UserId & "'"
  
   End If
   
   Set MyRs = Conn.Execute(stringconexao)
   If Not MyRs.EOF Then
      Usrid = MyRs.Fields("UsrID").Value
     ' UsrBrk = MyRs.Fields("UsrBrk").Value
      UsrDep = MyRs.Fields("UsrDep").Value
      'UsrExp = MyRs.Fields("UsrExp").Value
      UsrFun = MyRs.Fields("UsrFun").Value
      UsrNom = MyRs.Fields("UsrNom").Value
      UsrPwd = MyRs.Fields("UsrPwd").Value
      UsrLog = MyRs.Fields("UsrLog").Value
      SelectData = True
   End If
   CloseObjectRs MyRs
   Set MyRs = Nothing
SelectData_error:
   CloseObjectRs MyRs
   Set MyRs = Nothing
End Function

Public Function InsertData(Conn As ADODB.Connection) As Long
   On Error GoTo InsertData_error
   Dim MyRs As ADODB.Recordset 'local copy
   Set MyRs = New ADODB.Recordset
   MyRs.Open "SystemUsers", Conn, adOpenDynamic, adLockOptimistic
   MyRs.AddNew
   MyRs.Fields("UsrBrk").Value = UsrBrk
   MyRs.Fields("UsrDep").Value = UsrDep
   MyRs.Fields("UsrExp").Value = UsrExp
   MyRs.Fields("UsrFun").Value = UsrFun
   MyRs.Fields("UsrNom").Value = UsrNom
   MyRs.Fields("UsrPwd").Value = UsrPwd
   MyRs.Fields("UsrLog").Value = UsrLog
   MyRs.Update
   InsertData = MyRs.Fields("UsrID").Value
   CloseObjectRs MyRs
   Set MyRs = Nothing
InsertData_error:
   CloseObjectRs MyRs
   Set MyRs = Nothing
End Function

Public Function SelectAllUsers(Conn As ADODB.Connection) As ADODB.Recordset
   Dim rs As ADODB.Recordset
   
   
   Dim stringconexao As String
    Dim a As String
    Dim b As String
    Dim c As String
     Dim d As String
     Dim e As String
    
      

   If Form1.TipoConexao <> 4 Then
   stringconexao = "select * from SystemUsers Order By UsrNom"
   Else
   a = "SYSTEMUSERS"
      c = "USRNOM"
   stringconexao = "select *  from  " + """" + a + """" + "ORDER BY " + """" + c + """" + ""
   End If
   
    Set rs = Conn.Execute(stringconexao)
   

   Set SelectAllUsers = rs
   Set rs = Nothing
End Function

Private Sub CloseObjectRs(ObjRS As ADODB.Recordset)
   If Not (ObjRS Is Nothing) Then
      If ObjRS.State = adStateOpen Then ObjRS.Close
   End If
End Sub

Public Function DeleteUserGroups(Conn As ADODB.Connection) As Boolean
Dim stringconexao As String

    Dim a As String
    Dim b As String
    Dim c As String
     Dim d As String
     Dim e As String
    

   If Form1.TipoConexao <> 4 Then
   stringconexao = "Delete From SystemUsersGroups Where UsrID=" & Usrid
   Else
   a = "SYSTEMUSERSGROUPS"
      c = "USERID"
      
   stringconexao = "Delete  from  " + """" + a + """" + "Where " + """" + c + """" + " = '" & Usrid & "'"
   End If
   
   Conn.Execute (stringconexao)
   DeleteUserGroups = True
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

Public Function FindUser(Conn As ADODB.Connection, UserName As String) As Boolean
   On Error GoTo SelectData_error
   Dim MyRs As ADODB.Recordset 'local copy
   Dim stringconexao As String

    Dim a As String
    Dim b As String
    Dim c As String
     Dim d As String
     Dim e As String
    a = "USRLOG"
      c = "SYSTEMUSERS"
      
      
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
 If Form1.TipoConexao = 4 Then
mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.Path & "\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.Path & "\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.Path & "\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.Path & "\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.Path & "\GEOSAN.ini")
decriptada = FunDecripta(Senha)

  strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + decriptada + "; ByteaAsLongVarBinary=1;"

    conexao.Open strConn
      
      End If
      
      
      
      

      If Form1.TipoConexao <> 4 Then
   stringconexao = "Select * from  SystemUsers C Where C.UsrLog = '" & Trim(UserName) & "'"
 Else
   stringconexao = "Select * from  " + """" + c + """" + "Where " + """" + a + """" + " = '" & Trim(UserName) & "'"
   
   
   
   
   'MsgBox stringconexao
 
  
   End If
   
 If Form1.TipoConexao <> 4 Then
   Set MyRs = Conn.Execute(stringconexao)
   If Not MyRs.EOF Then
      Usrid = MyRs.Fields("USRID").Value
      UsrBrk = MyRs.Fields("USRBRK").Value
      UsrDep = MyRs.Fields("USRDEP").Value
      UsrExp = MyRs.Fields("USREXP").Value
      UsrFun = MyRs.Fields("USRFUN").Value
      UsrNom = MyRs.Fields("USRNOM").Value
      UsrPwd = MyRs.Fields("USRPWD").Value
      UsrLog = MyRs.Fields("USRLOG").Value
      FindUser = True
   End If
   Else
    Set MyRs = conexao.Execute(stringconexao)
   If Not MyRs.EOF Then
      Usrid = MyRs.Fields("USRID").Value
      UsrBrk = MyRs.Fields("USRBRK").Value
      UsrDep = MyRs.Fields("USRDEP").Value
      UsrExp = MyRs.Fields("USREXP").Value
      UsrFun = MyRs.Fields("USRFUN").Value
      UsrNom = MyRs.Fields("USRNOM").Value
      UsrPwd = MyRs.Fields("USRPWD").Value
      UsrLog = MyRs.Fields("USRLOG").Value
      FindUser = True
   End If
   
   
   End If
   CloseObjectRs MyRs
   Set MyRs = Nothing
SelectData_error:
   CloseObjectRs MyRs
   Set MyRs = Nothing
End Function


Public Static Function TipoConexao(tipo As String) As String
tipoDeConexao = tipo
Form1.TipoConexao2 (tipoDeConexao)


End Function



Public Function GetPermission(Conn As ADODB.Connection, Usrid As Long, RcsNom As String, Operacao As mOperation, Optional EnibirMsg As Boolean) As Boolean
   Dim mUpd As Boolean, mView As Boolean, mDel As Boolean, mIns As Boolean
   Dim RcsID As Integer, PermissionOK As Boolean
   Dim mGroups As String

  ' Form1.SetaTipoConexao (Conn.Provider)
   
   If UsersGroups.SelectGroupByUsers(Conn, Usrid, mGroups) Then
      If Resources.SelectResourcesByName(Conn, RcsNom, RcsID) Then
         If ResourcesGroups.SelectResourcesByGroupII(Conn, mGroups, RcsID, mUpd, mView, mDel, mIns) Then
            Select Case Operacao
               Case mSelect
                  PermissionOK = mView
               Case mUpdate
                  PermissionOK = mUpd
               Case mInsert
                  PermissionOK = mIns
               Case mDelete
                  PermissionOK = mDel
            End Select
            If PermissionOK Then
               GetPermission = True
            Else
               If Not EnibirMsg Then MsgBox "Usuário sem permissão", vbExclamation
            End If
         End If
      End If
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


