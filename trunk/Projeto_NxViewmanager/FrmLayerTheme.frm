VERSION 5.00
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TECOMC~1.DLL"
Object = "{F03ABD98-7B60-43E4-9934-DA5F0D19FDAC}#1.0#0"; "TeComViewManager.dll"
Object = "{EE78E37B-39BE-42FA-80B7-E525529739F7}#1.0#0"; "TECOMV~2.DLL"
Begin VB.Form FrmLayerTheme 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecione o  plano e informe o nome do novo tema"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   210
      TabIndex        =   5
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   315
      Left            =   3930
      TabIndex        =   4
      Top             =   1260
      Width           =   795
   End
   Begin VB.TextBox txtThemeName 
      Height          =   345
      Left            =   1050
      TabIndex        =   1
      Top             =   810
      Width           =   3675
   End
   Begin VB.ComboBox cboLayer 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3675
   End
   Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection2 
      Left            =   840
      OleObjectBlob   =   "FrmLayerTheme.frx":0000
      Top             =   600
   End
   Begin TeComViewDatabaseLibCtl.TeViewDatabase TeViewDatabase2 
      Left            =   240
      OleObjectBlob   =   "FrmLayerTheme.frx":0024
      Top             =   120
   End
   Begin TECOMVIEWMANAGERLibCtl.TeViewManager TeViewManager2 
      Left            =   720
      OleObjectBlob   =   "FrmLayerTheme.frx":0048
      Top             =   240
   End
   Begin TeComViewDatabaseLibCtl.TeViewDatabase TeViewDatabase1 
      Left            =   1920
      OleObjectBlob   =   "FrmLayerTheme.frx":006C
      Top             =   0
   End
   Begin TECOMVIEWMANAGERLibCtl.TeViewManager TeViewManager1 
      Left            =   3840
      OleObjectBlob   =   "FrmLayerTheme.frx":0090
      Top             =   120
   End
   Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
      Left            =   2880
      OleObjectBlob   =   "FrmLayerTheme.frx":00B4
      Top             =   120
   End
   Begin VB.Label Label2 
      Caption         =   "Tema:"
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Plano:"
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   270
      Width           =   1095
   End
End
Attribute VB_Name = "FrmLayerTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Confirm As Boolean
Dim aa As String
Dim b As String
Dim c As String
Dim d As String
Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
Dim tema As String
Dim conexao As New ADODB.Connection
Dim usuario As String
Dim strConn As String
Dim carrega As Integer
Dim layer2, theme2 As String
Dim manager As Object
Dim database As Object



Public Function Init(mtvw As Object, LayerName As String, ThemeName As String) As Boolean
   'LoozeXP1.InitSubClassing
   Confirm = False
   Dim a As Integer, rs As ADODB.Recordset
   
  
   

   
  aa = "name"
b = "te_layer"

   If TypeConn <> 4 Then
   Set rs = conn.Execute("Select name from te_layer order by name ")
   Else
   Set rs = conn.Execute("Select " + """" + aa + """" + " from " + """" + b + """" + " order by " + """" + aa + """" + "")
   End If
   
   
'   For A = 0 To mtvw.getLayerCount() - 1
'       cboLayer.AddItem mtvw.getLayerName(A)
'   Next
   cboLayer.Clear ' adicionado por Jonathas em 23/09/2008, elimina os anteriores ao recarregar
   While Not rs.EOF
      cboLayer.AddItem rs(0).Value
      rs.MoveNext
   Wend
   rs.Close
   Set rs = Nothing
   

   
   
   If cboLayer.ListCount > -1 Then cboLayer.ListIndex = 0
   txtThemeName = cboLayer.Text
   Me.Show vbModal
   LayerName = cboLayer.Text
   ThemeName = txtThemeName.Text
   layer2 = LayerName
   theme2 = ThemeName

   
   Init = Confirm
End Function







Public Function Temas(theme As String) As String

theme2 = theme
If carrega <> 10 Then
   If TypeConn = 4 Then

  
   mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.Path & "\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.Path & "\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.Path & "\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.Path & "\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.Path & "\GEOSAN.ini")
usuario = ReadINI("CONEXAO", "USER", App.Path & "\GEOSAN.ini")
decriptada = FunDecripta(Senha)
 strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + decriptada + "; ByteaAsLongVarBinary=1;"

    conexao.Open strConn

 TeAcXConnection1.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA
 
  
 TeViewManager1.UserName = usuario
 TeViewManager1.Provider = 4
 TeViewManager1.Connection = TeAcXConnection1.objectConnection_

  

 
   
 TeViewDatabase1.UserName = usuario
 TeViewDatabase1.Provider = 4
 TeViewDatabase1.Connection = TeAcXConnection1.objectConnection_

 ' TeViewManager1.Start
 carrega = 10
 End If
  End If

'TeViewManager1.Start


 ' TeViewManager1.Start
  ' TeViewManager1.saveAsLastView TeViewDatabase1.getActiveView
   'MsgBox TeViewManager1.visibleThemeStatus(theme2)
  ' TeViewManager1.Start
  
  ' tema = TeViewManager1.visibleThemeStatus(theme2)
   
Temas = tema



End Function


Public Function Temas2() As String

'Temas2 = tema

End Function


Private Sub cboLayer_Click()
   txtThemeName.Text = cboLayer.Text
End Sub

Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   Confirm = True
   Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'LoozeXP1.EndWinXPCSubClassing

End Sub

Private Sub txtThemeName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If cboLayer.ListIndex <> -1 Then
            If Trim(txtThemeName.Text) <> "" Then
                cmdOK_Click
            End If
        End If
    End If

End Sub
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
