VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TeComDatabase.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   660
      Left            =   3450
      TabIndex        =   0
      Top             =   1635
      Width           =   1440
   End
   Begin TECOMDATABASELibCtl.TeDatabase TeDatabase1 
      Left            =   855
      OleObjectBlob   =   "Form1.frx":0000
      Top             =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'
'
'    'variavel de conexao do banco principal
'
'    'Dim dbConn As New ADODB.Connection
'
'    'Variavel Contador
'
'    Dim iCont As Integer
'
'    'Variavel de coordenadas
'
'    Dim x As Double, y As Double
'
'
'
'
'    'estabelece a comunicacao com o banco principal
'
'    'dbConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projetos\Documentos\banco.mdb;Persist Security Info=False"
'
'    'configura o provider do banco principal
'
'    TeDatabase1.Provider = 1
'
'    'configura o componente para a conexao com o banco principal
'
'    TeDatabase1.Connection = Conn
'
'    If TeDatabase1.setCurrentLayer("Trechos") Then
'
'                    'Retorna a coordenada do ponto 2 da linha 0xx999
'
'            If TeDatabase1.getPointOfLine(0, "0xx999", 2, x, y) Then
'                MsgBox "Coordenada do Terceiro Ponto da linha : " & CStr(x) & "_ " & CStr(y)
'            End If
'    End If
'
'    'dbConn.Close
'
'    'Set dbConn = Nothing
'
'
'
'
'End Sub
