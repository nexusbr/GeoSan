VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmAtualiza 
   Caption         =   "Atualiza GeoSan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3480
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmAtualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Referências adicionadas:
'Microsoft Scripting Runtime - srcrun.dll - para saber a versão de uma aplicação
'Microsoft CDO for Windows 2000 Library - cdosys.dll
'Componentes adicionados:
'Microsoft Internet Transfer Control 6.0 - msinet.ocx - para fazer download de arquivos
'
'
'
Dim b() As Byte

Public ErroUsuario As New CPrintErro            'Para gerenciar os erros que por ventura ocorram


Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
    Dim carrega As CDownload
    Dim retorno As Boolean
    Dim versao As CGetVersion                               'gestão das versões de software que deverão ser atualizadas
    Dim numeroVersao As String
    Dim numeroAtualizacoes  As Integer                      'número total de atualizações a serem realizadas
    Dim i As Integer
    Dim nomeArquivo As String                               'nome do arquivo a ser atualizado
    Dim diretorio As String                                 'nome do drive e diretório onde o arquivo será atualizado (salvo)
    Dim versaoNova As String                                'numero da versão nova a ser atualizada
    
    'faz as configurações iniciais
    Set carrega = New CDownload
    Set versao = New CGetVersion
    numeroVersao = versao.ObtemVersao("D:\Desenv\GEOSAN_VB6_B\trunk\Projeto_AtualizaApp\geosanini.exe")
    carrega.diretorioServidor = "/download/GeoSan"
    carrega.url = "http://www.nexusbr.com"
    carrega.proxyPorta = "80"
    carrega.proxy = "NAO EXISTE"
    carrega.diretorioLocal = "c:\tempApp"
    
    Me.Show
    Screen.MousePointer = vbHourglass
    lblStatus.Caption = "Realizando download de atualizações ..."
    
    retorno = carrega.DownloadArquivo("Updates.txt")        'obtem a lista de atualizações disponíveis
    lblStatus.Caption = "Download completo!"
    numeroAtualizacoes = versao.VerificaAtualizacoes("c:\tempApp\Updates.txt")
    For i = 0 To numeroAtualizacoes - 1
        versao.SplitAtualizacoes i, nomeArquivo, diretorio, versaoNova
        retorno = carrega.DownloadArquivo(nomeArquivo)      'faz o download para o diretório local, da atualização
    Next
    
    
    Screen.MousePointer = vbDefault
    
    
    Dim MyVer As String
    MyVer = App.Major & "." & App.Minor & "." & App.Revision
    Open "c:\tempFtp\Version.Ver" For Output As #1
    Write #1, MyVer
    Close #1
    MsgBox "versão 1.1.0"
    Me.Show
    Screen.MousePointer = vbHourglass
    lblStatus.Caption = "Realizando download de atualizações ..."
    retorno = carrega.DownloadArquivo("Updates.txt")
    lblStatus.Caption = "Download completo!"
    Screen.MousePointer = vbDefault
    Command1.Visible = True
End Sub
