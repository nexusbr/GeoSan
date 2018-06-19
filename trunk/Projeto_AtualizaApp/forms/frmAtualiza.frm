VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAtualiza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualiza GeoSan"
   ClientHeight    =   4620
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1800
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton fecha 
      Caption         =   "Inicia GeoSan"
      Height          =   495
      Left            =   4418
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   9975
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   3840
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAtualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Encerra o programa de atualização
'
'
'
Private Sub fecha_Click()
    On Error GoTo Trata_Erro
    Shell ("C:\Arquivos de programas\GeoSan\GeoSan.exe")
    End
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmAtualiza", "fecha_Click", CStr(Err.Number), CStr(Err.Description), True, True
    End If
End Sub
' Carrega a caixa de diálogo que mostrará as atualizações
'
'
'
Private Sub Form_Load()
    On Error GoTo Trata_Erro
    
    Dim atualiza As CAtualiza                               'classe para atualizar tanto remoto para servidor quanto servidor para cliente
    Dim retorno As Boolean
    
    Set Email = New CEmail
    retorno = Email.leConfiguracoesEmail
    Set atualiza = New CAtualiza
    Set versao = New CGetVersion
    Me.Show
    frmAtualiza.fecha.Enabled = False
    Screen.MousePointer = vbHourglass
    mensagem = "Iniciando o download das atualizações ..."
    frmAtualiza.Text1 = mensagem
    Timer1.Enabled = True                                   'ativa o timer
    Me.ProgressBar1.Visible = True                          'ativa a visualização da barra de progresso
    retorno = atualiza.AtualizaDirRemoto                    'atualiza os arquivos do servidor da NEXUS para o servidor de arquivos do cliente
    retorno = atualiza.AtualizaAplicacaoLocal               'atualiza os arquivos do servidor do cliente para o diretório da máquina local da aplicação do cliente
    ErroUsuario.Registra "frmAtualiza", "Form_Load - Atualização realizada", CStr(Err.Number), CStr(Err.Description), False, True, mensagem
    mensagem = mensagem & vbCrLf & vbCrLf & "Final do processamento das atualizações"
    frmAtualiza.Text1 = mensagem
    frmAtualiza.ProgressBar1.Value = frmAtualiza.ProgressBar1.Max
    Timer1.Enabled = False               'ativa o timer
    Me.ProgressBar1.Visible = False      'ativa a visualização da barra de progresso
    Screen.MousePointer = vbDefault
    frmAtualiza.fecha.Enabled = True
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmAtualiza", "Form_Load", CStr(Err.Number), CStr(Err.Description), True, True
    End If
End Sub
' Não está mais sendo utilizado
Private Sub ObtemVersao_Click()
    On Error GoTo Trata_Erro
    Dim retorno As Boolean
    
    retorno = versao.ExisteArquivo("D:\Desenv\GEOSAN_VB6_B\trunk\Projeto_AtualizaApp\GeoSanIni.exe")
    If retorno = True Then
        MsgBox versao.ObtemVersaoArquivo("D:\Desenv\GEOSAN_VB6_B\trunk\Projeto_AtualizaApp\GeoSanIni.exe")
    End If

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmAtualiza", "ObtemVersao_Click", CStr(Err.Number), CStr(Err.Description), True, True
    End If
End Sub

' Evento disparado quando o usuário seleciona o X indicando que não quer mais atualizar o software nem entrar no GeoSan
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        End
    End If
End Sub

' Para barra de progresso
'
'
'
Private Sub Timer1_Timer()
    MousePointer = vbHourglass              'ativa a ampulheta
    'INICIAR                                 'inicia a conversão para o EPANET
    MousePointer = vbDefault                'desativa a ampulheta
    Timer1.Enabled = False                  'desativa o timer
    End
End Sub
