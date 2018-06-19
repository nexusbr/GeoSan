VERSION 5.00
Object = "{F141AAE7-BD4C-4793-88DB-10FE8CBA518C}#1.0#0"; "TeComPrinter.dll"
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Object = "{EE78E37B-39BE-42FA-80B7-E525529739F7}#1.0#0"; "TeComViewDatabase.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTePrinter 
   Caption         =   "Módulo de Impressão"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   9960
   StartUpPosition =   1  'CenterOwner
   Begin TeComPrinterLibCtl.TePrinter TePrinter1 
      Height          =   9375
      Left            =   120
   Begin VB.Menu mnArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnConfPagina 
         Caption         =   "Configurar Página"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnFechar 
         Caption         =   "Fechar"
      End
   End
   Begin VB.Menu mnEditar 
      Caption         =   "Editar"
      Begin VB.Menu mnFonte 
         Caption         =   "Fonte"
         Begin VB.Menu mnTamanho 
            Caption         =   "Tamanho"
            Begin VB.Menu mn8 
               Caption         =   "8"
            End
            Begin VB.Menu mn12 
               Caption         =   "12"
            End
            Begin VB.Menu mn16 
               Caption         =   "16"
            End
            Begin VB.Menu mn32 
               Caption         =   "32"
            End
         End
         Begin VB.Menu mnTipo 
            Caption         =   "Tipo"
            Begin VB.Menu mnArial 
               Caption         =   "Arial"
            End
            Begin VB.Menu mnCourrier 
               Caption         =   "Arial Negrito"
            End
         End
      End
   End
   Begin VB.Menu mnInserir 
      Caption         =   "Inserir"
      Begin VB.Menu mnMapa 
         Caption         =   "Mapa"
      End
      Begin VB.Menu mnImagem 
         Caption         =   "Imagem"
      End
      Begin VB.Menu mnTexto 
         Caption         =   "Texto"
      End
      Begin VB.Menu mnLinha 
         Caption         =   "Linha"
      End
      Begin VB.Menu mnRetangulo 
         Caption         =   "Retangulo"
      End
      Begin VB.Menu mnElipse 
         Caption         =   "Elipse"
      End
      Begin VB.Menu mnBarraEscala 
         Caption         =   "Barra de Escala"
      End
   End
   Begin VB.Menu mnAcaoMapa 
      Caption         =   "Mapa"
      Begin VB.Menu mnMover 
         Caption         =   "Mover"
      End
      Begin VB.Menu mnAproximar 
         Caption         =   "Aproximar"
      End
      Begin VB.Menu mnAfastar 
         Caption         =   "Afastar"
      End
      Begin VB.Menu mnEscala 
         Caption         =   "Definir Escala"
      End
   End
End
Attribute VB_Name = "frmTePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ObjName As String
Dim contador As Integer
Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
Dim username As String



Private Sub mnEscala_Click()
On Error GoTo saida
   Dim Scala As Double
   

   If CStr(ObjName) <> "" Then
  
      Scala = InputBox("", "Informe a escala do mapa")
      
      If Scala > 0 Then
         TePrinter1.setObjectPropertyValueByName ObjName, "MapScale_Property", CDbl(Scala)
         TePrinter1.execute
         
      End If
   Else
   
      MsgBox "Selecione o mapa antes de definir a escala.", vbInformation, ""
   
   End If
saida:

End Sub

Private Function MudaFonte(Tam As Double)
On Error GoTo saida
Dim tamTexto As String

   'retorna todos os nomes de propriedades
   'tamTexto = TePrinter1.getLayoutObjectPropertyName(ObjName, 5)

   TePrinter1.setObjectPropertyValue ObjName, 6, Tam
   TePrinter1.execute
   

saida:

End Function

Private Function MudaTipoFonte(tipo As String)
On Error GoTo saida

   'ObjName = Me.TePrinter1.getLayoutObjectPropertyValue("Text_2", 7) 'fonte
   
   TePrinter1.setObjectPropertyValue ObjName, 7, tipo
   TePrinter1.execute

saida:
End Function

Public Function id_objeto()
Dim qtd As Integer

   'VERIFICA SE ALGO SELECIONADO
   qtd = Me.TePrinter1.getLayoutSelectObjectCount
   If qtd > 0 Then
      
      'PASSA PARA O ObjName O NOME DO OBJETO SELECIONADO
      ObjName = Me.TePrinter1.getLayoutSelectObjectName(0)

   End If
   
   
   
End Function


Private Sub cmdSetScale_Click()
      Dim Texto As String
   Texto = InputBox("Informe o texto:")
   
   If Trim(Texto) <> "" Then
      TePrinter1.addText Texto
   End If
   
   TePrinter1.setObjectPropertyValueByName strObjSelecionado, "MapScale_Property", CDbl(txtScala.Text)
   TePrinter1.execute
End Sub


Private Sub Form_Load()

   TePrinter1.Provider = frmCanvas.TipoConexao
   
If (frmCanvas.TipoConexao = 4) Then
    
If (contador <> 10) Then

mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
username = ReadINI("CONEXAO", "USER", App.path & "\CONTROLES\GEOSAN.ini")
frmCanvas.FunDecripta (Senha)
decriptada = frmCanvas.Senha

TeAcXConnection1.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA
contador = 10

Else

TePrinter1.Provider = frmCanvas.TipoConexao
'TePrinter1.connection = TeAcXConnection1.objectConnection_

End If

Else
TePrinter1.Provider = frmCanvas.TipoConexao
TePrinter1.connection = Conn

End If

TePrinter1.execute
  
End Sub

Private Sub Form_Resize()
On Error GoTo Trata_Erro

TePrinter1.Width = 9960
TePrinter1.Height = Me.Height - 1250 '- tbMain.Height

Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   Exit Sub
End If

End Sub

Private Sub mn8_Click()
   MudaFonte 8
End Sub

Private Sub mn12_Click()
   
   
   TePrinter1.setObjectPropertyValue ObjName, 6, 12 'CLng(tamanho)
   TePrinter1.execute
   
End Sub

Private Sub mn16_Click()
   MudaFonte 16
End Sub

Private Sub mn32_Click()
   MudaFonte 32
End Sub



Private Sub mnAfastar_Click()
   TePrinter1.setMapMode zoomOut
End Sub

Private Sub mnAproximar_Click()
   TePrinter1.setMapMode zoomIn
End Sub

Private Sub mnArial_Click()
   MudaTipoFonte "Arial"
End Sub

Private Sub mnCourrier_Click()
   MudaTipoFonte "Arial Black"
End Sub

Private Sub mnLucConsole_Click()
   MudaTipoFonte "Lucida Console"
End Sub

Private Sub mnConfPagina_Click()
    If TePrinter1.orientation = portrait Then
         frmPageSetup.setOrientation True
    Else
         frmPageSetup.setOrientation False
    End If
    
    frmPageSetup.setPageSize TePrinter1.PaperSize
        
    frmPageSetup.Show 1, Me
    If frmPageSetup.getOK = True Then
            If frmPageSetup.getOrientation = True Then
                TePrinter1.orientation = portrait
            Else
                TePrinter1.orientation = landscape
            End If
            TePrinter1.PaperSize = frmPageSetup.getPageSize
            
            TePrinter1.execute
    End If
End Sub







Private Sub mnImagem_Click()
    cmmOpen.ShowOpen
    TePrinter1.addImage cmmOpen.FileName
End Sub

Private Sub mnImprimir_Click()
   
   TePrinter1.printExecute ("Teste")

End Sub

Private Sub mnLinha_Click()
   
   TePrinter1.addLine
   
End Sub



Private Sub mnMover_Click()
   TePrinter1.setMapMode pan
End Sub

Private Sub mnBarraEscala_Click()
   TePrinter1.addScale
End Sub




Private Sub mnMapa_Click()
         
   'COLETA AS VARIÁVEIS GLOBAIS PARA O MÓDULO DE IMPRESSÃO PASSADAS NO END PLOT VIEW
   'CanvasXmin_, CanvasYmin_, CanvasXmax_, CanvasYmax_

   'ADICIONA O MAPA QUE ESTAVA NA TELA DO GEOSAN
   'TePrinter1.addMap strViewAtiva_, strUser, CanvasXmin_, CanvasYmin_, CanvasXmax_, CanvasYmax_

If (frmCanvas.TipoConexao = 4) Then

'TePrinter1.provider = TeComPrinterLib.CONNECTION_TYPE.PostgreSQL;
TePrinter1.addDatabase
TePrinter1.execute
'TePrinter1.Open ("nexussql.gpl")
'TePrinter1.execute



TePrinter1.setObjectPropertyValueByName "Database_1", "left_Property", "15.0"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Database_1", "top_Property", "20.0"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Database_1", "width_Property", "11.0"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Database_1", "height_Property", "11.0"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Database_1", "angle_Property", "0.0"
TePrinter1.execute

TePrinter1.setObjectPropertyValueByName "Database_1", "username_Property", "postgres"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Database_1", "password_Property", "gustavo"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Database_1", "host_Property", "localhost"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Database_1", "databasename_Property", "teste"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Database_1", "databaseconnected_Property", "true"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Database_1", "databasetype_Property", "2.0"
TePrinter1.execute



TePrinter1.setObjectPropertyValueByName "Map_3", "viewname_Property", "Administrador"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Map_3", "viewuser_Property", "Administrador"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Map_3", "connection_Property", "Database_1"
TePrinter1.execute
         
TePrinter1.setObjectPropertyValueByName "Map_3", "left_Property", "15.0"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Map_3", "top_Property", "44.0"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Map_3", "width_Property", "300.0"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Map_3", "height_Property", "200.0"
TePrinter1.execute
TePrinter1.setObjectPropertyValueByName "Map_3", "angle_Property", "0.0"
TePrinter1.execute

TePrinter1.setObjectPropertyValueByName "Map_3", "mapscale_Property", "2200000"
TePrinter1.execute

TePrinter1.setObjectPropertyValueByName "Map_3", "fixedscale_Property", "False"
TePrinter1.execute




'TePrinter1.setObjectPropertyValueByName "Map_3", "connection_Property", "Database_1"
'TePrinter1.execute
         

'TePrinter1.addDatabase

'TePrinter1.execute
'TePrinter1.setObjectPropertyValueByName "Database_1", "databaseconnected_Property", "true"




End If

'update te_layer set lower_x = 715603;
'update te_layer set lower_y = 7702100;
'update te_layer set upper_x = 722075;
'update te_layer set upper_y = 7711330;
TePrinter1.addMap strViewAtiva_, strUser, CanvasXmin_, CanvasYmin_, CanvasXmax_, CanvasYmax_
'TePrinter1.addMap strViewAtiva_, strUser, CanvasXmin_, CanvasYmin_, CanvasXmax_, CanvasYmax_
   
End Sub

Private Sub mnRetangulo_Click()
   TePrinter1.addRectangle
End Sub

Private Sub mnElipse_Click()
   TePrinter1.addEllipse
End Sub

Private Sub mnTexto_Click()
   Dim Texto As String
   Texto = InputBox("Informe o texto:")
   
   If Trim(Texto) <> "" Then
      TePrinter1.addText Texto
   End If

End Sub


Private Sub TePrinter1_endProcess(ByVal partial As Boolean)

   DoEvents
   If (partial = False) Then
      'mnProperties_Click
      id_objeto
   End If
End Sub


