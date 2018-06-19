VERSION 5.00
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TeComDatabase.dll"
Object = "{C51C74EC-6107-4A01-8400-40B53BB20D42}#1.0#0"; "TeComExport.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{1A397116-3057-40EE-9ECA-6FA4CC1E5FC3}#1.0#0"; "NexusPM4.ocx"
Object = "{2CCABA93-B681-4E7F-8047-BD4D623301BA}#1.0#0"; "TeComImport.dll"
Object = "{91488A85-7250-4842-8681-87818334B791}#1.0#0"; "NxViewManager2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "  NEXUS - GeoSan - Tecnologia Terralib"
   ClientHeight    =   8445
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10035
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "FrmMain"
   StartUpPosition =   3  'Windows Default
   Tag             =   "0"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   3720
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":013E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0838
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3294
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":398E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4088
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4782
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5576
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":636A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A64
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":715E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7858
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":864C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":94C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A024
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A48E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AC08
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B302
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B9FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   4305
      Top             =   2430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pctSfondo 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7515
      Left            =   6090
      ScaleHeight     =   7515
      ScaleWidth      =   3945
      TabIndex        =   1
      Top             =   510
      Width           =   3945
      Begin NxViewManager.ViewManager ViewManager1 
         Height          =   1095
         Left            =   600
         TabIndex        =   11
         Top             =   3960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1931
      End
      Begin TECOMEXPORTLibCtl.TeExport TeExport2 
         Left            =   2880
         OleObjectBlob   =   "frmMain.frx":2C0F6
         Top             =   1680
      End
      Begin TECOMIMPORTLibCtl.TeImport TeImport1 
         Left            =   2880
         OleObjectBlob   =   "frmMain.frx":2C11A
         Top             =   5160
      End
      Begin PManager4.Manager Manager1 
         Height          =   855
         Left            =   600
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1508
      End
      Begin MSComDlg.CommonDialog cmmSalvaImg 
         Left            =   1590
         Top             =   7605
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdClose 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         Picture         =   "frmMain.frx":2C13E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   285
      End
      Begin VB.PictureBox picSplitter 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillColor       =   &H00808080&
         Height          =   7200
         Left            =   2250
         ScaleHeight     =   3135.189
         ScaleMode       =   0  'User
         ScaleWidth      =   780
         TabIndex        =   2
         Top             =   -90
         Visible         =   0   'False
         Width           =   72
      End
      Begin VB.Frame FrameEscala 
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   3675
         Begin VB.TextBox txtEscala 
            Height          =   285
            Left            =   1890
            TabIndex        =   5
            Top             =   210
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "Escala de Visualização:"
            Height          =   195
            Left            =   90
            TabIndex        =   6
            Top             =   240
            Width           =   1755
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         MultiRow        =   -1  'True
         MultiSelect     =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Temas"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Propriedades"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
         Left            =   2880
         OleObjectBlob   =   "frmMain.frx":2C4B0
         Top             =   2760
      End
      Begin TECOMDATABASELibCtl.TeDatabase TeDatabase1 
         Left            =   2760
         OleObjectBlob   =   "frmMain.frx":2C4D4
         Top             =   4080
      End
      Begin VB.Image imgSplitter 
         Height          =   7380
         Left            =   0
         MousePointer    =   9  'Size W E
         Top             =   0
         Width           =   150
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   8025
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            TextSave        =   "20:16"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Distância"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   900
      ButtonWidth     =   820
      ButtonHeight    =   794
      ImageList       =   "ImageList4"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   42
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "knew"
            Object.ToolTipText     =   "Novo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ksave"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kselection"
            Object.ToolTipText     =   "Selecionar"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kplotview"
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "krecompose"
            Object.ToolTipText     =   "Recompor"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "mnuPoligono"
            Object.ToolTipText     =   "Poligono de Seleção"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KFindCoordenadas"
            Object.ToolTipText     =   "Localizar Coordenada"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEncontraTexto"
            Object.ToolTipText     =   "Encontrar Textos"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEncontraConsumidor"
            Object.ToolTipText     =   "Encontrar Consumidores"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kzoomarea"
            Object.ToolTipText     =   "Zoom Área"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kpan"
            Object.ToolTipText     =   "Mover Visualização"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kundoview"
            Object.ToolTipText     =   "Voltar Visualização"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kredoview"
            Object.ToolTipText     =   "Avançar Visualização"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kzoomin"
            Object.ToolTipText     =   "Menos Zoom"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kzoomout"
            Object.ToolTipText     =   "Mais Zoom"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kdrawnetworkline"
            Object.ToolTipText     =   "Desenhar Rede"
            ImageIndex      =   16
            Style           =   1
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kmovenetworknode"
            Object.ToolTipText     =   "Mover Componente com Rede"
            ImageIndex      =   17
            Style           =   1
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kinsertnetworknode"
            Object.ToolTipText     =   "Desenhar Componente na Rede"
            ImageIndex      =   18
            Style           =   1
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "kdrawtext"
            Object.ToolTipText     =   "Desenhar Texto - Amarração"
            ImageIndex      =   19
            Style           =   1
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "kdrawline"
            Object.ToolTipText     =   "Desenhar Linha - Amarração"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "kdrawpoint"
            Object.ToolTipText     =   "Desenhar Ponto - Amarração"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kinsertdoc"
            Object.ToolTipText     =   "Inserir Documento(s)"
            ImageIndex      =   19
            Style           =   1
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kdrawramal"
            Object.ToolTipText     =   "Inserir Ramal"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kdrawramalAuto"
            Object.ToolTipText     =   "Inserir ramal automaticamente"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kdrawramalAddConsumer"
            Object.ToolTipText     =   "Adiciona um consumidor a um ramal"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kMoveConsumidorGPS"
            Object.ToolTipText     =   "Mover Consumidor GPS"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kdelete"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ksearchinnetwork"
            Object.ToolTipText     =   "Encontrar válvulas a partir da rede selecionada"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "ksearchattribute"
            Object.ToolTipText     =   "Pesquisa Geral"
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "kdrawintersection"
         EndProperty
         BeginProperty Button41 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "kConsumoLote"
            Object.ToolTipText     =   "Apresenta Consumo"
         EndProperty
         BeginProperty Button42 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "kMoveVertice"
            Object.ToolTipText     =   "Mover Vértice"
         EndProperty
      EndProperty
      Begin VB.Timer Timer1 
         Left            =   3240
         Top             =   840
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   9120
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Min             =   1e-4
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4830
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C4F8
            Key             =   "new_window"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C84A
            Key             =   "zoom_area"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CB9C
            Key             =   "zoom_in"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CEEE
            Key             =   "zoon_out"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D240
            Key             =   "undo_view"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D592
            Key             =   "redo_view"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D8E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DE36
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E188
            Key             =   "save"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E4DA
            Key             =   "fit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E82C
            Key             =   "insertramal"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EB7E
            Key             =   "foto"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EED0
            Key             =   "fonte"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F222
            Key             =   "fiti"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F574
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F8C6
            Key             =   "seta"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FC18
            Key             =   "world"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FF6A
            Key             =   "find_user"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":302BC
            Key             =   "insert_point"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3060E
            Key             =   "attach"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30960
            Key             =   "draw_network"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30CB2
            Key             =   "find_valvula"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31004
            Key             =   "declivity"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31356
            Key             =   "reflesh"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":316A8
            Key             =   "move_point"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":319FA
            Key             =   "pan"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31D4C
            Key             =   "registro"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3209E
            Key             =   "point"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":323F0
            Key             =   "point2"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":327BB
            Key             =   "line"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32B0D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5400
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   32
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32F5F
            Key             =   "new_window"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":332B1
            Key             =   "zoom_area"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33603
            Key             =   "zoom_in"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33955
            Key             =   "zoon_out"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33CA7
            Key             =   "undo_view"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33FF9
            Key             =   "redo_view"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3434B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3489D
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34BEF
            Key             =   "save"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34F41
            Key             =   "fit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35293
            Key             =   "insertramal"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":355E5
            Key             =   "foto"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35937
            Key             =   "fonte"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35C89
            Key             =   "fiti"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35FDB
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3632D
            Key             =   "seta"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3667F
            Key             =   "world"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":369D1
            Key             =   "find_user"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36D23
            Key             =   "insert_point"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37075
            Key             =   "attach"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":373C7
            Key             =   "draw_network"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37719
            Key             =   "find_valvula"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37A6B
            Key             =   "declivity"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37DBD
            Key             =   "reflesh"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3810F
            Key             =   "move_point"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38461
            Key             =   "pan"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":387B3
            Key             =   "registro"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38B05
            Key             =   "point"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38E57
            Key             =   "point2"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39222
            Key             =   "line"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39574
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":399C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   4320
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A453
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EFA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41FF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45049
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4809B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B0ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CF6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D38D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":593DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65431
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71483
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D4D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":89527
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":95579
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":961CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A221D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AE26F
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BA2C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C6313
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D2365
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D53B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D57FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D5C2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D6068
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnu_Arquive 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuExport 
         Caption         =   "Exportar"
         Begin VB.Menu mnuImagem 
            Caption         =   "Área Visualizada para Imagem"
         End
         Begin VB.Menu mnuExportLocalNos 
            Caption         =   "Localização de Nós com Cota"
         End
         Begin VB.Menu mnuExpAutoCad 
            Caption         =   "DXF"
         End
         Begin VB.Menu mnuExpCon 
            Caption         =   "Consumidores e Consumo"
         End
         Begin VB.Menu OdImport 
            Caption         =   "Novo DXF com OdImport"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnusep01001 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Importar"
         Begin VB.Menu mnuImportSIG 
            Caption         =   "SIG"
         End
         Begin VB.Menu mnuImportDXF 
            Caption         =   "DXF"
         End
         Begin VB.Menu mnuImpCotas 
            Caption         =   "Cotas"
         End
      End
      Begin VB.Menu mnusep011101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate_Demand 
         Caption         =   "Atualizar Consumos e Distribuir Demandas"
      End
      Begin VB.Menu mnuExporta_GeoSan 
         Caption         =   "Exporta consumidores, redes, ramais e nós no formato .shp"
      End
      Begin VB.Menu mnuAtualizaCotas 
         Caption         =   "Atualiza todas as cotas de todos os nós"
      End
      Begin VB.Menu mnuAutoLogin 
         Caption         =   "Logar Automaticamente"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Alterar Senha"
      End
      Begin VB.Menu mnuAtualizaNumCasaGps 
         Caption         =   "Atualizar números das casas dos coletores GPS"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Fechar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propriedades"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Editar"
      Begin VB.Menu mnuSelect 
         Caption         =   "Selecionar"
      End
      Begin VB.Menu mnu_Reflesh 
         Caption         =   "Atualizar"
      End
      Begin VB.Menu mnuRecompose 
         Caption         =   "Recompor"
      End
      Begin VB.Menu mnuFileBar71 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom Área"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Mover Visualização"
      End
      Begin VB.Menu mnuUndoView 
         Caption         =   "Voltar Visualização"
      End
      Begin VB.Menu mnuRedoView 
         Caption         =   "Avançar Visualização"
      End
      Begin VB.Menu mnuFileBar72 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinusZoom 
         Caption         =   "Menos Zoom"
      End
      Begin VB.Menu mnuMoreZoom 
         Caption         =   "Mais Zoom"
      End
      Begin VB.Menu mnuFileBar73 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDeleteInc 
         Caption         =   "Excluir Pontos Inconcistentes"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMapa 
      Caption         =   "&Mapa"
      Begin VB.Menu mnuDrawLineWater 
         Caption         =   "Desenhar Rede Água"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuMovePointWithLines 
         Caption         =   "Mover Componente c/ Rede"
      End
      Begin VB.Menu mnuDrawPointInLineWater 
         Caption         =   "Desenhar Componente na Rede"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditBar30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertDocs 
         Caption         =   "Inserir Documentos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditBar80 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrawRamal 
         Caption         =   "Desenhar Ramal"
      End
      Begin VB.Menu mnusep1234 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteLineWater 
         Caption         =   "Excluir"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnusep9999 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalcArea 
         Caption         =   "Calcular Área (m²)"
      End
      Begin VB.Menu mnuCalibrarZoom 
         Caption         =   "Calibrar Zoom"
      End
      Begin VB.Menu mnuDesenhaPoligono 
         Caption         =   "Desenhar Poligono"
      End
      Begin VB.Menu mnuCarregaPoligono 
         Caption         =   "Carregar Polígono"
      End
      Begin VB.Menu mnuDefEscala 
         Caption         =   "Definir Escala"
      End
      Begin VB.Menu mnuFixaIcone 
         Caption         =   "Fixar Ícone"
      End
      Begin VB.Menu mnusep0001 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Find_Object 
         Caption         =   "Encontrar Objeto"
      End
      Begin VB.Menu mnuEncontraTexto 
         Caption         =   "Encontrar Textos"
      End
      Begin VB.Menu mnuEncontraCoordenada 
         Caption         =   "Localizar Coordenadas"
      End
      Begin VB.Menu mnuLocConsumidores 
         Caption         =   "Localizar Consumidores"
      End
      Begin VB.Menu mnusep10000 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalvaImgMapa 
         Caption         =   "Salvar Imagem"
         Begin VB.Menu mnuBitmap 
            Caption         =   "BMP"
         End
         Begin VB.Menu mnuGIF 
            Caption         =   "GIF"
         End
         Begin VB.Menu mnuJPG 
            Caption         =   "JPG"
         End
         Begin VB.Menu mnuPNG 
            Caption         =   "PNG"
         End
         Begin VB.Menu mnuTIF 
            Caption         =   "TIF"
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "&Cadastros"
      Begin VB.Menu mnuTypes 
         Caption         =   "Tipos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   "Fornecedores"
      End
      Begin VB.Menu mnuManufacters 
         Caption         =   "Fabricantes"
      End
   End
   Begin VB.Menu mnuRel 
      Caption         =   "Relatórios"
      Begin VB.Menu MnuRelWl 
         Caption         =   "Rede de Água"
      End
      Begin VB.Menu MnuRelSl 
         Caption         =   "Rede de Esgoto"
      End
      Begin VB.Menu mnuRelRegistros 
         Caption         =   "Válvulas"
      End
      Begin VB.Menu mnuRelComponentesAgua 
         Caption         =   "Componentes de Água"
      End
      Begin VB.Menu mnuRelComponentesEsgoto 
         Caption         =   "Componentes de Esgoto"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administrar"
      Begin VB.Menu mnuRemoverPlano 
         Caption         =   "Remover Plano"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Usuários"
      End
      Begin VB.Menu mnuSep113 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectDatabase 
         Caption         =   "Banco de Dados"
      End
      Begin VB.Menu s999999 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudiagRede 
         Caption         =   "Diagnóstico de rede"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProdutividade 
         Caption         =   "Indicador de Produtividade"
         Begin VB.Menu mnuRedesAgua 
            Caption         =   "Redes de Água"
         End
         Begin VB.Menu mnuRamaisAgua 
            Caption         =   "Ligações de Água"
         End
         Begin VB.Menu mnusep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRedesEsgoto 
            Caption         =   "Redes de Esgoto"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Janela"
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Barra de Status"
         Checked         =   -1  'True
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Barra de Ferramentas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLayers 
         Caption         =   "Layers e Propriedades"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep016 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalculaZNo 
         Caption         =   "Calcular Z Nó Enquanto Desenha"
      End
      Begin VB.Menu mnuMultProperteis 
         Caption         =   "Propriedades Multiplas"
      End
      Begin VB.Menu mnuLoadAttributeByReference 
         Caption         =   "Carregar Attributos por Referência"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusep100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Cascata"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Alinhamento Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Alinhamento Vertical"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Conteúdo"
         Shortcut        =   ^H
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "Sobre"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TCanvas As frmCanvas
Const sglSplitLimit = 0

Private mbMoving        As Boolean
Private mstrprevfunc    As String
Private msngStartX      As Single

Dim conee As TeAcXConnection
Dim Abertura As Integer
Dim teac As TeAcXConnection
Dim a1 As TeImport
Dim a2 As TeDatabase




Private Sub cmdClose_Click()

   pctSfondo.Visible = False: mnuLayers.Checked = False
   
End Sub




Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    msngStartX = X
    
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    
    picSplitter.Visible = True
    picSplitter.Height = pctSfondo.Height
    mbMoving = True
    
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
        
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
    
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    pctSfondo.Width = pctSfondo.Width + msngStartX - X
    pctSfondo.Refresh
    picSplitter.Visible = False
    mbMoving = False
    MDIForm_Resize
    TabStrip1.Refresh
    
End Sub



Private Sub MDIForm_Activate()
   'MDIForm_Resize
End Sub
' Carrega o formulário principal. Rotina de entrada
'
'
Private Sub MDIForm_Load()
    Email.leConfiguracoesEmail                                  'lê as configurações de email para ele poder enviar mensagem de email para NEXUS caso ocorra um erro
    '''LoozeXP1.InitSubClassing
    Manager1.InitConn Conn, CInt(typeconnection)
    Manager1.GridVisibled False
    FrmMain.Timer1.Interval = 100                               'define o intervalo em que ele vai verificar se alguma tecla foi pressionada
    FrmMain.Timer1.Enabled = False                              'inicia com o timer desligado, só liga quando tiver cálculo intensivo
End Sub

Private Sub mnu_Find_Object_Click()
On Error GoTo Trata_Erro
   
   
   
   Dim Object_id_ As String, xmin As Double, ymin As Double, xmax As Double, ymax As Double
   
   Object_id_ = ""
   
   Object_id_ = InputBox("Informe o identificador do objeto", "Encontrar Objeto")
   
   If Trim(Object_id_) <> "" Then
      
      With ActiveForm.TCanvas
         .Normal
         If .addSelectObjectIds(Object_id_) = 1 Then
            .getSelectBox xmin, ymin, xmax, ymax
            .setWorld xmin - 1000, ymin - 1000, xmax + 1000, ymax + 1000
            .Select
            .setScale 1000
         Else
            MsgBox "Não foi encontrado a geometria referente ao atributo selecionado", vbExclamation
         End If
      End With
      
   End If
   
Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
ElseIf Err.Number = 91 Then
    MsgBox "Não há mapa ativo.", vbInformation, "Geosan"
Else
   
   PrintErro CStr(Me.Name), "Private Sub mnu_Find_Object_Click", CStr(Err.Number), CStr(Err.Description), True
   
End If
   
End Sub

' Atualiza todas as cotas existentes com os valores da interpolação do MDT
'
'
'
Private Sub mnuAtualizaCotas_Click()
     Dim setaZs As New CAcertaZsDosNos
     
     varGlobais.pararExecucao = False               'indica que iniciará sem sem informar que deverá parar a execução
     FrmMain.Timer1.Enabled = True                  'habilita o timer
     setaZs.AtribuiZs                               'chama método para atualizar todas as cotas da cidade toda
     FrmMain.Timer1.Enabled = False                 'deshabilita o timer
End Sub

Private Sub mnuAtualizaNumCasaGps_Click()
    Dim consumidorGps As clsConsumidorControler
    Set consumidorGps = New clsConsumidorControler
    
    Call consumidorGps.AtualizaNumeroGpsCasas
End Sub

Private Sub mnuAutoLogin_Click()
   
   Close #1
   Open App.path & "\controles\AutoLogin.txt" For Output As #1
   Print #1, strUser
   Close #1
   MsgBox "Definido login automático para " & strUser, vbInformation, ""
   
End Sub



Private Sub mnuCalcArea_Click()
   
   ActiveForm.TCanvas.ToolTipText = ""
   
   ActiveForm.TCanvas.calculateArea
   
   
   
End Sub

Private Sub mnuCalibrarZoom_Click()
   Dim calibrazoom As New frmCalibrarZoom
   calibrazoom.Show 1
End Sub



Private Sub mnuEncontraCoordenada_Click()
On Error GoTo Trata_Erro
    Dim X As Double, Y As Double
   
    X = InputBox("Informe a Coordena X ")
    Y = InputBox("Informe a Coordena Y ")
    
    If X <> 0 And Y <> 0 Then
        ActiveForm.TCanvas.setWorld X - 50, Y - 50, X + 50, Y + 50
        ActiveForm.TCanvas.plotView
    End If

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next

ElseIf Err.Number = 91 Or Err.Number = 13 Then
    'MsgBox "Não há mapa ativo.", vbInformation, "Geosan"
    Exit Sub
Else
   
   PrintErro CStr(Me.Name), "Public Function EncontraCoord()", CStr(Err.Number), CStr(Err.Description), True
   

End If

End Sub


Private Sub mnu_PatternCurves_Click()

   Dim frm As New frmEPANavegator
   
   frm.Init
   Set frm = Nothing
   
End Sub

Private Sub mnu_Reflesh_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kplotview")
   
End Sub
' Exportação de consumidores
'
'
'
Private Sub mnuExpCon_Click()
    frmExportaConsumos.Show
End Sub
' Exporta a localização dos nós com as suas respectivas cotas
'
'
'
Private Sub mnuExportLocalNos_Click()
    On Error GoTo Trata_Erro
    Dim CAMINHO As String
    Dim rs As New ADODB.Recordset
    Dim TB_GEOMETRIA As String
    Dim a As String
    Dim b As String
    Dim c As String
    Dim d As String
    Dim e As String
    Dim f As String
    Dim g As String
    Dim h As String
    Dim i As String
    Dim j As String
    Dim k As String
    Dim l As String
    Dim nomeArquivo As New CArquivo                                                             'para o usuário selecionar onde será salvo o arquivo
        
    varGlobais.pararExecucao = False               'indica que iniciará sem sem informar que deverá parar a execução
    FrmMain.Timer1.Enabled = True                  'habilita o timer
    CAMINHO = nomeArquivo.SelecionaDiretorio                                                    'solicita ao usuário a seleção de um diretório
    CAMINHO = CAMINHO + "\" + nomeArquivo.prefixo + "LOCALIZAÇÃO_NOS_REDE_AGUA.txt"  'coloca um prefixo de data e hora em que o arquivo será gerado
    a = "geom_table"
    b = "te_representation"
    c = "geom_type"
    d = "layer_id"
    e = "te_layer"
    f = "name"
    g = "WATERCOMPONENTS"
    If frmCanvas.TipoConexao <> 4 Then
        'retornar a tabela de geometria de pontos
        rs.Open "SELECT GEOM_TABLE FROM TE_REPRESENTATION WHERE GEOM_TYPE = 4 AND LAYER_ID IN (SELECT LAYER_ID FROM TE_LAYER WHERE NAME = 'WATERCOMPONENTS')", Conn, adOpenDynamic, adLockReadOnly
        If rs.EOF = False Then
            TB_GEOMETRIA = rs!GEOM_TABLE
        Else
            MsgBox "Não foi encontrada a tabela de Geometrias.", vbInformation, ""
        Exit Sub
        End If
    Else
        'SELECT "geom_table" FROM "te_representation" WHERE "geom_type" = '4' AND "layer_id" IN (SELECT "layer_id" FROM "te_layer" WHERE "name" = 'WATERCOMPONENTS')
        rs.Open "SELECT " + """" + a + """" + " FROM " + """" + b + """" + " WHERE " + """" + c + """" + " = '4' AND " + """" + d + """" + " IN (SELECT " + """" + d + """" + " FROM " + """" + e + """" + " WHERE " + """" + f + """" + " = 'WATERCOMPONENTS')", Conn, adOpenDynamic, adLockOptimistic
        If rs.EOF = False Then
            TB_GEOMETRIA = rs!GEOM_TABLE
        Else
            MsgBox "Não foi encontrada a tabela de Geometrias.", vbInformation, ""
            Exit Sub
        End If
    End If
    rs.Close
    If frmCanvas.TipoConexao = 1 Then
        'SELECIONA DO BANCO DE DADOS O CÓDIGO DO NÓ, COORDENADAS E COTA ATUAL
        rs.Open "SELECT W.GROUNDHEIGHT AS " + """" + "COTA" + """" + ", LEN(P.OBJECT_ID) AS " + """" + "TAM" + """" + ", P.OBJECT_ID ,P.X,P.Y FROM " & TB_GEOMETRIA & " P JOIN WATERCOMPONENTS W ON P.OBJECT_ID = W.OBJECT_ID_ ORDER BY TAM, OBJECT_ID"
    ElseIf frmCanvas.TipoConexao = 2 Then
        rs.Open "SELECT W.GROUNDHEIGHT AS " + """" + "COTA" + """" + ", P.OBJECT_ID AS " + """" + "TAM" + """" + ", P.OBJECT_ID ,P.X,P.Y FROM " & TB_GEOMETRIA & " P JOIN WATERCOMPONENTS W ON P.OBJECT_ID = W.OBJECT_ID_ ORDER BY TAM, OBJECT_ID"
    Else
        a = "INITIALGROUNDHEIGHT"
        b = "object_id"
        c = "x"
        d = "y"
        e = "WATERCOMPONENTS"
        f = LCase(TB_GEOMETRIA)
        g = "OBJECT_ID_"
        rs.Open "SELECT " + """" + e + """" + "." + """" + a + """" + " AS " + """" + "COTA" + """" + ", " + """" + f + """" + "." + """" + b + """" + " AS " + """" + "TAM" + """" + ", " + """" + f + """" + "." + """" + b + """" + " ," + """" + f + """" + "." + """" + c + """" + ", " + """" + f + """" + "." + """" + d + """" + " FROM  " + """" + f + """" + " JOIN " + """" + e + """" + "ON" + """" + f + """" + "." + """" + b + """" + " = " + """" + e + """" + "." + """" + g + """" + " ORDER BY " + """" + "TAM" + """" + ", " + """" + b + """" + "", Conn, adOpenDynamic, adLockOptimistic
    End If
    Screen.MousePointer = vbHourglass
    Open CAMINHO For Output As #1
    Print #1, "IDENTIFICADOR;COORD_X;COORD_Y;COTA"
    Do While Not rs.EOF = True
        DoEvents                                                                'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
        If varGlobais.pararExecucao = True Then
            varGlobais.pararExecucao = False
            Screen.MousePointer = vbDefault
            FrmMain.Timer1.Enabled = False                                      'deshabilita o timer
            Close #1
            rs.Close
            Exit Sub
        End If
        FrmMain.sbStatusBar.Panels(2).Text = "Nó: " & rs!object_id              'mostra na barra de status o nó que está sendo exportado
        Print #1, rs!object_id & ";" & rs!X & ";" & rs!Y & ";" & rs!cota
        rs.MoveNext
    Loop
    Close #1
    rs.Close
    Screen.MousePointer = vbDefault
    FrmMain.Timer1.Enabled = False                                              'deshabilita o timer
    MsgBox "Arquivo exportado em " & CAMINHO & ".", vbInformation, "Exportação Concluída!"
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       Screen.MousePointer = vbDefault
       FrmMain.Timer1.Enabled = False                                           'deshabilita o timer
       ErroUsuario.Registra "FrmMain", "mnuExportLocalNos_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub

Private Sub mnuBitmap_Click()
   SalvaImagem "BMP"
End Sub

Private Sub mnuLocConsumidores_Click()
   frmEncontraConsumidor.Show 1
End Sub

Private Sub mnuPNG_Click()
   SalvaImagem "PNG"
End Sub

Private Sub mnuGIF_Click()
   SalvaImagem "GIF"
End Sub

Private Sub mnuJPG_Click()
   SalvaImagem "JPG"
End Sub



Private Sub mnuTIF_Click()
   SalvaImagem "TIF"
End Sub

Private Function SalvaImagem(tipo As String)
On Error GoTo saida





If FrmMain.Tag = "0" Then
   
   MsgBox "É necessário um mapa para realizar esta função.", vbInformation, ""
   GoTo saida
End If

Dim TP As ImageType
Dim strCaminho As String
   
   'TP = "a"
   
   'PREPARA O FILTRO DO COMPONENTE DE LOCALIZAÇÃO DE ARQUIVOS
   If tipo = "BMP" Then
      Me.cmmSalvaImg.Filter = ".BMP|*.BMP"
      TP = 1 '= BMP
   ElseIf tipo = "GIF" Then
      Me.cmmSalvaImg.Filter = ".GIF|*.GIF"
      TP = 2 '= GIF
   ElseIf tipo = "JPG" Then
      Me.cmmSalvaImg.Filter = ".JPG|*.JPG"
      TP = 3 '= JPG
   ElseIf tipo = "PNG" Then
      Me.cmmSalvaImg.Filter = ".PNG|*.PNG"
      TP = 4 '= PNG
   ElseIf tipo = "TIF" Then
      Me.cmmSalvaImg.Filter = ".TIF|*.TIF"
      TP = 5 '= TIF
   End If
   
   'ABRE O CPMPONENTE DE LOCALIZAÇÃO DE ARQUIVOS
   Me.cmmSalvaImg.ShowOpen
   
   strCaminho = Me.cmmSalvaImg.FileName
         
   'SE NÃO FOI DEFINIDO UM CAMINHO, SAI DA FUNÇÃO
   If Trim(strCaminho) = "" Then
      GoTo saida
   End If
   
   'SALVA O ARQUIVO DE ACORDO COM O TIPO ESCOLHIDO
   'If tipo = "JPG" Then
      If TCanvas.TCanvas.saveImageToFile(strCaminho, TP) Then
         MsgBox "Imagem salva com sucesso!"

      Else
         MsgBox "Falha ao salvar imagem!"
      End If
   
saida:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      
      PrintErro CStr(Me.Name), "Private Function SalvaImagem(tipo As String)", CStr(Err.Number), CStr(Err.Description), True
      Err.Clear

   End If
End Function

Private Sub mnuFixaIcone_Click()
On Error GoTo Trata_Erro

    With ActiveForm.TCanvas
        
        If .fixedPoint = False Then
            .fixedPoint = True
            
            mnuFixaIcone.Checked = True
            
            Call WriteINI("MAPA", "FIXAR_ICONE", "SIM", App.path & "\CONTROLES\GEOSAN.INI")
            
        Else
            .fixedPoint = False
            mnuFixaIcone.Checked = False
            
            Call WriteINI("MAPA", "FIXAR_ICONE", "NAO", App.path & "\CONTROLES\GEOSAN.INI")
            
        End If
        .plotView
        
    End With
Trata_Erro:

End Sub


Private Sub mnuDefEscala_Click()
On Error GoTo Trata_Erro

   Dim Scala As String
   
   Scala = InputBox("Informe o valor: ", "Definição de Escala")
   
   If IsNumeric(Scala) Then
      canvasScale = CDbl(Scala)
   Else
      MsgBox "Valor inválido.", vbInformation, ""
   End If


Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   MsgBox Err.Number & " - " & Err.Description
End If

End Sub


Private Sub mnuCalcularRede_Click()

   Dim MyComponents As String
   
   With ActiveForm.TCanvas
      If .getCurrentLayer = "WATERLINES" Then
         If .getSelectCount(2) > 0 Then
            If obtemRede(ActiveForm.TCanvas) = True Then
               openForm
            End If
         Else
            MsgBox "Selecione a rede a ser calculada", vbExclamation
         End If
      Else
         .setCurrentLayer "WATERLINES"
         MsgBox "Somente é possivel calcular a rede selecionando os tubos", vbExclamation
      End If
   End With

End Sub

Private Sub mnuChangePassword_Click()

    
    frmTrocaSenha.txtUsuario.Text = strUser
    frmTrocaSenha.txtUsuario.Locked = True
    frmTrocaSenha.Show 1

'   Set Sec = CreateObject("NSecurity.AppMode")
'   If Sec.OpenUserChangePwd(Conn, Usuario.UsrId) Then
'      MsgBox "Senha alterada com sucesso", vbInformation
'   End If
'   Set Sec = Nothing
'
End Sub

Private Sub mnuDeleteInc_Click()

   ActiveForm.CorrigeBug
   
End Sub

Private Sub mnuDeleteLineWater_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kdelete")
   
End Sub

Private Sub mnuDrawLineWater_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kdrawnetworkline")
   
End Sub

Private Sub mnuDrawPointInLineWater_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kinsertnetworknode")
   
End Sub

Private Sub mnuDrawRamal_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kdrawramal")
   
End Sub

Private Sub mnuEncontraTexto_Click()
    'frmCanvas.TimerSetWorld.Enabled = True
    frmEncontraTexto.Show 1
    
End Sub

Private Sub mnuExpAutoCad_Click()

    Dim frm As New FrmExport
   
    'Se nao houver canvas aberto não é possivel exportar nada...
    If FrmMain.Tag > 0 Then
        frm.Init Conn, ActiveForm.TCanvas, Me
    Else
        MsgBox "Não é possível exportar quando não existe uma área de trabalho do mapa.", vbInformation, "Atenção!"
    End If
   
    'Set frm = Nothing
   
End Sub

Private Sub mnuExpCRD_Click()
On Error GoTo mnuExpCRD_Click_err

   Shell App.path & "\Ferramentas\Exporte EPANet.exe", vbNormalFocus
   Exit Sub


mnuExpCRD_Click_err:
   MsgBox "Programa Exporte EPANet.exe não encontrado", vbExclamation, ""
   'Resume
End Sub

Private Sub mnuFileExit_Click()

   End
   Close
   
End Sub



Private Sub mnuGroups_Click()

   Set Sec = CreateObject("NSecurity.AppMode")
   Sec.OpenGroups Conn
   Set Sec = Nothing
   
End Sub

Private Sub mnuHelpAbout_Click()

   frmAbout.Show
   
End Sub

Private Sub mnuImagem_Click()

    'Se nao houver canvas aberto não é possivel exportar nada...
    If FrmMain.Tag > 0 Then
        With Cdl
           .FileName = ""
           .Filter = "Bitmap (*.bmp)|*.bmp | GIF (*.gif) | *.gif | JPG (*.jpg) | *.jpg | PNG (*.png) | *.png | TIF (*.tif) | *.tif"
           .ShowOpen
           If .FileName <> "" Then
              ActiveForm.TCanvas.saveImageToFile Cdl.FileName, .FilterIndex - 1
           End If
        End With
    Else
        MsgBox "Não é possível exportar quando não existe uma área de trabalho do mapa.", vbInformation, "Atenção!"
    End If
   
End Sub

Private Sub mnuImpCotas_Click()
   frmImportarCotas.Show 1
End Sub

Public Function Conecta()

Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
If frmCanvas.TipoConexao <> 4 Then

   
   
   TeImport1.Provider = typeconnection
   TeImport1.connection = Conn
   
   TeDatabase1.Provider = typeconnection
   TeDatabase1.connection = Conn

Else




'Set teac = TeAcXConnection1
If frmCanvas.POSTB <> 10 Then
mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
frmCanvas.FunDecripta (Senha)
decriptada = frmCanvas.Senha

'If TeAcXConnection1.Open = False Then

 TeAcXConnection1.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA
 ' End If

 
 TeImport1.Provider = typeconnection
   TeImport1.connection = TeAcXConnection1.objectConnection_
   
   TeDatabase1.Provider = typeconnection
   TeDatabase1.connection = TeAcXConnection1.objectConnection_
  frmCanvas.POST2B (10)
'   Else
   
  '' a1.Provider = typeconnection
  ' a1.connection = teac.objectConnection_
   
  ' a2.Provider = typeconnection
  ' a2.connection = teac.objectConnection_
   End If
End If

End Function

Private Sub mnuImportDXF_Click()
   

  ' Set a1 = TeImport1
 ' Set a2 = TeDatabase1
   
Dim frm As New frmImportDxf

Conecta

   
        frm.Init Conn, TeImport1, TeDatabase1
   Set frm = Nothing
   'changeSelIntersectionPoint
End Sub

Private Sub mnuImportSIG_Click()
   
   Dim frm As New frmImportFile
   
   'Dim a1 As TeImport
  ' Dim a2 As TeDatabase
  ' Set a1 = TeImport1
 ' Set a2 = TeDatabase1
   

   
Conecta
   

      frm.Init Conn, TeImport1, TeDatabase1
   Set frm = Nothing
   
End Sub

Private Sub mnuInsertDocs_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kinsertdoc")
   
End Sub

Private Sub mnuInsertLabel_Click()

   FrmCreatTextForLayer.Init
   
End Sub

Private Sub mnuLayers_Click()
   
   Dim rs As ADODB.Recordset
Dim a, b, c As String
   Set rs = New ADODB.Recordset
a = "USRLOG"
b = "USRFUN"
c = "SYSTEMUSERS"



   If frmCanvas.TipoConexao <> 4 Then
   rs.Open "SELECT USRLOG, USRFUN FROM SYSTEMUSERS WHERE USRLOG = '" & strUser & "' ORDER BY USRLOG", Conn, adOpenDynamic, adLockReadOnly
   Else
   rs.Open "SELECT " + """" + a + """" + "," + """" + b + """" + " FROM " + """" + c + """" + " WHERE " + """" + a + """" + " = '" & strUser & "' ORDER BY " + """" + a + """" + "", Conn, adOpenDynamic, adLockOptimistic
   End If
   If rs.EOF = False Then
      If rs!UsrFun = 4 Then  'VISUALIZADOR
         
         frmLoginTema.Show 1
                  
         'pctSfondo.Visible = True
      Else
      
         pctSfondo.Visible = Not mnuLayers.Checked
         mnuLayers.Checked = Not mnuLayers.Checked
      
      End If
   End If
   rs.Close
   

   
End Sub

Private Sub mnuLoadAttributeByReference_Click()

   mnuLoadAttributeByReference.Checked = Not mnuLoadAttributeByReference.Checked
   
End Sub

Private Sub mnuManufacters_Click()

   FrmManufactures.Show
   
End Sub

Private Sub mnuMinusZoom_Click()
 
   tbToolBar_ButtonClick tbToolBar.Buttons("kzoomin")
   
End Sub

Private Sub mnuMoreZoom_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kzoomout")
   
End Sub

Private Sub mnuMove_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kpan")
   
End Sub

Private Sub mnuMovePointWithLines_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kmovenetworknode")
   
End Sub

Private Sub mnuMultProperteis_Click()

   mnuMultProperteis.Checked = Not mnuMultProperteis.Checked
   
End Sub
'Indica se é para calcular ou não a cota Z do nó enquanto estiver desenhando. Caso não exista o layer MDT ainda esta função é muito útil
'
Private Sub mnuCalculaZNo_Click()
    mnuCalculaZNo.Checked = Not mnuCalculaZNo.Checked
    If mnuCalculaZNo.Checked = True Then            'Indica se a aplicação deve calcular ou não o Z do nó enquanto o usuário está desenhando a rede
        varGlobais.deveCalcularZNo = True
    Else
        varGlobais.deveCalcularZNo = False
    End If
End Sub



Private Sub mnuRecompose_Click()
   tbToolBar_ButtonClick tbToolBar.Buttons("krecompose")
End Sub

Private Sub mnuRamaisAgua_Click()
    frmIndicProdutRamaisAgua.Show 1
End Sub

Private Sub mnuRedesAgua_Click()
    
    frmIndicProdutRedesDeAgua.TipoRede = "AGUA"
    frmIndicProdutRedesDeAgua.Show 1
    
    
End Sub

Private Sub mnuRedesEsgoto_Click()
    
    frmIndicProdutRedesDeAgua.TipoRede = "ESGOTO"
    frmIndicProdutRedesDeAgua.Show 1
    
End Sub

Private Sub mnuRedoView_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kredoview")
   
End Sub

Private Sub mnuRelComponentesAgua_Click()

   GeraRelatorioHtm ComponentsRede, "watercomponents"
   
End Sub

Private Sub mnuRelComponentesEsgoto_Click()

   GeraRelatorioHtm ComponentsRede, "sewercomponents"
   
End Sub

Private Sub mnuRelComponentsWaterFilter_Click()

   GeraRelatorioHtm ComponentsRede, "watercomponents", True
   
End Sub

Private Sub mnuRelRegistros_Click()

   GeraRelatorioHtm RegistrosEstadoEstado, ""
   
End Sub

Private Sub MnuRelSl_Click()

   OpenReport "sewer"
   
End Sub

Private Sub MnuRelWl_Click()

   OpenReport "water"
   
End Sub

Private Sub mnuRemoverPlano_Click()

   FrmRemoverPlano.Show vbModal
   
End Sub

Private Sub mnuSELECT_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kselection")
   
End Sub
' Redireciona para outro banco de dados geográfico, modificando a configuração do GEOSAN.INI
'
'
'
Private Sub mnuSELECTDatabase_Click()
    On Error GoTo Trata_Erro
    Dim cn As ADODB.connection, nC As Object
    
    'FrmConnection.Show (1)
    Set nC = CreateObject("NexusConnection.App")
    If nC.appNewRegistry(App.EXEName, cn) Then
        Conn.Close
        'Set Conn = cn
        'Shell App.path & "\" & App.EXEName & ".exe"
        MsgBox "Banco de dados redirecionado com sucesso." & Chr(13) & Chr(13) & "Reinicie o sistema para ativar.", vbInformation
        'Set cn = Nothing
        End
    End If
    Exit Sub
      
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "FrmMain", "mnuSELECTDatabase_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub

Private Sub mnuSuppliers_Click()

   FrmSuppliers.Show
   
End Sub



Private Sub mnuTypes_Click()

   FrmSelectTypes.Init
   
End Sub

Private Sub mnuUndoView_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kundoview")
   
End Sub
'Exporta as redes do GeoSan
'
'
Private Sub mnuExporta_GeoSan_Click()
    On Error GoTo Trata_Erro
    Dim retorno As Boolean
    Dim conexao As New ADODB.connection
    Dim diretorio As String                                                                                 'diretório para onde serão exportados os arquivos shape
    Dim prefixoArquivo As String                                                                            'prefixo com as datas, dos arquivos shp que serão exportados
    Dim nomeCompleto As String
    Dim nomeExportar As String
    Dim existemConsumidoresParaExportar As Boolean
  
    varGlobais.pararExecucao = False                            'indica que iniciará sem sem informar que deverá parar a execução
    diretorio = arquivo.SelecionaDiretorio
    If diretorio = "falhou" Then
        'MsgBox "Cancelada a seleção do diretório."
        Exit Sub
    End If
    prefixoArquivo = arquivo.prefixo
    nomeCompleto = diretorio + "\" + prefixoArquivo
    FrmMain.Timer1.Enabled = True                               'habilita o timer para permitir o usuário cancelar esta operação
    Screen.MousePointer = vbHourglass
    
    'exporta consumidores
    exp.InsereTabAtributoConsumidores
    exp.CriaTabelaConsumidores
    existemConsumidoresParaExportar = exp.InsereConsumidores    'insere os consumidores na tabela GS_CONSUMIDORES, para poder a partir dela exportar para shp
    If varGlobais.pararExecucao = True Then                     'usuário selecionou para parar tudo
        Exit Sub
    End If
    If existemConsumidoresParaExportar = True Then              'se a tabela GS_CONSUMIDORES foi preechida com todos os consumidores, exporta o shape.
        exp.AtivaExportacaoConsumidores
        conexao.Open Conn
        TeExport2.Provider = 1
        TeExport2.connection = conexao
        FrmMain.sbStatusBar.Panels(2).Text = "Criando shape de consumidores. Favor aguardar ..."            'mostra na barra de status o andamento da exportação
        nomeExportar = nomeCompleto & "gsConsumidores.shp"
        retorno = TeExport2.exportSHP(nomeExportar, "RAMAIS_AGUA", "GS_CONSUMIDORES")
        Screen.MousePointer = vbNormal
        If retorno Then
            'MsgBox "Exportação shape de ramais realizada com sucesso"
        Else
            MsgBox "Falha na exportação dos consumidores"
        End If
        conexao.Close
    Else
        ErroUsuario.Registra "FrmMain", "Exporta_Consumidores", CStr(Err.Number), CStr(Err.Description), False, glo.enviaEmails
    End If
    
    'exporta ramais
    exp.InsereTabAtributoRamais
    exp.CriaTabelaRamais
    exp.InsereRamais
    If varGlobais.pararExecucao = True Then                     'usuário selecionou para parar tudo
        Exit Sub
    End If
    exp.AtivaExportacaoRamais
    conexao.Open Conn
    TeExport2.Provider = 1
    TeExport2.connection = conexao
    FrmMain.sbStatusBar.Panels(2).Text = "Criando shape de ramais. Favor aguardar ..."                      'mostra na barra de status o andamento da exportação
    nomeExportar = nomeCompleto + "gsRamais.shp"
    retorno = TeExport2.exportSHP(nomeExportar, "RAMAIS_AGUA", "GS_RAMAIS")
    Screen.MousePointer = vbNormal
    If retorno Then
        'MsgBox "Exportação shape de ramais realizada com sucesso"
    Else
        MsgBox "Falha na exportação dos ramais."
    End If
    conexao.Close
        
    'prepara tabelas e atributos para exportar as redes
    exp.InsereTabAtributoRedes
    exp.CriaTabelaRedes
    exp.InsereRedes
    If varGlobais.pararExecucao = True Then                                                                 'usuário selecionou para parar tudo
        Exit Sub
    End If
    'exporta redes para o formato shape
    conexao.Open Conn
    TeExport2.Provider = 1
    TeExport2.connection = conexao
    FrmMain.sbStatusBar.Panels(2).Text = "Criando shape de redes. Favor aguardar ..."                       'mostra na barra de status o andamento da exportação
    nomeExportar = nomeCompleto + "gsRedes.shp"
    retorno = TeExport2.exportSHP(nomeExportar, "WATERLINES", "GS_REDES")
    Screen.MousePointer = vbNormal
    If retorno Then
        'MsgBox "Exportação shape de redes realizada com sucesso"
    Else
        MsgBox "Falha na exportação"
    End If
    conexao.Close
    
    'prepara tabelas e atributos para exportar os nós
    exp.InsereTabAtributoNos
    exp.CriaTabelaNos
    exp.InsereNos
    If varGlobais.pararExecucao = True Then                                                                 'usuário selecionou para parar tudo
        Exit Sub
    End If
    'exporta nós
    conexao.Open Conn
    TeExport2.Provider = 1
    TeExport2.connection = conexao
    exp.AtivaExportacaoNos
    FrmMain.sbStatusBar.Panels(2).Text = "Criando shape de nós. Favor aguardar ..."                         'mostra na barra de status o andamento da exportação
    nomeExportar = nomeCompleto + "gsNos.shp"
    retorno = TeExport2.exportSHP(nomeExportar, "WATERCOMPONENTS", "GS_NOS")
    exp.DesativaExportacaoNos
    Screen.MousePointer = vbNormal
    If retorno Then
        'MsgBox "Exportação dos nós realizada com sucesso"
    Else
        MsgBox "Falha na exportação"
    End If
    exp.AtivaRamaisGeoSan                                   'reativa te_representation, senão os ramais com os nós (ligações) não voltam a aparecer no GeoSan
    conexao.Close
    FrmMain.sbStatusBar.Panels(2).Text = "Exportação finalizada."
    FrmMain.Timer1.Enabled = False                          'deshabilita o timer
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        FrmMain.Timer1.Enabled = False                      'deshabilita o timer
        exp.AtivaRamaisGeoSan                               'reativa te_representation, senão os ramais com os nós (ligações) não voltam a aparecer no GeoSan
        ErroUsuario.Registra "FrmMain", "mnuExporta_GeoSan_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
' Atualiza os consumos médios nas ligações e distribui as demandas nos nós das redes
'
'
'
Private Sub mnuUpdate_Demand_Click()
    frmAtualizacaoConsumo.Show 1
End Sub

Private Sub mnuUsers_Click()

'   Set Sec = CreateObject("NSecurity.AppMode")
'   Sec.OpenUsers Conn
'   Set Sec = Nothing
   
   frmUserControle.Show 1
   'FrmUser.Show 1
   
End Sub



Private Sub mnuViewStatusBar_Click()

   sbStatusBar.Visible = Not mnuViewStatusBar.Checked
   mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
   
End Sub

Private Sub mnuViewToolbar_Click()

   tbToolBar.Visible = Not mnuViewToolbar.Checked
   mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
   
End Sub

Private Sub mnuWindowCascade_Click()

   Me.Arrange vbCascade
   
End Sub

Private Sub mnuWindowTileHorizontal_Click()

   Me.Arrange vbTileHorizontal
   
End Sub

Private Sub mnuWindowTileVertical_Click()

   Me.Arrange vbTileVertical
   
End Sub

Private Sub mnuZoom_Click()

   tbToolBar_ButtonClick tbToolBar.Buttons("kzoomarea")
   
End Sub

Private Sub OdImport_Click()
   
   Form1.Show 1
   
End Sub

Private Sub pctSfondo_Resize()

   SizeControls
   
End Sub

Private Sub MDIForm_Resize()

   SizeControls
   
End Sub
' Esta rotina controla a correta visualização do tamanho do gerenciador de propriedades, tree e itens perto dos mesmos
'
'
'
Public Sub SizeControls()
    On Error GoTo Trata_Erro
                                                                                        'a variável pctSfondo representa o lado esquerdo do gerenciador de propriedades
    With TabStrip1                                                                      'é a tab superior com as opções de tree e propriedades
        .Height = IIf(pctSfondo.Height < .Top, 100, pctSfondo.Height - .Top)
        .Width = IIf(pctSfondo.Width < .Left, 100, pctSfondo.Width - .Left)
    End With
    With Manager1                                                                       'gerenciador de propriedades
        .Width = IIf(pctSfondo.Width < .Left, 10, pctSfondo.Width - (.Left + 100))
        .Height = IIf(pctSfondo.Height < .Top, 10, pctSfondo.Height - (.Top + 100))
        .Resize pctSfondo.Width - 400, pctSfondo.Height - 1400
        .Top = 1340
        .Left = 300
    End With
    With ViewManager1                                                                   'gerenciador de tree
        .Width = IIf(pctSfondo.Width < .Left, 10, pctSfondo.Width - (.Left + 100))
        .Height = IIf(pctSfondo.Height < .Top, 10, pctSfondo.Height - (.Top + 100))
        .Top = 1350
        .Left = 300
    End With
    picSplitter.Height = pctSfondo.Height                                               'separador das duas colunas do gerenciador de propriedades
    imgSplitter.Height = pctSfondo.Height                                               'outro separador
    cmdClose.Left = pctSfondo.Width - 300                                               'ícone X de fechar
    cmdClose.Top = pctSfondo.Top - 350
    FrameEscala.Width = pctSfondo.Width - 300                                           'label da escala de visualização
    txtEscala.Width = pctSfondo.Width - 2350                                            'texto da escala de visualização
    Exit Sub
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = 380 Then
        Exit Sub
    Else
        PrintErro CStr(Me.Name), "Public Sub SizeControls()", CStr(Err.Number), CStr(Err.Description), True
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

On Error GoTo Trata_Erro
    
    
   If FrmMain.ViewManager1.mConn.State = 1 Then
      FrmMain.ViewManager1.mConn.Close
   End If
    
    If Conn.State = 1 Then
      Conn.Close
    End If
    

    Set Conn = Nothing
    'LoozeXP1.EndWinXPCSubClassing
    End
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
    
         'PrintErro CStr(Me.Name), "Private Sub MDIForm_Unload", CStr(Err.Number), CStr(Err.Description), True
         
    
    End If

End Sub
' Opção do menu de ícones que abre uma janela de desenho de mapa.
' Entra nesta rotina quando o usuário seleciona a ícone de que deseja uma nova janela de desenho
'
'
'
Private Sub mnuOpen_Click()
    Set TCanvas = New frmCanvas
    TCanvas.Init Conn, usuario.UseName
End Sub





Private Sub TabStrip1_Click()

   If TabStrip1.SelectedItem.index = 2 Then
      Manager1.Visible = True
      'Tv.Visible = False
      ViewManager1.Visible = False
   Else
      Manager1.Visible = False
      'Tv.Visible = True
      ViewManager1.Visible = True
   End If
   
End Sub

' Monitoramento dos eventos da barra de ícones. Evento de clique na barra de ferramentas
' Fica aguardando o usuário selecionar uma das ícones na barra de menu de ícones
' Caso a janela de desenho (canvas) não estiver aberto, não faz nada ainda
' Caso esteja aberta a seleção da ícone
' Caso selecione que é para abrir uma nova janela de desenho (canvas), abre a mesma
'
' Button - botão que foi selecionado pelo usuário
'
Public Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'If blnMonitorar = True Then
    On Error GoTo Trata_Erro
    Select Case Button.key
        Case "knew", ""
           mnuOpen_Click
        Case Else
           If Not ActiveForm Is Nothing Then
              If ActiveForm.Name = "frmCanvas" Then         'se o canvas de mapas está na tela
                  ActiveForm.Tb_SELECT Button.key           'indica a ativação do botão que foi selecionado
              End If
           End If
    End Select
    Exit Sub
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = 91 Then
        Err.Clear
        Exit Sub
    Else
        PrintErro CStr(Me.Name), "Não está encontrando o botão a ser selecionado em Public Sub tbToolBar_ButtonClick()", CStr(Err.Number), CStr(Err.Description), True
    End If
End Sub
'Foi comentado pois não estava sendo utilizado.
'Também não chamava o frm1TePrinter



'Private Sub TePrinter_Click()
'
'   'frmPrint As New frm1TePrinter
'
'   frmTePrinter.Show 1
'
'
'End Sub
' Configura um timer para caso o usuário selecione a tecla ESC ele pare a execução
'
' varGlobais.pararExecucao - contem a informação que deve ser configurada na rotina que deseja-se cancelar a execução. Lembrando-se de colocar um Doevents antes. Veja o exemplo abaixo
' o intervalo do timer está definido no MDIForm_Load
'
'DoEvents                                                            'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
'If varGlobais.pararExecucao = True Then
'    varGlobais.pararExecucao = False
'    Screen.MousePointer = vbNormal
'    Exit Sub
'End If
'
' O timer deve ser habilitado antes de entrar na rotina que requer cálculo intensivo. Veja o exemplo abaixo:
'FrmMain.Timer1.Enabled = True                               'habilita o timer
'
Private Sub Timer1_Timer()
    If GetAsyncKeyState(VK_ESCAPE) Then
        MsgBox ("Comando cancelado.")
        varGlobais.pararExecucao = True
    End If
End Sub

Private Sub txtEscala_KeyPress(KeyAscii As Integer)

'AO RECEBER UM COMANDO ENTER, É FORÇADO UM LOST_FOCUS

If KeyAscii = 13 Then
   
   txtEscala_LostFocus
   
End If

End Sub

Private Sub txtEscala_LostFocus()

   If IsNumeric(Me.txtEscala.Text) = True Then
 
      canvasScale = CDbl(Me.txtEscala.Text)
   
   End If
   
End Sub


Private Sub ViewManager1_onReset(ViewName As String)
On Error GoTo Trata_Erro

   Dim a As Integer, LayerNameStr As String
   
   For a = 1 To tbToolBar.Buttons.count
      If tbToolBar.Buttons.Item(a).Style = tbrCheck Then
         tbToolBar.Buttons(a).value = tbrUnpressed
      End If
   Next
   
   
   strLayerAtivo = TCanvas.TCanvas.getCurrentLayer
   
   With Me.ActiveForm.TCanvas
      For a = 0 To .getLayersToSnapCount() - 1
         If .getLayerToSnap(a, LayerNameStr) = 1 Then
            .removeLayerToSnap LayerNameStr
         End If
      Next
   End With
   
   tbToolBar.Buttons("kselection").value = tbrPressed
   
   Me.ActiveForm.Caption = "Vista: " & ViewName
   sbStatusBar.Panels(1).Text = "Modo de seleção: Selecione um objeto do plano referente ao tema ativo"

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        
        PrintErro CStr(Me.Name), "Private Sub ViewManager1_onReset", CStr(Err.Number), CStr(Err.Description), True
        
        
    End If
End Sub

Private Sub mnuCarregaPoligono_Click()


On Error GoTo Trata_Erro

   Dim rs As ADODB.Recordset
   Dim strsql As String
   
   blnPoligonoVirtual = False
   Dim a, b, c, d, e, f As String
   Set rs = New ADODB.Recordset
a = "layer_id"
b = "geom_table"
c = "te_representation"
d = "geom_type"
e = "te_layer"
f = "name"




     If frmCanvas.TipoConexao <> 4 Then
   strsql = "SELECT TL.LAYER_ID,TL.NAME,TR.GEOM_TABLE FROM TE_LAYER TL INNER JOIN TE_REPRESENTATION TR ON TL.LAYER_ID = TR.LAYER_ID WHERE TR.GEOM_TYPE = 1 AND TL.NAME = '" & strLayerAtivo & "'"
   Else
   strsql = "SELECT " + """" + e + """" + "." + """" + a + """" + "," + """" + e + """" + "." + """" + f + """" + "," + """" + c + """" + "." + """" + b + """" + " FROM " + """" + e + """" + " INNER JOIN " + """" + c + """" + " ON " + """" + e + """" + "." + """" + a + """" + " = " + """" + c + """" + "." + """" + a + """" + " WHERE " + """" + c + """" + "." + """" + d + """" + " = '1' AND " + """" + e + """" + "." + """" + f + """" + "= '" & strLayerAtivo & "'"
   End If
   
   
   
'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", strsql, App.path & "\DEBUG.INI"
 
   rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
   
   If rs.EOF = False Then
   
a = "object_id"
b = LCase(rs!GEOM_TABLE)



      If frmCanvas.TipoConexao <> 4 Then
      strsql = "SELECT COUNT(OBJECT_ID) AS " + """" + "QTD" + """" + " FROM " & rs!GEOM_TABLE
      Else
      strsql = "SELECT COUNT(" + """" + a + """" + ") AS " + """" + "QTD" + """" + " FROM " + """" + b + """" + ""
      End If
      rs.Close
      rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
      If rs!qtd = 0 Then
         MsgBox "O plano ativo não possui polígonos.", vbInformation, ""
         rs.Close
         Exit Sub
      End If
      rs.Close
      Me.MousePointer = vbHourglass
      frmAtualizarSetores.Show 1
      Me.MousePointer = vbDefault
   
   Else
      MsgBox "O plano ativo não possui polígonos.", vbInformation, ""
      rs.Close
   End If
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
         
         PrintErro CStr(Me.Name), "mnuCarregaPoligono_Click()", CStr(Err.Number), CStr(Err.Description), True
         
        
    End If

End Sub
' Permite o usuário desenhar um polígono que irá selecionar no mapa várias redes e consumidores para relatórios e exportação para o EPANET
'
'
''
Private Sub mnuDesenhaPoligono_Click()
    If ActiveForm.Name = "frmCanvas" Then
        ActiveForm.Tb_SELECT "mnuPoligono"
    End If
End Sub
