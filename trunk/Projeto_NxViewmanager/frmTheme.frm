VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTheme 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Propriedades do Tema"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFiltro 
      Caption         =   "Atributos"
      Enabled         =   0   'False
      Height          =   195
      Left            =   285
      TabIndex        =   32
      Top             =   2850
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   1815
      Top             =   5790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   345
      Left            =   4905
      TabIndex        =   39
      Top             =   5850
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   195
      TabIndex        =   38
      Top             =   5850
      Width           =   945
   End
   Begin VB.Frame Frame9 
      Caption         =   "Escala de Visualização"
      Height          =   735
      Left            =   165
      TabIndex        =   33
      Top             =   4980
      Width           =   5730
      Begin VB.TextBox txtMin 
         Height          =   315
         Left            =   1110
         TabIndex        =   35
         Text            =   "0"
         Top             =   300
         Width           =   1275
      End
      Begin VB.TextBox txtMax 
         Height          =   315
         Left            =   4260
         TabIndex        =   34
         Text            =   "0"
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Máxima"
         Height          =   285
         Left            =   3390
         TabIndex        =   37
         Top             =   330
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Mínima"
         Height          =   285
         Left            =   270
         TabIndex        =   36
         Top             =   330
         Width           =   1065
      End
   End
   Begin VB.Frame FrameFiltro 
      Caption         =   "Filtros Ativos"
      Height          =   1815
      Left            =   120
      TabIndex        =   31
      Top             =   2640
      Width           =   5715
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   345
         Left            =   4500
         TabIndex        =   74
         Top             =   255
         Width           =   1005
      End
      Begin VB.TextBox txtDataFim 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4290
         MaxLength       =   8
         TabIndex        =   72
         Top             =   1845
         Width           =   1020
      End
      Begin VB.TextBox txtDataInicio 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   70
         Top             =   1845
         Width           =   1020
      End
      Begin VB.ComboBox cboColunas 
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   54
         Top             =   720
         Width           =   1995
      End
      Begin VB.ComboBox cboFiltro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3690
         TabIndex        =   56
         Top             =   720
         Width           =   1845
      End
      Begin VB.ComboBox cboOperador 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmTheme.frx":0000
         Left            =   2205
         List            =   "frmTheme.frx":0002
         TabIndex        =   55
         Top             =   720
         Width           =   1365
      End
      Begin VB.ComboBox cboFiltro2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3690
         TabIndex        =   59
         Top             =   1155
         Width           =   1845
      End
      Begin VB.ComboBox cboOperador2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2205
         TabIndex        =   58
         Top             =   1155
         Width           =   1365
      End
      Begin VB.ComboBox cboColunas2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   57
         Top             =   1140
         Width           =   1995
      End
      Begin VB.ComboBox cboFiltroSub 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   53
         Top             =   270
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.ComboBox cboOperadorSub 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   270
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.ComboBox cboColunasSub 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   270
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CheckBox chkFiltraData 
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   69
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Data de Desenho DD/MM/AA "
         Enabled         =   0   'False
         Height          =   435
         Left            =   345
         TabIndex        =   75
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label6 
         Caption         =   "Fim"
         Enabled         =   0   'False
         Height          =   270
         Left            =   3930
         TabIndex        =   73
         Top             =   1860
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Início"
         Enabled         =   0   'False
         Height          =   270
         Left            =   1995
         TabIndex        =   71
         Top             =   1875
         Width           =   495
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   105
         X2              =   5490
         Y1              =   1635
         Y2              =   1635
      End
   End
   Begin TabDlg.SSTab mtab 
      Height          =   2295
      Left            =   165
      TabIndex        =   0
      Top             =   75
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   503
      TabMaxWidth     =   2117
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Polígonos"
      TabPicture(0)   =   "frmTheme.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPol(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPol(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Linhas"
      TabPicture(1)   =   "frmTheme.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fralines(0)"
      Tab(1).Control(1)=   "fralines(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Pontos"
      TabPicture(2)   =   "frmTheme.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frapoints(1)"
      Tab(2).Control(1)=   "frapoints(0)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Textos"
      TabPicture(3)   =   "frmTheme.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fratexts(1)"
      Tab(3).Control(1)=   "fratexts(0)"
      Tab(3).ControlCount=   2
      Begin VB.Frame fratexts 
         Caption         =   "Exemplo:"
         Height          =   1695
         Index           =   0
         Left            =   -71160
         TabIndex        =   25
         Top             =   450
         Width           =   1755
         Begin RichTextLib.RichTextBox txtExample 
            Height          =   1155
            Left            =   180
            TabIndex        =   62
            Top             =   300
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   2037
            _Version        =   393217
            TextRTF         =   $"frmTheme.frx":0074
         End
      End
      Begin VB.Frame frapoints 
         Caption         =   "Configuração"
         Height          =   1695
         Index           =   0
         Left            =   -74820
         TabIndex        =   20
         Top             =   450
         Width           =   3585
         Begin VB.CommandButton cmdColorTransp 
            Height          =   315
            Left            =   1050
            Picture         =   "frmTheme.frx":00F8
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   1140
            Width           =   705
         End
         Begin VB.TextBox txtAngle 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1890
            TabIndex        =   64
            Text            =   "0"
            Top             =   1110
            Width           =   645
         End
         Begin VB.CommandButton cmdIcone 
            Caption         =   "Inserir"
            Height          =   315
            Left            =   120
            TabIndex        =   63
            Top             =   1140
            Width           =   825
         End
         Begin VB.CheckBox chkVisibledPoint 
            Caption         =   "Visível"
            Height          =   255
            Left            =   2700
            TabIndex        =   48
            Top             =   1170
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.TextBox txtPointWidth 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1050
            TabIndex        =   46
            Text            =   "Text2"
            Top             =   510
            Width           =   615
         End
         Begin VB.ComboBox cboPointStyle 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   510
            Width           =   1515
         End
         Begin VB.CommandButton cmdCollorPoint 
            Height          =   315
            Left            =   90
            Picture         =   "frmTheme.frx":063A
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   510
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Transp."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   1080
            TabIndex        =   68
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Ícone"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   270
            TabIndex        =   66
            Top             =   930
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Angulo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   1920
            TabIndex        =   65
            Top             =   900
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Estilo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   1950
            TabIndex        =   24
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Tamanho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1020
            TabIndex        =   23
            Top             =   300
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Contorno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   22
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.Frame frapoints 
         Caption         =   "Exemplo:"
         Height          =   1695
         Index           =   1
         Left            =   -71160
         TabIndex        =   19
         Top             =   450
         Width           =   1755
         Begin VB.Shape shpPoint 
            BorderStyle     =   0  'Transparent
            FillStyle       =   0  'Solid
            Height          =   315
            Left            =   810
            Shape           =   3  'Circle
            Top             =   750
            Width           =   195
         End
      End
      Begin VB.Frame fralines 
         Caption         =   "Exemplo:"
         Height          =   1695
         Index           =   1
         Left            =   -71160
         TabIndex        =   18
         Top             =   450
         Width           =   1755
         Begin VB.Line Line1 
            X1              =   210
            X2              =   1560
            Y1              =   930
            Y2              =   930
         End
      End
      Begin VB.Frame fralines 
         Caption         =   "Configuração"
         Height          =   1695
         Index           =   0
         Left            =   -74820
         TabIndex        =   12
         Top             =   450
         Width           =   3585
         Begin VB.CheckBox chkVisibledLine 
            Caption         =   "Visível"
            Height          =   255
            Left            =   2310
            TabIndex        =   49
            Top             =   540
            Value           =   1  'Checked
            Width           =   885
         End
         Begin VB.TextBox txtLineWidth 
            Height          =   345
            Left            =   1050
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   510
            Width           =   975
         End
         Begin VB.CommandButton cmdCollorLine 
            Height          =   315
            Left            =   150
            Picture         =   "frmTheme.frx":0B7C
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   510
            Width           =   825
         End
         Begin VB.ComboBox cboLineStyle 
            Height          =   315
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1140
            Width           =   3225
         End
         Begin VB.Label Label1 
            Caption         =   "Contorno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   17
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Largura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1050
            TabIndex        =   16
            Top             =   300
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Estilo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   150
            TabIndex        =   15
            Top             =   930
            Width           =   2595
         End
      End
      Begin VB.Frame fraPol 
         Caption         =   "Configuração"
         Height          =   1695
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   450
         Width           =   3585
         Begin VB.CheckBox chkVisibledPolyguns 
            Caption         =   "Visível"
            Height          =   255
            Left            =   2610
            TabIndex        =   50
            Top             =   540
            Value           =   1  'Checked
            Width           =   885
         End
         Begin VB.TextBox txtPolBorderWidth 
            Height          =   315
            Left            =   1050
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   510
            Width           =   675
         End
         Begin VB.ComboBox cboPolStyleFill 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1170
            Width           =   1545
         End
         Begin VB.ComboBox cboPolStyleBorder 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   510
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.CommandButton cmdCollorPol 
            Height          =   315
            Left            =   150
            Picture         =   "frmTheme.frx":10BE
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1200
            Width           =   825
         End
         Begin VB.CommandButton cmdCollorPolBor 
            Height          =   315
            Left            =   150
            Picture         =   "frmTheme.frx":1600
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   510
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Estilo Fundo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1890
            TabIndex        =   11
            Top             =   930
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Estilo Borda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1890
            TabIndex        =   10
            Top             =   300
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "Largura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1050
            TabIndex        =   9
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Fundo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   210
            TabIndex        =   6
            Top             =   990
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Contorno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.Frame fraPol 
         Caption         =   "Exemplo:"
         Height          =   1695
         Index           =   1
         Left            =   3840
         TabIndex        =   1
         Top             =   450
         Width           =   1755
         Begin VB.Shape polyguns1 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   1245
            Left            =   210
            Shape           =   4  'Rounded Rectangle
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame fratexts 
         Caption         =   "Configuração"
         Height          =   1695
         Index           =   1
         Left            =   -74820
         TabIndex        =   26
         Top             =   450
         Width           =   3585
         Begin VB.CommandButton cmdTextCollor 
            Height          =   315
            Left            =   1140
            Picture         =   "frmTheme.frx":1B42
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   1140
            Width           =   825
         End
         Begin VB.CheckBox chkVisibledText 
            Caption         =   "Visível"
            Height          =   255
            Left            =   2610
            TabIndex        =   47
            Top             =   1200
            Value           =   1  'Checked
            Width           =   885
         End
         Begin VB.CheckBox chkFonteItalic 
            Caption         =   "Itálico"
            Height          =   315
            Left            =   210
            TabIndex        =   45
            Top             =   1230
            Width           =   975
         End
         Begin VB.CheckBox chkFonteBold 
            Caption         =   "Negrito"
            Height          =   315
            Left            =   210
            TabIndex        =   44
            Top             =   870
            Width           =   975
         End
         Begin VB.TextBox txtFonteSize 
            Height          =   315
            Left            =   210
            TabIndex        =   41
            Text            =   "Text2"
            Top             =   510
            Width           =   615
         End
         Begin VB.TextBox txtFonteStyle 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   510
            Width           =   1815
         End
         Begin VB.CommandButton cmdFonte 
            Height          =   315
            Left            =   3090
            Picture         =   "frmTheme.frx":2084
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   510
            Width           =   345
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Cor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   1170
            TabIndex        =   61
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Tamanho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   150
            TabIndex        =   29
            Top             =   300
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Estilo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1170
            TabIndex        =   28
            Top             =   300
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frmTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private tvm As Object, ThemeName As String, Confirm As Boolean
Private rs As ADODB.Recordset, NameIcon As String, InserirIcon As Boolean, ColorTransp As OLE_COLOR
Private blnMudaFiltro As Boolean
Dim rs2 As New ADODB.Recordset
Dim intTema As Integer                          'número do tema selecionado no menu da direita do GeoSan
Dim ThemeName2 As String                        'tema que está ativo na vista
Public LayerAtivo As String
Dim conexao As New ADODB.Connection
Dim man2 As Object
' Esta função é executada quando o usuário seleciona que deseja alterar as propriedades de visualização de um Tema
' Ela lê na tabela NXGS_FILT_TEMA os filtros anteriormente selecionados pelo usuário e então preenche na caixa de
' diálogo
'
Public Function Init(mtvm As Object, mtheme As String, mLayerName As String) As Boolean
    On Error GoTo Trata_Erro
    
    Dim vetor As Variant
    Dim str As String
    'LoozeXP1.InitSubClassing
    Call SaveLoadGlobalData("C:\Arquivos de programas\GeoSan" + "\controles\variaveisGlobais.txt", False) 'recupera as variáveis globais da aplicação principal do GeoSan
    Confirm = False
    Set tvm = mtvm
    LoadCboLine cboLineStyle
    LoadCboLine cboPolStyleBorder
    LoadCboPolygon cboPolStyleFill
    LoadCboPoints cboPointStyle
    Representarion_Visibled mtheme, tvm.getActiveView
    txtMin.Text = tvm.getMinScale(tvm.getActiveView, mtheme)
    txtMax.Text = tvm.getMaxScale(tvm.getActiveView, mtheme)
    Me.cboColunas.Clear
    Me.cboColunas2.Clear
    Me.cboFiltro.Clear
    Me.cboFiltro2.Clear
    cboOperador.Clear
    cboOperador.AddItem "Igual"
    cboOperador.AddItem "Maior"
    cboOperador.AddItem "Menor"
    cboOperador.AddItem "Diferente"
    cboOperador2.Clear
    cboOperador2.AddItem "Igual"
    cboOperador2.AddItem "Maior"
    cboOperador2.AddItem "Menor"
    cboOperador2.AddItem "Diferente"
    ThemeName = mtheme
    'LoadThemeFilter
    blnMudaFiltro = False
    Close #3
    Open glo.diretorioGeoSan + "\CONTROLES\FTema.txt" For Input As #3  ' Abre o arquivo que contém a lista com todos os temas existentes e filtros where da vista atual do usuário logado
    intTema = 0
    strCmdFiltro = ""
    'Localiza o número (intTema) do tema em que o usuário selecionou a direita na listas de temas do GeoSan
    Do While Not EOF(3)
        Line Input #3, str
        vetor = Split(str, ";")
        If CStr(vetor(1)) = ThemeName Then
            intTema = vetor(0)              'achou o tema selecionado pelo usuário
            Exit Do
        End If
        'MsgBox vetor(0) & " É O NÚMERO THEME_ID QUE IDENTIFICA O LAYER E É FEITO O SELECT"
        'MsgBox vetor(1) & " É O NOME DO LAYER"
        ' vetor(2) 'É O COMANDO DO FILTRO
    Loop
    Close #3
    Dim rs As New ADODB.Recordset
    '    Dim rs2 As New ADODB.Recordset
    '    rs2.Open "SELECT * FROM NXGS_FILT_TEMA", conn, adOpenKeyset, adLockOptimistic
    '    If rs2.EOF = True Then
    '        Rs.Open "SELECT * FROM TE_THEME", conn, adOpenKeyset, adLockOptimistic
    '        If Rs.EOF = False Then
    '            Do While Not Rs.EOF = True
    '                rs2.AddNew
    '                rs2!theme_id = Rs!theme_id
    ''                rs2!FILT_1 = rs!FILT_1
    ''                rs2!FILT_2 = rs!FILT_2
    ''                rs2!FILT_3 = rs!FILT_3
    '                rs2.Update
    '                Rs.MoveNext
    '            Loop
    '        End If
    '        Rs.Close
    '    End If
    '    rs2.Close
    'mudado 6-1-2011
    If intTema = 0 Then                 'desabilita a opção de filtro, pois nenhum tema foi encontrado, ou seja o tema selecionado não foi encontrato em Ftema.txt
        'Me.cmdModificar.Enabled = False
    Else
        'Agora é rodada uma querie junto ao banco de dados para descobrir a que layer pertence o tema que o usuário selecionou para alterar
        If TypeConn <> 4 Then
            rs.Open "SELECT NAME FROM TE_LAYER WHERE LAYER_ID = (SELECT LAYER_ID FROM TE_THEME WHERE THEME_ID = " & intTema & ")", conn, adOpenKeyset, adLockReadOnly
        Else
            rs.Open "SELECT " + """" + e + """" + " FROM " + """" + a + """" + " WHERE " + """" + c + """" + " = (SELECT " + """" + c + """" + " FROM " + """" + b + """" + " WHERE " + """" + d + """" + " = '" & intTema & "')", conn, adOpenDynamic, adLockOptimistic
        End If
        If rs.EOF = False Then
            LayerAtivo = rs!Name        'descobriu o nome do layer que o usuário selecionou
        End If
        rs.Close
        ThemeName2 = tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName)
        If ThemeName2 = "SEWERCOMPONENTS" Or ThemeName2 = "SEWERLINES" Then
            'cboColunas2.Visible = False
            'cboOperador2.Visible = False
            'cboFiltro2.Visible = False
            Me.chkFiltraData.Enabled = True
            chkFiltraData.Visible = True
            Line2.Visible = True
            Label7.Visible = True
            Label6.Visible = True
            Label5.Visible = True
            txtDataInicio.Visible = True
            txtDataFim.Visible = True
        End If
        If ThemeName2 = "WATERLINES" Or ThemeName2 = "WATERCOMPONENTS" Then
            Line2.Visible = False
            chkFiltraData.Visible = False
            Label7.Visible = False
            Label6.Visible = False
            Label5.Visible = False
            txtDataInicio.Visible = False
            txtDataFim.Visible = False
            Line2.Visible = False
            chkFiltraData.Visible = True
            Label7.Visible = True
            Label6.Visible = True
            Label5.Visible = True
            txtDataInicio.Visible = True
            txtDataFim.Visible = True
        End If
        'Carrega nos campos os filtros ja existentes
        Dim aa As String
        Dim bb As String
        aa = "NXGS_FILT_TEMA"
        bb = "THEME_ID"
        ' MsgBox "SELECT * FROM " + """" + aa + """" + " WHERE " + """" + bb + """" + " = '" & intTema & "'"
        'Prepara uma pesquisa para todos os filtros criados pelos usuários. Isso é feito pois precisamos preencher a caixa de diálogo com os filtros que já foram
        'entrados pelo usuário. Para isso ele roda a querie e procura pelo tema selecionado (intTema) os filtros existentes na tabela NXGS_FILT_TEMA
        If TypeConn <> 4 Then
            rs.Open "SELECT * FROM NXGS_FILT_TEMA WHERE THEME_ID = " & intTema, conn, adOpenKeyset, adLockOptimistic
        Else
            rs.Open "SELECT * FROM " + """" + aa + """" + " WHERE " + """" + bb + """" + " = '" & intTema & "'", conn, adOpenDynamic, adLockOptimistic
        End If
        'Preenche agora na caixa de diálogo com os dados dos filtros existentes para aquele tema
        If rs.EOF = False Then
            If (rs!FILT_1 & "A") <> "A" Then
                'Me.chkFiltro.Enabled = True
                Me.chkFiltro.Value = 1
                vetor = Split(rs!FILT_1, ";")
                Me.cboColunas.Text = vetor(0)
                Me.cboOperador.Text = vetor(1)
                Me.cboFiltro.Text = vetor(2)
            End If
            If (rs!FILT_2 & "A") <> "A" Then
                'Me.chkFiltro.Enabled = True
                Me.chkFiltro.Value = 1
                vetor = Split(rs!FILT_2, ";")
                Me.cboColunas2.Text = vetor(0)
                Me.cboOperador2.Text = vetor(1)
                Me.cboFiltro2.Text = vetor(2)
            End If
            If (rs!FILT_3 & "A") <> "A" Then
                'Me.chkFiltraData.Enabled = True
                Me.chkFiltraData.Value = 1
                'Me.txtDataInicio.Enabled = False
                'Me.txtDataFim.Enabled = False
                vetor = Split(rs!FILT_3, ";")
                Me.txtDataInicio.Text = vetor(0)
                Me.txtDataFim.Text = vetor(1)
            End If
        End If
        rs.Close
    End If
    Me.Show vbModal
    Init = Confirm
    Exit Function

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       Open App.Path & "GeoSanLog.txt" For Append As #1
       Print #1, Now & " - ViewManager.ctl - Public Function Init - " & Err.Number & " - " & Err.Description
       Close #1
       MsgBox "Um posssível erro foi identificado (Function NxViewManagerInit):" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo " + App.Path + " GeoSanLog.txt com informações desta ocorrência.", vbInformation
    End If
End Function

'Private Sub LoadThemeFilter()
'   If tvm.GetThemeWhere(tvm.getactiveview, ThemeName) = "" Then
'      chkFiltro.Value = 0
'   Else
'      chkFiltro.Value = 1
'      'GetThemeWhere tvm.GetThemeWhere(tvm.getActiveView, ThemeName), tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName), _
'                     Lv
'      LoadFilter tvm.getLayerNameFromTheme(tvm.getactiveview, ThemeName), tvm.GetThemeWhere(tvm.getactiveview, ThemeName), getPmsdp(tvm.getLayerNameFromTheme(tvm.getactiveview, ThemeName), 0, 0, TypeConn, conn)
'   End If
'End Sub
Private Sub Representarion_Visibled(mtheme As String, ActiveView As String)
   If tvm.existPolygon(ActiveView, mtheme) Then
      LoadThemeProperties 1, ActiveView, mtheme
      fraPol(0).Visible = True: fraPol(1).Visible = True: mtab.Tab = 0
   Else
      fraPol(0).Visible = False: fraPol(1).Visible = False
   End If
   If tvm.existLine(ActiveView, mtheme) Then
      fralines(0).Visible = True: fralines(1).Visible = True: mtab.Tab = 1
      LoadThemeProperties 2, ActiveView, mtheme
   Else
      fralines(0).Visible = False: fralines(1).Visible = False
   End If
   If tvm.existPoint(ActiveView, mtheme) Then
      frapoints(0).Visible = True: frapoints(1).Visible = True: mtab.Tab = 2
      LoadThemeProperties 4, ActiveView, mtheme
   Else
      frapoints(0).Visible = False: frapoints(1).Visible = False
   End If
   If tvm.existText(ActiveView, mtheme) Then
      fratexts(0).Visible = True: fratexts(1).Visible = True: mtab.Tab = 3
      LoadThemeProperties 128, ActiveView, mtheme
   Else
      fratexts(0).Visible = False: fratexts(1).Visible = False
   End If
End Sub

Private Sub LoadThemeProperties(Representation As Long, ActiveView As String, mtheme As String)
   On Error GoTo LoadThemeProperties_Err
   Dim a As Integer, Repvisibled As Integer
   With tvm
   
   'GetRepByTheme mtheme 'set chkvisibled
   Select Case Representation
      Case 1 'Poligono
         
         polyguns1.FillStyle = IIf(.getPolygonStyle(ActiveView, mtheme) = 0, 1, IIf(.getPolygonStyle(ActiveView, mtheme) = 1, 0, .getPolygonStyle(ActiveView, mtheme)))
         polyguns1.FillColor = .getPolygonColor(ActiveView, mtheme)
         polyguns1.BorderColor = .getPolygonContourColor(ActiveView, mtheme)
         'polyguns1.BackStyle = .getPolygonStyle(ActiveView, mTheme)
         polyguns1.BorderWidth = .getPolygonContourWidth(ActiveView, mtheme)
         txtPolBorderWidth.Text = polyguns1.BorderWidth
         cboPolStyleBorder.ListIndex = GetCboListIndex(-1, cboPolStyleFill, 2)
         'chkVisibledPolyguns.Value = IIf(.visiblePolygon(.getactiveview, ThemeName), 1, 0)
         
      Case 2 'linha

         Line1.BorderStyle = IIf(.getLineStyle(ActiveView, mtheme) > 15, 0, .getLineStyle(ActiveView, mtheme) + 1)
         Line1.BorderColor = .getLineColor(ActiveView, mtheme)
         Line1.BorderWidth = .getLineWidth(ActiveView, mtheme)
         txtLineWidth.Text = .getLineWidth(ActiveView, mtheme)
         cboLineStyle.ListIndex = GetCboListIndex(.getLineStyle(ActiveView, mtheme), cboLineStyle, 2)
         'chkVisibledLine.Value = IIf(.visibleLine(.getactiveview, ThemeName), 1, 0)
         
      Case 128 ' Texto
         txtFonteSize.Text = .getFontSize(ActiveView, mtheme)
         txtFonteStyle.Text = .getFontName(ActiveView, mtheme)
         txtExample.SelColor = .getTextColor(ActiveView, mtheme)
         txtExample.SelBold = .getTextBold(ActiveView, mtheme)
         txtExample.SelItalic = .getTextItalic(ActiveView, mtheme)
         
         'chkVisibledText.Value = IIf(.visibleText(.getactiveview, ThemeName), 1, 0)
         chkFonteBold.Value = IIf(.getTextBold(.getActiveView, ThemeName), 1, 0)
         chkFonteItalic.Value = IIf(.getTextItalic(.getActiveView, ThemeName), 1, 0)
      Case 4   ' Ponto
         shpPoint.FillColor = .getPointColor(ActiveView, mtheme)
         cboPointStyle.ListIndex = GetCboListIndex(.getPointStyle(ActiveView, mtheme), cboPointStyle, 4)
         txtPointWidth = .getPointSize(ActiveView, mtheme)
         'chkVisibledPoint.Value = IIf(.visiblePoint(.getactiveview, ThemeName), 1, 0)
         If .styleImageExist(ActiveView, mtheme) Then
            txtAngle.Text = .getStyleImageAngle(.getActiveView, mtheme)
            txtPointWidth.Text = .getStyleImageSize(.getActiveView, mtheme)
            cmdIcone.Caption = "Remover"
            NameIcon = "EXIT"
         End If
         
      Case Else 'Imagems/Outros (falta implementar)
   End Select
   End With
   Exit Sub
LoadThemeProperties_Err:
   Resume Next
End Sub




Private Sub cboLineStyle_Click()
   If cboLineStyle.ItemData(cboLineStyle.ListIndex) <= 5 Then
      Line1.BorderStyle = cboLineStyle.ItemData(cboLineStyle.ListIndex)
   Else
      Line1.BorderStyle = 0
   End If
End Sub

Private Sub cboPolStyleFill_Click()
   polyguns1.FillStyle = IIf(cboPolStyleFill.ItemData(cboPolStyleFill.ListIndex) = 0, 1, IIf(cboPolStyleFill.ItemData(cboPolStyleFill.ListIndex) = 1, 0, cboPolStyleFill.ItemData(cboPolStyleFill.ListIndex)))
End Sub



Private Sub chkFonteBold_Click()
   txtExample.SelBold = chkFonteBold.Value
End Sub

Private Sub chkFonteItalic_Click()
   txtExample.SelItalic = chkFonteItalic.Value
End Sub

Private Sub LoadLvFilter(FieldName As String, Operation As String, mValue As String, Tag As Integer)
   Dim itmx As ListItem
   If FieldName <> "" And Operation <> "" And mValue <> "" Then
      Set itmx = lv.ListItems.Add(, , FieldName)
         itmx.SubItems(1) = Operation
         itmx.SubItems(2) = mValue
         itmx.Tag = Tag
   End If
   Set itmx = Nothing
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdCollorLine_Click()
   Cdl.ShowColor
   Line1.BorderColor = Cdl.Color
End Sub

Private Sub cmdCollorPoint_Click()
   Cdl.ShowColor
   shpPoint.FillColor = Cdl.Color
End Sub

Private Sub cmdCollorPol_Click()
   Cdl.ShowColor
   polyguns1.FillColor = Cdl.Color
End Sub

Private Sub cmdCollorPolBor_Click()
   Cdl.ShowColor
   polyguns1.BorderColor = Cdl.Color
End Sub

Private Sub cmdColorTransp_Click()
   Cdl.ShowColor
   ColorTransp = Cdl.Color
End Sub

Private Sub cmdFonte_Click()
On Error GoTo CMDFONTE_ERR
   With Cdl
      .Color = txtExample.SelColor
      .FontSize = txtExample.SelFontSize
      .FontBold = txtExample.SelBold
      .FontItalic = txtExample.SelItalic
      .FontName = txtExample.Font
      .Flags = cdlCFScreenFonts
      .ShowFont
      txtExample.SelColor = .Color
      txtExample.SelFontSize = .FontSize
      txtExample.SelBold = .FontBold
      txtExample.SelItalic = .FontItalic
      txtExample.Font = .FontName
      txtFonteStyle.Text = .FontName
   End With
CMDFONTE_ERR:
End Sub


Private Sub cmdIcone_Click()
   If cmdIcone.Caption = "Inserir" Then
      Cdl.filename = ""
      Cdl.Filter = "Pictures (*.bmp;*.png;*.jpg)|*.bmp;*.ico;*.jpg"
      Cdl.ShowOpen
      If Not Cdl.filename = "" Then
         NameIcon = Cdl.filename
         InserirIcon = True
      End If
    Else
      cmdIcone.Enabled = False
      NameIcon = ""
    End If
End Sub



Private Sub cmdTextCollor_Click()
   Cdl.ShowColor
   txtExample.SelColor = Cdl.Color
End Sub



Private Sub Form_Unload(Cancel As Integer)
   'LoozeXP1.EndWinXPCSubClassing
End Sub

Private Sub txtAngle_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtFonteSize_Change()
   If txtFonteSize.Text <> "" Then
      txtExample.SelFontSize = txtFonteSize.Text
   End If
End Sub

Private Sub txtFonteSize_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyDelete, vbKeyTab, vbKeyBack, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
      
      Case Else
         KeyAscii = 0
   End Select
End Sub

Private Sub txtLineWidth_Change()
   On Error GoTo txtLineWidth_Change_err
   'If KeyAscii <> vbEmpty Then
      If txtLineWidth.Text <> "" Then
          If IsNumeric(txtLineWidth.Text) Then
              Line1.BorderWidth = txtLineWidth.Text
          Else
              MsgBox "Só é permitido números", vbExclamation
              txtLineWidth.Text = ""
          End If
      End If
   'End If
   Exit Sub
txtLineWidth_Change_err:
   txtLineWidth.Text = Line1.BorderWidth
End Sub

'Private Sub txtLineWidth_KeyPress(KeyAscii As Integer)
'
'   'If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
'
'End Sub

Private Sub txtPointWidth_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtPolBorderWidth_Change()
   polyguns1.BorderWidth = txtPolBorderWidth.Text
End Sub

Private Sub txtPolBorderWidth_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub cmdOK_Click()
On Error GoTo Trata_Erro
   Dim momento As String
   If txtFonteSize.Text = "" Then
      MsgBox "informe o tamanho do fonte do texto", vbExclamation
      Exit Sub
   End If
   
  ' With tvm
      If tvm.existPolygon(tvm.getActiveView, ThemeName) Then
         tvm.setPolygonStyle tvm.getActiveView, ThemeName, IIf(polyguns1.FillStyle = 0, 1, IIf(polyguns1.FillStyle = 1, 0, polyguns1.FillStyle))
         tvm.setPolygonColor tvm.getActiveView, ThemeName, polyguns1.FillColor
         tvm.setPolygonContourColor tvm.getActiveView, ThemeName, polyguns1.BorderColor
         tvm.setPolygonContourWidth tvm.getActiveView, ThemeName, polyguns1.BorderWidth
         tvm.setVisiblePolygonStatus tvm.getActiveView, ThemeName, IIf(chkVisibledPolyguns.Value = 1, True, False)
      End If
      If tvm.existLine(tvm.getActiveView, ThemeName) Then
         tvm.setLineStyle tvm.getActiveView, ThemeName, IIf(Line1.BorderStyle = 0, 107, Line1.BorderStyle - 1)
         tvm.setLineColor tvm.getActiveView, ThemeName, Line1.BorderColor
         tvm.setLineWidth tvm.getActiveView, ThemeName, Line1.BorderWidth
         tvm.setVisibleLineStatus tvm.getActiveView, ThemeName, IIf(chkVisibledLine.Value = 1, True, False)
      End If
      If tvm.existPoint(tvm.getActiveView, ThemeName) Then
         tvm.setPointColor tvm.getActiveView, ThemeName, shpPoint.FillColor
         tvm.setPointStyle tvm.getActiveView, ThemeName, cboPointStyle.ItemData(cboPointStyle.ListIndex)
         tvm.setPointSize tvm.getActiveView, ThemeName, txtPointWidth.Text
         tvm.setVisiblePointStatus tvm.getActiveView, ThemeName, IIf(chkVisibledPoint.Value = 1, True, False)
         If NameIcon = "" Then
            tvm.removeStyleImageFromTheme tvm.getActiveView, ThemeName
         Else
            If tvm.styleImageExist(tvm.getActiveView, ThemeName) = False Then
               tvm.importStyleImageToTheme tvm.getActiveView, ThemeName, NameIcon, 5, ColorTransp
            End If
            tvm.setStyleImageAngle tvm.getActiveView, ThemeName, txtAngle.Text
            tvm.setStyleImageSize tvm.getActiveView, ThemeName, txtPointWidth
         End If
      End If
      If tvm.existText(tvm.getActiveView, ThemeName) Then
         tvm.setFontSize tvm.getActiveView, ThemeName, txtFonteSize.Text
         tvm.setFontName tvm.getActiveView, ThemeName, txtFonteStyle.Text
         tvm.setTextColor tvm.getActiveView, ThemeName, txtExample.SelColor
         tvm.setTextBold tvm.getActiveView, ThemeName, txtExample.SelBold
         tvm.setTextitalic tvm.getActiveView, ThemeName, txtExample.SelItalic
         tvm.setVisibleTextStatus tvm.getActiveView, ThemeName, IIf(chkVisibledText.Value = 1, True, False)
      End If
      
      
      '.setMinScale .getactiveview, ThemeName, txtMin.Text
      
      tvm.setMinScale tvm.getActiveView, ThemeName, IIf(IsNumeric(txtMin.Text), txtMin.Text, "0")
      
      '.setMaxScale .getactiveview, ThemeName, txtMax.Text
      
      tvm.setMaxScale tvm.getActiveView, ThemeName, IIf(IsNumeric(txtMax.Text), txtMax.Text, "0")
      

      

   'End With
   Confirm = True
   Unload Me
   Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #1
        Open App.Path & "\GeoSanLog.txt" For Append As #1
        'Print #1, Now & " - NxViewManager - frmTheme - Private Sub cmdOK_Click() - " & Err.Number & " - " & Err.Description
        Close #1
        'MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & "Não foi possível estabelecer a conexão" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrencia.", vbInformation
    End If
End Sub



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

Public Function RetornaID2(nome As String, nome2 As String)


Dim atributo As String
Dim tabela As String

If (nome = "TIPO" Or nome = "MATERIAL" Or nome = "FORNECEDOR" Or nome = "FABRICANTE" Or nome = "LOCALIZAÇÃO" Or nome = "ESTADO" Or nome = "[LADO_DA_RUA]") Then
If (nome = "TIPO") Then
tabela = "WATERLINESTYPES"
End If

If (nome = "MATERIAL") Then
tabela = "X_MATERIAL"
'nome2 = "MATERIALID"
End If

If (nome = "FORNECEDOR") Then
tabela = "X_SUPPLIERS"
'nome2 = "SUPPLIERID"
End If

If (nome = "FABRICANTE") Then
tabela = "X_MANUFACTURERS"
'nome2 = "MANUFACTURERid"
End If

If (nome = "LOCALIZAÇÃO") Then
tabela = "X_LOCATION"
'nome2 = "LOCATIONID"
End If

If (nome = "ESTADO") Then
tabela = "X_STATE"
'nome2 = "STATEID"
End If

If (nome = "[LADO_DA_RUA]") Then
tabela = "X_SIDESTREET"
'nome2 = "SIDESTREET_ID"
End If


 If TypeConn <> 4 Then
 Set rs = conn.Execute("Select * from " + tabela + "")
 Else
 Set rs = conn.Execute("Select * from " + """" + tabela + """" + "")
 
 End If


If (nome = "MATERIAL") Then


     While Not rs.EOF
               If rs.Fields("MATERIALNAME").Value = nome2 Then
                  atributo = rs.Fields("MATERIALID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing


ElseIf (nome = "TIPO") Then
 While Not rs.EOF
               If rs.Fields("DESCRIPTION_").Value = nome2 Then
                  atributo = rs.Fields("ID_TYPE").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            ElseIf (nome = "FORNECEDOR") Then
 While Not rs.EOF
               If rs.Fields("COMPANYNAME").Value = nome2 Then
                  atributo = rs.Fields("SUPPLIERID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            
            ElseIf (nome = "FABRICANTE") Then
 While Not rs.EOF
               If rs.Fields("COMPANYNAME").Value = nome2 Then
                  atributo = rs.Fields("MANUFACTUREID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            
            ElseIf (nome = "LOCALIZAÇÃO") Then
 While Not rs.EOF
               If rs.Fields("LOCATIONNAME").Value = nome2 Then
                  atributo = rs.Fields("LOCATIONID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            
            ElseIf (nome = "ESTADO") Then
 While Not rs.EOF
               If rs.Fields("STATENAME").Value = nome2 Then
                  atributo = rs.Fields("STATEID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            ElseIf (nome = "[LADO_DA_RUA]") Then
 While Not rs.EOF
               If rs.Fields("DESCRIPTION").Value = nome2 Then
                  atributo = rs.Fields("SIDESTREET_ID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
             
            

End If


End If



If (atributo = "") Then
atributo = Me.cboFiltro2.Text
End If



RetornaID2 = atributo

End Function








Public Function RetornaID(nome As String, nome2 As String)


Dim atributo As String
Dim tabela As String
   ThemeName2 = tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName)
If (nome = "TIPO" Or nome = "MATERIAL" Or nome = "FORNECEDOR" Or nome = "FABRICANTE" Or nome = "LOCALIZAÇÃO" Or nome = "ESTADO" Or nome = "[LADO_DA_RUA]") Then
If (nome = "TIPO" And ThemeName2 = "WATERLINES") Then
tabela = "WATERLINESTYPES"
End If

If (nome = "TIPO" And ThemeName2 = "SEWERLINES") Then
tabela = "SEWERLINESTYPES"
End If

If (nome = "TIPO" And ThemeName2 = "WATERCOMPONENTS") Then
tabela = "WATERCOMPONENTSTYPES"
End If

If (nome = "TIPO" And ThemeName2 = "SEWERCOMPONENTS") Then
tabela = "SEWERCOMPONENTSTYPES"
End If



If (nome = "MATERIAL") Then
tabela = "X_MATERIAL"
'nome2 = "MATERIALID"
End If

If (nome = "FORNECEDOR") Then
tabela = "X_SUPPLIERS"
'nome2 = "SUPPLIERID"
End If

If (nome = "FABRICANTE") Then
tabela = "X_MANUFACTURERS"
'nome2 = "MANUFACTURERid"
End If

If (nome = "LOCALIZAÇÃO") Then
tabela = "X_LOCATION"
'nome2 = "LOCATIONID"
End If

If (nome = "ESTADO") Then
tabela = "X_STATE"
'nome2 = "STATEID"
End If

If (nome = "[LADO_DA_RUA]") Then
tabela = "X_SIDESTREET"
'nome2 = "SIDESTREET_ID"
End If









 If TypeConn <> 4 Then
 Set rs = conn.Execute("Select * from " + tabela + "")
 Else
 Set rs = conn.Execute("Select * from " + """" + tabela + """" + "")
 
 End If


If (nome = "MATERIAL") Then


     While Not rs.EOF
               If rs.Fields("MATERIALNAME").Value = nome2 Then
                  atributo = rs.Fields("MATERIALID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing


ElseIf (nome = "TIPO") Then
 While Not rs.EOF
               If rs.Fields("DESCRIPTION_").Value = nome2 Then
                  atributo = rs.Fields("ID_TYPE").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            ElseIf (nome = "FORNECEDOR") Then
 While Not rs.EOF
               If rs.Fields("COMPANYNAME").Value = nome2 Then
                  atributo = rs.Fields("SUPPLIERID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            
            ElseIf (nome = "FABRICANTE") Then
 While Not rs.EOF
               If rs.Fields("COMPANYNAME").Value = nome2 Then
                  atributo = rs.Fields("MANUFACTUREID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            
            ElseIf (nome = "LOCALIZAÇÃO") Then
 While Not rs.EOF
               If rs.Fields("LOCATIONNAME").Value = nome2 Then
                  atributo = rs.Fields("LOCATIONID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            
            ElseIf (nome = "ESTADO") Then
 While Not rs.EOF
               If rs.Fields("STATENAME").Value = nome2 Then
                  atributo = rs.Fields("STATEID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            
            ElseIf (nome = "[LADO_DA_RUA]") Then
 While Not rs.EOF
               If rs.Fields("DESCRIPTION").Value = nome2 Then
                  atributo = rs.Fields("SIDESTREET_ID").Value
               End If
              
               rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
             
            

End If


End If



If (atributo = "") Then
atributo = Me.cboFiltro.Text
End If



RetornaID = atributo

End Function
'Função responsável por retornar o nome da coluna a ser pesquisado na tabela, para o filtro da cláusula Where da tabela Theme_Id
'
'
Public Function Retorna(nome As String)
    Dim atributo As String
    
    If (nome = "TIPO") Then
        atributo = "ID_TYPE"
    End If
    If (nome = "[COTA DO TERRENO]") Then
        atributo = "INITIALGROUNDHEIGHT"
    End If
    If (nome = "[DEMANDA]") Then
        atributo = "DEMAND"
    End If
    If (nome = "[NÓ DE CÁLCULO]") Then
        atributo = "CALCULE_NODE"
    End If
    If (nome = "VALIDADE") Then
        atributo = "INFORMATIONVALIDITY"
    End If
    If (nome = "Observação") Then
        atributo = "NOTES"
    End If
    If (nome = "[NÃO_CONFORMIDADE]") Then
        atributo = "TROUBLE"
    End If
    If (nome = "[PADRÃO_CONSUMO]") Then
        atributo = "PATTERN"
    End If
    If (nome = "[SETOR]") Then
        atributo = "SECTOR"
    End If
    If (nome = "[TERRENO - COTA INICIAL]") Then
        If TypeConn <> 4 Then
            atributo = "INITIALGROUNDHEIGHT"
        Else
            atributo = "INITIALGROUNDHEIGHT"
        End If
    End If
    If (nome = "[COTA DO TERRENO]") Then
        If TypeConn <> 4 Then
            atributo = "INITIALGROUNDHEIGHT"
        Else
            atributo = "INITIALGROUNDHEIGHT"
        End If
    End If
    If (nome = "[COTA DO FUNDO]") Then
        If TypeConn <> 4 Then
            atributo = "FINALGROUNDHEIGHT"
        Else
            atributo = "FINALGROUNDHEIGHT"
        End If
    End If
    If (nome = "[TERRENO - COTA FINAL]") Then
        If TypeConn <> 4 Then
            atributo = "GROUNDHEIGHTFINAL"
        Else
            atributo = "FINALGROUNDHEIGHT"
        End If
    End If
    If (nome = "[TERRENO - COTA FINAL]") Then
        atributo = "FINALGROUNDHEIGHT"
    End If
    If (nome = "[PEÇA - COTA INICIAL]") Then
        atributo = "INITIALTUBEDEEPNESS"
    End If
    If (nome = "[ANO DE FABRICAÇÃO]") Then
        atributo = "YEAROFCONSTRUCTION"
    End If
    If (nome = "[PEÇA - COTA FINAL]") Then
        atributo = "FINALTUBEDEEPNESS"
    End If
    If (nome = "[DIAMETRO INTER.(MM)]") Then
        atributo = "INTERNALDIAMETER"
    End If
    If (nome = "[DIAMETRO EXT.(MM)]") Then
        atributo = "EXTERNALDIAMETER"
    End If
    If (nome = "[INICIAL COMPONENTE]") Then
        atributo = "INITIALCOMPONENT"
    End If
    If (nome = "[FINAL COMPONENTE]") Then
        atributo = "FINALCOMPONENT"
    End If
    If (nome = "DENSIDADE") Then
        atributo = "THICKNESS"
    End If
    If (nome = "MATERIAL") Then
        atributo = "MATERIAL"
    End If
    If (nome = "[COMPRIMENTO(M)]") Then
        atributo = "LENGTH"
    End If
    If (nome = "[COMPR. CALCULADO]") Then
        atributo = "LENGTHCALCULATED"
    End If
    If (nome = "FORNECEDOR") Then
        atributo = "SUPPLIER"
    End If
    If (nome = "FABRICANTE") Then
        atributo = "MANUFACTURER"
    End If
    If (nome = "LOCALIZAÇÃO") Then
        atributo = "LOCATION"
    End If
    If (nome = "ESTADO") Then
        atributo = "STATE"
    End If
    If (nome = "RUGOSIDADE") Then
        atributo = "ROUGHNESS"
    End If
    If (nome = "SETOR") Then
        atributo = "SECTOR"
    End If
    If (nome = "[LADO_DA_RUA]") Then
        atributo = "SIDESTREET"
    End If
    If (nome = "[DISTÂNCIA_DA_DIVISA]") Then
        atributo = "DIVIDEDDISTANCE"
    End If
    If (nome = "USUÁRIO") Then
        atributo = "USUARIO_LOG"
    End If
    If (nome = x) Then
        atributo = "MANUFACTURER"
    End If
    If (nome = "LINHA") Then
        atributo = "LINE_ID"
    End If
    If (nome = "DATA_DE_INSTALAÇÃO") Then
        atributo = "DATEINSTALLATION"
    End If
    If (nome = "PROFUNDIDADE_RAMAL") Then
        atributo = "PROFUNDIDADE_RAMAL"
    End If
    If (nome = "HIDROMETRADO") Then
        atributo = "HIDROMETRADO"
    End If
    If (nome = "ECONOMIAS") Then
        atributo = "ECONOMIAS"
    End If
    If (nome = "CONSUMO_LPS") Then
        atributo = "CONSUMO_LPS"
    End If
    If (nome = "DISTANCIA_TESTADA") Then
        atributo = "DISTANCIA_TESTADA"
    End If
    If (nome = "DISTANCIA_LADO") Then
        atributo = "DISTANCIA_LADO"
    End If
    If (nome = "COMPRIMENTO_RAMAL") Then
        atributo = "COMPRIMENTO_RAMAL"
    End If
    If (nome = "USUARIO_LOG") Then
        atributo = "USUARIO_LOG"
    End If
    Retorna = atributo
End Function
'Esta rotina modifica o filtro de pesquisa das entidades geográficas, ou seja a coluna [generate_attribute_where] da tabela [te_theme]
'o filtro (cláusula where) colocado nesta coluna, fará com que sejam apresentados apenas as geometrias que satisfação esta cláusula
'
'
'
Private Sub cmdModificar_Click()
    On Error GoTo Trata_Erro
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
    Dim mPROVEDOR As String
    Dim mSERVIDOR As String
    Dim mPORTA As String
    Dim mBANCO As String
    Dim mUSUARIO As String
    Dim Senha As String
    Dim decriptada As String
    Dim conexao As New ADODB.Connection
    Dim Filtro As String
    
    If cmdModificar.Caption = "Modificar" Then                                          'Caso tenha selecionado o botão indicando que deseja alterar o filtro de pesquisa
        Me.chkFiltro.Value = 0
        Me.chkFiltro.Enabled = True
        Me.cboColunas.Text = ""
        Me.cboColunas2.Text = ""
        Me.cboFiltro.Text = ""
        Me.cboFiltro2.Text = ""
        Me.cboOperador.Text = ""
        Me.cboOperador2.Text = ""
        Me.chkFiltraData.Value = 0
        If ThemeName2 = "SEWERLINES" Or ThemeName2 = "SEWERCOMPONENTS" Then             'se for esgoto ou nó de esgoto
            Me.txtDataInicio.Enabled = True
            Me.txtDataFim.Enabled = True
            Me.Label5.Enabled = True
            Me.Label6.Enabled = True
            Me.Label7.Enabled = True
            chkFiltraData.Visible = True
            Label7.Visible = True
            Label6.Visible = True
            Label5.Visible = True
            txtDataInicio.Visible = True
            txtDataFim.Visible = True
            Me.chkFiltraData.Enabled = True
        End If
        Me.Label5.Enabled = True
        Me.Label6.Enabled = True
        Me.Label7.Enabled = True
        Me.txtDataInicio.Text = ""
        Me.txtDataFim.Text = ""
        Me.cmdModificar.Caption = "Salvar"
    ElseIf cmdModificar.Caption = "Salvar" Then
        'PARTE QUE SALVA FILTROS
        'LIMPA O FILTRO E LOG DE FILTROS ANTERIORES
        Dim rs As New ADODB.Recordset
        a = "NXGS_FILT_TEMA"            'nesta tabela estão armazenados todos os filtros selecionados por todos os usuários, os quais estão associados a um tema da tabela [theme_id]
        b = "THEME_ID"
        'localiza o filtro que foi anteriormente entrado pelo usuário na caixa de diálogo, o qual está associado a um tema (theme_id)
        rs.Open "SELECT * FROM NXGS_FILT_TEMA WHERE theme_id = " & intTema, conn, adOpenKeyset, adLockOptimistic
        If rs.EOF = True Then ' se EOF = true significa que não existe ainda o thema no filtro.. é criado então
            rs.AddNew
            rs!theme_id = intTema
            rs.Update
            rs.Close
            Dim aa As String
            Dim bb As String
            rs.Open "SELECT * FROM NXGS_FILT_TEMA WHERE theme_id = " & intTema, conn, adOpenKeyset, adLockOptimistic
        End If
        If rs.EOF = False Then
            'Atualiza na tabela [NXGS_FILT_TEMA] os filtros 1 e 2, o filtro 3 não está atualizando, é sempre zerado
            'Isto registra apenas para cada usuário do GeoSan o filtro que está utilizando, não implica em nenhuma mudança de visualização
            If Me.chkFiltro.Value = 1 And Me.cboColunas.Text <> "" And Me.cboOperador.Text <> "" And Me.cboFiltro.Text <> "" Then
                rs!FILT_1 = Me.cboColunas.Text & ";" & Me.cboOperador.Text & ";" & Me.cboFiltro.Text
            Else
                rs!FILT_1 = ""
            End If
            If Me.chkFiltro.Value = 1 And Me.cboColunas2.Text <> "" And Me.cboOperador2.Text <> "" And Me.cboFiltro2.Text <> "" Then
                rs!FILT_2 = Me.cboColunas2.Text & ";" & Me.cboOperador2.Text & ";" & Me.cboFiltro2.Text
            Else
                rs!FILT_2 = ""
            End If
            rs!FILT_3 = ""
            rs.Update
        End If
        rs.Close
        'OS FILTROS 1 E 2 SÃO SALVOS PELO COMANDO ABAIXO
        ' If LayerAtivo = "RAMAIS_AGUA" Then
        If ThemeName2 = "WATERLINES" Or ThemeName2 = "SEWERLINES" Or ThemeName2 = "WATERCOMPONENTS" Or ThemeName2 = "SEWERCOMPONENTS" Or ThemeName2 = "RAMAIS_AGUA" Or ThemeName2 = "RAMAIS_AGUA_LIGACAO" Then
            'caso estivermos tratando de redes de água, esgoto, ou mesmo componentes de água ou esgoto
            FILT = Me.cboColunas.Text
            tabela = tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName)                                'obtem o nome do tema ativo
            'modifica o nome da tabela, pois quando o usuário seleciona o layer de ramais, são duas tabelas que estão associadas ao mesmo, a de ramais e a de ligações
            '
            '
            'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX AQUI ESTÁ O PROBLEMA ELE ACHA QUE É RAMAL E NAO NÓ
            '
            ' retirado do if abaixo FILT = "TIPO" Or
            If (FILT = "HIDROMETRADO" Or FILT = "ECONOMIAS" Or FILT = "CONSUMO_LPS") And ThemeName2 <> "WATERCOMPONENTS" Then     'RAMAIS_AGUA_LIGACAO
                tabela = "RAMAIS_AGUA_LIGACAO"
            ElseIf FILT = "DISTANCIA_TESTADA" Or FILT = "DISTANCIA_LADO" Or FILT = "COMPRIMENTO_RAMAL" Or FILT = "PROFUNDIDADE_RAMAL" Or FILT = "USUARIO_LOG" Then
                tabela = "RAMAIS_AGUA"
            End If
            'verifica o sinal da primeira comparação da variável com o valor
            If Me.cboOperador.Text = "Igual" Then
                sinal = "="
            ElseIf Me.cboOperador.Text = "Maior" Then
                sinal = ">"
            ElseIf Me.cboOperador.Text = "Menor" Then
                sinal = "<"
            ElseIf Me.cboOperador.Text = "Diferente" Then
                sinal = "<>"
            End If
            Dim sinal2 As String
            'verifica o sinal da segunda comparação da variável com o valor
            If Me.cboOperador2.Text = "Igual" Then
                sinal2 = "="
            ElseIf Me.cboOperador2.Text = "Maior" Then
                sinal2 = ">"
            ElseIf Me.cboOperador2.Text = "Menor" Then
                sinal2 = "<"
            ElseIf Me.cboOperador2.Text = "Diferente" Then
                sinal2 = "<>"
            End If
            Dim aaa As String
            aaa = "OBJECT_ID_"
            a = "object_id"
            b = tabela
            Dim a12, a13 As String
            a12 = Retorna(Me.cboColunas.Text)
            a13 = RetornaID(Me.cboColunas.Text, Me.cboFiltro.Text)
            Filtro = "object_id in (select object_id_ from " & tabela & " WHERE " + a12 + sinal & "'" & a13 & "'"
            If Me.cboOperador2.Text <> "" Then
                'adiciona o segunto filtro
                a12 = Retorna(Me.cboColunas2.Text)
                a13 = RetornaID2(Me.cboColunas2.Text, Me.cboFiltro2.Text)
                Filtro = Filtro + "AND " + a12 + sinal2 & "'" & a13 & "')"
                If Me.cboColunas.Text = "" Or Me.cboFiltro = "" Or Me.cboOperador.Text = "" Then
                    Filtro = ""
                End If
            Else
                'não precisa adicionar o terceiro filtro
                Filtro = Filtro + ")"
            End If
            If Me.cboColunas.Text = "" Or Me.cboFiltro = "" Or Me.cboOperador.Text = "" Then
                Filtro = ""
            End If
            a = "te_theme"
            b = "theme_id"
            'abre a conexão com a tabela de themas, onde existe a cláusula Where a ser preenchida
            'se SQLServer ou Oracle
            rs.Open "SELECT * FROM TE_THEME WHERE THEME_ID = " & intTema, conn, adOpenDynamic, adLockOptimistic
            'verifica se retornou pelo menos uma linha, deve retornar apenas uma, e então atualiza o filtro na cláusula where da tabela de temas, TE_THEME
            If rs.EOF = False Then ' se EOF = true significa que não existe ainda o thema no filtro.. é criado então
                rs!generate_attribute_where = Filtro
                rs.Update
            End If
            rs.Close
        Else
            'caso não seja rede de água e esgoto, nem componentes de água e esgoto. Por exemplo, caso sejam ramais
            ' GetThemeWhere
            'With tvm
            ' If Not (chkFiltro.Value = 1 And cboColunas.ListIndex = -1 And cboColunas2.ListIndex = -1) Then
            ' .setThemeWhere .getactiveview, ThemeName, IIf(chkFiltro.Value = 1, GetThemeWhere, "")
            ' End If
            '  End With
            'SALVAMENTO DE FILTRO DE DATA
            Dim nCmdo As String
            Dim DTA_INI As String
            Dim DTA_FIM As String
            If TypeConn = 1 Then  'CASO CONEXÃO SEJA SQL - formata a data
                'INVERTE A DATA E COLOCA HH,MM,SS
                DTA_INI = "'20" & Mid(Me.txtDataInicio.Text, 7, 2) & "-" & Mid(Me.txtDataInicio.Text, 4, 2) & "-" & Mid(Me.txtDataInicio.Text, 1, 2) & " 00:00:00.000'"
                DTA_FIM = "'20" & Mid(Me.txtDataFim.Text, 7, 2) & "-" & Mid(Me.txtDataFim.Text, 4, 2) & "-" & Mid(Me.txtDataFim.Text, 1, 2) & " 23:59:59.998'"
            End If
            If Me.chkFiltraData.Value = 1 Then 'SE O CHKFILTRODATA = SELECIONADO
                'Se o filtro por data foi selecionado
                If IsDate(Me.txtDataInicio.Text) = True And IsDate(Me.txtDataFim.Text) = True Then
                    If Len(Me.txtDataInicio.Text) = 8 And Len(Me.txtDataFim.Text) = 8 Then
                        'TRATAMENTO ESPECIAL CASO O FILTRO SEJA DATA DO CADASTRO, GRAVA A STRING DIRETO NO BANCO
                        Dim sGENERATE_ATTRIBUTE_WHERE As String
                        Dim sql1 As String
                        If TypeConn = 1 Then  'SQL
                            sGENERATE_ATTRIBUTE_WHERE = "object_id in(select a.object_id_ from WATERLINES A left join WATERLINESdata B on A.Object_id_=B.Object_id_ where A.DATALOG Between " & DTA_INI & " AND " & DTA_FIM & ")"
                        ElseIf TypeConn = 2 Then  'ORACLE
                            sGENERATE_ATTRIBUTE_WHERE = "object_id in(select a.object_id_ from WATERLINES A left join WATERLINESdata B on A.Object_id_=B.Object_id_ where A.DATALOG BETWEEN TO_DATE('" & Me.txtDataInicio.Text & "','DD/MM/YY" & "') and TO_DATE('" & Me.txtDataFim & "','DD/MM/YY" & "'))"
                        ElseIf TypeConn = 4 Then  'POSTGRES
                            a = "OBJECT_ID_"
                            b = "WATERLINES"
                            c = "WATERLINESDATA"
                            d = "TIPO"
                            e = "DATALOG"
                            sGENERATE_ATTRIBUTE_WHERE = "" + """" + a + """" + " in(select " + """" + b + """" + "." + """" + a + """" + " from " + """" + b + """" + " left join " + """" + c + """" + " on " + """" + b + """" + "." + """" + a + """" + " = " + """" + c + """" + "." + """" + a + """" + " where " + """" + b + """" + "." + """" + e + """" + " Between " + """" + "'" + """" + DTA_INI + """" + "'" + """" + " AND " + """" + "'" + """" + DTA_FIM + """" + "'" + """" + ")"
                        End If
                        Dim rsTheme As New ADODB.Recordset
                        a = "te_theme"
                        b = "theme_id"
                        rsTheme.Open "SELECT * FROM te_theme WHERE theme_id = " & intTema, conn, adOpenKeyset, adLockOptimistic
                        If rsTheme.EOF = False Then
                            a = "NXGS_FILT_TEMA"
                            b = "THEME_ID"
                            rs.Open "SELECT * FROM NXGS_FILT_TEMA WHERE theme_id = " & intTema, conn, adOpenKeyset, adLockOptimistic
                            If rs.EOF = True Then ' se EOF = true significa que não existe ainda o thema no filtro.. é criado então
                                rs.AddNew
                                rs!theme_id = intTema
                                rs.Update
                                rs.Close
                                rs.Open "SELECT * FROM NXGS_FILT_TEMA WHERE theme_id = '" & intTema & "'", conn, adOpenKeyset, adLockOptimistic
                            End If
                            If (rsTheme!generate_attribute_where & "a") = "a" Then 'testa se esta vazio, caso sim é salvo apenas o filtro de data
                                rsTheme!generate_attribute_where = sGENERATE_ATTRIBUTE_WHERE
                                rs!FILT_3 = Me.txtDataInicio.Text & ";" & Me.txtDataFim.Text
                            Else 'ALÉM DA DATA HÁ OUTROS FILTROS
                                nCmdo = Mid(rsTheme!generate_attribute_where, 1, (Len(rsTheme!generate_attribute_where) - 1)) 'pega todo o comando menos o ultimo parênteses
                                'acrescenta no comando o comando de filtro de data
                                If TypeConn = 1 Then 'SQL
                                    nCmdo = nCmdo & " AND A.DATALOG BETWEEN " & DTA_INI & " AND " & DTA_FIM & ")"
                                ElseIf TypeConn = 2 Then 'ORACLE
                                    nCmdo = nCmdo & " AND A.DATALOG BETWEEN TO_DATE('" & Me.txtDataInicio.Text & "','DD/MM/YY" & "') and TO_DATE('" & Me.txtDataFim & "','DD/MM/YY" & "'))"
                                Else
                                    Dim ss As String
                                    Dim ff As String
                                    ss = "DATALOG"
                                    ff = "A"
                                    nCmdo = nCmdo & " AND " + """" + ff + """" + "." + """" + ss + """" + " BETWEEN '" & DTA_INI & "' AND '" & DTA_FIM & "')"
                                End If
                                rsTheme!generate_attribute_where = nCmdo
                                rs!FILT_3 = Me.txtDataInicio.Text & ";" & Me.txtDataFim.Text
                            End If
                            rsTheme.Update
                            rs.Update
                        End If
                        rsTheme.Close
                        rs.Close
                        conn.Close
                        conn.Open
                    Else
                        MsgBox "As datas não possuem formatação correta. Utilize 'DD/MM/AA'", vbExclamation, "Filtro de Data"
                    End If
                Else
                    MsgBox "As datas não possuem formatação correta. Utilize 'DD/MM/AA'", vbExclamation, "Filtro de Data"
                End If
            End If
        End If
        cmdModificar.Caption = "Modificar"
        Me.chkFiltro.Enabled = False
        Me.chkFiltraData.Enabled = False
        Me.cboColunas.Enabled = False
        Me.cboColunas2.Enabled = False
        Me.cboFiltro.Enabled = False
        Me.cboFiltro2.Enabled = False
        Me.cboOperador.Enabled = False
        Me.cboOperador2.Enabled = False
        Me.txtDataInicio.Enabled = False
        Me.txtDataFim.Enabled = False
        Me.Label5.Enabled = False
        Me.Label6.Enabled = False
        Me.Label7.Enabled = False
        MsgBox "Filtro salvo com sucesso!", vbInformation, ""
    End If
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #1
        Open App.Path & "\GeoSanLog.txt" For Append As #1
        Print #1, Now & " - NxViewManager - frmTheme - Private Sub cmdModificar_Click() - " & Err.Number & " - " & Err.Description
        Close #1
        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrencia.", vbInformation
    End If
End Sub
Private Sub chkFiltraData_Click()

    If Me.chkFiltraData.Value = 0 Then
        Me.txtDataInicio.Enabled = False
        Me.txtDataFim.Enabled = False
        Me.Label5.Enabled = False
        Me.Label6.Enabled = False
    Else
        Me.txtDataInicio.Enabled = True
        Me.txtDataFim.Enabled = True
        Me.Label5.Enabled = True
        Me.Label6.Enabled = True
    End If
    
End Sub
'Habilita a edição dos filtros
'
'
Private Sub chkFiltro_Click()
    'QUANDO SE SELECIONA O CHECK BOX DE FILTRAR CAMPOS O SEGUINTE EVENTO É PROVOCADO
    ThemeName2 = tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName)
    If ThemeName2 = "SEWERLINES" Or ThemeName2 = "SEWERCOMPONENTS" Then
        'Me.chkFiltro.Enabled = True
    End If
    If chkFiltro.Value = 1 Then
        LoadCboAttribute
        Me.cboColunas.Enabled = True
        Me.cboColunas2.Enabled = True
        Me.cboFiltro.Enabled = True
        Me.cboFiltro2.Enabled = True
        Me.cboOperador.Enabled = True
        Me.cboOperador2.Enabled = True
    Else
        Me.cboColunas.Enabled = False
        Me.cboColunas2.Enabled = False
        Me.cboFiltro.Enabled = False
        Me.cboFiltro2.Enabled = False
        Me.cboOperador.Enabled = False
        Me.cboOperador2.Enabled = False
    End If
End Sub

Private Function GetThemeWhere() As String
    On Error GoTo GetThemeWhere_err
    Dim sql As String
   
   Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
Dim conexao2 As New ADODB.Connection
     ThemeName2 = tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName)
If TypeConn = 4 Then

mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.Path & "\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.Path & "\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.Path & "\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.Path & "\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.Path & "\GEOSAN.ini")
usuario = ReadINI("CONEXAO", "USER", App.Path & "\GEOSAN.ini")
decriptada = FunDecripta(Senha)
strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + decriptada + "; ByteaAsLongVarBinary=1;"

 conexao2.Open strConn
   
   
   End If
   
   
   
   Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
a = "te_layer"
b = "te_layer_table"

e = "layer_id"
f = "name"
   If TypeConn <> 4 Then
    Set rs = conn.Execute("Select * from te_layer l inner join te_layer_table t on t.layer_id=l.layer_id where l.name='" & tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName) & "'")
    'DE ACORDO COM A VARIAVEL NOME DO LAYER ATIVO, O SELECT ACIMA RETORNA UM RESULTADO CONFORME ABAIXO
    'LAYER_ID  PROJECTION_ID  NAME       LOWER_X   LOWER_Y    UPPER_X   UPPER_Y    INITIAL_TIME  FINAL_TIME TABLE_ID   LAYER_ID ATTR_TABLE UNIQUE_ID  ATTR_LINK   ATTR_INITIAL_TIME  ATTR_FINAL_TIME  ATTR_TIME_UNIT  ATTR_TABLE_TYPE  USER_NAME  INITIAL_TIME  FINAL_TIME
    
   '1         1              WATERLINES 292881,45 7422868,69 316554,47 7444782,19                          4          1        WATERLINES OBJECT_ID_ OBJECT_ID_                                      1               1

Dim az As String
az = "Select * from te_layer l inner join te_layer_table t on t.layer_id=l.layer_id where l.name='" & tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName) & "'"

'MsgBox "ARQUIVO DEBUG SALVO"
' WritePrivateProfileString "A", "A", az, App.Path & "\DEBUG.INI"




Else
 Set rs = conn.Execute("Select * from " + """" + a + """" + "  inner join " + """" + b + """" + " on " + """" + b + """" + "." + """" + e + """" + "=" + """" + a + """" + "." + """" + e + """" + " where " + """" + a + """" + "." + """" + f + """" + "='" & tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName) & "'")
    'DE




End If
 If Not rs.EOF Then
    rs.Close
  
  
     If TypeConn <> 4 Then
     
        Set rs = conn.Execute(getPmsdp(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName), 0, 0, TypeConn, conn))
        
        If UCase(cboColunas.Text) = "USUÁRIO" Then ' caso seja usuário, ignora o indice do combo filtro tratando como texto
            sql = "A." & rs.Fields(cboColunas.ListIndex).Name & GetOperador(cboOperador.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro.Text, "yyyymmdd"), cboFiltro.Text) & "''"
        
        ElseIf cboColunas.ListIndex >= 0 And cboFiltro.ListIndex >= 0 Then
            sql = "A." & rs.Fields(cboColunas.ListIndex).Name & GetOperador(cboOperador.ListIndex) & cboFiltro.ItemData(cboFiltro.ListIndex)

        ElseIf cboColunas.ListIndex >= 0 And cboFiltro.Text <> "" Then
            sql = "A." & rs.Fields(cboColunas.ListIndex).Name & GetOperador(cboOperador.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro.Text, "yyyymmdd"), cboFiltro.Text) & "''"
        
        End If
        

'MsgBox "primeiro sql = " & sql
        
        If UCase(cboColunas2.Text) = "USUÁRIO" Then  ' caso seja usuário, ignora o indice do combo filtro tratando como texto
            If sql <> "" Then
                sql = sql & " AND "
            End If
            sql = sql & "A." & rs.Fields(cboColunas2.ListIndex).Name & GetOperador(cboOperador2.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas2.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro2.Text, "yyyymmdd"), cboFiltro2.Text) & "''"
        
        ElseIf cboColunas2.ListIndex >= 0 And cboFiltro2.ListIndex >= 0 Then
            If sql <> "" Then
                sql = sql & " AND "

            End If
            sql = sql & "A." & rs.Fields(cboColunas2.ListIndex).Name & GetOperador(cboOperador2.ListIndex) & cboFiltro2.ListIndex
        ElseIf cboColunas2.ListIndex >= 0 And cboFiltro2.Text <> "" Then
            If sql <> "" Then
                sql = sql & " AND "
            End If
            sql = sql & "A." & rs.Fields(cboColunas2.ListIndex).Name & GetOperador(cboOperador2.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas2.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro2.Text, "yyyymmdd"), cboFiltro2.Text) & "''"
        End If
        
        Else
        
        '111111111111111
        
                Set rs = conn.Execute(getPmsdp(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName), 0, 0, TypeConn, conn))
        
        If UCase(cboColunas.Text) = "USUÁRIO" Then ' caso seja usuário, ignora o indice do combo filtro tratando como texto
            sql = "A." + """" & rs.Fields(cboColunas.ListIndex).Name & GetOperador(cboOperador.ListIndex) + """" & " ''" & IIf(UCase(rs.Fields(cboColunas.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro.Text, "yyyymmdd"), cboFiltro.Text) & "''"
        
        ElseIf cboColunas.ListIndex >= 0 And cboFiltro.ListIndex >= 0 Then
            sql = "A." + """" + rs.Fields(cboColunas.ListIndex).Name + """" & GetOperador(cboOperador.ListIndex) & cboFiltro.ItemData(cboFiltro.ListIndex) + """"

        ElseIf cboColunas.ListIndex >= 0 And cboFiltro.Text <> "" Then
            sql = "A." + """" & rs.Fields(cboColunas.ListIndex).Name + """" & GetOperador(cboOperador.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro.Text, "yyyymmdd"), cboFiltro.Text) & "''"
        
        End If
        

'MsgBox "primeiro sql = " & sql
        
        If UCase(cboColunas2.Text) = "USUÁRIO" Then  ' caso seja usuário, ignora o indice do combo filtro tratando como texto
            If sql <> "" Then
                sql = sql & " AND "
            End If
            sql = sql & "A." + """" & rs.Fields(cboColunas2.ListIndex).Name + """" & GetOperador(cboOperador2.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas2.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro2.Text, "yyyymmdd"), cboFiltro2.Text) & "''"
        
        ElseIf cboColunas2.ListIndex >= 0 And cboFiltro2.ListIndex >= 0 Then
            If sql <> "" Then
                sql = sql & " AND "

            End If
            sql = sql & "A." + """" & rs.Fields(cboColunas2.ListIndex).Name + """" & GetOperador(cboOperador2.ListIndex) & cboFiltro2.ListIndex
        ElseIf cboColunas2.ListIndex >= 0 And cboFiltro2.Text <> "" Then
            If sql <> "" Then
                sql = sql & " AND "
            End If
            sql = sql & "A." + """" & rs.Fields(cboColunas2.ListIndex).Name + """" & GetOperador(cboOperador2.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas2.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro2.Text, "yyyymmdd"), cboFiltro2.Text) & "''"
        End If
        
        
        End If
        
        

'MsgBox sql
    
        
        End If
 
a = "OBJECT_ID_"
b = UCase(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName))
c = b
d = UCase(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName))
e = "b"
f = "DATA"
   Dim sql2 As String
   'sql12 = sql
   If TypeConn <> 4 Then

        sql = "object_id in(select a.object_id_ from " & tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName) & " A left join " & tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName) & "data B on A.Object_id_=B.Object_id_ Where " & sql & ")"
          GetThemeWhere = convertQuery(sql, TypeConn)
        
      Else
       sql2 = """" + a + """" + " in(select " + """" + c + """" + "." + """" + a + """" + " from " + """" + c + """" + " left join " + """" + b + f + """" + " A on " + """" + c + """" + "." + """" + a + """" + "=" + """" + a + """" + " Where " & sql & ")"
     
     'MsgBox sql2
      End If
'MsgBox "sql FINAL da consulta " & sql

        'Momento em que a string de consulta é passada para gravação no banco de dados
Dim aa, bb As String
aa = "te_theme"
bb = "theme_id"


If TypeConn = 4 And ThemeName2 <> "WATERLINES" Then


    
            
            conexao2.Execute "UPDATE " + """" + aa + """" + " SET " + """" + "generate_attribute_where" + """" + " = '" & sql2 & "' WHERE " + """" + bb + """" + "='" & intTema & "'"
           

 End If

''' Codigo intermediário 06/01/09
'''    If Not rs.EOF Then
'''        rs.Close
'''        Set rs = conn.Execute(getPmsdp(tvm.getLayerNameFromTheme(tvm.getactiveview, ThemeName), 0, 0, TypeConn, conn))
'''
'''        If UCase(cboColunas.Text) = "USUÁRIO" Then ' caso seja usuário, ignora o indice do combo filtro tratando como texto
'''            sql = "A." & rs.Fields(cboColunas.ListIndex).Name & GetOperador(cboOperador.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro.Text, "yyyymmdd"), cboFiltro.Text) & "''"
'''
'''        ElseIf cboColunas.ListIndex >= 0 And cboFiltro.ListIndex >= 0 Then
'''            sql = "A." & rs.Fields(cboColunas.ListIndex).Name & GetOperador(cboOperador.ListIndex) & cboFiltro.ItemData(cboFiltro.ListIndex)
'''
'''        ElseIf cboColunas.ListIndex >= 0 And cboFiltro.Text <> "" Then
'''            sql = "A." & rs.Fields(cboColunas.ListIndex).Name & GetOperador(cboOperador.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro.Text, "yyyymmdd"), cboFiltro.Text) & "''"
'''
'''        End If
'''
''''MsgBox "primeiro sql = " & sql
'''
'''        If UCase(cboColunas2.Text) = "USUÁRIO" Then  ' caso seja usuário, ignora o indice do combo filtro tratando como texto
'''            sql = sql & "A." & rs.Fields(cboColunas2.ListIndex).Name & GetOperador(cboOperador2.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas2.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro2.Text, "yyyymmdd"), cboFiltro2.Text) & "''"
'''
'''        ElseIf cboColunas2.ListIndex >= 0 And cboFiltro2.ListIndex >= 0 Then
'''            If sql <> "" Then
'''                sql = sql & " AND "
'''
'''            End If
'''            sql = sql & "A." & rs.Fields(cboColunas2.ListIndex).Name & GetOperador(cboOperador2.ListIndex) & cboFiltro2.ListIndex
'''        ElseIf cboColunas2.ListIndex >= 0 And cboFiltro2.Text <> "" Then
'''            If sql <> "" Then
'''                sql = sql & " AND "
'''            End If
'''            sql = sql & "A." & rs.Fields(cboColunas2.ListIndex).Name & GetOperador(cboOperador2.ListIndex) & "''" & IIf(UCase(rs.Fields(cboColunas2.ListIndex).Name) = UCase("DateInstallation"), Format(cboFiltro2.Text, "yyyymmdd"), cboFiltro2.Text) & "''"
'''        End If
'''
''''MsgBox "segundo sql = " & sql
'''
'''        sql = "object_id in(select a.object_id_ from " & tvm.getLayerNameFromTheme(tvm.getactiveview, ThemeName) & " A left join " & tvm.getLayerNameFromTheme(tvm.getactiveview, ThemeName) & "data B on A.Object_id_=B.Object_id_ Where " & sql & ")"
'''
''''MsgBox "sql FINAL da consulta " & sql
'''
'''        'Momento em que a string de consulta é passada para gravação no banco de dados
'''        GetThemeWhere = convertQuery(sql, TypeConn)
'''
'''   End If
   
   rs.Close
   conn.Close
   conexao2.Close
   
   conn.Open
    conexao2.Open
   Exit Function
GetThemeWhere_err:
   MsgBox Err.Description & vbCrLf & sql
End Function

'Private Sub GetRepByTheme(mtheme As String)
'
'   chkVisibledPolyguns.Value = 0
'   chkVisibledLine.Value = 0
'   chkVisibledPoint.Value = 0
'   chkVisibledText.Value = 0
'   Select Case tvm.getThemeRepresentation(tvm.getActiveView, mtheme)
'     Case 1, 5, 7, 129, 131, 135          'Poligonus
'        chkVisibledPolyguns.Value = 1
'   End Select
'   Select Case tvm.getThemeRepresentation(tvm.getActiveView, mtheme)
'     Case 2, 3, 6, 7, 130, 134, 135 'lines
'        chkVisibledLine.Value = 1
'   End Select
'   Select Case tvm.getThemeRepresentation(tvm.getActiveView, mtheme)
'     Case 4, 5, 6, 7, 132, 134, 135
'        chkVisibledPoint.Value = 1
'   End Select
'   Select Case tvm.getThemeRepresentation(tvm.getActiveView, mtheme)
'     Case Is > 128
'        chkVisibledText.Value = 1
'   End Select
'
'End Sub

Private Sub cboColunas_Click()
    On Error GoTo Trata_Erro
    Dim a As String
    Dim b As String
    Dim c As String

    Me.cboFiltro.Clear
    If LayerAtivo = "RAMAIS_AGUA" Then
        'POSSIBILIDADE DE FILTROS POR:
        'cboColunas.AddItem "TIPO"
        'cboColunas.AddItem "HIDROMETRADO"
        'cboColunas.AddItem "ECONOMIAS"
        'cboColunas.AddItem "CONSUMO_LPS"
        'cboColunas.AddItem "DISTANCIA_TESTADA"
        'cboColunas.AddItem "DISTANCIA_LADO"
        'cboColunas.AddItem "COMPRIMENTO_RAMAL"
        'cboColunas.AddItem "PROFUNDIDADE_RAMAL"
        'cboColunas.AddItem "USUARIO_LOG"
        FILT = Me.cboColunas.Text
        If FILT = "TIPO" Or FILT = "HIDROMETRADO" Or FILT = "ECONOMIAS" Or FILT = "CONSUMO_LPS" Then 'RAMAIS_AGUA_LIGACAO
            tabela = "RAMAIS_AGUA_LIGACAO"
        ElseIf FILT = "DISTANCIA_TESTADA" Or FILT = "DISTANCIA_LADO" Or FILT = "COMPRIMENTO_RAMAL" Or FILT = "PROFUNDIDADE_RAMAL" Or FILT = "USUARIO_LOG" Then
            tabela = "RAMAIS_AGUA"
        End If
        Set rs = New ADODB.Recordset
        If TypeConn <> 4 Then
            rs.Open "SELECT DISTINCT " & Me.cboColunas.Text & " FROM " & tabela & " order by " & Me.cboColunas.Text, conn, adOpenForwardOnly, adLockReadOnly
        Else
            rs.Open "SELECT DISTINCT " + """" + Me.cboColunas.Text + """" + " FROM " + """" + tabela + """" + "order by " + """" + Me.cboColunas.Text + """", conn, adOpenDynamic, adLockOptimistic
        End If
        If rs.EOF = False Then
            Do While Not rs.EOF
                If Len(rs.Fields(0).Value) > 0 Then
                    Me.cboFiltro.AddItem rs.Fields(0).Value
                End If
                rs.MoveNext
            Loop
        End If
    Else
        Dim RsRef As ADODB.Recordset
        cboFiltro.Clear
        If cboColunas.ListIndex <> -1 Then
            If UCase(cboColunas.Text) = "USUÁRIO" Then
                a = "USRLOG"
                b = "SYSTEMUSERS"
                c = "USRLOG"
                If TypeConn <> 4 Then
                    Set rs = conn.Execute("SELECT USRLOG FROM SYSTEMUSERS ORDER BY USRLOG")
                Else
                    Set rs = conn.Execute("SELECT " + """" + a + """" + " FROM " + """" + b + """" + " ORDER BY " + """" + c + """" + "")
                End If
                If Not rs.EOF Then
                    While Not rs.EOF
                        cboFiltro.AddItem rs!USRLOG
                        rs.MoveNext
                    Wend
                End If
            Else
                a = "X_MANAGERPROPERTIES"
                b = "TABLENAMEIN"
                c = "FIELDNAMERIN"
                If TypeConn <> 4 Then
                    Set rs = conn.Execute("Select * from X_ManagerProperties Where Upper(TableNamein)='" & UCase(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName)) & _
                    "' and Upper(FieldNameRin) ='" & UCase(cboColunas.Text) & "'")
                Else
                    ' Dim d2 As String
                    'd2 = "Select * from " + """" + a + """" + " Where" + """" + b + """" + "='" & tvm.getLayerNameFromTheme(tvm.getactiveview, ThemeName) + _
                    '             "' and " + c + " ='" & cboColunas.Text & "'"
                    ' MsgBox d2
                    Set rs = conn.Execute("Select * from " + """" + a + """" + " Where" + """" + b + """" + "='" + tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName) + "' and " + """" + c + """" + " ='" & cboColunas.Text & "'")
                End If
                ' Dim a33 As String
                ' a33 = "Select * from " + """" + a + """" + " Where" + """" + b + """" + "='" & tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName) + _
                '  "' and " + """" + c + """" + " ='" & cboColunas.Text & "'"
                'MsgBox "ARQUIVO DEBUG SALVO"
                'WritePrivateProfileString "A", "A", a33, App.Path & "\DEBUG.INI"
                If Not rs.EOF Then
                    If TypeConn <> 4 Then
                        Set RsRef = conn.Execute("SELECT * FROM " & rs!TABLENAMEOUT)
                    Else
                        Set RsRef = conn.Execute("SELECT * FROM " + """" + UCase(rs!TABLENAMEOUT) + """")
                    End If
                    While Not RsRef.EOF
                        cboFiltro.AddItem RsRef(1)
                        cboFiltro.ItemData(cboFiltro.NewIndex) = RsRef(0)
                        RsRef.MoveNext
                    Wend
                    RsRef.Close
                End If
            End If
        Else
            cboOperador.ListIndex = -1
            cboFiltro.Clear
        End If
    End If
    rs.Close
    conn.Close
    conn.Open
    Set RsRef = Nothing
    Set rs = Nothing
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       MsgBox Err.Number & " - " & Err.Description
    End If
End Sub


Private Sub cboColunas2_Click()
On Error GoTo cboColunas2_Click_err
   
    Dim RsRef As ADODB.Recordset
    cboFiltro2.Clear
    
    If cboColunas2.ListIndex <> -1 Then
        If UCase(cboColunas2.Text) = "USUÁRIO" Then
      Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String

a = "USRLOG"
b = "SYSTEMUSERS"
c = "USRLOG"
   If TypeConn <> 4 Then
            Set rs = conn.Execute("SELECT USRLOG FROM SYSTEMUSERS ORDER BY USRLOG")
            Else
           Set rs = conn.Execute("SELECT " + """" + a + """" + " FROM " + """" + b + """" + " ORDER BY " + """" + c + """" + "")
            End If
            If Not rs.EOF Then
                 While Not rs.EOF
                     cboFiltro2.AddItem rs!USRLOG
                     rs.MoveNext
                 Wend
            End If
            rs.Close
            conn.Close
            conn.Open
        
        Else
        
  
a = "X_MANAGERPROPERTIES"
b = "TABLENAMEIN"
c = "FIELDNAMERIN"
   If TypeConn <> 4 Then
            Set rs = conn.Execute("Select * from X_ManagerProperties Where Upper(TableNamein)='" & UCase(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName)) & _
                         "' and Upper(FieldNameRin) ='" & UCase(cboColunas2.Text) & "'")
            
            If Not rs.EOF Then
            If TypeConn <> 4 Then
                 Set RsRef = conn.Execute("Select * from " & rs!TABLENAMEOUT)
                 Else
                    Set RsRef = conn.Execute("Select * from " & UCase(rs!TABLENAMEOUT))
                 End If
                 While Not RsRef.EOF
                     cboFiltro2.AddItem RsRef(1)
                     cboFiltro2.ItemData(cboFiltro2.NewIndex) = RsRef(0)
                     RsRef.MoveNext
                 Wend
                 RsRef.Close
            End If
            rs.Close
            conn.Close
            conn.Open
            
            Else
             Set rs = conn.Execute("Select * from " + """" + a + """" + " Where " + """" + b + """" + "='" & tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName) & _
                         "' and " + """" + c + """" + " ='" & cboColunas2.Text & "'")
                         
                Dim zaza As String
            
                         
                         
            
            If Not rs.EOF Then
                 Set RsRef = conn.Execute("Select * from " + """" + rs!TABLENAMEOUT + """")
                 While Not RsRef.EOF
                     cboFiltro2.AddItem RsRef(1)
                     cboFiltro2.ItemData(cboFiltro2.NewIndex) = RsRef(0)
                     RsRef.MoveNext
                 Wend
                 RsRef.Close
            End If
            rs.Close
            conn.Close
            conn.Open
            
            End If
            
        End If
    Else
        cboOperador2.ListIndex = -1
        cboFiltro2.Clear
    End If
   
  
    Set RsRef = Nothing
    Set rs = Nothing
   
   
Exit Sub
cboColunas2_Click_err:
   MsgBox Err.Description
End Sub
'Carrega os atributos de filtros possíveis para o tema selecionado, na caixa de diálogo para que o usuário possa selecioná-lo
'e realizar o filtro pelo mesmo
'
'
Private Sub LoadCboAttribute()
    Dim AttributeTable As String, AttributeLink As String, strCMD As String
    'LIMPA E CARREGA OS COMBOS DE ATRIBUTOS
    cboColunas.Clear
    cboColunas2.Clear
    'XXXXXX  AQUI O LAYER ATIVO PARA QUANTO SELECIONA RAMAIS, ESTÁ VINDO COMO WATERCOMPONENTS, QUANDO CORRIGIDO, RETIRAR ESTA LINHA
    'Agora iremos verificar pelos ifs qual o tipo de layer que estamos tratando
    If LayerAtivo = "RAMAIS_AGUA" Then
        '*** TRATANDO RAMAIS DE ÁGUA
        'aqui os atributos estão sendo carregados em hardcode
        cboColunas.AddItem "TIPO"
        cboColunas.AddItem "HIDROMETRADO"
        cboColunas.AddItem "ECONOMIAS"
        cboColunas.AddItem "CONSUMO_LPS"
        cboColunas.AddItem "DISTANCIA_TESTADA"
        cboColunas.AddItem "DISTANCIA_LADO"
        cboColunas.AddItem "COMPRIMENTO_RAMAL"
        cboColunas.AddItem "PROFUNDIDADE_RAMAL"
        cboColunas.AddItem "USUARIO_LOG"
        ' strCMD = getPmsdp(tvm.getLayerNameFromTheme(tvm.getactiveview, ThemeName), 2, 0, TypeConn, conn)
    Else
        ThemeName2 = tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName)
        If (ThemeName2 = "WATERLINES" Or ThemeName2 = "SEWERLINES") Then
            '*** TRATANDO REDES DE ÁGUA OU ESGOTO
            'carrega os atributos de redes
            If RetornaNomeAtr(conn, tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName), AttributeTable, AttributeLink) Then
                strCMD = ""
                strCMD = getPmsdp(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName), 2, 0, TypeConn, conn)
                strCMD = UCase(strCMD)
                Set rs = New ADODB.Recordset
                ' MsgBox "ARQUIVO DEBUG SALVO"
                'WritePrivateProfileString "A", "A", strCMD, App.Path & "\DEBUG.INI"
                Set rs = conn.Execute(strCMD)
                For a = 0 To rs.Fields.Count - 1
                    cboColunas.AddItem rs(a).Name
                    cboColunas.ItemData(cboColunas.NewIndex) = a
                    cboColunas2.AddItem rs(a).Name
                    cboColunas2.ItemData(cboColunas2.NewIndex) = a
                Next
            End If
        ElseIf (ThemeName2 = "WATERCOMPONENTS" Or ThemeName2 = "SEWERCOMPONENTS") Then
            '*** TRATANDO NÓS DAS REDES
            'carrega os atributos de nós
            If RetornaNomeAtr(conn, tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName), AttributeTable, AttributeLink) Then
                strCMD = ""
                strCMD = getPmsdp2(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName), 2, 0, TypeConn, conn)
                strCMD = UCase(strCMD)
                Set rs = New ADODB.Recordset
                ' MsgBox "ARQUIVO DEBUG SALVO"
                'WritePrivateProfileString "A", "A", strCMD, App.Path & "\DEBUG.INI"
                Set rs = conn.Execute(strCMD)
                For a = 0 To rs.Fields.Count - 1
                    cboColunas.AddItem rs(a).Name
                    cboColunas.ItemData(cboColunas.NewIndex) = a
                    cboColunas2.AddItem rs(a).Name
                    cboColunas2.ItemData(cboColunas2.NewIndex) = a
                Next
            End If
        End If
    End If
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If

End Sub

Private Sub cboFiltro_Click()
On Error GoTo Trata_Erro

   cboColunasSub.Clear
   cboFiltroSub.Clear
   cboOperadorSub.ListIndex = -1
   Dim Rsx As ADODB.Recordset
   Dim rsSubTypes As Recordset
   
   Dim strCommand As String
   
   Dim a As String
Dim b As String
Dim c As String

a = "SEWERCOMPONENTSSUBTYPES"
b = "USER_NAME"
c = "NAME"
   If TypeConn <> 4 Then
   Set rsSubTypes = conn.Execute("SELECT * FROM SEWERCOMPONENTSSubTypes") ' se não houver registros na tabela de subtipos, sai fora
   Else
   
    Set rsSubTypes = conn.Execute("SELECT * FROM " + """" + a + """" + "")
   End If
   If rsSubTypes.EOF = False Then
      strCommand = getPmssp(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName), cboFiltro.ItemData(cboFiltro.ListIndex), 0, conn, TypeConn)
      Set Rsx = conn.Execute(strCommand)
      If Rsx.EOF = False Then
         While Not Rsx.EOF
            cboColunasSub.AddItem Rsx(5)
            cboColunasSub.ItemData(cboColunasSub.NewIndex) = Rsx(7)
            Rsx.MoveNext
         Wend
         cboColunasSub.Enabled = True
         cboOperadorSub.Enabled = True
         cboFiltroSub.Enabled = True
      Else
         cboColunasSub.Enabled = False
         cboOperadorSub.Enabled = False
         cboFiltroSub.Enabled = False
      End If
      Rsx.Close
      Set Rsx = Nothing
   End If
   rsSubTypes.Close
   Set rsSubTypes = Nothing
           
'      Close #1
'      Open App.Path & "\GeoSanLog.txt" For Append As #1
'      Print #1, Now & " - " & strCommand
'      Close #1
   


Trata_Erro:

    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #1
        Open App.Path & "\GeoSanLog.txt" For Append As #1
        Print #1, Now & " - NxViewManager - frmTheme - Private Sub cboFiltro_Click() - " & Err.Number & " - " & Err.Description
        Close #1
        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrencia.", vbInformation
    End If
End Sub

Private Sub LoadCboAttributeAdc()
    cboColunasSub.Clear
    Set rs = conn.Execute(getPmssp(tvm.getLayerNameFromTheme(tvm.getActiveView, ThemeName), cboAttributeValue.ItemData(cboAttributeValue.ListIndex), 0, conn, TypeConn))
    While Not rs.EOF
        cboColunasSub.AddItem rs(5).Value
        cboColunasSub.ItemData(cboColunasSub.NewIndex) = rs(7).Value
        rs.MoveNext
    Wend
    rs.Close
    rs.Close
    conn.Close
    conn.Open
   
   Set rs = Nothing
End Sub

'Private Sub cboColunasSub_Click()
'   cboFiltroSub.Clear
'   If cboColunasSub.ListIndex = -1 Then Exit Sub
'   Set rs = conn.Execute("select value_, Option_ from " & tvm.getLayerNameFromTheme(tvm.getactiveview, ThemeName) & "selections where id_type = " & cboFiltro.ItemData(cboFiltro.ListIndex) & " and id_subtype = " & cboColunasSub.ItemData(cboColunasSub.ListIndex) & " order by option_")
'   While Not rs.EOF
'      cboFiltroSub.AddItem rs(1)
'      cboFiltroSub.ItemData(cboFiltroSub.NewIndex) = rs(0)
'      rs.MoveNext
'   Wend
'   rs.Close
'   Set rs = Nothing
'End Sub

Private Function GetOperador(Index As Integer) As String
   Select Case Index
      Case -1 'Não selecionado = -1
         GetOperador = "="
      Case 0 '"Igual"
         GetOperador = "="
      Case 1 '"Maior"
         GetOperador = ">"
      Case 2 '"Menor"
         GetOperador = "<"
      Case 3 '"Diferente"
         GetOperador = "<>"
   End Select
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

