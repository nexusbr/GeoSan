VERSION 5.00
Begin VB.Form FrmManufactures 
   Caption         =   "Fabricante"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "FrmManufactures.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Height          =   345
      Left            =   4560
      Picture         =   "FrmManufactures.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Editar"
      Top             =   6240
      Width           =   435
   End
   Begin VB.CommandButton Command2 
      Height          =   345
      Left            =   3960
      Picture         =   "FrmManufactures.frx":0B6C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Novo"
      Top             =   6240
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Height          =   345
      Left            =   5160
      Picture         =   "FrmManufactures.frx":13CE
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Salvar"
      Top             =   6240
      Width           =   435
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   30
      TabIndex        =   9
      Top             =   0
      Width           =   6015
      Begin VB.TextBox mskPostalCode 
         Height          =   315
         Left            =   1800
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox mskPhone 
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox mskFax 
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   5160
         Width           =   1815
      End
      Begin VB.TextBox txtCountry 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txtManufacturerId 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Text            =   "0"
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtContactName 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtContactTitle 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtRegion 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtHomePage 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   5640
         Width           =   4095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Left            =   960
         TabIndex        =   21
         Top             =   4680
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "País"
         Height          =   195
         Left            =   1320
         TabIndex        =   20
         Top             =   4200
         Width           =   330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Home Page"
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   5640
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         Height          =   195
         Left            =   1320
         TabIndex        =   18
         Top             =   5160
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Código Postal"
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Região"
         Height          =   195
         Left            =   1080
         TabIndex        =   16
         Top             =   3240
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   1080
         TabIndex        =   15
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   960
         TabIndex        =   14
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cargo do Contato"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Contato"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Companhia"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Fabricante"
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmManufactures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'editar dados do fornecedor
Private Sub cmdEdit_Click()
   Dim Man_ID As Long
   Dim frm As FrmSuppliersSub
   Set frm = New FrmSuppliersSub
   Man_ID = frm.init("X_Manufacturers")
   LoadManufacturer Man_ID
   Set frm = Nothing
End Sub


'inserir novo fornecedor
Private Sub cmdNew_Click()
   ClearFields
End Sub


'salvar dados do fornecedor
Private Sub cmdSave_Click()
On Error GoTo cmdSALVAR_ERROR
   If txtCompanyName = "" Then
      MsgBox "Nome da companhia obrigatório", vbExclamation, "GeoSan"
      Exit Sub
   End If

   Dim ClsMan As ClsX_Manufacturers
   Set ClsMan = New ClsX_Manufacturers
   ClsMan.Address = txtAddress.Text
   ClsMan.City = txtCity.Text
   ClsMan.CompanyName = txtCompanyName.Text
   ClsMan.ContactName = txtContactName.Text
   ClsMan.ContactTitle = txtContactTitle.Text
   ClsMan.Fax = mskFax.Text
   ClsMan.HomePage = txtHomePage.Text
   ClsMan.PostalCode = mskPostalCode.Text
   ClsMan.Region = txtRegion.Text
   ClsMan.ManufacturerID = txtManufacturerId.Text
   ClsMan.Country = txtCountry.Text
   ClsMan.Phone = mskPhone.Text
   ClsMan.InsertData Conn
   Set ClsMan = Nothing

   ClearFields
   MsgBox "Os dados foram gravados com sucesso", vbExclamation, "GeoSan"
   Exit Sub
cmdSALVAR_ERROR:
   MsgBox Err.Description, vbExclamation
End Sub


'pesquisa dados
Private Sub LoadManufacturer(Man_ID As Long)
   Dim ClsMan As ClsX_Manufacturers
   Set ClsMan = New ClsX_Manufacturers
   ClsMan.ManufacturerID = Man_ID
   ClsMan.UpdateData Conn
   txtAddress.Text = ClsMan.Address
   txtCity.Text = ClsMan.City
   txtCompanyName.Text = ClsMan.CompanyName
   txtContactName.Text = ClsMan.ContactName
   txtContactTitle.Text = ClsMan.ContactTitle
   mskFax.Text = IIf(ClsMan.Fax = "", "(___)____-____", ClsMan.Fax)
   txtHomePage.Text = ClsMan.HomePage
   mskPostalCode.Text = IIf(ClsMan.PostalCode = "", "_____-___", ClsMan.PostalCode)
   txtRegion.Text = ClsMan.Region
   txtManufacturerId.Text = ClsMan.ManufacturerID
   txtCountry.Text = ClsMan.Country
   mskPhone.Text = IIf(ClsMan.Phone = "", "(___)____-____", ClsMan.Phone)
   Set ClsMan = Nothing
End Sub


'limpar campos
Private Sub ClearFields()
   txtAddress = ""
   txtCity = ""
   txtCompanyName = ""
   txtContactName = ""
   txtContactTitle = ""
   txtCountry = ""
   txtHomePage = ""
   txtRegion = ""
   txtManufacturerId = 0
   mskPhone.Text = "(___)____-____"
   mskFax.Text = "(___)____-____"
   mskPostalCode = "_____-___"
   txtCompanyName.SetFocus
End Sub


Private Sub Command1_Click()
On Error GoTo cmdSALVAR_ERROR
   If txtCompanyName = "" Then
      MsgBox "Nome da companhia obrigatório", vbExclamation, "GeoSan"
      Exit Sub
   End If

   Dim ClsMan As ClsX_Manufacturers
   Set ClsMan = New ClsX_Manufacturers
   ClsMan.Address = txtAddress.Text
   ClsMan.City = txtCity.Text
   ClsMan.CompanyName = txtCompanyName.Text
   ClsMan.ContactName = txtContactName.Text
   ClsMan.ContactTitle = txtContactTitle.Text
   ClsMan.Fax = mskFax.Text
   ClsMan.HomePage = txtHomePage.Text
   ClsMan.PostalCode = mskPostalCode.Text
   ClsMan.Region = txtRegion.Text
   ClsMan.ManufacturerID = txtManufacturerId.Text
   ClsMan.Country = txtCountry.Text
   ClsMan.Phone = mskPhone.Text
   ClsMan.InsertData Conn
   Set ClsMan = Nothing

   ClearFields
   MsgBox "Os dados foram gravados com sucesso", vbExclamation, "GeoSan"
   Exit Sub
cmdSALVAR_ERROR:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub Command2_Click()
 ClearFields
End Sub

Private Sub Command3_Click()
   Dim Man_ID As Long
   Dim frm As FrmSuppliersSub
   Set frm = New FrmSuppliersSub
   Man_ID = frm.init("X_Manufacturers")
   LoadManufacturer Man_ID
   Set frm = Nothing
End Sub

Private Sub Form_Load()
   'LoozeXP1.InitSubClassing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'LoozeXP1.EndWinXPCSubClassing
End Sub

Private Sub mskPhone_GotFocus()
   mskPhone.SelStart = 0
   mskPhone.SelLength = Len(mskPhone.Text)
End Sub

Private Sub mskFax_GotFocus()
   mskFax.SelStart = 0
   mskFax.SelLength = Len(mskFax.Text)
End Sub

Private Sub mskPostalCode_GotFocus()
   mskPostalCode.SelStart = 0
   mskPostalCode.SelLength = Len(mskPostalCode.Text)
End Sub


