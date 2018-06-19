VERSION 5.00
Begin VB.Form FrmSuppliers 
   Caption         =   "Fornecedores"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "FrmSuppliers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Height          =   345
      Left            =   5100
      Picture         =   "FrmSuppliers.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salvar"
      Top             =   6270
      Width           =   435
   End
   Begin VB.CommandButton cmdNew 
      Height          =   345
      Left            =   3900
      Picture         =   "FrmSuppliers.frx":0B6C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Novo"
      Top             =   6270
      Width           =   435
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   345
      Left            =   4500
      Picture         =   "FrmSuppliers.frx":13CE
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Editar"
      Top             =   6270
      Width           =   435
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   60
      TabIndex        =   9
      Top             =   90
      Width           =   6015
      Begin VB.TextBox mskFax 
         Height          =   285
         Left            =   1800
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   5160
         Width           =   1815
      End
      Begin VB.TextBox mskPhone 
         Height          =   285
         Left            =   1800
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox mskPostalCode 
         Height          =   285
         Left            =   1800
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtHomePage 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   5640
         Width           =   4095
      End
      Begin VB.TextBox txtRegion 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox txtContactTitle 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox txtContactName 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtSupplierId 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Text            =   "0"
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtCountry 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Fornecedor"
         Height          =   195
         Left            =   600
         TabIndex        =   24
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Companhia"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Contato"
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cargo do Contato"
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   960
         TabIndex        =   20
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Região"
         Height          =   195
         Left            =   1080
         TabIndex        =   18
         Top             =   3240
         Width           =   510
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         Height          =   195
         Left            =   1320
         TabIndex        =   16
         Top             =   5160
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Home Page"
         Height          =   195
         Left            =   720
         TabIndex        =   15
         Top             =   5640
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "País"
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Top             =   4200
         Width           =   330
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Left            =   960
         TabIndex        =   13
         Top             =   4680
         Width           =   630
      End
   End
   Begin VB.Label Label15 
      Caption         =   "Salvar"
      Height          =   255
      Left            =   5310
      TabIndex        =   27
      Top             =   6600
      Width           =   465
   End
   Begin VB.Label Label14 
      Caption         =   "Editar"
      Height          =   255
      Left            =   4740
      TabIndex        =   26
      Top             =   6600
      Width           =   465
   End
   Begin VB.Label Label13 
      Caption         =   "Novo"
      Height          =   255
      Left            =   4110
      TabIndex        =   25
      Top             =   6600
      Width           =   465
   End
End
Attribute VB_Name = "FrmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'editar dados do fornecedor
Private Sub cmdEdit_Click()
Dim Sup_ID As Long
Dim frm As FrmSuppliersSub
Set frm = New FrmSuppliersSub
Sup_ID = frm.init("X_Suppliers")
LoadSupplier Sup_ID
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

   Dim ClsSup As ClsX_Suppliers
   Set ClsSup = New ClsX_Suppliers
   ClsSup.Address = txtAddress.Text
   ClsSup.City = txtCity.Text
   ClsSup.CompanyName = txtCompanyName.Text
   ClsSup.ContactName = txtContactName.Text
   ClsSup.ContactTitle = txtContactTitle.Text
   ClsSup.Fax = mskFax.Text
   ClsSup.HomePage = txtHomePage.Text
   ClsSup.PostalCode = mskPostalCode.Text
   ClsSup.Region = txtRegion.Text
   ClsSup.SupplierID = txtSupplierId.Text
   ClsSup.Country = txtCountry.Text
   ClsSup.Phone = mskPhone.Text
   ClsSup.InsertData Conn
   Set ClsSup = Nothing

   ClearFields
   MsgBox "Os dados foram gravados com sucesso", vbExclamation, "GeoSan"
   Exit Sub
cmdSALVAR_ERROR:
   MsgBox Err.Description, vbExclamation

End Sub


'pesquisa dados
Private Sub LoadSupplier(Sup_ID As Long)
   Dim ClsSup As ClsX_Suppliers
   Set ClsSup = New ClsX_Suppliers
   ClsSup.SupplierID = Sup_ID
   ClsSup.UpdateData Conn
   txtAddress.Text = ClsSup.Address
   txtCity.Text = ClsSup.City
   txtCompanyName.Text = ClsSup.CompanyName
   txtContactName.Text = ClsSup.ContactName
   txtContactTitle.Text = ClsSup.ContactTitle
   mskFax.Text = IIf(ClsSup.Fax = "", "(___)____-____", ClsSup.Fax)
   txtHomePage.Text = ClsSup.HomePage
   mskPostalCode.Text = IIf(ClsSup.PostalCode = "", "_____-___", ClsSup.PostalCode)
   txtRegion.Text = ClsSup.Region
   txtSupplierId.Text = ClsSup.SupplierID
   txtCountry.Text = ClsSup.Country
   mskPhone.Text = IIf(ClsSup.Phone = "", "(___)____-____", ClsSup.Phone)
   Set ClsSup = Nothing
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
   txtSupplierId = 0
   mskPhone.Text = "(___)____-____"
   mskFax.Text = "(___)____-____"
   mskPostalCode = "_____-___"
   txtCompanyName.SetFocus
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



