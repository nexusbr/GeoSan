VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmResources 
   Caption         =   "Recursos do Sistema"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   Icon            =   "FrmResources.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   210
         TabIndex        =   2
         Top             =   2850
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2250
         TabIndex        =   1
         Top             =   2850
         Width           =   855
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2730
         Top             =   1980
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmResources.frx":0320
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmResources.frx":1199
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView LvResources 
         Height          =   2535
         Left            =   210
         TabIndex        =   3
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private L As ListItems, SelectT As Boolean
Public Function Init(Conn As ADODB.Connection, GrpId As Integer, ByRef Itmx As ListItems) As Boolean
   Dim Usr As New NexusUsers.clsUsers, Rs As ADODB.Recordset, Rs_ As ADODB.Recordset
   Set Rs_ = New ADODB.Recordset
   Rs_.CursorLocation = adUseClient
   Rs_.LockType = adLockOptimistic
   Dim I As ListItem
   Set Rs = Usr.Resources.SelectResources(Conn)
   Set Rs_ = Usr.ResourcesGroups.SelectResourcesByGroup(Conn, GrpId)
   
   LvResources.ColumnHeaders.Item(1).Width = LvResources.Width
   
   While Not Rs.EOF
      Rs_.Filter = "rcsid=" & Rs.Fields("rcsid")
      If Rs_.EOF Then
         Set I = LvResources.ListItems.Add(, , Rs.Fields("RcsNom"), , 2)
            I.Tag = Rs.Fields("RcsId")
      End If
      Rs.MoveNext
   Wend
   Rs.Close
   Rs_.Close
   Set Rs_ = Nothing
   Set Rs = Nothing
   Me.Show vbModal
   If SelectT Then
      Set Itmx = L
      Init = SelectT
   End If
   Set Usr = Nothing
End Function

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOk_Click()
   If Not (LvResources.SelectedItem Is Nothing) Then
      Dim A As Integer
      For A = 1 To LvResources.ListItems.Count
         If Not LvResources.ListItems.Item(A).Selected Then
            LvResources.ListItems.Item(A).Tag = ""
         End If
      Next
      
      Set L = LvResources.ListItems
      SelectT = True
      Unload Me
   Else
      MsgBox "Selecione o recurso desejado e click em Ok, ou cancele para sair", vbInformation
   End If
End Sub


