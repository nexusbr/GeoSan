VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportarCotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importação de Cotas"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDL 
      Left            =   1530
      Top             =   1290
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFormato 
      Caption         =   "Formato"
      Height          =   405
      Left            =   3855
      TabIndex        =   4
      Top             =   1185
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arquivo Fonte"
      Height          =   795
      Left            =   210
      TabIndex        =   2
      Top             =   225
      Width           =   5985
      Begin VB.CommandButton cmdProcuraArquivo 
         Caption         =   "..."
         Height          =   360
         Left            =   5340
         TabIndex        =   5
         Top             =   315
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   210
         TabIndex        =   3
         Top             =   315
         Width           =   5040
      End
   End
   Begin VB.CommandButton cmdImportarCotas 
      Caption         =   "Iniciar"
      Height          =   405
      Left            =   5025
      TabIndex        =   1
      Top             =   1185
      Width           =   1110
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   285
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmImportarCotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
Private Sub cmdFormato_Click()
   MsgBox "O arquivo deverá possuir duas colunas de informação separadas por ponto e vírgula (;)" & Chr(13) & Chr(13) & _
   Chr(13) & Chr(13) & "A primeira coluna com o identificador do Nó" & _
   Chr(13) & Chr(13) & "A segunda coluna com o valor de cota" & _
   Chr(13) & Chr(13) & "Ex." & _
   Chr(13) & Chr(13) & "66;2.3", vbInformation, "Formato de arquivo"

End Sub

Private Sub cmdImportarCotas_Click()
On Error GoTo Trata_Erro
   'O FORMATO DO ARQUIVO DEVE SER TEXTO SEPARADO POR ;
Dim count As Integer
   Dim str As String, id_no As String, cota As String, contador As Long
   Dim rs As New ADODB.Recordset
   count = 0
   contador = 0
   If Me.Text1.Text <> "" Then
      str = Dir(CStr(Text1.Text))
      If str = "" Then
         MsgBox "Arquivo inexistente.", vbInformation, ""
         Exit Sub
      End If
   Else
      MsgBox "Arquivo inexistente.", vbInformation, ""
      Exit Sub
   End If

   Open Text1.Text For Input As #3
   Do While Not EOF(3)
      Line Input #3, str
      contador = contador + 1
   Loop
   ProgressBar1.Max = contador + 1
   ProgressBar1.value = 1
   ProgressBar1.Visible = True
   
   Close #3
   
   MousePointer = vbHourglass
   Dim numero As String
   Dim MyArray As Variant
   
   Open Text1.Text For Input As #3
   Do While Not EOF(3)
      DoEvents
      Line Input #3, str
      MyArray = Split(str, ";")
      
    id_no = Trim(MyArray(0))
      cota = Replace(Trim(MyArray(1)), ",", ".")
      numero = id_no
a = "WATERCOMPONENTS"
b = "INITIALGROUNDHEIGHT"
c = "OBJECT_ID_"
count = count + 1
If count > 1 Then

      If frmCanvas.TipoConexao <> 4 Then
      Conn.execute ("UPDATE WATERCOMPONENTS SET GROUNDHEIGHT = " & cota & " WHERE OBJECT_ID_ = '" & id_no & "'")
      Else
      Conn.execute ("UPDATE " + """" + a + """" + " SET " + """" + b + """" + " = '" & Round(cota) & "' WHERE " + """" + c + """" + " = '" + numero + "'")
      
     ' Dim coo As String
     ' coo = "UPDATE " + """" + a + """" + " SET " + """" + b + """" + " = '" & Round(cota) & "' WHERE " + """" + c + """" + " = '" + numero + "'"

'MsgBox "ARQUIVO DEBUG SALVO"
' WritePrivateProfileString "A", "A", coo, App.path & "\DEBUG.INI"
      
      
      
      End If
      ProgressBar1.value = ProgressBar1.value + 1
 End If
   Loop
   
   Close #3
   
   MousePointer = vbDefault
  
   MsgBox "Cotas de Nós de Redes atualizadas com sucesso!", vbInformation, "Processo Concluído"

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   MousePointer = vbDefault
   
   'PrintErro CStr(Me.Name), "cmdImportarCotas", CStr(Err.Number), CStr(Err.Description), True
      MsgBox "Cotas de Nós de Redes atualizadas com sucesso!", vbInformation, "Processo Concluído"
End If

Unload Me

End Sub

Private Sub cmdProcuraArquivo_Click()
        

   CDL.DialogTitle = "Localizar Arquivo"
   CDL.ShowOpen
        
   Me.Text1.Text = CDL.FileName
   
End Sub




