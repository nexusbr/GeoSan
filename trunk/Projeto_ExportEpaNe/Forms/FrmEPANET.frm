VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Begin VB.Form FrmEPANET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Exportação EPANET"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   ControlBox      =   0   'False
   Icon            =   "FrmEPANET.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Caminho de Exportação"
      Height          =   990
      Left            =   120
      TabIndex        =   4
      Top             =   210
      Width           =   6165
      Begin VB.TextBox txtArquivo 
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   375
         Width           =   5325
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   330
         Left            =   5550
         TabIndex        =   5
         Top             =   375
         Width           =   435
      End
   End
   Begin VB.TextBox txtTimer 
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Text            =   "20:00:00"
      Top             =   1335
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3450
      Top             =   1305
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3420
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   5190
      TabIndex        =   1
      Top             =   1335
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4035
      TabIndex        =   0
      Top             =   1335
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   165
      TabIndex        =   7
      Top             =   1335
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
      Left            =   4680
      OleObjectBlob   =   "FrmEPANET.frx":1CFA
      Top             =   120
   End
   Begin VB.Label Label4 
      Caption         =   "Horário"
      Height          =   225
      Left            =   645
      TabIndex        =   3
      Top             =   1395
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "FrmEPANET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public conn As ADODB.Connection
Public Provider As Integer
Public PLANO As String

Private rsTP As ADODB.Recordset
Private rsST As ADODB.Recordset

Dim i As Integer


Public Sub init()
   
   cmdConfirmar.Default = True
   
   txtArquivo.Text = App.Path & "\GEOEXP_EPANET" & Format(Now, "HHMM") & ".INP"
   
   Me.Show

End Sub



Private Sub cmdCancelar_Click()
   
   Cancelar = True
   
   Unload Me
End Sub



Private Function INICIAR()
On Error GoTo Trata_Erro
   
   Dim retval As String
   Dim usuario As String
   retval = Dir("C:\ARQUIVOS DE PROGRAMAS\GEOSAN\Controles\UserLog.txt")
   If retval <> "" Then 'verifica se o arquivo existe na pasta
      Open "C:\ARQUIVOS DE PROGRAMAS\GEOSAN\Controles\UserLog.txt" For Input As #3

      Line Input #3, usuario
      
      Close #3
   Else
      MsgBox "É necessário criar a seleção por polígono.", vbOKOnly + vbInformation, "Mensagem"
      End
   End If


   MousePointer = vbHourglass
   
   If conn.Provider <> "PostgreSQL.1" Then
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 111 WHERE MATERIAL = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 1 AND ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 120 WHERE MATERIAL = 2 AND ROUGHNESS = 0 ")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 110 WHERE MATERIAL = 3 AND ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 105 WHERE MATERIAL = 4 AND ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 90 WHERE MATERIAL = 5 AND ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 6 AND ROUGHNESS = 0")
   Else
     conn.Execute ("UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '111' WHERE " + """" + "MATERIAL" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '1' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '120' WHERE " + """" + "MATERIAL" + """" + " = '2' AND " + """" + "ROUGHNESS" + """" + " = '0' ")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '110' WHERE " + """" + "MATERIAL" + """" + " = '3' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '105' WHERE " + """" + "MATERIAL" + """" + " = '4' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '90' WHERE " + """" + "MATERIAL" + """" + " = '5' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '6' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   
   
   End If
   FrmEPANET.MousePointer = vbDefault
   
   
   
   Dim Rs As ADODB.Recordset
   Dim str As String
   Dim Tipo As String
   Dim setor As String
   Dim strtot As String
   

'   For i = 1 To lvTipoRede.ListItems.Count
'      If lvTipoRede.ListItems.Item(i).Checked Then
'         If Tipo = "" Then
'            Tipo = lvTipoRede.ListItems.Item(i).Tag
'         Else
'            Tipo = Tipo & "," & lvTipoRede.ListItems.Item(i).Tag
'         End If
'      End If
'   Next
'
'   For i = 1 To lvSetor.ListItems.Count
'      If lvSetor.ListItems.Item(i).Checked Then
'         If setor = "" Then
'            setor = lvSetor.ListItems.Item(i).Tag
'         Else
'            setor = setor & "," & lvSetor.ListItems.Item(i).Tag
'         End If
'      End If
'   Next
     If conn.Provider <> "PostgreSQL.1" Then
   conn.Execute ("UPDATE WATERLINES SET MATERIAL = 0 WHERE MATERIAL IS NULL")
      Else
      
       conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "MATERIAL" + """" + " = '0' WHERE " + """" + "MATERIAL" + """" + " IS NULL")
      End If
   
   'WHERE WATERLINES.id_type in(1,3,0,2,12) and WATERLINES.SECTOR IN (21)
   'str = "SELECT * FROM WATERLINES "
   
'   If Tipo <> "" Or setor <> "" Then
'      str = str & "WHERE "
'      If Tipo <> "" Then
'         str = str & "id_type in(" & Tipo & ") "
'      End If
'      If setor <> "" And Tipo <> "" Then
'         str = str & "and "
'      End If
'      If setor <> "" Then
'         str = str & "SECTOR IN (" & setor & ")"
'      End If
'   End If
   'MsgBox str
   
   If Provider = 1 Then

      str = "SELECT * FROM WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID "
      str = str & "WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & usuario & "' AND TIPO = 1)"

   ElseIf Provider = 2 Then
      
      str = "SELECT * FROM WATERLINES WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID "
      str = str & "WHERE EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE WATERLINES.LINE_ID = P.OBJECT_ID_ AND P.USUARIO = '" & usuario & "' AND P.TIPO = 1)"

   End If

    If conn.Provider = "PostgreSQL.1" Then
     str = "SELECT * FROM " + """" + "WATERLINES" + """" + " INNER JOIN " + """" + "X_MATERIAL" + """" + " ON " + """" + "WATERLINES" + """" + "." + """" + "MATERIAL" + """" + " = " + """" + "X_MATERIAL" + """" + "." + """" + "MATERIALID" + """" + " "
      str = str & "WHERE " + """" + "WATERLINES" + """" + "." + """" + "OBJECT_ID_" + """" + " IN (SELECT " + """" + "OBJECT_ID_" + """" + " FROM " + """" + "POLIGONO_SELECAO" + """" + " WHERE " + """" + "USUARIO" + """" + " = '" & usuario & "' AND " + """" + "TIPO" + """" + " = '1')"

    End If
   If conn.Provider <> "PostgreSQL.1" Then
   strtot = Replace(str, "SELECT *", "SELECT COUNT(*)")
      Else
      
       strtot = Replace(str, "SELECT *", "SELECT COUNT(*)")
      End If
   Set Rs = New ADODB.Recordset
      
       If conn.Provider <> "PostgreSQL.1" Then
   Rs.Open strtot, conn, adOpenDynamic, adLockReadOnly
   
   Else
     Rs.Open strtot, conn, adOpenDynamic, adLockOptimistic
     
   
   End If
   Me.ProgressBar1.Value = 1
   
   If Rs(0).Value > 0 Then
      Me.ProgressBar1.Max = Rs(0).Value
   Else
      MsgBox "Não há dados selecionados para exportar.", vbInformation, ""
      Exit Function
   End If
   
   Rs.Close
   Set Rs = Nothing
   
   Set Rs = New ADODB.Recordset
   Rs.Open str, conn, adOpenDynamic, adLockReadOnly
   
   If Rs.EOF = False Then
      ExportaEPANet Rs, conn
   Else
      MsgBox "Não há informações selecionadas para exportar.", vbInformation, ""
   End If


Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #2
        Open App.Path & "\LogErroExportEPANET.txt" For Append As #2
        Print #2, Now & "  - Private Sub cmdConfirmar_Click() - Linha: " & intLinhaCod & " - " & Err.Number & " - " & Err.Description
        Close #2
        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo LogErroExportEPANET.txt com informações desta ocorrencia.", vbInformation
    End If


End Function
Private Sub Timer1_Timer()

'   If txtTimer.Text = "" Then
      
      
      
      MousePointer = vbHourglass
      INICIAR
      
      MousePointer = vbDefault
      
      Timer1.Enabled = False
      End
      
'   Else
'      If IsDate(Me.txtTimer.Text) Then
'
'         If CDate(txtTimer.Text) < Format(Now, "HH:MM:SS") Then
'            INICIAR
'            Timer1.Enabled = False
'         End If
'
'      Else
'         MsgBox "Horário inválido"
'         Timer1.Enabled = False
'      End If
'   End If

End Sub




Private Sub cmdConfirmar_Click()
   Timer1.Enabled = True
   Me.ProgressBar1.Visible = True
   Me.cmdConfirmar.Enabled = False
   
End Sub


Private Sub cmdPath_Click()
   cdl.ShowSave
   cdl.FileName = txtArquivo.Text
   txtArquivo.Text = cdl.FileName
End Sub


Private Sub Command1_Click()

If MsgBox("Deseja aplicar fórmula Material x Rugosidade?", vbYesNo + vbQuestion, "Confirmar Ação") = vbYes Then
   FrmEPANET.MousePointer = vbHourglass
   
   

     If conn.Provider <> "PostgreSQL.1" Then
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 111 WHERE MATERIAL = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 1 AND ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 120 WHERE MATERIAL = 2 AND ROUGHNESS = 0 ")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 110 WHERE MATERIAL = 3 AND ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 105 WHERE MATERIAL = 4 AND ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 90 WHERE MATERIAL = 5 AND ROUGHNESS = 0")
   conn.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 6 AND ROUGHNESS = 0")
   Else
     conn.Execute ("UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE" + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '111' WHERE " + """" + "MATERIAL" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '1' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '120' WHERE " + """" + "MATERIAL" + """" + " = '2' AND " + """" + "ROUGHNESS" + """" + " = '0' ")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '110' WHERE " + """" + "MATERIAL" + """" + " = '3' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '105' WHERE " + """" + "MATERIAL" + """" + " = '4' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '90' WHERE " + """" + "MATERIAL" + """" + " = '5' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   conn.Execute ("UPDATE " + """" + "WATERLINES" + """" + " SET " + """" + "ROUGHNESS" + """" + " = '130' WHERE " + """" + "MATERIAL" + """" + " = '6' AND " + """" + "ROUGHNESS" + """" + " = '0'")
   
   
   End If
   
   
   
   
   
   FrmEPANET.MousePointer = vbDefault
   MsgBox "Fórmula aplicada com sucesso!", vbInformation, ""
End If


End Sub

