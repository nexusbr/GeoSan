VERSION 5.00
Begin VB.Form frmEpanet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   4650
      TabIndex        =   0
      Top             =   3135
      Width           =   1260
   End
End
Attribute VB_Name = "frmEpanet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

INICIAR

End Sub

Private Function INICIAR()
On Error GoTo Trata_Erro
   

'   MousePointer = vbHourglass
'   Conn.execute ("UPDATE WATERLINES SET ROUGHNESS = 0")
'   Conn.execute ("UPDATE WATERLINES SET ROUGHNESS = 111 WHERE MATERIAL = 0")
'   Conn.execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 1 AND ROUGHNESS = 0")
'   Conn.execute ("UPDATE WATERLINES SET ROUGHNESS = 120 WHERE MATERIAL = 2 AND ROUGHNESS = 0 ")
'   Conn.execute ("UPDATE WATERLINES SET ROUGHNESS = 110 WHERE MATERIAL = 3 AND ROUGHNESS = 0")
'   Conn.execute ("UPDATE WATERLINES SET ROUGHNESS = 105 WHERE MATERIAL = 4 AND ROUGHNESS = 0")
'   Conn.execute ("UPDATE WATERLINES SET ROUGHNESS = 90 WHERE MATERIAL = 5 AND ROUGHNESS = 0")
'   Conn.execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 6 AND ROUGHNESS = 0")
'   frmEpanet.MousePointer = vbDefault
   
   
   
   Dim Rs As ADODB.Recordset, str As String, Tipo  As String, setor As String, strtot As String
   

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
   
   conn.execute ("UPDATE WATERLINES SET MATERIAL = 0 WHERE MATERIAL IS NULL")
   
   str = "SELECT * FROM WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID "
   
   
   
   
   'WHERE WATERLINES.id_type in(1,3,0,2,12) and WATERLINES.SECTOR IN (21)
   'str = "SELECT * FROM WATERLINES "
   
   If Tipo <> "" Or setor <> "" Then
      str = str & "WHERE "
      If Tipo <> "" Then
         str = str & "id_type in(" & Tipo & ") "
      End If
      If setor <> "" And Tipo <> "" Then
         str = str & "and "
      End If
      If setor <> "" Then
         str = str & "SECTOR IN (" & setor & ")"
      End If
   End If
   'MsgBox str
   strtot = Replace(str, "SELECT *", "SELECT COUNT(*)")
   
   Set Rs = conn.execute(strtot)
   'frmOdometro.ProgressBar1.value = 1
   
   'frmOdometro.ProgressBar1.Max = Rs(0).value
   'frmOdometro.Show
   
   DoEvents
   
'   Close #2
'   Open App.Path & "\LogErroExportEPANET.txt" For Append As #2
'   Print #2, str
'   Close #2
   
   Set Rs = conn.execute(str)
   If Rs.EOF = False Then

   
      ExportaEPANet Rs, conn
   Else
      MsgBox "Não foi possivel carregar a tabela WATERLINES", vbInformation, ""
   End If

   Unload frmOdometro
   
   'End

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #2
        Open App.path & "\LogErroExportEPANET.txt" For Append As #2
        Print #2, Now & "  - Private Sub cmdConfirmar_Click() - Linha: " & intLinhaCod & " - " & Err.Number & " - " & Err.Description
        Close #2
        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo LogErroExportEPANET.txt com informações desta ocorrencia.", vbInformation
    End If


End Function
