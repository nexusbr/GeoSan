VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TeComDatabase.dll"
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Begin VB.Form frmAlteraNoPoligono 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Operações por Polígono"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Fechar"
      Height          =   405
      Left            =   2460
      TabIndex        =   2
      Top             =   2610
      Width           =   1095
   End
   Begin VB.CommandButton cmdAtivar 
      Caption         =   "Próximo >>"
      Height          =   405
      Left            =   3660
      TabIndex        =   1
      Top             =   2610
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione a operação desejada"
      Height          =   2235
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   4620
      Begin VB.OptionButton optLocTrechoRede 
         Caption         =   "Associa o ramal a rede mais próxima (até 5 m)"
         Height          =   435
         Left            =   165
         TabIndex        =   6
         ToolTipText     =   "Utilizada caso a rede tenha sido apagada e seja necessário ligar o ramal a outra nova rede"
         Top             =   1650
         Width           =   4170
      End
      Begin VB.OptionButton optRelatorios 
         Caption         =   "Relatórios em Arquivo Texto"
         Height          =   435
         Left            =   165
         TabIndex        =   5
         Top             =   1185
         Width           =   3450
      End
      Begin VB.OptionButton optAlteraConsumo 
         Caption         =   "Alterar Consumo de Ligações"
         Height          =   435
         Left            =   165
         TabIndex        =   4
         ToolTipText     =   "Força consumos em ligações, desconsiderando o consumo vindo do sistema comercial"
         Top             =   750
         Width           =   3450
      End
      Begin VB.OptionButton optExportEpanet 
         Caption         =   "Exportar para Epanet"
         Height          =   435
         Left            =   165
         TabIndex        =   3
         Top             =   345
         Value           =   -1  'True
         Width           =   3450
      End
      Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
         Left            =   3960
         OleObjectBlob   =   "frmAlteraNoPoligono.frx":0000
         Top             =   1680
      End
      Begin TECOMDATABASELibCtl.TeDatabase TeDatabase2 
         Left            =   3240
         OleObjectBlob   =   "frmAlteraNoPoligono.frx":0024
         Top             =   840
      End
      Begin TECOMDATABASELibCtl.TeDatabase TeDatabase1 
         Left            =   3360
         OleObjectBlob   =   "frmAlteraNoPoligono.frx":0048
         Top             =   240
      End
   End
End
Attribute VB_Name = "frmAlteraNoPoligono"
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
 Dim conexao As New ADODB.connection
  Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String

Dim strConn As String
Dim nStr As String
Dim usuario As String
Dim count1, count2 As Integer
'Subrotina acionada caso ele selecione o botão PRÓXIMO, para rodar o próximo passo com relação as tubulações, ramais e nós pré-selecionados
'
'
'
Private Sub cmdAtivar_Click()
    If Me.optExportEpanet.value = True Then
        'deseja exportar para o Epanet
        Shell glo.diretorioGeoSan + "\Exporte EPANet.exe", vbNormalFocus
    ElseIf Me.optAlteraConsumo.value = True Then
        'deseja alterar os consumos das ligações
        Me.Visible = False
        frmAlteraConsumoPorPoligono.Show 1
        ElseIf Me.optRelatorios.value = True Then
            'deseja emitir relatórios no formato texto
            Me.Visible = False
            frmRelatoriosAvancados.Show 1
            ElseIf Me.optLocTrechoRede.value = True Then
                'deseja localizar trechos de redes dos ramais selecionados
                If ATUALIZA_TRECHOS_RAMAIS_AGUA = True Then
                    MsgBox "Comando executado com sucesso!", vbInformation, ""
                End If
    End If
End Sub

Public Function ATUALIZA_TRECHOS_RAMAIS_AGUA() As Boolean

On Error GoTo Trata_Erro

   Dim X_LINHA As Double
   Dim Y_LINHA As Double
   Dim retorno As Long
   Dim WTC As ADODB.Recordset
   Dim rsPoligono As New ADODB.Recordset
   Dim Fator As Double
   Dim lngContaReloc As Long
   Dim strNaoLocalizados As String


   'TeDatabase1.Provider = frmCanvas.TipoConexao
   'TeDatabase1.Connection = frmCanvas.conexao
   
   
   'TeDatabase2.Provider = frmCanvas.TipoConexao
   'TeDatabase2.Connection = frmCanvas.conexao


   MousePointer = vbHourglass
   
   tb_linhas_ramais = ""
   'TB_LINHAS_RAMAIS = TeDatabase1.getRepresentationTableName("RAMAIS_AGUA", tpLINES)
   Dim rs As Recordset
   Set rs = New ADODB.Recordset
   
Dim sm As String
Dim sn As String
Dim so As String
Dim sp As String
Dim sq As String
Dim sr As String

sm = "geom_table"
sn = "te_representation"
so = "geom_type"
sp = "layer_id"
sq = "te_layer"
sr = "name"
         
   
   'retornar a tabela de geometria de linhas
   If frmCanvas.TipoConexao <> 4 Then
   
   rs.Open "SELECT GEOM_TABLE FROM TE_REPRESENTATION WHERE GEOM_TYPE = 2 AND LAYER_ID IN (SELECT LAYER_ID FROM TE_LAYER WHERE NAME = 'RAMAIS_AGUA')", Conn, adOpenDynamic, adLockReadOnly
   
   If rs.EOF = False Then
      tb_linhas_ramais = rs!GEOM_TABLE
        End If
        
      Else
      rs.Open "SELECT " + """" + sm + """" + " FROM " + """" + sn + """" + " WHERE " + """" + so + """" + " = '2' AND " + """" + sp + """" + " IN (SELECT " + """" + sp + """" + " FROM " + """" + sq + """" + " WHERE " + """" + sr + """" + " = 'RAMAIS_AGUA')", Conn, adOpenKeyset, adLockOptimistic
 
   If rs.EOF = False Then
     tb_linhas_ramais = rs!GEOM_TABLE
      End If
   End If
   rs.Close
   
      If tb_linhas_ramais = "" Then
      MsgBox "NÃO FOI POSSIVEL LOCALIZAR A TABELA DE GEOMETRIA RAMAIS DE AGUA"
      Exit Function
   End If
   
   
a = LCase(tb_linhas_ramais)
b = "OBJECT_ID_"
c = "OBJECT_ID_TRECHO"
d = "RAMAIS_AGUA"
e = "POLIGONO_SELECAO"
f = "USUARIO"
g = "TIPO"
h = "OBJECT_ID_"
i = a + "." + h
j = "OBJECT_ID_"
'k = a + "+""""+"."+"""" + j
   If frmCanvas.TipoConexao = 1 Then

      strsql = "SELECT " & tb_linhas_ramais & ".OBJECT_ID,RA.OBJECT_ID_TRECHO FROM " & tb_linhas_ramais & " " & tb_linhas_ramais & " INNER JOIN RAMAIS_AGUA RA ON " & tb_linhas_ramais & ".OBJECT_ID = RA.OBJECT_ID_ WHERE RA.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2)"

   ElseIf frmCanvas.TipoConexao = 2 Then

      strsql = "SELECT " & tb_linhas_ramais & ".OBJECT_ID,RA.OBJECT_ID_TRECHO FROM " & tb_linhas_ramais & " " & tb_linhas_ramais & " INNER JOIN RAMAIS_AGUA RA ON " & tb_linhas_ramais & ".OBJECT_ID = RA.OBJECT_ID_ WHERE EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE P.OBJECT_ID_ = RA.OBJECT_ID_ AND USUARIO = '" & strUser & "' AND TIPO = '2')"
ElseIf frmCanvas.TipoConexao = 4 Then

strsql = "SELECT " + """" + LCase(tb_linhas_ramais) + """" + "." + """" + "object_id" + """" + "," + """" + d + """" + "." + """" + c + """" + " FROM " + """" + a + """" + " INNER JOIN " + """" + d + """" + " ON " + """" + LCase(tb_linhas_ramais) + """" + "." + """" + "object_id" + """" + " = " + """" + d + """" + "." + """" + b + """" + " WHERE " + """" + d + """" + "." + """" + b + """" + " IN (SELECT " + """" + b + """" + " FROM " + """" + e + """" + " WHERE " + """" + f + """" + " = '" & strUser & "' AND " + """" + g + """" + " = '2')"
   'MsgBox strsql
  ' MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", "SELECT " + """" + tb_linhas_ramais + """" + "." + """" + "OBJECT_ID" + """" + "," + """" + d + """" + "." + """" + c + """" + " FROM " + """" + a + """" + " INNER JOIN " + """" + d + """" + " ON " + """" + tb_linhas_ramais + """" + "." + """" + "OBJECT_ID_" + """" + " = " + """" + d + """" + "." + """" + b + """" + " WHERE " + """" + d + """" + "." + """" + b + """" + " IN (SELECT " + """" + b + """" + " FROM " + """" + e + """" + " WHERE " + """" + f + """" + " = '" & strUser & "' AND " + """" + g + """" + " = '2')", App.path & "\DEBUG.INI"
   
   End If
   
   
   'Imprima CStr(strsql)
   Set WTC = New ADODB.Recordset
   WTC.Open strsql, Conn, adOpenKeyset, adLockOptimistic

   retorno = 0
   If WTC.EOF = False Then

mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
nStr = frmCanvas.FunDecripta(Senha)
decriptada = frmCanvas.Senha
  strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"
usuario = ReadINI("CONEXAO", "USER", App.path & "\CONTROLES\GEOSAN.ini")


If frmCanvas.TipoConexao <> 4 Then
If count1 <> 10 Then
 TeDatabase1.username = usuario
 TeDatabase1.Provider = frmCanvas.TipoConexao
 TeDatabase1.connection = Conn
 
 TeDatabase2.username = usuario
  TeDatabase2.Provider = typeconnection
 TeDatabase2.connection = Conn
 count1 = 10
End If

      TeDatabase2.setCurrentLayer "RAMAIS_AGUA"
      TeDatabase1.setCurrentLayer "WATERLINES"
      
      Else
      If count2 <> 10 Then
  TeAcXConnection1.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA

TeDatabase1.username = usuario
 TeDatabase1.Provider = frmCanvas.TipoConexao
 TeDatabase1.connection = TeAcXConnection1.objectConnection_
TeDatabase2.username = usuario
 TeDatabase2.Provider = frmCanvas.TipoConexao
 TeDatabase2.connection = TeAcXConnection1.objectConnection_

    conexao.Open strConn
     count2 = 10
      End If
       TeDatabase2.setCurrentLayer "RAMAIS_AGUA"
      TeDatabase1.setCurrentLayer "WATERLINES"
      End If

      Do While Not WTC.EOF = True


         retorno = TeDatabase2.getPointOfLine(0, WTC!object_id, 0, X_LINHA, Y_LINHA) 'RECUPERA EM X E Y A EXTREMIDADE DE UMA LINHA DE RAMAL

         If retorno = 1 Then

            qtd = TeDatabase1.locateGeometry(X_LINHA, Y_LINHA, tpLINES, 0.05) 'PROCURA SE HÁ REDE DE AGUA A NO MAXIMO 5 CENTÍMETROS DE DISTÂNCIA

            If qtd = 1 Then ' CASO 1, HÁ 1 REDE PASSANDO NA EXTREMIDADE DO RAMAL
a = "TB_LINHAS_RAMAIS"
b = "OBJECT_ID_"
c = "OBJECT_ID_TRECHO"
d = "RAMAIS_AGUA"
e = "POLIGONO_SELECAO"
f = "USUARIO"
g = "TIPO"
              If frmCanvas.TipoConexao <> 4 Then
               strsql = "UPDATE RAMAIS_AGUA SET OBJECT_ID_TRECHO = " & TeDatabase1.objectIds(0) & " WHERE OBJECT_ID_ = '" & WTC!object_id & "'"
               Else
               strsql = "UPDATE " + """" + d + """" + " SET " + """" + c + """" + " = '" & TeDatabase1.objectIds(0) & "' WHERE " + """" + b + """" + " = '" & WTC!object_id & "'"
               End If
               Conn.execute (strsql)

               lngContaReloc = lngContaReloc + 1

            ElseIf qtd = 0 Then
               Fator = 0.2
               Do While Not qtd = 1 And Fator < 5 ' executa um loop, aumentando a faixa de precisão, até que encontre 1 rede de agua a no máximo 5 metros

                  qtd = TeDatabase1.locateGeometry(X_LINHA, Y_LINHA, tpLINES, Fator)

                  If qtd = 1 Then
                      If frmCanvas.TipoConexao <> 4 Then
                     strsql = "UPDATE RAMAIS_AGUA SET OBJECT_ID_TRECHO = " & TeDatabase1.objectIds(0) & " WHERE OBJECT_ID_ = '" & WTC!object_id & "'"
                     Else
                      strsql = "UPDATE " + """" + d + """" + " SET " + """" + c + """" + " = '" & TeDatabase1.objectIds(0) & "' WHERE " + """" + b + """" + " = '" & WTC!object_id & "'"
                     End If
                     Conn.execute (strsql)

                     lngContaReloc = lngContaReloc + 1

                     Exit Do

                  End If

                  Fator = Fator + 0.1

               Loop
               If qtd <> 1 Then

                  lngContaNaoReloc = lngContaNaoReloc + 1

               End If
            Else

               lngContaNaoReloc = lngContaNaoReloc + 1

            End If
         End If

         WTC.MoveNext
      Loop
   End If

   MousePointer = vbDefault

   ATUALIZA_TRECHOS_RAMAIS_AGUA = True

   If lngContaReloc > 0 Then
      MsgBox "Foi relocalizada a rede para " & lngContaReloc & " ramais.", vbInformation, ""
   End If

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   ElseIf Err.Number = 52 Then
      Open App.path & "\" & strBanco & "_CORRETOR_BASE.TXT" For Append As #6 ' ABRE O ARQUIVO TEXTO PARA LOG
      Err.Clear
      Resume
   ElseIf Err.Number = 55 Then
      Err.Clear
      Resume Next
   Else
      MousePointer = vbDefault
      Close #6
      MsgBox Err.Number & " " & Err.Description
      ATUALIZA_TRECHOS_RAMAIS_AGUA = False
   End If
End Function

Private Sub cmdCancelar_Click()
   
   Unload Me

End Sub


