VERSION 5.00
Begin VB.Form frmApoio 
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox TeDatabase2 
      Height          =   480
      Left            =   1335
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   1095
      Width           =   1200
   End
   Begin VB.PictureBox TeDatabase1 
      Height          =   480
      Left            =   1005
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   480
      Width           =   1200
   End
End
Attribute VB_Name = "frmApoio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnDataBaseConectado As Boolean
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

Public Function ATUALIZA_TRECHOS_RAMAIS_AGUA()

On Error GoTo Trata_Erro

   Dim X_LINHA As Double
   Dim Y_LINHA As Double
   Dim retorno As Long
   Dim WTC As ADODB.Recordset
   Dim rsPoligono As New ADODB.Recordset
   Dim Fator As Double
   Dim lngContaReloc As Long
   Dim strNaoLocalizados As String

   If blnDataBaseConectado = False Then
      'EVITA CAUSAR UM ERRO DE RECONEXÃO CASO A TEDATABASE JA ESTEJA CONECTADA
      
      TeDatabase1.Provider = TpConexao
      TeDatabase1.Connection = Conn
      TeDatabase2.Provider = TpConexao
      TeDatabase2.Connection = Conn
      blnDataBaseConectado = True
   
   End If
   

   tb_linhas_ramais = TeDatabase1.getRepresentationTableName("RAMAIS_AGUA", tpLINES)
   If tb_linhas_ramais = "" Then
      MsgBox "NÃO FOI POSSIVEL LOCALIZAR A TABELA DE GEOMETRIA RAMAIS DE AGUA"
      Exit Function
   End If

   If TpConexao = 1 Then
   
      strsql = "SELECT " & tb_linhas_ramais & ".OBJECT_ID,RA.OBJECT_ID_TRECHO FROM " & tb_linhas_ramais & " " & tb_linhas_ramais & " INNER JOIN RAMAIS_AGUA RA ON " & tb_linhas_ramais & ".OBJECT_ID = RA.OBJECT_ID_ WHERE RA.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2)"
   
   ElseIf TpConexao = 2 Then
      
      strsql = "SELECT " & tb_linhas_ramais & ".OBJECT_ID,RA.OBJECT_ID_TRECHO FROM " & tb_linhas_ramais & " " & tb_linhas_ramais & " INNER JOIN RAMAIS_AGUA RA ON " & tb_linhas_ramais & ".OBJECT_ID = RA.OBJECT_ID_ WHERE EXISTS (SELECT 1 FROM POLIGONO_SELECAO P WHERE P.OBJECT_ID_ = RA.OBJECT_ID_ AND USUARIO = '" & strUser & "' AND TIPO = '2')"
   
    Else
    
a = tb_linhas_ramais
b = """OBJECT_ID"""
c = """OBJECT_ID_TRECHO"""
d = """RAMAIS_AGUA"""
e = """POLIGONO_SELECAO"""
f = """USUARIO"""
g = """tipo"""
i = "OBJECT_ID"
h = a + "." + i


    strsql = "SELECT ""+h+""," + d + "." + c + " FROM ""+ a+a+ "" INNER JOIN " + d + " ON "" h+"" = " + d + "." + b + " WHERE " + d + "." + b + "(SELECT " + d + " FROM " + e + " WHERE " + f + " = '" & strUser & "' AND " + g + " = '2')"
   
   
   
   End If
   
   'Imprima CStr(strsql)
   Set WTC = New ADODB.Recordset
   WTC.Open strsql, Conn, adOpenKeyset, adLockOptimistic

   retorno = 0
   If WTC.EOF = False Then

      TeDatabase2.setCurrentLayer "RAMAIS_AGUA"
      TeDatabase1.setCurrentLayer "WATERLINES"

      Do While Not WTC.EOF = True


         retorno = TeDatabase2.getPointOfLine(0, WTC!object_id, 0, X_LINHA, Y_LINHA) 'RECUPERA EM X E Y A EXTREMIDADE DE UMA LINHA DE RAMAL

         If retorno = 1 Then

            qtd = TeDatabase1.locateGeometry(X_LINHA, Y_LINHA, tpLINES, 0.05) 'PROCURA SE HÁ REDE DE AGUA A NO MAXIMO 5 CENTÍMETROS DE DISTÂNCIA

            If qtd = 1 Then ' CASO 1, HÁ 1 REDE PASSANDO NA EXTREMIDADE DO RAMAL
  If TpConexao = 1 Then
               strsql = "UPDATE RAMAIS_AGUA SET OBJECT_ID_TRECHO = " & TeDatabase1.objectIds(0) & " WHERE OBJECT_ID_ = '" & WTC!object_id & "'"
               ElseIf TpConexao = 4 Then
               a = tb_linhas_ramais
b = """OBJECT_ID"""
c = """OBJECT_ID_TRECHO"""
d = """RAMAIS_AGUA"""
e = """POLIGONO_SELECAO"""
f = """USUARIO"""
g = """tipo"""
i = "OBJECT_ID"
h = a + "." + i
                strsql = "UPDATE " + d + " SET " + c + " = " & TeDatabase1.objectIds(0) & " WHERE " + b + " = '" & WTC!object_id & "'"
             
               End If
               
               Conn.execute (strsql)

               lngContaReloc = lngContaReloc + 1

            ElseIf qtd = 0 Then
               Fator = 0.2
               Do While Not qtd = 1 And Fator < 5 ' executa um loop, aumentando a faixa de precisão, até que encontre 1 rede de agua a no máximo 5 metros

                  qtd = TeDatabase1.locateGeometry(X_LINHA, Y_LINHA, tpLINES, Fator)

                  If qtd = 1 Then


                   If TpConexao = 1 Then
            strsql = "UPDATE RAMAIS_AGUA SET OBJECT_ID_TRECHO = " & TeDatabase1.objectIds(0) & " WHERE OBJECT_ID_ = '" & WTC!object_id & "'"
                   ElseIf TpConexao = 4 Then
a = tb_linhas_ramais
b = """OBJECT_ID"""
c = """OBJECT_ID_TRECHO"""
d = """RAMAIS_AGUA"""
e = """POLIGONO_SELECAO"""
f = """USUARIO"""
g = """tipo"""
i = "OBJECT_ID"
h = a + "." + i
            strsql = "UPDATE " + d + " SET " + c + " = " & TeDatabase1.objectIds(0) & " WHERE " + b + " = '" & WTC!object_id & "'"
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

   WTC.Close

   If lngContaReloc > 0 Then
      MsgBox "Foi relocalizada a rede para " & lngContaReloc & " ramais.", vbInformation, ""
   End If
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   ElseIf Err.Number = 52 Then
      'Open App.path & "\" & strBanco & "_CORRETOR_BASE.TXT" For Append As #6 ' ABRE O ARQUIVO TEXTO PARA LOG
      Err.Clear
      Resume
   ElseIf Err.Number = 55 Then
      Err.Clear
      Resume Next
   Else
      MousePointer = vbDefault
      Close #6
      MsgBox Err.Number & " " & Err.Description

   End If
End Function
