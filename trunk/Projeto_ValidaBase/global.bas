Attribute VB_Name = "Module1"
'
'Modulo - string que contém em arquivo VB o erro ocorreu
'EVENTO - string que contém em que rotina o erro ocorreu
'ErrDescr - string com a descrição do erro ocorrido
'ExibeMensagem - se é para exibir ou não uma mensagem para o usuário
'linha - número da linha em que o erro ocorreu
'
Public Function PrintErro(ByVal Modulo As String, ByVal EVENTO As String, ByVal ErrNum As String, ByVal ErrDescr As String, ByVal ExibeMensagem As Boolean, Optional ByVal linha As Integer = 0)
      Close #1 'FECHA O ARQUIVO DE LOG
      Open App.Path & "\Controles\ValidaBaseLog.txt" For Append As #1
      Print #1, "DATA"; Tab(16); Now
      Print #1, "USUÁRIO"; Tab(16); strUser
      Print #1, "VERSÃO"; Tab(16); Versao_Geo
      Print #1, "MÓDULO"; Tab(16); Modulo
      Print #1, "EVENTO"; Tab(16); EVENTO
      Print #1, "LINHA"; Tab(16); CStr(linha)
      Print #1, "MOTIVO"; Tab(16); ErrNum
      Print #1, "DESCRIÇÃO"; Tab(16); ErrDescr
      Print #1, ""
      Print #1, "-----------------------------------------------------------------------------------------------------"
      Print #1, ""
      Close #1 'FECHA O ARQUIVO
      'SE O PARÂMETRO ExibeMensagem = True , EXIBE MENSAGEM PARA O USUÁRIO
      If ExibeMensagem = True Then
         MsgBox "A operação não pode ser completada, consulte o arquivo " & App.Path & "\Controles\ValidaBaseLog.txt" & " para maiores detalhes.", vbInformation
      End If
End Function
