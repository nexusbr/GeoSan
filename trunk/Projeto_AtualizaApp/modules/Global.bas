Attribute VB_Name = "Global"
'Referências adicionadas:
'Microsoft Scripting Runtime - srcrun.dll - para saber a versão de uma aplicação e copiar arquivos
'Microsoft CDO for Windows 2000 Library - cdosys.dll
'Componentes adicionados (controles):
'Microsoft Internet Transfer Control 6.0 - msinet.ocx - para fazer download de arquivos
'Microsoft Windows Comom Controls 6.0 (SP6) - mscomctl.ocx
'Microsoft Winsock Control 6.0 - mswsock.dll
'
' Arquivo GeoSan.ini
'[ATUALIZACAO]                  - informações para o GeoSan atualizar-se automaticamente
'WEB = NAO                      - indica que vai buscar as atualizações em um diretório local
'DIRETORIO=\download\GeoSan     - nome do sub-diretório onde irá buscar as atualizações, se for web a barra e normal ficando: /download/GeoSan
'URL=c:\tempFtp                 - nome do diretório (caso esteja loca a atualização) ou endereço web, p. ex.: http://www.nexusbr.com
'proxyPorta = NULO              - número da porta em que será buscada a atualização. Para buscar no site da NEXUS é porta 80
'proxy = NULO                   - endereço do proxy da interno da empresa
'DIRETORIOLOCAL=c:\tempApp      - nome do diretório completo para onde serão baixadas as atualizações
'USUARIO=nexus                  - nome do usuário para logar no proxy interno da empresa
'SENHA=senha                    - senha para logar no proxy interno da empresa
'

'Definição de variáveis globais
'
'
'
Public b() As Byte

Public ErroUsuario As New CPrintErro            'Para gerenciar os erros que por ventura ocorram
Public conf As New CArquivoIni                  'Para ler e escrever as configurações de trabalho do arquivo GEOSAN.INI
Public versao As CGetVersion                    'gestão das versões de software que deverão ser atualizadas
Public mensagem As String                       'mensagem do que está realizando para o usuário
Public Email As New CEmail                      'Classe responsável pelo envio de emails

Public Sub Main()
    MsgBox "test"
End Sub
