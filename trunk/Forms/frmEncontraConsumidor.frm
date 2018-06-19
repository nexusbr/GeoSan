VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TeComDatabase.dll"
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Object = "{EE78E37B-39BE-42FA-80B7-E525529739F7}#1.0#0"; "TeComViewDatabase.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEncontraConsumidor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localizar Consumidor"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Parte do texto"
      Height          =   720
      Left            =   285
      TabIndex        =   4
      Top             =   645
      Width           =   3750
      Begin VB.OptionButton optInicio 
         Caption         =   "Início"
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   330
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optFim 
         Caption         =   "Fim"
         Height          =   315
         Left            =   1185
         TabIndex        =   6
         Top             =   300
         Width           =   840
      End
      Begin VB.OptionButton optQQRParte 
         Caption         =   "Qualquer parte"
         Height          =   300
         Left            =   2055
         TabIndex        =   5
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.OptionButton optNomeCliente 
      Caption         =   "Nome do Cliente"
      Height          =   270
      Left            =   2265
      TabIndex        =   3
      Top             =   270
      Width           =   1665
   End
   Begin VB.OptionButton optNroLigacao 
      Caption         =   "Número da Ligação"
      Height          =   270
      Left            =   270
      TabIndex        =   2
      Top             =   270
      Value           =   -1  'True
      Width           =   1890
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Pesquisar"
      Default         =   -1  'True
      Height          =   360
      Left            =   4245
      TabIndex        =   1
      Top             =   1530
      Width           =   1080
   End
   Begin VB.TextBox txtBusca 
      Height          =   300
      Left            =   255
      TabIndex        =   0
      Top             =   1545
      Width           =   3795
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1995
      Left            =   210
      TabIndex        =   8
      Top             =   2025
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Localizado"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Eixo X"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Eixo Y"
         Object.Width           =   2540
      EndProperty
   End
   Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
      Left            =   3840
      OleObjectBlob   =   "frmEncontraConsumidor.frx":0000
      Top             =   0
   End
   Begin TeComViewDatabaseLibCtl.TeViewDatabase TeViewDatabase1 
      Left            =   4440
      OleObjectBlob   =   "frmEncontraConsumidor.frx":0024
      Top             =   120
   End
   Begin TECOMDATABASELibCtl.TeDatabase TeDatabase 
      Left            =   4320
      OleObjectBlob   =   "frmEncontraConsumidor.frx":0048
      Top             =   720
   End
End
Attribute VB_Name = "frmEncontraConsumidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Localiza os consumidores
'
'
'
Private Sub cmdLocalizar_Click()
    On Error GoTo Trata_Erro
    Dim aa As String
    Dim ab As String
    Dim ac As String
    Dim ad As String
    Dim ae As String
    Dim af As String
    Dim ag As String
    Dim ah As String
    Dim ai As String
    Dim aj As String
    Dim mPROVEDOR As String
    Dim mSERVIDOR As String
    Dim mPORTA As String
    Dim mBANCO As String
    Dim mUSUARIO As String
    Dim Senha As String
    Dim decriptada As String
    Dim tbPoints As String
    Dim str As String
    Dim rs As New ADODB.Recordset

    If (frmCanvas.TipoConexao = 4) Then
        If (frmCanvas.POST <> 10) Then
            mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
            mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
            mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
            mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
            Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
            frmCanvas.FunDecripta (Senha)
            decriptada = frmCanvas.Senha
            TeAcXConnection1.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA
            frmCanvas.POST2 (10)
            TeDatabase.Provider = frmCanvas.TipoConexao
            TeDatabase.connection = TeAcXConnection1.objectConnection_
        End If
    Else
        TeDatabase.Provider = frmCanvas.TipoConexao
        TeDatabase.connection = Conn
    End If
    'RECUPERA A TABELA QUE POSSUI OS PONTOS REFERENTES A RAMAIS AGUA
    tbPoints = TeDatabase.getRepresentationTableName("RAMAIS_AGUA", tpPOINTS)
    If Trim(Me.txtBusca.Text) = "" Then
        MsgBox "Informe o valor que deseja procurar", vbInformation, ""
        Me.txtBusca.SetFocus
        Exit Sub
    End If
    MousePointer = vbHourglass
    If Me.optNroLigacao.value = True Then
        If frmCanvas.TipoConexao <> 4 Then
            If Me.optInicio.value = True Then
                str = "SELECT RAL.OBJECT_ID_ AS " + """" + "ID" + """" + ",RAL.NRO_LIGACAO AS " + """" + "BUSCA" + """" + ",PT.X,PT.Y FROM "
                str = str & "RAMAIS_AGUA_LIGACAO RAL JOIN " & tbPoints & " PT ON PT.OBJECT_ID = RAL.OBJECT_ID_ "
                str = str & "WHERE RAL.NRO_LIGACAO like '" & Me.txtBusca.Text & "%'"
            ElseIf Me.optQQRParte.value = True Then
                str = "SELECT RAL.OBJECT_ID_ AS " + """" + "ID" + """" + ",RAL.NRO_LIGACAO AS " + """" + "BUSCA" + """" + ",PT.X,PT.Y FROM "
                str = str & "RAMAIS_AGUA_LIGACAO RAL JOIN " & tbPoints & " PT ON PT.OBJECT_ID = RAL.OBJECT_ID_ "
                str = str & "WHERE RAL.NRO_LIGACAO like '" & "%" & Me.txtBusca.Text & "%'"
            ElseIf Me.optFim.value = True Then
                str = "SELECT RAL.OBJECT_ID_ AS " + """" + "ID" + """" + ",RAL.NRO_LIGACAO AS " + """" + "BUSCA" + """" + ",PT.X,PT.Y FROM "
                str = str & "RAMAIS_AGUA_LIGACAO RAL JOIN " & tbPoints & " PT ON PT.OBJECT_ID = RAL.OBJECT_ID_ "
                str = str & "WHERE RAL.NRO_LIGACAO like '" & "%" & Me.txtBusca.Text & "%'"
            End If
        Else
            aa = "RAMAIS_AGUA_LIGACAO"
            ab = "OBJECT_ID_"
            ac = "NRO_LIGACAO"
            ad = LCase(tbPoints)
            ae = "x"
            af = "y"
            ag = "object_id"
            ah = "NXGS_V_LIG_COMERCIAL"
            ai = "CONSUMIDOR"
            If Me.optInicio.value = True Then
                str = "SELECT " + """" + aa + """" + "." + """" + ab + """" + " AS " + """" + "ID" + """" + "," + """" + aa + """" + "." + """" + ac + """" + " AS " + """" + "BUSCA" + """" + "," + """" + ad + """" + "." + """" + ae + """" + "," + """" + ad + """" + "." + """" + af + """" + " FROM "
                str = str & "" + """" + aa + """" + " JOIN " + """" + ad + """" + "  ON " + """" + ad + """" + "." + """" + ag + """" + "=" + """" + aa + """" + "." + """" + ab + """" + ""
                str = str & "WHERE " + """" + aa + """" + "." + """" + ac + """" + " like '" & Me.txtBusca.Text & "%'"
                '  MsgBox "ARQUIVO DEBUG SALVO"
                ' WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
            ElseIf Me.optQQRParte.value = True Then
                str = "SELECT " + """" + aa + """" + "." + """" + ab + """" + " AS " + """" + "ID" + """" + "," + """" + aa + """" + "." + """" + ac + """" + " AS " + """" + "BUSCA" + """" + "," + """" + ad + """" + "." + """" + ae + """" + "," + """" + ad + """" + "." + """" + af + """" + " FROM "
                str = str & "" + """" + aa + """" + " JOIN " + """" + ad + """" + "  ON " + """" + ad + """" + "." + """" + ag + """" + "=" + """" + aa + """" + "." + """" + ab + """" + ""
                str = str & "WHERE " + """" + aa + """" + "." + """" + ac + """" + " like '%" & Me.txtBusca.Text & "%'"
            ElseIf Me.optFim.value = True Then
                str = "SELECT " + """" + aa + """" + "." + """" + ab + """" + " AS " + """" + "ID" + """" + "," + """" + aa + """" + "." + """" + ac + """" + " AS " + """" + "BUSCA" + """" + "," + """" + ad + """" + "." + """" + ae + """" + "," + """" + ad + """" + "." + """" + af + """" + " FROM "
                str = str & "" + """" + aa + """" + " JOIN " + """" + ad + """" + "  ON " + """" + ad + """" + "." + """" + ag + """" + "=" + """" + aa + """" + "." + """" + ab + """" + ""
                str = str & "WHERE " + """" + aa + """" + "." + """" + ac + """" + " like '%" & Me.txtBusca.Text & "'"
            End If
        End If
    End If
    If Me.optNomeCliente.value = True Then
        If frmCanvas.TipoConexao <> 4 Then
            If Me.optInicio.value = True Then
                str = "SELECT RAL.OBJECT_ID_ AS " + """" + "ID" + """" + ",COM.CONSUMIDOR AS " + """" + "BUSCA" + """" + ",PT.X,PT.Y "
                str = str & "FROM NXGS_V_LIG_COMERCIAL COM "
                str = str & "JOIN RAMAIS_AGUA_LIGACAO RAL ON RAL.NRO_LIGACAO = COM.NRO_LIGACAO "
                str = str & "JOIN " & tbPoints & " PT ON RAL.OBJECT_ID_ = PT.OBJECT_ID "
                str = str & "WHERE COM.CONSUMIDOR LIKE '" & Me.txtBusca.Text & "%'"
            ElseIf Me.optQQRParte.value = True Then
                str = "SELECT RAL.OBJECT_ID_ AS " + """" + "ID" + """" + ",COM.CONSUMIDOR AS " + """" + "BUSCA" + """" + ",PT.X,PT.Y "
                str = str & "FROM NXGS_V_LIG_COMERCIAL COM "
                str = str & "JOIN RAMAIS_AGUA_LIGACAO RAL ON RAL.NRO_LIGACAO = COM.NRO_LIGACAO "
                str = str & "JOIN " & tbPoints & " PT ON RAL.OBJECT_ID_ = PT.OBJECT_ID "
                str = str & "WHERE COM.CONSUMIDOR LIKE '%" & Me.txtBusca.Text & "%'"
            ElseIf Me.optFim.value = True Then
                str = "SELECT RAL.OBJECT_ID_ AS " + """" + "ID" + """" + ",COM.CONSUMIDOR AS " + """" + "BUSCA" + """" + ",PT.X,PT.Y "
                str = str & "FROM NXGS_V_LIG_COMERCIAL COM "
                str = str & "JOIN RAMAIS_AGUA_LIGACAO RAL ON RAL.NRO_LIGACAO = COM.NRO_LIGACAO "
                str = str & "JOIN " & tbPoints & " PT ON RAL.OBJECT_ID_ = PT.OBJECT_ID "
                str = str & "WHERE COM.CONSUMIDOR LIKE '%" & Me.txtBusca.Text & "'"
            End If
        Else
            aa = "RAMAIS_AGUA_LIGACAO"
            ab = "OBJECT_ID_"
            ac = "NRO_LIGACAO"
            ad = LCase(tbPoints)
            ae = "x"
            af = "y"
            ag = "object_id"
            ah = "NXGS_V_LIG_COMERCIAL"
            ai = "CONSUMIDOR"
            If Me.optInicio.value = True Then
                str = "SELECT " + """" + aa + """" + "." + """" + ab + """" + " AS " + """" + "ID" + """" + "," + """" + ah + """" + "." + """" + ai + """" + " AS " + """" + "BUSCA" + """" + "," + """" + ad + """" + "." + """" + ae + """" + "," + """" + ad + """" + "." + """" + af + """" + " FROM "
                str = str & "" + """" + ah + """" + ""
                str = str & "JOIN " + """" + aa + """" + " ON " + """" + aa + """" + "." + """" + ac + """" + " = " + """" + ah + """" + "." + """" + ac + """" + " "
                str = str & "JOIN " + """" + LCase(tbPoints) + """" + "  ON " + """" + aa + """" + "." + """" + ab + """" + " = " + """" + ad + """" + "." + """" + ag + """" + " "
                str = str & "WHERE " + """" + ah + """" + "." + """" + ai + """" + " LIKE '" & Me.txtBusca.Text & "%'"
            ElseIf Me.optQQRParte.value = True Then
                str = "SELECT " + """" + aa + """" + "." + """" + ab + """" + " AS " + """" + "ID" + """" + "," + """" + ah + """" + "." + """" + ai + """" + " AS " + """" + "BUSCA" + """" + "," + """" + ad + """" + "." + """" + ae + """" + "," + """" + ad + """" + "." + """" + af + """" + " FROM "
                str = str & " " + """" + ah + """" + ""
                str = str & "JOIN " + """" + aa + """" + " ON " + """" + aa + """" + "." + """" + ac + """" + " = " + """" + ah + """" + "." + """" + ac + """" + " "
                str = str & "JOIN " + """" + LCase(tbPoints) + """" + "  ON " + """" + aa + """" + "." + """" + ab + """" + " = " + """" + ad + """" + "." + """" + ag + """" + " "
                str = str & "WHERE " + """" + ah + """" + "." + """" + ai + """" + " LIKE '%" & Me.txtBusca.Text & "%'"
            ElseIf Me.optFim.value = True Then
                str = "SELECT " + """" + aa + """" + "." + """" + ab + """" + " AS " + """" + "ID" + """" + "," + """" + ah + """" + "." + """" + ai + """" + " AS " + """" + "BUSCA" + """" + "," + """" + ad + """" + "." + """" + ae + """" + "," + """" + ad + """" + "." + """" + af + """" + " FROM "
                str = str & " " + """" + ah + """" + ""
                str = str & "JOIN " + """" + aa + """" + " ON " + """" + aa + """" + "." + """" + ac + """" + " = " + """" + ah + """" + "." + """" + ac + """" + " "
                str = str & "JOIN " + """" + LCase(tbPoints) + """" + "  ON " + """" + aa + """" + "." + """" + ab + """" + " = " + """" + ad + """" + "." + """" + ag + """" + " "
                str = str & "WHERE " + """" + ah + """" + "." + """" + ai + """" + " LIKE '%" & Me.txtBusca.Text & "'"
            End If
        End If
    End If
    rs.Open str, Conn, adOpenDynamic, adLockOptimistic
    Me.Lista.ListItems.Clear
    If rs.EOF = False Then
        'CARREGA NO FORM TODAS AS LIGAÇÕES DISPONIVEIS COM BASE NO PRÉ FILTRO
        Do While Not rs.EOF
            Set itmx = Lista.ListItems.Add(, , rs.Fields("BUSCA").value)
            itmx.SubItems(1) = IIf(IsNull(rs.Fields("X").value), "", rs.Fields("X").value)
            itmx.SubItems(2) = IIf(IsNull(rs.Fields("Y").value), "", rs.Fields("Y").value)
            itmx.Tag = rs.Fields("ID").value
            rs.MoveNext
        Loop
    End If
    rs.Close
    MousePointer = vbDefault
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmEncontraConsumidor", "cmdLocalizar_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
    MousePointer = vbDefault
End Sub
' Usuário selecionou com dois cliques do mouse um consumidor e irá fazer o zoom no mesmo
'
'
'
Private Sub Lista_DblClick()
    On Error GoTo Trata_Erro
    Dim i As Long
    Dim X As Double, Y As Double
    Dim xmin As Double
    Dim ymin As Double
    Dim xmax As Double
    Dim ymax As Double
    Dim a As String
    Dim Object_id_ As String
    
    X = Lista.SelectedItem.ListSubItems(1)
    Y = Lista.SelectedItem.ListSubItems(2)
    blnLocalizandoConsumidor = True
    xWorld = X 'carrega as variáveis públicas
    yWorld = Y 'carrega as variáveis públicas
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmEncontraConsumidor", "Lista_DblClick", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
    MousePointer = vbDefault
End Sub

