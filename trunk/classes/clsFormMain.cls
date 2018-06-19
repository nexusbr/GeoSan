VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
''local variable(s) to hold property value(s)
'Private mvarFrm As FrmMain 'local copy
'Private mbMoving        As Boolean
'Private mstrprevfunc    As String
'Private msngStartX      As Single
'Const sglSplitLimit = 0
'
'Public Property Set Frm(ByVal vData As FrmMain)
'    Set mvarFrm = vData
'End Property
'
'Public Property Get Frm() As FrmMain
'    Set Frm = mvarFrm
'End Property
'
'Public Sub Resize()
'   On Error GoTo Resize_error
'   With Frm
'      If .pctSfondo.Height > 50 Then
'         .StatusBar1.Panels.Item(2).Width = .Width - .StatusBar1.Panels.Item(1).Width
'         .TabStrip1.Height = .pctSfondo.Height - 50
'         .picSplitter.Height = .pctSfondo.Height
'         .imgSplitter.Height = .pctSfondo.Height
'         .Manager1.Width = .pctSfondo.Width
'         .Manager1.Width = .pctSfondo.Height + 1000
'         .Manager1.Resize .pctSfondo.Width - 300, .pctSfondo.Height
'         .ViewManager1.Width = .pctSfondo.Width - 1000
'         .ViewManager1.Height = .pctSfondo.Height - 50
'      End If
'   End With
'   On Error GoTo 0
'   Exit Sub
'Resize_error:
'
'End Sub
'
'Sub SizeControls(x As Single)
'   With Frm
'      .pctSfondo.Width = .pctSfondo.Width + msngStartX - x
'      .pctSfondo.Refresh
'      .TabStrip1.Width = .pctSfondo.Width
'   End With
'End Sub
'
'Public Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   msngStartX = x
'
'   With Frm
'       .picSplitter.Move .imgSplitter.Left, .imgSplitter.Top, .imgSplitter.Width \ 2, .imgSplitter.Height - 20
'   End With
'   Frm.picSplitter.Visible = True
'   Frm.picSplitter.Height = Frm.pctSfondo.Height
'
'   mbMoving = True
'End Sub
'
'Public Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim sglPos As Single
'    With Frm
'      If mbMoving Then
'          sglPos = x + .imgSplitter.Left
'          If sglPos < sglSplitLimit Then
'              .picSplitter.Left = sglSplitLimit
'          ElseIf sglPos > .Width - sglSplitLimit Then
'              .picSplitter.Left = .Width - sglSplitLimit
'          Else
'              .picSplitter.Left = sglPos
'          End If
'      End If
'     End With
'End Sub
'
'Public Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Call SizeControls(x)
'    Frm.picSplitter.Visible = False
'    mbMoving = False
'End Sub
