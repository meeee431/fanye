VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSearchVechile 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "搜索车辆"
   ClientHeight    =   465
   ClientLeft      =   4905
   ClientTop       =   6210
   ClientWidth     =   2850
   HelpContextID   =   7000120
   Icon            =   "frmSearchVechile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1695
   End
   Begin RTComctl3.CoolButton cmdSearch 
      Default         =   -1  'True
      Height          =   315
      Left            =   1830
      TabIndex        =   1
      Top             =   75
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "搜索"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSearchVechile.frx":000C
      PICN            =   "frmSearchVechile.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmSearchVechile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mnSearchIndex As Integer    '起始搜索号
Public mFormNum As Integer '判断哪个窗口调用

Private Sub cmdSearch_Click()
    '从起始搜索索引开始循环查找
    Dim i As Integer
 
     txtSearch.Text = Trim(txtSearch.Text)
  If mFormNum = 2 Then
    With frmQuerySheet.txtBusID
          For i = mnSearchIndex + 1 To .ListCount   '后部搜索
              If InStr(1, .List(i), txtSearch.Text, vbTextCompare) > 0 Then
                  GoTo FindIt
              End If
          Next i
          For i = 0 To mnSearchIndex           '前部搜索
              If InStr(1, .List(i), txtSearch.Text, vbTextCompare) > 0 Then
                  GoTo FindIt
              End If
          Next i
      End With
    Else
         With frmCarryLuggage.cboCarryVehicle
            For i = mnSearchIndex + 1 To .ListCount   '后部搜索
                If InStr(1, .List(i), txtSearch.Text, vbTextCompare) > 0 Then
                    GoTo FindIt
                End If
            Next i
            For i = 0 To mnSearchIndex           '前部搜索
                If InStr(1, .List(i), txtSearch.Text, vbTextCompare) > 0 Then
                    GoTo FindIt
                End If
            Next i
        End With
    End If
    
    Exit Sub
FindIt: '找到了,定位
    Unload Me
   If mFormNum = 0 Then
'    frmCarryLuggage.cboCarryVehicle.SetFocus
    frmCarryLuggage.cboCarryVehicle.ListIndex = i
   ElseIf mFormNum = 1 Then
    frmUpdateSheet.txtVehicle.Text = frmCarryLuggage.cboCarryVehicle.List(i)
   ElseIf mFormNum = 2 Then
    frmQuerySheet.txtBusID.Text = frmQuerySheet.txtBusID.List(i)
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    If KeyAscii = vbKeyF1 Then
        DisplayHelp Me
    End If
End Sub

Private Sub Form_Load()
  If mFormNum = 0 Then
    With frmCarryLuggage
        Top = .cboCarryVehicle.Top + 1600
        Left = Screen.Width / 2.5
    End With
  ElseIf mFormNum = 1 Then
    With frmUpdateSheet
        Top = .txtVehicle.Top + 2800
        Left = .txtVehicle.Left + 2000
    End With
  ElseIf mFormNum = 2 Then
    With frmQuerySheet
        Top = .txtBusID.Top + 800
        Left = .txtBusID.Left
    End With
  End If
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch.Text)
End Sub

Public Property Get StartSearchIndex() As Integer
    StartSearchIndex = mnSearchIndex
End Property

Public Property Let StartSearchIndex(ByVal pnNewIndex As Integer)
    mnSearchIndex = pnNewIndex
End Property
