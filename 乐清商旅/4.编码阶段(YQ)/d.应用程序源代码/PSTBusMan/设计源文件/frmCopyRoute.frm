VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmCopyRoute 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "线路复制"
   ClientHeight    =   1875
   ClientLeft      =   3585
   ClientTop       =   3225
   ClientWidth     =   5430
   Icon            =   "frmCopyRoute.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin FText.asFlatTextBox txtEndStationID 
      Height          =   300
      Left            =   1800
      TabIndex        =   8
      Top             =   1395
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4185
      TabIndex        =   7
      Top             =   630
      Width           =   1110
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmCopyRoute.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCopy 
      Height          =   330
      Left            =   4185
      TabIndex        =   6
      Top             =   225
      Width           =   1110
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "复制(&O)"
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
      MICON           =   "frmCopyRoute.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtRouteName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   990
      Width           =   2175
   End
   Begin VB.TextBox txtNewRoute 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   2
      Top             =   585
      Width           =   2175
   End
   Begin FText.asFlatTextBox txtRoute 
      Height          =   300
      Left            =   1800
      TabIndex        =   9
      Top             =   195
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "新线路终点站(&E):"
      Height          =   180
      Left            =   270
      TabIndex        =   5
      Top             =   1455
      Width           =   1440
   End
   Begin VB.Label lblNewRouteName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "新线路名称(&N):"
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   1050
      Width           =   1260
   End
   Begin VB.Label lblNewRouteID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "新线路代码(&I):"
      Height          =   180
      Left            =   270
      TabIndex        =   1
      Top             =   645
      Width           =   1260
   End
   Begin VB.Label lblRouteID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原线路代码(&G):"
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmCopyRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_bIsParent As Boolean '是否父窗体直接调用
Public m_szOldRouteID As String
Private m_oRoute As New Route


Private Sub CmdCopy_Click()
    Dim nResult As VbMsgBoxResult
    On Error GoTo ErrorHandle
    nResult = MsgBox("是否复制线路[" & txtNewRoute.Text & "]", vbYesNo + vbQuestion, Me.Caption)
    If nResult = vbNo Then Exit Sub
    SetBusy
    
    m_oRoute.Identify Trim(ResolveDisplay(txtRoute.Text))
    m_oRoute.CloneRoute txtNewRoute.Text, Trim(txtRouteName.Text), Trim(ResolveDisplay(txtEndStationID.Text))
    MsgBox "复制线路[" & txtNewRoute.Text & "]成功", vbInformation, Me.Caption
    If m_bIsParent Then
        frmAllRoute.AddList Trim(ResolveDisplay(txtNewRoute.Text))
    End If
    SetNormal
    Unload Me
    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub cmdClose_Click()
    Set m_oRoute = Nothing
    Unload Me
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    cmdCopy.Enabled = False
    m_oRoute.Init g_oActiveUser
    If m_szOldRouteID <> "" Then
        txtRoute.Text = m_szOldRouteID
    End If
End Sub



Private Sub txtEndStationID_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectStation(, False)
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtEndStationID.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub

Private Sub txtNewRoute_Change()
    IsEnabled
End Sub

Private Sub txtRoute_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectRoute(False)
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRoute.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))

End Sub

Private Sub txtRouteName_Change()
    IsEnabled
End Sub
Private Sub txtRoute_Change()
    IsEnabled
End Sub

Private Function IsEnabled()
    If txtRoute.Text <> "" And txtNewRoute.Text <> "" And txtRouteName.Text <> "" Then
        cmdCopy.Enabled = True
    Else
        cmdCopy.Enabled = False
    End If
End Function

