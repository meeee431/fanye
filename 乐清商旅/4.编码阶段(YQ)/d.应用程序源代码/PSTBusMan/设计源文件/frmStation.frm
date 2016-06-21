VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmStation 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "站点"
   ClientHeight    =   4455
   ClientLeft      =   3345
   ClientTop       =   3375
   ClientWidth     =   5820
   HelpContextID   =   2007801
   Icon            =   "frmStation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   450
      TabIndex        =   14
      Top             =   3990
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "帮助(&H)"
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
      MICON           =   "frmStation.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   3000
      TabIndex        =   12
      Top             =   3990
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "保存(&S)"
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
      MICON           =   "frmStation.frx":0166
      PICN            =   "frmStation.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4290
      TabIndex        =   13
      Top             =   3990
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭(&C)"
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
      MICON           =   "frmStation.frx":051C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   18
      Top             =   780
      Width           =   7215
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   15
      Top             =   0
      Width           =   7185
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增站点信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1890
      End
   End
   Begin VB.TextBox txtStation 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   1110
      Width           =   3015
   End
   Begin FText.asFlatTextBox txtAreaCode 
      Height          =   300
      Left            =   1920
      TabIndex        =   7
      Top             =   2340
      Width           =   3015
      _ExtentX        =   5318
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
   Begin VB.OptionButton optNoSell 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "不可售票(&G)"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3585
      TabIndex        =   11
      Top             =   3255
      Width           =   1380
   End
   Begin VB.OptionButton optCanSell 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "可售票站点(&Z)"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1905
      TabIndex        =   10
      Top             =   3270
      Value           =   -1  'True
      Width           =   1545
   End
   Begin VB.TextBox txtLocalCode 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   9
      Top             =   2745
      Width           =   3015
   End
   Begin VB.TextBox txtStationName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Top             =   1515
      Width           =   3015
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   1350
      Left            =   -120
      TabIndex        =   17
      Top             =   3750
      Width           =   8745
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "本地编码(&L):"
      Height          =   180
      Left            =   795
      TabIndex        =   8
      Top             =   2865
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "站点代码(&I):"
      Height          =   180
      Left            =   795
      TabIndex        =   0
      Top             =   1170
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "站点名称(&N):"
      Height          =   180
      Left            =   795
      TabIndex        =   2
      Top             =   1575
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地区(&D):"
      Height          =   180
      Left            =   795
      TabIndex        =   6
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label lblInputKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入码(&U):"
      Height          =   180
      Left            =   795
      TabIndex        =   4
      Top             =   1980
      Width           =   900
   End
End
Attribute VB_Name = "frmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmStation.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/27
'* Last Revision Date:2002/08/30
'* Brief Description:
'* Relational Document:UI_BS_SM_37.DOC
'**********************************************************
Public szStationID As String
Public m_bIsParent As Boolean '是否父窗体直接调用
Public Status As EFormStatus

Private m_oStation As New Station

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub
Private Sub cmdOk_Click()
    On Error GoTo ErrorHandle
    Select Case Status
    Case EFS_AddNew
        m_oStation.AddNew
        m_oStation.StationID = txtStation.Text
        m_oStation.StationName = txtStationName.Text
        m_oStation.LocalCode = txtLocalCode.Text
        m_oStation.AreaCode = ResolveDisplay(txtAreaCode.Text)
        m_oStation.StationInputCode = txtInput.Text
        If optCanSell.Value = True Then
            m_oStation.StaionLevel = TP_CanSellTicket
        Else
            m_oStation.StaionLevel = TP_CanNotSellTicket
        End If
        m_oStation.Update
        If m_bIsParent Then
            frmAllStation.AddList m_oStation.StationID
            txtStation.Text = ""
            txtStationName.Text = ""
            txtLocalCode.Text = ""
            txtInput.Text = ""
'            txtAreaCode.Text = ""
            txtStation.SetFocus
        End If
    Case EFS_Modify
        m_oStation.Identify txtStation.Text
        m_oStation.StationName = txtStationName.Text
        m_oStation.LocalCode = txtLocalCode.Text
        m_oStation.AreaCode = ResolveDisplay(txtAreaCode.Text)
        m_oStation.StationInputCode = txtInput.Text
        If optCanSell.Value = True Then
            m_oStation.StaionLevel = TP_CanSellTicket
        Else
            m_oStation.StaionLevel = TP_CanNotSellTicket
        End If
        m_oStation.Update
        If m_bIsParent Then
            frmAllStation.UpdateList Trim(frmAllStation.lvStation.SelectedItem.Text)
        End If
        Unload Me
    End Select
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    m_oStation.Init g_oActiveUser
    Select Case Status
    Case EFS_AddNew
    cmdOk.Caption = "新增(&A)"
    '       cmdCancelLock.Enabled = True
    frmStation.HelpContextID = 10000630
    Case EFS_Modify
        RefreshStation
        txtAreaCode.Enabled = True
        txtLocalCode.Enabled = False
        txtStation.Enabled = False
        frmStation.HelpContextID = 10000670
    Case EFS_Show
        RefreshStation
        cmdOk.Caption = "保存(&S)"
        txtAreaCode.Enabled = True
        txtLocalCode.Enabled = False
        txtStation.Enabled = False
        frmStation.HelpContextID = 10000670
    End Select
    cmdOk.Enabled = False
End Sub

Private Sub optCanSell_Click()
    IsSave
End Sub

Private Sub optNoSell_Click()
    IsSave
End Sub

Private Sub txtAreaCode_ButtonClick()
    Dim aszTemp() As String
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    
    aszTemp = oShell.SelectArea(False)
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtAreaCode.Text = aszTemp(1, 1) & "[" & aszTemp(1, 2) & "]"
End Sub

Private Sub txtAreaCode_Change()
    IsSave
End Sub

Private Sub txtInput_Change()
    IsSave
End Sub

Private Sub txtLocalCode_Change()
    IsSave
    If Status <> EFS_Modify Then
        txtStation.Text = ResolveDisplay(txtAreaCode.Text) & txtLocalCode.Text
    End If
End Sub

Public Sub RefreshStation()
    '刷新站点信息
    On Error GoTo ErrorHandle
    m_oStation.Identify szStationID
    txtAreaCode.Text = m_oStation.AreaCode & "[" & m_oStation.AreaName & "]"
    txtInput.Text = m_oStation.StationInputCode
    txtLocalCode.Text = m_oStation.LocalCode
    txtStationName.Text = m_oStation.StationName
    txtStation.Text = szStationID
    If m_oStation.StaionLevel = TP_CanSellTicket Then
        optCanSell.Value = True
        optNoSell.Value = False
    Else
        optCanSell.Value = False
        optNoSell.Value = True
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub IsSave()
    If txtAreaCode.Text = "" Or txtInput.Text = "" Or txtStationName.Text = "" Or txtStation.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtStation_Change()
    txtStation.Text = GetUnicodeBySize(txtStation.Text, 9)
    IsSave
End Sub

Private Sub txtStationName_Change()
    txtStationName.Text = GetUnicodeBySize(txtStationName.Text, 10)
    IsSave
End Sub
