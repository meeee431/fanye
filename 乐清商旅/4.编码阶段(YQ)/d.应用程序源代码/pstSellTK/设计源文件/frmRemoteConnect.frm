VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmRemoteConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "远程连接"
   ClientHeight    =   3780
   ClientLeft      =   3135
   ClientTop       =   2490
   ClientWidth     =   5670
   HelpContextID   =   3001801
   Icon            =   "frmRemoteConnect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdhelp 
      Height          =   315
      Left            =   150
      TabIndex        =   16
      Top             =   3420
      Width           =   915
      _ExtentX        =   1614
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
      MICON           =   "frmRemoteConnect.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   -990
      TabIndex        =   12
      Top             =   3780
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   "AviVideo"
      FileName        =   "liandh-1.avi"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "连接状态"
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   5355
      Begin VB.PictureBox ptProgress 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   960
         Picture         =   "frmRemoteConnect.frx":05A6
         ScaleHeight     =   555
         ScaleWidth      =   3105
         TabIndex        =   13
         Top             =   840
         Width           =   3105
         Begin VB.Image imgFailed 
            Height          =   540
            Left            =   1140
            Picture         =   "frmRemoteConnect.frx":0FC2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox ptEarth 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   4140
         Picture         =   "frmRemoteConnect.frx":12CC
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   14
         Top             =   630
         Width           =   795
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   450
         Picture         =   "frmRemoteConnect.frx":1BC4
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Height          =   1215
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5115
      End
   End
   Begin MSComctlLib.ImageList imglstIcon 
      Left            =   15
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemoteConnect.frx":24A4
            Key             =   "NotCon"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemoteConnect.frx":25FE
            Key             =   "Con"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo icboUnit 
      Height          =   330
      Left            =   1260
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "imglstIcon"
   End
   Begin RTComctl3.CoolButton cmdConnect 
      Default         =   -1  'True
      Height          =   315
      Left            =   3330
      TabIndex        =   2
      Top             =   3405
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "连接(&R)"
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
      MICON           =   "frmRemoteConnect.frx":2758
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   3405
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭"
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
      MICON           =   "frmRemoteConnect.frx":2774
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
      BackColor       =   &H00E0E0E0&
      Caption         =   "单位信息:"
      Height          =   915
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5355
      Begin VB.Label lblUnitAnnotation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1140
         TabIndex        =   10
         Top             =   540
         Width           =   90
      End
      Begin VB.Label lblIPAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3840
         TabIndex        =   9
         Top             =   300
         Width           =   90
      End
      Begin VB.Label lblUnitFullName 
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1140
         TabIndex        =   8
         Top             =   300
         Width           =   1890
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位注释:"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP地址:"
         Height          =   180
         Left            =   3120
         TabIndex        =   6
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位全称:"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "远程单位(&U):"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmRemoteConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_bOk As Boolean
Public m_szUnitID As String
Public m_szUnitName As String
Dim m_oUnit As Unit




Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    Dim szTemp As String
    Dim echoReturn As ICMP_ECHO_REPLY
    Dim szUnitIP As String
    On Error GoTo Error_Handle

    If icboUnit.Text <> "" Then
        On Error Resume Next
        BeginPlayAvi
        
   
        
'        Dim dtBegin As Date
'        dtBegin = Time
'        Do While DateDiff("s", dtBegin, Time) < 20
'        Loop
        On Error GoTo Error_Handle
        m_szUnitID = ResolveDisplay(icboUnit.Text, m_szUnitName)
        szUnitIP = m_oSell.GetUnitIP(m_szUnitID)
'        IsConnected szUnitIP, echoReturn
'        If echoReturn.status <> IP_SUCCESS Then
'            MsgBox "[一般性网络错误]连接服务器失败！！", vbInformation, "错误！"
'            StopPlayAvi False
'            Exit Sub
'        End If
        szTemp = m_oSell.SellUnitCode
        m_oSell.SellUnitCode = m_szUnitID
        m_oSell.SellUnitCode = szTemp
        m_bOk = True
        
        StopPlayAvi True
    End If
    Unload Me
    Exit Sub
Error_Handle:
    
    StopPlayAvi False
    ShowErrorMsg
End Sub

Private Sub cmdhelp_Click()
    DisplayHelp Me
End Sub

Private Sub Form_Load()
    Dim auiTemp() As TUnit
    Dim nCount As Integer
    Dim i As Integer
    auiTemp = m_oSell.GetServiceUnitInfo
    nCount = ArrayLength(auiTemp)
    
    For i = 1 To nCount
        
        If IsUnitConnected(auiTemp(i).szUnitID) Then
            icboUnit.ComboItems.Add , MakeDisplayString(1, auiTemp(i).szUnitID), MakeDisplayString(auiTemp(i).szUnitID, auiTemp(i).szUnitShortName), "Con"
        Else
            icboUnit.ComboItems.Add , MakeDisplayString(0, auiTemp(i).szUnitID), MakeDisplayString(auiTemp(i).szUnitID, auiTemp(i).szUnitShortName), "NotCon"
        End If
    Next
    Set m_oUnit = New Unit
    m_oUnit.Init m_oAUser
    If nCount > 0 Then
        icboUnit.ComboItems(1).Selected = True
        ShowUnitInfo
    End If
    MMControl1.hWndDisplay = ptProgress.hwnd

   
    m_bOk = False
    
End Sub
Private Function ShowUnitInfo() As Long
    
    If Not icboUnit.SelectedItem Is Nothing Then
        m_oUnit.Identify ResolveDisplay(icboUnit.SelectedItem.Text)
        lblUnitFullName.Caption = m_oUnit.UnitFullName
        lblIPAddress.Caption = m_oUnit.HostName
        lblUnitAnnotation.Caption = m_oUnit.UnitAnnotation
        If ResolveDisplay(icboUnit.SelectedItem.Key) = "1" Then
            cmdConnect.Enabled = False
            ShowConnectStatus 2
        Else
            cmdConnect.Enabled = True
            ShowConnectStatus 0
        End If
    End If
End Function

'判断给定的单位是否已经连接上了
Private Function IsUnitConnected(pszUnitID As String) As Boolean
    Dim i As Integer
    IsUnitConnected = False
    For i = 1 To MDISellTicket.tsUnit.Tabs.count
        If pszUnitID = MDISellTicket.tsUnit.Tabs(i).Tag Then
            IsUnitConnected = True
            Exit For
        End If
    Next
End Function

Private Sub icboUnit_Click()
    ShowUnitInfo
End Sub

'显示HTMLHELP,直接拷贝
Private Sub DisplayHelp(frmTemp As Form, Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long
    
    Select Case HelpType
        Case content
            lActiveControl = frmTemp.ActiveControl.HelpContextID
            If lActiveControl = 0 Then
                TopicID = Me.HelpContextID
                CallHTMLShowTopicID
            Else
                TopicID = lActiveControl
                CallHTMLShowTopicID
            End If
        Case Index
            CallHTMLHelpIndex
        Case Support
            TopicID = clSupportID
            CallHTMLShowTopicID
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub BeginPlayAvi()
    ptProgress.Visible = True
    MMControl1.Command = "open"
    
    MMControl1.Command = "play"
    
End Sub

Private Sub StopPlayAvi(pbConnected As Boolean)
    MMControl1.Command = "close"
    
    ShowConnectStatus IIf(pbConnected, 2, 1)
End Sub

'0,为还没有连接，1为连不上，2为已连接
Private Sub ShowConnectStatus(pnStatus As Integer)
    Select Case pnStatus
        Case 0
        ptProgress.Visible = False
        Case 1
        ptProgress.Visible = True
        imgFailed.Visible = True
        Case 2
        ptProgress.Visible = True
        imgFailed.Visible = False
    End Select
    
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
    MMControl1.Command = "seek"
    MMControl1.Command = "play"
End Sub



