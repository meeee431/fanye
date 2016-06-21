VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A635D87B-E561-11D2-A5F0-D56D5F7BA003}#3.2#0"; "asfbordr.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmChgStartTktNumber 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "更改起始票号和结束票号"
   ClientHeight    =   1740
   ClientLeft      =   2340
   ClientTop       =   3240
   ClientWidth     =   4845
   ControlBox      =   0   'False
   HelpContextID   =   4000020
   Icon            =   "frmChgStartTktNumber.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtEndFirst 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   330
      TabIndex        =   5
      Top             =   1200
      Width           =   825
   End
   Begin VB.TextBox txtEndLast 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   1770
   End
   Begin VB.TextBox txtLast 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      TabIndex        =   1
      Top             =   420
      Width           =   1770
   End
   Begin VB.TextBox txtFirst 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   330
      TabIndex        =   4
      Top             =   420
      Width           =   825
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3510
      TabIndex        =   7
      Top             =   1020
      Width           =   1125
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消(&C)"
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
      MICON           =   "frmChgStartTktNumber.frx":000C
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
      Default         =   -1  'True
      Height          =   345
      Left            =   3510
      TabIndex        =   6
      Top             =   540
      Width           =   1125
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmChgStartTktNumber.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin asFlatBorder.FlatBorderControl FlatBorderControl1 
      Left            =   2190
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      BorderColor     =   16744576
      Registered      =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   420
      Left            =   3075
      TabIndex        =   8
      Top             =   420
      Width           =   255
      _ExtentX        =   423
      _ExtentY        =   741
      _Version        =   393216
      BuddyControl    =   "txtLast"
      BuddyDispid     =   196611
      OrigLeft        =   3075
      OrigTop         =   405
      OrigRight       =   3315
      OrigBottom      =   825
      Max             =   10000000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   420
      Left            =   3075
      TabIndex        =   9
      Top             =   1200
      Width           =   255
      _ExtentX        =   423
      _ExtentY        =   741
      _Version        =   393216
      BuddyControl    =   "txtEndLast"
      BuddyDispid     =   196610
      OrigLeft        =   3075
      OrigTop         =   1200
      OrigRight       =   3315
      OrigBottom      =   1620
      Max             =   10000000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入结束票号(&E):"
      Height          =   225
      Left            =   330
      TabIndex        =   2
      Top             =   930
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入起始票号(&S):"
      Height          =   225
      Left            =   330
      TabIndex        =   0
      Top             =   150
      Width           =   2115
   End
End
Attribute VB_Name = "frmChgStartTktNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_bOk As Boolean
Public m_bNoCancel As Boolean



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If txtLast.Text <> "" And txtEndLast.Text <> "" Then
        m_lTicketNo = txtLast.Text
        m_lEndTicketNo = txtEndLast.Text
    Else
        m_lTicketNo = 0
        m_lEndTicketNo = 0
    End If
    If Len(txtFirst.Text) > m_oParam.TicketPrefixLen Then
        ShowMsg "票号前缀长度应小于系统参数限定值,无法更改"
        Exit Sub
    End If
    If Val(m_lEndTicketNo) < Val(m_lTicketNo) Then
        ShowMsg "结束票号应大于起始票号！"
        Exit Sub
    End If
    m_lEndTicketNoOld = m_lEndTicketNo
    
'    If CLng(txtLast.Text) > 0 Then
''        If m_oParam.TicketIfSupervise Then
'
'            Dim oTicketMan As New TicketMan
'
'            oTicketMan.Init m_oAUser
'
'            If Not oTicketMan.StartUserSheet(MakeTicketNo(m_lTicketNo, txtFirst.Text), MakeTicketNo(m_lEndTicketNo, txtEndFirst.Text)) Then
'                ShowMsg "票号没有领用记录"
'                Exit Sub
'            End If
'
''        End If
'    End If
    
    m_szTicketPrefix = txtFirst.Text
    m_bOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    cmdCancel.Enabled = Not m_bNoCancel
    txtFirst.Text = m_szTicketPrefix
    txtEndFirst.Text = txtFirst.Text
    txtLast.Text = m_lTicketNo
    m_lEndTicketNoOld = m_lEndTicketNo
    If m_lEndTicketNo = 0 Then
        txtEndLast.Text = txtLast.Text
'        ShowMsg "用户还未领票，请先去领票！"
        Exit Sub
    Else
        txtEndLast.Text = m_lEndTicketNo
    End If
    txtFirst.MaxLength = m_oParam.TicketPrefixLen
    txtEndFirst.MaxLength = txtFirst.MaxLength
    txtLast.MaxLength = TicketNoNumLen()
    txtEndLast.MaxLength = txtLast.MaxLength
    m_bOk = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = m_bNoCancel And (Not m_bOk)
End Sub

Private Sub txtEndFirst_Change()
    txtFirst.Text = txtEndFirst.Text
End Sub

Private Sub txtFirst_Change()
    txtEndFirst.Text = txtFirst.Text
End Sub

Private Sub txtLast_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lTemp As Long
    If txtLast.Text = "" Then
        lTemp = 0
    Else
        lTemp = CLng(txtLast.Text)
    End If
    
    If KeyCode = vbKeyDown Then
        lTemp = IIf(lTemp - 1 >= 0, lTemp - 1, 0)
        KeyCode = 0
        txtLast.Text = lTemp
    ElseIf KeyCode = vbKeyUp Then
        lTemp = lTemp + 1
        KeyCode = 0
        txtLast.Text = lTemp
    End If
    
    
End Sub

Private Sub txtEndLast_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lTemp As Long
    If txtEndLast.Text = "" Then
        lTemp = 0
    Else
        lTemp = CLng(txtEndLast.Text)
    End If
    
    If KeyCode = vbKeyDown Then
        lTemp = IIf(lTemp - 1 >= 0, lTemp - 1, 0)
        KeyCode = 0
        txtEndLast.Text = lTemp
    ElseIf KeyCode = vbKeyUp Then
        lTemp = lTemp + 1
        KeyCode = 0
        txtEndLast.Text = lTemp
    End If
    
    
End Sub

Private Sub txtLast_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEndLast_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

'显示HTMLHELP,直接拷贝
Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long
    
    Select Case HelpType
        Case content
            lActiveControl = Me.ActiveControl.HelpContextID
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
