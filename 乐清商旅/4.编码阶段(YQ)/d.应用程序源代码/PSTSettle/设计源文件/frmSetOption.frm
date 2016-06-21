VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   3930
   ClientLeft      =   3690
   ClientTop       =   3825
   ClientWidth     =   6225
   HelpContextID   =   20000200
   Icon            =   "frmSetOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "Modal"
   Begin VB.Frame Frame3 
      Caption         =   "检票提示音"
      Height          =   2775
      Left            =   330
      TabIndex        =   3
      Top             =   225
      Width           =   5565
      Begin VB.TextBox txtSoundFile 
         Height          =   315
         Left            =   2310
         TabIndex        =   7
         Top             =   840
         Width           =   3105
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "浏览(&B)"
         Height          =   315
         Left            =   4320
         TabIndex        =   4
         Top             =   1245
         Width           =   1095
      End
      Begin MSComctlLib.Slider sldPlayer 
         Height          =   495
         Left            =   2850
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2145
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   0
         SmallChange     =   0
      End
      Begin MCI.MMControl MMControl1 
         Height          =   375
         Left            =   2310
         TabIndex        =   6
         Top             =   1680
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   661
         _Version        =   393216
         AutoEnable      =   0   'False
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         PauseVisible    =   0   'False
         BackVisible     =   0   'False
         StepVisible     =   0   'False
         RecordVisible   =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   "WaveAudio"
         FileName        =   ""
      End
      Begin MSComctlLib.TreeView tvSoundEvent 
         Height          =   2115
         Left            =   120
         TabIndex        =   8
         Top             =   540
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   3731
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2310
         Top             =   2130
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "声音文件名(&F):"
         Height          =   180
         Left            =   2310
         TabIndex        =   10
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "事件(&E)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   15
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   315
      Left            =   2355
      TabIndex        =   0
      Top             =   3345
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   315
      Left            =   3585
      TabIndex        =   1
      Top             =   3345
      Width           =   1155
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   315
      Left            =   4875
      TabIndex        =   2
      Top             =   3345
      Width           =   1155
   End
End
Attribute VB_Name = "frmSetOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nLastNodeIndex As Integer               '上一次选择的Node的Index

Private Sub cmdBrowse_Click()
    Dim szFile As String
    dlgFile.Filter = "所有音效文件(*.wav)|*.wav"
    dlgFile.ShowOpen
    If dlgFile.FileName <> "" Then
        txtSoundFile.Text = dlgFile.FileName
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdok_Click()
On Error GoTo ErrHandle
    Dim oFreeReg As CFreeReg
    
    Set oFreeReg = New CFreeReg
    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    
    
    
    g_tEventSoundPath.CheckSheetCanceled = tvSoundEvent.Nodes("CheckSheetCanceled").Tag
    g_tEventSoundPath.CheckSheetNotExist = tvSoundEvent.Nodes("CheckSheetNotExist").Tag
    g_tEventSoundPath.CheckSheetSelected = tvSoundEvent.Nodes("CheckSheetSelected").Tag
    g_tEventSoundPath.CheckSheetSettled = tvSoundEvent.Nodes("CheckSheetSettled").Tag
    g_tEventSoundPath.CheckSheetValid = tvSoundEvent.Nodes("CheckSheetValid").Tag
    g_tEventSoundPath.ObjectNotSame = tvSoundEvent.Nodes("ObjectNotSame").Tag
    
    oFreeReg.SaveSetting m_cRegSoundKey, "CheckSheetCanceled", g_tEventSoundPath.CheckSheetCanceled
    oFreeReg.SaveSetting m_cRegSoundKey, "CheckSheetNotExist", g_tEventSoundPath.CheckSheetNotExist
    oFreeReg.SaveSetting m_cRegSoundKey, "CheckSheetSelected", g_tEventSoundPath.CheckSheetSelected
    oFreeReg.SaveSetting m_cRegSoundKey, "CheckSheetSettled", g_tEventSoundPath.CheckSheetSettled
    oFreeReg.SaveSetting m_cRegSoundKey, "CheckSheetValid", g_tEventSoundPath.CheckSheetValid
    oFreeReg.SaveSetting m_cRegSoundKey, "ObjectNotSame", g_tEventSoundPath.ObjectNotSame
    
    Unload Me
    Set oFreeReg = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub



Private Sub Form_Load()
    Dim n As Integer
    Dim fraTemp As Frame
    Dim oFreeReg As CFreeReg
 

On Error GoTo here
    AlignFormPos Me
    'CboType.ListIndex = 0
    Set oFreeReg = New CFreeReg
    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
        
    '初始化fraOption(1)
        MMControl1.TimeFormat = mciFormatMilliseconds
        tvSoundEvent.Nodes.Add , , "Root", "结算事件"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "CheckSheetNotExist", "路单不存在"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "CheckSheetCanceled", "路单已作废"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "CheckSheetSettled", "路单已结算"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "CheckSheetSelected", "路单已选择"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "CheckSheetValid", "路单有效"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "ObjectNotSame", "路单有效,但不在所要拆算时间或对象范围之内"
        
       
        tvSoundEvent.Nodes.Item("CheckSheetNotExist").Tag = g_tEventSoundPath.CheckSheetNotExist
        tvSoundEvent.Nodes.Item("CheckSheetCanceled").Tag = g_tEventSoundPath.CheckSheetCanceled
        tvSoundEvent.Nodes.Item("CheckSheetSettled").Tag = g_tEventSoundPath.CheckSheetSettled
        tvSoundEvent.Nodes.Item("CheckSheetSelected").Tag = g_tEventSoundPath.CheckSheetSelected
        tvSoundEvent.Nodes.Item("CheckSheetValid").Tag = g_tEventSoundPath.CheckSheetValid
        tvSoundEvent.Nodes.Item("ObjectNotSame").Tag = g_tEventSoundPath.ObjectNotSame
        
        
    
        tvSoundEvent.Nodes.Item("Root").Expanded = True
        tvSoundEvent.Nodes("CheckSheetNotExist").Selected = True
        txtSoundFile.Text = tvSoundEvent.Nodes("CheckSheetNotExist").Tag
        nLastNodeIndex = tvSoundEvent.SelectedItem.Index
            
    Exit Sub
here:
    ShowErrorMsg
End Sub




Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
    sldPlayer.ClearSel
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
    MMControl1.FileName = txtSoundFile.Text
    MMControl1.Command = "open"
    MMControl1.UpdateInterval = MMControl1.Length / 10
    MMControl1.StopEnabled = True
    MMControl1.Command = "play"

End Sub

Private Sub MMControl1_StatusUpdate()
    If MMControl1.Position < MMControl1.Length Then
        sldPlayer.Value = (MMControl1.Position / MMControl1.Length) * 10 + 1
    Else
        MMControl1.StopEnabled = False
        MMControl1.Command = "close"
        MMControl1.UpdateInterval = 0
        sldPlayer.Value = 0
    End If
End Sub

Private Sub MMControl1_StopClick(Cancel As Integer)
    MMControl1.Command = "Close"
End Sub

Private Sub tvSoundEvent_NodeClick(ByVal Node As MSComctlLib.Node)
    txtSoundFile.Text = Trim(txtSoundFile.Text)
    If txtSoundFile.Text <> "" And Dir(txtSoundFile.Text) = "" Or Right(txtSoundFile.Text, 1) = "\" Then
        MsgBox "此文件不存在！", vbExclamation, Me.Caption
        Dim nTmpLastIndex As Integer
        nTmpLastIndex = nLastNodeIndex
        nLastNodeIndex = Node.Index
        tvSoundEvent.Nodes(nTmpLastIndex).Selected = True
                
        txtSoundFile.SelStart = 0
        txtSoundFile.SelLength = Len(txtSoundFile.Text)
        txtSoundFile.SetFocus
        Exit Sub
    End If
    If Node.Key <> "Root" Then
        If Not cmdBrowse.Enabled Then cmdBrowse.Enabled = True
'        tvSoundEvent.Nodes(nLastNodeIndex).Tag = txtSoundFile.Text
        txtSoundFile.Text = Node.Tag
        nLastNodeIndex = Node.Index
        MMControl1.StopEnabled = False
        If txtSoundFile.Text = "" Then
            MMControl1.PlayEnabled = False
        End If
    Else
        cmdBrowse.Enabled = False
        txtSoundFile.Text = ""
    End If
End Sub


Private Sub txtSoundFile_Change()
    MMControl1.PlayEnabled = False
    MMControl1.StopEnabled = False
    If txtSoundFile.Text <> "" And Dir(txtSoundFile.Text) <> "" Then
        If Right(Trim(txtSoundFile.Text), 1) = "\" Then     '不是一个路径
            Exit Sub
        Else
            MMControl1.PlayEnabled = True
            MMControl1.StopEnabled = True
            tvSoundEvent.SelectedItem.Tag = txtSoundFile.Text
        End If
    End If
    If txtSoundFile.Text = "" Then
        tvSoundEvent.SelectedItem.Tag = ""
    End If
    
End Sub
