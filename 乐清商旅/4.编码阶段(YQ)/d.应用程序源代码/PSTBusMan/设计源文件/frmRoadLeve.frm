VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmRoadLevel 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "公路等级"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   HelpContextID   =   10000170
   Icon            =   "frmRoadLeve.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   540
      TabIndex        =   12
      Top             =   2730
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "帮助"
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
      MICON           =   "frmRoadLeve.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtRoadLevel 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2025
      TabIndex        =   1
      Top             =   1005
      Width           =   2640
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   5475
      TabIndex        =   8
      Top             =   0
      Width           =   5475
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   9
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增公路等级信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   2250
      End
   End
   Begin VB.TextBox txtSimpleName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2025
      TabIndex        =   3
      Top             =   1425
      Width           =   2640
   End
   Begin VB.TextBox txtLevelName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2025
      TabIndex        =   5
      Top             =   1830
      Width           =   2640
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Top             =   2730
      Width           =   1065
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
      MICON           =   "frmRoadLeve.frx":0166
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
      Left            =   4230
      TabIndex        =   7
      Top             =   2730
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "关闭(&O)"
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
      MICON           =   "frmRoadLeve.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   750
      Left            =   -120
      TabIndex        =   11
      Top             =   2490
      Width           =   8745
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公路等级简称(&S):"
      Height          =   180
      Left            =   540
      TabIndex        =   2
      Top             =   1455
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公路等级名称(&N):"
      Height          =   180
      Left            =   540
      TabIndex        =   4
      Top             =   1875
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公路等级代码(&I):"
      Height          =   180
      Left            =   540
      TabIndex        =   0
      Top             =   1035
      Width           =   1440
   End
End
Attribute VB_Name = "frmRoadLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus

Private moRoadLevel As RoadLevel '公路等级对象 RoadLevel
Public mszRoadLevel As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    
    Select Case Status
        Case EFormStatus.EFS_AddNew
            moRoadLevel.AddNew
            moRoadLevel.RoadLevelCode = txtRoadLevel.Text
            moRoadLevel.RoadLevelShortName = txtSimpleName.Text
            moRoadLevel.RoadLeveName = txtLevelName.Text
            moRoadLevel.Update
      Case EFormStatus.EFS_Modify
            moRoadLevel.Identify txtRoadLevel.Text
            moRoadLevel.RoadLevelShortName = txtSimpleName.Text
            moRoadLevel.RoadLeveName = txtLevelName.Text
            moRoadLevel.Update
    End Select
        
    Dim aszInfo(1 To 3) As String
    aszInfo(1) = Trim(txtRoadLevel.Text)
    aszInfo(2) = Trim(txtSimpleName.Text)
    aszInfo(3) = Trim(txtLevelName)
    
    '刷新基本信息窗体
    Dim oListItem As ListItem
    If Status = EFormStatus.EFS_Modify Then
        frmBaseInfo.UpdateList aszInfo
        Unload Me
        Exit Sub
    End If
    If Status = EFormStatus.EFS_AddNew Then
        frmBaseInfo.AddList aszInfo
        RefreshRoadLevel
        txtRoadLevel.SetFocus
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg

End Sub




Private Sub CoolButton1_Click()
DisplayHelp Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    
           Case vbKeyReturn
                SendKeys "{TAB}"
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    '布置窗体
    AlignFormPos Me
    
    Set moRoadLevel = CreateObject("STBase.RoadLevel")
    moRoadLevel.Init g_oActiveUser
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOk.Caption = "新增(&A)"
            RefreshRoadLevel
        Case EFormStatus.EFS_Modify
           txtRoadLevel.Enabled = False
           RefreshRoadLevel
        Case EFormStatus.EFS_Show
           txtRoadLevel.Enabled = False
           RefreshRoadLevel
    End Select
    cmdOk.Enabled = False
    
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub

Public Sub RefreshRoadLevel()
    If Status = EFS_AddNew Then
        txtRoadLevel.Text = ""
        txtSimpleName.Text = ""
        txtLevelName.Text = ""
    Else
        txtRoadLevel.Text = mszRoadLevel
        moRoadLevel.Identify Trim(mszRoadLevel)
        txtSimpleName.Text = moRoadLevel.RoadLevelShortName
        txtLevelName.Text = moRoadLevel.RoadLeveName
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moRoadLevel = Nothing
    SaveFormPos Me
End Sub

Private Sub txtLevelName_Change()
    IsSave
    FormatTextBoxBySize txtLevelName, 30
End Sub

Private Sub txtRoadLevel_Change()
    IsSave
    FormatTextBoxBySize txtRoadLevel, 4
End Sub

Private Sub IsSave()
    If Trim(txtRoadLevel.Text) = "" Or Trim(txtLevelName.Text) = "" Or Trim(txtSimpleName.Text) = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtSimpleName_Change()
    IsSave
    FormatTextBoxBySize txtSimpleName, 30
End Sub

