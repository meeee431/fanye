VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSeatType 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "座位类型"
   ClientHeight    =   3225
   ClientLeft      =   3270
   ClientTop       =   3555
   ClientWidth     =   5325
   HelpContextID   =   10000740
   Icon            =   "frmSeatType.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   360
      TabIndex        =   12
      Top             =   2700
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
      MICON           =   "frmSeatType.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1860
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1320
      Width           =   2910
   End
   Begin VB.TextBox txtSeatTypeID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1860
      MaxLength       =   25
      TabIndex        =   1
      Top             =   930
      Width           =   2910
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   8
      Top             =   0
      Width           =   6615
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
         Caption         =   "请修改或新增车次种类信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   2250
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3930
      TabIndex        =   7
      Top             =   2700
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmSeatType.frx":0028
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
      Height          =   315
      Left            =   2700
      TabIndex        =   6
      Top             =   2700
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
      MICON           =   "frmSeatType.frx":0044
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
      Height          =   840
      Left            =   -150
      TabIndex        =   11
      Top             =   2460
      Width           =   8745
   End
   Begin FText.asFlatMemo txtAnnotation 
      Height          =   630
      Left            =   1860
      TabIndex        =   5
      Top             =   1710
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   1111
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotForeColor=   -2147483628
      ButtonHotBackColor=   -2147483632
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   255
      Left            =   375
      TabIndex        =   4
      Top             =   1740
      Width           =   1020
   End
   Begin VB.Label lblObjectName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位类型名称(&N):"
      Height          =   180
      Left            =   375
      TabIndex        =   2
      Top             =   1380
      Width           =   1440
   End
   Begin VB.Label lblObjectA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位类型代码(&C):"
      Height          =   180
      Left            =   375
      TabIndex        =   0
      Top             =   990
      Width           =   1440
   End
End
Attribute VB_Name = "frmSeatType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus

Public mszSeatTypeID As String
Private moSeatType As SeatType

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrHandle
    Select Case Status
        Case EFormStatus.EFS_AddNew
            moSeatType.AddNew
            moSeatType.SeatTypeID = txtSeatTypeID.Text
            moSeatType.SeatTypeName = txtName.Text
            moSeatType.Annotation = txtAnnotation.Text
            moSeatType.Update
        Case EFormStatus.EFS_Modify
            moSeatType.Identify txtSeatTypeID.Text
            moSeatType.SeatTypeName = txtName.Text
            moSeatType.Annotation = txtAnnotation.Text
            moSeatType.Update
    End Select
        
    Dim aszInfo(1 To 3) As String
    aszInfo(1) = Trim(txtSeatTypeID.Text)
    aszInfo(2) = Trim(txtName.Text)
    aszInfo(3) = Trim(txtAnnotation.Text)
    
    '刷新基本信息窗体
    Dim oListItem As ListItem
    If Status = EFormStatus.EFS_Modify Then
        frmBaseInfo.UpdateList aszInfo
        Unload Me
        Exit Sub
    End If
    If Status = EFormStatus.EFS_AddNew Then
        frmBaseInfo.AddList aszInfo
        RefreshSeatType
        txtSeatTypeID.SetFocus
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
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
    
    Set moSeatType = CreateObject("STBase.SeatType")
    moSeatType.Init g_oActiveUser
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOk.Caption = "新增(&A)"
            RefreshSeatType
        Case EFormStatus.EFS_Modify
           txtSeatTypeID.Enabled = False
           RefreshSeatType
        Case EFormStatus.EFS_Show
           txtSeatTypeID.Enabled = False
           RefreshSeatType
    End Select
    cmdOk.Enabled = False
    
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub
Private Sub RefreshSeatType()
    If Status = EFS_AddNew Then
        txtSeatTypeID.Text = ""
        txtAnnotation.Text = ""
        txtName.Text = ""
    Else
        txtSeatTypeID.Text = mszSeatTypeID
        moSeatType.Identify mszSeatTypeID
        txtAnnotation.Text = moSeatType.Annotation
        txtName.Text = moSeatType.SeatTypeName
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set moSeatType = Nothing
    SaveFormPos Me
End Sub


Private Sub txtAnnotation_Change()
    IsSave
End Sub
Private Sub txtAnnotation_GotFocus()
    cmdOk.Default = False
End Sub

Private Sub txtAnnotation_LostFocus()
    cmdOk.Default = True
End Sub

Private Sub txtSeatTypeID_Change()
    IsSave
    FormatTextBoxBySize txtSeatTypeID, 3
End Sub
Private Sub IsSave()
    If txtSeatTypeID.Text = "" Or txtName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtName_Change()
    IsSave
    FormatTextBoxBySize txtName, 10
End Sub
