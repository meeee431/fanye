VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusType 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车次种类"
   ClientHeight    =   3255
   ClientLeft      =   1470
   ClientTop       =   2895
   ClientWidth     =   5295
   HelpContextID   =   1000080
   Icon            =   "frmBusType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   330
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
      MICON           =   "frmBusType.frx":0E42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtBusTypeID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1830
      MaxLength       =   25
      TabIndex        =   1
      Top             =   900
      Width           =   2910
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   2715
      TabIndex        =   6
      Top             =   2700
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "frmBusType.frx":0E5E
      PICN            =   "frmBusType.frx":0E7A
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
      MICON           =   "frmBusType.frx":1214
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1830
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1290
      Width           =   2910
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
      Left            =   1830
      TabIndex        =   5
      Top             =   1680
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
      Left            =   360
      TabIndex        =   4
      Top             =   1710
      Width           =   1020
   End
   Begin VB.Label lblObjectA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次种类代码(&I):"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label lblObjectName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次种类名称(&N):"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   1350
      Width           =   1440
   End
End
Attribute VB_Name = "frmBusType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus

Public mszBusTypeID As String
Private moBusType As BusType

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrHandle
    Select Case Status
        Case EFormStatus.EFS_AddNew
            moBusType.AddNew
            moBusType.BusTypeID = txtBusTypeID.Text
            moBusType.BusTypeName = txtName.Text
            moBusType.Annotation = txtAnnotation.Text
            moBusType.Update
        Case EFormStatus.EFS_Modify
            moBusType.Identify txtBusTypeID.Text
            moBusType.BusTypeName = txtName.Text
            moBusType.Annotation = txtAnnotation.Text
            moBusType.Update
    End Select
        
    Dim aszInfo(1 To 3) As String
    aszInfo(1) = Trim(txtBusTypeID.Text)
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
        RefreshBusType
        txtBusTypeID.SetFocus
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
    
    Set moBusType = CreateObject("STBase.BusType")
    moBusType.Init g_oActiveUser
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOk.Caption = "新增(&A)"
            RefreshBusType
        Case EFormStatus.EFS_Modify
           txtBusTypeID.Enabled = False
           RefreshBusType
        Case EFormStatus.EFS_Show
           txtBusTypeID.Enabled = False
           RefreshBusType
    End Select
    cmdOk.Enabled = False
    
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub
Private Sub RefreshBusType()
    If Status = EFS_AddNew Then
        txtBusTypeID.Text = ""
        txtAnnotation.Text = ""
        txtName = ""
    Else
        txtBusTypeID.Text = mszBusTypeID
        moBusType.Identify mszBusTypeID
        txtAnnotation.Text = moBusType.Annotation
        txtName = moBusType.BusTypeName
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set moBusType = Nothing
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

Private Sub txtBusTypeID_Change()
    IsSave
    FormatTextToNumeric txtBusTypeID, False, False
End Sub
Private Sub IsSave()
    If txtBusTypeID.Text = "" Or txtName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub
Private Sub txtName_Change()
    IsSave
    FormatTextBoxBySize txtName, 20
End Sub


