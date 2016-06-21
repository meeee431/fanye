VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Begin VB.Form frmVehicleType 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车型"
   ClientHeight    =   3465
   ClientLeft      =   2160
   ClientTop       =   3450
   ClientWidth     =   5760
   HelpContextID   =   10000100
   Icon            =   "frmVehicleType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   450
      TabIndex        =   16
      Top             =   3000
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
      MICON           =   "frmVehicleType.frx":038A
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
      Left            =   4410
      TabIndex        =   9
      Top             =   3000
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
      MICON           =   "frmVehicleType.frx":03A6
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
      Left            =   3180
      TabIndex        =   8
      Top             =   3000
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmVehicleType.frx":03C2
      PICN            =   "frmVehicleType.frx":03DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtVehicleType 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   1
      Top             =   930
      Width           =   3420
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   11
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -30
         TabIndex        =   12
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增车型信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   1890
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   780
      Left            =   -120
      TabIndex        =   10
      Top             =   2760
      Width           =   8745
   End
   Begin VB.TextBox txtVehicleTypeShortName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   3
      Top             =   1380
      Width           =   3405
   End
   Begin VB.TextBox txtVehicleTypeName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   5
      Top             =   1830
      Width           =   3405
   End
   Begin STSellCtl.ucUpDownText txtEndNumber 
      Height          =   300
      Left            =   4260
      TabIndex        =   14
      Top             =   2250
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      SelectOnEntry   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   100
      Value           =   "40"
   End
   Begin STSellCtl.ucUpDownText txtStartNumber 
      Height          =   300
      Left            =   1770
      TabIndex        =   15
      Top             =   2250
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      SelectOnEntry   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   100
      Value           =   "1"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型代码(&I):"
      Height          =   180
      Left            =   420
      TabIndex        =   0
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型简称(&S):"
      Height          =   180
      Left            =   420
      TabIndex        =   2
      Top             =   1425
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "起始座位号(&B):"
      Height          =   180
      Left            =   420
      TabIndex        =   6
      Top             =   2325
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型名称(&N):"
      Height          =   180
      Left            =   420
      TabIndex        =   4
      Top             =   1875
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "终止座位号(&E):"
      Height          =   180
      Left            =   2880
      TabIndex        =   7
      Top             =   2340
      Width           =   1260
   End
End
Attribute VB_Name = "frmVehicleType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus

Public mszVehicleType As String
Private moVehicleType As VehicleModel

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    Select Case Status
        Case EFormStatus.EFS_AddNew
            moVehicleType.AddNew
            moVehicleType.VehicleModelCode = txtVehicleType.Text
            moVehicleType.VehicleModelName = txtVehicleTypeName.Text
            moVehicleType.VehicleModelShortName = txtVehicleTypeShortName.Text
            moVehicleType.SeatCount = Val(txtEndNumber.Value) - Val(txtStartNumber.Value) + 1
            moVehicleType.StartSeatNumber = txtStartNumber.Value
            moVehicleType.Update
      Case EFormStatus.EFS_Modify
            moVehicleType.Identify txtVehicleType.Text
            moVehicleType.VehicleModelName = txtVehicleTypeName.Text
            moVehicleType.VehicleModelShortName = txtVehicleTypeShortName.Text
            moVehicleType.SeatCount = Val(txtEndNumber.Value) - Val(txtStartNumber.Value) + 1
            moVehicleType.StartSeatNumber = txtStartNumber.Value
            moVehicleType.Update
    End Select
        
    Dim aszInfo(1 To 3) As String
    aszInfo(1) = Trim(txtVehicleType.Text)
    aszInfo(2) = Trim(txtVehicleTypeShortName.Text)
    aszInfo(3) = moVehicleType.SeatCount
    
    '刷新基本信息窗体
    Dim oListItem As ListItem
    If Status = EFormStatus.EFS_Modify Then
        frmBaseInfo.UpdateList aszInfo
        Unload Me
        Exit Sub
    End If
    If Status = EFormStatus.EFS_AddNew Then
        frmBaseInfo.AddList aszInfo
        RefreshVechileType
        txtVehicleType.SetFocus
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
    
    Set moVehicleType = CreateObject("STBase.VehicleModel")
    moVehicleType.Init g_oActiveUser
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOk.Caption = "新增(&A)"
            RefreshVechileType
        Case EFormStatus.EFS_Modify
           txtVehicleType.Enabled = False
           RefreshVechileType
        Case EFormStatus.EFS_Show
           txtVehicleType.Enabled = False
           RefreshVechileType
    End Select
    cmdOk.Enabled = False
    
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub
Private Sub RefreshVechileType()
    If Status = EFS_AddNew Then
        txtVehicleType.Text = ""
        txtVehicleTypeName.Text = ""
        txtVehicleTypeShortName.Text = ""
'        txtEndNumber.Text = moVehicleType.StartSeatNumber + moVehicleType.SeatCount - 1
'        txtStartNumber.Text = moVehicleType.StartSeatNumber
    Else
        txtVehicleType.Text = mszVehicleType
        moVehicleType.Identify Trim(mszVehicleType)
        txtVehicleTypeName.Text = moVehicleType.VehicleModelName
        txtVehicleTypeShortName.Text = moVehicleType.VehicleModelShortName
        txtEndNumber.Value = moVehicleType.StartSeatNumber + moVehicleType.SeatCount - 1
        txtStartNumber.Value = moVehicleType.StartSeatNumber
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set moVehicleType = Nothing
    SaveFormPos Me
End Sub

Private Sub txtEndNumber_Change()
    IsSave
End Sub

Private Sub txtStartNumber_Change()
    IsSave
End Sub

Private Sub txtVehicleType_Change()
    IsSave
    FormatTextBoxBySize txtVehicleType, 3
End Sub


Private Sub IsSave()
    If txtVehicleType.Text = "" Or txtVehicleTypeName.Text = "" Or txtVehicleTypeShortName.Text = "" Or Val(txtStartNumber.Value) = 0 Or Val(txtEndNumber.Value) = 0 Then
    cmdOk.Enabled = False
    Else
    cmdOk.Enabled = True
    End If
End Sub

Private Sub txtVehicleTypeName_Change()
    IsSave
    FormatTextBoxBySize txtVehicleTypeName, 20
End Sub

Private Sub txtVehicleTypeShortName_Change()
    IsSave
    FormatTextBoxBySize txtVehicleTypeShortName, 10
End Sub
