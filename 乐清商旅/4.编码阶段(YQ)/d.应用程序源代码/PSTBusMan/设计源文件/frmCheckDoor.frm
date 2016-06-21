VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCheckDoor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "检票口"
   ClientHeight    =   3585
   ClientLeft      =   2310
   ClientTop       =   3105
   ClientWidth     =   5235
   HelpContextID   =   10000350
   Icon            =   "frmCheckDoor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   255
      TabIndex        =   14
      Top             =   3090
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "帮助(H)"
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
      MICON           =   "frmCheckDoor.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboOwnerStation 
      Height          =   300
      Left            =   1695
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   3030
   End
   Begin VB.TextBox txtCheckID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1695
      MaxLength       =   25
      TabIndex        =   1
      Top             =   885
      Width           =   3015
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   11
      Top             =   -30
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   12
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增检票口信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   2070
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3960
      TabIndex        =   9
      Top             =   3090
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
      MICON           =   "frmCheckDoor.frx":0166
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
      Left            =   2730
      TabIndex        =   8
      Top             =   3090
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
      MICON           =   "frmCheckDoor.frx":0182
      PICN            =   "frmCheckDoor.frx":019E
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
      Left            =   1695
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1275
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   840
      Left            =   -120
      TabIndex        =   10
      Top             =   2850
      Width           =   8745
   End
   Begin FText.asFlatMemo txtAnnotation 
      Height          =   630
      Left            =   1695
      TabIndex        =   7
      Top             =   2100
      Width           =   3015
      _ExtentX        =   5318
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "所属车站(&S):"
      Height          =   180
      Left            =   390
      TabIndex        =   4
      Top             =   1710
      Width           =   1080
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   255
      Left            =   390
      TabIndex        =   6
      Top             =   2130
      Width           =   1020
   End
   Begin VB.Label lblObjectA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检票口代码(&I):"
      Height          =   180
      Left            =   390
      TabIndex        =   0
      Top             =   945
      Width           =   1260
   End
   Begin VB.Label lblObjectName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N):"
      Height          =   180
      Left            =   390
      TabIndex        =   2
      Top             =   1335
      Width           =   720
   End
End
Attribute VB_Name = "frmCheckDoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus

Public mszCheckID As String
Private moCheckDoor As CheckGate

Private Sub cboOwnerStation_Change()
    IsSave
End Sub

Private Sub cboOwnerStation_Click()
    IsSave
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    Select Case Status
      Case EFormStatus.EFS_AddNew
            moCheckDoor.AddNew
            moCheckDoor.CheckGateCode = txtCheckID.Text
            moCheckDoor.CheckGateName = txtName.Text
            moCheckDoor.Annotation = txtAnnotation.Text
            moCheckDoor.SellStationID = ResolveDisplay(cboOwnerStation.Text)
            moCheckDoor.Update
      Case EFormStatus.EFS_Modify
            moCheckDoor.Identify txtCheckID.Text
            moCheckDoor.CheckGateName = txtName.Text
            moCheckDoor.Annotation = txtAnnotation.Text
            moCheckDoor.SellStationID = ResolveDisplay(cboOwnerStation.Text)
            moCheckDoor.Update
      End Select
        
    Dim aszInfo(1 To 3) As String
    aszInfo(1) = Trim(txtCheckID.Text)
    aszInfo(2) = Trim(txtName.Text)
    Dim szName As String
    ResolveDisplay cboOwnerStation.Text, szName
    aszInfo(3) = EncodeString(szName) & Trim(txtAnnotation.Text)
    
    '刷新基本信息窗体
    Dim oListItem As ListItem
    If Status = EFormStatus.EFS_Modify Then
        frmBaseInfo.UpdateList aszInfo
        Unload Me
        Exit Sub
    End If
    If Status = EFormStatus.EFS_AddNew Then
        frmBaseInfo.AddList aszInfo
        RefreshCheckDoor
        txtCheckID.SetFocus
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select

End Sub
Private Sub Form_Load()
    Dim nCount As Integer
    On Error GoTo ErrHandle
    '布置窗体
    AlignFormPos Me
    
    Set moCheckDoor = CreateObject("STBase.CheckGate")
    moCheckDoor.Init g_oActiveUser
    
    '添加售票站点
    Dim i As Integer
    With cboOwnerStation
    .Clear
    nCount = ArrayLength(g_atAllSellStation)
    For i = 1 To nCount
        .AddItem MakeDisplayString(g_atAllSellStation(i).szSellStationID, g_atAllSellStation(i).szSellStationFullName)
    Next i
    End With
    
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOk.Caption = "新增(&A)"
            RefreshCheckDoor
        Case EFormStatus.EFS_Modify
            txtCheckID.Enabled = False
           RefreshCheckDoor
        Case EFormStatus.EFS_Show
            txtCheckID.Enabled = False
           RefreshCheckDoor
    End Select
    cmdOk.Enabled = False
    
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moCheckDoor = Nothing
    SaveFormPos Me
End Sub


Public Sub RefreshCheckDoor()
    
    If Status = EFS_AddNew Then
        txtCheckID.Text = ""
        txtAnnotation.Text = ""
        txtName.Text = ""
    Else
        txtCheckID.Text = mszCheckID
        moCheckDoor.Identify mszCheckID
        txtAnnotation.Text = moCheckDoor.Annotation
        txtName.Text = moCheckDoor.CheckGateName
        Dim i As Integer
        For i = 0 To cboOwnerStation.ListCount - 1
            If moCheckDoor.SellStationID = ResolveDisplay(cboOwnerStation.List(i)) Then
                cboOwnerStation.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
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

Private Sub txtCheckID_Change()
    IsSave
    FormatTextBoxBySize txtCheckID, 2
End Sub

Private Sub IsSave()
    If txtCheckID.Text = "" Or txtName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtName_Change()
    IsSave
    FormatTextBoxBySize txtName, 50
End Sub
