VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSplitItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "拆算费用项"
   ClientHeight    =   4290
   ClientLeft      =   5970
   ClientTop       =   3285
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5730
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton optNotUsed 
      BackColor       =   &H00E0E0E0&
      Caption         =   "未使用"
      Height          =   315
      Left            =   3210
      TabIndex        =   16
      Top             =   1943
      Width           =   1095
   End
   Begin VB.OptionButton optUsed 
      BackColor       =   &H00E0E0E0&
      Caption         =   "使用"
      Height          =   255
      Left            =   2250
      TabIndex        =   15
      Top             =   1973
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   14
      Top             =   690
      Width           =   8775
   End
   Begin VB.ComboBox cboAllowModify 
      Height          =   300
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2940
      Width           =   2025
   End
   Begin VB.ComboBox cboType 
      Height          =   300
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2430
      Width           =   2025
   End
   Begin VB.TextBox txtSplitName 
      Height          =   315
      Left            =   2250
      TabIndex        =   1
      Text            =   "公建金"
      Top             =   1560
      Width           =   2025
   End
   Begin VB.TextBox txtSplitID 
      Height          =   315
      Left            =   2250
      TabIndex        =   0
      Text            =   "0002"
      Top             =   1050
      Width           =   2025
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   1740
      Left            =   -90
      TabIndex        =   8
      Top             =   3390
      Width           =   9465
      Begin RTComctl3.CoolButton cmdok 
         Default         =   -1  'True
         Height          =   330
         Left            =   2850
         TabIndex        =   4
         Top             =   330
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
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
         MICON           =   "frmSplitItem.frx":0000
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
         Height          =   330
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
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
         MICON           =   "frmSplitItem.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox ptCaption 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7965
      TabIndex        =   6
      Top             =   0
      Width           =   7965
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请设置折算费用项:"
         Height          =   180
         Left            =   420
         TabIndex        =   7
         Top             =   300
         Width           =   1530
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "是否允许修改:"
      Height          =   180
      Left            =   900
      TabIndex        =   13
      Top             =   3000
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "拆算类型:"
      Height          =   180
      Left            =   900
      TabIndex        =   12
      Top             =   2490
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "使用状态:"
      Height          =   180
      Left            =   900
      TabIndex        =   11
      Top             =   2010
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "拆算项名称:"
      Height          =   180
      Left            =   900
      TabIndex        =   10
      Top             =   1627
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "拆算项代码:"
      Height          =   180
      Left            =   930
      TabIndex        =   9
      Top             =   1117
      Width           =   990
   End
End
Attribute VB_Name = "frmSplitItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Status As EFormStatus
Private m_oSplitItem As New SplitItem
Public szSplitItemID As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
On Error GoTo here
    '保存更改
    m_oSplitItem.SplitItemID = txtSplitID.Text
    m_oSplitItem.SplitItemName = txtSplitName.Text
    m_oSplitItem.SplitStatus = IIf(optUsed.Value, 1, 0) 'chkUserStatus.Value
    m_oSplitItem.SplitType = cboType.ListIndex
    m_oSplitItem.AllowModify = cboAllowModify.ListIndex
    m_oSplitItem.Update
    Unload Me
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    On Error GoTo err
    AlignFormPos Me
    '初始化
    With Me
        With .cboType
            
            .Clear
            .AddItem "拆给对方公司", 0
            .AddItem "拆给站方", 1
            .AddItem "留给本公司", 2
            .ListIndex = 0
        End With
        .cboAllowModify.Clear
        .cboAllowModify.AddItem "不允许修改", 0
        .cboAllowModify.AddItem "允许修改", 1
        .cboAllowModify.ListIndex = 0
    End With
    txtSplitID.Enabled = False
    m_oSplitItem.Init g_oActiveUser
    m_oSplitItem.Identify szSplitItemID
    txtSplitID.Text = m_oSplitItem.SplitItemID
    txtSplitName = m_oSplitItem.SplitItemName
    If m_oSplitItem.SplitStatus = CS_SplitItemNotUse Then
        optNotUsed.Value = True
    Else
        optUsed.Value = True
        
'        optUsed.Enabled = False
'        optNotUsed.Enabled = False
    End If
    cboType.Text = cboType.List(m_oSplitItem.SplitType)
    cboAllowModify.Text = cboAllowModify.List(m_oSplitItem.AllowModify)
    

    Exit Sub
err:
    ShowErrorMsg
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    frmBaseInfo.lvObject.SelectedItem = txtSplitID.Text
    frmBaseInfo.lvObject.SelectedItem.SubItems(1) = txtSplitName.Text
    frmBaseInfo.lvObject.SelectedItem.SubItems(2) = GetSplitStatus(IIf(optUsed.Value, 1, 0))
    frmBaseInfo.lvObject.SelectedItem.SubItems(3) = GetSplitType(cboType.ListIndex)
    frmBaseInfo.lvObject.SelectedItem.SubItems(4) = GetAllowModify(cboAllowModify.ListIndex)
End Sub

