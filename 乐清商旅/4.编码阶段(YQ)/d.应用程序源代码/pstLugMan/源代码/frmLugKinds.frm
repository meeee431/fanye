VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmLugKinds 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "行包种类"
   ClientHeight    =   3780
   ClientLeft      =   4470
   ClientTop       =   2820
   ClientWidth     =   5385
   HelpContextID   =   2001001
   Icon            =   "frmLugKinds.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Top             =   3240
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭(&L)"
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
      MICON           =   "frmLugKinds.frx":0E42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtLugID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1485
      TabIndex        =   1
      Top             =   990
      Width           =   3210
   End
   Begin FText.asFlatMemo txtAnnotation 
      Height          =   810
      Left            =   1485
      TabIndex        =   5
      Top             =   1920
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   1429
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
         TabIndex        =   11
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增行包种类信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   2250
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1485
      TabIndex        =   3
      Top             =   1455
      Width           =   3210
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   2970
      TabIndex        =   6
      Top             =   3240
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
      MICON           =   "frmLugKinds.frx":0E5E
      PICN            =   "frmLugKinds.frx":0E7A
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
      Height          =   930
      Left            =   -150
      TabIndex        =   10
      Top             =   3000
      Width           =   8745
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   180
      Left            =   375
      TabIndex        =   4
      Top             =   1935
      Width           =   720
   End
   Begin VB.Label lblObjectName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N):"
      Height          =   180
      Left            =   345
      TabIndex        =   2
      Top             =   1470
      Width           =   720
   End
   Begin VB.Label lblObjectA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "种类代码(&C):"
      Height          =   180
      Left            =   345
      TabIndex        =   0
      Top             =   1050
      Width           =   1080
   End
End
Attribute VB_Name = "frmLugKinds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As eFormStatus
 '定义类型对象 A
Public mszLugID As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
On Error GoTo ErrHandle

    
    Select Case Status
        Case ST_AddObj
            '新增行包类型
            m_oLuggageKinds.AddNew
            m_oLuggageKinds.KindsCode = Trim(TxtLugID.Text)
            m_oLuggageKinds.KindsName = Trim(txtName.Text)
            m_oLuggageKinds.Annotation = Trim(txtAnnotation.Text)
             m_oLuggageKinds.Update
      Case ST_EditObj
            '修改行包类型
             m_oLuggageKinds.Identify mszLugID
            m_oLuggageKinds.KindsCode = Trim(TxtLugID.Text)
            m_oLuggageKinds.KindsName = Trim(txtName.Text)
            m_oLuggageKinds.Annotation = Trim(txtAnnotation.Text)
            m_oLuggageKinds.Update
    End Select
        
     '将值放入数组中，返回给基本信息窗口
    Dim aszInfo(0 To 3) As String
    aszInfo(0) = Trim(TxtLugID.Text)
    aszInfo(1) = Trim(txtName.Text)
    aszInfo(2) = Trim(txtAnnotation.Text)

    
    '刷新基本信息窗体
    Dim oListItem As ListItem
    If Status = ST_EditObj Then
        frmBaseInfo.UpdateList aszInfo
        Unload Me
        Exit Sub
    End If
    If Status = ST_AddObj Then
        frmBaseInfo.AddList aszInfo
        RefreshLug
        TxtLugID.SetFocus
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
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
    
'    Set moLug = CreateObject("STBase.Area")
    m_oLuggageKinds.Init m_oAUser
    Select Case Status
        Case ST_AddObj
           cmdOk.Caption = "新增(&A)"
           RefreshLug
        Case ST_EditObj
           TxtLugID.Enabled = False
           RefreshLug
'        Case eFormStatus.EFS_Show
'           TxtLugID.Enabled = False
'           RefreshLug
    End Select
    cmdOk.Enabled = False
    Exit Sub
ErrHandle:
    Status = ST_AddObj
    ShowErrorMsg
End Sub

Public Sub RefreshLug()
    If Status = ST_AddObj Then
        TxtLugID.Text = ""
        txtAnnotation.Text = ""
        txtName.Text = ""
    Else
        TxtLugID.Text = mszLugID
        m_oLuggageKinds.Identify Trim(mszLugID)
        txtAnnotation.Text = m_oLuggageKinds.Annotation
        txtName.Text = m_oLuggageKinds.KindsName

    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)

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

Private Sub TxtLugID_Change()
    IsSave
    FormatTextBoxBySize TxtLugID, 4
End Sub
Private Sub IsSave()
    If TxtLugID.Text = "" Or txtName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtName_Change()
    IsSave
    FormatTextBoxBySize txtName, 20
End Sub

