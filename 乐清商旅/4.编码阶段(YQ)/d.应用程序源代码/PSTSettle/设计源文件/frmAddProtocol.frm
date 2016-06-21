VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAddProtocol 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "拆算协议"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmAddProtocol.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5595
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkProtocol 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "设为默认协议"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1770
      TabIndex        =   2
      Top             =   1980
      Width           =   3315
   End
   Begin RTComctl3.CoolButton cmdok 
      Default         =   -1  'True
      Height          =   330
      Left            =   2910
      TabIndex        =   5
      Top             =   3720
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
      MICON           =   "frmAddProtocol.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdSetProtocol 
      Height          =   330
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "拆算协议项目"
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
      MICON           =   "frmAddProtocol.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtAnnotation 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   1770
      TabIndex        =   3
      Top             =   2550
      Width           =   3345
   End
   Begin VB.TextBox txtProtocolName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1785
      TabIndex        =   1
      Top             =   1545
      Width           =   3315
   End
   Begin VB.TextBox txtProtocolID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1785
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1095
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -180
      TabIndex        =   10
      Top             =   810
      Width           =   8775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   -120
      ScaleHeight     =   825
      ScaleWidth      =   8685
      TabIndex        =   7
      Top             =   0
      Width           =   8685
      Begin VB.Label lblContent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   450
         TabIndex        =   9
         Top             =   420
         Width           =   90
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "协议信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   780
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4170
      TabIndex        =   6
      Top             =   3720
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
      MICON           =   "frmAddProtocol.frx":0044
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
      Height          =   2880
      Left            =   -120
      TabIndex        =   11
      Top             =   3360
      Width           =   9465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认协议(&M):"
      Height          =   180
      Left            =   450
      TabIndex        =   15
      Top             =   2100
      Width           =   1080
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "协议代码(&I):"
      Height          =   285
      Left            =   420
      TabIndex        =   14
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "协议名称(&B):"
      Height          =   180
      Left            =   420
      TabIndex        =   13
      Top             =   1605
      Width           =   1080
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "注释(&A):"
      Height          =   195
      Left            =   420
      TabIndex        =   12
      Top             =   2640
      Width           =   1050
   End
End
Attribute VB_Name = "frmAddProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus
Public m_oProtocol As New Protocol
Public mszProtocolID As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
On Error GoTo ErrHandle
Dim szProtocol As Integer   '默认协议
    Select Case Status
        Case ST_AddObj
            '新增
            
            m_oProtocol.AddNew
            m_oProtocol.ProtocolID = txtProtocolID.Text
            m_oProtocol.ProtocolName = txtProtocolName.Text
            m_oProtocol.Annotation = txtAnnotation.Text
            If chkProtocol.Value = 1 Then
                m_oProtocol.DefaultMark = Default
            Else
                m_oProtocol.DefaultMark = NotDefaule
            End If
            m_oProtocol.Update
            cmdOk.Enabled = False
            
            
            
            frmProtocolItem.m_eStatus = AddStatus
            frmProtocolItem.m_szProtocolID = txtProtocolID.Text
            frmProtocolItem.Show vbModal
            frmBaseInfo.FillItemLists txtProtocolID.Text
            txtProtocolID.Text = ""
            txtProtocolName.Text = ""
            txtAnnotation.Text = ""
            chkProtocol.Value = 0

        Case ST_EditObj
            '修改
            m_oProtocol.ProtocolID = txtProtocolID.Text
            m_oProtocol.ProtocolName = txtProtocolName.Text
            m_oProtocol.Annotation = txtAnnotation.Text
            If chkProtocol.Value = 1 Then
                m_oProtocol.DefaultMark = Default
            Else
                m_oProtocol.DefaultMark = NotDefaule
            End If
            m_oProtocol.Update
            frmBaseInfo.FillItemLists txtProtocolID.Text
            Unload Me
    End Select
    

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
Dim rsTemp As Recordset
    '布置窗体
    AlignFormPos Me
    cmdOk.Enabled = False
    Select Case Status
        Case ST_AddObj
            txtProtocolID.Text = ""
            txtProtocolName.Text = ""
            txtAnnotation.Text = ""
            cmdSetProtocol.Visible = False
        Case ST_EditObj
            cmdSetProtocol.Visible = True
            txtProtocolID.Text = mszProtocolID
            m_oProtocol.Init g_oActiveUser
            m_oProtocol.Identify mszProtocolID
            txtProtocolName.Text = m_oProtocol.ProtocolName
            txtAnnotation.Text = m_oProtocol.Annotation
            If m_oProtocol.DefaultMark = Default Then
                chkProtocol.Value = 1
            Else
                chkProtocol.Value = 0
            End If
            txtProtocolID.Enabled = False
    End Select
    m_oProtocol.Init g_oActiveUser
    Exit Sub
ErrHandle:
    Status = ST_AddObj
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub cmdSetProtocol_Click()
    If cmdOk.Caption = "新增(&A)" Then
        frmProtocolItem.m_eStatus = AddStatus
        frmProtocolItem.m_szProtocolID = txtProtocolID.Text
        frmProtocolItem.ZOrder 0
        frmProtocolItem.Show vbModal
    Else
        frmProtocolItem.m_eStatus = ModifyStatus
        frmProtocolItem.m_szProtocolID = txtProtocolID.Text
        frmProtocolItem.ZOrder 0
        frmProtocolItem.Show vbModal

    End If
End Sub


Private Sub txtProtocolID_Change()
    If Trim(txtProtocolID.Text) = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub


