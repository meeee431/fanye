VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAddProtocol 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "拆算协议"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "frmAddProtocol.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6300
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "非默认协议"
      Height          =   315
      Left            =   4170
      TabIndex        =   17
      Top             =   2970
      Width           =   1545
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "随行默认协议"
      Height          =   345
      Left            =   2670
      TabIndex        =   16
      Top             =   2970
      Width           =   1485
   End
   Begin VB.OptionButton Option0 
      BackColor       =   &H00E0E0E0&
      Caption         =   "快件默认协议"
      Height          =   345
      Left            =   1200
      TabIndex        =   15
      Top             =   2970
      Width           =   1455
   End
   Begin RTComctl3.CoolButton cmdok 
      Height          =   330
      Left            =   3510
      TabIndex        =   3
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
      Default         =   -1  'True
      Height          =   330
      Left            =   1980
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "拆算协议项目"
      ENAB            =   0   'False
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
      Height          =   510
      Left            =   1980
      TabIndex        =   2
      Top             =   2130
      Width           =   3345
   End
   Begin VB.TextBox txtProtocolName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1995
      TabIndex        =   1
      Top             =   1635
      Width           =   3315
   End
   Begin VB.TextBox txtProtocolID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1995
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1185
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -180
      TabIndex        =   8
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
      TabIndex        =   5
      Top             =   0
      Width           =   8685
      Begin VB.Label lblContent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请填入协议的相关信息，如果设为默认协议，则不设协议自动按此协议。"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   420
         Width           =   5760
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
         Left            =   270
         TabIndex        =   6
         Top             =   120
         Width           =   780
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4800
      TabIndex        =   4
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
      Enabled         =   0   'False
      Height          =   2880
      Left            =   -120
      TabIndex        =   9
      Top             =   3360
      Width           =   9465
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "默认协议(&D):"
      Height          =   255
      Left            =   630
      TabIndex        =   13
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "协议代码(&I):"
      Height          =   285
      Left            =   630
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "协议名称(&B):"
      Height          =   180
      Left            =   630
      TabIndex        =   11
      Top             =   1695
      Width           =   1080
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "注释(&A):"
      Height          =   195
      Left            =   630
      TabIndex        =   10
      Top             =   2130
      Width           =   1050
   End
End
Attribute VB_Name = "frmAddProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As eFormStatus
Public iChkused As Integer
 '定义类型对象 A
Public mszProtocolID As String



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
Dim szProtocol As Integer   '默认协议
    If Option0.Value = True Then
        szProtocol = ELuggageProtocolDefault.LuggageProtocolDefaultGeneral
    ElseIf Option1.Value = True Then
        szProtocol = ELuggageProtocolDefault.LuggageProtocolDefaultMan
    ElseIf Option2.Value = True Then
        szProtocol = ELuggageProtocolDefault.LuggageProtocolNotDefault
    End If
    Select Case Status
        Case ST_AddObj
            '新增行包类型
         
            m_oProtocol.AddNew
            m_oProtocol.ProtocolID = Trim(txtProtocolID.Text)
            m_oProtocol.ProtocolName = Trim(txtProtocolName.Text)
            m_oProtocol.Annotation = Trim(txtAnnotation.Text)
            m_oProtocol.Default = szProtocol
            m_oProtocol.Update
        Case ST_EditObj
            '修改行包类型
            
            m_oProtocol.Identify mszProtocolID
            m_oProtocol.ProtocolID = Trim(txtProtocolID.Text)
            m_oProtocol.ProtocolName = Trim(txtProtocolName.Text)
            m_oProtocol.Annotation = Trim(txtAnnotation.Text)
            m_oProtocol.Default = szProtocol
            m_oProtocol.Update
    End Select
        
     '将值放入数组中，返回给基本信息窗口
    Dim aszInfo(0 To 3) As String
    aszInfo(0) = Trim(txtProtocolID.Text)
    aszInfo(1) = Trim(txtProtocolName.Text)
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
        txtProtocolID.SetFocus
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdSetProtocol_Click()
    frmProtocol.m_protocolID = Trim(txtProtocolID.Text)
    frmProtocol.lbProtocol.Caption = Trim(txtProtocolID.Text)
    frmProtocol.lbProtocolName.Caption = Trim(txtProtocolName.Text)
    If Option2.Value = True Then
        frmProtocol.lbDefault.Caption = "否"
    Else
        frmProtocol.lbDefault.Caption = "是"
    End If
    frmProtocol.lbRemark.Caption = Trim(txtAnnotation.Text)
    frmProtocol.Show vbModal
    
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
    m_oProtocol.Init m_oAUser
    Select Case Status
        Case ST_AddObj
           cmdOk.Caption = "新增(&A)"
           cmdSetProtocol.Enabled = False
           RefreshLug
           Option0.Value = False
           Option1.Value = False
           Option2.Value = True
        Case ST_EditObj
           txtProtocolID.Enabled = False
           cmdSetProtocol.Enabled = True
           RefreshLug
           Set rsTemp = m_oProtocol.GetAllProtocol(Trim(txtProtocolID.Text))
           If rsTemp.RecordCount > 0 Then
               If FormatDbValue(rsTemp!default_mark) = ELuggageProtocolDefault.LuggageProtocolNotDefault Then
                    Option2.Value = True
               ElseIf FormatDbValue(rsTemp!default_mark) = ELuggageProtocolDefault.LuggageProtocolDefaultGeneral Then
                    Option0.Value = True
               ElseIf FormatDbValue(rsTemp!default_mark) = ELuggageProtocolDefault.LuggageProtocolDefaultMan Then
                    Option1.Value = True
               End If
           End If
    End Select
  cmdOk.Enabled = False

    
    Exit Sub
ErrHandle:
    Status = ST_AddObj
    ShowErrorMsg
End Sub

Public Sub RefreshLug()
    If Status = ST_AddObj Then
        Me.Caption = "新增协议"
        cmdOk.Caption = "新增"
        txtProtocolID.Text = ""
        txtProtocolName.Text = ""
        txtAnnotation.Text = ""
    Else
        Me.Caption = "修改协议"
        cmdOk.Caption = "修改"
        txtProtocolID.Text = mszProtocolID
        m_oProtocol.Identify Trim(mszProtocolID)
        txtProtocolName.Text = m_oProtocol.ProtocolName
        txtAnnotation.Text = m_oProtocol.Annotation
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub Option0_Click()
    IsSave
End Sub

Private Sub Option1_Click()
    IsSave
End Sub

Private Sub Option2_Click()
    IsSave
End Sub

Private Sub txtAnnotation_Change()
    IsSave
End Sub
Private Sub IsSave()
    If txtProtocolID.Text = "" Or txtProtocolName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub
Private Sub txtProtocolID_Change()
    IsSave
    FormatTextBoxBySize txtProtocolID, 4
End Sub
Private Sub txtProtocolName_Change()
    IsSave
    FormatTextBoxBySize txtProtocolName, 20
End Sub

Public Sub SetAllProtocl(values As Integer)
Dim i As Integer, j As Integer
Dim nCount As Integer
Dim vbYesOrNo  As Variant
Dim aszTemp() As String
Dim aszReturn() As String
Dim aszReturn1() As String
Dim szVehicleProtocol() As TVehicleProtocol
nCount = ArrayLength(m_obase.GetVehicle)
ReDim aszReturn(1 To nCount, 1 To 10)
aszTemp = m_obase.GetVehicle
For i = 1 To nCount
   aszReturn(i, 1) = aszTemp(i, 1)
   aszReturn(i, 2) = aszTemp(i, 2)
Next i

Select Case values
       Case 0  '将所有的车辆取消为该协议
'            WriteProcessBar , i, nCount, "得到对象[" & rsItems!chinese_name & "]"
           vbYesOrNo = MsgBox("是否真的要将默认协议" & "[" & Trim(txtProtocolName.Text) & "]" & "取消吗?", vbQuestion + vbYesNo + vbDefaultButton2, "信息")
           If vbYesOrNo = vbYes Then
               ReDim aszReturn1(1 To nCount)
                For i = 1 To nCount
                   aszReturn1(i) = aszTemp(i, 1)
                Next i
                SetBusy
                m_oProtocol.DelVehicleProtocol aszReturn1, 1
                m_oProtocol.DelVehicleProtocol aszReturn1, 0
                SetNormal
           End If
                       
                
       Case 1  '将所有的车辆指定为该协议
             
            ReDim szVehicleProtocol(1 To nCount)
          vbYesOrNo = MsgBox("是否真的要将" & "[" & Trim(txtProtocolName.Text) & "]" & "设置成为默认的拆算协议吗", vbQuestion + vbYesNo + vbDefaultButton2, "信息")
              If vbYesOrNo = vbYes Then
                     '将随行车辆指定默认的拆算协议
                     SetBusy
                     ShowSBInfo "正在给所有的车辆设置,请稍候......" & "[" & Trim(txtProtocolName.Text) & "]", ESB_WorkingInfo
                    For i = 1 To nCount
'                     WriteProcessBar , i, nCount, "正在给[" & szVehicleProtocol(i).VehicleLicense & "]" & "设置拆算协议,请等待！"
                         szVehicleProtocol(i).ProtocolID = Trim(txtProtocolID.Text)
                         szVehicleProtocol(i).VehicleID = aszReturn(i, 1)
                         szVehicleProtocol(i).VehicleLicense = aszReturn(i, 2)
                         szVehicleProtocol(i).AcceptType = 0
                    Next i
                   m_oProtocol.SetVehicleProtocol szVehicleProtocol
                     '将快件车辆指定默认的拆算协议
                    For i = 1 To nCount
'                     WriteProcessBar , i, nCount, "正在给[" & szVehicleProtocol(i).VehicleLicense & "]" & "设置拆算协议,请等待！"
                         szVehicleProtocol(i).ProtocolID = Trim(txtProtocolID.Text)
                         szVehicleProtocol(i).VehicleID = aszReturn(i, 1)
                         szVehicleProtocol(i).VehicleLicense = aszReturn(i, 2)
                         szVehicleProtocol(i).AcceptType = 1
                    Next i
                    m_oProtocol.SetVehicleProtocol szVehicleProtocol
                    ShowSBInfo ""

                    SetNormal
                Else
                     m_oProtocol.Default = 0
                     SetNormal
              End If
   End Select
End Sub
