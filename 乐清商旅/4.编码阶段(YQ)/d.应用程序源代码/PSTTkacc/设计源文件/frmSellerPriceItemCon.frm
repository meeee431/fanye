VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSellerPriceItemCon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   240
   ClientTop       =   1935
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7200
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   7665
      TabIndex        =   7
      Top             =   0
      Width           =   7665
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   8
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -90
      TabIndex        =   6
      Top             =   690
      Width           =   7725
   End
   Begin VB.Frame Frame1 
      Caption         =   "报表说明"
      Height          =   555
      Left            =   1170
      TabIndex        =   4
      Top             =   6210
      Width           =   6975
      Begin VB.Label Label3 
         Caption         =   "按票种指定时间段，统计票种人数、金额。用于统计售票员的售票情况。"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6435
      End
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1410
      Width           =   4515
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   2580
      TabIndex        =   0
      Top             =   4620
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmSellerPriceItemCon.frx":0000
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
      Left            =   3960
      TabIndex        =   2
      Top             =   4620
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定"
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
      MICON           =   "frmSellerPriceItemCon.frx":001C
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
      Left            =   5370
      TabIndex        =   3
      Top             =   4620
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmSellerPriceItemCon.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   4560
      TabIndex        =   9
      Top             =   930
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   930
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin PSTTKAcc.AddDel adUser 
      Height          =   2535
      Left            =   540
      TabIndex        =   11
      Top             =   1785
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   4471
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonWidth     =   1215
      ButtonHeight    =   315
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -150
      TabIndex        =   12
      Top             =   4380
      Width           =   8745
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3420
      TabIndex        =   15
      Top             =   990
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   600
      TabIndex        =   14
      Top             =   990
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   600
      TabIndex        =   13
      Top             =   1470
      Width           =   900
   End
End
Attribute VB_Name = "frmSellerPriceItemCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IConditionForm

Const cszFileName = "售票员票价项简报模板.xls"

Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Private Sub adUser_DataChange()
    EnableOK
End Sub

Private Sub cboSellStation_Change()
   FillSellerEx
End Sub
Public Sub FillSellerEx()
    Dim oUnit As New Unit
    Dim aszUser() As String
    Dim aszUser2() As String
    Dim aszUser22() As String
    
    Dim oUser As New User
    Dim i As Integer, nUserCount As Integer
    Dim szRecentSeller As String
    Dim szTemp As String
    
    Dim nNoSelected As Integer, nSelected As Integer
    
    oUnit.Init m_oActiveUser
    oUnit.Identify m_oParam.UnitID
    aszUser = oUnit.GetAllUserEX(, ResolveDisplay(cboSellStation))
    nUserCount = ArrayLength(aszUser)
    If nUserCount > 0 Then
        
'        oUser.Init m_oActiveUser
        szRecentSeller = GetRecentSeller()
        
        nNoSelected = 0
        nSelected = 0
        
        For i = 1 To nUserCount
'            oUser.Identify aszUser(i)
            szTemp = MakeDisplayString(aszUser(i, 1), aszUser(i, 2))
            If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
                nNoSelected = nNoSelected + 1
                ReDim Preserve aszUser2(1 To nNoSelected)
                aszUser2(nNoSelected) = szTemp
            Else
                nSelected = nSelected + 1
                ReDim Preserve aszUser22(1 To nSelected)
                aszUser22(nSelected) = szTemp
            End If
        Next
    End If
    adUser.LeftData = aszUser2
    adUser.RightData = aszUser22
End Sub
Private Sub cboSellStation_Click()
    cboSellStation_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    Dim oSellerStat As New TicketSellerDim
    Dim aszUserID() As String
    Dim nSelUserCount As Integer
    Dim i As Integer
    
    On Error GoTo Error_Handle
    '生成Recordset
    nSelUserCount = ArrayLength(adUser.RightData)
'    dtpBeginDate.Value = CDate(Year(dtpBeginDate.Value) & "-" & Month(dtpBeginDate.Value) & "-01")
'    dtpEndDate.Value = DateAdd("D", -1, DateAdd("M", 1, dtpBeginDate.Value))
    If nSelUserCount > 0 Then
        oSellerStat.Init m_oActiveUser
        ReDim aszUserID(1 To nSelUserCount)
        For i = 1 To nSelUserCount
            aszUserID(i) = ResolveDisplay(adUser.RightData(i))
        Next
        Set m_rsData = oSellerStat.SellerPriceItemDateStat(aszUserID, dtpBeginDate.Value, dtpEndDate.Value)
    End If
    ReDim m_vaCustomData(1 To 3, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "制表人"
    m_vaCustomData(3, 2) = m_oActiveUser.UserID
    SaveRecentSeller adUser.RightData
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    m_bOk = False
    AlignFormPos Me

'    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
'    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    FillSellStation cboSellStation
    FillSellerEx
    
    EnableOK
End Sub

'填充售票员
Private Sub FillSeller()
    Dim oSysMan As New SystemMan
    Dim auiUserInfo() As TUserInfo
    Dim i As Integer, nUserCount As Integer
    Dim aszTemp() As String, aszTemp2() As String
    Dim nNoSelected As Integer, nSelected As Integer
    Dim szTemp As String
    Dim szRecentSeller As String
    
    oSysMan.Init m_oActiveUser
    auiUserInfo = oSysMan.GetAllUser()
    nUserCount = ArrayLength(auiUserInfo)
    If nUserCount > 0 Then
        szRecentSeller = GetRecentSeller()
        nNoSelected = 0
        nSelected = 0
        For i = 1 To nUserCount
            szTemp = MakeDisplayString(auiUserInfo(i).UserID, auiUserInfo(i).UserName)
            If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
                nNoSelected = nNoSelected + 1
                ReDim Preserve aszTemp(1 To nNoSelected)
                aszTemp(nNoSelected) = szTemp
            Else
                nSelected = nSelected + 1
                ReDim Preserve aszTemp2(1 To nSelected)
                aszTemp2(nSelected) = szTemp
            End If
        Next
    End If
    adUser.LeftData = aszTemp
    adUser.RightData = aszTemp2
End Sub

Private Sub EnableOK()
    Dim nCount As Integer
    nCount = ArrayLength(adUser.RightData)
    cmdOk.Enabled = IIf(nCount > 0, True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property

'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub




