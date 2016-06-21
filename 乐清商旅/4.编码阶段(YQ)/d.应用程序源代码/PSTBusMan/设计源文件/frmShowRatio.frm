VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmShowRatio 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "线路费率"
   ClientHeight    =   3945
   ClientLeft      =   1215
   ClientTop       =   2490
   ClientWidth     =   9180
   HelpContextID   =   10000560
   Icon            =   "frmShowRatio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9180
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   7545
      TabIndex        =   17
      Top             =   0
      Width           =   1515
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   150
         TabIndex        =   13
         Top             =   2010
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         MICON           =   "frmShowRatio.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   330
         Left            =   150
         TabIndex        =   10
         ToolTipText     =   "编辑检票口信息"
         Top             =   195
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "查询(&Q)"
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
         MICON           =   "frmShowRatio.frx":0166
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdAdd 
         Height          =   330
         Left            =   150
         TabIndex        =   11
         Top             =   585
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "新增(&A)"
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
         MICON           =   "frmShowRatio.frx":0182
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   330
         Left            =   150
         TabIndex        =   14
         Top             =   2400
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "帮助(&H)"
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
         MICON           =   "frmShowRatio.frx":019E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdSave 
         Height          =   330
         Left            =   150
         TabIndex        =   12
         Top             =   990
         Width           =   1200
         _ExtentX        =   2117
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
         MICON           =   "frmShowRatio.frx":01BA
         PICN            =   "frmShowRatio.frx":01D6
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
   Begin VB.ComboBox cboSeatType 
      Height          =   300
      Left            =   5535
      TabIndex        =   7
      Text            =   "(全部)"
      Top             =   390
      Width           =   1485
   End
   Begin VB.CheckBox chkSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "保留上次查询结果(&K)"
      Height          =   210
      Left            =   180
      TabIndex        =   16
      Top             =   3975
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.ComboBox cboRoadLevel 
      Height          =   300
      Left            =   3750
      TabIndex        =   5
      Text            =   "(全部)"
      Top             =   390
      Width           =   1485
   End
   Begin VB.ComboBox cboVehicleModel 
      Height          =   300
      Left            =   1965
      TabIndex        =   3
      Text            =   "(全部)"
      Top             =   390
      Width           =   1485
   End
   Begin VB.ComboBox cboArea 
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Text            =   "(全部)"
      Top             =   390
      Width           =   1485
   End
   Begin VB.TextBox txtChange 
      Height          =   285
      Left            =   3330
      TabIndex        =   15
      Top             =   1965
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgPass 
      Height          =   2670
      Left            =   180
      TabIndex        =   9
      Top             =   1170
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   4710
      _Version        =   393216
      Cols            =   10
      FixedCols       =   4
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      WordWrap        =   -1  'True
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位类型(&T)"
      Height          =   180
      Left            =   5535
      TabIndex        =   6
      Top             =   135
      Width           =   990
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1365
      X2              =   7260
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1380
      X2              =   7275
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公路等级(&R):"
      Height          =   180
      Left            =   3750
      TabIndex        =   4
      Top             =   135
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型(&M):"
      Height          =   180
      Left            =   1965
      TabIndex        =   2
      Top             =   135
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地区(&D):"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "费率列表(&L):"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   885
      Width           =   1320
   End
End
Attribute VB_Name = "frmShowRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmShowRatio.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/27
'* Last Revision Date:2002/09/03
'* Brief Description:
'* Relational Document:
'**********************************************************
Const cnNormalColor = vbBlack '正常的颜色
Const cnChangedColor = vbBlue '改变后的颜色

Private m_oBaseInfo As New BaseInfo
Private m_oCharge As New ChargeRatio

Private m_aszArea() As String
Private m_aszVehicleModel() As String
Private m_aszRoadlevel() As String
Private m_aszSeatType() As String



Private Sub cmdAdd_Click()
    frmChangeRatio.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub
Private Sub cmdQuery_Click()
    Dim szArea As String
    Dim szVehicleModel As String
    Dim szRoadLevel As String
    Dim szSeatType As String
    Dim atCharge() As TChargeRatioEx
    On Error GoTo ErrorHandle
    If cboArea.Text = "(全部)" Then
        szArea = ""
    Else
        szArea = m_aszArea(cboArea.ListIndex, 1)
    End If
    If cboRoadLevel.Text = "(全部)" Then
        szRoadLevel = ""
    Else
        szRoadLevel = m_aszRoadlevel(cboRoadLevel.ListIndex, 1)
    End If
    If cboVehicleModel.Text = "(全部)" Then
        szVehicleModel = ""
    Else
        szVehicleModel = m_aszVehicleModel(cboVehicleModel.ListIndex, 1)
    End If
    If cboSeatType.Text = "(全部)" Then
        szSeatType = ""
    Else
        szSeatType = ResolveDisplay(cboSeatType.Text)
    End If
    SetBusy
    atCharge = m_oCharge.GetAllChargeRatio(szVehicleModel, szArea, szRoadLevel, szSeatType)
    FillRatio atCharge
    SetNormal
    Exit Sub
ErrorHandle:
ShowErrorMsg
End Sub

Private Sub cmdSave_Click()
    Dim tChargeRs As TChargeRatio
    Dim nCols As Integer
    Dim i As Integer
    On Error GoTo ErrorHandle
    nCols = hfgPass.Cols
    For i = 1 To hfgPass.Rows - 1
        hfgPass.Row = i
        hfgPass.Col = 0
        If hfgPass.CellForeColor = cnChangedColor Then
            tChargeRs.szAreaCode = GetData(hfgPass.TextArray(i * nCols + 0), m_aszArea, True)
            tChargeRs.szRoadLevel = GetData(hfgPass.TextArray(i * nCols + 1), m_aszRoadlevel, True)
            tChargeRs.szVehicleModel = GetData(hfgPass.TextArray(i * nCols + 2), m_aszVehicleModel, True)
            tChargeRs.sgBaseCarriageRatio = Val(hfgPass.TextArray(i * nCols + 4))
            tChargeRs.sgRoadConstructFundRatio = Val(hfgPass.TextArray(i * nCols + 5))
            tChargeRs.szAnnotation = hfgPass.TextArray(i * nCols + 6)
            tChargeRs.szSeatType = ResolveDisplay(hfgPass.TextArray(i * nCols + 3))
            '保存到数据库
            m_oCharge.ModifyChargeRatio tChargeRs
            '回复原来的颜色
            hfgPass.Col = 0
            hfgPass.CellForeColor = cnNormalColor
            hfgPass.Col = 1
            hfgPass.CellForeColor = cnNormalColor
            hfgPass.Col = 2
            hfgPass.CellForeColor = cnNormalColor
            hfgPass.Col = 3
            hfgPass.CellForeColor = cnNormalColor
            hfgPass.Col = 4
            hfgPass.CellForeColor = cnNormalColor
            hfgPass.Col = 5
            hfgPass.CellForeColor = cnNormalColor
        End If
    Next
    cmdSave.Enabled = False
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    Dim i As Integer, nCount As Integer
    On Error GoTo ErrorHandle
    m_oCharge.Init g_oActiveUser
    m_oBaseInfo.Init g_oActiveUser
    
    '填充地区
    m_aszArea = m_oBaseInfo.GetAllArea
    nCount = ArrayLength(m_aszArea)
    cboArea.AddItem "(全部)"
    For i = 1 To nCount
        cboArea.AddItem m_aszArea(i, 2)
    Next
    
    '填充车型
    cboVehicleModel.AddItem "(全部)"
    m_aszVehicleModel = m_oBaseInfo.GetAllVehicleModel
    nCount = ArrayLength(m_aszVehicleModel)
    For i = 1 To nCount
        cboVehicleModel.AddItem m_aszVehicleModel(i, 2)
    Next
    '填充公路等级
    cboRoadLevel.AddItem "(全部)"
    m_aszRoadlevel = m_oBaseInfo.GetAllRoadLevel
    nCount = ArrayLength(m_aszRoadlevel)
    For i = 1 To nCount
        cboRoadLevel.AddItem m_aszRoadlevel(i, 2)
    Next
    '填充座位类型
    cboSeatType.AddItem "(全部)"
    m_aszSeatType = m_oBaseInfo.GetAllSeatType
    nCount = ArrayLength(m_aszSeatType)
    For i = 1 To nCount
        cboSeatType.AddItem MakeDisplayString(m_aszSeatType(i, 1), m_aszSeatType(i, 2))
    Next
    hfgPass.Redraw = True
    hfgPass.ColWidth(0) = 700
    hfgPass.ColWidth(1) = 1400
    hfgPass.ColWidth(2) = 800
    hfgPass.ColWidth(3) = 900
    hfgPass.ColWidth(4) = 1100
    hfgPass.ColWidth(5) = 900
    hfgPass.ColWidth(6) = 1000
    
'    hfgPass.RowHeight(0) = hfgPass.RowHeight(0) * 2
    hfgPass.TextArray(0) = "地区"
    hfgPass.TextArray(2) = "车型"
    hfgPass.TextArray(1) = "公路等级"
    hfgPass.TextArray(3) = "位型"
    hfgPass.TextArray(4) = "基本运价率"
    hfgPass.TextArray(5) = "公建金率"
    hfgPass.TextArray(6) = "备注"
    hfgPass.Cols = 7
    hfgPass.Rows = 2
    'hfgPass.FixedRows = 1
    cmdSave.Enabled = False
    Exit Sub
ErrorHandle:
        ShowErrorMsg
End Sub

Private Sub hfgPass_Click()
    Dim cbo As ComboBox
    On Error GoTo ErrorHandle
    If hfgPass.Row = 0 Then Exit Sub
    Select Case hfgPass.Col
    Case 4, 5, 6
        If hfgPass.TextMatrix(hfgPass.Row, 0) = "" Then Exit Sub
        txtChange.Visible = True
        txtChange.Height = hfgPass.CellHeight
        txtChange.Width = hfgPass.CellWidth
        txtChange.Top = hfgPass.Top + hfgPass.CellTop - 40
        txtChange.Left = hfgPass.Left + hfgPass.CellLeft - 20
        txtChange.Text = hfgPass.Text
        txtChange.SetFocus
    Case Else
    txtChange.Visible = False
    End Select
ErrorHandle:
 End Sub
Private Sub hfgPass_Scroll()
    txtChange.Visible = False
End Sub

Private Sub txtChange_Change()
    Dim nCol As Integer
    Dim szTemp As String
    Dim Col As Integer
    nCol = hfgPass.Col
    Select Case hfgPass.Col
    Case 4, 5
        szTemp = hfgPass.Text
        hfgPass.Text = Val(txtChange.Text)
    Case 6
        szTemp = hfgPass.Text
        hfgPass.Text = txtChange.Text
    End Select
    If Trim(szTemp) <> Trim(hfgPass.Text) Then
    nCol = hfgPass.Col
    hfgPass.Col = 0
    hfgPass.CellForeColor = cnChangedColor
    hfgPass.Col = 1
    hfgPass.CellForeColor = cnChangedColor
    hfgPass.Col = 2
    hfgPass.CellForeColor = cnChangedColor
    hfgPass.Col = 3
    hfgPass.CellForeColor = cnChangedColor
    hfgPass.Col = 4
    hfgPass.CellForeColor = cnChangedColor
    hfgPass.Col = nCol
    cmdSave.Enabled = True
    End If
End Sub


Public Sub FillRatio(atCharge() As TChargeRatioEx)
    '将费率进行填充
    Dim nCount As Integer
    Dim nSaveCount As Integer
    Dim i As Long
    Dim nCols As Integer
    nCount = ArrayLength(atCharge)
    hfgPass.Clear
    nCols = hfgPass.Cols
'    If chkSave.Value = 0 Then
'        hfgPass.Clear
'    Else
'        nSaveCount = hfgPass.Rows - 1
'    End If
    hfgPass.TextArray(0) = "地区"
    hfgPass.TextArray(1) = "公路等级"
    hfgPass.TextArray(2) = "车型"
    hfgPass.TextArray(3) = "位型"
    hfgPass.TextArray(4) = "基本运价率"
    hfgPass.TextArray(5) = "公建金率"
    hfgPass.TextArray(6) = "备注"
    If nCount = 0 Then Exit Sub
    hfgPass.Rows = nCount + nSaveCount + 1
    WriteProcessBar , , nCount, "正在填充费率"
    For i = nSaveCount + 1 To nCount + nSaveCount
        WriteProcessBar , i - nSaveCount, nCount, "正在填充费率"
        hfgPass.TextArray(i * nCols + 0) = atCharge(i - nSaveCount).szAreaName
        hfgPass.TextArray(i * nCols + 1) = atCharge(i - nSaveCount).szRoadLevelName
        hfgPass.TextArray(i * nCols + 2) = atCharge(i - nSaveCount).szVehicleModelName
        hfgPass.TextArray(i * nCols + 3) = MakeDisplayString(atCharge(i - nSaveCount).szSeatType, GetData(atCharge(i - nSaveCount).szSeatType, m_aszSeatType))
        hfgPass.TextArray(i * nCols + 4) = atCharge(i - nSaveCount).sgBaseCarriageRatio
        hfgPass.TextArray(i * nCols + 5) = atCharge(i - nSaveCount).sgRoadConstructFundRatio
        hfgPass.TextArray(i * nCols + 6) = atCharge(i - nSaveCount).szAnnotation
    Next
    hfgPass.Redraw = True
    hfgPass.FixedRows = 1
    txtChange.Visible = False
    WriteProcessBar False
End Sub
Private Function GetData(szSeatTypeID As String, aszTemp() As String, Optional bflg As Boolean = False) As String
    Dim i As Integer
    Dim nCount As Integer
    nCount = ArrayLength(aszTemp)
    If bflg = False Then
        GetData = "普通"
        For i = 1 To nCount
            If Trim(aszTemp(i, 1)) = Trim(szSeatTypeID) Then GetData = aszTemp(i, 2): Exit Function
        Next
    Else
        For i = 1 To nCount
            If Trim(aszTemp(i, 2)) = Trim(szSeatTypeID) Then GetData = aszTemp(i, 1): Exit Function
        Next
    End If
End Function

