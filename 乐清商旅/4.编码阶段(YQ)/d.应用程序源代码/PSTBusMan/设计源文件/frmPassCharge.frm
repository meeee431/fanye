VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmPassCharge 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "通行费"
   ClientHeight    =   4485
   ClientLeft      =   1785
   ClientTop       =   2190
   ClientWidth     =   7170
   HelpContextID   =   2004201
   Icon            =   "frmPassCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdQuery 
      Height          =   345
      Left            =   5865
      TabIndex        =   6
      Top             =   105
      Width           =   1170
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmPassCharge.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbSeatType 
      Height          =   300
      Left            =   3930
      TabIndex        =   3
      Text            =   "(全部)"
      Top             =   450
      Width           =   1800
   End
   Begin VB.ComboBox CboVehicleeModel 
      Height          =   300
      ItemData        =   "frmPassCharge.frx":0166
      Left            =   885
      List            =   "frmPassCharge.frx":0168
      TabIndex        =   1
      Text            =   "(全部)"
      Top             =   450
      Width           =   1680
   End
   Begin RTComctl3.CoolButton cmdModify 
      Height          =   345
      Left            =   5865
      TabIndex        =   8
      Top             =   975
      Width           =   1170
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "批量修改(&M)"
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
      MICON           =   "frmPassCharge.frx":016A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmStart 
      Interval        =   500
      Left            =   2430
      Top             =   2430
   End
   Begin RTComctl3.CoolButton cmdSave 
      Default         =   -1  'True
      Height          =   345
      Left            =   5865
      TabIndex        =   7
      Top             =   540
      Width           =   1170
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
      MICON           =   "frmPassCharge.frx":0186
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
      Height          =   345
      Left            =   5865
      TabIndex        =   9
      Top             =   1410
      Width           =   1170
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
      MICON           =   "frmPassCharge.frx":01A2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtChange 
      Height          =   240
      Left            =   3240
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   990
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgPass 
      Height          =   3195
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5636
      _Version        =   393216
      Cols            =   5
      BackColorBkg    =   14737632
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位类型(&T):"
      Height          =   180
      Left            =   2805
      TabIndex        =   2
      Top             =   510
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型(&T):"
      Height          =   180
      Left            =   105
      TabIndex        =   0
      Top             =   510
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "通行费设定(&P)"
      Height          =   180
      Left            =   105
      TabIndex        =   4
      Top             =   885
      Width           =   1170
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   975
      TabIndex        =   13
      Top             =   135
      Width           =   90
   End
   Begin VB.Label lblSectionName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   2805
      TabIndex        =   12
      Top             =   135
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "路段名称:"
      Height          =   180
      Left            =   1950
      TabIndex        =   11
      Top             =   135
      Width           =   810
   End
   Begin VB.Label lblSectionID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "路段代码:"
      Height          =   180
      Left            =   105
      TabIndex        =   10
      Top             =   135
      Width           =   810
   End
End
Attribute VB_Name = "frmPassCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmPassCharge.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/27
'* Last Revision Date:2002/09/02
'* Brief Description:通行费管理
'* Relational Document:
'**********************************************************
Public m_szSectionID As String

'Private m_tTransitCharge() As TTransitChargeInfo
Private m_oBaseInfo As New BaseInfo
Private m_oSection As New Section
Private m_aszVehicleModel() As String '车型
Private m_aszSeatType() As String '座位类型


Private Sub cmdModify_Click()
    frmPassRatioEx.szSectionID = m_szSectionID
    frmPassRatioEx.Show vbModal
    cmdQuery_Click
End Sub

Public Sub cmdQuery_Click()
    Dim szaVehicleModel As String
    Dim szSeatType As String
    Dim taTCharge() As TTransitChargeInfo
    Dim tTTransitChargeInfo As TTransitChargeInfo
    On Error GoTo ErrorHandle
    If CboVehicleeModel.Text = "(全部)" Then
        szaVehicleModel = ""
    Else
        szaVehicleModel = m_aszVehicleModel(CboVehicleeModel.ListIndex, 1)
    End If
    If cmbSeatType.Text = "(全部)" Then
        szSeatType = ""
    Else
        szSeatType = ResolveDisplay(cmbSeatType.Text)
    End If
    
    tTTransitChargeInfo.szSection = m_szSectionID
    tTTransitChargeInfo.szSeatType = szSeatType
    tTTransitChargeInfo.szVehicleType = szaVehicleModel
    taTCharge = m_oSection.GetPassChargeEX(tTTransitChargeInfo)
    FullRatio taTCharge
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
Private Sub cmdSave_Click()
    Dim i As Integer
    Dim tChargeRs As TTransitChargeInfo
    Dim nCols As Integer
    On Error GoTo ErrorHandle
    nCols = hfgPass.Cols
    For i = 1 To hfgPass.Rows - 1
        hfgPass.Row = i
        hfgPass.Col = 0
        If hfgPass.CellForeColor = vbBlue Then
            tChargeRs.szVehicleType = ResolveDisplay(hfgPass.TextArray(i * nCols + 1))
            tChargeRs.szPassCharge = Val(hfgPass.TextArray(i * nCols + 3))
            
            tChargeRs.szAnnotation = hfgPass.TextArray(i * nCols + 4)
            tChargeRs.szSeatType = ResolveDisplay(hfgPass.TextArray(i * nCols + 2))
            tChargeRs.szSection = m_szSectionID
            m_oSection.ModifyPassCharge tChargeRs
            hfgPass.Col = 0
            hfgPass.CellForeColor = vbBlack
            hfgPass.Col = 1
            hfgPass.CellForeColor = vbBlack
            hfgPass.Col = 2
            hfgPass.CellForeColor = vbBlack
        End If
    Next
    cmdSave.Enabled = False
    MsgBox "通行费保存成功", vbInformation, "路段"
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    hfgPass.TextArray(0) = "序"
    hfgPass.TextArray(1) = "车型"
    hfgPass.TextArray(2) = "位型"
    hfgPass.TextArray(3) = "通行费"
    hfgPass.TextArray(4) = "备注 "
    hfgPass.ColWidth(0) = 300
    hfgPass.ColWidth(1) = 1200
    hfgPass.ColWidth(2) = 1200
    hfgPass.ColWidth(3) = 700
    hfgPass.ColWidth(4) = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub hfgPass_Click()
    On Error GoTo ErrorHandle
    If hfgPass.Row >= 1 And hfgPass.TextMatrix(hfgPass.Row, 0) <> "" Then
        If hfgPass.Col = 3 Or hfgPass.Col = 4 Then
            txtChange.Visible = True
            
            txtChange.Height = hfgPass.CellHeight
            txtChange.Width = hfgPass.CellWidth
            txtChange.Top = hfgPass.Top + hfgPass.CellTop - 40
            txtChange.Left = hfgPass.Left + hfgPass.CellLeft - 20
            txtChange.Text = hfgPass.Text
            txtChange.SetFocus
        Else
            txtChange.Visible = False
        
        End If
    End If
ErrorHandle:
End Sub

Private Sub hfgPass_Scroll()
    txtChange.Visible = False
End Sub
Private Sub tmStart_Timer()
    Dim nCount As Integer
    Dim i As Integer
    Dim sgTemp() As TTransitChargeInfo
    tmStart.Enabled = False
    m_oSection.Init g_oActiveUser
    m_oSection.Identify m_szSectionID
    lblSection.Caption = m_oSection.SectionID
    lblSectionName.Caption = m_oSection.SectionName
    
    m_oBaseInfo.Init g_oActiveUser
    m_aszVehicleModel = m_oBaseInfo.GetAllVehicleModel
    m_aszSeatType = m_oBaseInfo.GetAllSeatType
    
    nCount = ArrayLength(m_aszSeatType)
    cmbSeatType.AddItem "(全部)"
    For i = 1 To nCount
        cmbSeatType.AddItem MakeDisplayString(m_aszSeatType(i, 1), m_aszSeatType(i, 2))
    Next
    nCount = ArrayLength(m_aszVehicleModel)
    CboVehicleeModel.AddItem "(全部)"
    For i = 1 To nCount
        CboVehicleeModel.AddItem MakeDisplayString(m_aszVehicleModel(i, 1), m_aszVehicleModel(i, 2))
    Next

    cmbSeatType.ListIndex = 0
    CboVehicleeModel.ListIndex = 0
    
    cmdQuery_Click
End Sub
Private Sub txtChange_Change()
    Dim dOldPassCharge As String
    Dim Col As Integer
    Col = hfgPass.Col
    If Col = 4 Then
        dOldPassCharge = hfgPass.Text
        hfgPass.Text = txtChange.Text
    Else
        dOldPassCharge = hfgPass.Text
        hfgPass.Text = Val(txtChange.Text)
    End If
    If dOldPassCharge <> Trim(hfgPass.Text) Then
        hfgPass.Col = 0
        hfgPass.CellForeColor = vbBlue
        hfgPass.Col = 1
        hfgPass.CellForeColor = vbBlue
        hfgPass.Col = 2
        hfgPass.CellForeColor = vbBlue
    End If
    hfgPass.Col = Col
    cmdSave.Enabled = True
End Sub

Private Function GetDateEX(szTypeSeatId As String, aszTemp() As String, Optional bflg As Boolean = False) As String
    Dim i As Integer
    Dim nCount As Integer
    nCount = ArrayLength(aszTemp)
    If bflg = False Then
        GetDateEX = "01[普通]"
        For i = 1 To nCount
            If Trim(aszTemp(i, 1)) = Trim(szTypeSeatId) Then GetDateEX = aszTemp(i, 1) & "[" & aszTemp(i, 2) & "]"
        Next
    Else
        For i = 1 To nCount
            If Trim(aszTemp(i, 1)) = Trim(szTypeSeatId) Then GetDateEX = aszTemp(i, 1) & "[" & aszTemp(i, 2) & "]"
        Next
    
    End If
End Function
Private Function FullRatio(taTCharge() As TTransitChargeInfo)
    Dim nSaveCount As Integer
    Dim i As Integer
    Dim nCount As Integer
    Dim nCols As Integer
    nCount = ArrayLength(taTCharge)
    hfgPass.Clear
    nCols = hfgPass.Cols
    'If chkSave.Value = 0 Then
    '    hfgPass.Clear
    'Else
    '    nSaveCount = hfgPass.Rows - 1
    'End If

    hfgPass.TextArray(0) = "序"
    hfgPass.TextArray(1) = "车型"
    hfgPass.TextArray(2) = "位型"
    hfgPass.TextArray(3) = "通行费"
    hfgPass.TextArray(4) = "备注"
    If nCount = 0 Then Exit Function
    hfgPass.Rows = IIf(nCount + nSaveCount + 1 > 1, nCount + nSaveCount + 1, 2)
    hfgPass.FixedRows = 1
    WriteProcessBar , , nCount

    For i = nSaveCount + 1 To nCount + nSaveCount
        WriteProcessBar , i - nSaveCount, nSaveCount
        hfgPass.TextArray(i * nCols + 0) = i
        hfgPass.TextArray(i * nCols + 1) = GetDateEX(taTCharge(i - nSaveCount).szVehicleType, m_aszVehicleModel, True)
        hfgPass.TextArray(i * nCols + 2) = GetDateEX(taTCharge(i - nSaveCount).szSeatType, m_aszSeatType)
        hfgPass.TextArray(i * nCols + 3) = taTCharge(i - nSaveCount).szPassCharge
        hfgPass.TextArray(i * nCols + 4) = taTCharge(i - nSaveCount).szAnnotation

    Next
    hfgPass.Redraw = True
    hfgPass.FixedCols = 2
    hfgPass.FixedRows = 1

    txtChange.Visible = False
    WriteProcessBar False
End Function

Private Sub txtChange_GotFocus()
    txtChange.SelStart = 0
    txtChange.SelLength = Len(txtChange.Text)
End Sub
