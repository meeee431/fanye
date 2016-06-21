VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusRoute 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车次站点"
   ClientHeight    =   4320
   ClientLeft      =   3420
   ClientTop       =   3570
   ClientWidth     =   6795
   HelpContextID   =   10000780
   Icon            =   "frmBusRoute.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboTCount 
      Height          =   300
      ItemData        =   "frmBusRoute.frx":014A
      Left            =   240
      List            =   "frmBusRoute.frx":0154
      TabIndex        =   3
      Top             =   3795
      Visible         =   0   'False
      Width           =   1170
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4440
      TabIndex        =   9
      Top             =   3840
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
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
      MICON           =   "frmBusRoute.frx":0166
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
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   6
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   7
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblSellStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001[乐清]"
         Height          =   180
         Left            =   2655
         TabIndex        =   11
         Top             =   330
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起点站:"
         Height          =   180
         Left            =   1920
         TabIndex        =   10
         Top             =   330
         Width           =   630
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点列表(&L):"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   1080
      End
   End
   Begin RTComctl3.CoolButton cmdRefresh 
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   3840
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "同步(&R)"
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
      MICON           =   "frmBusRoute.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboTime 
      Height          =   300
      ItemData        =   "frmBusRoute.frx":019E
      Left            =   510
      List            =   "frmBusRoute.frx":01A5
      TabIndex        =   4
      Top             =   4050
      Visible         =   0   'False
      Width           =   1380
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   3840
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
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
      MICON           =   "frmBusRoute.frx":01AF
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
      Default         =   -1  'True
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   3840
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmBusRoute.frx":01CB
      PICN            =   "frmBusRoute.frx":01E7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgRouteStation 
      Height          =   2940
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   5186
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      ScrollBars      =   2
      AllowUserResizing=   3
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Menu pmun_SellYesNo 
      Caption         =   "是否可售"
      Visible         =   0   'False
      Begin VB.Menu pmun_SellYes 
         Caption         =   "可售(不限)"
      End
      Begin VB.Menu pMun_SellNo 
         Caption         =   "不可售"
      End
   End
End
Attribute VB_Name = "frmBusRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oBus As Bus
Private m_tStationInfo() As TBusStationSellInfo
Private m_nStationCount As Integer
Private m_nStarSel As Integer
Private m_nEndSel As Integer
Private m_ny As Integer


Private Sub cboTCount_Change()
    If hfgRouteStation.Text = cboTCount.Text Then Exit Sub
    hfgRouteStation.Text = cboTCount.Text
    hfgRouteStation.CellForeColor = cvChangeColor
    cmdSave.Enabled = True
End Sub

Private Sub cboTCount_Click()
    If hfgRouteStation.Text = cboTCount.Text Then Exit Sub
    hfgRouteStation.Text = cboTCount.Text
    hfgRouteStation.CellForeColor = cvChangeColor
    cmdSave.Enabled = True
End Sub

Private Sub cboTime_Change()
    If hfgRouteStation.Text = cboTime.Text Then Exit Sub
    hfgRouteStation.Text = cboTime.Text
    hfgRouteStation.CellForeColor = cvChangeColor
    cmdSave.Enabled = True
End Sub

Private Sub cboTime_Click()
    If hfgRouteStation.Text = cboTime.Text Then Exit Sub
    hfgRouteStation.Text = cboTime.Text
    hfgRouteStation.CellForeColor = cvChangeColor
    cmdSave.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Private Sub cmdRefresh_Click()
    Dim oRoutePriceTable As New RoutePriceTable
    Dim oTicketPriceMan As New TicketPriceMan
    Dim atVehicleSetTypeInfo() As TBusVehicleSeatType
    Dim aszBusID() As String
    
    On Error GoTo ErrorHandle
    If MsgBox("是否同步车次站点信息", vbYesNo + vbExclamation + vbDefaultButton2, "计划") = vbYes Then
    
        SetBusy
        '同步车次站点
        ShowSBInfo "正在同步车次站点"
        m_oBus.RefreshPassStation
        FillStation
        
        '同步车次票价
        ShowSBInfo "正在同步车次票价"
        '得到该车次计算的条件
        oTicketPriceMan.Init g_oActiveUser
        ReDim aszBusID(1 To 1)
        aszBusID(1) = m_oBus.BusID
        atVehicleSetTypeInfo = oTicketPriceMan.GetAllBusVehicleTypeSeatType(aszBusID)
        
        
        oRoutePriceTable.Init g_oActiveUser
        oRoutePriceTable.Identify g_szExePriceTable
        '进行同步
        oRoutePriceTable.RefreshBusPriceRS atVehicleSetTypeInfo
        
        Set oRoutePriceTable = Nothing
        Set oTicketPriceMan = Nothing
        SetNormal
        ShowSBInfo ""
    End If
    Exit Sub
ErrorHandle:
    SetNormal
    ShowSBInfo ""
    Set oRoutePriceTable = Nothing
    Set oTicketPriceMan = Nothing
    ShowErrorMsg
End Sub

Private Sub cmdSave_Click()
    Dim tStationInfo As TBusStationSellInfo
    Dim i As Integer
    Dim nDateTemp As Integer
On Error GoTo ErrHandle
    SetBusy
    
    For i = 1 To m_nStationCount
        tStationInfo.szStationID = m_tStationInfo(i).szStationID
        tStationInfo.nMileage = m_tStationInfo(i).nMileage
        tStationInfo.sgFullPrice = m_tStationInfo(i).sgFullPrice
        tStationInfo.sgHalfPrice = m_tStationInfo(i).sgHalfPrice
        '填充站点限售张数
        If Trim(hfgRouteStation.TextArray(i * 5 + 3)) = "不限" Then
            tStationInfo.nLimitedSellCount = -1
        End If
        If Trim(hfgRouteStation.TextArray(i * 5 + 3)) = "不可售" Then
            tStationInfo.nLimitedSellCount = 0
        End If
        If Val(hfgRouteStation.TextArray(i * 5 + 3)) > 0 Then
            tStationInfo.nLimitedSellCount = Val(hfgRouteStation.TextArray(i * 5 + 3))
        End If
        '填充站点限售时间
        If Trim(hfgRouteStation.TextArray(i * 5 + 4)) = "不限" Then
            tStationInfo.sgLimitedSellTime = -1
        Else
             
            Dim nLen As Integer
            Dim szlimtTime As String
            nLen = InStr(1, hfgRouteStation.TextArray(i * 5 + 4), "小时")
            If nLen <> 0 Then
               szlimtTime = Left(hfgRouteStation.TextArray(i * 5 + 4), nLen - 1)
            Else
               szlimtTime = hfgRouteStation.TextArray(i * 5 + 4)
            End If
            
            If IsNumeric(szlimtTime) = False Then
               MsgBox "输入有误,应以**.**小时输入。, vbExclamation, Me.Caption"
               Exit Sub
            End If
            szlimtTime = Format(szlimtTime, ".00")
            tStationInfo.sgLimitedSellTime = CSng(szlimtTime)
            'tStationInfo.nLimitedSellTime = FormatDataAndSave(hfgRouteStation.TextArray(i * 5 + 4), m_oBus.BusType)
        End If
        m_oBus.ModifyPassStationSellInfo tStationInfo
    Next
    For i = 1 To m_nStationCount
    hfgRouteStation.Row = i
    hfgRouteStation.Col = 3
    hfgRouteStation.CellForeColor = vbBlack
    hfgRouteStation.Col = 4
    hfgRouteStation.CellForeColor = vbBlack
    Next
cmdSave.Enabled = False
SetNormal
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyEscape
       Unload Me
End Select
End Sub


'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:刷新出该车次的起点站
'===================================================b

Private Sub Form_Load()
    AlignFormPos Me
    FillStation
  
    lblSellStation.Caption = m_oBus.StartStationName
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub hfgRouteStation_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

 If Button = 2 Then
  Me.PopupMenu pmun_SellYesNo
 Else
  m_ny = y
  m_nStarSel = hfgRouteStation.Row
 End If
  

End Sub

'Private Sub hfgRouteStation_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Dim nTemp As Double
' If Button = 1 Then
'    m_nEndSel = CInt(y - m_y) \ (hfgRouteStation.CellHeight)
'    nTemp = (y - m_ny) Mod (hfgRouteStation.CellHeight)
'    If nTemp <> 0 Then
'      m_nEndSel = m_nEndSel + 1
'    End If
' End If
'End Sub

Private Sub hfgRouteStation_Scroll()
    cboTCount.Visible = False
    cboTime.Visible = False
End Sub

Private Sub hfgRouteStation_Click()
Const cnMargin = 15
On Error GoTo ErrHandle
    If hfgRouteStation.Col = 3 Then
        cboTCount.Top = hfgRouteStation.Top + hfgRouteStation.CellTop - cnMargin
        cboTCount.Left = hfgRouteStation.Left + hfgRouteStation.CellLeft
        cboTCount.Visible = True
        cboTime.Visible = False
        cboTCount.Text = hfgRouteStation.Text
        cboTCount.SetFocus
        Exit Sub
    Else
        cboTime.Visible = False
        cboTCount.Visible = False
    End If
    If hfgRouteStation.Col = 4 Then
        cboTime.Top = hfgRouteStation.Top + hfgRouteStation.CellTop - cnMargin
        cboTime.Left = hfgRouteStation.Left + hfgRouteStation.CellLeft
        cboTime.Visible = True
        cboTCount.Visible = False
        cboTime.Text = hfgRouteStation.Text
        cboTime.SetFocus
        Exit Sub
    Else
        cboTime.Visible = False
        cboTCount.Visible = False
    End If
ErrHandle:
End Sub

Public Sub Init(objBus As Bus)
    Set m_oBus = objBus
End Sub

Private Sub FillStation()
Dim i As Integer
On Error GoTo ErrHandle
    ShowSBInfo "获得车次站点属性"
    m_tStationInfo = m_oBus.GetPassStation
    m_nStationCount = ArrayLength(m_tStationInfo)
    hfgRouteStation.Redraw = False
    hfgRouteStation.Rows = m_nStationCount + 1
    hfgRouteStation.TextArray(0) = "站点代码"
    hfgRouteStation.TextArray(1) = "站点名称"
    hfgRouteStation.TextArray(2) = "里程数"
    hfgRouteStation.TextArray(3) = "限售张数"
    hfgRouteStation.TextArray(4) = "限售时间"
    hfgRouteStation.ColWidth(3) = 1200
    hfgRouteStation.ColWidth(4) = 1400
    cboTCount.AddItem "5张"
    cboTCount.AddItem "10张"
    cboTCount.AddItem "15张"
    cboTCount.AddItem "20张"
    cboTCount.AddItem "25张"
    cboTCount.AddItem "30张"
'    cboTime.AddItem "不限"
    
    If m_oBus.BusType = TP_RegularBus Then
      cboTime.AddItem "5小时"
      cboTime.AddItem "10小时"
      cboTime.AddItem "15小时"
      cboTime.AddItem "20小时"
      cboTime.AddItem "25小时"
      cboTime.AddItem "30小时"
    End If
    
    ShowSBInfo "获得车次站点属性"
    
    For i = 1 To m_nStationCount
        ShowSBInfo "获得站点" & m_tStationInfo(i).szStationName
        hfgRouteStation.TextArray(i * 5 + 0) = m_tStationInfo(i).szStationID
        hfgRouteStation.TextArray(i * 5 + 1) = m_tStationInfo(i).szStationName
        hfgRouteStation.TextArray(i * 5 + 2) = m_tStationInfo(i).nMileage
        hfgRouteStation.Row = i
        hfgRouteStation.Col = 3
        Select Case m_tStationInfo(i).nLimitedSellCount
           Case Is < 0: hfgRouteStation.Text = "不限": hfgRouteStation.CellForeColor = vbBlack
           Case 0: hfgRouteStation.Text = "不可售": hfgRouteStation.CellForeColor = vbGrayText
           Case Else: hfgRouteStation.Text = m_tStationInfo(i).nLimitedSellCount
        End Select
        hfgRouteStation.Row = i
        hfgRouteStation.Col = 4
        Select Case m_tStationInfo(i).sgLimitedSellTime
           Case Is <= 0: hfgRouteStation.Text = "不限": hfgRouteStation.CellForeColor = vbGrayText
           Case Else:
           cboTime.AddItem CStr(m_tStationInfo(i).sgLimitedSellTime)
           hfgRouteStation.Text = CStr(m_tStationInfo(i).sgLimitedSellTime)
        End Select
    Next
    ShowSBInfo ""
    hfgRouteStation.Redraw = True
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Private Sub pMun_SellNo_Click()
    Dim i As Integer
    With hfgRouteStation
        For i = .Row To .RowSel
            .TextMatrix(i, 3) = "不可售"
            If cboTCount.Visible = True Then cboTCount.ListIndex = 1
            
        Next i
        
    End With
    cmdSave.Enabled = True


End Sub

Private Sub pmun_SellYes_Click()
 Dim i As Integer
 Dim nCount As Integer
   
    With hfgRouteStation
        For i = .Row To .RowSel
            .TextMatrix(i, 3) = "不限"
            If cboTCount.Visible = True Then cboTCount.ListIndex = 1
            
        Next i
        
    End With
    cmdSave.Enabled = True

End Sub
