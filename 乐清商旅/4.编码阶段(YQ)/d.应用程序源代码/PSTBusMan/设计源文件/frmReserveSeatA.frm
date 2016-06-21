VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmReserveSeat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "预留座位"
   ClientHeight    =   4170
   ClientLeft      =   4410
   ClientTop       =   2490
   ClientWidth     =   7125
   HelpContextID   =   10000770
   Icon            =   "frmReserveSeatA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   3510
      TabIndex        =   8
      Top             =   3690
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
      MICON           =   "frmReserveSeatA.frx":038A
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
      Left            =   5970
      TabIndex        =   3
      Top             =   3690
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
      MICON           =   "frmReserveSeatA.frx":03A6
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
      Left            =   4740
      TabIndex        =   2
      Top             =   3690
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmReserveSeatA.frx":03C2
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
      Top             =   -30
      Width           =   7185
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   7
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车辆列表(&L):"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   330
         Width           =   1080
      End
   End
   Begin VB.ComboBox cboChange 
      Height          =   300
      ItemData        =   "frmReserveSeatA.frx":03DE
      Left            =   2010
      List            =   "frmReserveSeatA.frx":03E8
      TabIndex        =   5
      Top             =   4245
      Width           =   1050
   End
   Begin VB.TextBox txtChange 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   795
      TabIndex        =   4
      Top             =   4380
      Width           =   900
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgBusVe 
      Height          =   2820
      Left            =   30
      TabIndex        =   1
      Top             =   735
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4974
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   7
      FixedCols       =   3
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      AllowUserResizing=   2
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
End
Attribute VB_Name = "frmReserveSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cszLongReserve = "(长期预留)"
Const cszNoReserve = "(不预留)"

Public m_bIsParent As Boolean


Private m_oBus As Bus '车次对象的引用 Bus
Private m_szaBusVehicle() As String '1车次车辆
Private m_nVehicleCount As Integer '车辆数
Private m_tReserveSeat As TReserveSeatInfo
Private szSeatNoCount() As Integer



Private Sub cboChange_Change()
    Dim nCol As Integer
    Dim szTemp As String
    szTemp = hfgBusVe.Text
    hfgBusVe.Text = Trim(cboChange.Text)
    If Trim(szTemp) <> Trim(hfgBusVe.Text) Then
        nCol = hfgBusVe.Col
        hfgBusVe.Col = 0
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 1
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 2
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 3
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 4
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 5
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 6
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = nCol
        If nCol = 5 Then
            hfgBusVe.Col = 6
            hfgBusVe.Text = Trim(cboChange.Text)
            hfgBusVe.Col = 0
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 1
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 2
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 3
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 4
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 5
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 6
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = nCol
        End If
        If nCol = 5 And cboChange.Text = cszNoReserve Then
            hfgBusVe.Col = 5
            hfgBusVe.Text = Trim(cboChange.Text)
            hfgBusVe.Col = 0
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 1
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 2
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 3
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 4
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 5
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = 6
            hfgBusVe.CellForeColor = cvChangeColor
            hfgBusVe.Col = nCol
        End If
    cmdOk.Enabled = True
    End If
End Sub


Private Sub cboChange_Click()
    cboChange_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
Dim szaReCode() As String
Dim tBusVehicle As TBusVehicleInfo
Dim bDate As Boolean
Dim i As Integer
Dim j As Integer
On Error GoTo ErrHandle
    With hfgBusVe
        For i = 1 To .Rows - 1
            bDate = True
            .Row = i
            If .CellForeColor = vbBlue Then
                m_tReserveSeat.nSerialNo = Val(.TextArray(i * 7 + 0))
                m_tReserveSeat.nStartSeatNo = Val(.TextArray(i * 7 + 3))
'                GetSeatNoCount
                If .TextArray(i * 7 + 3) = 0 Then .TextArray(i * 7 + 3) = 1
                If AssserSeatCount(.TextArray(i * 7 + 3), i, True) = True Then MsgBox "预留开始座位号小于该车辆开始座位号,不能预留", vbInformation, "座位预留": Exit Sub
                If AssserSeatCount(.TextArray(i * 7 + 4), i) = True Then MsgBox "预留结束座号大于该车辆结束座位号,不能预留", vbInformation, "座位预留": Exit Sub
                m_tReserveSeat.nSeatCount = Val(.TextArray(i * 7 + 4)) - Val(.TextArray(i * 7 + 3)) + 1
                If m_tReserveSeat.nSeatCount <= 0 Then MsgBox "预留座位结束座号小于开始座位号,不能预留", vbInformation, "座位预留": Exit Sub
                If m_tReserveSeat.nSeatCount > (szSeatNoCount(i, 2) - szSeatNoCount(i, 1) + 1) Then MsgBox "预留座位数大于总座位数,不能预留", vbInformation, "座位预留": Exit Sub
                If .TextArray(i * 7 + 6) = cszLongReserve Then  '车辆是否长预留
                    If .TextArray(i * 7 + 5) = cszLongReserve Then
                        m_tReserveSeat.dtBeginDate = Now
                    Else
                        m_tReserveSeat.dtBeginDate = CDate(.TextArray(i * 7 + 5))
                    End If
                    m_tReserveSeat.dtEnddate = CDate(cszForeverDateStr)
                    bDate = False
                    m_oBus.ReserveSeat .TextArray(i * 7 + 0), m_tReserveSeat
                End If
                If .TextArray(i * 7 + 5) = cszNoReserve Then
                    m_oBus.UnReserveSeat .TextArray(i * 7 + 0)
                    bDate = False
                End If
                If bDate Then
                    m_tReserveSeat.dtBeginDate = CDate(.TextArray(i * 7 + 5))
                    m_tReserveSeat.dtEnddate = CDate(.TextArray(i * 7 + 6))
                    m_oBus.ReserveSeat .TextArray(i * 7 + 0), m_tReserveSeat
                End If
                For j = 0 To 6
                    .Col = j
                    .CellForeColor = vbBlack
                Next
            End If
        Next
    End With
    MsgBox "车次车辆座位预留成功", vbInformation, "计划"
    cmdOk.Enabled = False
Exit Sub
ErrHandle:
    Select Case err.Number
           Case 13: MsgBox "输入的日期不正确", vbCritical, "错误提示"
           Case Else: ShowErrorMsg
    End Select
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyEscape
       Unload Me
End Select
End Sub

Private Sub Form_Load()
    Dim bDate As Boolean
    Dim i As Integer
    On Error GoTo ErrHandle
    AlignFormPos Me
    
    m_nVehicleCount = ArrayLength(m_szaBusVehicle)
    hfgBusVe.TextArray(0) = "序"
    hfgBusVe.TextArray(1) = "车辆代码"
    hfgBusVe.TextArray(2) = "车辆车牌"
    hfgBusVe.TextArray(3) = "开始座位号"
    hfgBusVe.TextArray(4) = "结束座位号"
    hfgBusVe.TextArray(5) = "开始日期"
    hfgBusVe.TextArray(6) = "结束日期"
    hfgBusVe.Rows = m_nVehicleCount + 1
    hfgBusVe.ColWidth(0) = 400
    hfgBusVe.ColWidth(1) = 750
    hfgBusVe.ColWidth(2) = 1100
    hfgBusVe.ColWidth(3) = 1100
    hfgBusVe.ColWidth(4) = 1100
    hfgBusVe.ColWidth(5) = 1300
    hfgBusVe.ColWidth(6) = 1300
    cboChange.AddItem Format(Now, "YYYY-MM-DD")
    
    For i = 1 To m_nVehicleCount
        bDate = True
        hfgBusVe.TextArray(i * 7 + 0) = Trim(m_szaBusVehicle(i, 2)) '车辆代码
        hfgBusVe.TextArray(i * 7 + 1) = Trim(m_szaBusVehicle(i, 1)) '车辆车牌
        m_tReserveSeat = m_oBus.GetReserverSeat(Val(m_szaBusVehicle(i, 2)))
        hfgBusVe.TextArray(i * 7 + 2) = m_szaBusVehicle(i, 6) '车辆序号
        hfgBusVe.TextArray(i * 7 + 3) = m_tReserveSeat.nStartSeatNo  '开始座位号
        If m_tReserveSeat.nStartSeatNo + m_tReserveSeat.nSeatCount - 1 < 0 Then
            hfgBusVe.TextArray(i * 7 + 4) = 0
        Else
            hfgBusVe.TextArray(i * 7 + 4) = m_tReserveSeat.nStartSeatNo + m_tReserveSeat.nSeatCount - 1 '结束座位号
        End If
        If Format(m_tReserveSeat.dtBeginDate, "YYYY-MM-DD") = cszEmptyDateStr Then
            hfgBusVe.Row = i
            hfgBusVe.TextArray(i * 7 + 5) = cszNoReserve
            hfgBusVe.TextArray(i * 7 + 6) = cszNoReserve
            bDate = False
        End If
        If Format(m_tReserveSeat.dtEnddate, "YYYY-MM-DD") = cszForeverDateStr Then
            hfgBusVe.Row = i
            hfgBusVe.TextArray(i * 7 + 5) = Format(m_tReserveSeat.dtBeginDate, "YYYY-MM-DD")
            hfgBusVe.TextArray(i * 7 + 6) = cszLongReserve
            bDate = False
        End If
        If bDate Then
            hfgBusVe.TextArray(i * 7 + 5) = Format(m_tReserveSeat.dtBeginDate, "YYYY-MM-DD")
            hfgBusVe.TextArray(i * 7 + 6) = Format(m_tReserveSeat.dtEnddate, "YYYY-MM-DD")
        End If
    Next
    GetSeatNoCount
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub Init(vData As Object)
Set m_oBus = vData
m_szaBusVehicle = m_oBus.GetAllVehicle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub hfgBusVe_Click()
    Const cnMargin = 15
    If hfgBusVe.Row = 0 Then Exit Sub
    If hfgBusVe.Col = 3 Or hfgBusVe.Col = 4 Then
        cboChange.Visible = False
        txtChange.Visible = True
        txtChange.Height = hfgBusVe.CellHeight - 2 * cnMargin
        txtChange.Width = hfgBusVe.CellWidth
        txtChange.Top = hfgBusVe.Top + hfgBusVe.CellTop - cnMargin
        txtChange.Left = hfgBusVe.Left + hfgBusVe.CellLeft
        txtChange.Text = hfgBusVe.Text
        txtChange.SetFocus
        Exit Sub
    Else
        txtChange.Visible = False
        cboChange.Visible = False
    End If
    If hfgBusVe.Col = 6 Or hfgBusVe.Col = 5 Then
        If hfgBusVe.Col <> 5 Then
            If hfgBusVe.TextArray(hfgBusVe.Row * 7 + 5) = cszLongReserve Or hfgBusVe.TextArray(hfgBusVe.Row * 7 + 5) = cszNoReserve Then Exit Sub
        End If
        cboChange.Visible = True
        txtChange.Visible = False
        cboChange.Width = hfgBusVe.CellWidth
        cboChange.Top = hfgBusVe.Top + hfgBusVe.CellTop - cnMargin
        cboChange.Left = hfgBusVe.Left + hfgBusVe.CellLeft
        cboChange.Text = hfgBusVe.Text
        cboChange.SetFocus
    Else
        txtChange.Visible = False
        cboChange.Visible = False
    End If
End Sub
Private Sub hfgBusVe_Scroll()
    txtChange.Visible = False
End Sub

Private Sub txtChange_Change()
    Dim nCol As Integer
    Dim szTemp As String
    szTemp = hfgBusVe.Text
    hfgBusVe.Text = Trim(txtChange.Text)
    If Trim(szTemp) <> Trim(hfgBusVe.Text) Then
        nCol = hfgBusVe.Col
        hfgBusVe.Col = 0
        hfgBusVe.CellForeColor = vbBlue
        hfgBusVe.Col = 1
        hfgBusVe.CellForeColor = vbBlue
        hfgBusVe.Col = 2
        hfgBusVe.CellForeColor = vbBlue
        hfgBusVe.Col = 3
        hfgBusVe.CellForeColor = vbBlue
        hfgBusVe.Col = 4
        hfgBusVe.CellForeColor = vbBlue
        hfgBusVe.Col = 5
        hfgBusVe.CellForeColor = vbBlue
        hfgBusVe.Col = 6
        hfgBusVe.CellForeColor = vbBlue
        hfgBusVe.Col = nCol
    cmdOk.Enabled = True
    End If
End Sub

Private Function AssserSeatCount(nSeatNo As String, nI As Integer, Optional bflg As Boolean = False) As Boolean
    If bflg = True Then ' 起始座号
        If CInt(Val(nSeatNo)) < CInt(szSeatNoCount(nI, 1)) Then
            AssserSeatCount = True
        End If
    Else               '结束座号
        If CInt(Val(nSeatNo)) > CInt(szSeatNoCount(nI, 2)) Then
            AssserSeatCount = True
        End If
    End If
End Function

' szSeatNoCount(i, 2)---结束座号，szSeatNoCount(i, 1)--起始座号
Private Function GetSeatNoCount()
    Dim i As Integer
    Dim nCount As String
    If m_bIsParent Then
        nCount = frmArrangeBus.lvVehicle.ListItems.Count
        If nCount = 0 Then Exit Function
        ReDim szSeatNoCount(1 To nCount, 1 To 2)
        For i = 1 To nCount
            szSeatNoCount(i, 2) = CInt(frmArrangeBus.lvVehicle.ListItems(i).ListSubItems(4).Text) + CInt(frmArrangeBus.lvVehicle.ListItems(i).ListSubItems(7).Text) - 1
            szSeatNoCount(i, 1) = CInt(frmArrangeBus.lvVehicle.ListItems(i).ListSubItems(7).Text)
        Next
    End If
End Function
