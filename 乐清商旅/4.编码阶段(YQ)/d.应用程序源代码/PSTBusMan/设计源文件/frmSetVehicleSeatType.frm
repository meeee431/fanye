VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSetVehicleSeatType 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置座位类型"
   ClientHeight    =   3615
   ClientLeft      =   2895
   ClientTop       =   4905
   ClientWidth     =   7335
   Icon            =   "frmSetVehicleSeatType.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7335
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   5700
      TabIndex        =   9
      Top             =   120
      Width           =   1545
      Begin RTComctl3.CoolButton cmdOk 
         Height          =   330
         Left            =   195
         TabIndex        =   13
         Top             =   1065
         Width           =   1155
         _ExtentX        =   2037
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
         MICON           =   "frmSetVehicleSeatType.frx":014A
         PICN            =   "frmSetVehicleSeatType.frx":0166
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdExit 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   195
         TabIndex        =   12
         Top             =   2865
         Width           =   1155
         _ExtentX        =   2037
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
         MICON           =   "frmSetVehicleSeatType.frx":0500
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
         Left            =   195
         TabIndex        =   11
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "新增一行(&A)"
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
         MICON           =   "frmSetVehicleSeatType.frx":051C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdDelete 
         Height          =   330
         Left            =   195
         TabIndex        =   10
         Top             =   660
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "删除一行(&D)"
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
         MICON           =   "frmSetVehicleSeatType.frx":0538
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
   Begin VSFlex7LCtl.VSFlexGrid vsFg 
      Height          =   2610
      Left            =   120
      TabIndex        =   4
      Top             =   915
      Width           =   5505
      _cx             =   9710
      _cy             =   4604
      _ConvInfo       =   -1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   3
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSetVehicleSeatType.frx":0554
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin FText.asFlatTextBox txtVehicleId 
      Height          =   300
      Left            =   1290
      TabIndex        =   1
      Top             =   120
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   105
      X2              =   5505
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5535
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注:不设置的座位,默认为[普通]"
      Height          =   180
      Left            =   3045
      TabIndex        =   14
      Top             =   195
      Width           =   2520
   End
   Begin VB.Label lblEndSeatNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   180
      Left            =   4500
      TabIndex        =   8
      Top             =   615
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束座号:"
      Height          =   180
      Left            =   3540
      TabIndex        =   7
      Top             =   615
      Width           =   810
   End
   Begin VB.Label lblStartSeatNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   180
      Left            =   2670
      TabIndex        =   6
      Top             =   615
      Width           =   90
   End
   Begin VB.Label lbStartSeat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "起始座号:"
      Height          =   180
      Left            =   1710
      TabIndex        =   5
      Top             =   615
      Width           =   810
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   180
      Left            =   930
      TabIndex        =   3
      Top             =   615
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总座位:"
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   615
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆代码(&V):"
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   1080
   End
End
Attribute VB_Name = "frmSetVehicleSeatType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmSetVehicleType.frm
'* Project Name:PSTBusMan
'* Engineer:陈峰
'* Data Generated:2002/09/12
'* Last Revision Date:2002/09/12
'* Brief Description:车辆位型设置
'* Relational Document:
'**********************************************************
Const cnSerial = 0
Const cnVehicleID = 1
Const cnVehicleModel = 2
Const cnSeatType = 3
Const cnStartSeat = 4
Const cnEndSeat = 5
Public m_nEndSeatNo As Integer
Public m_nStartSeatNo As Integer
Public m_szVehicleId As String '车辆代码

Private m_oBase As New BaseInfo
Private m_aszSeatType() As String '总座位数
Private m_aszOldVehicleSeatType() As TVehcileSeatType '数据库中的车辆座位
Private m_szVehicleModelName As String '车辆
Private m_szLicenseTag  As String '车牌号
Private m_oVehicle As New Vehicle

Private Sub CmdAdd_Click()

    With vsFg
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, cnSerial) = .Row
        .TextMatrix(.Row, cnVehicleID) = m_szVehicleId
        .TextMatrix(.Row, cnVehicleModel) = m_szVehicleModelName
    End With
    cmdOK.Enabled = True
    CmdIsEnable
End Sub


Private Sub CmdDelete_Click()
    With vsFg
        vsFg.RemoveItem .Row
        RefreshSerial
        CmdIsEnable
        cmdOK.Enabled = True
    End With
End Sub


Public Sub cmdExit_Click()
    Set m_oBase = Nothing
    Unload Me
End Sub
    
    
Private Sub cmdOk_Click()
    '保存
    Dim aszVehicleSeatType() As String
    Dim i As Integer
    Dim j As Integer, k As Integer, l As Integer
    Dim UsableRows As Integer
    
    
    With vsFg
        For i = 1 To .Rows - 1
        '判断座位类型有没有选择,起始座位与结束座位是否为数字
            If .TextMatrix(i, cnSeatType) <> "" Then
                If IsNumeric(.TextMatrix(i, cnStartSeat)) And IsNumeric(.TextMatrix(i, cnEndSeat)) Then
                '-------------->应判断结束座位一定要大于起始座位
                    If CInt(.TextMatrix(i, cnStartSeat)) > CInt(.TextMatrix(i, cnEndSeat)) Then
                        MsgBox i & "行起始座位必须小于结束座位", vbExclamation, Me.Caption
                        Exit Sub
                    End If
                    If CInt(.TextMatrix(i, cnStartSeat)) < CInt(m_nStartSeatNo) Then
                        MsgBox i & "行设置的起始座位座位必须大于起始座位号", vbExclamation, Me.Caption
                        Exit Sub
                    End If
                    If CInt(.TextMatrix(i, cnEndSeat)) > CInt(m_nEndSeatNo) Then
                        MsgBox i & "行设置的结束座位必须小于结束座位号", vbExclamation, Me.Caption
                        Exit Sub
                    End If
                    If CInt(.TextMatrix(i, cnEndSeat)) - CInt(.TextMatrix(i, cnStartSeat) + 1) > CInt(lblCount.Caption) Then
                        MsgBox i & "行设置的座位总数必须小于总座位", vbExclamation, Me.Caption
                        Exit Sub
                    End If
                    UsableRows = UsableRows + 1
                Else
                    MsgBox i & "行起始与结束座位必须为数字", vbExclamation, Me.Caption
                    .Row = i
                    Exit Sub
                End If
            Else
                MsgBox "必须设置座位类型", vbExclamation, Me.Caption
                .Row = i
                Exit Sub
            End If
        Next i
        '统计行数
        If UsableRows = 0 Then
            m_oVehicle.DeleteVehicleSeatType
            Exit Sub
        End If
        ReDim Preserve aszVehicleSeatType(1 To UsableRows, 1 To 4)
        UsableRows = 0
        For i = 1 To .Rows - 1
            '判断座位类型有没有选择,起始座位与结束座位是否为数字
            If .TextMatrix(i, cnSeatType) <> "" Then
                If IsNumeric(.TextMatrix(i, cnStartSeat)) And IsNumeric(.TextMatrix(i, cnEndSeat)) Then
                    '-------------->应判断结束座位一定要大于起始座位
                    UsableRows = UsableRows + 1
                    aszVehicleSeatType(UsableRows, 1) = ResolveDisplay(.TextMatrix(UsableRows, cnSeatType))
                    aszVehicleSeatType(UsableRows, 2) = .TextMatrix(UsableRows, cnStartSeat)
                    aszVehicleSeatType(UsableRows, 3) = .TextMatrix(UsableRows, cnEndSeat)
                    aszVehicleSeatType(UsableRows, 4) = i
                End If
            End If
        Next i
    
        
        '判断座位的类型是不是重复
        For i = 1 To UsableRows
            For k = 1 To UsableRows
                If k <> i Then
                    '当起始座位与结束座位均小于或均大于其他行的起始座位与结束座位时,则说明其座位设置是不重复的
                    If Not ((CInt(aszVehicleSeatType(i, 2)) > CInt(aszVehicleSeatType(k, 2)) _
                    And CInt(aszVehicleSeatType(i, 2)) > CInt(aszVehicleSeatType(k, 3)) _
                    And CInt(aszVehicleSeatType(i, 3)) > CInt(aszVehicleSeatType(k, 2)) _
                    And CInt(aszVehicleSeatType(i, 3)) > CInt(aszVehicleSeatType(k, 3))) _
                    Or (CInt(aszVehicleSeatType(i, 2)) < CInt(aszVehicleSeatType(k, 2)) _
                    And CInt(aszVehicleSeatType(i, 2)) < CInt(aszVehicleSeatType(k, 3)) _
                    And CInt(aszVehicleSeatType(i, 3)) < CInt(aszVehicleSeatType(k, 2)) _
                    And CInt(aszVehicleSeatType(i, 3)) < CInt(aszVehicleSeatType(k, 3)))) Then
                    
                        '此座位已被其他类型定义,则出错
                        MsgBox "第" & aszVehicleSeatType(i, 4) & "行中定义的座位的类型与第" & aszVehicleSeatType(k, 4) & "行中定义的冲突,无法识别哪一个为准", vbExclamation, Me.Caption
                        .Row = i
                        
                        Exit Sub
                    End If
                
                End If
            Next k
        Next i
    
    End With
    On Error GoTo ErrorHandle
    m_oVehicle.UpdateVehicleSeatType aszVehicleSeatType
    cmdExit_Click
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    m_oBase.Init g_oActiveUser
    m_oVehicle.Init g_oActiveUser
    If m_szVehicleId <> "" Then
        txtVehicleId.Text = m_szVehicleId
        m_oVehicle.Identify m_szVehicleId
        
        lblCount.Caption = CStr(m_nEndSeatNo - m_nStartSeatNo + 1) 'm_oVehicle.SeatCount
        m_szVehicleModelName = m_oVehicle.VehicleModelName '座位类型
        m_szLicenseTag = m_oVehicle.LicenseTag
    End If
    lblStartSeatNo.Caption = str(m_nStartSeatNo)
    lblEndSeatNo.Caption = str(m_nEndSeatNo)
    FillVehicleSeat
    cmdOK.Enabled = False
End Sub

Private Sub FillVehicleSeat()
    '设置列头
    vsFg.TextMatrix(0, cnSerial) = "序号"
    vsFg.TextMatrix(0, cnVehicleID) = "车辆代号"
    vsFg.TextMatrix(0, cnVehicleModel) = "车型"
    vsFg.TextMatrix(0, cnSeatType) = "座位类型"
    vsFg.TextMatrix(0, cnStartSeat) = "起始座号"
    vsFg.TextMatrix(0, cnEndSeat) = "结束座号"
    vsFg.ColWidth(cnSeatType) = 1300
    m_aszSeatType = m_oBase.GetAllSeatType
    If m_szVehicleId <> "" Then
        m_aszOldVehicleSeatType = m_oBase.GetAllVehicleSeatTypeInfo(m_szVehicleId)
        AddCombList m_aszSeatType
        FillVehicleSeatType m_aszOldVehicleSeatType
    End If
End Sub

'添加网格组合框列表
Private Function AddCombList(aszSeatType() As String)
    Dim i As Integer
    Dim nCount As Integer
    Dim szTemp As String
    nCount = ArrayLength(aszSeatType)
    If nCount = 0 Then Exit Function
    For i = 1 To nCount - 1
        szTemp = szTemp & aszSeatType(i, 1) & "[" & aszSeatType(i, 2) & "]|"
    Next
    szTemp = szTemp & aszSeatType(i, 1) & "[" & aszSeatType(i, 2) & "]"
    
    vsFg.ColComboList(3) = szTemp
End Function

'填充车辆座位类型
Private Function FillVehicleSeatType(aszTemp() As TVehcileSeatType)
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    nCount = ArrayLength(aszTemp)
    If nCount = 0 Then
        Exit Function
    End If
    With vsFg
        .Rows = nCount + 1
        For i = 1 To nCount
            .TextMatrix(i, cnSerial) = i
            .TextMatrix(i, cnVehicleID) = aszTemp(i).szVehcileID
            .TextMatrix(i, cnVehicleModel) = aszTemp(i).szVehcileTypeName
            .TextMatrix(i, cnSeatType) = MakeDisplayString(Trim(aszTemp(i).szSeatTypeID), Trim(aszTemp(i).szSeatTypeName))
            .TextMatrix(i, cnStartSeat) = aszTemp(i).szStartSeatNo
            .TextMatrix(i, cnEndSeat) = aszTemp(i).szEndSeatNo
        Next i
    End With
End Function

Private Sub CmdIsEnable()
'是否删除可用
    With vsFg
        If .Rows = 1 Or .Row = 0 Or .Rows = 0 Then
        CmdDelete.Enabled = False
        Else
        CmdDelete.Enabled = True
        End If
    End With
End Sub
Private Sub RefreshSerial()
    Dim i As Integer
    With vsFg
        For i = 1 To .Rows
            .TextMatrix(.Row, cnSerial) = .Row
        Next i
    End With
End Sub

Private Sub txtVehicleId_ButtonClick()
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    Dim aszTmp() As String
    aszTmp = oShell.SelectVehicleEX(False)
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtVehicleId.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
End Sub


Private Sub vsFg_ChangeEdit()
    cmdOK.Enabled = True
End Sub

Private Sub vsFg_Click()
'查找匹配的座位类型
    Dim j As Integer
    Dim nSeatTypeCount As Integer
    With vsFg
        If .Col = cnSeatType Then
            nSeatTypeCount = ArrayLength(m_aszSeatType)
            
            For j = 0 To nSeatTypeCount - 1
                If Trim(m_aszSeatType(j + 1, 1)) = ResolveDisplay(.TextMatrix(.Row, cnSeatType)) Then
                    .ComboIndex = j
                End If
            Next j
        End If
    End With
End Sub

Private Sub vsFg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsFg
        If Col = cnStartSeat Or Col = cnEndSeat Then
            If Not IsNumeric(.EditText) Then
                MsgBox "必须为数字", vbExclamation, Me.Caption
                Cancel = True
            End If
        End If
    End With
End Sub
