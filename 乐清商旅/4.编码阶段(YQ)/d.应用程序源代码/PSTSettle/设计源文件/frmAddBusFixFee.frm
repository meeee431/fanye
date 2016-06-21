VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmAddBusFixFee 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "车次固定费用"
   ClientHeight    =   5550
   ClientLeft      =   240
   ClientTop       =   1935
   ClientWidth     =   8820
   Icon            =   "frmAddBusFixFee.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8820
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "车次信息"
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   510
      Width           =   7125
      Begin VB.Label Lable3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路:"
         Height          =   180
         Left            =   405
         TabIndex        =   4
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "绍兴杭州线"
         Height          =   180
         Left            =   960
         TabIndex        =   3
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Lable4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   3585
         TabIndex        =   2
         Top             =   285
         Width           =   810
      End
      Begin VB.Label lblStartUpTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8:55"
         Height          =   180
         Left            =   4470
         TabIndex        =   1
         Top             =   285
         Width           =   360
      End
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   390
      Left            =   7410
      TabIndex        =   5
      Top             =   330
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
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
      MICON           =   "frmAddBusFixFee.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFixFee 
      Height          =   4125
      Left            =   75
      TabIndex        =   6
      Top             =   1185
      Width           =   7140
      _cx             =   12594
      _cy             =   7276
      _ConvInfo       =   -1
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
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
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Editable        =   0
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   4635
      TabIndex        =   7
      Top             =   135
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   56819712
      CurrentDate     =   38553
   End
   Begin FText.asFlatTextBox txtBusID 
      Height          =   300
      Left            =   900
      TabIndex        =   8
      Top             =   135
      Width           =   2025
      _ExtentX        =   3572
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
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   7410
      TabIndex        =   9
      Top             =   1005
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
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
      MICON           =   "frmAddBusFixFee.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Lable1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次(&V):"
      Height          =   180
      Left            =   105
      TabIndex        =   11
      Top             =   195
      Width           =   720
   End
   Begin VB.Label Lable2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期(&D):"
      Height          =   180
      Left            =   3615
      TabIndex        =   10
      Top             =   195
      Width           =   720
   End
End
Attribute VB_Name = "frmAddBusFixFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_eFormStatus As EFormStatus

Public m_bIsParent As Boolean
Public m_szBusID As String '车次代码
Public m_dtDate As Date '日期
Private m_aszFixFeeItem() As String
Private m_atTemp() As TSplitItemInfo
Private m_oSplit As New Split

Private m_nCompanyCount As Integer '该车次的公司个数
Private m_aszCompany() As String

Const cnCols = 6
Const cnSerial = 0
Const cnCompanyID = 1
Const cnCompanyName = 2
Const cnFixFeeID = 3
Const cnFixFeeName = 4
Const cnValue = 5

Private Sub cmdClose_Click()
    ReDim m_aszCompany(1 To 1, 1 To 2)
    m_aszCompany(1, 1) = ""
    m_aszCompany(1, 2) = ""
    Unload Me
End Sub


Private Sub cmdok_Click()

    On Error GoTo ErrorHandle
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim nCount As Integer
    '创建记录集

    If txtBusID.Text = "" Then
        MsgBox "车次不能为空！", vbExclamation, Me.Caption
        Exit Sub
    End If

    With rsTemp.Fields
        .Append "bus_id", adChar, 5
        .Append "transport_company_id", adChar, 12
        .Append "bus_date", adDBDate
        For i = 1 To g_cnSplitItemCount
            .Append "split_item_" & i, adDouble
        Next i
        .Append "is_dec", adSmallInt
    End With
    
    With vsFixFee
        nCount = Int((.Rows - 1) / ArrayLength(m_aszFixFeeItem))
        
        rsTemp.Open
        For i = 1 To nCount
            rsTemp.AddNew
            
            rsTemp!bus_id = txtBusID.Text
            rsTemp!transport_company_id = .TextMatrix((i - 1) * ArrayLength(m_aszFixFeeItem) + 1, cnCompanyID)
    
            If Not g_szIsFixFeeUpdateEachMonth Then
                rsTemp!bus_date = cszEmptyDateStr
            Else
                rsTemp!bus_date = dtpDate.Value
            End If
            
            For j = 1 To g_cnSplitItemCount
                For k = (i - 1) * ArrayLength(m_aszFixFeeItem) + 1 To i * ArrayLength(m_aszFixFeeItem)
                    If j = Val(.TextMatrix(k, cnFixFeeID)) Then
                        rsTemp.Fields("split_item_" & j) = Val(.TextMatrix(k, cnValue))
                        Exit For
                    End If
                Next k
                If k = .Rows Then rsTemp.Fields("split_item_" & j) = 0
    
            Next j
            
            rsTemp.Update
        Next i
    End With



    Select Case m_eFormStatus
    Case AddStatus

        m_oSplit.AddBusFixFee rsTemp


'        If m_bIsParent Then
'            '刷新父窗口的信息
'            frmAllVehicleFixFee.m_szFormStatus = EFS_Bus
'            frmAllVehicleFixFee.AddList txtBusID.Text, dtpDate.Value
'        End If

        MsgBox "新增车次固定费用成功,车次=" & m_szBusID, vbInformation, Me.Caption

        txtBusID.Text = ""
        
        
        For i = 1 To vsFixFee.Rows - 1
            vsFixFee.TextMatrix(i, 1) = 0
        Next i
        
    Case ModifyStatus

        m_oSplit.EditBusFixFee rsTemp

'
'        If m_bIsParent Then
'            '刷新父窗口的信息
'            frmAllVehicleFixFee.m_szFormStatus = EFS_Bus
'            frmAllVehicleFixFee.UpdateList txtBusID.Text, dtpDate.Value
'        End If
        
        MsgBox "修改车次固定费用成功", vbInformation, Me.Caption
        Unload Me
    End Select

    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()

    txtBusID.Text = m_szBusID
    m_aszFixFeeItem = Split(g_szFixFeeItem, ",")
    m_oSplit.Init g_oActiveUser
    
    GetSplitItem
    
    AlignFormPos Me
    FillHead
    
    Me.Caption = "车次固定费用"
    If Not g_szIsFixFeeUpdateEachMonth Then
        dtpDate.Visible = False
        Lable2.Visible = False
    End If
    
    Select Case m_eFormStatus
    Case EFormStatus.AddStatus
        txtBusID.Text = ""
        dtpDate.Value = g_oParam.NowDate
        cmdOk.Caption = "新增(&A)"
        
    Case EFormStatus.ModifyStatus
    
        RefreshBusFixFee
        txtBusID.Enabled = False
        
    End Select
    
    
End Sub

'填充表格的列头
Private Sub FillHead()
    Dim nRows As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    m_nCompanyCount = ArrayLength(m_aszCompany)
    nRows = ArrayLength(m_aszFixFeeItem) * IIf(m_nCompanyCount = 0, 1, m_nCompanyCount)
    
    With vsFixFee
        .Clear
        .Rows = 1
        .Cols = cnCols
        .Rows = nRows + 1
        .AllowUserResizing = flexResizeColumns
        
        
        
        '显示列头的值
        
        .TextMatrix(0, cnCompanyName) = "公司"
        .TextMatrix(0, cnFixFeeID) = "固定费用代码"
        .TextMatrix(0, cnFixFeeName) = "固定费用名称"
        .TextMatrix(0, cnValue) = "金额"
        .Row = 0
        For i = 0 To cnCols - 1
            .Col = i
            .CellAlignment = flexAlignCenterCenter
        Next i
        
        '用split函数分解出来的数组是从0开始的
        '提取出固定项的项目名称
        For i = 0 To nRows - 1
            For j = 1 To ArrayLength(m_atTemp)
                If Val(m_aszFixFeeItem(i Mod ArrayLength(m_aszFixFeeItem))) = Val(m_atTemp(j).SplitItemID) Then
                    If m_nCompanyCount > 0 Then
                        k = Int((i + 1) / ArrayLength(m_aszFixFeeItem)) + IIf((i + 1) Mod ArrayLength(m_aszFixFeeItem) = 0, 0, 1)
                        .TextMatrix(i + 1, cnCompanyID) = m_aszCompany(k, 1)
                        .TextMatrix(i + 1, cnCompanyName) = m_aszCompany(k, 2)
                        
                    End If
                    .TextMatrix(i + 1, cnFixFeeName) = m_atTemp(j).SplitItemName
                    .TextMatrix(i + 1, cnFixFeeID) = m_atTemp(j).SplitItemID
                    Exit For
                End If
            Next j
        Next i
        
        .ColWidth(cnSerial) = 100
        .ColWidth(cnCompanyID) = 0
        .ColWidth(cnCompanyName) = 2000
        .ColWidth(cnFixFeeName) = 2000
        .ColWidth(cnValue) = 2000
        .ColWidth(cnFixFeeID) = 0 '固定费用代码不可见
        
        .MergeCol(cnCompanyName) = True
        
        
        .AllowUserResizing = flexResizeColumns
        '设置合并
        .MergeCells = flexMergeRestrictColumns
    End With
End Sub

'得到所有的结算项
Private Sub GetSplitItem()
    Dim oReport As New Report
    
    oReport.Init g_oActiveUser
    m_atTemp = oReport.GetSplitItemInfo(, True)
    
End Sub

Private Sub RefreshBusFixFee()
    On Error GoTo ErrorHandle
    
    
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    Dim oReport As New Report
    Dim k As Integer
    Dim nRow As Integer
    On Error GoTo ErrorHandle
    oReport.Init g_oActiveUser
    Set rsTemp = oReport.GetAllBusFixFee(m_szBusID, "", cszEmptyDateStr, cszForeverDateStr, "", -1)
    
    nCount = ArrayLength(m_aszFixFeeItem)
    If rsTemp.RecordCount > 0 Then
        txtBusID.Text = FormatDbValue(rsTemp!bus_id)
        dtpDate.Value = ToDBDate(FormatDbValue(rsTemp!bus_date))
        
        vsFixFee.Rows = nCount * rsTemp.RecordCount + 1
        For i = 1 To rsTemp.RecordCount
            With vsFixFee
    
                For j = 0 To nCount - 1
                    nRow = (i - 1) * nCount + j + 1
                
                
                    .TextMatrix(nRow, cnCompanyID) = FormatDbValue(rsTemp!transport_company_id)
                    .TextMatrix(nRow, cnCompanyName) = FormatDbValue(rsTemp!transport_company_name)
                    
                    
                    For k = 1 To ArrayLength(m_atTemp)
                        If Val(j + 1) = Val(m_atTemp(k).SplitItemID) Then
                            .TextMatrix(nRow, cnFixFeeName) = m_atTemp(k).SplitItemName
                            .TextMatrix(nRow, cnFixFeeID) = m_atTemp(k).SplitItemID
                            Exit For
                        End If
                    Next
                    
                    
                    .TextMatrix(nRow, cnValue) = FormatDbValue(rsTemp.Fields("split_item_" & Val(m_aszFixFeeItem(j))))
                            
                Next j
                
                
            End With
            rsTemp.MoveNext
        Next i
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub RefreshBusInfo()
    On Error GoTo ErrorHandle
    
    
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim j As Integer
    
    Dim oReport As New Report
    
    Dim oBus As New Bus
    
    On Error GoTo ErrorHandle
    oBus.Init g_oActiveUser
    oReport.Init g_oActiveUser
    m_szBusID = txtBusID.Text
    '得到线路及发车时间
    oBus.Identify m_szBusID
    lblRoute.Caption = oBus.RouteName
    lblStartUpTime.Caption = ToDBTime(oBus.StartUpTime)
    
    Set rsTemp = oReport.GetAllBusCompany(m_szBusID)
    If rsTemp.RecordCount > 0 Then
        ReDim m_aszCompany(1 To rsTemp.RecordCount, 1 To 2)
        
        For i = 1 To rsTemp.RecordCount
            m_aszCompany(i, 1) = FormatDbValue(rsTemp!transport_company_id)
            m_aszCompany(i, 2) = FormatDbValue(rsTemp!transport_company_short_name)
            rsTemp.MoveNext
        Next i
        FillHead
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me

End Sub



Private Sub txtBusID_ButtonClick()
    On Error GoTo err
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    
    
    oShell.Init g_oActiveUser

    aszTemp = oShell.SelectBus
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtBusID.Text = Trim(aszTemp(1, 1))
'    m_szBusID = txtBusID.Text
    
    RefreshBusInfo
    
    
    Exit Sub
err:
    ShowErrorMsg
End Sub


'将上个月该车辆负数的应结费用,放在对应的系统参数定义的项目上
Private Sub RefreshVehicleNegativeInfo()
    On Error GoTo err
    Dim oSplit As New Split
    Dim dbTotalVehicleSettlePrice As Double
    Dim i As Integer
    '如果不允许上个月该车辆的负数累加,则不刷新
    SetBusy
    If Not g_bAllowSellteTotalNegative Then Exit Sub
    oSplit.Init g_oActiveUser
    dbTotalVehicleSettlePrice = Val(oSplit.TotalVehicleSettlePrice(txtBusID.Text, GetFirstMonthDay(DateAdd("m", -1, dtpDate)), GetLastMonthDay(DateAdd("m", -1, dtpDate))))

    If dbTotalVehicleSettlePrice < 0 Then
        For i = 1 To vsFixFee.Rows - 1
            If Val(vsFixFee.TextMatrix(i, 3)) = Val(g_szSettleNegativeSplitItem) Then
                vsFixFee.TextMatrix(i, 1) = dbTotalVehicleSettlePrice
                Exit For
            End If
        Next i
    End If
    SetNormal
    Exit Sub
err:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub VsFixFee_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    On Error GoTo err
    '如果不允许上个月该车辆的负数累加,则不刷新
    If NewColSel = cnValue Then
        If Not g_bAllowSellteTotalNegative Then
            vsFixFee.Editable = flexEDKbdMouse
        Else
        
            If Val(vsFixFee.TextMatrix(NewRowSel, 3)) = Val(g_szSettleNegativeSplitItem) Then
                vsFixFee.Editable = flexEDNone
            Else
                vsFixFee.Editable = flexEDKbdMouse
            
            End If
        End If
    Else
        vsFixFee.Editable = flexEDNone
    End If
    Exit Sub
    
err:
    ShowErrorMsg
    
    
End Sub

