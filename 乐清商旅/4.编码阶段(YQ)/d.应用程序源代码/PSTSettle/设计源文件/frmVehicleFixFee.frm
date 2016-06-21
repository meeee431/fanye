VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmVehicleFixFee 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "固定费用"
   ClientHeight    =   5400
   ClientLeft      =   1530
   ClientTop       =   2385
   ClientWidth     =   8955
   Icon            =   "frmVehicleFixFee.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "车辆信息"
      Height          =   615
      Left            =   225
      TabIndex        =   7
      Top             =   540
      Width           =   7125
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "绍汽集团"
         Height          =   180
         Left            =   4095
         TabIndex        =   11
         Top             =   285
         Width           =   720
      End
      Begin VB.Label lblCom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "公司:"
         Height          =   180
         Left            =   3570
         TabIndex        =   10
         Top             =   285
         Width           =   450
      End
      Begin VB.Label lblLicenseTagNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙DB1199"
         Height          =   180
         Left            =   1875
         TabIndex        =   9
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblBus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车牌:"
         Height          =   180
         Left            =   1320
         TabIndex        =   8
         Top             =   270
         Width           =   450
      End
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   390
      Left            =   7560
      TabIndex        =   5
      Top             =   360
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
      MICON           =   "frmVehicleFixFee.frx":000C
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
      Left            =   225
      TabIndex        =   4
      Top             =   1215
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
      Rows            =   50
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
      Left            =   4785
      TabIndex        =   3
      Top             =   165
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   56819712
      CurrentDate     =   38553
   End
   Begin FText.asFlatTextBox txtVehicle 
      Height          =   300
      Left            =   1050
      TabIndex        =   2
      Top             =   165
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
      Left            =   7560
      TabIndex        =   6
      Top             =   1035
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
      MICON           =   "frmVehicleFixFee.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期(&D):"
      Height          =   180
      Left            =   3765
      TabIndex        =   1
      Top             =   225
      Width           =   720
   End
   Begin VB.Label lblCheCi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆(&V):"
      Height          =   180
      Left            =   255
      TabIndex        =   0
      Top             =   225
      Width           =   720
   End
End
Attribute VB_Name = "frmVehicleFixFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_eFormStatus As EFormStatus

Public m_bIsParent As Boolean
Public m_szVehicleID As String '车辆代码
Public m_dtDate As Date '日期
Private m_aszFixFeeItem() As String
Private m_atTemp() As TSplitItemInfo
Private m_oSplit As New Split



Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdok_Click()

    On Error GoTo ErrorHandle
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    '创建记录集
    If txtVehicle.Text = "" Then
        MsgBox "车辆不能为空！", vbExclamation, Me.Caption
        Exit Sub
    End If
    
            With rsTemp.Fields
                .Append "vehicle_id", adChar, 5
                .Append "bus_date", adDBDate
                For i = 1 To g_cnSplitItemCount
                    .Append "split_item_" & i, adDouble
                Next i
                .Append "is_dec", adSmallInt
            End With
            rsTemp.Open
            rsTemp.AddNew
            With vsFixFee
                rsTemp!vehicle_id = txtVehicle.Text
                rsTemp!bus_date = dtpDate.Value
            
                For i = 1 To g_cnSplitItemCount
                    For j = 1 To .Rows - 1
                        If i = Val(.TextMatrix(j, 2)) Then
                            rsTemp.Fields("split_item_" & i) = Val(.TextMatrix(j, 1))
                            Exit For
                        End If
                    Next j
                    If j = .Rows Then rsTemp.Fields("split_item_" & i) = 0
                
                Next i
                
            End With
            rsTemp.Update
            
            
    Select Case m_eFormStatus
    Case AddStatus
            m_oSplit.AddVehicleFixFee rsTemp
        
        
        If m_bIsParent Then
            '刷新父窗口的信息
            frmAllVehicleFixFee.AddList txtVehicle.Text, dtpDate.Value
        End If
        txtVehicle.Text = ""
        lblCompany.Caption = ""
        lblLicenseTagNo.Caption = ""
'        dtpDate.Value = Date
        For i = 1 To vsFixFee.Rows - 1
            vsFixFee.TextMatrix(i, 1) = 0
        Next i
        
    Case ModifyStatus
    
            m_oSplit.EditVehicleFixFee rsTemp
        
        If m_bIsParent Then
            '刷新父窗口的信息
            frmAllVehicleFixFee.UpdateList txtVehicle.Text, dtpDate.Value
        End If
        
        
        Unload Me
    End Select
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    m_aszFixFeeItem = Split(g_szFixFeeItem, ",")
    m_oSplit.Init g_oActiveUser
    
    GetSplitItem
    
    AlignFormPos Me
    FillHead
    lblCompany.Caption = ""
    lblLicenseTagNo.Caption = ""
    
        Me.Caption = "车辆固定费用"
        lblCheCi.Caption = "车辆(&V):"
        Frame1.Caption = "车辆信息"
        lblBus.Caption = "车牌:"
        lblCom.Caption = "公司:"
        If Not g_szIsFixFeeUpdateEachMonth Then
            dtpDate.Visible = False
            lblDate.Visible = False
        End If
    
    Select Case m_eFormStatus
    Case EFormStatus.AddStatus
        txtVehicle.Text = ""
        dtpDate.Value = g_oParam.NowDate
        cmdOk.Caption = "新增(&A)"
        
    Case EFormStatus.ModifyStatus
        
           RefreshInfo
        txtVehicle.Enabled = False
        
    End Select
    
    
End Sub

'填充表格的列头
Private Sub FillHead()
    Dim nRows As Integer
    Dim i As Integer
    Dim j As Integer
    
    nRows = ArrayLength(m_aszFixFeeItem)
    
    With vsFixFee
        .Cols = 3
        .Rows = nRows + 1
        .AllowUserResizing = flexResizeColumns
        
        
        
        '显示列头的值
        
        .TextMatrix(0, 0) = "固定费用名称"
        .TextMatrix(0, 1) = "金额"
        .TextMatrix(0, 2) = "固定费用代码"
        .Row = 0
        .Col = 0
        .CellAlignment = flexAlignCenterCenter
        .Row = 0
        .Col = 1
        .CellAlignment = flexAlignCenterCenter
        
        '用split函数分解出来的数组是从0开始的
        '提取出固定项的项目名称
        For i = 0 To nRows - 1
            For j = 1 To ArrayLength(m_atTemp)
                If Val(m_aszFixFeeItem(i)) = Val(m_atTemp(j).SplitItemID) Then
                    .TextMatrix(i + 1, 0) = m_atTemp(j).SplitItemName
                    .TextMatrix(i + 1, 2) = m_atTemp(j).SplitItemID
                    Exit For
                End If
            Next j
        Next i
        
    End With
    With vsFixFee
        .ColWidth(0) = 3000
        .ColWidth(1) = 2000
        .ColWidth(2) = 0
        
        
    End With
End Sub

'得到所有的结算项
Private Sub GetSplitItem()
    Dim oReport As New Report
    
    oReport.Init g_oActiveUser
    m_atTemp = oReport.GetSplitItemInfo(, True)
    
End Sub

Public Sub RefreshInfo()
    On Error GoTo ErrorHandle
    
    
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nCols As Integer
    Dim oReport As New Report
    Dim k As Integer
    
    On Error GoTo ErrorHandle
    oReport.Init g_oActiveUser
    Set rsTemp = oReport.GetAllVehicleFixFee(m_szVehicleID, "", m_dtDate, DateAdd("d", 1, m_dtDate), "", -1)
    
    nCols = ArrayLength(m_aszFixFeeItem)
    For i = 1 To rsTemp.RecordCount
        With vsFixFee
            txtVehicle.Text = FormatDbValue(rsTemp!vehicle_id) ', FormatDbValue(rsTemp!license_tag_no))
            dtpDate.Value = ToDBDate(FormatDbValue(rsTemp!bus_date))
            lblCompany.Caption = FormatDbValue(rsTemp!transport_company_name)
            lblLicenseTagNo.Caption = FormatDbValue(rsTemp!license_tag_no)
            
            For k = 0 To nCols - 1
                For j = 1 To .Rows - 1
                    If Val(m_aszFixFeeItem(k)) = Val(.TextMatrix(j, 2)) Then
                        .TextMatrix(j, 1) = FormatDbValue(rsTemp.Fields("split_item_" & m_aszFixFeeItem(k)))
                        Exit For
                    End If
                Next j
            Next k
            
            
        End With
        rsTemp.MoveNext
    Next i
    
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me

End Sub



Private Sub txtVehicle_ButtonClick()
    On Error GoTo err
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    Dim oBus As New Bus
    oShell.Init g_oActiveUser

        aszTemp = oShell.SelectVehicleEX
        Set oShell = Nothing
        If ArrayLength(aszTemp) = 0 Then Exit Sub
        txtVehicle.Text = Trim(aszTemp(1, 1))
        lblLicenseTagNo.Caption = Trim(aszTemp(1, 2))
        lblCompany.Caption = Trim(aszTemp(1, 3))
    RefreshVehicleNegativeInfo
    
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
    dbTotalVehicleSettlePrice = Val(oSplit.TotalVehicleSettlePrice(txtVehicle.Text, GetFirstMonthDay(DateAdd("m", -1, dtpDate)), GetLastMonthDay(DateAdd("m", -1, dtpDate))))

    If dbTotalVehicleSettlePrice < 0 Then
        For i = 1 To vsFixFee.Rows - 1
            If Val(vsFixFee.TextMatrix(i, 2)) = Val(g_szSettleNegativeSplitItem) Then
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
    If Not g_bAllowSellteTotalNegative Then
        vsFixFee.Editable = flexEDKbdMouse
    Else
    
        If Val(vsFixFee.TextMatrix(NewRowSel, 2)) = Val(g_szSettleNegativeSplitItem) Then
            vsFixFee.Editable = flexEDNone
        Else
            vsFixFee.Editable = flexEDKbdMouse
        
        End If
    End If
    
    Exit Sub
    
err:
    ShowErrorMsg
    
    
End Sub
