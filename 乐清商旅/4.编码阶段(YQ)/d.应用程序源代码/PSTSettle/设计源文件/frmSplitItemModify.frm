VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSplitItemModify 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "费用项修改"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmSplitItemModify.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6840
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox ptCaption 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7965
      TabIndex        =   16
      Top             =   0
      Width           =   7965
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请设置折算费用项:"
         Height          =   180
         Left            =   420
         TabIndex        =   17
         Top             =   300
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   15
      Top             =   690
      Width           =   8775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2505
      Left            =   225
      TabIndex        =   4
      Top             =   1440
      Width           =   2745
      Begin VB.TextBox txtQuantity 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1110
         TabIndex        =   14
         Top             =   1980
         Width           =   1425
      End
      Begin VB.TextBox txtStationPrice 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1110
         TabIndex        =   12
         Top             =   1530
         Width           =   1425
      End
      Begin VB.TextBox txtSettlePrice 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1110
         TabIndex        =   10
         Top             =   1110
         Width           =   1425
      End
      Begin VB.TextBox txtProtocol 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1110
         TabIndex        =   8
         Top             =   690
         Width           =   1425
      End
      Begin VB.TextBox txtObject 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1110
         TabIndex        =   6
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "人数:"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1980
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "协议:"
         Height          =   180
         Left            =   210
         TabIndex        =   11
         Top             =   750
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结给车站:"
         Height          =   180
         Left            =   210
         TabIndex        =   9
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应结票款:"
         Height          =   180
         Left            =   210
         TabIndex        =   7
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算对象:"
         Height          =   180
         Left            =   210
         TabIndex        =   5
         Top             =   330
         Width           =   810
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1740
      Left            =   -30
      TabIndex        =   0
      Top             =   4110
      Width           =   9465
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   4635
         TabIndex        =   2
         Top             =   240
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
         MICON           =   "frmSplitItemModify.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdok 
         Default         =   -1  'True
         Height          =   330
         Left            =   3375
         TabIndex        =   1
         Top             =   240
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
         MICON           =   "frmSplitItemModify.frx":0028
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
   Begin VSFlex7LCtl.VSFlexGrid vsSplitItem 
      Height          =   2955
      Left            =   3105
      TabIndex        =   3
      Top             =   1020
      Width           =   3000
      _cx             =   5292
      _cy             =   5212
      _ConvInfo       =   -1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   5
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   5
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSplitItemModify.frx":0044
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "对象的协议及汇总信息:"
      Height          =   330
      Left            =   255
      TabIndex        =   18
      Top             =   1170
      Width           =   2340
   End
End
Attribute VB_Name = "frmSplitItemModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'此窗口内的代码,均为垃圾代码,全部需要重新写过
'2005-07-22陈峰注



Public m_szObject As String
Public m_szProtocol As String
Public m_dbSettlePrice As Double
Public m_dbStationPrice As Double
Public m_nQuantity As Integer

Public m_bIsSave As Boolean

Public m_nType As Integer
Public m_eSettleObject As ESettleObjectType
Public m_szObjectID As String
Public m_szSettleSheetID As String



Dim m_aszSplitItem() As TSplitItemInfo

Private Sub cmdCancel_Click()
    m_bIsSave = False
    Unload Me
End Sub

Private Sub cmdok_Click()
On Error GoTo ErrHandle
    Dim i As Integer
    Dim j As Integer
    Dim atCompany(1 To 1) As TCompnaySettle
    Dim atVehicle(1 To 1) As TVehilceSettle
    Dim atBus(1 To 1) As TBusSettle
    Dim oSplit As New Split
    
    
    For j = 1 To vsSplitItem.Rows
        For i = 1 To ArrayLength(m_aszSplitItem)
        
            If Trim(m_aszSplitItem(i).SplitItemName) = Trim(vsSplitItem.TextMatrix(j - 1, 0)) Then
                Exit For
            End If
        Next i
        m_adbSplitItem(m_aszSplitItem(i).SplitItemID) = CDbl(vsSplitItem.TextMatrix(j - 1, 1))
    Next j
    m_bIsSave = True
    
    If m_nType = 2 Then
        If MsgBox("是否修改这些结算项？", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            oSplit.Init g_oActiveUser
            If m_eSettleObject = CS_SettleByVehicle Then

                atVehicle(1).SettlementSheetID = m_szSettleSheetID
                atVehicle(1).VehicleID = m_szObjectID
                atVehicle(1).SettlePrice = Val(txtSettlePrice.Text) + Val(txtStationPrice.Text)
                atVehicle(1).SettleStationPrice = Val(txtStationPrice.Text)
                
                For i = 1 To ArrayLength(m_aszSplitItem)
                    If Trim(m_aszSplitItem(i).SplitItemName) = Trim(vsSplitItem.TextMatrix(i - 1, 0)) Then
                        atVehicle(1).SplitItem(m_aszSplitItem(i).SplitItemID) = Val(vsSplitItem.TextMatrix(i - 1, 1))
                    End If
                Next i
                oSplit.UpdateVehicleSettleItem atVehicle
                
            ElseIf m_eSettleObject = CS_SettleByTransportCompany Then
                atCompany(1).SettlementSheetID = m_szSettleSheetID
                atCompany(1).CompanyID = m_szObjectID
                atCompany(1).SettlePrice = Val(txtSettlePrice.Text) + Val(txtStationPrice.Text)
                atCompany(1).SettleStationPrice = Val(txtStationPrice.Text)
                
                For i = 1 To ArrayLength(m_aszSplitItem)
                    If Trim(m_aszSplitItem(i).SplitItemName) = Trim(vsSplitItem.TextMatrix(i - 1, 0)) Then
                        atCompany(1).SplitItem(m_aszSplitItem(i).SplitItemID) = Val(vsSplitItem.TextMatrix(i - 1, 1))
                    End If
                Next i
                oSplit.UpdateCompanySettleItem atCompany
                
            ElseIf m_eSettleObject = CS_SettleByBus Then
                atBus(1).SettlementSheetID = m_szSettleSheetID
                atBus(1).BusID = m_szObjectID
                atBus(1).SettlePrice = Val(txtSettlePrice.Text) + Val(txtStationPrice.Text)
                atBus(1).SettleStationPrice = Val(txtStationPrice.Text)
                
                For i = 1 To ArrayLength(m_aszSplitItem)
                     If Trim(m_aszSplitItem(i).SplitItemName) = Trim(vsSplitItem.TextMatrix(i - 1, 0)) Then
                        atBus(1).SplitItem(m_aszSplitItem(i).SplitItemID) = Val(vsSplitItem.TextMatrix(i - 1, 1))
                    End If
                Next i
                oSplit.UpdateBusSettleItem atBus
            End If
            
            
            
        End If
        
    
    
    End If
    
    
    Unload Me
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    FillForm
End Sub

Private Sub FillForm()
On Error GoTo ErrHandle
    Dim i As Integer
    Dim j As Integer
    Dim szSplitResult As TSplitResult
    Dim m_oReport As New Report
    Dim nCount As Integer

    txtObject.Text = m_szObject
    txtProtocol.Text = m_szProtocol
    txtSettlePrice.Text = m_dbSettlePrice
    txtStationPrice.Text = m_dbStationPrice
    txtQuantity.Text = m_nQuantity
    
    
    m_oReport.Init g_oActiveUser

    '取得拆算项
    m_aszSplitItem = m_oReport.GetSplitItemInfo(, True)
    vsSplitItem.ColDataType(1) = flexDTDouble
    
    Select Case m_nType
    Case 1
        With frmWizSplitSettle
             
            If .lvCompany.ListItems.Count > 0 Then
            vsSplitItem.Rows = .lvCompany.SelectedItem.ListSubItems.Count - 5
                For i = 1 To .lvCompany.SelectedItem.ListSubItems.Count - 5
                    vsSplitItem.TextMatrix(i - 1, 0) = .lvCompany.ColumnHeaders.Item(i + 6).Text
                   vsSplitItem.TextMatrix(i - 1, 1) = .lvCompany.SelectedItem.ListSubItems(i + 5).Text
                Next i
            End If
        
            If .lvVehicle.ListItems.Count > 0 Then
            vsSplitItem.Rows = .lvVehicle.SelectedItem.ListSubItems.Count - 5
                For i = 1 To .lvVehicle.SelectedItem.ListSubItems.Count - 5
                    vsSplitItem.TextMatrix(i - 1, 0) = .lvVehicle.ColumnHeaders.Item(i + 6).Text
                    vsSplitItem.TextMatrix(i - 1, 1) = .lvVehicle.SelectedItem.ListSubItems(i + 5).Text
                Next i
            End If
            
            If .lvBus.ListItems.Count > 0 Then
            vsSplitItem.Rows = .lvBus.SelectedItem.ListSubItems.Count - 5
                For i = 1 To .lvBus.SelectedItem.ListSubItems.Count - 5
                    vsSplitItem.TextMatrix(i - 1, 0) = .lvBus.ColumnHeaders.Item(i + 6).Text
                    vsSplitItem.TextMatrix(i - 1, 1) = .lvBus.SelectedItem.ListSubItems(i + 5).Text
                Next i
            End If
        End With
    Case 2
        With frmSettleAttrib
             
            If .lvCompany.ListItems.Count > 0 Then
                vsSplitItem.Rows = .lvCompany.SelectedItem.ListSubItems.Count - 5
                For i = 1 To .lvCompany.SelectedItem.ListSubItems.Count - 5
                    vsSplitItem.TextMatrix(i - 1, 0) = .lvCompany.ColumnHeaders.Item(i + 6).Text
                    vsSplitItem.TextMatrix(i - 1, 1) = .lvCompany.SelectedItem.ListSubItems(i + 5).Text
                Next i
            End If
        
            If .lvVehicle.ListItems.Count > 0 Then
                vsSplitItem.Rows = .lvVehicle.SelectedItem.ListSubItems.Count - 5
                For i = 1 To .lvVehicle.SelectedItem.ListSubItems.Count - 5
                    vsSplitItem.TextMatrix(i - 1, 0) = .lvVehicle.ColumnHeaders.Item(i + 6).Text
                   vsSplitItem.TextMatrix(i - 1, 1) = .lvVehicle.SelectedItem.ListSubItems(i + 5).Text
                Next i
            End If
        
            If .lvBus.ListItems.Count > 0 Then
                vsSplitItem.Rows = .lvBus.SelectedItem.ListSubItems.Count - 5
                For i = 1 To .lvBus.SelectedItem.ListSubItems.Count - 5
                    vsSplitItem.TextMatrix(i - 1, 0) = .lvBus.ColumnHeaders.Item(i + 6).Text
                   vsSplitItem.TextMatrix(i - 1, 1) = .lvBus.SelectedItem.ListSubItems(i + 5).Text
                Next i
            End If
        End With

    End Select
    ReDim m_adbSplitItem(1 To g_cnSplitItemCount)
    vsSplitItem.Row = 1
    
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Unload Me
End Sub



Private Sub vsSplitItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '及时调整两个汇总金额
    Dim dbSettlePrice As Double
    Dim dbSettleStationPrice As Double
    Dim i As Integer
    Dim j As Integer
    dbSettlePrice = 0
    dbSettleStationPrice = 0
    '因为填充时是一一对应的，所以就不进行查找了
'    For i = 1 To vsSplitItem.Rows - 1
        For j = 0 To ArrayLength(m_aszSplitItem) - 1
'            If Trim(m_aszSplitItem(j).SplitItemName) = Trim(vsSplitItem.TextMatrix(i, 0)) Then
                
                If m_aszSplitItem(j + 1).SplitType = CS_SplitOtherCompany Then
                    dbSettlePrice = dbSettlePrice + Val(vsSplitItem.TextMatrix(j, 1))
                ElseIf m_aszSplitItem(j + 1).SplitType = CS_SplitStation Then
                    dbSettleStationPrice = dbSettleStationPrice + Val(vsSplitItem.TextMatrix(j, 1))
                End If
'                Exit For
'            End If
        Next j
'    Next i
    txtSettlePrice.Text = FormatMoney(dbSettlePrice - dbSettleStationPrice)
    txtStationPrice.Text = FormatMoney(dbSettleStationPrice)
    
    
End Sub

Private Sub vsSplitItem_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    If m_aszSplitItem(NewRowSel + 1).AllowModify = CS_SplitItemAllowModify Then
        vsSplitItem.Editable = flexEDKbdMouse
    Else
        vsSplitItem.Editable = flexEDNone
    End If
End Sub

