VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmFinSheetAttrib 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "行包结算单属性"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   HelpContextID   =   7000340
   Icon            =   "frmFinSheetAttrib.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7635
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   36
      Top             =   690
      Width           =   7695
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   3120
      Left            =   -30
      TabIndex        =   30
      Top             =   5400
      Width           =   8745
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   570
         TabIndex        =   35
         Top             =   300
         Visible         =   0   'False
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
         MICON           =   "frmFinSheetAttrib.frx":030A
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
         Height          =   330
         Left            =   6030
         TabIndex        =   33
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "关闭(&E)"
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
         MICON           =   "frmFinSheetAttrib.frx":0326
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
         Height          =   330
         Left            =   4740
         TabIndex        =   32
         Top             =   255
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
         MICON           =   "frmFinSheetAttrib.frx":0342
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   870
      Left            =   540
      TabIndex        =   16
      Top             =   810
      Width           =   6645
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmFinSheetAttrib.frx":035E
         Top             =   210
         Width           =   480
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算单代码:"
         Height          =   180
         Index           =   0
         Left            =   810
         TabIndex        =   29
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lblSheetID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "U000001"
         Height          =   195
         Left            =   1845
         TabIndex        =   28
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐公司:"
         Height          =   180
         Left            =   2940
         TabIndex        =   27
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblDesCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙江快客公司"
         Height          =   180
         Left            =   3780
         TabIndex        =   26
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算日期:"
         Height          =   180
         Left            =   2970
         TabIndex        =   25
         Top             =   540
         Width           =   810
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Index           =   0
         Left            =   810
         TabIndex        =   24
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblStateChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算月份:"
         Height          =   180
         Left            =   810
         TabIndex        =   23
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙江快客"
         Height          =   180
         Left            =   1650
         TabIndex        =   22
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblFinMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002年12月"
         Height          =   180
         Left            =   1680
         TabIndex        =   21
         Top             =   540
         Width           =   900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   810
         X2              =   6090
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         Index           =   0
         X1              =   810
         X2              =   6090
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label lblAcceptType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "快件行包"
         Height          =   180
         Left            =   3810
         TabIndex        =   20
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运方式:"
         Height          =   180
         Left            =   2970
         TabIndex        =   19
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblFinDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002-12-1至2002-12-31"
         Height          =   180
         Left            =   3810
         TabIndex        =   18
         Top             =   540
         Width           =   1890
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "作废"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4920
         TabIndex        =   17
         Top             =   270
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   3645
      Left            =   540
      TabIndex        =   2
      Top             =   1710
      Width           =   6645
      Begin MSComctlLib.ListView vsDetailPrice 
         Height          =   1755
         Left            =   795
         TabIndex        =   34
         Top             =   540
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   3096
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker dtpOperDate 
         Height          =   315
         Left            =   4485
         TabIndex        =   31
         Top             =   2850
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   61472771
         UpDown          =   -1  'True
         CurrentDate     =   37662
      End
      Begin VB.TextBox txtActSplitMoney 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1740
         TabIndex        =   5
         Top             =   2490
         Width           =   1275
      End
      Begin VB.TextBox txtOperator 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1740
         TabIndex        =   4
         Top             =   2850
         Width           =   1275
      End
      Begin VB.TextBox txtRemark 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1740
         TabIndex        =   3
         Top             =   3210
         Width           =   4365
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   180
         Picture         =   "frmFinSheetAttrib.frx":0668
         Top             =   285
         Width           =   480
      End
      Begin VB.Label lblProtocol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5：5"
         Height          =   180
         Left            =   3330
         TabIndex        =   15
         Top             =   270
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算协议:"
         Height          =   180
         Left            =   2340
         TabIndex        =   14
         Top             =   270
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实拆出(&A):"
         Height          =   180
         Left            =   810
         TabIndex        =   13
         Top             =   2550
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总运价:"
         Height          =   180
         Left            =   810
         TabIndex        =   12
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "220.3"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   1740
         TabIndex        =   11
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应拆出金额:"
         Height          =   180
         Left            =   4290
         TabIndex        =   10
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lblNeedSplitMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "220.3"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   5610
         TabIndex        =   9
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐人(&P):"
         Height          =   180
         Left            =   810
         TabIndex        =   8
         Top             =   2910
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐日期(&D):"
         Height          =   180
         Left            =   3360
         TabIndex        =   7
         Top             =   2910
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注(&R):"
         Height          =   180
         Left            =   810
         TabIndex        =   6
         Top             =   3270
         Width           =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         Index           =   1
         X1              =   780
         X2              =   6060
         Y1              =   2370
         Y2              =   2370
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   780
         X2              =   6060
         Y1              =   2385
         Y2              =   2385
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -15
      ScaleHeight     =   735
      ScaleWidth      =   7785
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      Begin VB.Image Image3 
         Height          =   855
         Left            =   2010
         Top             =   -30
         Width           =   5925
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包签发单信息:"
         Height          =   180
         Left            =   270
         TabIndex        =   1
         Top             =   270
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmFinSheetAttrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mSheetID As String
Public FormStatus As eFormStatus

Public g_oActiveUser As ActiveUser


Private oFinanceSheet As Object

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
'    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
'On Error GoTo ErrHandle
'   oFinanceSheet.Identify mSheetID
'   '更改后的属性赋值
'   oFinanceSheet.ActSplitPrice = CDbl(txtActSplitMoney.Text)
''   oFinanceSheet.Status = Trim(cboStatus.Text)
'   oFinanceSheet.OperatorName = Trim(txtOperator.Text)
'   oFinanceSheet.OperateTime = dtpOperDate.Value
'   oFinanceSheet.Remark = Trim(txtRemark.Text)
'   oFinanceSheet.Update
'
'   Unload Me
'Exit Sub
'ErrHandle:
'ShowErrorMsg
End Sub

Private Sub dtpOperDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
cmdOk.Enabled = True
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    AlignFormPos Me
    FormClear
    cmdOk.Enabled = False
    
    Select Case FormStatus
    Case ST_NormalObj
        txtActSplitMoney.Enabled = False
        txtOperator.Enabled = False
        dtpOperDate.Enabled = False
        txtRemark.Enabled = False
    Case ST_EditObj
        txtActSplitMoney.Enabled = True
        txtOperator.Enabled = True
        dtpOperDate.Enabled = True
        txtRemark.Enabled = True
    End Select
    
    Set oFinanceSheet = CreateObject("STLugDss.FinanceSheet")
    oFinanceSheet.Init g_oActiveUser
    oFinanceSheet.Identify mSheetID
    If GetFinTypeString(oFinanceSheet.Status) = mStatusReal Or GetFinTypeString(oFinanceSheet.Status) = mStatusCancel Then
        txtActSplitMoney.Enabled = False
        txtOperator.Enabled = False
        dtpOperDate.Enabled = False
        txtRemark.Enabled = False
        lblStatus.Visible = True
        If GetFinTypeString(oFinanceSheet.Status) = mStatusCancel Then
            lblStatus.ForeColor = vbRed
        End If
    End If
    '填充结算单信息
    FillSheetList
    '填充 vsDetailPrice
    FillvsDetailPrice
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'填充结算单表
Private Sub FillSheetList()
   lblSheetID.Caption = Trim(oFinanceSheet.SheetID)
   lblAcceptType.Caption = GetLuggageTypeString(oFinanceSheet.AcceptType)
   lblFinMonth.Caption = Format(oFinanceSheet.SettleMonth, "yyyy年mm月")
   lblFinDate.Caption = CStr(Format(oFinanceSheet.StartSettleDate, "yyyy-mm-dd")) & " 至 " & CStr(Format(oFinanceSheet.StopSettleDate, "yyyy-mm-dd"))
'   lblVehicle.Caption = Trim(oFinanceSheet.VehicleLicense)
'   lblOwner.Caption = Trim(oFinanceSheet.BusOwnerName)
   lblCompany.Caption = Trim(oFinanceSheet.CompanyName)
   lblDesCompany.Caption = Trim(oFinanceSheet.SplitCompanyName)
   lblProtocol.Caption = Trim(oFinanceSheet.ProtocolName)
   lblTotalPrice.Caption = CStr(oFinanceSheet.TotalPrice)
   lblNeedSplitMoney.Caption = CStr(oFinanceSheet.NeedSplitPrice)
   txtActSplitMoney.Text = oFinanceSheet.ActSplitPrice
   dtpOperDate.Value = oFinanceSheet.OperateTime
   txtOperator.Text = Trim(oFinanceSheet.OperatorName)
   txtRemark.Text = Trim(oFinanceSheet.Remark)
End Sub
'填充vsDetailPrice
Private Sub FillvsDetailPrice()
  On Error GoTo ErrHandle
      Dim i As Integer
      Dim nlen As Integer
      Dim rsTemp As Recordset
      Dim lvItem As ListItem
     '填充列首
     With vsDetailPrice.ColumnHeaders
         .Clear
         .Add , , "车牌号", 1000
         .Add , , "车主", 950
         .Add , , "费用代码", 0
         .Add , , "费用名称", 950
         .Add , , "拆出款", 950
         .Add , , "协议名称", 0
         .Add , , "参营公司简称"
         .Add , , "拆帐公司名称"
     End With
    vsDetailPrice.ListItems.Clear
    Set rsTemp = oFinanceSheet.GetVehicleInfo()
    If rsTemp.RecordCount = 0 Then Exit Sub
    For i = 1 To rsTemp.RecordCount
      Set lvItem = vsDetailPrice.ListItems.Add(, , FormatDbValue(rsTemp!license_tag_no))
       lvItem.SubItems(1) = FormatDbValue(rsTemp!owner_name)
       lvItem.SubItems(2) = FormatDbValue(rsTemp!charge_code)
       lvItem.SubItems(3) = FormatDbValue(rsTemp!charge_name)
       lvItem.SubItems(4) = FormatDbValue(rsTemp!split_out_money)
       lvItem.SubItems(5) = FormatDbValue(rsTemp!protocol_name)
       lvItem.SubItems(6) = FormatDbValue(rsTemp!transport_company_short_name)
       lvItem.SubItems(7) = FormatDbValue(rsTemp!split_company_name)
       vsDetailPrice.ListItems(i).Tag = FormatDbValue(rsTemp!vehicle_id)    '车号
       rsTemp.MoveNext
    Next i
   
  Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub
'清空界面
Private Sub FormClear()
   lblSheetID.Caption = ""
   lblAcceptType.Caption = ""
   lblFinMonth.Caption = ""
   lblFinDate.Caption = ""
   lblCompany.Caption = ""
   lblDesCompany.Caption = ""
   lblProtocol.Caption = ""
   lblTotalPrice.Caption = ""
   lblNeedSplitMoney.Caption = ""
   vsDetailPrice.ListItems.Clear
   txtOperator.Text = ""
   txtRemark.Text = ""
   lblStatus.Caption = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
  SaveFormPos Me
End Sub

Private Sub lblVehicle_Click()

End Sub

Private Sub txtActSplitMoney_Change()
 cmdOk.Enabled = True
End Sub

Private Sub txtOperator_Change()
cmdOk.Enabled = True
End Sub

Private Sub txtRemark_Change()
cmdOk.Enabled = True
End Sub

Private Sub vsDetailPrice_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If vsDetailPrice.SortOrder = lvwAscending Then
    vsDetailPrice.SortOrder = lvwDescending
 Else
    vsDetailPrice.SortOrder = lvwAscending
 End If
    vsDetailPrice.SortKey = ColumnHeader.Index - 1
    vsDetailPrice.Sorted = True
End Sub

