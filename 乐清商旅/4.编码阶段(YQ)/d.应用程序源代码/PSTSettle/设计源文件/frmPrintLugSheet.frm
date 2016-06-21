VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#1.4#0"; "RTReportlf.ocx"
Begin VB.Form frmPrintLugSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打印行包结算单"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9255
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ListView lvLugSheet 
      Height          =   3315
      Left            =   150
      TabIndex        =   9
      Top             =   1920
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   5847
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "行包结算单号"
         Object.Width           =   2540
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdveiw 
      Height          =   345
      Left            =   1050
      TabIndex        =   8
      Top             =   1020
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "预览(&V)"
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
      MICON           =   "frmPrintLugSheet.frx":0000
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
      Height          =   345
      Left            =   150
      TabIndex        =   7
      Top             =   1020
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   609
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
      MICON           =   "frmPrintLugSheet.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtLugSheetID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   6
      Text            =   "200340002"
      Top             =   1470
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -60
      ScaleHeight     =   735
      ScaleWidth      =   9345
      TabIndex        =   3
      Top             =   0
      Width           =   9345
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请填写要统计的行包结算单号:"
         Height          =   180
         Left            =   330
         TabIndex        =   4
         Top             =   270
         Width           =   2430
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   3390
         Picture         =   "frmPrintLugSheet.frx":0038
         Top             =   -30
         Width           =   5925
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "RTStation"
      Height          =   765
      Left            =   -30
      TabIndex        =   0
      Top             =   5370
      Width           =   9375
      Begin RTComctl3.CoolButton cmdCancel 
         Height          =   345
         Left            =   8130
         TabIndex        =   1
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
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
         MICON           =   "frmPrintLugSheet.frx":1522
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
         Height          =   345
         Left            =   7020
         TabIndex        =   2
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "确定(&E)"
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
         MICON           =   "frmPrintLugSheet.frx":153E
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
   Begin RTReportLF.RTReport RTReport1 
      Height          =   4245
      Left            =   2340
      TabIndex        =   5
      Top             =   990
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7488
   End
End
Attribute VB_Name = "frmPrintLugSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSheetData As Recordset
Dim szSheetCusTom() As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdveiw_Click()
    RTReport1.PrintView
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    
    txtLugSheetID.Text = ""
    lstLugSheet.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("您真的确定不打印行包结算单吗?") = vbYes Then
        Unload Me
        SaveFormPos Me
    Else
        Cancel = -1
    End If
End Sub

Private Sub PrintSheetReport()
On Error GoTo ErrHandle
    RTReport1.TemplateFile = App.Path & "\行包结算单.xls"
    RTReport1.ShowReport , szSheetCusTom
  
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub GetFinSheetInfo()
On Error GoTo ErrHandle

    Dim i As Integer
    Dim mStartDate As Date
    Dim mEndDate As Date
                 
    
    m_oLugFinSvr.Init g_oActiveUser
        Set rsSheetData = m_oLugFinSvr.PrintLugFinSheet(SheetID)
        If rsSheetData.RecordCount > 0 Then
        
        '创建自定义项目集
        ReDim szSheetCusTom(1 To 8, 1 To 2)
        szSheetCusTom(1, 1) = "结算期限"
        szSheetCusTom(1, 2) = Format(rsSheetData!settlement_start_time, "YYYY年MM月dd日") & "-" & Format(rsSheetData!settlement_end_time, "YYYY年MM月DD日")
        szSheetCusTom(2, 1) = "大写金额"
        szSheetCusTom(2, 2) = GetNumber(FormatDbValue(rsSheetData!need_split_out))
        Select Case FormatDbValue(rsSheetData!split_object_type)
               Case ObjectType.VehicleType
                    szSheetCusTom(3, 1) = "车牌号码"
                    szSheetCusTom(3, 2) = FormatDbValue(rsSheetData!split_object_name)
               Case ObjectType.TranportCompanyType
                    szSheetCusTom(3, 1) = "参运公司"
                    szSheetCusTom(3, 2) = FormatDbValue(rsSheetData!split_object_name)
        End Select
        szSheetCusTom(4, 1) = "结算单号"
        szSheetCusTom(4, 2) = FormatDbValue(rsSheetData!fin_sheet_id)
        szSheetCusTom(5, 1) = "行包运费"
        szSheetCusTom(5, 2) = FormatDbValue(rsSheetData!total_price)
        szSheetCusTom(6, 1) = "拆出金额"
        szSheetCusTom(6, 2) = FormatDbValue(rsSheetData!need_split_out)
        szSheetCusTom(7, 1) = "财务"
        szSheetCusTom(7, 2) = FormatDbValue(rsSheetData!Operator)
        szSheetCusTom(7, 1) = "拆算协议"
        szSheetCusTom(7, 2) = FormatDbValue(rsSheetData!protocol_name)
        
        WriteProcessBar True, , , "正在形成报表..."
        PrintSheetReport

    End If

    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

 '打开行包结算单属性
Private Sub ShowFinSheet()
    
    Dim oCommDialog As New STShell.CommDialog
    On Error GoTo ErrorHandle
    
    If lvFinSheets.SelectedItem Is Nothing Then Exit Sub
    oCommDialog.Init g_oActiveUser
'    oCommDialog.ShowLugFinSheet lvLugSheet.SelectedItem.Text, Trim(lvLugSheet.SelectedItem.SubItems(1))
    Set oCommDialog = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub lvLugSheet_DblClick()
    ShowFinSheet
End Sub
