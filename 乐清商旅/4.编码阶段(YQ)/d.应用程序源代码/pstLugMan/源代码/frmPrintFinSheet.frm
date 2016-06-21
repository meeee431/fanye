VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmPrintFinSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "结算单"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmPrintFinSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9165
   StartUpPosition =   3  '窗口缺省
   Begin RTComctl3.CoolButton cmdPrivew 
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   5850
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "打印预览(&V)"
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
      MICON           =   "frmPrintFinSheet.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdPrint 
      Height          =   345
      Left            =   6630
      TabIndex        =   0
      Top             =   5820
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "打印(&P)"
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
      MICON           =   "frmPrintFinSheet.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTReportLF.RTReport RTReport1 
      Height          =   5595
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   9869
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   7890
      TabIndex        =   2
      Top             =   5820
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmPrintFinSheet.frx":0342
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
Attribute VB_Name = "frmPrintFinSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSheetData As Recordset
Dim szSheetCusTom() As String
Public SheetID As String
Public OldSheetID As String  '用于重打的原结算单号
Public mRePrint As Boolean '是否是重打结算单
Private Sub cmdExit_Click()
    If MsgBox("还未打印结算单,退出吗?", vbYesNoCancel + vbQuestion, Me.Caption) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrHandle
    RTReport1.PrintReport True
    Unload Me
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub PrintSheetReport()
On Error GoTo ErrHandle
    RTReport1.TemplateFile = App.Path & "\行包结算单.xls"
    RTReport1.ShowReport , szSheetCusTom
  
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdPrivew_Click()
    RTReport1.PrintView
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    GetFinSheetInfo
End Sub

Private Sub GetFinSheetInfo()
On Error GoTo ErrHandle

    Dim i As Integer
    Dim mStartDate As Date
    Dim mEndDate As Date
                 
    
        m_oLugFinSvr.Init m_oAUser
    If mRePrint = False Then
        Set rsSheetData = m_oLugFinSvr.PrintLugFinSheet(SheetID)
    Else
        Set rsSheetData = m_oLugFinSvr.RSPrintLugFinSheet(SheetID, OldSheetID)
    End If
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

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    
    WriteProcessBar False, , , ""
    Set rsSheetData = Nothing
    Unload Me
End Sub
