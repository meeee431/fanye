VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Object = "{695ABF14-B2D8-11D2-A5ED-DE08DCF33612}#1.2#0"; "asfcombo.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmQuerySheet 
   BackColor       =   &H00E0E0E0&
   Caption         =   "行包签发单查询"
   ClientHeight    =   6150
   ClientLeft      =   3540
   ClientTop       =   2520
   ClientWidth     =   10260
   ControlBox      =   0   'False
   HelpContextID   =   7000100
   Icon            =   "frmQuerySheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   10260
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptResult 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6060
      Left            =   2940
      ScaleHeight     =   6060
      ScaleWidth      =   7380
      TabIndex        =   12
      Top             =   0
      Width           =   7380
      Begin RTReportLF.RTReport flReport 
         Height          =   5955
         Left            =   75
         TabIndex        =   18
         Top             =   75
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   10504
      End
   End
   Begin VB.PictureBox ptQuery 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   60
      ScaleHeight     =   5985
      ScaleWidth      =   2805
      TabIndex        =   0
      Top             =   30
      Width           =   2835
      Begin VB.OptionButton OptQueryMode5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定装卸工"
         Height          =   240
         Left            =   75
         TabIndex        =   20
         Top             =   4080
         Width           =   2115
      End
      Begin VB.OptionButton OptQueryMode4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "已废的签发单"
         Height          =   240
         Left            =   75
         TabIndex        =   19
         Top             =   3780
         Width           =   2115
      End
      Begin FCmbo.asFlatCombo txtBusID 
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   780
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   476
         ButtonDisabledForeColor=   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonHotBackColor=   8421504
         ButtonPressedBackColor=   0
         Text            =   ""
         ButtonBackColor =   8421504
         OfficeXPColors  =   -1  'True
      End
      Begin VB.TextBox txtluggage 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   780
         Width           =   2460
      End
      Begin VB.OptionButton OptQueryMode3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定签发的受理单号"
         Height          =   375
         Left            =   75
         TabIndex        =   15
         Top             =   3405
         Width           =   2265
      End
      Begin VB.TextBox txtSheetID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   780
         Width           =   2460
      End
      Begin VB.OptionButton optQueryMode1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定车辆"
         Height          =   285
         Left            =   75
         TabIndex        =   2
         Top             =   2910
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton OptQueryMode2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定签发单号"
         Height          =   285
         Left            =   75
         TabIndex        =   1
         Top             =   3165
         Width           =   1770
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   2280
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   71499779
         UpDown          =   -1  'True
         CurrentDate     =   37062
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Top             =   1680
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   71499779
         UpDown          =   -1  'True
         CurrentDate     =   36528
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HorizontalAlignment=   1
         Caption         =   "查询条件设定"
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   1440
         TabIndex        =   7
         Top             =   4575
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "关闭"
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
         MICON           =   "frmQuerySheet.frx":038A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   345
         Left            =   150
         TabIndex        =   8
         Top             =   4575
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
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
         MICON           =   "frmQuerySheet.frx":03A6
         PICN            =   "frmQuerySheet.frx":03C2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   225
         Left            =   90
         TabIndex        =   14
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "查询方式;"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   2685
         Width           =   900
      End
      Begin VB.Label lblQueryType 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "承运车辆(&T):"
         Height          =   180
         Left            =   135
         TabIndex        =   10
         Top             =   510
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "签发时间(&D):"
         Height          =   180
         Left            =   105
         TabIndex        =   9
         Top             =   1185
         Width           =   1080
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2640
         Y1              =   4425
         Y2              =   4425
      End
   End
End
Attribute VB_Name = "frmQuerySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'Last Modify By: 陆勇庆  2005-8-16
'Last Modify In:增加了装卸工的查询
'*******************************************************************************
Option Explicit
Dim mlMaxCount As Long
Dim mnLastSearchIndex As Integer
Private Sub Form_Deactivate()
    HideSheetNoLabel
    mdiMain.SetPrintEnabled False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
    frmSearchVechile.mFormNum = 2
    frmSearchVechile.StartSearchIndex = mnLastSearchIndex
    frmSearchVechile.Show vbModal
    mnLastSearchIndex = txtBusID.ListIndex
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF1 Then
        DisplayHelp Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    HideSheetNoLabel
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdQuery_Click()
    QuerySheet
End Sub
Private Sub QuerySheet()
On Error GoTo ErrHandle
    Dim rsTemp As Recordset
    Dim szStatus As Integer
    ReDim avCustmonData(1 To 3, 1 To 2)
    avCustmonData(1, 1) = "开始日期"
    avCustmonData(1, 2) = dtpStart.Value
    avCustmonData(2, 1) = "结束日期"
    avCustmonData(2, 2) = dtpEnd.Value
    avCustmonData(3, 1) = "查询方式"
    
    If optQueryMode1.Value = True Then
'        If txtBusID.Text = "" Then
'            MsgBox "承载车辆不能为空", vbExclamation, "行包签发"
'            Exit Sub
'        End If
        Set rsTemp = moLugSvr.GetCarrySheetRS(CDate(dtpStart), CDate(dtpEnd), , Trim(txtBusID.Text))
        avCustmonData(3, 2) = optQueryMode1.Caption & "且车辆为" & IIf(txtBusID.Text = "", "空", txtBusID.Text)
    ElseIf OptQueryMode2.Value = True Then
        If txtSheetID.Text = "" Then
            Set rsTemp = moLugSvr.GetCarrySheetRS(CDate(dtpStart), CDate(dtpEnd), Trim(txtSheetID.Text))
        Else
            Set rsTemp = moLugSvr.GetCarrySheetRS(, , Trim(txtSheetID.Text))
        End If
        avCustmonData(3, 2) = OptQueryMode2.Caption & "且签发单为" & IIf(txtSheetID.Text = "", "空", txtSheetID.Text)
    ElseIf OptQueryMode3.Value = True Then
        If txtluggage.Text = "" Then
            MsgBox "指定签发的受理单不能为空", vbExclamation, "行包签发"
            Exit Sub
        End If
        Set rsTemp = moLugSvr.GetCarrySheetRSEX(Trim(txtluggage.Text))
        avCustmonData(3, 2) = OptQueryMode3.Caption & "且指定签发的受理单为" & IIf(txtluggage.Text = "", "空", txtluggage.Text)
    ElseIf OptQueryMode4.Value Then
        Set rsTemp = moLugSvr.GetCarrySheetRS(CDate(dtpStart), CDate(dtpEnd), , , ELuggageSheetValidStatus.ST_LuggageSheetValidStatusCancel)
        avCustmonData(3, 2) = OptQueryMode4.Caption
    ElseIf OptQueryMode5.Value = True Then
        If txtSheetID.Text = "" Then
            Set rsTemp = moLugSvr.GetCarrySheetRS(CDate(dtpStart), CDate(dtpEnd))
        Else
            Set rsTemp = moLugSvr.GetCarrySheetRS(CDate(dtpStart), CDate(dtpEnd), , , , Trim(txtSheetID.Text))
        End If
        avCustmonData(3, 2) = OptQueryMode5.Caption & "且装卸工为" & IIf(txtSheetID.Text = "", "空", txtSheetID.Text)
    End If
    
    
    
    '显示查询的结果
'    If rsTemp.RecordCount > 0 Then
        ShowReport rsTemp, "行包签发单明细报表.xls", "行包签发单明细报表", avCustmonData
'    Else
'        MsgBox "符合条件的信息不存在！", vbInformation, "行包受理"
'    End If
    
    
    
    
  Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub



Private Sub Form_Load()
    Dim szaTemp() As String
    Dim nlen As Integer
    Dim i As Integer
    '得到所有车辆信息
    m_obase.Init m_oAUser
    szaTemp = m_obase.GetVehicle()
    nlen = ArrayLength(szaTemp)
    If nlen > 0 Then
    For i = 1 To nlen
    txtBusID.AddItem szaTemp(i, 2)
    Next i
    End If
    dtpStart.Value = Date
    dtpEnd.Value = DateAdd("s", -1, DateAdd("d", 1, dtpStart.Value))
    
    optQueryMode1.Value = True
    txtBusID.Visible = True
    txtluggage.Visible = False
    txtSheetID.Visible = False
End Sub

Private Sub Form_Resize()
 On Error Resume Next
    Const cnMargin = 50
    ptQuery.Move cnMargin, cnMargin, ptQuery.Width, Me.ScaleHeight - 2 * cnMargin
    ptResult.Move cnMargin + ptQuery.Width + 2 * cnMargin, cnMargin, Me.ScaleWidth - ptQuery.Width - 4 * cnMargin, Me.ScaleHeight - 2 * cnMargin
End Sub



Private Sub OptQueryMode3_Click()
    txtBusID.Visible = False
    txtSheetID.Visible = False
    txtluggage.Visible = True
    txtluggage.Text = ""
    lblQueryType.Caption = "指定签发的受理单号(&U)"
End Sub

Private Sub OptQueryMode1_Click()
    txtBusID.Visible = True
    txtBusID.Text = ""
    txtSheetID.Visible = False
    txtluggage.Visible = False
    lblQueryType.Caption = "承运车辆(&T):"
    
End Sub

Private Sub OptQueryMode2_Click()
    txtBusID.Visible = False
    txtSheetID.Visible = True
    txtluggage.Visible = False
    txtSheetID.Text = ""
    lblQueryType.Caption = "签发单号(&T):"
    
End Sub

Private Sub Form_Activate()
    SetSheetNoLabel False, g_szCarrySheetID


    mdiMain.SetPrintEnabled True
    
End Sub

Private Sub OptQueryMode5_Click()
    txtBusID.Visible = False
    txtSheetID.Visible = True
    txtluggage.Visible = False
    txtSheetID.Text = ""
    lblQueryType.Caption = "装卸工(&T):"
End Sub

Private Sub ptResult_Resize()
On Error Resume Next
    Const cnMargin = 80
    
    flReport.Move cnMargin, cnMargin, ptResult.ScaleWidth - 2 * cnMargin, ptResult.ScaleHeight - 3 * cnMargin
End Sub


Private Sub txtBusID_Change()
cmdQuery.Enabled = True
End Sub

Private Sub txtBusID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub txtSheetID_Change()
cmdQuery.Enabled = True
End Sub

Private Sub txtSheetID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
End Sub
Private Sub txtluggage_Change()
        cmdQuery.Enabled = True
End Sub

Private Sub txtluggage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Public Function ShowReport(prsData As Recordset, pszFileName As String, pszCaption As String, Optional pvaCustomData As Variant, Optional pnReportType As Integer = 0) As Long
    On Error GoTo Error_Handle
    Me.ZOrder 0
    Me.Show
    
    Dim arsTemp As Variant
    Dim aszTemp As Variant
'    Dim rsTemp As Recordset
    ReDim aszTemp(1 To 1)
    ReDim arsTemp(1 To 1)
    '赋票种
    aszTemp(1) = "托运费项"
    Set arsTemp(1) = g_rsPriceItem
    
'    Me.Caption = pszCaption
    
    WriteProcessBar True, , , "正在形成报表..."
    
    flReport.CustomStringCount = aszTemp
    flReport.CustomString = arsTemp
    flReport.SheetTitle = ""
    
    flReport.TemplateFile = App.Path & "\" & pszFileName
    flReport.LeftLabelVisual = True
    flReport.TopLabelVisual = True
    flReport.ShowReport prsData, pvaCustomData
    WriteProcessBar False, , , ""
    ShowSBInfo "共" & prsData.RecordCount & "条记录", ESB_ResultCountInfo
    Exit Function
Error_Handle:
    ShowErrorMsg
End Function




Public Sub ExportToFile()

    flReport.OpenDialog EDialogType.EXPORT_FILE
End Sub

Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    On Error GoTo ErrorHandle
    flReport.PrintReport pbShowDialog
    Exit Sub
ErrorHandle:
End Sub

Public Sub PreView()
    flReport.PrintView
End Sub

Public Sub PageSet()
    flReport.OpenDialog EDialogType.PAGESET_TYPE
End Sub

Public Sub PrintSet()
    flReport.OpenDialog EDialogType.PRINT_TYPE
End Sub
'导出文件
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = flReport.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'导出文件并打开
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = flReport.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub




Private Sub flReport_SetProgressRange(ByVal lRange As Variant)
    mlMaxCount = lRange
End Sub

Private Sub flReport_SetProgressValue(ByVal lValue As Variant)
    WriteProcessBar True, lValue, mlMaxCount
End Sub


