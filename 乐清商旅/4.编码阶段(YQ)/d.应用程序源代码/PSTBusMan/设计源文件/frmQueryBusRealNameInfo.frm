VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{A5E8F770-DA22-4EAF-B7BE-73B06021D09F}#1.1#0"; "ST6Report.ocx"
Begin VB.Form frmQueryBusRealNameInfo 
   BackColor       =   &H00F2D2BF&
   Caption         =   "车次实名制信息查询"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   15285
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptTop 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2D2BF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   0
      ScaleHeight     =   1365
      ScaleWidth      =   12405
      TabIndex        =   0
      Top             =   0
      Width           =   12405
      Begin VB.TextBox txtSheetNo 
         Height          =   300
         Left            =   6750
         TabIndex        =   16
         Top             =   967
         Width           =   1245
      End
      Begin VB.TextBox txtBusSerialNo 
         Height          =   300
         Left            =   4320
         TabIndex        =   15
         Top             =   967
         Width           =   1245
      End
      Begin VB.CheckBox chkStopCheck 
         BackColor       =   &H00F2D2BF&
         Caption         =   "已停检 "
         Height          =   315
         Left            =   3090
         TabIndex        =   9
         Top             =   75
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.OptionButton optCheck 
         BackColor       =   &H00F2D2BF&
         Caption         =   "按检票(&L)"
         Height          =   285
         Left            =   1980
         TabIndex        =   8
         Top             =   90
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.ComboBox cboSellStation 
         Height          =   300
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   967
         Width           =   1545
      End
      Begin VB.OptionButton optSell 
         BackColor       =   &H00F2D2BF&
         Caption         =   "按售票(&B)"
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   90
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpBusDate 
         Height          =   285
         Left            =   1290
         TabIndex        =   2
         Top             =   525
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   3670019
         CurrentDate     =   37512
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Height          =   330
         Left            =   5640
         TabIndex        =   11
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
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
         MICON           =   "frmQueryBusRealNameInfo.frx":0000
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
         Left            =   6960
         TabIndex        =   12
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
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
         MICON           =   "frmQueryBusRealNameInfo.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin FText.asFlatTextBox txtBusID 
         Height          =   300
         Left            =   4320
         TabIndex        =   10
         Top             =   510
         Width           =   1245
         _ExtentX        =   2196
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
         Registered      =   -1  'True
      End
      Begin VB.Label lblSheetNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路单号(&N):"
         Height          =   180
         Left            =   5850
         TabIndex        =   14
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label lblBusSerialNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次序号(&S):"
         Height          =   180
         Left            =   3240
         TabIndex        =   13
         Top             =   1027
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码(&B):"
         Height          =   180
         Index           =   0
         Left            =   3240
         TabIndex        =   7
         Top             =   577
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站(&D):"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车日期(&R):"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   570
         Width           =   1080
      End
   End
   Begin ST6Report.RTReport RTReport1 
      Height          =   3315
      Left            =   0
      TabIndex        =   3
      Top             =   1380
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   5847
   End
End
Attribute VB_Name = "frmQUeryBusRealNameInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_lRange As Long '写进度条用

Private Sub SetVisble(pbVisible As Boolean)
    If pbVisible = True Then
        lblBusSerialNo.Visible = True
        txtBusSerialNo.Visible = True
        lblSheetNo.Visible = True
        txtSheetNo.Visible = True
    Else
        lblBusSerialNo.Visible = False
        txtBusSerialNo.Visible = False
        txtBusSerialNo.Text = ""
        lblSheetNo.Visible = False
        txtSheetNo.Visible = False
        txtSheetNo.Text = ""
    End If
End Sub

Private Sub cmdQuery_Click()
    m_lRange = 0
    Query
End Sub

Private Sub Form_Load()
    dtpBusDate.Value = ToDBDate(Date)
    FillSellStation
End Sub

Private Sub Query()
On Error GoTo ErrHandle
'    '开始查询
    Dim rsTemp As Recordset
    Dim oTemp As Object
    Dim i As Integer
    Dim aszTemp() As String
    Dim szSellStationID As String
    Dim szVehcile As String
    Dim szBusStartTime As String
    Dim nMan As Integer
    Dim nWomen As Integer
    Dim nTotal As Integer
    
    Set oTemp = CreateObject("STDss.TicketSellerDim")
    
    szSellStationID = ""
    szVehcile = ""
    szBusStartTime = ""
    nMan = 0
    nWomen = 0
    nTotal = 0
    
    If Trim(txtBusID.Text) = "" Then MsgBox "车次不能为空！", vbExclamation, Me.Caption: Exit Sub
    
    If chkStopCheck.Value = vbChecked And optCheck.Value = True And txtBusID.Tag = "滚动车次" Then
        If Trim(txtBusSerialNo.Text) = "" And Trim(txtSheetNo.Text) = "" Then MsgBox "车次序号和客凭号不能同时为空！", vbExclamation, Me.Caption: Exit Sub
    End If
    
    szSellStationID = IIf(ResolveDisplay(cboSellStation.Text) = "(所有上车站)", "", ResolveDisplay(cboSellStation.Text))
    If szSellStationID = "" Then
        If MsgBox("未选择上车站", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If

    oTemp.Init g_oActiveUser
    Set rsTemp = oTemp.QueryBusRealNameInfo(dtpBusDate.Value, dtpBusDate.Value, szSellStationID, Trim(txtBusID.Text), IIf(optCheck.Value = True, True, False), IIf(chkStopCheck.Value = vbChecked, True, False), IIf(Trim(txtBusSerialNo.Text) = "", -1, Trim(txtBusSerialNo.Text)), Trim(txtSheetNo.Text))
    Set oTemp = Nothing
    nTotal = rsTemp.RecordCount
    
    If nTotal > 0 Then
        rsTemp.MoveFirst
        For i = 1 To nTotal
            If Trim(FormatDbValue(rsTemp!license_tag_no)) <> "" Then
                szVehcile = Trim(FormatDbValue(rsTemp!license_tag_no))
            End If
            If Trim(FormatDbValue(rsTemp!bus_start_time)) <> "" Then
                szBusStartTime = Format(FormatDbValue(rsTemp!bus_start_time), "YYYY-MM-DD HH:MM")
            End If
            If FormatDbValue(rsTemp!Sex) = "男" Then
                nMan = nMan + 1
            ElseIf FormatDbValue(rsTemp!Sex) = "女" Then
                nWomen = nWomen + 1
            End If
            rsTemp.MoveNext
        Next i
    End If
    
    Erase aszTemp
    ReDim aszTemp(1 To 7, 1 To 2)
    aszTemp(1, 1) = "发车时间"
    aszTemp(1, 2) = szBusStartTime
    aszTemp(2, 1) = "车次"
    aszTemp(2, 2) = Trim(txtBusID.Text)
    aszTemp(3, 1) = "车牌"
    aszTemp(3, 2) = szVehcile
    aszTemp(4, 1) = "男人"
    aszTemp(4, 2) = nMan & "人"
    aszTemp(5, 1) = "女人"
    aszTemp(5, 2) = nWomen & "人"
    aszTemp(6, 1) = "人数"
    aszTemp(6, 2) = nTotal & "人"
    aszTemp(7, 1) = "上车站"
    aszTemp(7, 2) = ResolveDisplayEx(cboSellStation.Text)

    '填充票价记录集
    '设置固定行,列可见性
    RTReport1.LeftLabelVisual = True
    RTReport1.TopLabelVisual = True
    RTReport1.TemplateFile = "车次实名制查询.xls"
    RTReport1.ShowReport rsTemp, aszTemp

    WriteProcessBar False
    ShowSBInfo
    ShowSBInfo "共有" & m_lRange & "条记录", ESB_ResultCountInfo
    Exit Sub
ErrHandle:
    Set oTemp = Nothing
    WriteProcessBar False
    ShowSBInfo
    ShowErrorMsg
End Sub

Private Sub Form_Activate()
    Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Const cnMargin = 50
    ptTop.Left = 0
    ptTop.Top = 0
    ptTop.Width = Me.ScaleWidth
    RTReport1.Left = cnMargin
    RTReport1.Top = ptTop.Height + cnMargin
'    RTReport1.Width = Me.ScaleWidth - IIf(abAction.Visible, abAction.Width, 0) - 2 * cnMargin
    RTReport1.Width = Me.ScaleWidth - 2 * cnMargin
    RTReport1.Height = Me.ScaleHeight - ptTop.Height - 2 * cnMargin
'    '当操作条关闭时间处理
'    If abAction.Visible Then
'        abAction.Move RTReport1.Width + cnMargin, RTReport1.Top
'        abAction.Height = RTReport1.Height
'    End If
End Sub

Private Sub cmdCancel_Click()
    '关闭
    Unload Me
End Sub

Private Sub optCheck_Click()
    chkStopCheck.Visible = True
End Sub

Private Sub optSell_Click()
    chkStopCheck.Visible = False
End Sub

Private Sub RTReport1_SetProgressRange(ByVal lRange As Variant)
    m_lRange = lRange
End Sub

Private Sub RTReport1_SetProgressValue(ByVal lValue As Variant)
    WriteProcessBar True, lValue, m_lRange, "正在填充统计信息..."
End Sub

Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    On Error GoTo ErrorHandle
    RTReport1.PrintReport pbShowDialog
    Exit Sub
ErrorHandle:
End Sub

Public Sub PreView()
    RTReport1.PrintView
End Sub

Public Sub PageSet()
    RTReport1.OpenDialog EDialogType.PAGESET_TYPE
End Sub

Public Sub PrintSet()
    RTReport1.OpenDialog EDialogType.PRINT_TYPE
End Sub
'导出文件
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = RTReport1.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'导出文件并打开
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = RTReport1.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub

Private Sub txtBusID_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectREBus(dtpBusDate.Value, False, False)
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtBusID.Text = aszTmp(1, 1)
    txtBusID.Tag = aszTmp(1, 2)
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub FillSellStation()
    Dim nCount As Integer
    Dim i As Integer
    cboSellStation.Clear
    nCount = ArrayLength(g_atAllSellStation)
    For i = 1 To nCount
        cboSellStation.AddItem MakeDisplayString(g_atAllSellStation(i).szSellStationID, g_atAllSellStation(i).szSellStationName)
    Next i
End Sub

Private Sub txtBusSerialNo_GotFocus()
    With txtBusSerialNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSheetNo_GotFocus()
    With txtSheetNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
