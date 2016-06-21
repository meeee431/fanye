VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmRouteTransport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "线路运量统计"
   ClientHeight    =   4125
   ClientLeft      =   4980
   ClientTop       =   3900
   ClientWidth     =   5010
   HelpContextID   =   60000170
   Icon            =   "frmRouteTransport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   2130
      TabIndex        =   17
      Top             =   2370
      Width           =   1815
   End
   Begin RTComctl3.CoolButton cmdChart 
      Height          =   315
      Left            =   60
      TabIndex        =   16
      Top             =   3660
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "图表"
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
      MICON           =   "frmRouteTransport.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   1230
      TabIndex        =   14
      Top             =   3660
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
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
      MICON           =   "frmRouteTransport.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1965
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -90
      TabIndex        =   10
      Top             =   690
      Width           =   6885
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   345
      Left            =   2490
      TabIndex        =   1
      Top             =   3660
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "确定"
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
      MICON           =   "frmRouteTransport.frx":0044
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
      Height          =   345
      Left            =   3750
      TabIndex        =   0
      Top             =   3660
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmRouteTransport.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   2130
      TabIndex        =   2
      Top             =   1470
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   2130
      TabIndex        =   3
      Top             =   990
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin FText.asFlatTextBox txtRouteID 
      Height          =   300
      Left            =   2130
      TabIndex        =   4
      Top             =   2745
      Width           =   1815
      _ExtentX        =   3201
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
      Locked          =   -1  'True
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      HelpContextID   =   60000170
      Left            =   -180
      TabIndex        =   11
      Top             =   3420
      Width           =   8745
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   8
      Top             =   0
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   9
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.CheckBox chkIsCompany 
      BackColor       =   &H00E0E0E0&
      Caption         =   "分参运公司统计"
      Height          =   285
      Left            =   870
      TabIndex        =   15
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   870
      TabIndex        =   13
      Top             =   2025
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   870
      TabIndex        =   7
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   870
      TabIndex        =   6
      Top             =   1530
      Width           =   1080
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "线路(&T):"
      Height          =   180
      Left            =   870
      TabIndex        =   5
      Top             =   2805
      Width           =   720
   End
End
Attribute VB_Name = "frmRouteTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm
Const cszFileName1 = "线路运量简报模板.xls"
Const cszFileName2 = "线路运量简报模板(分参运公司).xls"
Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant
Dim m_szCode As String
Dim m_szCode2 As String
Dim szStatus As Boolean


Private Sub chkIsCompany_Click()
    If chkIsCompany.Value = vbChecked Then
        szStatus = True
    Else
        szStatus = False
    End If
End Sub

'Private m_aszSum(1 To 38) As Double '总和

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Function MergeArea(prsArea1 As Recordset, prsArea2 As Recordset) As Recordset
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    
    
    With rsTemp.Fields
        .Append "route_id", adChar, 4
        .Append "route_name", adChar, 16
        .Append "bus_count", adInteger
        .Append "passenger_number", adDouble
        .Append "fact_float", adDouble
        .Append "total_float", adDouble
        .Append "total_number", adDouble
        .Append "full_seat_rate", adDouble
        .Append "fact_load_rate", adDouble
        .Append "total_ticket_price", adCurrency

    End With
    rsTemp.Open
    If Not (prsArea1 Is Nothing) Then
        If prsArea1.RecordCount > 0 Then
            prsArea1.MoveFirst
            For i = 1 To prsArea1.RecordCount
                rsTemp.AddNew
                rsTemp!route_id = prsArea1!route_id
                rsTemp!route_name = prsArea1!route_name
                rsTemp!bus_count = 0 ' prsArea1!bus_count
                rsTemp!passenger_number = prsArea1!passenger_number
                rsTemp!fact_float = prsArea1!fact_float
                rsTemp!total_float = prsArea1!total_float
                rsTemp!total_number = prsArea1!total_number
                rsTemp!full_seat_rate = prsArea1!full_seat_rate
                rsTemp!fact_load_rate = prsArea1!fact_load_rate
                rsTemp!total_ticket_price = prsArea1!total_ticket_price
                
                
        
                prsArea1.MoveNext
            Next i
            rsTemp.Update
        End If
        '赋记录集
    End If
    nCount = rsTemp.RecordCount
    If Not (prsArea2 Is Nothing) Then
        If prsArea2.RecordCount > 0 Then prsArea2.MoveFirst
        For i = 1 To prsArea2.RecordCount
            
            rsTemp.MoveFirst
            For j = 1 To nCount
                If prsArea2!route_id = rsTemp!route_id Then Exit For
                rsTemp.MoveNext
                
            Next j
            If j > nCount Then
                rsTemp.AddNew
                rsTemp!route_id = prsArea2!route_id
                rsTemp!route_name = prsArea2!route_name
                
                rsTemp!bus_count = prsArea2!fact_Bus_number
    
                rsTemp!passenger_number = 0
                rsTemp!fact_float = 0
                rsTemp!total_float = 0
                rsTemp!total_number = 0
                rsTemp!full_seat_rate = 0
                rsTemp!fact_load_rate = 0
                rsTemp!total_ticket_price = 0
            Else
                
                '已经存在
                rsTemp!bus_count = prsArea2!fact_Bus_number
                
            End If
            prsArea2.MoveNext
        Next i
        If rsTemp.RecordCount > 0 Then rsTemp.Update
        
    End If
    
    Set MergeArea = rsTemp
    
End Function

Private Sub cmdChart_Click()
    
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim frmTemp As frmChart
    Dim oRouteDim As New TicketRouteDim
    oRouteDim.Init m_oActiveUser
    Me.MousePointer = vbHourglass
    
    If MDIMain.m_szMethod = False Then
        Set rsTemp = oRouteDim.GetRouteTransport(dtpBeginDate.Value, dtpEndDate.Value, m_szCode, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)), szStatus)
    Else
        Set rsTemp = oRouteDim.GetRouteTransportByCheck(dtpBeginDate.Value, dtpEndDate.Value, m_szCode, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)), szStatus)
    End If
    
    Dim rsData As New Recordset
    With rsData.Fields
        .Append "route_id", adBSTR
        .Append "bus_count", adBigInt
    End With
    rsData.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsData.AddNew
        rsData!route_id = FormatDbValue(rsTemp!route_id)
        rsData!bus_count = FormatDbValue(rsTemp!bus_count)
        rsTemp.MoveNext
        rsData.Update
    Next i
    
    Dim rsdata2 As New Recordset
    With rsdata2.Fields
        .Append "route_id", adBSTR
        .Append "passenger_number", adBigInt
    End With
    rsdata2.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata2.AddNew
        rsdata2!route_id = FormatDbValue(rsTemp!route_id)
        rsdata2!passenger_number = FormatDbValue(rsTemp!passenger_number)
        rsTemp.MoveNext
        rsdata2.Update
    Next i
    
    Dim rsdata3 As New Recordset
    With rsdata3.Fields
        .Append "route_id", adBSTR
        .Append "fact_float", adBigInt
    End With
    rsdata3.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata3.AddNew
        rsdata3!route_id = FormatDbValue(rsTemp!route_id)
        rsdata3!fact_float = FormatDbValue(rsTemp!fact_float)
        rsTemp.MoveNext
        rsdata3.Update
    Next i
    
    Dim rsdata4 As New Recordset
    With rsdata4.Fields
        .Append "route_id", adBSTR
        .Append "fact_load_rate", adBigInt
    End With
    rsdata4.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata4.AddNew
        rsdata4!route_id = FormatDbValue(rsTemp!route_id)
        rsdata4!fact_load_rate = FormatDbValue(rsTemp!fact_load_rate)
        rsTemp.MoveNext
        rsdata4.Update
    Next i
    
    Dim rsdata5 As New Recordset
    With rsdata5.Fields
        .Append "route_id", adBSTR
        .Append "total_ticket_price", adBigInt
    End With
    rsdata5.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata5.AddNew
        rsdata5!route_id = FormatDbValue(rsTemp!route_id)
        rsdata5!total_ticket_price = FormatDbValue(rsTemp!total_ticket_price)
        rsTemp.MoveNext
        rsdata5.Update
    Next i

    Me.Hide
    Set frmTemp = New frmChart
    frmTemp.ClearChart
    frmTemp.AddChart "总趟次", rsData
    frmTemp.AddChart "人数", rsdata2
    frmTemp.AddChart "实际周转量", rsdata3
    frmTemp.AddChart "实载率", rsdata4
    frmTemp.AddChart "营收", rsdata5
    frmTemp.ShowChart "线路运量简报"
    Set frmTemp = Nothing
    Unload Me

    Exit Sub
Error_Handle:
    Set frmTemp = Nothing
    ShowErrorMsg
    
End Sub

Private Sub cmdok_Click()
    
    On Error GoTo Error_Handle
    '生成记录集

    Dim oRouteDim As New TicketRouteDim
    Dim rsArea1 As Recordset
    Dim rsArea2 As Recordset
    Dim rsTemp As New Recordset
    Dim i As Integer
    
    
    
    
    oRouteDim.Init m_oActiveUser
    Me.MousePointer = vbHourglass
    
    If MDIMain.m_szMethod = False Then
        Set rsArea1 = oRouteDim.GetRouteTransport(dtpBeginDate.Value, dtpEndDate.Value, m_szCode, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)), szStatus)
    Else
        Set rsArea1 = oRouteDim.GetRouteTransportByCheck(dtpBeginDate.Value, dtpEndDate.Value, m_szCode, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)), szStatus)
    End If
    
'    Set rsArea2 = oRouteDim.GetRouteTransportBusNum(dtpBeginDate.Value, dtpEndDate.Value, m_szCode)
'    If rsBusType2.RecordCount = 0 Then Exit Sub
'    If rsArea1.RecordCount <> 0 And rsArea2.RecordCount <> 0 Then
'        Set rsTemp = MergeArea(rsArea1, rsArea2)
'    ElseIf rsArea1.RecordCount <> 0 Then
        Set rsTemp = rsArea1
'    End If

    
    Set m_rsData = rsTemp
    ReDim m_vaCustomData(1 To 5, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY-MM-DD")
    
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY-MM-DD")
    
    m_vaCustomData(3, 1) = "线路"
    m_vaCustomData(3, 2) = IIf(txtRouteID.Text <> "", txtRouteID.Text, "所有线路")
    
    m_vaCustomData(4, 1) = "统计方式"
    If MDIMain.m_szMethod = False Then
        m_vaCustomData(4, 2) = "按售票统计"
    Else
        m_vaCustomData(4, 2) = "按检票统计"
    End If
    
    m_vaCustomData(5, 1) = "制表人"
    m_vaCustomData(5, 2) = m_oActiveUser.UserID
    m_bOk = True
    Me.MousePointer = vbDefault
    Unload Me
    Exit Sub
    
Error_Handle:
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub


Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
    m_szCode = ""
    m_bOk = False
    On Error GoTo Here
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = DateAdd("d", -1, dyNow)
    dtpEndDate.Value = DateAdd("d", -1, dyNow)
    
    FillSellStation cboSellStation
    
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If
    szStatus = False
    Exit Sub
Here:
    ShowMsg err.Description
End Sub

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    If szStatus = False Then
        IConditionForm_FileName = cszFileName1
    Else
        IConditionForm_FileName = cszFileName2
    End If
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property


Private Sub txtRouteID_ButtonClick()
    Dim aszRoute() As String
    aszRoute = m_oShell.SelectRoute(True)
    txtRouteID.Text = TeamToString(aszRoute, 2)

    m_szCode = TeamToString(aszRoute, 1)

End Sub


'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub

