VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCompanyFloatSimply 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "公司运量统计简报"
   ClientHeight    =   4365
   ClientLeft      =   4170
   ClientTop       =   3615
   ClientWidth     =   6510
   HelpContextID   =   6000070
   Icon            =   "frmCompanyFloatSimply.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   1320
      TabIndex        =   22
      Top             =   3030
      Width           =   1605
   End
   Begin RTComctl3.CoolButton cmdChart 
      Height          =   315
      Left            =   570
      TabIndex        =   21
      Top             =   3900
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
      MICON           =   "frmCompanyFloatSimply.frx":000C
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
      Left            =   2280
      TabIndex        =   20
      Top             =   3900
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
      MICON           =   "frmCompanyFloatSimply.frx":0028
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2565
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   0
      Top             =   660
      Width           =   6885
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   345
      Left            =   3720
      TabIndex        =   8
      Top             =   3900
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
      MICON           =   "frmCompanyFloatSimply.frx":0044
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
      Left            =   5100
      TabIndex        =   7
      Top             =   3900
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
      MICON           =   "frmCompanyFloatSimply.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -30
      ScaleHeight     =   735
      ScaleWidth      =   6615
      TabIndex        =   5
      Top             =   -30
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.ComboBox cboExtraStatus 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2030
      Width           =   1635
   End
   Begin VB.OptionButton optCompany 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按拆帐公司汇总"
      Height          =   285
      Left            =   3150
      TabIndex        =   3
      Top             =   1503
      Width           =   1605
   End
   Begin VB.OptionButton optCombine 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按车次段汇总"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   968
      Width           =   1455
   End
   Begin VB.ComboBox cboBusSection 
      Height          =   300
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2070
      Width           =   1905
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   1320
      TabIndex        =   9
      Top             =   1495
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1320
      TabIndex        =   10
      Top             =   945
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin FText.asFlatTextBox txtTransportCompanyID 
      Height          =   300
      Left            =   4440
      TabIndex        =   11
      Top             =   2070
      Width           =   1905
      _ExtentX        =   3360
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
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   12
      Top             =   3660
      Width           =   8745
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   2625
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   240
      TabIndex        =   17
      Top             =   1005
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   1555
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "补票状态(&S):"
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   2115
      Width           =   1080
   End
   Begin VB.Label lblCombine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次段序号(&R):"
      Height          =   180
      Left            =   3150
      TabIndex        =   14
      Top             =   2115
      Width           =   1260
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(T):"
      Height          =   180
      Left            =   3150
      TabIndex        =   13
      Top             =   2130
      Width           =   1080
   End
End
Attribute VB_Name = "frmCompanyFloatSimply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IConditionForm
Const cszFileName = "公司流量汇总简报模板.xls"

Private m_rsData As Recordset
Public m_bOk As Boolean
Private m_vaCustomData As Variant
Dim m_aszTemp() As String
Dim oDss As New TicketCompanyDim
Dim m_szCode As String

Public Function MergeRecordset(prsBusType1 As Recordset, prsBusType2 As Recordset) As Recordset


    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    With rsTemp.Fields
        .Append "split_company_id", adChar, 12
        .Append "split_company_name", adChar, 10
        .Append "passenger_number", adInteger
        .Append "bus_count", adInteger
        .Append "fact_float", adDouble
        .Append "total_float", adDouble
        .Append "total_number", adInteger
        .Append "full_seat_rate", adDouble
        .Append "fact_load_rate", adDouble
        .Append "Add_bus_count", adInteger
        .Append "add_passenger_number", adInteger
        .Append "add_fact_float", adDouble
        .Append "add_total_float", adDouble
        .Append "add_total_number", adInteger
        .Append "add_full_seat_rate", adDouble
        .Append "add_fact_load_rate", adDouble

    End With
    rsTemp.Open
    If Not (prsBusType1 Is Nothing) Then
    prsBusType1.MoveFirst
    For i = 1 To prsBusType1.RecordCount
        rsTemp.AddNew
        rsTemp!split_company_id = prsBusType1!split_company_id
        rsTemp!split_company_name = prsBusType1!split_company_name
        rsTemp!passenger_number = prsBusType1!passenger_number
        rsTemp!bus_count = 0
        rsTemp!fact_float = prsBusType1!fact_float
        rsTemp!total_float = prsBusType1!total_float
        rsTemp!total_number = prsBusType1!total_number
        rsTemp!full_seat_rate = prsBusType1!full_seat_rate
        rsTemp!fact_load_rate = prsBusType1!fact_load_rate
        rsTemp!add_bus_count = 0
        rsTemp!add_passenger_number = prsBusType1!add_passenger_number
        rsTemp!add_fact_float = prsBusType1!add_fact_float
        
        rsTemp!add_total_float = prsBusType1!add_total_float
        rsTemp!add_total_number = prsBusType1!add_total_number
        rsTemp!add_full_seat_rate = prsBusType1!add_full_seat_rate
        rsTemp!add_fact_load_rate = prsBusType1!add_fact_load_rate
        
        

        prsBusType1.MoveNext
    Next i
    rsTemp.Update
    '赋记录集
    End If
    nCount = rsTemp.RecordCount
    If Not (prsBusType2 Is Nothing) Then
        If prsBusType2.RecordCount > 0 Then prsBusType2.MoveFirst
        For i = 1 To prsBusType2.RecordCount
            
            rsTemp.MoveFirst
            For j = 1 To nCount
                If prsBusType2!split_company_id = rsTemp!split_company_id Then Exit For
                rsTemp.MoveNext
                
            Next j
            If j > nCount Then
                rsTemp.AddNew
                rsTemp!split_company_id = prsBusType2!split_company_id
                rsTemp!split_company_name = prsBusType2!split_company_name
                rsTemp!passenger_number = 0
                rsTemp!bus_count = prsBusType2!bus_count
                rsTemp!fact_float = 0
                rsTemp!total_float = 0
                rsTemp!total_number = 0
                rsTemp!full_seat_rate = 0
                rsTemp!fact_load_rate = 0
                rsTemp!add_bus_count = prsBusType2!add_bus_count
                rsTemp!add_passenger_number = 0
                rsTemp!add_fact_float = 0
                rsTemp!add_total_float = 0
                rsTemp!add_total_number = 0
                rsTemp!add_full_seat_rate = 0
                rsTemp!add_fact_load_rate = 0
        
        
            Else
                
                '已经存在
                rsTemp!bus_count = prsBusType2!bus_count
                rsTemp!add_bus_count = prsBusType2!add_bus_count
                
            End If
            prsBusType2.MoveNext
        Next i
        If Not (rsTemp Is Nothing) Then rsTemp.Update
    
    End If
    
    Set MergeRecordset = rsTemp
    

End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChart_Click()
    
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim frmTemp As frmChart
    If MDIMain.m_szMethod = False Then
        If optCompany.Value Then
            Set rsTemp = oDss.GetCompanyFloatSimpleCon(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        Else
            Set rsTemp = oDss.GetCombineFloatSimpleCon(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        End If
    Else
        If optCompany.Value Then
            Set rsTemp = oDss.GetCompanyFloatSimpleConCheck(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        Else
            Set rsTemp = oDss.GetCombineFloatSimpleConCheck(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        End If
    End If
    
    Dim rsData As New Recordset
    With rsData.Fields
        .Append "split_company_id", adBSTR
        .Append "bus_count", adBigInt
    End With
    rsData.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsData.AddNew
        rsData!split_company_id = FormatDbValue(rsTemp!split_company_id)
        rsData!bus_count = FormatDbValue(rsTemp!bus_count)
        rsTemp.MoveNext
        rsData.Update
    Next i
    
    Dim rsdata2 As New Recordset
    With rsdata2.Fields
        .Append "split_company_id", adBSTR
        .Append "passenger_number", adBigInt
    End With
    rsdata2.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata2.AddNew
        rsdata2!split_company_id = FormatDbValue(rsTemp!split_company_id)
        rsdata2!passenger_number = FormatDbValue(rsTemp!passenger_number)
        rsTemp.MoveNext
        rsdata2.Update
    Next i
    
    Dim rsdata3 As New Recordset
    With rsdata3.Fields
        .Append "split_company_id", adBSTR
        .Append "fact_float", adBigInt
    End With
    rsdata3.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata3.AddNew
        rsdata3!split_company_id = FormatDbValue(rsTemp!split_company_id)
        rsdata3!fact_float = FormatDbValue(rsTemp!fact_float)
        rsTemp.MoveNext
        rsdata3.Update
    Next i
    
    Dim rsdata4 As New Recordset
    With rsdata4.Fields
        .Append "split_company_id", adBSTR
        .Append "total_float", adBigInt
    End With
    rsdata4.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata4.AddNew
        rsdata4!split_company_id = FormatDbValue(rsTemp!split_company_id)
        rsdata4!total_float = FormatDbValue(rsTemp!total_float)
        rsTemp.MoveNext
        rsdata4.Update
    Next i
    
    Dim rsdata5 As New Recordset
    With rsdata5.Fields
        .Append "split_company_id", adBSTR
        .Append "full_seat_rate", adBigInt
    End With
    rsdata5.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata5.AddNew
        rsdata5!split_company_id = FormatDbValue(rsTemp!split_company_id)
        rsdata5!full_seat_rate = FormatDbValue(rsTemp!full_seat_rate)
        rsTemp.MoveNext
        rsdata5.Update
    Next i
   
      Dim rsdata6 As New Recordset
    With rsdata6.Fields
        .Append "split_company_id", adBSTR
        .Append "fact_load_rate", adBigInt
    End With
    rsdata6.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata6.AddNew
        rsdata6!split_company_id = FormatDbValue(rsTemp!split_company_id)
        rsdata6!fact_load_rate = FormatDbValue(rsTemp!fact_load_rate)
        rsTemp.MoveNext
        rsdata6.Update
    Next i
    
    Me.Hide
    Set frmTemp = New frmChart
    frmTemp.ClearChart
    frmTemp.AddChart "总趟次", rsData
    frmTemp.AddChart "人数", rsdata2
    frmTemp.AddChart "实际周转量", rsdata3
    frmTemp.AddChart "总周转量", rsdata4
    frmTemp.AddChart "上座率", rsdata5
    frmTemp.AddChart "实载率", rsdata6
    frmTemp.ShowChart "公司流量汇总简报"
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
    Dim rsTemp As Recordset
'    Dim rsBus As Recordset
'    Dim rsData As New Recordset
    Dim i As Integer
    If MDIMain.m_szMethod = False Then
        If optCompany.Value Then
            Set rsTemp = oDss.GetCompanyFloatSimpleCon(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    '        Set rsBus = oDss.GetBusNumByCompany(m_szCode, dtpBeginDate.Value, dtpEndDate.Value)
        Else
            Set rsTemp = oDss.GetCombineFloatSimpleCon(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    '        Set rsBus = oDss.GetBusNumByCompanySerial(dtpBeginDate.Value, dtpEndDate.Value, cboBusSection.Text)
        End If
    Else
        If optCompany.Value Then
            Set rsTemp = oDss.GetCompanyFloatSimpleConCheck(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        Else
            Set rsTemp = oDss.GetCombineFloatSimpleConCheck(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        End If
    End If
'    If rsTemp.RecordCount <> 0 And rsBus.RecordCount <> 0 Then
'        Set m_rsData = MergeRecordset(rsTemp, rsBus)
'    Else
'        Set m_rsData = rsTemp
'    End If
    
    Set m_rsData = rsTemp
    
    ReDim m_vaCustomData(1 To 6, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "补票状态"
    m_vaCustomData(3, 2) = cboExtraStatus.Text
    
    Dim szSellStation As String
    ResolveDisplay cboSellStation, szSellStation
    m_vaCustomData(4, 1) = "上车站"
    m_vaCustomData(4, 2) = szSellStation
    
    m_vaCustomData(5, 1) = "统计方式"
    If MDIMain.m_szMethod = False Then
        m_vaCustomData(5, 2) = "按售票统计"
    Else
        m_vaCustomData(5, 2) = "按检票统计"
    End If
    
    m_vaCustomData(6, 1) = "制表人"
    m_vaCustomData(6, 2) = m_oActiveUser.UserID
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub


Private Sub CoolButton1_Click()
DisplayHelp Me

End Sub

Private Sub Form_Load()
    oDss.Init m_oActiveUser
    m_szCode = ""
    m_bOk = False
    FillCombine
'    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
'    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    cboExtraStatus.AddItem "1[售票]"
    cboExtraStatus.AddItem "2[补票]"
    cboExtraStatus.AddItem "3[售票+补票]"
    
    cboExtraStatus.ListIndex = 2
    
    optCompany.Value = True
    SetVisible False
    FillSellStation cboSellStation
    
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If
    
    
End Sub

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property



Private Sub optCombine_Click()
    SetVisible True
End Sub

Private Sub optCompany_Click()
    SetVisible False
    
End Sub

Private Sub txtTransportCompanyID_ButtonClick()
    Dim aszTransportCompanyID() As String
    aszTransportCompanyID = m_oShell.SelectCompany
    txtTransportCompanyID.Text = TeamToString(aszTransportCompanyID, 2)
    
    m_szCode = TeamToString(aszTransportCompanyID, 1)
    
End Sub

Private Sub FillCombine()
    '填充唯一的车次序号
    Dim aszTemp() As String
    Dim i As Integer
    Dim nCount As Integer
    
    aszTemp = oDss.GetUniqueCombine
    nCount = ArrayLength(aszTemp)
    For i = 1 To nCount
        cboBusSection.AddItem aszTemp(i)
    Next i
    If cboBusSection.ListCount > 0 Then cboBusSection.ListIndex = 0
    
End Sub


Private Sub SetVisible(pbVisible As Boolean)
    lblCombine.Visible = pbVisible
    cboBusSection.Visible = pbVisible
    lblCompany.Visible = Not pbVisible
    txtTransportCompanyID.Visible = Not pbVisible
    
End Sub


'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub


