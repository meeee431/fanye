VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCompanyTranSimply 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "公司正常加班比较表"
   ClientHeight    =   4350
   ClientLeft      =   2220
   ClientTop       =   3720
   ClientWidth     =   6555
   HelpContextID   =   60000100
   Icon            =   "frmCompanyTranSimply.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   1290
      TabIndex        =   21
      Top             =   3030
      Width           =   1605
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   2190
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
      MICON           =   "frmCompanyTranSimply.frx":000C
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
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2595
      Width           =   1635
   End
   Begin VB.ComboBox cboBusSection 
      Height          =   300
      Left            =   4410
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2100
      Width           =   1905
   End
   Begin VB.OptionButton optCombine 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按车次段汇总"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   1020
      Width           =   1455
   End
   Begin VB.OptionButton optCompany 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按拆帐公司汇总"
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1605
   End
   Begin VB.ComboBox cboExtraStatus 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2115
      Width           =   1635
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   5070
      TabIndex        =   2
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
      MICON           =   "frmCompanyTranSimply.frx":0028
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
      Default         =   -1  'True
      Height          =   345
      Left            =   3630
      TabIndex        =   1
      Top             =   3900
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmCompanyTranSimply.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   0
      Top             =   660
      Width           =   6885
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   1290
      TabIndex        =   9
      Top             =   1545
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1290
      TabIndex        =   10
      Top             =   990
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin FText.asFlatTextBox txtTransportCompanyID 
      Height          =   300
      Left            =   4410
      TabIndex        =   11
      Top             =   2100
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -30
      ScaleHeight     =   735
      ScaleWidth      =   6615
      TabIndex        =   3
      Top             =   -30
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   4
         Top             =   270
         Width           =   990
      End
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
      Left            =   225
      TabIndex        =   19
      Top             =   2655
      Width           =   900
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(T):"
      Height          =   180
      Left            =   3120
      TabIndex        =   17
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Label lblCombine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次段序号(&R):"
      Height          =   180
      Left            =   3120
      TabIndex        =   16
      Top             =   2145
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "补票状态(&S):"
      Height          =   180
      Left            =   210
      TabIndex        =   15
      Top             =   2145
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   210
      TabIndex        =   14
      Top             =   1605
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   210
      TabIndex        =   13
      Top             =   1050
      Width           =   1080
   End
End
Attribute VB_Name = "frmCompanyTranSimply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IConditionForm
Const cszFileName = "公司运量统计简报模板.xls"

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

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset

'    Dim rsData As New Recordset
    Dim i As Integer
    
    Dim aszTemp() As String
    Dim szStation As String
    Dim nCount As Integer
    
'    Dim rsBus As Recordset
'    If optCompany.Value Then
'        Set rsTemp = oDss.GetCompanyFloatSimpleCon(txtTransportCompanyID.Text, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text))
'    Else
'        Set rsTemp = oDss.GetCombineFloatSimpleCon(dtpBeginDate.Value, dtpEndDate.Value, cboBusSection.Text, ResolveDisplay(cboExtraStatus.Text))
'    End If
    
    If optCompany.Value Then
        Set rsTemp = oDss.GetCompanyFloatSimpleCon(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
'        Set rsBus = oDss.GetBusNumByCompany(m_szCode, dtpBeginDate.Value, dtpEndDate.Value)
    Else
        Set rsTemp = oDss.GetCombineFloatSimpleCon(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
'        Set rsBus = oDss.GetBusNumByCompanySerial(dtpBeginDate.Value, dtpEndDate.Value, cboBusSection.Text)
    End If
    Set m_rsData = rsTemp
'    If rsTemp.RecordCount <> 0 And rsBus.RecordCount <> 0 Then
'        Set m_rsData = MergeRecordset(rsTemp, rsBus)
'    Else
'        Set m_rsData = rsTemp
'    End If
    ReDim m_vaCustomData(1 To 6, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "补票状态"
    m_vaCustomData(3, 2) = cboExtraStatus.Text
    Dim oStationDim As New TicketStationDim
    oStationDim.Init m_oActiveUser
    
    aszTemp = oStationDim.GetStationByParam(dtpBeginDate.Value, dtpEndDate.Value)
    nCount = ArrayLength(aszTemp)
    For i = 1 To nCount
        szStation = szStation & "   " & aszTemp(i, 2) & ":" & aszTemp(i, 3) & "人"
        
    Next i
    
    m_vaCustomData(4, 1) = "站点人数"
    m_vaCustomData(4, 2) = szStation
    
    Dim szSellStation As String
    ResolveDisplay cboSellStation, szSellStation
    m_vaCustomData(5, 1) = "上车站"
    m_vaCustomData(5, 2) = szSellStation
    
    
    m_vaCustomData(6, 1) = "制表人"
    m_vaCustomData(6, 2) = m_oActiveUser.UserID
    m_bOk = True
    Set oDss = Nothing
    Set oStationDim = Nothing
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




