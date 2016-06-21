VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmRouteTurnOver 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "线路营收"
   ClientHeight    =   4440
   ClientLeft      =   2550
   ClientTop       =   2250
   ClientWidth     =   6195
   HelpContextID   =   6000001
   Icon            =   "frmRouteTurnOver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   2220
      TabIndex        =   17
      Top             =   2310
      Width           =   2895
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   1740
      TabIndex        =   16
      Top             =   3960
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
      MICON           =   "frmRouteTurnOver.frx":000C
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
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1935
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   12
      Top             =   0
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   13
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   10
      Top             =   690
      Width           =   6885
   End
   Begin FText.asFlatTextBox txtAreaCode 
      Height          =   300
      Left            =   2220
      TabIndex        =   7
      Top             =   3195
      Width           =   2895
      _ExtentX        =   5106
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
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4770
      TabIndex        =   9
      Top             =   3960
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
      MICON           =   "frmRouteTurnOver.frx":0028
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
      Left            =   3360
      TabIndex        =   8
      Top             =   3990
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
      MICON           =   "frmRouteTurnOver.frx":0044
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
      Left            =   2220
      TabIndex        =   3
      Top             =   1485
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   2220
      TabIndex        =   1
      Top             =   990
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin FText.asFlatTextBox txtTransportCompanyID 
      Height          =   300
      Left            =   2220
      TabIndex        =   5
      Top             =   2700
      Width           =   2895
      _ExtentX        =   5106
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
      Left            =   -120
      TabIndex        =   11
      Top             =   3720
      Width           =   8745
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   960
      TabIndex        =   15
      Top             =   1995
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地区代码(&A):"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   3225
      Width           =   1080
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(&T):"
      Height          =   180
      Left            =   960
      TabIndex        =   4
      Top             =   2745
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   1530
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   1050
      Width           =   1080
   End
End
Attribute VB_Name = "frmRouteTurnOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm
Const cszFileName = "线路售票营收简报模板.xls"
Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant
Dim m_szCode As String
Dim m_szCode2 As String


Private m_aszSum(1 To 38) As Double '总和

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub EvalRecordSet(prsTemp As Recordset, prsArea As Recordset)
    Dim szArea As String
    Dim szAreaName As String
    Dim adbTemp(1 To 38) As Double
    Dim i As Integer
    Dim j As Integer
    
    Dim szStation As String
    Dim aszStation() As String
    Dim nCount As Integer
    
'    Dim szInterProvince As String
'    Dim szInCity As String
'    Dim szExtramural As String
    Dim szSum1 As String
    
    Dim szSubTotalItem As String
    Dim aszSubTotalItem() As String
    Dim nSubTotalItem As Integer
    '取得小计项参数
    szSubTotalItem = m_oParam.SubTotalItem1
    aszSubTotalItem = StringToTeam(szSubTotalItem)
    nSubTotalItem = ArrayLength(aszSubTotalItem)
    For i = 1 To 38
        m_aszSum(i) = 0
    Next i
'
'
'    szInterProvince = m_oParam.InterProvinceStationDetail
'
'    szInCity = m_oParam.InCityStationDetail
'
'    szExtramural = m_oParam.ExtramuralStationDetail
'
    If prsArea.RecordCount > 0 Then prsArea.MoveFirst
    Do While Not prsArea.EOF
        '如果车次种类已经一类结束则插入合计
        

        If szArea <> "" And prsArea!province_in_out <> szArea Then
            prsTemp.AddNew
            prsTemp!province_in_out = szAreaName
            prsTemp!end_station_id = "合计"
            prsTemp!station_name = "合计"
            prsTemp!passenger_number = adbTemp(1)
            prsTemp!total_ticket_price = adbTemp(2)
            prsTemp!ticket_return_charge = adbTemp(3)
            prsTemp!ticket_change_charge = adbTemp(4)
            prsTemp!base_price = adbTemp(5)
            For i = 1 To 15
                prsTemp("price_item_" & i) = adbTemp(5 + i)
            Next i
            
            prsTemp!some_sum = adbTemp(21)
            prsTemp!some_sum2 = adbTemp(22)
            
            For i = 1 To nSubTotalItem
                If aszSubTotalItem(i) = 0 Then
                    prsTemp!out_base_price = adbTemp(22 + i)
                Else
                    prsTemp("out_price_item_" & aszSubTotalItem(i)) = adbTemp(22 + i)
                    
                End If
            Next i
            '为总计用
            For i = 1 To 22 + nSubTotalItem
                m_aszSum(i) = m_aszSum(i) + adbTemp(i)
                adbTemp(i) = 0
            Next i
            
            
        End If
        '赋值
        
'        Select Case prsArea!province_in_out
'        Case "省外" 'EA_nOutProvince '省外
'            szStation = szInterProvince
'        Case "市内" 'EA_nInCity '市内
'            szStation = szInCity
'        Case "市外" 'EA_nOutCity '市外
'            szStation = szExtramural
'        End Select
'        aszStation = StringToTeam(szStation)
'        nCount = ArrayLength(aszStation)
'        For j = 1 To nCount
'            If RTrim(aszStation(j)) = RTrim(prsArea!end_station_id) Then
        prsTemp.AddNew
        prsTemp!province_in_out = prsArea!province_in_out
        prsTemp!end_station_id = prsArea!end_station_id
        prsTemp!station_name = prsArea!station_name
        prsTemp!passenger_number = prsArea!passenger_number
        prsTemp!total_ticket_price = prsArea!total_ticket_price
        prsTemp!ticket_return_charge = prsArea!ticket_return_charge
        prsTemp!ticket_change_charge = prsArea!ticket_change_charge
        prsTemp!base_price = prsArea!base_price
        For i = 1 To 15
            prsTemp("price_item_" & i) = prsArea("price_item_" & i)
        Next i
        
        prsTemp!some_sum = prsArea!some_sum
        prsTemp!some_sum2 = prsArea!some_sum2
        
        For i = 1 To nSubTotalItem
            If aszSubTotalItem(i) = 0 Then
                prsTemp!out_base_price = prsArea!out_base_price
            Else
                prsTemp("out_price_item_" & aszSubTotalItem(i)) = prsArea("out_price_item_" & aszSubTotalItem(i))
                
            End If
        Next i
            
'            End If
'        Next j
        '为计算合计用

        adbTemp(1) = adbTemp(1) + prsArea!passenger_number
        adbTemp(2) = adbTemp(2) + prsArea!total_ticket_price
        adbTemp(3) = adbTemp(3) + prsArea!ticket_return_charge
        adbTemp(4) = adbTemp(4) + prsArea!ticket_change_charge
        adbTemp(5) = adbTemp(5) + prsArea!base_price
        For i = 1 To 15
            adbTemp(5 + i) = adbTemp(5 + i) + prsArea("price_item_" & i)
        Next i
        
        adbTemp(21) = adbTemp(21) + prsArea!some_sum
        adbTemp(22) = adbTemp(22) + prsArea!some_sum2
        
        For i = 1 To nSubTotalItem
            If aszSubTotalItem(i) = 0 Then
                adbTemp(22 + i) = adbTemp(22 + i) + prsArea!out_base_price
            Else
                adbTemp(22 + i) = adbTemp(22 + i) + prsArea("out_price_item_" & Trim(aszSubTotalItem(i)))
                
            End If
        Next i
        
        szArea = prsArea!province_in_out
        szAreaName = Trim(prsArea!province_in_out)
        prsArea.MoveNext
    Loop
    '将最后一条记录进行赋值
    If prsArea.RecordCount > 0 Then
        '赋值
        prsTemp.AddNew
        prsTemp!province_in_out = szAreaName
        prsTemp!end_station_id = "合计"
        prsTemp!station_name = "合计"
        prsTemp!passenger_number = adbTemp(1)
        prsTemp!total_ticket_price = adbTemp(2)
        prsTemp!ticket_return_charge = adbTemp(3)
        prsTemp!ticket_change_charge = adbTemp(4)
        prsTemp!base_price = adbTemp(5)
        For i = 1 To 15
            prsTemp("price_item_" & i) = adbTemp(5 + i)
        Next i
        
        prsTemp!some_sum = adbTemp(21)
        prsTemp!some_sum2 = adbTemp(22)
        
        For i = 1 To nSubTotalItem
            If aszSubTotalItem(i) = 0 Then
                prsTemp!out_base_price = adbTemp(22 + i)
            Else
                prsTemp("out_price_item_" & aszSubTotalItem(i)) = adbTemp(22 + i)
                
            End If
        Next i
        '为总计用
        For i = 1 To 22 + nSubTotalItem
            m_aszSum(i) = m_aszSum(i) + adbTemp(i)
            adbTemp(i) = 0
        Next i
        
        prsTemp.AddNew
        prsTemp!province_in_out = "总计" 'Trim(prsArea!province_in_out) & "普通班车"
        prsTemp!end_station_id = ""
        prsTemp!station_name = ""
        prsTemp!passenger_number = m_aszSum(1)
        prsTemp!total_ticket_price = m_aszSum(2)
        prsTemp!ticket_return_charge = m_aszSum(3)
        prsTemp!ticket_change_charge = m_aszSum(4)
        prsTemp!base_price = m_aszSum(5)
        For i = 1 To 15
            prsTemp("price_item_" & i) = m_aszSum(5 + i)
        Next i
        
        prsTemp!some_sum = m_aszSum(21)
        prsTemp!some_sum2 = m_aszSum(22)
        
        For i = 1 To nSubTotalItem
            If aszSubTotalItem(i) = 0 Then
                prsTemp!out_base_price = m_aszSum(22 + i)
            Else
                prsTemp("out_price_item_" & aszSubTotalItem(i)) = m_aszSum(22 + i)
                
            End If
        Next i
        prsTemp.Update
    End If
End Sub



Private Sub cmdok_Click()
    
    On Error GoTo Error_Handle
    '生成记录集

    Dim oBusDim As New TicketRouteDim
    Dim rsArea As Recordset
    Dim rsTemp As New Recordset
    Dim i As Integer
    
    
    
    
    oBusDim.Init m_oActiveUser
    Me.MousePointer = vbHourglass
    
    '得到线路营收及其拆出等明细
    Set rsArea = oBusDim.GetRouteTurnOver(dtpBeginDate.Value, dtpEndDate.Value, m_szCode, m_szCode2, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    'province_in_out end_station_id station_name passenger_number total_ticket_price    ticket_return_charge  ticket_change_charge
    'base_price            price_item_1          price_item_2          price_item_3          price_item_4          price_item_5          price_item_6          price_item_7          price_item_8          price_item_9          price_item_10         price_item_11         price_item_12         price_item_13         price_item_14         price_item_15
    'some_sum2                     Some_Sum                     out_base_price          out_price_item_1      out_price_item_7      out_price_item_10     out_price_item_13     out_price_item_15
    With rsTemp.Fields
        .Append "province_in_out", adChar, 4
        .Append "end_station_id", adChar, 9
        .Append "station_name", adChar, 10
        .Append "passenger_number", adInteger
        .Append "total_ticket_price", adCurrency
        .Append "ticket_return_charge", adCurrency
        .Append "ticket_change_charge", adCurrency
        .Append "base_price", adCurrency
        For i = 1 To 15
            .Append "price_item_" & i, adCurrency
        Next i
        .Append "Some_Sum", adCurrency
        .Append "some_sum2", adCurrency
        .Append "out_base_price", adCurrency
        For i = 1 To 15
            .Append "out_price_item_" & i, adCurrency
        Next i

    End With
    rsTemp.Open
    EvalRecordSet rsTemp, rsArea
    
    Set m_rsData = rsTemp
    ReDim m_vaCustomData(1 To 5, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY-MM-DD")
    
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY-MM-DD")
    
    m_vaCustomData(3, 1) = "拆帐公司"
    m_vaCustomData(3, 2) = IIf(txtTransportCompanyID.Text <> "", txtTransportCompanyID.Text, "所有公司")
    
    m_vaCustomData(4, 1) = "地区"
    m_vaCustomData(4, 2) = IIf(txtAreaCode.Text <> "", txtAreaCode.Text, "所有地区")
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
    m_bOk = False
    m_szCode = ""
    m_szCode2 = ""
    
    On Error GoTo Here
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    FillSellStation cboSellStation
    
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If
    Exit Sub
Here:
    ShowMsg err.Description
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


Private Sub txtAreaCode_ButtonClick()
    Dim aszArea() As String
    aszArea = m_oShell.SelectArea
    txtAreaCode.Text = TeamToString(aszArea, 2)

    m_szCode2 = TeamToString(aszArea, 1)
    
End Sub

Private Sub txtTransportCompanyID_ButtonClick()
    Dim aszTransportCompanyID() As String
    aszTransportCompanyID = m_oShell.SelectCompany
    txtTransportCompanyID.Text = TeamToString(aszTransportCompanyID, 2)

    m_szCode = TeamToString(aszTransportCompanyID, 1)
    
End Sub


'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub
'
