VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAreaRouteSimply 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "地区线路统计"
   ClientHeight    =   3735
   ClientLeft      =   4860
   ClientTop       =   3615
   ClientWidth     =   4740
   HelpContextID   =   6000010
   Icon            =   "frmAreaRouteSimply.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   1890
      TabIndex        =   13
      Top             =   2550
      Width           =   2115
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   660
      TabIndex        =   12
      Top             =   3330
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      MICON           =   "frmAreaRouteSimply.frx":000C
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
      Height          =   315
      Left            =   3330
      TabIndex        =   5
      Top             =   3330
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
      MICON           =   "frmAreaRouteSimply.frx":0028
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
      ItemData        =   "frmAreaRouteSimply.frx":0044
      Left            =   1890
      List            =   "frmAreaRouteSimply.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2115
      Width           =   2115
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   1980
      TabIndex        =   4
      Top             =   3330
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      MICON           =   "frmAreaRouteSimply.frx":0048
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -90
      TabIndex        =   7
      Top             =   690
      Width           =   6885
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   1890
      TabIndex        =   3
      Top             =   1597
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   1080
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      HelpContextID   =   6000010
      Left            =   -120
      TabIndex        =   6
      Top             =   3000
      Width           =   8745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   585
      TabIndex        =   10
      Top             =   2175
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   585
      TabIndex        =   2
      Top             =   1664
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   585
      TabIndex        =   0
      Top             =   1147
      Width           =   1080
   End
End
Attribute VB_Name = "frmAreaRouteSimply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm
Const cszFileName = "地区线路简报模板.xls"
Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant



Private m_aszSum(1 To 8) As Double '总和
Private m_nKinds As Integer '需求平均的个数

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Function MergeBusType(prsBusType1 As Recordset, prsBusType2 As Recordset) As Recordset
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    With rsTemp.Fields
        .Append "bus_type", adInteger
        .Append "bus_type_name", adChar, 20
        .Append "route_id", adChar, 4
        .Append "Route_name", adChar, 16
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
    If Not (prsBusType1 Is Nothing) Then
    prsBusType1.MoveFirst
    For i = 1 To prsBusType1.RecordCount
        rsTemp.AddNew
        rsTemp!bus_type = prsBusType1!bus_type
        rsTemp!bus_type_name = prsBusType1!bus_type_name
        rsTemp!route_id = prsBusType1!route_id
        rsTemp!route_name = prsBusType1!route_name
        rsTemp!bus_count = 0 'prsBusType1!bus_count
        rsTemp!passenger_number = prsBusType1!passenger_number
        rsTemp!fact_float = prsBusType1!fact_float
        rsTemp!total_float = prsBusType1!total_float
        rsTemp!total_number = prsBusType1!total_number
        rsTemp!full_seat_rate = prsBusType1!full_seat_rate
        rsTemp!fact_load_rate = prsBusType1!fact_load_rate
        rsTemp!total_ticket_price = prsBusType1!total_ticket_price
        

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
                If prsBusType2!route_id = rsTemp!route_id And prsBusType2!bus_type = rsTemp!bus_type Then Exit For
                rsTemp.MoveNext
                
            Next j
            If j > nCount Then
                rsTemp.AddNew
                rsTemp!bus_type = prsBusType2!bus_type
                rsTemp!bus_type_name = prsBusType2!bus_type_name
                rsTemp!route_id = prsBusType2!route_id
                rsTemp!route_name = prsBusType2!route_name
                rsTemp!bus_count = prsBusType2!fact_Bus_number
    
                rsTemp!passenger_number = 0
                rsTemp!fact_float = 0
                rsTemp!total_float = 0
                rsTemp!total_number = 0
                rsTemp!full_seat_rate = 0
                rsTemp!fact_load_rate = 0
                rsTemp!total_ticket_price = 0
            Else
                
                '已经存在
                rsTemp!bus_count = prsBusType2!fact_Bus_number
                
            End If
            prsBusType2.MoveNext
        Next i
        If Not (rsTemp Is Nothing) Then rsTemp.Update
    
    End If
    
    Set MergeBusType = rsTemp
    
End Function



Private Function MergeArea(prsArea1 As Recordset, prsArea2 As Recordset) As Recordset
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    
    
    With rsTemp.Fields
        .Append "province_in_out", adChar, 4
        .Append "end_station_id", adChar, 9
        .Append "station_name", adChar, 16
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
                rsTemp!province_in_out = prsArea1!province_in_out
                rsTemp!end_station_id = prsArea1!end_station_id
                rsTemp!station_name = prsArea1!station_name
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
                If prsArea2!province_in_out = rsTemp!province_in_out And prsArea2!end_station_id = rsTemp!end_station_id Then Exit For
                rsTemp.MoveNext
                
            Next j
            If j > nCount Then
                rsTemp.AddNew
                rsTemp!province_in_out = prsArea2!province_in_out
                rsTemp!end_station_id = prsArea2!end_station_id
                
                rsTemp!station_name = prsArea2!station_name
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


Private Sub EvalRecBusType(prsTemp As Recordset, prsBusType As Recordset)
    Dim szBusType As String
    Dim adbTemp(1 To 8) As Double
    Dim i As Integer
    Dim szBusTypeName As String
    Dim j As Integer
    
    If prsBusType.RecordCount > 0 Then prsBusType.MoveFirst
    Do While Not prsBusType.EOF
        '如果车次种类已经一类结束则插入合计
        If szBusType <> "" And prsBusType!bus_type <> szBusType Then
            prsTemp.AddNew
            prsTemp!Kinds = szBusTypeName
            prsTemp!route_id = "合计"
            prsTemp!route_name = "合计"
            prsTemp!bus_count = adbTemp(1)
            prsTemp!passenger_number = adbTemp(2)
            prsTemp!fact_float = adbTemp(3)
            prsTemp!total_float = adbTemp(4)
            prsTemp!total_number = adbTemp(5)
            prsTemp!full_seat_rate = adbTemp(6) / j
            prsTemp!fact_load_rate = adbTemp(7) / j
            prsTemp!total_ticket_price = adbTemp(8)

            
            For i = 1 To 8
                If i = 6 Or i = 7 Then
                    m_aszSum(i) = m_aszSum(i) + adbTemp(i) / j
                Else
                    m_aszSum(i) = m_aszSum(i) + adbTemp(i)
                End If
                adbTemp(i) = 0
            Next i
            j = 0
            m_nKinds = m_nKinds + 1
        End If
        '赋值
        prsTemp.AddNew
        prsTemp!Kinds = prsBusType!bus_type_name
        prsTemp!route_id = prsBusType!route_id
        prsTemp!route_name = prsBusType!route_name
        prsTemp!bus_count = prsBusType!bus_count
        prsTemp!passenger_number = prsBusType!passenger_number
        prsTemp!fact_float = prsBusType!fact_float
        prsTemp!total_float = prsBusType!total_float
        prsTemp!total_number = prsBusType!total_number
        prsTemp!full_seat_rate = prsBusType!full_seat_rate
        prsTemp!fact_load_rate = prsBusType!fact_load_rate
        prsTemp!total_ticket_price = prsBusType!total_ticket_price
        '为计算合计用
        j = j + 1
        adbTemp(1) = adbTemp(1) + prsBusType!bus_count
        adbTemp(2) = adbTemp(2) + prsBusType!passenger_number
        adbTemp(3) = adbTemp(3) + prsBusType!fact_float
        adbTemp(4) = adbTemp(4) + prsBusType!total_float
        adbTemp(5) = adbTemp(5) + prsBusType!total_number
        adbTemp(6) = adbTemp(6) + prsBusType!full_seat_rate
        adbTemp(7) = adbTemp(7) + prsBusType!fact_load_rate
        adbTemp(8) = adbTemp(8) + prsBusType!total_ticket_price
        szBusType = prsBusType!bus_type
        szBusTypeName = prsBusType!bus_type_name
        prsBusType.MoveNext
    Loop
    '将最后一条记录进行赋值
    If prsBusType.RecordCount > 0 Then
        '赋值
        prsTemp.AddNew
        prsTemp!Kinds = szBusTypeName
        prsTemp!route_id = "合计"
        prsTemp!route_name = "合计"
        prsTemp!bus_count = adbTemp(1)
        prsTemp!passenger_number = adbTemp(2)
        prsTemp!fact_float = adbTemp(3)
        prsTemp!total_float = adbTemp(4)
        prsTemp!total_number = adbTemp(5)
        prsTemp!full_seat_rate = adbTemp(6) / j
        prsTemp!fact_load_rate = adbTemp(7) / j
        prsTemp!total_ticket_price = adbTemp(8)
        
        m_nKinds = m_nKinds + 1
        For i = 1 To 8
            If i = 6 Or i = 7 Then
                m_aszSum(i) = m_aszSum(i) + adbTemp(i) / j
            Else
                m_aszSum(i) = m_aszSum(i) + adbTemp(i)
            End If
        Next i
'        For i = 1 To 8
'            adbTemp(i) = 0
'        Next i
    
        prsTemp.Update
'        prsTemp.Close
    End If
    
End Sub


Private Sub EvalRecArea(prsTemp As Recordset, prsArea As Recordset)
    Dim szArea As String
    Dim szAreaName As String
    Dim adbTemp(1 To 8) As Double
    Dim i As Integer
    Dim j As Integer
    
    Dim szStation As String
    Dim aszStation() As String
    Dim nCount As Integer
    
    Dim szInterProvince As String
    Dim szInCity As String
    Dim szExtramural As String
    
    szInterProvince = m_oParam.InterProvinceStationDetail

    szInCity = m_oParam.InCityStationDetail

    szExtramural = m_oParam.ExtramuralStationDetail
    
    If prsArea.RecordCount > 0 Then prsArea.MoveFirst
    Do While Not prsArea.EOF
        '如果车次种类已经一类结束则插入合计
        If szArea <> "" And prsArea!province_in_out <> szArea Then
            prsTemp.AddNew
            prsTemp!Kinds = szAreaName
            prsTemp!route_id = "合计"
            prsTemp!route_name = "合计"
            prsTemp!bus_count = adbTemp(1)
            prsTemp!passenger_number = adbTemp(2)
            prsTemp!fact_float = adbTemp(3)
            prsTemp!total_float = adbTemp(4)
            prsTemp!total_number = adbTemp(5)
            prsTemp!full_seat_rate = adbTemp(6) / j
            prsTemp!fact_load_rate = adbTemp(7) / j
            prsTemp!total_ticket_price = adbTemp(8)
            
            For i = 1 To 8
                If i = 6 Or i = 7 Then
                    m_aszSum(i) = m_aszSum(i) + adbTemp(i) / j
                Else
                    m_aszSum(i) = m_aszSum(i) + adbTemp(i)
                End If
                adbTemp(i) = 0
            Next i
            
            j = 0
            m_nKinds = m_nKinds + 1
        End If
        '赋值
        
        Select Case prsArea!province_in_out
        Case "省外" 'EA_nOutProvince '省外
            szStation = szInterProvince
        Case "市内" 'EA_nInCity '市内
            szStation = szInCity
        Case "市外" 'EA_nOutCity '市外
            szStation = szExtramural
        End Select
        aszStation = StringToTeam(szStation)
        nCount = ArrayLength(aszStation)
        For i = 1 To nCount
            If RTrim(aszStation(i)) = RTrim(prsArea!end_station_id) Then
                prsTemp.AddNew
                prsTemp!Kinds = Trim(prsArea!province_in_out) & "普通班车"
                prsTemp!route_id = "" 'Trim(prsArea!end_station_id) 'route_id
                prsTemp!route_name = prsArea!station_name 'route_name
                prsTemp!bus_count = prsArea!bus_count
                prsTemp!passenger_number = prsArea!passenger_number
                prsTemp!fact_float = prsArea!fact_float
                prsTemp!total_float = prsArea!total_float
                prsTemp!total_number = prsArea!total_number
                prsTemp!full_seat_rate = prsArea!full_seat_rate
                prsTemp!fact_load_rate = prsArea!fact_load_rate
                prsTemp!total_ticket_price = prsArea!total_ticket_price
            End If
        Next i
        '为计算合计用
        j = j + 1
        adbTemp(1) = adbTemp(1) + prsArea!bus_count
        adbTemp(2) = adbTemp(2) + prsArea!passenger_number
        adbTemp(3) = adbTemp(3) + prsArea!fact_float
        adbTemp(4) = adbTemp(4) + prsArea!total_float
        adbTemp(5) = adbTemp(5) + prsArea!total_number
        adbTemp(6) = adbTemp(6) + prsArea!full_seat_rate
        adbTemp(7) = adbTemp(7) + prsArea!fact_load_rate
        adbTemp(8) = adbTemp(8) + prsArea!total_ticket_price
        szArea = prsArea!province_in_out
        szAreaName = Trim(prsArea!province_in_out) & "普通班车"
        prsArea.MoveNext
    Loop
    '将最后一条记录进行赋值
    If prsArea.RecordCount > 0 Then
        '赋值
        prsTemp.AddNew
        prsTemp!Kinds = szAreaName 'Trim(prsArea!province_in_out) & "普通班车"
        prsTemp!route_id = "合计"
        prsTemp!route_name = "合计"
        prsTemp!bus_count = adbTemp(1)
        prsTemp!passenger_number = adbTemp(2)
        prsTemp!fact_float = adbTemp(3)
        prsTemp!total_float = adbTemp(4)
        prsTemp!total_number = adbTemp(5)
        prsTemp!full_seat_rate = adbTemp(6) / j
        prsTemp!fact_load_rate = adbTemp(7) / j
        prsTemp!total_ticket_price = adbTemp(8)
        
        
        m_nKinds = m_nKinds + 1
        For i = 1 To 8
                If i = 6 Or i = 7 Then
                    m_aszSum(i) = m_aszSum(i) + adbTemp(i) / j
                Else
                    m_aszSum(i) = m_aszSum(i) + adbTemp(i)
                End If
        Next i
        
        
        
        prsTemp.AddNew
        prsTemp!Kinds = "总计" 'Trim(prsArea!province_in_out) & "普通班车"
        prsTemp!route_id = ""
        prsTemp!route_name = ""
        prsTemp!bus_count = m_aszSum(1)
        prsTemp!passenger_number = m_aszSum(2)
        prsTemp!fact_float = m_aszSum(3)
        prsTemp!total_float = m_aszSum(4)
        prsTemp!total_number = m_aszSum(5)
        prsTemp!full_seat_rate = m_aszSum(6) / m_nKinds
        prsTemp!fact_load_rate = m_aszSum(7) / m_nKinds
        prsTemp!total_ticket_price = m_aszSum(8)
        

        prsTemp.Update
    End If
End Sub



Private Sub cmdok_Click()
    
    On Error GoTo Error_Handle
    '生成记录集

    Dim oRouteDim As New TicketRouteDim
    Dim rsBusType1 As Recordset
    Dim rsBusType2 As Recordset
    Dim rsBusType As Recordset
    Dim rsArea1 As Recordset
    Dim rsArea2 As Recordset
    Dim rsArea As Recordset
    Dim rsTemp As New Recordset
    Dim i As Integer
    
    
    
    
    oRouteDim.Init m_oActiveUser
    Me.MousePointer = vbHourglass
    
    Set rsBusType1 = oRouteDim.GetRouteByBusTypeSimply(dtpBeginDate.Value, dtpEndDate.Value, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
'    If rsBusType1.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
    
    '得到除正常班次与滚动车次以外的车次的线路统计，按车次类型与线路进行分类列出
'
'    Set rsBusType2 = oRouteDim.GetBusNumByRoute(dtpBeginDate.Value, dtpEndDate.Value)
'    If rsBusType2.RecordCount = 0 Then Exit Sub
'    If rsBusType1.RecordCount <> 0 And rsBusType2.RecordCount <> 0 Then
'        Set rsBusType = MergeBusType(rsBusType1, rsBusType2)
'    ElseIf rsBusType1.RecordCount <> 0 Then
        Set rsBusType = rsBusType1
'    End If
'    If rsBusType.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
    
    Set rsArea1 = oRouteDim.GetRouteByAreaBusTypeSimply(dtpBeginDate.Value, dtpEndDate.Value, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
'    '得到按地区统计的人数及金额，并列出要求的线路的人数金额。
''    If rsArea1.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
'
'    Set rsArea2 = oRouteDim.GetBusNumByArea(dtpBeginDate.Value, dtpEndDate.Value)
''    If rsArea2.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
'    If rsArea1.RecordCount <> 0 And rsArea2.RecordCount <> 0 Then
'        Set rsArea = MergeArea(rsArea1, rsArea2)
'    ElseIf rsArea1.RecordCount <> 0 Then
        Set rsArea = rsArea1
'    End If
    
'    If rsArea.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
    
    'Set rsArea = rsArea1
    'rsTemp.Open
    
    With rsTemp.Fields
        .Append "Kinds", adChar, 20
        .Append "route_id", adChar, 4
        .Append "Route_name", adChar, 16
        .Append "bus_count", adInteger
        .Append "passenger_number", adDouble
        .Append "fact_float", adDouble
        .Append "total_float", adDouble
        .Append "total_number", adDouble
        .Append "full_seat_rate", adDouble
        .Append "fact_load_rate", adDouble
        .Append "total_ticket_price", adCurrency

    End With

    '给记录集赋值
    rsTemp.Open

    For i = 1 To 8
        m_aszSum(i) = 0
    Next i

    m_nKinds = 0
    If Not (rsBusType Is Nothing) Then
        If rsBusType.RecordCount <> 0 Then EvalRecBusType rsTemp, rsBusType
    End If
    If Not (rsArea Is Nothing) Then
        If rsArea.RecordCount <> 0 Then EvalRecArea rsTemp, rsArea
    End If
    
    
    Set m_rsData = rsTemp
    ReDim m_vaCustomData(1 To 4, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY-MM-DD")
    
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY-MM-DD")
    
    Dim szSellStation As String
    ResolveDisplay cboSellStation, szSellStation
    m_vaCustomData(3, 1) = "上车站"
    m_vaCustomData(3, 2) = szSellStation
  
    m_vaCustomData(4, 1) = "制表人"
    m_vaCustomData(4, 2) = m_oActiveUser.UserID
    
    
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


'Private Sub FillSellStation()
'    Dim oSystemMan As SystemMan
'    Dim asztemp() As String
'    On Error GoTo Here
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'    asztemp = oSystemMan.GetAllSellStation(g_szUnitID)
'    If m_oActiveUser.SellStationID = "" Then
'        cboSellStation.AddItem ""
'        For i = 1 To ArrayLength(asztemp)
'            cboSellStation.AddItem MakeDisplayString(asztemp(i, 1), asztemp(i, 2))
'
'        Next i
'    '否则只填充用户所属的上车站
'    Else
'        For i = 1 To ArrayLength(asztemp)
'            If m_oActiveUser.SellStationID = asztemp(i, 1) Then
'               cboSellStation.AddItem MakeDisplayString(asztemp(i, 1), asztemp(i, 2))
'               Exit For
'            End If
'        Next i
'    End If
'Here:
'    ShowMsg err.Description
'End Sub
