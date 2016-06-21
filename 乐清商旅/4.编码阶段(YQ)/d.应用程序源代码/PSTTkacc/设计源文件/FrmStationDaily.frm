VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form FrmStationDaily 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "站务简报"
   ClientHeight    =   4020
   ClientLeft      =   5040
   ClientTop       =   4065
   ClientWidth     =   5145
   HelpContextID   =   60000250
   Icon            =   "FrmStationDaily.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   1920
      TabIndex        =   15
      Top             =   2850
      Width           =   2310
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   990
      TabIndex        =   14
      Top             =   3510
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
      MICON           =   "FrmStationDaily.frx":000C
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2430
      Width           =   2310
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   2340
      TabIndex        =   2
      Top             =   3525
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
      MICON           =   "FrmStationDaily.frx":0028
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
      Left            =   -60
      TabIndex        =   10
      Top             =   690
      Width           =   6885
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
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
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3630
      TabIndex        =   1
      Top             =   3540
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
      MICON           =   "FrmStationDaily.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtBusID 
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Text            =   "49"
      Top             =   1950
      Width           =   2310
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Top             =   990
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   1920
      TabIndex        =   4
      Top             =   1470
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   11
      Top             =   3285
      Width           =   8745
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   570
      TabIndex        =   13
      Top             =   2490
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   570
      TabIndex        =   7
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "加班车前缀(&A):"
      Height          =   180
      Left            =   570
      TabIndex        =   6
      Top             =   2010
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   570
      TabIndex        =   5
      Top             =   1530
      Width           =   1080
   End
End
Attribute VB_Name = "FrmStationDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm
Const cszFileName = "车站站务简报模板.xls"
Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    
    Dim oDss As New TicketUnitDim
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As New Recordset
    Dim rsData As New Recordset
    Dim oTicketBusDim As New TicketUnitDim

   
    Dim cuToicketPriceTotal As Currency
    Dim nData As Double
    Dim i As Integer
    
    Dim oReScheme As New REScheme
    Dim aszBusInfo() As String
    Dim nCount As Integer
    Dim nSeatData() As Double
    
    Dim oPriceMan As New TicketPriceMan
    
    Dim szbusID As String
    Dim dyDateAddOne As Date
    Dim asgTotalPrice(0 To 15) As Single
    Dim j As Integer
    
    dyDateAddOne = DateAdd("d", 1, dtpEndDate.Value)
    
    
    If dtpBeginDate.Value > m_oParam.NowDate Then
        MsgBox "指定日期有误!"
        Exit Sub
    End If
    oTicketBusDim.Init m_oActiveUser
    Me.MousePointer = vbHourglass
    
    Set m_rsData = oTicketBusDim.GetCheckGateStatDetail(dtpBeginDate.Value, dyDateAddOne, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    
    Set rsTemp = oTicketBusDim.GetCheckGateStatSum(dtpBeginDate.Value, dyDateAddOne, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    '如果汇总数字存在则显示
    If rsTemp.RecordCount > 0 Then
        ReDim m_vaCustomData(1 To 14, 1 To 2)
        
        
        m_vaCustomData(1, 1) = "实开班次总座位"
        m_vaCustomData(1, 2) = rsTemp!total_seat_number
        m_vaCustomData(2, 1) = "实开班次总公里"
        m_vaCustomData(2, 2) = rsTemp!Mileage
        m_vaCustomData(3, 1) = "始发人数"
        m_vaCustomData(3, 2) = rsTemp!passenger_number
        m_vaCustomData(4, 1) = "始发周转量"
        m_vaCustomData(4, 2) = rsTemp!fact_float_number
        
        m_vaCustomData(5, 1) = "上座率"
        If m_vaCustomData(1, 2) <> 0 Then
            nData = m_vaCustomData(3, 2) / m_vaCustomData(1, 2)
            nData = Round(nData, 4)
        Else
            nData = 0
        End If
        m_vaCustomData(5, 2) = nData * 100 & "%"
        m_vaCustomData(6, 1) = "总周转量"
        m_vaCustomData(6, 2) = rsTemp!total_float_number
        
        m_vaCustomData(7, 1) = "实载率"
        
        
        If nData <> 0 Then
            nData = m_vaCustomData(4, 2) / m_vaCustomData(6, 2)
            nData = Round(nData, 4)
        Else
            nData = 0
        End If
        m_vaCustomData(7, 2) = nData * 100 & "%"
        m_vaCustomData(8, 1) = "开始日期"
        m_vaCustomData(8, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
        
        m_vaCustomData(9, 1) = "结束日期"
        m_vaCustomData(9, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
        
        
        '只有日报才有车次明细
        If dtpBeginDate.Value = dtpEndDate.Value Then
            oReScheme.Init m_oActiveUser
            If txtBusID.Text = "" Then
                szbusID = "9999999"
            Else
                If Right(txtBusID.Text, 1) <> "%" Then txtBusID.Text = txtBusID.Text & "%"
                szbusID = txtBusID.Text
            End If
            aszBusInfo = oReScheme.GetBus(dtpBeginDate.Value, szbusID)
            nCount = ArrayLength(aszBusInfo)
            m_vaCustomData(10, 1) = "加班车次"
            
            For i = 1 To nCount
                If Not ((aszBusInfo(i, 7) = ST_BusStopped) Or (aszBusInfo(i, 7) = ST_BusSlitpStop) Or (aszBusInfo(i, 7) = ST_BusReplace) Or (aszBusInfo(i, 7) = ST_BusMergeStopped)) Then
                    m_vaCustomData(10, 2) = m_vaCustomData(10, 2) & " " & aszBusInfo(i, 1)
                End If
            Next i
            '停班车次
            Set rsData = oTicketBusDim.GetStopBus(dtpBeginDate.Value, dyDateAddOne)
            m_vaCustomData(11, 1) = "停班车次"
            
            For i = 1 To rsData.RecordCount
                m_vaCustomData(11, 2) = m_vaCustomData(11, 2) & " " & rsData!bus_id
                rsData.MoveNext
            Next i
        End If
        m_vaCustomData(12, 1) = "收费项明细"
        
        
        On Error Resume Next
        
        
        asgTotalPrice(0) = asgTotalPrice(0) + rsTemp!base_price
        For j = 1 To 15
        asgTotalPrice(j) = asgTotalPrice(j) + rsTemp.Fields("price_item_" & j)
        Next j
        
        On Error GoTo 0
        
'        Dim aszTicketPriceItem() As String
'        Dim nTicketPriceItemCount As Integer
'
'        oPriceMan.Init m_oActiveUser
'        aszTicketPriceItem = oPriceMan.GetAllTicketItemEx() '得到所有的票价项
'        nTicketPriceItemCount = ArrayLength(aszTicketPriceItem)
'

        Dim aszTicketPriceItem() As String
        Dim nTicketPriceItemCount As Integer
'        Dim oScheme As New RegularScheme
        Dim aszTemp() As String
'        oScheme.Init m_oActiveUser
'        aszTemp = oScheme.GetRunPriceTable
        oPriceMan.Init m_oActiveUser
        aszTicketPriceItem = oPriceMan.GetAllTicketItem()
        nTicketPriceItemCount = ArrayLength(aszTicketPriceItem)



        m_vaCustomData(12, 2) = RTrim(aszTicketPriceItem(1, 2)) & ":" & Format(asgTotalPrice(0), "0.0")
        
        For i = 1 To nTicketPriceItemCount - 1
            If aszTicketPriceItem(i + 1, 3) = "1" Then '如果票价项是使用的
                m_vaCustomData(12, 2) = m_vaCustomData(12, 2) & "  " & aszTicketPriceItem(i + 1, 2) & ":" & Format(asgTotalPrice(i), "0.0")
            End If
        Next i
        
    
        Dim szSellStation As String
        ResolveDisplay cboSellStation, szSellStation
        m_vaCustomData(13, 1) = "上车站"
        m_vaCustomData(13, 2) = szSellStation
    
        m_vaCustomData(14, 1) = "制表人"
        m_vaCustomData(14, 2) = m_oActiveUser.UserID
    End If
    m_bOk = True
    Me.MousePointer = vbDefault
    Unload Me
    Exit Sub
    
Error_Handle:
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    m_bOk = False
    On Error GoTo Here
    

    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    FillSellStation cboSellStation
    txtBusID.Text = Trim(m_oParam.AdditionBusPreFix)
    
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If
    
    If Right(txtBusID.Text, 1) <> "%" Then txtBusID.Text = txtBusID.Text & "%"
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
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub



