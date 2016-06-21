VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form sxmain 
   Caption         =   "合作银行汽车票代理售票系统"
   ClientHeight    =   5190
   ClientLeft      =   1065
   ClientTop       =   990
   ClientWidth     =   9330
   ClipControls    =   0   'False
   Icon            =   "sxmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   9330
   Begin VB.Timer TimReConn 
      Interval        =   1000
      Left            =   3720
      Top             =   2280
   End
   Begin VB.Timer TimeRead 
      Interval        =   60000
      Left            =   4560
      Top             =   2280
   End
   Begin VB.Timer TcpTimer2 
      Index           =   0
      Left            =   3720
      Top             =   1200
   End
   Begin VB.Timer TcpTimer 
      Enabled         =   0   'False
      Index           =   0
      Left            =   1665
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock TcpServer 
      Index           =   0
      Left            =   2040
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   252
      Left            =   6360
      TabIndex        =   11
      Top             =   3360
      Width           =   732
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清除"
      Height          =   252
      Left            =   5320
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton cmdGetSchedules 
      Caption         =   "取车次"
      Height          =   252
      Left            =   4280
      TabIndex        =   9
      Top             =   3360
      Width           =   732
   End
   Begin VB.CommandButton cmdGetStations 
      Caption         =   "取站点"
      Height          =   252
      Left            =   3240
      TabIndex        =   8
      Top             =   3360
      Width           =   732
   End
   Begin VB.CommandButton cmdAccountTab 
      Caption         =   "对帐单"
      Height          =   252
      Left            =   2200
      TabIndex        =   7
      Top             =   3360
      Width           =   732
   End
   Begin VB.CommandButton cmdSchedulesTab 
      Caption         =   "车次表"
      Height          =   252
      Left            =   1160
      TabIndex        =   6
      Top             =   3360
      Width           =   732
   End
   Begin VB.CommandButton cmdStationsTab 
      Caption         =   "站点表"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   732
   End
   Begin VB.Frame Frame1 
      Caption         =   "客运中心网络服务器连接状态"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   0
      TabIndex        =   2
      Top             =   3855
      Width           =   9375
      Begin VB.CommandButton cmdConnect 
         Caption         =   "接入"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   165
         TabIndex        =   12
         Top             =   555
         Width           =   492
      End
      Begin VB.TextBox txtStatus 
         Height          =   288
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Text            =   " "
         Top             =   915
         Width           =   852
      End
      Begin VB.Label lblStartStation 
         AutoSize        =   -1  'True
         Caption         =   "客运中心"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   165
         TabIndex        =   4
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9330
      TabIndex        =   1
      Top             =   4890
      Width           =   9330
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "sxmain.frx":08CA
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   12
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "opDate"
         Caption         =   "交易日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "MachineIP"
         Caption         =   "前置机"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "BankID"
         Caption         =   "车站"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "OperatorID"
         Caption         =   "工号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TradeID"
         Caption         =   "交易码"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "TicketID"
         Caption         =   "票号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Amount"
         Caption         =   "数量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "SumMoney"
         Caption         =   "金额(元)"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "TradeOk"
         Caption         =   "交易结果"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   -1  'True
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   629.858
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      Top             =   4335
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   979
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=MSDASQL;dsn=sx;uid=sa;pwd=;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=sx;uid=sa;pwd=;"
      OLEDBFile       =   ""
      DataSourceName  =   "sx"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select opDate,MachineIP,BankID,OperatorID,TradeID,TicketID,Amount,SumMoney,TradeOk from DailyBook Order by opDate desc"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "sxmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Const cnControlMid = 1400 '多个控件的间隔

Dim dReConnetStation As Date
Dim ConnectedNum As Integer
Dim szReceiveStr(CONNECTEDMAX) As String
Dim szSendStr(CONNECTEDMAX) As String

Private Sub cmdCn_Click(Index As Integer)
End Sub


Private Sub cmdConnect_Click(Index As Integer)
    Dim szStr As String
    On Error GoTo here
    frmprompt.contents = "正在连接" & cszUnitName & BusNetName(Index) & "网络，请稍候..."
    frmprompt.Show
    szStr = oBusStation.ConnectStation(Index, BusNetIP(Index), UserName(Index), UserPWD(Index), "", "")
    If Trim(szStr) = "462" Then
        cmdConnect_Click (Index)
    End If
    Unload frmprompt
    txtStatus(Index) = oBusStation.GetCnStatus(Index)
    StartName(Index) = oBusStation.GetStartName(Index)
    Exit Sub
here:
    If Trim(szStr) = "462" Then
        cmdConnect_Click (Index)
    Else
        ShowErrMsg
        Unload frmprompt
    End If
End Sub

Private Sub TcpServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim i As Integer
    'debug.print "ConnectionRequest" & Index & "," & requestID
    
    If Index = 0 Then
    For i = 1 To ConnectedNum
      If TcpServer(i).State <> sckConnected Then
         TcpServer(i).Close
         szReceiveStr(i) = ""
         TcpServer(i).LocalPort = 0
         TcpServer(i).Accept requestID
         Exit Sub
      End If
    Next i
    ConnectedNum = ConnectedNum + 1
    Load TcpServer(ConnectedNum)
    szReceiveStr(ConnectedNum) = ""
    TcpServer(ConnectedNum).LocalPort = 0
    TcpServer(ConnectedNum).Accept requestID
  End If
End Sub

Private Sub TcpServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    'debug.print "DataArrival" & Index & "," & bytesTotal
If bytesTotal >= 119 Then
     TcpServer(Index).GetData szReceiveStr(Index)
'     'debug.print "len=" & Str(Len(szReceiveStr(Index))) & vbCr & szReceiveStr(Index)
     RxPkgProcess (Index)
   End If
End Sub

Public Sub RxPkgProcess(Index As Integer)
Dim szTradeID As String
Dim szStr As String
Dim szStartStationID As String
Dim szLenStation As Variant
Dim szALLstationLen As String
Dim szLenBus As Variant
Dim szALLBusLen As String
Dim szRecord1 As Integer
Dim szRecord2 As Integer
Dim szDir As String
Dim FileNo As Integer
Dim szCode As String
Dim rs As New ADODb.Recordset
Dim nLen As Long
On Error GoTo here
    szTradeID = GetTradeID(szReceiveStr(Index)) ' Mid(szReceiveStr(Index), 10, 4)
    szStartStationID = Trim(MidA(szReceiveStr(Index), 35, 1)) 'GetStartStationID(szReceiveStr(Index))
'    FileNo = FreeFile
    Select Case szTradeID
    
        Case InVaildUser
            
            szStr = oBusStation.VaildateUser(szReceiveStr(Index), szSendStr(Index))
            TcpServer(Index).SendData szSendStr(Index)
        Case GETSTATIONSID
            
'            If StationsStr = "" Then
                szCode = ""
                StationsStr = oBusStation.GetStationsStr(Val(szStartStationID), szRecord1, szCode)
                szLenStation = szRecord1
'            End If
            If szLenStation = 0 Then
                szALLstationLen = "无站点" & Space(10 - Len(szLenStation))
            Else
                szALLstationLen = szLenStation & Space(10 - Len(szLenStation))
            End If
            If Trim(szCode) = "" Then
                If szLenStation = 0 Then
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & "00000195" & Mid(szReceiveStr(Index), 10, 150) & "0001" & szALLstationLen & Mid(szReceiveStr(Index), 179, 21)
                                    nLen = LenA(szSendStr(Index))
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
                Else
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & "00000195" & Mid(szReceiveStr(Index), 10, 150) & "0000" & szALLstationLen & Mid(szReceiveStr(Index), 174, 21) & "|" & StationsStr & "@"
                    
                    nLen = LenA(szSendStr(Index))
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
                    
                End If
            Else
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & "00000195" & Mid(szReceiveStr(Index), 10, 150) & szCode
                    nLen = LenA(szSendStr(Index))
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            End If
'            MDLogFile szReceiveStr(Index), szSendStr(Index), E_BUYTICKETSID
            TcpServer(Index).SendData szSendStr(Index)
'            Open App.Path + "\station.dat" For Output As #3
'                Print #3, szSendStr(Index)
'            Close #3
            
            'Debug.Print szSendStr(Index)
        Case GETSCHEDULESID
'            If SchedulesStr = "" Then
                szCode = ""
                SchedulesStr = oBusStation.GetSchedulesStr(Val(szStartStationID), PackageToDate(GetBusOffDate(szReceiveStr(Index))), GetDestStationID(szReceiveStr(Index)), szReceiveStr(Index), szRecord2, szCode)
'            End If
                szLenBus = szRecord2
'            szALLBusLen = szLenBus & Space(10 - Len(szLenBus))
            If szLenBus = 0 Then
                szALLBusLen = "无车次" & Space(10 - Len(szLenBus))
            Else
                szALLBusLen = szLenBus & Space(10 - Len(szLenBus))
            End If
            If Trim(szCode) = "" Then
                If szLenBus = 0 Then
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & "00000195" & Mid(szReceiveStr(Index), 10, 150) & "0001" & szALLBusLen & Mid(szReceiveStr(Index), 179, 21)
                                    nLen = LenA(szSendStr(Index))
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
                Else
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & "00000195" & Mid(szReceiveStr(Index), 10, 150) & "0000" & szALLBusLen & Mid(szReceiveStr(Index), 174, 21) & "|" & SchedulesStr & "@"
                    
                    nLen = LenA(szSendStr(Index))
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
                End If
            Else
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & "00000195" & Mid(szReceiveStr(Index), 10, 150) & szCode
                                        nLen = LenA(szSendStr(Index))
                    szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            End If
            TcpServer(Index).SendData szSendStr(Index)
'            Open App.Path + "\bus.dat" For Output As #4
'                Print #4, szSendStr(Index)
'            Close #4
        Case BUYTICKETSID
            '=============================================
            '此处为何要判断是否为自动获取座位,需仔细看一下
            '=============================================
            'If Mid(szReceiveStr(Index), cnPosSeatID, 2) = cszAutoSeat And Mid(szReceiveStr(Index), cnPosRetCode, cnLenRetCode) = cszCorrectRetCode Then

                szStr = oBusStation.SellTickets(szReceiveStr(Index), szSendStr(Index))
                
                If Mid(szSendStr(Index), cnPosRetCode, cnLenRetCode) = cszCorrectRetCode Then
                    '开始计时器,如果时间超过8秒售票未成功,则自动废票处理
                    tcpok(Index) = 1
                    TcpTimer(Index).Interval = cnTimeOut
                    TcpTimer(Index).Enabled = True
                End If
                TcpServer(Index).SendData szSendStr(Index)
                MDLogFile szReceiveStr(Index), szSendStr(Index), E_BUYTICKETSID
        Case CANCELTICKETSID

            szStr = oBusStation.CancelTicket(szReceiveStr(Index), szSendStr(Index))
            TcpServer(Index).SendData szSendStr(Index)
            MDLogFile szReceiveStr(Index), szSendStr(Index), E_CANCELTICKETSID
        Case GETSEATSID
            On Error GoTo here
            szStr = oBusStation.GetSeats(szReceiveStr(Index), szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & "00000195" & Right(szReceiveStr(Index), FIXPKGLEN - cnLenBegin - cnLenLen) & "|" & szSendStr(Index) & "@"
'            szStr = oBusStation.GetSeatsex(szReceiveStr(Index), szSendStr(Index))
            nLen = LenA(szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            TcpServer(Index).SendData szSendStr(Index)
        Case GETSEAT_Self '自己开发的程序取座位用
        
             szStr = oBusStation.GetSeats(szReceiveStr(Index), szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & "00000195" & Right(szReceiveStr(Index), FIXPKGLEN - cnLenBegin - cnLenLen) & "|" & szSendStr(Index) & "@"
            nLen = LenA(szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            TcpServer(Index).SendData szSendStr(Index)
        Case GETTKINFO
            szStr = oBusStation.GetTicketInfo(szReceiveStr(Index), szSendStr(Index))
            nLen = LenA(szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            'szSendStr(Index) = szStr 'szSendStr(Index)
            TcpServer(Index).SendData szSendStr(Index)
        Case GETACCOUNTLISTID
            On Error GoTo here
            szStr = oBusStation.GetAccountList(szReceiveStr(Index), szSendStr(Index))
            nLen = LenA(szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            TcpServer(Index).SendData szSendStr(Index)
        Case GETTKPRICEID
            szStr = oBusStation.GetTkPrice(szReceiveStr(Index), szSendStr(Index))
            nLen = LenA(szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            TcpServer(Index).SendData szSendStr(Index)
        Case GETCHECKGATE
            On Error GoTo here
            CheckGatesStr = oBusStation.GetCheckGateStr(Val(szStartStationID))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & "00000195" & Right(szReceiveStr(Index), FIXPKGLEN - cnLenBegin - cnLenLen) & "|" & CheckGatesStr & "@"
            nLen = LenA(szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            TcpServer(Index).SendData szSendStr(Index)
        Case INTERNETSELL
'
'                szStr = oBusStation.InterNetSellTickets(szReceiveStr(Index), szSendStr(Index))
'
'                TcpServer(Index).SendData szSendStr(Index)
'                MDLogFile szReceiveStr(Index), szSendStr(Index), E_INTERNETSELL
        Case INTERNETCANCEL
               
'            szStr = oBusStation.UnInterNetTicket(szReceiveStr(Index), szSendStr(Index))
'
'            TcpServer(Index).SendData szSendStr(Index)
'            MDLogFile szReceiveStr(Index), szSendStr(Index), E_INTERNETCANCEL
        Case GetNetTKCOUNT
               
            szStr = oBusStation.GetAccountNetTicket(szReceiveStr(Index), rs, szSendStr(Index))
            If szStr = cszCorrectRetCode Then
                MakeNetTK rs
                Set rs = Nothing
            End If
            nLen = LenA(szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            TcpServer(Index).SendData szSendStr(Index)
            MDLogFile szReceiveStr(Index), szSendStr(Index), E_GetNetTKCOUNT
        Case GetTKCOUNT

            szStr = oBusStation.GetAccountTicket(szReceiveStr(Index), rs, szSendStr(Index))
            If szStr = cszCorrectRetCode Then
                MakeSellTK rs
                Set rs = Nothing
            End If
            nLen = LenA(szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
            TcpServer(Index).SendData szSendStr(Index)
            MDLogFile szReceiveStr(Index), szSendStr(Index), E_GetTKCOUNT
        Case GetNetTK

            szStr = oBusStation.GetInterNetTicket(szReceiveStr(Index), szSendStr(Index))

            TcpServer(Index).SendData szSendStr(Index)
            MDLogFile szReceiveStr(Index), szSendStr(Index), E_GetNetTK
        Case CancelPreNetTK

            szStr = oBusStation.CancelPreTK(szReceiveStr(Index), szSendStr(Index))

            TcpServer(Index).SendData szSendStr(Index)
            MDLogFile szReceiveStr(Index), szSendStr(Index), E_CancelPreNetTK
            
        Case GETALLSTATION '取全部站点

            szStr = oBusStation.GetAllStations(0, True, szReceiveStr(Index), szSendStr(Index))
            TcpServer(Index).SendData szSendStr(Index)
        Case GETALLSCHEDULESID '取全部车次

            szStr = oBusStation.GetAllSchedules(0, , szReceiveStr(Index), szSendStr(Index))
            TcpServer(Index).SendData szSendStr(Index)
        Case UpdCardID '更新银行卡号

            szStr = oBusStation.UpdBankCard(szReceiveStr(Index), szSendStr(Index))
            TcpServer(Index).SendData szSendStr(Index)
        Case QuerySellDetail
        
            szStr = oBusStation.SellDetail(szReceiveStr(Index), szSendStr(Index))
            
            TcpServer(Index).SendData szSendStr(Index)
            
        Case QuerySellCount
            szStr = oBusStation.SellCount(szReceiveStr(Index), szSendStr(Index))
            TcpServer(Index).SendData szSendStr(Index)
    End Select
    Exit Sub
here:
    If GetTradeID(szReceiveStr(Index)) = INTERNETSELL Or GetTradeID(szReceiveStr(Index)) = INTERNETCANCEL Or GetTradeID(szReceiveStr(Index)) = GetNetTK Then
        szSendStr(Index) = Left(szReceiveStr(Index), 242) & FormatLen(Left(err.Number, 4), 4) & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    Else
        szSendStr(Index) = Left(szReceiveStr(Index), cnPosRetCode - 1) & FormatLen(Left(err.Number, 4), 4) & FormatLen(MidA(err.Description, 1, 30), 30) & "@"
    End If
            nLen = LenA(szSendStr(Index))
            szSendStr(Index) = Left(szReceiveStr(Index), 1) & Format(nLen, "00000000") & Right(szSendStr(Index), Len(szSendStr(Index)) - 9)
     TcpServer(Index).SendData (szSendStr(Index))
     MDLogFile szReceiveStr(Index), szSendStr(Index), E_Error
End Sub

Private Sub TcpServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim nTkNums As Integer '票张数
    Dim szTkNo As String * cnLenTicketID
    Dim lTkNol As Long '票号长度
    Dim szStr As String
    Dim i As Integer
    Dim FileNo As Integer
    Dim szTradeID As String
    Dim szDir As String
    'debug.print "Error(" & Index & "," & Number & "," & Description & "," & Scode & "," & Source & "," & HelpFile & "," & HelpContext & "," & CancelDisplay
    TcpServer(Index).Close
    If szSendStr(Index) = "" Or GetTradeID(szSendStr(Index)) <> BUYTICKETSID Or GetTradeID(szSendStr(Index)) <> INTERNETSELL Then
        Exit Sub
    End If
    On Error GoTo AutoCancelErrHandle
    '当购票出错时,进行票的自动废票处理
    szReceiveStr(Index) = szSendStr(Index)
    szTradeID = GetTradeID(szSendStr(Index))
    nTkNums = CInt(GetTicketNum(szSendStr(Index)))
    szTkNo = GetTicketID(szSendStr(Index)) ' Mid(szSendStr(Index), 24, 7 + 1)
    lTkNol = CLng(szTkNo)
    If szTradeID = INTERNETSELL Then '网上预留取消
        szReceiveStr(Index) = Left(szReceiveStr(Index), cnLenBegin + cnLenLen) & AUTOCANCELINTTK & MidA(szReceiveStr(Index), 14, 313)
        szStr = oBusStation.UnInterNetTicket(szReceiveStr(Index), szSendStr(Index))
'        szStr = oBusStation.DailyRec(TcpServer(Index).LocalIP, szSendStr(Index))
        szSendStr(Index) = ""
    Else
        For i = 1 To nTkNums
            '此语句有问题的,未考虑车票有前缀的情况
            szTkNo = Trim(Str(lTkNol + i - 1))
            szTkNo = String(cnLenTicketID - Len(szTkNo), "0") & szTkNo
            szReceiveStr(Index) = Left(szReceiveStr(Index), cnLenBegin + cnLenLen) & AUTOCANCELTKID & _
              GetOperatorID(szReceiveStr(Index)) & GetOperatorBankID(szReceiveStr(Index)) & szTkNo & Right(szReceiveStr(Index), FIXPKGLEN - cnPosTicketType - 1 + cnLenTicketType) ', 161)
            szStr = oBusStation.CancelTicket(szReceiveStr(Index), szSendStr(Index))
            szSendStr(Index) = Left(szSendStr(Index), 9) & AUTOCANCELTKID & Right(szSendStr(Index), 177 + 1)
'            szStr = oBusStation.DailyRec(TcpServer(Index).LocalIP, szSendStr(Index))
        Next i
        szSendStr(Index) = ""
    End If
    Exit Sub
AutoCancelErrHandle:
    FileNo = FreeFile
               
            szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "Error.txt"
    If Dir(szDir) <> "" Then
            Open szDir For Append As #FileNo
                Print #FileNo, "发送" & Now & "自动取消无法完成" & szSendStr(Index)
            Close #FileNo
    Else
            Open szDir For Output As #FileNo
                Print #FileNo, "发送" & Now & "自动取消无法完成" & szSendStr(Index)
            Close #FileNo
    End If '写日志
End Sub

Private Sub TcpServer_SendComplete(Index As Integer)
'    Dim szDir As String
'    Dim FileNo As Integer
'    'debug.print "SendComplete(" & Index
'    Dim szTradeID As String
'    On Error GoTo here
'    szTradeID = GetTradeID(szSendStr(Index))
'    If szTradeID <> GETSTATIONSID And szTradeID <> GETSCHEDULESID And szTradeID <> GETCHECKGATE Then
'        '如果不为取站点，取车次才记录，否则记录太多了
'        If oBusStation.DailyRec(TcpServer(Index).RemoteHostIP, szSendStr(Index)) = cszCorrectRetCode Then
'            '   MsgBox "ok"
'        End If
'    End If
    ' szSendStr(index) = ""
    tcpok(Index) = 0
    tcpok2(Index) = 0
'    datPrimaryRS.Refresh
'    Exit Sub
'here:
'    FileNo = FreeFile
'
'            szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "Error.txt"
'    If Dir(szDir) <> "" Then
'            Open szDir For Append As #FileNo
'                Print #FileNo, "发送" & Now & "写本地数据库无法完成" & szSendStr(Index)
'            Close #FileNo
'    Else
'            Open szDir For Output As #FileNo
'                Print #FileNo, "发送" & Now & "写本地数据库无法完成" & szSendStr(Index)
'            Close #FileNo
'    End If '写日志
End Sub
Private Sub cmdAccountTab_Click()
    frmAcc.Show
End Sub

Private Sub cmdClear_Click()
    Dim szStr As String
    On Error GoTo here
    szStr = MsgBox("清除前务必生成流水文件、打印对帐单！确实要清除？", vbYesNo)
    If szStr = vbYes Then
        datPrimaryRS.Recordset.Close
        oBusStation.DailyClear
        datPrimaryRS.Refresh
    End If
    Exit Sub
here:
    ShowErrMsg
End Sub
Private Sub cmdExit_Click()
    Dim i As Integer
    If oBusStation Is Nothing Then
        Set oBusStation = Nothing
    End If
    Unload Me
    End
End Sub
Private Sub cmdGetSchedules_Click()
'    Dim i As Integer
'    Dim szStr As String
'    Dim IsToday As Boolean
'    On Error GoTo here
'    For i = 0 To ConnectedNum
'        If TcpServer(i).State <> sckClosed Then
'            TcpServer(i).Close
'        End If
'    Next i
'    szStr = MsgBox("获取当日可售车次？", vbYesNo)
'    If szStr = vbYes Then
'        IsToday = True
'    Else
'        IsToday = False
'    End If
'    For i = 0 To ConnectedNum
'        If txtStatus(i) = "接入" Then
'            szStr = MsgBox("获取" & StartName(i) & "车次信息?", vbYesNo)
'            'szStr = vbYes
'            If szStr = vbYes Then
'                frmprompt.contents = "正在获取" & cszUnitName & StartName(i) & "车次信息..."
'                frmprompt.Show
'                If oBusStation.GetAllSchedules(i, IsToday) <> cszCorrectRetCode Then
'                    MsgBox "获取车次错误！"
'                End If
'                Unload frmprompt
'            End If
'        End If
'    Next i
'    'SchedulesStr = oBusStation.GetSchedulesStr()
'    TcpServer(0).LocalPort = 8000
'    TcpServer(0).Listen
'    frmSch.Show
'    Exit Sub
'here:
'    ShowErrMsg
End Sub
Private Sub cmdGetStations_Click()
'    Dim i As Integer
'    Dim szStr As String
'    Dim IsToday As Boolean
'    On Error GoTo here
'    For i = 0 To ConnectedNum
'        If TcpServer(i).State <> sckClosed Then
'            TcpServer(i).Close
'        End If
'    Next i
'    szStr = MsgBox("获取当日可售站点？", vbYesNo)
'    If szStr = vbYes Then
'        IsToday = True
'    Else
'        IsToday = False
'    End If
'    For i = 0 To 0 'jyc 04.01.29
'        If txtStatus(i) = "接入" Then
'            szStr = MsgBox("获取" & StartName(i) & "站点信息?", vbYesNo)
'            If szStr = vbYes Then
'                frmprompt.contents = "正在获取" & cszUnitName & StartName(i) & "站点信息..."
'                frmprompt.Show
'                szStr = oBusStation.GetAllStations(i, IsToday)
'                If szStr <> cszCorrectRetCode Then
'                    MsgBox "获取" & cszUnitName & StartName(i) & "站点信息错误(" & _
'                    szStr & ")"
'                    If szStr = "462" Then
'                        szStr = oBusStation.ConnectStation(i, BusNetIP(i), UserName(i), UserPWD(i), "", "")
'                        txtStatus(i) = oBusStation.GetCnStatus(i)
'                        StartName(i) = oBusStation.GetStartName(i)
'                    End If
'                End If
'                Unload frmprompt
'            End If
'        Else
'            szStr = oBusStation.ConnectStation(i, BusNetIP(i), UserName(i), UserPWD(i), "", "")
'            txtStatus(i) = oBusStation.GetCnStatus(i)
'        End If
'    Next i
'    StationsStr = oBusStation.GetStationsStr(1)
'    TcpServer(0).LocalPort = 8000
'    TcpServer(0).Listen
'    frmSt.Show
'    Exit Sub
'here:
'    ShowErrMsg
End Sub
Private Sub cmdSchedulesTab_Click()
  frmSch.Show
End Sub
Private Sub cmdStationsTab_Click()
  frmSt.Show
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Dim szStr As String
    Dim FileNo As Integer
    Dim fso As New FileSystemObject   '文件对象
    On Error GoTo here
    datPrimaryRS.ConnectionString = GetAdodcConnectionStr
    
    Set oBusStation = CreateObject("SxBankBus.BusStation")
    StationsStr = ""
    SchedulesStr = ""
    CheckGatesStr = ""
    FileNo = FreeFile
    Open App.Path + "\sxicbcbus.ini" For Input As #FileNo
    On Error Resume Next
    '获取售票站的个数
    Input #FileNo, m_nSellStationCount
    '再获取其他信息
    For i = 0 To m_nSellStationCount - 1
        Input #FileNo, UserName(i) '用户
        Input #FileNo, UserPWD(i) '口令
        Input #FileNo, BusNetIP(i) '服务器名
        Input #FileNo, BusNetName(i) '服务器中文名
    Next i
    If err.Number <> 0 Then
        MsgBox "配置文件sxicbcbus.ini错误"
        End
    End If
    
'    For i = 0 To m_nSellStationCount - 1
'        frmprompt.Contents = "正在连接" & cszUnitName & BusNetName(i) & "网络，请稍候..."
'        frmprompt.Show
'        'szStr = oBusStation.ConnectStation(i, BusNetIP(i), UserName(i), UserPWD(i), "", "")
'        szStr = oBusStation.ConnectStation(i, BusNetIP(i), UserName(i), UserPWD(i), "", "")
'        Unload frmprompt
'    Next i
    For i = 0 To m_nSellStationCount - 1
        txtStatus(i) = oBusStation.GetCnStatus(i)
        StartName(i) = oBusStation.GetStartName(i)
    Next i
    lblStartStation(0).Caption = BusNetName(0)
    For i = 1 To m_nSellStationCount - 1
        Load lblStartStation(i)
        lblStartStation(i).Caption = BusNetName(i)
        lblStartStation(i).Left = lblStartStation(0).Left + i * cnControlMid
        lblStartStation(i).Top = lblStartStation(0).Top
        lblStartStation(i).Visible = True
        Load cmdConnect(i)
        cmdConnect(i).Left = cmdConnect(0).Left + i * cnControlMid
        cmdConnect(i).Top = cmdConnect(0).Top
        cmdConnect(i).Visible = True
        Load txtStatus(i)
        txtStatus(i).Left = txtStatus(0).Left + i * cnControlMid
        txtStatus(i).Top = txtStatus(0).Top
        txtStatus(i).Visible = True
        txtStatus(i).Text = ""
    Next i
    
    For i = 1 To CONNECTEDMAX - 1
        Load TcpTimer(i)
        Load TcpTimer2(i)
    Next i
    ConnectedNum = 0
    TcpServer(0).LocalPort = 8000 '8000'8001为宁波卖嘉兴票的端口
    TcpServer(0).Listen
    
    If Not fso.FolderExists(App.Path & "\Account") Then
        fso.CreateFolder App.Path & "\Account"
    End If
    
    If Not fso.FolderExists(App.Path & "\log") Then
        fso.CreateFolder App.Path & "\log"
    End If
    Exit Sub
here:
    ShowErrMsg
End Sub
Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  grdDataGrid.Height = grdDataGrid.RowHeight * 18
  'Me.ScaleHeight - datPrimaryRS.Height - 30 - picButtons.Height - Frame1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

'Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  'This will display the current record position for this recordset
'  datPrimaryRS.Caption = "日志记录: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
'End Sub

'Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  'This is where you put validation code
'  'This event gets called when the following actions occur
'  Dim bCancel As Boolean
'
'  Select Case adReason
'  Case adRsnAddNew
'  Case adRsnClose
'  Case adRsnDelete
'  Case adRsnFirstChange
'  Case adRsnMove
'  Case adRsnRequery
'  Case adRsnResynch
'  Case adRsnUndoAddNew
'  Case adRsnUndoDelete
'  Case adRsnUndoUpdate
'  Case adRsnUpdate
'  End Select
'
'  If bCancel Then adStatus = adStatusCancel
'End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub AutoCancelTickets(Index As Integer)
    Dim nTkNums As Integer
    Dim szTkNo As String * cnLenTicketID
    Dim lTkNol As Long
    Dim i As Integer
    Dim szStr As String
    Dim szDir As String
    Dim FileNo As Integer
    Dim a As String
    If szSendStr(Index) = "" Or GetTradeID(szSendStr(Index)) <> BUYTICKETSID Then
        Exit Sub
    End If
    On Error GoTo AutoCancelErrHandle

    FileNo = FreeFile
               
    szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "AUTOCancel.txt"
    If Dir(szDir) <> "" Then
        Open szDir For Append As #FileNo
            Print #FileNo, "接收" & Now & szReceiveStr(Index)
        Close #FileNo
    Else
        Open szDir For Output As #FileNo
            Print #FileNo, "接收" & Now & szReceiveStr(Index)
        Close #FileNo
    End If '写日志
'    szReceiveStr(Index) = Left(szSendStr(Index), cnPosTicketNum - 1) & "01" & Right(szSendStr(Index), FIXPKGLEN - cnPosTicketNum - cnLenTicketNum + 1) '01代表一张票
    nTkNums = CInt(GetTicketNum(szSendStr(Index)))
    szTkNo = GetTicketID(szSendStr(Index))
    lTkNol = CLng(szTkNo)
    For i = 1 To nTkNums
        szStr = Trim(Str(lTkNol + i - 1))
        '    szTkNo = String(7 + 1 - Len(szStr), "0") & szStr
'        a = szReceiveStr(Index)
        szTkNo = String(cnLenTicketID - Len(szTkNo), "0") & szStr
        szReceiveStr(Index) = Left(szReceiveStr(Index), cnLenBegin + cnLenLen) & AUTOCANCELTKID & FormatLen(GetOperatorID(szReceiveStr(Index)), 5) & FormatLen(GetOperatorBankID(szReceiveStr(Index)), 5) & FormatLen(szTkNo, 10) & Right(szReceiveStr(Index), FIXPKGLEN - cnPosTicketType + 1) ', 161)
        
        szStr = oBusStation.CancelTicket(szReceiveStr(Index), szSendStr(Index))
'        szSendStr(Index) = Left(szReceiveStr(Index), cnLenBegin + cnLenLen) & AUTOCANCELTKID & Right(szReceiveStr(Index), FIXPKGLEN - cnPosTradeID - cnLenTradeID + 1)
'        szStr = oBusStation.DailyRec(TcpServer(Index).LocalIP, szSendStr(Index))
        If Dir(szDir) <> "" Then
            Open szDir For Append As #FileNo
                Print #FileNo, "发送" & Now & szSendStr(Index)
            Close #FileNo
        Else
            Open szDir For Output As #FileNo
                Print #FileNo, "发送" & Now & szSendStr(Index)
            Close #FileNo
        End If
    Next i
    '  szSendStr(index) = ""

    Exit Sub
AutoCancelErrHandle:
    FileNo = FreeFile
               
            szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "AUTOCancel.txt"
    If Dir(szDir) <> "" Then
            Open szDir For Append As #FileNo
                Print #FileNo, "发送" & Now & "自动废票无法完成" & Str(err.Number) & szSendStr(Index)
            Close #FileNo
    Else
            Open szDir For Output As #FileNo
                Print #FileNo, "发送" & Now & "自动废票无法完成" & Str(err.Number) & szSendStr(Index)
            Close #FileNo
    End If '写日志
End Sub

 
Private Sub TcpTimer_Timer(Index As Integer)
    Dim szStr As String
    TcpTimer(Index).Enabled = False
    If tcpok(Index) = 1 Then
        AutoCancelTickets (Index)
        '   szStr = oBusStation.DailyRec(TcpServer(index).RemoteHostIP, szSendStr(index)) = cszCorrectRetCode
'        datPrimaryRS.Refresh
        TcpServer(Index).Close
        tcpok(Index) = 0
    End If
End Sub
 
Private Sub TcpTimer2_Timer(Index As Integer)
    Dim szStr As String
    TcpTimer2(Index).Enabled = False
    If tcpok2(Index) = 1 Then
        AutoCancelInterNetTickets (Index)
        TcpServer(Index).Close
        tcpok2(Index) = 0
    End If
End Sub
Private Sub AutoCancelInterNetTickets(Index As Integer)
    Dim nTkNums As Integer
    Dim szTkNo As String * cnLenTicketID
    Dim lTkNol As Long
    Dim i As Integer
    Dim szStr As String
    Dim szDir As String
    Dim FileNo As Integer
    If szSendStr(Index) = "" Or GetTradeID(szSendStr(Index)) <> INTERNETSELL Then
        Exit Sub
    End If
    On Error GoTo AutoCancelErrHandle
    FileNo = FreeFile
        szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "AUTOINTERCancel.txt"
    If Dir(szDir) <> "" Then
        Open szDir For Append As #FileNo
            Print #FileNo, "接收" & Now & szReceiveStr(Index)
        Close #FileNo
    Else
        Open szDir For Output As #FileNo
            Print #FileNo, "接收" & Now & szReceiveStr(Index)
        Close #FileNo
    End If '写日志
    
        szReceiveStr(Index) = Left(szReceiveStr(Index), cnLenBegin + cnLenLen) & AUTOCANCELINTTK & MidA(szReceiveStr(Index), 14, 313)
        szStr = oBusStation.UnInterNetTicket(szReceiveStr(Index), szSendStr(Index))
'        szStr = oBusStation.DailyRec(TcpServer(Index).LocalIP, szSendStr(Index))
        

    If Dir(szDir) <> "" Then
        Open szDir For Append As #FileNo
            Print #FileNo, "发送" & Now & szSendStr(Index)
        Close #FileNo
    Else
        Open szDir For Output As #FileNo
            Print #FileNo, "发送" & Now & szSendStr(Index)
        Close #FileNo
    End If
    Exit Sub
AutoCancelErrHandle:
    FileNo = FreeFile
               
            szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "AUTOINTERCancel.txt"
    If Dir(szDir) <> "" Then
            Open szDir For Append As #FileNo
                Print #FileNo, "发送" & Now & "网上票自动取消无法完成" & Str(err.Number) & szSendStr(Index)
            Close #FileNo
    Else
            Open szDir For Output As #FileNo
                Print #FileNo, "发送" & Now & "网上票自动取消无法完成" & Str(err.Number) & szSendStr(Index)
            Close #FileNo
    End If '写日志
End Sub
Private Sub MakeNetTK(rs As ADODb.Recordset) '生成网上售票报表
  Dim FileNo As Integer
  Dim szDir As String
  FileNo = FreeFile
  szDir = App.Path + "\Account\" & Format(Date, "YYYYMMDD") & "NetAccount.txt"
  If Dir(szDir) = "" Then
    Open szDir For Output As #FileNo

   Else
    Open szDir For Output As #FileNo

   End If
      Print #FileNo, "发车日期"; vbTab; "车次"; vbTab; "起点站"; vbTab; "终点站"; vbTab; "用户代码"; vbTab; _
          "座位号"; vbTab; "票价"; vbTab; "票状态"; "     "; "票号"; vbTab; "     "; "操作时间"; vbTab; "    "; "票号"; vbTab; "    "; "取票号"; vbTab; "             "; "验证码"
      Do While Not rs.EOF
          Print #FileNo, rs!StartTime; vbTab; Trim(rs!bus_id); vbTab; Trim(rs!StartID); vbTab; Trim(rs!DestID); vbTab; Trim(rs!operatorid); vbTab; vbTab; _
          Trim(rs!SeatNo); vbTab; Trim(rs!price); vbTab; Trim(rs!Status); vbTab; rs!ticketid; vbTab; rs!selldate; vbTab; rs!ticketid; vbTab; rs!GetTicketID; vbTab; rs!ValiDateID
          
        rs.MoveNext
      Loop
  Close #FileNo

End Sub
Private Sub MakeSellTK(rs As ADODb.Recordset) '生成售票报表
  Dim FileNo As Integer
  Dim szDir As String
  
  FileNo = FreeFile
  szDir = App.Path + "\Account\" & Format(Date, "YYYYMMDD") & "SellTKAccount.txt"
  If Dir(szDir) = "" Then
      Open szDir For Output As #FileNo

   Else
    
      Open szDir For Output As #FileNo

   End If
        Print #FileNo, "票号"; vbTab; "发车日期"; vbTab; "车次"; vbTab; "起点站"; vbTab; "终点站"; vbTab; "用户代码"; vbTab; _
            "座位号"; vbTab; "票价"; vbTab; "票状态"; vbTab; "操作时间"
        Do While Not rs.EOF
            Print #FileNo, rs!ID; vbTab; Trim(rs!StartTime); vbTab; Trim(rs!scheduleid); vbTab; Trim(rs!StartName); vbTab; Trim(rs!destname); vbTab; _
            Trim(rs!operatorid); vbTab; Trim(rs!seatid); vbTab; Trim(rs!price); vbTab; rs!Status; vbTab; rs!selldate
            
          rs.MoveNext
        Loop
    
      Close #FileNo
End Sub
Private Sub MDLogFile(szReceive As String, szSend As String, eType As E_FileType)
    Dim szLog As String
    Dim FileNo As Integer
    Dim szDir As String
    FileNo = FreeFile
    Select Case eType
        Case E_BUYTICKETSID
            szLog = ".txt"
        Case E_CANCELTICKETSID
            szLog = "Cancel.txt"
        Case E_INTERNETSELL
            szLog = "InternetSell.txt"
        Case E_INTERNETCANCEL
            szLog = "InterCancel.txt"
        Case E_GetNetTKCOUNT
            szLog = "GetNetTKCOUNT.txt"
        Case E_GetTKCOUNT
            szLog = "GetTKCOUNT.txt"
        Case E_GetNetTK
            szLog = "GetNetTK.txt"
        Case E_CancelPreNetTK
            szLog = "CancelPreNetTK.txt"
        Case E_Error
            szLog = "Error.txt"
    End Select
              
               
    szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & szLog
    If Dir(szDir) <> "" Then
        Open szDir For Append As #FileNo
            Print #FileNo, "接收" & Now & szReceive
            Print #FileNo, "发送" & Now & szSend
        Close #FileNo
    Else
        Open szDir For Output As #FileNo
            Print #FileNo, "接收" & Now & szReceive
            Print #FileNo, "发送" & Now & szSend
        Close #FileNo
    End If '写日志
End Sub

Private Sub TimeRead_Timer()
On Error GoTo here
    Dim FileNo As Integer
    Dim szReConnetStationHH As String
    Dim szReConnetStationMM As String
    Dim szReConnNum As String
    Dim i As Integer
    Dim szDir As String
    
    FileNo = FreeFile
    Open App.Path + "\TransToStation.ini" For Input As #FileNo
        Input #FileNo, szReConnetStationHH
        Input #FileNo, szReConnetStationMM
        Input #FileNo, szReConnNum
    
    szReConnetStationHH = Trim(Right(szReConnetStationHH, Len(szReConnetStationHH) - InStr(1, szReConnetStationHH, "]")))
    szReConnetStationMM = Trim(Right(szReConnetStationMM, Len(szReConnetStationMM) - InStr(1, szReConnetStationMM, "]")))
    szReConnNum = Trim(Right(szReConnNum, Len(szReConnNum) - InStr(1, szReConnNum, "]")))
    If CInt(szReConnNum) = 0 Then
        Close #FileNo
        Exit Sub
    End If
    ReDim ReConnStation(0 To CInt(szReConnNum - 1))
    For i = 0 To szReConnNum - 1
        Input #FileNo, ReConnStation(i) '重连接站的序号
    Next i
    Close #FileNo
    dReConnetStation = szReConnetStationHH & ":" & szReConnetStationMM
    Exit Sub
here:
    FileNo = FreeFile
               
            szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "ReConn.txt"
    If Dir(szDir) <> "" Then
            Open szDir For Append As #FileNo
                Print #FileNo, Now & "无法读取TransToStation.ini配置文件"
            Close #FileNo
    Else
            Open szDir For Output As #FileNo
                Print #FileNo, Now & "无法读取TransToStation.ini配置文件"
            Close #FileNo
    End If '写日志
End Sub

Private Sub TimReConn_Timer()
    If Format(Now, "HH:MM:SS") = Format(dReConnetStation, "HH:MM:SS") Then
        Dim i As Integer
        If ArrayLength(ReConnStation) = 0 Then Exit Sub
        For i = 0 To ArrayLength(ReConnStation) - 1
            ReConn (ReConnStation(i))
        Next i
    End If
End Sub
Private Sub ReConn(Index As Integer)

    Dim szStr As String
    Dim FileNo As Integer
    Dim szDir As String
    
    On Error GoTo here
    
    szStr = oBusStation.ConnectStation(Index, BusNetIP(Index), UserName(Index), UserPWD(Index), "", "")

    txtStatus(Index) = oBusStation.GetCnStatus(Index)
    StartName(Index) = oBusStation.GetStartName(Index)
    FileNo = FreeFile
               
            szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "ReConn.txt"
    If Dir(szDir) <> "" Then
            Open szDir For Append As #FileNo
                Print #FileNo, Now & "连接站点" & Index & "成功"
            Close #FileNo
    Else
            Open szDir For Output As #FileNo
                Print #FileNo, Now & "连接站点" & Index & "成功"
            Close #FileNo
    End If '写日志
    Exit Sub
here:
    FileNo = FreeFile
               
            szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "ReConn.txt"
    If Dir(szDir) <> "" Then
            Open szDir For Append As #FileNo
                Print #FileNo, Now & "无法连接站点" & Index
            Close #FileNo
    Else
            Open szDir For Output As #FileNo
                Print #FileNo, Now & "无法连接站点" & Index
            Close #FileNo
    End If '写日志
End Sub

