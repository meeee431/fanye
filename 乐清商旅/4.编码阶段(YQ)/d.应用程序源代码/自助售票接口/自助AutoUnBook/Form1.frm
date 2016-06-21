VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "NTSVC.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "浙江方苑自助售票自动取消服务端"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   600
      Top             =   840
   End
   Begin NTService.NTService NTService 
      Left            =   1800
      Top             =   1320
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "SelfSellAutoUnBook"
      ServiceName     =   "ProvinceSockService"
      StartMode       =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private aStation() As RTConnection

Private g_aszAllStartStation() As String
Private nOverTime As Integer
Private nQueryTime As Integer
Private nQueryMain As Integer

Private Sub Form_Load()

    RegService '注册服务


    Initialize
    Timer1.Enabled = True
    
End Sub


Public Sub Initialize()
On Error GoTo here:
    Dim i As Integer
    Dim FileNo As Integer
    Dim m_nSellStationCount As Integer
    Dim szTemp As String
    Dim szQueryTime As String
    Dim szOverTime As String
    '====================================================================
    '读取自定义数据
    '====================================================================
    m_nSellStationCount = 0
    FileNo = FreeFile
    Open App.Path + "\Param.ini" For Input As #FileNo
        Input #FileNo, szQueryTime '查询时间
        Input #FileNo, szOverTime '超时时间
        nQueryTime = Trim(Right(szQueryTime, Len(szQueryTime) - InStr(1, szQueryTime, "]")))
        nQueryMain = nQueryTime
        nOverTime = Trim(Right(szOverTime, Len(szOverTime) - InStr(1, szOverTime, "]")))
        m_nSellStationCount = 1
        If m_nSellStationCount > 0 Then
            ReDim aStation(1 To m_nSellStationCount)
            ReDim g_aszAllStartStation(1 To 1, 1 To 9)
            For i = 1 To m_nSellStationCount
                Input #FileNo, g_aszAllStartStation(i, 1) '用户名
                Input #FileNo, g_aszAllStartStation(i, 2) '密码
                Input #FileNo, g_aszAllStartStation(i, 3) '数据库名
                Input #FileNo, g_aszAllStartStation(i, 4) 'IP
                g_aszAllStartStation(i, 1) = Trim(Right(g_aszAllStartStation(i, 1), Len(g_aszAllStartStation(i, 1)) - InStr(1, g_aszAllStartStation(i, 1), "]")))
                g_aszAllStartStation(i, 2) = Trim(Right(g_aszAllStartStation(i, 2), Len(g_aszAllStartStation(i, 2)) - InStr(1, g_aszAllStartStation(i, 2), "]")))
                g_aszAllStartStation(i, 3) = Trim(Right(g_aszAllStartStation(i, 3), Len(g_aszAllStartStation(i, 3)) - InStr(1, g_aszAllStartStation(i, 3), "]")))
                g_aszAllStartStation(i, 4) = Trim(Right(g_aszAllStartStation(i, 4), Len(g_aszAllStartStation(i, 4)) - InStr(1, g_aszAllStartStation(i, 4), "]")))
  
            Next i
        End If
        If err.Number <> 0 Then
            err.Raise err.Number, "配置文件Param.ini错误"

        End If
    Close #FileNo


    '====================================================================
    For i = 1 To m_nSellStationCount
        Set aStation(i) = New RTConnection
        aStation(i).ConnectionString = GetConnectionStr()
        'WriteLog aStation(i).ConnectionString
        AutoUnBook
    Next
    Exit Sub
here:
    MsgBox err.Description
    End
End Sub
Private Sub RegService()
On Error GoTo ServiceError
    If Command = "/i" Then
        NTService.Interactive = True
            If NTService.Install Then
                NTService.SaveSetting "Parameters", "TimerInterval", "500"
                
            Else
                MsgBox NTService.DisplayName & "系统服务已安装或未能正确安装"
            End If
            End
    ElseIf Command = "/u" Then
        If NTService.Uninstall Then
           
        Else
            MsgBox NTService.DisplayName & "系统服务未能卸载"
        End If
        End
    End If
    NTService.ControlsAccepted = svcCtrlPauseContinue
    NTService.StartService
ServiceError:
    Call NTService.LogEvent(svcMessageError, svcEventError, "[" & err.Number & "] " & err.Description)
End Sub

Private Sub NTService_Start(SUCCESS As Boolean)
SUCCESS = True
End Sub

Private Sub AutoUnBook()
On Error GoTo here
    Dim i, j As Integer
    Dim szSql As String
    
    Dim rs As Recordset
    Dim lError As Long
    Dim szErrorString As String

    Dim szBusID As String
    Dim szBusDate As String
    Dim szSeatNo As String
    Dim BookTime As Date
    Dim szTemp As String
    Dim aTemp() As String
    Dim nDiff As Long

        
        szSql = "select * from Env_bus_seat_lst where bus_date>='" & Date & "' and remark like 'self|%'"

        aStation(1).BeginTrans
        Set rs = aStation(1).Execute(szSql)
        If rs.RecordCount > 0 Then
            For j = 1 To rs.RecordCount
                szTemp = FormatDbValue(rs!remark)
                aTemp = Split(szTemp, "|")
                BookTime = aTemp(1)
                nDiff = DateDiff("n", BookTime, Now)
                If nDiff > nOverTime Then
                    UnBookTK aStation(1), FormatDbValue(rs!bus_id), FormatDbValue(rs!bus_date), FormatDbValue(rs!seat_no)
                    WriteLog FormatDbValue(rs!bus_id) & FormatDbValue(rs!bus_date) & FormatDbValue(rs!seat_no) & "取消成功"
                End If
                rs.MoveNext
            Next

        End If
        aStation(1).CommitTrans



    Exit Sub
here:
    WriteLog err.Description
    aStation(1).RollbackTrans

   
End Sub
Public Function UnBookTK(poDb As RTConnection, paszBusID As String, paszBusDate As Date, pszSeatNo As String) As String     '取消预留
    
    Dim szSql As String
    Dim rsTemp2 As Recordset
    Dim szSeatNoTemplog As String
    Dim j As Integer
    

        szSql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & Trim(paszBusID) & "'" _
               & " AND  bus_date='" & ToDBDate(paszBusDate) & "'" _
               & "AND seat_no ='" & Trim(pszSeatNo) & "'" _
               & " AND Status <> 1"
        Set rsTemp2 = poDb.Execute(szSql)
        
        If rsTemp2.RecordCount > 0 Then
         szSeatNoTemplog = ""
         
          For j = 0 To rsTemp2.RecordCount - 1
           szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp2!seat_no) & "]"
           rsTemp2.MoveNext
          Next
          err.Raise "1234", "", "座位:" & szSeatNoTemplog

        End If
        szSql = "UPDATE Env_bus_seat_lst SET status=0 ," _
                & " remark='' " _
                & "    WHERE bus_id='" & Trim(paszBusID) & "' AND  bus_date='" & ToDBDate(Trim(paszBusDate)) & "' " _
                & " AND seat_no  ='" & Trim(pszSeatNo) & "'" _
                & " AND  Status=1"
        poDb.Execute szSql
    

        
                '修改环境车次的可售张数
                szSql = "UPDATE Env_bus_info SET  sale_seat_quantity=sale_seat_quantity+1 , "
                '这里默认座位类型为"坐位"
                szSql = szSql & "seat_remain=seat_remain + 1 "
                szSql = szSql & " WHERE bus_id='" & Trim(paszBusID) & "' AND bus_date='" & ToDBDate(Trim(paszBusDate)) & "'"
                poDb.Execute szSql

    



End Function

'把日志写入文件
Public Sub WriteLog(ByVal pszLog As String)

    Dim szDir As String
    Dim FileNo As Integer
    FileNo = FreeFile
    szDir = App.Path & "\log\" & Format(Date, "YYYYMMDD") & "KYServerLog.ini"
    If Dir(szDir) <> "" Then
        Open szDir For Append As #FileNo
            Print #1, Format(Now, "yyyy-mm-dd HH:MM:ss") & "  "; pszLog  ' 将文本数据写入文件。
            Print #1, ' 将空白行写入文件。
        Close #FileNo
    Else
        Open szDir For Output As #FileNo
            Print #1, Format(Now, "yyyy-mm-dd HH:MM:ss") & "  "; pszLog  ' 将文本数据写入文件。
            Print #1, ' 将空白行写入文件。
        Close #FileNo
    End If '写日志
End Sub

Private Sub Timer1_Timer()
    nQueryTime = nQueryTime - 1
    
    If nQueryTime = 0 Then
        Timer1.Enabled = False
        AutoUnBook
        Timer1.Enabled = True
        nQueryTime = nQueryMain
    ElseIf nQueryTime < 0 Then
        nQueryTime = nQueryMain
        
    End If
End Sub
