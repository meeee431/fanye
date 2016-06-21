Attribute VB_Name = "mdlMsgEvent"
'''事件处理模块
''Option Explicit
''
''Public Enum eEventId
''    AddBus = 1              '添加车次
''    AdjustTime = 2          '调整时间
''    ChangeBusCheckGate = 3  '更改检票口
''    ChangeBusSeat = 4       '更改座位
''    ChangeBusStandCount = 5 '某车次的站票数改变
''    ChangeBusTime = 6       '更改车次发车时间
''    ChangeParam = 7         '更改参数
''    ExStartCheckBus = 8     '补检车次
''    MergeBus = 9            '车次并班
''    RemoveBus = 10          '删除车次
''    ResumeBus = 11          '车次复班
''    StartCheckBus = 12      '开检车次
''    StopBus = 13            '车次停班
''    StopCheckBus = 14       '停检车次
''End Enum
''Public Sub RunMsgEvent(EventMode As eEventId, EventParam() As String)
'''    ' 参数注释
'''    ' *************************************
'''    ' EventMode:消息类型（消息Id）
'''    ' EventParam:参数数组
'''    ' *************************************
''
''    Dim nTmp As Integer
''    Dim tTmpBusInfo As tCheckBusLstInfo
''    Select Case EventMode
''        Case eEventId.AddBus    '新增车次
''            Event_AddBus EventParam(1)
''
''        Case eEventId.ChangeBusCheckGate    '更改车次检票口
''            Event_ChangeBusCheckGate EventParam(1), EventParam(3)
''
''        Case eEventId.ChangeBusTime     '更改车次时间
''            Event_ChangeBusTime EventParam(1), Format(EventParam(3), cszDateTimeStr)
''
''        Case eEventId.MergeBus          '并班车次
''            Event_MergeBus EventParam(1)
''
''        Case eEventId.RemoveBus         '删除车次
''            Event_RemoveBus EventParam(1)
''
''        Case eEventId.ResumeBus         '车次复班
''            Event_ResumeBus EventParam(1)
''
''        Case eEventId.StopBus           '车次停班
''            Event_StopBus EventParam(1)
''
''        Case eEventId.StartCheckBus     '车次开检
''            Event_StartCheckBus EventParam(1), Val(EventParam(3))
''
''        Case eEventId.ExStartCheckBus       '车次补检
''            Event_ExCheckBus EventParam(1)
''
''        Case eEventId.StopCheckBus      '车次停检
''            Event_StopCheckBus EventParam(1)
''
''    End Select
''End Sub
''
''Private Sub Event_AddBus(szBusId As String)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim oReBus As REBus
''        Set oReBus = New REBus
''        oReBus.Init g_oActiveUser
''        oReBus.Identify szBusId, Date
''
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        tTmpBusInfo.BusID = Trim(oReBus.BusID)
''        tTmpBusInfo.BusMode = oReBus.BusType
''        tTmpBusInfo.BusSerial = 0
''        tTmpBusInfo.CheckGate = Trim(oReBus.CheckGate)
''        tTmpBusInfo.Company = oReBus.CompanyName
''        tTmpBusInfo.EndStationName = oReBus.EndStationName
''        tTmpBusInfo.Owner = oReBus.OwnerName
''        tTmpBusInfo.Vehicle = oReBus.VehicleTag
''        tTmpBusInfo.StartupTime = oReBus.StartupTime
''        tTmpBusInfo.Status = EREBusStatus.ST_BusNormal
''        frmGateMoniter.m_cBusInfo.Addone tTmpBusInfo
''        '如果此信息正在显示，刷新之
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''        End If
''    End If
''End Sub
''Private Sub Event_RemoveBus(szBusId As String)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim szCheckGateID As String
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp > 0 Then
''            szCheckGateID = frmGateMoniter.m_cBusInfo.Item(nTmp).CheckGate
''            frmGateMoniter.m_cBusInfo.RemoveOne nTmp
''        Else
''            Exit Sub
''        End If
''        '如果此信息正在显示，刷新之
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = szCheckGateID Then
''            frmGateMoniter.RefreshlvCheckBus szCheckGateID
''        End If
''    End If
''End Sub
''Private Sub Event_StopBus(szBusId As String)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then Exit Sub
''
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp) '更新状态
''        tTmpBusInfo.Status = EREBusStatus.ST_BusStopped
''        frmGateMoniter.m_cBusInfo.UpdateOne tTmpBusInfo
''
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''        End If
''    End If
''End Sub
''
''Private Sub Event_ResumeBus(szBusId As String)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then Exit Sub
''
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp) '更新状态
''        tTmpBusInfo.Status = EREBusStatus.ST_BusNormal
''        frmGateMoniter.m_cBusInfo.UpdateOne tTmpBusInfo
''
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''        End If
''    End If
''End Sub
''
''Private Sub Event_MergeBus(szBusId As String)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then Exit Sub
''
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp) '更新状态
''        tTmpBusInfo.Status = EREBusStatus.ST_BusMergeStopped
''        frmGateMoniter.m_cBusInfo.UpdateOne tTmpBusInfo
''
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''        End If
''    End If
''
''End Sub
''Private Sub Event_ChangeBusTime(szBusId As String, dtNewTime As Date)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        Dim szCheckGateID As String
''
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then
''            Exit Sub
''        End If
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp)
''
''        tTmpBusInfo.StartupTime = dtNewTime
''        frmGateMoniter.m_cBusInfo.UpdateOne tTmpBusInfo
''        szCheckGateID = tTmpBusInfo.CheckGate
''
''        nTmp = m_cChkingBusInfo.FindItem(szBusId)
''        If nTmp > 0 Then    '更改下一检票车次发车时间
''            If m_cChkingBusInfo.Item(nTmp).BusID = szBusId Then
''                tTmpBusInfo = m_cChkingBusInfo
''                tTmpBusInfo.StartupTime = dtNewTime
''                m_cChkingBusInfo.UpdateOne tTmpBusInfo
''            End If
''        Else    '得到下一检票车次
''            If dtNewTime > Time Then
''                frmGateMoniter.GetChkingBusInfo szCheckGateID
''            End If
''        End If
''
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = szCheckGateID Then
''            frmGateMoniter.RefreshlvCheckBus szCheckGateID
''            frmGateMoniter.RefreshChkingBus szCheckGateID
''        End If
''    End If
''End Sub
''
''Private Sub Event_ChangeBusCheckGate(szBusId As String, szCheckGate As String)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        Dim szOldGate As String
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then
''            Exit Sub
''        End If
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp)
''        szOldGate = tTmpBusInfo.CheckGate
''        tTmpBusInfo.CheckGate = szCheckGate
''        frmGateMoniter.m_cBusInfo.UpdateOne tTmpBusInfo
''
''        nTmp = m_cChkingBusInfo.FindItemByGate(szOldGate)    '刷新该检票口的正检/下一检票车次信息
''        If nTmp > 0 Then
''            tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp)
''            If tTmpBusInfo.BusID = szBusId Then
''                frmGateMoniter.GetChkingBusInfo szOldGate    '得到最新的在检车次
''            End If
''        End If
''
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = szOldGate Then
''            frmGateMoniter.RefreshlvCheckBus szOldGate
''            frmGateMoniter.RefreshChkingBus szOldGate
''        End If
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = szCheckGate Then
''            frmGateMoniter.RefreshlvCheckBus szCheckGate
''        End If
''    End If
''End Sub
''
''
''Private Sub Event_StartCheckBus(szBusId As String, nSerialNo As Integer)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then
''            Exit Sub
''        End If
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp)
''        tTmpBusInfo.Status = EREBusStatus.ST_BusChecking
''        tTmpBusInfo.BusSerial = nSerialNo
''        frmGateMoniter.m_cBusInfo.UpdateOne tTmpBusInfo
''
''        frmGateMoniter.GetChkingBusInfo tTmpBusInfo.CheckGate
''            '刷新正检/下一检票车次信息
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''            frmGateMoniter.RefreshChkingBus tTmpBusInfo.CheckGate
''        End If
''        For nTmp = 1 To frmGateMoniter.lvCheckGate.ListItems.Count   '刷新检票口状态
''            If frmGateMoniter.lvCheckGate.ListItems(nTmp).Text = tTmpBusInfo.CheckGate Then
''                frmGateMoniter.lvCheckGate.ListItems(nTmp).SmallIcon = 1
''            End If
''        Next nTmp
''    End If
''End Sub
''Private Sub Event_ExCheckBus(szBusId As String)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''
''        Dim nOldSerialNo As Integer
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then
''            Exit Sub
''        End If
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp)
''        tTmpBusInfo.Status = EREBusStatus.ST_BusExtraChecking
''        frmGateMoniter.m_cBusInfo.UpdateOne tTmpBusInfo
''
''        frmGateMoniter.GetChkingBusInfo tTmpBusInfo.CheckGate
''        '刷新正检/下一检票车次信息
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''            frmGateMoniter.RefreshChkingBus tTmpBusInfo.CheckGate
''        End If
''        For nTmp = 1 To frmGateMoniter.lvCheckGate.ListItems.Count   '刷新检票口状态
''            If frmGateMoniter.lvCheckGate.ListItems(nTmp).Text = tTmpBusInfo.CheckGate Then
''                frmGateMoniter.lvCheckGate.ListItems(nTmp).SmallIcon = 1
''            End If
''        Next nTmp
''    End If
''End Sub
''
''Private Sub Event_StopCheckBus(szBusId As String)
''    '检票监控时处理
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        Dim nOldSerialNo As Integer
''
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then
''            Exit Sub
''        End If
''
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp)
''        If tTmpBusInfo.BusMode = EBusType.TP_RegularBus Then
''            tTmpBusInfo.Status = EREBusStatus.ST_BusStopCheck
''        Else
''            tTmpBusInfo.Status = EREBusStatus.ST_BusNormal
''        End If
''        frmGateMoniter.m_cBusInfo.UpdateOne tTmpBusInfo
''
''        nTmp = m_cChkingBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then    '如果是该车次不是(正检/下一检票车次)
''            Exit Sub
''        End If
''        frmGateMoniter.GetChkingBusInfo tTmpBusInfo.CheckGate    '得到最新（正检/下一检票车次）
''
''        '刷新正检/下一检票车次信息
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''            frmGateMoniter.RefreshChkingBus tTmpBusInfo.CheckGate
''        End If
''        For nTmp = 1 To frmGateMoniter.lvCheckGate.ListItems.Count   '刷新检票口状态
''            If frmGateMoniter.lvCheckGate.ListItems(nTmp).Text = tTmpBusInfo.CheckGate Then
''                frmGateMoniter.lvCheckGate.ListItems(nTmp).SmallIcon = 2
''            End If
''        Next nTmp
''    End If
''End Sub
''
