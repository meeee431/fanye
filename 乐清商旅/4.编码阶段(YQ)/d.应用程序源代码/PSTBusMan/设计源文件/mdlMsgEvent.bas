Attribute VB_Name = "mdlMsgEvent"
'''�¼�����ģ��
''Option Explicit
''
''Public Enum eEventId
''    AddBus = 1              '��ӳ���
''    AdjustTime = 2          '����ʱ��
''    ChangeBusCheckGate = 3  '���ļ�Ʊ��
''    ChangeBusSeat = 4       '������λ
''    ChangeBusStandCount = 5 'ĳ���ε�վƱ���ı�
''    ChangeBusTime = 6       '���ĳ��η���ʱ��
''    ChangeParam = 7         '���Ĳ���
''    ExStartCheckBus = 8     '���쳵��
''    MergeBus = 9            '���β���
''    RemoveBus = 10          'ɾ������
''    ResumeBus = 11          '���θ���
''    StartCheckBus = 12      '���쳵��
''    StopBus = 13            '����ͣ��
''    StopCheckBus = 14       'ͣ�쳵��
''End Enum
''Public Sub RunMsgEvent(EventMode As eEventId, EventParam() As String)
'''    ' ����ע��
'''    ' *************************************
'''    ' EventMode:��Ϣ���ͣ���ϢId��
'''    ' EventParam:��������
'''    ' *************************************
''
''    Dim nTmp As Integer
''    Dim tTmpBusInfo As tCheckBusLstInfo
''    Select Case EventMode
''        Case eEventId.AddBus    '��������
''            Event_AddBus EventParam(1)
''
''        Case eEventId.ChangeBusCheckGate    '���ĳ��μ�Ʊ��
''            Event_ChangeBusCheckGate EventParam(1), EventParam(3)
''
''        Case eEventId.ChangeBusTime     '���ĳ���ʱ��
''            Event_ChangeBusTime EventParam(1), Format(EventParam(3), cszDateTimeStr)
''
''        Case eEventId.MergeBus          '���೵��
''            Event_MergeBus EventParam(1)
''
''        Case eEventId.RemoveBus         'ɾ������
''            Event_RemoveBus EventParam(1)
''
''        Case eEventId.ResumeBus         '���θ���
''            Event_ResumeBus EventParam(1)
''
''        Case eEventId.StopBus           '����ͣ��
''            Event_StopBus EventParam(1)
''
''        Case eEventId.StartCheckBus     '���ο���
''            Event_StartCheckBus EventParam(1), Val(EventParam(3))
''
''        Case eEventId.ExStartCheckBus       '���β���
''            Event_ExCheckBus EventParam(1)
''
''        Case eEventId.StopCheckBus      '����ͣ��
''            Event_StopCheckBus EventParam(1)
''
''    End Select
''End Sub
''
''Private Sub Event_AddBus(szBusId As String)
''    '��Ʊ���ʱ����
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
''        '�������Ϣ������ʾ��ˢ��֮
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''        End If
''    End If
''End Sub
''Private Sub Event_RemoveBus(szBusId As String)
''    '��Ʊ���ʱ����
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
''        '�������Ϣ������ʾ��ˢ��֮
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = szCheckGateID Then
''            frmGateMoniter.RefreshlvCheckBus szCheckGateID
''        End If
''    End If
''End Sub
''Private Sub Event_StopBus(szBusId As String)
''    '��Ʊ���ʱ����
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then Exit Sub
''
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp) '����״̬
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
''    '��Ʊ���ʱ����
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then Exit Sub
''
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp) '����״̬
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
''    '��Ʊ���ʱ����
''    If frmGateMoniter.IsShow Then
''        Dim nTmp As Integer
''        Dim tTmpBusInfo As tCheckBusLstInfo
''        nTmp = frmGateMoniter.m_cBusInfo.FindItem(szBusId)
''        If nTmp = 0 Then Exit Sub
''
''        tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp) '����״̬
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
''    '��Ʊ���ʱ����
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
''        If nTmp > 0 Then    '������һ��Ʊ���η���ʱ��
''            If m_cChkingBusInfo.Item(nTmp).BusID = szBusId Then
''                tTmpBusInfo = m_cChkingBusInfo
''                tTmpBusInfo.StartupTime = dtNewTime
''                m_cChkingBusInfo.UpdateOne tTmpBusInfo
''            End If
''        Else    '�õ���һ��Ʊ����
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
''    '��Ʊ���ʱ����
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
''        nTmp = m_cChkingBusInfo.FindItemByGate(szOldGate)    'ˢ�¸ü�Ʊ�ڵ�����/��һ��Ʊ������Ϣ
''        If nTmp > 0 Then
''            tTmpBusInfo = frmGateMoniter.m_cBusInfo.Item(nTmp)
''            If tTmpBusInfo.BusID = szBusId Then
''                frmGateMoniter.GetChkingBusInfo szOldGate    '�õ����µ��ڼ쳵��
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
''    '��Ʊ���ʱ����
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
''            'ˢ������/��һ��Ʊ������Ϣ
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''            frmGateMoniter.RefreshChkingBus tTmpBusInfo.CheckGate
''        End If
''        For nTmp = 1 To frmGateMoniter.lvCheckGate.ListItems.Count   'ˢ�¼�Ʊ��״̬
''            If frmGateMoniter.lvCheckGate.ListItems(nTmp).Text = tTmpBusInfo.CheckGate Then
''                frmGateMoniter.lvCheckGate.ListItems(nTmp).SmallIcon = 1
''            End If
''        Next nTmp
''    End If
''End Sub
''Private Sub Event_ExCheckBus(szBusId As String)
''    '��Ʊ���ʱ����
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
''        'ˢ������/��һ��Ʊ������Ϣ
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''            frmGateMoniter.RefreshChkingBus tTmpBusInfo.CheckGate
''        End If
''        For nTmp = 1 To frmGateMoniter.lvCheckGate.ListItems.Count   'ˢ�¼�Ʊ��״̬
''            If frmGateMoniter.lvCheckGate.ListItems(nTmp).Text = tTmpBusInfo.CheckGate Then
''                frmGateMoniter.lvCheckGate.ListItems(nTmp).SmallIcon = 1
''            End If
''        Next nTmp
''    End If
''End Sub
''
''Private Sub Event_StopCheckBus(szBusId As String)
''    '��Ʊ���ʱ����
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
''        If nTmp = 0 Then    '����Ǹó��β���(����/��һ��Ʊ����)
''            Exit Sub
''        End If
''        frmGateMoniter.GetChkingBusInfo tTmpBusInfo.CheckGate    '�õ����£�����/��һ��Ʊ���Σ�
''
''        'ˢ������/��һ��Ʊ������Ϣ
''        If frmGateMoniter.lvCheckGate.SelectedItem.Text = tTmpBusInfo.CheckGate Then
''            frmGateMoniter.RefreshlvCheckBus tTmpBusInfo.CheckGate
''            frmGateMoniter.RefreshChkingBus tTmpBusInfo.CheckGate
''        End If
''        For nTmp = 1 To frmGateMoniter.lvCheckGate.ListItems.Count   'ˢ�¼�Ʊ��״̬
''            If frmGateMoniter.lvCheckGate.ListItems(nTmp).Text = tTmpBusInfo.CheckGate Then
''                frmGateMoniter.lvCheckGate.ListItems(nTmp).SmallIcon = 2
''            End If
''        Next nTmp
''    End If
''End Sub
''
