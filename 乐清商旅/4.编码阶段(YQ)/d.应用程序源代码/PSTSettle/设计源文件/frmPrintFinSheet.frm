VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmPrintFinSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���㵥"
   ClientHeight    =   7800
   ClientLeft      =   1890
   ClientTop       =   2055
   ClientWidth     =   10785
   Icon            =   "frmPrintFinSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10785
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4800
      Top             =   3660
   End
   Begin RTComctl3.CoolButton cmdPrevew 
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   7350
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "��ӡԤ��(&V)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmPrintFinSheet.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdPrint 
      Height          =   345
      Left            =   8250
      TabIndex        =   0
      Top             =   7320
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "��ӡ(&P)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmPrintFinSheet.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTReportLF.RTReport RTReport1 
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   12515
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   9510
      TabIndex        =   2
      Top             =   7320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "�ر�(&E)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmPrintFinSheet.frx":0342
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmPrintFinSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_rsStationData As Recordset
Private m_oReport As New Report

Private m_aszSheetCusTom() As String
Private m_aszStationCustom() As String
Public m_SheetID As String
Public m_OldSheetID As String  '�����ش��ԭ���㵥��
Public m_bRePrint As Boolean '�Ƿ����ش���㵥
'Public m_bNeedPrint As Boolean '�Ƿ���Ҫ��ӡ

Public m_szLugSettleSheetID As String
'Const cszSplitItemCount = 20
Dim m_oSettleSheet As New SettleSheet

Private Sub cmdExit_Click()

    Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error Resume Next
    
    RTReport1.PrintReport True
    
    
    On Error GoTo 0
    On Error GoTo ErrHandle
    m_oSettleSheet.SetPrint '����Ϊ�Ѵ�ӡ��
    Unload Me
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub




Private Sub cmdPrevew_Click()
    RTReport1.PrintView
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    
End Sub

Private Sub GetFinSheetInfo()

    Dim aszSplitItem() As TSplitItemInfo
    Dim nSplitItemCount As Integer
    Dim i As Integer
'    Dim rsSheetData As Recordset
    Dim szTemp As String
    Dim rsLugSheetRs As Recordset '�а���¼
    
    Dim szObjectName As String
    
    
    Dim rsCompanySettleLst As Recordset '��˾������ϸ
    Dim rsVehicleSettleLst As Recordset '����������ϸ
    Dim rsBusSettleLst As Recordset '���ν�����ϸ
    Dim rsSettleLst As Recordset '��ʱ������ϸ
    
'    Dim rsRemark As Recordset
    Dim rsRouteQuantity As Recordset '��·����
    
    Dim arsTemp As Variant
    Dim aszTemp As Variant
    Dim rsSplitItem As Recordset
    
    Dim dbExtraManCount As Double '��Ʊ��������
    Dim dbExtraTotalPrice As Double '��Ʊ����Ʊ��
    
    
    On Error GoTo ErrHandle
    
    m_oReport.Init g_oActiveUser
    If m_bRePrint Then
        m_oReport.ReprintSettleSheet m_OldSheetID, m_SheetID
    End If
    
    Set rsCompanySettleLst = m_oReport.GetSettleCompanyLst(m_SheetID)
    Set rsVehicleSettleLst = m_oReport.GetSettleVehicleLst(m_SheetID)
    Set rsBusSettleLst = m_oReport.GetSettleBusLst(m_SheetID)
    
    m_oSettleSheet.Init g_oActiveUser
    m_oSettleSheet.Identify m_SheetID
    aszSplitItem = m_oReport.GetSplitItemInfo
    nSplitItemCount = ArrayLength(aszSplitItem)
    '�ض������鳤��
    ReDim m_aszSheetCusTom(1 To nSplitItemCount + g_cnSplitItemCount + 29, 1 To 2)
        


    If m_oSettleSheet.SettleObject = CS_SettleByTransportCompany Then
        RTReport1.TemplateFile = App.Path & "\����˾·�����㵥.xls"
    ElseIf m_oSettleSheet.SettleObject = CS_SettleByVehicle Then
        RTReport1.TemplateFile = App.Path & "\������·�����㵥.xls"
    Else: RTReport1.TemplateFile = App.Path & "\������·�����㵥.xls"
    End If
    If rsCompanySettleLst.RecordCount + rsVehicleSettleLst.RecordCount > 1 Then
        Set rsSettleLst = rsVehicleSettleLst
    Else
        If m_oSettleSheet.SettleObject = CS_SettleByTransportCompany Then
            Set rsSettleLst = rsCompanySettleLst
        ElseIf m_oSettleSheet.SettleObject = CS_SettleByVehicle Then
            Set rsSettleLst = rsVehicleSettleLst
        Else: Set rsSettleLst = rsBusSettleLst
        End If
    End If
    For i = 1 To g_cnSplitItemCount
        m_aszSheetCusTom(i, 1) = "��Ŀ" & i
    Next i
    If rsSettleLst.RecordCount > 0 Then
        For i = 1 To g_cnSplitItemCount
            m_aszSheetCusTom(i, 2) = FormatDbValue(rsSettleLst.Fields("split_item_" & i)) ' IIf(FormatDbValue(rsSettleLst.Fields("split_item_" & i)) = 0, "", FormatDbValue(rsSettleLst.Fields("split_item_" & i)))
        Next i
    End If
    For i = 1 To nSplitItemCount
        m_aszSheetCusTom(g_cnSplitItemCount + i, 1) = "������" & i
        m_aszSheetCusTom(g_cnSplitItemCount + i, 2) = aszSplitItem(i).SplitItemName
    Next i
    nSplitItemCount = g_cnSplitItemCount + nSplitItemCount
    m_aszSheetCusTom(nSplitItemCount + 1, 1) = "����ʱ��"
    m_aszSheetCusTom(nSplitItemCount + 1, 2) = Format(m_oSettleSheet.SettleStartDate, "yyyy-MM-dd") & " - " & Format(m_oSettleSheet.SettleEndDate, "yyyy-MM-dd")
    m_aszSheetCusTom(nSplitItemCount + 2, 1) = "����·��"
    m_aszSheetCusTom(nSplitItemCount + 2, 2) = m_oSettleSheet.CheckSheetCount
    m_aszSheetCusTom(nSplitItemCount + 3, 1) = "���վ"
    m_aszSheetCusTom(nSplitItemCount + 3, 2) = g_oActiveUser.UserUnitName
    
    '�õ���Ʊ��������
    dbExtraManCount = GetExtraManCount(m_SheetID)
    
    '�õ���Ʊ����Ʊ��
    dbExtraTotalPrice = GetExtraTotalPrice(m_SheetID)
    
    m_aszSheetCusTom(nSplitItemCount + 4, 1) = "������"
    m_aszSheetCusTom(nSplitItemCount + 4, 2) = m_oSettleSheet.TotalQuantity + dbExtraManCount
    m_aszSheetCusTom(nSplitItemCount + 5, 1) = "��Ʊ��"
    m_aszSheetCusTom(nSplitItemCount + 5, 2) = m_oSettleSheet.TotalTicketPrice
    m_aszSheetCusTom(nSplitItemCount + 6, 1) = "������"
    m_aszSheetCusTom(nSplitItemCount + 6, 2) = m_oSettleSheet.Settler
    m_aszSheetCusTom(nSplitItemCount + 7, 1) = "�������"
    m_aszSheetCusTom(nSplitItemCount + 7, 2) = m_oSettleSheet.ObjectName
    m_aszSheetCusTom(nSplitItemCount + 8, 1) = "��·��ʽ"
    
'
'
'    '�õ���·�ļ��㹫ʽ
'    Set rsRemark = m_oReport.GetSettleRouteCalRemark(m_SheetID)
'    For i = 1 To rsRemark.RecordCount
'        If m_oSettleSheet.SettleObject = CS_SettleByTransportCompany Then
'            m_aszSheetCusTom(nSplitItemCount + 8, 2) = m_aszSheetCusTom(nSplitItemCount + 8, 2) & "��·[" & FormatDbValue(rsRemark!route_name) & "]����[" & FormatDbValue(rsRemark!vehicle_type_name) & "]:" & FormatDbValue(rsRemark!Annotation) & Chr(10)
'        Else
'            m_aszSheetCusTom(nSplitItemCount + 8, 2) = m_aszSheetCusTom(nSplitItemCount + 8, 2) & "��·[" & FormatDbValue(rsRemark!route_name) & "]:" & FormatDbValue(rsRemark!Annotation) & Chr(10)
'        End If
'        rsRemark.MoveNext
'    Next i
    m_aszSheetCusTom(nSplitItemCount + 9, 1) = "�а��˷�"
    m_aszSheetCusTom(nSplitItemCount + 10, 1) = "�а�������"
    m_aszSheetCusTom(nSplitItemCount + 11, 1) = "�а�����Э��"
    m_aszSheetCusTom(nSplitItemCount + 12, 1) = "�а���д���"
    m_aszSheetCusTom(nSplitItemCount + 13, 1) = "�а����㵥��"
    
    If rsSettleLst.RecordCount > 0 Then
        m_aszSheetCusTom(nSplitItemCount + 9, 2) = m_oSettleSheet.LuggageTotalBaseCarriage 'm_dbTotalPrice  'FormatDbValue(rsSettleLst!luggage_base_carriage) '
        m_aszSheetCusTom(nSplitItemCount + 10, 2) = m_oSettleSheet.LuggageTotalSettlePrice 'm_dbNeedSplitPrice 'FormatDbValue(rsSettleLst!luggage_settle_price) '
        m_aszSheetCusTom(nSplitItemCount + 11, 2) = m_oSettleSheet.LuggageProtocolName  'FormatDbValue(rsSettleLst!luggage_protocol_name) '
        m_aszSheetCusTom(nSplitItemCount + 12, 2) = GetNumber(m_oSettleSheet.LuggageTotalSettlePrice)  'GetNumber(FormatDbValue(rsSettleLst!luggage_settle_price)) '
        m_aszSheetCusTom(nSplitItemCount + 13, 2) = m_szLugSettleSheetID 'FormatDbValue(rsSettleLst!luggage_settle_id) '
        
    End If
    
    m_aszSheetCusTom(nSplitItemCount + 14, 1) = "·�����㵥��"
    m_aszSheetCusTom(nSplitItemCount + 14, 2) = m_SheetID
    
    
    Set rsRouteQuantity = m_oReport.GetSettleRouteQuantity(m_SheetID)
    
        
    m_aszSheetCusTom(nSplitItemCount + 15, 1) = "ʵ����"
    m_aszSheetCusTom(nSplitItemCount + 15, 2) = Val(m_oSettleSheet.LuggageTotalSettlePrice) + m_oSettleSheet.SettleLocalCompanyPrice ' Val(m_dbNeedSplitPrice) + m_oSettleSheet.SettleLocalCompanyPrice
    m_aszSheetCusTom(nSplitItemCount + 16, 1) = "��д���"
    m_aszSheetCusTom(nSplitItemCount + 16, 2) = GetNumber(Val(m_oSettleSheet.LuggageTotalSettlePrice) + m_oSettleSheet.SettleLocalCompanyPrice) 'GetNumber(Val(m_dbNeedSplitPrice) + m_oSettleSheet.SettleLocalCompanyPrice)
    
    If rsVehicleSettleLst.RecordCount > 0 Then
        m_aszSheetCusTom(nSplitItemCount + 17, 1) = "����"
        m_aszSheetCusTom(nSplitItemCount + 17, 2) = FormatDbValue(rsVehicleSettleLst!object_name)


'        m_aszSheetCusTom(nSplitItemCount + 18, 1) = "����Է�"
'        m_aszSheetCusTom(nSplitItemCount + 18, 2) = FormatDbValue(rsVehicleSettleLst!settle_other_price)
'        m_aszSheetCusTom(nSplitItemCount + 19, 1) = "�����վ"
'        m_aszSheetCusTom(nSplitItemCount + 19, 2) = FormatDbValue(rsVehicleSettleLst!settle_station_price)

        Dim oVehicle As New Vehicle
        oVehicle.Init g_oActiveUser
        oVehicle.Identify m_oSettleSheet.ObjectID
        
        m_aszSheetCusTom(nSplitItemCount + 18, 1) = "����"
        m_aszSheetCusTom(nSplitItemCount + 18, 2) = oVehicle.OwnerName
    End If
    On Error Resume Next
    If rsBusSettleLst.RecordCount > 0 Then
        m_aszSheetCusTom(nSplitItemCount + 17, 1) = "����"
        m_aszSheetCusTom(nSplitItemCount + 17, 2) = FormatDbValue(rsBusSettleLst!bus_id)


'        m_aszSheetCusTom(nSplitItemCount + 18, 1) = "����Է�"
'        m_aszSheetCusTom(nSplitItemCount + 18, 2) = FormatDbValue(rsVehicleSettleLst!settle_other_price)
'        m_aszSheetCusTom(nSplitItemCount + 19, 1) = "�����վ"
'        m_aszSheetCusTom(nSplitItemCount + 19, 2) = FormatDbValue(rsVehicleSettleLst!settle_station_price)
        
        Dim oBus As New Bus
        '���ǵ��Ӱ೵���޸üƻ�����,����Ҫ���ϴ��󲶻�

        oBus.Init g_oActiveUser
        oBus.Identify m_oSettleSheet.ObjectID
        
        m_aszSheetCusTom(nSplitItemCount + 18, 1) = "����ʱ��"
        m_aszSheetCusTom(nSplitItemCount + 18, 2) = Format(oBus.StartUpTime, "hh:mm")
        
    
    End If
    
    On Error GoTo 0
    On Error GoTo ErrHandle
    m_aszSheetCusTom(nSplitItemCount + 19, 1) = "���վ��"
    m_aszSheetCusTom(nSplitItemCount + 19, 2) = m_oSettleSheet.SettleStationPrice
    m_aszSheetCusTom(nSplitItemCount + 20, 1) = "����Է�"
    m_aszSheetCusTom(nSplitItemCount + 20, 2) = m_oSettleSheet.SettleOtherCompanyPrice
    
    
    m_aszSheetCusTom(nSplitItemCount + 21, 1) = "��λ"
    m_aszSheetCusTom(nSplitItemCount + 21, 2) = g_oActiveUser.UserUnitName
    m_aszSheetCusTom(nSplitItemCount + 22, 1) = "�������"
    m_aszSheetCusTom(nSplitItemCount + 22, 2) = Format(m_oSettleSheet.SettleStartDate, "yyyy")
    
    m_aszSheetCusTom(nSplitItemCount + 23, 1) = "�����·�"
    m_aszSheetCusTom(nSplitItemCount + 23, 2) = Format(m_oSettleSheet.SettleStartDate, "MM")
        
    m_aszSheetCusTom(nSplitItemCount + 24, 1) = "��·����"
    m_aszSheetCusTom(nSplitItemCount + 24, 2) = m_oSettleSheet.RouteName
        
    m_aszSheetCusTom(nSplitItemCount + 25, 1) = "��������"
    m_aszSheetCusTom(nSplitItemCount + 25, 2) = ToDBDate(m_oSettleSheet.SettleDate)
        
    m_aszSheetCusTom(nSplitItemCount + 26, 1) = "���˹�˾"
    m_aszSheetCusTom(nSplitItemCount + 26, 2) = m_oSettleSheet.TransportCompanyName
    
    m_aszSheetCusTom(nSplitItemCount + 27, 1) = "��ע"
    m_aszSheetCusTom(nSplitItemCount + 27, 2) = m_oSettleSheet.Annotation
    
    m_aszSheetCusTom(nSplitItemCount + 28, 1) = "��ӡ����"
    m_aszSheetCusTom(nSplitItemCount + 28, 2) = IIf(m_oSettleSheet.IsPrint = 0, "", "ע�����ǵ�" & m_oSettleSheet.IsPrint + 1 & "�δ�ӡ����ע��˶ԡ�")
    
    m_aszSheetCusTom(nSplitItemCount + 29, 1) = "·����"
    m_aszSheetCusTom(nSplitItemCount + 29, 2) = m_oSettleSheet.CheckSheetCount
    
    If rsVehicleSettleLst.RecordCount + rsCompanySettleLst.RecordCount > 1 Then
        Set m_rsStationData = rsCompanySettleLst
    Else
        Set m_rsStationData = rsRouteQuantity
    End If
    
    WriteProcessBar True, , , "�����γɱ���..."
'    RTReport1.CustomString

    '����ɾ����
    Dim rsTemp As Recordset
    If m_oSettleSheet.SettleObject = CS_SettleByVehicle Or m_oSettleSheet.SettleObject = CS_SettleByBus Then
          
        Dim rsCheckSheetTemp As Recordset
        Set rsCheckSheetTemp = m_oReport.GetCheckSheetInfo(m_oSettleSheet.SettleSheetID)
        Dim m_aszCheckSheetID() As String
        ReDim m_aszCheckSheetID(1 To rsCheckSheetTemp.RecordCount)
        rsCheckSheetTemp.MoveFirst
        For i = 1 To rsCheckSheetTemp.RecordCount
        m_aszCheckSheetID(i) = FormatDbValue(rsCheckSheetTemp!check_sheet_id)
        rsCheckSheetTemp.MoveNext
        Next i
        Set rsTemp = m_oReport.GetCheckSheetStationListEx(m_aszCheckSheetID, False)
          
'        Set rsTemp = m_oReport.TotalSettleStationQuantity(m_oSettleSheet.SettleSheetID)
    
    
        '����¼����Ϊ����
        Set m_rsStationData = MakeRecordset(rsTemp, dbExtraManCount, dbExtraTotalPrice)
    
    End If
    

    RTReport1.TopLabelVisual = True
    RTReport1.LeftLabelVisual = True
    RTReport1.ShowReport m_rsStationData, m_aszSheetCusTom
    
'    cmdPrint.Enabled = Not m_oSettleSheet.IsPrint
'    cmdPrevew.Enabled = Not m_oSettleSheet.IsPrint
    
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    WriteProcessBar False, , , ""
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    SetBusy
    
    
    
    GetFinSheetInfo
    SetNormal
End Sub

Private Function MakeRecordset(prsInfo As Recordset, pdbExtraManCount As Double, pdbExtraTotalPrice As Double) As Recordset
    
    Const cnColNumber = 2 '����
    Dim rsData As New Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nTemp As Integer
    
    '�½���¼��
    With rsData.Fields
        For j = 1 To cnColNumber
            .Append "station_name_" & j, adVarChar, 10
'            .Append "ticket_type_name_" & j, adChar, 12
            .Append "quantity_" & j, adChar, 8
            .Append "ticket_price_" & j, adChar, 12
        Next j
           
    End With
    '����¼��
    
    rsData.Open
    
    For i = 1 To prsInfo.RecordCount
        If i Mod cnColNumber = 1 Then
            '����һ��
            rsData.AddNew
            rsData!station_name_1 = FormatDbValue(prsInfo!station_name)
'            rsData!ticket_type_name_1 = FormatDbValue(prsInfo!ticket_type_name)
            rsData!quantity_1 = Trim(Str(FormatDbValue(prsInfo!Quantity)))
            rsData!ticket_price_1 = FormatDbValue(prsInfo!ticket_price)
            
            For j = 2 To cnColNumber
                rsData("station_name_" & j) = ""
'                rsData("ticket_type_name_" & j) = ""
                rsData("quantity_" & j) = ""
                rsData("ticket_price_" & j) = ""
            Next j
        Else
            nTemp = (i Mod cnColNumber)
            If nTemp = 0 Then nTemp = cnColNumber
            rsData("station_name_" & nTemp) = FormatDbValue(prsInfo!station_name)
'            rsData("ticket_type_name_" & nTemp) = FormatDbValue(prsInfo!ticket_type_name)
            rsData("quantity_" & nTemp) = Trim(Str(FormatDbValue(prsInfo!Quantity)))
            rsData("ticket_price_" & nTemp) = FormatDbValue(prsInfo!ticket_price)
        End If
        prsInfo.MoveNext
    Next i
    
    If pdbExtraManCount > 0 Then
        
        If i Mod cnColNumber = 1 Then
            '����һ��
            rsData.AddNew
            rsData!station_name_1 = "�ֹ���Ʊ"
'            rsData!ticket_type_name_1 = ""
            rsData!quantity_1 = Trim(Str(pdbExtraManCount))
            rsData!ticket_price_1 = pdbExtraTotalPrice '""
            
            For j = 2 To cnColNumber
                rsData("station_name_" & j) = ""
'                rsData("ticket_type_name_" & j) = ""
                rsData("quantity_" & j) = ""
                rsData("ticket_price_" & j) = ""
            Next j
        Else
            nTemp = (i Mod cnColNumber)
            If nTemp = 0 Then nTemp = cnColNumber
            rsData("station_name_" & nTemp) = "�ֹ���Ʊ"
'            rsData("ticket_type_name_" & nTemp) = ""
            rsData("quantity_" & nTemp) = Trim(Str(pdbExtraManCount))
            rsData("ticket_price_" & nTemp) = pdbExtraTotalPrice '""
        End If
    
    
    End If
    
    
    Set MakeRecordset = rsData

End Function

'�õ��ֹ���Ʊ��������
Private Function GetExtraManCount(pszSettleSheetID As String) As Double
    
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim dbTemp As Double
    Set rsTemp = m_oReport.GetExtraInfo(pszSettleSheetID)
    If rsTemp.RecordCount = 0 Then Exit Function
    dbTemp = 0
    '�����
    For i = 1 To rsTemp.RecordCount
        dbTemp = dbTemp + FormatDbValue(rsTemp!passenger_number)
        rsTemp.MoveNext
    Next i
    GetExtraManCount = dbTemp
    
    
End Function

'�õ��ֹ���Ʊ����Ʊ��
Private Function GetExtraTotalPrice(pszSettleSheetID As String) As Double
    
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim dbTemp As Double
    Set rsTemp = m_oReport.GetExtraInfo(pszSettleSheetID)
    If rsTemp.RecordCount = 0 Then Exit Function
    dbTemp = 0
    '�����
    For i = 1 To rsTemp.RecordCount
        dbTemp = dbTemp + FormatDbValue(rsTemp!total_ticket_price)
        rsTemp.MoveNext
    Next i
    GetExtraTotalPrice = dbTemp
    
    
End Function


