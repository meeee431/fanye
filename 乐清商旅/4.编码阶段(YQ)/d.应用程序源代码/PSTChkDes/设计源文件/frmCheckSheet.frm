VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{A5E8F770-DA22-4EAF-B7BE-73B06021D09F}#1.1#0"; "ST6Report.ocx"
Begin VB.Form frmCheckSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "·��"
   ClientHeight    =   6330
   ClientLeft      =   1245
   ClientTop       =   2025
   ClientWidth     =   9105
   Icon            =   "frmCheckSheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton cmdPrint 
      Height          =   345
      Left            =   6540
      TabIndex        =   0
      Top             =   5790
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
      MICON           =   "frmCheckSheet.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   7830
      TabIndex        =   1
      Top             =   5790
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ȡ��"
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
      MICON           =   "frmCheckSheet.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmStart 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5220
      Top             =   5760
   End
   Begin ST6Report.RTReport RTReport2 
      Height          =   735
      Left            =   510
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1296
   End
   Begin ST6Report.RTReport RTReport1 
      Height          =   5385
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   9499
   End
End
Attribute VB_Name = "frmCheckSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const csNoPrintPrompt = "��δ��ӡ·�����˳���?"
Private Const csMsgBoxTitle = "·��"

Const SheetGridLines = 10 'ԭ����7�е�
Public moChkTicket As CheckTicket      '��Ʊ����
Public g_oActiveUser As ActiveUser
Public mszSheetID As String             '·����
Public mbNoPrintPrompt As Boolean       '����NoPrintPrompt���ԣ���δ��ӡ�˳�ʱ�Ƿ���ʾ
Public mbExitAfterPrint As Boolean      '����ExitAfterPrint���ԣ������ӡ���Ƿ��˳�
Public mbViewMode As Boolean            '����ViewMode���ԣ�������ʾģʽ/���ܴ�ӡ


Private mbHasPrint As Boolean
Private mrsSheetData As Recordset   '·����¼��
Private maszSheetCustom() As String    '·���е��Զ�������





'��Ʊ����
Enum ECheckedTicketStatus
    NormalTicket = 1
    ChangedTicket = 2
    MergedTicket = 3
End Enum
Private Sub cmdExit_Click()
'On Error Resume Next
    If Not mbNoPrintPrompt And Not mbHasPrint Then
        If Not MsgboxEx(csNoPrintPrompt, vbYesNoCancel + vbQuestion, _
            csMsgBoxTitle) _
            = vbYes Then
            Exit Sub
        End If
    End If
    tmStart.Enabled = True
    Unload Me
End Sub

'���·������
Private Sub FillSheetReport()
    RTReport1.TemplateFile = App.Path & "\csshow.cll"
    RTReport1.ShowReport mrsSheetData, maszSheetCustom
    
    RTReport2.TemplateFile = App.Path & "\csprint.cll"
    RTReport2.ShowReport mrsSheetData, maszSheetCustom
    
End Sub
'��ӡ·������
Public Sub PrintSheetReport()
'    RTReport2.TemplateFile = App.Path & "\csprint.xls"
'    RTReport2.ShowReport mrsSheetData, maszSheetCustom
    On Error Resume Next
    RTReport2.PrintReport

    mbHasPrint = True
    
End Sub
Private Sub cmdPrint_Click()
On Error GoTo ErrHandle
    PrintSheetReport
    Unload Me
    
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub





Private Sub Form_Load()
On Error GoTo ErrHandle
    AlignFormPos Me
    If moChkTicket Is Nothing Then
        Set moChkTicket = New CheckTicket
        moChkTicket.Init g_oActiveUser
    End If

    tmStart.Enabled = False
    GetCheckSheetInfo
    FillSheetReport
    
    
    If mbViewMode Then
        '��ʾģʽ
        cmdPrint.Enabled = False
    Else
        cmdPrint.Enabled = True
        EvisibleCloseButton Me
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Unload Me
End Sub

'Private Sub RTReport1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyEscape Then
'        cmdExit_Click
'    End If
'End Sub

Private Sub tmStart_Timer()
''On Error GoTo ErrHandle
''    tmStart.Enabled = False
''    GetCheckSheetInfo
''    FillSheetReport
''
''    If mbViewMode Then
''        '��ʾģʽ
''        cmdPrint.Enabled = False
''    Else
''        cmdPrint.Enabled = True
''    End If
''    Exit Sub
''ErrHandle:
''    ShowErrorMsg
End Sub

'Public Sub GetCheckSheetInfo()
''********************************************************************
''ȡ��ָ��·�������еļ�Ʊ��Ϣ����ϸ·����Ϣ
''********************************************************************
'    Dim atSheetResult()  As TCheckSheetStationInfo
'    Dim tSheetInfo As TCheckSheetInfo
'    Dim nCount As Integer
'    Dim szStation As String
'    Dim i As Integer, j As Integer
'
'    tSheetInfo = moChkTicket.GetCheckSheetInfo(mszSheetID)
'    '�����Զ�����Ŀ
'    ReDim maszSheetCustom(1 To 10, 1 To 2)
'
'    '���ó�����Ϣ
'    Dim oVehicle As Vehicle
'    Set oVehicle = New Vehicle
'    oVehicle.Init g_oActiveUser
'    oVehicle.Identify tSheetInfo.szVehicleId
'    maszSheetCustom(1, 1) = "·����"
'    maszSheetCustom(1, 2) = mszSheetID
'    maszSheetCustom(2, 1) = "������λ"
'    maszSheetCustom(2, 2) = Trim(oVehicle.CompanyName)
'    maszSheetCustom(3, 1) = "����"
'    maszSheetCustom(3, 2) = Trim(oVehicle.LicenseTag)
'    maszSheetCustom(4, 1) = "����"
'    maszSheetCustom(4, 2) = Trim(tSheetInfo.szBusid) & IIf(tSheetInfo.nBusSerialNo > 0, "-" & tSheetInfo.nBusSerialNo, "")
'    maszSheetCustom(5, 1) = "����ʱ��"
'    maszSheetCustom(5, 2) = Format(tSheetInfo.dtStartUpTime, "HH:mm")
'    Dim oRoute As Route
'    Set oRoute = New Route
'    oRoute.Init g_oActiveUser
'    oRoute.Identify Trim(tSheetInfo.szRouteID)
'    maszSheetCustom(6, 1) = "��·"
'    maszSheetCustom(6, 2) = Trim(oRoute.RouteName)
'
'    '�õ���ƱԱ����
'    Dim szChecker As String
'    If tSheetInfo.szMakeSheetUser = g_oActiveUser.UserID Then
'        szChecker = MakeDisplayString(tSheetInfo.szMakeSheetUser, g_oActiveUser.UserName)
'    Else
'        Dim aszUsers() As String
'        aszUsers = moChkTicket.GetAllUser
'        For i = 1 To ArrayLength(aszUsers)
'            If aszUsers(i, 1) = tSheetInfo.szMakeSheetUser Then
'                szChecker = MakeDisplayString(tSheetInfo.szMakeSheetUser, aszUsers(i, 2))
'                Exit For
'            End If
'        Next i
'    End If
'    maszSheetCustom(7, 1) = "ǩ����"
'    maszSheetCustom(7, 2) = szChecker
'
''    maszSheetCustom(8, 1) = "��λ"
''    maszSheetCustom(8, 2) = Trim(m_oSysParam.LocalUnit.szUnitShortName)
''    maszSheetCustom(9, 1) = "����"
''    maszSheetCustom(9, 2) = Trim(LblPiece.Caption)
'    '�õ�·��վ����ϸ��Ϣ
'    atSheetResult = moChkTicket.GetCheckSheetStationInfo(mszSheetID)
'    nCount = ArrayLength(atSheetResult)
'    Dim aszSheetInfo() As String
'    If nCount > 0 Then
'        ReDim aszSheetInfo(1 To nCount, 1 To 12)
'    End If
'    j = 0
'    For i = 1 To nCount
'        If j = 0 Then
'            aszSheetInfo(1, 1) = atSheetResult(1).szStationID
'            j = 1
'        End If
'        If atSheetResult(i).szStationID <> aszSheetInfo(j, 1) Then
'                j = j + 1
'                aszSheetInfo(j, 1) = atSheetResult(i).szStationID
'        End If
'        If atSheetResult(i).sgMileage <> ECheckedTicketStatus.NormalTicket Then
'            aszSheetInfo(j, 2) = LeftAndRight(LeftAndRight(atSheetResult(i).szCheckSheet, False, "["), True, "]") & "(�Ĳ�)"
'        Else
'            aszSheetInfo(j, 2) = Trim(LeftAndRight(LeftAndRight(atSheetResult(i).szCheckSheet, False, "["), True, "]"))
'        End If
'        If atSheetResult(i).nTicketType = TP_FullPrice Then
'            aszSheetInfo(j, 3) = atSheetResult(i).nManCount
'            aszSheetInfo(j, 4) = atSheetResult(i).sgTicketPrice
'        End If
'        If atSheetResult(i).nTicketType = TP_HalfPrice Then
'            aszSheetInfo(j, 5) = atSheetResult(i).nManCount
'            aszSheetInfo(j, 6) = atSheetResult(i).sgTicketPrice
'        End If
'        If atSheetResult(i).nTicketType = TP_PreferentialTicket1 Then
'            aszSheetInfo(j, 7) = atSheetResult(i).nManCount
'            aszSheetInfo(j, 8) = atSheetResult(i).sgTicketPrice
'        End If
'        If atSheetResult(i).nTicketType = TP_PreferentialTicket2 Then
'            aszSheetInfo(j, 9) = atSheetResult(i).nManCount
'            aszSheetInfo(j, 10) = atSheetResult(i).sgTicketPrice
'        End If
'        If atSheetResult(i).nTicketType = TP_PreferentialTicket3 Then
'            aszSheetInfo(j, 11) = atSheetResult(i).nManCount
'            aszSheetInfo(j, 12) = atSheetResult(i).sgTicketPrice
'        End If
'        If atSheetResult(i).nTicketType = TP_FreeTicket Then        '��Ʊ����ȫƱ
'            aszSheetInfo(j, 3) = Val(aszSheetInfo(j, 3)) + atSheetResult(i).nManCount
'            aszSheetInfo(j, 4) = Val(aszSheetInfo(j, 4)) + atSheetResult(i).sgTicketPrice
'        End If
'    Next i
'
'
'    '����һ����¼��
'    Set mrsSheetData = New Recordset
'    mrsSheetData.CursorLocation = adUseClient
'    '����������֧�ֵ��ֶ�
'    mrsSheetData.Fields.Append "station_name", adVarChar, 30        'վ������
''    mrsSheetData.Fields.Append "mileage", adVarChar, 30             '���
'    mrsSheetData.Fields.Append "full_number", adVarChar, 30        'ȫƱ��
'    mrsSheetData.Fields.Append "full_price", adVarChar, 30        'ȫƱ���
'    mrsSheetData.Fields.Append "half_number", adVarChar, 30        '��Ʊ��
'    mrsSheetData.Fields.Append "half_price", adVarChar, 30        '��Ʊ���
'    mrsSheetData.Fields.Append "pre1_number", adVarChar, 30        '�Ż�Ʊ1��
'    mrsSheetData.Fields.Append "pre1_price", adVarChar, 30        '�Ż�Ʊ1���
'    mrsSheetData.Fields.Append "pre2_number", adVarChar, 30        '�Ż�Ʊ2��
'    mrsSheetData.Fields.Append "pre2_price", adVarChar, 30        '�Ż�Ʊ2���
'    mrsSheetData.Fields.Append "pre3_number", adVarChar, 30        '�Ż�Ʊ3��
'    mrsSheetData.Fields.Append "pre3_price", adVarChar, 30        '�Ż�Ʊ3���
'    mrsSheetData.Fields.Append "total_number", adVarChar, 30        '�Ż�Ʊ3��
'    mrsSheetData.Fields.Append "total_price", adVarChar, 30        '�Ż�Ʊ3���
'    mrsSheetData.Open
'    Dim aszTemp(1 To 14) As String
'    For i = 1 To SheetGridLines         '�����յļ�¼��
'        mrsSheetData.AddNew
'        If i > nCount Then
'            For j = 1 To mrsSheetData.Fields.Count
'                mrsSheetData.Fields(j - 1) = ""
'            Next j
'        Else
'            mrsSheetData.Fields("station_name") = aszSheetInfo(i, 2)
'            mrsSheetData.Fields("full_number") = aszSheetInfo(i, 3)
'            mrsSheetData.Fields("full_price") = aszSheetInfo(i, 4)
'            mrsSheetData.Fields("half_number") = aszSheetInfo(i, 5)
'            mrsSheetData.Fields("half_price") = aszSheetInfo(i, 6)
'            mrsSheetData.Fields("pre1_number") = aszSheetInfo(i, 7)
'            mrsSheetData.Fields("pre1_price") = aszSheetInfo(i, 8)
'            mrsSheetData.Fields("pre2_number") = aszSheetInfo(i, 9)
'            mrsSheetData.Fields("pre2_price") = aszSheetInfo(i, 10)
'            mrsSheetData.Fields("pre3_number") = aszSheetInfo(i, 11)
'            mrsSheetData.Fields("pre3_price") = aszSheetInfo(i, 12)
'            mrsSheetData.Fields("total_number") = Val(aszSheetInfo(i, 3)) + Val(aszSheetInfo(i, 5)) + Val(aszSheetInfo(i, 7)) + Val(aszSheetInfo(i, 9)) + Val(aszSheetInfo(i, 11))
'            mrsSheetData.Fields("total_price") = Val(aszSheetInfo(i, 4)) + Val(aszSheetInfo(i, 6)) + Val(aszSheetInfo(i, 8)) + Val(aszSheetInfo(i, 10)) + Val(aszSheetInfo(i, 12))
'            '����
'            For j = 3 To 12
'                aszTemp(j) = Val(aszTemp(j)) + Val(aszSheetInfo(i, j))
'            Next j
'
'            aszTemp(13) = Val(aszTemp(13)) + Val(mrsSheetData!total_number)
'            aszTemp(14) = Val(aszTemp(14)) + Val(mrsSheetData!total_price)
'
'        End If
'        mrsSheetData.Update
'    Next i
'    '����ϼ���
'    mrsSheetData.AddNew
'    mrsSheetData.Fields("station_name") = "�ϼ�"
'    mrsSheetData.Fields("full_number") = aszTemp(3)
'    mrsSheetData.Fields("full_price") = aszTemp(4)
'    mrsSheetData.Fields("half_number") = aszTemp(5)
'    mrsSheetData.Fields("half_price") = aszTemp(6)
'    mrsSheetData.Fields("pre1_number") = aszTemp(7)
'    mrsSheetData.Fields("pre1_price") = aszTemp(8)
'    mrsSheetData.Fields("pre2_number") = aszTemp(9)
'    mrsSheetData.Fields("pre2_price") = aszTemp(10)
'    mrsSheetData.Fields("pre3_number") = aszTemp(11)
'    mrsSheetData.Fields("pre3_price") = aszTemp(12)
'    mrsSheetData.Fields("total_number") = aszTemp(13)
'    mrsSheetData.Fields("total_price") = aszTemp(14)
'    mrsSheetData.Update
'End Sub

Public Sub GetCheckSheetInfo()
'********************************************************************
'ȡ��ָ��·�������еļ�Ʊ��Ϣ����ϸ·����Ϣ
'********************************************************************
    Dim atSheetResult()  As TCheckSheetStationInfoEx2
    Dim tSheetInfo As TCheckSheetInfo
    Dim nCount As Integer
    Dim szStation As String
    Dim i As Integer, j As Integer
    Dim szChecker As String
    Dim aszSheetInfo() As String
    Dim dbTotalMan As Double
    Dim dbTotalPrice As Double
    Dim dbTotalMileage As Double
    Dim aszTemp() As String
    
    Dim oRoute As Route
    Dim oParam As New SystemParam
    Dim nSpecialTicketTypePosition As Integer
    
    On Error GoTo ErrorHandle
    
    oParam.Init g_oActiveUser
    nSpecialTicketTypePosition = Val(oParam.SpecialTicketTypePosition)
    tSheetInfo = moChkTicket.GetCheckSheetInfo(mszSheetID)
    '�����Զ�����Ŀ
    ReDim maszSheetCustom(1 To 20, 1 To 2)
    
    '���ó�����Ϣ
    Dim oVehicle As Vehicle
    Set oVehicle = New Vehicle
    oVehicle.Init g_oActiveUser
    oVehicle.Identify tSheetInfo.szVehicleId
    maszSheetCustom(1, 1) = "·����"
    maszSheetCustom(1, 2) = "[" & mszSheetID & "]"
    maszSheetCustom(2, 1) = "������λ"
    maszSheetCustom(2, 2) = Trim(oVehicle.CompanyName)
    maszSheetCustom(3, 1) = "����"
    maszSheetCustom(3, 2) = Trim(oVehicle.LicenseTag)
    maszSheetCustom(4, 1) = "����"
    maszSheetCustom(4, 2) = Trim(tSheetInfo.szBusid) & IIf(tSheetInfo.nBusSerialNo > 0, "-" & tSheetInfo.nBusSerialNo, "")
    maszSheetCustom(5, 1) = "����ʱ��"
    maszSheetCustom(5, 2) = Format(tSheetInfo.dtStartUpTime, "DD HH:mm")
    Set oRoute = New Route
    oRoute.Init g_oActiveUser
    oRoute.Identify Trim(tSheetInfo.szRouteID)
    maszSheetCustom(6, 1) = "��·"
    maszSheetCustom(6, 2) = Trim(oRoute.RouteName)
    
    
    Dim oVehicleType As New VehicleModel
    oVehicleType.Init g_oActiveUser
    oVehicleType.Identify tSheetInfo.szVehicleModelID
    maszSheetCustom(7, 1) = "����"
    maszSheetCustom(7, 2) = Trim(oVehicleType.VehicleModelShortName)
    
    
    '�õ���ƱԱ����
'    If Trim(tSheetInfo.szMakeSheetUser) = Trim(g_oActiveUser.UserID) Then
'        szChecker = Trim(g_oActiveUser.UserName) 'MakeDisplayString(tSheetInfo.szMakeSheetUser, g_oActiveUser.UserName)
'    Else
'        Dim aszUsers() As String
'        aszUsers = moChkTicket.GetAllUser
'        For i = 1 To ArrayLength(aszUsers)
'            If Trim(aszUsers(i, 1)) = Trim(tSheetInfo.szMakeSheetUser) Then
'                szChecker = Trim(aszUsers(i, 2)) 'MakeDisplayString(Trim(tSheetInfo.szMakeSheetUser), Trim(aszUsers(i, 2)))
'                Exit For
'            End If
'        Next i
'    End If
    szChecker = Trim(tSheetInfo.szMakeSheetUser)
    maszSheetCustom(8, 1) = "ǩ����"
    maszSheetCustom(8, 2) = szChecker
    
    maszSheetCustom(9, 1) = "��λ"
    maszSheetCustom(9, 2) = Trim(g_szSellStationName)
    
    
    '�õ�·��վ����ϸ��Ϣ
'    atSheetResult = moChkTicket.GetCheckSheetStationInfo(mszSheetID)
    atSheetResult = moChkTicket.GetCheckSheetStationInfoEx(mszSheetID, tSheetInfo.szBusid, tSheetInfo.dtDate, tSheetInfo.nBusSerialNo)
    nCount = ArrayLength(atSheetResult)
    
    
    dbTotalMan = 0
    dbTotalPrice = 0
    dbTotalMileage = 0
    If nCount > 0 Then
        ReDim aszSheetInfo(1 To nCount, 1 To 20)
    End If
    j = 0
    For i = 1 To nCount
        If j = 0 Then
            aszSheetInfo(1, 1) = atSheetResult(1).szStationID
            aszSheetInfo(1, 13) = atSheetResult(i).sgMileage
            aszSheetInfo(1, 14) = atSheetResult(i).szSinglePrice
            j = 1
        End If
        If i > 1 Then
            If atSheetResult(i).szStationID <> aszSheetInfo(j, 1) Or (atSheetResult(i).szStationID = aszSheetInfo(j, 1) And atSheetResult(i).nTicketType = atSheetResult(i - 1).nTicketType And Val(atSheetResult(i).szSinglePrice) <> Val(atSheetResult(i - 1).szSinglePrice)) Or (atSheetResult(i).szStationID = aszSheetInfo(j, 1) And atSheetResult(i).nTicketType <> atSheetResult(i - 1).nTicketType And Val(atSheetResult(i).szSinglePrice) <> Val(atSheetResult(i - 1).szSinglePrice) And atSheetResult(i).nTicketType = nSpecialTicketTypePosition) Then     'And atSheetResult(i).szSinglePrice <> Val(aszSheetInfo(j, 14))
                    j = j + 1
                    aszSheetInfo(j, 1) = atSheetResult(i).szStationID
                    aszSheetInfo(j, 13) = atSheetResult(i).sgMileage
                    aszSheetInfo(j, 14) = atSheetResult(i).szSinglePrice
            End If
        End If
'        If atSheetResult(i).nCheckStatus <> ECheckedTicketStatus.NormalTicket Then
'            aszSheetInfo(j, 2) = LeftAndRight(LeftAndRight(atSheetResult(i).szCheckSheet, False, "["), True, "]") & "(�Ĳ�)"
'        Else
            aszSheetInfo(j, 2) = Trim(LeftAndRight(LeftAndRight(atSheetResult(i).szCheckSheet, False, "["), True, "]"))
'        End If
        If atSheetResult(i).nTicketType = TP_FullPrice Then
            aszSheetInfo(j, 3) = atSheetResult(i).nManCount
            aszSheetInfo(j, 4) = atSheetResult(i).sgTicketPrice
            aszSheetInfo(j, 15) = atSheetResult(i).szSinglePrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = TP_HalfPrice Then
            aszSheetInfo(j, 5) = atSheetResult(i).nManCount
            aszSheetInfo(j, 6) = atSheetResult(i).sgTicketPrice
            aszSheetInfo(j, 16) = atSheetResult(i).szSinglePrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = TP_PreferentialTicket1 Then
            aszSheetInfo(j, 7) = atSheetResult(i).nManCount
            aszSheetInfo(j, 8) = atSheetResult(i).sgTicketPrice
            aszSheetInfo(j, 17) = atSheetResult(i).szSinglePrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = TP_PreferentialTicket2 Then
            aszSheetInfo(j, 9) = atSheetResult(i).nManCount
            aszSheetInfo(j, 10) = atSheetResult(i).sgTicketPrice
            aszSheetInfo(j, 18) = atSheetResult(i).szSinglePrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = TP_PreferentialTicket3 Then
            aszSheetInfo(j, 11) = atSheetResult(i).nManCount
            aszSheetInfo(j, 12) = atSheetResult(i).sgTicketPrice
            aszSheetInfo(j, 19) = atSheetResult(i).szSinglePrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = TP_FreeTicket Then        '��Ʊ����ȫƱ
            aszSheetInfo(j, 3) = Val(aszSheetInfo(j, 3)) + atSheetResult(i).nManCount
            aszSheetInfo(j, 4) = Val(aszSheetInfo(j, 4)) + atSheetResult(i).sgTicketPrice
            aszSheetInfo(j, 20) = atSheetResult(i).szSinglePrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = nSpecialTicketTypePosition And atSheetResult(i).nTicketType <> TP_FullPrice Then '�����Ʊ��Ϊ��Ʊ,�Ҳ�ΪȫƱ,�򽫸�Ʊ��һ������ȫƱ����Ϊ��Ʊ��ͯ��������ӡ��.
        
            aszSheetInfo(j, 3) = Val(aszSheetInfo(j, 3)) + atSheetResult(i).nManCount
            aszSheetInfo(j, 4) = Val(aszSheetInfo(j, 4)) + atSheetResult(i).sgTicketPrice
            aszSheetInfo(j, 15) = atSheetResult(i).szSinglePrice
'            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
'            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
'            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
'            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
'            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        
        
    Next i
    
    
    maszSheetCustom(10, 1) = "�ϼ�����"
    maszSheetCustom(10, 2) = dbTotalMan
    maszSheetCustom(11, 1) = "�ϼƴ�д����"
    aszTemp = ApartBaseFig(CStr(dbTotalMan), True)
    maszSheetCustom(11, 2) = aszTemp(1) & "[" & dbTotalMan & "]"
    
    maszSheetCustom(12, 1) = "�ϼƽ��"
    maszSheetCustom(12, 2) = dbTotalPrice
    maszSheetCustom(13, 1) = "�ϼƴ�д���"
    aszTemp = ApartFig(dbTotalPrice)
    maszSheetCustom(13, 2) = GetNumber(dbTotalPrice) & "[" & dbTotalPrice & "]"
    
    
    maszSheetCustom(14, 1) = "�ϼ����"
    maszSheetCustom(14, 2) = dbTotalMileage
    maszSheetCustom(15, 1) = "�ϼƴ�д���"
    aszTemp = ApartBaseFig(CStr(dbTotalMileage), True)
    maszSheetCustom(15, 2) = aszTemp(1) & "[" & dbTotalMileage & "]"
    
    maszSheetCustom(16, 1) = "�յ�վ"
    maszSheetCustom(16, 2) = oRoute.EndStationName
    
    maszSheetCustom(17, 1) = "��Ʊ��"
    maszSheetCustom(17, 2) = tSheetInfo.szCheckGateName
    
    maszSheetCustom(18, 1) = "�������"
    maszSheetCustom(18, 2) = Format(tSheetInfo.dtStartUpTime, "YYYY-MM-DD")
    
    maszSheetCustom(19, 1) = "��ӡ����"
    maszSheetCustom(19, 2) = Format(oParam.NowDate, "YYYY-MM-DD")
    
    maszSheetCustom(20, 1) = "��ӡʱ��"
    maszSheetCustom(20, 2) = Format(oParam.NowTime, "HH:MM:SS")

    
    '����һ����¼��
    Set mrsSheetData = New Recordset
    mrsSheetData.CursorLocation = adUseClient
    '����������֧�ֵ��ֶ�
    mrsSheetData.Fields.Append "station_name", adVarChar, 30        'վ������
'    mrsSheetData.Fields.Append "mileage", adVarChar, 30             '���
    mrsSheetData.Fields.Append "full_number", adVarChar, 30        'ȫƱ��
    mrsSheetData.Fields.Append "full_price", adVarChar, 30        'ȫƱ���
    mrsSheetData.Fields.Append "half_number", adVarChar, 30        '��Ʊ��
    mrsSheetData.Fields.Append "half_price", adVarChar, 30        '��Ʊ���
    mrsSheetData.Fields.Append "pre1_number", adVarChar, 30        '�Ż�Ʊ1��
    mrsSheetData.Fields.Append "pre1_price", adVarChar, 30        '�Ż�Ʊ1���
    mrsSheetData.Fields.Append "pre2_number", adVarChar, 30        '�Ż�Ʊ2��
    mrsSheetData.Fields.Append "pre2_price", adVarChar, 30        '�Ż�Ʊ2���
    mrsSheetData.Fields.Append "pre3_number", adVarChar, 30        '�Ż�Ʊ3��
    mrsSheetData.Fields.Append "pre3_price", adVarChar, 30        '�Ż�Ʊ3���
    mrsSheetData.Fields.Append "total_number", adVarChar, 30        '�Ż�Ʊ3��
    mrsSheetData.Fields.Append "total_price", adVarChar, 30        '�Ż�Ʊ3���
    mrsSheetData.Open
    
    ReDim aszTemp(1 To 14)
    
    For i = 1 To SheetGridLines         '�����յļ�¼��
        mrsSheetData.AddNew
        If i > nCount Then
            For j = 1 To mrsSheetData.Fields.Count
                mrsSheetData.Fields(j - 1) = ""
            Next j
        Else
            mrsSheetData.Fields("station_name") = aszSheetInfo(i, 2)
            mrsSheetData.Fields("full_number") = aszSheetInfo(i, 3)
            mrsSheetData.Fields("full_price") = aszSheetInfo(i, 15)
            mrsSheetData.Fields("half_number") = aszSheetInfo(i, 5)
            mrsSheetData.Fields("half_price") = aszSheetInfo(i, 16)
            mrsSheetData.Fields("pre1_number") = aszSheetInfo(i, 7)
            mrsSheetData.Fields("pre1_price") = aszSheetInfo(i, 17)
            mrsSheetData.Fields("pre2_number") = aszSheetInfo(i, 9)
            mrsSheetData.Fields("pre2_price") = aszSheetInfo(i, 18)
            mrsSheetData.Fields("pre3_number") = aszSheetInfo(i, 11)
            mrsSheetData.Fields("pre3_price") = aszSheetInfo(i, 19)
            mrsSheetData.Fields("total_number") = Val(aszSheetInfo(i, 3)) + Val(aszSheetInfo(i, 5)) + Val(aszSheetInfo(i, 7)) + Val(aszSheetInfo(i, 11)) '+ Val(aszSheetInfo(i, 9)) 'ȥ��ЯͯƱ������
            mrsSheetData.Fields("total_price") = Val(aszSheetInfo(i, 4)) + Val(aszSheetInfo(i, 6)) + Val(aszSheetInfo(i, 8)) + Val(aszSheetInfo(i, 12)) '+ Val(aszSheetInfo(i, 10))'ȥ��ЯͯƱ�Ľ��
            '����
            For j = 3 To 12
                aszTemp(j) = Val(aszTemp(j)) + Val(aszSheetInfo(i, j))
            Next j

            aszTemp(13) = Val(aszTemp(13)) + Val(mrsSheetData!total_number)
            aszTemp(14) = Val(aszTemp(14)) + Val(mrsSheetData!total_price)

        End If
        mrsSheetData.Update
    Next i
    '����ϼ���
    mrsSheetData.AddNew
    mrsSheetData.Fields("station_name") = "" '"�ϼ�" '�״�ʱ,����Ҫ�ϼ�
    mrsSheetData.Fields("full_number") = aszTemp(3)
    mrsSheetData.Fields("full_price") = aszTemp(4)
    mrsSheetData.Fields("half_number") = aszTemp(5)
    mrsSheetData.Fields("half_price") = aszTemp(6)
    mrsSheetData.Fields("pre1_number") = aszTemp(7)
    mrsSheetData.Fields("pre1_price") = aszTemp(8)
    mrsSheetData.Fields("pre2_number") = aszTemp(9)
    mrsSheetData.Fields("pre2_price") = aszTemp(10)
    mrsSheetData.Fields("pre3_number") = aszTemp(11)
    mrsSheetData.Fields("pre3_price") = aszTemp(12)
    mrsSheetData.Fields("total_number") = aszTemp(13)
    mrsSheetData.Fields("total_price") = aszTemp(14)
    mrsSheetData.Update
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

