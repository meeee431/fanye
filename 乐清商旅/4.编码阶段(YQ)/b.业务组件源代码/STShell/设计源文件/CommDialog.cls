VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public m_oActiveUser As ActiveUser
Private WithEvents moShell As RTShell.SelectBox
Attribute moShell.VB_VarHelpID = -1
Private Enum eLugFormStatus
    AddStatus = 0
    ModifyStatus = 1
    ShowStatus = 2
    NotStatus = 3
End Enum



Private Sub Class_Initialize()
    Set moShell = New RTShell.SelectBox
End Sub

Private Sub Class_Terminate()
    Set moShell = Nothing
End Sub

'所有的类都有的私有类变量
'--------------------------->>
Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
'    m_oReg.Init "StationNet20", HKEY_CURRENT_USER, "Software\Topsoft" 'HKEY_LOCAL_MACHINE
End Sub
Private Sub AssertActiveUserIsNotThing()
    'If m_oActiveUser Is Nothing Then ShowError ERR_NoActiveUser
    If m_oActiveUser Is Nothing Then err.Raise ERR_ActiveUser, , "还未设置活动用户对象"
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    Set m_oActiveUser = vNewValue
    AssertActiveUserIsNotThing
End Property

'选择地区
Public Function SelectArea(Optional MultiSelect As Boolean = False) As String()
On Error GoTo ErrHandle
    Const cnColumnCount = 2
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "地区代码"
    aszColumns(2) = "地区名称"
    
    Set oBase = New BaseInfo
    oBase.Init m_oActiveUser
    aszTmp = oBase.GetAllArea()
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = aszTmp(i, 1)
        aszListDatas(i, 2) = aszTmp(i, 2)
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "地区列表"
    moShell.Title = "选择地区"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 1
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "area"
    moShell.ItemSmallIconIndex = "area"
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectArea = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function

'显示选择车次对话框
Public Function SelectBus(Optional MultiSelect As Boolean = False) As String()
On Error GoTo ErrHandle
    Const cnColumnCount = 3
    '验证身证
    AssertActiveUserIsNotThing

    Dim oBase As New BusProject
    Dim aszTmp As Variant

    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    Dim nDataCount As Long


    aszColumns(1) = "车次代码"
    aszColumns(2) = "发车时间"
    aszColumns(3) = "运行线路"
'    aszColumns(4) = "参运公司"

    ''以后要删除计划代码---->
'        Dim szProjectID As String
        Dim srTemp As TSchemeArrangement
        Dim oScheme As New RegularScheme
'        oScheme.Init m_oActiveUser
        Dim oParam As SystemParam
        Set oParam = New SystemParam
        oParam.Init m_oActiveUser
        srTemp = oScheme.GetExecuteBusProject(Date)
'        szProjectID = srTemp.szProjectID
        Set oParam = Nothing
        Set oScheme = Nothing
    ''<------------------------

    Set oBase = New BusProject
    oBase.Init m_oActiveUser
    oBase.Identify
    aszTmp = oBase.GetAllBus()
    nDataCount = ArrayLength(aszTmp)

    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = aszTmp(i, 1)
        aszListDatas(i, 2) = Format(aszTmp(i, 2), "hh:mm")
        aszListDatas(i, 3) = aszTmp(i, 4)
'        aszListDatas(i, 4) = aszTmp(i, 13)
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "车次列表"
    moShell.Title = "选择车次"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 1
    moShell.MultiColumn = "[1][2][3]"

    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "bus"
    moShell.ItemSmallIconIndex = "bus"
    moShell.EnableAddTool = False
    moShell.ShowBox

    SelectBus = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function

''Private Function RefreshBus(ByVal pszProjectID As String) As Long
''    Dim oProject As BusProject
''    Dim nDataCount As Integer, i As Integer
''    Dim liTemp As ListItem
''    Dim aszTmp As Variant
''    Set oProject = New BusProject
''
''    oProject.Init m_oActiveUser
''    oProject.Identify pszProjectID
''    aszTmp = oProject.GetAllBus()
''    nDataCount = ArrayLength(aszTmp)
''    frmSelVehicle.lvList.ListItems.Clear
''    For i = 1 To nDataCount
''        Set liTemp = frmSelVehicle.lvList.ListItems.Add(, GetEncodedKey(RTrim(aszTmp(i, 1))), RTrim(aszTmp(i, 1)))
''        liTemp.SmallIcon = "Bus"
''        liTemp.ListSubItems.Add , , Format(aszTmp(i, 2), "HH:MM:SS")
''        liTemp.ListSubItems.Add , , RTrim(RTrim(aszTmp(i, 4)))
''    Next
''End Function

'显示选择检票口对话框
Public Function SelectCheckGate(Optional MultiSelect As Boolean = True) As String()
On Error GoTo ErrHandle
    Const cnColumnCount = 3
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "检票口代码"
    aszColumns(2) = "检票口名称"
    aszColumns(3) = "所属车站"
    
    Set oBase = New BaseInfo
    oBase.Init m_oActiveUser
    aszTmp = oBase.GetAllCheckGate()
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = aszTmp(i, 1)
        aszListDatas(i, 2) = aszTmp(i, 2)
        aszListDatas(i, 3) = aszTmp(i, 5)
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "选择检票口"
    moShell.Title = "选择检票口"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 1
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "checkgate"
    moShell.ItemSmallIconIndex = "checkgate"
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectCheckGate = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function
''
'选择公司
Public Function SelectCompany(Optional MultiSelect As Boolean = False, Optional SplitCompany As Boolean = False) As String()
On Error GoTo ErrHandle
    Const cnColumnCount = 2
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "公司代码"
    aszColumns(2) = "公司名称"
    
    Set oBase = New BaseInfo
    oBase.Init m_oActiveUser
    aszTmp = oBase.GetCompany(, , SplitCompany)
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = aszTmp(i, 1)
        aszListDatas(i, 2) = aszTmp(i, 2)
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "公司列表"
    moShell.Title = "选择公司"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 2
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "company"
    moShell.ItemSmallIconIndex = "company"
    moShell.ShowBox
    
    SelectCompany = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
    
 End Function
''
'选择车主
Public Function SelectOwner(Optional CompanyID As String = "", Optional MultiSelect As Boolean = True) As String()
'CompanyID 所属公司 缺省为全部

On Error GoTo ErrHandle
    Const cnColumnCount = 2
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "车主代码"
    aszColumns(2) = "车主名称"
    
    If CompanyID = "" Then
        Set oBase = New BaseInfo
        oBase.Init m_oActiveUser
        aszTmp = oBase.GetOwner()
    Else
        Dim oCompany As Company
        Set oCompany = New Company
        oCompany.Init m_oActiveUser
        oCompany.Identify CompanyID
        aszTmp = oCompany.GetAllOwner()
        Set oCompany = Nothing
    End If
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = aszTmp(i, 1)
        aszListDatas(i, 2) = aszTmp(i, 2)
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "车主列表"
    moShell.Title = "选择" & IIf(CompanyID <> "", "公司" & EncodeString(CompanyID) & "的", "") & "车主"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 2
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "owner"
    moShell.ItemSmallIconIndex = "owner"
    moShell.ShowBox
    
    SelectOwner = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function

'显示选择线路对话框
Public Function SelectRoute(Optional MultiSelect As Boolean = False) As String()
On Error GoTo ErrHandle
    Const cnColumnCount = 3
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "线路代码"
    aszColumns(2) = "线路名称"
    aszColumns(3) = "终点站输入码"
    
    Set oBase = New BaseInfo
    oBase.Init m_oActiveUser
    aszTmp = oBase.GetRoute()
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = Trim(aszTmp(i, 1))
        aszListDatas(i, 2) = Trim(aszTmp(i, 2))
        aszListDatas(i, 3) = Trim(aszTmp(i, 8))
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "线路列表"
    moShell.Title = "选择线路"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 3
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "route"
    moShell.ItemSmallIconIndex = "route"
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectRoute = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function


'显示选择线路对话框
Public Function SelectBackRoute(Optional MultiSelect As Boolean = False, Optional IsBackRoute As Boolean = True) As String()
On Error GoTo ErrHandle
    Const cnColumnCount = 3
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As Object
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String
    If IsBackRoute Then
        Set oBase = CreateObject("STSettle.BackBaseInfo")
    Else
        Set oBase = CreateObject("STBase.BaseInfo")
    End If
    aszColumns(1) = "线路代码"
    aszColumns(2) = "线路名称"
    aszColumns(3) = "终点站输入码"
    
    oBase.Init m_oActiveUser
    aszTmp = oBase.GetRoute()
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = Trim(aszTmp(i, 1))
        aszListDatas(i, 2) = Trim(aszTmp(i, 2))
        aszListDatas(i, 3) = Trim(aszTmp(i, 7))
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "线路列表"
    moShell.Title = "选择线路"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 3
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "route"
    moShell.ItemSmallIconIndex = "route"
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectBackRoute = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function



Public Function SelectSection(Optional AreaID As String = "", Optional MultiSelect As Boolean = True) As String()
    '选择路段对话框
    '注意:当AreaID=""时则所有路段
On Error GoTo ErrHandle
    Const cnColumnCount = 2
    '验证身证
    AssertActiveUserIsNotThing
    Dim oArea As New Area
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String
    Dim nDataCount As Long

    
    aszColumns(1) = "路段代码"
    aszColumns(2) = "路段名称"
    If AreaID <> "" Then
        oArea.Init m_oActiveUser
        oArea.Identify AreaID
        aszTmp = oArea.GetAllSection()
    Else
        oBase.Init m_oActiveUser
        aszTmp = oBase.GetSection()
    End If
    
    
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = Trim(aszTmp(i, 1))
        aszListDatas(i, 2) = Trim(aszTmp(i, 2))
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "选择路段"
    moShell.Title = "路段列表"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 1
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "section"
    moShell.ItemSmallIconIndex = "section"
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectSection = moShell.SelectedItems
    Set oBase = Nothing
    Set oArea = Nothing
    
    Exit Function
ErrHandle:
    ShowErrorMsg

End Function


'选择车辆
Public Function SelectVehicle(Optional VehicleID As String, Optional CompanyID As String = "", Optional OwnerID As String = "", Optional BusType As String, Optional BusLicense As String, Optional MultiSelect As Boolean = False) As String()
'CompanyID 所属公司,缺省为全部
'OwnerID 所属车主,缺省为全部

On Error GoTo ErrHandle
    Const cnColumnCount = 7
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "车辆代码"
    aszColumns(2) = "车辆牌照"
    aszColumns(3) = "参运公司"
    aszColumns(4) = "车主"
    aszColumns(5) = "车型"
    aszColumns(6) = "座数"
    aszColumns(7) = "状态"
    
    Set oBase = New BaseInfo
    oBase.Init m_oActiveUser
    aszTmp = oBase.GetVehicle(VehicleID, CompanyID, OwnerID, BusType, BusLicense)
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = aszTmp(i, 1)
        aszListDatas(i, 2) = aszTmp(i, 2)
        aszListDatas(i, 3) = aszTmp(i, 4)
        aszListDatas(i, 4) = aszTmp(i, 5)
        aszListDatas(i, 5) = aszTmp(i, 8)
        aszListDatas(i, 6) = aszTmp(i, 3)
        aszListDatas(i, 7) = IIf(val(aszTmp(i, 6)) = ST_VehicleStop, "(停)", "")
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "参运车辆列表"
    moShell.Title = "选择" & IIf(CompanyID <> "", "公司" & EncodeString(CompanyID) & "的", "") & "车辆"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 2
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "vehicle"
    moShell.ItemSmallIconIndex = "vehicle"
    moShell.ShowBox
    
    SelectVehicle = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
    
End Function

Public Function SelectVehicleEX(Optional MultiSelect As Boolean = False) As String()
    On Error GoTo ErrHandle
    '验证身证
    AssertActiveUserIsNotThing
    Set frmVehicleQuery.m_oActiveUser = m_oActiveUser
    frmVehicleQuery.MultiSelect = MultiSelect
    frmVehicleQuery.Show vbModal
    SelectVehicleEX = frmVehicleQuery.GetSelectVehicle
    Unload frmVehicleQuery
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function


'显示选择车型对话框
Public Function SelectVehicleType(Optional MultiSelect As Boolean = False) As String()
On Error GoTo ErrHandle
    Const cnColumnCount = 2
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "车型代码"
    aszColumns(2) = "车型简称"
    
    Set oBase = New BaseInfo
    oBase.Init m_oActiveUser
    aszTmp = oBase.GetAllVehicleModel()
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = aszTmp(i, 1)
        aszListDatas(i, 2) = aszTmp(i, 2)
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "车型列表"
    moShell.Title = "选择车型"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 2
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "vehicletype"
    moShell.ItemSmallIconIndex = "vehicletype"
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectVehicleType = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function


'显示选择公路等级对话框
Public Function SelectRoadLevel(Optional MultiSelect As Boolean = False) As String()
On Error GoTo ErrHandle
    Const cnColumnCount = 2
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "公路等级代码"
    aszColumns(2) = "公路等级名称"
    
    Set oBase = New BaseInfo
    oBase.Init m_oActiveUser
    aszTmp = oBase.GetAllRoadLevel()
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = aszTmp(i, 1)
        aszListDatas(i, 2) = aszTmp(i, 2)
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "公路等级列表"
    moShell.Title = "选择公路等级"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 2
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "company"
    moShell.ItemSmallIconIndex = "company"
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectRoadLevel = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg

End Function



'因为以下的接口现在的结算里,没用了
'2006-01-12陈峰注释

'******************以下的选择在路单结算中用到*****************
'显示选择路单对话框
Public Function SelectCheckSheet(pszObject As String, pszObjectID As String, Optional pdyStartDate As Date = cszEmptyDateStr, Optional pdyEndDate As Date = cszForeverDateStr, Optional MultiSelect As Boolean = True) As String()
'On Error GoTo ErrHandle
'    Const cnColumnCount = 8
'    '验证身证
'    AssertActiveUserIsNotThing
'
'    Dim oFinance As Object
''    Dim aszTmp As Variant
'    Dim aszColumns(1 To cnColumnCount) As String
'    Dim aszListDatas() As String
'    Dim nDataCount As Long
'    Dim atCheckSheetInfo() As TCheckSheet
'
'    aszColumns(1) = "路单代码"
'    aszColumns(2) = "车次代码"
'    aszColumns(3) = "车次日期"
'    aszColumns(4) = "车牌"
'    aszColumns(5) = "线路"
'    aszColumns(6) = "时间"
'    aszColumns(7) = "参运公司"
'    aszColumns(8) = "拆帐公司"
'
'
'    Set oFinance = CreateObject("STSplit.Finance")
'    oFinance.Init m_oActiveUser
'
'    Select Case pszObjectID
'    Case "Bus"
'        atCheckSheetInfo = oFinance.GetAllCheckSheet(pdyStartDate, pdyEndDate, , pszObjectID, , , , , , , , Availability)
'    Case "Vehicle"
'        atCheckSheetInfo = oFinance.GetAllCheckSheet(pdyStartDate, pdyEndDate, , , , pszObjectID, , , , , , Availability)
'    Case "Owner"
'        atCheckSheetInfo = oFinance.GetAllCheckSheet(pdyStartDate, pdyEndDate, , , , , , , pszObjectID, , , Availability)
'    Case "Company"
'        atCheckSheetInfo = oFinance.GetAllCheckSheet(pdyStartDate, pdyEndDate, , , , , pszObjectID, , , , , Availability)
'    Case "SplitCompany"
'        atCheckSheetInfo = oFinance.GetAllCheckSheet(pdyStartDate, pdyEndDate, , , , , , pszObjectID, , , , Availability)
'    End Select
'
'    nDataCount = ArrayLength(atCheckSheetInfo)
'    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
'    Dim i As Long, j As Integer
'    For i = 1 To nDataCount
'        aszListDatas(i, 1) = atCheckSheetInfo(i).szCheckSheetID
'        aszListDatas(i, 2) = atCheckSheetInfo(i).szBusID
'        aszListDatas(i, 3) = atCheckSheetInfo(i).dtBusDate
'        aszListDatas(i, 4) = atCheckSheetInfo(i).szSellStationName
'        aszListDatas(i, 5) = atCheckSheetInfo(i).szLicenseTagNo
'        aszListDatas(i, 6) = atCheckSheetInfo(i).szRouteName
'        aszListDatas(i, 7) = atCheckSheetInfo(i).dtStartTime
'        aszListDatas(i, 8) = atCheckSheetInfo(i).szCompanyName
'        aszListDatas(i, 9) = atCheckSheetInfo(i).szSplitCompanyName
'
'    Next i
'    moShell.ListColumns = aszColumns
'    moShell.ListData = aszListDatas
'    moShell.DialogCaption = "选择路单"
'    moShell.Title = "路单列表"
'    moShell.MultiSelect = MultiSelect
'    moShell.MatchIndex = 1
'    moShell.MultiColumn = "[1][2][3][5][6][7][8][9][4]"
'
'    Set moShell.ItemIcons = frmResource.imlBigIcons
'    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
'    moShell.ItemIconIndex = "checksheet"
'    moShell.ItemSmallIconIndex = "checksheet"
'    moShell.EnableAddTool = False
'    moShell.ShowBox
'
'    SelectCheckSheet = moShell.SelectedItems
'    Set oFinance = Nothing
'    Exit Function
'ErrHandle:
'    ShowErrorMsg

End Function

'显示选择签发单对话框
Public Function SelectLuggageSheet(pszObject As String, pszObjectID As String, Optional pdyStartDate As Date = cszEmptyDateStr, Optional pdyEndDate As Date = cszForeverDateStr, Optional MultiSelect As Boolean = True) As String()
'On Error GoTo ErrHandle
'    Const cnColumnCount = 8
'    '验证身证
'    AssertActiveUserIsNotThing
'
'    Dim oFinance As Object
''    Dim aszTmp As Variant
'    Dim aszColumns(1 To cnColumnCount) As String
'    Dim aszListDatas() As String
'    Dim nDataCount As Long
'    Dim atLuggageSheet() As TLuggageSheet
'
'    aszColumns(1) = "签发单代码"
'    aszColumns(2) = "车次代码"
'    aszColumns(3) = "车次日期"
'    aszColumns(4) = "车牌"
'    aszColumns(5) = "线路"
'    aszColumns(6) = "时间"
'    aszColumns(7) = "参运公司"
'    aszColumns(8) = "拆帐公司"
'
'
'    Set oFinance = CreateObject("STSplit.Finance")
'    oFinance.Init m_oActiveUser
'
'    Select Case pszObject
'    Case "Bus"
'        atLuggageSheet = oFinance.GetAllLuggage(pdyStartDate, pdyEndDate, , pszObjectID, , , , , , , , Availability)
'    Case "Vehicle"
'        atLuggageSheet = oFinance.GetAllLuggage(pdyStartDate, pdyEndDate, , , , pszObjectID, , , , , , Availability)
'    Case "Owner"
'        atLuggageSheet = oFinance.GetAllLuggage(pdyStartDate, pdyEndDate, , , , , , , pszObjectID, , , Availability)
'    Case "Company"
'        atLuggageSheet = oFinance.GetAllLuggage(pdyStartDate, pdyEndDate, , , , , pszObjectID, , , , , Availability)
'    Case "SplitCompany"
'        atLuggageSheet = oFinance.GetAllLuggage(pdyStartDate, pdyEndDate, , , , , , pszObjectID, , , , Availability)
'    End Select
'
'    nDataCount = ArrayLength(atLuggageSheet)
'    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
'    Dim i As Long, j As Integer
'    For i = 1 To nDataCount
'        aszListDatas(i, 1) = atLuggageSheet(i).szLuggageSheetID
'        aszListDatas(i, 2) = atLuggageSheet(i).szBusID
'        aszListDatas(i, 3) = atLuggageSheet(i).dtBusDate
'        aszListDatas(i, 4) = atLuggageSheet(i).szLicenseTagNo
'        aszListDatas(i, 5) = atLuggageSheet(i).szRouteName
'        aszListDatas(i, 6) = atLuggageSheet(i).dtStartTime
'        aszListDatas(i, 7) = atLuggageSheet(i).szCompanyName
'        aszListDatas(i, 8) = atLuggageSheet(i).szSplitCompanyName
'
'    Next i
'    moShell.ListColumns = aszColumns
'    moShell.ListData = aszListDatas
'    moShell.DialogCaption = "选择签发单"
'    moShell.Title = "签发单列表"
'    moShell.MultiSelect = MultiSelect
'    moShell.MatchIndex = 1
'    moShell.MultiColumn = "[1][2][3][4][5][6][7][8]"
'
'    Set moShell.ItemIcons = frmResource.imlBigIcons
'    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
'    moShell.ItemIconIndex = "checksheet"
'    moShell.ItemSmallIconIndex = "checksheet"
'    moShell.EnableAddTool = False
'    moShell.ShowBox
'
'    SelectLuggageSheet = moShell.SelectedItems
'    Set oFinance = Nothing
'    Exit Function
'ErrHandle:
'    ShowErrorMsg

End Function

'显示选择公式对话框
Public Function SelectFormular(Optional MultiSelect As Boolean = False) As String()
'On Error GoTo ErrHandle
'    Const cnColumnCount = 3
'    '验证身证
'    AssertActiveUserIsNotThing
'
'    Dim oSplit As Object
'    Dim aszTmp As Variant
'
'    Dim aszColumns(1 To cnColumnCount) As String
'    Dim aszListDatas() As String
'
'
'    aszColumns(1) = "公式代码"
'    aszColumns(2) = "公式名称"
'    aszColumns(3) = "公式内容"
'    Set oSplit = CreateObject("STSplit.Split")
'    oSplit.Init m_oActiveUser
'    aszTmp = oSplit.GetFormular()
'
'    Dim nDataCount As Long
'    nDataCount = ArrayLength(aszTmp)
'    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
'    Dim i As Long, j As Integer
'    For i = 1 To nDataCount
'        aszListDatas(i, 1) = aszTmp(i, 1)
'        aszListDatas(i, 2) = aszTmp(i, 2)
'        aszListDatas(i, 3) = aszTmp(i, 3)
'    Next i
'    moShell.ListColumns = aszColumns
'    moShell.ListData = aszListDatas
'    moShell.DialogCaption = "选择公式"
'    moShell.Title = "公式列表"
'    moShell.MultiSelect = MultiSelect
'    moShell.MatchIndex = 1
'    moShell.MultiColumn = "[1][2][3]"
'
'    Set moShell.ItemIcons = frmResource.imlBigIcons
'    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
'    moShell.ItemIconIndex = "checksheet"
'    moShell.ItemSmallIconIndex = "checksheet"
'    moShell.EnableAddTool = False
'    moShell.ShowBox
'
'    SelectFormular = moShell.SelectedItems
'    Set oSplit = Nothing
'    Exit Function
'ErrHandle:
'    ShowErrorMsg

End Function


Public Function SelectCompensateBus(ByVal pdtBusDate As Date, Optional MultiSelect As Boolean = False) As String()
'On Error GoTo ErrHandle
'    Const cnColumnCount = 4
'    '验证身证
'    AssertActiveUserIsNotThing
'
'    Dim oSplit As Object
'    Dim atTmp() As TInputBusInfo
'
'    Dim aszColumns(1 To cnColumnCount) As String
'    Dim aszListDatas() As String
'
'
'    aszColumns(1) = "车次代码"
'    aszColumns(2) = "车次序号"
'    aszColumns(3) = "结算金额"
'    aszColumns(4) = "发车日期"
'
'    Set oSplit = CreateObject("STSplit.Split")
'    oSplit.Init m_oActiveUser
'    atTmp = oSplit.GetInputBusInfo(pdtBusDate)
'
'    Dim nDataCount As Long
'    nDataCount = ArrayLength(atTmp)
'    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
'    Dim i As Long, j As Integer
'    For i = 1 To nDataCount
'        aszListDatas(i, 1) = atTmp(i).szBusID
'        aszListDatas(i, 2) = atTmp(i).nSerialNumber
'        aszListDatas(i, 3) = atTmp(i).dbSettleMoney
'        aszListDatas(i, 4) = atTmp(i).dtBusStartDate
'
'    Next i
'    moShell.ListColumns = aszColumns
'    moShell.ListData = aszListDatas
'    moShell.DialogCaption = "补差车次选择"
'    moShell.Title = "补差车次列表"
'    moShell.MultiSelect = MultiSelect
'    moShell.MatchIndex = 1
'    moShell.MultiColumn = "[1][2][3][4]"
'
'    Set moShell.ItemIcons = frmResource.imlBigIcons
'    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
'    moShell.ItemIconIndex = "checksheet"
'    moShell.ItemSmallIconIndex = "checksheet"
'    moShell.EnableAddTool = False
'    moShell.ShowBox
'
'    SelectCompensateBus = moShell.SelectedItems
'    Set oSplit = Nothing
'    Exit Function
'ErrHandle:
'    ShowErrorMsg
    
End Function

'显示选择需要结算的路单对话框
Public Function SelectNeedSplitCS(pdtStartDate As Date, pdtEndDate As Date, Optional pnObjectType As Integer = 555, Optional pszObjectID As String = "", Optional MultiSelect As Boolean = True) As String()
'    On Error GoTo ErrHandle
'    Const cnColumnCount = 8
'    '验证身证
'    AssertActiveUserIsNotThing
'
'    Dim oFinance As Object
''    Dim aszTmp As Variant
'    Dim aszColumns(1 To cnColumnCount) As String
'    Dim aszListDatas() As String
'    Dim nDataCount As Long
'    Dim atCheckSheetInfo() As TCheckSheet
'
'    aszColumns(1) = "路单代码"
'    aszColumns(2) = "车次代码"
'    aszColumns(3) = "车次日期"
'    aszColumns(4) = "车牌"
'    aszColumns(5) = "线路"
'    aszColumns(6) = "时间"
'    aszColumns(7) = "参运公司"
'    aszColumns(8) = "拆帐公司"
'
'
'    Set oFinance = CreateObject("STSplit.Finance")
'    oFinance.Init m_oActiveUser
'
'    atCheckSheetInfo = oFinance.GetAppointedCheckSheet(pdtStartDate, pdtEndDate, , pnObjectType, pszObjectID, , Availability)
'
'    nDataCount = ArrayLength(atCheckSheetInfo)
'    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
'    Dim i As Long, j As Integer
'    For i = 1 To nDataCount
'        aszListDatas(i, 1) = atCheckSheetInfo(i).szCheckSheetID
'        aszListDatas(i, 2) = atCheckSheetInfo(i).szBusID
'        aszListDatas(i, 3) = atCheckSheetInfo(i).dtBusDate
'        aszListDatas(i, 4) = atCheckSheetInfo(i).szLicenseTagNo
'        aszListDatas(i, 5) = atCheckSheetInfo(i).szRouteName
'        aszListDatas(i, 6) = atCheckSheetInfo(i).dtStartTime
'        aszListDatas(i, 7) = atCheckSheetInfo(i).szCompanyName
'        aszListDatas(i, 8) = atCheckSheetInfo(i).szSplitCompanyName
'
'    Next i
'    moShell.ListColumns = aszColumns
'    moShell.ListData = aszListDatas
'    moShell.DialogCaption = "选择路单"
'    moShell.Title = "路单列表"
'    moShell.MultiSelect = MultiSelect
'    moShell.MatchIndex = 1
'    moShell.MultiColumn = "[1][2][3][4][5][6][7][8]"
'
'    Set moShell.ItemIcons = frmResource.imlBigIcons
'    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
'    moShell.ItemIconIndex = "checksheet"
'    moShell.ItemSmallIconIndex = "checksheet"
'    moShell.EnableAddTool = False
'    moShell.ShowBox
'
'    SelectNeedSplitCS = moShell.SelectedItems
'    Set oFinance = Nothing
'    Exit Function
'ErrHandle:
'    ShowErrorMsg
'
End Function

'显示选择协议
Public Function SelectProtocol(Optional MultiSelect As Boolean = False) As String()
'On Error GoTo ErrHandle
'    Const cnColumnCount = 3
'    '验证身证
'    AssertActiveUserIsNotThing
'
'    Dim oFinance As Object
'    Dim atTemp() As TProtocol
'    Dim aszColumns(1 To cnColumnCount) As String
'    Dim aszListDatas() As String
'
'
'    aszColumns(1) = "协议代码"
'    aszColumns(2) = "协议名称"
'    aszColumns(3) = "协议种类"
'    Set oFinance = CreateObject("STSplit.Finance")
'    oFinance.Init m_oActiveUser
'
'    atTemp = oFinance.GetAllProtocol()
'
'    Dim nDataCount As Long
'    nDataCount = ArrayLength(atTemp)
'    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
'    Dim i As Long, j As Integer
'    For i = 1 To nDataCount
'        aszListDatas(i, 1) = atTemp(i).szProtocolID
'        aszListDatas(i, 2) = atTemp(i).szProtocolName
'        aszListDatas(i, 3) = IIf(atTemp(i).nProtocolType = CommonBus, "长运车辆", "社会车辆")
'    Next i
'    moShell.ListColumns = aszColumns
'    moShell.ListData = aszListDatas
'    moShell.DialogCaption = "选择公式"
'    moShell.Title = "公式列表"
'    moShell.MultiSelect = MultiSelect
'    moShell.MatchIndex = 1
'    moShell.MultiColumn = "[1][2][3]"
'
'    Set moShell.ItemIcons = frmResource.imlBigIcons
'    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
'    moShell.ItemIconIndex = "checksheet"
'    moShell.ItemSmallIconIndex = "checksheet"
'    moShell.EnableAddTool = False
'    moShell.ShowBox
'
'    SelectProtocol = moShell.SelectedItems
'    Set oFinance = Nothing
'    Exit Function
'ErrHandle:
'    ShowErrorMsg

End Function


'显示选择路单结算单
Public Function SelectSettleCS(Optional pdtStartDate As Date = cszEmptyDateStr, Optional pdtEndDate As Date = cszForeverDateStr, Optional MultiSelect As Boolean = False) As String()
'    On Error GoTo ErrHandle
'    Const cnColumnCount = 8
'    '验证身证
'    AssertActiveUserIsNotThing
'
'    Dim oFinance As Object
'    Dim aszColumns(1 To cnColumnCount) As String
'    Dim aszListDatas() As String
'    Dim nDataCount As Long
'    Dim atTemp() As TSellet
'    Dim i As Long, j As Integer
'
'
'    aszColumns(1) = "结算单号"
'    aszColumns(2) = "结算日期"
'    aszColumns(3) = "起始日期"
'    aszColumns(4) = "结束日期"
'    aszColumns(5) = "结算单位"
'    aszColumns(6) = "领款人"
'    aszColumns(7) = "结算人"
'    aszColumns(8) = "复核人"
'
'    Set oFinance = CreateObject("STSplit.Finance")
'    oFinance.Init m_oActiveUser
'
'    atTemp = oFinance.GetAppointedSettleSheet(pdtStartDate, pdtEndDate)
'
'    nDataCount = ArrayLength(atTemp)
'    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
'
'    For i = 1 To nDataCount
'        aszListDatas(i, 1) = atTemp(i).szSettleSheetID
'        aszListDatas(i, 2) = atTemp(i).dtSettlementDate
'        aszListDatas(i, 3) = atTemp(i).dtStartDate
'        aszListDatas(i, 4) = atTemp(i).dtEndDate
'        aszListDatas(i, 5) = atTemp(i).szSettlementUnit
'        aszListDatas(i, 6) = atTemp(i).szMoneyReceiver
'        aszListDatas(i, 7) = atTemp(i).szSettlementExecutive
'        aszListDatas(i, 8) = atTemp(i).szSettlementChecker
'
'    Next i
'    moShell.ListColumns = aszColumns
'    moShell.ListData = aszListDatas
'    moShell.DialogCaption = "选择路单结算单"
'    moShell.Title = "路单结算单列表"
'    moShell.MultiSelect = MultiSelect
'    moShell.MatchIndex = 1
'    moShell.MultiColumn = "[1][2][3][4][5][6][7][8]"
'
'    Set moShell.ItemIcons = frmResource.imlBigIcons
'    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
'    moShell.ItemIconIndex = "checksheet"
'    moShell.ItemSmallIconIndex = "checksheet"
'    moShell.EnableAddTool = False
'    moShell.ShowBox
'
'    SelectSettleCS = moShell.SelectedItems
'    Set oFinance = Nothing
'    Exit Function
'ErrHandle:
'    ShowErrorMsg

End Function


'显示选择某车次的站点
Public Function SelectBusStation(pszBusID As String, pdyBusDate As Date, Optional MultiSelect As Boolean = False) As String()
'On Error GoTo ErrHandle
'    Const cnColumnCount = 5
'    '验证身证
'    AssertActiveUserIsNotThing
'
'    Dim oFinance As Object
'    Dim aszTemp() As String
'    Dim aszColumns(1 To cnColumnCount) As String
'    Dim aszListDatas() As String
'
'
'    aszColumns(1) = "站点名称"
'    aszColumns(2) = "站点代码"
'    aszColumns(3) = "站点输入码"
'    aszColumns(4) = "上车站代码"
'    aszColumns(5) = "上车站名称"
'    Set oFinance = CreateObject("STSplit.Finance")
'    oFinance.Init m_oActiveUser
'
'    aszTemp = oFinance.GetStationLst(pszBusID, pdyBusDate)
'
'    Dim nDataCount As Long
'    nDataCount = ArrayLength(aszTemp)
'    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
'    Dim i As Long, j As Integer
'    For i = 1 To nDataCount
'        aszListDatas(i, 1) = aszTemp(i, 1)
'        aszListDatas(i, 2) = aszTemp(i, 2)
'        aszListDatas(i, 3) = aszTemp(i, 3)
'        aszListDatas(i, 4) = aszTemp(i, 4)
'        aszListDatas(i, 5) = aszTemp(i, 5)
'    Next i
'    moShell.ListColumns = aszColumns
'    moShell.ListData = aszListDatas
'    moShell.DialogCaption = "选择公式"
'    moShell.Title = "公式列表"
'    moShell.MultiSelect = MultiSelect
'    moShell.MatchIndex = 2
'    moShell.MultiColumn = "[1][2][3][4][5]"
'
'    Set moShell.ItemIcons = frmResource.imlBigIcons
'    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
'    moShell.ItemIconIndex = "station"
'    moShell.ItemSmallIconIndex = "station"
'    moShell.EnableAddTool = False
'    moShell.ShowBox
'
'    SelectBusStation = moShell.SelectedItems
'    Set oFinance = Nothing
'    Exit Function
'ErrHandle:
'    ShowErrorMsg

End Function


'显示选择环境车次对话框
Public Function SelectREBus(ByRef pdyRunDate As Date, Optional pbAllowChangeDate As Boolean = False, Optional MultiSelect As Boolean = True) As String()
'pdyRunDate 中返回选择的日期
On Error GoTo ErrHandle
    Const cnColumnCount = 3
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oReScheme As STReSch.REScheme
    Dim aszTemp() As String
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "车次代码"
    aszColumns(2) = "发车时间"
    aszColumns(3) = "运行线路"
    Set oReScheme = CreateObject("STReSch.REScheme")
    oReScheme.Init m_oActiveUser
    
    aszListDatas = oReScheme.GetBus(pdyRunDate)
    
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "选择环境车次"
    moShell.Title = "环境车次列表"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 1
    moShell.MultiColumn = "[1][2][3]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "rebus"
    moShell.ItemSmallIconIndex = "rebus"
    
    moShell.AllowDateChange = pbAllowChangeDate
    moShell.SelectDate = pdyRunDate
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectREBus = moShell.SelectedItems
    pdyRunDate = moShell.SelectDate
    Set oReScheme = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function
'显示选择车站对话框
Public Function SelectStation(Optional pszAreaID As String, Optional MultiSelect As Boolean = False) As String()
'CompanyID 所属公司 缺省为全部

On Error GoTo ErrHandle
    Const cnColumnCount = 3
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oBase As New BaseInfo
    Dim aszTmp As Variant
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "站点代码"
    aszColumns(2) = "站点名称"
    aszColumns(3) = "站点输入码"
    
    Set oBase = New BaseInfo
    oBase.Init m_oActiveUser
    aszTmp = oBase.GetStation(pszAreaID)
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = Trim(aszTmp(i, 1))
        aszListDatas(i, 2) = Trim(aszTmp(i, 2))
        aszListDatas(i, 3) = Trim(aszTmp(i, 3))
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "站点列表"
    moShell.Title = "选择" & IIf(pszAreaID <> "", "地区" & EncodeString(pszAreaID) & "的", "") & "站点"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 3
    moShell.MultiColumn = "[1][2]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "station"
    moShell.ItemSmallIconIndex = "station"
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectStation = moShell.SelectedItems
    Set oBase = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function


Private Sub moShell_AddNewItem(ByVal FormCaption As String)
    Dim aszItem() As String
    Select Case FormCaption
'        Case "地区列表"
'            Set frmArea.g_oActiveUser = m_oActiveUser
'            frmArea.Status = EFormStatus.EFS_AddNew
'            frmArea.Show vbModal
'            If IsArray(frmArea.maszReturnItem) Then
'                ReDim aszItem(1 To 2)
'                aszItem(1) = frmArea.maszReturnItem(1)
'                aszItem(2) = frmArea.maszReturnItem(2)
'                moShell.InsertItemLine aszItem, True
'            End If
'            Unload frmArea
        Case "公司列表"
            Set frmCompany.g_oActiveUser = m_oActiveUser
            frmCompany.Status = eFormStatus.EFS_AddNew
            frmCompany.Show vbModal
            If IsArray(frmCompany.maszReturnItem) Then
                ReDim aszItem(1 To 2)
                aszItem(1) = frmCompany.maszReturnItem(1)
                aszItem(2) = frmCompany.maszReturnItem(2)
                moShell.InsertItemLine aszItem, True
            End If
            Unload frmCompany
'        Case "公路等级列表"
'            Set frmRoadLevel.g_oActiveUser = m_oActiveUser
'            frmRoadLevel.Status = EFormStatus.EFS_AddNew
'            frmRoadLevel.Show vbModal
'            If IsArray(frmRoadLevel.maszReturnItem) Then
'                ReDim aszItem(1 To 2)
'                aszItem(1) = frmRoadLevel.maszReturnItem(1)
'                aszItem(2) = frmRoadLevel.maszReturnItem(2)
'                moShell.InsertItemLine aszItem, True
'            End If
'            Unload frmRoadLevel
'        Case "车型列表"
'            Set frmVehicleType.g_oActiveUser = m_oActiveUser
'            frmVehicleType.Status = EFormStatus.EFS_AddNew
'            frmVehicleType.Show vbModal
'            If IsArray(frmVehicleType.maszReturnItem) Then
'                ReDim aszItem(1 To 2)
'                aszItem(1) = frmVehicleType.maszReturnItem(1)
'                aszItem(2) = frmVehicleType.maszReturnItem(2)
'                moShell.InsertItemLine aszItem, True
'            End If
'            Unload frmVehicleType
         Case "参运车辆列表"
            Set frmVehicle.g_oActiveUser = m_oActiveUser
            frmVehicle.Status = eFormStatus.EFS_AddNew
            frmVehicle.Show vbModal
            If IsArray(frmVehicle.maszReturnItem) Then
                ReDim aszItem(1 To 2)
                aszItem(1) = frmVehicle.maszReturnItem(1)
                aszItem(2) = frmVehicle.maszReturnItem(2)
                moShell.InsertItemLine aszItem, True
            End If
            Unload frmVehicle
         Case "车主列表"
            Set frmOwner.g_oActiveUser = m_oActiveUser
            frmOwner.Status = eFormStatus.EFS_AddNew
            frmOwner.Show vbModal
            If IsArray(frmOwner.maszReturnItem) Then
                ReDim aszItem(1 To 2)
                aszItem(1) = frmOwner.maszReturnItem(1)
                aszItem(2) = frmOwner.maszReturnItem(2)
                moShell.InsertItemLine aszItem, True
            End If
            Unload frmOwner
    End Select
End Sub

Private Sub moShell_DateChange(ByVal FormCaption As String)
    Dim oReScheme As STReSch.REScheme
    Dim aszTemp() As String
    Select Case FormCaption
    Case "选择环境车次"
        Set oReScheme = CreateObject("STReSch.REScheme")
        oReScheme.Init m_oActiveUser
        moShell.ListData = oReScheme.GetBus(moShell.SelectDate)
    End Select
        
End Sub

'显示车主信息
Public Sub ShowOwnerInfo(ByVal OwnerID As String)
On Error GoTo ErrHandle
    Dim ofrmOwner As New frmOwner
    Set ofrmOwner.g_oActiveUser = m_oActiveUser
    ofrmOwner.mszOwnerID = OwnerID
    ofrmOwner.Status = eFormStatus.EFS_Show
    ofrmOwner.Show vbModal
    Set ofrmOwner = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'显示车辆信息
Public Sub ShowVehicleInfo(ByVal VehicleID As String)
On Error GoTo ErrHandle
    Dim ofrmVehicle As New frmVehicle
    Set ofrmVehicle.g_oActiveUser = m_oActiveUser
    ofrmVehicle.mszVehicleId = VehicleID
    ofrmVehicle.Status = eFormStatus.EFS_Show
    ofrmVehicle.Show vbModal
    Set ofrmVehicle = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'显示车辆信息
Public Sub ShowCompanyInfo(ByVal CompanyID As String)
On Error GoTo ErrHandle
    Dim ofrmCompany As New frmCompany
    Set ofrmCompany.g_oActiveUser = m_oActiveUser
    ofrmCompany.mszCompanyID = CompanyID
    ofrmCompany.Status = eFormStatus.EFS_Show
    ofrmCompany.Show vbModal
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'显示车次信息
Public Sub ShowBusInfo(BusDate As Date, BusID As String)
On Error GoTo ErrHandle
    Dim oFrmBusInfo As New frmBusInfo
    oFrmBusInfo.SelfUser = m_oActiveUser
    oFrmBusInfo.BusID = BusID
    oFrmBusInfo.BusDate = BusDate
    oFrmBusInfo.Show vbModal
    Set oFrmBusInfo = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'显示车次检票信息
Public Sub ShowCheckInfo(BusDate As Date, BusID As String, Optional BusSerialNo As Integer = 0)
On Error GoTo ErrHandle
    Dim oFrmCheckInfo As New frmCheckBusInfo
    Set oFrmCheckInfo.g_oActiveUser = m_oActiveUser
    oFrmCheckInfo.mszBusID = BusID
    oFrmCheckInfo.mnBusSerialNo = BusSerialNo
    oFrmCheckInfo.mdtBusDate = BusDate
    Load oFrmCheckInfo
    oFrmCheckInfo.Show vbModal

    Exit Sub
ErrHandle:
'    ShowErrorMsg
End Sub
'显示车票信息
Public Sub ShowTicketInfo(pszTicketId As String)
On Error GoTo ErrHandle
   
    Dim ofrmTicketInfo As New frmTicketInfo
    Set ofrmTicketInfo.g_oActiveUser = m_oActiveUser
    ofrmTicketInfo.TicketID = pszTicketId
    ofrmTicketInfo.Show vbModal

    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'显示路单
Public Sub ShowCheckSheet(CheckSheetID As String)
On Error GoTo ErrHandle
    Dim oChkTicket As New CheckTicket
    oChkTicket.Init m_oActiveUser
    Dim ofrmSheet As New frmCheckSheet
    Set ofrmSheet.moChkTicket = oChkTicket
    Set ofrmSheet.g_oActiveUser = m_oActiveUser
    ofrmSheet.mszSheetID = CheckSheetID
    ofrmSheet.Show vbModal
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

' 根据车次路单调用路单显示窗体
Public Sub ShowCheckSheetEx(BusDate As Date, BusID As String, BusSerialNo As Integer)
'On Error GoTo ErrHandle
'    Dim oChkTicket As New CheckTicket
'    oChkTicket.Init m_oActiveUser
'    Dim ofrmSheet As New frmCheckSheet
'    Set ofrmSheet.moChkTicket = oChkTicket
'    Set ofrmSheet.g_oActiveUser = m_oActiveUser
'    ofrmSheet.SetSheetID BusDate, BusID, BusSerialNo
'    ofrmSheet.Show vbModal
'
'    Exit Sub
'ErrHandle:
'    ShowErrorMsg
End Sub

'显示用户信息(实际更改密码)对话框
Public Sub ShowUserInfo()
On Error GoTo ErrHandle
    Set frmUserPwd.m_oActiveUser = m_oActiveUser
    frmUserPwd.Show vbModal
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'显示行包结算单信息
Public Sub ShowLugFinSheet(SheetID As String, Status As String)
    On Error GoTo ErrHandle
    If Status = mStatusNo Or Status = mStatusReal Then
        frmFinSheetAttrib.FormStatus = NotStatus
    Else
        frmFinSheetAttrib.FormStatus = ShowStatus
    End If
    Set frmFinSheetAttrib.g_oActiveUser = m_oActiveUser
    frmFinSheetAttrib.mSheetID = Trim(SheetID)
    frmFinSheetAttrib.Show vbModal
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
    
End Sub




'选择地区
Public Function SelectSellStation(Optional SellStationID As String = "", Optional UnitID As String = "", Optional MultiSelect As Boolean = False) As String()
On Error GoTo ErrHandle
    Const cnColumnCount = 3
    '验证身证
    AssertActiveUserIsNotThing
    
    Dim oSysMan As New SystemMan
    Dim aszTmp() As TDepartmentInfo
        
    Dim aszColumns(1 To cnColumnCount) As String
    Dim aszListDatas() As String

    
    aszColumns(1) = "上车站代码"
    aszColumns(2) = "上车站名称"
    aszColumns(3) = "单位"
    
    
    oSysMan.Init m_oActiveUser
    aszTmp = oSysMan.GetAllSellStation()
    
    Dim nDataCount As Long
    nDataCount = ArrayLength(aszTmp)
    If nDataCount > 0 Then ReDim aszListDatas(1 To nDataCount, 1 To cnColumnCount)
    Dim i As Long, j As Integer
    For i = 1 To nDataCount
        aszListDatas(i, 1) = aszTmp(i).szSellStationID
        aszListDatas(i, 2) = aszTmp(i).szSellStationName
        aszListDatas(i, 3) = aszTmp(i).szUnitName
    Next i
    moShell.ListColumns = aszColumns
    moShell.ListData = aszListDatas
    moShell.DialogCaption = "上车站列表"
    moShell.Title = "选择上车站"
    moShell.MultiSelect = MultiSelect
    moShell.MatchIndex = 1
    moShell.MultiColumn = "[1][2][3]"
    
    Set moShell.ItemIcons = frmResource.imlBigIcons
    Set moShell.ItemSmallIcons = frmResource.imlSmallIcons
    moShell.ItemIconIndex = "station"
    moShell.ItemSmallIconIndex = "station"
    moShell.EnableAddTool = False
    moShell.ShowBox
    
    SelectSellStation = moShell.SelectedItems
    Set oSysMan = Nothing
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function



Public Sub ShowEnvScrollBusList(BusDate As Date, BusID As String)
    On Error GoTo ErrHandle
    '验证身证
    AssertActiveUserIsNotThing
    Set frmEnvScrollBusLst.m_oActiveUser = m_oActiveUser
    frmEnvScrollBusLst.m_dtBusDate = BusDate
    frmEnvScrollBusLst.m_szBusID = BusID
    
    frmEnvScrollBusLst.Show vbModal
    
    
    Unload frmEnvScrollBusLst
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub



