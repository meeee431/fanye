Attribute VB_Name = "mdFunsVisual"
Option Explicit

' *******************************************************************
' *  Source File Name: mdFunsVisual                                 *
' *  Brief Description: 提供用于界面的一系列功能函数                *
' *******************************************************************
Public Enum EStatusBarPanelArea '状态条区域
    areaUserInfo = 1
    areaProgramInfo = 2
    areaProcessInfo = 3
End Enum
Public Enum EIndexSystemButton  '系统按钮索引
    LogMan = 1
    ExportToFile = 3
    ExportAndOpen = 4
    PrintData = 6
    PrintPreview = 7
    ExitSystem = 9
End Enum
'Public Enum ESearchAreaIndex     '搜索位置
'    ESI_TextArea = 1             '主文本匹配
'    ESI_SubItemArea = -1            'SUBITEM区
'    ESI_WholeArea = -2              '全部区域
'End Enum

'
'''窗体的当前状态
'Public Enum eFormStatus
'    AddStatus = 0
'    ModifyStatus = 1
'    ShowStatus = 2
'    NotStatus = 3
'End Enum
''

'以下全局常量定义
'对齐字符串
Public Const cszLeftAlignString = "<"
Public Const cszRightAlignString = ">"
Public Const cszMiddleAlignString = "^"

Public Const cszPrefixFlag = "<<"
Public Const cszSuffixFlag = ">>"
Public Const cszUserName = "UserID"
Public Const cszUserPassword = "Password"



'以下全局变量定义

'以下API定义


Public Const HWND_BOTTOM = 1 '将窗口置于窗口列表底部
Public Const HWND_TOP = 0 '将窗口置于Z序列的顶部；Z序列代表在分级结构中，窗口针对一个给定级别的窗口显示的顺序
Public Const HWND_TOPMOST = -1 '将窗口置于列表顶部，并位于任何最顶部窗口的前面
Public Const HWND_NOTOPMOST = -2 '将窗口置于列表顶部，并位于任何最顶部窗口的后面

Public Const SWP_HIDEWINDOW = &H80 '隐藏窗口
Public Const SWP_NOACTIVATE = &H10 '不激活窗口
Public Const SWP_NOMOVE = &H2 '保持当前位置(x和y设定将被忽略)
Public Const SWP_NOREDRAW = &H8 '窗口不自动重画
Public Const SWP_NOSIZE = &H1 '保持当前大小(cx和cy会被忽略)
Public Const SWP_NOZORDER = &H4 '保持窗口在列表的当前位置(hWndInsertAfter将被忽略)
Public Const SWP_SHOWWINDOW = &H40 '显示窗口
Public Const SWP_FRAMECHANGED = &H20 '强迫一条WM_NCCALCSIZE消息进入窗口，即使窗口的大小没有改变
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED '围绕窗口画一个框

Public Const LB_SELECTSTRING = &H18C
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_FINDSTRING = &H14C


Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'判断是否有打印机常数、声明
Const PRINTER_ENUM_CONNECTIONS = &H4
Const PRINTER_ENUM_LOCAL = &H2
Public Declare Function EnumPrinters Lib "winspool.drv" Alias _
   "EnumPrintersA" (ByVal Flags As Long, ByVal name As String, _
   ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, _
   pcbNeeded As Long, pcReturned As Long) As Long
'关机声明
Public Const EWX_SHUTDOWN = 1
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Type POINTAPI
    X As Long
    y As Long
End Type
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPOINT As POINTAPI) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


' *******************************************************************
' *   Member Name: RealLocate                                       *
' *   Brief Description: 根据匹配串实时定位控件列表位置             *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************

Public Sub RealLocate(pszMatchSZ As String, poLocateObject As Object, Optional pvOption As Variant)
'参数注释
'*************************************
'pszMatchSZ(匹配串)
'poLocateObject(定位控件)
'pnMatchColumn(匹配字段)
'************************************
    Dim szObjectType As String
    Dim nTmp As Integer
    szObjectType = UCase(TypeName(poLocateObject))
    Select Case UCase(szObjectType)
        Case "COMBOBOX"
            If Not IsArray(pvOption) Then
                RealLocateCBO pszMatchSZ, poLocateObject
            Else
                nTmp = ArrayLength(pvOption)
                If nTmp = 1 Then
                    RealLocateCBO pszMatchSZ, poLocateObject, CBool(pvOption(1))
                Else
                    RealLocateCBO pszMatchSZ, poLocateObject, CBool(pvOption(1)), CBool(pvOption(2))
                End If
            End If
        Case "LISTBOX"
            RealLocateLST pszMatchSZ, poLocateObject
        Case "LISTVIEW"
            If Not IsArray(pvOption) Then
                RealLocateLVW pszMatchSZ, poLocateObject
            Else
                nTmp = ArrayLength(pvOption)
                If nTmp = 1 Then
                    RealLocateLVW pszMatchSZ, poLocateObject, CInt(pvOption(1))
                Else
                    RealLocateLVW pszMatchSZ, poLocateObject, CInt(pvOption(1)), CLng(pvOption(2))
                End If
            End If
        Case "MSFLEXGRID"
        Case "MSHFLEXGRID"
        Case "CELL"
    End Select
    
End Sub

' *******************************************************************
' *   Member Name: RealLocateCBO                                    *
' *   Brief Description: 根据匹配串实时下拉列表框位置               *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub RealLocateCBO(pszMatchSZ As String, poLocateObject As ComboBox, Optional pbShowLinked As Boolean = True, Optional pbMustSelect As Boolean = False)
'参数注释
'*************************************
'pszMatchSZ(匹配串)
'poLocateObject(定位控件)
'pbShowLinked(是否显示联想词组)
'pbMustSelect(是否必须选择一项)
'************************************
    
    Dim lIndex As Long
    If poLocateObject.ListCount = 0 Then Exit Sub
    lIndex = -1
    If Len(pszMatchSZ) > 0 Then
        lIndex = SendMessage(poLocateObject.hwnd, CB_FINDSTRING, -1, ByVal pszMatchSZ)
    End If
    If lIndex = -1 Then '无匹配串
        If Not pbShowLinked Then Exit Sub
        If pbMustSelect Then poLocateObject.ListIndex = 0
        Exit Sub
    Else
        If pbShowLinked Then
            lIndex = SendMessage(poLocateObject.hwnd, CB_SELECTSTRING, -1, ByVal pszMatchSZ)
        End If
    End If
    
    If pbShowLinked Then     '显示联想词组
        poLocateObject.SelStart = Len(pszMatchSZ)
        poLocateObject.SelLength = Len(poLocateObject.Text) - Len(pszMatchSZ)
    End If

End Sub
' *******************************************************************
' *   Member Name: RealLocateLST                                    *
' *   Brief Description: 根据匹配串实时列表框位置                   *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub RealLocateLST(pszMatchSZ As String, poLocateObject As ListBox)
'参数注释
'*************************************
'pszMatchSZ(匹配串)
'poLocateObject(定位控件)
'************************************
    Dim lIndex As Long
    lIndex = -1
    If Len(pszMatchSZ) > 0 Then
        lIndex = SendMessage(poLocateObject.hwnd, LB_SELECTSTRING, -1, ByVal pszMatchSZ)
    End If
    If lIndex = -1 Then '无匹配串
        Exit Sub
    End If
    
End Sub

' *******************************************************************
' *   Member Name: DropDownCBO                                      *
' *   Brief Description: 展开下拉列表框                             *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub DropDownCBO(poLocateObject As ComboBox)
    SendMessage poLocateObject.hwnd, CB_SHOWDROPDOWN, True, ByVal 0&
End Sub

' *******************************************************************
' *   Member Name: RealLocateLVW                                    *
' *   Brief Description: 根据匹配串实时ListView框位置               *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub RealLocateLVW(pszMatchSZ As String, poLocateObject As Object, Optional pnColumnIndex As Integer = 1, Optional plStartIndex As Long = 1)
'参数注释
'*************************************
'pszMatchSZ(匹配串)
'poLocateObject(定位控件)
'pnColumnIndex(定位字段)
'plStartIndex(开始搜索位置)
'************************************
    Dim i As Long
    Dim oFoundItem As Object
    Dim lFoundIndex As Long
        
        
    If pnColumnIndex < 1 Then Exit Sub
    If pnColumnIndex > poLocateObject.ColumnHeaders.count Then Exit Sub
    If plStartIndex < 1 Then plStartIndex = 1
    
    pszMatchSZ = UCase(pszMatchSZ)
    Dim nTmpLen As Integer
    nTmpLen = Len(pszMatchSZ)
    If nTmpLen = 0 Then Exit Sub
    
    lFoundIndex = -1
    If pnColumnIndex = 1 Then       '主文本区域
        Set oFoundItem = poLocateObject.FindItem(pszMatchSZ, , , 1)
        If Not (oFoundItem Is Nothing) Then
            lFoundIndex = oFoundItem.Index
        End If
    Else        '其他区域
        Dim szTmp As String
        For i = plStartIndex To poLocateObject.ListItems.count
            szTmp = UCase(poLocateObject.ListItems(i).SubItems(pnColumnIndex - 1))
            If pszMatchSZ = Left(szTmp, nTmpLen) Then
                lFoundIndex = i
                Exit For
            End If
        Next i
    End If
    If lFoundIndex = -1 Then Exit Sub
    '定位
    If oFoundItem Is Nothing Then
        Set oFoundItem = poLocateObject.ListItems(lFoundIndex)
    End If
    oFoundItem.EnsureVisible
    oFoundItem.Selected = True
End Sub
' *******************************************************************
' *   Member Name: FormatTextToNumeric                              *
' *   Brief Description: 以指定数字格式控制TEXTBOX的文本            *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Function FormatTextToNumeric(poTextBox As Control, Optional pbCanBeNegative As Boolean = True, Optional pbCanBeDecimal As Boolean = True)
'参数注释
'*************************************
'poTextBox(控件)
'pbCanBeNegative(是否可以为负数)
'pbCanBeDecimal(是否可以为小数)
'************************************
    poTextBox.Text = GetTextToNumeric(poTextBox.Text, pbCanBeNegative, pbCanBeDecimal)
    poTextBox.SelStart = Len(poTextBox.Text)
End Function

' *******************************************************************
' *   Member Name: BulidTreeByArray                              *
' *   Brief Description: 将经排序后的数组建成树结构            *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub BulidTreeByArray(poTree As TreeView, paszContent() As String)
'参数注释
'*************************************
'poTree(控件)
'paszContent(内容数组)两维数组,第一维代表元素,第二维代表层次
'************************************
    Dim lArrayLen As Long
    Dim nDeeply As Integer
    Dim aszCloumnValue() As String
    
    '验证有效性
    If UCase(TypeName(poTree)) <> "TREEVIEW" Then
        Exit Sub
    End If
    lArrayLen = ArrayLength(paszContent)
    nDeeply = ArrayLength(paszContent, 2) - 1   '去掉选项字符串项
    If lArrayLen = 0 Or nDeeply <= 0 Then
        Exit Sub
    End If
    
    
    ReDim aszCloumnValue(1 To nDeeply, 1 To 2)
    Dim i As Long, j As Long, k As Long
    For i = 1 To nDeeply
        aszCloumnValue(i, 1) = cszUnrepeatString      '保证不重复
        aszCloumnValue(i, 2) = "0"
    Next i
    
    Dim szKey As String, szParentKey As String
    Dim oNode As Node

    i = 1
    While i <= lArrayLen
        szKey = "K"
        For j = 1 To nDeeply
            szParentKey = szKey     '父结点键值
            
            If aszCloumnValue(j, 1) <> paszContent(i, j) Then
                If paszContent(i, j) = "" Then Exit For
                aszCloumnValue(j, 1) = paszContent(i, j)
                aszCloumnValue(j, 2) = Val(aszCloumnValue(j, 2)) + 1
                szKey = szKey & "_" & aszCloumnValue(j, 2)
                For k = j + 1 To nDeeply
                    aszCloumnValue(k, 1) = cszUnrepeatString
                Next k
'                oNode.Key = szKey
'                oNode.Text = paszContent(i, j)
                If j > 1 Then
                    Set oNode = poTree.Nodes.Add(szParentKey, tvwChild, szKey, paszContent(i, j))
                Else
                    Set oNode = poTree.Nodes.Add(, , szKey, paszContent(i, j))
                End If
                
                '进行辅助选项设置
                SetNodeOption oNode, paszContent(i, 0)
                
                If j < nDeeply Then   '继续子结点
                    GoTo flagLoop
                End If
            Else
                If j + 1 > nDeeply Then
                    aszCloumnValue(j, 2) = Val(aszCloumnValue(j, 2)) + 1
                End If
                szKey = szKey & "_" & aszCloumnValue(j, 2)
            End If
        Next j
        i = i + 1
flagLoop:
    Wend
        
    
End Sub


Private Sub SetNodeOption(poNode As Node, ByVal pszNodeOption As String)
    Dim aszOption() As String
    Dim szTmpKey As String, szTmpValue As String
    Dim k As Integer
    aszOption = SplitEncodeStringArray(pszNodeOption)
    Dim nlen As Integer
    nlen = ArrayLength(aszOption)
    If nlen > 0 Then
        For k = 1 To nlen
            UnEncodeKeyValue aszOption(k), szTmpKey, szTmpValue
            Select Case UCase(szTmpKey)
                Case "TAG"
                    poNode.Tag = szTmpValue
                Case "IMAGE"
                    poNode.Image = Val(szTmpValue)
                Case "SELECTEDIMAGE"
                    poNode.SelectedImage = Val(szTmpValue)
            End Select
        Next k
    End If
End Sub
Public Sub ArrayListView(pszButtonKey As String, poArrayObject As ListView)
    Select Case pszButtonKey
        Case "大图标"
            '应做:添加 '大图标' 按钮代码。
            poArrayObject.View = lvwIcon
        Case "小图标"
            '应做:添加 '小图标' 按钮代码。
            poArrayObject.View = lvwSmallIcon
        Case "列表"
            '应做:添加 '列表' 按钮代码。
            poArrayObject.View = lvwList
        Case "详细资料"
            '应做:添加 '详细资料' 按钮代码。
            poArrayObject.View = lvwReport
    End Select
End Sub

' *******************************************************************
' *   Member Name: AutoAlignListViewHeadWidth                              *
' *   Brief Description: 自动ListView的列头宽度            *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub AlignFormPos(poForm As Form)
'参数注释
'*************************************
'poForm(表单)
'************************************
    
    Dim szIniFileName As String '配置文件路径与名称
    szIniFileName = App.Path & "\HeadSet.ini"
    Dim szSectionName As String '块区
    szSectionName = poForm.name
    Dim lKeyValue As Long
    lKeyValue = GetPrivateProfileInt(szSectionName, "Left", -1, szIniFileName)
    If lKeyValue <> -1 Then  '缺省设置
        poForm.Left = lKeyValue
    End If
    lKeyValue = GetPrivateProfileInt(szSectionName, "Top", -1, szIniFileName)
    If lKeyValue <> -1 Then  '缺省设置
        poForm.Top = lKeyValue
    End If

End Sub
' *******************************************************************
' *   Member Name: AutoAlignListViewHeadWidth                              *
' *   Brief Description: 自动ListView的列头宽度            *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub SaveFormPos(poForm As Form)
'参数注释
'*************************************
'poForm(表单)
'************************************
    
    Dim szIniFileName As String '配置文件路径与名称
    szIniFileName = App.Path & "\HeadSet.ini"
    Dim szSectionName As String '块区
    szSectionName = poForm.name
    Dim lKeyValue As Long
    lKeyValue = poForm.Left
    WritePrivateProfileString szSectionName, "Left", lKeyValue, szIniFileName
    lKeyValue = poForm.Top
    WritePrivateProfileString szSectionName, "Top", lKeyValue, szIniFileName
End Sub


Private Function StringWidth(ByVal pszString As String) As Long
    Dim ncntFontSize As Integer
    '缺省为小五号字体
    ncntFontSize = 180 / 2    '小五号字为180twips(半个汉字为90twips)
    StringWidth = ncntFontSize * (LenA(pszString))
End Function
' *******************************************************************
' *   Member Name: AutoAlignListViewHeadWidth                              *
' *   Brief Description: 自动ListView的列头宽度            *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub AlignHeadWidth(pszFormName As String, poObject As Object)
'参数注释
'*************************************
'poListView(控件)
'************************************
'根据HeadWidth
    Dim lHeadWidth As Long
    
    Dim szObjectType As String
    szObjectType = UCase(TypeName(poObject))
    If szObjectType <> "LISTVIEW" And szObjectType <> "MSFLEXGRID" And szObjectType <> "MSHFLEXGRID" And szObjectType <> "VSFLEXGRID" Then Exit Sub
    Dim szIniFileName As String '配置文件路径与名称
    szIniFileName = App.Path & "\HeadSet.ini"
    Dim szSectionName As String '块区
    szSectionName = pszFormName
    Dim szKeyName As String     '键名
    szKeyName = poObject.name & "_head"
    Dim lKeyValue As Long       '键值
    Dim i As Integer
    
    Select Case szObjectType
        Case "LISTVIEW"
            With poObject.ColumnHeaders
            For i = 1 To .count
                lKeyValue = GetPrivateProfileInt(szSectionName, szKeyName & i, -1, szIniFileName)
                If lKeyValue = -1 Then  '缺省设置
                    lKeyValue = StringWidth(.Item(i) & "  ")
                End If
                .Item(i).Width = lKeyValue
            Next i
            End With
        Case "MSFLEXGRID", "MSHFLEXGRID", "VSFLEXGRID"
            For i = 0 To poObject.Cols - 1
                lKeyValue = GetPrivateProfileInt(szSectionName, szKeyName & i, -1, szIniFileName)
                If lKeyValue <> -1 Then
                    poObject.ColWidth(i) = lKeyValue
                End If
            Next i
    End Select


End Sub
' *******************************************************************
' *   Member Name: AutoAlignListViewHeadWidth                              *
' *   Brief Description: 自动ListView的列头宽度            *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub SaveHeadWidth(pszFormName As String, poObject As Object)
'参数注释
'*************************************
'poListView(控件)
'************************************
'根据HeadWidth
    Dim lHeadWidth As Long
    
    Dim szObjectType As String
    szObjectType = UCase(TypeName(poObject))
    If szObjectType <> "LISTVIEW" And szObjectType <> "MSFLEXGRID" And szObjectType <> "MSHFLEXGRID" And szObjectType <> "VSFLEXGRID" Then Exit Sub
    Dim szIniFileName As String '配置文件路径与名称
    szIniFileName = App.Path & "\HeadSet.ini"
    Dim szSectionName As String '块区
    szSectionName = pszFormName
    Dim szKeyName As String     '键名
    szKeyName = poObject.name & "_head"
    Dim lKeyValue As Long       '键值
    Dim i As Integer
    
    Select Case szObjectType
        Case "LISTVIEW"
            With poObject.ColumnHeaders
            For i = 1 To .count
                lKeyValue = .Item(i).Width
                WritePrivateProfileString szSectionName, szKeyName & i, lKeyValue, szIniFileName
            Next i
            End With
        Case "MSFLEXGRID", "MSHFLEXGRID", "VSFLEXGRID"
            For i = 0 To poObject.Cols - 1
                lKeyValue = poObject.ColWidth(i)
                WritePrivateProfileString szSectionName, szKeyName & i, lKeyValue, szIniFileName
            Next i
    End Select


End Sub



'Public Sub ReadRegWidth(pszAppName As String, pszSubKey As String, poObject As Control, pszSection As String, Optional plRoot As ERegRoot = HKEY_LOCAL_MACHINE, Optional pEControlType As EControlType = ListViewControl)
''例如：ReadRegWidth cszRegKeyCommon, cszRegKeyTPSchool, lvLog, "Log"
'
'    '得到注册表中的列表宽度
'    Dim oReg As New CFreeReg
'    Dim i As Integer
'    Dim nCount As Integer
'    On Error GoTo ErrorHandle
'    oReg.Init pszAppName, plRoot, pszSubKey
'
'    Select Case pEControlType
'        Case ListViewControl
'            nCount = poObject.ColumnHeaders.Count
'        Case MSHFlexGridControl
'            nCount = poObject.cols - 1
'        Case MSFlexGridControl
'            nCount = poObject.cols - 1
'    End Select
'
'    For i = 1 To nCount
'        Select Case pEControlType
'            Case ListViewControl
'                poObject.ColumnHeaders(i).Width = CSng(oReg.GetSetting(pszSection, cszWidth & CStr(i), "1440"))
'            Case MSHFlexGridControl
'                poObject.ColWidth(i) = CSng(oReg.GetSetting(pszSection, cszWidth & CStr(i), "1000"))
'            Case MSFlexGridControl
'                poObject.ColWidth(i) = CSng(oReg.GetSetting(pszSection, cszWidth & CStr(i), "1000"))
'        End Select
'
'
'    Next i
'
'Exit Sub
'ErrorHandle:
'    MsgBox err.Description, vbExclamation, "错误--" & err.Number
'End Sub


'Public Sub SaveRegWidth(pszAppName As String, pszSubKey As String, poObject As Object, pszSection As String, Optional plRoot As ERegRoot = HKEY_LOCAL_MACHINE, Optional pEControlType As EControlType = ListViewControl)
''例如：SaveRegWidth cszRegKeyCommon, cszRegKeyTPSchool, lvLog, "Log"
'    '保存注册表中的列表宽度
'    Dim oReg As New CFreeReg
'    Dim i As Integer
'    Dim nCount As Integer
'    On Error GoTo ErrorHandle
'    oReg.Init pszAppName, plRoot, pszSubKey
'
'
'    Select Case pEControlType
'        Case ListViewControl
'            nCount = poObject.ColumnHeaders.Count
'        Case MSHFlexGridControl
'            nCount = poObject.cols - 1
'        Case MSFlexGridControl
'            nCount = poObject.cols - 1
'    End Select
'
'    For i = 1 To nCount
'        Select Case pEControlType
'            Case ListViewControl
'                oReg.SaveSetting pszSection, cszWidth & CStr(i), poObject.ColumnHeaders(i).Width
'            Case MSHFlexGridControl
'                oReg.SaveSetting pszSection, cszWidth & CStr(i), poObject.ColWidth(i)
'            Case MSFlexGridControl
'                oReg.SaveSetting pszSection, cszWidth & CStr(i), poObject.ColWidth(i)
'        End Select
'    Next i
'    Exit Sub
'ErrorHandle:
'    MsgBox err.Description, vbExclamation, "错误--" & err.Number
'End Sub

' *******************************************************************
' *   Member Name: FormatTextBoxBySize                              *
' *   Brief Description: 按长度格式化TextBox中TEXT内容              *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub FormatTextBoxBySize(poTextBox As Object, pnSize As Integer)
'参数注释
'*************************************
'poTextBox(TextBox)
'pnSize(指定长度)
'*************************************
    Dim szTmp As String
    szTmp = GetUnicodeBySize(poTextBox.Text, pnSize)
    If szTmp <> poTextBox.Text Then
        poTextBox.Text = szTmp
        poTextBox.SelStart = Len(szTmp)
    End If
End Sub


'设置MSHFLEX的对齐方式
Public Sub SetAlign(oTemp As Object, paszFormatString() As String)
    Dim szFormat As String
    Dim i As Integer
    Dim nCount As Integer
    Dim szTemp As String
    With oTemp
        nCount = ArrayLength(paszFormatString)
        If oTemp.Cols <> nCount And nCount <> 0 Then
            MsgBox "格式数组长度必须与列数相同", vbExclamation, "错误"
            Exit Sub
        End If
        For i = 0 To .Cols - 1
            If nCount <> 0 Then
                szTemp = paszFormatString(i)
            Else
                szTemp = cszLeftAlignString
            End If
            szFormat = szFormat & szTemp & .TextMatrix(0, i) & "|"
        Next i
        szFormat = Left(szFormat, Len(szFormat) - 1)
        
        .FormatString = szFormat
    
    End With
End Sub


Public Sub ShowErrorMsg()
    MsgBox err.Description, vbExclamation, "错误-" & err.Number
End Sub


'将ListView排序
Public Sub SortListView(plvListView As Object, ByVal pnIndex As Integer)
    Dim nTemp As Integer
    If plvListView.Tag = "" Then
        nTemp = -1
    Else
        nTemp = CInt(plvListView.Tag)
    End If
    
    If nTemp = pnIndex - 1 Then
        plvListView.SortOrder = IIf(plvListView.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        plvListView.Tag = pnIndex - 1
        plvListView.SortOrder = lvwAscending
    End If
    plvListView.SortKey = pnIndex - 1
    plvListView.Sorted = True
    If Not plvListView.SelectedItem Is Nothing Then
        plvListView.SelectedItem.EnsureVisible
    End If
End Sub

'设置ListView的整行颜色
Public Sub SetListViewLineColor(plvList As ListView, plIndex As Variant, ByVal plColor As Long)
    Dim oListItem As ListItem
    Set oListItem = plvList.ListItems(plIndex)
    oListItem.ForeColor = plColor
    Dim i As Integer
    For i = 1 To oListItem.ListSubItems.count
        oListItem.ListSubItems(i).ForeColor = plColor
    Next i
'    plvList.Refresh
End Sub

Public Function SeekListIndex(poCombox As ComboBox, pszString As String) As Integer
    If poCombox.ListCount = 0 Then
        SeekListIndex = -1
        Exit Function
    End If
    Dim i As Integer
    For i = 0 To poCombox.ListCount - 1
        If UCase(pszString) = UCase(poCombox.List(i)) Then
            Exit For
        End If
    Next i
    If i = poCombox.ListCount Then
        i = -1  '找不到
    End If
    SeekListIndex = i
End Function



'///////////////////////////////
'判断是否有打印机
Public Function IsPrinterValid() As Boolean
Dim Success As Boolean, cbRequired As Long, cbBuffer As Long
Dim Buffer() As Long, nEntries As Long
   cbBuffer = 3072
   ReDim Buffer((cbBuffer \ 4) - 1) As Long
   Success = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or _
                         PRINTER_ENUM_LOCAL, _
                         vbNullString, _
                         1, _
                         Buffer(0), _
                         cbBuffer, _
                         cbRequired, _
                         nEntries)
   If nEntries <> 0 Then
        IsPrinterValid = True
   Else
        IsPrinterValid = False
   End If
End Function


'指示信息框显示函数
Public Function ShowMsg(pszMsg As String) As Integer
    MsgBox pszMsg, vbOKOnly Or vbInformation, "注意"
End Function

'关机函数
Public Sub CloseComputer()
'Dim nReturn As Long
'
'If MsgBox("是否关机？", vbInformation + vbYesNo, "关机") = vbYes Then
'
'    nReturn = ExitWindowsEx(EWX_SHUTDOWN, 0)
'End If

End Sub


'Public Function GetEncodedKey(ByVal pszOrgCode As String) As String
'    GetEncodedKey = "A" & pszOrgCode
'End Function

'////////////////////////////////////
'登录参数转换
Public Function TransferLoginParam(szLoginParam As String) As String
    Dim szResult As String
    Dim szUserName As String
    Dim szUserPassword As String
    Dim pszcommandin As String
    '取用户名
    szUserName = Trim(LeftAndRight(szLoginParam, True, ","))
    If szUserName = "" Then Exit Function
    
    szResult = MakeLoginString(cszUserName, szUserName)
    pszcommandin = LeftAndRight(szLoginParam, False, ",")

    '取用户口令
    szUserPassword = LeftAndRight(pszcommandin, True, ",")
    szResult = szResult & MakeLoginString(cszUserPassword, szUserPassword)
    
    TransferLoginParam = szResult
End Function



Private Function MakeLoginString(pszCmd As String, pszValue As String) As String
    MakeLoginString = cszPrefixFlag & pszCmd & "=" & pszValue & cszSuffixFlag
End Function


'/////////////////////////////////
'得到登录的参数
Public Function GetLoginParam(szLoginCommand As String, szLoginParam As String, Optional pnStartSearch As Integer = 1) As String
    Dim szTemp As String, szValue As String
    Dim nTemp1 As String, nTemp2 As String
    szValue = ""
    szTemp = cszPrefixFlag & Trim(szLoginParam) & "="
    nTemp1 = InStr(pnStartSearch, szLoginCommand, szTemp, vbTextCompare)
    If nTemp1 > 0 Then
        nTemp1 = nTemp1 + Len(szTemp)
        nTemp2 = InStr(nTemp1, szLoginCommand, cszSuffixFlag, vbTextCompare)
        If nTemp2 > 0 Then
            szValue = Trim(Mid(szLoginCommand, nTemp1, nTemp2 - nTemp1))
        End If
    End If
    GetLoginParam = szValue
End Function
'打开具有关联运行程序的文件
Public Function OpenLinkedFile(szFileName As String) As Boolean
    ShellExecute 0, "Open", szFileName, 0, "", vbNormalFocus
End Function

Public Function ReplaceEnterKey(pvParam As Variant) As Variant
On Error GoTo here
    
    ReplaceEnterKey = Replace(Replace(pvParam, Chr(10), ""), Chr(13), "")
    
    Exit Function
here:
    WriteErrorLog "Insurance-ReplaceEnterKey", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
End Function

'写错误日志
Public Sub WriteErrorLog(pszSubFunctionName As String, pszErrString As String)

    Dim szLog As String
    Dim FileNo As Integer
    Dim szDir As String
    FileNo = FreeFile
    szLog = "_sell.txt"
               
    szDir = App.Path & "\" & Format(Date, "YYYYMMDD") & szLog
    If Dir(szDir) <> "" Then
        Open szDir For Append As #FileNo
            Print #FileNo, Format(Now, "yyyy-mm-dd HH:MM:SS") & " " & pszSubFunctionName & " " & pszErrString
        Close #FileNo
    Else
        Open szDir For Output As #FileNo
            Print #FileNo, Format(Now, "yyyy-mm-dd HH:MM:SS") & " " & pszSubFunctionName & " " & pszErrString
        Close #FileNo
    End If
End Sub

'字符串累加通用算法
Public Function StrAdd(ByVal a As String, ByVal b As Integer)
    Dim nl As Integer, i As Integer, j As Integer
    Dim tema1 As String, tema2 As String, temf As String
    
    nl = Len(a)
    For i = nl To 1 Step -1
        If Asc(Mid(a, i, 1)) < 48 Or Asc(Mid(a, i, 1)) > 57 Then Exit For
    Next i
    
    tema1 = Left(a, i)
    tema2 = Right(a, nl - i)
    
    For j = 1 To nl - i
        temf = temf & "0"
    Next j
    
    StrAdd = tema1 & Format(Val(tema2) + b, temf)
End Function
