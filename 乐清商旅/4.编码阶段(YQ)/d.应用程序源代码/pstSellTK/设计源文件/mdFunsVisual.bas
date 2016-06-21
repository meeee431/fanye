Attribute VB_Name = "mdFunsVisual"
Option Explicit

' *******************************************************************
' *  Source File Name: mdFunsVisual                                 *
' *  Brief Description: �ṩ���ڽ����һϵ�й��ܺ���                *
' *******************************************************************
Public Enum EStatusBarPanelArea '״̬������
    areaUserInfo = 1
    areaProgramInfo = 2
    areaProcessInfo = 3
End Enum
Public Enum EIndexSystemButton  'ϵͳ��ť����
    LogMan = 1
    ExportToFile = 3
    ExportAndOpen = 4
    PrintData = 6
    PrintPreview = 7
    ExitSystem = 9
End Enum
'Public Enum ESearchAreaIndex     '����λ��
'    ESI_TextArea = 1             '���ı�ƥ��
'    ESI_SubItemArea = -1            'SUBITEM��
'    ESI_WholeArea = -2              'ȫ������
'End Enum

'
'''����ĵ�ǰ״̬
'Public Enum eFormStatus
'    AddStatus = 0
'    ModifyStatus = 1
'    ShowStatus = 2
'    NotStatus = 3
'End Enum
''

'����ȫ�ֳ�������
'�����ַ���
Public Const cszLeftAlignString = "<"
Public Const cszRightAlignString = ">"
Public Const cszMiddleAlignString = "^"

Public Const cszPrefixFlag = "<<"
Public Const cszSuffixFlag = ">>"
Public Const cszUserName = "UserID"
Public Const cszUserPassword = "Password"



'����ȫ�ֱ�������

'����API����


Public Const HWND_BOTTOM = 1 '���������ڴ����б�ײ�
Public Const HWND_TOP = 0 '����������Z���еĶ�����Z���д����ڷּ��ṹ�У��������һ����������Ĵ�����ʾ��˳��
Public Const HWND_TOPMOST = -1 '�����������б�������λ���κ�������ڵ�ǰ��
Public Const HWND_NOTOPMOST = -2 '�����������б�������λ���κ�������ڵĺ���

Public Const SWP_HIDEWINDOW = &H80 '���ش���
Public Const SWP_NOACTIVATE = &H10 '�������
Public Const SWP_NOMOVE = &H2 '���ֵ�ǰλ��(x��y�趨��������)
Public Const SWP_NOREDRAW = &H8 '���ڲ��Զ��ػ�
Public Const SWP_NOSIZE = &H1 '���ֵ�ǰ��С(cx��cy�ᱻ����)
Public Const SWP_NOZORDER = &H4 '���ִ������б�ĵ�ǰλ��(hWndInsertAfter��������)
Public Const SWP_SHOWWINDOW = &H40 '��ʾ����
Public Const SWP_FRAMECHANGED = &H20 'ǿ��һ��WM_NCCALCSIZE��Ϣ���봰�ڣ���ʹ���ڵĴ�Сû�иı�
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED 'Χ�ƴ��ڻ�һ����

Public Const LB_SELECTSTRING = &H18C
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_FINDSTRING = &H14C


Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'�ж��Ƿ��д�ӡ������������
Const PRINTER_ENUM_CONNECTIONS = &H4
Const PRINTER_ENUM_LOCAL = &H2
Public Declare Function EnumPrinters Lib "winspool.drv" Alias _
   "EnumPrintersA" (ByVal Flags As Long, ByVal name As String, _
   ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, _
   pcbNeeded As Long, pcReturned As Long) As Long
'�ػ�����
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
' *   Brief Description: ����ƥ�䴮ʵʱ��λ�ؼ��б�λ��             *
' *   Engineer: ½����                                              *
' *******************************************************************

Public Sub RealLocate(pszMatchSZ As String, poLocateObject As Object, Optional pvOption As Variant)
'����ע��
'*************************************
'pszMatchSZ(ƥ�䴮)
'poLocateObject(��λ�ؼ�)
'pnMatchColumn(ƥ���ֶ�)
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
' *   Brief Description: ����ƥ�䴮ʵʱ�����б��λ��               *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub RealLocateCBO(pszMatchSZ As String, poLocateObject As ComboBox, Optional pbShowLinked As Boolean = True, Optional pbMustSelect As Boolean = False)
'����ע��
'*************************************
'pszMatchSZ(ƥ�䴮)
'poLocateObject(��λ�ؼ�)
'pbShowLinked(�Ƿ���ʾ�������)
'pbMustSelect(�Ƿ����ѡ��һ��)
'************************************
    
    Dim lIndex As Long
    If poLocateObject.ListCount = 0 Then Exit Sub
    lIndex = -1
    If Len(pszMatchSZ) > 0 Then
        lIndex = SendMessage(poLocateObject.hwnd, CB_FINDSTRING, -1, ByVal pszMatchSZ)
    End If
    If lIndex = -1 Then '��ƥ�䴮
        If Not pbShowLinked Then Exit Sub
        If pbMustSelect Then poLocateObject.ListIndex = 0
        Exit Sub
    Else
        If pbShowLinked Then
            lIndex = SendMessage(poLocateObject.hwnd, CB_SELECTSTRING, -1, ByVal pszMatchSZ)
        End If
    End If
    
    If pbShowLinked Then     '��ʾ�������
        poLocateObject.SelStart = Len(pszMatchSZ)
        poLocateObject.SelLength = Len(poLocateObject.Text) - Len(pszMatchSZ)
    End If

End Sub
' *******************************************************************
' *   Member Name: RealLocateLST                                    *
' *   Brief Description: ����ƥ�䴮ʵʱ�б��λ��                   *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub RealLocateLST(pszMatchSZ As String, poLocateObject As ListBox)
'����ע��
'*************************************
'pszMatchSZ(ƥ�䴮)
'poLocateObject(��λ�ؼ�)
'************************************
    Dim lIndex As Long
    lIndex = -1
    If Len(pszMatchSZ) > 0 Then
        lIndex = SendMessage(poLocateObject.hwnd, LB_SELECTSTRING, -1, ByVal pszMatchSZ)
    End If
    If lIndex = -1 Then '��ƥ�䴮
        Exit Sub
    End If
    
End Sub

' *******************************************************************
' *   Member Name: DropDownCBO                                      *
' *   Brief Description: չ�������б��                             *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub DropDownCBO(poLocateObject As ComboBox)
    SendMessage poLocateObject.hwnd, CB_SHOWDROPDOWN, True, ByVal 0&
End Sub

' *******************************************************************
' *   Member Name: RealLocateLVW                                    *
' *   Brief Description: ����ƥ�䴮ʵʱListView��λ��               *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub RealLocateLVW(pszMatchSZ As String, poLocateObject As Object, Optional pnColumnIndex As Integer = 1, Optional plStartIndex As Long = 1)
'����ע��
'*************************************
'pszMatchSZ(ƥ�䴮)
'poLocateObject(��λ�ؼ�)
'pnColumnIndex(��λ�ֶ�)
'plStartIndex(��ʼ����λ��)
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
    If pnColumnIndex = 1 Then       '���ı�����
        Set oFoundItem = poLocateObject.FindItem(pszMatchSZ, , , 1)
        If Not (oFoundItem Is Nothing) Then
            lFoundIndex = oFoundItem.Index
        End If
    Else        '��������
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
    '��λ
    If oFoundItem Is Nothing Then
        Set oFoundItem = poLocateObject.ListItems(lFoundIndex)
    End If
    oFoundItem.EnsureVisible
    oFoundItem.Selected = True
End Sub
' *******************************************************************
' *   Member Name: FormatTextToNumeric                              *
' *   Brief Description: ��ָ�����ָ�ʽ����TEXTBOX���ı�            *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Function FormatTextToNumeric(poTextBox As Control, Optional pbCanBeNegative As Boolean = True, Optional pbCanBeDecimal As Boolean = True)
'����ע��
'*************************************
'poTextBox(�ؼ�)
'pbCanBeNegative(�Ƿ����Ϊ����)
'pbCanBeDecimal(�Ƿ����ΪС��)
'************************************
    poTextBox.Text = GetTextToNumeric(poTextBox.Text, pbCanBeNegative, pbCanBeDecimal)
    poTextBox.SelStart = Len(poTextBox.Text)
End Function

' *******************************************************************
' *   Member Name: BulidTreeByArray                              *
' *   Brief Description: �������������齨�����ṹ            *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub BulidTreeByArray(poTree As TreeView, paszContent() As String)
'����ע��
'*************************************
'poTree(�ؼ�)
'paszContent(��������)��ά����,��һά����Ԫ��,�ڶ�ά������
'************************************
    Dim lArrayLen As Long
    Dim nDeeply As Integer
    Dim aszCloumnValue() As String
    
    '��֤��Ч��
    If UCase(TypeName(poTree)) <> "TREEVIEW" Then
        Exit Sub
    End If
    lArrayLen = ArrayLength(paszContent)
    nDeeply = ArrayLength(paszContent, 2) - 1   'ȥ��ѡ���ַ�����
    If lArrayLen = 0 Or nDeeply <= 0 Then
        Exit Sub
    End If
    
    
    ReDim aszCloumnValue(1 To nDeeply, 1 To 2)
    Dim i As Long, j As Long, k As Long
    For i = 1 To nDeeply
        aszCloumnValue(i, 1) = cszUnrepeatString      '��֤���ظ�
        aszCloumnValue(i, 2) = "0"
    Next i
    
    Dim szKey As String, szParentKey As String
    Dim oNode As Node

    i = 1
    While i <= lArrayLen
        szKey = "K"
        For j = 1 To nDeeply
            szParentKey = szKey     '������ֵ
            
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
                
                '���и���ѡ������
                SetNodeOption oNode, paszContent(i, 0)
                
                If j < nDeeply Then   '�����ӽ��
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
        Case "��ͼ��"
            'Ӧ��:��� '��ͼ��' ��ť���롣
            poArrayObject.View = lvwIcon
        Case "Сͼ��"
            'Ӧ��:��� 'Сͼ��' ��ť���롣
            poArrayObject.View = lvwSmallIcon
        Case "�б�"
            'Ӧ��:��� '�б�' ��ť���롣
            poArrayObject.View = lvwList
        Case "��ϸ����"
            'Ӧ��:��� '��ϸ����' ��ť���롣
            poArrayObject.View = lvwReport
    End Select
End Sub

' *******************************************************************
' *   Member Name: AutoAlignListViewHeadWidth                              *
' *   Brief Description: �Զ�ListView����ͷ���            *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub AlignFormPos(poForm As Form)
'����ע��
'*************************************
'poForm(��)
'************************************
    
    Dim szIniFileName As String '�����ļ�·��������
    szIniFileName = App.Path & "\HeadSet.ini"
    Dim szSectionName As String '����
    szSectionName = poForm.name
    Dim lKeyValue As Long
    lKeyValue = GetPrivateProfileInt(szSectionName, "Left", -1, szIniFileName)
    If lKeyValue <> -1 Then  'ȱʡ����
        poForm.Left = lKeyValue
    End If
    lKeyValue = GetPrivateProfileInt(szSectionName, "Top", -1, szIniFileName)
    If lKeyValue <> -1 Then  'ȱʡ����
        poForm.Top = lKeyValue
    End If

End Sub
' *******************************************************************
' *   Member Name: AutoAlignListViewHeadWidth                              *
' *   Brief Description: �Զ�ListView����ͷ���            *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub SaveFormPos(poForm As Form)
'����ע��
'*************************************
'poForm(��)
'************************************
    
    Dim szIniFileName As String '�����ļ�·��������
    szIniFileName = App.Path & "\HeadSet.ini"
    Dim szSectionName As String '����
    szSectionName = poForm.name
    Dim lKeyValue As Long
    lKeyValue = poForm.Left
    WritePrivateProfileString szSectionName, "Left", lKeyValue, szIniFileName
    lKeyValue = poForm.Top
    WritePrivateProfileString szSectionName, "Top", lKeyValue, szIniFileName
End Sub


Private Function StringWidth(ByVal pszString As String) As Long
    Dim ncntFontSize As Integer
    'ȱʡΪС�������
    ncntFontSize = 180 / 2    'С�����Ϊ180twips(�������Ϊ90twips)
    StringWidth = ncntFontSize * (LenA(pszString))
End Function
' *******************************************************************
' *   Member Name: AutoAlignListViewHeadWidth                              *
' *   Brief Description: �Զ�ListView����ͷ���            *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub AlignHeadWidth(pszFormName As String, poObject As Object)
'����ע��
'*************************************
'poListView(�ؼ�)
'************************************
'����HeadWidth
    Dim lHeadWidth As Long
    
    Dim szObjectType As String
    szObjectType = UCase(TypeName(poObject))
    If szObjectType <> "LISTVIEW" And szObjectType <> "MSFLEXGRID" And szObjectType <> "MSHFLEXGRID" And szObjectType <> "VSFLEXGRID" Then Exit Sub
    Dim szIniFileName As String '�����ļ�·��������
    szIniFileName = App.Path & "\HeadSet.ini"
    Dim szSectionName As String '����
    szSectionName = pszFormName
    Dim szKeyName As String     '����
    szKeyName = poObject.name & "_head"
    Dim lKeyValue As Long       '��ֵ
    Dim i As Integer
    
    Select Case szObjectType
        Case "LISTVIEW"
            With poObject.ColumnHeaders
            For i = 1 To .count
                lKeyValue = GetPrivateProfileInt(szSectionName, szKeyName & i, -1, szIniFileName)
                If lKeyValue = -1 Then  'ȱʡ����
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
' *   Brief Description: �Զ�ListView����ͷ���            *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub SaveHeadWidth(pszFormName As String, poObject As Object)
'����ע��
'*************************************
'poListView(�ؼ�)
'************************************
'����HeadWidth
    Dim lHeadWidth As Long
    
    Dim szObjectType As String
    szObjectType = UCase(TypeName(poObject))
    If szObjectType <> "LISTVIEW" And szObjectType <> "MSFLEXGRID" And szObjectType <> "MSHFLEXGRID" And szObjectType <> "VSFLEXGRID" Then Exit Sub
    Dim szIniFileName As String '�����ļ�·��������
    szIniFileName = App.Path & "\HeadSet.ini"
    Dim szSectionName As String '����
    szSectionName = pszFormName
    Dim szKeyName As String     '����
    szKeyName = poObject.name & "_head"
    Dim lKeyValue As Long       '��ֵ
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
''���磺ReadRegWidth cszRegKeyCommon, cszRegKeyTPSchool, lvLog, "Log"
'
'    '�õ�ע����е��б���
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
'    MsgBox err.Description, vbExclamation, "����--" & err.Number
'End Sub


'Public Sub SaveRegWidth(pszAppName As String, pszSubKey As String, poObject As Object, pszSection As String, Optional plRoot As ERegRoot = HKEY_LOCAL_MACHINE, Optional pEControlType As EControlType = ListViewControl)
''���磺SaveRegWidth cszRegKeyCommon, cszRegKeyTPSchool, lvLog, "Log"
'    '����ע����е��б���
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
'    MsgBox err.Description, vbExclamation, "����--" & err.Number
'End Sub

' *******************************************************************
' *   Member Name: FormatTextBoxBySize                              *
' *   Brief Description: �����ȸ�ʽ��TextBox��TEXT����              *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub FormatTextBoxBySize(poTextBox As Object, pnSize As Integer)
'����ע��
'*************************************
'poTextBox(TextBox)
'pnSize(ָ������)
'*************************************
    Dim szTmp As String
    szTmp = GetUnicodeBySize(poTextBox.Text, pnSize)
    If szTmp <> poTextBox.Text Then
        poTextBox.Text = szTmp
        poTextBox.SelStart = Len(szTmp)
    End If
End Sub


'����MSHFLEX�Ķ��뷽ʽ
Public Sub SetAlign(oTemp As Object, paszFormatString() As String)
    Dim szFormat As String
    Dim i As Integer
    Dim nCount As Integer
    Dim szTemp As String
    With oTemp
        nCount = ArrayLength(paszFormatString)
        If oTemp.Cols <> nCount And nCount <> 0 Then
            MsgBox "��ʽ���鳤�ȱ�����������ͬ", vbExclamation, "����"
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
    MsgBox err.Description, vbExclamation, "����-" & err.Number
End Sub


'��ListView����
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

'����ListView��������ɫ
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
        i = -1  '�Ҳ���
    End If
    SeekListIndex = i
End Function



'///////////////////////////////
'�ж��Ƿ��д�ӡ��
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


'ָʾ��Ϣ����ʾ����
Public Function ShowMsg(pszMsg As String) As Integer
    MsgBox pszMsg, vbOKOnly Or vbInformation, "ע��"
End Function

'�ػ�����
Public Sub CloseComputer()
'Dim nReturn As Long
'
'If MsgBox("�Ƿ�ػ���", vbInformation + vbYesNo, "�ػ�") = vbYes Then
'
'    nReturn = ExitWindowsEx(EWX_SHUTDOWN, 0)
'End If

End Sub


'Public Function GetEncodedKey(ByVal pszOrgCode As String) As String
'    GetEncodedKey = "A" & pszOrgCode
'End Function

'////////////////////////////////////
'��¼����ת��
Public Function TransferLoginParam(szLoginParam As String) As String
    Dim szResult As String
    Dim szUserName As String
    Dim szUserPassword As String
    Dim pszcommandin As String
    'ȡ�û���
    szUserName = Trim(LeftAndRight(szLoginParam, True, ","))
    If szUserName = "" Then Exit Function
    
    szResult = MakeLoginString(cszUserName, szUserName)
    pszcommandin = LeftAndRight(szLoginParam, False, ",")

    'ȡ�û�����
    szUserPassword = LeftAndRight(pszcommandin, True, ",")
    szResult = szResult & MakeLoginString(cszUserPassword, szUserPassword)
    
    TransferLoginParam = szResult
End Function



Private Function MakeLoginString(pszCmd As String, pszValue As String) As String
    MakeLoginString = cszPrefixFlag & pszCmd & "=" & pszValue & cszSuffixFlag
End Function


'/////////////////////////////////
'�õ���¼�Ĳ���
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
'�򿪾��й������г�����ļ�
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

'д������־
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

'�ַ����ۼ�ͨ���㷨
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
