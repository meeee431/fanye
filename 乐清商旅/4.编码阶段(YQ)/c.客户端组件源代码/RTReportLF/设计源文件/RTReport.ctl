VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl RTReport 
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   ScaleHeight     =   3165
   ScaleWidth      =   4605
   ToolboxBitmap   =   "RTReport.ctx":0000
   Begin TTF160Ctl.F1Book F1Book 
      Height          =   2685
      Left            =   225
      TabIndex        =   0
      Top             =   210
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   4736
      _0              =   $"RTReport.ctx":0312
      _1              =   $"RTReport.ctx":071B
      _2              =   $"RTReport.ctx":0B24
      _3              =   $"RTReport.ctx":0F2D
      _4              =   $"RTReport.ctx":1336
      _count          =   5
      _ver            =   2
   End
   Begin MSComDlg.CommonDialog CommDialog 
      Left            =   4110
      Top             =   2670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "RTReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'睿通报表组件
Option Base 0
Option Explicit


'模板文档名
'##ModelId=3C5100CB023C
Private mszTemplateFile As String

'背景图片文件
'##ModelId=3C53252C0357
Private mszBackgroupFile As String
Private mbLeftLabelVisual As Boolean
Private mbTopLabelVisual As Boolean
Private mbAutoRowHeight As Boolean
Private mszSheetTitle As String


Private mlDataFormatLine As Long    '数据格式行
Private maszColumnStr() As String
Private matSortSum() As TFormatSumCell
Private matTotalSum() As TFormatSumCell
Private mlDataStartLine As Long
Dim mlStartFillLine As Long, mlEndFillLine As Long


Private m_arsCustomValue As Variant
Private m_aszCustomString As Variant
Private m_nCustomCount  As Integer 'm_aszCustomString的长度
Private Const cszNotFound = "*****"


Type TFormatSumCell '统计单元格
    col As Long     '数据列
    ToRow As Long   '与数据行的相对行
    Key As String   '类型
    KeyColumn As String
End Type
Type TMergeRange     '合并Range的信息
    Index As Long       '索引
    StartRow As Long
    EndRow As Long
    StartCol As Long
    EndCol As Long
End Type

Public Enum EFileType
    EXCEL_5_TYPE = 1
    EXCEL_97_TYPE = 2
    EXCEL_2000_TYPE = 3
End Enum
Public Enum EExportFileType
    HTML_TYPE = 1
    TEXT_TYPE = 2
    FORMULA_ONE_6_TYPE = 3
End Enum
Public Enum EDialogType
    EXPORT_FILE = 1
    SAVE_FILE = 2
    PRINTVIEW_TYPE = 3
    PRINT_TYPE = 4
    PAGESET_TYPE = 5
End Enum
    
'事件声明列表
Public Event Click(ByVal nRow As Long, ByVal nCol As Long)
Public Event RClick(ByVal nRow As Long, ByVal nCol As Long)
Public Event DblClick(ByVal nRow As Long, ByVal nCol As Long)
Public Event RDblClick(ByVal nRow As Long, ByVal nCol As Long)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Public Event StartReadTemplateFormat()          '开始解析模块文件
Public Event StartFillCustomData()              '开始填充自定义数据
Public Event StartFillContentData()             '开始填充记录
Public Event SetProgressRange(ByVal lRange)     '进度范围
Public Event SetProgressValue(ByVal lValue)     '进度事件

'打印
'##ModelId=3C5101060183
Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    On Error GoTo PrintErr

    F1Book.FilePrint pbShowDialog
    Exit Sub
PrintErr:
    Call RaiseError(Err.Number, "RTReport->Print", Err.Decription)
End Sub

'显示报表
'##ModelId=3C50FF8002A4
Public Sub ShowReport(Optional parsDataSource As Variant, Optional ByVal pvCustomData As Variant)
    On Error GoTo ShowReportErr
    Dim szModule As String
    szModule = "RTReport->ShowReport"

    '模板处理
    If mszTemplateFile = "" Then RaiseError ERR_TemplateFileNotFound, szModule

    '显示进度窗体
    ShowProgess 0, 100, True, "正在生成报表，请等待..."

'    F1Book.ClearRange 0, 0, F1Book
'    Cell1.ResetContent
    
    On Error Resume Next
    F1Book.ReadEx mszTemplateFile
    If Err Then
        On Error GoTo ShowReportErr
        RaiseError ERR_TemplateFileNotFound, szModule
    End If
    On Error GoTo ShowReportErr
    
                 
'    If CELL1.OpenFile(mszTemplateFile, "") <> 1 Then RaiseError ERR_TemplateFileNotFound, szModule
    
    '改观样式
    LeftLabelVisual = mbLeftLabelVisual
    TopLabelVisual = mbTopLabelVisual
    F1Book.ShowColHeading = TopLabelVisual
    F1Book.ShowRowHeading = LeftLabelVisual
    F1Book.TopRow = F1Book.MinRow
    F1Book.LeftCol = F1Book.MinCol
    F1Book.MaxRow = F1Book.LastRow
    F1Book.MaxCol = F1Book.LastCol
    If mszSheetTitle = "" Then
        F1Book.ShowTabs = F1TabsOff
    Else
        F1Book.ShowTabs = F1TabsBottom
        F1Book.Title = mszSheetTitle
        F1Book.SheetName(1) = mszSheetTitle
    End If
'    CELL1.ShowSheetLabel 0, 0

    Dim nRecordCount As Integer
    If IsNull(parsDataSource) Or VarType(parsDataSource) = vbError Then
       nRecordCount = 0
    Else
        If IsArray(parsDataSource) Then     '记录集数组
            nRecordCount = ArrayLength(parsDataSource)
        Else
            nRecordCount = 1
        End If
    End If

    Dim atmp() As TFormatSumCell
    matSortSum = atmp
    matTotalSum = atmp

    '填充自定义数据
    
        mlStartFillLine = F1Book.MinRow
        mlEndFillLine = F1Book.MaxRow
        
        FillCustomData pvCustomData
        
        '填充记录集
        Dim i As Integer
        ReadBodyDataFormat
        For i = 1 To nRecordCount
            If nRecordCount > 1 Then
                FillBodyData parsDataSource(i)
            Else
                FillBodyData parsDataSource
            End If
        Next i
        
    If nRecordCount > 0 Then
        '清除数据说明行
        DeleDataNoteLine
    End If
    '设置选中位置

    F1Book.Refresh
    F1Book.Row = F1Book.MinRow
    F1Book.col = F1Book.MinCol
    
    CloseProgess
    Exit Sub
ShowReportErr:
    CloseProgess
    MsgBox Err.Description, vbCritical, "错误:" & Err.Number
End Sub
'显示多记录报表
'##ModelId=3C50FF8002A4
Public Sub ShowMultiReport(Optional parsDataSource As Variant, Optional ByVal pvCustomData As Variant, Optional pnMinPageCount As Integer, Optional pbInsertPageSep As Boolean)
'pnMinPageCount 最小记录数
'pbInsertPageSep 是否分页显示
    On Error GoTo ShowReportErr
    Dim szModule As String
    szModule = "RTReport->ShowMultiReport"

    '模板处理
    If mszTemplateFile = "" Then RaiseError ERR_TemplateFileNotFound, szModule

    '显示进度窗体
    ShowProgess 0, 100, True, "正在生成报表，请等待..."

'    F1Book.ClearRange 0, 0, F1Book
'    Cell1.ResetContent
    
    On Error Resume Next
    F1Book.ReadEx mszTemplateFile
    If Err Then
        On Error GoTo ShowReportErr
        RaiseError ERR_TemplateFileNotFound, szModule
    End If
    On Error GoTo ShowReportErr
    
                 
'    If CELL1.OpenFile(mszTemplateFile, "") <> 1 Then RaiseError ERR_TemplateFileNotFound, szModule
    
    '改观样式
    LeftLabelVisual = mbLeftLabelVisual
    TopLabelVisual = mbTopLabelVisual
    F1Book.ShowColHeading = TopLabelVisual
    F1Book.ShowRowHeading = LeftLabelVisual
    F1Book.TopRow = F1Book.MinRow
    F1Book.LeftCol = F1Book.MinCol
    F1Book.MaxRow = F1Book.LastRow
    F1Book.MaxCol = F1Book.LastCol
    F1Book.Title = mszSheetTitle
    F1Book.SheetName(1) = mszSheetTitle
'    CELL1.ShowSheetLabel 0, 0

    Dim nRecordCount As Integer
    nRecordCount = ArrayLength(parsDataSource)
    If pnMinPageCount > nRecordCount Then
        nRecordCount = pnMinPageCount
    End If
    If nRecordCount = 0 Then RaiseError ERR_BodyDataInvalid, szModule
        
    mlStartFillLine = F1Book.MinRow
    mlEndFillLine = F1Book.LastRow
    Dim nTempLines As Integer
    nTempLines = F1Book.MaxRow
    Dim i As Integer, j As Integer, k As Integer
    For k = 1 To nRecordCount
        Dim atmp() As TFormatSumCell    '为了在每次调用时初始成空值
        matSortSum = atmp
        matTotalSum = atmp
    
        '填充自定义数据
'        Dim nStartCopyRow As Long
'        nStartCopyRow = F1Book.LastRow
        If k <> nRecordCount Then
            F1Book.SetSelection mlStartFillLine, F1Book.MinCol, mlEndFillLine, F1Book.LastCol
            F1Book.EditCopy
            If pbInsertPageSep Then F1Book.AddPageBreak
            F1Book.MaxRow = F1Book.MaxRow + nTempLines
            For i = mlStartFillLine To mlEndFillLine        '设置同高
                F1Book.RowHeight(i + nTempLines) = F1Book.RowHeight(i)
            Next i
            F1Book.SetSelection F1Book.LastRow + 1, F1Book.MinCol, F1Book.MaxRow, F1Book.LastCol
            F1Book.EditPaste
        End If

        Dim aszTmpCustomData() As String        '为了在每次调用时初始成空值
        Dim aszTmp() As String
        aszTmpCustomData = aszTmp
        
        If ArrayLength(pvCustomData, 1) >= k Then
            If ArrayLength(pvCustomData, 2) > 0 And ArrayLength(pvCustomData, 3) > 0 Then
                ReDim aszTmpCustomData(1 To ArrayLength(pvCustomData, 2), 1 To ArrayLength(pvCustomData, 3))
            End If
            For i = 1 To ArrayLength(pvCustomData, 2)
                For j = 1 To ArrayLength(pvCustomData, 3)
                    aszTmpCustomData(i, j) = pvCustomData(k, i, j)
                Next j
            Next i
            
            FillCustomData aszTmpCustomData
        End If
        If ArrayLength(parsDataSource) >= k Then
        '填充记录集
            ReadBodyDataFormat
            FillBodyData parsDataSource(k)
            
        '清除数据说明行
            DeleDataNoteLine
        End If
        
        mlStartFillLine = F1Book.LastRow - nTempLines + 1
        mlEndFillLine = F1Book.LastRow
    Next k
'    If mlDataStartLine > F1Book.MinRow Then     '删除多余的模板拷贝
'        F1Book.DeleteRange mlStartFillLine, F1Book.MinCol, mlEndFillLine, F1Book.LastCol, F1ShiftHorizontal
'        F1Book.MaxRow = F1Book.MaxRow - nTempLines
'    End If
        
    '设置选中位置
    F1Book.Refresh
    F1Book.Row = F1Book.MinRow
    F1Book.col = F1Book.MinCol
    
    CloseProgess
    Exit Sub
ShowReportErr:
    CloseProgess
    MsgBox Err.Description, vbCritical, "错误:" & Err.Number
End Sub


Private Sub FillCustomData(pvCustomData As Variant)
    Dim i As Long, j As Long, k As Long
    Dim nLeft As Integer, nRight As Integer
    Dim szString As String, szKey As String, szValue As String, szDestinString As String
    Dim szRight As String
    'Dim nCount As Integer
    'Dim aszString() As String
    'Dim l As Integer
    Dim szTemp As String
    Dim nDelColCount As Integer '删除的列数
    Dim nMinRow As Integer
    Dim nLastRow As Integer
    Dim nMinCol As Integer
    Dim nLastCol As Integer
    
    Dim oCellFormat As F1CellFormat
    Dim l As Integer
    
    Dim bNeedSetFont As Boolean '是否需要设置字体
    Dim szFontName As String '需要设置的字体
    Dim nFontSize As Integer   '需要设置的字体大小
    bNeedSetFont = IIf(ArrayLength(pvCustomData, 2) <= 2, False, True)
    
    '触发"开始填充自定义数据"事件
    RaiseEvent StartFillCustomData
    '数组长度
    
    With F1Book
        nMinRow = mlStartFillLine
        nLastRow = mlEndFillLine
        For i = nMinRow To nLastRow
            nDelColCount = 0
            nMinCol = .MinCol
            nLastCol = .LastCol
            For j = nMinCol To nLastCol
                
                szString = .TextRC(i, j - nDelColCount)
    '            szDestinString = szString
    
                szDestinString = LeftAndRight(szString, True, "[")
                szRight = LeftAndRight(szString, False, "[")
                If szRight <> "" Then szRight = "[" & szRight
                UnEncodeKeyValue szRight, szKey, szValue
                While Not (szKey = "" And szValue = "")
                    Select Case szKey
                        Case "日期", "时间"
                            If szValue = "" Then
                                If szKey = "日期" Then szValue = "YYYY-MM-DD" Else szValue = "HH:mm"
                            End If
                            szDestinString = szDestinString & Format(Now, szValue)
                        Case "自定义项目"
                            If VarType(pvCustomData) <> vbError Then
                                For k = 1 To ArrayLength(pvCustomData)
                                    If LCase(szValue) = LCase(pvCustomData(k, 1)) Then
                                        szDestinString = szDestinString & pvCustomData(k, 2)
                                        If bNeedSetFont Then
                                            szFontName = pvCustomData(k, 3)
                                            nFontSize = CInt(pvCustomData(k, 4))
                                        End If
                                        Exit For
                                    End If
                                Next k
                            End If
                        Case "空行"
                            
                        Case Else
                            szTemp = GetItemName(szKey, szValue)
                            If szTemp = cszNotFound Then
                                '如果未找到
                                szDestinString = szDestinString & LeftAndRight(szRight, True, "]") & "]"
                            ElseIf szTemp = "" Then
                                '有这个定制串,但未找到对应的值,则删除该列
                                
                                Dim oRange As F1RangeRef
                                Dim atMergeCells() As TMergeRange, nMergeCount As Long
                                nMergeCount = 0
                                For l = mlStartFillLine To .MaxRow
                                    F1Book.SetSelection l, j - nDelColCount, l, j - nDelColCount
                                    Set oRange = F1Book.SelectionEx(0)
                                    Set oCellFormat = F1Book.GetCellFormat
                                    If oCellFormat.MergeCells = True Then
                                         nMergeCount = nMergeCount + 1
                                         ReDim Preserve atMergeCells(1 To nMergeCount)      '将合并信息保留
                                         atMergeCells(nMergeCount).StartCol = oRange.StartCol
                                         atMergeCells(nMergeCount).EndCol = oRange.EndCol
                                         atMergeCells(nMergeCount).StartRow = oRange.StartRow
                                         atMergeCells(nMergeCount).EndRow = oRange.EndRow
                                        oCellFormat.MergeCells = False      '将合并列拆分以便于删除
                                        F1Book.SetCellFormat oCellFormat
                                        '将在原有合并格中的内容转移至后一格，为了保持数据的正确性
                                        If j - nDelColCount + 1 <= F1Book.LastCol Then
                                            F1Book.TextRC(l, j - nDelColCount + 1) = F1Book.TextRC(l, j - nDelColCount)
                                        End If
'                                        F1Book.SetSelection l, j - nDelColCount, l, j - nDelColCount
'                                        F1Book.EditCopy
'                                        F1Book.SetSelection l, j - nDelColCount + 1, l, j - nDelColCount + 1
'                                        F1Book.EditPasteValues
                                    End If
                                Next l
                                '删除无用的列
                                .DeleteRange -1, j - nDelColCount, -1, j - nDelColCount, F1ShiftHorizontal
                                F1Book.MaxCol = F1Book.MaxCol - 1
                                
                                nDelColCount = nDelColCount + 1
                                
                                '将原有合并情况恢复
                                For l = 1 To nMergeCount
                                    F1Book.SetSelection atMergeCells(l).StartRow, atMergeCells(l).StartCol, atMergeCells(l).EndRow, atMergeCells(l).EndCol - 1
                                    Set oCellFormat = F1Book.GetCellFormat
                                    oCellFormat.MergeCells = True
                                    F1Book.SetCellFormat oCellFormat
                                Next l
                                GoTo LoopNext
                            Else
                                '替换
                                szDestinString = szDestinString & szTemp
                            End If
                            '
                    End Select
                    szRight = LeftAndRight(szRight, False, "]")
                    UnEncodeKeyValue szRight, szKey, szValue
                    szDestinString = szDestinString & LeftAndRight(szRight, True, "[")
                Wend
    '            szDestinString = szDestinString & szRight
    
                If szString <> szDestinString Then
                    If IsNumeric(szDestinString) Then
                        .NumberRC(i, j - nDelColCount) = CDbl(szDestinString)
                    Else
                        .TextRC(i, j - nDelColCount) = szDestinString
                    End If
                    If bNeedSetFont Then
                        .Row = i
                        .col = j - nDelColCount
                        .SetFont szFontName, nFontSize, False, False, False, False, vbBlack, False, False
                    End If
                    If mbAutoRowHeight Then
                        .RowHeight(i) = GetRowHeight(i)
                    End If
                End If
LoopNext:
            Next j
        Next i

    End With
End Sub
Private Sub ReadBodyDataFormat()
    Dim i As Integer, j As Integer
    Dim szString As String, szKey As String, szValue As String
    Dim lRows As Long, lCols As Long, nArrayLen As Integer
    
    '触发"开始填充自定义数据"事件
    RaiseEvent StartReadTemplateFormat
    
    lRows = mlEndFillLine - mlStartFillLine + 1: lCols = F1Book.LastCol - F1Book.MinCol + 1
    For i = F1Book.MinCol To F1Book.LastCol
        For j = mlStartFillLine To mlEndFillLine
            szString = F1Book.TextRC(j, i)
            If szString <> "" Then
                UnEncodeKeyValue szString, szKey, szValue
            End If
            If szKey = "合并项目" Or szKey = "项目" Or szKey = "序号" Then
                GoTo Found
            End If
        Next j
    Next i
    mlDataFormatLine = -1
    Exit Sub
Found:
    mlDataFormatLine = j
    ReDim maszColumnStr(1 To lCols)
    For j = 1 To lCols
        maszColumnStr(j) = F1Book.TextRC(mlDataFormatLine, j + F1Book.MinCol - 1)
    Next j
    For i = mlDataFormatLine + 1 To mlStartFillLine + lRows - 1
        For j = 0 To lCols - 1
            szString = F1Book.TextRC(i, j + F1Book.MinCol)
            UnEncodeKeyValue szString, szKey, szValue
            Select Case szKey
                Case "小计", "日小计", "月小计", "年小计", "平均小计"
                    nArrayLen = ArrayLength(matSortSum) + 1
                    If nArrayLen <> 1 Then
                        ReDim Preserve matSortSum(1 To nArrayLen)
                    Else
                        ReDim matSortSum(1 To nArrayLen)
                    End If
                    matSortSum(nArrayLen).col = j + F1Book.MinCol
                    matSortSum(nArrayLen).ToRow = i - mlDataFormatLine
                    matSortSum(nArrayLen).Key = szKey
                    matSortSum(nArrayLen).KeyColumn = szValue
                Case "合计", "平均合计"
                    nArrayLen = ArrayLength(matTotalSum) + 1
                    If nArrayLen <> 1 Then
                        ReDim Preserve matTotalSum(1 To nArrayLen)
                    Else
                        ReDim matTotalSum(1 To nArrayLen)
                    End If
                    matTotalSum(nArrayLen).col = j + F1Book.MinCol
                    matTotalSum(nArrayLen).ToRow = i - mlDataFormatLine
                    matTotalSum(nArrayLen).Key = szKey
                    matTotalSum(nArrayLen).KeyColumn = szValue
            End Select
        Next j
    Next i
    
End Sub
Private Sub FillBodyData(prsDataSource As Variant)
    Dim i As Long, j As Long
    Dim nCurrRow As Long    '当前操作行
    Dim alRowBegin() As Long, alSumCount() As Long   '每一小计的起始行
    Dim avLostValue() As Variant
    Dim bNeedSort As Boolean

    '触发"开始填充数据"事件
    RaiseEvent StartFillContentData
    
    '分段
    If prsDataSource Is Nothing Then GoTo LastRun       'Exit Sub
    If prsDataSource.RecordCount = 0 Then GoTo LastRun
    
    '触发"设置进度范围"事件
    RaiseEvent SetProgressRange(prsDataSource.RecordCount)
    
    Dim nSortSum As Integer
    nSortSum = ArrayLength(matSortSum)
    If nSortSum = 0 Then GoTo FillRecordset

    ReDim avLostValue(1 To nSortSum)
    ReDim alRowBegin(1 To nSortSum)
    ReDim alSumCount(1 To nSortSum)

    '记录初始行
    For i = 1 To nSortSum
        alRowBegin(i) = mlDataFormatLine
        alSumCount(i) = 0
    Next i

FillRecordset:
    nCurrRow = mlDataFormatLine
    mlDataStartLine = mlDataFormatLine

    '添加记录空行
    F1Book.MaxRow = F1Book.MaxRow + prsDataSource.RecordCount
    F1Book.InsertRange mlDataFormatLine, 1, mlDataFormatLine + prsDataSource.RecordCount - 1, 1, F1ShiftRows + F1FixupPrepend
'    F1Book.InsertRow mlDataFormatLine, prsDataSource.RecordCount, 0
    For i = mlDataFormatLine To mlDataFormatLine + prsDataSource.RecordCount - 1
        F1Book.RowHeight(i) = F1Book.RowHeight(mlDataFormatLine + prsDataSource.RecordCount)
    Next i
    mlDataFormatLine = mlDataFormatLine + prsDataSource.RecordCount
    '开始填充数据
    F1Book.SelStartRow = mlDataFormatLine: F1Book.SelEndRow = mlDataFormatLine
    F1Book.SelStartCol = F1Book.MinCol: F1Book.SelEndCol = F1Book.LastCol
    F1Book.EditCopy
'    F1Book.CopyRange 0, mlDataFormatLine, F1Book.GetCols(0) - 1, mlDataFormatLine
    prsDataSource.MoveFirst
    For i = 1 To nSortSum
        avLostValue(i) = prsDataSource(matSortSum(i).KeyColumn).Value         '取第一行字段
    Next i
    FillCurrRowData prsDataSource, nCurrRow
    nCurrRow = nCurrRow + 1

    Dim szSortColumn    As String
    For i = 2 To prsDataSource.RecordCount      '比较小计字段
        prsDataSource.MoveNext
        bNeedSort = False
        szSortColumn = ""
        For j = 1 To nSortSum
            Select Case matSortSum(j).Key
                Case "小计", "平均小计"
                    If avLostValue(j) <> prsDataSource(matSortSum(j).KeyColumn).Value Then
                        avLostValue(j) = prsDataSource(matSortSum(j).KeyColumn).Value
                        szSortColumn = szSortColumn & EncodeString(CStr(j))
                        bNeedSort = True
                    Else
                        alSumCount(j) = alSumCount(j) + 1
                    End If
                Case "日小计"
                    If Format(avLostValue(j), "YYYYMMDD") <> Format(prsDataSource(matSortSum(j).KeyColumn).Value, "YYYYMMDD") Then
                        avLostValue(j) = prsDataSource(matSortSum(j).KeyColumn).Value
                        szSortColumn = szSortColumn & EncodeString(CStr(j))
                        bNeedSort = True
                    Else
                        alSumCount(j) = alSumCount(j) + 1
                    End If
                Case "月小计"
                    If Format(avLostValue(j), "YYYYMM") <> Format(prsDataSource(matSortSum(j).KeyColumn).Value, "YYYYMM") Then
                        avLostValue(j) = prsDataSource(matSortSum(j).KeyColumn).Value
                        szSortColumn = szSortColumn & EncodeString(CStr(j))
                        bNeedSort = True
                    Else
                        alSumCount(j) = alSumCount(j) + 1
                    End If
                Case "年小计"
                    If Format(avLostValue(j), "YYYY") <> Format(prsDataSource(matSortSum(j).KeyColumn).Value, "YYYY") Then
                        avLostValue(j) = prsDataSource(matSortSum(j).KeyColumn).Value
                        szSortColumn = szSortColumn & EncodeString(CStr(j))
                        bNeedSort = True
                    Else
                        alSumCount(j) = alSumCount(j) + 1
                    End If
                Case Else
                    alSumCount(j) = alSumCount(j) + 1
            End Select
        Next j
        If bNeedSort Then   '无需要小计
            InsertSumLine nCurrRow, alRowBegin, alSumCount, szSortColumn
'            F1Book.CopyRange 0, mlDataFormatLine, F1Book.GetCols(0) - 1, mlDataFormatLine
            F1Book.SelStartRow = mlDataFormatLine: F1Book.SelEndRow = mlDataFormatLine
            F1Book.SelStartCol = F1Book.MinCol: F1Book.SelEndCol = F1Book.LastCol
            F1Book.EditCopy
        End If
        FillCurrRowData prsDataSource, nCurrRow

        MergeSameColumn prsDataSource, nCurrRow
        nCurrRow = nCurrRow + 1
    Next i

LastRun:
    '设置最后一项小计栏
    szSortColumn = ""
    For j = 1 To nSortSum
        szSortColumn = szSortColumn & EncodeString(CStr(j))
    Next j
    InsertSumLine nCurrRow, alRowBegin, alSumCount, szSortColumn
    '设置总计栏
    
    If prsDataSource Is Nothing Then
        SetTotalSumLine True
    Else
        SetTotalSumLine IIf(prsDataSource.RecordCount > 0, False, True)
    End If
End Sub
Private Sub MergeSameColumn(prsRecordset As Variant, plCurrRow As Long)
On Error Resume Next
    Dim i As Integer
    Dim szKey As String, szValue As String
    Dim szTmpKey As String, szTmpValue As String
    Dim vTmpField1 As Variant, vTmpField2 As Variant

    For i = 1 To ArrayLength(maszColumnStr)
        UnEncodeKeyValue maszColumnStr(i), szKey, szValue

        If szKey <> "合并项目" Then GoTo NoMerge
        If plCurrRow = mlDataStartLine Then GoTo NoMerge    '第一行
        If i = 1 Then GoTo Merge
'        UnEncodeKeyValue maszColumnStr(i - 1), szTmpKey, szTmpValue
'        If szTmpKey <> "合并项目" Then GoTo Merge
'
'
'        '左侧检查
'        prsRecordset.MovePrevious
'        vTmpField1 = prsRecordset(szTmpValue).Value
'
'        prsRecordset.MoveNext
'        vTmpField2 = prsRecordset(szTmpValue).Value
'        If vTmpField1 <> vTmpField2 Then GoTo NoMerge
        
        '左侧检查
        Dim j As Integer
        For j = i - 1 To 1 Step -1
            UnEncodeKeyValue maszColumnStr(j), szTmpKey, szTmpValue
            If szTmpKey <> "合并项目" Then GoTo Merge
                
            prsRecordset.MovePrevious
            vTmpField1 = prsRecordset(szTmpValue).Value
            prsRecordset.MoveNext
            vTmpField2 = prsRecordset(szTmpValue).Value
            If vTmpField1 <> vTmpField2 Then GoTo NoMerge
        Next j
        
Merge:
        prsRecordset.MovePrevious
        vTmpField1 = prsRecordset(szValue).Value
        prsRecordset.MoveNext
        vTmpField2 = prsRecordset(szValue).Value
        If vTmpField1 <> vTmpField2 Then GoTo NoMerge

        Dim lRow1 As Long, lCol1 As Long, lRow2 As Long, lCol2 As Long
        F1Book.TextRC(plCurrRow, F1Book.MinCol + i - 1) = ""

        F1Book.SetSelection plCurrRow - 1, F1Book.MinCol + i - 1, plCurrRow, F1Book.MinCol + i - 1
        Dim oF1CellFormat As F1CellFormat
        Set oF1CellFormat = F1Book.GetCellFormat
        oF1CellFormat.MergeCells = True
        F1Book.SetCellFormat oF1CellFormat
NoMerge:
        If Err Then Err.Clear
    Next i
End Sub

Private Sub DeleDataNoteLine()
    Dim nHeight As Integer
    If ArrayLength(matSortSum) > 0 Then
        nHeight = matSortSum(1).ToRow
    Else
        nHeight = 0
    End If
    F1Book.DeleteRange mlDataFormatLine, F1Book.MinCol, mlDataFormatLine + nHeight, F1Book.LastCol, F1ShiftRows
    F1Book.MaxRow = F1Book.MaxRow - nHeight - 1
End Sub
Private Sub InsertSumLine(nCurrRow As Long, alBeginSumRow() As Long, alSumCount() As Long, szSumColumns As String)
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim aszColumns() As String
    Dim szTmpFormulte As String, szOldFormulte As String, szTmpFormulte2 As String
    With F1Book
        If szSumColumns = "" Then Exit Sub
        aszColumns = SplitEncodeStringArray(szSumColumns)
        If ArrayLength(aszColumns) = 0 Then Exit Sub
        .SelStartRow = mlDataFormatLine + matSortSum(1).ToRow: .SelEndRow = .SelStartRow
'        .SelStartRow = mlDataFormatLine + matSortSum(1).ToRow + 1: .SelEndRow = .SelStartRow
        .SelStartCol = .MinCol: .SelEndCol = .LastCol
        .EditCopy
        
        '插入一行小计行
        .MaxRow = F1Book.MaxRow + 1
        .InsertRange nCurrRow, F1Book.MinCol, nCurrRow, F1Book.LastCol, F1ShiftRows + F1FixupPrepend
'        .SelStartRow = mlDataFormatLine + matSortSum(1).ToRow + 1: .SelEndRow = .SelStartRow
        .SelStartRow = nCurrRow: .SelEndRow = .SelStartRow
        .SelStartCol = .MinCol: .SelEndCol = .LastCol
        .EditPaste
        '将小计行中统计标识符内容清除
        For i = F1Book.MinCol To F1Book.LastCol
            If UnEncodeString(.TextRC(nCurrRow, i)) <> "" Then
                .TextRC(nCurrRow, i) = ""
            End If
        Next i
        
        
        mlDataFormatLine = mlDataFormatLine + 1
        Dim szFunctionName As String
        Dim nColumnIndex As Integer
        For i = 1 To ArrayLength(aszColumns)
           nColumnIndex = Val(UnEncodeString(aszColumns(i)))
            Select Case matSortSum(nColumnIndex).Key
                Case "平均小计"
                    szFunctionName = "AVERAGE"
                Case Else       '其他均为合计
                    szFunctionName = "SUM"
            End Select
            szTmpFormulte = szFunctionName & " (" & ConvertCellName(matSortSum(nColumnIndex).col - 1, alBeginSumRow(nColumnIndex)) & _
                             ":" & ConvertCellName(matSortSum(nColumnIndex).col - 1, alBeginSumRow(nColumnIndex) + alSumCount(nColumnIndex)) & ")"
            .FormulaRC(nCurrRow, matSortSum(nColumnIndex).col) = szTmpFormulte
             For j = 1 To ArrayLength(matTotalSum)
                 If matTotalSum(j).col = matSortSum(nColumnIndex).col Then      '对应的合计栏
                    Select Case matTotalSum(j).Key
                        Case "平均合计"
                            szFunctionName = "AVERAGE"
                        Case Else       '其他均为合计
                            szFunctionName = "SUM"
                    End Select
                    szOldFormulte = .FormulaRC(mlDataFormatLine + matTotalSum(j).ToRow, matTotalSum(j).col)
                    If szOldFormulte <> "" Then     '拼装总公式
                        szTmpFormulte = StrReverse(LeftAndRight(StrReverse(szOldFormulte), False, ")")) & "," & szTmpFormulte & ")"
                    Else
                        szTmpFormulte = szFunctionName & " (" & szTmpFormulte & ")"
                    End If
                    .FormulaRC(mlDataFormatLine + matTotalSum(j).ToRow, matTotalSum(j).col) = szTmpFormulte
                     Exit For
                 End If
             Next j
            alBeginSumRow(nColumnIndex) = nCurrRow + 1
            alSumCount(nColumnIndex) = -1
        Next i
        For i = 1 To ArrayLength(matSortSum)    '对于其他行需要补充一行
            alSumCount(i) = alSumCount(i) + 1
        Next i

    End With
    '重新定位
    nCurrRow = nCurrRow + 1
End Sub
Private Sub SetTotalSumLine(pbRSIsNull As Boolean)
    '补充无对应小计栏的合计栏
    Dim i As Integer
    Dim szFormulte As String, szFunctionName As String
    For i = 1 To ArrayLength(matTotalSum)
        Select Case matTotalSum(i).Key
            Case "平均合计"
                szFunctionName = "AVERAGE"
            Case Else       '其他均为合计
                szFunctionName = "SUM"
        End Select

        If F1Book.FormulaRC(mlDataFormatLine + matTotalSum(i).ToRow, matTotalSum(i).col) = "" Then
            If Not pbRSIsNull Then
                szFormulte = szFunctionName & "(" & ConvertCellName(matTotalSum(i).col - 1, mlDataStartLine) & ":" & ConvertCellName(matTotalSum(i).col - 1, mlDataFormatLine - 1) & ")"
                F1Book.FormulaRC(mlDataFormatLine + matTotalSum(i).ToRow, matTotalSum(i).col) = szFormulte
            Else
                F1Book.NumberRC(mlDataFormatLine + matTotalSum(i).ToRow, matTotalSum(i).col) = "0"
            End If
        End If
    Next i

End Sub
Private Sub FillCurrRowData(rsData As Variant, nCurrRow As Long)
    On Error Resume Next
    With F1Book
        ShowProgess rsData.AbsolutePosition, rsData.RecordCount
        
        F1Book.SelStartRow = nCurrRow: F1Book.SelEndRow = nCurrRow
        F1Book.SelStartCol = F1Book.MinCol: F1Book.SelEndCol = F1Book.LastCol
'        F1Book.SetSelection nCurrRow, nCurrRow, F1Book.MinCol, F1Book.LastCol
        .EditPaste

        Dim i As Integer
        Dim szKey As String, szValue As String, szFormat As String
        For i = 1 To ArrayLength(maszColumnStr)
            szFormat = ""
            
            UnEncodeKeyValue maszColumnStr(i), szKey, szValue
            '2003年7月22日加
            If szKey = "序号" Then      '如果是序号列，则添加序号
                .TextRC(nCurrRow, .MinCol + i - 1) = rsData.AbsolutePosition
            Else
                If Not (IsNull(rsData.Fields(szValue).Value) Or IsEmpty(rsData.Fields(szValue).Value)) Then
                    If IsNumeric(rsData.Fields(szValue).Value) And TypeName(rsData.Fields(szValue).Value) <> "String" Then
                        .NumberRC(nCurrRow, .MinCol + i - 1) = rsData.Fields(szValue)
                    Else
                        If rsData.Fields(szValue).Type = adDate Or rsData.Fields(szValue).Type = adDBDate Or rsData.Fields(szValue).Type = adDBTime Or rsData.Fields(szValue).Type = adDBTimeStamp Then         '日期型字段的处理
                            If CDate(rsData.Fields(szValue)) <= CDate("1900-1-1") Then       '空日期不显示
                                .TextRC(nCurrRow, .MinCol + i - 1) = ""
                            Else
                                .TextRC(nCurrRow, .MinCol + i - 1) = RTrim(rsData.Fields(szValue))
                            End If
                        Else
                            .TextRC(nCurrRow, .MinCol + i - 1) = RTrim(rsData.Fields(szValue))
                        End If
                    End If
                Else
                    If IsNumeric(rsData.Fields(szValue).Value) Then
                        .NumberRC(nCurrRow, .MinCol + i - 1) = 0
                    Else
                        .TextRC(nCurrRow, .MinCol + i - 1) = ""
                    End If
                End If
            End If
                
            If Err Then
                Err.Clear
'                .TextRC(nCurrRow, .MinCol + i - 1) = ""
            End If
        Next i
    
        If mbAutoRowHeight Then
            .RowHeight(nCurrRow) = GetRowHeight(nCurrRow)
        End If
    End With
    
    
    '触发进度事件
    RaiseEvent SetProgressValue(rsData.AbsolutePosition)
End Sub


Private Function GetRowHeight(plRowIndex As Long) As Long
   Const cnTopMargin = 100
   Dim dbTextLen As Double
   Dim nSplitLines As Integer
   Dim oCellFormat As F1CellFormat
   Dim lOldHeight As Long
   lOldHeight = F1Book.RowHeight(plRowIndex)
   Dim lNewHeight As Long, lNewTmp As Long
   '**********************************************
   Dim i As Integer
   Dim lMergeColWidth As Long, j As Integer
   Dim szTemp As String
   Dim nBeginCol As Integer, nEndCol As Integer '表示起始列和结束列
   nBeginCol = F1Book.MaxCol
   nEndCol = nBeginCol
  
   For i = F1Book.MaxCol To F1Book.MinCol Step -1
      With F1Book
          .SetActiveCell plRowIndex, i
         Set oCellFormat = .GetCellFormat
         If oCellFormat.WordWrap = True Then
            If .Text = "" And oCellFormat.MergeCells = True Then  '如果值是空又是合并项，则表示与前面是合并项
               nEndCol = i
            Else
                '算出第一列的字数宽度
                  dbTextLen = LenA(.TextRC(plRowIndex, nEndCol)) * ScaleX(oCellFormat.FontSize, vbPoints, vbTwips)
                  For j = nBeginCol To nEndCol Step -1
                      lMergeColWidth = lMergeColWidth + .ColWidth(j)
                  Next j
                  nSplitLines = Int(dbTextLen / lMergeColWidth) + 1
                  lNewTmp = ScaleY(oCellFormat.FontSize, vbPoints, vbTwips) * nSplitLines + 2 * cnTopMargin
                 '重新计算
                  nBeginCol = i
                  nEndCol = i
            End If
         Else
           nBeginCol = i
           nEndCol = i
         End If
       End With
       If lNewTmp > lOldHeight Then
          lOldHeight = lNewTmp
       End If
    Next i
    If nEndCol < nBeginCol Then
       '先算出本合并单元的高度
           With F1Book
                .SetActiveCell plRowIndex, nEndCol
                Set oCellFormat = .GetCellFormat
                 '算出第一列的字数宽度
               dbTextLen = LenA(.TextRC(plRowIndex, nEndCol)) * ScaleX(oCellFormat.FontSize, vbPoints, vbTwips)
               For j = nBeginCol To nEndCol Step -1
                   lMergeColWidth = lMergeColWidth + .ColWidth(j)
               Next j
               nSplitLines = Int(dbTextLen / lMergeColWidth) + 1
               lNewTmp = ScaleY(oCellFormat.FontSize, vbPoints, vbTwips) * nSplitLines + 2 * cnTopMargin
           End With
    End If
    If lNewTmp > lOldHeight Then
       lNewHeight = lNewTmp + cnTopMargin
    Else
       lNewHeight = lOldHeight
    End If
     
    '**********************************************
    GetRowHeight = lNewHeight
End Function
'导出
'##ModelId=3C5100FE0145
Public Sub ExportTo(ByVal pszFileName As String, ByVal pnExportType As Integer)
    On Error GoTo SaveToErr
    Dim szModule As String
    szModule = "RTReport->SaveTo"

    Dim lType As Long
    Select Case pnExportType
        Case HTML_TYPE
            lType = F1FileHTML
        Case TEXT_TYPE
            lType = F1FileTabbedText
        Case FORMULA_ONE_6_TYPE
            lType = F1FileFormulaOne6
        Case Else
            RaiseError ERR_FileTypeNotSupport
    End Select
    On Error Resume Next
    F1Book.WriteEx pszFileName, lType
    If Err Then
        On Error GoTo SaveToErr
        RaiseError ERR_FileSaveError, szModule
    End If
    
    Exit Sub
SaveToErr:
    Call RaiseError(Err.Number, szModule, Err.Description)
End Sub

'打开常用对话框
'##ModelId=3C51010A0200
Public Function OpenDialog(ByVal pnDialogType As Integer) As String
    On Error GoTo OpenDialogErr
    Dim szFileName As String
    Select Case pnDialogType
        Case EDialogType.EXPORT_FILE
            CommDialog.Filter = "Microsoft Excel5.0/95 File (*.xls)|*.xls|Microsoft Excel97 File (*.xls)|*.xls|HTML File (*.htm)|*.htm|Tabbed Text File (*.txt)|*.txt|Formula One 6 File (*.vts)|*.vts"
            CommDialog.FilterIndex = 2
            CommDialog.CancelError = False
            CommDialog.DialogTitle = "当前表格导出至..."
            CommDialog.ShowSave
            szFileName = CommDialog.FileName
            If Trim(szFileName) = "" Then Exit Function
            Select Case UCase(Right(szFileName, 4))
                Case ".VTS"
                    ExportTo szFileName, EExportFileType.FORMULA_ONE_6_TYPE
                Case ".HTM"
                    ExportTo szFileName, EExportFileType.HTML_TYPE
                Case ".TXT"
                    ExportTo szFileName, EExportFileType.TEXT_TYPE
                Case ".XLS"
                    Select Case CommDialog.FilterIndex
                        Case 1
                            SaveTo szFileName, EFileType.EXCEL_5_TYPE
                        Case 2
                            SaveTo szFileName, EFileType.EXCEL_97_TYPE
'                        Case 3
'                            SaveTo szFileName, EFileType.EXCEL_2000_TYPE
                    End Select
            End Select
            OpenDialog = szFileName
        Case EDialogType.SAVE_FILE
            CommDialog.DialogTitle = "当前表格保存为..."
            CommDialog.Filter = "Microsoft Excel5.0/95 File (*.xls)|*.xls|Microsoft Excel97 File (*.xls)|*.xls|Microsoft Excel2000 File (*.xls)|*.xls"
            CommDialog.FilterIndex = 2
            CommDialog.CancelError = False
            CommDialog.ShowSave
            szFileName = CommDialog.FileName
            If Trim(szFileName) = "" Then Exit Function
            Select Case CommDialog.FilterIndex
                Case 1
                    SaveTo szFileName, EFileType.EXCEL_5_TYPE
                Case 2
                    SaveTo szFileName, EFileType.EXCEL_97_TYPE
                Case 3
                    SaveTo szFileName, EFileType.EXCEL_2000_TYPE
            End Select
            OpenDialog = szFileName
        Case EDialogType.PRINT_TYPE
            F1Book.FilePrintSetupDlg
        Case EDialogType.PAGESET_TYPE
            F1Book.FilePageSetupDlg
    End Select

    Exit Function
OpenDialogErr:
    Call RaiseError(Err.Number, "RTReport->OpenDialog", Err.Decription)
End Function
'保存
'##ModelId=3C510338036B
Public Sub SaveTo(ByVal pszFileName As String, Optional pnFileType As Integer = EXCEL_5_TYPE)
    On Error GoTo SaveToErr
    Dim szModule As String
    szModule = "RTReport->SaveTo"

    Dim lType As Long
    Select Case pnFileType
        Case EXCEL_5_TYPE
            lType = F1FileExcel5
        Case EXCEL_97_TYPE
            lType = F1FileExcel97
        Case EXCEL_2000_TYPE
            'lType = F1FileExcel2000
        Case Else
            RaiseError ERR_FileTypeNotSupport
    End Select
    On Error Resume Next
    F1Book.WriteEx pszFileName, lType
    If Err Then
        On Error GoTo SaveToErr
        RaiseError ERR_FileSaveError, szModule
    End If
    
    Exit Sub
SaveToErr:
    Call RaiseError(Err.Number, szModule, Err.Description)
End Sub

'打印预览
'##ModelId=3C5332B9035F
Public Sub PrintView()
    On Error GoTo PrintViewErr

    F1Book.FilePrintPreview

    Exit Sub
PrintViewErr:
    Call RaiseError(Err.Number, "RTReport->PrintView", Err.Description)
End Sub



'##ModelId=3C5338F30267
Public Property Get CellObject() As Object
    Set CellObject = F1Book
End Property

''##ModelId=3C5338F30177
'Public Property Get BackgroupFile() As String
'   Let BackgroupFile = mszBackgroupFile
'End Property
'
''##ModelId=3C5338F203E2
'Public Property Let BackgroupFile(ByVal pszValue As String)
'On Error Resume Next
'    Let mszBackgroupFile = pszValue
''    pszValue = Trim(pszValue)
''    If pszValue = "" Then F1Book.SetBackImage -1, 1, 0
''    F1Book.DeleteImage 1
''    F1Book.AddImage pszValue
''    F1Book.SetBackImage 1, 2, 0
''    F1Book.ReDraw
'End Property

'##ModelId=3C5338F202FC
Public Property Get TemplateFile() As String
   Let TemplateFile = mszTemplateFile
End Property

'显示列标
Public Property Let TopLabelVisual(ByVal pbValue As Boolean)
'    F1Book.ShowTopLabel IIf(pbValue, 1, 0), 0
    mbTopLabelVisual = pbValue
End Property
Public Property Get TopLabelVisual() As Boolean
   TopLabelVisual = mbTopLabelVisual
End Property
'显示行标
Public Property Let LeftLabelVisual(ByVal pbValue As Boolean)
'    F1Book.ShowSideLabel IIf(pbValue, 1, 0), 0
    mbLeftLabelVisual = pbValue
End Property
Public Property Get LeftLabelVisual() As Boolean
   LeftLabelVisual = mbLeftLabelVisual
End Property
'显示行标
Public Property Let SheetTitle(ByVal pszTitle As String)
'    F1Book.ShowSideLabel IIf(pbValue, 1, 0), 0
    mszSheetTitle = pszTitle
End Property
Public Property Get SheetTitle() As String
   SheetTitle = mszSheetTitle
End Property

'##ModelId=3C5338F2016B
Public Property Let TemplateFile(ByVal Value As String)
    Let mszTemplateFile = Value
End Property


Private Sub F1Book_Click(ByVal nRow As Long, ByVal nCol As Long)
    RaiseEvent Click(nRow, nCol)
End Sub

Private Sub F1Book_DblClick(ByVal nRow As Long, ByVal nCol As Long)
    RaiseEvent DblClick(nRow, nCol)
End Sub

Private Sub F1Book_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub F1Book_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub F1Book_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub F1Book_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub F1Book_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub F1Book_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub


Private Sub F1Book_RClick(ByVal nRow As Long, ByVal nCol As Long)
    RaiseEvent RClick(nRow, nCol)
End Sub

Private Sub F1Book_RDblClick(ByVal nRow As Long, ByVal nCol As Long)
    RaiseEvent RDblClick(nRow, nCol)
End Sub



Private Sub UserControl_Initialize()
    Debug.Print "UserControl_Initialize"

    '改观CELL组件
    F1Book.ShowColHeading = False
    F1Book.ShowRowHeading = False
    F1Book.ShowTabs = F1TabsOff
'    F1Book.ShowSheetLabel 0, 0
End Sub

Private Sub UserControl_InitProperties()
    Debug.Print "UserControl_InitProperties"
    
    mszTemplateFile = ""
    mszBackgroupFile = ""
    TopLabelVisual = True
    LeftLabelVisual = True
    mbAutoRowHeight = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Debug.Print "UserControl_ReadProperties"
End Sub

Private Sub UserControl_Resize()
    F1Book.Top = 0
    F1Book.Left = 0
    F1Book.Width = UserControl.Width
    F1Book.Height = UserControl.Height
End Sub

Private Sub UserControl_Terminate()
    Debug.Print "UserControl_Terminate"

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Debug.Print "UserControl_WriteProperties"
End Sub
Private Function ConvertCellName(plCol As Long, plRow As Long) As String
'    ConvertCellName = IIf(Int(plCol / 26) > 0, Chr(Asc("A") + Int(plCol / 26)), "")
'    ConvertCellName = ConvertCellName & Chr(Asc("A") + (plCol Mod 26))
'    ConvertCellName = ConvertCellName & plRow
    ConvertCellName = IIf(Int((plCol) / 26) > 0, Chr(Asc("A") + Int((plCol) / 26) - 1), "")
    ConvertCellName = ConvertCellName & Chr(Asc("A") + ((plCol) Mod 26))
    ConvertCellName = ConvertCellName & plRow
End Function

Public Property Let CustomString(ByVal arsTemp As Variant)
    m_arsCustomValue = arsTemp
End Property

Public Property Let CustomStringCount(aszTemp As Variant)
    m_aszCustomString = aszTemp
    m_nCustomCount = ArrayLength(m_aszCustomString)
End Property

Private Function GetItemName(szKey As String, szValue As String) As String
    Dim k As Integer
    Dim l As Integer
    Dim nCount As Integer
    For k = 1 To m_nCustomCount
        If szKey = m_aszCustomString(k) Then
            m_arsCustomValue(k).MoveFirst
            nCount = m_arsCustomValue(k).RecordCount
            For l = 1 To nCount
                If Val(szValue) = Val(m_arsCustomValue(k).Fields(0).Value) Then
                    '如果找到,则进行替换
                    '跳出循环
                    GetItemName = Trim(m_arsCustomValue(k).Fields(1).Value)
                    Exit Function
                End If
                m_arsCustomValue(k).MoveNext
            Next l
            '如果未找到
            If l > nCount Then GetItemName = ""
            Exit Function
        End If
    Next k
    GetItemName = cszNotFound
    
End Function

Public Property Get Enabled() As Boolean
    Enabled = F1Book.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    F1Book.Enabled = vNewValue
End Property

Public Property Get AutoRowHeight() As Boolean
    AutoRowHeight = mbAutoRowHeight
End Property

Public Property Let AutoRowHeight(ByVal vNewValue As Boolean)
    mbAutoRowHeight = vNewValue
End Property
Public Function ZoomReport(Optional ByVal pbFitPage As Boolean = False, Optional ByVal psgZoomRatio As Single = 100) As Single
'pbFitPage是否按页面调整系数
    Dim dbPageWidth As Single, dbPageHeight As Single
    Dim i As Long
    If pbFitPage Then
        For i = F1Book.MinCol To F1Book.LastCol
            dbPageWidth = dbPageWidth + F1Book.ColWidth(i)
        Next i
        For i = F1Book.MinRow To F1Book.LastRow
            dbPageHeight = dbPageHeight + F1Book.RowHeight(i)
        Next i
        Dim sgTmpWidthScale As Single
        Dim sgTmpHeightScale As Single
        sgTmpHeightScale = F1Book.Width / dbPageWidth
        sgTmpHeightScale = F1Book.Height / dbPageHeight
        psgZoomRatio = IIf(sgTmpHeightScale > sgTmpWidthScale, sgTmpHeightScale, sgTmpWidthScale) * 100
    End If
    psgZoomRatio = Round(psgZoomRatio, 2)
    If psgZoomRatio >= 10 And psgZoomRatio <= 400 Then
        F1Book.ViewScale = psgZoomRatio
    Else
        psgZoomRatio = F1Book.ViewScale
    End If
    ZoomReport = psgZoomRatio
End Function
