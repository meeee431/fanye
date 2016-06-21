VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clMantissa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:clMantissa.cls
'* Project Name:RTBusMan
'* Engineer:陈峰
'* Data Generated:2002/09/05
'* Last Revision Date:2002/09/13
'* Brief Description:专门作尾数处理用的类
'* Relational Document:
'**********************************************************

Private m_oPriceTable As RoutePriceTable '票价表对象
Private m_oF1Book As F1Book '对应的输入对象

Private m_rsPrice As Recordset '此处存放的是F1Book中所有打开的数据
Private m_rsUseItem As Recordset '传入可用票价项的记录集 , 为了找到该列对应的票价项的ID

Private m_nPriceItemStartCol As Integer '票价项起始列
Private m_nPriceItemStartRow As Integer '有数据的起始行
Private m_nTotalCol As Integer '总计所在的列
Private m_nMaxCol As Integer '舍入项(即进位项)所在列

Public Property Get PriceRs() As Recordset
    Set PriceRs = m_rsPrice
End Property

Public Property Let PriceRs(ByVal vNewValue As Recordset)
    Set m_rsPrice = vNewValue
End Property


Public Property Get PriceItemStartCol() As Integer
    PriceItemStartCol = m_nPriceItemStartCol
End Property

Public Property Let PriceItemStartCol(ByVal vNewValue As Integer)
    m_nPriceItemStartCol = vNewValue
End Property

Public Property Get PriceItemStartRow() As Integer
    PriceItemStartRow = m_nPriceItemStartRow
End Property

Public Property Let PriceItemStartRow(ByVal vNewValue As Integer)
    m_nPriceItemStartRow = vNewValue
End Property

Public Property Get TotalCol() As Integer
    TotalCol = m_nTotalCol
End Property

Public Property Let TotalCol(ByVal vNewValue As Integer)
    m_nTotalCol = vNewValue
End Property

Public Property Get MaxCol() As Integer
    MaxCol = m_nMaxCol
End Property

Public Property Let MaxCol(ByVal vNewValue As Integer)
    m_nMaxCol = vNewValue
End Property

Public Property Get oF1Book() As F1Book
    Set oF1Book = m_oF1Book
End Property

Public Property Let oF1Book(ByVal vNewValue As F1Book)
    Set m_oF1Book = vNewValue
End Property

Public Property Get oPriceTable() As RoutePriceTable
    Set oPriceTable = m_oPriceTable
End Property

Public Property Let oPriceTable(ByVal vNewValue As RoutePriceTable)
    Set m_oPriceTable = vNewValue
End Property

Public Property Get UseItemRs() As Recordset
    Set UseItemRs = m_rsUseItem
End Property

Public Property Let UseItemRs(ByVal vNewValue As Recordset)
    Set m_rsUseItem = vNewValue
End Property

'*****************************************************************************************************'
'描述:进行尾数处理,将相应的网格内的数据进行修改,但现在不处理传入的记录集数组的值的改变
'参数说明: pnStartRow 行,  pnEndRow 结束行,  ProgressVisible 进度条可视
'          pnDealCol: 如果是0  处理所有的列 , 否则处理对应的列  为0时主要在需要全体处理时用到
'*****************************************************************************************************
Public Sub SetTailCarry(pnStartRow As Long, pnEndRow As Long, Optional pnDealCol As Long = 0, Optional ProgressVisible As Boolean = False)
    Dim i As Integer
    Dim j As Long
    Dim vaData As Variant
    Dim sgTotalPrice As Single '总价值
    Dim ttDealValue As TDealValue
    Dim sgUpRiseValue As Single '舍入项值
    Dim sgTemp As Single
    Dim szPriceItem As String '票价项
    Dim szBusType As String '车次类型
    Dim szAreaCode As String '地区代码
    Dim nTemp As Integer
'    On Error GoTo ErrorHandle
    With m_oF1Book
        If ProgressVisible Then
            WriteProcessBar True, 1, pnEndRow - pnStartRow + 1, "正在进行尾数处理..."
        End If
        SetBusy
        
        For i = pnStartRow To pnEndRow
            If i = pnStartRow Then
                m_rsPrice.Move pnStartRow - m_nPriceItemStartRow, adBookmarkFirst
            Else
                m_rsPrice.MoveNext
            End If
            '得到车次类型
            szBusType = GetBusType
            szAreaCode = GetArea
            '取得舍入项的值
            vaData = Format(.NumberRC(i, m_nMaxCol), "0.00")  '"0.00" '
            sgUpRiseValue = Format(vaData, "0.00")
            sgTotalPrice = 0
            If pnDealCol = 0 Then
                '处理所有列
                For j = m_nPriceItemStartCol To .MaxCol
                    '从起始列到最后一列
                    vaData = Format(.NumberRC(i, j), "0.00")
                    '得到票价项的列名
                    szPriceItem = GetPriceItemID(j)
                    sgTemp = vaData
                    If szAreaCode <> "" Then ttDealValue = m_oPriceTable.GetColCarryValue(sgTemp, szAreaCode, szPriceItem, sgUpRiseValue, False, szBusType)
                    '得到处理后的值
                    If j = m_nMaxCol Then
                        sgUpRiseValue = ttDealValue.sgValue
                    End If
                    .NumberRC(i, j) = Format(ttDealValue.sgValue, "0.00")
                    sgTotalPrice = sgTotalPrice + ttDealValue.sgValue
                Next j
            Else
                '处理一列
                vaData = Format(.NumberRC(i, pnDealCol), "0.00")
                '得到票价项的列名
                szPriceItem = GetPriceItemID(pnDealCol)
                sgTemp = vaData
                If szAreaCode <> "" Then ttDealValue = m_oPriceTable.GetColCarryValue(sgTemp, szAreaCode, szPriceItem, sgUpRiseValue, False, szBusType)
                '得到处理后的值
                If pnDealCol = m_nMaxCol Then
                    sgUpRiseValue = ttDealValue.sgValue
                End If
                .NumberRC(i, pnDealCol) = Format(ttDealValue.sgValue, "0.00")
                sgTotalPrice = sgTotalPrice + ttDealValue.sgValue
                
                For j = m_nPriceItemStartCol To .MaxCol
                    If pnDealCol <> j Then
                        vaData = Format(.NumberRC(i, j), "0.00")
                        sgTotalPrice = sgTotalPrice + vaData
                    End If
                Next j
            End If
            '处理总票价
            If szAreaCode <> "" Then ttDealValue = m_oPriceTable.GetColCarryValue(sgTotalPrice, szAreaCode, cszItemBaseCarriage, sgUpRiseValue, True, szBusType)
'            Format(.numberrc(i, m_nTotalCol), "0.00") = ttDealValue.sgValue
            .NumberRC(i, m_nMaxCol) = Format(ttDealValue.sgLastColValue, "0.00")
            If ProgressVisible Then WriteProcessBar , i - pnStartRow, pnEndRow - pnStartRow + 1, "正在进行尾数处理..."
            
        Next i
        '*****刷新
        If ProgressVisible Then WriteProcessBar False
        SetNormal
    End With
    Exit Sub
ErrorHandle:
    WriteProcessBar False
    SetNormal
End Sub


Private Function GetBusType() As String
    '从记录集中得到车型
    If FormatDbValue(m_rsPrice!bus_type) = "" Then
        GetBusType = cnAllBusType
    Else
        GetBusType = FormatDbValue(m_rsPrice!bus_type)
    End If
End Function

Private Function GetArea() As String
    '从记录集 中得到地区
    GetArea = FormatDbValue(m_rsPrice!area_code)
    
End Function

Private Function GetPriceItemID(pnCol As Long) As String
    '从记录集中得到  票价项的代码
    '记录集移到对应的位置
    m_rsUseItem.Move pnCol - m_nPriceItemStartCol, 1
    GetPriceItemID = FormatDbValue(m_rsUseItem!price_item)
    
End Function

Private Sub Class_Terminate()
    '释放对象
    Set m_oF1Book = Nothing
    Set m_oPriceTable = Nothing
    Set m_rsPrice = Nothing
    Set m_rsUseItem = Nothing
End Sub
