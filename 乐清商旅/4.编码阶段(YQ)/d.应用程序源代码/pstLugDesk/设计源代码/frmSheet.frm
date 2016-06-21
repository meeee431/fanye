VERSION 5.00
Object = "{A5E8F770-DA22-4EAF-B7BE-73B06021D09F}#1.1#0"; "ST6Report.ocx"
Begin VB.Form frmSheet 
   Caption         =   "Form1"
   ClientHeight    =   1035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   2655
   StartUpPosition =   3  '窗口缺省
   Begin ST6Report.RTReport RTReport1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "frmSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moSheetData As AcceptSheet   '票据记录集
Private mszBusStartTime As String   '发车时间
Private maszSheetCustom() As String '票据中的自定义数据



'填充路单报表
Private Sub FillSheetReport()
On Error GoTo ErrHandler
    
    ReDim maszSheetCustom(1 To 16, 1 To 2)
    Dim mrsSheetData As Recordset
    Dim atPrice() As TLuggagePriceItem
    
    atPrice = moSheetData.PriceItems
    
    '构建自定义项目
    maszSheetCustom(1, 1) = "起运站"
    maszSheetCustom(1, 2) = Trim(moSheetData.DesStationName)
    maszSheetCustom(2, 1) = "品名"
    maszSheetCustom(2, 2) = Trim(moSheetData.LuggageName)
    maszSheetCustom(3, 1) = "件数"
    maszSheetCustom(3, 2) = Trim(moSheetData.Number)
    maszSheetCustom(4, 1) = "计重"
    maszSheetCustom(4, 2) = Trim(moSheetData.CalWeight)
    maszSheetCustom(5, 1) = "收货单位"
    maszSheetCustom(5, 2) = Trim(moSheetData.Picker)
    maszSheetCustom(6, 1) = "货号"
    maszSheetCustom(6, 2) = Trim(moSheetData.SheetID)
    maszSheetCustom(7, 1) = "费用项1"
    maszSheetCustom(7, 2) = Trim(atPrice(1).PriceValue)
    maszSheetCustom(8, 1) = "费用项2"
    maszSheetCustom(8, 2) = Trim(atPrice(3).PriceValue)
    maszSheetCustom(9, 1) = "费用项3"
    maszSheetCustom(9, 2) = Trim(atPrice(2).PriceValue) & "其它费" & atPrice(4).PriceValue
    maszSheetCustom(10, 1) = "代收运费"
    maszSheetCustom(10, 2) = Trim(0)
    maszSheetCustom(11, 1) = "合计(小写)"
    maszSheetCustom(11, 2) = Trim(atPrice(1).PriceValue + atPrice(2).PriceValue + atPrice(3).PriceValue + atPrice(4).PriceValue)
    maszSheetCustom(12, 1) = "合计(大写)"
    maszSheetCustom(12, 2) = GetNumber(atPrice(1).PriceValue + atPrice(2).PriceValue + atPrice(3).PriceValue + atPrice(4).PriceValue)
    maszSheetCustom(13, 1) = "工号"
    maszSheetCustom(13, 2) = Trim(m_oAUser.UserID)
    maszSheetCustom(14, 1) = "提货时间"
    maszSheetCustom(14, 2) = Format(moSheetData.OperateTime, "YYYY-MM-DD HH:mm")
    maszSheetCustom(15, 1) = "提货证件" '(托运人电话)
    maszSheetCustom(15, 2) = moSheetData.LuggageShipperPhone
    maszSheetCustom(16, 1) = "发车时间"
    maszSheetCustom(16, 2) = mszBusStartTime
    
    
    RTReport1.TemplateFile = App.Path & "\行包受理单.cll"
    RTReport1.ShowReport mrsSheetData, maszSheetCustom
    
     
    Exit Sub
ErrHandler:
    ShowErrorMsg
End Sub

'打印票据
Public Sub PrintSheetReport(ByVal poSheetData As AcceptSheet, ByVal pszBusStartTime As String)
    Set moSheetData = poSheetData
    mszBusStartTime = pszBusStartTime
    
    FillSheetReport

    On Error Resume Next
    RTReport1.PrintReport
    
End Sub


