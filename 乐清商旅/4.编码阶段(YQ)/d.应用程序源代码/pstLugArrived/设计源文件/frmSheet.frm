VERSION 5.00
Object = "{A5E8F770-DA22-4EAF-B7BE-73B06021D09F}#1.1#0"; "ST6Report.ocx"
Begin VB.Form frmSheet 
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   3660
   StartUpPosition =   3  '窗口缺省
   Begin ST6Report.RTReport RTReport1 
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   480
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

Private moSheetData As Package   '票据记录集
Private maszSheetCustom() As String '票据中的自定义数据



'填充路单报表
Private Sub FillSheetReport()
On Error GoTo ErrHandler
    
    ReDim maszSheetCustom(1 To 15, 1 To 2)
    Dim mrsSheetData As Recordset
    
    '构建自定义项目
    maszSheetCustom(1, 1) = "起运站"
    maszSheetCustom(1, 2) = IIf(moSheetData.StartStationName <> "", Trim(moSheetData.StartStationName), Trim(moSheetData.AreaType))
    maszSheetCustom(2, 1) = "品名"
    maszSheetCustom(2, 2) = Trim(moSheetData.PackageName)
    maszSheetCustom(3, 1) = "件数"
    maszSheetCustom(3, 2) = Trim(moSheetData.PackageNumber)
    maszSheetCustom(4, 1) = "计重"
    maszSheetCustom(4, 2) = Trim(moSheetData.CalWeight)
    maszSheetCustom(5, 1) = "收货单位"
    maszSheetCustom(5, 2) = Trim(moSheetData.Picker)
    maszSheetCustom(6, 1) = "货号"
    maszSheetCustom(6, 2) = Trim(moSheetData.PackageID)
    maszSheetCustom(7, 1) = "费用项1"
    maszSheetCustom(7, 2) = Trim(moSheetData.KeepCharge)
    maszSheetCustom(8, 1) = "费用项2"
    maszSheetCustom(8, 2) = Trim(moSheetData.LoadCharge)
    maszSheetCustom(9, 1) = "费用项3"
    maszSheetCustom(9, 2) = Trim(moSheetData.MoveCharge)
    maszSheetCustom(10, 1) = "代收运费"
    maszSheetCustom(10, 2) = Trim(moSheetData.TransitCharge)
    maszSheetCustom(11, 1) = "合计(小写)"
    maszSheetCustom(11, 2) = Trim(moSheetData.LoadCharge + moSheetData.KeepCharge + moSheetData.MoveCharge + moSheetData.TransitCharge)
    maszSheetCustom(12, 1) = "合计(大写)"
    maszSheetCustom(12, 2) = GetNumber(moSheetData.LoadCharge + moSheetData.KeepCharge + moSheetData.MoveCharge + moSheetData.TransitCharge)
    maszSheetCustom(13, 1) = "工号"
    maszSheetCustom(13, 2) = Trim(moSheetData.UserID)
    maszSheetCustom(14, 1) = "提货时间"
    maszSheetCustom(14, 2) = Format(moSheetData.PickTime, "YYYY-MM-DD HH:mm")
    maszSheetCustom(15, 1) = "提货证件"
    If moSheetData.PickerCreditID <> "" Then
        maszSheetCustom(15, 2) = Left(moSheetData.PickerCreditID, Len(moSheetData.PickerCreditID) - 4) & "****"
    Else
        maszSheetCustom(15, 2) = ""
    End If
    
    RTReport1.TemplateFile = App.Path & "\行包提货单.cll"
    RTReport1.ShowReport mrsSheetData, maszSheetCustom
    
     
    Exit Sub
ErrHandler:
    ShowErrorMsg
End Sub

'打印票据
Public Sub PrintSheetReport(ByVal poSheetData As Package)
    Set moSheetData = poSheetData
    FillSheetReport

    On Error Resume Next
    RTReport1.PrintReport
    
End Sub

