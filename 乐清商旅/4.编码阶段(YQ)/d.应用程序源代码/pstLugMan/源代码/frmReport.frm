VERSION 5.00
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#3.1#0"; "RTReportLF.ocx"
Begin VB.Form frmReport 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   3210
   ClientTop       =   4935
   ClientWidth     =   4920
   HelpContextID   =   7000370
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   4920
   WindowState     =   2  'Maximized
   Begin RTReportLF.RTReport flReport 
      Height          =   1440
      Left            =   795
      TabIndex        =   0
      Top             =   495
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   2540
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'Last Modify By: 陆勇庆  2005-8-16
'Last Modify In: ShowReport3的接口进行统一
'*******************************************************************************
Option Explicit
Public m_szCaption As String
Public m_lHelpContextID As Long
Private m_bNeedSave As Boolean
Private m_nReportType As Integer
Private m_szFileName As String '保存的文件名
Private m_nNeedMergeCol As Integer   '需要合并列
Private mlMaxCount As Long



Private Sub flReport_SetProgressRange(ByVal lRange As Variant)
    mlMaxCount = lRange
End Sub

Private Sub flReport_SetProgressValue(ByVal lValue As Variant)
    WriteProcessBar True, lValue, mlMaxCount
End Sub

Private Sub Form_Activate()
    mdiMain.SetPrintEnabled True
    mdiMain.lblTitle = Me.Caption
End Sub

Private Sub Form_Deactivate()
    mdiMain.SetPrintEnabled False
    mdiMain.lblTitle = ""
End Sub

Private Sub Form_Load()
    Me.HelpContextID = m_lHelpContextID
    
    Me.Caption = m_szCaption
'    If GetReportFormCount() <= 1 Then  True
    
End Sub



Private Sub Form_Resize()
    flReport.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub


Public Function ShowReport(prsData As Recordset, pszFileName As String, pszCaption As String, Optional pvaCustomData As Variant, Optional pnReportType As Integer = 0) As Long
    On Error GoTo Error_Handle
    Me.ZOrder 0
    Me.Show
    
    Dim arsTemp As Variant
    Dim aszTemp As Variant
'    Dim rsTemp As Recordset
    ReDim aszTemp(1 To 1)
    ReDim arsTemp(1 To 1)
    '赋票种
    aszTemp(1) = "托运费项"
    Set arsTemp(1) = g_rsPriceItem
    m_bNeedSave = True
    
    m_nReportType = pnReportType
    Me.Caption = pszCaption
    
    WriteProcessBar True, , , "正在形成报表..."
    
    flReport.CustomStringCount = aszTemp
    flReport.CustomString = arsTemp
    flReport.SheetTitle = ""
    
    flReport.TemplateFile = App.Path & "\" & pszFileName
    flReport.LeftLabelVisual = True
    flReport.TopLabelVisual = True
    flReport.ShowReport prsData, pvaCustomData
    WriteProcessBar False, , , ""
    ShowSBInfo "共" & prsData.RecordCount & "条记录", ESB_ResultCountInfo
    Exit Function
Error_Handle:
    ShowErrorMsg
End Function



Private Function GetReportFormCount()
    Dim frmTemp As Form
    Dim nCount As Integer
    nCount = 0
    For Each frmTemp In Forms
        If TypeName(frmTemp) = Me.name Then
            nCount = nCount + 1
        End If
    Next
    GetReportFormCount = nCount
End Function


Public Sub ExportToFile()

    flReport.OpenDialog EDialogType.EXPORT_FILE
End Sub


Public Function ShowReport2(aprsData() As Recordset, pszFileName As String, pszCaption As String, Optional pvaCustomData As Variant, Optional pnReportType As Integer = 0) As Long
    On Error GoTo Error_Handle
'    Me.Show
    
    Dim arsTemp As Variant
    Dim aszTemp As Variant
'    Dim rsTemp As Recordset
    ReDim aszTemp(1 To 1)
    ReDim arsTemp(1 To 1)
    '赋票种
    aszTemp(1) = "托运费项"
    Set arsTemp(1) = g_rsPriceItem
    m_bNeedSave = True
    m_nReportType = pnReportType
    Me.Caption = pszCaption
    
    WriteProcessBar True, , , "正在形成报表..."
    
    flReport.CustomStringCount = aszTemp
    flReport.CustomString = arsTemp
    flReport.LeftLabelVisual = True
    flReport.TopLabelVisual = True
    flReport.TemplateFile = App.Path & "\" & pszFileName
    flReport.ShowMultiReport aprsData, pvaCustomData
    WriteProcessBar True, , , ""
    Me.Show
    Exit Function
Error_Handle:
    ShowErrorMsg
End Function
Public Function ShowReport3(aprsData As Recordset, pszFileName As String, pszCaption As String, Optional pvaCustomData As Variant, Optional pnReportType As Integer = 0) As Long
     On Error GoTo Error_Handle
    Me.Show
    
    Dim arsTemp As Variant
    Dim aszTemp As Variant
'    Dim rsTemp As Recordset
    ReDim aszTemp(1 To 1)
    ReDim arsTemp(1 To 1)
    '赋票种
    aszTemp(1) = "托运费项"
    Set arsTemp(1) = g_rsPriceItem
    m_bNeedSave = True
    
    m_nReportType = pnReportType
    Me.Caption = pszCaption
    
    WriteProcessBar True, , , "正在形成报表..."
    
    flReport.CustomStringCount = aszTemp
    flReport.CustomString = arsTemp
    flReport.SheetTitle = ""
    
    flReport.TemplateFile = App.Path & "\" & pszFileName
    flReport.LeftLabelVisual = True
    flReport.TopLabelVisual = True
    flReport.ShowReport aprsData, pvaCustomData
    WriteProcessBar False, , , ""

    Exit Function
Error_Handle:
    ShowErrorMsg
End Function

Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    On Error GoTo ErrorHandle
    flReport.PrintReport pbShowDialog
    Exit Sub
ErrorHandle:
End Sub

Public Sub PreView()
    flReport.PrintView
End Sub

Public Sub PageSet()
    flReport.OpenDialog EDialogType.PAGESET_TYPE
End Sub

Public Sub PrintSet()
    flReport.OpenDialog EDialogType.PRINT_TYPE
End Sub
'导出文件
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = flReport.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'导出文件并打开
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = flReport.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mdiMain.SetPrintEnabled False
    mdiMain.lblTitle = ""
End Sub
