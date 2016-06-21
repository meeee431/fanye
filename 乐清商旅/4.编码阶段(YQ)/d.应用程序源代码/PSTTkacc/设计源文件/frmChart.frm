VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.0#0"; "MSOWC.DLL"
Begin VB.Form frmChart 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   1215
   ClientTop       =   2670
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   9540
   WindowState     =   2  'Maximized
   Begin VB.Frame fraTop 
      Height          =   615
      Left            =   5190
      TabIndex        =   1
      Top             =   1260
      Width           =   2835
      Begin VB.ComboBox cboChartType 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "显示方式:"
         Height          =   180
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   810
      End
   End
   Begin OWC.ChartSpace ChartSpace1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7935
      XMLData         =   $"frmChart.frx":0000
      ScreenUpdating  =   -1  'True
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mnChartCount As Integer
Dim mszTitle As String

Public Property Get Title() As String
    Title = mszTitle
End Property

Public Property Let Title(ByVal vNewValue As String)
    mszTitle = vNewValue
End Property
Public Sub ClearChart()
    ChartSpace1.Clear
End Sub
Public Sub ShowChart(Optional ByVal pszChartTitle As String)
    Me.Caption = pszChartTitle
    If pszChartTitle <> "" Then
        ChartSpace1.HasChartSpaceTitle = True
        ChartSpace1.ChartSpaceTitle.Caption = pszChartTitle
        ChartSpace1.ChartSpaceTitle.Font.Size = 16
        ChartSpace1.ChartSpaceTitle.Font.Color = vbRed
    End If
    Me.Show
    Me.ZOrder
End Sub

Public Sub AddChart(ByVal pszTitle As String, ByVal prsDataSource As Recordset)
    Dim nDataSourceIndex As Integer
    Dim i, nCount As Integer
    Dim owcSource As WCDataSource
    nCount = prsDataSource.RecordCount
    Set owcSource = ChartSpace1.ChartDataSources.Add
    owcSource.DataSourceType = chDataSourceTypeRecordset
    owcSource.DataSource = prsDataSource
    nDataSourceIndex = ChartSpace1.ChartDataSources.Count - 1
    
'    If ChartSpace1.Charts.Count > 1 Then
'        ChartSpace1.Border.Color = vbGreen
'    End If
    Dim oChart As WCChart
    Set oChart = ChartSpace1.Charts.Add()
    If pszTitle <> "" Then
        oChart.HasTitle = True
        oChart.Title.Caption = pszTitle
    Else
        oChart.HasTitle = False
    End If
    If prsDataSource.Fields.Count >= 2 Then
        oChart.HasLegend = True
    Else
        oChart.HasLegend = False
    End If
    
    If prsDataSource.Fields.Count = 2 Then
        oChart.SetData chDimCategories, nDataSourceIndex, prsDataSource.Fields(0).name
        oChart.SetData chDimValues, nDataSourceIndex, prsDataSource.Fields(1).name
        oChart.SeriesCollection(0).Caption = prsDataSource.Fields(1).name
    End If
    
    If prsDataSource.Fields.Count = 3 Then
        oChart.SetData chDimSeriesNames, nDataSourceIndex, prsDataSource.Fields(1).name
        oChart.SetData chDimCategories, nDataSourceIndex, prsDataSource.Fields(0).name
        oChart.SetData chDimValues, nDataSourceIndex, prsDataSource.Fields(2).name
    End If
    
    '当记录大于31条时，不显示数值
    If nCount <= 31 Then
        For i = 1 To oChart.SeriesCollection.Count
            oChart.SeriesCollection(i - 1).DataLabelsCollection.Add
        Next i
    End If
End Sub

Private Sub cboChartType_Click()
    Dim i As Integer
    For i = 0 To ChartSpace1.Charts.Count - 1
        ChartSpace1.Charts(i).Type = cboChartType.ItemData(cboChartType.ListIndex)
    Next i
     
End Sub



Private Sub Form_Load()
    ChartSpace1.ChartWrapCount = 2
    ChartSpace1.Clear

    cboChartType.Clear
    cboChartType.AddItem "柱形图"
    cboChartType.ItemData(0) = 0
    cboChartType.AddItem "条形图"
    cboChartType.ItemData(1) = 3
    cboChartType.AddItem "线形图"
    cboChartType.ItemData(2) = 6
    cboChartType.AddItem "点线图"
    cboChartType.ItemData(3) = 7
    cboChartType.AddItem "饼状图"
    cboChartType.ItemData(4) = 18
    cboChartType.AddItem "比率分布"
    cboChartType.ItemData(5) = 10
    cboChartType.AddItem "环形图"
    cboChartType.ItemData(6) = 32
    
    cboChartType.ListIndex = 0
    Call cboChartType_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraTop.Move Me.ScaleWidth - fraTop.Width, -100
    ChartSpace1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

