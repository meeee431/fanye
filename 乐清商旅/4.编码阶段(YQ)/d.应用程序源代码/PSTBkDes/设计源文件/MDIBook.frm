VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIBook 
   BackColor       =   &H8000000C&
   Caption         =   "票务预订"
   ClientHeight    =   5640
   ClientLeft      =   2670
   ClientTop       =   3270
   ClientWidth     =   7575
   Icon            =   "MDIBook.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdPrintSetup 
      Left            =   3840
      Top             =   1860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglToolbar 
      Left            =   1560
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":16AC2
            Key             =   "book"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":17914
            Key             =   "bookquery"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":181EE
            Key             =   "export"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":18348
            Key             =   "exportopen"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":184A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":185FC
            Key             =   "print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":18756
            Key             =   "Env"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":18A72
            Key             =   "Seat"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":18BCE
            Key             =   "Scheme"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIBook.frx":18EEA
            Key             =   "REBus"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ptStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   7575
      TabIndex        =   1
      Top             =   5310
      Width           =   7575
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   1500
         TabIndex        =   3
         Top             =   60
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.StatusBar sbStatus 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   0
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1764
               MinWidth        =   1764
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
               MinWidth        =   3528
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Object.Width           =   1764
               MinWidth        =   1764
               TextSave        =   "2016-5-17"
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   4560
         Top             =   120
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "book"
            Object.ToolTipText     =   "预定"
            ImageKey        =   "book"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "querybook"
            Object.ToolTipText     =   "预定信息查询"
            ImageKey        =   "bookquery"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "REBus"
            Object.ToolTipText     =   "车次属性"
            ImageKey        =   "REBus"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Seat"
            Object.ToolTipText     =   "座位信息"
            ImageKey        =   "Seat"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "exportopen"
            Object.ToolTipText     =   "导出并打开"
            ImageKey        =   "exportopen"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "export"
            Object.ToolTipText     =   "导出"
            ImageKey        =   "export"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "preview"
            Object.ToolTipText     =   "打印预览"
            ImageKey        =   "preview"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "print"
            Object.ToolTipText     =   "打印"
            ImageKey        =   "print"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_System 
      Caption         =   "系统(&S)"
      Begin VB.Menu mnu_Option 
         Caption         =   "选项(&O)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_UserProperty 
         Caption         =   "用户属性(&U)"
      End
      Begin VB.Menu mnu_Space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ExportFile 
         Caption         =   "导出文件(&F)..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_ExportAndOpen 
         Caption         =   "导出文件并打开(&T)..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Space7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Print 
         Caption         =   "打印(&P)"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_PrintPreview 
         Caption         =   "打印预览(&V)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Space8 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_PageSet 
         Caption         =   "页面设置(&G)"
      End
      Begin VB.Menu mnu_PrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnu_Space9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnu_Function 
      Caption         =   "功能(&F)"
      Begin VB.Menu mnu_Book 
         Caption         =   "预定(&B)"
      End
      Begin VB.Menu mnu_BookQuery 
         Caption         =   "预定查询(&Q)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu_Break3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_BookAttrib 
         Caption         =   "属性(&A)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Space4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Sort 
         Caption         =   "排序(&S)"
         Begin VB.Menu mnu_SortBy 
            Caption         =   "流水号"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnu_SortBy 
            Caption         =   "预定人"
            Index           =   1
         End
         Begin VB.Menu mnu_SortBy 
            Caption         =   "预定号"
            Index           =   2
         End
         Begin VB.Menu mnu_SortBy 
            Caption         =   "状态"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnu_RunEnv 
      Caption         =   "环境(&E)"
      Begin VB.Menu mnu_REBusQuery 
         Caption         =   "环境查询(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_REBusPro 
         Caption         =   "车次属性(&P)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Break2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_REBusSeat 
         Caption         =   "座位信息(&S)"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "窗口(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_HTitle 
         Caption         =   "水平平铺(&H)"
      End
      Begin VB.Menu mnu_VTitle 
         Caption         =   "垂直平铺(&V)"
      End
      Begin VB.Menu mnu_Cascade 
         Caption         =   "层叠(&C)"
      End
   End
   Begin VB.Menu mnu_Function2 
      Caption         =   "功能(&F)"
      Visible         =   0   'False
      Begin VB.Menu mnu_Refresh 
         Caption         =   "刷新(&R)"
      End
      Begin VB.Menu mnu_CancelBook 
         Caption         =   "取消预定(&C)"
      End
      Begin VB.Menu mnu_DeleteBook 
         Caption         =   "删除取消(&D)"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnu_HelpIndex 
         Caption         =   "帮助索引(&I)"
      End
      Begin VB.Menu mnu_HelpContent 
         Caption         =   "帮助内容(&C)"
      End
      Begin VB.Menu mnu_Space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "关于预定(&A)"
      End
   End
   Begin VB.Menu mnu_popREBus 
      Caption         =   "车次弹出"
      Visible         =   0   'False
      Begin VB.Menu mnu_popREBusPro 
         Caption         =   "车次属性(&P)"
      End
      Begin VB.Menu mnu_Break4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_popREBusSeat 
         Caption         =   "座位信息(&S)    Ctrl+Z"
      End
   End
End
Attribute VB_Name = "MDIBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnu_About_Click()
    Dim oShell As New CommShell
    oShell.ShowAbout "票务预订", "Ticket Booking System", "票务预订", MDIBook.Image1.Picture, App.Major, App.Minor, App.Revision
    
End Sub

Private Sub mnu_Book_Click()
    frmBook.Show vbModal
End Sub

Private Sub mnu_BookAttrib_Click()
    Dim frmTemp As frmQuery
    Set frmTemp = Me.ActiveForm
    frmTemp.lvSeatInfo_DblClick
End Sub

Private Sub mnu_BookQuery_Click()
    Dim frmTemp As New frmQuery
    frmTemp.Show
End Sub

Private Sub mnu_popREBusPro_Click()
    mnu_REBusPro_Click
End Sub

Private Sub mnu_popREBusSeat_Click()
    mnu_REBusSeat_Click
End Sub

Public Sub mnu_REBusPro_Click()
    frmREBusAttr.m_szBusID = frmREBus.m_szBusID
    frmREBusAttr.m_dtBusDate = frmREBus.m_dtBus
    frmREBusAttr.Show vbModal
End Sub

Private Sub mnu_CancelBook_Click()
    Dim frmQueryTemp As frmQuery
    Set frmQueryTemp = Me.ActiveForm
    If Not frmQueryTemp Is Nothing Then
        frmQueryTemp.UnBook
    End If

End Sub

Private Sub mnu_Cascade_Click()
    Arrange vbCascade
End Sub

Private Sub mnu_DeleteBook_Click()
    Dim frmQueryTemp As frmQuery
    Set frmQueryTemp = Me.ActiveForm
    If Not frmQueryTemp Is Nothing Then
        frmQueryTemp.DeleteBookInfo
    End If

End Sub

Private Sub mnu_Exit_Click()
    Unload Me
End Sub

Private Sub mnu_ExportAndOpen_Click()
    Dim frmTemp As frmQuery
    If Not Me.ActiveForm Is Nothing Then
        If TypeName(Me.ActiveForm) = "frmQuery" Then
            Set frmTemp = Me.ActiveForm
            frmTemp.ExportFile True
        End If
    End If
End Sub

Private Sub mnu_ExportFile_Click()
    Dim frmTemp As frmQuery
    If Not Me.ActiveForm Is Nothing Then
        If TypeName(Me.ActiveForm) = "frmQuery" Then
            Set frmTemp = Me.ActiveForm
            frmTemp.ExportFile False
        End If
    End If
    
End Sub

Private Sub mnu_HTitle_Click()
    Arrange vbTileHorizontal
End Sub

Private Sub mnu_REBusQuery_Click()
    frmREBus.Show
End Sub

Public Sub mnu_REBusSeat_Click()
    frmRESeat.m_szBusID = frmREBus.m_szBusID
    frmRESeat.m_dtBusDate = frmREBus.m_dtBus
    frmRESeat.Show vbModal
End Sub

Private Sub mnu_Refresh_Click()
    Dim frmQueryTemp As frmQuery
    Set frmQueryTemp = Me.ActiveForm
    If Not frmQueryTemp Is Nothing Then
        frmQueryTemp.FillBookInfo
    End If
End Sub

Private Sub mnu_SortBy_Click(Index As Integer)
    Dim frmTemp As frmQuery
    Dim i As Integer
    Set frmTemp = Me.ActiveForm
    For i = 1 To 4
        mnu_SortBy(i - 1).Checked = False
    Next
    Select Case Index
        Case 0
        SortListView frmTemp.lvSeatInfo, 1
        
        Case 1
        SortListView frmTemp.lvSeatInfo, 8
        
        Case 2
        SortListView frmTemp.lvSeatInfo, 5
        
        Case 3
        SortListView frmTemp.lvSeatInfo, 6
    End Select
    mnu_SortBy(Index).Checked = True
End Sub

Private Sub mnu_UserProperty_Click()
    Dim oShell As New CommDialog
    oShell.Init m_oActiveUser
    oShell.ShowUserInfo
End Sub

Private Sub mnu_Print_Click()
'    Dim frmTemp As frmQuery
'    Set frmTemp = Me.ActiveForm
'    Set CellExport1.ListViewSource = frmTemp.lvSeatInfo
'    CellExport1.SourceSelect = ListViewControl
'    CellExport1.PrintEx True
End Sub

Private Sub mnu_PrintPreview_Click()
'    Dim frmTemp As frmQuery
'    Set frmTemp = Me.ActiveForm
'    Set CellExport1.ListViewSource = frmTemp.lvSeatInfo
'    CellExport1.SourceSelect = ListViewControl
'    CellExport1.PrintPreview True
End Sub

Private Sub mnu_PrintSet_Click()
    On Error GoTo Error_Handle
    cdPrintSetup.flags = cdlPDPrintSetup
    cdPrintSetup.ShowPrinter
    Exit Sub
Error_Handle:
End Sub

Private Sub mnu_VTitle_Click()
    Arrange vbTileVertical
End Sub

Private Sub ptStatus_Resize()
    sbStatus.Move 0, 0, ptStatus.ScaleWidth, ptStatus.ScaleHeight
    
    Dim lTemp As Long
    lTemp = sbStatus.Width - (sbStatus.Panels(2).Width + sbStatus.Panels(3).Width)
    lTemp = IIf(lTemp > 1, lTemp, 1)
    sbStatus.Panels(1).Width = lTemp
    
    pbProgress.Move lTemp + 150, 50, sbStatus.Panels(2).Width - 200
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "book"
        mnu_Book_Click
        Case "querybook"
        mnu_BookQuery_Click
        Case "REBus"
        mnu_REBusPro_Click
        Case "Seat"
        mnu_REBusSeat_Click
        Case "export"
        mnu_ExportFile_Click
        Case "exportopen"
        mnu_ExportAndOpen_Click
        Case "preview"
        mnu_PrintPreview_Click
        Case "print"
        mnu_Print_Click
    End Select
End Sub


Public Sub EnableExportAndPrint(pbEnable As Boolean)
    mnu_ExportAndOpen.Enabled = pbEnable
    mnu_ExportFile.Enabled = pbEnable
    mnu_Print.Enabled = pbEnable
    mnu_PrintPreview.Enabled = pbEnable

    mnu_Sort.Enabled = pbEnable
    mnu_BookAttrib.Enabled = pbEnable
    Toolbar1.Buttons("export").Enabled = pbEnable
    Toolbar1.Buttons("exportopen").Enabled = pbEnable
    Toolbar1.Buttons("preview").Enabled = pbEnable
    Toolbar1.Buttons("print").Enabled = pbEnable
  
End Sub
