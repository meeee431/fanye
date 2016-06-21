VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDITicketMan 
   BackColor       =   &H8000000C&
   Caption         =   "票证管理"
   ClientHeight    =   7395
   ClientLeft      =   825
   ClientTop       =   960
   ClientWidth     =   10800
   HelpContextID   =   4000001
   Icon            =   "MDITicketMan.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDITicketMan.frx":16AC2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abMenuTool 
      Align           =   1  'Align Top
      Height          =   7395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10800
      _LayoutVersion  =   1
      _ExtentX        =   19050
      _ExtentY        =   13044
      _DataPath       =   ""
      Bands           =   "MDITicketMan.frx":16E5C
   End
End
Attribute VB_Name = "MDITicketMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abMenu_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)

End Sub


Private Sub mnu_TitleH_Click()
    Arrange vbTileHorizontal
End Sub

Private Sub mnu_TitleV_Click()
    Arrange vbTileVertical
End Sub
Private Sub mnu_Cascade_Click()
    Arrange vbCascade
End Sub

Private Sub mnu_ArrangeIcon_Click()
    Arrange vbArrangeIcons
End Sub
Private Sub mnu_HelpIndex_Click()
    DisplayHelp Me, Index
End Sub

Private Sub mnu_ModiyCompanyName_Click()
    frmModifyCompany.Show vbModal
End Sub
Private Sub mnu_HelpContent_Click()
    MDIMain.HelpContextID = 60000340
    DisplayHelp Me
End Sub
Private Sub mnu_About_Click()
    Dim oShell As New CommShell
    oShell.ShowAbout App.ProductName, "TicketMan", App.FileDescription, Me.Icon, App.Major, App.Minor, App.Revision
End Sub

Private Sub abMenuTool_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo ErrHandle
Select Case Tool.name
    
    Case "mnuGetTicket"
        frmGetTicket.Show vbModal
        
    Case "mnuRegTicket"
        frmManTicketMan.ZOrder 0
    Case "mnuRegNullTicket"
        frmManNullTicketMan.ZOrder 0
    
    '窗口
    Case "mnu_TitleH"
        mnu_TitleH_Click
    Case "mnu_TitleV"
        mnu_TitleV_Click
    Case "mnu_Cascade"
        mnu_Cascade_Click
    Case "mnu_ArrangeIcon"
        mnu_ArrangeIcon_Click
    '帮助
    Case "mnu_HelpIndex"
        mnu_HelpIndex_Click
    Case "mnu_HelpContent"
        mnu_HelpContent_Click
    Case "mnu_About"
        mnu_About_Click
    
'        '以下是系统部分
'        Case "tbn_system_print"
'            ActiveForm.PrintReport False
'        Case "mnu_system_print"
'            ActiveForm.PrintReport True
'        Case "tbn_system_printview", "mnu_system_printview"
'            ActiveForm.PreView
'        Case "mnu_PageOption"
'            '页面设置
'            ActiveForm.PageSet
'        Case "mnu_PrintOption"
'            '打印设置
'            ActiveForm.PrintSet
'        Case "tbn_system_export", "mnu_ExportFile"
'            ActiveForm.ExportFile
'        Case "tbn_system_exportopen", "mnu_ExportFileOpen"
'            ActiveForm.ExportFileOpen
        Case "mnuChgPassword"
            '修改口令
            ChangePassword
'        Case "mnu_SysExit", "tbn_system_exit"
'            ExitSystem
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub MDIForm_Load()
frmManTicketMan.Show
End Sub
Private Sub ChangePassword()
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.Init m_oActiveUser
    oShell.ShowUserInfo
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
