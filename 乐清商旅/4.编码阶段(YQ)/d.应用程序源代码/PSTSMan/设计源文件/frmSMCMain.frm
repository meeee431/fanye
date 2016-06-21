VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSMCMain 
   AutoRedraw      =   -1  'True
   Caption         =   "系统管理"
   ClientHeight    =   7515
   ClientLeft      =   1335
   ClientTop       =   2865
   ClientWidth     =   10470
   HelpContextID   =   5000001
   Icon            =   "frmSMCMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   10470
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrBug 
      Interval        =   500
      Left            =   510
      Top             =   5115
   End
   Begin ComCtl3.CoolBar cbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   741
      BandCount       =   1
      FixedOrder      =   -1  'True
      _CBWidth        =   10470
      _CBHeight       =   420
      _Version        =   "6.7.9782"
      MinHeight1      =   360
      NewRow1         =   0   'False
      Begin RTComctl3.FloatLabel flblMenu 
         Height          =   315
         Index           =   0
         Left            =   -1425
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "FloatLabel1"
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7155
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   635
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ptClient 
      BorderStyle     =   0  'None
      Height          =   4035
      Left            =   855
      ScaleHeight     =   4035
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   900
      Width           =   6795
      Begin VB.PictureBox ptRight 
         BorderStyle     =   0  'None
         Height          =   3300
         Left            =   2940
         ScaleHeight     =   3300
         ScaleWidth      =   3645
         TabIndex        =   7
         Top             =   60
         Width           =   3645
         Begin RTComctl3.FlatLabel lblRight 
            Height          =   360
            Left            =   435
            TabIndex        =   12
            Top             =   465
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OutnerStyle     =   2
            BorderWidth     =   0
            HorizontalAlignment=   1
            Caption         =   " 信息视图"
         End
         Begin VB.PictureBox ptDetail 
            BorderStyle     =   0  'None
            Height          =   2535
            Left            =   210
            ScaleHeight     =   2535
            ScaleWidth      =   2535
            TabIndex        =   8
            Top             =   1260
            Width           =   2535
            Begin RTComctl3.Spliter splDetail 
               Height          =   195
               Left            =   510
               TabIndex        =   11
               Top             =   930
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   344
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MinWidth        =   2000
               IsVertical      =   -1  'True
            End
            Begin MSComctlLib.ListView lvDetail2 
               Height          =   675
               Left            =   615
               TabIndex        =   10
               Top             =   1320
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   1191
               View            =   3
               Arrange         =   1
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
            Begin MSComctlLib.ListView lvDetail 
               Height          =   480
               Left            =   0
               TabIndex        =   9
               Top             =   165
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   847
               View            =   3
               Arrange         =   1
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
         End
      End
      Begin VB.PictureBox ptLeft 
         BorderStyle     =   0  'None
         Height          =   3555
         Left            =   -540
         ScaleHeight     =   3555
         ScaleWidth      =   3225
         TabIndex        =   5
         Top             =   60
         Width           =   3225
         Begin RTComctl3.FlatLabel lblLeft 
            Height          =   360
            Left            =   1425
            TabIndex        =   13
            Top             =   165
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OutnerStyle     =   2
            BorderWidth     =   0
            HorizontalAlignment=   1
            Caption         =   " 系统管理控制树"
         End
         Begin MSComctlLib.TreeView tvAll 
            Height          =   1695
            Left            =   1650
            TabIndex        =   6
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   2990
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   529
            LabelEdit       =   1
            Style           =   7
            BorderStyle     =   1
            Appearance      =   0
         End
      End
      Begin RTComctl3.Spliter Spliter1 
         Height          =   1935
         Left            =   2580
         TabIndex        =   4
         Top             =   780
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   3413
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MeWH            =   30
      End
   End
End
Attribute VB_Name = "frmSMCMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************
' *  Source File Name  : frmSMCMain                                 *
' *  Project Name: PSTSMan                                          *
' *  Engineer:                                              *
' *  Date Generated: 2002/08/19                                     *
' *  Last Revision Date : 2002/08/19                                *
' *  Brief Description   : 主窗体的显示载体                         *
' *******************************************************************

Option Explicit
Const cnNap = 20
'Private Type MENUITEMINFO
'    cbSize As Long
'    fMask As Long
'    fType As Long
'    fState As Long
'    wID As Long
'    hSubMenu As Long
'    hbmpChecked As Long
'    hbmpUnchecked As Long
'    dwItemData As Long
'    dwTypeData As String
'    cch As Long
'End Type
Const MIIM_STATE = &H1
Const MIIM_ID = &H2
Const MIIM_SUBMENU = &H4
Const MIIM_CHECKMARKS = &H8
Const MIIM_TYPE = &H10
Const MIIM_DATA = &H20
Const MF_STRING = &H0
Private Const WM_COMMAND = &H111
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long



Public m_frmMenu As frmStoreMenu
Private m_amnuMain()  As Menu
Private m_aszAltKey() As String
Private m_nAltKeyCount As Integer

Private Sub cbToolBar_DblClick()
    cbToolBar.Align = IIf(cbToolBar.Align = 1, 2, 1)
    LayoutForm

End Sub


Private Sub flblMenu_Click(Index As Integer)
    PopupMenu m_amnuMain(Index), , flblMenu(Index).Left + cbToolBar.Left, flblMenu(Index).Top + flblMenu(Index).Height + cbToolBar.Top   'cbToolBar.Height
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Index As Integer
        
    If Shift = 4 Then
        If Shift <> 18 Then
            For Index = 1 To m_nAltKeyCount
                If m_aszAltKey(Index) = UCase(Chr(KeyCode)) Then
                    PopupMenu m_amnuMain(Index), , flblMenu(Index).Left + cbToolBar.Left, flblMenu(Index).Top + flblMenu(Index).Height + cbToolBar.Top  'cbToolBar.Height
                End If
            Next
        End If
    Else
        DoMenuShortCut KeyCode, Shift
    End If
End Sub

Private Sub Form_Load()

    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2

    
    
    Set m_frmMenu = frmStoreMenu
    m_frmMenu.LoadMenuForm Me
    
    InitToolBar
    Spliter1.InitSpliter ptLeft, ptRight
    splDetail.InitSpliter lvDetail, lvDetail2
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_frmMenu.QueryUnload Cancel, UnloadMode

End Sub

Private Sub Form_Resize()
    LayoutForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload m_frmMenu
    Set m_frmMenu = Nothing
End Sub

Private Sub lvDetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static bAsc As Boolean
bAsc = Not bAsc

lvDetail.SortKey = ColumnHeader.Index - 1
If bAsc Then
    lvDetail.SortOrder = lvwAscending
Else
    lvDetail.SortOrder = lvwDescending
End If
lvDetail.Sorted = True
End Sub

Private Sub lvDetail2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static bAsc1 As Boolean
bAsc1 = Not bAsc1

lvDetail2.SortKey = ColumnHeader.Index - 1
If bAsc1 Then
    lvDetail2.SortOrder = lvwAscending
Else
    lvDetail2.SortOrder = lvwDescending
End If
lvDetail2.Sorted = True

End Sub

Private Sub ptClient_Resize()
    Spliter1.LayoutIt
End Sub

Public Sub InitToolBar()
    Dim i As Integer, j As Integer
    Dim szTemp As String
    Dim mnuTemp As Menu
    Dim lMenu As Long
    Dim aszMenu() As String, nMenuCount As Integer
    Dim miTemp As MENUITEMINFO, lResult As Long
    Dim szTemp2 As String
    Dim nTemp As Integer
    Dim k As Integer, szTemp4 As String
    If Not IsArrayEmpty(m_amnuMain) Then
        nTemp = ArrayLength(m_amnuMain)
        
        For i = 1 To nTemp
            Unload flblMenu(nTemp - i + 1)
        Next
        
    End If
    
    lMenu = GetMenu(m_frmMenu.hwnd)
    nMenuCount = GetMenuItemCount(lMenu)
    If nMenuCount > 0 Then
        ReDim m_amnuMain(1 To nMenuCount)
        ReDim m_aszAltKey(1 To nMenuCount)
        j = 1
        k = 1
        For i = 1 To nMenuCount
            miTemp.cbSize = LenB(miTemp)
            miTemp.fMask = MIIM_DATA Or MIIM_TYPE
            miTemp.cch = 19
            miTemp.dwTypeData = Space(20)
            lResult = GetMenuItemInfo(lMenu, i - 1, 1, miTemp)
            szTemp2 = Trim(miTemp.dwTypeData)
            szTemp2 = Left(szTemp2, Len(szTemp2) - 1) 'String length 减一位（值为零）
            
            Set mnuTemp = FindCaptionMenu(szTemp2)
            If mnuTemp.Visible Then
                Set m_amnuMain(j) = mnuTemp
                Load flblMenu(j)
                flblMenu(j).Caption = GetNoAndString(szTemp2)
                flblMenu(j).Enabled = mnuTemp.Enabled
                flblMenu(j).Move flblMenu(j - 1).Left + flblMenu(j - 1).Width, 50, Me.TextWidth(flblMenu(j).Caption) + 2 * 100
                flblMenu(j).Visible = True
                j = j + 1
                
                szTemp4 = GetAndChar(szTemp2)
                If szTemp4 <> "" Then
                    m_aszAltKey(k) = szTemp4
                    
                    k = k + 1
                End If
            End If
        Next
    Else
    
    End If
    m_nAltKeyCount = k - 1
    Exit Sub
here:
End Sub

Private Function FindCaptionMenu(pszCaption As String) As Menu
    Dim mnuTemp As Control
    For Each mnuTemp In m_frmMenu.Controls
        If TypeName(mnuTemp) = "Menu" Then
            If mnuTemp.Caption = pszCaption Then
                Set FindCaptionMenu = mnuTemp
                Exit For
            End If
        End If
    Next
End Function

Private Sub ptDetail_Resize()
    splDetail.LayoutIt
End Sub

Public Sub LayoutForm()
    Dim lTemp As Long
    If sbStatus.Visible Then
        lTemp = Me.ScaleHeight - (cbToolBar.Height + sbStatus.Height)
    Else
        lTemp = Me.ScaleHeight - cbToolBar.Height
    End If
    lTemp = IIf(lTemp > 0, lTemp, 0)
    If cbToolBar.Align = 1 Then
        ptClient.Move 0, cbToolBar.Height, Me.ScaleWidth, lTemp
    Else
        ptClient.Move 0, 0, Me.ScaleWidth, lTemp
    End If

End Sub

Private Sub ptLeft_Resize()
    PictureResize
End Sub

Public Sub PictureResize(Optional pbLeft As Boolean = True)
    Dim ptTemp As PictureBox
    Dim lblTemp As Control
    Dim oControl As Control
    Dim lTemp As Long
    If pbLeft Then
        Set ptTemp = ptLeft
        Set lblTemp = lblLeft
        Set oControl = tvAll
    Else
        Set ptTemp = ptRight
        Set lblTemp = lblRight
        Set oControl = ptDetail  'lvDetail
    End If
    lblTemp.Move 0, cnNap, ptTemp.ScaleWidth
    lTemp = ptTemp.ScaleHeight - (lblTemp.Height + 2 * cnNap)
    lTemp = IIf(lTemp > 0, lTemp, 0)
    oControl.Move 0, lblTemp.Height + 2 * cnNap, ptTemp.ScaleWidth, lTemp
End Sub

Private Sub ptRight_Resize()
    PictureResize False
End Sub

Private Sub DoMenuShortCut(KeyCode As Integer, Shift As Integer)
    Dim amiTemp() As MENUITEMINFO
    Dim nMenuItemCount As Integer
    Dim szTemp As String
    Dim i As Integer
    On Error GoTo ErrorHandle
    If Shift = 2 And KeyCode <> 17 Then
        amiTemp = GetMenuCaption()
        nMenuItemCount = UBound(amiTemp)
        On Error GoTo 0
        For i = 1 To nMenuItemCount
            szTemp = "Ctrl+" & UCase(Chr(KeyCode))
            If InStr(1, amiTemp(i).dwTypeData, szTemp, vbTextCompare) > 0 Then
                SendMessage m_frmMenu.hwnd, WM_COMMAND, amiTemp(i).wID, &O0
                Exit For
            End If
        Next
    End If
    Exit Sub

ErrorHandle:
    
End Sub

Private Function GetMenuCaption() As MENUITEMINFO()
    Dim amiMenuItemCaption()  As MENUITEMINFO
    Dim miTemp As MENUITEMINFO
    Dim nMenuCount As Integer
    Dim nMenuCount2 As Integer
    Dim oMenu As Control, lResult As Long
    Dim lMenu As Long, lSubMenu As Long, lSubMenuCount As Integer
    Dim j As Integer, i As Integer, k As Integer
    Dim szTemp2 As String
    nMenuCount2 = 0
    For Each oMenu In m_frmMenu.Controls
        If TypeName(oMenu) = "Menu" Then
            nMenuCount2 = nMenuCount2 + 1
        End If
    Next
    ReDim amiMenuItemCaption(1 To nMenuCount2)
    
    lMenu = GetMenu(m_frmMenu.hwnd)
    nMenuCount = GetMenuItemCount(lMenu)
    If nMenuCount > 0 Then
        j = 1
        For i = 1 To nMenuCount
            
            miTemp.cbSize = LenB(miTemp)
            miTemp.fMask = MIIM_DATA Or MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
            miTemp.cch = 19
            miTemp.dwTypeData = Space(20)
            lResult = GetMenuItemInfo(lMenu, i - 1, 1, miTemp)
            If miTemp.hSubMenu <> 0 Then
                lSubMenuCount = GetMenuItemCount(miTemp.hSubMenu)
                lSubMenu = miTemp.hSubMenu
                For k = 1 To lSubMenuCount
                    
                    miTemp.cbSize = LenB(miTemp)
                    miTemp.fMask = MIIM_DATA Or MIIM_TYPE Or MIIM_ID
                    miTemp.cch = 39
                    miTemp.dwTypeData = Space(40)
                    lResult = GetMenuItemInfo(lSubMenu, k - 1, 1, miTemp)
                    If miTemp.fType = 0 Then
                        szTemp2 = Trim(miTemp.dwTypeData)
                        szTemp2 = Left(szTemp2, Len(szTemp2) - 1) 'String length 减一位（值为零）
                        miTemp.dwTypeData = szTemp2
                        amiMenuItemCaption(j) = miTemp
                        j = j + 1
                    End If
                Next
            End If
        Next
   
    End If
    GetMenuCaption = amiMenuItemCaption
End Function

Private Function GetNoAndString(pszIn As String) As String
    Dim szTemp As String
    Dim nIndex As Integer
    nIndex = InStr(1, pszIn, "&", vbTextCompare)
    szTemp = pszIn
    If nIndex > 0 Then
        szTemp = Left(pszIn, nIndex - 1) & Right(pszIn, Len(pszIn) - nIndex)
    End If
    GetNoAndString = szTemp
End Function

Private Function GetAndChar(pszIn As String) As String
    Dim nIndex As Integer
    nIndex = InStr(1, pszIn, "&", vbTextCompare)
    If nIndex > 0 Then
        GetAndChar = Mid(pszIn, nIndex + 1, 1)
    End If
End Function


Private Sub DoDelUser()
End Sub

Private Sub DoDelGroup()
End Sub
