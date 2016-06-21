VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl AddDel2 
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   ScaleHeight     =   3165
   ScaleWidth      =   5985
   Begin MSComctlLib.ListView lvRight 
      Height          =   2475
      Left            =   3600
      TabIndex        =   8
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvLeft 
      Height          =   2535
      Left            =   180
      TabIndex        =   7
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cbAddAll 
      Caption         =   "新增全部>>"
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   1500
      Width           =   1215
   End
   Begin VB.CommandButton cbAddSel 
      Caption         =   "新增所选>"
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cbDelSel 
      Caption         =   "删除所选<"
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cbDelAll 
      Caption         =   "删除全部<<"
      Height          =   315
      Left            =   2100
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "左边的值(&L)"
      Height          =   180
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   990
   End
   Begin VB.Label lblRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "右边的值(&R)"
      Height          =   180
      Left            =   3720
      TabIndex        =   5
      Top             =   0
      Width           =   990
   End
   Begin VB.Label lblMovie 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      ForeColor       =   &H80000007&
      Height          =   180
      Left            =   2460
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "AddDel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'缺省属性值:
Const m_def_LeftData = 0
Const m_def_RightData = 0
Const m_def_LeftLabel = "左边的值(&L)"
Const m_def_RightLabel = "右边的值(&R)"

Const cbDefPlayMovie = True
Const cnDefSleepTime = 80
Const cnDefMoveStep = 400

'属性变量:
Private m_LeftData As Variant
Private m_RightData As Variant

Private m_bPlayMovie As Boolean
Private m_nSleepTime As Integer
Private m_nMoveStep As Integer

Private m_nBusy As Integer

'事件声明:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event DataChange()

Private m_lButtonWidth As Long
Private m_lButtonHeight As Long
Private m_lNap As Long
Private m_lButtonNap As Long
Private m_nColCount As Integer

Private m_nOrgLeftSort As Integer
Private m_nOrgRightSort As Integer
'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Dim oControl As Control
    For Each oControl In UserControl.Controls
        Set oControl.Font = New_Font
    Next
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub cbAddAll_Click()
    Dim i As Integer, j As Integer
    Dim nCount As Integer
    Dim liTemp As ListItem
    nCount = lvLeft.ListItems.Count
    
    For i = 1 To nCount
        MoveData True, 1
    Next
    
    EnableCommand
    RaiseEvent DataChange
End Sub

Private Sub cbAddSel_Click()
    LeftMoveToRight
End Sub

Private Sub cbDelAll_Click()
    Dim i As Integer
    Dim nCount As Integer
    nCount = lvRight.ListItems.Count
    For i = 1 To nCount
        MoveData False, 1
    Next
    
    EnableCommand
    RaiseEvent DataChange
End Sub

Private Sub cbDelSel_Click()
    RightMoveToLeft
End Sub

Private Sub lvLeft_Click()
    
    EnableCommand
End Sub

Private Sub lvLeft_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If m_nOrgLeftSort = ColumnHeader.Position Then
        lvLeft.SortOrder = IIf(lvLeft.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvLeft.SortOrder = lvwAscending
    End If
    lvLeft.SortKey = ColumnHeader.Position - 1
    lvLeft.Sorted = True
    m_nOrgLeftSort = ColumnHeader.Position
End Sub

Private Sub lvLeft_DblClick()
    LeftMoveToRight
End Sub

Private Sub lvRight_Click()
    
    EnableCommand
End Sub

Private Sub lvRight_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If m_nOrgRightSort = ColumnHeader.Position Then
        lvRight.SortOrder = IIf(lvRight.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvRight.SortOrder = lvwAscending
    End If
    lvRight.SortKey = ColumnHeader.Position - 1
    lvRight.Sorted = True
    m_nOrgRightSort = ColumnHeader.Position

End Sub

Private Sub lvRight_DblClick()
    RightMoveToLeft
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    m_lNap = 50
    m_lButtonNap = 200
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14,0,0,0
Public Property Get LeftData() As Variant
    'LeftData = m_LeftData
    
    Dim aszTemp() As String
    Dim i As Integer, j As Integer
    
    If lvLeft.ListItems.Count > 0 Then
        ReDim aszTemp(1 To lvLeft.ListItems.Count, 1 To m_nColCount)
        For i = 1 To lvLeft.ListItems.Count
            aszTemp(i, 1) = lvLeft.ListItems(i).Text
            For j = 1 To m_nColCount - 1
                aszTemp(i, j + 1) = lvLeft.ListItems(i).ListSubItems(j).Text
            Next
        Next
    End If
    
    LeftData = aszTemp
    
End Property

Public Property Let LeftData(ByVal New_LeftData As Variant)
    'm_LeftData = New_LeftData
    Dim i As Integer, nCount As Integer
    Dim j As Integer
    Dim liTemp As ListItem
    lvLeft.ListItems.Clear
    On Error GoTo here
    nCount = UBound(New_LeftData, 1)
    
    For i = 1 To nCount
        Set liTemp = lvLeft.ListItems.Add(, GetKeyStr(New_LeftData(i, 1)), New_LeftData(i, 1))
        For j = 1 To m_nColCount - 1
            liTemp.ListSubItems.Add , , New_LeftData(i, j + 1)
        Next
    Next
    
    EnableCommand
    
    Exit Property
here:
    EnableCommand

End Property

Public Property Get LeftLabel() As String
    LeftLabel = lblLeft.Caption
End Property

Public Property Let LeftLabel(ByVal New_LeftData As String)
    
    lblLeft.Caption = New_LeftData
    
    PropertyChanged "LeftLabel"
End Property


'注意！不要删除或修改下列被注释的行！
'MemberInfo=14,0,0,0
Public Property Get RightData() As Variant
    Dim aszTemp() As String
    Dim i As Integer, j As Integer
    
    If lvRight.ListItems.Count > 0 Then
        ReDim aszTemp(1 To lvRight.ListItems.Count, 1 To m_nColCount)
        For i = 1 To lvRight.ListItems.Count
            aszTemp(i, 1) = lvRight.ListItems(i).Text
            For j = 1 To m_nColCount - 1
                aszTemp(i, j + 1) = lvRight.ListItems(i).ListSubItems(j).Text
            Next
        Next
    End If
    
    RightData = aszTemp
End Property

Public Property Let RightData(ByVal New_RightData As Variant)
    Dim i As Integer, nCount As Integer
    Dim j As Integer
    Dim liTemp As ListItem
    lvRight.ListItems.Clear
    On Error GoTo here
    nCount = UBound(New_RightData, 1)
    
    For i = 1 To nCount
        Set liTemp = lvRight.ListItems.Add(, GetKeyStr(New_RightData(i, 1)), New_RightData(i, 1))
        For j = 1 To m_nColCount - 1
            liTemp.ListSubItems.Add , , New_RightData(i, j + 1)
        Next
    Next
    
    EnableCommand
    
    Exit Property
here:
    EnableCommand
End Property

Public Property Get RightLabel() As String
    RightLabel = lblRight.Caption
End Property

Public Property Let RightLabel(ByVal New_RightData As String)
    
    lblRight.Caption = New_RightData
    PropertyChanged "RightLabel"
End Property


'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Public Sub AddData(paszData As Variant, Optional pbLeft As Boolean = True)
    Dim liTemp As ListItem
    Dim i As Integer
    If pbLeft Then
        'lvLeft.AddItem pszData
        Set liTemp = lvLeft.ListItems.Add(, GetKeyStr(paszData(1)), paszData(1))
    Else
        Set liTemp = lvRight.ListItems.Add(, GetKeyStr(paszData(1)), paszData(1))
    End If
    
    For i = 1 To m_nColCount - 1
        liTemp.ListSubItems.Add , , paszData(i + 1)
    Next
     
    EnableCommand
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Public Sub RemoveData(pszData As String, Optional pbLeft As Boolean = True)
    Dim i As Integer
    Dim liTemp As ListItem
    If pbLeft Then
        Set liTemp = lvLeft.FindItem(GetKeyStr(pszData))
        lvLeft.ListItems.Remove liTemp.Key
    Else
        Set liTemp = lvRight.FindItem(GetKeyStr(pszData))
        lvRight.ListItems.Remove liTemp.Key
    End If
     
    EnableCommand
     
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Private Sub LeftMoveToRight()
    Dim i As Integer
    Dim nCount As Integer
    Dim nTemp As Integer
    Dim nTemp2 As Integer
    Dim szTemp As String
    If m_nBusy = 0 Then
        m_nBusy = m_nBusy + 1
        
        nTemp = 0
        nCount = lvLeft.ListItems.Count
        For i = 1 To nCount
            If lvLeft.ListItems(i - nTemp).Selected Then
                MoveData True, lvLeft.ListItems(i - nTemp).Key
'                szTemp = lvLeft.List(i - 1 - nTemp)
'                lvLeft.RemoveItem (i - 1 - nTemp)
'                SelfPlayMovie szTemp
'                lvRight.AddItem szTemp
                nTemp = nTemp + 1
                nTemp2 = i
            
            End If
        Next
        
        If lvLeft.ListItems.Count > 0 Then
            nTemp2 = IIf(nTemp2 > lvLeft.ListItems.Count, lvLeft.ListItems.Count, nTemp2)
            lvLeft.ListItems(nTemp2).Selected = True
        End If
        
        EnableCommand
        RaiseEvent DataChange
        
        m_nBusy = m_nBusy - 1
    End If
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Private Sub RightMoveToLeft()
    Dim i As Integer
    Dim nCount As Integer
    Dim nTemp As Integer
    Dim nTemp2 As Integer
    Dim szTemp As String
    If m_nBusy = 0 Then
        m_nBusy = m_nBusy + 1
        
        nTemp2 = 0
        nTemp = 0
        nCount = lvRight.ListItems.Count
        For i = 1 To nCount
            If lvRight.ListItems(i - nTemp).Selected Then
                MoveData False, lvRight.ListItems(i - nTemp).Key
'                szTemp = lvRight.List(i - 1 - nTemp)
'                lvRight.RemoveItem (i - 1 - nTemp)
'                SelfPlayMovie szTemp, False
'                lvLeft.AddItem szTemp
                
                nTemp = nTemp + 1
                nTemp2 = i
            End If
        Next
        
        If lvRight.ListItems.Count > 0 Then
            nTemp2 = IIf(nTemp2 > lvRight.ListItems.Count, lvRight.ListItems.Count, nTemp2)
            lvRight.ListItems(nTemp2).Selected = True
        End If
        EnableCommand
        RaiseEvent DataChange
        
        m_nBusy = m_nBusy - 1
    End If
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    Dim oControl As Control
    For Each oControl In UserControl.Controls
        Set oControl.Font = UserControl.Font
    Next
    
    m_LeftData = m_def_LeftData
    m_RightData = m_def_RightData
    m_bPlayMovie = cbDefPlayMovie
    m_nSleepTime = cnDefSleepTime
    m_nMoveStep = cnDefMoveStep
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim oControl As Control

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    For Each oControl In UserControl.Controls
        Set oControl.Font = UserControl.Font
    Next
    
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    
    m_bPlayMovie = PropBag.ReadProperty("PlayMovie", cbDefPlayMovie)
    m_nMoveStep = PropBag.ReadProperty("MoveStep", cnDefMoveStep)
    m_nSleepTime = PropBag.ReadProperty("SleepTime", cnDefSleepTime)
    
    cbAddAll.Caption = PropBag.ReadProperty("AddAllCaption", "新增全部>>")
    cbAddSel.Caption = PropBag.ReadProperty("AddSelCaption", "新增所选>")
    cbDelAll.Caption = PropBag.ReadProperty("DelAllCaption", "删除全部<<")
    cbDelSel.Caption = PropBag.ReadProperty("DelSelCaption", "删除所选<")
    
    cbAddAll.Width = PropBag.ReadProperty("ButtonWidth", 1215)
    cbAddAll.Height = PropBag.ReadProperty("ButtonHeight", 315)
    
    cbAddSel.Width = cbAddAll.Width
    cbAddSel.Height = cbAddAll.Height
    cbDelSel.Width = cbAddAll.Width
    cbDelSel.Height = cbAddAll.Height
    cbDelAll.Width = cbAddAll.Width
    cbDelAll.Height = cbAddAll.Height
    LayoutIt
    
    
    lblLeft.Caption = PropBag.ReadProperty("LeftLabel", m_def_LeftLabel)
    lblRight.Caption = PropBag.ReadProperty("RightLabel", m_def_RightLabel)

End Sub

Private Sub UserControl_Resize()
    LayoutIt
End Sub

Private Sub UserControl_Show()
    LayoutIt
    EnableCommand
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    
    Call PropBag.WriteProperty("AddAllCaption", cbAddAll.Caption, "新增全部>>")
    Call PropBag.WriteProperty("AddSelCaption", cbAddSel.Caption, "新增所选>")
    Call PropBag.WriteProperty("DelSelCaption", cbDelSel.Caption, "删除所选<")
    Call PropBag.WriteProperty("DelAllCaption", cbDelAll.Caption, "删除全部<<")

    Call PropBag.WriteProperty("LeftLabel", lblLeft.Caption, m_def_LeftLabel)
    Call PropBag.WriteProperty("RightLabel", lblRight.Caption, m_def_RightLabel)

    Call PropBag.WriteProperty("ButtonWidth", cbAddAll.Width, 315)
    Call PropBag.WriteProperty("ButtonHeight", cbAddAll.Height, 1215)
    
    Call PropBag.WriteProperty("PlayMovie", m_bPlayMovie, cbDefPlayMovie)
    Call PropBag.WriteProperty("SleepTime", m_nSleepTime, cnDefSleepTime)
    Call PropBag.WriteProperty("MoveStep", m_nMoveStep, cnDefMoveStep)

End Sub


Public Sub LayoutIt()
    Dim lListBoxWidth As Long, lListBoxHeight As Long
    Dim lTemp As Long, lTemp2 As Long
    
    m_lButtonHeight = cbAddAll.Height
    m_lButtonWidth = cbAddAll.Width
    
    lListBoxWidth = (UserControl.ScaleWidth - (2 * m_lNap + 2 * m_lButtonNap + m_lButtonWidth)) / 2
    lListBoxHeight = UserControl.ScaleHeight - (3 * m_lNap + lblLeft.Height)
    
    lListBoxWidth = IIf(lListBoxWidth > 0, lListBoxWidth, 0)
    lListBoxHeight = IIf(lListBoxHeight > 0, lListBoxHeight, 0)
    
    lblLeft.Move m_lNap, m_lNap
    lblRight.Move (m_lNap + 2 * m_lButtonNap + lListBoxWidth + m_lButtonWidth), m_lNap
    
    lvLeft.Move m_lNap, 2 * m_lNap + lblLeft.Height, lListBoxWidth, lListBoxHeight
    lvRight.Move (m_lNap + 2 * m_lButtonNap + lListBoxWidth + m_lButtonWidth), 2 * m_lNap + lblLeft.Height, lListBoxWidth, lListBoxHeight
    
    lTemp = (lListBoxHeight - 4 * m_lButtonHeight) / 5
    cbAddAll.Move m_lNap + lListBoxWidth + m_lButtonNap, 2 * m_lNap + lblLeft.Height + lTemp
    cbAddSel.Move m_lNap + lListBoxWidth + m_lButtonNap, 2 * m_lNap + lblLeft.Height + lTemp * 2 + m_lButtonHeight
    cbDelSel.Move m_lNap + lListBoxWidth + m_lButtonNap, 2 * m_lNap + lblLeft.Height + lTemp * 3 + m_lButtonHeight * 2
    cbDelAll.Move m_lNap + lListBoxWidth + m_lButtonNap, 2 * m_lNap + lblLeft.Height + lTemp * 4 + m_lButtonHeight * 3
End Sub

Private Sub EnableCommand()

    cbAddAll.Enabled = IIf(lvLeft.ListItems.Count = 0, False, True)
    cbAddSel.Enabled = IIf(lvLeft.SelectedItem Is Nothing, False, True)
    
    cbDelAll.Enabled = IIf(lvRight.ListItems.Count = 0, False, True)
    cbDelSel.Enabled = IIf(lvRight.SelectedItem Is Nothing, False, True)
    
End Sub

Public Property Get AddAllCaption() As String
    AddAllCaption = cbAddAll.Caption
End Property

Public Property Let AddAllCaption(ByVal vNewValue As String)
    cbAddAll.Caption = vNewValue
    PropertyChanged "AddAllCaption"
End Property

Public Property Get AddSelCaption() As String
    AddSelCaption = cbAddSel.Caption
End Property

Public Property Let AddSelCaption(ByVal vNewValue As String)
    cbAddSel.Caption = vNewValue
    PropertyChanged "AddSelCaption"

End Property

Public Property Get DelAllCaption() As String
    DelAllCaption = cbDelAll.Caption
End Property

Public Property Let DelAllCaption(ByVal vNewValue As String)
    cbDelAll.Caption = vNewValue
    PropertyChanged "DelAllCaption"
End Property

Public Property Get DelSelCaption() As String
    DelSelCaption = cbDelSel.Caption
End Property

Public Property Let DelSelCaption(ByVal vNewValue As String)
    cbDelSel.Caption = vNewValue
    PropertyChanged "DelSelCaption"
End Property


Public Property Get ButtonWidth() As Long
    ButtonWidth = cbAddAll.Width
End Property

Public Property Let ButtonWidth(ByVal vNewValue As Long)
    cbAddAll.Width = vNewValue
    cbAddSel.Width = cbAddAll.Width
    cbDelSel.Width = cbAddAll.Width
    cbDelAll.Width = cbAddAll.Width
    LayoutIt
    PropertyChanged "ButtonWidth"
    
End Property

Public Property Get ButtonHeight() As Long
    ButtonHeight = cbAddAll.Height
End Property

Public Property Let ButtonHeight(ByVal vNewValue As Long)
    cbAddAll.Height = vNewValue
    cbAddSel.Height = cbAddAll.Height
    cbDelSel.Height = cbAddAll.Height
    cbDelAll.Height = cbAddAll.Height
    LayoutIt
    PropertyChanged "ButtonHeight"
End Property

Public Sub SelfPlayMovie(pszMsg As String, Optional pbLeftToRight As Boolean = True)
    Dim i As Long
    Dim lTemp1 As Long, lTemp2 As Long, lTemp3 As Long
    If PlayMovie Then
        lTemp1 = lblLeft.Top ' lvLeft.Top
        lTemp2 = lvLeft.Left + lvLeft.Width
        lTemp3 = lvRight.Left
        
        lblMovie.Caption = pszMsg
        lblLeft.Visible = False
        lblRight.Visible = False
        lblMovie.Visible = True
        lTemp2 = lTemp2 - lblMovie.Width
        If pbLeftToRight Then
            
            For i = lTemp2 To lTemp3 Step m_nMoveStep
                lblMovie.Move i, lTemp1
                'DoEvents
                Sleep m_nSleepTime
            Next
        Else
            For i = lTemp3 To lTemp2 Step -m_nMoveStep
                lblMovie.Move i, lTemp1
                'DoEvents
                Sleep m_nSleepTime
            Next
        
        End If
        lblMovie.Visible = False
        lblLeft.Visible = True
        lblRight.Visible = True
    
    End If
End Sub

Public Property Get PlayMovie() As Boolean
    PlayMovie = m_bPlayMovie
End Property

Public Property Let PlayMovie(ByVal vNewValue As Boolean)
    m_bPlayMovie = vNewValue
    PropertyChanged "PlayMovie"
End Property

Public Property Get MoveStep() As Integer
    MoveStep = m_nMoveStep
End Property

Public Property Let MoveStep(ByVal vNewValue As Integer)
    m_nMoveStep = vNewValue
    PropertyChanged "MoveStep"
End Property

Public Property Get SleepTime() As Integer
    SleepTime = m_nSleepTime
    
End Property

Public Property Let SleepTime(ByVal vNewValue As Integer)
    m_nSleepTime = vNewValue
    PropertyChanged "SleepTime"
End Property


Private Sub MoveData(pbFromLeftToRight As Boolean, pnIndex As Variant)
    Dim lvFrom  As ListView, lvTo As ListView
    Dim liTemp As ListItem
    Dim i As Integer
    If pbFromLeftToRight Then
        Set lvFrom = lvLeft
        Set lvTo = lvRight
    Else
        Set lvFrom = lvRight
        Set lvTo = lvLeft
    End If
    Set liTemp = lvTo.ListItems.Add(, lvFrom.ListItems(pnIndex).Key, lvFrom.ListItems(pnIndex).Text)
    For i = 1 To m_nColCount - 1
        liTemp.ListSubItems.Add , lvFrom.ListItems(pnIndex).ListSubItems(i).Key, lvFrom.ListItems(pnIndex).ListSubItems(i).Text
    Next
    
    lvFrom.ListItems.Remove pnIndex
End Sub

Private Function DeletePrefix(ByVal pszStr As String, Optional pszPrefixChar As String = "A") As String
    Dim nTemp As Integer
    nTemp = Len(pszStr)
    If nTemp > 1 Then
        DeletePrefix = Right(pszStr, nTemp - 1)
    Else
        DeletePrefix = ""
    End If
End Function

Private Function GetKeyStr(ByVal pszStr As String, Optional pszPrefixChar As String = "A") As String
    GetKeyStr = pszPrefixChar & pszStr
End Function

Public Property Get ColumnHeaders() As Variant
    Dim aszTemp As Variant
    Dim i As Integer
    If m_nColCount > 0 Then
        ReDim aszTemp(1 To m_nColCount)
        For i = 1 To m_nColCount
            aszTemp(i) = lvLeft.ColumnHeaders(i).Text
        Next
        ColumnHeaders = aszTemp
    End If
End Property

Public Property Let ColumnHeaders(vNewValue As Variant)
    On Error GoTo here
    Dim i As Integer
    m_nColCount = UBound(vNewValue)
    lvLeft.ColumnHeaders.Clear
    lvRight.ColumnHeaders.Clear
    For i = 1 To m_nColCount
        lvLeft.ColumnHeaders.Add , , vNewValue(i)
        lvRight.ColumnHeaders.Add , , vNewValue(i)
    Next
    Exit Property
here:
End Property
