VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.UserControl AddDel 
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   ScaleHeight     =   3180
   ScaleWidth      =   6270
   Begin RTComctl3.CoolButton cbDelAll 
      Height          =   315
      Left            =   2340
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "删除全部<<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "AddDel.ctx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstRight 
      Appearance      =   0  'Flat
      Height          =   2550
      Left            =   3900
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
   Begin VB.ListBox lstLeft 
      Appearance      =   0  'Flat
      Height          =   2550
      Left            =   240
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1875
   End
   Begin RTComctl3.CoolButton cbAddAll 
      Height          =   315
      Left            =   2430
      TabIndex        =   2
      Top             =   450
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "新增全部>>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "AddDel.ctx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cbDelSel 
      Height          =   315
      Left            =   2550
      TabIndex        =   4
      Top             =   1740
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "删除所选<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "AddDel.ctx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cbAddSel 
      Height          =   315
      Left            =   2490
      TabIndex        =   3
      Top             =   2220
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "新增所选>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "AddDel.ctx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMovie 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      ForeColor       =   &H80000007&
      Height          =   180
      Left            =   2700
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已选列表(&R)"
      Height          =   180
      Left            =   3960
      TabIndex        =   6
      Top             =   180
      Width           =   990
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "待选列表(&L)"
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   990
   End
End
Attribute VB_Name = "AddDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'缺省属性值:
Const m_def_LeftData = 0
Const m_def_RightData = 0
Const m_def_LeftLabel = "待选列表(&L):"
Const m_def_RightLabel = "已选列表(&R):"

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
Attribute Click.VB_Description = "当用户在一个对象上按下并释放鼠标按钮时发生。"
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "当用户在一个对象上按下并释放鼠标按钮后再次按下并释放鼠标按钮时发生。"
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "当用户在拥有焦点的对象上按下任意键时发生。"
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "当用户按下和释放 ANSI 键时发生。"
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "当用户在拥有焦点的对象上释放键时发生。"
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "当用户在拥有焦点的对象上按下鼠标按钮时发生。"
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "当用户移动鼠标时发生。"
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "当用户在拥有焦点的对象上释放鼠标发生。"
Event DataChange()

Private m_lButtonWidth As Long
Private m_lButtonHeight As Long
Private m_lNap As Long
Private m_lButtonNap As Long

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
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
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "强制完全重画一个对象。"
    UserControl.Refresh
End Sub

Private Sub cbAddAll_Click()
    Dim i As Integer
    Dim nCount As Integer
    nCount = lstLeft.ListCount
    For i = 1 To nCount
        lstRight.AddItem lstLeft.List(0)
        lstLeft.RemoveItem 0
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
    nCount = lstRight.ListCount
    For i = 1 To nCount
        lstLeft.AddItem lstRight.List(0)
        lstRight.RemoveItem 0
    Next
    
    EnableCommand
    RaiseEvent DataChange
End Sub

Private Sub cbDelSel_Click()
    RightMoveToLeft
End Sub

Private Sub lstLeft_Click()
    
    EnableCommand
End Sub

Private Sub lstLeft_DblClick()
    LeftMoveToRight
End Sub

Private Sub lstRight_Click()
    
    EnableCommand
End Sub

Private Sub lstRight_DblClick()
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14,0,0,0
Public Property Get LeftData() As Variant
    'LeftData = m_LeftData
    Dim aszTemp() As String
    Dim i As Integer
    If lstLeft.ListCount > 0 Then
        ReDim aszTemp(1 To lstLeft.ListCount)
        For i = 1 To lstLeft.ListCount
            aszTemp(i) = lstLeft.List(i - 1)
        Next
    End If
    LeftData = aszTemp
    
End Property

Public Property Let LeftData(ByVal New_LeftData As Variant)
    'm_LeftData = New_LeftData
    Dim i As Integer, nCount As Integer
    lstLeft.Clear
    On Error GoTo Here
    nCount = UBound(New_LeftData)
    
    For i = 1 To nCount
        lstLeft.AddItem New_LeftData(i)
    Next
    
    EnableCommand
    'PropertyChanged "LeftData"
    Exit Property
Here:
    
    EnableCommand
    
    'PropertyChanged "LeftData"
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
    Dim i As Integer
    If lstRight.ListCount > 0 Then
        ReDim aszTemp(1 To lstRight.ListCount)
        For i = 1 To lstRight.ListCount
            aszTemp(i) = lstRight.List(i - 1)
        Next
    End If
    RightData = aszTemp
    'RightData = m_RightData
End Property

Public Property Let RightData(ByVal New_RightData As Variant)
    'm_RightData = New_RightData
    Dim i As Integer, nCount As Integer
    lstRight.Clear
    On Error GoTo Here
    nCount = UBound(New_RightData)
    
    For i = 1 To nCount
        lstRight.AddItem New_RightData(i)
    Next
    
    EnableCommand
    'PropertyChanged "RightData"
    Exit Property
Here:
    
    EnableCommand
    'PropertyChanged "RightData"
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
Public Sub AddData(pszData As String, Optional pbLeft As Boolean = True)
    If pbLeft Then
        lstLeft.AddItem pszData
    Else
        lstRight.AddItem pszData
    End If
     
    EnableCommand
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Public Sub RemoveData(pszData As String, Optional pbLeft As Boolean = True)
    Dim i As Integer
    If pbLeft Then
        For i = 1 To lstLeft.ListCount
            If lstLeft.List(i - 1) = pszData Then lstLeft.RemoveItem i - 1
        Next
        
    Else
        For i = 1 To lstRight.ListCount
            If lstRight.List(i - 1) = pszData Then lstRight.RemoveItem i - 1
        Next

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
        nCount = lstLeft.ListCount
        For i = 1 To nCount
            If lstLeft.Selected(i - 1 - nTemp) Then
                szTemp = lstLeft.List(i - 1 - nTemp)
                lstLeft.RemoveItem (i - 1 - nTemp)
'                SelfPlayMovie szTemp
                lstRight.AddItem szTemp
                nTemp = nTemp + 1
                nTemp2 = i
            End If
        Next
        
        If lstLeft.ListCount > 0 Then
            nTemp2 = IIf(nTemp2 > lstLeft.ListCount, lstLeft.ListCount, nTemp2)
            lstLeft.Selected(nTemp2 - 1) = True
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
        nCount = lstRight.ListCount
        For i = 1 To nCount
            If lstRight.Selected(i - 1 - nTemp) Then
                szTemp = lstRight.List(i - 1 - nTemp)
                lstRight.RemoveItem (i - 1 - nTemp)
'                SelfPlayMovie szTemp, False
                lstLeft.AddItem szTemp
                nTemp = nTemp + 1
                nTemp2 = i
            End If
        Next
        
        If lstRight.ListCount > 0 Then
            nTemp2 = IIf(nTemp2 > lstRight.ListCount, lstRight.ListCount, nTemp2)
            lstRight.Selected(nTemp2 - 1) = True
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
    
    lstLeft.Move m_lNap, 2 * m_lNap + lblLeft.Height, lListBoxWidth, lListBoxHeight
    lstRight.Move (m_lNap + 2 * m_lButtonNap + lListBoxWidth + m_lButtonWidth), 2 * m_lNap + lblLeft.Height, lListBoxWidth, lListBoxHeight
    
    lTemp = (lListBoxHeight - 4 * m_lButtonHeight) / 5
    cbAddAll.Move m_lNap + lListBoxWidth + m_lButtonNap, 2 * m_lNap + lblLeft.Height + lTemp
    cbAddSel.Move m_lNap + lListBoxWidth + m_lButtonNap, 2 * m_lNap + lblLeft.Height + lTemp * 2 + m_lButtonHeight
    cbDelSel.Move m_lNap + lListBoxWidth + m_lButtonNap, 2 * m_lNap + lblLeft.Height + lTemp * 3 + m_lButtonHeight * 2
    cbDelAll.Move m_lNap + lListBoxWidth + m_lButtonNap, 2 * m_lNap + lblLeft.Height + lTemp * 4 + m_lButtonHeight * 3
End Sub

Private Sub EnableCommand()

    cbAddAll.Enabled = IIf(lstLeft.ListCount = 0, False, True)
    cbAddSel.Enabled = IIf(lstLeft.SelCount = 0, False, True)
    
    cbDelAll.Enabled = IIf(lstRight.ListCount = 0, False, True)
    cbDelSel.Enabled = IIf(lstRight.SelCount = 0, False, True)
    
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
        lTemp1 = lblLeft.Top ' lstLeft.Top
        lTemp2 = lstLeft.Left + lstLeft.Width
        lTemp3 = lstRight.Left
        
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
