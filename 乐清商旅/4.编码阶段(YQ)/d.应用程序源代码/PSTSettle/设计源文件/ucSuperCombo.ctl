VERSION 5.00
Begin VB.UserControl ucSuperCombo 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2610
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   11.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1905
   ScaleWidth      =   2610
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   2  'OFF
      Index           =   1
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   1755
      IMEMode         =   2  'OFF
      Index           =   0
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   120
      Width           =   2490
   End
End
Attribute VB_Name = "ucSuperCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Event Declarations:
Event Change() 'MappingInfo=Combo1,Combo1,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click() 'MappingInfo=Combo1,Combo1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLECompleteDrag(Effect As Long) 'MappingInfo=Combo1,Combo1,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Combo1,Combo1,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=Combo1,Combo1,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=Combo1,Combo1,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=Combo1,Combo1,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=Combo1,Combo1,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Private m_Style As Integer
Private mListFields As String
Private mRowSource As Recordset
Private mBoundField As String
Private mBoundText As String
Private mFields() As String
Private mOldValue As String
Private mbChanged As Boolean

Public Property Get OldValue() As String
    OldValue = mOldValue
End Property

Private Sub Combo1_GotFocus(Index As Integer)
If Combo1(m_Style).SelLength = 0 Then
    Combo1(m_Style).SelStart = 0
    Combo1(m_Style).SelLength = Len(Combo1(m_Style).Text)
End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim nIndex As Integer
Dim myString As String
RaiseEvent KeyPress(KeyAscii)
If KeyAscii = 0 Then Exit Sub
If Combo1(m_Style).ListCount = 0 Then Exit Sub
If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeySpace Then
    myString = Left(Combo1(m_Style).Text, Combo1(m_Style).SelStart) & Chr(KeyAscii)
'如果键入的为BACKSPACE，则将选中部分向左移动一位
  ElseIf KeyAscii = vbKeyBack Then
        If Combo1(m_Style).SelStart <> 0 Then
            Combo1(m_Style).SelStart = Combo1(m_Style).SelStart - 1
        End If
        Combo1(m_Style).SelLength = Len(Combo1(m_Style).Text)
        KeyAscii = 0
        Exit Sub
'如果键入为空格或回车，则组合框失去焦点
 ElseIf KeyAscii = vbKeySpace Then KeyAscii = 0: SendTab: Exit Sub
' Or KeyAscii = vbKeyReturn
End If

'清除键入
KeyAscii = 0
nIndex = MatchedIndex(myString)
If nIndex >= 0 Then
    
    If nIndex <> Combo1(m_Style).ListIndex Then Combo1(m_Style).ListIndex = nIndex
    Combo1(m_Style).SelStart = Len(myString)
Else
    Beep
'    Combo1(m_Style).SelStart = 0
End If
Combo1(m_Style).SelLength = Len(Combo1(m_Style).Text)
End Sub

Public Function MatchedIndex(sz As String) As Integer
    Dim i As Integer
    Dim nIndex As Integer
    Dim nLen As Integer
    Dim nCount As Integer
    sz = Trim(sz)
    nLen = Len(sz)
    nIndex = -1
'总记数
    nCount = Combo1(m_Style).ListCount
'从当前选定位置开始查找
    For i = Combo1(m_Style).ListIndex To nCount
        If UCase(sz) = UCase(Left(Combo1(m_Style).List(i), nLen)) Then nIndex = i: Exit For
    Next
'若未找到
    If nIndex < 0 Then
'从开始查找
        For i = 0 To Combo1(m_Style).ListIndex
            If UCase(sz) = UCase(Left(Combo1(m_Style).List(i), nLen)) Then nIndex = i: Exit For
        Next
    End If
    MatchedIndex = nIndex
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    Combo1(m_Style).AddItem Item, Index
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Combo1(m_Style).BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Combo1(m_Style).BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub Combo1_Change(Index As Integer)
    mbChanged = True
    RaiseEvent Change
End Sub

Private Sub Combo1_Click(Index As Integer)
    mbChanged = True
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Combo1(m_Style).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Combo1(m_Style).Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ItemData
Public Property Get ItemData(ByVal Index As Integer) As Long
Attribute ItemData.VB_Description = "Returns/sets a specific number for each item in a ComboBox or ListBox control."
If Index = -1 Then
    ItemData = 0
  Else
    ItemData = Combo1(m_Style).ItemData(Index)
End If
End Property

Public Property Let ItemData(ByVal Index As Integer, ByVal New_ItemData As Long)
    Combo1(m_Style).ItemData(Index) = New_ItemData
    PropertyChanged "ItemData"
End Property

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListCount


Public Property Get List(ByVal Index As Integer) As String
If Index = -1 Then
    List = ""
  Else
    List = Combo1(m_Style).List(Index)
End If
End Property


Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = Combo1(m_Style).ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = Combo1(m_Style).ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    mbChanged = True
    Combo1(m_Style).ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = Combo1(m_Style).MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Combo1(m_Style).MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
'
'Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = Combo1(m_Style).MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    Combo1(m_Style).MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub Combo1_OLECompleteDrag(Index As Integer, Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    Combo1(m_Style).OLEDrag
End Sub

Private Sub Combo1_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,OLEDragMode
Public Property Get OLEDragMode() As Integer
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = Combo1(m_Style).OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As Integer)
    Combo1(m_Style).OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

Private Sub Combo1_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = Combo1(m_Style).OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    Combo1(m_Style).OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub Combo1_OLEGiveFeedback(Index As Integer, Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub Combo1_OLESetData(Index As Integer, Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub Combo1_OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,RemoveItem
Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
    Combo1(m_Style).RemoveItem Index
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = Combo1(m_Style).Sorted
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Style
Public Property Get Style() As Integer
Attribute Style.VB_Description = "Returns/sets a value that determines the type of control and the behavior of its list box portion."
    Style = m_Style
End Property
Public Property Let Style(new_Style As Integer)
    If new_Style = 0 Or new_Style = 1 Then
        Combo1(m_Style).Visible = False
        Combo1(new_Style).Width = Combo1(m_Style).Width
        m_Style = new_Style
        Combo1(m_Style).Visible = True
        Combo1(m_Style).Visible = True
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = Combo1(m_Style).Text
End Property

Public Property Let Text(ByVal New_Text As String)
    mbChanged = True
    Combo1(m_Style).Text = New_Text
    PropertyChanged "Text"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

    Combo1(m_Style).BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set Combo1(m_Style).Font = PropBag.ReadProperty("Font", Ambient.Font)
    Me.Style = PropBag.ReadProperty("Style", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'    combo1(m_style).ItemData(Index) = PropBag.ReadProperty("ItemData" & Index, 0)
'    combo1(m_style).ListIndex = PropBag.ReadProperty("ListIndex", 0)
'    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
'    combo1(m_style).MousePointer = PropBag.ReadProperty("MousePointer", 0)
'    combo1(m_style).OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
'    combo1(m_style).OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
'    combo1(m_style).Text = PropBag.ReadProperty("Text", "Combo1")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

    Call PropBag.WriteProperty("BackColor", Combo1(m_Style).BackColor, &H80000005)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Font", Combo1(m_Style).Font, Ambient.Font)
    Call PropBag.WriteProperty("Style", m_Style, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'    Call PropBag.WriteProperty("ItemData" & Index, combo1(m_style).ItemData(Index), 0)
'    Call PropBag.WriteProperty("ListIndex", combo1(m_style).ListIndex, 0)
'    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
'    Call PropBag.WriteProperty("MousePointer", combo1(m_style).MousePointer, 0)
'    Call PropBag.WriteProperty("OLEDragMode", combo1(m_style).OLEDragMode, 0)
'    Call PropBag.WriteProperty("OLEDropMode", combo1(m_style).OLEDropMode, 0)
'    Call PropBag.WriteProperty("Text", combo1(m_style).Text, "Combo1")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,NewIndex
Public Property Get NewIndex() As Integer
Attribute NewIndex.VB_Description = "Returns the index of the item most recently added to a control."
    NewIndex = Combo1(m_Style).NewIndex
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    Combo1(m_Style).Clear
End Sub

Public Property Get RowSource() As Recordset
    Set mRowSource = RowSource
End Property

Public Property Set RowSource(new_RowSource As Recordset)
    Set mRowSource = new_RowSource
    RefreshRows
    PropertyChanged "RowSource"
End Property

Private Sub RefreshRows()
    Dim sz As String
    sz = Combo1(m_Style).Text
    Combo1(m_Style).Clear
    AppendRows
'    Combo1(m_Style).Text = sz
    If sz <> "" Then
        Combo1(m_Style).ListIndex = MatchedIndex(sz)
    End If
'    RaiseEvent Click
End Sub

Private Sub AppendRows()
    On Error GoTo err
    If UBound(mFields) = 0 Then Exit Sub
    mRowSource.MoveFirst
    Do While Not mRowSource.EOF
        Combo1(m_Style).AddItem GetItemString
        Combo1(m_Style).ItemData(Combo1(m_Style).NewIndex) = mRowSource.Bookmark
        mRowSource.MoveNext
    Loop
    Dim n As Integer
    n = MatchedIndex(Combo1(m_Style).Text)
    If n >= 0 And Combo1(m_Style).Text <> "" Then Combo1(m_Style).ListIndex = n
    Exit Sub
err:
End Sub

Private Function FillSize(psz As String, pnLen As Integer, Optional pszSep As String = " ") As String
    If Len(psz) >= pnLen Then
        FillSize = Left(psz, pnLen)
    Else
        FillSize = psz & String(pnLen - Len(psz), pszSep)
    End If
End Function

Private Function GetItemString() As String
    Dim sz As String
    Dim i As Integer
    On Error GoTo err
    For i = 1 To UBound(mFields)
        If Val(mFields(i, 2)) > 0 Then
            GetItemString = GetItemString & FillSize(Trim(mRowSource.Fields(mFields(i, 1))), Val(mFields(i, 2))) & " "
        Else
            GetItemString = GetItemString & Trim(mRowSource.Fields(mFields(i, 1))) & mFields(i, 2)
        End If
    Next
    Exit Function
err:
End Function

Public Sub AppendWithFields(szListFields As String)
    Dim szTempListFields As String
    szTempListFields = Me.ListFields
    SetListFields szListFields
    AppendRows
    SetListFields szTempListFields
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=,,-1,Text
Public Property Get ListFields() As String
    ListFields = mListFields
End Property

Private Sub SetListFields(ByVal New_ListFields As String)
    Dim i As Integer
    Dim szTemp As String, szTemp1 As String, szTemp2 As String
    mListFields = Trim(New_ListFields)
    szTemp = mListFields
    i = 0
    Do While Not szTemp = ""
        SepByChar szTemp, ","
        i = i + 1
    Loop
    ReDim mFields(i, 2)
    
    szTemp = mListFields
    i = 0
    Do While Not szTemp = ""
        szTemp1 = SepByChar(szTemp, ",")
        szTemp2 = SepByChar(szTemp1, ":")
        If szTemp1 = "" Then szTemp1 = " "
        i = i + 1
        mFields(i, 1) = szTemp2
        mFields(i, 2) = szTemp1
    Loop
End Sub

Public Property Let ListFields(ByVal New_ListFields As String)
    SetListFields New_ListFields
    RefreshRows
    PropertyChanged "ListFields"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=,,-1,Text
Public Property Get BoundField() As String
    BoundField = mBoundField
End Property

Public Property Let BoundField(ByVal New_BoundField As String)
    mBoundField = Trim(New_BoundField)
    PropertyChanged "BoundField"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=,,-1,Text
Public Property Get BoundText() As String
    If mBoundField = "" Or Combo1(m_Style).ListIndex = -1 Then Exit Property
    If Not mRowSource Is Nothing Then
        Dim nBookmark As Integer
        If Combo1(m_Style).ListCount > 0 Then
            nBookmark = Combo1(m_Style).ItemData(Combo1(m_Style).ListIndex)
            If nBookmark > 0 Then
                mRowSource.Bookmark = CDbl(nBookmark)
                BoundText = mRowSource.Fields(mBoundField).Value
            End If
        End If
    End If
End Property


Private Sub UserControl_Resize()
    Combo1(m_Style).Move 0, 0, UserControl.ScaleWidth
'    Combo1(m_Style).Width = UserControl.Width - 50
    If m_Style = 0 Then
        Combo1(m_Style).Height = UserControl.Height
    Else
        UserControl.Height = Combo1(m_Style).Height
    End If
    
End Sub

Public Property Let SelStart(New_SelStart As Integer)
    Combo1(m_Style).SelStart = New_SelStart
End Property

Public Property Let SelLength(New_SelLength As Integer)
    Combo1(m_Style).SelLength = New_SelLength
End Property

Public Property Get Changed() As Boolean
    Changed = mbChanged
    mbChanged = False
End Property

'Public Sub MySetFocus()
'   Combo1(1).Enabled = True
'   Combo1(1).SetFocus
'   Combo1(1).Enabled = False
'
'End Sub
