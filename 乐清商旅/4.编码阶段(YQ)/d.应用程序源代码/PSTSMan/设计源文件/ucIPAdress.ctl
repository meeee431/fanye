VERSION 5.00
Begin VB.UserControl ucIPAddress 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   LockControls    =   -1  'True
   ScaleHeight     =   375
   ScaleWidth      =   2175
   Begin VB.TextBox txtAdd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   990
      MaxLength       =   3
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   0
      Width           =   360
   End
   Begin VB.TextBox txtAdd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   495
      MaxLength       =   3
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   0
      Width           =   360
   End
   Begin VB.TextBox txtAdd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   0
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   0
      Width           =   360
   End
   Begin VB.TextBox txtAdd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   0
      Width           =   360
   End
   Begin VB.Label lblMask 
      Height          =   240
      Left            =   1545
      TabIndex        =   7
      Top             =   60
      Width           =   585
   End
   Begin VB.Label lblDot 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "·"
      Height          =   180
      Index           =   0
      Left            =   345
      TabIndex        =   1
      Top             =   150
      Width           =   180
   End
   Begin VB.Label lblDot 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "·"
      Height          =   180
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   150
      Width           =   180
   End
   Begin VB.Label lblDot 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "·"
      Height          =   180
      Index           =   2
      Left            =   1335
      TabIndex        =   6
      Top             =   150
      Width           =   180
   End
End
Attribute VB_Name = "ucIPAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'缺省属性值:
Const m_def_ForeColor = 0
Const m_def_BackColor = 0
Const m_def_Enabled = 0
'属性变量:
Dim m_ForeColor As Long
Dim m_BackColor As Long
Dim m_Enabled As Boolean
'事件声明:
Event Validate(Cancel As Boolean)
Attribute Validate.VB_Description = "当控件把焦点移交到引起有效性验证的控件时发生。"
'Event KeyDown(KeyCode As Integer, Shift As Integer)


Private m_IPAddress As String
Private m_aIP(1 To 4) As String
Private m_IPNum As Integer
Public Function TextNotValid(szText As String) As Boolean
On Error GoTo errHandler
    Dim szTemp As String, nTemp As Integer
    szTemp = szText
    nTemp = LenB(szTemp)
    TextNotValid = False
        If nTemp = 0 Then
            SetTextZero
            TextNotValid = False
        Else
            If CInt(szTemp) > 255 Or CInt(szTemp) < 0 Then
                 MsgBox "0到255之间的整数.", vbInformation, "IP地址"
                 TextNotValid = True
            End If
        End If

    Exit Function
errHandler:
            MsgBox "0到255之间的整数.", vbInformation, "IP地址"
            TextNotValid = True
End Function

Private Sub txtAdd_Change(Index As Integer)
    m_IPAddress = txtAdd(0).Text & "." & txtAdd(1).Text & "." & txtAdd(2).Text & "." & txtAdd(3).Text
    m_aIP(Index + 1) = txtAdd(Index).Text
    m_IPNum = Index + 1
End Sub

Private Sub txtAdd_GotFocus(Index As Integer)
    txtAdd(Index).SelStart = 0
    txtAdd(Index).SelLength = Len(txtAdd(Index).Text)
End Sub

Private Sub txtAdd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Integer


    Select Case Index
        Case 0
            If KeyCode = vbKeyRight Then
                If TextNotValid(txtAdd(0)) = False Then
                   txtAdd(1).SetFocus
                Else
                    txtAdd(0).SetFocus
                End If
            End If
        Case 1
            If KeyCode = vbKeyRight Then
                If TextNotValid(txtAdd(1)) = False Then
                    txtAdd(2).SetFocus
                Else
                    txtAdd(1).SetFocus
                End If
            End If
            If KeyCode = vbKeyLeft Then
                If TextNotValid(txtAdd(1)) = False Then
                    txtAdd(0).SetFocus
                Else
                    txtAdd(1).SetFocus
                End If
            End If
        Case 2
            If KeyCode = vbKeyRight Then
                If TextNotValid(txtAdd(2)) = False Then
                    txtAdd(3).SetFocus
                Else
                    txtAdd(2).SetFocus
                End If
            End If
            If KeyCode = vbKeyLeft Then
                If TextNotValid(txtAdd(2)) = False Then
                    txtAdd(1).SetFocus
                Else
                    txtAdd(2).SetFocus
                End If
            End If
        Case 3
            If KeyCode = vbKeyLeft Then
                If TextNotValid(txtAdd(3)) = False Then
                    txtAdd(2).SetFocus
                Else
                    txtAdd(3).SetFocus
                End If
            End If
        Case Else
        '''
    End Select
End Sub



Private Sub txtAdd_Validate(Index As Integer, Cancel As Boolean)
    
   If TextNotValid(txtAdd(Index).Text) = True Then
       Cancel = True
       txtAdd(Index).SetFocus
    Else
        Cancel = False
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim i As Integer
    For i = 1 To 4
        m_aIP(i) = "0"
    Next i
    m_IPAddress = "0.0.0.0"
'    UserControl.Enabled = m_Enabled
End Sub

Private Sub UserControl_Resize()
    Dim lWidth As Long
    Dim i As Integer
    lWidth = (UserControl.ScaleWidth - lblDot(0).Width * 3) / 4
    For i = 1 To 4

        txtAdd(i - 1).Move (i - 1) * (lblDot(0).Width + lWidth), 0, lWidth, UserControl.ScaleHeight
        If i <> 4 Then
            lblDot(i - 1).Move i * lWidth + (i - 1) * lblDot(0).Width, IIf(UserControl.ScaleHeight - lblDot(0).Height > 0, UserControl.ScaleHeight - lblDot(0).Height, 0), lblDot(0).Width
        End If
    Next
    lblMask.Top = 0
    lblMask.Left = 0
    lblMask.Width = UserControl.ScaleWidth
    lblMask.Height = UserControl.ScaleHeight
End Sub
'
'
'
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    
'    UserControl.Enabled = New_Enabled

    Dim i As Integer
    For i = 0 To 3
        txtAdd(i).Enabled = New_Enabled
    Next i

    If New_Enabled = True Then
        lblMask.Visible = False
        lblMask.ZOrder 1
        For i = 0 To 3
            txtAdd(i).Visible = True
        Next i
    Else
        lblMask.Visible = True
        lblMask.ZOrder 0
        For i = 0 To 3
            txtAdd(i).Visible = False
        Next i
    End If
    
    PropertyChanged "Enabled"
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Enabled = m_def_Enabled
    m_ForeColor = m_def_ForeColor
    m_BackColor = m_def_BackColor
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
End Sub


Public Function GetIpAddress() As String
    GetIpAddress = m_IPAddress
End Function

Public Function GetIPNum() As Integer
    GetIPNum = m_IPNum

End Function
Public Sub SetIPDistri(szIPPart As String, IpNum As Integer)
    If IpNum < 5 And IpNum > 0 Then
        txtAdd(IpNum - 1).Text = szIPPart
        m_aIP(IpNum) = szIPPart
    End If
End Sub



Public Function GetIPDistri() As String()
    GetIPDistri = m_aIP
End Function

Private Sub SetTextZero()
    Dim i As Integer
    For i = 0 To 3
        If txtAdd(i).Text = "" Then
            txtAdd(i).Text = "0"
        End If
    Next i
        
    
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

