VERSION 5.00
Begin VB.Form frmSelect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmSelect"
   ClientHeight    =   3555
   ClientLeft      =   1170
   ClientTop       =   2745
   ClientWidth     =   7155
   HelpContextID   =   5000001
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin PSTSMan.AddDel2 adSelect 
      Height          =   2865
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   5054
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonWidth     =   1215
      ButtonHeight    =   315
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   315
      Left            =   5910
      TabIndex        =   1
      Top             =   3135
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   315
      Left            =   4620
      TabIndex        =   0
      Top             =   3135
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   7110
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   7110
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1
Public m_szCaption As String
Public m_bOk As String
Public m_aszSelect As Variant

Private Sub cmdCancel_Click()
    Unload Me
    m_bOk = False
End Sub

Private Sub cmdOK_Click()
    GetBackInfo
    Unload Me
    m_bOk = True
End Sub

Private Sub Form_Load()
    Dim aszTemp(1 To 2) As String
    Dim aszTemp1(1 To 1) As String
    Dim nLen As Integer, i As Integer
    Dim aszAllFunGroup() As String

    ReDim m_aszSelect(1)

    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2

    Me.Caption = m_szCaption
    Select Case m_szCaption
    Case "选择工作人员"
        aszTemp(1) = "用户代码"
        aszTemp(2) = "用户名"
        adSelect.ColumnHeaders = aszTemp
        nLen = ArrayLength(g_atUserInfo)
        For i = 1 To nLen
            aszTemp(1) = g_atUserInfo(i).UserID
            aszTemp(2) = g_atUserInfo(i).UserName
            Call adSelect.AddData(aszTemp)
        Next i
    Case "选择功能"
        aszTemp(1) = "功能ID"
        aszTemp(2) = "功能名"
        adSelect.ColumnHeaders = aszTemp
        nLen = ArrayLength(g_atAllFun)
        If nLen <> 0 Then
            For i = 1 To nLen
                aszTemp(1) = g_atAllFun(i).szFunctionCode
                aszTemp(2) = g_atAllFun(i).szFunctionName
                Call adSelect.AddData(aszTemp)
            Next i
        End If
    Case "选择功能组"
        aszAllFunGroup = frmStoreMenu.GetAllFunGroup
        aszTemp1(1) = "功能组名"
        adSelect.ColumnHeaders = aszTemp1
        nLen = ArrayLength(aszAllFunGroup)
        If nLen <> 0 Then
            For i = 1 To nLen
                If aszAllFunGroup(i) <> "" Then
                    aszTemp1(1) = aszAllFunGroup(i)
                    Call adSelect.AddData(aszTemp1)
                End If
            Next i
        End If
    End Select
    m_bOk = False
End Sub

Private Sub GetBackInfo()
    Dim aTempRight As Variant
    Dim i As Integer
    Dim nRight As Integer
    aTempRight = adSelect.RightData
    nRight = 0
    nRight = ArrayLength(aTempRight)
    Select Case m_szCaption
    Case "选择工作人员"
        If nRight = 0 Then
            ReDim m_aszSelect(1)
        Else
            ReDim m_aszSelect(1 To nRight)
            For i = 1 To nRight
                m_aszSelect(i) = aTempRight(i, 1)
            Next i
        End If
    Case "选择功能"
        If nRight = 0 Then
            ReDim m_aszSelect(1)
        Else
            ReDim m_aszSelect(1 To nRight)
            For i = 1 To nRight
                m_aszSelect(i) = aTempRight(i, 1)
            Next i
        End If
    Case "选择功能组"
        If nRight = 0 Then
            ReDim m_aszSelect(1)
        Else
            ReDim m_aszSelect(1 To nRight)
            For i = 1 To nRight
                m_aszSelect(i) = aTempRight(i, 1)
            Next i
        End If
    End Select
End Sub
