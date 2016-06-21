VERSION 5.00
Begin VB.Form frmMsgbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Title"
   ClientHeight    =   1395
   ClientLeft      =   3600
   ClientTop       =   4215
   ClientWidth     =   3915
   Icon            =   "frmMsgbox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Tag             =   "Modal"
   Begin VB.CommandButton cmdButton3 
      Caption         =   "BUTTON3"
      Default         =   -1  'True
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   930
      Width           =   1125
   End
   Begin VB.CommandButton cmdButton2 
      Caption         =   "BUTTON2"
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   930
      Width           =   1125
   End
   Begin VB.CommandButton cmdButton1 
      Caption         =   "BUTTON1"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   930
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   4110
      Picture         =   "frmMsgbox.frx":000C
      Top             =   1860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   4020
      Picture         =   "frmMsgbox.frx":044E
      Top             =   1950
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   4080
      Picture         =   "frmMsgbox.frx":0890
      Top             =   1860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   4020
      Picture         =   "frmMsgbox.frx":0CD2
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPrompt 
      Height          =   555
      Left            =   210
      Top             =   150
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Caption         =   "Prompt"
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Top             =   330
      Width           =   2640
   End
End
Attribute VB_Name = "frmMsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mszTitle As String
Dim mszPrompt As String
Dim meButtons As VbMsgBoxStyle
Dim meResult As VbMsgBoxResult      '返回结果
Dim mnDefault As Integer            '缺省按钮

Public Property Get Title() As String
    Title = mszTitle
End Property

Public Property Let Title(ByVal szNewTitle As String)
    mszTitle = szNewTitle
End Property

Public Property Get Prompt() As String
    Prompt = mszPrompt
End Property

Public Property Let Prompt(ByVal szNewPrompt As String)
    mszPrompt = szNewPrompt
End Property

Public Property Get Result() As VbMsgBoxResult
    Result = meResult
End Property


Public Property Get Buttons() As VbMsgBoxStyle
    Buttons = meButtons
End Property

Public Property Let Buttons(ByVal eNewButtons As VbMsgBoxStyle)
    meButtons = eNewButtons
End Property

Private Sub cmdButton1_Click()
    meResult = Val(cmdButton1.Tag)
    Unload Me
End Sub

Private Sub cmdButton2_Click()
    meResult = Val(cmdButton2.Tag)
    Unload Me
End Sub

Private Sub cmdButton3_Click()
    meResult = Val(cmdButton3.Tag)
    Unload Me
End Sub

Private Sub Form_Activate()
    Select Case mnDefault
        Case 2
            If cmdButton2.Visible Then
                cmdButton2.SetFocus
            End If
        Case 3
            If cmdButton3.Visible Then
                cmdButton3.SetFocus
            End If
        Case Else
            cmdButton1.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    '确定按钮排列
    Dim nCountButtons As Integer         '按钮数
    Select Case meButtons And 15    'meButtons & 0XOF
        Case vbOKCancel
            cmdButton1.Tag = vbOK
            cmdButton1.Caption = "确定"
            cmdButton2.Tag = vbCancel
            cmdButton2.Caption = "取消"
            nCountButtons = 2
        Case vbAbortRetryIgnore
            cmdButton1.Tag = vbAbort
            cmdButton1.Caption = "中止(&A)"
            cmdButton2.Tag = vbRetry
            cmdButton2.Caption = "重试(&R)"
            cmdButton3.Tag = vbIgnore
            cmdButton3.Caption = "忽略(&N)"
            nCountButtons = 3
        Case vbYesNoCancel
            cmdButton1.Tag = vbYes
            cmdButton1.Caption = "是(&Y)"
            cmdButton2.Tag = vbNo
            cmdButton2.Caption = "否(&N)"
            cmdButton3.Tag = vbCancel
            cmdButton3.Caption = "取消"
            nCountButtons = 3
        Case vbYesNo
            cmdButton1.Tag = vbYes
            cmdButton1.Caption = "是(&Y)"
            cmdButton2.Tag = vbNo
            cmdButton2.Caption = "否(&N)"
            nCountButtons = 2
        Case vbRetryCancel
            cmdButton1.Tag = vbRetry
            cmdButton1.Caption = "重试(&R)"
            cmdButton2.Tag = vbCancel
            cmdButton2.Caption = "取消"
            nCountButtons = 2
        Case Else
            cmdButton1.Tag = vbOK
            cmdButton1.Caption = "确定"
            nCountButtons = 1
    End Select
    cmdButton1.Visible = IIf(nCountButtons > 0, True, False)
    cmdButton2.Visible = IIf(nCountButtons > 1, True, False)
    cmdButton3.Visible = IIf(nCountButtons > 2, True, False)
        
    '确定标识图样
    imgPrompt.Visible = True
    Select Case (meButtons - (meButtons Mod 16)) And 127 'meButtons & OxF0
        Case vbCritical
            imgPrompt.Picture = Image1(3).Picture
        Case vbQuestion
            imgPrompt.Picture = Image1(2).Picture
        Case vbInformation
            imgPrompt.Picture = Image1(1).Picture
        Case vbExclamation
            imgPrompt.Picture = Image1(0).Picture
        Case Else
            imgPrompt.Visible = False
    End Select
    
    '布置缺省按钮
    Select Case (meButtons - (meButtons Mod 128)) And 1023 'meButtons & 0x0F00
        Case vbDefaultButton2
            mnDefault = 2
        Case vbDefaultButton3
            mnDefault = 3
        Case Else
            mnDefault = 1
    End Select
    
    Me.Caption = mszTitle
    lblPrompt.Caption = mszPrompt
    If Not imgPrompt.Visible Then
        lblPrompt.Left = imgPrompt.Left
    End If
    Dim nFormAllowWidth As Integer
    nFormAllowWidth = lblPrompt.Left + lblPrompt.Width + 300
    If nFormAllowWidth > Me.ScaleWidth Then
        If nFormAllowWidth < 2 * Me.ScaleWidth Then
            Me.Width = nFormAllowWidth
        Else
            Me.Width = 2 * Me.ScaleWidth
        End If
    End If
        
        
    Dim nSepWidth As Integer
    nSepWidth = cmdButton2.Left - cmdButton1.Left - cmdButton1.Width
    Select Case nCountButtons
        Case 1
            cmdButton1.Left = Me.ScaleWidth / 2 - cmdButton1.Width / 2
        Case 2
            cmdButton1.Left = Me.ScaleWidth / 2 - nSepWidth / 2 - cmdButton1.Width
            cmdButton2.Left = Me.ScaleWidth / 2 + nSepWidth / 2
        Case 3
            cmdButton1.Left = Me.ScaleWidth / 2 - cmdButton1.Width / 2 - cmdButton1.Width - nSepWidth
            cmdButton2.Left = Me.ScaleWidth / 2 - cmdButton1.Width / 2
            cmdButton3.Left = Me.ScaleWidth / 2 + cmdButton1.Width / 2 + nSepWidth
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If meResult = 0 Then
        If cmdButton3.Visible Then
            meResult = Val(cmdButton3.Tag)
        ElseIf cmdButton2.Visible Then
            meResult = Val(cmdButton2.Tag)
        Else
            meResult = Val(cmdButton1.Tag)
        End If
    End If
End Sub



