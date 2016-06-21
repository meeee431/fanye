VERSION 5.00
Begin VB.Form frmChangeRemotePassWord 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改远程登录密码"
   ClientHeight    =   1890
   ClientLeft      =   3480
   ClientTop       =   3735
   ClientWidth     =   5205
   HelpContextID   =   5001001
   Icon            =   "frmChangeRemotePassWord.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   165
      TabIndex        =   5
      Top             =   45
      Width           =   3600
      Begin VB.TextBox txtOld 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1410
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   270
         Width           =   1965
      End
      Begin VB.TextBox txtNew 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1410
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   735
         Width           =   1965
      End
      Begin VB.TextBox txtRe 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1410
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "旧密码(&O):"
         Height          =   180
         Left            =   255
         TabIndex        =   8
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新密码(&P):"
         Height          =   180
         Left            =   255
         TabIndex        =   7
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "确认密码(&S):"
         Height          =   180
         Left            =   255
         TabIndex        =   6
         Top             =   1200
         Width           =   1080
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   315
      Left            =   3930
      TabIndex        =   4
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   315
      Left            =   3930
      TabIndex        =   3
      Top             =   135
      Width           =   1095
   End
End
Attribute VB_Name = "frmChangeRemotePassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim g_szPassword As String

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
End Sub

Private Sub OKButton_Click()
    Dim oUnit As New Unit
    If txtNew.Text = txtRe.Text Then
        On Error GoTo ErrorHandle
        oUnit.Init g_oActUser
        oUnit.Identify g_alvItemText(1)
        oUnit.ChangePassword frmUnitBeUser.szRemoteUserID, txtOld.Text, txtNew.Text, txtRe.Text
        
        MsgBox "成功修改密码.", vbInformation, cszMsg
        
    Else
        MsgBox "两次输入的密码不同,重试.", vbInformation, cszMsg
    End If
    Set oUnit = Nothing
    Unload Me
Exit Sub
ErrorHandle:
    ShowErrorMsg
    Set oUnit = Nothing
    Unload Me
End Sub


Private Sub txtNew_Validate(Cancel As Boolean)
    If TextLongValidate(8, txtNew.Text) Then Cancel = True
End Sub

Private Sub txtOld_Validate(Cancel As Boolean)
    If TextLongValidate(8, txtOld.Text) Then Cancel = True
End Sub

Private Sub txtRe_Validate(Cancel As Boolean)
    If TextLongValidate(8, txtRe.Text) Then Cancel = True
End Sub
