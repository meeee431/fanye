VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmUserPwd 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�û�����"
   ClientHeight    =   3000
   ClientLeft      =   2970
   ClientTop       =   4005
   ClientWidth     =   5925
   Icon            =   "frmUserPwd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���Ŀ���"
      Height          =   1635
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   4215
      Begin VB.TextBox txtConfirmPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtNewPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtOldPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȷ���¿���(&R):"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1147
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�¿���(&N):"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   787
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɿ���(&D):"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   420
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�û�����"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4215
      Begin VB.Label lblUnitName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UnitName"
         Height          =   180
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblUserID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UserID"
         Height          =   180
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ:"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�û���:"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   630
      End
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   345
      Left            =   4530
      TabIndex        =   6
      Top             =   210
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ȷ��"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmUserPwd.frx":038A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4530
      TabIndex        =   7
      Top             =   660
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ȡ��"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmUserPwd.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmUserPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public m_oActiveUser As ActiveUser

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo here
    If ChangePassword(txtOldPassword.Text, txtNewPassword.Text, txtConfirmPassword.Text) Then
        MsgBox "��������ɹ����´ν�����ʹ��������Ŷ", vbOKOnly Or vbInformation, "��������"
        Unload Me
    End If
    Exit Sub
here:
    MsgBox err.Description, , "����"
End Sub


Private Sub Form_Load()
    lblUserID.Caption = m_cszOperatorID
    
    lblUnitName.Caption = m_cszOperatorBankID
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set m_oActiveUser = Nothing
End Sub

Private Function ChangePassword(pszOldPassword As String, pszNewPassword As String, pszConfirmPassword As String) As Boolean
    Dim odb As New ADODB.Connection
    Dim szSql As String
    Dim lAffect As Long
    Dim rsTemp As Recordset
    
    ChangePassword = False
    If pszNewPassword <> pszConfirmPassword Then
        MsgBox "�¿�����ȷ�Ͽ��һ��", vbOKOnly, "����"
        Exit Function
    Else
        odb.ConnectionString = GetConnectionStr
        odb.CursorLocation = adUseClient
        odb.Open
        szSql = "UPDATE user_info SET user_password = " & TransFieldValueToString(pszNewPassword) _
            & " WHERE operatorid = " & TransFieldValueToString(m_cszOperatorID) _
            & " AND user_password = " & TransFieldValueToString(pszOldPassword)
        odb.Execute szSql, lAffect
        If lAffect = 0 Then
            MsgBox "ԭ�����ȷ", vbOKOnly, "����"
            Exit Function
        Else
            ChangePassword = True
        End If
    End If
End Function
