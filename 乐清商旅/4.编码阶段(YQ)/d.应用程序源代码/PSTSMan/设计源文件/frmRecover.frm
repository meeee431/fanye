VERSION 5.00
Begin VB.Form frmRecover 
   BackColor       =   &H00FFFFFF&
   Caption         =   "恢复"
   ClientHeight    =   2190
   ClientLeft      =   3840
   ClientTop       =   2580
   ClientWidth     =   3975
   HelpContextID   =   5003401
   Icon            =   "frmRecover.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3975
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3165
      Top             =   1050
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   315
      Left            =   2850
      TabIndex        =   2
      Top             =   615
      Width           =   1015
   End
   Begin VB.CommandButton cmdRecover 
      Caption         =   "恢复&R)"
      Height          =   315
      Left            =   2850
      TabIndex        =   1
      Top             =   240
      Width           =   1015
   End
   Begin VB.ListBox lstUnitDel 
      Height          =   1860
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   240
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择待恢复单位(&U)"
      Height          =   180
      Left            =   60
      TabIndex        =   3
      Top             =   45
      Width           =   1530
   End
End
Attribute VB_Name = "frmRecover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'恢复被删除的单位
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
    frmStoreMenu.LoadUnitInfo
End Sub

Private Sub cmdRecover_Click()
    SetBusy
    Dim i As Integer, j As Integer
    Dim szTemp As String
    Dim oUnit As New Unit
    
    On Error GoTo ErrorHandle
    oUnit.Init g_oActUser
    
    For i = 0 To lstUnitDel.ListCount - 1
        If lstUnitDel.Selected(i) = True Then
            lstUnitDel.ListIndex = i
            szTemp = PartCode(lstUnitDel.Text)
            oUnit.Identify szTemp
            oUnit.ReCover
        End If
    Next i
    MsgBox "单位恢复成功!", vbInformation, cszMsg
    j = 0
    For i = 0 To lstUnitDel.ListCount - 1
        If lstUnitDel.Selected(i - j) = True Then
            lstUnitDel.RemoveItem (i - j)
            j = j + 1
        End If
    Next i
    '读数据库刷新内存
    frmStoreMenu.LoadCommonData
    
    
    SetNormal
Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub

Private Sub Form_Load()

    LoadForUnitDel

End Sub

Private Sub LoadForUnitDel()
    '显示删除的单位
    Dim i As Integer
    Dim nLen As Integer
    
    nLen = ArrayLength(g_atAllUnitDelTag)
    
    If nLen > 0 Then
    If g_atAllUnitDelTag(1).szUnitId <> Empty Then
        For i = 1 To nLen
            lstUnitDel.AddItem g_atAllUnitDelTag(i).szUnitId & "[" & g_atAllUnitDelTag(i).szUnitShortName & "]"
        Next i
    End If
    End If

End Sub


Private Sub Timer1_Timer()
    Dim i As Integer
    Dim bTemp As Boolean
    bTemp = False
    For i = 0 To lstUnitDel.ListCount - 1
        If lstUnitDel.Selected(i) = True Then
            bTemp = True
            Exit For
        End If
    Next i
    If bTemp = True Then
        cmdRecover.Enabled = True
    Else
        cmdRecover.Enabled = False
    End If
End Sub
