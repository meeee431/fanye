VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSeat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境--新增座位"
   ClientHeight    =   990
   ClientLeft      =   4500
   ClientTop       =   4155
   ClientWidth     =   4305
   HelpContextID   =   2007201
   Icon            =   "frmSeat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Top             =   495
      Width           =   1125
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消(&C)"
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
      MICON           =   "frmSeat.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   1125
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmSeat.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   300
      Left            =   2475
      TabIndex        =   5
      Top             =   510
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtEndSeat"
      BuddyDispid     =   196609
      OrigLeft        =   2475
      OrigTop         =   660
      OrigRight       =   2745
      OrigBottom      =   930
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   270
      Left            =   2475
      TabIndex        =   4
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   476
      _Version        =   393216
      BuddyControl    =   "txtStartSeat"
      BuddyDispid     =   196610
      OrigLeft        =   2415
      OrigTop         =   135
      OrigRight       =   2685
      OrigBottom      =   435
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtEndSeat 
      Height          =   315
      Left            =   1575
      TabIndex        =   3
      Text            =   "0"
      Top             =   510
      Width           =   1170
   End
   Begin VB.TextBox txtStartSeat 
      Height          =   300
      Left            =   1575
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束座位号(&E):"
      Height          =   180
      Left            =   225
      TabIndex        =   2
      Top             =   540
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始座位号(&S):"
      Height          =   180
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   1260
   End
End
Attribute VB_Name = "frmSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'                   TOP GROUP INC.
'* Copyright(C)1999 TOP GROUP INC.
'*
'* All rights reserved.No part of this program or publication
'* may be reproduced,transmitted,transcribed,stored in a
'* retrieval system,or translated intoany language or compute
'* language,in any form or by any means,electronic,mechanical,
'* magnetic,optical,chemical,biological,or otherwise,without
'* the prior written permission.
'*********************************************************
'
'**********************************************************
'* Source File Name:frmSeat.frm
'* Project Name:StationNet 2.0
'* Engineer:魏宏旭
'* Data Generated:1999/8/28
'* Last Revision Date:1999/10/6
'* Brief Description:新增座位
'* Relational Document:UI_BS_SM_34.DOC
'**********************************************************
Option Explicit
Public Enum eSeat
    AddSeat = 1
    DeleteSeat = 2
    ReserveSeat = 3
    UnReserveSeat = 4
End Enum
Public m_oREBus As REBus
Public efrmSeat As eSeat

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim i As Integer
On Error GoTo here
Me.MousePointer = vbHourglass
Select Case efrmSeat
       Case 1
        'ShowTBInfo "正在新增座位..."
        m_oREBus.AddSeat Val(txtEndSeat.Text) - Val(txtStartSeat.Text) + 1, txtStartSeat.Text
        'ShowTBInfo "新增座位完成"
        frmRESeat.FullSeat
       Case 2
        m_oREBus.DeleteSeat Val(txtEndSeat.Text) - Val(txtStartSeat.Text), txtStartSeat.Text
       Case 3
        For i = Val(txtStartSeat.Text) To Val(txtEndSeat.Text) - Val(txtStartSeat.Text)
        m_oREBus.ReserveSeat Format(i, "00")
        Next
       Case 4
        For i = Val(txtStartSeat.Text) To Val(txtEndSeat.Text) - Val(txtStartSeat.Text)
        m_oREBus.UnReserveSeat Format(i, "00")
        Next
End Select
Me.MousePointer = vbDefault
Unload Me
Exit Sub
here:
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyEscape
       Unload Me
End Select
End Sub

Private Sub Form_Load()
If m_oREBus.BusID = "" Then
    txtEndSeat.Enabled = False
    txtStartSeat.Enabled = False
    UpDown1.Enabled = False
    UpDown2.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False
End If
Select Case efrmSeat
       Case 1
       frmSeat.Caption = "新增座位"
       txtStartSeat.Text = Format(m_oREBus.TotalSeat + 1, "00")
       Case 2
       frmSeat.Caption = "删除座位"
       Case 3
       frmSeat.Caption = "预留座位"
       Case 4
       frmSeat.Caption = "取消预留"
End Select
cmdOk.Enabled = False
End Sub

Private Sub txtEndSeat_Change()
    If Val(txtEndSeat.Text) > Val(txtStartSeat.Text) Then
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
End Sub

Private Sub txtEndSeat_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdOk_Click
End Select
End Sub

Private Sub txtStartSeat_Change()
    If Val(txtEndSeat.Text) > Val(txtStartSeat.Text) Then
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
End Sub
