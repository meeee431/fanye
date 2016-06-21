VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmShowCheckSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "路单信息"
   ClientHeight    =   5985
   ClientLeft      =   1245
   ClientTop       =   1485
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9030
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   -30
      TabIndex        =   4
      Top             =   5220
      Width           =   9135
      Begin RTComctl3.CoolButton cmdCancel 
         Height          =   345
         Left            =   7860
         TabIndex        =   5
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "关闭(&E)"
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
         MICON           =   "frmShowCheckSheet.frx":0000
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
   Begin MSComctlLib.ListView lvCheckSheet 
      Height          =   4095
      Left            =   180
      TabIndex        =   3
      Top             =   1050
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9255
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6750
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShowCheckSheet.frx":001C
               Key             =   "checksheet"
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -30
         TabIndex        =   1
         Top             =   810
         Width           =   9285
      End
      Begin VB.Label lblSettleSheetID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200340001"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1620
         TabIndex        =   6
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算单号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   540
         TabIndex        =   2
         Top             =   330
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmShowCheckSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PI_CheckSheetID = 0
Const PI_BusDate = 1
Const PI_BusID = 2
Const PI_BusSerialNO = 3
Const PI_LicenseTagNo = 4
Const PI_CompanyName = 5
Const PI_RouteID = 6
Const PI_VehicleType = 7
Const PI_Owner = 8
Const PI_Checker = 9

Public m_szSettleSheetID As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    lblSettleSheetID.Caption = m_szSettleSheetID
    FillHead
    FillForm
End Sub

Private Sub FillHead()
On Error GoTo here
    With lvCheckSheet.ColumnHeaders
        .Add , , "路单代码"
        .Add , , "日期"
        .Add , , "车次代码"
        .Add , , "车次序号"
        .Add , , "车辆"
        .Add , , "参运公司"
        .Add , , "线路"
        .Add , , "车型"
        .Add , , "车主"
        .Add , , "检票员"
    End With
    AlignHeadWidth Me.name, lvCheckSheet
    Exit Sub
here:
    ShowErrorMsg

End Sub

Private Sub FillForm()
On Error GoTo here
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim lvItem As ListItem
    Dim m_oReport As New Report
    m_oReport.Init g_oActiveUser
    Set rsTemp = m_oReport.GetCheckSheetInfo(m_szSettleSheetID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    For i = 1 To rsTemp.RecordCount
       Set lvItem = lvCheckSheet.ListItems.Add(, , FormatDbValue(rsTemp!check_sheet_id), , "checksheet")
       lvItem.SubItems(PI_BusDate) = Format(FormatDbValue(rsTemp!bus_date), "yyyy-MM-dd")
       lvItem.SubItems(PI_BusID) = FormatDbValue(rsTemp!bus_id)
       lvItem.SubItems(PI_BusSerialNO) = FormatDbValue(rsTemp!bus_serial_no)
       lvItem.SubItems(PI_LicenseTagNo) = FormatDbValue(rsTemp!license_tag_no)
       lvItem.SubItems(PI_CompanyName) = FormatDbValue(rsTemp!transport_company_short_name)
       lvItem.SubItems(PI_RouteID) = FormatDbValue(rsTemp!route_name)
       lvItem.SubItems(PI_VehicleType) = FormatDbValue(rsTemp!vehicle_type_name)
       lvItem.SubItems(PI_Owner) = FormatDbValue(rsTemp!owner_name)
       lvItem.SubItems(PI_Checker) = FormatDbValue(rsTemp!Checker)
       rsTemp.MoveNext
    Next i
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    SaveHeadWidth Me.name, lvCheckSheet
    Unload Me
End Sub

Private Sub lvCheckSheet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If lvCheckSheet.SortOrder = lvwAscending Then
    lvCheckSheet.SortOrder = lvwDescending
 Else
    lvCheckSheet.SortOrder = lvwAscending
 End If
    lvCheckSheet.SortKey = ColumnHeader.Index - 1
    lvCheckSheet.Sorted = True
End Sub

Private Sub lvCheckSheet_DblClick()
    
    Dim oCommDialog As New STShell.CommDialog
    On Error GoTo here
    oCommDialog.Init g_oActiveUser
    oCommDialog.ShowCheckSheet lvCheckSheet.SelectedItem.Text
    Set oCommDialog = Nothing
    Exit Sub
here:
    ShowErrorMsg
End Sub
