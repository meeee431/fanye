VERSION 5.00
Object = "{A5E8F770-DA22-4EAF-B7BE-73B06021D09F}#1.1#0"; "ST6Report.ocx"
Begin VB.Form frmSheet 
   Caption         =   "Form1"
   ClientHeight    =   1035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   2655
   StartUpPosition =   3  '����ȱʡ
   Begin ST6Report.RTReport RTReport1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "frmSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moSheetData As AcceptSheet   'Ʊ�ݼ�¼��
Private mszBusStartTime As String   '����ʱ��
Private maszSheetCustom() As String 'Ʊ���е��Զ�������



'���·������
Private Sub FillSheetReport()
On Error GoTo ErrHandler
    
    ReDim maszSheetCustom(1 To 16, 1 To 2)
    Dim mrsSheetData As Recordset
    Dim atPrice() As TLuggagePriceItem
    
    atPrice = moSheetData.PriceItems
    
    '�����Զ�����Ŀ
    maszSheetCustom(1, 1) = "����վ"
    maszSheetCustom(1, 2) = Trim(moSheetData.DesStationName)
    maszSheetCustom(2, 1) = "Ʒ��"
    maszSheetCustom(2, 2) = Trim(moSheetData.LuggageName)
    maszSheetCustom(3, 1) = "����"
    maszSheetCustom(3, 2) = Trim(moSheetData.Number)
    maszSheetCustom(4, 1) = "����"
    maszSheetCustom(4, 2) = Trim(moSheetData.CalWeight)
    maszSheetCustom(5, 1) = "�ջ���λ"
    maszSheetCustom(5, 2) = Trim(moSheetData.Picker)
    maszSheetCustom(6, 1) = "����"
    maszSheetCustom(6, 2) = Trim(moSheetData.SheetID)
    maszSheetCustom(7, 1) = "������1"
    maszSheetCustom(7, 2) = Trim(atPrice(1).PriceValue)
    maszSheetCustom(8, 1) = "������2"
    maszSheetCustom(8, 2) = Trim(atPrice(3).PriceValue)
    maszSheetCustom(9, 1) = "������3"
    maszSheetCustom(9, 2) = Trim(atPrice(2).PriceValue) & "������" & atPrice(4).PriceValue
    maszSheetCustom(10, 1) = "�����˷�"
    maszSheetCustom(10, 2) = Trim(0)
    maszSheetCustom(11, 1) = "�ϼ�(Сд)"
    maszSheetCustom(11, 2) = Trim(atPrice(1).PriceValue + atPrice(2).PriceValue + atPrice(3).PriceValue + atPrice(4).PriceValue)
    maszSheetCustom(12, 1) = "�ϼ�(��д)"
    maszSheetCustom(12, 2) = GetNumber(atPrice(1).PriceValue + atPrice(2).PriceValue + atPrice(3).PriceValue + atPrice(4).PriceValue)
    maszSheetCustom(13, 1) = "����"
    maszSheetCustom(13, 2) = Trim(m_oAUser.UserID)
    maszSheetCustom(14, 1) = "���ʱ��"
    maszSheetCustom(14, 2) = Format(moSheetData.OperateTime, "YYYY-MM-DD HH:mm")
    maszSheetCustom(15, 1) = "���֤��" '(�����˵绰)
    maszSheetCustom(15, 2) = moSheetData.LuggageShipperPhone
    maszSheetCustom(16, 1) = "����ʱ��"
    maszSheetCustom(16, 2) = mszBusStartTime
    
    
    RTReport1.TemplateFile = App.Path & "\�а�����.cll"
    RTReport1.ShowReport mrsSheetData, maszSheetCustom
    
     
    Exit Sub
ErrHandler:
    ShowErrorMsg
End Sub

'��ӡƱ��
Public Sub PrintSheetReport(ByVal poSheetData As AcceptSheet, ByVal pszBusStartTime As String)
    Set moSheetData = poSheetData
    mszBusStartTime = pszBusStartTime
    
    FillSheetReport

    On Error Resume Next
    RTReport1.PrintReport
    
End Sub


