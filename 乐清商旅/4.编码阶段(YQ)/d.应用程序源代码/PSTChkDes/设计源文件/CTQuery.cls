VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowBusQuery(oActUser As Object)
    Dim ofrmQueryBus As New frmQueryBus
    ofrmQueryBus.SelfUser = oActUser
    ofrmQueryBus.Show , MDICheckMain
End Sub
Public Sub ShowTicketQuery(oActUser As Object)
    Dim ofrmQueryTicket As New frmQueryTicket
    ofrmQueryTicket.SelfUser = oActUser
    ofrmQueryTicket.Show , MDICheckMain
End Sub

