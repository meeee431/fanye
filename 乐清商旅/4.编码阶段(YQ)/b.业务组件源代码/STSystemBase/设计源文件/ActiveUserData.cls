VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActiveUserData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_lLoginEventID As Long '登录事件ID
Public m_lLoginEventID2 As Long '登录事件ID2
Public m_szUserID As String '登录用户ID
Public m_dtLoginTime As Date '登录时间
Public m_szLoginHost As String '登录工作站
Public m_dtLastTime As Date '最后使用时间
Public m_bInternetUser As Boolean '是否是INTERNET用户


'2002-11-12陈峰加
Public m_szSellStationID As String '用户所属售票站
