VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************
'* Source File Name:SystemAgent.cls
'* Project Name:STSystem.vbp
'* Engineer:陈峰
'* Data Generated:2002-08-15
'* Last Revision Date:2002-08-15
'* Brief Description:系统代理
'* Relational Document:
'**********************************************************

Option Explicit
Option Base 1

Public Enum EErrSystemAgent
    ERR_CanNotCreateObject = ERR_SystemAgent + 10
    ERR_CreateObjectErr = ERR_SystemAgent + 11
End Enum

Public Property Get LocalUnit() As TUnit
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim unTemp As TUnit
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT * from connect_unit_info WHERE " _
    & " unit_id=(SELECT parameter_value FROM  " & cszSystemParam & "  WHERE " _
    & " parameter_name='LocalUnitID' ) "
    
    Set rsTemp = oDb.Execute(szSql)
    
    unTemp.nUnitType = FormatDbValue(rsTemp!service_type)
    unTemp.szAnnotation = FormatDbValue(rsTemp!Annotation)
    unTemp.szIPAddress = FormatDbValue(rsTemp!ip_address)
    unTemp.szUnitFullName = FormatDbValue(rsTemp!unit_name)
    unTemp.szUnitID = Trim(FormatDbValue(rsTemp!unit_id))
    unTemp.szUnitShortName = FormatDbValue(rsTemp!unit_short_name)
    
    LocalUnit = unTemp
End Property


Public Function CreateServiceObject(ByVal pszObj As String) As Object
    Set CreateServiceObject = CreateObject(pszObj)
End Function




