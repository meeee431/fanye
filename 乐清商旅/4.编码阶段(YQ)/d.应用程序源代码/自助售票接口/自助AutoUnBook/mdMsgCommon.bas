Attribute VB_Name = "mdMsgCommon"
Option Explicit
Public Const cszSNSysMan = "SNSysMan"
Public Const cszSNRunEnv = "SNRunEnv"
Public Const cszSNChkTK = "SNChkTK"
Public Const cszSystem = "System"

Public Const cszAdjustTime = "AdjustTime"
Public Const cszChangeParam = "ChangeParam"

Public Const cszStopBus = "StopBus"
Public Const cszResumeBus = "ResumeBus"
Public Const cszMergeBus = "MergeBus"
Public Const cszUnSplit = "UnSplit"

Public Const cszRemoveBus = "RemoveBus"
Public Const cszChangeBusTime = "ChangeBusTime"

Public Const cszChangeBusCheckGate = "ChangeBusCheckGate"
Public Const cszChangeBusSeat = "ChangeBusSeat"

Public Const cszAddBus = "AddBus"
Public Const cszChangeBusStandCount = "ChangeBusStandCount"
Public Const cszStartCheckBus = "StartCheckBus"
Public Const cszStopCheckBus = "StopCheckBus"
Public Const cszExStartCheckBus = "ExStartCheckBus"

Public Const cszSlipBusLock = "SlitpBusLock"
Public Const cszSlipBus = "SlitpBus"
Public Const cszChangeSheetNo = "ChangeSheetNo"
Public Const cszMakeEnv = "MakeEnv"

Public Const cszDeleteReBusStation = "DeleteReBusStation"
Public Const cszInsertReBusStation = "InsertReBusStation"
Public Const cszModifyBusPrice = "ModifyBusPrice"
'µ×²ãÍ¨ÓÃ
'------------------------------
Public Const cszMsgPreFix = "RTMSG"
Public Const cszVersion = "Version"
Public Const cszUnit = "Unit"
Public Const cszSendTime = "SendTime"
Public Const cszMsgSource = "MsgSource"
Public Const cszMsgType = "MsgType"
Public Const cszValue = "Value"
Public Const cszSellStation = "SellStation"
Public Const cszInternetQuantity = "InternetQuantity"
Public Const cszCompany = "Company"
Public Const cszVehcile = "Vehicle"
Public Const cszChangeEnvStation = "ChangeEnvStation"
Public Const cszPrefixCmd = "<<"
Public Const cszSuffixCmd = ">>"
'-------------------------------
Public cszSOH As String
Public Function MakeCmdString(pszCmd As String, pszValue As String) As String
    MakeCmdString = cszPrefixCmd & pszCmd & "=" & pszValue & cszSuffixCmd
End Function

Public Function GetCmdValue(pszStr As String, pszCmd As String, Optional pnStartSearch As Integer = 1) As String
    Dim szTemp As String, szValue As String
    Dim nTemp1 As String, nTemp2 As String
    szValue = ""
    szTemp = cszPrefixCmd & Trim(pszCmd) & "="
    nTemp1 = InStr(pnStartSearch, pszStr, szTemp, vbTextCompare)
    If nTemp1 > 0 Then
        nTemp1 = nTemp1 + Len(szTemp)
        nTemp2 = InStr(nTemp1, pszStr, cszSuffixCmd, vbTextCompare)
        If nTemp2 > 0 Then
            szValue = Trim(Mid(pszStr, nTemp1, nTemp2 - nTemp1))
        End If
    End If
    GetCmdValue = szValue
End Function

