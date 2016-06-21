Attribute VB_Name = "mdlDisplay"
Option Explicit

#If DISPLAY = 1 Then
'win98中
'Declare Function dsbdll Lib "cky95h.DLL" (ByVal Port As Integer, ByVal OutString As String) As Integer
'win2000以上
Declare Function dsbdll Lib "ckyNTh.DLL" (ByVal Port As Integer, ByVal OutString As String) As Integer
#End If
Public g_lComPort As Long
Public Const cszComPort = "ComPort" 'Com端口字符串
        

'杭州东站显示屏情况
'1.欢迎来到XXXX车站 工号XXX为您服务 （两行）
'2.到站：XXXX 发车时间:yyyy-mm-dd（语音）                 （大写一行）
'3.收您：XXXX （语音）
'4.应找：XXXX （语音）
'5.转1



Public Sub SetInit()
    '初始化
    
#If DISPLAY = 1 Then
    dsbdll g_lComPort, "f"
    
#End If
    
End Sub

Public Sub SetClear(pnLine As Integer)
    '清空第几行
#If DISPLAY = 1 Then
    dsbdll g_lComPort, "$" & pnLine
    
#End If
End Sub

Public Sub SetUser(pnUserID As String)
    '显示工号
#If DISPLAY = 1 Then
    SetClear 1
    dsbdll g_lComPort, "# 玉环客运中心 欢迎您#"
    SetClear 2
    dsbdll g_lComPort, "#   工号:" & pnUserID & "为您服务#"
#End If
End Sub

Public Sub SetStationAndTime(pszStation As String, pszTime As String, Optional pnTicketNum As Integer = 0)
    '到站及时间
#If DISPLAY = 1 Then
    SetClear 1
    
    dsbdll g_lComPort, "#到站：" & pszStation & IIf(pnTicketNum <> 0, " " & pnTicketNum & "张#", "#")
    SetClear 2
    dsbdll g_lComPort, "#时间：" & Format(pszTime, "MM-dd hh:mm") & "#"
#End If
End Sub


Public Sub SetReceive(pdbMoney As Double)
    '实收
#If DISPLAY = 1 Then

    SetClear 2
    dsbdll g_lComPort, pdbMoney & "Y"
#End If
End Sub


Public Sub SetPay(pdbMoney As Double, pszStation As String, pszTime As String, Optional pnTicketNum As Integer = 0, Optional pdbInsurance As Double)
    '请付款
#If DISPLAY = 1 Then
    SetClear 1
    'dsbdll g_lComPort, "#到站：" & pszStation & IIf(pnTicketNum <> 0, IIf(Len(pszStation) < 3, " ", "") & pnTicketNum & "张", "") & " " & Format(pszTime, "mm月dd日") & "#"
    dsbdll g_lComPort, "#" & pszStation & IIf(pnTicketNum <> 0, IIf(Len(pszStation) < 3, "", "") & pnTicketNum & "张", "") & "" & Format(pszTime, "mm月dd日hh:mm") & "#"
    SetClear 2
    dsbdll g_lComPort, pdbMoney & "J"
    
    If pdbInsurance > 0 Then
        dsbdll g_lComPort, "t"
        dsbdll g_lComPort, pdbInsurance & "E"
    End If
#End If
End Sub

Public Sub SetPay2(pdbMoney As Double)
    '请付款
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, pdbMoney & "J"
#End If
End Sub

Public Sub SetReturn(pdbMoney As Double)
    '应找
#If DISPLAY = 1 Then
    SetClear 2
    '如果相等则显示谢谢
    'dsbdll g_lComPort, "X"
    '不等则显示金额
    dsbdll g_lComPort, pdbMoney & "Z"
#End If
End Sub


Public Sub SetThanks()
    '显示谢谢
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "X"
#End If
End Sub

Public Sub SetCal()
    '显示 找零请当面点清，谢谢
    
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "D"
#End If
End Sub


Public Sub SetWait()
    '显示等待
    
#If DISPLAY = 1 Then
    dsbdll g_lComPort, "W"
#End If
End Sub

Public Sub SetTicketNum()
    '显示要买几张
#If DISPLAY = 1 Then
    SetClear 2
    
    dsbdll g_lComPort, "c"
#End If
    
End Sub

Public Sub SetWhere()
    '显示要去哪儿
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "a"
#End If
End Sub

Public Sub SetCheck()
    '显示请您核对一下
#If DISPLAY = 1 Then
    SetClear 1
    dsbdll g_lComPort, "h"
#End If
End Sub

Public Sub SetQueue()
    '显示请排队购票,谢谢合作
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "k"
#End If
End Sub


Public Sub SetProtect()
    '显示请注意保管随身物品
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "i"
    
#End If
    
End Sub





Public Sub SetInsurance()
    '显示工号
'#If DISPLAY = 1 Then
''    SetClear 1
''    dsbdll g_lComPort, "#保险自愿，弃保请告知#"
'    dsbdll g_lComPort, "l"
'    SetUser m_oAUser.UserID
'#End If
End Sub
