VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LugFunLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_szLastError As String

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser
Private m_aszAllVehicleRatio() As String
Private m_bIsRelationWithVehicleType



Public Sub Init(AUser As ActiveUser)
    Dim oParam As New SystemParam
    
    Set SelfUser = AUser
    RefreshVehicleTypeRation
    oParam.Init m_oActiveUser
    '是否是与车型有关联
    m_bIsRelationWithVehicleType = False 'oParam.IsRelationWithVehicleType
    Set oParam = Nothing
    
End Sub

Private Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Private Property Set SelfUser(AUser As ActiveUser)
    Set m_oActiveUser = AUser
    m_oActiveUser.IAmLive
'    m_oDb.ConnectionString = GetConnectionStr(cszPILib)
    
End Property

Public Property Get LuggageFormulaInfo() As TPriceFormulaInfo()
    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
    Dim apfiTemp(1 To 8) As TPriceFormulaInfo
    apfiTemp(1).nFunParamCount = 2 '参数个数
    apfiTemp(1).szFunIntroduce = "行包基本运价的介绍（公里数*重量*费率）" '公式介绍
    apfiTemp(1).szFunName = "LuggageBaseCarriage" '具体使用的方法名
    apfiTemp(1).szFunChineseName = "基本运价" '公式中文名
    apfiTemp(1).bBaseCarriage = False '是不是基本运价公式
    ReDim apfiTemp(1).aszParamIntroduce(1 To 2)
    apfiTemp(1).aszParamIntroduce(1) = "费率" '参数的介绍
    apfiTemp(1).aszParamIntroduce(2) = "最小运价" '参数的介绍
    
    
    
    apfiTemp(2).nFunParamCount = 2
    apfiTemp(2).szFunIntroduce = "每件计算法的介绍（件数*费率）"
    apfiTemp(2).szFunName = "LuggagePerBaggageNumber"
    apfiTemp(2).szFunChineseName = "每件计算法"
    apfiTemp(2).bBaseCarriage = False
    
    ReDim apfiTemp(2).aszParamIntroduce(1 To 2)
    apfiTemp(2).aszParamIntroduce(1) = "费率"
    apfiTemp(2).aszParamIntroduce(2) = "超重费率"
    
    
    
    apfiTemp(3).nFunParamCount = 2
    apfiTemp(3).szFunIntroduce = "等件数计算法的介绍" & Chr(10) & "即件数折合成参数1的倍数如：3件按5件算,7件按10件算,再乘以费率"
    apfiTemp(3).szFunName = "LuggageEqBaggageNumber"
    apfiTemp(3).szFunChineseName = "等件数计算法"
    apfiTemp(3).bBaseCarriage = False
    ReDim apfiTemp(3).aszParamIntroduce(1 To 2)
    apfiTemp(3).aszParamIntroduce(1) = "计数单位（件数）"
    apfiTemp(3).aszParamIntroduce(2) = "费率"
        
'    apfiTemp(4).nFunParamCount = 2
'    apfiTemp(4).szFunIntroduce = "加超重行包的每件计算法的介绍" & Chr(10) & "件数*费率+超重件数*超重费率"
'    apfiTemp(4).szFunName = "LuggagePerBaggageNumberAddOver"
'    apfiTemp(4).szFunChineseName = "加超重行包的每件计算法"
'    apfiTemp(4).bBaseCarriage = False
'
'    ReDim apfiTemp(4).aszParamIntroduce(1 To 2)
'    apfiTemp(4).aszParamIntroduce(1) = "费率"
'    apfiTemp(4).aszParamIntroduce(2) = "超重费率"
    
    apfiTemp(4).nFunParamCount = 1
    apfiTemp(4).szCheckParamValidFunName = "AssertConstantParamValid"
    apfiTemp(4).szFunIntroduce = "常数公式:计算结果为参数1值"
    apfiTemp(4).szFunName = "LuggageConstantData"
    apfiTemp(4).szFunChineseName = "常数公式"
    apfiTemp(4).bBaseCarriage = False
    ReDim apfiTemp(4).aszParamIntroduce(1 To 1)
    apfiTemp(4).aszParamIntroduce(1) = "使对应的票价项的值为参数1 的值,范围为:大于或等于 0 且小于等于 5000"
    
    apfiTemp(5).nFunParamCount = 3
    apfiTemp(5).szFunIntroduce = "行包基本运价的件数计重取平均数计算介绍（公里数*平均计重*费率）   平均计重=计重/(件数+超重件数)" & vbCrLf _
                                            & "每件货物的运价为：" & vbCrLf _
                                            & "1．当计重<=最小计重时，运价=最小运价 " & vbCrLf _
                                            & "2．当计重>最小计重时，" & vbCrLf _
                                            & "如果计算运费<=最小运价时，运费=最小运价+（计重-最小计重）*费率*里程" & vbCrLf _
                                            & "如果计算运费>最小运价时，运费=计算运费"

    apfiTemp(5).szFunName = "LuggageBaseCarriageAvg"
    apfiTemp(5).szFunChineseName = "基本运价平均值"
    apfiTemp(5).bBaseCarriage = False
    ReDim apfiTemp(5).aszParamIntroduce(1 To 3)
    apfiTemp(5).aszParamIntroduce(1) = "费率"
    apfiTemp(5).aszParamIntroduce(2) = "最小运价"
    apfiTemp(5).aszParamIntroduce(3) = "最小计重"
    
    apfiTemp(6).nFunParamCount = 10
    apfiTemp(6).szFunIntroduce = "行包分段计算公式" & Chr(13) & Chr(10) _
        & "行包重量取平均数计算 (平均重量 = 计重 / 件数)" & Chr(13) & Chr(10) _
        & "0-[参数1]公斤之间(含[参数1]),每件的价格为[参数2]元；" & Chr(13) & Chr(10) _
        & "[参数1]-[参数3]公斤之间(含[参数3]),每件重量加[参数4]的公斤数,价格=[参数2]+[参数5]*((重量-[参数1])/[参数4])向上取整；" & Chr(13) & Chr(10) _
        & "[参数3]-[参数6]公斤之间(含[参数6]),每件重量加[参数7]的公斤数,价格=上一段的最高价格+[参数8]*((重量-[参数3])/[参数7])向上取整；" & Chr(13) & Chr(10) _
        & "超过[参数6]公斤数,价格=(重量/[参数9])向上取整*[参数10]"
    
    '        dbPrice = Param2 + CUint((Param3 - Param1), Param4) * Param5 + Param8 * nCount
    
    apfiTemp(6).szFunName = "LuggageConstantSeg"
    apfiTemp(6).szFunChineseName = "行包分段计算公式"
    apfiTemp(6).bBaseCarriage = False
    
    ReDim apfiTemp(6).aszParamIntroduce(1 To 10)
    apfiTemp(6).aszParamIntroduce(1) = "起步的公斤数"
    apfiTemp(6).aszParamIntroduce(2) = "起步的价格"
    apfiTemp(6).aszParamIntroduce(3) = "第二段公斤数"
    apfiTemp(6).aszParamIntroduce(4) = "重量的增量"
    apfiTemp(6).aszParamIntroduce(5) = "价格增量"
    apfiTemp(6).aszParamIntroduce(6) = "第三段的公斤数"
    apfiTemp(6).aszParamIntroduce(7) = "重量的增量"
    apfiTemp(6).aszParamIntroduce(8) = "价格增量"
    apfiTemp(6).aszParamIntroduce(9) = "超过第三段的除数"
    apfiTemp(6).aszParamIntroduce(10) = "乘数"
    
    apfiTemp(7).nFunParamCount = 7
    apfiTemp(7).szFunIntroduce = "行包分段计算公式分3段" & Chr(13) & Chr(10) _
        & "行包重量取平均数计算 (平均重量 = 计重 / 件数)" & Chr(13) & Chr(10) _
        & "0-[参数1]公斤之间(含[参数1]),每件的价格为[参数2]元；" & Chr(13) & Chr(10) _
        & "[参数1]-[参数3]公斤之间(含[参数3]),每件重量加[参数4]的公斤数,价格=[参数2]+[参数5]*((重量-[参数1])/[参数4])向上取整；" & Chr(13) & Chr(10) _
         & "超过[参数3]公斤数,价格=(重量/[参数6])向上取整*[参数7]"
    
    '        dbPrice = Param2 + CUint((Param3 - Param1), Param4) * Param5 + Param8 * nCount
    
    apfiTemp(7).szFunName = "LuggageConstantSeg3Seg"
    apfiTemp(7).szFunChineseName = "行包分段计算公式分3段"
    apfiTemp(7).bBaseCarriage = False
    
    ReDim apfiTemp(7).aszParamIntroduce(1 To 7)
    apfiTemp(7).aszParamIntroduce(1) = "起步的公斤数"
    apfiTemp(7).aszParamIntroduce(2) = "起步的价格"
    apfiTemp(7).aszParamIntroduce(3) = "第二段的公斤数"
    apfiTemp(7).aszParamIntroduce(4) = "重量的增量"
    apfiTemp(7).aszParamIntroduce(5) = "价格增量"
    apfiTemp(7).aszParamIntroduce(6) = "超过第二段的除数"
    apfiTemp(7).aszParamIntroduce(7) = "乘数"
    
    
    apfiTemp(8).nFunParamCount = 1 '参数个数
    apfiTemp(8).szFunIntroduce = "行包基本运价的介绍（公里数*重量*车型费率） 车型费率为该车辆对应的车型费率。 " '公式介绍
    apfiTemp(8).szFunName = "LuggageBaseCarriageByVehicleType" '具体使用的方法名
    apfiTemp(8).szFunChineseName = "基本运价（按车型费率算）" '公式中文名
    apfiTemp(8).bBaseCarriage = True '是不是基本运价公式
    
    ReDim apfiTemp(8).aszParamIntroduce(1 To 1)
    apfiTemp(8).aszParamIntroduce(1) = "最小运价"
    
    
    LuggageFormulaInfo = apfiTemp

End Property


'Public Sub Identify(BusID As String)
'    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
'
'End Sub

Private Sub Class_Initialize()
'    m_oDb.ConnectionString = GetConnectionStr(cszPILib)
End Sub

Public Function GetLastError() As String
    GetLastError = m_szLastError
    m_szLastError = ""
End Function

Private Sub SelfShowError(ByVal plErrNum As Long, Optional pszSource As String = "", Optional pszAddInfo As String = "", Optional ByVal pbWantOrgMsg As Boolean = False)
    On Error GoTo Error_Handle
    Dim nCount As Integer
    ShowError plErrNum, pszSource, pszAddInfo, pbWantOrgMsg
Error_Handle:
    m_szLastError = err.Description
    err.Raise err.Number
End Sub



'行包的基本运价
Public Function LuggageBaseCarriage(ByVal VehicleTypeCode As String, ByVal Mileage As Double, ByVal Weight As Double, ByVal BaggageNumber As Double, ByVal OverWeightNumber As Integer, ByVal Param1 As Double, ByVal Param2 As Double, ByVal Param3 As Double, ByVal Param4 As Double, ByVal Param5 As Double, ByVal Param6 As Double, ByVal Param7 As Double, ByVal Param8 As Double, ByVal Param9 As Double, ByVal Param10 As Double) As Double
    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
    
    'Mileage:里程
    'Weight:计重
    
    'Param1:每公里公斤费率
    'Param2:最小运价
    
    LuggageBaseCarriage = Mileage * Weight * Param1
    If LuggageBaseCarriage < Param2 Then
        LuggageBaseCarriage = Param2
    End If
    
End Function

'行包的平均计重基本运价
Public Function LuggageBaseCarriageAvg(ByVal VehicleTypeCode As String, ByVal Mileage As Double, ByVal Weight As Double, ByVal BaggageNumber As Double, ByVal OverWeightNumber As Integer, ByVal Param1 As Double, ByVal Param2 As Double, ByVal Param3 As Double, ByVal Param4 As Double, ByVal Param5 As Double, ByVal Param6 As Double, ByVal Param7 As Double, ByVal Param8 As Double, ByVal Param9 As Double, ByVal Param10 As Double) As Double
    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
    
    'Mileage:里程
    'Weight:计重
    
    'Param1:每公里公斤费率
    'Param2:最小运价
    'Param3:最小计重
    
    Dim mAvgWeight As Double
    Dim mLuggageBaseCarriageAvg As Double
    Dim mTotal As Double
    Dim i As Integer
    mAvgWeight = 0
    mLuggageBaseCarriageAvg = 0
    mTotal = 0
    For i = 1 To BaggageNumber + OverWeightNumber
        mAvgWeight = Weight / (BaggageNumber + OverWeightNumber)
'        mLuggageBaseCarriageAvg = Mileage * mAvgWeight * Param1
        If mAvgWeight < Param3 Then
            mLuggageBaseCarriageAvg = Param2
        ElseIf Mileage * mAvgWeight * Param1 <= Param2 Then
            mLuggageBaseCarriageAvg = Param2 + Mileage * (mAvgWeight - Param3) * Param1
        Else
            mLuggageBaseCarriageAvg = Mileage * mAvgWeight * Param1
        End If
        mTotal = mTotal + mLuggageBaseCarriageAvg
    Next i
    LuggageBaseCarriageAvg = mTotal
End Function




'行包的每件计算法
Public Function LuggagePerBaggageNumber(ByVal VehicleTypeCode As String, ByVal Mileage As Double, ByVal Weight As Double, ByVal BaggageNumber As Double, ByVal OverWeightNumber As Integer, ByVal Param1 As Double, ByVal Param2 As Double, ByVal Param3 As Double, ByVal Param4 As Double, ByVal Param5 As Double, ByVal Param6 As Double, ByVal Param7 As Double, ByVal Param8 As Double, ByVal Param9 As Double, ByVal Param10 As Double) As Double
    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
    
    'BaggageNumber:件数
    'OverWeightNumber:超重件数
    
    'Param1:每件费用
    'Param2:超重件数每件费用
    LuggagePerBaggageNumber = BaggageNumber * Param1 + OverWeightNumber * Param2
    
End Function

'行包的等件数计算法
'Param1 指示多少件为一个计数单位
Public Function LuggageEqBaggageNumber(ByVal VehicleTypeCode As String, ByVal Mileage As Double, ByVal Weight As Double, ByVal BaggageNumber As Double, ByVal OverWeightNumber As Integer, ByVal Param1 As Double, ByVal Param2 As Double, ByVal Param3 As Double, ByVal Param4 As Double, ByVal Param5 As Double, ByVal Param6 As Double, ByVal Param7 As Double, ByVal Param8 As Double, ByVal Param9 As Double, ByVal Param10 As Double) As Double
    'BaggageNumber:件数
    'OverWeightNumber:超重件数
    
    'Param1:计数单位
    'Param2:每个计数单位的费用
    'Param3:超重件数每件费用
    
    Dim nCount As Integer
    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
    
    nCount = Int(BaggageNumber / Param1)
    '如果不足参数1,按参数1算
    If nCount * Param1 < BaggageNumber Then
        nCount = nCount + 1
    End If
    
    LuggageEqBaggageNumber = nCount * Param1 * Param2 + OverWeightNumber * Param3
    
End Function

'行包常数公式
Public Function LuggageConstantData(ByVal VehicleTypeCode As String, ByVal Mileage As Double, ByVal Weight As Double, ByVal BaggageNumber As Double, ByVal OverWeightNumber As Integer, ByVal Param1 As Double, ByVal Param2 As Double, ByVal Param3 As Double, ByVal Param4 As Double, ByVal Param5 As Double, ByVal Param6 As Double, ByVal Param7 As Double, ByVal Param8 As Double, ByVal Param9 As Double, ByVal Param10 As Double) As Double
    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
    
    LuggageConstantData = Param1
    
End Function


Private Function LuggagePerBaggageNumberAddOver(ByVal VehicleTypeCode As String, ByVal Mileage As Double, ByVal Weight As Double, ByVal BaggageNumber As Double, ByVal OverWeightNumber As Integer, ByVal Param1 As Double, ByVal Param2 As Double, ByVal Param3 As Double, ByVal Param4 As Double, ByVal Param5 As Double, ByVal Param6 As Double, ByVal Param7 As Double, ByVal Param8 As Double, ByVal Param9 As Double, ByVal Param10 As Double) As Double
 
End Function

Private Function CUint(ByVal pdbDividend As Double, ByVal pdbDivisor As Double) As Integer
    '向上取整
    Dim nCount As Integer
    If pdbDivisor = 0 Then
        CUint = 0
        Exit Function
    End If
    nCount = Int(pdbDividend / pdbDivisor)
    '如果不足参数1,按参数1算
    If nCount * pdbDivisor < pdbDividend Then
        nCount = nCount + 1
    End If
    CUint = nCount
End Function

Public Function LuggageConstantSeg(ByVal VehicleTypeCode As String, ByVal Mileage As Double, ByVal Weight As Double, ByVal BaggageNumber As Double, ByVal OverWeightNumber As Integer, ByVal Param1 As Double, ByVal Param2 As Double, ByVal Param3 As Double, ByVal Param4 As Double, ByVal Param5 As Double, ByVal Param6 As Double, ByVal Param7 As Double, ByVal Param8 As Double, ByVal Param9 As Double, ByVal Param10 As Double) As Double
'        apfiTemp(6).szFunIntroduce = "行包分段计算公式" & Chr(13) & Chr(10) _
'        & ",行包重量取平均数计算 (平均重量 = 计重 / 件数)" & Chr(13) & Chr(10) _
'        & "0-[参数1]公斤之间(含[参数1])每件的价格为[参数2]元" & Chr(13) & Chr(10) _
'        & "[参数1]-[参数3]公斤之间(含[参数3])每件重量加[参数4]的公斤数,价格=[参数2]+[参数5]*((重量-[参数1])/[参数4])向上取整" & Chr(13) & Chr(10) _
'        & "[参数3]-[参数6]公斤之间(含[参数6])每件重量加[参数7]的公斤数,价格在[参数2]+[参数8]*((重量-[参数3])/[参数7])向上取整" & Chr(13) & Chr(10) _
'        & "超过[参数6]公斤数,价格=(重量/参数9)向上取整*参数10" & Chr(13) & Chr(10)
    
    Dim dbAvgWeight As Double
    Dim dbPrice As Double
    If BaggageNumber = 0 Then Exit Function
    
    Dim nCount As Integer
    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
    
    dbAvgWeight = Weight / BaggageNumber
    If dbAvgWeight > 0 And dbAvgWeight <= Param1 Then
        dbPrice = Param2
    ElseIf dbAvgWeight > Param1 And dbAvgWeight <= Param3 Then

        nCount = CUint(dbAvgWeight - Param1, Param4)

        dbPrice = Param2 + Param5 * nCount
        
    ElseIf dbAvgWeight > Param3 And dbAvgWeight <= Param6 Then
        
        nCount = CUint(dbAvgWeight - Param3, Param7)
        
        dbPrice = Param2 + CUint((Param3 - Param1), Param4) * Param5 + Param8 * nCount
        
    ElseIf dbAvgWeight > Param6 Then
        dbPrice = CUint(dbAvgWeight, Param9) * Param10
    End If
    
    LuggageConstantSeg = dbPrice * BaggageNumber '+ OverWeightNumber)

End Function


Public Function LuggageConstantSeg3Seg(ByVal VehicleTypeCode As String, ByVal Mileage As Double, ByVal Weight As Double, ByVal BaggageNumber As Double, ByVal OverWeightNumber As Integer, ByVal Param1 As Double, ByVal Param2 As Double, ByVal Param3 As Double, ByVal Param4 As Double, ByVal Param5 As Double, ByVal Param6 As Double, ByVal Param7 As Double, ByVal Param8 As Double, ByVal Param9 As Double, ByVal Param10 As Double) As Double

    Dim dbAvgWeight As Double
    Dim dbPrice As Double
    If BaggageNumber = 0 Then Exit Function
    
    Dim nCount As Integer
    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
    
    dbAvgWeight = Format(Weight / BaggageNumber, "0")
    If dbAvgWeight > 0 And dbAvgWeight <= Param1 Then
        dbPrice = Param2
    ElseIf dbAvgWeight > Param1 And dbAvgWeight <= Param3 Then

        nCount = CUint(dbAvgWeight - Param1, Param4)

        dbPrice = Param2 + Param5 * nCount
        
    ElseIf dbAvgWeight > Param3 Then

        dbPrice = CUint(dbAvgWeight, Param6) * Param7
    End If
    
    LuggageConstantSeg3Seg = dbPrice * BaggageNumber '+ OverWeightNumber)

End Function



'    apfiTemp(1).nFunParamCount = 0 '参数个数
'    apfiTemp(1).szFunIntroduce = "行包基本运价的介绍（公里数*重量*车型费率） 车型费率为该车辆对应的车型费率。 " '公式介绍
'    apfiTemp(1).szFunName = "LuggageBaseCarriageByVehicleType" '具体使用的方法名
'    apfiTemp(1).szFunChineseName = "基本运价（按车型费率算）" '公式中文名
'    apfiTemp(1).bBaseCarriage = True '是不是基本运价公式
   
'基本运价（按车型费率算）    : 公里数*重量*车型费率     车型费率为该车辆对应的车型费率
'行包的平均计重基本运价
Public Function LuggageBaseCarriageByVehicleType(ByVal VehicleTypeCode As String, ByVal Mileage As Double, ByVal Weight As Double, ByVal BaggageNumber As Double, ByVal OverWeightNumber As Integer, ByVal Param1 As Double, ByVal Param2 As Double, ByVal Param3 As Double, ByVal Param4 As Double, ByVal Param5 As Double, ByVal Param6 As Double, ByVal Param7 As Double, ByVal Param8 As Double, ByVal Param9 As Double, ByVal Param10 As Double) As Double
    AssertActiveUserValid m_oActiveUser, ERR_PriceItemFunLib
    
    'Mileage:里程
    'Weight:计重
    
    'Param1:最小运价
    'VehicleTypeCode为车型
    Dim dbRatio As Double
    dbRatio = GetVehicleTypeRatio(VehicleTypeCode)
    LuggageBaseCarriageByVehicleType = Mileage * Weight * dbRatio
    If LuggageBaseCarriageByVehicleType < Param1 Then
        LuggageBaseCarriageByVehicleType = Param1
    End If
End Function

'得到某车型的相应的费率
Private Function GetVehicleTypeRatio(ByVal VehicleTypeCode As String) As Double
    Dim i As Integer
    GetVehicleTypeRatio = 0
    '查找费率。
    For i = 1 To ArrayLength(m_aszAllVehicleRatio)
        If m_aszAllVehicleRatio(i, 1) = Trim(VehicleTypeCode) Then
            GetVehicleTypeRatio = m_aszAllVehicleRatio(i, 2)
            Exit Function
            
        End If
    Next i
    
    
End Function


'刷新所有的车型费率
Private Sub RefreshVehicleTypeRation()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    
    oDb.ConnectionString = GetConnectionStr()
    szSql = "SELECT * FROM vehicle_type_code "
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim m_aszAllVehicleRatio(1 To rsTemp.RecordCount, 1 To 2)
        For i = 1 To rsTemp.RecordCount
            m_aszAllVehicleRatio(i, 1) = FormatDbValue(rsTemp!vehicle_type_code)
            m_aszAllVehicleRatio(i, 2) = FormatDbValue(rsTemp!lug_base_transit_ratio)
            
            rsTemp.MoveNext
        Next i
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
    
    
End Sub










