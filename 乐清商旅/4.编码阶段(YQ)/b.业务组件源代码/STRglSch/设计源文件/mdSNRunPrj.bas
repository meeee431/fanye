Attribute VB_Name = "mdSNRunPrj"
Option Explicit

'内部用得到总票价
Public Function SelfGetTotalPrice(prsPriceInfo As Recordset) As Double
    Dim sgTemp As Double

    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!base_carriage)
    
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_1)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_2)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_3)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_4)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_5)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_6)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_7)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_8)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_9)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_10)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_11)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_12)    ' 改
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_13)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_14)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_15)
    SelfGetTotalPrice = sgTemp
End Function
