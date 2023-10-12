Option Explicit
Dim strArray, i, strText
strArray = Array("みかん", "りんご", "バナナ")
For i = 0 to Ubound(strArray)
    strText = strArray(i)
	MsgBox strText, , "配列データの表示"
Next
