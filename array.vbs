Option Explicit
Dim strArray, i, strText
strArray = Array("�݂���", "���", "�o�i�i")
For i = 0 to Ubound(strArray)
    strText = strArray(i)
	MsgBox strText, , "�z��f�[�^�̕\��"
Next
