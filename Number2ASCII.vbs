' =========================================================
' Tittle: Number2ASCII
' Version : v0.1
' Date: 13/11/2014
' Author: K Grey
' Comapny : KSL
' Comment: Convert a simple two digit number to ASCII and join to a string
' Version History  :
' =========================================================

Dim units, tens, result, offset

offset = 48 'ascii character for 0
result = ""
For tens = 0 + offset To 9 + offset
	For units = 0 + offset To 9 + offset
		result = "Vis_" + Chr(tens) + Chr(units)
		Debug.Write result & vbcrlf
	Next 
next