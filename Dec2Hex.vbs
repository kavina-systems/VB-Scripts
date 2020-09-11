' =========================================================
' Tittle: Dec2Hex
' Version : v0.1
' Date: 09/09/2009
' Author: K Grey
' Comapny : KSL
' Comment: Convert decimal value and display as hexadecimal
'		   This converts integer values only and strips the sign bit
'		   Note vb has a Hex() function that converts decimal to Hex
' Version History  :
' =========================================================

DecNumber = 12 '0000 000C
DecNumber = 255 '0000 00FF
DecNumber = 256 '0000 0100
'DecNumber = 12245 '0000 3039

Num = DecNumber

'strip sign bit
neg = 0
If Num <= 0 Then
	Num = Num * -1
	neg = 1
End If

Result = ""
Part = 0
Pass = 0

Do 
	'get loweset nibble and convert
	Part = Num And &hF
	
	'convert decimal to Hex A-F
	Select Case Part
		Case 10
			Result = "A" & Result
		Case 11
			Result = "B" & Result
		Case 12
			Result = "C" & Result
		Case 13
			Result = "D" & Result
		Case 14
			Result = "E" & Result
		Case 15
			Result = "F" & Result
		Case Else
			Result = Part & Result
	End select
	
	'shift number one nibble
	For n = 1 To 4
		Num = Int(Num / 2)
	Next
	Pass = Pass +1 'increment number of times to process nibbles
	
	If Pass = 4 Then
		Result = " " & Result
	End If
	
Loop Until Pass = 8 '32 bit number has 8 nibbles


Debug.Write "Decimal to Hexadecimal Converion" & vbCrLf
Debug.Write "Decimal Value " & DecNumber & vbCrLf
If neg Then
	Debug.Write "Sign Bit Has Been Removed" & vbCrLf
End If
Debug.Write "Hexadecimal Conversion " & Result & vbCrLf
