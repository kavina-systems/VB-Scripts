' =========================================================
' Tittle: Dec2Bin
' Version : v0.1
' Date: 09/09/2009
' Author: K Grey
' Comapny : KSL
' Comment: Convert decimal value and display as binary
'	 	   Binary number to be displayed in blocks of 4 bits
'		   This converts integer values only and strips the sign bit
' Version History  :
' =========================================================

DecNumber = 12 '0000 0000 0000 0000 0000 0000 0000 1100
DecNumber = 255 '0000 0000 0000 0000 0000 0000 1111 1111
DecNumber = 256 '0000 0000 0000 0000 0000 0001 0000 0000
DecNumber = 1112345 '0000 0000 0000 0000 0011 0000 0011 1001

Num = DecNumber

'strip sign bit
neg = 0
If Num <= 0 Then
	Num = Num * -1
	neg = 1
End If

bits = 0
'determine the number of bits required in conversion
nbits = Num
Do while nbits >= 1
	nbits = Int(nbits / 2)
	bits = bits +1
Loop 
bits = bits -1 
'adjust number of bits to be rounded to block of 4
r = bits Mod 4
bits = bits + (4 - r)
bits = bits -1 

'bits = 31 'set bits to fixed number if static display value required i.e. 31,15,7 etc

Result = ""

For n = bits To 0 Step -1
	If Num >= 2 ^ n Then
		Result = Result & "1"
		Num = Num - 2 ^ n
	Else
		Result = Result & "0"
	End If
	
	'add seperators every 4 bits
	Select Case n
		Case 4,8,12,16,20,24,28
		Result = Result & " "
	End Select
next

Debug.Write "Decimal to Binary Converion" & vbCrLf
Debug.Write "Decimal Value " & DecNumber & vbCrLf
If neg Then
	Debug.Write "Sign Bit Has Been Removed" & vbCrLf
End If
Debug.Write "                  MSB" & vbCrLf
Debug.Write "Binary Conversion " & Result & vbCrLf