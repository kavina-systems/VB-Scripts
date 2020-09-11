' =========================================================
' Tittle: Float2Hex
' Version : v0.1
' Date: 06/09/2009
' Author: K Grey
' Comapny : KSL
' Comment: IEEE-754 Hexadecimal Conversion From 32-bit 
'		   Decimal Floating-Point Representation
'
'	 Bits 
'	  31  30 29 28 27 26 25 24 23 22 21 20 19 18 17 16 15 14 13 12 11 10  9  8  7  6  5  4  3  2  1  0  
'	 +---+-----------------------+---------------------------------------------------------------------+ 
'	 | S | Exponent value        |                  Mantissa value                                     | 
'	 +---+-----------------------+---------------------------------------------------------------------+ 
'	  S = Sign Bit 
'
' Version History  :
' =========================================================

'number to be converted 
'DecNumber = 12.5 '41480000
'DecNumber = 0.25 '3E800000
'DecNumber = -12.5 'C1480000
DecNumber = -0.25 'BE800000

'convert ng number to pos
If DecNumber <= 0 Then
	decimalpart = DecNumber * -1
	neg = 1
Else 
	decimalpart = DecNumber
	neg = 0 
End If

'work out exponent get bits 23-30
exponent = 0
Do until decimalpart < 1 
	decimalpart = decimalpart / 2
	exponent = exponent +1
Loop
exponent = exponent + 127 'add bias

'compute mantissa
bit = 22 'first to be tested, mantissa bits 22-0
n = 0.5
complete = 0
mantissa = 0 'reset all bits
Do Until decimalpart = complete
	If complete + n  <= decimalpart  Then
		complete = complete + n
		mantissa = mantissa Or 2 ^ bit 'set corressponding bit
	End	if
	n = n /2
	bit = bit -1 'decrement to next bit
Loop
'the mantissa is normalized, by moving the bit patterns to the left 
'(each shift subtracts one from the exponent value) till the first 1 drops off
Do 
	mantissa = mantissa * 2 'shift to left
	exponent = exponent - 1 'adjust for mantissa
Loop Until mantissa And &h800000 'check bit 23 for roller over from final mantissa bit 22
mantissa = mantissa And &h7fffff 'clear rollover bits

'shift exponent to the left to 23 bit
For n = 0 To 22
	exponent = exponent * 2
Next

'compute result, result will be displayed as a decimal
Result = exponent Or mantissa ' join the two together
If neg = 1 Then
	Result = Result Or &h80000000  'set neg flag
End If
HexNumber = Hex(result) 'convert decimal into hex string

Debug.Write "IEEE-754 Floating-Point Conversion From 32-bit Hexadecimal Representation" & vbCrLf
Debug.Write "Decimal Value " & DecNumber & vbCrLf
Debug.Write "Hexadecimal " & Hexnumber & vbcrlf
