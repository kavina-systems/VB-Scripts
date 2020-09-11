' =========================================================
' Tittle: Hex2Float
' Version : v0.1
' Date: 06/09/2009
' Author: K Grey
' Comapny : KSL
' Comment: IEEE-754 Floating-Point Conversion From 32-bit 
'		   Hexadecimal Representation
'
'	 Bits 
'	  31  30 29 28 27 26 25 24 23 22 21 20 19 18 17 16 15 14 13 12 11 10  9  8  7  6  5  4  3  2  1  0  
'	 +---+-----------------------+---------------------------------------------------------------------+ 
'	 | S | Exponent value        |                  Mantissa value                                     | 
'	 +---+-----------------------+---------------------------------------------------------------------+ 
'	  S = Sign Bit 
'
' 		   Could be done using Python  
' 		   struct.unpack('!f', '\x41\xc4\x1d\xc3')
' Version History  :
' =========================================================

'number to be converted
'HexNumber = &h41480000 '12.5
'HexNumber = &h3E800000 '0.25
'HexNumber = &hC1480000 '-12.5
HexNumber = &hBE800000 '-0.25

'work out exponent get bits 23-30
'shift data along to clear the mantissa and leave the exponent & sign bit
exponent = HexNumber
For n = 0 To 22
	exponent = Int(exponent / 2)
Next
exponent = exponent And &hff 'mask of sign bit
exponent = exponent - 127
exponent = 2 ^ exponent

'check for neg value, sign bit31 true
If (Hexnumber And &h80000000) Then
	exponent = exponent * -1 'if neg value set exponent to neg
End If

'work out Mantissa bits 0-22
'check each bit inturn and store result
decimalpart = 1 
mantissa = 1
For n = 22 To 0 Step -1 'work in reverse order bits 22 > 0
	decimalpart = decimalpart / 2 'half decimal part each time for bit set test
    testbit = 2 ^ n 'work out what bit requires testing
    If (Hexnumber And testbit) Then 'determine if bit under test is true
    	mantissa = mantissa + decimalpart 'compute new mantissa
    End If
Next 

'compute result
Result = exponent * mantissa

Debug.Write "IEEE-754 Floating-Point Conversion From 32-bit Hexadecimal" & vbCrLf
Debug.Write "Hexadecimal " & Hex(HexNumber) & vbcrlf
Debug.Write "Decimal Value " & Result & vbCrLf
