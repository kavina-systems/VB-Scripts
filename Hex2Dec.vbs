' =========================================================
' Tittle: Hex2Dec
' Version : v0.1
' Date: 09/09/2009
' Author: K Grey
' Comapny : KSL
' Comment: Convert hexadecimal value and display as decimal
'		   This converts integer values only and ignores the sign bit
' Version History  :
' =========================================================

HexNumber = "0000000C" '12
HexNumber = "000000FF" '255
HexNumber = "00000100" '256
'HexNumber = "0001E240" '123456

Num = HexNumber

Result = 0
Factor = 1
Pos = Len(Num)
Converted = ""

Do 
	Part = Mid(Num,Pos,1) 'get each character in turn
	valid = 1
	
	'convert decimal to Hex A-F
	Select Case Part
		Case "0"
			Result = Result + (0 * Factor)
		Case "1"
			Result = Result + (1 * Factor)
		Case "2"
			Result = Result + (2 * Factor)
		Case "3"
			Result = Result + (3 * Factor)
		Case "4"
			Result = Result + (4 * Factor)
		Case "5"
			Result = Result + (5 * Factor)
		Case "6"
			Result = Result + (6 * Factor)
		Case "7"
			Result = Result + (7 * Factor)
		Case "8"
			Result = Result + (8 * Factor)
		Case "9"
			Result = Result + (9 * Factor)
		Case "A","a"
			Result = Result + (10 * Factor)
		Case "B","b"
			Result = Result + (11 * Factor)
		Case "C","c"
			Result = Result + (12 * Factor)
		Case "D","d"
			Result = Result + (13 * Factor)
		Case "E","e"
			Result = Result + (14 * Factor)
		Case "F","f"
			Result = Result + (15 * Factor)
		Case Else
			valid = 0
	End select
	
	'adjust multiplication factor for each nibble
	If valid then
		For n = 1 To 4
			Factor = Factor * 2
		Next
		Converted = Part & Converted
	End if
	
	Pos = Pos -1 'decrement number string position
	
Loop Until Pos = 0 


Debug.Write "Hexadecimal to Decimal Converion" & vbCrLf
Debug.Write "Hexadecimal Value Provided " & HexNumber & vbCrLf
Debug.Write "Hexadecimal Value Converted " & Converted & vbCrLf
Debug.Write "Decimal Conversion " & Result & vbCrLf
