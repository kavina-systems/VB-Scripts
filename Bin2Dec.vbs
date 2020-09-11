' =========================================================
' Tittle: Bin2Dec
' Version : v0.1
' Date: 09/09/2009
' Author: K Grey
' Comapny : KSL
' Comment: Convert binary value and display as decimal
'		   This converts integer values only and ignores the sign bit
' Version History  :
' =========================================================

BinNumber = "1100" '12
BinNumber = "10000" '16
BinNumber = "0000 0000 0000 0000 0000 0001 0000 0000" '256
'BinNumber = "0000 0000 0000 0000 0011 0000 0011 1001" '12345

Num = BinNumber

Result = 0
Factor = 1
Pos = Len(Num)
Converted = ""

Do 
	Part = Mid(Num,Pos,1) 'get each character in turn
	
	If Part = "1" Or Part = "0" Then
		If Part = "1" then
			Result = Result + Factor
		End if
		Factor = Factor * 2 'adjust multiplication factor for each valid bit
		Converted = Part & Converted
	End if
	
	Pos = Pos -1 'decrement number string position
	
Loop Until Pos = 0 


Debug.Write "Binary to Decimal Converion" & vbCrLf
Debug.Write "Binary Value Provided " & BinNumber & vbCrLf
Debug.Write "Binary Value Converted " & Converted & vbCrLf
Debug.Write "Decimal Conversion " & Result & vbCrLf
