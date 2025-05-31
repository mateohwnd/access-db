Attribute VB_Name = "mdlCurrencytoWords"
Option Compare Database
Option Explicit

Function CurrencyToWords(ByVal MyNumber)

  Dim Temp
         Dim Pesos, Centavos
         Dim DecimalPlace, count
 
         ReDim Place(9) As String
         Place(2) = " Thousand "
         Place(3) = " Million "
         Place(4) = " Billion "
 
         
' Convert MyNumber to a string, trimming extra spaces.

         MyNumber = Trim(Str(MyNumber))
 
         
' Find decimal place.

         DecimalPlace = InStr(MyNumber, ".")
 
         
' If we find decimal place...

         If DecimalPlace > 0 Then
            
' Convert Centavos

            Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            
' Hi! Note the above line Mid function it gives right portion

            
' after the decimal point

            
'if only . and no numbers such as 789. accures, mid returns nothing

            
' to avoid error we added 00

            
' Left function gives only left portion of the string with specified places here 2

 
 
            Centavos = ConvertTens(Temp)
 
 
            
' Strip off Centavos from remainder to convert.

            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
         End If
 
         count = 1
        If MyNumber <> "" Then
 
            
' Convert last 3 digits of MyNumber to Indian Pesos.

            Temp = ConvertHundreds(Right(MyNumber, 3))
 
            If Temp <> "" Then Pesos = Temp & Place(count) & Pesos
 
            If Len(MyNumber) > 3 Then
               
' Remove last 3 converted digits from MyNumber.

               MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
               MyNumber = ""
            End If
 
        End If
 
            
' convert last two digits to of mynumber

            count = 2
 
            Do While MyNumber <> ""
            Temp = ConvertTens(Right("0" & MyNumber, 2))
 
            If Temp <> "" Then Pesos = Temp & Place(count) & Pesos
            If Len(MyNumber) > 2 Then
               
' Remove last 2 converted digits from MyNumber.

               MyNumber = Left(MyNumber, Len(MyNumber) - 2)
 
            Else
               MyNumber = ""
            End If
            count = count + 1
 
            Loop
 
         
' Clean up Pesos. (extension)

        Select Case Pesos
            Case ""
               Pesos = ""
            Case "One"
               Pesos = "One Peso"
            Case Else
               Pesos = Pesos & " Pesos"
         End Select
 
         
' Clean up Centavos.

         Select Case Centavos
            Case ""
               Centavos = ""
            Case "One"
               Centavos = "One Centavo"
            Case Else
               Centavos = Centavos & " Centavos"
         End Select
         
         If Pesos = "" Then
         CurrencyToWords = Centavos & " Only"
         ElseIf Centavos = "" Then
         CurrencyToWords = Pesos & " Only"
         Else
         CurrencyToWords = Pesos & " and " & Centavos & " Only"
         End If
         
End Function

Private Function ConvertDigit(ByVal MyDigit)

        Select Case val(MyDigit)
            Case 1: ConvertDigit = "One"
            Case 2: ConvertDigit = "Two"
            Case 3: ConvertDigit = "Three"
            Case 4: ConvertDigit = "Four"
            Case 5: ConvertDigit = "Five"
            Case 6: ConvertDigit = "Six"
            Case 7: ConvertDigit = "Seven"
            Case 8: ConvertDigit = "Eight"
            Case 9: ConvertDigit = "Nine"
            Case Else: ConvertDigit = ""
         End Select

End Function

Private Function ConvertHundreds(ByVal MyNumber)
    Dim result As String


' Exit if there is nothing to convert.

         If val(MyNumber) = 0 Then Exit Function


' Append leading zeros to number.

         MyNumber = Right("000" & MyNumber, 3)


' Do we have a hundreds place digit to convert?

         If Left(MyNumber, 1) <> "0" Then
            result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
         End If


' Do we have a tens place digit to convert?

         If Mid(MyNumber, 2, 1) <> "0" Then
            result = result & ConvertTens(Mid(MyNumber, 2))
         Else

' If not, then convert the ones place digit.

            result = result & ConvertDigit(Mid(MyNumber, 3))
         End If

         ConvertHundreds = Trim(result)
End Function


Private Function ConvertTens(ByVal MyTens)
          Dim result As String


' Is value between 10 and 19?

         If val(Left(MyTens, 1)) = 1 Then
            Select Case val(MyTens)
               Case 10: result = "Ten"
               Case 11: result = "Eleven"
               Case 12: result = "Twelve"
               Case 13: result = "Thirteen"
               Case 14: result = "Fourteen"
               Case 15: result = "Fifteen"
               Case 16: result = "Sixteen"
               Case 17: result = "Seventeen"
               Case 18: result = "Eighteen"
               Case 19: result = "Nineteen"
               Case Else
            End Select
         Else

' .. otherwise it's between 20 and 99.

            Select Case val(Left(MyTens, 1))
               Case 2: result = "Twenty "
               Case 3: result = "Thirty "
               Case 4: result = "Forty "
               Case 5: result = "Fifty "
               Case 6: result = "Sixty "
               Case 7: result = "Seventy "
               Case 8: result = "Eighty "
               Case 9: result = "Ninety "
               Case Else
            End Select


' Convert ones place digit.

            result = result & ConvertDigit(Right(MyTens, 1))
         End If

         ConvertTens = result
End Function
