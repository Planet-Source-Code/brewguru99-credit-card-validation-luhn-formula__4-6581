<div align="center">

## Credit Card Validation \(Luhn Formula\)


</div>

### Description

Uses the Luhn formula to quickly validate a credit card. Basically all the digits except for the last one are summed together and the output is a single digit (0 to 9). This digit is compared with the last digit ensure a proper credit card number is entered (Does not actually confirm that is is a real number, just that it is likely to be one. Example: Entering "4000-0000-0000-0002" will pass the check, but "4000-0000-0000-0003" will not pass.)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BrewGuru99](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brewguru99.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__4-29.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brewguru99-credit-card-validation-luhn-formula__4-6581/archive/master.zip)





### Source Code

```
Function CheckCC(CCNo)
 Dim i, w, x, y
 y = 0
 CCNo = Replace(Replace(Replace(CStr(CCNo), "-", ""), " ", ""), ".", "") 'Ensure proper format of the input
 'Process digits from right to left, drop last digit if total length is even
 w = 2 * (Len(CCNo) Mod 2)
 For i = Len(CCNo) - 1 To 1 Step -1
  x = Mid(CCNo, i, 1)
  If IsNumeric(x) Then
   Select Case (i Mod 2) + w
    Case 0, 3 'Even Digit - Odd where total length is odd (eg. Visa vs. Amx)
     y = y + CInt(x)
    Case 1, 2 'Odd Digit - Even where total length is odd (eg. Visa vs. Amx)
     x = CInt(x) * 2
     If x > 9 Then
      'Break the digits (eg. 19 becomes 1 + 9)
      y = y + (x \ 10) + (x - 10)
     Else
      y = y + x
     End If
   End Select
  End If
 Next
 'Return the 10's complement of the total
 y = 10 - (y Mod 10)
 If y > 9 Then y = 0
 CheckCC = (CStr(y) = Right(CCNo, 1))
End Function
```

