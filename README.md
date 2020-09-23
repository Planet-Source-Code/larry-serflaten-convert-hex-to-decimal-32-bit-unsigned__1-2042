<div align="center">

## Convert Hex to Decimal \(32\-bit Unsigned\)


</div>

### Description

Converts Hex [0 - FFFFFFFF] to Decimal [0 - 4294967295] using Currency type to avoid the sign bit.
 
### More Info
 
A valid 1-8 character Hex String

A Currency value in the range of 0 - 4294967295


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Larry Serflaten](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/larry-serflaten.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/larry-serflaten-convert-hex-to-decimal-32-bit-unsigned__1-2042/archive/master.zip)





### Source Code

```
Function ConvertHex (H$) As Currency
Dim Tmp$
Dim lo1 As Integer, lo2 As Integer
Dim hi1 As Long, hi2 As Long
Const Hx = "&H"
Const BigShift = 65536
Const LilShift = 256, Two = 2
  Tmp = H
  'In case "&H" is present
  If UCase(Left$(H, 2)) = "&H" Then Tmp = Mid$(H, 3)
  'In case there are too few characters
  Tmp = Right$("0000000" & Tmp, 8)
  'In case it wasn't a valid number
  If IsNumeric(Hx & Tmp) Then
    lo1 = CInt(Hx & Right$(Tmp, Two))
    hi1 = CLng(Hx & Mid$(Tmp, 5, Two))
    lo2 = CInt(Hx & Mid$(Tmp, 3, Two))
    hi2 = CLng(Hx & Left$(Tmp, Two))
    ConvertHex = CCur(hi2 * LilShift + lo2) * BigShift + (hi1 * LilShift) + lo1
  End If
End Function
```

