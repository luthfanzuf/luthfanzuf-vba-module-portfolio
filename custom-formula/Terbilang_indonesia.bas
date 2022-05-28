Attribute VB_Name = "Module1"
' This function replaces the standard Excel MOD function
Function XLMod(a, b)
    XLMod = a - b * Fix(a / b)
End Function


Function terbilang(ByVal n As Double) As String
Dim satuan As Variant
satuan = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas")
Select Case n 'keadaan
Case 0 To 11
terbilang = " " + satuan(Fix(n))
Case 12 To 19
terbilang = terbilang(XLMod(n, 10)) + " Belas"
Case 20 To 99
terbilang = terbilang(Fix(n / 10)) + " Puluh" + terbilang(XLMod(n, 10))
Case 100 To 199
terbilang = " Seratus" + terbilang(n - 100)
Case 200 To 999
terbilang = terbilang(Fix(n / 100)) + " Ratus" + terbilang(XLMod(n, 100))
Case 1000 To 1999
terbilang = " Seribu" + terbilang(n - 1000)
Case 2000 To 999999
terbilang = terbilang(Fix(n / 1000)) + " Ribu" + terbilang(XLMod(n, 1000))
Case 1000000 To 999999999
terbilang = terbilang(Fix(n / 1000000)) + " Juta" + terbilang(XLMod(n, 1000000))
Case 1000000000 To 999999999999#
terbilang = terbilang(Fix(n / 1000000000)) + " Milyar" + terbilang(XLMod(n, 1000000000))
Case Else
terbilang = terbilang(Fix(n / 1000000000000#)) + " Triliun" + terbilang(XLMod(n, 1000000000000#))
End Select
End Function

Sub Panggil_func_terbilang()
nilai = ActiveCell.Value
nilai = terbilang(nilai)
ActiveCell.Offset(0, 1).Value = nilai
End Sub



