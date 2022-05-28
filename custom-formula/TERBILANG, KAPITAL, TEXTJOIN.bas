Attribute VB_Name = "Module5"
'Fungsi penterjemahan masing-masing angka
Private Function KeKata(Nomor)
TrjKata = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan")
KeKata = TrjKata(Nomor)
End Function

'Mulai penulisan Fungsi Terbilang
Public Function terbilang(Nilai_Angka, Optional Style = 4, Optional Satuan = "")
Angka = Fix(Abs(Nilai_Angka))
'Desimal dibelakang koma
des1 = Mid(Abs(Nilai_Angka), Len(Angka) + 2, 1)
des2 = Mid(Abs(Nilai_Angka), Len(Angka) + 3, 1)

If des2 = "" Then
    If des1 = "" Or des1 = "0" Then
    Koma = ""
    Else
    Koma = " Koma " & KeKata(des1)
    End If
ElseIf des2 = "0" Then
    If des1 = "0" Then
    Koma = ""
    ElseIf des1 = "1" Then
    Koma = " Koma Sepuluh"
    Else
    Koma = " Koma " & KeKata(des1) & " Puluh"
    End If
Else
    If des1 = "0" Then
    Koma = " koma nol " & KeKata(des2)
    ElseIf des1 = "1" Then
        If des2 = "1" Then
        Koma = " Koma Sebelas"
        Else
        Koma = " Koma " & KeKata(des2) & " Belas"
        End If
    Else
    Koma = " Koma " & KeKata(des1) & " Puluh " & KeKata(des2)
    End If
End If
'Misahin Angka
No1 = Left(Right(Angka, 1), 1)
No2 = Left(Right(Angka, 2), 1)
No3 = Left(Right(Angka, 3), 1)
No4 = Left(Right(Angka, 4), 1)
No5 = Left(Right(Angka, 5), 1)
No6 = Left(Right(Angka, 6), 1)
No7 = Left(Right(Angka, 7), 1)
No8 = Left(Right(Angka, 8), 1)
No9 = Left(Right(Angka, 9), 1)
No10 = Left(Right(Angka, 10), 1)
No11 = Left(Right(Angka, 11), 1)
No12 = Left(Right(Angka, 12), 1)
No13 = Left(Right(Angka, 13), 1)
No14 = Left(Right(Angka, 14), 1)
No15 = Left(Right(Angka, 15), 1)
'Satuan
If Len(Angka) >= 1 Then
    If Len(Angka) = 1 And No1 = 1 Then
    Nomor1 = "Satu"
    ElseIf Len(Angka) = 1 And No1 = 0 Then
    Nomor1 = "Nol"
    ElseIf No2 = "1" Then
        If No1 = "1" Then
        Nomor1 = "Sebelas"
        ElseIf No1 = "0" Then
        Nomor1 = "Sepuluh"
        Else
        Nomor1 = KeKata(No1) & " Belas"
        End If
    
    Else
    Nomor1 = KeKata(No1)
    End If
Else
Nomor1 = ""
End If

'Puluhan
If Len(Angka) >= 2 Then
    If No2 = 1 Or No2 = "0" Then
    Nomor2 = ""
    Else
    Nomor2 = KeKata(No2) & " Puluh "
    End If
Else
Nomor2 = ""
End If
'Ratusan
If Len(Angka) >= 3 Then
    If No3 = "1" Then
    Nomor3 = "seratus "
    ElseIf No3 = "0" Then
    Nomor3 = ""
    Else
    Nomor3 = KeKata(No3) & " Ratus "
    End If
Else
Nomor3 = ""
End If
'Ribuan
If Len(Angka) >= 4 Then
    If No6 = "0" And No5 = "0" And No4 = "0" Then
    Nomor4 = ""
    ElseIf (No4 = "1" And Len(Angka) = 4) Or (No6 = "0" And No5 = "0" And No4 = "1") Then
    Nomor4 = "seribu "
    ElseIf No5 = "1" Then
        If No4 = "1" Then
        Nomor4 = "sebelas ribu "
        ElseIf No4 = "0" Then
        Nomor4 = "sepuluh ribu "
        Else
        Nomor4 = KeKata(No4) & " belas ribu "
        End If

    Else
    Nomor4 = KeKata(No4) & " ribu "
    End If
Else
Nomor4 = ""
End If
'Puluhan ribu
If Len(Angka) >= 5 Then
    If No5 = "1" Or No5 = "0" Then
    Nomor5 = ""
    Else
    Nomor5 = KeKata(No5) & " puluh "
    End If
Else
Nomor5 = ""
End If
'Ratusan Ribu
If Len(Angka) >= 6 Then
    If No6 = "1" Then
    Nomor6 = "seratus "
    ElseIf No6 = "0" Then
    Nomor6 = ""
    Else
    Nomor6 = KeKata(No6) & " ratus "
    End If
Else
Nomor6 = ""
End If
'Jutaan
If Len(Angka) >= 7 Then
    If No9 = "0" And No8 = "0" And No7 = "0" Then
    Nomor7 = ""
    ElseIf No7 = "1" And Len(Angka) = 7 Then
    Nomor7 = "satu juta "
    ElseIf No8 = "1" Then
        If No7 = "1" Then
        Nomor7 = "sebelas juta "
        ElseIf No7 = "0" Then
        Nomor7 = "sepuluh juta "
        Else
        Nomor7 = KeKata(No7) & " belas juta "
        End If

    Else
    Nomor7 = KeKata(No7) & " juta "
    End If
Else
Nomor7 = ""
End If
'Puluhan juta
If Len(Angka) >= 8 Then
    If No8 = "1" Or No8 = "0" Then
    Nomor8 = ""
    Else
    Nomor8 = KeKata(No8) & " puluh "
    End If
Else
Nomor8 = ""
End If
'Ratusan juta
If Len(Angka) >= 9 Then
    If No9 = "1" Then
    Nomor9 = "seratus "
    ElseIf No9 = "0" Then
    Nomor9 = ""
    Else
    Nomor9 = KeKata(No9) & " ratus "
    End If
Else
Nomor9 = ""
End If
'Milyar
If Len(Angka) >= 10 Then
    If No12 = "0" And No11 = "0" And No10 = "0" Then
    Nomor10 = ""
    ElseIf No10 = "1" And Len(Angka) = 10 Then
    Nomor10 = "satu milyar "
    ElseIf No11 = "1" Then
        If No10 = "1" Then
        Nomor10 = "sebelas milyar "
        ElseIf No10 = "0" Then
        Nomor10 = "sepuluh milyar "
        Else
        Nomor10 = KeKata(No10) & " belas milyar "
        End If

    Else
    Nomor10 = KeKata(No10) & " milyar "
    End If
Else
Nomor10 = ""
End If
'Puluhan Milyar
If Len(Angka) >= 11 Then
    If No11 = "1" Or No11 = "0" Then
    Nomor11 = ""
    Else
    Nomor11 = KeKata(No11) & " puluh "
    End If
Else
Nomor11 = ""
End If
'Ratusan Milyar
If Len(Angka) >= 12 Then
    If No12 = "1" Then
    Nomor12 = "seratus "
    ElseIf No12 = "0" Then
    Nomor12 = ""
    Else
    Nomor12 = KeKata(No12) & " ratus "
    End If
Else
Nomor12 = ""
End If
'Triliun
If Len(Angka) >= 13 Then
    If No15 = "0" And No14 = "0" And No13 = "0" Then
    Nomor13 = ""
    ElseIf No13 = "1" And Len(Angka) = 13 Then
    Nomor13 = "satu triliun "
    ElseIf No14 = "1" Then
        If No13 = "1" Then
        Nomor13 = "sebelas triliun "
        ElseIf No13 = "0" Then
        Nomor13 = "sepuluh triliun "
        Else
        Nomor13 = KeKata(No13) & " belas triliun "
        End If

    Else
    Nomor13 = KeKata(No13) & " triliun "
    End If
Else
Nomor13 = ""
End If
'Puluhan triliun
If Len(Angka) >= 14 Then
    If No14 = "1" Or No14 = "0" Then
    Nomor14 = ""
    Else
    Nomor14 = KeKata(No14) & " puluh "
    End If
Else
Nomor14 = ""
End If
'Ratusan triliun
If Len(Angka) >= 15 Then
    If No15 = "1" Then
    Nomor15 = "seratus "
    ElseIf No15 = "0" Then
    Nomor15 = ""
    Else
    Nomor15 = KeKata(No15) & " ratus "
    End If
Else
Nomor15 = ""
End If

If Len(Angka) > 15 Then
bilang = "Digit Angka Terlalu Banyak"
Else
    If IsNull(Nilai_Angka) Then
    bilang = ""
    ElseIf Nilai_Angka < 0 Then
    bilang = "minus " & Trim(Nomor15 & Nomor14 & Nomor13 & Nomor12 & Nomor11 & Nomor10 & Nomor9 & Nomor8 & Nomor7 _
    & Nomor6 & Nomor5 & Nomor4 & Nomor3 & Nomor2 & Nomor1 & Koma & " " & Satuan)
    Else
    bilang = Trim(Nomor15 & Nomor14 & Nomor13 & Nomor12 & Nomor11 & Nomor10 & Nomor9 & Nomor8 & Nomor7 _
    & Nomor6 & Nomor5 & Nomor4 & Nomor3 & Nomor2 & Nomor1 & Koma & " " & Satuan)
    End If
End If
If Style = 4 Then
terbilang = StrConv(Left(bilang, 1), 1) & StrConv(Mid(bilang, 2, 1000), 2)
Else
terbilang = StrConv(bilang, Style)
End If
terbilang = Replace(terbilang, "  ", " ", 1, 1000, vbTextCompare)

End Function

Public Function Kapital(sWord As String) As String
    Kapital = StrConv(sWord, vbProperCase)
End Function

Public Function Textjoin(Sep As String, ParamArray TxtRng() As Variant) As String
On Error Resume Next
'Sep is the separator, set to "" if you don't want any separator. Separator must be string or single cell, not cell range
'TxtRng is the content you want to join. TxtRng can be string, single cell, cell range or array returned from an array function. Empty content will be ignored
Dim OutStr As String 'the output string
Dim i, j, k, l As Integer 'counters
Dim FinArr(), element As Variant 'the final array and a temporary element when transfering between the two arrays

'Go through each item of TxtRng(),  depending on the item type, transform and put it into FinArray()
i = 0 'the counter for TxtRng
j = 0 'the counter for FinArr
k = 0: l = 0 'the counters for the case of array from Excel array formula
Do While i < UBound(TxtRng) + 1
    If TypeName(TxtRng(i)) = "String" Then 'specified string like "t"
        ReDim Preserve FinArr(0 To j)
        FinArr(j) = "blah"
        FinArr(j) = TxtRng(i)
        j = j + 1
    ElseIf TypeName(TxtRng(i)) = "Range" Then 'single cell or range of cell like A1, A1:A2
        For Each element In TxtRng(i)
            ReDim Preserve FinArr(0 To j)
            FinArr(j) = element
            j = j + 1
        Next
    ElseIf TypeName(TxtRng(i)) = "Variant()" Then 'array returned from an Excel array formula
         For k = LBound(TxtRng(0), 1) To UBound(TxtRng(0), 1)
            For l = LBound(TxtRng(0), 2) To UBound(TxtRng(0), 2)
                ReDim Preserve FinArr(0 To j)
                FinArr(j) = TxtRng(0)(k, l)
                j = j + 1
            Next
         Next
    Else
        Textjoin = CVErr(xlErrValue)
        Exit Function
    End If
i = i + 1
Loop

'Put each element of the new array into the join string
For i = LBound(FinArr) To UBound(FinArr)
    If FinArr(i) <> "" Then 'Remove this line if you want to include empty strings
    OutStr = OutStr & FinArr(i) & Sep
    End If
Next
 Textjoin = Left(OutStr, Len(OutStr) - Len(Sep)) 'remove the ending separator

End Function
