Attribute VB_Name = "Module2"
Sub ADD_SUBTOTAL_3()
Attribute ADD_SUBTOTAL_3.VB_ProcData.VB_Invoke_Func = "w\n14"
    
    ' subtotal muat
    Dim e As Range, f As Range, g As Range, u As String
    Set e = Columns("Q").Rows(4) 'any column and/or start row you like
    If e = "" Then Set e = e.End(4)
    Do
    Set g = e
    If e.Offset(1) <> "" Then Set e = e.End(4)
    u = Range(g, e).Address(0, 0)
    Set f = e.Offset(1)
    Set e = e.End(4)
    f.Formula = "=subtotal(9, " & u & ")"
    ' f.Offset(, 1) = "sum " & u  'can omit this line if you want
    Loop Until e.Row = Rows.Count
    
    ' subtotal bongkar
    Dim e2 As Range, f2 As Range, g2 As Range, u2 As String
    Set e2 = Columns("T").Rows(4) 'any column and/or start row you like
    If e2 = "" Then Set e2 = e2.End(4)
    Do
    Set g2 = e2
    If e2.Offset(1) <> "" Then Set e2 = e2.End(4)
    u2 = Range(g2, e2).Address(0, 0)
    Set f2 = e2.Offset(1)
    Set e2 = e2.End(4)
    f2.Formula = "=subtotal(9, " & u2 & ")"
    ' f2.Offset(, 1) = "sum " & u2  'can omit this line if you want
    Loop Until e2.Row = Rows.Count
    
    ' subtotal price
    Dim e3 As Range, f3 As Range, g3 As Range, u3 As String
    Set e3 = Columns("Y").Rows(4) 'any column and/or start row you like
    If e3 = "" Then Set e3 = e3.End(4)
    Do
    Set g3 = e3
    If e3.Offset(1) <> "" Then Set e3 = e3.End(4)
    u3 = Range(g3, e3).Address(0, 0)
    Set f3 = e3.Offset(1)
    Set e3 = e3.End(4)
    f3.Formula = "=subtotal(9, " & u3 & ")"
    ' f3.Offset(, 1) = "sum " & u3  'can omit this line if you want
    Loop Until e3.Row = Rows.Count

End Sub


