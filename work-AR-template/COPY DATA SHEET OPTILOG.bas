Attribute VB_Name = "Module11"
Sub Copy_Template()
Attribute Copy_Template.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Copy_Template Macro
'
' Keyboard Shortcut: Ctrl+q
'
    Dim Actsheet As String
    Application.ScreenUpdating = False
    On Error Resume Next
    ActiveWorkbook.Sheets("Template SUMMARY").Visible = True
    ActiveWorkbook.Sheets("Template SUMMARY").Copy _
    After:=ActiveWorkbook.Sheets("Template SUMMARY")
    ActNm = ActiveSheet.Name
    ActiveSheet.Name = "SUMMARY"
    Sheets(ActiveSheet.Name).Visible = True
    ActiveWorkbook.Sheets("Template SUMMARY").Visible = False
    Application.ScreenUpdating = True

End Sub



Sub Copy_Optilog_Data_2()
Attribute Copy_Optilog_Data_2.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'
    Application.ScreenUpdating = True
    
    Sheets("optilog").Select
    Range("E2:E5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("D4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("C2:C5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("E4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("F2:F5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("F4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("R2:R5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("I4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("T2:T5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("J4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("Y2:Y5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("K4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("V2:V5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("L4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("W2:W5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("M4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("I2:I5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("N4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("Z2:Z5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("O4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("AA2:AA5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("P4").Select
    ActiveSheet.Paste
        
    Sheets("optilog").Select
    Range("AC2:AC5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("Q4").Select
    ActiveSheet.Paste
    Range("Q4:Q5004").Select
    Selection.TextToColumns Destination:=Range("Q4"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Selection.NumberFormat = "0.00"
        
    Sheets("optilog").Select
    Range("AD2:AD5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("R4").Select
    ActiveSheet.Paste

    Sheets("optilog").Select
    Range("AE2:AE5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("S4").Select
    ActiveSheet.Paste


    Sheets("optilog").Select
    Range("AF2:AF5000").Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("T4").Select
    ActiveSheet.Paste
    Range("T4:T5004").Select
    Selection.TextToColumns Destination:=Range("T4"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Selection.NumberFormat = "0.00"
    
    ' Shortage
    Range("U4").Select
    ActiveCell.FormulaR1C1 = "=(RC[-4]-RC[-1])/RC[-4]"
    Range("U4").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Selection.AutoFill Destination:=Range("U4:U5000"), Type:=xlFillDefault
    ' End Shortage
    
    ' Price/kg
    Range("V4").Select
    Selection.NumberFormat = "0.00"
    Selection.AutoFill Destination:=Range("V4:V5000"), Type:=xlFillDefault
    ' End Price/kg
    
    ' Sales
    Range("W4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-3]"
    Selection.Style = "Currency [0]"
    ' Selection.NumberFormat = "_-$* #,##0.0_-;-$* #,##0.0_-;_-$* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* ""-""??_-;_-@_-"
    Selection.AutoFill Destination:=Range("W4:W5000"), Type:=xlFillDefault
    ' End Price
    
    ' Tambahan
    Range("X4").Select
    Selection.Style = "Currency [0]"
    Selection.NumberFormat = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* ""-""??_-;_-@_-"
    Selection.AutoFill Destination:=Range("X4:X5000"), Type:=xlFillDefault
    ' End Tambahan
    
    ' TOTAL PRICE
    Range("Y4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]+RC[-2]"
    Selection.Style = "Currency [0]"
    Selection.NumberFormat = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* ""-""??_-;_-@_-"
    Selection.AutoFill Destination:=Range("Y4:Y5000"), Type:=xlFillDefault
    ' END TOTAL PRICE
    
    
    '==============================================
    ' index-match and compare summary sheet to optilog sheet

    Sheets("SUMMARY").Select
    Range("AB4").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(optilog!R2C22:R5000C22, MATCH(RC9, optilog!R2C18:R5000C18, 0),1)"
    Range("AC4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]=RC[-17]"
        
    Range("AD4").Select
    ActiveCell.FormulaR1C1 = _
        "=VALUE(INDEX(optilog!R2C29:R5000C29, MATCH(RC9, optilog!R2C18:R5000C18, 0),1))"
    Range("AE4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]=RC[-14]"
    
    'COMPARE NET UNLOAD'
    Range("AF4").Select
    ActiveCell.FormulaR1C1 = _
        "=VALUE(INDEX(optilog!R2C32:R5000C32, MATCH(SUMMARY!RC9, optilog!R2C18:R5000C18, 0),1))"
    Range("AG4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]=RC[-13]"
    'END COMPARE NET UNLOAD'
    
    ' COMPARE SALES'
    Range("AH4").Select
    ActiveCell.FormulaR1C1 = _
        "=VALUE(INDEX(optilog!R2C40:R5000C40, MATCH(SUMMARY!RC9, optilog!R2C18:R5000C18, 0),1))"
    Range("AI4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]=RC[-12]"
    'END COMPARE SALES'
    
    'COMPARE LOAD DATE'
    Range("AJ4").Select
    Selection.NumberFormat = "dd/mm/yyyy"
    ActiveCell.FormulaR1C1 = _
        "=VALUE(INDEX(optilog!R2C26:R5000C26, MATCH(SUMMARY!RC9, optilog!R2C18:R5000C18, 0),1))"
    Range("AK4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]=RC[-22]"
    'END COMPARE LOAD DATE'
    
    'COMPARE LOAD DATE'
    Range("AL4").Select
    Selection.NumberFormat = "dd/mm/yyyy"
    ActiveCell.FormulaR1C1 = _
        "=VALUE(INDEX(optilog!R2C30:R5000C30, MATCH(SUMMARY!RC9, optilog!R2C18:R5000C18, 0),1))"
    Range("AM4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]=RC[-21]"
    'END COMPARE LOAD DATE'
    

    
    
    ' Drag Autofill Rumus
    Range("AB4:AM4").Select
    Selection.AutoFill Destination:=Range("AB4:AM5000"), Type:=xlFillDefault
    
    
    ' conditional formating false comparison
    Range("AB4:AM5000").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=FALSE"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    '==============================================
    
    ' ADD BORDER TO SUMMARY TABLE
    Sheets("SUMMARY").Select
    Range("A4:Z5000").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ' END ADD BORDER TO SUMMARY TABLE
    
    ' ADD BORDER TO COMPARISON TABLE
    Sheets("SUMMARY").Select
    Range("AB4:AM5000").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ' END BORDER TO COMPARISON TABLE
    
    ' FILTER TABLE TO CLEAR YELLOW ROW
        
    Range("A4:AM5000").Select
    ActiveSheet.Range("$A$3:$AM$5000").AutoFilter Field:=10, Criteria1:=RGB(255 _
        , 255, 0), Operator:=xlFilterCellColor
    Selection.ClearContents
    Selection.Font.BOLD = True
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = 6
        .color = 65535
    End With
    ActiveSheet.Range("$A$3:$AF$5000").AutoFilter Field:=10
    
    ' END FILTER TABLE TO CLEAR YELLOW ROW
    
    
    ' FILTER TABLE TO REMOVE BLANK TRUCK PLATE
    
    Range("A3:AM5000").Select
    Selection.AutoFilter
    Range("A4:AM5000").Select
    ActiveSheet.Range("$A$3:$AM$5000").AutoFilter Field:=12, Criteria1:="="
    ActiveSheet.Range("$A$3:$AM$5000").AutoFilter Field:=17, Criteria1:="="
    ActiveSheet.Range("$A$3:$AM$5000").AutoFilter Field:=13, Operator:= _
        xlFilterNoFill
    
    
    Range("A4:AM5000").Select
    'Selection.SpecialCells(xlCellTypeVisible).Select
    
    Union(Range("A4:AM5000"), Selection.SpecialCells(xlCellTypeVisible)).Select
    
    Selection.EntireRow.Delete
    
    'With Selection
        '.Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        '.Parent.AutoFilterMode = False
    'End With
    
    ActiveSheet.Range("$A$3:$AM$5000").AutoFilter Field:=12
    ActiveSheet.Range("$A$3:$AM$5000").AutoFilter Field:=13
    ActiveSheet.Range("$A$3:$AM$5000").AutoFilter Field:=17
    
    ' END FILTER TABLE TO REMOVE BLANK TRUCK PLATE
    
    ' JUDUL SUMMARY
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "=""SUMMARY BONGKARAN ""&R[3]C[9]"
    Selection.Font.BOLD = True
    ' END JUDUL SUMMARY
    
    ' AUTOFIT KOLOM
    Columns("D:G").Select
    Selection.Columns.AutoFit
    
    Columns("I:Z").Select
    Selection.Columns.AutoFit
    
    Columns("AB:AM").Select
    Selection.Columns.AutoFit
    
    Columns("E").ColumnWidth = 16
    Columns("F").ColumnWidth = 12
    Columns("G").ColumnWidth = 12
    Columns("H").ColumnWidth = 5
    Columns("I").ColumnWidth = 16
    Columns("J").ColumnWidth = 24
    Columns("K").ColumnWidth = 24
    Columns("L").ColumnWidth = 11
    Columns("O").ColumnWidth = 11
    Columns("R").ColumnWidth = 11
    Columns("Q").ColumnWidth = 10
    Columns("T").ColumnWidth = 10
    Columns("U").ColumnWidth = 8
    Columns("V").ColumnWidth = 8
    Columns("W").ColumnWidth = 16
    Columns("X").ColumnWidth = 16
    Columns("Y").ColumnWidth = 16
    Columns("Z").ColumnWidth = 8
    Columns("AB").ColumnWidth = 9
    Columns("AC").ColumnWidth = 7
    Columns("AD").ColumnWidth = 9
    Columns("AE").ColumnWidth = 7
    Columns("AF").ColumnWidth = 9
    Columns("AG").ColumnWidth = 7
    Columns("AH").ColumnWidth = 9
    Columns("AI").ColumnWidth = 7
    Columns("AJ").ColumnWidth = 10
    Columns("AK").ColumnWidth = 7
    Columns("AL").ColumnWidth = 10
    Columns("AM").ColumnWidth = 7
    
    Range("F3").WrapText = True
    Range("G3").WrapText = True
    Range("Z3").WrapText = True
    Rows(3).Select
    Selection.Rows.AutoFit
    
    
    
    
    
    ' END AUTOFIT
    
    Application.ScreenUpdating = True
    
End Sub
