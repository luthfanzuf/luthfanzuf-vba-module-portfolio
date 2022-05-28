Attribute VB_Name = "Module3"
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function


Sub Copy_Sheet_Template_1()
Attribute Copy_Sheet_Template_1.VB_ProcData.VB_Invoke_Func = "q\n14"
    
    Dim Actsheet As String
    Dim answer1 As Integer
    Dim answer2 As Integer
    
    

    If WorksheetExists("SUMMARY") Then
        MsgBox ("SUMMARY SUDAH ADA")
        Sheets("SUMMARY").Select
        answer1 = MsgBox("Reset Semua Data di Sheet Summary?", vbQuestion + vbYesNo + vbDefaultButton2, "Reset Sheet Summary")
        
        If answer1 = vbYes Then
        
            ' Clear Filled Row
            Rows("1:5000").Select
            Rows("1:5001").Select
            Selection.AutoFilter
            Selection.ClearContents
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            Selection.Borders(xlEdgeTop).LineStyle = xlNone
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
            Selection.Borders(xlEdgeRight).LineStyle = xlNone
            Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("A4:AL6").Select
            Selection.ClearContents
            ' End Clear Filled Row
            
            ' Copy Row From Template
                ActiveWorkbook.Sheets("Template SUMMARY").Visible = True
                Sheets("Template SUMMARY").Select
                Rows("1:3").Select
                Selection.EntireRow.Hidden = False
                Rows("2:3").Select
                Selection.Copy
                Rows("2:2").Select
                Selection.EntireRow.Hidden = True
                Sheets("SUMMARY").Select
                Rows("2:2").Select
                ActiveSheet.Paste
                ActiveWorkbook.Sheets("Template SUMMARY").Visible = False
            ' End Copy Row From Template
            
            
            ' Freeze Pane Reset
            Range("J4").Select
            ActiveWindow.FreezePanes = False
            ActiveWindow.FreezePanes = True
            ' END Freeze Pane Reset
            
            
            MsgBox "Sheet Summary Cleared"
            
        Else
            MsgBox "Sheet Summary Presisted"
        End If
        
        answer2 = MsgBox("Copy Data dari Sheet Optilog?", vbQuestion + vbYesNo + vbDefaultButton2, "Copy Data Optilog")
                
        If answer2 = vbYes Then
                    
            MsgBox "Mengcopy Data Optilog"
            Copy_Optilog_Data_2
            
        Else
            MsgBox "Do Nothing"
        End If

        
    Else
        MsgBox ("SUMMARY TIDAK ADA")

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
        Range("I4").Select
        ActiveWindow.FreezePanes = True
        MsgBox ("SUMMARY SUDAH DIBUAT")
        Sheets("SUMMARY").Select
        
        answer2 = MsgBox("Copy Data dari Sheet Optilog?", vbQuestion + vbYesNo + vbDefaultButton2, "Copy Data Optilog")
        
        If answer2 = vbYes Then
            MsgBox "Mengcopy Data Optilog"
            Copy_Optilog_Data_2
            
        Else
            MsgBox "Do Nothing"
        End If
        
    End If

End Sub
