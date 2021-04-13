 Sub initalize()
    Sheets("Top Brewers").Select

    ActiveWindow.ActivateNext
    Sheets("Segments").Select
    Range("A1").Select
    Sheets("Packages").Select
    Range("A1").Select
    Sheets("Brands").Select
    Range("A1").Select
    Sheets("Families").Select
    Range("A1").Select
    Sheets("Brewers").Select
    Range("A1").Select

    ActiveWindow.ActivateNext
    
End Sub
Sub data_copy()
' Deletes all but the first row in your report workbook, then copies the values from the data workbook.
' First row is preserved for the formatting mask.
    
' Deleting the 2nd row and down in the report workbook. Preserves the first row for formatting.
    Range("A7").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.EntireRow.Delete
    
' Switches to the other active workbook (your data file) and copy/pastes data as values into B6
    ActiveWindow.ActivateNext
    Range("A7").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Range("A1").Select
    ActiveWindow.ActivateNext
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A6").Select
    
End Sub
Sub data_copy_sum()
' Copies the total row only.
' Only needed on the first tab, as all others are referencing that.

' Clears data from C4:V4
    Range("C4:V4").Select
    Selection.ClearContents
    
' Switches to the other active workbook (your data file) and copy/pastes sum data as values into C4:V4
    ActiveWindow.ActivateNext
    Range("B6:U6").Select
    Selection.Copy
    ActiveWindow.ActivateNext
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A6").Select
    
End Sub
Sub data_copy_segments()
' Less crazy than the normal data copy, since it's a small tab.

' Selects the data in the segments tab and clears it, preserving formatting.
    Range("B6:T19").Select
    Selection.ClearContents
    Range("B6").Select
    
' Switches to the other workbook and copy/pastes data as values
    ActiveWindow.ActivateNext
    Range("A7:S20").Select
    Selection.Copy
    ActiveWindow.ActivateNext
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A6").Select
    
    Range("A6:T35").Select
    ActiveWorkbook.Worksheets("Product Segments").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Product Segments").Sort.SortFields.Add2 Key:=Range _
        ("L6:L19"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Product Segments").Sort
        .SetRange Range("A6:T19")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A6").Select
    
End Sub
Sub fix_formatting()
' Reapplies formatting for the new data.

' Selects the first row of data, then copies the formatting to the end of the file.
    Range("A6:V6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
End Sub
Sub number_rows()
' Adds rank numbers to every row.

    ActiveCell.FormulaR1C1 = "1"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A6:A7").Select
    Range("A7").Activate
    Selection.AutoFill Destination:=Range("A6:A" & Range("E" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).Select
    Range("A6").Select
    
End Sub
Sub next_tab()
' Switches both active workbooks to the next tab.

    ActiveSheet.Next.Select
    ActiveWindow.ActivateNext
    ActiveSheet.Next.Select
    ActiveWindow.ActivateNext
    
End Sub
Sub home_cell()

' Resets all tabs to home cell.
' Uncomment the extra select statement when using workbooks with kombucha pivots (RMA/Grocery/HEB)

    Sheets("Segment BDC").Select
    Range("A2").Select
    ActiveSheet.Previous.Select
    Range("A7").Select
    ActiveSheet.Previous.Select
    Range("A7").Select
    ActiveSheet.Previous.Select
    Range("A7").Select
    ActiveSheet.Previous.Select
    Range("A7").Select
    ActiveSheet.Previous.Select
    Range("A7").Select
    ActiveSheet.Previous.Select
    Range("A7").Select
    ActiveSheet.Previous.Select
    Range("A7").Select
    ActiveSheet.Previous.Select
    Range("A7").Select
    ActiveSheet.Previous.Select
    Range("A6").Select
    ActiveSheet.Previous.Select
    Range("A6").Select
    ActiveSheet.Previous.Select
    Range("A6").Select
    ActiveSheet.Previous.Select
    Range("A6").Select
    ActiveSheet.Previous.Select
    Range("A6").Select
    ActiveWorkbook.RefreshAll
    
    ActiveWindow.ActivateNext
    Sheets("Segments").Select
    Range("A1").Select
    ActiveSheet.Previous.Select
    Range("A1").Select
    ActiveSheet.Previous.Select
    Range("A1").Select
    ActiveSheet.Previous.Select
    Range("A1").Select
    ActiveSheet.Previous.Select
    Range("A1").Select
    ActiveWindow.ActivateNext
    
End Sub
Sub get_pivot_data()

' switches to last tab and deletes everything
    Sheets("Segment BDC").Select
    Range("A2:J5000").Select
    Selection.EntireRow.Delete
    Range("A2").Select

' Switches to packages tab and copies in package names
    Sheets("Top Packages").Select
    Range("B6:B5000").Select
    Selection.Copy
    Range("A6").Select
    Sheets("Segment BDC").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select

' Switches to packages tab and copies in package sales data
    Sheets("Top Packages").Select
    Range("M6:O5000").Select
    Selection.Copy
    Range("A6").Select
    Sheets("Segment BDC").Select
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    
' Switches to packages tab and copies in package share data
    Sheets("Top Packages").Select
    Range("S6:S5000").Select
    Selection.Copy
    Range("A6").Select
    Sheets("Segment BDC").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select

End Sub
Sub get_pivot_info()

' Add product info and fill Down
    Sheets("Segment BDC").Select
    Range("F2").Select
    Range("F2").Formula2 = "=XLOOKUP(A2,'https://browndist-my.sharepoint.com/personal/andrews_browndistributing_com/Documents/NAA/IRI Reports/[AB SubSegment Values.xlsx]SubSegment Values'!$A:$A,'https://browndist-my.sharepoint.com/personal/andrews_browndistributing_com/Documents/NAA/IRI Reports/[AB SubSegment Values.xlsx]SubSegment Values'!$B:$G)"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F5000"), Type:=xlFillDefault
    Range("F2").Select
    'Application.Wait Now + #12:00:05 AM#
    
End Sub
Sub get_pivot_finalize()
    Sheets("Segment BDC").Select
    Range("A2:K5000").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    
' Format
    Range("A:K").Select
    Selection.NumberFormat = "0.00"
    Range("A2").Select
    ActiveWorkbook.RefreshAll
    
End Sub
Sub A_PREPARE_ALL()
' Runs subroutines for each tab.

' Checks for the required tabs to be open in a 2nd workbook, fails if they are not found.
' Also resets both workbooks to the first tab.
    Call initalize

' Brewers
    Call data_copy_sum
    Call data_copy
    Call fix_formatting
    Call number_rows
    Call next_tab
    
' Families
    Call data_copy
    Call fix_formatting
    Call number_rows
    Call next_tab

' Brands
    Call data_copy
    Call fix_formatting
    Call number_rows
    Call next_tab

' Packages
    Call data_copy
    Call fix_formatting
    Call number_rows
    Call next_tab
    
' Segments
    Call data_copy_segments
        
' Pivot Data Prep
    Call get_pivot_data
    Call get_pivot_info

    
End Sub
Sub B_FINALIZE_ALL()

    'Reset views and finalize
    Call get_pivot_finalize
    Call home_cell

End Sub

