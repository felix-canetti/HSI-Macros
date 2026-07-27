Sub Import_HSI_SubPages

'PURPOSE: To create the import files that will be used in the creation of the page36 subpages, ranging from pag36a all the way to page36L

'PROCEDURE:
 1. Find most recently imported case number from the Import Record and delete everything from it down.
'       > TIP: Select cell A3 and press CTRL + SHIFT + END to select the entire grid, then Filter out blanks on column A.
'       >      Highlight the most recently imported record yellow so it's easy to find again, then clear the filter. Go back to the yellow field
'       >      and press CTRL + SHIFT + END to select everything below it. E.g. if it is row 15, select cell A15 and press CTRL + SHIFT + END, then delete all selected rows.
'   2. Run this macro - it will create ten new Excels
'   3. Export the 'PAGE36 Serial' Smartfolder (Case Files > Reports > IT Folders) (No Hyperlinks)
'   4. In the PageSerial column, use Index/Match formula to find PageSerial from the exported PAGE36 Serial excel. Copy down and paste values in the column.
'   5. Save the workbook as the table name (PAGE36A/B/C/D/E/F/M/H/I/JK/L.csv)
'   6. Repeat with the other workbook
'   7. When importing both PAGE36A and PAGE36B and all other subpages, use 'Add All Entries' option

'Declare Variavbles
    Dim lastRow As Long
    Dim bkImport As String
    Dim stImport As String
    Dim bk36A As String
    Dim bk36B As String
    dim bk36C As String
    dim bk36D As String
    Dim bk36E As String
    Dim bk36F As String
    Dim bk36M As String
    Dim bk36H As String
    Dim bk36I As String
    Dim bk36K As String
    Dim bk36L As String

    Dim st36A As String
    Dim st36B As String
    dim st36C As String
    dim st36D As String
    Dim st36E As String
    Dim st36F As String
    Dim st36M As String
    Dim st36H As String
    Dim st36I As String
    Dim st36K As String
    Dim st36L As String

    Dim i As Iteger

    bkImport = ActiveWorkbook.Name
    stImport = ActiveWorkbook.ActiveSheet.Name

'Deleting the CountIF Column and Removing Conditional Formatting
    If Len(Range("AA1")) > 0 Then
        Range("AA:AA).Delete
    End If
    Cells.FormatConditions.Delete
End Sub 