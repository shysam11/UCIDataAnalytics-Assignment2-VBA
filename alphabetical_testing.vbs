' Steps:
' ----------------------------------------------------------------------------

'Create separate worksheets 2016,2015,2014 and move before the originally provided worksheets A-P
' 1. Loop through every worksheet .
'   a. If year(ColumnB)=2016 ,  Copy the worksheet contents and paste it into the Comb_sheet tab 2016
'   b. If year(ColumnB)=2015 ,  Copy the worksheet contents and paste it into the Comb_sheet tab 2015
'   c. If year(ColumnB)=2014 ,  Copy the worksheet contents and paste it into the Comb_sheet tab 2014

Sub yr_ticker()
'Dim yr As Integer
'yr = Year(Range("B:B"))
    
    ' Add sheets named 2016, 2015, 2014
    Sheets.Add.Name = "2016"
    Sheets.Add.Name = "2015"
    Sheets.Add.Name = "2014"
    'move created sheets to be before original test data sheets A-P
    Sheets("2014").Move Before:=Sheets(1)
    Sheets("2015").Move Before:=Sheets("2014")
    Sheets("2016").Move Before:=Sheets("2015")
    ' Specify the location of the combined sheet
    Set comb_sheet2016 = Worksheets("2016")
    Set comb_sheet2015 = Worksheets("2015")
    Set comb_sheet2014 = Worksheets("2014")

    ' Loop through all sheets
    For Each ws In Worksheets

        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        lastRow = comb_sheet2016.Cells(Rows.Count, "A").End(xlUp).Row + 1
        lastRow = comb_sheet2015.Cells(Rows.Count, "A").End(xlUp).Row + 1
        lastRow = comb_sheet2014.Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

        ' Copy the contents of each ticker sheet into the yearly combined sheets
        comb_sheet2016.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
        comb_sheet2015.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
        comb_sheet2014.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
    Next ws

    ' Copy the headers from sheet 1
    comb_sheet2016.Range("A1:G1").Value = A.Range("A1:G1").Value
    comb_sheet2015.Range("A1:G1").Value = Sheets(1).Range("A1:G1").Value
    comb_sheet2014.Range("A1:G1").Value = Sheets(1).Range("A1:G1").Value
    
    ' Autofit to display data
    comb_sheet2016.Columns("A:G").AutoFit
    comb_sheet2015.Columns("A:G").AutoFit
    comb_sheet2014.Columns("A:G").AutoFit
End Sub

