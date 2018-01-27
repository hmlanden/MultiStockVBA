Sub FindTotalTickerVolume()
' -------------------------------
' Set up all the variables I'm going to be using.
' -------------------------------

    ' Variable to hold a ticker's volume
    Dim CurrentTickerVolume As Double
    
    ' Variable to hold a ticker's name
    Dim NewTickerName As String
    
    ' Variable to hold the current "Section's" Ticker
    Dim CurrentTickerName As String
    
    ' Create the last row and last column variables
    Dim LastRow As Double
    Dim LastColumn As Double

    ' Variable to hold the current printing row
    Dim PrintingRow As Double

    
 
        
     For Each WorkingSheet In Worksheets
     
        ' -------------------------------
    ' Set up the values that the variables will have for the whole worksheet.
    ' Set printing row to start at 2 cuz we're going to have headers and set the initial value for the current tickername
    ' -------------------------------
    
        ' Set last row and last column
         LastColumn = WorkingSheet.Cells(1, Columns.Count).End(xlToLeft).Column
         LastRow = WorkingSheet.Cells(Rows.Count, 1).End(xlUp).Row

        ' set printing row
        PrintingRow = 2
        
        ' set initial "current" ticker name so everything will have something to compare to
        CurrentTickerName = WorkingSheet.Cells(2, 1)
        
        ' print the header names
        WorkingSheet.Cells(1, 9).Value = "Ticker"
        WorkingSheet.Cells(1, 10).Value = "Volume"
    
    ' -------------------------------
    ' Walk through each row that has content plus one extra so it doesn't skip the last ticker's values.
    ' Step 1: Set that row's ticker to the new name
    ' Step 2: Compare that row's ticker to the current section's ticker name.
    '               If they're the same, add the volume to the Current Ticker's volume and move on to the next row.
    '               If they're not the same, go to Step 3:
    ' Step 3: Print the Current Ticker's name (in column 9) and volume (in column 12).
    ' Step 4: Set the current ticker's name to the new name and set the volume equal to it.
    ' Step 5: Increment the printing row.
    ' -------------------------------
        For WorkingRow = 2 To LastRow + 1
            
            ' Set current working ticket name
            NewTickerName = WorkingSheet.Cells(WorkingRow, 1)
            
            ' if newTicker == Current Ticker
            If NewTickerName = CurrentTickerName Then
                CurrentTickerVolume = CurrentTickerVolume + WorkingSheet.Cells(WorkingRow, 7).Value

             Else
                ' Print the current values in the right place
                WorkingSheet.Cells(PrintingRow, 9) = CurrentTickerName
                WorkingSheet.Cells(PrintingRow, 10) = CurrentTickerVolume

                ' (re)set the variables to be equal to the new ticker's volume and name
                CurrentTickerName = WorkingSheet.Cells(WorkingRow, 1)
                CurrentTickerVolume = WorkingSheet.Cells(WorkingRow, 7)

                'increment printing row
                PrintingRow = PrintingRow + 1
            End If
        Next WorkingRow
        
    Next WorkingSheet
    
    
End Sub