Attribute VB_Name = "Module1"
Sub alphabet_testing()

' Define and set up variables
    ' Create worksheet
    Dim ws As Worksheet

 
' Start Loop to loop through all worksheets (A-P)
For Each ws In Worksheets

  
  ' Activate worksheet
  ws.Activate

    ' Set an initial variable for holding the alphabet ticker
        Dim Ticker As String
    ' Set an initial variable for holding the Last Row of Worksheet
        Dim LastRow As Long
        
    ' Set an initial variable for holding the total volume per alphabet ticker
        Dim Total_Volume As Double
        Total_Volume = 0
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    'set headers
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Total Stock Volume"
  
      'Set Last Row of Worksheet
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    ' Loop through all alphabets
    For i = 2 To LastRow

        ' Check if we are still within the same alphabet ticker, if  not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

         ' Set the alphabet ticker
         Ticker = ws.Cells(i, 1).Value

        ' Add to the Total Volume
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value

         ' Print the Alphabet Ticker in the Summary Table
        Range("K" & Summary_Table_Row).Value = Ticker

        ' Print the Total Volume to the Summary Table
        Range("L" & Summary_Table_Row).Value = Total_Volume

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
         ' Reset the Total Volume
         Total_Volume = 0

            ' If the cell immediately following a row is the same alphabet ticker...
            Else

          ' Add to the Total Volume
             Total_Volume = Total_Volume + ws.Cells(i, 7).Value

        End If

    ' Loop to next alphabet row
     Next i
     
    ' Loop to next worksheet
    Next ws
  
End Sub


