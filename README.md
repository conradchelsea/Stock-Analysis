#VBA Challenge

## Using Visual Basic for Applications to analyze and automate an excel stock market dataset for green energy companies. 

### Optimized by refacting code in VBA to uncover the volume and return percentage for solar company DAQO against other green energy companies from 2017-2018. 

## The conclusion proved refacting the VBA script icreased the speed of the outcome and that DAQO is had a poor performance in the stock market.  

### The 2017 data script ran faster than the previous worksheet by 0.8 of a second. 

### The 2018 data script also ran almost a second faster than the previous version. 

###
			 
			    'Number of rows to loop over
			    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
			    '1a) Create a ticker Index
			    tickerIndex = 0
			    '1b) Create three output arrays
			    Dim tickerVolumes(11) As Long
			    Dim tickerStartingPrices(11) As Single
			    Dim tickerEndingPrices(11) As Single
			    currentRowTicker = ""
			    currentRowStartPrice = 0
			    currentRowEndPrice = 0
			    currentRowVolume = 0	
				
				''2b) Loop over all the rows in the spreadsheet.
         For RowIndex = 2 To RowCount
         'Get current row Ticker
         currentRowTicker = Cells(RowIndex, 1).Value
         currentRowEndPrice = Cells(RowIndex, 6).Value
         currentRowVolume = Cells(RowIndex, 8).Value
                
          'First row set start price
           If Cells(RowIndex - 1, 1).Value <> currentRowTicker Then
           tickerStartingPrices(tickerIndex) = currentRowEndPrice
           End If
                
          '3a) Increase volume for current ticker
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + currentRowVolume
                
          '3c) check if the current row is the last row with the selected ticker
           If Cells(RowIndex + 1, 1).Value <> currentRowTicker Then
          'last row for ticker so set end price
           tickerEndingPrices(tickerIndex) = currentRowEndPrice
          'last row of ticker so increase tickerindex
           tickerIndex = tickerIndex + 1
                
           End If
				
				Next Row Index
        

### Being able to run the loop once as opposed to numerous times is more efficient, especially as the data sets get larger. The challenge with running the the loop by start to finish, beginning to end is that I am assuming the data is in order. If for some reason the data wasn't in order (the system I downloaded from changed, the person who usually puts the files together left, etc.) my data would be incorrect. 

### The advantages of refacting are it cuts down compute time (especially with large files purchased from Google or Amazon), saves on labor, money and resources. The disadvantages are the time it takes to edit the data. What is the opportinuty cost to have an engineer or analyst working on that as opposed to something else, potentially more valuable? 
 
## Both 2017 & 2018 ran faster after refactoring the VBA script. 
