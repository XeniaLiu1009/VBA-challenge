

Sub stock_function()
    
    'create four variables to hold ticker symbol, yearly change, percentage change and totale stock volume.
    Dim ticker_symbol As String
    Dim yearly_change, percent_change, total_stock_volume,yearly_start,yearly_end,vol_start As Double
    'Dim percent_change As Double
    'Dim total_stock_volume As Double
    'create other two variables to help calculate the percent change
    'Dim yearly_start As Double
    'Dim yearly_end As Double
    'create a counter 
    Dim counter as integer 
    'Dim vol_start as Double
    'create the row dimension 
    Dim LR As Long 
    
    'create four variables' title
    Range("I1").Value = "Ticker Symbol"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

 	'set a start counter
 	counter = 2
 	
 	'count the number of row
 	LR = Cells(Rows.count, 1).End(xLUp).Row
 	MsgBox LR

    'loop through all the stock for one year
    For i = 2 To LR 'row(data) = 753001 

    	'filter by the start of the year
        If Cells(i, 2).Value = 20180102 Then

            'the ticker symbol
            ticker_symbol = cells(i, 1) 'assign the ticker string to the real data (ticker)
            Cells(counter, 9) = ticker_symbol 'put the saved data in the ticker sybol in I column

            'opening price at the beginning of a given year
             yearly_start = Cells(i, 3) 'assign the opening price to the yearly start

             vol_start = Cells(i,7)
             total_stock_volume = vol_start
        Else
        total_stock_volume = total_stock_volume  + cells(i,7)

        End if 


        'filter by the end of the year
        If Cells(i, 2).Value = 20181231 Then
           
           'closing price at the end of that year
             yearly_end = Cells(i, 6) 'assign the ending price to the yearly end

            'yearly change from opening price at the beginning of a given year to the closing price at the end of that year
       		yearly_change = yearly_end - yearly_start
       		Cells(counter, 10) = yearly_change 'assign the J column the changing price

       		'the percent change from opening price at the beginning of a given year to the closing price at the end of that year
        	percentage_change = yearly_change / yearly_start
        	Cells(counter, 11) = FormatPercent(percentage_change) 'assign the K column the percentage change 

        	
        	'the total stock volume of the stock
        	Cells(counter, 12) = total_stock_volume

        	'add 1 to the counter
        	counter = counter + 1

        End If
  
    Next i
    
    MsgBox ("Competed!")
    
 For j = 2 To counter-1

 	If Cells(j,10) > 0 Then

 		Cells(j,10).Interior.ColorIndex = 4
    Else
    	Cells(j,10).Interior.ColorIndex = 3

	End If

 Next j

End Sub



'Bonus




Sub Greatest_Function():

    Dim maximum, minimum, max_total As Double
    
    'create five variables' title
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volumn"
    
    Set my_range = Worksheets("2018").Range("K1:K3001")
    Set my_range2 = Worksheets("2018").Range("L1:L3001")
    
    maximum = Application.WorksheetFunction.Max(my_range)
    Range("Q2").Value = FormatPercent(maximum)

    minimum = Application.WorksheetFunction.Min(my_range)
    Range("Q3").Value = FormatPercent(minimum)
    
    max_total = Application.WorksheetFunction.Max(my_range2)
    Range("Q4").Value = max_total

    for i = 1 To 3001

    	if Cells(i,11) = maximum then
    	Range("P2").value = Cells(i,9) 
        end if 

    	if Cells(i,11) = minimum then
    	Range("P3").value = Cells(i,9) 
        end if 
        
    	if Cells(i,12) = max_total then
    	Range("P4").value = Cells(i,9) 
   		end if

    Next i 
  
End Sub





Sub stock_three_function():

    Set my_range1 = ActiveWorkbook.Sheets("2018")
    Set my_range2 = ActiveWorkbook.Sheets("2019")
    Set my_range3 = ActiveWorkbook.Sheets("2020")
    
    
    MsgBox (my_range1.Cells(2, 3))
    MsgBox (my_range2.Cells(2, 3))
    MsgBox (my_range3.Cells(2, 3))

    Dim ticker_symbol_2018, ticker_symbol_2019, ticker_symbol_2020 As String
    Dim yearly_change_2018, percent_change_2018, total_stock_volume_2018, yearly_start_2018, yearly_end_2018, vol_start_2018 As Double
    Dim yearly_change_2019, percent_change_2019, total_stock_volume_2019, yearly_start_2019, yearly_end_2019, vol_start_2019 As Double
    Dim yearly_change_2020, percent_change_2020, total_stock_volume_2020, yearly_start_2020, yearly_end_2020, vol_start_2020 As Double

    Dim counter_2018, counter_2019, counter_2020 as integer 
    Dim LR_2018, LR_2019, LR_2020 As Long 
    
    'create four variables' title
    my_range1.Range("I1").Value = "Ticker Symbol"
    my_range1.Range("J1").Value = "Yearly Change"
    my_range1.Range("K1").Value = "Percent Change"
    my_range1.Range("L1").Value = "Total Stock Volume"

    my_range2.Range("I1").Value = "Ticker Symbol"
    my_range2.Range("J1").Value = "Yearly Change"
    my_range2.Range("K1").Value = "Percent Change"
    my_range2.Range("L1").Value = "Total Stock Volume"


    my_range3.Range("I1").Value = "Ticker Symbol"
    my_range3.Range("J1").Value = "Yearly Change"
    my_range3.Range("K1").Value = "Percent Change"
    my_range3.Range("L1").Value = "Total Stock Volume"


 	'set a start counter
 	counter_2018 = 2
 	counter_2019 = 2
 	counter_2020 = 2
 	
 	'count the number of row
 	LR_2018 = my_range1.Cells(Rows.count, 1).End(xLUp).Row
 	LR_2019 = my_range2.Cells(Rows.count, 1).End(xLUp).Row
 	LR_2020 = my_range3.Cells(Rows.count, 1).End(xLUp).Row
 	
 	MsgBox LR_2018
 	MsgBox LR_2019
 	MsgBox LR_2020

    'loop through all the stock for one year
    For i = 2 To LR_2018 'row(data) = 753001 

    	'filter by the start of the year
        If my_range1.Cells(i, 2).Value = 20180102 Then

            'the ticker symbol
            ticker_symbol_2018 = my_range1.cells(i, 1) 'assign the ticker string to the real data (ticker)
            my_range1.Cells(counter_2018, 9) = ticker_symbol_2018 'put the saved data in the ticker sybol in I column

            'opening price at the beginning of a given year
             yearly_start_2018 = my_range1.Cells(i, 3) 'assign the opening price to the yearly start

             vol_start_2018 = my_range1.Cells(i,7)
             total_stock_volume_2018 = vol_start_2018
        Else
        total_stock_volume_2018 = total_stock_volume_2018  + my_range1.cells(i,7)

        End if 


        'filter by the end of the year
        If my_range1.Cells(i, 2).Value = 20181231 Then
           
           'closing price at the end of that year
             yearly_end_2018 = my_range1.Cells(i, 6) 'assign the ending price to the yearly end

            'yearly change from opening price at the beginning of a given year to the closing price at the end of that year
       		yearly_change_2018 = yearly_end_2018 - yearly_start_2018
       		my_range1.Cells(counter_2018, 10) = yearly_change_2018 'assign the J column the changing price

       		'the percent change from opening price at the beginning of a given year to the closing price at the end of that year
        	percentage_change_2018 = yearly_change_2018 / yearly_start_2018
        	my_range1.Cells(counter_2018, 11) = FormatPercent(percentage_change_2018) 'assign the K column the percentage change 

        	
        	'the total stock volume of the stock
        	my_range1.Cells(counter_2018, 12) = total_stock_volume_2018

        	'add 1 to the counter
        	counter_2018 = counter_2018 + 1

        End If
  
    Next i
    
    MsgBox ("Competed!")
    
 For j = 2 To counter_2018-1

 	If my_range1.Cells(j,10) > 0 Then

 		my_range1.Cells(j,10).Interior.ColorIndex = 4
    Else
    	my_range1.Cells(j,10).Interior.ColorIndex = 3

	End If

 Next j

'loop through all the stock for one year
    For i = 2 To LR_2019 'row(data) = 753001 

    	'filter by the start of the year
        If my_range2.Cells(i, 2).Value = 20190102 Then

            'the ticker symbol
            ticker_symbol_2019 = my_range2.cells(i, 1) 'assign the ticker string to the real data (ticker)
            my_range2.Cells(counter_2019, 9) = ticker_symbol_2019 'put the saved data in the ticker sybol in I column

            'opening price at the beginning of a given year
             yearly_start_2019 = my_range2.Cells(i, 3) 'assign the opening price to the yearly start

             vol_start_2019 = my_range2.Cells(i,7)
             total_stock_volume_2019 = vol_start_2019
        Else
        total_stock_volume_2019 = total_stock_volume_2019  + my_range2.cells(i,7)

        End if 


        'filter by the end of the year
        If my_range2.Cells(i, 2).Value = 20191231 Then
           
           'closing price at the end of that year
             yearly_end_2019 = my_range2.Cells(i, 6) 'assign the ending price to the yearly end

            'yearly change from opening price at the beginning of a given year to the closing price at the end of that year
       		yearly_change_2019 = yearly_end_2019 - yearly_start_2019
       		my_range2.Cells(counter_2019, 10) = yearly_change_2019 'assign the J column the changing price

       		'the percent change from opening price at the beginning of a given year to the closing price at the end of that year
        	percentage_change_2019 = yearly_change_2019 / yearly_start_2019
        	my_range2.Cells(counter_2019, 11) = FormatPercent(percentage_change_2019) 'assign the K column the percentage change 

        	
        	'the total stock volume of the stock
        	my_range2.Cells(counter_2019, 12) = total_stock_volume_2019

        	'add 1 to the counter
        	counter_2019 = counter_2019 + 1

        End If
  
    Next i
    
    MsgBox ("Competed!")
    
 For j = 2 To counter_2019-1

 	If my_range2.Cells(j,10) > 0 Then

 		my_range2.Cells(j,10).Interior.ColorIndex = 4
    Else
    	my_range2.Cells(j,10).Interior.ColorIndex = 3

	End If

 Next j


 'loop through all the stock for one year
    For i = 2 To LR_2020 'row(data) = 753001 

    	'filter by the start of the year
        If my_range3.Cells(i, 2).Value = 20200102 Then

            'the ticker symbol
            ticker_symbol_2020 = my_range3.cells(i, 1) 'assign the ticker string to the real data (ticker)
            my_range3.Cells(counter_2020, 9) = ticker_symbol_2020 'put the saved data in the ticker sybol in I column

            'opening price at the beginning of a given year
             yearly_start_2020 = my_range3.Cells(i, 3) 'assign the opening price to the yearly start

             vol_start_2020 = my_range3.Cells(i,7)
             total_stock_volume_2020 = vol_start_2020
        Else
        total_stock_volume_2020 = total_stock_volume_2020  + my_range3.cells(i,7)

        End if 


        'filter by the end of the year
        If my_range3.Cells(i, 2).Value = 20201231 Then
           
           'closing price at the end of that year
             yearly_end_2020 = my_range3.Cells(i, 6) 'assign the ending price to the yearly end

            'yearly change from opening price at the beginning of a given year to the closing price at the end of that year
       		yearly_change_2020 = yearly_end_2020 - yearly_start_2020
       		my_range3.Cells(counter_2020, 10) = yearly_change_2020 'assign the J column the changing price

       		'the percent change from opening price at the beginning of a given year to the closing price at the end of that year
        	percentage_change_2020 = yearly_change_2020 / yearly_start_2020
        	my_range3.Cells(counter_2020, 11) = FormatPercent(percentage_change_2020) 'assign the K column the percentage change 

        	
        	'the total stock volume of the stock
        	my_range3.Cells(counter_2020, 12) = total_stock_volume_2020

        	'add 1 to the counter
        	counter_2020 = counter_2020 + 1

        End If
  
    Next i
    
    MsgBox ("Competed!")
    
 For j = 2 To counter_2020-1

 	If my_range3.Cells(j,10) > 0 Then

 		my_range3.Cells(j,10).Interior.ColorIndex = 4
    Else
    	my_range3.Cells(j,10).Interior.ColorIndex = 3

	End If

 Next j

    
End Sub










Sub Greatest_Function_three():


	Set my_range1 = ActiveWorkbook.Sheets("2018")
    Set my_range2 = ActiveWorkbook.Sheets("2019")
    Set my_range3 = ActiveWorkbook.Sheets("2020")

    Dim maximum_2018, maximum_2019, maximum_2020, minimum_2018, minimum_2019, minimum_2020, max_total_2018, max_total_2019, max_total_2020 As Double
    
    'create five variables' title
    my_range1.Range("P1").Value = "Ticker"
    my_range1.Range("Q1").Value = "Value"
    my_range1.Range("O2").Value = "Greatest % Increase"
    my_range1.Range("O3").Value = "Greatest % Decrease"
    my_range1.Range("O4").Value = "Greatest Total Volumn"

    my_range2.Range("P1").Value = "Ticker"
    my_range2.Range("Q1").Value = "Value"
    my_range2.Range("O2").Value = "Greatest % Increase"
    my_range2.Range("O3").Value = "Greatest % Decrease"
    my_range2.Range("O4").Value = "Greatest Total Volumn"

    my_range3.Range("P1").Value = "Ticker"
    my_range3.Range("Q1").Value = "Value"
    my_range3.Range("O2").Value = "Greatest % Increase"
    my_range3.Range("O3").Value = "Greatest % Decrease"
    my_range3.Range("O4").Value = "Greatest Total Volumn"

    
    Set pct_col_2018 = Worksheets("2018").Range("K1:K3001")
    Set vol_col_2018 = Worksheets("2018").Range("L1:L3001")

    Set pct_col_2019 = Worksheets("2019").Range("K1:K3001")
    Set vol_col_2019 = Worksheets("2019").Range("L1:L3001")


    Set pct_col_2020 = Worksheets("2020").Range("K1:K3001")
    Set vol_col_2020 = Worksheets("2020").Range("L1:L3001")

    
    maximum_2018 = Application.WorksheetFunction.Max(pct_col_2018)
    my_range1.Range("Q2").Value = FormatPercent(maximum_2018)
    maximum_2019 = Application.WorksheetFunction.Max(pct_col_2019)
    my_range2.Range("Q2").Value = FormatPercent(maximum_2019)
    maximum_2020 = Application.WorksheetFunction.Max(pct_col_2020)
    my_range3.Range("Q2").Value = FormatPercent(maximum_2020)


    minimum_2018 = Application.WorksheetFunction.Min(pct_col_2018)
    my_range1.Range("Q3").Value = FormatPercent(minimum_2018)
    minimum_2019 = Application.WorksheetFunction.Min(pct_col_2019)
    my_range2.Range("Q3").Value = FormatPercent(minimum_2019)
    minimum_2020 = Application.WorksheetFunction.Min(pct_col_2020)
    my_range3.Range("Q3").Value = FormatPercent(minimum_2020)
    
    max_total_2018 = Application.WorksheetFunction.Max(vol_col_2018)
    my_range1.Range("Q4").Value = max_total_2018
    max_total_2019 = Application.WorksheetFunction.Max(vol_col_2019)
    my_range2.Range("Q4").Value = max_total_2019
    max_total_2020 = Application.WorksheetFunction.Max(vol_col_2020)
    my_range3.Range("Q4").Value = max_total_2020

    for i = 1 To 3001

    	if my_range1.Cells(i,11) = maximum_2018 then
    	my_range1.Range("P2").value = my_range1.Cells(i,9) 
        end if 

    	if my_range1.Cells(i,11) = minimum_2018 then
    	my_range1.Range("P3").value = my_range1.Cells(i,9) 
        end if 
        
    	if my_range1.Cells(i,12) = max_total_2018 then
    	my_range1.Range("P4").value = my_range1.Cells(i,9) 
   		end if


   		if my_range2.Cells(i,11) = maximum_2019 then
    	my_range2.Range("P2").value = my_range2.Cells(i,9) 
        end if 

    	if my_range2.Cells(i,11) = minimum_2019 then
    	my_range2.Range("P3").value = my_range2.Cells(i,9) 
        end if 
        
    	if my_range2.Cells(i,12) = max_total_2019 then
    	my_range2.Range("P4").value = my_range2.Cells(i,9) 
   		end if


   		if my_range3.Cells(i,11) = maximum_2020 then
    	my_range3.Range("P2").value = my_range3.Cells(i,9) 
        end if 

    	if my_range3.Cells(i,11) = minimum_2020 then
    	my_range3.Range("P3").value = my_range3.Cells(i,9) 
        end if 
        
    	if my_range3.Cells(i,12) = max_total_2020 then
    	my_range3.Range("P4").value = my_range3.Cells(i,9) 
   		end if

    Next i 
  
End Sub

