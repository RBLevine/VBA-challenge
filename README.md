# VBA-challenge

This repository contains:

1. stockMarketLoop (VBScript Script File):
		This code, when implemented in a Module in VBA, will iterate through all sheet in the open
	workbook, and analyze the data. Note: this code will only work if the data is stored in the second
	row and below in columns A-G, with the first row being the Headers. In order from right to left, 
	the columns should be Ticker, Date, Open, High, Low, Close, and Volume. The Date column should be
	in yyyymmdd format. The data should be organized so that all tickers are grouped together in 
	ascending date order (January first down to the end of the year). The code will iterate through
	this data and print out a summary or the Yearly Change, Percent Change, and Total Stock Volume for 
	each ticker starting on the second row of columns I-L. It has conditional formatting for the Yearly
	Change columns, where positive change is highlighted in Green, and negative change is highlighted
	in red. Additionally, this code will provide a review of the data in columns I-L and display the
	Greatest Percent Increase, Greatest Percent Decrease, and Greatest Total Volume as well as the 
	corresponding Ticker for each of those values.
2. testing_database (Microsoft Excel Macro-Enabled Worksheet):
		This is the data provided to us for testing. I added a button labelled "Analyze" to quickly
	implement the code provided in the stockMarketLoop.vbs file. I also added a button labelled "Clear"
	that iterates through each worksheet in the workbook and deletes columns I through P so that the
	stockMarketLoop can be ran again.
3. clearData (VBS Script File):
		This code, when implemented in a Module in VBA, will iterate through all shees in the open
	workbook and delete rows I thorugh P. When ran through use of a Button, this speeds up the reset
	time on a workbook so that code can be ran again and more easily identify if it's running correctly.
4. Multiple_year_stock_data (Microsoft Excel Macro-Enabled Worksheet):
		This is the large data set that we need to run the code on. I have added the stockMarketLoop
	code to Module 1 and the clearData code to Module 2.
5. 2014-results (PNG File):
		This is a screen shot of the results my code prints to the excel sheet for the 2014 tab.
6. 2015-results (PNG File):
		This is a screen shot of the results my code prints to the excel sheet for the 2015 tab.
7. 2016-results (PNG File):
		This is a screen shot of the results my code prints to the excel sheet for the 2016 tab.
