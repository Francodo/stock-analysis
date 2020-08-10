# stock-analysis
Francis Odo

Background 
This project provides an Excel dataset on stock trade performance of some companies for the year 2017 and 2018 for practice of analysis using the Visual Basic Application (VBA) macro programming skills. The idea is to demonstrate and showcase the knowledge and understanding of basic stock analysis using VBA, which could be helpful in guiding stock trading decisions.
Objective
Refactor the All Stocks Analysis code by improving on it to be cleaner and efficient. In other words, achieve the same task with a more effective approach or methodology.
Code Plan
(1)	Set the output worksheet for Module 2 Challenge active. In this case “Module2Challenge”.

(2)	Create Header Row. Top left corner Cell A1 to indicate the year of analysis from the Input Box for the Year Value.

(3)	Create an Array of 12 Ticker symbols as String variable, Ticker Volume, Starting Price and Ending Price

(4)	Set Ticker Index to zero. Set Ticker Volume to zero

(5)	Create a loop that will go through the rows in data worksheet starting at row 2 all the way to end
In the loop find:
		tickerVolume()
		startingPrice()
		endingPrice()

(6)	Increment Ticker Index by +1

(7)	Output the collected data to the “Module2Challenge” worksheet
Approach
The core functionality of the program resides with the collection of three specific data in a loop through four specified arrays with “tickerIndex”.
Conclusion / Visualization of the observation
The program works effectively well despite underperforming stocks. However, it needs to be tested against larger size data to determine accuracy and reliability.

