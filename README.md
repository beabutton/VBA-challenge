# VBA-challenge

I began working on this challenge while we were still learning Excel and working on the Excel challenge. Due to that I tried to do some independent research into VBA utalizing the StackOverflow Q&A for various parts of the assignment. I found it extremely helpful to work backwards through something that works or mostly works to formulate my own methods. I utalized the below resources as well as a brief conversation about strategy with a fellow classmate, Roxanne. 

https://stackoverflow.com/questions/57279654/is-there-something-in-my-vba-code-that-is-causing-me-to-get-an-overflow-error

https://stackoverflow.com/questions/52122844/how-to-apply-a-vba-code-to-every-page-in-a-workbook-mine-does-part-of-the-code

https://stackoverflow.com/questions/62471422/vba-loop-how-to-get-ticker-symbols-into-ticker-column

https://stackoverflow.com/questions/20339067/excel-vba-overflow-error

https://excelchamps.com/vba/autofit/

Roxanne
If cells (c,11).value=application.worksheetfunction.max(ws.range(“K2:K”& yearly change lastrow)) then 
	cells(2,17) = cells(c,9)
	cells(2,17).numberformat="0.00%"
 
