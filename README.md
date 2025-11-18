8 Queens Puzzle in Excel VBA
This project provides a solution to the classic 8 Queens puzzle implemented in Excel VBA.
The goal of the puzzle is to place eight queens on a standard 8 x 8 chessboard such that no two queens attack each other.
________________________________________
 Features
•	Solves the 8 Queens puzzle using recursive backtracking
•	Visual board representation directly inside Excel
•	Highlights valid solutions
•	Easy to extend for different board sizes
•	Fully written in native VBA (no external dependencies)
________________________________________
How to Use
1.	Open the workbook (*.xlsm).
2.	Ensure macros are enabled.
3.	Make a double-click on the first row to launch a userForm.
4.	Click the Generate all solution button.
5.	Click the Show a random Solution button.
________________________________________
How It Works
The solver uses a classic backtracking algorithm:
1.	Place a queen in a safe column on the current row
2.	Move to the next row
3.	If stuck, backtrack and try a different column
4.	Stop when all queens are placed
The algorithm explores all valid solutions systematically.


