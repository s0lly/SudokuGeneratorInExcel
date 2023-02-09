# SudokuGenerator
A simple Sudoku Generator in Excel, with major logic driven by Excel formulae


This Excel file allows you to gain an intuitive understanding of how Sudoku puzzles can be generated in the basic setting of an Excel workbook.
The model uses Excel functions for the entirety of the underlying logic, with minimal VBA code to cycle through iterations of solutions / puzzles based on that logic. The logic / model is separated into two components:
 - The "Setup" tab generates a 9x9 cube of numbers that represent a Sudoku solution - not every set of 9x9 numbers is a solution, so a set of calculations are used to determine a valid solution.
     - The logic can get "stuck" on a non-solution path - therefore I use a macro to cycle through iterations until a valid solution can be found.
 - The "Solution" tab generates a Sudoku Puzzle based on the solution generated in the "Setup" tab. To do this, it begins with the solution and works backwards, removing elements one-by-one from the solution, until it cannot remove any more.
     - The logic for removing elements is relatively simple, and therefore only "Easy" to "Easy / Medium" puzzles can be generated.
     - This logic path can again get "stuck" and therefore another simple macro cycles through iterations until a reasonably good puzzle can be found (here defined as a puzzle with 36 or fewer elements remaining).
I hope this is of interest to those looking to understand how to generate somewhat complex puzzles, as well as in crafting up their own Excel models more generally!


Instructions

First, you need to go to the sheet named "Play". Next, you need to run the macro called "Generate Sudoku".
Voila! You can now try solve the puzzle at your own pace.

NOTE: this macro may take a short while to run - depending on how it quickly it can converge to a reasonable solution. It should not take more than 1 minute, and will be much faster.


by s0lly
https://www.youtube.com/c/s0lly
https://www.instagram.com/s0lly.gaming/
https://twitter.com/s0lly
https://www.twitch.tv/s0llygaming

THE MODEL IS PROVIDED WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT
IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE MODEL OR THE USE OR OTHER DEALINGS IN THE MODEL.
![image](https://user-images.githubusercontent.com/39304402/217695685-55e87863-01b6-4f77-a23e-f1df52ae5f10.png)

