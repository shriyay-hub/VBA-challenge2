Sai Shriya Yenugu's VBA challenge Project

Explanation:
Main Loop and Quarter Handling:
The script loops through each worksheet (ws) in the workbook.
Within each worksheet loop, it further loops through each quarter (1 to 4).

Quarter Detection:
It uses a helper function GetQuarter to determine which quarter a particular date belongs to based on its month.

Data Processing:
For each row in the current quarter, it calculates the quarterly change, percent change, and accumulates the total volume.
It checks if the ticker symbol changes to finalize the calculations for the current ticker.

Summary Calculation:
After processing each quarter, it updates the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume across all quarters.
Output:

Finally, it outputs the results in the first worksheet (ThisWorkbook.Worksheets(1)) as specified.

![Screenshot 2024-07-08 at 7 33 24 PM](https://github.com/shriyay-hub/VBA-challenge2/assets/173106830/589e627d-50ba-49a8-8b1d-667ae7dcde76)
![Screenshot 2024-07-08 at 7 27 19 PM](https://github.com/shriyay-hub/VBA-challenge2/assets/173106830/f04491da-1061-4bef-b672-0a7c383ac748)
