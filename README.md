# Module 2 - Hospital Operations

### Download Instructions
To run this solution, download the `Gurobi_Optimization` Python file to your local computer and ensure it is in the file path: `C:\Users\USER\Documents` where USER is your username. Ensure you locally download the two Excel files: `MSE 433 Hospital Operations Excel Tool` (ensure it is a .xlsm file since it has Macros enabled) and `TempOptiData` to the same file path. Failure to follow these steps will corrupt the VBA Macro file and render the solution unusable. **DO NOT** download this to the OneDrive documents folder. Only the local C:\ Documents folder.

### Initial Setup
1. Open the Gurobi Optimization python file on VS Code and ensure you click 'Trust' on the file if prompted.
2. On Line 7 of the python script, change the `username` variable to your computer username (Same as USER above in the download instructions). Save the Python file and close VS Code.
3. Open the MSE 433 Hospital Operations Excel Tool (Macro enabled). Do not turn on auto save. To save any work, click `CTRL+S`.
4. On the Excel file, open the Visual Basic code editor and navigate to `Module 2` and update the userName to your computer username. click `CTRL+S` to save this change and exit the Visual Basic code editor. Save the Excel file manually.

<img width="649" height="104" alt="image" src="https://github.com/user-attachments/assets/db2f35c0-cfc3-43c7-af23-c2c4a8ea2e49" />

## Steps to Run Solution
1. Navigate to the Dept Breakdown Excel sheet. This is where you will be acting as a hospital coordinator to specify the number of people arriving/departing the 4 departments and see recommendations.
2. Click the `NEW HOUR DATA ENTRY` button and input the number of people who arrived to each of the departments. **Note:** If you are at Hour 0, nobody is in the hospital yet, so you should not be putting any departures during Hour 0. Any hour after this, you can have departures, but ensure the departures from a department do not exceed the number of patients in that department. Click `Submit` on the form to display your provided info in the Dept Breakdown sheet.
3. Click the `RECOMMEND STAFFING` button to trigger the optimization model to run. This may take 10-15 seconds since the Macro button makes a call to Python where the Gurobi optimization code is and returns the results back to the Dept Breakdown sheet. Do not click out of the Excel file or anything in the file as this is happening.
4. Once the optimization runs, view the updated Dept Breakdown results and the score breakdown. To see more insights, go to the `Opti Data` and `Insights` tabs (do not modify anything in these sheets - just read).
5. To see visualizations of the results, return back to the Dept Breakdown sheet and click the `VISUALIZE DATA` button to see some graphs to summarize the current hospital performance.
6. Once you are done with the tool or wish to restart, click the `RESET` button on the Dept Breakdown sheet and the `RESET` button on the Trasfer Requests sheet.
   
**Note:** The Trasfer Requests sheet is just for the tool to randomly allocate transfers between departments. Do not modify this sheet.

### Solution Interface
**Dept Breakdown Sheet:**

<img width="1487" height="492" alt="image" src="https://github.com/user-attachments/assets/823ef147-7f19-4209-8b6b-726118ae2ae8"/>

**New Hour Data Entry Form:**

<img width="833" height="765" alt="image" src="https://github.com/user-attachments/assets/0b6f8bb5-4346-42df-8f62-e04b36c489a7"/>

**Insights Sheet:**

<img width="836" height="509" alt="image" src="https://github.com/user-attachments/assets/e4c57cc7-da4f-4a28-9615-a804d70dca9d"/>

**Data Visualization Sheet:**

<img width="659" height="488" alt="image" src="https://github.com/user-attachments/assets/9b720a59-89df-45a3-8df1-3c3cc264c51d"/>

