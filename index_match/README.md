<h1>Dynamic Index Match and Data Validation for Automated Cell Population</h1>

### Date:
September 2025

### Background/Context:
There is a table of sales by date and shirt color. Your stakeholder asks you to get a summary of sales by color for the previous month and to deliver this output every month. Instead of copying the title column from the first table and the next months column for each ouptut, you'd rather create an automated version to pull this data.

### Objective:
Build a flexible look up formula to summarize the sales by color to that quickly flips between months without having to manually search and copy and paste the correct data. This will save time and ensure accuracy by eliminating any room for user error.

<h2>Output:</h2>

<br />
<p align="left">
<img src="https://github.com/jameszil/pictures/blob/main/Excel/indexmatch1.png?raw=true" height="90%" width="90%" alt="Steps"/>
<br />
Here is an index match formula built out to pull a specific cell from the table on the left. In column I, whichever color and date I type will automatically find that value in the table and change the Value output in column I. Now I'll populate this again in another sheet summarizing the data, and adding a data validation dropdown box to provide a straight forward experience for end users. If we imagine a much larger table with many more date and color combinations, this could be a very useful lookup tool to quickly see the sales amount for each combination.
<br />
<br />
<img src="https://github.com/jameszil/pictures/blob/main/Excel/indexmatch%20dropdown.png?raw=true" height="50%" width="50%" alt="Steps"/>
<br />
<br />
<img src="https://github.com/jameszil/pictures/blob/main/Excel/indexmatch%20output.png?raw=true" height="90%" width="90%" alt="Steps"/>
<br />
You can see when I changed the data dropdown from 01/02/25 to 01/03/25 on the new sheet, the updated values for that date populated for each color in column J.
<br />
<br />
If we report this on a monthly basis and we are only wanting to see the previous month that closed out, I can easily snap to the next month for the next refresh instead of changing all the cell formulas to reference the right cell. Even better, I can quickly go back to previous months outputs with the click of a button without having to overwrite formulas or lose older summaries.
<br />

#### Real World Application

So why wouldn't you just the the original table? If your needs are very basic, then you may not need to use an index match. Index matches are useful to look up data from larger data sets or pull summarized data from another table for different viewing purposes. They are especially useful if you want to continuously pull data from one source to another while changing one or more factors such as a date field. Often times the index match table will provide a high level view of the more complex view that it is referencing. I had a situation at one of my jobs where a person was individually keying in cells to summarize a table for a powerpoint presentation every month and having to update each cell each month. I created a hidden column with all the aliases needed to correctly reference/match those cells and then built out the exact same chart using Index Match. Then I built out the data validation cell for the month so next month they could simply switch "July" to "August" and the chart would be ready to present. This was accomplished within an hour and saved the person close to half a day's worth of work every month.

#### Why not use VLOOKUP? 
- **Lookup Direction** - Index match can search for a value in any direction, not just left to right, and you don't have to count columns to populate like a vlookup.
- **Dynamic Ranges & Column Insertion** - It doesn't break if columns are added or deleted. Vlookup is not as dynamic; if a cell is moved, index match can still find it, whereas vlookup formulas will often get confused and break or reference the wrong cells when things get shifted, resulting in incorrect data.
- **Performance** - Index match is more efficient than vlookup. Vlookup often slows down reports.
- **Multiple Criteria** - Index match is more powerful for complex referencing. For example, Index match can be used with if statements. I have used a SUMIF INDEX MATCH to get a Year to Date total from a table on a different sheet.


