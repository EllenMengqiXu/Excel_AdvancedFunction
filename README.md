Top Advanced Excel Tips and Tricks
===

This file is about advanced excel funciton on Mac.
---

1. Advanced Transpose - when the source content changed the transpose copy will also change
Select blank size as the original size, directly type “=transpose(A1:F6)” (A1:F6 is the content that want to transpose), then shift+ctrl+enter (instead of enter)
2. Calendar picker (it is used when you need to put in date)
Developer —> add-ins —> search “date picker”
3. Slicers (make the data into table, make the table into columns)
insert —> table —> click anywhere in the data —> insert —> slicer
4. Scenario Manager - financial use - compare best worst and likely payment values
Select rate, term and principal —> what-if analysis —> scenario manager —>add best worst likely cases and show —> summary —> scenario summary —> output: payment and total
5. Convert function between different units of measurement
==CONVERT(A40, "m", "mi") (from meter to mile) - don’t forget double quotes - if you don’t know the symbol click convert and go to excel help
6. Get external data
Save link in word as txt file ended in .iqy (after paste the link and make it as link format)
Go to excel —>click top Data—> get external data—>run web query
7. Hide Cells
Right click on the cell that you want to hide—>format cell—>custom —>change 0 to ;;;
8. Delete blank rows
Highlight the columns and hit ctrl+G—>special—>blank—>ctrl + - —>shift cells up (this way is good when the entire row is blank; for some table when the entire information is incomplete go to VBA)
9. People Graph
MyAddin —> people Graph
10. Advanced filter
Copy header in a blank field —> input the criteria under headers —> go to advanced filter under data —> list range: original data; criteria range: header plus criteria range; copy to: the filter result
11. Networkdays function
=NETWORKDAYS(start_date,end_date,holiday)
12. Embedding (insert excel into word)
enter word —>insert —>object—>from file
13. Ctrl + A —> select and hightlight all selected —>find
14. Drop-Down List
Select the cells you want to add drop-down list —> data —>validation —> validation criteria allow: list —> pick up the list you want to choose —> click OK
15. Goal Seek - it is used when you set a y value and looking for x value
Go to data —> what if analysis —> goal seek —> set value: y cell; to value; by changing: x cell
16. solver —> located under data, similar with Goal seek but can add criteria.
17.  vlookup(E3, B3:C12,2,False)
Vlookup is used when you want to do match. In this example, E3 is the name, B3:C12 stands for name and ID(does not include headers), col_index_num is 2, which meansthe second column in the range(B3:C12) is what we are looking for, False means we are looking for 100% value instead of True for nearly 100% value.
Ctrl + ~ —>show formula
18. Concatenate —> add together
19. pivotTable --> count rows frequency.
choose the data you want to work on and insert pivotTable, put the title in the rows field that you want to count and put the same title in the value field but setting to count.
20. convert mixed-format (US & UK) to US format date
* column A is the mixed date
* column B: DATE(RIGHT(A2,4),MID(A2,4,2),LEFT(A2,2)) ---> UK version Date
* column C: ---> US version Date IFERROR(IF(ISNUMBER(A2),VALUE(TEXT(A2,"mm/dd/yyyy")),DATE(RIGHT(A2,4),LEFT(A2,FIND("/",A2)-1),MID(A2,FIND("/",A2)+1,2))),DATE(RIGHT(A2,4),LEFT(A2,FIND("/",A2)-1),MID(A2,FIND("/",A2)+1,1)))
* column D: IF(C2=A2,A2,B2)
21. use index to return multiple values (multiple vlookup)
* C3 is the lookup value, Column A is the same attribute as C3, Column B is what you want to return, 2 is 3(C2's column number) - 1.
* IFERROR(INDEX($B$2:$B$31133,SMALL(IF($C$3=$A$2:$A$31133,ROW($A$2:$A$31133)- MIN(ROW($A$2:$A$31133))+1,""),COLUMN()-2)), "N/A")
