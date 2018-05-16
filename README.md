VBA:
http://www.cpearson.com/Excel/Topic.aspx
https://peltiertech.com/
https://excelmacromastery.com/vba-articles/

Excel:
http://contextures.com/
https://www.youtube.com/user/ExcelIsFun

macro name standards
-seperate with _
-lowercase


- options to run a macro
	shortcut - (cannot use xls shortcuts; only a-z/A-Z)
	event triggering (click of button or on open of excel)
	developer mode - run button (F5)
	
	
	
absolute reference recording
relative reference recording

two types for procedures in macros - sub procedures and function procedures
macros always record sub procedures

function procedures are meant for calculations (e.g. excel formulae)
sub procedures are meant for actions (e.g. bold, color, italic)

Instruction Sheet as First Sheet for a macro file

Chapter1 - 
custom lists
recording/running macros

Chapter2 - 
DataTypes-
Byte, Integer, Long
Double
String
Date

sub procedures 
Syntax:

	Sub <procedure_name> ()
		<statements...>
	End Sub
	
Function Procedures
Syntax:

	Function <procedure_name>(variables with datatypes) as <datatypes>
		<statements.....>
	End Function
	
	
	
Chapter3
-assign macro

Cell Hierarchy
APPLICATION
	WORKBOOKS (Collection)
	WORKBOOK (Object)
		WORKSHEETS (Collection)
		WORKSHEET(Object)	
			RANGES (Collection)
			RANGE (Object)

			

Chapter6
Vlookup
WorkSheetFunction. //All Excel Formulas


Chapter 7
Looping types
	For Next 
		
			FOR <loop_counter_variable> = <startingValue> TO <endingValue>
				<statements...>
			Next <loop_counter_variable>
			
	For each
	Do loops
	
	
Chapter 8
 Pivot Table - for reporting/summarizing/data vizualisation
 - Headers are mandatory and unique
 - Avoid Merged Headers
 - Numeric cells should not be empty or have text
 - Clean up data
	
	
	
Chapter 9 - Outlook Integration

VBE - Tools - References - select Checkbox MicroSoft Outlook Library

	

	
