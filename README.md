java c
Advanced Excel - Topic 3
• DATE functions 
• Adding and Counting with Criteria: SUMIF, COUNTIF; SUMIFS, COUNTIFS DATE functions 
Activity 3a - Date functions and formats 
Open a   new spreadsheet to use to   experiment   with   date   calculations   and formats.   Save this spreadsheet as Dates.xlsx and   use it   to   try   out   the   following   activities.
1       Calculate the number of days between two dates 
Enter two   dates   as   shown.
This   is a   simple
subtraction formula   -
just subtract the   earlier         date from the   later   date.
Note:             If the   result   looks like a date and   not a   number, change   the   format   of the   cell to   General   or Number
2 Display the Current Date 
To display the current   date,   use   the      TODAY function. It   will   automatically
update to show the   current   date   if   you open the   workbook   on   a
different   day.
In   cell A5 type   in   =TODAY() 
Note:          Keyboard   shortcuts   to   enter the   current   date   as   a   static   value   (one   that   will   not   update):
Ctrl + ; (PC   or   MAC)
3 Date formats 
Try these different date   formats   (Home tab, Number group)   :
•             Short   Date
•               Long   Date
•             Click More Number Formats and   explore the   options
4 Calculate a person’s age 
A video   has   been supplied showing   how to calculate age   (Calculating Age.mp4)
Create a   new worksheet   in your file and   name the   worksheet Topic 3 (double click or   right click   on the   worksheet   tab   to   rename   it)
Calculate your age in years: 
a       In   cell A2,   enter   the   date   of your   birthday   (the   date   format   does   not   matter)
b       To subtract your   birthdate from   today’s   date,   enter   the   following   in   cell B2 =TODAY()-A2
c            Convert the   result to a   number   (number format)
d       Convert the   result to   years   by   dividing   by   the   number   of   days   in   a   year. To accurately calculate years, divide by 365.25 to allow for leap years. The result will show   the   number   of total   days.
Discussion Point: When we talk about someone’s age, do we use   decimal   places?
Calculate the age of someone in years, born in 1999, whose birthday is 1 week from today: 
a         In   cell A5 enter   the   function   for   today’s   date   =TODAY()
b         In   cell B5 enter   the   date   of   their   birth,   using   a   date   in   7   days’ time   but   in   the year   1999,   ie
If   today   is   4   May, then   use   11/5/1999.   If   today   is   31   May   then   use   7/6/1999
c            In cell D5 subtract the   birthdate from the   current   date,   and   divide   by   the   number of days   in   a   year
d         Convert to   a   number
e       Use   the   INT function to   convert   and   round   down   to   a   whole   number
f Repeat - In cell D6 subtract   the   birthdate   from   the   current   date,   and   divide   by   number of days   in   a   year
g Convert to a number and format showing 0 decimal places 
Discussion Point - Why would we use the INT function instead of formatting with zero decimals? 
5 Return (display) the Year, Month or Day Number 
If a cell contains a date   and time,   you   can   use   the   following   functions   to   show   only   the year,   month   number, or day   number from the cell   containing   the   date.
•             Enter   a   date   into   cell A10 
•               Try entering the following   in   cells as   shown:   =Year(A10)
=Month(A10)   =Day(A10)
6 Return (display) the Day of the Week 
To get the weekday   number for a   date,   use the WEEKDAY function.
For example, with a date   in cell B4,   this   formula   will   show   its   weekday   number   (Sunday   =   1,   Monday   =   2, etc.)   :
What day of the   week   were   you   born   on?
7       Save   your   file.
Challenge 3a 
The TEXT function allows you to display a   number as   text   by   using   format   codes.   The   syntax for the TEXT function   is
=TEXT(Value or cell you want to format, "Format code you want to apply") 
Use the TEXT function to display   the   day   of   the week or the   name   of the   month,   as
shown   in the example.
You   have used the COUNT function to count   the   number   of   cells   that   contain   numbers,
and the COUNTA function to count   non-blank cells   (ie, the   number   of   cells   that   have   any   data   in   them)   .    You   have   used SUM to   add   numbers.
You will  代 写Advanced Excel – Topic 3Processing
代做程序编程语言 now   learn   how to count and add selectively,   by   setting   specific   criteria   to   choose what to   count   or   what to   add.
Activity 3b - the COUNTIF function 
Use the file Counting and Adding.xlsx for this activity. This   workbook   has   three   worksheets.
1 Use the Sports worksheet 
This worksheet contains a   list of students and   the   sports   each   student   plays.
The COUNTIF function counts the number of   occurences   of   criteria   that   you   specify.
In the summary area   below the   list, use the   COUNTIF   function   count   the   number   of   students who   play each   sport.
a         In   cell B35 enter:
= CO U NT I F ( $ C $ 3: $ E $ 2 9 , A 3 5 )   
•               The range is   where   you
want the function to   count,   so in   this   example,   it   will
count each   occurrence   of
the criteria   in   the   range
$C$3: $E$29. The   range
is shaded   blue   in the   image   at   the   right.
•               The criteria is what   you   want to   match.      In this
example   it   is “Archery”   in cell A35. 
b         Autofill   down   to B44 and   check the   result.
There should   be 5 people   playing Volleyball
Examples of how =COUNTIF can be used =COUNTIF(A1:A10,10) count cells equal to 10 =COUNTIF(A1:A10,”>40”) count cells greater than 40 =COUNTIF(A1:A10,”frog”) count cells containing the word “frog” =COUNTIF(A1:A10,”<”B1) count cells where the value is less than that in cell B1 
2 Use the ClassMarks worksheet 
Refer   to   the   examples   above   to   enter   correct   formulas   in   cells D31:D33 
Activity 3c - the SUMIF function 
The SUMIF function adds numbers that correspond with   (relate to)   criteria that   you   specify.
The syntax for the SUMIF function   is therefore different   to   the   syntax   for   the   COUNTIF   function.
With SUMIF you   need to specify the   location   (range of   cells)   containing   the   numbers   to   be   added. This   is   the sum range. 
=SUM   IF(range,criteria,sum   range)
3 Use the Cars worksheet 
a         Experiment   –   use the   =SUMIF function where indicated.
b       Make   changes to   the   data   to   check   that   your   formula   is   working   correctly.
4       Save   your   file.
Activity 3d - the COUNTIFS and SUMIFS functions 
COUNTIFS and SUMIFS allow you to specify   more   than   one   criterion   to   be   matched   when counting   or   adding.
Note that the order of the arguments   in   countifs   is   different   to   sumifs.
You   may find the   Function Arguments   box (Windows) or   Formula Builder (MAC)   helpful.      The   image   below shows   how to add further criteria on Windows and   MAC.
1 Use the FruitSales worksheet   in the file Counting and Adding.xlsx.
a         Experiment   using the   =SUMIFS function, which allows   more than one   set   of
criteria to   be checked, adding values where specified.    The   order   of arguments   is   different -   see   example.
b       Make   changes   to   the   data   to   check   if your   formula   is   working   correctly.
c            Save your   file.
2 Use the SportsClub worksheet
a.    Use the   =COUNTIFS function, which allows   more   than   one   criteria   to   be   checked, counting each time the criteria   occur.   See   example:  
3 Save   your   file.
Challenge Activity 3e - Practice 
Use the file Transport Costs.xlsx for   this   activity.
1       Calculate   each   person’s   age   in   column E using   the   date   shown   in   cell B1 as   the
current date.    Show the age   in years   with   no   decimal   places.   (refer to   the   provided   video if you   are   unsure)
2       Enter a   formula   in   column F to   show   the   status   of   each   person.      Show   the   word Concession if   a   person   is   aged   under   18, or   aged   50 or   older,   otherwise   show Full Fare.    Use   absolute   referencing.
3       In column G enter   a formula   to   calculate   the   fare,   based   on   the   data   below.Bus   Concession            $4.50Bus   Full   Fare                      $6.25Train   Concession      $5.30Train   Full   Fare               $7.60
4       In the summary   area   at   the   bottom   of the   spreadsheet   enter   formulas to   show   the   required   information.
5       Place your   name,   class,   ZID   and   the   automatic   file   name   in   the   header.
6       Format   the   worksheet   to   improve   clarity   and   readability.
7       Save   your   file.







         
加QQ：99515681  WX：codinghelp
