<div align="center">

## The Daily Newbie \- Using the DateAdd\(\) Function


</div>

### Description

Explains how to use the Visual Basic DateAdd() function to add and subtract dates.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Beginner
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-the-daily-newbie-using-the-dateadd-function__1-22924/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type"
content="text/html; charset=iso-8859-1">
<title>Daily Newbie - 05/01/2001</title>
</head>
<body bgcolor="#FFFFFF">
<p> </p>
<p class="MsoTitle"><img width="100%" height="3"
v:shapes="_x0000_s1027"></p>
<p align="center" class="MsoTitle"><font size="7"><strong>The
Daily Newbie</strong></font></p>
<p align="center" class="MsoTitle"><strong>&#8220;To Start Things
Off Right&#8221;</strong></p>
<p align="center" class="MsoTitle"><font size="1">
May 3,
2001
</font></p>
<p align="center" class="MsoTitle"><img width="100%" height="3"
v:shapes="_x0000_s1027"></p>
<p align="center" class="MsoNormal" style="text-align:center"> </p>
<p align="center" class="MsoNormal" style="text-align:center"> </p>
<p class="MsoNormal"><font face="Arial"></font></p>
<p class="MsoNormal"><font size="2" face="Arial"></font></p>
<p class="MsoNormal"><font size="2" face="Arial"></font></p>
<p class="MsoNormal"
style="margin-left:135.0pt;text-indent:-135.0pt"><font size="2"
face="Arial"><strong>Today&#8217;s Keyword:</strong>
        </font><font
size="4" face="Arial"> DateAdd()</font></p>
<p class="MsoNormal"
style="margin-left:135.0pt;text-indent:-135.0pt"><font size="2"
face="Arial"><strong>Name Derived
From:  </strong>   </font>
 <font size="2" face="Arial">"Date Addition"</a></i> </em></font></p>
 </p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Used for: </strong>
Adding a specified time period to a date value.</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>VB Help Description: </strong>   Returns a Variant (Date) containing a date to which a specified time interval has been added.
</font></p>
<font size="2" face="Arial"><strong>Plain
English: </strong>Allows you to add a specified number of seconds, minutes, hours, days, weeks, months, quarters, or years to a date.<br><br>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Syntax:  </strong>       Val=DateAdd(Interval, Count, BaseDate)</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Usage:  </strong>        dtmNewDate = DateAdd("M", 8, "01/12/2000")</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Parameters:  </strong>
<br>
<font face = "arial" size="2">
<li><b>Interval</b> - The unit that you want to add to the Base Date. This can be:
	<blockquote>
		<blockquote>
	<li>s - Seconds
	<li>n - Minutes
	<li>h - Hours
	<li>d - Days
	<li>w - Weeks
	<li>m - Months
	<li>q - Quarter
	<li>yyyy - Year
		</blockquote>
	</blockquote>
<li><b>Count</b> - The number of days, weeks, etc. that you wish to add to the date.
<li><b>BaseDate</b> - The date that the interval is to be added to.
Example: <br>
<br>
To add two days to today's date:
<br><br>
<blockquote>
<code><font size="2">MsgBox DateAdd("d", 2, Date)</font></code>
</blockquote>
</font>
</font></p>
If you have not read the Daily Newbie on how VB stores date format, you may want to review it now <a href="http://www.planetsourcecode.com/xq/ASP/txtCodeId.22876/lngWId.1/qx/vb/scripts/ShowCode.htm"> by clicking here.</a>
 <br><br>
<br>
Today's code snippet prints a annual schedule of maintenance dates for a piece of equipment that must be maintained every 45 days.
</font></p>
<p class="MsoNormal"
style="margin-left:135.35pt;text-indent:-135.35pt"><font size="2"
face="Arial"><strong>Copy & Paste Code:</strong></font></p>
  <p class="MsoNormal"
  style="margin-left:135.35pt;text-indent:-135.35pt"><font
  size="2" face="Arial"></font></p>
    <pre>
<font size="2" face="Arial"><code></code></font></pre>
    <pre
    style="margin-left:1.25in;text-indent:.35pt;tab-stops:45.8pt 91.6pt 183.2pt 229.0pt 274.8pt 320.6pt 366.4pt 412.2pt 458.0pt 503.8pt 549.6pt 595.4pt 641.2pt 687.0pt 732.8pt"><font
size="3" face="Arial"><code>
<br><br>
Dim dtmStartDate As Date  'Holds original date
Dim dtmMaintDate As Date  'Holds incremented date
dtmStartDate = InputBox("Enter the date of the first maintenance:")
dtmMaintDate = dtmStartDate 'Start increment date at entered date
Debug.Print "Maintenance Schedule for Widget"
Debug.Print "================================"
<br><br>
<code>
Do
  dtmMaintDate =<b> DateAdd("d", 45, dtmMaintDate)</b>
	<br>
'		Print to the debug window (press "Ctrl" Key + "G" Key
'		to view the debug window
	<br>
  Debug.Print dtmMaintDate
'	keep going until the current maintenance date is
'	greater than the start date plus one year
Loop Until dtmMaintDate ><b> DateAdd("yyyy", 1, dtmStartDate)</b>
<br><br>
<br><br>
				</code></font></pre>
 <p class="MsoNormal"
 style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"> </p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Notes: </strong></font></p>
<font size="2" face="Arial">
The DateAdd function is extremely useful when you are writing time sensitive applications. You can accomplish with one function call what would take many, many lines of code without it.
<br><br><b>
Some general notes on DateAdd:
</b>
<br><br>
<li>Despite its name, you can subtract dates with DateAdd as well. This is accomplished by simply adding a negative number in the Count parameter.
<br>
<br>
<blockquote>
<code><font size="2">MsgBox DateAdd("d", -2, Date)</font></code>
</blockquote>
<br>
<li>DateAdd is aware of all of the calendar weirdness such as leap years. Using it to add an interval of one day to Feb. 28, 2001 will yield Feb. 29, while it will yield March 1 for 2002.
</body>
</html>

