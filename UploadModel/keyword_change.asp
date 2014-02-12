<html>

<%
FileID = Request("ID")
Set Conn = Server.CreateObject("ADODB.Connection")
SQL = "SELECT KW_Test.KeyWords FROM KW_Test WHERE KW_Test.FileID = '" & FileID & "'"
connstr = 'Your database connection sentence'
Conn.Open connstr
Set DS = Conn.Execute(SQL)

dim array(30)
dim t
t=0

do while DS.eof=false
array(t)= DS(0)
t = t+1
DS.movenext
Loop
DS.close

dim i
i=0

%>


<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>change keyword</title>

<script>
function displayResult()
{
var x=document.getElementById("mySelect");
var option=document.createElement("option");
option.text=document.getElementById("texttest").value;

try
  {
  // for IE earlier than version 8
  x.add(option,x.options[null]);
  }
catch (e)
  {
  x.add(option,null);
  }
}
</script>

</head>

<body>




<form name="input" action="submission.asp" method="get">

<table border="2">
	
	<tr>
		<td align="center">&nbsp;</td>
		<td align="center" width="14%">Target</td>
		<td align="center" width="14%">Applicants<br>&amp; Admits</td>
		<td align="center" width="14%">New<br>Student</td>
		<td align="center" width="14%">All<br>Student</td>
		<td align="center" width="14%">Degree<br>Recipients</td>
	</tr>
		
			<tr><td>Industry and Market</font></td>
				<td align=center><input type=checkbox name="myChk" value="1" <%if InStr("," & Join(array, ",") & ",",","& "Cell 1" & ",") > 0 then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="2" <%if InStr("," & Join(array, ",") & ",",","& "Cell 2" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="3" <%if InStr("," & Join(array, ",") & ",",","& "Cell 3" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="4" <%if InStr("," & Join(array, ",") & ",",","& "Cell 4" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="5" <%if InStr("," & Join(array, ",") & ",",","& "Cell 5" & ",") > 0  then%>checked<%else%><%end if%>></td>
						
			</tr>
			
			<tr><td>Benchmarks, Competition</font></td>
				<td align=center><input type=checkbox name="myChk" value="6" <%if InStr("," & Join(array, ",") & ",",","& "Cell 6" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="7" <%if InStr("," & Join(array, ",") & ",",","& "Cell 7" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="8" <%if InStr("," & Join(array, ",") & ",",","& "Cell 8" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="9" <%if InStr("," & Join(array, ",") & ",",","& "Cell 9" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="10" <%if InStr("," & Join(array, ",") & ",",","& "Cell 10" & ",") > 0  then%>checked<%else%><%end if%>></td>
				
			</tr>
			
			<tr><td>Profile & Patterns</font></td>
				<td align=center><input type=checkbox name="myChk" value="11" <%if InStr("," & Join(array, ",") & ",",","& "Cell 11" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="12" <%if InStr("," & Join(array, ",") & ",",","& "Cell 12" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="13" <%if InStr("," & Join(array, ",") & ",",","& "Cell 13" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="14" <%if InStr("," & Join(array, ",") & ",",","& "Cell 14" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="15" <%if InStr("," & Join(array, ",") & ",",","& "Cell 15" & ",") > 0  then%>checked<%else%><%end if%>></td>
			</tr>
			
			<tr><td>Student Perceptions</font></td>
				<td align=center><input type=checkbox name="myChk" value="16" <%if InStr("," & Join(array, ",") & ",",","& "Cell 16" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="17" <%if InStr("," & Join(array, ",") & ",",","& "Cell 17" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="18" <%if InStr("," & Join(array, ",") & ",",","& "Cell 18" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="19" <%if InStr("," & Join(array, ",") & ",",","& "Cell 19" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="20" <%if InStr("," & Join(array, ",") & ",",","& "Cell 20" & ",") > 0  then%>checked<%else%><%end if%>></td>
			</tr>
			
			<tr><td>Progress, Performance, Outcomes</font></td>
				<td align=center><input type=checkbox name="myChk" value="21" <%if InStr("," & Join(array, ",") & ",",","& "Cell 21" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="22" <%if InStr("," & Join(array, ",") & ",",","& "Cell 22" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="23" <%if InStr("," & Join(array, ",") & ",",","& "Cell 3" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="24" <%if InStr("," & Join(array, ",") & ",",","& "Cell 24" & ",") > 0  then%>checked<%else%><%end if%>></td>
				<td align=center><input type=checkbox name="myChk" value="25" <%if InStr("," & Join(array, ",") & ",",","& "Cell 25" & ",") > 0  then%>checked<%else%><%end if%>></td>
			
			</tr>
	</table>		
	
			<select id="mySelect" name = "mySlt" size=50 multiple>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Admission" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Admission">Admission</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Admission Summary" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt"  ="Admission Summary">Admission Summary</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Adult" & ",") > 0 then%>selected<%else%><%end if%> name= "mySlt" value ="Adult">Adult</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Advising" & ",") > 0  then%>selected<%end if%> name= "mySlt" value ="Advising">Advising</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Brand" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Brand">Brand</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Brown Bag" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Brown Bag">Brown Bag</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Brown Bag Powerpoint List" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Brown Bag Powerpoint List">Brown Bag Powerpoint Listg</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Budget-to-actuals enrollment report" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Budget-to-actuals enrollment report">Budget-to-actuals enrollment report</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Career" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Career">Career</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Catholic" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Catholic">Catholic</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Chicago Public Schools (CPS)" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Chicago Public Schools (CPS)">Chicago Public Schools (CPS)</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Communication" & ",") > 0  then%>selected<%else%><%end if%> value ="Communication">Communication</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Community College Report Card" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Community College Report Card">Community College Report Card</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Conference Presentation" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Conference Presentation">Conference Presentation</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Engagement" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Engagement">Engagement</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Enrollment Summary" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Enrollment Summary">Enrollment Summary</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Enrollment Update" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Enrollment Update">Enrollment Update</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Faculty/Staff Grid" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Faculty/Staff Grid">Faculty/Staff Grid</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Freshman" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Freshman">Freshman</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Freshman Admission Summary" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Freshman Admission Summary">Freshman Admission Summary</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Graduation" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Graduation">Graduation</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Hurricane Chart" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Hurricane Chart">Hurricane Chart</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Information Infrastructure Draft Folder" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Information Infrastructure Draft Folder">Information Infrastructure Draft Folder</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Institutional Studies Draft Folder" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Institutional Studies Draft Folder">Institutional Studies Draft Folder</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Internal" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Internal">Internal</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "IRMA Calendar" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="IRMA Calendar">IRMA Calendar</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "IRMA Minute" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="IRMA Minute">IRMA Minute</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Law" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Law">Law</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Low Income (Pell)" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Low Income (Pell)">Low Income (Pell)</option>			  
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Market Analytics Draft Folder" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Market Analytics Draft Folder">Market Analytics Draft Folder</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Market Share" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Market Share">Market Share</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Masters" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Masters">Masters</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "MS/IRA Working Folder" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="MS/IRA Working Folder">MS/IRA Working Folder</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "New Release" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="New Release">New Release</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Online" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Online">Online</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Profile/New Student" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Profile/New Student">Profile/New Student</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Program/Course Grid" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Program/Course Grid">Program/Course Grid</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Reenrollment Report" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Reenrollment Report">Reenrollment Report</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Registration Activity" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Registration Activity">Registration Activity</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Reporting Draft Folder" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Reporting Draft Folder">Reporting Draft Folder</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Retention" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Retention">Retention</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Science & Health" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Science & Health">Science & Health</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "SNL" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="SNL">SNL</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Student of color (minority, underrepresented minority, diversity)" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Student of color (minority, underrepresented minority, diversity)">Student of color (minority, underrepresented minority, diversity)</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Test Optional (ACT, SAT)" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Test Optional (ACT, SAT)">Test Optional (ACT, SAT)</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Transfer" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Transfer">Transfer</option>
			  <option <%if InStr("," & Join(array, ",") & ",",","& "Undergraduate" & ",") > 0  then%>selected<%else%><%end if%> name= "mySlt" value ="Undergraduate">Undergraduate</option>			  
			</select>			
			
			<p><b>Tip:</b> To select in multiple options, hold down the Ctrl key while selecting them</p>
		
			<!--<textarea id="texttest" name="texttest" >Delete me and Add what you want</textarea>  Add the keyword that you want--> 
			<!--<button type="button" onclick="displayResult()">Insert into the list</button>-->
			<button type="submit" name = FileID value=<%=FileID%> >Submit</button>
			<button type="reset" value="Reset">Reset</button>

		<br></br>
		

	</form>
</body>

</html>
