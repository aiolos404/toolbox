<html>


<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>submission</title>
</head>

<body>

	

<%
myChk=Request("myChk")
mySlt=Request("mySlt")
ID=Request("FileID")


dim array
dim array2
dim i,i2
i=0
i2=0
array = split(myChk,",")
array2 = split(mySlt,",")






Set Conn = Server.CreateObject("ADODB.Connection")
connstr = 'Your database connection sentence'
Conn.Open connstr
DeleteRow = "Delete from KW_Test where KW_Test.FileID = '" & ID & "' "
Set DS = Conn.Execute(DeleteRow) 
response.write ("Deleted all old keywords!")
response.write "<br>"





do while i<=ubound(array)

set DS1 = Conn.Execute("Insert into KW_Test (FileID,KeyWords) values('" & ID & "','Cell ' + '" & Ltrim(array(i))& "' )")
response.write ("You added the keyword Cell  " & Ltrim(array(i))& " !")
response.write "<br>"
i = i+1
loop



do while i2<=ubound(array2)

set DS2 = Conn.Execute("Insert into KW_Test (FileID,KeyWords) values('" & ID & "','" & Ltrim(array2(i2)) & "')")
response.write ("You added the keyword " & Ltrim(array2(i2))& " !")
response.write "<br>"
i2 = i2+1
loop
response.write ("You are all set!")
response.write "<br>"

%>
<input type="button" name="Submit" onclick="javascript:history.back(-1)" value="Go back"> 


</body>

</html>
