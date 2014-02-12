
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Search Output</title>
</head>
<body>
<table border="2">
	<tr>
		<td>Display Name</td>
		<td>File Title</td>		
	</tr>
<%
SearchWord = Request("search")

Set Conn = Server.CreateObject("ADODB.Connection")
SQL = 'Your querry sentence'
connstr = 'Your database connection sentence'
Conn.Open connstr
Set DS = Conn.Execute(SQL)	
While not DS.EOF

%>

<form method="POST" action="keyword_change.asp">	
	<tr>
		<td><%= DS(0) %></td>
		<td><a href=keyword_change.asp?ID=<%=DS(2)%>><%=DS(1)%></a></td>					
	</tr>
</form>
<%
DS.Movenext
Wend
DS.Close
Conn.Close
%>
</table>
</body>

</html>
