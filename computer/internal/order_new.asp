<html>
<head>
<title>New Order</title>
</head>

<body bgcolor="#ffffff" link="765f00" vlink="765f00">

<a href="/default.htm"><img src="title2.jpg" border=0></a>
<p>
<center>

<%
Set OBJdbConnection = Server.CreateObject("ADODB.Connection")  
OBJdbConnection.Open "comp"
SQLQuery = "SELECT name_last, name_first FROM customers" 
Set RScustomerList = OBJdbConnection.Execute(SQLQuery) 
%>

<% Do While Not RScustomerList.EOF %>

<%=RScustomerList("first_name")%> <%=RScustomerList("first_name")%>

<% 
RScustomerList.MoveNext 
Loop 
%>

</body>
</html>
