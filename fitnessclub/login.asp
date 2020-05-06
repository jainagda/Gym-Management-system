<html>
<body>
<%
	'Save the entered username and password
	Username = Request.Form("uname")	
	Password = Request.Form("psw")
	
	'Build connection with database
	set conn=Server.CreateObject("ADODB.Connection") 
	conn.Provider="Microsoft.Jet.OLEDB.4.0" 
	conn.Open(Server.Mappath("Gymdb.mdb")) 
 
set rs=Server.CreateObject("ADODB.recordset")
 sql="SELECT * FROM Login where login='"&Username&"' AND password='"&Password&"'" 
 rs.Open sql,conn,1
	'If there is no record with the entered username, close connection
	'and go back to index1 with QueryString'
	if rs.recordcount=0 then
	Response.Redirect "homepage3.html"
	
	else
	Response.Write "<script language=""javascript"">alert('successfully logged in');</script>"
	Response.Redirect "homepage1.html"
	end if
	Conn.Close
		
	
%>


</body>
</html>