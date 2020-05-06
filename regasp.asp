<html>
<body>

<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "C:\inetpub\wwwroot\planmyleave\plandb.mdb"

sql="INSERT INTO reg (fname,email,"
sql=sql & "psw,cont,gender,address)"
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("txtbxx") & "',"
sql=sql & "'" & Request.Form("mail") & "',"
sql=sql & "'" & Request.Form("psw1") & "',"
sql=sql & "'" & Request.Form("phno") & "',"
sql=sql & "'" & Request.Form("r") & "',"
sql=sql & "'" & Request.Form("add") & "')"

on error resume next
conn.Execute sql,recaffected
if err<>0 then
  Response.Write("No update permissions!")
else
  Response.Redirect "menu - Copy1.html"
end if
conn.close
%>

</body>
</html>
