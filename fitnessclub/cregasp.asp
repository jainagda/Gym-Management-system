 <html>
 <body>
 <%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "C:\\inetpub\\wwwroot\\fitnessclub\\Gymdb.mdb"
sql="INSERT INTO Creg (fname,lname,contact,email,gender,"
sql=sql & "cweight,height,address,city)"
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("fname") & "',"
sql=sql & "'" & Request.Form("lname") & "',"
sql=sql & "'" & Request.Form("contact") & "',"
sql=sql & "'" & Request.Form("email") & "',"
sql=sql & "'" & Request.Form("gender") & "',"
sql=sql & "'" & Request.Form("cweight") & "',"
sql=sql & "'" & Request.Form("height") & "',"
sql=sql & "'" & Request.Form("address") & "',"
sql=sql & "'" & Request.Form("city") & "')"
on error resume next
conn.Execute sql,recaffected
if err<>0 then
  Response.Write("No update permissions!")
else
 Response.Redirect "homepage3.html"
end if
conn.close
%>
</body>
</html>