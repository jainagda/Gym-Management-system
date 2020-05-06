 <html>
 <body>
 <%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "C:\\inetpub\\wwwroot\\fitnessclub\\Gymdb.mdb"
sql="INSERT INTO trainer(fname,lname,contact,email,gender,age,address,city,pincode,adhar,quali,work,bank,ifsc,accno,accountname)"
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("fname") & "',"
sql=sql & "'" &  Request.Form("lname") & "',"
sql=sql & "'" &  Request.Form("contact") & "',"
sql=sql & "'" &  Request.Form("email") & "',"
sql=sql & "'" &  Request.Form("gender") & "',"
sql=sql & "'" &  Request.Form("age") & "',"
sql=sql & "'" &  Request.Form("address") & "',"
sql=sql & "'" &  Request.Form("city") & "',"
sql=sql & "'" &  Request.Form("pincode") & "',"
sql=sql & "'" &  Request.Form("adhar") & "',"
sql=sql & "'" &  Request.Form("quali") & "',"
sql=sql & "'" &  Request.Form("work") & "',"
sql=sql & "'" &  Request.Form("bank") & "',"
sql=sql & "'" &  Request.Form("ifsc") & "',"
sql=sql & "'" &  Request.Form("accno") & "',"
sql=sql & "'" &  Request.Form("accountname") & "')"
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