<%
Dim oConn, oRs
Dim qry, connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
Dim damndata

db_server = "servername"
db_name = "databasename"
db_username = "databaseusername"
db_userpassword = "databaseuserpassword"
tablename = "databasetablename"
'fieldname = "street1"

response.write("Hello World! <br />")

connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open connectstr
Set damndata = Server.CreateObject("ADODB.Recordset")
qry = "SELECT software, license, ownername FROM " & tablename
damndata.Open qry, connectstr

	response.write("<html>")
	response.write("<head>")
	response.write("<script src='http://code.jquery.com/jquery-latest.min.js'></script>")
	response.write("<link rel='stylesheet' type='text/css' href='./css/table.css' />")
	response.write("</head>")
	response.write("<span><table border='1' align='left'>")
	response.write("<tr align='left'><th>Software</th><th>License</th><th>Ownername</th></tr>")
Do while not damndata.EOF

	response.write("<tr><td>" & damndata("Software") & "</span></td><td>" & damndata("License") & "</span></td><td>" & damndata("Ownername") & "</span></td><tr>")
	damndata.MoveNext
Loop
	response.write("</table></span>")
	response.write("</html>")

Set damndata = nothing
Set oConn = nothing

%>
