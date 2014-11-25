<%
'Author: Robert Holland
'Purpose: Display Cerner CCL files that may need to be checked due to the changes for iview.
'Currently working 201411241903 on PHX04134.

'Connection string
<!--#include file="ivdb.asp"-->

Dim oConn, oRs
Dim qry, connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
Dim dbdata

db_server = "phx04134.bhs.bannerhealth.com"
db_name = "iviewtest"
db_username = "iview"
db_userpassword = "3y3vi3W"
tablename = "iview_terms5"
fieldname = "street1"
connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open connectstr
Set dbdata = Server.CreateObject("ADODB.Recordset")

qry = "select distinct(program), code_snippet from " & tablename & " where term = 'Weight (kg)'"

dbdata.Open qry, connectstr
	response.write("<html>")
	response.write("<head>")
	response.write("<script src='http://code.jquery.com/jquery-latest.min.js'></script>")
	response.write("<link rel='stylesheet' type='text/css' href='./css/table.css' />")
	response.write("</head>")
	response.write("<span><table border='1' align='left'>")
	response.write("<table border='1' align='left'>")
	response.write("<tr align='left'><th>Program</th><th>Code_Snippet</th></tr>")
Do while not dbdata.EOF

	response.write("<tr><td>" & dbdata("program") & "</td><td>" & dbdata("code_snippet") & "</td><td><center><input type='checkbox' value='" & dbdata("program") & "' name='ckbox[]'></center></td><tr>")
	dbdata.MoveNext
Loop
	response.write("</table></span>")
	response.write("</body><html>")
	Set dbdata = nothing
	Set oConn = nothing

'Connection to SQL Server
%>


