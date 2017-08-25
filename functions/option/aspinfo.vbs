<%
'=======================================================================
'ASPの設定情報を出力する
'=======================================================================
Sub aspinfo
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "DTD/xhtml1-transitional.dtd">
<html><head>
<style type="text/css"><!--
body {background-color: #ffffff; color: #000000;}
body, td, th, h1, h2 {font-family: sans-serif;}
pre {margin: 0px; font-family: monospace;}
a:link {color: #000099; text-decoration: none; background-color: #ffffff;}
a:hover {text-decoration: underline;}
table {border-collapse: collapse;}
.center {text-align: center;}
.center table { margin-left: auto; margin-right: auto; text-align: left;}
.center th { text-align: center !important; }
td, th { border: 1px solid #000000; font-size: 75%; vertical-align: baseline;}
h1 {font-size: 150%;}
h2 {font-size: 125%;}
.p {text-align: left;}
.e {background-color: #ccccff; font-weight: bold; color: #000000;}
.h {background-color: #9999cc; font-weight: bold; color: #000000;}
.v {background-color: #cccccc; color: #000000;}
i {color: #666666; background-color: #cccccc;}
img {float: right; border: 0px;}
hr {width: 600px; background-color: #cccccc; border: 0px; height: 1px; color: #000000;}
//--></style>
<title>aspinfo()</title></head>
<body><div class="center">
<table border="0" cellpadding="3" width="600">
<tr class="h"><td>
<h1 class="p">ASP</h1>
</td></tr>
</table><br />
<h2>Request</h2>
<table border="0" cellpadding="3" width="600">
<tr class="h"><th colspan="2">ServerVariables</th></tr>
<%
Dim key
For Each key In Request.ServerVariables
    Response.Write "<tr>" & vbCrlf
    Response.Write "<td class=""e"">" & key & "</td>" & vbCrLf
    Response.Write "<td class=""v"">" & Request.ServerVariables(key) & "</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
Next
%>
<tr class="h"><th colspan="2">ClientCertificate</th></tr>
<%
For Each key In Request.ClientCertificate
    Response.Write "<tr>" & vbCrlf
    Response.Write "<td class=""e"">" & key & "</td>" & vbCrLf
    Response.Write "<td class=""v"">" & Request.ClientCertificate(key) & "</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
Next
%>

</table><br />
<h2>Application</h2>
<table border="0" cellpadding="3" width="600">
<tr class="h"><th colspan="2">Contents</th></tr>
<%
For Each key In Application.Contents
    Response.Write "<tr>" & vbCrlf
    Response.Write "<td class=""e"">" & key & "</td>" & vbCrLf
    Response.Write "<td class=""v"">" & Application.Contents(key) & "</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
Next
%>
</table><br />
</div></body></html>
<%
End Sub
%>
