<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn_bargain.asp" -->
<!--#include file="chk_login.asp" -->
<!--#include file="chk_level3.asp" -->
<%

if(request.form("bt_id") <> "") then 
Command1__mmid = request.form("bt_id")
else
response.Write("参数错误")
response.End
end if

if(request.form("bt_name") <> "") then 
Command1__mmname = request.form("bt_name")
else
response.Write("参数错误")
response.End
end if

%>
<%

set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = MM_conn_bargain_STRING
Command1.CommandText = "UPDATE t_btype  SET bt_name = '" + Replace(Command1__mmname, "'", "''") + "'  WHERE bt_id = " + Replace(Command1__mmid, "'", "''") + " "
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<meta http-equiv="refresh" content="1;URL=b_type.asp">
</head>

<body>
更新成功，等待自动返回... 
</body>
</html>
