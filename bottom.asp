<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn_bargain.asp" -->
<%
Dim rsc
Dim rsc_numRows

Set rsc = Server.CreateObject("ADODB.Recordset")
rsc.ActiveConnection = MM_conn_bargain_STRING
rsc.Source = "SELECT * FROM t_company"
rsc.CursorType = 0
rsc.CursorLocation = 2
rsc.LockType = 1
rsc.Open()

rsc_numRows = 0
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="Config/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<hr size="1" noshade>
<div align="center">电话:<%=(rsc.Fields.Item("c_tel").Value)%> 地址:<%=(rsc.Fields.Item("c_address").Value)%> 
  邮政编码:<%=(rsc.Fields.Item("c_code").Value)%><br>
  版权所有<%=(rsc.Fields.Item("c_name").Value)%></div>
</body>
</html>
<%
rsc.Close()
Set rsc = Nothing
%>
