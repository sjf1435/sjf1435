<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn_bargain.asp" -->
<!--#include file="chk_login.asp" -->
<!--#include file="chk_level3.asp" -->
<%
Dim rsa
Dim rsa_numRows

Set rsa = Server.CreateObject("ADODB.Recordset")
rsa.ActiveConnection = MM_conn_bargain_STRING
rsa.Source = "SELECT * FROM t_admin ORDER BY a_id DESC"
rsa.CursorType = 0
rsa.CursorLocation = 2
rsa.LockType = 1
rsa.Open()

rsa_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsa_numRows = rsa_numRows + Repeat1__numRows
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="Config/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body>
<h1>用户管理</h1>
<hr size="1">
<form action="a_addsave.asp" method="post" name="form1" onSubmit="MM_validateForm('a_username','','R','a_password','','R');return document.MM_returnValue">
  <table width="100%" align="center">
    <tr> 
      <td width="30%" height="30" align="right" nowrap class="bgcolor-left">用户名称:</td>
      <td height="30" class="bgcolor-right"> <input type="text" name="a_username" value="" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">用户密码:</td>
      <td height="30" class="bgcolor-right"> <input type="text" name="a_password" value="" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">用户等级:</td>
      <td height="30" class="bgcolor-right"> <select name="a_level">
          <option value="1" selected>录入员</option>
          <option value="2">审核员</option>
          <option value="3">超级管理员</option>
        </select> </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap>　</td>
      <td height="30"> <input type="submit" value="插入记录"> </td>
    </tr>
  </table>
</form>
<p>　</p>
<table width="100%" border="0" align="center" bordercolor="#999999" bgcolor="#FFFFFF" class="border-all">
  <tr align="center"> 
    <td height="30"><strong>用户名</strong></td>
    <td height="30"><strong>密码</strong></td>
    <td height="30"><strong>等级</strong></td>
    <td><strong>操作</strong></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rsa.EOF)) %>
  <tr align="center"> 
    <td height="30"><%=(rsa.Fields.Item("a_username").Value)%></td>
    <td height="30"><%=(rsa.Fields.Item("a_password").Value)%></td>
    <td height="30">
	<% 
	dim mm_level
	mm_level=cint(rsa.Fields.Item("a_level").Value)
	if mm_level = 1 then
	response.Write("录入员")
	end if
	if mm_level = 2 then
	response.Write("审核员")
	end if
	if mm_level = 3 then
	response.Write("超级管理员")
	end if
	%>
    </td>
    <td><a href="a_updata.asp?a_id=<%=(rsa.Fields.Item("a_id").Value)%>">修改</a> 
      <a href="a_del.asp?a_id=<%=(rsa.Fields.Item("a_id").Value)%>" onClick="javascript:return confirm('请确认要删除该用户 ')">删除</a> </td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsa.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
rsa.Close()
Set rsa = Nothing
%>