<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn_bargain.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("a_username"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization="a_level"
  MM_redirectLoginSuccess="frame.asp"
  MM_redirectLoginFailed="error.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_conn_bargain_STRING
  MM_rsUser.Source = "SELECT a_username, a_password"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM t_admin WHERE a_username='" & Replace(MM_valUsername,"'","''") &"' AND a_password='" & Replace(Request.Form("a_password"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("mm_name") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("mm_level") = (MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("mm_level") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
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
<form action="<%=MM_LoginAction%>" method="POST" name="form1" onSubmit="MM_validateForm('a_username','','R','a_password','','R');return document.MM_returnValue">
  <table width="100%" align="center">
    <tr> 
      <td width="30%" height="30">&nbsp;</td>
      <td height="30">&nbsp;</td>
    </tr>
    <tr> 
      <td height="30" align="right" class="bgcolor-left">用户名称:</td>
      <td height="30" class="bgcolor-right">
<input name="a_username" type="text" id="a_username"></td>
    </tr>
    <tr> 
      <td height="30" align="right" class="bgcolor-left">密码:</td>
      <td height="30" class="bgcolor-right">
<input name="a_password" type="password" id="a_password"></td>
    </tr>
    <tr>
      <td height="30" align="right">&nbsp;</td>
      <td height="30">
<input type="submit" name="Submit" value="登陆">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
