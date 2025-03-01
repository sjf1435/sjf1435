<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("mm_name")
  Session.Contents.Remove("mm_level")
  MM_logoutRedirectPage = "login.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="Config/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" align="center">
  <tr> 
    <td height="30" align="center" class="bgcolor-left">　</td>
  </tr>
  <tr> 
    <td height="30" align="center" class="bgcolor-right">
	<a href="<%= MM_Logout %>" target="_parent">登陆</a></td>
  </tr>
  <% if cint(session("mm_level")) <> 1 then %>
   <% if cint(session("mm_level")) = 3 then %>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="company.asp" target="main">公司管理</a></td>
  </tr>
  <% end if %>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="bargain.asp" target="main">合同管理</a></td>
  </tr>
  <% else %>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="b_add.asp" target="main">合同录入</a></td>
  </tr>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="b_list.asp" target="main">合同管理</a></td>
  </tr>
  <% end if %>
  <% if cint(session("mm_level")) = 3 then %>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="b_type.asp" target="main">合同类别</a></td>
  </tr>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="admin.asp" target="main">人员管理</a></td>
  </tr>
  <% end if %>
</table>
<p>　</p>
<table width="100%" align="center">
  <tr> 
    <td height="30" align="center" class="bgcolor-left">个人信息</td>
  </tr>
  <tr> 
    <td height="100" align="center" class="bgcolor-right">
<h3><%=session("mm_name")%></h3>
      <p>欢迎你的光临<br>
        你的级别是: 
        <% 
	dim mm_level
	mm_level=session("mm_level")
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
      </p></td>
  </tr>
</table>
<p>　</p>
</body>
</html>
