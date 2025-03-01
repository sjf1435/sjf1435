<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn_bargain.asp" -->
<!--#include file="chk_login.asp" -->
<!--#include file="chk_level3.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_conn_bargain_STRING
  MM_editTable = "t_admin"
  MM_editColumn = "a_id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "admin.asp"
  MM_fieldsStr  = "a_username|value|a_password|value|a_level|value"
  MM_columnsStr = "a_username|',none,''|a_password|',none,''|a_level|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsa__MMColParam
rsa__MMColParam = "1"
If (Request.QueryString("a_id") <> "") Then 
  rsa__MMColParam = Request.QueryString("a_id")
End If
%>
<%
Dim rsa
Dim rsa_numRows

Set rsa = Server.CreateObject("ADODB.Recordset")
rsa.ActiveConnection = MM_conn_bargain_STRING
rsa.Source = "SELECT * FROM t_admin WHERE a_id = " + Replace(rsa__MMColParam, "'", "''") + ""
rsa.CursorType = 0
rsa.CursorLocation = 2
rsa.LockType = 1
rsa.Open()

rsa_numRows = 0
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="Config/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <h1>用户修改</h1>
  <hr size="1">
  <table width="100%" align="center">
    <tr> 
      <td width="30%" height="30" align="right" nowrap class="bgcolor-left">用户名称:</td>
      <td height="30" class="bgcolor-right"> 
        <input type="text" name="a_username" value="<%=(rsa.Fields.Item("a_username").Value)%>" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">用户密码:</td>
      <td height="30" class="bgcolor-right"> 
        <input type="text" name="a_password" value="<%=(rsa.Fields.Item("a_password").Value)%>" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">等级:</td>
      <td height="30" class="bgcolor-right"> 
        <select name="a_level" class="px12">
          <option value="1" selected <%If (Not isNull((rsa.Fields.Item("a_level").Value))) Then If (1 = (rsa.Fields.Item("a_level").Value)) Then Response.Write("SELECTED") : Response.Write("")%>>录入员</option>
          <option value="2" <%If (Not isNull((rsa.Fields.Item("a_level").Value))) Then If (2 = (rsa.Fields.Item("a_level").Value)) Then Response.Write("SELECTED") : Response.Write("")%>>审核员</option>
          <option value="3" <%If (Not isNull((rsa.Fields.Item("a_level").Value))) Then If (3 = (rsa.Fields.Item("a_level").Value)) Then Response.Write("SELECTED") : Response.Write("")%>>超级管理员</option>
        </select> </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap>　</td>
      <td height="30"> 
        <input type="submit" value="更新记录"> </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rsa.Fields.Item("a_id").Value %>">
</form>
<p>　</p>
</body>
</html>
<%
rsa.Close()
Set rsa = Nothing
%>