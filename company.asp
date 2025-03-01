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
  MM_editTable = "t_company"
  MM_editColumn = "c_id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "c_name|value|c_tel|value|c_site|value|c_code|value|c_address|value"
  MM_columnsStr = "c_name|',none,''|c_tel|',none,''|c_site|',none,''|c_code|',none,''|c_address|',none,''"

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
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table width="100%" align="center">
    <tr> 
      <td width="30%" height="30" align="right" nowrap class="bgcolor-left">公司名称:</td>
      <td height="30" class="bgcolor-right"> 
        <input type="text" name="c_name" value="<%=(rsc.Fields.Item("c_name").Value)%>" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">公司电话:</td>
      <td height="30" class="bgcolor-right"> 
        <input type="text" name="c_tel" value="<%=(rsc.Fields.Item("c_tel").Value)%>" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">公司站点:</td>
      <td height="30" class="bgcolor-right"> 
        <input type="text" name="c_site" value="<%=(rsc.Fields.Item("c_site").Value)%>" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">邮政编码:</td>
      <td height="30" class="bgcolor-right"> 
        <input type="text" name="c_code" value="<%=(rsc.Fields.Item("c_code").Value)%>" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">公司地址:</td>
      <td height="30" class="bgcolor-right"> 
        <input type="text" name="c_address" value="<%=(rsc.Fields.Item("c_address").Value)%>" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap>&nbsp;</td>
      <td height="30"> 
        <input type="submit" value="更新记录"> </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rsc.Fields.Item("c_id").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rsc.Close()
Set rsc = Nothing
%>
