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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_conn_bargain_STRING
  MM_editTable = "t_btype"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "bt_name|value"
  MM_columnsStr = "bt_name|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
Dim rsbt
Dim rsbt_numRows

Set rsbt = Server.CreateObject("ADODB.Recordset")
rsbt.ActiveConnection = MM_conn_bargain_STRING
rsbt.Source = "SELECT * FROM t_btype ORDER BY bt_id DESC"
rsbt.CursorType = 0
rsbt.CursorLocation = 2
rsbt.LockType = 1
rsbt.Open()

rsbt_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsbt_numRows = rsbt_numRows + Repeat1__numRows
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="Config/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>

<body>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table width="100%" align="center">
    <tr> 
      <td width="50%" height="30" align="right" nowrap class="bgcolor-left">合同类型:</td>
      <td height="30" class="bgcolor-right"> 
        <input type="text" name="bt_name" value="" size="32"> </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap>&nbsp;</td>
      <td height="30"> 
        <input type="submit" value="插入记录"> </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<table width="100%">
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsbt.EOF)) 
%><form name="form2" method="post" action="bt_updata.asp">
  <tr> 
      <td width="50%" height="30" align="center" class="bgcolor-left"><input name="bt_id" type="hidden" id="bt_id" value="<%=(rsbt.Fields.Item("bt_id").Value)%>"> 
        <input name="bt_name" type="text" id="bt_name" value="<%=(rsbt.Fields.Item("bt_name").Value)%>" size="40"></td>
      <td height="30" align="center" class="bgcolor-right"> 
        <input type="submit" name="Submit" value="修改">
        <input name="按钮" type="button" onClick="MM_goToURL('parent','b_add.asp');return document.MM_returnValue" value="删除"> </td>
  </tr></form>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsbt.MoveNext()
Wend
%>
</table>
<p>&nbsp;</p>

</body>
</html>
<%
rsbt.Close()
Set rsbt = Nothing
%>
