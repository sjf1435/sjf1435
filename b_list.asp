<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn_bargain.asp" -->
<!--#include file="chk_login.asp" -->
<!--#include file="chk_level1.asp" -->
<%
Dim rsb__MMColParam
rsb__MMColParam = "1"
If (Session("mm_name") <> "") Then 
  rsb__MMColParam = Session("mm_name")
End If
%>
<%
Dim rsb
Dim rsb_numRows

Set rsb = Server.CreateObject("ADODB.Recordset")
rsb.ActiveConnection = MM_conn_bargain_STRING
rsb.Source = "SELECT * FROM t_bargain WHERE b_aname = '" + Replace(rsb__MMColParam, "'", "''") + "' ORDER BY b_id DESC"
rsb.CursorType = 0
rsb.CursorLocation = 2
rsb.LockType = 1
rsb.Open()

rsb_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 20
Repeat1__index = 0
rsb_numRows = rsb_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsb_total
Dim rsb_first
Dim rsb_last

' set the record count
rsb_total = rsb.RecordCount

' set the number of rows displayed on this page
If (rsb_numRows < 0) Then
  rsb_numRows = rsb_total
Elseif (rsb_numRows = 0) Then
  rsb_numRows = 1
End If

' set the first and last displayed record
rsb_first = 1
rsb_last  = rsb_first + rsb_numRows - 1

' if we have the correct record count, check the other stats
If (rsb_total <> -1) Then
  If (rsb_first > rsb_total) Then
    rsb_first = rsb_total
  End If
  If (rsb_last > rsb_total) Then
    rsb_last = rsb_total
  End If
  If (rsb_numRows > rsb_total) Then
    rsb_numRows = rsb_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsb_total = -1) Then

  ' count the total records by iterating through the recordset
  rsb_total=0
  While (Not rsb.EOF)
    rsb_total = rsb_total + 1
    rsb.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsb.CursorType > 0) Then
    rsb.MoveFirst
  Else
    rsb.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsb_numRows < 0 Or rsb_numRows > rsb_total) Then
    rsb_numRows = rsb_total
  End If

  ' set the first and last displayed record
  rsb_first = 1
  rsb_last = rsb_first + rsb_numRows - 1
  
  If (rsb_first > rsb_total) Then
    rsb_first = rsb_total
  End If
  If (rsb_last > rsb_total) Then
    rsb_last = rsb_total
  End If

End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rsb
MM_rsCount   = rsb_total
MM_size      = rsb_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsb_first = MM_offset + 1
rsb_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsb_first > MM_rsCount) Then
    rsb_first = MM_rsCount
  End If
  If (rsb_last > MM_rsCount) Then
    rsb_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="Config/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<h1>合同管理 </h1>
<hr size="1">
<%=(rsb_first)%> 到 <%=(rsb_last)%> (总共 <%=(rsb_total)%> 个记录) <br>
<br>
<table width="100%" border="0" align="center" bordercolor="#999999" bgcolor="#FFFFFF" class="border-all">
  <tr align="center"> 
    <td height="30"><strong>合同号</strong></td>
    <td height="30"><strong>发展商</strong></td>
    <td height="30"><strong>合同期</strong></td>
    <td height="30"><strong>日期</strong></td>
    <td height="30"><strong>是否签约</strong></td>
    <td height="30"><strong>合同总金额</strong></td>
    <td><strong>操作</strong></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rsb.EOF)) %>
  <tr align="center" bgcolor="#f7f7f7"> 
    <td height="3" colspan="7"></td>
  </tr>
  <tr align="center"> 
    <td height="30"><a href="b_updata.asp?b_id=<%=(rsb.Fields.Item("b_id").Value)%>"><%=(rsb.Fields.Item("b_num").Value)%></a></td>
    <td height="30"><%=(rsb.Fields.Item("b_company").Value)%></td>
    <td height="30"><%=(rsb.Fields.Item("b_datediff").Value)%>年</td>
    <td height="30"><%=(rsb.Fields.Item("b_date").Value)%></td>
    <td height="30"><%=(rsb.Fields.Item("b_sign").Value)%></td>
    <td height="30">RMB:<%= (rsb.Fields.Item("b_money").Value) %>元</td>
    <td> 
      <a href="b_updata1.asp?b_id=<%=(rsb.Fields.Item("b_id").Value)%>">修改</a> 
    </td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsb.MoveNext()
Wend
%>
</table>


<table border="0" width="50%" align="center">
  <tr> 
    <td width="23%" align="center"> <% If MM_offset <> 0 Then %>
      <a href="<%=MM_moveFirst%>">第一页</a> 
      <% End If ' end MM_offset <> 0 %> </td>
    <td width="31%" align="center"> <% If MM_offset <> 0 Then %>
      <a href="<%=MM_movePrev%>">前一页</a> 
      <% End If ' end MM_offset <> 0 %> </td>
    <td width="23%" align="center"> <% If Not MM_atTotal Then %>
      <a href="<%=MM_moveNext%>">下一页</a> 
      <% End If ' end Not MM_atTotal %> </td>
    <td width="23%" align="center"> <% If Not MM_atTotal Then %>
      <a href="<%=MM_moveLast%>">最后一页</a> 
      <% End If ' end Not MM_atTotal %> </td>
  </tr>
</table>
</body>
</html>
<%
rsb.Close()
Set rsb = Nothing
%>
