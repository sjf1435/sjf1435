<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="3"
MM_authFailedURL="login.asp"
MM_grantAccess=false
If Session("mm_name") <> "" Then
  If (false Or CStr(Session("mm_level"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("mm_level"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
