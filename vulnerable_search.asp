<%@ Language="VBScript" %>
<%
Option Explicit

Dim strSQL
Dim Grs

strSQL = "SELECT a.* FROM Articles a WHERE 1 = 1 "
strSQL = strSQL & "AND a.Title like '%" + Request("SearchString") + "%' "

Set Grs = DBHelper.ExecSQLReturnRS(strSQL, Nothing, Nothing)
%>
