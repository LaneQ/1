<!--#include file="manage/inc/config.asp"--> 
<%
response.Cookies(cookieName).Expires =  NOW() -1
response.Cookies(cookieName)("username")=""
response.Cookies(cookieName)("vip")=""
response.Write "<script language=javascript>alert('您已成功注销！');"
response.redirect "index.asp"
%>

