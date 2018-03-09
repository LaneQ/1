<%
if request.cookies(cookieName)("username")="" then
	call MsgBox("对不起，您还没有登录！","GoUrl","login.asp")
	response.end
end if
%>


