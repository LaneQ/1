<%
if request.cookies(cookieName)("username")="" then
	call MsgBox("�Բ�������û�е�¼��","GoUrl","login.asp")
	response.end
end if
%>


