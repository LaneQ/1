<%
if session("admin")="" then
	call MsgBox("���ȵ�¼��","GoUrl","login.asp")
	response.End
end if
%>

