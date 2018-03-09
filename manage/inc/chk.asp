<%
if session("admin")="" then
	call MsgBox("ÇëÏÈµÇÂ¼£¡","GoUrl","login.asp")
	response.End
end if
%>

