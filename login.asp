<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
if NOT isempty(request("LoginSubmit")) then
dim username,password
username=replace(trim(request("username")),"'","")
password=md5(replace(trim(request("password")),"'",""))

'if username="" or password="" then
'	call MsgBox("�Բ��𣬵�¼ʧ�ܣ��������ĵ�¼��������","None","None")
'end if

set rs=server.CreateObject("adodb.recordset")

rs.Open "select * from [user] where username='"&username&"' and password='"&password&"' " ,conn,1,3

if not(rs.bof and rs.eof) then
	if password=rs("password") then
		response.Cookies(cookieName)("username")=trim(request("username"))
		response.Cookies(cookieName)("vip")=rs("vip")
		rs("lastvst")=now()
		rs("loginnum")=rs("loginnum")+1
		rs.Update
		rs.Close
		set rs=nothing
		response.redirect "muser.asp"

	else
		call MsgBox("�Բ��������û�������������","Back","None")
	end if
else
	call MsgBox("�Բ��������û�������������","Back","None")
end if

end if


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>У԰�����</title>
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style2 {color: #000000}
-->
</style>

</head>

<body>
<!--#include file="head.htm"-->


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="219" align="left" valign="top"><!--#include file="left.asp"--></td>
    <td width="561" align="left" valign="top"> <br>      <br>        <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="18"><img src="images/w.gif"></td>
          <td style="color:#415373">�û���½</td>
        </tr>
        </table>        <br>        <form action="" method="post" name="loginfo" id="loginfo">
          <table border="0" cellpadding="0" cellspacing="10" class="t1">
            <tr>
              <td width="170" height="22" align="right" ><span class="style2">�û�����</span></td>
              <td width="235" height="22" ><input name="username" type="text" class="inputstyle" id="username">
              </td>
            </tr>
            <tr>
              <td height="22" align="right" ><span class="style2">���룺</span></td>
              <td height="22" ><input name="password" type="password" class="inputstyle" id="password"></td>
            </tr>
            <tr align="center">
              <td height="22" colspan="2" ><input type="reset" name="Submit" value="����">
                  <input name="LoginSubmit" ONCLICK="return checkuu();" type="submit" id="LoginSubmit" value="��¼">
                  <script language="JavaScript">
<!--
  function checkuu()
{
    if(checkspace(document.loginfo.username.value)) {
	document.loginfo.username.focus();
    alert("�û�������Ϊ�գ�");
	return false;
  }
    if(checkspace(document.loginfo.password.value)) {
	document.loginfo.password.focus();
    alert("���벻��Ϊ�գ�");
	return false;
  }
	
  }
//-->
                  </script></td>
            </tr>
          </table>
      </form></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


