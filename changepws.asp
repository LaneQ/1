<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 

<%
if NOT isempty(request("ChangePwsSubmit")) then
dim username
username=request.cookies(cookieName)("username")
set rs=server.CreateObject("adodb.recordset")
rs.open "select password from [user] where username='"&username&"'",conn,1,3
if md5(trim(request("password")))<>trim(rs("password")) then
	call MsgBox("�Բ����������ԭ�������","Back","None")
else
	rs("password")=md5(trim(request("password1")))
	rs.update
	rs.close
	set rs=nothing
	call MsgBox("������ĳɹ���","none","none")
end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>У԰�����</title>
<link href="style.css" rel="stylesheet" type="text/css">


</head>

<body>
<!--#include file="head.htm"-->


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="219" align="left" valign="top"><!--#include file="uleft.asp"-->      <br></td><td width="561" align="left" valign="top">
      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">�޸�����</td>
        </tr>
      </table>      <br>      <form name="form1" method="post" action="">
        <table  border="0" align="center" cellpadding="3" cellspacing="3">
          <tr>
            <td>�����룺</td>
            <td><input name="password" type="password" id="password"></td>
          </tr>
          <tr>
            <td>�����룺</td>
            <td><input name="password1" type="password" id="password1"></td>
          </tr>
          <tr>
            <td>ȷ���ϣ�</td>
            <td><input name="password2" type="password" id="password2"></td>
          </tr>
          <tr align="center">
            <td colspan="2"><input name="ChangePwsSubmit" type="submit" id="ChangePwsSubmit" value="ȷ��"></td>
          </tr>
        </table>
      </form>      <br>
      </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


