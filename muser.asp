<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%


set rs=server.CreateObject("adodb.recordset")
rs.open "select book from [user] where username='"&request.cookies(cookieName)("username")&"' ",conn,1,1

if rs.eof and rs.bof then
	call MsgBox("错误：此用户记录不存在！","Back","None")
end if

dim msg
msg=rs("book")
rs.close
set rs=nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
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
          <td style="color:#415373">消息中心</td>
        </tr>
      </table>      <br>      <br>      <table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr>
          <td><%=msg%></td>
        </tr>
      </table>      <br>      <br>      <br>
      </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>





