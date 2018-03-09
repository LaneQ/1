<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
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
    <td width="219" align="left" valign="top"><!--#include file="uleft.asp"-->      <br></td><td width="561" align="left" valign="top"><br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">统计信息</td>
        </tr>
      </table>      <br>      <br>      <%
      dim joindtm,myorder,successed,successedsum,lastorder
set rs=server.CreateObject("adodb.recordset")
rs.open "select adddate from [user] where username='"&request.cookies(cookieName)("username")&"'",conn,1,1
joindtm=rs("adddate")
rs.close
rs.open "select distinct(goods),actiondate from orders where username='"&request.cookies(cookieName)("username")&"' and state<6 ",conn,1,1
if rs.eof and rs.bof then
response.write ""
else
rs.movelast
lastorder=rs("actiondate")
myorder=rs.recordcount
end if
rs.close
rs.open "select sum(paid) as paid from orders where username='"&request.cookies(cookieName)("username")&"' and state<6 and state>3",conn,1,1 
if rs("paid")>0 then
successedsum=rs("paid")
else
successedsum=0
end if
rs.close
rs.open "select distinct(goods) from orders where username='"&request.cookies(cookieName)("username")&"' and state<6 and state>3",conn,1,1
successed=rs.recordcount
set rs=nothing 
%>      <table width=96% border=0 align=center cellpadding=2 cellspacing=1 bgcolor=#FFFFFF>
  <tr><td height=14 colspan=2 align="center"><br>以下是您在本站的一些历史记录的统计信息</td></tr>
  <tr><td height="5"></td></tr>
  <tr height=14  bgcolor=#FFFFFF>
  <td align=right>注册日期：</td><td width=56%>&nbsp;<% =joindtm %></td></tr>
  <tr height=14  bgcolor=#FFFFFF><td align=right>上次下单：</td><td>&nbsp;<% =lastorder %></td></tr>
  <tr height=14  bgcolor=#FFFFFF><td align=right>下单次数：</td><td>&nbsp;<% = myorder %>次</td></tr>
  <tr height=14  bgcolor=#FFFFFF><td align=right>成交次数：</td><td>&nbsp;<% =successed %>次</td></tr>
  <tr height=14  bgcolor=#FFFFFF><td align=right>成交金额：</td><td>&nbsp;<% =successedsum %>元</td></tr>
      </table></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


