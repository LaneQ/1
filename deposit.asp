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
    <td width="219" align="left" valign="top"><!--#include file="uleft.asp"-->      <br></td><td width="561" align="left" valign="top">
      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">积分查询</td>
        </tr>
      </table>      <br>      <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select score from [user] where username='"&request.cookies(cookieName)("username")&"' ",conn,1,1
response.Write "<table width=96% border=0 align=center cellpadding=1 cellspacing=1 bgcolor=#FFFFFF>"
response.Write "<form name=userinfo method=post action=saveprofile.asp?action=deposit>"
response.Write "<br><table width=96% border=0 align=center cellpadding=1 cellspacing=1>"
response.Write "<tr bgcolor=#FFFFFF><td colspan=2 STYLE='PADDING-LEFT: 20px'><font color=#FF3300>★</font> &nbsp;您目前的积分为： <font color=#FF3300>"&rs("score")&"</font> 分。</td></tr>"
response.Write "<tr><td bgcolor=#FFFFFF>当你积分达到2000分时你就可以成为本站的VIP会员，以后本站会推出积分专栏，就可以利用积分购买相应书籍！攒积分吧</td></tr>"
response.Write "</table>"

%>      <br>
      </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


