<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 

<%
dim id,username,action
action=request.QueryString("action")
username=trim(request.cookies(cookieName)("username"))
id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
select case action
	case "del"
		conn.execute "delete from orders where actionid="&request.QueryString("actionid")
		response.redirect "car.asp"
	case "add"
		rs.open "select id,username from orders where username='"&username&"' and id="&id&" and state=6",conn,1,1
		if not rs.eof and not rs.bof then
			call MsgBox("�Բ��𣬴���Ʒ�Ѵ��������Ĺ��ﳵ�У��������ظ���ӣ�","Close","None")
			response.end
			rs.close
		else
			rs.close
			rs.open "select id,username,state,paid from orders",conn,1,3
			rs.addnew
			rs("id")=id
			rs("username")=username
			rs("state")=6
			rs("paid")=0
			rs.update
			rs.close
			set rs=nothing
			call MsgBox("��Ʒ�ɹ���ӵ���Ĺ�������","Close","None")
			response.end
		end if
end select

rs.open "select orders.actionid,orders.id,product.name,product.price1,product.price2,product.discount from product inner join orders on product.id=orders.id where orders.username='"&request.cookies(cookieName)("username")&"' and orders.state=6",conn,1,1 
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
    <td width="219" align="left" valign="top"><!--#include file="uleft.asp"-->      <br></td><td width="561" align="left" valign="top"><br>        <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="18"><img src="images/w.gif"></td>
          <td style="color:#415373">���ﳵ</td>
        </tr>
        </table>
        <br>      <form name="form1" method="post" action="cart.asp">
  <table width="96%" border=0 align=center cellpadding=2 cellspacing=2>
  <tr height=20><td width=7% align=center>ѡ ��</td>
  <td width="45%" align=center>��Ʒ����</td>
  <td width="14%" align=center>�г���</td>
  <td width="13%" align=center>��Ա��</td>
  <td width="12%" align=center>�ۿ�</td>
  <td width="9%" align=center>ɾ��</td></tr>
  <%
do while not rs.eof
%>
  <tr bgcolor=#ffffff>
  <td align="center" bgcolor=#FFFFFF><input name=id type=checkbox checked value="<%=rs("id")%>"></td>
  <td STYLE="PADDING-LEFT: 5px" align=center><a href=vpro.asp?id=<%=rs("id")%> target=_blank><%=rs("name")%></a></td>		  
  <td align=center><%=rs("price1")%>Ԫ</td>	
  <td align=center><font color="#FF6600"><%=rs("price2")%>Ԫ</font></td>
  <td align=center><%=(rs("discount")*100)%>%</td>
  <td align="center">
  <a href="car.asp?action=del&actionid=<%=rs("actionid")%>"><img src=images/trash.gif width=15 height=17 border=0></a></a></td>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
%>
  <tr align="center"><td height=36 colspan=6 bgcolor=#FFFFFF><input type=submit name=Submit  value=ȥ�¶��� onclick="location='cart_t.asp'">
  <%
if action<>"addtocart" then
%>
  <input type=button name=Submit2 value=�����ɹ� onclick=javascript:window.close()>
  <%
end if
%>
  </td>
  </tr></table>        
        </form>        <p>&nbsp;</p></td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


