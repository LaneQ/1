<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 

<%
dim iaction,iid
iaction=request.QueryString("action")
iid=request.QueryString("id")

if iaction="add" then
	if request.cookies(cookieName)("username")="" then 
		call MsgBox("��û�е�½!","None","None")
	else

		set rs=server.CreateObject("adodb.recordset")
		rs.open "select id,username from orders where username='"&trim(request.cookies(cookieName)("username"))&"' and id="&iid&" and state=6",conn,1,1
		if not rs.eof and not rs.bof then
			call MsgBox("�Բ��𣬴���Ʒ�Ѵ��������Ĺ��ﳵ�У��������ظ���ӣ�","None","None")
			rs.close
		else
			rs.close
			rs.open "select id,username,state,paid from orders",conn,1,3
			rs.addnew
			rs("id")=iid
			rs("username")=trim(request.cookies(cookieName)("username"))
			rs("state")=6
			rs("paid")=0
			rs.update
			rs.close
			call MsgBox("��Ʒ�ɹ���ӵ���Ĺ�������","None","None")
		end if
		set rs=nothing
	end if
end if


dim iCarRs,iPrice,pNum
set iCarRs=server.CreateObject("adodb.recordset")
if request.cookies(cookieName)("username")="" then 
	iPrice=0
	pNum=0
else
	iCarRs.open "select count(*) as co,sum(product.vipprice) as vipsum,sum(product.price2) as psum from product inner join orders on product.id=orders.id where orders.username='"&request.cookies(cookieName)("username")&"' and orders.state=6",conn,1,1
	if(request.cookies(cookieName)("vip")) then
		iPrice=iCarRs("vipsum")
	else
		iPrice=ICarRs("psum")
	end if
	pNum=iCarRs("co")
	if pNum=0 then iPrice=0
	iCarRs.close
end if

iCarRs.open "select top 10 orders.id,product.name from product inner join orders on product.id=orders.id where orders.username='"&request.cookies(cookieName)("username")&"' and orders.state=6",conn,1,1 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ﳵ</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="200" border="0" cellspacing="0" cellpadding="0">
<%
if request.cookies(cookieName)("username")="" then 
%>
  <tr align="center">
    <td height="47">�㻹û�е�½<br><br>
        <a href="login.asp" target="_parent">��������½</a><br><br>
		<a href="forget.asp" target="_parent">��������</a></td>
  </tr>
    <tr align="center" valign="middle">
    <td><img src="images/cart_empty.gif" width="82" height="54"></td>
  </tr>

<%
else 
if iCarRs.recordcount=0 then
%>
  <tr align="center">
    <td height="47">
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select product.price2,product.vipprice,product.price1,orders.productnum from product inner join orders on product.id=orders.id where orders.state=1 and orders.username='"&trim(request.cookies(cookieName)("username"))&"' ",conn,1,1
dim shop,username

set shop=server.CreateObject("adodb.recordset")

shop.Open "select distinct(goods) from orders where username='"&request.cookies(cookieName)("username")&"' and state=1 ",conn,1,1
if  request.cookies(cookieName)("vip") = "True" then 

	if shop.recordcount=0 then
		response.write "��ӭ"&request.cookies(cookieName)("username")&"�������Ѿ���VIP�û�<br>��Ŀǰ��û��δ������<br>����:0.00Ԫ"
	else
		dim shopjiage
		do while not rs.eof
			shopjiage=round(shopjiage+rs("vipprice")*rs("productnum"),2)
			rs.movenext
		loop
		response.write "��ӭ"&request.cookies(cookieName)("username")&"�������Ѿ���VIP�û�<br>��Ŀǰ��"&shop.recordcount&"��δ������<br>���ƣ�"&shopjiage&"Ԫ(���ʷ�)"
	end if
else
	if shop.recordcount=0 then
		response.write "��ӭ"&request.cookies(cookieName)("username")&"������������ͨ�û�<br>��Ŀǰ��û��δ������<br>����:0.00Ԫ"
	else
		do while not rs.eof
			shopjiage=round(shopjiage+rs("price2")*rs("productnum"),2)
			rs.movenext
		loop
	response.write "��ӭ"&request.cookies(cookieName)("username")&"������������ͨ�û�<br>��Ŀǰ��"&shop.recordcount&"��δ������<br>���ƣ�"&shopjiage&"Ԫ(���ʷ�)"
	end if
end if

shop.Close
set shop=nothing
rs.close
set rs=nothing
%>
	
	
	</td>
  </tr>

<%
else 
dim ci
do while not iCarRs.eof
ci=ci+1
%>
  <tr align="left">
    <td style="PADDING-LEFT: 22px;"><%=ci%>.<a href="vpro.asp?id=<%=iCarRs("id")%>" target="_blank"><%=strvalue(iCarRs("name"),22)%></a></td>
  </tr>
  <% 
			  iCarRs.movenext
			  loop
			  end if %>
  <% end if 
  iCarRs.close
  set iCarRs=nothing
				

				%>
  <tr align="center">
    <td><img src="images/lineleft.gif" width="167" height="1"></td>
  </tr>
  <tr>
    <td align="center"><br>����<%=pNum%>����Ʒ|�ϼ�<%=iPrice%>Ԫ</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
</table>
</body>
</html>


