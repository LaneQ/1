<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 

<%
Dim rsvip,strvip
set rsvip=server.CreateObject("adodb.recordset")
rsvip.open "select vip from [user] where username='"&request.cookies(cookieName)("username")&"' ",conn,1,1
strvip = rsvip("vip")
rsvip.close
set rsvip=nothing

 
dim shijian,goods
dim userid,id,rs2,rs3,score
id=request("id")
userid=request("userid")
shijian=now()
goods=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
set rs2=server.CreateObject("adodb.recordset")
rs2.open "select id,name,price1,price2,vipprice,discount,score from product where id in ("&id&") order by id ",conn,1,1
goods=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)

do while not rs2.eof

	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from orders",conn,1,3
	rs.addnew
	score=score+rs2("score")
	rs("username")=trim(request.cookies(cookieName)("username"))
	rs("id")=rs2("id")
	rs("actiondate")=shijian
	rs("productnum")=CInt(Request("shop"&rs2("id")))
	rs("state")=1
	rs("goods")=goods
	rs("postcode")=int(request("postcode"))
	rs("recepit")=trim(request("recepit"))
	rs("address")=trim(request("address"))
	rs("paymethord")=int(request("paymethord"))
	rs("deliverymethord")=int(request("deliverymethord"))
	rs("sex")=int(request("sex"))
	rs("comments")=HTMLEncode2(trim(request("comments")))

	if  strvip = true then 
		rs("paid")=rs2("vipprice")*CInt(Request("shop"&rs2("id")))
	else
		rs("paid")=rs2("price2")*CInt(Request("shop"&rs2("id")))
	end if

	
	rs("realname")=trim(request("realname"))
	rs("useremail")=trim(request("useremail"))
	rs("usertel")=trim(request("usertel"))
	rs("userid")=userid
	rs.update
	rs.close
	conn.execute "delete from orders where username='"&request.cookies(cookieName)("username")&"' and id in ("&id&") and state=6"
	rs2.movenext
loop

rs2.close

rs2.open "select score from [user] where userid="&userid,conn,1,3
rs2("score")=rs2("score")+int(score)
rs2.Update
rs2.close
set rs2=nothing

set rs=server.CreateObject("adodb.recordset")
rs.open "select product.id,product.name,product.price1,vipprice,product.price2,orders.sex,orders.realname,orders.recepit,orders.goods,orders.postcode,orders.comments,orders.paymethord,orders.deliverymethord,orders.paid,orders.productnum from product inner join orders on product.id=orders.id where orders.username='"&request.cookies(cookieName)("username")&"' and state=1 and goods='"&goods&"' ",conn,1,1

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
          <td style="color:#415373">�������</td>
        </tr>
      </table>      <br>      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="3">
        <tr>
          <td height="25" align="center"><FONT SIZE="3"><B>��ϲ
                  <% =request.cookies(cookieName)("username") %>
          �����ѳɹ����ύ�˴˶�������ϸ��Ϣ����</B></FONT></td>
        </tr>
        <tr>
          <td height="18">�����ţ�<font color=#FF6600><%=rs("goods")%></font></td>
        </tr>
        <tr>
          <td height="18">��Ʒ�б�</td>
        </tr>
        <tr>
          <td>
            <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" >
              <tr align="center">
                <td>��Ʒ����</td>
                <td>�г���</td>
                <td>��Ա��</td>
                <td>VIP��Ա��</td>
                <td>����</td>
                <td> С��</td>
              </tr>
              <%dim iiii 
 do while not rs.eof
%>
              <tr BGCOLOR=#FFFFFF>
                <td><%response.Write "<a href=vpro.asp?id="&rs("id")&" target=_blank>"&rs("name")&"</a>"%></td>
                <td align="center"><%=rs("price1")%>Ԫ</td>
                <td align="center"><%=rs("price2")%>Ԫ</td>
                <td align="center"><%=rs("vipprice")%>Ԫ</td>
                <td align="center"><%=rs("productnum")%></td>
                <% 
set rsvip=server.CreateObject("adodb.recordset")
rsvip.open "select vip from [user] where username='"&request.cookies(cookieName)("username")&"' ",conn,1,1
if  rsvip("vip") = true then  %>
                <td align="center"><%=rs("vipprice")*rs("productnum")%>Ԫ</td>
              </tr>
              <%
 iiii=rs("vipprice")*rs("productnum")+iiii
else %>
        <td align="center"><%=rs("price2")*rs("productnum")%>Ԫ</td>
        </tr>
        <%
iiii=rs("price2")*rs("productnum")+iiii
end if
	rs.movenext
    loop
    rs.movefirst
	rsvip.close
	 set rsvip=nothing
    %>
        <tr align="center">
          <td colspan="6"><br>            <%response.write "<font color=#FF6600>��ѡ����ͻ���ʽ�ǣ�"

		  set rs3=server.CreateObject("adodb.recordset")
		  rs3.open "select * from delivery where deliveryid="&int(rs("deliverymethord")),conn,1,1
		  if rs3.eof and rs3.bof then
		  response.write "�ͻ���ʽ�Ѿ���ɾ��"
		  response.write "&nbsp;���ӷ��ã�0Ԫ&nbsp;&nbsp;&nbsp;���ƣ�"
		  response.write iiii&"Ԫ"
		  else
		  response.Write trim(rs3("subject"))
		  response.write "&nbsp;���ӷ��ã�"&rs3("fee")&"Ԫ&nbsp;&nbsp;&nbsp;���ƣ�"
		  response.write iiii+rs3("fee")&"Ԫ"
		  end if
		  rs3.close
		  set rs3=nothing
		%></td>
        </tr>
          </table>
          <br></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>������������<font color=#FF6600><%=trim(rs("realname"))%></font></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>�ջ���������<font color=#FF6600>
            <%response.Write trim(request("recepit"))
    if request("sex")=0 then
    response.Write "&nbsp;(����)"
    else
    response.Write "&nbsp;(Ůʿ)"
    end if%>
          </font></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>�ջ���ϸ��ַ��<font color=#FF6600><%=trim(request("address"))%></font></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>�ʱࣺ<font color=#FF6600><%=trim(request("postcode"))%></font>&nbsp;&nbsp;&nbsp;&nbsp;�绰��<font color=#FF6600><%=trim(request("usertel"))%></font>&nbsp;&nbsp;&nbsp;&nbsp;�����ʼ���<font color=#FF6600><%=trim(request("useremail"))%></font></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>�ͻ���ʽ��<font color=#FF6600>
            <%
      set rs3=server.CreateObject("adodb.recordset")
      rs3.open "select * from delivery where deliveryid="&request("deliverymethord"),conn,1,1
	  if rs3.eof and rs3.bof then
	  response.write "��ʽ�Ѿ���ɾ��"
	  else
      response.Write trim(rs3("subject"))
      end if
	  rs3.close
      set rs3=nothing
      %>
            </font>&nbsp;&nbsp;&nbsp;&nbsp;֧����ʽ��<font color=#FF6600>
            <%
      set rs3=server.CreateObject("adodb.recordset")
      rs3.open "select * from delivery where deliveryid="&request("paymethord"),conn,1,1
	  if rs3.eof and rs3.bof then
	  response.write "��ʽ�Ѿ���ɾ��"
	  else
      response.Write trim(rs3("subject"))
      end if
	  rs3.close
      set rs=nothing%>
          </font></td>
        </tr>
        <%if trim(request("comments"))<>"" then%>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>�������ԣ�<%=trim(request("comments"))%></td>
        </tr>
        <%end if%>
        <tr>
          <td height="18" ><br>
          ������һ����������ѡ���֧����ʽ���л����ʱ��ע������<font color="#FF0000">������</font>��<FONT COLOR="#FF0000">Ϊ�˸���ʱ��Ϊ����񣬵������һ��Ҫ�ǵõ�����<A HREF="myorder.asp" TARGET="_self"><B>�ʺ����޸���Ķ���<font color="#000000">״̬</font></B></A></FONT></td>
        </tr>
        <tr>
          <td height="18"  style='PADDING-LEFT: 100px'>
            <div align="right"><a href="#" onClick=javascript:window.close()> </a><font color="#999999"><FONT COLOR="#000000">������� ����ʱ�䣺<%=shijian%></FONT>&nbsp;</font></div></td>
        </tr>
        <tr>
          <td height="18" align="center"  ><input type="button" name="Submit" value="�ر�" onClick=javascript:window.close()></td>
        </tr>
      </table>      <br>      <br>
      
    </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


