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
          <td style="color:#415373">订单完成</td>
        </tr>
      </table>      <br>      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="3">
        <tr>
          <td height="25" align="center"><FONT SIZE="3"><B>恭喜
                  <% =request.cookies(cookieName)("username") %>
          ，您已成功的提交了此订单！详细信息如下</B></FONT></td>
        </tr>
        <tr>
          <td height="18">订单号：<font color=#FF6600><%=rs("goods")%></font></td>
        </tr>
        <tr>
          <td height="18">商品列表：</td>
        </tr>
        <tr>
          <td>
            <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" >
              <tr align="center">
                <td>商品名称</td>
                <td>市场价</td>
                <td>会员价</td>
                <td>VIP会员价</td>
                <td>数量</td>
                <td> 小计</td>
              </tr>
              <%dim iiii 
 do while not rs.eof
%>
              <tr BGCOLOR=#FFFFFF>
                <td><%response.Write "<a href=vpro.asp?id="&rs("id")&" target=_blank>"&rs("name")&"</a>"%></td>
                <td align="center"><%=rs("price1")%>元</td>
                <td align="center"><%=rs("price2")%>元</td>
                <td align="center"><%=rs("vipprice")%>元</td>
                <td align="center"><%=rs("productnum")%></td>
                <% 
set rsvip=server.CreateObject("adodb.recordset")
rsvip.open "select vip from [user] where username='"&request.cookies(cookieName)("username")&"' ",conn,1,1
if  rsvip("vip") = true then  %>
                <td align="center"><%=rs("vipprice")*rs("productnum")%>元</td>
              </tr>
              <%
 iiii=rs("vipprice")*rs("productnum")+iiii
else %>
        <td align="center"><%=rs("price2")*rs("productnum")%>元</td>
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
          <td colspan="6"><br>            <%response.write "<font color=#FF6600>您选择的送货方式是："

		  set rs3=server.CreateObject("adodb.recordset")
		  rs3.open "select * from delivery where deliveryid="&int(rs("deliverymethord")),conn,1,1
		  if rs3.eof and rs3.bof then
		  response.write "送货方式已经被删除"
		  response.write "&nbsp;附加费用：0元&nbsp;&nbsp;&nbsp;共计："
		  response.write iiii&"元"
		  else
		  response.Write trim(rs3("subject"))
		  response.write "&nbsp;附加费用："&rs3("fee")&"元&nbsp;&nbsp;&nbsp;共计："
		  response.write iiii+rs3("fee")&"元"
		  end if
		  rs3.close
		  set rs3=nothing
		%></td>
        </tr>
          </table>
          <br></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>订货人姓名：<font color=#FF6600><%=trim(rs("realname"))%></font></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>收货人姓名：<font color=#FF6600>
            <%response.Write trim(request("recepit"))
    if request("sex")=0 then
    response.Write "&nbsp;(先生)"
    else
    response.Write "&nbsp;(女士)"
    end if%>
          </font></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>收货详细地址：<font color=#FF6600><%=trim(request("address"))%></font></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>邮编：<font color=#FF6600><%=trim(request("postcode"))%></font>&nbsp;&nbsp;&nbsp;&nbsp;电话：<font color=#FF6600><%=trim(request("usertel"))%></font>&nbsp;&nbsp;&nbsp;&nbsp;电子邮件：<font color=#FF6600><%=trim(request("useremail"))%></font></td>
        </tr>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>送货方式：<font color=#FF6600>
            <%
      set rs3=server.CreateObject("adodb.recordset")
      rs3.open "select * from delivery where deliveryid="&request("deliverymethord"),conn,1,1
	  if rs3.eof and rs3.bof then
	  response.write "方式已经被删除"
	  else
      response.Write trim(rs3("subject"))
      end if
	  rs3.close
      set rs3=nothing
      %>
            </font>&nbsp;&nbsp;&nbsp;&nbsp;支付方式：<font color=#FF6600>
            <%
      set rs3=server.CreateObject("adodb.recordset")
      rs3.open "select * from delivery where deliveryid="&request("paymethord"),conn,1,1
	  if rs3.eof and rs3.bof then
	  response.write "方式已经被删除"
	  else
      response.Write trim(rs3("subject"))
      end if
	  rs3.close
      set rs=nothing%>
          </font></td>
        </tr>
        <%if trim(request("comments"))<>"" then%>
        <tr>
          <td height="18" style='PADDING-LEFT: 100px'>您的留言：<%=trim(request("comments"))%></td>
        </tr>
        <%end if%>
        <tr>
          <td height="18" ><br>
          请您在一周内依照您选择的支付方式进行汇款，汇款时请注明您的<font color="#FF0000">订单号</font>！<FONT COLOR="#FF0000">为了更及时得为你服务，当你汇完款，一定要记得到您的<A HREF="myorder.asp" TARGET="_self"><B>帐号中修改你的定单<font color="#000000">状态</font></B></A></FONT></td>
        </tr>
        <tr>
          <td height="18"  style='PADDING-LEFT: 100px'>
            <div align="right"><a href="#" onClick=javascript:window.close()> </a><font color="#999999"><FONT COLOR="#000000">订单完成 创建时间：<%=shijian%></FONT>&nbsp;</font></div></td>
        </tr>
        <tr>
          <td height="18" align="center"  ><input type="button" name="Submit" value="关闭" onClick=javascript:window.close()></td>
        </tr>
      </table>      <br>      <br>
      
    </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


