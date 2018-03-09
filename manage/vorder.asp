<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%
if session("rank")>1 then
	call Msgbox("你的权限不够！","Back","None")
	response.End
end if
%>

<%
dim action,goods,username
goods=request.QueryString("dan")
username=request.QueryString("username")
set rs=server.CreateObject("adodb.recordset")
rs.open "select product.id,product.name,product.price1,product.price2,product.vipprice,orders.actiondate,orders.sex,orders.realname,orders.recepit,orders.goods,orders.postcode,orders.comments,orders.paymethord,orders.deliverymethord,orders.state,orders.paid,orders.useremail,orders.usertel,orders.address,orders.productnum from product inner join orders on product.id=orders.id where orders.username='"&username&"' and goods='"&goods&"' ",conn,1,1
if rs.eof and rs.bof then
	call MsgBox("此订单中有商品已被管理员删除，无法进行正确计算。","Close","None")
	response.End
end if


action=request.QueryString("action")

select case action
	case "save"
		if request("state")<>"" then
			conn.execute "update orders set state="&request("state")&" where goods='"&goods&"' "
		end if
		call MsgBox("订单状态修改成功","GoUrl","porder.asp")
		response.Write goods
	case "del"
		conn.execute "delete from orders where goods='"&goods&"' "
		call MsgBox("订单删除成功！","GoUrl","porder.asp")
end select


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
<link href="../style.css" rel="stylesheet" type="text/css">


</head>

<body>
<!--#include file="head.htm"-->

<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="219" align="left" valign="top"><br>
      <!--#include file="menu.htm"-->

        <br></td><td width="561" align="left" valign="top">
      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">查看</td>
        </tr>
      </table>      <br>      <table width="100%" align="center" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
        <tr>
          <td colspan="2" align="center">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="90%" align="center">订单号：<%=goods%> 详细资料：</td>
                <td width="10%" align="center">
                  <input type="button" name="Submit4" value="打 印" onClick="javascript:window.print()">
                </td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td colspan="2">订单状态：
            <br>
            <br>            <table width="100%" align="center" border="0" cellspacing="1" cellpadding="0" >
              <form name="form1" method="post" action="vorder.asp?dan=<%=goods%>&action=save&username=<%=username%>">
                <tr >
                  <td align="right"  >
                    <%grade()%>
                    <br>
                    <br>
                    <br>
                    <input type="submit" name="Submit" value="修改订单状态"></td>
                </tr>
              </form>
          </table></td>
        </tr>
        <tr>
          <td colspan="2" valign="top"><br>
            商品列表：
            <br>
            <br>            <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
              <tr align="center" height="20">
                <td WIDTH="40%">商品名称</td>
                <td WIDTH="12%">订购数量</td>
                <td WIDTH="12%">市场价</td>
                <td WIDTH="12%">会员价格</td>
                <td WIDTH="12%">VIP会员价</td>
                <td WIDTH="12%">金额小计</td>
              </tr>
              <%dim iii
		do while not rs.eof%>
              <tr bgcolor="#FFFFFF" align="center" height="20">
                <td align="left">&nbsp;<a href=../product.asp?id=<%=rs("id")%> target=_blank><%=trim(rs("name"))%></a></td>
                <td><%=rs("productnum")%></td>
                <td><%=rs("price1")%>元</td>
                <td><%=rs("price2")%>元</td>
                <td><%=rs("vipprice")%>元</td>
                <td><%=rs("price2")*rs("productnum")%>元</td>
              </tr>
              <%iii=rs("paid")+iii
		rs.movenext
		loop
		rs.movefirst%>
              <tr>
                <td colspan="6" bgcolor="#FFFFFF" align="right" height="20">用户选择的送货方式是：
                    <%'
          dim rs2
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select * from delivery where deliveryid="&int(rs("deliverymethord")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "方式以被删除"
		  response.write "&nbsp;附加费用：0元"
		  response.write "&nbsp;&nbsp;金额总计："& round(iii,1) &"元"
		  else
		  response.write trim(rs2("subject"))
		  response.write "&nbsp;附加费用："&rs2("fee")&"元"
		  response.write "&nbsp;&nbsp;金额总计："&round(iii+rs2("fee"),1)&"元"
		  end if
		  rs2.Close
		  set rs2=nothing
		  %>
&nbsp;&nbsp;&nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="10" colspan="2"></td>
        </tr>
        <tr>
          <td width="13%" nowrap>订货人姓名：</td>
          <td width="87%"><%=trim(rs("realname"))%></td>
        </tr>
        <tr>
          <td>收货人姓名：</td>
          <td><%=trim(rs("recepit"))%></td>
        </tr>
        <tr>
          <td>收货地址：</td>
          <td><%=trim(rs("address"))%></td>
        </tr>
        <tr>
          <td>邮政编码：</td>
          <td><%=trim(rs("postcode"))%></td>
        </tr>
        <tr>
          <td>联系电话：</td>
          <td><%=trim(rs("usertel"))%></td>
        </tr>
        <tr>
          <td>电子邮件：</td>
          <td><%=trim(rs("useremail"))%></td>
        </tr>
        <tr>
          <td>送货方式：</td>
          <td>
            <%
    '///送货方式
    set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from delivery where deliveryid="&int(rs("deliverymethord")),conn,1,1
    if rs2.eof and rs2.bof then
	response.write "方式已被删除"
	else
	response.Write trim(rs2("subject"))
    end if
	rs2.close
    set rs2=nothing%>
          </td>
        </tr>
        <tr>
          <td>支付方式：</td>
          <td>
            <%
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select * from delivery where deliveryid="&int(rs("paymethord")),conn,1,1
          if rs2.eof and rs2.bof then
		  response.write "方式已被删除"
		  else
		  response.Write trim(rs2("subject"))
          end if
		  rs2.close
          set rs2=nothing %>
          </td>
        </tr>
        <tr>
          <td>用户留言：</td>
          <td><%=trim(rs("comments"))%></td>
        </tr>
        <tr>
          <td>下单日期：</td>
          <td><%=rs("actiondate")%></td>
        </tr>
        <tr>
          <td height="32" colspan="2" align="center">
            <input type="button" name="Submit3" value="删除订单" onClick="if(confirm('您确定要删除吗?')) location.href='vorder.asp?action=del&dan=<%=goods%>&username=<%=username%>';else return;">
&nbsp;&nbsp;
      <input type="button" name="Submit2" value="关闭窗口" onClick=javascript:window.close()>
          </td>
        </tr>
      </table>      <%sub grade()
select case rs("state")
case "1"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="state" type="checkbox" id="checkbox" value="checkbox" DISABLED>
      用户已划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox" value="dddd" DISABLED>
      服务商已收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
      服务商已发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      用户已收到货
      <%case "2"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="checkbox" type="checkbox" id="checkbox" value="adf" checked DISABLED>
      用户已划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="state" value="3" >
      服务商已收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox" value="checkbox" DISABLED>
      服务商已发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      用户已收到货
      <%case "3"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="checkbox" type="checkbox" id="checkbox" value="checkbox" checked DISABLED>
      用户已划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      服务商已收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="state" value="4" >
      服务商已发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      用户已收到货
      <%case "4"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="checkbox" type="checkbox" id="checkbox" value="2" checked DISABLED>
      用户已划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      服务商已收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
      服务商已发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox" value="checkbox" DISABLED>
      用户已收到货
      <%case "5"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="checkbox" type="checkbox" id="checkbox" value="2" checked DISABLED>
      用户已划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      服务商已收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
      服务商已发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
      用户已收到货
      <%end select
end sub%>      <br>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


