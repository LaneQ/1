<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%

dim action,goods,username
username=trim(request.cookies(cookieName)("username"))
if NOT isempty(request.QueryString("action")) then
	goods=request.QueryString("dan")
	action=request.QueryString("action")
	select case action
		case "save"
			if request("state")<>"" then
				set rs=server.CreateObject("adodb.recordset")
				rs.Open "select state from orders where goods='"&goods&"'",conn,1,3
					do while not rs.EOF
						rs("state")=request("state")
						rs.Update
						rs.MoveNext
					loop
			
				rs.Close
			
				if request("state")=5 then
					'response.write "select productnum,id from orders where state=5 and username='"&username&"'"
					'response.end
					rs.open "select productnum,id from orders where state=5 and username='"&username&"'",conn,1,1
					dim rsSeled
					set rsSeled=server.CreateObject("adodb.recordset")
						do while not rs.eof
							rsSeled.open "select solded from product where id="&rs("id"),conn,1,3
							rsSeled("solded")=rsSeled("solded")+rs("productnum")
							rsSeled.Update
							rsSeled.close
							rs.movenext
						loop
					set rsSeled=nothing
					rs.close
				end if
				'rs.close
				set rs=nothing
			
			end if
			call Msgbox("订单状态修改成功！","GoUrl","myorder.asp")
			response.End
		
		case "del"
			set rs=server.CreateObject("adodb.recordset")
			rs.open "select username,goods from orders where goods='"&goods&"' " ,conn,1,1

			if request.cookies(cookieName)("username")<>trim(rs("username")) then
				call Msgbox("response.Write ","Back","None")
				response.End
			end if
			conn.execute "delete from orders where goods='"&goods&"' "
			Call MsgBox("订单删除成功！","GoUrl","myorder.asp")
			response.end
	end select



end if
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
          <td style="color:#415373">订单详细资料</td>
        </tr>
      </table>      <%
goods=request.QueryString("dan")
set rs=server.CreateObject("adodb.recordset")
rs.open "select product.id,product.name,product.price1,product.price2,product.vipprice,orders.actiondate,orders.sex,orders.realname,orders.recepit,orders.goods,orders.postcode,orders.comments,orders.paymethord,orders.deliverymethord,orders.state,orders.paid,orders.useremail,orders.usertel,orders.address,orders.productnum from product inner join orders on product.id=orders.id where orders.username='"&request.cookies(cookieName)("username")&"' and goods='"&goods&"' ",conn,1,1
if rs.eof and rs.bof then
response.write "<center>此订单中有商品已被管理员删除，无法进行正确计算。<br>订单取消，请通知管理员或重新下订单！</center>"
response.End
end if
%>      <br>      <table width="98%" border="0" cellspacing="0" cellpadding="1"  align="center">
        <tr >
          <td colspan="2">
            <div align="center">订单号为：<%=goods%> ，详细资料如下：</div></td>
        </tr>
        <tr >
          <td colspan="2" valign="top">订单状态： <br>
            <br><form name="form1" method="post" action="chkorder.asp?dan=<%=goods%>&action=save">
            <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
            
              <tr>
                <td><%  grade() %></td>
              </tr>
              <tr>
                <td align="right">
                  <input type="submit" name="Submit" value="修改订单状态">
                </td>
              </tr>
            
          </table></form></td>
        </tr>
        <tr >
          <td colspan="2" valign="top">商品列表：<br>
            <br>            <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" >
            <tr  align="center">
              <td>商品名称</td>
              <td width="12%">订购数量</td>
              <td width="12%">市场价</td>
              <td width="12%">会员价</td>
              <td width="16%">VIP会员价</td>
              <td width="16%">金额小计</td>
            </tr>
            <%dim iii
		do while not rs.eof%>
            <tr>
              <td style='PADDING-LEFT: 5px'><a href=product.asp?id=<%=rs("id")%> target=_blank><%=trim(rs("name"))%></a></td>
              <td align="center"><%=rs("productnum")%></td>
              <td align="center"><%=rs("price1")%>元</td>
              <td align="center"><%=rs("price2")%>元</td>
              <td align="center"><%=rs("vipprice")%>元</td>
              <td align="center"><%=rs("paid")*rs("productnum")%>元</td>
            </tr>
            <%iii=rs("paid")+iii
		rs.movenext
		loop
		rs.movefirst%>
            <tr bgcolor= #FFFFFF>
              <td colspan="6" height="19" align="center"><br>
                您选择的送货方式是：
                  <%dim rs2
              set rs2=server.CreateObject("adodb.recordset")
              rs2.Open "select * from delivery where deliveryid="&rs("deliverymethord"),conn,1,1
              if rs2.EOF and rs2.BOF then
              response.Write "方式已经被删除"
              response.write "&nbsp;附加费用：0元"
		  response.write "&nbsp;&nbsp;金额总计："&iii&" 元"
              else
              response.Write trim(rs2("subject"))
              response.write "&nbsp;附加费用："&rs2("fee")&"元"
		  response.write "&nbsp;&nbsp;金额总计：<font color=red>"&iii+rs2("fee")&"</font>&nbsp;元"
		  end if
		  rs2.Close
		  set rs2=nothing
		  %>
                  <br>
                  <br>
</td>
            </tr>
          </table></td>
        </tr>
        <tr >
          <td width="17%" style='PADDING-LEFT: 10px'>订货人姓名：</td>
          <td width="83%" style='PADDING-LEFT: 10px'><%=trim(rs("realname"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>收货人姓名：</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("recepit"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>收货地址：</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("address"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>邮　　编：</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("postcode"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>联系电话：</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("usertel"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>电子邮件：</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("useremail"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>送货方式：</td>
          <td style='PADDING-LEFT: 10px'>
            <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from delivery where deliveryid="&rs("deliverymethord"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.Close
    set rs2=nothing%>
          </td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>支付方式：</td>
          <td style='PADDING-LEFT: 10px'>
            <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from delivery where deliveryid="&rs("paymethord"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.close
    set rs2=nothing%>
          </td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>您的留言：</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("comments"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>下单日期：</td>
          <td style='PADDING-LEFT: 10px'><%=rs("actiondate")%></td>
        </tr>
        <tr >
          <td height="32" colspan="2">            <div align="center">
              <%if rs("state")=1 then%>
              <input type="button" name="Submit3" value="删除订单" onClick="location.href='chkorder.asp?action=del&dan=<%=goods%>'">
              <%end if%>
              <input type="button" name="Submit2" value="关闭窗口" onclick="window.close()">
              </div></td>
        </tr>
      </table>      <%sub grade()
select case rs("state")
case "1"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="state" type="checkbox" id="state" value="2">
      用户已经划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
      服务商已经收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
      服务商已经发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      用户已经收到货
      <%case "2"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="checkbox" type="checkbox" id="state" value="2" checked DISABLED>
      用户已经划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
      服务商已经收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
      服务商已经发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      用户已经收到货
      <%case "3"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="checkbox" type="checkbox" id="state" value="2" checked DISABLED>
      用户已经划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      服务商已经收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
      服务商已经发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      用户已经收到货
      <%case "4"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="checkbox" type="checkbox" id="state" value="2" checked DISABLED>
      用户已经划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      服务商已经收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
      服务商已经发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="state" value="5" >
      用户已经收到货
      <%case "5"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      未作任何处理<span style='font-family:Wingdings;'>à</span>      <input name="checkbox" type="checkbox" id="state" value="2" checked DISABLED>
      用户已经划出款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      服务商已经收到款<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
      服务商已经发货<span style='font-family:Wingdings;'>à</span>      <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
      用户已经收到货
      <%end select
end sub%>      <br>      <br>
      </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


