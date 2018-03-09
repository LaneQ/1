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
if NOT isempty(request("SaveEditSubmit")) then
dim userid
userid=request.QueryString("id")
if userid="" then userid=request("userid")

set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from [user] where userid="&userid,conn,1,3
if trim(request("password"))<>"" then rs("password")=md5(trim(request("password")))
rs("realname")=trim(request("realname"))
rs("identify")=trim(request("identify"))
rs("mobile")=trim(request("mobile"))
rs("userqq")=trim(request("userqq"))
rs("useremail")=trim(request("useremail"))
rs("quesion")=trim(request("quesion"))
if trim(request("answer"))<>"" then rs("answer")=md5(trim(request("answer")))
rs("sex")=request("usersex")
rs("city")=trim(request("city"))
rs("address")=trim(request("address"))
rs("postcode")=trim(request("postcode"))
rs("usertel")=trim(request("usertel"))
rs("score")=trim(request("score"))

rs("book")=trim(request("book"))
rs("vip")=trim(request("vip"))
rs.Update
rs.Close
set rs=nothing
call MsgBox("操作成功!","None","None")
end if
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
          <td style="color:#415373">用户详细资料</td>
        </tr>
      </table>      <%dim vipuser
		userid=request.querystring("id")
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from [user] where userid="&userid ,conn,1,1
		if rs("vip") = true then
		    vipuser="VIP会员"
		  else
		    vipuser="普通会员"
		  end if
		  %>      <br>      <form name="form1" method="post" action="equser.asp?id=<%=userid%>">
        <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr >
            <td width="20%">&nbsp;用户名称：</td>
            <td width="80%">&nbsp;<font color=#FF0000><%=trim(rs("username"))%></font></td>
          </tr>
          <tr >
            <td>&nbsp;登录密码：</td>
            <td>&nbsp;
                <input name="password" type="text" id="password2" size="12">
        不改密码请为空!</td>
          </tr>
          <tr >
            <td>&nbsp;真实姓名：</td>
            <td> &nbsp;
                <input name="realname" type="text" id="realname" size="12" value=<%=trim(rs("realname"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;身份证号码</td>
            <td>&nbsp;
                <INPUT NAME="identify" TYPE="text" ID="identify" SIZE="18" VALUE=<%=trim(rs("identify"))%>></td>
          </tr>
          <tr >
            <td>&nbsp;电子邮件：</td>
            <td> &nbsp;
                <input name="useremail" type="text" id="useremail" value=<%=trim(rs("useremail"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;手机号码：</td>
            <td>&nbsp;
                <INPUT NAME="mobile" TYPE="text" ID="mobile" SIZE="15" VALUE=<%=trim(rs("mobile"))%>></td>
          </tr>
          <tr >
            <td>&nbsp;腾 讯 Q Q：</td>
            <td>&nbsp;
                <INPUT NAME="userqq" TYPE="text" ID="userqq" SIZE="15" VALUE=<%=trim(rs("userqq"))%>></td>
          </tr>
          <tr >
            <td>&nbsp;密码提问：</td>
            <td> &nbsp;
                <input name="quesion" type="text" id="quesion" value=<%=trim(rs("quesion"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;密码答案：</td>
            <td> &nbsp;
                <input name="answer" type="text" id="answer">
            </td>
          </tr>
          <tr >
            <td>&nbsp;收货人姓名：</td>
            <td> &nbsp;
                <input name="recepit" type="text" id="recepit" size="12" value=<%=trim(rs("recepit"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;收货人性别：</td>
            <td> &nbsp;
                <%if rs("sex")=0 then%>
                <input type="radio" name="usersex" value="1">
        男
        <input name="usersex" type="radio" value="0" checked>
        女
        <%else%>
        <input type="radio" name="usersex" value="1" checked>
        男
        <input name="usersex" type="radio" value="0" >
        女
        <%end if%>
            </td>
          </tr>
          <tr >
            <td>&nbsp;收货人省/市：</td>
            <td> &nbsp;
                <input name="city" type="text" id="city" size="12" value=<%=trim(rs("city"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;收货地址：</td>
            <td> &nbsp;
                <input name="address" type="text" id="address" size="30" value=<%=trim(rs("address"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;邮编：</td>
            <td>&nbsp;
                <input name="postcode" type="text" id="postcode" size="12" value=<%=rs("postcode")%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;电话：</td>
            <td> &nbsp;
                <input name="usertel" type="text" id="usertel" size="12" value=<%=trim(rs("usertel"))%>>
            </td>
          </tr>
          <tr  height="20">
            <td >&nbsp;送货方式：</td>
            <td > &nbsp;
                <%
dim rs2
set rs2=server.CreateObject("adodb.recordset")
rs2.open "select subject from delivery where methord=0 and deliveryidorder="&rs("deliverymethord"),conn,1,1

if rs2.recordcount=0 then
	response.write "没有此送货方式！"
else
	response.write rs2("subject")
end if
rs2.close
			
%>
            </td>
          </tr>
          <tr  height="20">
            <td>&nbsp;支付方式：</td>
            <td> &nbsp;
                <%
rs2.open "select subject from delivery where methord=1 and deliveryidorder="&rs("paymethord"),conn,1,1
if rs2.recordcount=0 then
	response.write "没有此支付方式！"
else
	response.write rs2("subject")
end if
set rs2=nothing

 %>
            </td>
          </tr>
          <tr >
            <td>&nbsp;用户积分：</td>
            <td>&nbsp;
                <INPUT NAME="score" TYPE="text" ID="usertel" SIZE="12" VALUE=<%=trim(rs("score"))%>>        </td>
          </tr>
          <tr >
            <td>&nbsp;会员级别</td>
            <td bgcolor="#FFFFFFFF">&nbsp; <input name="vip" type="radio" value="true" <%if vipuser="VIP会员" then response.write "checked"%>>
            VIP会员            <input type="radio" name="vip" value="false" <%if vipuser="普通会员" then response.write "checked"%>>
            普通会员</td>
          </tr>
          <tr  height="20">
            <td>&nbsp;注册时间：</td>
            <td>&nbsp;<%=rs("adddate")%></td>
          </tr>
          <tr  height="20">
            <td>&nbsp;上次登录时间：</td>
            <td>&nbsp;<%=rs("lastvst")%></td>
          </tr>
          <tr  height="20">
            <td>&nbsp;登录次数：</td>
            <td>&nbsp;<%=rs("loginnum")%>次</td>
          </tr>
          <tr  height="20">
            <td>&nbsp;已下订单数：</td>
            <td> 
              &nbsp;<%
			set rs2=server.CreateObject("adodb.recordset")
			rs2.open "select distinct(goods) from orders where username='"&trim(rs("username"))&"' ",conn,1,1
			response.write rs2.recordcount&"&nbsp;笔订单"
			rs2.close
			set rs2=nothing
			%>            </td>
          </tr>
          <tr >
            <td valign="top">&nbsp;系统广播</td>
            <td>&nbsp;
                <TEXTAREA NAME="book" ID="book" COLS="50" ROWS="5"><%=trim(rs("book"))%></TEXTAREA></td>
          </tr>
          <tr >
            <td height="28" colspan="2">&nbsp;查找此用户的所有定单：
                <select name="state" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
                  <base target=Right>
                  <option value="" selected>--选择查询状态--</option>
                  <option value="porder.asp?state=0&namekey=<%=trim(rs("username"))%>" >全部订单状态</option>
                  <option value="porder.asp?state=1&namekey=<%=trim(rs("username"))%>" >未作任何处理</option>
                  <option value="porder.asp?state=2&namekey=<%=trim(rs("username"))%>" >用户已经划出款</option>
                  <option value="porder.asp?state=3&namekey=<%=trim(rs("username"))%>" >服务商已经收到款</option>
                  <option value="porder.asp?state=4&namekey=<%=trim(rs("username"))%>" >服务商已经发货</option>
                  <option value="porder.asp?state=5&namekey=<%=trim(rs("username"))%>" >用户已经收到货</option>
                </select>
            </td>
          </tr>
          <%rs.close
			set rs=nothing
			conn.close
			%>
          <tr>
            <td height="28" colspan="2"  align="center"><input name="SaveEditSubmit" type="submit" id="SaveEditSubmit" value="确认提交">
&nbsp;&nbsp;&nbsp;
        <input type="button" name="Submit2" value="返回上一页" onClick='javascript:history.go(-1)'>
            </td>
          </tr>
        </table>
      </form>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


