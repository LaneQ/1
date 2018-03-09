<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
if NOT isempty(request("SaveAddrSubmit")) then
	dim username
	username=request.cookies(cookieName)("username")
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select * from [user] where username='"&username&"' ",conn,1,3
	
	rs("recepit")=trim(request("recepit"))
	rs("city")=trim(request("city"))
	rs("address")=trim(request("address"))
	rs("postcode")=cstr(request("postcode"))
	rs("usertel")=trim(request("usertel"))
'	rs("deliverymethord")=int(request("deliverymethord"))
'	rs("paymethord")=int(request("paymethord"))
'	rs("sex")=int(request("sex"))
'	rs("mobile")=int(request("mobile"))
'	rs("userqq")=int(request("userqq"))
	rs("deliverymethord")=request("deliverymethord")
	rs("paymethord")=request("paymethord")
	rs("sex")=request("sex")
	rs("mobile")=request("mobile")
	rs("userqq")=request("userqq")


	rs.Update
	rs.Close
	set rs=nothing
	call MsgBox("您的收货信息保存成功！","Back","None")
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
          <td style="color:#415373">收货资料</td>
        </tr>
      </table>      <script language="JavaScript">
	  function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}

function checkreceiveaddr()
{
   if(checkspace(document.receiveaddr.recepit.value)) {
	document.receiveaddr.recepit.focus();
    alert("对不起，请填写收货人姓名！");
	return false;
  }
  if(checkspace(document.receiveaddr.city.value)) {
	document.receiveaddr.city.focus();
    alert("对不起，请填写收货人所在省市！");
	return false;
  }
  if(checkspace(document.receiveaddr.address.value)) {
	document.receiveaddr.address.focus();
    alert("对不起，请填写收货人详细收货地址！");
	return false;
  }
  if(checkspace(document.receiveaddr.postcode.value)) {
	document.receiveaddr.postcode.focus();
    alert("对不起，请填写邮编！");
	return false;
  }
 
    if(checkspace(document.receiveaddr.usertel.value)) {
	document.receiveaddr.usertel.focus();
    alert("对不起，请留下您的电话！");
	return false;
  }
      if(checkspace(document.receiveaddr.deliverymethord.value)) {
	document.receiveaddr.deliverymethord.focus();
    alert("对不起，您还没有选择送货方式！");
	return false;
  }
      if(checkspace(document.receiveaddr.paymethord.value)) {
	document.receiveaddr.paymethord.focus();
    alert("对不起，您还没有选择支付方式！");
	return false;
  }
}
      </script>      <br>      <%
	  dim rs2

set rs=server.CreateObject("adodb.recordset")
rs.open "select recepit,recepit,city,address,postcode,usertel,mobile,userqq,deliverymethord,paymethord from [user] where username='"&request.cookies(cookieName)("username")&"' ",conn,1,1
%>      <form name=receiveaddr method=post action="">
  <table width=100% border=0 align=center cellpadding=1 cellspacing=2>
  
<tr align="center"><td colspan="2">请您仔细填写以下收货人的信息，以便所购买的商品能够及时投递。</td>
  </tr>
  <tr><td width=18% height="16" STYLE="PADDING-LEFT: 20px">收货人姓名：</td>
  <td width="82%" height="28"  STYLE="PADDING-LEFT: 20px"><input name=recepit type=text id=recepit size=12 value="<%=rs("recepit")%>"> &nbsp;&nbsp;&nbsp;性 &nbsp;别： <select name="sex" ID="Select1"><option value="0" selected>男</option><option value="1">女</option></select></td></tr>
  <tr><td height="16"  STYLE="PADDING-LEFT: 20px">收货人省/市</td><td height=28 bgcolor=#FFFFFF STYLE="PADDING-LEFT: 20px"><input name=city type=text id=city value="<%=rs("city")%>"></td></tr>
  <tr><td height="16"  STYLE="PADDING-LEFT: 20px">详细地址：</td><td height=28 bgcolor=#FFFFFF STYLE="PADDING-LEFT: 20px"><input name=address type=text id=address size=40 value="<%=rs("address")%>"></td></tr>
  <tr><td height="16"  STYLE="PADDING-LEFT: 20px">邮　　编：</td><td height=28 bgcolor=#FFFFFF STYLE="PADDING-LEFT: 20px"><input name=postcode type=text id=postcode value="<%=rs("postcode")%>"></td></tr>
  <tr><td height="17"  STYLE="PADDING-LEFT: 20px">电　　话：</td><td height=28 bgcolor=#FFFFFF STYLE="PADDING-LEFT: 20px"><input name=usertel type=text id=usertel value="<%=rs("usertel")%>"></td></tr>
  <tr><td height="17"  STYLE="PADDING-LEFT: 20px">手　　机：</td><td height=28 bgcolor=#FFFFFF STYLE="PADDING-LEFT: 20px"><input name=mobile type=text id=mobile value="<%=rs("mobile")%>"></td></tr>
  <tr><td height="17"  STYLE="PADDING-LEFT: 20px">腾讯  QQ：</td><td height=28 bgcolor=#FFFFFF STYLE="PADDING-LEFT: 20px"><input name=userqq type=text id=userqq value="<%=rs("userqq")%>"></td></tr>
  <tr>
  <td height="46"  STYLE="PADDING-LEFT: 20px">送货方式：</td>
  <td height=46  STYLE="PADDING-LEFT: 20px">
  <select name="deliverymethord" size="3" id="deliverymethord">
  <%

set rs2=server.CreateObject("adodb.recordset")
rs2.open "select * from delivery where methord=0 order by deliveryidorder",conn,1,1
do while not rs2.EOF
response.Write "<option value="&rs2("deliveryidorder")&">"&trim(rs2("subject"))&"</option>"
rs2.MoveNext
loop
rs2.Close
%>
  </select>
  </td></tr>
  <tr>
  <td height=58 bgcolor=#FFFFFF STYLE="PADDING-LEFT: 20px">支付方式：</td>
  <td height=58 bgcolor=#FFFFFF STYLE="PADDING-LEFT: 20px">
  <select name="paymethord" size="3" id="paymethord">
  <%
rs2.Open "select * from delivery where methord=1 order by deliveryidorder",conn,1,1
do while not rs2.EOF
response.Write "<option value="&rs2("deliveryidorder")&">"&trim(rs2("subject"))&"</option>"
rs2.MoveNext
loop
rs2.Close
set rs2=nothing
%>
  </select>
  </td>
  </tr>
  <tr align="center" bgcolor=#FFFFFF><td height=32 colspan=2 >
  <input name="SaveAddrSubmit" type="submit" id="SaveAddrSubmit" value="提交保存" onclick="return checkreceiveaddr();" >
  </td>
  </tr>
  </table>
        </form>      <%
rs.close
set rs=nothing
%>      <br>
      </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


