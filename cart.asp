<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%
dim id,i,iii ,books,bookscount,product
id=request("id")
if id="" then
	call MsgBox("��Ĺ�������û����Ʒ��","Back","None")
	response.end
end if
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
    <td width="219" align="left" valign="top"><!--#include file="left.asp"--></td>
    <td width="561" align="left" valign="top">
      <br>      <br>      <form name="form1" method="post" action="">
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
          
          <tr align="center">
            <td align="left">��Ʒ����</td>
            <td width="9%">�г���</td>
            <td width="9%"> ��Ա��</td>
            <td width="9%">�ۿ�</td>
            <td width="10%">VIP��Ա��</td>
            <td width="9%"> ����</td>
            <td width="12%"> С��</td>
            <td width="10%">�޸�����</td>
          </tr>
          <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select id,name,price1,price2,vipprice,discount from product where id in ("&id&") order by id",conn,1,1
	  iii=0
	  bookscount=request.QueryString("id").count
	  do while not rs.eof
	dim quatity 
	 Quatity = CInt( Request( "shop"&rs("id")) )
	If Quatity <=0 Then Quatity = 1
	%>
          <tr align="center">
            <td style='PADDING-LEFT: 5px' align="left"><%=trim(rs("name"))%>
                <input type=hidden name=name value=<%=trim(rs("name"))%>>
            </td>
            <td ><%=rs("price1")%>Ԫ</td>
            <input type=hidden name=price2 value=<%=rs("price2")%>>
            <td><%=rs("price2")%>Ԫ</td>
            <td><%=rs("discount")*100&"%"%></td>
            <td><%=rs("vipprice")%>Ԫ</td>
            <td><input name="<%="shop"& rs("id")%>" type="text" size="3" value="<%=Quatity%>" onKeyPress= "return regInput(this,	/^[0-9]*$/,	String.fromCharCode(event.keyCode))"onpaste	= "return regInput(this,/^[0-9]*$/, window.clipboardData.getData('Text'))"ondrop= "return regInput(this,/^[0-9]*$/,event.dataTransfer.getData('Text'))">
            </td>
            <td>
              <%
Dim rsvip,strvip,strdeposit,txtvip
set rsvip=server.CreateObject("adodb.recordset")
rsvip.open "select vip from [user] where username='"&request.cookies(cookieName)("username")&"' ",conn,1,1
strvip = rsvip("vip")
if  strvip = true then 
txtvip = "VIP��Ա"
if Quatity<=1 then
	  response.write rs("vipprice")*1&"Ԫ"
	  else
	  response.write rs("vipprice")*Quatity&"Ԫ"
	  end if	  
	  iii=rs("vipprice")*Quatity+iii
else
txtvip = "��ͨ��Ա"
if Quatity<=1 then
	  response.write rs("price2")*1&"Ԫ"
	  else
	  response.write rs("price2")*Quatity&"Ԫ"
	  end if	  
	  iii=rs("price2")*Quatity+iii
	  end if

	  %></td>
            <td WIDTH="12%" align="center"><input type="submit" name="Submit" value="�޸�"  onClick="this.form.action='cart.asp?action=modify';this.form.submit()">
            </td>
          </tr>
          <%if bookscount=1 then books=rs("id")
	rs.movenext
	loop
	rs.close
	  set rs=nothing%>
          <tr height="20">
            <td colspan="4" align="center">���� <font color="#FF0000">
              <% = txtvip %>
            </font> ��Ա</td>
            <td colspan="4" align="right"><font color="#FF0000">�ܼƣ�<%=iii%>Ԫ&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
          </tr>
          <tr>
            <td height="32" colspan="8" align="center"><input type="submit" name="Submit2" value="��һ��" onClick="this.form.action='checkout.asp';this.form.submit()" >
              <input name="id" type="hidden" id="id" value="<%=id%>">
  </td>
          </tr>
          
      </table>
    </form></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


