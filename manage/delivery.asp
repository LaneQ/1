<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%
if session("rank")>1 then
	call Msgbox("���Ȩ�޲�����","Back","None")
	response.End
end if
%>


<%
dim action,deliveryid
deliveryid=request.QueryString("id")
action=request.QueryString("action")
set rs=server.CreateObject("adodb.recordset")
select case action

case "deliverysave"
rs.open "select * from delivery where deliveryid="&deliveryid,conn,1,3
rs("subject")=trim(request("subject"))
rs("deliveryidorder")=request("deliveryidorder")
rs("fee")=request("fee")
rs("methord")=0
rs.update
rs.close
call MsgBox("�ɹ��޸����ͻ���ʽ��","GoUrl","delivery.asp?action=delivery")
response.End

case "deliveryadd"
rs.open "select * from delivery",conn,1,3
rs.addnew
rs("subject")=trim(request("subject"))
rs("deliveryidorder")=request("deliveryidorder")
rs("fee")=request("fee")
rs("methord")=0
rs.update
rs.close
call MsgBox("�ɹ�������µ��ͻ���ʽ��","GoUrl","delivery.asp?action=delivery")
response.End

case "deliverydel"
conn.execute "delete from delivery where deliveryid="&deliveryid
response.redirect "delivery.asp?action=delivery"

case "zhifudel"
conn.execute "delete from delivery where deliveryid="&deliveryid
response.redirect "delivery.asp?action=zhifu"

case "zhifusave"
rs.open "select * from delivery where deliveryid="&deliveryid,conn,1,3
rs("subject")=trim(request("subject"))
rs("deliveryidorder")=request("deliveryidorder")
rs("methord")=1
rs.update
rs.close
call MsgBox("�ɹ��޸���֧����ʽ��","GoUrl","delivery.asp?action=zhifu")
response.End

case "zhifuadd"
rs.open "select * from delivery",conn,1,3
rs.addnew
rs("subject")=trim(request("subject"))
rs("deliveryidorder")=request("deliveryidorder")
rs("methord")=1
rs.update
rs.close
call MsgBox("�ɹ�������µ�֧����ʽ��","GoUrl","delivery.asp?action=zhifu")

response.End
end select
set rs=nothing
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>У԰�����</title>
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
          <td style="color:#415373">�ͻ�/�������</td>
        </tr>
      </table>      <br>      <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
        <tr align="center" height="20">
          <td><a href="?action=delivery">�޸��ͻ���ʽ</a></td>
          <td><a href="?action=zhifu">�޸�֧����ʽ</a></td>
        </tr>
        <tr>
          <td height="100" colspan="2"><br>
              <%
			  select case action
	case "delivery"%>
              <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
                <tr align="center" height="20">
                  <td width="30%">�ͻ���ʽ</td>
                  <td width="25%">���ս��</td>
                  <td width="20%">�� ��</td>
                  <td width="25%">�� �� </td>
                </tr>
                <%dim i,j
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select * from delivery where methord=0 order by deliveryidorder",conn,1,1
		i=rs.recordcount
		do while not rs.eof%>
                <tr align="center">
                  <form name="form1" method="post" action="delivery.asp?action=deliverysave&id=<%=rs("deliveryid")%>">
                    <td><input name="subject" type="text" id="subject" size="14" value=<%=trim(rs("subject"))%>></td>
                    <td><input name="fee" type="text" id="fee" size="4" value=<%=rs("fee")%> onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
                      Ԫ</td>
                    <td><input name="deliveryidorder" type="text" id="deliveryidorder" size="2" value=<%=rs("deliveryidorder")%> onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))"></td>
                    <td><input type="submit" name="Submit" value="�� ��">
&nbsp;<a href="delivery.asp?action=deliverydel&id=<%=rs("deliveryid")%>" onClick="return confirm('��ȷ������ɾ��������')"><font color="#FF0000">ɾ��</font></a> </td>
                  </form>
                </tr>
                <%rs.movenext
		loop
		rs.close
		set rs=nothing%>
                <tr align="center">
                  <form name="form2" method="post" action="delivery.asp?action=deliveryadd">
                    <td><input name="subject" type="text" id="subject" size="14"></td>
                    <td><input name="fee" type="text" id="fee" size="4" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
                    Ԫ</td>
                    <td><input name="deliveryidorder" type="text" id="deliveryidorder" value=<%=i+1%> size="2" onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))"></td>
                    <td><input type="submit" name="Submit3" value="�� ��"></td>
                  </form>
                </tr>
            </table>
              <%case "zhifu"%>
              <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
                <tr align="center" height="20">
                  <td width="40%">֧����ʽ</td>
                  <td width="30%">�� ��</td>
                  <td width="30%">����</td>
                </tr>
                <%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select * from delivery where methord=1 order by deliveryidorder",conn,1,1
		  j=rs.recordcount
		  do while not rs.eof%>
                <tr align="center">
                  <form name="form1" method="post" action="delivery.asp?action=zhifusave&id=<%=rs("deliveryid")%>">
                    <td><input name="subject" type="text" id="subject" size="14" value=<%=trim(rs("subject"))%>></td>
                    <td><input name="deliveryidorder" type="text" id="deliveryidorder" size="2" value=<%=rs("deliveryidorder")%> onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))"></td>
                    <td><input type="submit" name="Submit2" value="ȷ ��">
&nbsp;<a href="delivery.asp?action=zhifudel&id=<%=rs("deliveryid")%>" onClick="return confirm('��ȷ������ɾ��������')"><font color="#FF0000">ɾ��</font></a> </td>
                  </form>
                  <%rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
                </tr>
                <tr align="center">
                  <form name="form1" method="post" action="delivery.asp?action=zhifuadd">
                    <td><input name="subject" type="text" id="subject" size="14"></td>
                    <td><input name="deliveryidorder" type="text" id="deliveryidorder" value=<%=j+1%> size="2" onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))"></td>
                    <td><input type="submit" name="Submit32" value="�� ��"></td>
                  </form>
                </tr>
            </table>
              <%end select%>
              <br>
          </td>
        </tr>
      </table>      <br>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


