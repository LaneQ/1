<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
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
          <td style="color:#415373">�ҵĶ���</td>
        </tr>
      </table>      <br>      <table border="0" cellpadding="0" cellspacing="0" align="center"  height="0" width="100%">
        <tr>
          <td>
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
              <tr align="center">
                <td><font color="#FF6600"><b><font color="#000000">�� �� �� ��</font></b></font></td>
              </tr>
              <tr>
                <td align="right">
                  <div align="right"><font color="#FF6600"><b></b></font></div>                  <select name="state" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
                    <option value="myorder.asp?state=0" selected>--��ѡ���ѯ״̬--</option>
                    <option value="myorder.asp?state=0" >ȫ������״̬</option>
                    <option value="myorder.asp?state=1" >δ���κδ���</option>
                    <option value="myorder.asp?state=2" >�û��Ѿ�������</option>
                    <option value="myorder.asp?state=3" >�������Ѿ��յ���</option>
                    <option value="myorder.asp?state=4" >�������Ѿ�����</option>
                    <option value="myorder.asp?state=5" >�û��Ѿ��յ���</option>
                  </select></td>
              </tr>
            </table>
            <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2" >
              <tr align="center">
                <td >������</td>
                <td >�ϼƽ��</td>
                <td >�ջ���</td>
                <td colspan="2" >������</td>
                <td >����</td>
                <td >״̬</td>
              </tr>
              <%set rs=server.CreateObject("adodb.recordset")
  dim state
  state=request.QueryString("state")
  if state=0 or state="" then
  select case state
  case "0"
  rs.open "select distinct(goods),realname,actiondate,recepit,deliverymethord,paymethord,state from orders where username='"&request.cookies(cookieName)("username")&"' and state<6 order by actiondate desc",conn,1,1
  case ""
  rs.open "select distinct(goods),realname,actiondate,recepit,deliverymethord,paymethord,state from orders where username='"&request.cookies(cookieName)("username")&"' and state<5 order by actiondate desc",conn,1,1
  end select
  else
  rs.open "select distinct(goods),realname,actiondate,recepit,deliverymethord,paymethord,state from orders where username='"&request.cookies(cookieName)("username")&"' and state="&state&" order by actiondate",conn,1,1
  end if

  do while not rs.eof
   %>
              <tr bgcolor=#ffffff align="center">
                <td ><a href="chkorder.asp?dan=<%=trim(rs("goods"))%>"><%=trim(rs("goods"))%></a></td>
                <td>
                <%
				  dim shop,rs2
	set rs2=server.CreateObject("adodb.recordset")
	rs2.open "select * from delivery where deliveryid="&rs("deliverymethord"),conn,1,1
	set shop=server.CreateObject("adodb.recordset")
	shop.open "select sum(paid) as paid from orders where goods='"&trim(rs("goods"))&"' ",conn,1,1
	response.write "<font color=#FF6600>"&round(shop("paid")+rs2("fee"),1)&"Ԫ</font>"
	shop.close
	set shop=nothing
	rs2.close
	set rs2=nothing%></td>
                <td><%=trim(rs("recepit"))%></td>
                <td colspan="2"><%=trim(rs("realname"))%></td>
                <td align="center"><%=trim(rs("actiondate"))%></td>
                <td><%select case rs("state")
	case "1"
	response.write "δ���κδ���"
	case "2"
	response.write "�û��Ѿ�������"
	case "3"
	response.write "�������Ѿ��յ���"
	case "4"
	response.write "�������Ѿ�����"
	case "5"
	response.write "�û��Ѿ��յ���"
	end select%></td>
              </tr>
              <tr bgcolor=#ffffff align="left">
                <td colspan="7" >���ʽ��
                <%set rs2=server.CreateObject("adodb.recordset")
			
        rs2.open "select * from delivery where  methord=1 and deliveryidorder="&rs("paymethord"),conn,1,1
        response.Write trim(rs2("subject"))
        rs2.close
        set rs2=nothing%> 
                ���������ջ���ʽ��
                <%set rs2=server.CreateObject("adodb.recordset")
        rs2.open "select * from delivery where  methord=0 and deliveryidorder="&rs("deliverymethord"),conn,1,1
        response.Write trim(rs2("subject"))
        rs2.close
        set rs2=nothing
        %></td>
              </tr>
              <tr bgcolor=#ffffff align="center">
                <td colspan="7" align="right" >&nbsp;</td>
              </tr>
              <%
   rs.movenext
  loop
  rs.close
  set rs=nothing%>
            </table>
      </table>      <br>
      </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


