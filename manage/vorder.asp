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
dim action,goods,username
goods=request.QueryString("dan")
username=request.QueryString("username")
set rs=server.CreateObject("adodb.recordset")
rs.open "select product.id,product.name,product.price1,product.price2,product.vipprice,orders.actiondate,orders.sex,orders.realname,orders.recepit,orders.goods,orders.postcode,orders.comments,orders.paymethord,orders.deliverymethord,orders.state,orders.paid,orders.useremail,orders.usertel,orders.address,orders.productnum from product inner join orders on product.id=orders.id where orders.username='"&username&"' and goods='"&goods&"' ",conn,1,1
if rs.eof and rs.bof then
	call MsgBox("�˶���������Ʒ�ѱ�����Աɾ�����޷�������ȷ���㡣","Close","None")
	response.End
end if


action=request.QueryString("action")

select case action
	case "save"
		if request("state")<>"" then
			conn.execute "update orders set state="&request("state")&" where goods='"&goods&"' "
		end if
		call MsgBox("����״̬�޸ĳɹ�","GoUrl","porder.asp")
		response.Write goods
	case "del"
		conn.execute "delete from orders where goods='"&goods&"' "
		call MsgBox("����ɾ���ɹ���","GoUrl","porder.asp")
end select


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
          <td style="color:#415373">�鿴</td>
        </tr>
      </table>      <br>      <table width="100%" align="center" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
        <tr>
          <td colspan="2" align="center">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="90%" align="center">�����ţ�<%=goods%> ��ϸ���ϣ�</td>
                <td width="10%" align="center">
                  <input type="button" name="Submit4" value="�� ӡ" onClick="javascript:window.print()">
                </td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td colspan="2">����״̬��
            <br>
            <br>            <table width="100%" align="center" border="0" cellspacing="1" cellpadding="0" >
              <form name="form1" method="post" action="vorder.asp?dan=<%=goods%>&action=save&username=<%=username%>">
                <tr >
                  <td align="right"  >
                    <%grade()%>
                    <br>
                    <br>
                    <br>
                    <input type="submit" name="Submit" value="�޸Ķ���״̬"></td>
                </tr>
              </form>
          </table></td>
        </tr>
        <tr>
          <td colspan="2" valign="top"><br>
            ��Ʒ�б�
            <br>
            <br>            <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
              <tr align="center" height="20">
                <td WIDTH="40%">��Ʒ����</td>
                <td WIDTH="12%">��������</td>
                <td WIDTH="12%">�г���</td>
                <td WIDTH="12%">��Ա�۸�</td>
                <td WIDTH="12%">VIP��Ա��</td>
                <td WIDTH="12%">���С��</td>
              </tr>
              <%dim iii
		do while not rs.eof%>
              <tr bgcolor="#FFFFFF" align="center" height="20">
                <td align="left">&nbsp;<a href=../product.asp?id=<%=rs("id")%> target=_blank><%=trim(rs("name"))%></a></td>
                <td><%=rs("productnum")%></td>
                <td><%=rs("price1")%>Ԫ</td>
                <td><%=rs("price2")%>Ԫ</td>
                <td><%=rs("vipprice")%>Ԫ</td>
                <td><%=rs("price2")*rs("productnum")%>Ԫ</td>
              </tr>
              <%iii=rs("paid")+iii
		rs.movenext
		loop
		rs.movefirst%>
              <tr>
                <td colspan="6" bgcolor="#FFFFFF" align="right" height="20">�û�ѡ����ͻ���ʽ�ǣ�
                    <%'
          dim rs2
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select * from delivery where deliveryid="&int(rs("deliverymethord")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "��ʽ�Ա�ɾ��"
		  response.write "&nbsp;���ӷ��ã�0Ԫ"
		  response.write "&nbsp;&nbsp;����ܼƣ�"& round(iii,1) &"Ԫ"
		  else
		  response.write trim(rs2("subject"))
		  response.write "&nbsp;���ӷ��ã�"&rs2("fee")&"Ԫ"
		  response.write "&nbsp;&nbsp;����ܼƣ�"&round(iii+rs2("fee"),1)&"Ԫ"
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
          <td width="13%" nowrap>������������</td>
          <td width="87%"><%=trim(rs("realname"))%></td>
        </tr>
        <tr>
          <td>�ջ���������</td>
          <td><%=trim(rs("recepit"))%></td>
        </tr>
        <tr>
          <td>�ջ���ַ��</td>
          <td><%=trim(rs("address"))%></td>
        </tr>
        <tr>
          <td>�������룺</td>
          <td><%=trim(rs("postcode"))%></td>
        </tr>
        <tr>
          <td>��ϵ�绰��</td>
          <td><%=trim(rs("usertel"))%></td>
        </tr>
        <tr>
          <td>�����ʼ���</td>
          <td><%=trim(rs("useremail"))%></td>
        </tr>
        <tr>
          <td>�ͻ���ʽ��</td>
          <td>
            <%
    '///�ͻ���ʽ
    set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from delivery where deliveryid="&int(rs("deliverymethord")),conn,1,1
    if rs2.eof and rs2.bof then
	response.write "��ʽ�ѱ�ɾ��"
	else
	response.Write trim(rs2("subject"))
    end if
	rs2.close
    set rs2=nothing%>
          </td>
        </tr>
        <tr>
          <td>֧����ʽ��</td>
          <td>
            <%
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select * from delivery where deliveryid="&int(rs("paymethord")),conn,1,1
          if rs2.eof and rs2.bof then
		  response.write "��ʽ�ѱ�ɾ��"
		  else
		  response.Write trim(rs2("subject"))
          end if
		  rs2.close
          set rs2=nothing %>
          </td>
        </tr>
        <tr>
          <td>�û����ԣ�</td>
          <td><%=trim(rs("comments"))%></td>
        </tr>
        <tr>
          <td>�µ����ڣ�</td>
          <td><%=rs("actiondate")%></td>
        </tr>
        <tr>
          <td height="32" colspan="2" align="center">
            <input type="button" name="Submit3" value="ɾ������" onClick="if(confirm('��ȷ��Ҫɾ����?')) location.href='vorder.asp?action=del&dan=<%=goods%>&username=<%=username%>';else return;">
&nbsp;&nbsp;
      <input type="button" name="Submit2" value="�رմ���" onClick=javascript:window.close()>
          </td>
        </tr>
      </table>      <%sub grade()
select case rs("state")
case "1"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="state" type="checkbox" id="checkbox" value="checkbox" DISABLED>
      �û��ѻ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox" value="dddd" DISABLED>
      ���������յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
      �������ѷ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      �û����յ���
      <%case "2"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="checkbox" type="checkbox" id="checkbox" value="adf" checked DISABLED>
      �û��ѻ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="state" value="3" >
      ���������յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox" value="checkbox" DISABLED>
      �������ѷ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      �û����յ���
      <%case "3"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="checkbox" type="checkbox" id="checkbox" value="checkbox" checked DISABLED>
      �û��ѻ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      ���������յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="state" value="4" >
      �������ѷ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      �û����յ���
      <%case "4"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="checkbox" type="checkbox" id="checkbox" value="2" checked DISABLED>
      �û��ѻ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      ���������յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
      �������ѷ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox" value="checkbox" DISABLED>
      �û����յ���
      <%case "5"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="checkbox" type="checkbox" id="checkbox" value="2" checked DISABLED>
      �û��ѻ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      ���������յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
      �������ѷ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
      �û����յ���
      <%end select
end sub%>      <br>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


