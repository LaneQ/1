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
			call Msgbox("����״̬�޸ĳɹ���","GoUrl","myorder.asp")
			response.End
		
		case "del"
			set rs=server.CreateObject("adodb.recordset")
			rs.open "select username,goods from orders where goods='"&goods&"' " ,conn,1,1

			if request.cookies(cookieName)("username")<>trim(rs("username")) then
				call Msgbox("response.Write ","Back","None")
				response.End
			end if
			conn.execute "delete from orders where goods='"&goods&"' "
			Call MsgBox("����ɾ���ɹ���","GoUrl","myorder.asp")
			response.end
	end select



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
    <td width="219" align="left" valign="top"><!--#include file="uleft.asp"-->      <br></td><td width="561" align="left" valign="top">
      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">������ϸ����</td>
        </tr>
      </table>      <%
goods=request.QueryString("dan")
set rs=server.CreateObject("adodb.recordset")
rs.open "select product.id,product.name,product.price1,product.price2,product.vipprice,orders.actiondate,orders.sex,orders.realname,orders.recepit,orders.goods,orders.postcode,orders.comments,orders.paymethord,orders.deliverymethord,orders.state,orders.paid,orders.useremail,orders.usertel,orders.address,orders.productnum from product inner join orders on product.id=orders.id where orders.username='"&request.cookies(cookieName)("username")&"' and goods='"&goods&"' ",conn,1,1
if rs.eof and rs.bof then
response.write "<center>�˶���������Ʒ�ѱ�����Աɾ�����޷�������ȷ���㡣<br>����ȡ������֪ͨ����Ա�������¶�����</center>"
response.End
end if
%>      <br>      <table width="98%" border="0" cellspacing="0" cellpadding="1"  align="center">
        <tr >
          <td colspan="2">
            <div align="center">������Ϊ��<%=goods%> ����ϸ�������£�</div></td>
        </tr>
        <tr >
          <td colspan="2" valign="top">����״̬�� <br>
            <br><form name="form1" method="post" action="chkorder.asp?dan=<%=goods%>&action=save">
            <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
            
              <tr>
                <td><%  grade() %></td>
              </tr>
              <tr>
                <td align="right">
                  <input type="submit" name="Submit" value="�޸Ķ���״̬">
                </td>
              </tr>
            
          </table></form></td>
        </tr>
        <tr >
          <td colspan="2" valign="top">��Ʒ�б�<br>
            <br>            <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" >
            <tr  align="center">
              <td>��Ʒ����</td>
              <td width="12%">��������</td>
              <td width="12%">�г���</td>
              <td width="12%">��Ա��</td>
              <td width="16%">VIP��Ա��</td>
              <td width="16%">���С��</td>
            </tr>
            <%dim iii
		do while not rs.eof%>
            <tr>
              <td style='PADDING-LEFT: 5px'><a href=product.asp?id=<%=rs("id")%> target=_blank><%=trim(rs("name"))%></a></td>
              <td align="center"><%=rs("productnum")%></td>
              <td align="center"><%=rs("price1")%>Ԫ</td>
              <td align="center"><%=rs("price2")%>Ԫ</td>
              <td align="center"><%=rs("vipprice")%>Ԫ</td>
              <td align="center"><%=rs("paid")*rs("productnum")%>Ԫ</td>
            </tr>
            <%iii=rs("paid")+iii
		rs.movenext
		loop
		rs.movefirst%>
            <tr bgcolor= #FFFFFF>
              <td colspan="6" height="19" align="center"><br>
                ��ѡ����ͻ���ʽ�ǣ�
                  <%dim rs2
              set rs2=server.CreateObject("adodb.recordset")
              rs2.Open "select * from delivery where deliveryid="&rs("deliverymethord"),conn,1,1
              if rs2.EOF and rs2.BOF then
              response.Write "��ʽ�Ѿ���ɾ��"
              response.write "&nbsp;���ӷ��ã�0Ԫ"
		  response.write "&nbsp;&nbsp;����ܼƣ�"&iii&" Ԫ"
              else
              response.Write trim(rs2("subject"))
              response.write "&nbsp;���ӷ��ã�"&rs2("fee")&"Ԫ"
		  response.write "&nbsp;&nbsp;����ܼƣ�<font color=red>"&iii+rs2("fee")&"</font>&nbsp;Ԫ"
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
          <td width="17%" style='PADDING-LEFT: 10px'>������������</td>
          <td width="83%" style='PADDING-LEFT: 10px'><%=trim(rs("realname"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>�ջ���������</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("recepit"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>�ջ���ַ��</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("address"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>�ʡ����ࣺ</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("postcode"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>��ϵ�绰��</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("usertel"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>�����ʼ���</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("useremail"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>�ͻ���ʽ��</td>
          <td style='PADDING-LEFT: 10px'>
            <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from delivery where deliveryid="&rs("deliverymethord"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.Close
    set rs2=nothing%>
          </td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>֧����ʽ��</td>
          <td style='PADDING-LEFT: 10px'>
            <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from delivery where deliveryid="&rs("paymethord"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.close
    set rs2=nothing%>
          </td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>�������ԣ�</td>
          <td style='PADDING-LEFT: 10px'><%=trim(rs("comments"))%></td>
        </tr>
        <tr >
          <td style='PADDING-LEFT: 10px'>�µ����ڣ�</td>
          <td style='PADDING-LEFT: 10px'><%=rs("actiondate")%></td>
        </tr>
        <tr >
          <td height="32" colspan="2">            <div align="center">
              <%if rs("state")=1 then%>
              <input type="button" name="Submit3" value="ɾ������" onClick="location.href='chkorder.asp?action=del&dan=<%=goods%>'">
              <%end if%>
              <input type="button" name="Submit2" value="�رմ���" onclick="window.close()">
              </div></td>
        </tr>
      </table>      <%sub grade()
select case rs("state")
case "1"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="state" type="checkbox" id="state" value="2">
      �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
      �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
      �������Ѿ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      �û��Ѿ��յ���
      <%case "2"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="checkbox" type="checkbox" id="state" value="2" checked DISABLED>
      �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
      �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
      �������Ѿ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      �û��Ѿ��յ���
      <%case "3"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="checkbox" type="checkbox" id="state" value="2" checked DISABLED>
      �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
      �������Ѿ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
      �û��Ѿ��յ���
      <%case "4"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="checkbox" type="checkbox" id="state" value="2" checked DISABLED>
      �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
      �������Ѿ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="state" value="5" >
      �û��Ѿ��յ���
      <%case "5"%>      <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
      δ���κδ���<span style='font-family:Wingdings;'>��</span>      <input name="checkbox" type="checkbox" id="state" value="2" checked DISABLED>
      �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
      �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
      �������Ѿ�����<span style='font-family:Wingdings;'>��</span>      <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
      �û��Ѿ��յ���
      <%end select
end sub%>      <br>      <br>
      </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


