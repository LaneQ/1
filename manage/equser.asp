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
call MsgBox("�����ɹ�!","None","None")
end if
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
          <td style="color:#415373">�û���ϸ����</td>
        </tr>
      </table>      <%dim vipuser
		userid=request.querystring("id")
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from [user] where userid="&userid ,conn,1,1
		if rs("vip") = true then
		    vipuser="VIP��Ա"
		  else
		    vipuser="��ͨ��Ա"
		  end if
		  %>      <br>      <form name="form1" method="post" action="equser.asp?id=<%=userid%>">
        <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr >
            <td width="20%">&nbsp;�û����ƣ�</td>
            <td width="80%">&nbsp;<font color=#FF0000><%=trim(rs("username"))%></font></td>
          </tr>
          <tr >
            <td>&nbsp;��¼���룺</td>
            <td>&nbsp;
                <input name="password" type="text" id="password2" size="12">
        ����������Ϊ��!</td>
          </tr>
          <tr >
            <td>&nbsp;��ʵ������</td>
            <td> &nbsp;
                <input name="realname" type="text" id="realname" size="12" value=<%=trim(rs("realname"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;���֤����</td>
            <td>&nbsp;
                <INPUT NAME="identify" TYPE="text" ID="identify" SIZE="18" VALUE=<%=trim(rs("identify"))%>></td>
          </tr>
          <tr >
            <td>&nbsp;�����ʼ���</td>
            <td> &nbsp;
                <input name="useremail" type="text" id="useremail" value=<%=trim(rs("useremail"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;�ֻ����룺</td>
            <td>&nbsp;
                <INPUT NAME="mobile" TYPE="text" ID="mobile" SIZE="15" VALUE=<%=trim(rs("mobile"))%>></td>
          </tr>
          <tr >
            <td>&nbsp;�� Ѷ Q Q��</td>
            <td>&nbsp;
                <INPUT NAME="userqq" TYPE="text" ID="userqq" SIZE="15" VALUE=<%=trim(rs("userqq"))%>></td>
          </tr>
          <tr >
            <td>&nbsp;�������ʣ�</td>
            <td> &nbsp;
                <input name="quesion" type="text" id="quesion" value=<%=trim(rs("quesion"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;����𰸣�</td>
            <td> &nbsp;
                <input name="answer" type="text" id="answer">
            </td>
          </tr>
          <tr >
            <td>&nbsp;�ջ���������</td>
            <td> &nbsp;
                <input name="recepit" type="text" id="recepit" size="12" value=<%=trim(rs("recepit"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;�ջ����Ա�</td>
            <td> &nbsp;
                <%if rs("sex")=0 then%>
                <input type="radio" name="usersex" value="1">
        ��
        <input name="usersex" type="radio" value="0" checked>
        Ů
        <%else%>
        <input type="radio" name="usersex" value="1" checked>
        ��
        <input name="usersex" type="radio" value="0" >
        Ů
        <%end if%>
            </td>
          </tr>
          <tr >
            <td>&nbsp;�ջ���ʡ/�У�</td>
            <td> &nbsp;
                <input name="city" type="text" id="city" size="12" value=<%=trim(rs("city"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;�ջ���ַ��</td>
            <td> &nbsp;
                <input name="address" type="text" id="address" size="30" value=<%=trim(rs("address"))%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;�ʱࣺ</td>
            <td>&nbsp;
                <input name="postcode" type="text" id="postcode" size="12" value=<%=rs("postcode")%>>
            </td>
          </tr>
          <tr >
            <td>&nbsp;�绰��</td>
            <td> &nbsp;
                <input name="usertel" type="text" id="usertel" size="12" value=<%=trim(rs("usertel"))%>>
            </td>
          </tr>
          <tr  height="20">
            <td >&nbsp;�ͻ���ʽ��</td>
            <td > &nbsp;
                <%
dim rs2
set rs2=server.CreateObject("adodb.recordset")
rs2.open "select subject from delivery where methord=0 and deliveryidorder="&rs("deliverymethord"),conn,1,1

if rs2.recordcount=0 then
	response.write "û�д��ͻ���ʽ��"
else
	response.write rs2("subject")
end if
rs2.close
			
%>
            </td>
          </tr>
          <tr  height="20">
            <td>&nbsp;֧����ʽ��</td>
            <td> &nbsp;
                <%
rs2.open "select subject from delivery where methord=1 and deliveryidorder="&rs("paymethord"),conn,1,1
if rs2.recordcount=0 then
	response.write "û�д�֧����ʽ��"
else
	response.write rs2("subject")
end if
set rs2=nothing

 %>
            </td>
          </tr>
          <tr >
            <td>&nbsp;�û����֣�</td>
            <td>&nbsp;
                <INPUT NAME="score" TYPE="text" ID="usertel" SIZE="12" VALUE=<%=trim(rs("score"))%>>        </td>
          </tr>
          <tr >
            <td>&nbsp;��Ա����</td>
            <td bgcolor="#FFFFFFFF">&nbsp; <input name="vip" type="radio" value="true" <%if vipuser="VIP��Ա" then response.write "checked"%>>
            VIP��Ա            <input type="radio" name="vip" value="false" <%if vipuser="��ͨ��Ա" then response.write "checked"%>>
            ��ͨ��Ա</td>
          </tr>
          <tr  height="20">
            <td>&nbsp;ע��ʱ�䣺</td>
            <td>&nbsp;<%=rs("adddate")%></td>
          </tr>
          <tr  height="20">
            <td>&nbsp;�ϴε�¼ʱ�䣺</td>
            <td>&nbsp;<%=rs("lastvst")%></td>
          </tr>
          <tr  height="20">
            <td>&nbsp;��¼������</td>
            <td>&nbsp;<%=rs("loginnum")%>��</td>
          </tr>
          <tr  height="20">
            <td>&nbsp;���¶�������</td>
            <td> 
              &nbsp;<%
			set rs2=server.CreateObject("adodb.recordset")
			rs2.open "select distinct(goods) from orders where username='"&trim(rs("username"))&"' ",conn,1,1
			response.write rs2.recordcount&"&nbsp;�ʶ���"
			rs2.close
			set rs2=nothing
			%>            </td>
          </tr>
          <tr >
            <td valign="top">&nbsp;ϵͳ�㲥</td>
            <td>&nbsp;
                <TEXTAREA NAME="book" ID="book" COLS="50" ROWS="5"><%=trim(rs("book"))%></TEXTAREA></td>
          </tr>
          <tr >
            <td height="28" colspan="2">&nbsp;���Ҵ��û������ж�����
                <select name="state" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
                  <base target=Right>
                  <option value="" selected>--ѡ���ѯ״̬--</option>
                  <option value="porder.asp?state=0&namekey=<%=trim(rs("username"))%>" >ȫ������״̬</option>
                  <option value="porder.asp?state=1&namekey=<%=trim(rs("username"))%>" >δ���κδ���</option>
                  <option value="porder.asp?state=2&namekey=<%=trim(rs("username"))%>" >�û��Ѿ�������</option>
                  <option value="porder.asp?state=3&namekey=<%=trim(rs("username"))%>" >�������Ѿ��յ���</option>
                  <option value="porder.asp?state=4&namekey=<%=trim(rs("username"))%>" >�������Ѿ�����</option>
                  <option value="porder.asp?state=5&namekey=<%=trim(rs("username"))%>" >�û��Ѿ��յ���</option>
                </select>
            </td>
          </tr>
          <%rs.close
			set rs=nothing
			conn.close
			%>
          <tr>
            <td height="28" colspan="2"  align="center"><input name="SaveEditSubmit" type="submit" id="SaveEditSubmit" value="ȷ���ύ">
&nbsp;&nbsp;&nbsp;
        <input type="button" name="Submit2" value="������һҳ" onClick='javascript:history.go(-1)'>
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


