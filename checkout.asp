<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select recepit,userid,sex,useremail,city,address,postcode,usertel,paymethord,deliverymethord,realname from [user] where username='"&request.cookies(cookieName)("username")&"'",conn,1,1
dim userid,id
id=request("id")
userid=rs("userid")
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
    <td width="219" align="left" valign="top"><!--#include file="uleft.asp"-->      <br></td><td width="561" align="left" valign="top">      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">����</td>
        </tr>
      </table>      <br>      <form action="vorder.asp" method="post" name="receiveaddr" id="receiveaddr">
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" >
          <tr >
		  <%
		  dim rs2
 set rs2=server.CreateObject("adodb.recordset")

rs2.open "select id from product where id in ("&id&") order by id",conn,1,1
do while not rs2.eof
%> <input name="<%="shop"&rs2("id")%>" type="hidden" value="<%=cint(request("shop"&rs2("id")))%>"> 
<%
rs2.movenext
loop
rs2.close
set rs2=nothing%> 
            <td width="114" style='PADDING-LEFT: 20px'>�ջ���������</td>
            <td width="436" height="28" style='PADDING-LEFT: 20px'> 
              <input name="recepit" type="text" id="recepit" size="12" value=<%=trim(rs("recepit"))%>>
&nbsp;&nbsp;�� ��
      <select name="sex" id="sex">
        <%select case rs("sex")
		  case ""
		  response.write "<option value=0 selected>��</option><option value=1>Ů</option>"
		  case "0"
		  response.write "<option value=0 selected>��</option><option value=1>Ů</option>"
		  case "1"
		  response.write "<option value=0 >��</option><option value=1 selected>Ů</option>"
		  end select%>
      </select>
      <input type=hidden name=realname value=<%=trim(rs("realname"))%>> <input name=userid type=hidden id="userid" value=<%=userid%>></td>
          </tr>
          <tr >
            <td style='PADDING-LEFT: 20px'>�ջ���ʡ/�У�</td>
            <td height="28" style='PADDING-LEFT: 20px'> <b>
              <input name="city" type="text" id="city" value=<%=trim(rs("city"))%>>
            </b></td>
          </tr>
          <tr >
            <td style='PADDING-LEFT: 20px'>��ϸ��ַ��</td>
            <td height="28" style='PADDING-LEFT: 20px'> <b>
              <input name="address" type="text" id="address" size="40" value=<%=trim(rs("address"))%>>
            </b></td>
          </tr>
          <tr >
            <td style='PADDING-LEFT: 20px'>�ʱࣺ</td>
            <td height="28" style='PADDING-LEFT: 20px'> 
            <input name="postcode" type="text" id="postcode" size="10" value=<%=rs("postcode")%>>            </td>
          </tr>
          <tr >
            <td style='PADDING-LEFT: 20px'>�绰��</td>
            <td height="28" style='PADDING-LEFT: 20px'>
            <input name="usertel" type="text" id="usertel" size="12" value=<%=trim(rs("usertel"))%>>            </td>
          </tr>
          <tr >
            <td style='PADDING-LEFT: 20px'>�����ʼ���</td>
            <td height="28" style='PADDING-LEFT: 20px'> 
            <input name="useremail" type="text" id="useremail" value=<%=trim(rs("useremail"))%>>            </td>
          </tr>
          <tr >
            <td height="32" style='PADDING-LEFT: 20px'>�ͻ���ʽ��</td>
            <td height="28" style='PADDING-LEFT: 20px'> <b>
              <%dim rs3
          set rs3=server.CreateObject("adodb.recordset")
          rs3.Open "select * from delivery where methord=0 order by deliveryidorder",conn,1,1
          response.Write "<select name=deliverymethord size="&rs3.recordcount&" id=deliverymethord>"
          do while not rs3.EOF
          response.Write "<option value="&rs3("deliveryidorder")
          if int(rs("deliverymethord"))=int(rs3("deliveryidorder")) then 
          response.Write " selected>"
          else
          response.Write ">"
          end if
          response.Write trim(rs3("subject"))&"</option>"
          rs3.MoveNext
          loop
          response.Write "</select>"
          rs3.Close
          set rs3=nothing
         %>
              <font color=red>�ͻ����������ڱ���</font></b></td>
          </tr>
          <tr >
            <td height="32" style='PADDING-LEFT: 20px'>֧����ʽ��</td>
            <td height="28" style='PADDING-LEFT: 20px'> 
            <%

          set rs3=server.CreateObject("adodb.recordset")
          rs3.open "select * from delivery where methord=1 order by deliveryidorder",conn,1,1
          response.Write "<select name=paymethord size="&rs3.recordcount&" id=paymethord>"
          do while not rs3.eof
          response.Write "<option value="&rs3("deliveryidorder")
          if int(rs("paymethord"))=int(rs3("deliveryidorder")) then
          response.Write " selected>"
          else
          response.Write ">"
          end if
          response.Write trim(rs3("subject"))&"</option>"
          rs3.movenext
          loop
          response.Write "</select>"
          rs3.close
          set rs3=nothing
	  rs.close
	  set rs=nothing%>            </td>
          </tr>
          <tr >
            <td height="32" valign="top" style='PADDING-LEFT: 20px'>�����ԣ�</td>
            <td height="28" style='PADDING-LEFT: 20px'> 
            <textarea name="comments" cols="40" rows="5" id="comments"></textarea>            </td>
          </tr>
          <tr align="center" >
            <td height="32" colspan="2" style='PADDING-LEFT: 20px'> <b>
              <input name="Submit" type="submit" id="Submit" style="height:20; font:9pt; BORDER-BOTTOM: #cccccc 1px groove; BORDER-RIGHT: #cccccc 1px groove; BACKGROUND-COLOR: #eeeeee" onClick="return ssother();"value="�ύ����">
              <input name="id" type="hidden" id="id" value="<%=id%>">
              <SCRIPT LANGUAGE="JavaScript">
//!--
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}

function ssother()
{
   if(checkspace(document.receiveaddr.recepit.value)) {
	document.receiveaddr.recepit.focus();
    alert("�Բ�������д�ջ���������");
	return false;
  }
  if(checkspace(document.receiveaddr.city.value)) {
	document.receiveaddr.city.focus();
    alert("�Բ�������д�ջ�������ʡ�У�");
	return false;
  }
  if(checkspace(document.receiveaddr.address.value)) {
	document.receiveaddr.address.focus();
    alert("�Բ�������д�ջ�����ϸ�ջ���ַ��");
	return false;
  }
  if(checkspace(document.receiveaddr.postcode.value)) {
	document.receiveaddr.postcode.focus();
    alert("�Բ�������д�ʱ࣡");
	return false;
  }
 
    if(checkspace(document.receiveaddr.usertel.value)) {
	document.receiveaddr.usertel.focus();
    alert("�Բ������������ĵ绰��");
	return false;
  }
      if(checkspace(document.receiveaddr.deliverymethord.value)) {
	document.receiveaddr.deliverymethord.focus();
    alert("�Բ�������û��ѡ���ͻ���ʽ��");
	return false;
  }
      if(checkspace(document.receiveaddr.paymethord.value)) {
	document.receiveaddr.paymethord.focus();
    alert("�Բ�������û��ѡ��֧����ʽ��");
	return false;
  }
  if(document.receiveaddr.useremail.value.length!=0)
  {
    if (document.receiveaddr.useremail.value.charAt(0)=="." ||        
         document.receiveaddr.useremail.value.charAt(0)=="@"||       
         document.receiveaddr.useremail.value.indexOf('@', 0) == -1 || 
         document.receiveaddr.useremail.value.indexOf('.', 0) == -1 || 
         document.receiveaddr.useremail.value.lastIndexOf("@")==document.receiveaddr.useremail.value.length-1 || 
         document.receiveaddr.useremail.value.lastIndexOf(".")==document.receiveaddr.useremail.value.length-1)
     {
      alert("Email��ַ��ʽ����ȷ��");
      document.receiveaddr.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email����Ϊ�գ�");
   document.receiveaddr.useremail.focus();
   return false;
   }
   
}
//-->
              </script></td>
          </tr>
        </table>
      </form>      <br>      <br>
    </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


