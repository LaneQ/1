<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%

if request.cookies(cookieName)("username")="" then
	response.Redirect "reg.asp"
	response.End
end if
set rs=server.CreateObject("adodb.recordset")

if NOT isempty(request("SaveSubmit")) then
	dim username
	username=request.cookies(cookieName)("username")
	rs.open "select useremail,realname,quesion,answer from [user] where username='"&username&"'",conn,1,3
	rs("useremail")=trim(request("useremail"))
	rs("realname")=trim(request("realname"))
	rs("quesion")=trim(request("quesion"))
	if trim(request("answer"))<>""then
		rs("answer")=md5(trim(request("answer")))
	end if
	rs.update
	rs.close
end if


rs.open "select useremail,vip,identify,quesion,realname from [user] where username='"&request.cookies(cookieName)("username")&"' ",conn,1,1
'rs.open "select useremail,vip,identify,quesion,realname from [user] where username='timesshop' ",conn,1,1
Dim Rank
Rank="��ͨ��Ա"
If rs("vip")=true then
Rank = "VIP��Ա"
End if
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
      <br>	  <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">��������</td>
        </tr>
      </table>	  <br>      <form name=userinfo method=post action="">
        <table border=0 align=center cellpadding=3 cellspacing=3>
          <tr>
            <td bgcolor=#FFFFFF >�� �� ���� <font color=#FF6600>
              <% = request.cookies(cookieName)("username") %>
            </font></td>
          </tr>
          <tr>
            <td bgcolor=#FFFFFF >��Ա���� [<b><font color=#FF6600>
              <% = Rank %>
            </font></b>]</td>
          </tr>
          <tr>
            <td bgcolor=#FFFFFF >E-Mail����
                <input name=useremail type=text id=useremail value=<% =trim(rs("useremail")) %>></td>
          </tr>
          <tr>
            <td bgcolor=#FFFFFF >��ʵ������
                <input name=realname type=text id=realname value=<% = trim(rs("realname"))%>></td>
          </tr>
          <tr>
            <td bgcolor=#FFFFFF >�������ʣ�
                <input name=quesion type=text id=quesion value=<% = trim(rs("quesion"))%>>
                (������������ʱʹ��)</td>
          </tr>
          <tr>
            <td bgcolor=#FFFFFF >����𰸣�
                <input name=answer type=text id=answer>
                (��������ʱ����֤�˴�)</td>
          </tr>
          <tr>
            <td height=32 align="center" bgcolor=#FFFFFF ><input name=SaveSubmit type=submit id="SaveSubmit" onClick='return checkuserinfo();' value=�ύ����>
              <script language="JavaScript" type="text/JavaScript">
function checkuserinfo()
{
 if(document.userinfo.useremail.value.length!=0)
  {
    if (document.userinfo.useremail.value.charAt(0)=="." ||        
         document.userinfo.useremail.value.charAt(0)=="@"||       
         document.userinfo.useremail.value.indexOf('@', 0) == -1 || 
         document.userinfo.useremail.value.indexOf('.', 0) == -1 || 
         document.userinfo.useremail.value.lastIndexOf("@")==document.userinfo.useremail.value.length-1 || 
         document.userinfo.useremail.value.lastIndexOf(".")==document.userinfo.useremail.value.length-1)
     {
      alert("Email��ַ��ʽ����ȷ��");
      document.userinfo.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email����Ϊ�գ�");
   document.userinfo.useremail.focus();
   return false;
   }
}
              
</script>
  <%
rs.close
set rs=nothing
%></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


