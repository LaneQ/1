<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
if NOT isempty(request("RegSubmit")) then 
	if session("regtimes")=1 then
		call MsgBox("对不起，您刚注册过用户!","Back","None")
		response.end
	end if

	set rs=server.CreateObject("adodb.recordset")
	rs.open "select username,useremail from [user] where username='"&trim(request("username"))&"' or useremail='"&trim(request("useremail"))&"'",conn,1,1
	if not rs.eof and not rs.bof then
		call MsgBox("您输入的用户名或Email地址已存在，请返回重新输入！","Back","None")
	end if
	rs.close
	rs.open "select * from [user]",conn,1,3
	rs.addnew
	rs("username")=trim(request("username"))
	rs("password")=md5(trim(request("password")))
	rs("useremail")=trim(request("useremail"))

	rs("quesion")=trim(request("quesion"))
	rs("answer")=md5(trim(request("answer")))

	rs("realname")=trim(request("realname"))
	'身份证
	rs("identify")=trim(request("identify"))
	
	rs("mobile")=trim(request("mobile"))
	rs("userqq")=trim(request("userqq"))

	
	rs("adddate")=now()
	rs("lastvst")=now()
	rs("loginnum")=0
	rs("postcode")=0

	rs("score")=0


	rs("paymethord")=0
	rs("deliverymethord")=0
	rs.update
	rs.close
	set rs=nothing
	response.Cookies(cookieName)("username")=trim(request("username"))
	response.Cookies(cookieName).expires=date+1
	session("regtimes")=1
	session.Timeout=1

	call MsgBox("注册成功！请到用户管理中心填详细资料！","GoUrl","muser.asp")
end if
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style2 {color: #000000}
-->
</style>

</head>

<body>
<!--#include file="head.htm"-->


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="left" valign="top"> <br>      <br>      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="18"><img src="images/w.gif"></td>
          <td style="color:#415373">用户注册</td>
        </tr>
        </table>      <br>      <br>      <FORM NAME="userinfo" METHOD="post" ACTION="" >
        <TABLE BORDER="0" ALIGN="center" CELLPADDING="2" CELLSPACING="1" >
          <tr>
            <TD colspan="2" ALIGN="center"><FONT COLOR="#cb6f00">请填写用户信息</FONT></TD>
          </tr>
          
            <TR>
              <TD align="right"><FONT COLOR="#cb6f00">*用户名：</FONT></TD>
              <TD><INPUT NAME="username" TYPE="text" ID="username2" >
        用户名长度不能小于2。</TD>
            </TR>
            <TR>
              <TD><div align="right"><FONT COLOR="#cb6f00">*密码：</FONT></div></TD>
              <TD>
                <INPUT NAME="password" TYPE="password" ID="password">
        长度必须大于6个字符。</TD>
            </TR>
            <TR>
              <TD><div align="right"><FONT COLOR="#cb6f00">*确认密码：</FONT> </div></TD>
              <TD>
                <INPUT NAME="password1" TYPE="password" ID="password1">
              </TD>
            </TR>
            <TR>
              <TD><div align="right"><FONT COLOR="#cb6f00">*E-Mail：</FONT> </div></TD>
              <TD>
                <INPUT NAME="useremail" TYPE="text" ID="useremail2">
        请您务必填写正确的E-mail地址，便于我们与您联系；</TD>
            </TR>
            <TR>
              <TD><div align="right"><FONT COLOR="#cb6f00">真实姓名： </FONT></div></TD>
              <TD>
                <INPUT NAME="realname" TYPE="text" ID="realname2">
        收货人姓名。</TD>
            </TR>
            <TR>
              <TD><div align="right"><FONT COLOR="#cb6f00">身份证号码： </FONT></div></TD>
              <TD>
                <input name="identify" type="text" id="userqq3" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
              此项信息用于必要时以核实身份，一经注册，便不可更改，请如实填写！</TD>
            </TR>
            <TR>
              <TD><div align="right"><FONT COLOR="#cb6f00">移动手机： </FONT></div></TD>
              <TD>
                <input name="mobile" type="text" id="userqq4" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
              请填写正确的号码，以便有急事联系。</TD>
            </TR>
            <TR>
              <TD><div align="right"><FONT COLOR="#cb6f00"> Q Q：</FONT> </div></TD>
              <TD>
        <input name="userqq" type="text" id="userqq" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
        网上联系        </TD>
            </TR>
            <TR>
              <TD><div align="right"><FONT COLOR=#cb6f00>密码提示： </FONT></div></TD>
              <TD>
                <INPUT NAME="quesion" TYPE="text" ID="quesion2">
              </TD>
            </TR>
            <TR>
              <TD><div align="right"><FONT COLOR=#cb6f00>密码答案： </FONT></div></TD>
              <TD>
                <INPUT NAME="answer" TYPE="text" ID="answer2">
              </TD>
            </TR>
            <TR>
              <TD colspan="2" align="center">
                <INPUT  TYPE="submit" ONCLICK="return check();" NAME="RegSubmit" STYLE="height:20; font:9pt; BORDER-BOTTOM: #cccccc 1px groove; BORDER-RIGHT: #cccccc 1px groove; BACKGROUND-COLOR: #eeeeee"VALUE="提交" >
                <input type="reset" name="Submit5" STYLE="height:20; font:9pt; BORDER-BOTTOM: #cccccc 1px groove; BORDER-RIGHT: #cccccc 1px groove; BACKGROUND-COLOR: #eeeeee" value="清除">
                <script language="JavaScript">
<!--
function check()
{
   if(checkspace(document.userinfo.username.value)) {
	document.userinfo.username.focus();
    alert("用户名不能为空，请重新输入！");
	return false;
  }
    if(checkspace(document.userinfo.username.value) || document.userinfo.username.value.length < 2) {
	document.userinfo.username.focus();
    alert("用户名长度不能小于2，请重新输入！");
	return false;
  }
    if(checkspace(document.userinfo.identify.value) || document.userinfo.identify.value.length < 15) {
	document.userinfo.identify.focus();
    alert("身份证号码长度不能小于15位，请重新输入！");
	return false;
  }
    if(checkspace(document.userinfo.password.value) || document.userinfo.password.value.length < 6) {
	document.userinfo.password.focus();
    alert("密码长度不能小于6，请重新输入！");
	return false;
  }
    if(document.userinfo.password.value != document.userinfo.password1.value) {
	document.userinfo.password.focus();
	document.userinfo.password.value = '';
	document.userinfo.password1.value = '';
    alert("两次输入的密码不同，请重新输入！");
	return false;
  }

 if(document.userinfo.useremail.value.length!=0)
  {
    if (document.userinfo.useremail.value.charAt(0)=="." ||        
         document.userinfo.useremail.value.charAt(0)=="@"||       
         document.userinfo.useremail.value.indexOf('@', 0) == -1 || 
         document.userinfo.useremail.value.indexOf('.', 0) == -1 || 
         document.userinfo.useremail.value.lastIndexOf("@")==document.userinfo.useremail.value.length-1 || 
         document.userinfo.useremail.value.lastIndexOf(".")==document.userinfo.useremail.value.length-1)
     {
      alert("Email地址格式不正确！");
      document.userinfo.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email不能为空！");
   document.userinfo.useremail.focus();
   return false;
   }

}


function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
//-->
                </script> <br>
              <br>
              <br></TD>
            </TR>
          
        </TABLE>
      </FORM></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


