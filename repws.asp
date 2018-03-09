<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
if request("username")="" then 
	call MsgBox("非法使用！","Back","None")
	response.end
end if 

dim tmp

set rs=server.CreateObject("adodb.recordset")

'提交修改密码
if not isempty(request("SubmitRePws"))  then
	if request("password")<>request("password2") then call MsgBox("再次输入密码不一致！","Back","None")
	rs.open "select password from [user] where username='"&trim(request("username"))&"'",conn,1,3
	rs("password")=md5(trim(request("password2")))
	rs.update
	rs.close
	call MsgBox("您的密码取回成功，请登录！","GoUrl","login.asp")
	response.end
end if


rs.open "select answer from [user] where username='"&trim(request("username"))&"' ",conn,1,1
tmp=trim(rs("answer"))
rs.close

if tmp<>md5(request("answer")) then
	call Msgbox("对不起，您输入的问题答案不正确","Back","None")
	response.end
end if

	
set rs=nothing

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
<link href="style.css" rel="stylesheet" type="text/css">
</style>
<!-- European format dd-mm-yyyy -->
<script language="JavaScript" src="calendar.js"></script>

</head>

<body>
<!--#include file="head.htm"-->


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="300" align="left" valign="top">
      <br>      <br>      <br>      <FORM name=frmdata  method=post action="">
          <table border=0 align="center" cellpadding=0 cellspacing=5>
            <tbody>
              <tr align="left">
                <td colspan="2"><table cellpadding="0" cellspacing="0" border="0">
                    <tr>
                      <td><img src="images/w.gif"></td>
                      <td style="color:#415373">取回密码</td>
                    </tr>
                  </table>
                    <br></td>
              </tr>
              <tr>
                <td width="108"  align=right>请输入新密码：</td>
                <td width="276"  align=left><input type="password" name="password">
                </td>
              </tr>
              <tr>
                <td  align=right>请确认新密码：</td>
                <td  align=left><input type="password" name="password2">
                </td>
              </tr>
              <tr>
                <td colspan="2"  align=center><input name="SubmitRePws" type="submit" id="Submit" value="提交">
                  <input name="username" type="hidden" id="username" value="<%=request("username")%>">
                </td>
              </tr>
          </table>
          <br>
          <br>
        </FORM>
	</td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


