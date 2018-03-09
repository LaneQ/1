<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
dim tmp,step,usernamee
step=request("step")

'第二步传递变量
if step=2  then 
	tmp=request.querystring("question")
	usernamee=request.querystring("usernamee")
end if

'确认是否存在此用户
set rs=server.CreateObject("adodb.recordset")
if not isempty(request("Submit")) and step=1 then
	rs.open "select quesion,answer from [user] where username='"&trim(request("username"))&"' ",conn,1,1
	if rs.eof and rs.bof then
		call MsgBox("查无此用户，请返回！","Back","None")
		response.end
	else
		response.redirect "forget.asp?step=2&question="&rs("quesion")&"&usernamee="&trim(request("username"))
	end if
	rs.close
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
      <br>      <br>      <br>      <%if step=1 or step="" then%><FORM name=frmdata  method=post action="">
        
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
              <td width="108"  align=right>请输入你的用户名：</td>
              <td width="276"  align=left><input type="text" name="username">
              <input type="submit" name="Submit" value="提交">
              <input name="step" type="hidden" id="step" value="1"></td>
            </tr>
        </table>
		</form><%end if%>
		<%if step=2 then%><FORM name=frmdata  method=post action="repws.asp">
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
              <td width="108"  align=right>您的密码提问：</td>
              <td width="276"  align=left><%=tmp%></td>
            </tr>
            <tr>
              <td  align=right>您的密码答案：</td>
              <td  align=left><input type="password" name="answer">              </td>
            </tr>
            <tr>
              <td colspan="2"  align=center><input name="Submit" type="submit" id="Submit" value="提交">
              <input name="username" type="hidden" id="username" value="<%=usernamee%>">              </td>
            </tr>
          </table>
        <br>
        <br>
      </FORM><%end if%>
	</td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


