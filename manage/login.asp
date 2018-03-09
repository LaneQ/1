<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%>
<!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"--> 


<%
if NOT isempty(request("LoginSubmit")) then
	dim admin,password
	admin=replace(trim(request("Name")),"'","")
	password=md5(replace(trim(request("Pws")),"'",""))
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select * from [admin] where admin='"&admin&"' and password='"&password&"' " ,conn,1,1
	if not(rs.bof and rs.eof) then
		if password=rs("password") then
			session("admin")=trim(rs("admin"))
			session("rank")=int(rs("rank"))
			session.Timeout=sessionLife
			rs.Close
			set rs=nothing
			response.Redirect "mpro.asp"
		else
			call MsgBox("登录失败！","Back","None")

		end if
	else
		call MsgBox("非法登陆！","Back","None")	
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
<link href="../style.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.table {font-family:Verdana;font-size:10px;color:#ffffff}
.style2 {color: #66FF00}
-->
</style>

</head>

<body>
<!--#include file="head.htm"-->

<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="left" valign="top"><br>      <form name="admininfo" method="post" action="" >
        <table  width="260" border="0" align="center" cellpadding="3" cellspacing="5">
          <tr bgcolor="#FFFFFF">
            <td colspan="2"><table border="0" align="left" cellpadding="0" cellspacing="0">
              <tr>
                <td width="18"><img src="../images/w.gif" width="18" height="18"></td>
                <td width="76" style="color:#415373">管理员登录</td>
              </tr>
            </table></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td width="59" bgcolor="#FFFFFF">管理员：</td>
            <td width="174"><input name="Name" type="text" id="admin2" size="12"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td bgcolor="#FFFFFF">密&nbsp;&nbsp;码：</td>
            <td><input name="Pws" type="password" id="Pws" size="12"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td colspan="2" align="center"><input name="LoginSubmit" type="submit" id="LoginSubmit"  value="登录">

              <input  type="reset" name="Submit2" value="清除" >
            </td>
          </tr>
        </table>
    </form>      </td>
  </tr>
</table>
<table width="100%" height="20"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" bgcolor="#415373"><span class="style1">联系我们</span></td>
  </tr>
</table>
</body>
</html>


