<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%
if session("rank")>1 then
	call Msgbox("你的权限不够！","Back","None")
	response.End
end if
%>

<%
'如果提交表单就建立Recoredset对像
If NOT IsEmpty (Request.Form) then 
	set rs=server.CreateObject("adodb.recordset")
	'取得处理id号
	dim id
	id=request("Id")

end if

'添加后台用户
If NOT IsEmpty (Request("AddHuser")) then
	rs.open "select * from admin",conn,1,3
	rs.addnew
	rs("admin")=trim(request("AddName"))
	rs("password")=md5(trim(request("AddPws")))
	rs("rank")=int(request("AddRank"))
	rs.update
	rs.close
	set rs=nothing
	call MsgBox("添加成功！","GoUrl","huser.asp")
end If

'删除后台用户
If NOT IsEmpty (request("Del")) then
	'取得Id号
	conn.execute ("delete from admin where id="&id)
	call MsgBox("删除成功！","GoUrl","huser.asp")

end If

'修改后台用户资料
if NOT IsEmpty (request("Modify")) then 
	'取得Id号
	rs.Open "select * from admin where id="&id,conn,1,3
	rs("admin")=trim(request("Name"))
	if trim(request("password"))<>"" then
		rs("password")=md5(trim(request("password")))
	end if
	rs("rank")=int(request("rank"))
	rs.Update
	rs.Close
	set rs=nothing
	call MsgBox("修改成功！","GoUrl","huser.asp")

end if


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
<link href="../style.css" rel="stylesheet" type="text/css">


</head>

<body>
<!--#include file="head.htm"-->

<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="219" align="left" valign="top"><br>
      <!--#include file="menu.htm"-->

        <br></td><td width="561" align="left" valign="top">      
        <br>        <table border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="../images/w.gif" width="18" height="18"></td>
            <td style="color:#415373">后台用户管理</td>
          </tr>
        </table>        <br>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="2">
            <tr align="center" bgcolor="#FFFFFF" class="bluefont" height="20">
              <td width="23%">管理员</td>
              <td width="21%">密 码</td>
              <td width="31%">权 限</td>
              <td width="25%">操 作</td>
            </tr>
            <%set rs=server.CreateObject("adodb.recordset")
        rs.Open "select * from admin order by rank",conn,1,1
        do while not rs.EOF%>
		<form name="form2" method="post" action="">
            <tr align="center" bgcolor="#FFFFFF" height="20">
              <td><input name="Name" type="text" id="Name" value="<%=trim(rs("admin"))%>" size="12"></td>
              <td><input name="Pws" type="text" id="Pws" size="12"></td>
              <td>
                <%select case rs("rank")
                case "1"
                response.Write "<input type=radio name=rank value=1 checked>管理&nbsp;<input name=rank type=radio value=2 >添加&nbsp;<input type=radio name=rank value=3>查看"
                case "2"
                response.Write "<input type=radio name=rank value=1>管理&nbsp;<input name=rank type=radio value=2 checked>添加&nbsp;<input type=radio name=rank value=3>查看"
				case "3"
				response.Write "<input type=radio name=rank value=1>管理&nbsp;<input name=rank type=radio value=2>添加&nbsp;<input type=radio name=rank value=3  checked>查看"
                end select%>
                <input name="Id" type="hidden" id="Id" value="<%=int(rs("id"))%>">
</td>
              <td><input name="Modify" type="submit" id="Modify" value="修改">
                  <input name="Del" type="submit" id="Del" value="删除">
              </td>
            </tr>
			</form>
            <%rs.movenext
        loop
        rs.close
        set rs=nothing
        %>
          </table>
          <br>          <br>          <table border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="../images/w.gif" width="18" height="18"></td>
            <td style="color:#415373">后台用户添加</td>
          </tr>
          </table>          <form name="form1" method="post" action="">

		<br>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr align="center" bgcolor="#FFFFFF" class="bluefont" height="20">
            <td width="21%">管理员</td>
            <td width="23%">密 码</td>
            <td width="35%">权 限</td>
            <td width="21%">操 作</td>
          </tr>
          <tr align="center" bgcolor="#FFFFFF">
            <td><input name="AddName" type="text" id="AddName" size="12"></td>
            <td><input name="AddPws" type="text" id="AddPws" size="12"></td>
            <td><input type="radio" name="AddRank" value="1">
      管理
        <input name="AddRank" type="radio" value="2" checked>
      添加
      <input type="radio" name="AddRank" value="3">
      查看</td>
            <td><input name="AddHuser" type="submit" id="AddHuser" value="添加"></td>
          </tr>
        </table>
          </form>          <br>          <table width="231" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="24"><img src="../images/w.gif" width="18" height="18"></td>
      <td width="207" style="color:#415373">操作注意事项及说明</td>
    </tr>
          </table>          <br>          <table width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
    <tr>
      <td><font color="#FF0000">·后台管理用户与前台用户毫无牵连。<br>
        ·添加人员只能添加、修改、删除商品资料。<br>
        ·查看人员可以管理商品评论和用户订单。<br>
        ·管理员拥有本站所有管理权限。<br>
        ·登录密码采用MD5不可逆转方式加密，如不修改密码，请留空。 <br>
        <br>
      </font></td>
    </tr>
          </table></td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


