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
if NOT isempty(request("MoveSubmit")) then
dim sortsid,categoryid
sortsid=int(request("sortsid"))
categoryid=int(request("categoryid"))
set rs=server.CreateObject("adodb.recordset")
rs.open "select sortsid,categoryid from sorts where sortsid="&sortsid ,conn,1,3
rs("categoryid")=categoryid
rs.Update
rs.Close
set rs=nothing
call MsgBox("转移成功！","None","None")
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
      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">商品类别转移</td>
        </tr>
      </table>      <br>      <br>      <form name="form1" method="post" action="">
        <table width="56%" border="0" align="center" cellpadding="1" cellspacing="1">
        
          <tr>
            <td width="56%" align="right">请选择您要转移的小类：</td>
            <td width="44%">
              <select name="sortsid" size="1" class="smallinput" >
                <%set rs=server.CreateObject("adodb.recordset")
                rs.Open "select sortsid,sorts from sorts order by sortsid",conn,1,1
                if rs.EOF and rs.BOF then
                response.Write "<option value=0>还没有分类</option>"
                else
                do while not rs.EOF
                %>
                <option value="<%=int(rs("sortsid"))%>"><%=trim(rs("sorts"))%></option>
                <%rs.MoveNext
                loop
                rs.Close
                set rs=nothing
                end if%>
              </select>
            </td>
          </tr>
          <tr>
            <td align="right">请选择所属大类：</td>
            <td>
              <select name="categoryid" size="1" class="smallinput" >
                <%set rs=server.CreateObject("adodb.recordset")
                rs.Open "select categoryid,category from category order by categoryorder",conn,1,1
                if rs.eof and rs.bof then
                response.Write "<option value=0>还没有分类</option>"
                else
                do while not rs.eof
                %>
                <option value="<%=int(rs("categoryid"))%>"><%=trim(rs("category"))%></option>
                <%rs.movenext
                loop
                rs.close
                set rs=nothing
                end if%>
              </select></td>
          </tr>
          <tr align="center">
            <td height="30" colspan="2"><input name="MoveSubmit" type="submit" id="MoveSubmit" value="确定转移"></td>
          </tr>
        </table>
      </form>      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">操作注意事项及说明</td>
        </tr>
      </table>      <table width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
        <tr>
          <td height="16"><font color="#FF0000">·转移小类的同时也转移小类下所有的商品。<br>
      ·转移后需要修改小分类的排序。</font></td>
        </tr>
      </table>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


