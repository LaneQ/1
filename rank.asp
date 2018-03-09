<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
dim id,action
action=request.QueryString("action")
id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
rs.open "select id,name,rank,ranknum from product where id="&id,conn,1,3

if NOT isempty(request("RankSubmit")) then
	if session("id")=id then
		Call MsgBox("对不起，您不能连续对同一本商品评级！","Close","None")
		response.End
	end if
	rs("rank")=rs("rank")+1
	rs("ranknum")=rs("ranknum")+request("radiobutton")
	rs.update
	Call MsgBox("您对这本商品的评论星级已成功提交！","Close","None")
	rs.close
	set rs=nothing
	session("id")=id
	session.timeout=1
	response.End
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body>
<!--#include file="head.htm"-->


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="219" align="left" valign="top"><!--#include file="left.asp"--></td>
    <td width="561" align="left" valign="top">
      <br>      <br>      <form name="form1" method="post" action="">
        <table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr>
            <td height="15" align="center" >
              
            我要评级</td>
          </tr>
          <tr>
            
            <td height="33">
              <table width="100%" border="0" cellpadding="2" cellspacing="1" >
                <tr >
                  <td colspan="2">您对[<%=trim(rs("name"))%>]商品的评级是：
                  <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td>
                </tr>
                <tr >
                  <td width="50%">
                    <input name="radiobutton" type="radio" value="10" checked>
                    <img src="images/rank/10.gif" width="79" height="14"></td>
                  <td width="54%">
                    <input type="radio" name="radiobutton" value="9">
                    <img src="images/rank/9.gif" width="79" height="14"></td>
                </tr>
                <tr >
                  <td>
                    <input type="radio" name="radiobutton" value="8">
                    <img src="images/rank/8.gif" width="79" height="14"></td>
                  <td>
                    <input type="radio" name="radiobutton" value="7">
                    <img src="images/rank/7.gif" width="79" height="14"></td>
                </tr>
                <tr >
                  <td>
                    <input type="radio" name="radiobutton" value="6">
                    <img src="images/rank/6.gif" width="79" height="14"></td>
                  <td>
                    <input type="radio" name="radiobutton" value="5">
                    <img src="images/rank/5.gif" width="79" height="14"></td>
                </tr>
                <tr >
                  <td>
                    <input type="radio" name="radiobutton" value="4">
                    <img src="images/rank/4.gif" width="79" height="14"></td>
                  <td>
                    <input type="radio" name="radiobutton" value="3">
                    <img src="images/rank/3.gif" width="79" height="14"></td>
                </tr>
                <tr >
                  <td>
                    <input type="radio" name="radiobutton" value="2">
                    <img src="images/rank/2.gif" width="79" height="14"></td>
                  <td>
                    <input type="radio" name="radiobutton" value="1">
                    <img src="images/rank/1.gif" width="79" height="14"></td>
                </tr>
                <tr >
                  <td colspan="2">
                    <div align="center">
                      <input name="RankSubmit" type="submit" id="RankSubmit" value="提交">
                  </div></td>
                </tr>
            </table></td>
            
        </tr>
        </table>
    </form></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


