<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.title {font-size:14px;color:#415373;font-weight:bold;}
-->
</style>

</head>

<body>
<!--#include file="head.htm"-->


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="219" align="left" valign="top"><!--#include file="left.asp"--></td>
    <td width="561" align="left" valign="top">
      <br>      <table width="100%" border="0" cellpadding="2" cellspacing="2">
  <%        
		  set rs=server.CreateObject("adodb.recordset")
		  rs.open "select category,categoryid from category",conn,1,1
		  if rs.eof then response.write "对不起！还没有添加任何的分类！"
		  do while not rs.eof
%>
	    <tr>
          <td>
            <table border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td><img src="images/w.gif"></td>
                <td><span class="title"><%=rs("category")%></span></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td style="PADDING-LEFT: 30px;">
		    <%
		  	dim rsSub
			set rsSub=server.CreateObject("adodb.recordset")
			rsSub.open "select sorts,sortsid from sorts where categoryid="&rs("categoryid")&" order by sortsorder",conn,1,1
			if rsSub.recordcount=0 then response.Write "对不起！此大类没有添加小类！"
			do while not rsSub.eof
			response.Write "<a href=sub.asp?aid="&rs("categoryid")&"&nid="&rsSub("sortsid")&">"&trim(rsSub("sorts"))&"</a>  "
			'response.write "<A href=class.asp?aid="&rs("categoryid")&"&nid="&rsSub("sortsid")&">"&trim(rsSub("sorts"))&"</A> | " 
             rsSub.movenext
			 loop
			 rsSub.close
			 set rsSub=nothing
			%>  
	      </td>
        </tr>
  <%
             rs.movenext
			 loop
			 rs.close
			 set rs=nothing
%>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>      <br>      </td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


