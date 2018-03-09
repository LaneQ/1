<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
dim id
id=request.QueryString("id")
dim prename,company,intro,predate,graph2,description,remarks,price
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from product where id="&id,conn,1,3
rs("viewnum")=rs("viewnum")+1
rs.update

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
      <br>      <table width="568"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="17%" height="130" align="center" valign="middle" class="shadow"><img src="<%=trim(rs("pic"))%>" width="85" height="125"></td>
          <td height="130" align="left" valign="top"><table width="100%"  border="0" cellspacing="2" cellpadding="0">
            <tr>
              <td colspan="2"><table border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td><img src="images/w.gif" width="18" height="18"></td>
                    <td><span class="booktitle"><%=trim(rs("name"))%></span></td>
                  </tr>
              </table></td>
            </tr>
            <tr class="bookinfo">
              <td width="50%" class="bookinfo">作　者：<%=trim(rs("author"))%></td>
              <td width="50%" class="bookinfo"> ISBN ：<%=trim(rs("type"))%></td>
            </tr>
            <tr class="bookinfo">
              <td width="50%"> 出版社：<%=trim(rs("mark"))%></td>
              <td width="50%"> 开　本：<%=trim(rs("format"))%> </td>
            </tr>
            <tr class="bookinfo">
              <td>出版日期：<%=trim(rs("productdate"))%></td>
              <td> 页　数：<%=trim(rs("pagenum"))%> </td>
            </tr>
            <tr class="bookinfo">
              <td> 装　帧：<%=trim(rs("introduce"))%> </td>
              <td> 版　次：<%=trim(rs("printed"))%> </td>
            </tr>
            <tr class="bookinfo">
              <td>定　价：<%=trim(rs("price1"))%> </td>
              <td>优惠价：<%=trim(rs("price2"))%></td>
            </tr>
            <tr class="bookinfo">
              <td>积　分：<%=rs("score")%></td>
              <td>VIP价格：<%=rs("vipprice")%></td>
            </tr>
            <tr class="bookinfo">
              <td>浏　览：<%=trim(rs("viewnum"))%></td>
              <td>购买：<%=trim(rs("solded"))%></td>
            </tr>
            <tr>
              <td colspan="2" align="center"><a href="icar.asp?id=<%=rs("id")%>&action=add" target="pcart"><img src="images/car.gif" width="23" height="20" border="0">购物车</a></td>
            </tr>
          </table></td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:10px;">            <table border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td><img src="images/w.gif" width="18" height="18"></td>
                <td style="padding-left:0px;"><strong><%=trim(rs("name"))%></strong></td>
              </tr>
            </table></td>
        </tr>
        <tr align="left">
          <td colspan="2" style="padding-left:40px;"> <%=trim(rs("detail"))%></td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:10px;"><table border="0" cellspacing="0" cellpadding="2">
            <tr>
              <td><img src="images/w.gif" width="18" height="18"></td>
              <td> <strong>目录 </strong> </td>
            </tr>
          </table></td>
        </tr>
        <tr align="left">
          <td colspan="2" style="padding-left:40px;"><%=trim(rs("content"))%></td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:10px;"><table border="0" cellpadding="2" cellspacing="0">
            <tr>
              <td><img src="images/w.gif" width="18" height="18"></td>
              <td> <strong>会员评级</strong> <a href="rank.asp?id=<%=id%>" target="_blank">发表您对这本书的评级</a></td>
            </tr>
          </table></td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:40px;"><%
		'用户评级
		
if rs("ranknum")>0 and rs("rank")>0 then
dim other
other=rs("ranknum")\rs("rank")
else
other=0
end if
response.write "<img src=images/rank/"&other&".gif alt=评论星级>"

		rs.close
%>            </td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:10px;"><table border="0" cellpadding="2" cellspacing="0">
            <tr>
              <td><img src="images/w.gif" width="18" height="18"></td>
              <td> <strong>会员评论</strong> <a href="comment.asp?id=<%=id%>" target="_blank">发表您对这本书的评论</a></td>
            </tr>
          </table> </td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:40px;">
		  <%
		rs.open "select * from review where id="&id&" and audit=1 ",conn,1,1
		if rs.eof and rs.bof then
		response.write "如果您用过本商品，或对本商品有所了解，欢迎您发表自己的评论。您的评论将被网络上成千上万的用户所共享，我们将对您的慷慨深表感谢。<br>"
		response.write "您的评论在提交后将经过我们的审核，也许您需要等待一些时间才可以看到。谢谢合作。"
		else
		do while not rs.eof 
				%>
        [<B><%=rs("title")%></B>@<%=rs("reviewdtm")%>]<BR>
                  <%=rs("reviewcontent")%><br>
 
            <%rs.movenext
		loop
		end if
		rs.close
		%>
		  </td>
        </tr>
        <tr align="left">
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr align="center">
          <td colspan="2"><TABLE WIDTH="96%" BORDER="0" CELLPADDING="0" CELLSPACING="0" align="center">
            <a name="pic"></a>
            <TR>
              <TD align="center"><B>本站发表用户评论，并不代表我们赞同或者支持用户的观点。我们的立场仅限于传播更多用户感兴趣的信息。</B></TD>
            </TR>
            <TR>
              <TD align="center">&nbsp;</TD>
            </TR>
            <TR>
              <TD align="center"><input type="button" name="Submit" value="关闭" onClick="window.close()"></TD>
            </TR>
          </TABLE></td>
        </tr>
        <tr align="center">
          <td colspan="2">&nbsp;</td>
        </tr>
    </table></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>
<%

set rs=nothing

%>

