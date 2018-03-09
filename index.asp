<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
dim i
set rs=server.CreateObject("adodb.recordset")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
<link href="style.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style></head>

<body>
<!--#include file="head.htm"-->


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="219" align="left" valign="top"><!--#include file="left.asp"--></td>
    <td width="561" align="left" valign="top">      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="44" align="center"><form name="form1" method="get" action="rsearch.asp">
              <br>
        书名：
        <INPUT name=name class="inputstyle">
        <select name=code>
                <option value="" selected>所有图书</option>
                <%

rs.open "select * from category",conn,1,1
do while not rs.eof
%>
                <option value="<%=rs("categoryid")%>" ><%=rs("category")%></option>
                <%
rs.movenext
loop
rs.close
%>
              </select>
        <input name="Submit" type="submit" class="inputstyle" value="快速搜索">
        <input type="button" name="Submit2" value="高级搜索" onclick="location='search.asp'">        
        <img src="images/line.gif" width="568" height="9">
		
          </form></td>
        </tr>
        <tr>
          <td align="left" valign="top"><img src="images/xssj.gif" width="221" height="28"></td>
        </tr>
        <tr>
          <td align="center" valign="top"><table width="568"  border="0" cellpadding="0" cellspacing="0">
<%
	  rs.open "select top 6 id,detail,discount,vipprice,author,desc,price1,price2,name,pic,mark from product order by id desc",conn,1,1
	  if rs.eof and rs.bof then
		  response.write "　　对不起，暂时还没有商品！"
		  end if
		  i=0
		  do while not rs.eof
%>
		  
		  
              <tr>
                <td width="17%" height="130" align="center" valign="middle" class="shadow"><a href="vpro.asp?id=<%=trim(rs("id"))%>" target="_blank"><img src="<%=trim(rs("pic"))%>" width="85" height="125" border="0"></a></td>
                <td width="33%" align="left" valign="top"><table width="100%"  border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td colspan="2"><img src="images/w.gif" width="18" height="18"><span class="booktitle"><%=strvalue(trim(rs("name")),24)%></span></td>
                    </tr>
                    <tr>
                      <td height="40" colspan="2" class="bookinfo"><%=trim(rs("desc"))%><br>
                      <br></td>
                    </tr>
                    <tr>
                      <td>定价:<span class="price1"><%=rs("price1") %></span>元</td>
                      <td>作者:<%=rs("author") %></td>
                    </tr>
                    <tr>
                      <td>优惠价:<span class="price2"><%=rs("price2") %></span>元</td>
                      <td>VIP价:<span class="viprice"><%=rs("vipprice") %></span></td>
                    </tr>
                    <tr>
                      <td colspan="2" align="center"><a href="icar.asp?id=<%=rs("id")%>&action=add" target="pcart"><img src="images/car.gif" width="23" height="20" border="0">加入购物车</a> </td>
                    </tr>
                </table></td>
				<%
				rs.movenext
				if rs.eof then
					response.write "<td width='17%'></td><td></td>"
				else
				%>
                <td width="17%" height="130" align="center" valign="middle" class="shadow"><a href="vpro.asp?id=<%=trim(rs("id"))%>" target="_blank"><img src="<%=trim(rs("pic"))%>" width="85" height="125" border="0"></a></td>
                <td width="33%" align="left" valign="top"><table width="100%"  border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td colspan="2"><img src="images/w.gif" width="18" height="18"><span class="booktitle"><%=strvalue(trim(rs("name")),24)%></span></td>
                    </tr>
                    <tr>
                      <td height="40" colspan="2" class="bookinfo"><%=trim(rs("desc"))%><br>
                      <br></td>
                    </tr>
                    <tr>
                      <td>定价:<span class="price1"><%=rs("price1") %></span>元</td>
                      <td>作者:<%=rs("author") %></td>
                    </tr>
                    <tr>
                      <td>优惠价:<span class="price2"><%=rs("price2") %></span>元</td>
                      <td>VIP价:<span class="viprice"><%=rs("vipprice") %></span>元</td>
                    </tr>
                    <tr>
                      <td colspan="2" align="center"><a href="icar.asp?id=<%=rs("id")%>&action=add" target="pcart"><img src="images/car.gif" width="23" height="20" border="0">加入购物车</a> </td>
                    </tr>
                </table></td>
				<%
				end if
				%>
              </tr>
			                <tr>
                <td colspan="4" align="center"><img src="images/line.gif" width="568" height="9"></td>
              </tr>

<%
i=i+1
			  if i>=5 then exit do
			  if not rs.eof then   rs.movenext
			  loop
			  rs.close
%>			  
			  
			  
              <tr align="right">
                <td colspan="4"><table width="100" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td align="left"><a href="new.asp"><img src="images/more_2.gif" width="42" height="15" border="0"></a></td>
                      <td width="10">&nbsp;</td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td colspan="4" align="center"><img src="images/line.gif" width="568" height="9"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td align="left" valign="top"><img src="images/tzts.gif" width="221" height="27"></td>
        </tr>
        <tr>
          <td><table width="568"  border="0" cellpadding="0" cellspacing="0">
            <%
	  rs.open "select top 6 id,detail,desc,discount,vipprice,author,price1,price2,name,pic,mark from product where recommend=1 order by adddate desc",conn,1,1

	  if rs.eof and rs.bof then
		  response.write "　　对不起，暂时还没有商品！"
		  'response.End
		  end if
		  i=0
		  do while not rs.eof
%>
            <tr>
              <td width="17%" height="130" align="center" valign="middle" class="shadow"><a href="vpro.asp?id=<%=trim(rs("id"))%>" target="_blank"><img src="<%=trim(rs("pic"))%>" width="85" height="125" border="0"></a></td>
              <td width="33%" align="left" valign="top"><table width="100%"  border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td colspan="2"><img src="images/w.gif" width="18" height="18"><span class="booktitle"><%=strvalue(trim(rs("name")),24)%></span></td>
                  </tr>
                  <tr>
                    <td height="40" colspan="2" class="bookinfo"><%=trim(rs("desc"))%><br>
                      <br></td>
                  </tr>
                  <tr>
                    <td>定价:<span class="price1"><%=rs("price1") %></span>元</td>
                    <td>作者:<%=rs("author") %></td>
                  </tr>
                  <tr>
                    <td>优惠价:<span class="price2"><%=rs("price2") %></span>元</td>
                    <td>VIP价:<span class="viprice"><%=rs("vipprice") %></span>元</td>
                  </tr>
                  <tr>
                    <td colspan="2" align="center"><a href="icar.asp?id=<%=rs("id")%>&action=add" target="pcart"><img src="images/car.gif" width="23" height="20" border="0">加入购物车</a> </td>
                  </tr>
              </table></td>
              <%
				rs.movenext

				if rs.eof then
					response.write "<td width='17%'></td><td></td>"
				else

				%>
              <td width="17%" height="130" align="center" valign="middle" class="shadow"><a href="vpro.asp?id=<%=trim(rs("id"))%>" target="_blank"><img src="<%=trim(rs("pic"))%>" width="85" height="125" border="0"></a></td>
              <td width="33%" align="left" valign="top"><table width="100%"  border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td colspan="2"><img src="images/w.gif" width="18" height="18"><span class="booktitle"><%=strvalue(trim(rs("name")),24)%></span></td>
                  </tr>
                  <tr>
                    <td height="40" colspan="2" class="bookinfo"><%=trim(rs("desc"))%><br>
                      <br></td>
                  </tr>
                  <tr>
                    <td>定价:<span class="price1"><%=rs("price1") %></span>元</td>
                    <td>作者:<%=rs("author") %></td>
                  </tr>
                  <tr>
                    <td>优惠价:<span class="price2"><%=rs("price2") %></span>元</td>
                    <td>VIP价:<span class="viprice"><%=rs("vipprice") %></span>元</td>
                  </tr>
                  <tr>
                    <td colspan="2" align="center"><a href="icar.asp?id=<%=rs("id")%>&action=add" target="pcart"><img src="images/car.gif" width="23" height="20" border="0">加入购物车</a> </td>
                  </tr>
              </table></td>
				<%
				end if
				%>


            </tr>
			            <tr>
              <td colspan="4" align="center"><img src="images/line.gif" width="568" height="9"></td>
            </tr>

            <%
i=i+1
			  if i>=5 then exit do
				if not rs.eof then   rs.movenext
			  loop
			  rs.close
			  set rs=nothing
%>
            <tr align="right">
              <td colspan="4"><table width="100" border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td align="left"><a href="commend.asp"><img src="images/more_2.gif" width="42" height="15" border="0"></a></td>
                    <td width="10">&nbsp;</td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td colspan="4" align="center"><img src="images/line.gif" width="568" height="9"></td>
            </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


