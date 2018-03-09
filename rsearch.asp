<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 

<%
if request.QueryString="" then 
	call MsgBox("错误：没有搜索条件!","GoUrl","search.asp")
	response.end
end if

'开始分页
Const MaxPerPage=5
dim totalPut   
dim CurrentPage
dim TotalPages
dim j
dim sql
dim nid,sortsid


if Not isempty(request.QueryString("page")) then
	currentPage=Cint(request.QueryString("page"))
else
	currentPage=1
end if 
set rs=server.CreateObject("adodb.recordset")


dim name
dim author
dim manufacturer
dim enabledate
dim expiredate
dim smallprice
dim largeprice
dim code
dim OrderField
dim Order

name=trim(request.QueryString("name"))
author=trim(request.QueryString("author"))
manufacturer=trim(request.QueryString("manufacturer"))
enabledate=trim(request.QueryString("enabledate"))
expiredate=trim(request.QueryString("expiredate"))
smallprice=trim(request.QueryString("smallprice"))
largeprice=trim(request.QueryString("largeprice"))
OrderField=trim(request.QueryString("OrderField"))
Order=trim(request.QueryString("Order"))
code=trim(request.QueryString("code"))

if OrderField="" then OrderField="adddate"
if Order="" then Order="DESC"

sql="select pagenum,name,mark,vipprice,id,author,productdate,price2,price1,discount,pic from product where 1=1 "

if name<>"" then
	sql=sql&"and name like '%"&name&"%' "
end if

if author<>"" then
	sql=sql&"and author like '%"&author&"%' "
end if

if manufacturer<>"" then
	sql=sql&"and mark like '%"&manufacturer&"%' "
end if

if code<>"" then
	sql=sql&"and categoryid like '%"&code&"%' "
end if

if smallprice<>"" then 
	smallprice=CDbl(smallprice)
	sql=sql&"and price2 >= "&smallprice
end if

if largeprice<>"" then
	largeprice=CDbl(largeprice)
	sql=sql&"and price2 <= "&largeprice
end if

if expiredate<>"" then
	expiredate=CDate(expiredate)
	sql=sql&"and productdate <= #"&expiredate&"#"
end if

if enabledate<>"" then
	enabledate=CDate(enabledate)
	sql=sql&"and productdate >= #"&enabledate&"#"
end if



sql=sql&" order by "&OrderField&" "&Order

rs.open sql,conn,1,1

if err.number<>0 then
	call MsgBox("数据库中无数据","Back","None")
	response.End
end if
if rs.eof And rs.bof then
	call MsgBox("对不起，找不到你所需的书籍！","Back","None")
	response.End
else
	totalPut=rs.recordcount

if currentpage<1 then
	urrentpage=1
end if

if (currentpage-1)*MaxPerPage>totalput then
	if (totalPut mod MaxPerPage)=0 then
		currentpage= totalPut \ MaxPerPage
	else
		currentpage= totalPut \ MaxPerPage + 1
	end if
end if

if currentPage=1 then
else
	if (currentPage-1)*MaxPerPage<totalPut then
		rs.move  (currentPage-1)*MaxPerPage
		dim bookmark
        bookmark=rs.bookmark
        
	else
		currentPage=1
	end if
	end if

end if

dim i
i=0

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
    <td width="561" align="left" valign="top">      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="left" valign="top"><br>
          <br></td>
        </tr>
        <tr>
          <td align="center" valign="top"><table width="568"  border="0" cellpadding="0" cellspacing="0">
<%
do while not rs.eof
%>
			  <tr>
                <td width="17%" height="130" align="center" valign="middle" class="shadow"><a href="vpro.asp?id=<%=trim(rs("id"))%>" target="_blank"><img src="<%=trim(rs("pic"))%>" width="85" height="125" border="0"></a></td>
                <td height="130" align="left" valign="top"><table width="100%"  border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td colspan="2"><img src="images/w.gif" width="18" height="18"><span class="booktitle"><%=trim(rs("name"))%></span></td>
                  </tr>
                  <tr class="bookinfo">
                    <td width="50%" height="12" class="bookinfo">作　者：<%=trim(rs("author"))%></td>
                    <td width="50%" class="bookinfo"> 出版社：<%=trim(rs("mark"))%></td>
                  </tr>
                  <tr class="bookinfo">
                    <td width="50%">日　期：<%=trim(rs("productdate"))%></td>
                    <td width="50%">VIP价格：<%=rs("vipprice")%></td>
                  </tr>
                  <tr class="bookinfo">
                    <td width="50%"> 定　价：<%=trim(rs("price1"))%></td>
                    <td width="50%"> 优惠价： <%=trim(rs("price2"))%></td>
                  </tr>
                  <tr>
                    <td colspan="2" align="center"><a href="icar.asp?id=<%=rs("id")%>&action=add" target="pcart"><img src="images/car.gif" width="23" height="20" border="0">购物车</a></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td colspan="2" align="center"><img src="images/line.gif" width="568" height="9"></td>
              </tr>
<%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
			loop
			rs.close
			set rs=nothing%>
                                                      <%  
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				
				Response.Write "<form method=Post action="&filename&">"  
				Response.Write "<p align='center' > "  
				If CurrentPage<2 Then  
					Response.Write "首页 上一页 "  
				Else  
					Response.Write "<a href="&filename&"page=1>首页</a> "  
					Response.Write "<a href="&filename&"page="&CurrentPage-1&" >上一页</a> "  
				End If
				
				If n-currentpage<1 Then  
					Response.Write "下一页 尾页"  
				Else  
					Response.Write "<a href="&filename&"page="&(CurrentPage+1)&" >"  
					Response.Write "下一页</a> <a href="&filename&"page="&n&">尾页</a>"  
				End If  
					Response.Write " 页次："&CurrentPage&"/"&n&"页 "  
					Response.Write " 共有<b>&nbsp;"&totalnumber&"&nbsp;</b>种商品 "&maxperpage&"种商品/页 " 
					Response.Write "转到：<input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'   value='GO' name='cndok'></form>"  
				End Function  
			%>
			  
              <tr align="center">
                <td colspan="2"><br>

                  <form name="form1" method=get action="">
				  				<INPUT TYPE="hidden" name="name" value="<%=name%>">
				<INPUT TYPE="hidden" name="author" value="<%=author%>">
				<INPUT TYPE="hidden" name="manufacturer" value="<%=manufacturer%>">
				<INPUT TYPE="hidden" name="enabledate" value="<%=enabledate%>">
				<INPUT TYPE="hidden" name="expiredate" value="<%=expiredate%>">
				<INPUT TYPE="hidden" name="smallprice" value="<%=smallprice%>">
				<INPUT TYPE="hidden" name="largeprice" value="<%=largeprice%>">
				<INPUT TYPE="hidden" name="code" value="<%=code%>">
				<INPUT TYPE="hidden" name="Order" value="<%=Order%>">
				<INPUT TYPE="hidden" name="OrderField" value="<%=OrderField%>">

                    <%
showpage totalput,MaxPerPage,"rsearch.asp?name="&name&"&author="&author&"&manufacturer="&manufacturer&"&enabledate="&enabledate&"&expiredate="&expiredate&"&smallprice="&smallprice&"&largeprice="&largeprice&"&code="&code&"&Order="&Order&"&OrderField="&OrderField&"&"
%>
                  </form> </td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
    </table></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


