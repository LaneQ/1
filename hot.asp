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

</head>

<body>
<!--#include file="head.htm"-->


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="219" align="left" valign="top"><!--#include file="left.asp"--></td>
    <td width="561" align="left" valign="top">      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="left" valign="top"><br>            <img src="images/cxtsph.gif" width="212" height="32"></td>
        </tr>
        <tr>
          <td align="center" valign="top"><table width="568"  border="0" cellpadding="0" cellspacing="0">
<%'开始分页
Const MaxPerPage=5
dim totalPut   
dim CurrentPage
dim TotalPages
dim j
dim sql
if Not isempty(request("page")) then
	currentPage=Cint(request("page"))
else
	currentPage=1
end if 
set rs=server.CreateObject("adodb.recordset")
rs.open "select top 100 pagenum,discount,score,name,mark,vipprice,id,author,productdate,price2,price1,discount,pic from product order by solded desc",conn,1,1

if err.number<>0 then
	call MsgBox("数据库中无数据","Back","None")
	response.End
end if
if rs.eof And rs.bof then
	call MsgBox("对不起，目前没有更新商品！","Back","None")
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
                    <td width="50%">优惠价： <%=trim(rs("price2"))%></td>
                    <td width="50%">日　期：<%=trim(rs("productdate"))%></td>
                  </tr>
                  <tr class="bookinfo">
                    <td>定　价：<%=trim(rs("price1"))%></td>
                    <td>VIP价格：<%=rs("vipprice")%></td>
                  </tr>
                  <tr class="bookinfo">
                    <td>折　扣：<%=trim(rs("discount")*100)%></td>
                    <td> 积　分：<%=rs("score")%></td>
                  </tr>
                  <tr class="bookinfo">
                    <td width="50%">&nbsp;</td>
                    <td width="50%">&nbsp; </td>
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
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>首页 上一页</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1 class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&" class='contents'>上一页</a> "  
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>下一页 尾页</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&" class='contents'>尾页</a>"  
				End If  
					Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"页</font> "  
					Response.Write "<font class='contents'> 共有<b>&nbsp;"&totalnumber&"&nbsp;</b>种商品 "&maxperpage&"种商品/页</font> " 
					Response.Write "<font class='contents'>转到：</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='GO' name='cndok'></form>"  
				End Function  
			%>
			  
              <tr align="center">
                <td colspan="2"><br>
                  <form name="form1" method="post" action="">
                    <%
showpage totalput,MaxPerPage,"hot.asp"
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


