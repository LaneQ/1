<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 


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
          <td style="color:#415373">商品订单</td>
        </tr>
      </table>      <br>      <%
	  dim state,namekey
namekey=trim(request("namekey"))
state=trim(request("state"))
if state="" then state=request.QueryString("state")
if namekey="" then namekey=request.querystring("namekey")
%>      <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1">
        <tr>
          <td align="right">
            <select name="select" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
              <base target=Right>
              <option value="porder.asp?state=0" >全部订单状态</option>
              <option value="porder.asp?state=1" >未作任何处理</option>
              <option value="porder.asp?state=2" >用户已经划出款</option>
              <option value="porder.asp?state=3" >服务商已经收到款</option>
              <option value="porder.asp?state=4" >服务商已经发货</option>
              <option value="porder.asp?state=5" >用户已经收到货</option>
            </select>
          </td>
        </tr>
      </table>      <%
				Const MaxPerPage=12 
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
	if namekey="" then

	if state=0 or state="" then
	select case state
	case "0"
  rs.open "select distinct(goods),userid,realname,actiondate,deliverymethord,paymethord,state from orders where state<6 order by actiondate desc",conn,1,1
  case ""
  rs.open "select distinct(goods),userid,realname,actiondate,deliverymethord,paymethord,state from orders where state<5 order by actiondate desc",conn,1,1
  end select
  else
  rs.open "select distinct(goods),userid,realname,actiondate,deliverymethord,paymethord,state from orders where  state="&state&" order by actiondate",conn,1,1
  end if
  else
  '//按用户查询
  if state=0 or state="" then
  rs.open "select distinct(goods),userid,realname,actiondate,deliverymethord,paymethord,state from orders where state<5 and username='"&namekey&"' order by actiondate desc",conn,1,1
  else
  rs.open "select distinct(goods),userid,realname,actiondate,deliverymethord,paymethord,state from orders where  state="&state&" and username='"&namekey&"'  order by actiondate",conn,1,1
  end if
  end if
		  
				if err.number<>0 then
				response.write "数据库中无数据"
				end if
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> 对不起，您选择的状态目前还没有订单！</p>"
   				else
	  				totalPut=rs.recordcount

      				if currentpage<1 then
          				currentpage=1
      				end if

      				if (currentpage-1)*MaxPerPage>totalput then
	   					if (totalPut mod MaxPerPage)=0 then
	     					currentpage= totalPut \ MaxPerPage
	   					else
	      					currentpage= totalPut \ MaxPerPage + 1
	   					end if
      				end if

       				if currentPage=1 then
            			showContent
            			showpage totalput,MaxPerPage,"porder.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"porder.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"porder.asp"
	      				end if
	   				end if
   				   				end if

   				sub showContent
       			dim i
	   			i=0

			%>      <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr bgcolor="#FFFFFF" align="center">
          <td>订单号</td>
          <td>下单用户</td>
          <td>订货人姓名</td>
          <td> 付款方式</td>
          <td> 收货方式</td>
          <td > 订单状态</td>
        </tr>
        <%do while not rs.eof
		dim shop,username
		  set shop=server.CreateObject("adodb.recordset")
		  shop.open "select username from [user] where userid="&rs("userid"),conn,1,1
		  username=trim(shop("username"))
		  shop.close
		  set shop=nothing
		  %>
        <tr bgcolor="#FFFFFF" align="center">
          <td align="left">&nbsp;<a href="vorder.asp?dan=<%=trim(rs("goods"))%>&username=<%=username%>"><%=trim(rs("goods"))%></a></td>
          <td><%=username%></td>
          <td><%=trim(rs("realname"))%></td>
          <td>
            <%dim rs2
          set rs2=server.CreateObject("adodb.recordset")
          rs2.open "select * from delivery where deliveryid="&int(rs("paymethord")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "方式已被删除"
		  else
          response.Write trim(rs2("subject"))
          end if
		  rs2.Close
          set rs2=nothing
          %>
          </td>
          <td>
            <%
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select * from delivery where deliveryid="&int(rs("deliverymethord")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "方式已被删除"
		  else
          response.Write trim(rs2("subject"))
          end if
		  rs2.close
          set rs2=nothing%>
          </td>
          <td>
            <%
		  select case rs("state")
	case "1"
	response.write "未作任何处理"
	case "2"
	response.write "用户已经划出款"
	case "3"
	response.write "服务商已经收到款"
	case "4"
	response.write "服务商已经发货"
	case "5"
	response.write "用户已经收到货"
	end select%>
          </td>
        </tr>
        <%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		loop
		rs.close
		set rs=nothing%>
      </table>      <%  
				End Sub   
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				if namekey="" then
				Response.Write "<form method=Post action="&filename&"?state="&state&">"  
				else
				Response.Write "<form method=Post action="&filename&"?state="&state&"&namekey="&namekey&">" 
				end if
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>首页 上一页</font> "  
				Else  
					if namekey="" then
					Response.Write "<a href="&filename&"?page=1&state="&state&" class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&state="&state&" class='contents'>上一页</a> "  
					ELSE
					Response.Write "<a href="&filename&"?page=1&state="&state&"&namekey="&namekey&" class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&state="&state&"&namekey="&namekey&" class='contents'>上一页</a> "
					end if  
				End If
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>下一页 尾页</font>"  
				Else 
				if namekey="" then
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&state="&state&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&"&state="&state&" class='contents'>尾页</a>"
					else
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&state="&state&"&namekey="&namekey&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&"&state="&state&"&namekey="&namekey&" class='contents'>尾页</a>" 
					end if 
				End If  
					Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"页</font> "  
					Response.Write "<font class='contents'> 共有"&totalnumber&"笔订单 "&maxperpage&"笔订单/页</font> " 
					Response.Write "<font class='contents'>转到：</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='GO' name='cndok'></form>"  
				End Function  
			%>      <br>      <table border="0" align="left" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">订单查找</td>
        </tr>
      </table>      <br>      <br>	  <form name="form1" method="post" action="">
        <table width="80%" border="0" align="left" cellpadding="1" cellspacing="1">
          <tr align="center">
            
            <td>按下单用户查询
                <input name="namekey" type="text" id="namekey" value="请输入用户名" size="14" onFocus="this.value=''">
                <select name="state" id="select2">
                  <option value="0" >全部订单状态</option>
                  <option value="1" >未作任何处理</option>
                  <option value="2" >用户已经划出款</option>
                  <option value="3" >服务商已经收到款</option>
                  <option value="4" >服务商已经发货</option>
                  <option value="5" >用户已经收到货</option>
                </select>
&nbsp;
          <input type="submit" name="Submit" value="查 询">
            </td>
            
        </tr>
        </table>
	    </form>	  <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


