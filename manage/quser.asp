<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%
if session("rank")>1 then
	call Msgbox("���Ȩ�޲�����","Back","None")
	response.End
end if
%>

<%

if NOT isempty(request("DelQuserSubmit")) then
	dim userid
	userid=request.QueryString("id")
	if userid="" then userid=request("userid")
	conn.execute "delete from [user] where userid in ("&userid&")"
	conn.execute "delete from orders where userid in ("&userid&")"
	response.Redirect "quser.asp"

end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>У԰�����</title>
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
          <td style="color:#415373">ǰ̨�û�����</td>
        </tr>
      </table>      <br>      <%
Const MaxPerPage=20 
dim totalPut,CurrentPage,TotalPages,j
	if Not isempty(request("page")) then
    	currentPage=Cint(request("page"))
   	else
    	currentPage=1
   	end if 			
dim namekey,checkbox,vipuser
checkbox=request("checkbox")
namekey=request("namekey")
if namekey="" then namekey=request.QueryString("namekey")
	if checkbox="" then checkbox=request.querystring("checkbox")
		set rs=server.CreateObject("adodb.recordset")
		if namekey="" then
			rs.open "select username,userid,realname,vip,score,loginnum,adddate from [user] order by adddate desc",conn,1,1
		else
		if checkbox=1 then
			rs.open "select username,userid,realname,vip,score,loginnum,adddate from [user] where username like '%"&namekey&"%' order by adddate desc",conn,1,1
		else
			rs.open "select username,userid,realname,vip,score,loginnum,adddate from [user] where username='"&namekey&"' order by adddate desc",conn,1,1
		end if
	end if
	if err.number<>0 then
		response.write "���ݿ���������"
	end if
				
  	if rs.eof And rs.bof then
    	Response.Write "<p align='center' class='contents'> �Բ���û�д��û���</p>"
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
        showpage totalput,MaxPerPage,"quser.asp"
	else
    	if (currentPage-1)*MaxPerPage<totalPut then
        	rs.move  (currentPage-1)*MaxPerPage
            dim bookmark
            bookmark=rs.bookmark
            showContent
            showpage totalput,MaxPerPage,"quser.asp"
        else
	    	currentPage=1
        	showContent
        	showpage totalput,MaxPerPage,"quser.asp"
	    end if
	end if
end if

sub showContent
	dim i
	i=0
%>      <form name="form1" method="post" action="">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr height="20">
            <td WIDTH="15%" align="center">�û���</td>
            <td WIDTH="15%" align="center">��ʵ����</td>
            <td WIDTH="20%" align="center">ע��ʱ��</td>
            <td WIDTH="10%" align="center">��Ա����</td>
            <td WIDTH="10%" align="center">����</td>
            <td WIDTH="10%" align="center"> ��¼����</td>
            <td WIDTH="10%" align="center">ѡ ��</td>
          </tr>
          <%do while not rs.eof
 		if rs("vip") = true then
		    vipuser="VIP��Ա"
		  else
		    vipuser="��ͨ��Ա"
		  end if
		  %>
          <tr align="center" height="20">
            <td align="center"><a href=equser.asp?id=<%=rs("userid")%>><%=trim(rs("username"))%></a></td>
            <td><%=trim(rs("realname"))%></td>
            <td><%=rs("adddate")%></td>
            <td><%=vipuser %></td>
            <td><%=rs("score")%></td>
            <td> <%=rs("loginnum")%>��</td>
            <td>
              <input name="userid" type="checkbox" id="userid" value="<%=rs("userid")%>" ></td>
          </tr>
          <%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		  loop%>
        </table>
        <br>
        <br>
        <div align="center">
          <input name="DelQuserSubmit" type="submit" id="DelQuserSubmit" onClick="return confirm('��ȷ��Ҫ����������')" value="ɾ����ѡ�û�">
  ȫѡ
  <input type="checkbox" name="checkbox2" value="Check All" onClick="mm()">
        </div>
      </form>      <%  
				End Sub   
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				
				Response.Write "<form method=Post action="&filename&"?checkbox="&checkbox&"&namekey="&namekey&">"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>��ҳ ��һҳ</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&checkbox="&checkbox&"&namekey="&namekey&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&checkbox="&checkbox&"&namekey="&namekey&" class='contents'>��һҳ</a> "  
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>��һҳ βҳ</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&checkbox="&checkbox&"&namekey="&namekey&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&checkbox="&checkbox&"&namekey="&namekey&" class='contents'>βҳ</a>"  
				End If  
					Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"ҳ</font> "  
					Response.Write "<font class='contents'> ����"&totalnumber&"��ע���û� " 
					Response.Write "<font class='contents'>ת����</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='GO' name='cndok'></form>"  
				End Function  
			%>      <table border="0" align="left" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">�����û�</td>
        </tr>
      </table>      <br>      <br>      <br>      <form name="form3" method="post" action="quser.asp?action=select">
        <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td align="center">���û�������:
                <input name="namekey" type="text" id="namekey4" size="12">
&nbsp;
      <input name="checkbox" type="checkbox" id="checkbox4" value="1" checked>
      ģ����ѯ
      <input type="submit" name="Submit2" value="�� ѯ"></td>
          </tr>
        </table>
      </form>            <br>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


