<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 

<%
dim selectm,selectkey,selectid
selectkey=trim(request(trim("selectkey")))
selectm=trim(request("selectm"))
if selectkey="" then
	selectkey=request.QueryString("selectkey")
end if

if selectm="" then
	selectm=request.QueryString("selectm")
end if
selectid=request("selectid")

if selectid<>"" then
	if session("rank")>1 then
	call Msgbox("���Ȩ�޲�����","Back","None")
	response.End
	end if
	conn.execute "delete from product where id in ("&selectid&")"
	response.Redirect "mpro.asp"
	response.End
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
          <td style="color:#415373">��Ʒ�鿴���޸�</td>
        </tr>
      </table>      <script language=javascript>
function test()
{
  if(!confirm('ȷ��ɾ����')) return false;
}


function mm()
{
   var a = document.getElementsByTagName("input");
   if(a[0].checked==true){
   for (var i=0; i<a.length; i++)
      if (a[i].type == "checkbox") a[i].checked = false;
   }
   else
   {
   for (var i=0; i<a.length; i++)
      if (a[i].type == "checkbox") a[i].checked = true;
   }
}
      </script>      <br>      <%
				Const MaxPerPage=20 
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
			select case selectm
			case ""
            rs.open "select id,name,adddate,mark,author from product order by adddate desc",conn,1,1
		    case "0"
			response.write "<center>�뷵��ѡ����Ҫ��ѯ�ķ�ʽ��<br><br><a href=javascript:history.go(-1)>���������һҳ</a></center>"
			response.End
			case "name"
			rs.open "select id,name,adddate,mark,author from product where name like '%"&selectkey&"%' order by adddate desc",conn,1,1
			case "zuozhe"
			rs.open "select id,name,adddate,mark,author from product where makein like '%"&selectkey&"%' order by adddate desc",conn,1,1
			case "chubanshe"
			rs.open "select id,name,adddate,mark,author from product where mark like '%"&selectkey&"%' order by adddate desc",conn,1,1
		  end select
		   	if err.number<>0 then
				response.write "���ݿ���������"
				end if
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> ���ݿ��������ݣ�</p>"
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
            			showpage totalput,MaxPerPage,"mpro.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"mpro.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"mpro.asp"
	      				end if
	   				end if
   				   				end if

   				sub showContent
       			dim i
	   			i=0%>      <br>      <form name="form2" method="post" action="">
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr bgcolor="#FFFFFF" align="center" height="20">
            <td height="20" bgcolor="#FFFFFF">����</td>
            <td height="20">����</td>
            <td height="20">������</td>
            <td height="20">����ʱ��</td>
            <td>ѡ��</td>
          </tr>
          <%
		  do while not rs.eof%>
          <tr bgcolor="#FFFFFF" align="center">
            <td align="left">&nbsp;<a href=epro.asp?id=<%=rs("id")%>>
              <% if len(trim(rs("name")))>20 then
			response.write left(trim(rs("name")),18)&"..."
			else
			response.write trim(rs("name"))
			end if%>
            </a></td>
            <td align="left">              <% if len(trim(rs("author")))>20 then
			response.write left(trim(rs("author")),18)&"..."
			else
			response.write trim(rs("author"))
			end if%>            </td>
            <td align="left">&nbsp;
                <%if len(trim(rs("mark")))>30 then
			response.write left(trim(rs("mark")),28)&"..."
			else 
			response.write trim(rs("mark"))
			end if%>
            </td>
            <td nowrap><%=rs("adddate")%></td>
            <td align="center"><input name="selectid" type="checkbox" id="selectid" value="<%=rs("id")%>"></td>
          </tr>
          <%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
          <tr bgcolor="#FFFFFF">
            <td height="30" colspan="5" align="right">ȫѡ
                <input type="checkbox" name="checkbox2" value="Check All" onClick="mm()">
&nbsp;
      <input type="submit" name="Submit" value="ɾ ��" onClick="return test();">
&nbsp;&nbsp; </td>
          </tr>
        </table>
      </form>      <%  
				End Sub   
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				
				Response.Write "<form method=Post action="&filename&"?selectm="&selectm&"&selectkey="&selectkey&" >"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>��ҳ ��һҳ</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>��һҳ</a> "  
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>��һҳ βҳ</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>βҳ</a>"  
				End If  
					Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"ҳ</font> "  
					Response.Write "<font class='contents'> ����"&totalnumber&"����Ʒ " 
					Response.Write "<font class='contents'>ת���ڣ�</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">ҳ"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='��ת' name='cndok' ></form>"  
				End Function  
			%>      <br>      <table border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">��Ʒ����</td>
        </tr>
      </table>      <form name="form1" method="post" action="">
        <table border="0" align="center" cellpadding="3" cellspacing="3">
          <tr bgcolor="#FFFFFF" align="center">
            <td>
              <input name="selectkey" type="text" id="selectkey2" onFocus="this.value=''" value="������ؽ���">
            </td>
            <td>
              <select name="selectm" id="select">
                <OPTION VALUE="name">����Ʒ����</OPTION>
                <OPTION VALUE="zuozhe">����Ʒ���</OPTION>
                <OPTION VALUE="chubanshe">����Ʒ����</OPTION>
              </select>
            </td>
            <td><input type="submit" name="Submit2" value="�� ѯ"></td>
          </tr>
        </table>
      </form>      <br>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


