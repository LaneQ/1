<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 

<%
dim action
action=request.QueryString("action")

select case action
	case "del"
		if request("audit").count=0 then
			call MsgBox("��û��ѡ��Ҫɾ�������ۣ�","None","None")
		else
			if session("rank")>1 then
			call Msgbox("���Ȩ�޲�����","Back","None")
			response.End
			end if
			conn.execute ("delete from review where reviewid in ("&request("audit")&")")
			call MsgBox("����ɾ���ɹ�!","None","None")
		end if
	case "audit"
		if request("audit").count=0 then
			call MsgBox("��û��ѡ��Ҫ��˵����ۣ�","None","None")
		else
			if session("rank")>1 then
			call Msgbox("���Ȩ�޲�����","Back","None")
			response.End
			end if
			conn.execute "update review set audit=1 where reviewid in ("&request("audit")&")"
			call MsgBox("������˳ɹ�!","None","None")
		end if
	case "delzhou"
		if session("rank")>1 then
		call Msgbox("���Ȩ�޲�����","Back","None")
		response.End
		end if

		dim theday
		theday=date-7
		conn.execute ("delete from review where reviewdtm<#"&theday&"# and audit=0")
		call MsgBox("һ��ǰδ�������ɾ���ɹ�!","None","None")
	case "delall"
		if session("rank")>1 then
		call Msgbox("���Ȩ�޲�����","Back","None")
		response.End
		end if

		conn.execute ("delete from review where audit=0")
		call MsgBox("����δ�������ɾ���ɹ�!","None","None")

end select

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
          <td style="color:#415373">��Ʒ����</td>
        </tr>
      </table>      <br>      <%dim dtype
dtype=request.QueryString("dtype")
if dtype="" then 
	dtype="no"
end if
%>      <form name="form1" method="post" action="">
  
      <table width="98%" border="0" align="center" cellpadding="2" cellspacing="2">
          <tr>
            <td bgcolor="#FFFFFF" align="center"><a href="procomm.asp?dtype=no">δ��˵�����</a></td>
            <td bgcolor="#FFFFFF" align="center"><a href="procomm.asp?dtype=yes">����˵�����</a></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td colspan="2"><%
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
		  select case dtype
		  case "no"
		  rs.open "select product.name,product.id,review.reviewid,review.reviewcontent,review.reviewdtm from review,product where product.id=review.id and review.audit=0",conn,1,1
		  case "yes"
		  		  rs.open "select product.name,product.id,review.reviewid,review.reviewcontent,review.reviewdtm from review,product where product.id=review.id and review.audit=1",conn,1,1
		  end select
				if err.number<>0 then
				response.write "���ݿ���������"
				end if
				if rs.eof And rs.bof then
       			Response.Write "<p align='center' class='contents'> Ŀǰ��û���κ����ۣ�</p>"
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
            			showpage totalput,MaxPerPage,"admincomment.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"admincomment.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"admincomment.asp"
	      				end if
	   				end if
   				   				end if

   				sub showContent
       			dim i
	   			i=0
			response.write "<table width=12 height=7 border=0 cellpadding=0 cellspacing=0><tr><td height=7></td></tr></table>"
			%>
              <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
                <tr bgcolor="#FFFFFF">
                  <td width="32%" align="center"><font color="#FF0000">������Ʒ����</font></td>
                  <td width="26%" align="center"><font color="#FF0000">��������</font></td>
                  <td width="33%" align="center"><font color="#FF0000">����ʱ��</font></td>
                  <td width="9%" align="center"><font color="#FF0000">�� ��</font></td>
                </tr>
                <%  
		 
		  do while not rs.eof%>
                <tr bgcolor="#FFFFFF" align="center">
                  <td>
                    <%
			response.write "<a href=../vpro.asp?id="&rs("id")&" target=_blank title="&trim(rs("name"))&">"&strvalue(trim(rs("name")),18)&"...</a>"
			%>
                  </td>
                  <td>
                    <%
			response.write "<a href=# title="&trim(rs("reviewcontent"))&">"&strvalue(trim(rs("reviewcontent")),20)&"...</a>"
			%>
  
                  </td>
                  <td><%=rs("reviewdtm")%></td>
                  <td>
                    <input name="audit" type="checkbox" id="audit3" value="<%=rs("reviewid")%>">
                  </td>
                </tr>
                <%i=i+1
		  if i>=MaxPerPage then Exit Do
		  rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		  %>
                <tr bgcolor="#FFFFFF">
                  <td height="30" colspan="4" align="center">
                    <%if dtype="no" then%>
                    <input type="submit" name="Submit" value="ͨ�����" onClick="this.form.action='procomm.asp?action=audit';this.form.submit()">
                    <%end if%>
      &nbsp;
                    <input type="button" name="Submit2" value="ɾ ��" onClick="this.form.action='procomm.asp?action=del';this.form.submit()">
&nbsp;&nbsp;ȫѡ
                <input type="checkbox" name="checkbox" value="Check All" onClick="mm()">
                  </td>
                </tr>
              </table>
              <%  
				End Sub   
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				
				Response.Write "<form method=Post action="&filename&"?action="&action&">"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>��ҳ ��һҳ</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&action="&action&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&action="&action&" class='contents'>��һҳ</a> "  
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>��һҳ βҳ</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&action="&action&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&action="&action&" class='contents'>βҳ</a>"  
				End If  
					Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"ҳ</font> "  
					Response.Write "<font class='contents'> ����"&totalnumber&"����¼ " 
					Response.Write "<font class='contents'>" 
					Response.Write "</form>"  
				End Function  
			%>            </td>
            
        </tr>
        </table>
      </form>      <table border="0" align="center" cellpadding="2" cellspacing="0">
        <tr>
          <td><input type="submit" name="Submit42" value="ɾ��һ��ǰδ�������" onClick="if(confirm('��ȷ������������?')) location.href='procomm.asp?action=delzhou';else return;">
            <input type="submit" name="Submit4" value="ɾ������δ�������" onClick="if(confirm('��ȷ������������?')) location.href='procomm.asp?action=delall';else return;"></td>
        </tr>
      </table>      <script language=javascript>
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
      </script>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


