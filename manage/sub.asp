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
dim categoryid,category,follows

dim action
if NOT isempty("action") then
dim url,i,abc
categoryid=request("categoryid")
category=request.QueryString("category")
action=request.QueryString("action")

select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sorts",conn,1,3
rs.AddNew
rs("sorts")=trim(request("sorts2"))
rs("sortsorder")=int(request("sortsorder2"))
rs("categoryid")=int(request("categoryid"))
rs("first")=int(request("first2"))
rs.Update
rs.Close
set rs=nothing
response.redirect "sub.asp?id="&categoryid&"&category="&category


case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sorts where sortsid="&request.QueryString("id"),conn,1,3
rs("sorts")=trim(request("sorts"))
rs("sortsorder")=int(request("sortsorder"))
rs("first")=int(request("first"))
rs.update
rs.close
set rs=nothing
response.redirect "sub.asp?id="&categoryid&"&category="&category

case "del"
categoryid=request.QueryString("categoryid")
conn.execute ("delete from sorts where sortsid="&request.QueryString("id"))
conn.execute ("delete from product where sortsid="&request.QueryString("id"))
response.redirect "sub.asp?id="&categoryid&"&category="&category

end select


end if

category=request.QueryString("sorts")
categoryid=request.QueryString("id")

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
          <td style="color:#415373">��ƷС�����</td>
        </tr>
      </table>      <br>      <select name="select" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
        <base target=Right>
        <option selected>ѡ����Ʒ����</option>
        <%set rs=server.createobject("adodb.recordset")
		rs.Open "select * from category order by categoryorder",conn,1,1
		do while not rs.eof %>
        <option value="sub.asp?id=<%=rs("categoryid")%>&category=<%=rs("category")%>" ><%=trim(rs("category"))%></option>
        <%rs.movenext
		loop
		rs.close
		set rs=nothing
		%>
      </select>      <%if request.QueryString("id")<>"" then
        response.Write "��ǰ��ѯ��"&request.QueryString("category")
        end if%>      <br>      <table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">
        <tr align="center" height="20">
          <td width="40%">��������</td>
          <td width="20%">��������</td>
          <td width="20%">һ������</td>
          <td width="20%">ȷ������</td>
        </tr>
        <%
        if categoryid="" then
        response.Write "<div align=center><font color=red>��ѡ�����ķ���</font></div>"
        else
        set rs=server.CreateObject("adodb.recordset")
        rs.Open "select * from sorts where categoryid="&categoryid&" order by sortsorder",conn,1,1
         if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>��û�з���</font></center>"
		  follows=0
		  else
         do while not rs.EOF
         %>
        <form name="form1" method="post" action="sub.asp?action=edit&id=<%=rs("sortsid")%>&category=<%=request.QueryString("category")%>">
          <tr align="center">
            <td><input name="sorts" type="text" id="sorts3" size="16" value="<%=trim(rs("sorts"))%>">
                <input name="categoryid" type="hidden" value="<%=request.QueryString("id")%>" id="categoryid"></td>
            <td><input name="sortsorder" type="text" id="sortsorder3" size="4" value="<%=int(rs("sortsorder"))%>"></td>
            <td><input name="first" type="checkbox" id="first22" value="1">
                <%if rs("first")=1 then
                response.Write "<font color=red>һ��</font>"
                else
                response.Write "����"
                end if%>
            </td>
            <td><input type="submit" name="Submit" value="�޸�">
&nbsp;<a href="sub.asp?id=<%=int(rs("sortsid"))%>&action=del&categoryid=<%=request.QueryString("id")%>&category=<%=request.QueryString("category")%>" onClick="return confirm('��ȷ������ɾ��������')"><font color=red>ɾ��</font></a> </td>
          </tr>
        </form>
        <%rs.movenext
        loop
        follows=rs.RecordCount
        rs.close
        set rs=nothing
        end if
        end if
				%>
      </table>      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">�����Ʒ����</td>
        </tr>
      </table>
      ��ǰ���ࣺ<%=request.QueryString("category")%> <br>      <table width="100%" align="center"border="0" cellpadding="2" cellspacing="1">
        <tr align="center" height="20">
          <td width="40%">��������</td>
          <td width="20%">��������</td>
          <td width="20%">һ������</td>
          <td width="20%">ȷ������</td>
        </tr>
        <form name="form2" method="post" action="sub.asp?action=add&category=<%=request.QueryString("category")%>">
          <tr align="center">
            <td><input name="sorts2" type="text" id="sorts22" size="16"></td>
            <td><input name="sortsorder2" type="text" id="sortsorder22" size="4" value="<%=follows+1%>">
                <input name="categoryid" type="hidden" value="<%=request.QueryString("id")%>"></td>
            <td><input name="first2" type="checkbox" id="first2" value="1">
                <font color="#FF0000">����</font></td>
            <td><input type="submit" name="Submit2" value="���"></td>
          </tr>
        </form>
      </table>      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">����ע�����˵��</td>
        </tr>
      </table>      <table width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
        <tr>
          <td height="16"><font color="#FF0000">����ע��������Ʋ�Ҫ���зǷ��ַ���<br>
      ������һ������ʱ�����û��ѡ���С����������Ϊһ�����࣬����Ϊ&quot;����&quot;���ࡣ<br>
      ������ɾ��������ͬʱ����ɾ���˷����µ�������Ʒ��</font></td>
        </tr>
      </table>      <br>      <br>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


