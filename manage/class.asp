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
dim action,categoryid
categoryid=request.QueryString("id")
action=request.querystring("action")
select case action

case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from category",conn,1,3
rs.AddNew
rs("category")=trim(request("category2"))
rs("categoryorder")=int(request("categoryorder2"))
rs("first")=int(request("first2"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "class.asp"

case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from category where categoryid="&categoryid,conn,1,3
rs("category")=trim(request("category"))
rs("categoryorder")=int(request("categoryorder"))

rs("first")=int(request("first"))

rs.Update
rs.Close
set rs=nothing
response.Redirect "class.asp"

case "del"
conn.execute ("delete from category where categoryid="&categoryid)
conn.execute ("delete from sorts where categoryid="&categoryid)
conn.execute ("delete from product where categoryid="&categoryid)
response.Redirect "class.asp"
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
          <td style="color:#415373">��Ʒ�������</td>
        </tr>
      </table>      <br>      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
        <tr align="center" bgcolor="#FFFFFF" height="20">
          <td width="30%">��������</td>
          <td width="20%"> ��������</td>
          <td width="20%">һ������</td>
          <td width="30%">ȷ������</td>
        </tr>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from category order by categoryorder",conn,1,1
		  dim follows
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>��û�з���</font></center>"
		  follows=0
		  else
		  do while not rs.EOF
		  %>
        <form name="form2" method="post" action="class.asp?action=edit&id=<%=int(rs("categoryid"))%>">
          <tr bgcolor="#FFFFFF" align="center">
            <td><input name="category" type="text" id="category3" size="12" value="<%=trim(rs("category"))%>"></td>
            <td><input name="categoryorder" type="text" id="categoryorder" size="4" value="<%=int(rs("categoryorder"))%>"></td>
            <td><input name="first" type=checkbox value="1">
                <%if rs("first")=1 then
                response.Write "<font color=red>һ��</font>"
                else
                response.Write "����"
                end if%>
            </td>
            <td><input type="submit" name="Submit" value="�޸�">
&nbsp; <a href="class.asp?id=<%=int(rs("categoryid"))%>&action=del" onClick="return confirm('��ȷ������ɾ��������')"><font color=red>ɾ��</font></a> </td>
          </tr>
        </form>
        <%rs.MoveNext
          loop
          follows=rs.RecordCount
          end if%>
      </table>      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">�����Ʒ����</td>
        </tr>
      </table>      <font color="#FFFFFF">����ע�����˵��</font><br>      <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr align="center" bgcolor="#FFFFFF" height="20">
          <td width="30%"> ��������</td>
          <td width="20%">��������</td>
          <td width="20%">һ������</td>
          <td width="30%">ȷ������</td>
        </tr>
        <form name="form1" method="post" action="class.asp?action=add">
          <tr align="center" bgcolor="#FFFFFF">
            <td><input name="category2" type="text" id="category22" size="12"></td>
            <td><input name="categoryorder2" type="text" id="categoryorder22" size="4" value="<%=follows+1%>"></td>
            <td><input name="first2" type="checkbox" id="first22" value="1"></td>
            <td><input type="submit" name="Submit3" value="���" onClick="return checkpro()">
            <script language="JavaScript">
<!--
function checkpro()
{
    if(checkspace(document.form1.category2.value)) {
	document.form1.category2.focus();
    alert("�������������");
	return false;
  }

    if(checkspace(document.form1.categoryorder2.value)) {
	document.form1.categoryorder2.focus();
    alert("�������������");
	return false;
  }

 
}
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}

function regInput(obj, reg, inputStr)
	{
		var docSel	= document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")	return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange	= obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
	}
//-->
                          
              
</script></td>
          </tr>
        </form>
      </table>      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">����ע�����˵��</td>
        </tr>
      </table>      <br>      <table width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
        <tr>
          <td height="20"><font color="#FF0000">����ע��������Ʋ�Ҫ���зǷ��ַ���<br>
      ������һ������󣬴˷��ཫ�ᱻ�г�����ҳ����Ŀ����<br>
      ������ɾ��������ͬʱ����ɾ���˴����°���������С�������Ʒ��</font></td>
        </tr>
      </table>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


