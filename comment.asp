<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
dim id
id=request.QueryString("id")
if NOT isempty(request("CommentSubmit")) then
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from review",conn,1,3
	rs.addnew
	rs("id")=id
	rs("title")=HTMLEncode2(trim(request("title")))
	rs("reviewcontent")=HTMLEncode2(trim(request("reviewcontent")))
	rs("reviewdtm")=now()
	rs("audit")=0
	rs.update
	rs.close
	set rs=nothing
	call MsgBox("���������ѳɹ��ύ�����������ͨ���󷽿ɷ�����","Close","None")
	response.End
end if
%> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>У԰�����</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body>
<!--#include file="head.htm"-->


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="219" align="left" valign="top"><!--#include file="left.asp"--></td>
    <td width="561" align="left" valign="top">      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr>
          <td ><DIV ALIGN="CENTER"><FONT COLOR="#FFFFFF" SIZE="4"><B>��������</B></FONT></DIV></td>
        </tr>
        <tr>
          <form name="reviewform" method="post" action="">
            <td>
              <table width="100%" border="0" cellpadding="2" cellspacing="1" >
                <tr >
                  <td width="23%">����������</td>
                  <td width="77%">
                    <input name="title" type="text" id="title" size="12">
                    <input name="id" type="hidden" id="id" value="<%=id%>">
</td>
                </tr>
                <tr >
                  <td valign="top">�������ģ�</td>
                  <td >
                    <textarea name="reviewcontent" cols="26" rows="5" id="reviewcontent"></textarea>
                  </td>
                </tr>
                <tr align="center" >
                  <td colspan="2">                    
                      <input name="CommentSubmit" type="submit" id="CommentSubmit" onClick="return check();" value="�ύ">
&nbsp;
                <input type="reset" name="Submit2"  value="���">
                <script language="javascript">
<!--
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function check()
{
  if(checkspace(document.reviewform.title.value)) {
	document.reviewform.title.focus();
    alert("����д����������");
	return false;
  }
  if(checkspace(document.reviewform.reviewcontent.value)) {
	document.reviewform.reviewcontent.focus();
    alert("����д�������ģ�");
	return false;
  }
	  }
	  //-->
                </script>
</div></td>
                </tr>
            </table></td>
          </form>
        </tr>
    </table></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>


