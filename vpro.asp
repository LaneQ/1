<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="manage/inc/config.asp"--> 
<!--#include file="inc/conn.asp"--> 
<%
dim id
id=request.QueryString("id")
dim prename,company,intro,predate,graph2,description,remarks,price
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from product where id="&id,conn,1,3
rs("viewnum")=rs("viewnum")+1
rs.update

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
    <td width="561" align="left" valign="top">
      <br>      <table width="568"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="17%" height="130" align="center" valign="middle" class="shadow"><img src="<%=trim(rs("pic"))%>" width="85" height="125"></td>
          <td height="130" align="left" valign="top"><table width="100%"  border="0" cellspacing="2" cellpadding="0">
            <tr>
              <td colspan="2"><table border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td><img src="images/w.gif" width="18" height="18"></td>
                    <td><span class="booktitle"><%=trim(rs("name"))%></span></td>
                  </tr>
              </table></td>
            </tr>
            <tr class="bookinfo">
              <td width="50%" class="bookinfo">�����ߣ�<%=trim(rs("author"))%></td>
              <td width="50%" class="bookinfo"> ISBN ��<%=trim(rs("type"))%></td>
            </tr>
            <tr class="bookinfo">
              <td width="50%"> �����磺<%=trim(rs("mark"))%></td>
              <td width="50%"> ��������<%=trim(rs("format"))%> </td>
            </tr>
            <tr class="bookinfo">
              <td>�������ڣ�<%=trim(rs("productdate"))%></td>
              <td> ҳ������<%=trim(rs("pagenum"))%> </td>
            </tr>
            <tr class="bookinfo">
              <td> װ��֡��<%=trim(rs("introduce"))%> </td>
              <td> �桡�Σ�<%=trim(rs("printed"))%> </td>
            </tr>
            <tr class="bookinfo">
              <td>�����ۣ�<%=trim(rs("price1"))%> </td>
              <td>�Żݼۣ�<%=trim(rs("price2"))%></td>
            </tr>
            <tr class="bookinfo">
              <td>�����֣�<%=rs("score")%></td>
              <td>VIP�۸�<%=rs("vipprice")%></td>
            </tr>
            <tr class="bookinfo">
              <td>䯡�����<%=trim(rs("viewnum"))%></td>
              <td>����<%=trim(rs("solded"))%></td>
            </tr>
            <tr>
              <td colspan="2" align="center"><a href="icar.asp?id=<%=rs("id")%>&action=add" target="pcart"><img src="images/car.gif" width="23" height="20" border="0">���ﳵ</a></td>
            </tr>
          </table></td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:10px;">            <table border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td><img src="images/w.gif" width="18" height="18"></td>
                <td style="padding-left:0px;"><strong><%=trim(rs("name"))%></strong></td>
              </tr>
            </table></td>
        </tr>
        <tr align="left">
          <td colspan="2" style="padding-left:40px;"> <%=trim(rs("detail"))%></td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:10px;"><table border="0" cellspacing="0" cellpadding="2">
            <tr>
              <td><img src="images/w.gif" width="18" height="18"></td>
              <td> <strong>Ŀ¼ </strong> </td>
            </tr>
          </table></td>
        </tr>
        <tr align="left">
          <td colspan="2" style="padding-left:40px;"><%=trim(rs("content"))%></td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:10px;"><table border="0" cellpadding="2" cellspacing="0">
            <tr>
              <td><img src="images/w.gif" width="18" height="18"></td>
              <td> <strong>��Ա����</strong> <a href="rank.asp?id=<%=id%>" target="_blank">���������Ȿ�������</a></td>
            </tr>
          </table></td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:40px;"><%
		'�û�����
		
if rs("ranknum")>0 and rs("rank")>0 then
dim other
other=rs("ranknum")\rs("rank")
else
other=0
end if
response.write "<img src=images/rank/"&other&".gif alt=�����Ǽ�>"

		rs.close
%>            </td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:10px;"><table border="0" cellpadding="2" cellspacing="0">
            <tr>
              <td><img src="images/w.gif" width="18" height="18"></td>
              <td> <strong>��Ա����</strong> <a href="comment.asp?id=<%=id%>" target="_blank">���������Ȿ�������</a></td>
            </tr>
          </table> </td>
        </tr>
        <tr align="left">
          <td height="30" colspan="2" style="padding-left:40px;">
		  <%
		rs.open "select * from review where id="&id&" and audit=1 ",conn,1,1
		if rs.eof and rs.bof then
		response.write "������ù�����Ʒ����Ա���Ʒ�����˽⣬��ӭ�������Լ������ۡ��������۽��������ϳ�ǧ������û����������ǽ������Ŀ�������л��<br>"
		response.write "�����������ύ�󽫾������ǵ���ˣ�Ҳ������Ҫ�ȴ�һЩʱ��ſ��Կ�����лл������"
		else
		do while not rs.eof 
				%>
        [<B><%=rs("title")%></B>@<%=rs("reviewdtm")%>]<BR>
                  <%=rs("reviewcontent")%><br>
 
            <%rs.movenext
		loop
		end if
		rs.close
		%>
		  </td>
        </tr>
        <tr align="left">
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr align="center">
          <td colspan="2"><TABLE WIDTH="96%" BORDER="0" CELLPADDING="0" CELLSPACING="0" align="center">
            <a name="pic"></a>
            <TR>
              <TD align="center"><B>��վ�����û����ۣ���������������ͬ����֧���û��Ĺ۵㡣���ǵ����������ڴ��������û�����Ȥ����Ϣ��</B></TD>
            </TR>
            <TR>
              <TD align="center">&nbsp;</TD>
            </TR>
            <TR>
              <TD align="center"><input type="button" name="Submit" value="�ر�" onClick="window.close()"></TD>
            </TR>
          </TABLE></td>
        </tr>
        <tr align="center">
          <td colspan="2">&nbsp;</td>
        </tr>
    </table></td>
  </tr>
</table>

<!--#include file="foot.htm"-->
</body>
</html>
<%

set rs=nothing

%>

