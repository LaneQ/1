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
'����ύ���ͽ���Recoredset����
If NOT IsEmpty (Request.Form) then 
	set rs=server.CreateObject("adodb.recordset")
	'ȡ�ô���id��
	dim id
	id=request("Id")

end if

'��Ӻ�̨�û�
If NOT IsEmpty (Request("AddHuser")) then
	rs.open "select * from admin",conn,1,3
	rs.addnew
	rs("admin")=trim(request("AddName"))
	rs("password")=md5(trim(request("AddPws")))
	rs("rank")=int(request("AddRank"))
	rs.update
	rs.close
	set rs=nothing
	call MsgBox("��ӳɹ���","GoUrl","huser.asp")
end If

'ɾ����̨�û�
If NOT IsEmpty (request("Del")) then
	'ȡ��Id��
	conn.execute ("delete from admin where id="&id)
	call MsgBox("ɾ���ɹ���","GoUrl","huser.asp")

end If

'�޸ĺ�̨�û�����
if NOT IsEmpty (request("Modify")) then 
	'ȡ��Id��
	rs.Open "select * from admin where id="&id,conn,1,3
	rs("admin")=trim(request("Name"))
	if trim(request("password"))<>"" then
		rs("password")=md5(trim(request("password")))
	end if
	rs("rank")=int(request("rank"))
	rs.Update
	rs.Close
	set rs=nothing
	call MsgBox("�޸ĳɹ���","GoUrl","huser.asp")

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
        <br>        <table border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="../images/w.gif" width="18" height="18"></td>
            <td style="color:#415373">��̨�û�����</td>
          </tr>
        </table>        <br>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="2">
            <tr align="center" bgcolor="#FFFFFF" class="bluefont" height="20">
              <td width="23%">����Ա</td>
              <td width="21%">�� ��</td>
              <td width="31%">Ȩ ��</td>
              <td width="25%">�� ��</td>
            </tr>
            <%set rs=server.CreateObject("adodb.recordset")
        rs.Open "select * from admin order by rank",conn,1,1
        do while not rs.EOF%>
		<form name="form2" method="post" action="">
            <tr align="center" bgcolor="#FFFFFF" height="20">
              <td><input name="Name" type="text" id="Name" value="<%=trim(rs("admin"))%>" size="12"></td>
              <td><input name="Pws" type="text" id="Pws" size="12"></td>
              <td>
                <%select case rs("rank")
                case "1"
                response.Write "<input type=radio name=rank value=1 checked>����&nbsp;<input name=rank type=radio value=2 >���&nbsp;<input type=radio name=rank value=3>�鿴"
                case "2"
                response.Write "<input type=radio name=rank value=1>����&nbsp;<input name=rank type=radio value=2 checked>���&nbsp;<input type=radio name=rank value=3>�鿴"
				case "3"
				response.Write "<input type=radio name=rank value=1>����&nbsp;<input name=rank type=radio value=2>���&nbsp;<input type=radio name=rank value=3  checked>�鿴"
                end select%>
                <input name="Id" type="hidden" id="Id" value="<%=int(rs("id"))%>">
</td>
              <td><input name="Modify" type="submit" id="Modify" value="�޸�">
                  <input name="Del" type="submit" id="Del" value="ɾ��">
              </td>
            </tr>
			</form>
            <%rs.movenext
        loop
        rs.close
        set rs=nothing
        %>
          </table>
          <br>          <br>          <table border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="../images/w.gif" width="18" height="18"></td>
            <td style="color:#415373">��̨�û����</td>
          </tr>
          </table>          <form name="form1" method="post" action="">

		<br>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr align="center" bgcolor="#FFFFFF" class="bluefont" height="20">
            <td width="21%">����Ա</td>
            <td width="23%">�� ��</td>
            <td width="35%">Ȩ ��</td>
            <td width="21%">�� ��</td>
          </tr>
          <tr align="center" bgcolor="#FFFFFF">
            <td><input name="AddName" type="text" id="AddName" size="12"></td>
            <td><input name="AddPws" type="text" id="AddPws" size="12"></td>
            <td><input type="radio" name="AddRank" value="1">
      ����
        <input name="AddRank" type="radio" value="2" checked>
      ���
      <input type="radio" name="AddRank" value="3">
      �鿴</td>
            <td><input name="AddHuser" type="submit" id="AddHuser" value="���"></td>
          </tr>
        </table>
          </form>          <br>          <table width="231" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="24"><img src="../images/w.gif" width="18" height="18"></td>
      <td width="207" style="color:#415373">����ע�����˵��</td>
    </tr>
          </table>          <br>          <table width="80%" border="0" align="center" cellpadding="5" cellspacing="0">
    <tr>
      <td><font color="#FF0000">����̨�����û���ǰ̨�û�����ǣ����<br>
        �������Աֻ����ӡ��޸ġ�ɾ����Ʒ���ϡ�<br>
        ���鿴��Ա���Թ�����Ʒ���ۺ��û�������<br>
        ������Աӵ�б�վ���й���Ȩ�ޡ�<br>
        ����¼�������MD5������ת��ʽ���ܣ��粻�޸����룬�����ա� <br>
        <br>
      </font></td>
    </tr>
          </table></td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


