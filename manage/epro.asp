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
dim id
id=request.QueryString("id")


If NOT IsEmpty (request("SaveEditSubmit")) then
	dim productdate,discount

	discount=round(request("price2")/request("price1"),2)
	
	if request("productdateyear")<>"" then
		productdate=trim(request("productdateyear"))&"��"&trim(request("productdatemonth"))&"��"
	else	
		productdate=""
	end if
	set rs=server.CreateObject("adodb.recordset")

	rs.Open "select * from product where id="&id,conn,1,3
	
	rs("productdate")=productdate 
	rs("discount")=discount 
	
	rs("name")=trim(request("name")) 
	
	
	rs("format")=request("format")
	
	if request("pagenum")<>"" then
		rs("pagenum")=int(request("pagenum"))
	else
		rs("pagenum")=0
	end if
	
	if request("printed")<>"" then
		rs("printed")=int(request("printed"))
	else
		rs("printed")=0
	end if

	
	
	rs("author")=trim(request("author"))
	
	rs("mark")=trim(request("mark"))
	rs("introduce")=trim(request("introduce")) 
	
	rs("price1")=trim(request("price1"))  
	rs("price2")=trim(request("price2"))  
	rs("vipprice")=trim(request("vipprice"))  
	
	rs("pic")=trim(request("pic")) 
	rs("categoryid")=int(request("categoryid"))
	rs("sortsid")=int(request("sortsid")) 
	
	rs("score")=request("score") 
	
	rs("type")=trim(request("type"))
	
	
	if request("detail")<>"" then
		rs("detail")=htmlencode2(request("detail"))
	end if

	if request("content")<>"" then
		rs("content")=htmlencode2(request("content"))
	end if

	if request("detail")<>"" then
		rs("desc")=htmlencode2(strvalue(request("detail") ,100))
	end if



	if request("recommend")=1 then  
		rs("recommend")=1
	else
		rs("recommend")=0
	end if
	rs.Update
	rs.Close
	set rs=nothing
	call MsgBox("�޸ĳɹ���","GoUrl","mpro.asp")
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

        <br></td><td width="561" align="left" valign="top">      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="../images/w.gif" width="18" height="18"></td>
          <td style="color:#415373">�༭��Ʒ����</td>
        </tr>
      </table>      <%
dim count
set rs=server.createobject("adodb.recordset")
rs.open "select * from sorts order by sortsorder ",conn,1,1
%>      <script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
<%
   count = 0
   do while not rs.eof 
%>
subcat[<%=count%>] = new Array("<%= trim(rs("sorts"))%>","<%= rs("categoryid")%>","<%= rs("sortsid")%>");
<%
        count = count + 1
        rs.movenext
        loop
        rs.close
%>
		
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.sortsid.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
             document.myform.sortsid.options[document.myform.sortsid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
      </script>      <br>      <form action="" method="post" name="myform" id="myform">
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
          <tr>
            <td width="21%" align="right"><span class="redfont">*</span>ѡ����ࣺ</td>
            <td colspan="2">
              <%
rs.open "select * from category order by categoryorder",conn,1,1
if rs.eof and rs.bof then
	response.write "���������Ŀ��"
else
	dim rs1
	set rs1=server.CreateObject("adodb.recordset")
	rs1.open "select * from product where id="&id,conn,1,1
%>
              <select name="categoryid" size="1" id="categoryid" onChange="changelocation(document.myform.categoryid.options[document.myform.categoryid.selectedIndex].value)">
                <option selected value="<%=rs("categoryid")%>"><%=trim(rs("category"))%></option>
                <%      dim selclass
         selclass=rs1("categoryid")
        rs.movenext
        do while not rs.eof

  response.write  "<option value="&rs("categoryid")
  if rs1("categoryid")=rs("categoryid") then response.write " selected "
  response.write ">"&trim(rs("category"))&"</option>"    
        rs.movenext
        loop
	end if
        rs.close
%>
              </select>
      С�ࣺ
      <select name="sortsid">
        <%rs.open "select * from sorts where categoryid="&selclass ,conn,1,1
if not(rs.eof and rs.bof) then
%>
        <option selected value="<%=rs("sortsid")%>"><%=rs("sorts")%></option>
        <% rs.movenext
do while not rs.eof
  response.write  "<option value="&rs("sortsid")
  if rs1("sortsid")=rs("sortsid") then response.write " selected "
  response.write ">"&trim(rs("sorts"))&"</option>"  
rs.movenext
loop
end if
        rs.close
        set rs = nothing
        
%>
      </select>
      <font color="#FF0000">&nbsp; </font></td>
          </tr>
          <tr>
            <td align="right"><span class="redfont">*</span>������ </td>
            <td colspan="2"><input name="name" type="text" id="name2" size="30" value="<%=trim(rs1("name"))%>"></td>
          </tr>
          <tr>
            <td align="right"><span class="redfont">*</span>���ߣ�</td>
            <td colspan="2"><input name="author" type="text" id="mark" size="30" value="<%=trim(rs1("author"))%>"></td>
          </tr>
          <tr>
            <td align="right"><span class="redfont">*</span>�� �� �磺</td>
            <td colspan="2"><input name="mark" type="text" id="mark2" size="30" value="<%=trim(rs1("mark"))%>"></td>
          </tr>
          <tr>
            <td align="right">װ֡�� </td>
            <td colspan="2"><input name="introduce" type="text" id="introduce2" size="30" value="<%=trim(rs1("introduce"))%>"></td>
          </tr>
          <tr>
            <td height="20" align="right"><span class="redfont">*</span>�������ڣ�</td>
            <td colspan="2"><select name="productdateyear" id="select">
                <%dim i
				for i=1996 to 2006 
				  response.write  "<option value="&i
  if int(left(rs1("productdate"),4))=i then response.write " selected "
  response.write ">"&i&"</option>" 
				next
			%>
              </select>
      ��
      <select name="productdatemonth" id="select2">
        <%dim b
				b=array("01","02","03","04","05","06","07","08","09","10","11","12")
				for i=1 to 12
				response.write  "<option value="&trim(b(i-1))				
  if  left(right(rs1("productdate"),3),2)=trim(b(i-1)) then response.write " selected "
  response.write ">"&i&"</option>"
				next%>
      </select>
      ��</td>
          </tr>
          <tr>
            <td align="right" valign="top"><span class="redfont">*</span>�۸� </td>
            <td colspan="2">�г��ۣ�
                <input name="price1" type="text" id="price12" size="5" onkeypress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))" value="<%=rs1("price1")%>">
      Ԫ����ͨ��Ա��
      <input name="price2" type="text" id="price22" size="5" onkeypress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))" value="<%=rs1("price2")%>">
      Ԫ, VIP��Ա�ۣ�
      <input name="vipprice" type="text" id="vipprice2" size="5" onkeypress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))" value="<%=rs1("vipprice")%>">
      Ԫ<br>
      ���ͻ��֣�
      <INPUT NAME="score" VALUE="<%=trim(rs1("score"))%>" SIZE="4" TYPE="text" ONKEYPRESS	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		ONPASTE		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ONDROP		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">
      ��</td>
          </tr>
          
          <tr>
            <td height="19" align="right">������</td>
            <td colspan="2"><input name="format" type="text" id="grade2" size="10" value="<%=trim(rs1("format"))%>"></td>
          </tr>
          <tr>
            <td height="19" align="right">��Σ�</td>
            <td colspan="2"><input name="printed" type="text" id="printed" value="<%=trim(rs1("printed"))%>" size="10"></td>
          </tr>
          <tr>
            <td height="19" align="right">ҳ���� </td>
            <td colspan="2"><input name="pagenum" type="text" id="pagenum" value="<%=trim(rs1("pagenum"))%>" size="10"></td>
          </tr>
          <tr>
            <td height="19" align="right"><span class="redfont">*</span>ISBN��</td>
            <td colspan="2"><input name="type" type="text" id="type" value="<%=trim(rs1("type"))%>" size="30">
            </td>
          </tr>
          <tr>
            <td align="right"><span class="redfont">*</span>��ƷͼƬ��</td>
            <td colspan="2"><input name="pic" type="text" id="pic2" size="30" value="<%=trim(rs1("pic"))%>">
&nbsp;&nbsp;
      <input type="button" name="Submit2" value="�� ��" onClick="javascript:window.open('upfile.asp','','width=580,height=160,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=no');return false;"></td>
          </tr>
          <tr>
            <td align="right" valign="top">��ϸ˵���� </td>
            <td colspan="2"><textarea name="detail" cols="46" rows="8" id="textarea"><% if rs1("detail")<>"" then response.write HTMLDecode(trim(rs1("detail")))%></textarea></td>
          </tr>
          <tr>
            <td align="right" valign="top">Ŀ¼��</td>
            <td colspan="2"><textarea name="content" cols="46" rows="8" id="textarea2"><%if rs1("content")<>"" then response.write HTMLDecode(trim(rs1("content")))%></textarea></td>
          </tr>
          <tr>
            <td height="32" colspan="3"><div align="center"> </div>              <div align="center">
                <input name="recommend" type="checkbox" id="recommend2" value="1" <%if rs1("recommend")=1 then response.write "checked"%>>
        �Ƽ�����Ʒ
        <input type="submit" name="SaveEditSubmit" value="�޸�" onClick="return check();">
              <input name="id" type="hidden" id="id" value="<%=id%>">
            </div></td>
          </tr>
        </table>
      </form>      <%rs1.close
set rs1=nothing
conn.close
set conn=nothing%>      <script language="JavaScript">
<!--
function checkpro()
{
    if(checkspace(document.myform.name.value)) {
	document.myform.name.focus();
    alert("������������");
	return false;
  }

	if(checkspace(document.myform.author.value)) {
	document.myform.author.focus();
    alert("���������ߣ�");
	return false;
  }
	if(checkspace(document.myform.mark.value)) {
	document.myform.mark.focus();
    alert("����������磡");
	return false;
  }

	if(checkspace(document.myform.type.value)) {
	document.myform.type.focus();
    alert("������ISBN��");
	return false;
  }

    if(checkspace(document.myform.price1.value)||document.myform.price1.value==0) {
	document.myform.price1.focus();
    alert("�������г��ۣ�");
	return false;
  }
    if(checkspace(document.myform.price2.value)||document.myform.price2.value==0) {
	document.myform.price2.focus();
    alert("�������Ա�ۣ�");
	return false;
  }
    if(checkspace(document.myform.vipprice.value)||document.myform.vipprice.value==0) {
	document.myform.vipprice.focus();
    alert("������VIP�ۣ�");
	return false;
  }


     if(checkspace(document.myform.price1.value)) {
	document.myform.price1.focus();
    alert("��������Ʒ�г��۸�");
	return false;
  }
     if(checkspace(document.myform.price2.value)) {
	document.myform.price2.focus();
    alert("��������Ʒ��Ա�۸�");
	return false;
  }
      if(checkspace(document.myform.vipprice.value)) {
	document.myform.vipprice.focus();
    alert("������VIP�»�Ա��Ʒ�۸�");
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
                          
      </script>      <br>      <br>
      </td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


