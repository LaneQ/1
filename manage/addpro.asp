<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"--> 
<!--#include file="inc/chk.asp"--> 
<%
if session("rank")>2 then
	call Msgbox("���Ȩ�޲�����","Back","None")
	response.End
end if
%>

<%

'��Ӳ�Ʒ
If NOT IsEmpty (request("AddProSubmit")) then
	dim productdate,discount
	discount=round(request("price2")/request("price1"),2)
	if request("productdateyear")<>"" then
		productdate=trim(request("productdateyear"))&"��"&trim(request("productdatemonth"))&"��"
	else
		productdate=""
	end if
	
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select * from product",conn,1,3
	rs.AddNew
	
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
	
	rs("solded")=0 
	rs("viewnum")=0 
	rs("adddate")=now() 
	rs("rank")=0  
	rs("ranknum")=0
	if request("detail")<>"" then
		rs("detail")=htmlencode2(request("detail"))
	end if

	if request("content")<>"" then
		rs("content")=htmlencode2(request("content"))
	end if

	if request("detail")<>"" then
		rs("desc")=htmlencode2(strvalue(request("detail") ,100))
	end if

	'�Ƿ��Ƽ���Ʒ
	if request("recommend")=1 then  
		rs("recommend")=1
	else
		rs("recommend")=0
	end if
	rs.Update
	rs.Close
	set rs=nothing
	call MsgBox("��ӳɹ���","GoUrl","addpro.asp")
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

        <br></td><td width="561" align="left" valign="top">      <script language = "JavaScript">

var onecount;
onecount=0;
subcat = new Array();
<%
'��ȡ�����ֶθ���JS����
dim count
set rs=server.createobject("adodb.recordset")
rs.open "select * from sorts order by sortsorder ",conn,1,1
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
      </script>      <br>      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="18"><img src="../images/w.gif" width="18" height="18"></td>
          <td width="66" style="color:#415373">�����Ʒ</td>
        </tr>
      </table>      <br>      <form action="" method="post" name="myform" id="myform">
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
          <tr>
            <td> <span class="redfont">*</span>ѡ����ࣺ</td>
            <td colspan="2">
<%
rs.open "select * from category order by categoryorder",conn,1,1
if rs.eof and rs.bof then
	call MsgBox("������ӷ���!","Back","None")
	response.end
else
%>
              <select name="categoryid" size="1" id="select2" onChange="changelocation(document.myform.categoryid.options[document.myform.categoryid.selectedIndex].value)">
                <option selected value="<%=rs("categoryid")%>"><%=trim(rs("category"))%></option>
<%      
 dim selclass
 selclass=rs("categoryid")
 rs.movenext
 do while not rs.eof
%>
                <option value="<%=rs("categoryid")%>"><%=trim(rs("category"))%></option>
                <%
 rs.movenext
 loop
end if
rs.close
%>
              </select>
      С�ࣺ
      <select name="sortsid">
<%
rs.open "select * from sorts where categoryid="&selclass ,conn,1,1
if not(rs.eof and rs.bof) then
%>
        <option value="<%=rs("sortsid")%>" selected><%=rs("sorts")%></option>
        <% rs.movenext
do while not rs.eof
%>
        <option value="<%=rs("sortsid")%>"><%=rs("sorts")%></option>
<%
rs.movenext
loop
end if
        rs.close
        set rs = nothing
        conn.Close
        set conn = nothing
%>
      </select>
            </td>
          </tr>
          <tr>
            <td><span class="redfont">*</span>������</td>
            <td colspan="2">
              <input name="name" type="text" id="name" size="30">
            </td>
          </tr>
          <tr>
            <td><span class="redfont">*</span>���ߣ�</td>
            <td colspan="2"><input name="author" type="text" id="author" size="20"></td>
          </tr>
          <tr>
            <td><span class="redfont">*</span>�� �� �磺              </td>
            <td colspan="2">
              <input name="mark" type="text" id="mark" size="30" ></td>
          </tr>
          <tr>
            <td>װ֡��</td>
            <td colspan="2">
              <input name="introduce" type="text" id="introduce" size="30" ></td>
          </tr>
          <tr>
            <td><span class="redfont">*</span>�������ڣ�</td>
            <td colspan="2">
              <select name="productdateyear" id="productdateyear">
                <%dim i
for i=year(now) to 1900 step -1
response.write "<option value="&i&">"&i&"</option>"
next
%>
              </select>
              ��
              <select name="productdatemonth" id="productdatemonth">
                <%for i=1 to 12
response.write "<option value="&i&">"&i&"</option>"
next%>
              </select>
      ��</td>
          </tr>
          <tr>
            <td><span class="redfont">*</span>�۸�              </td>
            <td colspan="2">�г��ۣ�
                <input name="price1" type="text" id="price1" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" VALUE="0" size="6" 
onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
                Ԫ�� ��Ա�ۣ�
                <input name="price2" type="text" id="price2" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" value="0" size="6" 
onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
                Ԫ�� VIP�ۣ�
                <INPUT NAME="vipprice" TYPE="text" ID="vipprice" ONKEYPRESS	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" value="0" SIZE="6" 
ONPASTE		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
ONDROP		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
                Ԫ<br> 
              ���ͻ��֣�
              <INPUT NAME="score" VALUE="0" SIZE="4" TYPE="text" ONKEYPRESS	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
ONPASTE		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
ONDROP		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">
                ��</td>
          </tr>
          
          <tr>
            <td>������</td>
            <td colspan="2"><input name="format" type="text" id="format" size="10"></td>
          </tr>
          <tr>
            <td>��Σ�</td>
            <td colspan="2">            <input name="printed" type="text" id="printed" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" size="6" 
onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))"></td>
          </tr>
          <tr>
            <td>ҳ����              </td>
            <td colspan="2">
              <input name="pagenum" type="text" id="pagenum" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" size="10" 
onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))"></td>
          </tr>
          <tr>
            <td><span class="redfont">*</span>ISBN��</td>
            <td colspan="2"><input name="type" type="text" id="type2" size="30"></td>
          </tr>
          <tr>
            <td><span class="redfont">*</span>��ƷͼƬ��              </td>
            <td colspan="2">
              <input name="pic" type="text" id="pic" size="30" VALUE="bookimages/emptybook.gif">
&nbsp;
      <input type="button" name="Submit2" value="�� ��" onClick="javascript:window.open('upfile.asp','','width=580,height=160,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=no');return false;"></td>
          </tr>
          <tr>
            <td>��ϸ˵����</td>
            <td colspan="2">
              <textarea name="detail" cols="46" rows="8" id="detail"></textarea>
            </td>
          </tr>
          <tr>
            <td valign="top">Ŀ¼��</td>
            <td colspan="2"><TEXTAREA NAME="content" COLS="46" ROWS="8" ID="content"></TEXTAREA>
            </td>
          </tr>
          <tr align="center">
            <td colspan="3">
                <input name="recommend" type="checkbox" id="recommend" value="1">
        �Ƽ�����Ʒ
        <input name="AddProSubmit" type="submit" id="AddProSubmit" onClick="return checkpro();" value="���">

<script language="JavaScript">
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
                          
</script></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


