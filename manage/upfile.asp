<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><!--#include file="inc/config.asp"-->
<!--#include file="inc/chk.asp"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>校园网书城</title>
<link href="../style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.bt {font-size: 9pt; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 0px; border-left-width: 0px; height: 16px; width: 80px; background-color: #eeeeee; cursor: hand}
.tx1 {height: 20px; width: 30px; font-size: 9pt; border: 1px solid; border-color: black black #000000; color: #000000}
-->
</style>


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
          <td style="color:#415373">上传文件</td>
        </tr>
      </table>      <form name="form1" method="post" action="supfile.asp" enctype="multipart/form-data" onsubmit="checkImage('file1')">
        <input type="hidden" name="act" value="upload">
        <table width="560" border="0" cellspacing="2" cellpadding="2" align="center" >
          <tr align="center" valign="middle">
            <td height="60" align="center"  >
        <input type="hidden" name="filepath" value="../bookimages/">
        文件：
        <input type="file" name="file1"  value="">
        <script language=javascript>
function checkImage(sId)
{
  if(( document.all[sId].value.indexOf(".gif") == -1) && (document.all[sId].value.indexOf(".jpg") == -1)) {
    alert("请选择gif或jpg的图象文件");
    event.returnValue = false;
    }
}
        </script>
        <input type="submit" name="Submit" value="提交">
小图片尺寸控制在85x125</td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<!--#include file="foot.htm"-->
</body>
</html>


