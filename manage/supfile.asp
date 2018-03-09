<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%


%><%OPTION EXPLICIT%>
<!--#include FILE="inc/upload.inc"-->
<html>
<head>
<title>图片上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<%
dim upload,file,formName,formPath,iCount,sname,fsobj,fsobjd
set upload=new upload_cwj

if upload.form("filepath")="" then
	call MsgBox("请输入要上传至的目录!","Back","None")
	set upload=nothing
	response.end
else
	'以年月为目录名存放图片
	formPath=upload.form("filepath")&year(now)&month(now)&"/"
	'如果目录不存在就建立
	set fsobjd=server.createobject("scripting.filesystemobject")
	if not fsobjd.folderexists(server.mappath(formPath)) then fsobjd.createfolder(server.mappath(formPath))
	set fsobjd=nothing

	if right(formPath,1)<>"/" then formPath=formPath&"/"&year(now)&month(now)&"/" 

end if

iCount=0

for each formName in upload.file 
	set file=upload.file(formName) 
	if file.FileSize>0 then     
		file.SaveAs Server.mappath(formPath&file.FileName) 
		response.write "<br><center><font size=2 color=red>上传成功，请复制下边剪切板中内容而后粘贴到商品图片剪切板内!</font></center><br>"
		
		dim thename,spp,paths
		
		thename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&right(file.filename,4)
		spp=file.filename
		file.filename=thename
		file.SaveAs Server.mappath(formPath&file.FileName)
		paths=server.mappath("../")&"\bookimages\"&year(now)&month(now)&"\"&spp
		set fsobj=server.CreateObject("scripting.filesystemobject")
		if fsobj.fileExists(""&paths&"") then
			fsobj.deletefile(""&paths&"")
		end if

	set fsobj=nothing
	response.write "<center><input type=text size=40 value=bookimages/"&year(now)&month(now)&"/"&file.filename&"><button onclick=window.clipboardData.setData('text',this.previousSibling.value)>复制</button><br><br><a href='javascript:window.close()'><font color=red size=2>关闭窗口</font></a></center>"
	iCount=iCount+1
	
	end if
	set file=nothing
next


set upload=nothing 
response.write "<font color=red size=2>"


%>
</body>
</html>

