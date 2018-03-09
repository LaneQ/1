<%
'打开数据库连接
dim conn,connstr,db,rs
db="db/db.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")

'如果你的服务器采用较老版本Access驱动，请用下面连接方法
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(db)

conn.Open connstr
%>



