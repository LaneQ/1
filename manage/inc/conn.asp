<%
'�����ݿ�����
dim conn,connstr,db,rs
db="db/db.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")

'�����ķ��������ý��ϰ汾Access�����������������ӷ���
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(db)

conn.Open connstr
%>



