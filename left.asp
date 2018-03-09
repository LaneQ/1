<%
dim ii,rsLeft
set rsLeft=server.CreateObject("adodb.recordset")


%>
            <script  language="javascript">
function Menu(term){
   if(term.style.display=="none"){
	  term.style.display="";
    }else{
	term.style.display="none";}

}

            </script>
<table width="212" border="0" align="left" cellpadding="0" cellspacing="0">
        <tr>
          <td align="left" valign="top"><img src="images/mycar_up_1.gif" width="212" height="47"></td>
        </tr>
        <tr>
          <td align="left" valign="top"><img src="images/mycar_up_2.gif" width="212" height="13"></td>
        </tr>
        <tr>
          <td height="188" align="center" valign="top" background="images/mycar_bg.gif" ><iframe name="pcart" src="icar.asp" width="185" height="188" scrolling="no" frameborder="0"></iframe></td>
        </tr>
        <tr>
          <td width="212" height="36" align="center" background="images/mycar_down.gif"><table width="161" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="26">&nbsp;</td>
                <td width="135"><a href="car.asp" target="_blank" class="whitefont">购物车/结帐</a>|<a href="muser.asp" target="_parent" class="whitefont">帐号</a>|<a href="logout.asp" target="_parent" class="whitefont">注消</a></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><img src="images/lmdh.gif" width="212" height="23"></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <%
		rsLeft.open "select category,categoryid from category where first=1",conn,1,1
			do while not rsLeft.eof
		%>
		<tr>
          <td height="23" background="images/lmenu_2.gif"  class="dao" onClick=Menu(<%="m"&rsLeft("categoryid")%>)><%=rsLeft("category")%></td>
        </tr>
        
		<tr>
          <td height="23" id=<%="m"&rsLeft("categoryid")%> style="display:none">
		  <table width="100%"  border="0" cellspacing="2" cellpadding="2">
            <%
		  	dim rsSubLeft
			set rsSubLeft=server.CreateObject("adodb.recordset")
			rsSubLeft.open "select sorts,sortsid from sorts where categoryid="&rsLeft("categoryid")&" and first=1 order by sortsorder",conn,1,1
			do while not rsSubLeft.eof
			%>
			<tr>
              <td class="daosub"><a href="sub.asp?aid=<%=rsLeft("categoryid")%>&nid=<%=rsSubLeft("sortsid")%>"><%=rsSubLeft("sorts")%></a></td>
            </tr>
            <%
			rsSubLeft.movenext
			loop

			rsSubLeft.close
			set rsSubLeft=nothing
			%>
		  </table></td>
        </tr>
        
		<%
		rsLeft.movenext
		loop
		
		rsLeft.close
		%>
        <tr>
          <td><img src="images/cxtsph.gif" width="212" height="32"></td>
        </tr>
        <tr>
          <td height="100" align="center" valign="top">
		  <table width="100%" border="0" cellspacing="2" cellpadding="1">
		  		  <%
		  rsLeft.open "select top 10 id,name,solded from product order by solded desc",conn,1,1
		  if rsLeft.eof and rsLeft.bof then
			  response.write "　　对不起， 暂时还没有商品！"
			  'response.End()
			  else
        		ii=0
			  do while not rsLeft.eof
			  %>

              <tr>
                <td align="left"><a href="vpro.asp?id=<%=rsLeft("id")%>" target="_blank"><img src="images/w.gif" width="18" height="18" border="0"><%=strvalue(trim(rsLeft("name")),20)%></a>(<%=rsLeft("solded")%>)</td>
              </tr>
			  <%ii=ii+1
			  if ii>=10 then exit do
			  rsLeft.movenext
			  loop
			  rsLeft.close
			  set rsLeft=nothing
			  end if
			  %>
          </table></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>


