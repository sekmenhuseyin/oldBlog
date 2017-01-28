<!-- #Include File ="tasarim.asp" --><%top:TheURL=lcase(Request.ServerVariables("QUERY_STRING"))
if instr(TheURL,"proje")>0 then
	TheID=1:TheLink="projelerim"
elseif instr(TheURL,"cv")>0 then
	TheID=2:TheLink="CV"
end if
if Session("Yonetici")=true and instr(TheURL,"/duzenle/")=0 then response.Write("<div style=""text-align:right;""><a href="""&SiteAdres&"/"&TheLink&"/Duzenle/"&TheID&"/"" title=""Duzenle"">Düzenle</a></div>")
if instr(TheURL,"/duzenle/")>0 then
	sql="select icerik from Xtra where id="&TheID:Set ObjRs=Server.CreateObject ("ADODB.RecordSet"):ObjRs.open sql,ObjCon,1,3
	%><h2>Düzenle <%=TheLink%></h2><br /><br /><form method="post" action="<%=SiteAdres%>/islem.asp"><input type="hidden" name="id" value="<%=TheID%>" /><textarea name="icerik" rows="40" cols="20" class="TextArea"><%=ObjRs(0)%></textarea><br /><input name="Kaydet" type="submit" value="Kaydet" /></form><script type="text/javascript" src="<%=SiteAdres%>/Contents/Editor/ckeditor.js"></script><script type="text/javascript">CKEDITOR.replace('icerik');</script><%
else
	Set ObjRs=ObjCon.Execute("select icerik from Xtra where id="&TheID):response.Write(ObjRs(0))
end if:ObjRs.close
bottom%>