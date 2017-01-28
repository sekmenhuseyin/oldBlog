<!-- #Include File ="tasarim.asp" --><%top:Set Rs=ObjCon.execute("SELECT Year([tarih]) AS Expr1 FROM Blog where gorunur=true ORDER BY Year([tarih]);")
if not Rs.eof or not Rs.bof then
TheFirstYear=Rs(0):rs.close
response.write "<div id=""arsiv"" class=""Mesaj""><br /><ul>"
for y=year(Date()) to TheFirstYear step -1
	response.write "<li><b>"&y&"</b><ul>"
	for a=12 to 1 step -1
		Set Rs=ObjCon.execute("Select * From blog where gorunur=true and year(tarih)="&y&" and month(tarih)="&a&" ORDER BY Year([tarih]);")
		Do While Not Rs.eof
			Response.Write "<li type=""circle""><b>"&day(Rs("tarih"))&" "&MonthName(a)&":</b> <a href="""&SiteAdres&"/blog/"&Rs("id")&"/"&Rs("baslik")&".html"" class=""l_b"">"&Rs("baslik")&"</a></li>"
			Rs.MoveNext
		loop
	next
	response.Write("</ul></li>")
next:Response.Write "</ul></div>"
end if
Rs.close
bottom%>