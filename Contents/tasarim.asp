<!-- #Include File = "abt-yapilandirma.asp" --><!-- #Include File = "abt-fonksiyonlar.asp" --><%Call WriteWebLog'web stats log
Z1=Request.Cookies("aBlog")("Ekleyen"):Z2=Request.Cookies("aBlog")("Eposta"):Z3=Request.Cookies("aBlog")("Web")
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" lang="TR" xml:lang="TR"><head>
<meta name="Generator" content="Dreamweaver CS2" /><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" /><meta http-equiv="Content-Language" content="tr" /><meta name="Content-Language" content="tr" />
<meta http-equiv="imagetoolbar" content="no" /><meta http-equiv="Distribution" content="global" /><meta http-equiv="Resource-type" content="document" /><meta http-equiv="Expires" content="<%=date()+1%>" />
<meta name="Copyright" content="Copyright C 2008 - <%=year(date())%> Huseyin Sekmenoglu" /><meta http-equiv="Copyright" content="Copyright C 2008 - <%=year(date())%> Huseyin Sekmenoglu" /><meta name="author" content="Huseyin Sekmenoglu" />
<meta name="googlebot" content="index,follow" /><meta name="robots" content="index,follow" /><meta name="robots" content="all" /><meta name="revisit-after" content="1" /><meta name="identifier-URL" content="<%=SiteAdres%>" />
<link rev="made" href="mailto:hus_as@yahoo.com" /><meta http-equiv="Reply-to" content="<%=EPosta%>" /><meta name="description" content="<%=SiteMetaDes%>" /><meta name="keywords" content="<%=SiteMetaKey%>" />
<link rel="icon" type="image/x-icon" href="<%=SiteAdres%>/contents/images/huseyin-sekmenoglu.ico" /><link rel="shortcut icon" href="<%=SiteAdres%>/contents/images/huseyin-sekmenoglu.ico" />
<link rel="alternate" type="application/rss+xml" title="Huseyin Sekmenoglu: Yazilar" href="http://feeds.feedburner.com/HuseyinSekmenoglu" />
<link rel="stylesheet" href="<%=SiteAdres%>/contents/css/screen.css" type="text/css" media="screen, projection" /><link rel="stylesheet" href="<%=SiteAdres%>/contents/css/print.css" type="text/css" media="print" /> 
<!--[if lt IE 8]><link rel="stylesheet" href="<%=SiteAdres%>/contents/css/ie.css" type="text/css" media="screen, projection" /><![endif]--><script type="text/javascript" src="<%=SiteAdres%>/contents/js/script.js"></script><%

sub top'''''''''''''''''''''''''''''''''''''''''''''''''sayfanin ust bolumu
SayfaAdres=Request.ServerVariables("QUERY_STRING"):if SayfaAdres>"" then SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
%><title><%=KonuBaslik%><%=SiteAdi%>&nbsp;-&nbsp;<%=SiteBaslik%></title>
</head><body><%Call NoIE6()%><div id="container"><div id="header"><a href="<%=SiteAdres%>/" title="<%=SiteAdres%>"><img src="<%=SiteAdres%>/contents/images/logo.jpg" alt="<%=SiteAdi%> Logo" width="93" height="120" class="logo" /><h1><%=SiteAdi%></h1><h2><%=SiteBaslik%></h2><img src="<%=SiteAdres%>/contents/images/bg.gif" alt="" width="580" height="90" class="sus" /></a></div>
<div id="navigation"><%Call TopNavigation
if Session("Yonetici")=true then Call YonetimMenusu:Session("Yonetici")=true:session.Timeout=20
Call SearchForm%></div>
<div id="content-container"><div id="content"><div id="content2"><%
end sub

sub bottom'''''''''''''''''''''''''''''''''''''''''''''''''sayfanin sol ve alt bolumu
%></div></div>
<div id="aside"><%'administration module
if Session("Yonetici")=true then
response.Write("<h3>Istatistik</h3><font class=""sayfaTelif"">"):Set ObjRs=ObjCon.Execute("Select COUNT(*) From yorum where onay=false")
if ObjRs(0)>0 then response.Write("<a href="""&SiteAdres&"/TumYorumlar"">Onaylanmamiº Yorum: "&ObjRs(0)&"</a><br />")
ObjRs.close:response.Write("Yeni Mesaj: "):Set ObjRs=ObjCon.Execute("Select COUNT(*) From BanaMesaj where yeni=true"):if ObjRs(0)=0 then response.Write("0") else response.Write("<a href="""&SiteAdres&"/TumMesajlar"">"&ObjRs(0)&" yeni mesaj</a>")
ObjRs.close:response.Write("<br />Bugunku Ziyaretci: "):Set ObjRs=ObjCon.Execute("Select COUNT(*) From WebStats where date=date()"):response.Write(ObjRs(0)):ObjRs.close
response.Write("<br />Dunku Ziyaretci: "):Set ObjRs=ObjCon.Execute("Select COUNT(*) From WebStats where date=date()-1"):response.Write(ObjRs(0)):ObjRs.close
response.Write("<br />Toplam Ziyaretci: "):Set ObjRs=ObjCon.Execute("Select COUNT(*) From WebStats"):response.Write(ObjRs(0)):ObjRs.close:response.Write("</font><div class=""ModulSep""><br /></div>")
else
Set ObjRs=ObjCon.Execute("Select * From siteayar"):ObjRs.close'sirf bir ObjRs tanimlamak icin...
end if
'popular module
response.Write("<h3>En Cok Okunanlar</h3>")
StrSql="select * from blog where okunma>0 and gorunur=true order by okunma desc,baslik asc;":ObjRs.Open StrSql,ObjCon,1,3:ObjRsCount=ObjRs.recordcount:if ObjRsCount>10 then ObjRsCount=10
for i=1 to ObjRsCount:Response.Write "<a href="""&SiteAdres&"/blog/"&ObjRs("id")&"/"&SEO_Olustur(ObjRs("baslik"))&".html"" class=""l_b"">"&ObjRs("baslik")&" <font class=""miniMavi"">("&ObjRs("okunma")&" okunma)</font></a><br />":ObjRs.Movenext:Next:ObjRs.close
'rastgele module
response.Write("<br /><div class=""ModulSep""><br /></div><h3>Bunu da Okuyabilirsin</h3>")
StrSql="select id,baslik from blog where gorunur=true;":ObjRs.Open StrSql,ObjCon,1,3:randomize:MyRandomBlog=Round(Rnd*(ObjRs.recordcount-1))
ObjRs.Move(MyRandomBlog)
Response.Write "<a href="""&SiteAdres&"/blog/"&ObjRs("id")&"/"&SEO_Olustur(ObjRs("baslik"))&".html"" class=""l_b"">"&ObjRs("baslik")&"</a><br />":ObjRs.close
'yorum module
response.Write("<br /><div class=""ModulSep""><br /></div><h3>Son yorumlar</h3>")
StrSql="select blog from yorum where onay=true order by tarih desc;":ObjRs.Open StrSql,ObjCon,1,3:ObjRsCount=ObjRs.recordcount:If ObjRs.Eof or ObjRs.Bof then Response.Write "<i>Henuz yorum yazilmamiº.</i><br />"
if ObjRsCount>10 then ObjRsCount=10
for i=1 to ObjRsCount:Set ObjRs2=ObjCon.Execute("select id,baslik from blog where id="&ObjRs("blog")):Response.Write "<a href="""&SiteAdres&"/blog/"&ObjRs2("id")&"/"&SEO_Olustur(ObjRs2("baslik"))&".html&#35;Mesaj-Yorumlar"" class=""l_b"">"&ObjRs2("baslik")&"</a><br />":ObjRs2.close:ObjRs.Movenext:Next:ObjRs.close
'etiket module
response.Write("<br /><div class=""ModulSep""><br /></div><h3>Etiket Bulutum</h3>")
StrSql="select * from etiket_ad order by etiket asc;":ObjRs.Open StrSql,ObjCon,1,3:ObjRsCount=ObjRs.recordcount
for i=0 to ObjRsCount-1
Set ObjRs2=ObjCon.Execute("SELECT Count(etiket_bulut.etiket_id) AS CountOfetiket_id FROM etiket_bulut INNER JOIN Blog ON etiket_bulut.blog_id=Blog.id WHERE Blog.gorunur=True AND etiket_bulut.etiket_id="&ObjRs("id"))
if not ObjRs2.eof and ObjRs2(0)>1 then boyut=Int(3+ObjRs2(0)*3):response.Write("<a href="""&SiteAdres&"/etiket/"&ObjRs("id")&"/"&SEO_Olustur(ObjRs("etiket"))&".html""><span style=""font-size:"&boyut&"px;"">"&ObjRs("Etiket")&"</span>&nbsp;<font class=""miniMavi"">("&ObjRs2(0)&")</font></a>&nbsp; ")
ObjRs2.close:ObjRs.movenext
next:ObjRs.close:response.Write("<br />")
'kategori module
response.Write("<br /><div class=""ModulSep""><br /></div><h3>Kategori Kategori</h3>")
StrSql="select * from kategoriler order by adi asc":ObjRs.Open StrSql,ObjCon,1,3:ObjRsCount=ObjRs.recordcount
for i=0 to ObjRsCount-1
Set ObjRs2=ObjCon.Execute("Select COUNT(*) From blog where kategori="&ObjRs("id"))
if ObjRs2(0)>0 then Response.Write "<a href="""&SiteAdres&"/kategori/"&ObjRs("id")&"/"&SEO_Olustur(ObjRs("adi"))&".html"" class=""l_m"">"&ObjRs("adi")&"</a> <font class=""miniMavi"">("&ObjRs2(0)&" yazi)</font><br />"
ObjRs2.close:ObjRs.movenext
next:ObjRs.close
'links module
response.Write("<br /><div class=""ModulSep""><br /></div><h3>Takip ettiklerim</h3>")
StrSql="select * from linkler order by baslik asc":ObjRs.Open StrSql,ObjCon,1,3:ObjRsCount=ObjRs.recordcount
for i=0 to ObjRsCount-1:Response.Write "<a href="""&ObjRs("adres")&""" class=""l_t"" target=""_blank"">"&ObjRs("baslik")&"</a><br />":ObjRs.movenext:next:ObjRs.close
Set ObjRs=nothing:Set ObjRs2=nothing
%><br /><br /></div></div>
<div id="footer"><div id="footerNavigation1"><ul><li><a href="<%=SiteAdres%>/" title="BLOG">BLOG</a></li><li><a href="<%=SiteAdres%>/arsiv/" title="Arºiv">Arºiv</a></li><li><a href="<%=SiteAdres%>/projelerim/" title="Projelerim">Projelerim</a></li><li><a href="<%=SiteAdres%>/iletisim/" title="Iletiºim">Iletiºim</a></li><li><a href="<%=SiteAdres%>/CV/">Oz Gecmiºim</a></li><li><a title="Valid XHTML 1.0 Transitional" href="http://validator.w3.org/check?uri=referer" target="_blank"><img src="<%=SiteAdres%>/contents/images/contact/valid.png" width="32" height="12" alt="Valid XHTML 1.0 Transitional" /></a></li><li><a href="http://feeds.feedburner.com/HuseyinSekmenoglu" target="_blank" title="Takip Et"><img src="<%=SiteAdres%>/Contents/images/contact/rss.jpg" alt="" width="12" height="12" />&nbsp;Takip Et</a></li><li><a href="#container" title="Yukari Cik"><img src="<%=SiteAdres%>/Contents/images/yukari_ok.png" alt="" width="12" height="9" />&nbsp;Yukari Cik</a></li></ul></div>
<div class="sayfaTelif clear"><a rel="license" href="http://creativecommons.org/licenses/by-sa/3.0/"><img alt="Creative Commons Lisansi" style="border-width:0" src="http://i.creativecommons.org/l/by-sa/3.0/88x31.png" /></a><br /><span xmlns:dct="http://purl.org/dc/terms/" href="http://purl.org/dc/dcmitype/Text" property="dct:title" rel="dct:type">BLOG</span> by <a xmlns:cc="http://creativecommons.org/ns#" href="huseyinsekmenoglu.net.tc" property="cc:attributionName" rel="cc:attributionURL">Huseyin Sekmenoglu</a> is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-sa/3.0/">Creative Commons "Alinti-Lisansi Devam Ettirme" 3.0 Unported License</a>.<br /><%=Link(SiteTelifHakki)%></div><%if Session("Yonetici")=true then
else%><div id="footerNavigation3"><form action="<%=SiteAdres%>/islem.asp" method="post">
<input type="text" name="Lakap" id="Lakap" value="Lakap" onfocus="javascript:delText('Lakap')" onblur="javascript:writeText('Lakap')" class="TextField GriInput" />&nbsp;
<input type="password" name="Parola" id="Parola" value="Parola" onfocus="javascript:delText('Parola')" onblur="javascript:writeText('Parola')" class="TextField GriInput" />&nbsp;
<input type="submit" name="LoginSubmit" value="Giriº" /></form><%end if
%></div></div></div>
</body></html><%
end sub

Sub NoIE6()%>
<!--[if lt IE 7]><div style='border: 1px solid #F7941D; background: #FEEFDA; text-align: center; clear: both; height: 75px; position: relative;'><div style='position: absolute; right: 3px; top: 3px; font-family: courier new; font-weight: bold;'><a href='#' onclick='javascript:this.parentNode.parentNode.style.display="none"; return false;'><img src='http://www.ie6nomore.com/files/theme/ie6nomore-cornerx.jpg' style='border: none;' alt='Close this notice'/></a></div><div style='width: 640px; margin: 0 auto; text-align: left; padding: 0; overflow: hidden; color: black;'><div style='width: 75px; float: left;'><img src='http://www.ie6nomore.com/files/theme/ie6nomore-warning.jpg' alt='Warning!'/></div><div style='width: 275px; float: left; font-family: Arial, sans-serif;'><div style='font-size: 14px; font-weight: bold; margin-top: 12px;'>You are using an outdated browser</div><div style='font-size: 12px; margin-top: 6px; line-height: 12px;'>For a better experience using this site, please upgrade to a modern web browser.</div></div><div style='width: 75px; float: left;'><a href='http://www.firefox.com' target='_blank'><img src='http://www.ie6nomore.com/files/theme/ie6nomore-firefox.jpg' style='border: none;' alt='Get Firefox'/></a></div><div style='width: 75px; float: left;'><a href='http://www.browserforthebetter.com/download.html' target='_blank'><img src='http://www.ie6nomore.com/files/theme/ie6nomore-ie8.jpg' style='border: none;' alt='Get Internet Explorer 8'/></a></div><div style='width: 73px; float: left;'><a href='http://www.apple.com/safari/download/' target='_blank'><img src='http://www.ie6nomore.com/files/theme/ie6nomore-safari.jpg' style='border: none;' alt='Get Safari 4'/></a></div><div style='float: left;'><a href='http://www.google.com/chrome' target='_blank'><img src='http://www.ie6nomore.com/files/theme/ie6nomore-chrome.jpg' style='border: none;' alt='Get Google Chrome'/></a></div></div></div><![endif]-->
<%end sub

Sub TopNavigation
response.Write("<ul id=""nav""><li class=""nav2""><a href="""&SiteAdres&"/"" title=""BLOG"">BLOG</a></li>"&_
"<li class=""nav2""><a href="""&SiteAdres&"/arsiv/"" title=""Arºiv"">Arºiv</a></li>"&_
"<li class=""nav2""><a href="""&SiteAdres&"/projelerim/"" title=""Projelerim"">Projelerim</a></li>"&_
"<li class=""nav2""><a href="""&SiteAdres&"/CV/"" title=""Oz Gecmiºim"">Oz Gecmiºim</a></li>"&_
"<li class=""nav2""><a href="""&SiteAdres&"/iletisim/"" title=""Iletiºim"">Iletiºim</a></li>"&_
"<li class=""nav2""><a href=""http://feeds.feedburner.com/HuseyinSekmenoglu"" target=""_blank""><img src="""&SiteAdres&"/Contents/images/contact/rss.jpg"" alt="""" width=""12"" height=""12"" />&nbsp;Takip Et</a></li></ul>")
End Sub
Sub YonetimMenusu
response.Write("<ul id=""nav"" class=""dropdown dropdown-horizontal""><li class=""dir"">Yonetim<ul>"&_
"<li><a href="""&SiteAdres&"/Yeni"" title=""Yeni yazi yaz"">Yeni yazi yaz</a></li>"&_
"<li><a href="""&SiteAdres&"/TumYorumlar"" title=""Yorumlar"">Yorumlar</a></li>"&_
"<li><a href="""&SiteAdres&"/TumLinkler"" title=""Linkler"">Linkler</a></li>"&_
"<li><a href="""&SiteAdres&"/TumEtiketler"" title=""Etiketler"">Etiketler</a></li>"&_
"<li><a href="""&SiteAdres&"/TumKategoriler"" title=""Kategoriler"">Kategoriler</a></li>"&_
"<li><a href="""&SiteAdres&"/Ayarlar"" title=""Ayarlar"">Ayarlar</a></li>"&_
"<li><a href="""&SiteAdres&"/DosyaYukle"" title=""Dosya Yukle"">Dosya Yukle</a></li>"&_
"<li><a href="""&SiteAdres&"/islem.asp?Logout=true"" title=""Cikiº yap"">Cikiº yap</a></li>"&_
"</ul></li></ul>")
End Sub
Sub SearchForm
if Request.ServerVariables("QUERY_STRING")>"" then if SayfaAdres(1)="ara" then txtAra=SayfaAdres(2):txtAra=right(txtAra,len(txtAra)-instr(txtAra,"=")):txtAra=TurkceHarfler(txtAra)
response.Write("<form action="""&SiteAdres&"/ara/"" method=""get"" name=""Arama"" class=""floatRight"" ><input name=""Ara"" id=""Ara"" type=""text"" class=""TextFieldArama")
if (txtAra)="" then response.Write(" GriInput"" value=""") else response.Write(""" value=""")
if (txtAra)<>"" then response.Write(txtAra) else response.Write("Ara")
response.Write(""" onfocus=""javascript:delText('Ara')"" onblur=""javascript:writeText('Ara')"" /></form>")
End Sub
%>