<%Sub BlogYaz(BlogID,TekBlogMu)
Set tmpObje=Server.CreateObject ("ADODB.RecordSet"):tmpObje.open "select * from blog where id="&BlogID&";",ObjCon,1,3
If tmpObje.Eof or tmpObje.Bof then
	Response.Write "<div class=""error"">�a��rmaya �al��t���n�z veri silinmi� veya anl�k bir sorun olu�mu� olabilir.</div>"
Else
	if Session("Yonetici")=true then
		if tmpObje("gorunur")=true then aktiftusu="<a href="""&SiteAdres&"/islem.asp?BlogPasif="&tmpObje("id")&""">Pasifle�tir</a>" else aktiftusu="<a href="""&SiteAdres&"/islem.asp?BlogAktif="&tmpObje("id")&""">Aktifle�tir</a>"
		if tmpObje("yorumlanabilir")=true then yorumtusu="<a href="""&SiteAdres&"/islem.asp?BlogYorumKapat="&tmpObje("id")&""">Yorumu Kapat</a>" else yorumtusu="<a href="""&SiteAdres&"/islem.asp?BlogYorumAc="&tmpObje("id")&""">Yorumu A�</a>"
		YonetimTuslari="<ul id=""nav"" class=""dropdown dropdown-horizontal floatRight""><li class=""dir"">Se�enekler<ul><li>"&yorumtusu&"</li><li>"&aktiftusu&"</li>"&_
		"<li><a href="""&SiteAdres&"/Duzenle/"&tmpObje("id")&""">D�zenle</a></li><li><a href="""&SiteAdres&"/TemizDuzenle/"&tmpObje("id")&""">G�venli D�zenle</a></li><li><a href="""&SiteAdres&"/islem.asp?BlogSil="&tmpObje("id")&""">Sil</a></li></ul></li></ul>" 
	else
		YonetimTuslari=""
	end if
	'b�l�nm�� m� b�l�nmemi� mi, 'devam�' yaz�s� olacak m� olmayacakm�
	if TekBlogMu=false and instr(tmpObje("metin"),"<hr />")>0 then Bolunmus=left(tmpObje("metin"),instr(tmpObje("metin"),"<hr />")-1)&"<a href="""&SiteAdres&"/blog/"&tmpObje("id")&"/"&SEO_Olustur(tmpObje("baslik"))&".html""><strong>Devam� �</strong></a><br /><br />" else	Bolunmus=tmpObje("metin")
	'etiket bulma ve yazd�rma i�lemleri
	Set tmpObje2=Server.CreateObject("ADODB.RecordSet"):tmpObje2.open "SELECT etiket_ad.etiket,etiket_ad.id FROM etiket_bulut INNER JOIN etiket_ad ON etiket_bulut.etiket_id = etiket_ad.id WHERE etiket_bulut.blog_id="&tmpObje("id")&" ORDER BY etiket_ad.etiket desc;",ObjCon,1,3
	tag_etiket="":while not tmpObje2.eof:tag_etiket=tag_etiket&", <a href="""&SiteAdres&"/etiket/"&tmpObje2("id")&"/"&SEO_Olustur(tmpObje2("etiket"))&".html"">"&tmpObje2("etiket")&"</a>":tmpObje2.movenext:wend:if tag_etiket<>"" then tag_etiket=right(tag_etiket,len(tag_etiket)-2)
	'kategorisi bulma ve yazd�rma i�lemi
	tmpObje2.close:tmpObje2.open "select * from kategoriler where id="&tmpObje("kategori")&";",ObjCon,1,3
	if not tmpObje2.eof then tmpObjeKat="<a href="""&SiteAdres&"/kategori/"&tmpObje("kategori")&"/"&SEO_Olustur(tmpObje2("adi"))&".html"" class=""l_m"">"&tmpObje2("adi")&"</a>"
	'yorum bulma ve yazd�rma i�lemleri
	tmpObje2.Close:tmpObje2.open "select * from yorum where onay=true and blog="&tmpObje("id")&";",ObjCon,1,3:tmpObjeYorum=tmpObje2.recordcount:tmpObje2.Close
	'burda da blog yaz�l�yor	
	Response.Write "<div id=""Mesaj-"&tmpObje("id")&""" class=""Mesaj"">"&_
	"<h3 class=""blogBaslik""><a href="""&SiteAdres&"/blog/"&tmpObje("id")&"/"&SEO_Olustur(tmpObje("baslik"))&".html"">"&tmpObje("baslik")&"</a></h3>"&_
	"<font class=""blogTarih"">"&TarihGoster(tmpObje("tarih"),tmpObje("saat"))&YonetimTuslari&" &nbsp;&nbsp; <font class=""BlogOkunma"">"&tmpObje("okunma")&" okunma</font> &nbsp;&nbsp; "&_
	"<font class=""BlogYorum""><a href="""&SiteAdres&"/blog/"&tmpObje("id")&"/"&SEO_Olustur(tmpObje("baslik"))&".html&#35;Mesaj-Yorumlar"">"&tmpObjeYorum&" yorum</a></font> &nbsp;&nbsp; "&_
	"<font class=""BlogKategori"">Kategori: "&tmpObjeKat&"</font> &nbsp;&nbsp; <font class=""BlogKategori"">Etiket: "&tag_etiket&"</font> &nbsp;&nbsp; "&_
	"</font><br /><br />"&_
	Video(Google(Resim(Link(Temizle(Bolunmus)))))
	if TekBlogMu=true then'<!-- AddThis Button BEGIN
		Response.Write "<div class=""addthis_toolbox addthis_default_style clear"">"&_
		"<a href=""http://www.addthis.com/bookmark.php?v=250&amp;username=sekmenhuseyin"" class=""addthis_button_compact"">Payla�</a>"&_
		"<span class=""addthis_separator"">|</span><a class=""addthis_button_preferred_1""></a><a class=""addthis_button_preferred_2""></a><a class=""addthis_button_preferred_3""></a><a class=""addthis_button_preferred_4""></a><a class=""addthis_button_preferred_5""></a>"&_
		"</div><script type=""text/javascript"">var addthis_config={""data_track_clickback"":true;ui_language:""tr"";};</script><script type=""text/javascript"" src=""http://s7.addthis.com/js/250/addthis_widget.js#username=sekmenhuseyin""></script><br />"
		'<!-- AddThis Button END -->
	end if
	Response.Write "</div>"
	'sadece bir blog g�steriliyorsa onun yorumlar� da g�sterilecek...
	if TekBlogMu=true then
		if Session("Yonetici")=true then XrtaSQL="" else XrtaSQL=" AND gorunur=true"
		tmpObje2.open "SELECT id,baslik FROM Blog WHERE id<"&tmpObje("id")&XrtaSQL&" ORDER BY id DESC;",ObjCon,1,3
		if not tmpObje2.eof then OncekiYazi="<a href="""&SiteAdres&"/blog/"&tmpObje2(0)&"/"&SEO_Olustur(tmpObje2(1))&".html""><strong>&lt;&lt; �nceki Yaz�</strong></a>" else OncekiYazi=""
		tmpObje2.close:tmpObje2.open "SELECT id,baslik FROM Blog WHERE id>"&tmpObje("id")&XrtaSQL&" ORDER BY id ASC;",ObjCon,1,3
		if not tmpObje2.eof then SonrakiYazi="<a href="""&SiteAdres&"/blog/"&tmpObje2(0)&"/"&SEO_Olustur(tmpObje2(1))&".html""><strong>Sonraki Yaz� &gt;&gt;</strong></a>" else SonrakiYazi=""
		tmpObje2.close:Response.Write "<div class=""clear""><font class=""floatLeft"">"&OncekiYazi&"</font><font class=""floatRight"">"&SonrakiYazi&"</font><br /><br /></div>"
		if Session("Yonetici")=true then StrSql="select * from yorum where blog="&tmpObje("id")&" order by onay desc,tarih desc;"  else StrSql="select * from yorum where blog="&tmpObje("id")&" and onay=true order by tarih desc;"
		Set tmpObjeYorumlar=Server.CreateObject ("ADODB.RecordSet"):tmpObjeYorumlar.Open StrSql, ObjCon, 1, 3
		if tmpObjeYorumlar.Eof or tmpObjeYorumlar.Bof then YorumVar=false else YorumVar=true'daha �nce yorum yaz�lm�� m� yaz�lmam�� m� ��ren.
		if tmpobje("yorumlanabilir")=true then
		if YorumVar=false then Response.Write "<hr /><div id=""Mesaj-Yorumlar""><i>Hen�z yorum yaz�lmam��. �lk yorum yazan sen ol.</i><br /><br /></div>"
		Response.Write yorumMSG&"<div id=""Mesaj-YorumYaz"" class=""Mesaj""><form action="""&SiteAdres&"/islem.asp"" method=""post"" id=""yorumform"">"
		Response.Write "<h2>Yorum Yaz</h2><br />"&_
		"<label for=""ekleyen"">�sim (*)&nbsp;<input type=""text"" id=""ekleyen"" name=""ekleyen"" class=""TextField"" value="""&Z1&""" /></label><br />"&_
		"<label for=""eposta"">e-Posta&nbsp;<input type=""text"" id=""eposta"" name=""eposta"" class=""TextField"" value="""&Z2&""" /></label><br />"&_
		"<label for=""web"">URL&nbsp;<input type=""text"" id=""web"" name=""web"" class=""TextField"" value="""&Z3&""" /></label>"&_
		"<label for=""metin"">Yorum&nbsp;</label><textarea id=""metin"" name=""metin"" rows=""10"" cols=""30""></textarea><br /><br />"&_
		"G�venlik (*)&nbsp;<br />A�a��daki yaz�n�n ayn�s�n� yan tarafa yaz ki insan oldu�unu anlayay�m<label for=""guvenlik"">"&_
		"<img src="""&SiteAdres&"/Contents/captcha.asp"" alt=""This Is CAPTCHA Image"" id=""captcha"" />&nbsp;<a href=""javascript:RefreshImage('captcha')"" style=""font-size:smaller"">Change image</a>"&_
		"<input type=""text"" id=""guvenlik"" name=""guvenlik"" class=""TextField"" value="""" /></label><br />"&_
		"<input type=""hidden"" name=""Blog"" value="""&BlogID&"""/><input type=""hidden"" name=""SEO"" value="""&SEO_Olustur(tmpObje("baslik"))&"""/>"&_
		"<input name=""YorumGonder"" type=""submit"" value=""G�nder"" />"
		Response.Write "</form></div>"
		end if		
		if YorumVar=true then
			Response.Write "<hr /><h2 id=""Mesaj-Yorumlar"">Yorumlar</h2>"
			for i = 1 to 10
				if tmpObjeYorumlar.eof then exit for
				if Session("Yonetici")=true then
				if tmpObjeYorumlar("onay")=false then OnayDugmeleri="&nbsp; <a href="""&SiteAdres&"/islem.asp?YorumOnayla="&tmpObjeYorumlar("id")&""">Onayla</a>&nbsp; -" else OnayDugmeleri="&nbsp; <a href="""&SiteAdres&"/islem.asp?YorumOnaylamama="&tmpObjeYorumlar("id")&""">Sakla</a>&nbsp; -"
				OnayDugmeleri=OnayDugmeleri&"&nbsp; <a href="""&SiteAdres&"/islem.asp?YorumSil="&tmpObjeYorumlar("id")&""">Sil</a>"
				end if
				Response.Write "<div id=""Yorumlar-"&tmpObjeYorumlar("id")&""" class=""Mesaj"
				if tmpObjeYorumlar("eposta")="YoneticiEposta" then
					response.Write("Admin"):tmpEkleyen2=SiteAdres&"/Contents/images/avatars/admin.jpg"
				else
				'avatar eklemek i�in avatar hesaplama kodlar�m
					tmpEkleyen=0:for j=1 to len(ekleyen):tmpEkleyen=tmpEkleyen+asc(mid(ekleyen,j,1)):next:tmpEkleyen2=SiteAdres&"/Contents/images/avatars/"&(tmpEkleyen mod 35)
					if tmpEkleyen mod 2=1 then tmpEkleyen2=tmpEkleyen2&".gif" else tmpEkleyen2=tmpEkleyen2&".jpg"
				'avatar eklemek i�in avatar hesaplama kodlar�m
				end if
				Response.Write """>"
				If Len(tmpObjeYorumlar("web")) < 1 then ekleyen=tmpObjeYorumlar("ekleyen") Else ekleyen="<a href=""http://"&Replace(tmpObjeYorumlar("web"),"http://","",1,-1,1)&""" target=""_blank"" rel=""no follow"">"&tmpObjeYorumlar("ekleyen")&"</a>"
				Response.Write "<img src="""&tmpEkleyen2&""" align=""left"" width=""40"" height=""40"" alt=""Avatar"" style=""margin-right:10px;""/>"&_
				"<font class=""blogTarih"">"&ekleyen&"<font class=""floatRight"">"&TarihGoster(tmpObjeYorumlar("tarih"),tmpObjeYorumlar("saat"))&OnayDugmeleri&_
				"</font></font><div class=""sep""></div>"&Temizle(tmpObjeYorumlar("metin"))&""
				Response.Write "</div>":tmpObjeYorumlar.Movenext
			Next
		end if:tmpObjeYorumlar.Close:Set tmpObjeYorumlar = nothing
	End If
End If
tmpObje.Close:Set tmpObje=nothing:Set tmpObje2=nothing
End Sub

Function TarihGoster(gelen,gelen2)
	if year(gelen)=year(date()) and month(gelen)=month(date()) and day(gelen)=day(date()) then
		tmpG="Bug�n, ":saat=hour(now())-hour(gelen2)
		if saat<>0 then
			tmpG=saat&" saat �nce"
		else
			dakika=minute(now())-minute(gelen2):if dakika<>0 then
				tmpG=tmpG&dakika&" dakika �nce"
			else
				saniye=second(now())-second(gelen2):if saniye<5 then tmpG="Az �nce" else tmpG=tmpG&saniye&" saniye �nce"
			end if
		end if
		TarihGoster=tmpG:Exit Function
	elseif year(gelen)=year(date()-1) and month(gelen)=month(date()-1) and day(gelen)=day(date()-1) then
		TarihGoster="D�n @ "&hour(gelen2)&":"&minute(gelen2):Exit Function
	elseif year(gelen)=year(date()-2) and month(gelen)=month(date()-2) and day(gelen)=day(date()-2) then
		TarihGoster="�nceki g�n @ "&hour(gelen2)&":"&minute(gelen2):Exit Function
	end if
	Aylarim=Array("","Ocak","�ubat","Mart","Nisan","May�s","Haziran","Temmuz","A�ustos","Eyl�l","Ekim","Kas�m","Aral�k")
	Gunlerim=Array("","Pazar","Pazartesi","Sal�","�ar�amba","Per�embe","Cuma","Cumartesi")
	TarihGoster=day(gelen)&" "&Aylarim(month(gelen))&" "&year(gelen)&", "&Gunlerim(Weekday(gelen))&" @ "&hour(gelen2)&":"&minute(gelen2)
End Function

Function Temizle(gelenveri)
	gelenveri=Replace(gelenveri ,"[code]","<div id=""kod"">")
	gelenveri=Replace(gelenveri ,"[/code]","</div>")
	gelenveri=Replace(gelenveri ,"[alinti]","<div id=""alinti"">")
	gelenveri=Replace(gelenveri ,"[/alinti]","</div>")
	gelenveri=Replace(gelenveri ,"<hr />","")
	Temizle=trim(gelenveri)
End Function

Function QS_Temizle(gelenveri)
if isnull(gelenveri)=true then exit function
gelenveri=Replace(gelenveri,"'","",1,-1,1)
gelenveri=Replace(gelenveri,"<","",1,-1,1)
gelenveri=Replace(gelenveri,">","",1,-1,1)
gelenveri=Replace(gelenveri,"&","",1,-1,1)
gelenveri=Replace(gelenveri,"%","",1,-1,1)
gelenveri=Replace(gelenveri,"?","",1,-1,1)
gelenveri=Replace(gelenveri,";","",1,-1,1)
gelenveri=Replace(gelenveri,"+","",1,-1,1)
gelenveri=Replace(gelenveri,"""","",1,-1,1)
gelenveri=Replace(gelenveri,"  "," ",1,-1,1)
QS_Temizle=trim(gelenveri)
end Function

Function TurkceHarfler(gelenveri)
gelenveri=Replace(gelenveri,"%FD","�",1,-1,1)
gelenveri=Replace(gelenveri,"%F0","�",1,-1,1)
gelenveri=Replace(gelenveri,"%FC","�",1,-1,1)
gelenveri=Replace(gelenveri,"%FE","�",1,-1,1)
gelenveri=Replace(gelenveri,"%E7","�",1,-1,1)
gelenveri=Replace(gelenveri,"%F6","�",1,-1,1)
gelenveri=Replace(gelenveri,"%DD","�",1,-1,1)
gelenveri=Replace(gelenveri,"%D0","�",1,-1,1)
gelenveri=Replace(gelenveri,"%DC","�",1,-1,1)
gelenveri=Replace(gelenveri,"%DE","�",1,-1,1)
gelenveri=Replace(gelenveri,"%C7","�",1,-1,1)
gelenveri=Replace(gelenveri,"%D6","�",1,-1,1)
gelenveri=Replace(gelenveri,"+"," ",1,-1,1)
TurkceHarfler=gelenveri
End Function

Sub Guvenlik(gelenveri)
	If gelenveri="" or IsNumeric(gelenveri)=False Then Response.Redirect SiteAdres&"/SayfaBulunamadi"
End Sub

Public Function Link(byVal Text) 
	Set objReg=New RegEXP:objReg.Global=True:objReg.IgnoreCase=True 
	objReg.Pattern="\[link:\s*(.+?)\]\s*(.+?)\[/link]" 
	Text=objReg.Replace(Text,"<a href=""$1"" target=""_blank"" class=""acik"">$2</a>") 
	Link=Text 
End Function

Public Function Resim(Text)
	Text=Replace(Text,"../upload",SiteAdres&"/upload",1,-1,1)
	Resim=Text 
End Function

Public Function Google(byVal Text) 
	Set objReg=New RegEXP:objReg.Global=True:objReg.IgnoreCase=True 
	objReg.Pattern="\[ara]\s*(.+?)\[/ara]" 
	Text=objReg.Replace(Text,"<a href=""http://www.google.com.tr/search?hl=tr&amp;q=$1"" target=""_blank"" class=""acik"">$1</a>") 
	Google=Text 
End Function

Public Function Video(byVal Text) 
	Set objReg=New RegEXP:objReg.Global=True:objReg.IgnoreCase=True 
	objReg.Pattern="\[video]\s*(.+?)\[/video]" 
	'Text=objReg.Replace(Text,"<center><object width=""450"" height=""334""><param name=""movie"" value=""http://www.youtube.com/v/$1""></param><param name=""wmode"" value=""transparent""></param><embed src=""http://www.youtube.com/v/$1"" type=""application/x-shockwave-flash"" wmode=""transparent"" width=""400"" height=""334""></embed></object></center>") 
	Text=objReg.Replace(Text,"<center><object type=""application/x-shockwave-flash"" data=""http://www.youtube.com/v/$1&amp;rel=1"" width=""450"" height=""334""><param name=""movie"" value=""http://www.youtube.com/v/$1&amp;rel=1"" /><param name=""FlashVars"" value=""playerMode=embedded"" /></object></center>")
	Video=Text 
End Function

Sub WriteWebLog%><!-- #Include File = "abt-info.asp" --><%Set ObjRs=Server.CreateObject ("ADODB.RecordSet")
IP=request.ServerVariables("REMOTE_HOST"):ObjRs.open "select * from WebStats where IP='" & IP & "' and Date=Date()",ObjCon,3,3
'(daha once kayit yoksa) ve (bot degilse) ve (cesitli arama motoru degilse) yeni kayit ekle
if ObjRs.eof and Instr(ua, "bot")=false and Instr(ua, "Yandex")=false and Instr(ua, "Lynxy")=false and Instr(ua, "web.archive.org")=false then
	if (Browser="Unknown" and OS="Unknown") or (len(trim(Browser))=0 and len(trim(OS))=0) then
	else
		ObjRs.close:ObjRs.open "select * from WebStats",ObjCon,3,3:ObjRs.AddNew:ObjRs("IP")=IP:ObjRs("Browser")=Browser:ObjRs("OS")=os:ObjRs.Update
	end if
end if
ObjRs.close:set ObjRs=nothing
if err.number<>0 then response.end()
End Sub

Sub BanaMailAtHaberimOlsun
	Dim iMsg, iConf, Flds:Set iMsg=CreateObject("CDO.Message"):Set iConf=CreateObject("CDO.Configuration"):Set Flds=iConf.Fields
	schema="http://schemas.microsoft.com/cdo/configuration/"
	Flds.Item(schema & "sendusing")=2
	Flds.Item(schema & "smtpserver")="smtp.gmail.com"
	Flds.Item(schema & "smtpserverport")=465
	Flds.Item(schema & "smtpauthenticate")=1
	Flds.Item(schema & "sendusername")="sekmenhuseyin@gmail.com"
	Flds.Item(schema & "sendpassword")="85liMartKedisi"
	Flds.Item(schema & "smtpusessl")=1
	Flds.Update
	With iMsg
	.To="sekmenhuseyin@gmail.com"
	.From="sekmenhuseyin@gmail.com"
	.Subject="You have new message"
	.HTMLBody="You have new message in yor blog. Please sign in and read it. "
	.Sender="H�seyin Sekmeno�lu - Blog"
	.Organization="H�seyin Sekmeno�lum"
	.ReplyTo="sekmenhuseyin@gmail.com"
	Set .Configuration=iConf
	SendEmailGmail=.Send
	End With
	Set iMsg=Nothing:Set iConf=Nothing:Set Flds=Nothing
End Sub

Function SEO_Olustur(BlogBaslik)
BlogBaslik=lcase(BlogBaslik)
BlogBaslik=Replace(BlogBaslik ,"!","",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"?","",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"'","",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"_","-",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"+","-",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"�","i",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"�","s",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"�","g",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"�","u",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"�","o",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"�","c",1,-1,1)
BlogBaslik=Replace(BlogBaslik ," ","-",1,-1,1)
if right(BlogBaslik,1)="-" then BlogBaslik=left(BlogBaslik,len(BlogBaslik)-1)
SEO_Olustur=BlogBaslik
End Function%>